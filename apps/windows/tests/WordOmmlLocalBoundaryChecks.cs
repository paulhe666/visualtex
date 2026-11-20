using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.Remoting.Messaging;
using System.Runtime.Remoting.Proxies;
using Word = Microsoft.Office.Interop.Word;

// Model the observed Word 2021 behavior: Range.OMaths returns OMath ranges
// clipped to that Range, even when Count==1 and NestingLevel==0. No live Word.
public static class WordOmmlLocalBoundaryChecks
{
    public sealed class Result
    {
        public string Case, Expected, Actual, Error;
        public bool Passed;
        public int GlobalReads, ContentOutsideHost, SelectionWrites, ParagraphExpansions;
        public string[] UnexpectedCalls;
    }
    private sealed class Proxy : RealProxy
    {
        private readonly Func<MethodInfo, object[], object> handler;
        internal Proxy(Type type, Func<MethodInfo, object[], object> handler) : base(type) { this.handler = handler; }
        public override IMessage Invoke(IMessage message)
        {
            var call = (IMethodCallMessage)message;
            try { return new ReturnMessage(handler((MethodInfo)call.MethodBase, call.Args), call.Args, call.ArgCount, call.LogicalCallContext, call); }
            catch (Exception e) { return new ReturnMessage(e, call); }
        }
    }
    private sealed class Host
    {
        internal int Start, End;
        internal Word.WdOMathType Type;
        internal Host(int start, int end, bool display) { Start = start; End = end; Type = display ? Word.WdOMathType.wdOMathDisplay : Word.WdOMathType.wdOMathInline; }
    }
    private sealed class Model
    {
        internal readonly List<string> Unexpected = new List<string>();
        internal int GlobalReads, OutsideExports, SelectionWrites, Expansions;
        internal readonly Host[] Hosts;
        internal readonly bool InTable, ReferenceResult;
        internal readonly Word.WdStoryType Story;
        internal readonly int ParagraphStart, ParagraphEnd;
        internal readonly Word.Document Document;
        internal Model(Host[] hosts, bool inTable, bool reference, Word.WdStoryType story, int paragraphStart = 0)
        {
            Hosts = hosts; InTable = inTable; ReferenceResult = reference; Story = story;
            ParagraphStart = paragraphStart; ParagraphEnd = paragraphStart + 100;
            Document = Make<Word.Document>("Document", delegate(MethodInfo m, object[] a) {
                if (m.Name == "get_OMaths" || m.Name == "get_Content" || m.Name == "get_Paragraphs" || m.Name == "get_Bookmarks" || m.Name == "get_InlineShapes") GlobalReads++;
                return Unknown("Document", m);
            });
        }
        internal Word.Range Range(int start, int end)
        {
            return Make<Word.Range>("Range", delegate(MethodInfo m, object[] a) {
                switch (m.Name)
                {
                    case "get_Start": return start;
                    case "get_End": return end;
                    case "get_StoryType": return Story;
                    case "get_Document": return Document;
                    case "get_Duplicate": return Range(start, end);
                    case "get_OMaths": return Maths(start, end);
                    case "get_Paragraphs": return Paragraphs();
                    case "get_Fields": return Fields();
                    case "get_Bookmarks": return Empty<Word.Bookmarks>("Bookmarks");
                    case "get_Information": return InTable;
                    case "MoveStart": { int n = Convert.ToInt32(a[1]); int old = start; start = Math.Max(ParagraphStart, start + n); return start - old; }
                    case "MoveEnd": { int n = Convert.ToInt32(a[1]); int old = end; end = Math.Min(ParagraphEnd, end + n); return end - old; }
                    case "Expand":
                        if ((Word.WdUnits)a[0] != Word.WdUnits.wdParagraph) throw new InvalidOperationException("Only local paragraph expansion is permitted.");
                        int added = (start - ParagraphStart) + (ParagraphEnd - end);
                        start = ParagraphStart; end = ParagraphEnd; Expansions++; return added;
                    case "get_WordOpenXML":
                        if (!Hosts.Any(h => h.Start == start && h.End == end)) OutsideExports++;
                        return "<m:oMath xmlns:m=\"http://schemas.openxmlformats.org/officeDocument/2006/math\"><m:r><m:t>x</m:t></m:r></m:oMath>";
                    case "Select": case "Delete": case "InsertXML": case "set_Text":
                        SelectionWrites++; throw new InvalidOperationException("Host discovery cannot mutate or select text.");
                    default: return Unknown("Range", m);
                }
            });
        }
        private Word.OMaths Maths(int start, int end)
        {
            var matching = Hosts.Where(h => start == end ? h.Start <= start && start <= h.End : start < h.End && end > h.Start).ToArray();
            return Make<Word.OMaths>("OMaths", delegate(MethodInfo m, object[] a) {
                if (m.Name == "get_Count") return matching.Length;
                if (m.Name == "get_Item")
                {
                    Host h = matching[Convert.ToInt32(a[0]) - 1];
                    return Make<Word.OMath>("OMath", delegate(MethodInfo mm, object[] aa) {
                        if (mm.Name == "get_Range") return Range(Math.Max(start, h.Start), Math.Min(end, h.End));
                        if (mm.Name == "get_Type") return h.Type;
                        if (mm.Name == "get_NestingLevel") return 0;
                        return Unknown("OMath", mm);
                    });
                }
                return Unknown("OMaths", m);
            });
        }
        private Word.Paragraphs Paragraphs()
        {
            return Make<Word.Paragraphs>("Paragraphs", delegate(MethodInfo m, object[] a) {
                if (m.Name == "get_Count") return 1;
                if (m.Name == "get_Item") return Make<Word.Paragraph>("Paragraph", delegate(MethodInfo mm, object[] aa) {
                    if (mm.Name == "get_Range") return Range(ParagraphStart, ParagraphEnd);
                    return Unknown("Paragraph", mm);
                });
                return Unknown("Paragraphs", m);
            });
        }
        private Word.Fields Fields()
        {
            return Make<Word.Fields>("Fields", delegate(MethodInfo m, object[] a) {
                if (m.Name == "get_Count") return ReferenceResult ? 1 : 0;
                if (m.Name == "get_Item" && ReferenceResult) return Make<Word.Field>("Field", delegate(MethodInfo mm, object[] aa) {
                    if (mm.Name == "get_Type") return Word.WdFieldType.wdFieldRef;
                    if (mm.Name == "get_Result") return Range(Hosts[0].Start, Hosts[0].End);
                    return Unknown("Field", mm);
                });
                return Unknown("Fields", m);
            });
        }
        private T Empty<T>(string label) where T : class { return Make<T>(label, delegate(MethodInfo m, object[] a) { if (m.Name == "get_Count") return 0; return Unknown(label, m); }); }
        private T Make<T>(string label, Func<MethodInfo, object[], object> handler) where T : class
        {
            return (T)new Proxy(typeof(T), delegate(MethodInfo m, object[] a) {
                if (m.Name == "GetType") return typeof(T);
                if (m.DeclaringType == typeof(object))
                {
                    if (m.Name == "ToString") return label;
                    if (m.Name == "GetHashCode") return label.GetHashCode();
                    if (m.Name == "Equals") return false;
                }
                return handler(m, a);
            }).GetTransparentProxy();
        }
        private object Unknown(string label, MethodInfo m) { Unexpected.Add(label + "." + m.Name); throw new NotSupportedException(label + "." + m.Name); }
    }
    private static object Property(object value, string name) { return value.GetType().GetProperty(name, BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic).GetValue(value, null); }
    public static Result[] Run(string assemblyPath)
    {
        Assembly assembly = Assembly.LoadFrom(assemblyPath);
        Type type = assembly.GetType("VisualTeX.WordVsto.WordFormulaHostResolver", true);
        Type kind = assembly.GetType("VisualTeX.WordVsto.WordFormulaHostKind", true);
        MethodInfo resolve = type.GetMethod("ResolveLocal", BindingFlags.Static | BindingFlags.NonPublic);
        object omml = Enum.Parse(kind, "Omml");
        var results = new List<Result>();
        Action<string, int, int, string, bool, bool, bool, Word.WdStoryType, int> check = (name, start, end, expected, display, inTable, reference, story, shift) => {
            var model = new Model(new[] { new Host(shift + 10, shift + 30, display), new Host(shift + 50, shift + 70, false) }, inTable, reference, story, shift);
            string actual = null, error = null;
            try
            {
                object host = resolve.Invoke(null, new object[] { model.Document, model.Range(shift + start, shift + end), omml });
                if (host == null) actual = "none";
                else
                {
                    object range = Property(host, "Range");
                    actual = (Convert.ToInt32(Property(range, "Start")) - shift) + ":" + (Convert.ToInt32(Property(range, "End")) - shift);
                    if (!Equals(Property(range, "StoryType"), story)) actual += ":wrong-story";
                    if (!Equals(Property(host, "WithinTable"), inTable)) actual += ":wrong-table";
                }
            }
            catch (TargetInvocationException e) { actual = e.InnerException is InvalidDataException ? "ambiguous" : "exception"; error = e.InnerException.ToString(); }
            results.Add(new Result { Case = name, Expected = expected, Actual = actual, Error = error,
                Passed = actual == expected && model.GlobalReads == 0 && model.OutsideExports == 0 && model.SelectionWrites == 0 && model.Unexpected.Count == 0,
                GlobalReads = model.GlobalReads, ContentOutsideHost = model.OutsideExports, SelectionWrites = model.SelectionWrites,
                ParagraphExpansions = model.Expansions, UnexpectedCalls = model.Unexpected.ToArray() });
        };
        var main = Word.WdStoryType.wdMainTextStory;
        check("exact-inline", 10, 30, "10:30", false, false, false, main, 0);
        check("clipped-prefix", 11, 30, "10:30", false, false, false, main, 0);
        check("clipped-suffix", 10, 29, "10:30", false, false, false, main, 0);
        check("interior-fragment", 17, 21, "10:30", false, false, false, main, 0);
        check("interior-caret", 18, 18, "10:30", false, false, false, main, 0);
        check("boundary-caret", 30, 30, "10:30", false, false, false, main, 0);
        check("second-formula-not-first", 56, 62, "50:70", false, false, false, main, 0);
        check("two-formulas-rejected", 20, 60, "ambiguous", false, false, false, main, 0);
        check("prose-between-formulas", 38, 42, "none", false, false, false, main, 0);
        check("prose-caret-near-formula", 32, 32, "none", false, false, false, main, 0);
        check("display-clipped-prefix", 11, 30, "10:30", true, false, false, main, 0);
        check("table-cell-local", 11, 29, "10:30", false, true, false, main, 0);
        check("header-story-local", 11, 29, "10:30", false, false, false, Word.WdStoryType.wdPrimaryHeaderStory, 0);
        check("far-document-local-only", 11, 29, "10:30", false, false, false, main, 1000000);
        check("reference-result-not-editable", 11, 29, "none", false, false, true, main, 0);
        return results.ToArray();
    }
}
