using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Runtime.Remoting.Messaging;
using System.Runtime.Remoting.Proxies;
using Word = Microsoft.Office.Interop.Word;

// Models the measured Word 2021 behavior: range.OMaths[i].Range is clipped to
// the requesting range, while ParentOMath exposes the physical equation. These
// tests forbid global discovery, XML export and mutation in host resolution.
public static class WordOmmlBoundaryChecks
{
    public sealed class Result
    {
        public string Case, Expected, Actual;
        public bool Passed;
        public int ParentReads, ContentExports, Mutations;
        public string[] Unexpected;
    }
    private sealed class Proxy : RealProxy
    {
        readonly Func<MethodInfo, object[], object> handler;
        internal Proxy(Type type, Func<MethodInfo, object[], object> handler) : base(type) { this.handler = handler; }
        public override IMessage Invoke(IMessage message)
        {
            var call = (IMethodCallMessage)message;
            try { var args = call.Args; var value = handler((MethodInfo)call.MethodBase, args); return new ReturnMessage(value, args, args.Length, call.LogicalCallContext, call); }
            catch (Exception e) { return new ReturnMessage(e, call); }
        }
    }
    private sealed class Model
    {
        internal readonly int[][] Equations = { new[] { 10, 20 }, new[] { 30, 55 }, new[] { 65, 75 } };
        internal readonly List<string> Unexpected = new List<string>();
        internal int ParentReads, Exports, Mutations;
        internal Word.Document Document;
        internal Model()
        {
            Document = Make<Word.Document>("Document", delegate(MethodInfo m, object[] a) {
                if (m.Name == "Range") return Range(Convert.ToInt32(a[0]), Convert.ToInt32(a[1]));
                // Document.OMaths/Content/Bookmarks scans must not be introduced.
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
                    case "get_StoryType": return Word.WdStoryType.wdMainTextStory;
                    case "get_Duplicate": return Range(start, end);
                    case "get_Document": return Document;
                    case "get_OMaths": return Maths(start, end);
                    case "get_Tables": return Empty<Word.Tables>("Tables");
                    case "get_Fields": return Empty<Word.Fields>("Fields");
                    case "get_Paragraphs": return Paragraphs();
                    case "get_Information": return false;
                    case "MoveStart": start = System.Math.Max(0, start + Convert.ToInt32(a[1])); return Convert.ToInt32(a[1]);
                    case "MoveEnd": end = System.Math.Min(100, end + Convert.ToInt32(a[1])); return Convert.ToInt32(a[1]);
                    case "Expand":
                        if ((Word.WdUnits)a[0] != Word.WdUnits.wdParagraph) return Unknown("Range", m);
                        int added = 100 - (end - start); start = 0; end = 100; return added;
                    case "get_WordOpenXML": Exports++; throw new InvalidOperationException("Host discovery must not export payloads.");
                    default: return Unknown("Range", m);
                }
            });
        }
        Word.OMaths Maths(int start, int end)
        {
            var found = Equations.Where(e => start == end ? e[0] <= start && start <= e[1] : e[0] < end && start < e[1]).ToArray();
            return Make<Word.OMaths>("OMaths", delegate(MethodInfo m, object[] a) {
                if (m.Name == "get_Count") return found.Length;
                if (m.Name == "get_Item")
                {
                    var equation = found[Convert.ToInt32(a[0]) - 1];
                    return Math(equation, MathMax(start, equation[0]), MathMin(end, equation[1]), false);
                }
                return Unknown("OMaths", m);
            });
        }
        static int MathMax(int x, int y) { return System.Math.Max(x, y); }
        static int MathMin(int x, int y) { return System.Math.Min(x, y); }
        Word.OMath Math(int[] equation, int start, int end, bool physical)
        {
            return Make<Word.OMath>("OMath", delegate(MethodInfo m, object[] a) {
                if (m.Name == "get_Range") return Range(start, end);
                if (m.Name == "get_Type") return Word.WdOMathType.wdOMathInline;
                if (m.Name == "get_NestingLevel") return 0;
                if (m.Name == "get_ParentOMath") { ParentReads++; return Math(equation, equation[0], equation[1], true); }
                return Unknown("OMath", m);
            });
        }
        Word.Paragraphs Paragraphs()
        {
            return Make<Word.Paragraphs>("Paragraphs", delegate(MethodInfo m, object[] a) {
                if (m.Name == "get_Count") return 1;
                if (m.Name == "get_Item") return Make<Word.Paragraph>("Paragraph", delegate(MethodInfo n, object[] b) {
                    if (n.Name == "get_Range") return Range(0, 100);
                    return Unknown("Paragraph", n);
                });
                return Unknown("Paragraphs", m);
            });
        }
        T Empty<T>(string name) where T : class { return Make<T>(name, delegate(MethodInfo m, object[] a) { if (m.Name == "get_Count") return 0; return Unknown(name, m); }); }
        T Make<T>(string name, Func<MethodInfo, object[], object> handler) where T : class
        {
            return (T)new Proxy(typeof(T), delegate(MethodInfo m, object[] a) {
                if (m.Name == "GetType" && a.Length == 0) return typeof(T);
                if (m.DeclaringType == typeof(object))
                {
                    if (m.Name == "ToString") return "OMML boundary double: " + name;
                    if (m.Name == "GetHashCode") return name.GetHashCode();
                    if (m.Name == "Equals") return false;
                }
                return handler(m, a);
            }).GetTransparentProxy();
        }
        object Unknown(string name, MethodInfo m)
        {
            if (m.Name.StartsWith("set_", StringComparison.Ordinal) || m.Name == "Delete" || m.Name == "InsertXML") Mutations++;
            Unexpected.Add(name + "." + m.Name);
            throw new NotSupportedException("Unmodelled call: " + name + "." + m.Name);
        }
    }
    static object Property(object value, string name) { return value.GetType().GetProperty(name, BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic).GetValue(value, null); }
    static string Address(object host)
    {
        if (host == null) return "none";
        var range = Property(host, "Range");
        return Property(range, "Start") + ":" + Property(range, "End");
    }
    public static Result[] Run(string assemblyPath)
    {
        var assembly = Assembly.LoadFrom(assemblyPath);
        var resolver = assembly.GetType("VisualTeX.WordVsto.WordFormulaHostResolver", true);
        var kind = Enum.Parse(assembly.GetType("VisualTeX.WordVsto.WordFormulaHostKind", true), "Omml");
        var flags = BindingFlags.Static | BindingFlags.Public | BindingFlags.NonPublic;
        var results = new List<Result>();
        var cases = new[] {
            new { Name="local-start-inside-surrogate", Start=31, End=55, Local=true, Expected="30:55" },
            new { Name="local-end-inside-surrogate", Start=30, End=54, Local=true, Expected="30:55" },
            new { Name="local-deep-interior-not-fixed-padding", Start=41, End=43, Local=true, Expected="30:55" },
            new { Name="local-caret-inside-equation", Start=41, End=41, Local=true, Expected="30:55" },
            new { Name="local-prose-remains-none", Start=85, End=85, Local=true, Expected="none" },
            new { Name="scope-start-inside-surrogate", Start=31, End=56, Local=false, Expected="30:55" },
            new { Name="scope-end-inside-surrogate", Start=30, End=54, Local=false, Expected="30:55" },
            new { Name="scope-deep-interior-excludes-neighbors", Start=41, End=43, Local=false, Expected="30:55" },
            new { Name="scope-two-partial-endpoints", Start=15, End=42, Local=false, Expected="10:20,30:55" },
            new { Name="scope-full-three-hosts", Start=0, End=100, Local=false, Expected="10:20,30:55,65:75" }
        };
        foreach (var test in cases)
        {
            var model = new Model();
            string actual;
            try
            {
                var method = resolver.GetMethod(test.Local ? "ResolveLocal" : "CaptureScopeKindIndex", flags);
                var value = method.Invoke(null, new object[] { model.Document, model.Range(test.Start, test.End), kind });
                actual = test.Local ? Address(value) : String.Join(",", ((IEnumerable)value).Cast<object>().Select(Address));
            }
            catch (Exception e) { actual = "ERROR: " + (e.InnerException ?? e).Message; }
            results.Add(new Result {
                Case = test.Name, Expected = test.Expected, Actual = actual,
                ParentReads = model.ParentReads, ContentExports = model.Exports, Mutations = model.Mutations,
                Unexpected = model.Unexpected.ToArray(),
                Passed = actual == test.Expected && model.Exports == 0 && model.Mutations == 0 && model.Unexpected.Count == 0
            });
        }
        return results.ToArray();
    }
}
