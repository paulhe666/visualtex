using System;
using System.Collections.Generic;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Runtime.Remoting.Messaging;
using System.Runtime.Remoting.Proxies;
using Word = Microsoft.Office.Interop.Word;

// These are managed interface doubles, not live Office objects. Run the same
// checks against the saved baseline DLL and the built candidate. A controlled
// export exception makes the unsafe old path observable without crashing Word.
public static class WordPassiveObservationChecks
{
    public sealed class Result
    {
        public string Case;
        public int ContentExports;
        public int ObjectActivations;
        public int ClipboardCopies;
        public int SelectionWrites;
        public int LayoutQueries;
        public string[] UnexpectedCalls;
        public bool BoundarySatisfied;
    }

    private sealed class InterfaceDouble : RealProxy
    {
        private readonly Func<MethodInfo, object[], object> invoke;
        internal InterfaceDouble(Type type, Func<MethodInfo, object[], object> handler) : base(type)
        { invoke = handler; }
        public override IMessage Invoke(IMessage message)
        {
            var call = (IMethodCallMessage)message;
            try
            {
                object value = invoke((MethodInfo)call.MethodBase, call.Args);
                return new ReturnMessage(value, call.Args, call.ArgCount, call.LogicalCallContext, call);
            }
            catch (Exception error) { return new ReturnMessage(error, call); }
        }
    }

    private sealed class Model
    {
        internal readonly List<string> Unexpected = new List<string>();
        internal int Exports, Activations, ClipboardCopies, SelectionWrites, LayoutQueries;
        internal readonly string ProgId;
        internal readonly bool MissingRegistration;
        internal Word.Application Application;
        internal Word.Document Document;
        internal Word.Selection Selection;
        internal Word.InlineShape Shape;
        internal Word.InlineShapes Shapes;
        internal Word.OLEFormat OleFormat;
        internal Word.OMaths EmptyMaths;
        internal Word.Range SelectedRange;

        internal Model(string progId, bool missingRegistration, bool collapsed, int caretPosition = 11, int shapeCount = 1)
        {
            ProgId = progId; MissingRegistration = missingRegistration;
            EmptyMaths = Make<Word.OMaths>("OMaths", delegate(MethodInfo m, object[] a) {
                if (m.Name == "get_Count") return 0;
                return Unknown("OMaths", m);
            });
            OleFormat = Make<Word.OLEFormat>("OLEFormat", delegate(MethodInfo m, object[] a) {
                if (m.Name == "get_ProgID")
                {
                    if (MissingRegistration) throw new COMException("Controlled missing registration", unchecked((int)0x80040154));
                    return ProgId;
                }
                if (m.Name == "get_Object" || m.Name == "Activate" || m.Name == "DoVerb")
                { Activations++; throw new InvalidOperationException("Third-party OLE activation is forbidden in observation."); }
                return Unknown("OLEFormat", m);
            });
            Shape = Make<Word.InlineShape>("Shape", delegate(MethodInfo m, object[] a) {
                switch (m.Name)
                {
                    case "get_Type": return Word.WdInlineShapeType.wdInlineShapeEmbeddedOLEObject;
                    case "get_OLEFormat": return OleFormat;
                    case "get_Range": return Range(10, 11);
                    case "get_AlternativeText": case "get_Title": return String.Empty;
                    default: return Unknown("Shape", m);
                }
            });
            Shapes = Make<Word.InlineShapes>("InlineShapes", delegate(MethodInfo m, object[] a) {
                if (m.Name == "get_Count") return shapeCount;
                if (m.Name == "get_Item") return Shape;
                return Unknown("InlineShapes", m);
            });
            SelectedRange = Range(collapsed ? caretPosition : 10, collapsed ? caretPosition : 11);
            Selection = Make<Word.Selection>("Selection", delegate(MethodInfo m, object[] a) {
                if (m.Name == "get_Range") return SelectedRange;
                if (m.Name == "get_Document") return Document;
                if (m.Name == "SetRange" || m.Name.StartsWith("set_", StringComparison.Ordinal))
                { SelectionWrites++; throw new InvalidOperationException("Foreign selection mutation is forbidden."); }
                return Unknown("Selection", m);
            });
            Document = Make<Word.Document>("Document", delegate(MethodInfo m, object[] a) {
                if (m.Name == "get_Application") return Application;
                return Unknown("Document", m);
            });
            Application = Make<Word.Application>("Application", delegate(MethodInfo m, object[] a) {
                if (m.Name == "get_ActiveDocument") return Document;
                if (m.Name == "get_Selection") return Selection;
                return Unknown("Application", m);
            });
        }

        private Word.Range Range(int start, int end)
        {
            return Make<Word.Range>("Range", delegate(MethodInfo m, object[] a) {
                switch (m.Name)
                {
                    case "get_Start": return start;
                    case "get_End": return end;
                    case "get_StoryType": return Word.WdStoryType.wdMainTextStory;
                    case "get_Duplicate": return Range(start, end);
                    case "get_InlineShapes": return Shapes;
                    case "get_OMaths": return EmptyMaths;
                    case "get_Document": return Document;
                    case "get_Information": LayoutQueries++; return false;
                    case "MoveStart": start += Convert.ToInt32(a[1]); return Convert.ToInt32(a[1]);
                    case "MoveEnd": end += Convert.ToInt32(a[1]); return Convert.ToInt32(a[1]);
                    case "Copy":
                        ClipboardCopies++;
                        throw new InvalidOperationException("Third-party clipboard capture is forbidden in observation.");
                    case "get_WordOpenXML":
                        Exports++;
                        throw new InvalidOperationException("Controlled export barrier: a passive caller must never reach this.");
                    default: return Unknown("Range", m);
                }
            });
        }

        private T Make<T>(string label, Func<MethodInfo, object[], object> handler) where T : class
        {
            return (T)new InterfaceDouble(typeof(T), delegate(MethodInfo method, object[] args) {
                if (method.Name == "GetType" && args.Length == 0) return typeof(T);
                if (method.DeclaringType == typeof(object))
                {
                    if (method.Name == "ToString") return "Managed test double: " + label;
                    if (method.Name == "GetHashCode") return label.GetHashCode();
                    if (method.Name == "Equals") return false;
                }
                return handler(method, args);
            }).GetTransparentProxy();
        }

        private object Unknown(string label, MethodInfo method)
        {
            Unexpected.Add(label + "." + method.Name);
            throw new NotSupportedException("Test double must explicitly model " + label + "." + method.Name);
        }
    }

    public static Result[] Run(string assemblyPath)
    {
        Assembly assembly = Assembly.LoadFrom(assemblyPath);
        Type serviceType = assembly.GetType("VisualTeX.WordVsto.WordFormulaService", true);
        var results = new List<Result>();
        foreach (string id in new[] { "Equation.DSMT4", "Unknown.ThirdParty.1", "missing-registration" })
        {
            bool missing = id == "missing-registration";
            foreach (bool caret in new[] { false, true })
            {
                var model = new Model(id, missing, caret);
                object service = Activator.CreateInstance(serviceType, new object[] { model.Application });
                string method = caret ? "NormalizeTypingCaretAfterInlineFormula" : "GetSelectedFormulaFontSize";
                serviceType.GetMethod(method, BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic)
                    .Invoke(service, caret ? new object[] { model.Selection } : new object[0]);
                results.Add(new Result {
                    Case = id + ":" + method,
                    ContentExports = model.Exports, ObjectActivations = model.Activations,
                    ClipboardCopies = model.ClipboardCopies,
                    SelectionWrites = model.SelectionWrites, UnexpectedCalls = model.Unexpected.ToArray(),
                    BoundarySatisfied = model.Exports == 0 && model.Activations == 0 && model.ClipboardCopies == 0 && model.SelectionWrites == 0 && model.Unexpected.Count == 0
                });
            }
        }
        // Ordinary text immediately before an owned OLE must not activate it or
        // request document layout. Ambiguous touching hosts also remain rejected.
        foreach (int count in new[] { 1, 2 })
        {
            var model = new Model("VisualTeX.Formula.1", false, true, 10, count);
            object service = Activator.CreateInstance(serviceType, new object[] { model.Application });
            serviceType.GetMethod("NormalizeTypingCaretAfterInlineFormula").Invoke(service, new object[] { model.Selection });
            results.Add(new Result {
                Case = count == 1 ? "visualtex-leading-caret-no-payload-or-layout" : "ambiguous-visualtex-caret-no-payload-or-layout",
                ContentExports = model.Exports, ObjectActivations = model.Activations,
                ClipboardCopies = model.ClipboardCopies, SelectionWrites = model.SelectionWrites,
                LayoutQueries = model.LayoutQueries, UnexpectedCalls = model.Unexpected.ToArray(),
                BoundarySatisfied = model.Exports == 0 && model.Activations == 0 && model.ClipboardCopies == 0
                    && model.SelectionWrites == 0 && model.LayoutQueries == 0 && model.Unexpected.Count == 0
            });
        }
        MethodInfo command = serviceType.GetMethod("ReadSelectedFormulaFontSizeForCommand", BindingFlags.Instance | BindingFlags.NonPublic);
        if (command != null)
        {
            var model = new Model("Equation.DSMT4", false, false);
            object service = Activator.CreateInstance(serviceType, new object[] { model.Application });
            command.Invoke(service, new object[0]);
            results.Add(new Result {
                Case = "explicit-command-reaches-content-barrier",
                ContentExports = model.Exports, ObjectActivations = model.Activations,
                ClipboardCopies = model.ClipboardCopies,
                SelectionWrites = model.SelectionWrites, UnexpectedCalls = model.Unexpected.ToArray(),
                BoundarySatisfied = model.Exports > 0 && model.ClipboardCopies == 0 && model.Unexpected.Count == 0 && model.SelectionWrites == 0
            });
        }
        return results.ToArray();
    }
}
