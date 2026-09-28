using System;
using System.Collections.Generic;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Runtime.Remoting.Messaging;
using System.Runtime.Remoting.Proxies;
using System.Threading;
using Forms = System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;

// Uses the real compiled ThisAddIn/WordFormulaService, but only managed Office
// interface doubles. No Office activation, user documents, registry changes or
// native renderer are involved. In particular, getEnabled is tested through the
// actual Ribbon entry point rather than by invoking only its command helper.
public static class WordRibbonObservationChecks
{
    public sealed class Result
    {
        public string Case;
        public bool BoundarySatisfied;
        public bool BehaviorSatisfied;
        public int ContentExports, ObjectActivations, ClipboardCopies, SelectionWrites;
        public string[] UnexpectedCalls;
        public string[] InvalidatedControls;
        public string Detail;
    }

    public sealed class RibbonSpy
    {
        public readonly List<string> Invalidated = new List<string>();
        public void InvalidateControl(string id) { Invalidated.Add(id); }
        public bool FontInvalidated
        {
            get
            {
                return Invalidated.Contains("VisualTeX.WordVsto.FontSize")
                    && Invalidated.Contains("VisualTeX.WordVsto.FontSizeDecrease")
                    && Invalidated.Contains("VisualTeX.WordVsto.FontSizeIncrease");
            }
        }
    }

    private sealed class InterfaceDouble : RealProxy
    {
        private readonly Func<MethodInfo, object[], object> handler;
        internal InterfaceDouble(Type type, Func<MethodInfo, object[], object> invoke) : base(type) { handler = invoke; }
        public override IMessage Invoke(IMessage message)
        {
            var call = (IMethodCallMessage)message;
            try
            {
                object[] args = call.Args;
                object result = handler((MethodInfo)call.MethodBase, args);
                return new ReturnMessage(result, args, args.Length, call.LogicalCallContext, call);
            }
            catch (Exception error) { return new ReturnMessage(error, call); }
        }
    }

    private sealed class Model
    {
        internal readonly List<string> Unexpected = new List<string>();
        internal int Exports, Activations, Copies, Writes, Calls, GeometryReads;
        internal string ExportPayload;
        internal int Start = 10, End = 11;
        internal readonly string ProgId;
        internal readonly bool MissingRegistration;
        internal readonly Word.WdInlineShapeType ShapeType;
        internal Word.Application Application;
        internal Word.Document Document;
        internal Word.Window Window;
        internal Word.Selection Selection;
        internal Word.InlineShape Shape;
        internal Word.OLEFormat OleFormat;
        internal Word.OMaths EmptyMaths;
        internal Word.Tables EmptyTables;
        internal Word.Fields EmptyFields;

        internal Model(string progId, bool missingRegistration, bool picture)
        {
            ProgId = progId;
            MissingRegistration = missingRegistration;
            ShapeType = picture ? Word.WdInlineShapeType.wdInlineShapePicture : Word.WdInlineShapeType.wdInlineShapeEmbeddedOLEObject;
            EmptyMaths = Make<Word.OMaths>("OMaths", delegate(MethodInfo m, object[] a) {
                if (m.Name == "get_Count") return 0;
                return Unknown("OMaths", m);
            });
            EmptyTables = Make<Word.Tables>("Tables", delegate(MethodInfo m, object[] a) {
                if (m.Name == "get_Count") return 0;
                return Unknown("Tables", m);
            });
            EmptyFields = Make<Word.Fields>("Fields", delegate(MethodInfo m, object[] a) {
                if (m.Name == "get_Count") return 0;
                return Unknown("Fields", m);
            });
            OleFormat = Make<Word.OLEFormat>("OLEFormat", delegate(MethodInfo m, object[] a) {
                if (m.Name == "get_ProgID")
                {
                    if (MissingRegistration) throw new COMException("Controlled missing registration", unchecked((int)0x80040154));
                    return ProgId;
                }
                if (m.Name == "get_Object" || m.Name == "Activate" || m.Name == "DoVerb")
                { Activations++; throw new InvalidOperationException("OLE activation barrier"); }
                return Unknown("OLEFormat", m);
            });
            Shape = Make<Word.InlineShape>("Shape", delegate(MethodInfo m, object[] a) {
                switch (m.Name)
                {
                    case "get_Type": return ShapeType;
                    case "get_OLEFormat": return OleFormat;
                    case "get_Range": return Range(10, 11);
                    case "get_AlternativeText": case "get_Title": return String.Empty;
                    default: return Unknown("Shape", m);
                }
            });
            Selection = Make<Word.Selection>("Selection", delegate(MethodInfo m, object[] a) {
                switch (m.Name)
                {
                    case "get_Range": return Range(Start, End);
                    case "get_Start": return Start;
                    case "get_End": return End;
                    case "get_Document": return Document;
                    case "SetRange": case "Collapse": case "TypeText":
                        Writes++; throw new InvalidOperationException("Selection mutation barrier");
                    default: return Unknown("Selection", m);
                }
            });
            Window = Make<Word.Window>("Window", delegate(MethodInfo m, object[] a) {
                if (m.Name == "get_Hwnd") return 1;
                if (m.Name == "GetPoint") { GeometryReads++; a[0] = 100; a[1] = 100; a[2] = 40; a[3] = 20; return null; }
                return Unknown("Window", m);
            });
            Document = Make<Word.Document>("Document", delegate(MethodInfo m, object[] a) {
                if (m.Name == "get_Application") return Application;
                if (m.Name == "get_Content") return Range(0, 100);
                if (m.Name == "Range") return Range(Convert.ToInt32(a[0]), Convert.ToInt32(a[1]));
                return Unknown("Document", m);
            });
            Application = Make<Word.Application>("Application", delegate(MethodInfo m, object[] a) {
                if (m.Name == "get_Selection") return Selection;
                if (m.Name == "get_ActiveDocument") return Document;
                if (m.Name == "get_ActiveWindow") return Window;
                return Unknown("Application", m);
            });
        }

        internal void Prose() { Start = End = 30; }
        internal void CaretAfterFormula() { Start = End = 11; }
        private Word.InlineShapes Shapes(bool contains)
        {
            return Make<Word.InlineShapes>("InlineShapes", delegate(MethodInfo m, object[] a) {
                if (m.Name == "get_Count") return contains ? 1 : 0;
                if (m.Name == "get_Item" && contains && Convert.ToInt32(a[0]) == 1) return Shape;
                return Unknown("InlineShapes", m);
            });
        }
        private Word.Paragraphs Paragraphs(int start)
        {
            Word.Paragraph paragraph = Make<Word.Paragraph>("Paragraph", delegate(MethodInfo m, object[] a) {
                if (m.Name == "get_Range") return start < 25 ? Range(0, 25) : Range(25, 100);
                return Unknown("Paragraph", m);
            });
            return Make<Word.Paragraphs>("Paragraphs", delegate(MethodInfo m, object[] a) {
                if (m.Name == "get_Count") return 1;
                if (m.Name == "get_Item" && Convert.ToInt32(a[0]) == 1) return paragraph;
                return Unknown("Paragraphs", m);
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
                    case "get_Document": return Document;
                    case "get_InlineShapes": return Shapes(start <= 10 && end >= 11);
                    case "get_OMaths": return EmptyMaths;
                    case "get_Tables": return EmptyTables;
                    case "get_Fields": return EmptyFields;
                    case "get_Paragraphs": return Paragraphs(start);
                    case "get_Text": return String.Empty;
                    case "MoveStart": start = Math.Max(0, start + Convert.ToInt32(a[1])); return Convert.ToInt32(a[1]);
                    case "MoveEnd": end = Math.Min(100, end + Convert.ToInt32(a[1])); return Convert.ToInt32(a[1]);
                    case "get_WordOpenXML":
                        Exports++;
                        if (ExportPayload != null) return ExportPayload;
                        throw new InvalidOperationException("Content export barrier");
                    case "Copy":
                        Copies++; throw new InvalidOperationException("Clipboard capture barrier");
                    default: return Unknown("Range", m);
                }
            });
        }
        private T Make<T>(string label, Func<MethodInfo, object[], object> invoke) where T : class
        {
            return (T)new InterfaceDouble(typeof(T), delegate(MethodInfo m, object[] a) {
                if (m.Name == "GetType" && a.Length == 0) return typeof(T);
                if (m.DeclaringType == typeof(object))
                {
                    if (m.Name == "ToString") return "Managed interface double: " + label;
                    if (m.Name == "GetHashCode") return label.GetHashCode();
                    if (m.Name == "Equals") return false;
                }
                Calls++;
                return invoke(m, a);
            }).GetTransparentProxy();
        }
        private object Unknown(string label, MethodInfo method)
        {
            Unexpected.Add(label + "." + method.Name);
            if (method.Name.StartsWith("set_", StringComparison.Ordinal)) Writes++;
            throw new NotSupportedException("Unmodelled call: " + label + "." + method.Name);
        }
    }

    private sealed class Harness : IDisposable
    {
        internal readonly object Addin, Service;
        internal readonly RibbonSpy Ribbon = new RibbonSpy();
        private readonly object dispatcher;
        internal Harness(Assembly assembly, Model model)
        {
            Addin = Activator.CreateInstance(assembly.GetType("VisualTeX.WordVsto.ThisAddIn", true));
            Service = Activator.CreateInstance(assembly.GetType("VisualTeX.WordVsto.WordFormulaService", true), new object[] { model.Application });
            dispatcher = Activator.CreateInstance(assembly.GetType("VisualTeX.WordVsto.OfficeUiDispatcher", true));
            Set("_application", model.Application);
            Set("_formulaService", Service);
            Set("_dispatcher", dispatcher);
            Set("_ribbonUi", Ribbon);
        }
        internal void Set(string field, object value) { Addin.GetType().GetField(field, BindingFlags.Instance | BindingFlags.NonPublic).SetValue(Addin, value); }
        internal object Read(string field) { return Addin.GetType().GetField(field, BindingFlags.Instance | BindingFlags.NonPublic).GetValue(Addin); }
        internal object Invoke(string method, params object[] args)
        {
            return Addin.GetType().GetMethod(method, BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic).Invoke(Addin, args);
        }
        internal bool Enabled() { return (bool)Invoke("GetFormulaFontSizeEnabled", new object[] { null }); }
        internal string Text() { return (string)Invoke("GetFormulaFontSizeText", new object[] { null }); }
        internal void SelectionChanged(Model model) { Invoke("OnWindowSelectionChange", model.Selection); Pump(); }
        internal void Activated(Model model) { Invoke("OnWindowActivate", model.Document, model.Window); Pump(); }
        // Pumps only this process's hidden WinForms dispatcher. No Word instance
        // exists in this test, and there is no external document polling.
        internal static void Pump() { Forms.Application.DoEvents(); Forms.Application.DoEvents(); }
        public void Dispose() { ((IDisposable)dispatcher).Dispose(); }
    }

    private static Result Observe(string name, Model model, RibbonSpy ribbon, bool behavior, string detail, bool explicitCommand)
    {
        return new Result {
            Case = name,
            ContentExports = model.Exports,
            ObjectActivations = model.Activations,
            ClipboardCopies = model.Copies,
            SelectionWrites = model.Writes,
            UnexpectedCalls = model.Unexpected.ToArray(),
            InvalidatedControls = ribbon.Invalidated.ToArray(),
            BoundarySatisfied = (explicitCommand ? model.Exports > 0 : model.Exports == 0)
                && model.Activations == 0 && model.Copies == 0 && model.Writes == 0 && model.Unexpected.Count == 0,
            BehaviorSatisfied = behavior,
            Detail = detail
        };
    }

    public static Result[] Run(string assemblyPath)
    {
        if (Thread.CurrentThread.GetApartmentState() != ApartmentState.STA) throw new InvalidOperationException("Run with Windows PowerShell -STA.");
        var assembly = Assembly.LoadFrom(assemblyPath);
        var results = new List<Result>();
        foreach (string identity in new[] { "Equation.DSMT4", "Unknown.ThirdParty.1", "missing-registration" })
        {
            var model = new Model(identity, identity == "missing-registration", false);
            using (var h = new Harness(assembly, model))
            {
                h.SelectionChanged(model);
                bool enabled = h.Enabled();
                string text = h.Text();
                bool claimed = (bool)h.Read("_nativeOleTargetActive");
                results.Add(Observe(identity + ":ribbon-command-reachable-without-content", model, h.Ribbon,
                    enabled && String.IsNullOrEmpty(text) && (identity == "Equation.DSMT4" ? claimed : !claimed),
                    "Enabled=" + enabled + "; Text=" + text + "; GestureTargetCached=" + claimed + "; GeometryReads=" + model.GeometryReads, false));
            }
            var caretModel = new Model(identity, identity == "missing-registration", false);
            caretModel.CaretAfterFormula();
            using (var h = new Harness(assembly, caretModel))
            {
                h.SelectionChanged(caretModel);
                results.Add(Observe(identity + ":full-collapsed-selection-event", caretModel, h.Ribbon,
                    true, "Complete event and queued dispatcher callbacks", false));
            }
            var commandModel = new Model(identity, identity == "missing-registration", false);
            using (var h = new Harness(assembly, commandModel))
            {
                bool enabled = h.Enabled();
                if (enabled)
                    h.Service.GetType().GetMethod("ReadSelectedFormulaFontSizeForCommand", BindingFlags.Instance | BindingFlags.NonPublic)
                        .Invoke(h.Service, new object[0]);
                results.Add(Observe(identity + ":enabled-command-reaches-content-barrier", commandModel, h.Ribbon,
                    enabled, "Checks Ribbon reachability BEFORE calling command helper; not a font-edit success test", true));
            }
        }
        foreach (bool activation in new[] { false, true })
        {
            var model = new Model("Equation.DSMT4", false, false);
            using (var h = new Harness(assembly, model))
            {
                h.SelectionChanged(model);
                bool before = h.Enabled();
                string text = h.Text();
                h.Ribbon.Invalidated.Clear();
                model.Prose();
                if (activation) h.Activated(model); else h.SelectionChanged(model);
                bool after = h.Enabled();
                results.Add(Observe(activation ? "window-activation-to-prose-invalidates-font" : "formula-to-prose-invalidates-unknown-font",
                    model, h.Ribbon, before && !after && String.IsNullOrEmpty(text) && h.Ribbon.FontInvalidated,
                    "Before=" + before + "; After=" + after + "; FontInvalidated=" + h.Ribbon.FontInvalidated, false));
            }
        }
        var prose = new Model("Equation.DSMT4", false, false);
        prose.Prose();
        using (var h = new Harness(assembly, prose))
        {
            h.SelectionChanged(prose);
            bool first = h.Enabled();
            h.Ribbon.Invalidated.Clear();
            prose.Start = prose.End = 40;
            h.SelectionChanged(prose);
            results.Add(Observe("prose-to-prose-no-redundant-invalidation", prose, h.Ribbon,
                !first && !h.Enabled() && h.Ribbon.Invalidated.Count == 0, "No formula target in either position", false));
        }
        var mutating = new Model("Equation.DSMT4", false, false);
        using (var h = new Harness(assembly, mutating))
        {
            h.Set("_formulaFormatMutationDepth", 1);
            h.SelectionChanged(mutating);
            bool enabled = h.Enabled();
            string text = h.Text();
            results.Add(Observe("own-mutation-no-object-model-observation", mutating, h.Ribbon,
                !enabled && String.IsNullOrEmpty(text) && mutating.Calls == 0, "Object-model calls=" + mutating.Calls, false));
        }
        var unrelated = new Model("Excel.Sheet.12", false, false);
        // A successfully returned Flat OPC document with no MathType OLE payload
        // must be rejected by the explicit semantic reader, not mutated as if the
        // Ribbon's candidate state had proved that the object was a formula.
        unrelated.ExportPayload = "<pkg:package xmlns:pkg=\"http://schemas.microsoft.com/office/2006/xmlPackage\"><pkg:part pkg:name=\"/word/document.xml\" pkg:contentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml\"><pkg:xmlData><w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:body><w:p><w:r><w:t>Not a formula</w:t></w:r></w:p></w:body></w:document></pkg:xmlData></pkg:part></pkg:package>";
        using (var h = new Harness(assembly, unrelated))
        {
            h.SelectionChanged(unrelated);
            bool enabled = h.Enabled();
            object value = h.Service.GetType().GetMethod("ReadSelectedFormulaFontSizeForCommand", BindingFlags.Instance | BindingFlags.NonPublic)
                .Invoke(h.Service, new object[0]);
            results.Add(Observe("unrelated-ole-inspection-does-not-authorize-mutation", unrelated, h.Ribbon,
                enabled && value == null && !(bool)h.Read("_nativeOleTargetActive"),
                "Unrelated content rejected; native gesture not claimed", true));
        }
        var picture = new Model("", false, true);
        using (var h = new Harness(assembly, picture))
        {
            h.SelectionChanged(picture);
            results.Add(Observe("plain-picture-is-not-an-ole-command-candidate", picture, h.Ribbon,
                !h.Enabled() && !(bool)h.Read("_nativeOleTargetActive"), "Ordinary image remains outside formula commands", false));
        }
        return results.ToArray();
    }
}
