using System.Xml.Linq;
using VisualTeX.WordVsto;
using Word = Microsoft.Office.Interop.Word;

namespace VisualTeX.VstoFlowAcceptance;

internal static partial class Program
{
    private static void RunNativeNumberStructureAcceptance(string root)
    {
        Directory.CreateDirectory(root);
        XNamespace m = "http://schemas.openxmlformats.org/officeDocument/2006/math";
        var template = XElement.Parse(WordOmmlConverter.BuildWordNativeNumberedOmml(
            "<m:oMath xmlns:m=\"" + m + "\"><m:r><m:t>x+1</m:t></m:r></m:oMath>",
            "Equation", null, 1, "-", "1", "0"));
        XElement Row(XElement xml) => xml.Element(m + "eqArr")!.Element(m + "e")!;
        XElement Label(XElement xml) => Row(xml).Element(m + "d")!.Element(m + "e")!;
        XElement Text(string value) => new(m + "r", new XElement(m + "t", value));
        var cases = new (string Name, Action<XElement> Change, bool Native, bool FieldFree)[]
        {
            ("native-seq", _ => { }, true, false),
            ("literal-label", xml => Label(xml).ReplaceNodes(Text("0-1")), false, true),
            ("no-hash", xml => Row(xml).Elements(m + "r").Last().Element(m + "t")!.Value = "+", false, false),
            ("square-delimiter", xml => Row(xml).Element(m + "d")!.Element(m + "dPr")!.ReplaceNodes(
                new XElement(m + "begChr", new XAttribute(m + "val", "[")),
                new XElement(m + "endChr", new XAttribute(m + "val", "]"))), false, false),
            ("trailing-math", xml => Row(xml).Add(Text("+z")), false, false),
            ("root-sibling", xml => xml.Add(Text("+z")), false, false),
            ("second-row", xml => xml.Element(m + "eqArr")!.Add(new XElement(m + "e", Text("y"))), false, false),
            ("seq-in-body", xml =>
            {
                var field = Label(xml).Nodes().ToArray();
                Label(xml).ReplaceNodes(Text("7"));
                Row(xml).AddFirst(field);
            }, false, false),
        };
        Word.Application? app = null;
        try
        {
            app = CreateWordApplication(visible: false);
            foreach (var test in cases)
            {
                Word.Document? document = null; Word.Range? point = null; Word.Range? inserted = null;
                Word.Range? exact = null; Word.Range? number = null; Word.OMath? math = null;
                try
                {
                    document = app.Documents.Add();
                    var xml = new XElement(template); test.Change(xml);
                    point = document.Range(0, 0);
                    inserted = WordOmmlConverter.ReplaceWithPreparedOmmlDirect(document, point,
                        xml.ToString(SaveOptions.DisableFormatting), display: true, mathFontName: document.OMathFontName);
                    math = document.OMaths[1]; exact = math.Range;
                    File.WriteAllText(Path.Combine(root, test.Name + ".xml"), exact.WordOpenXML);
                    AssertEqual(test.Native, WordNativeOmmlNumbering.TryResolveUnownedNativeStructure(document, exact, out _),
                        test.Name + ": incorrect SEQ ownership classification");
                    number = WordNativeOmmlNumbering.TryResolveNumberContentRange(exact);
                    var host = new WordFormulaHostDescriptor
                    {
                        Kind = WordFormulaHostKind.Omml, DisplayMode = "block",
                        Range = new WordFormulaRangeAddress { Start = exact.Start, End = exact.End, StoryType = exact.StoryType },
                    };
                    AssertEqual(test.FieldFree, WordOmmlHostWriter.HasFieldFreeWordNativeNumberHost(document, host),
                        test.Name + ": incorrect literal-number classification");
                    if (test.Native || test.FieldFree) AssertTrue(number is not null, test.Name + ": missing native label owner");
                    Console.WriteLine("PASS native structural ownership: " + test.Name);
                }
                finally
                {
                    Release(number); Release(exact); Release(math); Release(inserted); Release(point);
                    document?.Close(Word.WdSaveOptions.wdDoNotSaveChanges); Release(document);
                }
            }
        }
        finally { QuitWordApplicationIfOwned(app); Release(app); }
    }
}
