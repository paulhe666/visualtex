using System.Text.Json;
using System.Xml.Linq;
using VisualTeX.WordVsto;
using Word = Microsoft.Office.Interop.Word;

namespace VisualTeX.VstoFlowAcceptance;
internal static partial class Program
{
    private static void RunNativeLiteralPrefixProbe(string root)
    {
        Directory.CreateDirectory(root);
        var observations = new List<object>();
        Word.Application? app = null;
        var previous = WordEquationNumbering.GetDefaultEquationNumberFormatId();
        try
        {
            app = CreateWordApplication(visible: false);
            foreach (var strategy in new[] { "whole", "characters" })
            {
                Word.Document? doc = null;
                try
                {
                    doc = app.Documents.Add();
                    doc.Content.Text = "4 Chapter\r4.2 Section\r\r";
                    for (var i = 1; i <= 2; i++)
                    {
                        var heading = doc.Paragraphs[i].Range;
                        try { object style = i == 1 ? Word.WdBuiltinStyle.wdStyleHeading1 : Word.WdBuiltinStyle.wdStyleHeading2; heading.set_Style(ref style); }
                        finally { Release(heading); }
                    }
                    InsertNativeNumberedFixture(doc, doc.Paragraphs[3].Range, "x+1", 0, "", "1", 1);
                    var service = new WordFormulaService(app);
                    service.SetEquationNumberFormat(EquationNumberFormat.Heading1DashId);
                    void Observe(string stage)
                    {
                        var math = doc!.OMaths[1]; var exact = math.Range;
                        try
                        {
                            var xml = exact.WordOpenXML;
                            XNamespace m = "http://schemas.openxmlformats.org/officeDocument/2006/math";
                            var normal = XElement.Parse(WordOmmlConverter.ExtractSingleOMath(xml)).Descendants(m + "nor").Count();
                            var number = WordNativeOmmlNumbering.ResolveNumberContentRange(exact);
                            try { observations.Add(new { Strategy = strategy, Stage = stage, NormalRuns = normal, Live = number.Text, XmlLabel = WordOmmlConverter.ReadWordNativeEquationNumberText(xml) }); }
                            finally { Release(number); }
                            File.WriteAllText(Path.Combine(root, strategy + "-" + stage + ".xml"), xml);
                            Console.WriteLine(strategy + " " + stage + " normal=" + normal);
                        }
                        finally { Release(exact); Release(math); }
                    }
                    Observe("after-service");
                    for (var pass = 1; pass <= 2; pass++)
                    {
                        var math = doc.OMaths[1]; var exact = math.Range;
                        Word.Range? prefix = null; Word.Fields? fields = null;
                        try
                        {
                            prefix = WordNativeOmmlNumbering.ResolveNumberContentRange(exact);
                            fields = prefix.Fields;
                            for (var i = 1; i <= fields.Count; i++)
                            {
                                var f = fields[i]; var code = f.Code;
                                try { if (f.Type == Word.WdFieldType.wdFieldSequence && WordNumberingFieldCollection.IsCurrentField(f)) { prefix.End = code.Start - 1; break; } }
                                finally { Release(code); Release(f); }
                            }
                            var functions = math.Functions;
                            try
                            {
                                if (strategy == "whole") { var normal = functions.Add(prefix, Word.WdOMathFunctionType.wdOMathFunctionNormalText); Release(normal); }
                                else
                                {
                                    var start = prefix.Start; var end = prefix.End;
                                    for (var i = start; i < end; i++)
                                    {
                                        var atom = doc.Range(i, i + 1);
                                        try { var normal = functions.Add(atom, Word.WdOMathFunctionType.wdOMathFunctionNormalText); Release(normal); }
                                        finally { Release(atom); }
                                    }
                                }
                            }
                            finally { Release(functions); }
                        }
                        finally { Release(fields); Release(prefix); Release(exact); Release(math); }
                        Observe("after-normal-" + pass);
                    }
                    service.UpdateEquationNumbers(); Observe("after-update");
                    var path = Path.Combine(root, strategy + ".docx");
                    doc.SaveAs2(path, Word.WdSaveFormat.wdFormatXMLDocument);
                    doc.Close(Word.WdSaveOptions.wdDoNotSaveChanges); Release(doc); doc = null;
                    doc = app.Documents.Open(path, ReadOnly: false, Visible: false, AddToRecentFiles: false);
                    Observe("after-reopen");
                }
                finally { try { doc?.Close(Word.WdSaveOptions.wdDoNotSaveChanges); } catch { } Release(doc); }
            }
            File.WriteAllText(Path.Combine(root, "observations.json"), JsonSerializer.Serialize(observations, new JsonSerializerOptions { WriteIndented = true }));
        }
        finally { try { QuitWordApplicationIfOwned(app); } catch { } Release(app); WordEquationNumbering.SetDefaultEquationNumberFormatPreference(previous); ForceComCleanup(); }
    }
}
