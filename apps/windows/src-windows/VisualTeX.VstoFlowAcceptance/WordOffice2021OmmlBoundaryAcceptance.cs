using Extensibility;
using Microsoft.Office.Core;
using VisualTeX.WindowsOffice.Contracts;
using Word = Microsoft.Office.Interop.Word;

namespace VisualTeX.VstoFlowAcceptance;

internal static partial class Program
{
    private static void RunWordOffice2021OmmlBoundaryAcceptance(
        VisualTeXSessionClient client,
        string artifactRoot)
    {
        Directory.CreateDirectory(artifactRoot);
        Word.Application? application = null;
        Word.Document? document = null;
        COMAddIns? installedAddIns = null;
        COMAddIn? installedAddIn = null;
        VisualTeX.WordVsto.ThisAddIn? addIn = null;
        Array custom = Array.Empty<object>();
        try
        {
            application = CreateWordApplication(visible: false);
            Console.WriteLine($"Word OMML boundary acceptance: version={application.Version}, build={application.Build}.");
            installedAddIns = application.COMAddIns;
            try
            {
                object addInIndex = "VisualTeX.WordVsto";
                installedAddIn = installedAddIns.Item(ref addInIndex);
                if (installedAddIn.Connect)
                    installedAddIn.Connect = false;
            }
            catch
            {
                Release(installedAddIn);
                installedAddIn = null;
            }

            document = application.Documents.Add();
            addIn = new VisualTeX.WordVsto.ThisAddIn();
            addIn.OnConnection(application, ext_ConnectMode.ext_cm_AfterStartup, addIn, ref custom);

            void InsertInline(string latex, string mathMl)
            {
                var existing = SnapshotSessionIds();
                addIn.OnInsertInlineOmml(new object());
                var sessionId = WaitForNewSession(existing, "word", TimeSpan.FromSeconds(30));
                var session = client.GetSessionAsync(sessionId, CancellationToken.None)
                    .GetAwaiter().GetResult();
                Commit(
                    client,
                    session,
                    "inline",
                    FormulaOleContract.WordOmmlMode,
                    latex,
                    numbered: false,
                    mathMl: mathMl);
                var final = WaitForTerminal(client, sessionId, TimeSpan.FromSeconds(45));
                AssertEqual("completed", final.Status,
                    final.Error ?? "Word 2021 inline OMML insertion failed.");
                client.CloseEditorAsync(sessionId, CancellationToken.None).GetAwaiter().GetResult();
                WaitForAddInIdle(addIn, TimeSpan.FromSeconds(10));
            }

            void EditSelectedInline(string latex, string mathMl)
            {
                var existing = SnapshotSessionIds();
                addIn.OnEditSelected(new object());
                var sessionId = WaitForNewSession(existing, "word", TimeSpan.FromSeconds(30));
                var session = WaitForUnchangedEditorReady(
                    client,
                    sessionId,
                    TimeSpan.FromSeconds(15));
                AssertEqual(FormulaOleContract.WordOmmlMode, session.ObjectMode,
                    "Word 2021 OMML boundary edit opened the wrong object mode.");
                Commit(
                    client,
                    session,
                    "inline",
                    FormulaOleContract.WordOmmlMode,
                    latex,
                    numbered: false,
                    mathMl: mathMl);
                var final = WaitForTerminal(client, sessionId, TimeSpan.FromSeconds(45));
                AssertEqual("completed", final.Status,
                    final.Error ?? "Word 2021 inline OMML edit failed.");
                client.CloseEditorAsync(sessionId, CancellationToken.None).GetAwaiter().GetResult();
                WaitForAddInIdle(addIn, TimeSpan.FromSeconds(10));
            }

            string DescribeParagraphCharacters(int paragraphIndex)
            {
                Word.Range? paragraphRange = null;
                var parts = new List<string>();
                try
                {
                    paragraphRange = document.Paragraphs[paragraphIndex].Range;
                    for (var position = paragraphRange.Start; position < paragraphRange.End; position++)
                    {
                        Word.Range? character = null;
                        Word.Font? font = null;
                        Word.OMaths? characterMaths = null;
                        try
                        {
                            character = document.Range(position, position + 1);
                            font = character.Font;
                            characterMaths = character.OMaths;
                            var value = character.Text ?? string.Empty;
                            var codepoint = value.Length == 0 ? -1 : value[0];
                            parts.Add($"{position}:U+{codepoint:X4}:H={font.Hidden}:M={characterMaths.Count}:{value.Replace("\r", "<P>")}");
                        }
                        finally
                        {
                            Release(characterMaths);
                            Release(font);
                            Release(character);
                        }
                    }
                    return string.Join(" | ", parts);
                }
                finally { Release(paragraphRange); }
            }

            Console.WriteLine("[Word 2021 OMML boundary 1/2] Editing with following prose...");
            application.Selection.TypeText("prefix ");
            InsertInline(
                "x+y",
                "<math xmlns=\"http://www.w3.org/1998/Math/MathML\"><mi>x</mi><mo>+</mo><mi>y</mi></math>");
            application.Selection.TypeText(" suffix");
            Console.WriteLine("  before edit: " + DescribeParagraphCharacters(1));
            document.OMaths[1].Range.Select();
            EditSelectedInline(
                "x+y1",
                "<math xmlns=\"http://www.w3.org/1998/Math/MathML\"><mi>x</mi><mo>+</mo><mi>y</mi><mn>1</mn></math>");
            Console.WriteLine("  after edit:  " + DescribeParagraphCharacters(1));
            Word.OMath? editedMath = null;
            Word.Range? editedMathRange = null;
            Word.Range? firstParagraphRange = null;
            Word.Range? following = null;
            try
            {
                editedMath = document.OMaths[1];
                editedMathRange = editedMath.Range;
                firstParagraphRange = document.Paragraphs[1].Range;
                following = document.Range(editedMathRange.End, firstParagraphRange.End - 1);
                AssertEqual(" suffix", following.Text ?? string.Empty,
                    "Word 2021 inline OMML edit changed following prose or left guard spaces outside the equation.");
                AssertEqual(0, following.Font.Hidden,
                    "Word 2021 inline OMML edit left following prose hidden.");
            }
            finally
            {
                Release(following);
                Release(firstParagraphRange);
                Release(editedMathRange);
                Release(editedMath);
            }

            Console.WriteLine("[Word 2021 OMML boundary 2/2] Editing at paragraph end and typing ordinary prose...");
            application.Selection.EndKey(Word.WdUnits.wdStory);
            application.Selection.TypeParagraph();
            application.Selection.TypeText("end ");
            InsertInline(
                "a+b",
                "<math xmlns=\"http://www.w3.org/1998/Math/MathML\"><mi>a</mi><mo>+</mo><mi>b</mi></math>");
            document.OMaths[2].Range.Select();
            EditSelectedInline(
                "a+b1",
                "<math xmlns=\"http://www.w3.org/1998/Math/MathML\"><mi>a</mi><mo>+</mo><mi>b</mi><mn>1</mn></math>");
            application.Selection.TypeText(" tail");

            AssertEqual(2, document.OMaths.Count,
                "Typing after the edited paragraph-end OMML changed the equation inventory.");
            Word.OMath? finalMath = null;
            Word.Range? finalMathRange = null;
            try
            {
                finalMath = document.OMaths[2];
                finalMathRange = finalMath.Range;
                var finalMathText = finalMathRange.Text ?? string.Empty;
                AssertTrue(finalMathText.IndexOf("tail", StringComparison.Ordinal) < 0,
                    "Word 2021 absorbed ordinary text into the edited inline OMath.");
            }
            finally
            {
                Release(finalMathRange);
                Release(finalMath);
            }
            Word.Range? secondParagraphRange = null;
            Word.Range? finalFollowingText = null;
            Word.Paragraphs? finalMathParagraphs = null;
            try
            {
                finalMath = document.OMaths[2];
                finalMathRange = finalMath.Range;
                finalMathParagraphs = finalMathRange.Paragraphs;
                AssertEqual(
                    1,
                    finalMathParagraphs.Count,
                    "Word 2021 edited inline OMML escaped its owning paragraph.");
                secondParagraphRange = finalMathParagraphs[1].Range;
                finalFollowingText = document.Range(
                    finalMathRange.End,
                    secondParagraphRange.End - 1);
                AssertEqual(" tail", finalFollowingText.Text ?? string.Empty,
                    "Word 2021 did not leave the caret on the ordinary-text side after inline OMML edit.");
                AssertEqual(0, finalFollowingText.Font.Hidden,
                    "Text typed after the edited Word 2021 inline OMML is hidden.");
            }
            finally
            {
                Release(finalMathParagraphs);
                Release(finalFollowingText);
                Release(secondParagraphRange);
                Release(finalMathRange);
                Release(finalMath);
                finalMathRange = null;
                finalMath = null;
            }

            var path = Path.Combine(artifactRoot, "Word-2021-OMML-Boundary.docx");
            document.SaveAs2(path, Word.WdSaveFormat.wdFormatXMLDocument);
            Console.WriteLine($"Word 2021 inline OMML boundary acceptance passed. Artifact: {path}");
        }
        finally
        {
            if (addIn is not null)
            {
                try { addIn.OnDisconnection(ext_DisconnectMode.ext_dm_UserClosed, ref custom); } catch { }
            }
            if (installedAddIn is not null)
            {
                try { installedAddIn.Connect = true; } catch { }
            }
            Release(installedAddIn);
            Release(installedAddIns);
            if (document is not null)
            {
                try { document.Close(Word.WdSaveOptions.wdDoNotSaveChanges); } catch { }
            }
            Release(document);
            try { QuitWordApplicationIfOwned(application); } catch { }
            Release(application);
            ForceComCleanup();
        }
    }
}
