using Microsoft.Win32;
using System.Xml.Linq;
using VisualTeX.WindowsOffice.VstoShared;
using VisualTeX.WordVsto;
using Word = Microsoft.Office.Interop.Word;

namespace VisualTeX.VstoFlowAcceptance;

internal static partial class Program
{
    private static void RunWordMathTypeOpenXmlCloneAcceptance(string artifactRoot)
    {
        Directory.CreateDirectory(artifactRoot);
        var sourceDocx = Path.GetFullPath(Path.Combine(
            AppContext.BaseDirectory,
            "..", "..", "..", "..", "..",
            "artifacts", "mathtype-native-editor",
            "VisualTeX-MathType7-NativeEditor-5f04f8b3545e444a824705446e314ba1.docx"));
        if (!File.Exists(sourceDocx))
            throw new FileNotFoundException("Synchronized MathType source fixture is missing.", sourceDocx);

        var target = Path.Combine(artifactRoot, "MathType7-WordOpenXml-Clone.docx");
        File.Copy(sourceDocx, target, overwrite: true);

        Word.Application? application = null;
        Word.Document? document = null;
        Word.InlineShape? sourceShape = null;
        Word.InlineShape? cloneShape = null;
        Word.InlineShape? rewrittenShape = null;
        Word.InlineShape? unregisteredShape = null;
        Word.Range? sourceRange = null;
        Word.Range? insertion = null;
        Word.OLEFormat? format = null;
        try
        {
            application = CreateWordApplication(visible: false);
            document = application.Documents.Open(target, ReadOnly: false, Visible: false);
            AssertEqual(1, document.InlineShapes.Count,
                "WordOpenXML MathType fixture must begin with one inline OLE object.");
            sourceShape = document.InlineShapes[1];
            sourceRange = sourceShape.Range;

            Console.WriteLine("[MathType WordOpenXML 1/4] Reading the real MathType OLE as Flat OPC WordOpenXML...");
            var wordOpenXml = sourceRange.WordOpenXML;
            File.WriteAllText(Path.Combine(artifactRoot, "source-wordopenxml.xml"), wordOpenXml);
            var package = XDocument.Parse(wordOpenXml, LoadOptions.PreserveWhitespace);
            XNamespace pkg = "http://schemas.microsoft.com/office/2006/xmlPackage";
            var parts = package.Descendants(pkg + "part")
                .Select(element => new
                {
                    Name = (string?)element.Attribute(pkg + "name") ?? string.Empty,
                    ContentType = (string?)element.Attribute(pkg + "contentType") ?? string.Empty,
                    BinaryLength = element.Descendants()
                        .FirstOrDefault(child => child.Name.LocalName == "binaryData")
                        ?.Value.Trim().Length ?? 0,
                })
                .ToArray();
            foreach (var part in parts)
                Console.WriteLine(
                    $"  part name='{part.Name}' contentType='{part.ContentType}' binaryBase64Chars={part.BinaryLength}");
            AssertTrue(parts.Any(part => part.Name.IndexOf("/embeddings/", StringComparison.OrdinalIgnoreCase) >= 0),
                "MathType WordOpenXML did not contain an embedded OLE package part.");
            AssertTrue(parts.Any(part => part.Name.IndexOf("/media/", StringComparison.OrdinalIgnoreCase) >= 0),
                "MathType WordOpenXML did not contain a presentation image package part.");

            Console.WriteLine("[MathType WordOpenXML 2/4] Reinserting the unmodified Flat OPC package without OLE clipboard activation...");
            var before = SnapshotMathTypeProcessIds();
            insertion = document.Range(document.Content.End - 1, document.Content.End - 1);
            insertion.InsertParagraphBefore();
            insertion.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
            insertion.InsertXML(wordOpenXml);
            Thread.Sleep(250);
            var after = SnapshotMathTypeProcessIds();
            var started = after.Except(before).ToArray();
            Console.WriteLine(
                "  MathType processes started by InsertXML: "
                + (started.Length == 0 ? "none" : string.Join(", ", started)));
            AssertEqual(0, started.Length,
                "Word InsertXML unexpectedly started MathType while cloning its existing package parts.");

            Console.WriteLine("[MathType WordOpenXML 3/4] Verifying the cloned object identity and semantic CFB...");
            AssertEqual(2, document.InlineShapes.Count,
                "InsertXML did not materialize a second inline MathType OLE object.");
            cloneShape = document.InlineShapes[2];
            format = cloneShape.OLEFormat;
            var progId = format.ProgID ?? string.Empty;
            Console.WriteLine($"  cloned ProgID='{progId}', size={cloneShape.Width:0.##}x{cloneShape.Height:0.##}.");
            AssertEqual("Equation.DSMT4", progId,
                "WordOpenXML clone changed the MathType OLE ProgID.");
            var sourceLatex = MathMlToLatexConverter.Convert(
                    MathTypeOleStorage.ReadMathMl(MathTypeOleStorage.CaptureCompoundFile(sourceShape)))
                .Trim();
            var clonedLatex = MathMlToLatexConverter.Convert(
                    MathTypeOleStorage.ReadMathMl(MathTypeOleStorage.CaptureCompoundFile(cloneShape)))
                .Trim();
            AssertMathTypeLatexEquivalent(sourceLatex, clonedLatex,
                "WordOpenXML clone changed the embedded MathType MTEF.");
            var clonedPreview = ReadInlineShapeEnhancedMetafile(cloneShape);
            var clonedInk = DescribeEmfInkBounds(clonedPreview);
            AssertTrue(!string.Equals(clonedInk, "empty", StringComparison.Ordinal),
                "Unmodified genuine MathType Flat OPC InsertXML created a visually blank live OLE.");
            Console.WriteLine($"  cloned live preview={clonedInk}; bytes={clonedPreview.Length}.");

            Console.WriteLine("[MathType WordOpenXML 4/5] Rewriting CFB + WMF preview + size entirely inside Flat OPC...");
            var previewSvg = Path.Combine(artifactRoot, "flat-opc-preview.svg");
            File.WriteAllText(
                previewSvg,
                "<svg xmlns=\"http://www.w3.org/2000/svg\" width=\"240\" height=\"80\" viewBox=\"0 0 240 80\"><rect x=\"1\" y=\"1\" width=\"238\" height=\"78\" fill=\"white\"/><text x=\"8\" y=\"54\" font-size=\"42\">VisualTeX OPC</text></svg>");
            var previewEmf = OfficeOlePreview.CreateVectorEmfFromSvg(previewSvg, 240, 80);
            const string targetMathMl =
                "<math xmlns=\"http://www.w3.org/1998/Math/MathML\"><mfrac><mrow><mi>m</mi><mo>+</mo><mn>1</mn></mrow><mi>n</mi></mfrac></math>";
            var sourceFragment = MathTypeWordOpenXml.Read(wordOpenXml);
            var rewrittenCompound = MathTypeOleStorage.RewriteMathTypeCompoundFile(
                sourceFragment.CompoundFile,
                targetMathMl,
                inline: true).CompoundFile;
            var rewrittenWordOpenXml = MathTypeWordOpenXml.Rewrite(
                wordOpenXml,
                rewrittenCompound,
                previewEmf,
                widthPt: 96f,
                heightPt: 32f);
            File.WriteAllText(
                Path.Combine(artifactRoot, "rewritten-wordopenxml.xml"),
                rewrittenWordOpenXml);
            var rewrittenFragment = MathTypeWordOpenXml.Read(rewrittenWordOpenXml);
            var rewrittenLatex = MathMlToLatexConverter.Convert(
                    MathTypeOleStorage.ReadMathMl(rewrittenFragment.CompoundFile))
                .Trim();
            AssertMathTypeLatexEquivalent(
                "\\frac{m+1}{n}",
                rewrittenLatex,
                "Flat OPC rewrite did not carry the new MathType MTEF before Word insertion.");
            var previewDifferenceBeforeInsert = MeasureEmfPixelDifference(
                File.ReadAllBytes(previewEmf),
                rewrittenFragment.PreviewWmf);
            Console.WriteLine(
                $"  rewritten Flat OPC preview pixel difference before InsertXML={previewDifferenceBeforeInsert:0.0000}.");
            AssertTrue(
                previewDifferenceBeforeInsert < 0.08,
                "EMF-to-placeable-WMF conversion changed the VisualTeX preview too much.");
            AssertTrue(
                Math.Abs(rewrittenFragment.WidthPt - 96f) < 0.1f
                && Math.Abs(rewrittenFragment.HeightPt - 32f) < 0.1f,
                "Flat OPC rewrite did not update the VML MathType object size.");

            var rewriteBeforeProcesses = SnapshotMathTypeProcessIds();
            insertion.SetRange(document.Content.End - 1, document.Content.End - 1);
            insertion.InsertParagraphBefore();
            insertion.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
            insertion.InsertXML(rewrittenWordOpenXml);
            Thread.Sleep(250);
            var rewriteAfterProcesses = SnapshotMathTypeProcessIds();
            AssertEqual(
                0,
                rewriteAfterProcesses.Except(rewriteBeforeProcesses).Count(),
                "Rewritten Flat OPC InsertXML unexpectedly started MathType.");
            AssertEqual(3, document.InlineShapes.Count,
                "Rewritten Flat OPC did not materialize a third MathType OLE object.");
            rewrittenShape = document.InlineShapes[3];
            var materializedFragment = MathTypeWordOpenXml.Read(rewrittenShape);
            var materializedLatex = MathMlToLatexConverter.Convert(
                    MathTypeOleStorage.ReadMathMl(materializedFragment.CompoundFile))
                .Trim();
            AssertMathTypeLatexEquivalent(
                "\\frac{m+1}{n}",
                materializedLatex,
                "Word changed the rewritten MTEF while materializing Flat OPC.");
            var materializedPreviewDifference = MeasureEmfPixelDifference(
                File.ReadAllBytes(previewEmf),
                materializedFragment.PreviewWmf);
            Console.WriteLine(
                $"  materialized Word media preview pixel difference={materializedPreviewDifference:0.0000}.");
            AssertTrue(
                materializedPreviewDifference < 0.08,
                "Word did not preserve the VisualTeX WMF preview from rewritten Flat OPC.");
            var rewrittenLivePreview = ReadInlineShapeEnhancedMetafile(rewrittenShape);
            var rewrittenLiveInk = DescribeEmfInkBounds(rewrittenLivePreview);
            AssertTrue(!string.Equals(rewrittenLiveInk, "empty", StringComparison.Ordinal),
                "Rewritten genuine MathType Flat OPC InsertXML created a visually blank live OLE.");
            Console.WriteLine($"  rewritten live preview={rewrittenLiveInk}; bytes={rewrittenLivePreview.Length}.");

            Console.WriteLine("[MathType WordOpenXML inline] Replacing one inline OLE between ordinary prose through hidden staging + FormattedText...");
            Word.Range? leftTextRange = null;
            Word.Range? rightTextRange = null;
            Word.Range? targetObjectRange = null;
            Word.Range? stagedObjectRange = null;
            Word.Range? scratchInsertion = null;
            Word.Document? scratchDocument = null;
            Word.InlineShape? stagedShape = null;
            try
            {
                leftTextRange = sourceShape.Range.Duplicate;
                leftTextRange.Collapse(Word.WdCollapseDirection.wdCollapseStart);
                leftTextRange.Text = "VTLEFT ";
                rightTextRange = sourceShape.Range.Duplicate;
                rightTextRange.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                rightTextRange.Text = " VTRIGHT";

                var inlineFragment = MathTypeWordOpenXml.Read(sourceShape);
                const string inlineTargetMathMl =
                    "<math xmlns=\"http://www.w3.org/1998/Math/MathML\"><mrow><mi>q</mi><mo>+</mo><mn>2</mn></mrow></math>";
                var inlineCompound = MathTypeOleStorage.RewriteMathTypeCompoundFile(
                    inlineFragment.CompoundFile,
                    inlineTargetMathMl,
                    inline: true).CompoundFile;
                var inlineReplacementXml = MathTypeWordOpenXml.Rewrite(
                    inlineFragment.WordOpenXml,
                    inlineCompound,
                    previewEmf,
                    widthPt: 72f,
                    heightPt: 24f);
                var inlineShapeCount = document.InlineShapes.Count;
                var inlineBeforeProcesses = SnapshotMathTypeProcessIds();

                scratchDocument = application.Documents.Add(Visible: false);
                scratchInsertion = scratchDocument.Range(0, 0);
                scratchInsertion.InsertXML(inlineReplacementXml);
                AssertEqual(
                    1,
                    scratchDocument.InlineShapes.Count,
                    "Hidden staging document did not materialize exactly one MathType OLE object.");
                stagedShape = scratchDocument.InlineShapes[1];
                stagedObjectRange = stagedShape.Range.Duplicate;
                targetObjectRange = sourceShape.Range.Duplicate;
                targetObjectRange.FormattedText = stagedObjectRange;
                Thread.Sleep(150);

                var inlineAfterProcesses = SnapshotMathTypeProcessIds();
                AssertEqual(
                    0,
                    inlineAfterProcesses.Except(inlineBeforeProcesses).Count(),
                    "Hidden staging + FormattedText unexpectedly started MathType.");
                AssertEqual(
                    inlineShapeCount,
                    document.InlineShapes.Count,
                    "FormattedText MathType replacement changed the user document OLE object count.");
                var prose = document.Content.Text ?? string.Empty;
                AssertEqual(
                    1,
                    prose.Split(new[] { "VTLEFT" }, StringSplitOptions.None).Length - 1,
                    "FormattedText MathType replacement duplicated or removed the left surrounding prose.");
                AssertEqual(
                    1,
                    prose.Split(new[] { "VTRIGHT" }, StringSplitOptions.None).Length - 1,
                    "FormattedText MathType replacement duplicated or removed the right surrounding prose.");
                var inlineMatchCount = 0;
                for (var index = 1; index <= document.InlineShapes.Count; index++)
                {
                    Word.InlineShape? candidate = null;
                    try
                    {
                        candidate = document.InlineShapes[index];
                        MathTypeWordOpenXml.Fragment candidateFragment;
                        try { candidateFragment = MathTypeWordOpenXml.Read(candidate); }
                        catch { continue; }
                        var candidateLatex = MathMlToLatexConverter.Convert(
                                MathTypeOleStorage.ReadMathMl(candidateFragment.CompoundFile))
                            .Trim();
                        if (NormalizeMathTypeLatex(candidateLatex)
                            == NormalizeMathTypeLatex("q+2"))
                            inlineMatchCount++;
                    }
                    finally { Release(candidate); }
                }
                AssertEqual(
                    1,
                    inlineMatchCount,
                    "FormattedText replacement did not leave exactly one q+2 MathType object between the prose.");
            }
            finally
            {
                Release(targetObjectRange);
                Release(stagedObjectRange);
                Release(stagedShape);
                Release(scratchInsertion);
                if (scratchDocument is not null)
                {
                    try { scratchDocument.Close(Word.WdSaveOptions.wdDoNotSaveChanges); } catch { }
                }
                Release(scratchDocument);
                Release(rightTextRange);
                Release(leftTextRange);
            }

            Console.WriteLine("[MathType WordOpenXML 5/6] Proving InsertXML does not require a registered OLE server...");
            const string unregisteredProgId = "Equation.VisualTeXUnregisteredProbe";
            var unregisteredClsid = new Guid("4D8DD123-0B42-4E80-B744-6B504A3B7F91");
            AssertTrue(
                Type.GetTypeFromProgID(unregisteredProgId, throwOnError: false) is null,
                "The acceptance probe ProgID unexpectedly exists in the Windows registry.");
            using (var registryProbe = Registry.ClassesRoot.OpenSubKey(
                "CLSID\\{" + unregisteredClsid.ToString("D") + "}"))
            {
                AssertTrue(
                    registryProbe is null,
                    "The acceptance probe CLSID unexpectedly exists in the Windows registry.");
            }
            var unregisteredPackage = XDocument.Parse(
                rewrittenWordOpenXml,
                LoadOptions.PreserveWhitespace);
            XNamespace pkgProbe = "http://schemas.microsoft.com/office/2006/xmlPackage";
            XNamespace office = "urn:schemas-microsoft-com:office:office";
            var unregisteredOleObject = unregisteredPackage
                .Descendants(office + "OLEObject")
                .Single();
            unregisteredOleObject.SetAttributeValue("ProgID", unregisteredProgId);
            var unregisteredCompound = MathTypeOleStorage.RewriteCompoundFileRootClsid(
                rewrittenCompound,
                unregisteredClsid);
            AssertEqual(
                unregisteredClsid,
                MathTypeOleStorage.ReadCompoundFileRootClsid(unregisteredCompound),
                "Acceptance probe did not rewrite the CFB root CLSID.");
            var unregisteredEmbeddingPart = unregisteredPackage
                .Descendants(pkgProbe + "part")
                .Single(part => ((string?)part.Attribute(pkgProbe + "name") ?? string.Empty)
                    .IndexOf("/embeddings/", StringComparison.OrdinalIgnoreCase) >= 0);
            var unregisteredBinary = unregisteredEmbeddingPart.Element(pkgProbe + "binaryData")
                ?? throw new InvalidDataException("Acceptance Flat OPC embedding has no binaryData.");
            unregisteredBinary.Value = Convert.ToBase64String(unregisteredCompound);
            var unregisteredWordOpenXml = unregisteredPackage.ToString(SaveOptions.DisableFormatting);
            var unregisteredBeforeProcesses = SnapshotMathTypeProcessIds();
            insertion.SetRange(document.Content.End - 1, document.Content.End - 1);
            insertion.InsertParagraphBefore();
            insertion.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
            insertion.InsertXML(unregisteredWordOpenXml);
            Thread.Sleep(250);
            var unregisteredAfterProcesses = SnapshotMathTypeProcessIds();
            AssertEqual(
                0,
                unregisteredAfterProcesses.Except(unregisteredBeforeProcesses).Count(),
                "InsertXML with an unregistered OLE ProgID unexpectedly started MathType.");
            AssertEqual(4, document.InlineShapes.Count,
                "InsertXML could not materialize an OLE object whose ProgID is unregistered.");
            unregisteredShape = document.InlineShapes[4];
            var unregisteredRange = unregisteredShape.Range;
            try
            {
                var materializedUnregisteredPackage = XDocument.Parse(
                    unregisteredRange.WordOpenXML,
                    LoadOptions.PreserveWhitespace);
                var materializedEmbeddingPart = materializedUnregisteredPackage
                    .Descendants(pkgProbe + "part")
                    .Single(part => ((string?)part.Attribute(pkgProbe + "name") ?? string.Empty)
                        .IndexOf("/embeddings/", StringComparison.OrdinalIgnoreCase) >= 0);
                var materializedBinary = materializedEmbeddingPart.Element(pkgProbe + "binaryData")
                    ?? throw new InvalidDataException(
                        "Materialized unregistered OLE Flat OPC has no binaryData.");
                var materializedCompound = Convert.FromBase64String(
                    new string(materializedBinary.Value
                        .Where(character => !char.IsWhiteSpace(character))
                        .ToArray()));
                AssertEqual(
                    unregisteredClsid,
                    MathTypeOleStorage.ReadCompoundFileRootClsid(materializedCompound),
                    "Word required/rewrote the unregistered OLE CLSID during InsertXML.");
            }
            finally { Release(unregisteredRange); }
            Console.WriteLine(
                "  unregistered ProgID + unregistered CFB CLSID materialized successfully with no OLE server lookup/launch.");

            Console.WriteLine("[MathType WordOpenXML 6/6] Saving/reopening clone and rewritten objects...");
            document.Save();
            document.Close(Word.WdSaveOptions.wdSaveChanges);
            Release(format); format = null;
            Release(unregisteredShape); unregisteredShape = null;
            Release(rewrittenShape); rewrittenShape = null;
            Release(cloneShape); cloneShape = null;
            Release(sourceShape); sourceShape = null;
            Release(sourceRange); sourceRange = null;
            Release(insertion); insertion = null;
            Release(document); document = null;
            application.Quit(Word.WdSaveOptions.wdDoNotSaveChanges);
            Release(application); application = null;

            var embeddings = ReadDocxOleEmbeddings(target);
            AssertEqual(4, embeddings.Count,
                "Saved WordOpenXML document did not retain all four OLE package parts.");
            AssertTrue(
                embeddings.Any(compound =>
                    MathTypeOleStorage.ReadCompoundFileRootClsid(compound) == unregisteredClsid),
                "Saved Word document did not preserve the unregistered OLE root CLSID.");
            Console.WriteLine(
                "MathType WordOpenXML acceptance passed: Word cloned Equation.DSMT4 and materialized VisualTeX-rewritten MTEF + WMF preview + size from Flat OPC without launching MathType.");
        }
        finally
        {
            Release(format);
            Release(unregisteredShape);
            Release(rewrittenShape);
            Release(cloneShape);
            Release(sourceShape);
            Release(sourceRange);
            Release(insertion);
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
