using System.Drawing;
using System.Drawing.Imaging;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text.Json;
using System.Xml.Linq;
using VisualTeX.WindowsOffice.Contracts;
using VisualTeX.WordVsto;
using Word = Microsoft.Office.Interop.Word;

namespace VisualTeX.VstoFlowAcceptance;

internal static partial class Program
{
    // This is a diagnostic acceptance, not a source-Position preservation rule.
    // MathType and VisualTeX have different picture bounds. Only real Word pixels
    // relative to unchanged prose can establish whether the whole equation sank.
    private sealed class InlineBaselineSeed
    {
        public int SourceOrdinal { get; set; }
        public string Family { get; set; } = "";
        public string Latex { get; set; } = "";
        public string MathMl { get; set; } = "";
        public double FontSizePt { get; set; }
        public int SeedObjectPositionPt { get; set; }
        public float SeedWidthPt { get; set; }
        public float SeedHeightPt { get; set; }
        internal string FlatOpc = "";
    }

    private sealed class InlineBaselineSample
    {
        public int Ordinal { get; set; }
        public string Family { get; set; } = "";
        public string Latex { get; set; } = "";
        public string Phase { get; set; } = "";
        public string Alignment { get; set; } = "";
        public int AlignmentValue { get; set; }
        public int ObjectCharacterCode { get; set; }
        public int ObjectPositionPt { get; set; }
        public int HostInstructionPositionPt { get; set; }
        public int PrefixPositionPt { get; set; }
        public int SuffixPositionPt { get; set; }
        public float WidthPt { get; set; }
        public float HeightPt { get; set; }
        public double SemanticFontSizePt { get; set; }
        public int? CapturedSourcePositionPt { get; set; }
        public float? CapturedSourceBottomWhitespacePt { get; set; }
        public double? ExportHeightPx { get; set; }
        public double? ExportBaselinePx { get; set; }
        public int? ExpectedPolicyPositionPt { get; set; }
        public Rectangle ObjectScreen { get; set; }
        public Rectangle PrefixScreen { get; set; }
        public Rectangle SuffixScreen { get; set; }
        public Rectangle CaptureScreen { get; set; }
        public Rectangle FormulaInk { get; set; }
        public Rectangle PrefixInk { get; set; }
        public Rectangle SuffixInk { get; set; }
        public double ScreenPixelsPerPoint { get; set; }
        public double ProseBaselineYPx { get; set; }
        public double? FormulaEqualsAxisYPx { get; set; }
        public int EqualsAxisCandidateCount { get; set; }
        public double InkTopFromProseBaselinePt { get; set; }
        public double InkBottomFromProseBaselinePt { get; set; }
        public double? EqualsAxisFromProseBaselinePt { get; set; }
        public string Screenshot { get; set; } = "";
    }

    private static class InlineBaselineCaptureNative
    {
        [DllImport("user32.dll")] internal static extern IntPtr GetForegroundWindow();
    }

    private static void RunWordConversionInlineBaselineAcceptance(string artifactRoot)
    {
        AssertTrue(!AttachActiveWord, "Inline baseline acceptance refuses the user's active Word instance.");
        var seedPath = Environment.GetEnvironmentVariable("VISUALTEX_INLINE_BASELINE_MATHTYPE_SEED");
        AssertTrue(!string.IsNullOrWhiteSpace(seedPath) && File.Exists(seedPath),
            "Set VISUALTEX_INLINE_BASELINE_MATHTYPE_SEED to a saved MathType fixture. "
            + "Use cp146-large-body-mathTypeOle-200/result.docx for saved simple/script/tall cases.");
        Directory.CreateDirectory(artifactRoot);
        var seedHash = InlineBaselineFileHash(seedPath!);
        var ownedSeedPath = Path.Combine(artifactRoot, "owned-seed-copy.docx");
        File.Copy(seedPath!, ownedSeedPath, overwrite: false);
        var seeds = ReadInlineBaselineSeeds(ownedSeedPath);
        File.WriteAllText(Path.Combine(artifactRoot, "seed-manifest.json"),
            JsonSerializer.Serialize(new { Source = Path.GetFullPath(seedPath!), Sha256 = seedHash, Seeds = seeds },
                new JsonSerializerOptions { WriteIndented = true }));
        var modes = new[] {
            Word.WdBaselineAlignment.wdBaselineAlignAuto,
            Word.WdBaselineAlignment.wdBaselineAlignBaseline,
            Word.WdBaselineAlignment.wdBaselineAlignCenter,
            Word.WdBaselineAlignment.wdBaselineAlignTop,
            Word.WdBaselineAlignment.wdBaselineAlignFarEast50
        };
        var measurements = new List<object>();
        foreach (var mode in modes)
        {
            var name = mode.ToString().Replace("wdBaselineAlign", "");
            var caseRoot = Path.Combine(artifactRoot, name);
            Directory.CreateDirectory(caseRoot);
            var sourcePath = Path.Combine(caseRoot, "source-MathType-inline.docx");
            var targetPath = Path.Combine(caseRoot, "target-VisualTeX-inline.docx");
            CreateInlineBaselineFixture(sourcePath, seeds, mode);
            List<InlineBaselineSample> source;
            using (var host = new WordPerformanceHost(sourcePath))
            {
                var service = new WordFormulaService(host.Application);
                var plan = service.CaptureFormulaFormatConversionPlan(true,
                    FormulaOleContract.MathTypeOleMode, FormulaOleContract.NativeOleMode);
                AssertInlineBaselineSemantics(plan, seeds, "source " + name);
                source = CaptureInlineBaselineSamples(host, plan, seeds, mode, caseRoot, "source");
                // Save the owned copy under its target name before mutation. The
                // source fixture and the original supplied seed remain untouched.
                host.Document.SaveAs2(targetPath, Word.WdSaveFormat.wdFormatXMLDocument,
                    AddToRecentFiles: false);
                host.AddIn.OnConvertMathTypeToVisualTeXDocument(new object());
                WaitForAddInIdleFast(host.AddIn, TimeSpan.FromMinutes(10));
                AssertEqual(0, service.CountFormulaObjectsForLatex(true, FormulaOleContract.MathTypeOleMode),
                    name + " source objects remain after conversion");
                var after = service.CaptureFormulaFormatConversionPlan(true,
                    FormulaOleContract.NativeOleMode, FormulaOleContract.MathTypeOleMode);
                AssertInlineBaselineSemantics(after, seeds, "converted " + name);
                host.Document.Save();
            }
            // Read actual persisted Word rendering, after the conversion process
            // has released its OLE servers, instead of trusting a transient preview.
            using (var host = new WordPerformanceHost(targetPath))
            {
                var plan = new WordFormulaService(host.Application).CaptureFormulaFormatConversionPlan(
                    true, FormulaOleContract.NativeOleMode, FormulaOleContract.MathTypeOleMode);
                AssertInlineBaselineSemantics(plan, seeds, "reopened " + name);
                var target = CaptureInlineBaselineSamples(host, plan, seeds, mode, caseRoot, "target-reopened");
                for (var i = 0; i < seeds.Count; i++)
                {
                    var before = source[i]; var after = target[i];
                    var axisDelta = after.EqualsAxisFromProseBaselinePt - before.EqualsAxisFromProseBaselinePt;
                    var topDelta = after.InkTopFromProseBaselinePt - before.InkTopFromProseBaselinePt;
                    var bottomDelta = after.InkBottomFromProseBaselinePt - before.InkBottomFromProseBaselinePt;
                    var note = axisDelta.HasValue
                        ? (Math.Abs(axisDelta.Value) > 1.25
                            ? "Measured equation-axis displacement exceeds 1.25 pt; inspect both screenshots and host alignment."
                            : "Equation-axis displacement is within 1.25 pt; glyph/script fidelity still requires screenshot review.")
                        : "No unique equals-axis pair; compare the real screenshots. Ink bounds alone do not establish baseline correctness.";
                    measurements.Add(new {
                        Alignment = name, Seed = seeds[i], Source = before, Target = after,
                        PositionDeltaPt = after.ObjectPositionPt - before.ObjectPositionPt,
                        HeightDeltaPt = after.HeightPt - before.HeightPt,
                        InkTopDeltaPt = topDelta, InkBottomDeltaPt = bottomDelta,
                        EqualsAxisDeltaPt = axisDelta,
                        Direction = axisDelta.HasValue ? (axisDelta > 0 ? "down" : axisDelta < 0 ? "up" : "unchanged") : "undetermined",
                        Assessment = note,
                        // Source Position is not an oracle. Explicit Center/Top/
                        // FarEast50 use Word's host policy, not baseline descent.
                        VisualVerdict = "ReviewRequired"
                    });
                }
            }
            WriteInlineBaselineReport(artifactRoot, measurements, seedHash, complete: false);
            Console.WriteLine("[INLINE BASELINE CAPTURED] " + name + " formulas=" + seeds.Count
                + "; source/target-reopened screenshots and measured offsets written; visual review pending.");
        }
        AssertEqual(seedHash, InlineBaselineFileHash(seedPath!), "Original MathType seed changed.");
        WriteInlineBaselineReport(artifactRoot, measurements, seedHash, complete: true);
        Console.WriteLine("[INLINE BASELINE EVIDENCE READY] Structural/policy checks passed; "
            + measurements.Count + " source/target pairs captured. This is not a visual-pass claim.");
    }

    private static string InlineBaselineFileHash(string path)
    {
        using var stream = File.OpenRead(path);
        using var hash = SHA256.Create();
        return BitConverter.ToString(hash.ComputeHash(stream)).Replace("-", "");
    }

    private static void WriteInlineBaselineReport(string root, List<object> cases, string hash, bool complete)
    {
        File.WriteAllText(Path.Combine(root, "inline-baseline-report.json"), JsonSerializer.Serialize(new {
            CaptureComplete = complete, StructuralChecks = complete ? "Passed" : "InProgress",
            VisualVerdict = "ReviewRequired", SeedSha256 = hash,
            Interpretation = "Positive deltas mean downward relative to identical HH prose. "
                + "Do not treat source Font.Position as correct by definition. "
                + "Center and Top deliberately use zero offset; FarEast50 uses half the excess height above the semantic em. "
                + "The equals detector is independent pixel evidence, not a formula semantic or typography oracle.",
            Cases = cases
        }, new JsonSerializerOptions { WriteIndented = true }));
    }

    private static List<InlineBaselineSeed> ReadInlineBaselineSeeds(string ownedPath)
    {
        using var host = new WordPerformanceHost(ownedPath);
        var plan = new WordFormulaService(host.Application).CaptureFormulaFormatConversionPlan(
            true, FormulaOleContract.MathTypeOleMode, FormulaOleContract.NativeOleMode);
        var requested = Environment.GetEnvironmentVariable("VISUALTEX_INLINE_BASELINE_SOURCE_ORDINALS");
        var ordinals = string.IsNullOrWhiteSpace(requested)
            ? (plan.Targets.Count >= 5 ? new[] { 1, 3, 5 } : Enumerable.Range(1, Math.Min(3, plan.Targets.Count)).ToArray())
            : requested!.Split(',').Select(value => int.Parse(value.Trim(), CultureInfo.InvariantCulture)).ToArray();
        AssertTrue(ordinals.Length >= 3 && ordinals.Length <= 5 && ordinals.Distinct().Count() == ordinals.Length,
            "Select three to five distinct saved source ordinals for focused simple/script/tall coverage.");
        var result = new List<InlineBaselineSeed>();
        foreach (var ordinal in ordinals)
        {
            AssertTrue(ordinal > 0 && ordinal <= plan.Targets.Count, "Source ordinal out of range: " + ordinal);
            var target = plan.Targets[ordinal - 1];
            AssertTrue(!string.IsNullOrWhiteSpace(target.SourceMathMl), "Saved MathType source has no readable MathML: " + ordinal);
            var xml = XDocument.Parse(target.SourceMathMl!);
            var names = new HashSet<string>(xml.Descendants().Select(node => node.Name.LocalName));
            var text = string.Concat(xml.Descendants().Where(node => node.Name.LocalName == "mo").Select(node => node.Value));
            var family = names.Contains("mtable") ? "matrix"
                : names.Contains("mfrac") ? "fraction"
                : names.Contains("msqrt") || names.Contains("mroot") ? "radical"
                : text.Contains("∫") ? "integral" : text.Contains("∑") ? "sum"
                : names.Contains("msup") || names.Contains("msub") || names.Contains("msubsup") ? "script" : "simple";
            var shape = InlineBaselineFindShape(host.Document, target);
            Word.Range? range = null; Word.Range? character = null;
            try
            {
                range = shape.Range;
                character = InlineBaselineObjectCharacter(host.Document, range);
                result.Add(new InlineBaselineSeed {
                    SourceOrdinal = ordinal, Family = family, Latex = target.Latex, MathMl = target.SourceMathMl!,
                    FontSizePt = target.FontSizePt, SeedObjectPositionPt = character.Font.Position,
                    SeedWidthPt = shape.Width, SeedHeightPt = shape.Height, FlatOpc = range.WordOpenXML
                });
            }
            finally { Release(character); Release(range); Release(shape); }
        }
        AssertTrue(result.Any(item => item.Family == "simple" || item.Family == "script")
            && result.Any(item => item.Family == "fraction" || item.Family == "matrix" || item.Family == "integral")
            && result.Any(item => XDocument.Parse(item.MathMl).Descendants().Any(node =>
                node.Name.LocalName is "msup" or "msub" or "msubsup")),
            "Seed selection lacks simple/script/tall coverage. Use saved source ordinals 1,3,5 from the rich fixture.");
        return result;
    }

    private static void CreateInlineBaselineFixture(string path, List<InlineBaselineSeed> seeds, Word.WdBaselineAlignment mode)
    {
        using var host = new WordPerformanceHost(null);
        var doc = host.Document;
        doc.Content.Text = "MathType to VisualTeX inline baseline evidence: " + mode + "\r";
        for (var i = 0; i < seeds.Count; i++)
        {
            var seed = seeds[i];
            var start = doc.Content.End - 1;
            var heading = "CASE " + (i + 1).ToString("00") + " " + seed.Family + "\r";
            var tail = doc.Range(start, start);
            try { tail.Text = heading + "HH=HH \uE000 HH=HH\r"; } finally { Release(tail); }
            var insert = doc.Range(start + heading.Length + 6, start + heading.Length + 7);
            try { insert.InsertXML(seed.FlatOpc); } finally { Release(insert); }
            var shape = doc.InlineShapes[doc.InlineShapes.Count];
            var range = shape.Range;
            var paragraph = range.Paragraphs[1].Range;
            var prefix = doc.Range(paragraph.Start, range.Start);
            var suffix = doc.Range(range.End, paragraph.End - 1);
            try
            {
                AssertEqual("HH=HH ", prefix.Text, "Fixture import consumed leading prose.");
                AssertEqual(" HH=HH", suffix.Text, "Fixture import changed its paragraph/trailing prose.");
                paragraph.ParagraphFormat.BaseLineAlignment = mode;
                paragraph.ParagraphFormat.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceSingle;
                paragraph.ParagraphFormat.SpaceBefore = 0;
                paragraph.ParagraphFormat.SpaceAfter = 16;
                paragraph.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
                paragraph.ParagraphFormat.KeepTogether = -1;
                foreach (var prose in new[] { prefix, suffix })
                {
                    prose.Font.Name = "Times New Roman";
                    prose.Font.Size = (float)FormulaFontSize.Normalize(seed.FontSizePt);
                    prose.Font.Position = 0; prose.Font.Bold = 0; prose.Font.Italic = 0;
                    prose.Font.Color = Word.WdColor.wdColorBlue;
                }
            }
            finally { Release(suffix); Release(prefix); Release(paragraph); Release(range); Release(shape); }
        }
        AssertEqual(seeds.Count, doc.InlineShapes.Count, "Fixture host count.");
        doc.SaveAs2(path, Word.WdSaveFormat.wdFormatXMLDocument, AddToRecentFiles: false);
    }

    private static void AssertInlineBaselineSemantics(WordFormulaFormatConversionPlan plan,
        List<InlineBaselineSeed> seeds, string label)
    {
        AssertEqual(seeds.Count, plan.Targets.Count, label + " formula count");
        for (var i = 0; i < seeds.Count; i++)
        {
            var target = plan.Targets[i];
            AssertEqual("inline", target.DisplayMode, label + " formula became display " + i);
            AssertTrue(!target.Numbered, label + " formula acquired numbering " + i);
            if (target.SourceMathMl is not null)
                AssertEqual(WordNativeOmmlSemanticComparer.ComputeMathMlSemanticSignature(seeds[i].MathMl),
                    WordNativeOmmlSemanticComparer.ComputeMathMlSemanticSignature(target.SourceMathMl),
                    label + " saved-source semantics " + i);
            else
                AssertEqual(seeds[i].Latex, target.Latex, label + " saved-source LaTeX " + i);
        }
    }

    private static Word.InlineShape InlineBaselineFindShape(Word.Document doc, WordFormulaFormatConversionTarget target)
    {
        var range = WordFormulaOperationLocator.ResolveCapturedRange(doc, target.SourceObjectId);
        var shapes = range.InlineShapes;
        try { AssertEqual(1, shapes.Count, "Captured formula does not own exactly one inline shape."); return shapes[1]; }
        finally { Release(shapes); Release(range); }
    }

    private static Word.Range InlineBaselineObjectCharacter(Word.Document doc, Word.Range host)
    {
        for (var position = host.Start; position < host.End; position++)
        {
            var probe = doc.Range(position, position + 1);
            if (probe.Text == "\u0001") return probe;
            Release(probe);
        }
        throw new InvalidDataException("No actual U+0001 object result character in the local OLE host.");
    }

    private static Rectangle InlineBaselineScreenRect(Word.Window window, Word.Range range)
    {
        window.GetPoint(out var left, out var top, out var width, out var height, range);
        AssertTrue(width > 0 && height > 0, "Word did not provide a visible range rectangle.");
        return new Rectangle(left, top, width, height);
    }

    private static List<InlineBaselineSample> CaptureInlineBaselineSamples(WordPerformanceHost host,
        WordFormulaFormatConversionPlan plan, List<InlineBaselineSeed> seeds,
        Word.WdBaselineAlignment mode, string root, string phase)
    {
        var samples = new List<InlineBaselineSample>();
        host.Application.Visible = true;
        host.Document.Activate();
        var window = host.Document.ActiveWindow;
        var handle = ConversionCaptureNative.GetAncestor(new IntPtr(window.Hwnd), 2);
        try
        {
            window.View.Type = Word.WdViewType.wdPrintView;
            window.View.ShowAll = false;
            window.View.ShowFieldCodes = false;
            window.View.Zoom.Percentage = 150;
            ConversionCaptureNative.ShowWindow(handle, 9);
            ConversionCaptureNative.SetForegroundWindow(handle);
            for (var i = 0; i < plan.Targets.Count; i++)
            {
                var target = plan.Targets[i];
                var shape = InlineBaselineFindShape(host.Document, target);
                var range = shape.Range;
                var character = InlineBaselineObjectCharacter(host.Document, range);
                var paragraph = range.Paragraphs[1].Range;
                var prefix = host.Document.Range(paragraph.Start, range.Start);
                var suffix = host.Document.Range(range.End, paragraph.End - 1);
                var instruction = host.Document.Range(range.Start, range.Start + 1);
                try
                {
                    AssertEqual("HH=HH ", prefix.Text, phase + " leading prose");
                    AssertEqual(" HH=HH", suffix.Text, phase + " trailing prose");
                    AssertEqual((int)mode, (int)paragraph.ParagraphFormat.BaseLineAlignment, phase + " paragraph alignment");
                    AssertEqual(0, prefix.Font.Position, phase + " leading prose baseline");
                    AssertEqual(0, suffix.Font.Position, phase + " trailing prose baseline");
                    window.ScrollIntoView(paragraph, true);
                    host.Application.Selection.SetRange(paragraph.End - 1, paragraph.End - 1);
                    System.Windows.Forms.Application.DoEvents();
                    Thread.Sleep(180);
                    AssertEqual(handle, ConversionCaptureNative.GetAncestor(InlineBaselineCaptureNative.GetForegroundWindow(), 2),
                        "Owned Word window is not foreground; refusing to capture another application's pixels.");
                    var objectRect = InlineBaselineScreenRect(window, character);
                    var prefixRect = InlineBaselineScreenRect(window, prefix);
                    var suffixRect = InlineBaselineScreenRect(window, suffix);
                    var crop = Rectangle.Union(Rectangle.Union(objectRect, prefixRect), suffixRect);
                    crop.Inflate(16, 14);
                    var screen = System.Windows.Forms.Screen.FromHandle(handle).Bounds;
                    AssertTrue(screen.Contains(crop), "Evidence row is clipped by the screen; capture is not trustworthy.");
                    using var bitmap = new Bitmap(crop.Width, crop.Height);
                    using (var graphics = Graphics.FromImage(bitmap))
                        graphics.CopyFromScreen(crop.Location, Point.Empty, bitmap.Size);
                    var shot = Path.Combine(root, phase + "-formula-" + (i + 1).ToString("00") + ".png");
                    bitmap.Save(shot, ImageFormat.Png);
                    Rectangle Local(Rectangle value) => new(value.X - crop.X, value.Y - crop.Y, value.Width, value.Height);
                    var formulaInk = InlineBaselineInk(bitmap, Local(objectRect), blue: false);
                    var prefixInk = InlineBaselineInk(bitmap, Local(prefixRect), blue: true);
                    var suffixInk = InlineBaselineInk(bitmap, Local(suffixRect), blue: true);
                    AssertTrue(!formulaInk.IsEmpty && !prefixInk.IsEmpty && !suffixInk.IsEmpty,
                        phase + " missing formula or blue HH prose pixels");
                    AssertTrue(Math.Abs(prefixInk.Bottom - suffixInk.Bottom) <= 2,
                        phase + " prose anchors are not on the same visible line");
                    var pixelsPerPoint = objectRect.Height / (double)shape.Height;
                    AssertTrue(pixelsPerPoint > 0.5 && pixelsPerPoint < 8, "Implausible Word object pixel scale.");
                    var baseline = (prefixInk.Bottom + suffixInk.Bottom) / 2.0;
                    var axes = InlineBaselineEqualsAxes(bitmap, Local(objectRect),
                        FormulaFontSize.Normalize(seeds[i].FontSizePt) * pixelsPerPoint);
                    var hasOneEquals = XDocument.Parse(seeds[i].MathMl).Descendants()
                        .Count(node => node.Name.LocalName == "mo" && node.Value == "=") == 1;
                    double? axis = hasOneEquals && axes.Count == 1 ? axes[0] : null;
                    var metadata = WordFormulaMetadataReader.TryRead(shape);
                    var metric = new InlineBaselineSample {
                        Ordinal = i + 1, Family = seeds[i].Family, Latex = target.Latex, Phase = phase,
                        Alignment = mode.ToString(), AlignmentValue = (int)mode,
                        ObjectCharacterCode = character.Text[0], ObjectPositionPt = character.Font.Position,
                        HostInstructionPositionPt = instruction.Font.Position,
                        PrefixPositionPt = prefix.Font.Position, SuffixPositionPt = suffix.Font.Position,
                        WidthPt = shape.Width, HeightPt = shape.Height,
                        SemanticFontSizePt = seeds[i].FontSizePt,
                        CapturedSourcePositionPt = target.SourceInlineWordPosition,
                        CapturedSourceBottomWhitespacePt = target.SourceInlineBottomWhitespacePoints,
                        ExportHeightPx = metadata?.RenderHeightPx, ExportBaselinePx = metadata?.Baseline,
                        ObjectScreen = objectRect, PrefixScreen = prefixRect, SuffixScreen = suffixRect,
                        CaptureScreen = crop, FormulaInk = formulaInk, PrefixInk = prefixInk, SuffixInk = suffixInk,
                        ScreenPixelsPerPoint = pixelsPerPoint, ProseBaselineYPx = baseline,
                        FormulaEqualsAxisYPx = axis, EqualsAxisCandidateCount = axes.Count,
                        InkTopFromProseBaselinePt = (formulaInk.Top - baseline) / pixelsPerPoint,
                        InkBottomFromProseBaselinePt = (formulaInk.Bottom - baseline) / pixelsPerPoint,
                        EqualsAxisFromProseBaselinePt = axis.HasValue ? (axis.Value - baseline) / pixelsPerPoint : null,
                        Screenshot = shot
                    };
                    if (phase.StartsWith("target", StringComparison.Ordinal))
                    {
                        AssertTrue(metadata is not null, "Converted object lost VisualTeX metadata.");
                        metric.ExpectedPolicyPositionPt = WordInlineAlignment.CalculateFontPositionForHost(
                            (WordInlineHostAlignment)(int)mode, shape.Height,
                            (float)(metadata!.RenderHeightPx ?? 0), (float?)metadata.Baseline,
                            existingFontPosition: null, sourceSemanticFontSizePoints: FormulaFontSize.ResolveSemanticFontSize(metadata),
                            targetSemanticFontSizePoints: FormulaFontSize.ResolveSemanticFontSize(metadata));
                        AssertEqual(metric.ExpectedPolicyPositionPt.Value, metric.ObjectPositionPt,
                            phase + " persisted object does not follow its host-alignment policy");
                    }
                    samples.Add(metric);
                    File.WriteAllText(Path.Combine(root, phase + "-measurements.json"),
                        JsonSerializer.Serialize(samples, new JsonSerializerOptions { WriteIndented = true }));
                }
                finally
                {
                    Release(instruction); Release(suffix); Release(prefix); Release(paragraph);
                    Release(character); Release(range); Release(shape);
                }
            }
        }
        finally { Release(window); host.Application.Visible = false; }
        return samples;
    }

    private static bool InlineBaselineIsInk(Color color, bool blue) => blue
        ? color.B > 140 && color.R < 140 && color.G < 140
        : color.R < 170 && color.G < 170 && color.B < 170;

    private static Rectangle InlineBaselineInk(Bitmap image, Rectangle region, bool blue)
    {
        region.Intersect(new Rectangle(Point.Empty, image.Size));
        var left = region.Right; var top = region.Bottom; var right = -1; var bottom = -1;
        for (var y = region.Top; y < region.Bottom; y++)
        for (var x = region.Left; x < region.Right; x++)
            if (InlineBaselineIsInk(image.GetPixel(x, y), blue))
            { left = Math.Min(left, x); top = Math.Min(top, y); right = Math.Max(right, x); bottom = Math.Max(bottom, y); }
        return right < left ? Rectangle.Empty : Rectangle.FromLTRB(left, top, right + 1, bottom + 1);
    }

    // An equals sign supplies a stable visible math axis even when the picture's
    // top/bottom change with superscripts or renderer padding. Only report a
    // unique two-component pair; ambiguity remains visible rather than guessed.
    private static List<double> InlineBaselineEqualsAxes(Bitmap image, Rectangle region, double emPixels)
    {
        region.Intersect(new Rectangle(Point.Empty, image.Size));
        var visited = new bool[region.Width, region.Height];
        var bars = new List<Rectangle>();
        var queue = new Queue<Point>();
        for (var yy = 0; yy < region.Height; yy++)
        for (var xx = 0; xx < region.Width; xx++)
        {
            if (visited[xx, yy] || !InlineBaselineIsInk(image.GetPixel(region.X + xx, region.Y + yy), blue: false)) continue;
            visited[xx, yy] = true; queue.Enqueue(new Point(xx, yy));
            var minX = xx; var maxX = xx; var minY = yy; var maxY = yy; var pixels = 0;
            while (queue.Count > 0)
            {
                var p = queue.Dequeue(); pixels++;
                minX = Math.Min(minX, p.X); maxX = Math.Max(maxX, p.X);
                minY = Math.Min(minY, p.Y); maxY = Math.Max(maxY, p.Y);
                for (var dy = -1; dy <= 1; dy++)
                for (var dx = -1; dx <= 1; dx++)
                {
                    var x = p.X + dx; var y = p.Y + dy;
                    if (x < 0 || y < 0 || x >= region.Width || y >= region.Height || visited[x, y]) continue;
                    if (!InlineBaselineIsInk(image.GetPixel(region.X + x, region.Y + y), blue: false)) continue;
                    visited[x, y] = true; queue.Enqueue(new Point(x, y));
                }
            }
            var r = Rectangle.FromLTRB(region.X + minX, region.Y + minY, region.X + maxX + 1, region.Y + maxY + 1);
            if (r.Width >= emPixels * 0.25 && r.Width <= emPixels * 1.2
                && r.Height <= Math.Max(3, emPixels * 0.22) && r.Width >= r.Height * 2
                && pixels >= r.Width * r.Height * 0.6) bars.Add(r);
        }
        var axes = new List<double>();
        for (var i = 0; i < bars.Count; i++)
        for (var j = i + 1; j < bars.Count; j++)
        {
            var a = bars[i]; var b = bars[j];
            var gap = b.Top - a.Bottom;
            if (Math.Abs(a.Left - b.Left) <= 2 && Math.Abs(a.Width - b.Width) <= 2
                && gap >= Math.Max(1, emPixels * 0.04) && gap <= emPixels * 0.45)
                axes.Add((a.Top + a.Bottom + b.Top + b.Bottom) / 4.0);
        }
        return axes;
    }
}
