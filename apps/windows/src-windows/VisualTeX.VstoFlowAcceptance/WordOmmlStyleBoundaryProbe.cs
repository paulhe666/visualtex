using System.Text.Json;
using System.Xml.Linq;
using VisualTeX.WordVsto;
using Word = Microsoft.Office.Interop.Word;

namespace VisualTeX.VstoFlowAcceptance;
internal static partial class Program
{
    private static void RunWordOmmlStyleBoundaryProbe(string artifactRoot)
    {
        if (Environment.GetEnvironmentVariable("VISUALTEX_OMML_NATIVE_OPERATOR_CONTROL") == "1")
        {
            using var host = new WordPerformanceHost(null);
            host.Document.Content.Text = "a∥b\r";
            Word.Range? range = null; Word.OMaths? maths = null; Word.Range? added = null; Word.OMath? equation = null;
            Word.Range? result = null; Word.Characters? characters = null;
            var native = new List<object>();
            try
            {
                range = host.Document.Range(0, 3); maths = host.Document.OMaths; added = maths.Add(range);
                equation = maths[1]; equation.BuildUp(); result = equation.Range; characters = result.Characters;
                for (var index = 1; index <= characters.Count; index++)
                {
                    Word.Range? character = null; Word.Font? font = null;
                    try { character = characters[index]; font = character.Font;
                        native.Add(new { index, text = character.Text, italic = font.Italic, bold = font.Bold, font = font.Name }); }
                    finally { Release(font); Release(character); }
                }
                var projection = CaptureOmmlStyleProbe(host.Document, artifactRoot, "native-default", "buildup");
                File.WriteAllText(Path.Combine(artifactRoot, "native-default-properties.json"), JsonSerializer.Serialize(new { native, projection }));
                host.Save(Path.Combine(artifactRoot, "native-default.docx"));
            }
            finally { Release(characters); Release(result); Release(equation); Release(added); Release(maths); Release(range); }
            return;
        }
        var existingRoot = Environment.GetEnvironmentVariable("VISUALTEX_OMML_STYLE_EXISTING");
        if (!string.IsNullOrWhiteSpace(existingRoot))
        {
            var properties = new List<object>();
            foreach (var name in new[] { "F02", "italic", "normal", "bold-italic" })
            {
                using var host = new WordPerformanceHost(Path.Combine(existingRoot, name + ".docx"));
                Word.OMaths? maths = null; Word.OMath? equation = null; Word.Range? range = null; Word.Characters? characters = null;
                try
                {
                    maths = host.Document.OMaths; equation = maths[1]; range = equation.Range; characters = range.Characters;
                    for (var i = 1; i <= characters.Count; i++)
                    {
                        Word.Range? character = null; Word.Font? font = null;
                        try { character = characters[i]; font = character.Font;
                            properties.Add(new { name, index = i, text = character.Text, italic = font.Italic, bold = font.Bold, font = font.Name }); }
                        finally { Release(font); Release(character); }
                    }
                }
                finally { Release(characters); Release(range); Release(equation); Release(maths); }
            }
            File.WriteAllText(Path.Combine(artifactRoot, "native-font-properties.json"), JsonSerializer.Serialize(properties));
            return;
        }
        const string ns = "http://www.w3.org/1998/Math/MathML";
        var evidencePath = Environment.GetEnvironmentVariable("VISUALTEX_OMML_STYLE_SOURCE")!;
        using var evidence = JsonDocument.Parse(File.ReadAllText(evidencePath));
        var runItalicProbe = Environment.GetEnvironmentVariable("VISUALTEX_OMML_RUN_ITALIC_PROBE") == "1";
        var cases = Environment.GetEnvironmentVariable("VISUALTEX_OMML_MONOSPACE_PROBE") == "1"
            ? new[] {
                (Name: "monospace-latin", MathMl: $"<math xmlns='{ns}'><mi mathvariant='monospace'>x</mi></math>"),
                (Name: "monospace-digits", MathMl: $"<math xmlns='{ns}'><mn mathvariant='monospace'>09</mn></math>"),
                (Name: "monospace-mixed", MathMl: $"<math xmlns='{ns}'><mi mathvariant='monospace'>x</mi><mo>+</mo><mi mathvariant='normal'>y</mi></math>"),
                (Name: "monospace-script", MathMl: $"<math xmlns='{ns}'><msup><mi mathvariant='monospace'>x</mi><mn mathvariant='monospace'>2</mn></msup></math>")
            }
            : Environment.GetEnvironmentVariable("VISUALTEX_OMML_SCRIPT_ROLE_PROBE") == "1"
            ? new[] {
                (Name: "script-role-operator-paren", MathMl: $"<math xmlns='{ns}'><msup><mo mathvariant='normal'>)</mo><mn>2</mn></msup></math>"),
                (Name: "script-role-operator-bars", MathMl: $"<math xmlns='{ns}'><msub><mo mathvariant='normal'>‖</mo><mi>i</mi></msub></math>"),
                (Name: "script-role-identifier", MathMl: $"<math xmlns='{ns}'><msub><mi mathvariant='normal'>abc</mi><mi>i</mi></msub></math>"),
                (Name: "script-role-prose-symbol", MathMl: $"<math xmlns='{ns}'><msub><mtext>∥</mtext><mi>i</mi></msub></math>")
            }
            : runItalicProbe
            ? new[] {
                (Name: "word-italic-with-math-style", MathMl: $"<math xmlns='{ns}'><mi>a</mi><mo mathvariant='italic'>∥</mo><mi>b</mi></math>"),
                (Name: "word-italic-without-math-style", MathMl: $"<math xmlns='{ns}'><mi>a</mi><mo mathvariant='italic'>∥</mo><mi>b</mi></math>")
            }
            : Environment.GetEnvironmentVariable("VISUALTEX_OMML_FENCE_PROBE") == "1"
            ? new[] {
                (Name: "original-F09", MathMl: $"<math xmlns='{ns}'><mrow data-mjx-texclass='INNER'><mo data-mjx-texclass='OPEN'>⟮</mo><mi>x</mi><mo data-mjx-texclass='CLOSE'>⟯</mo></mrow><mo>+</mo><mrow data-mjx-texclass='INNER'><mo data-mjx-texclass='OPEN' fence='true' stretchy='true' symmetric='true'>⏐</mo><mi>y</mi><mo data-mjx-texclass='CLOSE' fence='true' stretchy='true' symmetric='true'></mo></mrow></math>"),
                (Name: "group-fence", MathMl: $"<math xmlns='{ns}'><mfenced open='⟮' close='⟯' separators=''><mi>x</mi></mfenced></math>"),
                (Name: "arrowvert-fence", MathMl: $"<math xmlns='{ns}'><mfenced open='⏐' close='' separators=''><mi>y</mi></mfenced></math>"),
                (Name: "round-fence-control", MathMl: $"<math xmlns='{ns}'><mfenced open='(' close=')' separators=''><mi>x</mi></mfenced></math>")
            }
            : Environment.GetEnvironmentVariable("VISUALTEX_OMML_SPACE_PROBE") == "1"
            ? new[] {
                (Name: "one-em", MathMl: $"<math xmlns='{ns}'><mi>a</mi><mspace width='1.000em'/><mi>b</mi></math>"),
                (Name: "two-em", MathMl: $"<math xmlns='{ns}'><mi>a</mi><mspace width='2em'/><mi>b</mi></math>"),
                (Name: "zero-em", MathMl: $"<math xmlns='{ns}'><mi>a</mi><mspace width='0em'/><mi>b</mi></math>")
            }
            : new[] {
            (Name: "F02", MathMl: evidence.RootElement.GetProperty("expected").GetString()!),
            (Name: "italic", MathMl: $"<math xmlns='{ns}'><mo mathvariant='italic'>∥</mo></math>"),
            (Name: "normal", MathMl: $"<math xmlns='{ns}'><mo mathvariant='normal'>∥</mo></math>"),
            (Name: "bold-italic", MathMl: $"<math xmlns='{ns}'><mo mathvariant='bold-italic'>∥</mo></math>")
        };
        if (evidence.RootElement.TryGetProperty("cases", out var suppliedCases))
            cases = suppliedCases.EnumerateArray().Select(item => (
                Name: item.GetProperty("name").GetString()!,
                MathMl: item.GetProperty("mathMl").GetString()!)).ToArray();
        var rows = new List<object>();
        foreach (var item in cases)
        {
            var omml = WordOmmlConverter.TransformMathMlToOmml(item.MathMl);
            if (evidence.RootElement.TryGetProperty("nativeOmmlControls", out var controls)
                && controls.TryGetProperty(item.Name, out var nativeControl))
                omml = nativeControl.GetString()!;
            if (runItalicProbe)
            {
                XNamespace m = "http://schemas.openxmlformats.org/officeDocument/2006/math";
                XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
                var projection = XDocument.Parse(omml);
                foreach (var run in projection.Descendants(m + "r").Where(run => run.Element(m + "t")?.Value.Contains("∥") == true))
                {
                    if (item.Name == "word-italic-without-math-style")
                        run.Element(m + "rPr")?.Elements(m + "sty").Remove();
                    var wordProperties = new XElement(w + "rPr", new XElement(w + "i", new XAttribute(w + "val", "true")));
                    var mathProperties = run.Element(m + "rPr");
                    if (mathProperties is null) run.AddFirst(wordProperties);
                    else mathProperties.AddAfterSelf(wordProperties);
                }
                omml = projection.ToString(SaveOptions.DisableFormatting);
            }
            File.WriteAllText(Path.Combine(artifactRoot, item.Name + "-input.xml"), omml);
            var path = Path.Combine(artifactRoot, item.Name + ".docx");
            using (var host = new WordPerformanceHost(null))
            {
                host.Document.Content.Text = "\r";
                Word.Range? insertion = null;
                Word.OMaths? maths = null;
                Word.OMath? equation = null;
                Word.Range? equationRange = null;
                try
                {
                    insertion = host.Document.Content;
                    var package = XDocument.Parse(insertion.WordOpenXML);
                    XNamespace word = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
                    var inputDocument = XDocument.Parse(WordOmmlConverter.BuildDocumentXml(omml));
                    package.Descendants(word + "body").Single().ReplaceNodes(
                        inputDocument.Descendants(word + "body").Single().Elements().Select(element => new XElement(element)));
                    var inputPackage = package.ToString(SaveOptions.DisableFormatting);
                    File.WriteAllText(Path.Combine(artifactRoot, item.Name + "-insert-input.xml"), inputPackage);
                    insertion.InsertXML(inputPackage);
                    rows.Add(CaptureOmmlStyleProbe(host.Document, artifactRoot, item.Name, "insert"));
                    maths = host.Document.OMaths;
                    AssertEqual(1, maths.Count, item.Name + " direct insertion count");
                    equation = maths[1];
                    var skipBuildUp = evidence.RootElement.TryGetProperty("skipBuildUp", out var skip) && skip.GetBoolean();
                    if (!skipBuildUp) equation.BuildUp();
                    rows.Add(CaptureOmmlStyleProbe(host.Document, artifactRoot, item.Name, skipBuildUp ? "preserved-structure" : "buildup"));
                    equationRange = equation.Range;
                    WordFormulaHostLayout.ApplyOmmlLocalTypography(equationRange, 12);
                    rows.Add(CaptureOmmlStyleProbe(host.Document, artifactRoot, item.Name, "typography"));
                    host.Save(path);
                }
                finally { Release(equationRange); Release(equation); Release(maths); Release(insertion); }
            }
            using (var reopened = new WordPerformanceHost(path))
                rows.Add(CaptureOmmlStyleProbe(reopened.Document, artifactRoot, item.Name, "reopen"));
        }
        File.WriteAllText(Path.Combine(artifactRoot, "stages.json"), JsonSerializer.Serialize(rows,
            new JsonSerializerOptions { WriteIndented = true }));
        Console.WriteLine("[OMML BOUNDARY PROBE] " + cases.Length + " owned documents captured; no screenshots.");
    }

    private static object CaptureOmmlStyleProbe(Word.Document document, string root, string name, string stage)
    {
        Word.Range? content = null;
        try
        {
            content = document.Content;
            var xml = content.WordOpenXML;
            File.WriteAllText(Path.Combine(root, name + "-" + stage + ".xml"), xml);
            XNamespace math = "http://schemas.openxmlformats.org/officeDocument/2006/math";
            var parsed = XDocument.Parse(xml);
            var formula = parsed.Descendants(math + "oMath").Single();
            var omml = formula.ToString(SaveOptions.DisableFormatting);
            var mathMl = WordOmmlConverter.TransformOmmlToMathMl(omml, display: false);
            File.WriteAllText(Path.Combine(root, name + "-" + stage + ".mathml"), mathMl);
            if (name.StartsWith("script-role-", StringComparison.Ordinal))
            {
                var script = XDocument.Parse(mathMl).Descendants().Single(node => node.Name.LocalName is "msub" or "msup");
                var baseToken = script.Elements().First();
                while (baseToken.Name.LocalName == "mrow" && !baseToken.HasAttributes &&
                       baseToken.Nodes().Count() == 1 && baseToken.Elements().Count() == 1)
                    baseToken = baseToken.Elements().Single();
                var expected = name switch
                {
                    "script-role-operator-paren" => (Kind: "mo", Text: ")"),
                    "script-role-operator-bars" => (Kind: "mo", Text: "‖"),
                    "script-role-identifier" => (Kind: "mi", Text: "abc"),
                    _ => (Kind: "mtext", Text: "∥"),
                };
                AssertEqual(expected.Kind, baseToken.Name.LocalName, name + " " + stage + " script base role");
                AssertEqual(expected.Text, baseToken.Value, name + " " + stage + " script base text");
                AssertEqual("normal", MathMlTokenStyleContract.ResolveEffectiveVariant(baseToken), name + " " + stage + " script base style");
            }
            if (name.StartsWith("monospace-", StringComparison.Ordinal))
            {
                XNamespace presentation = "http://www.w3.org/1998/Math/MathML";
                var projected = XDocument.Parse(mathMl);
                var tokens = projected.Descendants().Where(node => node.Name.Namespace == presentation &&
                    node.Name.LocalName is "mi" or "mn" or "mo" or "mtext").ToArray();
                if (name == "monospace-digits")
                {
                    AssertEqual("09", string.Concat(tokens.Select(node => node.Value)), name + " " + stage + " text");
                    foreach (var token in tokens)
                    {
                        AssertEqual("mn", token.Name.LocalName, name + " " + stage + " role");
                        AssertEqual("monospace", MathMlTokenStyleContract.ResolveEffectiveVariant(token), name + " " + stage + " style");
                    }
                }
                else
                {
                    var expected = name == "monospace-mixed"
                        ? new[] { (Role: "mi", Text: "x", Style: "monospace"), (Role: "mo", Text: "+", Style: "normal"), (Role: "mi", Text: "y", Style: "normal") }
                        : name == "monospace-script"
                            ? new[] { (Role: "mi", Text: "x", Style: "monospace"), (Role: "mn", Text: "2", Style: "monospace") }
                            : new[] { (Role: "mi", Text: "x", Style: "monospace") };
                    AssertEqual(expected.Length, tokens.Length, name + " " + stage + " token count");
                    for (var index = 0; index < expected.Length; index++)
                    {
                        AssertEqual(expected[index].Role, tokens[index].Name.LocalName, name + " " + stage + " role");
                        AssertEqual(expected[index].Text, tokens[index].Value, name + " " + stage + " text");
                        AssertEqual(expected[index].Style, MathMlTokenStyleContract.ResolveEffectiveVariant(tokens[index]), name + " " + stage + " style");
                    }
                    if (name == "monospace-script")
                        AssertEqual(1, projected.Descendants(presentation + "msup").Count(), name + " " + stage + " script count");
                }
                AssertEqual(0, formula.Descendants(math + "nor").Count(), name + " " + stage + " no prose workaround");
            }
            return new { name, stage, omml, mathMl };
        }
        finally { Release(content); }
    }
}
