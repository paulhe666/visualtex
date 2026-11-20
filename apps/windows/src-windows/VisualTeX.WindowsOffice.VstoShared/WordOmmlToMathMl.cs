using System.Globalization;
using System.Text;
using System.Xml.Linq;
using static VisualTeX.WordVsto.WordOmmlConverter;
using static VisualTeX.WordVsto.WordOmmlFromMathMl;

namespace VisualTeX.WordVsto;

/// <summary>
/// The MathML Word's OMML-to-MathML transform returns, with what the transform loses restored: accents, matrix
/// alignment, fractions, limits, scripts, function names and math text spelling.
/// </summary>
internal static class WordOmmlToMathMl
{
    internal static void RestoreOmmlAccentSemantics(string omml, XElement mathMlRoot)
    {
        XNamespace officeMath = MathNamespace;
        XNamespace mathMl = "http://www.w3.org/1998/Math/MathML";
        var source = XDocument.Parse(omml, LoadOptions.PreserveWhitespace);
        var sourceAccents = source.Descendants(officeMath + "acc").ToArray();
        if (sourceAccents.Length == 0) return;

        // Office's bundled OMML->MathML stylesheet still contains legacy symbol
        // mappings for m:acc. In current Word documents that can turn a perfectly
        // valid Unicode combining accent (for example U+0305/U+20D7/U+0307) into
        // '-', '?' or mojibake. The OMML itself is the authoritative source, so
        // restore the accent mark after the XSL transform before MathType/LaTeX
        // consumers see it.
        var targetAccents = mathMlRoot
            .DescendantsAndSelf(mathMl + "mover")
            .Where(element => string.Equals(
                (string?)element.Attribute("accent"),
                "true",
                StringComparison.OrdinalIgnoreCase))
            .ToArray();
        var count = Math.Min(sourceAccents.Length, targetAccents.Length);
        for (var index = 0; index < count; index++)
        {
            var sourceAccent = sourceAccents[index];
            var officeCharacter = sourceAccent
                .Element(officeMath + "accPr")?
                .Element(officeMath + "chr")?
                .Attribute(officeMath + "val")?
                .Value;
            var canonicalCharacter = CanonicalMathMlAccentCharacter(officeCharacter);
            var targetAccent = targetAccents[index];
            var children = targetAccent.Elements().ToArray();
            if (children.Length < 2) continue;
            var restoredMark = new XElement(mathMl + "mo",
                new XAttribute("accent", "true"), canonicalCharacter);
            XNamespace word = WordNamespace;
            var controlRun = sourceAccent.Element(officeMath + "accPr")?
                .Element(officeMath + "ctrlPr")?.Element(word + "rPr");
            var bold = controlRun?.Element(word + "b");
            var italic = controlRun?.Element(word + "i");
            if (bold is not null || italic is not null)
            {
                bool Enabled(XElement? property) => property is not null
                    && (string?)property.Attribute(word + "val") is not ("0" or "false" or "off");
                var variant = Enabled(bold)
                    ? (Enabled(italic) ? "bold-italic" : "bold")
                    : (Enabled(italic) ? "italic" : "normal");
                restoredMark.SetAttributeValue("mathvariant", variant);
            }
            var color = MathMlColorContract.Normalize("#" + (string?)controlRun?.Element(word + "color")?.Attribute(word + "val"));
            if (color is not null) restoredMark.SetAttributeValue("mathcolor", color);
            children[1].ReplaceWith(restoredMark);
        }
    }

    // The stylesheet keeps only the first column's justification of an m:m.
    // Record every column (mcJc, repeated by m:count) so the LaTeX reader can
    // tell `aligned` (right, left, right, …) from a matrix. m:m and m:eqArr
    // are the only sources of mtable, so the two pre-order lists line up.
    internal static void RestoreOmmlMatrixColumnAlignment(string omml, XElement mathMlRoot)
    {
        XNamespace officeMath = MathNamespace;
        XNamespace mathMl = "http://www.w3.org/1998/Math/MathML";
        var source = XDocument.Parse(omml, LoadOptions.PreserveWhitespace);
        var sourceTables = source.Descendants()
            .Where(element => element.Name == officeMath + "m" || element.Name == officeMath + "eqArr")
            .ToArray();
        var targetTables = mathMlRoot.DescendantsAndSelf(mathMl + "mtable").ToArray();
        if (sourceTables.Length == 0 || sourceTables.Length != targetTables.Length)
            return;

        for (var index = 0; index < sourceTables.Length; index++)
        {
            if (sourceTables[index].Name != officeMath + "m") continue;
            var columns = (sourceTables[index].Element(officeMath + "mPr")?
                    .Elements(officeMath + "mcs") ?? Enumerable.Empty<XElement>())
                .SelectMany(group => group.Elements(officeMath + "mc"))
                .SelectMany(column =>
                {
                    var properties = column.Element(officeMath + "mcPr");
                    var count = int.TryParse(
                        (string?)properties?.Element(officeMath + "count")?.Attribute(officeMath + "val"),
                        out var parsed) && parsed > 0 ? Math.Min(parsed, 64) : 1;
                    var justification = ((string?)properties?.Element(officeMath + "mcJc")?
                        .Attribute(officeMath + "val") ?? "center").Trim().ToLowerInvariant();
                    return Enumerable.Repeat(justification, count);
                })
                .ToArray();
            if (columns.Length > 1)
                targetTables[index].SetAttributeValue("data-vt-columnalign", string.Join(" ", columns));
        }
    }

    internal static void RestoreOmmlNoBarFractionSemantics(string omml, XElement mathMlRoot)
    {
        XNamespace officeMath = MathNamespace;
        XNamespace mathMl = "http://www.w3.org/1998/Math/MathML";
        var source = XDocument.Parse(omml, LoadOptions.PreserveWhitespace);
        var sourceFractions = source.Descendants(officeMath + "f").ToArray();
        var targetFractions = mathMlRoot.DescendantsAndSelf(mathMl + "mfrac").ToArray();
        if (sourceFractions.Length == 0 || sourceFractions.Length != targetFractions.Length)
            return;

        for (var index = 0; index < sourceFractions.Length; index++)
        {
            var type = sourceFractions[index]
                .Element(officeMath + "fPr")?
                .Element(officeMath + "type")?
                .Attribute(officeMath + "val")?
                .Value;
            if (!string.Equals(type, "noBar", StringComparison.OrdinalIgnoreCase)) continue;
            targetFractions[index].SetAttributeValue("linethickness", "0");
        }
    }

    internal static void RestoreMergedNumericPunctuationTokens(XElement mathMlRoot)
    {
        XNamespace mathMl = "http://www.w3.org/1998/Math/MathML";
        foreach (var number in mathMlRoot.DescendantsAndSelf(mathMl + "mn").ToArray())
        {
            var value = number.Value;
            if (string.IsNullOrWhiteSpace(value)
                || (value.IndexOf(',') < 0 && value.IndexOf(';') < 0))
                continue;
            if (!value.All(character => char.IsDigit(character)
                    || character == ','
                    || character == ';'))
                continue;

            var tokens = new List<XElement>();
            var start = 0;
            for (var index = 0; index <= value.Length; index++)
            {
                var atEnd = index == value.Length;
                var punctuation = !atEnd && (value[index] == ',' || value[index] == ';');
                if (!atEnd && !punctuation) continue;
                if (index > start)
                    tokens.Add(new XElement(mathMl + "mn", value.Substring(start, index - start)));
                if (punctuation)
                    tokens.Add(new XElement(mathMl + "mo", value[index].ToString()));
                start = index + 1;
            }
            if (tokens.Count <= 1) continue;

            if (number.Parent?.Name == mathMl + "mrow")
                number.ReplaceWith(tokens.Cast<object>().ToArray());
            else
                number.ReplaceWith(new XElement(mathMl + "mrow", tokens));
        }
    }

    internal static void RestoreMathSymbolTextTokens(XElement mathMlRoot)
    {
        XNamespace mathMl = "http://www.w3.org/1998/Math/MathML";
        foreach (var text in mathMlRoot.DescendantsAndSelf(mathMl + "mtext").ToArray())
        {
            var value = text.Value;
            if (string.IsNullOrEmpty(value) || !ContainsOnlyMathSymbols(value)) continue;
            text.Name = mathMl + "mo";
        }
    }

    internal static void RestoreWordCanonicalAsteriskOperators(
        XElement mathMlRoot)
    {
        XNamespace mathMl = "http://www.w3.org/1998/Math/MathML";
        foreach (var token in mathMlRoot
                     .DescendantsAndSelf()
                     .Where(element =>
                         element.Name == mathMl + "mi"
                         || element.Name == mathMl + "mo")
                     .Where(element =>
                         !element.HasElements
                         && string.Equals(element.Value, "*", StringComparison.Ordinal))
                     .ToList())
        {
            // Word canonicalizes U+2217 ASTERISK OPERATOR to ASCII '*' when a
            // professional equation is materialized. Its reverse stylesheet then
            // misclassifies the punctuation character as <mi>, losing both the
            // operator token type and the original \ast semantics. In mathematical
            // token positions, canonicalize that Word spelling back to the Unicode
            // operator. Deliberate text-mode '*' remains mtext and is untouched.
            token.Name = mathMl + "mo";
            token.Value = RestoreWordAsteriskSpelling(token.Value);
        }
    }

    private static string RestoreWordAsteriskSpelling(string mathematicalText)
        => mathematicalText.Replace('*', '∗');

    internal static void NormalizeLegacyMathAlphabetSpelling(
        XDocument document,
        XNamespace math)
    {
        // Word may materialize legacy Unicode mathematical letter symbols as a
        // base Latin letter plus m:scr (for example ℜ -> R + fraktur), while the
        // converter can emit the equivalent precomposed Unicode symbol. Canonicalize
        // only the standard one-to-one legacy spellings so a real script/style
        // change still produces a different content signature.
        var spellings = new Dictionary<char, (char Base, string Script)>
        {
            ['ℭ'] = ('C', "fraktur"),
            ['ℌ'] = ('H', "fraktur"),
            ['ℑ'] = ('I', "fraktur"),
            ['ℜ'] = ('R', "fraktur"),
            ['ℨ'] = ('Z', "fraktur"),
            ['ℬ'] = ('B', "script"),
            ['ℰ'] = ('E', "script"),
            ['ℱ'] = ('F', "script"),
            ['ℋ'] = ('H', "script"),
            ['ℐ'] = ('I', "script"),
            ['ℒ'] = ('L', "script"),
            ['ℳ'] = ('M', "script"),
            ['ℛ'] = ('R', "script"),
            ['ℯ'] = ('e', "script"),
            ['ℊ'] = ('g', "script"),
            ['ℴ'] = ('o', "script"),
            ['ℂ'] = ('C', "double-struck"),
            ['ℍ'] = ('H', "double-struck"),
            ['ℕ'] = ('N', "double-struck"),
            ['ℙ'] = ('P', "double-struck"),
            ['ℚ'] = ('Q', "double-struck"),
            ['ℝ'] = ('R', "double-struck"),
            ['ℤ'] = ('Z', "double-struck"),
        };

        foreach (var run in document.Descendants(math + "r").ToArray())
        {
            var text = run.Element(math + "t");
            if (text is null || text.Value.Length != 1
                || !spellings.TryGetValue(text.Value[0], out var spelling))
                continue;

            var properties = run.Element(math + "rPr");
            var script = properties?.Element(math + "scr");
            var currentScript = (string?)script?.Attribute(math + "val");
            if (script is not null
                && !string.Equals(currentScript, spelling.Script, StringComparison.Ordinal))
                continue;

            if (properties is null)
            {
                properties = new XElement(math + "rPr");
                run.AddFirst(properties);
            }
            if (script is null)
            {
                script = new XElement(
                    math + "scr",
                    new XAttribute(math + "val", spelling.Script));
                properties.AddFirst(script);
            }
            else if (script.Attribute(math + "val") is null)
            {
                script.SetAttributeValue(math + "val", spelling.Script);
            }
            text.Value = spelling.Base.ToString();
        }
    }

    internal static void NormalizeMathScriptRunBoundaries(
        XDocument document,
        XNamespace math)
    {
        foreach (var run in document.Descendants(math + "r").ToArray())
        {
            var properties = run.Element(math + "rPr");
            if (properties?.Element(math + "scr") is null) continue;
            var texts = run.Elements(math + "t").ToArray();
            if (texts.Length != 1
                || run.Elements().Any(element => element.Name != math + "rPr" && element.Name != math + "t"))
                continue;

            var value = texts[0].Value;
            var segments = new List<(string Text, bool ScriptSensitive)>();
            var builder = new StringBuilder();
            bool? currentSensitive = null;
            for (var index = 0; index < value.Length;)
            {
                var category = CharUnicodeInfo.GetUnicodeCategory(value, index);
                var scalarLength = char.IsHighSurrogate(value[index])
                    && index + 1 < value.Length
                    && char.IsLowSurrogate(value[index + 1])
                    ? 2
                    : 1;
                var sensitive = category is UnicodeCategory.UppercaseLetter
                    or UnicodeCategory.LowercaseLetter
                    or UnicodeCategory.TitlecaseLetter
                    or UnicodeCategory.ModifierLetter
                    or UnicodeCategory.OtherLetter
                    or UnicodeCategory.DecimalDigitNumber
                    or UnicodeCategory.LetterNumber;
                if (currentSensitive.HasValue && currentSensitive.Value != sensitive)
                {
                    segments.Add((builder.ToString(), currentSensitive.Value));
                    builder.Clear();
                }
                currentSensitive = sensitive;
                builder.Append(value, index, scalarLength);
                index += scalarLength;
            }
            if (builder.Length > 0 && currentSensitive.HasValue)
                segments.Add((builder.ToString(), currentSensitive.Value));
            if (segments.Count == 1)
            {
                if (!segments[0].ScriptSensitive)
                {
                    // A math-alphabet script has no defined effect on a
                    // symbol-only run. Word removes it while materializing
                    // operators such as U+2201 COMPLEMENT, so compare the
                    // prepared and native trees through that same canonical
                    // spelling. Script/fraktur/double-struck on letters and
                    // digits remains semantic and is deliberately retained.
                    properties.Element(math + "scr")?.Remove();
                    if (!properties.HasElements
                        && !properties.HasAttributes
                        && string.IsNullOrWhiteSpace(properties.Value))
                        properties.Remove();
                }
                continue;
            }

            foreach (var segment in segments)
            {
                XElement? segmentProperties = new XElement(properties);
                if (!segment.ScriptSensitive)
                {
                    segmentProperties.Element(math + "scr")?.Remove();
                    if (!segmentProperties.HasElements
                        && !segmentProperties.HasAttributes
                        && string.IsNullOrWhiteSpace(segmentProperties.Value))
                        segmentProperties = null;
                }
                var segmentText = new XElement(math + "t", segment.Text);
                foreach (var attribute in texts[0].Attributes())
                    segmentText.SetAttributeValue(attribute.Name, attribute.Value);
                var replacement = new XElement(math + "r");
                if (segmentProperties is not null) replacement.Add(segmentProperties);
                replacement.Add(segmentText);
                run.AddBeforeSelf(replacement);
            }
            run.Remove();
        }
    }

    internal static string NormalizeImportedMathTextSpelling(string mathematicalText) =>
        RestoreWordAsteriskSpelling(
            mathematicalText
                .Replace('\u2212', '-')
                .Replace("\u2032", "'")
                .Replace("\u2033", "''")
                .Replace("\u2034", "'''")
                .Replace("\u2057", "''''"));

    private static bool ContainsOnlyMathSymbols(string value)
    {
        var sawSymbol = false;
        for (var index = 0; index < value.Length;)
        {
            var category = CharUnicodeInfo.GetUnicodeCategory(value, index);
            var scalarLength = char.IsHighSurrogate(value[index])
                && index + 1 < value.Length
                && char.IsLowSurrogate(value[index + 1])
                ? 2
                : 1;
            if (category != UnicodeCategory.MathSymbol) return false;
            sawSymbol = true;
            index += scalarLength;
        }
        return sawSymbol;
    }

    internal static void RestoreOmmlLimitBaseSemantics(string omml, XElement mathMlRoot)
    {
        XNamespace officeMath = MathNamespace;
        XNamespace mathMl = "http://www.w3.org/1998/Math/MathML";
        var source = XDocument.Parse(omml, LoadOptions.PreserveWhitespace);
        var limits = source
            .Descendants()
            .Where(element => element.Name == officeMath + "limLow"
                || element.Name == officeMath + "limUpp")
            .ToArray();
        if (limits.Length == 0) return;

        var usedTargets = new HashSet<XElement>();
        foreach (var limit in limits)
        {
            var baseSlot = limit.Element(officeMath + "e");
            if (baseSlot is null) continue;
            var visibleRuns = baseSlot
                .Descendants(officeMath + "r")
                .Where(run => run.Elements(officeMath + "t")
                    .Any(text => !string.IsNullOrWhiteSpace(text.Value)))
                .ToArray();
            if (visibleRuns.Length == 0) continue;

            var baseText = string.Concat(
                visibleRuns.SelectMany(run => run.Elements(officeMath + "t"))
                    .Select(text => text.Value));
            if (string.IsNullOrWhiteSpace(baseText)) continue;

            // Word's BuildUp/save materialization can discard m:nor from the base
            // run of m:limLow/m:limUpp even though the visible operator is still a
            // native limit operator. Preserve explicit m:nor whenever it survives;
            // when it does not, only infer upright grouping for TeX's standard
            // named operators. This restores lim/max/min/sup/inf/limsup/... without
            // ever joining an arbitrary italic variable sequence such as l*i*m.
            var explicitlyNormal = visibleRuns.All(run =>
                run.Element(officeMath + "rPr")?
                    .Element(officeMath + "nor") is not null);
            if (!explicitlyNormal && !IsStandardTexOperatorName(baseText))
                continue;

            var lower = limit.Name == officeMath + "limLow";
            var targetNames = lower
                ? new HashSet<XName> { mathMl + "msub", mathMl + "munder" }
                : new HashSet<XName> { mathMl + "msup", mathMl + "mover" };
            var target = mathMlRoot
                .DescendantsAndSelf()
                .Where(element => targetNames.Contains(element.Name)
                    && !usedTargets.Contains(element))
                .FirstOrDefault(element =>
                {
                    var children = element.Elements().ToArray();
                    return children.Length >= 2
                        && string.Equals(
                            FlattenMathMlTokenText(children[0]),
                            baseText,
                            StringComparison.Ordinal);
                });
            if (target is null) continue;

            var targetChildren = target.Elements().ToArray();
            targetChildren[0].ReplaceWith(
                new XElement(
                    mathMl + "mi",
                    new XAttribute("mathvariant", "normal"),
                    baseText));
            usedTargets.Add(target);
        }
    }

    internal static void RestoreOmmlScriptBaseRunSemantics(string omml, XElement mathMlRoot)
    {
        XNamespace officeMath = MathNamespace;
        XNamespace mathMl = "http://www.w3.org/1998/Math/MathML";
        var source = XDocument.Parse(omml, LoadOptions.PreserveWhitespace);
        var scripts = source
            .Descendants()
            .Where(element => element.Name == officeMath + "sSub"
                || element.Name == officeMath + "sSup"
                || element.Name == officeMath + "sSubSup")
            .ToArray();
        if (scripts.Length == 0) return;

        var usedTargets = new HashSet<XElement>();
        foreach (var script in scripts)
        {
            var baseSlot = script.Element(officeMath + "e");
            if (baseSlot is null) continue;
            var baseText = string.Concat(
                baseSlot.Descendants(officeMath + "t").Select(text => text.Value));
            if (string.IsNullOrWhiteSpace(baseText)) continue;

            XName targetName;
            if (script.Name == officeMath + "sSub") targetName = mathMl + "msub";
            else if (script.Name == officeMath + "sSup") targetName = mathMl + "msup";
            else targetName = mathMl + "msubsup";

            var target = mathMlRoot
                .DescendantsAndSelf(targetName)
                .Where(element => !usedTargets.Contains(element))
                .FirstOrDefault(element =>
                {
                    var children = element.Elements().ToArray();
                    return children.Length >= 2
                        && string.Equals(
                            FlattenMathMlTokenText(children[0]),
                            baseText,
                            StringComparison.Ordinal);
                });
            if (target is null) continue;
            // Claim the corresponding base even when no repair is required.
            // A preceding prose/operator/mixed base with the same text must not
            // steal a later upright identifier's target.
            usedTargets.Add(target);

            // This recovery joins one native identifier run split by Office's
            // projection. A descendant run inside a radical, fraction or nested
            // script does not own the enclosing base slot.
            var baseContents = baseSlot.Elements()
                .Where(element => element.Name != officeMath + "ePr")
                .ToArray();
            if (baseContents.Length != 1 || baseContents[0].Name != officeMath + "r")
                continue;
            var properties = baseContents[0].Element(officeMath + "rPr");
            var projectedBase = target.Elements().First();
            if (IsMathBooleanEnabled(properties, officeMath + "nor"))
            {
                // Normal Text is explicit role authority, including a prose
                // symbol promoted by RestoreMathSymbolTextTokens earlier.
                // Keep projected typography and wrappers while restoring prose.
                foreach (var token in projectedBase.DescendantsAndSelf().Where(element =>
                             element.Name == mathMl + "mi" || element.Name == mathMl + "mo"
                             || element.Name == mathMl + "mn"))
                    token.Name = mathMl + "mtext";
                continue;
            }
            var style = (string?)properties?.Element(officeMath + "sty")?.Attribute(officeMath + "val");
            if (!string.Equals(style, "p", StringComparison.OrdinalIgnoreCase)) continue;

            // Plain style describes typography, not token role. Only a projected
            // run of upright identifiers can be joined into an upright mi.
            // Retain projected mo, mn, mtext and structural nodes unchanged.
            if (projectedBase.DescendantsAndSelf().Any(element =>
                    element.Name != mathMl + "mrow" && element.Name != mathMl + "mi"))
                continue;
            var identifiers = projectedBase.DescendantsAndSelf(mathMl + "mi").ToArray();
            if (identifiers.Length == 0 || identifiers.Any(identifier =>
                    MathMlTokenStyleContract.ResolveEffectiveVariant(identifier) != "normal"))
                continue;
            // The joined identifier keeps the run's color (issue #35: a red
            // \mathrm{M}_{...} lost its red); differently colored parts stay apart.
            var color = MathMlColorContract.Effective(identifiers[0]);
            if (identifiers.Any(identifier => MathMlColorContract.Effective(identifier) != color))
                continue;
            projectedBase.ReplaceWith(
                new XElement(
                    mathMl + "mi",
                    new XAttribute("mathvariant", "normal"),
                    color is not null && color != MathMlColorContract.Effective(target)
                        ? new XAttribute("mathcolor", color)
                        : null,
                    baseText));
        }
    }

    internal static void RestoreOmmlFunctionApplicationRuns(string omml, XElement mathMlRoot)
    {
        XNamespace officeMath = MathNamespace;
        XNamespace mathMl = "http://www.w3.org/1998/Math/MathML";
        var source = XDocument.Parse(omml, LoadOptions.PreserveWhitespace);
        var sourceRuns = source.Descendants(officeMath + "r").ToArray();
        for (var index = 0; index + 1 < sourceRuns.Length; index++)
        {
            var run = sourceRuns[index];
            var properties = run.Element(officeMath + "rPr");
            var style = properties?
                .Element(officeMath + "sty")?
                .Attribute(officeMath + "val")?
                .Value;
            var explicitNormal = properties?.Element(officeMath + "nor") is not null
                || string.Equals(style, "p", StringComparison.OrdinalIgnoreCase);
            if (!explicitNormal) continue;

            var functionName = string.Concat(
                run.Elements(officeMath + "t").Select(text => text.Value));
            if (string.IsNullOrWhiteSpace(functionName)) continue;

            var nextRun = sourceRuns[index + 1];
            var nextText = string.Concat(
                nextRun.Elements(officeMath + "t").Select(text => text.Value));
            if (nextText.Length == 0 || nextText[0] != '\u2061') continue;

            CollapseMathMlUprightSequence(mathMlRoot, mathMl, functionName);
        }
    }

    private static bool CollapseMathMlUprightSequence(
        XElement root,
        XNamespace mathMl,
        string expectedText)
    {
        foreach (var parent in root.DescendantsAndSelf().ToArray())
        {
            var children = parent.Elements().ToArray();
            for (var start = 0; start < children.Length; start++)
            {
                var matched = new List<XElement>();
                var builder = new StringBuilder();
                for (var cursor = start; cursor < children.Length; cursor++)
                {
                    var child = children[cursor];
                    if (child.Name != mathMl + "mi") break;
                    var variant = child.Attribute("mathvariant")?.Value ?? string.Empty;
                    if (variant.IndexOf("normal", StringComparison.OrdinalIgnoreCase) < 0
                        && variant.IndexOf("upright", StringComparison.OrdinalIgnoreCase) < 0)
                        break;
                    matched.Add(child);
                    builder.Append(child.Value);
                    var candidate = builder.ToString();
                    if (!expectedText.StartsWith(candidate, StringComparison.Ordinal)) break;
                    if (!string.Equals(candidate, expectedText, StringComparison.Ordinal)) continue;

                    var replacement = new XElement(
                        mathMl + "mi",
                        new XAttribute("mathvariant", "normal"),
                        expectedText);
                    matched[0].AddBeforeSelf(replacement);
                    foreach (var element in matched) element.Remove();
                    return true;
                }
            }
        }
        return false;
    }

    private static bool IsStandardTexOperatorName(string value)
    {
        // Plain TeX/LaTeX operator names whose conventional math presentation is
        // upright. Keep this intentionally finite; arbitrary alphabetic runs must
        // remain variables even when Word happens to place them in a limit object.
        return value is "arccos" or "arcsin" or "arctan" or "arg"
            or "cos" or "cosh" or "cot" or "coth" or "csc"
            or "deg" or "det" or "dim" or "exp" or "gcd" or "hom"
            or "inf" or "ker" or "lg" or "lim" or "liminf" or "limsup"
            or "ln" or "log" or "max" or "min" or "Pr" or "sec"
            or "sin" or "sinh" or "sup" or "tan" or "tanh";
    }

    private static string FlattenMathMlTokenText(XElement element)
    {
        XNamespace mathMl = "http://www.w3.org/1998/Math/MathML";
        if (element.Name.Namespace != mathMl) return string.Empty;
        if (element.Name.LocalName is "mi" or "mn" or "mo" or "mtext")
            return element.Value;
        if (element.Name.LocalName is not ("mrow" or "mstyle" or "mpadded" or "semantics"))
            return string.Empty;
        var builder = new StringBuilder();
        foreach (var child in element.Elements())
        {
            var value = FlattenMathMlTokenText(child);
            if (value.Length == 0 && child.HasElements) return string.Empty;
            builder.Append(value);
        }
        return builder.ToString();
    }

    private static string CanonicalMathMlAccentCharacter(string? officeCharacter) =>
        officeCharacter switch
        {
            null or "" => "^", // OMML's omitted m:chr means the default hat.
            "\u0302" => "^",
            "\u0303" => "~",
            "\u20D7" => "→",
            "\u20D6" => "←",
            "\u20E1" => "↔",
            "\u0305" => "¯",
            "\u0307" => "˙",
            "\u0308" => "¨",
            "\u030C" => "ˇ",
            "\u0306" => "˘",
            "\u0301" => "´",
            "\u0300" => "`",
            "\u030A" => "˚",
            _ => officeCharacter,
        };

    internal static XElement CanonicalizeMathMlElement(XElement source)
    {
        XNamespace mathMl = "http://www.w3.org/1998/Math/MathML";
        var result = new XElement(mathMl + source.Name.LocalName);
        foreach (var attribute in source.Attributes())
        {
            if (attribute.IsNamespaceDeclaration) continue;
            result.SetAttributeValue(attribute.Name, attribute.Value);
        }
        foreach (var node in source.Nodes())
        {
            switch (node)
            {
                case XElement child:
                    result.Add(CanonicalizeMathMlElement(child));
                    break;
                case XCData cdata:
                    result.Add(new XCData(cdata.Value));
                    break;
                case XText text:
                    result.Add(new XText(text.Value));
                    break;
            }
        }
        return result;
    }
}
