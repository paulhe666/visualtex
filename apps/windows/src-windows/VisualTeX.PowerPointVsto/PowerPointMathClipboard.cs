using System.Text;
using System.Xml;
using System.Xml.Linq;

namespace VisualTeX.PowerPointVsto;

/// <summary>
/// PowerPoint's native mathematical TEXT clipboard package ("Art::Text
/// ClipFormat"): a DrawingML text fragment holding one a14:m equation, which
/// TextRange.Paste turns into a native equation in the target text. This is the
/// same package VisualTeX for macOS pastes (src-tauri office/powerpoint_omml.rs,
/// clipboard_entries and build_stored_zip): the same OMML in gives the same bytes
/// out on both platforms.
/// </summary>
internal static class PowerPointMathClipboard
{
    internal const string Format = "Art::Text ClipFormat";

    private const string M = "http://schemas.openxmlformats.org/officeDocument/2006/math";
    private const string A = "http://schemas.openxmlformats.org/drawingml/2006/main";
    private const string R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    private const string A14 = "http://schemas.microsoft.com/office/drawing/2010/main";
    private const string LC = "http://schemas.openxmlformats.org/drawingml/2006/lockedCanvas";
    private const string Rel = "http://schemas.openxmlformats.org/package/2006/relationships";

    private sealed class Node
    {
        internal string Name = string.Empty;
        internal readonly List<KeyValuePair<string, string>> Attributes = new();
        internal readonly List<Node> Children = new();
        internal string? Text;
    }

    /// <summary>The complete package (a stored ZIP) for one equation.</summary>
    internal static byte[] Build(string omml, double fontSize, bool display) =>
        BuildStoredZip(Entries(omml, fontSize, display));

    /// <summary>
    /// An edit's package: the edited equation's native zone (<paramref name="zone"/>,
    /// its a14:m element as saved in the slide) with the new equation in place of
    /// its own, as VisualTeX for macOS pastes an edit into the copied original
    /// context (replace_native_edit_context). The equation keeps its zone's
    /// justification and inline or display form.
    /// </summary>
    internal static byte[] BuildEdit(string omml, double fontSize, string zone) =>
        BuildStoredZip(EditEntries(omml, fontSize, zone));

    internal static List<KeyValuePair<string, string>> EditEntries(string omml, double fontSize, string zone)
    {
        const string drawingPart = "clipboard/drawings/drawing1.xml";
        var entries = Entries(omml, fontSize, display: false);
        var slot = entries.FindIndex(entry => entry.Key == drawingPart);
        var drawing = entries[slot].Value;
        var start = drawing.IndexOf("<a14:m>", StringComparison.Ordinal);
        var end = drawing.IndexOf("</a14:m>", start, StringComparison.Ordinal) + "</a14:m>".Length;
        var generated = XElement.Parse(
            $"<root xmlns:a=\"{A}\" xmlns:a14=\"{A14}\" xmlns:m=\"{M}\">{drawing.Substring(start, end - start)}</root>");
        var math = generated.Descendants(XName.Get("oMath", M)).Single();
        XElement original;
        try
        {
            original = XElement.Parse(zone, LoadOptions.PreserveWhitespace);
        }
        catch (XmlException error)
        {
            throw new InvalidDataException("The edited native equation's zone is not valid XML.", error);
        }
        if (original.Name != XName.Get("m", A14))
            throw new InvalidDataException("The edited equation is not inside a native PowerPoint math zone.");
        var equations = original.Descendants(XName.Get("oMath", M)).ToArray();
        if (equations.Length != 1)
            throw new InvalidDataException("The edited native math zone does not hold exactly one equation.");
        equations[0].ReplaceWith(math);
        foreach (var (prefix, name) in new[] { ("a", A), ("a14", A14), ("m", M) })
            if (original.Attribute(XNamespace.Xmlns + prefix) is null)
                original.SetAttributeValue(XNamespace.Xmlns + prefix, name);
        entries[slot] = new KeyValuePair<string, string>(drawingPart,
            drawing.Substring(0, start) + original.ToString(SaveOptions.DisableFormatting) + drawing.Substring(end));
        return entries;
    }

    internal static List<KeyValuePair<string, string>> Entries(string omml, double fontSize, bool display)
    {
        if (double.IsNaN(fontSize) || double.IsInfinity(fontSize) || fontSize < 1 || fontSize > 512)
            throw new InvalidDataException("Invalid native equation font size");
        var size = (long)Math.Round(fontSize * 100.0, MidpointRounding.AwayFromZero);
        var math = Serialize(Parse(omml), size);
        var paragraphs = new StringBuilder("<a:p>");
        if (display) paragraphs.Append("<a:pPr algn=\"ctr\"><a:buNone/></a:pPr>");
        paragraphs.Append("<a14:m><m:oMathPara><m:oMathParaPr><m:jc m:val=\"centerGroup\"/></m:oMathParaPr>")
            .Append(math)
            .Append("</m:oMathPara></a14:m><a:endParaRPr sz=\"").Append(size).Append("\"/></a:p>");
        const string xfrm = "<a:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"0\" cy=\"0\"/></a:xfrm>";
        var drawing =
            $"<a:graphic xmlns:a=\"{A}\" xmlns:a14=\"{A14}\" xmlns:m=\"{M}\"><a:graphicData uri=\"{LC}\"><lc:lockedCanvas xmlns:lc=\"{LC}\"><a:nvGrpSpPr><a:cNvPr id=\"0\" name=\"\"/><a:cNvGrpSpPr/></a:nvGrpSpPr><a:grpSpPr><a:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"0\" cy=\"0\"/><a:chOff x=\"0\" y=\"0\"/><a:chExt cx=\"0\" cy=\"0\"/></a:xfrm></a:grpSpPr><a:sp><a:nvSpPr><a:cNvPr id=\"0\" name=\"\"/><a:cNvSpPr/></a:nvSpPr><a:spPr>{xfrm}<a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom></a:spPr><a:txSp><a:txBody><a:bodyPr><a:noAutofit/></a:bodyPr><a:lstStyle/>{paragraphs}</a:txBody>{xfrm}</a:txSp></a:sp></lc:lockedCanvas></a:graphicData></a:graphic>";
        return new List<KeyValuePair<string, string>>
        {
            new("[Content_Types].xml", "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\"><Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/><Default Extension=\"xml\" ContentType=\"application/xml\"/><Override PartName=\"/clipboard/drawings/drawing1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.drawing+xml\"/><Override PartName=\"/clipboard/theme/theme1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.theme+xml\"/></Types>"),
            new("_rels/.rels", Rels("drawing", "clipboard/drawings/drawing1.xml")),
            new("clipboard/drawings/drawing1.xml", drawing),
            new("clipboard/drawings/_rels/drawing1.xml.rels", Rels("theme", "../theme/theme1.xml")),
            new("clipboard/theme/theme1.xml", Theme()),
        };
    }

    private static string Escape(string value)
    {
        var builder = new StringBuilder(value.Length);
        foreach (var c in value)
        {
            switch (c)
            {
                case '<': builder.Append("&lt;"); break;
                case '>': builder.Append("&gt;"); break;
                case '&': builder.Append("&amp;"); break;
                case '\'': builder.Append("&apos;"); break;
                case '"': builder.Append("&quot;"); break;
                default: builder.Append(c); break;
            }
        }
        return builder.ToString();
    }

    private static Node Parse(string xml)
    {
        var settings = new XmlReaderSettings
        {
            DtdProcessing = DtdProcessing.Prohibit,
            IgnoreComments = true,
            IgnoreProcessingInstructions = false,
            XmlResolver = null,
        };
        var stack = new Stack<Node>();
        var holder = new Node();
        stack.Push(holder);
        using (var reader = XmlReader.Create(new StringReader(xml), settings))
        {
            while (reader.Read())
            {
                switch (reader.NodeType)
                {
                    case XmlNodeType.Element:
                    {
                        var node = new Node { Name = reader.Name };
                        var empty = reader.IsEmptyElement;
                        if (reader.MoveToFirstAttribute())
                        {
                            do
                            {
                                if (reader.Name == "xmlns:m" && reader.Value != M)
                                    throw new InvalidDataException("Invalid Office Math namespace");
                                node.Attributes.Add(new KeyValuePair<string, string>(reader.Name, reader.Value));
                            }
                            while (reader.MoveToNextAttribute());
                            reader.MoveToElement();
                        }
                        stack.Peek().Children.Add(node);
                        if (!empty) stack.Push(node);
                        if (stack.Count > 256) throw new InvalidDataException("Office Math XML is nested too deeply");
                        break;
                    }
                    case XmlNodeType.EndElement:
                        if (stack.Count < 2) throw new InvalidDataException("Unbalanced Office Math XML");
                        stack.Pop();
                        break;
                    case XmlNodeType.Text:
                    case XmlNodeType.CDATA:
                    case XmlNodeType.Whitespace:
                    case XmlNodeType.SignificantWhitespace:
                        if (stack.Count > 1) stack.Peek().Children.Add(new Node { Text = reader.Value });
                        else if (!string.IsNullOrWhiteSpace(reader.Value))
                            throw new InvalidDataException("Unexpected text outside Office Math");
                        break;
                    case XmlNodeType.ProcessingInstruction:
                    case XmlNodeType.DocumentType:
                        throw new InvalidDataException("Office Math XML cannot contain declarations or processing instructions");
                }
            }
        }
        if (stack.Count != 1 || holder.Children.Count != 1)
            throw new InvalidDataException("Expected exactly one Office Math root");
        var root = holder.Children[0];
        if (root.Name != "m:oMath" || !root.Attributes.Any(a => a.Key == "xmlns:m" && a.Value == M))
            throw new InvalidDataException("Expected a namespaced m:oMath root");
        return root;
    }

    private static string? Property(Node node, string name, string key) =>
        node.Children.FirstOrDefault(child => child.Name == name)?
            .Attributes.Where(a => a.Key == key).Select(a => a.Value).FirstOrDefault();

    private static string RunProperties(Node? word, long size)
    {
        var flags = new StringBuilder();
        if (word is not null)
        {
            foreach (var (w, a) in new[] { ("w:b", "b"), ("w:i", "i") })
            {
                var node = word.Children.FirstOrDefault(child => child.Name == w);
                if (node is null) continue;
                var enabled = !node.Attributes.Any(attribute =>
                    attribute.Key == "w:val" && (attribute.Value == "0" || attribute.Value == "false" || attribute.Value == "off"));
                flags.Append(' ').Append(a).Append("=\"").Append(enabled ? 1 : 0).Append('"');
            }
        }
        var xml = new StringBuilder($"<a:rPr sz=\"{size}\"{flags}>");
        var color = word is null ? null : Property(word, "w:color", "w:val");
        if (color is { Length: 6 } && color.All(Uri.IsHexDigit))
            xml.Append("<a:solidFill><a:srgbClr val=\"").Append(color).Append("\"/></a:solidFill>");
        foreach (var (key, tag, fallback) in new[]
                 {
                     ("w:ascii", "a:latin", (string?)"Cambria Math"),
                     ("w:eastAsia", "a:ea", null),
                     ("w:cs", "a:cs", null),
                 })
        {
            var font = (word is null ? null : Property(word, "w:rFonts", key)) ?? fallback;
            if (font is not null) xml.Append('<').Append(tag).Append(" typeface=\"").Append(Escape(font)).Append("\"/>");
        }
        return xml.Append("</a:rPr>").ToString();
    }

    // OMML properties of on/off type (ST_OnOff). PowerPoint's native clipboard
    // importer drops their numeric values (a radical's degree, an operator's empty
    // limit and a matrix's placeholders reappear), so they are written in its own
    // canonical on/off form.
    private static readonly HashSet<string> OnOff = new(StringComparer.Ordinal)
    {
        "m:aln", "m:alnScr", "m:degHide", "m:diff", "m:dispDef", "m:grow", "m:hideBot", "m:hideLeft",
        "m:hideRight", "m:hideTop", "m:lit", "m:maxDist", "m:noBreak", "m:nor", "m:objDist", "m:opEmu",
        "m:plcHide", "m:show", "m:smallFrac", "m:strikeBLTR", "m:strikeH", "m:strikeTLBR", "m:strikeV",
        "m:subHide", "m:supHide", "m:transp", "m:wrapRight", "m:zeroAsc", "m:zeroDesc", "m:zeroWid",
    };

    private static string Serialize(Node node, long size)
    {
        if (node.Text is not null) return Escape(node.Text);
        if (node.Name == "w:rPr") return RunProperties(node, size);
        if (!node.Name.StartsWith("m:", StringComparison.Ordinal))
            throw new InvalidDataException($"Unsupported element in native math: {node.Name}");
        var xml = new StringBuilder("<").Append(node.Name);
        foreach (var attribute in node.Attributes)
        {
            // The shared writer also declares MathML's namespace; no element uses it.
            if (attribute.Key == "xmlns:w" || attribute.Key == "xmlns:mml") continue;
            var value = attribute.Key == "m:val" && OnOff.Contains(node.Name)
                ? attribute.Value switch { "1" or "true" => "on", "0" or "false" => "off", _ => attribute.Value }
                : attribute.Value;
            xml.Append(' ').Append(attribute.Key).Append("=\"").Append(Escape(value)).Append('"');
        }
        xml.Append('>');
        var inserted = node.Children.Any(child => child.Name == "w:rPr");
        foreach (var child in node.Children)
        {
            if (node.Name == "m:r" && !inserted && child.Name != "m:rPr" && child.Text is null)
            {
                xml.Append(RunProperties(null, size));
                inserted = true;
            }
            xml.Append(Serialize(child, size));
        }
        return xml.Append("</").Append(node.Name).Append('>').ToString();
    }

    private static string Rels(string kind, string target) =>
        $"<Relationships xmlns=\"{Rel}\"><Relationship Id=\"rId1\" Type=\"{R}/{kind}\" Target=\"{target}\"/></Relationships>";

    private static string Theme()
    {
        var colors = string.Concat(new[]
        {
            ("dk1", "000000"), ("lt1", "FFFFFF"), ("dk2", "44546A"), ("lt2", "E7E6E6"), ("accent1", "4472C4"),
            ("accent2", "ED7D31"), ("accent3", "A5A5A5"), ("accent4", "FFC000"), ("accent5", "5B9BD5"),
            ("accent6", "70AD47"), ("hlink", "0563C1"), ("folHlink", "954F72"),
        }.Select(c => $"<a:{c.Item1}><a:srgbClr val=\"{c.Item2}\"/></a:{c.Item1}>"));
        const string fill = "<a:solidFill><a:schemeClr val=\"phClr\"/></a:solidFill>";
        const string fonts = "<a:latin typeface=\"Cambria Math\"/><a:ea typeface=\"\"/><a:cs typeface=\"\"/>";
        string Repeat(string value) => string.Concat(Enumerable.Repeat(value, 3));
        return $"<a:clipboardTheme xmlns:a=\"{A}\"><a:themeElements><a:clrScheme name=\"Office\">{colors}</a:clrScheme><a:fontScheme name=\"Office\"><a:majorFont>{fonts}</a:majorFont><a:minorFont>{fonts}</a:minorFont></a:fontScheme><a:fmtScheme name=\"Office\"><a:fillStyleLst>{Repeat(fill)}</a:fillStyleLst><a:lnStyleLst>{Repeat($"<a:ln w=\"12700\">{fill}<a:prstDash val=\"solid\"/></a:ln>")}</a:lnStyleLst><a:effectStyleLst>{Repeat("<a:effectStyle><a:effectLst/></a:effectStyle>")}</a:effectStyleLst><a:bgFillStyleLst>{Repeat(fill)}</a:bgFillStyleLst></a:fmtScheme></a:themeElements><a:clrMap bg1=\"lt1\" tx1=\"dk1\" bg2=\"lt2\" tx2=\"dk2\" accent1=\"accent1\" accent2=\"accent2\" accent3=\"accent3\" accent4=\"accent4\" accent5=\"accent5\" accent6=\"accent6\" hlink=\"hlink\" folHlink=\"folHlink\"/></a:clipboardTheme>";
    }

    // build_stored_zip: stored entries, UTF-8 names, DOS date 1980-01-01.
    internal static byte[] BuildStoredZip(IReadOnlyList<KeyValuePair<string, string>> entries)
    {
        using var output = new MemoryStream();
        using var writer = new BinaryWriter(output);
        var records = new List<(byte[] Name, uint Crc, uint Size, uint Offset)>();
        foreach (var entry in entries)
        {
            var name = Encoding.UTF8.GetBytes(entry.Key);
            var contents = Encoding.UTF8.GetBytes(entry.Value);
            var offset = (uint)output.Position;
            var crc = Crc32(contents);
            writer.Write(0x04034b50u); writer.Write((ushort)20); writer.Write((ushort)0x0800);
            writer.Write((ushort)0); writer.Write((ushort)0); writer.Write((ushort)33);
            writer.Write(crc); writer.Write((uint)contents.Length); writer.Write((uint)contents.Length);
            writer.Write((ushort)name.Length); writer.Write((ushort)0);
            writer.Write(name); writer.Write(contents);
            records.Add((name, crc, (uint)contents.Length, offset));
        }
        var centralOffset = (uint)output.Position;
        foreach (var record in records)
        {
            writer.Write(0x02014b50u); writer.Write((ushort)20); writer.Write((ushort)20); writer.Write((ushort)0x0800);
            writer.Write((ushort)0); writer.Write((ushort)0); writer.Write((ushort)33);
            writer.Write(record.Crc); writer.Write(record.Size); writer.Write(record.Size);
            writer.Write((ushort)record.Name.Length); writer.Write((ushort)0); writer.Write((ushort)0);
            writer.Write((ushort)0); writer.Write((ushort)0); writer.Write(0u); writer.Write(record.Offset);
            writer.Write(record.Name);
        }
        var centralSize = (uint)output.Position - centralOffset;
        writer.Write(0x06054b50u); writer.Write((ushort)0); writer.Write((ushort)0);
        writer.Write((ushort)records.Count); writer.Write((ushort)records.Count);
        writer.Write(centralSize); writer.Write(centralOffset); writer.Write((ushort)0);
        writer.Flush();
        return output.ToArray();
    }

    private static uint Crc32(byte[] bytes)
    {
        var crc = 0xffffffffu;
        foreach (var value in bytes)
        {
            crc ^= value;
            for (var bit = 0; bit < 8; bit++)
            {
                var mask = 0u - (crc & 1);
                crc = (crc >> 1) ^ (0xedb88320u & mask);
            }
        }
        return ~crc;
    }
}
