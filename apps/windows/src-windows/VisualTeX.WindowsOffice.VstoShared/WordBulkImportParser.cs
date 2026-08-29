using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Text.RegularExpressions;
using VisualTeX.WindowsOffice.Contracts;

namespace VisualTeX.WindowsOffice.VstoShared;

internal enum WordBulkSourceFormat
{
    Auto,
    Markdown,
    Latex,
}

internal enum WordBulkFormulaObjectMode
{
    Omml,
    Ole,
    MathType,
}

internal enum WordBulkBlockKind
{
    Paragraph,
    Heading,
    Bullet,
    Numbered,
    Quote,
    Code,
    DisplayFormula,
}

internal sealed class WordBulkRun
{
    internal string Id { get; set; } = Guid.NewGuid().ToString("D");
    internal bool IsFormula { get; set; }
    internal string Text { get; set; } = string.Empty;
    internal string Latex { get; set; } = string.Empty;
    internal bool Bold { get; set; }
    internal bool Italic { get; set; }
    internal bool Code { get; set; }
    internal bool Strike { get; set; }
    internal bool Underline { get; set; }
    internal string DisplayMode { get; set; } = "inline";
    internal string? EquationTag { get; set; }
}

internal sealed class WordBulkBlock
{
    internal WordBulkBlockKind Kind { get; set; }
    internal int Level { get; set; }
    internal List<WordBulkRun> Runs { get; set; } = new();
}

internal sealed class WordBulkImportDocument
{
    internal WordBulkSourceFormat SourceFormat { get; set; }
    internal WordBulkFormulaObjectMode FormulaObjectMode { get; set; }
    internal List<WordBulkBlock> Blocks { get; set; } = new();
    internal List<string> Warnings { get; set; } = new();
    internal int FormulaCount => Blocks.Sum(block => block.Runs.Count(run => run.IsFormula));
    internal int InlineFormulaCount => Blocks.Sum(block =>
        block.Runs.Count(run => run.IsFormula && run.DisplayMode == "inline"));
    internal int DisplayFormulaCount => Blocks.Sum(block =>
        block.Runs.Count(run => run.IsFormula && run.DisplayMode == "block"));
    internal int TextCharacterCount => Blocks.Sum(block =>
        block.Runs.Where(run => !run.IsFormula).Sum(run => run.Text.Length));
}

internal sealed class WordLatexFormulaSpan
{
    internal string Id { get; set; } = Guid.NewGuid().ToString("D");
    internal int Start { get; set; }
    internal int Length { get; set; }
    internal string Latex { get; set; } = string.Empty;
    internal string DisplayMode { get; set; } = "inline";
}

internal static class WordBulkImportParser
{
    private sealed class SerializedDocument
    {
        [JsonPropertyName("format")]
        public string Format { get; set; } = "markdown";

        [JsonPropertyName("blocks")]
        public List<SerializedBlock> Blocks { get; set; } = new();

        [JsonPropertyName("warnings")]
        public List<string> Warnings { get; set; } = new();
    }

    private sealed class SerializedBlock
    {
        [JsonPropertyName("kind")]
        public string Kind { get; set; } = "paragraph";

        [JsonPropertyName("level")]
        public int Level { get; set; }

        [JsonPropertyName("runs")]
        public List<SerializedRun> Runs { get; set; } = new();
    }

    private sealed class SerializedRun
    {
        [JsonPropertyName("kind")]
        public string Kind { get; set; } = "text";

        [JsonPropertyName("text")]
        public string? Text { get; set; }

        [JsonPropertyName("latex")]
        public string? Latex { get; set; }

        [JsonPropertyName("display")]
        public bool Display { get; set; }

        [JsonPropertyName("equationTag")]
        public string? EquationTag { get; set; }

        [JsonPropertyName("bold")]
        public bool Bold { get; set; }

        [JsonPropertyName("italic")]
        public bool Italic { get; set; }

        [JsonPropertyName("code")]
        public bool Code { get; set; }

        [JsonPropertyName("strike")]
        public bool Strike { get; set; }

        [JsonPropertyName("underline")]
        public bool Underline { get; set; }
    }

    private static readonly Regex MarkdownHeading = new(
        @"^(?<marks>#{1,6})\s+(?<text>.+?)\s*#*\s*$",
        RegexOptions.Compiled);
    private static readonly Regex MarkdownBullet = new(
        @"^(?<indent>\s*)[-+*]\s+(?<text>.+)$",
        RegexOptions.Compiled);
    private static readonly Regex MarkdownNumbered = new(
        @"^(?<indent>\s*)\d+[.)]\s+(?<text>.+)$",
        RegexOptions.Compiled);
    private static readonly Regex LatexSection = new(
        @"^\s*\\(?<kind>part|chapter|section|subsection|subsubsection|paragraph|subparagraph)\*?\{(?<text>.*)\}\s*$",
        RegexOptions.Compiled | RegexOptions.IgnoreCase);
    private static readonly Regex LatexEnvironmentStart = new(
        @"^\s*\\begin\{(?<name>equation\*?|align\*?|gather\*?|multline\*?|displaymath)\}\s*$",
        RegexOptions.Compiled | RegexOptions.IgnoreCase);
    private static readonly Regex LatexItem = new(
        @"^\s*\\item(?:\s*\[[^\]]*\])?\s*(?<text>.*)$",
        RegexOptions.Compiled);
    private static readonly Regex LatexTheoremEnvironmentStart = new(
        @"^\s*\\begin\{(?<name>theorem|lemma|proposition|corollary|definition|proof|remark|example|exercise|assumption|axiom|claim|conjecture|criterion|fact|notation|observation|problem|question|solution)\*?\}(?:\s*\[(?<title>[^\]]*)\])?\s*$",
        RegexOptions.Compiled | RegexOptions.IgnoreCase);
    private static readonly Regex LatexTheoremEnvironmentEnd = new(
        @"^\s*\\end\{(?:theorem|lemma|proposition|corollary|definition|proof|remark|example|exercise|assumption|axiom|claim|conjecture|criterion|fact|notation|observation|problem|question|solution)\*?\}\s*$",
        RegexOptions.Compiled | RegexOptions.IgnoreCase);

    internal static WordBulkImportDocument ParseSerialized(
        string serialized,
        WordBulkFormulaObjectMode objectMode)
    {
        if (string.IsNullOrWhiteSpace(serialized))
            throw new InvalidDataException("批量导入没有返回可用的文档结构。");
        if (serialized.Length > 5_000_000)
            throw new InvalidDataException("批量导入文档结构不能超过 5 MB。");

        SerializedDocument? wire;
        try
        {
            wire = JsonSerializer.Deserialize<SerializedDocument>(
                serialized,
                new JsonSerializerOptions
                {
                    PropertyNameCaseInsensitive = true,
                    MaxDepth = 128,
                });
        }
        catch (JsonException error)
        {
            throw new InvalidDataException("批量导入返回了无效的文档结构。", error);
        }
        if (wire is null || wire.Blocks.Count == 0)
            throw new InvalidDataException("批量导入没有找到可以插入 Word 的内容。");
        if (wire.Blocks.Count > 10_000)
            throw new InvalidDataException("批量导入包含过多段落（上限 10000）。");

        var blocks = new List<WordBulkBlock>(wire.Blocks.Count);
        foreach (var sourceBlock in wire.Blocks)
        {
            var kind = sourceBlock.Kind.Trim().ToLowerInvariant() switch
            {
                "heading" => WordBulkBlockKind.Heading,
                "bullet" => WordBulkBlockKind.Bullet,
                "numbered" => WordBulkBlockKind.Numbered,
                "quote" => WordBulkBlockKind.Quote,
                "code" => WordBulkBlockKind.Code,
                "display" => WordBulkBlockKind.DisplayFormula,
                _ => WordBulkBlockKind.Paragraph,
            };
            var runs = new List<WordBulkRun>();
            foreach (var sourceRun in sourceBlock.Runs)
            {
                if (string.Equals(sourceRun.Kind, "formula", StringComparison.OrdinalIgnoreCase))
                {
                    var latex = sourceRun.Latex?.Trim() ?? string.Empty;
                    if (latex.Length == 0) continue;
                    runs.Add(new WordBulkRun
                    {
                        IsFormula = true,
                        Latex = latex,
                        DisplayMode = sourceRun.Display || kind == WordBulkBlockKind.DisplayFormula
                            ? "block"
                            : "inline",
                        EquationTag = sourceRun.EquationTag,
                    });
                }
                else
                {
                    runs.Add(new WordBulkRun
                    {
                        Text = sourceRun.Text ?? string.Empty,
                        Bold = sourceRun.Bold,
                        Italic = sourceRun.Italic,
                        Code = sourceRun.Code,
                        Strike = sourceRun.Strike,
                        Underline = sourceRun.Underline,
                    });
                }
            }
            if (kind == WordBulkBlockKind.DisplayFormula)
            {
                var formula = runs.FirstOrDefault(run => run.IsFormula);
                if (formula is null) continue;
                blocks.Add(new WordBulkBlock
                {
                    Kind = kind,
                    Level = 0,
                    Runs = new List<WordBulkRun> { formula },
                });
                continue;
            }
            if (runs.Count == 0) runs.Add(new WordBulkRun { Text = string.Empty });
            blocks.Add(new WordBulkBlock
            {
                Kind = kind,
                Level = Math.Min(8, Math.Max(0, sourceBlock.Level)),
                Runs = runs,
            });
        }

        if (blocks.Count == 0)
            throw new InvalidDataException("批量导入没有找到可以插入 Word 的内容。");
        var formulaCount = blocks.Sum(block => block.Runs.Count(run => run.IsFormula));
        if (formulaCount > 1_000)
            throw new InvalidDataException("批量导入包含过多公式（上限 1000）。");

        return new WordBulkImportDocument
        {
            SourceFormat = string.Equals(wire.Format, "latex", StringComparison.OrdinalIgnoreCase)
                ? WordBulkSourceFormat.Latex
                : WordBulkSourceFormat.Markdown,
            FormulaObjectMode = objectMode,
            Blocks = blocks,
            Warnings = wire.Warnings
                .Where(warning => !string.IsNullOrWhiteSpace(warning))
                .Select(warning => warning.Trim())
                .Take(256)
                .ToList(),
        };
    }

    internal static WordBulkImportDocument Parse(
        string source,
        WordBulkSourceFormat sourceFormat,
        WordBulkFormulaObjectMode objectMode)
    {
        if (source is null) throw new ArgumentNullException(nameof(source));
        if (source.Length > 5_000_000)
            throw new InvalidDataException("批量导入内容不能超过 5 MB。");

        var format = sourceFormat == WordBulkSourceFormat.Auto
            ? DetectFormat(source)
            : sourceFormat;
        var warnings = new List<string>();
        var normalized = NormalizeSource(source, format, warnings);
        var blocks = ParseBlocks(normalized, format, warnings);
        if (blocks.Count == 0)
            throw new InvalidDataException("没有找到可以插入 Word 的文字或公式。");
        if (blocks.Count > 10_000)
            throw new InvalidDataException("批量导入包含过多段落（上限 10000）。");
        var formulaCount = blocks.Sum(block => block.Runs.Count(run => run.IsFormula));
        if (formulaCount > 1_000)
            throw new InvalidDataException("批量导入包含过多公式（上限 1000）。");

        return new WordBulkImportDocument
        {
            SourceFormat = format,
            FormulaObjectMode = objectMode,
            Blocks = blocks,
            Warnings = warnings,
        };
    }

    /// <summary>
    /// Finds LaTeX math delimiters without rewriting the surrounding Word text.
    /// This intentionally shares the delimiter semantics used by batch import,
    /// while preserving exact source offsets for in-place redraw.
    /// </summary>
    internal static List<WordLatexFormulaSpan> FindFormulaSpans(string source)
    {
        if (source is null) throw new ArgumentNullException(nameof(source));
        if (source.Length > 5_000_000)
            throw new InvalidDataException("LaTeX 重绘范围不能超过 5 MB。");

        var spans = new List<WordLatexFormulaSpan>();
        for (var index = 0; index < source.Length;)
        {
            if (source[index] == '$' && !IsEscaped(source, index))
            {
                if (index + 1 < source.Length && source[index + 1] == '$')
                {
                    var end = FindUnescapedSequence(source, "$$", index + 2);
                    if (end >= index + 2)
                    {
                        AddFormulaSpan(
                            spans,
                            source,
                            index,
                            end + 2,
                            index + 2,
                            end,
                            "block",
                            normalizeDisplay: true);
                        index = end + 2;
                        continue;
                    }
                }
                else
                {
                    var end = FindUnescaped(source, '$', index + 1);
                    if (end > index + 1)
                    {
                        AddFormulaSpan(
                            spans,
                            source,
                            index,
                            end + 1,
                            index + 1,
                            end,
                            "inline",
                            normalizeDisplay: false);
                        index = end + 1;
                        continue;
                    }
                }
            }

            if (source[index] == '\\'
                && index + 1 < source.Length
                && (source[index + 1] == '(' || source[index + 1] == '[')
                && !IsEscaped(source, index))
            {
                var display = source[index + 1] == '[';
                var endToken = display ? "\\]" : "\\)";
                var end = FindUnescapedSequence(source, endToken, index + 2);
                if (end > index + 2)
                {
                    AddFormulaSpan(
                        spans,
                        source,
                        index,
                        end + endToken.Length,
                        index + 2,
                        end,
                        display ? "block" : "inline",
                        normalizeDisplay: display);
                    index = end + endToken.Length;
                    continue;
                }
            }

            if (source[index] == '\\' && !IsEscaped(source, index))
            {
                var environment = Regex.Match(
                    source.Substring(index),
                    @"^\\begin\{(?<name>equation\*?|align\*?|gather\*?|multline\*?|displaymath)\}",
                    RegexOptions.IgnoreCase);
                if (environment.Success)
                {
                    var name = environment.Groups["name"].Value;
                    var endToken = $"\\end{{{name}}}";
                    var bodyStart = index + environment.Length;
                    var end = FindUnescapedSequence(source, endToken, bodyStart);
                    if (end >= bodyStart)
                    {
                        var latex = NormalizeDisplayEnvironmentLatex(
                            name,
                            source.Substring(bodyStart, end - bodyStart));
                        if (!string.IsNullOrWhiteSpace(latex))
                        {
                            spans.Add(new WordLatexFormulaSpan
                            {
                                Start = index,
                                Length = end + endToken.Length - index,
                                Latex = latex,
                                DisplayMode = "block",
                            });
                        }
                        index = end + endToken.Length;
                        continue;
                    }
                }
            }

            index++;
        }

        if (spans.Count > 1_000)
            throw new InvalidDataException("LaTeX 重绘包含过多公式（上限 1000）。");
        return spans;
    }

    private static void AddFormulaSpan(
        ICollection<WordLatexFormulaSpan> spans,
        string source,
        int sourceStart,
        int sourceEnd,
        int bodyStart,
        int bodyEnd,
        string displayMode,
        bool normalizeDisplay)
    {
        if (bodyEnd <= bodyStart) return;
        var body = source.Substring(bodyStart, bodyEnd - bodyStart);
        var latex = normalizeDisplay
            ? NormalizeDelimitedDisplayLatex(body)
            : body.Trim();
        if (latex.Length == 0) return;
        spans.Add(new WordLatexFormulaSpan
        {
            Start = sourceStart,
            Length = sourceEnd - sourceStart,
            Latex = latex,
            DisplayMode = displayMode,
        });
    }

    private static WordBulkSourceFormat DetectFormat(string source)
    {
        if (Regex.IsMatch(
                source,
                @"\\(?:documentclass|usepackage|RequirePackage|newcommand|renewcommand|providecommand|DeclareMathOperator\*?|DeclarePairedDelimiter\w*|newtheorem\*?|begin\{[A-Za-z@*]+\}|\[|\(|text(?:bf|it|tt)\{|emph\{|item(?:\s|\[)|(?:part|chapter|section|subsection|subsubsection|paragraph|subparagraph)\*?\{)",
                RegexOptions.IgnoreCase))
            return WordBulkSourceFormat.Latex;
        return WordBulkSourceFormat.Markdown;
    }

    private static bool TryReadLatexInlineLiteral(
        string text,
        int index,
        out string content,
        out int nextIndex)
    {
        content = string.Empty;
        nextIndex = index;
        if (index < 0 || index >= text.Length || text[index] != '\\' || IsEscaped(text, index))
            return false;
        var match = Regex.Match(
            text.Substring(index),
            @"^\\(?:verb|lstinline)\*?(?![A-Za-z@])(?:\[[^\]\r\n]*\])?",
            RegexOptions.IgnoreCase);
        if (!match.Success) return false;
        var delimiterIndex = index + match.Length;
        if (delimiterIndex >= text.Length) return false;
        var delimiter = text[delimiterIndex];
        if (char.IsLetterOrDigit(delimiter) || char.IsWhiteSpace(delimiter)) return false;
        var close = text.IndexOf(delimiter, delimiterIndex + 1);
        if (close < 0) return false;
        content = text.Substring(delimiterIndex + 1, close - delimiterIndex - 1);
        nextIndex = close + 1;
        return true;
    }

    private static int FindLatexInlineLiteralEnd(string text, int index)
    {
        if (TryReadLatexInlineLiteral(text, index, out _, out var nextIndex))
            return nextIndex;
        if (index < 0 || index >= text.Length || text[index] != '\\' || IsEscaped(text, index))
            return -1;
        var match = Regex.Match(
            text.Substring(index),
            @"^\\(?:verb|lstinline)\*?(?![A-Za-z@])(?:\[[^\]\r\n]*\])?",
            RegexOptions.IgnoreCase);
        if (!match.Success) return -1;
        var lineEnd = text.IndexOf('\n', index + match.Length);
        return lineEnd >= 0 ? lineEnd : text.Length;
    }

    private static int FindLatexDocumentToken(string source, string token, int startIndex = 0)
    {
        for (var index = Math.Max(0, startIndex); index < source.Length;)
        {
            var literalEnd = FindLatexInlineLiteralEnd(source, index);
            if (literalEnd > index)
            {
                index = literalEnd;
                continue;
            }
            if (source[index] == '%' && !IsEscaped(source, index))
            {
                var lineEnd = source.IndexOf('\n', index + 1);
                index = lineEnd >= 0 ? lineEnd + 1 : source.Length;
                continue;
            }
            if (source[index] == '\\' && !IsEscaped(source, index))
            {
                var literalEnvironment = Regex.Match(
                    source.Substring(index),
                    @"^\\begin\{(?<name>verbatim\*?|lstlisting\*?)\}(?:\[[^\]\r\n]*\])?",
                    RegexOptions.IgnoreCase);
                if (literalEnvironment.Success)
                {
                    var endToken = $"\\end{{{literalEnvironment.Groups["name"].Value}}}";
                    var environmentEnd = source.IndexOf(
                        endToken,
                        index + literalEnvironment.Length,
                        StringComparison.OrdinalIgnoreCase);
                    index = environmentEnd >= 0
                        ? environmentEnd + endToken.Length
                        : source.Length;
                    continue;
                }
                if (index <= source.Length - token.Length
                    && source.IndexOf(token, index, token.Length, StringComparison.OrdinalIgnoreCase) == index)
                    return index;
            }
            index++;
        }
        return -1;
    }

    private static int FindLatexCommentStart(string line)
    {
        for (var index = 0; index < line.Length; index++)
        {
            var literalEnd = FindLatexInlineLiteralEnd(line, index);
            if (literalEnd > index)
            {
                index = literalEnd - 1;
                continue;
            }
            if (line[index] == '%' && !IsEscaped(line, index)) return index;
        }
        return -1;
    }

    private static string NormalizeSource(
        string source,
        WordBulkSourceFormat format,
        ICollection<string> warnings)
    {
        var normalized = source
            .Replace("\r\n", "\n")
            .Replace('\r', '\n')
            .Trim('\uFEFF', ' ', '\t', '\n');
        if (format != WordBulkSourceFormat.Latex) return normalized;

        const string beginToken = "\\begin{document}";
        const string endToken = "\\end{document}";
        var begin = FindLatexDocumentToken(normalized, beginToken);
        if (begin >= 0)
        {
            var contentStart = begin + beginToken.Length;
            var end = FindLatexDocumentToken(normalized, endToken, contentStart);
            normalized = end >= 0
                ? normalized.Substring(contentStart, end - contentStart)
                : normalized.Substring(contentStart);
            if (end < 0) warnings.Add("LaTeX 文档缺少 \\end{document}，已导入其余内容。");
        }

        var lines = normalized.Split('\n');
        var cleaned = new StringBuilder(normalized.Length);
        var literalEnvironmentEnd = string.Empty;
        foreach (var line in lines)
        {
            var trimmed = line.Trim();
            if (!string.IsNullOrEmpty(literalEnvironmentEnd))
            {
                cleaned.Append(line).Append('\n');
                if (trimmed.Equals(literalEnvironmentEnd, StringComparison.OrdinalIgnoreCase))
                    literalEnvironmentEnd = string.Empty;
                continue;
            }

            var literalStart = Regex.Match(
                trimmed,
                @"^\\begin\{(?<name>verbatim|lstlisting)\}(?:\[[^\]]*\])?\s*$",
                RegexOptions.IgnoreCase);
            if (literalStart.Success)
            {
                literalEnvironmentEnd = $"\\end{{{literalStart.Groups["name"].Value}}}";
                cleaned.Append(line).Append('\n');
                continue;
            }

            var comment = FindLatexCommentStart(line);
            cleaned.Append(comment >= 0 ? line.Substring(0, comment) : line)
                .Append('\n');
        }
        return cleaned.ToString().Trim();
    }

    private static List<WordBulkBlock> ParseBlocks(
        string source,
        WordBulkSourceFormat format,
        ICollection<string> warnings)
    {
        var blocks = new List<WordBulkBlock>();
        var paragraph = new List<string>();
        var lines = source.Split('\n').ToList();
        var inCodeFence = false;
        var codeFenceEnd = string.Empty;
        var codeFenceDescription = string.Empty;
        var code = new StringBuilder();
        var listModes = new Stack<string>();
        var quote = new List<string>();
        var inLatexQuote = false;

        void FlushParagraph()
        {
            if (paragraph.Count == 0) return;
            var text = string.Join(" ", paragraph).Trim();
            paragraph.Clear();
            if (text.Length == 0) return;
            blocks.Add(new WordBulkBlock
            {
                Kind = WordBulkBlockKind.Paragraph,
                Runs = ParseInlineRuns(text, format, warnings),
            });
        }

        void FlushQuote()
        {
            if (quote.Count == 0) return;
            var text = string.Join(" ", quote).Trim();
            quote.Clear();
            if (text.Length == 0) return;
            blocks.Add(new WordBulkBlock
            {
                Kind = WordBulkBlockKind.Quote,
                Runs = ParseInlineRuns(text, format, warnings),
            });
        }

        void FinishCodeBlock(string? warning = null)
        {
            blocks.Add(new WordBulkBlock
            {
                Kind = WordBulkBlockKind.Code,
                Runs = new List<WordBulkRun>
                {
                    new() { Text = code.ToString().TrimEnd('\r', '\n'), Code = true },
                },
            });
            code.Clear();
            inCodeFence = false;
            codeFenceEnd = string.Empty;
            codeFenceDescription = string.Empty;
            if (!string.IsNullOrWhiteSpace(warning)) warnings.Add(warning!);
        }

        for (var index = 0; index < lines.Count; index++)
        {
            var raw = lines[index];
            var trimmed = raw.Trim();

            if (format == WordBulkSourceFormat.Markdown
                && trimmed.StartsWith("```", StringComparison.Ordinal))
            {
                FlushParagraph();
                FlushQuote();
                if (inCodeFence && codeFenceEnd == "```")
                {
                    FinishCodeBlock();
                }
                else if (!inCodeFence)
                {
                    inCodeFence = true;
                    codeFenceEnd = "```";
                    codeFenceDescription = "Markdown 代码块";
                }
                else
                {
                    code.AppendLine(raw);
                }
                continue;
            }
            if (format == WordBulkSourceFormat.Latex && !inCodeFence)
            {
                var codeStart = Regex.Match(
                    trimmed,
                    @"^\\begin\{(?<name>verbatim|lstlisting)\}(?:\[[^\]]*\])?\s*$",
                    RegexOptions.IgnoreCase);
                if (codeStart.Success)
                {
                    FlushParagraph();
                    FlushQuote();
                    var environment = codeStart.Groups["name"].Value;
                    inCodeFence = true;
                    codeFenceEnd = $"\\end{{{environment}}}";
                    codeFenceDescription = $"LaTeX {environment} 环境";
                    continue;
                }
            }
            if (inCodeFence)
            {
                if (trimmed.Equals(codeFenceEnd, StringComparison.OrdinalIgnoreCase))
                    FinishCodeBlock();
                else
                    code.AppendLine(raw);
                continue;
            }

            if (TryReadEmbeddedDisplayFormula(
                    lines,
                    ref index,
                    format,
                    out var prefixText,
                    out var displayLatex,
                    out var suffixText,
                    out var warning))
            {
                if (!string.IsNullOrWhiteSpace(prefixText))
                {
                    if (inLatexQuote)
                        quote.Add(prefixText.Trim());
                    else
                        paragraph.Add(prefixText.Trim());
                }
                FlushParagraph();
                FlushQuote();
                if (!string.IsNullOrWhiteSpace(warning)) warnings.Add(warning!);
                if (!string.IsNullOrWhiteSpace(displayLatex))
                {
                    var split = FormulaEquationTag.Extract(displayLatex);
                    blocks.Add(new WordBulkBlock
                    {
                        Kind = WordBulkBlockKind.DisplayFormula,
                        Runs = new List<WordBulkRun>
                        {
                            new()
                            {
                                IsFormula = true,
                                Latex = split.Latex,
                                EquationTag = split.EquationTag,
                                DisplayMode = "block",
                            },
                        },
                    });
                }
                if (!string.IsNullOrWhiteSpace(suffixText))
                    lines.Insert(index + 1, suffixText);
                continue;
            }

            if (TryReadDisplayFormula(lines, ref index, format, out displayLatex, out warning))
            {
                FlushParagraph();
                FlushQuote();
                if (!string.IsNullOrWhiteSpace(warning)) warnings.Add(warning!);
                if (!string.IsNullOrWhiteSpace(displayLatex))
                {
                    var split = FormulaEquationTag.Extract(displayLatex);
                    blocks.Add(new WordBulkBlock
                    {
                        Kind = WordBulkBlockKind.DisplayFormula,
                        Runs = new List<WordBulkRun>
                        {
                            new()
                            {
                                IsFormula = true,
                                Latex = split.Latex,
                                EquationTag = split.EquationTag,
                                DisplayMode = "block",
                            },
                        },
                    });
                }
                continue;
            }

            if (format == WordBulkSourceFormat.Latex)
            {
                var theoremStart = LatexTheoremEnvironmentStart.Match(trimmed);
                if (theoremStart.Success)
                {
                    FlushParagraph();
                    FlushQuote();
                    var label = LatexTheoremLabel(theoremStart.Groups["name"].Value);
                    var title = theoremStart.Groups["title"].Value.Trim();
                    blocks.Add(new WordBulkBlock
                    {
                        Kind = WordBulkBlockKind.Paragraph,
                        Runs = new List<WordBulkRun>
                        {
                            new()
                            {
                                Text = title.Length > 0
                                    ? $"{label}（{title}）："
                                    : $"{label}：",
                                Bold = true,
                            },
                        },
                    });
                    continue;
                }
                if (LatexTheoremEnvironmentEnd.IsMatch(trimmed))
                {
                    FlushParagraph();
                    FlushQuote();
                    continue;
                }

                if (Regex.IsMatch(
                        trimmed,
                        @"^\\begin\{(?:quote|quotation)\}\s*$",
                        RegexOptions.IgnoreCase))
                {
                    FlushParagraph();
                    FlushQuote();
                    inLatexQuote = true;
                    continue;
                }
                if (Regex.IsMatch(
                        trimmed,
                        @"^\\end\{(?:quote|quotation)\}\s*$",
                        RegexOptions.IgnoreCase))
                {
                    FlushParagraph();
                    FlushQuote();
                    inLatexQuote = false;
                    continue;
                }

                var listStart = Regex.Match(
                    trimmed,
                    @"^\\begin\{(?<name>itemize|enumerate)\}\s*$",
                    RegexOptions.IgnoreCase);
                if (listStart.Success)
                {
                    FlushParagraph();
                    FlushQuote();
                    listModes.Push(
                        listStart.Groups["name"].Value.Equals(
                            "enumerate",
                            StringComparison.OrdinalIgnoreCase)
                            ? "numbered"
                            : "bullet");
                    continue;
                }
                var listEnd = Regex.Match(
                    trimmed,
                    @"^\\end\{(?<name>itemize|enumerate)\}\s*$",
                    RegexOptions.IgnoreCase);
                if (listEnd.Success)
                {
                    FlushParagraph();
                    FlushQuote();
                    if (listModes.Count == 0)
                    {
                        warnings.Add($"忽略了没有对应开始标记的 {trimmed}。");
                    }
                    else
                    {
                        listModes.Pop();
                    }
                    continue;
                }
                var section = LatexSection.Match(trimmed);
                if (section.Success)
                {
                    FlushParagraph();
                    FlushQuote();
                    var level = section.Groups["kind"].Value.ToLowerInvariant() switch
                    {
                        "part" => 1,
                        "chapter" => 1,
                        "section" => 1,
                        "subsection" => 2,
                        "subsubsection" => 3,
                        "paragraph" => 4,
                        _ => 5,
                    };
                    blocks.Add(new WordBulkBlock
                    {
                        Kind = WordBulkBlockKind.Heading,
                        Level = level,
                        Runs = ParseInlineRuns(section.Groups["text"].Value, format, warnings),
                    });
                    continue;
                }
                var item = LatexItem.Match(trimmed);
                if (item.Success)
                {
                    FlushParagraph();
                    FlushQuote();
                    if (listModes.Count == 0)
                        warnings.Add("检测到列表外的 \\item，已按一级项目符号导入。");
                    blocks.Add(new WordBulkBlock
                    {
                        Kind = listModes.Count > 0 && listModes.Peek() == "numbered"
                            ? WordBulkBlockKind.Numbered
                            : WordBulkBlockKind.Bullet,
                        Level = Math.Max(0, listModes.Count - 1),
                        Runs = ParseInlineRuns(item.Groups["text"].Value, format, warnings),
                    });
                    continue;
                }
            }
            else
            {
                if (trimmed.StartsWith(">", StringComparison.Ordinal))
                {
                    FlushParagraph();
                    quote.Add(trimmed.TrimStart('>').TrimStart());
                    continue;
                }
                FlushQuote();

                var heading = MarkdownHeading.Match(raw);
                if (heading.Success)
                {
                    FlushParagraph();
                    blocks.Add(new WordBulkBlock
                    {
                        Kind = WordBulkBlockKind.Heading,
                        Level = heading.Groups["marks"].Value.Length,
                        Runs = ParseInlineRuns(heading.Groups["text"].Value, format, warnings),
                    });
                    continue;
                }
                var bullet = MarkdownBullet.Match(raw);
                if (bullet.Success)
                {
                    FlushParagraph();
                    blocks.Add(new WordBulkBlock
                    {
                        Kind = WordBulkBlockKind.Bullet,
                        Level = MarkdownListLevel(bullet.Groups["indent"].Value),
                        Runs = ParseInlineRuns(bullet.Groups["text"].Value, format, warnings),
                    });
                    continue;
                }
                var numbered = MarkdownNumbered.Match(raw);
                if (numbered.Success)
                {
                    FlushParagraph();
                    blocks.Add(new WordBulkBlock
                    {
                        Kind = WordBulkBlockKind.Numbered,
                        Level = MarkdownListLevel(numbered.Groups["indent"].Value),
                        Runs = ParseInlineRuns(numbered.Groups["text"].Value, format, warnings),
                    });
                    continue;
                }
            }

            if (trimmed.Length == 0)
            {
                FlushParagraph();
                FlushQuote();
                continue;
            }
            if (inLatexQuote)
                quote.Add(trimmed);
            else
                paragraph.Add(trimmed);
        }

        if (inCodeFence)
            FinishCodeBlock($"{codeFenceDescription}未闭合，已导入到文末。");
        FlushParagraph();
        FlushQuote();
        if (inLatexQuote)
            warnings.Add("LaTeX quote/quotation 环境未闭合，已导入到文末。");
        if (listModes.Count > 0)
            warnings.Add($"LaTeX 文档有 {listModes.Count} 个列表环境未闭合，已导入其余内容。");
        return blocks;
    }

    private static string LatexTheoremLabel(string environment) =>
        environment.Trim().ToLowerInvariant() switch
        {
            "theorem" => "定理",
            "lemma" => "引理",
            "proposition" => "命题",
            "corollary" => "推论",
            "definition" => "定义",
            "proof" => "证明",
            "remark" => "注记",
            "example" => "例",
            "exercise" => "练习",
            "assumption" => "假设",
            "axiom" => "公理",
            "claim" => "断言",
            "conjecture" => "猜想",
            "criterion" => "判据",
            "fact" => "事实",
            "notation" => "记号",
            "observation" => "观察",
            "problem" => "问题",
            "question" => "问题",
            "solution" => "解",
            _ => environment,
        };

    private static int MarkdownListLevel(string indentation)
    {
        if (string.IsNullOrEmpty(indentation)) return 0;
        var columns = 0;
        foreach (var character in indentation)
            columns += character == '\t' ? 4 : 1;
        return Math.Min(8, Math.Max(0, columns / 2));
    }

    private static bool TryReadEmbeddedDisplayFormula(
        IList<string> lines,
        ref int index,
        WordBulkSourceFormat format,
        out string prefix,
        out string latex,
        out string suffix,
        out string? warning)
    {
        prefix = string.Empty;
        latex = string.Empty;
        suffix = string.Empty;
        warning = null;
        if (index < 0 || index >= lines.Count) return false;

        var raw = lines[index];
        var starts = new List<(int Position, string StartToken, string EndToken, string? Environment)>();
        var dollars = FindUnescapedSequence(raw, "$$", 0);
        if (dollars >= 0)
            starts.Add((dollars, "$$", "$$", null));
        var bracket = FindUnescapedSequence(raw, "\\[", 0);
        if (bracket >= 0)
            starts.Add((bracket, "\\[", "\\]", null));
        if (format == WordBulkSourceFormat.Latex)
        {
            var environment = Regex.Match(
                raw,
                @"\\begin\{(?<name>equation\*?|align\*?|gather\*?|multline\*?|displaymath)\}",
                RegexOptions.IgnoreCase);
            if (environment.Success && !IsEscaped(raw, environment.Index))
            {
                var name = environment.Groups["name"].Value;
                starts.Add((
                    environment.Index,
                    environment.Value,
                    $"\\end{{{name}}}",
                    name));
            }
        }
        if (starts.Count == 0) return false;

        var start = starts.OrderBy(candidate => candidate.Position).First();
        prefix = raw.Substring(0, start.Position).TrimEnd();
        var builder = new StringBuilder();
        var firstContentStart = start.Position + start.StartToken.Length;
        for (var cursor = index; cursor < lines.Count; cursor++)
        {
            var line = lines[cursor];
            var searchStart = cursor == index ? firstContentStart : 0;
            var end = FindUnescapedSequence(line, start.EndToken, searchStart);
            if (end >= 0)
            {
                if (end > searchStart)
                    builder.Append(line.Substring(searchStart, end - searchStart));
                suffix = line.Substring(end + start.EndToken.Length).TrimStart();
                index = cursor;
                latex = start.Environment is null
                    ? NormalizeDelimitedDisplayLatex(builder.ToString())
                    : NormalizeDisplayEnvironmentLatex(
                        start.Environment,
                        builder.ToString());
                return true;
            }
            if (searchStart < line.Length)
                builder.Append(line.Substring(searchStart));
            if (cursor + 1 < lines.Count) builder.Append('\n');
        }

        index = lines.Count - 1;
        latex = start.Environment is null
            ? NormalizeDelimitedDisplayLatex(builder.ToString())
            : NormalizeDisplayEnvironmentLatex(start.Environment, builder.ToString());
        warning = start.Environment is null
            ? $"行间公式缺少结束标记 {start.EndToken}，已导入到文末。"
            : $"LaTeX 环境 {start.Environment} 未闭合，已导入到文末。";
        return true;
    }

    private static bool TryReadDisplayFormula(
        IReadOnlyList<string> lines,
        ref int index,
        WordBulkSourceFormat format,
        out string latex,
        out string? warning)
    {
        latex = string.Empty;
        warning = null;
        var trimmed = lines[index].Trim();
        if (trimmed.StartsWith("$$", StringComparison.Ordinal))
        {
            return ReadDelimited(lines, ref index, "$$", "$$", out latex, out warning);
        }
        if (trimmed.StartsWith("\\[", StringComparison.Ordinal))
        {
            return ReadDelimited(lines, ref index, "\\[", "\\]", out latex, out warning);
        }
        if (format != WordBulkSourceFormat.Latex) return false;
        var match = LatexEnvironmentStart.Match(trimmed);
        if (!match.Success) return false;
        var environment = match.Groups["name"].Value;
        var endToken = $"\\end{{{environment}}}";
        var builder = new StringBuilder();
        for (var cursor = index + 1; cursor < lines.Count; cursor++)
        {
            var line = lines[cursor];
            if (line.Trim().Equals(endToken, StringComparison.OrdinalIgnoreCase))
            {
                index = cursor;
                latex = NormalizeDisplayEnvironmentLatex(
                    environment,
                    builder.ToString());
                return true;
            }
            builder.Append(line).Append('\n');
        }
        index = lines.Count - 1;
        latex = NormalizeDisplayEnvironmentLatex(
            environment,
            builder.ToString());
        warning = $"LaTeX 环境 {environment} 未闭合，已导入到文末。";
        return true;
    }

    private static string NormalizeDisplayEnvironmentLatex(
        string environment,
        string body)
    {
        // Physical source line breaks inside TeX math are ordinary whitespace.
        // Preserve explicit row commands such as `\\\\`, but do not reinterpret
        // pretty-printed equation/equation* source as VisualTeX formula rows.
        var formulaBody = Regex.Replace(
            Regex.Replace(
                body,
                @"\\label\s*\{[^{}]*\}",
                string.Empty,
                RegexOptions.IgnoreCase),
            @"\\(?:notag|nonumber)\b",
            string.Empty,
            RegexOptions.IgnoreCase);
        var normalizedBody = Regex.Replace(
                formulaBody
                    .Replace("\r\n", "\n")
                    .Replace('\r', '\n'),
                @"[ \t]*\n+[ \t]*",
                " ")
            .Trim();
        var baseEnvironment = environment.TrimEnd('*').ToLowerInvariant();
        return baseEnvironment switch
        {
            // MathJax receives a formula body that is already in display math
            // mode. Convert document-level AMS structures to their embeddable
            // counterparts so alignment markers and row breaks remain valid.
            "align" => $"\\begin{{aligned}}{normalizedBody}\\end{{aligned}}",
            "gather" => $"\\begin{{gathered}}{normalizedBody}\\end{{gathered}}",
            // MathJax/MathLive do not consistently support a document-level
            // multline environment inside an existing display formula. Keep
            // its row structure in a stable, editable gathered environment.
            "multline" => $"\\begin{{gathered}}{normalizedBody}\\end{{gathered}}",
            _ => normalizedBody,
        };
    }

    private static string NormalizeDelimitedDisplayLatex(string body)
    {
        var normalized = Regex.Replace(
                Regex.Replace(
                    body,
                    @"\\label\s*\{[^{}]*\}",
                    string.Empty,
                    RegexOptions.IgnoreCase),
                @"\\(?:notag|nonumber)\b",
                string.Empty,
                RegexOptions.IgnoreCase)
            .Replace("\r\n", "\n")
            .Replace('\r', '\n');
        // Newlines inside $$...$$ and \[...\] are ordinary TeX whitespace.
        // Do not reinterpret source formatting as VisualTeX formula rows:
        // doing so can split paired \left/\right delimiters across aligned rows.
        return Regex.Replace(normalized, @"[ \t]*\n+[ \t]*", " ").Trim();
    }

    private static bool ReadDelimited(
        IReadOnlyList<string> lines,
        ref int index,
        string startToken,
        string endToken,
        out string latex,
        out string? warning)
    {
        var first = lines[index].Trim();
        var afterStart = first.Substring(startToken.Length);
        var sameLineEnd = afterStart.IndexOf(endToken, StringComparison.Ordinal);
        if (sameLineEnd >= 0)
        {
            latex = NormalizeDelimitedDisplayLatex(
                afterStart.Substring(0, sameLineEnd));
            warning = null;
            return true;
        }
        var builder = new StringBuilder();
        if (afterStart.Length > 0) builder.Append(afterStart).Append('\n');
        for (var cursor = index + 1; cursor < lines.Count; cursor++)
        {
            var line = lines[cursor];
            var end = line.IndexOf(endToken, StringComparison.Ordinal);
            if (end >= 0)
            {
                builder.Append(line.Substring(0, end));
                index = cursor;
                latex = NormalizeDelimitedDisplayLatex(builder.ToString());
                warning = null;
                return true;
            }
            builder.Append(line).Append('\n');
        }
        index = lines.Count - 1;
        latex = NormalizeDelimitedDisplayLatex(builder.ToString());
        warning = $"行间公式缺少结束标记 {endToken}，已导入到文末。";
        return true;
    }

    private static List<WordBulkRun> ParseInlineRuns(
        string text,
        WordBulkSourceFormat format,
        ICollection<string> warnings)
    {
        var runs = new List<WordBulkRun>();
        ParseInlineSegment(text, format, false, false, false, runs, warnings);
        var merged = MergeTextRuns(runs);
        if (format == WordBulkSourceFormat.Latex)
            merged = NormalizeLatexInlineBoundaryWhitespace(merged);
        if (merged.Count == 0) merged.Add(new WordBulkRun { Text = string.Empty });
        return merged;
    }

    private static void ParseInlineSegment(
        string text,
        WordBulkSourceFormat format,
        bool bold,
        bool italic,
        bool code,
        ICollection<WordBulkRun> runs,
        ICollection<string> warnings)
    {
        var buffer = new StringBuilder();
        void Flush()
        {
            if (buffer.Length == 0) return;
            runs.Add(new WordBulkRun
            {
                Text = DecodeText(buffer.ToString(), format),
                Bold = bold,
                Italic = italic,
                Code = code,
            });
            buffer.Clear();
        }

        for (var index = 0; index < text.Length;)
        {
            if (text[index] == '$' && !IsEscaped(text, index))
            {
                var end = FindUnescaped(text, '$', index + 1);
                if (end > index + 1)
                {
                    Flush();
                    runs.Add(new WordBulkRun
                    {
                        IsFormula = true,
                        Latex = text.Substring(index + 1, end - index - 1).Trim(),
                        DisplayMode = "inline",
                    });
                    index = end + 1;
                    continue;
                }
            }
            if (index + 1 < text.Length
                && text[index] == '\\'
                && text[index + 1] == '(')
            {
                var end = text.IndexOf("\\)", index + 2, StringComparison.Ordinal);
                if (end > index + 2)
                {
                    Flush();
                    runs.Add(new WordBulkRun
                    {
                        IsFormula = true,
                        Latex = text.Substring(index + 2, end - index - 2).Trim(),
                        DisplayMode = "inline",
                    });
                    index = end + 2;
                    continue;
                }
            }

            if (format == WordBulkSourceFormat.Markdown)
            {
                var tripleDelimiter = text.IndexOf("***", index, StringComparison.Ordinal) == index
                    ? "***"
                    : text.IndexOf("___", index, StringComparison.Ordinal) == index
                        ? "___"
                        : null;
                if (tripleDelimiter is not null && !IsEscaped(text, index))
                {
                    var end = FindMarkdownClosingDelimiter(
                        text,
                        tripleDelimiter,
                        index + tripleDelimiter.Length);
                    if (end > index + tripleDelimiter.Length)
                    {
                        Flush();
                        ParseInlineSegment(
                            text.Substring(
                                index + tripleDelimiter.Length,
                                end - index - tripleDelimiter.Length),
                            format,
                            true,
                            true,
                            code,
                            runs,
                            warnings);
                        index = end + tripleDelimiter.Length;
                        continue;
                    }
                }
                var strongDelimiter = text.IndexOf("**", index, StringComparison.Ordinal) == index
                    ? "**"
                    : text.IndexOf("__", index, StringComparison.Ordinal) == index
                        ? "__"
                        : null;
                if (strongDelimiter is not null && !IsEscaped(text, index))
                {
                    var end = FindMarkdownClosingDelimiter(
                        text,
                        strongDelimiter,
                        index + strongDelimiter.Length);
                    if (end > index + strongDelimiter.Length)
                    {
                        Flush();
                        ParseInlineSegment(
                            text.Substring(
                                index + strongDelimiter.Length,
                                end - index - strongDelimiter.Length),
                            format,
                            true,
                            italic,
                            code,
                            runs,
                            warnings);
                        index = end + strongDelimiter.Length;
                        continue;
                    }
                }
                if ((text[index] == '*' || text[index] == '_') && !IsEscaped(text, index))
                {
                    var marker = text[index];
                    var end = FindUnescaped(text, marker, index + 1);
                    if (end > index + 1)
                    {
                        Flush();
                        ParseInlineSegment(
                            text.Substring(index + 1, end - index - 1),
                            format,
                            bold,
                            true,
                            code,
                            runs,
                            warnings);
                        index = end + 1;
                        continue;
                    }
                }
                if (text[index] == '`')
                {
                    var end = text.IndexOf('`', index + 1);
                    if (end > index + 1)
                    {
                        Flush();
                        runs.Add(new WordBulkRun
                        {
                            Text = text.Substring(index + 1, end - index - 1),
                            Bold = bold,
                            Italic = italic,
                            Code = true,
                        });
                        index = end + 1;
                        continue;
                    }
                }
            }
            else if (text[index] == '\\')
            {
                if (TryReadLatexInlineLiteral(text, index, out var literal, out var literalEnd))
                {
                    Flush();
                    runs.Add(new WordBulkRun
                    {
                        Text = literal,
                        Bold = bold,
                        Italic = italic,
                        Code = true,
                    });
                    index = literalEnd;
                    goto ContinueOuter;
                }
                var explicitSpacing = Regex.Match(
                    text.Substring(index),
                    @"^\\(?<name>quad|qquad)(?![A-Za-z@])",
                    RegexOptions.IgnoreCase);
                if (explicitSpacing.Success)
                {
                    Flush();
                    runs.Add(new WordBulkRun
                    {
                        Text = explicitSpacing.Groups["name"].Value.Equals(
                            "qquad",
                            StringComparison.OrdinalIgnoreCase)
                            ? "\u00A0\u00A0"
                            : "\u00A0",
                        Bold = bold,
                        Italic = italic,
                        Code = code,
                    });
                    index += explicitSpacing.Length;
                    goto ContinueOuter;
                }
                var declaration = Regex.Match(
                    text.Substring(index),
                    @"^\\(?:tiny|scriptsize|footnotesize|small|normalsize|large|huge|centering|raggedright|raggedleft)(?![A-Za-z@])",
                    RegexOptions.IgnoreCase);
                if (declaration.Success)
                {
                    Flush();
                    index += declaration.Length;
                    goto ContinueOuter;
                }
                foreach (var command in new[]
                         {
                             (Name: "\\textbf{", Bold: true, Italic: italic, Code: code),
                             (Name: "\\textit{", Bold: bold, Italic: true, Code: code),
                             (Name: "\\emph{", Bold: bold, Italic: true, Code: code),
                             (Name: "\\texttt{", Bold: bold, Italic: italic, Code: true),
                         })
                {
                    if (text.IndexOf(command.Name, index, StringComparison.Ordinal) != index)
                        continue;
                    var open = index + command.Name.Length - 1;
                    var close = FindMatchingBrace(text, open);
                    if (close > open)
                    {
                        Flush();
                        ParseInlineSegment(
                            text.Substring(open + 1, close - open - 1),
                            format,
                            command.Bold,
                            command.Italic,
                            command.Code,
                            runs,
                            warnings);
                        index = close + 1;
                        goto ContinueOuter;
                    }
                }
            }

            buffer.Append(text[index]);
            index++;
            ContinueOuter:;
        }
        Flush();
    }

    private static bool IsTightLatexBoundaryCharacter(char character)
    {
        return character is >= '\u2E80' and <= '\u2FFF'
            or >= '\u3000' and <= '\u303F'
            or >= '\u3040' and <= '\u30FF'
            or >= '\u31C0' and <= '\u31EF'
            or >= '\u3400' and <= '\u4DBF'
            or >= '\u4E00' and <= '\u9FFF'
            or >= '\uAC00' and <= '\uD7AF'
            or >= '\uF900' and <= '\uFAFF'
            or >= '\uFF00' and <= '\uFFEF'
            or '，' or '。' or '！' or '？' or '；' or '：' or '、'
            or '》' or '）' or '】' or '」' or '』' or '”' or '’' or '…'
            or ',' or '.' or '!' or '?' or ';' or ':'
            or '(' or ')' or '[' or ']' or '{' or '}';
    }

    private static List<WordBulkRun> NormalizeLatexInlineBoundaryWhitespace(
        List<WordBulkRun> runs)
    {
        for (var index = 0; index < runs.Count; index++)
        {
            var run = runs[index];
            if (!run.IsFormula || run.DisplayMode != "inline") continue;

            if (index > 0 && !runs[index - 1].IsFormula)
            {
                var previous = runs[index - 1];
                var visible = previous.Text.TrimEnd(' ', '\t');
                if (visible.Length < previous.Text.Length
                    && visible.Length > 0
                    && IsTightLatexBoundaryCharacter(visible[visible.Length - 1]))
                    previous.Text = visible;
            }

            if (index + 1 < runs.Count && !runs[index + 1].IsFormula)
            {
                var next = runs[index + 1];
                var visible = next.Text.TrimStart(' ', '\t');
                if (visible.Length < next.Text.Length
                    && visible.Length > 0
                    && IsTightLatexBoundaryCharacter(visible[0]))
                    next.Text = visible;
            }
        }
        runs.RemoveAll(run => !run.IsFormula && run.Text.Length == 0);
        return runs;
    }

    private static List<WordBulkRun> MergeTextRuns(IEnumerable<WordBulkRun> source)
    {
        var merged = new List<WordBulkRun>();
        foreach (var run in source)
        {
            var previous = merged.LastOrDefault();
            if (!run.IsFormula
                && previous is not null
                && !previous.IsFormula
                && previous.Bold == run.Bold
                && previous.Italic == run.Italic
                && previous.Code == run.Code
                && previous.Strike == run.Strike
                && previous.Underline == run.Underline)
            {
                merged[merged.Count - 1] = new WordBulkRun
                {
                    Id = previous.Id,
                    Text = previous.Text + run.Text,
                    Bold = previous.Bold,
                    Italic = previous.Italic,
                    Code = previous.Code,
                    Strike = previous.Strike,
                    Underline = previous.Underline,
                };
            }
            else
            {
                merged.Add(run);
            }
        }
        return merged;
    }

    private static string DecodeText(string value, WordBulkSourceFormat format)
    {
        if (format == WordBulkSourceFormat.Markdown)
        {
            return Regex.Replace(value, @"\\([\\`*_{}\[\]()#+\-.!$])", "$1");
        }
        return value
            .Replace("~", "\u00A0")
            .Replace("\\ ", "\u00A0")
            .Replace("\\%", "%")
            .Replace("\\_", "_")
            .Replace("\\&", "&")
            .Replace("\\#", "#")
            .Replace("\\$", "$")
            .Replace("\\{", "{")
            .Replace("\\}", "}")
            .Replace("\\textbackslash{}", "\\")
            .Replace("\\newline", "\v")
            .Replace("\\\\", "\v");
    }

    private static int FindMatchingBrace(string text, int open)
    {
        var depth = 0;
        for (var index = open; index < text.Length; index++)
        {
            if (text[index] == '{' && !IsEscaped(text, index)) depth++;
            if (text[index] == '}' && !IsEscaped(text, index))
            {
                depth--;
                if (depth == 0) return index;
            }
        }
        return -1;
    }

    private static int FindMarkdownClosingDelimiter(
        string text,
        string delimiter,
        int start)
    {
        var cursor = Math.Max(0, start);
        while (cursor <= text.Length - delimiter.Length)
        {
            var found = text.IndexOf(delimiter, cursor, StringComparison.Ordinal);
            if (found < 0) return -1;
            if (!IsEscaped(text, found))
            {
                if (delimiter.Length == 2
                    && found + delimiter.Length < text.Length
                    && text[found + delimiter.Length] == delimiter[0])
                {
                    return found + 1;
                }
                return found;
            }
            cursor = found + delimiter.Length;
        }
        return -1;
    }

    private static int FindUnescaped(string text, char target, int start)
    {
        for (var index = start; index < text.Length; index++)
        {
            if (text[index] == target && !IsEscaped(text, index)) return index;
        }
        return -1;
    }

    private static int FindUnescapedSequence(string text, string target, int start)
    {
        if (string.IsNullOrEmpty(target)) return -1;
        for (var index = Math.Max(0, start); index <= text.Length - target.Length; index++)
        {
            if (!text.AsSpan(index, target.Length).SequenceEqual(target.AsSpan())) continue;
            if (!IsEscaped(text, index)) return index;
        }
        return -1;
    }

    private static bool IsEscaped(string text, int index)
    {
        var slashes = 0;
        for (var cursor = index - 1; cursor >= 0 && text[cursor] == '\\'; cursor--) slashes++;
        return slashes % 2 == 1;
    }
}
