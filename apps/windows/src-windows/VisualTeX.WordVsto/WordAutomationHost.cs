using System.Text.Json;
using VisualTeX.WindowsOffice.VstoShared;
using Application = Microsoft.Office.Interop.Word.Application;

namespace VisualTeX.WordVsto;

/// <summary>Word and WPS Writer operations for agents (docs/office-mcp/DESIGN.md §5).</summary>
internal sealed class WordAutomationHost : IOfficeAutomationHost
{
    private readonly Application _application;
    private readonly WordFormulaService _service;

    internal WordAutomationHost(Application application, WordFormulaService service)
    {
        _application = application;
        _service = service;
    }

    public string App => "word";

    public string ProductVersion
    {
        get
        {
            try
            {
                var version = _application.Version;
                var build = _application.Build;
                return build.StartsWith(version, StringComparison.Ordinal) ? build : $"{version}.{build}";
            }
            catch { return string.Empty; }
        }
    }

    public IReadOnlyList<string> Capabilities { get; } = new[] { "read" };

    public object? Execute(string operation, JsonElement arguments) => operation switch
    {
        "list_documents" => _service.AutomationListDocuments(),
        "get_selection" => _service.AutomationGetSelection(OfficeAutomationText.String(arguments, "document")),
        "list_formulas" => _service.AutomationListFormulas(
            OfficeAutomationText.RequiredString(arguments, "document"),
            OfficeAutomationText.Integer(arguments, "offset", 0, 0, int.MaxValue),
            OfficeAutomationText.Integer(arguments, "limit", 50, 1, 200),
            OfficeAutomationText.Integer(arguments, "contextChars", 40, 0, 120),
            OfficeAutomationText.Kinds(arguments)),
        "read_formula" => _service.AutomationReadFormula(
            OfficeAutomationText.RequiredString(arguments, "document"),
            OfficeAutomationText.RequiredString(arguments, "formulaId")),
        _ => throw OfficeAutomationText.UnknownOperation(operation),
    };
}
