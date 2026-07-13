namespace VisualTeX.WindowsOffice.Contracts;

public interface IWordFormulaService
{
    OfficeSelection GetSelection();
    OfficeObjectResult InsertInlineFormula(SessionInfo session);
    OfficeObjectResult InsertDisplayFormula(SessionInfo session);
    OfficeObjectResult ReplaceFormula(SessionInfo session);
    int UpdateEquationNumbers();
}

public sealed class OfficeSelection
{
    public string Host { get; set; } = string.Empty;
    public string? DocumentId { get; set; }
    public string? ObjectId { get; set; }
    public bool ReadOnly { get; set; }
    public string? FormulaId { get; set; }
    public FormulaMetadata? Metadata { get; set; }
}

public sealed class OfficeObjectResult
{
    public string FormulaId { get; set; } = string.Empty;
    public string? DocumentId { get; set; }
    public string? ObjectId { get; set; }
}
