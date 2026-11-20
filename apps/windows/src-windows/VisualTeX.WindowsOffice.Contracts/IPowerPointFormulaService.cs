namespace VisualTeX.WindowsOffice.Contracts;

public interface IPowerPointFormulaService
{
    OfficeSelection GetSelection();
    OfficeObjectResult InsertFormula(SessionInfo session);
    OfficeObjectResult ReplaceFormula(SessionInfo session);
    OfficeObjectResult MarkFormula(string formulaId);
    void DeleteFormula(string formulaId);
}
