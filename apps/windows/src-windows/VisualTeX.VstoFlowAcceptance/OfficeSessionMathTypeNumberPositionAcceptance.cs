using VisualTeX.WindowsOffice.Contracts;

namespace VisualTeX.VstoFlowAcceptance;

internal static partial class Program
{
    private static void RunOfficeSessionMathTypeNumberPositionAcceptance(
        VisualTeXSessionClient client)
    {
        var lineId = Guid.NewGuid().ToString("D");
        var request = new CreateVstoSessionRequest
        {
            Mode = "create",
            Host = "word",
            Title = "MathType numbering position Session acceptance",
            Lines = new List<FormulaLine>
            {
                new() { Id = lineId, Latex = "x=1" },
            },
            ActiveLineId = lineId,
            CodeFormat = "latex",
            DisplayMode = "block",
            ObjectMode = FormulaOleContract.MathTypeOleMode,
            Numbered = true,
            MathTypeNumberPosition = "right",
            FontSizePt = 10.5,
            AutoCommitOnClose = true,
        };

        var session = client.CreateSessionAsync(request, CancellationToken.None)
            .GetAwaiter()
            .GetResult();
        if (!string.Equals(session.MathTypeNumberPosition, "right", StringComparison.Ordinal))
            throw new InvalidDataException(
                "New MathType numbered Session did not preserve the initial right number position.");

        var autosaved = client.PatchAsync(
                session.Id,
                new
                {
                    mathTypeNumberPosition = "left",
                    status = "editing",
                },
                CancellationToken.None)
            .GetAwaiter()
            .GetResult();
        if (!string.Equals(autosaved.MathTypeNumberPosition, "left", StringComparison.Ordinal))
            throw new InvalidDataException(
                "MathType number-position autosave did not persist left.");

        var committing = client.PatchAsync(
                session.Id,
                new
                {
                    mathTypeNumberPosition = "left",
                    status = "committing",
                    dirty = true,
                    exportResult = new
                    {
                        svg = "<svg xmlns=\"http://www.w3.org/2000/svg\" viewBox=\"0 0 10 10\"></svg>",
                        svgBase64 = "PHN2ZyB4bWxucz0iaHR0cDovL3d3dy53My5vcmcvMjAwMC9zdmciIHZpZXdCb3g9IjAgMCAxMCAxMCI+PC9zdmc+",
                        pngBase64 = (string?)null,
                        width = 10.0,
                        height = 10.0,
                        baseline = 8.0,
                    },
                },
                CancellationToken.None)
            .GetAwaiter()
            .GetResult();
        if (!string.Equals(committing.MathTypeNumberPosition, "left", StringComparison.Ordinal)
            || !string.Equals(committing.Status, "committing", StringComparison.Ordinal))
            throw new InvalidDataException(
                "MathType number-position commit did not preserve the autosaved left value.");

        client.PatchAsync(
                session.Id,
                new
                {
                    status = "cancelled",
                    explicitCancel = true,
                    error = (string?)null,
                },
                CancellationToken.None)
            .GetAwaiter()
            .GetResult();

        Console.WriteLine(
            "[Session MathType number position] create right -> autosave left -> committing left passed through the real Companion HTTP API without HTTP 400.");
    }
}
