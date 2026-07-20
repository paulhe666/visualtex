using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;

namespace VisualTeX.WindowsOffice.Contracts;

public interface IWindowsOfficeBackend
{
    Task<OfficeBridgeResponse> HandleAsync(
        OfficeBridgeRequest request,
        CancellationToken cancellationToken);

    IReadOnlyList<OfficeBridgeEvent> GetEventsAfter(long cursor);
}
