namespace VisualTeX.WordVsto;

/// <summary>
/// The \x01CompObj stream of a MathType Equation.DSMT4 storage, byte for byte as
/// MathType writes it. WriteFmtUserTypeStg takes the ProgID from the registry and
/// leaves it empty when no MathType is registered; Word then reports a freshly
/// inserted object with an empty OLEFormat.ProgID for the rest of the session
/// while its XML still says Equation.DSMT4.
/// </summary>
internal static class MathTypeCompObj
{
    internal const string ProgId = "Equation.DSMT4";
    internal const string ClipboardFormat = "MathType EF";

    internal static byte[] Build(Guid clsid, string userType, string? progId)
    {
        using var data = new MemoryStream();
        using var writer = new BinaryWriter(data);
        void WriteAnsi(string value)
        {
            var bytes = System.Text.Encoding.ASCII.GetBytes(value + "\0");
            writer.Write(bytes.Length);
            writer.Write(bytes);
        }
        writer.Write(0xFFFE0001u);
        writer.Write(0x00000A03u);
        writer.Write(0xFFFFFFFFu);
        writer.Write(clsid.ToByteArray());
        WriteAnsi(userType);
        WriteAnsi(ClipboardFormat);
        WriteAnsi(string.IsNullOrWhiteSpace(progId) ? ProgId : progId!);
        writer.Write(0x71B239F4u);
        writer.Write(0);
        writer.Write(0);
        writer.Write(0);
        writer.Flush();
        return data.ToArray();
    }
}
