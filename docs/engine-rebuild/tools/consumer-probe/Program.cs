using System.Reflection;
using System.Security.Cryptography;
using System.Text.Json;
using System.Runtime.InteropServices;
using VisualTeX.WindowsOffice.Contracts;

// Calls only pure contract methods in a byte-identical copy of the INSTALLED assembly.
// No Office COM activation, registration, installation, or production rebuild.
var assembly = typeof(OfficeFormulaSizing).Assembly;
var hash = Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(assembly.Location)));
const string expected = "49EDA828D2B5928C4E53F8E91943B76CFB866298900BC13B8B3A610B14F0B073";
if (hash != expected) throw new InvalidOperationException("Installed contract assembly changed; inspect baseline before continuing.");
int checks = 0;
void Equal(double actual, double wanted) {
    if (Math.Abs(actual - wanted) > 1e-6) throw new InvalidOperationException($"Expected {wanted}, received {actual}");
    checks++;
}
var natural = OfficeFormulaSizing.NaturalSize(96, 48);
Equal(natural.Width, 72); Equal(natural.Height, 36);
var tiny = OfficeFormulaSizing.NaturalSize(1, 1);
Equal(tiny.Width, 1); Equal(tiny.Height, 1);
var edited = OfficeFormulaSizing.EditedSize(18.75f, 8.25f, 25, 11, 50, 11);
Equal(edited.Width, 37.5); Equal(edited.Height, 8.25);
Equal(FormulaFontSize.Normalize(13.25), 13.25);
Equal(FormulaFontSize.Normalize(null), 14);
Equal(FormulaOleContract.ProtocolVersion, 2);
Equal(FormulaOleContract.StorageSchemaVersion, 1);
var abi = typeof(IVisualTeXFormulaObject).GetMethods().OrderBy(m => m.Name)
    .Select(m => new { name = m.Name, dispId = m.GetCustomAttribute<DispIdAttribute>()?.Value }).ToArray();
foreach (var method in abi) {
    int wanted = method.name switch { "InitializeFromFiles" => 1, "UpdateFromFiles" => 2, "GetFormulaJson" => 3, _ => throw new Exception("Unexpected ABI method") };
    Equal(method.dispId ?? -1, wanted);
}
if (typeof(IVisualTeXFormulaObject).GUID.ToString().ToUpperInvariant() != "6C672AF0-7321-4D21-B325-868CB34592C2") throw new Exception("COM IID changed");
checks++;
// A wrong-unit oracle must fail; this assertion is not a rewritten consumer.
if (Math.Abs(natural.Width - 96) < 1e-6) throw new Exception("Incorrect px-to-pt mutation was not detected");
checks++;
Console.WriteLine(JsonSerializer.Serialize(new { status = "PASS", checks, installedAssemblySha256 = hash,
    natural = new { widthPt = natural.Width, heightPt = natural.Height }, abi, officeComCalls = 0 }));
