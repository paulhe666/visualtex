using System.Text.Json;
using VisualTeX.WordVsto;

namespace VisualTeX.VstoFlowAcceptance;
internal static partial class Program
{
    // Read native serialization from newly created, contained OLE objects only.
    // No Word application, active object, user selection or global clipboard.
    private static void RunMathTypeNativeCharacterOracle(string root)
    {
        var cases = new[] {
            ("plain-control", "<mi>x</mi><mo>+</mo><mn>1</mn>"),
            ("double-dot-group", "<mover accent='true'><mrow><mi>a</mi><mi>b</mi></mrow><mo>¨</mo></mover>"),
            ("double-dot-primes", "<mover accent='true'><mrow><mi>′</mi><mi>′</mi></mrow><mo>¨</mo></mover>"),
            ("bold-hbar", "<mi mathvariant='bold-italic'>ℏ</mi>"),
            ("bold-jmath", "<mi mathvariant='bold-italic'>ȷ</mi>"),
            ("bold-real", "<mi mathvariant='bold'>ℜ</mi>"),
        };
        var results = new List<object>();
        foreach (var item in cases)
        {
            var source = "<math xmlns='http://www.w3.org/1998/Math/MathML'>" + item.Item2 + "</math>";
            try
            {
                // Separate native CF_TEXT control for SDKs that advertise no
                // MathML SetData. This remains diagnostic-only, never a product
                // import or semantic-validation fallback.
                var tex = MathMlToLatexConverter.Convert(source);
                var native = MathTypeOleInterop.CreateServerBackedCompoundFileFromText("$" + tex + "$");
                File.WriteAllBytes(Path.Combine(root, item.Item1 + ".ole"), native);
                var bytes = MathTypeOleStorage.ReadEquationNative(native);
                File.WriteAllBytes(Path.Combine(root,item.Item1 + ".equation-native"),bytes);
                var actual = MathTypeOleStorage.ReadMathMl(native);
                File.WriteAllText(Path.Combine(root,item.Item1 + ".mathml"),actual);
                results.Add(new { Name=item.Item1, Source=source, Actual=actual, Native=Convert.ToBase64String(bytes) });
                Console.WriteLine(item.Item1 + "=" + actual);
            }
            catch(Exception e)
            {
                results.Add(new { Name=item.Item1, Source=source, Error=e.ToString() });
                Console.WriteLine(item.Item1 + " ERROR=" + e.Message);
            }
            File.WriteAllText(Path.Combine(root,"results.json"),JsonSerializer.Serialize(results,new JsonSerializerOptions { WriteIndented=true }));
        }
    }
}
