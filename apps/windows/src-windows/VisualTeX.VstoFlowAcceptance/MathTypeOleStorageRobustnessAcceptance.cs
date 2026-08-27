using System.Diagnostics;
using VisualTeX.WordVsto;

namespace VisualTeX.VstoFlowAcceptance;

internal static partial class Program
{
    private static void RunMathTypeOleStorageRobustnessAcceptance()
    {
        static void AssertRejected(byte[] payload, string description)
        {
            bool accepted;
            try
            {
                accepted = MathTypeOleStorage.LooksLikeMathTypeCompoundFile(payload);
            }
            catch (Exception error)
            {
                throw new InvalidDataException(
                    $"MathType compound-file detection threw for {description}.",
                    error);
            }
            AssertTrue(!accepted,
                $"MathType compound-file detection accepted malformed input: {description}.");
        }

        AssertRejected(Array.Empty<byte>(), "empty input");
        AssertRejected(new byte[] { 0xD0, 0xCF, 0x11, 0xE0 }, "truncated CFB signature");
        AssertRejected(new byte[511], "one-byte-short CFB header");

        var random = new Random(0x56495458);
        var stopwatch = Stopwatch.StartNew();
        for (var index = 0; index < 750; index++)
        {
            var length = index < 256
                ? index
                : random.Next(0, 256 * 1024);
            var payload = new byte[length];
            random.NextBytes(payload);
            if (payload.Length >= 8 && index % 3 == 0)
            {
                payload[0] = 0xD0;
                payload[1] = 0xCF;
                payload[2] = 0x11;
                payload[3] = 0xE0;
                payload[4] = 0xA1;
                payload[5] = 0xB1;
                payload[6] = 0x1A;
                payload[7] = 0xE1;
            }

            try
            {
                _ = MathTypeOleStorage.LooksLikeMathTypeCompoundFile(payload);
            }
            catch (Exception error)
            {
                throw new InvalidDataException(
                    $"MathType compound-file detection was not failure-safe for deterministic mutation {index} ({payload.Length} bytes).",
                    error);
            }
        }
        stopwatch.Stop();

        AssertTrue(stopwatch.Elapsed < TimeSpan.FromSeconds(15),
            $"MathType malformed-storage rejection took too long: {stopwatch.Elapsed.TotalSeconds:F2}s.");
        Console.WriteLine(
            $"[MATHTYPE STORAGE ROBUSTNESS] 750 deterministic malformed/truncated payloads were rejected or safely ignored in {stopwatch.ElapsedMilliseconds}ms without exceptions or hangs.");
    }
}
