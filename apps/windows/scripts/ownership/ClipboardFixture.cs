using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;

// The runner inserts the production nested class verbatim. Only native boundaries are faked.
static class Harness
{
/* PRODUCTION_CLIPBOARD_TRANSACTION */
    static readonly IDataObject Previous = new Data();
    static readonly IDataObject Source = new Data();
    static readonly List<object> Released = new();
    static string Failure = "";
    static uint Sequence;
    static int Restores, Uninitializes;
    static object? Clipboard;
    static void Reset(string failure = "")
    {
        Failure = failure; Sequence = 100; Restores = 0; Uninitializes = 0;
        Clipboard = Previous; Released.Clear();
    }
    static void FailAt(string name) { if (Failure == name) throw new InvalidOperationException(name); }
    static int OleInitialize(IntPtr _) { FailAt("initialize"); return 0; }
    static void OleUninitialize() { Uninitializes++; }
    static uint GetClipboardSequenceNumber() => Sequence;
    static IDataObject TryGetOleClipboard(out object value)
    { value = Previous; FailAt("previous"); return Previous; }
    static IDataObject GetOleClipboard(out object value)
    { value = Source; FailAt("source"); return Source; }
    static byte[] ReadEmbeddedObject(IDataObject _) { FailAt("read"); return new byte[] { 1 }; }
    static void ValidateCompoundFile(byte[] _) { FailAt("validate"); }
    static void Release(object? value) { if (value != null) Released.Add(value); }
    static int OleSetClipboard(object? value)
    {
        if (ReferenceEquals(value, Previous)) { Restores++; FailAt("restore"); }
        Clipboard = value; Sequence++; return 0;
    }
    static void Copy() { Clipboard = Source; Sequence++; FailAt("copy-after-write"); }
    sealed class MathTypeOleClipboardProxy
    {
        public MathTypeOleClipboardProxy(IDataObject _, byte[] bytes, string? path, bool prefer, bool external) { }
        public int StorageWriteCount => 0;
    }
    sealed class Data : IDataObject
    {
        public void GetData(ref FORMATETC f, out STGMEDIUM m) => throw new NotSupportedException();
        public void GetDataHere(ref FORMATETC f, ref STGMEDIUM m) => throw new NotSupportedException();
        public int QueryGetData(ref FORMATETC f) => throw new NotSupportedException();
        public int GetCanonicalFormatEtc(ref FORMATETC f, out FORMATETC o) => throw new NotSupportedException();
        public void SetData(ref FORMATETC f, ref STGMEDIUM m, bool release) => throw new NotSupportedException();
        public IEnumFORMATETC EnumFormatEtc(DATADIR d) => throw new NotSupportedException();
        public int DAdvise(ref FORMATETC f, ADVF a, IAdviseSink s, out int connection) => throw new NotSupportedException();
        public void DUnadvise(int c) => throw new NotSupportedException();
        public int EnumDAdvise(out IEnumSTATDATA? d) => throw new NotSupportedException();
    }
    static void Check(bool passed, string label) { if (!passed) throw new Exception(label); }
    static void ExpectFailure(Action action, string label)
    {
        try { action(); } catch (InvalidOperationException e) when (e.Message == label) { return; }
        throw new Exception("Expected failure: " + label);
    }
    static void Main()
    {
        var count = 0;
        foreach (var point in new[] { "previous", "copy-after-write", "source", "read", "validate" })
        {
            Reset(point);
            ExpectFailure(() => new ClipboardTransaction(Copy), point);
            Check(Uninitializes == 1, point + " OLE lifetime");
            Check(Released.Contains(Previous), point + " previous release");
            Check(Released.Contains(Source) == (point == "source" || point == "read" || point == "validate"), point + " source release");
            Check(Restores == (point == "previous" ? 0 : 1), point + " restoration count");
            Check(ReferenceEquals(Clipboard, Previous), point + " clipboard preserved");
            count++;
        }
        Reset("copy-before-write");
        ExpectFailure(() => new ClipboardTransaction(() => FailAt("copy-before-write")), "copy-before-write");
        Check(Restores == 0 && Uninitializes == 1 && Released.Count == 1, "unchanged failed copy"); count++;
        Reset("initialize");
        ExpectFailure(() => new ClipboardTransaction(Copy), "initialize");
        Check(Uninitializes == 0 && Released.Count == 0, "no resources before init"); count++;
        foreach (var replacement in new[] { false, true })
        {
            Reset(); var tx = new ClipboardTransaction(Copy);
            if (replacement) tx.SetReplacementClipboard(new byte[] {1}, null, standaloneExternal: true);
            tx.Dispose(); tx.Dispose();
            Check(Restores == 1 && Uninitializes == 1 && Released.Count == 2, "balanced once");
            Check(ReferenceEquals(Clipboard, Previous), "original clipboard restored"); count++;
        }
        Reset(); var changed = new ClipboardTransaction(Copy);
        var external = new object(); OleSetClipboard(external); changed.Dispose();
        Check(Restores == 0 && ReferenceEquals(Clipboard, external), "new clipboard retained");
        Check(Uninitializes == 1 && Released.Count == 2, "external change resources released"); count++;
        Reset(); var brokenRestore = new ClipboardTransaction(Copy); Failure = "restore";
        brokenRestore.Dispose();
        Check(Uninitializes == 1 && Released.Count == 2, "restoration failure releases resources"); count++;
        Console.WriteLine($"Clipboard transaction ownership passed {count}/{count}; native boundaries mocked, no system clipboard used.");
    }
}
