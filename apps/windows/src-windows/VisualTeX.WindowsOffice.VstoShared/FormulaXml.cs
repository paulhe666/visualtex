using System.Xml;
using System.Xml.Linq;

namespace VisualTeX.WordVsto;

// Input boundary for formula converters. Enforce complexity WHILE loading the
// document, before recursive normalization/serialization. No extra preflight
// parse, external entities, or content normalization is involved.
internal static class FormulaXml
{
    internal const int MaximumCharacters = 4_000_000;
    internal const int MaximumDepth = 256;
    internal const int MaximumElements = 250_000;

    internal static XDocument Parse(string xml, string kind, LoadOptions options = LoadOptions.PreserveWhitespace)
    {
        using var reader = CreateReader(xml, kind);
        return XDocument.Load(reader, options);
    }

    internal static XmlReader CreateReader(string xml, string kind)
    {
        if (xml is null) throw new ArgumentNullException(nameof(xml));
        if (xml.Length > MaximumCharacters)
            throw new InvalidDataException($"{kind} exceeds the formula limit of {MaximumCharacters} characters.");
        return new BoundedReader(XmlReader.Create(new StringReader(xml), new XmlReaderSettings
        {
            DtdProcessing = DtdProcessing.Prohibit,
            XmlResolver = null,
            MaxCharactersInDocument = MaximumCharacters,
            IgnoreWhitespace = false,
            CloseInput = true,
        }), kind);
    }

    private sealed class BoundedReader : XmlReader
    {
        private readonly XmlReader _reader;
        private readonly string _kind;
        private int _elements;

        internal BoundedReader(XmlReader reader, string kind) { _reader = reader; _kind = kind; }
        public override bool Read()
        {
            if (!_reader.Read()) return false;
            if (_reader.NodeType == XmlNodeType.Element)
            {
                if (_reader.Depth >= MaximumDepth)
                    throw new InvalidDataException($"{_kind} exceeds the formula nesting limit of {MaximumDepth} levels.");
                if (++_elements > MaximumElements)
                    throw new InvalidDataException($"{_kind} exceeds the formula limit of {MaximumElements} elements.");
            }
            return true;
        }

        public override int AttributeCount => _reader.AttributeCount;
        public override string BaseURI => _reader.BaseURI;
        public override int Depth => _reader.Depth;
        public override bool EOF => _reader.EOF;
        public override bool HasValue => _reader.HasValue;
        public override bool IsEmptyElement => _reader.IsEmptyElement;
        public override string LocalName => _reader.LocalName;
        public override string NamespaceURI => _reader.NamespaceURI;
        public override XmlNameTable NameTable => _reader.NameTable;
        public override XmlNodeType NodeType => _reader.NodeType;
        public override string Prefix => _reader.Prefix;
        public override ReadState ReadState => _reader.ReadState;
        public override string Value => _reader.Value;
        public override string GetAttribute(int i) => _reader.GetAttribute(i);
        public override string? GetAttribute(string name) => _reader.GetAttribute(name);
        public override string? GetAttribute(string name, string? namespaceURI) => _reader.GetAttribute(name, namespaceURI);
        public override string? LookupNamespace(string prefix) => _reader.LookupNamespace(prefix);
        public override bool MoveToAttribute(string name) => _reader.MoveToAttribute(name);
        public override bool MoveToAttribute(string name, string? ns) => _reader.MoveToAttribute(name, ns);
        public override void MoveToAttribute(int i) => _reader.MoveToAttribute(i);
        public override bool MoveToElement() => _reader.MoveToElement();
        public override bool MoveToFirstAttribute() => _reader.MoveToFirstAttribute();
        public override bool MoveToNextAttribute() => _reader.MoveToNextAttribute();
        public override bool ReadAttributeValue() => _reader.ReadAttributeValue();
        public override void ResolveEntity() => _reader.ResolveEntity();
        public override void Close() => _reader.Close();
    }
}
