using System.Globalization;
using System.Xml.Linq;
using static VisualTeX.WordVsto.MathTypeMtefCodec;
using static VisualTeX.WordVsto.MathTypeMtefAssembly;
using static VisualTeX.WordVsto.MathTypeMtefCharacters;
using static VisualTeX.WordVsto.MathTypeMtefPrefix;
using static VisualTeX.WordVsto.MathTypeMtefSignature;
using static VisualTeX.WordVsto.MathTypeMtefTemplates;

namespace VisualTeX.WordVsto;

/// <summary>
/// Native glyphs and fonts: the glyph contracts MathType stores, its merged alternates, font definitions and
/// explicit character styles.
/// </summary>
internal static class MathTypeMtefGlyphs
{
    // Exact native glyph contracts, shared by font binding, CHAR emission and
    // decoding. Do not infer a glyph from a font's Unicode coverage. The corner
    // positions are Design Science's eumat1 AFM/post names, cross-checked with
    // lcdf-typetools texglyphlist and native WMF pixels. E92F is MathType's own
    // shortparallel MTCode; its EuclidMath1 position 89 renders the short bars.
    internal sealed class NativeGlyphAuthority
    {
        internal NativeGlyphAuthority(int scalar, bool alternate, string font,
            string encoding, int mtCode, byte encoded8, string tokenClass, string glyphName)
        {
            Scalar = scalar; Alternate = alternate; Font = font; Encoding = encoding;
            MtCode = mtCode; Encoded8 = encoded8; TokenClass = tokenClass; GlyphName = glyphName;
        }
        internal int Scalar { get; }
        internal bool Alternate { get; }
        internal string Font { get; }
        internal string Encoding { get; }
        internal int MtCode { get; }
        internal byte Encoded8 { get; }
        internal string TokenClass { get; }
        internal string GlyphName { get; }
    }

    private static readonly NativeGlyphAuthority[] NativeGlyphAuthorities =
    {
        // MT Extra's planckover2pi slot is the TeX hbar glyph. Bind it through
        // the same native font/style path as other alternate glyphs; selecting
        // the built-in Extra face directly discards the identifier's style.
        new(0x210F, true, "MT Extra", "MTExtra", 0x210F, 0x68, "mi", "planckover2pi"),
        new(0x231C, false, "Euclid Math One", "EuclidMath1", 0x231C, 0x84, "mo", "rightanglenw"),
        new(0x231D, false, "Euclid Math One", "EuclidMath1", 0x231D, 0x86, "mo", "rightanglene"),
        new(0x231E, false, "Euclid Math One", "EuclidMath1", 0x231E, 0x85, "mo", "rightanglesw"),
        new(0x231F, false, "Euclid Math One", "EuclidMath1", 0x231F, 0x87, "mo", "rightanglese"),
        new(0x2225, true, "Euclid Math One", "EuclidMath1", 0xE92F, 0x89, "mo", "parallelvar"),
        new(0x2226, false, "Euclid Math One", "EuclidMath1", 0x2226, 0x8D, "mo", "notparallelto"),
        new(0x2226, true, "Euclid Math One", "EuclidMath1", 0xEA2F, 0x8B, "mo", "notparalleltovar"),
        // Vendor Texvc input + Euclid AFM/post/cmap authority; alternate slots
        // remain distinct from their ordinary Unicode counterparts.
        new(0x2223, true, "Euclid Math One", "EuclidMath1", 0xE92E, 0x88, "mo", "dividesvar"),
        new(0x2224, false, "Euclid Math One", "EuclidMath1", 0x2224, 0x8C, "mo", "doesnotdivide"),
        new(0x2224, true, "Euclid Math One", "EuclidMath1", 0xEA2E, 0x8A, "mo", "doesnotdividevar"),
        new(0x2216, true, "Euclid Math One", "EuclidMath1", 0x2216, 0x80, "mo", "integerdivide"),
        new(0x228A, false, "Euclid Math Two", "EuclidMath2", 0x228A, 0xD6, "mo", "subsetnotequal"),
        new(0x228A, true, "Euclid Math Two", "EuclidMath2", 0xEA42, 0xD8, "mo", "notsubsetoreql"),
        new(0x228B, false, "Euclid Math Two", "EuclidMath2", 0x228B, 0xD7, "mo", "supersetnotequal"),
        new(0x228B, true, "Euclid Math Two", "EuclidMath2", 0xEA43, 0xD9, "mo", "notsupersetoreql"),
        new(0x2288, false, "Euclid Math Two", "EuclidMath2", 0x2288, 0xDA, "mo", "notreflexsubset"),
        new(0x2288, true, "Euclid Math Two", "EuclidMath2", 0xEA0C, 0xE0, "mo", "notreflexsubsetdbl"),
        new(0x2289, false, "Euclid Math Two", "EuclidMath2", 0x2289, 0xDB, "mo", "notreflexsuperset"),
        new(0x2289, true, "Euclid Math Two", "EuclidMath2", 0xEA0B, 0xE1, "mo", "notreflexsupersetdbl"),
        new(0x2ACB, false, "Euclid Math Two", "EuclidMath2", 0xEA44, 0xDC, "mo", "subsetornotdbleql"),
        new(0x2ACB, true, "Euclid Math Two", "EuclidMath2", 0xEA46, 0xDE, "mo", "notsubsetordblequal"),
        new(0x2ACC, false, "Euclid Math Two", "EuclidMath2", 0xEA45, 0xDD, "mo", "supersetornotdbleql"),
        new(0x2ACC, true, "Euclid Math Two", "EuclidMath2", 0xEA47, 0xDF, "mo", "notsupersetordblequal"),
        new(0x2220, false, "Symbol", "Symbol", 0x2220, 0xD0, "mi", "angle"),
        new(0x2221, false, "MT Extra", "MTExtra", 0x2221, 0x52, "mi", "measuredangle"),
        new(0x2222, false, "MT Extra", "MTExtra", 0x2222, 0x53, "mi", "sphericalangle"),
        // MathType 7.9 importer authority: the font position MathType itself
        // stores for each command (MTSetEqnFromLangStr with MathJax MathML and
        // TeX, read back as MTEF). Plain glyphs only; italic Euclid letters and
        // the Greek alphabets keep their own paths above.
        new(0x00AE, false, "Symbol", "Symbol", 0x00AE, 0xD2, "mi", "circledR"),
        new(0x2035, true, "MT Extra", "MTExtra", 0x2035, 0x60, "mi", "backprime"),
        new(0x2127, false, "MT Extra", "MTExtra", 0x2127, 0x4A, "mi", "mho"),
        new(0x219A, false, "Euclid Math One", "EuclidMath1", 0x219A, 0xBC, "mo", "nleftarrow"),
        new(0x219B, false, "Euclid Math One", "EuclidMath1", 0x219B, 0xBD, "mo", "nrightarrow"),
        new(0x219E, false, "Euclid Math One", "EuclidMath1", 0x219E, 0xB3, "mo", "twoheadleftarrow"),
        new(0x21A0, false, "Euclid Math One", "EuclidMath1", 0x21A0, 0xB2, "mo", "twoheadrightarrow"),
        new(0x21A2, false, "Euclid Math One", "EuclidMath1", 0x21A2, 0xB1, "mo", "leftarrowtail"),
        new(0x21A3, false, "Euclid Math One", "EuclidMath1", 0x21A3, 0xB0, "mo", "rightarrowtail"),
        new(0x21AB, false, "Euclid Math One", "EuclidMath1", 0x21AB, 0xCD, "mo", "looparrowleft"),
        new(0x21AC, false, "Euclid Math One", "EuclidMath1", 0x21AC, 0xCC, "mo", "looparrowright"),
        new(0x21AD, false, "Euclid Math One", "EuclidMath1", 0x21AD, 0xCF, "mo", "leftrightsquigarrow"),
        new(0x21AE, false, "Euclid Math One", "EuclidMath1", 0x21AE, 0xC0, "mo", "nleftrightarrow"),
        new(0x21B0, false, "Euclid Math One", "EuclidMath1", 0x21B0, 0xCB, "mo", "Lsh"),
        new(0x21B1, false, "Euclid Math One", "EuclidMath1", 0x21B1, 0xCA, "mo", "Rsh"),
        new(0x21B6, false, "Euclid Math One", "EuclidMath1", 0x21B6, 0xD1, "mo", "curvearrowleft"),
        new(0x21B7, false, "Euclid Math One", "EuclidMath1", 0x21B7, 0xD0, "mo", "curvearrowright"),
        new(0x21BA, false, "Euclid Math One", "EuclidMath1", 0x21BA, 0xD3, "mo", "circlearrowleft"),
        new(0x21BB, false, "Euclid Math One", "EuclidMath1", 0x21BB, 0xD2, "mo", "circlearrowright"),
        new(0x21BC, false, "Euclid Math One", "EuclidMath1", 0x21BC, 0xB4, "mo", "leftharpoonup"),
        new(0x21BE, false, "Euclid Math One", "EuclidMath1", 0x21BE, 0xB8, "mo", "upharpoonright"),
        new(0x21BF, false, "Euclid Math One", "EuclidMath1", 0x21BF, 0xBA, "mo", "upharpoonleft"),
        new(0x21C1, false, "Euclid Math One", "EuclidMath1", 0x21C1, 0xB7, "mo", "rightharpoondown"),
        new(0x21C2, false, "Euclid Math One", "EuclidMath1", 0x21C2, 0xB9, "mo", "downharpoonright"),
        new(0x21C3, false, "Euclid Math One", "EuclidMath1", 0x21C3, 0xBB, "mo", "downharpoonleft"),
        new(0x21C4, false, "MT Extra", "MTExtra", 0x21C4, 0x80, "mo", "rightleftarrows"),
        new(0x21C6, false, "Euclid Math One", "EuclidMath1", 0x21C6, 0xC6, "mo", "leftrightarrows"),
        new(0x21C7, false, "Euclid Math One", "EuclidMath1", 0x21C7, 0xC2, "mo", "leftleftarrows"),
        new(0x21C8, false, "Euclid Math One", "EuclidMath1", 0x21C8, 0xC4, "mo", "upuparrows"),
        new(0x21C9, false, "Euclid Math One", "EuclidMath1", 0x21C9, 0xC3, "mo", "rightrightarrows"),
        new(0x21CA, false, "Euclid Math One", "EuclidMath1", 0x21CA, 0xC5, "mo", "downdownarrows"),
        new(0x21CB, false, "Euclid Math One", "EuclidMath1", 0x21CB, 0xC7, "mo", "leftrightharpoons"),
        new(0x21CC, true, "MT Extra", "MTExtra", 0x21CC, 0x83, "mo", "rightleftharpoons"),
        new(0x21CD, false, "Euclid Math One", "EuclidMath1", 0x21CD, 0xBE, "mo", "nLeftarrow"),
        new(0x21CE, false, "Euclid Math One", "EuclidMath1", 0x21CE, 0xC1, "mo", "nLeftrightarrow"),
        new(0x21CF, false, "Euclid Math One", "EuclidMath1", 0x21CF, 0xBF, "mo", "nRightarrow"),
        new(0x21DA, false, "Euclid Math One", "EuclidMath1", 0x21DA, 0xC9, "mo", "Lleftarrow"),
        new(0x21DB, false, "Euclid Math One", "EuclidMath1", 0x21DB, 0xC8, "mo", "Rrightarrow"),
        new(0x21DD, false, "Euclid Math One", "EuclidMath1", 0x21DD, 0xCE, "mo", "rightsquigarrow"),
        new(0x220D, false, "Symbol", "Symbol", 0x220D, 0x27, "mo", "backepsilon"),
        new(0x2214, false, "Euclid Math One", "EuclidMath1", 0x2214, 0xEA, "mo", "dotplus"),
        new(0x2216, false, "Euclid Math One", "EuclidMath1", 0x2216, 0x82, "mo", "setminus"),
        new(0x2218, false, "MT Extra", "MTExtra", 0x2218, 0x6F, "mo", "circ"),
        new(0x2235, false, "MT Extra", "MTExtra", 0x2235, 0x51, "mo", "because"),
        new(0x223C, false, "Symbol", "Symbol", 0x223C, 0x7E, "mo", "sim"),
        new(0x2240, false, "Euclid Math One", "EuclidMath1", 0x2240, 0xAA, "mo", "wr"),
        new(0x2241, false, "Euclid Math One", "EuclidMath1", 0x2241, 0xAD, "mo", "nsim"),
        new(0x2242, false, "Euclid Math One", "EuclidMath1", 0x2242, 0xAB, "mo", "eqsim"),
        new(0x2247, false, "Euclid Math One", "EuclidMath1", 0x2247, 0xAE, "mo", "ncong"),
        new(0x224A, false, "Euclid Math One", "EuclidMath1", 0x224A, 0xAC, "mo", "approxeq"),
        new(0x224D, false, "Euclid Math One", "EuclidMath1", 0x224D, 0xA9, "mo", "asymp"),
        new(0x224E, false, "Euclid Math One", "EuclidMath1", 0x224E, 0xA7, "mo", "Bumpeq"),
        new(0x224F, false, "Euclid Math One", "EuclidMath1", 0x224F, 0xA6, "mo", "bumpeq"),
        new(0x2251, false, "Euclid Math One", "EuclidMath1", 0x2251, 0xA4, "mo", "doteqdot"),
        new(0x2253, false, "Euclid Math One", "EuclidMath1", 0x2253, 0xA3, "mo", "risingdotseq"),
        new(0x2256, false, "Euclid Math One", "EuclidMath1", 0x2256, 0xA2, "mo", "eqcirc"),
        new(0x2257, false, "Euclid Math One", "EuclidMath1", 0x2257, 0xA1, "mo", "circeq"),
        new(0x225C, false, "MT Extra", "MTExtra", 0x225C, 0x40, "mo", "triangleq"),
        new(0x2268, false, "Euclid Math Two", "EuclidMath2", 0x2268, 0x98, "mo", "lneqq"),
        new(0x2269, false, "Euclid Math Two", "EuclidMath2", 0x2269, 0x99, "mo", "gneqq"),
        new(0x2270, false, "Euclid Math Two", "EuclidMath2", 0x2270, 0x92, "mo", "nleq"),
        new(0x2271, false, "Euclid Math Two", "EuclidMath2", 0x2271, 0x93, "mo", "ngeq"),
        new(0x2272, false, "Euclid Math Two", "EuclidMath2", 0x2272, 0x88, "mo", "lesssim"),
        new(0x2273, false, "Euclid Math Two", "EuclidMath2", 0x2273, 0x89, "mo", "gtrsim"),
        new(0x2276, false, "Euclid Math Two", "EuclidMath2", 0x2276, 0xA3, "mo", "lessgtr"),
        new(0x2277, false, "Euclid Math Two", "EuclidMath2", 0x2277, 0xA4, "mo", "gtrless"),
        new(0x227C, false, "Euclid Math Two", "EuclidMath2", 0x227C, 0xB0, "mo", "preccurlyeq"),
        new(0x227D, false, "Euclid Math Two", "EuclidMath2", 0x227D, 0xB1, "mo", "succcurlyeq"),
        new(0x227E, false, "Euclid Math Two", "EuclidMath2", 0x227E, 0xB6, "mo", "precsim"),
        new(0x227F, false, "Euclid Math Two", "EuclidMath2", 0x227F, 0xB7, "mo", "succsim"),
        new(0x2280, false, "Euclid Math Two", "EuclidMath2", 0x2280, 0xBA, "mo", "nprec"),
        new(0x2281, false, "Euclid Math Two", "EuclidMath2", 0x2281, 0xBB, "mo", "nsucc"),
        new(0x228E, false, "Euclid Math Two", "EuclidMath2", 0x228E, 0xE2, "mo", "uplus"),
        new(0x228F, false, "Euclid Math Two", "EuclidMath2", 0x228F, 0xF0, "mo", "sqsubset"),
        new(0x2290, false, "Euclid Math Two", "EuclidMath2", 0x2290, 0xF1, "mo", "sqsupset"),
        new(0x2291, false, "Euclid Math Two", "EuclidMath2", 0x2291, 0xF4, "mo", "sqsubseteq"),
        new(0x2292, false, "Euclid Math Two", "EuclidMath2", 0x2292, 0xF5, "mo", "sqsupseteq"),
        new(0x2293, false, "Euclid Math Two", "EuclidMath2", 0x2293, 0xF3, "mo", "sqcap"),
        new(0x2294, false, "Euclid Math Two", "EuclidMath2", 0x2294, 0xF2, "mo", "sqcup"),
        new(0x2296, false, "Euclid Math One", "EuclidMath1", 0x2296, 0x21, "mo", "ominus"),
        new(0x2298, false, "Euclid Math One", "EuclidMath1", 0x2298, 0x25, "mo", "oslash"),
        new(0x229A, false, "Euclid Math One", "EuclidMath1", 0x229A, 0x22, "mo", "circledcirc"),
        new(0x229B, false, "Euclid Math One", "EuclidMath1", 0x229B, 0x23, "mo", "circledast"),
        new(0x229E, false, "Euclid Math One", "EuclidMath1", 0x229E, 0x28, "mo", "boxplus"),
        new(0x229F, false, "Euclid Math One", "EuclidMath1", 0x229F, 0x27, "mo", "boxminus"),
        new(0x22A0, false, "Euclid Math One", "EuclidMath1", 0x22A0, 0x29, "mo", "boxtimes"),
        new(0x22A1, false, "Euclid Math One", "EuclidMath1", 0x22A1, 0x2A, "mo", "boxdot"),
        new(0x22A2, false, "Euclid Math One", "EuclidMath1", 0x22A2, 0x90, "mo", "vdash"),
        new(0x22A3, false, "Euclid Math One", "EuclidMath1", 0x22A3, 0x94, "mo", "dashv"),
        new(0x22A4, false, "Euclid Math One", "EuclidMath1", 0x22A4, 0x95, "mi", "top"),
        new(0x22A8, false, "Euclid Math One", "EuclidMath1", 0x22A8, 0x91, "mo", "models"),
        new(0x22A9, false, "Euclid Math One", "EuclidMath1", 0x22A9, 0x92, "mo", "Vdash"),
        new(0x22AA, false, "Euclid Math One", "EuclidMath1", 0x22AA, 0x93, "mo", "Vvdash"),
        new(0x22AC, false, "Euclid Math One", "EuclidMath1", 0x22AC, 0x96, "mo", "nvdash"),
        new(0x22AD, false, "Euclid Math One", "EuclidMath1", 0x22AD, 0x97, "mo", "nvDash"),
        new(0x22AE, false, "Euclid Math One", "EuclidMath1", 0x22AE, 0x98, "mo", "nVdash"),
        new(0x22AF, false, "Euclid Math One", "EuclidMath1", 0x22AF, 0x99, "mo", "nVDash"),
        new(0x22B4, false, "Euclid Math Two", "EuclidMath2", 0x22B4, 0x28, "mo", "trianglelefteq"),
        new(0x22B5, false, "Euclid Math Two", "EuclidMath2", 0x22B5, 0x27, "mo", "trianglerighteq"),
        new(0x22B8, false, "Euclid Math One", "EuclidMath1", 0x22B8, 0xA8, "mo", "multimap"),
        new(0x22BA, false, "Euclid Math One", "EuclidMath1", 0x22BA, 0xFA, "mo", "intercal"),
        new(0x22BB, false, "Euclid Math Two", "EuclidMath2", 0x22BB, 0xA9, "mo", "veebar"),
        new(0x22BC, false, "Euclid Math Two", "EuclidMath2", 0x22BC, 0xAA, "mo", "barwedge"),
        new(0x22C6, false, "Euclid Math One", "EuclidMath1", 0x22C6, 0xE5, "mo", "star"),
        new(0x22C7, false, "Euclid Math One", "EuclidMath1", 0x22C7, 0xE6, "mo", "divideontimes"),
        new(0x22C9, false, "Euclid Math One", "EuclidMath1", 0x22C9, 0xE2, "mo", "ltimes"),
        new(0x22CA, false, "Euclid Math One", "EuclidMath1", 0x22CA, 0xE3, "mo", "rtimes"),
        new(0x22CB, false, "Euclid Math One", "EuclidMath1", 0x22CB, 0xE0, "mo", "leftthreetimes"),
        new(0x22CC, false, "Euclid Math One", "EuclidMath1", 0x22CC, 0xE1, "mo", "rightthreetimes"),
        new(0x22CE, false, "Euclid Math Two", "EuclidMath2", 0x22CE, 0xCB, "mo", "curlyvee"),
        new(0x22CF, false, "Euclid Math Two", "EuclidMath2", 0x22CF, 0xCA, "mo", "curlywedge"),
        new(0x22D0, false, "Euclid Math Two", "EuclidMath2", 0x22D0, 0xD0, "mo", "Subset"),
        new(0x22D1, false, "Euclid Math Two", "EuclidMath2", 0x22D1, 0xD1, "mo", "Supset"),
        new(0x22D2, false, "Euclid Math Two", "EuclidMath2", 0x22D2, 0xD3, "mo", "Cap"),
        new(0x22D3, false, "Euclid Math Two", "EuclidMath2", 0x22D3, 0xD2, "mo", "Cup"),
        new(0x22D4, false, "Euclid Math Two", "EuclidMath2", 0x22D4, 0xE3, "mo", "pitchfork"),
        new(0x22D6, false, "Euclid Math Two", "EuclidMath2", 0x22D6, 0x80, "mo", "lessdot"),
        new(0x22D7, false, "Euclid Math Two", "EuclidMath2", 0x22D7, 0x81, "mo", "gtrdot"),
        new(0x22D8, false, "Euclid Math Two", "EuclidMath2", 0x22D8, 0x8C, "mo", "lll"),
        new(0x22D9, false, "Euclid Math Two", "EuclidMath2", 0x22D9, 0x8D, "mo", "ggg"),
        new(0x22DA, false, "Euclid Math Two", "EuclidMath2", 0x22DA, 0xA5, "mo", "lesseqgtr"),
        new(0x22DB, false, "Euclid Math Two", "EuclidMath2", 0x22DB, 0xA6, "mo", "gtreqless"),
        new(0x22DE, false, "Euclid Math Two", "EuclidMath2", 0x22DE, 0xB2, "mo", "curlyeqprec"),
        new(0x22DF, false, "Euclid Math Two", "EuclidMath2", 0x22DF, 0xB3, "mo", "curlyeqsucc"),
        new(0x22E0, true, "Euclid Math Two", "EuclidMath2", 0x22E0, 0xBE, "mo", "npreceq"),
        new(0x22E1, true, "Euclid Math Two", "EuclidMath2", 0x22E1, 0xBF, "mo", "nsucceq"),
        new(0x22E6, false, "Euclid Math Two", "EuclidMath2", 0x22E6, 0x9E, "mo", "lnsim"),
        new(0x22E7, false, "Euclid Math Two", "EuclidMath2", 0x22E7, 0x9F, "mo", "gnsim"),
        new(0x22E8, false, "Euclid Math Two", "EuclidMath2", 0x22E8, 0xC6, "mo", "precnsim"),
        new(0x22E9, false, "Euclid Math Two", "EuclidMath2", 0x22E9, 0xC7, "mo", "succnsim"),
        new(0x22EA, false, "Euclid Math Two", "EuclidMath2", 0x22EA, 0x2A, "mo", "ntriangleleft"),
        new(0x22EB, false, "Euclid Math Two", "EuclidMath2", 0x22EB, 0x29, "mo", "ntriangleright"),
        new(0x22EC, false, "Euclid Math Two", "EuclidMath2", 0x22EC, 0x2C, "mo", "ntrianglelefteq"),
        new(0x22ED, false, "Euclid Math Two", "EuclidMath2", 0x22ED, 0x2B, "mo", "ntrianglerighteq"),
        new(0x2322, false, "MT Extra", "MTExtra", 0x2322, 0x29, "mo", "frown"),
        new(0x2323, false, "MT Extra", "MTExtra", 0x2323, 0x28, "mo", "smile"),
        new(0x24C8, false, "Euclid Math One", "EuclidMath1", 0x24C8, 0x26, "mi", "circledS"),
        new(0x25B3, false, "MT Extra", "MTExtra", 0x25B3, 0x56, "mo", "triangle"),
        new(0x266E, false, "Euclid Math One", "EuclidMath1", 0x266E, 0xE8, "mi", "natural"),
        new(0x2A7D, false, "Euclid Math Two", "EuclidMath2", 0x2A7D, 0x84, "mo", "leqslant"),
        new(0x2A7E, false, "Euclid Math Two", "EuclidMath2", 0x2A7E, 0x85, "mo", "geqslant"),
        // Characters MathType's TeX importer stores in place of the MathJax
        // character (see MathTypeStoredCharacters).
        new(0x03DC, false, "Euclid Math One", "EuclidMath1", 0x03DC, 0xF4, "mi", "digamma"),
        new(0x22B2, false, "MT Extra", "MTExtra", 0x22B2, 0x3C, "mo", "triangleleft"),
        new(0x22B3, false, "MT Extra", "MTExtra", 0x22B3, 0x3E, "mo", "triangleright"),
        new(0x22DC, false, "Euclid Math Two", "EuclidMath2", 0x22DC, 0x82, "mo", "eqslantless"),
        new(0x22DD, false, "Euclid Math Two", "EuclidMath2", 0x22DD, 0x83, "mo", "eqslantgtr"),
        new(0x2306, false, "Euclid Math Two", "EuclidMath2", 0x2306, 0xAB, "mo", "doublebarwedge"),
        new(0x25A1, false, "MT Extra", "MTExtra", 0x25A1, 0x57, "mi", "square/Box"),
        new(0x2666, false, "Symbol", "Symbol", 0x2666, 0xA8, "mi", "blacklozenge"),
        // MathType 7.9 TeX importer authority, used where it keeps the
        // character and differs from the MathML importer: upright Greek
        // capitals, Hebrew letters and MathType's private negated relations.
new(0x0393, false, "Symbol", "Symbol", 0x0393, 0x47, "mi", "Gamma/varGamma"),
        new(0x0394, false, "Symbol", "Symbol", 0x0394, 0x44, "mi", "Delta/varDelta"),
        new(0x0398, false, "Symbol", "Symbol", 0x0398, 0x51, "mi", "Theta/varTheta"),
        new(0x039B, false, "Symbol", "Symbol", 0x039B, 0x4C, "mi", "Lambda/varLambda"),
        new(0x039E, false, "Symbol", "Symbol", 0x039E, 0x58, "mi", "Xi/varXi"),
        new(0x03A0, false, "Symbol", "Symbol", 0x03A0, 0x50, "mi", "Pi/varPi"),
        new(0x03A3, false, "Symbol", "Symbol", 0x03A3, 0x53, "mi", "Sigma/varSigma"),
        new(0x03A5, false, "Symbol", "Symbol", 0x03A5, 0x55, "mi", "Upsilon/varUpsilon"),
        new(0x03A6, false, "Symbol", "Symbol", 0x03A6, 0x46, "mi", "Phi/varPhi"),
        new(0x03A8, false, "Symbol", "Symbol", 0x03A8, 0x59, "mi", "Psi/varPsi"),
        new(0x03A9, false, "Symbol", "Symbol", 0x03A9, 0x57, "mi", "Omega/varOmega"),
        new(0x2132, false, "Euclid Math One", "EuclidMath1", 0x2132, 0xF5, "mi", "Finv"),
        new(0x2136, false, "Euclid Math One", "EuclidMath1", 0x2136, 0xFB, "mi", "beth"),
        new(0x2137, false, "Euclid Math One", "EuclidMath1", 0x2137, 0xFC, "mi", "gimel"),
        new(0x2138, false, "Euclid Math One", "EuclidMath1", 0x2138, 0xFD, "mi", "daleth"),
        new(0x2141, false, "Euclid Math One", "EuclidMath1", 0xED01, 0xF6, "mi", "Game"),
        new(0x2201, false, "Euclid Math One", "EuclidMath1", 0x2201, 0xF0, "mi", "complement"),
        new(0x2204, false, "Euclid Math One", "EuclidMath1", 0x2204, 0xF3, "mi", "nexists"),
        new(0x2270, true, "Euclid Math Two", "EuclidMath2", 0xEA06, 0x9C, "mo", "nleqq"),
        new(0x2271, true, "Euclid Math Two", "EuclidMath2", 0xEA07, 0x9D, "mo", "ngeqq"),
        new(0x22C5, true, "MT Extra", "MTExtra", 0xE98F, 0x67, "mo", "centerdot"),
        new(0x2720, false, "Euclid Math One", "EuclidMath1", 0x2720, 0xE4, "mi", "maltese"),
        new(0x2A85, false, "Euclid Math Two", "EuclidMath2", 0xE932, 0x8A, "mo", "lessapprox"),
        new(0x2A86, false, "Euclid Math Two", "EuclidMath2", 0xE933, 0x8B, "mo", "gtrapprox"),
        new(0x2A87, false, "Euclid Math Two", "EuclidMath2", 0xEA34, 0x90, "mo", "lneq"),
        new(0x2A88, false, "Euclid Math Two", "EuclidMath2", 0xEA35, 0x91, "mo", "gneq"),
        new(0x2A89, false, "Euclid Math Two", "EuclidMath2", 0xEA32, 0xA1, "mo", "lnapprox"),
        new(0x2A8A, false, "Euclid Math Two", "EuclidMath2", 0xEA33, 0xA2, "mo", "gnapprox"),
        new(0x2A8B, false, "Euclid Math Two", "EuclidMath2", 0xE922, 0xA7, "mo", "lesseqqgtr"),
        new(0x2A8C, false, "Euclid Math Two", "EuclidMath2", 0xE92D, 0xA8, "mo", "gtreqqless"),
        new(0x2AB5, false, "Euclid Math Two", "EuclidMath2", 0xEA40, 0xC4, "mo", "precneqq"),
        new(0x2AB6, false, "Euclid Math Two", "EuclidMath2", 0xEA41, 0xC5, "mo", "succneqq"),
        new(0x2AB7, false, "Euclid Math Two", "EuclidMath2", 0xE93A, 0xB8, "mo", "precapprox"),
        new(0x2AB8, false, "Euclid Math Two", "EuclidMath2", 0xE93B, 0xB9, "mo", "succapprox"),
        new(0x2AB9, false, "Euclid Math Two", "EuclidMath2", 0xEA3A, 0xC8, "mo", "precnapprox"),
        new(0x2ABA, false, "Euclid Math Two", "EuclidMath2", 0xEA3B, 0xC9, "mo", "succnapprox"),
        new(0x2AC5, false, "Euclid Math Two", "EuclidMath2", 0xE90C, 0xD4, "mo", "subseteqq"),
        new(0x2AC6, false, "Euclid Math Two", "EuclidMath2", 0xE90B, 0xD5, "mo", "supseteqq"),
    };

    // Variant commands MathType's importers store as another command's glyph
    // (the same font position): \vartriangle is \triangle, \thicksim is \sim.
    // The variant is written as that command, and the round-trip signature
    // compares the same canonical form.
    private static readonly Dictionary<int, (int Scalar, bool Alternate)> MathTypeMergedAlternates = new()
    {
        [0x2216] = (0x2216, false), // \smallsetminus -> \setminus
        [0x223C] = (0x223C, false), // \thicksim -> \sim
        [0x2268] = (0x2268, false), // \lvertneqq -> \lneqq
        [0x2269] = (0x2269, false), // \gvertneqq -> \gneqq
        [0x22A8] = (0x22A8, false), // \vDash -> \models
        [0x2322] = (0x2322, false), // \smallfrown -> \frown
        [0x2323] = (0x2323, false), // \smallsmile -> \smile
        [0x25B3] = (0x25B3, false), // \vartriangle -> \triangle
        [0x2A87] = (0x2270, true),  // \nleqslant -> \nleqq
        [0x2A88] = (0x2271, true),  // \ngeqslant -> \ngeqq
    };

    // Commands MathType's TeX importer stores as another character, whose
    // glyph MathType draws (it has none for the MathJax character):
    // \bullet is U+2022, \eqslantless is U+22DC, \blacklozenge is U+2666.
    private static readonly Dictionary<int, int> MathTypeStoredCharacters = new()
    {
        [0x03DD] = 0x03DC, // \digamma
        [0x2219] = 0x2022, // \bullet
        [0x229D] = 0x2296, // \circleddash
        [0x25B8] = 0x25B6, // \blacktriangleright
        [0x25B9] = 0x22B3, // \triangleright
        [0x25C2] = 0x25C0, // \blacktriangleleft
        [0x25C3] = 0x22B2, // \triangleleft
        [0x29EB] = 0x2666, // \blacklozenge
        [0x2A5E] = 0x2306, // \doublebarwedge
        [0x2A95] = 0x22DC, // \eqslantless
        [0x2A96] = 0x22DD, // \eqslantgtr
        [0x2AB0] = 0x227D, // \succeq
    };

    internal static void MergeMathTypeAlternates(XElement root, bool keepNativeTokens)
    {
        foreach (var token in root.DescendantsAndSelf().Where(e => e.Name.LocalName is "mi" or "mo"
                     && (string?)e.Attribute("data-mjx-alternate") != "1"
                     && e.Value.Trim().Length == 1
                     && MathTypeStoredCharacters.ContainsKey(e.Value.Trim()[0])
                     && !(keepNativeTokens && e.Attribute("data-mtef-mtcode") is not null)).ToArray())
            token.Value = ((char)MathTypeStoredCharacters[token.Value.Trim()[0]]).ToString();
        foreach (var token in root.DescendantsAndSelf().Where(e => e.Name.LocalName is "mi" or "mo"
                     && (string?)e.Attribute("data-mjx-alternate") == "1"
                     && e.Value.Trim().Length == 1
                     && MathTypeMergedAlternates.ContainsKey(e.Value.Trim()[0])
                     && !(keepNativeTokens && e.Attribute("data-mtef-mtcode") is not null)).ToArray())
        {
            var merged = MathTypeMergedAlternates[token.Value.Trim()[0]];
            token.Value = ((char)merged.Scalar).ToString();
            if (!merged.Alternate) token.Attribute("data-mjx-alternate")!.Remove();
        }
    }

    internal static NativeGlyphAuthority? ResolveNativeGlyphAuthority(XElement token)
    {
        if (token.Name.LocalName is not ("mo" or "mi") || token.Value.Length != 1)
            return null;
        // A glyph authority is one font position with a character style;
        // alphabets such as sans-serif keep their own font binding.
        if (MathMlTokenStyleContract.ResolveEffectiveVariant(token) is not
            (null or "" or "normal" or "bold" or "italic" or "bold-italic"))
            return null;
        var alternate = (string?)token.Attribute("data-mjx-alternate") == "1";
        return NativeGlyphAuthorities.FirstOrDefault(glyph =>
            glyph.Scalar == token.Value[0] && glyph.Alternate == alternate);
    }

    internal static NativeGlyphAuthority? ResolveNativeGlyphAuthority(
        string? font, string? encoding, int scalar, int encoded, bool hasMtCode) =>
        NativeGlyphAuthorities.FirstOrDefault(glyph =>
            string.Equals(font, glyph.Font, StringComparison.OrdinalIgnoreCase)
            && string.Equals(encoding, glyph.Encoding, StringComparison.OrdinalIgnoreCase)
            && (encoded < 0 || encoded == glyph.Encoded8)
            && (!hasMtCode || scalar == glyph.MtCode
                || scalar == glyph.Scalar && encoded == glyph.Encoded8));

    private const string HbarPlainGlyphMarker = "data-vtx-hbar-plain";

    // MT Extra's plain hbar is the slanted \hbar glyph itself: no style to add.
    internal static string? NativeGlyphCharacterVariant(byte style, NativeGlyphAuthority glyph) =>
        style switch
        {
            1 => "bold", 2 => "italic", 3 => "bold-italic",
            _ => glyph.Scalar == 0x210F ? null : glyph.TokenClass == "mi" ? "normal" : null,
        };

    internal static void EmitExplicitFontDefinitions(
        XElement math,
        byte[] sourceMtef,
        List<byte> output)
    {
        var definitionsStart = output.Count;
        var existing = CountAvailableDefinitions(sourceMtef);
        var nativeStyles = ReadNativeCharacterStyleBindings(sourceMtef);
        foreach (var token in math.DescendantsAndSelf().Where(e => !e.HasElements
                     && e.Name.LocalName is "mi" or "mo" or "mtext"
                     && e.Attribute("data-mtef-mtcode") is null
                     && e.Value.Any(FarEastGlyphSubstitutes.ContainsKey)).ToArray())
            token.Value = new string(token.Value
                .Select(c => FarEastGlyphSubstitutes.TryGetValue(c, out var covered) ? covered : c).ToArray());
        // A letterlike character (U+210B script H, U+212D fraktur C, ...) is its
        // base letter in that alphabet. Without the declared variant no alphabet
        // font binds below and MathType shows the plain letter.
        foreach (var token in math.DescendantsAndSelf().Where(e => e.Name.LocalName == "mi"
                     && e.Value.Trim().Length == 1
                     && e.Attribute("data-mtef-mtcode") is null
                     && MathMlTokenStyleContract.ResolveDeclaredVariant(e) is null
                     && TryNormalizeLetterlikeForWrite(e.Value.Trim()[0], out _, out _)).ToArray())
        {
            TryNormalizeLetterlikeForWrite(token.Value.Trim()[0], out var letter, out var letterVariant);
            token.Value = ((char)letter).ToString();
            token.SetAttributeValue("mathvariant", letterVariant);
        }
        MergeMathTypeAlternates(math, keepNativeTokens: true);
        // A bare hbar (Word's OMML gives <mi>ℏ</mi>) is MathJax's \hbar: the same
        // implicit-italic MT Extra glyph. Without MathJax's alternate marker it
        // was written in the upright built-in MT Extra face and read back upright.
        foreach (var token in math.DescendantsAndSelf().Where(e => e.Name.LocalName == "mi"
                     && e.Value.Trim() == "ℏ"
                     && e.Attribute("data-mtef-mtcode") is null
                     && e.Attribute("data-mjx-alternate") is null
                     && MathMlTokenStyleContract.ResolveDeclaredVariant(e) is null).ToArray())
            token.SetAttributeValue("data-mjx-alternate", "1");
        // MT Extra's hbar is drawn slanted in its plain style, which is how both
        // of MathType's importers store \hbar. Italic on top slants it twice, so
        // its implicit or explicit italic is that plain glyph (bold stays bold).
        foreach (var token in math.DescendantsAndSelf().Where(e => e.Name.LocalName == "mi"
                     && e.Value.Trim() == "ℏ").ToArray())
        {
            var declared = MathMlTokenStyleContract.ResolveDeclaredVariant(token);
            if (declared is "italic" or "normal") token.Attribute("mathvariant")?.Remove();
            else if (declared == "bold-italic") token.SetAttributeValue("mathvariant", "bold");
            if (MathMlTokenStyleContract.ResolveDeclaredVariant(token) is null)
                token.SetAttributeValue(HbarPlainGlyphMarker, "1");
        }
        // MTEF persists scalar CHARs. Split only the already-canonical equivalent
        // operator runs before binding fonts, so every glyph receives its style.
        foreach (var token in math.DescendantsAndSelf().Where(e => e.Name.LocalName == "mo"
                     && IsScalarOperatorRun(e.Value)).ToArray())
            token.ReplaceWith(EnumerateBmpScalars(token.Value).Select(scalar =>
                new XElement(token.Name, token.Attributes(), char.ConvertFromUtf32(scalar))));
        // An operator outside Times New Roman's code pages (e.g. U+25B3) has no
        // styled identifier form; bind it to SimSun/WinAllCodePages below.
        // An explicit normal variant (MathJax's \triangle is
        // <mi mathvariant="normal">) is the same upright SimSun glyph; in the
        // Times binding MathType draws its missing-glyph diamond instead.
        foreach (var token in math.DescendantsAndSelf().Where(e => e.Value.Trim().Length == 1
                     && (e.Name.LocalName == "mo"
                         || e.Name.LocalName == "mi" && IsMathematicalSymbolScalar(e.Value.Trim()[0]))
                     && IsFarEastCodePageOperator(e.Value.Trim()[0])
                     && MathMlTokenStyleContract.ResolveDeclaredVariant(e) is null or "normal"
                     && ResolveNativeGlyphAuthority(e) is null
                     && !HasValidPreservedExplicitTypeface(e, existing.FontStyleDefinitions, nativeStyles)).ToArray())
        {
            // A symbol mi keeps its implicit italic, as SemanticSignature does.
            var upright = token.Name.LocalName == "mo"
                || MathMlTokenStyleContract.ResolveDeclaredVariant(token) == "normal";
            token.SetAttributeValue(FarEastOperatorMarker, upright ? "normal" : "italic");
            token.Name = token.Name.Namespace + "mo";
            token.Attribute("mathvariant")?.Remove();
        }
        // A single mathematical-symbol mi inherits italic when no variant was
        // supplied. Inherited variants have already been materialized. Resolve
        // this before changing token class or binding a native font, so a plain
        // native glyph cannot masquerade as an implicit-italic MathML token.
        foreach (var token in math.DescendantsAndSelf().Where(e => e.Name.LocalName == "mi"
                     && e.Value.Trim().Length == 1
                     && (IsMathematicalSymbolToken(e.Value.Trim())
                         || ResolveNativeGlyphAuthority(e) is not null
                         || IsLowerGreek(e.Value.Trim()[0]) || IsUpperGreek(e.Value.Trim()[0]))
                     && MathMlTokenStyleContract.ResolveDeclaredVariant(e) is null))
        {
            if (TryNormalizeLetterlikeForWrite(token.Value.Trim()[0], out _, out _)) continue;
            if (token.Attribute(HbarPlainGlyphMarker) is not null) continue;
            var effective = MathMlTokenStyleContract.ResolveEffectiveVariant(token);
            var scalar = token.Value.Trim()[0];
            var greekFace = IsLowerGreek(scalar) ? TypefaceLowerGreek
                : IsUpperGreek(scalar) ? TypefaceUpperGreek : 0;
            if (greekFace != 0 && nativeStyles.TryGetValue(greekFace, out var binding)
                && TryNativeCharacterStyle(effective, out var requestedBits)
                && requestedBits == binding.CharacterStyle)
                continue;
            token.SetAttributeValue("mathvariant", effective);
        }
        foreach (var token in math.DescendantsAndSelf())
        {
            if (!int.TryParse((string?)token.Attribute("data-mtef-explicit-typeface"), out var face)
                || !nativeStyles.TryGetValue(face, out var actual)) continue;
            var requested = token.Attribute(HbarPlainGlyphMarker) is not null
                ? "normal" // the slanted MT Extra hbar in its plain style
                : MathMlTokenStyleContract.ResolveEffectiveVariant(token);
            if (TryNativeCharacterStyle(requested, out var bits) && bits != actual.CharacterStyle)
            {
                // A new font/style must not inherit an old font-position byte.
                foreach (var attribute in token.Attributes().Where(a =>
                             a.Name.LocalName.StartsWith("data-mtef-", StringComparison.Ordinal)).ToArray())
                    attribute.Remove();
            }
        }
        foreach (var token in math.DescendantsAndSelf().Where(e => e.Name.LocalName == "mi"
                     && ResolveNativeGlyphAuthority(e) is not null
                     && !HasValidPreservedExplicitTypeface(e, existing.FontStyleDefinitions, nativeStyles)))
            token.Name = token.Name.Namespace + ResolveNativeGlyphAuthority(token)!.TokenClass;
        // A mathematical glyph has the same native character mapping whether
        // a frontend calls it mi or mo. Keep explicit styles on that mapping.
        foreach (var token in math.DescendantsAndSelf().Where(e => e.Name.LocalName == "mi"
                     && IsMathematicalSymbolToken(e.Value.Trim())
                     && ResolveNativeGlyphAuthority(e)?.TokenClass != "mi"
                     && ((string?)e.Attribute("mathvariant") ?? string.Empty) is "bold" or "italic" or "bold-italic"))
            token.Name = token.Name.Namespace + "mo";
        var specs = new[]
        {
            (Variant: "normal", Encoding: "WindowsANSI", Font: "Times New Roman", CharacterStyle: (byte)0),
            (Variant: "bold", Encoding: "WindowsANSI", Font: "Times New Roman", CharacterStyle: (byte)1),
            (Variant: "italic", Encoding: "WindowsANSI", Font: "Times New Roman", CharacterStyle: (byte)2),
            (Variant: "script", Encoding: "EuclidMath1", Font: "Euclid Math One", CharacterStyle: (byte)0),
            (Variant: "bold-script", Encoding: "EuclidMath1", Font: "Euclid Math One", CharacterStyle: (byte)1),
            (Variant: "double-struck", Encoding: "EuclidMath2", Font: "Euclid Math Two", CharacterStyle: (byte)0),
            (Variant: "fraktur", Encoding: "EuclidFraktur", Font: "Euclid Fraktur", CharacterStyle: (byte)0),
            (Variant: "bold-fraktur", Encoding: "EuclidFraktur", Font: "Euclid Fraktur", CharacterStyle: (byte)1),
            (Variant: "sans-serif", Encoding: "WindowsANSI", Font: "Arial", CharacterStyle: (byte)0),
            (Variant: "bold-sans-serif", Encoding: "WindowsANSI", Font: "Arial", CharacterStyle: (byte)1),
            (Variant: "sans-serif-italic", Encoding: "WindowsANSI", Font: "Arial", CharacterStyle: (byte)2),
            (Variant: "sans-serif-bold-italic", Encoding: "WindowsANSI", Font: "Arial", CharacterStyle: (byte)3),
            (Variant: "monospace", Encoding: "WindowsANSI", Font: "Courier New", CharacterStyle: (byte)0),
            (Variant: "bold-italic", Encoding: "WindowsANSI", Font: "Times New Roman", CharacterStyle: (byte)3),
            (Variant: "native-glyph", Encoding: "Symbol", Font: "Symbol", CharacterStyle: (byte)0),
            (Variant: "native-glyph", Encoding: "MTExtra", Font: "MT Extra", CharacterStyle: (byte)0),
            // Font is the prefix's own typeface-12 FONT_DEF (SimSun under its
            // localized name); MathType rejects the English name "SimSun".
            (Variant: "far-east-operator", Encoding: "WinAllCodePages", Font: "", CharacterStyle: (byte)0),
            (Variant: "far-east-operator-italic", Encoding: "WinAllCodePages", Font: "", CharacterStyle: (byte)2),
        };
        var createdEncodings = 0;
        var createdFonts = 0;
        var createdStyles = 0;
        foreach (var spec in specs)
        {
            var tokens = math.DescendantsAndSelf()
                .Where(element => !IsVectorStyleLetter(element)
                    && !(element.Attribute(HbarPlainGlyphMarker) is not null
                        && element.Attribute("data-mtef-explicit-typeface") is null))
                .Where(element => ResolveNativeGlyphAuthority(element) is { } glyph
                    ? spec.CharacterStyle == 0 && glyph.Font == spec.Font && glyph.Encoding == spec.Encoding
                    : MatchesExplicitFontSpec(element, spec.Variant))
                .Where(element => !HasValidPreservedExplicitTypeface(
                    element,
                    existing.FontStyleDefinitions, nativeStyles))
                .ToArray();
            if (tokens.Length == 0) continue;

            if (spec.Variant.StartsWith("far-east-operator", StringComparison.Ordinal))
            {
                // Without a typeface-12 font the operator keeps its old Times binding.
                if (!nativeStyles.TryGetValue(TypefaceFarEastText, out var farEast) || farEast.FontDefinition <= 0)
                    continue;
                var farEastStyle = FindPrefixFontStyleIndex(sourceMtef, farEast.FontDefinition, spec.CharacterStyle);
                if (farEastStyle is null)
                {
                    farEastStyle = existing.FontStyleDefinitions + ++createdStyles;
                    output.Add(RecordFontStyleDef);
                    WriteUnsigned(output, farEast.FontDefinition);
                    output.Add(spec.CharacterStyle);
                }
                foreach (var token in tokens)
                    token.SetAttributeValue("data-mtef-explicit-typeface", -farEastStyle.Value);
                continue;
            }

            var existingStyleIndex = FindPrefixFontStyleIndex(
                sourceMtef,
                spec.Font,
                spec.CharacterStyle);
            if (existingStyleIndex is not null)
            {
                foreach (var token in tokens)
                {
                    token.SetAttributeValue("data-mtef-explicit-typeface", -existingStyleIndex.Value);
                    token.SetAttributeValue("data-mtef-explicit-font", spec.Font);
                }
                continue;
            }

            var encodingIndex = 4 + existing.EncodingDefinitions + ++createdEncodings;
            var fontDefinitionIndex = existing.FontDefinitions + ++createdFonts;
            var fontStyleIndex = existing.FontStyleDefinitions + ++createdStyles;

            output.Add(RecordEncodingDef);
            WriteNullTerminatedAscii(output, spec.Encoding);
            output.Add(RecordFontDef);
            WriteUnsigned(output, encodingIndex);
            WriteNullTerminatedAscii(output, spec.Font);
            output.Add(RecordFontStyleDef);
            WriteUnsigned(output, fontDefinitionIndex);
            output.Add(spec.CharacterStyle);

            foreach (var token in tokens)
            {
                token.SetAttributeValue("data-mtef-explicit-typeface", -fontStyleIndex);
                token.SetAttributeValue("data-mtef-explicit-font", spec.Font);
            }
        }
        MaterializeExplicitCharacterStyles(math, sourceMtef, output, definitionsStart);
    }

    // Resolve font/style through the native glyph's actual encoding. A bold
    // symbol is NOT an identifier in the Vector face, nor a Unicode scalar in
    // an arbitrary Times font. Retain the canonical CHAR's MTCode/font-position
    // and bind a FONT_STYLE_DEF to the SAME FONT_DEF with the requested bits.
    // This applies to Symbol, MT Extra, Euclid and text fallback characters.
    private static void MaterializeExplicitCharacterStyles(
        XElement math, byte[] sourceMtef, List<byte> output, int definitionsStart)
    {
        // FUNCTION APPLICATION (U+2061) is stored as no CHAR at all, so a style
        // on it has nothing to bind to and no visible effect.
        var tokens = math.DescendantsAndSelf().Where(e => (e.Name.LocalName is "mo" or "mi")
            && e.Value.Length == 1
            && e.Value[0] != (char)0x2061
            && !IsVectorStyleLetter(e)
            && (((string?)e.Attribute("mathvariant") ?? string.Empty) is "bold" or "italic" or "bold-italic"
                || TryGetBasicStyledLegacyIdentifierGlyph(e, out _, out _, out _))).ToArray();
        if (tokens.Length == 0) return;
        var fonts = new List<string>();
        var styles = new List<(int Font, byte Bits)>();
        var builtin = ReadNativeCharacterStyleBindings(sourceMtef)
            .Where(item => item.Key > 0).ToDictionary(item => item.Key, item => item.Value.FontDefinition);
        void Definition(byte[] data, ref int position)
        {
            var record = data[position++];
            if (record == RecordEncodingDef) position = SkipNullTerminated(data, position);
            else if (record == RecordFontDef)
            {
                _ = ReadUnsigned(data, ref position);
                fonts.Add(ReadPrefixNullTerminatedString(data, ref position));
            }
            else if (record == RecordFontStyleDef)
            {
                var index = ReadUnsigned(data, ref position);
                Require(data,position,1); styles.Add((index,data[position++]));
            }
            else throw new InvalidDataException("Unexpected record in native character font definitions.");
        }
        foreach(var record in ReadPrefixLayout(sourceMtef).Records.Concat(ReadRootLeadingFormattingRecords(sourceMtef)))
        {
            var cursor=record.Start;
            if(record.Type is RecordEncodingDef or RecordFontDef or RecordFontStyleDef)
                Definition(sourceMtef,ref cursor);

        }
        var added=output.Skip(definitionsStart).ToArray();
        for(var cursor=0;cursor<added.Length;) Definition(added,ref cursor);
        int StyledTypeface(int typeface, byte bits)
        {
            var font=0;
            if(typeface<0 && -typeface<=styles.Count) font=styles[-typeface-1].Font;
            else builtin.TryGetValue(typeface,out font);
            if(font==0)
            {
                var family=typeface is TypefaceSymbol or TypefaceLowerGreek or TypefaceUpperGreek ? "Symbol"
                    : typeface==TypefaceMtExtra ? "MT Extra" : "Times New Roman";
                font=fonts.FindIndex(name=>string.Equals(name,family,StringComparison.OrdinalIgnoreCase))+1;
            }
            if(font<1 || font>fonts.Count) throw new InvalidDataException("No declared font owns the styled native operator.");
            var style=styles.FindIndex(s=>s.Font==font && s.Bits==bits)+1;
            if(style==0)
            {
                output.Add(RecordFontStyleDef); WriteUnsigned(output,font); output.Add(bits);
                styles.Add((font,bits)); style=styles.Count;
            }
            return -style;
        }
        // A large-operator template can use a different encoded glyph than
        // its fixed CHAR form. Bind the template glyph to its own native
        // font instead of reusing the ordinary character's font position.
        void BindTemplateTypefaces(XElement token, byte bits)
        {
            if(TryResolveNativeBigOperator(token,out var big,allowAmbiguousUnionIntersection:true))
                token.SetAttributeValue("data-mtef-template-typeface",StyledTypeface(big.Typeface,bits));
            if(TryResolveNativeIntegral(token,out _))
            {
                token.SetAttributeValue("data-mtef-template-typeface",StyledTypeface(TypefaceSymbol,bits));
                token.SetAttributeValue("data-mtef-adornment-typeface",StyledTypeface(TypefaceMtExtra,bits));
            }
        }
        foreach(var token in tokens)
        {
            var variant=MathMlTokenStyleContract.ResolveEffectiveVariant(token)!;
            var bits=(byte)(variant=="bold" ? 1 : variant=="italic" ? 2 : 3);
            // A decoded native CHAR already owns its exact explicit style.
            // Preserve its encoded8/encoded16/no-MTCode contract byte-for-byte;
            // a lone operator read back from its CHAR still needs the style
            // of the template glyph it becomes once limits are added.
            if (HasValidPreservedExplicitTypeface(token, styles.Count, ReadNativeCharacterStyleBindings(sourceMtef))
                && int.TryParse((string?)token.Attribute("data-mtef-explicit-typeface"), out var retained)
                && (retained > 0 || styles[-retained - 1].Bits == bits))
            {
                BindTemplateTypefaces(token, bits);
                continue;
            }
            var character=new List<byte>();
            // Bind style to the exact CHAR selected by this token's encoder.
            // Identifier glyphs must not bypass explicit style merely because
            // their native glyph is supplied by a legacy font map.
            if (TryGetBasicStyledLegacyIdentifierGlyph(token, out var mappedCode, out var mappedFace, out var mappedPosition))
                EmitScalar(mappedCode, mappedFace, character, includeEncoded8: true, encoded8Override: mappedPosition);
            else if (token.Name.LocalName == "mi") EmitIdentifier(token,character);
            else EmitOperator(token,character);
            if(character.Count<5 || character[0]!=RecordChar || (character[1] & 0x29)!=0)
                throw new InvalidDataException("Styled native operator must resolve to one undecorated CHAR.");
            var cursor=2;
            var lead=character[cursor++];
            var typeface=lead==255 ? (character[cursor++] | character[cursor++]<<8)-32768 : lead-128;
            var scalar=character[cursor++] | character[cursor++]<<8;
            byte? encoded=(character[1]&CharEncoded8)!=0 ? character[cursor++] : null;
            if(cursor!=character.Count) throw new InvalidDataException("Styled operator CHAR has unsupported trailing data.");
            var styled=StyledTypeface(typeface,bits);
            token.SetAttributeValue("data-mtef-explicit-typeface",styled);
            token.SetAttributeValue("data-mtef-mtcode",scalar);
            token.SetAttributeValue("data-mtef-encoded8",encoded);
            token.SetAttributeValue("data-mtef-explicit-font",fonts[styles[-styled-1].Font-1]);
            BindTemplateTypefaces(token, bits);
        }
    }

    // MathType's Vector style is its bold upright letter style: its import
    // stores \mathbf on Latin and Greek letters there (Times New Roman bold),
    // not in an explicit font style. A natively read style stays as it is.
    private static bool IsVectorStyleLetter(XElement token) =>
        token.Name.LocalName == "mi"
        && string.Equals(((string?)token.Attribute("mathvariant"))?.Trim(), "bold", StringComparison.Ordinal)
        && token.Attribute("data-mtef-explicit-typeface") is null
        && token.Value.Length == 1
        && (token.Value[0] is >= 'A' and <= 'Z' or >= 'a' and <= 'z'
            || token.Value[0] >= 0x0391 && token.Value[0] <= 0x03A9 && token.Value[0] != 0x03A2 // Greek capitals
            || token.Value[0] >= 0x03B1 && token.Value[0] <= 0x03C9); // Greek small letters

    internal static string NativeCharacterVariant(byte style) => style switch
    { 1 => "bold", 2 => "italic", 3 => "bold-italic", _ => "normal" };

    internal static string NativeTextFontVariant(string font, byte style) =>
        string.Equals(font, "Courier New", StringComparison.OrdinalIgnoreCase) ? "monospace"
        : string.Equals(font, "Arial", StringComparison.OrdinalIgnoreCase) ? style switch
        { 1 => "bold-sans-serif", 2 => "sans-serif-italic", 3 => "sans-serif-bold-italic", _ => "sans-serif" }
        : NativeCharacterVariant(style);

    internal static bool TryNativeCharacterStyle(string? variant, out byte style)
    {
        style = variant switch
        { "normal" or "upright" => 0, "bold" => 1, "italic" => 2, "bold-italic" => 3, _ => byte.MaxValue };
        return style != byte.MaxValue;
    }

    internal static Dictionary<int, (int FontDefinition, byte CharacterStyle)> ReadNativeCharacterStyleBindings(byte[] data)
    {
        var result = new Dictionary<int, (int FontDefinition, byte CharacterStyle)>();
        var explicitStyle = 0;
        foreach (var record in ReadPrefixLayout(data).Records.Concat(ReadRootLeadingFormattingRecords(data)))
        {
            var cursor = record.Start + 1;
            if (record.Type == RecordFontStyleDef)
            {
                var font = ReadUnsigned(data, ref cursor);
                Require(data, cursor, 1);
                result[-++explicitStyle] = (font, data[cursor]);
            }
            else if (record.Type == RecordEqnPrefs)
            {
                cursor++;
                cursor = SkipDimensionArray(data, cursor);
                cursor = SkipDimensionArray(data, cursor);
                Require(data, cursor, 1);
                var count = data[cursor++];
                for (var face = 1; face <= count; face++)
                {
                    var font = ReadUnsigned(data, ref cursor);
                    if (font == 0) continue;
                    Require(data, cursor, 1);
                    result[face] = (font, data[cursor++]);
                }
            }
        }
        return result;
    }

    private static bool HasValidPreservedExplicitTypeface(
        XElement element,
        int availableFontStyleDefinitions,
        IReadOnlyDictionary<int, (int FontDefinition, byte CharacterStyle)> nativeStyles)
    {
        return int.TryParse(
                (string?)element.Attribute("data-mtef-explicit-typeface"),
                NumberStyles.Integer,
                CultureInfo.InvariantCulture,
                out var typeface)
            && (typeface < 0 && -typeface <= availableFontStyleDefinitions
                || typeface > 0 && nativeStyles.TryGetValue(typeface, out var nativeStyle)
                    && TryNativeCharacterStyle(element.Attribute(HbarPlainGlyphMarker) is not null
                        ? "normal" // the slanted MT Extra hbar in its plain style
                        : MathMlTokenStyleContract.ResolveEffectiveVariant(element), out var bits)
                    && bits == nativeStyle.CharacterStyle)
            && (int.TryParse(
                (string?)element.Attribute("data-mtef-mtcode"),
                NumberStyles.Integer,
                CultureInfo.InvariantCulture,
                out _)
                || IsTrueMathMlBoolean((string?)element.Attribute("data-mtef-no-mtcode")));
    }

    private static bool TryGetBasicStyledLegacyIdentifierGlyph(
        XElement token, out int mtCode, out int typeface, out byte encoded8)
    {
        mtCode = typeface = 0; encoded8 = 0;
        if (token.Name.LocalName != "mi" || token.HasElements || token.Value.Length != 1
            || MathMlTokenStyleContract.ResolveEffectiveVariant(token) is not ("bold" or "italic" or "bold-italic")
            || (string?)token.Attribute("data-mjx-alternate") == "1"
            || ResolveNativeGlyphAuthority(token) is not null)
            return false;
        var scalar = token.Value[0];
        // Alphabet-specific Greek and explicit alternate glyphs keep their
        // established authorities. This is the existing native symbol tuple,
        // selected before adding weight/slant, not a Unicode-font fallback.
        if (TryEuclidMathOneGreekVariantEncoded8(scalar, out _)
            || !(scalar > 0x7F || IsMathematicalSymbolScalar(scalar)))
            return false;
        return TryResolveLegacyMathTypeGlyph(scalar, out mtCode, out typeface, out encoded8)
            && typeface is TypefaceSymbol or TypefaceMtExtra;
    }

    private static bool MatchesExplicitFontSpec(XElement element, string variant)
    {
        if (variant.StartsWith("far-east-operator", StringComparison.Ordinal)
            || element.Attribute(FarEastOperatorMarker) is not null)
            return (string?)element.Attribute(FarEastOperatorMarker) switch
            {
                "normal" => variant == "far-east-operator",
                "italic" => variant == "far-east-operator-italic",
                _ => false,
            };
        if (TryGetBasicStyledLegacyIdentifierGlyph(element, out _, out _, out _)) return false;
        var local = element.Name.LocalName;
        var scalars = EnumerateBmpScalars(element.Value).ToArray();
        if (scalars.Length == 0) return false;
        if (local == "mo" && scalars.Length > 1 && scalars.All(IsLetterScalar)
            && variant is "bold" or "italic" or "bold-italic")
            return MathMlTokenStyleContract.ResolveEffectiveVariant(element) == variant;

        if (variant == "normal")
        {
            // Explicit upright alphabetic identifiers need their own font style:
            // the built-in Text face reads back as prose and merges neighboring
            // identifiers across math spaces. Greek faces also inherit slant
            // from preferences. Bind Latin/Greek identifiers to the existing
            // Times New Roman contract with style bits zero; keep prose, mixed
            // literals, default Greek and symbol/operator mappings unchanged.
            return local == "mi"
                && string.Equals(((string?)element.Attribute("mathvariant"))?.Trim(),
                    "normal", StringComparison.OrdinalIgnoreCase)
                && scalars.All(scalar => scalar <= 0x7F && IsLetterScalar(scalar)
                    || IsLowerGreek(scalar) || IsUpperGreek(scalar)
                    || TryEuclidMathOneGreekVariantEncoded8(scalar, out _));
        }
        if (local == "mn")
            return string.Equals(((string?)element.Attribute("mathvariant"))?.Trim(), variant, StringComparison.OrdinalIgnoreCase);
        // Built-in Greek faces inherit equation preferences: UpperGreek is
        // commonly upright, so it cannot encode an explicit italic request.
        // Use the same explicit Times font/style contract as upright Greek.
        // Leave unstyled Greek and ordinary Latin identifier preferences alone.
        if (variant is "bold" or "italic")
            return variant == "italic" && local == "mi"
                && string.Equals(((string?)element.Attribute("mathvariant"))?.Trim(),
                    variant, StringComparison.OrdinalIgnoreCase)
                && scalars.All(scalar => IsLowerGreek(scalar) || IsUpperGreek(scalar));
        if (local == "mi")
        {
            if (string.Equals(
                    ((string?)element.Attribute("mathvariant"))?.Trim(),
                    variant,
                    StringComparison.OrdinalIgnoreCase))
            {
                // The standard blackboard set letters use MathType's built-in
                // MT Extra style, so they do not require an explicit EuclidMath2
                // prefix even when MathML calls them double-struck.
                return !string.Equals(variant, "double-struck", StringComparison.OrdinalIgnoreCase)
                    || !scalars.All(scalar => TryMtExtraDoubleStruck(scalar, out _, out _));
            }

            if (string.Equals(variant, "script", StringComparison.OrdinalIgnoreCase)
                && (string.IsNullOrWhiteSpace((string?)element.Attribute("mathvariant"))
                    || string.Equals((string?)element.Attribute("mathvariant"), "italic", StringComparison.OrdinalIgnoreCase))
                && (scalars.All(scalar => TryEuclidMathOneGreekVariantEncoded8(scalar, out _))
                    || scalars.All(scalar => TryEuclidMathOneDotlessJ(scalar, out _, out _))))
            {
                // MathType 7 stores TeX \epsilon / \varrho / \varkappa and
                // dotless-j in Euclid Math One even when MathML does not carry
                // a script mathvariant.
                return true;
            }
        }

        if (local == "mo")
        {
            if (string.Equals(variant, "script", StringComparison.OrdinalIgnoreCase)
                && scalars.All(scalar => TryEuclidMathOneOperatorEncoded8(scalar, out _)))
            {
                return true;
            }
            if (string.Equals(variant, "double-struck", StringComparison.OrdinalIgnoreCase)
                && scalars.All(scalar => TryEuclidMathTwoOperatorEncoded8(scalar, out _)))
            {
                return true;
            }
        }

        return false;
    }

    private static int? FindPrefixFontStyleIndex(
        byte[] mtef, int fontDefinition, byte expectedCharacterStyle)
    {
        var styleIndex = 0;
        foreach (var record in ReadPrefixLayout(mtef).Records.Where(r => r.Type == RecordFontStyleDef))
        {
            var position = record.Start + 1;
            styleIndex++;
            if (ReadUnsigned(mtef, ref position) == fontDefinition && mtef[position] == expectedCharacterStyle)
                return styleIndex;
        }
        return null;
    }

    private static int? FindPrefixFontStyleIndex(
        byte[] mtef, string fontName, byte expectedCharacterStyle)
    {
        var fontDefinitions = new List<string>();
        var styleIndex = 0;
        foreach (var record in ReadPrefixLayout(mtef).Records)
        {
            var position = record.Start + 1;
            if (record.Type == RecordFontDef)
            {
                _ = ReadUnsigned(mtef, ref position);
                fontDefinitions.Add(ReadPrefixNullTerminatedString(mtef, ref position));
            }
            else if (record.Type == RecordFontStyleDef)
            {
                var fontDefinitionIndex = ReadUnsigned(mtef, ref position);
                var actualCharacterStyle = mtef[position];
                styleIndex++;
                if (actualCharacterStyle == expectedCharacterStyle
                    && fontDefinitionIndex > 0 && fontDefinitionIndex <= fontDefinitions.Count
                    && string.Equals(fontDefinitions[fontDefinitionIndex - 1], fontName, StringComparison.OrdinalIgnoreCase))
                    return styleIndex;
            }
        }
        return null;
    }

    internal static string ReadPrefixNullTerminatedString(
        byte[] data,
        ref int position)
    {
        var start = position;
        while (position < data.Length && data[position] != 0) position++;
        if (position >= data.Length)
            throw new EndOfStreamException("MTEF prefix string is not null terminated.");
        var value = System.Text.Encoding.Default.GetString(
            data,
            start,
            position - start);
        position++;
        return value;
    }

    private static (int EncodingDefinitions, int FontDefinitions, int FontStyleDefinitions)
        CountAvailableDefinitions(byte[] mtef)
    {
        var records = ReadPrefixLayout(mtef).Records
            .Concat(ReadRootLeadingFormattingRecords(mtef))
            .ToArray();
        return (records.Count(item => item.Type == RecordEncodingDef),
            records.Count(item => item.Type == RecordFontDef),
            records.Count(item => item.Type == RecordFontStyleDef));
    }

    internal static void WriteNullTerminatedAscii(List<byte> output, string value)
    {
        output.AddRange(System.Text.Encoding.ASCII.GetBytes(value));
        output.Add(0);
    }
}
