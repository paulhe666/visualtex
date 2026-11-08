//! The fonts dvisvgm embeds, made right for the browser:
//! - dvisvgm (3.2 at least) sometimes writes a TrueType font whose `head` says long `loca` offsets while the
//!   `loca` table holds short ones; browsers then reject the font (OTS: "loca: Failed to read offset") and draw
//!   the page's text in a fallback font. The flag is set to what the table holds.
//! - CJK fonts map a Kangxi radical (U+2F45 ⽅) and the ideograph (U+65B9 方) to one glyph; dvisvgm writes the
//!   radical, so the page's text copies and searches as radicals. The ideograph is added to the font's cmap and
//!   the text written with it.

use std::io::Read;

use base64::engine::general_purpose::STANDARD as BASE64;
use base64::Engine;

fn be16(b: &[u8], i: usize) -> Option<u16> {
    Some(u16::from_be_bytes([*b.get(i)?, *b.get(i + 1)?]))
}
fn be32(b: &[u8], i: usize) -> Option<u32> {
    Some(u32::from_be_bytes([*b.get(i)?, *b.get(i + 1)?, *b.get(i + 2)?, *b.get(i + 3)?]))
}

/// The tables of a WOFF 1 or sfnt font: (tag, data).
fn tables(font: &[u8]) -> Option<(u32, Vec<([u8; 4], Vec<u8>)>)> {
    let mut out = vec![];
    if font.starts_with(b"wOFF") {
        let flavor = be32(font, 4)?;
        let n = be16(font, 12)? as usize;
        for i in 0..n {
            let e = 44 + 20 * i;
            let tag = [font[e], font[e + 1], font[e + 2], font[e + 3]];
            let (off, comp, orig) = (be32(font, e + 4)? as usize, be32(font, e + 8)? as usize, be32(font, e + 12)? as usize);
            let raw = font.get(off..off + comp)?;
            let data = if comp < orig {
                let mut d = Vec::with_capacity(orig);
                flate2::read::ZlibDecoder::new(raw).read_to_end(&mut d).ok()?;
                d
            } else {
                raw.to_vec()
            };
            out.push((tag, data));
        }
        Some((flavor, out))
    } else {
        let flavor = be32(font, 0)?;
        let n = be16(font, 4)? as usize;
        for i in 0..n {
            let e = 12 + 16 * i;
            let tag = [font[e], font[e + 1], font[e + 2], font[e + 3]];
            let (off, len) = (be32(font, e + 8)? as usize, be32(font, e + 12)? as usize);
            out.push((tag, font.get(off..off + len)?.to_vec()));
        }
        Some((flavor, out))
    }
}

fn checksum(data: &[u8]) -> u32 {
    let mut sum: u32 = 0;
    for chunk in data.chunks(4) {
        let mut w = [0u8; 4];
        w[..chunk.len()].copy_from_slice(chunk);
        sum = sum.wrapping_add(u32::from_be_bytes(w));
    }
    sum
}

fn sfnt(flavor: u32, tables: &[([u8; 4], Vec<u8>)]) -> Vec<u8> {
    let n = tables.len() as u16;
    let mut pow = 1u16;
    let mut log = 0u16;
    while pow * 2 <= n {
        pow *= 2;
        log += 1;
    }
    let mut out = vec![];
    out.extend_from_slice(&flavor.to_be_bytes());
    out.extend_from_slice(&n.to_be_bytes());
    out.extend_from_slice(&(pow * 16).to_be_bytes());
    out.extend_from_slice(&log.to_be_bytes());
    out.extend_from_slice(&(n * 16 - pow * 16).to_be_bytes());
    let mut offset = 12 + 16 * tables.len();
    let mut body = vec![];
    for (tag, data) in tables {
        out.extend_from_slice(tag);
        out.extend_from_slice(&checksum(data).to_be_bytes());
        out.extend_from_slice(&(offset as u32).to_be_bytes());
        out.extend_from_slice(&(data.len() as u32).to_be_bytes());
        body.extend_from_slice(data);
        while body.len() % 4 != 0 {
            body.push(0);
        }
        offset = 12 + 16 * tables.len() + body.len();
    }
    out.extend_from_slice(&body);
    out
}

fn parse_cmap(cmap: &[u8]) -> Option<std::collections::BTreeMap<u32, u16>> {
    let n = be16(cmap, 2)? as usize;
    let mut best: Option<(u16, usize)> = None;
    for i in 0..n {
        let e = 4 + 8 * i;
        let (platform, encoding, offset) = (be16(cmap, e)?, be16(cmap, e + 2)?, be32(cmap, e + 4)? as usize);
        let format = be16(cmap, offset)?;
        let rank = match (platform, encoding, format) {
            (3, 10, 12) | (0, 4, 12) | (0, 6, 12) => 3,
            (3, 1, 4) | (0, 3, 4) | (0, _, 4) => 2,
            _ => 0,
        };
        if rank > 0 && best.map(|(r, _)| rank > r).unwrap_or(true) {
            best = Some((rank, offset));
        }
    }
    let (_, off) = best?;
    let t = &cmap[off..];
    let mut map = std::collections::BTreeMap::new();
    match be16(t, 0)? {
        4 => {
            let seg = be16(t, 6)? as usize / 2;
            let (ends, starts, deltas, ranges) = (14, 16 + 2 * seg, 16 + 4 * seg, 16 + 6 * seg);
            for i in 0..seg {
                let (end, start) = (be16(t, ends + 2 * i)? as u32, be16(t, starts + 2 * i)? as u32);
                let delta = be16(t, deltas + 2 * i)?;
                let range = be16(t, ranges + 2 * i)? as usize;
                if start == 0xFFFF {
                    continue;
                }
                for c in start..=end {
                    let g = if range == 0 {
                        (c as u16).wrapping_add(delta)
                    } else {
                        let at = ranges + 2 * i + range + 2 * (c - start) as usize;
                        let g = be16(t, at)?;
                        if g == 0 {
                            0
                        } else {
                            g.wrapping_add(delta)
                        }
                    };
                    if g != 0 {
                        map.insert(c, g);
                    }
                }
            }
        }
        12 => {
            let groups = be32(t, 12)? as usize;
            for i in 0..groups {
                let e = 16 + 12 * i;
                let (a, b, g) = (be32(t, e)?, be32(t, e + 4)?, be32(t, e + 8)?);
                for c in a..=b.min(a + 0xFFFF) {
                    map.insert(c, (g + c - a) as u16);
                }
            }
        }
        _ => return None,
    }
    Some(map)
}

/// A cmap with a (3,1) format 4 and a (3,10) format 12 subtable for `map`.
fn write_cmap(map: &std::collections::BTreeMap<u32, u16>) -> Vec<u8> {
    let bmp: Vec<(u32, u16)> = map.iter().filter(|(c, _)| **c < 0xFFFF).map(|(c, g)| (*c, *g)).collect();
    let seg = bmp.len() + 1;
    let mut f4 = vec![];
    let push16 = |v: &mut Vec<u8>, x: u16| v.extend_from_slice(&x.to_be_bytes());
    let length = 16 + 8 * seg;
    push16(&mut f4, 4);
    push16(&mut f4, length as u16);
    push16(&mut f4, 0);
    push16(&mut f4, (seg * 2) as u16);
    let mut pow = 1usize;
    let mut log = 0u16;
    while pow * 2 <= seg {
        pow *= 2;
        log += 1;
    }
    push16(&mut f4, (pow * 2) as u16);
    push16(&mut f4, log);
    push16(&mut f4, (seg * 2 - pow * 2) as u16);
    for (c, _) in &bmp {
        push16(&mut f4, *c as u16);
    }
    push16(&mut f4, 0xFFFF);
    push16(&mut f4, 0);
    for (c, _) in &bmp {
        push16(&mut f4, *c as u16);
    }
    push16(&mut f4, 0xFFFF);
    for (c, g) in &bmp {
        push16(&mut f4, g.wrapping_sub(*c as u16));
    }
    push16(&mut f4, 1);
    for _ in 0..seg {
        push16(&mut f4, 0);
    }
    let mut f12 = vec![];
    push16(&mut f12, 12);
    push16(&mut f12, 0);
    f12.extend_from_slice(&((16 + 12 * map.len()) as u32).to_be_bytes());
    f12.extend_from_slice(&0u32.to_be_bytes());
    f12.extend_from_slice(&(map.len() as u32).to_be_bytes());
    for (c, g) in map {
        f12.extend_from_slice(&c.to_be_bytes());
        f12.extend_from_slice(&c.to_be_bytes());
        f12.extend_from_slice(&(*g as u32).to_be_bytes());
    }
    let mut out = vec![];
    push16(&mut out, 0);
    push16(&mut out, 2);
    for (enc, off) in [(1u16, 20u32), (10u16, 20 + f4.len() as u32)] {
        push16(&mut out, 3);
        push16(&mut out, enc);
        out.extend_from_slice(&off.to_be_bytes());
    }
    out.extend_from_slice(&f4);
    out.extend_from_slice(&f12);
    out
}

fn unified(c: u32) -> Option<u32> {
    crate::cjk_table::CJK_UNIFIED.binary_search_by_key(&c, |(a, _)| *a).ok().map(|i| crate::cjk_table::CJK_UNIFIED[i].1)
}

/// The font repaired as an sfnt, or None when it is fine (or not one this can read).
pub fn repair(font: &[u8]) -> Option<Vec<u8>> {
    let (flavor, mut tabs) = tables(font)?;
    let find = |tabs: &[([u8; 4], Vec<u8>)], t: &[u8; 4]| tabs.iter().position(|(tag, _)| tag == t);
    let mut changed = false;
    if let (Some(h), Some(l), Some(m)) = (find(&tabs, b"head"), find(&tabs, b"loca"), find(&tabs, b"maxp")) {
        let glyphs = be16(&tabs[m].1, 4)? as usize;
        let loca_len = tabs[l].1.len();
        let format = be16(&tabs[h].1, 50)?;
        let wanted: Option<u16> = if loca_len == (glyphs + 1) * 2 {
            Some(0)
        } else if loca_len == (glyphs + 1) * 4 {
            Some(1)
        } else {
            None
        };
        if let Some(w) = wanted {
            if w != format {
                let head = &mut tabs[h].1;
                head[50..52].copy_from_slice(&w.to_be_bytes());
                changed = true;
            }
        }
    }
    if let Some(c) = find(&tabs, b"cmap") {
        if let Some(mut map) = parse_cmap(&tabs[c].1) {
            let extra: Vec<(u32, u16)> = map.iter().filter_map(|(cp, g)| unified(*cp).filter(|u| !map.contains_key(u)).map(|u| (u, *g))).collect();
            if !extra.is_empty() {
                map.extend(extra);
                tabs[c].1 = write_cmap(&map);
                changed = true;
            }
        }
    }
    if !changed {
        return None;
    }
    if let Some(h) = find(&tabs, b"head") {
        tabs[h].1[8..12].copy_from_slice(&[0, 0, 0, 0]);
    }
    Some(sfnt(flavor, &tabs))
}

/// Every embedded font of a dvisvgm SVG checked; broken ones replaced by their repaired form; the text's CJK
/// radicals written as the ideographs.
pub fn repair_svg_fonts(svg: &str) -> String {
    let fonts = repair_fonts(svg);
    if !fonts.chars().any(|c| unified(c as u32).is_some()) {
        return fonts;
    }
    fonts.chars().map(|c| unified(c as u32).and_then(char::from_u32).unwrap_or(c)).collect()
}

fn repair_fonts(svg: &str) -> String {
    let key = "src:url(data:application/x-font-";
    if !svg.contains(key) {
        return svg.to_string();
    }
    let mut out = String::with_capacity(svg.len());
    let mut rest = svg;
    while let Some(i) = rest.find(key) {
        out.push_str(&rest[..i]);
        rest = &rest[i..];
        let Some(comma) = rest.find(";base64,") else { break };
        let Some(end) = rest[comma..].find(')') else { break };
        let data = &rest[comma + 8..comma + end];
        let fixed = BASE64.decode(data).ok().and_then(|bytes| repair(&bytes));
        match fixed {
            Some(font) => {
                out.push_str("src:url(data:font/ttf;base64,");
                out.push_str(&BASE64.encode(font));
                out.push(')');
                rest = &rest[comma + end + 1..];
                // the format() hint that followed no longer applies
                if let Some(r) = rest.strip_prefix(" format('woff')").or_else(|| rest.strip_prefix(" format('truetype')")) {
                    rest = r;
                }
            }
            None => {
                out.push_str(&rest[..comma + end + 1]);
                rest = &rest[comma + end + 1..];
            }
        }
    }
    out.push_str(rest);
    out
}

#[cfg(test)]
mod tests {
    use super::*;

    fn font(format: u16, loca_entry: usize) -> Vec<u8> {
        let mut head = vec![0u8; 54];
        head[50..52].copy_from_slice(&format.to_be_bytes());
        let mut maxp = vec![0u8; 6];
        maxp[4..6].copy_from_slice(&3u16.to_be_bytes());
        let loca = vec![0u8; 4 * loca_entry];
        sfnt(0x0001_0000, &[(*b"head", head), (*b"loca", loca), (*b"maxp", maxp)])
    }

    #[test]
    fn cmap_round_trip_and_radicals() {
        let mut map = std::collections::BTreeMap::new();
        map.insert(0x41u32, 3u16);
        map.insert(0x2F45u32, 7u16);
        map.insert(0x1F600u32, 9u16);
        let parsed = parse_cmap(&write_cmap(&map)).unwrap();
        assert_eq!(parsed, map);
        assert_eq!(unified(0x2F45), Some(0x65B9));
        assert_eq!(repair_svg_fonts("<text>⽅程</text>"), "<text>方程</text>");
    }

    #[test]
    fn repairs_wrong_loca_flag() {
        // 3 glyphs, short offsets (4 entries × 2 bytes) but the flag says long
        let broken = font(1, 2);
        let fixed = repair(&broken).expect("repaired");
        let (_, t) = tables(&fixed).unwrap();
        assert_eq!(be16(&t[0].1, 50), Some(0));
        assert!(repair(&font(0, 2)).is_none());
        assert!(repair(&font(1, 4)).is_none());
    }
}
