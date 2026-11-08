//! Character widths from TFM files (the DVI positions of TFM fonts' characters need them).

use std::collections::HashMap;

#[derive(Clone, Debug, Default)]
pub struct Tfm {
    pub bc: u32,
    /// widths as fractions of the design size, by character code from `bc`
    pub widths: Vec<f64>,
}

impl Tfm {
    pub fn parse(data: &[u8]) -> Option<Tfm> {
        let half = |i: usize| -> Option<u32> { Some(((*data.get(2 * i)? as u32) << 8) | *data.get(2 * i + 1)? as u32) };
        let word = |i: usize| -> Option<i32> {
            let b = data.get(4 * i..4 * i + 4)?;
            Some(i32::from_be_bytes([b[0], b[1], b[2], b[3]]))
        };
        let lh = half(1)? as usize;
        let bc = half(2)?;
        let ec = half(3)?;
        let nw = half(4)? as usize;
        if ec < bc || ec > 255 + 0xffff {
            return Some(Tfm { bc: 0, widths: vec![] });
        }
        let n = (ec - bc + 1) as usize;
        let char_info = 6 + lh;
        let width_base = char_info + n;
        let mut widths = Vec::with_capacity(n);
        for i in 0..n {
            let info = word(char_info + i)?;
            let wi = ((info as u32) >> 24) as usize;
            let w = if wi == 0 || wi >= nw { 0.0 } else { word(width_base + wi)? as f64 / (1u64 << 20) as f64 };
            widths.push(w);
        }
        Some(Tfm { bc, widths })
    }

    /// Width of `ch` scaled to `scale` (both in DVI units).
    pub fn width(&self, ch: u32, scale: i32) -> f64 {
        if ch < self.bc {
            return 0.0;
        }
        self.widths.get((ch - self.bc) as usize).map(|w| w * scale as f64).unwrap_or(0.0)
    }
}

/// The widths of a document's TFM fonts, by DVI font number.
#[derive(Default)]
pub struct Metrics {
    pub fonts: HashMap<u32, (Tfm, i32)>,
}

impl Metrics {
    pub fn width(&self, font: u32, ch: u32) -> f64 {
        self.fonts.get(&font).map(|(t, s)| t.width(ch, *s)).unwrap_or(0.0)
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn reads_a_small_tfm() {
        // lf lh bc ec nw nh nd ni nl nk ne np, then 2 header words, 2 char infos, 3 widths
        let mut d: Vec<u8> = vec![];
        for v in [0u16, 2, 65, 66, 3, 1, 1, 1, 0, 0, 0, 0] {
            d.extend_from_slice(&v.to_be_bytes());
        }
        d.extend_from_slice(&[0; 8]);
        d.extend_from_slice(&[1, 0, 0, 0]); // 'A' -> width 1
        d.extend_from_slice(&[2, 0, 0, 0]); // 'B' -> width 2
        d.extend_from_slice(&0i32.to_be_bytes());
        d.extend_from_slice(&(1i32 << 19).to_be_bytes()); // 0.5
        d.extend_from_slice(&(1i32 << 20).to_be_bytes()); // 1.0
        let t = Tfm::parse(&d).unwrap();
        assert_eq!(t.width(65, 1000), 500.0);
        assert_eq!(t.width(66, 1000), 1000.0);
        assert_eq!(t.width(67, 1000), 0.0);
    }
}
