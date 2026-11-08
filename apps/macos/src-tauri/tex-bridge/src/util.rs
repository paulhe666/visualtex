//! Small shared pieces: affine matrices, a stable hash, number formatting.

/// An affine map in PDF order: (x, y) -> (a x + c y + e, b x + d y + f).
#[derive(Clone, Copy, Debug, PartialEq)]
pub struct Mat {
    pub a: f64,
    pub b: f64,
    pub c: f64,
    pub d: f64,
    pub e: f64,
    pub f: f64,
}

impl Mat {
    pub const ID: Mat = Mat { a: 1.0, b: 0.0, c: 0.0, d: 1.0, e: 0.0, f: 0.0 };

    pub fn new(a: f64, b: f64, c: f64, d: f64, e: f64, f: f64) -> Mat {
        Mat { a, b, c, d, e, f }
    }

    pub fn translate(x: f64, y: f64) -> Mat {
        Mat::new(1.0, 0.0, 0.0, 1.0, x, y)
    }

    /// `self · other`: `other` applied first, then `self` (column vectors).
    pub fn mul(&self, o: &Mat) -> Mat {
        Mat {
            a: self.a * o.a + self.c * o.b,
            b: self.b * o.a + self.d * o.b,
            c: self.a * o.c + self.c * o.d,
            d: self.b * o.c + self.d * o.d,
            e: self.a * o.e + self.c * o.f + self.e,
            f: self.b * o.e + self.d * o.f + self.f,
        }
    }

    pub fn apply(&self, x: f64, y: f64) -> (f64, f64) {
        (self.a * x + self.c * y + self.e, self.b * x + self.d * y + self.f)
    }

    pub fn det(&self) -> f64 {
        self.a * self.d - self.b * self.c
    }

    pub fn invert(&self) -> Option<Mat> {
        let det = self.det();
        if det.abs() < 1e-12 {
            return None;
        }
        let a = self.d / det;
        let b = -self.b / det;
        let c = -self.c / det;
        let d = self.a / det;
        Some(Mat { a, b, c, d, e: -(a * self.e + c * self.f), f: -(b * self.e + d * self.f) })
    }

    pub fn is_identity(&self) -> bool {
        (self.a - 1.0).abs() < 1e-9
            && self.b.abs() < 1e-9
            && self.c.abs() < 1e-9
            && (self.d - 1.0).abs() < 1e-9
            && self.e.abs() < 1e-6
            && self.f.abs() < 1e-6
    }

    /// The SVG `matrix(...)` value.
    pub fn svg(&self) -> String {
        format!("matrix({} {} {} {} {} {})", num(self.a), num(self.b), num(self.c), num(self.d), num(self.e), num(self.f))
    }
}

/// A number for SVG/PDF text: at most 4 decimals, no trailing zeros, no "-0".
pub fn num(v: f64) -> String {
    if !v.is_finite() {
        return "0".into();
    }
    let mut s = format!("{:.4}", v);
    if s.contains('.') {
        while s.ends_with('0') {
            s.pop();
        }
        if s.ends_with('.') {
            s.pop();
        }
    }
    if s == "-0" {
        s = "0".into();
    }
    s
}

/// FNV-1a, 64 bit: a hash that stays the same between runs and builds.
#[derive(Clone, Copy)]
pub struct Fnv(pub u64);

impl Default for Fnv {
    fn default() -> Self {
        Fnv(0xcbf2_9ce4_8422_2325)
    }
}

impl Fnv {
    pub fn write(&mut self, bytes: &[u8]) {
        for &b in bytes {
            self.0 ^= b as u64;
            self.0 = self.0.wrapping_mul(0x0000_0100_0000_01b3);
        }
    }
    pub fn write_u64(&mut self, v: u64) {
        self.write(&v.to_le_bytes());
    }
    pub fn finish(&self) -> u64 {
        self.0
    }
}

pub fn hash_bytes(bytes: &[u8]) -> u64 {
    let mut h = Fnv::default();
    h.write(bytes);
    h.finish()
}

pub fn hex(v: u64) -> String {
    format!("{:016x}", v)
}

/// XML attribute text.
pub fn xml_escape(s: &str) -> String {
    s.replace('&', "&amp;").replace('<', "&lt;").replace('>', "&gt;").replace('\'', "&apos;").replace('"', "&quot;")
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn matrix_inverse_round_trip() {
        let m = Mat::new(2.0, 0.5, -0.3, 1.5, 10.0, -4.0);
        let r = m.mul(&m.invert().unwrap());
        assert!(r.is_identity());
        let (x, y) = m.apply(1.0, 2.0);
        assert!((x - (2.0 - 0.6 + 10.0)).abs() < 1e-12 && (y - (0.5 + 3.0 - 4.0)).abs() < 1e-12);
    }

    #[test]
    fn numbers_are_short() {
        assert_eq!(num(1.0), "1");
        assert_eq!(num(-0.00001), "0");
        assert_eq!(num(2.50004), "2.5");
        assert_eq!(num(-3.14159), "-3.1416");
    }
}
