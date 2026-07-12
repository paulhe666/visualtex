use std::fs::{self, File};
use std::io::BufWriter;
use std::path::Path;

fn main() {
    ensure_icon();
    tauri_build::build()
}

fn ensure_icon() {
    let path = Path::new("icons/icon.png");
    if path.exists() {
        return;
    }
    fs::create_dir_all(path.parent().expect("icon parent")).expect("create icon directory");

    const SIZE: u32 = 128;
    let mut pixels = vec![0_u8; (SIZE * SIZE * 4) as usize];
    for y in 0..SIZE {
        for x in 0..SIZE {
            let index = ((y * SIZE + x) * 4) as usize;
            let inside = rounded_square(x as f32, y as f32, SIZE as f32, 25.0);
            if !inside {
                continue;
            }
            pixels[index] = 42 + (x * 25 / SIZE) as u8;
            pixels[index + 1] = 78 + (y * 28 / SIZE) as u8;
            pixels[index + 2] = 225;
            pixels[index + 3] = 255;

            let v = distance_to_segment(x as f32, y as f32, 25.0, 31.0, 49.0, 95.0) < 5.0
                || distance_to_segment(x as f32, y as f32, 49.0, 95.0, 72.0, 31.0) < 5.0;
            let t = ((72..=108).contains(&x) && (30..=40).contains(&y))
                || ((86..=96).contains(&x) && (35..=97).contains(&y));
            if v || t {
                pixels[index] = 255;
                pixels[index + 1] = 255;
                pixels[index + 2] = 255;
                pixels[index + 3] = 255;
            }
        }
    }

    let file = File::create(path).expect("create generated icon");
    let writer = BufWriter::new(file);
    let mut encoder = png::Encoder::new(writer, SIZE, SIZE);
    encoder.set_color(png::ColorType::Rgba);
    encoder.set_depth(png::BitDepth::Eight);
    let mut png = encoder.write_header().expect("write icon header");
    png.write_image_data(&pixels).expect("write icon pixels");
}

fn rounded_square(x: f32, y: f32, size: f32, radius: f32) -> bool {
    let inner_min = radius;
    let inner_max = size - radius - 1.0;
    if (x >= inner_min && x <= inner_max) || (y >= inner_min && y <= inner_max) {
        return true;
    }
    let center_x = if x < inner_min { inner_min } else { inner_max };
    let center_y = if y < inner_min { inner_min } else { inner_max };
    (x - center_x).powi(2) + (y - center_y).powi(2) <= radius.powi(2)
}

fn distance_to_segment(x: f32, y: f32, start_x: f32, start_y: f32, end_x: f32, end_y: f32) -> f32 {
    let dx = end_x - start_x;
    let dy = end_y - start_y;
    let length_squared = dx * dx + dy * dy;
    let projection = (((x - start_x) * dx + (y - start_y) * dy) / length_squared).clamp(0.0, 1.0);
    let closest_x = start_x + projection * dx;
    let closest_y = start_y + projection * dy;
    ((x - closest_x).powi(2) + (y - closest_y).powi(2)).sqrt()
}
