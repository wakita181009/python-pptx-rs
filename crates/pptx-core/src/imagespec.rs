//! Image header sniffing for `add_picture`.
//!
//! Mirrors what python-pptx observes through PIL: canonical extension per
//! detected format (filename extension is ignored), pixel size, and dpi with
//! python-pptx's `int_dpi` normalization (non-numeric / <1 / >2048 → 72).
//! EMF is reported as `wmf`, matching PIL's WmfImagePlugin.

use crate::error::{Error, Result};

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ImageSpec {
    /// Canonical extension: `bmp` `gif` `jpg` `png` `tiff` `wmf`.
    pub ext: &'static str,
    pub content_type: &'static str,
    pub px_width: u32,
    pub px_height: u32,
    pub horz_dpi: u32,
    pub vert_dpi: u32,
}

const EMU_PER_INCH: i64 = 914_400;

impl ImageSpec {
    /// Native size in EMU (python-pptx `ImagePart._native_size`).
    pub fn native_size_emu(&self) -> (i64, i64) {
        let width = EMU_PER_INCH as f64 * self.px_width as f64 / self.horz_dpi as f64;
        let height = EMU_PER_INCH as f64 * self.px_height as f64 / self.vert_dpi as f64;
        (width as i64, height as i64)
    }
}

/// python-pptx `Image.dpi.int_dpi`: out-of-range or invalid → 72.
fn int_dpi(dpi: f64) -> u32 {
    let rounded = dpi.round();
    if !rounded.is_finite() || !(1.0..=2048.0).contains(&rounded) {
        72
    } else {
        rounded as u32
    }
}

fn unsupported() -> Error {
    Error::InvalidPackage(
        "unsupported image format, expected one of: BMP, GIF, JPEG, PNG, TIFF, WMF".into(),
    )
}

pub fn sniff(blob: &[u8]) -> Result<ImageSpec> {
    if blob.starts_with(b"\x89PNG\r\n\x1a\n") {
        return sniff_png(blob);
    }
    if blob.starts_with(b"\xff\xd8") {
        return sniff_jpeg(blob);
    }
    if blob.starts_with(b"GIF87a") || blob.starts_with(b"GIF89a") {
        return sniff_gif(blob);
    }
    if blob.starts_with(b"BM") {
        return sniff_bmp(blob);
    }
    if blob.starts_with(b"II*\x00") || blob.starts_with(b"MM\x00*") {
        return sniff_tiff(blob);
    }
    if blob.starts_with(b"\xd7\xcd\xc6\x9a") {
        return sniff_wmf_placeable(blob);
    }
    if blob.len() >= 44 && blob.starts_with(b"\x01\x00\x00\x00") && &blob[40..44] == b" EMF" {
        return sniff_emf(blob);
    }
    Err(unsupported())
}

fn be16(b: &[u8], at: usize) -> Option<u32> {
    Some(u32::from(*b.get(at)?) << 8 | u32::from(*b.get(at + 1)?))
}

fn be32(b: &[u8], at: usize) -> Option<u32> {
    let v: [u8; 4] = b.get(at..at + 4)?.try_into().ok()?;
    Some(u32::from_be_bytes(v))
}

fn le16(b: &[u8], at: usize) -> Option<u32> {
    Some(u32::from(*b.get(at)?) | u32::from(*b.get(at + 1)?) << 8)
}

fn le32(b: &[u8], at: usize) -> Option<i64> {
    let v: [u8; 4] = b.get(at..at + 4)?.try_into().ok()?;
    Some(i64::from(i32::from_le_bytes(v)))
}

fn sniff_png(blob: &[u8]) -> Result<ImageSpec> {
    // IHDR must be the first chunk; pHYs (unit 1 = meter) carries density.
    let (mut width, mut height) = (None, None);
    let (mut hdpi, mut vdpi) = (72u32, 72u32);
    let mut at = 8;
    while at + 8 <= blob.len() {
        let len = be32(blob, at).ok_or_else(unsupported)? as usize;
        let chunk = &blob[at + 4..at + 8];
        let data = at + 8;
        match chunk {
            b"IHDR" => {
                width = be32(blob, data);
                height = be32(blob, data + 4);
            }
            b"pHYs" if len >= 9 => {
                if blob.get(data + 8) == Some(&1) {
                    let ppm_x = be32(blob, data).ok_or_else(unsupported)?;
                    let ppm_y = be32(blob, data + 4).ok_or_else(unsupported)?;
                    hdpi = int_dpi(f64::from(ppm_x) * 0.0254);
                    vdpi = int_dpi(f64::from(ppm_y) * 0.0254);
                }
            }
            b"IDAT" | b"IEND" => break,
            _ => {}
        }
        at = data + len + 4;
    }
    let (px_width, px_height) = width.zip(height).ok_or_else(unsupported)?;
    Ok(ImageSpec {
        ext: "png",
        content_type: "image/png",
        px_width,
        px_height,
        horz_dpi: hdpi,
        vert_dpi: vdpi,
    })
}

fn sniff_jpeg(blob: &[u8]) -> Result<ImageSpec> {
    let (mut hdpi, mut vdpi) = (72u32, 72u32);
    let mut at = 2;
    loop {
        // resync to the next 0xFF marker byte, skipping fill bytes
        while blob.get(at) == Some(&0xff) && blob.get(at + 1) == Some(&0xff) {
            at += 1;
        }
        if blob.get(at) != Some(&0xff) {
            return Err(unsupported());
        }
        let marker = *blob.get(at + 1).ok_or_else(unsupported)?;
        match marker {
            0xd8 | 0x01 | 0xd0..=0xd7 => {
                at += 2;
                continue;
            }
            _ => {}
        }
        let len = be16(blob, at + 2).ok_or_else(unsupported)? as usize;
        let data = at + 4;
        match marker {
            // SOF0-15, minus DHT(C4)/JPG(C8)/DAC(CC)
            0xc0..=0xcf if !matches!(marker, 0xc4 | 0xc8 | 0xcc) => {
                let px_height = be16(blob, data + 1).ok_or_else(unsupported)?;
                let px_width = be16(blob, data + 3).ok_or_else(unsupported)?;
                return Ok(ImageSpec {
                    ext: "jpg",
                    content_type: "image/jpeg",
                    px_width,
                    px_height,
                    horz_dpi: hdpi,
                    vert_dpi: vdpi,
                });
            }
            // APP0/JFIF: units 1 = dpi, 2 = dots-per-cm
            0xe0 if len >= 14 && blob.get(data..data + 5) == Some(b"JFIF\x00") => {
                let units = *blob.get(data + 7).ok_or_else(unsupported)?;
                let xd = be16(blob, data + 8).ok_or_else(unsupported)?;
                let yd = be16(blob, data + 10).ok_or_else(unsupported)?;
                match units {
                    1 => {
                        hdpi = int_dpi(f64::from(xd));
                        vdpi = int_dpi(f64::from(yd));
                    }
                    2 => {
                        hdpi = int_dpi(f64::from(xd) * 2.54);
                        vdpi = int_dpi(f64::from(yd) * 2.54);
                    }
                    _ => {}
                }
            }
            0xd9 | 0xda => return Err(unsupported()),
            _ => {}
        }
        at += 2 + len;
    }
}

fn sniff_gif(blob: &[u8]) -> Result<ImageSpec> {
    let px_width = le16(blob, 6).ok_or_else(unsupported)?;
    let px_height = le16(blob, 8).ok_or_else(unsupported)?;
    Ok(ImageSpec {
        ext: "gif",
        content_type: "image/gif",
        px_width,
        px_height,
        horz_dpi: 72,
        vert_dpi: 72,
    })
}

fn sniff_bmp(blob: &[u8]) -> Result<ImageSpec> {
    let header_size = le32(blob, 14).ok_or_else(unsupported)?;
    if header_size == 12 {
        // BITMAPCOREHEADER: u16 dimensions, no density
        let px_width = le16(blob, 18).ok_or_else(unsupported)?;
        let px_height = le16(blob, 20).ok_or_else(unsupported)?;
        return Ok(ImageSpec {
            ext: "bmp",
            content_type: "image/bmp",
            px_width,
            px_height,
            horz_dpi: 72,
            vert_dpi: 72,
        });
    }
    let px_width = le32(blob, 18).ok_or_else(unsupported)?.unsigned_abs() as u32;
    // negative height = top-down DIB
    let px_height = le32(blob, 22).ok_or_else(unsupported)?.unsigned_abs() as u32;
    let ppm_x = le32(blob, 38).ok_or_else(unsupported)?;
    let ppm_y = le32(blob, 42).ok_or_else(unsupported)?;
    Ok(ImageSpec {
        ext: "bmp",
        content_type: "image/bmp",
        px_width,
        px_height,
        horz_dpi: int_dpi(ppm_x as f64 * 0.0254),
        vert_dpi: int_dpi(ppm_y as f64 * 0.0254),
    })
}

fn sniff_tiff(blob: &[u8]) -> Result<ImageSpec> {
    let big_endian = blob.starts_with(b"MM");
    let rd16 = |at: usize| -> Option<u32> {
        if big_endian {
            be16(blob, at)
        } else {
            le16(blob, at)
        }
    };
    let rd32 = |at: usize| -> Option<u32> {
        if big_endian {
            be32(blob, at)
        } else {
            le32(blob, at).map(|v| v as u32)
        }
    };
    let ifd = rd32(4).ok_or_else(unsupported)? as usize;
    let count = rd16(ifd).ok_or_else(unsupported)? as usize;

    let (mut width, mut height) = (None, None);
    let (mut x_res, mut y_res) = (None::<f64>, None::<f64>);
    let mut res_unit = 1u32; // TIFF default is "none"; PIL only maps units 2/3 to dpi
    for i in 0..count {
        let entry = ifd + 2 + i * 12;
        let tag = rd16(entry).ok_or_else(unsupported)?;
        let field_type = rd16(entry + 2).ok_or_else(unsupported)?;
        let value_at = entry + 8;
        let short_or_long = || -> Option<u32> {
            match field_type {
                3 => rd16(value_at),
                4 => rd32(value_at),
                _ => None,
            }
        };
        let rational = || -> Option<f64> {
            let offset = rd32(value_at)? as usize;
            let num = rd32(offset)?;
            let den = rd32(offset + 4)?;
            (den != 0).then(|| f64::from(num) / f64::from(den))
        };
        match tag {
            256 => width = short_or_long(),
            257 => height = short_or_long(),
            282 => x_res = rational(),
            283 => y_res = rational(),
            296 => res_unit = short_or_long().unwrap_or(1),
            _ => {}
        }
    }
    let (px_width, px_height) = width.zip(height).ok_or_else(unsupported)?;
    let to_dpi = |res: Option<f64>| match (res, res_unit) {
        (Some(r), 2) => int_dpi(r),
        (Some(r), 3) => int_dpi(r * 2.54),
        _ => 72,
    };
    Ok(ImageSpec {
        ext: "tiff",
        content_type: "image/tiff",
        px_width,
        px_height,
        horz_dpi: to_dpi(x_res),
        vert_dpi: to_dpi(y_res),
    })
}

fn sniff_wmf_placeable(blob: &[u8]) -> Result<ImageSpec> {
    // PIL normalizes placeable-WMF size to 72 dpi from twips-per-inch
    let i16_at = |at: usize| -> Option<i64> {
        let v: [u8; 2] = blob.get(at..at + 2)?.try_into().ok()?;
        Some(i64::from(i16::from_le_bytes(v)))
    };
    let x0 = i16_at(6).ok_or_else(unsupported)?;
    let y0 = i16_at(8).ok_or_else(unsupported)?;
    let x1 = i16_at(10).ok_or_else(unsupported)?;
    let y1 = i16_at(12).ok_or_else(unsupported)?;
    let inch = le16(blob, 14).ok_or_else(unsupported)? as i64;
    if inch == 0 {
        return Err(unsupported());
    }
    Ok(ImageSpec {
        ext: "wmf",
        content_type: "image/x-wmf",
        px_width: ((x1 - x0) * 72 / inch).max(0) as u32,
        px_height: ((y1 - y0) * 72 / inch).max(0) as u32,
        horz_dpi: 72,
        vert_dpi: 72,
    })
}

fn sniff_emf(blob: &[u8]) -> Result<ImageSpec> {
    // PIL reports EMF as format "WMF"; size from bounds (device px), dpi from
    // the frame rectangle (0.01 mm units)
    let x0 = le32(blob, 8).ok_or_else(unsupported)?;
    let y0 = le32(blob, 12).ok_or_else(unsupported)?;
    let x1 = le32(blob, 16).ok_or_else(unsupported)?;
    let y1 = le32(blob, 20).ok_or_else(unsupported)?;
    let fl = le32(blob, 24).ok_or_else(unsupported)?;
    let ft = le32(blob, 28).ok_or_else(unsupported)?;
    let fr = le32(blob, 32).ok_or_else(unsupported)?;
    let fb = le32(blob, 36).ok_or_else(unsupported)?;
    if fr == fl || fb == ft {
        return Err(unsupported());
    }
    let xdpi = 2540.0 * (x1 - x0) as f64 / (fr - fl) as f64;
    let ydpi = 2540.0 * (y1 - y0) as f64 / (fb - ft) as f64;
    Ok(ImageSpec {
        ext: "wmf",
        content_type: "image/x-wmf",
        px_width: (x1 - x0).max(0) as u32,
        px_height: (y1 - y0).max(0) as u32,
        horz_dpi: int_dpi(xdpi),
        vert_dpi: int_dpi(ydpi),
    })
}

#[cfg(test)]
mod tests {
    use super::*;

    fn png(width: u32, height: u32, phys: Option<(u32, u32, u8)>) -> Vec<u8> {
        let mut b = b"\x89PNG\r\n\x1a\n".to_vec();
        b.extend(13u32.to_be_bytes());
        b.extend(b"IHDR");
        b.extend(width.to_be_bytes());
        b.extend(height.to_be_bytes());
        b.extend([8, 6, 0, 0, 0]);
        b.extend([0u8; 4]); // crc (unchecked)
        if let Some((ppm_x, ppm_y, unit)) = phys {
            b.extend(9u32.to_be_bytes());
            b.extend(b"pHYs");
            b.extend(ppm_x.to_be_bytes());
            b.extend(ppm_y.to_be_bytes());
            b.push(unit);
            b.extend([0u8; 4]);
        }
        b.extend(0u32.to_be_bytes());
        b.extend(b"IEND");
        b.extend([0u8; 4]);
        b
    }

    #[test]
    fn sniffs_png_size_and_dpi() {
        let spec = sniff(&png(640, 480, Some((5906, 5906, 1)))).unwrap();
        assert_eq!(
            (spec.ext, spec.px_width, spec.px_height, spec.horz_dpi),
            ("png", 640, 480, 150)
        );
    }

    #[test]
    fn png_without_phys_defaults_to_72() {
        let spec = sniff(&png(10, 20, None)).unwrap();
        assert_eq!((spec.horz_dpi, spec.vert_dpi), (72, 72));
    }

    #[test]
    fn native_size_uses_dpi() {
        let spec = sniff(&png(144, 72, None)).unwrap();
        assert_eq!(spec.native_size_emu(), (1_828_800, 914_400));
    }

    #[test]
    fn sniffs_jpeg_jfif() {
        let mut b = vec![0xff, 0xd8];
        // APP0 JFIF, units=1, 300x200 dpi
        b.extend([0xff, 0xe0, 0x00, 0x10]);
        b.extend(b"JFIF\x00");
        b.extend([1, 2, 1, 0x01, 0x2c, 0x00, 0xc8, 0, 0]);
        // SOF0: precision 8, height 30, width 40
        b.extend([0xff, 0xc0, 0x00, 0x11, 8, 0, 30, 0, 40]);
        let spec = sniff(&b).unwrap();
        assert_eq!(
            (
                spec.ext,
                spec.px_width,
                spec.px_height,
                spec.horz_dpi,
                spec.vert_dpi
            ),
            ("jpg", 40, 30, 300, 200)
        );
    }

    #[test]
    fn sniffs_gif() {
        let mut b = b"GIF89a".to_vec();
        b.extend([0x20, 0x00, 0x10, 0x00]);
        let spec = sniff(&b).unwrap();
        assert_eq!((spec.ext, spec.px_width, spec.px_height), ("gif", 32, 16));
    }

    #[test]
    fn sniffs_bmp_with_density() {
        let mut b = vec![0u8; 54];
        b[0] = b'B';
        b[1] = b'M';
        b[14] = 40;
        b[18..22].copy_from_slice(&100i32.to_le_bytes());
        b[22..26].copy_from_slice(&(-50i32).to_le_bytes());
        b[38..42].copy_from_slice(&2835i32.to_le_bytes()); // 72 dpi
        b[42..46].copy_from_slice(&2835i32.to_le_bytes());
        let spec = sniff(&b).unwrap();
        assert_eq!(
            (spec.ext, spec.px_width, spec.px_height, spec.horz_dpi),
            ("bmp", 100, 50, 72)
        );
    }

    #[test]
    fn sniffs_tiff_le() {
        // II header, IFD at 8 with width/height/resolution(inch)
        let mut b = b"II*\x00".to_vec();
        b.extend(8u32.to_le_bytes());
        b.extend(5u16.to_le_bytes());
        let entry = |tag: u16, ftype: u16, value: u32| {
            let mut e = Vec::new();
            e.extend(tag.to_le_bytes());
            e.extend(ftype.to_le_bytes());
            e.extend(1u32.to_le_bytes());
            e.extend(value.to_le_bytes());
            e
        };
        b.extend(entry(256, 3, 200));
        b.extend(entry(257, 3, 100));
        b.extend(entry(282, 5, 74)); // offset of rational
        b.extend(entry(283, 5, 82));
        b.extend(entry(296, 3, 2));
        b.extend(0u32.to_le_bytes());
        assert_eq!(b.len(), 74);
        b.extend(150u32.to_le_bytes());
        b.extend(1u32.to_le_bytes());
        b.extend(300u32.to_le_bytes());
        b.extend(1u32.to_le_bytes());
        let spec = sniff(&b).unwrap();
        assert_eq!(
            (
                spec.ext,
                spec.px_width,
                spec.px_height,
                spec.horz_dpi,
                spec.vert_dpi
            ),
            ("tiff", 200, 100, 150, 300)
        );
    }

    #[test]
    fn sniffs_placeable_wmf() {
        let mut b = vec![0u8; 22];
        b[0..4].copy_from_slice(&[0xd7, 0xcd, 0xc6, 0x9a]);
        b[6..8].copy_from_slice(&0i16.to_le_bytes());
        b[8..10].copy_from_slice(&0i16.to_le_bytes());
        b[10..12].copy_from_slice(&2880i16.to_le_bytes());
        b[12..14].copy_from_slice(&1440i16.to_le_bytes());
        b[14..16].copy_from_slice(&1440u16.to_le_bytes());
        let spec = sniff(&b).unwrap();
        assert_eq!(
            (spec.ext, spec.px_width, spec.px_height, spec.horz_dpi),
            ("wmf", 144, 72, 72)
        );
    }

    #[test]
    fn sniffs_emf_as_wmf() {
        let mut b = vec![0u8; 88];
        b[0..4].copy_from_slice(&[0x01, 0x00, 0x00, 0x00]);
        b[8..12].copy_from_slice(&0i32.to_le_bytes());
        b[12..16].copy_from_slice(&0i32.to_le_bytes());
        b[16..20].copy_from_slice(&96i32.to_le_bytes());
        b[20..24].copy_from_slice(&96i32.to_le_bytes());
        b[24..28].copy_from_slice(&0i32.to_le_bytes());
        b[28..32].copy_from_slice(&0i32.to_le_bytes());
        b[32..36].copy_from_slice(&2540i32.to_le_bytes()); // 1 inch
        b[36..40].copy_from_slice(&2540i32.to_le_bytes());
        b[40..44].copy_from_slice(b" EMF");
        let spec = sniff(&b).unwrap();
        assert_eq!((spec.ext, spec.px_width, spec.horz_dpi), ("wmf", 96, 96));
    }

    #[test]
    fn rejects_unknown_format() {
        assert!(sniff(b"not an image").is_err());
    }
}
