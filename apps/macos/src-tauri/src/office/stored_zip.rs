//! Minimal writer for uncompressed ("stored") ZIP packages.
//!
//! Office staging documents are tiny and are read back immediately by Word or
//! PowerPoint, so compression buys nothing; a stored package keeps the bytes
//! deterministic.

fn crc32(bytes: &[u8]) -> u32 {
    let mut crc = 0xffff_ffff_u32;
    for byte in bytes {
        crc ^= u32::from(*byte);
        for _ in 0..8 {
            let mask = 0_u32.wrapping_sub(crc & 1);
            crc = (crc >> 1) ^ (0xedb8_8320 & mask);
        }
    }
    !crc
}

fn push_zip_u16(output: &mut Vec<u8>, value: u16) {
    output.extend_from_slice(&value.to_le_bytes());
}

fn push_zip_u32(output: &mut Vec<u8>, value: u32) {
    output.extend_from_slice(&value.to_le_bytes());
}

struct StoredZipEntry {
    name: Vec<u8>,
    crc: u32,
    size: u32,
    offset: u32,
}

pub fn build_stored_zip<N, C>(entries: &[(N, C)]) -> Result<Vec<u8>, String>
where
    N: AsRef<str>,
    C: AsRef<[u8]>,
{
    let entry_count = u16::try_from(entries.len())
        .map_err(|_| "ZIP package has too many entries".to_string())?;
    let mut output = Vec::new();
    let mut records = Vec::with_capacity(entries.len());

    for (name, contents) in entries {
        let name = name.as_ref();
        let contents = contents.as_ref();
        let name_bytes = name.as_bytes();
        let name_length = u16::try_from(name_bytes.len())
            .map_err(|_| "ZIP package contains an overlong entry name".to_string())?;
        let size = u32::try_from(contents.len())
            .map_err(|_| "ZIP package entry is too large".to_string())?;
        let offset = u32::try_from(output.len())
            .map_err(|_| "ZIP package is too large".to_string())?;
        let checksum = crc32(contents);

        push_zip_u32(&mut output, 0x0403_4b50);
        push_zip_u16(&mut output, 20);
        push_zip_u16(&mut output, 0x0800);
        push_zip_u16(&mut output, 0);
        push_zip_u16(&mut output, 0);
        push_zip_u16(&mut output, 33);
        push_zip_u32(&mut output, checksum);
        push_zip_u32(&mut output, size);
        push_zip_u32(&mut output, size);
        push_zip_u16(&mut output, name_length);
        push_zip_u16(&mut output, 0);
        output.extend_from_slice(name_bytes);
        output.extend_from_slice(contents);

        records.push(StoredZipEntry {
            name: name_bytes.to_vec(),
            crc: checksum,
            size,
            offset,
        });
    }

    let central_offset = u32::try_from(output.len())
        .map_err(|_| "ZIP package is too large".to_string())?;
    for record in &records {
        let name_length = u16::try_from(record.name.len())
            .map_err(|_| "ZIP package contains an overlong entry name".to_string())?;
        push_zip_u32(&mut output, 0x0201_4b50);
        push_zip_u16(&mut output, 20);
        push_zip_u16(&mut output, 20);
        push_zip_u16(&mut output, 0x0800);
        push_zip_u16(&mut output, 0);
        push_zip_u16(&mut output, 0);
        push_zip_u16(&mut output, 33);
        push_zip_u32(&mut output, record.crc);
        push_zip_u32(&mut output, record.size);
        push_zip_u32(&mut output, record.size);
        push_zip_u16(&mut output, name_length);
        push_zip_u16(&mut output, 0);
        push_zip_u16(&mut output, 0);
        push_zip_u16(&mut output, 0);
        push_zip_u16(&mut output, 0);
        push_zip_u32(&mut output, 0);
        push_zip_u32(&mut output, record.offset);
        output.extend_from_slice(&record.name);
    }
    let central_size = u32::try_from(output.len())
        .map_err(|_| "ZIP package is too large".to_string())?
        .checked_sub(central_offset)
        .ok_or_else(|| "ZIP package central directory is invalid".to_string())?;

    push_zip_u32(&mut output, 0x0605_4b50);
    push_zip_u16(&mut output, 0);
    push_zip_u16(&mut output, 0);
    push_zip_u16(&mut output, entry_count);
    push_zip_u16(&mut output, entry_count);
    push_zip_u32(&mut output, central_size);
    push_zip_u32(&mut output, central_offset);
    push_zip_u16(&mut output, 0);
    Ok(output)
}
