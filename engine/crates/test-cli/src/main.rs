use std::io::{self, BufRead, Write};
use visualtex_engine_core::{Runtime, ErrorCode, MAX_WIRE_BYTES, error_json};

// Newline-delimited JSON. Keep the host buffer bounded too, rather than waiting
// until an unbounded read_line has already allocated a malicious request.
fn read_request(reader: &mut impl BufRead) -> io::Result<Option<Result<String, &'static str>>> {
    let mut bytes = Vec::new();
    let mut oversized = false;
    let mut saw_data = false;
    loop {
        let chunk = reader.fill_buf()?;
        if chunk.is_empty() {
            if !saw_data { return Ok(None); }
            break;
        }
        saw_data = true;
        let end = chunk.iter().position(|b| *b == b'\n');
        let count = end.map_or(chunk.len(), |p| p + 1);
        let content = if end.is_some() { &chunk[..count - 1] } else { &chunk[..count] };
        if !oversized {
            if bytes.len().saturating_add(content.len()) > MAX_WIRE_BYTES {
                oversized = true;
                bytes.clear();
            } else {
                bytes.extend_from_slice(content);
            }
        }
        reader.consume(count);
        if end.is_some() { break; }
    }
    if oversized { return Ok(Some(Err("resourceLimit"))); }
    if bytes.last() == Some(&b'\r') { bytes.pop(); }
    Ok(Some(String::from_utf8(bytes).map_err(|_| "invalidUtf8")))
}

fn main() -> io::Result<()> {
    let input = io::stdin();
    let mut reader = input.lock();
    let output = io::stdout();
    let mut writer = io::BufWriter::new(output.lock());
    let mut runtime = Runtime::default();
    // Test-host resource discovery only; filenames never enter the pure core.
    let mut args = std::env::args_os().skip(1);
    while let Some(flag) = args.next() {
        if flag != "--font" { return Err(io::Error::new(io::ErrorKind::InvalidInput, "Expected --font <path>")); }
        let path = args.next().ok_or_else(|| io::Error::new(io::ErrorKind::InvalidInput, "Missing font path"))?;
        let file = std::fs::File::open(path)?;
        let mut bytes = Vec::new();
        let limit = visualtex_engine_core::fonts::MAX_FONT_BYTES as u64 + 1;
        io::Read::read_to_end(&mut io::Read::take(file, limit), &mut bytes)?;
        runtime.register_font(&bytes, 0).map_err(|e| io::Error::new(io::ErrorKind::InvalidData, e.message))?;
    }
    while let Some(request) = read_request(&mut reader)? {
        let response = match request {
            Ok(request) => runtime.dispatch_json(&request),
            Err("resourceLimit") => error_json(ErrorCode::ResourceLimit, "Request exceeds the byte budget"),
            Err(_) => error_json(ErrorCode::InvalidInput, "Request is not valid UTF-8"),
        };
        writeln!(writer, "{response}")?;
        writer.flush()?;
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::io::Cursor;
    #[test]
    fn bounded_lines_recover_after_oversized_request() {
        let input = format!("{}\n{{}}\r\n", "x".repeat(MAX_WIRE_BYTES + 1));
        let mut cursor = Cursor::new(input);
        assert_eq!(read_request(&mut cursor).unwrap(), Some(Err("resourceLimit")));
        assert_eq!(read_request(&mut cursor).unwrap(), Some(Ok("{}".into())));
        assert_eq!(read_request(&mut cursor).unwrap(), None);
    }
    #[test]
    fn invalid_utf8_is_rejected_and_final_line_is_read() {
        let mut cursor = Cursor::new(vec![255, b'\n', b'{', b'}']);
        assert_eq!(read_request(&mut cursor).unwrap(), Some(Err("invalidUtf8")));
        assert_eq!(read_request(&mut cursor).unwrap(), Some(Ok("{}".into())));
    }
}
