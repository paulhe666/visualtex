//! IBus's serialized objects (IBusText, IBusLookupTable, IBusPropList, as ibus-daemon reads them
//! in a D-Bus variant: a struct of the type's name, its attachments and its fields) and where the
//! IBus bus is.

use std::collections::HashMap;
use std::path::PathBuf;
use zbus::zvariant::{StructureBuilder, Value};

/// IBusAttribute types and underline styles (ibusattribute.h).
const ATTR_TYPE_UNDERLINE: u32 = 1;
const ATTR_TYPE_BACKGROUND: u32 = 3;
const ATTR_UNDERLINE_SINGLE: u32 = 1;

fn attachments() -> Value<'static> {
    Value::from(HashMap::<String, Value<'static>>::new())
}

fn structure(fields: Vec<Value<'static>>) -> Value<'static> {
    let mut builder = StructureBuilder::new();
    for field in fields {
        builder = builder.append_field(field);
    }
    Value::Structure(builder.build().expect("an IBus object has fields"))
}

fn variant(value: Value<'static>) -> Value<'static> {
    Value::Value(Box::new(value))
}

/// An IBusText, underlined from `underline.0` to `underline.1` (characters) when given.
pub fn text(content: &str, underline: Option<(u32, u32)>) -> Value<'static> {
    text_with_background(content, underline, None)
}

/// An IBusText with a part underlined and a part on a background (`background`: the range and an
/// RGB colour).
pub fn text_with_background(content: &str, underline: Option<(u32, u32)>, background: Option<(u32, u32, u32)>) -> Value<'static> {
    let mut attributes: Vec<Value<'static>> = Vec::new();
    if let Some((start, end, color)) = background.filter(|(start, end, _)| end > start) {
        attributes.push(structure(vec![
            Value::from("IBusAttribute"),
            attachments(),
            Value::from(ATTR_TYPE_BACKGROUND),
            Value::from(color),
            Value::from(start),
            Value::from(end),
        ]));
    }
    if let Some((start, end)) = underline.filter(|(start, end)| end > start) {
        attributes.push(structure(vec![
            Value::from("IBusAttribute"),
            attachments(),
            Value::from(ATTR_TYPE_UNDERLINE),
            Value::from(ATTR_UNDERLINE_SINGLE),
            Value::from(start),
            Value::from(end),
        ]));
    }
    let list = structure(vec![Value::from("IBusAttrList"), attachments(), Value::from(attributes)]);
    structure(vec![Value::from("IBusText"), attachments(), Value::from(content.to_string()), variant(list)])
}

/// An IBusLookupTable showing `candidates` (with `labels`), the `cursor`-th highlighted.
pub fn lookup_table(candidates: &[String], labels: &[String], cursor: u32) -> Value<'static> {
    let page_size = candidates.len().clamp(1, 16) as u32;
    let texts: Vec<Value<'static>> = candidates.iter().map(|candidate| text(candidate, None)).collect();
    let labels: Vec<Value<'static>> = labels.iter().take(candidates.len()).map(|label| text(label, None)).collect();
    structure(vec![
        Value::from("IBusLookupTable"),
        attachments(),
        Value::from(page_size),
        Value::from(cursor.min(page_size - 1)),
        Value::from(true),
        Value::from(false),
        // IBUS_ORIENTATION_SYSTEM: as the panel is set up
        Value::from(2i32),
        Value::from(texts),
        Value::from(labels),
    ])
}

/// An IBusProperty: a label (中 / 英, the panel's input mode) with no menu.
fn property(key: &str, label: &str, tooltip: &str, symbol: &str) -> Value<'static> {
    let empty_list = structure(vec![Value::from("IBusPropList"), attachments(), Value::from(Vec::<Value<'static>>::new())]);
    structure(vec![
        Value::from("IBusProperty"),
        attachments(),
        Value::from(key.to_string()),
        // PROP_TYPE_NORMAL
        Value::from(0u32),
        variant(text(label, None)),
        Value::from(String::new()),
        variant(text(tooltip, None)),
        Value::from(true),
        Value::from(true),
        // PROP_STATE_UNCHECKED
        Value::from(0u32),
        variant(empty_list),
        variant(text(symbol, None)),
    ])
}

pub const MODE_PROPERTY: &str = "InputMode";

/// The input mode property: Chinese (pinyin) or English.
pub fn mode_property(ascii: bool) -> Value<'static> {
    if ascii {
        property(MODE_PROPERTY, "英", "English", "英")
    } else {
        property(MODE_PROPERTY, "中", "Chinese (pinyin)", "中")
    }
}

pub fn property_list(properties: Vec<Value<'static>>) -> Value<'static> {
    structure(vec![Value::from("IBusPropList"), attachments(), Value::from(properties)])
}

/// The IBus bus: $IBUS_ADDRESS, else the address ibus-daemon wrote for this display
/// ($XDG_CONFIG_HOME/ibus/bus/<machine id>-<host>-<display>, as ibus_get_socket_path names it).
pub fn address() -> Option<String> {
    if let Some(address) = std::env::var("IBUS_ADDRESS").ok().filter(|value| !value.is_empty()) {
        return Some(address);
    }
    let machine = ["/var/lib/dbus/machine-id", "/etc/machine-id"]
        .iter()
        .find_map(|path| std::fs::read_to_string(path).ok())
        .map(|id| id.trim().to_string())?;
    let (host, number) = match std::env::var("WAYLAND_DISPLAY").ok().filter(|value| !value.is_empty()) {
        Some(wayland) => ("unix".to_string(), wayland),
        None => {
            let display = std::env::var("DISPLAY").ok().unwrap_or_else(|| ":0".to_string());
            let (host, rest) = display.split_once(':').unwrap_or(("", display.as_str()));
            let number = rest.split('.').next().unwrap_or("0").to_string();
            (if host.is_empty() { "unix".to_string() } else { host.to_string() }, number)
        }
    };
    let config = std::env::var_os("XDG_CONFIG_HOME")
        .map(PathBuf::from)
        .or_else(|| std::env::var_os("HOME").map(|home| PathBuf::from(home).join(".config")))?;
    let file = config.join("ibus").join("bus").join(format!("{machine}-{host}-{number}"));
    let content = std::fs::read_to_string(file).ok()?;
    content
        .lines()
        .find_map(|line| line.strip_prefix("IBUS_ADDRESS="))
        .map(str::to_string)
}
