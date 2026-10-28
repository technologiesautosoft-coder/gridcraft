//! Excel 2003 XML Spreadsheet import, including XML reports distributed with an `.xls` suffix.
//! Reads typed values, cached formula results, basic styles, dimensions and merged cells.
//! The namespace and element/attribute names follow Microsoft's Office 2003 XML schemas.
//! This is independent of the BIFF reader; no code from another spreadsheet is used.

use std::collections::{HashMap, HashSet};
use std::sync::Arc;

use chrono::{Datelike, NaiveDateTime, Timelike};
use gridcraft_core::{CellError, CellRef, DateSystem, MAX_COLS, MAX_ROWS, RangeRef, Value};
use gridcraft_model::{BorderLine, BorderStyle, Cell, Color, Fill, HAlign, LineInfo, NumFmt, Sheet, Style, StyleId, VAlign, Workbook};
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::ResolveResult;
use quick_xml::{NsReader, XmlVersion};

use crate::xml::{self, El};
use crate::{IoError, ReadReport};

const NS: &[u8] = b"urn:schemas-microsoft-com:office:spreadsheet";
const MAX_BYTES: usize = 32 * 1024 * 1024;
const MAX_NODES: usize = 250_000;
const MAX_CELLS: u64 = 1_000_000;
const MAX_STYLES: usize = 4096;
const MAX_SHEETS: usize = 256;

fn error(message: impl std::fmt::Display) -> IoError {
    IoError::Format(format!("XML Spreadsheet: {message}"))
}

fn bound_ns(ns: ResolveResult<'_>) -> bool {
    matches!(ns, ResolveResult::Bound(n) if n.as_ref() == NS)
}

/// Only the root in the SpreadsheetML 2003 namespace identifies this format. A CSV containing
/// XML snippets, a different XML document or an XLSX namespace must not take this path.
pub(crate) fn is_spreadsheetml(bytes: &[u8]) -> bool {
    let head = bytes.get(..bytes.len().min(64 * 1024)).unwrap_or(bytes);
    let decoded = xml::to_utf8(head);
    let mut reader = NsReader::from_reader(decoded.as_ref());
    for _ in 0..256 {
        match reader.read_event() {
            Ok(Event::Start(e) | Event::Empty(e)) => {
                let (ns, name) = reader.resolver_mut().resolve_element(e.name());
                return bound_ns(ns) && name.as_ref() == b"Workbook";
            }
            Ok(Event::Decl(_) | Event::PI(_) | Event::Comment(_) | Event::DocType(_)) => {}
            Ok(Event::Text(t)) if t.as_ref().iter().all(u8::is_ascii_whitespace) => {}
            _ => return false,
        }
    }
    false
}

/// Import only; the caller must save a new `.xlsx` rather than overwrite the XML source.
pub fn read_spreadsheetml(bytes: &[u8]) -> Result<(Workbook, ReadReport), IoError> {
    if bytes.len() > MAX_BYTES {
        return Err(IoError::TooLarge("XML Spreadsheet exceeds 32 MiB".into()));
    }
    let root = parse(bytes)?;
    let mut wb = Workbook::new();
    wb.sheets.clear();
    let mut report = ReadReport::default();
    let style_nodes: HashMap<&str, &El> =
        root.child("Styles").into_iter().flat_map(|e| e.kids("Style")).filter_map(|e| e.attr("ID").map(|id| (id, e))).collect();
    if style_nodes.len() > MAX_STYLES {
        return Err(IoError::TooLarge("XML Spreadsheet contains more than 4096 styles".into()));
    }
    let default = resolve_style("Default", &style_nodes, &mut Vec::new())?;
    let mut styles = HashMap::new();
    styles.insert("Default", wb.styles.intern(default));
    for id in style_nodes.keys() {
        styles.insert(*id, wb.styles.intern(resolve_style(id, &style_nodes, &mut Vec::new())?));
    }
    let default_id = styles.get("Default").copied().unwrap_or_default();
    let mut budget = 0u64;
    for worksheet in root.kids("Worksheet") {
        if wb.sheets.len() >= MAX_SHEETS {
            return Err(IoError::TooLarge("XML Spreadsheet contains more than 256 sheets".into()));
        }
        let name = worksheet.attr("Name").unwrap_or("Sheet1");
        wb.check_sheet_name(name, None).map_err(error)?;
        let mut sheet = Sheet::new(name);
        if let Some(table) = worksheet.child("Table") {
            read_table(table, &mut sheet, &styles, default_id, &mut budget, &mut report)?;
        }
        for unsupported in ["WorksheetOptions", "AutoFilter", "Names", "ConditionalFormatting", "DataValidation"] {
            if worksheet.child(unsupported).is_some() || worksheet.child(&format!("unsupported:{unsupported}")).is_some() {
                let feature = match unsupported {
                    "WorksheetOptions" => "print and sheet view settings",
                    "AutoFilter" => "filters",
                    "Names" => "defined names",
                    "ConditionalFormatting" => "conditional formatting",
                    _ => "data validation",
                };
                report.warn(format!("XML Spreadsheet import does not preserve {feature}"));
            }
        }
        wb.sheets.push(Arc::new(sheet));
    }
    if wb.sheets.is_empty() {
        return Err(error("the workbook contains no worksheets"));
    }
    if root.child("Names").is_some() {
        report.warn("XML Spreadsheet: defined names are not imported");
    }
    Ok((wb, report))
}

/// A strict, namespace-aware bounded tree. Unlike the lenient XLSX recovery parser, standalone
/// XML must have one complete root. DTDs/entities are refused and never fetched or expanded.
fn parse(bytes: &[u8]) -> Result<El, IoError> {
    let data = xml::to_utf8(bytes);
    let mut reader = NsReader::from_reader(data.as_ref());
    let mut stack: Vec<El> = Vec::new();
    let mut root = None;
    let mut nodes = 0usize;
    loop {
        let event = reader.read_event().map_err(error)?;
        let empty = matches!(&event, Event::Empty(_));
        match event {
            Event::Start(e) | Event::Empty(e) => {
                nodes += 1;
                if nodes > MAX_NODES || stack.len() >= 64 {
                    return Err(IoError::TooLarge("XML Spreadsheet has too many elements or is nested too deeply".into()));
                }
                let element = element(&mut reader, &e)?;
                if stack.is_empty() && (root.is_some() || element.name != "Workbook") {
                    return Err(error("expected one SpreadsheetML Workbook root"));
                }
                if empty {
                    finish(&mut stack, &mut root, element)?;
                } else {
                    stack.push(element);
                }
            }
            Event::End(_) => {
                let element = stack.pop().ok_or_else(|| error("unexpected closing element"))?;
                finish(&mut stack, &mut root, element)?;
            }
            Event::Text(t) => {
                let text = t.xml_content(XmlVersion::Explicit1_0).map_err(error)?;
                append_text(&mut stack, &text)?;
            }
            Event::CData(t) => {
                let text = std::str::from_utf8(t.as_ref()).map_err(error)?;
                append_text(&mut stack, text)?;
            }
            Event::GeneralRef(r) => {
                let text = match r.resolve_char_ref().map_err(error)? {
                    Some(c) => c.to_string(),
                    None => match r.as_ref() {
                        b"amp" => "&".into(),
                        b"lt" => "<".into(),
                        b"gt" => ">".into(),
                        b"quot" => "\"".into(),
                        b"apos" => "'".into(),
                        _ => return Err(error("unknown XML entity")),
                    },
                };
                append_text(&mut stack, &text)?;
            }
            Event::DocType(_) => return Err(error("DOCTYPE is not supported")),
            Event::Eof => break,
            _ => {}
        }
    }
    if !stack.is_empty() {
        return Err(error("the XML is truncated"));
    }
    root.ok_or_else(|| error("no Workbook root"))
}

fn element(reader: &mut NsReader<&[u8]>, event: &BytesStart<'_>) -> Result<El, IoError> {
    let (ns, local) = reader.resolver_mut().resolve_element(event.name());
    let local = String::from_utf8_lossy(local.as_ref()).into_owned();
    let name = if bound_ns(ns) { local } else { format!("unsupported:{local}") };
    let mut el = El { name, ..Default::default() };
    for attribute in event.attributes() {
        let attribute = attribute.map_err(error)?;
        let (ns, local) = reader.resolver_mut().resolve_attribute(attribute.key);
        if bound_ns(ns) {
            let key = String::from_utf8_lossy(local.as_ref()).into_owned();
            let value = attribute.normalized_value(XmlVersion::Explicit1_0).map_err(error)?.into_owned();
            if value.len() > 1024 * 1024 {
                return Err(IoError::TooLarge("XML Spreadsheet attribute exceeds 1 MiB".into()));
            }
            el.attrs.push((key, value));
        }
    }
    Ok(el)
}

fn finish(stack: &mut [El], root: &mut Option<El>, element: El) -> Result<(), IoError> {
    if let Some(parent) = stack.last_mut() {
        parent.children.push(element);
    } else if root.replace(element).is_some() {
        return Err(error("multiple roots"));
    }
    Ok(())
}

fn append_text(stack: &mut [El], text: &str) -> Result<(), IoError> {
    if let Some(current) = stack.last_mut() {
        if current.text.len().saturating_add(text.len()) > 1024 * 1024 {
            return Err(IoError::TooLarge("XML Spreadsheet cell text exceeds 1 MiB".into()));
        }
        current.text.push_str(text);
    } else if !text.trim().is_empty() {
        return Err(error("text outside the Workbook root"));
    }
    Ok(())
}

fn resolve_style<'a>(id: &'a str, nodes: &HashMap<&'a str, &'a El>, visiting: &mut Vec<&'a str>) -> Result<Style, IoError> {
    if visiting.len() >= 32 || visiting.contains(&id) {
        return Err(error("cyclic or excessively deep style inheritance"));
    }
    visiting.push(id);
    let mut style = match nodes.get(id).and_then(|e| e.attr("Parent")).or(if id != "Default" { Some("Default") } else { None }) {
        Some(parent) => resolve_style(parent, nodes, visiting)?,
        None => Style::default(),
    };
    if let Some(e) = nodes.get(id) {
        apply_style(&mut style, e);
    }
    visiting.pop();
    Ok(style)
}

fn color(s: Option<&str>) -> Option<Color> {
    let hex = s?.strip_prefix('#')?;
    (hex.len() == 6).then(|| u32::from_str_radix(hex, 16).ok()).flatten().map(Color::Rgb)
}

fn apply_style(style: &mut Style, e: &El) {
    if let Some(font) = e.child("Font") {
        if let Some(name) = font.attr("FontName") {
            style.font.name = name.to_string();
            style.font.minor = false;
        }
        if let Some(size) = font.attr_f64("Size").filter(|n| *n > 0.0 && *n <= 409.0) {
            style.font.size = size as f32;
        }
        if let Some(bold) = font.attr_bool("Bold") {
            style.font.bold = bold;
        }
        if let Some(italic) = font.attr_bool("Italic") {
            style.font.italic = italic;
        }
        if let Some(c) = color(font.attr("Color")) {
            style.font.color = c;
        }
    }
    if let Some(fill) = e.child("Interior")
        && fill.attr("Pattern") == Some("Solid")
        && let Some(c) = color(fill.attr("Color"))
    {
        style.fill = Fill::solid(c);
    }
    if let Some(alignment) = e.child("Alignment") {
        if let Some(horizontal) = alignment.attr("Horizontal") {
            style.align.h = match horizontal {
                "Left" => HAlign::Left,
                "Center" => HAlign::Center,
                "Right" => HAlign::Right,
                "Justify" => HAlign::Justify,
                "Fill" => HAlign::Fill,
                "CenterAcrossSelection" => HAlign::CenterAcross,
                _ => HAlign::General,
            };
        }
        if let Some(vertical) = alignment.attr("Vertical") {
            style.align.v = match vertical {
                "Top" => VAlign::Top,
                "Center" => VAlign::Center,
                _ => VAlign::Bottom,
            };
        }
        if let Some(wrap) = alignment.attr_bool("WrapText") {
            style.align.wrap = wrap;
        }
        if let Some(shrink) = alignment.attr_bool("ShrinkToFit") {
            style.align.shrink = shrink;
        }
    }
    if let Some(format) = e.child("NumberFormat").and_then(|f| f.attr("Format")) {
        style.num_fmt = NumFmt(format.into());
    }
    for border in e.child("Borders").into_iter().flat_map(|b| b.kids("Border")) {
        let line = BorderLine {
            style: match border.attr("LineStyle") {
                Some("Double") => BorderStyle::Double,
                Some("Dash") => BorderStyle::Dashed,
                Some("Dot") => BorderStyle::Dotted,
                Some("None") | None => BorderStyle::None,
                _ => match border.attr_u32("Weight").unwrap_or(1) {
                    0 => BorderStyle::Hair,
                    2 => BorderStyle::Medium,
                    3 => BorderStyle::Thick,
                    _ => BorderStyle::Thin,
                },
            },
            color: color(border.attr("Color")).unwrap_or(Color::Auto),
        };
        match border.attr("Position") {
            Some("Top") => style.borders.top = line,
            Some("Bottom") => style.borders.bottom = line,
            Some("Left") => style.borders.left = line,
            Some("Right") => style.borders.right = line,
            _ => {}
        }
    }
}

fn number(e: &El, key: &str, default: u32, max: u32) -> Result<u32, IoError> {
    let n = match e.attr(key) {
        Some(value) => value.parse::<u32>().map_err(|_| error(format!("invalid {key}")))?,
        None => default,
    };
    if n > max {
        return Err(error(format!("{key} exceeds worksheet bounds")));
    }
    Ok(n)
}

fn index(e: &El, next: u32, max: u32) -> Result<u32, IoError> {
    let n = number(e, "Index", next, max)?;
    if n == 0 || n < next {
        return Err(error("indices must be positive and increasing"));
    }
    Ok(n - 1)
}

fn style_id(e: &El, styles: &HashMap<&str, StyleId>, fallback: StyleId, report: &mut ReadReport) -> StyleId {
    match e.attr("StyleID") {
        Some(id) => styles.get(id).copied().unwrap_or_else(|| {
            report.warn(format!("XML Spreadsheet: unknown style {id}"));
            fallback
        }),
        None => fallback,
    }
}

fn dimension(e: &El, key: &str) -> Option<f32> {
    e.attr_f64(key).filter(|n| *n >= 0.0 && *n <= 32767.0).map(|n| (n * 4.0 / 3.0) as f32)
}

fn charge(budget: &mut u64, count: u64) -> Result<(), IoError> {
    *budget = budget.saturating_add(count);
    if *budget > MAX_CELLS {
        return Err(IoError::TooLarge("XML Spreadsheet expands beyond 1 million cells or dimensions".into()));
    }
    Ok(())
}

fn read_table(
    table: &El,
    sheet: &mut Sheet,
    styles: &HashMap<&str, StyleId>,
    default: StyleId,
    budget: &mut u64,
    report: &mut ReadReport,
) -> Result<(), IoError> {
    if number(table, "TopCell", 1, MAX_ROWS)? != 1 || number(table, "LeftCell", 1, MAX_COLS)? != 1 {
        return Err(error("offset tables (TopCell/LeftCell) are not supported"));
    }
    if let Some(height) = dimension(table, "DefaultRowHeight") {
        sheet.default_row_height = height;
    }
    if let Some(width) = dimension(table, "DefaultColumnWidth") {
        sheet.default_col_width = width;
    }
    let mut col_next = 1;
    for column in table.kids("Column") {
        let start = index(column, col_next, MAX_COLS)?;
        let span = number(column, "Span", 0, MAX_COLS - 1)?;
        let end = start.checked_add(span).filter(|c| *c < MAX_COLS).ok_or_else(|| error("column span exceeds worksheet bounds"))?;
        charge(budget, u64::from(end - start + 1))?;
        for col in start..=end {
            sheet.cols.insert(
                col,
                LineInfo {
                    size: dimension(column, "Width"),
                    hidden: column.flag("Hidden", false),
                    style: Some(style_id(column, styles, default, report)),
                    ..Default::default()
                },
            );
        }
        col_next = end + 2;
    }
    let mut occupied = HashSet::new();
    let mut row_next = 1;
    for row in table.kids("Row") {
        let r = index(row, row_next, MAX_ROWS)?;
        charge(budget, 1)?;
        row_next = r + 2;
        let row_style = style_id(row, styles, default, report);
        sheet.rows.insert(
            r,
            LineInfo {
                size: dimension(row, "Height"),
                hidden: row.flag("Hidden", false),
                style: row.attr("StyleID").map(|_| row_style),
                custom: row.attr("Height").is_some(),
                ..Default::default()
            },
        );
        let mut cell_next = 1;
        for cell in row.kids("Cell") {
            let c = index(cell, cell_next, MAX_COLS)?;
            let across = number(cell, "MergeAcross", 0, MAX_COLS - 1)?;
            let down = number(cell, "MergeDown", 0, MAX_ROWS - 1)?;
            let right = c.checked_add(across).filter(|v| *v < MAX_COLS).ok_or_else(|| error("merge exceeds worksheet width"))?;
            let bottom = r.checked_add(down).filter(|v| *v < MAX_ROWS).ok_or_else(|| error("merge exceeds worksheet height"))?;
            charge(budget, u64::from(across + 1) * u64::from(down + 1))?;
            let inherited = if row.attr("StyleID").is_some() { row_style } else { sheet.cols.get(&c).and_then(|c| c.style).unwrap_or(default) };
            let style = style_id(cell, styles, inherited, report);
            for rr in r..=bottom {
                for cc in c..=right {
                    let at = CellRef::new(rr, cc);
                    if !occupied.insert(at) {
                        return Err(error("overlapping cells or merges"));
                    }
                    sheet.set_cell(at, Cell { style, ..Default::default() });
                }
            }
            let mut value = read_value(cell.child("Data"))?;
            if cell.attr("Formula").is_some() {
                report.warn("XML Spreadsheet formulas use their saved results and will not recalculate; formula import is not supported yet");
                if cell.child("Data").is_none() {
                    value = Value::Text("[Unavailable XML formula result]".into());
                }
            }
            sheet.set_cell(CellRef::new(r, c), Cell { value, style, ..Default::default() });
            if across != 0 || down != 0 {
                sheet.merges.push(RangeRef::new(CellRef::new(r, c), CellRef::new(bottom, right)));
            }
            cell_next = right + 2;
        }
    }
    Ok(())
}

fn read_value(data: Option<&El>) -> Result<Value, IoError> {
    let Some(data) = data else { return Ok(Value::Empty) };
    let text = &data.text;
    Ok(match data.attr("Type").unwrap_or("String") {
        "String" => {
            if !data.children.is_empty() {
                return Err(error("rich text inside Data is not supported"));
            }
            Value::Text(text.as_str().into())
        }
        "Number" => Value::Number(text.trim().parse::<f64>().ok().filter(|n| n.is_finite()).ok_or_else(|| error("invalid numeric value"))?),
        "Boolean" => Value::Bool(match text.trim() {
            "1" | "true" => true,
            "0" | "false" => false,
            _ => return Err(error("invalid boolean value")),
        }),
        "Error" => Value::Error(CellError::parse(text.trim()).ok_or_else(|| error("unknown cell error"))?),
        "DateTime" => {
            let dt = NaiveDateTime::parse_from_str(text.trim(), "%Y-%m-%dT%H:%M:%S%.f").map_err(error)?;
            let day = gridcraft_core::date::serial_from_ymd(DateSystem::D1900, i64::from(dt.year()), i64::from(dt.month()), i64::from(dt.day()))
                .ok_or_else(|| error("date outside the spreadsheet date range"))?;
            Value::Number(day + (f64::from(dt.num_seconds_from_midnight()) + f64::from(dt.nanosecond()) / 1e9) / 86400.0)
        }
        _ => return Err(error("unsupported Data type")),
    })
}
