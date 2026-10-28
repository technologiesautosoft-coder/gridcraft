//! Synthetic XML Spreadsheet 2003 fixtures, never real operational workbooks.

use gridcraft_core::{CellRef, Value};

use crate::{Format, read_spreadsheetml, sniff};

fn book(body: &str) -> String {
    format!(
        r#"<?xml version="1.0"?><Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet" xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet">{body}</Workbook>"#
    )
}

fn sheet(rows: &str) -> String {
    book(&format!(r#"<Worksheet ss:Name="Report"><Table>{rows}</Table></Worksheet>"#))
}

#[test]
fn spreadsheetml_sniff_requires_the_root_namespace_and_handles_unicode_boms() {
    let xml = sheet(r#"<Row><Cell><Data ss:Type="String">Crème &amp; café</Data></Cell></Row>"#);
    for bytes in [
        xml.as_bytes().to_vec(),
        [vec![0xEF, 0xBB, 0xBF], xml.as_bytes().to_vec()].concat(),
        [vec![0xFF, 0xFE], xml.encode_utf16().flat_map(u16::to_le_bytes).collect()].concat(),
        [vec![0xFE, 0xFF], xml.encode_utf16().flat_map(u16::to_be_bytes).collect()].concat(),
    ] {
        assert_eq!(sniff(&bytes), Format::SpreadsheetXml);
        let (wb, _) = read_spreadsheetml(&bytes).unwrap();
        assert_eq!(wb.sheets[0].value(CellRef::new(0, 0)), Value::Text("Crème & café".into()));
    }
    for text in [
        r#"<Workbook xmlns="urn:unrelated"><Worksheet/></Workbook>"#,
        "a,b\n<Workbook>,other",
        r#"<other><Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet"/></other>"#,
    ] {
        assert_ne!(sniff(text.as_bytes()), Format::SpreadsheetXml);
    }
    let prefixed = r#"<s:Workbook xmlns:s="urn:schemas-microsoft-com:office:spreadsheet"><s:Worksheet s:Name="One"><s:Table/></s:Worksheet><s:Worksheet s:Name="Two"><s:Table/></s:Worksheet></s:Workbook>"#;
    assert_eq!(read_spreadsheetml(prefixed.as_bytes()).unwrap().0.sheets.len(), 2);
}

#[test]
fn spreadsheetml_preserves_typed_values_and_warns_when_formula_caches_are_used() {
    let xml = sheet(
        r#"<Row><Cell><Data ss:Type="Number">12.25</Data></Cell>
        <Cell><Data ss:Type="Boolean">0</Data></Cell><Cell><Data ss:Type="Error">#DIV/0!</Data></Cell>
        <Cell><Data ss:Type="DateTime">2026-10-16T12:00:00.000</Data></Cell>
        <Cell ss:Formula="=RC[-4]*2"><Data ss:Type="Number">24.5</Data></Cell>
        <Cell ss:Formula="=UNKNOWN()"/></Row>"#,
    );
    let (wb, report) = read_spreadsheetml(xml.as_bytes()).unwrap();
    let sh = &wb.sheets[0];
    assert_eq!(sh.value(CellRef::new(0, 0)), Value::Number(12.25));
    assert_eq!(sh.value(CellRef::new(0, 1)), Value::Bool(false));
    assert_eq!(sh.value(CellRef::new(0, 2)), Value::Error(gridcraft_core::CellError::Div0));
    let date = gridcraft_core::date::serial_from_ymd(gridcraft_core::DateSystem::D1900, 2026, 10, 16).unwrap() + 0.5;
    assert_eq!(sh.value(CellRef::new(0, 3)), Value::Number(date));
    assert_eq!(sh.value(CellRef::new(0, 4)), Value::Number(24.5));
    assert!(sh.cell(CellRef::new(0, 4)).unwrap().formula.is_none());
    assert_eq!(sh.value(CellRef::new(0, 5)), Value::Text("[Unavailable XML formula result]".into()));
    assert!(report.warnings.iter().any(|w| w.contains("will not recalculate")));
}

#[test]
fn spreadsheetml_reads_dimensions_in_points_and_resolves_parent_styles() {
    let xml = book(
        r##"<Styles><Style ss:ID="Default"><Font ss:Size="14"/></Style>
        <Style ss:ID="Child" ss:Parent="Parent"><NumberFormat ss:Format="0.000"/></Style>
        <Style ss:ID="Parent"><Font ss:Bold="1"/><Interior ss:Pattern="Solid" ss:Color="#FFFF00"/></Style></Styles>
        <Worksheet ss:Name="Report"><Table ss:DefaultRowHeight="21">
        <Column ss:Index="2" ss:Span="1" ss:Width="90" ss:StyleID="Child"/>
        <Row ss:Index="3" ss:Height="30"><Cell ss:Index="2"><Data ss:Type="Number">10</Data></Cell></Row>
        </Table></Worksheet>"##,
    );
    let (wb, _) = read_spreadsheetml(xml.as_bytes()).unwrap();
    let sh = &wb.sheets[0];
    assert_eq!(sh.default_row_height, 28.0);
    assert_eq!(sh.rows.get(&2).unwrap().size, Some(40.0));
    assert_eq!(sh.cols.get(&1).unwrap().size, Some(120.0));
    assert_eq!(sh.cols.get(&2).unwrap().size, Some(120.0));
    let style = wb.styles.get(sh.style_id(CellRef::new(2, 1)));
    assert!(style.font.bold);
    assert_eq!(style.font.size, 14.0);
    assert_eq!(style.fill.fg, gridcraft_model::Color::Rgb(0xFFFF00));
    assert_eq!(style.num_fmt.0.as_ref(), "0.000");
}

#[test]
fn spreadsheetml_rejects_malformed_documents_and_hostile_expansion() {
    for rows in [
        r#"<Row ss:Index="0"/>"#,
        r#"<Row><Cell ss:MergeAcross="4294967295"/></Row>"#,
        r#"<Row><Cell ss:MergeDown="1048575" ss:MergeAcross="16383"/></Row>"#,
        r#"<Row><Cell ss:MergeDown="1"/></Row><Row><Cell/></Row>"#,
        r#"<Row><Cell ss:Index="4"/><Cell ss:Index="3"/></Row>"#,
        r#"<Row><Cell><Data ss:Type="Number">NaN</Data></Cell></Row>"#,
    ] {
        assert!(read_spreadsheetml(sheet(rows).as_bytes()).is_err(), "accepted {rows}");
    }
    let valid = sheet("<Row/>");
    for xml in [
        valid.trim_end_matches("</Workbook>").to_string(),
        format!("{valid}{valid}"),
        valid.replace("<Row/>", "<Row></Cell>"),
        valid.replace("<Row/>", &format!("{}{}", "<a>".repeat(65), "</a>".repeat(65))),
        valid.replace("<?xml version=\"1.0\"?>", r#"<!DOCTYPE Workbook [<!ENTITY secret SYSTEM "file:///etc/passwd">]>"#),
        book(r#"<Styles><Style ss:ID="Default" ss:Parent="Default"/></Styles><Worksheet ss:Name="Report"/>"#),
    ] {
        assert!(read_spreadsheetml(xml.as_bytes()).is_err(), "accepted malformed XML");
    }
}

#[test]
fn spreadsheetml_rejects_namespace_declaration_bombs() {
    let declarations: String = (0..10_000).map(|i| format!(" xmlns:n{i}=\"urn:{i}\"")).collect();
    let xml = sheet("<Row/>").replace("<Workbook ", &format!("<Workbook{declarations} "));
    assert!(read_spreadsheetml(xml.as_bytes()).is_err());
}
