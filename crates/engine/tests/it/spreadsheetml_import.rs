//! Synthetic SpreadsheetML 2003 exports, exercising the same dispatch as File > Open.

use gridcraft_core::{CellRef, RangeRef, Value};
use gridcraft_engine::io;
use serde_json::json;

#[test]
fn xml_import_saves_as_xlsx_and_does_not_overwrite_the_source() {
    let xml = br#"<Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet" xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet"><Worksheet ss:Name="Report"><Table><Row><Cell><Data ss:Type="Number">42</Data></Cell></Row></Table></Worksheet></Workbook>"#;
    for name in ["report.xls", "report.xml", "misnamed.csv"] {
        let mut session = gridcraft_engine::Session::new();
        session.new_workbook();
        session.execute("file.open", json!({"name":name,"base64":io::base64_encode(xml)})).unwrap();
        assert!(session.doc().unwrap().path.is_none());
        assert!(session.doc().unwrap().display_title().ends_with("(imported).xlsx"));
        assert!(session.execute("file.saveBytes", json!({"format":"xls"})).is_err());
        assert!(session.execute("file.saveBytes", json!({"format":"xlsx"})).is_ok());
        session.execute("cell.set", json!({"cell":"A1","input":"Keep edited data"})).unwrap();
        assert!(session.execute("file.open", json!({"name":"broken.xls","base64":io::base64_encode(&xml[..xml.len()-10])})).is_err());
        assert_eq!(session.doc().unwrap().wb.sheets[0].value(CellRef::new(0, 0)), Value::Text("Keep edited data".into()));
        assert!(session.doc().unwrap().is_dirty());
    }
}

#[test]
fn xml_named_xls_opens_as_a_formatted_workbook_instead_of_raw_xml() {
    let xml = br##"<?xml version="1.0"?>
<?mso-application progid="Excel.Sheet"?>
<Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet" xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet">
<Styles><Style ss:ID="Default"><Font ss:FontName="Calibri" ss:Size="14"/></Style>
<Style ss:ID="Heading"><Font ss:Bold="1"/><Interior ss:Color="#F4A06F" ss:Pattern="Solid"/></Style>
<Style ss:ID="Number"><NumberFormat ss:Format="0.00"/></Style></Styles>
<Worksheet ss:Name="Production"><Table ss:DefaultRowHeight="20">
<Column ss:Width="90"/><Column ss:Width="120"/>
<Row ss:Height="26"><Cell ss:StyleID="Heading" ss:MergeAcross="1"><Data ss:Type="String">Daily report</Data></Cell></Row>
<Row><Cell ss:StyleID="Number"><Data ss:Type="Number">42.5</Data></Cell><Cell><Data ss:Type="String">A &amp; B</Data></Cell></Row>
<Row ss:Index="5"><Cell ss:Index="3"><Data ss:Type="Boolean">1</Data></Cell></Row>
</Table></Worksheet></Workbook>"##;
    for name in ["daily.xls", "daily.xml", "misnamed.csv"] {
        let (wb, _) = io::open_bytes(name, xml).unwrap();
        let sheet = wb.sheet(0).unwrap();
        assert_eq!(sheet.name, "Production");
        assert_eq!(sheet.value(CellRef::new(0, 0)), Value::Text("Daily report".into()));
        assert_eq!(sheet.value(CellRef::new(1, 0)), Value::Number(42.5));
        assert_eq!(sheet.value(CellRef::new(1, 1)), Value::Text("A & B".into()));
        assert_eq!(sheet.value(CellRef::new(4, 2)), Value::Bool(true));
        assert_eq!(sheet.merges, vec![RangeRef::parse("A1:B1").unwrap()]);
        let title = wb.styles.get(sheet.style_id(CellRef::new(0, 0)));
        assert!(title.font.bold);
        assert_eq!(title.font.size, 14.0);
        assert_eq!(wb.styles.get(sheet.style_id(CellRef::new(1, 0))).num_fmt.0.as_ref(), "0.00");
        let saved = io::save_bytes(&wb, "daily.xlsx").unwrap();
        let (again, _) = io::open_bytes("daily.xlsx", &saved).unwrap();
        assert_eq!(again.sheets[0].value(CellRef::new(4, 2)), Value::Bool(true));
        assert_eq!(again.sheets[0].merges, sheet.merges);
    }
}
