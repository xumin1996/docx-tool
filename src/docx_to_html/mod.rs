use docx_rs::{
    BuildXML, Document, DocumentChild, Docx, Justification, TableAlignmentType, TableChild,
    WidthType, read_docx,
};

#[test]
pub fn to_json() {
    // 读取docx
    let docx_content = include_bytes!("../../asset/测试.docx");
    let mut docx: Docx = read_docx(docx_content).unwrap();

    // 遍历
    for child in docx.document.children {
        if let DocumentChild::Paragraph(paragraph) = child {
            println!(
                "{}",
                serde_json::to_string_pretty(&paragraph).unwrap_or("".to_string())
            );
        }
    }
}
