use docx_rs::{Docx, read_docx};
use docx_tool::sql_parser::DocxDb;
use gluesql::core::store::GStore;
use gluesql::prelude::Glue;
use gluesql::{
    core::{
        ast::ColumnDef,
        data::Value,
        data::{Schema, SchemaParseError},
        error::FetchError,
        store::{DataRow, RowIter, Store},
    },
    prelude::{DataType, Error, Key, Result},
};

#[tokio::main(flavor = "current_thread")]
async fn main() -> Result<(), Box<dyn std::error::Error>> {
    // 读取docx
    let docx_content = include_bytes!("../../asset/接口.docx");

    let mut docx: Docx = read_docx(docx_content)?;
    let store = DocxDb {
        docx: &mut docx.document,
    };
    let mut glue: Glue<DocxDb> = Glue::new(store);

    let result = glue
        .execute("update tables set justification = 'center', width = 100, width_type = 'dxa'")
        .await?;
    for item in result {
        println!("{:?}", item);
    }

    println!("{:?}", serde_json::to_string(&glue.storage.docx));

    let path = std::path::Path::new("out.docx");
    let file = std::fs::File::create(path)?;
    docx.build().pack(file);

    Ok(())
}
