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

#[async_std::main]
async fn main() -> Result<(), Box<dyn std::error::Error>> {
    // 读取docx
    let docx_content = include_bytes!("../../asset/db-test.docx");

    let docx: Docx = read_docx(docx_content)?;
    let store = DocxDb {
        docx: docx.document,
    };
    let mut glue: Glue<DocxDb> = Glue::new(store);

    let result = glue.execute("select * from doc_table;").await?;
    for item in result {
        println!("{:?}", item);
    }

    Ok(())
}
