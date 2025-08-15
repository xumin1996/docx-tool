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

    let docx: Docx = read_docx(docx_content)?;
    let store = DocxDb {
        docx: docx.document,
    };
    let mut glue: Glue<DocxDb> = Glue::new(store);

    let result = glue
        .execute(
            "select hash, row_number, column_number, json_content, 1+1 as cal_number from tables limit 1",
        )
        .await?;
    for item in result {
        println!("{:?}", item);
    }

    Ok(())
}
