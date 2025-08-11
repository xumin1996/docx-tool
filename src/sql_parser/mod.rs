use async_trait::async_trait;
use docx_rs::{Docx, read_docx};
use gluesql::{
    core::{
        data::Schema,
        error::FetchError,
        store::{DataRow, RowIter, Store},
    },
    prelude::{Error, Key, Result},
};

struct DocxDb {}

#[async_trait(?Send)]
impl Store for DocxDb {
    async fn fetch_schema(&self, table_name: &str) -> Result<Option<Schema>> {
        Result::Ok(Option::None)
    }

    async fn fetch_all_schemas(&self) -> Result<Vec<Schema>> {
        Result::Ok(vec![])
    }

    async fn fetch_data(&self, table_name: &str, key: &Key) -> Result<Option<DataRow>> {
        Result::Ok(Option::None)
    }

    async fn scan_data<'a>(&'a self, table_name: &str) -> Result<RowIter<'a>> {
        Result::Err(Error::Fetch(FetchError::TableNotFound("".to_string())))
    }
}
