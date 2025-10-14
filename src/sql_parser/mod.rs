use std::{collections::HashMap, str::FromStr};

use async_trait::async_trait;
use docx_rs::{
    BuildXML, Document, DocumentChild, Docx, Justification, TableAlignmentType, TableChild,
    WidthType, read_docx,
};
use futures::stream::{self, StreamExt};
use gluesql::{
    core::{
        ast::ColumnDef,
        data::{Schema, SchemaParseError, Value},
        error::FetchError,
        store::{
            AlterTable, CustomFunction, CustomFunctionMut, DataRow, Index, IndexMut, Metadata,
            RowIter, Store, StoreMut, Transaction,
        },
    },
    prelude::{DataType, Error, Key, Result},
};
use sha2::{Digest, Sha256};
use std::mem;

pub mod cell;
pub mod tables;

pub struct DocxDb<'a> {
    pub docx: &'a mut Document,
    tables: tables::Tables,
    cell: cell::Cell,
}

impl<'a> DocxDb<'a> {
    pub fn new(docx: &mut Document) -> DocxDb {
        DocxDb {
            docx: docx,
            tables: tables::Tables,
            cell: cell::Cell,
        }
    }
}

#[async_trait(?Send)]
impl<'b> Store for DocxDb<'b> {
    async fn fetch_schema(&self, table_name: &str) -> Result<Option<Schema>> {
        let schemas = self.fetch_all_schemas().await;

        if let Ok(schema_list) = schemas {
            let schema_op = schema_list
                .iter()
                .filter(|item| item.table_name == table_name)
                .map(|item| item.clone())
                .nth(0);
            return Ok(schema_op);
        } else {
            return Result::Err(Error::Schema(SchemaParseError::CannotParseDDL));
        }
    }

    async fn fetch_all_schemas(&self) -> Result<Vec<Schema>> {
        let mut schemas: Vec<Schema> = Vec::new();
        schemas.extend(self.tables.fetch_all_schemas());
        schemas.extend(self.cell.fetch_all_schemas());
        Result::Ok(schemas)
    }

    async fn fetch_data(&self, table_name: &str, key: &Key) -> Result<Option<DataRow>> {
        // 查找
        if self.tables.table_name() == table_name {
            return self.tables.fetch_data(self.docx, key).await;
        }
        if self.cell.table_name() == table_name {
            return self.cell.fetch_data(self.docx, key).await;
        }

        return Result::Ok(None);
    }

    // todo 修改为stream格式
    async fn scan_data<'a>(&'a self, table_name: &str) -> Result<RowIter<'a>> {
        // 查找
        if self.tables.table_name() == table_name {
            return self.tables.scan_data(self.docx).await;
        }
        if self.cell.table_name() == table_name {
            return self.cell.scan_data(self.docx).await;
        }

        return Ok(Box::pin(stream::iter(vec![])));
    }
}

impl<'b> Index for DocxDb<'b> {}
impl<'b> Metadata for DocxDb<'b> {}
impl<'b> CustomFunction for DocxDb<'b> {}

#[async_trait(?Send)]
impl<'b> StoreMut for DocxDb<'b> {
    async fn insert_schema(&mut self, _schema: &Schema) -> Result<()> {
        let msg = "[Storage] StoreMut::insert_schema is not supported".to_owned();

        Err(Error::StorageMsg(msg))
    }

    async fn delete_schema(&mut self, _table_name: &str) -> Result<()> {
        let msg = "[Storage] StoreMut::delete_schema is not supported".to_owned();

        Err(Error::StorageMsg(msg))
    }

    async fn append_data(&mut self, _table_name: &str, _rows: Vec<DataRow>) -> Result<()> {
        let msg = "[Storage] StoreMut::append_data is not supported".to_owned();

        Err(Error::StorageMsg(msg))
    }

    async fn insert_data(&mut self, table_name: &str, _rows: Vec<(Key, DataRow)>) -> Result<()> {
        // 查找
        if self.tables.table_name() == table_name {
            return self.tables.insert_data(self.docx, _rows).await;
        }
        if self.cell.table_name() == table_name {
            return self.cell.insert_data(self.docx, _rows).await;
        }

        Ok(())
    }

    async fn delete_data(&mut self, _table_name: &str, _keys: Vec<Key>) -> Result<()> {
        let msg = "[Storage] StoreMut::delete_data is not supported".to_owned();

        Err(Error::StorageMsg(msg))
    }
}
impl<'b> IndexMut for DocxDb<'b> {}
impl<'b> AlterTable for DocxDb<'b> {}
impl<'b> Transaction for DocxDb<'b> {}
impl<'b> CustomFunctionMut for DocxDb<'b> {}

#[test]
pub fn to_xml() {
    // 读取docx
    let docx_content = include_bytes!("../../asset/测试.docx");
    let mut docx: Docx = read_docx(docx_content).unwrap();
    let docx_json = docx.json();
    println!("{docx_json}");
}
