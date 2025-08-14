use std::collections::HashMap;

use async_trait::async_trait;
use docx_rs::{Document, DocumentChild, Docx, read_docx};
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

pub struct DocxDb {
    pub docx: Document,
}

#[async_trait(?Send)]
impl Store for DocxDb {
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
        Result::Ok(vec![Schema {
            table_name: "doc_table".to_string(),
            column_defs: Some(vec![
                ColumnDef {
                    name: "hash".to_string(),
                    data_type: DataType::Text,
                    nullable: false,
                    default: None,
                    unique: None,
                    comment: Some("table".to_string()),
                },
                ColumnDef {
                    name: "json_content".to_string(),
                    data_type: DataType::Text,
                    nullable: false,
                    default: None,
                    unique: None,
                    comment: Some("表格的json形式".to_string()),
                },
            ]),
            indexes: vec![],
            engine: None,
            foreign_keys: vec![],
            comment: None,
        }])
    }

    async fn fetch_data(&self, table_name: &str, key: &Key) -> Result<Option<DataRow>> {
        // 只支持table
        if table_name != table_name {
            return Result::Ok(None);
        }

        // 查找
        if let Ok(mut rowIter) = self.scan_data(table_name).await {
            while let Some(row_result) = rowIter.next().await {
                if let Ok(row) = row_result {
                    if row.0 == *key {
                        return Ok(Some(row.1.clone()));
                    }
                }
            }
        }

        return Result::Ok(None);
    }

    // todo 修改为stream格式
    async fn scan_data<'a>(&'a self, table_name: &str) -> Result<RowIter<'a>> {
        let mut tables = Vec::new();
        for doc_child in &self.docx.children {
            if let DocumentChild::Table(t_box) = doc_child {
                let table_json_str = serde_json::to_string(t_box).unwrap_or("".to_string());
                let mut hasher = Sha256::new();
                hasher.update(table_json_str.as_bytes());
                let result = hasher.finalize();
                let hash_hex = hex::encode(result);

                let key = Key::Str(hash_hex.clone());
                let mut hm: HashMap<String, Value> = HashMap::new();
                hm.insert("hash".to_string(), Value::Str(hash_hex.clone()));
                hm.insert("json_content".to_string(), Value::Str(table_json_str.clone()));
                let data_row = DataRow::Map(hm);

                tables.push(Ok((key, data_row)));
            }
        }
        return Ok(Box::pin(stream::iter(tables)));
    }
}

impl Index for DocxDb {}
impl Metadata for DocxDb {}
impl CustomFunction for DocxDb {}
impl StoreMut for DocxDb {}
impl IndexMut for DocxDb {}
impl AlterTable for DocxDb {}
impl Transaction for DocxDb {}
impl CustomFunctionMut for DocxDb {}
