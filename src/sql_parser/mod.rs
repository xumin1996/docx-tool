use std::collections::HashMap;

use async_trait::async_trait;
use docx_rs::{Document, DocumentChild, Docx, read_docx};
use futures::stream::{self, StreamExt};
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
use sha2::{Digest, Sha256};

struct DocxDb {
    docx: Document,
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
            table_name: "table".to_string(),
            column_defs: Some(vec![ColumnDef {
                name: "hash".to_string(),
                data_type: DataType::Text,
                nullable: false,
                default: None,
                unique: None,
                comment: Some("table".to_string()),
            }]),
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
        if let Ok(rowIter) = self.scan_data(table_name).await {
            while let row_op = rowIter.next().await {
                if let Some(row_result) = row_op {
                    if let Ok(row) = row_result {
                        if row.0 == *key {
                            return Ok(Some(row.1.clone()));
                        }
                    }
                }
            }
        }

        return Result::Ok(None);
    }

    async fn scan_data<'a>(&'a self, table_name: &str) -> Result<RowIter<'a>> {
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
                let data_row = DataRow::Map(hm);

                return Ok(Box::pin(stream::iter(vec![Ok((key, data_row))])));
            }
        }
        Result::Err(Error::Fetch(FetchError::TableNotFound("".to_string())))
    }
}
