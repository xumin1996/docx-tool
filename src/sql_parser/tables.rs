use std::{collections::HashMap, str::FromStr};

use async_trait::async_trait;
use docx_rs::{
    BorderType, Document, DocumentChild, Docx, Justification, TableAlignmentType, TableBorder,
    TableBorderPosition, TableChild, WidthType, read_docx,
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

pub struct Tables;

impl Tables {
    pub fn table_name(&self) -> String {
        "tables".to_string()
    }

    pub fn fetch_all_schemas(&self) -> Vec<Schema> {
        vec![Schema {
            table_name: "tables".to_string(),
            column_defs: Some(vec![
                ColumnDef {
                    name: "hash".to_string(),
                    data_type: DataType::Text,
                    nullable: false,
                    default: None,
                    unique: None,
                    comment: Some("表格".to_string()),
                },
                ColumnDef {
                    name: "row_number".to_string(),
                    data_type: DataType::Uint32,
                    nullable: false,
                    default: None,
                    unique: None,
                    comment: Some("行数".to_string()),
                },
                ColumnDef {
                    name: "column_number".to_string(),
                    data_type: DataType::Uint32,
                    nullable: false,
                    default: None,
                    unique: None,
                    comment: Some("列数".to_string()),
                },
                ColumnDef {
                    name: "width".to_string(),
                    data_type: DataType::Uint32,
                    nullable: true,
                    default: None,
                    unique: None,
                    comment: Some("表格宽度".to_string()),
                },
                ColumnDef {
                    name: "width_type".to_string(),
                    data_type: DataType::Text,
                    nullable: true,
                    default: None,
                    unique: None,
                    comment: Some("表格宽度类型".to_string()),
                },
                ColumnDef {
                    name: "justification".to_string(),
                    data_type: DataType::Text,
                    nullable: true,
                    default: None,
                    unique: None,
                    comment: Some("对齐方式".to_string()),
                },
                ColumnDef {
                    name: "borders_top".to_string(),
                    data_type: DataType::Text,
                    nullable: true,
                    default: None,
                    unique: None,
                    comment: Some("上边框".to_string()),
                },
                ColumnDef {
                    name: "borders_left".to_string(),
                    data_type: DataType::Text,
                    nullable: true,
                    default: None,
                    unique: None,
                    comment: Some("左边框".to_string()),
                },
                ColumnDef {
                    name: "borders_bottom".to_string(),
                    data_type: DataType::Text,
                    nullable: true,
                    default: None,
                    unique: None,
                    comment: Some("底边框".to_string()),
                },
                ColumnDef {
                    name: "borders_right".to_string(),
                    data_type: DataType::Text,
                    nullable: true,
                    default: None,
                    unique: None,
                    comment: Some("右边框".to_string()),
                },
                ColumnDef {
                    name: "borders_inside_h".to_string(),
                    data_type: DataType::Text,
                    nullable: true,
                    default: None,
                    unique: None,
                    comment: Some("水平内部边框".to_string()),
                },
                ColumnDef {
                    name: "borders_inside_v".to_string(),
                    data_type: DataType::Text,
                    nullable: true,
                    default: None,
                    unique: None,
                    comment: Some("垂直内部边框".to_string()),
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
        }]
    }

    pub async fn fetch_data(&self, docx: &Document, key: &Key) -> Result<Option<DataRow>> {
        // 查找
        if let Ok(mut rowIter) = self.scan_data(docx).await {
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
    pub async fn scan_data<'a>(&self, docx: &Document) -> Result<RowIter<'a>> {
        let mut tables = Vec::new();
        for doc_child in &docx.children {
            if let DocumentChild::Table(t_box) = doc_child {
                let table_json_str = serde_json::to_string(t_box).unwrap_or("".to_string());
                let mut hasher = Sha256::new();
                hasher.update(table_json_str.as_bytes());
                let result = hasher.finalize();
                let hash_hex = hex::encode(result);

                // 表格的行数和列数
                let row_number = t_box.rows.len();
                let column_number = t_box
                    .rows
                    .get(0)
                    .map(|item| {
                        if let TableChild::TableRow(table_row) = item {
                            return table_row.cells.len();
                        } else {
                            return 0;
                        }
                    })
                    .unwrap_or(0);

                // 使用json读取属性
                let property_value: serde_json::Value =
                    serde_json::to_value(&t_box.property).unwrap_or(serde_json::Value::Null);

                let key = Key::Str(hash_hex.clone());
                let mut hm: HashMap<String, Value> = HashMap::new();
                hm.insert("hash".to_string(), Value::Str(hash_hex.clone()));
                hm.insert("row_number".to_string(), Value::I32(row_number as i32));
                hm.insert(
                    "column_number".to_string(),
                    Value::U32(column_number as u32),
                );
                hm.insert(
                    "width".to_string(),
                    Value::U32(
                        property_value
                            .get("width")
                            .and_then(|item| item.get("width"))
                            .and_then(|item| item.as_u64())
                            .and_then(|item| Some(item as u32))
                            .unwrap_or(0u32),
                    ),
                );
                hm.insert(
                    "width_type".to_string(),
                    Value::Str(
                        property_value
                            .get("width")
                            .and_then(|item| item.get("widthType"))
                            .and_then(|item| item.as_str())
                            .unwrap_or("")
                            .to_string(),
                    ),
                );
                hm.insert(
                    "justification".to_string(),
                    Value::Str(
                        property_value
                            .get("justification")
                            .and_then(|item| item.as_str())
                            .unwrap_or("")
                            .to_string(),
                    ),
                );
                hm.insert(
                    "borders_top".to_string(),
                    property_value
                        .get("borders")
                        .and_then(|item| item.get("top"))
                        .and_then(|item| item.as_str())
                        .map(|item| Value::Str(item.to_string()))
                        .unwrap_or(Value::Null),
                );
                hm.insert(
                    "borders_left".to_string(),
                    property_value
                        .get("borders")
                        .and_then(|item| item.get("left"))
                        .and_then(|item| item.as_str())
                        .map(|item| Value::Str(item.to_string()))
                        .unwrap_or(Value::Null),
                );
                hm.insert(
                    "borders_bottom".to_string(),
                    property_value
                        .get("borders")
                        .and_then(|item| item.get("bottom"))
                        .and_then(|item| item.as_str())
                        .map(|item| Value::Str(item.to_string()))
                        .unwrap_or(Value::Null),
                );
                hm.insert(
                    "borders_right".to_string(),
                    property_value
                        .get("borders")
                        .and_then(|item| item.get("right"))
                        .and_then(|item| item.as_str())
                        .map(|item| Value::Str(item.to_string()))
                        .unwrap_or(Value::Null),
                );
                hm.insert(
                    "borders_inside_h".to_string(),
                    property_value
                        .get("borders")
                        .and_then(|item| item.get("insideH"))
                        .and_then(|item| item.as_str())
                        .map(|item| Value::Str(item.to_string()))
                        .unwrap_or(Value::Null),
                );
                hm.insert(
                    "borders_inside_v".to_string(),
                    property_value
                        .get("borders")
                        .and_then(|item| item.get("insideV"))
                        .and_then(|item| item.as_str())
                        .map(|item| Value::Str(item.to_string()))
                        .unwrap_or(Value::Null),
                );

                let data_row = DataRow::Map(hm);
                tables.push(Ok((key, data_row)));
            }
        }
        return Ok(Box::pin(stream::iter(tables)));
    }

    pub async fn insert_data(&self, docx: &mut Document, _rows: Vec<(Key, DataRow)>) -> Result<()> {
        // 查找
        for doc_child in &mut docx.children {
            if let DocumentChild::Table(t_box) = doc_child {
                let table_json_str = serde_json::to_string(&t_box).unwrap_or("".to_string());
                let mut hasher = Sha256::new();
                hasher.update(table_json_str.as_bytes());
                let result = hasher.finalize();
                let hash_hex = hex::encode(result);
                let hash_key = Key::Str(hash_hex);

                for row in &_rows {
                    if row.0 == hash_key {
                        if let DataRow::Map(kvs) = &row.1 {
                            for kv in kvs.iter() {
                                if kv.0 == "width" {
                                    if let Value::U32(width) = kv.1 {
                                        // 使用json读取属性
                                        let property_value: serde_json::Value =
                                            serde_json::to_value(&t_box.property)
                                                .unwrap_or(serde_json::Value::Null);
                                        let pre_width = property_value
                                            .get("width")
                                            .and_then(|item| item.get("width"))
                                            .and_then(|item| item.as_u64())
                                            .and_then(|item| Some(item as usize))
                                            .unwrap_or(0usize);
                                        let pre_width_type = property_value
                                            .get("width")
                                            .and_then(|item| item.get("widthType"))
                                            .and_then(|item| item.as_str())
                                            .and_then(|item| WidthType::from_str(item).ok())
                                            .unwrap_or(WidthType::Unsupported);

                                        let property = mem::take(&mut t_box.property);
                                        t_box.property =
                                            property.width(*width as usize, pre_width_type);
                                    }
                                }
                                if kv.0 == "width_type" {
                                    if let Value::Str(width_type) = kv.1 {
                                        // 使用json读取属性
                                        let property_value: serde_json::Value =
                                            serde_json::to_value(&t_box.property)
                                                .unwrap_or(serde_json::Value::Null);
                                        let pre_width = property_value
                                            .get("width")
                                            .and_then(|item| item.get("width"))
                                            .and_then(|item| item.as_u64())
                                            .and_then(|item| Some(item as usize))
                                            .unwrap_or(0usize);
                                        let pre_width_type = WidthType::from_str(width_type)
                                            .ok()
                                            .unwrap_or(WidthType::Auto);

                                        let property = mem::take(&mut t_box.property);
                                        t_box.property = property.width(pre_width, pre_width_type);
                                    }
                                }
                                if kv.0 == "justification" {
                                    if let Value::Str(prop_value) = kv.1 {
                                        if let Ok(align) = TableAlignmentType::from_str(prop_value)
                                        {
                                            let property = mem::take(&mut t_box.property);
                                            t_box.property = property.align(align);
                                        }
                                    }
                                }
                                if kv.0 == "borders_top" {
                                    if let Value::Str(border_value) = kv.1 {
                                        // 使用json读取属性
                                        let value: serde_json::Value =
                                            serde_json::from_str(&border_value)
                                                .unwrap_or(serde_json::Value::Null);

                                        let mut table_border =
                                            TableBorder::new(TableBorderPosition::Top);

                                        // 颜色
                                        if let Some(color) =
                                            value.get("color").and_then(|item| item.as_str())
                                        {
                                            table_border = table_border.color(color);
                                        }

                                        // 线条宽度
                                        if let Some(size) = value
                                            .get("size")
                                            .and_then(|item| item.as_u64())
                                            .and_then(|item| Some(item as usize))
                                        {
                                            table_border = table_border.size(size);
                                        }

                                        // 线条类型
                                        if let Some(border_type) = value
                                            .get("borderType")
                                            .and_then(|item| item.as_str())
                                            .and_then(|item| BorderType::from_str(item).ok())
                                        {
                                            table_border = table_border.border_type(border_type);
                                        }

                                        let property = mem::take(&mut t_box.property);
                                        t_box.property = property.set_border(table_border);
                                    }
                                }
                                if kv.0 == "borders_left" {
                                    if let Value::Str(border_value) = kv.1 {
                                        // 使用json读取属性
                                        let value: serde_json::Value =
                                            serde_json::from_str(&border_value)
                                                .unwrap_or(serde_json::Value::Null);

                                        let mut table_border =
                                            TableBorder::new(TableBorderPosition::Left);

                                        // 颜色
                                        if let Some(color) =
                                            value.get("color").and_then(|item| item.as_str())
                                        {
                                            table_border = table_border.color(color);
                                        }

                                        // 线条宽度
                                        if let Some(size) = value
                                            .get("size")
                                            .and_then(|item| item.as_u64())
                                            .and_then(|item| Some(item as usize))
                                        {
                                            table_border = table_border.size(size);
                                        }

                                        // 线条类型
                                        if let Some(border_type) = value
                                            .get("borderType")
                                            .and_then(|item| item.as_str())
                                            .and_then(|item| BorderType::from_str(item).ok())
                                        {
                                            table_border = table_border.border_type(border_type);
                                        }

                                        let property = mem::take(&mut t_box.property);
                                        t_box.property = property.set_border(table_border);
                                    }
                                }
                                if kv.0 == "borders_bottom" {
                                    if let Value::Str(border_value) = kv.1 {
                                        // 使用json读取属性
                                        let value: serde_json::Value =
                                            serde_json::from_str(&border_value)
                                                .unwrap_or(serde_json::Value::Null);

                                        let mut table_border =
                                            TableBorder::new(TableBorderPosition::Bottom);

                                        // 颜色
                                        if let Some(color) =
                                            value.get("color").and_then(|item| item.as_str())
                                        {
                                            table_border = table_border.color(color);
                                        }

                                        // 线条宽度
                                        if let Some(size) = value
                                            .get("size")
                                            .and_then(|item| item.as_u64())
                                            .and_then(|item| Some(item as usize))
                                        {
                                            table_border = table_border.size(size);
                                        }

                                        // 线条类型
                                        if let Some(border_type) = value
                                            .get("borderType")
                                            .and_then(|item| item.as_str())
                                            .and_then(|item| BorderType::from_str(item).ok())
                                        {
                                            table_border = table_border.border_type(border_type);
                                        }

                                        let property = mem::take(&mut t_box.property);
                                        t_box.property = property.set_border(table_border);
                                    }
                                }
                                if kv.0 == "borders_right" {
                                    if let Value::Str(border_value) = kv.1 {
                                        // 使用json读取属性
                                        let value: serde_json::Value =
                                            serde_json::from_str(&border_value)
                                                .unwrap_or(serde_json::Value::Null);

                                        let mut table_border =
                                            TableBorder::new(TableBorderPosition::Right);

                                        // 颜色
                                        if let Some(color) =
                                            value.get("color").and_then(|item| item.as_str())
                                        {
                                            table_border = table_border.color(color);
                                        }

                                        // 线条宽度
                                        if let Some(size) = value
                                            .get("size")
                                            .and_then(|item| item.as_u64())
                                            .and_then(|item| Some(item as usize))
                                        {
                                            table_border = table_border.size(size);
                                        }

                                        // 线条类型
                                        if let Some(border_type) = value
                                            .get("borderType")
                                            .and_then(|item| item.as_str())
                                            .and_then(|item| BorderType::from_str(item).ok())
                                        {
                                            table_border = table_border.border_type(border_type);
                                        }

                                        let property = mem::take(&mut t_box.property);
                                        t_box.property = property.set_border(table_border);
                                    }
                                }
                                if kv.0 == "borders_inside_h" {
                                    if let Value::Str(border_value) = kv.1 {
                                        // 使用json读取属性
                                        let value: serde_json::Value =
                                            serde_json::from_str(&border_value)
                                                .unwrap_or(serde_json::Value::Null);

                                        let mut table_border =
                                            TableBorder::new(TableBorderPosition::InsideH);

                                        // 颜色
                                        if let Some(color) =
                                            value.get("color").and_then(|item| item.as_str())
                                        {
                                            table_border = table_border.color(color);
                                        }

                                        // 线条宽度
                                        if let Some(size) = value
                                            .get("size")
                                            .and_then(|item| item.as_u64())
                                            .and_then(|item| Some(item as usize))
                                        {
                                            table_border = table_border.size(size);
                                        }

                                        // 线条类型
                                        if let Some(border_type) = value
                                            .get("borderType")
                                            .and_then(|item| item.as_str())
                                            .and_then(|item| BorderType::from_str(item).ok())
                                        {
                                            table_border = table_border.border_type(border_type);
                                        }

                                        let property = mem::take(&mut t_box.property);
                                        t_box.property = property.set_border(table_border);
                                    }
                                }
                                if kv.0 == "borders_inside_v" {
                                    if let Value::Str(border_value) = kv.1 {
                                        // 使用json读取属性
                                        let value: serde_json::Value =
                                            serde_json::from_str(&border_value)
                                                .unwrap_or(serde_json::Value::Null);

                                        let mut table_border =
                                            TableBorder::new(TableBorderPosition::InsideV);

                                        // 颜色
                                        if let Some(color) =
                                            value.get("color").and_then(|item| item.as_str())
                                        {
                                            table_border = table_border.color(color);
                                        }

                                        // 线条宽度
                                        if let Some(size) = value
                                            .get("size")
                                            .and_then(|item| item.as_u64())
                                            .and_then(|item| Some(item as usize))
                                        {
                                            table_border = table_border.size(size);
                                        }

                                        // 线条类型
                                        if let Some(border_type) = value
                                            .get("borderType")
                                            .and_then(|item| item.as_str())
                                            .and_then(|item| BorderType::from_str(item).ok())
                                        {
                                            table_border = table_border.border_type(border_type);
                                        }

                                        let property = mem::take(&mut t_box.property);
                                        t_box.property = property.set_border(table_border);
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }

        Ok(())
    }
}
