use std::{collections::HashMap, iter, str::FromStr};

use async_trait::async_trait;
use docx_rs::{
    BorderType, Document, DocumentChild, Docx, Justification, Paragraph, ParagraphChild, RunChild,
    TableAlignmentType, TableCellBorder, TableCellBorderPosition, TableCellContent,
    TableCellProperty, TableChild, TableRowChild, WidthType, border_position, read_docx,
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

pub struct Cell;

impl Cell {
    pub fn table_name(&self) -> String {
        "cell".to_string()
    }

    pub fn fetch_all_schemas(&self) -> Vec<Schema> {
        vec![Schema {
            table_name: "cell".to_string(),
            column_defs: Some(vec![
                ColumnDef {
                    name: "hash".to_string(),
                    data_type: DataType::Text,
                    nullable: false,
                    default: None,
                    unique: None,
                    comment: Some("cell的哈希".to_string()),
                },
                ColumnDef {
                    name: "table_hash".to_string(),
                    data_type: DataType::Text,
                    nullable: false,
                    default: None,
                    unique: None,
                    comment: Some("表格的哈希".to_string()),
                },
                ColumnDef {
                    name: "content".to_string(),
                    data_type: DataType::Text,
                    nullable: true,
                    default: None,
                    unique: None,
                    comment: Some("cell内容".to_string()),
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
                    comment: Some("cell的json形式".to_string()),
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
        let mut cells = Vec::new();
        for doc_child in &docx.children {
            if let DocumentChild::Table(t_box) = doc_child {
                let table_json_str = serde_json::to_string(t_box).unwrap_or("".to_string());
                let mut hasher = Sha256::new();
                hasher.update(table_json_str.as_bytes());
                let result = hasher.finalize();
                let table_hash_hex = hex::encode(result);

                // 遍历cell
                for row in &t_box.rows {
                    if let TableChild::TableRow(table_row) = row {
                        for cell in &table_row.cells {
                            if let TableRowChild::TableCell(table_cell) = cell {
                                // cell的文本内容
                                let runs = table_cell
                                    .children
                                    .iter()
                                    .flat_map(|item: &TableCellContent| {
                                        if let TableCellContent::Paragraph(paragraph) = item {
                                            paragraph.children.iter()
                                        } else {
                                            [].iter()
                                        }
                                    })
                                    .flat_map(|item| {
                                        if let ParagraphChild::Run(run) = item {
                                            run.children.iter()
                                        } else {
                                            [].iter()
                                        }
                                    })
                                    .map(|item| {
                                        if let RunChild::Text(run_text) = item {
                                            run_text.text.clone()
                                        } else {
                                            "".to_string()
                                        }
                                    })
                                    .collect::<Vec<String>>();
                                let content = runs.join("");

                                let table_json_str =
                                    serde_json::to_string(table_cell).unwrap_or("".to_string());
                                let mut hasher = Sha256::new();
                                hasher.update(table_json_str.as_bytes());
                                let result = hasher.finalize();
                                let cell_hash_hex = hex::encode(result);

                                // 使用json读取属性
                                let property_value: serde_json::Value =
                                    serde_json::to_value(&table_cell.property)
                                        .unwrap_or(serde_json::Value::Null);

                                let key = Key::Str(cell_hash_hex.clone());
                                let mut hm: HashMap<String, Value> = HashMap::new();
                                hm.insert("hash".to_string(), Value::Str(cell_hash_hex.clone()));
                                hm.insert(
                                    "table_hash".to_string(),
                                    Value::Str(table_hash_hex.clone()),
                                );
                                hm.insert("content".to_string(), Value::Str(content.clone()));
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
                                cells.push(Ok((key, data_row)));
                            }
                        }
                    }
                }
            }
        }
        return Ok(Box::pin(stream::iter(cells)));
    }

    pub async fn insert_data(&self, docx: &mut Document, _rows: Vec<(Key, DataRow)>) -> Result<()> {
        // 查找
        for doc_child in &mut docx.children {
            if let DocumentChild::Table(t_box) = doc_child {
                // 遍历cell
                for row in &mut t_box.rows {
                    if let TableChild::TableRow(table_row) = row {
                        for cell in &mut table_row.cells {
                            let TableRowChild::TableCell(table_cell) = cell;
                            let cell_json_str =
                                serde_json::to_string(&table_cell).unwrap_or("".to_string());
                            let mut hasher = Sha256::new();
                            hasher.update(cell_json_str.as_bytes());
                            let result = hasher.finalize();
                            let cell_hash_hex = hex::encode(result);
                            let hash_key = Key::Str(cell_hash_hex);

                            for row in &_rows {
                                if row.0 == hash_key {
                                    if let DataRow::Map(kvs) = &row.1 {
                                        for kv in kvs.iter() {
                                            if kv.0 == "width" {
                                                if let Value::U32(width) = kv.1 {
                                                    // 使用json读取属性
                                                    let property_value: serde_json::Value =
                                                        serde_json::to_value(&table_cell.property)
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
                                                        .and_then(|item| {
                                                            WidthType::from_str(item).ok()
                                                        })
                                                        .unwrap_or(WidthType::Unsupported);

                                                    let property =
                                                        mem::take(&mut table_cell.property);
                                                    table_cell.property = property
                                                        .width(*width as usize, pre_width_type);
                                                }
                                            }
                                            if kv.0 == "width_type" {
                                                if let Value::Str(width_type) = kv.1 {
                                                    // 使用json读取属性
                                                    let property_value: serde_json::Value =
                                                        serde_json::to_value(&table_cell.property)
                                                            .unwrap_or(serde_json::Value::Null);
                                                    let pre_width = property_value
                                                        .get("width")
                                                        .and_then(|item| item.get("width"))
                                                        .and_then(|item| item.as_u64())
                                                        .and_then(|item| Some(item as usize))
                                                        .unwrap_or(0usize);
                                                    let pre_width_type =
                                                        WidthType::from_str(width_type)
                                                            .ok()
                                                            .unwrap_or(WidthType::Auto);

                                                    let property =
                                                        mem::take(&mut table_cell.property);
                                                    table_cell.property =
                                                        property.width(pre_width, pre_width_type);
                                                }
                                            }
                                            if kv.0 == "borders_top" {
                                                if let Value::Str(border_value) = kv.1 {
                                                    let property =
                                                        mem::take(&mut table_cell.property);
                                                    table_cell.property = self.set_border(
                                                        property,
                                                        border_value,
                                                        TableCellBorderPosition::Top,
                                                    );
                                                }
                                            }
                                            if kv.0 == "borders_left" {
                                                if let Value::Str(border_value) = kv.1 {
                                                    let property =
                                                        mem::take(&mut table_cell.property);
                                                    table_cell.property = self.set_border(
                                                        property,
                                                        border_value,
                                                        TableCellBorderPosition::Left,
                                                    );
                                                }
                                            }
                                            if kv.0 == "borders_bottom" {
                                                if let Value::Str(border_value) = kv.1 {
                                                    let property =
                                                        mem::take(&mut table_cell.property);
                                                    table_cell.property = self.set_border(
                                                        property,
                                                        border_value,
                                                        TableCellBorderPosition::Bottom,
                                                    );
                                                }
                                            }
                                            if kv.0 == "borders_right" {
                                                if let Value::Str(border_value) = kv.1 {
                                                    let property =
                                                        mem::take(&mut table_cell.property);
                                                    table_cell.property = self.set_border(
                                                        property,
                                                        border_value,
                                                        TableCellBorderPosition::Right,
                                                    );
                                                }
                                            }
                                            if kv.0 == "borders_inside_h" {
                                                if let Value::Str(border_value) = kv.1 {
                                                    let property =
                                                        mem::take(&mut table_cell.property);
                                                    table_cell.property = self.set_border(
                                                        property,
                                                        border_value,
                                                        TableCellBorderPosition::InsideH,
                                                    );
                                                }
                                            }
                                            if kv.0 == "borders_inside_v" {
                                                if let Value::Str(border_value) = kv.1 {
                                                    let property =
                                                        mem::take(&mut table_cell.property);
                                                    table_cell.property = self.set_border(
                                                        property,
                                                        border_value,
                                                        TableCellBorderPosition::InsideV,
                                                    );
                                                }
                                            }
                                        }
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

    fn set_border(
        &self,
        property: TableCellProperty,
        border_value: &String,
        border_position: TableCellBorderPosition,
    ) -> TableCellProperty {
        // 使用json读取属性
        let value: serde_json::Value =
            serde_json::from_str(&border_value).unwrap_or(serde_json::Value::Null);

        let mut table_border = TableCellBorder::new(border_position);

        // 颜色
        if let Some(color) = value.get("color").and_then(|item| item.as_str()) {
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

        return property.set_border(table_border);
    }
}
