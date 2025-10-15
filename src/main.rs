use base64::{Engine, engine::general_purpose};
use bytes::Bytes;
use clap::{Arg, Command};
use docx_handlebars::render_handlebars;
use serde::{Deserialize, Serialize};
use serde_json::Value;
use std::collections::HashMap;

use crate::swagger::*;

mod docx_to_html;
mod swagger;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let matches = Command::new("docx-tools")
        .about("根据json和docx模板生成目标docx的工具")
        .arg(
            Arg::new("swagger")
                .long("swagger")
                .value_parser(clap::value_parser!(String))
                .help("基于swagger生成接口文档，目前支持swagger 2.0"),
        )
        .arg(
            Arg::new("docx-model")
                .long("model")
                .value_parser(clap::value_parser!(String))
                .help("docx的模板路径"),
        )
        .arg(
            Arg::new("model-json")
                .long("json")
                .value_parser(clap::value_parser!(String))
                .help("docx的模板填充的json数据文件路径"),
        )
        .arg(
            Arg::new("output")
                .long("output")
                .value_parser(clap::value_parser!(String))
                .help("输出文件名"),
        )
        .get_matches();

    let mut output_file_name: String = "output.docx".to_string();
    if let Some(output) = matches.get_one::<String>("output") {
        output_file_name = output.clone();
    }

    // 解析swagger并生成文档
    if let Some(swagger_path) = matches.get_one::<String>("swagger") {
        let swagger_bytes = get_file_bytes(&swagger_path)?;

        // 生成docx文件
        let r = parse_swagger_and_gen_docx(&swagger_bytes, &output_file_name);
        if let Err(e) = r {
            println!("parse_swagger_and_gen_docx fail, {e:?}");
        }

        return Ok(());
    }

    // 通用的模板
    if let Some(model_path) = matches.get_one::<String>("docx-model") {
        if let Some(json_path) = matches.get_one::<String>("model-json") {
            let template_bytes = get_file_bytes(&model_path)?;
            let json_bytes = get_file_bytes(&json_path)?;
            let mut value: Value = serde_json::from_slice(&json_bytes)?;

            // 处理图片路径
            image_to_base64(&mut value);

            // 渲染模板
            // println!("{}", serde_json::to_string_pretty(&value)?);
            let result = render_handlebars(template_bytes, &value)?;

            // 保存
            std::fs::write(output_file_name, result)?;

            return Ok(());
        }
    }

    Ok(())
}

fn image_to_base64(value: &mut Value) {
    match value {
        Value::Object(map) => {
            let mut add_items: HashMap<String, Value> = HashMap::new();
            for (k, v) in map.iter_mut() {
                if k.ends_with(".image") {
                    if let Value::String(map_value) = v {
                        let content = get_file_bytes(map_value).unwrap_or(vec![]);
                        *v = Value::String(general_purpose::STANDARD.encode(&content));
                        add_items.insert(
                            k.strip_suffix(".image").unwrap_or(k).to_string(),
                            Value::String(general_purpose::STANDARD.encode(&content)),
                        );
                    }
                }
                image_to_base64(v);
            }
            // 添加
            map.extend(add_items);
        }
        Value::Array(arr) => {
            for ele in arr {
                image_to_base64(ele);
            }
        }
        _ => {}
    }
}

fn get_file_bytes(path: &str) -> Result<Vec<u8>, Box<dyn std::error::Error>> {
    // 判断是网络文件还是本地文件// 创建同步客户端
    if path.starts_with("http") {
        let response = ureq::get(path).call()?.body_mut().read_to_vec()?;
        Ok(response)
    } else {
        // 普通文件
        let file_bytes = std::fs::read(path)?;

        return Ok(file_bytes);
    }
}
