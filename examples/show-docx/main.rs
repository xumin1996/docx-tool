use base64::{Engine, engine::general_purpose};
use bytes::Bytes;
use clap::{Arg, Command};
use docx_handlebars::render_handlebars;
use docx_rs::{read_docx, Docx};
use serde::{Deserialize, Serialize};
use serde_json::Value;
use std::collections::HashMap;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    // 读取docx
    let docx_content = std::fs::read("template/官方template.docx")?;

    let docx = read_docx(&docx_content)?;
    let t = docx.document;
    println!("{:?}", t);
    
    Ok(())
}