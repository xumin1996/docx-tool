use base64::{Engine, engine::general_purpose};
use docx_handlebars::render_handlebars;
use serde_json::json;

fn main() -> Result<(), Box<dyn std::error::Error>> {
    // 读取docx模板文件
    let template_bytes = std::fs::read("assets/model.docx")?;

    // 准备数据
    let data = json!({
        "cheng_xu_miao_shu": [
            "描述内容1",
            "描述内容2",
            "描述内容3"
        ],
        "shu_ru_xiang": [
            {
                "shu_ru_xiang": "旧码",
                "lei_xing": "String1"
            },
            {
                "shu_ru_xiang": "新码",
                "lei_xing": "String2"
            },
        ],
        "image_base64": general_purpose::STANDARD.encode(&std::fs::read("/home/x/Pictures/stream_water_street_1379018_1920x1080.jpg")?)
    });

    // 渲染模板
    let result = render_handlebars(template_bytes, &data)?;

    // 保存结果
    std::fs::write("output.docx", result)?;

    Ok(())
}
