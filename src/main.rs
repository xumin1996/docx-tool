use base64::{Engine, engine::general_purpose};
use clap::{Arg, ArgAction, Command};
use docx_handlebars::render_handlebars;
use serde::{Deserialize, Serialize};
use serde_json::Value;
use serde_json::json;
use std::collections::HashMap;

const SWAGGER_DOCX_MODEL: &[u8] = include_bytes!("../asset/swagger-model.docx");

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let matches = Command::new("docx-tools")
        .about("根据json和docx模板生成目标docx的工具")
        .arg(
            Arg::new("swagger")
                .long("swagger")
                .value_parser(clap::value_parser!(String))
                .help("基于swagger生成接口文档，目前支持swagger 2.0"),
        )
        .get_matches();

    // 解析swagger并生成文档
    if let Some(swagger_path) = matches.get_one::<String>("swagger") {
        // todo 判断是网络文件还是本地文件// 创建同步客户端
        let client = reqwest::blocking::Client::new();
        // 发送GET请求并获取响应
        let response = client.get(swagger_path).send()?;

        let sw: SwaggerDocument = serde_json::from_slice(response.bytes().unwrap().as_ref())?;
        // println!("{:?}", sw);

        // 生成docx的模板对象
        let mut apis: HashMap<String, Vec<DocxApiInfo>> = HashMap::new();
        for tag in sw.tags {
            apis.insert(tag.name.clone(), vec![]);
        }

        for urls in sw.paths {
            let url = urls.0;
            let method_infos = urls.1;
            for methods in method_infos {
                let method = methods.0;
                let operation = methods.1;
                let doc_api_info = DocxApiInfo {
                    name: operation.summary.clone().unwrap_or("".to_string()),
                    desc: operation.summary.clone().unwrap_or("".to_string()),
                    url: url.clone(),
                    method: method,
                    api_type: "".to_string(),
                    return_type: "*/*".to_string(),
                    query_params: vec![],
                    return_params: vec![],
                };

                // tags
                for tag in operation.tags {
                    if let Some(vec) = apis.get_mut(&tag) {
                        vec.push(doc_api_info.clone());
                    }
                }
            }
        }
        let docx_project = DocxProjectInfo {
            name: sw.info.title.clone(),
            apis: apis,
        };
        println!("{:?}", docx_project);

        // 渲染模板
        let result = render_handlebars(
            SWAGGER_DOCX_MODEL.to_vec(),
            &serde_json::to_value(&docx_project)?,
        )?;

        // 保存
        std::fs::write("output.docx", result)?;
    }

    // // 读取docx模板文件
    // let template_bytes = std::fs::read("template/model.docx")?;

    // // 准备数据
    // let data = json!({
    //     "cheng_xu_miao_shu": [
    //         "描述内容1",
    //         "描述内容2",
    //         "描述内容3"
    //     ],
    //     "shu_ru_xiang": [
    //         {
    //             "shu_ru_xiang": "旧码",
    //             "lei_xing": "String1"
    //         },
    //         {
    //             "shu_ru_xiang": "新码",
    //             "lei_xing": "String2"
    //         },
    //     ],
    //     "image_base64": general_purpose::STANDARD.encode(&std::fs::read("/home/x/Pictures/stream_water_street_1379018_1920x1080.jpg")?)
    // });

    // // 渲染模板
    // let result = render_handlebars(template_bytes, &data)?;

    // // 保存结果
    // std::fs::write("output.docx", result)?;

    Ok(())
}

#[derive(Debug, Serialize, Deserialize)]
pub struct SwaggerDocument {
    pub swagger: String,
    pub info: Info,
    pub host: String,
    pub basePath: String,
    pub tags: Vec<Tag>,
    pub paths: HashMap<String, HashMap<String, Operation>>,
    pub securityDefinitions: HashMap<String, SecurityDefinition>,
    pub definitions: HashMap<String, Definition>,
}

#[derive(Debug, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct Info {
    pub version: String,
    pub title: String,
}

#[derive(Debug, Serialize, Deserialize)]
pub struct Tag {
    pub name: String,
    pub description: Option<String>,
}

#[derive(Debug, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct Operation {
    pub tags: Vec<String>,
    pub summary: Option<String>,
    pub operation_id: String,
    pub produces: Vec<String>,
    pub parameters: Option<Vec<Parameter>>,
    pub responses: HashMap<String, Response>,
    pub security: Option<Vec<HashMap<String, Vec<String>>>>,
    pub consumes: Option<Vec<String>>,
}

#[derive(Debug, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct Parameter {
    pub name: String,
    pub in_: String,
    pub description: Option<String>,
    pub required: bool,
    #[serde(rename = "type")]
    pub param_type: Option<String>,
    pub format: Option<String>,
    pub schema: Option<SchemaRef>,
}

#[derive(Debug, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct Response {
    pub description: String,
    pub schema: Option<SchemaRef>,
}

#[derive(Debug, Serialize, Deserialize)]
#[serde(untagged)]
pub enum SchemaRef {
    Ref {
        #[serde(rename = "$ref")]
        ref_: String,
        original_ref: Option<String>,
    },
    Object(Schema),
}

#[derive(Debug, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct Schema {
    #[serde(rename = "type")]
    pub type_: Option<String>,
    pub properties: Option<HashMap<String, Property>>,
    pub title: Option<String>,
}

#[derive(Debug, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct Property {
    #[serde(rename = "type")]
    pub type_: String,
    pub description: Option<String>,
    pub format: Option<String>,
}

#[derive(Debug, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct SecurityDefinition {
    #[serde(rename = "type")]
    pub type_: String,
    pub name: String,
    pub in_: String,
}

#[derive(Debug, Serialize, Deserialize)]
#[serde(untagged)]
pub enum Definition {
    Object(Schema),
    Other(Value),
}

// 下面是docx的模板
#[derive(Debug, Serialize, Deserialize, Clone)]
pub struct DocxProjectInfo {
    // 项目名称
    name: String,

    // 接口描述
    apis: HashMap<String, Vec<DocxApiInfo>>,
}

#[derive(Debug, Serialize, Deserialize, Clone)]
pub struct DocxApiInfo {
    // 接口名称
    name: String,

    // 接口描述
    desc: String,

    // url
    url: String,

    // 请求方式
    method: String,

    // 请求类型
    api_type: String,

    // 请求类型
    return_type: String,

    // 请求参数列表
    query_params: Vec<DocxParamInfo>,

    // 返回参数列表
    return_params: Vec<DocxReturnParamInfo>,
}

#[derive(Debug, Serialize, Deserialize, Clone)]
pub struct DocxParamInfo {
    // 参数名
    name: String,

    // 数据类型
    data_type: String,

    // 参数类型
    param_type: String,

    // 是否必填
    required: String,

    // 说明
    desc: String,
}

#[derive(Debug, Serialize, Deserialize, Clone)]
pub struct DocxReturnParamInfo {
    // 返回属性名
    name: String,

    // 类型
    data_type: String,

    // 说明
    desc: String,
}
