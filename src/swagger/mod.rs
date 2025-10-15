use docx_handlebars::render_handlebars;
use serde::{Deserialize, Serialize};
use serde_json::{Map, Number, Value};
use std::{
    cell::Ref,
    collections::{HashMap, HashSet},
};

const SWAGGER_DOCX_MODEL: &[u8] = include_bytes!("../../asset/template/swagger-model.docx");

pub fn parse_swagger_and_gen_docx(
    swagger_bytes: &Vec<u8>,
    output_file_name: &String,
) -> Result<(), Box<dyn std::error::Error>> {
    let sw: SwaggerDocument = serde_json::from_slice(&swagger_bytes)?;

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

            // 请求参数
            let mut query_params: Vec<DocxParamInfo> = vec![];
            if let Some(params) = operation.parameters {
                for param in params {
                    if let Some(schema) = param.schema {
                        if let SchemaRef::Ref { ref_, original_ref } = schema {
                            let mut ps = param_by_definitions(
                                &original_ref.unwrap_or("".to_string()),
                                &sw.definitions,
                            );
                            // 在每个参数前面加上"body."
                            ps.iter_mut()
                                .for_each(|item| item.name = format!("body.{}", item.name));
                            query_params.extend(ps);
                        }
                    } else {
                        query_params.push(DocxParamInfo {
                            name: param.name.clone(),
                            data_type: param.param_type.clone().unwrap_or("".to_string()),
                            param_type: param.in_,
                            required: if param.required {
                                "Y".to_string()
                            } else {
                                "N".to_string()
                            },
                            desc: param.description.clone().unwrap_or("".to_string()),
                        });
                    }
                }
            }

            // 状态码
            let mut status_codes: Vec<DocxStatusCode> = vec![];
            for ele in &operation.responses {
                status_codes.push(DocxStatusCode {
                    code: ele.0.clone(),
                    desc: ele.1.description.clone(),
                    explain: "".to_string(),
                });
            }

            // 返回参数
            let mut return_params: Vec<DocxReturnParamInfo> = vec![];
            if let Some(response) = &operation.responses.get("200") {
                let description = response.description.clone();
                if let Some(schema) = &response.schema {
                    if let SchemaRef::Ref { ref_, original_ref } = schema {
                        let mut ps = response_by_definitions(
                            original_ref.as_ref().unwrap_or(&"".to_string()),
                            &sw.definitions,
                            &mut HashSet::<&String>::new(),
                        );
                        // 在每个参数前面加上"body."
                        ps.iter_mut()
                            .for_each(|item| item.name = format!("body.{}", item.name));
                        return_params.extend(ps);
                    }
                }
            }

            // 示例
            let mut example_object = serde_json::Value::Object(Map::new());
            if let Some(response) = &operation.responses.get("200") {
                let description = response.description.clone();
                if let Some(schema) = &response.schema {
                    if let SchemaRef::Ref { ref_, original_ref } = schema {
                        fill_value_by_definitions(
                            original_ref.as_ref().unwrap_or(&"".to_string()),
                            &mut example_object,
                            &sw.definitions,
                            &mut HashSet::<&String>::new(),
                        );
                    }
                }
            }

            let doc_api_info = DocxApiInfo {
                name: operation.summary.clone().unwrap_or("".to_string()),
                desc: operation.summary.clone().unwrap_or("".to_string()),
                url: url.clone(),
                method: method,
                api_type: "".to_string(),
                return_type: "*/*".to_string(),
                query_params: query_params,
                status_codes: status_codes,
                return_params: return_params,
                return_params_example: serde_json::to_string(&example_object)
                    .unwrap_or("".to_string()),
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
    println!("{}", serde_json::to_string_pretty(&docx_project)?);

    // 渲染模板
    let result = render_handlebars(
        SWAGGER_DOCX_MODEL.to_vec(),
        &serde_json::to_value(&docx_project)?,
    )?;

    // 保存
    std::fs::write(output_file_name, result)?;

    return Ok(());
}

// 获得返回属性（嵌套获取）
fn response_by_definitions<'a>(
    original_ref: &'a String,
    definitions: &'a HashMap<String, Definition>,
    used_name: &mut HashSet<&'a String>,
) -> Vec<DocxReturnParamInfo> {
    // 检查是否循环引用
    if used_name.contains(original_ref) {
        return vec![];
    }
    used_name.insert(original_ref);

    let mut ps: Vec<DocxReturnParamInfo> = vec![];
    if let Some(definition) = definitions.get(original_ref) {
        if let Definition::Object(scheme) = definition {
            if let Some(hm) = &scheme.properties {
                for ele in hm {
                    let name = ele.0;
                    let prop = ele.1;
                    let type_ = &prop.type_;
                    if let Some(type_value) = type_ {
                        let data_type = type_value.clone();
                        if "array" == data_type {
                            // 列表
                            if let Some(schema) = &prop.items {
                                if let SchemaRef::Ref { ref_, original_ref } = schema {
                                    if let Some(original_ref_value) = original_ref {
                                        let mut pst = response_by_definitions(
                                            original_ref_value,
                                            &definitions,
                                            &mut used_name.clone(),
                                        );
                                        // 在每个参数前面加上"[]."
                                        pst.iter_mut().for_each(|item| {
                                            item.name = format!("{}.[].{}", name, item.name)
                                        });
                                        ps.extend(pst);
                                    }
                                } else if let SchemaRef::Primitives {
                                    type_,
                                    description,
                                    format,
                                } = schema
                                {
                                    // 属性
                                    let spi = DocxReturnParamInfo {
                                        // todo 优化
                                        name: format!(
                                            "{}.[].{}",
                                            name,
                                            format.clone().unwrap_or("".to_string())
                                        ),
                                        data_type: type_.clone().unwrap_or("".to_string()),
                                        desc: prop.description.clone().unwrap_or("".to_string()),
                                    };
                                    ps.push(spi);
                                }
                            }
                        } else {
                            // 属性
                            let spi = DocxReturnParamInfo {
                                name: name.clone(),
                                data_type: data_type,
                                desc: prop.description.clone().unwrap_or("".to_string()),
                            };
                            ps.push(spi);
                        }
                    } else if let Some(original_ref_value) = &prop.original_ref {
                        // 对象
                        let mut pst = response_by_definitions(
                            original_ref_value,
                            &definitions,
                            &mut used_name.clone(),
                        );
                        // 在每个参数前面加上"."
                        pst.iter_mut()
                            .for_each(|item| item.name = format!("{}.{}", name, item.name));
                        ps.extend(pst);
                    }
                }
            }
        }
    }

    return ps;
}

// 属性填充Value
fn fill_value_by_definitions<'a>(
    original_ref: &'a String,
    value: &mut Value,
    definitions: &'a HashMap<String, Definition>,
    used_name: &mut HashSet<&'a String>,
) {
    // 检查是否循环引用
    if used_name.contains(original_ref) {
        return;
    }
    used_name.insert(original_ref);

    if let Some(definition) = definitions.get(original_ref) {
        if let Definition::Object(scheme) = definition {
            if let Some(hm) = &scheme.properties {
                for ele in hm {
                    let name = ele.0;
                    let prop = ele.1;
                    let type_ = &prop.type_;
                    if let Some(type_value) = type_ {
                        let data_type = type_value.clone();
                        if "array" == data_type {
                            // 列表
                            if let Some(schema) = &prop.items {
                                if let SchemaRef::Ref { ref_, original_ref } = schema {
                                    if let Some(original_ref_value) = original_ref {
                                        let mut value_item = Value::Object(Map::new());
                                        fill_value_by_definitions(
                                            original_ref_value,
                                            &mut value_item,
                                            &definitions,
                                            &mut used_name.clone(),
                                        );
                                        value.as_object_mut().unwrap().insert(
                                            name.to_string(),
                                            Value::Array(vec![value_item]),
                                        );
                                    }
                                } else if let SchemaRef::Primitives {
                                    type_,
                                    description,
                                    format,
                                } = schema
                                {
                                    // 属性
                                    // todo 空数组
                                    value
                                        .as_object_mut()
                                        .unwrap()
                                        .insert(name.to_string(), Value::Array(vec![]));
                                }
                            }
                        } else {
                            // 属性
                            value.as_object_mut().unwrap().insert(
                                name.to_string(),
                                gen_example_value(&name, &data_type, &prop.example),
                            );
                        }
                    } else if let Some(original_ref_value) = &prop.original_ref {
                        // 对象
                        let mut value_item = Value::Object(Map::new());
                        fill_value_by_definitions(
                            original_ref_value,
                            &mut value_item,
                            &definitions,
                            &mut used_name.clone(),
                        );
                        value
                            .as_object_mut()
                            .unwrap()
                            .insert(name.to_string(), value_item);
                    }
                }
            }
        }
    }
}

fn param_by_definitions(
    original_ref: &String,
    definitions: &HashMap<String, Definition>,
) -> Vec<DocxParamInfo> {
    let mut ps: Vec<DocxParamInfo> = vec![];
    if let Some(definition) = definitions.get(original_ref) {
        if let Definition::Object(scheme) = definition {
            let reqwest_long = vec![];
            let require = scheme.required.as_ref().unwrap_or(&reqwest_long);
            if let Some(hm) = &scheme.properties {
                for ele in hm {
                    let name = ele.0;
                    let prop = ele.1;
                    let type_ = &prop.type_;
                    if let Some(type_value) = type_ {
                        let spi = DocxParamInfo {
                            name: name.clone(),
                            data_type: type_value.clone(),
                            param_type: "".to_string(),
                            required: if require.contains(name) {
                                "Y".to_string()
                            } else {
                                "N".to_string()
                            },
                            desc: prop.description.clone().unwrap_or("".to_string()),
                        };
                        ps.push(spi);
                    }
                }
            }
        }
    }

    return ps;
}

// 生成测试数据
fn gen_example_value(name: &String, value_type: &String, example: &Option<Value>) -> Value {
    // 有示例数据则使用示例数据
    if let Some(example_value) = example {
        // return Value::String(example_value.clone());
        return example_value.clone();
    }

    if "integer" == value_type {
        return Value::Number(Number::from(1u32));
    }
    if "boolean" == value_type {
        return Value::Bool(false);
    }
    if "string" == value_type {
        // 日期
        if name.to_lowercase().contains("time") || name.to_lowercase().contains("date") {
            return Value::String("2025-10-13 20:26:09".to_string());
        }
        return Value::String("string".to_string());
    }

    return Value::Null;
}

#[derive(Debug, Serialize, Deserialize)]
pub struct SwaggerDocument {
    pub swagger: String,
    pub info: Info,
    pub host: String,
    pub basePath: Option<String>,
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
    // 引用对象
    Ref {
        #[serde(rename = "$ref")]
        ref_: String,
        #[serde(rename = "originalRef")]
        original_ref: Option<String>,
    },
    // 原始类型 int long
    Primitives {
        #[serde(rename = "type")]
        type_: Option<String>,
        description: Option<String>,
        format: Option<String>,
    },
    Object(Schema),
}

#[derive(Debug, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct Schema {
    #[serde(rename = "type")]
    pub type_: Option<String>,
    pub required: Option<Vec<String>>,
    pub properties: Option<HashMap<String, Property>>,
    pub title: Option<String>,
}

#[derive(Debug, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct Property {
    #[serde(rename = "type")]
    pub type_: Option<String>,
    pub description: Option<String>,
    pub format: Option<String>,
    pub example: Option<Value>,
    pub items: Option<SchemaRef>,
    #[serde(rename = "$ref")]
    pub ref_: Option<String>,
    #[serde(rename = "originalRef")]
    pub original_ref: Option<String>,
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
    pub name: String,

    // 接口描述
    pub apis: HashMap<String, Vec<DocxApiInfo>>,
}

#[derive(Debug, Serialize, Deserialize, Clone)]
pub struct DocxApiInfo {
    // 接口名称
    pub name: String,

    // 接口描述
    pub desc: String,

    // url
    pub url: String,

    // 请求方式
    pub method: String,

    // 请求类型
    pub api_type: String,

    // 请求类型
    pub return_type: String,

    // 请求参数列表
    pub query_params: Vec<DocxParamInfo>,

    // 状态码
    pub status_codes: Vec<DocxStatusCode>,

    // 返回参数列表
    pub return_params: Vec<DocxReturnParamInfo>,

    // 返回参数示例
    pub return_params_example: String,
}

#[derive(Debug, Serialize, Deserialize, Clone)]
pub struct DocxParamInfo {
    // 参数名
    pub name: String,

    // 数据类型
    pub data_type: String,

    // 参数类型
    pub param_type: String,

    // 是否必填
    pub required: String,

    // 说明
    pub desc: String,
}

#[derive(Debug, Serialize, Deserialize, Clone)]
pub struct DocxStatusCode {
    // 状态码
    pub code: String,

    // 描述
    pub desc: String,

    // 说明
    pub explain: String,
}

#[derive(Debug, Serialize, Deserialize, Clone)]
pub struct DocxReturnParamInfo {
    // 返回属性名
    pub name: String,

    // 类型
    pub data_type: String,

    // 说明
    pub desc: String,
}
