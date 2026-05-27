// Actix-приложение для работы с docx-шаблонами документов.

/* Функциональность приложения включает конструктор и заполнитель шаблонов документов 
в формате .docx, возможность кастомизации типовых вносимых форм данных */

// Основной поток данных:
// 1. пользователь загружает документ с расширением .docx в конструктор;
// 2. word/document.xml извлекается из docx-архива как zip;
// 3. xml конвертируется в HTML-предпросмотр;
// 4. выбранные фрагменты заменяются плейсхолдерами вида {{name:type:format}};
// 5. при заполнении шаблона значения форматируются и записываются обратно в docx-файл.

/* Является портируемым прототипом, поэтому использует простую работу с файловой системой 
для хранения файлов без использования специализированных баз данных*/


use actix_files::{Files, NamedFile};
use actix_multipart::Multipart;
use actix_web::error::{ErrorBadRequest, ErrorInternalServerError, ErrorNotFound, ErrorPayloadTooLarge};
use actix_web::http::header::LOCATION;
use actix_web::{web, App, Error, HttpResponse, HttpServer};
use chrono::{Datelike, Local, NaiveDate, NaiveTime, Timelike};
use futures::{StreamExt, TryStreamExt};
use quick_xml::escape::escape;
use quick_xml::events::{BytesEnd, BytesStart, BytesText, Event};
use quick_xml::reader::Reader;
use quick_xml::writer::Writer;
use regex::Regex;
use roxmltree::{Document, Node, ParsingOptions};
use serde::{Deserialize, Serialize};
use serde_json::{from_str, json, to_string_pretty};
use std::collections::{HashMap, HashSet};
use std::fs::{self, File, metadata};
use std::io::{BufWriter, Cursor, Read, Write};
use std::path::{Component, Path, PathBuf};
use std::sync::{Arc, RwLock};
use tera::{Context, Tera};
use webbrowser;
use zip::write::FileOptions;
use zip::{CompressionMethod, ZipArchive, ZipWriter};

const TEMP_CONSTRUCTOR_FILE: &str = "temp_constructor.docx";
const DOC_TEMPLATES_DIR: &str = "doc_templates";
const SAVED_DIR: &str = "saved";
const TEMPLATES_DIR: &str = "templates";
const CUSTOM_TYPES_FILE: &str = "custom_types.json";
const MAX_FILE_SIZE: usize = 20 * 1024 * 1024;
const DOC_XML_PATH: &str = "word/document.xml";
// Константы формулы аппаратной погрешности при измерении уровня освещенности
const HARDWARE_ERROR_FACTOR: f64 = 4.0;
const HARDWARE_ERROR_DIVISOR: f64 = 3.4641;
const DEFAULT_MAIN_RELATIVE_ERROR_PERCENT: &str = "8";
const DEFAULT_ADDITIONAL_RELATIVE_ERROR_PERCENT: &str = "0";
// Служебные пространства имен xml
const WORD_NS: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const XML_NS: &str = "http://www.w3.org/XML/1998/namespace";
const XMLNS_NS: &str = "http://www.w3.org/2000/xmlns/";


#[derive(Serialize, Deserialize, Clone, Debug)]
struct Settings {
    templates_dir: PathBuf,
    saved_dir: PathBuf,
}
// Пока хардкод путей к папкам шаблонов и сохраненных документов
impl Default for Settings {
    fn default() -> Self {
        Settings {
            templates_dir: PathBuf::from(DOC_TEMPLATES_DIR),
            saved_dir: PathBuf::from(SAVED_DIR),
        }
    }
}

type AppState = Arc<RwLock<Settings>>;
fn load_settings() -> Settings {
    Settings::default()
}

// Единица замены, созданной в конструкторе
// Позиция считается по unicode
#[derive(Deserialize, Clone, Debug)]
pub struct Replacement {
    pub old: String, // исходный текст в docx-файле 
    pub insert: String, // плейсхолдер для замены
    pub paragraph_index: usize, // индекс параграфа
    pub offset: usize, // позиция внутри параграфа
}

#[derive(Debug, Clone, Serialize)]
struct Placeholder {
    ph: String,
    name: String,
    field_type: String,
    format: Option<String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
struct CustomType {
    key: String,
    name: String,
    options: Vec<String>,
}

// Один текстовый узел w:t внутри w:r
#[derive(Clone, Debug)]
struct TextNode {
    run_idx: usize,
    text: String,
}

// План низкоуровневой замены внутри xml параграфа с учетом разрывов текста по w:r
#[derive(Clone, Debug)]
struct ReplacePlan {
    start_run_idx: usize,
    start_char_in_run: usize,
    end_run_idx: usize,
    end_char_in_run: usize,
    insert: String,
}

// Диапазон найденного плейсхолдера для заполнения
#[derive(Debug, Clone)]
struct PlaceholderMatch {
    start_char: usize,
    end_char: usize,
    body: String,
}

// Пространства имен xml, обнаруженные в исходном docx-файле
// для предотвращения потери совместимости между разными версиями Word/LibreOffice
#[derive(Clone, Debug, Default)]
struct NamespaceContext {
    declarations: Vec<(String, String)>,
    uri_to_prefix: HashMap<String, String>,
}

#[derive(Default)]
struct ConstructorRequest {
    action: Option<String>,
    replacements: Vec<Replacement>,
    template_name: String,
    force: bool,
    target_folder: String,
}

fn safe_file_name(raw: &str) -> Result<String, Error> {
    let name = Path::new(raw)
        .file_name()
        .and_then(|s| s.to_str())
        .ok_or_else(|| ErrorBadRequest("Invalid file name"))?;
    if name.is_empty() || name.contains('\0') {
        return Err(ErrorBadRequest("Invalid file name"));
    }
    Ok(name.to_string())
} 

fn ensure_docx_name(raw: &str) -> Result<String, Error> {
    let file_name = safe_file_name(raw)?;
    if file_name.ends_with(".docx") {
        Ok(file_name)
    } else {
        Ok(format!("{file_name}.docx"))
    }
}

// Разрешены только относительные пути .docx внутри настроенной директории шаблонов
fn validate_relative_docx_path(raw: &str) -> Result<PathBuf, Error> {
    let path = Path::new(raw);
    if raw.trim().is_empty() || path.is_absolute() {
        return Err(ErrorNotFound("Шаблон не найден"));
    }
    for component in path.components() {
        match component {
            Component::Normal(part) => {
                let part = part.to_string_lossy();
                if part.contains('\\') || part.contains('\0') {
                    return Err(ErrorNotFound("Шаблон не найден"));
                }
            }
            _ => return Err(ErrorNotFound("Шаблон не найден")),
        }
    }
    if path.extension().and_then(|s| s.to_str()) != Some("docx") {
        return Err(ErrorNotFound("Шаблон не найден"));
    }
    Ok(path.to_path_buf())
}

fn resolve_template_file(templates_dir: &Path, raw: &str) -> Result<PathBuf, Error> {
    Ok(templates_dir.join(validate_relative_docx_path(raw)?))
}

fn validate_folder_name(raw: &str) -> Result<Option<String>, Error> {
    let folder = raw.trim();
    if folder.is_empty() || folder == "." {
        return Ok(None);
    }
    if folder.contains('/') || folder.contains('\\') || folder == ".." || folder.starts_with('.') || folder.contains('\0') {
        return Err(ErrorBadRequest("Некорректное имя папки"));
    }
    Ok(Some(folder.to_string()))
}

fn list_docx_files(dir: &Path) -> Vec<String> {
    let mut files = Vec::new();
    if let Ok(entries) = fs::read_dir(dir) {
        for entry in entries.flatten() {
            let path = entry.path();
            let Some(file_name) = entry.file_name().to_str().map(str::to_string) else {
                continue;
            };
            if path.is_file() && file_name.ends_with(".docx") {
                files.push(file_name);
            }
        }
    }
    files.sort();
    files
}

fn list_files_sorted_by_mtime(dir: &Path) -> Result<Vec<String>, Error> {
    let mut entries: Vec<_> = fs::read_dir(dir)
        .map_err(ErrorInternalServerError)?
        .filter_map(|entry| entry.ok())
        .filter_map(|entry| {
            let file_name = entry.file_name().into_string().ok()?;
            let modified = entry.metadata().ok()?.modified().ok()?;
            Some((file_name, modified))
        })
        .collect();
    entries.sort_by(|a, b| b.1.cmp(&a.1));
    Ok(entries.into_iter().map(|(name, _)| name).collect())
}

/*===================== XML/DOCX =====================*/

fn normalize_custom_type_key(raw: &str) -> String {
    let mut key = String::with_capacity(raw.len());
    let mut prev_sep = false;
    for ch in raw.trim().chars() {
        let mapped = if ch.is_whitespace() || ":{},;\"'".contains(ch) {
            '_'
        } else {
            ch
        };
        if mapped == '_' {
            if !prev_sep {
                key.push('_');
                prev_sep = true;
            }
        } else {
            key.push(mapped);
            prev_sep = false;
        }
    }
    key.trim_matches('_').to_string()
}

fn normalize_custom_type(mut item: CustomType) -> CustomType {
    item.name = item.name.trim().to_string();
    item.key = normalize_custom_type_key(if item.key.trim().is_empty() {
        &item.name
    } else {
        &item.key
    });
    let mut seen = HashSet::new();
    item.options = item
        .options
        .into_iter()
        .map(|value| value.trim().to_string())
        .filter(|value| !value.is_empty() && seen.insert(value.clone()))
        .collect();
    item
}

fn load_custom_types() -> Result<Vec<CustomType>, Box<dyn std::error::Error>> {
    let path = Path::new(CUSTOM_TYPES_FILE);
    if !path.exists() {
        return Ok(Vec::new());
    }
    let content = fs::read_to_string(path)?;
    if content.trim().is_empty() {
        return Ok(Vec::new());
    }
    let value: serde_json::Value = serde_json::from_str(&content)?;
    let mut items = match value {
        serde_json::Value::Array(_) => serde_json::from_str::<Vec<CustomType>>(&content)?,
        serde_json::Value::Object(map) => {
            if map.contains_key("key") && map.contains_key("name") && map.contains_key("options") {
                vec![serde_json::from_value::<CustomType>(serde_json::Value::Object(
                    map,
                ))?]
            } else {
                let mut parsed = Vec::new();
                for (raw_key, raw_value) in map {
                    match raw_value {
                        serde_json::Value::Object(mut item) => {
                            let key = item
                                .remove("key")
                                .and_then(|value| value.as_str().map(str::to_string))
                                .unwrap_or_else(|| raw_key.clone());
                            let name = item
                                .remove("name")
                                .and_then(|value| value.as_str().map(str::to_string))
                                .unwrap_or_else(|| raw_key.clone());
                            let options = item
                                .remove("options")
                                .and_then(|value| {
                                    serde_json::from_value::<Vec<String>>(value).ok()
                                })
                                .unwrap_or_default();
                            parsed.push(CustomType { key, name, options });
                        }
                        serde_json::Value::Array(options) => {
                            parsed.push(CustomType {
                                key: raw_key.clone(),
                                name: raw_key,
                                options: serde_json::from_value::<Vec<String>>(
                                    serde_json::Value::Array(options),
                                )?,
                            });
                        }
                        _ => {}
                    }
                }
                parsed
            }
        }
        _ => Vec::new(),
    };
    items = items
        .into_iter()
        .map(normalize_custom_type)
        .filter(|item| !item.key.is_empty() && !item.name.is_empty())
        .collect();
    items.sort_by_cached_key(|item| item.name.to_lowercase());
    Ok(items)
}

fn save_custom_types(items: &[CustomType]) -> Result<(), Box<dyn std::error::Error>> {
    fs::write(CUSTOM_TYPES_FILE, to_string_pretty(items)?)?;
    Ok(())
}

fn render_html(tera: &Tera, template: &str, context: &Context) -> Result<HttpResponse, Error> {
    let rendered = tera
        .render(template, context)
        .map_err(ErrorInternalServerError)?;
    Ok(HttpResponse::Ok()
        .content_type("text/html; charset=utf-8")
        .body(rendered))
}

fn found(location: impl Into<String>) -> HttpResponse {
    HttpResponse::Found()
        .append_header((LOCATION, location.into()))
        .finish()
}

fn escape_html(s: &str) -> String {
    s.replace('&', "&amp;")
        .replace('<', "&lt;")
        .replace('>', "&gt;")
        .replace('\"', "&quot;")
        .replace('\'', "&apos;")
        .replace(
            "javascript:",
            "&#106;&#97;&#118;&#97;&#115;&#99;&#114;&#105;&#112;&#116;:",
        )
}

fn escape_attr(s: &str) -> String {
    escape_html(s).replace('"', "&quot;")
}

/* Word/LibreOffice часто разбивают один текстовый фрагмент в docx-документе 
на несколько run'ов (w:r), из-за чего замена текста не может быть произведена гарантированно 
корректно с помощью простых строковых замен w:t. Для решения данной проблемы используется
обработка xml-узлов с нормализацией и поиском с учетом вариаций w:r.

Чтобы гарантировать корректность применяемых замен, места для них рассчитываются 
по позициям старого текста, включая индекс параграфа, первый символ текста для замены 
внутри параграфа (offset) и сам заменяемый текст (old). Если при расчете позиции и содержимое 
заменяемого текста несовпадают, замена не применяется.

Строится карта символов по run'ам и затем переписываются только затронутые w:t, 
сохраняя форматирование первого run'а, что позволяет сохранять значительную часть форматирования 
на документах со сложными стилями */


// Вычисляет смещение текущего текстового узла внутри run'а
fn get_char_offset_in_run(text_nodes: &[TextNode], node_idx: usize) -> usize {
    let target_run_idx = text_nodes[node_idx].run_idx;
    let mut offset = 0usize;
    for node in text_nodes.iter().take(node_idx) {
        if node.run_idx == target_run_idx {
            offset += node.text.chars().count();
        }
    }
    offset
}

fn collect_all_paragraphs<'a>(root: Node<'a, 'a>) -> Vec<Node<'a, 'a>> {
    root.descendants().filter(|n| n.has_tag_name("p")).collect()
}

fn collect_text_nodes<'a>(p_node: Node<'a, 'a>) -> Vec<TextNode> {
    let mut nodes = Vec::new();
    let mut run_idx = 0usize;
    for run in p_node
        .children()
        .filter(|n| n.is_element() && n.tag_name().name() == "r")
    {
        for t in run
            .children()
            .filter(|n| n.is_element() && n.tag_name().name() == "t")
        {
            if let Some(text) = t.text() {
                nodes.push(TextNode {
                    run_idx,
                    text: text.to_string(),
                });
            }
        }
        run_idx += 1;
    }
    nodes
}

fn build_char_map(text_nodes: &[TextNode]) -> (String, Vec<(usize, usize)>) {
    let mut full_text = String::new();
    let mut char_map = Vec::new();
    for (node_idx, node) in text_nodes.iter().enumerate() {
        for (local_idx, ch) in node.text.chars().enumerate() {
            full_text.push(ch);
            char_map.push((node_idx, local_idx));
        }
    }
    (full_text, char_map)
}

fn find_placeholder_matches(text: &str) -> Vec<PlaceholderMatch> {
    let re = Regex::new(r"\{\{?\s*([^{}]+?)\s*\}\}?").unwrap();
    re.captures_iter(text)
        .filter_map(|caps| {
            let whole = caps.get(0)?;
            let body = caps.get(1)?.as_str().trim().to_string();
            if body.is_empty() {
                return None;
            }
            let start_char = text[..whole.start()].chars().count();
            let end_char = start_char + whole.as_str().chars().count();
            Some(PlaceholderMatch {
                start_char,
                end_char,
                body,
            })
        })
        .collect()
}

// Преобразует диапазон в полном тексте параграфа в диапазон run'ов
fn build_plan_from_offsets(
    text_nodes: &[TextNode],
    char_map: &[(usize, usize)],
    start_offset: usize,
    end_offset: usize,
    insert: String,
) -> Option<ReplacePlan> {
    if end_offset == 0 || end_offset > char_map.len() {
        return None;
    }
    let start = char_map[start_offset];
    let end = char_map[end_offset - 1];
    let start_char_in_run = get_char_offset_in_run(text_nodes, start.0) + start.1;
    let end_char_in_run = get_char_offset_in_run(text_nodes, end.0) + end.1 + 1;
    Some(ReplacePlan {
        start_run_idx: text_nodes[start.0].run_idx,
        start_char_in_run,
        end_run_idx: text_nodes[end.0].run_idx,
        end_char_in_run,
        insert,
    })
}

fn sort_replace_plans(plans: &mut [ReplacePlan]) {
    plans.sort_by(|a, b| {
        a.start_run_idx
            .cmp(&b.start_run_idx)
            .then_with(|| a.start_char_in_run.cmp(&b.start_char_in_run))
    });
}

fn read_docx_document_xml(path: &Path) -> Result<String, Box<dyn std::error::Error>> {
    let mut archive = ZipArchive::new(File::open(path)?)?;
    let mut xml = String::new();
    archive.by_name(DOC_XML_PATH)?.read_to_string(&mut xml)?;
    Ok(xml)
}

fn read_uploaded_docx_xml(path: &Path) -> Result<String, Error> {
    let file = File::open(path).map_err(ErrorInternalServerError)?;
    let mut archive = ZipArchive::new(file).map_err(|_| ErrorBadRequest("Некорректный файл"))?;
    let mut xml = String::new();
    archive
        .by_name(DOC_XML_PATH)
        .map_err(|_| ErrorBadRequest("В docx отсутствует word/document.xml"))?
        .read_to_string(&mut xml)
        .map_err(ErrorInternalServerError)?;
    Ok(xml)
}

fn rewrite_docx_document_xml<F>(
    input_path: &Path,
    output_path: &Path,
    mut rewrite: F,
) -> Result<(), Box<dyn std::error::Error>>
where
    F: FnMut(&str) -> Result<String, Box<dyn std::error::Error>>,
{
    let mut archive = ZipArchive::new(File::open(input_path)?)?;
    let tmp_path = output_path.with_extension("tmp");
    {
        let out_file = File::create(&tmp_path)?;
        let mut buffered_out = BufWriter::with_capacity(8 * 1024 * 1024, out_file);
        let mut new_archive = ZipWriter::new(&mut buffered_out);
        let options =
            FileOptions::<'_, ()>::default().compression_method(CompressionMethod::Deflated);
        for i in 0..archive.len() {
            let mut part = archive.by_index(i)?;
            let name = part.name().to_string();
            new_archive.start_file(name.clone(), options)?;
            if name == DOC_XML_PATH {
                let mut xml = String::new();
                part.read_to_string(&mut xml)?;
                new_archive.write_all(rewrite(&xml)?.as_bytes())?;
            } else {
                std::io::copy(&mut part, &mut new_archive)?;
            }
        }
        new_archive.finish()?;
        buffered_out.flush()?;
    }
    if output_path.exists() {
        fs::remove_file(output_path)?;
    }
    fs::rename(&tmp_path, output_path)?;
    Ok(())
}

fn extract_namespace_context(xml: &str) -> NamespaceContext {
    let start = xml.find("<w:document").unwrap_or(0);
    let rest = &xml[start..];
    let end_rel = rest.find('>').unwrap_or(rest.len());
    let head = &rest[..end_rel];
    let re = Regex::new(
        r#"xmlns(?::([A-Za-z_][A-Za-z0-9_.-]*))?\s*=\s*(?:"([^"]*)"|'([^']*)')"#,
    )
    .unwrap();
    let mut ctx = NamespaceContext::default();
    let mut seen = HashSet::new();
    for caps in re.captures_iter(head) {
        let prefix = caps.get(1).map(|m| m.as_str()).unwrap_or("").to_string();
        let uri = caps
            .get(2)
            .or_else(|| caps.get(3))
            .map(|m| m.as_str())
            .unwrap_or("")
            .to_string();
        if seen.insert((prefix.clone(), uri.clone())) {
            ctx.uri_to_prefix
                .entry(uri.clone())
                .or_insert(prefix.clone());
            ctx.declarations.push((prefix, uri));
        }
    }
    if !ctx.declarations.iter().any(|(prefix, _)| prefix == "w") {
        ctx.declarations
            .push(("w".to_string(), WORD_NS.to_string()));
        ctx.uri_to_prefix
            .entry(WORD_NS.to_string())
            .or_insert_with(|| "w".to_string());
    }
    if !ctx.declarations.iter().any(|(prefix, _)| prefix == "xml") {
        ctx.declarations
            .push(("xml".to_string(), XML_NS.to_string()));
        ctx.uri_to_prefix
            .entry(XML_NS.to_string())
            .or_insert_with(|| "xml".to_string());
    }
    ctx
}

fn qname_from_ns(local: &str, ns: Option<&str>, ns_ctx: &NamespaceContext) -> String {
    match ns {
        Some(XML_NS) => format!("xml:{local}"),
        Some(uri) => {
            if let Some(prefix) = ns_ctx.uri_to_prefix.get(uri) {
                if prefix.is_empty() {
                    local.to_string()
                } else {
                    format!("{prefix}:{local}")
                }
            } else if uri == WORD_NS {
                format!("w:{local}")
            } else {
                local.to_string()
            }
        }
        None => local.to_string(),
    }
}

// Собирает набор quick-xml событий обратно в строковый xml-фрагмент параграфа
fn paragraph_events_to_xml(
    events: &[Event<'static>],
) -> Result<String, Box<dyn std::error::Error>> {
    let mut out = Vec::new();
    let mut writer = Writer::new(Cursor::new(&mut out));
    for event in events {
        writer.write_event(event.clone())?;
    }
    Ok(String::from_utf8(out)?)
}

// Оборачивает xml-фрагмент параграфа в временный <root> с неймспейсом для корректной обработки префиксов
fn wrap_fragment_in_root(fragment: &str, ns_ctx: &NamespaceContext) -> String {
    let mut wrapped = String::from("<root");
    for (prefix, uri) in &ns_ctx.declarations {
        if prefix.is_empty() {
            wrapped.push_str(&format!(r#" xmlns="{}""#, escape_attr(uri)));
        } else {
            wrapped.push_str(&format!(r#" xmlns:{}="{}""#, prefix, escape_attr(uri)));
        }
    }
    wrapped.push('>');
    wrapped.push_str(fragment);
    wrapped.push_str("</root>");
    wrapped
}

fn collect_paragraph_events<R: std::io::BufRead>(
    reader: &mut Reader<R>,
    start_event: Event<'static>,
) -> Result<Vec<Event<'static>>, Box<dyn std::error::Error>> {
    let mut events = vec![start_event];
    let mut depth = 1usize;
    let mut buf = Vec::new();
    while depth > 0 {
        buf.clear();
        let event = reader.read_event_into(&mut buf)?.into_owned();
        match &event {
            Event::Start(e) if e.name().as_ref() == b"w:p" => depth += 1,
            Event::End(e) if e.name().as_ref() == b"w:p" => depth -= 1,
            Event::Eof => return Err("Неожиданный конец XML внутри w:p".into()),
            _ => {}
        }
        events.push(event);
    }
    Ok(events)
}

// Главный xml-стриминг
fn rewrite_document_xml_by_paragraph_with_index<F>(
    xml: &str,
    mut planner: F,
) -> Result<String, Box<dyn std::error::Error>>
where
    F: for<'a> FnMut(
        usize,
        Node<'a, 'a>,
    ) -> Result<Option<Vec<ReplacePlan>>, Box<dyn std::error::Error>>,
{
    let ns_ctx = extract_namespace_context(xml);
    let mut reader = Reader::from_str(xml);
    reader.trim_text(false);
    let mut writer_output = Vec::new();
    let mut writer = Writer::new(Cursor::new(&mut writer_output));
    let mut buf = Vec::new();
    let mut paragraph_index = 0usize;
    loop {
        buf.clear();
        let event = reader.read_event_into(&mut buf)?.into_owned();
        if matches!(&event, Event::Start(e) if e.name().as_ref() == b"w:p") {
            // предполагается отсутствие необходимости в плейсхолдерах, пересекающих разные w:p
            let paragraph_events = collect_paragraph_events(&mut reader, event)?;
            let fragment = paragraph_events_to_xml(&paragraph_events)?;
            let wrapped = wrap_fragment_in_root(&fragment, &ns_ctx);
            let paragraph_doc = Document::parse_with_options(
                &wrapped,
                ParsingOptions {
                    allow_dtd: false,
                    ..Default::default()
                },
            )?;
            let p_node = paragraph_doc
                .root_element()
                .children()
                .find(|n| n.is_element() && n.tag_name().name() == "p")
                .ok_or("Не удалось извлечь параграф из XML-фрагмента")?;
            match planner(paragraph_index, p_node)? {
                Some(plans) if !plans.is_empty() => {
                    generate_paragraph(&mut writer, p_node, &plans, &ns_ctx)?;
                }
                _ => {
                    for ev in paragraph_events {
                        writer.write_event(ev)?;
                    }
                }
            }
            paragraph_index += 1;
            continue;
        }
        match event {
            Event::Eof => break,
            other => writer.write_event(other)?,
        }
    }
    Ok(String::from_utf8(writer_output)?)
}

fn rewrite_document_xml_by_paragraph<F>(
    xml: &str,
    mut planner: F,
) -> Result<String, Box<dyn std::error::Error>>
where
    F: for<'a> FnMut(
        Node<'a, 'a>,
    ) -> Result<Option<Vec<ReplacePlan>>, Box<dyn std::error::Error>>,
{
    rewrite_document_xml_by_paragraph_with_index(xml, |_, p_node| planner(p_node))
}

// В случае несовпадения вывод ошибки в браузере
fn analyze_paragraph(
    p_node: Node<'_, '_>,
    reps: &[Replacement],
) -> Result<(Vec<ReplacePlan>, Vec<String>), Box<dyn std::error::Error>> {
    let text_nodes = collect_text_nodes(p_node);
    let (full_text, char_map) = build_char_map(&text_nodes);
    let full_len = full_text.chars().count();
    let mut plans = Vec::with_capacity(reps.len());
    let mut errors = Vec::new();
    for rep in reps {
        let old_len = rep.old.chars().count();
        if old_len == 0 {
            errors.push(format!(
                "Параграф {}: пустой фрагмент для замены",
                rep.paragraph_index
            ));
            continue;
        }
        if rep.offset + old_len > full_len {
            errors.push(format!(
                "В параграфе {}: позиция {} + длина {} превышает длину текста {}. Замена '{}' на '{}'",
                rep.paragraph_index, rep.offset, old_len, full_len, rep.old, rep.insert
            ));
            continue;
        }
        let extracted: String = full_text.chars().skip(rep.offset).take(old_len).collect();
        if extracted != rep.old {
            errors.push(format!(
                "Параграф {}, позиция {}: ожидалось '{}', найдено '{}'. Замена на '{}'",
                rep.paragraph_index, rep.offset, rep.old, extracted, rep.insert
            ));
            continue;
        }
        let start = char_map[rep.offset];
        let end = char_map[rep.offset + old_len - 1];
        let start_char_in_run = get_char_offset_in_run(&text_nodes, start.0) + start.1;
        let end_char_in_run = get_char_offset_in_run(&text_nodes, end.0) + end.1 + 1;
        plans.push(ReplacePlan {
            start_run_idx: text_nodes[start.0].run_idx,
            start_char_in_run,
            end_run_idx: text_nodes[end.0].run_idx,
            end_char_in_run,
            insert: rep.insert.clone(),
        });
    }
    sort_replace_plans(&mut plans);
    Ok((plans, errors))
}

fn copy_attributes_to_start(
    attrs: roxmltree::Attributes<'_, '_>,
    start: &mut BytesStart<'_>,
    ns_ctx: &NamespaceContext,
) {
    for attr in attrs {
        if attr.namespace() == Some(XMLNS_NS) {
            continue;
        }
        let qname = qname_from_ns(attr.name(), attr.namespace(), ns_ctx);
        start.push_attribute((qname.as_bytes(), attr.value().as_bytes()));
    }
}

fn write_node<W: Write>(
    writer: &mut Writer<W>,
    node: Node,
    ns_ctx: &NamespaceContext,
) -> Result<(), Box<dyn std::error::Error>> {
    if node.is_text() {
        if let Some(text) = node.text() {
            if !text.is_empty() {
                writer.write_event(Event::Text(BytesText::from_escaped(escape(text))))?;
            }
        }
        return Ok(());
    }
    if !node.is_element() {
        return Ok(());
    }
    let qname = qname_from_ns(node.tag_name().name(), node.tag_name().namespace(), ns_ctx);
    let mut start = BytesStart::new(qname.as_str());
    copy_attributes_to_start(node.attributes(), &mut start, ns_ctx);
    writer.write_event(Event::Start(start))?;
    for child in node.children() {
        write_node(writer, child, ns_ctx)?;
    }
    writer.write_event(Event::End(BytesEnd::new(qname.as_str())))?;
    Ok(())
}

// Пишет w:r с новым текстом, сохраняя w:rPr
fn write_run_with_text<W: Write>(
    writer: &mut Writer<W>,
    run_node: Node,
    new_text: &str,
    ns_ctx: &NamespaceContext,
) -> Result<(), Box<dyn std::error::Error>> {
    let run_qname = qname_from_ns(
        run_node.tag_name().name(),
        run_node.tag_name().namespace(),
        ns_ctx,
    );
    writer.write_event(Event::Start(BytesStart::new(run_qname.as_str())))?;
    for child in run_node.children() {
        if child.is_element() && child.tag_name().name() == "rPr" {
            write_node(writer, child, ns_ctx)?;
        }
    }
    if !new_text.is_empty() {
        let text_qname = run_node
            .children()
            .find(|n| n.is_element() && n.tag_name().name() == "t")
            .map(|t_node| {
                qname_from_ns(
                    t_node.tag_name().name(),
                    t_node.tag_name().namespace(),
                    ns_ctx,
                )
            })
            .unwrap_or_else(|| "w:t".to_string());
        let mut t_start = BytesStart::new(text_qname.as_str());
        t_start.push_attribute(("xml:space", "preserve"));
        writer.write_event(Event::Start(t_start))?;
        writer.write_event(Event::Text(BytesText::from_escaped(escape(new_text))))?;
        writer.write_event(Event::End(BytesEnd::new(text_qname.as_str())))?;
    }
    writer.write_event(Event::End(BytesEnd::new(run_qname.as_str())))?;
    Ok(())
}


// Сложность связана с тем, что замена может начинаться в одном w:r, заканчиваться в другом и пересекать служебные proofErr
fn generate_paragraph<W: Write>(
    writer: &mut Writer<W>,
    p_node: Node,
    plans: &[ReplacePlan],
    ns_ctx: &NamespaceContext,
) -> Result<(), Box<dyn std::error::Error>> {
    let p_qname =
        qname_from_ns(p_node.tag_name().name(), p_node.tag_name().namespace(), ns_ctx);
    let mut p_start = BytesStart::new(p_qname.as_str());
    copy_attributes_to_start(p_node.attributes(), &mut p_start, ns_ctx);
    writer.write_event(Event::Start(p_start))?;
    let mut run_idx = 0usize;
    for child in p_node.children() {
        if !child.is_element() {
            write_node(writer, child, ns_ctx)?;
            continue;
        }
        if child.tag_name().name() != "r" {
            let is_proof_err_inside_placeholder = child.tag_name().name() == "proofErr"
                && plans
                    .iter()
                    .any(|plan| plan.start_run_idx < run_idx && run_idx <= plan.end_run_idx);
            if !is_proof_err_inside_placeholder {
                write_node(writer, child, ns_ctx)?;
            }
            continue;
        }
        let orig_text: String = child
            .children()
            .filter(|c| c.is_element() && c.tag_name().name() == "t")
            .filter_map(|n| n.text())
            .collect();
        let char_count = orig_text.chars().count();
        if char_count == 0 {
            write_node(writer, child, ns_ctx)?;
            run_idx += 1;
            continue;
        }
        let char_to_byte = |pos: usize| {
            orig_text
                .char_indices()
                .nth(pos)
                .map(|(byte, _)| byte)
                .unwrap_or(orig_text.len())
        };
        let mut segments: Vec<String> = Vec::new();
        let mut cursor = 0usize;
        for plan in plans {
            if plan.end_run_idx < run_idx {
                continue;
            }
            if plan.start_run_idx > run_idx {
                break;
            }
            if plan.start_run_idx < run_idx && plan.end_run_idx > run_idx {
                cursor = char_count;
                break;
            }
            if plan.start_run_idx == run_idx && plan.end_run_idx == run_idx {
                if plan.start_char_in_run > cursor {
                    let start_byte = char_to_byte(cursor);
                    let end_byte = char_to_byte(plan.start_char_in_run);
                    segments.push(orig_text[start_byte..end_byte].to_string());
                }
                segments.push(plan.insert.clone());
                cursor = plan.end_char_in_run.min(char_count);
                continue;
            }
            if plan.start_run_idx == run_idx {
                if plan.start_char_in_run > cursor {
                    let start_byte = char_to_byte(cursor);
                    let end_byte = char_to_byte(plan.start_char_in_run);
                    segments.push(orig_text[start_byte..end_byte].to_string());
                }
                segments.push(plan.insert.clone());
                cursor = char_count;
                continue;
            }
            if plan.end_run_idx == run_idx {
                cursor = plan.end_char_in_run.min(char_count);
            }
        }
        if cursor < char_count {
            let start_byte = char_to_byte(cursor);
            segments.push(orig_text[start_byte..].to_string());
        }
        for segment in segments.into_iter().filter(|s| !s.is_empty()) {
            write_run_with_text(writer, child, &segment, ns_ctx)?;
        }
        run_idx += 1;
    }
    writer.write_event(Event::End(BytesEnd::new(p_qname.as_str())))?;
    Ok(())
}

fn build_fill_plans(
    p_node: Node<'_, '_>,
    replacements: &HashMap<String, String>,
) -> Vec<ReplacePlan> {
    let text_nodes = collect_text_nodes(p_node);
    let (full_text, char_map) = build_char_map(&text_nodes);
    let mut plans = Vec::new();
    for ph in find_placeholder_matches(&full_text) {
        let Some(value) = replacements.get(&ph.body) else {
            continue;
        };
        if let Some(plan) = build_plan_from_offsets(
            &text_nodes,
            &char_map,
            ph.start_char,
            ph.end_char,
            value.clone(),
        ) {
            plans.push(plan);
        }
    }
    sort_replace_plans(&mut plans);
    plans
}

// Применяет замены конструктора к временно загруженному docx, сохраняет новый шаблон
// Диагностика ошибок замен
fn apply_constructor_replacements(
    template_path: &Path,
    output_path: &Path,
    mut replacements: Vec<Replacement>,
    force: bool,
) -> Result<(), Box<dyn std::error::Error>> {
    let file_size = metadata(template_path)?.len() as usize;
    if file_size > MAX_FILE_SIZE {
        return Err("Файл слишком большой".into());
    }
    replacements.sort_by(|a, b| {
        a.paragraph_index
            .cmp(&b.paragraph_index)
            .then_with(|| b.offset.cmp(&a.offset))
    });
    let mut replacements_by_paragraph: HashMap<usize, Vec<Replacement>> = HashMap::new();
    for rep in replacements {
        replacements_by_paragraph
            .entry(rep.paragraph_index)
            .or_default()
            .push(rep);
    }
    let mut all_errors = Vec::new();
    rewrite_docx_document_xml(template_path, output_path, |xml| {
        rewrite_document_xml_by_paragraph_with_index(xml, |paragraph_index, p_node| {
            let reps_for_paragraph = replacements_by_paragraph
                .get(&paragraph_index)
                .cloned()
                .unwrap_or_default();
            if reps_for_paragraph.is_empty() {
                return Ok(None);
            }
            let (plans, mut paragraph_errors) = analyze_paragraph(p_node, &reps_for_paragraph)?;
            all_errors.append(&mut paragraph_errors);
            Ok((!plans.is_empty()).then_some(plans))
        })
    })?;
    if !force && !all_errors.is_empty() {
        let _ = fs::remove_file(output_path);
        return Err(
            format!("ошибки сохранения полей:\n{}", all_errors.join("\n")).into(),
        );
    }
    Ok(())
}

/*===================== XML TO HTML =====================*/

fn build_raw_text(p_node: Node<'_, '_>) -> String {
    collect_text_nodes(p_node)
        .into_iter()
        .map(|node| node.text)
        .collect::<String>()
}
fn paragraph_xml_to_html(p_node: Node<'_, '_>, para_index: &mut usize) -> String {
    let raw_text = build_raw_text(p_node);
    let mut replaced = String::new();
    let mut cursor_char = 0usize;
    for ph in find_placeholder_matches(&raw_text) {
        let start_byte = raw_text
            .char_indices()
            .nth(ph.start_char)
            .map(|(idx, _)| idx)
            .unwrap_or(raw_text.len());
        let cursor_byte = raw_text
            .char_indices()
            .nth(cursor_char)
            .map(|(idx, _)| idx)
            .unwrap_or(raw_text.len());
        replaced.push_str(&escape_html(&raw_text[cursor_byte..start_byte]));
        let name = ph.body.split(':').next().unwrap_or(&ph.body).trim();
        replaced.push_str(&format!(
            r#"<span class="placeholder" data-ph="{}" contenteditable="false">{}</span>"#,
            escape_attr(&ph.body),
            escape_html(name)
        ));
        cursor_char = ph.end_char;
    }
    let cursor_byte = raw_text
        .char_indices()
        .nth(cursor_char)
        .map(|(idx, _)| idx)
        .unwrap_or(raw_text.len());
    replaced.push_str(&escape_html(&raw_text[cursor_byte..]));
    let html = format!(r#"<p data-para-index="{}">{}</p>"#, *para_index, replaced);
    *para_index += 1;
    html
}
fn table_xml_to_html(tbl_node: Node<'_, '_>, para_index: &mut usize) -> String {
    let mut html = String::from(r#"<table class="table table-bordered">"#);
    for tr in tbl_node.descendants().filter(|n| n.has_tag_name("tr")) {
        html.push_str("<tr>");
        for tc in tr.descendants().filter(|n| n.has_tag_name("tc")) {
            html.push_str("<td>");
            for p in tc.descendants().filter(|n| n.has_tag_name("p")) {
                html.push_str(&paragraph_xml_to_html(p, para_index));
            }
            html.push_str("</td>");
        }
        html.push_str("</tr>");
    }
    html.push_str("</table>");
    html
}
fn docx_xml_to_html(xml: &str) -> Result<String, Box<dyn std::error::Error>> {
    let doc = Document::parse_with_options(
        xml,
        ParsingOptions {
            allow_dtd: false,
            ..Default::default()
        },
    )?;
    let root = doc.root();
    let mut html = String::new();
    let mut para_index = 0usize;
    if let Some(body) = root.descendants().find(|n| n.has_tag_name("body")) {
        for child in body.children() {
            match child.tag_name().name() {
                "p" => html.push_str(&paragraph_xml_to_html(child, &mut para_index)),
                "tbl" => html.push_str(&table_xml_to_html(child, &mut para_index)),
                _ => {}
            }
        }
    }
    Ok(html)
}

// Для совместимости отсутствующий тип считается text
fn parse_placeholder(ph: &str) -> (String, String, Option<String>) {
    let mut parts = ph.split(':').map(str::trim);
    let name = parts.next().unwrap_or("").to_string();
    let field_type = parts.next().unwrap_or("text").to_string();
    let format = parts.next().map(str::to_string);
    (name, field_type, format)
}
fn parse_f64_loose(value: &str) -> Option<f64> {
    let normalized = value.trim().replace(',', ".");
    if normalized.is_empty() {
        return None;
    }
    normalized.parse::<f64>().ok()
}

// Рассчет аппаратной погрешности при изменении освещенности по специальной формуле
// Ошибка рассчитывается как корень суммы квадратов основной и дополнительной относительных погрешностей, масштабированный коэффициентом 4
pub fn calculate_error(
    measurement_result: f64,
    main_relative_error_percent: f64,
    additional_relative_error_percent: f64,
) -> f64 {
    HARDWARE_ERROR_FACTOR
        * (((main_relative_error_percent / 100.0 * measurement_result) / HARDWARE_ERROR_DIVISOR).powi(2)
            + ((additional_relative_error_percent / 100.0 * measurement_result) / HARDWARE_ERROR_DIVISOR).powi(2))
            .sqrt()
}

fn format_float_trimmed(value: f64) -> String {
    format!("{value:.10}")
        .trim_end_matches('0')
        .trim_end_matches('.')
        .to_string()
}

fn format_hardware_error_value(
    measurement_result: f64,
    main_relative_error_percent: f64,
    additional_relative_error_percent: f64,
) -> String {
    let error = calculate_error(
        measurement_result,
        main_relative_error_percent,
        additional_relative_error_percent,
    );
    let measurement_text = format_float_trimmed(measurement_result);
    let error_text = format!("{error:.2}");
    format!("{measurement_text}±{error_text}")
}

// Встроенные типы и форматы
fn format_value(type_: &str, format_: Option<&str>, value: &str) -> String {
    match type_ {
        "date" => {
            if let Ok(date) = NaiveDate::parse_from_str(value, "%Y-%m-%d") {
                if let Some("full") = format_ {
                    let months = [
                        "января", "февраля", "марта", "апреля", "мая", "июня", 
                        "июля", "августа", "сентября", "октября", "ноября", "декабря",
                    ];
                    format!(
                        "\"{:02}\" {} {} г.",
                        date.day(),
                        months[date.month() as usize - 1],
                        date.year()
                    )
                } else {
                    date.format("%d.%m.%Y").to_string()
                }
            } else {
                value.to_string()
            }
        }
        "time" => {
            if let Ok(time) = NaiveTime::parse_from_str(value, "%H:%M") {
                if let Some("hour_min") = format_ {
                    format!("{:02} час. {:02} мин", time.hour(), time.minute())
                } else {
                    time.format("%H:%M").to_string()
                }
            } else {
                value.to_string()
            }
        }
        "hardware_error" => value.to_string(),
        _ => value.to_string(),
    }
}

fn get_placeholders_from_template(
    template_path: &Path,
) -> Result<HashSet<String>, Box<dyn std::error::Error>> {
    let xml = read_docx_document_xml(template_path)?;
    let doc = Document::parse_with_options(
        &xml,
        ParsingOptions {
            allow_dtd: false,
            ..Default::default()
        },
    )?;
    let mut set = HashSet::new();
    for p_node in collect_all_paragraphs(doc.root()) {
        let raw_text = build_raw_text(p_node);
        for ph in find_placeholder_matches(&raw_text) {
            set.insert(ph.body);
        }
    }
    Ok(set)
}

async fn parse_fill_form(mut payload: Multipart) -> Result<HashMap<String, String>, Error> {
    let mut form = HashMap::new();
    let mut total_size = 0usize;
    while let Some(mut field) = payload.try_next().await? {
        let name = field.name().to_string();
        if name.is_empty() {
            return Err(ErrorBadRequest("Поле формы без имени"));
        }
        let value = read_field_bytes_limited(
            &mut field,
            &mut total_size,
            "Слишком большой payload формы",
        )
        .await?;
        let value = String::from_utf8(value)
            .map_err(|_| ErrorBadRequest("Форма содержит некорректный UTF-8"))?;
        form.insert(name, value);
    }
    Ok(form)
}

fn fill_docx_template(
    template_path: &Path,
    output_path: &Path,
    replacements: &HashMap<String, String>,
) -> Result<(), Box<dyn std::error::Error>> {
    rewrite_docx_document_xml(template_path, output_path, |xml| {
        rewrite_document_xml_by_paragraph(xml, |p_node| {
            let plans = build_fill_plans(p_node, replacements);
            Ok((!plans.is_empty()).then_some(plans))
        })
    })
}

async fn read_field_bytes_limited(
    field: &mut actix_multipart::Field,
    total_size: &mut usize,
    error_message: &'static str,
) -> Result<Vec<u8>, Error> {
    let mut bytes = Vec::new();
    while let Some(chunk) = field.next().await {
        let chunk = chunk.map_err(ErrorInternalServerError)?;
        *total_size += chunk.len();
        if *total_size > MAX_FILE_SIZE {
            return Err(ErrorPayloadTooLarge(error_message));
        }
        bytes.extend_from_slice(&chunk);
    }
    Ok(bytes)
}

async fn read_text_field(field: &mut actix_multipart::Field) -> Result<String, Error> {
    let mut total_size = 0usize;
    let bytes = read_field_bytes_limited(field, &mut total_size, "Превышен лимит формы").await?;
    Ok(String::from_utf8_lossy(&bytes).into_owned())
}

async fn save_uploaded_file(
    field: &mut actix_multipart::Field,
    path: &Path,
) -> Result<(), Error> {
    let mut file = File::create(path).map_err(ErrorInternalServerError)?;
    let mut total_size = 0usize;
    while let Some(chunk) = field.next().await {
        let chunk = chunk.map_err(ErrorInternalServerError)?;
        total_size += chunk.len();
        if total_size > MAX_FILE_SIZE {
            return Err(ErrorPayloadTooLarge("Файл слишком большой"));
        }
        file.write_all(&chunk).map_err(ErrorInternalServerError)?;
    }
    Ok(())
}

async fn parse_constructor_request(mut payload: Multipart) -> Result<ConstructorRequest, Error> {
    let mut request = ConstructorRequest::default();
    while let Some(mut field) = payload.try_next().await? {
        match field.name() {
            "action" => request.action = Some(read_text_field(&mut field).await?),
            "replacements" => {
                let content = read_text_field(&mut field).await?;
                request.replacements = from_str(&content).unwrap_or_default();
            }
            "template_name" => request.template_name = read_text_field(&mut field).await?,
            "force" => request.force = read_text_field(&mut field).await? == "true",
            "target_folder" => request.target_folder = read_text_field(&mut field).await?,
            "template_file" => {
                save_uploaded_file(&mut field, Path::new(TEMP_CONSTRUCTOR_FILE)).await?;
            }
            _ => while field.next().await.is_some() {},
        }
    }
    Ok(request)
}

async fn custom_types_get() -> Result<HttpResponse, Error> {
    let items = load_custom_types().map_err(ErrorInternalServerError)?;
    Ok(HttpResponse::Ok().json(items))
}

async fn custom_types_upsert(item: web::Json<CustomType>) -> Result<HttpResponse, Error> {
    let item = normalize_custom_type(item.into_inner());
    if item.name.is_empty() {
        return Err(ErrorBadRequest("Имя типа не может быть пустым"));
    }
    if item.key.is_empty() {
        return Err(ErrorBadRequest("Ключ типа не может быть пустым"));
    }
    if item.options.is_empty() {
        return Err(ErrorBadRequest("Добавьте хотя бы один вариант текста"));
    }
    let mut items = load_custom_types().map_err(ErrorInternalServerError)?;
    if let Some(existing) = items.iter_mut().find(|existing| existing.key == item.key) {
        *existing = item.clone();
    } else {
        items.push(item.clone());
    }
    items.sort_by_cached_key(|item| item.name.to_lowercase());
    save_custom_types(&items).map_err(ErrorInternalServerError)?;
    Ok(HttpResponse::Ok().json(item))
}

async fn custom_types_delete(path: web::Path<String>) -> Result<HttpResponse, Error> {
    let key = normalize_custom_type_key(&path.into_inner());
    if key.is_empty() {
        return Err(ErrorBadRequest("Некорректный ключ типа"));
    }
    let mut items = load_custom_types().map_err(ErrorInternalServerError)?;
    let original_len = items.len();
    items.retain(|item| item.key != key);
    if items.len() == original_len {
        return Err(ErrorNotFound("Тип не найден"));
    }
    save_custom_types(&items).map_err(ErrorInternalServerError)?;
    Ok(HttpResponse::Ok().finish())
}

/*===================== HANDLERS =====================*/

async fn constructor_get(tera: web::Data<Tera>) -> Result<HttpResponse, Error> {
    render_html(&tera, "constructor.html", &Context::new())
}

async fn constructor_post(
    payload: Multipart,
    tera: web::Data<Tera>,
    settings: web::Data<AppState>,
) -> Result<HttpResponse, Error> {
    let request = parse_constructor_request(payload).await?;
    let temp_path = Path::new(TEMP_CONSTRUCTOR_FILE);
    match request.action.as_deref() {
        Some("upload") => {
            if !temp_path.exists() {
                return Err(ErrorBadRequest("Файл не загружен"));
            }
            let xml = read_uploaded_docx_xml(temp_path)?;
            let html = docx_xml_to_html(&xml)
                .map_err(|e| ErrorInternalServerError(format!("Не удалось разобрать документ: {e}")))?;
            let mut context = Context::new();
            context.insert(
                "info_message",
                "Файл загружен. Выделяйте текст для добавления полей",
            );
            context.insert("html", &html);
            render_html(&tera, "constructor.html", &context)
        }
        Some("save") => {
            if !temp_path.exists() {
                return Err(ErrorBadRequest("Временный файл шаблона не найден"));
            }
            if request.template_name.trim().is_empty() {
                return Err(ErrorBadRequest("Имя шаблона не может быть пустым"));
            }
            let file_name = ensure_docx_name(&request.template_name)?;
            let settings = settings.read().unwrap();
            let save_dir = match validate_folder_name(&request.target_folder)? {
                Some(folder_name) => {
                    let dir = settings.templates_dir.join(folder_name);
                    fs::create_dir_all(&dir).map_err(ErrorInternalServerError)?;
                    dir
                }
                None => settings.templates_dir.clone(),
            };
            let save_path = save_dir.join(file_name);
            match apply_constructor_replacements(
                temp_path,
                &save_path,
                request.replacements,
                request.force,
            ) {
                Ok(_) => {
                    let _ = fs::remove_file(temp_path);
                    Ok(found("/"))
                }
                Err(error) => Ok(HttpResponse::BadRequest().json(json!({
                    "error": error.to_string()
                }))),
            }
        }
        _ => Err(ErrorBadRequest("Неизвестное действие")),
    }
}

// Для выбора сохранения в конструкторе
async fn get_template_folders(
    settings: web::Data<AppState>,
) -> Result<HttpResponse, Error> {
    let settings = settings.read().unwrap();
    let folders = group_template_files(&settings.templates_dir)?
        .into_iter()
        .map(|folder| folder.path_prefix.trim_end_matches('/').to_string())
        .collect::<Vec<_>>();
    Ok(HttpResponse::Ok().json(folders))
}

#[derive(Serialize)]
struct FolderInfo {
    name: String,
    path_prefix: String,
    files: Vec<String>,
}

// Группирует шаблоны по корню и одноуровневым подпапкам
fn group_template_files(templates_dir: &Path) -> Result<Vec<FolderInfo>, Error> {
    let mut folders = Vec::new();
    let mut subdirs = Vec::new();
    if templates_dir.is_dir() {
        for entry in fs::read_dir(templates_dir).map_err(ErrorInternalServerError)? {
            let entry = entry.map_err(ErrorInternalServerError)?;
            if entry.path().is_dir() {
                if let Some(name) = entry.file_name().to_str() {
                    subdirs.push(name.to_string());
                }
            }
        }
    }
    subdirs.sort();
    folders.push(FolderInfo {
        name: "doc_templates (корень)".to_string(),
        path_prefix: String::new(),
        files: list_docx_files(templates_dir),
    });
    for dir_name in subdirs {
        folders.push(FolderInfo {
            files: list_docx_files(&templates_dir.join(&dir_name)),
            path_prefix: format!("{dir_name}/"),
            name: dir_name,
        });
    }
    Ok(folders)
}

async fn index(
    tera: web::Data<Tera>,
    settings: web::Data<AppState>,
    query: web::Query<HashMap<String, String>>,
) -> Result<HttpResponse, Error> {
    let settings = settings.read().unwrap();
    let folders = group_template_files(&settings.templates_dir)
        .map_err(ErrorInternalServerError)?;

    let mut context = Context::new();
    context.insert("folders", &folders);
    if let Some(success) = query.get("success") {
        context.insert("success_message", success);
    }
    if let Some(error) = query.get("error") {
        context.insert("error_message", error);
    }

    render_html(&tera, "index.html", &context)
}

fn build_placeholder_structs(
    template_path: &Path,
) -> Result<Vec<Placeholder>, Box<dyn std::error::Error>> {
    let mut placeholders: Vec<String> = get_placeholders_from_template(template_path)?
        .into_iter()
        .collect();
    placeholders.sort();
    Ok(placeholders
        .into_iter()
        .map(|ph| {
            let (name, field_type, format) = parse_placeholder(&ph);
            Placeholder {
                ph,
                name,
                field_type,
                format,
            }
        })
        .collect())
}

// Если шаблон содержит некорректные плейсхолдеры, форма не показывается
async fn fill_get(
    tera: web::Data<Tera>,
    settings: web::Data<AppState>,
    path: web::Path<String>,
) -> Result<HttpResponse, Error> {
    let relative_path = path.into_inner();
    let settings = settings.read().unwrap();
    let template_path = resolve_template_file(&settings.templates_dir, &relative_path)?;
    if !template_path.exists() {
        return Err(ErrorNotFound("Шаблон не найден"));
    }
    let placeholder_structs =
        build_placeholder_structs(&template_path).map_err(ErrorInternalServerError)?;
    let mut context = Context::new();
    context.insert("filename", &relative_path);
    if placeholder_structs.is_empty() {
        let preview_html = read_docx_document_xml(&template_path)
            .and_then(|xml| docx_xml_to_html(&xml))
            .unwrap_or_default();
        context.insert("error_message", "в этом документе нет полей для ввода");
        context.insert("invalid_placeholders", &Vec::<String>::new());
        context.insert("preview_html", &preview_html);
    } else {
        context.insert("placeholders", &placeholder_structs);
        context.insert("error_message", &Option::<String>::None);
    }
    render_html(&tera, "fill.html", &context)
}

fn format_form_value(ph: &str, form: &HashMap<String, String>) -> String {
    let (_name, field_type, format) = parse_placeholder(ph);
    if field_type != "hardware_error" {
        let value = form.get(ph).cloned().unwrap_or_default();
        return format_value(&field_type, format.as_deref(), &value);
    }
    let measurement_raw = form.get(&format!("{ph}__measurement_result")).cloned().unwrap_or_default();
    let main_raw = form
        .get(&format!("{ph}__main_relative_error_percent"))
        .cloned()
        .unwrap_or_else(|| DEFAULT_MAIN_RELATIVE_ERROR_PERCENT.to_string());
    let additional_raw = form
        .get(&format!("{ph}__additional_relative_error_percent"))
        .cloned()
        .unwrap_or_else(|| DEFAULT_ADDITIONAL_RELATIVE_ERROR_PERCENT.to_string());
    match (
        parse_f64_loose(&measurement_raw),
        parse_f64_loose(&main_raw),
        parse_f64_loose(&additional_raw),
    ) {
        (Some(measurement), Some(main), Some(additional)) => {
            format_hardware_error_value(measurement, main, additional)
        }
        _ => String::new(),
    }
}

async fn fill_post(
    payload: Multipart,
    settings: web::Data<AppState>,
    path: web::Path<String>,
) -> Result<HttpResponse, Error> {
    let relative_path = path.into_inner();
    let form = parse_fill_form(payload).await?;
    let settings = settings.read().unwrap();
    let template_path = resolve_template_file(&settings.templates_dir, &relative_path)?;
    if !template_path.exists() {
        return Err(ErrorNotFound("Шаблон не найден"));
    }
    let placeholders =
        get_placeholders_from_template(&template_path).map_err(ErrorInternalServerError)?;
    if placeholders.is_empty() {
        return Ok(found(format!("/fill/{relative_path}?error=no_placeholders")));
    }
    let replacements: HashMap<String, String> = placeholders
        .iter()
        .map(|ph| (ph.clone(), format_form_value(ph, &form)))
        .collect();
    let template_base = template_path
        .file_stem()
        .and_then(|s| s.to_str())
        .unwrap_or("filled");
    let today = Local::now().format("%Y-%m-%d_%H-%M-%S").to_string();
    let output_name = format!("{template_base}-{today}.docx");
    let save_path = settings.saved_dir.join(output_name);
    fill_docx_template(&template_path, &save_path, &replacements)
        .map_err(ErrorInternalServerError)?;
    Ok(found("/history"))
}

async fn history(
    tera: web::Data<Tera>,
    settings: web::Data<AppState>,
) -> Result<HttpResponse, Error> {
    let settings = settings.read().unwrap();
    let files = list_files_sorted_by_mtime(&settings.saved_dir)?;
    let mut context = Context::new();
    context.insert("files", &files);
    render_html(&tera, "history.html", &context)
}

async fn serve_template(
    settings: web::Data<AppState>,
    path: web::Path<String>,
) -> Result<NamedFile, Error> {
    let settings = settings.read().unwrap();
    let file_path = resolve_template_file(&settings.templates_dir, &path.into_inner())?;
    NamedFile::open(file_path).map_err(ErrorNotFound)
}

async fn serve_saved(
    settings: web::Data<AppState>,
    path: web::Path<String>,
) -> Result<NamedFile, Error> {
    let settings = settings.read().unwrap();
    let file_path = settings.saved_dir.join(safe_file_name(&path.into_inner())?);
    NamedFile::open(file_path).map_err(ErrorNotFound)
}

fn delete_file_and_redirect(
    file_path: PathBuf,
    not_found_message: &'static str,
    redirect_to: &'static str,
) -> Result<HttpResponse, Error> {
    if !file_path.exists() {
        return Err(ErrorNotFound(not_found_message));
    }
    fs::remove_file(&file_path).map_err(ErrorInternalServerError)?;
    Ok(found(redirect_to))
}

async fn delete_template(
    settings: web::Data<AppState>,
    path: web::Path<String>,
) -> Result<HttpResponse, Error> {
    let settings = settings.read().unwrap();
    delete_file_and_redirect(
        resolve_template_file(&settings.templates_dir, &path.into_inner())?,
        "Шаблон не найден",
        "/",
    )
}

async fn delete_saved(
    settings: web::Data<AppState>,
    path: web::Path<String>,
) -> Result<HttpResponse, Error> {
    let settings = settings.read().unwrap();
    delete_file_and_redirect(
        settings.saved_dir.join(safe_file_name(&path.into_inner())?),
        "Файл не найден",
        "/history",
    )
}

/*===================== main =====================*/

fn ensure_custom_types_file() -> std::io::Result<()> {
    if !Path::new(CUSTOM_TYPES_FILE).exists() {
        fs::write(CUSTOM_TYPES_FILE, "[]")?;
    }
    Ok(())
}

#[actix_web::main]
async fn main() -> std::io::Result<()> {
    fs::create_dir_all(DOC_TEMPLATES_DIR)?;
    fs::create_dir_all(SAVED_DIR)?;
    ensure_custom_types_file()?;
    let settings = load_settings();
    let app_state: web::Data<AppState> = web::Data::new(Arc::new(RwLock::new(settings)));
    let tera = Tera::new(&format!("{TEMPLATES_DIR}/**/*"))
        .map_err(|e| std::io::Error::new(std::io::ErrorKind::Other, e))?;
    let server = HttpServer::new(move || {
        App::new()
            .app_data(web::Data::new(tera.clone()))
            .app_data(app_state.clone())
            .service(Files::new("/static", "./static").show_files_listing())
            .route("/", web::get().to(index))
            .route("/constructor", web::get().to(constructor_get))
            .route("/constructor", web::post().to(constructor_post))
            .route("/api/custom-types", web::get().to(custom_types_get))
            .route("/api/custom-types", web::post().to(custom_types_upsert))
            .route("/api/custom-types/{key}", web::delete().to(custom_types_delete))
            .route("/fill/{filename:.*}", web::get().to(fill_get))
            .route("/fill/{filename:.*}", web::post().to(fill_post))
            .route("/history", web::get().to(history))
            .route("/templates/{filename:.*}", web::get().to(serve_template))
            .route("/saved/{file}", web::get().to(serve_saved))
            .route("/delete_template/{filename:.*}", web::post().to(delete_template))
            .route("/delete_saved/{file}", web::post().to(delete_saved))
            .route("/api/template-folders", web::get().to(get_template_folders))
    })
    .bind(("127.0.0.1", 8080))?;
    let _ = webbrowser::open("http://127.0.0.1:8080/");
    server.run().await
}