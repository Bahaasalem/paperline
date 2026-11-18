use std::path::PathBuf;
use std::sync::Mutex;

use percent_encoding::percent_decode_str;
use tauri::ipc::{InvokeBody, Request, Response};
use tauri::{Emitter, Manager};

/// Files the system asked us to open before the page was ready to receive them.
#[derive(Default)]
struct Pending(Mutex<(Vec<String>, bool)>);

const OPENABLE: &[&str] = &[
    "pline", "docx", "xlsx", "pptx", "pdf", "csv", "tsv", "txt", "md", "html", "htm", "png", "jpg", "jpeg", "webp", "gif", "bmp",
];

fn openable(args: impl IntoIterator<Item = String>) -> Vec<String> {
    args.into_iter()
        .filter(|a| !a.starts_with('-'))
        .filter(|a| {
            let p = PathBuf::from(a);
            p.is_file()
                && p.extension()
                    .and_then(|e| e.to_str())
                    .map(|e| OPENABLE.contains(&e.to_ascii_lowercase().as_str()))
                    .unwrap_or(false)
        })
        .collect()
}

fn header(request: &Request, name: &str) -> Result<String, String> {
    let raw = request
        .headers()
        .get(name)
        .and_then(|v| v.to_str().ok())
        .ok_or_else(|| format!("missing {name}"))?;
    Ok(percent_decode_str(raw).decode_utf8_lossy().into_owned())
}

fn body<'a>(request: &'a Request<'_>) -> Result<&'a [u8], String> {
    match request.body() {
        InvokeBody::Raw(data) => Ok(data),
        _ => Err("expected raw bytes".into()),
    }
}

#[tauri::command]
fn take_opened_files(state: tauri::State<Pending>) -> Vec<String> {
    let mut p = state.0.lock().unwrap();
    p.1 = true; // the page is listening from now on
    std::mem::take(&mut p.0)
}

#[tauri::command]
fn read_file(path: String) -> Result<Response, String> {
    std::fs::read(&path).map(Response::new).map_err(|e| e.to_string())
}

/// Writes the bytes to a path the user picked in the save dialog.
#[tauri::command]
fn write_file(request: Request) -> Result<(), String> {
    let path = header(&request, "path")?;
    std::fs::write(path, body(&request)?).map_err(|e| e.to_string())
}

/// Printing on macOS: write the PDF to a temporary file and open it in the default viewer (Preview).
#[tauri::command]
fn open_for_print(app: tauri::AppHandle, request: Request) -> Result<(), String> {
    use tauri_plugin_opener::OpenerExt;
    let name = header(&request, "name")?;
    let safe: String = name.chars().map(|c| if r#"/\:*?"<>|"#.contains(c) { '_' } else { c }).collect();
    let dir = std::env::temp_dir().join("paperline-print");
    std::fs::create_dir_all(&dir).map_err(|e| e.to_string())?;
    let path = dir.join(safe);
    std::fs::write(&path, body(&request)?).map_err(|e| e.to_string())?;
    app.opener().open_path(path.to_string_lossy(), None::<&str>).map_err(|e| e.to_string())
}

fn deliver(app: &tauri::AppHandle, files: Vec<String>) {
    if files.is_empty() {
        return;
    }
    {
        // until the page asks for pending files it cannot hear events, so keep them
        let state = app.state::<Pending>();
        let mut p = state.0.lock().unwrap();
        if !p.1 {
            p.0.extend(files);
            return;
        }
    }
    let _ = app.emit("open-files", files);
    if let Some(w) = app.get_webview_window("main") {
        let _ = w.unminimize();
        let _ = w.set_focus();
    }
}

#[cfg_attr(mobile, tauri::mobile_entry_point)]
pub fn run() {
    let mut builder = tauri::Builder::default();
    #[cfg(desktop)]
    {
        builder = builder.plugin(tauri_plugin_single_instance::init(|app, argv, _cwd| {
            deliver(app, openable(argv.into_iter().skip(1)));
        }));
    }
    builder
        .plugin(tauri_plugin_dialog::init())
        .plugin(tauri_plugin_opener::init())
        .manage(Pending::default())
        .setup(|app| {
            // Windows and Linux pass files as arguments
            let files = openable(std::env::args().skip(1));
            app.state::<Pending>().0.lock().unwrap().0.extend(files);
            Ok(())
        })
        .invoke_handler(tauri::generate_handler![take_opened_files, read_file, write_file, open_for_print])
        .build(tauri::generate_context!())
        .expect("error while building Paperline")
        .run(|_app, _event| {
            // macOS sends files through an "open documents" event
            #[cfg(any(target_os = "macos", target_os = "ios"))]
            if let tauri::RunEvent::Opened { urls } = _event {
                let files = urls.into_iter().filter_map(|u| u.to_file_path().ok()).map(|p| p.to_string_lossy().into_owned()).collect();
                deliver(_app, files);
            }
        });
}
