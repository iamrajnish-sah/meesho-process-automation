"""Session-scoped upload storage for the web UI."""
import json
import os
import secrets
import shutil
import tempfile
from typing import Optional
from urllib.parse import urlparse

from werkzeug.utils import secure_filename

RAW_DATA_NAME = "raw_data"
MASTER_NAME = "master_workbook"
SECONDARY_NAME = "secondary_upload"
GEMINI_CSV_NAME = "gemini_converted_raw.csv"
META_NAME = "upload_meta.json"

ALLOWED_RAW = {".csv", ".xlsx", ".xlsm", ".xls"}
ALLOWED_MASTER = {".xlsx", ".xlsm"}
ALLOWED_SECONDARY = {
    ".png", ".jpg", ".jpeg", ".gif", ".webp",
    ".mp3", ".wav", ".m4a", ".ogg", ".aac",
    ".xlsx", ".xlsm", ".xls", ".pdf", ".docx",
}


def new_session_id() -> str:
    return secrets.token_hex(12)


def session_dir(session_id: str) -> str:
    base = os.path.join(tempfile.gettempdir(), "meesho_web_uploads")
    path = os.path.join(base, session_id)
    os.makedirs(path, exist_ok=True)
    return path


def _meta_path(session_id: str) -> str:
    return os.path.join(session_dir(session_id), META_NAME)


def load_meta(session_id: str) -> dict:
    path = _meta_path(session_id)
    if os.path.exists(path):
        with open(path, encoding="utf-8") as f:
            return json.load(f)
    return {}


def save_meta(session_id: str, meta: dict) -> None:
    with open(_meta_path(session_id), "w", encoding="utf-8") as f:
        json.dump(meta, f, indent=2)


def _save_upload(session_id: str, storage_name: str, file_storage, allowed_ext: set) -> str:
    if not file_storage or not file_storage.filename:
        raise ValueError("No file selected.")
    ext = os.path.splitext(file_storage.filename)[1].lower()
    if ext not in allowed_ext:
        raise ValueError(f"File type {ext!r} not allowed.")
    dest = os.path.join(session_dir(session_id), storage_name + ext)
    file_storage.save(dest)
    meta = load_meta(session_id)
    meta[storage_name] = dest
    meta[f"{storage_name}_original"] = secure_filename(file_storage.filename)
    save_meta(session_id, meta)
    return dest


def save_raw_data(session_id: str, file_storage) -> str:
    return _save_upload(session_id, RAW_DATA_NAME, file_storage, ALLOWED_RAW)


def save_master(session_id: str, file_storage) -> str:
    return _save_upload(session_id, MASTER_NAME, file_storage, ALLOWED_MASTER)


def save_secondary(session_id: str, file_storage) -> str:
    return _save_upload(session_id, SECONDARY_NAME, file_storage, ALLOWED_SECONDARY)


def save_secondary_url(session_id: str, url: str) -> str:
    url = (url or "").strip()
    if not url:
        raise ValueError("URL is empty.")
    parsed = urlparse(url)
    if parsed.scheme not in ("http", "https"):
        raise ValueError("URL must start with http:// or https://")

    try:
        import urllib.request

        req = urllib.request.Request(url, headers={"User-Agent": "MeeshoPipelineWeb/1.0"})
        with urllib.request.urlopen(req, timeout=30) as resp:
            data = resp.read()
            ctype = resp.headers.get("Content-Type", "")
    except Exception as exc:
        raise ValueError(f"Could not fetch URL: {exc}") from exc

    ext = os.path.splitext(parsed.path)[1].lower()
    if not ext:
        if "pdf" in ctype:
            ext = ".pdf"
        elif "png" in ctype:
            ext = ".png"
        elif "jpeg" in ctype or "jpg" in ctype:
            ext = ".jpg"
        else:
            ext = ".bin"

    dest = os.path.join(session_dir(session_id), SECONDARY_NAME + ext)
    with open(dest, "wb") as f:
        f.write(data)

    meta = load_meta(session_id)
    meta[SECONDARY_NAME] = dest
    meta[f"{SECONDARY_NAME}_original"] = url
    meta["secondary_source"] = "url"
    save_meta(session_id, meta)
    return dest


def save_gemini_csv(session_id: str, csv_text: str) -> str:
    dest = os.path.join(session_dir(session_id), GEMINI_CSV_NAME)
    with open(dest, "w", newline="", encoding="utf-8") as f:
        f.write(csv_text)
    meta = load_meta(session_id)
    meta["gemini_csv"] = dest
    save_meta(session_id, meta)
    return dest


def get_raw_data_path(session_id: str) -> Optional[str]:
    meta = load_meta(session_id)
    path = meta.get(RAW_DATA_NAME)
    if path and os.path.exists(path):
        return path
    gemini = meta.get("gemini_csv")
    return gemini if gemini and os.path.exists(gemini) else None


def get_master_path(session_id: str) -> Optional[str]:
    meta = load_meta(session_id)
    path = meta.get(MASTER_NAME)
    return path if path and os.path.exists(path) else None


def get_secondary_path(session_id: str) -> Optional[str]:
    meta = load_meta(session_id)
    path = meta.get(SECONDARY_NAME)
    return path if path and os.path.exists(path) else None


def cleanup_session(session_id: str) -> None:
    path = session_dir(session_id)
    if os.path.isdir(path):
        shutil.rmtree(path, ignore_errors=True)
