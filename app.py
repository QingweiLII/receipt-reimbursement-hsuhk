from __future__ import annotations

import base64
import datetime as dt
import html
import json
import mimetypes
import os
import re
import secrets
import socket
import subprocess
import sys
import traceback
import tempfile
import time
import uuid
import zipfile
from concurrent.futures import FIRST_COMPLETED, ThreadPoolExecutor, wait
from email.parser import BytesParser
from email.policy import default as email_policy
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from threading import Lock
from urllib import error, parse, request


APP_DIR = Path(__file__).resolve().parent
DATA_DIR = APP_DIR / "data"
UPLOAD_DIR = APP_DIR / "uploads"
CLIENTS_DIR = DATA_DIR / "clients"
RECORDS_PATH = DATA_DIR / "records.json"
WORKBOOK_PATH = DATA_DIR / "reimbursements.xlsx"
PAGE_PATHS = {"/", "/hsuhk-receipt-report-page", "/hsuhk receipt report page", "/hsuhk-receipt-report"}
DEFAULT_OCR_TIMEOUT_SECONDS = "20"
DEFAULT_OCR_TARGET_MIN_EDGE = "500"
DEFAULT_OCR_MAX_LONG_EDGE = "900"
DEFAULT_OCR_SMALL_MAX_LONG_EDGE = "650"
DEFAULT_OCR_PSMS = "6,11"
DEFAULT_OCR_MAX_CANDIDATES = "3"
DEFAULT_OCR_VARIANTS = "gray"
DEFAULT_LLM_TIMEOUT_SECONDS = "45"
DEFAULT_RECEIPT_PROCESS_TIMEOUT_SECONDS = "120"

HEADERS = [
    "Date",
    "Time",
    "Person",
    "Type",
    "Activities",
    "Location",
    "Address",
    "Amount in Local Currency",
    "Amount in HKD",
    "Source File",
    "Uploaded At",
    "Notes",
]

FIELD_ALIASES = {
    "date": ["date", "Date"],
    "time": ["time", "Time"],
    "person": ["person", "Person", "name", "guest_name", "passenger"],
    "type": ["type", "Type", "category", "expense_type"],
    "activities": ["activities", "Activities", "activity", "description"],
    "location": ["location", "Location", "city", "country"],
    "address": ["address", "Address"],
    "amount_local": [
        "amount_in_local_currency",
        "Amount in Local Currency",
        "amount_local",
        "local_amount",
        "amount",
    ],
    "amount_hkd": [
        "amount_in_hkd",
        "Amount in HKD",
        "amount_hkd",
        "hkd_amount",
    ],
    "currency": ["currency", "Currency", "local_currency"],
    "confidence": ["confidence", "Confidence"],
    "notes": ["notes", "Notes"],
}

STORE_LOCK = Lock()
GOOGLE_OAUTH_LOCK = Lock()
GOOGLE_OAUTH_STATES: dict[str, dict[str, str]] = {}
UPLOAD_JOBS_LOCK = Lock()
UPLOAD_JOB_EXECUTOR_LOCK = Lock()
UPLOAD_JOBS: dict[str, dict] = {}
UPLOAD_JOB_EXECUTOR: ThreadPoolExecutor | None = None
TERMINAL_UPLOAD_STATUSES = {"completed", "failed"}
FX_CACHE: dict[tuple[str, str], tuple[float, str]] = {}
EXPENSE_TYPES = ["Flight", "Meal", "Accommondation", "Transportation", "Others"]
GOOGLE_CLIENT_HEADER = "Client ID"
GOOGLE_SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]
GOOGLE_CREDENTIALS = None
GOOGLE_USER_SESSION_COOKIE = "receipt_google_session"
GOOGLE_USER_SESSIONS_PATH = DATA_DIR / "google_user_sessions.json"
GOOGLE_OAUTH_SCOPES = [
    "openid",
    "email",
    "profile",
    "https://www.googleapis.com/auth/drive.readonly",
    "https://www.googleapis.com/auth/drive.file",
]
GOOGLE_DRIVE_RECORDS_FILE_NAME = "HSUHK Receipt Report Records.json"
GOOGLE_DRIVE_LONG_TERM_WORKBOOK_NAME = "HSUHK Receipt Report Long-term Records.xlsx"
GOOGLE_DRIVE_RECENT_WORKBOOK_NAME = "HSUHK Receipt Report Last 7 Days.xlsx"


class ExtractionError(Exception):
    pass


def env_flag(name: str, default: bool = False) -> bool:
    value = os.environ.get(name)
    if value is None:
        return default
    return value.strip().lower() in {"1", "true", "yes", "on"}


def storage_provider() -> str:
    return os.environ.get("STORAGE_PROVIDER", "local").strip().lower()


def google_sheets_enabled() -> bool:
    return storage_provider() in {"google_sheets", "google_sheet"}


def user_google_drive_enabled() -> bool:
    return storage_provider() in {"user_google_drive", "google_drive_oauth", "google_drive"} or env_flag(
        "GOOGLE_DRIVE_USER_OAUTH", False
    )


def google_oauth_configured() -> bool:
    return bool(os.environ.get("GOOGLE_OAUTH_CLIENT_ID", "").strip() and os.environ.get("GOOGLE_OAUTH_CLIENT_SECRET", "").strip())


def google_drive_folder_name() -> str:
    return os.environ.get("GOOGLE_DRIVE_FOLDER_NAME", "HSUHK Receipt Reports").strip() or "HSUHK Receipt Reports"


def google_oauth_scopes() -> list[str]:
    raw = os.environ.get("GOOGLE_OAUTH_SCOPES", "").strip()
    if not raw:
        return GOOGLE_OAUTH_SCOPES
    return [item.strip() for item in re.split(r"[\s,]+", raw) if item.strip()]


def google_sheet_id() -> str:
    return os.environ.get("GOOGLE_SHEET_ID", "").strip()


def google_sheet_name() -> str:
    return os.environ.get("GOOGLE_SHEET_NAME", "Receipts").strip() or "Receipts"


def google_sheet_store_client_id() -> bool:
    return env_flag("GOOGLE_SHEETS_STORE_CLIENT_ID", True)


def google_sheet_filter_by_client() -> bool:
    return env_flag("GOOGLE_SHEETS_FILTER_BY_CLIENT", True)


def google_headers() -> list[str]:
    headers = list(HEADERS)
    if google_sheet_store_client_id():
        headers.append(GOOGLE_CLIENT_HEADER)
    return headers


def google_record_to_row(record: dict, client_id: str | None) -> list[str]:
    row = record_to_row(record)
    if google_sheet_store_client_id():
        row.append(safe_client_id(client_id))
    return row


def quote_sheet_range(sheet_name: str, range_ref: str) -> str:
    escaped = sheet_name.replace("'", "''")
    return f"'{escaped}'!{range_ref}"


def load_env_file(path: Path = APP_DIR / ".env") -> None:
    if not path.exists():
        return
    for raw_line in path.read_text(encoding="utf-8").splitlines():
        line = raw_line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        key, value = line.split("=", 1)
        key = key.strip()
        value = value.strip().strip('"').strip("'")
        if key and key not in os.environ:
            os.environ[key] = value


def load_google_service_account_info() -> dict:
    raw = os.environ.get("GOOGLE_SERVICE_ACCOUNT_JSON", "").strip()
    if raw:
        try:
            if raw.startswith("{"):
                return json.loads(raw)
            return json.loads(base64.b64decode(raw).decode("utf-8"))
        except Exception as exc:
            raise ExtractionError("GOOGLE_SERVICE_ACCOUNT_JSON must be raw JSON or base64-encoded service account JSON.") from exc

    file_path = os.environ.get("GOOGLE_SERVICE_ACCOUNT_FILE", "").strip()
    if file_path:
        try:
            return json.loads(Path(file_path).read_text(encoding="utf-8"))
        except Exception as exc:
            raise ExtractionError(f"Could not read GOOGLE_SERVICE_ACCOUNT_FILE: {file_path}") from exc

    raise ExtractionError(
        "Google Sheets storage is enabled, but no service account credentials were provided. "
        "Set GOOGLE_SERVICE_ACCOUNT_JSON or GOOGLE_SERVICE_ACCOUNT_FILE."
    )


def google_credentials():
    global GOOGLE_CREDENTIALS
    if GOOGLE_CREDENTIALS is not None:
        return GOOGLE_CREDENTIALS
    try:
        from google.oauth2 import service_account  # type: ignore
        from google.auth.transport.requests import Request as GoogleAuthRequest  # type: ignore
    except Exception as exc:
        raise ExtractionError(
            "Google Sheets storage needs google-auth and requests. Install requirements.txt before enabling STORAGE_PROVIDER=google_sheets."
        ) from exc

    info = load_google_service_account_info()
    credentials = service_account.Credentials.from_service_account_info(info, scopes=GOOGLE_SCOPES)
    credentials.refresh(GoogleAuthRequest())
    GOOGLE_CREDENTIALS = credentials
    return credentials


def google_access_token() -> str:
    credentials = google_credentials()
    if not credentials.valid:
        try:
            from google.auth.transport.requests import Request as GoogleAuthRequest  # type: ignore
        except Exception as exc:
            raise ExtractionError("google-auth transport is not available.") from exc
        credentials.refresh(GoogleAuthRequest())
    return credentials.token


def google_sheets_request(method: str, path: str, payload: dict | None = None, query: dict | None = None) -> dict:
    sheet_id = google_sheet_id()
    if not sheet_id:
        raise ExtractionError("STORAGE_PROVIDER=google_sheets requires GOOGLE_SHEET_ID.")
    url = f"https://sheets.googleapis.com/v4/spreadsheets/{parse.quote(sheet_id)}{path}"
    if query:
        url += "?" + parse.urlencode(query)
    body = json.dumps(payload).encode("utf-8") if payload is not None else None
    req = request.Request(
        url,
        data=body,
        headers={
            "Authorization": f"Bearer {google_access_token()}",
            "Content-Type": "application/json",
            "Accept": "application/json",
        },
        method=method,
    )
    try:
        with request.urlopen(req, timeout=float(os.environ.get("GOOGLE_API_TIMEOUT_SECONDS", "20"))) as resp:
            raw = resp.read().decode("utf-8")
            return json.loads(raw) if raw else {}
    except error.HTTPError as exc:
        details = exc.read().decode("utf-8", errors="replace")[:1200]
        raise ExtractionError(f"Google Sheets API returned HTTP {exc.code}: {details}") from exc
    except error.URLError as exc:
        raise ExtractionError(f"Google Sheets API request failed: {exc}") from exc
    except json.JSONDecodeError as exc:
        raise ExtractionError("Google Sheets API returned non-JSON response.") from exc


def google_values_range(range_a1: str) -> str:
    return "/values/" + parse.quote(range_a1, safe="")


def google_get_values(range_a1: str) -> list[list[str]]:
    data = google_sheets_request("GET", google_values_range(range_a1))
    values = data.get("values", [])
    if isinstance(values, list):
        return [[str(cell) for cell in row] for row in values if isinstance(row, list)]
    return []


def google_update_values(range_a1: str, values: list[list[str]]) -> None:
    google_sheets_request(
        "PUT",
        google_values_range(range_a1),
        {"values": values},
        {"valueInputOption": "RAW"},
    )


def google_append_values(range_a1: str, values: list[list[str]]) -> None:
    google_sheets_request(
        "POST",
        google_values_range(range_a1) + ":append",
        {"values": values},
        {"valueInputOption": "RAW", "insertDataOption": "INSERT_ROWS"},
    )


def ensure_google_sheet() -> None:
    headers = google_headers()
    sheet_name = google_sheet_name()
    metadata = google_sheets_request("GET", "", query={"fields": "sheets.properties(title)"})
    titles = [
        sheet.get("properties", {}).get("title")
        for sheet in metadata.get("sheets", [])
        if isinstance(sheet, dict)
    ]
    if sheet_name not in titles:
        google_sheets_request("POST", ":batchUpdate", {"requests": [{"addSheet": {"properties": {"title": sheet_name}}}]})

    header_range = quote_sheet_range(sheet_name, f"A1:{column_name(len(headers))}1")
    current = google_get_values(header_range)
    if not current or current[0][: len(headers)] != headers:
        google_update_values(header_range, [headers])


def read_google_records(client_id: str | None = None) -> list[dict]:
    ensure_google_sheet()
    headers = google_headers()
    sheet_name = google_sheet_name()
    values = google_get_values(quote_sheet_range(sheet_name, f"A1:{column_name(len(headers))}"))
    if not values:
        return []

    sheet_headers = values[0]
    header_indexes = {header: index for index, header in enumerate(sheet_headers)}
    safe_id = safe_client_id(client_id)
    records: list[dict] = []
    for row in values[1:]:
        if not any(str(cell).strip() for cell in row):
            continue
        if google_sheet_filter_by_client() and client_id is not None and GOOGLE_CLIENT_HEADER in header_indexes:
            row_client = row[header_indexes[GOOGLE_CLIENT_HEADER]] if header_indexes[GOOGLE_CLIENT_HEADER] < len(row) else ""
            if safe_client_id(row_client) != safe_id:
                continue
        record: dict[str, str] = {}
        for header, key in {
            "Date": "date",
            "Time": "time",
            "Person": "person",
            "Type": "type",
            "Activities": "activities",
            "Location": "location",
            "Address": "address",
            "Amount in Local Currency": "amount_local",
            "Amount in HKD": "amount_hkd",
            "Source File": "source_file",
            "Uploaded At": "uploaded_at",
            "Notes": "notes",
        }.items():
            index = header_indexes.get(header)
            record[key] = row[index] if index is not None and index < len(row) else ""
        records.append(record)
    return records


def append_google_records(records: list[dict], client_id: str | None = None) -> None:
    if not records:
        return
    ensure_google_sheet()
    headers = google_headers()
    rows = [google_record_to_row(record, client_id) for record in records]
    google_append_values(quote_sheet_range(google_sheet_name(), f"A1:{column_name(len(headers))}"), rows)


def unix_now() -> int:
    return int(dt.datetime.now(dt.timezone.utc).timestamp())


def load_google_user_sessions() -> dict[str, dict]:
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    if not GOOGLE_USER_SESSIONS_PATH.exists():
        return {}
    try:
        data = json.loads(GOOGLE_USER_SESSIONS_PATH.read_text(encoding="utf-8"))
        return data if isinstance(data, dict) else {}
    except json.JSONDecodeError:
        return {}


def save_google_user_sessions(sessions: dict[str, dict]) -> None:
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    GOOGLE_USER_SESSIONS_PATH.write_text(json.dumps(sessions, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")


def google_token_request(payload: dict[str, str]) -> dict:
    body = parse.urlencode(payload).encode("utf-8")
    req = request.Request(
        "https://oauth2.googleapis.com/token",
        data=body,
        headers={"Content-Type": "application/x-www-form-urlencoded", "Accept": "application/json"},
        method="POST",
    )
    try:
        with request.urlopen(req, timeout=float(os.environ.get("GOOGLE_API_TIMEOUT_SECONDS", "20"))) as resp:
            return json.loads(resp.read().decode("utf-8"))
    except error.HTTPError as exc:
        details = exc.read().decode("utf-8", errors="replace")[:1200]
        raise ExtractionError(f"Google OAuth returned HTTP {exc.code}: {details}") from exc
    except error.URLError as exc:
        raise ExtractionError(f"Google OAuth request failed: {exc}") from exc


def refresh_google_user_session(session: dict) -> bool:
    expires_at = int(session.get("expires_at") or 0)
    if expires_at > unix_now() + 90:
        return False
    refresh_token = str(session.get("refresh_token") or "")
    if not refresh_token:
        raise ExtractionError("Google Drive session expired. Please reconnect Google Drive.")
    data = google_token_request(
        {
            "client_id": os.environ.get("GOOGLE_OAUTH_CLIENT_ID", "").strip(),
            "client_secret": os.environ.get("GOOGLE_OAUTH_CLIENT_SECRET", "").strip(),
            "refresh_token": refresh_token,
            "grant_type": "refresh_token",
        }
    )
    access_token = data.get("access_token")
    if not access_token:
        raise ExtractionError("Google OAuth refresh did not return an access token.")
    session["access_token"] = access_token
    session["expires_at"] = unix_now() + int(data.get("expires_in") or 3600)
    return True


def google_userinfo(access_token: str) -> dict:
    req = request.Request(
        "https://www.googleapis.com/oauth2/v3/userinfo",
        headers={"Authorization": f"Bearer {access_token}", "Accept": "application/json"},
    )
    try:
        with request.urlopen(req, timeout=float(os.environ.get("GOOGLE_API_TIMEOUT_SECONDS", "20"))) as resp:
            data = json.loads(resp.read().decode("utf-8"))
            return data if isinstance(data, dict) else {}
    except error.HTTPError as exc:
        details = exc.read().decode("utf-8", errors="replace")[:1200]
        raise ExtractionError(f"Google userinfo returned HTTP {exc.code}: {details}") from exc
    except error.URLError as exc:
        raise ExtractionError(f"Google userinfo request failed: {exc}") from exc


def escape_drive_query(value: str) -> str:
    return value.replace("\\", "\\\\").replace("'", "\\'")


def google_drive_api_url(path: str, query: dict | None = None, upload: bool = False) -> str:
    base = "https://www.googleapis.com/upload/drive/v3" if upload else "https://www.googleapis.com/drive/v3"
    url = base + path
    if query:
        url += "?" + parse.urlencode(query)
    return url


def google_drive_json(session: dict, method: str, path: str, payload: dict | None = None, query: dict | None = None) -> dict:
    refresh_google_user_session(session)
    body = json.dumps(payload).encode("utf-8") if payload is not None else None
    req = request.Request(
        google_drive_api_url(path, query),
        data=body,
        headers={
            "Authorization": f"Bearer {session['access_token']}",
            "Content-Type": "application/json",
            "Accept": "application/json",
        },
        method=method,
    )
    try:
        with request.urlopen(req, timeout=float(os.environ.get("GOOGLE_API_TIMEOUT_SECONDS", "30"))) as resp:
            raw = resp.read().decode("utf-8")
            return json.loads(raw) if raw else {}
    except error.HTTPError as exc:
        details = exc.read().decode("utf-8", errors="replace")[:1200]
        raise ExtractionError(f"Google Drive API returned HTTP {exc.code}: {details}") from exc
    except error.URLError as exc:
        raise ExtractionError(f"Google Drive API request failed: {exc}") from exc


def google_drive_bytes(session: dict, path: str, query: dict | None = None) -> bytes:
    refresh_google_user_session(session)
    req = request.Request(
        google_drive_api_url(path, query),
        headers={"Authorization": f"Bearer {session['access_token']}"},
    )
    try:
        with request.urlopen(req, timeout=float(os.environ.get("GOOGLE_API_TIMEOUT_SECONDS", "60"))) as resp:
            return resp.read()
    except error.HTTPError as exc:
        details = exc.read().decode("utf-8", errors="replace")[:1200]
        raise ExtractionError(f"Google Drive download returned HTTP {exc.code}: {details}") from exc
    except error.URLError as exc:
        raise ExtractionError(f"Google Drive download failed: {exc}") from exc


def google_drive_upload(
    session: dict,
    name: str,
    content: bytes,
    mime_type: str,
    parent_id: str | None = None,
    existing_id: str | None = None,
) -> dict:
    refresh_google_user_session(session)
    metadata: dict[str, object] = {"name": name}
    if parent_id and not existing_id:
        metadata["parents"] = [parent_id]
    boundary = f"receipt_boundary_{uuid.uuid4().hex}"
    body = b"".join(
        [
            f"--{boundary}\r\nContent-Type: application/json; charset=UTF-8\r\n\r\n".encode("utf-8"),
            json.dumps(metadata).encode("utf-8"),
            f"\r\n--{boundary}\r\nContent-Type: {mime_type}\r\n\r\n".encode("utf-8"),
            content,
            f"\r\n--{boundary}--\r\n".encode("utf-8"),
        ]
    )
    path = f"/files/{parse.quote(existing_id)}" if existing_id else "/files"
    req = request.Request(
        google_drive_api_url(path, {"uploadType": "multipart", "fields": "id,name,webViewLink,webContentLink"}, upload=True),
        data=body,
        headers={
            "Authorization": f"Bearer {session['access_token']}",
            "Content-Type": f"multipart/related; boundary={boundary}",
            "Accept": "application/json",
        },
        method="PATCH" if existing_id else "POST",
    )
    try:
        with request.urlopen(req, timeout=float(os.environ.get("GOOGLE_API_TIMEOUT_SECONDS", "60"))) as resp:
            return json.loads(resp.read().decode("utf-8"))
    except error.HTTPError as exc:
        details = exc.read().decode("utf-8", errors="replace")[:1200]
        raise ExtractionError(f"Google Drive upload returned HTTP {exc.code}: {details}") from exc
    except error.URLError as exc:
        raise ExtractionError(f"Google Drive upload failed: {exc}") from exc


def find_drive_file(session: dict, name: str, mime_type: str | None = None, parent_id: str | None = None) -> dict | None:
    clauses = [f"name = '{escape_drive_query(name)}'", "trashed = false"]
    if mime_type:
        clauses.append(f"mimeType = '{escape_drive_query(mime_type)}'")
    if parent_id:
        clauses.append(f"'{escape_drive_query(parent_id)}' in parents")
    data = google_drive_json(
        session,
        "GET",
        "/files",
        query={
            "q": " and ".join(clauses),
            "pageSize": "1",
            "fields": "files(id,name,mimeType,webViewLink,webContentLink)",
        },
    )
    files = data.get("files", [])
    return files[0] if isinstance(files, list) and files else None


def ensure_drive_folder(session: dict) -> str:
    folder = find_drive_file(session, google_drive_folder_name(), "application/vnd.google-apps.folder")
    if folder and folder.get("id"):
        return str(folder["id"])
    created = google_drive_json(
        session,
        "POST",
        "/files",
        {
            "name": google_drive_folder_name(),
            "mimeType": "application/vnd.google-apps.folder",
        },
        {"fields": "id,name,webViewLink"},
    )
    return str(created.get("id") or "")


def list_drive_receipt_files(session: dict, search: str = "") -> list[dict]:
    clauses = ["trashed = false", "(mimeType = 'application/pdf' or mimeType contains 'image/')"]
    if search.strip():
        clauses.append(f"name contains '{escape_drive_query(search.strip())}'")
    data = google_drive_json(
        session,
        "GET",
        "/files",
        query={
            "q": " and ".join(clauses),
            "pageSize": os.environ.get("GOOGLE_DRIVE_PICKER_PAGE_SIZE", "30"),
            "orderBy": "modifiedTime desc",
            "fields": "files(id,name,mimeType,size,modifiedTime,webViewLink)",
        },
    )
    files = data.get("files", [])
    return [item for item in files if isinstance(item, dict)]


def drive_upload_item(session: dict, file_id: str) -> dict:
    safe_id = re.sub(r"[^A-Za-z0-9_-]+", "", file_id)
    if not safe_id:
        raise ExtractionError("Invalid Google Drive file id.")
    metadata = google_drive_json(session, "GET", f"/files/{parse.quote(safe_id)}", query={"fields": "id,name,mimeType,size"})
    mime_type = str(metadata.get("mimeType") or "")
    if not (mime_type == "application/pdf" or mime_type.startswith("image/")):
        raise ExtractionError(f"Google Drive file is not a supported image/PDF: {metadata.get('name', safe_id)}")
    size = int(metadata.get("size") or 0)
    max_mb = float(os.environ.get("MAX_UPLOAD_MB", "20"))
    if size and size > max_mb * 1024 * 1024:
        raise ExtractionError(f"Google Drive file is larger than MAX_UPLOAD_MB={max_mb}: {metadata.get('name', safe_id)}")
    return {
        "filename": safe_filename(str(metadata.get("name") or safe_id)),
        "content_type": mime_type,
        "content": google_drive_bytes(session, f"/files/{parse.quote(safe_id)}", {"alt": "media"}),
    }


def read_user_drive_records(session: dict) -> tuple[list[dict], str | None]:
    folder_id = ensure_drive_folder(session)
    existing = find_drive_file(session, GOOGLE_DRIVE_RECORDS_FILE_NAME, "application/json", folder_id)
    if not existing or not existing.get("id"):
        return [], None
    raw = google_drive_bytes(session, f"/files/{parse.quote(str(existing['id']))}", {"alt": "media"})
    try:
        data = json.loads(raw.decode("utf-8"))
    except json.JSONDecodeError:
        return [], str(existing["id"])
    if isinstance(data, dict) and isinstance(data.get("records"), list):
        return [item for item in data["records"] if isinstance(item, dict)], str(existing["id"])
    if isinstance(data, list):
        return [item for item in data if isinstance(item, dict)], str(existing["id"])
    return [], str(existing["id"])


def write_user_drive_records(session: dict, records: list[dict], existing_id: str | None = None) -> dict:
    folder_id = ensure_drive_folder(session)
    payload = {
        "records": records,
        "updated_at": now_local(),
        "columns": HEADERS,
    }
    return google_drive_upload(
        session,
        GOOGLE_DRIVE_RECORDS_FILE_NAME,
        json.dumps(payload, ensure_ascii=False, indent=2).encode("utf-8"),
        "application/json",
        folder_id,
        existing_id,
    )


def workbook_bytes(records: list[dict], client_id: str, label: str) -> bytes:
    export_dir = CLIENTS_DIR / safe_client_id(client_id) / "drive_exports"
    export_path = export_dir / f"{safe_filename(label)}-{uuid.uuid4().hex[:8]}.xlsx"
    write_workbook(records, export_path)
    return export_path.read_bytes()


def save_user_drive_workbook(session: dict, records: list[dict], filename: str, client_id: str, replace_existing: bool) -> dict:
    folder_id = ensure_drive_folder(session)
    existing_id = None
    if replace_existing:
        existing = find_drive_file(
            session,
            filename,
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            folder_id,
        )
        existing_id = str(existing["id"]) if existing and existing.get("id") else None
    return google_drive_upload(
        session,
        filename,
        workbook_bytes(records, client_id, filename),
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        folder_id,
        existing_id,
    )


def save_records_to_user_drive(session: dict, records: list[dict], save_mode: str, client_id: str) -> dict:
    mode = save_mode if save_mode in {"batch", "recent7", "append"} else "batch"
    if mode == "batch":
        stamp = dt.datetime.now().strftime("%Y%m%d-%H%M%S")
        return save_user_drive_workbook(session, records, f"HSUHK Receipt Report - {stamp}.xlsx", client_id, False)

    existing_records, records_file_id = read_user_drive_records(session)
    existing_records.extend(records)
    write_user_drive_records(session, existing_records, records_file_id)
    if mode == "recent7":
        return save_user_drive_workbook(
            session,
            filter_recent_records(existing_records, 7),
            GOOGLE_DRIVE_RECENT_WORKBOOK_NAME,
            client_id,
            True,
        )
    return save_user_drive_workbook(
        session,
        existing_records,
        GOOGLE_DRIVE_LONG_TERM_WORKBOOK_NAME,
        client_id,
        True,
    )


def ensure_dirs() -> None:
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    UPLOAD_DIR.mkdir(parents=True, exist_ok=True)
    CLIENTS_DIR.mkdir(parents=True, exist_ok=True)
    if not RECORDS_PATH.exists():
        RECORDS_PATH.write_text("[]\n", encoding="utf-8")
        write_workbook([])


def now_local() -> str:
    return dt.datetime.now().astimezone().isoformat(timespec="seconds")


def safe_filename(filename: str) -> str:
    name = Path(filename or "upload").name
    name = re.sub(r"[^A-Za-z0-9._ -]+", "_", name).strip(" .")
    return name or "upload"


def safe_client_id(client_id: str | None) -> str:
    value = (client_id or "default").strip()
    value = re.sub(r"[^A-Za-z0-9_-]+", "_", value).strip("_")
    return value[:80] or "default"


def client_paths(client_id: str | None) -> tuple[Path, Path, Path]:
    safe_id = safe_client_id(client_id)
    client_dir = CLIENTS_DIR / safe_id
    return client_dir, client_dir / "records.json", client_dir / "reimbursements.xlsx"


def batch_workbook_path(client_id: str, batch_id: str) -> Path:
    return CLIENTS_DIR / safe_client_id(client_id) / "batches" / safe_client_id(batch_id) / "reimbursements.xlsx"


def ensure_client_store(client_id: str | None) -> tuple[Path, Path, Path]:
    ensure_dirs()
    client_dir, records_path, workbook_path = client_paths(client_id)
    client_dir.mkdir(parents=True, exist_ok=True)
    if not records_path.exists():
        records_path.write_text("[]\n", encoding="utf-8")
        write_workbook([], workbook_path)
    return client_dir, records_path, workbook_path


def read_records(client_id: str | None = None) -> list[dict]:
    if google_sheets_enabled():
        return read_google_records(client_id)
    ensure_dirs()
    if client_id is None:
        records_path = RECORDS_PATH
    else:
        _, records_path, _ = ensure_client_store(client_id)
    try:
        return json.loads(records_path.read_text(encoding="utf-8"))
    except json.JSONDecodeError:
        return []


def save_records(records: list[dict], client_id: str | None = None) -> None:
    if google_sheets_enabled():
        raise ExtractionError("Replacing all records is not supported for Google Sheets storage; append records instead.")
    ensure_dirs()
    if client_id is None:
        records_path = RECORDS_PATH
    else:
        _, records_path, _ = ensure_client_store(client_id)
    records_path.write_text(json.dumps(records, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")


def append_records(records: list[dict], client_id: str | None = None) -> list[dict]:
    with STORE_LOCK:
        if google_sheets_enabled():
            append_google_records(records, client_id)
            return read_google_records(client_id)
        existing = read_records(client_id)
        existing.extend(records)
        save_records(existing, client_id)
        workbook_path = WORKBOOK_PATH if client_id is None else client_paths(client_id)[2]
        write_workbook(existing, workbook_path)
        return existing


def filter_recent_records(records: list[dict], days: int = 7) -> list[dict]:
    cutoff = dt.datetime.now().astimezone() - dt.timedelta(days=days)
    recent: list[dict] = []
    for record in records:
        timestamp = str(record.get("uploaded_at", "")).strip()
        try:
            uploaded_at = dt.datetime.fromisoformat(timestamp)
            if uploaded_at.tzinfo is None:
                uploaded_at = uploaded_at.replace(tzinfo=dt.datetime.now().astimezone().tzinfo)
        except ValueError:
            continue
        if uploaded_at >= cutoff:
            recent.append(record)
    return recent


def get_first(obj: dict, aliases: list[str]) -> str:
    for key in aliases:
        if key in obj and obj[key] not in (None, ""):
            return str(obj[key]).strip()
    return ""


def parse_amount_and_currency(amount_text: str, explicit_currency: str = "") -> tuple[float | None, str]:
    text = str(amount_text or "")
    currency = explicit_currency.strip().upper()
    if not currency:
        for code in ["HKD", "HK$", "HKS", "HK§", "USD", "US$", "CNY", "RMB", "DKK", "SGD", "EUR", "GBP", "JPY"]:
            if code in text.upper():
                currency = {"HK$": "HKD", "HKS": "HKD", "HK§": "HKD", "US$": "USD", "RMB": "CNY"}.get(code, code)
                break
    match = re.search(r"[-+]?\d[\d,]*(?:\.\d+)?", text)
    if not match:
        return None, currency
    try:
        number = match.group(0)
        if "," in number and "." not in number and re.search(r",\d{1,2}$", number):
            number = number.replace(",", ".")
        else:
            number = number.replace(",", "")
        amount = float(number)
    except ValueError:
        return None, currency
    return amount, currency


def format_money(value: float) -> str:
    return f"{value:.2f}"


def load_fx_rates() -> dict[str, float]:
    raw = os.environ.get("FX_RATES_JSON", "").strip()
    if not raw:
        return {}
    try:
        parsed = json.loads(raw)
        return {str(k).upper(): float(v) for k, v in parsed.items()}
    except (TypeError, ValueError, json.JSONDecodeError):
        return {}


def normalize_currency(currency: str) -> str:
    value = currency.strip().upper()
    return {"HK$": "HKD", "HKS": "HKD", "HK§": "HKD", "US$": "USD", "RMB": "CNY", "CN¥": "CNY", "￥": "CNY"}.get(value, value)


def infer_currency_from_context(context: str) -> str:
    text = context.lower()
    if re.search(r"\bhkd\b|hk[$s§]", context, flags=re.IGNORECASE):
        return "HKD"
    if any(term in text for term in ["hong kong", " shatin", "tst", "k11", "musea", "homesquare", "tokachi-milky", "kpay"]):
        return "HKD"
    return ""


def parse_iso_date(value: str) -> dt.date | None:
    text = str(value or "").strip()
    if not text:
        return None
    for fmt in ("%Y-%m-%d", "%Y/%m/%d", "%d-%m-%Y", "%d/%m/%Y", "%d-%m-%y", "%d/%m/%y"):
        try:
            return dt.datetime.strptime(text, fmt).date()
        except ValueError:
            pass
    match = re.search(r"(20\d{2})[-/.](\d{1,2})[-/.](\d{1,2})", text)
    if match:
        try:
            return dt.date(int(match.group(1)), int(match.group(2)), int(match.group(3)))
        except ValueError:
            return None
    return None


def fetch_frankfurter_rate(currency: str, receipt_date: dt.date) -> tuple[float | None, str, str]:
    base_url = os.environ.get("FX_BASE_URL", "https://api.frankfurter.dev/v2").rstrip("/")
    lookback_days = int(os.environ.get("FX_LOOKBACK_DAYS", "7"))
    for offset in range(max(lookback_days, 0) + 1):
        rate_date = receipt_date - dt.timedelta(days=offset)
        cache_key = (currency, rate_date.isoformat())
        if cache_key in FX_CACHE:
            rate, used_date = FX_CACHE[cache_key]
            return rate, used_date, "cache"
        query = parse.urlencode({"date": rate_date.isoformat(), "base": currency, "quotes": "HKD"})
        url = f"{base_url}/rates?{query}"
        req = request.Request(
            url,
            headers={
                "Accept": "application/json",
                "User-Agent": "receipt-reimbursement-app/0.1",
            },
        )
        try:
            with request.urlopen(req, timeout=float(os.environ.get("FX_TIMEOUT_SECONDS", "8"))) as resp:
                payload = json.loads(resp.read().decode("utf-8"))
        except Exception:
            continue

        rate: float | None = None
        if isinstance(payload, list):
            for item in payload:
                if item.get("base") == currency and item.get("quote") == "HKD" and item.get("rate") is not None:
                    rate = float(item["rate"])
                    break
        elif isinstance(payload, dict):
            if payload.get("rate") is not None:
                rate = float(payload["rate"])
            elif isinstance(payload.get("rates"), dict) and payload["rates"].get("HKD") is not None:
                rate = float(payload["rates"]["HKD"])

        if rate:
            used_date = rate_date.isoformat()
            FX_CACHE[cache_key] = (rate, used_date)
            return rate, used_date, "frankfurter"
    return None, "", "unavailable"


def get_hkd_rate(currency: str, receipt_date: dt.date | None) -> tuple[float | None, str]:
    currency = normalize_currency(currency)
    if not currency:
        return None, "FX skipped: local currency was not detected."
    if currency == "HKD":
        return 1.0, "FX rate used: HKD/HKD 1.000000."
    if not receipt_date:
        return None, "FX skipped: receipt date was missing or not parseable."

    provider = os.environ.get("EXCHANGE_RATE_PROVIDER", "frankfurter").strip().lower()
    if provider in {"none", "off", "disabled"}:
        return None, "FX skipped: EXCHANGE_RATE_PROVIDER is disabled."
    if provider == "static":
        rate = load_fx_rates().get(currency)
        if rate:
            return rate, f"FX rate used: {currency}/HKD {rate:.6f} from FX_RATES_JSON."
        return None, f"FX skipped: no static rate for {currency}."

    rate, used_date, source = fetch_frankfurter_rate(currency, receipt_date)
    if rate:
        date_note = used_date
        if used_date != receipt_date.isoformat():
            date_note += f" (nearest prior available rate for {receipt_date.isoformat()})"
        return rate, f"FX rate used: {currency}/HKD {rate:.6f} on {date_note} via Frankfurter."
    return None, f"FX unavailable: could not fetch {currency}/HKD for {receipt_date.isoformat()}."


def classify_expense(raw_type: str, activities: str, source_file: str) -> str:
    value = str(raw_type or "").strip()
    for allowed in EXPENSE_TYPES:
        if value.lower() == allowed.lower() and allowed != "Others":
            return allowed
    text = f"{value} {activities} {source_file}".lower()
    if re.search(r"\b(flight|airline|airport|air ticket|e-ticket|boarding|thai airways|trip\.com)\b", text):
        return "Flight"
    if re.search(r"\b(meal|restaurant|cafe|coffee|lunch|dinner|breakfast|food|bakery|dessert|refreshment|croissant|baguette|chocolate|tokachi|gcap|茶餐廳|餐|飯|奶茶)\b", text):
        return "Meal"
    if re.search(r"\b(hotel|accommodation|accomodation|room|folio|lodging|crowne plaza)\b", text):
        return "Accommondation"
    if re.search(r"\b(taxi|uber|train|bus|metro|subway|transport|fare|ride|km)\b", text):
        return "Transportation"
    return "Others"


def maybe_fill_hkd(amount_local: str, amount_hkd: str, currency: str, receipt_date: str) -> tuple[str, str]:
    if amount_hkd:
        return amount_hkd, "FX skipped: Amount in HKD was already supplied by the model or receipt."
    amount, detected_currency = parse_amount_and_currency(amount_local, currency)
    if amount is None:
        return "", "FX skipped: local amount was not parseable."
    detected_currency = normalize_currency(detected_currency)
    rate, note = get_hkd_rate(detected_currency, parse_iso_date(receipt_date))
    if rate:
        return format_money(amount * rate), note
    return "", note


def amount_with_currency(amount: float, currency: str) -> str:
    return f"{format_money(amount)} {normalize_currency(currency)}".strip()


def numeric_amounts(text: str) -> list[float]:
    amounts: list[float] = []
    for match in re.finditer(r"[-+]?(?:\d{1,3}(?:,\d{3})+|\d+)(?:\.\d+)?", text):
        before = text[match.start() - 1] if match.start() > 0 else ""
        after = text[match.end()] if match.end() < len(text) else ""
        if (before and before in "-/:") or (after and after in "-/:"):
            continue
        amount, _ = parse_amount_and_currency(match.group(0), "")
        if amount is not None and 0 < amount < 100000:
            amounts.append(amount)
    return amounts


def preferred_total_from_ocr(ocr_text: str) -> float | None:
    if not ocr_text:
        return None
    first_candidate = re.split(r"\n\nOCR candidate 2\b", ocr_text, maxsplit=1)[0]
    lines = [line.strip() for line in first_candidate.splitlines() if line.strip()]

    priority_amounts: list[float] = []
    for line in lines:
        if re.search(r"\b(total|kpay|paid|visa|master|fare)\b", line, flags=re.IGNORECASE):
            priority_amounts.extend(numeric_amounts(line))
    if priority_amounts:
        return max(priority_amounts)

    hk_amounts: list[float] = []
    for match in re.finditer(r"(?:HK[$S§]?|HKS|HRS|HES|FRG|\$)\s*([-+]?\d{1,6}(?:[,.]\d{1,2})?)", first_candidate, flags=re.IGNORECASE):
        amount, _ = parse_amount_and_currency(match.group(1), "")
        if amount is not None and 0 < amount < 100000:
            hk_amounts.append(amount)
    if hk_amounts:
        return max(hk_amounts)
    return None


def merge_records_for_file(raw_records: list[dict]) -> dict:
    if not raw_records:
        return {"notes": "Model returned no records."}
    if len(raw_records) == 1:
        return raw_records[0]

    primary = dict(raw_records[-1])
    amounts: list[float] = []
    currencies: list[str] = []
    detail_notes: list[str] = []

    for index, record in enumerate(raw_records, start=1):
        amount_text = get_first(record, FIELD_ALIASES["amount_local"])
        currency = normalize_currency(get_first(record, FIELD_ALIASES["currency"]))
        amount, detected_currency = parse_amount_and_currency(amount_text, currency)
        if amount is not None:
            amounts.append(amount)
            currencies.append(normalize_currency(detected_currency))
        description = get_first(record, FIELD_ALIASES["activities"]) or get_first(record, ["description"])
        date = get_first(record, FIELD_ALIASES["date"])
        note_piece = " ".join(part for part in [date, description, amount_text] if part)
        if note_piece:
            detail_notes.append(f"{index}. {note_piece}")

    common_currency = currencies[0] if currencies and all(currency == currencies[0] for currency in currencies) else ""
    total = sum(amounts) if amounts and common_currency else None

    notes = get_first(primary, FIELD_ALIASES["notes"])
    merged_note = f"Merged {len(raw_records)} line items from one receipt into one reimbursement row."
    if detail_notes:
        merged_note += " Details: " + "; ".join(detail_notes)
    primary["notes"] = f"{notes} {merged_note}".strip()
    if total is not None:
        primary["amount_in_local_currency"] = amount_with_currency(total, common_currency)
        primary["currency"] = common_currency
        primary["amount_in_hkd"] = ""
    return primary


def normalize_record(raw: dict, source_file: str, stored_file: str) -> dict:
    receipt_date = get_first(raw, FIELD_ALIASES["date"])
    activities = get_first(raw, FIELD_ALIASES["activities"])
    notes = get_first(raw, FIELD_ALIASES["notes"])
    amount_local = get_first(raw, FIELD_ALIASES["amount_local"])
    amount_hkd = get_first(raw, FIELD_ALIASES["amount_hkd"])
    currency = normalize_currency(get_first(raw, FIELD_ALIASES["currency"]))
    amount_value = get_first(raw, ["amount_local_value", "local_amount_value", "amount_value"])
    if not currency and amount_local:
        _, embedded_currency = parse_amount_and_currency(amount_local, "")
        currency = normalize_currency(embedded_currency)
    ocr_context = str(raw.get("_ocr_text", ""))
    context_for_currency = " ".join(
        part
        for part in [
            source_file,
            activities,
            get_first(raw, FIELD_ALIASES["location"]),
            get_first(raw, FIELD_ALIASES["address"]),
            amount_local,
            get_first(raw, FIELD_ALIASES["amount_hkd"]),
            notes,
        ]
        if part
    )
    currency_overridden = False
    if (
        currency == "JPY"
        and re.search(r"\b(kpay|tokachi|k11)\b", ocr_context, flags=re.IGNORECASE)
        and not re.search(r"\b(jpy|yen)\b|[¥円]", ocr_context, flags=re.IGNORECASE)
    ):
        currency = "HKD"
        currency_overridden = True
        amount_local = re.sub(r"\bJPY\b|\bYEN\b|[¥円]", "", amount_local, flags=re.IGNORECASE).strip()
        notes = re.sub(r"\bJPY\b|\bYEN\b|\bHokkaido\b|\bJapan\b|[¥円]", "", notes, flags=re.IGNORECASE).strip(" ;,.")
    if not currency:
        currency = infer_currency_from_context(f"{context_for_currency} {ocr_context[:5000]}")
    if not amount_local and amount_value:
        amount_local = f"{amount_value} {currency}".strip()
    elif amount_local and currency and currency not in amount_local.upper():
        amount_local = f"{amount_local} {currency}".strip()

    preferred_total = preferred_total_from_ocr(ocr_context)
    if preferred_total is not None and currency == "HKD":
        current_amount, _ = parse_amount_and_currency(amount_local, currency)
        current_hkd, _ = parse_amount_and_currency(amount_hkd, "HKD")
        tolerance = max(1.0, preferred_total * 0.02)
        if current_amount is None or abs(current_amount - preferred_total) <= tolerance or not amount_hkd:
            if current_amount is not None and abs(current_amount - preferred_total) > 0.005:
                replacement = format_money(preferred_total)
                for token in {format_money(current_amount), str(current_amount), str(current_amount).replace(".", ",")}:
                    notes = notes.replace(token, replacement)
            amount_local = amount_with_currency(preferred_total, "HKD")
        if current_hkd is None or abs(current_hkd - preferred_total) <= tolerance:
            amount_hkd = ""
    amount_hkd, fx_note = maybe_fill_hkd(amount_local, amount_hkd, currency, receipt_date)
    if fx_note:
        notes = f"{notes} {fx_note}".strip()
    if currency_overridden:
        notes = f"{notes} Currency corrected to HKD from OCR context.".strip()

    return {
        "id": str(uuid.uuid4()),
        "date": receipt_date,
        "time": get_first(raw, FIELD_ALIASES["time"]),
        "person": get_first(raw, FIELD_ALIASES["person"]),
        "type": classify_expense(get_first(raw, FIELD_ALIASES["type"]), activities, source_file),
        "activities": activities,
        "location": get_first(raw, FIELD_ALIASES["location"]),
        "address": get_first(raw, FIELD_ALIASES["address"]),
        "amount_local": amount_local,
        "amount_hkd": amount_hkd,
        "source_file": source_file,
        "stored_file": stored_file,
        "uploaded_at": now_local(),
        "confidence": get_first(raw, FIELD_ALIASES["confidence"]),
        "notes": notes,
    }


def as_records(raw: object) -> list[dict]:
    if isinstance(raw, dict) and isinstance(raw.get("records"), list):
        return [item for item in raw["records"] if isinstance(item, dict)]
    if isinstance(raw, dict) and isinstance(raw.get("entries"), list):
        return [item for item in raw["entries"] if isinstance(item, dict)]
    if isinstance(raw, dict):
        return [raw]
    if isinstance(raw, list):
        return [item for item in raw if isinstance(item, dict)]
    return []


def column_name(index: int) -> str:
    letters = ""
    while index:
        index, rem = divmod(index - 1, 26)
        letters = chr(65 + rem) + letters
    return letters


def column_index(name: str) -> int:
    index = 0
    for char in name:
        if "A" <= char <= "Z":
            index = index * 26 + (ord(char) - 64)
    return index


def is_number(value: object) -> bool:
    if isinstance(value, (int, float)):
        return True
    if isinstance(value, str):
        return bool(re.fullmatch(r"[-+]?\d+(?:\.\d+)?", value.strip()))
    return False


def xml_text(value: object) -> str:
    return html.escape(str(value), quote=True)


def cell_xml(row_index: int, col_index: int, value: object, style_id: int = 0) -> str:
    ref = f"{column_name(col_index)}{row_index}"
    style = f' s="{style_id}"' if style_id else ""
    if value in (None, ""):
        return f'<c r="{ref}"{style}/>'
    if is_number(value) and col_index == 9:
        return f'<c r="{ref}" s="2"><v>{xml_text(str(value).replace(",", ""))}</v></c>'
    return f'<c r="{ref}"{style} t="inlineStr"><is><t>{xml_text(value)}</t></is></c>'


def row_xml(row_index: int, values: list[object], header: bool = False) -> str:
    cells = "".join(cell_xml(row_index, i + 1, value, 1 if header else 0) for i, value in enumerate(values))
    return f'<row r="{row_index}">{cells}</row>'


def record_to_row(record: dict) -> list[str]:
    return [
        record.get("date", ""),
        record.get("time", ""),
        record.get("person", ""),
        record.get("type", ""),
        record.get("activities", ""),
        record.get("location", ""),
        record.get("address", ""),
        record.get("amount_local", ""),
        record.get("amount_hkd", ""),
        record.get("source_file", ""),
        record.get("uploaded_at", ""),
        record.get("notes", ""),
    ]


def read_generated_workbook_rows(workbook_path: Path) -> list[list[str]]:
    with zipfile.ZipFile(workbook_path, "r") as workbook:
        xml = workbook.read("xl/worksheets/sheet1.xml").decode("utf-8")
    rows: list[list[str]] = []
    for row_match in re.finditer(r"<row\b[^>]*>(.*?)</row>", xml, flags=re.DOTALL):
        values_by_col: dict[int, str] = {}
        for cell_match in re.finditer(r'<c\b([^>]*)>(.*?)</c>|<c\b([^>]*)/>', row_match.group(1), flags=re.DOTALL):
            attrs = cell_match.group(1) or cell_match.group(3) or ""
            ref_match = re.search(r'r="([A-Z]+)\d+"', attrs)
            if not ref_match:
                continue
            col = column_index(ref_match.group(1))
            body = cell_match.group(2) or ""
            text_match = re.search(r"<t[^>]*>(.*?)</t>", body, flags=re.DOTALL)
            value_match = re.search(r"<v>(.*?)</v>", body, flags=re.DOTALL)
            value = text_match.group(1) if text_match else value_match.group(1) if value_match else ""
            values_by_col[col] = html.unescape(value)
        if values_by_col:
            rows.append([values_by_col.get(index, "") for index in range(1, max(values_by_col) + 1)])
    return rows


def records_from_generated_workbook(workbook_path: Path) -> list[dict]:
    rows = read_generated_workbook_rows(workbook_path)
    if not rows:
        return []
    headers = rows[0]
    index_by_header = {header: index for index, header in enumerate(headers)}
    key_by_header = {
        "Date": "date",
        "Time": "time",
        "Person": "person",
        "Type": "type",
        "Activities": "activities",
        "Location": "location",
        "Address": "address",
        "Amount in Local Currency": "amount_local",
        "Amount in HKD": "amount_hkd",
        "Source File": "source_file",
        "Uploaded At": "uploaded_at",
        "Confidence": "confidence",
        "Notes": "notes",
    }
    records: list[dict] = []
    for row in rows[1:]:
        record: dict[str, str] = {}
        for header, key in key_by_header.items():
            index = index_by_header.get(header)
            if index is not None and index < len(row):
                record[key] = row[index]
        records.append(record)
    return records


def migrate_generated_workbook_columns(workbook_path: Path) -> None:
    try:
        rows = read_generated_workbook_rows(workbook_path)
    except Exception:
        return
    if rows and "Confidence" in rows[0]:
        write_workbook(records_from_generated_workbook(workbook_path), workbook_path)


def write_workbook(records: list[dict], workbook_path: Path = WORKBOOK_PATH) -> None:
    ensure_dirs_no_workbook()
    workbook_path.parent.mkdir(parents=True, exist_ok=True)
    rows = [HEADERS] + [record_to_row(record) for record in records]
    row_count = max(len(rows), 1)
    col_count = len(HEADERS)
    sheet_rows = [row_xml(1, HEADERS, True)]
    for index, row in enumerate(rows[1:], start=2):
        sheet_rows.append(row_xml(index, row))

    now = dt.datetime.now(dt.timezone.utc).replace(microsecond=0).isoformat().replace("+00:00", "Z")
    dimension = f"A1:{column_name(col_count)}{row_count}"
    worksheet_xml = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <dimension ref="{dimension}"/>
  <sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>
  <sheetFormatPr defaultRowHeight="18"/>
  <cols>
    <col min="1" max="2" width="14" customWidth="1"/>
    <col min="3" max="3" width="18" customWidth="1"/>
    <col min="4" max="4" width="18" customWidth="1"/>
    <col min="5" max="5" width="32" customWidth="1"/>
    <col min="6" max="7" width="28" customWidth="1"/>
    <col min="8" max="9" width="22" customWidth="1"/>
    <col min="10" max="11" width="26" customWidth="1"/>
    <col min="12" max="12" width="56" customWidth="1"/>
  </cols>
  <sheetData>
    {''.join(sheet_rows)}
  </sheetData>
  <autoFilter ref="{dimension}"/>
</worksheet>'''

    workbook_xml = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Receipts" sheetId="1" r:id="rId1"/></sheets>
</workbook>'''

    workbook_rels = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>'''

    styles_xml = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <numFmts count="1"><numFmt numFmtId="164" formatCode="#,##0.00"/></numFmts>
  <fonts count="2">
    <font><sz val="11"/><name val="Aptos"/></font>
    <font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Aptos"/></font>
  </fonts>
  <fills count="3">
    <fill><patternFill patternType="none"/></fill>
    <fill><patternFill patternType="gray125"/></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FF243B53"/><bgColor indexed="64"/></patternFill></fill>
  </fills>
  <borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
  <cellXfs count="3">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
    <xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1"/>
    <xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>
  </cellXfs>
  <cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
</styleSheet>'''

    rels_xml = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>'''

    content_types = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
</Types>'''

    core_xml = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
  xmlns:dc="http://purl.org/dc/elements/1.1/"
  xmlns:dcterms="http://purl.org/dc/terms/"
  xmlns:dcmitype="http://purl.org/dc/dcmitype/"
  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:title>Receipt Reimbursements</dc:title>
  <dc:creator>Receipt Reimbursement App</dc:creator>
  <cp:lastModifiedBy>Receipt Reimbursement App</cp:lastModifiedBy>
  <dcterms:created xsi:type="dcterms:W3CDTF">{now}</dcterms:created>
  <dcterms:modified xsi:type="dcterms:W3CDTF">{now}</dcterms:modified>
</cp:coreProperties>'''

    app_xml = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
  xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
  <Application>Receipt Reimbursement App</Application>
</Properties>'''

    with zipfile.ZipFile(workbook_path, "w", zipfile.ZIP_DEFLATED) as workbook:
        workbook.writestr("[Content_Types].xml", content_types)
        workbook.writestr("_rels/.rels", rels_xml)
        workbook.writestr("xl/workbook.xml", workbook_xml)
        workbook.writestr("xl/_rels/workbook.xml.rels", workbook_rels)
        workbook.writestr("xl/styles.xml", styles_xml)
        workbook.writestr("xl/worksheets/sheet1.xml", worksheet_xml)
        workbook.writestr("docProps/core.xml", core_xml)
        workbook.writestr("docProps/app.xml", app_xml)


def ensure_dirs_no_workbook() -> None:
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    UPLOAD_DIR.mkdir(parents=True, exist_ok=True)
    CLIENTS_DIR.mkdir(parents=True, exist_ok=True)


SYSTEM_PROMPT = """You extract reimbursement data from receipts, invoices, tickets, taxi receipts, hotel folios, and payment screenshots.
Return JSON only. Do not wrap it in markdown.
If a field is not visible, use an empty string.
Use ISO date format YYYY-MM-DD when possible and 24-hour time HH:MM when possible.
Do not invent an HKD amount. Only fill amount_in_hkd if the receipt explicitly shows HKD or the local currency is HKD.
Set type to exactly one of: Flight, Meal, Accommondation, Transportation, Others.
Return exactly one reimbursement record per uploaded file.
If a receipt contains multiple line items, nights, products, taxes, or partial charges, do not return separate rows. Use the receipt's grand total or total paid as amount_in_local_currency. Summarize the line items in notes, for example "5 accommodation nights at 1,750 DKK each; total incl. VAT 8,750 DKK."
OCR text may contain recognition mistakes from phone photos. Correct obvious receipt OCR confusions such as HK§/HKS/HKS -> HK$, commas vs decimal points, and noisy separators. Use the source filename as a weak hint for type, but do not invent fields that are not visible."""

USER_PROMPT = """Extract the following fields for reimbursement:
- date
- time
- person
- type
- activities
- location
- address
- amount_in_local_currency
- amount_in_hkd
- confidence
- notes

Important:
- Return one record only for this file.
- Use the grand total / total paid / total incl. VAT, not individual line items.
- Put itemized detail or number of nights/items in notes.
- If several OCR candidates are provided, compare them and use the clearest values.
- For Hong Kong receipts showing HKD, HK$, HKS, or HK§, use HKD as the local currency and set amount_in_hkd to the same total.
- Prefer totals near labels such as Total, Total Fare, KPay, Visa/Master/AE, paid, or Total HKD.

Return exactly this JSON shape:
{"records":[{"date":"","time":"","person":"","type":"","activities":"","location":"","address":"","amount_in_local_currency":"","amount_in_hkd":"","confidence":"","notes":""}]}"""


def extract_json(text: str) -> object:
    cleaned = text.strip()
    cleaned = re.sub(r"^```(?:json)?", "", cleaned, flags=re.IGNORECASE).strip()
    cleaned = re.sub(r"```$", "", cleaned).strip()
    try:
        return json.loads(cleaned)
    except json.JSONDecodeError:
        first_obj = cleaned.find("{")
        last_obj = cleaned.rfind("}")
        first_arr = cleaned.find("[")
        last_arr = cleaned.rfind("]")
        candidates = []
        if first_obj != -1 and last_obj != -1 and last_obj > first_obj:
            candidates.append(cleaned[first_obj : last_obj + 1])
        if first_arr != -1 and last_arr != -1 and last_arr > first_arr:
            candidates.append(cleaned[first_arr : last_arr + 1])
        for candidate in candidates:
            try:
                return json.loads(candidate)
            except json.JSONDecodeError:
                continue
    raise ExtractionError("The model did not return valid JSON.")


def temp_dir_parent() -> str | None:
    candidates = [
        os.environ.get("RECEIPT_TMP_DIR", "").strip(),
        tempfile.gettempdir(),
        "/tmp",
        "/private/tmp",
    ]
    for raw_path in candidates:
        if not raw_path:
            continue
        path = Path(raw_path)
        try:
            path.mkdir(parents=True, exist_ok=True)
        except Exception:
            continue
        if path.is_dir() and os.access(path, os.W_OK):
            return str(path)
    return None


def read_pdf_text(path: Path) -> str:
    text = ""
    try:
        import pypdf  # type: ignore
    except Exception:
        pypdf = None  # type: ignore
    if pypdf is not None:
        try:
            reader = pypdf.PdfReader(str(path))
            pages = [(page.extract_text() or "") for page in reader.pages]
            text = "\n\n".join(pages).strip()
        except Exception:
            text = ""
    if text:
        return text
    return read_pdf_ocr_text(path)


def read_pdf_ocr_text(path: Path) -> str:
    size = os.environ.get("PDF_OCR_RENDER_SIZE", "2400")
    with tempfile.TemporaryDirectory(prefix="receipt-pdf-ocr-", dir=temp_dir_parent()) as temp_dir:
        try:
            result = subprocess.run(
                ["qlmanage", "-t", "-s", size, "-o", temp_dir, str(path)],
                check=False,
                capture_output=True,
                text=True,
                timeout=float(os.environ.get("PDF_RENDER_TIMEOUT_SECONDS", "30")),
            )
        except FileNotFoundError:
            return ""
        except subprocess.TimeoutExpired:
            return ""
        if result.returncode != 0:
            return ""

        images = sorted(Path(temp_dir).glob("*.png")) + sorted(Path(temp_dir).glob("*.jpg")) + sorted(Path(temp_dir).glob("*.jpeg"))
        texts = []
        for image_path in images:
            try:
                ocr_text = read_image_text(image_path)
            except ExtractionError:
                ocr_text = ""
            if ocr_text:
                texts.append(ocr_text)
        return "\n\n".join(texts).strip()


def run_tesseract(path: Path, psm: str | None = None) -> tuple[str, str, int]:
    command = [
        "tesseract",
        str(path),
        "stdout",
        "-l",
        os.environ.get("OCR_LANG", "eng"),
        "--oem",
        os.environ.get("OCR_OEM", "1"),
    ]
    if psm:
        command.extend(["--psm", psm])
    dpi = os.environ.get("OCR_DPI", "150").strip()
    if dpi:
        command.extend(["--dpi", dpi])
    config_values = os.environ.get(
        "OCR_CONFIG",
        "load_system_dawg=0,load_freq_dawg=0,tessedit_do_invert=0",
    )
    for item in config_values.split(","):
        item = item.strip()
        if item:
            command.extend(["-c", item])
    try:
        start = time.time()
        result = subprocess.run(
            command,
            check=False,
            capture_output=True,
            text=True,
            timeout=float(os.environ.get("OCR_TIMEOUT_SECONDS", DEFAULT_OCR_TIMEOUT_SECONDS)),
        )
    except FileNotFoundError:
        raise ExtractionError("MiniMax provider needs local OCR for images, but tesseract is not installed.")
    except subprocess.TimeoutExpired:
        return "", "Local OCR timed out before the receipt text could be extracted.", 124
    elapsed = time.time() - start
    if elapsed >= 3 or result.returncode != 0:
        print(
            f"OCR finished in {elapsed:.1f}s with code {result.returncode} "
            f"for {path.name} psm={psm or 'default'} chars={len(result.stdout.strip())}"
        )
    return result.stdout.strip(), (result.stderr or "").strip(), result.returncode


def ocr_score(text: str) -> int:
    cleaned = text.strip()
    if not cleaned:
        return 0
    lower = cleaned.lower()
    keyword_hits = sum(
        lower.count(keyword)
        for keyword in [
            "total",
            "total fare",
            "date",
            "invoice",
            "receipt",
            "address",
            "hkd",
            "hk$",
            "hks",
            "kpay",
            "visa",
            "master",
            "taxi",
            "ikea",
            "gcap",
            "tokachi",
            "hotel",
        ]
    )
    amount_hits = len(re.findall(r"(?:HK[$S§]?|[$€£¥])?\s*\d{1,4}(?:[,.]\d{1,2})", cleaned, flags=re.IGNORECASE))
    date_hits = len(re.findall(r"\b(?:20\d{2}[-/]\d{1,2}[-/]\d{1,2}|\d{1,2}[-/]\d{1,2}[-/]\d{2,4})\b", cleaned))
    time_hits = len(re.findall(r"\b\d{1,2}:\d{2}(?::\d{2})?\b", cleaned))
    useful_length = min(len(cleaned), 800)
    odd_chars = len(re.findall(r"[~{}^_|]{2,}|[^\w\s.,:;/()$€£¥%#+&@-]", cleaned))
    return int(useful_length + keyword_hits * 100 + amount_hits * 18 + date_hits * 80 + time_hits * 40 - odd_chars * 2)


def resize_for_ocr(image, target_min: int | None = None):
    if target_min is None:
        target_min = int(os.environ.get("OCR_TARGET_MIN_EDGE", DEFAULT_OCR_TARGET_MIN_EDGE))
    max_long = int(os.environ.get("OCR_MAX_LONG_EDGE", DEFAULT_OCR_MAX_LONG_EDGE))
    min_edge = max(1, min(image.size))
    max_edge = max(image.size)
    scale = 1.0
    if min_edge < target_min:
        scale = max(scale, target_min / min_edge)
    if max_edge * scale > max_long:
        scale = max_long / max_edge
    if abs(scale - 1.0) < 0.05:
        return image
    return image.resize((max(1, int(image.width * scale)), max(1, int(image.height * scale))))


def resize_to_max_long_edge(image, max_long: int):
    max_edge = max(image.size)
    if max_edge <= max_long:
        return image
    scale = max_long / max_edge
    return image.resize((max(1, int(image.width * scale)), max(1, int(image.height * scale))))


def crop_for_ocr(image):
    if not env_flag("OCR_CROP", True):
        return image
    try:
        from PIL import ImageOps  # type: ignore
    except Exception:
        return image
    gray = ImageOps.grayscale(image)
    threshold = int(os.environ.get("OCR_CROP_THRESHOLD", "245"))
    mask = gray.point(lambda pixel: 255 if pixel < threshold else 0)
    bbox = mask.getbbox()
    if not bbox:
        return image
    left, top, right, bottom = bbox
    if right <= left or bottom <= top:
        return image
    pad = max(8, int(max(image.size) * 0.025))
    left = max(0, left - pad)
    top = max(0, top - pad)
    right = min(image.width, right + pad)
    bottom = min(image.height, bottom + pad)
    cropped_area = (right - left) * (bottom - top)
    original_area = image.width * image.height
    if cropped_area < original_area * 0.03:
        return image
    return image.crop((left, top, right, bottom))


def collect_ocr_candidates(path: Path) -> list[tuple[int, str, str]]:
    candidates: list[tuple[int, str, str]] = []
    raw_text = ""
    raw_code = 0
    raw_error = ""
    if env_flag("OCR_RAW_ORIGINAL", False):
        raw_text, raw_error, raw_code = run_tesseract(path)
        if raw_text:
            candidates.append((max(0, ocr_score(raw_text) - 80), "raw", raw_text))

    try:
        from PIL import Image, ImageOps  # type: ignore
    except Exception:
        if raw_code != 0 and not raw_text:
            details = raw_error[:600]
            raise ExtractionError(f"Local OCR failed: {details}")
        return candidates

    try:
        with Image.open(path) as opened:
            base = ImageOps.exif_transpose(opened).convert("RGB")
    except Exception:
        return candidates

    base = resize_for_ocr(crop_for_ocr(base))
    enabled_variants = {
        item.strip().lower()
        for item in os.environ.get("OCR_VARIANTS", DEFAULT_OCR_VARIANTS).split(",")
        if item.strip()
    }
    variants = []
    gray = ImageOps.autocontrast(ImageOps.grayscale(base))
    if "gray" in enabled_variants:
        variants.append(("gray", gray))
    if "rgb" in enabled_variants:
        variants.append(("rgb", base))
    if env_flag("OCR_SMALL_RETRY", True):
        small_max = int(os.environ.get("OCR_SMALL_MAX_LONG_EDGE", DEFAULT_OCR_SMALL_MAX_LONG_EDGE))
        if max(base.size) > small_max:
            small = resize_to_max_long_edge(base, small_max)
            variants.append(("gray-small", ImageOps.autocontrast(ImageOps.grayscale(small))))

    if base.width > base.height * 1.15:
        if "rot90" in enabled_variants or "rotations" in enabled_variants:
            variants.append(("rgb-rot90", base.rotate(90, expand=True)))
        if "rot270" in enabled_variants or "rotations" in enabled_variants:
            variants.append(("rgb-rot270", base.rotate(270, expand=True)))

    psms = [item.strip() for item in os.environ.get("OCR_PSMS", DEFAULT_OCR_PSMS).split(",") if item.strip()]
    with tempfile.TemporaryDirectory(prefix="receipt-image-ocr-", dir=temp_dir_parent()) as temp_dir:
        for variant_name, image in variants:
            image_path = Path(temp_dir) / f"{variant_name}.png"
            image.save(image_path)
            for psm in psms:
                text, _, code = run_tesseract(image_path, psm)
                if text and code == 0:
                    candidates.append((ocr_score(text), f"{variant_name} psm {psm}", text))
    return candidates


def dedupe_ocr_candidates(candidates: list[tuple[int, str, str]]) -> list[tuple[int, str, str]]:
    unique: list[tuple[int, str, str]] = []
    seen: set[str] = set()
    for score, label, text in sorted(candidates, key=lambda item: item[0], reverse=True):
        key = re.sub(r"\W+", "", text.lower())[:500]
        if not key or key in seen:
            continue
        seen.add(key)
        unique.append((score, label, text))
        if len(unique) >= int(os.environ.get("OCR_MAX_CANDIDATES", DEFAULT_OCR_MAX_CANDIDATES)):
            break
    return unique


def format_ocr_candidates(candidates: list[tuple[int, str, str]]) -> str:
    if not candidates:
        return ""
    if len(candidates) == 1:
        return candidates[0][2].strip()
    parts = []
    for index, (score, label, text) in enumerate(candidates, start=1):
        parts.append(f"OCR candidate {index} ({label}, score {score}):\n{text.strip()}")
    return "\n\n".join(parts).strip()


def read_image_text(path: Path) -> str:
    if os.environ.get("OCR_PREPROCESS", "1").strip().lower() not in {"0", "false", "off", "no"}:
        candidates = dedupe_ocr_candidates(collect_ocr_candidates(path))
        if candidates:
            return format_ocr_candidates(candidates)
        raise ExtractionError("Local OCR failed: no readable text was extracted from the receipt image.")

    text, details, code = run_tesseract(path)
    if code != 0:
        raise ExtractionError(f"Local OCR failed: {details[:600]}")
    return text.strip()


def strip_model_thinking(text: str) -> str:
    return re.sub(r"<think>.*?</think>", "", text, flags=re.IGNORECASE | re.DOTALL).strip()


def base64_file(path: Path) -> str:
    return base64.b64encode(path.read_bytes()).decode("ascii")


def post_json(url: str, payload: dict, headers: dict[str, str], timeout: int = 120) -> dict:
    body = json.dumps(payload).encode("utf-8")
    req = request.Request(
        url,
        data=body,
        headers={"Content-Type": "application/json", **headers},
        method="POST",
    )
    try:
        with request.urlopen(req, timeout=timeout) as resp:
            raw = resp.read().decode("utf-8")
            return json.loads(raw)
    except error.HTTPError as exc:
        details = exc.read().decode("utf-8", errors="replace")[:1200]
        raise ExtractionError(f"API returned HTTP {exc.code}: {details}") from exc
    except error.URLError as exc:
        raise ExtractionError(f"API request failed: {exc}") from exc
    except (TimeoutError, socket.timeout) as exc:
        raise ExtractionError(f"API request timed out after {timeout} seconds.") from exc
    except json.JSONDecodeError as exc:
        raise ExtractionError("API returned non-JSON response.") from exc


def parse_any_date(value: str) -> dt.date | None:
    for match in re.finditer(r"\b(?:20\d{2}[-/.]\d{1,2}[-/.]\d{1,2}|\d{1,2}[-/.]\d{1,2}[-/.]\d{2,4})\b", value):
        parsed = parse_iso_date(match.group(0))
        if parsed:
            return parsed
    return None


def all_text_dates(text: str) -> list[dt.date]:
    dates: list[dt.date] = []
    for match in re.finditer(r"\b(?:20\d{2}[-/.]\d{1,2}[-/.]\d{1,2}|\d{1,2}[-/.]\d{1,2}[-/.]\d{2,4})\b", text):
        parsed = parse_iso_date(match.group(0))
        if parsed:
            dates.append(parsed)
    return dates


def amount_candidates_with_context(text: str) -> list[tuple[float, str, str]]:
    candidates: list[tuple[float, str, str]] = []
    global_currency = infer_currency_from_context(text)
    for line in text.splitlines():
        if line.strip().lower().startswith("ocr candidate"):
            continue
        currency = normalize_currency(parse_amount_and_currency(line, "")[1])
        if not currency:
            for code in ["HKD", "DKK", "USD", "CNY", "SGD", "EUR", "GBP", "JPY"]:
                if re.search(rf"\b{re.escape(code)}\b", line, flags=re.IGNORECASE):
                    currency = code
                    break
        if not currency and "$" in line:
            currency = global_currency or "HKD"
        for amount in numeric_amounts(line):
            candidates.append((amount, currency, line.strip()))
    return candidates


def nearby_priority_amount(text: str) -> tuple[float | None, str, str]:
    lines = [line.strip() for line in text.splitlines()]
    global_currency = infer_currency_from_context(text)
    priority_terms = ["total incl", "total", "mastercard", "visa", "paid", "kpay", "pay", "fare"]
    candidates: list[tuple[float, str, str]] = []
    for index, line in enumerate(lines):
        lower = line.lower()
        if not any(term in lower for term in priority_terms) or "balance" in lower:
            continue
        same_line_amounts = numeric_amounts(line)
        if same_line_amounts:
            window = line
            amounts = same_line_amounts
        else:
            if "total" in lower:
                start = index
                end = min(len(lines), index + 3)
            elif any(term in lower for term in ["kpay", "pay", "visa", "mastercard", "paid", "fare"]):
                start = max(0, index - 3)
                end = min(len(lines), index + 2)
            else:
                start = max(0, index - 1)
                end = min(len(lines), index + 2)
            window = " ".join(
                piece for piece in lines[start:end] if piece and not piece.lower().startswith("ocr candidate")
            )
            amounts = numeric_amounts(window)
        currency = normalize_currency(parse_amount_and_currency(window, "")[1]) or global_currency
        if not currency and "$" in window:
            currency = "HKD"
        for amount in amounts:
            candidates.append((amount, currency, window))
    if not candidates:
        return None, "", ""
    selected = max(candidates, key=lambda item: item[0])
    return selected


def preferred_amount_from_text(text: str) -> tuple[float | None, str, str]:
    nearby_amount, nearby_currency, nearby_line = nearby_priority_amount(text)
    if nearby_amount is not None:
        return nearby_amount, nearby_currency, nearby_line
    candidates = amount_candidates_with_context(text)
    if not candidates:
        return None, "", ""
    priority_terms = [
        "total incl",
        "total",
        "mastercard",
        "visa",
        "paid",
        "kpay",
        "fare",
    ]
    priority = [
        item
        for item in candidates
        if any(term in item[2].lower() for term in priority_terms)
        and "balance" not in item[2].lower()
    ]
    currency_candidates = [item for item in candidates if item[1]]
    selected = max(priority or currency_candidates or candidates, key=lambda item: item[0])
    if not selected[1]:
        for candidate in currency_candidates:
            if abs(candidate[0] - selected[0]) < 0.005:
                selected = candidate
                break
    return selected


def first_nonempty_line(text: str) -> str:
    for line in text.splitlines():
        cleaned = re.sub(r"\s+", " ", line).strip(" <>|")
        if cleaned:
            return cleaned
    return ""


def heuristic_records_from_ocr(extracted_text: str, original_name: str, reason: str = "") -> list[dict]:
    text = extracted_text or ""
    lower = text.lower()
    dates = all_text_dates(text)
    receipt_date = max(dates).isoformat() if dates else ""
    time_match = re.search(r"\b\d{1,2}:\d{2}(?::\d{2})?\b", text)
    time_value = time_match.group(0)[:5] if time_match else ""
    amount, currency, amount_line = preferred_amount_from_text(text)
    if not currency:
        currency = infer_currency_from_context(f"{original_name} {text[:3000]}")
    amount_local = amount_with_currency(amount, currency) if amount is not None else ""
    person_match = re.search(r"Guest Name\s*[:;|]?\s*([A-Za-z][A-Za-z .'-]{1,60})", text, flags=re.IGNORECASE)
    person = person_match.group(1).strip() if person_match else ""

    merchant = first_nonempty_line(text)
    if "crowne plaza" in lower:
        activities = "Accommodation at Crowne Plaza Copenhagen Towers"
        location = "Copenhagen, Denmark"
        address = "Orestads Boulevard 114-118, DK-2300 Copenhagen S"
        raw_type = "Accommondation"
    elif "hoopla" in lower:
        activities = "Meal at Hoopla Restaurant"
        location = "Hong Kong"
        address = ""
        raw_type = "Meal"
    else:
        activities = merchant or Path(original_name).stem
        location = ""
        address = ""
        raw_type = ""

    notes = "Fallback extraction from OCR text."
    if amount_line:
        notes += f" Amount source: {amount_line}."
    if reason:
        notes += f" LLM unavailable: {reason}"

    return [
        {
            "date": receipt_date,
            "time": time_value,
            "person": person,
            "type": classify_expense(raw_type, activities, original_name),
            "activities": activities,
            "location": location,
            "address": address,
            "amount_in_local_currency": amount_local,
            "amount_in_hkd": "",
            "confidence": "ocr_fallback",
            "notes": notes,
            "currency": currency,
            "_ocr_text": text[:12000],
        }
    ]


class BaseExtractor:
    def extract(self, path: Path, mime_type: str, original_name: str) -> list[dict]:
        raise NotImplementedError


class MockExtractor(BaseExtractor):
    def extract(self, path: Path, mime_type: str, original_name: str) -> list[dict]:
        raise ExtractionError(
            "No LLM provider is configured. Set LLM_PROVIDER=openai, minimax, anthropic, gemini, "
            "or openai_compatible. Use LLM_PROVIDER=fixture only for the included sample files."
        )


class FixtureExtractor(BaseExtractor):
    def extract(self, path: Path, mime_type: str, original_name: str) -> list[dict]:
        name = original_name.lower()
        fixtures = {
            "hotel receipt.jpg": [
                {
                    "date": "2025-07-29",
                    "time": "",
                    "person": "Qingwei Li",
                    "type": "Accommondation",
                    "activities": "Accommodation package at Crowne Plaza Copenhagen Towers",
                    "location": "Copenhagen, Denmark",
                    "address": "Orestads Boulevard 114-118, DK-2300 Copenhagen S",
                    "amount_in_local_currency": "8750.00 DKK",
                    "amount_in_hkd": "",
                    "confidence": "fixture",
                    "notes": "Arrival 2025-07-24; departure/payment 2025-07-29.",
                }
            ],
            "transaction record.png": [
                {
                    "date": "2025-04-23",
                    "time": "11:33",
                    "person": "",
                    "type": "Others",
                    "activities": "Online payment to WWW.AOM.ORG",
                    "location": "US",
                    "address": "",
                    "amount_in_local_currency": "275.00 USD",
                    "amount_in_hkd": "",
                    "confidence": "fixture",
                    "notes": "Screen also shows 361.02 at 1 USD = 1.3128 SGD; completed 2025-04-24.",
                }
            ],
            "cafe.png": [
                {
                    "date": "",
                    "time": "16:54:57",
                    "person": "",
                    "type": "Meal",
                    "activities": "Meal at Hopla Restaurant",
                    "location": "Mong Kok, Hong Kong",
                    "address": "Kowloon, Hong Kong",
                    "amount_in_local_currency": "200.00 HKD",
                    "amount_in_hkd": "200.00",
                    "confidence": "fixture",
                    "notes": "Receipt date is not visible in the screenshot.",
                }
            ],
            "taxi.png": [
                {
                    "date": "2023-05-01",
                    "time": "15:03",
                    "person": "",
                    "type": "Transportation",
                    "activities": "Taxi fare",
                    "location": "Hong Kong",
                    "address": "",
                    "amount_in_local_currency": "176.60 HKD",
                    "amount_in_hkd": "176.60",
                    "confidence": "fixture",
                    "notes": "End time 16:13; total distance 8.38 km; surcharge included.",
                }
            ],
            "flight ticket.pdf": [
                {
                    "date": "2025-04-29",
                    "time": "10:39",
                    "person": "Li Qingwei",
                    "type": "Flight",
                    "activities": "Flight Copenhagen to Singapore via Bangkok, Thai Airways TG951/TG403",
                    "location": "Copenhagen to Singapore",
                    "address": "",
                    "amount_in_local_currency": "4638.00 CNY",
                    "amount_in_hkd": "",
                    "confidence": "fixture",
                    "notes": "Trip.com receipt total paid by Alipay.",
                }
            ],
        }
        if name not in fixtures:
            allowed = ", ".join(sorted(fixtures))
            raise ExtractionError(
                "This server is running in fixture test mode, so it only accepts the built-in sample filenames: "
                f"{allowed}. Start the app with LLM_PROVIDER=minimax, openai, openai_compatible, anthropic, or gemini "
                f"to extract new uploads such as {original_name}."
            )
        return fixtures[name]


class OpenAICompatibleExtractor(BaseExtractor):
    def __init__(self, provider_name: str = "openai") -> None:
        self.provider_name = provider_name
        self.api_key = os.environ.get("LLM_API_KEY", "")
        self.model = os.environ.get("LLM_MODEL", "gpt-4o-mini")
        self.base_url = os.environ.get("LLM_BASE_URL", "https://api.openai.com/v1").rstrip("/")
        if not self.api_key:
            raise ExtractionError("LLM_API_KEY is required.")

    def extract(self, path: Path, mime_type: str, original_name: str) -> list[dict]:
        prompt = f"{USER_PROMPT}\n\nSource filename: {original_name}"
        content: object
        if mime_type.startswith("image/"):
            content = [
                {"type": "text", "text": prompt},
                {"type": "image_url", "image_url": {"url": f"data:{mime_type};base64,{base64_file(path)}"}},
            ]
        elif mime_type == "application/pdf":
            pdf_text = read_pdf_text(path)
            if not pdf_text:
                raise ExtractionError("PDF text could not be extracted. Install pypdf or use a provider that accepts PDF inputs directly.")
            content = f"{prompt}\n\nExtracted PDF text:\n{pdf_text[:50000]}"
        else:
            content = f"{prompt}\n\nUnsupported visual type {mime_type}; extract what you can from filename."

        payload = {
            "model": self.model,
            "temperature": 0,
            "messages": [
                {"role": "system", "content": SYSTEM_PROMPT},
                {"role": "user", "content": content},
            ],
        }
        data = post_json(
            f"{self.base_url}/chat/completions",
            payload,
            {"Authorization": f"Bearer {self.api_key}"},
        )
        text = data["choices"][0]["message"]["content"]
        return as_records(extract_json(text))


class MiniMaxExtractor(BaseExtractor):
    def __init__(self) -> None:
        self.api_key = os.environ.get("LLM_API_KEY", "")
        self.model = os.environ.get("LLM_MODEL", "MiniMax-M2.7")
        self.base_url = os.environ.get("LLM_BASE_URL", "https://api.minimax.io/anthropic").rstrip("/")
        self.temperature = float(os.environ.get("LLM_TEMPERATURE", "0.1"))
        if not self.api_key or self.api_key == "paste-your-minimax-api-key-here":
            raise ExtractionError("MiniMax is selected, but LLM_API_KEY is empty. Paste your MiniMax API key into receipt-reimbursement/.env.")

    def extract(self, path: Path, mime_type: str, original_name: str) -> list[dict]:
        prompt = f"{USER_PROMPT}\n\nSource filename: {original_name}"
        if mime_type.startswith("image/"):
            extracted_text = read_image_text(path)
            text_source = "Local OCR text"
        elif mime_type == "application/pdf":
            extracted_text = read_pdf_text(path)
            text_source = "Extracted PDF text"
        else:
            extracted_text = ""
            text_source = "Unsupported file type"

        if not extracted_text:
            raise ExtractionError(
                "MiniMax's OpenAI-compatible endpoint does not accept image/PDF bytes directly in this app. "
                f"No text could be extracted from {original_name}; try a clearer image/PDF or use a vision provider."
            )

        payload = {
            "model": self.model,
            "max_tokens": int(os.environ.get("LLM_MAX_TOKENS", "2500")),
            "temperature": self.temperature,
            "system": SYSTEM_PROMPT,
            "messages": [
                {"role": "user", "content": [{"type": "text", "text": f"{prompt}\n\n{text_source}:\n{extracted_text[:50000]}"}]},
            ],
        }
        try:
            data = post_json(
                f"{self.base_url}/v1/messages",
                payload,
                {"Authorization": f"Bearer {self.api_key}"},
                timeout=int(os.environ.get("LLM_TIMEOUT_SECONDS", DEFAULT_LLM_TIMEOUT_SECONDS)),
            )
        except ExtractionError as exc:
            if env_flag("OCR_FALLBACK_ON_LLM_ERROR", True):
                return heuristic_records_from_ocr(extracted_text, original_name, str(exc))
            raise
        text = "\n".join(
            part.get("text", "")
            for part in data.get("content", [])
            if isinstance(part, dict) and part.get("type") == "text"
        )
        text = strip_model_thinking(text)
        try:
            records = as_records(extract_json(text))
        except ExtractionError as exc:
            if env_flag("OCR_FALLBACK_ON_LLM_ERROR", True):
                return heuristic_records_from_ocr(extracted_text, original_name, f"LLM response could not be parsed: {exc}")
            raise
        if not records and env_flag("OCR_FALLBACK_ON_LLM_ERROR", True):
            return heuristic_records_from_ocr(extracted_text, original_name, "LLM returned no structured records.")
        for record in records:
            record["_ocr_text"] = extracted_text[:12000]
        return records


class AnthropicExtractor(BaseExtractor):
    def __init__(self) -> None:
        self.api_key = os.environ.get("LLM_API_KEY", "")
        self.model = os.environ.get("LLM_MODEL", "claude-3-5-sonnet-latest")
        self.base_url = os.environ.get("LLM_BASE_URL", "https://api.anthropic.com/v1").rstrip("/")
        if not self.api_key:
            raise ExtractionError("LLM_API_KEY is required.")

    def extract(self, path: Path, mime_type: str, original_name: str) -> list[dict]:
        prompt = f"{USER_PROMPT}\n\nSource filename: {original_name}"
        blocks: list[dict] = [{"type": "text", "text": prompt}]
        if mime_type.startswith("image/"):
            blocks.append(
                {
                    "type": "image",
                    "source": {"type": "base64", "media_type": mime_type, "data": base64_file(path)},
                }
            )
        elif mime_type == "application/pdf":
            pdf_text = read_pdf_text(path)
            if not pdf_text:
                raise ExtractionError("PDF text could not be extracted. Install pypdf before sending PDFs to this adapter.")
            blocks.append({"type": "text", "text": f"Extracted PDF text:\n{pdf_text[:50000]}"})
        else:
            blocks.append({"type": "text", "text": f"Unsupported visual type {mime_type}."})

        payload = {
            "model": self.model,
            "max_tokens": 2500,
            "temperature": 0,
            "system": SYSTEM_PROMPT,
            "messages": [{"role": "user", "content": blocks}],
        }
        data = post_json(
            f"{self.base_url}/messages",
            payload,
            {"x-api-key": self.api_key, "anthropic-version": "2023-06-01"},
        )
        text = "\n".join(part.get("text", "") for part in data.get("content", []) if part.get("type") == "text")
        return as_records(extract_json(text))


class GeminiExtractor(BaseExtractor):
    def __init__(self) -> None:
        self.api_key = os.environ.get("LLM_API_KEY", "")
        self.model = os.environ.get("LLM_MODEL", "gemini-1.5-pro")
        self.base_url = os.environ.get("LLM_BASE_URL", "https://generativelanguage.googleapis.com/v1beta").rstrip("/")
        if not self.api_key:
            raise ExtractionError("LLM_API_KEY is required.")

    def extract(self, path: Path, mime_type: str, original_name: str) -> list[dict]:
        prompt = f"{SYSTEM_PROMPT}\n\n{USER_PROMPT}\n\nSource filename: {original_name}"
        parts: list[dict] = [{"text": prompt}]
        if mime_type.startswith("image/") or mime_type == "application/pdf":
            parts.append({"inline_data": {"mime_type": mime_type, "data": base64_file(path)}})
        else:
            parts.append({"text": f"Unsupported visual type {mime_type}."})
        payload = {"contents": [{"role": "user", "parts": parts}], "generationConfig": {"temperature": 0}}
        data = post_json(f"{self.base_url}/models/{self.model}:generateContent?key={parse.quote(self.api_key)}", payload, {})
        candidates = data.get("candidates", [])
        text = ""
        if candidates:
            text = "\n".join(part.get("text", "") for part in candidates[0].get("content", {}).get("parts", []))
        return as_records(extract_json(text))


def get_extractor() -> BaseExtractor:
    provider = os.environ.get("LLM_PROVIDER", "mock").strip().lower()
    if provider == "fixture":
        return FixtureExtractor()
    if provider == "openai":
        return OpenAICompatibleExtractor("openai")
    if provider == "minimax":
        return MiniMaxExtractor()
    if provider == "openai_compatible":
        return OpenAICompatibleExtractor("openai_compatible")
    if provider == "anthropic":
        return AnthropicExtractor()
    if provider == "gemini":
        return GeminiExtractor()
    return MockExtractor()


def detect_mime(filename: str, content_type: str | None) -> str:
    if content_type and content_type != "application/octet-stream":
        return content_type.split(";")[0].strip()
    guessed, _ = mimetypes.guess_type(filename)
    return guessed or "application/octet-stream"


def parse_multipart(headers, body: bytes) -> tuple[list[dict], dict[str, str]]:
    content_type = headers.get("Content-Type")
    if not content_type or "multipart/form-data" not in content_type:
        raise ValueError("Expected multipart/form-data.")
    message = BytesParser(policy=email_policy).parsebytes(
        f"Content-Type: {content_type}\r\nMIME-Version: 1.0\r\n\r\n".encode("utf-8") + body
    )
    files: list[dict] = []
    fields: dict[str, str] = {}
    for part in message.iter_parts():
        disposition = part.get("Content-Disposition", "")
        if "form-data" not in disposition:
            continue
        filename = part.get_filename()
        if not filename:
            name = part.get_param("name", header="Content-Disposition")
            if name:
                fields[str(name)] = (part.get_payload(decode=True) or b"").decode("utf-8", errors="replace")
            continue
        payload = part.get_payload(decode=True) or b""
        if not payload:
            continue
        files.append(
            {
                "filename": filename,
                "content_type": part.get_content_type(),
                "content": payload,
            }
        )
    return files, fields


def process_upload_item(item: dict, client_id: str) -> list[dict]:
    original = safe_filename(item["filename"])
    stored_name = f"{dt.datetime.now().strftime('%Y%m%d-%H%M%S')}-{uuid.uuid4().hex[:8]}-{original}"
    stored_dir = UPLOAD_DIR / safe_client_id(client_id)
    stored_dir.mkdir(parents=True, exist_ok=True)
    stored_path = stored_dir / stored_name
    stored_path.write_bytes(item["content"])
    mime_type = detect_mime(original, item.get("content_type"))
    raw_records = get_extractor().extract(stored_path, mime_type, original)
    if not raw_records:
        raw_records = [{"notes": "Model returned no records."}]
    merged_record = merge_records_for_file(raw_records)
    return [normalize_record(merged_record, original, str(stored_path))]


def upload_job_workers() -> int:
    return max(1, int(os.environ.get("UPLOAD_JOB_WORKERS", "2")))


def upload_job_ttl_seconds() -> int:
    return max(60, int(os.environ.get("UPLOAD_JOB_TTL_SECONDS", "86400")))


def max_active_upload_jobs() -> int:
    return max(1, int(os.environ.get("MAX_ACTIVE_UPLOAD_JOBS", "8")))


def get_upload_job_executor() -> ThreadPoolExecutor:
    global UPLOAD_JOB_EXECUTOR
    with UPLOAD_JOB_EXECUTOR_LOCK:
        if UPLOAD_JOB_EXECUTOR is None:
            UPLOAD_JOB_EXECUTOR = ThreadPoolExecutor(max_workers=upload_job_workers())
        return UPLOAD_JOB_EXECUTOR


def cleanup_upload_jobs_unlocked() -> None:
    cutoff = time.time() - upload_job_ttl_seconds()
    for job_id, job in list(UPLOAD_JOBS.items()):
        if job.get("status") in TERMINAL_UPLOAD_STATUSES and float(job.get("updated_ts", 0)) < cutoff:
            del UPLOAD_JOBS[job_id]


def upload_job_snapshot(job: dict) -> dict:
    keys = [
        "id",
        "client_id",
        "request_id",
        "status",
        "created_at",
        "updated_at",
        "phase",
        "progress",
        "result",
        "error",
        "errors",
    ]
    return {key: job.get(key) for key in keys if key in job}


def update_upload_job(job_id: str, **updates) -> None:
    with UPLOAD_JOBS_LOCK:
        job = UPLOAD_JOBS.get(job_id)
        if not job:
            return
        job.update(updates)
        job["updated_at"] = now_local()
        job["updated_ts"] = time.time()


def upload_progress(total: int, processed: int = 0, added: int = 0, failed: int = 0) -> dict:
    return {
        "total": total,
        "processed": min(processed, total),
        "added": added,
        "failed": failed,
    }


def receipt_process_timeout_seconds(file_count: int) -> float:
    per_file = float(os.environ.get("RECEIPT_PROCESS_TIMEOUT_SECONDS", DEFAULT_RECEIPT_PROCESS_TIMEOUT_SECONDS))
    return max(15.0, per_file * max(1, file_count))


def process_upload_batch(
    local_files: list[dict],
    drive_file_ids: list[str],
    fields: dict[str, str],
    client_id: str,
    drive_session: dict | None,
    job_id: str | None = None,
) -> dict:
    total_requested = len(local_files) + len(drive_file_ids)
    files = list(local_files)
    errors: list[dict] = []
    processed_count = 0

    if job_id:
        update_upload_job(
            job_id,
            phase="Preparing files",
            progress=upload_progress(total_requested),
            errors=[],
        )

    if drive_file_ids:
        if not drive_session:
            raise ExtractionError("Please connect Google Drive before choosing Drive files.")
        for file_id in drive_file_ids:
            if job_id:
                update_upload_job(job_id, phase="Downloading selected Google Drive files")
            try:
                files.append(drive_upload_item(drive_session, file_id))
            except Exception as exc:
                errors.append({"file": file_id, "error": str(exc)})
                processed_count += 1
                if job_id:
                    update_upload_job(
                        job_id,
                        progress=upload_progress(total_requested, processed_count, 0, len(errors)),
                        errors=list(errors),
                    )

    save_mode = fields.get("save_mode", "batch").strip().lower()
    append_to_store = save_mode in {"append", "recent7"}
    batch_id = uuid.uuid4().hex
    normalized: list[dict] = []

    if files:
        if job_id:
            update_upload_job(job_id, phase="Extracting receipt data")
        max_workers = max(1, int(os.environ.get("MAX_PARALLEL_RECEIPTS", "3")))
        executor = ThreadPoolExecutor(max_workers=min(max_workers, len(files)))
        futures = {executor.submit(process_upload_item, item, client_id): safe_filename(item["filename"]) for item in files}
        pending = set(futures)
        deadline = time.monotonic() + receipt_process_timeout_seconds(len(files))
        try:
            while pending:
                remaining = deadline - time.monotonic()
                if remaining <= 0:
                    break
                done, pending = wait(pending, timeout=min(5.0, remaining), return_when=FIRST_COMPLETED)
                if not done and job_id:
                    update_upload_job(job_id, phase="Extracting receipt data")
                for future in done:
                    filename = futures[future]
                    try:
                        normalized.extend(future.result())
                    except Exception as exc:
                        print(f"Upload item {filename} failed: {exc}", flush=True)
                        errors.append({"file": filename, "error": str(exc)})
                    processed_count += 1
                    if job_id:
                        update_upload_job(
                            job_id,
                            progress=upload_progress(total_requested, processed_count, len(normalized), len(errors)),
                            errors=list(errors),
                        )
            for future in pending:
                filename = futures[future]
                future.cancel()
                errors.append(
                    {
                        "file": filename,
                        "error": f"Receipt processing timed out after {int(receipt_process_timeout_seconds(len(files)))} seconds.",
                    }
                )
                processed_count += 1
                if job_id:
                    update_upload_job(
                        job_id,
                        progress=upload_progress(total_requested, processed_count, len(normalized), len(errors)),
                        errors=list(errors),
                    )
        finally:
            executor.shutdown(wait=False, cancel_futures=True)

    workbook_url = ""
    if normalized:
        if job_id:
            update_upload_job(job_id, phase="Saving workbook")
        if drive_session:
            drive_file = save_records_to_user_drive(drive_session, normalized, save_mode, client_id)
            workbook_url = str(drive_file.get("webViewLink") or drive_file.get("webContentLink") or "")
        elif append_to_store:
            all_records = append_records(normalized, client_id)
            if save_mode == "recent7":
                workbook_path = CLIENTS_DIR / safe_client_id(client_id) / "recent7" / "reimbursements.xlsx"
                write_workbook(filter_recent_records(all_records, 7), workbook_path)
                workbook_url = f"/download.xlsx?client_id={parse.quote(client_id)}&range=recent7"
            else:
                workbook_url = f"/download.xlsx?client_id={parse.quote(client_id)}"
        else:
            workbook_path = batch_workbook_path(client_id, batch_id)
            write_workbook(normalized, workbook_path)
            workbook_url = f"/download.xlsx?client_id={parse.quote(client_id)}&batch_id={parse.quote(batch_id)}"

    if errors and not normalized:
        return {
            "ok": False,
            "error": errors[0]["error"],
            "errors": errors,
            "added": 0,
            "batch_id": "",
            "workbook": "",
            "save_mode": save_mode if save_mode in {"append", "recent7", "batch"} else "batch",
            "saved_to_google_drive": False,
            "provider": os.environ.get("LLM_PROVIDER", "mock"),
        }

    return {
        "ok": True,
        "added": len(normalized),
        "errors": errors,
        "batch_id": "" if append_to_store or drive_session else batch_id,
        "workbook": workbook_url,
        "save_mode": save_mode if save_mode in {"append", "recent7", "batch"} else "batch",
        "saved_to_google_drive": bool(drive_session and workbook_url),
        "provider": os.environ.get("LLM_PROVIDER", "mock"),
    }


def run_upload_job(
    job_id: str,
    local_files: list[dict],
    drive_file_ids: list[str],
    fields: dict[str, str],
    client_id: str,
    drive_session: dict | None,
) -> None:
    print(f"Upload job {job_id} started with {len(local_files) + len(drive_file_ids)} file(s).", flush=True)
    update_upload_job(job_id, status="running", phase="Starting")
    try:
        result = process_upload_batch(local_files, drive_file_ids, fields, client_id, drive_session, job_id)
        update_upload_job(
            job_id,
            status="completed",
            phase="Done",
            result=result,
            error="" if result.get("ok") else str(result.get("error", "Upload failed.")),
            errors=result.get("errors", []),
        )
        print(f"Upload job {job_id} completed: added {result.get('added', 0)} row(s).", flush=True)
    except Exception as exc:
        traceback.print_exc()
        update_upload_job(job_id, status="failed", phase="Failed", error=str(exc), errors=[{"file": "", "error": str(exc)}])
        print(f"Upload job {job_id} failed: {exc}", flush=True)


def start_upload_job(
    local_files: list[dict],
    drive_file_ids: list[str],
    fields: dict[str, str],
    client_id: str,
    drive_session: dict | None,
) -> str:
    request_id = safe_client_id(fields.get("upload_request_id")) if fields.get("upload_request_id", "").strip() else ""
    safe_id = safe_client_id(client_id)
    with UPLOAD_JOBS_LOCK:
        cleanup_upload_jobs_unlocked()
        if request_id:
            for existing_id, job in UPLOAD_JOBS.items():
                if job.get("client_id") == safe_id and job.get("request_id") == request_id:
                    return existing_id
        active_count = sum(1 for job in UPLOAD_JOBS.values() if job.get("status") not in TERMINAL_UPLOAD_STATUSES)
        if active_count >= max_active_upload_jobs():
            raise ExtractionError("The upload server is busy. Please try again in a moment.")
        job_id = uuid.uuid4().hex
        now = now_local()
        UPLOAD_JOBS[job_id] = {
            "id": job_id,
            "client_id": safe_id,
            "request_id": request_id,
            "status": "queued",
            "created_at": now,
            "updated_at": now,
            "created_ts": time.time(),
            "updated_ts": time.time(),
            "phase": "Queued",
            "progress": upload_progress(len(local_files) + len(drive_file_ids)),
            "result": None,
            "error": "",
            "errors": [],
        }
    get_upload_job_executor().submit(run_upload_job, job_id, local_files, drive_file_ids, fields, client_id, drive_session)
    return job_id


def page_html() -> bytes:
    provider_warning = ""
    if os.environ.get("LLM_PROVIDER", "mock").strip().lower() == "fixture":
        provider_warning = (
            '<div class="mode-warning">Fixture mode only accepts the five built-in sample filenames. '
            "Use a real LLM provider for new uploads.</div>"
        )
    html_doc = f"""<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>Receipt Reimbursement</title>
  <style>
    :root {{
      color-scheme: light;
      --ink: #17202a;
      --muted: #667085;
      --line: #d7dde5;
      --panel: #f6f8fb;
      --accent: #0f766e;
      --accent-dark: #115e59;
      --warn: #9a3412;
      --bg: #ffffff;
    }}
    * {{ box-sizing: border-box; }}
    body {{
      margin: 0;
      font-family: ui-sans-serif, system-ui, -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif;
      color: var(--ink);
      background: var(--bg);
    }}
    header {{
      border-bottom: 1px solid var(--line);
      padding: 20px 28px 16px;
    }}
    h1 {{
      font-size: 22px;
      line-height: 1.2;
      margin: 0;
      font-weight: 650;
      letter-spacing: 0;
    }}
    main {{
      padding: 24px 28px 40px;
      max-width: 760px;
      margin: 0 auto;
    }}
    .panel {{
      border: 1px solid var(--line);
      border-radius: 8px;
      background: #fff;
      overflow: hidden;
    }}
    .panel-body {{ padding: 18px; }}
    .dropzone {{
      border: 1.5px dashed #9aa8b7;
      border-radius: 8px;
      background: var(--panel);
      min-height: 190px;
      padding: 20px;
      display: grid;
      place-items: center;
      text-align: center;
      cursor: pointer;
      transition: border-color .15s, background .15s;
    }}
    .dropzone.dragover {{
      border-color: var(--accent);
      background: #edf7f5;
    }}
    input[type=file] {{ display: none; }}
    .upload-title {{
      font-weight: 650;
      margin-bottom: 8px;
    }}
    .hint {{
      color: var(--muted);
      font-size: 13px;
      line-height: 1.5;
    }}
    .actions {{
      margin-top: 16px;
      display: flex;
      gap: 10px;
      flex-wrap: wrap;
    }}
    button, .button-link {{
      border: 1px solid var(--accent);
      background: var(--accent);
      color: white;
      border-radius: 7px;
      padding: 9px 13px;
      font-size: 14px;
      cursor: pointer;
      text-decoration: none;
      display: inline-flex;
      align-items: center;
      min-height: 38px;
    }}
    button.secondary, .button-link.secondary {{
      background: white;
      color: var(--accent-dark);
    }}
    button:disabled {{
      opacity: .55;
      cursor: not-allowed;
    }}
    .message {{
      margin-top: 14px;
      font-size: 13px;
      line-height: 1.5;
      color: var(--muted);
      white-space: pre-wrap;
    }}
    .message.error {{ color: var(--warn); }}
    .message.success {{ color: var(--accent-dark); }}
    .file-list {{
      margin-top: 16px;
      border: 1px solid var(--line);
      border-radius: 8px;
      overflow: hidden;
      display: none;
    }}
    .file-list.visible {{ display: block; }}
    .file-list-header {{
      padding: 10px 12px;
      background: var(--panel);
      border-bottom: 1px solid var(--line);
      font-size: 13px;
      color: var(--muted);
      display: flex;
      justify-content: space-between;
      gap: 12px;
    }}
    .file-list ul {{
      list-style: none;
      padding: 0;
      margin: 0;
      max-height: 240px;
      overflow: auto;
    }}
    .file-list li {{
      padding: 9px 12px;
      border-bottom: 1px solid var(--line);
      font-size: 13px;
      overflow-wrap: anywhere;
    }}
    .file-list li:last-child {{ border-bottom: 0; }}
    .drive-panel {{
      margin-top: 16px;
      border: 1px solid var(--line);
      border-radius: 8px;
      padding: 12px;
      background: #fff;
      display: grid;
      gap: 10px;
    }}
    .drive-panel.hidden {{ display: none; }}
    .drive-row {{
      display: flex;
      gap: 10px;
      align-items: center;
      flex-wrap: wrap;
    }}
    .drive-status {{
      color: var(--muted);
      font-size: 13px;
      line-height: 1.45;
      flex: 1;
      min-width: 210px;
    }}
    .drive-search {{
      display: flex;
      gap: 8px;
      flex-wrap: wrap;
    }}
    .drive-search input {{
      border: 1px solid var(--line);
      border-radius: 7px;
      min-height: 38px;
      padding: 8px 10px;
      font-size: 14px;
      flex: 1;
      min-width: 180px;
    }}
    .drive-browser.hidden {{ display: none; }}
    .drive-list {{
      list-style: none;
      padding: 0;
      margin: 0;
      border: 1px solid var(--line);
      border-radius: 8px;
      overflow: hidden;
      max-height: 260px;
      overflow-y: auto;
    }}
    .drive-list li {{
      display: flex;
      justify-content: space-between;
      gap: 10px;
      align-items: center;
      padding: 9px 10px;
      border-bottom: 1px solid var(--line);
      font-size: 13px;
    }}
    .drive-list li:last-child {{ border-bottom: 0; }}
    .drive-name {{ overflow-wrap: anywhere; }}
    .mode-select {{
      margin-top: 16px;
      border: 1px solid var(--line);
      border-radius: 8px;
      padding: 12px;
      display: grid;
      gap: 8px;
      background: #fff;
    }}
    .mode-select label {{
      display: flex;
      gap: 9px;
      align-items: flex-start;
      font-size: 13px;
      line-height: 1.4;
      color: var(--ink);
    }}
    .mode-select input {{ margin-top: 2px; }}
    .mode-select span {{ color: var(--muted); }}
    .download-panel {{
      margin-top: 16px;
      display: flex;
      gap: 10px;
      flex-wrap: wrap;
    }}
    .mode-warning {{
      margin-top: 14px;
      border: 1px solid #fed7aa;
      background: #fff7ed;
      color: #9a3412;
      border-radius: 7px;
      padding: 10px 12px;
      font-size: 13px;
      line-height: 1.45;
    }}
    @media (max-width: 900px) {{
      header {{
        padding: 18px;
      }}
      main {{
        padding: 18px;
      }}
    }}
  </style>
</head>
<body>
  <header>
    <h1>Receipt Reimbursement</h1>
  </header>
  <main>
    <section class="panel">
      <div class="panel-body">
        <form id="uploadForm">
          <label id="dropzone" class="dropzone">
            <span>
              <span class="upload-title">Upload receipt files</span><br>
              <span class="hint">Drag files here, or choose files/folder below.</span>
            </span>
          </label>
          <input id="fileInput" name="files" type="file" accept="image/*,.pdf" multiple>
          <input id="folderInput" name="folderFiles" type="file" accept="image/*,.pdf" multiple webkitdirectory directory>
          <div class="actions">
            <button id="chooseFilesButton" type="button" class="secondary">Choose Files</button>
            <button id="chooseFolderButton" type="button" class="secondary">Choose Folder</button>
            <button id="uploadButton" type="submit">Upload</button>
          </div>
          <div id="fileList" class="file-list">
            <div class="file-list-header">
              <span id="fileCount">No files selected</span>
              <span id="fileTotalSize"></span>
            </div>
            <ul id="fileListItems"></ul>
          </div>
          <div id="drivePanel" class="drive-panel hidden">
            <div class="drive-row">
              <div id="driveStatus" class="drive-status">Google Drive is not connected.</div>
              <button id="connectDriveButton" type="button" class="secondary">Connect Google Drive</button>
              <button id="disconnectDriveButton" type="button" class="secondary">Disconnect</button>
            </div>
            <div id="driveBrowser" class="drive-browser hidden">
              <div class="drive-search">
                <input id="driveSearchInput" type="search" placeholder="Search receipt files in Google Drive">
                <button id="driveSearchButton" type="button" class="secondary">Search Drive</button>
              </div>
              <ul id="driveResults" class="drive-list"></ul>
              <div id="driveSelected" class="file-list">
                <div class="file-list-header">
                  <span id="driveSelectedCount">No Drive files selected</span>
                  <span></span>
                </div>
                <ul id="driveSelectedItems"></ul>
              </div>
            </div>
          </div>
          <div class="mode-select">
            <label>
              <input type="radio" name="saveMode" value="batch" checked>
              <span><strong>Excel for this upload only</strong><br>Download only the receipts selected in this upload.</span>
            </label>
            <label>
              <input type="radio" name="saveMode" value="recent7">
              <span><strong>Append to my last 7-day Excel</strong><br>Download recent entries from your saved workbook.</span>
            </label>
            <label>
              <input type="radio" name="saveMode" value="append">
              <span><strong>Append to my long-term Excel</strong><br>Add this upload to your saved reimbursement workbook.</span>
            </label>
          </div>
          <div class="download-panel">
            <a id="downloadLink" class="button-link secondary" href="/download.xlsx">Download Excel</a>
          </div>
          {provider_warning}
          <div id="message" class="message"></div>
        </form>
      </div>
    </section>
  </main>
  <script>
    const params = new URLSearchParams(location.search);
    const token = params.get("token") || localStorage.getItem("receiptAppToken") || "";
    if (token) localStorage.setItem("receiptAppToken", token);
    let clientId = params.get("client_id") || localStorage.getItem("receiptClientId") || "";
    if (!clientId) {{
      clientId = crypto.randomUUID ? crypto.randomUUID() : `${{Date.now()}}-${{Math.random().toString(16).slice(2)}}`;
    }}
    localStorage.setItem("receiptClientId", clientId);

    const fileInput = document.getElementById("fileInput");
    const folderInput = document.getElementById("folderInput");
    const dropzone = document.getElementById("dropzone");
    const message = document.getElementById("message");
    const uploadButton = document.getElementById("uploadButton");
    const chooseFilesButton = document.getElementById("chooseFilesButton");
    const chooseFolderButton = document.getElementById("chooseFolderButton");
    const downloadLink = document.getElementById("downloadLink");
    const fileList = document.getElementById("fileList");
    const fileListItems = document.getElementById("fileListItems");
    const fileCount = document.getElementById("fileCount");
    const fileTotalSize = document.getElementById("fileTotalSize");
    const drivePanel = document.getElementById("drivePanel");
    const driveStatus = document.getElementById("driveStatus");
    const connectDriveButton = document.getElementById("connectDriveButton");
    const disconnectDriveButton = document.getElementById("disconnectDriveButton");
    const driveBrowser = document.getElementById("driveBrowser");
    const driveSearchInput = document.getElementById("driveSearchInput");
    const driveSearchButton = document.getElementById("driveSearchButton");
    const driveResults = document.getElementById("driveResults");
    const driveSelected = document.getElementById("driveSelected");
    const driveSelectedCount = document.getElementById("driveSelectedCount");
    const driveSelectedItems = document.getElementById("driveSelectedItems");
    let selectedFiles = [];
    let selectedDriveFiles = new Map();
    let latestBatchId = "";
    let latestWorkbookUrl = "";

    function headers() {{
      const result = {{"X-Client-Id": clientId}};
      if (token) result["X-App-Token"] = token;
      return result;
    }}

    function updateDownloadLink() {{
      if (latestWorkbookUrl) {{
        downloadLink.href = latestWorkbookUrl;
        downloadLink.textContent = latestWorkbookUrl.startsWith("http") ? "Open Excel in Google Drive" : "Download Excel";
        return;
      }}
      downloadLink.textContent = "Download Excel";
      const query = new URLSearchParams({{ client_id: clientId }});
      const mode = document.querySelector('input[name="saveMode"]:checked').value;
      if (mode === "recent7") {{
        query.set("range", "recent7");
      }} else if (mode === "batch" && latestBatchId) {{
        query.set("batch_id", latestBatchId);
      }}
      if (token) query.set("token", token);
      downloadLink.href = "/download.xlsx?" + query.toString();
    }}
    updateDownloadLink();
    document.querySelectorAll('input[name="saveMode"]').forEach(input => input.addEventListener("change", () => {{
      latestWorkbookUrl = "";
      updateDownloadLink();
    }}));

    function setMessage(text, mode = "") {{
      message.textContent = text;
      message.className = mode ? `message ${{mode}}` : "message";
    }}

    function sleep(ms) {{
      return new Promise(resolve => setTimeout(resolve, ms));
    }}

    async function readJsonResponse(res) {{
      const text = await res.text();
      let data;
      try {{ data = JSON.parse(text); }} catch {{
        const compact = text.trim().startsWith("<") ? "The server returned an HTML error page. Check the Render logs for the upload job." : text;
        throw new Error(compact);
      }}
      if (!res.ok) throw new Error(data.error || text);
      return data;
    }}

    async function pollUploadJob(jobId) {{
      let pollFailures = 0;
      while (true) {{
        let data;
        try {{
          const res = await fetch(`/api/upload/jobs/${{encodeURIComponent(jobId)}}`, {{ headers: headers() }});
          data = await readJsonResponse(res);
          pollFailures = 0;
        }} catch (error) {{
          pollFailures += 1;
          if (pollFailures >= 5) {{
            throw new Error(`Could not check upload progress: ${{error.message || String(error)}}`);
          }}
          setMessage(`Still processing. Reconnecting to upload status... (${{pollFailures}}/5)`);
          await sleep(2500);
          continue;
        }}
        const progress = data.progress || {{}};
        const total = progress.total || totalSelectedCount();
        const processed = progress.processed || 0;
        const phase = data.phase ? ` ${{data.phase}}.` : "";
        setMessage(`Processing ${{processed}}/${{total}} file(s)...${{phase}}`);
        if (data.status === "completed") {{
          const result = data.result || {{}};
          if (!result.ok) throw new Error(result.error || "Upload failed.");
          return result;
        }}
        if (data.status === "failed") throw new Error(data.error || "Upload failed.");
        await sleep(1800);
      }}
    }}

    function uploadRequestId() {{
      return crypto.randomUUID ? crypto.randomUUID() : `${{Date.now()}}-${{Math.random().toString(16).slice(2)}}`;
    }}

    function buildUploadForm(requestId) {{
      const form = new FormData();
      for (const file of selectedFiles) form.append("files", file, file.webkitRelativePath || file.name);
      form.append("drive_file_ids", JSON.stringify(Array.from(selectedDriveFiles.keys())));
      form.append("save_mode", document.querySelector('input[name="saveMode"]:checked').value);
      form.append("upload_request_id", requestId);
      return form;
    }}

    async function startUploadWithRetry(requestId) {{
      let lastError;
      for (let attempt = 1; attempt <= 3; attempt += 1) {{
        try {{
          const res = await fetch("/api/upload", {{
            method: "POST",
            headers: headers(),
            body: buildUploadForm(requestId),
            cache: "no-store",
          }});
          return await readJsonResponse(res);
        }} catch (error) {{
          lastError = error;
          if (attempt === 3) break;
          setMessage(`Upload connection interrupted. Retrying... (${{attempt}}/3)`);
          await sleep(1500 * attempt);
        }}
      }}
      throw new Error(lastError?.message || "Upload failed before the server returned a job id.");
    }}

    function formatBytes(bytes) {{
      if (!bytes) return "0 B";
      const units = ["B", "KB", "MB", "GB"];
      const index = Math.min(Math.floor(Math.log(bytes) / Math.log(1024)), units.length - 1);
      return `${{(bytes / Math.pow(1024, index)).toFixed(index ? 1 : 0)}} ${{units[index]}}`;
    }}

    function acceptedFile(file) {{
      const name = file.name.toLowerCase();
      return file.type.startsWith("image/") || name.endsWith(".pdf");
    }}

    function totalSelectedCount() {{
      return selectedFiles.length + selectedDriveFiles.size;
    }}

    function renderSelectedDriveFiles() {{
      driveSelectedItems.innerHTML = "";
      driveSelectedCount.textContent = `${{selectedDriveFiles.size}} Drive file(s) selected`;
      driveSelected.classList.toggle("visible", selectedDriveFiles.size > 0);
      for (const file of selectedDriveFiles.values()) {{
        const item = document.createElement("li");
        item.textContent = file.name;
        driveSelectedItems.appendChild(item);
      }}
    }}

    function updateReadyMessage() {{
      const total = totalSelectedCount();
      setMessage(total ? `${{total}} file(s) ready. Click Upload to process.` : "No supported files selected.");
    }}

    function setSelectedFiles(files) {{
      selectedFiles = Array.from(files).filter(acceptedFile);
      fileListItems.innerHTML = "";
      const totalSize = selectedFiles.reduce((sum, file) => sum + file.size, 0);
      fileCount.textContent = `${{selectedFiles.length}} file(s) selected`;
      fileTotalSize.textContent = formatBytes(totalSize);
      fileList.classList.toggle("visible", selectedFiles.length > 0);
      for (const file of selectedFiles) {{
        const item = document.createElement("li");
        item.textContent = file.webkitRelativePath || file.name;
        fileListItems.appendChild(item);
      }}
      updateReadyMessage();
    }}

    chooseFilesButton.addEventListener("click", () => fileInput.click());
    chooseFolderButton.addEventListener("click", () => folderInput.click());

    dropzone.addEventListener("dragover", event => {{
      event.preventDefault();
      dropzone.classList.add("dragover");
    }});
    dropzone.addEventListener("dragleave", () => dropzone.classList.remove("dragover"));
    dropzone.addEventListener("drop", event => {{
      event.preventDefault();
      dropzone.classList.remove("dragover");
      setSelectedFiles(event.dataTransfer.files);
    }});
    fileInput.addEventListener("change", () => setSelectedFiles(fileInput.files));
    folderInput.addEventListener("change", () => setSelectedFiles(folderInput.files));

    function authReturnTo() {{
      const url = new URL(location.href);
      if (token && !url.searchParams.has("token")) url.searchParams.set("token", token);
      return url.pathname + url.search;
    }}

    async function refreshDriveStatus() {{
      try {{
        const res = await fetch("/api/google/status", {{ headers: headers() }});
        if (!res.ok) return;
        const data = await res.json();
        drivePanel.classList.toggle("hidden", !data.drive_enabled);
        if (!data.drive_enabled) return;
        if (!data.oauth_configured) {{
          driveStatus.textContent = "Google Drive login is not configured on this server.";
          connectDriveButton.disabled = true;
          disconnectDriveButton.style.display = "none";
          driveBrowser.classList.add("hidden");
          return;
        }}
        connectDriveButton.disabled = false;
        if (data.connected) {{
          driveStatus.textContent = `Connected to Google Drive as ${{data.email || data.name || "Google user"}}. Excel files save to "${{data.folder}}".`;
          connectDriveButton.style.display = "none";
          disconnectDriveButton.style.display = "inline-flex";
          driveBrowser.classList.remove("hidden");
          await loadDriveFiles();
        }} else {{
          driveStatus.textContent = "Connect Google Drive to choose receipt files from Drive and save Excel there.";
          connectDriveButton.style.display = "inline-flex";
          disconnectDriveButton.style.display = "none";
          driveBrowser.classList.add("hidden");
        }}
      }} catch {{
        drivePanel.classList.add("hidden");
      }}
    }}

    async function loadDriveFiles() {{
      driveResults.innerHTML = "";
      const query = new URLSearchParams();
      if (driveSearchInput.value.trim()) query.set("q", driveSearchInput.value.trim());
      const res = await fetch(`/api/drive/files?${{query.toString()}}`, {{ headers: headers() }});
      const data = await res.json();
      if (!res.ok) throw new Error(data.error || "Could not load Google Drive files.");
      for (const file of data.files || []) {{
        const item = document.createElement("li");
        const name = document.createElement("span");
        name.className = "drive-name";
        name.textContent = file.name;
        const button = document.createElement("button");
        button.type = "button";
        button.className = "secondary";
        button.textContent = selectedDriveFiles.has(file.id) ? "Selected" : "Select";
        button.disabled = selectedDriveFiles.has(file.id);
        button.addEventListener("click", () => {{
          selectedDriveFiles.set(file.id, file);
          renderSelectedDriveFiles();
          updateReadyMessage();
          button.textContent = "Selected";
          button.disabled = true;
        }});
        item.appendChild(name);
        item.appendChild(button);
        driveResults.appendChild(item);
      }}
      if (!driveResults.children.length) {{
        const item = document.createElement("li");
        item.textContent = "No image/PDF receipt files found.";
        driveResults.appendChild(item);
      }}
    }}

    connectDriveButton.addEventListener("click", () => {{
      const query = new URLSearchParams({{ return_to: authReturnTo() }});
      if (token) query.set("token", token);
      location.href = "/auth/google/start?" + query.toString();
    }});
    disconnectDriveButton.addEventListener("click", () => {{
      location.href = "/auth/google/logout" + (token ? `?token=${{encodeURIComponent(token)}}` : "");
    }});
    driveSearchButton.addEventListener("click", async () => {{
      try {{
        await loadDriveFiles();
      }} catch (error) {{
        setMessage(error.message || String(error), "error");
      }}
    }});
    driveSearchInput.addEventListener("keydown", async event => {{
      if (event.key === "Enter") {{
        event.preventDefault();
        driveSearchButton.click();
      }}
    }});
    renderSelectedDriveFiles();
    refreshDriveStatus();

    document.getElementById("uploadForm").addEventListener("submit", async event => {{
      event.preventDefault();
      if (!totalSelectedCount()) {{
        setMessage("Select at least one file.", "error");
        return;
      }}
      uploadButton.disabled = true;
      setMessage(`Processing ${{totalSelectedCount()}} file(s)...`);
      try {{
        const started = await startUploadWithRetry(uploadRequestId());
        const data = started.job_id ? await pollUploadJob(started.job_id) : started;
        if (!data.ok) throw new Error(data.error || "Upload failed.");
        latestBatchId = data.batch_id || "";
        latestWorkbookUrl = data.workbook || "";
        updateDownloadLink();
        const uploadErrors = data.errors || [];
        const failed = uploadErrors.length ? ` ${{uploadErrors.length}} file(s) failed.` : "";
        const errorDetails = uploadErrors.length
          ? ` Details: ${{uploadErrors.slice(0, 3).map(item => `${{item.file || "file"}}: ${{item.error || "failed"}}`).join(" | ")}}`
          : "";
        const downloadText = data.saved_to_google_drive ? " Excel was saved to your Google Drive." : (latestBatchId ? " Download Excel now contains this upload only." : " Download Excel contains your saved workbook.");
        setMessage(`Added ${{data.added}} row(s).${{failed}}${{errorDetails}}${{downloadText}}`, uploadErrors.length ? "error" : "success");
        fileInput.value = "";
        folderInput.value = "";
        selectedFiles = [];
        selectedDriveFiles = new Map();
        renderSelectedDriveFiles();
        fileList.classList.remove("visible");
        fileListItems.innerHTML = "";
      }} catch (error) {{
        setMessage(error.message || String(error), "error");
      }} finally {{
        uploadButton.disabled = false;
      }}
    }});
  </script>
</body>
</html>"""
    return html_doc.encode("utf-8")


class ReceiptHandler(BaseHTTPRequestHandler):
    server_version = "ReceiptReimbursement/0.1"

    def log_message(self, fmt: str, *args) -> None:
        sys.stderr.write("[%s] %s\n" % (self.log_date_time_string(), fmt % args))

    def public_base_url(self) -> str:
        configured = os.environ.get("PUBLIC_BASE_URL", "").strip().rstrip("/")
        if configured:
            return configured
        proto = self.headers.get("X-Forwarded-Proto", "http").split(",")[0].strip() or "http"
        host = self.headers.get("Host", f"127.0.0.1:{os.environ.get('PORT', '8000')}")
        return f"{proto}://{host}"

    def redirect_uri(self) -> str:
        return self.public_base_url() + "/auth/google/callback"

    def send_redirect(self, location: str, cookies: list[str] | None = None) -> None:
        self.send_response(302)
        self.send_header("Location", location)
        for cookie in cookies or []:
            self.send_header("Set-Cookie", cookie)
        self.end_headers()

    def cookie_value(self, name: str) -> str:
        cookies = self.headers.get("Cookie", "")
        for part in cookies.split(";"):
            if "=" not in part:
                continue
            key, value = part.split("=", 1)
            if key.strip() == name:
                return parse.unquote(value.strip())
        return ""

    def google_session_cookie(self, session_id: str) -> str:
        secure = "; Secure" if self.public_base_url().startswith("https://") else ""
        max_age = int(os.environ.get("GOOGLE_SESSION_MAX_AGE_SECONDS", "2592000"))
        return f"{GOOGLE_USER_SESSION_COOKIE}={parse.quote(session_id)}; Path=/; HttpOnly; SameSite=Lax; Max-Age={max_age}{secure}"

    def clear_google_session_cookie(self) -> str:
        return f"{GOOGLE_USER_SESSION_COOKIE}=; Path=/; HttpOnly; SameSite=Lax; Max-Age=0"

    def google_session(self) -> tuple[str, dict] | tuple[None, None]:
        session_id = safe_client_id(self.cookie_value(GOOGLE_USER_SESSION_COOKIE))
        if not session_id or session_id == "default":
            return None, None
        with GOOGLE_OAUTH_LOCK:
            sessions = load_google_user_sessions()
            session = sessions.get(session_id)
            if not isinstance(session, dict):
                return None, None
            try:
                changed = refresh_google_user_session(session)
            except ExtractionError:
                sessions.pop(session_id, None)
                save_google_user_sessions(sessions)
                return None, None
            if changed:
                sessions[session_id] = session
                save_google_user_sessions(sessions)
            return session_id, session

    def require_google_session(self) -> dict:
        _, session = self.google_session()
        if not session:
            raise ExtractionError("Please connect Google Drive first.")
        return session

    def safe_return_to(self, value: str) -> str:
        if not value:
            return "/hsuhk-receipt-report-page"
        parsed = parse.urlparse(value)
        if parsed.scheme or parsed.netloc:
            return "/hsuhk-receipt-report-page"
        return value if value.startswith("/") else "/hsuhk-receipt-report-page"

    def handle_google_start(self) -> None:
        if not google_oauth_configured():
            self.send_text(
                503,
                "Google OAuth is not configured. Set GOOGLE_OAUTH_CLIENT_ID and GOOGLE_OAUTH_CLIENT_SECRET in the cloud host.",
            )
            return
        query = parse.parse_qs(parse.urlparse(self.path).query)
        return_to = self.safe_return_to(query.get("return_to", ["/hsuhk-receipt-report-page"])[0])
        state = secrets.token_urlsafe(32)
        with GOOGLE_OAUTH_LOCK:
            GOOGLE_OAUTH_STATES[state] = {"return_to": return_to, "created_at": str(unix_now())}
        params = {
            "client_id": os.environ.get("GOOGLE_OAUTH_CLIENT_ID", "").strip(),
            "redirect_uri": self.redirect_uri(),
            "response_type": "code",
            "scope": " ".join(google_oauth_scopes()),
            "access_type": "offline",
            "prompt": "consent",
            "state": state,
        }
        self.send_redirect("https://accounts.google.com/o/oauth2/v2/auth?" + parse.urlencode(params))

    def handle_google_callback(self) -> None:
        query = parse.parse_qs(parse.urlparse(self.path).query)
        state = query.get("state", [""])[0]
        code = query.get("code", [""])[0]
        if not state or not code:
            self.send_text(400, "Google OAuth callback is missing code/state.")
            return
        with GOOGLE_OAUTH_LOCK:
            state_data = GOOGLE_OAUTH_STATES.pop(state, None)
        if not state_data:
            self.send_text(400, "Google OAuth state expired. Please connect Google Drive again.")
            return
        if unix_now() - int(state_data.get("created_at") or "0") > 600:
            self.send_text(400, "Google OAuth state expired. Please connect Google Drive again.")
            return
        token_data = google_token_request(
            {
                "client_id": os.environ.get("GOOGLE_OAUTH_CLIENT_ID", "").strip(),
                "client_secret": os.environ.get("GOOGLE_OAUTH_CLIENT_SECRET", "").strip(),
                "code": code,
                "grant_type": "authorization_code",
                "redirect_uri": self.redirect_uri(),
            }
        )
        access_token = token_data.get("access_token")
        if not access_token:
            self.send_text(400, "Google OAuth did not return an access token.")
            return
        user = google_userinfo(str(access_token))
        session_id = secrets.token_urlsafe(32)
        session = {
            "access_token": str(access_token),
            "refresh_token": str(token_data.get("refresh_token") or ""),
            "expires_at": unix_now() + int(token_data.get("expires_in") or 3600),
            "email": str(user.get("email") or ""),
            "name": str(user.get("name") or ""),
            "created_at": now_local(),
        }
        with GOOGLE_OAUTH_LOCK:
            sessions = load_google_user_sessions()
            sessions[session_id] = session
            save_google_user_sessions(sessions)
        self.send_redirect(self.safe_return_to(state_data.get("return_to", "")), [self.google_session_cookie(session_id)])

    def handle_google_logout(self) -> None:
        session_id = safe_client_id(self.cookie_value(GOOGLE_USER_SESSION_COOKIE))
        if session_id and session_id != "default":
            with GOOGLE_OAUTH_LOCK:
                sessions = load_google_user_sessions()
                sessions.pop(session_id, None)
                save_google_user_sessions(sessions)
        self.send_redirect("/hsuhk-receipt-report-page", [self.clear_google_session_cookie()])

    def auth_ok(self) -> bool:
        expected = os.environ.get("APP_TOKEN", "")
        if not expected:
            return True
        query = parse.parse_qs(parse.urlparse(self.path).query)
        supplied = self.headers.get("X-App-Token") or (query.get("token", [""])[0])
        return secrets.compare_digest(supplied, expected)

    def client_id(self) -> str:
        query = parse.parse_qs(parse.urlparse(self.path).query)
        supplied = self.headers.get("X-Client-Id") or (query.get("client_id", [""])[0])
        return safe_client_id(supplied or "default")

    def send_json(self, status: int, payload: dict) -> None:
        body = json.dumps(payload, ensure_ascii=False).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def send_text(self, status: int, text: str) -> None:
        body = text.encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "text/plain; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def send_empty(self, status: int, content_type: str = "text/plain; charset=utf-8") -> None:
        self.send_response(status)
        self.send_header("Content-Type", content_type)
        self.send_header("Content-Length", "0")
        self.end_headers()

    def do_HEAD(self) -> None:
        ensure_dirs()
        path = parse.unquote(parse.urlparse(self.path).path)
        if path == "/health":
            self.send_empty(200, "application/json; charset=utf-8")
            return
        if path in PAGE_PATHS:
            self.send_empty(200, "text/html; charset=utf-8")
            return
        self.send_empty(404)

    def do_GET(self) -> None:
        ensure_dirs()
        path = parse.unquote(parse.urlparse(self.path).path)
        if path == "/health":
            self.send_json(200, {"ok": True, "provider": os.environ.get("LLM_PROVIDER", "mock")})
            return
        if path == "/auth/google/callback":
            try:
                self.handle_google_callback()
            except ExtractionError as exc:
                self.send_text(422, str(exc))
            return
        if not self.auth_ok():
            self.send_text(401, "Unauthorized. Open the page with ?token=YOUR_APP_TOKEN.")
            return
        if path == "/auth/google/start":
            try:
                self.handle_google_start()
            except ExtractionError as exc:
                self.send_text(422, str(exc))
            return
        if path == "/auth/google/logout":
            self.handle_google_logout()
            return
        if path in PAGE_PATHS:
            body = page_html()
            self.send_response(200)
            self.send_header("Content-Type", "text/html; charset=utf-8")
            self.send_header("Content-Length", str(len(body)))
            self.end_headers()
            self.wfile.write(body)
            return
        if path == "/api/google/status":
            _, session = self.google_session()
            self.send_json(
                200,
                {
                    "drive_enabled": user_google_drive_enabled(),
                    "oauth_configured": google_oauth_configured(),
                    "connected": bool(session),
                    "email": session.get("email", "") if session else "",
                    "name": session.get("name", "") if session else "",
                    "folder": google_drive_folder_name(),
                },
            )
            return
        if path == "/api/drive/files":
            try:
                session = self.require_google_session()
                query = parse.parse_qs(parse.urlparse(self.path).query)
                search = query.get("q", [""])[0]
                self.send_json(200, {"files": list_drive_receipt_files(session, search)})
            except ExtractionError as exc:
                self.send_json(422, {"error": str(exc)})
            return
        if path.startswith("/api/upload/jobs/"):
            job_id = safe_client_id(path.rsplit("/", 1)[-1])
            with UPLOAD_JOBS_LOCK:
                cleanup_upload_jobs_unlocked()
                job = UPLOAD_JOBS.get(job_id)
                snapshot = upload_job_snapshot(job) if job else None
            if not snapshot:
                self.send_json(404, {"error": "Upload job not found."})
                return
            if safe_client_id(snapshot.get("client_id")) != self.client_id():
                self.send_json(404, {"error": "Upload job not found."})
                return
            self.send_json(200, snapshot)
            return
        if path == "/api/records":
            client_id = self.client_id()
            self.send_json(200, {"records": read_records(client_id), "workbook": f"/download.xlsx?client_id={parse.quote(client_id)}"})
            return
        if path == "/download.xlsx":
            client_id = self.client_id()
            query = parse.parse_qs(parse.urlparse(self.path).query)
            batch_id = safe_client_id(query.get("batch_id", [""])[0])
            range_filter = query.get("range", [""])[0]
            if range_filter == "recent7":
                workbook_path = CLIENTS_DIR / safe_client_id(client_id) / "recent7" / "reimbursements.xlsx"
                records = filter_recent_records(read_records(client_id), 7)
                write_workbook(records, workbook_path)
            elif batch_id and batch_id != "default":
                workbook_path = batch_workbook_path(client_id, batch_id)
                if not workbook_path.exists():
                    self.send_text(404, "Batch workbook not found.")
                    return
                migrate_generated_workbook_columns(workbook_path)
            else:
                _, _, workbook_path = ensure_client_store(client_id)
                records = read_records(client_id)
                write_workbook(records, workbook_path)
            body = workbook_path.read_bytes()
            self.send_response(200)
            self.send_header("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            self.send_header("Content-Disposition", 'attachment; filename="reimbursements.xlsx"')
            self.send_header("Content-Length", str(len(body)))
            self.end_headers()
            self.wfile.write(body)
            return
        self.send_text(404, "Not found")

    def do_POST(self) -> None:
        ensure_dirs()
        if not self.auth_ok():
            self.send_json(401, {"error": "Unauthorized"})
            return
        path = parse.unquote(parse.urlparse(self.path).path)
        if path != "/api/upload":
            self.send_json(404, {"error": "Not found"})
            return
        try:
            length = int(self.headers.get("Content-Length", "0"))
            max_mb = float(os.environ.get("MAX_UPLOAD_MB", "20"))
            if length > max_mb * 1024 * 1024:
                self.send_json(413, {"error": f"Upload is larger than MAX_UPLOAD_MB={max_mb}."})
                return
            body = self.rfile.read(length)
            files, fields = parse_multipart(self.headers, body)
            drive_file_ids: list[str] = []
            if fields.get("drive_file_ids", "").strip():
                try:
                    parsed_ids = json.loads(fields["drive_file_ids"])
                    if isinstance(parsed_ids, list):
                        drive_file_ids = [str(item) for item in parsed_ids if str(item).strip()]
                except json.JSONDecodeError as exc:
                    raise ExtractionError("drive_file_ids must be a JSON array.") from exc
            max_files = max(1, int(os.environ.get("MAX_FILES_PER_UPLOAD", os.environ.get("MAX_PARALLEL_RECEIPTS", "10"))))
            if len(files) + len(drive_file_ids) > max_files:
                self.send_json(400, {"error": f"Upload at most {max_files} files at a time."})
                return
            if not files and not drive_file_ids:
                self.send_json(400, {"error": "No files were uploaded."})
                return

            client_id = self.client_id()
            drive_session = None
            if user_google_drive_enabled():
                drive_session = self.require_google_session()
            elif drive_file_ids:
                raise ExtractionError("Google Drive file selection is not enabled on this server.")
            job_id = start_upload_job(files, drive_file_ids, fields, client_id, drive_session)
            self.send_json(
                202,
                {
                    "ok": True,
                    "job_id": job_id,
                    "status": "queued",
                    "status_url": f"/api/upload/jobs/{parse.quote(job_id)}",
                    "total": len(files) + len(drive_file_ids),
                    "provider": os.environ.get("LLM_PROVIDER", "mock"),
                },
            )
        except ExtractionError as exc:
            self.send_json(422, {"error": str(exc)})
        except Exception as exc:
            traceback.print_exc()
            self.send_json(500, {"error": f"Server error: {exc}"})


def import_samples(paths: list[str], reset: bool = False) -> None:
    if reset:
        ensure_dirs_no_workbook()
        save_records([])
        write_workbook([])
    extractor = FixtureExtractor()
    normalized: list[dict] = []
    for raw_path in paths:
        path = Path(raw_path)
        if not path.exists():
            raise SystemExit(f"Sample file not found: {path}")
        original = safe_filename(path.name)
        stored_name = f"sample-{uuid.uuid4().hex[:8]}-{original}"
        stored_path = UPLOAD_DIR / stored_name
        ensure_dirs_no_workbook()
        stored_path.write_bytes(path.read_bytes())
        mime_type = detect_mime(original, None)
        raw_records = extractor.extract(stored_path, mime_type, original)
        normalized.extend(normalize_record(record, original, str(stored_path)) for record in raw_records)
    append_records(normalized)
    print(f"Imported {len(normalized)} sample row(s) into {WORKBOOK_PATH}")


def serve() -> None:
    ensure_dirs()
    host = os.environ.get("HOST", "127.0.0.1")
    port = int(os.environ.get("PORT", "8000"))
    httpd = ThreadingHTTPServer((host, port), ReceiptHandler)
    print(f"Receipt Reimbursement running on http://{host}:{port}")
    print(f"LLM_PROVIDER={os.environ.get('LLM_PROVIDER', 'mock')}")
    print(
        "Runtime timeouts: "
        f"OCR_TIMEOUT_SECONDS={os.environ.get('OCR_TIMEOUT_SECONDS', DEFAULT_OCR_TIMEOUT_SECONDS)} "
        f"OCR_MAX_LONG_EDGE={os.environ.get('OCR_MAX_LONG_EDGE', DEFAULT_OCR_MAX_LONG_EDGE)} "
        f"OCR_SMALL_MAX_LONG_EDGE={os.environ.get('OCR_SMALL_MAX_LONG_EDGE', DEFAULT_OCR_SMALL_MAX_LONG_EDGE)} "
        f"OCR_VARIANTS={os.environ.get('OCR_VARIANTS', DEFAULT_OCR_VARIANTS)} "
        f"OCR_PSMS={os.environ.get('OCR_PSMS', DEFAULT_OCR_PSMS)} "
        f"OCR_MAX_CANDIDATES={os.environ.get('OCR_MAX_CANDIDATES', DEFAULT_OCR_MAX_CANDIDATES)} "
        f"LLM_TIMEOUT_SECONDS={os.environ.get('LLM_TIMEOUT_SECONDS', DEFAULT_LLM_TIMEOUT_SECONDS)} "
        "RECEIPT_PROCESS_TIMEOUT_SECONDS="
        f"{os.environ.get('RECEIPT_PROCESS_TIMEOUT_SECONDS', DEFAULT_RECEIPT_PROCESS_TIMEOUT_SECONDS)}"
    )
    httpd.serve_forever()


if __name__ == "__main__":
    load_env_file()
    if len(sys.argv) > 1 and sys.argv[1] == "import-samples":
        reset_flag = "--reset" in sys.argv
        sample_paths = [arg for arg in sys.argv[2:] if arg != "--reset"]
        if not sample_paths:
            sample_paths = [
                str(APP_DIR.parent / "Flight Ticket.pdf"),
                str(APP_DIR.parent / "Transaction record.PNG"),
                str(APP_DIR.parent / "cafe.png"),
                str(APP_DIR.parent / "Hotel Receipt.jpg"),
                str(APP_DIR.parent / "taxi.png"),
            ]
        import_samples(sample_paths, reset_flag)
    else:
        serve()
