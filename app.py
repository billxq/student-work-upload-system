from __future__ import annotations

import base64
import html
import json
import re
import secrets
import time
import urllib.parse
import zipfile
import xml.etree.ElementTree as ET
from dataclasses import dataclass
from http import cookies
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from typing import Dict, List, Optional


BASE_DIR = Path(__file__).resolve().parent
UPLOAD_ROOT = BASE_DIR / "uploads"
DATA_DIR = BASE_DIR / "data"
DATA_DIR.mkdir(exist_ok=True)

DEFAULT_PASSWORD = "111111"
SESSION_COOKIE = "student_upload_session"

XLSX_NS = {
    "a": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
}

INVALID_PATH_CHARS = re.compile(r'[\\/:*?"<>|]+')
INDEXED_FILE_RE = re.compile(r"^(?P<base>.+)_(?P<index>\d+)\.jpg$", re.IGNORECASE)
MULTIPART_BOUNDARY_RE = re.compile(r'boundary=(?:"([^"]+)"|([^;]+))')
CONTENT_DISPOSITION_PARAM_RE = re.compile(r';\s*([^=]+)="?([^";]*)"?')


@dataclass(frozen=True)
class Student:
    class_name: str
    name: str
    student_id: str


def find_workbook_path() -> Path:
    candidates = list(BASE_DIR.glob("*.xlsx"))
    if not candidates:
        raise FileNotFoundError("未找到 Excel 文件")
    for candidate in candidates:
        if "学生名单" in candidate.name:
            return candidate
    return candidates[0]


def parse_shared_strings(zf: zipfile.ZipFile) -> List[str]:
    if "xl/sharedStrings.xml" not in zf.namelist():
        return []

    root = ET.fromstring(zf.read("xl/sharedStrings.xml"))
    strings: List[str] = []
    for si in root.findall("a:si", XLSX_NS):
        text = "".join(t.text or "" for t in si.iterfind(".//a:t", XLSX_NS))
        strings.append(text)
    return strings


def cell_value(cell: ET.Element, shared_strings: List[str]) -> str:
    value_node = cell.find("a:v", XLSX_NS)
    if value_node is None or value_node.text is None:
        return ""
    raw = value_node.text
    if cell.attrib.get("t") == "s":
        try:
            return shared_strings[int(raw)]
        except (ValueError, IndexError):
            return raw
    return raw


def parse_workbook_students(path: Path) -> List[Student]:
    students: List[Student] = []
    with zipfile.ZipFile(path) as zf:
        workbook = ET.fromstring(zf.read("xl/workbook.xml"))
        rels = ET.fromstring(zf.read("xl/_rels/workbook.xml.rels"))
        rel_map = {
            rel.attrib["Id"]: rel.attrib["Target"]
            for rel in rels
            if rel.attrib.get("Target")
        }
        sheets = workbook.find("a:sheets", XLSX_NS)
        if sheets is None:
            return students

        first_sheet = next(iter(sheets), None)
        if first_sheet is None:
            return students

        sheet_target = rel_map[first_sheet.attrib["{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"]]
        if not sheet_target.startswith("worksheets/"):
            sheet_target = f"worksheets/{sheet_target}"

        shared_strings = parse_shared_strings(zf)
        sheet_root = ET.fromstring(zf.read(f"xl/{sheet_target}"))
        rows = sheet_root.findall(".//a:sheetData/a:row", XLSX_NS)
        if not rows:
            return students

        headers = [cell_value(cell, shared_strings).strip() for cell in rows[0].findall("a:c", XLSX_NS)]
        header_map = {name: idx for idx, name in enumerate(headers)}

        class_idx = header_map.get("班级")
        name_idx = header_map.get("学生姓名")
        id_idx = header_map.get("学生上海市学籍号")

        if class_idx is None or name_idx is None or id_idx is None:
            raise ValueError("Excel 表头必须包含：班级、学生姓名、学生上海市学籍号")

        for row in rows[1:]:
            values = [cell_value(cell, shared_strings).strip() for cell in row.findall("a:c", XLSX_NS)]
            if max(class_idx, name_idx, id_idx) >= len(values):
                continue
            class_name = values[class_idx]
            name = values[name_idx]
            student_id = values[id_idx]
            if not (class_name and name and student_id):
                continue
            students.append(Student(class_name=class_name, name=name, student_id=student_id))
    return students


def load_students() -> Dict[str, Student]:
    workbook_path = find_workbook_path()
    students = parse_workbook_students(workbook_path)
    return {student.student_id: student for student in students}


def find_student_by_login_code(students: Dict[str, Student], login_code: str) -> Optional[Student]:
    login_code = login_code.strip()
    if not login_code:
        return None

    if login_code.isdigit() and len(login_code) == 8:
        matches = [student for student in students.values() if student.student_id.endswith(login_code)]
        if len(matches) == 1:
            return matches[0]
        return None

    return students.get(login_code)


def sanitize_component(value: str) -> str:
    cleaned = INVALID_PATH_CHARS.sub("_", value.strip())
    cleaned = cleaned.replace("/", "_").replace("\\", "_")
    cleaned = cleaned.strip(" .")
    return cleaned or "未命名"


def ensure_directory(path: Path) -> None:
    path.mkdir(parents=True, exist_ok=True)


def load_json(path: Path, default):
    if not path.exists():
        return default
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except json.JSONDecodeError:
        return default


def save_json(path: Path, data) -> None:
    path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")


def append_upload_record(record: dict) -> None:
    path = DATA_DIR / "upload_records.json"
    records = load_json(path, [])
    records.append(record)
    save_json(path, records)


def parse_multipart_form(content_type: str, body: bytes) -> Dict[str, List[dict]]:
    match = MULTIPART_BOUNDARY_RE.search(content_type)
    if not match:
        return {}

    boundary = (match.group(1) or match.group(2) or "").encode("utf-8")
    if not boundary:
        return {}

    delimiter = b"--" + boundary
    fields: Dict[str, List[dict]] = {}
    for raw_part in body.split(delimiter):
        raw_part = raw_part.strip()
        if not raw_part or raw_part == b"--":
            continue
        if raw_part.startswith(b"--"):
            raw_part = raw_part[2:]
        raw_part = raw_part.lstrip(b"\r\n").rstrip(b"\r\n")
        if not raw_part:
            continue
        header_blob, separator, payload = raw_part.partition(b"\r\n\r\n")
        if not separator:
            continue

        headers = {}
        for line in header_blob.split(b"\r\n"):
            if b":" not in line:
                continue
            key, value = line.split(b":", 1)
            headers[key.strip().lower()] = value.strip()

        disposition = headers.get(b"content-disposition", b"").decode("utf-8", errors="ignore")
        if "form-data" not in disposition:
            continue

        name_match = re.search(r'name="([^"]+)"', disposition)
        if not name_match:
            continue
        name = name_match.group(1)
        filename_match = re.search(r'filename="([^"]*)"', disposition)
        filename = filename_match.group(1) if filename_match else None
        content_type_value = headers.get(b"content-type", b"").decode("utf-8", errors="ignore")
        if payload.endswith(b"\r\n"):
            payload = payload[:-2]

        fields.setdefault(name, []).append(
            {
                "filename": filename,
                "content_type": content_type_value,
                "data": payload,
            }
        )
    return fields


def get_cookie_value(cookie_header: Optional[str], name: str) -> Optional[str]:
    if not cookie_header:
        return None
    jar = cookies.SimpleCookie()
    jar.load(cookie_header)
    if name in jar:
        return jar[name].value
    return None


def school_header_data_uri() -> str:
    asset_path = BASE_DIR / "assets" / "school-header.png"
    if not asset_path.exists():
        return ""
    data = base64.b64encode(asset_path.read_bytes()).decode("ascii")
    return f"data:image/png;base64,{data}"


def favicon_data_uri() -> str:
    asset_path = BASE_DIR / "assets" / "icon.png"
    if not asset_path.exists():
        return ""
    data = base64.b64encode(asset_path.read_bytes()).decode("ascii")
    return f"data:image/png;base64,{data}"


def page_header(title: str, subtitle: str = "") -> str:
    banner = school_header_data_uri()
    banner_html = f'<img class="login-banner" src="{banner}" alt="上海师范大学附属青浦实验小学">' if banner else ""
    subtitle_html = f'<p class="login-subtitle">{html.escape(subtitle)}</p>' if subtitle else ""
    return f"""
      {banner_html}
      <h3 class="login-title">{html.escape(title)}</h3>
      {subtitle_html}
    """


def render_page(title: str, body: str) -> bytes:
    favicon = favicon_data_uri()
    favicon_link = f'<link rel="icon" type="image/png" href="{favicon}">' if favicon else ""
    html_doc = f"""<!doctype html>
<html lang="zh-CN">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1, viewport-fit=cover">
  <title>{html.escape(title)}</title>
  {favicon_link}
  <style>
    :root {{
      --bg1: #eff5ff;
      --bg2: #dce9ff;
      --card: rgba(255, 255, 255, 0.96);
      --text: #1d2748;
      --muted: #64748b;
      --primary: #2e67f8;
      --primary-dark: #1650f0;
      --border: #d7dfec;
    }}
    * {{ box-sizing: border-box; }}
    html, body {{ min-height: 100%; }}
    body {{
      margin: 0;
      font-family: "Microsoft YaHei", "PingFang SC", "Noto Sans SC", sans-serif;
      color: var(--text);
      min-height: 100vh;
      display: flex;
      align-items: center;
      justify-content: center;
      padding: clamp(12px, 3vw, 28px);
      background:
        radial-gradient(circle at top left, rgba(255, 255, 255, 0.78), transparent 30%),
        radial-gradient(circle at bottom right, rgba(120, 168, 255, 0.18), transparent 32%),
        linear-gradient(160deg, var(--bg1), var(--bg2));
    }}
    .shell {{
      width: min(100%, 560px);
    }}
    .card {{
      background: var(--card);
      border: 1px solid rgba(255, 255, 255, 0.92);
      border-radius: 24px;
      box-shadow: 0 20px 70px rgba(77, 102, 152, 0.16);
      backdrop-filter: blur(16px);
      padding: clamp(14px, 2.4vw, 22px);
    }}
    .login-banner {{
      width: min(100%, 430px);
      display: block;
      margin: 0 auto 12px;
      border-radius: 14px;
      box-shadow: 0 10px 28px rgba(74, 103, 160, 0.12);
    }}
    .login-title {{
      margin: 4px 0 8px;
      text-align: center;
      font-size: clamp(17px, 2.2vw, 25px);
      line-height: 1.1;
      letter-spacing: 0.01em;
      color: #18315a;
      font-weight: 900;
      max-width: 22ch;
      margin-left: auto;
      margin-right: auto;
    }}
    .login-subtitle {{
      margin: 0 auto 14px;
      text-align: center;
      color: #64748b;
      font-size: clamp(13px, 1.4vw, 15px);
      line-height: 1.6;
      max-width: 460px;
    }}
    .login-form {{
      width: min(100%, 430px);
      margin: 0 auto;
    }}
    label {{
      display: block;
      font-weight: 800;
      margin: 12px 0 6px;
      font-size: 14px;
      color: #0f172a;
    }}
    input, button {{ font: inherit; }}
    input[type="text"], input[type="password"], input[type="file"] {{
      width: 100%;
      padding: 12px 14px;
      border: 1px solid var(--border);
      border-radius: 12px;
      background: #fff;
      outline: none;
      color: #1f2937;
      font-size: 14px;
    }}
    input[type="text"]:focus, input[type="password"]:focus {{
      border-color: rgba(46, 103, 248, 0.58);
      box-shadow: 0 0 0 4px rgba(46, 103, 248, 0.1);
    }}
    input::placeholder {{
      color: #94a3b8;
    }}
    .login-form .actions {{
      margin-top: 16px;
      display: block;
    }}
    .login-form .btn {{
      width: 100%;
      min-height: 46px;
      border-radius: 12px;
      font-size: 17px;
      box-shadow: 0 14px 26px rgba(46, 103, 248, 0.28);
    }}
    .actions {{
      display: flex;
      gap: 12px;
      align-items: center;
      margin-top: 20px;
      flex-wrap: wrap;
    }}
    .btn {{
      display: inline-flex;
      align-items: center;
      justify-content: center;
      padding: 13px 18px;
      border: 0;
      border-radius: 14px;
      background: linear-gradient(135deg, var(--primary), var(--primary-dark));
      color: white;
      text-decoration: none;
      cursor: pointer;
      font-weight: 700;
      box-shadow: 0 12px 30px rgba(37,99,235,0.28);
    }}
    .btn.secondary {{
      background: rgba(15, 23, 42, 0.08);
      color: var(--text);
      box-shadow: none;
    }}
    .alert {{
      padding: 10px 12px;
      border-radius: 12px;
      margin-bottom: 12px;
      line-height: 1.6;
    }}
    .alert.error {{
      background: rgba(220, 38, 38, 0.09);
      border: 1px solid rgba(220, 38, 38, 0.2);
      color: #991b1b;
    }}
    .alert.success {{
      background: rgba(16, 185, 129, 0.11);
      border: 1px solid rgba(16, 185, 129, 0.25);
      color: #065f46;
    }}
    .meta-grid {{
      display: grid;
      gap: 12px;
      margin: 16px 0 0;
    }}
    .meta-item {{
      padding: 14px 16px;
      border-radius: 16px;
      border: 1px solid rgba(148, 163, 184, 0.22);
      background: rgba(248, 250, 252, 0.92);
    }}
    .meta-item strong {{
      display: block;
      margin-bottom: 4px;
    }}
    .hint {{
      margin-top: 10px;
      color: #64748b;
      font-size: 13px;
      line-height: 1.7;
      text-align: center;
    }}
    .file-list {{
      margin: 12px 0 0;
      padding-left: 18px;
    }}
    .file-list li {{
      margin: 6px 0;
      word-break: break-all;
    }}
    @media (max-width: 820px) {{
      body {{
        align-items: flex-start;
      }}
      .card {{
        border-radius: 18px;
        padding: 14px 12px 16px;
      }}
      .login-title {{
        font-size: clamp(18px, 6vw, 22px);
        max-width: 100%;
      }}
      .login-form, .login-banner {{
        width: 100%;
      }}
      .login-form .btn {{
        min-height: 44px;
        font-size: 16px;
      }}
      label {{
        font-size: 13px;
      }}
      input[type="text"], input[type="password"], input[type="file"] {{
        font-size: 14px;
      }}
    }}
  </style>
</head>
<body>
  <div class="shell">
    <main class="card">
      {body}
    </main>
  </div>
</body>
</html>"""
    return html_doc.encode("utf-8")


def login_form(error: str = "", student_id: str = "", name: str = "") -> bytes:
    error_block = f'<div class="alert error">{html.escape(error)}</div>' if error else ""
    body = f"""
      {page_header("上师大附小项目化作品提交系统", "请输入学籍号后 8 位、姓名和密码进入上传页面。")}
      {error_block}
      <form class="login-form" method="post" action="/login">
        <label for="student_id">学籍号后 8 位</label>
        <input id="student_id" name="student_id" type="text" value="{html.escape(student_id)}" autocomplete="off" inputmode="numeric" maxlength="8" placeholder="例如：20250101" required>

        <label for="name">姓名</label>
        <input id="name" name="name" type="text" value="{html.escape(name)}" autocomplete="off" placeholder="请输入姓名" required>

        <label for="password">密码</label>
        <input id="password" name="password" type="password" value="" placeholder="默认：111111" required>

        <div class="actions">
          <button class="btn" type="submit">登录</button>
        </div>
      </form>
      <div class="hint">如果登录失败，请检查学籍号后 8 位和姓名是否正确。</div>
    """
    return render_page("上师大附小项目化作品提交系统", body)


def upload_form(student: Student, message: str = "", error: str = "") -> bytes:
    msg_block = f'<div class="alert success">{html.escape(message)}</div>' if message else ""
    err_block = f'<div class="alert error">{html.escape(error)}</div>' if error else ""
    body = f"""
      {page_header("上传作品", f"当前登录学生：{student.class_name} 班 {student.name}（{student.student_id}）")}
      {msg_block}
      {err_block}
      <form class="login-form" method="post" action="/upload" enctype="multipart/form-data">
        <label for="work_type">作品类型</label>
        <input id="work_type" name="work_type" type="text" placeholder="例如：绘画作品 / 手工作品 / 语文作业" required>

        <label for="photos">上传作品图片</label>
        <input id="photos" name="photos" type="file" accept="image/*" multiple required>

        <div class="actions">
          <button class="btn" type="submit">上传</button>
          <a class="btn secondary" href="/logout">退出登录</a>
        </div>
      </form>
      <div class="hint">支持一次选择多张照片。系统会自动重命名为 <code>班级/姓名/作品类型_1.jpg</code> 的形式。</div>
    """
    return render_page("学生作品收集系统 - 上传", body)


def success_page(student: Student, upload_result: dict) -> bytes:
    file_items = "".join(
        f"<li>{html.escape(item)}</li>" for item in upload_result.get("saved_files", [])
    )
    body = f"""
      {page_header("上传成功", "作品已经保存完成，下面是本次上传的信息。")}
      <div class="alert success">上传成功</div>
      <div class="meta-grid">
        <div class="meta-item"><strong>学籍号</strong>{html.escape(student.student_id)}</div>
        <div class="meta-item"><strong>姓名</strong>{html.escape(student.name)}</div>
        <div class="meta-item"><strong>班级</strong>{html.escape(student.class_name)}</div>
        <div class="meta-item"><strong>作品类型</strong>{html.escape(upload_result.get('work_type', ''))}</div>
        <div class="meta-item"><strong>上传数量</strong>{html.escape(str(upload_result.get('count', 0)))}</div>
        <div class="meta-item"><strong>保存路径</strong>{html.escape(upload_result.get('target_dir', ''))}</div>
      </div>
      <h3 style="margin:22px 0 10px;">已保存文件</h3>
      <ul class="file-list">{file_items}</ul>
      <div class="actions">
        <a class="btn" href="/">继续上传</a>
        <a class="btn secondary" href="/logout">退出登录</a>
      </div>
    """
    return render_page("学生作品收集系统 - 上传成功", body)


class StudentUploadHandler(BaseHTTPRequestHandler):
    students = load_students()
    sessions: Dict[str, dict] = {}

    def log_message(self, format: str, *args) -> None:  # noqa: A003
        return

    def _send_html(self, body: bytes, status: int = 200, cookies_to_set: Optional[List[str]] = None) -> None:
        self.send_response(status)
        self.send_header("Content-Type", "text/html; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        if cookies_to_set:
            for cookie in cookies_to_set:
                self.send_header("Set-Cookie", cookie)
        self.end_headers()
        self.wfile.write(body)

    def _current_session(self) -> Optional[dict]:
        session_id = get_cookie_value(self.headers.get("Cookie"), SESSION_COOKIE)
        if not session_id:
            return None
        return self.sessions.get(session_id)

    def _require_student(self) -> Optional[Student]:
        session = self._current_session()
        if not session:
            return None
        student_id = session.get("student_id")
        if not student_id:
            return None
        return self.students.get(student_id)

    def do_GET(self) -> None:
        path = urllib.parse.urlparse(self.path).path
        if path == "/":
            student = self._require_student()
            if student:
                upload_result = self._current_session().get("last_upload") if self._current_session() else None
                if upload_result:
                    self._send_html(success_page(student, upload_result))
                else:
                    self._send_html(upload_form(student))
            else:
                self._send_html(login_form())
            return

        if path == "/success":
            student = self._require_student()
            session = self._current_session()
            if student and session and session.get("last_upload"):
                self._send_html(success_page(student, session["last_upload"]))
            else:
                self._send_html(login_form())
            return

        if path == "/logout":
            session_id = get_cookie_value(self.headers.get("Cookie"), SESSION_COOKIE)
            if session_id and session_id in self.sessions:
                self.sessions.pop(session_id, None)
            cookie = cookies.SimpleCookie()
            cookie[SESSION_COOKIE] = ""
            cookie[SESSION_COOKIE]["path"] = "/"
            cookie[SESSION_COOKIE]["max-age"] = 0
            self._send_html(login_form(), cookies_to_set=[cookie.output(header="").strip()])
            return

        self.send_error(404, "Not Found")

    def do_POST(self) -> None:
        path = urllib.parse.urlparse(self.path).path
        if path == "/login":
            self.handle_login()
            return
        if path == "/upload":
            self.handle_upload()
            return
        self.send_error(404, "Not Found")

    def handle_login(self) -> None:
        length = int(self.headers.get("Content-Length", "0"))
        raw = self.rfile.read(length).decode("utf-8", errors="ignore")
        fields = urllib.parse.parse_qs(raw, keep_blank_values=True)

        student_id = fields.get("student_id", [""])[0].strip()
        name = fields.get("name", [""])[0].strip()
        password = fields.get("password", [""])[0]

        if not student_id or not name or not password:
            self._send_html(login_form("请完整填写学籍号后 8 位、姓名和密码。", student_id, name))
            return

        student = find_student_by_login_code(self.students, student_id)
        if not student:
            self._send_html(login_form("未找到该学籍号后 8 位，请检查是否输入正确。", student_id, name))
            return

        if student.name != name:
            self._send_html(login_form("姓名与学籍号后 8 位不匹配，请重新确认。", student_id, name))
            return

        if password != DEFAULT_PASSWORD:
            self._send_html(login_form("密码错误，默认密码为 111111。", student_id, name))
            return

        session_id = secrets.token_urlsafe(24)
        self.sessions[session_id] = {
            "student_id": student.student_id,
            "login_at": time.time(),
        }
        cookie = cookies.SimpleCookie()
        cookie[SESSION_COOKIE] = session_id
        cookie[SESSION_COOKIE]["path"] = "/"
        cookie[SESSION_COOKIE]["httponly"] = True
        self.send_response(303)
        self.send_header("Location", "/")
        self.send_header("Set-Cookie", cookie.output(header="").strip())
        self.end_headers()

    def handle_upload(self) -> None:
        student = self._require_student()
        if not student:
            self._send_html(login_form("请先登录后再上传作品。"))
            return

        content_type = self.headers.get("Content-Type", "")
        if "multipart/form-data" not in content_type:
            self._send_html(upload_form(student, error="上传失败：表单格式不正确。"))
            return

        body = self.rfile.read(int(self.headers.get("Content-Length", "0")))
        form = parse_multipart_form(content_type, body)

        work_type_items = form.get("work_type", [])
        work_type = ""
        if work_type_items:
            work_type = work_type_items[0]["data"].decode("utf-8", errors="ignore").strip()
        photos = form.get("photos")

        if not work_type:
            self._send_html(upload_form(student, error="请先填写作品类型。"))
            return

        if photos is None:
            self._send_html(upload_form(student, error="请至少选择一张作品图片。"))
            return

        file_items = photos or []
        valid_files = []
        for item in file_items:
            filename = item.get("filename") or ""
            file_bytes = item.get("data", b"")
            if not filename or not file_bytes:
                continue
            valid_files.append((filename, file_bytes))

        if not valid_files:
            self._send_html(upload_form(student, error="没有读取到有效图片，请重新选择文件。"))
            return

        class_dir = sanitize_component(student.class_name)
        name_dir = sanitize_component(student.name)
        work_dir_name = sanitize_component(work_type)
        target_dir = UPLOAD_ROOT / class_dir / name_dir
        ensure_directory(target_dir)

        existing_indexes = []
        for existing in target_dir.glob(f"{work_dir_name}_*.jpg"):
            match = INDEXED_FILE_RE.match(existing.name)
            if match and match.group("base") == work_dir_name:
                existing_indexes.append(int(match.group("index")))
        next_index = max(existing_indexes, default=0) + 1

        saved_files = []
        for offset, (_, content) in enumerate(valid_files, start=0):
            index = next_index + offset
            filename = f"{work_dir_name}_{index}.jpg"
            target_path = target_dir / filename
            with target_path.open("wb") as f:
                f.write(content)
            saved_files.append(str(target_path.relative_to(BASE_DIR)).replace("\\", "/"))

        upload_result = {
            "work_type": work_type,
            "count": len(saved_files),
            "saved_files": saved_files,
            "target_dir": str(target_dir.relative_to(BASE_DIR)).replace("\\", "/"),
            "uploaded_at": time.strftime("%Y-%m-%d %H:%M:%S"),
        }

        append_upload_record(
            {
                "student_id": student.student_id,
                "name": student.name,
                "class_name": student.class_name,
                "work_type": work_type,
                "count": len(saved_files),
                "saved_files": saved_files,
                "uploaded_at": upload_result["uploaded_at"],
            }
        )

        session = self._current_session()
        if session is not None:
            session["last_upload"] = upload_result

        self.send_response(303)
        self.send_header("Location", "/success")
        self.end_headers()


def main() -> None:
    ensure_directory(UPLOAD_ROOT)
    server = ThreadingHTTPServer(("127.0.0.1", 8000), StudentUploadHandler)
    print("学生作品收集系统已启动：http://127.0.0.1:8000")
    print(f"读取的学生名单：{find_workbook_path().name}")
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\n服务器已停止")
    finally:
        server.server_close()


if __name__ == "__main__":
    main()
