from __future__ import annotations

import io
import json
import base64
import html
import mimetypes
import os
import posixpath
import re
import socket
import shutil
import time
import urllib.error
import urllib.parse
import urllib.request
import uuid
import zipfile
from html.parser import HTMLParser
from pathlib import Path
from typing import Any

from dotenv import load_dotenv


load_dotenv(Path(__file__).resolve().parents[1] / ".env")


class MinerUConfigError(RuntimeError):
    pass


class MinerUExtractionError(RuntimeError):
    pass


def user_facing_mineru_error(error: Exception) -> str:
    detail = str(error)
    if isinstance(error, MinerUConfigError):
        return "PDF 识别服务 MinerU 尚未配置，请设置 MINERU_BASE_URL、MINERU_API_TOKEN 和 MINERU_PUBLIC_BASE_URL 后重试。"
    if "timed out" in detail or "TimeoutError" in detail:
        return (
            "PDF 识别服务 MinerU 连接超时。请确认 MINERU_BASE_URL 指向的 MinerU 服务已启动且本机可访问，"
            "并确认 MINERU_PUBLIC_BASE_URL 是 MinerU 服务可访问的后端地址。"
        )
    if "Failed to download MinerU result" in detail:
        return "PDF 识别结果下载失败，请检查 MinerU 返回的结果地址是否可访问。"
    if "Timed out waiting for MinerU task" in detail:
        return "PDF 识别任务等待超时，MinerU 已接收任务但长时间没有返回结果，请稍后重试或检查 MinerU 队列。"
    return f"PDF 识别服务 MinerU 失败：{detail}"


STAGING_DIR = Path(__file__).resolve().parent / "temp" / "mineru_public"
EXTRACT_DIR = Path(__file__).resolve().parent / "temp" / "mineru_extract"
TERMINAL_FAILED = {"failed", "failure", "error", "canceled", "cancelled"}
IMAGE_MD_RE = re.compile(r"!\[([^\]]*)\]\(([^)]+)\)")
DETAILS_RE = re.compile(r"\n?<details>\s*<summary>.*?</summary>.*?</details>\s*", re.S | re.I)
HTML_TABLE_RE = re.compile(r"<table\b.*?</table>", re.S | re.I)
HTML_TAG_RE = re.compile(r"<[^>]+>")


def _required_env(name: str) -> str:
    value = os.getenv(name, "").strip()
    if not value:
        raise MinerUConfigError(
            f"{name} is required for PDF extraction. Please set MINERU_BASE_URL, "
            "MINERU_API_TOKEN and MINERU_PUBLIC_BASE_URL."
        )
    return value


def _token() -> str:
    raw = os.getenv("MINERU_API_TOKEN", "").strip() or _required_env("MINERU_API_KEY")
    return raw[7:].strip() if raw.lower().startswith("bearer ") else raw


def _base_url() -> str:
    return _required_env("MINERU_BASE_URL").rstrip("/")


def _public_base_url() -> str:
    return _required_env("MINERU_PUBLIC_BASE_URL").rstrip("/")


def _model_version() -> str:
    return os.getenv("MINERU_MODEL_VERSION", "vlm").strip() or "vlm"


def _trust_env() -> bool:
    value = os.getenv("MINERU_TRUST_ENV", "1").strip().lower()
    return value not in {"0", "false", "no", "off"}


def _urlopen(request: urllib.request.Request | str, timeout: int):
    if _trust_env():
        return urllib.request.urlopen(request, timeout=timeout)
    opener = urllib.request.build_opener(urllib.request.ProxyHandler({}))
    return opener.open(request, timeout=timeout)


def _validate_public_base_url(url: str) -> None:
    parsed = urllib.parse.urlparse(url)
    if parsed.scheme not in {"http", "https"} or not parsed.netloc or not parsed.hostname:
        raise MinerUConfigError("MINERU_PUBLIC_BASE_URL must be a valid http(s) URL.")
    host = parsed.hostname.strip().lower()
    if host in {"localhost", "127.0.0.1", "::1"}:
        raise MinerUConfigError(
            "MINERU_PUBLIC_BASE_URL must be reachable by MinerU; localhost/127.0.0.1 cannot be used."
        )


def _probe_mineru_service(timeout: int) -> None:
    url = f"{_base_url()}/extract/task/__codex_startup_probe__"
    req = urllib.request.Request(url, method="GET", headers={"Authorization": f"Bearer {_token()}"})
    try:
        with _urlopen(req, timeout=timeout) as resp:
            resp.read(1)
    except urllib.error.HTTPError as exc:
        detail = exc.read().decode("utf-8", errors="replace")
        if exc.code in {400, 404, 405}:
            return
        if exc.code in {401, 403}:
            raise MinerUConfigError(f"MinerU authentication failed during startup check: HTTP {exc.code}") from exc
        raise MinerUExtractionError(f"MinerU startup check HTTP {exc.code}: {detail}") from exc
    except (TimeoutError, socket.timeout, urllib.error.URLError) as exc:
        raise MinerUExtractionError(f"MinerU startup check failed: {exc}") from exc


def validate_mineru_environment(*, timeout: int | None = None) -> dict[str, str]:
    """Validate the mandatory MinerU-only PDF extraction path before serving requests."""
    base_url = _base_url()
    _token()
    public_base_url = _public_base_url()
    _validate_public_base_url(public_base_url)
    _probe_mineru_service(timeout or int(os.getenv("MINERU_STARTUP_TIMEOUT_SECONDS", "10")))
    return {
        "base_url": base_url,
        "public_base_url": public_base_url,
        "model_version": _model_version(),
    }


def _json_request(method: str, url: str, payload: dict[str, Any] | None = None, timeout: int = 30) -> dict[str, Any]:
    data = None
    headers = {"Authorization": f"Bearer {_token()}"}
    if payload is not None:
        data = json.dumps(payload).encode("utf-8")
        headers["Content-Type"] = "application/json"
    req = urllib.request.Request(url, data=data, method=method, headers=headers)
    try:
        with _urlopen(req, timeout=timeout) as resp:
            return json.loads(resp.read().decode("utf-8"))
    except urllib.error.HTTPError as exc:
        detail = exc.read().decode("utf-8", errors="replace")
        raise MinerUExtractionError(f"MinerU HTTP {exc.code}: {detail}") from exc
    except urllib.error.URLError as exc:
        raise MinerUExtractionError(f"MinerU request failed: {exc}") from exc


def _download(url: str, timeout: int = 180) -> bytes:
    try:
        with _urlopen(url, timeout=timeout) as resp:
            return resp.read()
    except urllib.error.URLError as exc:
        raise MinerUExtractionError(f"Failed to download MinerU result: {exc}") from exc


def _stage_pdf(pdf_path: Path) -> tuple[str, Path, str]:
    STAGING_DIR.mkdir(parents=True, exist_ok=True)
    file_id = f"{uuid.uuid4().hex}.pdf"
    staged = STAGING_DIR / file_id
    shutil.copy2(pdf_path, staged)
    return file_id, staged, f"{_public_base_url()}/api/handwriting/mineru_files/{file_id}"


def _safe_zip_members(zip_bytes: bytes) -> list[tuple[zipfile.ZipInfo, str]]:
    members: list[tuple[zipfile.ZipInfo, str]] = []
    with zipfile.ZipFile(io.BytesIO(zip_bytes)) as zf:
        for info in zf.infolist():
            name = urllib.parse.unquote(info.filename).replace("\\", "/")
            if not name or name.endswith("/") or name.startswith("/"):
                continue
            normalized = posixpath.normpath(name)
            parts = [part for part in normalized.split("/") if part]
            if not parts or parts[0] == "__MACOSX" or any(part in {".", ".."} for part in parts):
                continue
            members.append((info, "/".join(parts)))
    return members


def _extract_zip(zip_bytes: bytes, out_dir: Path) -> None:
    out_dir.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(io.BytesIO(zip_bytes)) as zf:
        for info, normalized in _safe_zip_members(zip_bytes):
            target = out_dir / normalized
            target.parent.mkdir(parents=True, exist_ok=True)
            target.write_bytes(zf.read(info))


def _submit_task(file_url: str) -> str:
    body = _json_request(
        "POST",
        f"{_base_url()}/extract/task",
        {"url": file_url, "model_version": _model_version()},
    )
    data = body.get("data") or {}
    task_id = str(data.get("task_id") or data.get("id") or "").strip()
    if not task_id:
        raise MinerUExtractionError(f"MinerU did not return task_id: {body}")
    return task_id


def _poll_task(task_id: str) -> str:
    interval = float(os.getenv("MINERU_POLL_INTERVAL_SECONDS", "2"))
    timeout = float(os.getenv("MINERU_TIMEOUT_SECONDS", "600"))
    deadline = time.monotonic() + timeout
    while True:
        body = _json_request("GET", f"{_base_url()}/extract/task/{task_id}")
        data = body.get("data") or {}
        zip_url = str(data.get("full_zip_url") or "").strip()
        if zip_url:
            return zip_url
        status = str(data.get("status") or data.get("state") or "").strip().lower()
        if status in TERMINAL_FAILED:
            raise MinerUExtractionError(f"MinerU task failed: task_id={task_id}, status={status}")
        if time.monotonic() >= deadline:
            raise MinerUExtractionError(f"Timed out waiting for MinerU task: {task_id}")
        time.sleep(interval)


def _inline_local_image_source(source: str, base_dir: Path | None) -> str:
    source = (source or "").strip()
    if not source or source.startswith("data:image/") or re.match(r"^https?://", source, re.I):
        return source
    if base_dir is None:
        return source

    parsed = urllib.parse.urlparse(source)
    if parsed.scheme and parsed.scheme != "file":
        return source
    raw_path = urllib.parse.unquote(parsed.path if parsed.scheme == "file" else source)
    path = Path(raw_path)
    if not path.is_absolute():
        path = base_dir / path
    if not path.exists() or not path.is_file():
        return source

    mime = mimetypes.guess_type(path.name)[0] or "image/png"
    if not mime.startswith("image/"):
        mime = "image/png"
    encoded = base64.b64encode(path.read_bytes()).decode("ascii")
    return f"data:{mime};base64,{encoded}"


class _HTMLTableTextParser(HTMLParser):
    def __init__(self) -> None:
        super().__init__(convert_charrefs=True)
        self.rows: list[list[str]] = []
        self._row: list[str] | None = None
        self._cell: list[str] | None = None

    def handle_starttag(self, tag: str, attrs: list[tuple[str, str | None]]) -> None:
        tag = tag.lower()
        if tag == "tr":
            self._finish_row()
            self._row = []
        elif tag in {"td", "th"}:
            if self._row is None:
                self._row = []
            self._finish_cell()
            self._cell = []
        elif tag in {"br", "p", "div"} and self._cell is not None:
            self._cell.append(" ")

    def handle_endtag(self, tag: str) -> None:
        tag = tag.lower()
        if tag in {"td", "th"}:
            self._finish_cell()
        elif tag == "tr":
            self._finish_row()

    def handle_data(self, data: str) -> None:
        if self._cell is not None:
            self._cell.append(data)

    def close(self) -> None:
        super().close()
        self._finish_row()

    def _finish_cell(self) -> None:
        if self._cell is None:
            return
        if self._row is None:
            self._row = []
        text = re.sub(r"\s+", " ", "".join(self._cell)).strip()
        self._row.append(text)
        self._cell = None

    def _finish_row(self) -> None:
        self._finish_cell()
        if self._row is not None and any(cell.strip() for cell in self._row):
            self.rows.append(self._row)
        self._row = None


def _html_table_to_text(table_html: str) -> str:
    parser = _HTMLTableTextParser()
    try:
        parser.feed(table_html)
        parser.close()
    except Exception:
        text = HTML_TAG_RE.sub(" ", html.unescape(table_html))
        return re.sub(r"\s+", " ", text).strip()

    rows = parser.rows
    if not rows:
        text = HTML_TAG_RE.sub(" ", html.unescape(table_html))
        return re.sub(r"\s+", " ", text).strip()

    column_count = max(len(row) for row in rows)
    widths = [
        max((len(row[index]) for row in rows if index < len(row)), default=0)
        for index in range(column_count)
    ]
    lines = []
    for row in rows:
        padded = [
            row[index].ljust(widths[index]) if index < len(row) else "".ljust(widths[index])
            for index in range(column_count)
        ]
        lines.append("  ".join(padded).rstrip())
    return "\n".join(lines)


def sanitize_mineru_markdown(markdown: str, base_dir: Path | None = None) -> str:
    markdown = markdown.replace("\r\n", "\n").replace("\r", "\n")
    markdown = DETAILS_RE.sub("\n", markdown)
    markdown = HTML_TABLE_RE.sub(lambda m: f"\n\n{_html_table_to_text(m.group(0))}\n\n", markdown)
    markdown = IMAGE_MD_RE.sub(
        lambda m: f"![{m.group(1)}]({_inline_local_image_source(m.group(2), base_dir)})",
        markdown,
    )
    markdown = re.sub(r"\n{3,}", "\n\n", markdown)
    return markdown.strip() + "\n"


def _collect_markdown(extract_dir: Path) -> str:
    candidates = sorted(
        extract_dir.rglob("*.md"),
        key=lambda p: (0 if p.name == "full.md" else 1, str(p)),
    )
    if not candidates:
        raise MinerUExtractionError("MinerU result did not contain Markdown files.")
    parts = []
    for path in candidates:
        text = sanitize_mineru_markdown(path.read_text(encoding="utf-8", errors="ignore"), path.parent)
        if text.strip():
            parts.append(text.strip())
    return "\n\n".join(parts).strip() + "\n"


def extract_pdf_to_markdown(pdf_path: Path) -> dict[str, Any]:
    file_id, staged, public_url = _stage_pdf(pdf_path)
    out_dir = EXTRACT_DIR / file_id.removesuffix(".pdf")
    try:
        task_id = _submit_task(public_url)
        zip_url = _poll_task(task_id)
        zip_bytes = _download(zip_url)
        _extract_zip(zip_bytes, out_dir)
        markdown = _collect_markdown(out_dir)
        return {
            "markdown": markdown,
            "source": "mineru",
            "warnings": [],
            "metadata": {"task_id": task_id, "model_version": _model_version()},
        }
    finally:
        staged.unlink(missing_ok=True)
        shutil.rmtree(out_dir, ignore_errors=True)
