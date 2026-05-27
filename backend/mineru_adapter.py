from __future__ import annotations

import io
import json
import os
import posixpath
import re
import shutil
import time
import urllib.error
import urllib.parse
import urllib.request
import uuid
import zipfile
from pathlib import Path
from typing import Any


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


def sanitize_mineru_markdown(markdown: str) -> str:
    markdown = markdown.replace("\r\n", "\n").replace("\r", "\n")
    markdown = DETAILS_RE.sub("\n", markdown)
    markdown = IMAGE_MD_RE.sub(lambda m: f"[图片:{m.group(0)}]", markdown)
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
        text = sanitize_mineru_markdown(path.read_text(encoding="utf-8", errors="ignore"))
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
