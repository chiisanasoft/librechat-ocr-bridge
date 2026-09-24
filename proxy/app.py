"""OpenAI-compatible proxy in front of Ollama that makes PDF attachments usable.

LibreChat sends PDFs as OpenAI `file` content parts, which Ollama rejects with
"400 invalid message format". This proxy rewrites them before forwarding:

- OCR models (OCR_MODELS, e.g. glm-ocr): each PDF page is rendered to an image
  and OCR'd in its own request; the prompt is mapped to a supported task prompt
  ("Text Recognition:" / "Table Recognition:" / "Formula Recognition:").
- Vision models: PDF pages become `image_url` parts.
- Other models: the PDF's text layer is extracted and sent as a `text` part;
  scanned pages (mostly an image, often with a poor embedded OCR layer) are
  OCR'd with the OCR model instead.

OCR output is streamed through a repetition guard, since small OCR models tend
to loop (e.g. ": [table]: [table]...") after finishing a page. HTML tables in
OCR output are converted to Markdown tables and, for OCR models, also exported
as an .xlsx workbook served from /exports/ with a download link in the reply.

Everything else is passed through unchanged (including streaming).
"""

import base64
import json
import logging
import os
import re
import time
import uuid
from pathlib import Path
from urllib.parse import quote

import fitz  # PyMuPDF
import httpx
from starlette.applications import Starlette
from starlette.requests import Request
from starlette.responses import FileResponse, JSONResponse, Response, StreamingResponse
from starlette.routing import Route

import tables

UPSTREAM = os.environ.get("OLLAMA_UPSTREAM", "http://host.docker.internal:11434").rstrip("/")
PDF_DPI = int(os.environ.get("PDF_DPI", "200"))
PDF_MAX_PAGES = int(os.environ.get("PDF_MAX_PAGES", "20"))
OCR_MODELS = [m.strip() for m in os.environ.get("OCR_MODELS", "glm-ocr").split(",") if m.strip()]
OCR_MAX_TOKENS = int(os.environ.get("OCR_MAX_TOKENS", "8192"))
# Model used to OCR scanned PDFs for non-vision models ("" disables it)
OCR_FALLBACK_MODEL = os.environ.get("OCR_FALLBACK_MODEL", "glm-ocr:latest")
SCAN_IMAGE_COVERAGE = float(os.environ.get("SCAN_IMAGE_COVERAGE", "0.5"))
EXPORT_DIR = Path(os.environ.get("EXPORT_DIR", "/data/exports"))
EXPORT_TTL_HOURS = float(os.environ.get("EXPORT_TTL_HOURS", "72"))
# Browser-facing base URL of this proxy, used for the Excel download link
PUBLIC_BASE_URL = os.environ.get("PUBLIC_BASE_URL", "http://localhost:3082").rstrip("/")

OCR_TASK_PROMPTS = ("Text Recognition:", "Table Recognition:", "Formula Recognition:")

log = logging.getLogger("ollama-proxy")
logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")

client = httpx.AsyncClient(timeout=httpx.Timeout(None, connect=10.0))
_vision_cache: dict[str, bool] = {}


def base_model(model: str) -> str:
    return model.split(":", 1)[0]


def is_ocr_model(model: str) -> bool:
    return base_model(model) in OCR_MODELS


async def is_vision_model(model: str) -> bool:
    if model not in _vision_cache:
        try:
            r = await client.post(f"{UPSTREAM}/api/show", json={"model": model})
            _vision_cache[model] = "vision" in r.json().get("capabilities", [])
        except Exception:
            log.exception("capability lookup failed for %s", model)
            return False
    return _vision_cache[model]


# ---------------------------------------------------------------- PDF helpers

def decode_data_url(data_url: str) -> tuple[str, bytes]:
    header, _, payload = data_url.partition(",")
    mime = header.removeprefix("data:").split(";", 1)[0]
    return mime, base64.b64decode(payload)


def page_numbers(doc: fitz.Document, pages: list[int] | None) -> list[int]:
    if pages:
        selected = [p - 1 for p in pages if 1 <= p <= doc.page_count]
        if selected:
            return selected
    return list(range(min(doc.page_count, PDF_MAX_PAGES)))


def pdf_to_image_parts(pdf: bytes, pages: list[int] | None = None) -> list[dict]:
    with fitz.open(stream=pdf, filetype="pdf") as doc:
        parts = []
        for i in page_numbers(doc, pages):
            png = doc[i].get_pixmap(dpi=PDF_DPI).tobytes("png")
            url = "data:image/png;base64," + base64.b64encode(png).decode()
            parts.append({"type": "image_url", "image_url": {"url": url}})
        return parts


def is_scanned(page: fitz.Page) -> bool:
    """True when images cover most of the page, i.e. any text layer is a scanner's OCR."""
    area = abs(page.rect)
    covered = sum(abs(fitz.Rect(info["bbox"]) & page.rect) for info in page.get_image_info())
    return area > 0 and covered / area >= SCAN_IMAGE_COVERAGE


async def pdf_to_text(pdf: bytes, filename: str, pages_wanted: list[int] | None,
                      prompt: str, headers: dict) -> str:
    with fitz.open(stream=pdf, filetype="pdf") as doc:
        pages = []
        for i in page_numbers(doc, pages_wanted):
            text = doc[i].get_text().strip()
            if OCR_FALLBACK_MODEL and (not text or is_scanned(doc[i])):
                png = doc[i].get_pixmap(dpi=PDF_DPI).tobytes("png")
                img = {"type": "image_url",
                       "image_url": {"url": "data:image/png;base64," + base64.b64encode(png).decode()}}
                log.info("OCR fallback page=%d model=%s prompt=%r", i + 1, OCR_FALLBACK_MODEL, prompt)
                text = "".join([t async for t in ocr_page(OCR_FALLBACK_MODEL, img, prompt, headers, {})])
                text, _ = tables.convert_html_tables(text)
            pages.append(f"--- Page {i + 1} ---\n{text.strip()}")
    return f'File: "{filename}"\n\n' + "\n\n".join(pages)


def file_part_payload(part: dict) -> tuple[str, str, bytes] | None:
    """Returns (filename, mime, bytes) for an OpenAI `file` part with inline data."""
    if part.get("type") != "file":
        return None
    f = part.get("file") or {}
    data_url = f.get("file_data") or ""
    if not data_url.startswith("data:"):
        return None
    mime, data = decode_data_url(data_url)
    return f.get("filename") or "document", mime, data


# ---------------------------------------------------------------- rewriting

async def rewrite_file_parts(body: dict, headers: dict) -> None:
    """Replace `file` parts in all messages with parts Ollama understands."""
    vision = None
    for msg in body.get("messages", []):
        content = msg.get("content")
        if not isinstance(content, list):
            continue
        new_content = []
        for part in content:
            payload = file_part_payload(part)
            if payload is None:
                new_content.append(part)
                continue
            filename, mime, data = payload
            if mime == "application/pdf":
                if vision is None:
                    vision = await is_vision_model(body.get("model", ""))
                request_text = message_text(content)
                pages = requested_pages(request_text)
                if vision:
                    new_content.extend(pdf_to_image_parts(data, pages))
                else:
                    prompt = ocr_prompt(request_text)
                    text = await pdf_to_text(data, filename, pages, prompt, headers)
                    new_content.append({"type": "text", "text": text})
            elif mime.startswith("image/"):
                url = "data:" + mime + ";base64," + base64.b64encode(data).decode()
                new_content.append({"type": "image_url", "image_url": {"url": url}})
            else:
                text = data.decode("utf-8", errors="replace")
                new_content.append({"type": "text", "text": f'File: "{filename}"\n\n{text}'})
        msg["content"] = new_content


def message_text(content: list) -> str:
    return "\n".join(p.get("text", "") for p in content if p.get("type") == "text")


def requested_pages(text: str) -> list[int] | None:
    """Parses page hints such as "first page", "page 3", "2ページ目", "1〜3ページ"."""
    t = text.lower()
    if re.search(r"\bfirst page\b|最初のページ|1枚目", t):
        return [1]
    m = re.search(r"(\d+)\s*[-~〜～]\s*(\d+)\s*ページ|pages?\s*(\d+)\s*-\s*(\d+)", t)
    if m:
        a, b = (int(x) for x in (m.group(1) or m.group(3), m.group(2) or m.group(4)))
        return list(range(a, b + 1))
    nums = re.findall(r"(\d+)\s*ページ目|\bpage\s*(\d+)", t)
    pages = sorted({int(a or b) for a, b in nums})
    return pages or None


def ocr_prompt(text: str) -> str:
    stripped = text.strip()
    for p in OCR_TASK_PROMPTS:
        if stripped.startswith(p):
            return p
    t = stripped.lower()
    if re.search(r"table|表", t):
        return "Table Recognition:"
    if re.search(r"formula|equation|latex|数式", t):
        return "Formula Recognition:"
    return "Text Recognition:"


def build_ocr_jobs(body: dict) -> tuple[list[dict], str, str] | None:
    """For OCR models: the images/pages of the last user turn, the task prompt and a source name."""
    users = [m for m in body.get("messages", []) if m.get("role") == "user"]
    if not users or not isinstance(users[-1].get("content"), list):
        return None
    content = users[-1]["content"]
    text = message_text(content)
    pages = requested_pages(text)
    images: list[dict] = []
    source = "ocr"
    for part in content:
        payload = file_part_payload(part)
        if payload:
            source = Path(payload[0]).stem or source
        if payload and payload[1] == "application/pdf":
            images.extend(pdf_to_image_parts(payload[2], pages))
        elif payload and payload[1].startswith("image/"):
            url = "data:" + payload[1] + ";base64," + base64.b64encode(payload[2]).decode()
            images.append({"type": "image_url", "image_url": {"url": url}})
        elif part.get("type") == "image_url":
            images.append(part)
    if not images:
        return None
    return images, ocr_prompt(text), source


# ---------------------------------------------------------------- OCR execution

REPEAT_MIN_SPAN = 300  # chars of back-to-back repetition treated as a loop
REPEAT_MIN_COUNT = 10
REPEAT_MAX_UNIT = 60
HOLD_BACK = 600  # chars withheld while streaming so a loop can be cut before it is sent


def repetition_start(text: str) -> int | None:
    """Index where a trailing loop (one unit repeated back-to-back) begins, if any."""
    tail = text[-(REPEAT_MIN_SPAN * 2):]
    for size in range(1, REPEAT_MAX_UNIT + 1):
        unit = tail[-size:]
        count, pos = 0, len(tail)
        while pos >= size and tail[pos - size:pos] == unit:
            count, pos = count + 1, pos - size
        if count >= REPEAT_MIN_COUNT and count * size >= REPEAT_MIN_SPAN:
            # The loop may extend past the inspected tail; walk back over the full text
            pos = len(text) - count * size
            while pos >= size and text[pos - size:pos] == unit:
                pos -= size
            return pos
    return None


async def ocr_page(model: str, img: dict, prompt: str, headers: dict, usage: dict):
    """Streams OCR text for one image, cutting the output when the model starts looping."""
    req = {"model": model, "stream": True, "stream_options": {"include_usage": True},
           "max_tokens": OCR_MAX_TOKENS, "temperature": 0,
           "messages": [{"role": "user", "content": [img, {"type": "text", "text": prompt}]}]}
    text, sent = "", 0
    async with client.stream("POST", f"{UPSTREAM}/v1/chat/completions", json=req, headers=headers) as r:
        if r.status_code != 200:
            err = (await r.aread()).decode(errors="replace")
            yield f"[OCR error: {err}]"
            return
        async for line in r.aiter_lines():
            if not line.startswith("data: {"):
                continue
            data = json.loads(line[6:])
            for k, v in (data.get("usage") or {}).items():
                if k in usage:
                    usage[k] += v
            for choice in data.get("choices") or []:
                text += (choice.get("delta") or {}).get("content") or ""
            cut = repetition_start(text)
            if cut is not None:
                log.info("OCR loop detected; truncating %d chars", len(text) - cut)
                text = text[:cut]
                break
            if len(text) - HOLD_BACK > sent:
                yield text[sent:len(text) - HOLD_BACK]
                sent = len(text) - HOLD_BACK
    if len(text) > sent:
        yield text[sent:]


def chunk(cid: str, model: str, delta: dict, finish: str | None = None) -> str:
    obj = {
        "id": cid, "object": "chat.completion.chunk", "created": int(time.time()), "model": model,
        "choices": [{"index": 0, "delta": delta, "finish_reason": finish}],
    }
    return f"data: {json.dumps(obj, ensure_ascii=False)}\n\n"


def export_xlsx(source: str, sheets: list[tuple[str, tables.Table]]) -> str:
    """Writes the workbook and returns its download URL; prunes expired exports."""
    EXPORT_DIR.mkdir(parents=True, exist_ok=True)
    cutoff = time.time() - EXPORT_TTL_HOURS * 3600
    for old in EXPORT_DIR.glob("*/*.xlsx"):
        if old.stat().st_mtime < cutoff:
            old.unlink(missing_ok=True)
            old.parent.rmdir()
    token = uuid.uuid4().hex
    name = re.sub(r'[\\/:*?"<>|]', "_", source) + ".xlsx"
    (EXPORT_DIR / token).mkdir()
    tables.write_xlsx(str(EXPORT_DIR / token / name), sheets)
    return f"{PUBLIC_BASE_URL}/exports/{token}/{quote(name)}"


async def run_ocr(body: dict, images: list[dict], prompt: str, source: str, headers: dict) -> Response:
    model = body["model"]
    cid = f"chatcmpl-{uuid.uuid4().hex[:12]}"
    multi = len(images) > 1
    usage = {"prompt_tokens": 0, "completion_tokens": 0, "total_tokens": 0}

    async def pieces():
        # Pages are buffered so their HTML tables can be converted as a whole
        sheets: list[tuple[str, tables.Table]] = []
        for i, img in enumerate(images, 1):
            raw = "".join([t async for t in ocr_page(model, img, prompt, headers, usage)])
            text, found = tables.convert_html_tables(raw)
            for n, t in enumerate(found, 1):
                sheets.append((f"Page{i}" + (f"_{n}" if len(found) > 1 else ""), t))
            heading = f"## Page {i}\n\n" if multi else ""
            yield ("\n\n" if i > 1 else "") + heading + text
        if sheets:
            try:
                url = export_xlsx(source, sheets)
                log.info("exported %d table(s) to %s", len(sheets), url)
                yield f"\n\n---\n\n📥 [Excelでダウンロード ({len(sheets)}表)]({url})"
            except Exception:
                log.exception("xlsx export failed")
                yield "\n\n---\n\n(Excelファイルの作成に失敗しました)"

    if not body.get("stream"):
        content = "".join([t async for t in pieces()])
        return JSONResponse({
            "id": cid, "object": "chat.completion", "created": int(time.time()), "model": model,
            "choices": [{"index": 0, "message": {"role": "assistant", "content": content},
                         "finish_reason": "stop"}],
            "usage": usage,
        })

    include_usage = (body.get("stream_options") or {}).get("include_usage")

    async def gen():
        yield chunk(cid, model, {"role": "assistant", "content": ""})
        async for t in pieces():
            yield chunk(cid, model, {"content": t})
        yield chunk(cid, model, {}, "stop")
        if include_usage:
            obj = {"id": cid, "object": "chat.completion.chunk", "created": int(time.time()),
                   "model": model, "choices": [], "usage": usage}
            yield f"data: {json.dumps(obj)}\n\n"
        yield "data: [DONE]\n\n"

    return StreamingResponse(gen(), media_type="text/event-stream")


# ---------------------------------------------------------------- HTTP

HOP_HEADERS = {"host", "content-length", "connection", "accept-encoding", "transfer-encoding"}


def forward_headers(request: Request) -> dict:
    return {k: v for k, v in request.headers.items() if k.lower() not in HOP_HEADERS}


async def passthrough(request: Request, body: bytes) -> Response:
    url = f"{UPSTREAM}{request.url.path}"
    if request.url.query:
        url += f"?{request.url.query}"
    req = client.build_request(request.method, url, content=body, headers=forward_headers(request))
    r = await client.send(req, stream=True)
    headers = {k: v for k, v in r.headers.items() if k.lower() not in HOP_HEADERS | {"content-encoding"}}

    async def body_iter():
        try:
            async for b in r.aiter_raw():
                yield b
        finally:
            await r.aclose()

    return StreamingResponse(body_iter(), status_code=r.status_code, headers=headers)


async def chat_completions(request: Request) -> Response:
    raw = await request.body()
    try:
        body = json.loads(raw)
    except ValueError:
        return await passthrough(request, raw)

    model = body.get("model", "")
    has_files = any(
        isinstance(m.get("content"), list) and any(p.get("type") == "file" for p in m["content"])
        for m in body.get("messages", [])
    )

    if is_ocr_model(model):
        built = build_ocr_jobs(body)
        if built:
            images, prompt, source = built
            log.info("OCR model=%s pages=%d prompt=%r", model, len(images), prompt)
            return await run_ocr(body, images, prompt, source, forward_headers(request))

    if has_files:
        await rewrite_file_parts(body, forward_headers(request))
        log.info("rewrote file parts for model=%s", model)
        raw = json.dumps(body).encode()

    return await passthrough(request, raw)


async def any_route(request: Request) -> Response:
    return await passthrough(request, await request.body())


async def health(_: Request) -> Response:
    return JSONResponse({"ok": True})


async def download(request: Request) -> Response:
    token, name = request.path_params["token"], request.path_params["name"]
    if not re.fullmatch(r"[0-9a-f]{32}", token) or "/" in name or name.startswith("."):
        return Response(status_code=404)
    path = EXPORT_DIR / token / name
    if not path.is_file():
        return Response("このファイルは期限切れか存在しません", status_code=404)
    return FileResponse(path, filename=name, media_type=(
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"))


app = Starlette(routes=[
    Route("/healthz", health),
    Route("/exports/{token}/{name}", download, methods=["GET"]),
    Route("/v1/chat/completions", chat_completions, methods=["POST"]),
    Route("/{path:path}", any_route, methods=["GET", "POST", "PUT", "DELETE", "PATCH", "HEAD", "OPTIONS"]),
])
