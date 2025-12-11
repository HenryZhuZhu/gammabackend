import os
import io
import time
from typing import List, Dict, Any

import requests
from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import JSONResponse, StreamingResponse
from pptx import Presentation


# ================= 配置区域 =================

# 在 Render 环境变量中配置（不要写死在代码里）
GAMMA_API_KEY = os.getenv("GAMMA_API_KEY")          # 必填：你的 Gamma API Key
GAMMA_TEMPLATE_ID = os.getenv("GAMMA_TEMPLATE_ID")  # 必填：你在 Gamma 里选好的模板 gammaId

# 可选：如果你想强制指定主题 / 文件夹，可以额外加这些环境变量
GAMMA_THEME_ID = os.getenv("GAMMA_THEME_ID")        # 选填：themeId，不填则用模板的主题:contentReference[oaicite:3]{index=3}
GAMMA_FOLDER_IDS = os.getenv("GAMMA_FOLDER_IDS")    # 选填：逗号分隔的 folderId 列表
GAMMA_EXPORT_FORMAT = os.getenv("GAMMA_EXPORT_AS", "pdf")  # "pdf" 或 "pptx":contentReference[oaicite:4]{index=4}

GAMMA_BASE_URL = "https://public-api.gamma.app/v1.0"

# 轮询生成状态的配置：每 5 秒查一次，最多等 2 分钟:contentReference[oaicite:5]{index=5}
POLL_INTERVAL_SECONDS = 5
MAX_POLL_SECONDS = 120

app = FastAPI(title="AIStoryteller Gamma Backend (Template Mode)")

# CORS：先全开，方便你在 Netlify 前端调试；以后可以改成只允许你的域名
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],  # 比如改成 ["https://你的netlify域名.netlify.app"]
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)


# =============== 工具函数：从 PPT 提取文本 ===============

def extract_ppt_structure_and_text(file_bytes: bytes) -> Dict[str, Any]:
    """
    输入 PPTX 二进制，输出：
      - slides: 用于前端预览的结构（slides -> shapes -> text）
      - outline_text: 传给 Gamma 的文本大纲（包含每页的内容）
    """
    try:
        from io import BytesIO
        prs = Presentation(BytesIO(file_bytes))
    except Exception as e:
        raise HTTPException(status_code=400, detail=f"Failed to open PPTX: {e}")

    slides_data: List[Dict[str, Any]] = []
    outline_parts: List[str] = []

    for idx, slide in enumerate(prs.slides, start=1):
        shapes_data = []
        slide_text_lines = []

        for shape in slide.shapes:
            text = ""
            if hasattr(shape, "has_text_frame") and shape.has_text_frame:
                paragraphs = [p.text for p in shape.text_frame.paragraphs if p.text]
                text = "\n".join(paragraphs).strip()

            if text:
                shapes_data.append({"text": text})
                slide_text_lines.append(text)

        slides_data.append(
            {
                "index": idx,
                "shapes": shapes_data,
            }
        )

        if slide_text_lines:
            outline_parts.append(f"Slide {idx}:\n" + "\n".join(slide_text_lines))
        else:
            outline_parts.append(f"Slide {idx}:\n(No visible text content)")

    # 把每一页的内容用 --- 分隔，作为发送给 Gamma 的“原始内容”:contentReference[oaicite:6]{index=6}
    outline_text = "\n---\n".join(outline_parts)

    return {
        "slides": slides_data,
        "outline_text": outline_text,
    }


# =============== 调用 Gamma Create-from-template ===============

def call_gamma_from_template(prompt_text: str) -> str:
    """
    使用 Gamma 的“Create from template”接口创建新的 gamma 内容：
    POST https://public-api.gamma.app/v1.0/generations/from-template

    返回 generationId，用于后续轮询。
    """
    if not GAMMA_API_KEY:
        raise HTTPException(
            status_code=500,
            detail="GAMMA_API_KEY is not set. Please configure it in Render environment variables.",
        )

    if not GAMMA_TEMPLATE_ID:
        raise HTTPException(
            status_code=500,
            detail="GAMMA_TEMPLATE_ID is not set. Please set it to your template gammaId.",
        )

    url = f"{GAMMA_BASE_URL}/generations/from-template"

    # 解析 folderIds（如果配置了的话）
    folder_ids = None
    if GAMMA_FOLDER_IDS:
        folder_ids = [f.strip() for f in GAMMA_FOLDER_IDS.split(",") if f.strip()]

    payload: Dict[str, Any] = {
        "gammaId": GAMMA_TEMPLATE_ID,        # 模板 ID（必填）
        "prompt": prompt_text,               # 输入内容+指令（必填）
        "exportAs": GAMMA_EXPORT_FORMAT,     # "pdf" 或 "pptx"
    }

    if GAMMA_THEME_ID:
        payload["themeId"] = GAMMA_THEME_ID  # 覆盖模板主题（可选）

    if folder_ids:
        payload["folderIds"] = folder_ids    # 保存到指定文件夹（可选）

    headers = {
        "Content-Type": "application/json",
        "X-API-KEY": GAMMA_API_KEY,
    }

    try:
        resp = requests.post(url, json=payload, headers=headers, timeout=60)
    except Exception as e:
        raise HTTPException(status_code=502, detail=f"Failed to call Gamma: {e}")

    # ✅ 关键修复点：接受所有 2xx 为成功，而不是只认 200
    if not resp.ok:
        # 非 2xx：真正的错误，直接把 Gamma 返回内容透传出去方便调试
        raise HTTPException(
            status_code=resp.status_code,
            detail=f"Gamma create-from-template error: {resp.text}",
        )

    try:
        data = resp.json()
    except Exception as e:
        raise HTTPException(
            status_code=500,
            detail=f"Gamma response is not valid JSON: {e}, raw: {resp.text}",
        )

    generation_id = data.get("generationId")
    if not generation_id:
        raise HTTPException(
            status_code=500,
            detail=f"Gamma did not return generationId. Raw response: {data}",
        )

    return generation_id


def poll_gamma_result(generation_id: str) -> Dict[str, Any]:
    """
    轮询 Gamma 生成状态 + 获取文件 URL：
    GET https://public-api.gamma.app/v1.0/generations/{generationId} :contentReference[oaicite:13]{index=13}

    - status: pending / completed / error
    - gammaUrl: 在线编辑地址
    - file URLs: 如果 exportAs=pdf/pptx，会返回对应的文件 URL
    """
    if not GAMMA_API_KEY:
        raise HTTPException(
            status_code=500,
            detail="GAMMA_API_KEY is not set. Please configure it in Render environment variables.",
        )

    url = f"{GAMMA_BASE_URL}/generations/{generation_id}"
    headers = {
        "X-API-KEY": GAMMA_API_KEY,
        "accept": "application/json",
    }

    start_time = time.time()
    while True:
        try:
            resp = requests.get(url, headers=headers, timeout=30)
        except Exception as e:
            raise HTTPException(status_code=502, detail=f"Failed to poll Gamma: {e}")

        if resp.status_code != 200:
            raise HTTPException(
                status_code=resp.status_code,
                detail=f"Gamma poll error: {resp.text}",
            )

        data = resp.json()
        status = data.get("status")
        if status == "completed":
            return data
        elif status in ("failed", "error"):
            raise HTTPException(
                status_code=500,
                detail=f"Gamma generation failed: {data}",
            )

        elapsed = time.time() - start_time
        if elapsed > MAX_POLL_SECONDS:
            raise HTTPException(
                status_code=504,
                detail="Gamma generation timed out. Please try again.",
            )

        time.sleep(POLL_INTERVAL_SECONDS)


def download_gamma_file(gamma_result: Dict[str, Any], expected_format: str = "pdf") -> bytes:
    """
    从 Gamma 轮询结果中找到 PDF / PPTX 的 URL，然后下载文件字节。:contentReference[oaicite:14]{index=14}

    文档说明 GET /generations/{id} 会返回：
      - gammaUrl（在线编辑）
      - 文件 URL（如果你在请求里设置了 exportAs=pdf/pptx）
    但示例没有完全列出字段名，所以这里做一些兼容处理。
    """
    file_url = None

    # 常见结构：{"fileUrls": {"pdf": "...", "pptx": "..."}} 或 {"files": {...}}
    file_urls = gamma_result.get("fileUrls") or gamma_result.get("files") or {}
    if isinstance(file_urls, dict):
        file_url = file_urls.get(expected_format)

    # 兜底字段名（以防将来文档更新）
    if not file_url and expected_format == "pdf":
        file_url = gamma_result.get("pdfUrl")
    if not file_url and expected_format == "pptx":
        file_url = gamma_result.get("pptxUrl")

    if not file_url:
        raise HTTPException(
            status_code=500,
            detail=f"Gamma result completed but did not include {expected_format} file URL. Raw result: {gamma_result}",
        )

    try:
        resp = requests.get(file_url, timeout=120)
    except Exception as e:
        raise HTTPException(status_code=502, detail=f"Failed to download file from Gamma: {e}")

    if resp.status_code != 200:
        raise HTTPException(
            status_code=resp.status_code,
            detail=f"Failed to download file from Gamma: {resp.text}",
        )

    return resp.content


# =============== API 路由 ===============

@app.get("/health")
def health_check():
    return {"status": "ok"}


@app.post("/api/parse_ppt")
async def parse_ppt(file: UploadFile = File(...)):
    """
    接收用户上传的 PPTX，解析文字结构，返回给前端做预览用。
    """
    if not file.filename.lower().endswith(".pptx"):
        raise HTTPException(status_code=400, detail="Only .pptx files are supported")

    file_bytes = await file.read()
    result = extract_ppt_structure_and_text(file_bytes)

    # 只返回 slides，outline_text 不发给前端（只在服务器上用来喂给 Gamma）
    return JSONResponse({"slides": result["slides"]})


@app.post("/api/beautify")
async def beautify_ppt(file: UploadFile = File(...)):
    """
    核心流程：
      1. 接收 PPTX
      2. 提取文本 → outline_text
      3. 组织一个 prompt，告诉 Gamma 如何基于模板改写这个 deck
      4. 调用 Create-from-template 接口
      5. 轮询生成状态
      6. 下载 PDF/PPTX 并返回给前端
    """
    if not file.filename.lower().endswith(".pptx"):
        raise HTTPException(status_code=400, detail="Only .pptx files are supported")

    file_bytes = await file.read()
    parsed = extract_ppt_structure_and_text(file_bytes)
    outline_text = parsed["outline_text"]

    # 组织 prompt：说明用途 + 附上原始 slide 文本
    # 你可以根据自己的业务风格继续微调这段 instruction
    prompt = (
        "You are an expert presentation designer. "
        "Please take the following slide contents from a partly-finished deck and "
        "rebuild them using the provided Gamma template. "
        "Keep the key points and slide structure, but improve clarity, hierarchy, and flow. "
        "Make it suitable for semiconductor testing / AI product reviews, with a professional tone. "
        "Avoid inventing unrelated business details.\n\n"
        "Here is the slide content to transform:\n\n"
        f"{outline_text}"
    )

    # 2. 调用 Create-from-template
    generation_id = call_gamma_from_template(prompt)

    # 3. 轮询结果
    result = poll_gamma_result(generation_id)

    # 4. 下载文件（默认 pdf，如果你把 GAMMA_EXPORT_AS 改成 pptx，就会下 PPTX）
    file_bytes_out = download_gamma_file(result, expected_format=GAMMA_EXPORT_FORMAT)

    filename_root = os.path.splitext(file.filename)[0]
    ext = "pdf" if GAMMA_EXPORT_FORMAT.lower() == "pdf" else "pptx"
    output_filename = f"{filename_root}_gamma_beautified.{ext}"

    media_type = (
        "application/pdf"
        if ext == "pdf"
        else "application/vnd.openxmlformats-officedocument.presentationml.presentation"
    )

    return StreamingResponse(
        io.BytesIO(file_bytes_out),
        media_type=media_type,
        headers={
            "Content-Disposition": f'attachment; filename="{output_filename}"'
        },
    )
