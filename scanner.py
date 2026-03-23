# scanner.py
# ─── Gemini-powered OCR and contact extraction (google-genai SDK) ─────────────

import re
import json
import time
from pathlib import Path

import pdfplumber
from docx import Document
from PIL import Image
from google import genai
from google.genai import types

from config import (
    GEMINI_API_KEY, GEMINI_MODEL,
    REQUEST_TIMEOUT, MAX_RETRIES, RETRY_DELAY,
    DOCX_TEXT_LIMIT, PDF_TEXT_LIMIT, GARBAGE_THRESHOLD
)

# ─── Supported extensions ─────────────────────────────────────────────────────

SUPPORTED_IMAGE_EXTS = {".jpg", ".jpeg", ".png", ".tiff", ".tif", ".bmp", ".webp"}
SUPPORTED_PDF_EXTS   = {".pdf"}
SUPPORTED_WORD_EXTS  = {".docx"}
ALL_SUPPORTED        = SUPPORTED_IMAGE_EXTS | SUPPORTED_PDF_EXTS | SUPPORTED_WORD_EXTS

# ─── Gemini client (single instance, reused for all calls) ───────────────────

CLIENT = genai.Client(api_key=GEMINI_API_KEY)

# ─── Prompts ──────────────────────────────────────────────────────────────────

IMAGE_PROMPT = """You are a contact information extractor specializing in visiting cards.

Look at this image carefully. Extract the following:
- Full name of the person (not the company name)
- Phone number (mobile preferred, include country code if visible)
- Email address

Return ONLY this JSON and nothing else:
{"name": "...", "phone": "...", "email": "..."}

Rules:
- If a field is not visible or not present, use null
- Do not guess or hallucinate any information
- Do not return any text outside the JSON object
- If multiple phone numbers exist, prefer the mobile number
- If multiple emails exist, prefer the personal or direct one"""


def build_text_prompt(text: str) -> str:
    return f"""You are a contact information extractor.

The following text is from a document — it may be a resume,
business profile, letterhead, or contact sheet.
Only the first section of the document is provided.

Extract the following about the PRIMARY person in the document:
- Their full name (not a company name, not a department)
- Their direct phone number (mobile preferred over landline)
- Their personal or professional email address

Return ONLY this JSON and nothing else:
{{"name": "...", "phone": "...", "email": "..."}}

Rules:
- If a field does not exist, use null
- Do not guess or hallucinate
- Do not return any text outside the JSON object
- If multiple people are mentioned, pick the most prominent one
- If multiple phones exist, prefer mobile
- The name must be a real human name, not a job title or company

Text:
{text}"""


def build_retry_prompt(content: str) -> str:
    return f"""Your previous response was not valid JSON.

Extract contact information and return ONLY this exact format:
{{"name": "...", "phone": "...", "email": "..."}}

Use null for any missing fields.
No explanation. No extra text. JSON only.

Content:
{content}"""


# ─── Generation config ────────────────────────────────────────────────────────

GEN_CONFIG = types.GenerateContentConfig(
    temperature=0.0,
    max_output_tokens=256,
)

# ─── Gemini callers ───────────────────────────────────────────────────────────

def _call_gemini_with_image(image_path: Path) -> str:
    """Send an image file to Gemini Vision and return raw text response."""
    img = Image.open(image_path)

    for attempt in range(MAX_RETRIES):
        try:
            response = CLIENT.models.generate_content(
                model=GEMINI_MODEL,
                contents=[IMAGE_PROMPT, img],
                config=GEN_CONFIG,
            )
            return response.text.strip()
        except Exception as e:
            err = str(e)
            if attempt < MAX_RETRIES - 1:
                wait = 15 if ("429" in err or "quota" in err.lower() or "rate" in err.lower()) else RETRY_DELAY
                time.sleep(wait)
                continue
            raise RuntimeError(_classify_error(err))


def _call_gemini_with_text(text: str) -> str:
    """Send a text prompt to Gemini and return raw text response."""
    prompt = build_text_prompt(text)

    for attempt in range(MAX_RETRIES):
        try:
            response = CLIENT.models.generate_content(
                model=GEMINI_MODEL,
                contents=prompt,
                config=GEN_CONFIG,
            )
            return response.text.strip()
        except Exception as e:
            err = str(e)
            if attempt < MAX_RETRIES - 1:
                wait = 15 if ("429" in err or "quota" in err.lower() or "rate" in err.lower()) else RETRY_DELAY
                time.sleep(wait)
                continue
            raise RuntimeError(_classify_error(err))


def _call_gemini_retry(content: str) -> str:
    """Stricter retry when first response was bad JSON."""
    prompt = build_retry_prompt(content)
    try:
        response = CLIENT.models.generate_content(
            model=GEMINI_MODEL,
            contents=prompt,
            config=GEN_CONFIG,
        )
        return response.text.strip()
    except Exception as e:
        raise RuntimeError(_classify_error(str(e)))


def _classify_error(err: str) -> str:
    """Map raw API errors to clean user-facing messages."""
    e = err.lower()
    if "api_key" in e or "api key" in e or "401" in err:
        return "Invalid API key — check config.py"
    if "403" in err or "permission" in e:
        return "API key doesn't have permission for this operation"
    if "429" in err or "quota" in e or "rate" in e:
        return "Rate limit reached — wait a moment and try again"
    if "500" in err or "internal" in e:
        return "Gemini server error — try again shortly"
    if "timeout" in e or "timed out" in e:
        return "Request timed out — check your connection"
    if "name resolution" in e or "connection" in e or "network" in e:
        return "Cannot reach Gemini API — check your internet connection"
    if "400" in err or "invalid" in e:
        return "Bad request — file may be corrupted or unsupported"
    return f"API error: {err}"


# ─── JSON parsing ─────────────────────────────────────────────────────────────

def _parse_json_response(raw: str) -> dict | None:
    """Extract JSON from Gemini response, handling markdown fences."""
    cleaned = re.sub(r"```(?:json)?", "", raw).replace("```", "").strip()

    try:
        return json.loads(cleaned)
    except json.JSONDecodeError:
        pass

    match = re.search(r"\{.*?\}", cleaned, re.DOTALL)
    if match:
        try:
            return json.loads(match.group())
        except json.JSONDecodeError:
            pass

    return None


def _validate_and_clean(data: dict) -> dict:
    """Validate and clean extracted contact fields."""
    name  = data.get("name")
    phone = data.get("phone")
    email = data.get("email")

    # Name
    if name and isinstance(name, str):
        name = name.strip()
        if name.isupper():
            name = name.title()
        company_markers = ["inc", "ltd", "pvt", "llc", "corp", "co.", "company"]
        if any(m in name.lower() for m in company_markers):
            name = None
    else:
        name = None

    # Phone
    if phone and isinstance(phone, str):
        phone = phone.strip()
        digits = re.sub(r"\D", "", phone)
        if len(digits) < 7:
            phone = None
    else:
        phone = None

    # Email
    if email and isinstance(email, str):
        email = email.strip().lower()
        if "@" not in email or "." not in email.split("@")[-1]:
            email = None
    else:
        email = None

    return {
        "name":  name  or "N/A",
        "phone": phone or "N/A",
        "email": email or "N/A",
    }


# ─── Text extraction helpers ──────────────────────────────────────────────────

def _is_garbage(text: str) -> bool:
    printable = sum(c.isprintable() and not c.isspace() for c in text)
    return printable < GARBAGE_THRESHOLD


def _extract_pdf_text(path: Path) -> str | None:
    try:
        texts = []
        with pdfplumber.open(path) as pdf:
            for page in pdf.pages:
                t = page.extract_text()
                if t:
                    texts.append(t)
        combined = "\n".join(texts).strip()
        if _is_garbage(combined):
            return None
        return combined[:PDF_TEXT_LIMIT]
    except Exception as e:
        if "encrypted" in str(e).lower() or "password" in str(e).lower():
            raise RuntimeError("File is encrypted — cannot read")
        raise RuntimeError(f"Could not open PDF: {e}")


def _extract_docx_text(path: Path) -> str | None:
    try:
        doc = Document(path)
        lines = [p.text.strip() for p in doc.paragraphs if p.text.strip()]
        combined = "\n".join(lines)
        if _is_garbage(combined):
            return None
        return combined[:DOCX_TEXT_LIMIT]
    except Exception as e:
        raise RuntimeError(f"Could not open Word document: {e}")


# ─── Main scan entry point ────────────────────────────────────────────────────

def scan_file(file_path: str) -> dict:
    path = Path(file_path)

    if not path.exists():
        return _error("File not found — check the path")

    ext = path.suffix.lower()
    if ext not in ALL_SUPPORTED:
        return _error(f"Unsupported format '{ext}' — use JPG, PNG, PDF, or DOCX")

    try:
        if ext in SUPPORTED_IMAGE_EXTS:
            return _scan_image(path)
        elif ext in SUPPORTED_PDF_EXTS:
            return _scan_pdf(path)
        elif ext in SUPPORTED_WORD_EXTS:
            return _scan_docx(path)
    except RuntimeError as e:
        return _error(str(e))
    except Exception as e:
        return _error(f"Unexpected error: {e}")


def _scan_image(path: Path) -> dict:
    try:
        raw = _call_gemini_with_image(path)
    except RuntimeError as e:
        return _error(str(e))
    return _process_raw_response(raw, str(path))


def _scan_pdf(path: Path) -> dict:
    try:
        text = _extract_pdf_text(path)
    except RuntimeError as e:
        return _error(str(e))

    if text:
        try:
            raw = _call_gemini_with_text(text)
        except RuntimeError as e:
            return _error(str(e))
        return _process_raw_response(raw, text)
    else:
        return _error(
            "This PDF appears to be a scanned image with no text layer. "
            "Export the card as a JPG or PNG and scan that instead."
        )


def _scan_docx(path: Path) -> dict:
    try:
        text = _extract_docx_text(path)
    except RuntimeError as e:
        return _error(str(e))

    if not text:
        return _error("Document appears to be empty")

    try:
        raw = _call_gemini_with_text(text)
    except RuntimeError as e:
        return _error(str(e))

    return _process_raw_response(raw, text)


def _process_raw_response(raw: str, original_content: str) -> dict:
    data = _parse_json_response(raw)

    if data is None:
        try:
            raw2 = _call_gemini_retry(
                original_content if len(original_content) < 500 else raw
            )
            data = _parse_json_response(raw2)
        except RuntimeError as e:
            return _error(str(e))

    if data is None:
        return _error("Could not parse contact data from this file")

    contacts = _validate_and_clean(data)
    missing  = [k for k, v in contacts.items() if v == "N/A"]

    if len(missing) == 3:
        return _error("No contact information found in this file")

    return {
        "status":  "success" if not missing else "partial",
        "message": "Successfully scanned." if not missing
                   else f"Partial data — {', '.join(missing)} not found",
        "data": contacts
    }


def _error(message: str) -> dict:
    return {"status": "error", "message": message, "data": None}
