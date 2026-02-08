from __future__ import annotations

import os
from typing import Optional

MAX_IMAGE_DIM = 1600  # Reduce OCR load for large images


class OcrError(RuntimeError):
    """Raised when OCR extraction fails."""


def extract_text(image_path: str) -> str:
    try:
        from PIL import Image  # type: ignore[import-not-found]
    except Exception as exc:
        raise OcrError("Pillow is not installed.") from exc

    try:
        import pytesseract  # type: ignore[import-not-found]
    except Exception as exc:
        raise OcrError("pytesseract is not installed.") from exc

    tesseract_cmd = os.getenv("TESSERACT_CMD")
    if tesseract_cmd:
        pytesseract.pytesseract.tesseract_cmd = tesseract_cmd

    image = Image.open(image_path).convert("RGB")
    max_dim = max(image.size)
    if max_dim > MAX_IMAGE_DIM:
        scale = MAX_IMAGE_DIM / max_dim
        new_size = (int(image.size[0] * scale), int(image.size[1] * scale))
        image = image.resize(new_size, Image.LANCZOS)
    text = pytesseract.image_to_string(image, lang="kor+eng")
    return text.strip()


def extract_text_from_pil_image(image) -> str:  # type: ignore[no-untyped-def]
    """
    OCR helper for in-memory PIL images.
    """
    try:
        import pytesseract  # type: ignore[import-not-found]
    except Exception as exc:
        raise OcrError("pytesseract is not installed.") from exc

    try:
        text = pytesseract.image_to_string(image, lang="kor+eng")
        return (text or "").strip()
    except Exception as exc:
        raise OcrError(str(exc)) from exc
