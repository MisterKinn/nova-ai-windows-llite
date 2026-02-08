from __future__ import annotations

import sys
from dataclasses import dataclass
from typing import Literal, Optional, Tuple


def _debug(msg: str) -> None:
    sys.stderr.write(f"[Layout Debug] {msg}\n")
    sys.stderr.flush()


ContainerTemplate = Literal["header.hwp", "box.hwp", "box_white.hwp"]


@dataclass(frozen=True)
class ContainerDetection:
    """
    Result of container detection on an image.

    - template: which HWP template should be inserted
    - rect: (x, y, w, h) bounding box of the container (in original image coordinates),
            if detected. For header-only detection (no border), rect can be None.
    - has_view_text: whether "<보기>"-like header text was detected
    - border_score: 0..1 score indicating border strength along detected rectangle
    """

    template: Optional[ContainerTemplate]
    rect: Optional[Tuple[int, int, int, int]]
    has_view_text: bool
    border_score: float


def detect_container(image_path: str) -> ContainerDetection:
    """
    Detect a <보기>/box-like container and choose the correct template.

    Heuristics:
    - Detect the best rectangle candidate (container border) using edge + contour geometry.
    - Compute border strength along the rectangle perimeter.
    - Detect '<보기>' text using pytesseract word boxes; tolerate spaced '< 보 기 >'.
    - Decision:
        - If view text is found: header.hwp
        - Else if rectangle exists and border strong: box.hwp
        - Else if rectangle exists and border weak: box_white.hwp
        - Else: template=None (no container)
    """
    _debug(f"Detecting container for: {image_path}")

    has_view_text, view_bbox = _detect_view_text_bbox(image_path)
    _debug(f"View text detected: {has_view_text}, bbox: {view_bbox}")
    
    rect, border_score = _detect_best_rectangle(image_path)
    _debug(f"Rectangle detected: {rect}, border_score: {border_score:.3f}")

    template: Optional[ContainerTemplate] = None
    # If OCR fails to read '<보기>' (common when border breaks), infer from border gap pattern.
    if (not has_view_text) and rect is not None:
        try:
            if _infer_view_from_border_gap(image_path, rect):
                has_view_text = True
        except Exception:
            pass

    if has_view_text:
        template = "header.hwp"
    elif rect is not None:
        template = "box.hwp" if border_score >= 0.35 else "box_white.hwp"

    # If view text exists but we also have a rectangle, prefer the rectangle as the rect.
    # If view text exists but no rectangle, keep rect=None and still use header.hwp.
    _ = view_bbox  # currently unused but kept for future refinements
    return ContainerDetection(
        template=template,
        rect=rect,
        has_view_text=has_view_text,
        border_score=float(border_score),
    )


def _detect_view_text_bbox(image_path: str) -> tuple[bool, Optional[Tuple[int, int, int, int]]]:
    """
    Returns (has_view_text, bbox).
    bbox is best-effort union bbox of the detected '<보기>' token(s).
    """

    try:
        from PIL import Image  # type: ignore[import-not-found]
        import pytesseract  # type: ignore[import-not-found]
    except Exception:
        return False, None

    # Respect optional environment override used elsewhere in Nova AI Lite.
    try:
        import os

        tesseract_cmd = os.getenv("TESSERACT_CMD")
        if tesseract_cmd:
            pytesseract.pytesseract.tesseract_cmd = tesseract_cmd
    except Exception:
        pass

    try:
        img = Image.open(image_path).convert("RGB")
    except Exception:
        return False, None

    # Fallback: raw OCR string (bbox not guaranteed, but better than missing the signal)
    try:
        raw = pytesseract.image_to_string(img, lang="kor+eng")
        # Remove all whitespace (spaces/newlines/tabs) for robust matching.
        raw_norm = "".join((raw or "").split())
        raw_norm = raw_norm.replace("〈", "<").replace("〉", ">").replace("《", "<").replace("》", ">")
        raw_norm = raw_norm.replace("＜", "<").replace("＞", ">")
        if "보기" in raw_norm or "<보기>" in raw_norm or "<보기" in raw_norm:
            # We don't know bbox here; return True with bbox=None.
            return True, None
        # spaced "<보 기>" style
        if ("보" in raw_norm and "기" in raw_norm) and ("<" in raw_norm or ">" in raw_norm):
            return True, None
    except Exception:
        pass

    try:
        data = pytesseract.image_to_data(img, lang="kor+eng", output_type=pytesseract.Output.DICT)
    except Exception:
        return False, None

    n = len(data.get("text", []))
    tokens: list[tuple[str, int, int, int, int, int]] = []
    for i in range(n):
        txt = (data["text"][i] or "").strip()
        # Some Korean headers are low-confidence; keep a low threshold.
        conf = int(float(data.get("conf", ["-1"] * n)[i] or -1))
        # If token looks like a header marker, keep it even with low conf.
        if conf < 5:
            txt_probe = (data["text"][i] or "").strip()
            if not any(ch in txt_probe for ch in ("보", "기", "<", ">", "〈", "〉", "《", "》", "＜", "＞")):
                continue
        if not txt:
            continue
        x = int(data["left"][i])
        y = int(data["top"][i])
        w = int(data["width"][i])
        h = int(data["height"][i])
        line = int(data.get("line_num", [0] * n)[i] or 0)
        tokens.append((txt, x, y, w, h, line))

    # Direct hit: '보기' or '<보기>' variants
    for txt, x, y, w, h, _line in tokens:
        normalized = txt.replace(" ", "")
        normalized = normalized.replace("〈", "<").replace("〉", ">").replace("《", "<").replace("》", ">")
        normalized = normalized.replace("＜", "<").replace("＞", ">")
        if "보기" in normalized:
            return True, (x, y, w, h)
        if "<보기>" in normalized:
            return True, (x, y, w, h)

    # Spaced pattern: '<' '보' '기' '>' (or without brackets)
    # Group by line and y proximity.
    by_line: dict[int, list[tuple[str, int, int, int, int]]] = {}
    for txt, x, y, w, h, line in tokens:
        by_line.setdefault(line, []).append((txt, x, y, w, h))
    for line, items in by_line.items():
        if not line:
            continue
        items.sort(key=lambda t: t[1])
        texts = [t[0] for t in items]
        joined = "".join(texts).replace(" ", "")
        joined = joined.replace("〈", "<").replace("〉", ">").replace("《", "<").replace("》", ">")
        joined = joined.replace("＜", "<").replace("＞", ">")
        if "보기" not in joined:
            continue
        # best-effort bbox: union of tokens that include 보/기 or brackets on that line
        xs: list[int] = []
        ys: list[int] = []
        xe: list[int] = []
        ye: list[int] = []
        for txt, x, y, w, h in items:
            t = txt.replace(" ", "")
            if any(ch in t for ch in ("보", "기", "<", ">", "〈", "〉", "《", "》", "＜", "＞")):
                xs.append(x)
                ys.append(y)
                xe.append(x + w)
                ye.append(y + h)
        if xs:
            return True, (min(xs), min(ys), max(xe) - min(xs), max(ye) - min(ys))

    return False, None


def _detect_best_rectangle(image_path: str) -> tuple[Optional[Tuple[int, int, int, int]], float]:
    """
    Detect the most likely container rectangle and return (rect, border_score).
    rect is (x, y, w, h) in original image coords.
    border_score is 0..1 edge density along perimeter.
    """

    try:
        import cv2  # type: ignore[import-not-found]
        import numpy as np  # type: ignore[import-not-found]
    except Exception:
        return None, 0.0

    img = cv2.imread(image_path)
    if img is None:
        return None, 0.0
    h, w = img.shape[:2]
    if h < 10 or w < 10:
        return None, 0.0

    # 이미지가 너무 크면 리사이즈해서 처리
    max_dim = 2000
    scale = 1.0
    if max(h, w) > max_dim:
        scale = max_dim / max(h, w)
        new_w, new_h = int(w * scale), int(h * scale)
        img = cv2.resize(img, (new_w, new_h))
        _debug(f"Resized image from {w}x{h} to {new_w}x{new_h} (scale={scale:.3f})")
        h, w = new_h, new_w

    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)

    # Line-based detection (more robust when borders are broken by '<보기>' header).
    try:
        bw = cv2.adaptiveThreshold(
            gray, 255, cv2.ADAPTIVE_THRESH_MEAN_C, cv2.THRESH_BINARY_INV, 21, 10
        )
        # Extract long horizontal/vertical strokes.
        hk = max(30, w // 25)
        vk = max(30, h // 25)
        h_kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (hk, 1))
        v_kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (1, vk))
        horizontal = cv2.erode(bw, h_kernel, iterations=1)
        horizontal = cv2.dilate(horizontal, h_kernel, iterations=1)
        vertical = cv2.erode(bw, v_kernel, iterations=1)
        vertical = cv2.dilate(vertical, v_kernel, iterations=1)
        grid = cv2.add(horizontal, vertical)
        grid = cv2.dilate(grid, np.ones((3, 3), np.uint8), iterations=1)
        contours, _ = cv2.findContours(grid, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
        edge_map = grid
        _debug(f"Grid-based contours found: {len(contours)}")
    except Exception as e:
        _debug(f"Grid method failed: {e}, using Canny fallback")
        # Edge fallback
        edges = cv2.Canny(gray, 40, 140)
        edges = cv2.dilate(edges, np.ones((3, 3), np.uint8), iterations=1)
        contours, _ = cv2.findContours(edges, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
        edge_map = edges
    candidates: list[Tuple[int, int, int, int, float]] = []
    img_area = float(w * h)
    _debug(f"Image size: {w}x{h}, area: {img_area}")

    for cnt in contours:
        x, y, ww, hh = cv2.boundingRect(cnt)
        area = float(ww * hh)
        if area < img_area * 0.02:
            continue
        if area > img_area * 0.90:
            continue
        if ww < 40 or hh < 20:
            continue
        aspect = ww / max(1.0, float(hh))
        if aspect < 1.1 or aspect > 30:
            continue
        # Exclude near full-page frames and very near borders (likely page edge)
        margin = int(min(w, h) * 0.02)
        if x <= margin and y <= margin:
            continue
        if (x + ww) >= (w - margin) and (y + hh) >= (h - margin):
            continue

        score = _border_score_on_rect(edge_map, (x, y, ww, hh))
        _debug(f"Candidate: ({x},{y},{ww},{hh}) aspect={aspect:.2f} score={score:.3f}")
        candidates.append((x, y, ww, hh, score))

    _debug(f"Total candidates after filtering: {len(candidates)}")
    if not candidates:
        return None, 0.0

    # Choose the best: prioritize larger area, then stronger border.
    candidates.sort(key=lambda t: (t[2] * t[3], t[4]), reverse=True)
    x, y, ww, hh, score = candidates[0]
    
    # 리사이즈한 경우 원래 좌표로 복원
    if scale < 1.0:
        x = int(x / scale)
        y = int(y / scale)
        ww = int(ww / scale)
        hh = int(hh / scale)
        _debug(f"Restored coordinates to original scale: ({x},{y},{ww},{hh})")
    
    return (int(x), int(y), int(ww), int(hh)), float(score)


def _border_score_on_rect(edges, rect: Tuple[int, int, int, int]) -> float:
    import numpy as np  # type: ignore[import-not-found]

    x, y, w, h = rect
    hh, ww = edges.shape[:2]
    x0 = max(0, x)
    y0 = max(0, y)
    x1 = min(ww - 1, x + w)
    y1 = min(hh - 1, y + h)
    if x1 <= x0 or y1 <= y0:
        return 0.0

    # Sample a 2px band around the perimeter and compute edge density.
    band = 2
    top = edges[max(0, y0 - band) : min(hh, y0 + band), x0:x1]
    bottom = edges[max(0, y1 - band) : min(hh, y1 + band), x0:x1]
    left = edges[y0:y1, max(0, x0 - band) : min(ww, x0 + band)]
    right = edges[y0:y1, max(0, x1 - band) : min(ww, x1 + band)]

    total_pixels = float(top.size + bottom.size + left.size + right.size)
    if total_pixels <= 1:
        return 0.0
    edge_pixels = float(np.count_nonzero(top) + np.count_nonzero(bottom) + np.count_nonzero(left) + np.count_nonzero(right))
    return max(0.0, min(1.0, edge_pixels / total_pixels))


def _infer_view_from_border_gap(image_path: str, rect: Tuple[int, int, int, int]) -> bool:
    """
    Heuristic for '<보기>' header when OCR misses it:
    - In many test sheets, the top border is "broken" around the centered header text.
    - We detect a strong top border on the left/right thirds with a weak middle third.
    """
    try:
        import cv2  # type: ignore[import-not-found]
        import numpy as np  # type: ignore[import-not-found]
    except Exception:
        return False

    img = cv2.imread(image_path)
    if img is None:
        return False
    h, w = img.shape[:2]
    if h < 10 or w < 10:
        return False

    x, y, ww, hh = rect
    x0 = max(0, x)
    x1 = min(w, x + ww)
    y0 = max(0, y)
    if x1 - x0 < 60:
        return False

    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    bw = cv2.adaptiveThreshold(
        gray, 255, cv2.ADAPTIVE_THRESH_MEAN_C, cv2.THRESH_BINARY_INV, 21, 10
    )
    hk = max(30, w // 25)
    h_kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (hk, 1))
    horizontal = cv2.erode(bw, h_kernel, iterations=1)
    horizontal = cv2.dilate(horizontal, h_kernel, iterations=1)

    seg_w = (x1 - x0) // 3
    if seg_w <= 0:
        return False

    # Search a small vertical window below y0 to find the best "top border" band.
    band = 3
    max_side_avg = 0.0
    best_mid_ratio = 1.0
    search_h = min(24, max(6, hh // 6))
    for dy in range(0, search_h):
        yy = min(h - 1, y0 + dy)
        top_band = horizontal[max(0, yy - band) : min(h, yy + band), x0:x1]
        if top_band.size <= 0:
            continue
        left = top_band[:, 0:seg_w]
        mid = top_band[:, seg_w : 2 * seg_w]
        right = top_band[:, 2 * seg_w : (3 * seg_w)]
        left_score = float(np.count_nonzero(left)) / max(1.0, float(left.size))
        mid_score = float(np.count_nonzero(mid)) / max(1.0, float(mid.size))
        right_score = float(np.count_nonzero(right)) / max(1.0, float(right.size))
        side_avg = (left_score + right_score) / 2.0
        if side_avg > max_side_avg:
            max_side_avg = side_avg
            best_mid_ratio = (mid_score / side_avg) if side_avg > 1e-6 else 1.0

    if max_side_avg < 0.12:
        return False
    # Middle must be significantly weaker than sides
    return best_mid_ratio < 0.45


def crop_inside_rect(image_path: str, rect: Tuple[int, int, int, int], *, inset: int = 4) -> Optional["object"]:
    """
    Return a PIL Image cropped to the inside of rect (excluding border by `inset`).
    """
    try:
        from PIL import Image  # type: ignore[import-not-found]
    except Exception:
        return None
    try:
        img = Image.open(image_path).convert("RGB")
    except Exception:
        return None
    x, y, w, h = rect
    x0 = max(0, x + inset)
    y0 = max(0, y + inset)
    x1 = min(img.width, x + w - inset)
    y1 = min(img.height, y + h - inset)
    if x1 <= x0 or y1 <= y0:
        return None
    return img.crop((x0, y0, x1, y1))


def mask_rect_on_image(image_path: str, rect: Tuple[int, int, int, int], *, pad: int = 2) -> Optional["object"]:
    """
    Return a PIL Image with the given rect area masked to white.
    """
    try:
        from PIL import Image, ImageDraw  # type: ignore[import-not-found]
    except Exception:
        return None
    try:
        img = Image.open(image_path).convert("RGB")
    except Exception:
        return None
    x, y, w, h = rect
    x0 = max(0, x - pad)
    y0 = max(0, y - pad)
    x1 = min(img.width, x + w + pad)
    y1 = min(img.height, y + h + pad)
    draw = ImageDraw.Draw(img)
    draw.rectangle([x0, y0, x1, y1], fill=(255, 255, 255))
    return img

