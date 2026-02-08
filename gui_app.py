from __future__ import annotations

import concurrent.futures
import os
import sys
import queue
import threading
import math
import tempfile
import time
from pathlib import Path

# Allow running this file directly (python gui_app.py) by ensuring the
# package parent directory is on sys.path.
if __package__ in (None, ""):
    pkg_parent = Path(__file__).resolve().parent.parent
    if str(pkg_parent) not in sys.path:
        sys.path.insert(0, str(pkg_parent))

from PySide6.QtCore import Qt, QTimer, QThread, Signal, QEvent
from PySide6.QtWidgets import (
    QApplication,
    QWidget,
    QVBoxLayout,
    QHBoxLayout,
    QLabel,
    QPushButton,
    QMessageBox,
    QFileDialog,
    QTextEdit,
    QListWidget,
    QListWidgetItem,
    QMenu,
    QStackedLayout,
    QSizePolicy,
    QProgressBar,
)
from PySide6.QtGui import QColor, QPalette, QGuiApplication, QImage
from PySide6.QtWidgets import QStyledItemDelegate, QStyle

from ai_client import AIClient, AIClientError
from hwp_controller import HwpController, HwpControllerError
from ocr_pipeline import extract_text, extract_text_from_pil_image, OcrError
from layout_detector import detect_container, crop_inside_rect, mask_rect_on_image
from script_runner import ScriptRunner, ScriptCancelled
from backend.oauth_desktop import get_stored_user, start_oauth_flow, logout_user, is_logged_in
from backend.firebase_profile import (
    refresh_user_profile_from_firebase,
    get_ai_usage,
    increment_ai_usage,
    get_remaining_usage,
    check_usage_limit,
    get_plan_limit,
    force_refresh_usage,
    PLAN_LIMITS,
)


class LoginWorker(QThread):
    """OAuth 로그인을 백그라운드에서 처리"""
    finished = Signal(bool)  # True if login successful
    
    def run(self) -> None:
        try:
            user = start_oauth_flow(timeout=300)
            self.finished.emit(user is not None and bool(user.get("uid")))
        except Exception:
            self.finished.emit(False)


class AIWorker(QThread):
    finished = Signal(object)
    error = Signal(str)
    progress = Signal(int, str)
    item_finished = Signal(int, str)

    def __init__(self, image_paths: list[str]) -> None:
        super().__init__()
        self._image_paths = image_paths

    def run(self) -> None:  # type: ignore[override]
        import sys
        def _log(msg: str) -> None:
            sys.stderr.write(f"[GUI Debug] {msg}\n")
            sys.stderr.flush()
        
        try:
            total = len(self._image_paths)
            results: list[str] = [""] * total
            _log(f"Starting AI generation for {total} images")

            def _job(idx: int, image_path: str) -> str:
                _log(f"[{idx}] Processing: {image_path}")
                # 1 image : 1 AIClient (1-to-1 mapping, safe for concurrency)
                try:
                    client = AIClient()
                except Exception as e:
                    _log(f"[{idx}] AIClient creation failed: {e}")
                    raise
                def _extract_code(text: str) -> str:
                    cleaned = (text or "").strip()
                    if cleaned.startswith("```"):
                        lines = cleaned.split("\n")[1:]
                        if lines and lines[-1].strip() == "```":
                            lines = lines[:-1]
                        return "\n".join(lines).strip()
                    return cleaned

                def _sanitize_part(script: str) -> str:
                    code = _extract_code(script)
                    if not code:
                        return ""
                    out_lines: list[str] = []
                    for line in code.splitlines():
                        s = line.strip()
                        if not s:
                            out_lines.append(line)
                            continue
                        # Prevent nested template/placeholder/box insertions inside parts.
                        if s.startswith("insert_template("):
                            continue
                        if s.startswith("focus_placeholder("):
                            continue
                        if s.startswith("insert_box(") or s == "insert_box()":
                            continue
                        if s.startswith("insert_view_box(") or s == "insert_view_box()":
                            continue
                        if s.startswith("exit_box(") or s == "exit_box()":
                            continue
                        out_lines.append(line)
                    return "\n".join(out_lines).strip()

                # 1) Full OCR (fallback context)
                _log(f"[{idx}] Starting OCR...")
                ocr_text_full = ""
                try:
                    ocr_text_full = extract_text(image_path)
                    _log(f"[{idx}] OCR done, length: {len(ocr_text_full)}")
                except Exception as e:
                    _log(f"[{idx}] OCR failed (skipping): {type(e).__name__}: {e}")
                    ocr_text_full = ""

                # 2) Detect container + split generation when possible
                _log(f"[{idx}] Detecting container...")
                det = detect_container(image_path)
                _log(f"[{idx}] Container detected: template={det.template}, rect={det.rect}")
                if det.template and det.rect:
                    _log(f"[{idx}] Building region images...")
                    # Build region images
                    try:
                        outside_img = mask_rect_on_image(image_path, det.rect)
                        _log(f"[{idx}] Outside image: {type(outside_img)}")
                    except Exception as e:
                        _log(f"[{idx}] mask_rect_on_image failed: {e}")
                        outside_img = None
                    
                    try:
                        inside_img = crop_inside_rect(image_path, det.rect)
                        _log(f"[{idx}] Inside image: {type(inside_img)}")
                    except Exception as e:
                        _log(f"[{idx}] crop_inside_rect failed: {e}")
                        inside_img = None

                    tmp_dir = Path(tempfile.gettempdir()) / "nova_ai"
                    try:
                        tmp_dir.mkdir(parents=True, exist_ok=True)
                    except Exception:
                        tmp_dir = Path.cwd()

                    outside_path = ""
                    inside_path = ""
                    try:
                        if outside_img is not None:
                            fp = tmp_dir / f"nova_ai_outside_{os.getpid()}_{idx}.png"
                            outside_img.save(fp, format="PNG")
                            outside_path = str(fp)
                    except Exception:
                        outside_path = ""
                    try:
                        if inside_img is not None:
                            fp = tmp_dir / f"nova_ai_inside_{os.getpid()}_{idx}.png"
                            inside_img.save(fp, format="PNG")
                            inside_path = str(fp)
                    except Exception:
                        inside_path = ""

                    outside_ocr = ""
                    inside_ocr = ""
                    try:
                        if outside_img is not None:
                            outside_ocr = extract_text_from_pil_image(outside_img)
                    except OcrError:
                        outside_ocr = ""
                    try:
                        if inside_img is not None:
                            inside_ocr = extract_text_from_pil_image(inside_img)
                    except OcrError:
                        inside_ocr = ""

                    _log(f"[{idx}] Calling AI for OUTSIDE content...")
                    outside_script_raw = client.generate_script_for_image(
                        outside_path or image_path,
                        description=(
                            "Type ONLY the content OUTSIDE/BEFORE the box container. "
                            "This includes the problem statement and equation. "
                            "Do NOT include ㄱ. ㄴ. ㄷ. conditions - those go INSIDE the box. "
                            "Do NOT include the answer choices (①②③④⑤)."
                        ),
                        ocr_text=outside_ocr or ocr_text_full,
                    )
                    _log(f"[{idx}] Outside AI response length: {len(outside_script_raw) if outside_script_raw else 0}")
                    
                    _log(f"[{idx}] Calling AI for INSIDE content...")
                    # For inside content, use the FULL image so AI can find the ㄱ. ㄴ. ㄷ. conditions
                    inside_script_raw = client.generate_script_for_image(
                        image_path,  # Use full image, not cropped inside
                        description=(
                            "Type ONLY the ㄱ. ㄴ. ㄷ. (or ㄱ, ㄴ, ㄷ, ㄹ) conditions that should go INSIDE the box. "
                            "These are the numbered conditions like 'ㄱ. k=0이면...' or 'ㄴ. k=3이면...' "
                            "Do NOT include the problem text before the box. "
                            "Do NOT include answer choices (①②③④⑤)."
                        ),
                        ocr_text=inside_ocr or ocr_text_full,
                    )
                    _log(f"[{idx}] Inside AI response length: {len(inside_script_raw) if inside_script_raw else 0}")
                    
                    _log(f"[{idx}] Calling AI for CHOICES content...")
                    # For choices (①②③④⑤), use the FULL image
                    choices_script_raw = client.generate_script_for_image(
                        image_path,
                        description=(
                            "Type ONLY the answer choices (①②③④⑤ or ① ㄱ  ② ㄱ, ㄴ  ③ ㄱ, ㄷ  ④ ㄴ, ㄷ  ⑤ ㄱ, ㄴ, ㄷ). "
                            "These are the multiple choice options at the bottom of the problem. "
                            "Do NOT include the problem text. "
                            "Do NOT include ㄱ. ㄴ. ㄷ. conditions."
                        ),
                        ocr_text=ocr_text_full,
                    )
                    _log(f"[{idx}] Choices AI response length: {len(choices_script_raw) if choices_script_raw else 0}")

                    outside_part = _sanitize_part(outside_script_raw or "")
                    inside_part = _sanitize_part(inside_script_raw or "")
                    choices_part = _sanitize_part(choices_script_raw or "")
                    
                    _log(f"[{idx}] Outside part preview: {outside_part[:200] if outside_part else 'EMPTY'}...")
                    _log(f"[{idx}] Inside part preview: {inside_part[:200] if inside_part else 'EMPTY'}...")
                    _log(f"[{idx}] Choices part preview: {choices_part[:200] if choices_part else 'EMPTY'}...")

                    # Template structure:
                    # 1. Insert box template
                    # 2. @@@ = placeholder for content BEFORE the box (problem text)
                    # 3. ### = placeholder for content INSIDE the box (ㄱ. ㄴ. ㄷ. conditions)
                    # 4. &&& = placeholder for content AFTER the box (answer choices ①②③④⑤)
                    combined = "\n".join(
                        [
                            f"insert_template('{det.template}')",
                            "focus_placeholder('@@@')",
                            outside_part,
                            "focus_placeholder('###')",
                            inside_part,
                            "focus_placeholder('&&&')",
                            choices_part,
                        ]
                    ).strip()
                    _log(f"[{idx}] Combined script length: {len(combined)}")
                    return combined

                if det.template and not det.rect:
                    # Header text detected but rectangle not confidently found:
                    # enforce template/placeholder workflow and let the model separate.
                    _log(f"[{idx}] Template detected (no rect): {det.template}")
                    script_raw = client.generate_script_for_image(image_path, ocr_text=ocr_text_full) or ""
                    _log(f"[{idx}] AI response length: {len(script_raw)}")
                    script_body = _sanitize_part(script_raw)
                    combined = "\n".join(
                        [
                            f"insert_template('{det.template}')",
                            "focus_placeholder('@@@')",
                            script_body,
                            "focus_placeholder('###')",
                            "",
                        ]
                    ).strip()
                    return combined

                # No container detected: default behavior
                _log(f"[{idx}] No container detected, calling AI...")
                raw_result = client.generate_script_for_image(image_path, ocr_text=ocr_text_full) or ""
                _log(f"[{idx}] AI response length: {len(raw_result)}")
                if not raw_result.strip():
                    _log(f"[{idx}] WARNING: Empty AI response!")
                return _extract_code(raw_result)

            # If you need to cap concurrency (rate limiting), set NOVA_AI_MAX_WORKERS.
            max_workers_env = os.getenv("NOVA_AI_MAX_WORKERS")
            max_workers = total
            if max_workers_env:
                try:
                    max_workers = max(1, min(total, int(max_workers_env)))
                except Exception:
                    max_workers = total

            with concurrent.futures.ThreadPoolExecutor(max_workers=max_workers) as ex:
                future_to_idx: dict[concurrent.futures.Future[str], int] = {}
                for idx, image_path in enumerate(self._image_paths):
                    self.progress.emit(idx, "생성중")
                    future_to_idx[ex.submit(_job, idx, image_path)] = idx

                for fut in concurrent.futures.as_completed(future_to_idx):
                    idx = future_to_idx[fut]
                    try:
                        text = fut.result() or ""
                        results[idx] = text
                        if text.strip():
                            self.progress.emit(idx, "코드 생성 완료")
                        else:
                            self.progress.emit(idx, "오류(빈 결과)")
                        # Notify UI for incremental typing / preview.
                        self.item_finished.emit(idx, text)
                    except Exception as exc:
                        results[idx] = ""
                        self.progress.emit(idx, f"오류: {exc}")
                        self.item_finished.emit(idx, "")
            self.finished.emit(results)
        except Exception as exc:
            self.error.emit(str(exc))


class NovaAILiteWindow(QWidget):
    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle("Nova AI Lite")
        self.setMinimumWidth(520)
        self.setAcceptDrops(True)

        # Profile state (populated from get_stored_user() and Firebase)
        self.profile_uid: str | None = None
        self.profile_display_name: str = "사용자"
        self.profile_plan: str = "Free"
        self.profile_avatar_url: str | None = None
        self._login_worker: LoginWorker | None = None

        layout = QVBoxLayout(self)

        self.selected_images: list[str] = []
        self.generated_code: str = ""
        self.generated_codes: list[str] = []
        self._generated_codes_by_index: list[str] = []
        self._gen_statuses: list[str] = []
        self._ai_worker: AIWorker | None = None
        self._typed_indexes: set[int] = set()
        self._next_auto_type_index: int = 0
        self._auto_type_has_inserted_any: bool = False
        self._auto_type_pending_idx: int | None = None
        self._skipped_indexes: set[int] = set()
        self._typing_worker: "TypingWorker | None" = None
        self.filename_label = QLabel("감지된 파일: (없음)")
        self.filename_label.setTextInteractionFlags(Qt.TextInteractionFlag.TextSelectableByMouse)
        self.typing_status_label = QLabel("")
        self.typing_status_label.setStyleSheet("color: #bbb;")
        self.order_title = QLabel("타이핑 순서: (없음)")
        self.order_title.setTextInteractionFlags(Qt.TextInteractionFlag.TextSelectableByMouse)
        self.order_list = OrderListWidget()
        self.order_list.setMinimumHeight(260)
        self.order_list.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
        self.order_list.setSelectionMode(QListWidget.SelectionMode.SingleSelection)
        self.order_list.setDragDropMode(QListWidget.DragDropMode.InternalMove)
        self.order_list.setDefaultDropAction(Qt.DropAction.MoveAction)
        self.order_list.setDragEnabled(True)
        self.order_list.setAcceptDrops(True)
        self.order_list.setDropIndicatorShown(True)
        self._order_delegate = OrderListDelegate(self.order_list)
        self.order_list.setItemDelegate(self._order_delegate)
        self.order_list.itemClicked.connect(self._on_order_item_clicked)
        self.order_list.model().rowsMoved.connect(self._on_order_rows_moved)
        self.order_list.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.order_list.customContextMenuRequested.connect(self._on_order_context_menu)
        self.order_list.filesDropped.connect(self._on_files_dropped)

        self.btn_ai_type = QPushButton("AI 타이핑")

        self.code_view = QTextEdit()
        self.code_view.setReadOnly(False)
        self.code_view.setFixedHeight(200)
        self._generated_code_label = QLabel("생성된 코드")
        self._code_type_btn = QPushButton("코드 타이핑")
        self._code_type_btn.setEnabled(False)
        self._generated_container = QWidget()
        gen_layout = QVBoxLayout(self._generated_container)
        gen_layout.setContentsMargins(0, 0, 0, 0)
        gen_layout.setSpacing(8)
        gen_header = QHBoxLayout()
        gen_header.addWidget(self._generated_code_label)
        gen_header.addStretch(1)
        gen_header.addWidget(self._code_type_btn)
        gen_layout.addLayout(gen_header)
        gen_layout.addWidget(self.code_view)
        # Hidden until user presses the typing-order button (AI 타이핑).
        self._generated_container.setVisible(False)

        # Typing order container (status + list + bottom row)
        order_container = QWidget()
        order_layout = QVBoxLayout(order_container)
        order_layout.setContentsMargins(0, 0, 0, 0)
        order_layout.setSpacing(8)
        order_status_row = QHBoxLayout()
        order_status_row.addWidget(self.typing_status_label)
        order_status_row.addStretch(1)
        order_layout.addLayout(order_status_row)
        list_stack_container = QWidget()
        list_stack_container.setMinimumHeight(260)
        list_stack_container.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
        list_stack = QStackedLayout(list_stack_container)
        self._empty_placeholder = DropPlaceholder()
        self._empty_placeholder.setMinimumHeight(260)
        self._empty_placeholder.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
        self._empty_placeholder.clicked.connect(self.on_upload_image)
        self._empty_placeholder.filesDropped.connect(self._on_files_dropped)
        list_stack.addWidget(self._empty_placeholder)
        list_stack.addWidget(self.order_list)
        order_layout.addWidget(list_stack_container, 1)
        self._order_list_stack = list_stack
        layout.addSpacing(20)
        
        # 상단 사용자 정보 및 로그인 버튼
        user_row = QHBoxLayout()
        self._user_label = QLabel("")
        self._user_label.setStyleSheet("color: #aaa; font-size: 12px;")
        self._token_progress = QProgressBar()
        self._token_progress.setFixedWidth(120)
        self._token_progress.setFixedHeight(16)
        self._token_progress.setTextVisible(True)
        self._token_progress.setFormat("%v/%m")
        self._token_progress.setStyleSheet("""
            QProgressBar {
                border: 1px solid #555;
                border-radius: 4px;
                background-color: #333;
                text-align: center;
                color: #fff;
                font-size: 10px;
            }
            QProgressBar::chunk {
                background-color: #4CAF50;
                border-radius: 3px;
            }
        """)
        self._login_btn = QPushButton("로그인")
        self._login_btn.setFixedWidth(80)
        self._logout_btn = QPushButton("로그아웃")
        self._logout_btn.setFixedWidth(80)
        self._logout_btn.setVisible(False)
        
        user_row.addWidget(self._user_label)
        user_row.addWidget(self._token_progress)
        user_row.addStretch(1)
        user_row.addWidget(self._login_btn)
        user_row.addWidget(self._logout_btn)
        layout.addLayout(user_row)
        layout.addSpacing(10)
        
        self._login_btn.clicked.connect(self._on_login_clicked)
        self._logout_btn.clicked.connect(self._on_logout_clicked)
        
        title = QLabel("Nova AI Lite")
        title.setAlignment(Qt.AlignmentFlag.AlignHCenter | Qt.AlignmentFlag.AlignVCenter)
        title.setStyleSheet("font-size: 26px; font-weight: 700;")
        layout.addWidget(title)
        subtitle = QLabel("가장 빠른 한글 AI 입력기")
        subtitle.setAlignment(Qt.AlignmentFlag.AlignHCenter | Qt.AlignmentFlag.AlignVCenter)
        subtitle.setStyleSheet("font-size: 13px; color: #bbb;")
        layout.addWidget(subtitle)
        layout.addSpacing(20)
        top_action_row = QHBoxLayout()
        top_action_row.addWidget(self.filename_label)
        top_action_row.addStretch(1)
        top_action_row.addWidget(self.btn_ai_type)
        layout.addLayout(top_action_row)
        layout.addSpacing(12)
        layout.addWidget(order_container)
        layout.addWidget(self._generated_container)

        self.btn_ai_type.clicked.connect(self.on_ai_type_run)
        self._code_type_btn.clicked.connect(self._on_code_type_clicked)
        self.code_view.textChanged.connect(self._on_code_view_changed)

        self._timer = QTimer(self)
        self._timer.setInterval(1000)
        self._timer.timeout.connect(self.update_filename)
        self._timer.start()
        self.update_filename()
        self._auto_type_after_ai = False
        self._current_code_index = -1
        self._current_code_path: str | None = None
        self._code_view_updating = False

        # Animate "생성중" status in the list.
        self._status_anim_timer = QTimer(self)
        self._status_anim_timer.setInterval(50)
        self._status_anim_timer.timeout.connect(self._tick_status_animation)
        self._status_anim_timer.start()

        # Capture ESC globally to stop typing even during long operations.
        app = QApplication.instance()
        if app is not None:
            app.installEventFilter(self)
        
        # 초기 사용자 상태 업데이트 (로컬 캐시에서)
        self._load_stored_user()
        self._update_user_status()
        
        # Firebase에서 최신 프로필 동기화 (백그라운드)
        QTimer.singleShot(500, self._refresh_profile_from_firebase)

    def _load_stored_user(self) -> None:
        """로컬 캐시에서 저장된 사용자 정보 로드"""
        user = get_stored_user()
        if user and user.get("uid"):
            self.profile_uid = user.get("uid")
            self.profile_display_name = user.get("name") or "사용자"
            self.profile_plan = user.get("tier") or "Free"
            self.profile_avatar_url = user.get("photo_url")
        else:
            self.profile_uid = None
            self.profile_display_name = "사용자"
            self.profile_plan = "Free"
            self.profile_avatar_url = None

    def _refresh_profile_from_firebase(self) -> None:
        """Firebase에서 최신 프로필과 사용량 동기화"""
        if not self.profile_uid:
            return
        try:
            # Force refresh usage from Firebase (bypass cache on startup)
            force_refresh_usage()
            
            profile = refresh_user_profile_from_firebase()
            if profile:
                self.profile_plan = profile.get("tier") or self.profile_plan
                self.profile_display_name = profile.get("display_name") or self.profile_display_name
                self._update_user_status()
        except Exception:
            pass

    def _update_user_status(self) -> None:
        """사용자 로그인 상태 및 사용량 정보 업데이트"""
        if self.profile_uid:
            # 로그인 상태
            tier = self.profile_plan or "Free"
            tier_label = {"Free": "무료", "free": "무료", "Standard": "Standard", "Plus": "Plus", "Pro": "Pro"}.get(tier, tier)
            self._user_label.setText(f"{self.profile_display_name} ({tier_label})")
            
            # 사용량 프로그레스 바 업데이트
            usage = get_ai_usage(self.profile_uid)
            limit = get_plan_limit(tier)
            self._token_progress.setMaximum(limit)
            self._token_progress.setValue(usage)
            
            # 남은 횟수 또는 한도 초과 표시
            remaining = limit - usage
            if remaining <= 0:
                self._token_progress.setFormat(f"한도 초과! 업그레이드 필요")
            elif remaining <= 5:
                self._token_progress.setFormat(f"{usage}/{limit} (남은 횟수: {remaining})")
            else:
                self._token_progress.setFormat(f"{usage}/{limit}")
            self._token_progress.setVisible(True)
            
            # 사용량에 따라 색상 변경
            usage_ratio = usage / limit if limit > 0 else 0
            if usage_ratio >= 1.0:
                chunk_color = "#d32f2f"  # 진한 빨강 (한도 초과)
            elif usage_ratio >= 0.9:
                chunk_color = "#f44336"  # 빨강
            elif usage_ratio >= 0.7:
                chunk_color = "#ff9800"  # 주황
            else:
                chunk_color = "#4CAF50"  # 초록
            self._token_progress.setStyleSheet(f"""
                QProgressBar {{
                    border: 1px solid #555;
                    border-radius: 4px;
                    background-color: #333;
                    text-align: center;
                    color: #fff;
                    font-size: 10px;
                }}
                QProgressBar::chunk {{
                    background-color: {chunk_color};
                    border-radius: 3px;
                }}
            """)
            
            self._login_btn.setVisible(False)
            self._logout_btn.setVisible(True)
        else:
            # 비로그인 상태
            self._user_label.setText("로그인하면 사용량을 추적할 수 있습니다")
            self._token_progress.setVisible(False)
            self._login_btn.setVisible(True)
            self._logout_btn.setVisible(False)

    def _on_login_clicked(self) -> None:
        """로그인 버튼 클릭 - 브라우저로 OAuth 로그인"""
        self._login_btn.setEnabled(False)
        self._login_btn.setText("로그인 중...")
        
        # 백그라운드에서 OAuth 플로우 시작
        self._login_worker = LoginWorker()
        self._login_worker.finished.connect(self._on_login_finished)
        self._login_worker.start()

    def _on_login_finished(self, success: bool) -> None:
        """OAuth 로그인 완료"""
        self._login_btn.setEnabled(True)
        self._login_btn.setText("로그인")
        
        if success:
            self._load_stored_user()
            self._update_user_status()
            QMessageBox.information(self, "로그인 성공", f"환영합니다, {self.profile_display_name}!")
            # Firebase에서 최신 정보 가져오기
            QTimer.singleShot(100, self._refresh_profile_from_firebase)
        else:
            QMessageBox.warning(self, "로그인 실패", "로그인이 취소되었거나 실패했습니다.")

    def _on_logout_clicked(self) -> None:
        """로그아웃 버튼 클릭"""
        reply = QMessageBox.question(
            self,
            "로그아웃",
            "로그아웃하시겠습니까?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No
        )
        if reply == QMessageBox.StandardButton.Yes:
            logout_user()
            self.profile_uid = None
            self.profile_display_name = "사용자"
            self.profile_plan = "Free"
            self.profile_avatar_url = None
            self._update_user_status()

    def _tick_status_animation(self) -> None:
        try:
            self._order_delegate.advance()
            self.order_list.viewport().update()
        except Exception:
            pass

    def _connect(self) -> HwpController:
        controller = HwpController()
        controller.connect()
        return controller

    def update_filename(self) -> None:
        try:
            filename = HwpController.get_current_filename()
            if filename:
                self.filename_label.setText(f"감지된 파일: {filename}")
            else:
                self.filename_label.setText("감지된 파일: (없음)")
        except Exception:
            self.filename_label.setText("감지된 파일: (오류)")

    def on_upload_image(self) -> None:
        file_paths, _ = QFileDialog.getOpenFileNames(
            self,
            "사진 선택",
            "",
            "Images (*.png *.jpg *.jpeg *.bmp *.gif *.webp *.pdf);;All Files (*)",
        )
        if file_paths:
            self._set_selected_images(file_paths)

    def _on_files_dropped(self, file_paths: list[str]) -> None:
        if file_paths:
            self._set_selected_images(file_paths)

    def on_ai_run(self) -> None:
        self._start_ai_run(auto_type=False)

    def on_ai_type_run(self) -> None:
        self._set_typing_status("기다리는 중")
        self._generated_container.setVisible(True)
        self._start_ai_run(auto_type=True)

    def _start_ai_run(self, auto_type: bool) -> None:
        if not self.selected_images:
            QMessageBox.warning(self, "안내", "먼저 사진을 업로드하세요.")
            return
        if self._ai_worker and self._ai_worker.isRunning():
            return
        # Prevent reordering/removal while generation is running.
        self._set_order_editable(False)
        self._auto_type_after_ai = auto_type
        self.generated_code = ""
        self.generated_codes = []
        self._generated_codes_by_index = [""] * len(self.selected_images)
        self._gen_statuses = ["대기"] * len(self.selected_images)
        self._typed_indexes = set()
        self._next_auto_type_index = 0
        self._auto_type_has_inserted_any = False
        self._skipped_indexes = set()
        self._render_order_list()
        self.code_view.setPlainText("")
        self._ai_worker = AIWorker(self.selected_images)
        self._ai_worker.finished.connect(self._on_ai_finished)
        self._ai_worker.error.connect(self._on_ai_error)
        self._ai_worker.progress.connect(self._on_ai_progress)
        self._ai_worker.item_finished.connect(self._on_ai_item_finished)
        self._ai_worker.start()

    def on_type_run(self) -> None:
        if not self.generated_codes and not self.generated_code.strip():
            QMessageBox.warning(self, "안내", "먼저 AI 실행을 눌러 코드를 생성하세요.")
            return
        script = self._build_typing_script()
        if not script.strip():
            QMessageBox.warning(self, "안내", "실행할 코드가 없습니다.")
            return
        self._ensure_typing_worker()
        self._typing_worker.enqueue(-1, script)

    def _render_order_list(self) -> None:
        self.order_list.clear()
        if not self.selected_images:
            self.order_title.setText("")
            return
        self.order_title.setText("타이핑 순서:")
        for idx, path in enumerate(self.selected_images):
            name = os.path.basename(path)
            status = self._gen_statuses[idx] if idx < len(self._gen_statuses) else "대기"
            item = QListWidgetItem(f"{idx + 1}. {name} - {status}")
            item.setData(Qt.ItemDataRole.UserRole, path)
            self.order_list.addItem(item)

    def _on_ai_progress(self, idx: int, status: str) -> None:
        if idx < 0:
            return
        if idx >= len(self._gen_statuses):
            self._gen_statuses.extend(["대기"] * (idx + 1 - len(self._gen_statuses)))
        self._gen_statuses[idx] = status
        self._render_order_list()

    def _run_typing(self) -> None:
        # Deprecated: typing now runs in a worker thread to allow ESC cancellation.
        self.on_type_run()

    def _run_typing_script(self, script: str) -> None:
        # Deprecated: typing now runs in a worker thread to allow ESC cancellation.
        if not script.strip():
            return
        self._ensure_typing_worker()
        self._typing_worker.enqueue(-1, script)

    def _build_typing_script(self) -> str:
        if self.generated_codes:
            cleaned = [code.strip() for code in self.generated_codes if code.strip()]
            separator = "\ninsert_paragraph()\n" * 4
            return separator.join(cleaned)
        return self.generated_code

    def _on_ai_item_finished(self, idx: int, text: str) -> None:
        """Called when a single image's code generation finishes (success or fail)."""
        if idx < 0:
            return
        if idx >= len(self._generated_codes_by_index):
            # Defensive: keep arrays consistent.
            self._generated_codes_by_index.extend([""] * (idx + 1 - len(self._generated_codes_by_index)))
        self._generated_codes_by_index[idx] = (text or "").strip()
        if idx < len(self.generated_codes):
            self.generated_codes[idx] = self._generated_codes_by_index[idx]
        if idx == self._current_code_index:
            if not self.code_view.hasFocus() or not self.code_view.toPlainText().strip():
                self._set_code_view_text(self._generated_codes_by_index[idx])
            self._update_code_type_button_state()
        # Auto-typing: type incrementally in order as soon as possible.
        if self._auto_type_after_ai:
            self._try_auto_type()

    def _try_auto_type(self) -> None:
        """Type completed items in order while generation continues."""
        if not self._auto_type_after_ai:
            return
        total = len(self.selected_images)
        if total <= 0:
            return
        if self._auto_type_pending_idx is not None:
            return

        # Type sequentially (1 -> 2 -> 3 ...) only when each is ready.
        while self._next_auto_type_index < total and self._auto_type_pending_idx is None:
            idx = self._next_auto_type_index

            status = self._gen_statuses[idx] if idx < len(self._gen_statuses) else "대기"
            # Not ready yet (still generating or not started).
            if status in ("대기", "생성중"):
                self._set_typing_status("기다리는 중")
                return

            code = (self._generated_codes_by_index[idx] or "").strip()
            # If generation failed/empty, skip and continue to the next item.
            if not code:
                if idx not in self._skipped_indexes:
                    self._skipped_indexes.add(idx)
                    if idx < len(self._gen_statuses):
                        self._gen_statuses[idx] = "생성 실패(건너뜀)"
                    self._render_order_list()
                self._next_auto_type_index += 1
                continue

            separator = ""
            if self._auto_type_has_inserted_any:
                separator = "insert_paragraph()\n" * 4
            script = f"{separator}{code}\n"

            self._ensure_typing_worker()
            self._auto_type_pending_idx = idx
            if idx < len(self._gen_statuses):
                self._gen_statuses[idx] = "타이핑 대기"
            self._render_order_list()
            self._set_typing_status("타이핑 중")
            self._typing_worker.enqueue(idx, script)
            return

    def _on_ai_finished(self, results: object) -> None:
        if not isinstance(results, list):
            results = [results]
        raw_codes = [str(item or "").strip() for item in results]

        total = len(self.selected_images)
        if len(raw_codes) < total:
            raw_codes.extend([""] * (total - len(raw_codes)))
        raw_codes = raw_codes[:total]

        self._generated_codes_by_index = raw_codes
        ok_count = sum(1 for c in raw_codes if c.strip())
        all_ok = (total > 0) and (ok_count == total)

        self._render_order_list()

        # Store results for manual typing as well.
        self.generated_codes = raw_codes
        self.generated_code = raw_codes[0] if total == 1 else ""

        # Ensure any remaining ready items are typed (in case signals arrived late).
        if self._auto_type_after_ai:
            self._try_auto_type()
            if self._next_auto_type_index >= total and self._auto_type_pending_idx is None:
                self._auto_type_after_ai = False
                self._set_typing_status("")

        if not all_ok:
            failed_indexes = [i + 1 for i, code in enumerate(raw_codes) if not code.strip()]
            QMessageBox.warning(
                self,
                "안내",
                f"일부 문제에서 코드 생성이 실패했습니다: {failed_indexes}\n"
                "실패한 항목은 다시 시도하거나, 성공한 항목만 수동으로 타이핑할 수 있습니다.",
            )
        self._set_order_editable(True)
        self._update_code_type_button_state()
        
        # 토큰 사용량 업데이트
        self._update_user_status()

    def _on_ai_error(self, message: str) -> None:
        self._render_order_list()
        QMessageBox.critical(self, "AI 오류", message)
        self._auto_type_after_ai = False
        self._set_order_editable(True)
        
        # 토큰 사용량 업데이트
        self._update_user_status()

    def _ensure_typing_worker(self) -> None:
        if self._typing_worker and self._typing_worker.isRunning():
            return
        self._typing_worker = TypingWorker()
        self._typing_worker.item_started.connect(self._on_typing_item_started)
        self._typing_worker.item_finished.connect(self._on_typing_item_finished)
        self._typing_worker.cancelled.connect(self._on_typing_cancelled)
        self._typing_worker.error.connect(self._on_typing_error)
        self._typing_worker.start()

    def _on_typing_item_started(self, idx: int) -> None:
        if idx >= 0 and idx < len(self._gen_statuses):
            self._gen_statuses[idx] = "타이핑중"
            self._render_order_list()
        self._set_typing_status("타이핑 중")

    def _on_typing_item_finished(self, idx: int) -> None:
        if idx >= 0 and idx < len(self._gen_statuses):
            self._gen_statuses[idx] = "타이핑 완료"
            self._render_order_list()
        if idx >= 0 and self._auto_type_pending_idx == idx:
            self._auto_type_pending_idx = None
            self._auto_type_has_inserted_any = True
            self._next_auto_type_index = idx + 1
            if self._auto_type_after_ai:
                self._try_auto_type()
            if self._next_auto_type_index >= len(self.selected_images):
                self._auto_type_after_ai = False
                self._set_typing_status("")

    def _on_typing_cancelled(self) -> None:
        # Stop auto-type chain, keep generated code for manual re-run.
        self._auto_type_after_ai = False
        self._auto_type_pending_idx = None
        self._set_typing_status("")
        QMessageBox.information(self, "안내", "타이핑이 중단되었습니다.")

    def _on_typing_error(self, message: str) -> None:
        self._auto_type_after_ai = False
        self._auto_type_pending_idx = None
        self._set_typing_status("")
        QMessageBox.critical(self, "타이핑 오류", message)

    def _cancel_typing(self) -> None:
        if self._typing_worker and self._typing_worker.isRunning():
            self._typing_worker.cancel()

    def _save_clipboard_image(self) -> str:
        clipboard = QGuiApplication.clipboard()
        if clipboard is None:
            return ""
        img = clipboard.image()
        if img is None or img.isNull():
            return ""
        tmp_dir = Path(tempfile.gettempdir()) / "nova_ai"
        try:
            tmp_dir.mkdir(parents=True, exist_ok=True)
        except Exception:
            tmp_dir = Path.cwd()
        file_name = f"nova_ai_clip_{os.getpid()}_{time.time_ns()}.png"
        file_path = tmp_dir / file_name
        try:
            saved = img.save(str(file_path), "PNG")
        except Exception:
            saved = False
        return str(file_path) if saved else ""

    def _try_paste_image(self) -> bool:
        clipboard = QGuiApplication.clipboard()
        if clipboard is None:
            return False
        mime = clipboard.mimeData()
        if mime is None or not mime.hasImage():
            return False
        path = self._save_clipboard_image()
        if not path:
            return False
        new_paths = list(self.selected_images)
        new_paths.append(path)
        self._set_selected_images(new_paths)
        return True

    def eventFilter(self, obj, event):  # type: ignore[override]
        try:
            if event.type() == QEvent.Type.KeyPress and event.key() == Qt.Key.Key_Escape:
                self._cancel_typing()
                return True
            if (
                event.type() == QEvent.Type.KeyPress
                and event.key() == Qt.Key.Key_V
                and event.modifiers() & Qt.KeyboardModifier.ControlModifier
            ):
                if self._try_paste_image():
                    return True
        except Exception:
            pass
        return super().eventFilter(obj, event)

    def _on_order_item_clicked(self, item: QListWidgetItem) -> None:
        idx = self.order_list.row(item)
        if idx < 0 or idx >= len(self._generated_codes_by_index):
            self._current_code_index = -1
            self._current_code_path = None
            self._set_code_view_text("")
            self._update_code_type_button_state()
            return
        self._current_code_index = idx
        path = item.data(Qt.ItemDataRole.UserRole)
        self._current_code_path = path if isinstance(path, str) else None
        code = self._generated_codes_by_index[idx] or ""
        self._set_code_view_text(code)
        self._update_code_type_button_state()

    def _on_order_rows_moved(self, *args) -> None:
        # Rebuild selected_images order based on list widget items.
        if not self.selected_images:
            return
        new_paths: list[str] = []
        for i in range(self.order_list.count()):
            item = self.order_list.item(i)
            path = item.data(Qt.ItemDataRole.UserRole)
            if isinstance(path, str) and path:
                new_paths.append(path)
        if not new_paths or len(new_paths) != len(self.selected_images):
            return
        old_index_by_path = {p: i for i, p in enumerate(self.selected_images)}
        self.selected_images = new_paths
        if self._generated_codes_by_index:
            self._generated_codes_by_index = [
                self._generated_codes_by_index[old_index_by_path[p]]
                for p in new_paths
                if p in old_index_by_path
            ]
        if self._gen_statuses:
            self._gen_statuses = [
                self._gen_statuses[old_index_by_path[p]]
                for p in new_paths
                if p in old_index_by_path
            ]
        self._render_order_list()
        if self._current_code_path and self._current_code_path in new_paths:
            self._current_code_index = new_paths.index(self._current_code_path)
            code = self._generated_codes_by_index[self._current_code_index] or ""
            self._set_code_view_text(code)
        else:
            self._current_code_index = -1
            self._current_code_path = None
            self._set_code_view_text("")
        self._update_code_type_button_state()

    def _on_order_context_menu(self, pos) -> None:
        if not self._is_order_editable():
            return
        item = self.order_list.itemAt(pos)
        if item is None:
            return
        menu = QMenu(self)
        remove_action = menu.addAction("항목 제거")
        action = menu.exec(self.order_list.mapToGlobal(pos))
        if action == remove_action:
            self._remove_order_item(item)

    def _remove_order_item(self, item: QListWidgetItem) -> None:
        if not self._is_order_editable():
            QMessageBox.information(self, "안내", "생성 중에는 항목을 변경할 수 없습니다.")
            return
        idx = self.order_list.row(item)
        if idx < 0 or idx >= len(self.selected_images):
            return
        self.selected_images.pop(idx)
        if idx < len(self._generated_codes_by_index):
            self._generated_codes_by_index.pop(idx)
        if idx < len(self._gen_statuses):
            self._gen_statuses.pop(idx)
        self._render_order_list()

    def _set_selected_images(self, file_paths: list[str]) -> None:
        self.selected_images = [path for path in file_paths if path]
        if not self.selected_images:
            self.order_list.clear()
            self._gen_statuses = []
            self._generated_codes_by_index = []
            self._current_code_index = -1
            self._current_code_path = None
            self._set_code_view_text("")
            self._update_code_type_button_state()
            self._update_order_list_visibility()
            return
        order_lines = [
            f"{idx + 1}. {os.path.basename(path)}"
            for idx, path in enumerate(self.selected_images)
        ]
        self.order_title.setText("타이핑 순서:\n" + "\n".join(order_lines))
        self._generated_codes_by_index = [""] * len(self.selected_images)
        self._gen_statuses = ["대기"] * len(self.selected_images)
        self._render_order_list()
        self._current_code_index = -1
        self._current_code_path = None
        self._set_code_view_text("")
        self._update_code_type_button_state()
        self._update_order_list_visibility()

    def _set_order_editable(self, enabled: bool) -> None:
        if enabled:
            self.order_list.setDragDropMode(QListWidget.DragDropMode.InternalMove)
            self.order_list.setDragEnabled(True)
            self.order_list.setAcceptDrops(True)
            self.order_list.setDropIndicatorShown(True)
        else:
            self.order_list.setDragDropMode(QListWidget.DragDropMode.NoDragDrop)
            self.order_list.setDragEnabled(False)
            self.order_list.setAcceptDrops(False)
            self.order_list.setDropIndicatorShown(False)

    def _is_order_editable(self) -> bool:
        return self.order_list.dragDropMode() != QListWidget.DragDropMode.NoDragDrop

    def _update_order_list_visibility(self) -> None:
        if not self.selected_images:
            self._order_list_stack.setCurrentWidget(self._empty_placeholder)
        else:
            self._order_list_stack.setCurrentWidget(self.order_list)

    def _set_typing_status(self, text: str) -> None:
        self.typing_status_label.setText(text)

    def _set_code_view_text(self, text: str) -> None:
        self._code_view_updating = True
        try:
            self.code_view.setPlainText(text or "")
        finally:
            self._code_view_updating = False

    def _update_code_type_button_state(self) -> None:
        idx = self._current_code_index
        if idx < 0 or idx >= len(self._generated_codes_by_index):
            self._code_type_btn.setEnabled(False)
            return
        code = self._generated_codes_by_index[idx] or ""
        self._code_type_btn.setEnabled(bool(code.strip()))

    def _sync_current_code_from_view(self) -> None:
        idx = self._current_code_index
        if idx < 0 or idx >= len(self._generated_codes_by_index):
            return
        text = self.code_view.toPlainText()
        self._generated_codes_by_index[idx] = text
        if idx < len(self.generated_codes):
            self.generated_codes[idx] = text
        self._update_code_type_button_state()

    def _on_code_view_changed(self) -> None:
        if self._code_view_updating:
            return
        self._sync_current_code_from_view()

    def _on_code_type_clicked(self) -> None:
        idx = self._current_code_index
        if idx < 0 or idx >= len(self._generated_codes_by_index):
            QMessageBox.warning(self, "안내", "먼저 항목을 선택하세요.")
            return
        self._sync_current_code_from_view()
        script = (self._generated_codes_by_index[idx] or "").strip()
        if not script:
            QMessageBox.warning(self, "안내", "실행할 코드가 없습니다.")
            return
        # Ensure only the selected item is typed.
        self._auto_type_after_ai = False
        self._auto_type_pending_idx = None
        self._set_typing_status("타이핑 중")
        self._ensure_typing_worker()
        self._typing_worker.enqueue(idx, f"{script}\n")

    def dragEnterEvent(self, event) -> None:  # type: ignore[override]
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
        else:
            event.ignore()

    def dropEvent(self, event) -> None:  # type: ignore[override]
        urls = event.mimeData().urls()
        if not urls:
            return
        file_paths = [url.toLocalFile() for url in urls if url.toLocalFile()]
        if file_paths:
            self._set_selected_images(file_paths)


def _is_rpc_unavailable_message(message: str) -> bool:
    return (
        "RPC 서버를 사용할 수 없습니다" in message
        or "RPC server is unavailable" in message
        or "0x800706BA" in message
        or "-2147023174" in message
    )


class TypingWorker(QThread):
    item_started = Signal(int)
    item_finished = Signal(int)
    cancelled = Signal()
    error = Signal(str)

    def __init__(self) -> None:
        super().__init__()
        self._q: "queue.Queue[tuple[int, str]]" = queue.Queue()
        self._cancel = threading.Event()

    def enqueue(self, idx: int, script: str) -> None:
        if not script.strip():
            return
        self._q.put((idx, script))

    def cancel(self) -> None:
        self._cancel.set()
        # best-effort drain
        try:
            while True:
                self._q.get_nowait()
        except Exception:
            pass

    def run(self) -> None:  # type: ignore[override]
        # COM init (best-effort) to safely control HWP from this thread.
        pythoncom = None
        try:
            import pythoncom  # type: ignore
        except Exception:
            pythoncom = None
        if pythoncom is not None:
            try:
                pythoncom.CoInitialize()
            except Exception:
                pass

        controller: HwpController | None = None
        runner: ScriptRunner | None = None
        try:
            while True:
                if self._cancel.is_set():
                    self.cancelled.emit()
                    return
                try:
                    idx, script = self._q.get(timeout=0.1)
                except Exception:
                    continue

                if self._cancel.is_set():
                    self.cancelled.emit()
                    return

                if controller is None:
                    controller = HwpController()
                    controller.connect()
                    runner = ScriptRunner(controller)

                self.item_started.emit(idx)
                try:
                    assert runner is not None
                    runner.run(script, cancel_check=self._cancel.is_set)
                except ScriptCancelled:
                    self.cancelled.emit()
                    return
                except HwpControllerError as exc:
                    msg = str(exc)
                    if _is_rpc_unavailable_message(msg):
                        try:
                            controller = HwpController()
                            controller.connect()
                            runner = ScriptRunner(controller)
                            runner.run(script, cancel_check=self._cancel.is_set)
                        except Exception as retry_exc:
                            self.error.emit(str(retry_exc))
                            return
                    else:
                        self.error.emit(msg)
                        return
                except Exception as exc:
                    msg = str(exc)
                    if _is_rpc_unavailable_message(msg):
                        try:
                            controller = HwpController()
                            controller.connect()
                            runner = ScriptRunner(controller)
                            runner.run(script, cancel_check=self._cancel.is_set)
                        except Exception as retry_exc:
                            self.error.emit(str(retry_exc))
                            return
                    else:
                        self.error.emit(msg)
                        return
                self.item_finished.emit(idx)
        finally:
            if pythoncom is not None:
                try:
                    pythoncom.CoUninitialize()
                except Exception:
                    pass


class OrderListWidget(QListWidget):
    filesDropped = Signal(list)

    def dragEnterEvent(self, event) -> None:  # type: ignore[override]
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
            return
        super().dragEnterEvent(event)

    def dragMoveEvent(self, event) -> None:  # type: ignore[override]
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
            return
        super().dragMoveEvent(event)

    def dropEvent(self, event) -> None:  # type: ignore[override]
        if event.mimeData().hasUrls():
            paths = [url.toLocalFile() for url in event.mimeData().urls() if url.toLocalFile()]
            if paths:
                self.filesDropped.emit(paths)
                event.acceptProposedAction()
                return
        super().dropEvent(event)


class DropPlaceholder(QWidget):
    clicked = Signal()
    filesDropped = Signal(list)

    def __init__(self, parent=None) -> None:
        super().__init__(parent)
        self.setAcceptDrops(True)
        self.setStyleSheet("background-color: #2d2d2d; border-radius: 8px;")
        self.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
        layout = QVBoxLayout(self)
        layout.setContentsMargins(12, 12, 12, 12)
        layout.setSpacing(6)
        text_label = QLabel("사진을 넣으려면 드래그앤드롭하세요")
        text_label.setAlignment(Qt.AlignmentFlag.AlignHCenter | Qt.AlignmentFlag.AlignVCenter)
        text_label.setStyleSheet("color: #999; background-color: transparent;")
        layout.addStretch(1)
        layout.addWidget(text_label)
        layout.addStretch(1)

    def mousePressEvent(self, event) -> None:  # type: ignore[override]
        if event.button() == Qt.MouseButton.LeftButton:
            self.clicked.emit()
            return
        super().mousePressEvent(event)

    def dragEnterEvent(self, event) -> None:  # type: ignore[override]
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
            return
        super().dragEnterEvent(event)

    def dropEvent(self, event) -> None:  # type: ignore[override]
        if event.mimeData().hasUrls():
            paths = [url.toLocalFile() for url in event.mimeData().urls() if url.toLocalFile()]
            if paths:
                self.filesDropped.emit(paths)
                event.acceptProposedAction()
                return
        super().dropEvent(event)


class OrderListDelegate(QStyledItemDelegate):
    """
    Draw the status part with a wavy black->white animation when status == "생성중".
    Item text format is expected: "{n}. {name} - {status}".
    """

    def __init__(self, parent=None) -> None:
        super().__init__(parent)
        self._phase = 0.0

    def advance(self) -> None:
        self._phase += 0.25
        if self._phase > 1e9:
            self._phase = 0.0

    def paint(self, painter, option, index) -> None:  # type: ignore[override]
        opt = option
        self.initStyleOption(opt, index)

        text = opt.text or ""
        # Let the style draw the background/selection, but we will custom draw the text.
        opt_text_backup = opt.text
        opt.text = ""
        style = opt.widget.style() if opt.widget else QApplication.style()
        style.drawControl(QStyle.ControlElement.CE_ItemViewItem, opt, painter, opt.widget)
        opt.text = opt_text_backup

        # Determine colors with contrast fallback.
        base_color = opt.palette.color(QPalette.ColorRole.Text)
        if opt.state & QStyle.StateFlag.State_Selected:
            base_color = opt.palette.color(QPalette.ColorRole.HighlightedText)
        else:
            bg = opt.palette.color(QPalette.ColorRole.Base)
            # If text color is too close to background, fall back to WindowText or light gray.
            if abs(base_color.red() - bg.red()) + abs(base_color.green() - bg.green()) + abs(base_color.blue() - bg.blue()) < 60:
                base_color = opt.palette.color(QPalette.ColorRole.WindowText)
                if abs(base_color.red() - bg.red()) + abs(base_color.green() - bg.green()) + abs(base_color.blue() - bg.blue()) < 60:
                    base_color = QColor(220, 220, 220)

        # Prepare text rect.
        rect = opt.rect.adjusted(8, 0, -8, 0)
        fm = opt.fontMetrics
        y = rect.y() + (rect.height() + fm.ascent() - fm.descent()) // 2
        x = rect.x()

        # Split into prefix and status.
        sep = " - "
        if sep not in text:
            painter.setPen(base_color)
            painter.drawText(rect, Qt.AlignmentFlag.AlignVCenter | Qt.AlignmentFlag.AlignLeft, text)
            return

        prefix, status = text.rsplit(sep, 1)
        prefix_with_sep = prefix + sep

        painter.save()
        painter.setFont(opt.font)

        # Draw prefix normally.
        painter.setPen(base_color)
        painter.drawText(x, y, prefix_with_sep)
        x += fm.horizontalAdvance(prefix_with_sep)

        status_text = status.strip()

        # Color mapping for status labels.
        status_colors = {
            "대기": QColor(150, 150, 150),
            "생성중": None,  # animated gray -> white
            "타이핑중": None,  # animated gray -> white
            "타이핑 대기": QColor(200, 190, 120),
            "코드 생성 완료": QColor(120, 200, 140),
            "타이핑 완료": QColor(130, 200, 255),
            "오류(빈 결과)": QColor(220, 120, 120),
            "생성 실패(건너뜀)": QColor(220, 150, 90),
        }

        if status_text not in status_colors:
            painter.setPen(base_color)
            painter.drawText(x, y, status)
            painter.restore()
            return

        target = status_colors[status_text]
        if target is not None:
            painter.setPen(target)
            painter.drawText(x, y, status)
            painter.restore()
            return

        # Animated gray -> white only (no vertical wobble).
        speed = 1.2
        phase = self._phase * speed
        # 0..1 pulse
        t = (math.sin(phase) * 0.5) + 0.5
        gray = int(round(140 + (255 - 140) * t))
        painter.setPen(QColor(gray, gray, gray))
        painter.drawText(x, y, status)

        painter.restore()


def main() -> None:
    app = QApplication(sys.argv)
    window = NovaAILiteWindow()
    window.setAttribute(Qt.WidgetAttribute.WA_DeleteOnClose)
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
