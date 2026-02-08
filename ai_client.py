from __future__ import annotations

import os
import sys
import base64
from pathlib import Path
from typing import Optional

def _debug(msg: str) -> None:
    sys.stderr.write(msg + "\n")
    sys.stderr.flush()

from prompt_loader import get_image_instructions_prompt
from backend.oauth_desktop import get_stored_user
from backend.firebase_profile import (
    check_usage_limit,
    increment_ai_usage,
    get_remaining_usage,
    get_plan_limit,
    PLAN_LIMITS,
)


MAX_IMAGE_DIM = 1600  # Reduce model processing time for large images


SYSTEM_PROMPT = """
You are generating a minimal Python script for HWP automation.
Use ONLY the following functions:
- insert_text("text")
- insert_paragraph()
- insert_small_paragraph()
- insert_equation("hwp_equation_syntax")
- insert_latex_equation("latex_math")
- insert_template("header.hwp|box.hwp|box_white.hwp")
- focus_placeholder("@@@|###")
- insert_box()
- exit_box()
- insert_view_box()
- insert_table(rows, cols, cell_data=[...], align_center=False, exit_after=True)
- set_bold(True/False)
- set_underline(True/False)
- set_table_border_white()
- set_align_right_next_line()
- set_align_justify_next_line()

Return ONLY Python code. No explanations.

수학 문제라고 판단되면 코드 맨 위에 아래 한 줄을 추가한다 ( [CODE] 표시는 쓰지 말 것 ):
MATH_CHOICES_EQUATION = True
""".strip()

try:
    from dotenv import load_dotenv
except Exception:
    load_dotenv = None


class AIClientError(RuntimeError):
    """Raised when AI client setup or call fails."""


def _load_env() -> None:
    if load_dotenv is None:
        return
    candidates = [
        Path(__file__).resolve().parent / ".env",
        Path.cwd() / ".env",
    ]
    for path in candidates:
        if path.exists():
            load_dotenv(dotenv_path=path)
            break


def _resolve_model(model: Optional[str]) -> str:
    if model:
        return model
    env_model = os.getenv("NOVA_AI_MODEL") or os.getenv("GEMINI_MODEL")
    return env_model or "gemini-2.5-flash"


class AIClient:
    def __init__(self, api_key: Optional[str] = None, model: Optional[str] = None, check_usage: bool = True) -> None:
        _load_env()
        self.api_key = api_key or os.getenv("GEMINI_API_KEY")
        if not self.api_key:
            raise AIClientError("GEMINI_API_KEY is missing.")

        try:
            import google.generativeai as genai
        except Exception as exc:
            raise AIClientError("google-generativeai package is not installed.") from exc

        genai.configure(api_key=self.api_key)
        self._genai = genai
        self.model = _resolve_model(model)
        self._check_usage = check_usage

    def _get_user_info(self) -> tuple[str | None, str]:
        """현재 사용자 정보 반환: (uid, tier)"""
        user = get_stored_user()
        if user and user.get("uid"):
            return user.get("uid"), user.get("tier", "Free")
        return None, "Free"

    def _check_usage_limit(self) -> None:
        """사용량 제한 체크"""
        if not self._check_usage:
            return
        
        uid, tier = self._get_user_info()
        if not uid:
            return  # 비로그인 상태에서는 체크 안함
        
        if not check_usage_limit(uid, tier):
            limit = get_plan_limit(tier)
            tier_label = {"Free": "무료", "free": "무료", "Standard": "Standard", "Plus": "Plus", "Pro": "Pro"}.get(tier, tier)
            
            # 업그레이드 안내 메시지
            upgrade_msg = ""
            if tier.lower() == "free":
                upgrade_msg = "\n\n💡 Standard 플랜으로 업그레이드하면 하루 50회까지 사용 가능!"
            elif tier.lower() == "standard":
                upgrade_msg = "\n\n💡 Plus 플랜으로 업그레이드하면 하루 220회까지 사용 가능!"
            elif tier.lower() == "plus":
                upgrade_msg = "\n\n💡 Pro 플랜으로 업그레이드하면 하루 660회까지 사용 가능!"
            
            raise AIClientError(
                f"⚠️ 일일 사용량 한도 초과! ({limit}/{limit})\n"
                f"현재 플랜: {tier_label}"
                f"{upgrade_msg}\n\n"
                "nova-ai.work에서 플랜을 업그레이드하거나 내일 다시 시도해주세요."
            )

    def _record_usage(self) -> None:
        """사용량 기록"""
        if not self._check_usage:
            return
        
        uid, tier = self._get_user_info()
        if not uid:
            return
        
        increment_ai_usage(uid)

    def _encode_image_to_base64(self, image_path: str) -> Optional[str]:
        try:
            with open(image_path, "rb") as image_file:
                return base64.b64encode(image_file.read()).decode("utf-8")
        except Exception:
            return None

    def generate_script(self, prompt: str, image_path: Optional[str] = None) -> str:
        if not prompt.strip():
            return ""

        # 사용량 제한 체크
        self._check_usage_limit()

        try:
            model = self._genai.GenerativeModel(self.model)
            if image_path:
                from PIL import Image  # type: ignore[import-not-found]

                image = Image.open(image_path).convert("RGB")
                max_dim = max(image.size)
                if max_dim > MAX_IMAGE_DIM:
                    scale = MAX_IMAGE_DIM / max_dim
                    new_size = (int(image.size[0] * scale), int(image.size[1] * scale))
                    image = image.resize(new_size, Image.LANCZOS)
                response = model.generate_content([prompt, image])
            else:
                response = model.generate_content(prompt)
            
            # 응답 텍스트 안전하게 추출
            result_text = ""
            try:
                if hasattr(response, "text"):
                    result_text = response.text
                elif hasattr(response, "parts") and response.parts:
                    result_text = "".join(part.text for part in response.parts if hasattr(part, "text"))
                elif hasattr(response, "candidates") and response.candidates:
                    for candidate in response.candidates:
                        if hasattr(candidate, "content") and hasattr(candidate.content, "parts"):
                            for part in candidate.content.parts:
                                if hasattr(part, "text"):
                                    result_text += part.text
            except Exception as text_err:
                _debug(f"[AI Debug] 응답 텍스트 추출 실패: {text_err}")
                # 차단 사유 확인
                if hasattr(response, "prompt_feedback"):
                    _debug(f"[AI Debug] Prompt feedback: {response.prompt_feedback}")
                if hasattr(response, "candidates") and response.candidates:
                    for i, c in enumerate(response.candidates):
                        if hasattr(c, "finish_reason"):
                            _debug(f"[AI Debug] Candidate {i} finish_reason: {c.finish_reason}")
                        if hasattr(c, "safety_ratings"):
                            _debug(f"[AI Debug] Candidate {i} safety_ratings: {c.safety_ratings}")
                result_text = ""
            
        except AIClientError:
            raise
        except Exception as exc:
            _debug(f"[AI Debug] generate_content 예외: {exc}")
            raise AIClientError(str(exc)) from exc

        if not result_text.strip():
            # 빈 결과일 때 디버그 정보 출력
            _debug(f"[AI Debug] 빈 응답 받음")
            if hasattr(response, "prompt_feedback"):
                _debug(f"[AI Debug] Prompt feedback: {response.prompt_feedback}")
            if hasattr(response, "candidates") and response.candidates:
                for i, c in enumerate(response.candidates):
                    if hasattr(c, "finish_reason"):
                        _debug(f"[AI Debug] Candidate {i} finish_reason: {c.finish_reason}")
                    if hasattr(c, "safety_ratings"):
                        _debug(f"[AI Debug] Candidate {i} safety_ratings: {c.safety_ratings}")
            return ""
        
        # 성공 시 사용량 기록
        self._record_usage()

        return result_text.strip()

    def build_prompt(
        self,
        description: str,
        image_path: Optional[str] = None,
        ocr_text: str = "",
    ) -> str:
        parts = [SYSTEM_PROMPT]
        if image_path:
            instructions = get_image_instructions_prompt()
            if instructions:
                parts.append(instructions)
        if ocr_text:
            parts.append(
                "OCR extracted text (use this to improve accuracy; "
                "verify with the image and fix obvious OCR errors):\n"
                f"{ocr_text}"
            )
        if description:
            parts.append(f"User request: {description}")
        else:
            parts.append("User request: Extract the image content and type it into HWP.")
        return "\n\n".join(parts)

    def generate_script_for_image(
        self, image_path: str, description: str = "", ocr_text: str = ""
    ) -> str:
        prompt = self.build_prompt(description, image_path=image_path, ocr_text=ocr_text)
        return self.generate_script(prompt, image_path=image_path)
