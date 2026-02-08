from __future__ import annotations

from pathlib import Path

PROMPT_DIR = Path(__file__).resolve().parent / "prompts"


def _read_prompt(filename: str) -> str:
    path = PROMPT_DIR / filename
    if not path.exists():
        return ""
    text = path.read_text(encoding="utf-8")
    marker = "[PROMPT]"
    if marker in text:
        text = text.split(marker, 1)[1]
    lines = [line for line in text.splitlines() if not line.strip().startswith("#")]
    return "\n".join(lines).strip()


def get_image_instructions_prompt() -> str:
    return _read_prompt("image_instructions_prompt.txt")
