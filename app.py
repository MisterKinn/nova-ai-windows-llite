from __future__ import annotations

import argparse
import sys
from pathlib import Path
from typing import Optional

# Allow running this file directly (python app.py) by ensuring the
# package parent directory is on sys.path (so `import litepro...` works).
if __package__ in (None, ""):
    pkg_parent = Path(__file__).resolve().parent.parent
    if str(pkg_parent) not in sys.path:
        sys.path.insert(0, str(pkg_parent))

from ai_client import AIClient, AIClientError
from hwp_controller import HwpController, HwpControllerError
from script_runner import ScriptRunner


SYSTEM_PROMPT = """
You are generating a minimal Python script for HWP automation.
Use ONLY the following functions:
- insert_text("text")
- insert_paragraph()
- insert_equation("hwp_equation_syntax")
- insert_latex_equation("latex_math")
- insert_template("header.hwp|box.hwp|box_white.hwp")
- focus_placeholder("@@@|###")

Return ONLY Python code. No explanations.
""".strip()


def _extract_code(text: str) -> str:
    if not text:
        return ""
    cleaned = text.strip()
    if cleaned.startswith("```"):
        lines = cleaned.split("\n")[1:]
        if lines and lines[-1].strip() == "```":
            lines = lines[:-1]
        return "\n".join(lines).strip()
    return cleaned


def _read_file(path: Path) -> str:
    return path.read_text(encoding="utf-8")


def cmd_detect(_: argparse.Namespace) -> int:
    titles = HwpController.find_hwp_windows()
    if not titles:
        print("HWP 창을 찾지 못했습니다.")
        return 1
    print("감지된 HWP 창:")
    for title in titles:
        print(f"- {title}")
    return 0


def _connect_controller() -> HwpController:
    controller = HwpController()
    controller.connect()
    return controller


def cmd_insert_text(args: argparse.Namespace) -> int:
    controller = _connect_controller()
    controller.insert_text(args.text)
    if args.paragraph:
        controller.insert_paragraph()
    return 0


def cmd_insert_equation(args: argparse.Namespace) -> int:
    controller = _connect_controller()
    if args.latex:
        controller.insert_latex_equation(
            args.equation,
            font_size_pt=args.font_size,
            eq_font_name=args.font_name,
            treat_as_char=not args.no_treat_as_char,
            ensure_newline=args.newline,
        )
    else:
        controller.insert_equation(
            args.equation,
            font_size_pt=args.font_size,
            eq_font_name=args.font_name,
            treat_as_char=not args.no_treat_as_char,
            ensure_newline=args.newline,
        )
    return 0


def cmd_run_script(args: argparse.Namespace) -> int:
    controller = _connect_controller()
    runner = ScriptRunner(controller)
    script = _read_file(Path(args.file))
    runner.run(script, log=print)
    return 0


def cmd_ai_generate(args: argparse.Namespace) -> int:
    prompt = f"{SYSTEM_PROMPT}\n\nUser request: {args.description}"
    try:
        client = AIClient(model=args.model)
        result = client.generate_script(prompt)
    except AIClientError as exc:
        print(f"AI 오류: {exc}")
        return 1

    code = _extract_code(result)
    if args.output:
        Path(args.output).write_text(code, encoding="utf-8")
        print(f"저장 완료: {args.output}")
    else:
        print(code)
    return 0


def cmd_ai_run(args: argparse.Namespace) -> int:
    prompt = f"{SYSTEM_PROMPT}\n\nUser request: {args.description}"
    try:
        client = AIClient(model=args.model)
        result = client.generate_script(prompt)
    except AIClientError as exc:
        print(f"AI 오류: {exc}")
        return 1

    code = _extract_code(result)
    if not code:
        print("AI가 빈 스크립트를 반환했습니다.")
        return 1

    controller = _connect_controller()
    runner = ScriptRunner(controller)
    runner.run(code, log=print)
    return 0


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="LitePro - minimal HWP automation")
    subparsers = parser.add_subparsers(dest="command", required=True)

    detect = subparsers.add_parser("detect", help="HWP 창 감지")
    detect.set_defaults(func=cmd_detect)

    insert_text = subparsers.add_parser("insert-text", help="텍스트 입력")
    insert_text.add_argument("text", help="입력할 텍스트")
    insert_text.add_argument("--paragraph", action="store_true", help="입력 후 줄바꿈")
    insert_text.set_defaults(func=cmd_insert_text)

    insert_eq = subparsers.add_parser("insert-equation", help="수식 입력 (HwpEqn)")
    insert_eq.add_argument("equation", help="HwpEqn 수식 문자열")
    insert_eq.add_argument("--font-size", type=float, default=10.0)
    insert_eq.add_argument("--font-name", default="HyhwpEQ")
    insert_eq.add_argument("--no-treat-as-char", action="store_true")
    insert_eq.add_argument("--newline", action="store_true")
    insert_eq.add_argument("--latex", action="store_true", help="LaTeX 입력으로 변환 후 삽입")
    insert_eq.set_defaults(func=cmd_insert_equation)

    insert_latex = subparsers.add_parser("insert-latex-equation", help="LaTeX 수식 입력")
    insert_latex.add_argument("equation", help="LaTeX 수식 문자열")
    insert_latex.add_argument("--font-size", type=float, default=10.0)
    insert_latex.add_argument("--font-name", default="HyhwpEQ")
    insert_latex.add_argument("--no-treat-as-char", action="store_true")
    insert_latex.add_argument("--newline", action="store_true")
    insert_latex.set_defaults(
        func=lambda args: cmd_insert_equation(
            argparse.Namespace(
                equation=args.equation,
                font_size=args.font_size,
                font_name=args.font_name,
                no_treat_as_char=args.no_treat_as_char,
                newline=args.newline,
                latex=True,
            )
        )
    )

    run_script = subparsers.add_parser("run-script", help="스크립트 실행")
    run_script.add_argument("--file", required=True, help="파이썬 스크립트 경로")
    run_script.set_defaults(func=cmd_run_script)

    ai_gen = subparsers.add_parser("ai-generate", help="AI로 스크립트 생성")
    ai_gen.add_argument("description", help="요청 설명")
    ai_gen.add_argument("--model", default="gemini-3-flash-preview")
    ai_gen.add_argument("--output", help="저장할 파일 경로")
    ai_gen.set_defaults(func=cmd_ai_generate)

    ai_run = subparsers.add_parser("ai-run", help="AI로 생성 후 실행")
    ai_run.add_argument("description", help="요청 설명")
    ai_run.add_argument("--model", default="gemini-3-flash-preview")
    ai_run.set_defaults(func=cmd_ai_run)

    return parser


def main(argv: Optional[list[str]] = None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)
    try:
        return args.func(args)
    except HwpControllerError as exc:
        print(f"HWP 오류: {exc}")
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
