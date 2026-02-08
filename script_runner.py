from __future__ import annotations

import textwrap
import traceback
import re
from typing import Callable, Dict, List
import ast

from hwp_controller import HwpController


LogFn = Callable[[str], None]
CancelCheck = Callable[[], bool]

SAFE_BUILTINS: Dict[str, object] = {
    "range": range,
    "len": len,
    "min": min,
    "max": max,
    "enumerate": enumerate,
    "sum": sum,
    "print": print,
    "abs": abs,
}


class ScriptCancelled(RuntimeError):
    """Raised when script execution is cancelled."""


class ScriptRunner:
    def __init__(self, controller: HwpController) -> None:
        self._controller = controller

    def _split_concat_calls(self, line: str) -> List[str]:
        if " + " not in line:
            return [line]
        parts: List[str] = []
        buf: List[str] = []
        quote: str | None = None
        escaped = False
        i = 0
        while i < len(line):
            ch = line[i]
            if escaped:
                buf.append(ch)
                escaped = False
                i += 1
                continue
            if ch == "\\":
                buf.append(ch)
                escaped = True
                i += 1
                continue
            if ch in ("'", '"'):
                if quote is None:
                    quote = ch
                elif quote == ch:
                    quote = None
                buf.append(ch)
                i += 1
                continue
            # split only on " + " outside quotes
            if quote is None and line[i:i+3] == " + ":
                part = "".join(buf).strip()
                if part:
                    parts.append(part)
                buf = []
                i += 3
                continue
            buf.append(ch)
            i += 1
        tail = "".join(buf).strip()
        if tail:
            parts.append(tail)
        return parts if parts else [line]

    def _repair_multiline_calls(self, lines: List[str]) -> List[str]:
        def _count_unescaped(text: str, quote: str) -> int:
            count = 0
            escaped = False
            for ch in text:
                if escaped:
                    escaped = False
                    continue
                if ch == "\\":
                    escaped = True
                    continue
                if ch == quote:
                    count += 1
            return count

        repaired: List[str] = []
        buffer: List[str] = []
        quote_char: str | None = None
        for line in lines:
            if quote_char is None:
                if "insert_text(" in line or "insert_equation(" in line or "insert_latex_equation(" in line:
                    if _count_unescaped(line, "'") % 2 == 1:
                        quote_char = "'"
                        buffer = [line]
                        continue
                    if _count_unescaped(line, '"') % 2 == 1:
                        quote_char = '"'
                        buffer = [line]
                        continue
                repaired.append(line)
            else:
                buffer.append(line)
                count = sum(_count_unescaped(chunk, quote_char) for chunk in buffer)
                if count % 2 == 0:
                    joined = " ".join(part.strip() for part in buffer)
                    repaired.append(joined)
                    buffer = []
                    quote_char = None
        if buffer:
            joined = " ".join(part.strip() for part in buffer)
            if quote_char == "'" and not joined.strip().endswith("')"):
                joined = f"{joined}')"
            elif quote_char == '"' and not joined.strip().endswith('")'):
                joined = f'{joined}")'
            repaired.append(joined)
        return repaired

    def _sanitize_unterminated_equation_strings(self, script: str) -> str:
        lines = script.split("\n")
        out: List[str] = []
        for line in lines:
            if "insert_equation('" in line and line.count("'") % 2 == 1:
                out.append(line + "')")
            elif 'insert_equation("' in line and line.count('"') % 2 == 1:
                out.append(line + '")')
            else:
                out.append(line)
        return "\n".join(out)

    def _normalize_inline_calls(self, script: str) -> str:
        targets = ("insert_text(", "insert_equation(", "insert_latex_equation(")
        out: List[str] = []
        i = 0
        in_call = False
        quote_char: str | None = None
        quote_open = False
        while i < len(script):
            if not in_call:
                for t in targets:
                    if script.startswith(t, i):
                        in_call = True
                        break
            ch = script[i]
            if in_call:
                if ch in ("'", '"'):
                    if quote_char is None:
                        quote_char = ch
                        quote_open = True
                    elif quote_char == ch:
                        quote_open = not quote_open
                        if not quote_open:
                            quote_char = None
                if ch in ("\n", "\r", "\u2028", "\u2029"):
                    out.append(" ")
                    i += 1
                    continue
                if ch == ")" and not quote_open:
                    in_call = False
            out.append(ch)
            i += 1
        if in_call and quote_open:
            out.append(quote_char or "'")
            out.append(")")
        return "".join(out)

    def _sanitize_multiline_strings(self, script: str) -> str:
        out: List[str] = []
        quote_char: str | None = None
        escaped = False
        for ch in script:
            if escaped:
                out.append(ch)
                escaped = False
                continue
            if ch == "\\":
                out.append(ch)
                escaped = True
                continue
            if ch in ("'", '"'):
                if quote_char is None:
                    quote_char = ch
                elif quote_char == ch:
                    quote_char = None
                out.append(ch)
                continue
            if ch in ("\n", "\r", "\u2028", "\u2029") and quote_char is not None:
                out.append(" ")
                continue
            out.append(ch)
        if quote_char is not None:
            out.append(quote_char)
        return "".join(out)

    def _strip_code_markers(self, script: str) -> str:
        lines = script.split("\n")
        cleaned: List[str] = []
        for line in lines:
            stripped = line.strip()
            if stripped in ("[CODE]", "[/CODE]", "CODE"):
                continue
            cleaned.append(line)
        return "\n".join(cleaned)

    def _normalize_primes_in_equations(self, script: str) -> str:
        """
        Normalize prime notation inside insert_equation/insert_latex_equation strings.
        - Replace \\prime or \\Prime or unicode primes with apostrophe (')
        """
        def _fix(s: str) -> str:
            s = s.replace("′", "'").replace("’", "'")
            s = re.sub(r"\\+prime\b", "'", s, flags=re.IGNORECASE)
            # Some models emit backslash as prime marker: F\  -> F'
            # Only convert when backslash is NOT starting a command (e.g. \sqrt).
            s = re.sub(r"\\'+", "'", s)  # remove escaped apostrophes: \' -> '
            s = re.sub(r"([A-Za-z])\\(?![A-Za-z])", r"\1'", s)
            s = re.sub(r"\brm\s*([A-Za-z])\s*\\(?![A-Za-z])", r"rm\1'", s)
            # Special rule: F prime should be 'rm F prime' (with single spaces).
            s = re.sub(r"\brm\s*F\s*'", "rm F prime", s)
            s = re.sub(r"\brm\s*F\s*\\\\(?![A-Za-z])", "rm F prime", s)
            s = re.sub(r"\brm\s*F\s*prime\b", "rm F prime", s, flags=re.IGNORECASE)
            # Prime with rm should be tight: rm X' -> rmX'
            s = re.sub(r"\brm\s+([A-Za-z])'", r"rm\1'", s)
            return s

        pattern = re.compile(r"(insert_(?:equation|latex_equation)\()(['\"])(.*?)(\2\))", re.DOTALL)

        def repl(m: re.Match) -> str:
            return f"{m.group(1)}{m.group(2)}{_fix(m.group(3))}{m.group(4)}"

        return pattern.sub(repl, script)

    def _ensure_score_right_align(self, lines: List[str]) -> List[str]:
        out: List[str] = []
        score_re = re.compile(
            r"^\s*(insert_(?:text|equation|latex_equation))\(\s*(['\"])\s*(\[\s*(\d+)\s*점\s*\])\s*\2\s*\)\s*$"
        )
        need_extra_blank_line = False
        in_line_content = False
        for idx, line in enumerate(lines):
            stripped = line.strip()

            # Track whether the current line already has content (since last paragraph break)
            if stripped == "insert_paragraph()":
                in_line_content = False
                if need_extra_blank_line:
                    # This paragraph can serve as the blank line after score.
                    need_extra_blank_line = False
                out.append(line)
                continue

            if stripped == "insert_small_paragraph()":
                in_line_content = False
                if need_extra_blank_line:
                    need_extra_blank_line = False
                out.append(line)
                continue

            if need_extra_blank_line and stripped:
                # Ensure exactly one blank line after score before the next content.
                out.append("insert_paragraph()")
                need_extra_blank_line = False

            m = score_re.match(line)
            if m:
                # Remove extra blank lines before score (keep at most ONE paragraph break)
                while out and out[-1].strip() in ("insert_small_paragraph()", "insert_paragraph()"):
                    last = out[-1].strip()
                    if last == "insert_paragraph()":
                        # If there is another paragraph right before, drop extras
                        if len(out) >= 2 and out[-2].strip() == "insert_paragraph()":
                            out.pop()
                            continue
                        # Keep exactly one paragraph break
                        break
                    # Small paragraph before score creates visible blank space; remove it
                    out.pop()

                # Ensure score starts on a new line (single paragraph break only)
                if out and out[-1].strip() != "insert_paragraph()":
                    out.append("insert_paragraph()")
                in_line_content = False

                # Right align score line
                prev = out[-1].strip() if out else ""
                if prev != "set_align_right_next_line()":
                    out.append("set_align_right_next_line()")

                # Force score to be plain text (not equation)
                score_num = m.group(4)
                out.append(f"insert_text('[{score_num}점]')")
                out.append("insert_paragraph()")  # move to next line after score
                in_line_content = False
                need_extra_blank_line = True  # ensure one blank line below the score
                continue

            out.append(line)
            if stripped:
                in_line_content = True
        return out

    def _sanitize_tabs(self, lines: List[str]) -> List[str]:
        """
        Only keep insert_text('\\t') when it immediately precedes an insert_equation(...) line.
        Otherwise replace it with a single space.
        """
        out: List[str] = []
        i = 0
        while i < len(lines):
            line = lines[i]
            if line.strip() == "insert_text('\\t')" or line.strip() == 'insert_text("\\t")':
                j = i + 1
                while j < len(lines) and not lines[j].strip():
                    j += 1
                if j < len(lines) and lines[j].lstrip().startswith("insert_equation("):
                    out.append(line)
                else:
                    out.append("insert_text(' ')")
                i += 1
                continue
            out.append(line)
            i += 1
        return out

    def _normalize_placeholders(self, lines: List[str]) -> List[str]:
        """
        Ensure placeholder usage order is stable.
        - After entering the box placeholder (###), any later @@@ is treated as
          "move after box" to type choices outside.
        """
        out: List[str] = []
        seen_inside = False
        inserted_inside = False
        saw_template = False
        has_choices_placeholder = False
        saw_outside = False
        saw_after_box = False
        fp_re = re.compile(r"^\s*focus_placeholder\(\s*(['\"])(.*?)\1\s*\)\s*$")
        box_item_re = re.compile(r"^\s*insert_text\(\s*['\"]\s*[ㄱㄴㄷ]\.")
        content_re = re.compile(
            r"^\s*(insert_text|insert_equation|set_bold|set_align_justify_next_line|set_align_right_next_line)\("
        )
        box_start_re = re.compile(
            r"^\s*insert_text\(\s*['\"]\s*(\(|○|◎|●|•|ㄱ\.|ㄴ\.|ㄷ\.|가\.|나\.|다\.)"
        )
        choice_re = re.compile(r"^\s*insert_(?:text|equation)\(\s*['\"].*①")
        for line in lines:
            stripped = line.strip()
            if stripped.startswith("insert_template(") and any(
                name in stripped for name in ("header.hwp", "box.hwp", "box_white.hwp")
            ):
                saw_template = True
                if (
                    "header.hwp" in stripped
                    or "box_white.hwp" in stripped
                    or "box.hwp" in stripped
                ):
                    has_choices_placeholder = True
                out.append(line)
                continue
            m = fp_re.match(stripped)
            if not m:
                if saw_template and not saw_outside and content_re.match(stripped):
                    out.append("focus_placeholder('@@@')")
                    saw_outside = True
                if (
                    not seen_inside
                    and saw_template
                    and has_choices_placeholder
                    and saw_outside
                    and not saw_after_box
                    and (
                        stripped == "set_align_justify_next_line()"
                        or box_item_re.match(stripped)
                        or box_start_re.match(stripped)
                    )
                ):
                    out.append("focus_placeholder('###')")
                    seen_inside = True
                    inserted_inside = True
                if (
                    not seen_inside
                    and saw_template
                    and saw_outside
                    and box_item_re.match(stripped)
                ):
                    if out and out[-1].strip() == "set_align_justify_next_line()":
                        out.pop()
                        out.append("focus_placeholder('###')")
                        out.append("set_align_justify_next_line()")
                    else:
                        out.append("focus_placeholder('###')")
                    seen_inside = True
                    inserted_inside = True
                if (
                    seen_inside
                    and saw_template
                    and has_choices_placeholder
                    and not saw_after_box
                    and choice_re.match(stripped)
                ):
                    out.append("exit_box()")
                    out.append("insert_paragraph()")
                    out.append("focus_placeholder('&&&')")
                    saw_after_box = True
                out.append(line)
                continue
            marker = m.group(2)
            if marker == "###":
                if not inserted_inside:
                    seen_inside = True
                    out.append(line)
                continue
            if marker == "@@@":
                saw_outside = True
                if seen_inside:
                    out.append("exit_box()")
                    out.append("insert_paragraph()")
                    continue
                # If we're using a template with placeholders, consume @@@ here.
                if saw_template:
                    out.append(line)
                continue
            if marker == "&&&":
                if has_choices_placeholder:
                    saw_after_box = True
                    if seen_inside:
                        out.append("exit_box()")
                        out.append("insert_paragraph()")
                        out.append(line)
                        continue
                    out.append(line)
                    continue
                # If template has no &&& placeholder, ignore this marker.
                continue
            out.append(line)
        return out

    def _execute_fallback(
        self, script: str, log_fn: LogFn, cancel_check: CancelCheck | None = None
    ) -> None:
        funcs_no_args = {
            "insert_paragraph": self._controller.insert_paragraph,
            "insert_box": self._controller.insert_box,
            "exit_box": self._controller.exit_box,
            "insert_view_box": self._controller.insert_view_box,
            "insert_small_paragraph": self._controller.insert_small_paragraph,
            "set_align_right_next_line": self._controller.set_align_right_next_line,
            "set_align_justify_next_line": self._controller.set_align_justify_next_line,
            "set_table_border_white": self._controller.set_table_border_white,
        }
        funcs_one_str = {
            "insert_text": self._controller.insert_text,
            "insert_equation": self._controller.insert_equation,
            "insert_latex_equation": self._controller.insert_latex_equation,
            "insert_template": self._controller.insert_template,
            "focus_placeholder": self._controller.focus_placeholder,
        }
        funcs_one_int = {
            "set_char_width_ratio": self._controller.set_char_width_ratio,
        }

        i = 0
        text = script
        names = sorted(
            list(funcs_no_args.keys())
            + list(funcs_one_str.keys())
            + list(funcs_one_int.keys())
            + ["set_bold", "set_underline", "insert_table"],
            key=len,
            reverse=True,
        )
        while i < len(text):
            if cancel_check and cancel_check():
                raise ScriptCancelled("cancelled")
            matched = None
            for name in names:
                if text.startswith(name + "(", i):
                    matched = name
                    break
            if not matched:
                i += 1
                continue
            i += len(matched) + 1  # skip name + '('
            # parse args until matching ')', respecting quotes
            args = []
            depth = 1
            quote = None
            escaped = False
            while i < len(text) and depth > 0:
                ch = text[i]
                if escaped:
                    args.append(ch)
                    escaped = False
                    i += 1
                    continue
                if ch == "\\":
                    args.append(ch)
                    escaped = True
                    i += 1
                    continue
                if quote:
                    if ch == quote:
                        quote = None
                    args.append(ch)
                    i += 1
                    continue
                if ch in ("'", '"'):
                    quote = ch
                    args.append(ch)
                    i += 1
                    continue
                if ch == "(":
                    depth += 1
                elif ch == ")":
                    depth -= 1
                    if depth == 0:
                        i += 1
                        break
                args.append(ch)
                i += 1
            arg_str = "".join(args).strip()

            try:
                if cancel_check and cancel_check():
                    raise ScriptCancelled("cancelled")
                if matched in funcs_no_args:
                    funcs_no_args[matched]()
                elif matched in funcs_one_str:
                    s = ""
                    if arg_str.startswith(("'", '"')):
                        q = arg_str[0]
                        end = arg_str.find(q, 1)
                        if end == -1:
                            s = arg_str[1:]
                        else:
                            s = arg_str[1:end]
                    else:
                        s = arg_str
                    funcs_one_str[matched](s)
                elif matched == "set_bold":
                    val = "true" in arg_str.lower()
                    self._controller.set_bold(val)
                elif matched == "set_underline":
                    if not arg_str:
                        self._controller.set_underline()
                    else:
                        val = "true" in arg_str.lower()
                        self._controller.set_underline(val)
                elif matched in funcs_one_int:
                    try:
                        val = int(float(arg_str)) if arg_str else 0
                        funcs_one_int[matched](val)
                    except Exception:
                        pass
                elif matched == "insert_table":
                    # best-effort parse using literal_eval on args tuple
                    try:
                        node = ast.parse(f"f({arg_str})", mode="eval")
                        call = node.body  # type: ignore[attr-defined]
                        if isinstance(call, ast.Call):
                            eval_args = [ast.literal_eval(a) for a in call.args]
                            eval_kwargs = {kw.arg: ast.literal_eval(kw.value) for kw in call.keywords if kw.arg}
                            self._controller.insert_table(*eval_args, **eval_kwargs)
                    except Exception:
                        pass
            except Exception as exc:
                log_fn(f"[Fallback] {matched} failed: {exc}")

    def run(
        self,
        script: str,
        log: LogFn | None = None,
        *,
        cancel_check: CancelCheck | None = None,
    ) -> None:
        log_fn = log or (lambda *_: None)
        cleaned = textwrap.dedent(script or "").strip()
        # Normalize line separators (Windows CRLF / unicode separators)
        cleaned = (
            cleaned.replace("\r\n", "\n")
            .replace("\r", "\n")
            .replace("\u2028", "\n")
            .replace("\u2029", "\n")
        )
        if cleaned.startswith("```"):
            lines = cleaned.split("\n")[1:]
            if lines and lines[-1].strip() == "```":
                lines = lines[:-1]
            cleaned = "\n".join(lines).strip()
        cleaned = self._strip_code_markers(cleaned).strip()

        if not cleaned:
            log_fn("빈 스크립트라서 실행하지 않았습니다.")
            return

        # Normalize newlines inside any quoted strings
        cleaned = self._sanitize_multiline_strings(cleaned)
        # Normalize newlines inside insert_* calls
        cleaned = self._normalize_inline_calls(cleaned)
        # Fix unterminated equation strings on same line
        cleaned = self._sanitize_unterminated_equation_strings(cleaned)
        # Normalize prime notation inside equation strings
        cleaned = self._normalize_primes_in_equations(cleaned)
        expanded_lines: List[str] = []
        for line in self._repair_multiline_calls(cleaned.split("\n")):
            for sub_line in self._split_concat_calls(line):
                expanded_lines.append(sub_line)
        expanded_lines = self._normalize_placeholders(expanded_lines)
        expanded_lines = self._ensure_score_right_align(expanded_lines)
        expanded_lines = self._sanitize_tabs(expanded_lines)
        # Do not post-process choices; keep model output as-is.
        cleaned = "\n".join(expanded_lines).strip()

        def _wrap0(fn: Callable[[], None]) -> Callable[[], None]:
            def _inner() -> None:
                if cancel_check and cancel_check():
                    raise ScriptCancelled("cancelled")
                return fn()

            return _inner

        def _wrap1(fn: Callable[[str], None]) -> Callable[[str], None]:
            def _inner(arg: str) -> None:
                if cancel_check and cancel_check():
                    raise ScriptCancelled("cancelled")
                return fn(arg)

            return _inner

        def _wrap_bold(fn: Callable[[bool], None]) -> Callable[[bool], None]:
            def _inner(enabled: bool = True) -> None:
                if cancel_check and cancel_check():
                    raise ScriptCancelled("cancelled")
                return fn(enabled)

            return _inner

        def _wrap_underline(fn: Callable[[bool | None], None]) -> Callable[[bool | None], None]:
            def _inner(enabled: bool | None = None) -> None:
                if cancel_check and cancel_check():
                    raise ScriptCancelled("cancelled")
                return fn(enabled)

            return _inner

        def _wrap_table(fn: Callable[..., None]) -> Callable[..., None]:
            def _inner(*args, **kwargs) -> None:  # type: ignore[no-untyped-def]
                if cancel_check and cancel_check():
                    raise ScriptCancelled("cancelled")
                return fn(*args, **kwargs)

            return _inner

        env: Dict[str, object] = {
            "__builtins__": SAFE_BUILTINS,
            "insert_text": _wrap1(self._controller.insert_text),
            "insert_paragraph": _wrap0(self._controller.insert_paragraph),
            "insert_small_paragraph": _wrap0(self._controller.insert_small_paragraph),
            "insert_equation": _wrap1(self._controller.insert_equation),
            "insert_latex_equation": _wrap1(self._controller.insert_latex_equation),
            "insert_template": _wrap1(self._controller.insert_template),
            "focus_placeholder": _wrap1(self._controller.focus_placeholder),
            "insert_box": _wrap0(self._controller.insert_box),
            "exit_box": _wrap0(self._controller.exit_box),
            "insert_view_box": _wrap0(self._controller.insert_view_box),
            "insert_table": _wrap_table(self._controller.insert_table),
            "set_bold": _wrap_bold(self._controller.set_bold),
            "set_underline": _wrap_underline(self._controller.set_underline),
            "set_char_width_ratio": self._controller.set_char_width_ratio,
            "set_table_border_white": _wrap0(self._controller.set_table_border_white),
            "set_align_right_next_line": _wrap0(self._controller.set_align_right_next_line),
            "set_align_justify_next_line": _wrap0(self._controller.set_align_justify_next_line),
        }

        log_fn("스크립트 실행 시작")
        try:
            if cancel_check and cancel_check():
                raise ScriptCancelled("cancelled")
            exec(cleaned, env, {})
        except SyntaxError:
            log_fn("[Fallback] SyntaxError detected, running fallback parser.")
            self._execute_fallback(cleaned, log_fn, cancel_check=cancel_check)
        except ScriptCancelled:
            log_fn("스크립트 실행 취소됨")
            raise
        except Exception as exc:
            log_fn(traceback.format_exc())
            raise exc
        else:
            log_fn("스크립트 실행 완료")
