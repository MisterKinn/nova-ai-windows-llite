from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Any
import subprocess


@dataclass
class EquationOptions:
    font_size_pt: float = 8.0
    eq_font_name: str = "HyhwpEQ"
    treat_as_char: bool = True
    ensure_newline: bool = False


class EquationAutomationError(RuntimeError):
    """Raised when low-level equation actions cannot be executed."""


class LatexConversionError(RuntimeError):
    """Raised when LaTeX conversion fails unexpectedly."""


NODE_CLI = Path(__file__).resolve().parents[1] / "node_eqn" / "hwp_eqn_cli.js"


def latex_to_hwpeqn(latex: str, timeout: float = 10.0) -> str:
    text = (latex or "").strip()
    if not text:
        return ""

    if not NODE_CLI.exists():
        return latex

    try:
        result = subprocess.run(
            ["node", str(NODE_CLI)],
            input=text,
            capture_output=True,
            text=True,
            timeout=timeout,
            check=True,
        )
        output = result.stdout.strip()
        return output or latex
    except FileNotFoundError:
        return latex
    except subprocess.CalledProcessError:
        return latex
    except subprocess.TimeoutExpired:
        return latex
    except Exception as exc:
        raise LatexConversionError(str(exc)) from exc


def insert_equation_control(
    hwp: Any,
    hwpeqn: str,
    *,
    options: EquationOptions | None = None,
) -> None:
    text = (hwpeqn or "").strip()
    if not text:
        return

    opts = options or EquationOptions()

    try:
        action = hwp.HAction
        param_sets = hwp.HParameterSet
    except AttributeError as exc:
        raise EquationAutomationError("HAction interface is missing on this HWP session.") from exc

    try:
        eq_param = param_sets.HEqEdit
        action.GetDefault("EquationCreate", eq_param.HSet)
        eq_param.EqFontName = opts.eq_font_name
        eq_param.string = text
        eq_param.BaseUnit = _point_to_hwp_unit(hwp, opts.font_size_pt)
        action.Execute("EquationCreate", eq_param.HSet)

        if hasattr(hwp, "FindCtrl"):
            hwp.FindCtrl()
        shape_param = param_sets.HShapeObject
        action.GetDefault("EquationPropertyDialog", shape_param.HSet)
        shape_param.HSet.SetItem("ShapeType", 3)
        shape_param.Version = "Equation Version 60"
        shape_param.EqFontName = opts.eq_font_name
        shape_param.HSet.SetItem("ApplyTo", 0)
        shape_param.HSet.SetItem("TreatAsChar", 1 if opts.treat_as_char else 0)
        action.Execute("EquationPropertyDialog", shape_param.HSet)

        hwp.Run("Cancel")
        action.Run("MoveRight")
        if opts.ensure_newline:
            action.Run("BreakPara")
    except EquationAutomationError:
        raise
    except Exception as exc:
        raise EquationAutomationError(f"Failed to insert equation: {exc}") from exc


def _point_to_hwp_unit(hwp: Any, point: float) -> float:
    if point <= 0:
        return 0.0
    if hasattr(hwp, "PointToHwpUnit"):
        return hwp.PointToHwpUnit(point)
    return point * 100.0
