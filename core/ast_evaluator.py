"""
Safe AST-based arithmetic expression evaluator for Excel formula inputs.
Evaluates formulas like '=16000000*2+4567' without using insecure eval().
"""

import ast
import operator
import re
from typing import Union, Optional, Tuple

SAFE_OPERATORS = {
    ast.Add: operator.add,
    ast.Sub: operator.sub,
    ast.Mult: operator.mul,
    ast.Div: operator.truediv,
    ast.FloorDiv: operator.floordiv,
    ast.Pow: operator.pow,
    ast.USub: operator.neg,
    ast.UAdd: operator.pos,
}

IGNORE_PATTERNS = {
    "na", "n/a", "nil", "do not fill", "donotfill", "none", "-", "--", "null", "blank", "n.a.", "not applicable"
}


def is_ignored_text(val: any) -> bool:
    """Checks whether a cell value is an ignored placeholder like 'NA', 'Nil', 'DO NOT FILL'."""
    if val is None:
        return True
    s = str(val).strip().lower()
    if s in IGNORE_PATTERNS or s == "":
        return True
    return False


def safe_eval_expr(node: ast.AST) -> Union[int, float]:
    """Recursively evaluates an AST node containing only arithmetic operations."""
    if isinstance(node, ast.Expression):
        return safe_eval_expr(node.body)
    elif isinstance(node, ast.Constant):  # Python 3.8+ numeric constant
        if isinstance(node.value, (int, float)):
            return node.value
        raise ValueError(f"Unsupported constant type: {type(node.value)}")
    elif isinstance(node, ast.BinOp):
        left = safe_eval_expr(node.left)
        right = safe_eval_expr(node.right)
        op_type = type(node.op)
        if op_type in SAFE_OPERATORS:
            return SAFE_OPERATORS[op_type](left, right)
        raise ValueError(f"Unsupported operator: {op_type}")
    elif isinstance(node, ast.UnaryOp):
        operand = safe_eval_expr(node.operand)
        op_type = type(node.op)
        if op_type in SAFE_OPERATORS:
            return SAFE_OPERATORS[op_type](operand)
        raise ValueError(f"Unsupported unary operator: {op_type}")
    else:
        raise ValueError(f"Unsupported AST node: {type(node)}")


def evaluate_cell_value(val: any) -> Tuple[Optional[Union[int, float, str]], str]:
    """
    Parses a cell value:
    - If None or ignored text, returns (None, 'ignored')
    - If numeric (int/float), returns (number, 'numeric')
    - If arithmetic formula string (e.g. '=16000000*2+4567'), evaluates and returns (number, 'formula_evaluated')
    - If non-arithmetic formula or text, returns (raw_text, 'text')
    """
    if is_ignored_text(val):
        return (None, "ignored")

    if isinstance(val, (int, float)):
        return (val, "numeric")

    val_str = str(val).strip()

    # Check if string is a numeric literal directly (e.g. "12345" or "12.34" or "1,234.56")
    cleaned_num = val_str.replace(",", "")
    try:
        if "." in cleaned_num:
            return (float(cleaned_num), "numeric")
        else:
            return (int(cleaned_num), "numeric")
    except ValueError:
        pass

    # Check if formula string
    if val_str.startswith("="):
        expr_str = val_str[1:].strip().replace(",", "")
        # Only evaluate if it looks like arithmetic (no function calls like SUM, IF, etc.)
        if not re.search(r"[a-zA-Z_]\w*\(", expr_str):
            try:
                tree = ast.parse(expr_str, mode="eval")
                evaluated = safe_eval_expr(tree)
                return (evaluated, "formula_evaluated")
            except Exception:
                pass
        return (val_str, "formula_text")

    return (val_str, "text")
