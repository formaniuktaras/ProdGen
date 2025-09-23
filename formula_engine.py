"""Formula evaluation engine with Google Sheets-like syntax.

This module exposes :class:`FormulaEngine`, which can parse and evaluate
expressions that mimic a subset of Google Sheets formulas.  The engine is
designed to be safe (no ``eval`` usage), extensible (custom functions can be
registered), and convenient for templating data exports where variables are
embedded in ``{{variable}}`` placeholders.

The supported feature-set includes arithmetic, comparisons, string
concatenation, and a collection of spreadsheet-inspired functions such as
``IF``, ``SUM``, ``TEXT`` and ``TODAY``.  Arguments inside formulas are
separated with semicolons to match the locale-specific style used in the
examples (e.g. ``=SUM(1; 2; 3)``).
"""

from __future__ import annotations

from dataclasses import dataclass
from datetime import date, datetime, time
from decimal import Decimal, ROUND_HALF_UP
import inspect
import math
import re
from typing import Any, Callable, Dict, Iterable, Iterator, List, Mapping, Optional, Sequence, Tuple


class FormulaError(Exception):
    """Raised when a formula cannot be parsed or evaluated."""


@dataclass(frozen=True)
class _Token:
    """Single lexical token produced by the tokenizer."""

    type: str
    value: Any


class _Tokenizer:
    """Turn a formula string into a stream of tokens."""

    _NUMBER_RE = re.compile(r"[+-]?((\d+\.\d*)|(\d*\.\d+)|(\d+))(e[+-]?\d+)?", re.IGNORECASE)

    def __init__(self, text: str) -> None:
        self.text = text
        self.length = len(text)
        self.pos = 0

    def tokenize(self) -> List[_Token]:
        tokens: List[_Token] = []
        while self.pos < self.length:
            ch = self.text[self.pos]
            if ch.isspace():
                self.pos += 1
                continue
            if ch == '{' and self._peek('{'):
                tokens.append(self._consume_variable())
                continue
            if ch == '"':
                tokens.append(self._consume_string())
                continue
            if ch.isdigit() or (ch in '+-' and self._peek_digit()):
                tokens.append(self._consume_number())
                continue
            if ch.isalpha() or ch == '_':
                tokens.append(self._consume_identifier())
                continue
            if ch == '(':
                tokens.append(_Token('LPAREN', ch))
                self.pos += 1
                continue
            if ch == ')':
                tokens.append(_Token('RPAREN', ch))
                self.pos += 1
                continue
            if ch in ';,':
                tokens.append(_Token('SEMI', ch))
                self.pos += 1
                continue
            if ch == '&':
                tokens.append(_Token('OP', ch))
                self.pos += 1
                continue
            # multi-character comparisons first
            if self.text.startswith('>=', self.pos):
                tokens.append(_Token('COMPARE', '>='))
                self.pos += 2
                continue
            if self.text.startswith('<=', self.pos):
                tokens.append(_Token('COMPARE', '<='))
                self.pos += 2
                continue
            if self.text.startswith('<>', self.pos):
                tokens.append(_Token('COMPARE', '<>'))
                self.pos += 2
                continue
            if ch in '=<>':
                tokens.append(_Token('COMPARE', ch))
                self.pos += 1
                continue
            if ch in '+-*/^':
                tokens.append(_Token('OP', ch))
                self.pos += 1
                continue
            raise FormulaError(f"Unexpected character '{ch}' at position {self.pos}.")
        tokens.append(_Token('EOF', None))
        return tokens

    def _peek(self, char: str) -> bool:
        return self.pos + 1 < self.length and self.text[self.pos + 1] == char

    def _peek_digit(self) -> bool:
        return self.pos + 1 < self.length and self.text[self.pos + 1].isdigit()

    def _consume_variable(self) -> _Token:
        end = self.text.find('}}', self.pos)
        if end == -1:
            raise FormulaError("Unterminated variable placeholder.")
        name = self.text[self.pos + 2 : end].strip()
        if not name:
            raise FormulaError("Empty variable placeholder encountered.")
        self.pos = end + 2
        return _Token('VAR', name)

    def _consume_string(self) -> _Token:
        self.pos += 1  # skip opening quote
        buffer: List[str] = []
        while self.pos < self.length:
            ch = self.text[self.pos]
            if ch == '"':
                if self._peek('"'):
                    buffer.append('"')
                    self.pos += 2
                    continue
                self.pos += 1
                return _Token('STRING', ''.join(buffer))
            if ch == '\\':
                self.pos += 1
                if self.pos >= self.length:
                    raise FormulaError("Invalid escape sequence in string literal.")
                escape = self.text[self.pos]
                escapes = {'n': '\n', 't': '\t', 'r': '\r', '"': '"', '\\': '\\'}
                buffer.append(escapes.get(escape, escape))
                self.pos += 1
                continue
            buffer.append(ch)
            self.pos += 1
        raise FormulaError("Unterminated string literal.")

    def _consume_number(self) -> _Token:
        match = self._NUMBER_RE.match(self.text, self.pos)
        if not match:
            raise FormulaError(f"Invalid number at position {self.pos}.")
        literal = match.group(0)
        self.pos = match.end()
        if any(c in literal for c in '.eE'):
            value: Any = float(literal)
        else:
            value = int(literal)
        return _Token('NUMBER', value)

    def _consume_identifier(self) -> _Token:
        start = self.pos
        self.pos += 1
        while self.pos < self.length and (self.text[self.pos].isalnum() or self.text[self.pos] in '._'):
            self.pos += 1
        ident = self.text[start:self.pos]
        upper = ident.upper()
        if upper == 'TRUE':
            return _Token('BOOLEAN', True)
        if upper == 'FALSE':
            return _Token('BOOLEAN', False)
        if upper in {'NULL', 'NONE'}:
            return _Token('NULL', None)
        return _Token('IDENT', ident)


class _Parser:
    """Produce an abstract syntax tree from the token stream."""

    _OP_PRECEDENCE: Dict[str, int] = {
        '^': 4,
        '*': 3,
        '/': 3,
        '+': 2,
        '-': 2,
        '&': 2,
    }

    def __init__(self, tokens: Sequence[_Token]) -> None:
        self.tokens = tokens
        self.pos = 0

    def parse(self) -> Any:
        expression = self._parse_expression()
        self._expect('EOF')
        return expression

    def _parse_expression(self, min_precedence: int = 0, stop_types: Optional[Sequence[str]] = None) -> Any:
        stop = set(stop_types or ())
        node = self._parse_prefix(stop)
        while True:
            token = self._peek()
            if token.type in stop or token.type == 'EOF':
                break
            if token.type == 'OP':
                precedence = self._OP_PRECEDENCE[token.value]
                if precedence < min_precedence:
                    break
                self._advance()
                next_min = precedence + (0 if token.value == '^' else 1)
                rhs = self._parse_expression(next_min, stop)
                node = ('binop', token.value, node, rhs)
                continue
            if token.type == 'COMPARE':
                precedence = 1
                if precedence < min_precedence:
                    break
                op = token.value
                self._advance()
                rhs = self._parse_expression(precedence + 1, stop)
                node = ('compare', op, node, rhs)
                continue
            break
        return node

    def _parse_prefix(self, stop: set[str]) -> Any:
        token = self._advance()
        if token.type == 'NUMBER':
            return ('literal', token.value)
        if token.type == 'STRING':
            return ('literal', token.value)
        if token.type == 'BOOLEAN':
            return ('literal', token.value)
        if token.type == 'NULL':
            return ('literal', None)
        if token.type == 'VAR':
            return ('var', token.value)
        if token.type == 'OP' and token.value in ('+', '-'):
            operand = self._parse_expression(self._OP_PRECEDENCE['^'], stop)
            return ('unary', token.value, operand)
        if token.type == 'LPAREN':
            expr = self._parse_expression(0, stop_types=('RPAREN',))
            self._expect('RPAREN')
            return expr
        if token.type == 'IDENT':
            next_token = self._peek()
            if next_token.type == 'LPAREN':
                self._advance()  # consume LPAREN
                args = self._parse_arguments()
                return ('func', token.value.upper(), args)
            raise FormulaError(f"Unexpected identifier '{token.value}'. Variables must use {{}} notation.")
        raise FormulaError(f"Unexpected token {token.type!r} in expression.")

    def _parse_arguments(self) -> List[Any]:
        args: List[Any] = []
        if self._peek().type == 'RPAREN':
            self._advance()
            return args
        while True:
            args.append(self._parse_expression(0, stop_types=('SEMI', 'RPAREN')))
            token = self._peek()
            if token.type == 'SEMI':
                self._advance()
                continue
            if token.type == 'RPAREN':
                self._advance()
                break
            raise FormulaError("Expected ';' or ')' in function arguments.")
        return args

    def _peek(self) -> _Token:
        return self.tokens[self.pos]

    def _advance(self) -> _Token:
        token = self.tokens[self.pos]
        self.pos += 1
        return token

    def _expect(self, token_type: str) -> _Token:
        token = self._advance()
        if token.type != token_type:
            raise FormulaError(f"Expected token {token_type!r}, found {token.type!r}.")
        return token


class FormulaEngine:
    """Evaluates spreadsheet-like formulas in a controlled environment."""

    FUNCTIONS: Dict[str, Callable[..., Any]] = {}

    @classmethod
    def register_function(cls, name: str, func: Callable[..., Any]) -> None:
        """Register a callable under the provided uppercase name."""

        if not callable(func):
            raise TypeError("Registered function must be callable.")
        cls.FUNCTIONS[name.upper()] = func

    @classmethod
    def _prepare_formula(cls, formula: str) -> Tuple[Optional[str], bool]:
        expression = formula.strip()
        if not expression:
            return None, True
        if expression.startswith('='):
            expression = expression[1:]
            if not expression.strip():
                return None, True
            expression = expression.strip()
        return expression, False

    @classmethod
    def _parse_tokens(cls, expression: str) -> Any:
        tokenizer = _Tokenizer(expression)
        tokens = tokenizer.tokenize()
        parser = _Parser(tokens)
        return parser.parse()

    @classmethod
    def parse(cls, formula: str) -> Any:
        if formula is None:
            raise FormulaError("Formula is empty.")
        if not isinstance(formula, str):
            raise FormulaError("Formula must be a string.")
        expression, is_blank = cls._prepare_formula(formula)
        if expression is None or is_blank:
            raise FormulaError("Formula is empty.")
        return cls._parse_tokens(expression)

    @classmethod
    def describe(cls, formula: str) -> Dict[str, Any]:
        ast = cls.parse(formula)
        variables: set[str] = set()
        functions: set[str] = set()
        cls._collect_metadata(ast, variables, functions)
        return {"ast": ast, "variables": variables, "functions": functions}

    @classmethod
    def _collect_metadata(cls, node: Any, variables: set[str], functions: set[str]) -> None:
        node_type = node[0]
        if node_type == 'var':
            variables.add(node[1])
            return
        if node_type == 'func':
            functions.add(node[1])
            for arg in node[2]:
                cls._collect_metadata(arg, variables, functions)
            return
        if node_type == 'unary':
            cls._collect_metadata(node[2], variables, functions)
            return
        if node_type in ('binop', 'compare'):
            cls._collect_metadata(node[2], variables, functions)
            cls._collect_metadata(node[3], variables, functions)
            return

    @classmethod

    def evaluate(cls, formula: str, context: Optional[Mapping[str, Any]] = None) -> Any:
        """Evaluate a single formula using the supplied context."""

        if formula is None:
            return None
        if not isinstance(formula, str):
            return formula
        expression, is_blank = cls._prepare_formula(formula)
        if expression is None or is_blank:
            return ''
        ast = cls._parse_tokens(expression)
        eval_context = dict(context or {})
        return cls._evaluate_node(ast, eval_context)

    @classmethod
    def _evaluate_node(cls, node: Any, context: Mapping[str, Any]) -> Any:
        node_type = node[0]
        if node_type == 'literal':
            return node[1]
        if node_type == 'var':
            name = node[1]
            if name not in context:
                raise FormulaError(f"Unknown variable '{name}'.")
            value = context[name]
            if callable(value):
                value = value()
            return value
        if node_type == 'unary':
            op, operand_node = node[1], node[2]
            operand = cls._evaluate_node(operand_node, context)
            number = cls._to_number(operand)
            if op == '+':
                return number
            return -number
        if node_type == 'binop':
            op, left_node, right_node = node[1], node[2], node[3]
            left = cls._evaluate_node(left_node, context)
            right = cls._evaluate_node(right_node, context)
            return cls._apply_operator(op, left, right)
        if node_type == 'compare':
            op, left_node, right_node = node[1], node[2], node[3]
            left = cls._evaluate_node(left_node, context)
            right = cls._evaluate_node(right_node, context)
            return cls._apply_comparison(op, left, right)
        if node_type == 'func':
            func_name, args_nodes = node[1], node[2]
            func = cls.FUNCTIONS.get(func_name)
            if func is None:
                raise FormulaError(f"Unknown function '{func_name}'.")
            args = [cls._evaluate_node(arg, context) for arg in args_nodes]
            try:
                return cls._call_function(func, args, context)
            except FormulaError:
                raise
            except Exception as exc:  # pragma: no cover - safety net
                raise FormulaError(f"Error executing function '{func_name}': {exc}") from exc
        raise FormulaError(f"Unsupported node type '{node_type}'.")

    @classmethod
    def _call_function(cls, func: Callable[..., Any], args: Sequence[Any], context: Mapping[str, Any]) -> Any:
        signature = inspect.signature(func)
        kwargs: Dict[str, Any] = {}
        for param in signature.parameters.values():
            if param.kind in (param.POSITIONAL_OR_KEYWORD, param.KEYWORD_ONLY):
                if param.name == 'context':
                    kwargs['context'] = context
                elif param.name == 'engine':
                    kwargs['engine'] = cls
        return func(*args, **kwargs)

    @staticmethod
    def _to_number(value: Any) -> float:
        if value is None:
            return 0.0
        if isinstance(value, bool):
            return 1.0 if value else 0.0
        if isinstance(value, (int, float)):
            return float(value)
        if isinstance(value, Decimal):
            return float(value)
        if isinstance(value, (date, datetime)):
            base = datetime.combine(date(1899, 12, 30), time())
            delta = (value - base) if isinstance(value, datetime) else (datetime.combine(value, time()) - base)
            return delta.days + delta.seconds / 86400
        if isinstance(value, str):
            stripped = value.strip()
            if stripped == '':
                return 0.0
            try:
                return float(stripped) if any(c in stripped for c in '.eE') else float(int(stripped))
            except ValueError as exc:
                raise FormulaError(f"Cannot convert '{value}' to a number.") from exc
        raise FormulaError(f"Unsupported value type '{type(value).__name__}' for numeric coercion.")

    @staticmethod
    def _apply_operator(op: str, left: Any, right: Any) -> Any:
        if op == '&':
            return ''.join(_flatten_text([left, right]))
        left_num = FormulaEngine._to_number(left)
        right_num = FormulaEngine._to_number(right)
        if op == '+':
            return FormulaEngine._normalize_number(left_num + right_num)
        if op == '-':
            return FormulaEngine._normalize_number(left_num - right_num)
        if op == '*':
            return FormulaEngine._normalize_number(left_num * right_num)
        if op == '/':
            if right_num == 0:
                raise FormulaError("Division by zero.")
            return FormulaEngine._normalize_number(left_num / right_num)
        if op == '^':
            return FormulaEngine._normalize_number(math.pow(left_num, right_num))
        raise FormulaError(f"Unsupported operator '{op}'.")

    @staticmethod
    def _apply_comparison(op: str, left: Any, right: Any) -> bool:
        left_cmp, left_is_num = _coerce_comparable(left)
        right_cmp, right_is_num = _coerce_comparable(right)
        if left_is_num and right_is_num:
            left_value, right_value = left_cmp, right_cmp
        else:
            left_value, right_value = left_cmp, right_cmp
        if op == '=':
            return left_value == right_value
        if op == '<>':
            return left_value != right_value
        if op == '>':
            return left_value > right_value
        if op == '>=':
            return left_value >= right_value
        if op == '<':
            return left_value < right_value
        if op == '<=':
            return left_value <= right_value
        raise FormulaError(f"Unsupported comparison operator '{op}'.")

    @staticmethod
    def _normalize_number(value: float) -> Any:
        if isinstance(value, float) and value.is_integer():
            return int(value)
        return value


def _flatten(args: Iterable[Any]) -> Iterator[Any]:
    for value in args:
        if isinstance(value, (list, tuple)):
            yield from _flatten(value)
        else:
            yield value


def _flatten_text(args: Iterable[Any]) -> Iterator[str]:
    for value in _flatten(args):
        if value is None:
            continue
        yield str(value)


def _coerce_comparable(value: Any) -> Tuple[Any, bool]:
    try:
        return FormulaEngine._to_number(value), True
    except FormulaError:
        if isinstance(value, (datetime, date, time)):
            return value, False
        if isinstance(value, bool):
            return value, False
        if value is None:
            return '', False
        return str(value), False


def _round_half_up(number: Any, digits: int = 0) -> float:
    decimal_value = Decimal(str(FormulaEngine._to_number(number)))
    quantizer = Decimal('1').scaleb(-digits)
    result = decimal_value.quantize(quantizer, rounding=ROUND_HALF_UP)
    if digits <= 0:
        return int(result)
    return float(result)


def _round_with_mode(number: Any, digits: int, mode: str) -> float:
    number_value = FormulaEngine._to_number(number)
    power = 10 ** digits
    if digits >= 0:
        scaled = number_value * power
    else:
        scaled = number_value / (10 ** (-digits))
    if mode == 'up':
        scaled = math.ceil(scaled) if scaled >= 0 else math.floor(scaled)
    else:
        scaled = math.floor(scaled) if scaled >= 0 else math.ceil(scaled)
    if digits >= 0:
        return scaled / power
    return scaled * (10 ** (-digits))


def _ensure_text(value: Any) -> str:
    return '' if value is None else str(value)


def _normalize_datetime(value: Any) -> datetime:
    if isinstance(value, datetime):
        return value
    if isinstance(value, date):
        return datetime.combine(value, time())
    if isinstance(value, time):
        today = datetime.combine(date.today(), value)
        return today
    if isinstance(value, str):
        text = value.strip()
        for parse in (datetime.fromisoformat, _parse_date, _parse_time):
            try:
                parsed = parse(text)
            except (ValueError, TypeError):
                continue
            if isinstance(parsed, datetime):
                return parsed
            if isinstance(parsed, date):
                return datetime.combine(parsed, time())
            if isinstance(parsed, time):
                return datetime.combine(date.today(), parsed)
    raise FormulaError(f"Cannot interpret '{value}' as a date/time value.")


def _parse_date(text: str) -> date:
    return date.fromisoformat(text)


def _parse_time(text: str) -> time:
    return time.fromisoformat(text)


def _format_text(value: Any, fmt: str) -> str:
    if isinstance(value, (datetime, date, time)):
        dt = _normalize_datetime(value)
        pattern = fmt
        replacements = [
            ('YYYY', '%Y'),
            ('YY', '%y'),
            ('MM', '%m'),
            ('DD', '%d'),
            ('HH', '%H'),
            ('hh', '%I'),
            ('mm', '%M'),
            ('ss', '%S'),
        ]
        for token, repl in replacements:
            pattern = pattern.replace(token, repl)
        return dt.strftime(pattern)
    try:
        number = FormulaEngine._to_number(value)
    except FormulaError:
        try:
            return fmt.format(value=value)
        except Exception:
            return _ensure_text(value)
    if '#' in fmt or '0' in fmt:
        if '.' in fmt:
            decimals = len(fmt.split('.', 1)[1])
            rounded = _round_half_up(number, decimals)
            return f"{rounded:.{decimals}f}"
        return str(_round_half_up(number, 0))
    try:
        return format(number, fmt)
    except Exception:
        return str(number)


def _truthy(value: Any) -> bool:
    if isinstance(value, str):
        return value.strip() != ''
    if isinstance(value, (list, tuple, set, dict)):
        return len(value) > 0
    return bool(value)


def _is_blank(value: Any) -> bool:
    if value is None:
        return True
    if isinstance(value, str):
        return value.strip() == ''
    if isinstance(value, (list, tuple)):
        return len(value) == 0
    return False


def _build_default_functions() -> Dict[str, Callable[..., Any]]:
    def func_if(condition: Any, true_value: Any, false_value: Any = None) -> Any:
        return true_value if _truthy(condition) else false_value

    def func_ifs(*args: Any) -> Any:
        if len(args) < 2 or len(args) % 2 != 0:
            raise FormulaError("IFS requires condition/value pairs.")
        for condition, value in zip(args[0::2], args[1::2]):
            if _truthy(condition):
                return value
        raise FormulaError("IFS did not match any condition.")

    def func_switch(expression: Any, *cases: Any) -> Any:
        if not cases:
            raise FormulaError("SWITCH requires at least one case.")
        default = None
        if len(cases) % 2 == 1:
            default = cases[-1]
            cases = cases[:-1]
        for case, value in zip(cases[0::2], cases[1::2]):
            if expression == case:
                return value
        if default is not None:
            return default
        raise FormulaError("SWITCH did not match any case.")

    def func_sum(*args: Any) -> float:
        total = 0.0
        for value in _flatten(args):
            total += FormulaEngine._to_number(value)
        return FormulaEngine._normalize_number(total)

    def func_average(*args: Any) -> float:
        values = [FormulaEngine._to_number(value) for value in _flatten(args)]
        if not values:
            raise FormulaError("AVERAGE requires at least one numeric value.")
        return FormulaEngine._normalize_number(sum(values) / len(values))

    def func_min(*args: Any) -> Any:
        values = [FormulaEngine._to_number(value) for value in _flatten(args)]
        if not values:
            raise FormulaError("MIN requires at least one numeric value.")
        return FormulaEngine._normalize_number(min(values))

    def func_max(*args: Any) -> Any:
        values = [FormulaEngine._to_number(value) for value in _flatten(args)]
        if not values:
            raise FormulaError("MAX requires at least one numeric value.")
        return FormulaEngine._normalize_number(max(values))

    def func_round(number: Any, digits: Any = 0) -> float:
        return _round_half_up(number, int(digits))

    def func_roundup(number: Any, digits: Any = 0) -> float:
        return _round_with_mode(number, int(digits), 'up')

    def func_rounddown(number: Any, digits: Any = 0) -> float:
        return _round_with_mode(number, int(digits), 'down')

    def func_len(text: Any) -> int:
        return len(_ensure_text(text))

    def func_concat(*args: Any) -> str:
        return ''.join(_flatten_text(args))

    def func_lower(text: Any) -> str:
        return _ensure_text(text).lower()

    def func_upper(text: Any) -> str:
        return _ensure_text(text).upper()

    def func_proper(text: Any) -> str:
        return ' '.join(word.capitalize() for word in _ensure_text(text).split())

    def func_trim(text: Any) -> str:
        return ' '.join(_ensure_text(text).split())

    def func_substitute(text: Any, old: Any, new: Any, occurrence: Any = None) -> str:
        source = _ensure_text(text)
        old_text = _ensure_text(old)
        new_text = _ensure_text(new)
        if occurrence is None:
            return source.replace(old_text, new_text)
        index = int(occurrence)
        if index <= 0:
            return source
        parts = source.split(old_text)
        if len(parts) <= index:
            return source
        rebuilt: List[str] = []
        for i, part in enumerate(parts[:-1], start=1):
            rebuilt.append(part)
            if i == index:
                rebuilt.append(new_text)
            else:
                rebuilt.append(old_text)
        rebuilt.append(parts[-1])
        return ''.join(rebuilt)

    def func_replace(text: Any, start: Any, length: Any, new_text: Any) -> str:
        source = _ensure_text(text)
        start_index = max(int(start) - 1, 0)
        count = max(int(length), 0)
        return source[:start_index] + _ensure_text(new_text) + source[start_index + count :]

    def func_left(text: Any, count: Any = 1) -> str:
        source = _ensure_text(text)
        return source[: max(int(count), 0)]

    def func_right(text: Any, count: Any = 1) -> str:
        source = _ensure_text(text)
        length = max(int(count), 0)
        if length == 0:
            return ''
        return source[-length:]

    def func_mid(text: Any, start: Any, length: Any) -> str:
        source = _ensure_text(text)
        start_index = max(int(start) - 1, 0)
        end_index = start_index + max(int(length), 0)
        return source[start_index:end_index]

    def func_search(needle: Any, haystack: Any, start: Any = 1) -> int:
        text = _ensure_text(haystack).lower()
        target = _ensure_text(needle).lower()
        position = text.find(target, max(int(start) - 1, 0))
        if position == -1:
            raise FormulaError("SEARCH could not find the specified text.")
        return position + 1

    def func_find(needle: Any, haystack: Any, start: Any = 1) -> int:
        text = _ensure_text(haystack)
        target = _ensure_text(needle)
        position = text.find(target, max(int(start) - 1, 0))
        if position == -1:
            raise FormulaError("FIND could not find the specified text.")
        return position + 1

    def func_split(text: Any, delimiter: Any) -> List[str]:
        return _ensure_text(text).split(_ensure_text(delimiter))

    def func_value(text: Any) -> float:
        return FormulaEngine._to_number(text)

    def func_to_text(value: Any) -> str:
        return _ensure_text(value)

    def func_text(value: Any, fmt: Any) -> str:
        return _format_text(value, _ensure_text(fmt))

    def func_now() -> datetime:
        return datetime.now()

    def func_today() -> date:
        return date.today()

    def func_date(year: Any, month: Any, day: Any) -> date:
        return date(int(year), int(month), int(day))

    def func_time(hour: Any, minute: Any, second: Any = 0) -> time:
        return time(int(hour), int(minute), int(second))

    def func_year(value: Any) -> int:
        return _normalize_datetime(value).year

    def func_month(value: Any) -> int:
        return _normalize_datetime(value).month

    def func_day(value: Any) -> int:
        return _normalize_datetime(value).day

    def func_hour(value: Any) -> int:
        return _normalize_datetime(value).hour

    def func_minute(value: Any) -> int:
        return _normalize_datetime(value).minute

    def func_second(value: Any) -> int:
        return _normalize_datetime(value).second

    def func_and(*args: Any) -> bool:
        return all(_truthy(value) for value in args)

    def func_or(*args: Any) -> bool:
        return any(_truthy(value) for value in args)

    def func_not(value: Any) -> bool:
        return not _truthy(value)

    def func_isnumber(value: Any) -> bool:
        try:
            FormulaEngine._to_number(value)
        except FormulaError:
            return False
        return True

    def func_istext(value: Any) -> bool:
        return isinstance(value, str)

    def func_isblank(value: Any) -> bool:
        return _is_blank(value)

    functions: Dict[str, Callable[..., Any]] = {
        'IF': func_if,
        'IFS': func_ifs,
        'SWITCH': func_switch,
        'SUM': func_sum,
        'AVERAGE': func_average,
        'MIN': func_min,
        'MAX': func_max,
        'ROUND': func_round,
        'ROUNDUP': func_roundup,
        'ROUNDDOWN': func_rounddown,
        'LEN': func_len,
        'CONCAT': func_concat,
        'CONCATENATE': func_concat,
        'LOWER': func_lower,
        'UPPER': func_upper,
        'PROPER': func_proper,
        'TRIM': func_trim,
        'SUBSTITUTE': func_substitute,
        'REPLACE': func_replace,
        'LEFT': func_left,
        'RIGHT': func_right,
        'MID': func_mid,
        'SEARCH': func_search,
        'FIND': func_find,
        'SPLIT': func_split,
        'VALUE': func_value,
        'TO_TEXT': func_to_text,
        'TEXT': func_text,
        'NOW': func_now,
        'TODAY': func_today,
        'DATE': func_date,
        'TIME': func_time,
        'YEAR': func_year,
        'MONTH': func_month,
        'DAY': func_day,
        'HOUR': func_hour,
        'MINUTE': func_minute,
        'SECOND': func_second,
        'AND': func_and,
        'OR': func_or,
        'NOT': func_not,
        'ISNUMBER': func_isnumber,
        'ISTEXT': func_istext,
        'ISBLANK': func_isblank,
    }
    return functions


FormulaEngine.FUNCTIONS.update(_build_default_functions())

