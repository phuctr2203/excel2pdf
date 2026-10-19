"""Excel number-format engine.

Turns a cell value plus a format code such as ``#,##0.00;[Red](#,##0.00)`` or
``dd/mm/yyyy hh:mm`` into the text Excel would display (en-US separators).
"""

from __future__ import annotations

import datetime as _dt
import math
import re
from dataclasses import dataclass
from decimal import ROUND_HALF_UP, Decimal
from functools import lru_cache
from typing import List, Optional, Tuple

RGB = Tuple[float, float, float]

_NAMED_COLORS = {
    "black": (0.0, 0.0, 0.0),
    "white": (1.0, 1.0, 1.0),
    "red": (1.0, 0.0, 0.0),
    "green": (0.0, 128 / 255, 0.0),
    "blue": (0.0, 0.0, 1.0),
    "yellow": (1.0, 1.0, 0.0),
    "magenta": (1.0, 0.0, 1.0),
    "cyan": (0.0, 1.0, 1.0),
}

MONTHS = ["January", "February", "March", "April", "May", "June", "July",
          "August", "September", "October", "November", "December"]
DAYS = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]

_EPOCH_1900 = _dt.datetime(1899, 12, 30)
_EPOCH_1904 = _dt.datetime(1904, 1, 1)


@dataclass
class Formatted:
    text: str
    color: Optional[RGB] = None
    # ``*x`` in a format repeats x to fill the cell; fill_at is the index in
    # ``text`` where the repetition goes (accounting formats use ``* ``).
    fill_char: Optional[str] = None
    fill_at: int = -1
    is_number: bool = False


# --------------------------------------------------------------------------
# Tokenizer
# --------------------------------------------------------------------------

@dataclass
class Tok:
    kind: str  # lit fill color cond elapsed date ampm general digit dot comma pct exp slash at fsec
    value: object = None


@dataclass
class Section:
    tokens: List[Tok]
    color: Optional[RGB] = None
    color_index: Optional[int] = None
    cond: Optional[Tuple[str, float]] = None
    is_date: bool = False
    has_ampm: bool = False
    is_text: bool = False
    has_digits: bool = False
    has_general: bool = False


def _split_sections(fmt: str) -> List[str]:
    out, cur, i, n = [], [], 0, len(fmt)
    while i < n:
        ch = fmt[i]
        if ch == '"':
            j = fmt.find('"', i + 1)
            j = n - 1 if j < 0 else j
            cur.append(fmt[i:j + 1])
            i = j + 1
            continue
        if ch == "\\" and i + 1 < n:
            cur.append(fmt[i:i + 2])
            i += 2
            continue
        if ch == "[":
            j = fmt.find("]", i)
            j = n - 1 if j < 0 else j
            cur.append(fmt[i:j + 1])
            i = j + 1
            continue
        if ch == ";":
            out.append("".join(cur))
            cur = []
            i += 1
            continue
        cur.append(ch)
        i += 1
    out.append("".join(cur))
    return out


_COND_RE = re.compile(r"^(<=|>=|<>|<|>|=)\s*(-?[\d.]+(?:[eE][-+]?\d+)?)$")


def _tokenize(src: str) -> Section:
    toks: List[Tok] = []
    sec = Section(tokens=toks)
    i, n = 0, len(src)
    lower = src.lower()
    while i < n:
        ch = src[i]
        lc = ch.lower()
        if ch == '"':
            j = src.find('"', i + 1)
            if j < 0:
                j = n
            toks.append(Tok("lit", src[i + 1:j]))
            i = j + 1
        elif ch == "\\":
            if i + 1 < n:
                toks.append(Tok("lit", src[i + 1]))
            i += 2
        elif ch == "_":
            toks.append(Tok("lit", " "))
            i += 2
        elif ch == "*":
            if i + 1 < n:
                toks.append(Tok("fill", src[i + 1]))
            i += 2
        elif ch == "[":
            j = src.find("]", i)
            if j < 0:
                j = n
            body = src[i + 1:j]
            bl = body.lower()
            i = j + 1
            if bl in _NAMED_COLORS:
                sec.color = _NAMED_COLORS[bl]
            elif bl.startswith("color") and bl[5:].isdigit():
                sec.color_index = int(bl[5:])
            elif body.startswith("$"):
                sym = body[1:].split("-", 1)[0]
                if sym:
                    toks.append(Tok("lit", sym))
            elif bl and set(bl) <= set("hms") and len(set(bl)) == 1:
                toks.append(Tok("elapsed", (bl[0], len(bl))))
                sec.is_date = True
            else:
                m = _COND_RE.match(body.strip())
                if m:
                    sec.cond = (m.group(1), float(m.group(2)))
                # [DBNum1], [natNum], locale ids... ignored
        elif lower.startswith("general", i):
            toks.append(Tok("general"))
            sec.has_general = True
            i += 7
        elif lower.startswith("am/pm", i):
            toks.append(Tok("ampm", "AM/PM" if src[i] == "A" else "am/pm"))
            sec.has_ampm = True
            sec.is_date = True
            i += 5
        elif lower.startswith("a/p", i):
            toks.append(Tok("ampm", "A/P" if src[i] == "A" else "a/p"))
            sec.has_ampm = True
            sec.is_date = True
            i += 3
        elif lc in "ymdhseg" and not (lc == "e" and i + 1 < n and src[i + 1] in "+-"):
            j = i
            while j < n and src[j].lower() == lc:
                j += 1
            toks.append(Tok("date", (lc, j - i)))
            sec.is_date = True
            i = j
        elif lc == "b" and i + 1 < n and src[i + 1] in "12":
            i += 2  # Buddhist calendar marker: ignore
        elif ch in "0#?":
            toks.append(Tok("digit", ch))
            sec.has_digits = True
            i += 1
        elif ch == ".":
            # fractional seconds: ".0", ".00" right after a seconds token
            j = i + 1
            while j < n and src[j] == "0":
                j += 1
            prev = _last_date_tok(toks)
            if j > i + 1 and prev is not None and prev.kind in ("date", "elapsed") and prev.value[0] == "s":
                toks.append(Tok("fsec", j - i - 1))
                i = j
            else:
                toks.append(Tok("dot"))
                i += 1
        elif ch == ",":
            toks.append(Tok("comma"))
            i += 1
        elif ch == "%":
            toks.append(Tok("pct"))
            i += 1
        elif ch in "eE" and i + 1 < n and src[i + 1] in "+-":
            toks.append(Tok("exp", src[i + 1]))
            i += 2
        elif ch == "/":
            toks.append(Tok("slash"))
            i += 1
        elif ch == "@":
            toks.append(Tok("at"))
            sec.is_text = True
            i += 1
        else:
            toks.append(Tok("lit", ch))
            i += 1

    if sec.is_date:
        _resolve_minutes(toks)
        if sec.has_digits and not any(t.kind == "date" for t in toks):
            pass
    return sec


def _last_date_tok(toks: List[Tok]) -> Optional[Tok]:
    for t in reversed(toks):
        if t.kind in ("date", "elapsed"):
            return t
    return None


def _resolve_minutes(toks: List[Tok]) -> None:
    """'m' means minutes after an hour token or before a seconds token."""
    dated = [t for t in toks if t.kind in ("date", "elapsed")]
    for k, t in enumerate(dated):
        if t.kind != "date" or t.value[0] != "m" or t.value[1] > 2:
            continue
        prev = dated[k - 1] if k > 0 else None
        nxt = dated[k + 1] if k + 1 < len(dated) else None
        if (prev is not None and prev.value[0] == "h") or (nxt is not None and nxt.value[0] == "s"):
            t.value = ("M", t.value[1])  # capital M = minutes


@lru_cache(maxsize=2048)
def parse_format(fmt: str) -> Tuple[Section, ...]:
    if not fmt:
        fmt = "General"
    return tuple(_tokenize(s) for s in _split_sections(fmt))


# --------------------------------------------------------------------------
# Value conversion
# --------------------------------------------------------------------------

def to_serial(value, date1904: bool = False) -> Optional[float]:
    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    if isinstance(value, _dt.datetime):
        epoch = _EPOCH_1904 if date1904 else _EPOCH_1900
        delta = value.replace(tzinfo=None) - epoch
        serial = delta.days + delta.seconds / 86400.0 + delta.microseconds / 86400e6
        if not date1904 and serial < 61:
            serial -= 1  # Excel's fictitious 1900-02-29
        return serial
    if isinstance(value, _dt.date):
        return to_serial(_dt.datetime(value.year, value.month, value.day), date1904)
    if isinstance(value, _dt.time):
        return (value.hour * 3600 + value.minute * 60 + value.second + value.microsecond / 1e6) / 86400.0
    if isinstance(value, _dt.timedelta):
        return value.total_seconds() / 86400.0
    return None


def _serial_to_datetime(serial: float, date1904: bool) -> _dt.datetime:
    days = math.floor(serial)
    frac = serial - days
    if date1904:
        base = _EPOCH_1904 + _dt.timedelta(days=days)
    elif days >= 61:
        base = _EPOCH_1900 + _dt.timedelta(days=days)
    elif days == 60:
        base = _dt.datetime(1900, 2, 28)  # Excel's fictitious 1900-02-29
    else:
        base = _dt.datetime(1899, 12, 31) + _dt.timedelta(days=days)
    return base + _dt.timedelta(seconds=frac * 86400.0)


def _round15(x: float) -> float:
    if x == 0 or not math.isfinite(x):
        return x
    return float(f"{x:.15g}")


# --------------------------------------------------------------------------
# Public API
# --------------------------------------------------------------------------

def format_value(value, fmt: Optional[str], date1904: bool = False, indexed_color=None) -> Formatted:
    """Format ``value`` with Excel format code ``fmt``.

    ``indexed_color`` is an optional callable mapping ``[ColorN]`` to RGB.
    """
    if value is None:
        return Formatted("")
    if isinstance(value, bool):
        return Formatted("TRUE" if value else "FALSE")
    sections = parse_format(fmt or "General")

    if isinstance(value, str):
        sec = _text_section(sections)
        if sec is None:
            return Formatted(value)
        return _render_text(value, sec, indexed_color)

    serial = to_serial(value, date1904)
    if serial is None:
        return Formatted(str(value))
    if not math.isfinite(serial):
        return Formatted("#NUM!")
    serial = _round15(serial)

    sec, use_abs, explicit_sign = _pick_section(sections, serial)
    color = sec.color
    if sec.color_index is not None and indexed_color is not None:
        color = indexed_color(sec.color_index + 7)

    if sec.is_date:
        if serial < 0:
            return Formatted("#" * 10, color, is_number=True)
        text, fill_char, fill_at = _render_date(serial, sec, date1904)
        return Formatted(text, color, fill_char, fill_at, is_number=True)

    x = abs(serial) if use_abs else serial
    negative = x < 0
    x = abs(x)
    text, fill_char, fill_at = _render_number(x, sec)
    if negative and explicit_sign:
        # Excel keeps the sign even when the value rounds to zero ("-0").
        text = "-" + text
        if fill_at >= 0:
            fill_at += 1
    return Formatted(text, color, fill_char, fill_at, is_number=True)


def _text_section(sections: Tuple[Section, ...]) -> Optional[Section]:
    if len(sections) >= 4:
        return sections[3]
    for s in sections:
        if s.is_text:
            return s
    return None


def _render_text(value: str, sec: Section, indexed_color) -> Formatted:
    out: List[str] = []
    fill_char, fill_at = None, -1
    for t in sec.tokens:
        if t.kind == "at":
            out.append(value)
        elif t.kind == "lit":
            out.append(t.value)  # type: ignore[arg-type]
        elif t.kind == "fill":
            fill_char, fill_at = t.value, len("".join(out))  # type: ignore[assignment]
    if not any(t.kind == "at" for t in sec.tokens) and not out:
        out.append(value)
    color = sec.color
    if sec.color_index is not None and indexed_color is not None:
        color = indexed_color(sec.color_index + 7)
    return Formatted("".join(out), color, fill_char, fill_at)


def _cond_match(cond: Tuple[str, float], x: float) -> bool:
    op, v = cond
    return {
        "<": x < v, "<=": x <= v, ">": x > v, ">=": x >= v, "=": x == v, "<>": x != v,
    }[op]


def _pick_section(sections: Tuple[Section, ...], x: float) -> Tuple[Section, bool, bool]:
    """Return (section, use_absolute_value, show_minus_sign)."""
    num = [s for i, s in enumerate(sections) if i < 3 and not (s.is_text and not s.has_digits and not s.is_date and not s.has_general)]
    if not num:
        num = [parse_format("General")[0]]
    if any(s.cond for s in num):
        for s in num:
            if s.cond is None or _cond_match(s.cond, x):
                return s, False, True
        return num[-1], False, True
    if len(num) == 1:
        return num[0], False, True
    if x > 0 or (x == 0 and len(num) < 3):
        return num[0], False, True
    if x < 0:
        return num[1], True, False
    return num[2], False, True


# --------------------------------------------------------------------------
# General format
# --------------------------------------------------------------------------

def format_general(x: float, max_chars: int = 11) -> str:
    if x == 0:
        return "0"
    neg = x < 0
    a = abs(x)
    avail = max_chars - (1 if neg else 0)
    if a.is_integer() and a < 10 ** avail:
        s = str(int(a))
    else:
        s = None
        if 1e-9 <= a < 10 ** avail:
            int_digits = max(1, int(math.floor(math.log10(a))) + 1)
            decimals = avail - int_digits - 1
            if decimals >= 0:
                cand = _fixed(a, decimals)
                if "." in cand:
                    cand = cand.rstrip("0").rstrip(".")
                if cand not in ("0", "") and _sig_ok(cand, a):
                    s = cand
        if s is None:
            s = _general_sci(a, avail)
    return ("-" if neg else "") + s


def _sig_ok(text: str, a: float) -> bool:
    # Prefer scientific when fixed notation lost every significant digit or
    # the number is very small (Excel switches below 1E-4 when digits drop).
    try:
        v = float(text)
    except ValueError:
        return False
    return v != 0 and a >= 1e-4


def _general_sci(a: float, avail: int) -> str:
    exp = int(math.floor(math.log10(a)))
    mant = a / 10 ** exp
    # mantissa digits so the result fits: d.ddddE+XX
    exp_len = 4 if abs(exp) < 100 else 5
    decimals = max(0, avail - exp_len - 2)
    mstr = _fixed(mant, decimals)
    if float(mstr) >= 10:
        exp += 1
        mstr = _fixed(a / 10 ** exp, decimals)
    if "." in mstr:
        mstr = mstr.rstrip("0").rstrip(".")
    sign = "+" if exp >= 0 else "-"
    return f"{mstr}E{sign}{abs(exp):02d}"


def _fixed(x: float, decimals: int) -> str:
    q = Decimal(repr(x)).quantize(Decimal(1).scaleb(-decimals), rounding=ROUND_HALF_UP)
    return format(q, "f")


# --------------------------------------------------------------------------
# Numbers
# --------------------------------------------------------------------------

def _render_number(x: float, sec: Section) -> Tuple[str, Optional[str], int]:
    toks = sec.tokens
    if sec.has_general and not sec.has_digits:
        out: List[str] = []
        fill_char, fill_at = None, -1
        for t in toks:
            if t.kind == "general":
                out.append(format_general(x))
            elif t.kind == "lit":
                out.append(t.value)  # type: ignore[arg-type]
            elif t.kind == "fill":
                fill_char, fill_at = t.value, len("".join(out))  # type: ignore[assignment]
            elif t.kind == "pct":
                out.append("%")
        return "".join(out), fill_char, fill_at
    if not sec.has_digits:
        # literal-only section, e.g. "-" for zero in accounting formats
        out = []
        fill_char, fill_at = None, -1
        for t in toks:
            if t.kind == "lit":
                out.append(t.value)  # type: ignore[arg-type]
            elif t.kind == "fill":
                fill_char, fill_at = t.value, len("".join(out))  # type: ignore[assignment]
            elif t.kind == "pct":
                out.append("%")
            elif t.kind == "general":
                out.append(format_general(x))
        return "".join(out), fill_char, fill_at

    pct = sum(1 for t in toks if t.kind == "pct")
    x = x * (100 ** pct)

    if any(t.kind == "slash" for t in toks):
        res = _render_fraction(x, toks)
        if res is not None:
            return res

    exp_idx = next((i for i, t in enumerate(toks) if t.kind == "exp"), None)
    dot_idx = next((i for i, t in enumerate(toks) if t.kind == "dot"), None)
    mant_end = exp_idx if exp_idx is not None else len(toks)
    int_end = dot_idx if dot_idx is not None and dot_idx < mant_end else mant_end

    int_ph = [i for i in range(int_end) if toks[i].kind == "digit"]
    frac_ph = [i for i in range(int_end, mant_end) if toks[i].kind == "digit"]
    exp_ph = [i for i in range(mant_end, len(toks)) if toks[i].kind == "digit"] if exp_idx is not None else []

    # thousands separators vs scaling commas
    grouping = False
    scale_commas = 0
    last_int_digit = int_ph[-1] if int_ph else -1
    for i, t in enumerate(toks[:mant_end]):
        if t.kind != "comma":
            continue
        if i < int_end and int_ph and int_ph[0] < i < last_int_digit:
            grouping = True
        elif (int_ph or frac_ph) and i > (frac_ph[-1] if frac_ph and i > int_end else last_int_digit):
            scale_commas += 1
    x = x / (1000 ** scale_commas)

    exponent = 0
    if exp_idx is not None:
        if x != 0:
            n_int = max(1, len(int_ph))
            exponent = int(math.floor(math.log10(x)))
            if len(int_ph) > 1 and any(toks[i].value == "#" for i in int_ph):
                exponent = int(math.floor(exponent / n_int) * n_int)  # engineering
            else:
                exponent -= n_int - 1
            mant = x / (10 ** exponent)
            if float(_fixed(mant, len(frac_ph))) >= 10 ** n_int:
                exponent += n_int if len(int_ph) > 1 and any(toks[i].value == "#" for i in int_ph) else 1
                mant = x / (10 ** exponent)
            x = mant

    num = _fixed(x, len(frac_ph))
    int_str, _, frac_str = num.partition(".")
    if int_str == "0":
        int_str = ""

    out: List[Optional[str]] = [None] * len(toks)

    # integer digits, right to left
    out_prefix = ""
    digits = list(int_str)
    for k, ti in enumerate(reversed(int_ph)):
        ph = toks[ti].value
        if digits:
            d = digits.pop()
            out[ti] = d
        else:
            out[ti] = "0" if ph == "0" else (" " if ph == "?" else "")
    if digits and int_ph:
        out[int_ph[0]] = "".join(digits) + (out[int_ph[0]] or "")
    elif digits and not int_ph:
        # no integer placeholders: integer digits go before the decimal point
        out_prefix = "".join(digits)

    if grouping and int_ph:
        # rebuild integer segment with separators (assumes no literals inside)
        seg = "".join(out[i] or "" for i in int_ph)
        core = seg.lstrip(" ")
        pad = seg[: len(seg) - len(core)]
        if core:
            core = f"{int(core):,}" if core.isdigit() else core
        for i in int_ph:
            out[i] = ""
        out[int_ph[0]] = pad + core

    # fraction digits, left to right, trimming '#'/'?' trailing zeros
    fdigits = list(frac_str)
    for k, ti in enumerate(frac_ph):
        out[ti] = fdigits[k] if k < len(fdigits) else "0"
    for ti in reversed(frac_ph):
        ph = toks[ti].value
        if out[ti] != "0" or ph == "0":
            break
        out[ti] = "" if ph == "#" else " "

    # exponent
    if exp_idx is not None:
        sign = "-" if exponent < 0 else ("+" if toks[exp_idx].value == "+" else "")
        min_len = sum(1 for i in exp_ph if toks[i].value == "0")
        estr = str(abs(exponent)).rjust(max(min_len, 1), "0")
        out[exp_idx] = "E" + sign
        for i in exp_ph:
            out[i] = ""
        if exp_ph:
            out[exp_ph[0]] = estr

    text: List[str] = []
    fill_char, fill_at = None, -1
    for i, t in enumerate(toks):
        if t.kind in ("digit", "exp"):
            if out[i]:
                text.append(out[i])  # type: ignore[arg-type]
        elif t.kind == "dot":
            if out_prefix and not int_ph:
                text.append(out_prefix)
                out_prefix = ""
            text.append(".")
        elif t.kind == "lit":
            text.append(t.value)  # type: ignore[arg-type]
        elif t.kind == "pct":
            text.append("%")
        elif t.kind == "fill":
            fill_char, fill_at = t.value, len("".join(text))  # type: ignore[assignment]
        elif t.kind == "general":
            text.append(format_general(x))
        elif t.kind == "comma":
            pass
    s = "".join(text)
    if out_prefix:
        s = out_prefix + s
    return s, fill_char, fill_at


def _render_fraction(x: float, toks: List[Tok]) -> Optional[Tuple[str, Optional[str], int]]:
    si = next(i for i, t in enumerate(toks) if t.kind == "slash")
    # numerator: contiguous digit tokens right before the slash
    j = si - 1
    while j >= 0 and toks[j].kind == "digit":
        j -= 1
    num_ph = list(range(j + 1, si))
    if not num_ph:
        return None
    int_ph = [i for i in range(0, j + 1) if toks[i].kind == "digit"]
    # denominator: digit placeholders or a literal number
    k = si + 1
    den_ph: List[int] = []
    fixed_den = None
    while k < len(toks) and toks[k].kind == "digit":
        den_ph.append(k)
        k += 1
    if not den_ph:
        lit = ""
        while k < len(toks) and toks[k].kind == "lit" and str(toks[k].value).isdigit():
            lit += str(toks[k].value)
            k += 1
        if not lit:
            return None
        fixed_den = int(lit)
    end = k

    whole = math.floor(x) if int_ph else 0
    frac = x - whole
    if fixed_den:
        den = fixed_den
        numer = int(Decimal(repr(frac * den)).quantize(Decimal(1), rounding=ROUND_HALF_UP))
    else:
        numer, den = _best_fraction(frac, 10 ** len(den_ph) - 1)
    if int_ph and numer == den:
        whole += 1
        numer = 0

    def pad(value: str, ph: List[int], left: bool) -> str:
        need = len(ph)
        if len(value) >= need:
            return value
        fillc = " " if any(toks[i].value == "?" for i in ph) else ""
        return (fillc * (need - len(value)) + value) if left else (value + fillc * (need - len(value)))

    text: List[str] = []
    fill_char, fill_at = None, -1
    for i, t in enumerate(toks[: num_ph[0]]):
        if i in int_ph:
            if i == int_ph[-1]:
                if whole == 0 and numer != 0:
                    ws = "" if all(toks[p].value == "#" for p in int_ph) else ("0" if any(toks[p].value == "0" for p in int_ph) else " ")
                else:
                    ws = str(whole)
                text.append(ws)
        elif t.kind == "lit":
            text.append(str(t.value))
        elif t.kind == "fill":
            fill_char, fill_at = str(t.value), len("".join(text))
    if numer == 0 and int_ph:
        blank = len(num_ph) + 1 + (len(den_ph) if den_ph else len(str(fixed_den)))
        text.append(" " * blank)
    else:
        text.append(pad(str(numer), num_ph, True))
        text.append("/")
        text.append(pad(str(den), den_ph, False) if den_ph else str(den))
    for t in toks[end:]:
        if t.kind == "lit":
            text.append(str(t.value))
        elif t.kind == "pct":
            text.append("%")
    s = "".join(text)
    if not int_ph and numer == 0:
        s = "0"
    return s, fill_char, fill_at


def _best_fraction(f: float, max_den: int) -> Tuple[int, int]:
    best = (0, 1)
    best_err = abs(f)
    for d in range(1, max_den + 1):
        n = round(f * d)
        err = abs(f - n / d)
        if err < best_err - 1e-12:
            best, best_err = (n, d), err
            if err < 1e-12:
                break
    return best


# --------------------------------------------------------------------------
# Dates and times
# --------------------------------------------------------------------------

def _render_date(serial: float, sec: Section, date1904: bool) -> Tuple[str, Optional[str], int]:
    toks = sec.tokens
    fdigits = max((t.value for t in toks if t.kind == "fsec"), default=0)
    # round to the displayed precision
    unit = 86400 * (10 ** fdigits)
    serial = round(serial * unit) / unit
    dt = _serial_to_datetime(serial, date1904)
    total_seconds = serial * 86400.0

    out: List[str] = []
    fill_char, fill_at = None, -1
    for t in toks:
        k = t.kind
        if k == "lit":
            out.append(str(t.value))
        elif k == "fill":
            fill_char, fill_at = str(t.value), len("".join(out))
        elif k == "elapsed":
            unit_c, ln = t.value  # type: ignore[misc]
            if unit_c == "h":
                v = int(math.floor(total_seconds / 3600 + 1e-9))
            elif unit_c == "m":
                v = int(math.floor(total_seconds / 60 + 1e-9))
            else:
                v = int(math.floor(total_seconds + 1e-9))
            out.append(str(v).rjust(ln, "0"))
        elif k == "date":
            c, ln = t.value  # type: ignore[misc]
            out.append(_date_part(c, ln, dt, sec.has_ampm, serial))
        elif k == "ampm":
            am = dt.hour < 12
            style = str(t.value)
            if style in ("AM/PM", "am/pm"):
                txt = "AM" if am else "PM"
                out.append(txt if style[0] == "A" else txt.lower())
            else:
                txt = "A" if am else "P"
                out.append(txt if style[0] == "A" else txt.lower())
        elif k == "fsec":
            frac = round(dt.microsecond / 1e6, int(t.value))  # type: ignore[arg-type]
            out.append("." + f"{frac:.{t.value}f}".split(".")[1])
        elif k == "digit":
            out.append("0" if t.value == "0" else "")
        elif k == "dot":
            out.append(".")
        elif k == "general":
            out.append(format_general(serial))
        elif k == "slash":
            out.append("/")
        elif k == "comma":
            out.append(",")
        elif k == "pct":
            out.append("%")
    return "".join(out), fill_char, fill_at


def _date_part(c: str, ln: int, dt: _dt.datetime, ampm: bool, serial: float) -> str:
    if c in ("y", "e"):
        if c == "e" or ln > 2:
            return f"{dt.year:04d}"
        return f"{dt.year % 100:02d}"
    if c == "m":  # month
        if ln == 1:
            return str(dt.month)
        if ln == 2:
            return f"{dt.month:02d}"
        if ln == 3:
            return MONTHS[dt.month - 1][:3]
        if ln == 5:
            return MONTHS[dt.month - 1][0]
        return MONTHS[dt.month - 1]
    if c == "M":  # minutes
        return str(dt.minute) if ln == 1 else f"{dt.minute:02d}"
    if c == "d":
        if ln == 1:
            return str(dt.day)
        if ln == 2:
            return f"{dt.day:02d}"
        wd = dt.weekday()
        if serial < 1:
            wd = 5  # serial 0 is "Saturday, January 0, 1900"
        return DAYS[wd][:3] if ln == 3 else DAYS[wd]
    if c == "h":
        h = dt.hour
        if ampm:
            h = h % 12 or 12
        return str(h) if ln == 1 else f"{h:02d}"
    if c == "s":
        return str(dt.second) if ln == 1 else f"{dt.second:02d}"
    return ""


def is_date_format(fmt: Optional[str]) -> bool:
    if not fmt:
        return False
    return any(s.is_date for s in parse_format(fmt)[:1])
