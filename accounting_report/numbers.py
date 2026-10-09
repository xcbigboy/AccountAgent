import re
import unicodedata
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP, localcontext

from .errors import ReportError

UNITS = {"元": Decimal(1), "千元": Decimal(1000), "万元": Decimal(10000)}


def normalize_label(value):
    return re.sub(r"\s+", "", unicodedata.normalize("NFKC", str(value))).rstrip(":")


def amount(value, blank_zero=False):
    if value is None or (isinstance(value, str) and not value.strip()):
        if blank_zero:
            return Decimal(0)
        raise ReportError("金额为空，不能自动视为零")
    if isinstance(value, bool) or not isinstance(value, (str, int, float, Decimal)):
        raise ReportError(f"不是金额类型：{value!r}")
    text = unicodedata.normalize("NFKC", str(value)).strip()
    negative = text.startswith("(") and text.endswith(")")
    if negative:
        text = text[1:-1].strip()
    text = re.sub(r"^[¥￥]", "", text)
    pattern = r"[+-]?(?:\d+|\d{1,3}(?:,\d{3})+)(?:\.\d+)?"
    if not re.fullmatch(pattern, text) or (negative and text.startswith(("+", "-"))):
        raise ReportError(f"无效金额：{value!r}；不接受百分号、横杠、科学计数文本或混合单位")
    try:
        result = Decimal(text.replace(",", ""))
    except InvalidOperation as exc:
        raise ReportError(f"无效金额：{value!r}") from exc
    if not result.is_finite() or len(result.as_tuple().digits) > 28:
        raise ReportError("金额不是有限数值，或超过 28 位精度")
    return -result if negative else result


def display(value, source_unit, target_unit, decimals=2, grouping=True, negative="minus"):
    with localcontext() as ctx:
        ctx.prec = 60
        converted = value * UNITS[source_unit] / UNITS[target_unit]
        rounded = converted.quantize(Decimal(1).scaleb(-decimals), rounding=ROUND_HALF_UP)
        if rounded == 0:
            rounded = abs(rounded)
        text = format(abs(rounded), f"{',' if grouping else ''}.{decimals}f")
        if rounded < 0:
            text = f"({text})" if negative == "parentheses" else "-" + text
        return text, converted, rounded
