import json
import re
from pathlib import Path

from .errors import ReportError
from .numbers import UNITS, amount

ID = r"[A-Za-z][A-Za-z0-9_.]*"
CELL = r"[A-Z]{1,3}[1-9][0-9]*"


def require(condition, message):
    if not condition:
        raise ReportError("配置错误：" + message)


def keys(obj, allowed, required, where):
    require(isinstance(obj, dict), f"{where} 必须是对象")
    require(not set(obj) - set(allowed), f"{where} 有未知键 {set(obj) - set(allowed)}")
    require(set(required) <= set(obj), f"{where} 缺少 {set(required) - set(obj)}")


def integer(x, low, high):
    return type(x) is int and low <= x <= high


def validate(c):
    keys(c, ["version", "source", "metadata", "fields", "checks", "format"],
         ["version", "source", "metadata", "fields", "checks"], "根对象")
    require(type(c["version"]) is int and c["version"] == 1, "version 必须为 1")
    s = c["source"]
    keys(s, ["sheet", "unit", "blocks"], ["sheet", "unit", "blocks"], "source")
    require(isinstance(s["sheet"], str) and s["sheet"], "sheet 不能为空")
    require(s["unit"] in UNITS, "source.unit 只支持元、千元、万元")
    require(isinstance(s["blocks"], list) and s["blocks"], "blocks 不能为空")
    block_ids = set()
    for b in s["blocks"]:
        keys(b, ["id", "label_column", "start_row", "end_row", "period_columns"],
             ["id", "label_column", "start_row", "end_row", "period_columns"], "block")
        require(isinstance(b["id"], str) and b["id"] not in block_ids, "block.id 必须唯一")
        block_ids.add(b["id"])
        require(re.fullmatch(r"[A-Z]{1,3}", b["label_column"]) is not None, "科目列须为列字母")
        require(integer(b["start_row"], 1, 1048576) and integer(b["end_row"], b["start_row"], b["start_row"] + 50000), "扫描行范围无效或超过五万行")
        require(isinstance(b["period_columns"], dict) and b["period_columns"], "期间列不能为空")
        require(all(isinstance(k, str) and isinstance(v, str) and re.fullmatch(r"[A-Z]{1,3}", v) for k, v in b["period_columns"].items()), "期间列无效")
    require(isinstance(c["metadata"], list) and c["metadata"], "metadata 至少需要一个来源标识校验")
    for m in c["metadata"]:
        keys(m, ["cell", "expected"], ["cell", "expected"], "metadata")
        require(isinstance(m["cell"], str) and re.fullmatch(CELL, m["cell"]), "metadata.cell 无效")
        require(isinstance(m["expected"], str) and m["expected"], "metadata.expected 必须是非空文本")
    require(isinstance(c["fields"], list) and c["fields"], "fields 不能为空")
    require(any(isinstance(f, dict) and "target" in f for f in c["fields"]), "至少一个字段须设置 Word target")
    field_ids = set()
    blocks = {b["id"]: b for b in s["blocks"]}
    for f in c["fields"]:
        keys(f, ["id", "source", "target", "blank_zero", "min", "max"], ["id", "source"], "field")
        require(isinstance(f["id"], str) and re.fullmatch(ID, f["id"]) and f["id"] not in field_ids, "field.id 无效或重复")
        field_ids.add(f["id"])
        src = f["source"]
        keys(src, ["block", "labels", "period", "cell", "label_cell"], [], f"{f['id']}.source")
        require(isinstance(src.get("labels"), list) and src["labels"] and all(isinstance(x, str) and x.strip() for x in src["labels"]), "每个来源必须配置 labels，用于科目匹配或坐标防漂移")
        if "cell" in src:
            require(set(src) == {"cell", "label_cell", "labels"}, "坐标来源需要 cell、label_cell、labels，不能混用 block")
            require(all(isinstance(src[k], str) and re.fullmatch(CELL, src[k]) for k in ["cell", "label_cell"]), "来源坐标无效")
        else:
            require(set(src) == {"block", "labels", "period"}, "科目来源需要 block、labels、period")
            require(src["block"] in blocks and src["period"] in blocks[src["block"]]["period_columns"], "未知 block 或 period")
        require(type(f.get("blank_zero", False)) is bool, "blank_zero 必须是布尔值")
        for bound in ["min", "max"]:
            if bound in f:
                amount(f[bound])
        if "min" in f and "max" in f:
            require(amount(f["min"]) <= amount(f["max"]), "min 不得大于 max")
        if "target" in f:
            t = f["target"]
            keys(t, ["placeholder", "occurrences", "table", "row", "cell", "expected_text"], [], "target")
            if "placeholder" in t:
                require(set(t) <= {"placeholder", "occurrences"}, "占位符不能混用表格定位")
                require(isinstance(t["placeholder"], str) and re.fullmatch(ID, t["placeholder"]), "placeholder 无效")
                require(integer(t.get("occurrences", 1), 1, 10000), "occurrences 必须为正整数")
            else:
                require(set(t) == {"table", "row", "cell", "expected_text"}, "表格目标须有 table、row、cell、expected_text")
                require(all(integer(t[k], 0, 10000) for k in ["table", "row", "cell"]), "Word 索引须从 0 开始")
                require(isinstance(t["expected_text"], str), "expected_text 必须是文本")
    require(isinstance(c["checks"], list) and c["checks"], "至少需要一个勾稽检查")
    for check in c["checks"]:
        keys(check, ["name", "left", "right", "tolerance"], ["name", "left", "right", "tolerance"], "check")
        require(isinstance(check["name"], str) and check["name"], "检查名称不能为空")
        for side in ["left", "right"]:
            require(isinstance(check[side], list) and check[side] and len(set(check[side])) == len(check[side]), "检查两边须为不重复的字段列表")
            require(set(check[side]) <= field_ids, "勾稽检查引用了未知字段")
        require(amount(check["tolerance"]) >= 0, "容差不能为负，单位与 Excel 一致")
    fmt = c.get("format", {})
    keys(fmt, ["unit", "decimals", "grouping", "negative"], [], "format")
    require(fmt.get("unit", s["unit"]) in UNITS, "输出单位无效")
    require(integer(fmt.get("decimals", 2), 0, 8), "decimals 必须为 0 至 8")
    require(type(fmt.get("grouping", True)) is bool, "grouping 必须是布尔值")
    require(fmt.get("negative", "minus") in ["minus", "parentheses"], "negative 无效")
    return c


def load_config(path):
    def unique(pairs):
        d = {}
        for k, v in pairs:
            if k in d:
                raise ReportError(f"JSON 键重复：{k}")
            d[k] = v
        return d
    try:
        c = json.loads(Path(path).read_text(encoding="utf-8-sig"), object_pairs_hook=unique)
        return validate(c)
    except ReportError:
        raise
    except (ValueError, TypeError, KeyError) as exc:
        raise ReportError(f"配置文件无效：{exc}") from exc
