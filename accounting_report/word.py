import re
from collections import Counter
from zipfile import ZipFile
from lxml import etree

from .errors import ReportError

NS = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
W = "{" + NS["w"] + "}"
TOKEN = re.compile(r"\{\{([A-Za-z][A-Za-z0-9_.]*)\}\}")
PART = re.compile(r"word/(document|header\d+|footer\d+)\.xml$")


def parse(data):
    return etree.fromstring(data, etree.XMLParser(resolve_entities=False, no_network=True))


def texts(paragraph):
    # A text box may contain nested paragraphs; never count their text twice.
    nodes = []
    for n in paragraph.iter():
        if n.tag not in [W + "t", W + "tab", W + "br", W + "cr"]:
            continue
        parent = n.getparent()
        while parent is not None and parent.tag != W + "p":
            parent = parent.getparent()
        if parent is paragraph:
            nodes.append(n)
    return nodes


def node_text(node):
    return (node.text or "") if node.tag == W + "t" else "\n"


def replace_span(nodes, start, end, value):
    position = 0
    inserted = False
    for node in nodes:
        text = node_text(node)
        finish = position + len(text)
        if finish > start and position < end:
            if node.tag != W + "t":
                raise ReportError("占位符跨越换行或制表符，不能安全填报")
            a, b = max(0, start - position), min(len(text), end - position)
            node.text = text[:a] + (value if not inserted else "") + text[b:]
            node.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
            inserted = True
        position = finish


class WordTemplate:
    def __init__(self, path):
        with ZipFile(path) as archive:
            infos = archive.infolist()
            if len({i.filename for i in infos}) != len(infos):
                raise ReportError("Word 压缩包有重复成员")
            if sum(i.file_size for i in infos) > 100 * 1024 * 1024:
                raise ReportError("Word 解压内容超过 100 MB")
            self.files = [(i, archive.read(i.filename)) for i in infos]
        self.roots = {i.filename: parse(data) for i, data in self.files if PART.fullmatch(i.filename)}
        if "word/document.xml" not in self.roots:
            raise ReportError("不是有效的 DOCX 文档")
        for name, root in self.roots.items():
            if root.xpath("//w:ins | //w:del | //w:moveFrom | //w:moveTo", namespaces=NS):
                raise ReportError(f"{name} 含修订，请接受或拒绝修订后重试")
        for i, data in self.files:
            if i.filename == "word/settings.xml":
                root = parse(data)
                if any(n.get(W + "enforcement") not in ["0", "false", "off"] for n in root.findall("w:documentProtection", NS)):
                    raise ReportError("Word 文档启用了编辑保护，请解除保护后重试")
            if i.filename.startswith("word/") and i.filename.endswith(".xml") and i.filename not in self.roots:
                root = parse(data)
                paragraph_texts = ["".join(node_text(n) for n in texts(p)) for p in root.iter(W + "p")]
                if b"{{" in data or any("{{" in text or "}}" in text for text in paragraph_texts):
                    raise ReportError(f"占位符位于未支持的部分：{i.filename}")
        self.modified = set()

    def paragraphs(self):
        for name, root in self.roots.items():
            for p in root.iter(W + "p"):
                yield name, p

    def inventory(self):
        result = []
        for name, p in self.paragraphs():
            text = "".join(node_text(n) for n in texts(p))
            if TOKEN.search(text) and p.xpath(".//w:fldSimple | .//w:fldChar | .//w:instrText", namespaces=NS):
                raise ReportError(f"{name} 占位符所在段落含 Word 域；请把占位符移至普通文本段落")
            for m in TOKEN.finditer(text):
                result.append({"placeholder": m.group(1), "part": name, "paragraph": text})
        root = self.roots["word/document.xml"]
        cells = []
        for ti, table in enumerate(root.findall("w:body/w:tbl", NS)):
            for ri, row in enumerate(table.findall("w:tr", NS)):
                for ci, cell in enumerate(row.findall("w:tc", NS)):
                    cells.append({"table": ti, "row": ri, "cell": ci,
                                  "text": "".join(cell.xpath(".//w:t/text()", namespaces=NS))})
        return {"placeholders": result, "table_cells": cells}

    def fill(self, config, records):
        values = {r["id"]: r["formatted"] for r in records}
        replacements, expected, logs = {}, {}, []
        cells = set()
        for f in config["fields"]:
            t = f.get("target")
            if t is None:
                continue
            if "placeholder" in t:
                token = t["placeholder"]
                if token in replacements:
                    raise ReportError(f"Word 目标重复：{token}")
                replacements[token] = values[f["id"]]
                expected[token] = t.get("occurrences", 1)
            else:
                index = (t["table"], t["row"], t["cell"])
                if index in cells:
                    raise ReportError(f"Word 单元格目标重复：{index}")
                cells.add(index)
                try:
                    tables = self.roots["word/document.xml"].findall("w:body/w:tbl", NS)
                    cell = tables[index[0]].findall("w:tr", NS)[index[1]].findall("w:tc", NS)[index[2]]
                except IndexError as exc:
                    raise ReportError(f"Word 单元格不存在：{index}") from exc
                ps = cell.findall("w:p", NS)
                if len(ps) != 1 or cell.find("w:tbl", NS) is not None:
                    raise ReportError(f"单元格 {index} 须为一个段落且不能嵌套表格；复杂单元格请用占位符")
                nodes = texts(ps[0])
                actual = "".join(node_text(n) for n in nodes)
                if actual != t["expected_text"]:
                    raise ReportError(f"单元格 {index} 原文不符，期望 {t['expected_text']!r}，实际 {actual!r}")
                if actual:
                    replace_span(nodes, 0, len(actual), values[f["id"]])
                else:
                    node = next((n for n in nodes if n.tag == W + "t"), None)
                    if node is None:
                        run = ps[0].find("w:r", NS)
                        if run is None:
                            run = etree.SubElement(ps[0], W + "r")
                        node = etree.SubElement(run, W + "t")
                    node.text = values[f["id"]]
                self.modified.add("word/document.xml")
                logs.append({"field": f["id"], "target": list(index), "before": actual, "after": values[f["id"]]})
        counts = Counter(x["placeholder"] for x in self.inventory()["placeholders"])
        if set(counts) - set(replacements):
            raise ReportError(f"Word 含未映射占位符：{sorted(set(counts) - set(replacements))}")
        for token, count in expected.items():
            if counts[token] != count:
                raise ReportError(f"占位符 {token} 出现 {counts[token]} 次，配置要求 {count} 次")
        for name, p in self.paragraphs():
            nodes = texts(p)
            text = "".join(node_text(n) for n in nodes)
            for m in reversed(list(TOKEN.finditer(text))):
                replace_span(nodes, m.start(), m.end(), replacements[m.group(1)])
                self.modified.add(name)
                logs.append({"target": m.group(1), "part": name, "before": m.group(0), "after": replacements[m.group(1)]})
        for name, p in self.paragraphs():
            text = "".join(node_text(n) for n in texts(p))
            if "{{" in text or "}}" in text:
                raise ReportError(f"{name} 仍有未填或格式错误的占位符")
        return logs

    def save(self, path):
        with ZipFile(path, "w") as archive:
            for info, original in self.files:
                data = etree.tostring(self.roots[info.filename], encoding="UTF-8", xml_declaration=True, standalone=True) if info.filename in self.modified else original
                archive.writestr(info, data)
