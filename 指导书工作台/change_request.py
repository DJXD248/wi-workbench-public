# -*- coding: utf-8 -*-
"""一句话需求：用大白话在已有指导书 xlsx 上批量修改，并按规矩就地登记变更。

规矩（用户 2026-10-20 定）：
  1. 凡是被改过的页面，要在该页「内容变更 Change the content」栏写明变更内容＋日期；
  2. 清单页也要登记"哪一页改了什么"。
栏位坐标随文件而变（生成版在 I32、规范化版可能在 E32，甚至清单页没有这栏），
所以一律**按标签文字找栏**，不写死坐标；找不到就走兜底登记（见 _fallback_register）。

输入文件永远只读，结果另存到 成品\，文件名带日期。
"""
from __future__ import annotations

import os
import re
import shutil
from dataclasses import dataclass, field
from datetime import date

import openpyxl

SEP = "；;，,、\n"
_SPLIT = re.compile(r"([；;，,、\n])")
_LABEL_CHANGE = re.compile(r"内容变更|Change the content", re.I)
_LABEL_CONTENT = re.compile(r"^(内容|Content)", re.I)
_LABEL_DATE = re.compile(r"^(日期|Date)", re.I)
_STEP_MARK = re.compile(r"^(第([一二三四五六七八九十百\d]+)步|Step\s*(\d+))", re.I)
_TOKEN = re.compile(r"[A-Za-z0-9][A-Za-z0-9\-\./]*|[一-鿿]+|[^\sA-Za-z0-9一-鿿]+")

_CN = {"零": 0, "一": 1, "二": 2, "三": 3, "四": 4, "五": 5,
       "六": 6, "七": 7, "八": 8, "九": 9}


def _to_int(s: str) -> int:
    s = s.strip()
    if s.isdigit():
        return int(s)
    if "十" in s:
        left, right = s.split("十", 1)
        tens = _CN.get(left, 1) if left else 1
        ones = _CN.get(right, 0) if right else 0
        return tens * 10 + ones
    return _CN.get(s, 0)


# ============================================================
# 1. 句式解析
# ============================================================

@dataclass
class TextOp:
    kind: str                 # "replace" | "delete"
    source: str               # 要替换掉的原文 / 要删除的关键词
    target: str = ""          # 替换成的新文字（delete 时为空）
    raw: str = ""             # 原句片段，回执里给用户对着看

    def describe(self) -> str:
        if self.kind == "replace":
            return f"「{self.source}」改为「{self.target}」"
        return f"删除「{self.source}」"


_REPLACE = re.compile(
    r"[把将]\s*((?:所有|全部|的|之)?[^，,。；;]*?)"
    r"\s*(?:改成|改为|换成|替换为|替换成|修改为|修改成|变更为|更新为)\s*"
    r"([^，,。；;]+)")
_DELETE = re.compile(r"(?:删除|删掉|去掉|去除|移除)\s*((?:所有|全部)?[^，,。；;]+)")
_FILLER = re.compile(r"^(所有的|所有|全部|把|将|里的|中的|的)+")


def parse_request(sentence: str) -> list[TextOp]:
    """把一句大白话拆成文字操作。认不出来就返回空表（由界面如实告诉用户）。"""
    ops: list[TextOp] = []
    taken: list[tuple[int, int]] = []

    for m in _REPLACE.finditer(sentence):
        ops.append(TextOp("replace", _clean(m.group(1)), _clean(m.group(2), head_only=False), m.group(0)))
        taken.append(m.span())

    for m in _DELETE.finditer(sentence):
        if any(a < m.end() and m.start() < b for a, b in taken):
            continue
        kw = _clean(m.group(1))
        if kw:
            ops.append(TextOp("delete", kw, "", m.group(0)))
    return [op for op in ops if op.source]


def _clean(s: str, head_only: bool = True) -> str:
    s = (s or "").strip().strip("「」\"'“”")
    if head_only:
        s = _FILLER.sub("", s)
    else:
        s = re.sub(r"^(所有的|全部)", "", s)
    return s.strip("。，,；;、 ")


# ============================================================
# 2. 逐格改写
# ============================================================

def _segments(text: str, keyword: str) -> tuple[str, int]:
    """按分隔符切条，摘掉含关键词的那一条；返回新文字与删除条数"""
    parts = _SPLIT.split(text)
    kept, removed = [], 0
    for i in range(0, len(parts), 2):
        seg = parts[i]
        sepa = parts[i + 1] if i + 1 < len(parts) else ""
        if keyword in seg:
            removed += 1
            if not kept and sepa:            # 删的是开头一条：别留个孤立分隔符
                sepa = ""
            continue
        kept.append((seg, sepa))
    if not removed:
        return text, 0
    out = "".join(s + p for s, p in kept)
    if out.endswith(("；", ";", "，", ",", "、")):
        out = out[:-1]
    return out, removed


def _token_pairs(source: str, target: str) -> list[tuple[str, str]]:
    """整串找不到时的兜底：Excel 把"型号"和"名称"拆在相邻两格，
    所以把短语按"字母数字段 / 中文段"切开后一一对应地换。切不齐就不换（宁可不动）。"""
    a, b = _TOKEN.findall(source), _TOKEN.findall(target)
    if len(a) != len(b) or len(a) < 2:
        return []
    pairs = []
    for x, y in zip(a, b):
        if x == y:
            continue
        x_code = bool(re.fullmatch(r"[A-Za-z0-9][A-Za-z0-9\-\./]*", x))
        y_code = bool(re.fullmatch(r"[A-Za-z0-9][A-Za-z0-9\-\./]*", y))
        if x_code != y_code:
            return []                      # 一段是型号一段是中文，说明切歪了，放弃
        pairs.append((x, y))
    return pairs


@dataclass
class SheetEdit:
    sheet: str
    page_no: str = ""
    hits: dict = field(default_factory=dict)      # 操作描述 -> 处数
    split: int = 0                                 # 有几格是"型号/名称分在两格"按词换的
    cell_notes: list = field(default_factory=list)

    def summary(self) -> str:
        return "；".join(f"{k}（{v}处）" for k, v in self.hits.items())


def apply_text_ops(wb, ops: list[TextOp]) -> list[SheetEdit]:
    """在所有 sheet 上做文字替换/删除（清单页与工序页一起，"所有的"就是这个意思）。

    一格一格的判据：这格里能找到整句原文就整句换；找不到才允许"拆开换"
    （Excel 常把型号放一格、名称放隔壁格）。这样"所有"真的覆盖到每一格，
    又不会在同一个格里换两遍。"""
    edits: list[SheetEdit] = []
    for ws in wb.worksheets:
        ed = SheetEdit(ws.title, page_no=_page_no(ws))
        for row in ws.iter_rows():
            for cell in row:
                v = cell.value
                if not isinstance(v, str) or not v.strip():
                    continue
                if _LABEL_CHANGE.search(v):        # 别把自己写的登记又改一遍
                    continue
                new = v
                for op in ops:
                    if op.kind == "delete":
                        cand, n = _segments(new, op.source)
                        if n:
                            ed.hits[op.describe()] = ed.hits.get(op.describe(), 0) + n
                            new = cand
                    else:
                        n = new.count(op.source)
                        if n:
                            ed.hits[op.describe()] = ed.hits.get(op.describe(), 0) + n
                            new = new.replace(op.source, op.target)
                        else:
                            done = False
                            for x, y in _token_pairs(op.source, op.target):
                                if x in new:
                                    new = new.replace(x, y)
                                    done = True
                            if done:
                                ed.hits[op.describe()] = ed.hits.get(op.describe(), 0) + 1
                                ed.split += 1
                if new != v:
                    cell.value = new
                    ed.cell_notes.append(cell.coordinate)
        if ed.hits:
            edits.append(ed)
    return edits


_PROC_LABEL = re.compile(r"^(工序号|Working procedure)", re.I)
_ANY_LABEL = re.compile(r"工序号|Working procedure|工位名称|Testing part|页码|Page number", re.I)


def _page_no(ws) -> str:
    """给这一页一个清单认得出来的名字：优先工序号（ZT-1 / DEMO-1），其次页码。
    清单页自己也有"工序号"表头，那里读到的是表头文字，不能当页名用。"""
    for row in ws.iter_rows(min_row=1, max_row=8):
        for c in row:
            if isinstance(c.value, str) and _PROC_LABEL.match(c.value.strip()):
                v = _get(ws, c.row, c.column + 2)
                if v and not _ANY_LABEL.search(v):
                    return v
    for coord in ((4, 11), (2, 12)):
        v = _get(ws, coord[0], coord[1])
        if v and not _ANY_LABEL.search(v):
            return v
    return ""


# ============================================================
# 3. 变更登记：工序页变更栏 ＋ 清单页
# ============================================================

def find_change_block(ws):
    """按标签找到该页的变更登记栏。返回 (内容列, 日期列, [槽行号]) 或 None。

    两种形状都要认：
      A 工序页——标签独占一行（I32:E 皆然），下一两行才是"内容Content｜日期Date"表头；
      B 清单页——标签行自己就是表头，且"日期Date"可能在内容列的**左边**（A8 日期 / B8 内容），
        还可能同一行并排两组（规范化出来的老文件就是这样），取最左边那组。
    """
    label = None
    for row in ws.iter_rows():
        for c in row:
            if isinstance(c.value, str) and _LABEL_CHANGE.search(c.value):
                label = c
                break
        if label:
            break
    if not label:
        return None

    c1, c2 = _merge_span(ws, label)
    date_col = _date_col_in_row(ws, label.row, c1, c2)
    if date_col is not None:                       # 形状 B：标签行即表头
        head_row, col = label.row, c1
    else:                                          # 形状 A：另找内容表头行
        head_row = col = None
        for r in range(label.row, label.row + 4):
            for c in range(1, ws.max_column + 1):
                v = ws.cell(r, c).value
                if isinstance(v, str) and _LABEL_CONTENT.match(v.strip()) \
                        and not _LABEL_CHANGE.search(v):
                    head_row, col = r, _merge_span(ws, ws.cell(r, c))[0]
                    break
            if head_row:
                break
        if not head_row:
            return None
        _, e = _merge_span(ws, ws.cell(head_row, col))
        date_col = _date_col_in_row(ws, head_row, col, e)
    if date_col is None:
        return None

    want = (col, span_width(ws, head_row, col))
    slots, r = [], head_row + 1
    while r <= ws.max_row and len(slots) < 6:
        if _merge_span(ws, ws.cell(r, col)) != want:
            break
        if _next_header_in_row(ws, r):       # 再往下就是工序表了，别把登记写进去
            break
        slots.append(r)
        r += 1
    return (col, date_col, slots) if slots else None


_NEXT_HEADER = re.compile(r"工序号|Working procedure|工位名称|Testing part", re.I)


def _next_header_in_row(ws, row) -> bool:
    """表头字样只在最左两列、且是短标签时才算。
    我们自己写的登记文字里也带"工序号"三个字，长段落不能当表头（否则下一轮跑
    这文件时，槽位会被自己写的内容截掉）。"""
    for c in range(1, 3):
        v = ws.cell(row, c).value
        if isinstance(v, str) and len(v) <= 24 and _NEXT_HEADER.search(v):
            return True
    return False


def span_width(ws, row, col) -> int:
    return _merge_span(ws, ws.cell(row, col))[1]


def _merge_span(ws, cell) -> tuple[int, int]:
    for m in ws.merged_cells.ranges:
        if m.min_row <= cell.row <= m.max_row and m.min_col <= cell.column <= m.max_col:
            return m.min_col, m.max_col
    return cell.column, cell.column


def _date_col_in_row(ws, row, content_first: int, content_last: int):
    """日期标签可能挨在内容栏的左边也可能在右边；先左后右找"""
    for c in range(content_first - 1, 0, -1):
        v = ws.cell(row, c).value
        if isinstance(v, str) and _LABEL_DATE.match(v.strip()):
            return c
    for c in range(content_last + 1, min(content_last + 4, ws.max_column + 1)):
        v = ws.cell(row, c).value
        if isinstance(v, str) and _LABEL_DATE.match(v.strip()):
            return c
    return None


def _put(ws, row: int, col: int, value) -> None:
    """合并区只有左上角可写，日期列落在合并区里时按左上角写"""
    for m in ws.merged_cells.ranges:
        if m.min_row <= row <= m.max_row and m.min_col <= col <= m.max_col:
            row, col = m.min_row, m.min_col
            break
    ws.cell(row, col).value = value


def _get(ws, row: int, col: int) -> str:
    for m in ws.merged_cells.ranges:
        if m.min_row <= row <= m.max_row and m.min_col <= col <= m.max_col:
            row, col = m.min_row, m.min_col
            break
    v = ws.cell(row, col).value
    return "" if v is None else str(v).strip()


def write_block_entry(ws, block, content: str, today: str) -> str:
    """往变更栏第一个空槽写一条；槽满了就并进最后一槽。返回落点描述。"""
    col, dcol, slots = block
    for r in slots:
        if not _get(ws, r, col) and not _get(ws, r, dcol):
            _put(ws, r, col, content)
            _put(ws, r, dcol, today)
            return f"第 {r} 行变更栏"
    last = slots[-1]
    old = _get(ws, last, col)
    _put(ws, last, col, f"{old}\n{content}" if old else content)
    _put(ws, last, dcol, today)
    return f"第 {last} 行变更栏（栏位已满，并写在一起）"


def _count_slots(ws, block) -> int:
    col, dcol, slots = block
    return sum(1 for r in slots if not _get(ws, r, col) and not _get(ws, r, dcol))


def _append_register(ws, entries: list[str], today: str, title: str) -> str:
    """没有变更栏、或栏位已经写满：在本表最下面追加登记，不动任何原有格子。
    同名追加栏已存在就接着往下写，第二回跑不会重复建一个表头。"""
    head = f"变更登记Change record（工具追加）｜{title}"
    r = None
    for row in ws.iter_rows(min_col=1, max_col=1):
        c = row[0]
        if isinstance(c.value, str) and c.value.startswith(head):
            r = c.row + 2
            break
    if r is None:
        r = ws.max_row + 2
        for m in ws.merged_cells.ranges:
            r = max(r, m.max_row + 2)
        ws.cell(r, 1).value = head
        ws.cell(r + 1, 1).value = "日期Date"
        ws.cell(r + 1, 2).value = "内容变更Change the content"
        r += 2
    while ws.cell(r, 1).value is not None or ws.cell(r, 2).value is not None:
        r += 1
    for content in entries:
        ws.cell(r, 1).value = today
        ws.cell(r, 2).value = content
        r += 1
    return "已在本表下方追加登记"


def register_changes(wb, edits: list[SheetEdit], today: str,
                     image_notes: dict[str, list[str]], ops_text: str = "") -> list[str]:
    """规矩落地：每页在自己的「内容变更」栏写明改了什么＋日期；
    清单页再登记"哪一页改了什么"。清单栏位放不下时改成一条汇总＋下方追加明细，
    绝不把十几条硬塞进同一个格子（塞进去会被行高截掉，等于没写）。"""
    notes = []
    cover = next((ws for ws in wb.worksheets if "清单" in ws.title), wb.worksheets[0])
    per_page: list[tuple[str, str]] = []        # (页名, 改了什么)

    for ed in edits:
        ws = wb[ed.sheet]
        lines = [f"{k}（{v}处）" for k, v in ed.hits.items()]
        lines += image_notes.get(ed.sheet, [])
        if not lines:
            continue
        content = "；".join(lines)
        label = _page_label(ed)
        if ws is cover:
            # 清单页自己的改动就记在下面那份"哪一页改了什么"里，
            # 不然它会先占掉变更栏的格子，让后面真正要登记的页数没地方写
            notes.append(f"{label}：本表有改动，已并入清单登记")
        else:
            block = find_change_block(ws)
            if block and _count_slots(ws, block):
                notes.append(f"{label}：已写入{write_block_entry(ws, block, content, today)}")
            else:
                notes.append(f"{label}：{_append_register(ws, [content], today, label)}")
        per_page.append((label, content))

    if per_page:
        block = find_change_block(cover)
        space = _count_slots(cover, block) if block else 0
        detail = [f"{p}：{c}" for p, c in per_page]
        summary = f"一句话需求批量修改：{ops_text or '见下方明细'}；共涉及 {len(detail)} 页"
        if space >= len(detail):
            for line in detail:
                write_block_entry(cover, block, line, today)
            notes.append(f"清单页：{len(detail)} 页逐条写入变更栏")
        elif space >= 1:
            for line in detail[:space]:
                write_block_entry(cover, block, line, today)
            _append_register(cover, [summary] + detail[space:], today, "清单")
            notes.append(f"清单页：变更栏还有 {space} 格，逐条写了前 {space} 页；"
                         f"其余 {len(detail) - space} 页连同汇总追加在本表下方")
        else:
            _append_register(cover, [summary] + detail, today, "清单")
            notes.append(f"清单页：变更栏已无空格，{len(detail)} 页明细＋汇总追加在本表下方")
    return notes


def _page_label(ed: SheetEdit) -> str:
    """清单页要说清楚是哪一页：用表名，附带工序号（清单上的工序表就是按它排的）"""
    if ed.page_no and ed.page_no != ed.sheet:
        return f"「{ed.sheet}」页（工序号 {ed.page_no}）"
    return f"「{ed.sheet}」页"


# ============================================================
# 4. 配图更换
# ============================================================

def list_steps(ws) -> list[dict]:
    """列出这一页出现的步骤标记（给界面当参照，不参与归属判断）"""
    out = []
    for row in ws.iter_rows():
        for c in row:
            if isinstance(c.value, str):
                m = _STEP_MARK.match(c.value.strip())
                if m:
                    out.append({"step": _to_int(m.group(2) or m.group(3)),
                                "row": c.row, "col": c.column,
                                "where": c.coordinate})
    return sorted(out, key=lambda s: (s["row"], s["col"]))


def list_images(ws) -> list[dict]:
    """列出这一页每张配图的序号、锚点位置与显示尺寸。
    外来版式的步骤标记列位不统一（实测见过 A5/C5 各一步），所以**按序号点选**最稳，
    另附"上方最近的步骤标记"作提示，不拿它当判据。"""
    marks = list_steps(ws)
    out = []
    for i, im in enumerate(getattr(ws, "_images", [])):
        frm = getattr(im.anchor, "_from", None)
        if frm is None:
            continue
        row1, col1 = frm.row + 1, frm.col + 1
        above = [m for m in marks if m["row"] <= row1]
        hint = ""
        if above:
            near = min(above, key=lambda m: (row1 - m["row"], abs(m["col"] - col1)))
            hint = f"疑似第{near['step']}步（{near['where']}）"
        out.append({"index": i, "row": row1, "col": col1,
                    "pos": f"第{row1}行第{col1}列", "w": int(im.width or 0),
                    "h": int(im.height or 0), "hint": hint})
    return sorted(out, key=lambda x: (x["row"], x["col"]))


def replace_image(ws, image_index: int, new_path: str) -> str:
    """按序号换掉一张配图：锚点与显示尺寸原样保留，所以版面不动"""
    imgs = getattr(ws, "_images", [])
    if not (0 <= image_index < len(imgs)):
        raise RuntimeError(f"「{ws.title}」没有编号 {image_index} 的配图"
                           f"（共 {len(imgs)} 张）")
    old = imgs[image_index]
    frm = getattr(old.anchor, "_from", None)
    from openpyxl.drawing.image import Image as XLImage
    tmp = os.path.join(os.environ.get("TEMP", "."),
                       f"_wi_newimg_{os.getpid()}_{id(old)}.png")
    if new_path.lower().endswith(".png"):
        shutil.copyfile(new_path, tmp)
    else:
        from PIL import Image as PILImage
        PILImage.open(new_path).convert("RGB").save(tmp, "PNG")
    new = XLImage(tmp)
    new.width, new.height = old.width, old.height
    new.anchor = old.anchor
    imgs[image_index] = new
    at = f"第{frm.row + 1}行第{frm.col + 1}列" if frm else "原位"
    return f"配图 #{image_index + 1}（{at}）已换成 {os.path.basename(new_path)}"


# ============================================================
# 5. 对外入口
# ============================================================

@dataclass
class Report:
    ops: list[TextOp] = field(default_factory=list)
    edits: list[SheetEdit] = field(default_factory=list)
    notes: list[str] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)
    output: str = ""

    def preview_text(self) -> str:
        if not self.ops:
            return "这句话我没能拆出可执行的修改。请用这两种句式之一：\n" \
                   "    把所有的A改成B\n    删除所有A\n"
        lines = [f"将要执行：{i + 1}. {op.describe()}" for i, op in enumerate(self.ops)]
        lines.append("")
        if self.edits:
            for ed in self.edits:
                tail = f"（其中 {ed.split} 格是型号与名称分在两格，按词对应换）" if ed.split else ""
                lines.append(f"{_page_label(ed)}： {ed.summary()}{tail}")
            n = sum(sum(ed.hits.values()) for ed in self.edits)
            lines.append("")
            lines.append(f"合计 {len(self.edits)} 页、{n} 处。"
                         "改完会在每页「内容变更」栏写明改了什么＋日期，"
                         "清单页同步登记哪一页改了什么。")
        else:
            lines.append("⚠ 现有一页都没命中，执行不会改动任何文字。"
                         "请把要改的原话贴得更准（可只取其中一段）。")
        return "\n".join(lines)


def plan(src: str, sentence: str, today: str | None = None) -> Report:
    """只读试跑：算出会命中哪些页、各几处，不写任何文件"""
    today = today or date.today().isoformat()
    wb = openpyxl.load_workbook(src)
    ops = parse_request(sentence)
    edits = apply_text_ops(wb, ops) if ops else []
    rep = Report(ops=ops, edits=edits)
    if not ops:
        rep.warnings.append("句式没认出来")
    return rep


def run(src: str, sentence: str, out_dir: str | None = None,
        image_ops: list[tuple] | None = None, today: str | None = None) -> Report:
    """执行：另存副本到 out_dir（默认与输入同目录），原文件不动。
    image_ops = [(工作表名, 配图序号0起, 新图片路径), ...]"""
    today = today or date.today().isoformat()
    wb = openpyxl.load_workbook(src)
    rep = Report()
    rep.ops = parse_request(sentence)
    if rep.ops:
        rep.edits = apply_text_ops(wb, rep.ops)

    image_notes: dict[str, list[str]] = {}
    for sheet, image_index, path in (image_ops or []):
        image_notes.setdefault(sheet, []).append(
            replace_image(wb[sheet], image_index, path))
    for sheet, lines in image_notes.items():
        if sheet not in {ed.sheet for ed in rep.edits}:
            rep.edits.append(SheetEdit(sheet, page_no=_page_no(wb[sheet])))

    if not rep.edits:
        rep.warnings.append("没有任何命中，未写出文件")
        return rep
    rep.notes = register_changes(
        wb, rep.edits, today, image_notes,
        ops_text="；".join(op.describe() for op in rep.ops))

    base = os.path.splitext(os.path.basename(src))[0]
    out_dir = out_dir or os.path.dirname(os.path.abspath(src))
    os.makedirs(out_dir, exist_ok=True)
    out = os.path.join(out_dir, f"{base}_改_{today}.xlsx")
    n = 1
    while os.path.exists(out):
        out = os.path.join(out_dir, f"{base}_改_{today}_{n}.xlsx")
        n += 1
    wb.save(out)
    rep.output = out
    return rep
