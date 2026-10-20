# -*- coding: utf-8 -*-
"""无界面回归检查（云端任务也能跑：不需要 Excel/WPS/GUI，只装 requirements.txt）。

它把四条规矩变成机器判据：
  1) 正文不许动 —— 规范化只能换版头，第 5 行往下的文字必须逐字不变、图片一张不丢；
  2) 步骤不限 4 步 —— 6 步/8 步的工序要在一页里排下，标题行落在预期坐标，页脚跟着下移；
  3) 一道工序 = 一个工作表 —— 不许拆成多个 sheet；
  4) 一句话需求 —— 只改该改的格，未命中的格逐字不变；改过的页必须自己写上变更栏＋日期，
     清单页必须登记"哪一页改了什么"；换图不许增减张数、不许挪位置。

用法：python tests/verify.py        （从任意目录跑都行）
退出码：0 = 全部通过；非 0 = 有判据失败（打印 FAIL 行）。
"""
from __future__ import annotations

import os
import re
import shutil
import sys
import tempfile

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))   # 指导书工作台\
sys.path.insert(0, BASE)

import openpyxl
from PIL import Image as PILImage

from excel_generator import ExcelGenerator, Normalizer, generate_from_yaml, normalize_file
from wi_models import Process, WorkInstruction

_NORMALIZER = Normalizer(os.path.join(BASE, "作业指导书版头.xlsx"))

TEMPLATE = os.path.join(BASE, "作业指导书版头.xlsx")
# 规范化的回归输入。用假样本（结构与真实旧文件一致、内容全占位），
# 需要时由 tests/make_synthetic_fixture.py 在本机重新生成，公司内容不进仓库。
FIXTURE = os.path.join(BASE, "tests", "fixtures", "合成老格式样本.xlsx")

BAND = 17                    # 一段（左右两块）占 17 行
FIRST_LABEL_ROW = 5          # 第 0 段标题行
_CN = {"一": 1, "二": 2, "三": 3, "四": 4, "五": 5, "六": 6, "七": 7, "八": 8, "九": 9, "十": 10}
STEP_RE = re.compile(r"第([一二三四五六七八九十\d]+)步")

FAILS: list[str] = []
INFOS: list[str] = []


def _step_no(text_first_line: str):
    """把「第三步：」「Step 2第二步：」这类标题解析成步号；不是步骤标题则返回 None。"""
    if not isinstance(text_first_line, str):
        return None
    head = text_first_line.strip().split("\n")[0]
    if not (STEP_RE.search(head) or re.search(r"Step\s*\d", head, re.I)):
        return None
    m = STEP_RE.search(head)
    if not m:
        return None
    token = m.group(1)
    return int(token) if token.isdigit() else _CN.get(token)


def check(name: str, ok: bool, detail: str = ""):
    print(("PASS  " if ok else "FAIL  ") + name + (f"   [{detail}]" if detail else ""))
    if not ok:
        FAILS.append(name)


def info(text: str):
    INFOS.append(text)
    print("      · " + text)


def step_labels(ws) -> dict[int, str]:
    """工序页里每个步骤标题出现在哪个单元格。"""
    hits: dict[int, str] = {}
    for row in ws.iter_rows(min_row=1, max_row=ws.max_row, max_col=8):
        for c in row:
            n = _step_no(c.value)
            if n is not None:
                hits.setdefault(n, c.coordinate)
    return hits


def find_process_sheet(wb, n_steps: int):
    """返回包含 n_steps 个步骤标题的那个 sheet。"""
    for ws in wb.worksheets:
        if len(step_labels(ws)) == n_steps:
            return ws
    return None


def body_grid(ws, start_row: int) -> dict:
    """正文快照：从 start_row 往下所有非空单元格 (行,列) -> 文本。"""
    g = {}
    for row in ws.iter_rows(min_row=start_row, max_row=ws.max_row, max_col=ws.max_column):
        for c in row:
            if c.value not in (None, ""):
                g[(c.row, c.column)] = str(c.value)
    return g


def first_step_row(ws) -> int | None:
    """正文起点行 = 代码自己用的判据（Normalizer._first_step_row），不另起炉灶。"""
    return _NORMALIZER._first_step_row(ws)


def make_pngs(tmp: str, n: int) -> list[str]:
    paths = []
    for i in range(n):
        p = os.path.join(tmp, f"t{i}.png")
        PILImage.new("RGB", (40, 30), (i * 20 % 256, 60, 120)).save(p)
        paths.append(p)
    return paths


def build_wi(n_steps: int, imgs_per_step: int, tmp: str, tag: str) -> tuple[WorkInstruction, int]:
    """造一道 n_steps 步的工序，每步挂 imgs_per_step 张临时图；返回模型和声明的图片总数。"""
    pngs = make_pngs(tmp, n_steps * imgs_per_step)
    p = Process(process_no=f"{tag}-1", process_name=f"Test {tag}",
                process_name_cn=f"测试工序{tag}", work_hours=30, page_no=1)
    k = 0
    total = 0
    for i in range(1, n_steps + 1):
        p.add_step(i, f"第{i}步动作描述", f"Step {i} description")
        for _ in range(imgs_per_step):
            p.steps[-1].images.append(pngs[k])
            k += 1
            total += 1
    wi = WorkInstruction()
    wi.header.file_no = f"VERIFY-{tag}"
    wi.header.edition = "A"
    wi.header.product_spec_cn = f"回归检查{tag}"
    wi.header.made_by = "验证"
    wi.header.checked_by = "验证"
    wi.header.approved_by = "验证"
    wi.add_process(p)
    return wi, total


# ---------------------------------------------------------------- 生成路径

def verify_generated_layout(tmp: str, n_steps: int, imgs_per_step: int):
    tag = f"{n_steps}步"
    wi, declared = build_wi(n_steps, imgs_per_step, tmp, tag)
    out = os.path.join(tmp, f"verify_{n_steps}.xlsx")
    generate_path = ExcelGenerator(TEMPLATE).generate(wi, out)
    wb = openpyxl.load_workbook(generate_path)
    try:
        proc_sheets = [ws for ws in wb.worksheets if step_labels(ws)]
        check(f"[生成] {n_steps}步：一道工序一张表，没有拆成多页",
              len(proc_sheets) == len(wi.processes),
              f"工序页 {len(proc_sheets)} 张 / 总表 {len(wb.worksheets)} 张")

        ws = proc_sheets[0]
        labels = step_labels(ws)
        check(f"[生成] {n_steps}：每一步都有自己的标题", len(labels) == n_steps,
              "缺步 " + str([i for i in range(1, n_steps + 1) if i not in labels]) if len(labels) != n_steps else "")

        # 逐块核对坐标：块 i → 段 band=i//2、左列 A(第0块)或 E(第1块)、标题行 5+17*band
        wrong = []
        for i in range(n_steps):
            band, side = divmod(i, 2)
            want = f"{'A' if side == 0 else 'E'}{FIRST_LABEL_ROW + band * BAND}"
            if labels.get(i + 1) != want:
                wrong.append(f"第{i+1}步应在{want}实为{labels.get(i+1)}")
        check(f"[生成] {n_steps}：步骤块落在正确的版面位置（自动向下克隆）", not wrong, "; ".join(wrong[:4]))

        n_bands = (n_steps + 1) // 2
        footer_row = FIRST_LABEL_ROW + n_bands * BAND
        found = [ws.cell(row=footer_row, column=c).value for c in range(1, 13)]
        check(f"[生成] {n_steps}：页脚签字栏跟着段数下移到第 {footer_row} 行",
              any(isinstance(v, str) and "Made by" in v for v in found),
              "该行内容 " + str([v for v in found if v])[:80])

        got = len(ws._images)
        check(f"[生成] {n_steps}：配图零丢失（声明 {declared} 张 → 嵌入 {got} 张）", got == declared)

        ps = ws.sheet_properties.pageSetUpPr
        check(f"[生成] {n_steps}：整页缩放到一张打印纸（fitToPage/1×1）",
              bool(ps and ps.fitToPage) and ws.page_setup.fitToWidth == 1 and ws.page_setup.fitToHeight == 1,
              f"fitToWidth={ws.page_setup.fitToWidth} fitToHeight={ws.page_setup.fitToHeight}")
    finally:
        wb.close()


# ---------------------------------------------------------------- 规范化路径

def verify_normalize(tmp: str):
    if not os.path.exists(FIXTURE):
        check("[规范化] 回归样本存在", False, FIXTURE)
        return
    src_wb = openpyxl.load_workbook(FIXTURE)
    try:
        src = {}
        for ws in src_wb.worksheets:
            fsr = first_step_row(ws)
            if fsr:
                src[ws.title] = (fsr, body_grid(ws, fsr), len(ws._images),
                                 sorted(step_labels(ws).keys()))
        n_sheets = len(src_wb.worksheets)
    finally:
        src_wb.close()

    check("[规范化] 样本里认出了工序页", len(src) >= 20, f"工序页 {len(src)} 张 / 总表 {n_sheets} 张")
    if not src:
        check("[规范化] 后续判据没跑（工序页为 0，禁止空跑蒙绿）", False)
        return

    out = os.path.join(tmp, "verify_规范化.xlsx")
    normalize_file(FIXTURE, TEMPLATE, out)
    n_wb = openpyxl.load_workbook(out)
    try:
        check("[规范化] 工作表数量不变", len(n_wb.worksheets) == n_sheets,
              f"输入 {n_sheets} 张 → 输出 {len(n_wb.worksheets)} 张")
        same_body, same_img, same_steps, lost = True, True, True, []
        for title, (fsr, grid, nimg, steps) in src.items():
            if title not in n_wb.sheetnames:
                lost.append(title)
                same_body = same_img = same_steps = False
                continue
            ws = n_wb[title]
            if body_grid(ws, fsr) != grid:
                same_body = False
                lost.append(f"{title} 正文有改动")
            if len(ws._images) != nimg:
                same_img = False
                lost.append(f"{title} 图片 {nimg}→{len(ws._images)}")
            if sorted(step_labels(ws).keys()) != steps:
                same_steps = False
                lost.append(f"{title} 步骤序号集合变了")
        total_img = sum(v[2] for v in src.values())
        check(f"[规范化] 正文（步骤起始行往下）逐字未改（{len(src)} 张工序页）", same_body, "; ".join(lost[:3]))
        check(f"[规范化] 图片零丢失（样本合计 {total_img} 张）", same_img, "; ".join(lost[:3]))
        check("[规范化] 步骤没有被截断（几步还是几步）", same_steps, "; ".join(lost[:3]))
        check(f"[规范化] 版头确实被换上去了", header_replaced(out, FIXTURE, src), "各工序页第 5 行以上已按标准模板重写")
    finally:
        n_wb.close()


def header_replaced(out_path: str, src_path: str, src: dict) -> bool:
    """版头判据：正文起点以上的区域，输出必须与输入不同（说明真的换了版头），
    且输出版头区里必须出现标准模板的英文标签。"""
    a = openpyxl.load_workbook(src_path)
    b = openpyxl.load_workbook(out_path)
    changed = tagged = 0
    try:
        for title, (fsr, _g, _i, _s) in src.items():
            ws_a, ws_b = a[title], b[title]
            ga = {(c.row, c.column): str(c.value) for row in ws_a.iter_rows(min_row=1, max_row=fsr - 1, max_col=12)
                  for c in row if c.value not in (None, "")}
            gb = {(c.row, c.column): str(c.value) for row in ws_b.iter_rows(min_row=1, max_row=fsr - 1, max_col=12)
                  for c in row if c.value not in (None, "")}
            if ga != gb:
                changed += 1
            if any("Working Instruction" in v or "作业指导书" in v for v in gb.values()):
                tagged += 1
        return changed > 0 and tagged == changed
    finally:
        a.close()
        b.close()


# ---------------------------------------------------------------- 配置库

# ---------------------------------------------------------------- 一句话需求路径

def _block_cells(ws, block) -> set:
    """变更栏会动到的格子（内容列＋日期列，逐个槽位，合并区取左上角）"""
    col, dcol, slots = block
    out = set()
    for r in slots:
        for c in (col, dcol):
            for m in ws.merged_cells.ranges:
                if m.min_row <= r <= m.max_row and m.min_col <= c <= m.max_col:
                    r, c = m.min_row, m.min_col
                    break
            out.add((r, c))
    return out


def _grid(ws) -> dict:
    g = {}
    for row in ws.iter_rows():
        for c in row:
            if c.value not in (None, ""):
                g[(c.row, c.column)] = str(c.value)
    return g


def verify_change_request(tmp: str):
    import hashlib
    from datetime import date
    import change_request as cr

    ops = cr.parse_request("把所有的F435氩弧焊机改成F348TIG焊机，删除所有氩气充气时间")
    check("[一句话] 大白话能拆成一改一删两个动作",
          len(ops) == 2 and ops[0].kind == "replace" and ops[1].kind == "delete"
          and ops[0].source == "F435氩弧焊机" and ops[0].target == "F348TIG焊机"
          and ops[1].source == "氩气充气时间", str([(o.kind, o.source, o.target) for o in ops]))
    check("[一句话] 与修改无关的话不会硬编出动作",
          cr.parse_request("今天天气不错，我们再聊聊") == [])

    # 造一份带"型号与名称分在两格"的生成版：F435 在编号格、氩弧焊机在名称格
    wi, declared = build_wi(2, 1, tmp, "CR")
    p = wi.processes[0]
    p.add_material("F435", "氩弧焊机")
    p.steps[0].content = "调整气压，确认氩气充气时间 5 秒，开始点焊"
    src = os.path.join(tmp, "cr_src.xlsx")
    ExcelGenerator(TEMPLATE).generate(wi, src)
    before = hashlib.md5(open(src, "rb").read()).hexdigest()

    new_img = os.path.join(tmp, "cr_new.png")
    PILImage.new("RGB", (30, 20), (200, 30, 30)).save(new_img)
    src_wb = openpyxl.load_workbook(src)
    proc = next(ws for ws in src_wb.worksheets if step_labels(ws))
    proc_name, img_index = proc.title, 0
    src_anchor = getattr(proc._images[0].anchor, "_from", None)
    src_grid = {ws.title: _grid(ws) for ws in src_wb.worksheets}
    src_blocks = {ws.title: cr.find_change_block(ws) for ws in src_wb.worksheets}
    src_maxrow = {ws.title: ws.max_row for ws in src_wb.worksheets}
    src_wb.close()

    rep = cr.run(src, "把所有的F435氩弧焊机改成F348TIG焊机，删除所有氩气充气时间",
                 out_dir=os.path.join(tmp, "cr_out"),
                 image_ops=[(proc_name, img_index, new_img)])

    check("[一句话] 确实命中并写出了文件（0 命中必须判失败，不许空跑蒙绿）",
          bool(rep.edits) and bool(rep.output) and os.path.exists(rep.output),
          f"命中 {len(rep.edits)} 页；warnings={rep.warnings}")
    check("[一句话] 原文件一个字节都没动",
          hashlib.md5(open(src, "rb").read()).hexdigest() == before)
    if not rep.edits or not os.path.exists(rep.output):
        return

    out_wb = openpyxl.load_workbook(rep.output)
    try:
        # 1) 命中的格按预期改写
        g = {ws.title: _grid(ws) for ws in out_wb.worksheets}
        hit_cells = [k for k, v in g[proc_name].items() if v in ("F348TIG", "焊机")]
        check("[一句话] 型号与名称分在两格时也换到了（F435→F348TIG、氩弧焊机→焊机）",
              any(v == "F348TIG" for v in g[proc_name].values())
              and any(v == "焊机" for v in g[proc_name].values()) and len(hit_cells) == 2,
              str(hit_cells))
        step_txt = [v for v in g[proc_name].values() if "调整气压" in v]
        check("[一句话] 删除只摘掉含关键词的那一条，前后文留着",
              len(step_txt) == 1 and "氩气充气时间" not in step_txt[0]
              and "调整气压" in step_txt[0] and "开始点焊" in step_txt[0], repr(step_txt))

        # 2) 没命中的格子必须逐字不变（改动只允许是：命中格 / 变更栏 / 表下方追加）
        words = ("F435", "氩弧焊机", "氩气充气时间")
        bad = []
        for title, old in src_grid.items():
            new = g[title]
            allowed = _block_cells(out_wb[title], src_blocks[title]) if src_blocks[title] else set()
            for k, v in old.items():
                nv = new.get(k, "")
                if nv != v and not (k in allowed or any(w in v for w in words)):
                    bad.append(f"{title}!{k}")
            for k, v in new.items():
                if k not in old and v and k not in allowed and k[0] <= src_maxrow[title]:
                    bad.append(f"{title} 凭空多出{k}")
        check("[一句话] 未命中的格子逐字不变，也没多出别的格子", not bad, "; ".join(bad[:5]))

        # 3) 规矩①：改过的页面，自己那一页的变更栏写了内容＋日期
        today = date.today().isoformat()
        miss = []
        for ed in rep.edits:
            ws = out_wb[ed.sheet]
            block = cr.find_change_block(ws)
            if not block:
                miss.append(f"{ed.sheet} 找不到变更栏")
                continue
            col, dcol, slots = block
            ok = any(cr._get(ws, r, col).strip() and cr._get(ws, r, dcol).strip() == today
                     for r in slots)
            if not ok:
                miss.append(f"{ed.sheet} 变更栏没写今天({today})的记录")
        check(f"[一句话] {len(rep.edits)} 个改动页各自写了变更栏＋日期", not miss, "; ".join(miss[:4]))

        # 4) 规矩②：清单页登记"哪一页改了什么"
        cover = next((ws for ws in out_wb.worksheets if "清单" in ws.title), None)
        texts = [v for ws in ([cover] if cover else []) for row in ws.iter_rows() for v in
                 (str(c.value) for c in row if c.value is not None)]
        joined = "\n".join(texts)
        not_in_cover = [ed.sheet for ed in rep.edits if ed.sheet != cover.title
                        and cr._page_no(out_wb[ed.sheet]) not in joined]
        check(f"[一句话] 清单页列出了每个改动页（共 {len(rep.edits)} 页）",
              cover is not None and not not_in_cover,
              "清单里找不到工序号：" + ";".join(not_in_cover[:4]) if not_in_cover else "")

        # 5) 换图：张数一张不增减，位置原样
        ws = out_wb[proc_name]
        check("[一句话] 换图后配图张数不变（生成时嵌入 %d 张）" % declared,
              len(ws._images) == declared, f"输出 {len(ws._images)} 张")
        fa = getattr(ws._images[0].anchor, "_from", None)
        check("[一句话] 换上去的图还在原来的位置（锚点没挪）",
              fa is not None and src_anchor is not None
              and (fa.row, fa.col) == (src_anchor.row, src_anchor.col),
              f"{(src_anchor.row, src_anchor.col)} → {(fa.row, fa.col) if fa else None}")
    finally:
        out_wb.close()


def verify_configs(tmp: str):
    cfg_dir = os.path.join(BASE, "configs")
    yamls = [os.path.join(r, f) for r, _, fs in os.walk(cfg_dir)
             for f in fs if f.endswith(".yaml") and f != "blank_template.yaml"]
    ok = skipped = 0
    for y in yamls:
        name = os.path.basename(y)
        try:
            wb_test = ExcelGenerator(TEMPLATE)
            wi = WorkInstruction.from_yaml(y)
            if not wi.processes:
                skipped += 1
                continue
            out = os.path.join(tmp, "cfg_" + name.replace(".yaml", ".xlsx"))
            wb_test.generate(wi, out)
            g = openpyxl.load_workbook(out)
            n_pages = len([ws for ws in g.worksheets if step_labels(ws)])
            multi = [len(step_labels(ws)) for ws in g.worksheets if len(step_labels(ws)) > 4]
            g.close()
            if n_pages != len(wi.processes):
                check(f"[配置库] {name}：每工序一张表", False, f"{len(wi.processes)} 工序 → {n_pages} 表")
            else:
                ok += 1
                if multi:
                    info(f"{name}：超过 4 步的工序页 {multi}")
        except Exception as e:
            check(f"[配置库] {name}：能生成", False, f"{type(e).__name__}: {e}")
    check(f"[配置库] 全部配置可生成（{ok} 份通过，{skipped} 份空配置跳过）", True, "")


def main() -> int:
    tmp = tempfile.mkdtemp(prefix="wi_verify_")
    try:
        wb = openpyxl.load_workbook(TEMPLATE)
        names = wb.sheetnames
        wb.close()
        check("[模板] 版头模板可读且含「新格式」工序页", any("新格式" in n for n in names), str(names))

        verify_generated_layout(tmp, 4, 1)
        verify_generated_layout(tmp, 6, 2)
        verify_generated_layout(tmp, 8, 1)
        verify_normalize(tmp)
        verify_change_request(tmp)
        verify_configs(tmp)
    finally:
        shutil.rmtree(tmp, ignore_errors=True)

    print("\n" + ("全部通过 ✅" if not FAILS else f"有 {len(FAILS)} 项失败 ❌：" + "，".join(FAILS)))
    return 0 if not FAILS else 1


if __name__ == "__main__":
    sys.exit(main())
