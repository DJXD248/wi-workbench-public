# -*- coding: utf-8 -*-
"""无界面回归检查（云端任务也能跑：不需要 Excel/WPS/GUI，只装 requirements.txt）。

它把四条规矩变成机器判据：
  1) 正文不许动 —— 规范化只能换版头，第 5 行往下的文字必须逐字不变、图片一张不丢；
  2) 步骤不限 4 步 —— 6 步/8 步的工序要在一页里排下，标题行落在预期坐标，页脚跟着下移；
  3) 一道工序 = 一个工作表 —— 不许拆成多个 sheet；
  4) 一句话需求 —— 只改该改的格，未命中的格逐字不变；改过的页必须自己写上变更栏＋日期，
     清单页必须登记"哪一页改了什么"；换图不许增减张数、不许挪位置；
     没有变更栏或栏已写满时走表尾追加登记，重复追加不重建表头（2026-10-22 补）；
     生成侧另测 10 步仍不拆页（证明"步骤不限"没有隐性上限）。
  5) 无步号工序页（2026-10-22 用户口径"按图＋完整描述为一步"）：按配图锚点分步并换版头，
     正文逐字未改、每张图归描述在它上方的那一步、完整描述一句不丢；
     没有配图也没有步号的页维持老行为（告警、不换版头），不许瞎换。

用法：python tests/verify.py        （从任意目录跑都行）
退出码：0 = 全部通过；非 0 = 有判据失败（打印 FAIL 行）。
"""
from __future__ import annotations

import os
import re
import shutil
import sys
import tempfile
from pathlib import Path

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


# ---------------------------------------------------------------- 版面几何（图片位置/尺寸）

# 模板「新格式」真实列宽（XML 实测：G/H 由一条 min=7 max=8 的定义合并设置，都是 15）
COL_W = {"A": 15, "B": 12.875, "C": 17.5, "D": 14.25,
         "E": 16.5, "F": 13.875, "G": 15, "H": 15}
EMU_PER_PX = 9525
ROW_PX = 17.65 * 96 / 72          # 模板行高 17.65pt = 23.533px


def verify_image_geometry(tmp: str):
    """版面几何判据（2026-10-22 补，专防「吞字/位置不正确/大小不一」复发）：
      1. 任何图片底部不许越过文字区顶线——压住步骤描述就是吞字；
      2. 每块最顶上的图片必须贴着图区顶起排——不允许标签和图之间空一大截；
      3. 列宽不许有分组/重复定义——H 列曾因 15/13 两条打架而版面错乱；
      4. 同块同排图片必须等宽——方案甲；锚点 colOff 必须落在所在列内。
    横图/竖图/一步8图两排三种构型各跑一遍。"""
    def build(counts, size, tag: str) -> WorkInstruction:
        """counts: 每一步各挂几张图。"""
        pngs = []
        for i in range(sum(counts)):
            p = os.path.join(tmp, f"geo_{tag}_{i}.png")
            PILImage.new("RGB", size, (80, 110, 160)).save(p)
            pngs.append(p)
        p = Process(process_no=f"{tag}-1", process_name=f"Test {tag}",
                    process_name_cn=f"几何{tag}", work_hours=30, page_no=1)
        k = 0
        for i, cnt in enumerate(counts, 1):
            p.add_step(i, f"第{i}步动作描述", f"Step {i} description")
            for _ in range(cnt):
                p.steps[-1].images.append(pngs[k])
                k += 1
        wi = WorkInstruction()
        wi.header.file_no = f"GEO-{tag}"
        wi.header.edition = "A"
        wi.header.product_spec_cn = f"版面几何{tag}"
        wi.header.made_by = wi.header.checked_by = wi.header.approved_by = "几何"
        wi.add_process(p)
        return wi

    # 注意：工序页必须是 4 步——模板自带 4 个占位标题，2 步的页会被占位标题
    # 干扰识别（历史行为），所以一步 8 图的两排场景用 4 步 + 首步 8 图来构造
    for counts, size, tag in (([2, 2, 2, 2], (320, 240), "横图"),
                              ([2, 2, 2, 2], (240, 320), "竖图"),
                              ([8, 1, 1, 1], (320, 240), "两排")):
        n_steps = len(counts)
        total = sum(counts)
        wi = build(counts, size, tag)
        out = os.path.join(tmp, f"geo_{tag}.xlsx")
        ExcelGenerator(TEMPLATE).generate(wi, out)
        wb = openpyxl.load_workbook(out)
        try:
            ws = find_process_sheet(wb, n_steps)
            if ws is None:
                check(f"[版面] {tag}：找到工序页", False, "未找到")
                continue

            # 每张图：锚点行/列 + colOff/rowOff + 尺寸，全部换算成 px
            imgs_info = []
            for img in ws._images:
                a = img.anchor
                frm, ext = a._from, a.ext
                row0, col0 = frm.row, frm.col
                top = row0 * ROW_PX + (frm.rowOff or 0) / EMU_PER_PX
                left = (sum(COL_W[chr(65 + k)] * 7.5 for k in range(col0))
                        + (frm.colOff or 0) / EMU_PER_PX)
                wpx, hpx = ext.cx / EMU_PER_PX, ext.cy / EMU_PER_PX
                imgs_info.append((row0, col0, top, left, wpx, hpx))
            check(f"[版面] {tag}：图片全部嵌入（{len(imgs_info)} 张）",
                  len(imgs_info) == total,
                  f"期望 {total} 实得 {len(imgs_info)}")
            if len(imgs_info) != total:
                continue

            # 判据1：图片底部不许越过文字区顶线（吞字防线）
            press = []
            for (row0, col0, top, left, wpx, hpx) in imgs_info:
                band = (row0 - FIRST_LABEL_ROW) // BAND if row0 >= FIRST_LABEL_ROW else 0
                r_lab = FIRST_LABEL_ROW + band * BAND
                txt_top = (r_lab + 12) * ROW_PX        # 文字区首行（1-indexed r_lab+13）的行首
                if top + hpx > txt_top + 1:
                    press.append(f"图@行{row0 + 1}底{top + hpx:.0f}px 越过文字线 {txt_top:.0f}px")
            check(f"[版面] {tag}：没有图片压住步骤文字（吞字防线）", not press,
                  "; ".join(press[:4]))

            # 判据2：每块最顶上的图必须贴着图区顶（位置防线）
            gap_bad = []
            for band in range((max(r[0] for r in imgs_info) - FIRST_LABEL_ROW) // BAND + 1):
                r_lab = FIRST_LABEL_ROW + band * BAND
                img_top = r_lab * ROW_PX               # 图区首行（1-indexed r_lab+1）的行首
                rows_in = [t for (r0, c, t, l, w, h) in imgs_info
                           if r_lab <= r0 <= r_lab + 11]
                if rows_in and min(rows_in) > img_top + 6:
                    gap_bad.append(f"第{band + 1}段图顶空白 {min(rows_in) - img_top:.0f}px")
            check(f"[版面] {tag}：图片贴着标题下方起排（不留大白条）", not gap_bad,
                  "; ".join(gap_bad[:4]))

            # 判据3：列宽无分组/重复定义（版面错乱防线）
            cover = {}
            for key, cd in ws.column_dimensions.items():
                if cd.min is None or cd.max is None:
                    continue
                if cd.max < 1 or cd.min > 8:
                    continue
                for c in range(max(cd.min, 1), min(cd.max, 8) + 1):
                    cover.setdefault(c, []).append((key, cd.width))
            conflict = [f"第{c}列:{[w for _, w in lst]}"
                        for c, lst in sorted(cover.items())
                        if len({round(w, 3) for _, w in lst}) > 1]
            check(f"[版面] {tag}：列宽无冲突定义（H 列不再 15/13 打架）", not conflict,
                  "; ".join(conflict[:4]))

            # 判据4：同块同排等宽（方案甲）+ 锚点 colOff 落在所在列内
            groups = {}
            for (row0, col0, top, left, wpx, hpx) in imgs_info:
                groups.setdefault((row0, col0 // 4), []).append(wpx)
            uneven = [f"排{k}:{[round(w) for w in v]}"
                      for k, v in groups.items() if len(v) > 1 and max(v) - min(v) > 1]
            check(f"[版面] {tag}：同排图片等宽（方案甲）", not uneven,
                  "; ".join(uneven[:4]))
            off_bad = []
            for (row0, col0, top, left, wpx, hpx) in imgs_info:
                col_letter = chr(65 + col0)
                col_px = COL_W.get(col_letter, 8.43) * 7.5
                if left - sum(COL_W[chr(65 + k)] * 7.5 for k in range(col0)) > col_px + 0.5:
                    off_bad.append(f"图@{col_letter}{row0 + 1} colOff 越过列宽")
            check(f"[版面] {tag}：锚点偏移合法（colOff 不越列宽）", not off_bad,
                  "; ".join(off_bad[:4]))
        finally:
            wb.close()

    # 判据5（源码级防回归）：_fix_image_dimensions 里像素→EMU 必须乘 9525。
    # 这里误写过 914400（把像素当英寸），一旦 ext=0 的后处理路径被触发，
    # 图片会放大 96 倍——虽然正常流程不走该路径，也必须钉死。
    gen_src = Path(BASE, "excel_generator.py").read_text(encoding="utf-8")
    m = re.search(r"def _fix_image_dimensions\(.*?(?=\n    def |\nclass |\Z)",
                  gen_src, re.S)
    body = m.group(0) if m else ""
    code_only = "\n".join(l.split("#", 1)[0] for l in body.splitlines())  # 剥掉注释再查
    bad_mul = re.findall(r"\*\s*914400", code_only)
    detail = (f"函数体仍有 {len(bad_mul)} 处 * 914400 乘法" if bad_mul
              else ("函数体缺失" if not body else "像素→EMU = px×9525"))
    check("[版面] 后处理像素→EMU 换算正确（9525，不再放大 96 倍）",
          bool(body) and not bad_mul and "* 9525" in code_only, detail)


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


def verify_change_request_fallback(tmp: str):
    """变更登记的兜底路径：没有变更栏、或栏位已写满时，必须在本表下方追加登记，
    且对已存在的追加表接着往下写、不重复建表头。
    真实老文件大多没有变更栏（交接文档 §4：栏位坐标随文件而变），
    而这条路以前一条断言都没有——补上。"""
    from datetime import date
    import change_request as cr
    today = date.today().isoformat()
    HEAD = "变更登记Change record（工具追加）"

    def _heads(ws):
        return [c for row in ws.iter_rows(min_col=1, max_col=1) for c in row
                if isinstance(c.value, str) and c.value.startswith(HEAD)]

    def _has_today(ws):
        return any(isinstance(c.value, str) and c.value.strip() == today
                   for row in ws.iter_rows() for c in row)

    # --- 形状一：整份文件没有任何变更栏（老文件常态）---
    wb = openpyxl.Workbook()
    cover = wb.active
    cover.title = "清单"
    cover["A1"] = "工序号"
    cover["B1"] = "工位名称"
    cover["A2"] = "P-1"
    cover["B2"] = "测试工位 F435"
    proc = wb.create_sheet("工序P-1")
    proc["A1"] = "工序号"
    proc["B1"] = "P-1"
    proc["A3"] = "把 F435 装到台上，确认氩气充气时间 5 秒"
    src = os.path.join(tmp, "cr_fb_nobar.xlsx")
    wb.save(src)
    wb.close()

    rep = cr.run(src, "把所有的F435改成F348TIG，删除所有氩气充气时间",
                 out_dir=os.path.join(tmp, "cr_fb1_out"))
    check("[一句话] 兜底：没有变更栏的老文件也能改并写出文件",
          bool(rep.edits) and bool(rep.output) and os.path.exists(rep.output),
          f"命中 {len(rep.edits)} 页；warnings={rep.warnings}")
    if os.path.exists(rep.output):
        wb2 = openpyxl.load_workbook(rep.output)
        try:
            bad = []
            for title in ("清单", "工序P-1"):
                ws = wb2[title]
                if len(_heads(ws)) != 1 or not _has_today(ws):
                    bad.append(f"{title} 表头{len(_heads(ws))}个/"
                               f"日期{'有' if _has_today(ws) else '缺'}")
            check("[一句话] 兜底：无栏文件每张改过的表都追加了登记＋今天日期",
                  not bad, "; ".join(bad))
            g = {ws.title: _grid(ws) for ws in wb2.worksheets}
            # 只查正文格 A3：追加的登记文本本身就含"F435"字样（那是在记录这次改动），
            # 判"换没换干净"要看正文，不能把登记区搅进来
            a3 = g["工序P-1"].get((3, 1), "")
            check("[一句话] 兜底：追加登记不影响正文替换本身（该换的换到、该删的删掉）",
                  "F435" not in a3 and "F348TIG" in a3
                  and "氩气充气时间" not in a3, repr(a3))
        finally:
            wb2.close()

    # --- 形状二：有变更栏但槽位已写满 → 也走追加，不硬塞、不动旧记录 ---
    wb = openpyxl.Workbook()
    cover = wb.active
    cover.title = "清单"
    cover["A1"] = "工序号"
    cover["B1"] = "工位名称"
    cover["A2"] = "P-2"
    cover["B2"] = "测试工位 F435"
    proc = wb.create_sheet("工序P-2")
    proc["A1"] = "工序号"
    proc["B1"] = "P-2"
    proc["A3"] = "把 F435 装到台上"
    proc["A5"] = "内容变更"                                  # 标签独占一行（形状 A）
    proc["A6"] = "内容"
    proc["B6"] = "日期"
    proc["A7"] = "更早的记录"
    proc["B7"] = "2020-01-01"                                # 槽位预填＝写满
    src2 = os.path.join(tmp, "cr_fb_full.xlsx")
    wb.save(src2)
    wb.close()

    rep2 = cr.run(src2, "把所有的F435改成F348TIG",
                  out_dir=os.path.join(tmp, "cr_fb2_out"))
    if rep2.output and os.path.exists(rep2.output):
        wb3 = openpyxl.load_workbook(rep2.output)
        try:
            ws = wb3["工序P-2"]
            old_slot = (cr._get(ws, 7, 1), cr._get(ws, 7, 2))
            check("[一句话] 兜底：变更栏写满时改走表尾追加，且旧记录一个字节没动",
                  len(_heads(ws)) == 1 and _has_today(ws)
                  and old_slot == ("更早的记录", "2020-01-01"),
                  f"表头{len(_heads(ws))}个；第7行仍为{old_slot}")
        finally:
            wb3.close()
    else:
        check("[一句话] 兜底：变更栏写满时改走表尾追加", False,
              f"warnings={rep2.warnings}")

    # --- 回归：只读访问不许撑大表格边界（2026-10-22 栽过：_page_no 的只读访问
    # 会创建空单元格把 max_row 撑大，害得表尾空行被误判成变更栏空槽、登记写飞）---
    wb4 = openpyxl.load_workbook(src2)
    try:
        ws4 = wb4["工序P-2"]
        before = cr.find_change_block(ws4)
        cr._page_no(ws4)                     # 触发过 bug 的那个只读函数
        after = cr.find_change_block(ws4)
        check("[一句话] 回归：只读访问不撑大表格边界，变更栏槽位不漂移",
              before == after and before is not None,
              f"{before} → {after}")
    finally:
        wb4.close()

    # --- 单元级：对已存在的追加表接着写，不重复建表头 ---
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "工序P-3"
    cr._append_register(ws, ["第一次登记"], "2026-01-01", "「工序P-3」页")
    cr._append_register(ws, ["第二次登记"], "2026-01-02", "「工序P-3」页")
    hs = _heads(ws)
    all_text = " ".join(str(c.value) for row in ws.iter_rows() for c in row if c.value)
    check("[一句话] 兜底：同一张表第二遍追加不重建表头，两条登记都在",
          len(hs) == 1 and "第一次登记" in all_text and "第二次登记" in all_text,
          f"表头{len(hs)}个")


def verify_no_step_mark(tmp: str):
    """无步号工序页（用户 2026-10-22 口径：按"图＋完整描述"为一步）。
    以前这种页版头直接不换、提取步骤全空；现在按配图位置分步并换版头，
    但正文仍然一字不动（硬规则①）。没有配图的页维持老行为，不许瞎换。"""
    import io
    from openpyxl.drawing.image import Image as XLImage

    def _png(color):
        b = io.BytesIO()
        PILImage.new("RGB", (50, 40), color).save(b, "PNG")
        b.seek(0)
        return b

    # 版式：描述紧贴图上方（一步 = 图＋它的完整描述），中间隔空行
    wb = openpyxl.Workbook()
    cover = wb.active
    cover.title = "清单"
    cover["A1"] = "工序号"
    cover["A2"] = "ZT-91"
    ws = wb.create_sheet("ZT-91")
    ws["A1"] = "DEMO CONTROLS"                       # 版头特征行（往上吞的停止线）
    ws["A2"] = "作业指导书"
    ws["A3"] = "文件编号 WI-DEMO-91"
    ws["A6"] = "检查设备外观是否完好，确认无松动和异响"   # 步骤1描述，紧贴图1上方
    ws["A13"] = "调整压力表示数到规定值并记录"           # 步骤2描述，紧贴图2上方
    ws["I5"] = "设备/工夹具"                          # 右侧表格（挡内容区右界）
    ws["I6"] = "压力表"
    ws["A25"] = "编制"                               # 页脚起点
    ws.add_image(XLImage(_png((200, 60, 60))), "A7")
    ws.add_image(XLImage(_png((60, 60, 200))), "A14")
    src = os.path.join(tmp, "nostep_src.xlsx")
    wb.save(src)
    wb.close()

    out = os.path.join(tmp, "nostep_out.xlsx")
    wi = normalize_file(src, TEMPLATE, out)

    # 1) 按图分成 2 步
    proc = wi.processes[0] if wi.processes else None
    n_steps = len(proc.steps) if proc else 0
    check("[无步号] 按图＋完整描述认出步骤（2 张图 → 2 步）", n_steps == 2,
          f"实到 {n_steps} 步")

    src_wb = openpyxl.load_workbook(src)
    out_wb = openpyxl.load_workbook(out)
    try:
        ws_in, ws_out = src_wb["ZT-91"], out_wb["ZT-91"]
        fsr = _NORMALIZER._first_step_row(ws_in)
        check("[无步号] 找得到正文起点（从第一张图往上吞描述，空行/版头字样前停）",
              fsr == 6, f"fsr={fsr}")
        if fsr:
            hdr = " ".join(str(c.value) for row in ws_out.iter_rows(min_row=1, max_row=fsr - 1)
                           for c in row if c.value is not None)
            check("[无步号] 版头确实被换上去了（以前这种页不换）",
                  "Working Instruction" in hdr or "作业指导书" in hdr, hdr[:60])
            check("[无步号] 正文（起点行往下）逐字未改",
                  body_grid(ws_out, fsr) == body_grid(ws_in, fsr), "")
            check("[无步号] 图片零丢失",
                  len(ws_out._images) == len(ws_in._images) == 2,
                  f"输入 {len(ws_in._images)} → 输出 {len(ws_out._images)}")
            texts = " ".join(str(c.value) for row in ws_out.iter_rows(min_row=fsr)
                             for c in row if isinstance(c.value, str))
            check("[无步号] 两步的描述都没被误当版头删掉",
                  "检查设备外观是否完好" in texts and "调整压力表示数" in texts, "")
    finally:
        src_wb.close()
        out_wb.close()

    # 2) 步骤—配图配对 + 描述完整：走真实的带图归位链路
    #    （_parse_old_file_with_images：图片按锚点落进步骤块几何归位）。
    #    用户口径：一张图＋它的完整描述 = 一步；图归描述在它上方的那一步，
    #    两步文字各归各、一句不丢（13 字无标点的描述也曾被当成图片说明丢掉）。
    _NORMALIZER._img_output_dir = Path(os.path.join(tmp, "nostep_imgs"))
    wb5 = openpyxl.load_workbook(src)
    try:
        parsed = _NORMALIZER._parse_old_file_with_images(wb5, [], src)
        proc = parsed.processes[0] if parsed.processes else None
        per_step = [len(s.images) for s in proc.steps] if proc else []
        check("[无步号] 每张图归描述在它上方的那一步（两步各 1 张图）",
              per_step == [1, 1], f"各步配图数 {per_step}")
        if proc and len(proc.steps) >= 2:
            t1, t2 = proc.steps[0].content, proc.steps[1].content
            check("[无步号] 两步的完整描述都提取到，没被当成图片说明丢掉",
                  "检查设备外观" in t1 and "调整压力表示数" in t2,
                  f"步骤1[{t1[:18]}] 步骤2[{t2[:18]}]")
        else:
            check("[无步号] 两步的完整描述都提取到，没被当成图片说明丢掉",
                  False, f"实到 {len(proc.steps) if proc else 0} 步")
    finally:
        wb5.close()

    # 没有配图也没有步号的页：维持老行为（告警、不换版头），不许瞎换
    wb2 = openpyxl.Workbook()
    cov2 = wb2.active
    cov2.title = "清单"
    cov2["A1"] = "工序号"
    cov2["A2"] = "ZT-92"
    ws2 = wb2.create_sheet("ZT-92")
    ws2["A1"] = "DEMO CONTROLS"
    ws2["A2"] = "文件编号 WI-DEMO-92"
    ws2["A5"] = "把工件放到工作台上，用压板固定"
    src2 = os.path.join(tmp, "nostep_noimg.xlsx")
    wb2.save(src2)
    wb2.close()
    out2 = os.path.join(tmp, "nostep_noimg_out.xlsx")
    normalize_file(src2, TEMPLATE, out2)
    a = openpyxl.load_workbook(src2)
    b = openpyxl.load_workbook(out2)
    try:
        check("[无步号] 无图无步号的页维持老行为：版头不换、一字不动",
              _grid(a["ZT-92"]) == _grid(b["ZT-92"]), "")
    finally:
        a.close()
        b.close()


def verify_configs(tmp: str):
    cfg_dir = os.path.join(BASE, "configs")
    # 防复发：顶层系列目录名去掉空格后必须唯一。2026-10-22 删过一对
    # "AI 演示系列" / "AI演示系列" 重复目录——看板统计按系列名汇总，
    # 近似重复会把一个系列算成两个，GUI 里两个同名文件也容易点错。
    seen, dup = {}, []
    for d in sorted(os.listdir(cfg_dir)):
        if not os.path.isdir(os.path.join(cfg_dir, d)):
            continue
        key = d.replace(" ", "")
        if key in seen:
            dup.append(f"{seen[key]} 与 {d}")
        else:
            seen[key] = d
    check("[配置库] 系列目录名无近似重复（去空格后唯一）", not dup, "; ".join(dup))

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


def verify_ui_quick_actions():
    """工具箱「快速操作」按钮表门禁（2026-10-24 用户实测"底部按钮被窗口裁掉"修复后新增）。

    不依赖 GUI/Excel：直接 import wi_workbench 读模块级 QUICK_ACTIONS 常量
    与 _build_tools_tab 源码，断言——
      1) 恰 12 个按钮（历史全量功能，一个不能少也不能多）；
      2) 按钮文字互不重复；
      3) 每个条目声明的 App 方法真实存在（构建时 getattr 不会 AttributeError）；
      4) 排布是两列（divmod 行优先），且带 Canvas 纵向滚动容器＋滚轮绑定——
         防止有人图省事改回单列纵排，把小屏幕上又被裁回来。
    """
    import inspect

    import wi_workbench as wiw

    qa = wiw.QUICK_ACTIONS
    check("[UI] 快速操作按钮恰 12 个", len(qa) == 12, f"实际 {len(qa)} 个")
    texts = [t for t, _ in qa]
    dup = sorted({t for t in texts if texts.count(t) > 1})
    check("[UI] 按钮文字互不重复", not dup, "; ".join(dup))
    missing = [f"「{t}」→{n}" for t, n in qa
               if not callable(getattr(wiw.App, n, None))]
    check("[UI] 每个按钮声明的 App 方法真实存在", not missing, "; ".join(missing))
    body = inspect.getsource(wiw.App._build_tools_tab)
    check("[UI] 快速操作按两列排布（divmod 行优先）", "divmod(i, 2)" in body)
    check("[UI] 工具箱页带纵向滚动容器＋滚轮绑定（小屏不裁切）",
          "tk.Canvas(tab" in body and "MouseWheel" in body)
    check("[UI] 构建取自模块级 QUICK_ACTIONS（非内联散列表）",
          "QUICK_ACTIONS" in body and "quick_actions = [" not in body)


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
        verify_generated_layout(tmp, 10, 1)      # 超过 8 步：证明"步骤不限"没有隐性上限
        verify_image_geometry(tmp)               # 版面几何：吞字/位置/等宽/列宽冲突防线
        verify_normalize(tmp)
        verify_change_request(tmp)
        verify_change_request_fallback(tmp)      # 登记兜底路径：无栏/栏满/重复追加
        verify_no_step_mark(tmp)                 # 无步号页：按图＋完整描述分步并换版头
        verify_configs(tmp)
        verify_ui_quick_actions()          # 工具箱按钮表：12 项两列排布＋滚动容器
    finally:
        shutil.rmtree(tmp, ignore_errors=True)

    print("\n" + ("全部通过 ✅" if not FAILS else f"有 {len(FAILS)} 项失败 ❌：" + "，".join(FAILS)))
    return 0 if not FAILS else 1


if __name__ == "__main__":
    sys.exit(main())
