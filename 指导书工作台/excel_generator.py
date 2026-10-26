# -*- coding: utf-8 -*-
"""
Excel 生成器核心模块
策略：直接以《作业指导书版头.xlsx》为底本打开，把数据填进对应格子；
     工序页通过复制模板内的标准工序表生成——格式、合并单元格、
     列宽行高全部原样保留。
"""
from __future__ import annotations
from typing import List, Optional
import re
import io
import copy
import shutil
from pathlib import Path

try:
    import openpyxl
    from openpyxl.styles import Alignment, Font
    from openpyxl.drawing.image import Image as XLImage
    from openpyxl.drawing.spreadsheet_drawing import OneCellAnchor, TwoCellAnchor, AnchorMarker
    from openpyxl.drawing.xdr import XDRPositiveSize2D
    from openpyxl.utils.units import pixels_to_EMU
except ImportError:
    print("⚠ 需安装 openpyxl: pip install openpyxl")
    openpyxl = None
    XLImage = None

from wi_models import WorkInstruction, Header, Process, RevisionRecord

# 1 像素 = 914400 EMU (at 96 DPI)
PIXELS_TO_EMU_FACTOR = 914400


# ============================================================
# 标准版头模板坐标映射（依据《作业指导书版头.xlsx》实测结构）
# ============================================================

COVER_SHEET_KEY = "清单"           # 封面/清单页（第1个sheet，名称含「清单」）
PROCESS_TMPL_KEY = "新格式"         # 工序页标准样式表

COVER_CELLS = {                     # 封面页版头信息
    "file_no": "I2", "edition": "I3",
    "made_by": "I4", "checked_by": "I5",
    "approved_by": "I6", "issued_date": "I7",
    "note": "B4",
}
COVER_REVISION_FIRST_ROW = 9           # 变更记录首行
COVER_REVISION_LAST_ROW = 12           # 变更记录末行（模板容量 4 条，超出自动扩容）
COVER_REVISION_ROWS = range(COVER_REVISION_FIRST_ROW, COVER_REVISION_LAST_ROW + 1)  # A=日期 B(合并B:I)=内容
COVER_PROCESS_FIRST_ROW = 15            # A=工序号 B(合并B:I)=工位名称 J=工时
COVER_PROCESS_LAST_ROW = 43             # J44 为 =SUM 公式行（模板容量 29 个，超出自动扩容）
COVER_SUM_ROW = 44                      # 合计行（随扩容下移，公式区间同步重写）

PROC_CELLS = {                          # 工序页（新格式）
    "product_spec": "G2",
    "product_name": "G3",
    "process_title": "G4",
    "work_hours": "J2",
    "page_no": "L2",
    "file_no": "K3",
    "process_no": "K4",
    "accessories": "K5",
    "protected_tools": "K6",
}
# 四块步骤区：(起始列, 标题行)。块内布局 = 标题1行 + 图片区12行 + 文字区4行
PROC_STEP_BLOCKS = [("A", 5), ("E", 5), ("A", 22), ("E", 22)]
PROC_IMG_ROWS = 12
PROC_TXT_ROWS = 4
PROC_IMG_COLS = [("A", "B", "C", "D"), ("E", "F", "G", "H"),
                 ("A", "B", "C", "D"), ("E", "F", "G", "H")]
# 步骤区几何：每「段(band)」= 1 标题行 + 12 图片行 + 4 文字行 = 17 行，含左右两列块。
# 模板自带 2 段（第0段标题行5、第1段标题行22）。工序步数超过 4 时，按需向下克隆整段版式，
# 页脚签字行随之推到所有步骤段之后——一道工序仍排在一页内，不强拆成 4 步、也不丢步骤。
STEP_BAND_ROWS = 1 + PROC_IMG_ROWS + PROC_TXT_ROWS   # 17
STEP_BAND_FIRST_LABEL_ROW = 5                         # 第0段标题行
STEP_BAND_STYLE_SRC_LABEL_ROW = 22                    # 克隆版式取第1段（下半段样式）
PROC_IMG_MAX_PER_STEP = 4           # 同一步最多几张图（受列数限制，每步最多4张横向并排）
PROC_EQUIP_ROWS = range(9, 15)      # I=型号/编号 K=名称 L=数量
PROC_MATERIAL_ROWS = range(18, 26)  # I=编号 K=名称/规格 L=数量
PROC_QC_ROWS = range(29, 32)        # I=项目 K=标准 L=参数
# 页脚三栏签署（模板原为五栏，重排为 编制/审核/批准）
PROC_SIGN_GROUPS = [(1, 3, "Made by编制"), (5, 7, "Checked by审核"),
                    (9, 11, "Approved by批准")]   # (标签起始列号, 值起始列号, 标签)
# 一页版式的近似行数上限：工序页 max_row 不超过它时视为「约一页内容」，
# 打印设 fitToHeight=1 缩进单页——页脚不会溢出成只有两行签字的废页
# （2026-10-26 ZT 实测 35 页里 8 页是这种孤页）；超过则 fitToHeight=0
# 允许跨页，长工序硬压单页会把版头纵向压缩（B1/B3 修过的现象）。
ONE_PAGE_MAX_ROW = 45

_EN_LETTER = re.compile(r"[A-Za-z]")

_ILLEGAL_SHEET_CHARS = re.compile(r"[\[\]:*?/\\]")


def _fmt_qty(q) -> str:
    try:
        f = float(q)
        return str(int(f)) if f == int(f) else f"{f:g}"
    except (TypeError, ValueError):
        return str(q or "")


class ExcelGenerator:
    """作业指导书 Excel 生成器（模板原位填充）"""

    def __init__(self, template_path: str):
        if not openpyxl:
            raise RuntimeError("openpyxl 未安装")
        self.template_path = template_path
        self._temp_images: list = []  # 临时图片文件，生成后清理

    def generate(self, wi: WorkInstruction, output_path: str,
                 template_type: str = "auto") -> str:
        warnings: List[str] = []
        wb = openpyxl.load_workbook(self.template_path)

        cover = self._find_sheet(wb, COVER_SHEET_KEY)
        proc_tmpl = self._find_sheet(wb, PROCESS_TMPL_KEY)
        if cover is None or proc_tmpl is None:
            raise RuntimeError(
                f"模板结构不符：需要「清单」封面页和「新格式」工序页，"
                f"实际包含: {wb.sheetnames}")

        self._fill_cover(cover, wi, warnings)

        # 清理模板文件：只保留封面页和工序模板页，移除多余的模板 sheets
        template_sheets_to_keep = {cover, proc_tmpl}
        for ws in list(wb.worksheets):
            if ws not in template_sheets_to_keep:
                wb.remove(ws)

        used = set(wb.sheetnames)
        keep = {cover}
        for page_no, process in enumerate(wi.processes, 1):
            # 版头页码按工序顺序递增：解析从不产 page_no，Process.page_no
            # 默认又是 1，不重编的话每一页版头都标「1」（2026-10-26 ZT 实测
            # 27 页全为 1）
            process.page_no = page_no
            ws = wb.copy_worksheet(proc_tmpl)
            ws.title = self._unique_sheet_name(process.process_no, used)
            self._fill_process_sheet(ws, wi, process, warnings)
            keep.add(ws)

        for ws in list(wb.worksheets):
            if ws not in keep:
                wb.remove(ws)
        cover.title = "清单"

        if wi.change_records:
            self._write_change_sheet(wb, wi)

        Path(output_path).parent.mkdir(parents=True, exist_ok=True)
        wb.save(output_path)

        try:
            self._fix_image_dimensions(output_path)
        except Exception:
            # 后处理失败不影响主流程——用户仍获得可用文件（图片位置可能需手动微调）
            pass

        # 清理临时图片文件
        if hasattr(self, '_temp_images'):
            for tmp in self._temp_images:
                tmp.unlink(missing_ok=True)
            self._temp_images.clear()

        print(f"✓ 生成完成: {output_path}")
        for w in warnings:
            print(f"  ⚠ {w}")
        return output_path

    def _insert_cover_rows(self, ws, at: int, count: int, src_row: int) -> None:
        """在清单封面第 at 行处插入 count 个空行，样式/行高/合并区仿自 src_row。

        为什么不能直接用 ws.insert_rows：openpyxl 的 insert_rows 不搬合并区、
        不搬行高、不更新公式——封面里的 B:I 数据合并和合计行 SUM 公式会留在
        原地，数据与版式整体错位。这里全部手工处理：拆合并→从下往上搬
        （值+逐属性样式+行高）→合并区整体下移恢复→新行仿 src_row 补样式。
        """
        if count <= 0 or at < 1:
            return
        max_col = ws.max_column
        max_row = ws.max_row

        def _clone_style(src, dst):
            if src.has_style:
                dst.font = copy.copy(src.font)
                dst.fill = copy.copy(src.fill)
                dst.border = copy.copy(src.border)
                dst.alignment = copy.copy(src.alignment)
                dst.protection = copy.copy(src.protection)
                dst.number_format = src.number_format

        # 1) 合并区：跨插入点的拉长、整体在插入点下方的下移，先拆后补
        moved, stretched = [], []
        for m in list(ws.merged_cells.ranges):
            if m.min_row >= at:
                moved.append((m.min_row, m.min_col, m.max_row, m.max_col))
                ws.unmerge_cells(str(m))
            elif m.max_row >= at:  # 跨插入点：合并区向下生长 count 行
                stretched.append((m.min_row, m.min_col, m.max_row, m.max_col))
                ws.unmerge_cells(str(m))

        # 2) 从下往上搬单元格（值 + 样式 + 行高）；搬空即新行
        for r in range(max_row, at - 1, -1):
            for c in range(1, max_col + 1):
                src, dst = ws.cell(r, c), ws.cell(r + count, c)
                dst.value = src.value
                _clone_style(src, dst)
                src.value = None
            if ws.row_dimensions[r].height:
                ws.row_dimensions[r + count].height = ws.row_dimensions[r].height

        # 3) 恢复合并区
        for r1, c1, r2, c2 in moved:
            ws.merge_cells(start_row=r1 + count, start_column=c1,
                           end_row=r2 + count, end_column=c2)
        for r1, c1, r2, c2 in stretched:
            ws.merge_cells(start_row=r1, start_column=c1,
                           end_row=r2 + count, end_column=c2)

        # 4) 新行仿 src_row：逐格样式 + 行高 + 同幅 B:I 合并（若有）
        span = None
        for m in ws.merged_cells.ranges:
            if m.min_row == src_row and m.max_row == src_row and m.min_col == 2:
                span = (m.min_col, m.max_col)
                break
        for i in range(count):
            r = at + i
            for c in range(1, max_col + 1):
                _clone_style(ws.cell(src_row, c), ws.cell(r, c))
            if ws.row_dimensions[src_row].height:
                ws.row_dimensions[r].height = ws.row_dimensions[src_row].height
            if span:
                ws.merge_cells(start_row=r, start_column=span[0],
                               end_row=r, end_column=span[1])

    # ---------- 封面（清单）页 ----------
    def _fill_cover(self, ws, wi: WorkInstruction, warnings: List[str]):
        h = wi.header
        values = {
            COVER_CELLS["file_no"]: h.file_no,
            COVER_CELLS["edition"]: h.edition,
            COVER_CELLS["made_by"]: h.made_by,
            COVER_CELLS["checked_by"]: h.checked_by,
            COVER_CELLS["approved_by"]: h.approved_by,
            COVER_CELLS["issued_date"]: h.issued_date,
        }
        if h.note:
            values[COVER_CELLS["note"]] = h.note
        for coord, v in values.items():
            if v:  # 非空值才写入
                ws[coord] = v
                self._apply_en_font(ws[coord])
            else:  # 空值或 None，清除单元格（避免模板默认值残留）
                ws[coord] = None

        capacity = list(COVER_REVISION_ROWS)
        # 变更记录超出模板 4 行容量时向下扩容，绝不截断
        # （用户 2026-10-24 实测：DEMO 文件 13 条变更曾被截成 4 条）。
        extra_rev = max(0, len(h.revision_history) - len(capacity))
        if extra_rev:
            self._insert_cover_rows(ws, COVER_REVISION_LAST_ROW + 1, extra_rev,
                                    COVER_REVISION_LAST_ROW)
        for row, rev in zip(range(COVER_REVISION_FIRST_ROW,
                                  COVER_REVISION_LAST_ROW + 1 + extra_rev),
                            h.revision_history):
            self._set(ws, f"A{row}", rev.date)
            self._apply_en_font(ws[f"A{row}"])
            text = rev.content
            if rev.author:
                text = f"{text}（{rev.author}）" if text else rev.author
            self._set(ws, f"B{row}", text)
            self._apply_en_font(ws[f"B{row}"])

        # 工序清单超出 29 行容量时在合计行之前扩容，合计公式随动重写，
        # 有幾個工序列幾個（用户 2026-10-24 实测：60+ 工序被截成 29 个）。
        cap_proc = COVER_PROCESS_LAST_ROW - COVER_PROCESS_FIRST_ROW + 1
        extra_proc = max(0, len(wi.processes) - cap_proc)
        if extra_proc:
            self._insert_cover_rows(ws, COVER_SUM_ROW + extra_rev, extra_proc,
                                    COVER_PROCESS_LAST_ROW + extra_rev)
        proc_first = COVER_PROCESS_FIRST_ROW + extra_rev
        proc_last = COVER_PROCESS_LAST_ROW + extra_rev + extra_proc
        for row, p in zip(range(proc_first, proc_last + 1), wi.processes):
            self._set(ws, f"A{row}", p.process_no)
            self._set(ws, f"B{row}", p.process_name_cn or p.process_name)
            self._set(ws, f"J{row}", p.work_hours if p.work_hours else None)
            self._apply_en_font(ws[f"B{row}"])

        # 合计行随扩容下移后，SUM 区间同步重写（openpyxl 插行不更新公式）
        sum_row = COVER_SUM_ROW + extra_rev + extra_proc
        sum_cell = ws.cell(sum_row, 10)
        if isinstance(sum_cell.value, str) and sum_cell.value.startswith("=SUM"):
            sum_cell.value = f"=SUM(J{proc_first}:J{proc_last})"
        # 行数变多后允许跨页打印，否则 fitToHeight=1 会把整页缩成蚊子版
        if extra_rev or extra_proc:
            try:
                ws.page_setup.fitToHeight = 0
            except Exception:
                pass

    # ---------- 工序页 ----------
    @staticmethod
    def _step_block(i: int):
        """第 i 块步骤的 (标题列, 标题行, 图片列元组)。每段 2 块、高 17 行。"""
        band, side = divmod(i, 2)
        r_lab = STEP_BAND_FIRST_LABEL_ROW + band * STEP_BAND_ROWS
        if side == 0:
            return "A", r_lab, ("A", "B", "C", "D")
        return "E", r_lab, ("E", "F", "G", "H")

    def _ensure_step_bands(self, ws, n_bands: int) -> int:
        """工序步数超过模板自带的 2 段（4 块）时，向下克隆段版式并把页脚签字推到段之后。
        返回页脚起始行。n_bands<=2 时不动模板，返回 39。"""
        band = STEP_BAND_ROWS
        if n_bands <= 2:
            return 39
        delta = (n_bands - 2) * band
        old_r0 = 39
        maxr = ws.max_row

        # 1) 把页脚区（含样式/值/行高/合并）整体下移 delta 行
        snap = []
        for r in range(old_r0, maxr + 1):
            snap.append([(ws.cell(row=r, column=c).value,
                          copy.copy(ws.cell(row=r, column=c)._style))
                         for c in range(1, 14)])
        hsnap = {r: ws.row_dimensions[r].height for r in range(old_r0, maxr + 1)}
        for m in list(ws.merged_cells.ranges):
            if m.min_row >= old_r0:
                ws.unmerge_cells(str(m))
        for r in range(old_r0, maxr + 1):
            for c in range(1, 14):
                ws.cell(row=r, column=c).value = None
        for i, rowdata in enumerate(snap):
            tgt = old_r0 + i + delta
            for c, (val, st) in enumerate(rowdata, start=1):
                cell = ws.cell(row=tgt, column=c)
                cell._style = st
                if val is not None:
                    cell.value = val
        for src_r, h in hsnap.items():
            if h is not None:
                ws.row_dimensions[src_r + delta].height = h
        footer_row = old_r0 + delta

        # 2) 用模板第1段（行22-38）的版式克隆出第2段及以后（仅 A-H 列样式+行高）
        src_r0 = STEP_BAND_STYLE_SRC_LABEL_ROW
        for b in range(2, n_bands):
            tgt_r0 = STEP_BAND_FIRST_LABEL_ROW + b * band
            for rr in range(band):
                srow = src_r0 + rr
                trow = tgt_r0 + rr
                if ws.row_dimensions[srow].height is not None:
                    ws.row_dimensions[trow].height = ws.row_dimensions[srow].height
                for c in range(1, 9):
                    ws.cell(row=trow, column=c)._style = copy.copy(
                        ws.cell(row=srow, column=c)._style)

        # 3) 清掉新段所在行右侧(I-M)残留的页脚/备注样式，避免出现多余边框
        from openpyxl.styles import Border, PatternFill
        for r in range(old_r0, footer_row):
            for c in range(9, 14):
                cell = ws.cell(row=r, column=c)
                cell.value = None
                cell.border = Border()
                cell.fill = PatternFill()
        return footer_row

    @staticmethod
    def _split_grouped_cols(ws, max_col: int = 8):
        """把第 1~max_col 列范围内「一条定义覆盖多列」的列宽拆成逐列独立定义。

        模板《新格式》把 G、H 两列并成一条 min=7 max=8 的 col 定义（都=15）。
        openpyxl 读 column_dimensions['H'] 拿到的却是默认值 13，不是真实的 15；
        更糟的是只要读过 H，保存时就会写出两条互相冲突的第 8 列定义（15 和 13），
        Excel/WPS 取哪一条不确定——整列可能变窄、版面错乱。这里主动拆开，
        读写一致，从根上消除这个偏差。
        """
        from openpyxl.utils import get_column_letter as _gl
        for key, cd in list(ws.column_dimensions.items()):
            if cd.min is None or cd.max is None or cd.max <= cd.min:
                continue
            if cd.min > max_col:
                continue
            w = cd.width
            for c in range(cd.min, min(cd.max, max_col) + 1):
                tgt = ws.column_dimensions[_gl(c)]
                tgt.width = w          # customWidth 由 width 自动派生，无需手动设
                tgt.min = c
                tgt.max = c
                try:
                    tgt.style = cd.style
                except Exception:
                    pass
            # 原分组条目：超出 max_col 的剩余列保留覆盖，否则删除整条
            if cd.max > max_col:
                cd.min = max_col + 1
            else:
                del ws.column_dimensions[key]

    @staticmethod
    def _col_at(ws, cols, px: float):
        """px（相对块左缘的像素数）落在块内哪一列、列内偏移多少。

        返回 (列字母, 列内偏移px)，偏移恒在 [0, 该列宽) 内——图片锚点用，
        保证 colOff 合法、不依赖跨列偏移这种各渲染器行为不一的手法。
        """
        acc = 0.0
        for c in cols:
            w = (ws.column_dimensions[c].width or 8.43) * 7.5
            if px < acc + w:
                return c, px - acc
            acc += w
        return cols[-1], 0.0

    def _fill_process_sheet(self, ws, wi: WorkInstruction,
                            p: Process, warnings: List[str]):
        from openpyxl.worksheet.properties import PageSetupProperties
        ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
        ws.page_setup.fitToWidth = 1
        ws.page_setup.fitToHeight = 1
        h = wi.header
        wrap = Alignment(wrap_text=True, vertical="center")

        def put(coord, value, do_wrap=False):
            if value in (None, ""):
                return
            self._set(ws, coord, value)
            if do_wrap:
                ws[coord].alignment = wrap
            self._apply_en_font(ws[coord])

        put(PROC_CELLS["product_spec"], h.product_spec or h.product_spec_cn, True)
        put(PROC_CELLS["product_name"], h.product_spec_cn, True)
        title = "\n".join(x for x in (p.process_name, p.process_name_cn) if x)
        put(PROC_CELLS["process_title"], title, True)
        hours_disp = (p.work_hours_display or "").strip().lstrip("约").strip()
        put(PROC_CELLS["work_hours"], hours_disp or f"{p.work_hours}秒")
        put(PROC_CELLS["page_no"], p.page_no)
        put(PROC_CELLS["file_no"], h.file_no)
        put(PROC_CELLS["process_no"], p.process_no)
        put(PROC_CELLS["accessories"], p.accessories)

        steps = list(p.steps)
        n_blocks = max(len(steps), 1)
        n_bands = (n_blocks + 1) // 2
        footer_row = self._ensure_step_bands(ws, n_bands)
        # 根治 H 列宽读取偏差：模板把 G/H 并成一条 col 定义，openpyxl 读 H 会拿到
        # 默认值 13 而非真实 15，且保存时写出两条冲突定义——先拆成逐列独立定义
        self._split_grouped_cols(ws)
        # 整张工序页（含扩展出来的多段步骤）缩放到一页：一道工序不拆成多页
        ws.page_setup.fitToWidth = 1
        ws.page_setup.fitToHeight = 1
        if n_bands > 4:
            warnings.append(f"工序 {p.process_no} 有 {len(steps)} 步（{n_bands} 段），"
                            f"打印时整页缩放会很小，建议拆分工序或精简配图")

        bodies = ["\n".join(x for x in (s.content_en, s.content) if x)
                  for s in steps]
        for i in range(n_blocks):
            col0, r_lab, cols = self._step_block(i)
            lab = ws[f"{col0}{r_lab}"]
            if i < len(steps):
                self._set(ws, f"{col0}{r_lab}", f"Step {i + 1}第{i + 1}步：")
                self._ensure_black_font(lab)
                self._apply_en_font(lab)
                self._put_step_text(ws, cols, r_lab, bodies[i])
            else:
                lab.value = None
        self._embed_step_images(ws, p, warnings, n_blocks)

        self._fill_rows(ws, PROC_EQUIP_ROWS, p.equipment, ("I", "K", "L"),
                        lambda e: (e.code or "编号未知",
                                   f"{e.name}\n{e.operating_time}" if e.operating_time else e.name,
                                   ""),
                        "设备/工夹具", p, warnings, is_equipment=True)
        self._fill_rows(ws, PROC_MATERIAL_ROWS, p.materials, ("I", "K", "L"),
                        lambda m: (m.code,
                                   " ".join(x for x in (m.name, m.spec) if x),
                                   f"{_fmt_qty(m.qty)}{m.unit}".strip()),
                        "物料/辅料", p, warnings)
        self._fill_rows(ws, PROC_QC_ROWS, p.qc_items, ("I", "K", "L"),
                        lambda q: (q.item, q.spec, q.parameter),
                        "质检项目", p, warnings)

        self._fill_sign_footer(ws, p, h, footer_row)

    # ---------- 步骤文字区 / 页脚 ----------
    @staticmethod
    def _apply_en_font(cell):
        """凡含英文字母的文本格统一 Times New Roman（保留原字号加粗）"""
        if _EN_LETTER.search(str(cell.value or "")):
            f = cell.font.copy()
            f.name = "Times New Roman"
            cell.font = f

    @staticmethod
    def _ensure_black_font(cell):
        """强制黑色字体，避免继承模板单元格的红色"""
        from openpyxl.styles import Color
        f = cell.font.copy()
        f.color = Color(rgb="FF000000")
        cell.font = f

    @staticmethod
    def _disp_len(text: str) -> float:
        return sum(2 if ord(ch) > 127 else 1.1 for ch in text)

    def _put_step_text(self, ws, cols, r_lab, body: str):
        """文字区=图片区下方 4 行，合并整块；文字放不下就自动加高首行"""
        t0 = r_lab + 1 + PROC_IMG_ROWS
        t1 = r_lab + PROC_IMG_ROWS + PROC_TXT_ROWS
        rng = f"{cols[0]}{t0}:{cols[3]}{t1}"
        from openpyxl.utils import column_index_from_string as _ci
        c0i, c3i = _ci(cols[0]), _ci(cols[3])
        for m in list(ws.merged_cells.ranges):
            if (m.min_row >= t0 and m.max_row <= t1
                    and m.min_col >= c0i and m.max_col <= c3i
                    and str(m) != rng):
                ws.unmerge_cells(str(m))
        if str(rng) not in [str(x) for x in ws.merged_cells.ranges]:
            ws.merge_cells(rng)
        anchor = f"{cols[0]}{t0}"
        if body:
            self._set(ws, anchor, body)
        cell = ws[anchor]
        cell.alignment = Alignment(wrap_text=True, vertical="top", horizontal="left")
        self._ensure_black_font(cell)
        self._apply_en_font(cell)
        if not body:
            return
        width_chars = sum(ws.column_dimensions[c].width or 8.43 for c in cols)
        cap = max(width_chars / 2 - 2, 10)
        lines = sum(max(1, int(self._disp_len(seg) / cap) + 1)
                    for seg in body.split("\n"))
        needed = lines * 14.5 + 6
        cur = sum(ws.row_dimensions[r].height or 17.65 for r in range(t0, t1 + 1))
        if needed > cur:
            ws.row_dimensions[t0].height = (ws.row_dimensions[t0].height or 17.65) + (needed - cur)

    def _fill_sign_footer(self, ws, p, h, footer_row: int = 39):
        """页脚重排为三栏：Made by编制 / Checked by审核 / Approved by批准"""
        r0, r1 = footer_row, footer_row + 1
        for m in list(ws.merged_cells.ranges):
            if m.min_row >= r0 and m.max_row <= r1:
                ws.unmerge_cells(str(m))
        for row in (r0, r1):
            for c in range(1, 13):
                ws.cell(row=row, column=c).value = None
        from openpyxl.utils import get_column_letter as _gl
        values = [(p.made_by or h.made_by, p.checked_by or h.checked_by,
                   p.approved_by or h.approved_by),
                  (p.process_date or h.issued_date,) * 3]
        labels = [("Made by编制", "Checked by审核", "Approved by批准"),
                  ("Date日期", "Date日期", "Date日期")]
        for r_i, row in enumerate((r0, r1)):
            for g_i, (lab_col, val_col, _) in enumerate(PROC_SIGN_GROUPS):
                ws.merge_cells(f"{_gl(lab_col)}{row}:{_gl(lab_col + 1)}{row}")
                ws.merge_cells(f"{_gl(val_col)}{row}:{_gl(val_col + 1)}{row}")
                lc, vc = f"{_gl(lab_col)}{row}", f"{_gl(val_col)}{row}"
                ws[lc] = labels[r_i][g_i]
                v = values[r_i][g_i]
                if v:
                    ws[vc] = v
                ws[lc].alignment = Alignment(wrap_text=True, vertical="center")
                ws[vc].alignment = Alignment(wrap_text=True, vertical="center")
                self._apply_en_font(ws[lc])
                self._apply_en_font(ws[vc])

    @staticmethod
    def _fill_rows(ws, rows, items, cols, mapper, label, p, warnings, is_equipment=False):
        rows = list(rows)
        if len(items) > len(rows):
            warnings.append(f"工序 {p.process_no} 的{label}有 {len(items)} 项，"
                            f"表格仅 {len(rows)} 行，超出项并入最后一格")
        from openpyxl.styles import PatternFill
        red_fill = PatternFill(start_color="FFFF0000", end_color="FFFF0000", fill_type="solid")
        for row, item in zip(rows, items):
            a, b, c = mapper(item)
            lines = 1
            for col, val in zip(cols, (a, b, c)):
                ExcelGenerator._set(ws, f"{col}{row}", val)
                ExcelGenerator._apply_en_font(ws[f"{col}{row}"])
                if val:
                    cell = ws[f"{col}{row}"]
                    cell.alignment = Alignment(wrap_text=True, vertical="center")
                    cap = max((ws.column_dimensions[col].width or 8.43) / 2, 4)
                    lines = max(lines, sum(max(1, int(ExcelGenerator._disp_len(seg) / cap) + 1)
                                           for seg in str(val).split("\n")))
                    # 设备编号为空或显示"编号未知"时标记红色背景
                    if is_equipment and col == cols[0] and (not val or val == "编号未知"):
                        cell.fill = red_fill
                        cell.font = cell.font.copy(color="FFFF0000")
            need = lines * 14.5 + 4
            cur = ws.row_dimensions[row].height or 17.65
            if need > cur:
                ws.row_dimensions[row].height = need
        if len(items) > len(rows):
            row = rows[-1]
            anchor = f"{cols[0]}{row}"
            extra = "\n".join(" ".join(str(x) for x in mapper(it) if x)
                              for it in items[len(rows):])
            ws[anchor] = f"{ws[anchor]}\n{extra}" if ws[anchor].value else extra
            ws[anchor].alignment = Alignment(wrap_text=True, vertical="top")

    # ---------- 步骤配图 ----------
    def _embed_step_images(self, ws, p: Process, warnings: List[str],
                           n_blocks: Optional[int] = None):
        """把每一步的 images 横向并排插进对应块"""
        if XLImage is None:
            if any(getattr(s, "images", None) for s in p.steps):
                warnings.append(f"工序 {p.process_no} 配有图片，但 openpyxl/Pillow 不可用，已跳过")
            return
        steps = list(p.steps)
        n_blocks = n_blocks if n_blocks is not None else len(PROC_STEP_BLOCKS)
        # 收集每块图片：每一步的图片归入自己那一块（不再把第 5 步起并入第 4 段）
        block_imgs = [[] for _ in range(n_blocks)]
        for i, s in enumerate(steps):
            raw = getattr(s, "images", None) or []
            if isinstance(raw, str):
                raw = re.split(r"[;|,]", raw)
            paths = [str(q).strip() for q in raw if str(q).strip()]
            bi = i if i < n_blocks else n_blocks - 1
            block_imgs[bi].extend(paths)
        # 落单图（锚点落在步骤块外）按当前图片最少的块补齐，保证零丢失
        for q in (getattr(p, "_extracted_images", None) or []):
            q = str(q).strip()
            if not q:
                continue
            bi = min(range(n_blocks), key=lambda k: len(block_imgs[k]))
            block_imgs[bi].append(q)
        for i in range(n_blocks):
            paths = block_imgs[i]
            if not paths:
                continue
            col0, r_lab, cols = self._step_block(i)
            ncols = len(cols)
            # 图片区可用高度按真实行高逐行累加（pt→px），上下共留 8px 呼吸；
            # 一步多于 ncols 张图时纵向排多行，每排高度均分
            n_total = len(paths)
            n_rows = (n_total + ncols - 1) // ncols
            img_r0, img_r1 = r_lab + 1, r_lab + PROC_IMG_ROWS   # 1-indexed 闭区间
            avail_px = sum(ws.row_dimensions[r].height or 17.65
                           for r in range(img_r0, img_r1 + 1)) * 96 / 72 - 8
            slot_h = max(avail_px / n_rows, 30)      # 每排槽位高(px)
            max_h = max(int(slot_h) - 2, 28)
            rows_per_img = max(PROC_IMG_ROWS // n_rows, 1)
            # 方案甲（等宽整齐）的落地：块内图片按排均分槽宽——一步 1 张图时
            # 图可占到整块宽、2 张时各占半块，3 张以上按列均分；同排图一律
            # 等宽不溢出、互不重叠（H 列宽偏差已由 _split_grouped_cols 根治）
            block_w = sum((ws.column_dimensions[c].width or 8.43) * 7.5 for c in cols)
            per_row = min(n_total, ncols)            # 每排几张
            slot_w = max(block_w / per_row, 40)
            # 图从图区顶行开始：r_lab 是 1-indexed 标题行，0-indexed 图区顶 = r_lab
            # （曾按 r_lab==5→10、else r_lab+5 起锚，平白多空 5 行；band1 还偏 1 行）
            base_row = r_lab

            for j, q in enumerate(paths):
                fp = Path(q)
                if not fp.is_absolute():
                    fp = Path(__file__).resolve().parent / fp
                if not fp.exists():
                    warnings.append(f"工序 {p.process_no} 第{i + 1}步图片不存在：{q}")
                    continue
                try:
                    from PIL import Image as PILImage
                    pil_img = PILImage.open(fp)

                    if pil_img.size[0] > 0 and pil_img.size[1] > 0:
                        ratio_w = (slot_w - 4) / pil_img.size[0]   # 槽内留 4px 间距
                        ratio_h = max_h / pil_img.size[1]
                        ratio = min(ratio_w, ratio_h)
                        new_w = max(int(round(pil_img.size[0] * ratio)), 30)
                        new_h = max(int(round(pil_img.size[1] * ratio)), 30)
                    else:
                        new_w, new_h = pil_img.size

                    if new_h > max_h:
                        new_w = int(new_w * max_h / pil_img.size[1]) if pil_img.size[1] > 0 else new_w
                        new_h = max_h
                    pil_img = pil_img.resize((new_w, new_h), PILImage.LANCZOS)

                    PX_TO_EMU = 914400 // 96   # 9525

                    buf = io.BytesIO()
                    if pil_img.mode == 'RGBA':
                        background = PILImage.new('RGB', pil_img.size, (255, 255, 255))
                        background.paste(pil_img, mask=pil_img.split()[3])
                        pil_img = background
                    pil_img.save(buf, format='JPEG')
                    buf.seek(0)

                    im = XLImage(buf)
                    im._user_width = new_w * PX_TO_EMU
                    im._user_height = new_h * PX_TO_EMU

                    # 图左缘 = 槽内水平居中；锚点列 = 图左缘落在的那一列，
                    # colOff = 图左缘相对该列左缘的偏移（恒在 [0, 列宽) 内）。
                    # colOff/rowOff 单位是 EMU，1px = 9525 EMU。曾经误写 2 * 914400
                    # （=192px），每张图被向右下各推 2 英寸——图砸文字区、盖住步骤
                    # 描述，这是「位置不正确 / 吞字」的最大元凶。
                    slot_start = (j % per_row) * slot_w
                    img_left = slot_start + max(0.0, (slot_w - 4 - new_w) / 2)
                    anchor_col_letter, col_off_px = self._col_at(ws, cols, img_left)
                    anchor_col = ord(anchor_col_letter) - ord("A")
                    anchor_row = base_row + (j // per_row) * rows_per_img
                    marker = AnchorMarker(
                        col=anchor_col,
                        row=anchor_row,
                        colOff=pixels_to_EMU(col_off_px),
                        rowOff=pixels_to_EMU(2))     # 顶对齐：图贴着标题行下方

                    im.anchor = OneCellAnchor(_from=marker)
                    im.anchor.ext = XDRPositiveSize2D(
                        cx=new_w * PX_TO_EMU,
                        cy=new_h * PX_TO_EMU)
                    ws.add_image(im)
                except Exception as e:
                    warnings.append(f"工序 {p.process_no} 第{i + 1}步图片插入失败：{e}")

    def _fix_image_dimensions(self, output_path: str):
        """后处理 xlsx ZIP：对每个 <ext cx='0' cy='0'> 元素，用 PIL 读取实际图片像素尺寸并写入 EMU。"""
        try:
            import zipfile
            from xml.etree import ElementTree as ET
            import tempfile
            import shutil
            import os as _os_mod

            SD = 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing'
            A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
            R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
            PKG_NS = 'http://schemas.openxmlformats.org/package/2006/relationships'

            with zipfile.ZipFile(output_path, 'r') as zin:
                namelist = set(zin.namelist())

                # 读取所有需要的内容到内存（drawing XML + rels + media）
                raw_drawings = {}
                raw_rels = {}
                media_files = {}
                for name in namelist:
                    if name.startswith('xl/drawings/') and name.endswith('.xml') and '_rels' not in name:
                        raw_drawings[name] = zin.read(name)
                    elif '/_rels/' in name and name.endswith('.rels'):
                        raw_rels[name] = zin.read(name)
                    elif name.startswith('xl/media/'):
                        media_files[name] = zin.read(name)

            if not raw_drawings:
                return

            tmp_path = None
            try:
                for drawing_name, raw_xml in raw_drawings.items():
                    root = ET.fromstring(raw_xml)

                    # iter() 同时支持默认 namespace 和带前缀的格式
                    anchors = list(root.iter(f'{{{SD}}}oneCellAnchor'))
                    if not anchors:
                        anchors = list(root.iter(f'{{{SD}}}twoCellAnchor'))
                    if not anchors:
                        continue

                    changed_any = False
                    for anchor in anchors:
                        ext_elem = anchor.find(f'{{{SD}}}ext')
                        if ext_elem is None:
                            continue
                        cx = ext_elem.get('cx', '')
                        cy = ext_elem.get('cy', '')
                        # 只修 cx=0 且 cy=0 的元素（openpyxl 默认值）
                        if cx != '0' or cy != '0':
                            continue

                        # 找 blip → rId
                        pic = anchor.find(f'.//{{{SD}}}pic')
                        blip = pic.find(f'.//{{{A}}}blip') if pic is not None else None
                        if blip is None:
                            continue
                        r_id = blip.get(f'{{{R}}}embed')

                        # rId → media file path via rels
                        rels_key = f'{_os_mod.path.dirname(drawing_name)}/_rels/{_os_mod.path.basename(drawing_name)}.rels'
                        media_path = None
                        rels_raw = raw_rels.get(rels_key)
                        if rels_raw:
                            rels_root = ET.fromstring(rels_raw)
                            for rel in rels_root.findall(f'{{{PKG_NS}}}Relationship'):
                                if rel.get('Id') == r_id:
                                    target = rel.get('Target', '').lstrip('/')
                                    if target in media_files:
                                        media_path = target
                                    break

                        # 如果 rels 查找失败，尝试 image{N}.jpeg 模式
                        if media_path is None:
                            num = ''.join(filter(str.isdigit, r_id or ''))
                            candidate = f'xl/media/image{num}.jpeg'
                            if candidate in media_files:
                                media_path = candidate

                        if not media_path:
                            continue

                        try:
                            from PIL import Image
                            img = Image.open(io.BytesIO(media_files[media_path]))
                            w, h = img.size
                            # 像素 → EMU：1px = 9525 EMU（914400/96）。
                            # 曾经误写 w * 914400（把像素当英寸），一旦触发就把图放大 96 倍
                            ext_elem.set('cx', str(w * 9525))
                            ext_elem.set('cy', str(h * 9525))
                            changed_any = True
                        except Exception:
                            pass

                    if changed_any:
                        raw_drawings[drawing_name] = ET.tostring(root, encoding='unicode')

                # 重写 ZIP
                tmp_path = tempfile.mktemp(suffix='.xlsx')
                with zipfile.ZipFile(output_path, 'r') as zin_r:
                    with zipfile.ZipFile(tmp_path, 'w', zipfile.ZIP_DEFLATED) as zout:
                        for item in zin_r.infolist():
                            data = zin_r.read(item.filename)
                            if item.filename in raw_drawings and isinstance(raw_drawings[item.filename], str):
                                data = raw_drawings[item.filename].encode('utf-8')
                            zout.writestr(item, data)

                shutil.move(tmp_path, output_path)
                tmp_path = None  # moved successfully
            finally:
                if tmp_path and _os_mod.path.exists(tmp_path):
                    _os_mod.unlink(tmp_path)
        except Exception:
            pass

    # ---------- 辅助 ----------
    @staticmethod
    def _set(ws, coord, value):
        """写单元格；若目标落在合并区非左上角，先拆分该合并区再写"""
        from openpyxl.cell.cell import MergedCell
        cell = ws[coord]
        if isinstance(cell, MergedCell):
            for rng in list(ws.merged_cells.ranges):
                if coord in rng:
                    ws.unmerge_cells(str(rng))
                    break
            cell = ws[coord]
        cell.value = value

    @staticmethod
    def _find_sheet(wb, keyword: str):
        for name in wb.sheetnames:
            if keyword in name:
                return wb[name]
        return None

    @staticmethod
    def _unique_sheet_name(base, used: set) -> str:
        """使用工序号作为sheet名，避免与现有sheet名冲突"""
        name = _ILLEGAL_SHEET_CHARS.sub("", str(base or "")).strip()[:31] or "工序"
        if name in used:
            i = 2
            while f"{name[:28]}-{i}" in used:
                i += 1
            name = f"{name[:28]}-{i}"
        used.add(name)  # 添加后防止后续重复
        return name

    @staticmethod
    def _write_change_sheet(wb, wi: WorkInstruction):
        ws = wb.create_sheet("变更记录")
        ws.append(["修改内容", "审核", "批准", "日期"])
        for r in wi.change_records:
            ws.append([r.modification, r.checked_by, r.approved_by, r.date])


# ============================================================
# 单文件规范化器：把旧文件转为标准结构
# ============================================================

class Normalizer:
    """单文件规范化：读取任意旧版作业指导书，提取数据，按标准模板重写"""

    def __init__(self, template_path: str):
        self.generator = ExcelGenerator(template_path)

    _STEP_MARK = re.compile(r"^(第[一二三四五六七八九十\d]+步|Step\s*\d+)",
                            re.IGNORECASE)

    # 版头特征字样：无步号页面往上吞"该步描述"时的停止线（用户 2026-10-22 口径）
    _HEADER_HINTS = (
        "公司", "Company", "Working Instruction", "作业指导书",
        "文件编号", "File NO", "Edition", "版本",
        "编制", "Made by", "审核", "Checked by", "批准", "Approved by",
        "工序号", "Working procedure", "工位名称", "Testing part",
        "页码", "Page number", "工时", "Working hour",
        "附属文件", "Accessories",
    )

    # 版头标签字样：字段右侧取到值时的防污染名单。_get_right_cell_value
    # 会顺延右侧 1-5 格取首个非空——字段本身无值时，抓到的是下一列的
    # 标签（如"文件编号File NO"），把它回填进新版头就乱套了。
    # 2026-10-23 用户实测（9STT RC-1）：工位名称无值 → G4 被
    # "文件编号File NO"污染。
    _LABEL_WORDS = (
        "文件编号", "File NO", "工序号", "Working procedure",
        "工位名称", "Testing part", "产品规格", "Specification",
        "产品名称", "页码", "Page number", "工时", "Working hour",
        "编制", "Made by", "审核", "Checked", "批准", "Approved",
        "Video Number", "检验依据", "劳保用具", "附属文件",
        "Accessories", "Edition", "版本",
    )

    @classmethod
    def _looks_like_label(cls, text: str) -> bool:
        """取到的"值"是不是下一个标签的字样（顺延抓取污染防护）。"""
        return any(w in text for w in cls._LABEL_WORDS)

    @classmethod
    def _image_anchor_rows(cls, ws) -> list:
        """页面所有配图的锚点行（1-indexed，去重升序）。"""
        rows = set()
        for im in getattr(ws, "_images", []):
            fr = getattr(im.anchor, "_from", None)
            if fr is not None and getattr(fr, "row", None) is not None:
                rows.add(fr.row + 1)
        return sorted(rows)

    @classmethod
    def _first_step_row_by_image(cls, ws) -> Optional[int]:
        """无步号工序页的正文起点（用户 2026-10-22 口径：按"图＋完整描述"为一步）。

        从第一张配图的行往上，吞掉属于该步的描述文字；遇到空行或版头特征行就停。
        没有配图返回 None——调用方按老行为告警、不换版头：
        宁可不换版头，也绝不能把正文当版头删掉（硬规则①）。"""
        img_rows = cls._image_anchor_rows(ws)
        if not img_rows:
            return None
        r = img_rows[0]
        while r > 2:
            # 只看左侧 8 列（步骤内容区）：右侧的设备/工夹具、附属文件等表格
            # 不算步骤描述，不许把正文起点往上拽——2026-10-22 验收抓到过：
            # 右表标签正好压在第一条描述上一行，整行扫描会把版头区少算一行，
            # 旧表格残在新版头下面。左 8 列与 _first_step_row 找步号同一个口径。
            # 只读 _cells、不新建空单元格（openpyxl 一读就建的坑，
            # 详见 change_request._peek 的注释）。
            parts = []
            for c in range(1, 9):
                cell = ws._cells.get((r - 1, c))
                if cell is not None and isinstance(cell.value, str) and cell.value.strip():
                    parts.append(cell.value.strip())
            line = " ".join(parts)
            if not line or any(h in line for h in cls._HEADER_HINTS):
                break
            r -= 1
        return r

    def _first_step_row(self, ws) -> Optional[int]:
        """返回第一个步骤标记所在行（1-indexed），只看左侧 8 列前 12 行；找不到返回 None。"""
        for row in ws.iter_rows(min_row=1, max_row=min(ws.max_row, 12), max_col=8):
            for cell in row:
                if isinstance(cell.value, str) and self._STEP_MARK.match(cell.value.strip()):
                    return cell.row
        return self._first_step_row_by_image(ws)

    def _replace_header(self, ws, ref, fsr, ref_hdr_rows, wi, p, copy):
        """把 ws 顶部版头条（第 1..fsr-1 行）替换为参考模板 ref 的版头样式，
        并回填该工序页字段值。第 fsr 行往下（步骤/图片/文字/表格/页脚）完全不动。"""
        from openpyxl.cell.cell import MergedCell
        hdr_end = fsr - 1
        max_col = max(ws.max_column, ref.max_column)

        # 1) 删除锚定在版头区的图片（正文区图片保留）
        def _anchor_row(im):
            fr = getattr(im.anchor, '_from', None) or getattr(im.anchor, 'start', None)
            return getattr(fr, 'row', None)
        try:
            ws._images = [im for im in ws._images
                          if _anchor_row(im) is None or _anchor_row(im) >= hdr_end]
        except Exception:
            pass

        # 2) 取消版头区合并并清空值
        for m in list(ws.merged_cells.ranges):
            if m.min_row <= hdr_end:
                ws.unmerge_cells(str(m))
        for r in range(1, hdr_end + 1):
            for c in range(1, max_col + 1):
                ws.cell(r, c).value = None

        # 3) 从参考版头逐格复制值+样式（标签、公司块、标题），并同步行高
        n = min(hdr_end, ref_hdr_rows)
        for r in range(1, n + 1):
            for c in range(1, max_col + 1):
                src = ref.cell(r, c)
                dst = ws.cell(r, c)
                dst.value = src.value
                if src.has_style:
                    # 跨工作簿复制样式，绝不能直接搬 _style：StyleArray 里的
                    # fontId/alignmentId 等下标是模板簿样式表的，目标簿里没有
                    # 对应条目——保存时轻则张冠李戴套错样式，重则直接
                    # IndexError 崩在 write_stylesheet（2026-10-22 无步号
                    # 验收就这么崩过：目标簿是新建的薄样式表）。逐属性赋值
                    # 才会把字体/填充/边框/对齐/保护/数字格式重新登记进
                    # 目标簿的样式表。同簿内搬样式（如 _ensure_step_bands）
                    # 仍可继续用 _style 直拷。
                    dst.font = copy.copy(src.font)
                    dst.fill = copy.copy(src.fill)
                    dst.border = copy.copy(src.border)
                    dst.alignment = copy.copy(src.alignment)
                    dst.protection = copy.copy(src.protection)
                    dst.number_format = src.number_format
            if ref.row_dimensions[r].height:
                ws.row_dimensions[r].height = ref.row_dimensions[r].height

        # 4) 复制参考版头内的合并单元格（落在 1..n 行）
        for m in ref.merged_cells.ranges:
            if m.min_row <= n and m.max_row <= n:
                try:
                    ws.merge_cells(start_row=m.min_row, start_column=m.min_col,
                                   end_row=m.max_row, end_column=m.max_col)
                except Exception:
                    pass

        # 5) 回填字段值（坐标取自参考工序页 PROC_CELLS）
        h = wi.header
        title = "\n".join(x for x in (p.process_name, p.process_name_cn) if x)
        hours = ((p.work_hours_display or "").strip().lstrip("约").strip()
                 or (f"{p.work_hours}秒" if p.work_hours else ""))
        vals = {
            PROC_CELLS["product_spec"]: h.product_spec or h.product_spec_cn,
            PROC_CELLS["product_name"]: h.product_spec_cn,
            PROC_CELLS["process_title"]: title,
            PROC_CELLS["work_hours"]: hours,
            PROC_CELLS["page_no"]: p.page_no,
            PROC_CELLS["file_no"]: h.file_no,
            PROC_CELLS["process_no"]: p.process_no,
        }
        for coord, value in vals.items():
            if value in (None, ""):
                continue
            cell = ws[coord]
            if isinstance(cell, MergedCell):
                for m in ws.merged_cells.ranges:
                    if coord in m:
                        ws.cell(m.min_row, m.min_col).value = str(value)
                        break
            else:
                cell.value = str(value)

    # ---------- 清单封面：老封面 → 模板新格式封面 ----------
    @staticmethod
    def _looks_like_cover(ws) -> bool:
        """兜底判断第 1 张表是不是清单封面（表名不含「清单」时看 A1）。
        工序页 A1 是「作业指导书 Working Instruction」，不含「清单/listing」。"""
        a1 = ws["A1"].value
        if not isinstance(a1, str):
            return False
        low = a1.lower()
        return "清单" in a1 or "listing" in low or "sop list" in low

    def _copy_sheet(self, ref, ws) -> None:
        """跨工作簿整表复制（封面用）：值+样式+合并+行高列宽+页面设置。

        样式绝不能直拷 _style——StyleArray 里的 fontId/fillId 等下标是
        模板簿样式表的，目标簿里没有对应条目，保存时轻则张冠李戴、
        重则 IndexError 崩在 write_stylesheet（2026-10-22 无步号验收
        踩过）。逐属性赋值才会把样式重新登记进目标簿的样式表。
        """
        for r in range(1, ref.max_row + 1):
            for c in range(1, ref.max_column + 1):
                src = ref.cell(r, c)
                dst = ws.cell(r, c)
                dst.value = src.value
                if src.has_style:
                    dst.font = copy.copy(src.font)
                    dst.fill = copy.copy(src.fill)
                    dst.border = copy.copy(src.border)
                    dst.alignment = copy.copy(src.alignment)
                    dst.protection = copy.copy(src.protection)
                    dst.number_format = src.number_format
            if ref.row_dimensions[r].height:
                ws.row_dimensions[r].height = ref.row_dimensions[r].height
        for col, dim in ref.column_dimensions.items():
            if dim.width:
                ws.column_dimensions[col].width = dim.width
        for m in ref.merged_cells.ranges:
            try:
                ws.merge_cells(str(m))
            except Exception:
                pass
        try:
            from openpyxl.worksheet.properties import PageSetupProperties
            ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
            ws.page_setup.fitToWidth = 1
            # 封面含多行工序/变更清单，必须允许竖向跨页，否则 fitToHeight=1
            # 会把整页缩成「蚊子版」，右侧 I–L 版头（工时/页码/文件编号/工序号）被切掉
            ws.page_setup.fitToHeight = 0
            ws.page_setup.orientation = ref.page_setup.orientation
            ws.page_setup.paperSize = ref.page_setup.paperSize
        except Exception:
            pass

    @staticmethod
    def _capture_variant_table(cover, input_path=None):
        """从老封面右侧捕获「变体表」（产品编号/工时/备注等），原样保留到新封面。

        返回 (top_row, { (r,c): (value, font, fill, border, alignment, num_fmt, protection) },
        {col_letter: width})；没找到返回 None。

        判定：在 K:M 列（11..13）找表头行——K 列含「编号/变体/型号」且同行
        L 列含「工时」或 M 列含「备注/说明」，即视为变体表起点。向下收集到
        连续两行全空为止（兼容表头+若干数据行的紧凑布局）。样式从 cover 取；
        值优先用 data_only 读取（变体表工时列常是公式如 =247+50+30+64+25，
        须取显示值 416/436，否则搬过去的是公式），源文件带缓存值时生效。
        """
        import copy as _copy
        from openpyxl.utils import get_column_letter
        # data_only 读取计算值（公式→显示值）；样式仍从 cover 取（两load样式一致）
        dov_cover = None
        if input_path:
            try:
                dov_wb = openpyxl.load_workbook(input_path, data_only=True)
                for ws in dov_wb.worksheets:
                    if ws.title == cover.title:
                        dov_cover = ws
                        break
            except Exception:
                dov_cover = None

        def _val(r, c):
            if dov_cover is not None:
                v = dov_cover.cell(r, c).value
                if v is not None:
                    return v
            return cover.cell(r, c).value

        found = None
        for r in range(1, min(cover.max_row, 30) + 1):
            k = _val(r, 11)
            if isinstance(k, str) and any(w in k for w in ("编号", "变体", "型号")):
                l = _val(r, 12)
                m = _val(r, 13)
                if (isinstance(l, str) and "工时" in l) or (
                        isinstance(m, str) and ("备注" in m or "说明" in m)):
                    found = r
                    break
        if not found:
            return None
        hr = found
        widths = {}
        for c in (11, 12, 13):
            cl = get_column_letter(c)
            d = cover.column_dimensions.get(cl)
            if d and d.width:
                widths[cl] = d.width
        data = {}
        r = hr
        empty_streak = 0
        while r <= cover.max_row and empty_streak < 2:
            vals = [_val(r, 11), _val(r, 12), _val(r, 13)]
            if all(v is None for v in vals):
                empty_streak += 1
                r += 1
                continue
            empty_streak = 0
            for c in (11, 12, 13):
                cell = cover.cell(r, c)  # 样式从非data_only表取
                data[(r, c)] = (
                    _val(r, c),
                    _copy.copy(cell.font) if cell.has_style else None,
                    _copy.copy(cell.fill) if cell.has_style else None,
                    _copy.copy(cell.border) if cell.has_style else None,
                    _copy.copy(cell.alignment) if cell.has_style else None,
                    cell.number_format,
                    _copy.copy(cell.protection) if cell.has_style else None,
                )
            r += 1
        if not data:
            return None
        return hr, data, widths

    def _convert_cover(self, wb, tmpl_wb, cover, wi, warnings: List[str],
                       input_path=None) -> None:
        """把老清单封面整体换成模板的新格式封面，字段按解析结果回填。

        用户 2026-10-23 实测要求：规范化的文件清单页也要换成新格式
        （此前封面明确跳过不动）。映射关系：
          文件编号/版本/编制/审核/批准/发放日期 → I2..I7（_fill_cover）
          变更记录 → A9 起（模板 4 行，超出自动向下扩容，不截断）
          工序清单（工序号/工位名称/工时） → A15 起（模板 29 行，超出自动扩容，
            SUM 合计公式随动重写，不截断）
          变体表（产品编号/工时/备注）→ 老封面右侧 K:M 原样搬回新封面右侧
            （用户 2025-12-16 拍板保留）
        产品规格在新封面没有格子，但已由 _replace_header 写进每张工序页
        版头 G2，不丢。
        表名保持原封面名（含「清单」即可被各处识别），位置仍是第 1 张表。
        """
        ref = ExcelGenerator._find_sheet(tmpl_wb, COVER_SHEET_KEY)
        if ref is None:
            warnings.append("模板缺少清单封面页，封面未转换")
            return
        # 先抓老封面右侧变体表（删除旧表前），换完封面再贴回右侧
        variant = self._capture_variant_table(cover, input_path)
        if len(getattr(cover, "_images", [])):
            warnings.append("老封面带有图片，换新封面后不保留（新封面为模板样式）")
        old_name = cover.title
        del wb[old_name]
        ws = wb.create_sheet(title=old_name, index=0)
        self._copy_sheet(ref, ws)
        self.generator._fill_cover(ws, wi, warnings)
        # 变体表贴回新封面右侧（仅当封面未因超额扩容插行、行号未错位时）
        if variant:
            hr, data, widths = variant
            extra_rev = max(0, len(wi.header.revision_history)
                            - len(COVER_REVISION_ROWS))
            cap_proc = COVER_PROCESS_LAST_ROW - COVER_PROCESS_FIRST_ROW + 1
            extra_proc = max(0, len(wi.processes) - cap_proc)
            if extra_rev == 0 and extra_proc == 0:
                for cl, w in widths.items():
                    ws.column_dimensions[cl].width = w
                for (r, c), (value, font, fill, border, alignment,
                             num_fmt, protection) in data.items():
                    cell = ws.cell(r, c)
                    cell.value = value
                    if font is not None:
                        cell.font = font
                        cell.fill = fill
                        cell.border = border
                        cell.alignment = alignment
                        cell.number_format = num_fmt
                        cell.protection = protection
                print(f"  ✓ 封面变体表已保留至右侧（K{hr}:M）")
            else:
                warnings.append("封面变更/工序超额扩容，变体表行号可能错位，已跳过贴回")
        print(f"  ✓ 清单封面已换为新格式（{old_name}）")

    def normalize(self, input_path: str, output_path: str) -> WorkInstruction:
        """规范化：以源文件为底，把清单封面和各工序页版头换成参考模板样式，
        正文（步骤图片+文字+表格+页脚）原样保留——有几步就是几步，绝不截断。

        封面（清单页）整体替换为模板新格式封面，老封面字段按
        _parse_header/_parse_revisions/工序列表回填（用户 2026-10-23 实测要求）。"""
        import copy
        print(f"正在规范化（正文原样保留）: {input_path}")

        wb = openpyxl.load_workbook(input_path)
        tmpl_wb = openpyxl.load_workbook(self.generator.template_path)
        ref = ExcelGenerator._find_sheet(tmpl_wb, PROCESS_TMPL_KEY)
        if ref is None:
            tmpl_wb.close()
            raise RuntimeError("模板缺少「新格式」工序页，无法规范化版头")
        ref_hdr_rows = (self._first_step_row(ref) or 5) - 1

        cover = ExcelGenerator._find_sheet(wb, COVER_SHEET_KEY) or wb.worksheets[0]
        wi = WorkInstruction()
        wi.header = self._parse_header(cover)
        wi.header.revision_history = self._parse_revisions(cover)

        warnings: List[str] = []
        # 附件类页面（可选项/附件/资料/参考/变体/附录/选项）：正文与版头照常
        # 保留，但【不计入封面工序清单】——用户 2025-12-16 实测明确要求
        # 「附件页（1可选项等）不需要在清单页列出」。按标题关键词排除，
        # 因为所有页（含附件页）都有步骤标记，无法靠「有无步骤」区分。
        ATTACH_KW = ("附件", "可选项", "资料", "参考", "变体", "附录", "选项")
        page_seq = 0        # 版头页码：按页在文档中的顺序递增
        seen_proc_no = {}   # 工序号 -> 首次出现的 sheet 名（重复检测用）
        for ws in wb.worksheets:
            if ws is cover or COVER_SHEET_KEY in ws.title:
                continue
            p = self._parse_process_sheet(ws)
            if not p or not (p.process_no or p.process_name_cn):
                continue
            is_attach = any(kw in ws.title for kw in ATTACH_KW)
            fsr = self._first_step_row(ws)
            if not fsr or fsr <= 1:
                warnings.append(f"工序 {p.process_no or ws.title} 未找到步骤起始行，版头未替换")
                continue
            # 版头无论是否附件都换成标准样式（否则新旧版头混排，且门禁要求
            # 所有工序页版头已换）。附件页仅不进封面清单。
            # 页码重编：解析从不产 page_no，Process.page_no 默认又是 1，
            # 不重编的话每一页版头都标「1」（2026-10-26 ZT 实测 27 页全为 1）。
            # 附件页也是一页纸，同样编号。
            page_seq += 1
            p.page_no = page_seq
            # 重复工序号检测：只对「两页都会进封面清单」的组合警告——
            # 附件页不进清单（A3），与正文页同号不算清单重复
            if (not is_attach and p.process_no
                    and p.process_no in seen_proc_no):
                warnings.append(
                    f"工序号 {p.process_no} 在「{seen_proc_no[p.process_no]}」和"
                    f"「{ws.title}」两页重复出现，封面清单会各登记一行，请人工确认")
            elif not is_attach and p.process_no:
                seen_proc_no[p.process_no] = ws.title
            self._replace_header(ws, ref, fsr, ref_hdr_rows, wi, p, copy)
            try:
                from openpyxl.worksheet.properties import PageSetupProperties
                ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
                ws.page_setup.fitToWidth = 1
                # 清除源文件遗留的打印区域：实测样本自带 Print_Area=$A$1:$H$40，
                # 右侧 I–L 版头（工时/页码/文件编号/工序号）正好被排除在打印
                # 区外——数据都在，打印/导 PDF 时整块消失（2026-10-26 ZT 实测
                # 16 页中招）。清掉后打印范围回到整个使用区域。
                try:
                    ws.print_area = None
                except Exception:
                    pass
                # fitToHeight 按需：fitToHeight=0 让高度自然流动，但「略超一页」
                # 的工序页会把页脚（编制/审核/批准）挤成只有两行字的孤页
                # （2026-10-26 ZT 实测 35 页里 8 页是废页）；近一页内容改回
                # fitToHeight=1 缩进单页消废页，真正多页的长工序仍保持 0
                # 允许跨页（硬压单页会把版头纵向压缩，B1/B3 修过的现象）。
                ws.page_setup.fitToHeight = (1 if ws.max_row <= ONE_PAGE_MAX_ROW
                                             else 0)
            except Exception:
                pass
            if not is_attach:
                wi.add_process(p)

        # 清单封面也换成模板新格式（用户 2026-10-23 实测要求）。
        # 放在工序页循环之后：wi.processes 已收集全，封面工序清单才列得全。
        if COVER_SHEET_KEY in cover.title or self._looks_like_cover(cover):
            self._convert_cover(wb, tmpl_wb, cover, wi, warnings, input_path)

        wb.save(output_path)
        wb.close()
        tmpl_wb.close()
        for w in warnings:
            print("  ⚠ " + w)
        print(f"✓ 规范化完成（正文原样）: {output_path}  工序 {len(wi.processes)} 个")
        return wi

    def _parse_old_file_with_images(self, wb, all_images, input_path) -> WorkInstruction:
        """带图片信息的完整解析"""
        wi = WorkInstruction()

        cover = next((ws for ws in wb.worksheets if "清单" in ws.title),
                     wb.worksheets[0])
        wi.header = self._parse_header(cover)
        wi.header.revision_history = self._parse_revisions(cover)

        img_dir = getattr(self, '_img_output_dir', None)
        if isinstance(img_dir, str):
            # 外部调用方可能给字符串路径；不转 Path 的话下面 img_dir / ...
            # 会 TypeError，又被 except Exception 静默吞掉——配图悄无声息全丢
            img_dir = Path(img_dir)
        if not img_dir and all_images:
            img_dir = Path(input_path).parent / "_norm_imgs" / Path(input_path).stem
        if img_dir is not None:
            # 目录不存在时 write_bytes 会抛异常，又被下面逐图的
            # except Exception 静默吞掉——配图悄无声息全丢。
            # 2026-10-22 无步号验收就是这么丢的，这里先建好目录。
            img_dir.mkdir(parents=True, exist_ok=True)

        for ws in wb.worksheets:
            if ws is cover or "清单" in ws.title:
                continue
            process = self._parse_process_sheet(ws)
            if not process or not (process.process_no or process.process_name_cn):
                continue

            # 按锚点位置把图片归到对应步骤块
            geom = getattr(process, '_step_geom', {}) or {}
            n_assigned = 0
            leftover = []
            for img in ws._images:
                if not hasattr(img, '_data'):
                    continue
                try:
                    data = img._data()
                    if not data:
                        continue
                    img_path = img_dir / f"{id(img)}.png"
                    img_path.write_bytes(data)
                except Exception:
                    continue
                # 取锚点行列（0-indexed）
                arow = acol = None
                try:
                    anc = img.anchor
                    fr = getattr(anc, '_from', None) or getattr(anc, 'start', None)
                    if fr is not None:
                        arow, acol = getattr(fr, 'row', None), getattr(fr, 'col', None)
                except Exception:
                    pass
                idx = None
                if arow is not None and acol is not None:
                    for k, (r0, r1, c0, c1) in geom.items():
                        if r0 <= arow < r1 and c0 <= acol < c1:
                            idx = k
                            break
                if idx is not None and idx < len(process.steps):
                    process.steps[idx].images.append(str(img_path))
                    n_assigned += 1
                else:
                    leftover.append(str(img_path))

            if leftover:
                process._extracted_images = leftover
            # 保留条件：有步骤（含空步骤的附件页）或有图片
            if process.steps or n_assigned:
                wi.add_process(process)

        return wi

    def _parse_old_file(self, path: str) -> WorkInstruction:
        """解析旧版 Excel，提取结构化数据"""
        wb = openpyxl.load_workbook(path, data_only=True)
        wi = WorkInstruction()

        cover = next((ws for ws in wb.worksheets if "清单" in ws.title),
                     wb.worksheets[0])
        wi.header = self._parse_header(cover)
        wi.header.revision_history = self._parse_revisions(cover)

        # 解析工序 sheet（跳过纯附件/资料 sheet）
        for ws in wb.worksheets:
            if ws is cover or "清单" in ws.title:
                continue
            title = ws.title.lower()
            # 跳过明显的非工序页
            if any(kw in title for kw in ["配件附件目录", "protected tools"]):
                # 仍可能是子工序页，继续处理
                pass
            process = self._parse_process_sheet(ws)
            if process and (process.process_no or process.process_name_cn or process.steps):
                wi.add_process(process)

        wb.close()
        return wi

    def _parse_header(self, ws) -> Header:
        """从版头 sheet 按关键字提取基本信息（只看前 8 行）

        两条防抢字段规则（用户 2026-10-24 实测补）：
        1. 标签格都是短字样（"文件编号File NO" 12 字），超过 24 字的长文本
           （变更记录正文等）即使含"版本/编制"字样也不算标签，不许抢字段；
        2. 同一字段先非空匹配获胜——版头块在表格最上方，后出现的同名文本
          （如变更表里提到"文件编号"）不能覆盖它。
        """
        from datetime import datetime, date as pydate
        h = Header()
        seen = set()  # 已确定取值的字段（Header 有非空默认值，不能靠真假判断）

        def take(field: str, value: str):
            if value and field not in seen:
                setattr(h, field, value)
                seen.add(field)

        for row in ws.iter_rows(min_row=1, max_row=min(8, ws.max_row), max_col=15):
            for cell in row:
                if not cell.value or not isinstance(cell.value, str):
                    continue
                val = cell.value.strip()
                # 跳过纯日期格（不应作为文本值处理）
                if re.match(r"^\d{4}[-/年.]\d{1,2}[-/月.]\d{1,2}$", val):
                    continue
                # 长文本不是标签格（变更记录正文动辄半句话）
                if len(val) > 24:
                    continue
                if "文件编号" in val or "File NO" in val:
                    v = self._get_right_cell_value(ws, cell.row, cell.column)
                    if not v:  # 标签和值同格时，取标签字样之后的剩余部分
                        v = val.replace("文件编号", "").replace("File NO", "").strip()
                    take("file_no", v)
                elif "版本" in val or "Edition" in val:
                    v = self._get_right_cell_value(ws, cell.row, cell.column)
                    if not v:
                        v = val.replace("版本", "").replace("Edition", "").strip()
                    take("edition", v)
                elif ("产品规格" in val or "Specification" in val) and "名称" not in val:
                    v = self._get_right_cell_value(ws, cell.row, cell.column)
                    if not v:
                        v = val.replace("产品规格", "").replace("Specification", "").strip()
                    take("product_spec", v)
                elif "产品名称" in val or "Product Name" in val or "Name" in val and "规格" not in val:
                    v = self._get_right_cell_value(ws, cell.row, cell.column)
                    if v:
                        h.product_spec_cn = v
                elif "编制" in val or "Made by" in val:
                    v = self._get_right_cell_value(ws, cell.row, cell.column)
                    if not v:
                        v = val.replace("编制", "").replace("Made by", "").strip()
                    take("made_by", v)
                elif "审核" in val or "Checked" in val:
                    v = self._get_right_cell_value(ws, cell.row, cell.column)
                    if not v:
                        v = val.replace("审核", "").replace("Checked by", "").strip()
                    take("checked_by", v)
                elif "批准" in val or "Approved" in val:
                    v = self._get_right_cell_value(ws, cell.row, cell.column)
                    if not v:
                        v = val.replace("批准", "").replace("Approved by", "").strip()
                    take("approved_by", v)
                elif "发放日期" in val or "Issued" in val:
                    take("issued_date",
                         self._get_right_cell_value(ws, cell.row, cell.column,
                                                    allow_date=True))
        return h

    def _parse_revisions(self, ws) -> List[RevisionRecord]:
        """同一行内「日期格 + 右侧文本格」视为一条变更记录。

        支持双列变更表：一行内可能并排两对（如 A/B 一对 + D/E 一对）。
        旧逻辑每行只取「首个」日期格，会把第二对（右列）整对漏掉——
        ZT 样本 2025-12-08「ZT-2工序工时36秒修改为100秒」因此丢失。
        现改为收集本行【所有】日期格，逐格取其右侧到下一个日期格之间的
        首个文本，一对不漏。

        扫描范围覆盖整张变更表（老封面变更表动辄 10+ 条，2026-10-24 实测
        DEMO 文件 13 条曾被截成 4 条）；遇到工序汇总表头（工序号/Working
        procedure）即停，不会把工序行误认成变更记录。解析结果不再截断，
        封面容量由 _fill_cover 动态扩容兜底。
        """
        from datetime import datetime, date
        revs = []
        skip_kw = ("内容变更", "日期Date", "发放日期", "Issued")
        date_pat = re.compile(r"^\d{4}[-/年.]\d{1,2}[-/月.]\d{1,2}")
        for row in ws.iter_rows(min_row=4, max_row=min(ws.max_row, 80), max_col=10):
            first = row[0].value if row else None
            if isinstance(first, str) and ("工序号" in first
                                           or "Working procedure" in first):
                break  # 到工序汇总表了，变更表结束
            cells = list(row)
            # 收集本行所有日期格（0-indexed 位置）
            date_pos = [i for i, c in enumerate(cells)
                        if isinstance(c.value, (datetime, date))
                        or (isinstance(c.value, str)
                            and date_pat.match(c.value.strip()))]
            if not date_pos:
                continue
            for di, i in enumerate(date_pos):
                nxt = date_pos[di + 1] if di + 1 < len(date_pos) else len(cells)
                content = ""
                for c in cells[i + 1:nxt]:
                    v = c.value
                    if (v and not isinstance(v, (datetime, date))
                            and not any(k in str(v) for k in skip_kw)):
                        content = str(v).strip()
                        break
                if content:
                    d = cells[i].value
                    ds = d.strftime("%Y-%m-%d") if hasattr(d, "strftime") else str(d).strip()
                    revs.append(RevisionRecord(date=ds, content=content[:200], author=""))
        return revs

    def _get_right_cell_value(self, ws, row: int, col: int,
                              allow_date: bool = False) -> str:
        """获取右侧 1-5 格内第一个有意义的非空值（跳过日期/数字）

        allow_date=True 时日期格也当值返回（格式化成 YYYY/M/D）——
        发放日期在真实文件里常以日期类型存储，旧逻辑直接跳过导致
        该字段整体丢失（用户 2026-10-24 实测）。
        """
        from datetime import datetime, date as pydate
        for c in range(col + 1, min(col + 6, ws.max_column + 1)):
            cell = ws.cell(row=row, column=c)
            if cell.value is None:
                continue
            # 跳过纯日期值（转为字符串会污染目标字段）；allow_date 时放行
            if isinstance(cell.value, (datetime, pydate)):
                if allow_date:
                    d = cell.value
                    return f"{d.year}/{d.month}/{d.day}"
                continue
            # 跳过纯数字（工时等应通过专门字段提取）
            if isinstance(cell.value, (int, float)):
                # 但允许小数形式的版本号等
                s = str(cell.value).strip()
                # 如果是长数字串（如编号），当作文本返回
                if len(s) > 3 or "." in s:
                    continue
            text = str(cell.value).strip()
            if text and len(text) > 0:
                return text
        return ""

    def _parse_process_sheet(self, ws) -> Optional[Process]:
        """解析工序详情页（旧格式：所有数据混在同一 sheet 里，按关键字分段）"""
        from datetime import datetime, date as pydate
        p = Process()
        found_no = False

        # ---------- 1. 先检查 G4 直接工序号（部分源文件格式把工序号放 G4）----------
        g4_val = ws['G4'].value
        if g4_val and isinstance(g4_val, str):
            s = g4_val.strip()
            # 如果 G4 内容是编号类且不含标签关键词，直接作为工序号
            if s and not any(kw in s for kw in ("工序号", "Working procedure", "工位名称", "Testing part")):
                if re.match(r'^[A-Za-z0-9\-]', s):
                    p.process_no = s
                    found_no = True

        all_rows = []
        for row in ws.iter_rows(min_row=1, max_row=min(60, ws.max_row), max_col=15, values_only=True):
            all_rows.append(row)

        if not all_rows:
            return None

        # ---------- 2. 遍历前 8 行提取其他页头信息 ----------
        for r_idx, row in enumerate(all_rows[:8]):
            for col_idx, cell_val in enumerate(row):
                if cell_val is None or not isinstance(cell_val, str):
                    continue
                val = cell_val.strip()
                # 跳过纯日期和短字符串关键词格本身
                if len(val) < 3:
                    continue
                if re.match(r"^\d{4}[-/年.]\d{1,2}[-/月.]\d{1,2}$", val):
                    continue

                # 工序号（仅在 G4 没拿到时才从此路径解析；只认第一个）
                if not found_no and ("工序号" in val
                                     or ("Working procedure" in val and "NO" in val)):
                    v = self._get_right_cell_value(ws, r_idx + 1, col_idx + 1)
                    if v and not v.isdigit() and not self._looks_like_label(v):
                        p.process_no, found_no = v, True
                # 工位名称/工时与工序号谁先出现无关，必须各自独立匹配：
                # 旧版写成 if found_no 的 elif 链，工时标签排在工序号前面的文件
                # （实测 9STT RC-1：工时在第2行、工序号在第3行）会被整段跳过，
                # 规范化后新版头工时、工位名称双双丢值——2026-10-23 用户实测。
                # 右侧取值再加 _looks_like_label 防护：字段无值时顺延抓到的是
                # 下一个标签（如"文件编号File NO"），不能回填进新版头。
                if "工位名称" in val or "Testing part" in val:
                    v = self._get_right_cell_value(ws, r_idx + 1, col_idx + 1)
                    if v and not self._looks_like_label(v):
                        p.process_name_cn = v
                if "工时" in val and "hour" in val.lower():
                    wh = self._get_right_cell_value(ws, r_idx + 1, col_idx + 1)
                    if wh and not self._looks_like_label(wh):
                        # 版头工时可能是算式（实测 ZT-2 写「45*2+10=100S」）：
                        # 优先取等号右侧的结果值，否则清单里填 45、版头却显示
                        # 100S，自相矛盾（2026-10-26 ZT 实测）。无等号
                        # （如「20S」）才取首个数字。
                        m_eq = re.search(r"=\s*(\d+(?:\.\d+)?)", str(wh))
                        if m_eq:
                            p.work_hours = int(float(m_eq.group(1)))
                        else:
                            nums = re.findall(r"\d+", str(wh))
                            if nums:
                                p.work_hours = int(nums[0])
                        p.work_hours_display = wh

        # ---------- 3. 收集所有步骤文字 ----------
        STEP_SECTION_LABELS = (
            "附属文件", "Accessories", "劳保用具", "Protected tools",
            "设备/工夹具", "Machines", "物料", "Material", "工位质量检测",
            "Quality Insp", "编制", "审核", "批准", "日期", "完成图",
            "注：", "★★", "根据产品图纸要求", "可选项", "交替动作",
            "瞬间动作", "测试部分", "Testing part", "工序号", "Working procedure",
            "页码", "Page number", "工时", "Working hour",
            "产品规格", "Specification", "产品名称", "File NO", "文件编号",
            "Edition", "版本", "内容变更", "Change the conte",
            "压力测试", "项", "项目", "标准", "参数", "内容", "型号/编号",
            "Type/code", "编         号Code", "名称Name", "数量Quantity",
        )

        def _is_step_content(text, loose=False):
            """判断是否为步骤说明文本（不改动内容，只过滤明显非步骤的行）。

            loose=True 只用于无步号页（用户 2026-10-22 口径"按图＋完整描述
            为一步"的那种）：那里没有步号标签可混淆，图上方/下方的短句就是
            完整描述，不能再按"15 字以内且无标点"当图片说明丢掉——
            2026-10-22 验收抓到过 13 字无标点的描述被整句丢成空步骤。
            有步号的页维持原口径（loose=False），行为一字不变。"""
            if not text or len(text) < 4:
                return False
            t = text.strip()
            # 标签匹配：只匹配完整标签或标签开头（后跟：或空格），避免误杀正文中的标签词
            for bl in STEP_SECTION_LABELS:
                if t == bl or t.startswith(bl + "：") or t.startswith(bl + " ") or t.startswith(bl + "/"):
                    return False
            if re.match(r'^[A-Za-z0-9\-\.]+$', t):
                return False
            if re.match(r'^[A-Z]{1,3}\d{4,}$', t):
                return False
            if loose:
                return True
            # 跳过图片说明/标签（短文本、无标点、纯名词短语）
            if len(t) < 15 and '，' not in t and '。' not in t and '；' not in t and '、' not in t:
                return False
            return True

        # ---------- 按块位置提取每个步骤的文字 ----------
        _CN_NUM = {'一': 1, '二': 2, '三': 3, '四': 4, '五': 5,
                   '六': 6, '七': 7, '八': 8, '九': 9, '十': 10}

        def _parse_step_num(s):
            m = re.search(r"第([一二三四五六七八九十]+|\d+)步", s)
            if m:
                tok = m.group(1)
                return int(tok) if tok.isdigit() else _CN_NUM.get(tok)
            m = re.search(r"Step\s*(\d+)", s, re.IGNORECASE)
            if m:
                return int(m.group(1))
            return None

        # 收集步骤标记（0-indexed 行、列、步号）
        markers = []
        for r_idx, row in enumerate(all_rows):
            for c_idx, cell_val in enumerate(row):
                if cell_val is None or not isinstance(cell_val, str):
                    continue
                s = cell_val.strip()
                if not (re.match(r"^第[一二三四五六七八九十\d]+步", s)
                        or re.match(r"^Step\s*\d+", s, re.IGNORECASE)):
                    continue
                num = _parse_step_num(s)
                if num:
                    markers.append((r_idx, c_idx, num))

        # 去重：同位置保留一个，同步号只保留首次出现
        seen_pos, seen_num = set(), set()
        dedup = []
        for r, c, n in sorted(markers):
            if (r, c) in seen_pos or n in seen_num:
                continue
            seen_pos.add((r, c)); seen_num.add(n)
            dedup.append((r, c, n))
        markers = sorted(dedup)  # 按行、列排序

        virtual = False
        if not markers:
            virtual = True
            # 无步号页面（用户 2026-10-22 口径"按图＋完整描述为一步"）：
            # 拿配图锚点行当步骤领地边界——一步 = 一张（同行并排算一步）图 ＋ 它的完整描述。
            # 第 1 步从正文起点（第一张图往上吞描述，见 _first_step_row_by_image）开始；
            # 第 k 步（k≥2）从第 k-1 张图的下一行开始：每张图和它上方/下方的描述
            # 归同一图所在的步——图归描述在它上方的那一步，图下面的描述归下一步
            # （2026-10-22 验收抓到过 off-by-one：第 2 步从第 1 张图的行算起，
            #  图 1 被并进第 2 步，描述和图没配上对）。
            # 列统一取最左（0），让块文字覆盖整个内容区（右侧表格由 content_limit 挡着）。
            img_rows = self._image_anchor_rows(ws)
            if img_rows:
                first_top = self._first_step_row_by_image(ws) or img_rows[0]
                starts = [first_top] + [r + 1 for r in img_rows[:-1]]
                for i, top in enumerate(starts):
                    markers.append((top - 1, 0, i + 1))
                markers.sort()

        # 行/列边界：标记所在的行、列构成网格
        marker_rows = sorted(set(r for r, c, n in markers))
        marker_cols = sorted(set(c for r, c, n in markers))

        # 内容区右边界（右侧表格从此列开始）
        content_limit = ws.max_column
        for r_idx, row in enumerate(all_rows[:12]):
            for c_idx, cv in enumerate(row):
                # 只在第 5 列（E）及以右认右表标签：旧格式里右侧信息表
                # （附属文件/劳保用具/设备工夹具）最早也从 E 列起（真实样本
                # 如此），A-D 列是步骤描述的地盘。且标签格必须以关键词开头，
                # 不能按子串匹配——2026-10-22 无步号验收抓到过：A 列描述
                # "检查设备外观是否完好"含"设备"二字，被子串匹配误判成右表
                # 起点，内容区右界被压成 0 列，步骤文字和配图全归不了位。
                # 右界宁可探测不到（步骤文字多带几列右表内容），
                # 也绝不能把内容区压垮。
                if c_idx < 4:
                    continue
                if isinstance(cv, str) and any(
                        cv.strip().startswith(lb) for lb in (
                        "附属文件", "Accessories", "劳保用具", "Protected",
                        "设备", "物料", "Material", "工位质量", "Machines",
                        "Quality")):
                    content_limit = min(content_limit, c_idx)
        # 内容区下边界（页脚从此行开始）
        footer_row = ws.max_row
        for r_idx, row in enumerate(all_rows):
            for c_idx, cv in enumerate(row):
                if isinstance(cv, str) and any(
                        lb in cv for lb in ("编制", "Made by", "审核", "Checked")):
                    footer_row = min(footer_row, r_idx)

        # 每个标记对应一个块：行 [r, 下一标记行)、列 [c, 下一标记列)
        step_blocks = []  # (step_no, text, geom)
        for r, c, n in markers:
            r1 = next((mr for mr in marker_rows if mr > r), footer_row)
            c1 = next((mc for mc in marker_cols if mc > c), content_limit)
            texts = []
            for rr in range(r, r1):
                if rr >= len(all_rows):
                    break
                row = all_rows[rr]
                for cc in range(c, min(c1, len(row))):
                    cv = row[cc]
                    if cv is None or isinstance(cv, (datetime, pydate)):
                        continue
                    if not isinstance(cv, str):
                        continue
                    t = cv.strip()
                    if _is_step_content(t, loose=virtual):
                        texts.append(t)
            step_blocks.append((n, "；".join(texts), (r, r1, c, c1)))

        # 按步号升序生成步骤，记录每步在源文件中的块几何（供图片按位置归位）
        step_blocks.sort(key=lambda x: x[0])
        p._step_geom = {}
        for idx, (n, text, geom) in enumerate(step_blocks[:6]):
            p.add_step(idx + 1, text)
            p._step_geom[idx] = geom

        # ---------- 4. 解析附属文件（Accessories 下方单元格）----------
        accessories_parts = []
        for r_idx, row in enumerate(all_rows):
            for col_idx, cell_val in enumerate(row):
                if cell_val is None or not isinstance(cell_val, str):
                    continue
                s = str(cell_val).strip()
                if "Accessories" in s or "附属文件" in s:
                    # 下一行同列区域有值
                    for nr in [r_idx + 1, r_idx + 2]:
                        if nr < len(all_rows):
                            nc = col_idx + 1
                            if nc < len(all_rows[nr]) and all_rows[nr][nc]:
                                v = all_rows[nr][nc]
                                if v and isinstance(v, str) and len(str(v).strip()) > 0:
                                    accessories_parts.append(str(v).strip())
                            elif nc < len(all_rows[nr]) and all_rows[nr][col_idx]:
                                v = all_rows[nr][col_idx]
                                if v and isinstance(v, str) and len(str(v).strip()) > 0:
                                    accessories_parts.append(str(v).strip())
        p.accessories = "；".join(accessories_parts[:3])

        # ---------- 5. 解析物料/Materials 表格 ----------
        in_materials = False
        mat_header_found = False
        mat_row_count = 0
        for r_idx, row in enumerate(all_rows):
            row_str = " ".join(str(v or "") for v in row if v)
            if "物料" in row_str or "Material" in row_str:
                in_materials = True
                mat_header_found = True
                mat_row_count = 0
                continue
            if in_materials:
                # 遇到下一区就停止
                if ("质量" in row_str and "项目" in row_str) or "质检" in row_str \
                        or "设备" in row_str and "Machine" in row_str.lower():
                    in_materials = False
                    continue
                if mat_header_found and mat_row_count > 0:
                    # 尝试提取：编号、名称、数量
                    codes = [str(v or "") for v in row if v and isinstance(v, str)]
                    names = [str(v or "") for v in row if v and isinstance(v, str) and len(str(v)) > 1]
                    if codes and mat_row_count <= 8:  # 最多 8 条物料
                        code = codes[0]
                        if any(ch.isdigit() for ch in code) and len(code) > 3:
                            name = names[1] if len(names) > 1 else ""
                            p.add_material(code, name)
                    mat_row_count += 1
                elif not mat_header_found:
                    mat_row_count += 1

        # ---------- 6. 解析设备/Equipment 表格 ----------
        in_equipment = False
        eq_row_count = 0
        for r_idx, row in enumerate(all_rows):
            row_str = " ".join(str(v or "") for v in row if v)
            if "设备" in row_str and ("夹" in row_str or "Machine" in row_str or "工夹具" in row_str):
                in_equipment = True
                eq_row_count = 0
                continue
            if in_equipment:
                if ("物料" in row_str and "材料" not in row_str) or ("质量" in row_str and "项目" in row_str) \
                        or (eq_row_count > 5 and not any(str(v or "").strip() for v in row[:3])):
                    in_equipment = False
                    continue
                if eq_row_count > 0:
                    vals = [str(v or "").strip() for v in row if v is not None]
                    if len(vals) >= 2:
                        code = next((v for v in vals if any(c.isdigit() for c in v) and len(v) > 2), "")
                        name = next((v for v in vals if v != code and len(v) > 0), "")
                        if code:
                            p.add_equipment(code, name)
                eq_row_count += 1

        # ---------- 7. 解析质检/QC 表格 ----------
        in_qc = False
        qc_header_found = False
        qc_row_count = 0
        for r_idx, row in enumerate(all_rows):
            row_str = " ".join(str(v or "") for v in row if v)
            if "质量" in row_str and "项目" in row_str:
                in_qc = True
                qc_header_found = True
                qc_row_count = 0
                continue
            if in_qc:
                if "编制" in row_str or "批准" in row_str or "日期" in row_str:
                    in_qc = False
                    continue
                if qc_header_found and qc_row_count > 0:
                    vals = [str(v or "").strip() for v in row if v is not None]
                    if len(vals) >= 3 and qc_row_count <= 5:
                        item = vals[0] if vals[0] and len(vals[0]) > 1 else ""
                        spec = vals[1] if len(vals) > 1 and vals[1] else ""
                        param = vals[2] if len(vals) > 2 and vals[2] else ""
                        if item:
                            p.add_qc_item(item, spec, param)
                qc_row_count += 1

        # ---------- 8. 如果没找到工序号，用 sheet 名（补上 ZT- 前缀）----------
        if not p.process_no:
            if not found_no and not p.process_name_cn and not p.steps:
                return None
            t = ws.title[:31]
            p.process_no = t if t.startswith("ZT-") else f"ZT-{t}"
        return p


# ============================================================
# 便捷函数
# ============================================================

def generate_from_yaml(yaml_path: str, template_path: str, output_path: str) -> str:
    """从 YAML 配置直接生成 Excel"""
    wi = WorkInstruction.from_yaml(yaml_path)
    gen = ExcelGenerator(template_path)
    return gen.generate(wi, output_path)


def normalize_file(input_path: str, template_path: str, output_path: str) -> WorkInstruction:
    """规范化单个旧文件"""
    norm = Normalizer(template_path)
    return norm.normalize(input_path, output_path)


if __name__ == "__main__":
    import sys
    if len(sys.argv) > 1:
        cmd = sys.argv[1]
        if cmd == "gen":
            generate_from_yaml(sys.argv[2] if len(sys.argv) > 2 else "configs/DEMO_series/DEMO_demo.yaml",
                               sys.argv[3] if len(sys.argv) > 3 else "作业指导书版头.xlsx",
                               sys.argv[4] if len(sys.argv) > 4 else "output/generated.xlsx")
        elif cmd == "norm":
            normalize_file(sys.argv[2],
                           sys.argv[3] if len(sys.argv) > 3 else "作业指导书版头.xlsx",
                           sys.argv[4] if len(sys.argv) > 4 else "output/normalized.xlsx")
    else:
        print("用法:")
        print("  python excel_generator.py gen <yaml配置> [模板] [输出]")
        print("  python excel_generator.py norm <旧文件> [模板] [输出]")
