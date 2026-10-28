# -*- coding: utf-8 -*-
"""作业指导书工作台 - 图形界面
双击「启动工作台.bat」或运行 python wi_workbench.py
"""
from __future__ import annotations

import copy
import os
import subprocess
import sys
import threading
from pathlib import Path

if getattr(sys, "frozen", False):          # PyInstaller 打包后：exe 所在目录即工作目录
    exe_dir = Path(sys.executable).resolve().parent
    # 如果 exe 在"发布包"子目录中，说明是从发布包启动的，工作目录应上推到项目根
    if exe_dir.name == "发布包" and exe_dir.parent.exists():
        APP_DIR = exe_dir.parent
    else:
        APP_DIR = exe_dir
else:
    APP_DIR = Path(__file__).resolve().parent
    sys.path.insert(0, str(APP_DIR))
os.chdir(APP_DIR)

import tkinter as tk
from tkinter import ttk, filedialog, messagebox

try:
    from wi_models import (WorkInstruction, Header, Process,
                           ChangeRecordDetail, create_sb_template,
                           create_welding_template)
    from excel_generator import ExcelGenerator, normalize_file
    IMPORT_ERROR = None
except Exception as e:
    IMPORT_ERROR = e

TEMPLATE_DEFAULT = APP_DIR / "作业指导书版头.xlsx"
CONFIGS_DIR = APP_DIR / "configs"
OUTPUT_DIR = APP_DIR / "成品"
MEDIA_DIR = APP_DIR / "素材图片"

HEADER_FIELDS = [
    ("company_name", "公司名称"),
    ("company_subtitle", "工厂/子标题"),
    ("doc_title", "文档标题"),
    ("doc_title_en", "英文标题"),
    ("file_no", "文件编号 *"),
    ("edition", "版本"),
    ("product_spec", "产品规格(英文) *"),
    ("product_spec_cn", "产品规格(中文)"),
    ("made_by", "编制人"),
    ("checked_by", "审核人"),
    ("approved_by", "批准人"),
    ("issued_date", "发放日期"),
    ("note", "备注(关键工序说明)"),
]

PROCESS_FIELDS = [
    ("process_no", "工序号 *"),
    ("process_name", "工位名称(英文)"),
    ("process_name_cn", "工位名称(中文)"),
    ("work_hours", "工时(秒)"),
    ("page_no", "页码"),
    ("video_no", "视频编号"),
    ("accessories", "附属文件编号"),
    ("made_by", "编制人"),
    ("checked_by", "审核人"),
    ("approved_by", "批准人"),
    ("process_date", "日期"),
]

SUBTABLES = {
    "steps": {
        "label": "作业步骤 (最多6步)",
        "cols": [("step_no", "步号", 50), ("content", "步骤内容(中文)", 260),
                 ("content_en", "步骤内容(英文)", 220),
                 ("images", "配图(多个用;分隔)", 180)],
    },
    "materials": {
        "label": "物料 / 辅料",
        "cols": [("code", "编号", 140), ("name", "名称", 120), ("spec", "规格", 100),
                 ("qty", "数量", 60), ("unit", "单位", 60)],
    },
    "equipment": {
        "label": "设备 / 工夹具",
        "cols": [("code", "型号/编号", 120), ("name", "名称", 160),
                 ("operating_time", "运行参数", 180)],
    },
    "qc_items": {
        "label": "质检项目",
        "cols": [("item", "检测项目", 150), ("spec", "标准", 150),
                 ("parameter", "测量方法/参数", 150)],
    },
}

REVISION_TABLE = {
    "label": "变更记录 (版头简表)",
    "cols": [("date", "日期", 100), ("content", "变更内容", 400), ("author", "变更人", 100)],
}
CHANGE_RECORD_TABLE = {
    "label": "变更记录 (独立 Sheet4 明细)",
    "cols": [("modification", "修改内容", 300), ("checked_by", "审核", 90),
             ("approved_by", "批准", 90), ("date", "日期", 100)],
}


def blank_process(series: str = "") -> dict:
    p = Process(process_no=f"{series}-1" if series else "1",
                process_name="", process_name_cn="", work_hours=0,
                work_hours_display="0秒", page_no=1)
    return p.to_dict()


def imgs_to_str(v) -> str:
    """图片路径列表 → 子表格里的分号文本"""
    if isinstance(v, (list, tuple)):
        return ";".join(str(x) for x in v)
    return str(v or "")


def imgs_to_list(v) -> list:
    """分号/竖线/逗号分隔的文本 → 图片路径列表"""
    if isinstance(v, (list, tuple)):
        return [str(x).strip() for x in v if str(x).strip()]
    return [x.strip() for x in str(v or "").replace("|", ";").replace(",", ";").split(";") if x.strip()]


def reveal_in_folder(path):
    """在资源管理器中打开并选中文件。
    注意：文件对话框返回的是正斜杠路径，explorer 的 /select 只认反斜杠，必须先规范化。"""
    p = Path(str(path).replace("/", "\\")).resolve()
    try:
        if p.exists():
            subprocess.Popen(["explorer.exe", f"/select,{p}"])
        elif p.parent.exists():
            subprocess.Popen(["explorer.exe", str(p.parent)])
        else:
            os.startfile(str(APP_DIR))
    except Exception:
        try:
            os.startfile(str(p.parent if p.parent.exists() else APP_DIR))
        except Exception:
            pass


class RowEditDialog(tk.Toplevel):
    def __init__(self, parent, title, cols, initial=None):
        super().__init__(parent)
        self.title(title)
        self.resizable(False, False)
        self.result = None
        initial = initial or {}
        self.vars = {}
        frm = ttk.Frame(self, padding=12)
        frm.pack(fill="both", expand=True)
        for i, (key, label, _w) in enumerate(cols):
            ttk.Label(frm, text=label).grid(row=i, column=0, sticky="e", pady=3)
            var = tk.StringVar(value=str(initial.get(key, "")))
            self.vars[key] = var
            ttk.Entry(frm, textvariable=var, width=48).grid(
                row=i, column=1, sticky="we", padx=(8, 0), pady=3)
        btns = ttk.Frame(frm)
        btns.grid(row=len(cols), column=0, columnspan=2, pady=(14, 0))
        ttk.Button(btns, text="确定", command=self._ok).pack(side="left", padx=6)
        ttk.Button(btns, text="取消", command=self.destroy).pack(side="left", padx=6)
        self.bind("<Return>", lambda e: self._ok())
        self.bind("<Escape>", lambda e: self.destroy())
        self.transient(parent)
        self.grab_set()
        self.wait_visibility()
        self.focus_set()

    def _ok(self):
        self.result = {k: v.get() for k, v in self.vars.items()}
        self.destroy()


class SubTableFrame(ttk.LabelFrame):
    """通用子表格: 增/改/删/上下移动"""

    def __init__(self, master, spec, on_change=None):
        super().__init__(master, text=spec["label"], padding=6)
        self.spec = spec
        self.on_change = on_change
        self.rows: list[dict] = []

        body = ttk.Frame(self)
        body.pack(fill="both", expand=True)
        col_ids = [c[0] for c in spec["cols"]]
        self.tree = ttk.Treeview(body, columns=col_ids, show="headings", height=5)
        for key, label, width in spec["cols"]:
            self.tree.heading(key, text=label)
            self.tree.column(key, width=width, anchor="w")
        vsb = ttk.Scrollbar(body, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        self.tree.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")
        self.tree.bind("<Double-1>", lambda e: self.edit_selected())

        bar = ttk.Frame(self)
        bar.pack(fill="x", pady=(4, 0))
        for text, cmd in (("添加", self.add_row), ("修改", self.edit_selected),
                          ("删除", self.delete_selected), ("上移", lambda: self.move(-1)),
                          ("下移", lambda: self.move(1))):
            ttk.Button(bar, text=text, command=cmd, width=6).pack(side="left", padx=2)

        self._apply_view()

    def set_rows(self, rows):
        self.rows = copy.deepcopy(rows or [])
        self._apply_view()

    def get_rows(self):
        return copy.deepcopy(self.rows)

    def _apply_view(self):
        self.tree.delete(*self.tree.get_children())
        for r in self.rows:
            self.tree.insert("", "end", values=[r.get(k, "") for k, _l, _w in self.spec["cols"]])

    def _selected_idx(self):
        sel = self.tree.selection()
        if not sel:
            return None
        return self.tree.index(sel[0])

    def add_row(self):
        dlg = RowEditDialog(self.winfo_toplevel(), f"添加 - {self.spec['label']}", self.spec["cols"])
        dlg.wait_window()
        if dlg.result is not None:
            self.rows.append(dlg.result)
            self._apply_view()
            if self.on_change:
                self.on_change()

    def edit_selected(self):
        idx = self._selected_idx()
        if idx is None:
            messagebox.showinfo("提示", "请先选中一行")
            return
        dlg = RowEditDialog(self.winfo_toplevel(), f"修改 - {self.spec['label']}",
                            self.spec["cols"], self.rows[idx])
        dlg.wait_window()
        if dlg.result is not None:
            self.rows[idx] = dlg.result
            self._apply_view()
            self.tree.selection_set(self.tree.get_children()[idx])
            if self.on_change:
                self.on_change()

    def delete_selected(self):
        idx = self._selected_idx()
        if idx is None:
            return
        if messagebox.askyesno("确认", "删除选中行？"):
            del self.rows[idx]
            self._apply_view()
            if self.on_change:
                self.on_change()

    def move(self, delta):
        idx = self._selected_idx()
        if idx is None:
            return
        j = idx + delta
        if 0 <= j < len(self.rows):
            self.rows[idx], self.rows[j] = self.rows[j], self.rows[idx]
            self._apply_view()
            self.tree.selection_set(self.tree.get_children()[j])
            if self.on_change:
                self.on_change()


# 「工具箱 → 快速操作」的按钮表。提到模块级是为了无头可测：
# tests/verify.py 直接 import 本常量做断言（12 项、文字不重、方法名都真实存在），
# 不依赖 GUI 即可防止"加/改/删按钮时漏改或排布倒退"。
# 条目为（按钮文字, App 方法名）；方法名在构建时用 getattr(self, name) 解析。
QUICK_ACTIONS: list[tuple[str, str]] = [
    ("生成演示示例文件（DEMOB + 焊接）", "run_demo"),
    ("校验当前填写内容", "run_validate"),
    ("导入视频自动抽帧（生成带图步骤草稿）", "import_video"),
    ("自动创建工序（视频抽帧 +AI 起草）", "auto_create_process"),
    ("AI 模型配置（添加自定义大模型 API）", "ai_model_config"),
    ("一键接入探测器已录入的免费模型", "connect_detector_models"),
    ("AI 起草步骤文字（看当前工序配图写文字）", "run_ai_draft"),
    ("导入图纸参数（从工程图/BOM表提取 QC Items）", "import_blueprint"),
    ("导入图纸全参（物料+设备+QC 自动填入三表）", "import_blueprint_full"),
    ("清理素材（删除本文件用过的帧图和视频）", "cleanup_media"),
    ("清理全部测试文件（视频+抽帧+测试输出）", "cleanup_all_test_files"),
    ("同步进度到云端看板", "sync_progress"),
    ("打开输出目录 成品/", "open成品_dir"),
]


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("作业指导书工作台")
        self.geometry("1180x780")
        self.minsize(980, 640)
        try:
            ttk.Style(self).theme_use("vista")
        except Exception:
            pass

        self.doc: dict = self._blank_doc()
        self.current_yaml: Path | None = None
        self.dirty = False
        self._loading = False

        self._build_toolbar()
        self._build_tabs()
        self._build_status()
        self.protocol("WM_DELETE_WINDOW", self.on_close)
        self.load_doc_into_ui(self.doc)

    # ---------- 数据 ----------
    @staticmethod
    def _blank_doc() -> dict:
        return {
            "header": Header().to_dict(),
            "processes": [],
            "change_records": [],
        }

    def _collect_ui(self) -> dict:
        self._commit_current_process_to_list()
        header = {k: v.get() for k, v in self.header_vars.items()}
        header["revision_history"] = self.revision_tbl.get_rows()
        return {
            "header": header,
            "processes": copy.deepcopy(self.process_rows),
            "change_records": self.change_tbl.get_rows(),
        }

    def build_work_instruction(self) -> WorkInstruction:
        doc = self._collect_ui()
        for p in doc["processes"]:
            for s in p.get("steps", []) or []:
                s["images"] = imgs_to_list(s.get("images"))
        wi = WorkInstruction()
        wi.header = Header.from_dict(doc["header"])
        wi.processes = [Process.from_dict(p) for p in doc["processes"]]
        wi.change_records = [ChangeRecordDetail(**r) for r in doc["change_records"]]
        wi.recalc()
        return wi

    # ---------- UI 构建 ----------
    def _build_toolbar(self):
        bar = ttk.Frame(self, padding=(8, 6))
        bar.pack(fill="x")
        for text, cmd in (("新建", self.new_doc), ("打开YAML…", self.open_yaml),
                          ("保存", self.save_yaml), ("另存为…", self.save_yaml_as)):
            ttk.Button(bar, text=text, command=cmd).pack(side="left", padx=3)
        ttk.Separator(bar, orient="vertical").pack(side="left", fill="y", padx=8)
        ttk.Button(bar, text="生成 Excel", command=self.generate_excel).pack(side="left", padx=3)
        self.file_label = ttk.Label(bar, text="未打开配置文件（新文档）", foreground="#555")
        self.file_label.pack(side="right")

    def _build_status(self):
        self.status = tk.StringVar(value="就绪")
        ttk.Label(self, textvariable=self.status, padding=(10, 4), anchor="w").pack(fill="x", side="bottom")

    def _build_tabs(self):
        self.nb = ttk.Notebook(self)
        self.nb.pack(fill="both", expand=True, padx=8, pady=4)
        self._build_header_tab()
        self._build_process_tab()
        self._build_tools_tab()

    # --- 版头 tab ---
    def _build_header_tab(self):
        tab = ttk.Frame(self.nb, padding=10)
        self.nb.add(tab, text="  1. 版头信息  ")
        left = ttk.LabelFrame(tab, text="基本信息（*为必填）", padding=8)
        left.pack(side="left", fill="both", expand=True)
        self.header_vars: dict[str, tk.StringVar] = {}
        for i, (key, label) in enumerate(HEADER_FIELDS):
            ttk.Label(left, text=label + "：").grid(row=i, column=0, sticky="e", pady=3, padx=(0, 6))
            var = tk.StringVar()
            self.header_vars[key] = var
            e = ttk.Entry(left, textvariable=var, width=46)
            e.grid(row=i, column=1, sticky="we", pady=3)
            var.trace_add("write", self._mark_dirty)
        left.columnconfigure(1, weight=1)

        wrap = ttk.Frame(tab)
        wrap.pack(side="left", fill="both", expand=True, padx=(10, 0))
        self.revision_tbl = SubTableFrame(wrap, REVISION_TABLE, on_change=self._mark_dirty)
        self.revision_tbl.pack(fill="x")
        self.change_tbl = SubTableFrame(wrap, CHANGE_RECORD_TABLE, on_change=self._mark_dirty)
        self.change_tbl.pack(fill="both", expand=True, pady=(10, 0))

    # --- 工序 tab ---
    def _build_process_tab(self):
        tab = ttk.Frame(self.nb, padding=10)
        self.nb.add(tab, text="  2. 工序内容  ")

        nav = ttk.Frame(tab)
        nav.pack(side="left", fill="y")
        self.proc_list = tk.Listbox(nav, width=26, exportselection=False,
                                    font=("Microsoft YaHei UI", 10))
        self.proc_list.pack(fill="y", expand=True)
        self.proc_list.bind("<<ListboxSelect>>", self.on_process_select)
        bar = ttk.Frame(nav)
        bar.pack(fill="x", pady=6)
        for text, cmd in (("新增", self.add_process), ("删除", self.delete_process),
                          ("上移", lambda: self.move_process(-1)),
                          ("下移", lambda: self.move_process(1))):
            ttk.Button(bar, text=text, command=cmd, width=5).pack(side="left", padx=1)

        right = ttk.Frame(tab)
        right.pack(side="left", fill="both", expand=True, padx=(12, 0))
        self.process_vars: dict[str, tk.StringVar] = {}
        form = ttk.LabelFrame(right, text="当前工序", padding=8)
        form.pack(fill="x")
        for i, (key, label) in enumerate(PROCESS_FIELDS):
            r, c = divmod(i, 3)
            cell = ttk.Frame(form)
            cell.grid(row=r, column=c, sticky="we", padx=4, pady=3)
            ttk.Label(cell, text=label + "：").pack(side="left")
            var = tk.StringVar()
            self.process_vars[key] = var
            ttk.Entry(cell, textvariable=var, width=16).pack(side="left", fill="x", expand=True)
        form.columnconfigure((0, 1, 2), weight=1)

        tables = ttk.Frame(right)
        tables.pack(fill="both", expand=True, pady=(8, 0))
        self.sub_tables: dict[str, SubTableFrame] = {}
        col = 0
        for key, spec in SUBTABLES.items():
            st = SubTableFrame(tables, spec, on_change=self._mark_dirty)
            st.grid(row=col // 2, column=col % 2, sticky="nsew", padx=4, pady=4)
            self.sub_tables[key] = st
            col += 1
        tables.columnconfigure((0, 1), weight=1)
        tables.rowconfigure((0, 1), weight=1)

        self.process_rows: list[dict] = []
        self.proc_index = -1

    def _build_tools_tab(self):
        tab = ttk.Frame(self.nb)
        self.nb.add(tab, text="  3. 工具箱  ")

        # 内容超出窗口时纵向可滚：小屏幕/大字体下按钮不再被窗口边缘裁掉
        # （2026-10-24 用户实测：快速操作 12 个按钮 + 上面两个框总高约 730px，
        #  超出 1180x780 窗口可用高度，"同步进度到云端看板"等底部按钮不可见）
        qcanvas = tk.Canvas(tab, borderwidth=0, highlightthickness=0)
        qsb = ttk.Scrollbar(tab, orient="vertical", command=qcanvas.yview)
        qcanvas.configure(yscrollcommand=qsb.set)
        qsb.pack(side="right", fill="y")
        qcanvas.pack(side="left", fill="both", expand=True)
        inner = ttk.Frame(qcanvas, padding=16)
        qwin = qcanvas.create_window((0, 0), window=inner, anchor="nw")
        self.tools_canvas = qcanvas  # 防 GC，并便于后续扩展

        def _qsync(_e=None):
            qcanvas.configure(scrollregion=qcanvas.bbox("all"))

        inner.bind("<Configure>", _qsync)
        qcanvas.bind("<Configure>", lambda e: qcanvas.itemconfigure(qwin, width=e.width))
        self.bind_all("<MouseWheel>",
                      lambda e: qcanvas.yview_scroll(int(-e.delta / 120), "units"))

        g1 = ttk.LabelFrame(inner, text="规范化同事发来的旧文件", padding=10)
        g1.pack(fill="x", pady=(0, 12))
        ttk.Label(g1, text="旧文件：").grid(row=0, column=0, sticky="e")
        self.norm_in = tk.StringVar()
        ttk.Entry(g1, textvariable=self.norm_in, width=60).grid(row=0, column=1, sticky="we", padx=4)
        ttk.Button(g1, text="浏览…", command=self.pick_norm_in).grid(row=0, column=2)
        ttk.Label(g1, text="同时导出YAML(可选)：").grid(row=1, column=0, sticky="e", pady=(6, 0))
        self.norm_yaml = tk.StringVar()
        ttk.Entry(g1, textvariable=self.norm_yaml, width=60).grid(row=1, column=1, sticky="we", padx=4, pady=(6, 0))
        ttk.Button(g1, text="浏览…", command=self.pick_norm_yaml).grid(row=1, column=2, pady=(6, 0))
        ttk.Button(g1, text="开始规范化（生成到 成品/ 并载入编辑器）",
                   command=self.run_normalize).grid(row=2, column=1, sticky="w", pady=(8, 0))
        g1.columnconfigure(1, weight=1)

        g1b = ttk.LabelFrame(inner, text="一句话需求（大白话批量改，改完自动登记变更）", padding=10)
        g1b.pack(fill="x", pady=(0, 12))
        ttk.Label(g1b, text="要改的文件：").grid(row=0, column=0, sticky="e")
        self.cr_in = tk.StringVar()
        ttk.Entry(g1b, textvariable=self.cr_in, width=60).grid(row=0, column=1, sticky="we", padx=4)
        ttk.Button(g1b, text="浏览…", command=self.pick_cr_in).grid(row=0, column=2)
        ttk.Label(g1b, text="一句话：").grid(row=1, column=0, sticky="e", pady=(6, 0))
        self.cr_say = tk.StringVar()
        ttk.Entry(g1b, textvariable=self.cr_say, width=60).grid(row=1, column=1, sticky="we", padx=4, pady=(6, 0))
        ttk.Button(g1b, text="先看会改哪里", command=self.run_cr_preview).grid(row=1, column=2, pady=(6, 0))
        self.cr_pending = []
        self.cr_img_label = ttk.Label(g1b, text="配图更换：未挂任何换图项")
        self.cr_img_label.grid(row=2, column=1, sticky="w", pady=(6, 0))
        ttk.Button(g1b, text="挂一张换图…", command=self.pick_cr_image).grid(row=2, column=2, pady=(6, 0))
        ttk.Button(g1b, text="执行（原文件不动，另存副本到 成品/）",
                   command=self.run_change_request).grid(row=3, column=1, sticky="w", pady=(8, 0))
        g1b.columnconfigure(1, weight=1)

        g2 = ttk.LabelFrame(inner, text="快速操作", padding=10)
        g2.pack(fill="x")
        # 两列排布（按 QUICK_ACTIONS 顺序行优先）：默认视图即可看全全部按钮，
        # 更小的屏幕再配合上面的滚动条
        self._qa_buttons = {}
        for i, (text, name) in enumerate(QUICK_ACTIONS):
            r, c = divmod(i, 2)
            b = ttk.Button(g2, text=text, command=getattr(self, name))
            b.grid(row=r, column=c, sticky="w", padx=(0, 8), pady=2)
            self._qa_buttons[name] = b
        g2.columnconfigure((0, 1), weight=1)

        # 探测器接入按钮的状态提示（无探测器时红字、按钮灰置）
        self._detector_hint = tk.Label(g2, text="", fg="red",
                                       wraplength=640, justify="left")
        self._apply_detector_button_state()

    # ---------- 文档载入 ----------
    def load_doc_into_ui(self, doc: dict):
        self.doc = doc
        hdr = dict(doc.get("header") or {})
        for key, var in self.header_vars.items():
            var.set(str(hdr.get(key, "")))
        self.revision_tbl.set_rows(hdr.get("revision_history", []))
        self.change_tbl.set_rows(doc.get("change_records", []))
        self.process_rows = copy.deepcopy(doc.get("processes", []))
        for p in self.process_rows:
            for s in p.get("steps", []) or []:
                s["images"] = imgs_to_str(s.get("images"))
        self._refresh_proc_list()
        if self.process_rows:
            self.proc_list.selection_set(0)
            self.show_process(0)
        else:
            self.proc_index = -1
            self._clear_process_form()
        self.dirty = False
        self._update_file_label()

    def _update_file_label(self):
        name = str(self.current_yaml) if self.current_yaml else "未打开配置文件（新文档）"
        mark = "  ●未保存" if self.dirty else ""
        self.file_label.configure(text=name + mark)

    def _mark_dirty(self, *a):
        if not self._loading:
            self.dirty = True
            self._update_file_label()

    # ---------- 工序列表管理 ----------
    def _refresh_proc_list(self):
        self.proc_list.delete(0, "end")
        for i, p in enumerate(self.process_rows):
            label = p.get("process_no") or f"(第{i+1}项)"
            name = p.get("process_name_cn") or p.get("process_name") or ""
            self.proc_list.insert("end", f"{i+1}. {label}  {name[:12]}")

    def _commit_current_process_to_list(self):
        if self.proc_index < 0 or self.proc_index >= len(self.process_rows):
            return
        p = self.process_rows[self.proc_index]
        for key, var in self.process_vars.items():
            p[key] = var.get()
        try:
            p["work_hours"] = int(float(p.get("work_hours") or 0))
        except ValueError:
            pass
        p["work_hours_display"] = f"{p['work_hours']}秒"
        try:
            p["page_no"] = int(float(p.get("page_no") or 1))
        except ValueError:
            p["page_no"] = 1
        for key, st in self.sub_tables.items():
            p[key] = st.get_rows()
        if p.get("steps") and len(p["steps"]) > 6:
            p["steps"] = p["steps"][:6]

    def on_process_select(self, _e=None):
        sel = self.proc_list.curselection()
        if not sel:
            return
        idx = sel[0]
        if idx == self.proc_index:
            return
        self._commit_current_process_to_list()
        self.show_process(idx)

    def show_process(self, idx):
        if not (0 <= idx < len(self.process_rows)):
            return
        self._loading = True
        self.proc_index = idx
        p = self.process_rows[idx]
        for key, var in self.process_vars.items():
            var.set(str(p.get(key, "")))
        for key, st in self.sub_tables.items():
            st.set_rows(p.get(key, []))
        self._loading = False

    def _clear_process_form(self):
        self._loading = True
        for var in self.process_vars.values():
            var.set("")
        for st in self.sub_tables.values():
            st.set_rows([])
        self._loading = False

    def add_process(self):
        self._commit_current_process_to_list()
        series = self.header_vars["product_spec_cn"].get().split("系列")[0] if \
            self.header_vars["product_spec_cn"].get() else ""
        num = str(self.process_vars["process_no"].get()).rsplit("-", 1)[-1]
        p = blank_process(series)
        if num.isdigit():
            p["process_no"] = f"{series}-{int(num)+1}" if series else str(int(num) + 1)
        p["page_no"] = len(self.process_rows) + 1
        self.process_rows.append(p)
        self._refresh_proc_list()
        self.proc_list.selection_set(len(self.process_rows) - 1)
        self.show_process(len(self.process_rows) - 1)
        self.dirty = True
        self._update_file_label()

    def delete_process(self):
        sel = self.proc_list.curselection()
        if not sel:
            messagebox.showinfo("提示", "请先在左侧选择要删除的工序")
            return
        idx = sel[0]
        name = self.process_rows[idx].get("process_no", f"第{idx+1}项")
        if not messagebox.askyesno("确认", f"删除工序「{name}」？"):
            return
        self.proc_index = -1
        del self.process_rows[idx]
        self._refresh_proc_list()
        if self.process_rows:
            new = min(idx, len(self.process_rows) - 1)
            self.proc_list.selection_set(new)
            self.show_process(new)
        else:
            self._clear_process_form()
        self.dirty = True
        self._update_file_label()

    def move_process(self, delta):
        sel = self.proc_list.curselection()
        if not sel:
            return
        idx = sel[0]
        j = idx + delta
        if 0 <= j < len(self.process_rows):
            self._commit_current_process_to_list()
            self.process_rows[idx], self.process_rows[j] = self.process_rows[j], self.process_rows[idx]
            self.proc_index = j
            self._refresh_proc_list()
            self.proc_list.selection_set(j)
            self.show_process(j)
            self.dirty = True
            self._update_file_label()

    # ---------- 文件操作 ----------
    def _confirm_discard(self) -> bool:
        if not self.dirty:
            return True
        r = messagebox.askyesnocancel("未保存的修改", "当前有未保存的修改，先保存吗？\n（是=保存 · 否=放弃 · 取消=不操作）")
        if r is None:
            return False
        if r:
            return self.save_yaml()
        return True

    def new_doc(self):
        if not self._confirm_discard():
            return
        self.current_yaml = None
        self.load_doc_into_ui(self._blank_doc())
        self.status.set("已新建空白文档")

    def open_yaml(self):
        if not self._confirm_discard():
            return
        path = filedialog.askopenfilename(title="选择 YAML 配置",
                                          initialdir=str(CONFIGS_DIR if CONFIGS_DIR.exists() else APP_DIR),
                                          filetypes=[("YAML 配置", "*.yaml *.yml")])
        if not path:
            return
        try:
            wi = WorkInstruction.from_yaml(path)
        except Exception as e:
            messagebox.showerror("打开失败", f"文件无法读取：\n{e}")
            return
        self.current_yaml = Path(path)
        self.load_doc_into_ui({
            "header": wi.header.to_dict(),
            "processes": [p.to_dict() for p in wi.processes],
            "change_records": [vars(r) for r in wi.change_records],
        })
        self.status.set(f"已打开: {path}")

    def save_yaml(self, save_as=False):
        if self.current_yaml is None or save_as:
            CONFIGS_DIR.mkdir(parents=True, exist_ok=True)
            init = str(self.current_yaml or (CONFIGS_DIR / "untitled.yaml"))
            path = filedialog.asksaveasfilename(title="保存 YAML 配置", initialfile=Path(init).name,
                                                defaultextension=".yaml",
                                                initialdir=str(Path(init).parent),
                                                filetypes=[("YAML 配置", "*.yaml")])
            if not path:
                return False
            self.current_yaml = Path(path)
        try:
            wi = self.build_work_instruction()
            wi.to_yaml(str(self.current_yaml))
        except Exception as e:
            messagebox.showerror("保存失败", str(e))
            return False
        self.dirty = False
        self._update_file_label()
        self.status.set(f"已保存: {self.current_yaml}")
        return True

    def save_yaml_as(self):
        return self.save_yaml(save_as=True)

    def generate_excel(self):
        try:
            wi = self.build_work_instruction()
        except Exception as e:
            messagebox.showerror("数据错误", str(e))
            return
        errors = wi.header.validate()
        if not wi.processes:
            errors.append("至少需要一个工序（见“工序内容”页）")
        if errors:
            messagebox.showwarning("请先补全", "以下必填项未完成：\n- " + "\n- ".join(errors))
            self.nb.select(0)
            return
        if not TEMPLATE_DEFAULT.exists():
            messagebox.showerror("缺少模板", f"找不到版头模板：\n{TEMPLATE_DEFAULT}")
            return
        OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
        default_name = (wi.header.file_no or "作业指导书") + "_生成版.xlsx"
        path = filedialog.asksaveasfilename(title="生成 Excel", initialfile=default_name,
                                            defaultextension=".xlsx", initialdir=str(OUTPUT_DIR),
                                            filetypes=[("Excel 文件", "*.xlsx")])
        if not path:
            return
        try:
            gen = ExcelGenerator(str(TEMPLATE_DEFAULT))
            gen.generate(wi, path)
        except Exception as e:
            messagebox.showerror("生成失败", f"{e}")
            self.status.set(f"生成失败: {e}")
            return
        if messagebox.askyesno("生成成功", f"已生成：\n{path}\n\n打开所在文件夹吗？"):
            reveal_in_folder(path)
        self.status.set(f"已生成: {path}")

    # ---------- 工具箱 ----------
    def pick_norm_in(self):
        path = filedialog.askopenfilename(title="选择旧文件", filetypes=[("Excel 文件", "*.xlsx *.xlsm")])
        if path:
            self.norm_in.set(path)

    def pick_norm_yaml(self):
        CONFIGS_DIR.mkdir(parents=True, exist_ok=True)
        path = filedialog.asksaveasfilename(title="YAML 导出路径", defaultextension=".yaml",
                                            initialdir=str(CONFIGS_DIR))
        if path:
            self.norm_yaml.set(path)

    def run_normalize(self):
        src = self.norm_in.get().strip()
        if not src or not Path(src).exists():
            messagebox.showinfo("提示", "请先选择要规范化的旧文件")
            return
        if not TEMPLATE_DEFAULT.exists():
            messagebox.showerror("缺少模板", f"找不到版头模板：\n{TEMPLATE_DEFAULT}")
            return
        OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
        out = str(OUTPUT_DIR / f"{Path(src).stem}_规范化版.xlsx")
        try:
            wi = normalize_file(src, str(TEMPLATE_DEFAULT), out)
        except Exception as e:
            messagebox.showerror("规范化失败", str(e))
            return
        yaml_path = self.norm_yaml.get().strip()
        if yaml_path:
            try:
                wi.to_yaml(yaml_path)
            except Exception as e:
                messagebox.showwarning("YAML 导出失败", str(e))
        self.status.set(f"规范化完成: {out}")
        if messagebox.askyesno("规范化完成",
                               f"已生成：\n{out}\n\n要把提取出来的内容载入编辑器检查/修改吗？\n（会替换当前未保存的内容）"):
            self.current_yaml = Path(yaml_path) if yaml_path else None
            self.load_doc_into_ui({
                "header": wi.header.to_dict(),
                "processes": [p.to_dict() for p in wi.processes],
                "change_records": [vars(r) for r in wi.change_records],
            })
            self.dirty = bool(yaml_path is None)
            self._update_file_label()

    # ---------- 一句话需求 ----------
    def pick_cr_in(self):
        OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
        path = filedialog.askopenfilename(title="选要修改的指导书文件",
                                         initialdir=str(OUTPUT_DIR),
                                         filetypes=[("Excel", "*.xlsx")])
        if path:
            self.cr_in.set(path)

    def _cr_src(self):
        src = self.cr_in.get().strip().strip('"')
        if not src or not Path(src).exists():
            messagebox.showinfo("提示", "请先选择要修改的文件（成品/ 里的，或规范化出来的那份）")
            return None
        return src

    def _cr_pending_text(self):
        if not self.cr_pending:
            return "配图更换：未挂任何换图项"
        return "配图更换：已挂 {} 项（执行时一起生效）\n    ".format(len(self.cr_pending)) + \
               "\n    ".join(f"· {s} 图#{i + 1} → {Path(p).name}" for s, i, p in self.cr_pending)

    def _cr_refresh_pending(self):
        if hasattr(self, "cr_img_label"):
            self.cr_img_label.config(text=self._cr_pending_text())

    def run_cr_preview(self):
        src = self._cr_src()
        if not src:
            return
        say = self.cr_say.get().strip()
        try:
            import change_request as cr
            rep = cr.plan(src, say)
        except Exception as e:
            messagebox.showerror("预览失败", str(e))
            return
        text = rep.preview_text()
        if self.cr_pending:
            text += "\n\n另外挂着这些换图：\n    " + "\n    ".join(
                f"· {s} 图#{i + 1} → {Path(p).name}" for s, i, p in self.cr_pending)
        messagebox.showinfo("先看会改哪里", text)
        self.status.set("一句话需求：已预览")

    def run_change_request(self):
        src = self._cr_src()
        if not src:
            return
        say = self.cr_say.get().strip()
        if not say and not self.cr_pending:
            messagebox.showinfo("提示", "一句话和换图至少得有一项")
            return
        OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
        try:
            import change_request as cr
            rep = cr.run(src, say, out_dir=str(OUTPUT_DIR),
                         image_ops=list(self.cr_pending))
        except Exception as e:
            messagebox.showerror("执行失败", f"{e}\n\n（原文件没有被改动）")
            return
        if not rep.output:
            messagebox.showwarning("没有改动可保存",
                                   "这句话在这份文件里一处都没命中，换图也没挂上，所以没有写出新文件。\n"
                                   "请把要改的原话贴准一点（可以从单元格里原样复制一段）。")
            return
        self.cr_pending.clear()
        self._cr_refresh_pending()
        detail = "\n".join(f"· {cr._page_label(ed)} {ed.summary()}"
                           for ed in rep.edits if ed.hits)
        reg = "\n".join("· " + n for n in rep.notes)
        messagebox.showinfo(
            "改好了",
            f"原文件没动，新文件：\n{rep.output}\n\n【实际改动】\n{detail or '（只换了图）'}"
            f"\n\n【变更登记】\n{reg}")
        self.status.set(f"一句话需求完成: {Path(rep.output).name}")

    def pick_cr_image(self):
        src = self._cr_src()
        if not src:
            return
        import change_request as cr
        try:
            import openpyxl
            wb = openpyxl.load_workbook(src)
        except Exception as e:
            messagebox.showerror("读不了这个文件", str(e))
            return

        dlg = tk.Toplevel(self)
        dlg.title("挂一张换图：选页面 → 选图号 → 挑新图")
        dlg.geometry("640x420")
        bar = ttk.Frame(dlg, padding=8)
        bar.pack(fill="x")
        ttk.Label(bar, text="页面：").pack(side="left")
        pages = [ws.title for ws in wb.worksheets]
        cur = ttk.Combobox(bar, values=pages, width=32)
        cur.set(self.process_rows[self.proc_index].get("sheet") if
                0 <= self.proc_index < len(self.process_rows) and
                self.process_rows[self.proc_index].get("sheet") in pages else pages[0])
        cur.pack(side="left", padx=4)

        tree = ttk.Treeview(dlg, columns=("no", "pos", "size", "hint"),
                            show="headings", height=12)
        for cid, txt, w in (("no", "图号", 60), ("pos", "所在位置", 140),
                            ("size", "显示尺寸", 120), ("hint", "旁边最近的步骤", 220)):
            tree.heading(cid, text=txt)
            tree.column(cid, width=w, anchor="w")
        tree.pack(fill="both", expand=True, padx=8, pady=4)

        def load():
            tree.delete(*tree.get_children())
            ws = wb[cur.get()]
            for d in cr.list_images(ws):
                tree.insert("", "end", values=(f"#{d['index'] + 1}", d["pos"],
                                               f"{d['w']}×{d['h']}", d["hint"]))
        ttk.Button(bar, text="列出这一页的配图", command=load).pack(side="left", padx=4)
        load()

        def attach():
            sel = tree.selection()
            if not sel:
                messagebox.showwarning("提示", "先在表里选中一张图", parent=dlg)
                return
            no = int(tree.item(sel[0])["values"][0].lstrip("#")) - 1
            path = filedialog.askopenfilename(title="挑新图", parent=dlg,
                                              filetypes=[("图片", "*.png *.jpg *.jpeg *.bmp")])
            if not path:
                return
            self.cr_pending.append((cur.get(), no, path))
            self._cr_refresh_pending()
            messagebox.showinfo("已挂上", f"{cur.get()} 图#{no + 1} → {Path(path).name}\n"
                                         f"\n执行（改文字那步一起）时才真正写入新文件。", parent=dlg)
            dlg.destroy()

        btns = ttk.Frame(dlg, padding=8)
        btns.pack(fill="x")
        ttk.Button(btns, text="挑新图并挂上", command=attach).pack(side="left")
        ttk.Button(btns, text="清空已挂的换图",
                   command=lambda: (self.cr_pending.clear(), self._cr_refresh_pending())).pack(
            side="left", padx=8)
        ttk.Button(btns, text="关闭", command=dlg.destroy).pack(side="right")

    def run_demo(self):
        try:
            sb = create_sb_template()
            weld = create_welding_template()
            CONFIGS_DIR.joinpath("DEMO_series").mkdir(parents=True, exist_ok=True)
            CONFIGS_DIR.joinpath("Welding_DEMOA").mkdir(parents=True, exist_ok=True)
            OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
            sb.to_yaml(str(CONFIGS_DIR / "DEMO_series" / "DEMO_demo.yaml"))
            weld.to_yaml(str(CONFIGS_DIR / "Welding_DEMOA" / "DEMOA_demo.yaml"))
            msg = "已生成演示配置：\n- configs/DEMO_series/DEMO_demo.yaml\n- configs/Welding_DEMOA/DEMOA_demo.yaml"
            if TEMPLATE_DEFAULT.exists():
                gen = ExcelGenerator(str(TEMPLATE_DEFAULT))
                gen.generate(sb, str(OUTPUT_DIR / "DEMO_demo_生成版.xlsx"))
                gen.generate(weld, str(OUTPUT_DIR / "DEMOA_demo_生成版.xlsx"))
                msg += "\n\n已生成演示 Excel 到 成品/"
            messagebox.showinfo("完成", msg)
            self.status.set("演示文件已生成")
        except Exception as e:
            messagebox.showerror("失败", str(e))

    def run_validate(self):
        try:
            wi = self.build_work_instruction()
        except Exception as e:
            messagebox.showerror("数据错误", str(e))
            return
        errors = wi.header.validate()
        for i, p in enumerate(wi.processes, 1):
            if not p.process_no:
                errors.append(f"第{i}个工序缺少工序号")
        lines = [f"文档: {wi.header.file_no or '(未填文件编号)'} v{wi.header.edition}",
                 f"产品: {wi.header.product_spec}",
                 f"工序数: {len(wi.processes)}",
                 f"合计工时: {wi.total_work_hours} 秒"]
        for p in wi.processes:
            lines.append(f"  {p.process_no}: {p.process_name_cn or p.process_name} | "
                         f"{p.work_hours_display} | {len(p.steps)}步 | "
                         f"物料{len(p.materials)} 设备{len(p.equipment)} 质检{len(p.qc_items)}")
        if errors:
            lines.insert(0, "⚠ 问题：\n- " + "\n- ".join(errors) + "\n")
        messagebox.showinfo("校验结果", "\n".join(lines))

    def import_video(self):
        if getattr(self, "_vid_busy", False):
            messagebox.showinfo("提示", "上一个视频还在处理中，稍等几秒")
            return
        try:
            from video_frames import extract_frames
        except Exception as e:
            messagebox.showerror("缺少组件", f"视频功能不可用：{e}")
            return
        vid = filedialog.askopenfilename(
            title="选择操作视频",
            filetypes=[("视频文件", "*.mp4 *.avi *.mov *.wmv *.mkv *.mpeg *.flv"), ("所有文件", "*.*")])
        if not vid:
            return
        dlg = RowEditDialog(self, "抽帧设置（直接确定用默认值）",
                            [("interval", "兜底间隔秒数", 100),
                             ("max", "最多帧数", 100),
                             ("start", "起始帧（从第几张开始，默认1）", 100),
                             ("stride", "取帧间隔（每几张取1张，默认1）", 100),
                             ("group", "每几张图合并为一步（1-3，默认1）", 100)])
        dlg.wait_window()
        if dlg.result is None:
            return
        try:
            interval = float(dlg.result.get("interval") or 3)
            maxf = int(float(dlg.result.get("max") or 12))
        except ValueError:
            interval, maxf = 3.0, 12
        try:
            start = max(1, int(float(dlg.result.get("start") or 1)))
        except ValueError:
            start = 1
        try:
            stride = max(1, int(float(dlg.result.get("stride") or 1)))
        except ValueError:
            stride = 1
        try:
            group = min(3, max(1, int(float(dlg.result.get("group") or 1))))
        except ValueError:
            group = 1
        outdir = MEDIA_DIR / Path(vid).stem
        self._vid_busy = True
        self.status.set("正在抽帧，请稍候…")
        tip = "正在抽帧（画面切换处优先，否则按间隔均匀截）…"

        def work():
            try:
                frames, mode = extract_frames(vid, outdir, interval=interval,
                                              max_frames=max(1, min(maxf, 40)))
                try:
                    from video_frames import video_duration
                    dur = video_duration(vid)
                except Exception:
                    dur = 0
            except Exception as e:
                self.after(0, lambda: (self._vid_busy_done(),
                                       messagebox.showerror("抽帧失败", str(e))))
                return
            self.after(0, lambda: self._video_done(Path(vid).stem, frames, mode,
                                                   start, stride, group, dur))

        threading.Thread(target=work, daemon=True).start()

    def _vid_busy_done(self):
        self._vid_busy = False
        self.status.set("就绪")

    def _video_done(self, stem, frames, mode, start=1, stride=1, group=1, dur=0):
        self._vid_busy = False
        outdir = MEDIA_DIR / stem
        sel = frames[start - 1::stride] or frames
        chunks = [sel[i:i + group] for i in range(0, len(sel), group)] if group > 1 else [[f] for f in sel]
        self.status.set(f"抽帧完成（{mode}）: {len(frames)} 张，选用 {len(sel)} 张 → {len(chunks)} 步")
        dur_txt = f"，视频时长约 {int(round(dur))} 秒" if dur else ""
        load = messagebox.askyesno(
            "抽帧完成",
            f"共抽 {len(frames)} 张（{mode}）{dur_txt}。\n"
            f"按起点/间隔选用 {len(sel)} 张，每 {group} 张合并为一步 → 共 {len(chunks)} 步。\n"
            f"图片已存到：\n{outdir}\n\n"
            "是 = 载入「工序内容」当前工序的步骤表（覆盖原步骤，图片路径已填好，请补文字）\n"
            "否 = 只保存图片")
        if load:
            if self.proc_index < 0:
                messagebox.showwarning("提示", "请先在「工序内容」页选择或新增一个工序，再导入一次视频")
            else:
                self.sub_tables["steps"].set_rows([
                    {"step_no": i + 1, "content": "", "content_en": "",
                     "images": imgs_to_str(ch)}
                    for i, ch in enumerate(chunks)])
                if dur:
                    self.process_vars["work_hours"].set(str(int(round(dur))))
                self._mark_dirty()
                self.nb.select(1)
        reveal_in_folder(outdir)

    def auto_create_process(self):
        """自动创建工序：视频抽帧 + AI 起草步骤文字 + 填充到表单"""
        if getattr(self, "_auto_busy", False):
            messagebox.showinfo("提示", "上一个操作还在处理中，稍等几秒")
            return

        # 先选视频
        vid = filedialog.askopenfilename(
            title="选择操作视频",
            filetypes=[("视频文件", "*.mp4 *.avi *.mov *.wmv *.mkv *.mpeg *.flv"), ("所有文件", "*.*")])
        if not vid:
            return

        # 检查是否有当前工序
        if self.proc_index < 0:
            messagebox.showwarning("提示", "请先在「工序内容」页选择或新增一个工序，再自动创建")
            return

        p = self.process_rows[self.proc_index]
        pname = str(p.get("process_name_cn") or p.get("process_name") or "工序")

        # 抽帧设置
        dlg = RowEditDialog(self, "自动创建工序设置",
                            [("max", "最多帧数（每步 1-3 张）", 100),
                             ("group", "每几张图合并为一步（1-3）", 100)])
        dlg.wait_window()
        if dlg.result is None:
            return
        try:
            maxf = int(float(dlg.result.get("max") or 12))
        except ValueError:
            maxf = 12
        try:
            group = min(3, max(1, int(float(dlg.result.get("group") or 1))))
        except ValueError:
            group = 1

        outdir = MEDIA_DIR / f"auto_{Path(vid).stem}"
        self._auto_busy = True
        self.status.set("自动创建工序中：抽帧 + AI 起草…")

        def work():
            try:
                from video_frames import extract_frames, video_duration
                # 步骤 1: 抽帧
                self.after(0, lambda: self.status.set("正在抽帧…"))
                frames, mode = extract_frames(vid, outdir, max_frames=max(1, min(maxf, 40)))
                dur = video_duration(vid)

                # 步骤 2: 分组
                sel = frames
                chunks = [sel[i:i + group] for i in range(0, len(sel), group)] if group > 1 else [[f] for f in sel]

                if not chunks:
                    raise RuntimeError("抽帧后没有有效步骤")

                # 步骤 3: AI 起草每步文字
                self.after(0, lambda: self.status.set(f"AI 起草中（{len(chunks)} 步）…"))
                try:
                    import ai_draft
                    if getattr(self, "_ai_cands", None) is None:
                        self._ai_cands = ai_draft.vision_candidates(
                            lambda s: self.after(0, lambda t=s: self.status.set("AI：" + t)))
                    drafter = ai_draft.StepDrafter(self._ai_cands)
                except Exception as e:
                    self.after(0, lambda: (self._auto_busy_done(),
                                           messagebox.showerror("AI 不可用", f"无法调用 AI：{e}\n\n请先配置 API key 或检查网络")))
                    return

                steps = []
                done = 0
                for i, chunk in enumerate(chunks):
                    self.after(0, lambda n=i + 1: self.status.set(f"AI 正在写第 {n} 步…"))
                    try:
                        txt = drafter.draft(chunk, process_name=pname, step_no=i + 1)
                        if txt:
                            steps.append({
                                "step_no": i + 1,
                                "content": txt,
                                "content_en": "",
                                "images": imgs_to_str(chunk)
                            })
                            done += 1
                        else:
                            steps.append({
                                "step_no": i + 1,
                                "content": "",
                                "content_en": "",
                                "images": imgs_to_str(chunk)
                            })
                    except Exception as e:
                        steps.append({
                            "step_no": i + 1,
                            "content": f"[AI 失败：{str(e)[:30]}]",
                            "content_en": "",
                            "images": imgs_to_str(chunk)
                        })

                # 步骤 4: 填充到表单
                self.after(0, lambda: self._auto_done(outdir, done, len(chunks), dur, steps))

            except Exception as e:
                self.after(0, lambda: (self._auto_busy_done(),
                                       messagebox.showerror("自动创建失败", str(e))))

        threading.Thread(target=work, daemon=True).start()

    def _auto_busy_done(self):
        self._auto_busy = False
        self.status.set("就绪")

    def _auto_done(self, outdir, done, total, dur, steps):
        self._auto_busy = False
        self.status.set("就绪")

        # 填充到当前工序
        if self.proc_index >= 0 and self.proc_index < len(self.process_rows):
            self.sub_tables["steps"].set_rows(steps)
            if dur:
                self.process_vars["work_hours"].set(str(int(round(dur))))
                self.process_vars["work_hours_display"].set(f"{int(round(dur))}秒")
            self._mark_dirty()
            self.nb.select(1)  # 切到工序页

        tip = f"自动创建完成：{done}/{total} 步已写入 AI 文字\n"
        tip += f"图片已存到：\n{outdir}\n\n"
        tip += f"工时：{int(round(dur))}秒\n\n"
        tip += "请检查步骤文字，必要时手动修改。"

        messagebox.showinfo("自动创建完成", tip)
        reveal_in_folder(outdir)

    def ai_model_config(self):
        """打开 AI 模型配置窗口，让用户添加/编辑自定义大模型 API"""
        try:
            import json
        except ImportError:
            messagebox.showerror("错误", "JSON 模块不可用")
            return

        config_path = APP_DIR / "ai_models.json"

        # 读取现有配置
        if config_path.exists():
            try:
                with open(config_path, "r", encoding="utf-8") as f:
                    config = json.load(f)
            except Exception as e:
                config = {"models": []}
        else:
            config = {"models": []}

        models = config.get("models", [])

        # 创建配置窗口
        dlg = tk.Toplevel(self)
        dlg.title("AI 模型配置")
        dlg.geometry("600x500")
        dlg.resizable(False, False)

        frm = ttk.Frame(dlg, padding=12)
        frm.pack(fill="both", expand=True)

        ttk.Label(frm, text="AI 模型配置（用于离线 AI 起草）", font=("Microsoft YaHei UI", 10, "bold")).grid(
            row=0, column=0, columnspan=2, pady=(0, 10))

        ttk.Label(frm, text="说明：").grid(row=1, column=0, sticky="w", pady=(0, 5))
        ttk.Label(frm, text="在此添加自定义大模型 API，工作台的 AI 功能将优先使用这些配置。", wraplength=500).grid(
            row=2, column=0, columnspan=2, sticky="w", pady=(0, 10))

        # 模型列表
        list_frame = ttk.LabelFrame(frm, text="已配置的模型", padding=6)
        list_frame.grid(row=3, column=0, columnspan=2, sticky="nsew", pady=(0, 10))

        list_frame.columnconfigure(0, weight=1)
        list_frame.rowconfigure(1, weight=1)

        listbox = tk.Listbox(list_frame, height=8, font=("Microsoft YaHei UI", 9))
        listbox.pack(fill="both", expand=True)

        # 填充列表
        for i, m in enumerate(models):
            status = "✓" if m.get("api_key") and not m["api_key"].startswith("sk-请替换") else "✗"
            listbox.insert("end", f"{status} {m.get('name', '未命名')} - {m.get('model', '无模型')}")

        # 按钮
        btn_frame = ttk.Frame(list_frame)
        btn_frame.pack(fill="x", pady=(5, 0))

        def add_model():
            """添加新模型"""
            model_dlg = tk.Toplevel(dlg)
            model_dlg.title("添加模型")
            model_dlg.geometry("450x300")
            model_dlg.resizable(False, False)

            f = ttk.Frame(model_dlg, padding=12)
            f.pack(fill="both", expand=True)

            vars = {}
            labels = [
                ("name", "模型名称（如：智谱 GLM-4V）"),
                ("base_url", "API 地址（如：https://open.bigmodel.cn/api/paas/v4）"),
                ("model", "模型名称（如：glm-4v）"),
                ("api_key", "API Key（以 sk- 开头）"),
            ]
            for i, (key, label) in enumerate(labels):
                ttk.Label(f, text=label + "：").grid(row=i, column=0, sticky="e", pady=5)
                var = tk.StringVar()
                vars[key] = var
                entry = ttk.Entry(f, textvariable=var, width=40)
                entry.grid(row=i, column=1, sticky="we", padx=(8, 0), pady=5)

            def save():
                new_model = {
                    "name": vars["name"].get() or "自定义模型",
                    "base_url": vars["base_url"].get().strip(),
                    "model": vars["model"].get().strip(),
                    "api_key": vars["api_key"].get().strip(),
                    "type": "custom"
                }
                if not new_model["base_url"] or not new_model["model"]:
                    messagebox.showwarning("提示", "API 地址和模型名称不能为空")
                    return
                models.append(new_model)
                config["models"] = models
                with open(config_path, "w", encoding="utf-8") as f:
                    json.dump(config, f, ensure_ascii=False, indent=2)
                model_dlg.destroy()
                dlg.destroy()
                messagebox.showinfo("成功", "模型已添加")

            btns = ttk.Frame(f)
            btns.grid(row=len(labels), column=0, columnspan=2, pady=(10, 0))
            ttk.Button(btns, text="保存", command=save).pack(side="left", padx=5)
            ttk.Button(btns, text="取消", command=model_dlg.destroy).pack(side="left", padx=5)

        def edit_model(idx):
            """编辑模型"""
            if idx < 0 or idx >= len(models):
                return
            m = models[idx]

            model_dlg = tk.Toplevel(dlg)
            model_dlg.title("编辑模型")
            model_dlg.geometry("450x300")
            model_dlg.resizable(False, False)

            f = ttk.Frame(model_dlg, padding=12)
            f.pack(fill="both", expand=True)

            vars = {}
            labels = [
                ("name", "模型名称"),
                ("base_url", "API 地址"),
                ("model", "模型名称"),
                ("api_key", "API Key"),
            ]
            for i, (key, label) in enumerate(labels):
                ttk.Label(f, text=label + "：").grid(row=i, column=0, sticky="e", pady=5)
                var = tk.StringVar(value=m.get(key, ""))
                vars[key] = var
                entry = ttk.Entry(f, textvariable=var, width=40)
                entry.grid(row=i, column=1, sticky="we", padx=(8, 0), pady=5)

            def save():
                models[idx] = {
                    "name": vars["name"].get() or "自定义模型",
                    "base_url": vars["base_url"].get().strip(),
                    "model": vars["model"].get().strip(),
                    "api_key": vars["api_key"].get().strip(),
                    "type": m.get("type", "custom")
                }
                config["models"] = models
                with open(config_path, "w", encoding="utf-8") as f:
                    json.dump(config, f, ensure_ascii=False, indent=2)
                model_dlg.destroy()
                dlg.destroy()
                messagebox.showinfo("成功", "模型已更新")

            btns = ttk.Frame(f)
            btns.grid(row=len(labels), column=0, columnspan=2, pady=(10, 0))
            ttk.Button(btns, text="保存", command=save).pack(side="left", padx=5)
            ttk.Button(btns, text="取消", command=model_dlg.destroy).pack(side="left", padx=5)

        def delete_model(idx):
            """删除模型"""
            if idx < 0 or idx >= len(models):
                return
            if messagebox.askyesno("确认", f"确定删除模型「{models[idx].get('name', '?')}」？"):
                models.pop(idx)
                config["models"] = models
                with open(config_path, "w", encoding="utf-8") as f:
                    json.dump(config, f, ensure_ascii=False, indent=2)
                dlg.destroy()
                messagebox.showinfo("成功", "模型已删除")

        listbox.bind("<Double-1>", lambda e: edit_model(listbox.curselection()[0] if listbox.curselection() else -1))

        btns = ttk.Frame(frm)
        btns.grid(row=4, column=0, columnspan=2, pady=(0, 5))
        ttk.Button(btns, text="添加模型", command=add_model).pack(side="left", padx=3)
        ttk.Button(btns, text="编辑选中", command=lambda: edit_model(listbox.curselection()[0] if listbox.curselection() else -1)).pack(side="left", padx=3)
        ttk.Button(btns, text="删除选中", command=lambda: delete_model(listbox.curselection()[0] if listbox.curselection() else -1)).pack(side="left", padx=3)

        # 说明文字
        ttk.Label(frm, text="提示：", foreground="#666").grid(row=5, column=0, sticky="w", pady=(10, 0))
        ttk.Label(frm, text="1. 配置后需重启工作台才能生效\n2. 配置文件位置：ai_models.json（与 exe 同目录）\n3. 支持 OpenAI 兼容 API 格式", wraplength=500, justify="left").grid(
            row=6, column=0, columnspan=2, sticky="w", pady=(0, 5))

        dlg.transient(self)
        dlg.grab_set()
        dlg.wait_window()

    def open成品_dir(self):
        OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
        reveal_in_folder(OUTPUT_DIR)

    def connect_detector_models(self):
        """一键接入同目录「免费大模型自动探测器」已录入的免费模型。

        离线读取探测器的 keys.json + 免费模型清单，把候选直接写入 self._ai_cands，
        后续 AI 起草/图纸解析即用这些模型，不再走 60 秒实时探测。
        探测器缺失或没录入 key 时由 ai_draft 抛 DraftError，这里转成提示。
        """
        import ai_draft
        try:
            cands = ai_draft.detector_candidates(
                "vision", lambda s: self.status.set("探测器：" + s))
        except Exception as e:
            messagebox.showerror("接入失败", str(e))
            return
        self._ai_cands = cands
        self.status.set(
            f"已接入探测器 {len(cands)} 个免费模型，AI 功能将直接使用，无需重新探测。")

    def _apply_detector_button_state(self):
        """根据同目录是否有探测器，决定「一键接入」按钮是否可用，并给红字提示。"""
        import ai_draft
        d = ai_draft.find_detector_dir()
        btn = self._qa_buttons.get("connect_detector_models")
        if d is None:
            if btn:
                btn.configure(state="disabled")
            self._detector_hint.configure(
                text="⚠ 未检测到「免费大模型自动探测器」：此按钮已禁用。\n"
                     "请把探测器放在工作台同目录后重启工作台。",
                fg="red")
            self._detector_hint.pack(fill="x", pady=(4, 0))
        else:
            if btn:
                btn.configure(state="normal")
            self._detector_hint.configure(
                text=f"已检测到探测器：{d}", fg="#0a7d2b")
            self._detector_hint.pack(fill="x", pady=(4, 0))

    def run_ai_draft(self):
        if getattr(self, "_ai_busy", False):
            messagebox.showinfo("提示", "AI 还在写上一条，稍等几秒")
            return
        if self.proc_index < 0 or self.proc_index >= len(self.process_rows):
            messagebox.showwarning("提示", "请先在「工序内容」页选择或新增一个工序")
            return
        rows = self.sub_tables["steps"].get_rows()
        targets = [i for i, r in enumerate(rows) if imgs_to_list(r.get("images"))]
        if not targets:
            messagebox.showwarning(
                "提示", "当前工序的步骤表里没有配图。\n"
                "先用「导入视频自动抽帧」生成带图步骤，或在步骤的配图列填图片路径")
            return
        blank = [i for i in targets if not str(rows[i].get("content", "")).strip()]
        mode = messagebox.askyesno(
            "AI 起草步骤文字",
            f"当前工序共 {len(rows)} 步，其中 {len(targets)} 步带配图，{len(blank)} 步文字还是空白。\n\n"
            "是 = 覆盖重写所有带图步骤的文字\n"
            "否 = 只补写空白文字的步骤")
        todo = targets if mode else blank
        if not todo:
            messagebox.showinfo("提示", "没有需要补写的步骤（空白步骤都没有配图？）")
            return
        p = self.process_rows[self.proc_index]
        pname = str(p.get("process_name_cn") or p.get("process_name") or "")
        self._ai_busy = True
        self.nb.select(1)
        self.status.set("AI 起草中：先探测可用渠道（约1分钟，之后每步几秒钟）…")

        def work():
            try:
                import ai_draft
                if getattr(self, "_ai_cands", None) is None:
                    self._ai_cands = ai_draft.vision_candidates(
                        lambda s: self.after(0, lambda t=s: self.status.set("AI：" + t)))
                drafter = ai_draft.StepDrafter(self._ai_cands)
            except Exception as e:
                self.after(0, lambda: self._ai_done(False, f"AI 起草失败：{e}"))
                return
            done, fails, err = 0, [], None
            for i in todo:
                r = rows[i]
                try:
                    n = int(str(r.get("step_no") or i + 1))
                except ValueError:
                    n = i + 1
                self.after(0, lambda n=n: self.status.set(f"AI 正在写第 {n} 步…"))
                try:
                    txt = drafter.draft(imgs_to_list(r.get("images")),
                                        process_name=pname, step_no=n)
                except Exception as e:
                    err = str(e)[:100]
                    fails.append(f"第{n}步：{err}")
                    if "渠道" in err or "key" in err.lower():
                        break  # 渠道整体故障就不再耗时间
                else:
                    if txt:
                        done += 1
                        self.after(0, lambda idx=i, t=txt: self._ai_patch(idx, t))
                    else:
                        fails.append(f"第{n}步：模型返回空内容")
            tip = f"写好 {done} 步"
            if done == 0:
                self._ai_cands = None  # 全失败：下次重新探测渠道
            if fails:
                tip += f"，失败 {len(fails)} 步：\n" + "\n".join(fails[:6])
            self.after(0, lambda: self._ai_done(done > 0, tip))

        threading.Thread(target=work, daemon=True).start()

    def _ai_patch(self, idx, text):
        st = self.sub_tables["steps"]
        if 0 <= idx < len(st.rows):
            st.rows[idx]["content"] = text
            st._apply_view()
            self._mark_dirty()

    def _ai_done(self, ok, tip):
        self._ai_busy = False
        self.status.set("就绪")
        (messagebox.showinfo if ok else messagebox.showerror)("AI 起草", tip)

    def import_blueprint(self):
        """从工程图纸/BOM表/工艺卡图片中提取 QC Items。"""
        if getattr(self, "_ai_busy", False):
            messagebox.showinfo("提示", "AI 还在处理上一条，稍等几秒")
            return
        if self.proc_index < 0 or self.proc_index >= len(self.process_rows):
            messagebox.showwarning("提示", "请先在「工序内容」页选择或新增一个工序")
            return
        p = self.process_rows[self.proc_index]
        pname = str(p.get("process_name_cn") or p.get("process_name") or "")
        # 选图纸文件
        files = filedialog.askopenfilenames(
            title="选择工程图纸/BOM表/工艺卡图片（可多选）",
            filetypes=[("图片文件", "*.png *.jpg *.jpeg *.bmp *.gif *.tif *.tiff"), ("所有文件", "*.*")]
        )
        if not files:
            return
        self._ai_busy = True
        self.status.set(f"导入图纸参数中：已选 {len(files)} 张图纸，正在解析…")
        self.nb.select(2)  # 切到 QC Items 页签等待结果

        def work():
            try:
                import ai_draft
                cands = ai_draft.vision_candidates(
                    lambda s: self.after(0, lambda t=s: self.status.set(t)))
                parser = ai_draft.BlueprintParser(cands)
                self.after(0, lambda: self.status.set("正在识别图纸参数…"))
                items = parser.parse(files, process_name=pname)
                if not items:
                    self.after(0, lambda: self._ai_done(False, "未从图纸中识别到质量检验参数。\n请确保图片包含明确的检验要求标注。"))
                    return
                # 弹窗展示并确认
                summary_lines = []
                for idx, item in enumerate(items, 1):
                    spec = item.get("spec", "").strip()
                    param = item.get("parameter", "").strip()
                    summary_lines.append(f"{idx}. {item['item']}" + (f"  — 规格：{spec}" if spec else "") + (f"；方法：{param}" if param else ""))
                detail_text = "\n".join(summary_lines)
                result_ok = self.after(0, lambda: None)  # placeholder
                # 用同步对话框（线程安全：在主线程执行）
                self.after(0, lambda: self._confirm_blueprint(items, detail_text))
            except Exception as e:
                err_msg = str(e)[:300]
                self.after(0, lambda: self._ai_done(False, f"导入图纸失败：\n{err_msg}"))

        threading.Thread(target=work, daemon=True).start()

    def _confirm_blueprint(self, items, detail_text):
        """在 GUI 线程中弹出确认框。"""
        text = f"从图纸中识别到以下 QC Items（共 {len(items)} 项）：\n\n{detail_text}\n\n确认写入当前工序的 QC Items 表格？"
        ok = messagebox.askyesno("导入图纸参数 - 确认", text)
        if ok:
            qc_tbl = self.sub_tables["qc_items"]
            qc_tbl.set_rows(items)
            self._mark_dirty()
            self.status.set("已导入 QC Items")
            self._ai_done(True, f"成功导入 {len(items)} 项 QC Items")
        else:
            self.status.set("已取消导入")
            self._ai_done(True, "已取消")

    def import_blueprint_full(self):
        """全面解析图纸参数：物料+BOM+设备+QC，自动填入三表。"""
        if getattr(self, "_ai_busy", False):
            messagebox.showinfo("提示", "AI 还在处理上一条，稍等几秒")
            return
        if self.proc_index < 0 or self.proc_index >= len(self.process_rows):
            messagebox.showwarning("提示", "请先在「工序内容」页选择或新增一个工序")
            return
        p = self.process_rows[self.proc_index]
        pname = str(p.get("process_name_cn") or p.get("process_name") or "")
        files = filedialog.askopenfilenames(
            title="选择工程图纸/BOM表/工艺卡图片（可多选）",
            filetypes=[("图片文件", "*.png *.jpg *.jpeg *.bmp *.gif *.tif *.tiff"), ("所有文件", "*.*")]
        )
        if not files:
            return
        self._ai_busy = True
        self.status.set(f"全面解析图纸参数：已选 {len(files)} 张图纸，正在提取物料/设备/QC…")
        self.nb.select(2)

        def work():
            try:
                import ai_draft
                cands = ai_draft.vision_candidates(
                    lambda s: self.after(0, lambda t=s: self.status.set(t)))
                parser = ai_draft.FullParseParser(cands)
                self.after(0, lambda: self.status.set("正在识别图纸中的物料、设备和质检参数…"))
                result = parser.parse(files, process_name=pname)

                mats = result.get("materials", [])
                eqs = result.get("equipment", [])
                qcs = result.get("qc_items", [])

                summary_lines = []
                if mats:
                    summary_lines.append(f"📦 物料 ({len(mats)} 项):")
                    for r in mats[:10]:
                        summary_lines.append(f"  - {r['code']}: {r['name']} | {r['spec']} x{int(r['qty'])}{r['unit']}")
                if eqs:
                    summary_lines.append(f"🔧 设备/工夹具 ({len(eqs)} 项):")
                    for r in eqs[:10]:
                        summary_lines.append(f"  - {r['code']}: {r['name']} | {r['operating_time']}")
                if qcs:
                    summary_lines.append(f"✅ 质检项目 ({len(qcs)} 项):")
                    for r in qcs[:10]:
                        spec_part = f" 标准：{r['spec']}" if r['spec'] else ""
                        param_part = f"；方法：{r['parameter']}" if r['parameter'] else ""
                        summary_lines.append(f"  - {r['item']}{spec_part}{param_part}")

                summary = "\n".join(summary_lines) if summary_lines else "图中未识别到可用参数。"
                total = len(mats) + len(eqs) + len(qcs)

                self.after(0, lambda: self._confirm_blueprint_full(
                    result, f"从图纸中完整解析了以下参数（共 {total} 项）：\n\n{summary}\n\n是否写入对应表格？"))
            except Exception as e:
                err_msg = str(e)[:300]
                self.after(0, lambda: self._ai_done(False, f"全面解析失败：\n{err_msg}"))

        threading.Thread(target=work, daemon=True).start()

    def _confirm_blueprint_full(self, result, detail_text):
        """全面解析确认后批量填入三个子表。"""
        from wi_models import Material, Equipment, QCItem
        ok = messagebox.askyesno("导入图纸全参 - 确认", detail_text + "\n\n是 = 全部写入  否 = 暂不写入")
        if not ok:
            self.status.set("已取消导入")
            self._ai_done(True, "已取消")
            return

        applied = 0
        if result.get("materials"):
            for m in result["materials"]:
                mat = Material(**m)
                # 避免重复（按编号判断）
                existing_codes = {x.code for x in self.sub_tables["materials"].rows if x.get("code")}
                if m["code"] not in existing_codes:
                    self.sub_tables["materials"].rows.append(dict(m))
                    applied += 1
            self.sub_tables["materials"]._apply_view()

        if result.get("equipment"):
            for e in result["equipment"]:
                if e["code"] not in {x.get("code", "") for x in self.sub_tables["equipment"].rows}:
                    self.sub_tables["equipment"].rows.append(dict(e))
                    applied += 1
            self.sub_tables["equipment"]._apply_view()

        if result.get("qc_items"):
            for q in result["qc_items"]:
                if q["item"] not in {x.get("item", "") for x in self.sub_tables["qc_items"].rows}:
                    self.sub_tables["qc_items"].rows.append(dict(q))
                    applied += 1
            self.sub_tables["qc_items"]._apply_view()

        self._mark_dirty()
        self.status.set(f"已导入 {applied} 项图纸参数")
        self._ai_done(True, f"全面解析完成：物料{sum(1 for m in (result.get('materials',[])) if any(x.get('code')==m['code'] for x in self.sub_tables['materials'].rows))}项、设备{sum(1 for e in (result.get('equipment',[])) if any(x.get('code')==e['code'] for x in self.sub_tables['equipment'].rows))}项、QC{sum(1 for q in (result.get('qc_items',[])) if any(x.get('item')==q['item'] for x in self.sub_tables['qc_items'].rows))}项，共 {applied} 项已写入。")

    VIDEO_EXTS = (".mp4", ".avi", ".mov", ".wmv", ".mkv", ".mpeg", ".flv")

    def cleanup_media(self):
        """删除当前配置文件实际引用过的抽帧目录（及工作目录里的同名视频）。
        只认当前这份 YAML 里图片路径指向的 素材图片 下的一级目录，其余一律不动。"""
        doc = self._collect_ui()
        dirs = {}
        for p in doc["processes"]:
            for s in p.get("steps", []) or []:
                for img in imgs_to_list(s.get("images")):
                    fp = Path(img)
                    if not fp.is_absolute():
                        fp = APP_DIR / fp
                    try:
                        fp = fp.resolve()
                        rel = fp.relative_to(MEDIA_DIR.resolve())
                    except (ValueError, OSError):
                        continue
                    if len(rel.parts) >= 2:
                        d = MEDIA_DIR / rel.parts[0]
                        dirs[d] = dirs.get(d, 0) + 1
        if not dirs:
            messagebox.showinfo("清理素材",
                                "当前这份文件没有引用「素材图片」下的帧图，没有可清理的东西。\n"
                                "（本按钮只删当前配置引用过的帧目录，不碰其他文件）")
            return
        plans, lines = [], []
        for d in sorted(dirs):
            files = [f for f in d.iterdir() if f.is_file()] if d.is_dir() else []
            vids = [cand for ext in self.VIDEO_EXTS
                    for cand in (APP_DIR / (d.name + ext),
                                 APP_DIR / "素材" / (d.name + ext))
                    if cand.is_file()]
            size = sum(f.stat().st_size for f in files) + sum(v.stat().st_size for v in vids)
            if not files and not vids:
                continue
            plans.append((files, vids, d))
            mb = size / 1048576
            lines.append(f"{d.name}：{len(files)} 张帧图"
                         + (f" + 视频 {', '.join(v.name for v in vids)}" if vids else "")
                         + f"（约 {mb:.1f} MB）")
        if not plans:
            messagebox.showinfo("清理素材", "引用的帧目录都已经是空的，无需清理。")
            return
        ans = messagebox.askyesno(
            "清理素材",
            "将删除以下内容（只删当前这份作业指导书用过的素材）：\n\n"
            + "\n".join(lines)
            + "\n\n已生成的 Excel 不受影响；但删掉后重新生成 Excel 时工序页将没有配图。\n"
              "确认这份指导书已经做完、不再需要重生成时再点「是」。\n\n是 = 删除 否 = 取消")
        if not ans:
            return
        freed, failed = 0, []
        for files, vids, d in plans:
            for f in files + vids:
                try:
                    freed += f.stat().st_size
                    f.unlink()
                except Exception as e:
                    failed.append(f"{f.name}: {e}")
            try:
                if d.is_dir() and not any(d.iterdir()):
                    d.rmdir()
            except Exception:
                pass
        tip = f"已释放约 {freed / 1048576:.1f} MB"
        if failed:
            tip += "\n部分文件删除失败（可能被其他程序打开）：\n" + "\n".join(failed[:5])
        self.status.set("素材清理完成")
        messagebox.showinfo("清理素材", tip)

    def cleanup_all_test_files(self):
        """一次性清理所有测试产生的文件（视频、抽帧、测试输出）。"""
        import shutil

        # 要清理的目标
        to_clean = []  # list of (path, desc)

        # 1. 素材目录下的视频
        vid_dir = APP_DIR / "素材"
        if vid_dir.is_dir():
            for f in sorted(vid_dir.iterdir()):
                if f.is_file() and f.suffix.lower() in self.VIDEO_EXTS:
                    to_clean.append((f, f"素材/{f.name}"))

        # 2. 素材图片下以 VID_ 开头的目录（自动创建工序产生的）
        pic_dir = APP_DIR / "素材图片"
        if pic_dir.is_dir():
            for d in sorted(pic_dir.iterdir()):
                if d.is_dir() and d.name.startswith("VID_"):
                    files = [f for f in d.iterdir() if f.is_file()]
                    if files:
                        to_clean.append((d, f"素材图片/{d.name}/ ({len(files)} 张)"))

        # 3. 成品目录下文件名含「规范化版」或「demo」或「测试」的 xlsx
        out_dir = OUTPUT_DIR
        if out_dir.is_dir():
            for f in sorted(out_dir.iterdir()):
                if f.is_file() and f.suffix.lower() == ".xlsx":
                    base = f.name.lower()
                    if ("规范化版" in f.name or "demo" in base or "测试" in base or "zT系列" in base):
                        to_clean.append((f, f"成品/{f.name}"))

        if not to_clean:
            messagebox.showinfo("清理测试文件", "没有找到可清理的测试文件。\n当前状态很干净！")
            return

        lines = []
        total_size = 0
        for p, desc in to_clean:
            try:
                size = p.stat().st_size
            except OSError:
                size = 0
            mb = size / 1048576
            total_size += size
            lines.append(f"  {desc}")

        tip = "将删除以下测试文件（约 {:.1f} MB）：\n\n".format(total_size / 1048576) + "\n".join(lines)
        ans = messagebox.askyesno(
            "清理全部测试文件",
            tip + "\n确认后这些文件将无法恢复！\n\n是 = 删除  否 = 取消")
        if not ans:
            return

        freed, failed = 0, []
        for p, _desc in to_clean:
            try:
                freed += p.stat().st_size
                if p.is_dir():
                    shutil.rmtree(p)
                else:
                    p.unlink()
            except Exception as e:
                failed.append(f"{p.name}: {e}")

        msg = f"已清理，释放约 {freed / 1048576:.1f} MB"
        if failed:
            msg += "\n部分文件删除失败（可能被其他程序打开）：\n" + "\n".join(failed[:5])
        self.status.set("测试文件清理完成")
        messagebox.showinfo("清理测试文件", msg)

    def sync_progress(self):
        if getattr(self, "_sync_busy", False):
            messagebox.showinfo("提示", "上一次同步还没结束，稍等几秒")
            return
        try:
            from site_sync import load_config, collect_snapshot, upload, SyncError
            cfg = load_config(APP_DIR)
        except Exception as e:
            messagebox.showerror("无法同步", str(e))
            return
        self._sync_busy = True
        self.status.set("正在同步进度到云端看板…")

        def work():
            try:
                snap = collect_snapshot(CONFIGS_DIR, OUTPUT_DIR)
                if snap["summary"]["docs"] == 0:
                    raise SyncError("configs/ 里没有任何 YAML 配置，先去「版头信息/工序内容」填写并保存")
                upload(snap, cfg)
                tip = (f"同步成功！共 {snap['summary']['docs']} 份配置、"
                       f"{snap['summary']['processes']} 个工序。\n打开看板网页即可看到。")
            except Exception as e:
                tip = f"同步失败：{e}"
                ok = False
            else:
                ok = True
            self.after(0, lambda: self._sync_done(ok, tip))

        threading.Thread(target=work, daemon=True).start()

    def _sync_done(self, ok, tip):
        self._sync_busy = False
        self.status.set("就绪")
        (messagebox.showinfo if ok else messagebox.showerror)("云端同步", tip)

    def on_close(self):
        if self._confirm_discard():
            self.destroy()


def main():
    root = tk.Tk()
    root.withdraw()
    if IMPORT_ERROR is not None:
        messagebox.showerror("启动失败",
                             f"程序组件缺失：\n{IMPORT_ERROR}\n\n"
                             "如使用源码版请先安装依赖：\npip install openpyxl pyyaml")
        sys.exit(1)
    try:
        root.destroy()
        App().mainloop()
    except Exception:
        import traceback
        err = traceback.format_exc()
        try:
            with open(APP_DIR / "启动错误.log", "a", encoding="utf-8") as f:
                from datetime import datetime
                f.write(f"\n===== {datetime.now()} =====\n{err}")
        except OSError:
            pass
        try:
            messagebox.showerror("程序出错",
                                 f"发生错误，详情已写入同目录「启动错误.log」：\n\n{err[-800:]}")
        except tk.TclError:
            pass
        sys.exit(1)


if __name__ == "__main__":
    main()
