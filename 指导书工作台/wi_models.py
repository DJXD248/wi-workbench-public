# -*- coding: utf-8 -*-
"""
作业指导书数据模型 Python 类
支持：YAML 序列化、字段校验、工时自动汇总、模板渲染数据准备
"""
from __future__ import annotations
from dataclasses import dataclass, field, asdict
from typing import List, Optional, Dict, Any
from datetime import date, datetime
import yaml
import copy


# ============================================================
# 基础数据类
# ============================================================

@dataclass
class RevisionRecord:
    """变更记录条目"""
    date: str = ""           # YYYY-MM-DD
    content: str = ""        # 变更内容
    author: str = ""         # 变更人
    
    def to_row_data(self) -> List[str]:
        return [self.date, self.content, self.author]


@dataclass
class Step:
    """工序步骤"""
    step_no: int = 1
    content: str = ""        # 中文内容
    content_en: str = ""     # 英文内容（可选）
    images: List[str] = field(default_factory=list)  # 图片路径列表
    
    def full_content(self) -> str:
        """返回中英对照文本"""
        if self.content_en:
            return f"Step {self.step_no}第{self.step_no}步：\n{self.content}\n{self.content_en}"
        return f"第{self.step_no}步：{self.content}"


@dataclass
class Material:
    """物料/辅料"""
    code: str = ""           # 编号
    name: str = ""           # 名称
    spec: str = ""           # 规格
    qty: float = 1           # 数量
    unit: str = "PCS"        # 单位
    
    def to_row_data(self) -> List[str]:
        return [self.code, self.name, self.spec, str(self.qty), self.unit]


@dataclass
class Equipment:
    """设备/工夹具"""
    code: str = ""           # 型号/编号
    name: str = ""           # 名称
    operating_time: str = "" # 设备运行时间/参数
    
    def to_row_data(self) -> List[str]:
        return [self.code, self.name, self.operating_time]


@dataclass
class QCItem:
    """质检项目"""
    item: str = ""           # 项目
    spec: str = ""           # 标准
    parameter: str = ""      # 参数
    
    def to_row_data(self) -> List[str]:
        return [self.item, self.spec, self.parameter]


@dataclass
class ChangeRecordDetail:
    """Sheet4 变更记录详细条目"""
    modification: str = ""   # 修改内容
    checked_by: str = ""     # 审核
    approved_by: str = ""    # 批准
    date: str = ""           # 日期
    
    def to_row_data(self) -> List[str]:
        return [self.modification, self.checked_by, self.approved_by, self.date]


# ============================================================
# 核心模型
# ============================================================

@dataclass
class Header:
    """Sheet1: 文档版头"""
    # 公司/品牌
    company_name: str = "DEMO CONTROLS"
    company_subtitle: str = "示例温控器工厂"
    doc_title: str = "作业指导书"
    doc_title_en: str = "Working Instruction"
    
    # 文档编号
    file_no: str = ""
    edition: str = "A"
    product_spec: str = ""
    product_spec_cn: str = ""
    
    # 审批
    made_by: str = ""
    checked_by: str = ""
    approved_by: str = ""
    issued_date: str = ""
    
    # 变更记录（简表，Sheet1 区域）
    revision_history: List[RevisionRecord] = field(default_factory=list)
    
    # 备注
    note: str = "(★) Means key processes 表示关键工序"
    
    def validate(self) -> List[str]:
        """返回错误信息列表，空表示通过"""
        errors = []
        if not self.file_no:
            errors.append("file_no 不能为空")
        if not self.product_spec:
            errors.append("product_spec 不能为空")
        return errors
    
    def to_dict(self) -> dict:
        d = asdict(self)
        d['revision_history'] = [asdict(r) for r in self.revision_history]
        return d
    
    @classmethod
    def from_dict(cls, data: dict) -> 'Header':
        revs = [RevisionRecord(**r) for r in data.get('revision_history', [])]
        data = copy.deepcopy(data)
        data['revision_history'] = revs
        return cls(**data)


@dataclass
class Process:
    """Sheet2/3: 单个工序完整数据"""
    # 基本信息
    process_no: str = ""           # 工序号，如 DEMO-1, DEMOA-4
    process_name: str = ""         # 工位名称（中英）
    process_name_cn: str = ""      # 工位名称中文
    work_hours: int = 0            # 工时(秒)
    work_hours_display: str = ""   # 显示文本，如 "56秒"
    page_no: int = 1               # 页码
    video_no: str = ""             # 视频编号
    
    # 步骤 (最多 6 步)
    steps: List[Step] = field(default_factory=list)
    
    # 附属文件
    accessories: str = ""
    
    # 物料/设备/质检
    materials: List[Material] = field(default_factory=list)
    equipment: List[Equipment] = field(default_factory=list)
    qc_items: List[QCItem] = field(default_factory=list)
    
    # 签名区（可继承自 header，也可单独覆盖）
    made_by: str = ""
    checked_by: str = ""
    approved_by: str = ""
    process_date: str = ""
    
    def add_step(self, step_no: int, content: str, content_en: str = "") -> Step:
        step = Step(step_no=step_no, content=content, content_en=content_en)
        self.steps.append(step)
        return step
    
    def add_material(self, code: str, name: str, spec: str = "", qty: float = 1, unit: str = "PCS") -> Material:
        m = Material(code=code, name=name, spec=spec, qty=qty, unit=unit)
        self.materials.append(m)
        return m
    
    def add_equipment(self, code: str, name: str, operating_time: str = "") -> Equipment:
        e = Equipment(code=code, name=name, operating_time=operating_time)
        self.equipment.append(e)
        return e
    
    def add_qc_item(self, item: str, spec: str = "", parameter: str = "") -> QCItem:
        q = QCItem(item=item, spec=spec, parameter=parameter)
        self.qc_items.append(q)
        return q
    
    def total_material_rows(self) -> int:
        return max(len(self.materials), len(self.equipment))
    
    def to_dict(self) -> dict:
        return {
            'process_no': self.process_no,
            'process_name': self.process_name,
            'process_name_cn': self.process_name_cn,
            'work_hours': self.work_hours,
            'work_hours_display': self.work_hours_display,
            'page_no': self.page_no,
            'video_no': self.video_no,
            'steps': [asdict(s) for s in self.steps],
            'accessories': self.accessories,
            'materials': [asdict(m) for m in self.materials],
            'equipment': [asdict(e) for e in self.equipment],
            'qc_items': [asdict(q) for q in self.qc_items],
            'made_by': self.made_by,
            'checked_by': self.checked_by,
            'approved_by': self.approved_by,
            'process_date': self.process_date,
        }
    
    @classmethod
    def from_dict(cls, data: dict) -> 'Process':
        data = copy.deepcopy(data)
        data['steps'] = [Step(**s) for s in data.get('steps', [])]
        data['materials'] = [Material(**m) for m in data.get('materials', [])]
        data['equipment'] = [Equipment(**e) for e in data.get('equipment', [])]
        data['qc_items'] = [QCItem(**q) for q in data.get('qc_items', [])]
        return cls(**data)


@dataclass
class WorkInstruction:
    """完整作业指导书文档模型"""
    header: Header = field(default_factory=Header)
    processes: List[Process] = field(default_factory=list)
    change_records: List[ChangeRecordDetail] = field(default_factory=list)
    
    # 计算字段
    total_work_hours: int = 0
    total_pages: int = 1
    
    def __post_init__(self):
        self.recalc()
    
    def recalc(self):
        """重新计算汇总字段"""
        self.total_work_hours = sum(p.work_hours for p in self.processes)
        self.total_pages = len(self.processes)  # 每工序至少1页
    
    def add_process(self, process: Process) -> Process:
        self.processes.append(process)
        self.recalc()
        return process
    
    def add_change_record(self, record: ChangeRecordDetail):
        self.change_records.append(record)
    
    # ---- YAML 序列化 ----
    def to_yaml(self, path: str):
        """保存为 YAML 配置文件"""
        data = {
            'header': self.header.to_dict(),
            'processes': [p.to_dict() for p in self.processes],
            'change_records': [asdict(r) for r in self.change_records],
        }
        with open(path, 'w', encoding='utf-8') as f:
            yaml.dump(data, f, allow_unicode=True, sort_keys=False, indent=2)
    
    @classmethod
    def from_yaml(cls, path: str) -> 'WorkInstruction':
        """从 YAML 加载"""
        with open(path, 'r', encoding='utf-8') as f:
            data = yaml.safe_load(f)
        wi = cls()
        wi.header = Header.from_dict(data.get('header', {}))
        wi.processes = [Process.from_dict(p) for p in data.get('processes', [])]
        wi.change_records = [ChangeRecordDetail(**r) for r in data.get('change_records', [])]
        wi.recalc()
        return wi
    
    # ---- 生成器用的扁平数据 ----
    def to_generator_data(self) -> Dict[str, Any]:
        """转换为生成器易用的扁平字典"""
        return {
            'header': self.header.to_dict(),
            'processes': [p.to_dict() for p in self.processes],
            'change_records': [asdict(r) for r in self.change_records],
            'total_work_hours': self.total_work_hours,
            'total_pages': self.total_pages,
            'generated_at': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
        }


# ============================================================
# 预设/工厂函数
# ============================================================

def create_sb_template() -> WorkInstruction:
    """创建 DEMOB 系列示例模板"""
    wi = WorkInstruction()
    
    # 版头
    wi.header = Header(
        file_no="WI-DEMO-W-0183-2024",
        edition="A",
        product_spec="demo series controller",
        product_spec_cn="演示系列组件作业指导书",
        made_by="张三",
        checked_by="李四",
        approved_by="王五",
        issued_date=date.today().strftime('%Y-%m-%d'),
        revision_history=[
            RevisionRecord(date="2024-01-15", content="Initial release", author="张三"),
            RevisionRecord(date="2024-03-20", content="Updated vacuum value 50Pa->1000Pa", author="李四"),
        ]
    )
    
    # 工序 DEMO-1
    p1 = Process(
        process_no="DEMO-1",
        process_name="Stamping DEMO000AA0001A",
        process_name_cn="冲压 DEMO000AA0001A",
        work_hours=56,
        work_hours_display="56秒",
        page_no=1,
    )
    p1.add_step(1, "将材料 BA13010DEMOB003A 放置精密矫正机前，需将百分表调零", 
                "Before putting material BA13010DEMOB003A on straightening machine, adjust dial gauge to 0")
    p1.add_step(2, "调零完成后，将材料放入矫正机，通过时百分表显示 0.2 状态",
                "After zeroing, put material in machine, gauge should show 0.2 when passing through")
    p1.add_step(3, "取出材料，目视检查表面无划伤、变形",
                "Take out material, visual inspect surface no scratch/deformation")
    p1.add_material("BA13010DEMOB003A", "SUS301", "0.1mm", 1, "PCS")
    p1.add_equipment("F0733-01", "精密矫正仪", "压力表 0.2MPa")
    p1.add_qc_item("材料平整度", "≤0.2mm", "百分表测量")
    p1.add_qc_item("表面质量", "无划伤/变形", "目视全检")
    wi.add_process(p1)
    
    # 工序 DEMO-2
    p2 = Process(
        process_no="DEMO-2",
        process_name="Stamping DEMO000AA0006A",
        process_name_cn="冲压 DEMO000AA0006A",
        work_hours=48,
        work_hours_display="48秒",
        page_no=2,
    )
    p2.add_step(1, "第二冲压工序...", "Second stamping process...")
    p2.add_material("BA13010DEMOB004A", "SUS304", "0.15mm", 1, "PCS")
    p2.add_equipment("F0733-02", "精密冲床", "冲压力 2.5T")
    p2.add_qc_item("尺寸精度", "±0.05mm", "投影仪测量")
    wi.add_process(p2)
    
    # 更多工序可继续添加...
    
    return wi


def create_welding_template() -> WorkInstruction:
    """创建焊接系列示例模板 (DEMOA)"""
    wi = WorkInstruction()
    
    wi.header = Header(
        file_no="WI-DEMO-0002",
        edition="B",
        product_spec="DEMOA2Q0000HIS01",
        product_spec_cn="DEMOA 焊接作业指导书",
        made_by="张三",
        checked_by="李四",
        approved_by="王五",
        issued_date=date.today().strftime('%Y-%m-%d'),
        revision_history=[
            RevisionRecord(date="2021-03-09", content="Initial release", author="张三"),
            RevisionRecord(date="2022-06-15", content="DEMOA-4 删除焊枪气体压力值", author="张三"),
        ]
    )
    
    # DEMOA-1 激光打印
    p1 = Process(
        process_no="DEMOA-1",
        process_name="Laser printing production cycle & batch barcode",
        process_name_cn="激光打印生产周期及批号条码",
        work_hours=56,
        work_hours_display="56秒",
        page_no=1,
    )
    p1.add_step(1, "激光打印机设置参数", "Set laser printer parameters")
    p1.add_step(2, "扫描工单条码", "Scan work order barcode")
    p1.add_step(3, "确认打印内容正确", "Verify print content correct")
    p1.add_equipment("EQ-0001", "激光打印机", "功率 20W")
    p1.add_qc_item("打印清晰度", "字符完整可读", "目视检查")
    p1.add_qc_item("条码扫描率", "100% 可读", "扫码枪验证")
    wi.add_process(p1)
    
    # DEMOA-2 卷板
    p2 = Process(
        process_no="DEMOA-2",
        process_name="Rolling",
        process_name_cn="卷板",
        work_hours=120,
        work_hours_display="120秒",
        page_no=2,
    )
    p2.add_step(1, "卷板机调整间隙", "Adjust rolling machine gap")
    p2.add_step(2, "送料卷制", "Feed and roll")
    p2.add_step(3, "检查圆度", "Check roundness")
    p2.add_equipment("F0832-01", "卷板机", "间隙 1.2mm")
    p2.add_qc_item("圆度", "≤0.1mm", "圆度仪测量")
    wi.add_process(p2)
    
    # 可继续添加 DEMOA-3 到 DEMOA-21...
    
    return wi


# ============================================================
# 测试入口
# ============================================================

if __name__ == "__main__":
    import os
    
    # 创建示例配置目录
    os.makedirs("configs/DEMO_series", exist_ok=True)
    os.makedirs("configs/Welding_DEMOA", exist_ok=True)
    os.makedirs("configs/Welding_9SU", exist_ok=True)
    os.makedirs("成品", exist_ok=True)
    
    # 生成示例 YAML
    sb_wi = create_sb_template()
    sb_wi.to_yaml("configs/DEMO_series/DEMO_demo.yaml")
    print(f"✓ 生成 DEMOB 示例: configs/DEMO_series/DEMO_demo.yaml ({len(sb_wi.processes)} 个工序, 合计 {sb_wi.total_work_hours}秒)")
    
    weld_wi = create_welding_template()
    weld_wi.to_yaml("configs/Welding_DEMOA/DEMOA_demo.yaml")
    print(f"✓ 生成焊接示例: configs/Welding_DEMOA/DEMOA_demo.yaml ({len(weld_wi.processes)} 个工序, 合计 {weld_wi.total_work_hours}秒)")
    
    # 验证读取
    loaded = WorkInstruction.from_yaml("configs/DEMO_series/DEMO_demo.yaml")
    print(f"✓ 读取验证: {loaded.header.file_no}, {len(loaded.processes)} 个工序")
    
    print("\n=== 使用方法 ===")
    print("1. 修改 configs/ 下的 YAML 文件")
    print("2. 运行 python quick_start.py 生成 Excel")
    print("3. 或在代码中: wi = WorkInstruction.from_yaml(...); wi.to_yaml(...)")