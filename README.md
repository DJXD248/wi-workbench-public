# 作业指导书模块化生成系统 - 使用说明

> **先看路径**：程序代码全部在子目录 **`D:\工作自动化\指导书工作台\`**，日常入口是
> `指导书工作台\启动工作台.bat`（或发布包里的 exe）。本文下面部分示例写的是老版本的根目录路径，按此对应。
> **接手开发**请先读同目录的 **《交接文档.md》** 与 **《更新日志.md》**；这个项目已纳入 Git 版本管理：真数据在私有仓 `DJXD248/work-automation`（云端读不到，别试），**你现在 clone 的这个 `DJXD248/wi-workbench-public` 是它的净化公开镜像**（见《公开仓说明.md》）。
> **要开一个云端会话替你做**：直接复制 **《云端开场白.md》** 里的那段话发给它，一条命令的验收闸是 `指导书工作台\tests\verify.py`。

## 日常使用（推荐）：图形界面工作台

**双击《指导书工作台\启动工作台.bat》即可打开窗口**（发布包版在 `指导书工作台\发布包\作业指导书工作台.exe`），全程填表操作，不需要写任何命令、不需要碰 YAML 文件。

| 界面位置 | 干什么 |
|---|---|
| 顶部按钮：新建 / 打开YAML / 保存 / 另存为 | 管理配置文件（自动读写 configs/ 下的 YAML） |
| 顶部按钮：**生成 Excel** | 按《作业指导书版头.xlsx》标准版式输出到 output/ |
| 标签页 1「版头信息」 | 填文件编号、版本、产品规格、编制/审核/批准，及两张变更表 |
| 标签页 2「工序内容」 | 左侧工序列表可增删/上下移；右侧填工序基本信息，四个子表分别添加步骤、物料、设备、质检项（双击=修改） |
| 标签页 3「工具箱」 | 规范化同事发来的旧 Excel（自动提取内容并可载入编辑器修改）；**导入视频自动抽帧**（操作视频→自动截帧，弹窗可选**起始帧/取帧间隔/每几张图合并为一步**，多图横排同步；载入时按视频时长自动填工时）；**AI 起草步骤文字**（让大模型看当前工序的配图自动写步骤文字，需 keys.json，详见《用户使用指南.md》）；**清理素材**（做完后删除本文件用过的帧图和同名视频）；生成演示示例文件；校验当前填写；同步进度到云端看板；一键打开输出目录 |

详细操作步骤请看同目录的 **《用户使用指南.md》**（面向使用者，写清了每个按钮怎么用、注意事项）。

首次使用建议：打开「工具箱」→ 点「生成演示示例文件」→ 再「打开YAML」选 `configs/DEMO_series/DEMO_demo.yaml` 看看示例长什么样。

## 目录结构
```
D:\工作自动化\                     # Git 仓库根（私有远端：DJXD248/work-automation）
├── README.md / 更新日志.md / 交接文档.md   # 文档（每次改动都要更新后两份）
├── 云端开场白.md                   # 复制这段话发给新的云端会话，让它接手本项目
├── .gitignore / .gitattributes    # 入库边界：密钥、素材、成品、发布包、构建产物都不入库
├── 指导书工作台\                  # ★ 主程序（下面这些文件都在这个子目录里）
│   ├── 启动工作台.bat             # ★ 双击打开图形界面
│   ├── wi_workbench.py            # 图形界面程序
│   ├── 用户使用指南.md            # ★ 面向使用者的详细图文指南
│   ├── 作业指导书版头.xlsx        # 标准版头模板（封面=「清单」页，工序页=「新格式」页）
│   ├── wi_schema.yaml             # 统一数据模型定义（文档用）
│   ├── wi_models.py               # Python 数据模型类
│   ├── excel_generator.py         # 核心生成器 + 规范化器（版式与图片排布都在这）
│   ├── video_frames.py            # 视频抽帧（导入视频自动截关键画面）
│   ├── ai_draft.py                # AI 起草步骤文字（调用大模型看图写文字）
│   ├── probe.py / providers.json  # 大模型渠道清单（从「大模型探测」同步来的副本）
│   ├── site_sync.py               # 云端进度看板同步（真实口令在 site_sync.json，不入库）
│   ├── configs\                   # YAML 配置库（按产品系列分目录）
│   ├── 素材\ / 素材图片\ / 成品\   # 视频源、帧图、生成结果（都不入库）
│   ├── 发布包\ + *.zip            # 免安装 exe 与压缩包（不入库，走 GitHub Releases）
│   └── 作业指导书模板-*.xlsx       # 公司原始参考模板（太大，不入库）
├── 签名工具\                      # SmartScreen 信任证书 + 重打包后重新签名的脚本
├── 看板站点\                      # 云端进度看板的网页源码（Qoder Sites）
├── 大模型探测\                    # 独立子项目：免费大模型自动探测器
└── 旧脚本存档\                    # 一次性调试脚本（打包临时产物不入库）
```

> ⚠ 命令行工具 `normalize_single.py`、`quick_start.py` 现在在 `旧脚本存档\`，是图形界面之外的备用方式；
> 日常一律走 `指导书工作台\启动工作台.bat`。

> ⚠ 重新打包注意：每次用 PyInstaller 重新生成 exe 后，签名会失效，必须再用
> `签名工具\签名exe.ps1` 重新签名，并重建 zip；公司电脑首次使用需运行「安装信任证书.bat」。

> 生成策略说明（2026-10-18 现状）：生成器**直接以模板文件为底本原位填值**，
> 工序页通过复制模板内「新格式」表生成，原有格式、合并单元格、列宽行高全部保留。
> 每工序版式：步骤块**图在上、动作在下**，每块最多 3 图**横向并排**，文字自动合并整块+加高行确保**完整可见**，
> 含英文的字体统一 Times New Roman，正文/标题**英文在中文前**；页脚**三栏** 编制/审核/批准；工时按实测秒数不带"约"。
> 另右侧设备6行、物料8行、质检3行。
> **步骤数量不限 4 步**：一道工序有 5、6、8 步时，工序页自动向下克隆整段版面（每段 17 行、含左右两块），
> 页脚签字栏跟着下移，整页缩放到一张打印纸——**一道工序不会被拆成两页，也不会截断步骤**。
>
> 规范化说明（2026-10-18 定档的硬规则）：「规范化旧文件」**只重做版头**（第 5 行以上），
> 正文（每步图片+文字、右侧四张表、页脚签字）**原样保留、一张图都不丢**，无论原文有没有缺图缺字。

## 依赖安装
```bash
pip install openpyxl pyyaml
```
或使用国内镜像：
```bash
pip install openpyxl pyyaml -i https://pypi.tuna.tsinghua.edu.cn/simple
```

## 快速开始

> 下面的 `quick_start.py` 是早期命令行入口，脚本已归档在 `旧脚本存档\`，它里面的路径写法还是老目录结构，
> **日常不必使用**；所有功能都以图形界面（`指导书工作台\启动工作台.bat`）为准。留在这里仅供查阅数据流与参数含义。

### 1. 生成演示文件（首次运行）
```bash
cd D:\工作自动化
python quick_start.py demo
```
这将创建：
- `configs/` 下的示例 YAML 配置
- `output/` 下的生成 Excel（需有《作业指导书版头.xlsx》模板）

### 2. 从 YAML 生成标准作业指导书
```bash
# 编辑配置后生成
python quick_start.py gen configs/DEMO_series/DEMO_demo.yaml

# 指定输出路径
python quick_start.py gen configs/DEMO_series/DEMO_demo.yaml -o output/DEMOB正式版.xlsx
```

### 3. 规范化同事发来的旧文件
```bash
# 基本用法：自动输出到 output/ 目录
python quick_start.py norm "同事发来的作业指导书.xlsx"

# 指定输出并同时导出 YAML 配置（便于后续修改）
python quick_start.py norm "旧文件.xlsx" -o "output/标准版.xlsx" -y "configs/提取的配置.yaml"
```

### 4. 创建新产品系列的空白配置
```bash
python quick_start.py new 9SU3A
```
这将创建 `configs/9SU3A/9SU3A_template.yaml`，编辑后即可生成。

### 5. 验证 YAML 配置
```bash
python quick_start.py validate configs/DEMO_series/DEMO_demo.yaml
```

## YAML 配置文件详解

### 版头字段 (header)
```yaml
header:
  company_name: "DEMO CONTROLS"          # 公司名
  company_subtitle: "示例温控器工厂"        # 工厂/子标题
  doc_title: "作业指导书"                  # 文档标题
  doc_title_en: "Working Instruction"     # 英文标题
  file_no: "WI-DEMO-W-0183-2024"         # 文件编号（建议统一规则）
  edition: "A"                             # 版本
  product_spec: "demo series controller"        # 产品规格（英文）
  product_spec_cn: "演示系列组件作业指导书"  # 产品规格（中文）
  made_by: "张三"                          # 编制人
  checked_by: "李四"                       # 审核人
  approved_by: "王五"                      # 批准人
  issued_date: "2024-01-15"               # 发放日期
  revision_history:                        # 变更记录
    - date: "2024-01-15"
      content: "Initial release"
      author: "张三"
    - date: "2024-03-20"
      content: "Updated vacuum value 50Pa->1000Pa"
      author: "李四"
  note: "(★) Means key processes 表示关键工序"
```

### 工序字段 (processes 列表)
```yaml
processes:
  - process_no: "DEMO-1"                     # 工序号
    process_name: "Stamping DEMO000AA0001A" # 工位名称（英文）
    process_name_cn: "冲压 DEMO000AA0001A"   # 工位名称（中文）
    work_hours: 56                         # 工时（秒，数字）
    work_hours_display: "56秒"             # 工时显示文本
    page_no: 1                             # 页码
    video_no: ""                           # 视频编号（可选）
    steps:                                 # 步骤列表（数量不限，工序页自动向下扩版面）
      - step_no: 1
        content: "将材料放置精密矫正机前，需将百分表调零"
        content_en: "Before putting material on straightening machine, adjust dial gauge to 0"
    accessories: "WI-DEMO-E-00014-2024"   # 附属文件
    materials:                             # 物料/辅料
      - code: "BA13010DEMOB003A"
        name: "SUS301"
        spec: "0.1mm"
        qty: 1
        unit: "PCS"
    equipment:                             # 设备/工夹具
      - code: "F0733-01"
        name: "精密矫正仪"
        operating_time: "压力表 0.2MPa"
    qc_items:                              # 质检项目
      - item: "材料平整度"
        spec: "≤0.2mm"
        parameter: "百分表测量"
      - item: "表面质量"
        spec: "无划伤/变形"
        parameter: "目视全检"
```

## 常用工作流

### 工作流 A：新建作业指导书
1. `python quick_start.py new 产品系列名` → 创建空白配置
2. 编辑 `configs/产品系列名/配置.yaml` 填入实际数据
3. `python quick_start.py gen configs/产品系列名/配置.yaml` → 生成标准 Excel
4. 发给车间使用

### 工作流 B：维护现有文件（同事发来旧版）
1. 收到旧文件 → `python quick_start.py norm "旧文件.xlsx" -y "configs/提取.yaml"`
2. 检查生成的 `output/旧文件_规范化版.xlsx` 是否正确
3. 如需微调：编辑 `configs/提取.yaml` → 重新 `gen`
4. 归档 YAML 到 `configs/` 对应产品系列目录

### 工作流 C：批量更新版头（公司名变更、编号规则变更等）
1. 修改 `configs/` 下所有 YAML 的 `header` 字段
2. 批量重新生成：写个小脚本遍历 `configs/` 目录调用 `generate_from_yaml()`

## 进阶：在 Python 代码中使用

```python
from wi_models import WorkInstruction, create_sb_template
from excel_generator import ExcelGenerator

# 方式1：从模板创建
wi = create_sb_template()
wi.header.file_no = "新编号"
wi.header.edition = "B"

# 方式2：从 YAML 加载
wi = WorkInstruction.from_yaml("configs/DEMO_series/DEMO_demo.yaml")

# 方式3：手动构建
wi = WorkInstruction()
wi.header.file_no = "WI-DEMO-W-XXXX-2024"
# ... 填充其他字段 ...
from wi_models import Process
p = Process(process_no="DEMO-1", process_name="...", work_hours=60)
p.add_step(1, "步骤1")
p.add_material("代码", "名称", "规格", 1, "PCS")
wi.add_process(p)

# 生成 Excel
gen = ExcelGenerator("作业指导书版头.xlsx")
gen.generate(wi, "output/自定义.xlsx")
```

## 文件编号建议规则
```
WI-DEMO-W-{年份}{月份:02d}-{流水号:04d}
例：DEMO-RDD-W-202403-0001
```
可在 `wi_models.py` 中添加自动生成逻辑。

## 常见问题

**Q: 生成的 Excel 样式和模板不一致？**
A: 生成器会复制模板的列宽、行高、合并单元格、字体、边框等。如有细微差异，可微调 `excel_generator.py` 中的 `HEADER_CELL_MAP` 和 `PROCESS_CELL_MAP` 坐标。

**Q: 工序超过 4 步、6 步怎么办？**
A: 直接填就行，不用拆工序。工序页会自动向下克隆新的步骤区，页脚签字栏跟着下移，整道工序仍然缩放到一张打印纸上。只有图片实在过多、缩到看不清时，才考虑人为拆成两道工序。

**Q: 图片怎么处理？**
A: 已支持自动插图：在 YAML 的 `steps[i].images` 里列图片路径（多图用分号分隔），生成 Excel 时同一步最多 3 张图**横向并排**插进对应步骤块的图区（图在上、文字在下）。图片可由「工具箱→导入视频自动抽帧」生成（弹窗可选起始帧/取帧间隔/每几张图合并为一步），步骤文字可由「工具箱→AI 起草步骤文字」自动补写。

**Q: 如何支持 Sheet3 结构（另一种工序页）？**
A: 在 `excel_generator.py` 中添加 `_generate_process_sheet_v2` 方法，按 Sheet3 坐标映射写入，通过 `template_type` 参数切换。

## 版本迭代建议
1. **v1.x (当前)**：YAML 配置 + 模板生成（工序页版式对齐真实模板：图上文下/横排/英文在前/Times New Roman/三栏签署）+ 单文件规范化 + 视频抽帧（可选起点/间隔/每N帧合并一步+按视频填工时）+ 图片自动插入 + AI 起草步骤文字 + 云端进度看板
2. **v2.0**：多用户协作、版本管理、审批流集成

## 联系支持
如有问题，请检查：
1. Python 版本 ≥ 3.8
2. `openpyxl`、`pyyaml` 已安装
3. `作业指导书版头.xlsx` 在工作目录下
4. 输出目录 `output/` 有写入权限