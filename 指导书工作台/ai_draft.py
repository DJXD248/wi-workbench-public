# -*- coding: utf-8 -*-
"""AI 起草：看步骤配图，自动写一句操作步骤文字；看工程图纸，提取质检参数。

复用 vendored 的 probe.py（来自 大模型探测 项目）选渠道，调 OpenAI 兼容的
chat/completions 多模态接口。API key 读 exe/脚本同目录 keys.json，
开发目录回退读 大模型探测 子目录的 keys.json。
keys.json 含私密 key，绝不进 zip/git。
"""
from __future__ import annotations

import base64
import io
import json
import sys
import urllib.parse
from pathlib import Path

import probe

HERE = probe.HERE

# 名字像视觉但实际不是对话模型的黑名单片段
BAD_MODEL_HINTS = ("edit", "embedding", "bge", "search", "rerank", "clip", "ocr-")
# 文生图模型（只画图不看图）：名字带 image 但不带这些片的才排除
VISION_WORD_HINTS = ("vl", "vision", "visual", "mm", "multimodal")


def _is_chat_vision(name: str) -> bool:
    n = name.lower()
    if "image" in n and not any(v in n for v in VISION_WORD_HINTS):
        return False  # 如 Qwen/Qwen-Image 是文生图
    return True



class DraftError(Exception):
    pass


def load_keys() -> dict:
    """按顺序找 keys.json：exe/脚本同目录 → 大模型探测 子目录"""
    candidates = [HERE / "keys.json"]
    if not getattr(sys, "frozen", False):
        candidates.append(HERE / "大模型探测" / "keys.json")
    for c in candidates:
        if c.exists():
            k = probe.load_keys(c)
            if k:
                return k
    return {}


def load_custom_models() -> list:
    """从 ai_models.json 加载用户自定义的 AI 模型配置。

    返回 [{name, base_url, key, model, type}, ...]
    """
    config_path = HERE / "ai_models.json"
    if not config_path.exists():
        return []

    try:
        with open(config_path, "r", encoding="utf-8") as f:
            config = json.load(f)
    except Exception as e:
        return []

    models = config.get("models", [])
    result = []
    for m in models:
        if not m.get("api_key") or m["api_key"].startswith("sk-请替换"):
            continue  # 跳过未配置的模型
        result.append({
            "name": m.get("name", "自定义模型"),
            "base_url": m.get("base_url", ""),
            "key": m.get("api_key", ""),
            "model": m.get("model", ""),
            "type": m.get("type", "custom")
        })
    return result


def find_detector_dir() -> "Path | None":
    """定位随包附带的「免费大模型自动探测器」目录。

    与 load_keys 的回退路径一致：探测器是工作台同目录下的 `大模型探测`
    子文件夹（发布包里就自带它）。找不到返回 None。
    """
    base = probe.HERE
    cand = base / "大模型探测"
    return cand if cand.is_dir() else None


def _read_json(path) -> dict:
    try:
        return json.loads(Path(path).read_text(encoding="utf-8"))
    except Exception:
        return {}


def detector_candidates(family: str = "vision", on_note=None) -> list:
    """一键接入探测器「已录入的免费模型」，离线返回 AI 候选（不联网探测）。

    数据来源：
      - `大模型探测/keys.json`：用户在探测器里「设置 → 填写 API key」录入的 key，
        按 provider id 存。只有录入了 key 的渠道才会被接入。
      - `大模型探测/给Qoder的免费模型清单.json`：探测器已联网核实过的免费模型清单，
        每条模型带 `free` 标记与 `vision`/`chat` 验证结果——这是离线拿免费模型名的权威来源。
    若清单缺失，回退到 `大模型探测/providers.json` 的 free_rule + probe_model。

    返回 [{name, base_url, key, model}, ...]，结构与 vision_candidates 一致，
    可直接喂给 StepDrafter / BlueprintParser。
    """
    d = find_detector_dir()
    if d is None:
        raise DraftError(
            "未检测到「免费大模型自动探测器」目录（应与工作台同目录）。")

    raw_keys = _read_json(d / "keys.json") or {}
    # 过滤掉空值/纯空白/示例占位（探测器 keys.example.json 里写的是"请把…填在这里"）
    keys = {}
    for k, v in raw_keys.items():
        if isinstance(v, str):
            v = v.strip()
        if v and not v.startswith("请把"):
            keys[k] = v
    if not keys:
        raise DraftError(
            "探测器目录存在，但还没录入任何 key。\n\n"
            "请先双击「免费大模型自动探测器」，在「设置 → 填写 API key」里"
            "粘贴 key 并保存，再回来点这个按钮。")

    out = []
    curated = _read_json(d / "给Qoder的免费模型清单.json")
    if curated:
        prov_path = d / "providers.json"
        providers = probe.load_providers(prov_path) if prov_path.exists() else []
        pmap = {p["id"]: p for p in providers}
        for ch in curated.get("channels", []):
            kid = ch.get("id")
            if kid not in keys:
                continue  # 只接入用户已录入 key 的渠道
            prov = pmap.get(kid, {})
            base_url = ch.get("base_url") or prov.get("base_url", "")
            if not base_url:
                continue
            for m in ch.get("models", []):
                if not m.get("free"):
                    continue
                mid = m.get("id")
                if not mid:
                    continue
                if family == "vision":
                    vis = m.get("vision")
                    if isinstance(vis, dict):
                        pass  # 探测器已联网验证过能看图
                    elif probe.is_vision_model(mid) and _is_chat_vision(mid):
                        pass
                    else:
                        continue
                out.append({"name": ch.get("name", kid),
                            "base_url": base_url,
                            "key": keys[kid], "model": mid})

    if not out:
        # 回退：用 providers.json 的 free_rule + probe_model（离线、可能只含文本模型）
        for p in probe.load_providers(d / "providers.json"):
            kid = p["id"]
            if kid not in keys:
                continue
            models = (probe.free_models_of(p.get("models", []), p.get("free_rule"))
                      or ([p["probe_model"]] if p.get("probe_model") else []))
            for mid in models[:3]:
                if family == "vision" and not (
                        probe.is_vision_model(mid) and _is_chat_vision(mid)):
                    continue
                out.append({"name": p.get("name", kid),
                            "base_url": p["base_url"],
                            "key": keys[kid], "model": mid})

    if not out:
        raise DraftError(
            "探测器里录入的渠道没有可用的免费模型（可能是 free_rule 未配置或模型清单缺失）。\n"
            "可改用「AI 模型配置」手动添加自定义模型。")

    # 与 vision_candidates 行为一致：把用户自定义模型也并进来
    out.extend(load_custom_models())
    if on_note:
        note = f"已接入探测器 {len(out)} 个免费模型（离线，无需重新探测）"
        if any(c.get("type") == "custom" for c in out):
            note += "（含自定义模型）"
        on_note(note)
    return out


def vision_candidates(on_note=None) -> list:
    """探测全部渠道，返回按可用度排序的看图模型候选：
    [{name, base_url, key, model}, ...]"""
    keys = load_keys()
    if not keys:
        raise DraftError(
            "找不到任何 API key。\n\n请先用「免费大模型自动探测器」注册领 key，"
            f"然后把 key 填到这里的 keys.json：\n{HERE / 'keys.json'}")
    providers = probe.load_providers()
    if on_note:
        on_note("正在探测哪些渠道当前可用（约几十秒）…")
    results = probe.probe_all(providers, keys, total_budget=60)
    usable = [r for r in results
              if r["status"] in (probe.STATUS_OK_KEY, probe.STATUS_OK_NOKEY)
              and not r["id"].startswith(("ollama", "lmstudio"))]
    usable.sort(key=lambda r: (0 if r["status"] == probe.STATUS_OK_KEY else 1,
                               r["latency_ms"] if r["latency_ms"] is not None else 10 ** 9))
    pmap = {p["id"]: p for p in providers}
    out = []
    for r in usable:
        vis = [m for m in r["models"]
               if probe.is_vision_model(m) and _is_chat_vision(m)
               and not any(b in m.lower() for b in BAD_MODEL_HINTS)]
        if r["free_models"]:
            vis = [m for m in vis if m in r["free_models"]] or vis
        prov = pmap.get(r["id"])
        if not prov:
            continue
        for m in vis[:3]:
            out.append({"name": r["name"], "base_url": prov["base_url"],
                        "key": keys.get(r["id"], ""), "model": m})

    # 添加用户自定义模型
    custom_models = load_custom_models()
    for cm in custom_models:
        if cm["model"] and cm["base_url"]:
            out.append({
                "name": cm["name"],
                "base_url": cm["base_url"],
                "key": cm["key"],
                "model": cm["model"],
                "type": cm.get("type", "custom")
            })

    if not out:
        raise DraftError("所有可用渠道里都找不到能看图的模型。\n"
                         "可先用探测器工具确认网络与 key，或改用免费的智谱/魔搭渠道。")
    if on_note:
        first = out[0]
        note = f"共 {len(out)} 个看图模型候选，首选 {first['name']}/{first['model']}"
        if custom_models:
            note += f"（含 {len(custom_models)} 个自定义模型）"
        on_note(note)
    return out


def _image_data_url(path: str, max_side=900, quality=72) -> str:
    from PIL import Image
    p = Path(path)
    if not p.exists():
        raise DraftError(f"图片不存在：{p}")
    img = Image.open(p).convert("RGB")
    if max(img.size) > max_side:
        img.thumbnail((max_side, max_side))
    buf = io.BytesIO()
    img.save(buf, format="JPEG", quality=quality)
    return "data:image/jpeg;base64," + base64.b64encode(buf.getvalue()).decode()


def _call_channel(channel: dict, images: list, process_name: str, step_no: int,
                  timeout=75) -> str:
    import requests
    parts = [{"type": "text", "text":
              f"这是装配作业指导书「{process_name}」工序第{step_no}步的"
              f"{len(images)}张连续操作照片（按时间先后排列）。"
              "请用一句中文祈使句写出这一步的操作内容，不超过26个字，动词开头"
              "（例如“将引线插入端子壳体”“用扭力扳手拧紧螺母”）。"
              "只输出这句话本身，不要编号、不要引号、不要其他解释。"}]
    for img in images[:3]:
        parts.append({"type": "image_url",
                      "image_url": {"url": _image_data_url(img)}})
    url = urllib.parse.urljoin(channel["base_url"].rstrip("/") + "/", "chat/completions")
    body = {"model": channel["model"],
            "messages": [{"role": "user", "content": parts}],
            "max_tokens": 120}
    headers = {"Content-Type": "application/json"}
    if channel["key"]:
        headers["Authorization"] = f"Bearer {channel['key']}"
    r = requests.post(url, json=body, headers=headers, timeout=timeout)
    if r.status_code != 200:
        raise DraftError(f"{channel['name']}/{channel['model']} 返回 {r.status_code}："
                         f"{r.text[:160]}")
    try:
        content = r.json()["choices"][0]["message"]["content"]
    except Exception:
        raise DraftError(f"{channel['name']} 返回内容无法解析：{r.text[:160]}")
    lines = (content or "").strip().strip('"\"\"').splitlines()
    return lines[0].strip() if lines else ""


def _call_blueprint(channel: dict, images: list, process_name: str, timeout=90) -> list:
    """从工程图纸/BOM表/工艺卡中提取 QC Items。

    返回 JSON 数组：[{"item": "检验项目名称", "spec": "规格要求", "parameter": "测量方法"}]
    """
    import requests
    text_prompt = (
        f"这是一份「{process_name}\"工序相关的工程图纸、BOM表或工艺卡。\n"
        "请识别其中的**质量检验要求**（QC Items），以 JSON 格式输出，"
        "不要包含任何其他文字说明。格式如下：\n"
        "[\n"
        '  {"item": "推拉力读数", "spec": "50±10N", "parameter": "使用数显推拉力计读取，记录于点检表"},\n'
        '  {"item": "外观检查", "spec": "无毛刺、无划伤", "parameter": "目视检查"}\n'
        "]\n\n"
        "提取规则：\n"
        "1. item = 检验项目名称（中文）\n"
        "2. spec = 规格上限/下限或公差范围（如有数值则尽量精确）\n"
        "3. parameter = 检验方法或依据（简要描述）\n"
        "如果图纸中没有明确的检验参数，请输出空数组 []。\n"
        "只输出 JSON，不要代码块标记。"
    )
    parts = [{"type": "text", "text": text_prompt}]
    for img in images[:4]:
        parts.append({"type": "image_url",
                      "image_url": {"url": _image_data_url(img, max_side=1200, quality=85)}})
    url = urllib.parse.urljoin(channel["base_url"].rstrip("/") + "/", "chat/completions")
    body = {
        "model": channel["model"],
        "messages": [{"role": "user", "content": parts}],
        "max_tokens": 500,
        "response_format": {"type": "json_object"},
    }
    headers = {"Content-Type": "application/json"}
    if channel["key"]:
        headers["Authorization"] = f"Bearer {channel['key']}"
    r = requests.post(url, json=body, headers=headers, timeout=timeout)
    if r.status_code != 200:
        raise DraftError(f"{channel['name']}/{channel['model']} 返回 {r.status_code}："
                         f"{r.text[:160]}")
    try:
        content = r.json()["choices"][0]["message"]["content"]
        # 处理可能的代码块包裹
        content = content.strip()
        if content.startswith("```"):
            content = content.split("```")[1]
            if content.startswith("json"):
                content = content[4:]
            content = content.rstrip("```").strip()
        result = json.loads(content)
        if isinstance(result, list):
            # 规范化每一项
            items = []
            for row in result:
                item = row.get("item", "")
                if not item:
                    continue
                items.append({
                    "item": item,
                    "spec": row.get("spec", ""),
                    "parameter": row.get("parameter", ""),
                })
            return items
        elif isinstance(result, dict):
            # 可能是 {qc_items: [...]} 结构
            rows = result.get("qc_items", result.get("items", [result]))
            items = []
            for row in rows:
                item = row.get("item", "")
                if not item:
                    continue
                items.append({
                    "item": item,
                    "spec": row.get("spec", ""),
                    "parameter": row.get("parameter", ""),
                })
            return items
    except Exception as e:
        raise DraftError(f"{channel['name']} 返回内容无法解析为 JSON：{str(e)[:160]}")
    raise DraftError(f"{channel['name']} 返回格式不符合预期")


class StepDrafter:
    """逐句起草；某个渠道失败自动换下一个候选，成功后沿用最稳的那个。"""

    def __init__(self, candidates: list):
        self.cands = candidates
        self.idx = 0
        self.worked = None

    def draft(self, images, *, process_name="", step_no=1) -> str:
        order = ([self.worked] if self.worked else []) + \
                self.cands[self.idx:] + self.cands[:self.idx]
        errs = []
        for i, ch in enumerate(order):
            try:
                text = _call_channel(ch, images, process_name, step_no)
                if text:
                    self.worked = ch
                    if self.worked not in self.cands[:self.idx + 1]:
                        self.idx = self.cands.index(ch)
                    return text
                errs.append(f"{ch['name']}: 空回复")
            except Exception as e:
                errs.append(str(e)[:120])
            self.idx = min(self.idx + i + 1, len(self.cands))
        raise DraftError("所有看图渠道都失败：\n- " + "\n- ".join(errs[-4:]))


# ---- 蓝图解析器 -------------------------------------------------------

class BlueprintParser:
    """解析工程图纸中的 QC Items；某个渠道失败自动换下一个候选。"""

    def __init__(self, candidates: list):
        self.cands = candidates
        self.idx = 0
        self.worked = None

    def parse(self, images, *, process_name="") -> list:
        """返回 [{item, spec, parameter}, ...]"""
        order = ([self.worked] if self.worked else []) + \
                self.cands[self.idx:] + self.cands[:self.idx]
        errs = []
        for i, ch in enumerate(order):
            try:
                items = _call_blueprint(ch, images, process_name)
                if items:
                    self.worked = ch
                    if self.worked not in self.cands[:self.idx + 1]:
                        self.idx = self.cands.index(ch)
                    return items
                # 空结果不重试同一渠道
                errs.append(f"{ch['name']}: 未识别到 QC 参数")
            except Exception as e:
                errs.append(str(e)[:120])
            self.idx = min(self.idx + i + 1, len(self.cands))
        raise DraftError("所有看图渠道都失败：\n- " + "\n- ".join(errs[-4:]))


# ---- 全参解析器（物料+BOM+设备+QC） ------------------------------------

def _call_full_parse(channel: dict, images: list, process_name: str, timeout=90) -> dict:
    """从工程图纸/BOM表/工艺卡中提取完整参数（物料+设备+QC）。

    返回 JSON 对象：{materials: [...], equipment: [...], qc_items: [...]}
    每一项包含该表的所有字段。
    """
    import requests
    text_prompt = (
        f"这是一份「{process_name}\"工序相关的工程图纸、BOM表或工艺卡。\n"
        "请全面识别其中的**所有参数信息**，按以下三类输出 JSON，不要遗漏：\n\n"
        '  "materials": [\n'
        '    {"code": "编号", "name": "名称", "spec": "规格型号", "qty": 数量(数字), "unit": "单位(如PCS/m等)"},\n'
        '  ],\n'
        '  "equipment": [\n'
        '    {"code": "型号/编号", "name": "设备/工夹具名称", "operating_time": "运行参数(如压力/温度/速度等)"},\n'
        '  ],\n'
        '  "qc_items": [\n'
        '    {"item": "检测项目名称", "spec": "标准/公差", "parameter": "测量方法/依据"},\n'
        '  ]\n\n'
        "注意：\n"
        "1. 数量 qty 必须是纯数字整数或小数\n"
        "2. 规格 spec 尽量保留原单位的精确值（如 ±0.05mm、≥50N、120°C）\n"
        "3. 如果某类在图纸中找不到就输出空数组 []\n"
        "4. 严格按此 JSON 结构输出，不要加任何文字解释"
    )
    parts = [{"type": "text", "text": text_prompt}]
    for img in images[:6]:
        parts.append({"type": "image_url",
                      "image_url": {"url": _image_data_url(img, max_side=1200, quality=85)}})
    url = urllib.parse.urljoin(channel["base_url"].rstrip("/") + "/", "chat/completions")
    body = {
        "model": channel["model"],
        "messages": [{"role": "user", "content": parts}],
        "max_tokens": 800,
        "response_format": {"type": "json_object"},
    }
    headers = {"Content-Type": "application/json"}
    if channel["key"]:
        headers["Authorization"] = f"Bearer {channel['key']}"
    r = requests.post(url, json=body, headers=headers, timeout=timeout)
    if r.status_code != 200:
        raise DraftError(f"{channel['name']}/{channel['model']} 返回 {r.status_code}："
                         f"{r.text[:160]}")
    try:
        content = r.json()["choices"][0]["message"]["content"]
        content = content.strip()
        if content.startswith("```"):
            content = content.split("```")[1]
            if content.startswith("json"):
                content = content[4:]
            content = content.rstrip("```").strip()
        result = json.loads(content)

        materials = []
        for row in (result.get("materials") or []):
            code = str(row.get("code", "")).strip()
            if code:
                try:
                    qty_v = float(row.get("qty", 1))
                except (ValueError, TypeError):
                    qty_v = 1
                materials.append({
                    "code": code,
                    "name": str(row.get("name", "")),
                    "spec": str(row.get("spec", "")),
                    "qty": qty_v,
                    "unit": str(row.get("unit", "PCS")),
                })

        equipment = []
        for row in (result.get("equipment") or []):
            code = str(row.get("code", "")).strip()
            if code or str(row.get("name", "")).strip():
                equipment.append({
                    "code": code,
                    "name": str(row.get("name", "")),
                    "operating_time": str(row.get("operating_time", "")),
                })

        qc_items = []
        for row in (result.get("qc_items") or []):
            item = str(row.get("item", "")).strip()
            if item:
                qc_items.append({
                    "item": item,
                    "spec": str(row.get("spec", "")),
                    "parameter": str(row.get("parameter", "")),
                })

        return {
            "materials": materials,
            "equipment": equipment,
            "qc_items": qc_items,
        }
    except Exception as e:
        raise DraftError(f"{channel['name']} 返回内容无法解析为 JSON：{str(e)[:160]}")
    raise DraftError("图纸参数解析失败")


class FullParseParser:
    """解析工程图纸完整参数（物料+设备+QC），某个渠道失败自动换下一个候选。

    返回 {materials: [...], equipment: [...], qc_items: [...]}
    """

    def __init__(self, candidates: list):
        self.cands = candidates
        self.idx = 0
        self.worked = None

    def parse(self, images, *, process_name="") -> dict:
        order = ([self.worked] if self.worked else []) + \
                self.cands[self.idx:] + self.cands[:self.idx]
        errs = []
        for i, ch in enumerate(order):
            try:
                result = _call_full_parse(ch, images, process_name)
                if any(len(v) > 0 for v in [result.get("materials", []), result.get("equipment", []), result.get("qc_items", [])]):
                    self.worked = ch
                    if self.worked not in self.cands[:self.idx + 1]:
                        self.idx = self.cands.index(ch)
                    return result
                errs.append(f"{ch['name']}: 未识别到任何参数")
            except Exception as e:
                errs.append(str(e)[:120])
            self.idx = min(self.idx + i + 1, len(self.cands))
        raise DraftError("所有看图渠道都失败：\n- " + "\n- ".join(errs[-4:]))


if __name__ == "__main__":
    imgs = sys.argv[1:]
    cands = vision_candidates(print)
    for c in cands[:8]:
        print("候选:", c["name"], c["model"])
    if imgs:
        d = StepDrafter(cands)
        print("起草:", d.draft(imgs, process_name="焊接", step_no=1))
