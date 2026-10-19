# -*- coding: utf-8 -*-
"""免费大模型自动探测 —— 核心探测逻辑（纯函数，可被其他项目直接 import 复用）。

用法（命令行/其他项目）:
    from probe import probe_all, load_providers, load_keys, pick_best
    results = probe_all(load_providers(), load_keys())
    best = pick_best(results, family="vision")   # 或 "chat"
本文件不依赖 tkinter，只做网络探测和判定。
"""

import json
import sys
import time
import urllib.parse
from pathlib import Path

if getattr(sys, "frozen", False):  # PyInstaller 打包后：数据文件跟着 exe 走
    HERE = Path(sys.executable).resolve().parent
    _BUNDLED = Path(getattr(sys, "_MEIPASS", HERE))
else:
    HERE = Path(__file__).resolve().parent
    _BUNDLED = HERE

PROVIDERS_FILE = HERE / "providers.json"
if not PROVIDERS_FILE.exists() and (_BUNDLED / "providers.json").exists():
    PROVIDERS_FILE = _BUNDLED / "providers.json"
KEYS_FILE = HERE / "keys.json"

DEFAULT_TIMEOUT = 9  # 每个供应商单独超时，秒

STATUS_OK_NOKEY = "ok_nokey"      # ✅ 无需账号即可用
STATUS_OK_KEY = "ok_key"          # 🔑 已关联可用
STATUS_NEED_KEY = "need_key"      # ⚠️ 需注册领key
STATUS_UNREACHABLE = "unreachable"  # ❌ 当前网络不可达

STATUS_LABELS = {
    STATUS_OK_NOKEY: "✅ 无需账号即可用",
    STATUS_OK_KEY: "🔑 已关联可用",
    STATUS_NEED_KEY: "⚠️ 需注册领key",
    STATUS_UNREACHABLE: "❌ 当前网络不可达",
}

# 判断"视觉模型"的常见命名片段（各家命名习惯）
VISION_HINTS = ("vision", "-vl", "vl-", "4v", "4.5-v", "ocr", "image",
                "llava", "minicpm-v", "qwen-vl", "glm-4v", "internvl",
                "cambrian", "moondream", "janus")


def load_providers(path: Path = PROVIDERS_FILE) -> list:
    data = json.loads(Path(path).read_text(encoding="utf-8"))
    return [p for p in data.get("providers", []) if p.get("enabled", True)]


def load_keys(path: Path = KEYS_FILE) -> dict:
    p = Path(path)
    if not p.exists():
        return {}
    try:
        return json.loads(p.read_text(encoding="utf-8"))
    except Exception:
        return {}


def is_vision_model(name: str) -> bool:
    n = name.lower()
    return any(h in n for h in VISION_HINTS)


def free_models_of(models: list, rule: str | None) -> list:
    """按 providers.json 的 free_rule 挑出免费模型。
    all=清单里都是免费额度内 | suffix_free=带 :free 后缀 | contains:a,b=名字含关键词
    | official:名1,名2=人工在官方定价页核实过的免费名单逐字精确匹配（最硬口径）
    | 其他=判不出"""
    rule = (rule or "").strip()
    if rule == "all":
        return list(models)
    if rule == "suffix_free":
        return [m for m in models if ":free" in m.lower()]
    if rule.startswith("official:"):
        off = {s.strip().lower() for s in rule[9:].split(",") if s.strip()}
        return [m for m in models if m.lower() in off]
    if rule.startswith("contains:"):
        subs = [s.strip().lower() for s in rule[9:].split(",") if s.strip()]
        return [m for m in models if any(s in m.lower() for s in subs)]
    return []


def _headers(key: str) -> dict:
    h = {"Accept": "application/json"}
    if key:
        h["Authorization"] = f"Bearer {key}"
    return h


def _post_chat(session, base_url: str, model: str, key: str, timeout: float):
    url = urllib.parse.urljoin(base_url.rstrip("/") + "/", "chat/completions")
    body = {
        "model": model,
        "messages": [{"role": "user", "content": "hi"}],
        "max_tokens": 1,
    }
    return session.post(url, json=body, headers=_headers(key), timeout=timeout)


def _probe_one(session, provider: dict, key: str, timeout: float) -> dict:
    """探测单个供应商，返回一条结果记录。任何异常都吞掉并归类，绝不抛出。"""
    base = provider["base_url"].rstrip("/") + "/"
    pid = provider["id"]
    has_key = bool(key)
    result = {
        "id": pid,
        "name": provider.get("name", pid),
        "status": STATUS_UNREACHABLE,
        "latency_ms": None,
        "model_count": 0,
        "models": [],
        "sample_model": "",
        "notes": provider.get("notes", ""),
        "promo": provider.get("promo", ""),
        "free_models": [],
        "need_proxy": provider.get("need_proxy", False),
        "register_url": provider.get("register_url", ""),
        "has_key": has_key,
        "detail": "",
    }
    t0 = time.monotonic()
    try:
        # 第一步：OpenAI 兼容的模型清单接口
        resp = session.get(base + "models", headers=_headers(key), timeout=timeout)
        result["latency_ms"] = int((time.monotonic() - t0) * 1000)
        if resp.status_code in (200, 201, 204):
            try:
                data = resp.json()
                items = data.get("data") or data.get("models") or data.get("items") or []
                names = sorted({m.get("id") or m.get("name") for m in items
                                if isinstance(m, dict) and (m.get("id") or m.get("name"))})
            except Exception:
                names = []
            result["models"] = names
            result["model_count"] = len(names)
            result["free_models"] = free_models_of(names, provider.get("free_rule"))
            if names:
                result["sample_model"] = names[0]
            result["status"] = STATUS_OK_KEY if has_key else STATUS_OK_NOKEY
            result["detail"] = f"GET /models 返回 {resp.status_code}"
            # 没带 key 却列出了清单：再发一条试聊，确认真能免费对话，
            # 防止把"只公开名单、调用仍要 key"的渠道误标成 ✅
            if not has_key and provider.get("needs_key", True):
                check = provider.get("probe_model") or (names[0] if names else "")
                if check:
                    try:
                        resp3 = _post_chat(session, base, check, "", timeout)
                        if resp3.status_code in (401, 403):
                            result["status"] = STATUS_NEED_KEY
                            result["detail"] = ("模型清单公开，但真实对话仍需 key"
                                                f"（试聊 {check} 返回 {resp3.status_code}）")
                        elif resp3.status_code in (200, 201, 204):
                            result["detail"] = f"GET /models 200，且免 key 试聊 {check} 成功"
                    except Exception:
                        pass
            return result
        if resp.status_code in (401, 403):
            result["status"] = STATUS_NEED_KEY
            result["detail"] = f"GET /models 返回 {resp.status_code}（需要有效 key）"
            return result
        if resp.status_code == 404:
            result["detail"] = "GET /models 返回 404，改用试聊探测"
        else:
            result["detail"] = f"GET /models 返回 {resp.status_code}，改用试聊探测"

        # 第二步：退化为一条极短试聊（max_tokens=1）
        probe_model = provider.get("probe_model") or ""
        if not probe_model:
            probe_model = result["models"][0] if result["models"] else "default"
        if not has_key and provider.get("needs_key", True):
            # 明确要 key 且用户还没填，不必浪费一次试聊
            result["status"] = STATUS_NEED_KEY
            result["detail"] += "；该渠道需要 key，尚未关联"
            return result
        t1 = time.monotonic()
        resp2 = _post_chat(session, base, probe_model, key, timeout)
        if result["latency_ms"] is None or resp2.status_code in (200, 201):
            result["latency_ms"] = int((time.monotonic() - t1) * 1000)
        if resp2.status_code in (200, 201, 204):
            result["status"] = STATUS_OK_KEY if has_key else STATUS_OK_NOKEY
            result["sample_model"] = probe_model
            if result["model_count"] == 0:
                result["model_count"] = 1
                result["models"] = [probe_model]
                result["free_models"] = free_models_of([probe_model], provider.get("free_rule"))
            result["detail"] = f"试聊 {probe_model} 成功"
        elif resp2.status_code in (401, 403):
            result["status"] = STATUS_NEED_KEY
            result["detail"] = f"试聊返回 {resp2.status_code}（key 无效或未提供）"
        else:
            result["status"] = STATUS_NEED_KEY if resp2.status_code in (
                400, 404, 422, 429) else STATUS_UNREACHABLE
            result["detail"] = f"试聊返回 {resp2.status_code}"
    except Exception as e:
        result["latency_ms"] = int((time.monotonic() - t0) * 1000)
        result["status"] = STATUS_UNREACHABLE
        cls = type(e).__name__
        if "Timeout" in cls:
            result["detail"] = f"连接超时（{timeout:.0f} 秒）"
        elif "Connection" in cls or "resolve" in str(e).lower() or "Failed to establish" in str(e):
            result["detail"] = "连接被拒绝或域名解析失败"
        else:
            result["detail"] = f"{cls}: {e}"
    return result


def probe_one_provider(provider: dict, key: str = "", timeout: float = DEFAULT_TIMEOUT) -> dict:
    import requests
    with requests.Session() as s:
        return _probe_one(s, provider, key, timeout)


def probe_all(providers: list, keys: dict | None = None,
              timeout: float = DEFAULT_TIMEOUT, max_workers: int = 4,
              on_result=None, total_budget: float = 75) -> list:
    """并发探测全部供应商。on_result(result_dict, done, total) 可用于进度回调。
    每个供应商一个守护线程；无论谁卡死，本函数最多 total_budget 秒后返回，
    没跑完的渠道按"超时不可达"记——保证界面永远不会卡在探测中。"""
    import queue as _queue
    import threading
    import requests
    keys = keys or {}
    total = len(providers)
    q: _queue.Queue = _queue.Queue()
    seen: dict = {}

    def run(p):
        try:
            with requests.Session() as s:
                r = _probe_one(s, p, keys.get(p["id"], ""), timeout)
        except Exception as e:
            r = _base_result(p, keys)
            r["detail"] = f"探测线程异常：{type(e).__name__}: {e}"
        q.put(r)

    for p in providers:
        threading.Thread(target=run, args=(p,), daemon=True).start()

    deadline = time.monotonic() + total_budget
    done = 0
    while done < total:
        left = deadline - time.monotonic()
        if left <= 0:
            break
        try:
            r = q.get(timeout=min(left, 1.0))
        except _queue.Empty:
            continue
        if r["id"] in seen:
            continue
        seen[r["id"]] = r
        done += 1
        if on_result:
            try:
                on_result(r, done, total)
            except Exception:
                pass
    for p in providers:
        if p["id"] not in seen:
            r = _base_result(p, keys)
            r["latency_ms"] = int(total_budget * 1000)
            r["detail"] = f"超过 {total_budget:.0f} 秒仍未响应，按不可达处理（多为海外源未挂代理）"
            seen[p["id"]] = r
            done += 1
            if on_result:
                try:
                    on_result(r, done, total)
                except Exception:
                    pass
    order = {p["id"]: i for i, p in enumerate(providers)}
    ordered = list(seen.values())
    if ordered:
        ordered.sort(key=lambda r: (
            [STATUS_OK_KEY, STATUS_OK_NOKEY, STATUS_NEED_KEY,
             STATUS_UNREACHABLE].index(r["status"]) if r["status"] in
            (STATUS_OK_KEY, STATUS_OK_NOKEY, STATUS_NEED_KEY, STATUS_UNREACHABLE) else 9,
            order.get(r["id"], 99)))
    return ordered


def _base_result(provider: dict, keys: dict) -> dict:
    return {
        "id": provider["id"],
        "name": provider.get("name", provider["id"]),
        "status": STATUS_UNREACHABLE,
        "latency_ms": None,
        "model_count": 0,
        "models": [],
        "sample_model": "",
        "notes": provider.get("notes", ""),
        "promo": provider.get("promo", ""),
        "free_models": [],
        "need_proxy": provider.get("need_proxy", False),
        "register_url": provider.get("register_url", ""),
        "has_key": bool(keys.get(provider["id"])),
        "detail": "",
    }


def pick_best(results: list, family: str = "chat") -> dict | None:
    """从探测结果里挑一个最适合用的免费渠道。
    family="vision" 要求能处理图片的模型；family="chat" 普通对话。
    优先级：已关联key > 无需key可用；本地源排最后（避免占用本机资源做首选）。"""
    usable = [r for r in results if r["status"] in (STATUS_OK_KEY, STATUS_OK_NOKEY)]
    if family == "vision":
        def names(r):
            return [m for m in r["models"] if is_vision_model(m)]
    else:
        def names(r):
            return [m for m in r["models"] if not is_vision_model(m)] or r["models"]
    usable = [r for r in usable if names(r)]
    if not usable:
        return None
    usable.sort(key=lambda r: (
        0 if not r["id"].startswith(("ollama", "lmstudio")) else 1,
        0 if r["status"] == STATUS_OK_KEY else 1,
        r["latency_ms"] if r["latency_ms"] is not None else 10 ** 9,
    ))
    best = usable[0]
    out = dict(best)
    out["picked_model"] = names(best)[0]
    return out


if __name__ == "__main__":
    rs = probe_all(load_providers(), load_keys())
    for r in rs:
        print(f"{STATUS_LABELS[r['status']]:<12} {r['name']:<24} "
              f"{str(r['latency_ms']) + 'ms' if r['latency_ms'] is not None else '-':>8} "
              f"模型数:{r['model_count']:<4} 免费:{len(r.get('free_models', [])):<4} {r['sample_model']}")
    print("\n最佳 chat 渠道:", (pick_best(rs, 'chat') or {}).get("name", "无"))
    print("最佳 vision 渠道:", (pick_best(rs, 'vision') or {}).get("name", "无"))
