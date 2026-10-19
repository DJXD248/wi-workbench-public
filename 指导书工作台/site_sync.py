# -*- coding: utf-8 -*-
"""把本地 configs/output 的工作进度打包成快照，上传到 Qoder 站点看板。

同目录需要 site_sync.json：{"url": "...", "token": "..."}
"""
from __future__ import annotations

import json
import socket
from datetime import datetime
from pathlib import Path
from urllib.error import HTTPError, URLError
from urllib.parse import urlparse
from urllib.request import Request, urlopen

import yaml

CONFIG_FILE = "site_sync.json"


class SyncError(Exception):
    pass


def load_config(app_dir) -> dict:
    p = Path(app_dir) / CONFIG_FILE
    if not p.exists():
        raise SyncError(f"未找到同步配置文件：{p}\n（应由维护人员放在程序同目录）")
    try:
        cfg = json.loads(p.read_text(encoding="utf-8"))
    except Exception as e:
        raise SyncError(f"同步配置文件读取失败：{e}")
    if not cfg.get("url") or not cfg.get("token"):
        raise SyncError("同步配置缺少 url 或 token")
    return cfg


def _int(v, default=0):
    try:
        return int(float(v))
    except (TypeError, ValueError):
        return default


def _ts(p: Path) -> str:
    return datetime.fromtimestamp(p.stat().st_mtime).strftime("%Y-%m-%d %H:%M")


def collect_snapshot(configs_dir, output_dir) -> dict:
    docs = []
    configs_dir = Path(configs_dir)
    for f in sorted(configs_dir.rglob("*.yaml")):
        try:
            doc = yaml.safe_load(f.read_text(encoding="utf-8")) or {}
        except Exception:
            continue
        if not isinstance(doc, dict):
            continue
        hdr = doc.get("header") or {}
        procs = doc.get("processes") or []
        work = sum(_int(p.get("work_hours")) for p in procs if isinstance(p, dict))
        rel = f.relative_to(configs_dir)
        series = rel.parts[0] if len(rel.parts) > 1 else rel.stem
        docs.append({
            "series": str(series),
            "file": str(rel),
            "file_no": str(hdr.get("file_no") or "")[:120],
            "edition": str(hdr.get("edition") or "")[:20],
            "product": str(hdr.get("product_spec_cn") or hdr.get("product_spec") or "")[:200],
            "processes": len([p for p in procs if isinstance(p, dict)]),
            "work_hours": work,
            "updated": _ts(f),
        })
    docs = docs[:200]

    recent = []
    output_dir = Path(output_dir)
    if output_dir.exists():
        xls = [f for f in output_dir.glob("*.xlsx") if not f.name.startswith("~$")]
        for f in sorted(xls, key=lambda x: x.stat().st_mtime, reverse=True)[:8]:
            recent.append({
                "name": f.name[:120],
                "time": _ts(f),
                "size_kb": max(1, round(f.stat().st_size / 1024)),
            })

    return {
        "schema": 1,
        "synced_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "source_pc": socket.gethostname()[:60],
        "summary": {
            "series": len({d["series"] for d in docs}),
            "docs": len(docs),
            "processes": sum(d["processes"] for d in docs),
            "work_hours": sum(d["work_hours"] for d in docs),
        },
        "docs": docs,
        "recent_output": recent,
    }


def _http_hint(code: int) -> str:
    if code in (401, 403):
        return f"同步口令校验未通过（{code}），请联系维护人员核对口令"
    if code == 404:
        return "没找到站点接口（404），请确认站点已发布"
    if code == 503:
        return "站点服务未就绪（503），请稍后重试或让维护人员检查配置"
    return f"上传失败，站点返回 HTTP {code}"


def upload(snapshot: dict, cfg: dict, timeout: int = 25) -> None:
    url = cfg["url"]
    origin = urlparse(url)
    req = Request(url, data=json.dumps(snapshot, ensure_ascii=False).encode("utf-8"),
                  method="POST")
    req.add_header("Content-Type", "application/json")
    req.add_header("x-sync-token", cfg["token"])
    # 网关对写操作校验同源，非浏览器程序需自带站点 Origin 头
    req.add_header("Origin", f"{origin.scheme}://{origin.netloc}")
    try:
        with urlopen(req, timeout=timeout) as resp:
            body = resp.read().decode("utf-8", "replace")
    except HTTPError as e:
        raise SyncError(_http_hint(e.code))
    except (URLError, TimeoutError, OSError) as e:
        raise SyncError(f"网络不通或站点无响应：{getattr(e, 'reason', e)}")
    if body.lstrip().startswith("<"):
        raise SyncError("站点返回了登录页，说明访问权限还没放开，请联系维护人员检查站点设置")
    try:
        out = json.loads(body)
    except Exception:
        raise SyncError("站点返回了无法解析的内容")
    if not out.get("ok"):
        raise SyncError({"invalid_token": "同步口令不对，请联系维护人员更新 site_sync.json",
                        "service_not_configured": "站点还没配置同步口令，请联系维护人员",
                        "invalid_schema": "快照格式不符，请更新工作台程序后重试",
                        "too_large": "数据量超出限制",
                        }.get(out.get("error"), f"站点保存失败（{out.get('error')}）"))
