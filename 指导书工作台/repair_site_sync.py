# -*- coding: utf-8 -*-
"""安全修复 site_sync.json：剔除其中的不可见控制字符后写回。

用法（在 site_sync.json 所在目录执行）：
    python repair_site_sync.py            # 就地修复 ./site_sync.json
    python repair_site_sync.py 路径/xxx.json   # 指定文件

行为：
- 只剔除 0x00~0x1F 中除 \\n \\r \\t 外的控制字符（它们通常是复制粘贴带入的
  脏字符，会让 json.loads 直接报「同步配置文件读取失败」）；url/token 里本就
  不含这类字符，剔除是安全的。
- 不直接覆盖原文件，先写一份 `原名.repaired.json`，校验通过后再提示你手动替换，
  避免误伤。
"""
import json
import sys
from pathlib import Path

TARGET = sys.argv[1] if len(sys.argv) > 1 else "site_sync.json"


def main() -> int:
    p = Path(TARGET)
    if not p.exists():
        print(f"未找到文件：{p}")
        return 1
    raw = p.read_text(encoding="utf-8")
    cleaned = "".join(ch for ch in raw if ch >= " " or ch in "\n\r\t")
    if cleaned == raw:
        print("未发现需要剔除的控制字符，文件本身应是干净的。")
        return 0
    try:
        cfg = json.loads(cleaned)
    except Exception as e:
        print(f"剔除后仍无法解析：{e}")
        return 1
    if not cfg.get("url") or not cfg.get("token"):
        print("警告：修复后配置缺少 url 或 token，请手动核对。")
    out = p.with_name(p.name + ".repaired.json")
    out.write_text(cleaned, encoding="utf-8")
    print(f"已生成修复文件：{out}")
    print(f"请核对内容后，用它替换原 {p.name}（原文件未改动）。")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
