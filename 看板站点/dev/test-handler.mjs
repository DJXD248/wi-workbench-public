// 本地验证：用存储替身跑通 POST 上传 -> GET 读取 -> 错误分支
import { handler } from "../functions/handler.mjs";
import { createFakeStorage } from "./storage-fake.mjs";

const storage = createFakeStorage();
const env = (n) => (n === "WI_SYNC_TOKEN" ? "test-token" : undefined);
const call = (req) => handler({ request: req, storage, env });

const snapshot = {
  schema: 1,
  synced_at: "2026-10-02 21:00:00",
  source_pc: "家里电脑",
  summary: { series: 2, docs: 3, processes: 12, work_hours: 640 },
  docs: [
    { series: "DEMO_series", file: "DEMO_demo.yaml", file_no: "WI-DEMO-W-0183-2024", edition: "A",
      product: "演示系列组件", processes: 5, work_hours: 300, updated: "2026-09-30 14:02" },
  ],
  recent_output: [{ name: "DEMOB正式版.xlsx", time: "2026-10-01 10:11", size_kb: 88 }],
};

const assert = (cond, label) => {
  console.log((cond ? "PASS" : "FAIL") + "  " + label);
  if (!cond) process.exitCode = 1;
};

let r = await call(new Request("http://x/functions/v1/app/snapshot", { method: "GET" }));
let b = await r.json();
assert(r.status === 200 && b.ok && b.empty, "空看板 GET 返回 empty");

r = await call(new Request("http://x/functions/v1/app/snapshot", {
  method: "POST", headers: { "content-type": "application/json" }, body: JSON.stringify(snapshot),
}));
assert(r.status === 401, "缺 token 的 POST 被拒");

r = await call(new Request("http://x/functions/v1/app/snapshot", {
  method: "POST", headers: { "content-type": "application/json", "x-sync-token": "wrong" }, body: JSON.stringify(snapshot),
}));
assert(r.status === 401, "错 token 的 POST 被拒");

r = await call(new Request("http://x/functions/v1/app/snapshot", {
  method: "POST", headers: { "content-type": "application/json", "x-sync-token": "test-token" }, body: JSON.stringify(snapshot),
}));
b = await r.json();
assert(r.status === 200 && b.ok && b.saved_bytes > 0, "带 token 的 POST 保存成功");

r = await call(new Request("http://x/functions/v1/app/snapshot", { method: "GET" }));
b = await r.json();
assert(b.ok && !b.empty && b.snapshot.docs[0].file === "DEMO_demo.yaml", "GET 读回刚上传的快照");

r = await call(new Request("http://x/functions/v1/app/snapshot", {
  method: "POST", headers: { "x-sync-token": "test-token" }, body: "{坏json",
}));
assert(r.status === 400, "坏 JSON 被拒");

r = await call(new Request("http://x/functions/v1/app/snapshot", {
  method: "POST", headers: { "x-sync-token": "test-token" }, body: JSON.stringify({ schema: 2, docs: [], summary: {}, synced_at: "x" }),
}));
assert(r.status === 400, "schema 不符被拒");

r = await call(new Request("http://x/functions/v1/app/unknown", { method: "GET" }));
assert(r.status === 404, "未知路径不返回成功");

const envNoSecret = (n) => (n === "WI_SYNC_TOKEN" ? undefined : "");
r = await handler({ request: new Request("http://x/functions/v1/app/snapshot", { method: "POST" }), storage, env: envNoSecret });
assert(r.status === 503, "Secret 未配置时明确报 service_not_configured");

console.log(process.exitCode ? "存在失败用例" : "全部通过");
