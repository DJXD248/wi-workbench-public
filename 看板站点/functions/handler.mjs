const json = (body, status = 200, extra = {}) => Response.json(body, {
  status,
  headers: { "cache-control": "no-store", ...extra },
});

// 作业指导书快照读写：GET/POST .../app/snapshot
// POST 校验 x-sync-token（应用 Secret WI_SYNC_TOKEN），快照存 Storage wi-dashboard/snapshot.json
const SNAPSHOT_PATH = "wi-dashboard/snapshot.json";
const MAX_BYTES = 512 * 1024;

export async function handleDashboard({ request, env = () => undefined, storage = null }) {
  const url = new URL(request.url);
  const segs = url.pathname.split("/").filter(Boolean);
  if ((segs[segs.length - 1] || "") === "snapshot" || url.searchParams.get("action") === "snapshot") {
    if (!storage) return json({ ok: false, error: "storage_unavailable" }, 503);
    if (request.method === "GET") return getSnapshot(storage);
    if (request.method === "POST") return postSnapshot(request, storage, env);
    return json({ ok: false, error: "method_not_allowed" }, 405, { allow: "GET, POST" });
  }

  return json({ error: "not_found" }, 404);
}

async function getSnapshot(storage) {
  try {
    const res = await storage.download(SNAPSHOT_PATH);
    if (!res.ok) return json({ ok: true, empty: true });
    const data = JSON.parse(await res.text());
    if (!isSnapshot(data)) return json({ ok: true, empty: true });
    return json({ ok: true, empty: false, snapshot: data });
  } catch {
    // 尚无快照或读取失败都按"还没有数据"处理，页面显示空状态
    return json({ ok: true, empty: true });
  }
}

async function postSnapshot(request, storage, env) {
  const expected = env("WI_SYNC_TOKEN");
  if (!expected) return json({ ok: false, error: "service_not_configured" }, 503);
  const got = request.headers.get("x-sync-token") || "";
  if (got !== expected) return json({ ok: false, error: "invalid_token" }, 401);

  const text = await readLimitedText(request, MAX_BYTES);
  if (text === null) return json({ ok: false, error: "too_large" }, 413);
  let data;
  try {
    data = JSON.parse(text);
  } catch {
    return json({ ok: false, error: "invalid_json" }, 400);
  }
  if (!isSnapshot(data)) return json({ ok: false, error: "invalid_schema" }, 400);

  try {
    await storage.upload(SNAPSHOT_PATH, new TextEncoder().encode(text), {
      contentType: "application/json",
      upsert: true,
    });
    const back = await storage.download(SNAPSHOT_PATH);
    if (!back.ok || (await back.text()).length !== text.length) {
      return json({ ok: false, error: "save_verify_failed" }, 500);
    }
  } catch {
    return json({ ok: false, error: "storage_unavailable" }, 503);
  }
  return json({ ok: true, saved_bytes: text.length });
}

function isSnapshot(d) {
  return (
    d && typeof d === "object" && d.schema === 1 &&
    typeof d.synced_at === "string" && Array.isArray(d.docs) &&
    d.docs.length <= 200 && d.docs.every(
      (x) => x && typeof x === "object" && typeof x.file === "string" && x.file.length <= 200) &&
    typeof d.summary === "object" && d.summary !== null
  );
}

// 声明长度只用于提前拒绝，实际以流式读取的字节数为准
async function readLimitedText(request, maxBytes) {
  const declared = Number(request.headers.get("content-length") || "0");
  if (declared > maxBytes) return null;
  const reader = request.body?.getReader();
  if (!reader) return null;
  const chunks = [];
  let received = 0;
  for (;;) {
    const { done, value } = await reader.read();
    if (done) break;
    received += value.byteLength;
    if (received > maxBytes) {
      await reader.cancel();
      return null;
    }
    chunks.push(value);
  }
  const buf = new Uint8Array(received);
  let off = 0;
  for (const c of chunks) {
    buf.set(c, off);
    off += c.byteLength;
  }
  return new TextDecoder().decode(buf);
}
