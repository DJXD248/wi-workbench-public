// 本地函数服务入口（需安装 Deno；本机未装，验证用 dev/test-handler.mjs 代替）
import { handler } from "../functions/handler.mjs";
import { createFakeStorage } from "./storage-fake.mjs";

const storage = createFakeStorage();
Deno.serve({ hostname: "127.0.0.1", port: 8000 },
  (request) => handler({ request, storage, env: (n) => (n === "WI_SYNC_TOKEN" ? "test-token" : undefined) }));
