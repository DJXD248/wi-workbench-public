import { createClient } from "npm:@supabase/supabase-js@2.57.4";
// _qoder/storage.mjs 由平台在部署时注入，不要打包该目录
import { storage } from "./_qoder/storage.mjs";
import { serveSite } from "./adapter.mjs";
import { handleDashboard } from "./handler.mjs";

const getEnv = (name: string) => Deno.env.get(name);

Deno.serve(serveSite((ctx) => handleDashboard({ ...ctx, storage, env: getEnv }), {
  createClient,
  env: getEnv,
}));
