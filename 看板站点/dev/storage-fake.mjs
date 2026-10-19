// 本地开发用的无凭证存储替身，仅存在于 dev/，不会进入发布包
export function createFakeStorage() {
  const store = new Map();
  return {
    async upload(objectPath, bytes, { upsert = false } = {}) {
      if (!upsert && store.has(objectPath)) {
        throw new Error("object_exists");
      }
      const text = new TextDecoder().decode(bytes);
      store.set(objectPath, text);
      return { Key: objectPath, Size: bytes.byteLength };
    },
    async download(objectPath) {
      if (!store.has(objectPath)) {
        return new Response("Not Found", { status: 404 });
      }
      return new Response(store.get(objectPath), {
        status: 200,
        headers: { "content-type": "application/json" },
      });
    },
  };
}
