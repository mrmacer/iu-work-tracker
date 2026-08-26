import assert from "node:assert/strict";
import test from "node:test";

test("server-renders the IU Work Tracker shell", async () => {
  const workerUrl = new URL("../dist/server/index.js", import.meta.url);
  workerUrl.searchParams.set("test", `${process.pid}-${Date.now()}`);
  const { default: worker } = await import(workerUrl.href);
  const response = await worker.fetch(new Request("http://localhost/", { headers: { accept: "text/html" } }), { ASSETS: { fetch: async () => new Response("Not found", { status:404 }) }, DB: undefined }, { waitUntil(){}, passThroughOnException(){} });
  assert.equal(response.status, 200);
  const html = await response.text();
  assert.match(html, /IU Work Tracker/);
  assert.match(html, /Log work/i);
  assert.match(html, /Preparing your workspace/);
  assert.match(html, /STEM \/ ORBIT/);
  assert.doesNotMatch(html, /codex-preview/);
});
