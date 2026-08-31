import assert from "node:assert/strict";
import test from "node:test";
import { spawn } from "node:child_process";
import path from "node:path";
import { fileURLToPath } from "node:url";

// Replaces the pre-Patch-3A version of this test, which imported the Cloudflare Worker
// build output (dist/server/index.js) and called its .fetch(request, env, ctx) directly —
// that build no longer exists under real Next.js. This preserves the same intent (prove the
// actual production server renders the app shell correctly, not a mocked/unit-level render)
// by starting the real `next start` production server and making a real HTTP request against
// it. Assumes `npm run build` (next build) already produced .next/, which npm test's script
// chain guarantees by running build before this test.

const projectRoot = path.dirname(path.dirname(fileURLToPath(import.meta.url)));
const PORT = 39217;

async function waitForServer(url, timeoutMs = 30000) {
  const start = Date.now();
  let lastError;
  while (Date.now() - start < timeoutMs) {
    try {
      return await fetch(url);
    } catch (error) {
      lastError = error;
      await new Promise((resolve) => setTimeout(resolve, 200));
    }
  }
  throw new Error(`Server at ${url} did not respond within ${timeoutMs}ms: ${lastError}`);
}

test("server-renders the IU Work Tracker shell", async () => {
  const server = spawn("npx", ["next", "start", "-p", String(PORT)], {
    cwd: projectRoot,
    stdio: ["ignore", "pipe", "pipe"],
  });
  server.on("error", () => {});

  try {
    const response = await waitForServer(`http://localhost:${PORT}/`);
    assert.equal(response.status, 200);
    const html = await response.text();
    assert.match(html, /IU Work Tracker/);
    assert.match(html, /Log work/i);
    assert.match(html, /Preparing your workspace/);
    assert.match(html, /STEM \/ ORBIT/);
    assert.doesNotMatch(html, /codex-preview/);
  } finally {
    server.kill("SIGTERM");
  }
});
