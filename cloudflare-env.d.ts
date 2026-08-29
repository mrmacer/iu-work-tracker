declare namespace Cloudflare {
  interface Env {
    DB: D1Database;
    /** Server-only secret for Inbox Intelligence. Never NEXT_PUBLIC_*; never sent to the browser. */
    ANTHROPIC_API_KEY?: string;
  }
}
