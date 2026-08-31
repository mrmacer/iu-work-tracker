# IU Work Tracker

A Next.js App Router application. See [docs/AI_HANDOFF.md](docs/AI_HANDOFF.md) for current project state and [docs/PRODUCT_VISION.md](docs/PRODUCT_VISION.md) for product context.

## Prerequisites

- Node.js `>=22.13.0`

## Quick Start

```bash
npm install
npm run dev
```

## Useful Commands

- `npm run dev`: start the Next.js development server
- `npm run build`: production build
- `npm start`: run the production build
- `npm test`: build, then run the automated test suite
- `npm run typecheck`: TypeScript check
- `npm run lint`: ESLint

## Persistence

SharePoint (via delegated Microsoft Graph access, `@azure/msal-browser`) is the durable production data store when a Microsoft account is configured and signed in. Without an authenticated Microsoft connection, the app runs against a non-durable in-memory store for local development, testing, and preview — see `docs/AI_HANDOFF.md` for details.
