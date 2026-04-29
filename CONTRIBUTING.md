# Contributing

Thanks for your interest in contributing to spo-link-inventory.

## Development environment

You need:

- Node.js 18 – 22
- Bash (git-bash on Windows works)
- Azure CLI 2.50+
- Azure Functions Core Tools v4
- (For SPFx) `npm install` inside `spfx/link-inventory-admin/` — uses Heft toolchain, separate from `func/`

The two packages are **standalone** — no workspaces, no shared dependency. The URL normalizer logic that's used by both is small enough that we duplicated it (in `func/src/services/urlNormalizer.ts` and `spfx/.../components/linkSuggestionClient.ts`) rather than introducing a build-time dependency that breaks Linux Consumption deploys.

## Local dev loop

For the function:

```bash
cd func
npm install
npm run build
npm test          # 100+ tests, all hermetic
npm start         # local dev (Azurite required)
```

For the SPFx web part:

```bash
cd spfx/link-inventory-admin
npm install
npm run build:dev   # full TypeScript + sass + webpack pass
npm run start       # workbench on https://localhost:4321
```

## Adding a new function endpoint

The Azure Functions v4 programming model has a sharp edge: every `functions/<name>.ts` file MUST be imported in `src/index.ts`. Adding a new function file without the matching import results in the endpoint silently missing at runtime — TypeScript builds clean, deploys succeed, but `/api/<your-route>` returns 404. Always update `src/index.ts` when you add a function file.

When the new endpoint reads an environment variable at module-init time (a `const X = process.env.FOO` at the top level of a service file), add a default for it in `func/src/__tests__/test-setup.ts` or your tests will fail at module import.

## Style & shape

- **No tenant-specific values in source.** Anything tenant-specific (hostnames, GUIDs, group names) goes through `func/src/services/config.ts` (server side) or web part properties (client side).
- **OBO is the only auth path.** All endpoints require a user bearer token via `AadHttpClient`. There is no function-key fallback in this codebase.
- **Idempotent storage operations.** Writers use ETag-optimistic concurrency or upsert-by-key. Don't introduce read-modify-write without thinking about concurrent runs.
- **Comments explain WHY, not WHAT.** Most workarounds in this codebase are SharePoint quirks (URL encoding rules, list view threshold, classic-site detection). When you fix one, leave a comment describing the SP behavior so the next person doesn't undo it.

## Testing

- Unit tests in `func/src/__tests__/` use the Node `node --test` runner. No Jest, no Mocha. Tests are loaded via a `--import` setup file that pre-populates env vars consumed at module-init time.
- Integration tests against a real tenant: not committed. Test scripts live under `.tmp-*` patterns and are gitignored — run them ad-hoc.
- SPFx tests: not currently set up; PRs welcome.

## Pull requests

1. Open an issue first for anything non-trivial — easier to align on scope.
2. Branch from `main`. No feature branches required for tiny changes; for anything bigger, name the branch by topic.
3. CI runs `func` build+test on every PR and `spfx` build:dev on every PR. Both must pass.
4. Squash on merge.

## Deploy after a change

If you change func/ code and want to redeploy your dev tenant:

```bash
cd func
npm run build
func azure functionapp publish $FUNCTION_APP_NAME --typescript --build remote
```

Settings + role assignments + the federated credential persist — don't need to redo Bicep / setup-entra unless infra itself changes.

## Reporting security issues

Don't open a public issue for security findings. Email `<hi at charlie dot tools>` with details.  I'll respond within a few days.
