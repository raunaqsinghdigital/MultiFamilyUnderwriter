# Changelog

All notable changes to this project are documented here.
Format follows [Keep a Changelog](https://keepachangelog.com/en/1.1.0/).

---

## [Unreleased]

---

## [1.1.0] — 2026-03-03

### Added
- **Public hosting config** — `render.yaml` and `requirements.txt` for zero-config deployment to Render (free tier)
- **Supabase Auth integration** — email/password login via `@supabase/supabase-js@2` (CDN); JWT-based sessions with no extra Python packages (HS256 verified using stdlib `hmac` + `hashlib`)
- **RBAC** — `admin` and `analyst` roles stored in Supabase `app_metadata.role`; read from `raw_app_meta_data` in the JWT payload
- **CORS origin whitelisting** — `ALLOWED_ORIGINS` env var (comma-separated); cross-origin requests from unlisted domains receive `403`; same-origin and direct requests always allowed
- **OPTIONS preflight handler** — `do_OPTIONS()` on `UnderwriterHandler` for browser CORS preflight support
- **`UnderwriterServer` class** — subclass of `ThreadingHTTPServer` that carries `allowed_origins` on the server object
- **Login overlay UI** — full-screen login card shown until a valid Supabase session is established; handles sign-in errors inline
- **Product selector nav** — horizontal pill-tab nav above the toolbar; Buy & Hold Underwriter (active) + Value Add Underwriter (Coming Soon / disabled)
- **User chip in toolbar** — shows signed-in email, role badge (Admin/Analyst), and Sign Out button
- **Session expiry handling** — any `401` API response re-shows the login overlay with "Session expired" message
- **`HOST` / `PORT` env var support** — server now reads bind address and port from environment (required for Render)
- **Startup warnings** — server logs a warning if `SUPABASE_JWT_SECRET` is not set

### Changed
- **Admin mode gate** — replaced insecure `?admin=1` URL parameter with JWT role claim (`isAdminMode` is now a function `isAdminMode()` checking `currentUser.role === "admin"`)
- **Default bind address** — changed from `127.0.0.1` (localhost only) to `0.0.0.0` (all interfaces) to support public hosting
- **All API fetch calls** — now include `Authorization: Bearer <token>` header via `authHeaders()` helper
- **App shell** — `<main>` is hidden by default and only shown after successful authentication
- **Toolbar eyebrow text** — updated from "Local Analysis Workspace" to "Buy & Hold Analysis Workspace"
- **README** — completely rewritten with deployment guide, Supabase setup, RBAC reference, and troubleshooting section

### Security
- All `/api/*` endpoints now require a valid Supabase Bearer token (return `401` if missing or invalid, `401` if expired)
- `/api/admin/*` endpoints additionally require `role=admin` in JWT `raw_app_meta_data` (return `403` otherwise)
- Static file endpoints remain public (required for the login page to load)
- Path traversal protection unchanged (existing)

---

## [1.0.0] — 2025

### Added
- Initial local-only multifamily underwriting tool
- Pure Python HTTP server using stdlib `ThreadingHTTPServer` (no framework, no dependencies)
- Workbook model loaded from `data/workbook_model.json` (740 KB) at startup
- Formula engine: tokenizer → recursive-descent parser → AST evaluator
- Supported functions: `ABS`, `ROUND`, `ROUNDDOWN`, `MAX`, `MIN`, `SUM`, `AVERAGE`, `POWER`, `IF`, `AND`, `OR`, `NOT`
- Four analyst tabs: Rent Roll, Valuation, Returns, Sensitivity Analysis
- Admin: Mortgage Logic tab (gated by `?admin=1` URL param — client-side only)
- Formula override persistence via `data/admin_overrides.json`
- Thread-safe engine with `threading.Lock()`
- Dynamic Rent Roll row generation based on unit count input
- PDF report tab with printable summary
- Sensitivity analysis panel
- CLI args: `--host`, `--port`, `--model`, `--formulas`, `--overrides`
