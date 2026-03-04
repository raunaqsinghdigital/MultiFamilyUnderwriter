# Changelog

All notable changes to this project are documented here.
Format follows [Keep a Changelog](https://keepachangelog.com/en/1.1.0/).

---

## [Unreleased]

---

## [1.2.0] — 2026-03-04

### Added
- **Value Add Underwriter** — second product tab alongside Buy & Hold; fully isolated state, model, and formula engine; accessible after login via the product nav at the top of the workspace
- `data/value_add_model.json` — workbook model extracted from the Value Add Excel template (`09_MFM_Property_Analyzer_VALUE_ADD_Rev1_-_11212_124_ST.xls`) covering Rent Roll, Valuation, and Refinance sheets
- `data/value_add_formulas_expanded.json` — 127 pre-computed formula entries for the Value Add engine (Valuation + Refinance + Rent Roll totals)
- `data/value_add_admin_overrides.json` — separate admin mortgage formula override file for Value Add (isolated from Buy & Hold's `admin_overrides.json`)
- `scripts/extract_value_add_model.py` — one-time extraction script: decrypts the Value Add XLS (VelvetSweatshop default password), reads cell values via `xlrd`, detects input cells by fill color `(255, 255, 204)`, and outputs the two JSON data files
- `?product=buy-hold|value-add` query param on `/api/model`, `/api/calculate`, and `/api/admin/*` endpoints — routes each request to the correct isolated `AppState` instance
- `_resolve_state()` helper on `UnderwriterHandler` — resolves the correct `AppState` (Buy & Hold or Value Add) from the `?product=` query param; returns 503 if Value Add model files are missing
- Separate admin override persistence per product: edits to Buy & Hold mortgage formulas only touch `admin_overrides.json`; Value Add only touches `value_add_admin_overrides.json`
- Per-product frontend state cache (`_productStateCache`) — saves and restores `workbookModel`, `rentRollState`, `sensitivityState`, `latestFormulaValues`, and `hasSuccessfulCalculation` when switching tabs
- `VA_*` metadata constants in `app.js` — parallel to Buy & Hold constants; cover `VA_VALUATION_SECTIONS`, `VA_REFINANCE_SECTIONS`, `VA_VALUATION_RESULT_FIELDS`, `VA_REFINANCE_RESULT_FIELDS`, `VA_VALUATION_INPUT_HINTS`, etc.
- Product-aware metadata getter functions (`getValuationSections()`, `getValuationResultFields()`, `getPercentWholeInputKeys()`, etc.) — all rendering functions now call these instead of hardcoded BH constants
- Value Add-specific rent roll injection (`_apply_dynamic_rent_roll_va`) — maps Rent Roll columns C/D/E/G/H (VA layout) and sets `Rent Roll!G15` (current total) + `Rent Roll!H15` (projected total) which feed `Valuation!G10` and `Refinance!G10` via formula chain

### Changed
- **Product selector nav** — Value Add tab enabled (was `disabled` / "Coming Soon"); click switches workspace and API product context
- **API URL references** — `modelUrl`, `calculateUrl`, `adminMortgageUrl`, `adminMortgageOverrideUrl` constants replaced with `getModelUrl()`, `getCalculateUrl()`, `getAdminMortgageUrl()`, `getAdminMortgageOverrideUrl()` functions that append the active product param
- **`renderWorkbook`** — sheet routing extended to handle `Refinance` (→ `renderValuationPanel`) and `Return` (→ `renderReturnsPanel`) alongside existing BH sheet names
- **`AppState`** — accepts `product` parameter; passes it to `PurePythonUnderwriterEngine`
- **`PurePythonUnderwriterEngine`** — accepts `product` parameter; `_apply_dynamic_rent_roll` dispatches to the correct product-specific method
- **`main()`** — loads both Buy & Hold and Value Add models at startup; attaches both as `server.buy_hold_state` / `server.value_add_state`; prints which products loaded
- **README** — Value Add Underwriter marked ✅ Available; file structure updated with new data and scripts files; opening description updated

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
