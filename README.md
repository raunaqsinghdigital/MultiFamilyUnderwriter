# Multifamily Underwriter

A browser-based multifamily real estate underwriting tool. Currently ships with the **Buy & Hold Underwriter**. A **Value Add Underwriter** is planned for a future release.

---

## Architecture

```
Browser
  └─ Static files (HTML / CSS / JS) ──── served by Python HTTP server on Render
  └─ POST /api/calculate ──────────────► Python formula engine (pure stdlib)
  └─ GET  /api/model ──────────────────► Returns workbook structure + inputs
  └─ /api/admin/* ─────────────────────► Admin-only endpoints (role=admin JWT claim)

Authentication
  └─ Supabase Auth (email/password)
       └─ Issues HS256 JWT on login
       └─ JWT verified by Python server (stdlib hmac + hashlib, no packages)
       └─ Role stored in app_metadata.role → read from raw_app_meta_data in JWT
```

**Hosting:** [Render](https://render.com) (free tier — spins down after 15 min idle; first request after idle takes ~30 s)
**Auth:** [Supabase Auth](https://supabase.com) (free tier — 50,000 MAU)

---

## Product Roadmap

| Product | Status |
|---------|--------|
| Buy & Hold Underwriter | ✅ Available |
| Value Add Underwriter | 🔜 Coming Soon |

---

## Prerequisites

- A [Supabase](https://supabase.com) account (free)
- A [Render](https://render.com) account (free)
- The repo pushed to GitHub

---

## One-Time Setup

### Step 1 — Create a Supabase Project

1. Go to [supabase.com](https://supabase.com) → **New project**
2. Note your project name/ref (e.g. `abcdefghij`)
3. Navigate to **Project Settings → API**
4. Copy these three values — you'll need them throughout this setup:

| Value | Where to find it |
|-------|-----------------|
| `SUPABASE_URL` | Project Settings → API → Project URL |
| `SUPABASE_ANON_KEY` | Project Settings → API → Project API keys → `anon` `public` |
| `SUPABASE_JWT_SECRET` | Project Settings → API → JWT Settings → JWT Secret |

> **Security note:** `SUPABASE_URL` and `SUPABASE_ANON_KEY` are safe to put in frontend JS (they're public keys protected by Row Level Security). `SUPABASE_JWT_SECRET` must stay server-side only — never commit it.

---

### Step 2 — Paste Supabase Public Credentials into the Frontend

Open [static/app.js](static/app.js) and replace the placeholder values near the top:

```js
const SUPABASE_URL = "https://YOUR_PROJECT_REF.supabase.co";
const SUPABASE_ANON_KEY = "YOUR_ANON_KEY";
```

Replace with your actual values from Step 1.

---

### Step 3 — Create Users in Supabase

1. Go to **Supabase Dashboard → Authentication → Users**
2. Click **Invite user** → enter the email address
3. Supabase sends an email invite; the user sets their own password

---

### Step 4 — Assign Roles (admin or analyst)

Roles are stored in `app_metadata` (only the service role key can write this — users cannot change their own role).

Get the `service_role_key` from: **Project Settings → API → Project API keys → `service_role` `secret`**

Get the user's UID from: **Authentication → Users** (shown in the user list)

Then run the following curl command (replace `<project-ref>`, `<service_role_key>`, and `<user-uid>`):

```bash
# Set role to "admin"
curl -X PATCH \
  https://<project-ref>.supabase.co/auth/v1/admin/users/<user-uid> \
  -H "Authorization: Bearer <service_role_key>" \
  -H "Content-Type: application/json" \
  -d '{"app_metadata": {"role": "admin"}}'

# Set role to "analyst"
curl -X PATCH \
  https://<project-ref>.supabase.co/auth/v1/admin/users/<user-uid> \
  -H "Authorization: Bearer <service_role_key>" \
  -H "Content-Type: application/json" \
  -d '{"app_metadata": {"role": "analyst"}}'
```

> **Security note:** Never expose the `service_role_key` in frontend code or commit it to git.

After setting the role, the user must **sign out and sign back in** for the new role to appear in their JWT.

---

### Step 5 — Deploy to Render

1. Push this repo to GitHub (if not already done)
2. Go to [render.com](https://render.com) → **New → Web Service**
3. Connect your GitHub repo
4. Render will auto-detect `render.yaml` and configure the service
5. Click **Create Web Service**
6. Once the service is created, go to the **Environment** tab and add:

| Key | Value |
|-----|-------|
| `SUPABASE_JWT_SECRET` | Paste the JWT Secret from Supabase Step 1 |

7. Click **Save Changes** — Render will redeploy automatically
8. Your app is live at `https://multifamily-underwriter.onrender.com` (or similar)

---

## Local Development

```bash
# No dependencies to install — pure Python stdlib
python app.py
# Server starts at http://0.0.0.0:8000

# Use a specific host/port
python app.py --host 127.0.0.1 --port 3000

# Run with a JWT secret for local auth testing
SUPABASE_JWT_SECRET=your-secret python app.py
```

> **Note:** The `?admin=1` URL trick has been removed. Admin features are now gated by the `admin` role in your Supabase JWT. Log in with an admin-role account to see the Admin: Mortgage Logic tab.

---

## Environment Variables Reference

| Variable | Required | Default | Description |
|----------|----------|---------|-------------|
| `SUPABASE_JWT_SECRET` | **Yes** (production) | `""` | JWT signing secret from Supabase. If empty, all API calls return 401. |
| `ALLOWED_ORIGINS` | No | `""` (allow all) | Comma-separated allowed CORS origins. Example: `https://www.ramavasa.com`. Empty = no restriction. |
| `HOST` | No | `0.0.0.0` | Server bind address. |
| `PORT` | No | `8000` | Server port. Injected automatically by Render. |

---

## CORS / Domain Whitelisting

To restrict which domains can make cross-origin API requests, set `ALLOWED_ORIGINS` in your Render environment:

```
https://www.ramavasa.com,https://ramavasa.com
```

- Requests **without** an `Origin` header (direct browser access, curl) are always allowed.
- Cross-origin requests from domains **not** in the list receive a `403` response.
- Leave empty to allow all origins (suitable for direct standalone access).

---

## RBAC Reference

| Feature | Analyst | Admin |
|---------|---------|-------|
| Rent Roll tab | ✅ | ✅ |
| Valuation tab | ✅ | ✅ |
| Returns tab | ✅ | ✅ |
| Sensitivity Analysis tab | ✅ | ✅ |
| Admin: Mortgage Logic tab | ❌ | ✅ |
| Edit/save mortgage formula overrides | ❌ | ✅ |

---

## User Management Reference

| Task | How |
|------|-----|
| Create a user | Supabase Dashboard → Authentication → Users → Invite user |
| Set role to admin | `curl PATCH /auth/v1/admin/users/<uid>` with `app_metadata.role = "admin"` (see Step 4) |
| Set role to analyst | Same curl, with `"role": "analyst"` |
| Remove a user | Supabase Dashboard → Authentication → Users → Delete |
| User forgets password | Supabase Dashboard → Authentication → Users → Send reset email |

---

## Troubleshooting

| Problem | Cause | Fix |
|---------|-------|-----|
| App loads slowly on first visit | Render free tier spins down after 15 min idle | Wait ~30 s for cold start; upgrade to paid tier to disable |
| Login returns "Invalid login credentials" | Wrong email/password | Reset password via Supabase dashboard |
| Admin tab not visible after role change | JWT cached; old token lacks `admin` role | Sign out and sign back in |
| All API calls return 401 | `SUPABASE_JWT_SECRET` not set in Render | Add it in Render → Environment vars |
| CORS blocked in browser console | `ALLOWED_ORIGINS` set but domain not in list | Add domain to `ALLOWED_ORIGINS` in Render env vars |
| Supabase project paused | Free tier pauses after 1 week idle | Visit Supabase dashboard → click **Restore** |

---

## File Structure

```
MultiFamilyUnderwriter/
├── app.py                    # Python HTTP server + formula engine (1,500+ lines)
├── requirements.txt          # Empty — stdlib only; exists for Render Python detection
├── render.yaml               # Render deployment config
├── run_local.sh              # Local dev shortcut: python app.py
├── README.md                 # This file
├── CHANGELOG.md              # Version history
├── static/
│   ├── index.html            # Single-page app shell
│   ├── app.js                # Frontend logic (~4,000 lines)
│   └── styles.css            # UI styles
└── data/
    ├── workbook_model.json   # Core model: sheets, cells, input metadata (740 KB)
    ├── formulas_expanded.json # Pre-computed formula map
    ├── admin_overrides.json  # Persisted admin mortgage formula edits
    └── backup_1_1_sheet.json # Backup of hidden "1.1" worksheet
```
