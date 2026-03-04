# Multifamily Underwriter

A browser-based multifamily real estate underwriting tool with two products: **Buy & Hold Underwriter** and **Value Add Underwriter**. Both are available after login via a tab switcher at the top of the workspace.

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
| Value Add Underwriter | ✅ Available |

---

## Prerequisites

- A [Supabase](https://supabase.com) account (free)
- A [Render](https://render.com) account (free)
- The repo pushed to GitHub

---

## One-Time Setup

> **Important — follow these steps in order.** Inviting users before deploying to Render causes invite emails to link to `localhost:3000` (Supabase's default), which will not work.

---

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

### Step 3 — Deploy to Render

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
8. Note your live URL (e.g. `https://multifamily-underwriter.onrender.com`) — you need it in the next step

---

### Step 4 — Configure Supabase Redirect URL

Before inviting users, Supabase must know where to redirect invite email links. Without this, invite emails link to `localhost:3000` (Supabase's default) and cannot be used.

1. Go to **Supabase Dashboard → Authentication → URL Configuration**
2. Set **Site URL** to your Render URL (e.g. `https://multifamily-underwriter.onrender.com`)
3. Under **Redirect URLs**, add: `https://multifamily-underwriter.onrender.com/**`
4. Click **Save**

---

### Step 5 — Create Users in Supabase

1. Go to **Supabase Dashboard → Authentication → Users**
2. Click **Invite user** → enter the email address
3. Supabase sends an invite email; the user clicks the link and is taken to the app to set their password

> **Note:** The invite link redirects to your Render URL (configured in Step 4). On first visit the app shows a "Set Your Password" screen — the user creates a password and is immediately taken into the workbook.

---

### Step 6 — Assign Roles (admin or analyst)

Roles are stored in `app_metadata` (only the service role key can write this — users cannot change their own role).

#### Option A — SQL Editor (recommended, no keys required)

1. Go to **Supabase Dashboard → SQL Editor**
2. Run the query for the role you want to assign (replace the email address):

```sql
-- Set role to "admin"
UPDATE auth.users
SET raw_app_meta_data = raw_app_meta_data || '{"role": "admin"}'::jsonb
WHERE email = 'user@example.com';

-- Set role to "analyst"
UPDATE auth.users
SET raw_app_meta_data = raw_app_meta_data || '{"role": "analyst"}'::jsonb
WHERE email = 'user@example.com';
```

The `||` operator merges the new key into the existing metadata — `provider` and `providers` fields are preserved. The dashboard session handles authentication; no service role key required.

#### Option B — curl with silent prompt (fallback)

If you prefer the API, use `read -rs` so the key is never echoed to the terminal or saved in shell history:

```bash
# Prompts silently — key is not stored in ~/.zsh_history
read -rs SERVICE_ROLE_KEY

# Set role to "admin"
curl -X PATCH \
  https://<project-ref>.supabase.co/auth/v1/admin/users/<user-uid> \
  -H "Authorization: Bearer $SERVICE_ROLE_KEY" \
  -H "Content-Type: application/json" \
  -d '{"app_metadata": {"role": "admin"}}'

unset SERVICE_ROLE_KEY  # clear from memory immediately
```

Get `<project-ref>` from: **Project Settings → API → Project URL**
Get `<user-uid>` from: **Authentication → Users** (shown in the user list)

> **Security note:** Never paste the `service_role_key` inline in a command — it will be saved to shell history and visible in process listings.

After setting the role, the user must **sign out and sign back in** for the new role to appear in their JWT.

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
| Set role to admin | SQL Editor: `UPDATE auth.users SET raw_app_meta_data = raw_app_meta_data \|\| '{"role":"admin"}'::jsonb WHERE email = '...'` (see Step 6) |
| Set role to analyst | Same query, with `"role": "analyst"` |
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
| Invite email links to localhost:3000 | Supabase Site URL not configured | Set Site URL in Auth → URL Configuration to your Render URL (Step 4) |
| "OTP expired" when clicking invite link | Invite link used after Site URL was wrong; new invite needed | Fix Site URL (Step 4), then re-invite the user |

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
├── scripts/
│   └── extract_value_add_model.py  # One-time: XLS → value_add_model.json
└── data/
    ├── workbook_model.json           # Buy & Hold model (sheets, cells, input metadata)
    ├── formulas_expanded.json        # Buy & Hold pre-computed formula map
    ├── admin_overrides.json          # Buy & Hold admin mortgage formula overrides
    ├── value_add_model.json          # Value Add model (generated by extract script)
    ├── value_add_formulas_expanded.json  # Value Add pre-computed formula map
    ├── value_add_admin_overrides.json    # Value Add admin mortgage overrides
    └── backup_1_1_sheet.json         # Backup of hidden "1.1" worksheet
```
