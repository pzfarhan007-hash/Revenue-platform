# Village Portal — Stage 1 (browser-only)

Flask backend that serves the Jamabandi Document Builder from a
per-village Google Sheet, behind a username + PIN login. Deploys to
Render. Every operational workflow — migrating a village from Excel,
adding users, rotating PINs — runs inside the browser on the deployed
app. No terminal needed after you push the code to GitHub.

Stage 2 (pending-changes button + diff modal) and Stage 3 (dual-PIN
write-back) come next. Stage 1 is read-only; the sheet is never
modified by patwaris logging in, only by the admin page.

---

## What you'll do, in order

Everything in these two sections is browser-only. Zero `python`, zero
`git` command-line.

### A. One-time cloud setup (10 min)

1. **Google Cloud project** — create a project, enable **Google Sheets API** and **Google Drive API**, create a service account, download its JSON key. (You've already done this.)
2. **Master Users sheet** — create a blank Google Sheet named e.g. "Village Portal — Users". Share with the service-account email as **Editor**. Copy its sheet ID (the long string in the URL between `/d/` and `/edit`). The admin page auto-creates the `Users` tab + header row on first write, so you can leave the sheet blank for now.
3. **Push code to GitHub** — create a new empty repo on github.com (e.g. `village-portal`). On the empty repo's page, click **"uploading an existing file"**, drag the unzipped contents of `village-portal.zip` into the browser, commit. No `git` command.
4. **Deploy on Render** — New → Blueprint → point at the repo. `render.yaml` is picked up automatically.
5. **Set four env vars in Render's Environment tab:**
   - `SECRET_KEY` — leave as auto-generated.
   - `ADMIN_PIN` — your personal admin login PIN. Pick a strong one.
   - `MASTER_USERS_SHEET_ID` — the ID from step 2.
   - `GOOGLE_SERVICE_ACCOUNT_JSON` — open the service-account JSON file in a text editor, select all, paste the whole thing into the value box. (Render preserves the newlines inside `private_key` — don't worry about them.)
6. **Deploy.** First boot takes 30 seconds on the free tier. Then open your Render URL + `/healthz` — it should say `ok`.

### B. First login + first village (5 min)

7. Go to the Render URL + `/login`. Username: `admin`. PIN: whatever you set in `ADMIN_PIN`. You land on the **Admin dashboard**.
8. **Create a blank Google Sheet** for your first village (e.g. "Jamabandi — Zalpora"). Share it with the service-account email as **Editor**. Copy its sheet ID.
9. On the Admin dashboard → **Migrate a Village** panel:
   - Choose your Zalpora `.xlsx` file.
   - Paste the village sheet ID.
   - Click **⇪ Migrate**.
   - You should see a green flash: `Migrated: N khewats, N khatas, N khasras, N portions, N mutations`.
10. On the same dashboard → **Add a User** panel:
    - Username: e.g. `patwari_zalpora`
    - Village: e.g. `Zalpora` (what shows in the UI header)
    - Login PIN: the PIN the patwari will type to log in
    - Sheet ID: the same village sheet ID
    - Role: `editor` (keep as default)
    - Click **+ Add user**.
11. Log out. Log back in as `patwari_zalpora` + their PIN. The Jamabandi builder loads with the village data pre-populated. Full v32 feature set works (preview, PDF, Word, Points, Seal, En/Ur, manual mode).

Repeat steps 8-10 for every village you add.

---

## What's in this repo

| File / folder                | Role                                                         |
|------------------------------|--------------------------------------------------------------|
| `app.py`                     | Flask routes, session, login + admin decorators              |
| `auth.py`                    | bcrypt verification, bootstrap admin, Users-sheet lookup     |
| `sheets.py`                  | gspread client — read village, write village, add user       |
| `migrate.py`                 | Zalpora Excel → six-tab parser; also usable as a CLI         |
| `templates/jamabandi.html`   | The v32 HTML, upload box replaced by a `/api/village-data` fetch and a village + logout banner |
| `templates/login.html`       | Login form                                                   |
| `templates/admin.html`       | Admin dashboard — migrate, add user, rotate PIN, hash PIN    |
| `requirements.txt`           | Pinned deps                                                  |
| `Procfile`, `render.yaml`    | Deployment                                                   |
| `.gitignore`                 | Protects secrets                                             |
| `.env.example`               | Template for local-dev env vars                              |
| `users.json.example`         | Fallback for running without a Users sheet (dev only)        |

---

## Admin dashboard — what each panel does

| Panel              | Does                                                        |
|--------------------|-------------------------------------------------------------|
| Migrate a Village  | Reads your uploaded `.xlsx`, parses it with the same logic v32 uses, writes six tabs into the target Google Sheet. Overwrites existing tabs with those six names. |
| Add a User         | Appends one row to the master Users sheet. Bcrypt-hashes both the login PIN and the (optional) sync PIN. If the Users sheet doesn't have the required header row, it's created automatically. |
| Current Users      | Lists everyone in the Users sheet, with a one-click **Rotate login PIN** control per row. (Rotate sync PIN comes in Stage 3.) |
| Hash a PIN         | Standalone utility — paste a PIN, get its bcrypt hash. Only needed if you prefer editing the Users sheet directly in Google Sheets. |

---

## What Stage 1 does NOT do yet

- **No write-back from the Jamabandi builder.** Patwaris can edit owner names, cultivators, portions etc. inside the UI, but those edits live in the browser only. Refresh = lose them. (Apply button + diff modal = Stage 2; dual-PIN write-back = Stage 3.)
- No admin-level view of villages (list of sheets, last-sync time). That's Stage 4.

---

## Local dev (optional — only if you want to run on your laptop)

```bash
python -m venv .venv && source .venv/bin/activate      # Mac/Linux
python -m venv .venv && .venv\Scripts\activate         # Windows

pip install -r requirements.txt
cp .env.example .env                                   # fill the values
export $(grep -v '^#' .env | xargs)                    # Mac/Linux

python app.py                                          # http://localhost:5000
```

For dev without a master sheet, leave `MASTER_USERS_SHEET_ID` empty
and provide `users.json` (see `users.json.example`). The CLI
`python migrate.py <xlsx> <sheet_id>` still works too.

---

## Security notes

- Service-account JSON is a credential. Never commit it to GitHub. The `.gitignore` already blocks `service-account.json`, `*.env`, and `users.json`.
- If you've ever uploaded or pasted the JSON somewhere non-private (chats, email, screenshots), **rotate the key**: Service Accounts → your account → Keys → Add key → Create new → download → delete the old key. Takes 30 seconds.
- The `ADMIN_PIN` env var is also a credential — the Render dashboard treats env-var values as secrets. Pick something you'd be comfortable storing there and nowhere else.
