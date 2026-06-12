# Setup — Auth, File Storage (RBAC) & Recent Uploads

Two steps: run the SQL, then deploy. Everything else is already wired into index.html.

## Step 1 — Run the database setup (safe to re-run)
1. Open: https://supabase.com/dashboard/project/tytzjbvmjtdfxfvftigq/sql/new
2. Paste all of `supabase_setup.sql` → **Run**.

This creates / updates:
- `app_users` — managers, bcrypt-hashed passwords, a `session_token` per login.
- `uploads` — upload history + the stored file path + the owner (`uploaded_by_id`).
- Functions: `register_user`, `login_user` (returns a token), `whoami`, `add_upload`,
  `delete_upload` (RBAC), `admin_reset_password`.
- A **public Storage bucket** `report-files` so any signed-in manager can open files.

## Step 2 — Deploy to Vercel (free)
Put `index.html` (+ `vercel.json`) in your repo and import at vercel.com (preset **Other**,
no build, output dir `.`). The Supabase URL + anon key are already in `index.html`.

## What you get
- **Login / Register** with User ID + password (no email). Sessions persist across refresh
  and are validated against the database on load (`whoami`).
- **Recent uploads, shared:** every processed file is uploaded to Storage and listed for ALL
  managers under that report type. Click a file name to open it in a new tab.
- **RBAC delete:** the trash icon appears ONLY on files you uploaded. The delete itself is
  enforced in the database (`delete_upload` checks your session token's user against the
  file's owner) and also removes the stored file — others simply can't delete your files.
- **Styled Excel export (ExcelJS):** navy bold header, bordered cells, zebra striping,
  date/number formats, auto-fit column widths, frozen header + auto-filter.
- **Multi-select Department filter** (checkbox dropdown), 3D/glass UI, and a footer.

## EDIT before you go live
Open `index.html`, find the `COMPANY` block near the top, and replace the placeholders:
```js
const COMPANY = {
  name: "Vision Infra Equipment Solutions Ltd.",
  address: "…your real corporate address…",   // EDIT
  linkedin: "https://www.linkedin.com/company/your-company-here",  // EDIT
  email: "info@visioninfra.example"            // EDIT
};
```

## Honest notes
- Files up to ~45 MB are stored for preview (Supabase free per-file limit is ~50 MB). Bigger
  files (your 75 MB .xls) are listed as history but not stored — Save As .xlsx (~12 MB) first.
- The bucket is **public**: anyone with a file's URL can open it. Good for an internal team;
  if you need per-user file access control, that requires Supabase Auth (JWT) — ask and I'll
  switch it. Delete is properly owner-restricted regardless.
- The anon key in the frontend is public-by-design and safe. Never expose the service_role key.

## Forgot a password?
Run in SQL editor: `select public.admin_reset_password('theuserid','NewPass123');`
