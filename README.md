# Daily/Weekly/Quarter Planner System

A personal productivity system built around six life categories (GLS/Work, Book Keeping,
Administration, Student Ministry, Personal, Lopez Home), designed with ADHD, mental health,
and a July 2026 bariatric surgery recovery window in mind.

## Structure

- **app/** — the working planner
  - `daily_planner_cloud.html` — cross-device synced version (Supabase-backed, requires setup — see `supabase_schema.sql`)
  - `daily_planner_local.html` — standalone version (localStorage only, no setup required, no cross-device sync)
  - `supabase_schema.sql` — run this once in a Supabase project's SQL Editor to enable the cloud version
- **docs/** — printable/editable Word templates and filled examples
  - `daily_planner_template.docx`, `weekly_review_template.docx`, `quarter_month_plan.docx` — blank templates
  - `weekly_review_Jul13-19.docx`, `weekly_review_Jul20-26.docx`, `weekly_review_Jul27-Aug02_RECOVERY.docx` — filled for the real July 2026 plan
  - `daily_planner_Jul14_TODAY.docx`, `daily_planner_Jul30_RECOVERY.docx` — filled daily examples
- **scripts/** — Node.js source (using the `docx` package) that generates the .docx files in `docs/`

## Deploying the cloud app

1. Create a free Supabase project, run `app/supabase_schema.sql` in its SQL Editor.
2. Paste your Supabase Project URL and anon key into the `SUPABASE_URL` / `SUPABASE_ANON_KEY`
   constants near the top of `app/daily_planner_cloud.html`.
3. Deploy the `app/` folder to a static host (GitHub Pages, Netlify, etc).
4. Open the deployed URL in Safari on each device, create one account, and log into that
   same account everywhere to sync.

## Rotation & fixed schedule (baked into the app's seed data)

- Mon: Book Keeping · Tue: Administration · Wed: Student Ministry · Thu: Personal · Fri: Lopez Home · Sat/Sun: Family
- Tue 6–9pm: family movie night · Thu 6–7pm: prayer · Sat: Sabbath (nothing scheduled) · Sun 8am–1pm: church
- Recovery window auto-suggested Jul 29 – Aug 15, 2026 (surgery on Jul 29)
