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
- **supabase/functions/breakdown-goal/** — Edge Function that powers the "Break it down" AI goal-breakdown feature (see below)

## Deploying the cloud app

1. Create a free Supabase project, run `app/supabase_schema.sql` in its SQL Editor.
2. Paste your Supabase Project URL and anon key into the `SUPABASE_URL` / `SUPABASE_ANON_KEY`
   constants near the top of `app/daily_planner_cloud.html`.
3. Deploy the `app/` folder to a static host (GitHub Pages, Netlify, etc).
4. Open the deployed URL in Safari on each device, create one account, and log into that
   same account everywhere to sync.

## Setting up "Break it down" (AI goal breakdown)

The "Break it down" feature (the Big Goal box and the 🧩 buttons on tasks) turns an overwhelming
goal into 3-5 tiny physical next steps. It calls a small Supabase Edge Function so your Anthropic
API key never has to sit in the browser.

1. Get an Anthropic API key at [console.anthropic.com](https://console.anthropic.com) (Settings →
   API Keys). Add a small amount of billing credit — usage for this feature is typically a fraction
   of a cent per breakdown.
2. Install the Supabase CLI if you haven't already (`brew install supabase/tap/supabase` on a Mac),
   then run `supabase login` and `supabase link --project-ref <your-project-ref>` from this repo's
   root (your project ref is in your Supabase project's URL/settings).
3. Set the two secrets the function needs:
   ```
   supabase secrets set ANTHROPIC_API_KEY=sk-ant-...
   supabase secrets set PLANNER_SHARED_SECRET=<make up any long random string>
   ```
4. Deploy it: `supabase functions deploy breakdown-goal --no-verify-jwt`
5. Supabase will print your function's URL (looks like
   `https://<project-ref>.supabase.co/functions/v1/breakdown-goal`). Paste that into
   `BREAKDOWN_FUNCTION_URL` near the top of `app/daily_planner_local.html`, and paste the same
   random string from step 3 into `BREAKDOWN_SHARED_SECRET` right below it.

The `--no-verify-jwt` flag is what lets the app call this function without a full Supabase login
flow — the `x-planner-secret` header is what stands in for auth instead. This is a reasonable
tradeoff for a personal single-user tool, but keep in mind that secret lives in the HTML file
itself, so don't publish this file somewhere public without swapping to real Supabase Auth first.

## Rotation & fixed schedule (baked into the app's seed data)

- Mon: Book Keeping · Tue: Administration · Wed: Student Ministry · Thu: Personal · Fri: Lopez Home · Sat/Sun: Family
- Tue 6–9pm: family movie night · Thu 6–7pm: prayer · Sat: Sabbath (nothing scheduled) · Sun 8am–1pm: church
- Recovery window auto-suggested Jul 29 – Aug 15, 2026 (surgery on Jul 29)
