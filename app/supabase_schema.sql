-- Daily Planner — Supabase schema
-- Run this once in your Supabase project's SQL Editor (Dashboard → SQL Editor → New query → paste → Run)

-- 1. One row per day, per user. `data` holds the full planner record as JSON
--    (priorities, lanes, habits, brain dump, etc.) — same shape the app already uses.
create table if not exists planner_days (
  id uuid primary key default gen_random_uuid(),
  user_id uuid not null references auth.users(id) on delete cascade,
  date date not null,
  data jsonb not null,
  updated_at timestamptz not null default now(),
  unique (user_id, date)
);

-- 2. One row per user, tracking the last date they opened the app.
--    This is what powers the "new day / Morning Review" detection.
create table if not exists planner_meta (
  user_id uuid primary key references auth.users(id) on delete cascade,
  last_date date,
  updated_at timestamptz not null default now()
);

-- 3. Row Level Security: each user can only ever read/write their own rows.
--    This is what makes it safe to ship the public "anon" key inside the app file.
alter table planner_days enable row level security;
alter table planner_meta enable row level security;

create policy "select own days" on planner_days
  for select using (auth.uid() = user_id);
create policy "insert own days" on planner_days
  for insert with check (auth.uid() = user_id);
create policy "update own days" on planner_days
  for update using (auth.uid() = user_id);

create policy "select own meta" on planner_meta
  for select using (auth.uid() = user_id);
create policy "insert own meta" on planner_meta
  for insert with check (auth.uid() = user_id);
create policy "update own meta" on planner_meta
  for update using (auth.uid() = user_id);

-- 4. Keep updated_at fresh automatically on every save.
create or replace function set_updated_at()
returns trigger as $$
begin
  new.updated_at = now();
  return new;
end;
$$ language plpgsql;

create trigger planner_days_updated_at
  before update on planner_days
  for each row execute function set_updated_at();

create trigger planner_meta_updated_at
  before update on planner_meta
  for each row execute function set_updated_at();
