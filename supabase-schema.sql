create table if not exists public.work_tasks (
  id text primary key,
  payload jsonb not null,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now()
);

create table if not exists public.saved_outputs (
  id text primary key,
  payload jsonb not null,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now()
);

create table if not exists public.app_meta (
  key text primary key,
  value jsonb,
  updated_at timestamptz not null default now()
);

alter table public.work_tasks enable row level security;
alter table public.saved_outputs enable row level security;
alter table public.app_meta enable row level security;

drop policy if exists "Allow anon read work tasks" on public.work_tasks;
drop policy if exists "Allow anon write work tasks" on public.work_tasks;
drop policy if exists "Allow anon update work tasks" on public.work_tasks;
drop policy if exists "Allow anon delete work tasks" on public.work_tasks;
drop policy if exists "Allow anon read saved outputs" on public.saved_outputs;
drop policy if exists "Allow anon write saved outputs" on public.saved_outputs;
drop policy if exists "Allow anon update saved outputs" on public.saved_outputs;
drop policy if exists "Allow anon read app meta" on public.app_meta;
drop policy if exists "Allow anon write app meta" on public.app_meta;
drop policy if exists "Allow anon update app meta" on public.app_meta;

create policy "Allow anon read work tasks"
  on public.work_tasks for select
  to anon
  using (true);

create policy "Allow anon write work tasks"
  on public.work_tasks for insert
  to anon
  with check (true);

create policy "Allow anon update work tasks"
  on public.work_tasks for update
  to anon
  using (true)
  with check (true);

create policy "Allow anon delete work tasks"
  on public.work_tasks for delete
  to anon
  using (true);

create policy "Allow anon read saved outputs"
  on public.saved_outputs for select
  to anon
  using (true);

create policy "Allow anon write saved outputs"
  on public.saved_outputs for insert
  to anon
  with check (true);

create policy "Allow anon update saved outputs"
  on public.saved_outputs for update
  to anon
  using (true)
  with check (true);

create policy "Allow anon read app meta"
  on public.app_meta for select
  to anon
  using (true);

create policy "Allow anon write app meta"
  on public.app_meta for insert
  to anon
  with check (true);

create policy "Allow anon update app meta"
  on public.app_meta for update
  to anon
  using (true)
  with check (true);
