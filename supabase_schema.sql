-- ============================================================
-- 포워딩 견적 시스템 - 팀 공용 DB 스키마
-- 실행 방법: Supabase 대시보드 → SQL Editor → 붙여넣고 Run (1회)
-- 계정 생성: Authentication → Users → Add user (관리자가 담당자별로 생성)
-- 보안: RLS 활성화 - 로그인(authenticated)한 담당자만 읽기/쓰기 가능.
--       소규모 팀 상부상조 방침에 따라 담당자 간 격리는 하지 않음.
-- ============================================================

-- 요율표 세트 (rate_master / safe_rate / charge_rule 공용)
create table if not exists public.rate_sets (
  id text primary key,
  name text not null,
  kind text not null,               -- rate_master | safe_rate | charge_rule
  enabled boolean not null default true,
  source text default 'file',       -- file | ai | seed
  scope text default 'personal',    -- common | personal | client
  client text default '',           -- scope=client 일 때 고객사명
  created_label text default '',
  rows jsonb not null default '[]',
  updated_by text,
  updated_at timestamptz default now()
);

-- 저장 견적
create table if not exists public.quotes (
  id text primary key,
  name text not null,
  saved_label text default '',
  customer text default '',
  data jsonb not null,
  updated_by text,
  updated_at timestamptz default now()
);

-- 고객사명 자동완성 성능용 인덱스
create index if not exists quotes_customer_idx on public.quotes (customer);

-- RLS: 로그인한 담당자 전원 읽기/쓰기 허용
alter table public.rate_sets enable row level security;
alter table public.quotes enable row level security;

drop policy if exists "team all rate_sets" on public.rate_sets;
create policy "team all rate_sets" on public.rate_sets
  for all to authenticated using (true) with check (true);

drop policy if exists "team all quotes" on public.quotes;
create policy "team all quotes" on public.quotes
  for all to authenticated using (true) with check (true);
