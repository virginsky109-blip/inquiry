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

-- 안전운임(전사 공통, 국토부 고시 전국표) 전용 테이블.
-- rate_sets.rows(jsonb) 한 덩어리에 넣기엔 행이 너무 많아(수만 행) 실제 컬럼으로 쪼개
-- port/place로 인덱스 조회한다. 안전위탁운임·운수사업자간운임은 애초에 저장하지 않는다.
create table if not exists public.safe_rates_common (
  id bigint generated always as identity primary key,
  port text not null,
  place text not null,
  size text not null,
  amount numeric not null
);
create index if not exists safe_rates_common_port_idx on public.safe_rates_common (port);
create index if not exists safe_rates_common_place_idx on public.safe_rates_common (place);

-- RLS: 로그인한 담당자 전원 읽기/쓰기 허용. 단, scope='common' 요율(부대비용)과
-- 안전운임 전국표는 로그인 안 한 담당자도 조회는 가능해야 하므로 anon 읽기 정책을 별도로 둔다.
alter table public.rate_sets enable row level security;
alter table public.quotes enable row level security;
alter table public.safe_rates_common enable row level security;

drop policy if exists "team all rate_sets" on public.rate_sets;
create policy "team all rate_sets" on public.rate_sets
  for all to authenticated using (true) with check (true);

drop policy if exists "anon reads common rate_sets" on public.rate_sets;
create policy "anon reads common rate_sets" on public.rate_sets
  for select to anon using (scope = 'common');

drop policy if exists "team all quotes" on public.quotes;
create policy "team all quotes" on public.quotes
  for all to authenticated using (true) with check (true);

drop policy if exists "team all safe_rates_common" on public.safe_rates_common;
create policy "team all safe_rates_common" on public.safe_rates_common
  for all to authenticated using (true) with check (true);

drop policy if exists "anon reads safe_rates_common" on public.safe_rates_common;
create policy "anon reads safe_rates_common" on public.safe_rates_common
  for select to anon using (true);
