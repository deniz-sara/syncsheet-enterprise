-- ============================================================================
-- SyncSheet Enterprise - Supabase PostgreSQL Tablo Kurulumu
-- Bu SQL kodunu Supabase Dashboard -> SQL Editor ekranına yapıştırıp "Run" deyin.
-- ============================================================================

-- 1) Kullanıcı Hesapları ve Oda-PIN Geçmişi Tablosu
create table if not exists public.syncsheet_users (
  username text primary key,
  data jsonb not null,
  updated_at timestamptz not null default now()
);

-- 2) Ortak Çalışma Excel Odaları, Tablo Verisi ve Versiyon Geçmişi Tablosu
create table if not exists public.syncsheet_rooms (
  room_code text primary key,
  data jsonb not null,
  updated_at timestamptz not null default now()
);

-- İndeksler (Hızlı sorgulama için)
create index if not exists idx_syncsheet_users_updated_at on public.syncsheet_users (updated_at desc);
create index if not exists idx_syncsheet_rooms_updated_at on public.syncsheet_rooms (updated_at desc);

-- Not: Next.js API Route'ları üzerinden hem SUPABASE_SERVICE_ROLE_KEY hem de SUPABASE_ANON_KEY
-- ile sorunsuz okuma/yazma yapılabilmesi için RLS politikası:
alter table public.syncsheet_users enable row level security;
alter table public.syncsheet_rooms enable row level security;

drop policy if exists "Allow backend API access on syncsheet_users" on public.syncsheet_users;
create policy "Allow backend API access on syncsheet_users"
  on public.syncsheet_users
  for all
  using (true)
  with check (true);

drop policy if exists "Allow backend API access on syncsheet_rooms" on public.syncsheet_rooms;
create policy "Allow backend API access on syncsheet_rooms"
  on public.syncsheet_rooms
  for all
  using (true)
  with check (true);
