-- 役割分担ボード用：Supabase セットアップSQL
-- SupabaseダッシュボードのSQL Editorに貼り付けて「Run」を押してください。

-- 1) データ保存用テーブル（今のGoogleスプレッドシートのKVシートに相当）
create table if not exists kv_store (
  key text primary key,
  value jsonb not null,
  updated_at timestamptz not null default now(),
  updated_by uuid references auth.users(id)
);

-- 2) スタッフ管理者用：ユーザーごとの役割（管理者／デイサービス／訪問介護）
create table if not exists profiles (
  id uuid primary key references auth.users(id) on delete cascade,
  display_name text not null,
  role text not null check (role in ('admin','day','home')),
  created_at timestamptz not null default now()
);

-- 3) Row Level Security（ここが「サーバー側でしか守れない」本当の防御になります）
alter table kv_store enable row level security;
alter table profiles enable row level security;

-- ログイン済みのユーザーだけが kv_store を読み書きできる
create policy "logged in users can read kv" on kv_store
  for select using (auth.role() = 'authenticated');
create policy "logged in users can write kv" on kv_store
  for insert with check (auth.role() = 'authenticated');
create policy "logged in users can update kv" on kv_store
  for update using (auth.role() = 'authenticated');

-- 自分自身のプロフィール（役割）だけは本人が読める
create policy "users can read own profile" on profiles
  for select using (auth.uid() = id);

-- 管理者（admin）はスタッフ全員のプロフィールを読める
create policy "admins can read all profiles" on profiles
  for select using (
    exists (select 1 from profiles p where p.id = auth.uid() and p.role = 'admin')
  );
