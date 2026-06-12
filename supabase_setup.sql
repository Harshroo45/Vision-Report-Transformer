-- ============================================================================
--  Vision Infra · Report Analyzer — Supabase setup  (v2: tokens + files + RBAC)
--  Run this ONCE (safe to re-run):  Dashboard → SQL Editor → New query → Run
-- ============================================================================
create extension if not exists pgcrypto with schema extensions;

-- ---------------------------------------------------------------------------
-- 1) USERS  (User ID + password login, bcrypt-hashed, + a session token)
-- ---------------------------------------------------------------------------
create table if not exists public.app_users (
  id            uuid primary key default gen_random_uuid(),
  name          text not null,
  user_id       text unique not null,
  password_hash text not null,
  session_token uuid,
  created_at    timestamptz default now()
);
alter table public.app_users enable row level security;     -- locked: no policies
alter table public.app_users add column if not exists session_token uuid;  -- for existing tables

create or replace function public.register_user(p_name text, p_user_id text, p_password text)
returns json language plpgsql security definer set search_path = public, extensions as $$
declare v_id uuid;
begin
  p_user_id := lower(trim(p_user_id));
  if length(coalesce(p_name,''))=0 or length(p_user_id)=0 or length(coalesce(p_password,''))<4 then
    return json_build_object('ok',false,'error','Name, User ID and a password (min 4 chars) are required.');
  end if;
  if exists (select 1 from public.app_users where user_id=p_user_id) then
    return json_build_object('ok',false,'error','That User ID is already taken.');
  end if;
  insert into public.app_users(name,user_id,password_hash)
    values (trim(p_name),p_user_id,crypt(p_password,gen_salt('bf'))) returning id into v_id;
  return json_build_object('ok',true,'name',trim(p_name),'user_id',p_user_id);
end; $$;

create or replace function public.login_user(p_user_id text, p_password text)
returns json language plpgsql security definer set search_path = public, extensions as $$
declare v record; v_token uuid;
begin
  p_user_id := lower(trim(p_user_id));
  select * into v from public.app_users where user_id=p_user_id;
  if v.id is null then return json_build_object('ok',false,'error','Invalid User ID or password.'); end if;
  if v.password_hash = crypt(p_password, v.password_hash) then
    v_token := gen_random_uuid();
    update public.app_users set session_token=v_token where id=v.id;
    return json_build_object('ok',true,'name',v.name,'user_id',v.user_id,'token',v_token);
  end if;
  return json_build_object('ok',false,'error','Invalid User ID or password.');
end; $$;

-- resolve a token back to a user (used to restore a session on page load)
create or replace function public.whoami(p_token uuid)
returns json language plpgsql security definer set search_path = public as $$
declare v record;
begin
  select * into v from public.app_users where session_token=p_token;
  if v.id is null then return json_build_object('ok',false); end if;
  return json_build_object('ok',true,'name',v.name,'user_id',v.user_id);
end; $$;

-- admin: reset a manager's password (run manually in SQL editor)
create or replace function public.admin_reset_password(p_user_id text, p_new_password text)
returns json language plpgsql security definer set search_path = public, extensions as $$
begin
  if length(coalesce(p_new_password,''))<4 then
    return json_build_object('ok',false,'error','New password must be at least 4 characters.'); end if;
  update public.app_users set password_hash=crypt(p_new_password,gen_salt('bf')) where user_id=lower(trim(p_user_id));
  if not found then return json_build_object('ok',false,'error','No user with that User ID.'); end if;
  return json_build_object('ok',true,'message','Password updated.');
end; $$;

-- ---------------------------------------------------------------------------
-- 2) UPLOADS  (history + the stored file path + owner for RBAC)
-- ---------------------------------------------------------------------------
create table if not exists public.uploads (
  id             uuid primary key default gen_random_uuid(),
  report_type    text not null,
  file_name      text not null,
  file_size      bigint,
  row_count      int,
  uploaded_by    text,            -- display name
  uploaded_by_id text,            -- user_id (owner, for RBAC)
  storage_path   text,            -- path in the report-files bucket (nullable)
  uploaded_at    timestamptz default now()
);
alter table public.uploads add column if not exists uploaded_by_id text;
alter table public.uploads add column if not exists storage_path text;
alter table public.uploads enable row level security;
drop policy if exists "uploads read"   on public.uploads;
drop policy if exists "uploads insert" on public.uploads;
create policy "uploads read" on public.uploads for select to anon, authenticated using (true);
-- inserts & deletes go through the token-checked functions below (not direct)

create or replace function public.add_upload(p_token uuid, p_report_type text, p_file_name text,
  p_file_size bigint, p_row_count int, p_storage_path text)
returns json language plpgsql security definer set search_path = public as $$
declare v record; v_id uuid;
begin
  select * into v from public.app_users where session_token=p_token;
  if v.id is null then return json_build_object('ok',false,'error','Not signed in.'); end if;
  insert into public.uploads(report_type,file_name,file_size,row_count,uploaded_by,uploaded_by_id,storage_path)
    values(p_report_type,p_file_name,p_file_size,p_row_count,v.name,v.user_id,p_storage_path)
    returning id into v_id;
  return json_build_object('ok',true,'id',v_id);
end; $$;

-- RBAC: only the uploader (by token) can delete; also removes the stored file
create or replace function public.delete_upload(p_token uuid, p_id uuid)
returns json language plpgsql security definer set search_path = public, storage as $$
declare v record; u record;
begin
  select * into v from public.app_users where session_token=p_token;
  if v.id is null then return json_build_object('ok',false,'error','Not signed in.'); end if;
  select * into u from public.uploads where id=p_id;
  if u.id is null then return json_build_object('ok',false,'error','File not found.'); end if;
  if coalesce(u.uploaded_by_id,'') <> v.user_id then
    return json_build_object('ok',false,'error','Only the user who uploaded this file can delete it.'); end if;
  if u.storage_path is not null then
    begin delete from storage.objects where bucket_id='report-files' and name=u.storage_path; exception when others then null; end;
  end if;
  delete from public.uploads where id=p_id;
  return json_build_object('ok',true);
end; $$;

revoke all on function public.register_user(text,text,text) from public;
revoke all on function public.login_user(text,text) from public;
grant execute on function public.register_user(text,text,text) to anon, authenticated;
grant execute on function public.login_user(text,text) to anon, authenticated;
grant execute on function public.whoami(uuid) to anon, authenticated;
grant execute on function public.add_upload(uuid,text,text,bigint,int,text) to anon, authenticated;
grant execute on function public.delete_upload(uuid,uuid) to anon, authenticated;
grant execute on function public.admin_reset_password(text,text) to anon, authenticated;

-- ---------------------------------------------------------------------------
-- 3) STORAGE  (public bucket so any signed-in manager can open files)
-- ---------------------------------------------------------------------------
insert into storage.buckets (id,name,public) values ('report-files','report-files',true)
  on conflict (id) do update set public=true;
drop policy if exists "report files upload" on storage.objects;
create policy "report files upload" on storage.objects for insert to anon, authenticated
  with check (bucket_id='report-files');
-- public read is automatic for a public bucket; deletes happen via delete_upload()

-- Done.
