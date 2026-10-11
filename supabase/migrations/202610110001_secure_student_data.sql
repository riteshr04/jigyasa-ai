create table if not exists public.documents (
    id uuid primary key,
    owner_id uuid not null references auth.users (id) on delete cascade,
    name text not null check (char_length(name) between 1 and 255),
    ext text not null check (ext in ('pdf', 'pptx', 'docx', 'txt')),
    size bigint not null check (size > 0 and size <= 26214400),
    added_at timestamptz not null default now(),
    text_content text not null,
    storage_path text not null unique
);

create index if not exists documents_owner_added_at_idx
    on public.documents (owner_id, added_at);

alter table public.documents enable row level security;
revoke all on public.documents from anon, authenticated;
grant all on public.documents to service_role;

insert into storage.buckets (id, name, public, file_size_limit, allowed_mime_types)
values (
    'study-materials',
    'study-materials',
    false,
    26214400,
    array[
        'application/pdf',
        'application/vnd.openxmlformats-officedocument.presentationml.presentation',
        'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
        'text/plain',
        'application/octet-stream'
    ]
)
on conflict (id) do update
set public = false,
    file_size_limit = excluded.file_size_limit,
    allowed_mime_types = excluded.allowed_mime_types;

create or replace function public.enforce_document_storage_quota()
returns trigger
language plpgsql
security definer
set search_path = public, auth, pg_temp
as $$
declare
    used_bytes bigint;
    used_documents bigint;
begin
    perform 1 from auth.users where id = new.owner_id for update;
    select coalesce(sum(size), 0), count(*)
      into used_bytes, used_documents
      from public.documents
     where owner_id = new.owner_id
       and id <> new.id;
    if used_documents >= 25 then
        raise exception 'Account document limit exceeded.' using errcode = '23514';
    end if;
    if used_bytes + new.size > 104857600 then
        raise exception 'Account storage limit exceeded.' using errcode = '23514';
    end if;
    return new;
end;
$$;

drop trigger if exists documents_enforce_storage_quota on public.documents;
create trigger documents_enforce_storage_quota
before insert or update of size, owner_id on public.documents
for each row execute function public.enforce_document_storage_quota();

create table if not exists public.daily_ai_usage (
    owner_id uuid not null references auth.users (id) on delete cascade,
    usage_date date not null,
    questions_used integer not null default 0 check (questions_used between 0 and 200),
    primary key (owner_id, usage_date)
);

alter table public.daily_ai_usage enable row level security;
revoke all on public.daily_ai_usage from anon, authenticated;
grant all on public.daily_ai_usage to service_role;

create or replace function public.consume_ai_question_quota(p_user_id uuid, p_daily_limit integer)
returns integer
language plpgsql
security definer
set search_path = public, pg_temp
as $$
declare
    used integer;
    utc_today date := (now() at time zone 'UTC')::date;
begin
    if p_daily_limit < 1 or p_daily_limit > 200 then
        raise exception 'Invalid daily AI question limit.' using errcode = '22023';
    end if;
    insert into public.daily_ai_usage (owner_id, usage_date, questions_used)
    values (p_user_id, utc_today, 0)
    on conflict (owner_id, usage_date) do nothing;
    update public.daily_ai_usage
       set questions_used = questions_used + 1
     where owner_id = p_user_id
       and usage_date = utc_today
       and questions_used < p_daily_limit
    returning questions_used into used;
    if found then
        return used;
    end if;
    select questions_used into used
      from public.daily_ai_usage
     where owner_id = p_user_id and usage_date = utc_today;
    return -used;
end;
$$;

revoke all on function public.consume_ai_question_quota(uuid, integer) from public, anon, authenticated;
grant execute on function public.consume_ai_question_quota(uuid, integer) to service_role;
