alter table public.documents
    drop constraint if exists documents_ext_check;
alter table public.documents
    add constraint documents_ext_check
    check (ext in ('pdf', 'pptx', 'docx', 'txt', 'mp4', 'mov', 'webm'));

update storage.buckets
   set allowed_mime_types = array[
       'application/pdf',
       'application/vnd.openxmlformats-officedocument.presentationml.presentation',
       'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
       'text/plain',
       'application/octet-stream',
       'video/mp4',
       'video/quicktime',
       'video/webm'
   ]
 where id = 'study-materials';

create table if not exists public.account_learning_state (
    owner_id uuid primary key references auth.users (id) on delete cascade,
    state jsonb not null default '{"history":[],"quizScores":[],"chatMessages":[]}'::jsonb,
    updated_at timestamptz not null default now(),
    constraint account_learning_state_shape_check check (
        jsonb_typeof(state) = 'object'
        and jsonb_typeof(state->'history') = 'array'
        and jsonb_typeof(state->'quizScores') = 'array'
        and jsonb_typeof(state->'chatMessages') = 'array'
        and octet_length(state::text) <= 1048576
    )
);

alter table public.account_learning_state enable row level security;
revoke all on public.account_learning_state from anon, authenticated;
grant all on public.account_learning_state to service_role;
