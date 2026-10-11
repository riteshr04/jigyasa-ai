create or replace function public.refund_ai_question_quota(p_user_id uuid)
returns integer
language plpgsql
security definer
set search_path = public, pg_temp
as $$
declare
    remaining integer;
    utc_today date := (now() at time zone 'UTC')::date;
begin
    update public.daily_ai_usage
       set questions_used = questions_used - 1
     where owner_id = p_user_id
       and usage_date = utc_today
       and questions_used > 0
    returning questions_used into remaining;
    if found then
        return remaining;
    end if;
    return 0;
end;
$$;

revoke all on function public.refund_ai_question_quota(uuid) from public, anon, authenticated;
grant execute on function public.refund_ai_question_quota(uuid) to service_role;
