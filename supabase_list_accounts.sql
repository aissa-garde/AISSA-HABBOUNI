create or replace function public.toxgarde_list_accounts(p_token text)
returns table(doctor text, username text, role text, active boolean)
language plpgsql
security definer
set search_path = public, extensions
as $$
declare
  s record;
begin
  select * into s
  from public.toxgarde_sessions
  where token_hash=encode(digest(p_token,'sha256'),'hex')
    and expires_at>now()
  limit 1;

  if not found or s.role<>'admin' then
    raise exception 'Accès administrateur requis';
  end if;

  return query
  select a.doctor,a.username,a.role,a.active
  from public.toxgarde_accounts a
  order by case when a.role='admin' then 0 else 1 end,a.doctor,a.username;
end;
$$;

grant execute on function public.toxgarde_list_accounts(text)
to anon,authenticated;
