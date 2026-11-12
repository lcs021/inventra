-- Desfaz 0003. Se houver cargos diferentes entre empresas, mantém os da empresa mais antiga.
begin;
drop policy if exists "ver cargos" on public.cargos;
drop policy if exists "admin edita cargos" on public.cargos;
delete from public.cargos c using (select id from public.galpoes order by criado limit 1) p
  where c.galpao_id <> p.id;
alter table public.cargos drop constraint cargos_pkey;
alter table public.cargos drop column galpao_id;
alter table public.cargos add constraint cargos_pkey primary key (nome);
create policy "ver cargos" on public.cargos for select to authenticated using (true);
create policy "admin edita cargos" on public.cargos for all to authenticated
  using (public.is_admin() and public.meu_galpao() = (select id from public.galpoes order by criado limit 1))
  with check (public.is_admin() and public.meu_galpao() = (select id from public.galpoes order by criado limit 1));
create or replace function public.perms_efetivas() returns text[]
language sql stable security definer set search_path = public as $$
  select coalesce((
    select case
      when pf.admin or (pf.aprovado and pf.cargo = 'Chefe de Inventário') then array['*']
      when pf.aprovado then coalesce((select c.perms from public.cargos c where c.nome = pf.cargo), '{}')
      else '{}'::text[] end
    from public.perfis pf where pf.id = auth.uid()), '{}'::text[])
$$;
create or replace function public.audita_cargos() returns trigger
language plpgsql security definer set search_path = public as $$
begin
  perform public.audita('cargo_alterado', coalesce(new.nome, old.nome), jsonb_build_object('operacao', lower(tg_op),
    'de', case when tg_op <> 'INSERT' then to_jsonb(old.perms) end, 'para', case when tg_op <> 'DELETE' then to_jsonb(new.perms) end), null);
  return coalesce(new, old);
end $$;
grant execute on function public.meu_galpao() to anon, authenticated;
commit;
