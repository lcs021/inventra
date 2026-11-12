-- 0003 — Cargos passam a pertencer a cada empresa (galpão) + fecha execução anônima de meu_galpao()
-- Corrige: qualquer admin de qualquer empresa editava os cargos (policy lia galpoes sob RLS).
-- Reversível: rollback/0003_down.sql. Rode primeiro no projeto de TESTE e faça backup antes do real.
-- Compatibilidade: o site precisa gravar cargos com onConflict 'galpao_id,nome' (ver README, Fase 3).
begin;

-- 1. coluna de empresa; cópia dos cargos atuais para cada galpão existente
alter table public.cargos add column if not exists galpao_id uuid references public.galpoes(id) on delete cascade;
alter table public.cargos drop constraint if exists cargos_pkey;

insert into public.cargos (nome, perms, galpao_id)
select c.nome, c.perms, g.id
from public.cargos c cross join public.galpoes g
where c.galpao_id is null;

delete from public.cargos where galpao_id is null;

alter table public.cargos alter column galpao_id set not null;
alter table public.cargos alter column galpao_id set default public.meu_galpao();
alter table public.cargos add constraint cargos_pkey primary key (galpao_id, nome);

-- 2. policies: cada empresa vê e edita só os próprios cargos (edita = admin da empresa)
drop policy if exists "ver cargos" on public.cargos;
drop policy if exists "admin edita cargos" on public.cargos;
create policy "ver cargos" on public.cargos for select to authenticated
  using (galpao_id = (select public.meu_galpao()));
create policy "admin edita cargos" on public.cargos for all to authenticated
  using ((select public.is_admin()) and galpao_id = (select public.meu_galpao()))
  with check ((select public.is_admin()) and galpao_id = (select public.meu_galpao()));

-- 3. permissões efetivas leem o cargo da empresa do próprio usuário
create or replace function public.perms_efetivas() returns text[]
language sql stable security definer set search_path = public as $$
  select coalesce((
    select case
      when pf.admin or (pf.aprovado and pf.cargo = 'Chefe de Inventário') then array['*']
      when pf.aprovado then coalesce((select c.perms from public.cargos c
                                       where c.nome = pf.cargo and c.galpao_id = pf.galpao_id), '{}')
      else '{}'::text[] end
    from public.perfis pf where pf.id = auth.uid()), '{}'::text[])
$$;

-- 4. auditoria do cargo registra a empresa
create or replace function public.audita_cargos() returns trigger
language plpgsql security definer set search_path = public as $$
begin
  perform public.audita('cargo_alterado', coalesce(new.nome, old.nome), jsonb_build_object('operacao', lower(tg_op),
    'de', case when tg_op <> 'INSERT' then to_jsonb(old.perms) end, 'para', case when tg_op <> 'DELETE' then to_jsonb(new.perms) end),
    coalesce(new.galpao_id, old.galpao_id));
  return coalesce(new, old);
end $$;

-- 5. visitante sem login não executa meu_galpao()
revoke execute on function public.meu_galpao() from public, anon;
grant execute on function public.meu_galpao() to authenticated;

commit;
