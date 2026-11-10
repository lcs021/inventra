-- =====================================================================
-- v106 — Back-end do Controle de Relatórios / INVENTRA
-- Rode no Supabase > SQL Editor. Pode rodar mais de uma vez (idempotente).
-- Não apaga dados, não muda colunas, não muda quem vê o quê.
--
-- 1. ctrl_definir(): grava as datas do controle (dados.ctrl) de forma atômica.
--    Antes, os dois apps liam o registro, mudavam no navegador e regravavam o
--    "dados" inteiro: se duas pessoas mexessem ao mesmo tempo, uma apagava a
--    alteração da outra. Agora o banco junta a mudança na hora de gravar.
--    O histórico (quem/quando) passa a ser carimbado pelo servidor.
-- 2. Regras de acesso (RLS) mais rápidas: as funções de permissão passam a ser
--    calculadas uma vez por consulta, não uma vez por linha.
-- 3. Alterar/apagar exigem também a permissão do módulo (mod:Categoria) —
--    ninguém altera às cegas registros de um módulo que não consegue ver.
-- 4. Índices para as consultas que os apps fazem.
-- =====================================================================

begin;  -- tudo ou nada: se algo falhar, nada muda

-- ---------- 1. Gravação atômica do controle ----------
create or replace function public.ctrl_definir(p_id uuid, p_campos jsonb)
returns jsonb
language plpgsql
security invoker              -- roda com as permissões de quem chama: o RLS continua valendo
set search_path = public
as $$
declare
  k text; v text;
  quem text := coalesce((select nullif(trim(nome), '') from public.perfis where id = auth.uid()), 'Usuário');
  agora text := to_char(now() at time zone 'utc', 'YYYY-MM-DD"T"HH24:MI:SS.MS"Z"');
  novos jsonb := '[]'::jsonb;
  res jsonb;
begin
  if auth.uid() is null then
    raise exception 'Sessão expirada' using errcode = '42501';
  end if;
  if p_campos is null or jsonb_typeof(p_campos) <> 'object' or p_campos = '{}'::jsonb then
    raise exception 'Nada para gravar' using errcode = '22023';
  end if;
  for k, v in select key, coalesce(value #>> '{}', '') from jsonb_each(p_campos) loop
    if k not in ('ent', 'ivano', 'envio', 'ret') then
      raise exception 'Campo de controle inválido: %', left(k, 40) using errcode = '22023';
    end if;
    if jsonb_typeof(p_campos -> k) not in ('string', 'null') then
      raise exception 'Data inválida' using errcode = '22023';
    end if;
    if v <> '' then
      if v !~ '^\d{4}-\d{2}-\d{2}$' then raise exception 'Data inválida' using errcode = '22023'; end if;
      perform v::date;
      if v::date > current_date + 1 then raise exception 'Data no futuro' using errcode = '22023'; end if;
    end if;
    novos := novos || jsonb_build_object('k', k, 'v', v, 'por', quem, 't', agora);
  end loop;

  -- Uma única instrução: lê e grava a linha com trava, sem janela para corrida.
  update public.ocorrencias o
     set dados = jsonb_set(
           coalesce(o.dados, '{}'::jsonb), '{ctrl}',
           (coalesce(o.dados -> 'ctrl', '{}'::jsonb)
              || (select coalesce(jsonb_object_agg(key, coalesce(value #>> '{}', '')), '{}'::jsonb) from jsonb_each(p_campos)))
           || jsonb_build_object('log',
                (select coalesce(jsonb_agg(e order by n), '[]'::jsonb) from (
                   select e, n from jsonb_array_elements(
                     coalesce(case when jsonb_typeof(o.dados -> 'ctrl' -> 'log') = 'array' then o.dados -> 'ctrl' -> 'log' end, '[]'::jsonb) || novos
                   ) with ordinality as x(e, n)
                   order by n desc limit 20) y)),
           true)
   where o.id = p_id
  returning o.dados into res;

  if res is null then
    -- não existe, é de outro galpão ou o usuário não tem permissão (o RLS esconde)
    raise exception 'Sem permissão para alterar este registro' using errcode = '42501';
  end if;
  return res;
end
$$;
revoke all on function public.ctrl_definir(uuid, jsonb) from public, anon;
grant execute on function public.ctrl_definir(uuid, jsonb) to authenticated;

-- ---------- 2. Permissões calculadas uma vez por consulta ----------
-- '*' = acesso total (admin ou Chefe de Inventário aprovado); senão, as permissões do cargo.
-- Mesma regra de tem_perm(), só que devolvida de uma vez.
create or replace function public.perms_efetivas() returns text[]
language sql stable security definer set search_path = public as $$
  select coalesce((
    select case
      when pf.admin or (pf.aprovado and pf.cargo = 'Chefe de Inventário') then array['*']
      when pf.aprovado then coalesce((select c.perms from public.cargos c where c.nome = pf.cargo), '{}')
      else '{}'::text[] end
    from public.perfis pf where pf.id = auth.uid()), '{}'::text[])
$$;
revoke all on function public.perms_efetivas() from public, anon;
grant execute on function public.perms_efetivas() to authenticated;

drop policy if exists "ver ocorrencias" on public.ocorrencias;
create policy "ver ocorrencias" on public.ocorrencias for select to authenticated
  using ((select public.is_aprovado())
     and galpao_id = (select public.meu_galpao())
     and ((select public.perms_efetivas()) && array['*', 'mod:' || coalesce(categoria, 'Relatório')]));

drop policy if exists "registrar ocorrencias" on public.ocorrencias;
create policy "registrar ocorrencias" on public.ocorrencias for insert to authenticated
  with check ((select public.is_aprovado())
     and autor = (select auth.uid())
     and galpao_id = (select public.meu_galpao())
     and (select public.tem_perm('registrar'))
     and ((select public.perms_efetivas()) && array['*', 'mod:' || coalesce(categoria, 'Relatório')]));

drop policy if exists "tratar ocorrencias" on public.ocorrencias;
create policy "tratar ocorrencias" on public.ocorrencias for update to authenticated
  using ((select public.tem_perm('garantia'))
     and galpao_id = (select public.meu_galpao())
     and ((select public.perms_efetivas()) && array['*', 'mod:' || coalesce(categoria, 'Relatório')]))
  with check ((select public.tem_perm('garantia'))
     and galpao_id = (select public.meu_galpao())
     and ((select public.perms_efetivas()) && array['*', 'mod:' || coalesce(categoria, 'Relatório')]));

drop policy if exists "apagar ocorrencias" on public.ocorrencias;
create policy "apagar ocorrencias" on public.ocorrencias for delete to authenticated
  using ((select public.tem_perm('apagar'))
     and galpao_id = (select public.meu_galpao())
     and ((select public.perms_efetivas()) && array['*', 'mod:' || coalesce(categoria, 'Relatório')]));

drop policy if exists "ver perfis" on public.perfis;
create policy "ver perfis" on public.perfis for select to authenticated
  using (id = (select auth.uid()) or ((select public.is_aprovado()) and galpao_id = (select public.meu_galpao())));
drop policy if exists "admin altera perfis" on public.perfis;
create policy "admin altera perfis" on public.perfis for update to authenticated
  using ((select public.is_admin()) and galpao_id = (select public.meu_galpao()))
  with check ((select public.is_admin()) and galpao_id = (select public.meu_galpao()));

drop policy if exists "ver setores" on public.setores;
create policy "ver setores" on public.setores for select to authenticated
  using ((select public.is_aprovado()) and galpao_id = (select public.meu_galpao()));
drop policy if exists "admin edita setores" on public.setores;
create policy "admin edita setores" on public.setores for all to authenticated
  using ((select public.is_admin()) and galpao_id = (select public.meu_galpao()))
  with check ((select public.is_admin()) and galpao_id = (select public.meu_galpao()));

-- ---------- 3. Índices ----------
-- lista do Controle/INVENTRA: por galpão e categoria, em ordem de criação
create index if not exists ocorrencias_galpao_cat_criado on public.ocorrencias (galpao_id, categoria, criado);
-- chaves estrangeiras usadas em junções e no "on delete set null"
create index if not exists ocorrencias_autor_idx on public.ocorrencias (autor);
create index if not exists perfis_galpao_idx on public.perfis (galpao_id);
create index if not exists atividade_galpao_quando on public.atividade (galpao_id, quando desc);

commit;

analyze public.ocorrencias;
analyze public.perfis;
