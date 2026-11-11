-- DIAGNÓSTICO SOMENTE LEITURA — INVENTRA / Controle de Relatórios
-- Cole no SQL Editor do Supabase (projeto REAL) e execute. Não altera nada:
-- é um único SELECT. O resultado é uma tabela (secao | item | detalhe); copie/cole de volta.
-- Não retorna senhas, e-mails nem dados de ocorrências — só a estrutura e contagens.

with
esperadas_func(nome) as (values
  ('meu_galpao'),('is_aprovado'),('is_admin'),('tem_perm'),('tem_cargo'),('perms_efetivas'),
  ('ctrl_definir'),('lista_galpoes'),('usuario_existe'),('login_email'),('registrar_evento'),
  ('admin_redefinir_senha'),('excluir_usuario'),('audita'),('pasta_do_meu_galpao')),
esperadas_tab(nome) as (values
  ('galpoes'),('perfis'),('ocorrencias'),('setores'),('cargos'),('atividade'),
  ('auditoria'),('login_falhas'),('ajustes'))

select '1. tabelas' as secao, e.nome as item,
       case when t.tablename is null then 'AUSENTE'
            when t.rowsecurity then 'existe, RLS ligada'
            else 'existe, RLS DESLIGADA  <-- risco' end as detalhe
from esperadas_tab e
left join pg_tables t on t.schemaname='public' and t.tablename=e.nome

union all
select '2. funções', e.nome,
       case when count(p.oid)=0 then 'AUSENTE (migração não aplicada?)'
            else 'existe'
              || case when bool_or(p.prosecdef) then ', security definer' else '' end
              || case when bool_or(has_function_privilege('anon', p.oid, 'execute')) then ', EXECUTÁVEL POR ANON' else ', anon bloqueado' end
       end
from esperadas_func e
left join pg_proc p on p.proname=e.nome and p.pronamespace='public'::regnamespace
group by e.nome

union all
select '3. policies', tablename||' / '||policyname, cmd||' para '||array_to_string(roles,',')
from pg_policies where schemaname='public'

union all
select '4. privilégios do anon em tabelas', table_name, string_agg(privilege_type, ',' order by privilege_type)
from information_schema.role_table_grants
where table_schema='public' and grantee='anon' group by table_name

union all
select '5. colunas galpao_id', t.table_name,
       case when c.column_name is null then 'sem galpao_id' else 'tem galpao_id ('||c.is_nullable||' nulo)' end
from information_schema.tables t
left join information_schema.columns c
  on c.table_schema=t.table_schema and c.table_name=t.table_name and c.column_name='galpao_id'
where t.table_schema='public' and t.table_type='BASE TABLE'

union all
select '6. contagens', x.item, x.n::text from (
  select 'galpoes' item, count(*) n from public.galpoes
  union all select 'perfis', count(*) from public.perfis
  union all select 'perfis sem galpao_id', count(*) from public.perfis where galpao_id is null
  union all select 'ocorrencias', count(*) from public.ocorrencias
  union all select 'ocorrencias sem galpao_id', count(*) from public.ocorrencias where galpao_id is null
  union all select 'setores', count(*) from public.setores
  union all select 'cargos', count(*) from public.cargos
) x

union all
select '7. unicidade', i.indexname, i.indexdef
from pg_indexes i where i.schemaname='public' and i.tablename='perfis'

union all
select '8. triggers', event_object_table||' / '||trigger_name, action_timing||' '||event_manipulation
from information_schema.triggers where trigger_schema='public'

union all
select '9. storage', 'bucket '||id, case when public then 'PÚBLICO <-- risco' else 'privado' end
from storage.buckets

union all
select '10. extensões', extname, extversion from pg_extension

order by 1, 2;
