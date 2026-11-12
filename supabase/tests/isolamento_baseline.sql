-- BASELINE DE ISOLAMENTO ENTRE EMPRESAS (galpões)
-- Roda no banco LOCAL/TESTE (nunca no real). Cada teste imprime PASS ou FAIL.
-- Hoje alguns FAIL são esperados: documentam as lacunas que a Fase 3 deve corrigir.
-- Depois da Fase 3 todos devem dar PASS.
\set ON_ERROR_STOP off
\pset format unaligned
\pset tuples_only on

begin;
-- Duas empresas, cada uma com um admin e um auxiliar aprovados
insert into auth.users(id,email) values
 ('a0000000-0000-0000-0000-00000000000a','admin.a@t.local'),
 ('a0000000-0000-0000-0000-0000000000a1','aux.a@t.local'),
 ('b0000000-0000-0000-0000-00000000000b','admin.b@t.local'),
 ('b0000000-0000-0000-0000-0000000000b1','aux.b@t.local');
insert into public.galpoes(id,nome,criado) values
 ('aaaaaaaa-0000-0000-0000-000000000001','Empresa A', now()-interval '2 day'),
 ('bbbbbbbb-0000-0000-0000-000000000002','Empresa B', now()-interval '1 day');
delete from public.perfis;
insert into public.perfis(id,nome,email,cargo,aprovado,admin,usuario,galpao_id) values
 ('a0000000-0000-0000-0000-00000000000a','Admin A','admin.a@t.local','Chefe de Inventário',true,true,'admin_a','aaaaaaaa-0000-0000-0000-000000000001'),
 ('a0000000-0000-0000-0000-0000000000a1','Aux A','aux.a@t.local','Auxiliar de Inventário',true,false,'aux_a','aaaaaaaa-0000-0000-0000-000000000001'),
 ('b0000000-0000-0000-0000-00000000000b','Admin B','admin.b@t.local','Chefe de Inventário',true,true,'admin_b','bbbbbbbb-0000-0000-0000-000000000002'),
 ('b0000000-0000-0000-0000-0000000000b1','Aux B','aux.b@t.local','Auxiliar de Inventário',true,false,'aux_b','bbbbbbbb-0000-0000-0000-000000000002');
insert into public.ocorrencias(id,tipo,endereco,autor,galpao_id) values
 ('a1000000-0000-0000-0000-000000000001','Falta','A-01','a0000000-0000-0000-0000-0000000000a1','aaaaaaaa-0000-0000-0000-000000000001'),
 ('b1000000-0000-0000-0000-000000000001','Falta','B-01','b0000000-0000-0000-0000-0000000000b1','bbbbbbbb-0000-0000-0000-000000000002');

create temp table resultado(teste text, ok boolean, detalhe text);
grant all on resultado to public;

-- helper: executa SQL como um usuário e devolve o texto do resultado/erro
create or replace function pg_temp.como(uid text, role_ text, sql_ text) returns text
language plpgsql as $$
declare r text;
begin
  perform set_config('request.jwt.claims', case when uid is null then '' else json_build_object('sub',uid,'role',role_)::text end, true);
  execute 'set local role '||role_;
  begin
    execute sql_ into r;
  exception when others then r := 'ERRO: '||sqlerrm;
  end;
  execute 'reset role';
  return r;
end $$;

-- T1: auxiliar da A não enxerga ocorrência da B
insert into resultado select 'T1 ocorrências: A não vê dados de B',
  pg_temp.como('a0000000-0000-0000-0000-0000000000a1','authenticated',
    $q$select count(*)::text from public.ocorrencias where galpao_id='bbbbbbbb-0000-0000-0000-000000000002'$q$)='0', null;

-- T2: auxiliar da A não consegue inserir ocorrência na B
insert into resultado select 'T2 ocorrências: A não grava na B',
  pg_temp.como('a0000000-0000-0000-0000-0000000000a1','authenticated',
    $q$with i as (insert into public.ocorrencias(tipo,endereco,autor,galpao_id) values ('Falta','X','a0000000-0000-0000-0000-0000000000a1','bbbbbbbb-0000-0000-0000-000000000002') returning 1) select count(*)::text from i$q$) like 'ERRO%', null;

-- T3: admin da A não altera perfil de usuário da B
insert into resultado select 'T3 perfis: admin A não altera usuário de B',
  pg_temp.como('a0000000-0000-0000-0000-00000000000a','authenticated',
    $q$with u as (update public.perfis set nome='HACK' where id='b0000000-0000-0000-0000-0000000000b1' returning 1) select count(*)::text from u$q$)='0', null;

-- T4: admin da A não vê setores da B (setores são por galpão)
insert into resultado select 'T4 setores: A não vê setores de B',
  pg_temp.como('a0000000-0000-0000-0000-00000000000a','authenticated',
    $q$select count(*)::text from public.setores where galpao_id='bbbbbbbb-0000-0000-0000-000000000002'$q$)='0', null;

-- T5 (LACUNA GRAVE): a policy 'admin edita cargos' compara com o galpão mais antigo, mas a subconsulta roda sob RLS
-- e só enxerga o próprio galpão => qualquer admin de qualquer empresa edita os cargos globais.
insert into public.cargos(nome,perms,galpao_id) select 'Conferente', array['ver'], 'aaaaaaaa-0000-0000-0000-000000000001'
  where exists (select 1 from information_schema.columns where table_schema='public' and table_name='cargos' and column_name='galpao_id');
insert into resultado select 'T5 cargos: admin de B não altera cargos da empresa A',
  case when exists (select 1 from information_schema.columns where table_schema='public' and table_name='cargos' and column_name='galpao_id')
   then pg_temp.como('b0000000-0000-0000-0000-00000000000b','authenticated',
    $q$with u as (update public.cargos set perms='{}' where nome='Conferente' returning 1) select count(*)::text from u$q$)='0'
   else pg_temp.como('b0000000-0000-0000-0000-00000000000b','authenticated',
    $q$with u as (update public.cargos set perms='{}' returning 1) select count(*)::text from u$q$)='0' end, null;
insert into resultado select 'T5b cargos: admin de A altera os cargos da própria empresa',
  pg_temp.como('a0000000-0000-0000-0000-00000000000a','authenticated',
    $q$with u as (update public.cargos set perms=array['ver','x'] where nome='Conferente' returning 1) select count(*)::text from u$q$)='1', null;
insert into resultado select 'T5c cargos: usuário da B não enxerga cargos da A',
  pg_temp.como('b0000000-0000-0000-0000-0000000000b1','authenticated',$q$select count(*)::text from public.cargos where nome='Conferente'$q$)='0', null;

-- T6 (LACUNA estrutural): cargos precisam pertencer a uma empresa (hoje a tabela é global, sem galpao_id)
insert into resultado select 'T6 cargos: tabela tem coluna de empresa',
  exists (select 1 from information_schema.columns where table_schema='public' and table_name='cargos' and column_name in ('galpao_id','organizacao_id')),
  'cargos é global';

-- T7 (LACUNA): visitante anônimo não deve listar empresas
insert into resultado select 'T7 lista_galpoes: anônimo não lista empresas',
  pg_temp.como(null,'anon',$q$select count(*)::text from public.lista_galpoes()$q$) in ('0') or
  pg_temp.como(null,'anon',$q$select count(*)::text from public.lista_galpoes()$q$) like 'ERRO%', null;

-- T8 (LACUNA): nome de usuário só precisa ser único dentro da empresa (hoje é único no sistema todo)
insert into resultado select 'T8 usuário: mesmo login permitido em empresas diferentes',
  pg_temp.como(null,'postgres',$q$with i as (update public.perfis set usuario='admin_a' where id='b0000000-0000-0000-0000-0000000000b1' returning 1) select count(*)::text from i$q$)='1', 'hoje: índice perfis_usuario_unico global';

-- T9 (LACUNA): anônimo não executa meu_galpao()
insert into resultado select 'T9 meu_galpao: anônimo sem acesso',
  not has_function_privilege('anon','public.meu_galpao()','execute'), null;

-- T10: tabela ajustes (global) não deve ser lida por usuário comum
insert into resultado select 'T10 ajustes: usuário comum não lê',
  pg_temp.como('a0000000-0000-0000-0000-0000000000a1','authenticated',$q$select count(*)::text from public.ajustes$q$) in ('0') or
  pg_temp.como('a0000000-0000-0000-0000-0000000000a1','authenticated',$q$select count(*)::text from public.ajustes$q$) like 'ERRO%', null;

-- T11: auxiliar não vira admin sozinho
insert into resultado select 'T11 perfis: auxiliar não se promove a admin',
  pg_temp.como('a0000000-0000-0000-0000-0000000000a1','authenticated',
    $q$with u as (update public.perfis set admin=true where id='a0000000-0000-0000-0000-0000000000a1' returning 1) select count(*)::text from u$q$)<>'1'
  or (select not admin from public.perfis where id='a0000000-0000-0000-0000-0000000000a1'), null;

-- T12: troca de galpao_id por auxiliar bloqueada
insert into resultado select 'T12 perfis: auxiliar não muda de empresa',
  (pg_temp.como('a0000000-0000-0000-0000-0000000000a1','authenticated',
    $q$with u as (update public.perfis set galpao_id='bbbbbbbb-0000-0000-0000-000000000002' where id='a0000000-0000-0000-0000-0000000000a1' returning 1) select count(*)::text from u$q$)<>'1')
  or (select galpao_id='aaaaaaaa-0000-0000-0000-000000000001' from public.perfis where id='a0000000-0000-0000-0000-0000000000a1'), null;

select case when ok then 'PASS' else 'FAIL' end||'  '||teste||coalesce('  ('||detalhe||')','') from resultado order by teste;
select 'RESUMO: '||count(*) filter (where ok)||' PASS / '||count(*) filter (where not ok)||' FAIL' from resultado;
rollback;
