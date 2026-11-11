-- Inventário CD: estrutura do banco (rode uma vez no SQL Editor do Supabase)

-- Tabelas
create table if not exists public.perfis (
  id uuid primary key references auth.users(id) on delete cascade,
  nome text not null default '',
  email text,
  cargo text not null default 'Auxiliar de Inventário',
  aprovado boolean not null default false,
  admin boolean not null default false,
  criado timestamptz not null default now()
);
alter table public.perfis add column if not exists recusado boolean not null default false;
alter table public.perfis add column if not exists usuario text;
create unique index if not exists perfis_usuario_unico on public.perfis (lower(usuario));

create table if not exists public.ocorrencias (
  id uuid primary key default gen_random_uuid(),
  tipo text not null,
  endereco text not null,
  sku text,
  produto text,
  fab text,
  qtd integer,
  obs text,
  fotos text[] not null default '{}',
  videos text[] not null default '{}',
  autor uuid references public.perfis(id) on delete set null,
  status text not null default 'aberta',
  hist jsonb not null default '[]'::jsonb,
  criado timestamptz not null default now()
);

-- Categoria do registro (Garantia ou Relatório)
alter table public.ocorrencias add column if not exists categoria text not null default 'Relatório';
alter table public.ocorrencias add column if not exists fornecedor text;
alter table public.ocorrencias add column if not exists id_quebra text;
alter table public.ocorrencias add column if not exists gar_status text;
alter table public.ocorrencias add column if not exists origem text;

alter table public.perfis enable row level security;
alter table public.ocorrencias enable row level security;

-- Funções de permissão
create or replace function public.is_aprovado() returns boolean
language sql stable security definer set search_path = public as $$
  select coalesce((select aprovado or admin from public.perfis where id = auth.uid()), false)
$$;

create or replace function public.is_admin() returns boolean
language sql stable security definer set search_path = public as $$
  select coalesce((select admin from public.perfis where id = auth.uid()), false)
$$;

create or replace function public.tem_cargo(cargos text[]) returns boolean
language sql stable security definer set search_path = public as $$
  select coalesce((select admin or (aprovado and cargo = any(cargos)) from public.perfis where id = auth.uid()), false)
$$;


-- Permissões por cargo (o Chefe de Inventário pode tudo; o Auxiliar pode o que o administrador liberar)
create table if not exists public.cargos (
  nome text primary key,
  perms text[] not null default '{}'
);
alter table public.cargos enable row level security;
drop policy if exists "ver cargos" on public.cargos;
create policy "ver cargos" on public.cargos for select to authenticated using (true);
drop policy if exists "admin edita cargos" on public.cargos;
create policy "admin edita cargos" on public.cargos for all to authenticated using (public.is_admin()) with check (public.is_admin());
insert into public.cargos (nome, perms) values
  ('Auxiliar de Inventário', array['mod:Garantia','mod:Relatório','mod:Minucioso','mod:Consertos Baiano','registrar','resumo','garantia'])
on conflict (nome) do nothing;

create or replace function public.tem_perm(p text) returns boolean
language sql stable security definer set search_path = public as $$
  select coalesce((
    select pf.admin or (pf.aprovado and (pf.cargo = 'Chefe de Inventário'
      or exists (select 1 from public.cargos c where c.nome = pf.cargo and p = any(c.perms))))
    from public.perfis pf where pf.id = auth.uid()), false)
$$;

-- Ao criar uma conta, cria o perfil. A primeira conta vira administrador.
create or replace function public.novo_usuario() returns trigger
language plpgsql security definer set search_path = public as $$
declare primeiro boolean;
begin
  select not exists (select 1 from public.perfis) into primeiro;
  insert into public.perfis (id, nome, email, cargo, aprovado, admin, usuario)
  values (new.id, coalesce(new.raw_user_meta_data->>'nome', ''), new.email,
          case when primeiro then 'Chefe de Inventário' else 'Auxiliar de Inventário' end, primeiro, primeiro,
          nullif(regexp_replace(trim(coalesce(new.raw_user_meta_data->>'usuario', '')), '\s+', ' ', 'g'), ''));
  return new;
end
$$;

drop trigger if exists ao_criar_usuario on auth.users;
create trigger ao_criar_usuario after insert on auth.users
  for each row execute function public.novo_usuario();

-- Regras de acesso: perfis
drop policy if exists "ver perfis" on public.perfis;
create policy "ver perfis" on public.perfis for select to authenticated
  using (id = auth.uid() or public.is_aprovado());
drop policy if exists "admin altera perfis" on public.perfis;
create policy "admin altera perfis" on public.perfis for update to authenticated
  using (public.is_admin()) with check (public.is_admin());

-- Regras de acesso: ocorrências
drop policy if exists "ver ocorrencias" on public.ocorrencias;
create policy "ver ocorrencias" on public.ocorrencias for select to authenticated
  using (public.is_aprovado() and public.tem_perm('mod:' || coalesce(categoria, 'Relatório')));
drop policy if exists "registrar ocorrencias" on public.ocorrencias;
create policy "registrar ocorrencias" on public.ocorrencias for insert to authenticated
  with check (public.is_aprovado() and autor = auth.uid() and public.tem_perm('registrar') and public.tem_perm('mod:' || coalesce(categoria, 'Relatório')));
drop policy if exists "tratar ocorrencias" on public.ocorrencias;
create policy "tratar ocorrencias" on public.ocorrencias for update to authenticated
  using (public.tem_perm('garantia'))
  with check (public.tem_perm('garantia'));
drop policy if exists "apagar ocorrencias" on public.ocorrencias;
create policy "apagar ocorrencias" on public.ocorrencias for delete to authenticated
  using (public.tem_perm('apagar'));

-- Fotos e vídeos (pasta privada, até 50 MB por arquivo)
insert into storage.buckets (id, name, public, file_size_limit)
values ('midias', 'midias', false, 52428800)
on conflict (id) do nothing;

drop policy if exists "ver midias" on storage.objects;
create policy "ver midias" on storage.objects for select to authenticated
  using (bucket_id = 'midias' and public.is_aprovado());
drop policy if exists "enviar midias" on storage.objects;
create policy "enviar midias" on storage.objects for insert to authenticated
  with check (bucket_id = 'midias' and public.is_aprovado());
drop policy if exists "apagar midias" on storage.objects;
create policy "apagar midias" on storage.objects for delete to authenticated
  using (bucket_id = 'midias' and (owner_id = auth.uid()::text or public.tem_perm('apagar')));

-- Excluir usuário não liberado (somente administrador)
create or replace function public.excluir_usuario(alvo uuid) returns void
language plpgsql security definer set search_path = public as $$
begin
  if not public.is_admin() then raise exception 'Apenas o administrador pode excluir usuários'; end if;
  if alvo = auth.uid() then raise exception 'Você não pode excluir a si mesmo'; end if;
  if exists (select 1 from public.perfis where id = alvo and (admin or aprovado)) then
    raise exception 'Só é possível excluir usuários não liberados';
  end if;
  delete from auth.users where id = alvo;
end
$$;
revoke all on function public.excluir_usuario(uuid) from public, anon;
grant execute on function public.excluir_usuario(uuid) to authenticated;

-- Minucioso: setores e contagem
alter table public.ocorrencias add column if not exists setor text;
alter table public.ocorrencias add column if not exists chefe text;
alter table public.ocorrencias add column if not exists qtd_sistema integer;
alter table public.ocorrencias add column if not exists qtd_contada integer;

create table if not exists public.setores (
  nome text primary key,
  chefe text,
  ordem integer not null default 0
);
alter table public.setores enable row level security;
drop policy if exists "ver setores" on public.setores;
create policy "ver setores" on public.setores for select to authenticated using (public.is_aprovado());
drop policy if exists "admin edita setores" on public.setores;
create policy "admin edita setores" on public.setores for all to authenticated using (public.is_admin()) with check (public.is_admin());


-- Minucioso: ficha de falta de produtos
alter table public.ocorrencias add column if not exists custo numeric;
alter table public.ocorrencias add column if not exists falta numeric;
alter table public.ocorrencias add column if not exists motivo_chefe text;
alter table public.ocorrencias add column if not exists motivo_inventario text;
alter table public.ocorrencias add column if not exists motivo_allan text;

-- Login por usuário: descobre o e-mail a partir do nome de usuário
create or replace function public.email_do_usuario(u text) returns text
language sql stable security definer set search_path = public as $$
  select email from public.perfis
  where lower(usuario) = lower(regexp_replace(trim(u), '\s+', ' ', 'g'))
  limit 1
$$;
grant execute on function public.email_do_usuario(text) to anon, authenticated;

-- Usuário do administrador (mesma senha de antes)
update public.perfis set usuario = 'Lucas Carvalho'
where id = (select id from public.perfis where admin order by criado limit 1)
  and not exists (select 1 from public.perfis where lower(usuario) = 'lucas carvalho');

-- Cadastro sem e-mail: gera um endereço interno (caixa do administrador com +código)
create or replace function public.email_novo_usuario(u text) returns text
language sql stable security definer set search_path = public, auth as $$
  select split_part(x.e, '@', 1) || '+u' || substr(md5(lower(regexp_replace(trim(u), '\s+', ' ', 'g'))), 1, 10) || '@' || split_part(x.e, '@', 2)
  from (
    select coalesce(nullif(p.email, ''), au.email) as e, p.criado
    from public.perfis p join auth.users au on au.id = p.id
    where p.admin
    order by p.criado
    limit 1
  ) x
  where x.e like '%@%'
$$;
grant execute on function public.email_novo_usuario(text) to anon, authenticated;

-- Consertos Baiano
alter table public.ocorrencias add column if not exists requisitante text;
alter table public.ocorrencias add column if not exists valor_conserto numeric;
alter table public.ocorrencias add column if not exists data_ref date;

-- Cargos: Chefe de Inventário e Auxiliar de Inventário
alter table public.perfis alter column cargo set default 'Auxiliar de Inventário';
update public.perfis set cargo = 'Auxiliar de Inventário' where cargo not in ('Chefe de Inventário','Auxiliar de Inventário','Visitante');

-- Administrador redefine a senha de um usuário (a senha antiga não é visível; fica guardada só em forma criptografada)
create extension if not exists pgcrypto with schema extensions;
create or replace function public.admin_redefinir_senha(alvo uuid, nova text) returns void
language plpgsql security definer set search_path = public, extensions, auth as $$
begin
  if not public.is_admin() then raise exception 'Apenas o administrador pode redefinir senhas'; end if;
  if length(coalesce(nova, '')) < 6 then raise exception 'A senha precisa ter pelo menos 6 caracteres'; end if;
  update auth.users set encrypted_password = extensions.crypt(nova, extensions.gen_salt('bf')), updated_at = now() where id = alvo;
  if not found then raise exception 'Usuário não encontrado'; end if;
end
$$;
revoke all on function public.admin_redefinir_senha(uuid, text) from public, anon;
grant execute on function public.admin_redefinir_senha(uuid, text) to authenticated;

-- Relatório de Avaria
alter table public.ocorrencias add column if not exists dados jsonb;

-- Log de atividade (quem criou, editou ou apagou)
create table if not exists public.atividade (
  id bigint generated always as identity primary key,
  quando timestamptz not null default now(),
  usuario uuid default auth.uid(),
  acao text not null,
  categoria text,
  ref_id text,
  resumo text
);
create index if not exists atividade_quando_idx on public.atividade (quando desc);
alter table public.atividade enable row level security;
drop policy if exists "grava atividade" on public.atividade;
drop policy if exists "admin ve atividade" on public.atividade;
create policy "grava atividade" on public.atividade for insert to authenticated with check (public.is_aprovado() and usuario = auth.uid());
create policy "admin ve atividade" on public.atividade for select to authenticated using (public.is_admin());

-- =====================================================================
-- Vários galpões no mesmo sistema (cada galpão só enxerga os próprios dados)
-- Para criar um galpão novo:  insert into public.galpoes (nome) values ('Nome do galpão');
-- A primeira pessoa que se cadastrar nele escolhendo esse galpão vira Chefe (administrador).
-- =====================================================================
create table if not exists public.galpoes (
  id uuid primary key default gen_random_uuid(),
  nome text not null,
  criado timestamptz not null default now()
);
create unique index if not exists galpoes_nome_unico on public.galpoes (lower(nome));
insert into public.galpoes (nome) select 'Galpão principal' where not exists (select 1 from public.galpoes);
alter table public.galpoes enable row level security;

alter table public.perfis add column if not exists galpao_id uuid references public.galpoes(id);
alter table public.ocorrencias add column if not exists galpao_id uuid references public.galpoes(id);
alter table public.setores add column if not exists galpao_id uuid references public.galpoes(id);
alter table public.atividade add column if not exists galpao_id uuid references public.galpoes(id);

-- tudo o que já existe passa a pertencer ao primeiro galpão
update public.perfis set galpao_id = (select id from public.galpoes order by criado limit 1) where galpao_id is null;
update public.ocorrencias set galpao_id = (select id from public.galpoes order by criado limit 1) where galpao_id is null;
update public.setores set galpao_id = (select id from public.galpoes order by criado limit 1) where galpao_id is null;
update public.atividade set galpao_id = (select id from public.galpoes order by criado limit 1) where galpao_id is null;

create or replace function public.meu_galpao() returns uuid
language sql stable security definer set search_path = public as $$
  select galpao_id from public.perfis where id = auth.uid()
$$;

alter table public.ocorrencias alter column galpao_id set default public.meu_galpao();
alter table public.setores alter column galpao_id set default public.meu_galpao();
alter table public.atividade alter column galpao_id set default public.meu_galpao();
alter table public.setores alter column galpao_id set not null;
alter table public.setores drop constraint if exists setores_pkey;
alter table public.setores add primary key (galpao_id, nome);

-- lista de galpões para a tela de cadastro (só nome)
create or replace function public.lista_galpoes() returns table (id uuid, nome text)
language sql stable security definer set search_path = public as $$
  select g.id, g.nome from public.galpoes g order by g.nome
$$;
grant execute on function public.lista_galpoes() to anon, authenticated;

drop policy if exists "ver galpao" on public.galpoes;
create policy "ver galpao" on public.galpoes for select to authenticated using (id = public.meu_galpao());

-- Novo usuário: entra no galpão escolhido; o primeiro de cada galpão vira Chefe
create or replace function public.novo_usuario() returns trigger
language plpgsql security definer set search_path = public as $$
declare g uuid; primeiro boolean;
begin
  begin g := nullif(new.raw_user_meta_data->>'galpao', '')::uuid; exception when others then g := null; end;
  if g is null or not exists (select 1 from public.galpoes where id = g) then
    select id into g from public.galpoes order by criado limit 1;
  end if;
  select not exists (select 1 from public.perfis where galpao_id = g) into primeiro;
  insert into public.perfis (id, nome, email, cargo, aprovado, admin, usuario, galpao_id)
  values (new.id, coalesce(new.raw_user_meta_data->>'nome', ''), new.email,
          case when primeiro then 'Chefe de Inventário' else 'Auxiliar de Inventário' end, primeiro, primeiro,
          nullif(regexp_replace(trim(coalesce(new.raw_user_meta_data->>'usuario', '')), '\s+', ' ', 'g'), ''), g);
  return new;
end
$$;

-- Regras de acesso, agora separadas por galpão
drop policy if exists "ver perfis" on public.perfis;
create policy "ver perfis" on public.perfis for select to authenticated
  using (id = auth.uid() or (public.is_aprovado() and galpao_id = public.meu_galpao()));
drop policy if exists "admin altera perfis" on public.perfis;
create policy "admin altera perfis" on public.perfis for update to authenticated
  using (public.is_admin() and galpao_id = public.meu_galpao())
  with check (public.is_admin() and galpao_id = public.meu_galpao());

drop policy if exists "ver ocorrencias" on public.ocorrencias;
create policy "ver ocorrencias" on public.ocorrencias for select to authenticated
  using (public.is_aprovado() and galpao_id = public.meu_galpao() and public.tem_perm('mod:' || coalesce(categoria, 'Relatório')));
drop policy if exists "registrar ocorrencias" on public.ocorrencias;
create policy "registrar ocorrencias" on public.ocorrencias for insert to authenticated
  with check (public.is_aprovado() and autor = auth.uid() and galpao_id = public.meu_galpao() and public.tem_perm('registrar') and public.tem_perm('mod:' || coalesce(categoria, 'Relatório')));
drop policy if exists "tratar ocorrencias" on public.ocorrencias;
create policy "tratar ocorrencias" on public.ocorrencias for update to authenticated
  using (public.tem_perm('garantia') and galpao_id = public.meu_galpao())
  with check (public.tem_perm('garantia') and galpao_id = public.meu_galpao());
drop policy if exists "apagar ocorrencias" on public.ocorrencias;
create policy "apagar ocorrencias" on public.ocorrencias for delete to authenticated
  using (public.tem_perm('apagar') and galpao_id = public.meu_galpao());

drop policy if exists "ver setores" on public.setores;
create policy "ver setores" on public.setores for select to authenticated
  using (public.is_aprovado() and galpao_id = public.meu_galpao());
drop policy if exists "admin edita setores" on public.setores;
create policy "admin edita setores" on public.setores for all to authenticated
  using (public.is_admin() and galpao_id = public.meu_galpao())
  with check (public.is_admin() and galpao_id = public.meu_galpao());

drop policy if exists "grava atividade" on public.atividade;
create policy "grava atividade" on public.atividade for insert to authenticated
  with check (public.is_aprovado() and usuario = auth.uid() and galpao_id = public.meu_galpao());
drop policy if exists "admin ve atividade" on public.atividade;
create policy "admin ve atividade" on public.atividade for select to authenticated
  using (public.is_admin() and galpao_id = public.meu_galpao());

-- O administrador só mexe em usuários do próprio galpão
create or replace function public.excluir_usuario(alvo uuid) returns void
language plpgsql security definer set search_path = public as $$
begin
  if not public.is_admin() then raise exception 'Apenas o administrador pode excluir usuários'; end if;
  if alvo = auth.uid() then raise exception 'Você não pode excluir a si mesmo'; end if;
  if not exists (select 1 from public.perfis where id = alvo and galpao_id = public.meu_galpao()) then
    raise exception 'Usuário de outro galpão';
  end if;
  if exists (select 1 from public.perfis where id = alvo and (admin or aprovado)) then
    raise exception 'Só é possível excluir usuários não liberados';
  end if;
  delete from auth.users where id = alvo;
end
$$;
create or replace function public.admin_redefinir_senha(alvo uuid, nova text) returns void
language plpgsql security definer set search_path = public, extensions, auth as $$
begin
  if not public.is_admin() then raise exception 'Apenas o administrador pode redefinir senhas'; end if;
  if length(coalesce(nova, '')) < 6 then raise exception 'A senha precisa ter pelo menos 6 caracteres'; end if;
  if not exists (select 1 from public.perfis where id = alvo and galpao_id = public.meu_galpao()) then
    raise exception 'Usuário de outro galpão';
  end if;
  update auth.users set encrypted_password = extensions.crypt(nova, extensions.gen_salt('bf')), updated_at = now() where id = alvo;
  if not found then raise exception 'Usuário não encontrado'; end if;
end
$$;

-- Setores iniciais do Minucioso (só se a lista estiver vazia)
insert into public.setores (galpao_id, nome, ordem)
select (select id from public.galpoes order by criado limit 1), v.nome, v.ordem
from (values ('Expedição',0),('Shopee',1),('Mercado Livre',2),('Cofre',3),('1º andar',4),
  ('2º andar',5),('3º andar',6),('Carga e descarga',7),('Recebimento',8),('Parachoque',9)) as v(nome, ordem)
where not exists (select 1 from public.setores);

-- ===== Segurança (v103) =====
-- 1. Ninguém muda autor, data de criação ou galpão de um registro depois de criado
create or replace function public.protege_ocorrencia() returns trigger
language plpgsql as $$
begin
  if (old.autor is not null and new.autor is distinct from old.autor and new.autor is not null)
     or new.criado is distinct from old.criado
     or new.galpao_id is distinct from old.galpao_id then
    raise exception 'Autor, data de criação e galpão do registro não podem ser alterados';
  end if;
  return new;
end $$;
drop trigger if exists protege_ocorrencia on public.ocorrencias;
create trigger protege_ocorrencia before update on public.ocorrencias
  for each row execute function public.protege_ocorrencia();

-- 2. Perfis: identidade fixa, e o administrador não tira o próprio acesso sem querer
create or replace function public.protege_perfil() returns trigger
language plpgsql as $$
begin
  if new.id is distinct from old.id or new.criado is distinct from old.criado
     or new.galpao_id is distinct from old.galpao_id or new.email is distinct from old.email then
    raise exception 'Identidade, e-mail e galpão do usuário não podem ser alterados';
  end if;
  if auth.uid() is not null and old.id = auth.uid()
     and (new.admin is distinct from old.admin or new.aprovado is distinct from old.aprovado) then
    raise exception 'Você não pode remover o seu próprio acesso de administrador';
  end if;
  return new;
end $$;
drop trigger if exists protege_perfil on public.perfis;
create trigger protege_perfil before update on public.perfis
  for each row execute function public.protege_perfil();

-- 3. Senhas redefinidas pelo administrador: mínimo de 8 caracteres
create or replace function public.admin_redefinir_senha(alvo uuid, nova text) returns void
language plpgsql security definer set search_path = public, extensions, auth as $$
begin
  if not public.is_admin() then raise exception 'Apenas o administrador pode redefinir senhas'; end if;
  if length(coalesce(nova, '')) < 8 then raise exception 'A senha precisa ter pelo menos 8 caracteres'; end if;
  if not exists (select 1 from public.perfis where id = alvo and galpao_id = public.meu_galpao()) then
    raise exception 'Usuário de outro galpão';
  end if;
  update auth.users set encrypted_password = extensions.crypt(nova, extensions.gen_salt('bf')), updated_at = now() where id = alvo;
  if not found then raise exception 'Usuário não encontrado'; end if;
end
$$;
revoke all on function public.admin_redefinir_senha(uuid, text) from public, anon;
grant execute on function public.admin_redefinir_senha(uuid, text) to authenticated;

-- 4. Fotos e vídeos: só imagem/vídeo, cada pessoa envia para a própria pasta,
--    e só o mesmo galpão enxerga (antes qualquer galpão aprovado via qualquer arquivo)
update storage.buckets set allowed_mime_types = array['image/*','video/*'] where id = 'midias';

create or replace function public.pasta_do_meu_galpao(caminho text) returns boolean
language sql stable security definer set search_path = public, storage as $$
  select exists (select 1 from public.perfis p
    where p.id::text = (storage.foldername(caminho))[1] and p.galpao_id = public.meu_galpao())
$$;
revoke all on function public.pasta_do_meu_galpao(text) from public, anon;
grant execute on function public.pasta_do_meu_galpao(text) to authenticated;

drop policy if exists "ver midias" on storage.objects;
create policy "ver midias" on storage.objects for select to authenticated
  using (bucket_id = 'midias' and public.is_aprovado() and public.pasta_do_meu_galpao(name));
drop policy if exists "enviar midias" on storage.objects;
create policy "enviar midias" on storage.objects for insert to authenticated
  with check (bucket_id = 'midias' and public.is_aprovado() and (storage.foldername(name))[1] = auth.uid()::text);
drop policy if exists "apagar midias" on storage.objects;
create policy "apagar midias" on storage.objects for delete to authenticated
  using (bucket_id = 'midias' and ((storage.foldername(name))[1] = auth.uid()::text
         or (public.tem_perm('apagar') and public.pasta_do_meu_galpao(name))));

-- 5. Funções internas não precisam ficar abertas para visitantes sem login
revoke execute on function public.is_aprovado() from public, anon;
revoke execute on function public.is_admin() from public, anon;
revoke execute on function public.tem_perm(text) from public, anon;
revoke execute on function public.tem_cargo(text[]) from public, anon;
grant execute on function public.is_aprovado(), public.is_admin(), public.tem_perm(text), public.tem_cargo(text[]) to authenticated;
-- ===== fim da segurança =====

-- ===== Segurança (v104) — auditoria de 07/11/2026 =====
-- Pode rodar mais de uma vez. Não apaga nenhum registro de avaria, garantia ou usuário.

-- 6. Login por usuário sem expor e-mails + bloqueio no servidor após 5 senhas erradas (15 min)
create table if not exists public.login_falhas (
  usuario text not null,
  quando timestamptz not null default now()
);
create index if not exists login_falhas_idx on public.login_falhas (usuario, quando desc);
alter table public.login_falhas enable row level security;
revoke all on public.login_falhas from public, anon, authenticated;

create or replace function public.login_email(u text, s text) returns jsonb
language plpgsql volatile security definer set search_path = public, extensions, auth as $$
declare
  k text := lower(regexp_replace(trim(coalesce(u, '')), '\s+', ' ', 'g'));
  e text; h text; n int;
begin
  if k = '' or s is null or length(s) > 200 or length(k) > 120 then
    return jsonb_build_object('ok', false);
  end if;
  delete from public.login_falhas where quando < now() - interval '1 day';
  select count(*) into n from public.login_falhas where usuario = k and quando > now() - interval '15 minutes';
  if n >= 5 then return jsonb_build_object('ok', false, 'bloqueado', true); end if;

  select au.email, au.encrypted_password into e, h
  from public.perfis p join auth.users au on au.id = p.id
  where lower(p.usuario) = k limit 1;

  if h is not null and h like '$2%' and h = extensions.crypt(s, h) then
    delete from public.login_falhas where usuario = k;
    return jsonb_build_object('ok', true, 'email', e);
  end if;
  if h is null then perform extensions.crypt(s, extensions.gen_salt('bf', 10)); end if;  -- mesmo tempo de resposta
  insert into public.login_falhas (usuario) values (k);
  return jsonb_build_object('ok', false, 'bloqueado', n + 1 >= 5);
end
$$;
revoke all on function public.login_email(text, text) from public;
grant execute on function public.login_email(text, text) to anon, authenticated;

-- 7. Cadastro: verifica se o usuário já existe sem devolver e-mail
create or replace function public.usuario_existe(u text) returns boolean
language sql stable security definer set search_path = public as $$
  select exists (select 1 from public.perfis
    where lower(usuario) = lower(regexp_replace(trim(coalesce(u, '')), '\s+', ' ', 'g')))
$$;
revoke all on function public.usuario_existe(text) from public;
grant execute on function public.usuario_existe(text) to anon, authenticated;

-- 8. A função antiga que entregava o e-mail de qualquer usuário deixa de ficar aberta
revoke execute on function public.email_do_usuario(text) from public, anon, authenticated;

-- 9. Endereço interno dos cadastros: pode usar uma caixa dedicada em vez da caixa do administrador
--    Para usar: insert into public.ajustes (chave, valor) values ('email_base_cadastro', 'cadastros.inventra@gmail.com')
--               on conflict (chave) do update set valor = excluded.valor;
create table if not exists public.ajustes (chave text primary key, valor text);
alter table public.ajustes enable row level security;
revoke all on public.ajustes from public, anon, authenticated;

create or replace function public.email_novo_usuario(u text) returns text
language sql stable security definer set search_path = public, auth as $$
  select split_part(x.e, '@', 1) || '+u' || substr(md5(lower(regexp_replace(trim(u), '\s+', ' ', 'g'))), 1, 10) || '@' || split_part(x.e, '@', 2)
  from (
    select coalesce(
      (select nullif(valor, '') from public.ajustes where chave = 'email_base_cadastro'),
      (select coalesce(nullif(p.email, ''), au.email) from public.perfis p join auth.users au on au.id = p.id
        where p.admin order by p.criado limit 1)) as e
  ) x
  where x.e like '%@%'
    and length(trim(coalesce(u, ''))) between 3 and 60
    and not exists (select 1 from public.perfis where lower(usuario) = lower(regexp_replace(trim(u), '\s+', ' ', 'g')))
$$;
revoke all on function public.email_novo_usuario(text) from public;
grant execute on function public.email_novo_usuario(text) to anon, authenticated;

-- 10. Galpão novo: o primeiro cadastro só vira Chefe se for a pessoa indicada ao criar o galpão
--     Para criar um galpão:  insert into public.galpoes (nome, chefe_usuario) values ('Nome do galpão', 'Usuário do chefe');
alter table public.galpoes add column if not exists chefe_usuario text;

create or replace function public.novo_usuario() returns trigger
language plpgsql security definer set search_path = public as $$
declare g uuid; primeiro boolean; usu text; chefe text;
begin
  begin g := nullif(new.raw_user_meta_data->>'galpao', '')::uuid; exception when others then g := null; end;
  if g is null or not exists (select 1 from public.galpoes where id = g) then
    select id into g from public.galpoes order by criado limit 1;
  end if;
  usu := nullif(regexp_replace(trim(coalesce(new.raw_user_meta_data->>'usuario', '')), '\s+', ' ', 'g'), '');
  select chefe_usuario into chefe from public.galpoes where id = g;
  primeiro := not exists (select 1 from public.perfis where galpao_id = g)
    and (not exists (select 1 from public.perfis)                         -- primeiro uso do sistema
         or (chefe is not null and lower(chefe) = lower(coalesce(usu, '')))); -- chefe indicado
  insert into public.perfis (id, nome, email, cargo, aprovado, admin, usuario, galpao_id)
  values (new.id, coalesce(new.raw_user_meta_data->>'nome', ''), new.email,
          case when primeiro then 'Chefe de Inventário' else 'Auxiliar de Inventário' end, primeiro, primeiro, usu, g);
  return new;
end
$$;

-- 11. Cargos valem para todos os galpões: só o administrador do galpão principal altera
drop policy if exists "admin edita cargos" on public.cargos;
create policy "admin edita cargos" on public.cargos for all to authenticated
  using (public.is_admin() and public.meu_galpao() = (select id from public.galpoes order by criado limit 1))
  with check (public.is_admin() and public.meu_galpao() = (select id from public.galpoes order by criado limit 1));
-- ===== fim da segurança v104 =====

-- ===== Segurança (v105) — auditoria avançada de 09/11/2026 =====
-- Pode rodar mais de uma vez. Não apaga avarias, garantias nem usuários.
-- Requer o bloco v104 (login_email, login_falhas). Este arquivo já o inclui antes desta parte.

-- 12. Registro de auditoria (somente leitura para administradores; ninguém grava direto)
create table if not exists public.auditoria (
  id bigserial primary key,
  quando timestamptz not null default now(),
  galpao_id uuid,
  ator uuid,
  ator_nome text,
  evento text not null,
  alvo text,
  detalhe jsonb not null default '{}'::jsonb,
  ip text
);
create index if not exists auditoria_galpao_quando on public.auditoria (galpao_id, quando desc);
create index if not exists auditoria_evento on public.auditoria (evento, quando desc);
alter table public.auditoria enable row level security;
revoke all on public.auditoria from public, anon, authenticated;
grant select on public.auditoria to authenticated;
revoke all on sequence public.auditoria_id_seq from public, anon, authenticated;
drop policy if exists "admin le auditoria" on public.auditoria;
create policy "admin le auditoria" on public.auditoria for select to authenticated
  using (public.is_admin() and (galpao_id = public.meu_galpao()
         or (galpao_id is null and public.meu_galpao() = (select id from public.galpoes order by criado limit 1))));

create or replace function public.ip_requisicao() returns text
language plpgsql stable set search_path = public as $$
declare h json; v text;
begin
  begin h := nullif(current_setting('request.headers', true), '')::json; exception when others then return null; end;
  if h is null then return null; end if;
  v := coalesce(h->>'x-forwarded-for', h->>'x-real-ip', h->>'cf-connecting-ip');
  return nullif(trim(split_part(coalesce(v, ''), ',', 1)), '');
end $$;

create or replace function public.audita(ev text, alvo text, det jsonb, gal uuid default null) returns void
language plpgsql security definer set search_path = public as $$
declare u uuid := auth.uid(); nm text; g uuid := gal;
begin
  if u is not null then
    select nome, coalesce(g, galpao_id) into nm, g from public.perfis where id = u;
  end if;
  insert into public.auditoria (galpao_id, ator, ator_nome, evento, alvo, detalhe, ip)
  values (g, u, nm, left(ev, 60), left(alvo, 200), coalesce(det, '{}'::jsonb), public.ip_requisicao());
  if random() < 0.01 then delete from public.auditoria where quando < now() - interval '365 days'; end if;
end $$;
revoke all on function public.audita(text, text, jsonb, uuid) from public, anon, authenticated;
revoke all on function public.ip_requisicao() from public, anon, authenticated;

-- 13. Gatilhos de auditoria: permissões, cargos, exclusões e entregas
create or replace function public.audita_perfis() returns trigger
language plpgsql security definer set search_path = public as $$
begin
  if tg_op = 'INSERT' then
    perform public.audita('cadastro', coalesce(new.usuario, new.nome), jsonb_build_object('cargo', new.cargo, 'admin', new.admin, 'aprovado', new.aprovado), new.galpao_id);
  elsif tg_op = 'UPDATE' then
    if new.cargo is distinct from old.cargo or new.admin is distinct from old.admin
       or new.aprovado is distinct from old.aprovado or new.recusado is distinct from old.recusado then
      perform public.audita('permissao_alterada', coalesce(new.usuario, new.nome), jsonb_build_object(
        'de', jsonb_build_object('cargo', old.cargo, 'admin', old.admin, 'aprovado', old.aprovado, 'recusado', old.recusado),
        'para', jsonb_build_object('cargo', new.cargo, 'admin', new.admin, 'aprovado', new.aprovado, 'recusado', new.recusado)), new.galpao_id);
    end if;
  elsif tg_op = 'DELETE' then
    perform public.audita('usuario_excluido', coalesce(old.usuario, old.nome), jsonb_build_object('cargo', old.cargo, 'admin', old.admin), old.galpao_id);
    return old;
  end if;
  return new;
end $$;
drop trigger if exists audita_perfis on public.perfis;
create trigger audita_perfis after insert or update or delete on public.perfis for each row execute function public.audita_perfis();

create or replace function public.audita_cargos() returns trigger
language plpgsql security definer set search_path = public as $$
begin
  perform public.audita('cargo_alterado', coalesce(new.nome, old.nome), jsonb_build_object('operacao', lower(tg_op),
    'de', case when tg_op <> 'INSERT' then to_jsonb(old.perms) end, 'para', case when tg_op <> 'DELETE' then to_jsonb(new.perms) end), null);
  return coalesce(new, old);
end $$;
drop trigger if exists audita_cargos on public.cargos;
create trigger audita_cargos after insert or update or delete on public.cargos for each row execute function public.audita_cargos();

create or replace function public.audita_ocorrencias() returns trigger
language plpgsql security definer set search_path = public as $$
declare chaves text[];
begin
  if tg_op = 'DELETE' then
    perform public.audita('registro_excluido', coalesce(old.dados->>'ref', old.sku, old.id::text),
      jsonb_build_object('id', old.id, 'categoria', old.categoria, 'sku', old.sku, 'produto', left(old.produto, 120), 'autor', old.autor), old.galpao_id);
    return old;
  end if;
  if (old.dados->'ctrl'->>'ent') is distinct from (new.dados->'ctrl'->>'ent') then
    perform public.audita('entrega_alterada', coalesce(new.dados->>'ref', new.sku),
      jsonb_build_object('id', new.id, 'de', old.dados->'ctrl'->>'ent', 'para', new.dados->'ctrl'->>'ent'), new.galpao_id);
  end if;
  select array_agg(k order by k) into chaves from (
    select key as k from jsonb_each(to_jsonb(new)) n
    where key not in ('dados') and n.value is distinct from (to_jsonb(old)->key)
    union
    select 'dados.' || key from jsonb_each(coalesce(new.dados, '{}'::jsonb)) n
    where key <> 'ctrl' and n.value is distinct from (coalesce(old.dados, '{}'::jsonb)->key)
    union
    select 'dados.' || key from jsonb_each(coalesce(old.dados, '{}'::jsonb)) o
    where key <> 'ctrl' and not (coalesce(new.dados, '{}'::jsonb) ? key)
  ) x;
  if chaves is not null then
    perform public.audita('registro_alterado', coalesce(new.dados->>'ref', new.sku, new.id::text),
      jsonb_build_object('id', new.id, 'categoria', new.categoria, 'campos', to_jsonb(chaves)), new.galpao_id);
  end if;
  return new;
end $$;
drop trigger if exists audita_ocorrencias on public.ocorrencias;
create trigger audita_ocorrencias after update or delete on public.ocorrencias for each row execute function public.audita_ocorrencias();

-- 14. Validação no servidor (vale mesmo se alguém chamar a API direto)
create or replace function public.valida_ocorrencia() returns trigger
language plpgsql set search_path = public as $$
declare ent text := new.dados->'ctrl'->>'ent';
begin
  if (tg_op = 'INSERT' or new.tipo is distinct from old.tipo) and position('<' in coalesce(new.tipo, '')) > 0 then
    raise exception 'Tipo inválido' using errcode = '22023';
  end if;
  if (tg_op = 'INSERT' or new.categoria is distinct from old.categoria) and position('<' in coalesce(new.categoria, '')) > 0 then
    raise exception 'Categoria inválida' using errcode = '22023';
  end if;
  if (tg_op = 'INSERT' or new.sku is distinct from old.sku) and length(coalesce(new.sku, '')) > 120 then
    raise exception 'Código muito longo' using errcode = '22023';
  end if;
  if (tg_op = 'INSERT' or new.produto is distinct from old.produto) and length(coalesce(new.produto, '')) > 1000 then
    raise exception 'Descrição muito longa' using errcode = '22023';
  end if;
  if ent is not null and ent <> '' and (tg_op = 'INSERT' or ent is distinct from (old.dados->'ctrl'->>'ent')) then
    if ent !~ '^\d{4}-\d{2}-\d{2}$' then raise exception 'Data de entrega inválida' using errcode = '22023'; end if;
    perform ent::date;
    if ent::date > current_date + 1 then raise exception 'Data de entrega no futuro' using errcode = '22023'; end if;
  end if;
  if pg_column_size(new.dados) > 1048576
     and (tg_op = 'INSERT' or pg_column_size(new.dados) > pg_column_size(old.dados)) then
    raise exception 'Registro grande demais' using errcode = '22023';
  end if;
  return new;
end $$;
drop trigger if exists valida_ocorrencia on public.ocorrencias;
create trigger valida_ocorrencia before insert or update on public.ocorrencias for each row execute function public.valida_ocorrencia();

create or replace function public.valida_nome_cadastro() returns trigger
language plpgsql set search_path = public as $$
begin
  if (tg_op = 'INSERT' or new.nome is distinct from old.nome)
     and (position('<' in coalesce(new.nome, '')) > 0 or length(coalesce(new.nome, '')) > 80) then
    raise exception 'Nome inválido' using errcode = '22023';
  end if;
  return new;
end $$;
drop trigger if exists valida_nome_cargo on public.cargos;
create trigger valida_nome_cargo before insert or update on public.cargos for each row execute function public.valida_nome_cadastro();
drop trigger if exists valida_nome_setor on public.setores;
create trigger valida_nome_setor before insert or update on public.setores for each row execute function public.valida_nome_cadastro();

-- 15. Login: bloqueio progressivo por usuário + bloqueio por IP (ataque a vários usuários) + registro
alter table public.login_falhas add column if not exists ip text;
create index if not exists login_falhas_ip_idx on public.login_falhas (ip, quando desc);

create or replace function public.login_email(u text, s text) returns jsonb
language plpgsql volatile security definer set search_path = public, extensions, auth as $$
declare
  k text := lower(regexp_replace(trim(coalesce(u, '')), '\s+', ' ', 'g'));
  e text; h text; g uuid; f15 int; f24 int; fip int; ult timestamptz; espera interval; vip text := public.ip_requisicao();
begin
  if k = '' or s is null or length(s) > 200 or length(k) > 120 then
    return jsonb_build_object('ok', false);
  end if;
  delete from public.login_falhas where quando < now() - interval '2 days';
  -- muitas falhas vindas do mesmo IP (tentando vários usuários): bloqueia o IP por 15 minutos
  if vip is not null then
    select count(*) into fip from public.login_falhas where login_falhas.ip = vip and quando > now() - interval '15 minutes';
    if fip >= 30 then
      return jsonb_build_object('ok', false, 'bloqueado', true, 'minutos', 15);
    end if;
  end if;
  select count(*) filter (where quando > now() - interval '15 minutes'), count(*), max(quando)
    into f15, f24, ult from public.login_falhas where usuario = k and quando > now() - interval '24 hours';
  espera := case when f24 >= 20 then interval '24 hours' when f24 >= 10 then interval '1 hour' when f15 >= 5 then interval '15 minutes' else null end;
  if espera is not null and ult > now() - espera then
    return jsonb_build_object('ok', false, 'bloqueado', true, 'minutos', ceil(extract(epoch from (ult + espera - now())) / 60));
  end if;

  select au.email, au.encrypted_password, p.galpao_id into e, h, g
  from public.perfis p join auth.users au on au.id = p.id
  where lower(p.usuario) = k limit 1;

  if h is not null and h like '$2%' and h = extensions.crypt(s, h) then
    delete from public.login_falhas where usuario = k;
    perform public.audita('login_ok', k, '{}'::jsonb, g);
    return jsonb_build_object('ok', true, 'email', e);
  end if;
  if h is null then perform extensions.crypt(s, extensions.gen_salt('bf', 10)); end if;
  insert into public.login_falhas (usuario, ip) values (k, vip);
  f15 := f15 + 1; f24 := f24 + 1;
  perform public.audita(case when f15 = 5 or f24 in (10, 20) then 'login_bloqueado' else 'login_falha' end, k,
    jsonb_build_object('usuario_existe', h is not null, 'falhas_15min', f15, 'falhas_24h', f24), g);
  return jsonb_build_object('ok', false, 'bloqueado', f15 >= 5 or f24 >= 10,
    'minutos', case when f24 >= 20 then 1440 when f24 >= 10 then 60 when f15 >= 5 then 15 else 0 end);
end
$$;
revoke all on function public.login_email(text, text) from public;
grant execute on function public.login_email(text, text) to anon, authenticated;

-- 16. Eventos enviados pelas telas (saída do sistema), com limite de uso
create or replace function public.registrar_evento(tipo text) returns void
language plpgsql security definer set search_path = public as $$
declare n int;
begin
  if auth.uid() is null or tipo not in ('logout', 'logout_todos') then return; end if;
  select count(*) into n from public.auditoria where ator = auth.uid() and quando > now() - interval '1 hour';
  if n >= 120 then return; end if;
  perform public.audita(tipo, null, '{}'::jsonb, null);
end $$;
revoke all on function public.registrar_evento(text) from public, anon;
grant execute on function public.registrar_evento(text) to authenticated;

-- 17. Ações administrativas também ficam registradas
do $$ begin
  if to_regprocedure('public.admin_redefinir_senha(uuid,text)') is not null then
    execute 'create or replace function public.audita_senha_admin() returns trigger language plpgsql security definer set search_path = public, auth as $f$
      begin
        if new.encrypted_password is distinct from old.encrypted_password and auth.uid() is not null and auth.uid() <> new.id then
          perform public.audita(''senha_redefinida_admin'', (select coalesce(usuario, nome) from public.perfis where id = new.id), ''{}''::jsonb, null);
        end if;
        return new;
      end $f$';
    execute 'drop trigger if exists audita_senha_admin on auth.users';
    execute 'create trigger audita_senha_admin after update on auth.users for each row execute function public.audita_senha_admin()';
  end if;
exception when insufficient_privilege then
  raise notice 'Sem permissão para criar gatilho em auth.users; redefinições de senha não serão auditadas.';
end $$;
-- ===== fim da segurança v105 =====

notify pgrst, 'reload schema';

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
