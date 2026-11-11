# supabase/ — banco versionado e testes

Nada aqui é carregado pelo site. São scripts para diagnosticar, versionar e testar o banco.

## Estrutura
- `migrations/0001_base_instalacao_inventra.sql` — instalação base (extraída do SQL que estava embutido no módulo).
- `migrations/0002_seguranca_v106.sql` — endurecimento de segurança (v106).
- `diagnostico/diagnostico_somente_leitura.sql` — **um único SELECT**, não altera nada. Mostra tabelas, RLS, policies, funções expostas ao `anon`, índices, triggers e contagens (sem dados pessoais).
- `tests/shim_supabase.sql` — simula `auth`, `storage` e papéis do Supabase num Postgres comum.
- `tests/isolamento_baseline.sql` — 12 testes de isolamento entre duas empresas. Hoje 5 falham de propósito; devem passar após a Fase 3.
- `run_local.sh` — recria o banco local de teste e roda tudo.

## Rodar localmente
`bash supabase/run_local.sh` (Postgres 16).

## Diagnosticar o banco real (somente leitura)
1. Supabase → SQL Editor → cole `diagnostico_somente_leitura.sql` → Run.
2. Envie o resultado de volta (nenhum e-mail, senha ou ocorrência é exibido).

## Criar o ambiente de teste (staging)
1. supabase.com → New project (plano gratuito), nome `inventra-teste`.
2. SQL Editor: rode `0001`, depois `0002`.
3. **Nunca** coloque a chave `service_role` no site; use apenas a `anon` em `config.js`.
4. Só depois de passar no teste, a mesma migração vai ao banco real, com backup antes.

## Achados do baseline (a corrigir na Fase 3)
- **Grave:** qualquer admin de qualquer empresa edita `cargos` (a policy consulta `galpoes` sob RLS e sempre vê só o próprio galpão).
- `cargos` e `ajustes` são globais, sem empresa.
- `lista_galpoes()` e `meu_galpao()` executáveis por visitante anônimo.
- Nome de usuário é único no sistema todo, não por empresa.
