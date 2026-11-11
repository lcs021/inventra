#!/usr/bin/env bash
# Sobe um Postgres 16 descartável, aplica shim + migrações e roda diagnóstico e testes.
# Uso: bash supabase/run_local.sh        (precisa de postgresql-16 instalado)
set -euo pipefail
cd "$(dirname "$0")"
BIN=/usr/lib/postgresql/16/bin; D=${PGDATA_TMP:-/tmp/pgdata}; PORT=${PGPORT_TMP:-5544}
RUN=""; [ "$(id -u)" = 0 ] && RUN="su postgres -c"
run() { if [ -n "$RUN" ]; then $RUN "$*"; else sh -c "$*"; fi; }
[ -d "$D" ] || { mkdir -p "$D"; [ -n "$RUN" ] && chown postgres "$D"; run "$BIN/initdb -D $D -U postgres" >/dev/null; }
run "$BIN/pg_ctl -D $D -o '-p $PORT -k /tmp' -l /tmp/inventra_pg.log -w start" >/dev/null 2>&1 || true
P="psql -h /tmp -p $PORT -U postgres -v ON_ERROR_STOP=1 -q"
$P -d postgres -c "drop database if exists inv_test" -c "create database inv_test"
for f in tests/shim_supabase.sql migrations/*.sql; do echo ">> $f"; $P -d inv_test -f "$f" >/dev/null; done
echo ">> diagnóstico"; $P -d inv_test -f diagnostico/diagnostico_somente_leitura.sql | tail -3
echo ">> testes de isolamento"; psql -h /tmp -p $PORT -U postgres -d inv_test -q -f tests/isolamento_baseline.sql 2>&1 | grep -E "^(PASS|FAIL|RESUMO)"
