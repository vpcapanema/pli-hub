#!/usr/bin/env bash
# Deploy do PLI Hub no servidor de producao (ubuntu@56.125.163.194).
# Uso: ssh no servidor, depois:  cd /home/ubuntu/local-github-vm && ./deploy/deploy.sh
set -euo pipefail

cd "$(dirname "$0")/.."

echo "==> Atualizando o codigo a partir do GitHub"
git pull --ff-only origin master

echo "==> Reconstruindo e subindo o container"
docker compose up -d --build

echo "==> Aguardando o container ficar saudavel"
for _ in $(seq 1 30); do
  status=$(docker inspect pli-hub --format '{{.State.Health.Status}}' 2>/dev/null || echo unknown)
  [ "$status" = "healthy" ] && break
  sleep 2
done
echo "    status: ${status:-unknown}"

echo "==> Conferindo a resposta local do container"
curl -fsS http://127.0.0.1:8080/health && echo
curl -fsS -o /dev/null -w "    /pli-hub/ -> HTTP %{http_code}\n" http://127.0.0.1:8080/pli-hub/

echo "==> Deploy concluido: https://56.125.163.194/pli-hub/"
