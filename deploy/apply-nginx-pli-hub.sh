#!/usr/bin/env bash
set -euo pipefail

TARGET_SNIPPET=/etc/nginx/snippets/pli-hub.conf
SITE_CONF=/etc/nginx/sites-available/default

sudo cp deploy/nginx-pli-hub.conf "$TARGET_SNIPPET"

if ! sudo grep -q "include /etc/nginx/snippets/pli-hub.conf;" "$SITE_CONF"; then
  sudo sed -i '/server_name _;/a\
\n    include /etc/nginx/snippets/pli-hub.conf;' "$SITE_CONF"
fi

sudo nginx -t
sudo systemctl reload nginx

echo "OK: rota /pli-hub ativa no Nginx"
