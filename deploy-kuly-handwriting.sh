#!/usr/bin/env bash
set -euo pipefail

ROOT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
REMOTE="${KULI_REMOTE:-kuly}"
REMOTE_APP_DIR="${HANDWRITING_REMOTE_APP_DIR:-/opt/kuly/handwriting-web}"
SITE_URL="${HANDWRITING_SITE_URL:-https://kuly.com.cn}"

echo "==> Deploying handwriting workbench to ${REMOTE}:${REMOTE_APP_DIR}"

rsync -az --delete \
  --exclude='.git/' \
  --exclude='.DS_Store' \
  --exclude='.env' \
  --exclude='backend/.venv/' \
  --exclude='frontend/node_modules/' \
  --exclude='frontend/dist/' \
  --exclude='logs/' \
  --exclude='output/' \
  --exclude='backend/tasks.db' \
  "${ROOT_DIR}/" "${REMOTE}:${REMOTE_APP_DIR}/"

ssh "${REMOTE}" \
  "HANDWRITING_APP_DIR='${REMOTE_APP_DIR}' HANDWRITING_SITE_URL='${SITE_URL}' bash -s" <<'REMOTE_SCRIPT'
set -euo pipefail

cd "${HANDWRITING_APP_DIR}"
export PATH="${HOME}/.local/bin:${PATH}"

if [ ! -f .env ]; then
  cat > .env <<ENV
# Handwriting workbench production placeholder.
# MinerU cloud API is configured here for the long-lived production deployment.
MINERU_BASE_URL=https://mineru.net/api/v4
MINERU_API_TOKEN=replace-with-cloud-mineru-token
MINERU_PUBLIC_BASE_URL=${HANDWRITING_SITE_URL}/handwriting-api
MINERU_MODEL_VERSION=vlm
MINERU_TRUST_ENV=1
NO_PROXY=
FONT_ASSETS_DIR=${HANDWRITING_APP_DIR}/ttf_files
FONT_ASSETS_BUNDLED_DIR=${HANDWRITING_APP_DIR}/ttf_files
ENV
  chmod 600 .env
  echo "==> Created placeholder .env. Update MinerU values before expecting PDF extraction to work."
fi

echo "==> Build handwriting frontend"
cd frontend
npm ci --legacy-peer-deps
VUE_APP_PUBLIC_PATH=/handwriting/ VUE_APP_API_BASE_URL=/handwriting-api npm run build

echo "==> Prepare backend environment"
cd "${HANDWRITING_APP_DIR}/backend"
python3 -m venv .venv
.venv/bin/python -m pip install --upgrade pip wheel
.venv/bin/pip install -r requirements.txt

echo "==> Install systemd unit"
sudo tee /etc/systemd/system/handwriting-backend.service >/dev/null <<UNIT
[Unit]
Description=Kuli handwriting backend
After=network.target

[Service]
Type=simple
User=ubuntu
Group=ubuntu
WorkingDirectory=${HANDWRITING_APP_DIR}/backend
EnvironmentFile=${HANDWRITING_APP_DIR}/.env
Environment=PYTHONUNBUFFERED=1
ExecStart=${HANDWRITING_APP_DIR}/backend/.venv/bin/python app.py
Restart=always
RestartSec=5

[Install]
WantedBy=multi-user.target
UNIT

echo "==> Install nginx handwriting routes"
sudo rm -f /etc/nginx/conf.d/kuly-handwriting.conf
sudo tee /etc/nginx/snippets/kuly-handwriting.conf >/dev/null <<NGINX
location = /handwriting {
    return 301 /handwriting/;
}

location = /handwriting/index.html {
    alias ${HANDWRITING_APP_DIR}/frontend/dist/index.html;
}

location ^~ /handwriting/ {
    alias ${HANDWRITING_APP_DIR}/frontend/dist/;
    try_files \$uri \$uri/ /handwriting/index.html;
}

location ^~ /handwriting-api/ {
    rewrite ^/handwriting-api/(.*)\$ /\$1 break;
    proxy_pass http://127.0.0.1:5005;
    proxy_http_version 1.1;
    proxy_set_header Host \$host;
    proxy_set_header X-Real-IP \$remote_addr;
    proxy_set_header X-Forwarded-For \$proxy_add_x_forwarded_for;
    proxy_set_header X-Forwarded-Proto \$scheme;
    proxy_set_header Upgrade \$http_upgrade;
    proxy_set_header Connection "upgrade";
    proxy_read_timeout 300s;
    proxy_send_timeout 300s;
}
NGINX

sudo chmod o+x /opt/kuly

sudo python3 - <<'PY'
from pathlib import Path

path = Path("/etc/nginx/sites-enabled/kuly.conf")
text = path.read_text()
text = text.replace("    include /etc/nginx/conf.d/kuly-handwriting.conf;\n\n", "")
needle = "    location /api/ {"
include_line = "    include /etc/nginx/snippets/kuly-handwriting.conf;\n\n"
if include_line not in text:
    text = text.replace(needle, include_line + needle, 1)
path.write_text(text)
PY

if ! sudo grep -q 'include /etc/nginx/snippets/kuly-handwriting.conf;' /etc/nginx/sites-enabled/kuly.conf; then
  sudo python3 - <<'PY'
from pathlib import Path

path = Path("/etc/nginx/sites-enabled/kuly.conf")
text = path.read_text()
needle = "    location /api/ {"
include_line = "    include /etc/nginx/snippets/kuly-handwriting.conf;\n\n"
if include_line not in text:
    text = text.replace(needle, include_line + needle, 1)
    path.write_text(text)
PY
fi

echo "==> Reload nginx"
sudo nginx -t
sudo systemctl reload nginx

echo "==> Start handwriting backend if MinerU config is usable"
sudo systemctl daemon-reload
sudo systemctl enable handwriting-backend >/dev/null
if sudo systemctl restart handwriting-backend; then
  sleep 3
  if systemctl is-active --quiet handwriting-backend; then
    echo "==> handwriting-backend active"
  else
    echo "WARN: handwriting-backend did not stay active. MinerU placeholder/private endpoint may be unavailable."
    sudo journalctl -u handwriting-backend --no-pager -n 40 || true
  fi
else
  echo "WARN: handwriting-backend restart failed. MinerU placeholder/private endpoint may be unavailable."
  sudo journalctl -u handwriting-backend --no-pager -n 40 || true
fi

echo "==> Verify handwriting frontend path"
curl -fsSI http://127.0.0.1/handwriting/ >/dev/null
echo "==> Handwriting deploy finished"
REMOTE_SCRIPT

echo "==> Done: ${SITE_URL}/handwriting/"
