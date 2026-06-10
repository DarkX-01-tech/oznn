#!/usr/bin/env bash
set -euo pipefail
ROOT="$(cd "$(dirname "$0")/.." && pwd)"

echo "==> Ensuring MariaDB is running"
sudo service mariadb start >/dev/null 2>&1 || true

echo "==> Initializing database schema and seed data"
sudo mysql < "$ROOT/scripts/init-dev-db.sql"

echo "==> Dev setup complete"
echo "    Start server: SERVER_WEB_ROOT=$ROOT /opt/axonasp/axonasp-http"
echo "    Public menu:  http://localhost:8801/yemek_index.asp"
echo "    Admin login:  http://localhost:8801/giris.asp"
