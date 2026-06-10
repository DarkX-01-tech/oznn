# AGENTS.md

## Project overview

**oznn** is a hospital cafeteria meal-management system ("yemek listesi") written in **Classic ASP (VBScript)**. The public menu is served from `yemek_index.asp`; admin workflows live under `panel.asp`, `yemek_ekle.asp`, `yemek_liste.asp`, and related pages.

Application code lives on feature branches (for example `cursor/menu-degisiklik-istatistik-df8c`). `main` currently contains only a README stub — check out a feature branch before developing.

Production targets **Windows IIS** with an **Access** database. `database/Connection.asp` and `giris.asp` are deployment-local files and are not committed in all branches.

## Cursor Cloud specific instructions

### Runtime stack (Linux dev)

| Component | Purpose |
|-----------|---------|
| **AxonASP 2.2** (`/opt/axonasp/axonasp-http`) | Classic ASP runtime on Linux |
| **MariaDB** (`oznn_yemek` database) | Dev database substitute for Access |
| `database/Connection.asp` | MySQL connection string (dev) |
| `giris.asp` | Dev login stub (accepts any username) |

### Starting services

MariaDB must be running before the app:

```bash
sudo service mariadb start
```

Start the ASP server (use tmux for long-running processes):

```bash
SESSION_NAME="axonasp-http"
tmux -f /exec-daemon/tmux.portal.conf has-session -t "=$SESSION_NAME" 2>/dev/null \
  || tmux -f /exec-daemon/tmux.portal.conf new-session -d -s "$SESSION_NAME" -c /opt/axonasp -- "${SHELL:-zsh}" -l
tmux -f /exec-daemon/tmux.portal.conf send-keys -t "$SESSION_NAME:0.0" \
  'cd /opt/axonasp && SERVER_WEB_ROOT=/workspace ./axonasp-http' C-m
```

URLs:

- Public menu: http://localhost:8801/yemek_index.asp
- Dev login: http://localhost:8801/giris.asp
- Admin panel: http://localhost:8801/panel.asp (requires session via `giris.asp`)

### Database setup

See `scripts/setup-dev.sh` and `scripts/init-dev-db.sql`. Credentials: user `oznn`, password `oznn_dev`, database `oznn_yemek`.

### SQL compatibility caveats

The codebase uses **Microsoft Access SQL** (`#date#` literals, `SELECT TOP N`, `aktif = True`). On Linux dev with MariaDB:

- `SELECT` queries using `YEAR()` / `MONTH()` work for menu listing and statistics pages.
- `INSERT`/`UPDATE` with `#date#` or `SELECT TOP` will fail until run on Windows/Access or queries are adapted.

Admin write flows (`yemek_ekle.asp`, `yemek_kisi_giris.asp`) may not work fully on Linux dev; read-only pages (`yemek_index.asp`, `yemek_istatistik.asp`, `panel.asp`) are the reliable test targets.

### Lint / tests

No automated lint or test suite is configured. Verify changes by loading affected `.asp` pages through AxonASP and checking HTTP 200 responses.

### Branch workflow

Check out the relevant feature branch before working:

```bash
git fetch origin
git checkout cursor/menu-degisiklik-istatistik-df8c   # or another feature branch
```
