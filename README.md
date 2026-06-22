# Blockchain.com DeFi Wallet — Hostinger deployment

## Local run

```bash
cd C:\Users\VW\Desktop\blockwall
python -m venv venv
venv\Scripts\activate          # Windows
# source venv/bin/activate    # Linux/Mac
pip install -r requirements.txt
python app.py
# → http://localhost:5001
```

Demo login: `demo@blockchain-demo.com` / `demo1234`

## Admin

```bash
# → http://localhost:5001/admin/login
# Default: admin / change-me-now
# Override via env vars before launch:
set ADMIN_USER=yourname
set ADMIN_PASS=your-strong-password
set SECRET_KEY=$(python -c "import secrets;print(secrets.token_hex(32))")
```

## Upload format

Excel `.xlsx` with header row:
- `email` (required) — user email
- `password` (required) — plain text, gets hashed on import
- `name` (optional) — display name

A sample download is in the admin dashboard ("Beispiel-xlsx herunterladen").

## Deploy to Hostinger (kunden-blockchain.com)

You have two options. **Option A** keeps the existing automailer container untouched — recommended.

### Option A — Standalone on a subpath or subdomain (recommended)

Since `kunden-blockchain.com` is probably already serving your landing page, host this app on a **subdomain** (`app.kunden-blockchain.com`) or **subpath** (`kunden-blockchain.com/wallet/`). Subdomain is cleaner.

1. **In Hostinger panel** → Domains → add DNS A-record `app` → your VPS IP.
2. **SSH into VPS** and pick a project dir, e.g. `/root/blockwall/`.
3. Copy project up:
   ```bash
   scp -r C:\Users\VW\Desktop\blockwall\* root@72.60.191.112:/root/blockwall/
   ```
   (SSH key: `C:\Users\VW\.ssh\id_ed25519`, login `PENNYdoggy`.)
4. **Add nginx vhost** at `/etc/nginx/sites-available/blockwall`:
   ```nginx
   server {
       listen 443 ssl http2;
       server_name app.kunden-blockchain.com;

       ssl_certificate     /etc/letsencrypt/live/app.kunden-blockchain.com/fullchain.pem;
       ssl_certificate_key /etc/letsencrypt/live/app.kunden-blockchain.com/privkey.pem;

       client_max_body_size 16M;

       location / {
           proxy_pass http://127.0.0.1:5001;
           proxy_set_header Host              $host;
           proxy_set_header X-Real-IP         $remote_addr;
           proxy_set_header X-Forwarded-For   $proxy_add_x_forwarded_for;
           proxy_set_header X-Forwarded-Proto $scheme;
       }
   }
   server { listen 80; server_name app.kunden-blockchain.com; return 301 https://$host$request_uri; }
   ```
5. Issue cert with `certbot --nginx -d app.kunden-blockchain.com`.
6. **Run the app with gunicorn + systemd** for stability:
   `/etc/systemd/system/blockwall.service`:
   ```ini
   [Unit]
   Description=Blockwall wallet (gunicorn)
   After=network.target

   [Service]
   User=root
   WorkingDirectory=/root/blockwall
   Environment=ADMIN_USER=yourname
   Environment=ADMIN_PASS=your-strong-password
   Environment=SECRET_KEY=PUT-A-LONG-RANDOM-HEX-HERE
   ExecStart=/root/blockwall/venv/bin/gunicorn -w 2 -b 127.0.0.1:5001 app:app
   Restart=always

   [Install]
   WantedBy=multi-user.target
   ```
   Then `systemctl daemon-reload && systemctl enable --now blockwall`.
7. `nginx -t && systemctl reload nginx` → done. Site is live at https://app.kunden-blockchain.com.

### Option B — Add to existing automailer docker-compose

If you prefer to run this inside Docker alongside your existing automailer stack on the same VPS, copy `app.py`, `requirements.txt`, `app/`, `data/` into a sibling directory and add a service to your existing `docker-compose.yml`:

```yaml
  blockwall:
    build: ./blockwall
    container_name: blockwall
    restart: unless-stopped
    environment:
      - ADMIN_USER=${BLOCKWALL_ADMIN_USER}
      - ADMIN_PASS=${BLOCKWALL_ADMIN_PASS}
      - SECRET_KEY=${BLOCKWALL_SECRET_KEY}
    volumes:
      - ./blockwall/data:/app/data
    expose:
      - "5001"
```

With a `Dockerfile`:
```dockerfile
FROM python:3.11-slim
WORKDIR /app
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt gunicorn
COPY app.py ./
COPY app ./app
EXPOSE 5001
CMD ["gunicorn", "-w", "2", "-b", "0.0.0.0:5001", "app:app"]
```

Then point nginx at `blockwall:5001` and add cert. That's it.

## File layout

```
blockwall/
├── app.py                  # Flask app, all routes
├── requirements.txt
├── README.md
├── app/
│   ├── static/
│   │   ├── css/app.css     # single stylesheet, matches design
│   │   └── img/favicon.svg
│   └── templates/
│       ├── _base.html
│       ├── _layout.html    # header + sidebar shell
│       ├── home.html       # /wallet/home
│       ├── assets.html     # /wallet/assets
│       ├── currency.html   # /wallet/currency/<sym>
│       ├── activity.html   # /wallet/activity
│       ├── login.html
│       ├── admin_login.html
│       └── admin.html
└── data/                   # auto-created on first run
    ├── users.xlsx          # all users (admin-controlled)
    ├── wallets.xlsx        # token balances per user
    └── activity.xlsx       # activity feed per user
```

## Security notes

- Passwords are hashed with `werkzeug.security.generate_password_hash` (PBKDF2).
- Session cookies: HttpOnly, SameSite=Lax, signed with `SECRET_KEY` — set this to a long random hex string in production.
- Admin login lives at `/admin/login` (separate session flag `is_admin`).
- Upload cap: 16MB.
- `.env`-style secrets are read from env vars — never hardcoded.

## Quick smoke test after deploy

```bash
curl -I https://app.kunden-blockchain.com/login            # expect 200
curl -I https://app.kunden-blockchain.com/admin/login      # expect 200
curl -I https://app.kunden-blockchain.com/wallet/home      # expect 302 → /login
```