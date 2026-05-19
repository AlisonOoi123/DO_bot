#!/bin/bash
# DO Bot — One-time production server setup script
# Run as root (or with sudo) on a fresh Ubuntu/Debian server.
# Usage: bash setup_server.sh YOUR_DOMAIN

set -e
DOMAIN=${1:?'Usage: bash setup_server.sh YOUR_DOMAIN'}
APP_DIR=/opt/do_bot
LOG_DIR=/var/log/do_bot

echo "=== [1/7] Installing system packages ==="
apt-get update -q
apt-get install -y python3 python3-venv python3-pip nginx certbot python3-certbot-nginx

echo "=== [2/7] Creating app directory ==="
mkdir -p "$APP_DIR" "$LOG_DIR"
chown www-data:www-data "$LOG_DIR"

echo "=== [3/7] Cloning / copying code ==="
# If running from the repo directory, copy files over.
rsync -av --exclude='__pycache__' --exclude='*.pyc' --exclude='.git' \
    --exclude='config.txt' --exclude='data/*.xlsx' \
    ./ "$APP_DIR/"
chown -R www-data:www-data "$APP_DIR"

echo "=== [4/7] Creating Python virtual environment ==="
python3 -m venv "$APP_DIR/venv"
"$APP_DIR/venv/bin/pip" install --upgrade pip -q
"$APP_DIR/venv/bin/pip" install -r "$APP_DIR/requirements.txt" -q

echo "=== [5/7] Creating data directory and .env file ==="
mkdir -p "$APP_DIR/data"
chown www-data:www-data "$APP_DIR/data"

ENV_FILE="$APP_DIR/.env"
if [ ! -f "$ENV_FILE" ]; then
    cp "$APP_DIR/.env.example" "$ENV_FILE"
    chmod 600 "$ENV_FILE"
    echo ""
    echo ">>> IMPORTANT: Edit $ENV_FILE and fill in your real credentials:"
    echo "    META_ACCESS_TOKEN, META_PHONE_NUMBER_ID, PUBLIC_BASE_URL"
    echo ""
fi

echo "=== [6/7] Installing systemd service ==="
# Replace placeholder domain in service file (already correct — no domain there)
cp "$APP_DIR/do_bot.service" /etc/systemd/system/do_bot.service
systemctl daemon-reload
systemctl enable do_bot

echo "=== [7/7] Setting up Nginx + SSL ==="
# Replace placeholder in nginx config
sed "s/YOUR_DOMAIN_HERE/$DOMAIN/g" "$APP_DIR/nginx_do_bot.conf" \
    > "/etc/nginx/sites-available/do_bot"
ln -sf /etc/nginx/sites-available/do_bot /etc/nginx/sites-enabled/do_bot
rm -f /etc/nginx/sites-enabled/default

# Obtain SSL certificate (requires domain DNS already pointing to this server)
certbot --nginx -d "$DOMAIN" --non-interactive --agree-tos -m "admin@$DOMAIN"

nginx -t && systemctl reload nginx

echo ""
echo "=== Setup complete ==="
echo "Next steps:"
echo "  1. Copy your data files to $APP_DIR/data/:"
echo "       ZSDOROUTEWRH.xlsx"
echo "  2. Copy master lorry.xlsx to $APP_DIR/"
echo "  3. Edit $ENV_FILE with real credentials"
echo "  4. Start the bot: systemctl start do_bot"
echo "  5. Check logs:   journalctl -u do_bot -f"
echo "  6. Update Meta webhook URL to: https://$DOMAIN/webhook"
