#!/bin/bash
# DO Bot — Production server setup (ngrok static tunnel, no domain needed)
# Run as root on Ubuntu 20.04+ / Debian 11+
# Usage: bash setup_server.sh

set -e
APP_DIR=/opt/do_bot
LOG_DIR=/var/log/do_bot

echo "=== [1/6] Installing system packages ==="
apt-get update -q
apt-get install -y python3 python3-venv python3-pip curl unzip

echo "=== [2/6] Installing ngrok ==="
if ! command -v ngrok &>/dev/null; then
    curl -sSL https://ngrok-agent.s3.amazonaws.com/ngrok.asc \
        | tee /etc/apt/trusted.gpg.d/ngrok.asc >/dev/null
    echo "deb https://ngrok-agent.s3.amazonaws.com buster main" \
        > /etc/apt/sources.list.d/ngrok.list
    apt-get update -q && apt-get install -y ngrok
fi

echo "=== [3/6] Creating directories ==="
mkdir -p "$APP_DIR" "$LOG_DIR"
chown www-data:www-data "$LOG_DIR"

echo "=== [4/6] Copying code ==="
rsync -av --exclude='__pycache__' --exclude='*.pyc' --exclude='.git' \
    --exclude='config.txt' --exclude='data/*.xlsx' \
    ./ "$APP_DIR/"
chown -R www-data:www-data "$APP_DIR"

echo "=== [5/6] Creating Python virtual environment ==="
python3 -m venv "$APP_DIR/venv"
"$APP_DIR/venv/bin/pip" install --upgrade pip -q
"$APP_DIR/venv/bin/pip" install -r "$APP_DIR/requirements.txt" -q

echo "=== [6/6] Creating data directory and config files ==="
mkdir -p "$APP_DIR/data"
chown -R www-data:www-data "$APP_DIR/data"

# Create .env if not present
ENV_FILE="$APP_DIR/.env"
if [ ! -f "$ENV_FILE" ]; then
    cp "$APP_DIR/.env.example" "$ENV_FILE"
    chmod 600 "$ENV_FILE"
fi

# Install systemd services
cp "$APP_DIR/do_bot.service"       /etc/systemd/system/do_bot.service
cp "$APP_DIR/ngrok_do_bot.service" /etc/systemd/system/ngrok_do_bot.service
systemctl daemon-reload
systemctl enable do_bot ngrok_do_bot

echo ""
echo "=== Setup complete! ==="
echo ""
echo "Before starting, complete these steps:"
echo ""
echo "  1. Copy your data files:"
echo "       cp ZSDOROUTEWRH.xlsx       $APP_DIR/data/"
echo "       cp 'master lorry.xlsx'     $APP_DIR/"
echo ""
echo "  2. Fill in credentials:"
echo "       nano $APP_DIR/.env"
echo "       (set META_ACCESS_TOKEN, META_PHONE_NUMBER_ID)"
echo ""
echo "  3. Fill in your ngrok authtoken:"
echo "       nano $APP_DIR/ngrok_do_bot.yml"
echo "       (replace YOUR_NGROK_AUTHTOKEN_HERE)"
echo ""
echo "  4. Start both services:"
echo "       systemctl start do_bot"
echo "       systemctl start ngrok_do_bot"
echo ""
echo "  5. Check they are running:"
echo "       journalctl -u do_bot        -f"
echo "       journalctl -u ngrok_do_bot  -f"
echo ""
echo "  6. Meta webhook URL (already set, no change needed):"
echo "       https://degraded-sincerity-glue.ngrok-free.dev/webhook"
