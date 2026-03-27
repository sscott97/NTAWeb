#!/bin/bash
set -e

echo ""
echo "================================================"
echo "  NTAWeb Server Setup"
echo "================================================"
echo ""

# ── 1. Swap file (R package compilation needs it on 1GB RAM) ──
echo "[1/9] Adding 2GB swap..."
if [ ! -f /swapfile ]; then
  sudo fallocate -l 2G /swapfile
  sudo chmod 600 /swapfile
  sudo mkswap /swapfile
  sudo swapon /swapfile
  echo '/swapfile none swap sw 0 0' | sudo tee -a /etc/fstab > /dev/null
  echo "     Swap added."
else
  echo "     Swap already exists, skipping."
fi

# ── 2. System update + dependencies ──
echo "[2/9] Updating system and installing dependencies..."
sudo apt-get update -y -q
sudo apt-get install -y -q \
  python3 python3-pip python3-venv \
  nginx git curl wget \
  software-properties-common dirmngr \
  libssl-dev libcurl4-openssl-dev libxml2-dev \
  libfontconfig1-dev libharfbuzz-dev libfribidi-dev \
  libfreetype6-dev libpng-dev libtiff5-dev libjpeg-dev \
  gfortran libgfortran5 \
  iptables-persistent

# ── 3. Install R 4.x ──
echo "[3/9] Installing R 4.x..."
wget -qO- https://cloud.r-project.org/bin/linux/ubuntu/marutter_pubkey.asc \
  | sudo tee /etc/apt/trusted.gpg.d/cran_ubuntu_key.asc > /dev/null
echo "deb https://cloud.r-project.org/bin/linux/ubuntu $(lsb_release -cs)-cran40/" \
  | sudo tee /etc/apt/sources.list.d/cran.list > /dev/null
sudo apt-get update -y -q
sudo apt-get install -y -q r-base r-base-dev

# ── 4. Install R packages ──
echo "[4/9] Installing R packages (this takes 5-10 min)..."
sudo Rscript -e "
  options(repos = list(CRAN = 'https://cloud.r-project.org'))
  pkgs <- c('readxl','jsonlite','ggplot2','dplyr','readr','tidyr',
            'cowplot','minpack.lm','scales','plotly','tidyverse')
  install.packages(pkgs, Ncpus = 1, quiet = TRUE)
  cat('R packages installed.\n')
"

# ── 5. Create app directory + Python venv ──
echo "[5/9] Setting up Python environment..."
sudo mkdir -p /var/www/ntaweb/app
sudo chown -R ubuntu:ubuntu /var/www/ntaweb

python3 -m venv /var/www/ntaweb/venv
/var/www/ntaweb/venv/bin/pip install --upgrade pip -q
/var/www/ntaweb/venv/bin/pip install flask openpyxl pillow gunicorn -q

# ── 6. Open OS firewall ports ──
echo "[6/9] Opening firewall ports 80 and 443..."
sudo iptables -I INPUT 6 -m state --state NEW -p tcp --dport 80 -j ACCEPT 2>/dev/null || true
sudo iptables -I INPUT 6 -m state --state NEW -p tcp --dport 443 -j ACCEPT 2>/dev/null || true
sudo netfilter-persistent save

# ── 7. Configure nginx ──
echo "[7/9] Configuring nginx..."
sudo tee /etc/nginx/sites-available/ntaweb > /dev/null <<'NGINX'
server {
    listen 80;
    server_name _;

    client_max_body_size 50M;

    location / {
        proxy_pass         http://127.0.0.1:8000;
        proxy_set_header   Host              $host;
        proxy_set_header   X-Real-IP         $remote_addr;
        proxy_set_header   X-Forwarded-For   $proxy_add_x_forwarded_for;
        proxy_read_timeout 300s;
        proxy_send_timeout 300s;
        proxy_connect_timeout 75s;
    }
}
NGINX

sudo ln -sf /etc/nginx/sites-available/ntaweb /etc/nginx/sites-enabled/ntaweb
sudo rm -f /etc/nginx/sites-enabled/default
sudo nginx -t
sudo systemctl enable nginx
sudo systemctl restart nginx

# ── 8. Create gunicorn systemd service ──
echo "[8/9] Creating gunicorn service..."
sudo tee /etc/systemd/system/ntaweb.service > /dev/null <<'SERVICE'
[Unit]
Description=NTAWeb Flask Application
After=network.target

[Service]
User=ubuntu
Group=ubuntu
WorkingDirectory=/var/www/ntaweb/app
Environment="PATH=/var/www/ntaweb/venv/bin"
ExecStart=/var/www/ntaweb/venv/bin/gunicorn \
    --workers 2 \
    --bind 127.0.0.1:8000 \
    --timeout 300 \
    --keep-alive 5 \
    --log-level info \
    --access-logfile /var/log/ntaweb-access.log \
    --error-logfile /var/log/ntaweb-error.log \
    app:app
Restart=always
RestartSec=5

[Install]
WantedBy=multi-user.target
SERVICE

sudo systemctl daemon-reload
sudo systemctl enable ntaweb

# ── 9. Done ──
echo "[9/9] Setup complete!"
echo ""
echo "================================================"
echo "  Next step: upload your app files, then run:"
echo "  sudo systemctl start ntaweb"
echo "  Your site will be at http://144.21.50.245"
echo "================================================"
echo ""
