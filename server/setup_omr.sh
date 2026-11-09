#!/usr/bin/env bash
# OMR čitač anketa - postavljanje na Oracle poslužitelj (Ubuntu 22.04/24.04, ARM ili x86),
# uz aplikacije koje već rade na istom poslužitelju (npr. Career tests). Njih ne dira.
#
#   sudo bash setup_omr.sh                      -> https://ankete.<IP-s-crticama>.sslip.io
#   sudo bash setup_omr.sh ankete.example.hr    -> https://ankete.example.hr  (DNS A zapis -> IP poslužitelja)
#   sudo bash setup_omr.sh --nova-lozinka       -> nova lozinka za stranicu
#
# Ponovno pokretanje instalira najnoviju verziju programa s GitHuba; ključ, prijava i lozinka ostaju.
set -euo pipefail

REPO=https://github.com/maksvin-tech/chitachanketa.git
APP=/opt/omr-ankete
VENV=/opt/omr-ankete-venv
DATA=/var/lib/omr-ankete
PORT=8502
WEB_USER=ured
SITE=/etc/caddy/sites/omr-ankete.caddy
PASSFILE=/root/omr-ankete-lozinka.txt
export DEBIAN_FRONTEND=noninteractive

log()  { printf '\n\033[1;34m==> %s\033[0m\n' "$*"; }
fail() { printf '\n\033[1;31mZaustavljeno: %s\033[0m\n' "$*"; exit 1; }

[[ $EUID -eq 0 ]] || fail "pokrenite sa sudo:  sudo bash setup_omr.sh"
. /etc/os-release
[[ "${ID:-}" == "ubuntu" ]] || fail "potreban je Ubuntu, a poslužitelj ima ${PRETTY_NAME:-nepoznat sustav}."

NEW_PASSWORD=0
DOMAIN=""
for arg in "$@"; do
  case "$arg" in
    --nova-lozinka) NEW_PASSWORD=1 ;;
    *) DOMAIN="$arg" ;;
  esac
done
IP="$(curl -s -m 5 https://ifconfig.me || true)"
if [[ -z "$DOMAIN" ]]; then
  [[ "$IP" =~ ^[0-9]+\.[0-9]+\.[0-9]+\.[0-9]+$ ]] || fail "javna IP adresa nije pronađena; navedite adresu: sudo bash setup_omr.sh ankete.example.hr"
  DOMAIN="ankete.${IP//./-}.sslip.io"
fi

# ---------------------------------------------------------------- paketi
log "Instaliram potrebne pakete"
apt-get update -y
apt-get install -y git curl ca-certificates gnupg python3-venv python3-pip libgl1 libglib2.0-0 \
  debian-keyring debian-archive-keyring apt-transport-https iptables-persistent
if ! command -v caddy >/dev/null; then
  log "Instaliram Caddy (web poslužitelj s automatskim HTTPS-om)"
  curl -1sLf 'https://dl.cloudsmith.io/public/caddy/stable/gpg.key' | gpg --dearmor --yes -o /usr/share/keyrings/caddy-stable-archive-keyring.gpg
  curl -1sLf 'https://dl.cloudsmith.io/public/caddy/stable/debian.deb.txt' > /etc/apt/sources.list.d/caddy-stable.list
  apt-get update -y
  apt-get install -y caddy
fi

# ---------------------------------------------------------------- program
log "Preuzimam program s GitHuba u $APP"
id omr >/dev/null 2>&1 || useradd --system --home-dir "$DATA" --shell /usr/sbin/nologin omr
if [[ -d "$APP/.git" ]]; then
  git -c safe.directory="$APP" -C "$APP" fetch --depth 1 origin main
  git -c safe.directory="$APP" -C "$APP" reset --hard origin/main
else
  rm -rf "$APP"
  git clone --depth 1 "$REPO" "$APP"
fi
chmod -R a+rX "$APP"

log "Instaliram Python biblioteke (prvi put nekoliko minuta)"
[[ -x "$VENV/bin/python" ]] || python3 -m venv "$VENV"
"$VENV/bin/pip" install --quiet --upgrade pip
"$VENV/bin/pip" install --quiet -r "$APP/requirements.txt"

mkdir -p "$DATA"
chown -R omr:omr "$DATA"
chmod 700 "$DATA"

# ---------------------------------------------------------------- servis
log "Pokrećem program kao servis (najviše 1,5 procesora i 6 GB memorije)"
cat > /etc/systemd/system/omr-ankete.service <<UNIT
[Unit]
Description=OMR citac anketa (Streamlit)
After=network.target

[Service]
User=omr
Group=omr
WorkingDirectory=$APP
Environment=OMR_SERVER=1
Environment=OMR_DATA_DIR=$DATA
Environment=COMPUTERNAME=oracle-posluzitelj
Environment=HOME=$DATA
Environment=PYTHONDONTWRITEBYTECODE=1
ExecStart=$VENV/bin/streamlit run omr_aplikacija.py --server.address 127.0.0.1 --server.port $PORT --server.headless true --server.maxUploadSize 1000 --browser.gatherUsageStats false
Restart=always
RestartSec=3
CPUQuota=150%
MemoryMax=6G
NoNewPrivileges=true
ProtectSystem=strict
ProtectHome=true
ReadWritePaths=$DATA
PrivateTmp=true

[Install]
WantedBy=multi-user.target
UNIT
systemctl daemon-reload
systemctl enable omr-ankete >/dev/null
systemctl restart omr-ankete

# ---------------------------------------------------------------- lozinka
if [[ $NEW_PASSWORD -eq 1 || ! -s "$PASSFILE" ]]; then
  PASSWORD="$(python3 -c "import secrets; a='abcdefghjkmnpqrstuvwxyz23456789'; print(''.join(secrets.choice(a) for _ in range(12)))")"
  printf 'Adresa: https://%s\nKorisničko ime: %s\nLozinka: %s\n' "$DOMAIN" "$WEB_USER" "$PASSWORD" > "$PASSFILE"
  chmod 600 "$PASSFILE"
else
  PASSWORD="$(sed -n 's/^Lozinka: //p' "$PASSFILE")"
  sed -i "s|^Adresa: .*|Adresa: https://$DOMAIN|" "$PASSFILE"
fi
HASH="$(caddy hash-password --plaintext "$PASSWORD")"

# ---------------------------------------------------------------- Caddy
log "Dodajem adresu $DOMAIN u Caddy (ostale stranice na poslužitelju ostaju kakve jesu)"
mkdir -p /etc/caddy/sites
cat > "$SITE" <<CADDY
$DOMAIN {
	encode zstd gzip
	# Javne stranice za Google (početna stranica i pravila privatnosti) - bez lozinke
	handle /info* {
		root * $APP/server/web
		file_server
	}
	handle {
		basic_auth {
			$WEB_USER $HASH
		}
		reverse_proxy 127.0.0.1:$PORT
	}
	header Strict-Transport-Security "max-age=31536000"
}
CADDY
grep -qF "import /etc/caddy/sites/*.caddy" /etc/caddy/Caddyfile 2>/dev/null || \
  printf '\n# Dodatne aplikacije na ovom poslužitelju (npr. OMR ankete)\nimport /etc/caddy/sites/*.caddy\n' >> /etc/caddy/Caddyfile
if ! caddy validate --config /etc/caddy/Caddyfile --adapter caddyfile >/tmp/omr-caddy-check.txt 2>&1; then
  rm -f "$SITE"
  cat /tmp/omr-caddy-check.txt
  fail "Caddy postavke nisu valjane - stranica nije dodana, ostale aplikacije rade kao prije."
fi
systemctl reload caddy 2>/dev/null || systemctl restart caddy

for p in 80 443; do
  iptables -C INPUT -p tcp -m state --state NEW --dport "$p" -j ACCEPT 2>/dev/null || iptables -I INPUT 1 -p tcp -m state --state NEW --dport "$p" -j ACCEPT
done
netfilter-persistent save >/dev/null 2>&1 || true

# ---------------------------------------------------------------- provjera
log "Provjeravam"
ok=0
for i in $(seq 1 30); do
  if curl -fs -m 3 "http://127.0.0.1:$PORT/_stcore/health" >/dev/null; then ok=1; break; fi
  sleep 2
done
[[ $ok -eq 1 ]] && echo "Program radi." || echo "Program se još ne javlja. Pogledajte: journalctl -u omr-ankete -n 50"

echo
echo "==================================================================="
echo " Gotovo. Otvorite:  https://$DOMAIN"
echo "   korisničko ime:  $WEB_USER"
echo "   lozinka:         $PASSWORD"
echo " (Lozinka je spremljena; prikaz: sudo cat $PASSFILE)"
echo
echo " Prvi put: u lijevom stupcu učitajte google_credentials.json i prijavite se u Google Drive."
echo " Nova verzija programa:  sudo bash setup_omr.sh   (isti poziv ponovno)"
echo " Dnevnik programa:       journalctl -u omr-ankete -f"
echo " Ako se ponovno pokrene setup za Career tests, nakon njega pokrenite i ovu skriptu."
echo "==================================================================="
