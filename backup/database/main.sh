#!/usr/bin/env bash
#
# backup_mysql.sh
#
# Automated MySQL/MariaDB backup with gzip compression, rotation,
# and logging. Designed to be run manually or via cron.
#
# Usage:
#   ./backup_mysql.sh
#
# Config is read from a .env file in the same directory as this script.
# See .env.example for the required variables.

set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
ENV_FILE="$SCRIPT_DIR/.env"

# --- Load config -----------------------------------------------------------

if [ -f "$ENV_FILE" ]; then
  set -a
  # shellcheck disable=SC1090
  source "$ENV_FILE"
  set +a
else
  echo "❌ Missing .env file at $ENV_FILE (copy .env.example to get started)" >&2
  exit 1
fi

: "${DB_HOST:?DB_HOST is not set in .env}"
: "${DB_PORT:=3306}"
: "${DB_USER:?DB_USER is not set in .env}"
: "${DB_PASSWORD:?DB_PASSWORD is not set in .env}"
: "${DB_NAME:?DB_NAME is not set in .env}"
: "${BACKUP_DIR:=$SCRIPT_DIR/backups}"
: "${RETENTION_DAYS:=7}"

mkdir -p "$BACKUP_DIR"

TIMESTAMP="$(date +"%Y-%m-%d_%H-%M-%S")"
BACKUP_FILE="$BACKUP_DIR/${DB_NAME}_${TIMESTAMP}.sql.gz"
LOG_FILE="$BACKUP_DIR/backup.log"

log() {
  echo "[$(date +"%Y-%m-%d %H:%M:%S")] $1" | tee -a "$LOG_FILE"
}

# --- Sanity checks -----------------------------------------------------------

if ! command -v mysqldump >/dev/null 2>&1; then
  log "❌ mysqldump not found. Install the MySQL/MariaDB client tools first."
  exit 1
fi

# --- Credentials -------------------------------------------------------------
# Written to a temp file instead of passed as --password on the command line,
# so the password doesn't show up in `ps` output or shell history.

CNF_FILE="$(mktemp)"
chmod 600 "$CNF_FILE"
trap 'rm -f "$CNF_FILE"' EXIT

cat > "$CNF_FILE" <<EOF
[client]
host=$DB_HOST
port=$DB_PORT
user=$DB_USER
password=$DB_PASSWORD
EOF

# --- Backup --------------------------------------------------------------

log "🚀 Starting backup for database '$DB_NAME'..."

if mysqldump \
  --defaults-extra-file="$CNF_FILE" \
  --single-transaction \
  --quick \
  --routines \
  --triggers \
  --events \
  "$DB_NAME" | gzip > "$BACKUP_FILE"; then

  FILE_SIZE="$(du -h "$BACKUP_FILE" | cut -f1)"
  log "✅ Backup successful: $BACKUP_FILE ($FILE_SIZE)"
else
  log "❌ Backup failed for database '$DB_NAME'"
  rm -f "$BACKUP_FILE"
  exit 1
fi

# --- Rotation --------------------------------------------------------------
# Deletes backups for this database older than RETENTION_DAYS.

log "🧹 Cleaning up backups older than $RETENTION_DAYS day(s)..."

find "$BACKUP_DIR" -maxdepth 1 -name "${DB_NAME}_*.sql.gz" -type f -mtime +"$RETENTION_DAYS" -print | while read -r old_file; do
  rm -f "$old_file"
  log "🗑️  Deleted old backup: $old_file"
done

log "🎉 Done."