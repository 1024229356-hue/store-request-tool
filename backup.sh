#!/usr/bin/env bash
set -euo pipefail

cd "$(dirname "$0")"

backup_dir="backups/$(date +%Y%m%d_%H%M%S)"
mkdir -p "$backup_dir"

if [ -f "data/tickets.db" ]; then
    cp "data/tickets.db" "$backup_dir/tickets.db"
else
    echo "Warning: data/tickets.db not found, skipped database backup."
fi

if [ -d "uploads" ]; then
    cp -a "uploads" "$backup_dir/uploads"
else
    mkdir -p "$backup_dir/uploads"
    echo "Warning: uploads/ not found, created empty uploads backup directory."
fi

echo "Backup complete: $backup_dir"
