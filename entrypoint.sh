#!/bin/sh
# Render entrypoint: seed /app/data with bundled xlsx files on first run,
# when the persistent disk is empty.
set -e

SEED_DIR="/app/seed"
DATA_DIR="/app/data"

mkdir -p "$DATA_DIR"

# If the data dir has no xlsx files, copy the bundled seed
# (disk is fresh OR the bundled files were lost — both cases handled)
if [ -z "$(ls -A "$DATA_DIR" 2>/dev/null)" ]; then
    echo "[entrypoint] Empty data dir, copying seed xlsx files..."
    if [ -d "$SEED_DIR" ]; then
        cp -n "$SEED_DIR"/*.xlsx "$DATA_DIR"/ 2>/dev/null || true
    fi
else
    echo "[entrypoint] Data dir has files, leaving as-is."
fi

ls -la "$DATA_DIR"

# Hand off to the original CMD
exec "$@"
