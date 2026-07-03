#!/usr/bin/env bash
set -euo pipefail

cd "$(dirname "$0")"

python3 -m venv .venv
. .venv/bin/activate

python -m pip install --upgrade pip
python -m pip install -r requirements.txt

mkdir -p data uploads

echo "Deploy setup complete."
echo "Start with: .venv/bin/python -m uvicorn main:app --host 0.0.0.0 --port 8701"
