#!/bin/bash
set -e

echo "=== Alembicマイグレーションを実行 ==="
alembic upgrade head

echo "=== 記入例データをシード ==="
python scripts/fix_all_examples.py

echo "=== アプリケーションを起動 ==="
exec uvicorn app.main:app --host 0.0.0.0 --port 8080
