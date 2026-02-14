#!/bin/bash
set -e

echo "=== テーブルを作成（SQLAlchemy create_all）==="
python -c "
import asyncio
from app.database import init_db
asyncio.run(init_db())
print('テーブル作成完了')
"

echo "=== Alembicをhead状態にスタンプ（マイグレーション済みとしてマーク）==="
alembic stamp head

echo "=== 記入例データをシード ==="
python scripts/fix_all_examples.py

echo "=== アプリケーションを起動 ==="
exec uvicorn app.main:app --host 0.0.0.0 --port 8080
