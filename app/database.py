"""
データベース接続とセッション管理

このモジュールは、SQLAlchemyを使用した非同期データベース接続の
セットアップとセッション管理を提供します。
"""

import os

from sqlalchemy.ext.asyncio import AsyncSession, async_sessionmaker, create_async_engine

from app.models import Base

# 環境変数からデータベースURLを取得
DATABASE_URL = os.getenv(
    "DATABASE_URL",
    "mysql+aiomysql://founding_user:founding_pass@localhost:3306/founding_agent",
)

# Cloud SQL接続名（Cloud Run環境で設定。例: project:asia-northeast1:instance）
CLOUD_SQL_CONNECTION_NAME = os.getenv("CLOUD_SQL_CONNECTION_NAME")

# connect_argsの設定（Cloud Run環境ではUnix socketを使用）
_connect_args = {"charset": "utf8mb4"}
if CLOUD_SQL_CONNECTION_NAME:
    _connect_args["unix_socket"] = f"/cloudsql/{CLOUD_SQL_CONNECTION_NAME}"

# 非同期エンジンの作成
engine = create_async_engine(
    DATABASE_URL,
    echo=os.getenv("DEBUG", "False").lower() == "true",
    pool_pre_ping=True,  # 接続プールの健全性チェック
    pool_recycle=3600,  # 1時間ごとに接続をリサイクル
    connect_args=_connect_args,
)

# 非同期セッションファクトリ
AsyncSessionLocal = async_sessionmaker(
    engine,
    class_=AsyncSession,
    expire_on_commit=False,
)


async def get_db():
    """
    FastAPI の依存性注入で使用するデータベースセッションを提供します。

    Yields:
        AsyncSession: データベースセッション

    Example:
        ```python
        @app.get("/items")
        async def read_items(db: AsyncSession = Depends(get_db)):
            result = await db.execute(select(Item))
            return result.scalars().all()
        ```
    """
    async with AsyncSessionLocal() as session:
        try:
            yield session
        finally:
            await session.close()


async def init_db():
    """
    データベーステーブルを初期化します。

    既存のテーブルがない場合、全てのモデルに基づいてテーブルを作成します。
    本番環境ではAlembicマイグレーションの使用を推奨します。

    Note:
        アプリケーション起動時に一度だけ呼び出してください。
    """
    async with engine.begin() as conn:
        # 開発環境では既存テーブルを削除して再作成（本番環境では使用しないこと）
        # await conn.run_sync(Base.metadata.drop_all)

        # 全テーブル作成
        await conn.run_sync(Base.metadata.create_all)


async def close_db():
    """
    データベース接続を閉じます。

    アプリケーション終了時に呼び出してください。
    """
    await engine.dispose()
