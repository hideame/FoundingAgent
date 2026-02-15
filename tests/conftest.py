"""
pytest configuration file
テストの共通設定とフィクスチャ
"""

from unittest.mock import AsyncMock, MagicMock

import pytest
from fastapi.testclient import TestClient


@pytest.fixture
def test_client():
    """FastAPI TestClientのフィクスチャ"""
    from app.main import app
    return TestClient(app)


@pytest.fixture
def mock_async_db():
    """非同期データベースセッションのモック"""
    mock = MagicMock()
    mock.execute = AsyncMock()
    mock.commit = AsyncMock()
    mock.rollback = AsyncMock()
    mock.close = AsyncMock()
    mock.refresh = AsyncMock()
    mock.add = MagicMock()
    return mock


@pytest.fixture
def mock_gemini_service():
    """Gemini AIサービスのモック"""
    mock = MagicMock()
    mock.start_chat = MagicMock()
    mock.send_message = AsyncMock(return_value="モック応答")
    mock.restore_chat_session = MagicMock()
    return mock


@pytest.fixture
def mock_session_store():
    """セッションストアのモック"""
    mock = MagicMock()
    mock.load_session = AsyncMock(return_value=None)
    mock.save_session = AsyncMock()
    mock.create_session = AsyncMock(return_value="test-session-id")
    return mock
