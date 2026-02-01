import logging
import os

import vertexai
from vertexai.generative_models import Content, GenerativeModel, Part

# ロガーの設定
logger = logging.getLogger(__name__)


class GeminiService:
    def __init__(self, project_id: str = None, location: str = None):
        self.project_id = project_id or os.getenv("GOOGLE_CLOUD_PROJECT")
        self.location = location or os.getenv("GOOGLE_CLOUD_LOCATION", "us-central1")

        # Vertex AIの初期化
        # 認証情報が設定されていない環境でのエラーを防ぐため、try-exceptで囲む（開発用）
        try:
            vertexai.init(project=self.project_id, location=self.location)
            # モデルIDを完全に指定して安定させる
            self.model = GenerativeModel("gemini-2.0-flash-lite-001")
            self.is_initialized = True
        except Exception as e:
            logger.error(f"Vertex AI initialization failed: {e}")
            self.is_initialized = False

    async def generate_greeting(self) -> str:
        """
        チャット開始時の挨拶メッセージを生成します。
        """
        if not self.is_initialized:
            return "（システムエラー: Google Cloudへの接続設定が必要です。）<br>こんにちは！創業計画書の作成をお手伝いします。"

        system_instruction = """
        あなたは起業家を支援するAIエージェント「Founder's Pilot」です。
        日本政策金融公庫の創業計画書を作成するために、ユーザーに優しく寄り添いながらヒアリングを行います。

        まずは最初の挨拶として、自己紹介と、ユーザーの事業内容について尋ねるメッセージを簡潔に作成してください。
        トーン＆マナー：プロフェッショナルだが親しみやすい、励ますような口調。
        """

        try:
            prompt = f"{system_instruction}\n\nUser: 挨拶をお願いします。\nAgent:"
            response = await self.model.generate_content_async(prompt)
            return response.text
        except Exception as e:
            logger.error(f"Generate greeting failed: {e}")
            return f"申し訳ありません。現在AIシステムにアクセスできません。<br>エラー詳細: {str(e)}<br>地域: {self.location}<br>もう一度お試しください。"
