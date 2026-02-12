"""
Google Gemini API (Vertex AI) との連携を担当するサービスモジュールです。

チャットセッションの管理、プロンプトの構築、モデルからの応答生成、
および事業計画書の生成ロジックをカプセル化しています。
"""

import logging
import os
import traceback

import vertexai
from vertexai.generative_models import Content, GenerativeModel, Part

# ロガーの設定
logger = logging.getLogger(__name__)


class GeminiService:
    """
    Google Vertex AI Geminiモデルへのインターフェースを提供するクラスです。

    チャットセッションの状態管理は簡易的にメモリ上で行いますが、
    永続化層（SessionStore）からの復元にも対応しています。
    """

    def __init__(self, project_id: str = None, location: str = None):
        """
        GeminiServiceを初期化します。

        環境変数 `GOOGLE_CLOUD_PROJECT` および `GOOGLE_CLOUD_LOCATION` が設定されていることを前提とします。

        Args:
            project_id (str, optional): Google Cloud Project ID. Defaults to None.
            location (str, optional): Google Cloud Region. Defaults to None.
        """
        self.project_id = project_id or os.getenv("GOOGLE_CLOUD_PROJECT")
        self.location = location or os.getenv("GOOGLE_CLOUD_LOCATION", "us-central1")

        # Vertex AIの初期化
        # 認証情報が設定されていない環境でのエラーを防ぐため、try-exceptで囲む（開発用）
        try:
            vertexai.init(project=self.project_id, location=self.location)
            # モデルIDを完全に指定して安定させる
            self.model = GenerativeModel("gemini-2.0-flash-lite-001")

            # 会話履歴を保持する辞書 (Session ID -> ChatSession)
            self.sessions = {}

            self.is_initialized = True
        except Exception as e:
            logger.error(f"Vertex AI initialization failed: {e}")
            self.is_initialized = False

    async def start_chat_session(
        self, session_id: str, initial_task: str = None, examples_text: str = None
    ) -> str:
        """
        新しいチャットセッションを開始します。

        システムプロンプトを設定し、必要であれば特定のタスク（initial_task）に焦点を当てた
        初期メッセージを生成して返します。

        Args:
            session_id (str): セッションID
            initial_task (str, optional): チャット開始時にフォーカスするタスクID（例: "motivation"）

        Returns:
            str: AIからの初期挨拶メッセージ
        """
        if not self.is_initialized:
            return "（システムエラー: Google Cloudへの接続設定が必要です。）"

        system_instruction = """
        あなたは起業家を支援するAIエージェント「Founder's Pilot」です。
        日本政策金融公庫の創業計画書を作成するために、ヒアリングを行い、内容を具体化します。

        【重要：進め方の方針】
        ユーザーは「もっと少ないやり取りで結論を出したい」と望んでいます。
        質問は最小限にし、早めに「ドラフト案」を提示して合意形成を図ってください。

        【絶対禁止事項】
        - Markdown記法は一切使用しないでください。
        - **太字**、# 見出し、- リスト、```コードブロック```、*斜体* 等は禁止です。
        - 全てプレーンテキストを使用し、箇条書きが必要な場合は「・」(中黒)や番号記述のみを用いてください。
        - 強調が必要な場合は「」(かぎかっこ)で囲むか、そのまま記述してください。
        - 不要な空行や改行を入れないでください。段落の区切りは1行の改行のみとし、複数の空行は使用しないでください。

        【特殊マーカーの使用】
        以下のマーカーは、システムがUIを生成するために使用します。適切な場面で必ず使用してください：

        1. `[[INDUSTRY_SELECTOR]]`: 業種選択ボタンを表示します。初回の挨拶時に使用してください。
        2. `[[DRAFT_PROPOSED]]`: OKボタンと修正ボタンを表示します。ドラフト提示時に使用してください。
        3. `[[CONTENT_START]]`～`[[CONTENT_END]]`: ドラフト内容を囲みます。ドラフト提示時に使用してください。

        【行動ルール】
        1. **ドラフトの提示（必須形式）**: ヒアリング内容から創業計画書に記載する文章（ドラフト）を作成し、「以下の内容でよろしいでしょうか？」と提案してください。
             - 重要1: ドラフト内容は必ず `[[CONTENT_START]]` と `[[CONTENT_END]]` で囲んでください。
             - 重要2: 応答の最後に必ず `[[DRAFT_PROPOSED]]` を付けてください。
             - 形式例:
               以下の内容でよろしいでしょうか？
               [[CONTENT_START]]
               （ドラフト内容をここに記述）
               [[CONTENT_END]]
               [[DRAFT_PROPOSED]]
        2. **一問一答の原則**: 左側のタスクリストのチェックボックスを一つずつ埋めていくため、**必ず一つの項目ずつ**ヒアリングとドラフト提示を行ってください。「創業の動機」と「略歴」をまとめて聞く等は禁止です。
        3. **承認後の遷移**: ユーザーから「この内容でOKです」「OKです」「OK」「承認します」などの承認メッセージを受け取った場合、**次の項目のヒアリングに移行してください**。
             - 承認されたドラフトを再掲しないでください。
             - 次の項目の記入例を示しながらヒアリングを開始してください。
             - `[[COMPLETED:xxx]]` マーカーは出力しないでください。システムが自動的に処理します。
        4. **業種に応じた参考例の使用**: ユーザーが「〇〇業を選択しました」と伝えてきた業種の参考例のみを使用してください。AIが独自に業種を判断・変更してはいけません。システムがメッセージ内に「記入例（参考）」として該当セクションの例を提示します。それを優先して使用してください。
        5. **記入例の提示（絶対必須・スキップ禁止）**: 各項目についてヒアリングを開始する際、**必ず**選択した業種の参考例から該当セクションを引用し、「ソフトウェア開発業の場合はこのように記載します：〔引用〕」という形式で示してください。
             - 重要: 記入例の提示は全項目で必須です。1項目たりともスキップしてはいけません。
             - 記入例を示した後に「あなたの場合はいかがですか？」と質問してください。
        トーンはお客様に寄り添うプロフェッショナルな姿勢で、しかし簡潔に。

        【参考資料：創業計画書の記入例】
        ユーザーの事業に最も近い業種の例を選択して使用し、ドラフト作成時の文体や具体性の参考にしてください。内容はユーザーの回答に合わせて完全にカスタマイズしてください。

        {examples_section}
        """

        # DBから渡された記入例テキストでプレースホルダーを置換
        system_instruction = system_instruction.format(
            examples_section=examples_text or "（記入例データを読み込めませんでした）"
        )

        # 新しいチャットセッションを開始
        chat_session = self.model.start_chat()
        self.sessions[session_id] = chat_session

        # 最初のプロンプトを作成
        first_prompt = f"{system_instruction}\n\n"
        if initial_task:
            first_prompt += f"ユーザーは「{initial_task}」の項目を作成したいと考えています。まずはこの項目についてヒアリングを開始してください。"
        else:
            first_prompt += """まずは最初の挨拶として、簡潔な自己紹介をしてください。

その後、必ず以下のマーカーを出力してください（このマーカーがあると、システムが業種選択ボタンを表示します）：

[[INDUSTRY_SELECTOR]]

マーカーの前に「まずは、あなたの事業に最も近い業種を選択してください。」などの簡単な説明を加えてください。
マーカーの後には何も書かないでください。"""

        # システムプロンプトを最初のメッセージとして送信（または内部的に保持）
        try:
            response = await chat_session.send_message_async(first_prompt)
            return response.text
        except Exception as e:
            logger.error(f"Start chat session failed: {e}")
            return "申し訳ありません。チャットの開始に失敗しました。"

    async def generate_response(self, session_id: str, user_message: str) -> str:
        """
        継続中のセッションでユーザーのメッセージに対する応答を生成します。
        """
        if not self.is_initialized:
            return "（システムエラー: Google Cloudへの接続設定が必要です。）"

        chat_session = self.sessions.get(session_id)
        if not chat_session:
            # セッションがない場合は新規作成
            return await self.start_chat_session(session_id)

        try:
            # セッションを使ってメッセージを送信（履歴は自動で維持される）
            response = await chat_session.send_message_async(user_message)
            return response.text
        except Exception as e:
            logger.error(f"Generate response failed: {e}")
            traceback.print_exc()
            return f"申し訳ありません。エラーが発生しました。（詳細: {str(e)}）"

    async def generate_business_plan(self, session_id: str) -> str:
        """
        これまでの会話内容を基に、創業計画書の全体ドラフトを作成します。
        """
        if not self.is_initialized:
            return "（システムエラー: Google Cloudへの接続設定が必要です。）"

        chat_session = self.sessions.get(session_id)
        if not chat_session:
            return "エラー: セッションが見つかりません。"

        prompt = """
        これまでのヒアリング内容を基に、日本政策金融公庫の創業計画書の形式で、全体のドラフトを作成してください。
        以下の項目を含めて適切に構造化してください：
        1. 創業の動機
        2. 経営者の略歴等
        3. 取扱商品・サービス
        4. 従業員
        5. 取引先・取引関係等
        6. 関連企業
        7. お借入の状況
        8. 必要な資金と調達方法
        9. 事業の見通し
        10. 自由記述欄

        出力形式は、各項目を見出しとして、内容を詳細に記述したプレーンテキストでお願いします。
        （マークダウンは見出し程度なら可としますが、過度な装飾は避けてください。読みやすい形式でお願いします。）
        """
        try:
            response = await chat_session.send_message_async(prompt)
            return response.text
        except Exception as e:
            logger.error(f"Generate business plan failed: {e}")
            return "申し訳ありません。創業計画書の作成に失敗しました。"

    def get_chat_history(self, session_id: str) -> list:
        """
        セッションのチャット履歴を取得します（JSON形式でシリアライズ用に変換）。
        """
        chat_session = self.sessions.get(session_id)
        if not chat_session:
            return []

        history = []
        for content in chat_session.history:
            parts = []
            for part in content.parts:
                parts.append({"text": part.text})
            history.append({"role": content.role, "parts": parts})
        return history

    def restore_chat_session(self, session_id: str, history_data: list):
        """
        履歴データからチャットセッションを復元します。
        """
        if not self.is_initialized:
            return

        # Historyオブジェクトの再構築
        history_objects = []
        for item in history_data:
            parts = [Part.from_text(p["text"]) for p in item["parts"]]
            content = Content(role=item["role"], parts=parts)
            history_objects.append(content)

        # 履歴付きでセッションを開始
        self.sessions[session_id] = self.model.start_chat(history=history_objects)

    def reset_chat_session(self, session_id: str):
        """
        指定されたセッションIDのチャット履歴をリセットします。
        """
        if session_id in self.sessions:
            del self.sessions[session_id]
