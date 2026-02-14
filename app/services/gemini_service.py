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
             - 記入例はシステムが画面に自動表示するため、AIがテキストで引用・再掲することは禁止です。
             - `[[COMPLETED:xxx]]` マーカーは出力しないでください。システムが自動的に処理します。
        4. **業種に応じた参考例の使用**: 【参考資料】の記入例はドラフト作成時の文体・粒度の参考にするためのものです。AIが独自に業種を判断・変更してはいけません。記入例のテキストをそのままユーザーへの返答に含めないでください（システムが画面に表示します）。
        5. **記入例の非表示**: 記入例はシステムが画面上のカードとして自動表示します。AIは記入例を返答中に引用・転記してはいけません。各項目のヒアリング開始時は「〇〇についてお伺いします。あなたの場合はいかがですか？」のように、質問のみを返してください。
        6. **経営者の略歴のドラフト（絶対厳守）**:
           「経営者の略歴等」のドラフト作成時は、以下のルールを必ず守ってください。

           【許可されること】
           - 西暦（例: 2000年4月）を和暦（例: 平成12年4月）に変換する
           - ユーザーが書いた内容をそのままドラフトに含める

           【絶対に禁止されること】
           - ユーザーが書いていない「終了年月」「退職年月」を推測・追記すること
           - ユーザーが書いていない「～」（期間の区切り）を日付の後に付け加えること
           - ユーザーが書いていない行・エントリー（「現在：創業準備中」など）を追加すること
           - 前後のエントリーから在職期間を計算して補完すること

           【正しいドラフト例】
           ユーザー入力:「2000年4月 A社入社（〇〇に従事）」
           正しいドラフト:「平成12年4月：A社入社（〇〇に従事）」
           誤ったドラフト:「平成12年4月～平成28年3月：A社入社（〇〇に従事）」（終了年月を追加→禁止）
           誤ったドラフト:「平成12年4月～：A社入社（〇〇に従事）」（「～」を追加→禁止）

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

    async def verify_section_draft(
        self, section_name: str, draft: str, example: str
    ) -> dict:
        """
        セクションのドラフト内容を記入例と比較して改善点を確認します。
        チャットセッションとは独立した単発のAPIコールで検証を行います。

        Args:
            section_name (str): セクション名（例: "創業の動機"）
            draft (str): ユーザーのドラフト内容
            example (str): 参考となる記入例

        Returns:
            dict: {"has_issues": bool, "feedback": str}
        """
        if not self.is_initialized:
            return {"has_issues": False, "feedback": ""}

        # 「なし」系の回答は検証不要（借入なし・関連企業なし 等）
        _draft_normalized = draft.strip().replace("　", "").replace(" ", "")
        _NASHI_PATTERNS = {"なし", "特になし", "該当なし", "ありません", "ない", "なし。", "特になし。", "該当なし。"}
        if _draft_normalized in _NASHI_PATTERNS:
            return {"has_issues": False, "feedback": ""}

        prompt = f"""
以下は創業計画書の「{section_name}」セクションのドラフトと、参考となる記入例です。

【ユーザーのドラフト】
{draft}

【参考となる記入例】
{example}

記入例を「合格ライン」として、ドラフトが合格ラインを明確に下回っている場合のみフィードバックしてください。

判定基準（すべて満たす場合のみ FEEDBACK）：
- 記入例に含まれている重要な項目・情報カテゴリが、ドラフトから完全に抜け落ちている
- かつ、その欠落がこのセクションの目的を損なう程度である

以下は指摘しないでください：
- 記入例より記載量が少ない、簡潔である（量の違いは問題なし）
- 具体的な数値や社名が入っていない（個人情報・未確定事項は書かないケースが多い）
- 表現・文体・構成が記入例と異なる
- ユーザーの事業内容が記入例の業種と異なる（当然の違い）
- 少しでも改善余地がある程度の指摘

迷ったら RESULT: OK にしてください。

以下のフォーマットで回答してください：
- **や*などのMarkdown記法は使用しないでください。プレーンテキストで記述してください。

問題がない場合：
RESULT: OK

合格ラインを明確に下回っている場合のみ：
RESULT: FEEDBACK
FEEDBACK:
・（欠落している重要項目と、なぜ必要かの説明）
"""
        try:
            response = await self.model.generate_content_async(prompt)
            return self._parse_section_verification(response.text)
        except Exception as e:
            logger.error(f"Verify section draft failed: {e}")
            return {"has_issues": False, "feedback": ""}

    def _parse_section_verification(self, text: str) -> dict:
        """
        セクション検証結果テキストをパースして構造化データに変換します。
        """
        lines = text.strip().split("\n")
        has_issues = False
        feedback_lines = []
        in_feedback = False

        for line in lines:
            line = line.strip()
            if line.startswith("RESULT:"):
                if "FEEDBACK" in line:
                    has_issues = True
            elif line == "FEEDBACK:":
                in_feedback = True
            elif in_feedback and line.startswith("・"):
                feedback_lines.append(line)

        feedback = "\n".join(feedback_lines) if feedback_lines else ""
        return {"has_issues": has_issues, "feedback": feedback}

    async def verify_business_plan(self, sections: dict) -> dict:
        """
        創業計画書の内容を検証し、整合性の問題や不足事項を確認します。
        チャットセッションとは独立した単発のAPIコールで検証を行います。

        Args:
            sections (dict): セクションキーと内容のdict

        Returns:
            dict: {"status": "ok"|"issues"|"error", "issues": [{"severity": "error"|"warning", "text": str}]}
        """
        if not self.is_initialized:
            return {"status": "error", "issues": []}

        section_labels = {
            "motivation": "1. 創業の動機",
            "background": "2. 経営者の略歴等",
            "service": "3. 取扱商品・サービス",
            "employees": "4. 従業員",
            "partners": "5. 取引先・取引関係等",
            "related_companies": "6. 関連企業",
            "loans": "7. お借入の状況",
            "funds": "8. 必要な資金と調達方法",
            "outlook": "9. 事業の見通し",
            "free_description": "10. 自由記述欄",
        }

        sections_text = ""
        for key, label in section_labels.items():
            content = sections.get(key) or "（未記入）"
            sections_text += f"【{label}】\n{content}\n\n"

        prompt = f"""
以下は日本政策金融公庫への創業計画書の各セクションの内容です。

{sections_text}

この創業計画書を以下の観点で検証してください：
1. 数値の整合性（資金調達の合計と内訳が一致しているか、売上・経費・利益の計算が妥当か）
2. 必須項目の不足（重要な情報が未記入または極端に薄い項目）
3. 論理的な矛盾（異なるセクション間で矛盾する記述がないか）

以下のフォーマットで回答してください：

問題がない場合：
STATUS: OK

問題がある場合：
STATUS: ISSUES
ISSUES:
- [重要] （重大な問題の説明）
- [確認] （確認が推奨される点の説明）

ISSUESの項目は3件以内にまとめてください。「未記入」のセクションへの指摘は不要です。
"""
        try:
            response = await self.model.generate_content_async(prompt)
            return self._parse_verification_result(response.text)
        except Exception as e:
            logger.error(f"Verify business plan failed: {e}")
            return {"status": "error", "issues": []}

    def _parse_verification_result(self, text: str) -> dict:
        """
        AIの検証結果テキストをパースして構造化データに変換します。
        """
        lines = text.strip().split("\n")
        status = "ok"
        issues = []
        in_issues = False

        for line in lines:
            line = line.strip()
            if line.startswith("STATUS:"):
                if "ISSUES" in line:
                    status = "issues"
            elif line == "ISSUES:":
                in_issues = True
            elif in_issues and line.startswith("-"):
                issue_text = line[1:].strip()
                if "[重要]" in issue_text:
                    severity = "error"
                    body = issue_text.replace("[重要]", "").strip()
                elif "[確認]" in issue_text:
                    severity = "warning"
                    body = issue_text.replace("[確認]", "").strip()
                else:
                    severity = "warning"
                    body = issue_text
                if body:
                    issues.append({"severity": severity, "text": body})

        return {"status": status, "issues": issues}

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
