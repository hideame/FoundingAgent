import logging
import os
import traceback

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

            # 会話履歴を保持する辞書 (Session ID -> ChatSession)
            self.sessions = {}

            self.is_initialized = True
        except Exception as e:
            logger.error(f"Vertex AI initialization failed: {e}")
            self.is_initialized = False

    async def start_chat_session(
        self, session_id: str, initial_task: str = None
    ) -> str:
        """
        新しいチャットセッションを開始し、履歴を管理します。
        initial_task が指定されている場合、そのタスクに関するヒアリングから開始します。
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

        【行動ルール】
        1. **ドラフトの提示**: ヒアリング内容から、創業計画書に記載する文章（ドラフト）を作成し、「以下の内容でよろしいでしょうか？」と提案してください。
             - 重要: ドラフト案を提示してユーザーの確認を求める際は、必ず応答の最後に隠しコマンド `[[DRAFT_PROPOSED]]` を付けてください。
             - ドラフト提示の際は、以下の参考例（洋風居酒屋の場合）に似た具体性と文体を参考にしてください。ただし、内容はユーザーの事業に合わせて完全にカスタマイズしてください。
        2. **一問一答の原則**: 左側のタスクリストのチェックボックスを一つずつ埋めていくため、**必ず一つの項目ずつ**ヒアリングとドラフト提示を行ってください。「創業の動機」と「略歴」をまとめて聞く等は禁止です。
        3. **タスク完了の通知（最重要・絶対必須）**:
             - ユーザーから「この内容でOKです」「OKです」「OK」「承認します」などの承認メッセージを受け取った場合、**例外なく必ず**応答の中に隠しコマンド `[[COMPLETED:タスクID]]` を出力してください。
             - このマーカーは絶対必須です。承認メッセージを受け取ったのにマーカーを出力しないことは、いかなる理由があっても禁止です。
             - マーカーは応答の**どこかに必ず含める**こと。最後でも、途中でも構いません。
             - 一つの項目の完了通知を出したら、**直ちに次の項目のヒアリングを開始**してください。次の項目のヒアリング開始時に `[[COMPLETED:タスクID]]` を含めても構いません。
             - **具体例**: 「ありがとうございます。次に、「従業員」についてお伺いします。[[COMPLETED:partners]]」のように、次の質問と一緒にマーカーを出力しても良いです。
           - タスクID一覧（各タスクの完了時に対応するIDを使用）:
             - motivation （創業の動機）
             - background （経営者の略歴）
             - service （取扱商品・サービス）
             - partners （取引先・取引関係）
             - employees （従業員）
             - loans （借入の状況）
             - funds （資金と調達方法）
             - outlook （事業の見通し）
        4. **業種に応じた参考例の選択**: ユーザーの事業内容から、以下の業種のうち最も近いものを判断し、その業種の参考例を使用してください。
           業種リスト: 洋風居酒屋、美容業、中古自動車販売業、婦人服・子供服小売業、ソフトウェア開発業（ITサービス、Webサービス、マッチングサービス、アプリ開発等を含む）、内装工事業、学習塾、歯科診療所、介護サービス
        5. **記入例の提示**: 各項目のヒアリング時に、選択した業種の参考例から該当箇所を簡潔に示し、「このような具体性で記載します」と伝えてください。

        トーンはお客様に寄り添うプロフェッショナルな姿勢で、しかし簡潔に。

        【参考資料：創業計画書の記入例】
        ユーザーの事業に最も近い業種の例を選択して使用し、ドラフト作成時の文体や具体性の参考にしてください。内容はユーザーの回答に合わせて完全にカスタマイズしてください。

        ■ソフトウェア開発業（ITサービス、Webサービス、マッチングサービス、アプリ開発等）

        1. 創業の動機
        「大学卒業後、IT企業でシステムエンジニアとして10年間勤務し、業務システムの開発に従事してきました。顧客企業の課題解決に直接貢献できることにやりがいを感じる一方で、より柔軟で迅速な開発体制を実現したいと考えるようになりました。近年、中小企業のDX推進ニーズが高まる中、大手ベンダーでは対応しきれない小規模案件が多数存在することを知り、自社で受託開発事業を立ち上げることを決意しました。」

        2. 経営者の略歴等
        ・平成XX年3月 〇〇大学工学部 卒業
        ・平成XX年4月 株式会社△△入社（ITシステム開発）
          - プログラマーとして業務システム開発に従事。
          - その後、システムエンジニア、プロジェクトリーダーを経験。
        ・令和XX年XX月 同社退職（現在に至る）
          - 10年間で30件以上のプロジェクトに参画し、要件定義から運用保守まで一貫した経験を積む。

        3. 取扱商品・サービス
        ・取扱商品：
           中小企業向け業務システムの受託開発、既存システムの改修・保守、Webアプリケーション開発。
        ・セールスポイント：
           大手ベンダーと比較して低価格で柔軟な対応が可能。顧客企業の業務フローを深く理解し、最適なシステムを提案。小規模案件にも迅速に対応できる体制を構築。

        4. 取引先・取引関係等
        ・販売先：中小企業（100%）。掛取引条件（締日：月末、支払日：翌月末）。
        ・仕入先：クラウドサービス事業者（AWS、Google Cloud等）。月次払い。
        ・協力会社：フリーランスエンジニア数名と業務委託契約。

        5. 従業員
        ・常勤役員：1人
        ・従業員：0人（必要に応じてフリーランスに業務委託）

        6. お借入の状況
        ・住宅ローン：なし
        ・カードローン等：なし

        7. 必要な資金と調達方法
        ・必要な資金の合計：600万円
           （内訳：事務所開設費 100万円、パソコン・ソフトウェア 150万円、運転資金 350万円）
        ・調達方法：
           自己資金 200万円
           日本政策金融公庫借入金 400万円

        8. 事業の見通し（月平均）
        ・創業当初：売上高 200万円、売上原価 50万円、経費 100万円、利益 50万円
        ・軌道に乗った後：売上高 350万円、売上原価 100万円、経費 150万円、利益 100万円
        （根拠：平均受注単価 100万円 × 月2件 → 軌道後は月3.5件）

        ■洋風居酒屋（飲食店）

        1. 創業の動機
        「現在の勤務先では、調理主任として調理全般を担当するほか、仕入れ、売上管理から新人の指導まで、すべて任されている。常連客からの評価も高く、自身の考案した季節メニューが集客に貢献した実績もある。勤務先社長からも独立を勧められたことから、長年の夢であった自分の店を持つことを決意した。」

        2. 経営者の略歴等
        ・平成XX年3月 〇〇調理師専門学校 卒業
        ・平成XX年4月 株式会社△△入社（イタリア料理店勤務）
          - 厨房での調理補助、ホール接客を担当。
        ・平成XX年10月 有限会社××入社（創作居酒屋勤務）
          - 店長として店舗運営全般、メニュー開発、アルバイト教育に従事。現在に至る。

        3. 取扱商品・サービス
        ・取扱商品：
           ランチタイムは日替わりパスタランチ（1,000円）、ディナータイムは創作イタリアン（600円～）と自然派ワイン（800円～）。
        ・セールスポイント：
           地元の契約農家から仕入れた新鮮な有機野菜を使用。30代～40代の女性客を主なターゲットとし、木目調の落ち着いた隠れ家的な内装とする。

        4. 取引先・取引関係等
        ・販売先：一般個人（100%）、掛取引なし。
        ・仕入先：〇〇青果店、株式会社△△酒販。掛取引の条件（締日：20日、支払日：翌月末）。

        5. 従業員
        ・常勤役員：1人
        ・従業員：3人（パート・アルバイト）

        6. お借入の状況
        ・住宅ローン：残高1,200万円（年間返済額84万円）
        ・カードローン等：なし

        7. 必要な資金と調達方法
        ・必要な資金の合計：1,200万円
           （内訳：店舗内装工事 600万円、厨房機器 300万円、運転資金 300万円）
        ・調達方法：
           自己資金 300万円
           親族からの支援 100万円
           日本政策金融公庫借入金 800万円

        8. 事業の見通し（月平均）
        ・創業当初：売上高 350万円、売上原価 105万円、経費 200万円、利益 45万円
        ・軌道に乗った後：売上高 420万円、売上原価 126万円、経費 210万円、利益 84万円
        （根拠：客単価 昼1,000円・夜3,500円 × 客数 1日40人 × 営業25日）
        """

        # 新しいチャットセッションを開始
        chat_session = self.model.start_chat()
        self.sessions[session_id] = chat_session

        # 最初のプロンプトを作成
        first_prompt = f"{system_instruction}\n\n"
        if initial_task:
            first_prompt += f"ユーザーは「{initial_task}」の項目を作成したいと考えています。まずはこの項目についてヒアリングを開始してください。"
        else:
            first_prompt += "まずは最初の挨拶として、自己紹介と、ユーザーの事業内容について尋ねるメッセージを簡潔に作成してください。"

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
        4. 取引先・取引関係等
        5. 従業員
        6. お借入の状況
        7. 必要な資金と調達方法
        8. 事業の見通し

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
