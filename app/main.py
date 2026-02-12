import re
import uuid
import zipfile
from io import BytesIO
from pathlib import Path

import openpyxl
from dotenv import load_dotenv
from fastapi import Cookie, Depends, FastAPI, Form, Request, Response
from fastapi.responses import HTMLResponse, StreamingResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
from openpyxl.utils import get_column_letter, range_boundaries
from sqlalchemy.ext.asyncio import AsyncSession

from sqlalchemy import select

from app.database import close_db, get_db, init_db
from app.models import ExampleContent, Session as SessionModel
from app.services.gemini_service import GeminiService
from app.store import session_store

"""
Founder's Pilot アプリケーションのメインエントリポイントです。

FastAPIを使用したWebサーバーとして機能し、以下の役割を担います。
- 静的ファイルとHTMLテンプレートの配信
- クライアント（ブラウザ）からのチャットメッセージの受信とAI応答の返却
- セッション管理（Cookieを使用）
- 創業計画書の編集、保存、生成
- 完成した計画書（Excel）と記入例（PDF）のZIPダウンロード提供

主なエンドポイント:
- `/`: アプリケーションのルート（ダッシュボード）
- `/chat/*`: チャット機能関連（開始、メッセージ送信）
- `/plan/*`: 計画書の編集、保存、生成、ダウンロード
"""

# Load environment variables first
load_dotenv()

app = FastAPI(
    title="Founder's Pilot",
    description="""
    ## 創業計画書作成支援アプリケーション

    AIエージェントとの対話を通じて、創業計画書を作成するためのWebアプリケーションです。

    ### 主な機能
    - **AIチャット**: Gemini APIを使用した対話型ヒアリング
    - **計画書作成**: 10項目の創業計画書を段階的に作成
    - **編集機能**: 作成した計画書の編集・保存
    - **Excelエクスポート**: 日本政策金融公庫のテンプレート形式でダウンロード
    - **記入例提供**: 業種別の記入例PDFを自動選択して提供

    ### エンドポイント
    - `/`: ダッシュボード（HTMLページ）
    - `/chat/*`: チャット機能（AIとの対話）
    - `/plan/*`: 計画書の編集・保存・生成・ダウンロード
    """,
    version="1.0.0",
    contact={
        "name": "Founder's Pilot Support",
    },
)


# アプリケーション起動時の処理
@app.on_event("startup")
async def startup_event():
    """
    アプリケーション起動時にデータベースを初期化します。
    """
    await init_db()
    print("[INFO] Database initialized successfully")


# アプリケーション終了時の処理
@app.on_event("shutdown")
async def shutdown_event():
    """
    アプリケーション終了時にデータベース接続を閉じます。
    """
    await close_db()
    print("[INFO] Database connection closed")


# Initialize Gemini Service
gemini_service = GeminiService()

# --- 定数の定義 ---
# 大項目のフロー（ステッパー用）
ROADMAP_STEPS = [
    {"id": 1, "title": "構想・準備", "status": "current"},
    {"id": 2, "title": "創業計画書作成", "status": "upcoming"},
    {"id": 3, "title": "面談対策", "status": "upcoming"},
    {"id": 4, "title": "融資実行", "status": "upcoming"},
]

# 現在のフェーズ（創業計画書）のタスクリスト
TASKS = [
    {
        "id": "motivation",
        "title": "1. 創業の動機（創業されるのは、どのような目的、動機からですか）",
        "desc": "なぜこの事業を始めるのか、熱意と目的を言語化しましょう。",
        "status": "pending",
    },
    {
        "id": "background",
        "title": "2. 経営者の略歴等（略歴については、勤務先名だけではなく、担当業務や役職、身につけた技能等についても記載してください）",
        "desc": "これまでの経験が事業にどう活きるか整理します。",
        "status": "pending",
    },
    {
        "id": "service",
        "title": "3. 取扱商品・サービス",
        "desc": "商品・サービスの強みや特徴を明確にします。",
        "status": "pending",
    },
    {
        "id": "employees",
        "title": "4. 従業員",
        "desc": "常勤役員、従業員数などの体制を計画します。",
        "status": "pending",
    },
    {
        "id": "partners",
        "title": "5. 取引先・取引関係等",
        "desc": "販売先や仕入先、掛取引の条件などを整理します。",
        "status": "pending",
    },
    {
        "id": "related_companies",
        "title": "6. 関連企業（お申込人もしくは法人代表者または配偶者の方がご経営されている企業がある場合にご記入ください）",
        "desc": "関連する企業との関係性を整理します。",
        "status": "pending",
    },
    {
        "id": "loans",
        "title": "7. お借入の状況（法人の場合、代表者の方のお借入）",
        "desc": "個人的な借り入れや住宅ローンなどの状況を確認します。",
        "status": "pending",
    },
    {
        "id": "funds",
        "title": "8. 必要な資金と調達方法",
        "desc": "設備資金・運転資金の総額と、自己資金・借入金のバランスを計算します。",
        "status": "pending",
    },
    {
        "id": "outlook",
        "title": "9. 事業の見通し（月平均）",
        "desc": "創業当初と軌道に乗った後の売上・利益予測を立てます。",
        "status": "pending",
    },
    {
        "id": "free_description",
        "title": "10. 自由記述欄（アピールポイント、事業を行ううえでの悩み、希望するアドバイス等）",
        "desc": "事業のアピールポイントや悩み、希望するアドバイスなどを記載します。",
        "status": "pending",
    },
]

# Static files and Templates setup
BASE_DIR = Path(__file__).resolve().parent
app.mount("/static", StaticFiles(directory=str(BASE_DIR / "static")), name="static")
templates = Jinja2Templates(directory=str(BASE_DIR / "templates"))


# --- ヘルパー関数 ---
def extract_sections_from_text(plan_text: str) -> dict:
    """
    創業計画書の全文テキストから各セクションを抽出します。

    Args:
        plan_text (str): 創業計画書の全文（マークダウン形式）

    Returns:
        dict: 各セクションの辞書 {"motivation": "...", "background": "...", ...}
    """
    heading_order = [
        ("motivation", "創業の動機"),
        ("background", "経営者の略歴等"),
        ("service", "取扱商品・サービス"),
        ("employees", "従業員"),
        ("partners", "取引先・取引関係等"),
        ("related_companies", "関連企業"),
        ("loans", "お借入の状況"),
        ("funds", "必要な資金と調達方法"),
        ("outlook", "事業の見通し"),
        ("free_description", "自由記述欄"),
    ]

    def build_section_pattern(current_label, next_labels):
        """見出しとその次の見出しまでの内容を抽出する正規表現パターンを生成"""
        next_part = "|".join(re.escape(label) for label in next_labels)
        if next_labels:
            return rf"^\s*#{{{0, 6}}}\s*\d*\.?\s*{re.escape(current_label)}\s*(?:\n|:|：)\s*(.*?)(?=^\s*#{{{0, 6}}}\s*\d*\.?\s*(?:{next_part}))"
        else:
            return rf"^\s*#{{{0, 6}}}\s*\d*\.?\s*{re.escape(current_label)}\s*(?:\n|:|：)\s*(.*?)$"

    def normalize_section_text(content: str) -> str:
        content = content.strip()
        content = re.sub(r"[\*#]+", "", content)
        content = re.sub(r"\n\s*\n+", "\n\n", content)
        return content.strip()

    sections = {}
    for idx, (key, label) in enumerate(heading_order):
        next_labels = [next_label for _, next_label in heading_order[idx + 1 :]]
        pattern = build_section_pattern(label, next_labels)
        match = re.search(pattern, plan_text, re.DOTALL | re.MULTILINE)
        if match:
            sections[key] = normalize_section_text(match.group(1))
        else:
            sections[key] = None

    return sections


def build_plan_text_from_sections(sections: dict) -> str:
    """
    各セクションから創業計画書の全文テキストを再構築します。

    Args:
        sections (dict): 各セクションの辞書

    Returns:
        str: 創業計画書の全文（マークダウン形式）
    """
    heading_order = [
        ("motivation", "1. 創業の動機"),
        ("background", "2. 経営者の略歴等"),
        ("service", "3. 取扱商品・サービス"),
        ("employees", "4. 従業員"),
        ("partners", "5. 取引先・取引関係等"),
        ("related_companies", "6. 関連企業"),
        ("loans", "7. お借入の状況"),
        ("funds", "8. 必要な資金と調達方法"),
        ("outlook", "9. 事業の見通し"),
        ("free_description", "10. 自由記述欄"),
    ]

    plan_text_parts = []
    for key, title in heading_order:
        content = sections.get(key)
        if content:
            plan_text_parts.append(f"## {title}\n\n{content}")

    return "\n\n".join(plan_text_parts)


def extract_single_section_from_response(ai_response: str, task_id: str) -> str | None:
    """
    AI応答から特定のタスクに対応するセクションの内容を抽出します。

    AI応答内の [[CONTENT_START]] と [[CONTENT_END]] マーカーで囲まれた
    コンテンツを取得します。

    Args:
        ai_response (str): AIからの応答テキスト
        task_id (str): タスクID（例: "motivation", "background"）

    Returns:
        str | None: 抽出されたセクション内容。見つからない場合はNone
    """
    print(f"[DEBUG] Extracting content for task: {task_id}")
    print(f"[DEBUG] AI response length: {len(ai_response)} chars")
    print(f"[DEBUG] AI response preview: {ai_response[:300]}")

    # [[CONTENT_START]] と [[CONTENT_END]] の間のコンテンツを取得
    start_marker = "[[CONTENT_START]]"
    end_marker = "[[CONTENT_END]]"

    start_pos = ai_response.find(start_marker)
    end_pos = ai_response.find(end_marker)

    if start_pos == -1 or end_pos == -1:
        print(f"[WARNING] Content markers not found in AI response")
        print(f"[WARNING]   start_marker found: {start_pos != -1}")
        print(f"[WARNING]   end_marker found: {end_pos != -1}")
        return None

    if start_pos >= end_pos:
        print(f"[ERROR] Invalid marker positions: start={start_pos}, end={end_pos}")
        return None

    # マーカー間のコンテンツを取得（マーカー自体は除外）
    content_start = start_pos + len(start_marker)
    content = ai_response[content_start:end_pos]

    # 前後の空白・改行を削除（内部の改行は保持）
    content = content.strip()

    print(f"[DEBUG] ✓ Extracted content:")
    print(f"[DEBUG]   Length: {len(content)} chars")
    print(f"[DEBUG]   Preview: {content[:200]}...")
    print(f"[DEBUG]   Contains newlines: {chr(10) in content}")
    print(f"[DEBUG]   Number of lines: {len(content.splitlines())}")

    return content if content else None


# 業種・セクションの表示名定義
INDUSTRY_DISPLAY_NAMES = {
    "software": "ソフトウェア開発業（ITサービス、Webサービス、マッチングサービス、アプリ開発等）",
    "restaurant": "洋風居酒屋（飲食店）",
    "beauty": "美容業",
    "car_sales": "中古自動車販売業",
    "apparel": "婦人服・子供服小売業",
    "construction": "内装工事業",
    "cram_school": "学習塾",
    "dentist": "歯科診療所",
    "care_service": "介護サービス",
}

SECTION_DISPLAY_NAMES = {
    "motivation": "1. 創業の動機",
    "background": "2. 経営者の略歴等",
    "service": "3. 取扱商品・サービス",
    "employees": "4. 従業員",
    "partners": "5. 取引先・取引関係等",
    "related_companies": "6. 関連企業",
    "loans": "7. お借入の状況",
    "funds": "8. 必要な資金と調達方法",
    "outlook": "9. 事業の見通し（月平均）",
}

SECTION_ORDER = [
    "motivation", "background", "service", "employees", "partners",
    "related_companies", "loans", "funds", "outlook",
]

INDUSTRY_ORDER = [
    "software", "restaurant", "beauty", "car_sales", "apparel",
    "construction", "cram_school", "dentist", "care_service",
]


def convert_western_to_wareki(text: str) -> str:
    """
    テキスト中の西暦年月表記（例: 2010年4月）を和暦（元号）表記に変換する。
    既に和暦表記になっている箇所は変換しない。

    元号の境界:
    - 昭和: 1926〜1989/1/7
    - 平成: 1989/1/8〜2019/4/30
    - 令和: 2019/5/1〜
    """

    def year_month_to_wareki(year: int, month: int | None) -> str:
        if year > 2019 or (year == 2019 and month is not None and month >= 5):
            wareki_year = year - 2018
            era = "令和"
        elif year == 2019:
            # 2019年1〜4月 → 平成31年
            wareki_year = 31
            era = "平成"
        elif year > 1989 or (year == 1989 and month is not None and month >= 2):
            wareki_year = year - 1988
            era = "平成"
        elif year == 1989:
            # 1989年1月 → 昭和64年
            wareki_year = 64
            era = "昭和"
        elif year >= 1926:
            wareki_year = year - 1925
            era = "昭和"
        else:
            return f"{year}年"  # 大正以前はそのまま

        year_str = "元" if wareki_year == 1 else str(wareki_year)
        return f"{era}{year_str}年"

    def replace_match(m: re.Match) -> str:
        year = int(m.group(1))
        month_str = m.group(2)  # "4月" or None
        month = int(m.group(3)) if m.group(3) else None
        wareki = year_month_to_wareki(year, month)
        return wareki + (month_str or "")

    # 西暦4桁（1900〜2099）の年月表記を対象にする。直前が数字の場合は除外
    pattern = r'(?<!\d)((?:19|20)\d{2})年((\d{1,2})月)?'
    return re.sub(pattern, replace_match, text)


async def build_examples_text(db: AsyncSession) -> str:
    """
    example_contentsテーブルから全業種の記入例を取得し、
    システムプロンプト用のテキストに整形して返します。
    """
    result = await db.execute(select(ExampleContent))
    examples = result.scalars().all()

    # 業種・セクションでインデックス化
    examples_dict: dict[str, dict[str, str]] = {}
    for ex in examples:
        examples_dict.setdefault(ex.industry_type, {})[ex.section_key] = ex.example_text

    parts = []
    for industry in INDUSTRY_ORDER:
        if industry not in examples_dict:
            continue
        display_name = INDUSTRY_DISPLAY_NAMES.get(industry, industry)
        parts.append(f"■{display_name}")
        parts.append("")
        for section_key in SECTION_ORDER:
            if section_key not in examples_dict[industry]:
                continue
            section_name = SECTION_DISPLAY_NAMES.get(section_key, section_key)
            parts.append(section_name)
            parts.append(examples_dict[industry][section_key])
            parts.append("")

    return "\n".join(parts)


async def get_section_example(
    db: AsyncSession, industry_type: str, section_key: str
) -> str | None:
    """
    指定した業種・セクションの記入例をDBから取得して返します。
    見つからない場合はNoneを返します。
    """
    if not industry_type or not section_key:
        return None
    result = await db.execute(
        select(ExampleContent).where(
            ExampleContent.industry_type == industry_type,
            ExampleContent.section_key == section_key,
        )
    )
    example = result.scalar_one_or_none()
    return example.example_text if example else None


@app.get("/", response_class=HTMLResponse)
async def read_root(
    request: Request,
    session_id: str | None = Cookie(default=None),
    db: AsyncSession = Depends(get_db),
):
    """
    アプリケーションのメインページ（ダッシュボード）を表示します。

    CookieからセッションIDを取得し、セッションが存在する場合は
    これまでのタスク進捗状況とチャット履歴を復元して表示します。

    Args:
        request (Request): FastAPIのリクエストオブジェクト
        session_id (str | None): クライアントのCookieから取得したセッションID
        db (AsyncSession): データベースセッション

    Returns:
        TemplateResponse: レンダリングされた `index.html`
    """
    # セッションIDがあればデータをロード
    user_tasks = TASKS  # デフォルト
    chat_messages_html = ""  # ここで履歴HTML文字列を作る

    if session_id:
        data = await session_store.load_session(db, session_id)
        if data:
            # データベースから取得したタスク状態をTASKSテンプレートにマージ
            task_states = data.get("task_states", {})
            user_tasks = []
            for task in TASKS:
                task_copy = task.copy()
                # データベースに保存されている状態があればそれを使用
                if task["id"] in task_states:
                    task_copy["status"] = task_states[task["id"]]
                user_tasks.append(task_copy)

            history = data.get("chat_history", [])

            # Gemini履歴を復元
            if history:
                gemini_service.restore_chat_session(session_id, history)

            # 表示用の履歴HTMLを再構築
            # ここでは簡易的に、historyの構造から role=user or model を判定してテンプレート適用
            # 本来は templates/components/message.html を再利用して文字列結合する
            for msg in history:
                role = msg["role"]
                # model -> "text", user -> "text"
                text_content = ""
                if "parts" in msg:
                    for part in msg["parts"]:
                        if "text" in part:
                            text_content += part["text"]

                is_user = role == "user"

                msg_html = templates.get_template("components/message.html").render(
                    {"request": request, "message": text_content, "is_user": is_user}
                )
                chat_messages_html += msg_html

    return templates.TemplateResponse(
        "index.html",
        {
            "request": request,
            "steps": ROADMAP_STEPS,
            "tasks": user_tasks,
            "initial_chat_history": chat_messages_html,
        },
    )


@app.get("/chat/start", response_class=HTMLResponse)
async def start_chat(
    request: Request,
    task_id: str | None = None,
    db: AsyncSession = Depends(get_db),
):
    """
    新しいチャットセッションを開始します。

    新規にUUIDを生成してセッションIDとし、AIエージェントからの初期挨拶メッセージを取得して返します。
    特定のタスクIDが指定された場合、そのタスクのヒアリングから開始するようにAIに指示します。

    Args:
        request (Request): FastAPIのリクエストオブジェクト
        task_id (str | None): 個別のタスクから開始する場合のタスクID（例: "motivation"）
        db (AsyncSession): データベースセッション

    Returns:
        TemplateResponse: 初期メッセージを含む `components/chat_interface.html`
    """
    session_id = str(uuid.uuid4())

    # データベースにセッションレコードを作成
    new_session = SessionModel(id=session_id)
    db.add(new_session)
    await db.commit()
    print(f"[DEBUG] Created new session: {session_id}")

    initial_task_title = None
    if task_id:
        # TASKSからタイトルを検索
        for task in TASKS:
            if task["id"] == task_id:
                initial_task_title = task["title"]
                break

    # DBから記入例テキストを構築してシステムプロンプトに渡す
    examples_text = await build_examples_text(db)

    # エージェントからの初期挨拶を生成
    # GeminiService.start_chat_session requires session_id
    tGreeting = await gemini_service.start_chat_session(session_id, initial_task_title, examples_text)

    # 業種選択マーカーの検出
    industry_selector_html = ""
    if "[[INDUSTRY_SELECTOR]]" in tGreeting:
        tGreeting = tGreeting.replace("[[INDUSTRY_SELECTOR]]", "")
        # 業種選択ボタンのHTMLを生成
        industry_selector_html = templates.get_template(
            "components/industry_selector.html"
        ).render({"request": request})

    # 完全なチャットインターフェースを返す（chat_interface.html テンプレートを使用）
    response = templates.TemplateResponse(
        "components/chat_interface.html",
        {
            "request": request,
            "message": tGreeting,
            "is_user": False,
            "additional_content": industry_selector_html,  # 業種選択ボタンを追加コンテンツとして渡す
        },
    )
    # 簡易的にCookieでセッションIDを管理
    response.set_cookie(key="session_id", value=session_id)
    return response


@app.post("/chat/approve_draft", response_class=HTMLResponse)
async def approve_draft(
    request: Request,
    task_id: str = Form(...),
    session_id: str | None = Cookie(default=None),
    db: AsyncSession = Depends(get_db),
):
    """
    ドラフト承認ボタン（OK）がクリックされた際の処理。

    Geminiに依存せず、確実に次のステップに進みます。

    処理フロー:
    1. 直前のAI応答からドラフト内容を抽出
    2. タスクを完了状態に更新
    3. データベースに保存
    4. 次のタスクの質問を生成
    """
    if not session_id:
        return HTMLResponse("Session not found", status_code=400)

    print(f"[DEBUG] /chat/approve_draft - task_id: {task_id}, session_id: {session_id}")

    # セッションデータをロード
    import copy

    current_tasks = copy.deepcopy(TASKS)
    session_data = await session_store.load_session(db, session_id)

    if session_data and "task_states" in session_data:
        task_states = session_data["task_states"]
        current_tasks = []
        for task in TASKS:
            task_copy = task.copy()
            if task["id"] in task_states:
                task_copy["status"] = task_states[task["id"]]
            current_tasks.append(task_copy)

    # チャット履歴から直前のAI応答を取得
    history_data = gemini_service.get_chat_history(session_id)
    last_ai_response = None
    if history_data:
        # 最後のAIメッセージを探す
        for msg in reversed(history_data):
            if msg.get("role") == "model":
                last_ai_response = msg.get("parts", [{}])[0].get("text", "")
                break

    print(
        f"[DEBUG] Last AI response preview: {last_ai_response[:200] if last_ai_response else 'None'}"
    )

    # ドラフト内容を抽出
    draft_content = None
    if last_ai_response:
        draft_content = extract_single_section_from_response(last_ai_response, task_id)

    if not draft_content:
        print(f"[WARNING] Could not extract draft content for task {task_id}")
        # フォールバック: 直前のAI応答全体を使用
        if last_ai_response:
            # マーカーを除去して使用
            draft_content = (
                last_ai_response.replace("[[CONTENT_START]]", "")
                .replace("[[CONTENT_END]]", "")
                .replace("[[DRAFT_PROPOSED]]", "")
                .strip()
            )

    # background（経営者の略歴等）は西暦→和暦に変換
    if draft_content and task_id == "background":
        converted = convert_western_to_wareki(draft_content)
        if converted != draft_content:
            print("[DEBUG] 和暦変換を適用しました（background）")
            draft_content = converted

    # タスクを完了状態に更新
    task_found = False
    for task in current_tasks:
        if task["id"] == task_id:
            task["status"] = "done"
            task_found = True
            print(f"[DEBUG] Task {task_id} marked as done")
            break

    if not task_found:
        print(f"[ERROR] Task {task_id} not found in task list")
        return HTMLResponse("Invalid task ID", status_code=400)

    # ドラフト内容をデータベースに保存
    if draft_content:
        section_update = {task_id: draft_content}
        await session_store.save_session(
            db,
            session_id,
            current_tasks,
            history_data,
            sections=section_update,
        )
        print(f"[DEBUG] Saved draft content for task {task_id}")

    # 次のタスクを見つける
    next_task = None
    for i, task in enumerate(TASKS):
        if task["id"] == task_id:
            if i + 1 < len(TASKS):
                next_task = TASKS[i + 1]
            break

    # 応答メッセージを生成
    user_msg_html = templates.get_template("components/message.html").render(
        {"request": request, "message": "この内容でOKです", "is_user": True}
    )

    # 次のセクションの記入例をDBから取得してAIメッセージに添付
    industry_type = session_data.get("industry_type") if session_data else None
    ai_ok_message = "この内容でOKです。"
    if next_task and industry_type:
        next_example = await get_section_example(db, industry_type, next_task["id"])
        if next_example:
            section_name = SECTION_DISPLAY_NAMES.get(next_task["id"], next_task["title"])
            ai_ok_message += f"\n\n【{section_name} の記入例（参考）】\n{next_example}"
            print(f"[DEBUG] 記入例を添付: {next_task['id']} ({industry_type})")

    # AIに「この内容でOKです」を伝えて、次のタスクの質問を生成させる
    ai_next_response = await gemini_service.generate_response(
        session_id, ai_ok_message
    )

    # マーカーとコンテンツの削除処理
    import re

    # [[CONTENT_START]] ~ [[CONTENT_END]] の部分を削除（承認済みコンテンツは再表示不要）
    ai_next_response = re.sub(
        r"\[\[CONTENT_START\]\].*?\[\[CONTENT_END\]\]",
        "",
        ai_next_response,
        flags=re.DOTALL,
    )

    # [[COMPLETED:xxx]] マーカーを削除
    ai_next_response = re.sub(r"\[\[COMPLETED:[a-z_]+\]\]", "", ai_next_response)

    # 過剰な空白・改行を整理
    ai_next_response = re.sub(r"\n{3,}", "\n\n", ai_next_response)
    ai_next_response = ai_next_response.strip()

    # フォールバック: Geminiが適切な応答を返さない場合
    if not ai_next_response or len(ai_next_response.strip()) < 20:
        if next_task:
            ai_next_response = f"承認ありがとうございます。\n\nそれでは次に、「{next_task['title']}」についてお伺いします。\n\n{next_task['desc']}\n\nどのような内容でしょうか？"
        else:
            ai_next_response = "承認ありがとうございます。\n\nすべての項目が完了しました！左側の「計画書を確認・編集」から内容を確認し、必要に応じて編集してください。"

    ai_msg_html = templates.get_template("components/message.html").render(
        {"request": request, "message": ai_next_response, "is_user": False}
    )

    # チェックボックス更新用のOOB HTML
    task_update_html = f"""
    <input id="task-{task_id}"
           name="task-{task_id}"
           type="checkbox"
           class="h-4 w-4 rounded border-gray-300 text-blue-600 focus:ring-blue-600"
           checked
           hx-swap-oob="true">
    """

    response = HTMLResponse(content=user_msg_html + ai_msg_html + task_update_html)
    response.set_cookie(key="session_id", value=session_id)

    return response


@app.post("/chat/select_industry", response_class=HTMLResponse)
async def select_industry(
    request: Request,
    industry: str = Form(...),
    session_id: str | None = Cookie(default=None),
    db: AsyncSession = Depends(get_db),
):
    """
    業種選択ボタンがクリックされた際の処理。

    業種を保存し、AIの応答を含む完全なチャットインターフェースを返します。
    """
    if not session_id:
        return HTMLResponse("Session not found", status_code=400)

    # 業種タイプのマッピング（番号 → 内部キー）
    industry_mapping = {
        "1": "restaurant",
        "2": "beauty",
        "3": "car_sales",
        "4": "apparel",
        "5": "software",
        "6": "construction",
        "7": "cram_school",
        "8": "dentist",
        "9": "care_service",
    }

    # AIに伝える業種名ラベル（番号 → 日本語表示名）
    industry_label_mapping = {
        "1": "洋風居酒屋（飲食店）",
        "2": "美容業",
        "3": "中古自動車販売業",
        "4": "婦人服・子供服小売業",
        "5": "ソフトウェア開発業（ITサービス、Webサービス、マッチングサービス、アプリ開発等）",
        "6": "内装工事業",
        "7": "学習塾",
        "8": "歯科診療所",
        "9": "介護サービス",
    }

    industry_type = industry_mapping.get(industry)
    if not industry_type:
        return HTMLResponse("Invalid industry selection", status_code=400)

    # データベースに業種タイプを保存
    await session_store.update_industry_type(db, session_id, industry_type)
    print(f"[DEBUG] Industry type set to: {industry_type}")

    # AIに業種名を明示して伝える（数字だけだと誤判定する恐れがあるため）
    industry_label = industry_label_mapping.get(industry, industry)
    ai_message = f"「{industry_label}」を選択しました。"

    # 最初のセクション（創業の動機）の記入例をDBから取得してメッセージに添付
    first_example = await get_section_example(db, industry_type, "motivation")
    if first_example:
        section_name = SECTION_DISPLAY_NAMES.get("motivation", "1. 創業の動機")
        ai_message += f"\n\n【{section_name} の記入例（{industry_label}の場合）】\n{first_example}"

    ai_response_text = await gemini_service.generate_response(session_id, ai_message)

    # チャット履歴を保存
    history_data = gemini_service.get_chat_history(session_id)
    await session_store.save_session(db, session_id, [], history_data)

    # 完全なチャットインターフェースを返す
    return templates.TemplateResponse(
        "components/chat_interface.html",
        {
            "request": request,
            "message": ai_response_text,
            "is_user": False,
            "show_industry_selection": f"{industry}",  # 選択された業種番号をユーザーメッセージとして表示
        },
    )


@app.post("/chat/message", response_class=HTMLResponse)
async def chat_message(
    request: Request,
    user_message: str = Form(...),
    session_id: str | None = Cookie(default=None),
    db: AsyncSession = Depends(get_db),
):
    """
    ユーザーからのチャットメッセージを処理し、AIからの応答を返します。

    主な処理フロー:
    1. ユーザーメッセージをチャット履歴に追加
    2. Gemini Service を用いてAI応答を生成
    3. AI応答内の特殊コマンド（マーカー）を検出し、以下の処理を実行
       - `[[DRAFT_PROPOSED]]`: ドラフト提案ボタン（OK/修正）を表示
       - `[[COMPLETED:{task_id}]]`: 指定されたタスクを完了状態(done)に更新し、チェックボックスをオンにする
    4. 更新されたセッション状態（タスク、履歴）を保存

    Args:
        request (Request): FastAPIのリクエストオブジェクト
        user_message (str): フォームから送信されたユーザーのメッセージ
        session_id (str | None): CookieセッションID
        db (AsyncSession): データベースセッション

    Returns:
        HTMLResponse: ユーザーメッセージ、AIメッセージ、およびもしあればUI更新用HTML（OOB Swap）を含むフラグメント
    """
    # ユーザーのメッセージを表示するためのHTML
    user_msg_html = templates.get_template("components/message.html").render(
        {"request": request, "message": user_message, "is_user": True}
    )

    # セッションIDがない場合のフォールバック（途中でCookieが消えた場合など）
    # 本来はエラーにするか再接続を促すべきだが、ここでは新規作成して続行
    if not session_id:
        session_id = str(uuid.uuid4())

    # AI応答の生成
    ai_response_text = await gemini_service.generate_response(session_id, user_message)

    print(f"[DEBUG] AI Response preview: {ai_response_text[:200]}")

    # 業種選択の検出と処理
    industry_mapping = {
        "1": "restaurant",
        "2": "beauty",
        "3": "car_sales",
        "4": "apparel",
        "5": "software",
        "6": "construction",
        "7": "cram_school",
        "8": "dentist",
        "9": "care_service",
    }

    # ユーザーメッセージから業種番号を検出
    user_message_stripped = user_message.strip()
    if user_message_stripped in industry_mapping:
        industry_type = industry_mapping[user_message_stripped]
        await session_store.update_industry_type(db, session_id, industry_type)
        print(f"[DEBUG] Industry type set to: {industry_type}")

    # タスク完了マーカーの検出と処理
    task_update_html = ""
    draft_buttons_html = ""
    industry_selector_html = ""

    # 業種選択マーカーの検出
    if "[[INDUSTRY_SELECTOR]]" in ai_response_text:
        ai_response_text = ai_response_text.replace("[[INDUSTRY_SELECTOR]]", "")
        # 業種選択ボタンのHTMLを生成
        industry_selector_html = templates.get_template(
            "components/industry_selector.html"
        ).render({"request": request})

    # セッションからタスク状態をロード（deep copy）
    import copy

    current_tasks = copy.deepcopy(TASKS)  # デフォルト値
    session_data = await session_store.load_session(db, session_id)
    if session_data and "task_states" in session_data:
        # データベースから取得したタスク状態をTASKSテンプレートにマージ
        task_states = session_data["task_states"]
        current_tasks = []
        for task in TASKS:
            task_copy = task.copy()
            # データベースに保存されている状態があればそれを使用
            if task["id"] in task_states:
                task_copy["status"] = task_states[task["id"]]
            current_tasks.append(task_copy)

    # ドラフト提示マーカーの検出（フォールバック検出を含む）
    # [[DRAFT_PROPOSED]]がなくても、ドラフト的な内容を含む場合はOKボタンを表示する
    DRAFT_PHRASES = ["でよろしいでしょうか", "いかがでしょうか", "ご確認ください", "以下の内容で"]
    has_draft_proposed = "[[DRAFT_PROPOSED]]" in ai_response_text
    if not has_draft_proposed:
        if "[[CONTENT_START]]" in ai_response_text:
            has_draft_proposed = True
            print("[DEBUG] Fallback: [[CONTENT_START]] detected as draft proposal")
        elif any(phrase in ai_response_text for phrase in DRAFT_PHRASES):
            has_draft_proposed = True
            print("[DEBUG] Fallback: draft confirmation phrase detected")

    if has_draft_proposed:
        ai_response_text = ai_response_text.replace("[[DRAFT_PROPOSED]]", "")

        # 現在進行中のタスク（最初のpendingタスク）を特定
        current_task_id = None
        for task in current_tasks:
            if task["status"] == "pending":
                current_task_id = task["id"]
                break

        if current_task_id:
            draft_buttons_html = f"""
            <div class="flex gap-4 mt-2 mb-4 ml-12">
                <button hx-post="/chat/approve_draft"
                        hx-vals='{{"task_id": "{current_task_id}"}}'
                        hx-target="#chat-history"
                        hx-swap="beforeend"
                        class="px-4 py-2 bg-blue-600 text-white rounded-lg text-sm hover:bg-blue-700 font-bold transition-colors">
                    OK(次の項目へ)
                </button>
                <button hx-post="/chat/message"
                        hx-vals='{{"user_message": "文面を修正したいです"}}'
                        hx-target="#chat-history"
                        hx-swap="beforeend"
                        class="px-4 py-2 bg-white border border-gray-300 text-gray-700 rounded-lg text-sm hover:bg-gray-50 transition-colors">
                    文面を修正する
                </button>
            </div>
            """
            print(f"[DEBUG] Draft buttons created for task: {current_task_id}")
        else:
            print("[WARNING] No pending task found for draft buttons")

    print(
        f"[DEBUG] Loaded tasks: {[t['id'] + ':' + t['status'] for t in current_tasks]}"
    )

    # コンテンツマーカーを削除（表示用。ドラフト内容はOKボタン経由で保存するため表示不要）
    ai_response_text = ai_response_text.replace("[[CONTENT_START]]", "")
    ai_response_text = ai_response_text.replace("[[CONTENT_END]]", "")

    # [[COMPLETED:xxx]] マーカーをAIが誤出力した場合のクリーンアップ（保存処理は行わない）
    ai_response_text = re.sub(r"\n*\[\[COMPLETED:[a-z_]+\]\]\n*", "\n\n", ai_response_text)

    # 過剰な改行を整理
    ai_response_text = re.sub(r"\n{3,}", "\n\n", ai_response_text)

    ai_msg_html = templates.get_template("components/message.html").render(
        {"request": request, "message": ai_response_text, "is_user": False}
    )

    response = HTMLResponse(
        content=user_msg_html
        + ai_msg_html
        + industry_selector_html
        + draft_buttons_html
        + task_update_html
    )

    # セッションIDを再設定（有効期限延長などの効果もあるが、とりあえず念の為）
    response.set_cookie(key="session_id", value=session_id)

    # --- セッション状態の保存 ---
    # タスク完了時に既にセクションとタスク状態を保存している場合があるため、
    # ここでは主にチャット履歴の最終更新を行う
    # 1. GeminiServiceから現在のチャット履歴を取得
    history_data = gemini_service.get_chat_history(session_id)
    # 2. タスク状態とチャット履歴を保存（セクションは既に保存済みの場合はスキップ）
    await session_store.save_session(db, session_id, current_tasks, history_data)

    return response


@app.get(
    "/plan/edit",
    response_class=HTMLResponse,
    summary="計画書編集画面の取得",
    description="""
    創業計画書の編集画面（エディタ）を取得します。

    ⚠️ **Swagger UIでの直接テストはできません**
    - このエンドポイントはCookieによるセッション管理が必要です
    - Swagger UIではCookieパラメータを正しく送信できません

    **テスト方法:**
    - 下記の「Try it out」→「session_id」を入力→「Execute」で表示されるCurlコマンドをコピー
    - ターミナルで実行
    """,
)
async def edit_plan(
    request: Request,
    session_id: str | None = Cookie(default=None),
    db: AsyncSession = Depends(get_db),
):
    """
    創業計画書の編集画面（エディタ）を取得します。

    セッションに保存されている現在の計画書セクションを読み込み、
    全文として結合してから編集用フォームにセットして返します。

    Args:
        request (Request): FastAPIのリクエストオブジェクト
        session_id (str | None): CookieセッションID
        db (AsyncSession): データベースセッション

    Returns:
        TemplateResponse: `components/plan_editor.html`
    """
    if not session_id:
        return HTMLResponse("Session not found", status_code=400)

    data = await session_store.load_session(db, session_id)
    print(f"[DEBUG] /plan/edit - session_id: {session_id}")
    print(f"[DEBUG] /plan/edit - data loaded: {data is not None}")

    sections = {}
    if data and data.get("sections"):
        sections = data["sections"]
        print(f"[DEBUG] /plan/edit - sections keys: {list(sections.keys())}")
        print(
            f"[DEBUG] /plan/edit - sections with content: {[k for k, v in sections.items() if v]}"
        )
    else:
        # 空のセクション辞書を用意（テンプレートでエラーを防ぐ）
        sections = {
            "motivation": None,
            "background": None,
            "service": None,
            "employees": None,
            "partners": None,
            "related_companies": None,
            "loans": None,
            "funds": None,
            "outlook": None,
            "free_description": None,
        }
        print("[DEBUG] /plan/edit - No sections found, using empty template")

    return templates.TemplateResponse(
        "components/plan_editor.html",
        {"request": request, "sections": sections},
    )


@app.post(
    "/plan/save",
    response_class=HTMLResponse,
    summary="計画書の保存",
    description="""
    編集された創業計画書テキストを保存します。

    ⚠️ **Swagger UIでの直接テストはできません**
    - このエンドポイントはCookieによるセッション管理が必要です
    - Swagger UIではCookieパラメータを正しく送信できません

    **テスト方法:**
    - 下記の「Try it out」→各フィールドと「session_id」を入力→「Execute」で表示されるCurlコマンドをコピー
    - ターミナルで実行
    """,
)
async def save_plan(
    request: Request,
    motivation: str = Form(default=""),
    background: str = Form(default=""),
    service: str = Form(default=""),
    employees: str = Form(default=""),
    partners: str = Form(default=""),
    related_companies: str = Form(default=""),
    loans: str = Form(default=""),
    funds: str = Form(default=""),
    outlook: str = Form(default=""),
    free_description: str = Form(default=""),
    session_id: str | None = Cookie(default=None),
    db: AsyncSession = Depends(get_db),
):
    """
    編集された創業計画書テキストを保存します。

    各セクションのフォームデータを受け取り、データベースに保存し、
    保存後は閲覧モード（Viewer）のHTMLを返します。

    Args:
        request (Request): FastAPIのリクエストオブジェクト
        motivation - free_description (str): 各セクションのフォームデータ
        session_id (str | None): CookieセッションID
        db (AsyncSession): データベースセッション

    Returns:
        TemplateResponse: `components/plan_viewer.html`
    """
    if not session_id:
        return HTMLResponse("Session not found", status_code=400)

    # 現在のセッションデータをロード
    data = await session_store.load_session(db, session_id)
    if not data:
        data = {"task_states": {}, "chat_history": []}

    # デバッグ: 受信したフォームデータを確認
    print(f"[DEBUG] /plan/save - Received background data:")
    print(f"[DEBUG]   Length: {len(background)} chars")
    print(f"[DEBUG]   Raw: {repr(background[:200])}")
    print(f"[DEBUG]   Contains newlines: {'\\n' in background or '\\r' in background}")

    # フォームデータからセクション辞書を構築
    # strip()は前後の空白のみ削除し、内部の改行は保持する
    sections = {
        "motivation": motivation.strip() if motivation.strip() else None,
        "background": background.strip() if background.strip() else None,
        "service": service.strip() if service.strip() else None,
        "employees": employees.strip() if employees.strip() else None,
        "partners": partners.strip() if partners.strip() else None,
        "related_companies": related_companies.strip()
        if related_companies.strip()
        else None,
        "loans": loans.strip() if loans.strip() else None,
        "funds": funds.strip() if funds.strip() else None,
        "outlook": outlook.strip() if outlook.strip() else None,
        "free_description": free_description.strip()
        if free_description.strip()
        else None,
    }

    print(
        f"[DEBUG] /plan/save - Saving {len([v for v in sections.values() if v])} sections with content"
    )
    if sections.get("background"):
        print(
            f"[DEBUG] /plan/save - background after strip: {repr(sections['background'][:200])}"
        )

    # タスク状態をTASKS形式に変換
    task_states = data.get("task_states", {})
    current_tasks = []
    for task in TASKS:
        task_copy = task.copy()
        if task["id"] in task_states:
            task_copy["status"] = task_states[task["id"]]
        current_tasks.append(task_copy)

    # 保存
    history = data.get("chat_history", [])
    await session_store.save_session(
        db, session_id, current_tasks, history, sections=sections
    )

    # 表示用に全文テキストを再構築
    plan_text = build_plan_text_from_sections(sections)

    # 閲覧モードのHTMLを返す
    return templates.TemplateResponse(
        "components/plan_viewer.html",
        {"request": request, "plan_text": plan_text, "stepper_oob": ""},
    )


@app.post(
    "/plan/generate",
    response_class=HTMLResponse,
    summary="計画書のドラフト生成",
    description="""
    これまでのチャット内容を元に、創業計画書のドラフトを生成します。

    ⚠️ **Swagger UIでの直接テストはできません**
    - このエンドポイントはCookieによるセッション管理が必要です
    - Swagger UIではCookieパラメータを正しく送信できません

    **テスト方法:**
    - 下記の「Try it out」→「session_id」を入力→「Execute」で表示されるCurlコマンドをコピー
    - ターミナルで実行
    """,
)
async def generate_plan(
    request: Request,
    session_id: str | None = Cookie(default=None),
    db: AsyncSession = Depends(get_db),
):
    """
    これまでのチャット内容を元に、創業計画書のドラフトを生成します。

    Gemini APIを使用して、会話履歴全体から創業計画書形式のテキストを生成し、
    各セクションに分割してデータベースに保存します。
    また、ステッパー（進捗表示）を「創業計画書作成」フェーズに進めます。

    Args:
        request (Request): FastAPIのリクエストオブジェクト
        session_id (str | None): CookieセッションID
        db (AsyncSession): データベースセッション

    Returns:
        TemplateResponse: 生成された計画書を表示する `components/plan_viewer.html`
    """
    if not session_id:
        # セッションがない場合はエラー表示
        return HTMLResponse(
            '<div class="flex items-center justify-center h-full text-red-500">セッションが見つかりません。最初からやり直してください。</div>',
            status_code=400,
        )

    # データベースからセッションデータをロードし、チャット履歴を復元
    current_data = await session_store.load_session(db, session_id)
    if not current_data:
        return HTMLResponse(
            '<div class="flex items-center justify-center h-full text-red-500">セッションデータが見つかりません。最初からやり直してください。</div>',
            status_code=400,
        )

    # チャット履歴をGeminiServiceに復元
    history = current_data.get("chat_history", [])
    if history:
        gemini_service.restore_chat_session(session_id, history)

    # Geminiでドラフト生成
    plan_text = await gemini_service.generate_business_plan(session_id)

    # マーカーを除去（表示用に不要な制御コマンドを削除）
    plan_text = re.sub(r"\[\[DRAFT_PROPOSED\]\]", "", plan_text)
    plan_text = re.sub(r"\[\[COMPLETED:[a-z_]+\]\]", "", plan_text)

    # --- 既存のセクションデータを取得 ---
    # 注: current_data は既に上でロード済み
    existing_sections = {}
    if current_data and current_data.get("sections"):
        existing_sections = current_data["sections"]
        print(
            f"[DEBUG] /plan/generate - Found existing sections: {list(existing_sections.keys())}"
        )
        print(
            f"[DEBUG] /plan/generate - Sections with data: {[k for k, v in existing_sections.items() if v]}"
        )

    # 全文テキストをセクション別に分割
    new_sections = extract_sections_from_text(plan_text)

    # 既存のセクションデータと新規生成データをマージ（既存データを優先）
    sections = {}
    for key in [
        "motivation",
        "background",
        "service",
        "employees",
        "partners",
        "related_companies",
        "loans",
        "funds",
        "outlook",
        "free_description",
    ]:
        # 既存データがあればそれを使用、なければ新規生成データを使用
        if existing_sections.get(key):
            sections[key] = existing_sections[key]
            print(f"[DEBUG] /plan/generate - Using existing data for {key}")
        elif new_sections.get(key):
            sections[key] = new_sections[key]
            print(f"[DEBUG] /plan/generate - Using new generated data for {key}")
        else:
            sections[key] = None

    # --- 生成されたプランの保存 ---
    if current_data:
        # タスク状態をTASKS形式に変換
        task_states = current_data.get("task_states", {})
        saved_tasks = []
        for task in TASKS:
            task_copy = task.copy()
            if task["id"] in task_states:
                task_copy["status"] = task_states[task["id"]]
            saved_tasks.append(task_copy)

        saved_history = current_data.get("chat_history", [])
        await session_store.save_session(
            db, session_id, saved_tasks, saved_history, sections=sections
        )

    # ステップの状態を更新 (Conceptual: 1->Completed, 2->Current)
    current_steps = [s.copy() for s in ROADMAP_STEPS]
    current_steps[0]["status"] = "completed"
    current_steps[1]["status"] = "current"

    # ステッパーHTMLをOOB更新用にレンダリング
    stepper_html = templates.get_template("components/stepper.html").render(
        {"steps": current_steps}
    )
    # stepper.html内のnav要素をhx-swap-oob="true"にするために置換 (またはtemplate側で対応)
    # ここでは簡易的に文字列置換で hx-swap-oob 属性を注入
    stepper_html = stepper_html.replace(
        '<nav aria-label="Progress" id="stepper-nav">',
        '<nav aria-label="Progress" id="stepper-nav" hx-swap-oob="true">',
    )

    return templates.TemplateResponse(
        "components/plan_viewer.html",
        {"request": request, "plan_text": plan_text, "stepper_oob": stepper_html},
    )


@app.get(
    "/plan/download_excel",
    summary="計画書のExcelダウンロード",
    description="""
    作成完了した創業計画書をExcel形式でZIP圧縮してダウンロードします。

    ⚠️ **Swagger UIでの直接テストはできません**
    - このエンドポイントはCookieによるセッション管理が必要です
    - Swagger UIではCookieパラメータを正しく送信できません

    **テスト方法1: curlでファイル保存（推奨）**
    ```bash
    curl -X 'GET' \\
      'http://localhost:8000/plan/download_excel' \\
      -H 'Cookie: session_id=あなたのセッションID' \\
      -o 創業計画書.zip
    ```

    **テスト方法2: ブラウザで直接アクセス（最も簡単）**
    - ブラウザで `http://localhost:8000/plan/download_excel` を開く
    - 自動的にダウンロードが開始されます
    """,
)
async def download_plan_excel(
    session_id: str | None = Cookie(default=None),
    db: AsyncSession = Depends(get_db),
):
    """
    作成完了した創業計画書をExcel形式でZIP圧縮してダウンロードします。

    1. セッションから計画書の各セクションデータを取得
    2. `templates/startup_plan_template.xlsx` を読み込み、対応するセルに転記
    3. 計画書の内容から業種を推測し、適切な記入例PDF（`static/templates/examples/*.pdf`）を選択
    4. ExcelファイルとPDFファイルをZIPアーカイブにまとめてストリーミング返却

    Args:
        session_id (str | None): CookieセッションID
        db (AsyncSession): データベースセッション

    Returns:
        StreamingResponse: ZIPファイル（application/zip）
    """
    if not session_id:
        return Response("Session not found", status_code=400)

    data = await session_store.load_session(db, session_id)
    if not data or not data.get("sections"):
        return Response("Plan data not found", status_code=404)

    sections = data.get("sections")

    # テンプレート読み込み
    template_path = BASE_DIR / "static" / "templates" / "startup_plan_template.xlsx"
    workbook = openpyxl.load_workbook(template_path)
    sheet = workbook.active

    # テンプレート内の見出しセルを探して、転記先セルを動的に推定する
    label_to_key = {
        "創業の動機": "motivation",
        "経営者の略歴等": "background",
        "取扱商品・サービス": "service",
        "従業員": "employees",
        "取引先・取引関係等": "partners",
        "関連企業": "related_companies",
        "お借入の状況": "loans",
        "必要な資金と調達方法": "funds",
        "事業の見通し": "outlook",
        "自由記述欄": "free_description",
    }

    label_cells = {}
    for row in sheet.iter_rows():
        for cell in row:
            if isinstance(cell.value, str):
                value = cell.value.strip()
                for label, key in label_to_key.items():
                    if label in value:
                        label_cells[key] = cell

    def choose_target_cell(label_cell):
        label_row = label_cell.row
        label_col = label_cell.column
        candidates = []
        for merged_range in sheet.merged_cells.ranges:
            min_col, min_row, max_col, max_row = range_boundaries(str(merged_range))
            if min_row >= label_row + 1 and min_row <= label_row + 12:
                width = max_col - min_col + 1
                height = max_row - min_row + 1
                if width >= 6:
                    if min_col <= label_col <= max_col or min_col > label_col:
                        candidates.append(
                            (min_row, -(width * height), min_col, str(merged_range))
                        )
        if candidates:
            candidates.sort()
            selected = candidates[0][3]
            min_col, min_row, _, _ = range_boundaries(selected)
            return f"{get_column_letter(min_col)}{min_row}"
        return f"{get_column_letter(label_col + 1)}{label_row}"

    mapping = {key: None for key in label_to_key.values()}
    for key, cell in label_cells.items():
        mapping[key] = choose_target_cell(cell)

    print("[DEBUG] Cell mapping detected:")
    for key, addr in mapping.items():
        print(f"  {key}: {addr}")

    # フォールバック用の静的マッピング（テンプレート検出できない場合）
    fallback_mapping = {
        "motivation": "A9",
        "background": "A16",
        "service": "A27",
        "employees": "T21",
        "partners": "M27",
        "related_companies": "A36",
        "loans": "M36",
        "funds": "A45",
        "outlook": "M45",
        "free_description": "A54",
    }
    for key, addr in fallback_mapping.items():
        if not mapping.get(key):
            mapping[key] = addr

    # 特定のセルを手動で上書き（動的検出が正しくない場合）
    mapping["background"] = "H13"  # 経営者の略歴等

    # データベースから取得したセクションデータをExcelに転記
    print("[DEBUG] Starting Excel data export...")
    print(f"[DEBUG] Sections to export: {list(sections.keys())}")
    print(f"[DEBUG] Sections with content: {[k for k, v in sections.items() if v]}")

    for key, content in sections.items():
        if content:
            # 「創業の動機」は特殊処理（60文字ごとに分割してB7～B12に転記）
            if key == "motivation":
                print(
                    "[DEBUG] Processing motivation (創業の動機) with 60-char splitting..."
                )
                # 改行を削除して1つの文字列にする
                text = content.replace("\n", "")
                print(f"[DEBUG] Total length: {len(text)} chars")

                # 60文字ごとに分割（最大6行）
                max_rows = 6
                char_limit = 60

                for idx in range(max_rows):
                    start_pos = idx * char_limit
                    end_pos = start_pos + char_limit

                    if start_pos >= len(text):
                        break

                    line_text = text[start_pos:end_pos]
                    row_num = 7 + idx  # B7, B8, B9, B10, B11, B12

                    sheet[f"B{row_num}"].value = line_text
                    print(
                        f"[DEBUG]   ✓ Wrote to B{row_num}: '{line_text[:40]}...' ({len(line_text)} chars)"
                    )

                print(
                    f"[DEBUG] ✓ Successfully processed motivation in {min((len(text) + char_limit - 1) // char_limit, max_rows)} rows"
                )
                continue

            # 「経営者の略歴等」は特殊処理（複数行に分割して転記）
            if key == "background":
                print(
                    f"[DEBUG] Processing background (略歴) with special formatting..."
                )
                lines = content.strip().split("\n")
                print(f"[DEBUG] Found {len(lines)} lines in background")

                # 最大6行まで処理（B13/H13, B14/H14, B15/H15, B16/H16, B17/H17, B18/H18）
                for idx, line in enumerate(lines[:6]):
                    line = line.strip()
                    if not line:
                        continue

                    row_num = 13 + idx  # 13, 14, 15, 16, 17, 18

                    # 年月と内容を分離（例: "・平成XX年3月 〇〇大学工学部 卒業"）
                    # 「・」や「-」で始まる場合は除去
                    line = line.lstrip("・-− ")

                    # 年月部分を抽出（例: "平成XX年3月" または "20XX年X月"）
                    year_month_match = re.match(
                        r"^([0-9]{4}年[0-9]{1,2}月|[平成令和昭和]+[0-9元]{1,2}年[0-9]{1,2}月)",
                        line,
                    )

                    if year_month_match:
                        year_month = year_month_match.group(1)
                        content_text = line[len(year_month) :].strip()
                    else:
                        # 年月が見つからない場合は、全体を内容として扱う
                        year_month = ""
                        content_text = line

                    # B列に年月、H列に内容を書き込み
                    if year_month:
                        sheet[f"B{row_num}"].value = year_month
                        print(
                            f"[DEBUG]   ✓ Wrote year/month to B{row_num}: '{year_month}'"
                        )

                    if content_text:
                        sheet[f"H{row_num}"].value = content_text
                        print(
                            f"[DEBUG]   ✓ Wrote content to H{row_num}: '{content_text[:50]}...'"
                        )

                print(
                    f"[DEBUG] ✓ Successfully processed background with {min(len(lines), 6)} rows"
                )
                continue

            # その他のセクションは通常通り処理
            cell_addr = mapping.get(key)
            if cell_addr:
                try:
                    sheet[cell_addr].value = content
                    print(
                        f"[DEBUG] ✓ Successfully wrote {key} ({len(content)} chars) to {cell_addr}"
                    )
                    print(f"[DEBUG]   Preview: '{content[:80]}...'")
                except Exception as e:
                    print(f"[ERROR] ✗ Could not write {key} to cell {cell_addr}: {e}")
            else:
                print(f"[WARNING] No cell mapping found for {key}")
        else:
            print(f"[DEBUG] Skipping {key} (no content)")

    # メモリ上のバイナリとして保存 (Excel)
    excel_output = BytesIO()
    workbook.save(excel_output)
    excel_output.seek(0)

    # 業種判定とPDF選択
    examples_dir = BASE_DIR / "static" / "templates" / "examples"
    pdf_filename = None

    # 業種タイプ→PDFファイルのマッピング
    industry_pdf_map = {
        "software": "software_example.pdf",
        "restaurant": "restaurant_example.pdf",
        "beauty": "beauty_example.pdf",
        "apparel": "apparel_example.pdf",
        "construction": "construction_example.pdf",
        "cram_school": "cram_school_example.pdf",
        "care_service": "care_service_example.pdf",
        "car_sales": "car_sales_example.pdf",
        "dentist": "dentist_example.pdf",
    }

    # セッションに保存された業種タイプを取得
    industry_type = data.get("industry_type")
    if industry_type and industry_type in industry_pdf_map:
        pdf_filename = industry_pdf_map[industry_type]
        print(
            f"[DEBUG] Using industry type from session: {industry_type} -> {pdf_filename}"
        )
    else:
        # 業種タイプが保存されていない場合はエラー
        print(
            "[WARNING] Industry type not found in session. User may have skipped selection."
        )
        # デフォルトのPDFを使用しない（業種選択は必須）
        pdf_filename = None

    # ZIPファイルの作成
    zip_output = BytesIO()
    with zipfile.ZipFile(zip_output, "w", zipfile.ZIP_DEFLATED) as zf:
        # Add Excel
        excel_filename = f"創業計画書_{session_id[:8]}.xlsx"
        zf.writestr(excel_filename, excel_output.getvalue())

        # Add PDF if found
        if pdf_filename:
            pdf_path = examples_dir / pdf_filename
            print(f"[DEBUG] Looking for PDF at: {pdf_path}")
            if pdf_path.exists():
                zf.write(pdf_path, arcname="創業計画書記入例.pdf")
            else:
                print(f"[WARNING] PDF file not found: {pdf_path}")
        else:
            print("[DEBUG] No matching industry found for PDF example.")

    zip_output.seek(0)
    zip_filename = f"創業計画書一式_{session_id[:8]}.zip"
    # URLエンコードしたファイル名をRFC 5987形式で指定
    from urllib.parse import quote

    encoded_filename = quote(zip_filename)
    headers = {
        "Content-Disposition": f"attachment; filename*=UTF-8''{encoded_filename}"
    }

    return StreamingResponse(
        zip_output,
        headers=headers,
        media_type="application/zip",
    )


@app.post(
    "/reset",
    response_class=HTMLResponse,
    summary="セッションのリセット",
    description="""
    現在のセッションを完全にリセットし、初期状態に戻します。

    ⚠️ **Swagger UIでの直接テストはできません**
    - このエンドポイントはCookieによるセッション管理が必要です
    - Swagger UIではCookieパラメータを正しく送信できません

    **テスト方法:**
    - 下記の「Try it out」→「session_id」を入力→「Execute」で表示されるCurlコマンドをコピー
    - ターミナルで実行
    """,
)
async def reset_session(
    request: Request,
    session_id: str | None = Cookie(default=None),
    db: AsyncSession = Depends(get_db),
):
    """
    現在のセッションを完全にリセットし、初期状態に戻します。

    サーバー側のセッションデータと会話履歴を削除し、
    トップページへリダイレクトします。

    Args:
        request (Request): FastAPIのリクエストオブジェクト
        session_id (str | None): CookieセッションID
        db (AsyncSession): データベースセッション

    Returns:
        Response: HX-Redirectヘッダーを含むレスポンス
    """
    if session_id:
        # セッションデータを削除
        await session_store.delete_session(db, session_id)
        # チャットセッションもリセット
        gemini_service.reset_chat_session(session_id)

    # htmxのリダイレクト機能を使用
    from fastapi.responses import Response

    response = Response(status_code=200)
    response.headers["HX-Redirect"] = "/"
    return response
