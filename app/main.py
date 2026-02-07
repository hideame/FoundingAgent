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

from app.database import close_db, get_db, init_db
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

app = FastAPI(title="Founder's Pilot")


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

    タスク完了時に、AI応答にドラフト提案が含まれている場合、
    そのセクションの内容だけを抽出して返します。

    Args:
        ai_response (str): AIからの応答テキスト
        task_id (str): タスクID（例: "motivation", "background"）

    Returns:
        str | None: 抽出されたセクション内容。見つからない場合はNone
    """
    # タスクIDから日本語の見出しにマッピング
    task_to_label = {
        "motivation": "創業の動機",
        "background": "経営者の略歴等",
        "service": "取扱商品・サービス",
        "employees": "従業員",
        "partners": "取引先・取引関係等",
        "related_companies": "関連企業",
        "loans": "お借入の状況",
        "funds": "必要な資金と調達方法",
        "outlook": "事業の見通し",
        "free_description": "自由記述欄",
    }

    label = task_to_label.get(task_id)
    if not label:
        return None

    print(f"[DEBUG] Searching for label: '{label}' in AI response")

    # 見出しパターンでセクションを抽出（柔軟なパターンマッチング）
    # "## 1. 創業の動機" や "創業の動機:" などに対応
    pattern = rf"(?:##\s*\d*\.?\s*)?{re.escape(label)}\s*(?:\n|:|：)\s*(.*?)(?=(?:##\s*\d*\.?\s*(?:{'|'.join(task_to_label.values())})|$))"

    print(f"[DEBUG] Regex pattern: {pattern[:200]}")

    match = re.search(pattern, ai_response, re.DOTALL | re.MULTILINE)
    if match:
        print(f"[DEBUG] Match found! Extracting content...")
        content = match.group(1).strip()
        # マークダウン記号やマーカーを除去
        content = re.sub(r"[\*#]+", "", content)
        content = re.sub(r"\[\[DRAFT_PROPOSED\]\]", "", content)
        content = re.sub(r"\[\[COMPLETED:[a-z_]+\]\]", "", content)
        content = re.sub(r"\n\s*\n+", "\n\n", content)
        return content.strip()
    else:
        print(f"[DEBUG] No match found for label '{label}'")
        print(
            f"[DEBUG] Available labels in response: {[label_name for label_name in task_to_label.values() if label_name in ai_response]}"
        )

    return None


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

    initial_task_title = None
    if task_id:
        # TASKSからタイトルを検索
        for task in TASKS:
            if task["id"] == task_id:
                initial_task_title = task["title"]
                break

    # エージェントからの初期挨拶を生成
    # GeminiService.start_chat_session requires session_id
    tGreeting = await gemini_service.start_chat_session(session_id, initial_task_title)

    response = templates.TemplateResponse(
        "components/chat_interface.html",
        {
            "request": request,
            "message": tGreeting,
            "is_user": False,
        },
    )
    # 簡易的にCookieでセッションIDを管理
    response.set_cookie(key="session_id", value=session_id)
    return response


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

    # タスク完了マーカーの検出と処理
    task_update_html = ""
    draft_buttons_html = ""

    # ドラフト提示マーカーの検出
    if "[[DRAFT_PROPOSED]]" in ai_response_text:
        ai_response_text = ai_response_text.replace("[[DRAFT_PROPOSED]]", "")
        draft_buttons_html = """
        <div class="flex gap-4 mt-2 mb-4 ml-12">
            <button hx-post="/chat/message"
                    hx-vals='{"user_message": "この内容でOKです"}'
                    hx-target="#chat-history"
                    hx-swap="beforeend"
                    class="px-4 py-2 bg-blue-600 text-white rounded-lg text-sm hover:bg-blue-700 font-bold transition-colors">
                OK（次の項目へ）
            </button>
            <button hx-post="/chat/message"
                    hx-vals='{"user_message": "文面を修正したいです"}'
                    hx-target="#chat-history"
                    hx-swap="beforeend"
                    class="px-4 py-2 bg-white border border-gray-300 text-gray-700 rounded-lg text-sm hover:bg-gray-50 transition-colors">
                文面を修正する
            </button>
        </div>
        """

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

    print(
        f"[DEBUG] Loaded tasks: {[t['id'] + ':' + t['status'] for t in current_tasks]}"
    )

    match = re.search(r"\[\[COMPLETED:([a-z_]+)\]\]", ai_response_text)
    if match:
        completed_task_id = match.group(1)
        print(f"[DEBUG] Found COMPLETED marker for task: {completed_task_id}")
        # マーカーを応答テキストから削除
        ai_response_text = ai_response_text.replace(match.group(0), "")

        # タスクステータスの更新（セッション内のタスクリストを更新）
        for task in current_tasks:
            if task["id"] == completed_task_id:
                task["status"] = "done"
                print(f"[DEBUG] Updated task {completed_task_id} to done")

                # 更新されたチェックボックスのHTMLを生成 (OOB Swap用)
                # targetは id="task-{id}"
                task_update_html = f"""
                <input id="task-{completed_task_id}"
                       name="task-{completed_task_id}"
                       type="checkbox"
                       class="h-4 w-4 rounded border-gray-300 text-blue-600 focus:ring-blue-600"
                       checked
                       hx-swap-oob="true">
                """
                print(
                    f"[DEBUG] Generated OOB HTML for checkbox: task-{completed_task_id}"
                )

                # --- セクション内容の抽出と保存 ---
                # AI応答から該当セクションの内容を抽出
                print(f"[DEBUG] AI response for extraction: {ai_response_text[:500]}")
                section_content = extract_single_section_from_response(
                    ai_response_text, completed_task_id
                )

                if section_content:
                    print(
                        f"[DEBUG] Extracted section content for {completed_task_id}: {section_content[:100]}..."
                    )
                    # 該当セクションのみをデータベースに保存
                    # 既存のチャット履歴は後でまとめて保存するので、ここでは sections だけ更新
                    # 注: save_sessionは既存セクションを上書きしないので、部分更新が可能
                    section_update = {completed_task_id: section_content}

                    # 現在のタスク状態とチャット履歴を取得して保存
                    history_data_temp = gemini_service.get_chat_history(session_id)
                    await session_store.save_session(
                        db,
                        session_id,
                        current_tasks,
                        history_data_temp,
                        sections=section_update,
                    )
                    print(f"[DEBUG] Saved section {completed_task_id} to database")
                else:
                    print(
                        f"[WARNING] Could not extract section content for {completed_task_id}"
                    )

                break

    ai_msg_html = templates.get_template("components/message.html").render(
        {"request": request, "message": ai_response_text, "is_user": False}
    )

    response = HTMLResponse(
        content=user_msg_html + ai_msg_html + draft_buttons_html + task_update_html
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


@app.get("/plan/edit", response_class=HTMLResponse)
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
    plan_text = ""
    if data and data.get("sections"):
        # セクションから全文を再構築
        plan_text = build_plan_text_from_sections(data["sections"])

    return templates.TemplateResponse(
        "components/plan_editor.html",
        {"request": request, "plan_text": plan_text},
    )


@app.post("/plan/save", response_class=HTMLResponse)
async def save_plan(
    request: Request,
    plan_text: str = Form(...),
    session_id: str | None = Cookie(default=None),
    db: AsyncSession = Depends(get_db),
):
    """
    編集された創業計画書テキストを保存します。

    全文テキストを各セクションに分割してデータベースに保存し、
    保存後は閲覧モード（Viewer）のHTMLを返します。

    Args:
        request (Request): FastAPIのリクエストオブジェクト
        plan_text (str): フォームから送信された計画書テキスト
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

    # 全文テキストをセクション別に分割
    sections = extract_sections_from_text(plan_text)

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

    # 閲覧モードのHTMLを返す
    return templates.TemplateResponse(
        "components/plan_viewer.html",
        {"request": request, "plan_text": plan_text, "stepper_oob": ""},
    )


@app.post("/plan/generate", response_class=HTMLResponse)
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

    # Geminiでドラフト生成
    plan_text = await gemini_service.generate_business_plan(session_id)

    # マーカーを除去（表示用に不要な制御コマンドを削除）
    plan_text = re.sub(r"\[\[DRAFT_PROPOSED\]\]", "", plan_text)
    plan_text = re.sub(r"\[\[COMPLETED:[a-z_]+\]\]", "", plan_text)

    # 全文テキストをセクション別に分割
    sections = extract_sections_from_text(plan_text)

    # --- 生成されたプランの保存 ---
    current_data = await session_store.load_session(db, session_id)
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


@app.get("/plan/download_excel")
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

    # データベースから取得したセクションデータをExcelに転記
    for key, content in sections.items():
        if content:
            cell_addr = mapping.get(key)
            if cell_addr:
                try:
                    sheet[cell_addr].value = content
                    print(
                        f"[DEBUG] Successfully wrote '{content[:50]}...' to {cell_addr}"
                    )
                except Exception as e:
                    print(f"[ERROR] Could not write to cell {cell_addr}: {e}")

    # メモリ上のバイナリとして保存 (Excel)
    excel_output = BytesIO()
    workbook.save(excel_output)
    excel_output.seek(0)

    # 業種判定とPDF選択（全セクションの内容を結合して判定）
    all_content = " ".join([str(v) for v in sections.values() if v])
    examples_dir = BASE_DIR / "static" / "templates" / "examples"
    pdf_filename = None

    # 業種キーワードマッチング
    if re.search(
        r"飲食|居酒屋|カフェ|レストラン|食堂|ランチ|ディナー|料理", all_content
    ):
        pdf_filename = "restaurant_example.pdf"
    elif re.search(r"美容|サロン|ヘア|ネイル|エステ|カット|パーマ", all_content):
        pdf_filename = "beauty_example.pdf"
    elif re.search(r"自動車|中古車|車両|整備|板金|ガソリン", all_content):
        pdf_filename = "car_sales_example.pdf"
    elif re.search(
        r"アパレル|洋服|服飾|衣料|婦人服|子供服|ベビー服|ファッション|ブティック",
        all_content,
    ):
        pdf_filename = "apparel_example.pdf"
    elif re.search(
        r"内装|工事|建築|リフォーム|リノベーション|施工|塗装|配管|電気工事", all_content
    ):
        pdf_filename = "construction_example.pdf"
    elif re.search(
        r"学習塾|塾|予備校|個別指導|教室|スクール|講師|生徒|授業|受験", all_content
    ):
        pdf_filename = "cram_school_example.pdf"
    elif re.search(
        r"歯科|歯医者|デンタル|クリニック|矯正|インプラント|衛生士", all_content
    ):
        pdf_filename = "dentist_example.pdf"
    elif re.search(
        r"介護|デイサービス|福祉|ヘルパー|ケアマネ|老人|高齢者|リハビリ|支援",
        all_content,
    ):
        pdf_filename = "care_service_example.pdf"
    elif re.search(
        r"ソフトウェア|システム|アプリ|開発|IT|Web|ウェブ|エンジニア|プログラミング",
        all_content,
    ):
        pdf_filename = "software_example.pdf"

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


@app.post("/reset", response_class=HTMLResponse)
async def reset_session(
    request: Request,
    session_id: str | None = Cookie(default=None),
    db: AsyncSession = Depends(get_db),
):
    """
    現在のセッションを完全にリセットし、初期状態に戻します。

    サーバー側のセッションデータと会話履歴を削除し、
    トップページへリダイレクト（クライアントサイドリダイレクト）を行います。

    Args:
        request (Request): FastAPIのリクエストオブジェクト
        session_id (str | None): CookieセッションID
        db (AsyncSession): データベースセッション

    Returns:
        HTMLResponse: リダイレクト用のJavaScriptを含むHTML
    """
    if session_id:
        # セッションデータを削除
        await session_store.delete_session(db, session_id)
        # チャットセッションもリセット
        gemini_service.reset_chat_session(session_id)

    # ページ全体をリロード
    return HTMLResponse(
        """
        <script>
            window.location.href = '/';
        </script>
        """,
        status_code=200,
    )
