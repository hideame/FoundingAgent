import os
import re
import uuid
import zipfile
from io import BytesIO
from pathlib import Path

import openpyxl
from dotenv import load_dotenv
from fastapi import Cookie, FastAPI, Form, Request, Response
from fastapi.responses import HTMLResponse, StreamingResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
from openpyxl.utils import get_column_letter, range_boundaries

from app.services.gemini_service import GeminiService
from app.store import session_store

# Load environment variables first
load_dotenv()

app = FastAPI(title="Founder's Pilot")

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
        "title": "1. 創業の動機",
        "desc": "なぜこの事業を始めるのか、熱意と目的を言語化しましょう。",
        "status": "pending",
    },
    {
        "id": "background",
        "title": "2. 経営者の略歴等",
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
        "id": "partners",
        "title": "4. 取引先・取引関係等",
        "desc": "販売先や仕入先、掛取引の条件などを整理します。",
        "status": "pending",
    },
    {
        "id": "employees",
        "title": "5. 従業員",
        "desc": "常勤役員、従業員数などの体制を計画します。",
        "status": "pending",
    },
    {
        "id": "loans",
        "title": "6. お借入の状況",
        "desc": "個人的な借り入れや住宅ローンなどの状況を確認します。",
        "status": "pending",
    },
    {
        "id": "funds",
        "title": "7. 必要な資金と調達方法",
        "desc": "設備資金・運転資金の総額と、自己資金・借入金のバランスを計算します。",
        "status": "pending",
    },
    {
        "id": "outlook",
        "title": "8. 事業の見通し",
        "desc": "創業当初と軌道に乗った後の売上・利益予測を立てます。",
        "status": "pending",
    },
]

# Static files and Templates setup
BASE_DIR = Path(__file__).resolve().parent
app.mount("/static", StaticFiles(directory=str(BASE_DIR / "static")), name="static")
templates = Jinja2Templates(directory=str(BASE_DIR / "templates"))


@app.get("/", response_class=HTMLResponse)
async def read_root(request: Request, session_id: str | None = Cookie(default=None)):
    # セッションIDがあればデータをロード
    user_tasks = TASKS  # デフォルト
    chat_messages_html = ""  # ここで履歴HTML文字列を作る

    if session_id:
        data = session_store.load_session(session_id)
        if data:
            user_tasks = data.get("tasks", TASKS)
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
async def start_chat(request: Request, task_id: str | None = None):
    """
    チャットを開始するエンドポイント。
    新規セッションIDを発行してCookieにセットし、初期メッセージを返します。
    task_id が指定されている場合、そのタスクに合わせて初期メッセージを生成します。
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
):
    """
    ユーザーからのメッセージを受け取り、チャット履歴に追加して応答を返します。
    Cookieからsession_idを取得して利用します。
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
    session_data = session_store.load_session(session_id)
    if session_data and "tasks" in session_data:
        current_tasks = copy.deepcopy(session_data["tasks"])

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
    # 1. GeminiServiceから現在のチャット履歴を取得
    history_data = gemini_service.get_chat_history(session_id)
    # 2. 更新されたタスク状態を保存
    session_store.save_session(session_id, current_tasks, history_data)

    return response


@app.get("/plan/edit", response_class=HTMLResponse)
async def edit_plan(request: Request, session_id: str | None = Cookie(default=None)):
    """
    創業計画書の編集画面を表示します。
    """
    if not session_id:
        return HTMLResponse("Session not found", status_code=400)

    data = session_store.load_session(session_id)
    plan_text = ""
    if data:
        plan_text = data.get("plan_text", "")

    return templates.TemplateResponse(
        "components/plan_editor.html",
        {"request": request, "plan_text": plan_text},
    )


@app.post("/plan/save", response_class=HTMLResponse)
async def save_plan(
    request: Request,
    plan_text: str = Form(...),
    session_id: str | None = Cookie(default=None),
):
    """
    編集された創業計画書を保存し、閲覧モードに戻ります。
    """
    if not session_id:
        return HTMLResponse("Session not found", status_code=400)

    # 現在のセッションデータをロード
    data = session_store.load_session(session_id)
    if not data:
        data = {"tasks": TASKS, "chat_history": []}

    # 保存
    tasks = data.get("tasks", TASKS)
    history = data.get("chat_history", [])
    session_store.save_session(session_id, tasks, history, plan_text=plan_text)

    # 閲覧モードのHTMLを返す (plan_viewer.htmlのOOB部分は不要なので、plan_viewer.htmlの中身だけを返したいが、
    # plan_viewer.htmlはOOB含んでいる。
    # OOB部分は空にしてレンダリングするか、viewer専用のHTMLを分けるべきだが、
    # ここではOOB変数を空文字にしてレンダリングする。

    return templates.TemplateResponse(
        "components/plan_viewer.html",
        {"request": request, "plan_text": plan_text, "stepper_oob": ""},
    )


@app.post("/plan/generate", response_class=HTMLResponse)
async def generate_plan(
    request: Request, session_id: str | None = Cookie(default=None)
):
    """
    創業計画書のドラフトを生成し、表示します。
    ステッパーの状態も更新します。
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

    # --- 生成されたプランの保存 ---
    # セッションストアに plan_text を保存しておく
    current_data = session_store.load_session(session_id)
    if current_data:
        saved_tasks = current_data.get("tasks", TASKS)
        saved_history = current_data.get("chat_history", [])
        session_store.save_session(
            session_id, saved_tasks, saved_history, plan_text=plan_text
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
async def download_plan_excel(session_id: str | None = Cookie(default=None)):
    """
    保存された創業計画書データをExcelテンプレートに書き込み、ダウンロードさせます。
    """
    if not session_id:
        return Response("Session not found", status_code=400)

    data = session_store.load_session(session_id)
    if not data or not data.get("plan_text"):
        return Response("Plan text not found", status_code=404)

    plan_text = data.get("plan_text")

    # テンプレート読み込み
    template_path = BASE_DIR / "static" / "templates" / "startup_plan_template.xlsx"
    workbook = openpyxl.load_workbook(template_path)
    sheet = workbook.active

    # --- plan_textの解析とマッピング ---
    # テキストからセクションを抽出してセルに埋め込む
    # 出力フォーマットの揺れに対応できるよう、見出しの前後パターンを緩めに扱う

    heading_order = [
        ("motivation", "創業の動機"),
        ("background", "経営者の略歴等"),
        ("service", "取扱商品・サービス"),
        ("partners", "取引先・取引関係等"),
        ("employees", "従業員"),
        ("loans", "お借入の状況"),
        ("funds", "必要な資金と調達方法"),
        ("outlook", "事業の見通し"),
    ]

    def build_section_pattern(current_label, next_labels):
        next_part = "|".join(re.escape(label) for label in next_labels)
        return rf"(?:^|\n)\s*(?:\d+[\.|\s]*)?{re.escape(current_label)}\s*(?:\n|:|：)\s*(.*?)(?=\n\s*(?:\d+[\.|\s]*)?(?:{next_part})\s*(?:\n|:|：)|\Z)"

    sections = {}
    for idx, (key, label) in enumerate(heading_order):
        next_labels = [next_label for _, next_label in heading_order[idx + 1 :]]
        pattern = (
            build_section_pattern(label, next_labels)
            if next_labels
            else build_section_pattern(label, [])
        )
        sections[key] = pattern

    def normalize_section_text(content: str) -> str:
        content = content.strip()
        content = re.sub(r"[\*#]+", "", content)
        content = re.sub(r"\n\s*\n+", "\n\n", content)
        return content.strip()

    def extract_section(pattern, text):
        match = re.search(pattern, text, re.DOTALL | re.IGNORECASE)
        if match:
            return normalize_section_text(match.group(1))
        return ""

    print(f"[DEBUG] plan_text length: {len(plan_text)}")
    print(f"[DEBUG] plan_text preview: {plan_text[:500]}")

    # テンプレート内の見出しセルを探して、転記先セルを動的に推定する
    label_to_key = {
        "創業の動機": "motivation",
        "経営者の略歴等": "background",
        "取扱商品・サービス": "service",
        "取引先・取引関係等": "partners",
        "従業員": "employees",
        "お借入の状況": "loans",
        "必要な資金と調達方法": "funds",
        "事業の見通し": "outlook",
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

    # フォールバック用の静的マッピング（テンプレート検出できない場合）
    fallback_mapping = {
        "motivation": "A9",
        "background": "A16",
        "service": "A27",
        "partners": "M27",
        "employees": "T21",
        "loans": "M36",
        "funds": "A45",
        "outlook": "M45",
    }
    for key, addr in fallback_mapping.items():
        if not mapping.get(key):
            mapping[key] = addr

    # セル結合されていることが多いので、左上のセルに値をセットする
    extracted_any = False
    for key, pattern in sections.items():
        content = extract_section(pattern, plan_text)
        print(f"[DEBUG] Extracted {key}: {content[:100] if content else '(empty)'}...")
        cell_addr = mapping.get(key)
        if cell_addr and content:
            extracted_any = True
            # 結合セルの場合でも安全に値を設定
            try:
                cell = sheet[cell_addr]
                # MergedCellの場合、unmergeしてから書き込む
                if isinstance(cell, openpyxl.cell.cell.MergedCell):
                    # 結合範囲を探して解除
                    for merged_range in list(sheet.merged_cells.ranges):
                        if cell.coordinate in merged_range:
                            sheet.unmerge_cells(str(merged_range))
                            break
                # 値を設定
                sheet[cell_addr].value = content
                print(f"[DEBUG] Successfully wrote to {cell_addr}")
            except Exception as e:
                print(f"Warning: Could not write to cell {cell_addr}: {e}")
                continue

    if not extracted_any and plan_text.strip():
        fallback_cell = mapping.get("motivation")
        if fallback_cell:
            try:
                sheet[fallback_cell].value = normalize_section_text(plan_text)
                print(
                    "[DEBUG] No sections matched. Wrote full plan_text to motivation cell as fallback."
                )
            except Exception as e:
                print(f"Warning: Could not write fallback plan_text: {e}")

    # メモリ上のバイナリとして保存 (Excel)
    excel_output = BytesIO()
    workbook.save(excel_output)
    excel_output.seek(0)

    # 業種判定とPDF選択
    examples_dir = BASE_DIR / "static" / "templates" / "examples"
    pdf_filename = None

    # Determine keywords based on plan_text
    # Simple keyword matching to guess industry
    if re.search(r"飲食|居酒屋|カフェ|レストラン|食堂|ランチ|ディナー|料理", plan_text):
        pdf_filename = "restaurant_example.pdf"
    elif re.search(r"美容|サロン|ヘア|ネイル|エステ|カット|パーマ", plan_text):
        pdf_filename = "beauty_example.pdf"
    elif re.search(r"自動車|中古車|車両|整備|板金|ガソリン", plan_text):
        pdf_filename = "car_sales_example.pdf"
    elif re.search(
        r"アパレル|洋服|服飾|衣料|婦人服|子供服|ベビー服|ファッション|ブティック",
        plan_text,
    ):
        pdf_filename = "apparel_example.pdf"
    elif re.search(
        r"内装|工事|建築|リフォーム|リノベーション|施工|塗装|配管|電気工事", plan_text
    ):
        pdf_filename = "construction_example.pdf"
    elif re.search(
        r"学習塾|塾|予備校|個別指導|教室|スクール|講師|生徒|授業|受験", plan_text
    ):
        pdf_filename = "cram_school_example.pdf"
    elif re.search(
        r"歯科|歯医者|デンタル|クリニック|矯正|インプラント|衛生士", plan_text
    ):
        pdf_filename = "dentist_example.pdf"
    elif re.search(
        r"介護|デイサービス|福祉|ヘルパー|ケアマネ|老人|高齢者|リハビリ|支援", plan_text
    ):
        pdf_filename = "care_service_example.pdf"
    elif re.search(
        r"ソフトウェア|システム|アプリ|開発|IT|Web|ウェブ|エンジニア|プログラミング",
        plan_text,
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
                # 日本語ファイル名にする: "記入例_業種名.pdf" のようにするとなお良いが、
                # シンプルに "記入例.pdf" または元のファイル名でも良い。
                # ここではわかりやすく "創業計画書記入例.pdf" とする
                zf.write(pdf_path, arcname="創業計画書記入例.pdf")
            else:
                print(f"[WARNING] PDF file not found: {pdf_path}")
        else:
            print("[DEBUG] No matching industry found for PDF example.")

    zip_output.seek(0)
    zip_filename = f"創業計画書一式_{session_id[:8]}.zip"
    headers = {"Content-Disposition": f'attachment; filename="{zip_filename}"'}

    return StreamingResponse(
        zip_output,
        headers=headers,
        media_type="application/zip",
    )


@app.post("/reset", response_class=HTMLResponse)
async def reset_session(
    request: Request, session_id: str | None = Cookie(default=None)
):
    """
    セッションをリセットします。
    チャット履歴、タスク進捗を初期化し、再び最初からやり直せるようにします。
    """
    if session_id:
        # セッションデータを削除
        session_store.delete_session(session_id)
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
