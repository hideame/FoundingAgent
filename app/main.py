import os
from pathlib import Path

from dotenv import load_dotenv

# Load environment variables first
load_dotenv()

from fastapi import FastAPI, Form, Request
from fastapi.responses import HTMLResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates

from app.services.gemini_service import GeminiService

app = FastAPI(title="Founder's Pilot")

# Initialize Gemini Service
gemini_service = GeminiService()

# Static files and Templates setup
BASE_DIR = Path(__file__).resolve().parent
app.mount("/static", StaticFiles(directory=str(BASE_DIR / "static")), name="static")
templates = Jinja2Templates(directory=str(BASE_DIR / "templates"))


@app.get("/", response_class=HTMLResponse)
async def read_root(request: Request):
    return templates.TemplateResponse("index.html", {"request": request})


@app.get("/chat/start", response_class=HTMLResponse)
async def start_chat(request: Request):
    """
    チャットを開始するエンドポイント。
    チャットインターフェースと初期メッセージを返します。
    """
    # エージェントからの初期挨拶を生成
    tGreeting = await gemini_service.generate_greeting()

    return templates.TemplateResponse(
        "components/chat_interface.html",
        {
            "request": request,
            "message": tGreeting,
            "is_user": False,
        },
    )


@app.post("/chat/message", response_class=HTMLResponse)
async def chat_message(request: Request, user_message: str = Form(...)):
    """
    ユーザーからのメッセージを受け取り、チャット履歴に追加して応答を返します。
    (今回はモック実装としてエコーバック+αを返す)
    """
    # ユーザーのメッセージを表示するためのHTML
    user_msg_html = templates.get_template("components/message.html").render(
        {"request": request, "message": user_message, "is_user": True}
    )

    # 実際にはここでGeminiServiceに履歴を投げて応答を得る
    # response_text = await gemini_service.generate_response(user_message, history)

    # 仮の応答（後で実装）
    ai_response_text = f"承知いたしました。「{user_message}」ですね。それについてもう少し詳しく教えていただけますか？"

    ai_msg_html = templates.get_template("components/message.html").render(
        {"request": request, "message": ai_response_text, "is_user": False}
    )

    # ユーザーメッセージとAI応答を結合して返す
    return HTMLResponse(content=user_msg_html + ai_msg_html)


@app.get("/health")
async def health_check():
    return {"status": "ok"}
