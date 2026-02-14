"""
アプリケーションのデータ永続化を担当するモジュールです。

各ユーザーセッションのタスク進捗状況、チャット履歴、および
作成された事業計画書のテキストデータをMySQLデータベースに保存・管理します。
"""

from datetime import datetime

from sqlalchemy import delete, select
from sqlalchemy.ext.asyncio import AsyncSession
from sqlalchemy.orm import selectinload

from app.models import BusinessPlan, ChatMessage, Session, TaskState


class SessionStore:
    """
    セッションデータをデータベース上で管理するクラスです。

    各セッションデータは sessions テーブルに保存され、
    関連する事業計画、チャット履歴、タスク状態が紐付けられます。
    """

    async def save_session(
        self,
        db: AsyncSession,
        session_id: str,
        tasks_state: list,
        chat_history: list,
        sections: dict = None,
    ):
        """
        セッションの状態（タスク進捗とチャット履歴、計画書セクション）をデータベースに保存します。

        既存のデータがある場合、指定されなかったセクションは既存の値を維持します。

        Args:
            db (AsyncSession): データベースセッション
            session_id (str): 保存対象のセッションID
            tasks_state (list): タスクリストの辞書配列
            chat_history (list): チャット履歴の辞書配列 (Gemini API形式)
            sections (dict, optional): 更新する事業計画書セクションの辞書
                例: {"motivation": "創業の動機の内容", "background": "経営者の略歴等の内容"}
        """
        try:
            # セッションの取得または作成
            result = await db.execute(
                select(Session)
                .options(
                    selectinload(Session.business_plan),
                    selectinload(Session.chat_messages),
                    selectinload(Session.task_states),
                )
                .where(Session.id == session_id)
            )
            session = result.scalar_one_or_none()

            if not session:
                # 新規セッション作成
                session = Session(id=session_id)
                db.add(session)
                await db.flush()

            # タスク状態の更新
            # 既存のタスク状態を削除して再作成
            await db.execute(
                delete(TaskState).where(TaskState.session_id == session_id)
            )
            for task in tasks_state:
                # main.pyのTASKS形式に対応: {"id": "task_id", "status": "done"/"pending"}
                task_id = task.get("id", "")
                task_status = task.get("status", "pending")
                task_state = TaskState(
                    session_id=session_id,
                    task_key=task_id,
                    completed=1 if task_status == "done" else 0,
                )
                db.add(task_state)

            # チャット履歴の更新
            # 既存のメッセージを削除して再作成
            await db.execute(
                delete(ChatMessage).where(ChatMessage.session_id == session_id)
            )
            for msg in chat_history:
                chat_msg = ChatMessage(
                    session_id=session_id,
                    role=msg.get("role", ""),
                    content=msg.get("parts", [{}])[0].get("text", ""),
                )
                db.add(chat_msg)

            # 事業計画セクションの更新
            if sections is not None:
                # None値のセクションをフィルタリング（既存データを保持）
                valid_sections = {k: v for k, v in sections.items() if v is not None}

                if valid_sections:  # 有効なセクションがある場合のみ処理
                    if not session.business_plan:
                        # 新規作成
                        session.business_plan = BusinessPlan(session_id=session_id)
                        db.add(session.business_plan)

                    # セクションを更新（None以外のみ）
                    for key, value in valid_sections.items():
                        if hasattr(session.business_plan, key):
                            setattr(session.business_plan, key, value)
                        else:
                            print(f"[WARNING] save_session - {key} is not a valid attribute")

                    session.business_plan.updated_at = datetime.utcnow()

            await db.commit()

        except Exception as e:
            await db.rollback()
            print(f"Failed to save session {session_id}: {e}")
            raise

    async def load_session(self, db: AsyncSession, session_id: str) -> dict | None:
        """
        指定されたセッションIDのデータをデータベースから読み込みます。

        Args:
            db (AsyncSession): データベースセッション
            session_id (str): 読み込むセッションID

        Returns:
            dict | None: セッションデータの辞書。セッションが存在しない場合は None を返します。
        """
        try:
            result = await db.execute(
                select(Session)
                .options(
                    selectinload(Session.business_plan),
                    selectinload(Session.chat_messages),
                    selectinload(Session.task_states),
                )
                .where(Session.id == session_id)
            )
            session = result.scalar_one_or_none()

            if not session:
                return None

            # タスク状態を辞書配列に変換
            # データベースにはstatus(done/pending)のみ保存されているので、
            # main.pyでTASKSテンプレートとマージする必要がある
            # ここではidとstatusのみを返す
            task_states_dict = {}
            for task in session.task_states:
                task_states_dict[task.task_key] = (
                    "done" if task.completed else "pending"
                )

            # チャット履歴をGemini API形式に変換
            chat_history = [
                {
                    "role": msg.role,
                    "parts": [{"text": msg.content}],
                }
                for msg in session.chat_messages
            ]

            # 事業計画の各セクション
            sections = {}
            if session.business_plan:
                for section_key in [
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
                    value = getattr(session.business_plan, section_key, None)
                    sections[section_key] = value

            # 業種タイプをbusiness_planから取得
            industry_type = None
            if session.business_plan:
                industry_type = session.business_plan.industry_type

            return {
                "task_states": task_states_dict,  # {"task_id": "done"/"pending"}
                "chat_history": chat_history,
                "sections": sections,
                "industry_type": industry_type,  # 業種タイプを追加
            }

        except Exception as e:
            print(f"Failed to load session {session_id}: {e}")
            return None

    async def delete_session(self, db: AsyncSession, session_id: str):
        """
        指定されたセッションのデータを完全に削除します。
        データベースからセッションと関連データを削除します（カスケード削除）。

        Args:
            db (AsyncSession): データベースセッション
            session_id (str): 削除するセッションID
        """
        try:
            await db.execute(delete(Session).where(Session.id == session_id))
            await db.commit()
        except Exception as e:
            await db.rollback()
            print(f"Failed to delete session {session_id}: {e}")
            raise

    async def update_business_plan_sections(
        self,
        db: AsyncSession,
        session_id: str,
        sections: dict,
    ):
        """
        事業計画書の各セクションを個別に更新します。

        Args:
            db (AsyncSession): データベースセッション
            session_id (str): セッションID
            sections (dict): 更新するセクションの辞書
                例: {
                    "motivation": "創業の動機の内容",
                    "background": "経営者の略歴等の内容",
                    ...
                }
        """
        try:
            result = await db.execute(
                select(BusinessPlan).where(BusinessPlan.session_id == session_id)
            )
            business_plan = result.scalar_one_or_none()

            if not business_plan:
                # 事業計画が存在しない場合は新規作成
                business_plan = BusinessPlan(session_id=session_id)
                db.add(business_plan)

            # セクションを更新
            for key, value in sections.items():
                if hasattr(business_plan, key):
                    setattr(business_plan, key, value)

            business_plan.updated_at = datetime.utcnow()
            await db.commit()

        except Exception as e:
            await db.rollback()
            print(f"Failed to update business plan sections for {session_id}: {e}")
            raise

    async def update_industry_type(
        self, db: AsyncSession, session_id: str, industry_type: str
    ):
        """
        事業計画の業種タイプを更新します。
        business_plansレコードが存在しない場合は新規作成します。

        Args:
            db (AsyncSession): データベースセッション
            session_id (str): セッションID
            industry_type (str): 業種タイプ（software, restaurant, beauty等）
        """
        try:
            # business_plansレコードを取得または作成
            result = await db.execute(
                select(BusinessPlan).where(BusinessPlan.session_id == session_id)
            )
            business_plan = result.scalar_one_or_none()

            if business_plan:
                # 既存レコードを更新
                business_plan.industry_type = industry_type
                business_plan.updated_at = datetime.utcnow()
            else:
                # 新規レコードを作成
                business_plan = BusinessPlan(
                    session_id=session_id,
                    industry_type=industry_type,
                )
                db.add(business_plan)

            await db.commit()

        except Exception as e:
            await db.rollback()
            print(f"Failed to update industry type for {session_id}: {e}")
            raise


session_store = SessionStore()
