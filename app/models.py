"""
SQLAlchemyデータベースモデル定義

このモジュールは、アプリケーションで使用する全てのデータベーステーブルのモデルを定義します。
"""

from datetime import datetime

from sqlalchemy import Column, DateTime, ForeignKey, Integer, String, Text
from sqlalchemy.orm import declarative_base, relationship

Base = declarative_base()


class Session(Base):
    """
    ユーザーセッション情報を管理するテーブル

    各ユーザーの訪問セッションを一意に識別し、
    関連する事業計画やチャット履歴との紐付けを行います。
    """

    __tablename__ = "sessions"

    id = Column(String(36), primary_key=True, comment="セッションID（UUID）")
    created_at = Column(
        DateTime, default=datetime.utcnow, nullable=False, comment="作成日時"
    )
    updated_at = Column(
        DateTime,
        default=datetime.utcnow,
        onupdate=datetime.utcnow,
        nullable=False,
        comment="更新日時",
    )

    # リレーション
    business_plan = relationship(
        "BusinessPlan",
        back_populates="session",
        uselist=False,
        cascade="all, delete-orphan",
    )
    chat_messages = relationship(
        "ChatMessage", back_populates="session", cascade="all, delete-orphan"
    )
    task_states = relationship(
        "TaskState", back_populates="session", cascade="all, delete-orphan"
    )


class BusinessPlan(Base):
    """
    創業計画書の各項目データを管理するテーブル

    日本政策金融公庫の創業計画書フォーマットに基づく9つの項目を格納します。
    """

    __tablename__ = "business_plans"

    id = Column(Integer, primary_key=True, autoincrement=True, comment="事業計画ID")
    session_id = Column(
        String(36),
        ForeignKey("sessions.id", ondelete="CASCADE"),
        nullable=False,
        unique=True,
        comment="セッションID",
    )

    # 創業計画書の9項目
    motivation = Column(Text, comment="1. 創業の動機")
    background = Column(Text, comment="2. 経営者の略歴等")
    service = Column(Text, comment="3. 取扱商品・サービス")
    employees = Column(Text, comment="4. 従業員")
    partners = Column(Text, comment="5. 取引先・取引関係等")
    related_companies = Column(Text, comment="6. 関連企業")
    loans = Column(Text, comment="7. お借入の状況")
    funds = Column(Text, comment="8. 必要な資金と調達方法")
    outlook = Column(Text, comment="9. 事業の見通し")

    # 全文テキスト（マークダウン形式の完全版）
    full_text = Column(Text, comment="完全なドラフトテキスト（マークダウン）")

    created_at = Column(
        DateTime, default=datetime.utcnow, nullable=False, comment="作成日時"
    )
    updated_at = Column(
        DateTime,
        default=datetime.utcnow,
        onupdate=datetime.utcnow,
        nullable=False,
        comment="更新日時",
    )

    # リレーション
    session = relationship("Session", back_populates="business_plan")


class ChatMessage(Base):
    """
    チャット履歴を管理するテーブル

    ユーザーとAI（Gemini）の対話履歴を時系列で保存します。
    """

    __tablename__ = "chat_messages"

    id = Column(Integer, primary_key=True, autoincrement=True, comment="メッセージID")
    session_id = Column(
        String(36),
        ForeignKey("sessions.id", ondelete="CASCADE"),
        nullable=False,
        comment="セッションID",
    )
    role = Column(String(20), nullable=False, comment="発言者（user/model）")
    content = Column(Text, nullable=False, comment="メッセージ内容")
    created_at = Column(
        DateTime, default=datetime.utcnow, nullable=False, comment="作成日時"
    )

    # リレーション
    session = relationship("Session", back_populates="chat_messages")


class TaskState(Base):
    """
    タスク進捗状態を管理するテーブル

    ステッパーUIに表示される各タスクの完了状態を保存します。
    """

    __tablename__ = "task_states"

    id = Column(Integer, primary_key=True, autoincrement=True, comment="タスク状態ID")
    session_id = Column(
        String(36),
        ForeignKey("sessions.id", ondelete="CASCADE"),
        nullable=False,
        comment="セッションID",
    )
    task_key = Column(
        String(50), nullable=False, comment="タスクキー（motivation, background等）"
    )
    completed = Column(
        Integer, default=0, nullable=False, comment="完了フラグ（0: 未完了, 1: 完了）"
    )
    created_at = Column(
        DateTime, default=datetime.utcnow, nullable=False, comment="作成日時"
    )
    updated_at = Column(
        DateTime,
        default=datetime.utcnow,
        onupdate=datetime.utcnow,
        nullable=False,
        comment="更新日時",
    )

    # リレーション
    session = relationship("Session", back_populates="task_states")
