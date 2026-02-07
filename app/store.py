"""
アプリケーションのデータ永続化を担当するモジュールです。

各ユーザーセッションのタスク進捗状況、チャット履歴、および
作成された事業計画書のテキストデータをJSONファイルとしてファイルシステム上に保存・管理します。
"""

import json
from pathlib import Path

DATA_DIR = Path(__file__).resolve().parent.parent / "data"
DATA_DIR.mkdir(parents=True, exist_ok=True)


class SessionStore:
    """
    セッションデータをファイルシステム上で管理するクラスです。

    各セッションデータは data/ ディレクトリ直下に `{session_id}.json` という名前で保存されます。
    JSONファイルには、タスクの進捗状況、チャット履歴、生成された計画書テキストが含まれます。
    """

    def __init__(self):
        """
        SessionStoreを初期化します。
        データ保存ディレクトリ（DATA_DIR）を設定します。
        """
        self.data_dir = DATA_DIR

    def _get_file_path(self, session_id: str) -> Path:
        """
        指定されたセッションIDに対応するJSONファイルのパスを返します。

        Args:
            session_id (str): セッションID

        Returns:
            Path: JSONファイルのパス
        """
        return self.data_dir / f"{session_id}.json"

    def save_session(
        self,
        session_id: str,
        tasks_state: list,
        chat_history: list,
        plan_text: str = None,
    ):
        """
        セッションの状態（タスク進捗とチャット履歴、計画書ドラフト）をJSONファイルとして保存します。

        既存のデータがある場合、指定されなかったパラメータ（特にplan_text）は
        既存の値を維持するようにマージ処理が行われます。

        Args:
            session_id (str): 保存対象のセッションID
            tasks_state (list): タスクリストの辞書配列
            chat_history (list): チャット履歴の辞書配列 (Gemini API形式)
            plan_text (str, optional): 生成された事業計画書のテキスト全文。指定しない場合は既存値を維持します。
        """
        data = {"tasks": tasks_state, "chat_history": chat_history}
        if plan_text is not None:
            data["plan_text"] = plan_text
        # 既存のdataがあれば読み込んでマージしたほうが安全だが、
        # 今回はstore側でplan_textだけ更新するメソッドがないので
        # 呼び出し元で気をつけるか、ここをloadしてからmergeにする。
        # シンプルに実装するため、loadしてmergeする。
        existing_data = self.load_session(session_id)
        if existing_data:
            # plan_textがNone（引数未指定）の場合、既存のplan_textを維持する
            if plan_text is None:
                data["plan_text"] = existing_data.get("plan_text")
            # tasksやchat_historyは常に最新が渡される前提

        try:
            with open(self._get_file_path(session_id), "w", encoding="utf-8") as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
        except Exception as e:
            print(f"Failed to save session {session_id}: {e}")

    def load_session(self, session_id: str) -> dict | None:
        """
        指定されたセッションIDのデータをファイルから読み込みます。

        Args:
            session_id (str): 読み込むセッションID

        Returns:
            dict | None: セッションデータの辞書。ファイルが存在しない場合は None を返します。
        """
        file_path = self._get_file_path(session_id)
        if not file_path.exists():
            return None

        try:
            with open(file_path, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception as e:
            print(f"Failed to load session {session_id}: {e}")
            return None

    def delete_session(self, session_id: str):
        """
        指定されたセッションのデータを完全に削除します。
        ファイルシステムからJSONファイルを削除します。

        Args:
            session_id (str): 削除するセッションID
        """
        file_path = self._get_file_path(session_id)
        if file_path.exists():
            try:
                file_path.unlink()
            except Exception as e:
                print(f"Failed to delete session {session_id}: {e}")


session_store = SessionStore()
