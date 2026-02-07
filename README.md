# FoundingAgent

日本政策金融公庫の創業計画書作成を支援するAIエージェントアプリケーション

## 機能

- ✅ AI（Google Gemini）による対話形式の創業計画書作成支援
- ✅ 9つのセクションに分けた段階的な計画書作成
- ✅ 作成した計画書をExcel形式でダウンロード
- ✅ 業種別の記入例PDF付きZIPダウンロード
- ✅ MySQLデータベースによるセッション管理

## 技術スタック

- **Backend**: FastAPI (Python 3.12)
- **Database**: MySQL 8.0 + SQLAlchemy (非同期ORM)
- **AI**: Google Gemini API
- **Frontend**: Jinja2 Templates + HTML/CSS/JavaScript
- **Excel**: openpyxl
- **Testing**: pytest

## セットアップ

### 1. 環境変数の設定

`.env.example`を`.env`にコピーして、必要な値を設定します：

```bash
cp .env.example .env
```

`.env`ファイルを編集：

```env
# Google Gemini API Key（必須）
GOOGLE_API_KEY=your_actual_google_api_key_here

# データベース接続URL
DATABASE_URL=mysql+aiomysql://founding_user:founding_pass@localhost:3306/founding_agent

# デバッグモード
DEBUG=True
```

### 2. MySQLデータベースの起動

Docker Composeを使用してMySQLを起動します：

```bash
# MySQLコンテナを起動
docker-compose up -d

# 起動確認
docker-compose ps
```

### 3. Python仮想環境のセットアップ

```bash
# 仮想環境の作成
python3.12 -m venv .venv

# 仮想環境の有効化（macOS/Linux）
source .venv/bin/activate

# 仮想環境の有効化（Windows）
.venv\Scripts\activate

# 依存パッケージのインストール
pip install -r requirements.txt
```

### 4. アプリケーションの起動

```bash
# 開発サーバーを起動（ホットリロード有効）
uvicorn app.main:app --reload
```

ブラウザで `http://localhost:8000` にアクセスしてください。

## データベース管理

### テーブル構成

- `sessions`: セッション管理
- `business_plans`: 創業計画書の9項目データ
- `chat_messages`: チャット履歴
- `task_states`: タスク進捗状態

### マイグレーション（Alembic）

```bash
# 初回マイグレーション環境の初期化
alembic init alembic

# マイグレーションファイルの自動生成
alembic revision --autogenerate -m "Initial migration"

# マイグレーション実行
alembic upgrade head

# ロールバック
alembic downgrade -1
```

### データベースのリセット

```bash
# Docker Composeでデータベースを完全にリセット
docker-compose down -v
docker-compose up -d
```

## テスト

```bash
# 全テストを実行
pytest

# 詳細出力付き
pytest -v

# 特定のテストファイルのみ実行
pytest tests/test_excel_export.py -v
```

## ドキュメント生成

```bash
# APIドキュメントを生成（pdoc使用）
pdoc app -o docs

# ドキュメントをブラウザでプレビュー
pdoc app
```

## プロジェクト構成

```
FoundingAgent/
├── app/
│   ├── main.py                    # FastAPIアプリケーション
│   ├── models.py                  # SQLAlchemyモデル
│   ├── database.py                # データベース接続設定
│   ├── store.py                   # セッションストア（DB操作）
│   ├── services/
│   │   └── gemini_service.py      # Gemini API統合
│   ├── static/
│   │   └── templates/
│   │       ├── startup_plan_template.xlsx
│   │       └── examples/          # 業種別記入例PDF
│   └── templates/                 # Jinja2テンプレート
├── tests/                         # pytestテスト
├── data/                          # JSONセッションデータ（旧）
├── docker-compose.yml             # MySQL環境定義
├── requirements.txt               # Python依存パッケージ
├── .env.example                   # 環境変数テンプレート
└── README.md
```

## 開発ガイド

### VS Code タスク

`.vscode/tasks.json`に定義されたタスク：

- **Generate Docs**: pdocでドキュメント生成
- **Preview Docs**: pdocドキュメントをブラウザでプレビュー

`Cmd+Shift+P` → "Tasks: Run Task"から実行可能

### コーディング規約

- Python 3.12の型ヒントを積極的に使用
- docstringはGoogle形式で記述
- 非同期処理には`async/await`を使用
- SQLAlchemyの非同期APIを使用

## トラブルシューティング

### MySQLに接続できない

```bash
# MySQLコンテナのログを確認
docker-compose logs mysql

# MySQLコンテナに直接接続
docker-compose exec mysql mysql -u founding_user -p founding_agent
# パスワード: founding_pass
```

### データベーステーブルが作成されない

アプリケーション起動時に自動作成されますが、手動で確認する場合：

```bash
# MySQLコンテナに接続
docker-compose exec mysql mysql -u founding_user -p founding_agent

# テーブル一覧を確認
SHOW TABLES;

# テーブル構造を確認
DESCRIBE business_plans;
```

## ライセンス

MIT License

## 作者

hide
