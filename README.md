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
- **AI**: Google Gemini API via **Vertex AI**
- **Frontend**: Jinja2 Templates + HTML/CSS/JavaScript
- **Excel**: openpyxl
- **Testing**: pytest
- **Cloud**: Google Cloud Platform (Vertex AI)

## セットアップ

### 1. Google Cloud / Vertex AI のセットアップ

このプロジェクトは **Vertex AI** を使用してGemini APIにアクセスします。

#### 1.1 Google Cloud プロジェクトの作成

1. [Google Cloud Console](https://console.cloud.google.com/) にアクセス
2. 新しいプロジェクトを作成（または既存のプロジェクトを使用）
3. プロジェクトIDをメモしておく

#### 1.2 Vertex AI APIの有効化

```bash
# Google Cloud CLIをインストール（未インストールの場合）
# https://cloud.google.com/sdk/docs/install

# 認証
gcloud auth login

# プロジェクトを設定
gcloud config set project YOUR_PROJECT_ID

# Vertex AI APIを有効化
gcloud services enable aiplatform.googleapis.com
```

#### 1.3 Application Default Credentials (ADC) の設定

```bash
# ローカル開発用の認証情報を設定
gcloud auth application-default login
```

ブラウザが開くので、Google アカウントでログインしてください。

#### 1.4 環境変数の設定

`.env.example`を`.env`にコピーして、必要な値を設定します：

```bash
cp .env.example .env
```

`.env`ファイルを編集：

```env
# Google Cloud / Vertex AI（必須）
GOOGLE_CLOUD_PROJECT=your-project-id-here
GOOGLE_CLOUD_LOCATION=us-central1

# データベース接続URL
DATABASE_URL=mysql+aiomysql://founding_user:founding_pass@localhost:3306/founding_agent

# デバッグモード
DEBUG=True
```

**注意:**
- `GOOGLE_CLOUD_PROJECT`: 1.1で作成したプロジェクトID
- `GOOGLE_CLOUD_LOCATION`: リージョン（デフォルトは `us-central1`、東京リージョンは `asia-northeast1`）

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

## API ドキュメント（Swagger UI）

FastAPIには自動生成されるAPI ドキュメントが組み込まれています。

### Swagger UI（インタラクティブなAPI ドキュメント）

アプリケーション起動後、以下のURLにアクセスしてください：

```
http://localhost:8000/docs
```

**主な機能:**
- 📋 全APIエンドポイントの一覧表示
- 🔍 各エンドポイントのリクエスト/レスポンススキーマ確認
- 🧪 ブラウザから直接APIをテスト（Try it out機能）
- 📥 リクエストボディのサンプル表示
- 📤 レスポンスの例とステータスコード

### ReDoc（読みやすいAPI ドキュメント）

より読みやすいドキュメント形式が必要な場合：

```
http://localhost:8000/redoc
```

**特徴:**
- 📖 3カラムレイアウトで見やすい
- 🔎 検索機能付き
- 📋 エンドポイントのグループ化表示

### OpenAPI スキーマ（JSON）

APIスキーマを直接取得する場合：

```
http://localhost:8000/openapi.json
```

このJSONファイルは、Postmanなどの外部ツールへのインポートにも使用できます。

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

## ドキュメント

このプロジェクトには2種類のドキュメントがあります：

### 1. API ドキュメント（Swagger UI / ReDoc）

**用途**: REST APIエンドポイントの仕様確認とテスト

- **Swagger UI**: `http://localhost:8000/docs`（インタラクティブ）
- **ReDoc**: `http://localhost:8000/redoc`（読みやすい）
- **OpenAPI JSON**: `http://localhost:8000/openapi.json`（スキーマファイル）

FastAPIが自動生成するため、追加の設定は不要です。

### 2. コードドキュメント（pdoc）

**用途**: Pythonコードの内部実装、クラス、関数の詳細説明

```bash
# コードドキュメントを生成（静的HTML）
pdoc app -o docs

# ドキュメントをブラウザでプレビュー（開発サーバー起動）
pdoc app
```

生成されたHTMLは `docs/` ディレクトリに保存されます。

> **💡 使い分け**:
> - **API利用者向け** → Swagger UI / ReDoc
> - **開発者向け（コード理解）** → pdoc

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

### Vertex AI / Gemini APIのエラー

#### 「チャットの開始に失敗しました」と表示される

**原因:** Vertex AIの認証情報が設定されていない、またはAPIが有効化されていない

**解決方法:**

1. **環境変数の確認**
   ```bash
   # .envファイルを確認
   cat .env

   # 以下が設定されているか確認
   # GOOGLE_CLOUD_PROJECT=your-project-id
   # GOOGLE_CLOUD_LOCATION=us-central1
   ```

2. **Application Default Credentials (ADC) の設定**
   ```bash
   # 認証情報を設定
   gcloud auth application-default login

   # 認証状態を確認
   gcloud auth application-default print-access-token
   ```

3. **Vertex AI APIの有効化確認**
   ```bash
   # APIが有効化されているか確認
   gcloud services list --enabled | grep aiplatform

   # 有効化されていない場合
   gcloud services enable aiplatform.googleapis.com
   ```

4. **プロジェクトIDの確認**
   ```bash
   # 現在のプロジェクトを確認
   gcloud config get-value project

   # プロジェクト一覧を表示
   gcloud projects list
   ```

#### Google Cloud の無料枠について

- 新規ユーザーは **$300の無料クレジット**（90日間有効）
- 詳細: https://cloud.google.com/free/docs/gcp-free-tier/#free-trial

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
