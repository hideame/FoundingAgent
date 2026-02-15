# FoundingAgent

日本政策金融公庫の創業計画書作成を支援するAIエージェントアプリケーション

## 機能

- ✅ AI（Google Gemini）による対話形式の創業計画書作成支援
- ✅ 10項目に分けた段階的な計画書作成
- ✅ AI検証による創業計画書のフィードバック（[重要]/[確認]マーカー付き）
- ✅ 作成した計画書をExcel形式でダウンロード
- ✅ 業種別の記入例PDF付きZIPダウンロード
- ✅ 面談対策画面（ビジネスサポートプラザ案内）
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
- `GOOGLE_CLOUD_LOCATION`: 東京リージョン（`asia-northeast1`）はVertex AIで選択できるモデルの種類が少ないため、デフォルトリージョンの `us-central1` を使用します。

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

### 4. 初回データ初期化

アプリを一度起動してテーブルを作成してから、記入例データをシードします：

```bash
# テーブルを作成（SQLAlchemy create_all）
python -c "import asyncio; from app.database import init_db; asyncio.run(init_db())"

# 業種別の記入例データをDBに登録（初回のみ）
python scripts/fix_all_examples.py
```

### 5. アプリケーションの起動

```bash
# 開発サーバーを起動（ホットリロード有効）
uvicorn app.main:app --reload
```

ブラウザで `http://localhost:8000` にアクセスしてください。

## テスト

```bash
# 全テストを実行
pytest

# 詳細出力付き
pytest -v

# 特定のテストファイルのみ実行
pytest tests/test_excel_export.py -v
```

## データベース管理

### テーブル構成

- `sessions`: セッション管理
- `business_plans`: 創業計画書の9項目データ
- `chat_messages`: チャット履歴
- `task_states`: タスク進捗状態

### マイグレーション（Alembic）

> **注意:** 初回セットアップに `alembic upgrade head` は使用しないでください。
> 初回マイグレーションがテーブル作成ではなく `ALTER TABLE` のため、新規DBに実行するとエラーになります。
> テーブル作成は `create_all()`（アプリ起動時に自動実行）で行います。

以下のコマンドは、**モデル変更後に新しいマイグレーションファイルを作成・適用する**際に使用します：

```bash
# モデル変更後、マイグレーションファイルを自動生成
alembic revision --autogenerate -m "Add new column"

# マイグレーションを適用（既存DBへの差分適用）
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

## プロジェクト構成

```
FoundingAgent/
├── alembic/                       # データベースマイグレーション
│   ├── env.py
│   └── versions/                  # マイグレーションファイル
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
│   └── templates/
│       ├── index.html             # メインページ
│       └── components/            # コンポーネントテンプレート
│           ├── chat_interface.html
│           ├── interview_prep.html    # 面談対策画面
│           ├── plan_viewer.html
│           ├── plan_editor.html
│           └── stepper.html
├── document/                      # プロジェクトドキュメント
│   └── loan_flow.md              # 融資フロー説明
├── scripts/                       # ユーティリティスクリプト
│   └── fix_all_examples.py       # 記入例データシード
├── tests/                         # pytestテスト
│   ├── conftest.py
│   └── test_excel_export.py
├── .env.example                   # 環境変数テンプレート
├── CLAUDE.md                      # Claude Code用プロジェクト指示
├── docker-compose.yml             # MySQL環境定義
├── requirements.txt               # Python依存パッケージ
└── README.md
```

## 開発ガイド

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

# MySQLコンテナに直接接続（文字化けを防ぐためにUTF-8を指定）
docker-compose exec mysql mysql -u founding_user -pfounding_pass --default-character-set=utf8mb4 founding_agent
```

**注意:** 本番環境では、セキュリティ上の理由から環境変数やシークレット管理を使用してください。

### データベーステーブルが作成されない

アプリケーション起動時に自動作成されますが、手動で確認する場合：

```bash
# MySQLコンテナに接続
docker-compose exec mysql mysql -u founding_user -pfounding_pass founding_agent

# テーブル一覧を確認
SHOW TABLES;

# テーブル構造を確認
DESCRIBE business_plans;

# example_contentsテーブルの確認（日本語データ）
DESCRIBE example_contents;
SELECT COUNT(*) FROM example_contents;

# 特定の業種・セクションの記入例を確認
SELECT industry_type, section_key, SUBSTRING(example_text, 1, 200) as preview
FROM example_contents
WHERE industry_type='software' AND section_key='motivation'\G
```
