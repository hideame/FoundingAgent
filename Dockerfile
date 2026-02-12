FROM python:3.12-slim

WORKDIR /app

# 依存関係インストール（キャッシュ活用のため先にコピー）
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# アプリケーションコードをコピー
COPY . .

# 起動スクリプトに実行権限を付与
RUN chmod +x startup.sh

EXPOSE 8080

CMD ["./startup.sh"]
