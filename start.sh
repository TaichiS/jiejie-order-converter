#!/bin/bash

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
URL="http://127.0.0.1:5099"
PORT=5099

cd "$SCRIPT_DIR"

# 砍掉佔用 port 的舊有程序
OLD_PID=$(lsof -ti tcp:$PORT 2>/dev/null)
if [ -n "$OLD_PID" ]; then
    echo "關閉舊有程序 (PID: $OLD_PID)..."
    kill $OLD_PID
    sleep 0.5
fi

# 背景啟動 server
uv run python main.py &
SERVER_PID=$!

# 等待 server 就緒
echo "啟動中..."
for i in $(seq 1 20); do
    if curl -s "$URL" > /dev/null 2>&1; then
        break
    fi
    sleep 0.5
done

# 開啟瀏覽器
open "$URL"

echo "Server 已啟動 (PID: $SERVER_PID)"
echo "按 Ctrl+C 關閉 server"

# 等待 server 結束
wait $SERVER_PID
