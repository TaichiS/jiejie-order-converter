#!/bin/bash
cd "$(dirname "$0")"

uv run python3 -c "
import sqlite3
conn = sqlite3.connect('data/conversion.db')
conn.execute('DELETE FROM conversion_errors')
conn.execute('DELETE FROM conversion_logs')
conn.commit()
conn.close()
print('DB 清除完成')
"

find output -type f -delete
echo "output/ 清除完成"

find archive -type f -delete
echo "archive/ 清除完成"
