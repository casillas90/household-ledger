#!/bin/bash
# 개인 데이터(private/private_data.json)를 암호화해 공개 저장소용 private_data.enc.json 으로 만든다.
# - 키: private/data_key.txt (64자리 hex, 없으면 새로 생성). 같은 값을 Apps Script 스크립트 속성 DATA_KEY 에 넣어야 함.
# - private/ 폴더는 .gitignore 대상이라 GitHub 에 올라가지 않음.
set -euo pipefail
cd "$(dirname "$0")"

KEY_FILE="private/data_key.txt"
PLAIN="private/private_data.json"
OUT="private_data.enc.json"

[ -f "$PLAIN" ] || { echo "$PLAIN 이 없습니다." >&2; exit 1; }
if [ ! -f "$KEY_FILE" ]; then
  openssl rand -hex 32 > "$KEY_FILE"
  chmod 600 "$KEY_FILE"
  echo "새 데이터 키를 만들었습니다: $KEY_FILE (Apps Script 스크립트 속성 DATA_KEY 에 넣어 주세요)"
fi

KEY="$(tr -d '\n' < "$KEY_FILE")"
IV="$(openssl rand -hex 16)"
DATA="$(openssl enc -aes-256-cbc -K "$KEY" -iv "$IV" -in "$PLAIN" -a -A)"
printf '{"v":1,"alg":"AES-256-CBC","iv":"%s","data":"%s"}\n' "$IV" "$DATA" > "$OUT"
echo "암호화 완료: $OUT"
