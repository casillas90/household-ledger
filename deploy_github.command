#!/bin/bash
# 다정 & 선준 가계부 대시보드 - GitHub Pages 자동 배포 스크립트
cd "$(dirname "$0")"

GH_BIN="$HOME/bin/gh"
if [ ! -f "$GH_BIN" ]; then
    GH_BIN="gh"
fi

clear
echo "============================================================"
echo "  🚀 다정 & 선준 가계부 대시보드 - GitHub Pages 자동 배포기"
echo "============================================================"
echo ""

# 1. GitHub CLI 로그인 상태 확인
if ! "$GH_BIN" auth status >/dev/null 2>&1; then
    echo "🔑 GitHub 로그인이 필요합니다. 브라우저 인증을 시작합니다..."
    echo "   (화면에 표시되는 영문 8자리 코드를 복사하고 웹 브라우저에서 승인해 주세요)"
    echo ""
    "$GH_BIN" auth login -w -p https -h github.com
fi

# 2. 로그인된 사용자 이름 조회
USERNAME=$("$GH_BIN" api user -q .login 2>/dev/null)
if [ -z "$USERNAME" ]; then
    echo "❌ GitHub 로그인 정보를 가져오지 못했습니다. 다시 실행해 주세요."
    read -p "엔터를 누르면 종료합니다..."
    exit 1
fi

echo ""
echo "✅ GitHub 계정 확인 완료: $USERNAME"
REPO_NAME="household-ledger"
echo "📦 GitHub 저장소 설정 중 ($USERNAME/$REPO_NAME)..."

# 3. Git 브랜치 확인 및 코드 푸시
git branch -M main 2>/dev/null || true
git add .
git commit -m "feat: GitHub Pages 배포" 2>/dev/null || true

if "$GH_BIN" repo view "$USERNAME/$REPO_NAME" >/dev/null 2>&1; then
    echo "ℹ️  기존 저장소가 존재합니다. 최신 코드를 푸시합니다..."
    git remote remove origin 2>/dev/null || true
    git remote add origin "https://github.com/$USERNAME/$REPO_NAME.git"
    git push -u origin main
else
    echo "✨ GitHub에 새 저장소($REPO_NAME)를 생성하고 코드를 푸시합니다..."
    "$GH_BIN" repo create "$REPO_NAME" --public --source=. --remote=origin --push
fi

# 4. GitHub Pages 활성화
echo "🌐 GitHub Pages 웹 호스팅 켜는 중..."
"$GH_BIN" api "repos/$USERNAME/$REPO_NAME/pages" -X POST -f source='{"branch":"main","path":"/"}' 2>/dev/null || true

PAGES_URL="https://$USERNAME.github.io/$REPO_NAME/"

echo ""
echo "============================================================"
echo "  🎉 축하합니다! GitHub 웹링크 배포 설정이 완료되었습니다!"
echo "============================================================"
echo ""
echo "  🔗 아이폰 & 외부 접속 영구 웹링크:"
echo "     👉 $PAGES_URL"
echo ""
echo "  💡 유용한 안내:"
echo "  1. 이제 맥북을 꺼두어도, 아이폰이 Wi-Fi가 아닌 LTE/5G 데이터 상태여도"
echo "     언제 어디서든 위 링크로 접속할 수 있습니다!"
echo "  2. 아이폰 Safari에서 위 링크 접속 후 [공유 ⎋] ➔ [홈 화면에 추가 ➕]를"
echo "     누르면 진짜 금융 앱처럼 깔끔하게 홈 화면에 설치됩니다."
echo "  (GitHub Pages 최초 반영에 약 1~2분 정도 소요될 수 있습니다)"
echo "============================================================"
echo ""
read -p "완료되었습니다. 창을 닫으려면 [Enter] 키를 누르세요..."
