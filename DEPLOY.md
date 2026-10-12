# 📱 아이폰 & 외부 접속을 위한 무료 웹 배포 가이드

다정 & 선준 가계부 대시보드는 백엔드 서버 없이 동작하는 순수 정적 웹 애플리케이션으로, **GitHub Pages**나 **Vercel**을 통해 **100% 평생 무료**로 고유 인터넷 주소(URL)를 가질 수 있습니다.

---

## 🎉 배포 완료! 현재 운영 중인 접속 링크

👉 **실시간 웹링크: [https://casillas90.github.io/household-ledger/](https://casillas90.github.io/household-ledger/)**

- 이제 맥북을 꺼두어도, 아이폰이 LTE/5G 데이터 상태여도 언제 어디서든 접속할 수 있습니다.
- 카카오톡이나 메시지로 다정님께 위 링크를 공유하시면 둘 다 언제든 확인하실 수 있습니다.

---

## 🛠 수동으로 직접 배포하고 싶으실 때

### 1단계: GitHub 새 저장소 만들기
1. [GitHub.com](https://github.com)에 로그인합니다.
2. 우측 상단 **[+] ➔ [New repository]**를 클릭합니다.
3. Repository name에 `household-ledger`를 입력합니다.
4. **Public**을 선택하고 **[Create repository]**를 클릭합니다.

### 2단계: 맥북 터미널에서 코드 올리기
맥북 터미널을 열고 아래 명령어를 실행합니다:

```bash
cd "/Users/swdjsj/Desktop/가계부대시보드"

# GitHub 원격 저장소 연결 (내 아이디로 변경)
git branch -M main
git remote add origin https://github.com/<내GitHub아이디>/household-ledger.git
git push -u origin main
```

### 3단계: GitHub Pages 켜기 (10초 소요)
1. 생성한 GitHub 저장소 화면 상단의 **[Settings]** 탭 클릭
2. 좌측 메뉴에서 **[Pages]** 클릭
3. **Build and deployment** 항목의 **Branch**를 `main` 브랜치, 폴더는 `/ (root)`로 선택 후 **[Save]** 클릭
4. 약 1분 후 상단에 생성된 나만의 접속 링크가 나타납니다:
   👉 **`https://<내깃허브아이디>.github.io/household-ledger/`**

---

## ⚡️ 방법 2: Vercel로 배포 (초고속 드래그 앤 드롭)

GitHub 계정이 없거나 복잡한 설정이 싫다면 Vercel을 사용할 수 있습니다.

1. [Vercel.com](https://vercel.com) 회원가입 (무료)
2. **Add New... ➔ Project** 클릭
3. 맥북의 `가계부대시보드` 폴더를 화면에 드래그 앤 드롭
4. **Deploy** 버튼을 누르면 즉시 고유 도메인(예: `https://my-ledger-xxx.vercel.app`)이 발급됩니다.

---

## 📲 아이폰에서 가계부 앱으로 등록하기

웹 링크가 생성되면 아이폰에서 접속한 뒤:
1. 아이폰 Safari로 해당 주소에 접속합니다.
2. 화면 하단 중앙의 **[공유 버튼 ⎋]** (네모 위 화살표) 탭
3. **[홈 화면에 추가 ➕]** 탭
4. 우측 상단 **[추가]**를 누르면, 홈 화면에 예쁜 전용 가계부 앱 아이콘이 생깁니다!

---

## 🔐 Apps Script 업데이트 & 접근 키 설정 (v3.7)

### 1. 새 코드 반영 (배포 주소는 그대로 유지)
1. 구글 시트 ➔ **[확장 프로그램] ➔ [Apps Script]**
2. 기존 코드를 모두 지우고 `google_apps_script.js` 내용을 붙여넣은 뒤 저장
3. **[배포] ➔ [배포 관리]** ➔ 기존 웹 앱 배포의 ✏️(편집) ➔ 버전: **새 버전** ➔ **[배포]**
   - ⚠️ [새 배포]를 누르면 주소가 바뀌어 `index.html`의 `APPS_SCRIPT_URL`도 바꿔야 합니다.
4. 사이트를 새로고침했을 때 상단 "데이터 점검"에 **"Apps Script가 최신 버전이 아님"**이 사라지면 완료

### 2. 접근 키 설정 (선택, 권장)
1. Apps Script 편집기 왼쪽 ⚙️ **[프로젝트 설정]** ➔ 맨 아래 **[스크립트 속성] ➔ [속성 추가]**
2. 속성: `DASHBOARD_KEY`, 값: 원하는 비밀 문구 (예: 길고 추측하기 어려운 문장) ➔ 저장
3. 사이트를 열면 접근 키를 묻는 창이 한 번 뜹니다. 같은 값을 입력하면 그 기기(브라우저)에 기억됩니다.
   - 아이폰·맥북 등 기기마다 한 번씩 입력
   - 키를 바꾸면 다음 접속 때 다시 묻습니다.

키를 설정하면 시트 실시간 데이터는 키가 있어야만 볼 수 있습니다.
단, `index.html` 안에 직접 들어 있는 기본(과거) 데이터와 GitHub 저장소 기록은 키와 무관하게 공개 상태입니다.

### 3. 진단용 시트 덤프
헤더 인식 문제를 확인할 때만 사용: `APPS_SCRIPT_URL?key=<키>&debug=1`
(접근 키가 설정된 경우에만 응답에 `debugSheets`가 포함됩니다)

### 4. 테스트 실행
```bash
osascript -l JavaScript test_apps_script.js && osascript -l JavaScript test_validation.js
```
