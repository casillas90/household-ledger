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
