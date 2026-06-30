# Avikus Newsletter (Issue 01) — Outlook 메일 작성·발송 가이드

이 문서는 **Avikus Newsletter (Issue 01 / 창간호 "Who is Avikus?")** 를
**Outlook 본문에 디자인이 그대로 박힌 채로** 새 메일 창으로 띄우는 도구의 사용법입니다.
이미지는 첨부파일이 아니라 **본문 인라인(CID)** 으로 들어가, 클래식·신형 Outlook 모두에서 표시됩니다.

처음 쓰는 사람도 따라 할 수 있도록 **설치 → 실행 → 수정 → 발송 주의사항** 순서로 정리했습니다.

---

## 1. 폴더 구성

```
newsletter_new/
├─ README.md                          저장소 전체 안내(시작점)
├─ docs/                              사용법 문서 모음
│  ├─ 사용법_Avikus_Newsletter.md      이 문서(원본)
│  ├─ 사용법_Avikus_Newsletter.docx    이 문서의 Word 버전
│  └─ build_docx.py                    md → docx 변환기(문서 갱신용)
├─ issues/                            발송 호(號)별 자체완결 폴더
│  ├─ vol1_hinas/         ★ 현재 발송용 — HiNAS Newsletter Vol.1
│  │  ├─ run.bat                       더블클릭 실행
│  │  ├─ create_outlook_email.py       Outlook 메일 생성 스크립트
│  │  ├─ newsletter_email.html         메일 본문(디자인/글)
│  │  ├─ HiNAS_뉴스레터_사용가이드.docx  Vol.1 전용 가이드
│  │  ├─ README.md                     Vol.1 안내
│  │  └─ assets/                       본문 이미지(cid 연결)
│  └─ issue01_avikus/    Avikus Newsletter Issue 01 — "Who is Avikus?"
│     ├─ run.bat                       더블클릭 실행
│     ├─ create_outlook_email.py       Outlook 메일 생성 스크립트
│     ├─ newsletter_email.html         메일 본문(디자인/글)
│     ├─ 뉴스레터_일간호.docx            원본 기획 문서(글 출처)
│     └─ assets/                       본문 이미지(cid 연결)
│        ├─ logo_avikus_2.png    Avikus 로고 (cid:avikus_logo)
│        ├─ feat_nav2.png        HiNAS Navigation 화면 (cid:feat_nav2)
│        ├─ feat_svm1.png        HiNAS SVM 화면 (cid:feat_svm1)
│        ├─ feat_control1.png    HiNAS Control 화면 1 (cid:feat_control1)
│        ├─ feat_control2.png    HiNAS Control 화면 2 (cid:feat_control2)
│        ├─ feat_cloud.png       HiNAS Cloud 화면 (cid:feat_cloud)
│        ├─ _orig/               둥근 모서리 적용 전 원본 백업
│        └─ _unused/             현재 안 쓰는 이미지 보관
└─ _source/                           원본 디자인 소스(Monthly Newsletter.pptx) 보관
```

> 각 호 폴더는 **자체완결**입니다 — 그 폴더 안의 `create_outlook_email.py` · `newsletter_email.html` · `assets/` 만으로 독립 실행됩니다.
> 이미지는 HTML의 `cid:이름` 과 스크립트의 `IMAGES` 딕셔너리 키가 **이름으로 연결**됩니다. 둘의 이름이 같아야 본문에 표시됩니다.

---

## 2. 사전 준비 (최초 1회만)

### 2-1. Outlook 데스크톱 앱
- **Windows용 Outlook 데스크톱 앱**이 설치되어 있고 **로그인**되어 있어야 합니다.
- 웹 브라우저용 Outlook만으로는 동작하지 않습니다.

### 2-2. Python 3 설치
1. https://www.python.org/downloads/ 에서 Python 3 다운로드
2. 설치 **첫 화면에서 `Add python.exe to PATH` 체크박스를 반드시 켜고** 설치

설치 확인 — 명령 프롬프트(또는 PowerShell)에서:

```
py --version
```

버전이 나오면 정상입니다. (예: `Python 3.12.x`)

> `py` 가 안 되면 `python --version` 으로도 시도해 보세요.
> "Microsoft Store" 가 열리면 Python 이 안 깔린 상태입니다 → 위 1~2단계 다시 진행.

### 2-3. 필요한 패키지 설치 (pywin32)
명령 프롬프트에서 한 번만:

```
py -m pip install pywin32
```

> `py` 가 안 되면 `python -m pip install pywin32` 로 시도하세요.

---

## 3. 메일 만들기 (실행)

**가장 쉬운 방법 ⭐** — 발송할 호 폴더(예: `issues/vol1_hinas/` 또는 `issues/issue01_avikus/`)에서 **`run.bat` 더블클릭**.
→ Python 확인 → pywin32 자동 설치 → 메일 생성까지 한 번에 진행됩니다.

**명령줄로 직접 실행하려면:**

1. 발송할 호 폴더(예: `issues/issue01_avikus/`)의 주소창에 `cmd` 를 입력하고 Enter → 그 폴더에서 명령 프롬프트가 열립니다.
2. 아래 명령 실행:

```
py create_outlook_email.py
```

→ **뉴스레터가 본문에 박힌 채로** Outlook 새 메일 창이 열립니다.
→ **받는 사람 입력 → 보내기** 누르면 끝.

> 실행 시 아래처럼 인라인 삽입 로그가 뜨면 정상입니다:
> ```
>   inline: avikus_logo <- logo_avikus_2.png
>   inline: feat_nav2 <- feat_nav2.png
>   ...
> ```

---

## 4. 내용·디자인 수정

| 바꾸고 싶은 것 | 수정할 곳 |
|---|---|
| 글, 문구, 레이아웃 | 해당 호 폴더의 `newsletter_email.html` 편집 후 다시 실행 |
| 이미지 교체 | 해당 호 폴더의 `assets/` 안 **같은 이름** 파일로 교체 (자동 반영) |
| 새 이미지 추가 | HTML에 `<img src="cid:새이름">` 추가 + 스크립트 `IMAGES` 에 `"새이름": "파일.png"` 등록 |
| 제목 / 받는 사람 / 발송 방식 | 해당 호 폴더의 `create_outlook_email.py` (아래 5장) |

HTML은 메모장으로도 열 수 있지만, **VS Code** 같은 편집기를 권장합니다.

---

## 5. 제목·받는 사람·발송 방식 변경

해당 호 폴더의 `create_outlook_email.py` 파일을 편집기로 열어 아래 부분을 수정합니다.

### 제목

```python
mail.Subject = "Avikus Newsletter — Issue 01: Who is Avikus?"
```

### 받는 사람 자동 입력 (선택)
아래 줄의 맨 앞 `#` 을 지우고 주소를 넣습니다. 여러 명은 세미콜론으로 구분:

```python
mail.To = "a@avikus.ai; b@avikus.ai"
```

### 발송 방식
스크립트 맨 아래 한 줄:

```python
mail.Display(False)   # 메일 창만 열기 (기본, 권장)
# mail.Send()         # ← 이렇게 바꾸면 검토 없이 즉시 발송
```

> 처음에는 반드시 `Display` 로 **눈으로 확인 후 직접 보내기**를 권장합니다.

---

## 6. ⚠️ 발송 시 주의사항 (꼭 읽어주세요)

실수로 **옛 버전 메일이 같이 전송되는 사고**가 있었습니다. 원인과 예방법입니다.

### 왜 옛 메일이 같이 나갔나? (큐잉)
- Outlook 은 "보내기" 를 눌러도 즉시 전송하지 않고, 먼저 **보낼 편지함(Outbox)** 에 잠시 넣어둡니다(= 큐잉/대기).
- 예전 실행 때 띄워둔 **옛 초안 창**이 닫히지 않고 보낼 편지함에 남아 있으면,
  새 메일을 보낼 때 **대기 중이던 옛 메일까지 한 배치로 같이 전송**됩니다.

### 예방 수칙
1. 실행 **전후로 열려 있는 옛 메일 작성 창을 모두 닫기** (저장 안 함).
2. 새로 뜬 창에서 **받는 사람 입력 → 그 창에서만 보내기**.
3. 보내기 직후 **보낼 편지함(Outbox)이 비었는지** 확인.
   (안 비어 있으면 인터넷 연결/오프라인 설정 확인)

---

## 7. 임시보관함(Drafts) 초안 정리

스크립트를 여러 번 실행하면 **받는사람 없는 테스트 초안**이 임시보관함에 쌓일 수 있습니다.
실수 발송 위험을 줄이려면 주기적으로 정리하세요.

- Outlook **임시보관함** 열기 → 제목이 `Avikus Newsletter — Issue 01...` 이고
  **받는사람이 비어 있는** 초안들을 선택해 삭제.
- 삭제한 초안은 **지운 편지함**으로 가므로 복구도 가능합니다.

---

## 8. 잘못 보낸 메일 회수

사내(@avikus.ai) 동료에게 보냈다면 회수 가능성이 높습니다. **빠를수록 성공률↑**.

**클래식 Outlook**
1. **보낸 편지함** → 해당 메일 더블클릭으로 열기
2. **메시지** 탭 → **작업(Actions)** → **이 메시지 회수(Recall This Message)**
3. **"읽지 않은 복사본 삭제"** 선택 → 확인

**새 Outlook / 웹 버전**
1. **보낸 편지함** → 해당 메일 열기
2. 우측 상단 **··· (추가 작업)** → **메시지 회수(Recall message)**

> 상대가 **이미 읽었으면 회수는 실패**합니다.
> 이 경우 같은 분께 "착오 발송이니 무시 부탁드립니다" 한 줄 안내가 가장 확실합니다.

---

## 9. 문제 해결 (FAQ)

| 증상 | 해결 |
|---|---|
| `py` / `python` 입력 시 Microsoft Store 가 열림 | Python 미설치. python.org 에서 설치(PATH 체크). 그래도 뜨면 설정 ▸ 앱 ▸ 고급 앱 설정 ▸ 앱 실행 별칭 에서 `python.exe`/`python3.exe` 끄기 |
| `ModuleNotFoundError: win32com` | `py -m pip install pywin32` 실행 |
| 메일 창이 안 열림 | Outlook 데스크톱 앱이 실행·로그인 상태인지 확인 |
| 이미지가 안 보임 | `assets/` 에 해당 PNG 가 있고, 파일명이 스크립트 `IMAGES` 키와 일치하는지 확인 |
| 사진 모서리가 각져 보임 | Outlook 데스크톱은 CSS 둥근 모서리를 무시함 → 이미지 파일 자체에 둥근 모서리를 입혀야 함(이미 적용됨, 원본은 `issues/issue01_avikus/assets/_orig/`) |
| 한글 콘솔 로그가 깨짐 | 표시상의 문제일 뿐 메일 생성에는 영향 없음 |

---

## 10. 이미지 가공 참고 (선택)

제품 사진은 보기 좋게 두 가지 가공이 되어 있습니다. (다시 할 때만 참고)

- **둥근 모서리**: Outlook 데스크톱은 CSS `border-radius` 를 무시하므로,
  PNG 파일 자체의 모서리를 투명하게 깎아 둥글게 만들었습니다.
- **흰 배경 크롭**: NAV 사진(`feat_nav2.png`)은 아래쪽 흰 여백을 잘라냈습니다.
- 가공 전 **원본은 `issues/issue01_avikus/assets/_orig/`** 에 백업되어 있습니다.
- 이미지 가공에는 추가 패키지가 필요합니다: `py -m pip install pillow`

---

*이 문서를 Word(.docx)로 다시 만들려면: `docs/` 폴더에서 `py build_docx.py`*
