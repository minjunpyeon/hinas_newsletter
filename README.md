# Avikus Newsletter

뉴스레터를 **Outlook 본문에 디자인 그대로(이미지 인라인 CID) 박은 채로** 새 메일 창으로 띄우는 도구 모음입니다.
이미지는 첨부파일이 아니라 본문에 인라인으로 들어가서, 클래식·신형 Outlook 모두에서 표시됩니다. (Windows 전용)

## 📁 구조

```
newsletter_new/
├─ docs/                    사용법 가이드 (시작 전 읽기)
│  ├─ 사용법_Avikus_Newsletter.md     ← 전체 사용법 (원본)
│  ├─ 사용법_Avikus_Newsletter.docx   ← Word 버전
│  └─ build_docx.py                   md → docx 변환기
├─ issues/                  발송 호(號)별 자체완결 폴더
│  ├─ vol1_hinas/      ★ 현재 발송용 — HiNAS Newsletter Vol.1
│  └─ issue01_avikus/    Avikus Newsletter Issue 01 — "Who is Avikus?"
└─ _source/                 원본 디자인 소스(.pptx) 보관
```

각 호 폴더는 **자체완결**입니다 — 그 폴더 안의 `create_outlook_email.py` · `newsletter_email.html` · `assets/` 만으로 독립 실행됩니다.

## 🚀 발송 방법 (가장 쉬운 방법)

발송할 호 폴더로 들어가 **`run.bat` 더블클릭**.
→ Python 확인 → 필요한 패키지(pywin32) 자동 설치 → Outlook 새 메일 창이 **본문에 박힌 채로** 열림.
→ 받는 사람 입력 → **보내기**.

| 보낼 것 | 폴더 |
|---|---|
| **HiNAS Newsletter Vol.1** (현재 발송용) | `issues/vol1_hinas/` |
| Avikus Newsletter Issue 01 | `issues/issue01_avikus/` |

> 명령줄로 직접 실행하려면 해당 호 폴더에서: `py create_outlook_email.py`

### 📨 대량 발송 (선사 등 다수 수신자 BCC 한 통)

1. 수신자 엑셀을 **CSV(UTF-8)로 저장**해 [`recipients/`](recipients/) 폴더에 넣기
   (열 위치·제목 무관 — 이메일 형태면 자동 추출, 중복 제거)
2. 발송할 호 폴더에서: `py send_bulk.py`
3. 전원이 **BCC**에 들어간 메일 1통이 열림 → **인원수 확인 후 보내기**

자세한 규칙은 [`recipients/README.md`](recipients/README.md) 참고.
실제 수신자 목록은 개인정보라 git에 커밋되지 않아(`.gitignore` 제외).

## 📖 자세한 사용법

- 전체 가이드: [`docs/사용법_Avikus_Newsletter.md`](docs/사용법_Avikus_Newsletter.md)
  (설치 → 실행 → 내용/디자인 수정 → ⚠️ 발송 주의사항 → 메일 회수 → FAQ)
- Vol.1 전용 안내: [`issues/vol1_hinas/README.md`](issues/vol1_hinas/README.md)

## ✏️ 내용·디자인 수정 (요약)

| 바꾸고 싶은 것 | 수정할 곳 |
|---|---|
| 글·문구·레이아웃 | 해당 호 폴더의 `newsletter_email.html` |
| 이미지 교체 | 해당 호 폴더의 `assets/` 안 **같은 이름** 파일로 교체 |
| 제목 / 받는 사람 / 발송 방식 | 해당 호 폴더의 `create_outlook_email.py` |

HTML의 `cid:이름` 과 스크립트 `IMAGES` 딕셔너리의 키가 **이름으로 연결**됩니다.

## ⚠️ 사전 준비 (최초 1회)

1. **Outlook 데스크톱 앱** 설치·로그인 (웹 Outlook 단독으로는 동작 안 함)
2. **Python 3** 설치 — 설치 첫 화면에서 `Add python.exe to PATH` 체크
   (`pywin32` 패키지는 `run.bat`이 자동 설치)
