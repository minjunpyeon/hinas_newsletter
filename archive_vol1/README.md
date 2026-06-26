# HiNAS Newsletter → Outlook 메일 자동 생성

PPT(암호화된 구형 포맷)에서 추출한 이미지로 만든 뉴스레터를
**Outlook 본문에 인라인(CID)으로 박아서** 새 메일 창을 여는 도구입니다.
이미지가 첨부파일로 보이지 않고, 클래식·신형 Outlook 모두에서 표시됩니다.

---

## 📁 폴더 구성

```
archive_vol1/
├─ run.bat                          # ★ 더블클릭하면 한 번에 실행 (Python 체크+설치+실행)
├─ newsletter_email.html            # 이메일 본문 HTML (글/디자인 수정은 여기)
├─ create_outlook_email.py          # Outlook 메일 자동 생성 스크립트
├─ README.md                        # 이 문서
├─ HiNAS_뉴스레터_사용가이드.docx   # 사용 가이드 (Word)
└─ assets/                          # 본문에 들어가는 이미지 (cid:이름 으로 연결)
   ├─ feat_hd.png         # HiNAS 헤더 이미지 (cid:logo_hd)
   ├─ avikus_wordmark.png # Avikus 로고 (cid:logo_avikus)
   ├─ feat_control1.png   # HiNAS Control 화면 — Optimization (cid:feat_control1)
   ├─ feat_control2.png   # HiNAS Control 화면 — HiNAS Control (cid:feat_control2)
   ├─ badge_1.png         # 배지 1 (cid:badge_1)
   ├─ badge_2.png         # 배지 2 (cid:badge_2)
   └─ badge_3.png         # 배지 3 (cid:badge_3)
```

---

## ✅ 사전 준비 (최초 1회)

1. **Outlook 데스크톱 앱**이 설치·로그인되어 있어야 합니다. (Windows 전용)
2. **Python 3** 설치 — 아직 없다면:
   - https://www.python.org/downloads/ 에서 다운로드
   - 설치 **첫 화면에서 `Add python.exe to PATH` 체크박스를 반드시 켜고** 설치
   - `pywin32` 패키지는 `run.bat`이 **자동으로 설치**하므로 따로 안 해도 됩니다.

> 💡 Python이 이미 설치돼 있는지 모르겠으면 그냥 `run.bat`을 더블클릭해 보세요.
> 없으면 설치 안내 메시지가 뜹니다.

> 이미지 가공(추출/선명화)을 다시 할 때만 `Pillow`, `olefile`이 추가로 필요합니다.
> 메일 생성만 할 거면 위 준비만으로 충분합니다.

---

## 🚀 사용법

### 1) 메일 만들기 (가장 쉬운 방법 ⭐)

압축을 푼 폴더에서 **`run.bat`을 더블클릭**하세요.

→ Python 확인 → 필요한 패키지(pywin32) 자동 설치 → 메일 생성까지 한 번에 진행됩니다.
→ 뉴스레터가 **본문에 박힌 채로** Outlook 새 메일 창이 열립니다.
받는 사람 입력 → **보내기** 누르면 끝.

> 명령줄로 직접 실행하고 싶다면 폴더에서:
> ```bash
> python create_outlook_email.py
> ```

### 2) 받는 사람 자동 입력

`create_outlook_email.py`에서 아래 줄의 주석(`#`)을 풀고 주소 입력:

```python
# mail.To = "someone@example.com"
```

여러 명은 세미콜론으로 구분: `"a@x.com; b@y.com"`

### 3) 검토 없이 바로 발송

스크립트 맨 아래 한 줄만 변경:

```python
mail.Display(False)   # 메일 창 열기 (기본)
mail.Send()           # ← 이렇게 바꾸면 즉시 발송
```

> 처음에는 `Display`로 확인 후 보내는 것을 권장합니다.

### 4) 제목 변경

```python
mail.Subject = "HiNAS Newsletter — Vol.1"   # 원하는 제목으로 수정
```

---

## ✏️ 내용·디자인 수정

| 바꾸고 싶은 것 | 수정할 곳 |
|---|---|
| 글, 문구, 레이아웃 | `newsletter_email.html` 편집 후 스크립트 재실행 |
| 이미지 교체 | `assets/` 안의 **같은 이름** 파일로 교체 (자동 반영) |
| 새 이미지 추가 | HTML에 `<img src="cid:새이름">` 추가 + 스크립트 `IMAGES` 딕셔너리에 `"새이름": "파일.png"` 등록 |
| 제목 / 받는 사람 / 발송 방식 | `create_outlook_email.py` 상단·하단 |

이미지는 HTML의 `cid:이름` 과 스크립트 `IMAGES`의 키가 **이름으로 연결**됩니다.

---

## ⚠️ 참고 / 제약

- **이메일은 자바스크립트·hover·애니메이션이 대부분 차단**됩니다.
  - "동적" 요소는 클릭 버튼(Contact our team)·mailto 링크로 구현.
  - hover 확대 효과는 신형 Outlook/웹에서만 작동(데스크톱은 자연 무시).
- 전체 메일 용량은 약 **1.3MB** — 전송에 문제없습니다.
- 원본 PPT는 **민감도 레이블/IRM으로 암호화**되어 있어 이미지를 직접 추출할 수 없습니다.
  이미지는 PowerPoint로 슬라이드를 고해상도 렌더링한 뒤 좌표 기준으로 잘라낸 것입니다.

---

## 🛠️ 문제 해결

| 증상 | 해결 |
|---|---|
| `python` 입력 시 **Microsoft Store가 열림** / 아무 반응 없음 | Python 미설치 또는 가짜 별칭 때문. python.org에서 설치(PATH 체크). 그래도 Store가 뜨면 **설정 ▸ 앱 ▸ 고급 앱 설정 ▸ 앱 실행 별칭**에서 `python.exe`/`python3.exe`를 끄세요. |
| `pip`가 없다고 나옴 | Python이 안 깔린 상태. python.org에서 설치(PATH 체크) 후 `run.bat` 재실행. `pip` 대신 `python -m pip ...` 사용 권장. |
| `run.bat`이 창만 깜빡이고 닫힘 | 폴더 경로에 문제 없음 — 메시지를 보려면 cmd 창에서 직접 `run.bat` 실행하거나, 표시되는 안내를 따르세요. |
| `ModuleNotFoundError: win32com` | `run.bat`이 자동 설치합니다. 수동은 `python -m pip install pywin32` |
| 메일 창이 안 열림 | Outlook 데스크톱 앱이 실행/로그인 상태인지 확인 |
| 이미지가 안 보임 | `assets/` 폴더에 PNG 3개가 있는지, 파일명이 스크립트 `IMAGES`와 일치하는지 확인 |
| 한글이 깨져 보임(콘솔 로그) | 표시상의 문제일 뿐, 메일 생성에는 영향 없음 |
