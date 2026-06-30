# -*- coding: utf-8 -*-
"""
HiNAS Newsletter -> Outlook 본문(HTML) + 이미지 CID 인라인 삽입.
실행하면 Outlook 새 메일 창이 '본문에 뉴스레터가 들어간 채로' 열립니다.
(자동 발송 X — 검토 후 직접 '보내기'를 누르세요.)
"""
import os
import win32com.client

BASE = os.path.dirname(os.path.abspath(__file__))
HTML_PATH = os.path.join(BASE, "newsletter_email.html")

# 이미지(assets) 폴더 탐색:
#   1) 같은 폴더의 assets/  (독립 실행 — 권장)
#   2) 상위 폴더의 assets/  (구버전 호환)
def _find_assets():
    for cand in (os.path.join(BASE, "assets"),
                 os.path.join(os.path.dirname(BASE), "assets")):
        if os.path.isdir(cand):
            return cand
    # 못 찾으면 기본값(같은 폴더 assets)을 반환 — 아래에서 경고 출력
    return os.path.join(BASE, "assets")

ASSETS = _find_assets()

# HTML 안의 cid:이름  ->  실제 이미지 파일
IMAGES = {
    "logo_hd":       "feat_hd.png",
    "logo_avikus":   "avikus_wordmark.png",
    "feat_control1": "feat_control1.png",
    "feat_control2": "feat_control2.png",
    "badge_1":       "badge_1.png",
    "badge_2":       "badge_2.png",
    "badge_3":       "badge_3.png",
}

# MAPI property tags
PR_ATTACH_CONTENT_ID  = "http://schemas.microsoft.com/mapi/proptag/0x3712001F"
PR_ATTACH_MIME_TAG    = "http://schemas.microsoft.com/mapi/proptag/0x370E001F"
PR_ATTACHMENT_HIDDEN  = "http://schemas.microsoft.com/mapi/proptag/0x7FFE000B"

# 메일 제목 (대량 발송 send_bulk.py 도 이 값을 그대로 사용)
SUBJECT = "HiNAS Newsletter — Vol.1"


def load_html():
    """본문 HTML 을 읽어 문자열로 반환."""
    with open(HTML_PATH, "r", encoding="utf-8") as f:
        return f.read()


def attach_inline_images(mail):
    """IMAGES 딕셔너리대로 mail 에 CID 인라인 이미지를 첨부. 삽입한 개수 반환."""
    count = 0
    for cid, fname in IMAGES.items():
        path = os.path.join(ASSETS, fname)
        if not os.path.exists(path):
            print("  (건너뜀, 파일 없음):", path)
            continue
        att = mail.Attachments.Add(path, 1)  # 1 = olByValue
        pa = att.PropertyAccessor
        pa.SetProperty(PR_ATTACH_CONTENT_ID, cid)      # cid: 참조와 연결
        pa.SetProperty(PR_ATTACH_MIME_TAG, "image/png")
        try:
            pa.SetProperty(PR_ATTACHMENT_HIDDEN, True) # 첨부 목록에서 숨김(인라인 전용)
        except Exception:
            pass
        print("  인라인 삽입:", cid, "<-", fname)
        count += 1
    return count


def main():
    if not os.path.exists(HTML_PATH):
        print("[오류] HTML 파일을 찾을 수 없습니다:", HTML_PATH)
        return
    if not os.path.isdir(ASSETS):
        print("[경고] assets 폴더를 찾을 수 없습니다:", ASSETS)
        print("       이미지가 본문에 표시되지 않을 수 있습니다.")
    else:
        print("[확인] assets 폴더:", ASSETS)

    outlook = win32com.client.Dispatch("Outlook.Application")
    mail = outlook.CreateItem(0)  # 0 = olMailItem
    mail.Subject = SUBJECT
    # 받는 사람은 검토 후 직접 입력하세요. 필요하면 아래 주석을 해제:
    # mail.To = "someone@example.com"

    mail.HTMLBody = load_html()  # 본문에 직접 기재
    attach_inline_images(mail)

    mail.Display(False)  # 메일 창 열기 (발송 X). 바로 보내려면 mail.Send()
    print("\n완료: Outlook 새 메일 창이 열렸습니다. 내용 확인 후 '보내기'를 누르세요.")

if __name__ == "__main__":
    main()
