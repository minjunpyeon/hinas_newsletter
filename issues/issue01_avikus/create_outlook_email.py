# -*- coding: utf-8 -*-
"""
Avikus Newsletter (Issue 01 / 창간호) -> Outlook 본문(HTML) + 이미지 CID 인라인 삽입.
docx(뉴스레터_일간호.docx) 내용을 바탕으로 만든 newsletter_email.html 을
Outlook 새 메일 본문에 '박힌 채로' 띄웁니다. (자동 발송 X — 검토 후 직접 '보내기')
"""
import os
import win32com.client

BASE = os.path.dirname(os.path.abspath(__file__))
HTML_PATH = os.path.join(BASE, "newsletter_email.html")
ASSETS = os.path.join(BASE, "assets")

# HTML 안의 cid:이름  ->  실제 이미지 파일
IMAGES = {
    "avikus_logo":   "logo_avikus_2.png",    # Avikus 정식 로고 (비킹 심볼 포함)
    "feat_nav2":     "feat_nav2.png",        # NAV 제품 사진 (흰 배경 크롭)
    "feat_svm1":     "feat_svm1.png",        # SVM 제품 사진
    "feat_control1": "feat_control1.png",    # CONTROL 제품 사진 1
    "feat_control2": "feat_control2.png",    # CONTROL 제품 사진 2
    "feat_cloud":    "feat_cloud.png",       # CLOUD 제품 사진
}

# MAPI property tags
PR_ATTACH_CONTENT_ID  = "http://schemas.microsoft.com/mapi/proptag/0x3712001F"
PR_ATTACH_MIME_TAG    = "http://schemas.microsoft.com/mapi/proptag/0x370E001F"
PR_ATTACHMENT_HIDDEN  = "http://schemas.microsoft.com/mapi/proptag/0x7FFE000B"

# 메일 제목 (대량 발송 send_bulk.py 도 이 값을 그대로 사용)
SUBJECT = "Avikus Newsletter — Issue 01: Who is Avikus?"


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
        print("  inline:", cid, "<-", fname)
        count += 1
    return count


def main():
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
