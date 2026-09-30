# -*- coding: utf-8 -*-
"""
HiNAS Newsletter Vol.2 -> Outlook HTML email with inline CID images.
Outlook new mail window opens pre-filled. Review and click Send manually.

Before running:
  1. Fill in all DUB_* URLs below (create each link at dub.sh first)
  2. Confirm asset image filenames are correct in IMAGES dict
"""
import os
import win32com.client

# ============================================================
# dub.sh 트래킹 URL
# ============================================================
DUB_METHODOLOGY = "https://dub.sh/sRRMtVd"
DUB_ROI         = "https://dub.sh/O7g8PGN"
DUB_KMTC        = "https://dub.sh/caVvwNs"
DUB_ALMI        = "https://dub.sh/2Ei75zj"
DUB_700         = "https://dub.sh/BW59t44"
DUB_CIDO        = "https://dub.sh/rmqO1oZ"
DUB_WEBSITE     = "https://dub.sh/SzEWcYD"
# ============================================================

BASE      = os.path.dirname(os.path.abspath(__file__))
HTML_PATH = os.path.join(BASE, "newsletter_email_vol2.html")
ASSETS    = os.path.join(os.path.dirname(BASE), "assets")

# cid 이름 (HTML 내 cid:xxx)  ->  assets/ 파일명
# 이미지 추가 시 여기에 넣고 HTML의 주석도 해제하세요
IMAGES = {
    "logo_hd":     "feat_hd.png",
    "logo_avikus": "avikus_wordmark.png",
    # --- 콘텐츠 이미지 (추가 예정) ---
    # "feat_nav_banner":   "feat_nav2.png",
    # "feat_kmtc_img":     "feat_control1.png",
    # "feat_almi_img":     "feat_control2.png",
    # "feat_fleet_img":    "feat_cloud.png",
    # "feat_retrofit_img": "feat_svm1.png",
}

PR_ATTACH_CONTENT_ID = "http://schemas.microsoft.com/mapi/proptag/0x3712001F"
PR_ATTACH_MIME_TAG   = "http://schemas.microsoft.com/mapi/proptag/0x370E001F"
PR_ATTACHMENT_HIDDEN = "http://schemas.microsoft.com/mapi/proptag/0x7FFE000B"

def main():
    with open(HTML_PATH, "r", encoding="utf-8") as f:
        html = f.read()

    # dub.sh URL 주입
    html = html.replace("{{DUB_METHODOLOGY}}", DUB_METHODOLOGY)
    html = html.replace("{{DUB_ROI}}",         DUB_ROI)
    html = html.replace("{{DUB_KMTC}}",        DUB_KMTC)
    html = html.replace("{{DUB_ALMI}}",        DUB_ALMI)
    html = html.replace("{{DUB_700}}",         DUB_700)
    html = html.replace("{{DUB_CIDO}}",        DUB_CIDO)
    html = html.replace("{{DUB_WEBSITE}}",     DUB_WEBSITE)

    outlook = win32com.client.Dispatch("Outlook.Application")
    mail = outlook.CreateItem(0)
    mail.Subject = "HiNAS Newsletter — Vol.2"
    # mail.To = "someone@example.com"

    mail.HTMLBody = html

    for cid, fname in IMAGES.items():
        path = os.path.join(ASSETS, fname)
        if not os.path.exists(path):
            print(f"  (건너뜀, 파일 없음): {path}")
            continue
        att = mail.Attachments.Add(path, 1)
        pa  = att.PropertyAccessor
        pa.SetProperty(PR_ATTACH_CONTENT_ID, cid)
        pa.SetProperty(PR_ATTACH_MIME_TAG, "image/png")
        try:
            pa.SetProperty(PR_ATTACHMENT_HIDDEN, True)
        except Exception:
            pass
        print(f"  인라인 삽입: {cid} <- {fname}")

    mail.Display(False)
    print("\n완료: Outlook 새 메일 창이 열렸습니다. 내용 확인 후 '보내기'를 누르세요.")

if __name__ == "__main__":
    main()
