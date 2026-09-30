# -*- coding: utf-8 -*-
"""
newsletter_email.html 을 '보이는 그대로' PDF 로 뽑는 도구.

메일 HTML 은 이미지를 cid:이름 (이메일 인라인 참조) 로 부르기 때문에 브라우저로
그냥 열면 이미지가 안 나옵니다. 그래서:
  1) create_outlook_email.py 의 IMAGES 대로 cid:이름 → 실제 이미지(base64) 로 치환
     → 자체완결 인쇄용 HTML 생성
  2) Windows 기본 Edge(또는 Chrome) 헤드리스로 PDF 변환
     → 브라우저가 없거나 실패하면, 생성된 HTML 을 직접 열어 Ctrl+P ▸ PDF 로 저장

세 가지 레이아웃을 한 번에 생성합니다:
  - newsletter_top.pdf  : 섹션 1개 = 페이지 1장, 페이지 '위쪽' 정렬(전 버전)
  - newsletter.pdf      : 섹션 1개 = 페이지 1장, 페이지 '세로 중앙' 정렬(현재 버전)
  - newsletter_full.pdf : <hr> 로 안 나누고 원본 그대로 쭉 이어서 출력(연속형)

실행:  py build_pdf.py
필요:  별도 설치 없음 (Edge/Chrome 중 하나만 있으면 됨)
"""
import os
import base64
import subprocess

import create_outlook_email as base  # HTML_PATH / ASSETS / IMAGES 재사용

BASE = os.path.dirname(os.path.abspath(__file__))

# 각 섹션 축소 배율. 1.0=원본. 큰 섹션이 한 장에 안 들어가면 살짝 줄임.
# (환경변수 BUILD_SCALE 로 임시 override 가능)
SCALE = float(os.environ.get("BUILD_SCALE", "0.92"))

# 생성할 작업 목록: (레이아웃, 출력 PDF, 중간 HTML)
JOBS = [
    ("top",    "newsletter_top.pdf",  "newsletter_top_print.html"),
    ("center", "newsletter.pdf",      "newsletter_print.html"),
    ("raw",    "newsletter_full.pdf", "newsletter_full_print.html"),
]

# 섹션을 나누는 구분선 행(HTML 원본과 정확히 일치) — 이 두 곳에서 페이지가 나뉨.
# (표지 상단의 헤더 밑줄 <hr> 은 이 패턴과 달라서 건드리지 않음)
SECTION_SEP = ('<tr><td style="padding:30px 44px 0 44px;">'
               '<hr style="border:0; border-top:1.5px solid #1F3A5F; margin:0;"></td></tr>')
# top 레이아웃: 구분선 자리에 넣을 페이지 나눔 행
PAGE_BREAK = ('<tr style="page-break-before:always; break-before:page;">'
              '<td style="padding:0;height:0;line-height:0;font-size:0;">&#8203;</td></tr>')
# center 레이아웃: 640px 카드 표를 섹션별로 쪼개는 기준점
CARD_OPEN = ('<table role="presentation" width="640" cellpadding="0" cellspacing="0" border="0" '
             'style="width:640px; max-width:640px; background:#ffffff; border-collapse:collapse;">')
END_COMMENT = '<!-- ============ /MAIN CARD ============ -->'

# Edge / Chrome 실행 파일 후보
BROWSERS = [
    r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
    r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
    r"C:\Program Files\Google\Chrome\Application\chrome.exe",
    r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
]


def _data_uri(path):
    with open(path, "rb") as f:
        b64 = base64.b64encode(f.read()).decode("ascii")
    return "data:image/png;base64," + b64


def _print_style(layout):
    """레이아웃별 인쇄 스타일(배경색 인쇄 + 페이지 구성)."""
    color = "  * { -webkit-print-color-adjust: exact !important; print-color-adjust: exact !important; }\n"
    if layout == "center":
        return (
            "\n<style>\n" + color +
            "  @media print {\n"
            "    @page { size: A4; margin: 0; }\n"          # 여백은 .page 안쪽 padding 으로
            "    html, body { margin: 0 !important; }\n"
            "    .page { min-height: 100vh; box-sizing: border-box; padding: 10mm;\n"
            "            display: flex; flex-direction: column; justify-content: center;\n"
            "            page-break-after: always; }\n"
            "    .page:last-child { page-break-after: auto; }\n"
            f"    .page > table {{ zoom: {SCALE}; }}\n"
            "  }\n"
            "</style>\n"
        )
    if layout == "raw":
        # 원본 그대로 연속 출력 — 축소 없이(1.0) A4 여백만
        return (
            "\n<style>\n" + color +
            "  @media print { @page { size: A4; margin: 10mm; } }\n"
            "</style>\n"
        )
    # top
    return (
        "\n<style>\n" + color +
        "  @media print {\n"
        "    @page { size: A4; margin: 10mm; }\n"
        f"    html {{ zoom: {SCALE}; }}\n"
        "  }\n"
        "</style>\n"
    )


def _apply_layout(html, layout):
    """레이아웃에 맞게 페이지 나눔 구조를 적용."""
    if layout == "raw":
        # <hr> 로 안 나누고 원본 그대로 연속 출력
        print("  [raw] 페이지 나눔 없이 원본 그대로 연속 출력")
        return html

    cnt = html.count(SECTION_SEP)
    if not cnt:
        print("  [주의] 섹션 구분선 패턴을 못 찾음 — 페이지 나눔 없이 진행")
        return html

    if layout == "center":
        if CARD_OPEN in html and END_COMMENT in html:
            # 외곽 래퍼 td 여백 제거(페이지 정렬이 밀리지 않게)
            html = html.replace('<td align="center" style="padding:28px 12px;">',
                                '<td align="center" style="padding:0;">', 1)
            # 카드 표를 섹션별로 쪼개 각각 .page div 로 감쌈
            html = html.replace(CARD_OPEN, '<div class="page">\n' + CARD_OPEN, 1)
            html = html.replace(SECTION_SEP, '</table>\n</div>\n<div class="page">\n' + CARD_OPEN)
            html = html.replace(END_COMMENT, '</div>\n' + END_COMMENT, 1)
            print(f"  [center] 섹션 {cnt + 1}개를 각 페이지 세로 중앙에 배치")
        else:
            print("  [주의] 카드 패턴을 못 찾음 — center 레이아웃 생략")
    else:  # top
        html = html.replace(SECTION_SEP, PAGE_BREAK)
        print(f"  [top] 섹션 {cnt + 1}개를 각 페이지 위쪽에 배치")
    return html


def build_print_html(layout, out_name):
    """cid → 이미지(base64) 치환 + 스타일/레이아웃 적용한 인쇄용 HTML 생성."""
    html = base.load_html()
    for cid in sorted(base.IMAGES, key=len, reverse=True):  # 긴 이름부터(부분일치 방지)
        img = os.path.join(base.ASSETS, base.IMAGES[cid])
        if not os.path.exists(img):
            print("  (건너뜀, 파일 없음):", img)
            continue
        html = html.replace("cid:" + cid, _data_uri(img))

    style = _print_style(layout)
    if "</head>" in html:
        html = html.replace("</head>", style + "</head>", 1)
    else:
        html = style + html

    html = _apply_layout(html, layout)

    out_path = os.path.join(BASE, out_name)
    with open(out_path, "w", encoding="utf-8") as f:
        f.write(html)
    return out_path


def find_browser():
    for exe in BROWSERS:
        if os.path.exists(exe):
            return exe
    return None


def to_pdf(browser, html_path, pdf_path):
    """Edge/Chrome 헤드리스로 html → pdf. 성공하면 True."""
    url = "file:///" + html_path.replace("\\", "/")
    cmd = [
        browser,
        "--headless=new",
        "--disable-gpu",
        "--no-pdf-header-footer",            # 머리글/바닥글 제거
        "--run-all-compositor-stages-before-draw",
        "--virtual-time-budget=10000",       # 렌더 완료 대기
        "--print-to-pdf=" + pdf_path,
        url,
    ]
    try:
        subprocess.run(cmd, check=True, timeout=120,
                       stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    except Exception:
        cmd[1] = "--headless"  # 구버전 호환
        try:
            subprocess.run(cmd, check=True, timeout=120,
                           stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        except Exception as e:
            print("  [실패] PDF 변환 중 오류:", e)
            return False
    return os.path.exists(pdf_path)


def main():
    if not os.path.exists(base.HTML_PATH):
        print("[오류] HTML 을 찾을 수 없습니다:", base.HTML_PATH)
        return

    browser = find_browser()
    for layout, pdf_name, html_name in JOBS:
        print(f"\n=== [{layout}] {pdf_name} ===")
        html_path = build_print_html(layout, html_name)
        pdf_path = os.path.join(BASE, pdf_name)
        if not browser:
            print("  Edge/Chrome 미발견 — 아래 HTML 을 열고 Ctrl+P ▸ PDF 저장:")
            print("  ", html_path)
            continue
        if to_pdf(browser, html_path, pdf_path):
            print("  완료 →", pdf_path)
        else:
            print("  변환 실패 — 아래 HTML 을 열고 Ctrl+P ▸ PDF 저장:")
            print("  ", html_path)


if __name__ == "__main__":
    main()
