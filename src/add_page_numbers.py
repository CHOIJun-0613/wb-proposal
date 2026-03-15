"""
pptx 폴더의 PPTX 파일들에 슬라이드 마스터 기반 연속 페이지 번호 추가
- 각 슬라이드에 직접 추가된 __page_num__ 텍스트박스 제거
- 슬라이드 마스터/레이아웃 우측 하단에 슬라이드 번호 placeholder 추가
- firstSlideNum 설정으로 파일 간 연속 페이지 번호 유지
"""

import os
import glob
import uuid
import xlsxwriter
from pathlib import Path
from datetime import datetime
from dotenv import load_dotenv
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.oxml.ns import qn
from lxml import etree

# 프로젝트 루트의 .env 로드 (src/ 의 상위 폴더)
_ROOT = Path(__file__).parent.parent
load_dotenv(_ROOT / ".env")

# 환경 변수
PPTX_DIR      = _ROOT / os.getenv("PPTX_DIR", "./pptx")
OUTPUT_DIR    = _ROOT / os.getenv("OUTPUT_DIR", "./output")
FONT_SIZE_PT  = float(os.getenv("PAGE_NUM_FONT_SIZE", "11"))
COLOR_HEX     = os.getenv("PAGE_NUM_COLOR", "606060")
MARGIN_RIGHT  = Inches(float(os.getenv("PAGE_NUM_MARGIN_RIGHT", "0.3")))
MARGIN_BOTTOM = Inches(float(os.getenv("PAGE_NUM_MARGIN_BOTTOM", "0.2")))
NUM_W         = Inches(float(os.getenv("PAGE_NUM_WIDTH", "1.0")))
NUM_H         = Pt(24)


def remove_page_num_textboxes(prs):
    """슬라이드별로 추가된 __page_num__ 텍스트박스 제거"""
    count = 0
    for slide in prs.slides:
        to_remove = [s._element for s in slide.shapes if s.name == "__page_num__"]
        for elem in to_remove:
            elem.getparent().remove(elem)
            count += 1
    return count


def _has_sldnum_placeholder(slide_like):
    spTree = slide_like.shapes._spTree
    for sp in spTree.findall(".//" + qn("p:sp")):
        ph = sp.find(".//" + qn("p:ph"))
        if ph is not None and ph.get("type") == "sldNum":
            return True
    return False


def _get_max_sp_id(spTree):
    max_id = 0
    for elem in spTree.iter():
        id_val = elem.get("id")
        if id_val is not None:
            try:
                max_id = max(max_id, int(id_val))
            except ValueError:
                pass
    return max_id


def _build_sldnum_sp_xml(sp_id, x, y, w, h):
    field_guid = "{" + str(uuid.uuid4()).upper() + "}"
    sz = int(FONT_SIZE_PT * 100)
    return f"""<p:sp xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
               xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <p:nvSpPr>
    <p:cNvPr id="{sp_id}" name="Slide Number Placeholder"/>
    <p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>
    <p:nvPr><p:ph type="sldNum" sz="quarter" idx="12"/></p:nvPr>
  </p:nvSpPr>
  <p:spPr>
    <a:xfrm>
      <a:off x="{int(x)}" y="{int(y)}"/>
      <a:ext cx="{int(w)}" cy="{int(h)}"/>
    </a:xfrm>
    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
    <a:noFill/>
    <a:ln><a:noFill/></a:ln>
  </p:spPr>
  <p:txBody>
    <a:bodyPr/>
    <a:lstStyle/>
    <a:p>
      <a:pPr algn="r"/>
      <a:fld id="{field_guid}" type="slidenum">
        <a:rPr lang="ko-KR" sz="{sz}" b="0" dirty="0">
          <a:solidFill><a:srgbClr val="{COLOR_HEX}"/></a:solidFill>
        </a:rPr>
        <a:t>‹#›</a:t>
      </a:fld>
    </a:p>
  </p:txBody>
</p:sp>"""


def _update_sldnum_position(slide_like, x, y, w, h):
    spTree = slide_like.shapes._spTree
    for sp in spTree.findall(".//" + qn("p:sp")):
        ph = sp.find(".//" + qn("p:ph"))
        if ph is not None and ph.get("type") == "sldNum":
            spPr = sp.find(qn("p:spPr"))
            if spPr is None:
                continue
            xfrm = spPr.find(qn("a:xfrm"))
            if xfrm is None:
                xfrm = etree.SubElement(spPr, qn("a:xfrm"))
            off = xfrm.find(qn("a:off"))
            if off is None:
                off = etree.SubElement(xfrm, qn("a:off"))
            ext = xfrm.find(qn("a:ext"))
            if ext is None:
                ext = etree.SubElement(xfrm, qn("a:ext"))
            off.set("x", str(int(x)))
            off.set("y", str(int(y)))
            ext.set("cx", str(int(w)))
            ext.set("cy", str(int(h)))


def add_sldnum_to_master_and_layouts(prs):
    sw = prs.slide_width
    sh = prs.slide_height
    x = sw - NUM_W - MARGIN_RIGHT
    y = sh - NUM_H - MARGIN_BOTTOM

    for target in [prs.slide_master] + list(prs.slide_layouts):
        if _has_sldnum_placeholder(target):
            _update_sldnum_position(target, x, y, NUM_W, NUM_H)
        else:
            spTree = target.shapes._spTree
            sp_id = _get_max_sp_id(spTree) + 1
            spTree.append(etree.fromstring(_build_sldnum_sp_xml(sp_id, x, y, NUM_W, NUM_H)))


def set_first_slide_num(prs, num):
    prs._element.set("firstSlideNum", str(num))


def count_slides(filepath):
    return len(Presentation(filepath).slides)


def save_index_excel(rows, out_dir):
    """목차페이지정보.xlsx 생성"""
    xlsx_path = out_dir / "목차페이지정보.xlsx"
    wb = xlsxwriter.Workbook(str(xlsx_path))
    ws = wb.add_worksheet("목차")

    # 스타일
    hdr_fmt = wb.add_format({
        "bold": True, "align": "center", "valign": "vcenter",
        "bg_color": "#4472C4", "font_color": "#FFFFFF",
        "border": 1, "font_size": 11,
    })
    cell_fmt = wb.add_format({
        "align": "left", "valign": "vcenter", "border": 1, "font_size": 10,
    })
    num_fmt = wb.add_format({
        "align": "center", "valign": "vcenter", "border": 1, "font_size": 10,
    })

    # 헤더
    headers = ["파일명", "슬라이드 수", "시작페이지", "끝페이지"]
    col_widths = [70, 12, 12, 12]
    for col, (h, w) in enumerate(zip(headers, col_widths)):
        ws.write(0, col, h, hdr_fmt)
        ws.set_column(col, col, w)
    ws.set_row(0, 20)

    # 데이터
    for row_idx, (fname, n_slides, start, end) in enumerate(rows, start=1):
        ws.write(row_idx, 0, fname,    cell_fmt)
        ws.write(row_idx, 1, n_slides, num_fmt)
        ws.write(row_idx, 2, start,    num_fmt)
        ws.write(row_idx, 3, end,      num_fmt)
        ws.set_row(row_idx, 18)

    wb.close()
    return xlsx_path


def main():
    files = sorted(glob.glob(str(PPTX_DIR / "*.pptx")))

    if not files:
        print(f"파일 없음: {PPTX_DIR}")
        return

    # 타임스탬프 출력 폴더 생성
    timestamp = datetime.now().strftime("%Y%m%d-%H%M%S")
    out_dir = OUTPUT_DIR / timestamp
    out_dir.mkdir(parents=True, exist_ok=True)

    print(f"설정값: PPTX_DIR={PPTX_DIR}, FONT={FONT_SIZE_PT}pt, COLOR=#{COLOR_HEX}")
    print(f"출력 폴더: {out_dir}")
    print(f"총 {len(files)}개 파일\n")

    print("[1/2] 슬라이드 수 파악 중...")
    slide_counts = [count_slides(f) for f in files]

    print("\n[2/2] 마스터 슬라이드 번호 설정 중...")
    page_start = 1
    index_rows = []
    for i, filepath in enumerate(files):
        prs = Presentation(filepath)

        removed = remove_page_num_textboxes(prs)
        add_sldnum_to_master_and_layouts(prs)
        set_first_slide_num(prs, page_start)

        out_path = out_dir / os.path.basename(filepath)
        prs.save(str(out_path))

        page_end = page_start + slide_counts[i] - 1
        index_rows.append((os.path.basename(filepath), slide_counts[i], page_start, page_end))

        print(f"  {os.path.basename(filepath)}")
        print(f"    페이지 {page_start} ~ {page_end}  ({slide_counts[i]}슬라이드)"
              + (f"  [텍스트박스 {removed}개 제거]" if removed else ""))
        page_start = page_end + 1

    # 목차 엑셀 저장
    xlsx_path = save_index_excel(index_rows, out_dir)
    print(f"\n완료! 총 {page_start - 1}페이지  →  {out_dir}")
    print(f"목차 엑셀: {xlsx_path.name}")


if __name__ == "__main__":
    main()
