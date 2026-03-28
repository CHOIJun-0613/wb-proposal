"""
pptx 폴더의 PPTX 파일들의 시작페이지 번호를 연속되도록 설정
- firstSlideNum 설정으로 파일 간 연속 페이지 번호 유지
"""

import os
import glob
import xlsxwriter
from pathlib import Path
from datetime import datetime
from dotenv import load_dotenv
from pptx import Presentation

# 프로젝트 루트의 .env 로드 (src/ 의 상위 폴더)
_ROOT = Path(__file__).parent.parent
load_dotenv(_ROOT / ".env")

# 환경 변수
PPTX_DIR      = _ROOT / os.getenv("PPTX_DIR", "./pptx")
OUTPUT_DIR    = _ROOT / os.getenv("OUTPUT_DIR", "./output")


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

    print(f"설정값: PPTX_DIR={PPTX_DIR}")
    print(f"출력 폴더: {out_dir}")
    print(f"총 {len(files)}개 파일\n")

    print("[1/2] 슬라이드 수 파악 중...")
    slide_counts = [count_slides(f) for f in files]

    print("\n[2/2] 마스터 슬라이드 번호 설정 중...")
    page_start = 1
    index_rows = []
    for i, filepath in enumerate(files):
        prs = Presentation(filepath)

        set_first_slide_num(prs, page_start)

        out_path = out_dir / os.path.basename(filepath)
        prs.save(str(out_path))

        page_end = page_start + slide_counts[i] - 1
        index_rows.append((os.path.basename(filepath), slide_counts[i], page_start, page_end))

        print(f"  {os.path.basename(filepath)}")
        print(f"    페이지 {page_start} ~ {page_end}  ({slide_counts[i]}슬라이드)")
        page_start = page_end + 1

    # 목차 엑셀 저장
    xlsx_path = save_index_excel(index_rows, out_dir)
    print(f"\n완료! 총 {page_start - 1}페이지  →  {out_dir}")
    print(f"목차 엑셀: {xlsx_path.name}")


if __name__ == "__main__":
    main()
