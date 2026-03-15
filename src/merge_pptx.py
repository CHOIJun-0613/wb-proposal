"""
PPTX 파일 병합 스크립트
pptx/ 폴더의 모든 파일을 하나의 PPTX로 병합 (ZIP 레벨 조작)
"""

import zipfile, os, re, glob
from pathlib import Path
from datetime import datetime
from dotenv import load_dotenv
from lxml import etree

# 프로젝트 루트의 .env 로드 (src/ 의 상위 폴더)
_ROOT = Path(__file__).parent.parent
load_dotenv(_ROOT / ".env")

# 환경 변수
PPTX_DIR     = _ROOT / os.getenv("PPTX_DIR", "./pptx")
OUTPUT_DIR   = _ROOT / os.getenv("OUTPUT_DIR", "./output")
MERGE_OUTPUT = os.getenv("MERGE_OUTPUT", "III.기술부문.pptx")

RELS_NS = 'http://schemas.openxmlformats.org/package/2006/relationships'
CT_NS   = 'http://schemas.openxmlformats.org/package/2006/content-types'
PML_NS  = 'http://schemas.openxmlformats.org/presentationml/2006/main'
R_NS    = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'


# ── 헬퍼 ──────────────────────────────────────────────────────────────────────

def xparse(data):
    return etree.fromstring(data)

def xbytes(elem):
    return etree.tostring(elem, xml_declaration=True, encoding='UTF-8', standalone=True)

def rels_of(path):
    """'ppt/slides/slide1.xml' → 'ppt/slides/_rels/slide1.xml.rels'"""
    d, f = path.rsplit('/', 1)
    return f'{d}/_rels/{f}.rels'

def resolve(base, rel_target):
    """상대 경로 → ZIP 내 절대 경로"""
    parts = (base.rsplit('/', 1)[0] + '/' + rel_target).split('/')
    stack = []
    for p in parts:
        if p == '..':
            if stack: stack.pop()
        elif p and p != '.':
            stack.append(p)
    return '/'.join(stack)

def max_num(files, pattern):
    n = 0
    for p in files:
        m = re.search(pattern, p)
        if m:
            try: n = max(n, int(m.group(1)))
            except: pass
    return n


# ── 메인 병합 클래스 ───────────────────────────────────────────────────────────

class Merger:

    def load_base(self, path):
        self.files = {}
        with zipfile.ZipFile(path, 'r') as zf:
            for n in zf.namelist():
                self.files[n] = zf.read(n)

        self.next_slide = max_num(self.files, r'ppt/slides/slide(\d+)\.xml$') + 1
        self.next_media = max_num(self.files, r'ppt/media/[^\d]*(\d+)\.\w+$') + 1
        self.next_chart = max_num(self.files, r'ppt/charts/chart(\d+)\.xml$') + 1
        self.next_diag  = max_num(self.files, r'ppt/diagrams/data(\d+)\.xml$') + 1
        self.next_embed = max_num(self.files, r'ppt/embeddings/[^\d]*(\d+)\.') + 1

        self.prs      = xparse(self.files['ppt/presentation.xml'])
        self.prs_rels = xparse(self.files['ppt/_rels/presentation.xml.rels'])
        self.ct       = xparse(self.files['[Content_Types].xml'])

        sldIdLst = self.prs.find(f'{{{PML_NS}}}sldIdLst')
        self.max_sld_id = max(
            (int(el.get('id', 0)) for el in sldIdLst), default=256
        )
        n_slides = self.next_slide - 1
        print(f"  기준 파일: {os.path.basename(path)} ({n_slides}슬라이드)")

    def append(self, path):
        src = {}
        with zipfile.ZipFile(path, 'r') as zf:
            for n in zf.namelist():
                src[n] = zf.read(n)

        slides = sorted(
            [n for n in src if re.match(r'ppt/slides/slide\d+\.xml$', n)],
            key=lambda x: int(re.search(r'(\d+)', os.path.basename(x)).group(1))
        )
        for s in slides:
            self._copy_slide(src, s)
        print(f"  추가: {os.path.basename(path)} ({len(slides)}슬라이드)")

    # ── 슬라이드 복사 ──────────────────────────────────────────────────────────

    def _copy_slide(self, src, slide_path):
        ns = self.next_slide
        self.next_slide += 1
        new_slide = f'ppt/slides/slide{ns}.xml'
        new_rels  = f'ppt/slides/_rels/slide{ns}.xml.rels'

        old_rels = rels_of(slide_path)
        if old_rels in src:
            rels_xml = xparse(src[old_rels])
            diag_map = {}

            for rel in rels_xml.findall(f'{{{RELS_NS}}}Relationship'):
                target = rel.get('Target', '')
                if not target or target.startswith('#') or '://' in target:
                    continue
                abs_t = resolve(slide_path, target)

                if '/media/' in abs_t:
                    new_p = self._copy_media(src, abs_t)
                    if new_p:
                        rel.set('Target', '../media/' + os.path.basename(new_p))

                elif '/charts/' in abs_t and abs_t.endswith('.xml'):
                    new_p = self._copy_chart(src, abs_t)
                    if new_p:
                        rel.set('Target', '../charts/' + os.path.basename(new_p))

                elif '/diagrams/' in abs_t:
                    m = re.search(r'(\d+)', os.path.basename(abs_t))
                    old_n = int(m.group(1)) if m else 0
                    if old_n not in diag_map:
                        diag_map[old_n] = self.next_diag
                        self.next_diag += 1
                    new_p = self._copy_diag(src, abs_t, diag_map[old_n])
                    if new_p:
                        rel.set('Target', '../diagrams/' + os.path.basename(new_p))

                elif '/embeddings/' in abs_t:
                    new_p = self._copy_embed(src, abs_t)
                    if new_p:
                        rel.set('Target', '../embeddings/' + os.path.basename(new_p))

            self.files[new_rels] = xbytes(rels_xml)

        self.files[new_slide] = src[slide_path]

        self.max_sld_id += 1
        rid = f'rId_s{ns}'
        sldIdLst = self.prs.find(f'{{{PML_NS}}}sldIdLst')
        el = etree.SubElement(sldIdLst, f'{{{PML_NS}}}sldId')
        el.set('id', str(self.max_sld_id))
        el.set(f'{{{R_NS}}}id', rid)

        rel_el = etree.SubElement(self.prs_rels, f'{{{RELS_NS}}}Relationship')
        rel_el.set('Id', rid)
        rel_el.set('Type', f'{R_NS}/slide')
        rel_el.set('Target', f'slides/slide{ns}.xml')

        pn = f'/ppt/slides/slide{ns}.xml'
        existing = {e.get('PartName') for e in self.ct.findall(f'{{{CT_NS}}}Override')}
        if pn not in existing:
            ov = etree.SubElement(self.ct, f'{{{CT_NS}}}Override')
            ov.set('PartName', pn)
            ov.set('ContentType',
                   'application/vnd.openxmlformats-officedocument.presentationml.slide+xml')

    # ── 리소스 복사 헬퍼 ───────────────────────────────────────────────────────

    def _copy_media(self, src, abs_path):
        if abs_path not in src:
            return None
        ext  = os.path.splitext(abs_path)[1]
        base = re.sub(r'\d+', '', os.path.splitext(os.path.basename(abs_path))[0]) or 'image'
        new_path = f'ppt/media/{base}{self.next_media}{ext}'
        self.next_media += 1
        self.files[new_path] = src[abs_path]
        self._add_ct(src, abs_path, new_path)
        return new_path

    def _copy_chart(self, src, abs_path):
        n = self.next_chart
        self.next_chart += 1
        new_chart = f'ppt/charts/chart{n}.xml'
        new_crels = f'ppt/charts/_rels/chart{n}.xml.rels'

        if abs_path in src:
            self.files[new_chart] = src[abs_path]
            self._add_ct(src, abs_path, new_chart)

        old_crels = rels_of(abs_path)
        if old_crels in src:
            crels_xml = xparse(src[old_crels])
            for rel in crels_xml.findall(f'{{{RELS_NS}}}Relationship'):
                tgt = rel.get('Target', '')
                if not tgt: continue
                abs_t = resolve(abs_path, tgt)
                if '/embeddings/' in abs_t:
                    new_p = self._copy_embed(src, abs_t)
                    if new_p:
                        rel.set('Target', '../embeddings/' + os.path.basename(new_p))
                elif '/media/' in abs_t:
                    new_p = self._copy_media(src, abs_t)
                    if new_p:
                        rel.set('Target', '../media/' + os.path.basename(new_p))
            self.files[new_crels] = xbytes(crels_xml)

        return new_chart

    def _copy_diag(self, src, abs_path, new_num):
        if abs_path not in src:
            return None
        fname     = os.path.basename(abs_path)
        new_fname = re.sub(r'\d+', str(new_num), fname, count=1)
        new_path  = f'ppt/diagrams/{new_fname}'
        if new_path not in self.files:
            self.files[new_path] = src[abs_path]
            self._add_ct(src, abs_path, new_path)
        return new_path

    def _copy_embed(self, src, abs_path):
        if abs_path not in src:
            return None
        fname = os.path.basename(abs_path)
        name, ext = os.path.splitext(fname)
        new_name = re.sub(r'\d+', str(self.next_embed), name, count=1)
        if new_name == name:
            new_name = f'{name}_{self.next_embed}'
        new_path = f'ppt/embeddings/{new_name}{ext}'
        while new_path in self.files:
            self.next_embed += 1
            new_name = re.sub(r'\d+', str(self.next_embed), name, count=1)
            if new_name == name:
                new_name = f'{name}_{self.next_embed}'
            new_path = f'ppt/embeddings/{new_name}{ext}'
        self.next_embed += 1
        self.files[new_path] = src[abs_path]
        self._add_ct(src, abs_path, new_path)
        return new_path

    def _add_ct(self, src, old_abs, new_abs):
        old_pn = '/' + old_abs
        new_pn = '/' + new_abs
        existing = {e.get('PartName') for e in self.ct.findall(f'{{{CT_NS}}}Override')}
        if new_pn in existing:
            return
        try:
            src_ct = xparse(src['[Content_Types].xml'])
            for ov in src_ct.findall(f'{{{CT_NS}}}Override'):
                if ov.get('PartName') == old_pn:
                    new_ov = etree.SubElement(self.ct, f'{{{CT_NS}}}Override')
                    new_ov.set('PartName', new_pn)
                    new_ov.set('ContentType', ov.get('ContentType', ''))
                    return
            ext = os.path.splitext(new_abs)[1].lstrip('.')
            tgt_exts = {d.get('Extension', '').lower()
                        for d in self.ct.findall(f'{{{CT_NS}}}Default')}
            for df in src_ct.findall(f'{{{CT_NS}}}Default'):
                if df.get('Extension', '').lower() == ext.lower() and ext.lower() not in tgt_exts:
                    new_df = etree.SubElement(self.ct, f'{{{CT_NS}}}Default')
                    new_df.set('Extension', df.get('Extension', ''))
                    new_df.set('ContentType', df.get('ContentType', ''))
                    return
        except Exception:
            pass

    # ── 저장 ──────────────────────────────────────────────────────────────────

    def set_first_slide_num(self, n):
        self.prs.set('firstSlideNum', str(n))

    def save(self, output_path):
        self.files['ppt/presentation.xml'] = xbytes(self.prs)
        self.files['ppt/_rels/presentation.xml.rels'] = xbytes(self.prs_rels)
        self.files['[Content_Types].xml'] = xbytes(self.ct)

        with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zf:
            for name, data in self.files.items():
                zf.writestr(name, data)

        total = max_num(self.files, r'ppt/slides/slide(\d+)\.xml$')
        size_mb = os.path.getsize(output_path) / 1024 / 1024
        print(f"\n저장 완료: {output_path}")
        print(f"총 {total}슬라이드 / {size_mb:.1f} MB")


# ── 실행 ──────────────────────────────────────────────────────────────────────

def main():
    files = sorted(glob.glob(str(PPTX_DIR / "*.pptx")))
    if not files:
        print(f"파일 없음: {PPTX_DIR}")
        return

    # 타임스탬프 출력 폴더 생성
    timestamp = datetime.now().strftime("%Y%m%d-%H%M%S")
    out_dir = OUTPUT_DIR / timestamp
    out_dir.mkdir(parents=True, exist_ok=True)
    output = out_dir / MERGE_OUTPUT

    print(f"설정값: PPTX_DIR={PPTX_DIR}")
    print(f"출력 폴더: {out_dir}")
    print(f"총 {len(files)}개 파일 병합 시작\n")
    m = Merger()
    m.load_base(files[0])
    for f in files[1:]:
        m.append(f)
    m.set_first_slide_num(1)
    m.save(str(output))


if __name__ == '__main__':
    main()
