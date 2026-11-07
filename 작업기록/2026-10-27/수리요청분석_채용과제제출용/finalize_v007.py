from pathlib import Path
import json, shutil, hashlib
from pptx import Presentation
import pymupdf as fitz

W = Path(__file__).resolve().parent
R = W.parents[2]
V = W / 'versions'
Q = W / '검수_v007'

sha = lambda f: hashlib.sha256(f.read_bytes()).hexdigest()

for ext in ['pptx', 'pdf']:
    src = V / f'repair2_채용과제_분석결과_v007.{ext}'
    dst = W / f'repair2_채용과제_분석결과.{ext}'
    shutil.copy2(src, dst)
    assert sha(src) == sha(dst)

# 2nd Reopen Check on Latest Files
pp = Presentation(W / 'repair2_채용과제_분석결과.pptx')
doc = fitz.open(W / 'repair2_채용과제_분석결과.pdf')
assert len(pp.slides) == len(doc) == 13
norm = lambda t: ''.join(t.split())
for i, s in enumerate(pp.slides):
    pdf_txt = norm(doc[i].get_text())
    for sh in s.shapes:
        if sh.has_text_frame:
            assert norm(sh.text) in pdf_txt, (i + 1, sh.text)

charts = [list(se.values) for sl in pp.slides for sh in sl.shapes if sh.has_chart for se in sh.chart.series]
assert charts == [[39, 89, 63, 90, 58], [11.6, 160.4, 71.4]]

p_texts = [norm(doc[i].get_text()) for i in range(13)]
assert norm('노현동 / 데이터 분석 / 2026.10.26') in p_texts[0]
assert norm('데이터 분석 채용 과제') not in p_texts[0]

r = json.loads((Q / 'validation.json').read_text(encoding='utf-8'))
assert not r['overflow'] and not r['text_mismatch'] and not r['shape_bounds'] and r['raster_images'] == 0
r['latest_second_check'] = {
    'status': 'pass',
    'version_hash_match': True,
    'text_reopen_match': True,
    'v007_cover_text_verified': True,
    'charts': charts,
    'pptx_sha256': sha(W / 'repair2_채용과제_분석결과.pptx'),
    'pdf_sha256': sha(W / 'repair2_채용과제_분석결과.pdf')
}
r['visual_review'] = '1쪽 표지 하단(노현동 / 데이터 분석 / 2026.10.26) 및 상단(DATA ANALYSIS) 반영·실제 PDF 렌더 검수 완료 / 넘침 0·래스터 0'
(Q / 'validation.json').write_text(json.dumps(r, ensure_ascii=False, indent=2), encoding='utf-8')

readme_txt = (W / 'README.md').read_text(encoding='utf-8')
readme_txt = readme_txt.replace('최신 v006 / 13장 / 2026-10-27', '최신 v007 / 13장 / 2026-10-27')
readme_txt = readme_txt.replace('검수_v006/validation.json', '검수_v007/validation.json')
readme_txt = readme_txt.replace(
    '버전 보존: `versions/repair2_채용과제_분석결과_v006.pptx` 및 `.pdf` (이전 `v003`, `v004`, `v005` 보존)',
    '버전 보존: `versions/repair2_채용과제_분석결과_v007.pptx` 및 `.pdf` (이전 `v003`, `v004`, `v005`, `v006` 보존)'
)
readme_txt = readme_txt.replace('## v006 주요 반영 내역', '## v007 · v006 주요 반영 내역\n- **1페이지(표지)**: 하단 표기를 `노현동 / 데이터 분석 / 2026.10.26`으로 수정 (상단 영문 `DATA ANALYSIS`)')
(W / 'README.md').write_text(readme_txt, encoding='utf-8')

ch_txt = (W / 'CHANGELOG.md').read_text(encoding='utf-8')
ch_txt = ch_txt.replace(
    '# 변경 기록\n',
    '# 변경 기록\n## v007 / 2026-10-27\n- 1쪽 표지 텍스트 `데이터 분석 채용 과제` → `데이터 분석` 수정 (`노현동 / 데이터 분석 / 2026.10.26`, 상단 `DATA ANALYSIS`)\n- 보존: `versions/repair2_채용과제_분석결과_v007.pptx` 및 `.pdf` (이전 `v003~v006` 유지)\n- 최신: `repair2_채용과제_분석결과.pptx` 및 `.pdf`\n\n'
)
(W / 'CHANGELOG.md').write_text(ch_txt, encoding='utf-8')

h = R / '작업인계.md'
shutil.copy2(h, V / '작업인계_v006.md')
h_txt = h.read_text(encoding='utf-8')
h_txt = h_txt.replace('최신 제출용 v006 13장', '최신 제출용 v007 13장')
h_txt = h_txt.replace('v006 반영:', 'v007 반영: 1쪽 표지 `데이터 분석` 표기 수정 및 v006 반영 사항(')
h_txt = h_txt.replace('문구 다듬기 완수', '문구 다듬기) 완수')
h_txt = h_txt.replace('v003/v004/v005/v006은', 'v003~v007은')
h_txt = h_txt.replace('v006 전체 13장', 'v007 전체 13장')
h_txt = h_txt.replace('v007로 보존하고', 'v008로 보존하고')
h.write_text(h_txt, encoding='utf-8')
print('v007 finalized and verified successfully')
