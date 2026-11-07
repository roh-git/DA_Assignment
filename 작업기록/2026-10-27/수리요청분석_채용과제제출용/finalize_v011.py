from pathlib import Path
import json, shutil, hashlib
from pptx import Presentation
import pymupdf as fitz

W = Path(__file__).resolve().parent
R = W.parents[2]
V = W / 'versions'
Q = W / '검수_v011'

sha = lambda f: hashlib.sha256(f.read_bytes()).hexdigest()

for ext in ['pptx', 'pdf']:
    src = V / f'repair2_채용과제_분석결과_v011.{ext}'
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
    if i >= 1:
        assert norm('repair2_분석.ipynb |') in pdf_txt, f"Missing notebook ref on slide {i+1}"
    for sh in s.shapes:
        if sh.has_text_frame:
            assert norm(sh.text) in pdf_txt, (i + 1, sh.text)

charts = [list(se.values) for sl in pp.slides for sh in sl.shapes if sh.has_chart for se in sh.chart.series]
assert charts == [[39, 89, 63, 90, 58], [11.6, 160.4, 71.4]]

r = json.loads((Q / 'validation.json').read_text(encoding='utf-8'))
assert not r['overflow'] and not r['text_mismatch'] and not r['shape_bounds'] and r['raster_images'] == 0
r['latest_second_check'] = {
    'status': 'pass',
    'version_hash_match': True,
    'text_reopen_match': True,
    'notebook_location_footer_verified': True,
    'charts': charts,
    'pptx_sha256': sha(W / 'repair2_채용과제_분석결과.pptx'),
    'pdf_sha256': sha(W / 'repair2_채용과제_분석결과.pdf')
}
r['visual_review'] = '전 슬라이드(2~13쪽) 하단 출처를 repair2_분석.ipynb 내 실제 섹션 번호·제목으로 일관되게 통일 및 실제 PDF 렌더 검수 완료 / 넘침 0·래스터 0'
(Q / 'validation.json').write_text(json.dumps(r, ensure_ascii=False, indent=2), encoding='utf-8')

readme_txt = (W / 'README.md').read_text(encoding='utf-8')
readme_txt = readme_txt.replace('최신 v010 / 13장 / 2026-10-27', '최신 v011 / 13장 / 2026-10-27')
readme_txt = readme_txt.replace('검수_v010/validation.json', '검수_v011/validation.json')
readme_txt = readme_txt.replace(
    '버전 보존: `versions/repair2_채용과제_분석결과_v010.pptx` 및 `.pdf` (이전 `v003~v009` 보존)',
    '버전 보존: `versions/repair2_채용과제_분석결과_v011.pptx` 및 `.pdf` (이전 `v003~v010` 보존)'
)
readme_txt = readme_txt.replace(
    '- **하단 푸터**: 슬라이드 하단의 `상세 보고서 5–7쪽`(3·4쪽), `상세 보고서 9쪽`(7쪽) 표기 일괄 삭제',
    '- **하단 푸터 통일**: 전 슬라이드(2~13쪽) 하단을 `repair2_분석.ipynb  |  <해당 섹션 번호 및 제목>` 형식으로 통일하여 노트북 내 실제 분석 위치를 명확히 안내'
)
(W / 'README.md').write_text(readme_txt, encoding='utf-8')

ch_txt = (W / 'CHANGELOG.md').read_text(encoding='utf-8')
ch_txt = ch_txt.replace(
    '# 변경 기록\n',
    '# 변경 기록\n## v011 / 2026-10-27\n- 전 슬라이드(2~13쪽) 하단 푸터를 `repair2_분석.ipynb  |  <섹션 번호·제목>` 형식으로 통일하여 노트북 내 위치를 일관되게 명시\n- 보존: `versions/repair2_채용과제_분석결과_v011.pptx` 및 `.pdf` (이전 `v003~v010` 유지)\n- 최신: `repair2_채용과제_분석결과.pptx` 및 `.pdf`\n\n'
)
(W / 'CHANGELOG.md').write_text(ch_txt, encoding='utf-8')

h = R / '작업인계.md'
shutil.copy2(h, V / '작업인계_v010.md')
h_txt = h.read_text(encoding='utf-8')
h_txt = h_txt.replace('최신 제출용 v010 13장', '최신 제출용 v011 13장')
h_txt = h_txt.replace('v010 반영: 하단 상세 보고서 쪽수 표기 일괄 삭제,', 'v011 반영: 2~13쪽 하단 푸터를 `repair2_분석.ipynb | 섹션 번호·제목`으로 통일,')
h_txt = h_txt.replace('v003~v010은', 'v003~v011은')
h_txt = h_txt.replace('v010 전체 13장', 'v011 전체 13장')
h_txt = h_txt.replace('v011로 보존하고', 'v012로 보존하고')
h.write_text(h_txt, encoding='utf-8')
print('v011 finalized and verified successfully')
