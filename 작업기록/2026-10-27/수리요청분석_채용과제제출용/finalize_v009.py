from pathlib import Path
import json, shutil, hashlib
from pptx import Presentation
import pymupdf as fitz

W = Path(__file__).resolve().parent
R = W.parents[2]
V = W / 'versions'
Q = W / '검수_v009'

sha = lambda f: hashlib.sha256(f.read_bytes()).hexdigest()

for ext in ['pptx', 'pdf']:
    src = V / f'repair2_채용과제_분석결과_v009.{ext}'
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
assert norm('날짜 기록의 정합성을 검증하고 분석 대상을 구분했다') in p_texts[3]
assert norm('최초 860행에서 날짜가 행 순서대로 하루씩 증가하는 패턴 확인') in p_texts[3]

r = json.loads((Q / 'validation.json').read_text(encoding='utf-8'))
assert not r['overflow'] and not r['text_mismatch'] and not r['shape_bounds'] and r['raster_images'] == 0
r['latest_second_check'] = {
    'status': 'pass',
    'version_hash_match': True,
    'text_reopen_match': True,
    'v009_slide04_verified': True,
    'charts': charts,
    'pptx_sha256': sha(W / 'repair2_채용과제_분석결과.pptx'),
    'pdf_sha256': sha(W / 'repair2_채용과제_분석결과.pdf')
}
r['visual_review'] = '4쪽 제목(날짜 기록의 정합성을 검증하고 분석 대상을 구분했다), 검증 결과(최초 860행에서 날짜가 행 순서대로 하루씩 증가하는 패턴 확인), 배지·숫자(860건·226쌍) 중앙 정렬 맞춤 보정 및 실제 PDF 렌더 검수 완료 / 넘침 0·래스터 0'
(Q / 'validation.json').write_text(json.dumps(r, ensure_ascii=False, indent=2), encoding='utf-8')

readme_txt = (W / 'README.md').read_text(encoding='utf-8')
readme_txt = readme_txt.replace('최신 v008 / 13장 / 2026-10-27', '최신 v009 / 13장 / 2026-10-27')
readme_txt = readme_txt.replace('검수_v008/validation.json', '검수_v009/validation.json')
readme_txt = readme_txt.replace(
    '버전 보존: `versions/repair2_채용과제_분석결과_v008.pptx` 및 `.pdf` (이전 `v003~v007` 보존)',
    '버전 보존: `versions/repair2_채용과제_분석결과_v009.pptx` 및 `.pdf` (이전 `v003~v008` 보존)'
)
readme_txt = readme_txt.replace(
    '상단 제목을 `날짜 이상을 확인하고 임의 보정 대신 분석 대상을 구분했다`',
    '상단 제목을 `날짜 기록의 정합성을 검증하고 분석 대상을 구분했다`, 행 순서 검사 설명을 `최초 860행에서 날짜가 행 순서대로 하루씩 증가하는 패턴 확인`으로 다듬고 `860건`·`226쌍` 배지·숫자 중앙 정렬 맞춤'
)
(W / 'README.md').write_text(readme_txt, encoding='utf-8')

ch_txt = (W / 'CHANGELOG.md').read_text(encoding='utf-8')
ch_txt = ch_txt.replace(
    '# 변경 기록\n',
    '# 변경 기록\n## v009 / 2026-10-27\n- 4쪽 제목(`날짜 기록의 정합성을 검증하고 분석 대상을 구분했다`), 검증 결과(`최초 860행에서 날짜가 행 순서대로 하루씩 증가하는 패턴 확인`) 수정 및 좌측 배지·숫자(`860건`·`226쌍`) 중앙 정렬 맞춤 보정\n- 보존: `versions/repair2_채용과제_분석결과_v009.pptx` 및 `.pdf` (이전 `v003~v008` 유지)\n- 최신: `repair2_채용과제_분석결과.pptx` 및 `.pdf`\n\n'
)
(W / 'CHANGELOG.md').write_text(ch_txt, encoding='utf-8')

h = R / '작업인계.md'
shutil.copy2(h, V / '작업인계_v008.md')
h_txt = h.read_text(encoding='utf-8')
h_txt = h_txt.replace('최신 제출용 v008 13장', '최신 제출용 v009 13장')
h_txt = h_txt.replace('v008 반영:', 'v009 반영: 4쪽 제목·검증결과 문구 수정 및 860건·226쌍 중앙 맞춤 보정,')
h_txt = h_txt.replace('v003~v008은', 'v003~v009는')
h_txt = h_txt.replace('v008 전체 13장', 'v009 전체 13장')
h_txt = h_txt.replace('v009로 보존하고', 'v010으로 보존하고')
h.write_text(h_txt, encoding='utf-8')
print('v009 finalized and verified successfully')
