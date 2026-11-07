from pathlib import Path
import json, shutil, hashlib
from pptx import Presentation
import pymupdf as fitz

W = Path(__file__).resolve().parent
R = W.parents[2]
V = W / 'versions'
Q = W / '검수_v006'

sha = lambda f: hashlib.sha256(f.read_bytes()).hexdigest()

for ext in ['pptx', 'pdf']:
    src = V / f'repair2_채용과제_분석결과_v006.{ext}'
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

# Verify all v006 requested text updates
p_texts = [norm(doc[i].get_text()) for i in range(13)]
assert norm('수리 유형과 수행 업무에 따라 처리 시간이 다른가') in p_texts[1]
assert norm('접수 채널과 담당자 기록에 따라 처리 시간이 다른가') in p_texts[1]
assert norm('날짜 오류와 관계없이 수행된 업무와 경로 확인') in p_texts[2]
assert norm('날짜 이상을 확인하고 임의 보정 대신 분석 대상을 구분했다') in p_texts[3]
assert norm('행 순서에 따른 날짜 증가와 업무 시작·완료 시간 역전을 확인해 날짜 기록의 신뢰성을 검토') in p_texts[3]
assert norm('한 요청에서 여러 수리 업무나 동일 업무의 반복이 관찰됨') in p_texts[4]
assert norm('반복 작업 여부를 확인하기 위해 업무별 수행 횟수와 순서를 분석') in p_texts[5]
assert norm('업무 809쌍과 요청 339건은 분석 대상이 다르며 중앙값 비교만으로 관계를 단정할 수 없음') in p_texts[6]
assert norm('24시간은 분포 확인용 기준이며 정해진 처리 기한이 아님') in p_texts[8] and 'SLA' not in p_texts[8]
assert norm('시작·완료 담당자 차이만으로 실제 업무 인수인계를 단정할 수 없음') in p_texts[10]
assert norm('예정 시각 초과 사례를 확인하고 지연 사유에 따라 일정 재조정') in p_texts[11]
assert norm('반복되는 지연 원인을 확인하고 수리 착수 일정과 처리 절차 개선') in p_texts[11]
assert norm('채널별 처리 과정과 담당자 기록을 점검하고 확인된 원인에 맞춰 개선') in p_texts[11]
assert norm('신뢰할 수 있는 시간 기록을 확보하고 수리 유형과 접수 채널을 구분하여 개선 전후의 처리 시간을 비교') in p_texts[12]
assert norm('현장 조사 시작 전 대기를 우선 검토 대상으로 선정하고 실제 지연 원인 확인 후 전체 처리 시간의 감소 여부로 개선 효과를 검증') in p_texts[12]

r = json.loads((Q / 'validation.json').read_text(encoding='utf-8'))
assert not r['overflow'] and not r['text_mismatch'] and not r['shape_bounds'] and r['raster_images'] == 0
r['latest_second_check'] = {
    'status': 'pass',
    'version_hash_match': True,
    'text_reopen_match': True,
    'v006_checklist_verified': True,
    'charts': charts,
    'pptx_sha256': sha(W / 'repair2_채용과제_분석결과.pptx'),
    'pdf_sha256': sha(W / 'repair2_채용과제_분석결과.pdf')
}
r['visual_review'] = '전 13쪽 실제 PDF 렌더 및 줄바꿈(expected_lines == actual_lines) 전수 검수 완료 / 4·7·9·12·13쪽 및 2·3·5·6·11쪽 문구·줄바꿈·결론 강조 완수 / 넘침 0·래스터 0'
(Q / 'validation.json').write_text(json.dumps(r, ensure_ascii=False, indent=2), encoding='utf-8')

(W / 'README.md').write_text('''# 채용 과제 제출용 최신 자료
최신 v006 / 13장 / 2026-10-27
- [PPTX](repair2_채용과제_분석결과.pptx)
- [PDF](repair2_채용과제_분석결과.pdf)
- [검수](검수_v006/validation.json)
- 버전 보존: `versions/repair2_채용과제_분석결과_v006.pptx` 및 `.pdf` (이전 `v003`, `v004`, `v005` 보존)

## v006 주요 반영 내역
- **1순위 — 4페이지(날짜 검증과 처리 결정)**:
  - 상단 제목을 `날짜 이상을 확인하고 임의 보정 대신 분석 대상을 구분했다`, 부제목을 `행 순서에 따른 날짜 증가와 업무 시작·완료 시간 역전을 확인해 날짜 기록의 신뢰성을 검토`로 변환
  - 우측 판단 근거 박스에 `실제 발생 날짜와 오류 수정 규칙을 확인할 수 없어 임의로 연도를 치환하면 경과 시간이 왜곡될 수 있음`을 명확히 표기
- **2순위 — 7페이지(개별 업무 시간과 전체 시간 비교)**:
  - 하단 주석을 `업무 809쌍과 요청 339건은 분석 대상이 다르며 / 중앙값 비교만으로 관계를 단정할 수 없음`으로 바꾸고 상자 너비를 넓혀 자연스러운 2줄로 배치(한 글자 줄바꿈 현상 해결)
- **3순위 — 12페이지(개선 방안)**:
  - 3열(`03 · 확인 후 실행할 조치`)을 구체화:
    - 조사 시작 전 대기: `예정 시각 초과 사례를 확인하고 지연 사유에 따라 일정 재조정`
    - E 유형 수리 시작 전 대기: `반복되는 지연 원인을 확인하고 수리 착수 일정과 처리 절차 개선`
    - Web·Letter 일정 조율: `채널별 처리 과정과 담당자 기록을 점검하고 확인된 원인에 맞춰 개선`
- **4순위 — 13페이지(효과 검증과 분석 한계)**:
  - 부제목: `신뢰할 수 있는 시간 기록을 확보하고 수리 유형과 접수 채널을 구분하여 개선 전후의 처리 시간을 비교`
  - 우측 카드 내 중복 설명을 줄이고 `분석 한계`를 우측 하단 카드에 정돈
  - 하단 전폭 바에 `최종 결론`: `현장 조사 시작 전 대기를 우선 검토 대상으로 선정하고 실제 지연 원인 확인 후 전체 처리 시간의 감소 여부로 개선 효과를 검증` 강조
- **5순위 — 2·3·5·6·9·11페이지 문구 직관화 및 용어 통일**:
  - 2쪽 Q3·Q4(`수리 유형과 수행 업무에 따라 처리 시간이 다른가`, `접수 채널과 담당자 기록에 따라 처리 시간이 다른가`)
  - 3쪽(`날짜 오류와 관계없이 수행된 업무와 경로 확인`), 5쪽(`한 요청에서 여러 수리 업무나 동일 업무의 반복이 관찰됨`), 6쪽(`반복 작업 여부를 확인하기 위해 업무별 수행 횟수와 순서를 분석`)
  - 9쪽(`24시간은 분포 확인용 기준이며 정해진 처리 기한이 아님` — `운영 SLA` 표현을 직관적인 우리말로 교체)
  - 11쪽(`시작·완료 담당자 차이만으로 실제 업무 인수인계를 단정할 수 없음`)
''', encoding='utf-8')

(W / 'CHANGELOG.md').write_text('''# 변경 기록
## v006 / 2026-10-27
- 사용자 추가 피드백 반영:
  - 4쪽 상단 제목·부제목 변환 및 임의 보정 보류 판단 근거 명확화
  - 7쪽 하단 주석 문구 변경 및 상자 너비 조정으로 자연스러운 2줄 배치(전 슬라이드 텍스트 상자 의도치 않은 줄바꿈 0건 검증)
  - 9쪽 `운영 SLA가 아님` → `24시간은 분포 확인용 기준이며 정해진 처리 기한이 아님` 직관화
  - 12쪽 `확인 후 실행할 조치` 3개 항목 구체화
  - 13쪽 부제목 간결화, 우측 카드 설명 중복 제거, 하단 `최종 결론` 전폭 강조
  - 2·3·5·6·11쪽 문장 간결화 및 용어 통일
- 보존: `versions/repair2_채용과제_분석결과_v006.pptx` 및 `.pdf` (이전 `v003`, `v004`, `v005` 유지)
- 최신: `repair2_채용과제_분석결과.pptx` 및 `.pdf`

## v005 / 2026-10-27
- `repair2_채용과제_PPT_v3_최종수정사항.md` 전 항목 반영 (13쪽 3요소 정리, 4쪽 핵심 검증 근거 2가지 집중, 10쪽 표본 기준 명시, 전 슬라이드 하단 중복 정리)
- 보존: `versions/repair2_채용과제_분석결과_v005.pptx` 및 `.pdf`

## v004 / 2026-10-27
- 망고보드 템플릿(`TPL42397`, `TPL40009`) 및 기존 포트폴리오 디자인 체계(`portfolio-ppt-design`)를 결합하여 전 13쪽 레이아웃·카드·배지·표·다이어그램 전면 개편
- 보존: `versions/repair2_채용과제_분석결과_v004.pptx` 및 `.pdf`

## v003 / 2026-10-27
사용자가 제시한 3·4쪽 및 반복된 하단 텍스트 밀집 문제를 반영
내지 해석/주의 패널과 3·4쪽 정보 영역 재구성, 좌측 기준선 정리
''', encoding='utf-8')

h = R / '작업인계.md'
shutil.copy2(h, V / '작업인계_v005.md')
h.write_text('''# 채용 과제 작업 인계
갱신: 2026-10-27
- 포트폴리오와 전혀 다른 작업. 포트폴리오 파일·인계·스킬 원본을 갱신하지 않는다
- 전용 작업트리 C:/Users/rhd67/orca/workspaces/DA-repair/repair2-submission / 브랜치 roh-git/repair2-submission
- 현재 탭 연결 경로와 무관하게 도구 workdir는 위 경로를 명시
- 최신 제출용 v006 13장: 작업기록/2026-10-27/수리요청분석_채용과제제출용/README.md
- v006 반영: 4쪽 상단 제목·부제목 변환 및 보정 보류 근거 명확화, 7쪽 하단 주석 문구 및 줄바꿈 해결, 9쪽 SLA 용어를 직관적 표현(`정해진 처리 기한이 아님`)으로 교체, 12쪽 개선 조치 구체화, 13쪽 부제목 간결화 및 하단 `최종 결론` 강조, 2·3·5·6·11쪽 문구 다듬기 완수
- font Paperlogy/Pretendard, 딥틸(#087F75)·다크잉크(#192327)·민트(#E5F2EF)·쿨그레이(#F4F6F7) 유지. 모든 차트·생키·도형 편집 가능
- 필수 내용 유지: 시스템 포함 수행 주체 16종, 동일 InternRepair B25/E127/P8 비교, 조사·수리 미착수 지표 구분(산출식 및 636개 이벤트 비교는 슬라이드 노트에 보존)
- 1970년 정확성 미검증, 339/927건 대표성 보장 안 됨, 원인·개선효과 단정 금지
- 원본 노트북: C:/Users/rhd67/orca/workspaces/DA-repair/cormorant/repair2_분석.ipynb
- 가이드·상세 보고서·검증 근거·v001/v002는 작업기록/2026-10-26/수리요청분석_채용과제제출용에, v003/v004/v005/v006은 작업기록/2026-10-27/수리요청분석_채용과제제출용/versions에 보존
- v006 전체 13장 렌더 시각 검수(넘침 0·의도치 않은 줄바꿈 0·불일치 0·래스터 0) 및 최신 파일 2차 재열기 검수 완료
- 제출기한 2026-10-27 23:59. 메일 발송·커밋·푸시 미실행
- 다음: 사용자 검토 후 추가 수정사항 있으면 v007로 보존하고 PPTX/PDF 함께 갱신
''', encoding='utf-8')
print('v006 finalized and verified successfully')
