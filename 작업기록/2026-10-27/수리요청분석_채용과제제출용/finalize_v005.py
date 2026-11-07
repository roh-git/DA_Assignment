from pathlib import Path
import json, shutil, hashlib
from pptx import Presentation
import pymupdf as fitz

W = Path(__file__).resolve().parent
R = W.parents[2]
V = W / 'versions'
Q = W / '검수_v005'

sha = lambda f: hashlib.sha256(f.read_bytes()).hexdigest()

for ext in ['pptx', 'pdf']:
    src = V / f'repair2_채용과제_분석결과_v005.{ext}'
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

# Verify checklist items from repair2_채용과제_PPT_v3_최종수정사항.md
p_texts = [norm(doc[i].get_text()) for i in range(13)]
assert norm('SCOPE 02 · 일반 경로') in p_texts[0]
assert norm('860건') in p_texts[3] and norm('226쌍') in p_texts[3] and norm('636') not in p_texts[3]
assert norm('1970년 기록도 실제 날짜의 정확성이 검증된 것은 아님') in p_texts[3]
assert norm('단일 또는 복합 가능 · 반복 수행 관찰') in p_texts[4]
assert norm('복합 수리 91건과 InternRepair 반복 38건은 두 집단이 겹칠 수 있어 합산 불가') in p_texts[5]
assert norm('809 업무쌍') in p_texts[6] and norm('339 케이스') in p_texts[6]
assert norm('70.8%는 절감 가능 비율이 아니며 정상적인 예약 대기와 실제 운영 지연은 현재 기록만으로 구분 불가') in p_texts[7]
assert norm('24시간은 탐색 기준이며 운영 SLA가 아님') in p_texts[8]
assert norm('E 유형 전체 · 수리 시작 전 대기') in p_texts[9] and norm('143건 · 중앙값 25.2시간') in p_texts[9]
assert norm('같은 내부 수리(InternRepair)만 수행한 요청 비교') in p_texts[9] and norm('25.9시간') in p_texts[9]
assert norm('담당자가 다르다는 기록이 실제 인수인계의 증거는 아님') in p_texts[10]
assert norm('수리 요청 전체 처리 시간') in p_texts[12] and norm('중앙값 · P90') in p_texts[12]
assert norm('신규 지표는 추가 데이터 수집 이후 계산 가능하며 개선은 아직 실행·검증되지 않음') in p_texts[12]
assert norm('날짜 정확성 미검증 · 예정 시각 미기록 · 시간 분석 대상 339건(일반 경로 927건의 36.6% · 외부 수리 완료 시각 없음)') in p_texts[12]

r = json.loads((Q / 'validation.json').read_text(encoding='utf-8'))
assert not r['overflow'] and not r['text_mismatch'] and not r['shape_bounds'] and r['raster_images'] == 0
r['latest_second_check'] = {
    'status': 'pass',
    'version_hash_match': True,
    'text_reopen_match': True,
    'checklist_verified': True,
    'charts': charts,
    'pptx_sha256': sha(W / 'repair2_채용과제_분석결과.pptx'),
    'pdf_sha256': sha(W / 'repair2_채용과제_분석결과.pdf')
}
r['visual_review'] = '전 13쪽 실제 PDF 렌더 및 시트 검수 완료 / repair2_채용과제_PPT_v3_최종수정사항.md 전 항목(4·10·13쪽 중점 정리 및 하단 중복 제거·용어 통일) 반영 / 넘침 0·래스터 0'
(Q / 'validation.json').write_text(json.dumps(r, ensure_ascii=False, indent=2), encoding='utf-8')

(W / 'README.md').write_text('''# 채용 과제 제출용 최신 자료
최신 v005 / 13장 / 2026-10-27
- [PPTX](repair2_채용과제_분석결과.pptx)
- [PDF](repair2_채용과제_분석결과.pdf)
- [검수](검수_v005/validation.json)
- 버전 보존: `versions/repair2_채용과제_분석결과_v005.pptx` 및 `.pdf` (이전 `v003`, `v004` 보존)
- 수정 지침 보존: `sources/repair2_채용과제_PPT_v3_최종수정사항.md`

## 최종 수정사항 반영 (v005)
- **전체 원칙**: `v004`의 13장 구성·컨테이너 카드·필(Pill) 배지·색상 체계를 유지하면서, `repair2_채용과제_PPT_v3_최종수정사항.md`의 3대 우선 정리 대상(13쪽·4쪽·10쪽)과 하단 중복 문장 정리를 반영
- **우선순위 1 — 13페이지(효과 검증과 분석 한계) 3요소 정리**:
  - ① 좌측(`개선 과제별 핵심 측정 지표`): 산출식 박스를 발표 노트·상세 보고서로 옮기고 3개 과제별 측정 지표와 필수 전제(`신규 지표는 추가 데이터 수집 이후 계산 가능하며 개선은 아직 실행·검증되지 않음`)를 선명하게 배치
  - ② 우측(`최종 성과 지표`): 다크 히어로 카드를 전체 높이로 확대하여 `수리 요청 전체 처리 시간` / `중앙값 · P90` 강조
  - ③ 하단 1줄(`분석 한계`): 5개 항목을 핵심 1줄(`날짜 정확성 미검증 · 예정 시각 미기록 · 시간 분석 대상 339건(일반 경로 927건의 36.6% · 외부 수리 완료 시각 없음)`)로 압축
- **우선순위 2 — 4페이지(날짜 검증과 처리 결정) 핵심 검증 2가지 집중**:
  - 좌측: `636개 이벤트` 비교를 발표 노트·상세 보고서로 옮기고 가장 직관적인 핵심 근거 2가지(`860건: 2020~2022년 날짜가 CSV 행 순서대로 하루씩 증가`, `226쌍: 시작보다 완료 시각이 먼저 기록된 업무가 존재`)만 큼직하게 배치
  - 우측: `판단(실제 날짜와 오류 수정 규칙을 확인할 수 없어 연도 일괄 치환을 보류)` → `처리(원본은 유지하고 분석 목적에 따라 데이터 분리: 업무 흐름 1,000건 · 개별 업무 시간 3,151쌍 · 일반 경로 시간 339건)` 순으로 정리
  - 하단: 우측 박스와 중복되던 결론 문구를 제거하고 `해석 시 주의(1970년 기록도 실제 날짜의 정확성이 검증된 것은 아님)`만 1줄로 유지
- **우선순위 3 — 10페이지(수리 유형별 비교) 표본 기준 명시**:
  - 우측 상단 카드 제목을 `E 유형 전체 · 수리 시작 전 대기`로, 수치를 `143건 · 중앙값 25.2시간`으로 명시
  - 우측 하단 표 제목을 `같은 내부 수리(InternRepair)만 수행한 요청 비교`(`B 25건 0.0시간 / E 127건 25.9시간 / P 8건 0.9시간`)로 구분해 `25.2시간`과 `25.9시간`의 표본 차이를 즉시 이해할 수 있게 보완
- **전체 슬라이드 다듬기**:
  - 1쪽 표지 `SCOPE 02 · 일반 경로`로 용어 통일
  - 2·3·5·6·7·8·9·11·12쪽 하단 박스에서 본문 카드 숫자를 그대로 반복하거나 다음 슬라이드 예고를 적던 문구를 걷어내고, 각 슬라이드에서 꼭 필요한 해석·주의사항만 1줄 바로 정돈
''', encoding='utf-8')

(W / 'CHANGELOG.md').write_text('''# 변경 기록
## v005 / 2026-10-27
- `repair2_채용과제_PPT_v3_최종수정사항.md` 전 항목 반영:
  - 13쪽(산출식 노트 이동·최종 성과 지표 히어로 카드 확대·분석 한계 하단 1줄 압축)
  - 4쪽(핵심 검증 근거 2가지 `860건`·`226쌍` 집중 배치, 우측 판단·데이터 분리 정리, 하단 중복 제거)
  - 10쪽(`E 유형 전체 143건 · 중앙값 25.2시간` vs `같은 내부 수리(InternRepair)만 수행한 요청 비교 127건 · 25.9시간` 표본 기준 명시)
  - 1쪽(`SCOPE 02 · 일반 경로` 용어 통일) 및 2·3·5·6·7·8·9·11·12쪽 하단 중복 문구 정리·1줄 통합
- 전 13쪽 실제 PDF 렌더 시각 검수, 텍스트 넘침 0건, 문구 불일치 0건, 래스터 이미지 0건(네이티브 차트·도형 100%), 최신 경로 재열기 2차 검수 완료
- 보존: `versions/repair2_채용과제_분석결과_v005.pptx` 및 `.pdf` (이전 `v003`, `v004` 유지)
- 최신: `repair2_채용과제_분석결과.pptx` 및 `.pdf`

## v004 / 2026-10-27
- 망고보드 템플릿(`TPL42397`, `TPL40009`) 및 기존 포트폴리오 디자인 체계(`portfolio-ppt-design`)를 결합하여 전 13쪽 레이아웃·카드·배지·표·다이어그램 전면 개편
- 전 13쪽 실제 PDF 렌더 시각 검수, 텍스트 넘침 0건, 문구 불일치 0건, 래스터 이미지 0건(네이티브 차트·도형 100%), 최신 경로 재열기 2차 검수 완료
- 보존: `versions/repair2_채용과제_분석결과_v004.pptx` 및 `.pdf` (이전 `v003` 유지)

## v003 / 2026-10-27
사용자가 제시한 3·4쪽 및 반복된 하단 텍스트 밀집 문제를 반영
내지 해석/주의 패널과 3·4쪽 정보 영역 재구성, 좌측 기준선 정리
전체 렌더·넘침0·내용 불일치0·이미지0, 최신 파일 재열기 및 버전 해시 확인
이전 날짜의 v001·v002 및 26일 제작 중 v003 보존
''', encoding='utf-8')

h = R / '작업인계.md'
shutil.copy2(h, V / '작업인계_v004.md')
h.write_text('''# 채용 과제 작업 인계
갱신: 2026-10-27
- 포트폴리오와 전혀 다른 작업. 포트폴리오 파일·인계·스킬 원본을 갱신하지 않는다
- 전용 작업트리 C:/Users/rhd67/orca/workspaces/DA-repair/repair2-submission / 브랜치 roh-git/repair2-submission
- 현재 탭 연결 경로와 무관하게 도구 workdir는 위 경로를 명시
- 최신 제출용 v005 13장: 작업기록/2026-10-27/수리요청분석_채용과제제출용/README.md
- v005 반영: v004의 컨테이너 카드·배지·컬러 체계를 유지하고 `repair2_채용과제_PPT_v3_최종수정사항.md` 지침(13쪽 3요소 정리 및 최종 성과 지표 확대, 4쪽 핵심 근거 860건·226쌍 집중 및 하단 중복 제거, 10쪽 E유형 전체 143건 25.2h vs 동일 InternRepair 127건 25.9h 표본 구분, 1쪽 일반 경로 용어 통일, 전 슬라이드 하단 중복 문구 1줄 정리) 완수
- font Paperlogy/Pretendard, 딥틸(#087F75)·다크잉크(#192327)·민트(#E5F2EF)·쿨그레이(#F4F6F7) 유지. 모든 차트·생키·도형 편집 가능
- 필수 내용 유지: 시스템 포함 수행 주체 16종, 동일 InternRepair B25/E127/P8 비교, 조사·수리 미착수 지표 구분(산출식 및 636개 이벤트 비교는 슬라이드 노트에 보존)
- 1970년 정확성 미검증, 339/927건 대표성 보장 안 됨, 원인·개선효과 단정 금지
- 원본 노트북: C:/Users/rhd67/orca/workspaces/DA-repair/cormorant/repair2_분석.ipynb
- 가이드·상세 보고서·검증 근거·v001/v002는 작업기록/2026-10-26/수리요청분석_채용과제제출용에, v003/v004/v005는 작업기록/2026-10-27/수리요청분석_채용과제제출용/versions에 보존
- v005 전체 13장 렌더 시각 검수(넘침 0·불일치 0·래스터 0) 및 최신 파일 2차 재열기 검수 완료
- 제출기한 2026-10-27 23:59. 메일 발송·커밋·푸시 미실행
- 다음: 사용자 검토 후 추가 수정사항 있으면 v006으로 보존하고 PPTX/PDF 함께 갱신
''', encoding='utf-8')
print('v005 finalized and verified successfully')
