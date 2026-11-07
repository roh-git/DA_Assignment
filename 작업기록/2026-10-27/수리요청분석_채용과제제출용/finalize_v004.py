from pathlib import Path
import json, shutil, hashlib
from pptx import Presentation
import pymupdf as fitz

W = Path(__file__).resolve().parent
R = W.parents[2]
V = W / 'versions'
Q = W / '검수_v004'

sha = lambda f: hashlib.sha256(f.read_bytes()).hexdigest()

for ext in ['pptx', 'pdf']:
    src = V / f'repair2_채용과제_분석결과_v004.{ext}'
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

r = json.loads((Q / 'validation.json').read_text(encoding='utf-8'))
assert not r['overflow'] and not r['text_mismatch'] and not r['shape_bounds'] and r['raster_images'] == 0
r['latest_second_check'] = {
    'status': 'pass',
    'version_hash_match': True,
    'text_reopen_match': True,
    'charts': charts,
    'pptx_sha256': sha(W / 'repair2_채용과제_분석결과.pptx'),
    'pdf_sha256': sha(W / 'repair2_채용과제_분석결과.pdf')
}
r['visual_review'] = '전 13쪽 실제 PDF 렌더 및 시트 검수 완료 / 망고보드(TPL42397·TPL40009) 및 포트폴리오 카드·배지·영역 분리 체계 반영 / 넘침 0·래스터 0'
(Q / 'validation.json').write_text(json.dumps(r, ensure_ascii=False, indent=2), encoding='utf-8')

(W / 'README.md').write_text('''# 채용 과제 제출용 최신 자료
최신 v004 / 13장 / 2026-10-27
- [PPTX](repair2_채용과제_분석결과.pptx)
- [PDF](repair2_채용과제_분석결과.pdf)
- [검수](검수_v004/validation.json)
- 버전 보존: `versions/repair2_채용과제_분석결과_v004.pptx` 및 `.pdf` (이전 `v003` 보존)

## 디자인 전면 개편 (v004)
- **레퍼런스 반영**: 망고보드 템플릿 2종(`TPL42397` 카드·배지·스텝 구조, `TPL40009` 딥틸·세이지·슬레이트 배색 및 KPI·3단 연결 매트릭스) + 사용자 포트폴리오(`portfolio-ppt-design`)의 컨테이너 카드·내부 헤더 배너·필(Pill) 라벨 체계 통합
- **공통 레이아웃 구조**:
  - 상단 카테고리 필 배지(`02 · 분석 목적과 검토 질문`) 및 우측 페이지 배지(`Page 02`) + 제목/부제 하단 구분선으로 헤더와 본문 영역 명확히 분리
  - 모든 본문 요소를 둥근 컨테이너 카드(`#F4F6F7`)와 내부 강조 배너(`#E5F2EF` / `#192327`) 및 필 배지로 구조화하여 맨바닥 텍스트 나열 제거
  - 하단 해석/주의 영역을 둥근 필 배지(`해석` / `해석 시 주의`)가 포함된 카드 형태로 정돈
- **장별 주요 시각 개선**:
  - `01` 표지: 다크 딥틸 배경 + 상단 포인트 바 + 3대 분석 범위(`1,000건` / `927건` / `339건`) 벤토 카드 배치
  - `02` 검토 질문: `Q1~Q4` 4열 세로 스텝 카드 + 하단 다크 4단계 분석 진행 순서 바
  - `03` 데이터 개요: 좌측 데이터 구조·결측 요약 카드(`13개` 업무 / `16종` 수행 주체) + 우측 3단계 표본 선별 카드
  - `04` 날짜 검증: 좌측 3대 검증 흰색 내부 카드 + 우측 다크 히어로 결정 카드(`원본 유지`) 및 목적별 표본 카드
  - `05` 업무 흐름: 좌측 생키 컨테이너 카드 + 우측 3단계(접수·조사 / 수리 3종 / 완료·종결) 카드형 프로세스 다이어그램
  - `06` 반복 작업: 상단 좌우 포트폴리오형 구성 카드(단일 836·복합 91 비중 막대 + 반복 38건) + 하단 `caseID 124` 타임라인 카드
  - `07` 개별 vs 전체 시간: 좌우 대칭 포트폴리오 카드(내부 헤더 배너 + `4분` vs `404.8분` 서브카드 + `78.2h` vs `77.2h` 비교 막대 카드)
  - `08` 구간 시간 비중: 상단 5개 마일스톤 흐름 및 4개 구간 비중 카드(`70.8%` 딥틸 강조) + 하단 다크 100% 누적 비중 스트립
  - `09~11` 심층 분석: 좌측 차트/표 컨테이너 카드 + 우측 히어로 KPI 스택 카드 및 담당자 비교 카드
  - `12` 개선 방안: `TPL40009` 15쪽 스타일의 3열 연결형 매트릭스(`01 개선 검토 대상` → `02 추가로 기록할 정보` → `03 확인 후 실행할 조치`)
  - `13` 효과 검증과 한계: 좌측 신규 수집 지표 필-로우 카드 및 산출식 박스 + 우측 최종 성과 지표 다크 카드 및 `01~05` 한계 배지 리스트
''', encoding='utf-8')

(W / 'CHANGELOG.md').write_text('''# 변경 기록
## v004 / 2026-10-27
- 망고보드 템플릿(`TPL42397`, `TPL40009`) 및 기존 포트폴리오 디자인 체계(`portfolio-ppt-design`)를 결합하여 전 13쪽 레이아웃·카드·배지·표·다이어그램 전면 개편
- 전 13쪽 실제 PDF 렌더 시각 검수, 텍스트 넘침 0건, 문구 불일치 0건, 래스터 이미지 0건(네이티브 차트·도형 100%), 최신 경로 재열기 2차 검수 완료
- 보존: `versions/repair2_채용과제_분석결과_v004.pptx` 및 `.pdf` (이전 `v003` 유지)
- 최신: `repair2_채용과제_분석결과.pptx` 및 `.pdf`

## v003 / 2026-10-27
사용자가 제시한 3·4쪽 및 반복된 하단 텍스트 밀집 문제를 반영
내지 해석/주의 패널과 3·4쪽 정보 영역 재구성, 좌측 기준선 정리
전체 렌더·넘침0·내용 불일치0·이미지0, 최신 파일 재열기 및 버전 해시 확인
이전 날짜의 v001·v002 및 26일 제작 중 v003 보존
''', encoding='utf-8')

h = R / '작업인계.md'
shutil.copy2(h, V / '작업인계_v003.md')
h.write_text('''# 채용 과제 작업 인계
갱신: 2026-10-27
- 포트폴리오와 전혀 다른 작업. 포트폴리오 파일·인계·스킬 원본을 갱신하지 않는다
- 전용 작업트리 C:/Users/rhd67/orca/workspaces/DA-repair/repair2-submission / 브랜치 roh-git/repair2-submission
- 현재 탭 연결 경로와 무관하게 도구 workdir는 위 경로를 명시
- 최신 제출용 v004 13장: 작업기록/2026-10-27/수리요청분석_채용과제제출용/README.md
- v004 반영: 망고보드(TPL42397·TPL40009) 및 사용자 포트폴리오 디자인 체계를 결합해 전 13장 컨테이너 카드·헤더 배너·필(Pill) 배지·3열 연결 매트릭스·KPI 카드로 전면 리디자인
- font Paperlogy/Pretendard, 딥틸(#087F75)·다크잉크(#192327)·민트(#E5F2EF)·쿨그레이(#F4F6F7) 유지. 모든 차트·생키·도형 편집 가능
- 필수 내용 유지: 시스템 포함 수행 주체 16종, 동일 InternRepair B25/E127/P8 비교, 조사·수리 미착수 지표 구분
- 1970년 정확성 미검증, 339/927건 대표성 보장 안 됨, 원인·개선효과 단정 금지
- 원본 노트북: C:/Users/rhd67/orca/workspaces/DA-repair/cormorant/repair2_분석.ipynb
- 가이드·상세 보고서·검증 근거·v001/v002는 작업기록/2026-10-26/수리요청분석_채용과제제출용에, v003/v004는 작업기록/2026-10-27/수리요청분석_채용과제제출용/versions에 보존
- v004 전체 13장 렌더 시각 검수(넘침 0·불일치 0·래스터 0) 및 최신 파일 2차 재열기 검수 완료
- 제출기한 2026-10-27 23:59. 메일 발송·커밋·푸시 미실행
- 다음: 사용자 검토 후 추가 수정사항 있으면 v005로 보존하고 PPTX/PDF 함께 갱신
''', encoding='utf-8')
print('v004 finalized and verified successfully')
