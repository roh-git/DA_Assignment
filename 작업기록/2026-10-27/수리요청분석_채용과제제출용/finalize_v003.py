from pathlib import Path
import json,shutil,hashlib
from pptx import Presentation
from pptx.util import Inches
import pymupdf as fitz
import win32com.client
W=Path(__file__).resolve().parent;R=W.parents[2]
old=R/'작업기록/2026-10-26/수리요청분석_채용과제제출용'
Q=W/'검수_v003';V=W/'versions';out=V/'repair2_채용과제_분석결과_v003.pptx';pdf=out.with_suffix('.pdf')
assert not out.exists()
p=Presentation(old/'versions/repair2_채용과제_분석결과_v003.pptx')
# Keep both segments of the 836/91 composition bar contiguous.
for sh in p.slides[5].shapes:
 if abs(sh.top/Inches(1)-3.35)<.02 and sh.left>Inches(6) and sh.left<Inches(7) and abs(sh.height/Inches(1)-.64)<.02:
  sh.left-=Inches(.1)
p.save(out)
qa=(old/'revise_v002.py').read_text(encoding='utf-8-sig')
qa=qa[qa.index("app=win32com.client.Dispatch"):].replace("'version':2","'version':3")
exec(compile(qa,'qa_v003','exec'))
r=json.loads((Q/'validation.json').read_text(encoding='utf-8'))
assert not r['overflow'] and not r['text_mismatch'] and not r['shape_bounds'] and r['raster_images']==0
sha=lambda f:hashlib.sha256(f.read_bytes()).hexdigest()
for ext in ['pptx','pdf']:
 src=V/f'repair2_채용과제_분석결과_v003.{ext}';dst=W/f'repair2_채용과제_분석결과.{ext}'
 assert not dst.exists()
 shutil.copy2(src,dst);assert sha(src)==sha(dst)
pp=Presentation(W/'repair2_채용과제_분석결과.pptx');doc=fitz.open(W/'repair2_채용과제_분석결과.pdf')
assert len(pp.slides)==len(doc)==13
norm=lambda t:''.join(t.split())
for i,s in enumerate(pp.slides):
 for sh in s.shapes:
  if sh.has_text_frame:assert norm(sh.text) in norm(doc[i].get_text())
charts=[list(se.values) for sl in pp.slides for sh in sl.shapes if sh.has_chart for se in sh.chart.series]
assert charts==[[39,89,63,90,58],[11.6,160.4,71.4]]
r['latest_second_check']={'status':'pass','version_hash_match':True,'text_reopen_match':True,'charts':charts}
r['visual_review']='13쪽 렌더 확인 / 3·4쪽 영역 재구성 / 내지 해석·주의 구분 / 6쪽 비중 막대 연결 간격 마지막 확인 예정'
(Q/'validation.json').write_text(json.dumps(r,ensure_ascii=False,indent=2),encoding='utf-8')
(W/'README.md').write_text('''# 채용 과제 제출용 최신 자료
최신 v003 / 13장 / 2026-10-27
- [PPTX](repair2_채용과제_분석결과.pptx)
- [PDF](repair2_채용과제_분석결과.pdf)
- [검수](검수_v003/validation.json)
- 버전 보존: versions/repair2_채용과제_분석결과_v003.pptx 및 PDF
## 디자인 변경
- frontend-design: 3쪽 정보·표본·해석 순서, 4쪽 검증/결정 영역을 분리
- ui-ux-pro-max: 콘텐츠에 맞춘 상자 높이, 직접 수치 표기, 여백과 읽는 순서 점검
- hallmark: 근거 없이 장식을 추가하지 않고 정보 역할에 맞는 영역 구분 적용
- portfolio-ppt-design: 기존 폰트·색상 유지, 바깥쪽 기준선 통일과 실제 PDF 렌더 검수
- 내지 하단의 긴 텍스트를 청록 해석 영역과 회색 주의 영역으로 재구성
- 3쪽 업무 종류/수행 주체 요약과 넓은 표 행, 4쪽 검증 근거와 원본 유지 결정을 구분
- 자세한 기존 해석·주석은 발표자 노트에도 보존
## 보존 자료
- v001·v002와 원본 근거는 ../../2026-10-26/수리요청분석_채용과제제출용 에 보존
- 제작: 10월26일 revise_v003.py·finish_v003.py, 당일 finalize_v003.py
- 포트폴리오와 별도인 DA-repair/repair2-submission 작업트리에서만 관리
''',encoding='utf-8')
(W/'CHANGELOG.md').write_text('''# 변경 기록
## v003 / 2026-10-27
사용자가 제시한 3·4쪽 및 반복된 하단 텍스트 밀집 문제를 반영
내지 해석/주의 패널과 3·4쪽 정보 영역 재구성, 좌측 기준선 정리
전체 렌더·넘침0·내용 불일치0·이미지0, 최신 파일 재열기 및 버전 해시 확인
이전 날짜의 v001·v002 및 26일 제작 중 v003 보존
최신: repair2_채용과제_분석결과.pptx 및 PDF
''',encoding='utf-8')
h=R/'작업인계.md';shutil.copy2(h,V/'작업인계_v002.md')
h.write_text('''# 채용 과제 작업 인계
갱신: 2026-10-27
- 포트폴리오와 전혀 다른 작업. 포트폴리오 파일·인계·스킬 원본을 갱신하지 않는다
- 전용 작업트리 C:/Users/rhd67/orca/workspaces/DA-repair/repair2-submission / 브랜치 roh-git/repair2-submission
- 현재 탭 연결 경로와 무관하게 도구 workdir는 위 경로를 명시
- 최신 제출용 v003 13장: 작업기록/2026-10-27/수리요청분석_채용과제제출용/README.md
- 사용자 피드백: 텍스트 나열과 하단 밀집, 기준선 불일치. 디자인 스킬 적극 적용 요청
- 3쪽 정보·표본 구분, 4쪽 검증 근거/처리 결정 분리, 내지 하단 해석/주의 패널, 문장 축약과 정렬 반영
- font Paperlogy/Pretendard, 청록/회색/흰색 유지. 차트·도형 편집 가능
- v002 필수 내용 유지: 시스템 포함 수행 주체16종, 동일InternRepair B25/E127/P8 비교, 조사·수리 미착수 지표 구분
- 1970년 정확성 미검증, 339/927건 대표성 보장 안 됨, 원인·개선효과 단정 금지
- 원본 노트북: C:/Users/rhd67/orca/workspaces/DA-repair/cormorant/repair2_분석.ipynb
- 가이드·상세 보고서·검증 근거·v001/v002·기존 제작코드는 작업기록/2026-10-26/수리요청분석_채용과제제출용에 보존
- v003 전체 렌더 및 최신 파일2차 재열기 통과. 최종 6쪽 막대 간격 수정 확인 뒤 검수 기록 완료
- 제출기한2026-10-27 23:59. 메일 발송·커밋·푸시 미실행
- 다음: 사용자 검토 수정사항 반영 시 v004로 보존하고 PPTX/PDF 함께 갱신
''',encoding='utf-8')
print('v003 latest saved and reopened')
