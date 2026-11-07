from pathlib import Path
import hashlib,json,shutil
from pptx import Presentation
import pymupdf as fitz
W=Path(__file__).resolve().parent;R=W.parents[2];V=W/'versions';Q=W/'검수_v002'
sha=lambda x:hashlib.sha256(x.read_bytes()).hexdigest()
r=json.loads((Q/'validation.json').read_text(encoding='utf-8'))
assert not r['overflow'] and not r['text_mismatch'] and not r['shape_bounds']
assert r['raster_images']==0
for ext in ['pptx','pdf']:
 old=W/f'repair2_채용과제_분석결과.{ext}'
 assert sha(old) in {sha(V/f'repair2_채용과제_분석결과_v001.{ext}'),sha(V/f'repair2_채용과제_분석결과_v002.{ext}')}
 shutil.copy2(V/f'repair2_채용과제_분석결과_v002.{ext}',old)
 assert sha(old)==sha(V/f'repair2_채용과제_분석결과_v002.{ext}')
p=Presentation(W/'repair2_채용과제_분석결과.pptx');d=fitz.open(W/'repair2_채용과제_분석결과.pdf')
assert len(d)==len(p.slides)==13
norm=lambda s:''.join(s.split())
for i,s in enumerate(p.slides):
 for sh in s.shapes:
  if sh.has_text_frame:assert norm(sh.text) in norm(d[i].get_text()),(i,sh.text)
full='\n'.join(pg.get_text() for pg in d)
assert '16명' not in full and '분기' not in full
assert norm('업무 수행 주체 16종') in norm(d[2].get_text())
for t in ['25건','127건','8건','0.0시간','25.9시간','0.9시간','11.6','160.4','71.4']:assert norm(t) in norm(d[9].get_text()),t
for t in ['조사 예정 시각 초과 미착수율','수리 예정 시각 초과 미착수율','× 100','중앙값 및 P90']:assert norm(t) in norm(d[12].get_text()),t
charts=[]
for i,sl in enumerate(p.slides):
 for sh in sl.shapes:
  if sh.has_chart:charts.append({'page':i+1,'series':[list(se.values) for se in sh.chart.series]})
assert charts==[{'page':9,'series':[[39.0,89.0,63.0,90.0,58.0]]},{'page':10,'series':[[11.6,160.4,71.4]]}],charts
r.update({'visual_review':'13쪽 전체 렌더 검수 완료 / 추가 수정 5·10·13쪽 재검수 완료','latest_second_check':{'status':'pass','pages':13,'pptx_pdf_text_match':True,'version_latest_sha256_match':True,'pptx_sha256':sha(W/'repair2_채용과제_분석결과.pptx'),'pdf_sha256':sha(W/'repair2_채용과제_분석결과.pdf')},'native_chart_values':charts,'checklist_required_items':'03 수행 주체·선별 설명 / 10 B·E·P 비교·표본 / 13 조사·수리 지표 및 최종 성과 반영'})
(Q/'validation.json').write_text(json.dumps(r,ensure_ascii=False,indent=2),encoding='utf-8')
(W/'최종점검_반영내역_v002.md').write_text('''# 최종점검 반영 및 디자인 검수 / v002
기준: sources/repair2_채용과제_PPT_최종점검_수정사항.md
13장 구성과 기존 청록색·흰색·회색 체계를 유지했다
## 내용 반영
| 페이지 | 반영 사항 |
|---|---|
| 03 | 업무 수행 주체 16종(시스템 기록 포함)으로 수정, 3,631→3,151쌍 선별 설명 간소화, 339/927건(36.6%) 제한을 본문으로 이동 |
| 04 | 날짜 순서·시작과 완료·복원 가능성의 질문을 수치와 연결, 연도 치환 및 삭제 보류 이유 유지 |
| 05 | 일반·예외 경로 연결선 대비 강화, 예외 안내 화살표가 상자에 가리지 않도록 수정 |
| 06 | 복합91건·반복38건 요약, 29건 시간순 확인 조건 유지, 재작업 단정 금지 |
| 07 | 78.2h와77.2h의 기술 비교로 한정, 관련성 부재·인과 결론을 내리지 않음을 본문에 표시 |
| 08–09 | 70.8% 합계 비중과 절감 가능 비율 구분, 대기 분포 수치·탐색 기준 유지 |
| 10 | InternRepair 단독·1:1 연결160건의 B25건/0.0h·E127건/25.9h·P8건/0.9h 비교표, P 소표본 주의 본문에 표시 |
| 11 | 채널·담당자 비교 수치 유지, 인수인계·역량 차이로 단정하지 않는 해석 유지 |
| 12 | 예정 초과 미착수 확인·사유별 일정 조정·지연 단계 개선으로 행동 구체화 |
| 13 | 조사와 수리 예정 시각 초과 미착수율 분리, 신규 수집 후 계산·분모/분자 조건·×100 표기, 최종 성과인 전체 시간 중앙값/P90 강조 |
## 디자인
- portfolio-ppt-design: Paperlogy/Pretendard, 일정한 좌우 기준선, 실제 렌더 기반 정렬 검수
- frontend-design: 데이터·판단·해석의 읽는 순서를 분리하고 내용에 맞춘 표·차트 구성 유지
- ui-ux-pro-max: 색뿐 아니라 숫자·유형·표본 직접 표기, 차트 숫자14pt 및 핵심 제한사항 본문화
- hallmark: 새 템플릿 도입 없이 기존 디자인을 다듬고 불필요한 장식·중첩 카드를 억제
- 전체 페이지의 해석 영역과 출처 영역을 구분, 단계명12pt·제목28pt, 표와 본문 간격 조정
- 각주10.5pt와 출처9.5pt는 보조 정보에 한정하며 핵심 제한사항은 본문14.5pt로 표시
- 10쪽 비교표의 유형·건수·시간 정렬, 13쪽 최종 성과 패널 크기 조정
- 웹 전용 CSS/인터랙션 검사는 정적 PPT에 적용하지 않음
## 근거·검수
- originator 원본 목록에 System 포함, 결측 제외 고유값16개 확인
- 노트북294셀 출력으로 동일InternRepair 비교 수치 대조
- 신규 수치 외 기존 수치는 v001 검증 결과 재사용, native chart 값과 최신 PDF 표시 재확인
- 13쪽 전체 렌더 확인, 수정 후 5·10·13쪽 재확인
- 넘침0 / 텍스트 불일치0 / 슬라이드 밖 도형0 / PDF 래스터 이미지0
- 최신 파일 재열기와 v002 보존본 해시 일치 2차 확인
- 기존v001과 원본 가이드·노트북·포트폴리오 파일 보존
''',encoding='utf-8')
readme=W/'README.md';shutil.copy2(readme,V/'README_v001.md')
text=readme.read_text(encoding='utf-8').replace('최신: v001','최신: v002').replace('검수_v001/validation.json','검수_v002/validation.json').replace('versions/repair2_채용과제_분석결과_v001.pptx','versions/repair2_채용과제_분석결과_v002.pptx')
text+='\n## v002 수정\n[최종점검 반영·디자인 검수](최종점검_반영내역_v002.md) / 기존v001 보존\n'
readme.write_text(text,encoding='utf-8')
with (W/'CHANGELOG.md').open('a',encoding='utf-8') as f:f.write('''
## v002 / 2026-10-26
- 최종점검 가이드 필수03·10·13쪽 및 권장04·06·07·12쪽 반영
- 전체 단계·제목·해석·출처 위계 조정, 표본·중요 제한사항 본문화, 표 정렬·차트 수치·흐름도 대비 보강
- v001 PPTX/PDF 보존, v002를 최신 경로에 반영
- 13장 렌더 및 최신 파일2차 검수 통과, 상세: 최종점검_반영내역_v002.md
- 최신: repair2_채용과제_분석결과.pptx 및 PDF
''')
handoff=R/'작업인계.md';shutil.copy2(handoff,V/'작업인계_제출용_v001.md')
t=handoff.read_text(encoding='utf-8').replace('최신 제출용 v001 13장','최신 제출용 v002 13장')
t+='\n## 최종점검 반영 v002\n- 사용자 최종점검 MD를 sources에 보존, 필수/권장 수정과 전체 간격·위계·정렬·가독성 검수 완료\n- 수행 주체16종(시스템 포함), InternRepair 동일조건 B25/E127/P8 비교, 조사·수리 미착수율 구분, 전체 시간 중앙값/P90 최종 성과 강조\n- 13쪽 전체 렌더, 보완한5·10·13쪽 추가 확인, 최신 PPTX/PDF 재열기2차 검수 통과\n- 코드 revise_v002.py → finish_layout_v002.py → polish_flow_v002.py, 검수/최신화 finalize_v002.py. 기존 버전을 덮어쓰지 않도록 다음 수정은 새 번호 사용\n- 상세 반영: 작업기록/2026-10-26/수리요청분석_채용과제제출용/최종점검_반영내역_v002.md\n'
handoff.write_text(t,encoding='utf-8')
print(json.dumps({'version':2,'pages':13,'latest_second_check':'pass','checklist':'pass','charts':charts},ensure_ascii=False))

