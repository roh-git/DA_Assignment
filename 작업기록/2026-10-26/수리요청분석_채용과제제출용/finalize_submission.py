from pathlib import Path
import pandas as pd,json,hashlib,shutil
from pptx import Presentation
import pymupdf as fitz
W=Path(__file__).resolve().parent;R=W.parents[2];Q=W/'검수_v001'
src=Path('C:/Users/rhd67/orca/workspaces/DA-repair/cormorant')
df=pd.read_csv(src/'repair2.csv')
df['timestamp']=pd.to_datetime(df['date']+' '+df['time'],format='%Y-%m-%d %H:%M',errors='coerce')
events=df[df.eventtype.isin(['start','complete'])]
cnt=events.groupby(['caseID','taskID','eventtype']).size().unstack(fill_value=0)
keys=cnt[(cnt.start==1)&(cnt.complete==1)].index
pairs=events.set_index(['caseID','taskID']).loc[keys].reset_index().pivot(index=['caseID','taskID'],columns='eventtype',values='timestamp')
pairs['minutes']=(pairs.complete-pairs.start).dt.total_seconds()/60
tp=pairs[pairs.start.dt.year.eq(1970)&pairs.complete.dt.year.eq(1970)&pairs.minutes.ge(0)]
ok=df.groupby('caseID').timestamp.agg(lambda x:x.notna().all() and x.dt.year.eq(1970).all())
general=set(df.loc[df.taskID.eq('TicketReady'),'caseID'])
ids=ok[ok].index.intersection(list(general))
dur=df[df.caseID.isin(ids)].groupby('caseID').timestamp.agg(['min','max'])
dur['hours']=(dur['max']-dur['min']).dt.total_seconds()/3600
arr=tp.xs('ArrangeSurvey',level='taskID').minutes
comparison=dur.join(arr.rename('arr'))
comp=comparison.groupby(comparison.arr.gt(60)).hours.agg(['count','median'])
assert len(df)==13262 and df.caseID.nunique()==1000
assert len(tp)==3151 and len(ids)==339
assert (pairs.minutes<0).sum()==226
assert len(arr)==809 and round(arr.median(),1)==4 and round(arr.quantile(.9),1)==404.8
assert comp.loc[True,'count']==75 and comp.loc[False,'count']==264
assert round(comp.loc[True,'median'],1)==78.2 and round(comp.loc[False,'median'],1)==77.2
sus=df[df.timestamp.dt.year.between(2020,2022)]
assert len(sus)==860 and sus.timestamp.dt.normalize().diff().dropna().eq(pd.Timedelta(days=1)).all()
sha=lambda p:hashlib.sha256(p.read_bytes()).hexdigest()
assert sha(W/'sources/repair2_분석.ipynb')==sha(src/'repair2_분석.ipynb')
checks={'source_csv_sha256':sha(src/'repair2.csv'),'source_notebook_sha256':sha(src/'repair2_분석.ipynb'),'raw_events':len(df),'cases':1000,'valid_task_pairs':len(tp),'general_time_cases':len(ids),'reversed_pairs':int((pairs.minutes<0).sum()),'arrange_median_min':arr.median(),'arrange_p90_min':arr.quantile(.9),'whole_duration_comparison':comp.round(1).to_dict(),'row_date_rule_matches':860}
(Q/'新規数値検証.json'.replace('新規数値検証','추가수치_원본재계산')).write_text(json.dumps(checks,ensure_ascii=False,indent=2),encoding='utf-8')
report=json.loads((Q/'validation.json').read_text(encoding='utf-8'))
assert not report['overflow'] and not report['mismatch'] and not report['bounds'] and report['raster_images']==0
report['visual_review']='13쪽 전체 렌더 검수 완료: 제목·본문·표·차트 라벨·페이지 번호 확인'
latest=[]
for ext in ['pptx','pdf']:
 source=W/f'versions/repair2_채용과제_분석결과_v001.{ext}'
 target=W/f'repair2_채용과제_분석결과.{ext}'
 assert not target.exists()
 shutil.copy2(source,target)
 assert sha(source)==sha(target)
 latest.append(target)
p=Presentation(latest[0]);d=fitz.open(latest[1])
assert len(p.slides)==len(d)==13
norm=lambda x:''.join(x.split())
for i,s in enumerate(p.slides):
 for sh in s.shapes:
  if sh.has_text_frame:assert norm(sh.text) in norm(d[i].get_text()),(i,sh.text)
text='\n'.join(pg.get_text() for pg in d)
for t in ['860/860','226/3,627','636개','78.2','77.2','70.8%','36.6%']:assert t in text,t
assert '분기' not in text and '분석 과제 발표' not in text
report['latest_second_check']={'status':'pass','pages':13,'pptx_sha256':sha(latest[0]),'pdf_sha256':sha(latest[1]),'version_latest_hash_match':True,'text_reopen_match':True}
(Q/'validation.json').write_text(json.dumps(report,ensure_ascii=False,indent=2),encoding='utf-8')
(W/'README.md').write_text('''# 수리 요청 프로세스 분석 — 채용 과제 제출용
최신: v001 / 13장 / 2026-10-26
- [PPTX](repair2_채용과제_분석결과.pptx)
- [PDF](repair2_채용과제_분석결과.pdf)
- [검수 기록](검수_v001/validation.json)
- [원본 수치 재계산](검수_v001/추가수치_원본재계산.json)
- 버전 보존: versions/repair2_채용과제_분석결과_v001.pptx 및 PDF.
- reference의 발표자료 v008은 참고본이며 이번 제출본과 다르다.
## 구성
01 분석 개요 / 02 목적과 검토 질문 / 03 데이터·표본 / 04 날짜 검증·보정 보류 /
05 업무 흐름 / 06 반복 업무 / 07 개별 업무와 전체 시간 /
08 구간 비중 / 09 조사 전 대기 분포 / 10 수리 유형 /
11 접수 채널·담당자 / 12 개선 제안 / 13 검증 지표·한계.
## 제출 관점
- 사전 등록 가설로 포장하지 않고 초기 질문과 탐색 후 추가 비교를 구분했다.
- 날짜 오류 근거와 실제 날짜 복원 가능성을 구분했다. 1970년을 실제 연도로 인정하지 않는다.
- 339건은 일반 경로의 36.6%이며 전체 대표성을 보장하지 않는다.
- 원인과 관련성, 정상 예약 대기와 실제 지연, 제안과 실행 성과를 구분했다.
- 차트·생키·비중 막대는 편집 가능한 PPT 요소다. 잘라 붙인 그래프 이미지가 없다.
## 디자인
Paperlogy 6 SemiBold 제목 / Pretendard 본문 / 청록 강조와 흰 내지·짙은 표지.
portfolio-ppt-design의 여백·폰트·렌더 검수, frontend-design의 내용별 구성,
ui-ux-pro-max의 숫자 직접 표기와 색 이외 구분, hallmark의 불필요한 카드·장식 억제를 적용했다.
웹 전용 CSS·키보드·애니메이션 검사는 정적 PPT에 적용하지 않았다.
## 작업 분리
DA-repair/repair2-submission 전용 작업트리와 roh-git/repair2-submission 브랜치.
포트폴리오 문서·디자인 스킬 원본·기존 발표 파일은 수정하지 않았다.
포트폴리오 공용 인계에서 과제 진행 블록만 제거했고 변경 전 인계는 reference에 보존했다.
현재 대화 탭의 UI 연결 경로 자체는 포트폴리오로 남아 있으므로, 재개 시 반드시 전용 작업트리 경로를 지정한다.
''',encoding='utf-8')
(W/'CHANGELOG.md').write_text('''# 변경 기록
## v001 / 2026-10-26
- 포트폴리오와 다른 DA-repair 저장소의 전용 작업트리에서 생성.
- 기존 발표용 10장에서 채용 과제 독자용 13장으로 재구성.
- 초기 검토 질문, 날짜 검증과 보정 보류, 업무 시간·전체 시간 비교 추가.
- 데이터 표본 조건과 대표성 한계, 개선별 측정 지표를 본문에 표시.
- 원본 CSV로 새 주요 수치를 재계산하고 노트북·상세 보고서와 대조.
- 13장 렌더 검수, 넘침0·텍스트 불일치0·이미지0, 최신 PPTX/PDF 재열기 및 버전 해시 일치 확인.
- 최신 파일: repair2_채용과제_분석결과.pptx / repair2_채용과제_분석결과.pdf.
- 버전 보존: versions/repair2_채용과제_분석결과_v001.pptx 및 PDF.
''',encoding='utf-8')
handoff=R/'작업인계.md'
shutil.copy2(handoff,W/'versions/작업인계_분리초기_v001.md')
handoff.write_text('''# 채용 과제 작업 인계
갱신: 2026-10-26
- 사용자 확정: 포트폴리오와 전혀 다른 작업. 포트폴리오 파일·스킬·인계·최신본에 반영하지 않는다.
- 전용 작업트리 C:/Users/rhd67/orca/workspaces/DA-repair/repair2-submission, 브랜치 roh-git/repair2-submission, 부모 없음.
- 현재 대화 탭 UI 경로는 포트폴리오일 수 있으므로 모든 도구 workdir를 위 경로로 지정한다.
- 최신 제출용 v001 13장: 작업기록/2026-10-26/수리요청분석_채용과제제출용/README.md.
- PPTX/PDF 모두 생성, 전체 렌더 검수·넘침0·텍스트불일치0·래스터이미지0, 최신 경로 재열기2차 확인 완료.
- 초기 질문, 날짜 검증·보정 보류, 개별 업무와 전체 시간 비교를 추가했다.
- 1970년을 실제 연도로 확정하지 않음. 860행 날짜 규칙·226쌍 역전·636개 서로 다른 업무의 시각 일치가 복원 불가 판단 근거.
- 전체 업무 구성1,000건/업무 시간3,151쌍/일반 시간339건(927건의36.6%). 대표성 미보장.
- 조율업무 중앙값4분/P90 404.8분; 조율1시간 초과75건 전체78.2h vs 이하264건77.2h. 원본 CSV 재계산 완료.
- 모든 시간은 기록상 탐색값. 관련성/원인과 예약 대기/실제 지연을 구분. 효과는 미실행 제안.
- 원본 노트북: C:/Users/rhd67/orca/workspaces/DA-repair/cormorant/repair2_분석.ipynb. sources 사본과 해시 일치.
- sources에 사용자 가이드·상세 보고서. reference에 기존 발표용v008·제작 코드·분리 전 포트폴리오 인계 보존.
- 새 제작코드 build_submission.py, 검수·최신생성 finalize_submission.py. 이미 존재하는 v001은 덮어쓰지 않도록 코드 assert 적용.
- 포트폴리오 공용 인계의 현재 수리 분석 블록만 백업 후 제거. 기존 포트폴리오 파일과 기존 v008 보존.
- 제출기한2026-10-27 23:59. 메일 발송·커밋·푸시 미실행.
- 다음: 사용자의 검토 수정사항 반영. 변경 시 다음 버전으로 보존, PPTX/PDF 동시 갱신.
''',encoding='utf-8')
print(json.dumps({'status':'complete','pages':13,'latest':[str(f) for f in latest],'second_check':'pass','new_numbers':'recomputed from source CSV'},ensure_ascii=False))
