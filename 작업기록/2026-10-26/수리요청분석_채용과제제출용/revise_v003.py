from pathlib import Path
import ast,json,re
from copy import deepcopy
from pptx import Presentation
from pptx.util import Inches,Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE,MSO_CONNECTOR
from pptx.enum.text import PP_ALIGN,MSO_ANCHOR
from pptx.enum.chart import XL_CHART_TYPE,XL_LABEL_POSITION,XL_LEGEND_POSITION,XL_TICK_MARK
from pptx.chart.data import CategoryChartData
from pptx.oxml.xmlchemy import OxmlElement
W=Path(__file__).resolve().parent;V=W/'versions';Q=W/'검수_v003';Q.mkdir(exist_ok=True)
out=V/'repair2_채용과제_분석결과_v003.pptx';pdf=out.with_suffix('.pdf')
assert not out.exists()
p=Presentation(V/'repair2_채용과제_분석결과_v002.pptx')
m=ast.parse((W/'reference/build_v004.py').read_text(encoding='utf-8-sig'))
ns=[n for n in m.body if (isinstance(n,ast.FunctionDef) and n.name in {'rgb','rect','txt','line','node','bar','dot','ribbon'}) or (isinstance(n,ast.Assign) and any(isinstance(t,ast.Name) and t.id in {'C','F'} for t in n.targets))]
exec(compile(ast.Module(body=ns,type_ignores=[]),'helpers','exec'))
def rewrite(sh,t):
 tf=sh.text_frame;par=tf.paragraphs[0];rp=deepcopy(par.runs[0]._r.rPr) if par.runs and par.runs[0]._r.rPr is not None else None
 tf.clear()
 for i,text in enumerate(t.split('\n')):
  x=tf.paragraphs[0] if i==0 else tf.add_paragraph()
  r=x.add_run();r.text=text
  if rp is not None:r._r.insert(0,deepcopy(rp))
def clearbody(s):
 for sh in list(s.shapes):
  if Inches(2.4)<=sh.top<Inches(6.2):sh._element.getparent().remove(sh._element)
def intent(s,t):
 for sh in s.shapes:
  if sh.has_text_frame and abs(sh.top/Inches(1)-1.79)<.02:rewrite(sh,t)
# 03: explicit metric group and roomy sample table
s=p.slides[2];clearbody(s)
intent(s,'요청(caseID)과 업무(taskID)를 기준으로 기록을 정리하고 시간 계산에 필요한 표본을 선별')
for x,n,label in [(.65,'13개','업무 종류'),(6.8,'16종','업무 수행 주체 / 시스템 포함')]:
 rect(s,x,2.48,5.88,.73,'surface')
 txt(s,x+.2,2.62,1.25,.48,n,26,color='accent',bold=True)
 txt(s,x+1.65,2.65,4,.32,label,16,bold=True)
rect(s,.65,3.48,12.03,.45,'ink')
for x,w,t in [(.85,2.75,'분석 대상'),(3.9,1.55,'규모'),(5.7,6.7,'선별 기준')]:txt(s,x,3.57,w,.26,t,14,color='white',bold=True)
for i,(a,b,c) in enumerate([
 ('업무 구성·경로','1,000건','원본 이벤트 보존\n날짜와 무관한 업무 존재·구성 확인'),
 ('개별 업무 시간','3,151쌍','시작·완료가 각 1회인 3,631쌍에서 선별\n모두 1970년 기록이며 완료 시각 ≥ 시작 시각'),
 ('일반 경로 전체 시간','339건','모든 시각이 존재하고 1970년인 402건\n이 중 예외 경로 63건 제외')]):
 y=4.08+i*.66
 if i==2:rect(s,.65,y-.05,12.03,.61,'pale')
 txt(s,.85,y,2.75,.48,a,16,bold=True)
 txt(s,3.9,y,1.55,.48,b,22,bold=True)
 txt(s,5.7,y-.02,6.7,.53,c,14.5)
txt(s,.65,6.02,12.03,.23,'Date·Time 결합 후 시작·완료 시각 연결 / 날짜·시각 결측 13행 / 수리 속성 결측은 예외 경로에서 확인',11.5,color='muted')
# 04: two distinct zones, evidence versus a single chosen action
s=p.slides[3];clearbody(s)
intent(s,'비정상적인 날짜 패턴을 검증했지만 실제 발생 연도와 날짜 변환 규칙은 확인하지 못함')
rect(s,.65,2.48,7.25,3.48,'surface')
rect(s,8.15,2.48,4.53,3.48,'pale')
txt(s,.9,2.68,6.75,.34,'날짜 복원 가능성 검증',18,bold=True)
for i,(label,num,desc) in enumerate([
 ('행 순서 검사','860/860행','2020~2022년 날짜가\nCSV 행 순서대로 하루씩 증가'),
 ('시작·완료 검사','226/3,627쌍','완료가 시작보다 빠른 기록\n비교 가능한 업무 쌍에서 확인'),
 ('같은 시각 비교','636개 이벤트','다른 연도 기록과 시각은 같지만\n모두 서로 다른 업무')]):
 y=3.22+i*.86
 txt(s,.9,y,2.35,.25,label,13,color='muted')
 txt(s,.9,y+.29,2.35,.4,num,22,color='accent',bold=True)
 txt(s,3.55,y+.14,4.07,.55,desc,15)
 if i<2:line(s,.9,y+.75,7.65,y+.75,'rule')
txt(s,8.4,2.68,4.03,.34,'처리 결정',18,bold=True)
txt(s,8.4,3.21,4.03,.48,'원본 유지',28,color='accent',bold=True)
txt(s,8.4,3.87,4.03,.64,'연도 치환과 행 삭제는 보류\n실제 날짜를 복원할 근거가 없음',15)
line(s,8.4,4.72,12.43,4.72,'light')
txt(s,8.4,4.91,4.03,.28,'목적별로 나눠 사용',14,bold=True)
txt(s,8.4,5.32,4.03,.56,'업무 구성 1,000건\n시간 3,151쌍 / 일반 경로 339건',15)
txt(s,.65,6.02,12.03,.23,'연도별 이벤트: 1970년 12,383행 / 2020–2022년 860행 / 1905년 6행 / 날짜·시각 결측 13행',11.5,color='muted')
# All inner slides: replace small text footer with labeled, padded summary/caution panels.
summaries=[
 ('검토 질문에서 출발해 데이터 품질부터 확인','E 유형·특정 채널 비교는 탐색 후 추가\n사전 등록된 가설은 아님'),
 ('업무 기록은 전체 보존\n시간 계산만 조건에 맞는 표본으로 제한','1970년 실제 연도는 미확인\n339건의 전체 대표성은 보장 안 됨'),
 ('날짜 복원 근거가 없어 원본을 유지\n이후 분석은 목적별 표본으로 진행','연도 치환·행 삭제 보류\n1970년 기록도 정확성 미확인'),
 ('일반 경로에서도 복합 수리와 반복 업무가 관찰됨','업무 흐름의 구조 요약\n종결 관련 두 업무의 순서는 고정 아님'),
 ('복합 수리 91건 · InternRepair 반복 38건\n다음으로 업무 시간과 전체 시간을 비교','두 집단은 겹칠 수 있어 합산 불가\n반복 원인·재작업 여부는 미확인'),
 ('전체 시간 중앙값 78.2h vs 77.2h\n업무 사이의 대기도 추가 확인','중앙값만으로 관계를 단정하지 않음\n809쌍과 339건은 별도 표본'),
 ('조사 전 대기가 전체 시간 합계의 70.8%\n다음으로 대기의 분포를 확인','절감 가능 비율이 아님\n예약 대기와 실제 지연 구분 불가'),
 ('24시간 초과 211/339건(62.2%)\n이어서 수리 유형별 차이를 비교','24시간은 탐색 기준\n운영 SLA가 아님'),
 ('동일 InternRepair 조건에서도\nE 유형의 긴 착수 전 대기가 관찰됨','P 유형은 8건으로 해석에 주의\n수리 유형을 원인으로 단정 불가'),
 ('접수 채널과 담당자 기록에 따른 차이 확인\n처리 방식과 담당자가 다른 이유를 추가 점검','담당자 차이가 인수인계의 증거는 아님\n실제 지연 원인은 미확인'),
 ('원인 확인 후 시범 적용\n구간 시간과 전체 처리 시간을 함께 비교','신규 기록 수집이 선행\n자재·인력 문제로 단정하지 않음'),
 ('예정 시각·지연 사유를 기록한 뒤\n대기와 전체 완료 시간의 변화를 비교','현재는 실제 지연 원인 미확인\n개선 효과는 아직 검증 전')]
for idx,s in enumerate(p.slides):
 if idx==0:continue
 oldnotes=s.notes_slide.notes_text_frame.text
 for sh in list(s.shapes):
  y=sh.top/Inches(1)
  if 6.18<=y<7.19:
   sh._element.getparent().remove(sh._element)
   continue
  if sh.has_text_frame:
   # Outer text grid aligns with the title; retain interior card padding.
   if abs(sh.left/Inches(1)-.75)<.015 or abs(sh.left/Inches(1)-.8)<.015:
    sh.left=Inches(.65)
   if abs(y-.25)<.02:sh.left=Inches(.65)
 insight,caution=summaries[idx-1]
 rect(s,.65,6.38,7.65,.68,'pale')
 rect(s,.65,6.38,.045,.68,'accent')
 txt(s,.84,6.57,.65,.25,'해석',12,color='accent',bold=True)
 txt(s,1.6,6.44,6.42,.56,insight,15,bold=True)
 rect(s,8.5,6.38,4.18,.68,'surface')
 txt(s,8.69,6.43,3.8,.2,'해석 시 주의',10.5,color='muted',bold=True)
 txt(s,8.69,6.66,3.8,.33,caution,11.5)
 # Wider right caution panel text needs 2 lines; preserve 12pt where possible.
 for sh in list(s.shapes):
  if sh.has_text_frame and sh.text==caution:
   sh.top=Inches(6.61);sh.height=Inches(.42)
   for par in sh.text_frame.paragraphs:
    for r in par.runs:r.font.size=Pt(12)
 # Source remains isolated beneath the summary, with matched outer alignment.
 for sh in s.shapes:
  if sh.has_text_frame and sh.top>Inches(7.19):
   if sh.left<Inches(12):sh.left=Inches(.65)
 s.notes_slide.notes_text_frame.text=oldnotes+'\n\n가독성 개정 요약\n'+insight+'\n해석 시 주의\n'+caution
 for sh in s.shapes:
  for ef in sh._element.xpath('.//a:effectRef'):ef.set('idx','0')
  for sp in sh._element.xpath('./p:spPr'):
   for el in list(sp):
    if el.tag.endswith('}effectLst') or el.tag.endswith('}effectDag'):sp.remove(el)
   sp.append(OxmlElement('a:effectLst'))
p.save(out)
# Reuse the already validated export and geometry/text checks, with this version's paths.
qa=(W/'revise_v002.py').read_text(encoding='utf-8-sig')
qa=qa[qa.index("app=win32com.client.Dispatch"):].replace("'version':2","'version':3")
import pymupdf as fitz
import win32com.client
exec(compile(qa,'qa_v003','exec'))
