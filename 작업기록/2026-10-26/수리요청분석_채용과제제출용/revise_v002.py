from pathlib import Path
import ast,json,re,hashlib,shutil
from copy import deepcopy
from pptx import Presentation
from pptx.util import Inches,Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE,MSO_CONNECTOR
from pptx.enum.text import PP_ALIGN,MSO_ANCHOR
from pptx.enum.chart import XL_CHART_TYPE,XL_LABEL_POSITION,XL_LEGEND_POSITION,XL_TICK_MARK
from pptx.chart.data import CategoryChartData
from pptx.oxml.xmlchemy import OxmlElement
import pymupdf as fitz
import win32com.client
W=Path(__file__).resolve().parent;Q=W/'검수_v002';Q.mkdir(exist_ok=True)
V=W/'versions';out=V/'repair2_채용과제_분석결과_v002.pptx';pdf=out.with_suffix('.pdf')
assert not out.exists()
p=Presentation(V/'repair2_채용과제_분석결과_v001.pptx')
m=ast.parse((W/'reference/build_v004.py').read_text(encoding='utf-8-sig'))
ns=[n for n in m.body if (isinstance(n,ast.FunctionDef) and n.name in {'rgb','rect','txt','line','node','bar','dot','ribbon'}) or (isinstance(n,ast.Assign) and any(isinstance(t,ast.Name) and t.id in {'C','F'} for t in n.targets))]
exec(compile(ast.Module(body=ns,type_ignores=[]),'helpers','exec'))
def rewrite(sh,t):
 tf=sh.text_frame;par=tf.paragraphs[0]
 rp=deepcopy(par.runs[0]._r.rPr) if par.runs and par.runs[0]._r.rPr is not None else None
 tf.clear();par=tf.paragraphs[0]
 for i,text in enumerate(t.split('\n')):
  x=par if i==0 else tf.add_paragraph()
  if i and par._p.pPr is not None:x._p.insert(0,deepcopy(par._p.pPr))
  r=x.add_run();r.text=text
  if rp is not None:r._r.insert(0,deepcopy(rp))
def size(sh,n):
 for par in sh.text_frame.paragraphs:
  for run in par.runs:run.font.size=Pt(n)
def change(sl,needle,text):
 for sh in sl.shapes:
  if sh.has_text_frame and needle in sh.text:rewrite(sh,text);return sh
 raise ValueError(needle)
def bodyclear(sl):
 for sh in list(sl.shapes):
  if Inches(2.4)<=sh.top<Inches(6.2):
   sh._element.getparent().remove(sh._element)
def conclusion(sl,t):
 for sh in sl.shapes:
  if sh.has_text_frame and abs(sh.top/Inches(1)-6.22)<.02:rewrite(sh,t);return
def foot(sl,t):
 for sh in sl.shapes:
  if sh.has_text_frame and abs(sh.top/Inches(1)-6.8)<.02:rewrite(sh,t);return
def heading(sl,t):
 for sh in sl.shapes:
  if sh.has_text_frame and abs(sh.top/Inches(1)-.68)<.02:rewrite(sh,t);return
def intent(sl,t):
 for sh in sl.shapes:
  if sh.has_text_frame and abs(sh.top/Inches(1)-1.79)<.02:rewrite(sh,t);return
# 03: simplify selection and accurately identify actors.
s=p.slides[2];bodyclear(s)
intent(s,'caseID는 요청 1건이고 taskID는 수행 업무를 뜻함\nDate·Time을 결합하고 start·complete 기록을 연결해 업무별 경과 시간을 계산')
txt(s,.75,2.5,11.8,.36,'업무 종류 13개 · 업무 수행 주체 16종(시스템 기록 포함)',18,bold=True)
txt(s,.75,2.97,11.8,.3,'날짜·시각 결측 13행 / 수리 속성 결측은 예외 경로에서 확인',14,color='muted')
rect(s,.75,3.48,11.8,.43,'ink')
for x,w,t in [(.88,2.65,'분석 대상'),(3.7,1.6,'규모'),(5.45,6.85,'선별 기준')]:txt(s,x,3.54,w,.28,t,14,color='white',bold=True)
for i,(a,b,c) in enumerate([
 ('업무 구성·경로','1,000건','원본 이벤트 보존 / 업무 존재와 구성 확인'),
 ('개별 업무 시간','3,151쌍','시작·완료 각 1회 / 모두 1970년 / 완료 ≥ 시작'),
 ('일반 경로 전체 시간','339건','모든 시각 존재·1970년인 402건에서 예외 63건 제외')]):
 y=4.04+i*.59
 if i==2:rect(s,.75,y-.06,11.8,.53,'pale')
 txt(s,.88,y,2.65,.36,a,15.5,bold=True)
 txt(s,3.7,y,1.6,.36,b,19,bold=True)
 txt(s,5.45,y,6.85,.36,c,14.5)
txt(s,.75,5.84,11.8,.3,'시작·완료가 각 1회인 3,631쌍 중 시각 조건을 충족한 3,151쌍을 분석',15)
conclusion(s,'1970년은 실제 연도로 확인된 값이 아님\n339건은 일반 경로 927건의 36.6%로 전체 일반 요청을 대표한다고 단정할 수 없음')
foot(s,'업무 구성 1,000건과 시간 분석 339건은 서로 다른 범위 / 업무 쌍의 시간 역전 검증은 다음 페이지')
# 04: make numerical evidence answer clear questions.
s=p.slides[3];bodyclear(s)
intent(s,'여러 연도가 섞인 기록에서 날짜 규칙과 시간 역전을 점검\n이상 패턴은 확인했지만 실제 발생 연도와 날짜 변환 규칙은 확인하지 못함')
txt(s,.75,2.5,7.15,.35,'원래 날짜를 복원할 수 있는가',18,bold=True)
line(s,8.0,2.5,8.0,6.02)
txt(s,8.35,2.5,4.2,.35,'대안과 처리 결정',18,bold=True)
for i,(question,n,desc) in enumerate([
 ('날짜가 실제 업무 순서를 반영하는가','860/860행','2020~2022년 날짜가 CSV 행 순서대로 하루씩 증가'),
 ('업무 시작·완료 순서가 맞는가','226/3,627쌍','비교 가능한 업무 중 완료가 시작보다 빠른 기록 존재'),
 ('같은 시각이면 날짜도 복원할 수 있는가','636개 이벤트','다른 연도 기록과 시각은 같지만 모두 서로 다른 업무')]):
 y=3.06+i*.98
 txt(s,.75,y,7.0,.32,question,15.5,bold=True)
 txt(s,.75,y+.4,2.4,.43,n,22,color='accent',bold=True)
 txt(s,3.25,y+.36,4.35,.56,desc,14.5)
for i,(a,b) in enumerate([
 ('연도 일괄 치환 보류','1970 → 2020으로 바꿀\n실제 날짜·변환 규칙 근거가 없음'),
 ('이상 날짜 행 삭제 보류','업무 구성과 반복 기록까지\n함께 사라지므로 원본 유지'),
 ('목적별 표본으로 나눠 분석','업무 구성 1,000건 보존\n시간은 3,151쌍·일반 경로 339건')]):
 y=3.06+i*.98
 txt(s,8.35,y,4.15,.32,a,16,bold=True,color='accent' if i==2 else 'ink')
 txt(s,8.35,y+.38,4.15,.58,b,14.5)
conclusion(s,'날짜 이상을 발견한 것만으로 실제 날짜를 확정할 수는 없음\n원본을 보존하고 시간 계산을 제한했으며 1970년 기록의 정확성도 미확인')
foot(s,'연도별 이벤트: 1970년 12,383행 / 2020–2022년 860행 / 1905년 6행 / 날짜·시각 결측 13행')
# 06: plain summary while retaining conditions and example.
s=p.slides[5]
conclusion(s,'일반 경로 927건 중 복합 수리 91건과 InternRepair 반복 38건을 확인\n반복 원인과 재작업 여부는 알 수 없어 다음으로 업무 시간과 전체 시간을 비교')
sh=change(s,'InternRepair 시각이 모두','InternRepair 시각이 모두 1970년인\n29건을 시간순으로 확인\n28건에서 완료 후 다시 시작')
size(sh,15)
# 07: descriptive comparison, not a causal claim.
s=p.slides[6]
heading(s,'조율 업무가 1시간을 넘는 요청에서도\n전체 처리 시간 중앙값 차이는 작았다')
conclusion(s,'조율 1시간 초과 집단은 78.2시간으로 비교 집단의 77.2시간과 큰 차이가 없었음\n중앙값 비교만으로 관련성이나 인과관계를 판단하지 않고 업무 사이 대기도 확인')
foot(s,'업무 시간 809쌍과 전체 시간 339건은 별도 표본 / 분과 시간을 직접 합산하지 않음')
# 08: calculation definition moved to visible interpretation.
s=p.slides[7]
conclusion(s,'70.8%는 339건의 전체 시간 합계에서 이 구간이 차지하는 비중이며 절감 가능 비율이 아님\n정상적인 예약 대기일 수 있으므로 다음으로 대기의 분포를 확인')
foot(s,'구간 시간 합계 ÷ 전체 시간 합계 / 반올림 합계 100.1% / 위 흐름도의 간격은 시간 비례가 아님')
# 10: compare all B/E/P under identical repair conditions.
s=p.slides[9];bodyclear(s)
intent(s,'수리 유형별 전체 처리 시간을 비교한 뒤 수리 시작 전 대기를 확인\n업무 구성의 차이를 줄이기 위해 InternRepair(내부 수리)만 수행한 요청끼리 추가 비교')
txt(s,.75,2.5,7.05,.38,'전체 처리 시간 중앙값 / 339건 / 시간',17,bold=True)
bar(s,.75,3.08,7.05,2.95,['B (46건)','E (143건)','P (150건)'],[('중앙값',[11.6,160.4,71.4],'bar')],200,fmt='0.0',accents={0:['bar','accent','bar']})
line(s,8.05,2.5,8.05,6.04)
txt(s,8.4,2.5,4.15,.34,'수리 시작 전 대기 / 338건',16,bold=True)
txt(s,8.4,2.91,4.15,.28,'Survey complete → 첫 수리 start',12.5,color='muted')
txt(s,8.4,3.28,4.15,.46,'E 25.2시간',27,color='accent',bold=True)
txt(s,8.4,3.83,4.15,.31,'중앙값 / B 0.0 · E 25.2 · P 0.0시간',13)
txt(s,8.4,4.16,4.15,.26,'B 46건 / E 143건 / P 149건',12.5,color='muted')
line(s,8.4,4.55,12.55,4.55)
txt(s,8.4,4.57,4.15,.32,'InternRepair 단독·1:1 연결 / 160건',14.5,bold=True)
for x,w,t in [(8.4,.8,'유형'),(9.3,1,'건수'),(10.45,2.1,'대기 중앙값')]:txt(s,x,4.96,w,.25,t,12.5,color='muted',align=PP_ALIGN.RIGHT if x>8.4 else None)
for i,(a,n,v) in enumerate([('B','25건','0.0시간'),('E','127건','25.9시간'),('P','8건','0.9시간')]):
 y=5.29+i*.27
 if i==1:rect(s,8.32,y-.02,4.3,.27,'pale')
 txt(s,8.4,y,.8,.24,a,14,bold=i==1)
 txt(s,9.3,y,1,.24,n,14,bold=i==1,align=PP_ALIGN.RIGHT)
 txt(s,10.45,y,2.1,.24,v,14,bold=i==1,align=PP_ALIGN.RIGHT)
conclusion(s,'같은 InternRepair만 수행한 요청에서도 E 유형의 긴 대기가 관찰됨\n수리 유형이 원인이라는 결론은 아니며 P 유형은 8건으로 표본이 적어 해석에 주의')
foot(s,'표와 우측 수치는 수리 착수 전 대기 / InternRepair start→complete 수행 시간과 구분 / B·E·P 의미는 미확인')
# 12: improve actionable table with generous row spacing.
s=p.slides[11];bodyclear(s)
heading(s,'예정 시각과 진행 상태를 기록하고\n실제 지연이 확인된 단계부터 개선한다')
intent(s,'기록상 긴 대기를 곧바로 불필요한 지연으로 보지 않고 확인할 정보와 실행할 조치를 연결')
rect(s,.75,2.55,11.8,.49,'ink')
for x,w,t in [(.9,2.85,'개선 검토 대상'),(4.05,3.65,'추가로 기록할 정보'),(8.08,4.28,'확인 후 실행할 조치')]:txt(s,x,2.65,w,.28,t,15,color='white',bold=True)
rows=[
 ('조사 시작 전 대기\n전체 시간 합계의 70.8%','조사 예정 시각·착수 상태\n예정 초과 사유','예정 시각을 넘긴 미착수 요청 확인\n사유에 따라 조사 일정 조정'),
 ('E 유형 수리 시작 전 대기\n중앙값 25.2시간','조사 완료 이후의 진행 상태\n수리 예정 시각·지연 사유','반복적인 지연 사유를 확인\n해당 단계부터 절차 개선'),
 ('Web·Letter 일정 조율\n1시간 초과 194건','채널별 업무 진행 상태\n시작·완료 담당자 기록 차이','기록 차이와 실제 지연 여부 확인\n지연이 나타난 처리 절차 점검')]
for i,(a,b,c) in enumerate(rows):
 y=3.31+i*.92
 if i==0:rect(s,.75,y-.1,11.8,.82,'pale')
 for x,w,t in [(.9,2.85,a),(4.05,3.65,b),(8.08,4.28,c)]:txt(s,x,y,w,.61,t,15)
 if i<2:line(s,.75,y+.8,12.55,y+.8)
conclusion(s,'제안은 실행 전 검토안이며 자재 부족이나 인력 과부하를 확인된 원인으로 가정하지 않음\n원인을 확인한 뒤 시범 적용하고 구간별 시간과 요청 전체의 처리 시간을 함께 비교')
foot(s,'예정 시각·진행 상태·지연 사유는 신규 수집 항목 / 우선 검토 순서는 절감 효과가 입증된 순서가 아님')
# 13: distinguish survey/repair measures and show a final outcome hierarchy.
s=p.slides[12];bodyclear(s)
heading(s,'예정 시각과 지연 사유를 새로 기록하고\n대기와 전체 완료 시간이 함께 줄었는지 검증한다')
intent(s,'정확한 시각 기록으로 새 기준선을 마련한 뒤 시범 개선 전후를 비교\n수리 유형과 접수 채널의 구성 차이도 확인해 비교 가능성을 점검')
txt(s,.75,2.5,7.2,.35,'신규 수집 후 계산할 지표',18,bold=True)
for i,(a,b) in enumerate([
 ('현장 조사 일정','조사 예정 시각 초과 미착수율 · 착수 지연 시간'),
 ('E 유형 수리 착수','수리 예정 시각 초과 미착수율 · 수리 전 대기'),
 ('Web·Letter 일정 조율','ArrangeSurvey 1시간 초과율 · 처리 시간 P90')]):
 y=3.08+i*.73
 txt(s,.75,y,7.2,.28,a,13.5,color='muted')
 txt(s,.75,y+.32,7.2,.35,b,16,bold=True)
txt(s,.75,5.47,7.2,.58,'미착수율(%) = 예정 시각이 지났지만 미착수인 요청 수\n÷ 조회 시점까지 시작 예정이었던 전체 요청 수 × 100',12.5)
line(s,8.1,2.5,8.1,6.03)
rect(s,8.45,2.5,4.1,1.42,'ink')
txt(s,8.65,2.68,3.7,.29,'최종 성과 지표',14,color='white')
txt(s,8.65,3.06,3.7,.72,'케이스 전체 처리 시간\n중앙값 및 P90',19,color='white',bold=True)
txt(s,8.45,4.08,4.1,.32,'해석의 한계',16,bold=True)
txt(s,8.45,4.54,4.1,1.42,'1970년 날짜의 실제 정확성 미확인\n339건은 일반 경로의 36.6%\n예정 시각이 없어 실제 지연 구분 불가\n외부 수리(ExternRepair) 완료 시각 없음\n1시간·24시간은 탐색 기준이며 SLA 아님',13.5)
conclusion(s,'조사 시작 전 대기의 비중은 컸지만 실제 지연 원인은 확인되지 않음\n예정 시각과 지연 사유를 기록한 뒤 개선 전후의 대기와 전체 완료 시간을 함께 비교')
foot(s,'미착수율: 아직 예정 시각이 안 된 요청은 분모에서 제외 / 이미 착수한 요청은 미착수 건수에서 제외')
s.notes_slide.notes_text_frame.text='조사와 수리 지표는 각각 조사 예정 시각과 수리 예정 시각을 기준으로 따로 계산한다. 조사·수리 착수 지연 시간은 max(실제 착수−예정 착수,0)이며 조회 시점 미착수 건은 완료된 지연 시간과 구분한다. 예정 시각 초과 미착수율(%)=(조회 시점에 예정 시각이 지났지만 아직 시작하지 않은 요청 수 / 조회 시점까지 시작 예정이었던 전체 요청 수)×100. 예정 시각이 미래인 요청은 분모에서 제외하고 이미 시작한 요청은 분자에서 제외한다. 예정 시각은 기존 로그에 없어 신규 수집 후 계산한다. 유형·채널 구성과 관찰기간을 맞춰 전후 비교하며 인과 효과로 단정하지 않는다.'
# Whole-deck polish: retain body layouts, align interpretation/source bands, improve notes.
for idx,s in enumerate(p.slides):
 if idx==0:continue
 for sh in s.shapes:
  if sh.has_text_frame:
   t=sh.text
   # Sentence-ending punctuation only; retain decimal points and thousands commas.
   t=re.sub(r'(?<!\d)\.(?=\s|$)','',t)
   t=t.replace('뒤,','뒤').replace('다. ','다 ')
   if t!=sh.text:rewrite(sh,t)
   y=sh.top/Inches(1)
   if abs(y-.25)<.02:size(sh,12)
   if abs(y-.68)<.02:size(sh,28)
   if abs(y-6.22)<.02:
    sh.left=Inches(.75);sh.top=Inches(6.32);sh.height=Inches(.51);size(sh,14.5)
   if abs(y-6.8)<.02:
    sh.left=Inches(.75);sh.width=Inches(11.8);sh.top=Inches(6.91);sh.height=Inches(.2);size(sh,10.5)
   if y>7.1:
    sh.top=Inches(7.23) if sh.left<Inches(12) else Inches(7.2)
  elif abs(sh.top/Inches(1)-7.08)<.02:
   sh.top=Inches(7.16)
 line(s,.75,6.2,12.55,6.2,'rule',1)
 # Improve native chart labels without increasing every body shape.
 for sh in s.shapes:
  if sh.has_chart:
   chart=sh.chart
   chart.category_axis.tick_labels.font.size=Pt(14)
   chart.value_axis.tick_labels.font.size=Pt(12)
   for plot in chart.plots:
    if plot.has_data_labels:plot.data_labels.font.size=Pt(14)
 # Visible body is authoritative; keep supporting notes synchronized.
 if idx!=12:
  s.notes_slide.notes_text_frame.text='검토용 근거와 해석\n'+'\n'.join(sh.text for sh in s.shapes if sh.has_text_frame)+'\n시간은 기록상 탐색 값이며 실제 날짜·인과관계·개선 효과를 확정하지 않는다'
 for sh in s.shapes:
  for ef in sh._element.xpath('.//a:effectRef'):ef.set('idx','0')
  for sp in sh._element.xpath('./p:spPr'):
   for el in list(sp):
    if el.tag.endswith('}effectLst') or el.tag.endswith('}effectDag'):sp.remove(el)
   sp.append(OxmlElement('a:effectLst'))
p.save(out)
app=win32com.client.Dispatch('PowerPoint.Application');deck=app.Presentations.Open(str(out),False,False,False)
overflow=[]
try:
 for sl in deck.Slides:
  for sh in sl.Shapes:
   if sh.HasTextFrame and sh.TextFrame.HasText:
    tr=sh.TextFrame.TextRange
    if tr.BoundHeight>sh.Height+2 or tr.BoundWidth>sh.Width+2:overflow.append({'page':sl.SlideIndex,'text':tr.Text,'h':sh.Height,'bh':tr.BoundHeight,'w':sh.Width,'bw':tr.BoundWidth})
 deck.Save();deck.SaveAs(str(pdf),32)
finally:deck.Close()
pp=Presentation(out);d=fitz.open(pdf);norm=lambda x:''.join(x.split());mismatch=[];bounds=[]
for i,sl in enumerate(pp.slides):
 for sh in sl.shapes:
  if sh.has_text_frame and norm(sh.text) not in norm(d[i].get_text()):mismatch.append([i+1,sh.text])
  if sh.left<0 or sh.top<0 or sh.left+sh.width>pp.slide_width+1000 or sh.top+sh.height>pp.slide_height+1000:bounds.append([i+1,sh.name])
 d[i].get_pixmap(matrix=fitz.Matrix(1.5,1.5)).save(Q/f'page-{i+1:02}.png')
for start in range(0,13,4):
 sheet=fitz.open();pg=sheet.new_page(width=1600,height=900)
 for j,i in enumerate(range(start,min(start+4,13))):pg.show_pdf_page(fitz.Rect(j%2*800,j//2*450,(j%2+1)*800,(j//2+1)*450),d,i)
 pg.get_pixmap().save(Q/f'sheet-{start+1:02}.png')
report={'version':2,'pages':len(d),'overflow':overflow,'text_mismatch':mismatch,'shape_bounds':bounds,'raster_images':sum(len(pg.get_images()) for pg in d),'fonts':sorted({x[3] for pg in d for x in pg.get_fonts()}),'visual_review':'pending','latest_second_check':'pending'}
(Q/'validation.json').write_text(json.dumps(report,ensure_ascii=False,indent=2),encoding='utf-8')
print(json.dumps(report,ensure_ascii=False,indent=2))
