from pathlib import Path
import ast,json,re,hashlib,shutil
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
R=Path(__file__).resolve().parents[3]
W=Path(__file__).resolve().parent
Q=W/'검수_v001';Q.mkdir(exist_ok=True)
V=W/'versions'
out=V/'repair2_채용과제_분석결과_v001.pptx'
pdf=out.with_suffix('.pdf')
assert not out.exists()
module=ast.parse((W/'reference/build_v004.py').read_text(encoding='utf-8-sig'))
selected=[]
for n in module.body:
 if isinstance(n,ast.FunctionDef) and n.name in {'rgb','rect','txt','line','node','bar','dot','ribbon'}:selected.append(n)
 if isinstance(n,ast.Assign) and any(isinstance(t,ast.Name) and t.id in {'C','F'} for t in n.targets):selected.append(n)
exec(compile(ast.Module(body=selected,type_ignores=[]),'helpers','exec'))
p=Presentation(W/'reference/repair2_수리요청프로세스_분석발표.pptx')
old=list(p.slides)
def page(role,title,intent,source,conclusion,foot=''):
 s=p.slides.add_slide(p.slide_layouts[6])
 s.background.fill.solid();s.background.fill.fore_color.rgb=rgb('bg')
 txt(s,.65,.25,12.03,.27,role,11,color='muted')
 txt(s,.65,.68,12.03,1.01,title,27,font='head')
 txt(s,.65,1.79,12.03,.58,intent,16,color='muted')
 txt(s,.75,6.22,11.8,.51,conclusion,14.5)
 if foot:txt(s,.65,6.8,12.03,.21,foot,9.5,color='muted')
 line(s,.65,7.08,12.68,7.08)
 txt(s,.65,7.16,11.3,.2,'repair2 분석 / '+source,9.5,color='muted')
 txt(s,12.1,7.13,.58,.25,'00',10,color='muted',align=PP_ALIGN.RIGHT)
 return s
def rewrite(shape,text):
 tf=shape.text_frame
 base=tf.paragraphs[0]
 template=base.runs[0] if base.runs else None
 props=(template.font.name,template.font.size,template.font.bold) if template else None
 # Preserve all formatting of first run, including colour and East Asian typeface.
 from copy import deepcopy
 rp=deepcopy(template._r.rPr) if template is not None and template._r.rPr is not None else None
 tf.clear()
 base=tf.paragraphs[0];base.clear()
 for i,t in enumerate(text.split('\n')):
  par=base if i==0 else tf.add_paragraph()
  if i:par._p.insert(0,deepcopy(base._p.pPr)) if base._p.pPr is not None else None
  run=par.add_run();run.text=t
  if rp is not None:run._r.insert(0,deepcopy(rp))
def change(sl,needle,text):
 for sh in sl.shapes:
  if sh.has_text_frame and needle in sh.text:
   rewrite(sh,text);return
 raise ValueError(needle)
# Cover: submission audience and a readable scope.
change(old[0],'노현동 /','노현동 / 데이터 분석 채용 과제 / 2026.10.26')
change(old[0],'수리 요청의 업무 흐름','네덜란드 임대 주택 기관의 수리 요청 데이터\n업무 흐름과 기록상 대기를 분석해 개선 검토 대상을 제안')
# New 02: initial questions, without claiming preregistration.
q=page('분석 목적과 검토 질문','수리 요청의 지연과 반복이\n어디에서 나타나는지 확인하고자 했다',
 '업무 수행 시간과 업무 사이 대기를 구분하고, 업무 구성·수리 유형·접수 방식과의 관계를 살펴봤다',
 '분석 질문 재구성 / 노트북 7–15',
 '아래 질문은 분석 의도를 정리한 것이다. E 유형과 특정 채널의 세부 비교는 탐색 결과에서 이어졌다.',
 '사전 등록된 가설이나 인과 검증으로 제시하지 않음')
rows=[
 ('시간이 길어지는 곳은 어디인가','개별 업무 시간과 업무 사이 대기를 나눠 비교'),
 ('동일한 수리 업무가 반복되는가','업무 종류, 수행 횟수, 시작·완료 순서를 확인'),
 ('수리 유형·업무 구성과 관련되는가','RepairType별 전체 시간과 같은 업무 조건을 비교'),
 ('채널·담당자 기록과 관련되는가','접수 채널과 시작·완료 담당자별 시간 분포를 비교')]
for i,(a,b) in enumerate(rows):
 y=2.65+i*.67
 txt(q,.8,y,5.3,.42,a,18,bold=True);txt(q,6.3,y,6.1,.42,b,15)
 line(q,.8,y+.56,12.5,y+.56)
txt(q,.8,5.55,11.7,.42,'데이터 품질 확인  →  업무 구성·순서 확인  →  시간 비교  →  개선과 검증 지표 제안',16,color='accent',bold=True)
# Replace old data page with sample definitions in the visible page.
data=page('데이터 개요와 전처리','1,000건의 수리 요청과\n13,262개 업무 기록을 분석했다',
 'caseID는 요청 1건, taskID는 수행 업무다. start·complete 기록으로 업무 쌍을 구성하고 Date·Time을 결합했다.',
 '1–5 / 8–9 / 상세 보고서 5–7쪽',
 '1970년은 실제 연도로 확정한 값이 아니다. 다른 연도가 섞이지 않은 기록의 시각 차이만 탐색했다.',
 '일반 경로 시간 표본은 339/927건(36.6%) / 전체 일반 요청을 대표한다고 단정할 수 없음')
txt(data,.8,2.5,11.7,.42,'기본 확인: 13개 업무 · 16명 담당자 / 날짜·시각 결측 13행 / 수리 속성 결측은 예외 경로와 연결',15)
headers=[(.8,2.7,'분석 단위'),(3.7,1.7,'사용 규모'),(5.6,6.8,'선별 기준과 해석')]
rect(data,.75,3.13,11.8,.45,'ink')
for x,w,t in headers:txt(data,x,3.19,w,.28,t,13,color='white',bold=True)
for i,(a,b,c) in enumerate([
 ('업무 구성·경로','1,000건','원본 이벤트 보존. 날짜와 무관한 업무 존재·구성 확인'),
 ('개별 업무 시간','3,151쌍','시작·완료 각 1회, 모두 1970년, 완료 ≥ 시작'),
 ('일반 경로 전체 시간','339건','모든 시각 존재·1970년인 402건 중 예외 63건 제외')]):
 y=3.72+i*.62
 if i==2:rect(data,.75,y-.05,11.8,.58,'pale')
 txt(data,.8,y,2.8,.4,a,15,bold=True);txt(data,3.7,y,1.7,.4,b,18,bold=True)
 txt(data,5.6,y,6.75,.4,c,14)
txt(data,.8,5.65,11.7,.36,'1:1 연결 후보 3,631쌍 → 시간 비교 가능 3,627쌍 → 보수적 조건 충족 3,151쌍',15)
# New 04: verified evidence and decision.
date=page('날짜 검증과 처리 결정','날짜 이상을 확인했지만\n1970년을 2020년으로 바꾸지는 않았다',
 '이상 패턴을 찾는 것과 실제 날짜를 복원하는 것은 다르다. 기준 연도와 변환 규칙을 확인할 근거가 없었다.',
 '9-1–9-5 / 상세 보고서 5–7쪽',
 '연도 치환은 보류하고 원본을 유지했다. 업무 구성은 전체, 시간 계산은 조건을 충족한 표본으로 나눴다.',
 '1970년 기록의 실제 날짜 정확성 미확인 / 1970 12,383행 · 2020–2022 860행 · 1905 6행 · 결측 13행')
txt(date,.8,2.55,6.4,.35,'관찰 → 검증 → 복원 가능성 확인',17,bold=True)
for i,(n,a,b) in enumerate([
 ('860/860행','원본 행 번호와 날짜 증가 규칙이 일치','0~859행: 2020-01-06부터 매일 1일씩 증가'),
 ('226/3,627쌍','완료가 시작보다 빠르게 기록','시간 역전 226쌍 모두 시작·완료 연도 불일치'),
 ('636개 이벤트','같은 케이스·시각의 1970년 기록과 연결','모두 서로 다른 업무. 시각 일치만으로 날짜 확정 불가')]):
 y=3.12+i*.88
 txt(date,.8,y,2.1,.4,n,21,color='accent',bold=True)
 txt(date,3,y,4.2,.36,a,14.5,bold=True)
 txt(date,.8,y+.43,6.6,.32,b,13,color='muted')
line(date,7.68,2.58,7.68,5.95)
txt(date,8.04,2.55,4.3,.35,'처리 대안과 판단',17,bold=True)
for i,(a,b) in enumerate([
 ('연도만 2020년으로 치환','실제 연도·오류 규칙을 검증할 수 없어 보류'),
 ('이상 날짜 행을 모두 삭제','업무 구성·반복 정보까지 잃으므로 보류'),
 ('원본 유지 + 목적별 표본 선별','기록을 보존하면서 시간 해석 범위를 제한')]):
 y=3.12+i*.88
 txt(date,8.04,y,4.4,.36,a,15,bold=True,color='accent' if i==2 else 'ink')
 txt(date,8.04,y+.4,4.4,.49,b,14)
# New 07: different unit and sample explicitly separated.
tasktime=page('개별 업무 시간과 전체 시간 비교','일정 조율 업무의 긴 처리 시간만으로는\n요청 전체의 지연을 설명하기 어려웠다',
 'ArrangeSurvey(현장 조사 일정 조율)의 일부 긴 처리 기록을 확인한 뒤, 케이스 전체 시간과의 관계를 비교했다.',
 '12-1–12-2 / 상세 보고서 9쪽',
 '긴 업무 시간이 관찰됐지만 전체 시간 중앙값 차이는 작았다. 이어서 업무 사이 대기를 비교했다.',
 '서로 다른 표본·단위를 직접 합산하지 않음 / 중앙값 비교만으로 관련성 부재나 인과관계를 확정하지 않음')
txt(tasktime,.8,2.57,5.4,.4,'업무 시작→완료 / 809쌍 / 분',17,bold=True)
for y,a,b in [(3.3,'중앙값','4분'),(4.12,'P90','404.8분')]:
 txt(tasktime,.8,y,2,.4,a,17);txt(tasktime,3.1,y,2.8,.5,b,30,color='accent',bold=True)
txt(tasktime,.8,5.19,5.2,.7,'1시간 초과 194쌍(24.0%)\nP90: 기록의 하위 90%와 상위 10%를 나누는 경계',13)
line(tasktime,6.35,2.6,6.35,5.95)
txt(tasktime,6.75,2.57,5.6,.4,'전체 경과 시간 중앙값 / 339건 / 시간',17,bold=True)
for i,(label,val,n) in enumerate([('조율 1시간 초과',78.2,75),('조율 1시간 이하',77.2,264)]):
 y=3.32+i*1.16
 txt(tasktime,6.75,y,5.5,.35,f'{label} / {n}건',15)
 rect(tasktime,6.75,y+.5,4.2*val/90,.26,'accent' if i==0 else 'bar')
 txt(tasktime,11,y+.37,1.4,.5,f'{val}h',23,bold=True)
# Replacement final page: explicit metrics and limits in body.
last=page('효과 검증과 분석 한계','신뢰할 수 있는 기록으로 기준선을 세우고\n요청 전체의 처리 시간이 줄었는지 검증한다',
 '예정·실제 시각과 지연 사유를 수집한 뒤, 수리 유형·접수 채널별 구성을 함께 확인하며 시범 개선 전후를 비교한다.',
 '9-4–9-5 / 15-5 / 개선 후 측정 설계',
 '개별 구간이 빨라져도 전체 완료 시간이 줄지 않을 수 있다. 전체 시간 중앙값과 P90을 함께 확인한다.',
 '개선안은 미실행 제안이며 절감률·인력 부족·자재 부족을 현재 자료로 단정하지 않음')
txt(last,.8,2.55,6.2,.4,'제안별 측정 지표',18,bold=True)
for i,(a,b) in enumerate([
 ('조사 일정 관리','조사 착수 지연 시간 · 예정 초과 미착수율'),
 ('E 유형 수리 착수 관리','수리 시작 전 대기 · 예정 초과 미착수율'),
 ('Web·Letter 처리 점검','조율 업무 1시간 초과율 · 처리 시간 P90')]):
 y=3.14+i*.73
 txt(last,.8,y,6.25,.3,a,15,bold=True)
 txt(last,.8,y+.32,6.25,.34,b,14.5)
txt(last,.8,5.46,6.25,.58,'미착수율 = 조회 시점에 예정이 지났으나 미착수인 건\n÷ 조회 시점까지 착수 예정인 전체 건',12.5,color='muted')
line(last,7.3,2.6,7.3,6.03)
txt(last,7.7,2.55,4.8,.4,'해석의 한계',18,bold=True)
txt(last,7.7,3.14,4.7,2.75,
 '1970년 날짜의 실제 정확성 미확인\n339건은 일반 경로 927건의 36.6%\n예정 시각이 없어 예약 대기·지연 구분 불가\nExternRepair(외부 수리) 완료 시각 없음\n1시간·24시간 기준은 탐색용이며 SLA 아님\n담당자 차이·처리 건수만으로 원인 확정 불가',14)
# Refresh interpretation in retained pages and visible task translations.
change(old[3],'이어서 전체 처리 시간 중','반복 기록만으로 재작업이나 불필요한 반복을 확정할 수는 없다\n먼저 개별 업무 시간과 요청 전체 시간의 관계를 살펴봤다')
change(old[6],'실제 수리 업무','수리 시작→완료')
# Assemble 13-page reading order; drop superseded pages and their relationships.
order=[old[0],q,data,date,old[2],old[3],tasktime,old[4],old[5],old[6],old[7],old[8],last]
ids={p.part.related_slide(el.rId).part:el for el in list(p.slides._sldIdLst)}
for el in list(p.slides._sldIdLst):p.slides._sldIdLst.remove(el)
keep={s.part for s in order}
for key,el in ids.items():
 if key not in keep:p.part.drop_rel(el.rId)
for s in order:p.slides._sldIdLst.append(ids[s.part])
for i,s in enumerate(p.slides):
 for sh in s.shapes:
  if sh.has_text_frame and sh.left>Inches(12) and sh.top>Inches(7):
   rewrite(sh,f'{i+1:02}')
  for ef in sh._element.xpath('.//a:effectRef'):ef.set('idx','0')
  for sp in sh._element.xpath('./p:spPr'):
   for a in list(sp):
    if a.tag.endswith('}effectLst') or a.tag.endswith('}effectDag'):sp.remove(a)
   sp.append(OxmlElement('a:effectLst'))
 visible='\n'.join(sh.text for sh in s.shapes if sh.has_text_frame)
 s.notes_slide.notes_text_frame.text='読者向け補足'.replace('読者向け補足','검토용 근거와 해석')+'\n'+visible+'\n\n시간은 기록상 탐색 값이며 실제 날짜·인과관계·절감 효과가 검증된 것은 아니다. 원본 분석 노트북의 해당 절과 상세 보고서를 참고한다.'
p.core_properties.title='수리 요청 프로세스 분석 및 개선 제안 — 채용 과제'
p.core_properties.subject='업무 흐름, 날짜 품질 검증, 기록상 대기 분석과 개선 제안'
p.save(out)
app=win32com.client.Dispatch('PowerPoint.Application')
deck=app.Presentations.Open(str(out),False,False,False)
overflow=[]
try:
 for sl in deck.Slides:
  for sh in sl.Shapes:
   if sh.HasTextFrame and sh.TextFrame.HasText:
    tr=sh.TextFrame.TextRange
    if tr.BoundHeight>sh.Height+2 or tr.BoundWidth>sh.Width+2:overflow.append({'page':sl.SlideIndex,'text':tr.Text,'h':sh.Height,'bh':tr.BoundHeight,'w':sh.Width,'bw':tr.BoundWidth})
 deck.Save();deck.SaveAs(str(pdf),32)
finally:deck.Close()
pp=Presentation(out);doc=fitz.open(pdf);norm=lambda t:''.join(t.split());mismatch=[];bounds=[]
for i,s in enumerate(pp.slides):
 for sh in s.shapes:
  if sh.has_text_frame and norm(sh.text) not in norm(doc[i].get_text()):mismatch.append([i+1,sh.text])
  if sh.left<0 or sh.top<0 or sh.left+sh.width>pp.slide_width+1000 or sh.top+sh.height>pp.slide_height+1000:bounds.append([i+1,sh.name])
 doc[i].get_pixmap(matrix=fitz.Matrix(1.4,1.4)).save(Q/f'page-{i+1:02}.png')
for start in range(0,13,4):
 sheet=fitz.open();pg=sheet.new_page(width=1600,height=900)
 for j,i in enumerate(range(start,min(start+4,13))):pg.show_pdf_page(fitz.Rect(j%2*800,j//2*450,(j%2+1)*800,(j//2+1)*450),doc,i)
 pg.get_pixmap().save(Q/f'sheet-{start+1:02}.png')
report={'version':1,'pages':len(doc),'overflow':overflow,'mismatch':mismatch,'bounds':bounds,'raster_images':sum(len(pg.get_images()) for pg in doc),'native_charts':sum(sh.has_chart for sl in pp.slides for sh in sl.shapes),'fonts':sorted({x[3] for pg in doc for x in pg.get_fonts()}),'visual_review':'pending','latest_second_check':'pending'}
(Q/'validation.json').write_text(json.dumps(report,ensure_ascii=False,indent=2),encoding='utf-8')
print(json.dumps(report,ensure_ascii=False,indent=2))


