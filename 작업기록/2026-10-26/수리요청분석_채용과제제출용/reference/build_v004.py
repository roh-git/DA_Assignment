from pathlib import Path
import json, re, shutil, hashlib, math
from datetime import datetime
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE, MSO_CONNECTOR
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.chart import XL_CHART_TYPE, XL_LABEL_POSITION, XL_LEGEND_POSITION, XL_TICK_MARK
from pptx.chart.data import CategoryChartData
from pptx.oxml.xmlchemy import OxmlElement
import pymupdf as fitz
import win32com.client

R=Path.cwd(); D=R/'작업기록/2026-10-26'; W=D/'수리요청분석'; V=D/'versions/수리요청분석'
for p in [W,V,W/'sources',W/'previews']: p.mkdir(parents=True,exist_ok=True)
v=max([int(x.stem.rsplit('_v',1)[1]) for x in V.glob('*_v*.pptx')]+[0])+1
name=f'repair2_수리요청프로세스_분석발표_v{v:03}'
OUT=V/(name+'.pptx'); PDF=OUT.with_suffix('.pdf'); Q=W/f'검수_v{v:03}'; Q.mkdir(exist_ok=True)
NB=Path('C:/Users/rhd67/orca/workspaces/DA-repair/cormorant/repair2_분석.ipynb')
GUIDE=Path('C:/Users/rhd67/Downloads/repair2_PPT_제작가이드_최종.md')
sha=lambda p:hashlib.sha256(p.read_bytes()).hexdigest()
master_paths=[R/'노현동_포트폴리오.pdf', R/'최종본/노현동_포트폴리오.pptx']
baseline={str(p.relative_to(R)):sha(p) for p in master_paths}
nb=json.loads(NB.read_text(encoding='utf-8')); guide=GUIDE.read_text(encoding='utf-8')
for src in [NB,GUIDE]:
 dst=W/'sources'/src.name
 if not dst.exists():shutil.copy2(src,dst)
 else:assert sha(src)==sha(dst),'Source changed: preserve separately'
ids=[12,15,23,58,59,86,87,88,105,108,109,110,138,158,168,174,183,187,189,209,243,250,268,286,294,301,306]
evidence={str(i):{'type':nb['cells'][i]['cell_type'],'source':''.join(nb['cells'][i]['source']),'outputs':'\n'.join(''.join(o.get('text',o.get('data',{}).get('text/plain',[]))) for o in nb['cells'][i].get('outputs',[]))} for i in ids}
(W/'sources/근거발췌.json').write_text(json.dumps(evidence,ensure_ascii=False,indent=2),encoding='utf-8')
checks={23:['13262','1000'],87:['339','63'],138:['836','91','38'],187:['39','89','63','90','58'],209:['70.8','22.5','4.9'],286:['25.2','143'],294:['25.9','127'],243:['332.0','704','105'],168:['109','85','224','203']}
for i,terms in checks.items():
 text=''.join(str(x) for x in evidence[str(i)].values())
 assert all(t in text for t in terms),(i,terms)
C={'bg':'FFFFFF','ink':'202C32','muted':'52616A','accent':'087F75','pale':'E5F2EF','surface':'F1F4F5','rule':'D5DEE1','bar':'9AA8AE','light':'B8DAD3','white':'FFFFFF'}
F={'head':'Paperlogy 6 SemiBold','body':'Pretendard'}
p=Presentation();p.slide_width=Inches(13.333333);p.slide_height=Inches(7.5)
p.core_properties.title='수리 요청 프로세스 분석 및 개선 제안';p.core_properties.author='Roh Hyundong';p.core_properties.subject='repair2 분석 과제 발표자료 / 포트폴리오와 별도 관리'
SL=[]
def rgb(k):return RGBColor.from_string(C.get(k,k))
def rect(s,x,y,w,h,c='surface',line=None,rad=False):
 sh=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE if rad else MSO_SHAPE.RECTANGLE, Inches(x),Inches(y),Inches(w),Inches(h));sh.fill.solid();sh.fill.fore_color.rgb=rgb(c);sh.line.fill.background() if not line else None
 if line:sh.line.color.rgb=rgb(line);sh.line.width=Pt(.7)
 if rad:sh.adjustments[0]=.08
 # Remove inherited shadows
 pr=sh._element.spPr; ef=OxmlElement('a:effectLst');pr.append(ef)
 return sh

def txt(s,x,y,w,h,t,size=16,color='ink',bold=False,font='body',align=None):
 sh=s.shapes.add_textbox(Inches(x),Inches(y),Inches(w),Inches(h));tf=sh.text_frame
 tf.clear();tf.word_wrap=True;tf.margin_left=tf.margin_right=tf.margin_top=tf.margin_bottom=0
 tf.vertical_anchor=MSO_ANCHOR.MIDDLE
 for i,line in enumerate(t.split('\n')):
  q=tf.paragraphs[0] if i==0 else tf.add_paragraph();q.space_before=Pt(0);q.space_after=Pt(0);q.line_spacing=1.12
  if align is not None:q.alignment=align
  r=q.add_run();r.text=line;r.font.name=F[font];r.font.size=Pt(size);r.font.bold=bold;r.font.color.rgb=rgb(color)
  rp=r._r.get_or_add_rPr();ea=OxmlElement('a:ea');ea.set('typeface',F[font]);rp.append(ea)
 return sh

def line(s,x1,y1,x2,y2,color='rule',width=1.1,arrow=False):
 sh=s.shapes.add_connector(MSO_CONNECTOR.STRAIGHT,Inches(x1),Inches(y1),Inches(x2),Inches(y2));sh.line.color.rgb=rgb(color);sh.line.width=Pt(width)
 if arrow:
  end=OxmlElement('a:tailEnd');end.set('type','triangle');sh.line._get_or_add_ln().append(end)
 return sh

def dot(s,x,y,r=.055,color='accent'):
 sh=s.shapes.add_shape(MSO_SHAPE.OVAL,Inches(x-r),Inches(y-r),Inches(2*r),Inches(2*r));sh.fill.solid();sh.fill.fore_color.rgb=rgb(color);sh.line.fill.background();return sh

def node(s,x,y,w,h,t,accent=False,size=14):
 rect(s,x,y,w,h,'pale' if accent else 'surface');return txt(s,x+.06,y,w-.12,h,t,size,align=PP_ALIGN.CENTER)

def ribbon(s,x1,x2,y1,y2,h,color):
 # Editable freeform Sankey band, identical thickness at both ends
 pts=[]
 for j in range(31):
  u=j/30;f=3*u*u-2*u*u*u;pts.append((Inches(x1+(x2-x1)*u),Inches(y1+(y2-y1)*f)))
 for j in range(30,-1,-1):
  u=j/30;f=3*u*u-2*u*u*u;pts.append((Inches(x1+(x2-x1)*u),Inches(y1+(y2-y1)*f+h)))
 fb=s.shapes.build_freeform(int(pts[0][0]),int(pts[0][1]));fb.add_line_segments([(int(x),int(y)) for x,y in pts[1:]],close=True);sh=fb.convert_to_shape();sh.fill.solid();sh.fill.fore_color.rgb=rgb(color);sh.line.fill.background();return sh

def slide(title,sub,src,foot='',dark=False):
 s=p.slides.add_slide(p.slide_layouts[6]);s.background.fill.solid();s.background.fill.fore_color.rgb=rgb('ink' if dark else 'bg');SL.append(s)
 if not dark:
  txt(s,.65,.43,12.03,1.04,title,27,font='head');txt(s,.65,1.56,12.03,.62,sub,16,color='muted')
  if foot:txt(s,.65,6.58,12.03,.39,foot,10.5,color='muted')
  line(s,.65,7.08,12.68,7.08);txt(s,.65,7.16,11.3,.2,'repair2 분석  /  '+src,9.5,color='muted');txt(s,12.13,7.13,.55,.25,f'{len(SL):02}',10,color='muted',align=PP_ALIGN.RIGHT)
 return s

def note(s,page,extra=''):
 m=re.search(r'### '+f'{page:02}'+r'\s.*?(?=\n### |\n## |\Z)',guide,re.S)
 s.notes_slide.notes_text_frame.text=(m.group(0) if m else '')+'\n\n제작·근거 메모\n'+extra+'\n시간값은 실제 날짜 신뢰성을 검증하지 못한 탐색적인 기록상 경과 시간이며 확정된 절감 가능 시간이나 인과효과가 아님'

def bar(s,x,y,w,h,cats,series,maxval,fmt='0',legend=False,accents=None):
 data=CategoryChartData();data.categories=cats
 for label,values,color in series:data.add_series(label,values)
 ch=s.shapes.add_chart(XL_CHART_TYPE.BAR_CLUSTERED,Inches(x),Inches(y),Inches(w),Inches(h),data).chart
 ch.has_title=False;ch.has_legend=legend
 if legend:
  ch.legend.position=XL_LEGEND_POSITION.BOTTOM;ch.legend.include_in_layout=False;ch.legend.font.name=F['body'];ch.legend.font.size=Pt(13)
 plot=ch.plots[0];plot.gap_width=70;plot.has_data_labels=True
 dl=plot.data_labels;dl.position=XL_LABEL_POSITION.OUTSIDE_END;dl.font.name=F['body'];dl.font.size=Pt(13);dl.font.color.rgb=rgb('ink');dl.number_format=fmt
 ca=ch.category_axis;ca.reverse_order=True;ca.tick_labels.font.name=F['body'];ca.tick_labels.font.size=Pt(14);ca.tick_labels.font.color.rgb=rgb('ink');ca.major_tick_mark=XL_TICK_MARK.NONE;ca.minor_tick_mark=XL_TICK_MARK.NONE;ca.format.line.fill.background()
 va=ch.value_axis;va.minimum_scale=0;va.maximum_scale=maxval;va.major_unit=maxval/4;va.tick_labels.font.name=F['body'];va.tick_labels.font.size=Pt(11);va.tick_labels.font.color.rgb=rgb('muted');va.major_tick_mark=XL_TICK_MARK.NONE;va.minor_tick_mark=XL_TICK_MARK.NONE;va.has_major_gridlines=True;va.major_gridlines.format.line.color.rgb=rgb('rule');va.major_gridlines.format.line.width=Pt(.5);va.format.line.fill.background();va.tick_labels.number_format='0'
 for k,(lab,vals,col) in enumerate(series):
  se=ch.series[k];se.format.fill.solid();se.format.fill.fore_color.rgb=rgb(col);se.format.line.fill.background()
  if accents and k in accents:
   for i,c in enumerate(accents[k]):se.points[i].format.fill.solid();se.points[i].format.fill.fore_color.rgb=rgb(c);se.points[i].format.line.fill.background()
 # Set East Asian fonts on chart text runs/defaults
 for e in ch._chartSpace.iter():
  if e.tag.endswith('}defRPr') or e.tag.endswith('}rPr'):
   ea=OxmlElement('a:ea');ea.set('typeface',F['body']);e.append(ea)
 return ch

# 01 — typographic cover with evidence inventory rather than promotional KPI cards
s=slide('','','',dark=True)
line(s,.7,.8,12.6,.8,'muted',.8)
txt(s,.75,1.35,10.7,1.8,'수리 요청 프로세스\n분석 및 개선 제안',39,color='white',font='head')
txt(s,.78,3.5,10.9,.7,'처리 구조를 재구성하고 주요 대기 구간을 확인해\n추가 기록과 운영 개선의 우선순위를 제안한다',20,color='white')
for x,num,label in [(.8,'1,000','전체 수리 요청 / 건'),(4.8,'927','일반 경로 / 건'),(8.8,'339','시간 분석 표본 / 건')]:
 txt(s,x,5.0,3.4,.75,num,38,color='white',font='head');txt(s,x,5.86,3.4,.4,label,16,color='white')
txt(s,.8,6.92,11.7,.25,'노현동  /  분석 과제 발표  /  2026.10.26',11,color='white')
note(s,1,'표지 70.8% 미사용 / 1,000·927·339건 분석 단위 구분')

# 02 — branching sample selection
s=slide('데이터 품질을 검토하고\n분석 대상을 구분하다','업무 구조는 전체 요청으로 확인하고 시간 계산은 각 분석의 선별 조건에 맞춰 수행한다','1–5 / 9-5','날짜·시간 결측 13행 / 연도 1905·1970·2020–2022 혼재 / 임의 날짜 보정 보류')
node(s,.7,2.55,2.5,1.15,'전체 1,000건\n13,262개 이벤트',size=18)
line(s,3.2,3.12,3.65,3.12);line(s,3.65,2.88,3.65,4.69);line(s,3.65,2.88,4.1,2.88,arrow=True);line(s,3.65,4.69,4.1,4.69,arrow=True)
node(s,4.1,2.45,3.3,.88,'업무·경로 분석 1,000건',accent=True,size=18);txt(s,7.8,2.49,4.7,.7,'시각 결측이 있어도\n업무 기록은 보존',16)
node(s,4.1,4.1,3.3,1.17,'시각 완전 케이스 402건\n모든 이벤트가 1970년 기록',size=16);line(s,7.4,4.69,8.15,4.69,arrow=True)
node(s,8.15,4.1,4.43,1.17,'일반 경로 시간 분석 339건\n시간 신뢰성은 별도 검증 필요',accent=True,size=17)
line(s,.75,5.67,12.58,5.67);txt(s,.75,5.86,11.8,.45,'업무별 start–complete 분석: 1:1 연결 3,631쌍 중 유효 3,151쌍',16)
note(s,2,'3,151쌍은 업무 이벤트 연결 단위 / 339건은 일반 경로 케이스 단위 / 일반 경로 전체의 36.6%가 시간 분석 조건 충족')

# 03 — Sankey split (width encodes counts) + native workflow (not count-scaled)
s=slide('수리 요청은 일반·예외 경로와\n수리 업무로 분기된다','경로 구성은 생키의 폭으로 표현하고 일반 경로의 업무 순서는 별도 흐름으로 요약한다','8-2 / 10-3–10-5','흐름도는 구조 요약 / FirstContact 1건 미기록 / ExternRepair 완료 시각 없음 / 종결 관련 두 업무의 순서는 고정하지 않음')
txt(s,.7,2.35,3.0,.36,'전체 요청 1,000건',17,bold=True)
H=1.3; hg=H*.927;he=H*.073
ribbon(s,1.05,2.55,2.95,2.85,hg,'light');ribbon(s,1.05,2.55,2.95+hg,5.32,he,'bar')
rect(s,.92,2.95,.12,H,'ink');rect(s,2.55,2.85,.12,hg,'accent');rect(s,2.55,5.32,.12,he,'bar')
txt(s,.7,4.44,1.7,.55,'FirstContact',13)
txt(s,2.8,2.86,1.1,.7,'일반\n927건',16,bold=True);txt(s,2.8,5.07,1.15,.55,'예외 73건',14,bold=True)
txt(s,.75,5.7,3.05,.5,'InformClientWrongPlace',12.5)
line(s,3.75,3.2,4.1,3.2,arrow=True)
for x,w,t in [(4.1,2.0,'MakeTicket'),(6.55,2.1,'ArrangeSurvey'),(9.1,2.0,'Survey')]:node(s,x,2.87,w,.65,t)
line(s,6.1,3.2,6.55,3.2,arrow=True);line(s,8.65,3.2,9.1,3.2,arrow=True)
txt(s,6.5,3.58,3.6,.3,'보조 안내: InformClientSurvey',11.5,color='muted')
line(s,10.1,3.52,10.1,3.99);line(s,5.1,3.99,10.85,3.99)
for x,w,t in [(4.05,2.15,'InternRepair'),(6.6,2.4,'ImmediateRepair'),(9.4,2.75,'ExternRepair')]:
 line(s,x+w/2,3.99,x+w/2,4.2,arrow=True);node(s,x,4.2,w,.62,t,accent=True,size=13.5);line(s,x+w/2,4.82,x+w/2,5.1)
line(s,5.125,5.1,10.775,5.1);txt(s,6.0,5.12,6.5,.29,'단일·복합·반복 수행 가능',12,color='muted')
line(s,5.125,5.1,5.125,5.7,arrow=True)
node(s,4.05,5.7,2.2,.62,'RepairReady',size=13.5);line(s,6.25,6.01,6.6,6.01,arrow=True)
node(s,6.6,5.7,3.2,.62,'ReadyInformClient\nSendTicketToFinAdmin',size=12.0);line(s,9.8,6.01,10.2,6.01,arrow=True);node(s,10.2,5.7,2.05,.62,'TicketReady',size=13.5)
note(s,3,'생키 정량 폭은 일반 927/1,000와 예외 73/1,000에만 적용 / 이후 선은 건수 비례가 아닌 논리적 요약 / 분기 후 수리 업무별 건수는 중복 가능하므로 합산하지 않음')

# 04 — part-to-whole + independent repeated work observation
s=slide('일부 케이스에서 복합 수리와\n반복 수행이 관찰된다','일반 경로 927건의 업무 구성과 InternRepair 반복 기록을 서로 다른 기준으로 살펴본다','10-3 / 10-5 / 12-5','복합 수리와 반복 기록은 겹칠 수 있음 / 반복 수행이 재작업인지는 현재 데이터로 확인할 수 없음')
txt(s,.7,2.55,6.5,.4,'일반 경로 927건의 수리 업무 구성',18,bold=True)
rect(s,.75,3.3,6.35*836/927,.66,'bar');rect(s,.75+6.35*836/927,3.3,6.35*91/927,.66,'accent')
txt(s,.75,4.1,3.7,.8,'단일 수리 836건\n90.2%',19);txt(s,4.6,4.1,2.9,.8,'복합 수리 91건\n9.8%',19)
line(s,7.65,2.55,7.65,4.95);txt(s,8.1,2.57,4.4,.4,'InternRepair 반복 기록',18,bold=True);txt(s,8.1,3.15,4.4,.82,'38건',36,font='head');txt(s,8.1,4.13,4.4,.7,'시간순 검토 29건 중 28건은\n완료 후 다시 시작하는 형태',16)
line(s,.75,5.25,12.58,5.25)
txt(s,.75,5.46,2.55,.64,'실제 기록 순서 예시\ncaseID 124',13,color='muted')
for x,time,event in [(3.55,'11:41','start'),(5.78,'15:21','complete'),(8.01,'19:45','start'),(10.24,'23:03','complete')]:
 dot(s,x,5.85);txt(s,x-.6,5.39,1.5,.3,time,14,align=PP_ALIGN.CENTER);txt(s,x-.6,6.0,1.5,.3,event,13,align=PP_ALIGN.CENTER)
line(s,3.55,5.85,10.24,5.85,'bar');txt(s,11.1,5.53,1.45,.67,'同日記録\n간격 비례 아님'.replace('同日記録','동일 일자'),11,color='muted')
note(s,4,'caseID124: 1970-01-18 11:41 start → 15:21 complete → 19:45 start → 23:03 complete / 타임라인은 순서만 표현하며 시간 간격 비례 아님')

# 05 — explicit time-share strip and process milestones
s=slide('조사 시작 전 대기가\n전체 처리 시간 합계의 70.8%를 차지한다','일반 경로 시간 표본 339건의 기록상 경과 시간을 네 구간으로 나누어 비교한다','12-2-1 / 12-4-1','비중 = 해당 구간 시간 합계 / 전체 시간 합계 / 반올림 합계 100.1% / 절감 가능 시간으로 해석하지 않음')
txt(s,.75,2.5,7.6,.47,'ArrangeSurvey complete → Survey start',21,bold=True);txt(s,9.4,2.42,3.05,.8,'70.8%',37,font='head',align=PP_ALIGN.RIGHT)
vals=[1.9,70.8,22.5,4.9];colors=['bar','accent','bar','light']
# Reference-inspired whole/part visual, native editable shapes, no raster asset
rect(s,.65,3.2,12.03,1.37,'ink',rad=True)
rect(s,.88,3.43,11.57,.39,'bar',rad=True)
rect(s,.88,3.43,11.57*.708,.39,'accent',rad=True)
rect(s,.88+11.57*.708-.16,3.43,.16,.39,'accent')
rect(s,.88+11.57*.708-.027,3.43,.054,.39,'ink')
txt(s,.88,4.0,7.8,.33,'조사 시작 전 대기 70.8%',14,color='white',bold=True)
txt(s,9.25,4.0,3.2,.33,'나머지 세 구간',14,color='white',align=PP_ALIGN.RIGHT)
labels=[('접수→조사 준비 완료','1.9%'),('조사 준비 완료→조사 시작','70.8%'),('조사 시작→수리 완료','22.5%'),('수리 완료→종결','4.9%')]
for j,(lab,num) in enumerate(labels):
 x=.75+j*3.02;rect(s,x,4.84,.13,.25,colors[j]);txt(s,x+.23,4.75,2.65,.38,num,19,bold=True);txt(s,x,5.27,2.8,.55,lab,13.5)
txt(s,.75,6.01,11.8,.38,'전체 처리 시간 중앙값 77.3시간  /  22.5% 구간에는 Survey와 수리 관련 경과 시간이 함께 포함',14,color='muted')
note(s,5,'구간 경계: FirstContact complete → ArrangeSurvey complete → Survey start → RepairReady complete → TicketReady complete / 합계 비중과 케이스별 비중 중앙값 66.3%를 혼용하지 않음')

# 06 — editable native PowerPoint chart with embedded data
s=slide('ArrangeSurvey → Survey 대기는\n다수의 케이스에서 길게 나타난다','339건 중 211건에서 24시간을 초과하는 기록상 대기가 관찰된다','12-3-1–12-3-3','P90은 약 90%가 그 값 이하인 경계 / 24시간은 분석 경계이며 SLA 기준이 아님 / 정상 예약 대기와 지연은 구분 불가')
txt(s,.75,2.4,7.25,.35,'조사 시작 전 대기 분포 / 339건 / 단위: 건',15,bold=True)
bar(s,.7,2.94,7.2,3.39,['0시간','0시간 초과–24시간 이하','24시간 초과–72시간 이하','72시간 초과–168시간 이하','168시간 초과'],[('케이스 수',[39,89,63,90,58],'bar')],120,accents={0:['bar','bar','accent','accent','accent']})
line(s,8.36,2.54,8.36,6.25)
for y,val,lab in [(2.53,'54.5시간','대기 시간 중앙값'),(3.78,'195.7시간','대기 시간 P90'),(5.03,'62.2%','24시간 초과 211 / 339건')]:
 txt(s,8.8,y,3.75,.62,val,29,font='head');txt(s,8.8,y+.65,3.75,.4,lab,15)
note(s,6,'분포 39+89+63+90+58=339 / 24시간 초과 63+90+58=211 / native chart에는 원값을 포함한 편집용 통합문서 내장')

# 07 — grouped comparison and restricted-condition replication
s=slide('E 유형은 수리 시작 전 대기가\n상대적으로 길다','동일한 InternRepair 단독 업무로 범위를 좁혀도 E 유형의 긴 대기가 유지된다','14-1 / 14-4-1–14-4-2','모든 시간은 기록상 경과 시간 / 구간 중앙값의 합은 전체 중앙값이 아님 / E 유형 전체 처리 시간 중앙값 160.4시간')
txt(s,.75,2.4,7.2,.42,'수리 유형별 구간 중앙값 / 338건 / 단위: 시간(h)',15,bold=True)
bar(s,.7,2.94,7.25,3.52,['B (46건)','E (143건)','P (149건)'],[('수리 시작 전 대기',[0,25.2,0],'accent'),('수리 시작→RepairReady',[3.2,4.0,2.2],'bar')],32,fmt='0.0',legend=True,accents={0:['light','accent','light']})
rect(s,8.45,2.5,4.13,3.77,'surface');txt(s,8.75,2.72,3.5,.6,'동일 업무 조건으로 재확인',17,bold=True);txt(s,8.75,3.42,3.5,.7,'E 유형 127건\nInternRepair 단독 1:1 연결',15)
for yy,label,val in [(4.42,'수리 시작 전 대기','25.9h'),(5.2,'실제 InternRepair','4.0h')]:
 txt(s,8.75,yy,2.5,.36,label,14);txt(s,11.15,yy-.08,1.15,.52,val,21,font='head',align=PP_ALIGN.RIGHT)
txt(s,8.75,5.85,3.5,.25,'단독 비교 전체 160건: B 25 / E 127 / P 8',10.5,color='muted')
note(s,7,'좌측 338건의 수리 시작 이후는 RepairReady까지의 경과 시간 / 우측 160건의 실제 업무는 InternRepair start→complete / 우측 B 대기0.0h·실제3.2h / E25.9h·4.0h / P0.9h·4.0h')

# 08 — channel rate chart + start/complete recorder comparison
s=slide('Web·Letter 접수에서 ArrangeSurvey\n장시간 처리 사례가 집중된다','1시간 초과 194건은 모두 두 채널에서 발생하며 시작·완료 담당자 차이와도 관련성이 보인다','12-1-2–12-1-4 / 13-3–13-4','ArrangeSurvey 내부 start→complete 기록상 경과 시간 / 담당자 차이가 실제 인수인계나 지연 원인이라는 뜻은 아님')
txt(s,.75,2.4,7.2,.42,'ArrangeSurvey 1시간 초과율 / 809건 / 단위: %',15,bold=True)
bar(s,.7,2.94,7.0,3.35,['Web (109/207건)','Letter (85/175건)','Phone (0/224건)','Personal (0/203건)'],[('1시간 초과율',[52.7,48.6,0,0],'bar')],80,fmt='0.0',accents={0:['accent','accent','bar','bar']})
line(s,8.12,2.53,8.12,6.25);txt(s,8.55,2.5,4,.64,'시작·완료 담당자별 중앙값\nArrangeSurvey / 809건',16,bold=True)
for y,label,n,val in [(3.45,'동일 담당자','704건','4분'),(4.49,'서로 다른 담당자','105건','332분')]:
 txt(s,8.55,y,2.5,.33,label,15);txt(s,8.55,y+.39,2.0,.28,n,13,color='muted');txt(s,11.05,y,1.5,.67,val,25,font='head',align=PP_ALIGN.RIGHT)
txt(s,8.55,5.55,3.9,.65,'동일 담당자의 Web·Letter에서도\n1시간 초과 89건 관찰',15)
note(s,8,'오른쪽 값은 업무 처리 시간 중앙값 / Web 중앙값119분·Letter5분 / 같은 채널에서 담당자 상이 집단의 중앙값 Letter317.5분·Web374분 / 인수인계 실제 발생 여부 미확인')

# 09 — editorial comparison rows, rather than three generic cards
s=slide('확인된 문제를 기준으로\n세 가지 개선 과제를 제안한다','우선순위는 검토 순서이며 절감효과의 크기나 확인된 원인을 뜻하지 않는다','15-2–15-5','원인을 추가 기록한 뒤 확인된 조건에만 운영 조치를 시범 적용 / 자재 부족이나 인력 과부하를 원인으로 단정하지 않음')
cols=[.75,1.55,5.3,9.0];heads=['순위','관찰한 문제','추가로 기록할 정보','검토할 운영 조치']
rect(s,.65,2.42,12.03,.56,'ink')
for x,w,t in zip(cols,[.7,3.5,3.45,3.4],heads):txt(s,x,2.45,w,.48,t,15,color='white',bold=True)
rows=[('1','조사 착수 전 대기\n전체 시간 합계의 70.8%','조사 예정 시각\n배정 담당자와 착수 상태','예정 시각 초과 미착수 확인\n사유에 따른 재배정·일정 조정'),('2','E 유형 수리 전 대기\n중앙값 25.2시간','수리 예정 시각\n착수 전 상태와 지연 사유','확인된 지연 요인에 맞춰\n착수 일정과 진행 상태 관리'),('3','ArrangeSurvey 내부 처리\nWeb·Letter 장시간 사례','세부 업무 상태\n담당자 변경 사유','채널별 처리 절차 점검\n변경 전후 기록 방식 표준화')]
for i,row in enumerate(rows):
 y=3.07+i*.99
 if i==0:rect(s,.65,y-.02,12.03,.94,'pale')
 for x,w,t in zip(cols,[.65,3.5,3.45,3.4],row):txt(s,x,y+.07,w,.74,t,16 if x>1 else 23,bold=x<1,font='head' if x<1 else 'body')
 line(s,.75,y+.95,12.58,y+.95)
note(s,9,'현재 CSV에서 예정 시각과 지연 사유를 계산할 수 없음 / 모두 새로 수집할 항목 / 반복 수리는 정상 복수 수행과 재작업을 구분하도록 작업 사유 기록을 보완')

# 10 — measurable next steps, definitions and limitations
s=slide('추가 기록과 효과 검증으로\n개선 효과를 확인한다','신뢰할 수 있는 시각으로 기준선을 다시 세우고 구간 개선이 전체 기간 단축으로 이어지는지 확인한다','9-4–9-5 / 15-5','수리 유형·접수 채널을 구분해 전후 비교 / 구간별 대기가 다른 구간으로 이동하는지도 함께 확인')
txt(s,.75,2.39,7.5,.4,'개선 전후 함께 볼 지표',18,bold=True)
items=[('조사 착수 지연 시간','max(0, Survey 실제 시작 − 조사 예정 시각)'),('예정 시각 초과 미착수율','예정 시각 초과 미착수 / 예정 시각이 지난 전체'),('수리 시작 전 대기','Survey complete → 첫 수리 start 중앙값·P90'),('ArrangeSurvey 장시간 처리','채널별 1시간 초과율과 처리 시간 P90'),('전체 처리 시간','케이스별 전체 경과 시간 중앙값·P90')]
for i,(a,b) in enumerate(items):
 y=3.01+i*.56;line(s,.75,y-.07,8.16,y-.07);txt(s,.75,y,2.7,.4,a,14,bold=True);txt(s,3.5,y,4.65,.4,b,12.5)
rect(s,8.62,2.42,3.95,3.67,'surface');txt(s,8.9,2.67,3.38,.4,'해석과 검증의 경계',18,bold=True)
for y,t in [(3.37,'1970년 날짜의 정확성 미검증'),(3.98,'예정 시각과 지연 사유 미기록'),(4.59,'ExternRepair 완료 시각 부재'),(5.2,'예상 절감효과 수치 제시 보류')]:txt(s,8.9,y,3.4,.52,t,15)
txt(s,.75,6.07,11.8,.4,'기록 확보 → 원인 확인 → 시범 적용 → 구간·전체 재측정',18,bold=True)
note(s,10,'예정 시각 초과 미착수율은 동일 관측 시점 기준 / 추가 신뢰 시각 수집 후 새 기준선 설정 / 개선 성과 미실행이며 절감률 임의 제시 없음')

assert len(p.slides)==10
# Explicitly disable inherited theme effects on every native object
for sl in p.slides:
 for sh in sl.shapes:
  for ef in sh._element.xpath('.//a:effectRef'): ef.set('idx','0')
  for sp in sh._element.xpath('./p:spPr'):
   for old in list(sp):
    if old.tag.endswith('}effectLst') or old.tag.endswith('}effectDag'):sp.remove(old)
   sp.append(OxmlElement('a:effectLst'))
p.save(OUT)
# Notes and source mapping remain native, no screenshot/raster illustrations
bounds=[]
for i,s in enumerate(p.slides,1):
 for sh in s.shapes:
  if sh.left<0 or sh.top<0 or sh.left+sh.width>p.slide_width+1000 or sh.top+sh.height>p.slide_height+1000:bounds.append([i,sh.name])
assert not bounds,bounds
app=win32com.client.Dispatch('PowerPoint.Application')
deck=app.Presentations.Open(str(OUT),False,False,False)
over=[]
try:
 for sl in deck.Slides:
  for sh in sl.Shapes:
   if sh.HasTextFrame and sh.TextFrame.HasText:
    tr=sh.TextFrame.TextRange
    if tr.BoundHeight>sh.Height+2 or tr.BoundWidth>sh.Width+2:
     over.append({'page':sl.SlideIndex,'shape':sh.Name,'text':tr.Text,'height':sh.Height,'boundHeight':tr.BoundHeight,'width':sh.Width,'boundWidth':tr.BoundWidth})
 deck.Save();deck.SaveAs(str(PDF),32)
finally:deck.Close()
# Verify after Office save and render all actual PDF pages
pp=Presentation(OUT);doc=fitz.open(PDF);norm=lambda t:''.join(t.split())
mismatch=[]
for i,s in enumerate(pp.slides):
 for sh in s.shapes:
  if sh.has_text_frame and norm(sh.text) not in norm(doc[i].get_text()):mismatch.append([i+1,sh.text])
 for chs in s.shapes:
  if chs.has_chart:
   for se in chs.chart.series:
    assert len(se.values)==len(chs.chart.plots[0].categories)
 doc[i].get_pixmap(matrix=fitz.Matrix(1.5,1.5)).save(Q/f'page-{i+1:02}.png')
for start in [0,5]:
 sheet=fitz.open();pg=sheet.new_page(width=1600,height=1350)
 for j,i in enumerate(range(start,start+5)):
  xx=j%2*800;yy=j//2*450;pg.show_pdf_page(fitz.Rect(xx,yy,xx+800,yy+450),doc,i)
 pg.get_pixmap().save(Q/f'전체-{start+1:02}.png')
fonts=sorted({x[3] for pg in doc for x in pg.get_fonts()})
images=sum(len(pg.get_images()) for pg in doc)
charts=sum(sh.has_chart for s in pp.slides for sh in s.shapes)
assert charts==3 and images==0,(charts,images)
assert all(sha(pth)==baseline[str(pth.relative_to(R))] for pth in master_paths)
assert any('Paperlogy' in f for f in fonts) and any('Pretendard' in f for f in fonts),fonts
assert len(doc)==10
report={'version':v,'pptx':str(OUT.relative_to(R)),'pdf':str(PDF.relative_to(R)),'pages':10,'native_charts_with_workbooks':charts,'raster_images_in_pdf':images,'text_mismatch':mismatch,'overflow':over,'shape_bounds':bounds,'fonts':fonts,'master_hashes_unchanged':baseline,'sources':{str(NB):sha(NB),str(GUIDE):sha(GUIDE)},'visual_review':'pending','latest_second_review':'pending','hallmark_scope':'PPTに適用可能な構造・誠実性・一色強調・反テンプレート原則のみ。CSS専用lintはPPTX非対応のため対象外'.replace('PPTに適用可能な構造・誠実性・一色強調・反テンプレート原則のみ。CSS専用lintはPPTX非対応のため対象外','PPT에 해당하는 정보 구조·수치 정직성·단일 강조색·반복 카드 회피 원칙 적용 / CSS 전용 lint는 PPTX 미지원으로 대상 제외')}
(Q/'validation.json').write_text(json.dumps(report,ensure_ascii=False,indent=2),encoding='utf-8')
(W/'current_build.json').write_text(json.dumps(report,ensure_ascii=False,indent=2),encoding='utf-8')
(W/'디자인_및_구성근거.md').write_text('''# 수리 요청 프로세스 발표자료 디자인·구성

포트폴리오와 분리한 10장 분석 과제 발표자료

- 기존 디자인 스킬: Paperlogy 6 SemiBold 제목 / Pretendard 본문 / 16:9 / PPTX·PDF 동시 관리와 실제 렌더 검수
- frontend-design: 수리 요청의 실제 업무 흐름과 대기 시간을 중심 소재로 선택 / 표지 타이포·표본 분기·생키·시간 비중·비교 차트·운영 조치표로 구조 변화
- ui-ux-pro-max: 로컬 디자인 시스템 검색에서 산업용 분석의 Swiss·최소 장식·높은 대비를 채택 / 웹 랜딩 페이지 CTA·영문 폰트 추천은 발표자료에 맞지 않아 제외 / 기존 국문 폰트 유지
- Hallmark: 꾸민 카드 반복·파이·장식 아이콘·그라데이션 배제 / 단일 청록 강조색 / 검증된 수치만 사용 / 웹 CSS·테마 무작위 추첨·웹 렌더 명령은 PPTX 형식에 적용하지 않음
- 색 토큰: 흰 배경 FFFFFF / 본문 202C32 / 보조 52616A / 강조 087F75 / 연강조 E5F2EF / 구분 D5DEE1
- 표지와 모든 제목은 한 가지 글자색 / 단일 강조색은 주요 구간·막대·생키 분기에만 사용
- 대안 검토: 어두운 대시보드형은 발표 화면의 정보 밀도로 제외 / 포스터형은 과정 설명 부족으로 제외 / 흐름도와 근거 비교를 중심으로 한 분석 보고서형 선택
- 표지 1,000·927·339 / 전처리 1,000·402·339 / 개별 업무 시간 3,151쌍은 단위 구분
- p03 일반927·예외73의 생키 폭만 정량 / 이후 순서선은 비례하지 않는 요약 / ExternRepair 완료 추정 없음
- p05 70.8%는 구간 시간 합계 비중 / 케이스별 대기 비중 중앙값66.3%와 구분
- p06·07·08은 편집용 통합문서를 내장한 PowerPoint 기본 차트 / 나머지 도식도 모든 요소가 편집 가능한 기본 도형·텍스트
- 분석 전체를 재실행하지 않고 사용자가 지정한 노트북 저장 출력과 가이드의 수치를 대조 / 개선안은 제안이며 실행 성과 아님
- 발표자 노트에 각 장의 원가이드·정의·한계와 근거를 포함

평가기준 대응: p02 전처리 / p03 프로세스 / p04–06 반복·대기 / p07–08 관련 요소 / p09–10 개선 제안과 효과 검증
''',encoding='utf-8')
print(json.dumps({'version':v,'pptx':str(OUT),'pdf':str(PDF),'pages':10,'charts':charts,'images':images,'overflow':over,'mismatch':mismatch,'previews':str(Q),'fonts':fonts},ensure_ascii=False,indent=2))
