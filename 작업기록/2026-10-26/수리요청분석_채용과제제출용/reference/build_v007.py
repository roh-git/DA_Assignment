from pathlib import Path
import ast,json,hashlib,re,shutil
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
R=Path.cwd();D=R/'작업기록/2026-10-26';W=D/'수리요청분석';V=D/'versions/수리요청분석';Q=W/'검수_v007';Q.mkdir(exist_ok=True)
out=V/'repair2_수리요청프로세스_분석발표_v007.pptx';pdf=out.with_suffix('.pdf');assert not out.exists()
sha=lambda f:hashlib.sha256(f.read_bytes()).hexdigest()
masters={str(x.relative_to(R)):sha(x) for x in [R/'노현동_포트폴리오.pdf',R/'최종본/노현동_포트폴리오.pptx']}
module=ast.parse((V/'build_v004.py').read_text(encoding='utf-8-sig'));sel=[]
for n in module.body:
 if isinstance(n,ast.FunctionDef) and n.name in {'rgb','rect','txt','line','node','bar','dot','ribbon'}:sel.append(n)
 if isinstance(n,ast.Assign) and any(isinstance(t,ast.Name) and t.id in {'C','F'} for t in n.targets):sel.append(n)
exec(compile(ast.Module(body=sel,type_ignores=[]),'native_helpers','exec'))
p=Presentation();p.slide_width=Inches(13.333333);p.slide_height=Inches(7.5);p.core_properties.title='수리 요청 프로세스 분석 및 개선 제안';p.core_properties.author='Roh Hyundong'
G=(W/'sources/PPT_분석스토리_수정가이드_원본.md').read_text(encoding='utf-8')
page_notes=[]
def page(role,title,intent,source,conclusion,foot=''):
 s=p.slides.add_slide(p.slide_layouts[6]);s.background.fill.solid();s.background.fill.fore_color.rgb=rgb('bg')
 txt(s,.65,.25,12.03,.27,role,11,color='muted')
 txt(s,.65,.68,12.03,1.01,title,27,font='head')
 txt(s,.65,1.79,12.03,.58,intent,16,color='muted')
 txt(s,.75,6.22,11.8,.51,conclusion,14.5)
 if foot:txt(s,.65,6.8,12.03,.21,foot,9.5,color='muted')
 line(s,.65,7.08,12.68,7.08);txt(s,.65,7.16,11.3,.2,'repair2 분석 / '+source,9.5,color='muted');txt(s,12.1,7.13,.58,.25,f'{len(p.slides):02}',10,color='muted',align=PP_ALIGN.RIGHT)
 return s

def note(s,n,explanation):
 m=re.search(r'### '+f'{n:02}'+r'\..*?(?=\n### |\n## |\Z)',G,re.S)
 t=m.group(0) if m else ''
 t=t.replace('분기된다','나뉜다').replace('분기되는','나뉘는').replace('업무 분기','업무 구분').replace('분기와','구분과')
 s.notes_slide.notes_text_frame.text='발표 설명\n'+explanation+'\n\n제작 기준과 해석 범위\n'+t+'\n모든 시간은 날짜 정확성이 미검증된 탐색적인 기록상 경과 시간이며 인과관계나 확정된 절감효과가 아님'
 page_notes.append(explanation)

def task(s,x,y,w,name,meaning):
 rect(s,x,y,w,.75,'surface');txt(s,x+.07,y+.04,w-.14,.3,name,13,bold=True,align=PP_ALIGN.CENTER);txt(s,x+.07,y+.39,w-.14,.28,meaning,12.5,align=PP_ALIGN.CENTER)
# 01
s=p.slides.add_slide(p.slide_layouts[6]);s.background.fill.solid();s.background.fill.fore_color.rgb=rgb('ink')
txt(s,.75,.4,11.8,.3,'분석 개요',11,color='white');line(s,.75,.9,12.58,.9,'muted')
txt(s,.75,1.45,11.8,1.6,'수리 요청 프로세스\n분석 및 개선 제안',39,color='white',font='head')
txt(s,.78,3.52,11.5,.85,'수리 요청의 업무 흐름을 파악하고\n어디에서 시간이 길어지는지 살펴봤다',21,color='white')
for x,val,label in [(.8,'1,000','전체 수리 요청 / 건'),(4.8,'927','일반 경로 / 건'),(8.8,'339','시간 분석 대상 / 건')]:
 txt(s,x,5.08,3.4,.73,val,38,color='white',font='head');txt(s,x,5.93,3.5,.4,label,16,color='white')
txt(s,.8,6.94,11.8,.25,'노현동 / 분석 과제 발표 / 2026.10.26',11,color='white')
note(s,1,'처리 시간이 길어지는 구간과 반복 업무를 찾기 위해 수리 요청1,000건을 분석했습니다. 먼저 데이터의 신뢰성과 업무 흐름을 확인했고, 분석 가능한339건에서 업무 처리와 업무 사이 대기를 나누어 살펴봤습니다. 그 결과를 추가 기록과 개선 제안으로 연결하겠습니다.')
# 02
s=page('데이터 개요와 전처리','데이터 품질을 점검하고\n분석 목적에 따라 대상을 구분했다','업무 흐름과 시간 계산에 각각 어떤 기록을 사용할 수 있는지 먼저 확인했다','1–5 / 9-5','날짜를 임의로 고치지 않고 분석 범위를 나눴다\n시간 결과는 탐색적으로 해석하고 업무 흐름은 전체 1,000건에서 확인했다','날짜·시간 결측13행 / 1905·1970·2020–2022년 혼재 / 1970년 기록의 실제 날짜 정확성은 미확인')
node(s,.75,2.76,2.5,1.03,'전체 1,000건\n13,262개 이벤트',size=18)
line(s,3.25,3.28,3.65,3.28);line(s,3.65,3.02,3.65,4.76);line(s,3.65,3.02,4.1,3.02,arrow=True);line(s,3.65,4.76,4.1,4.76,arrow=True)
node(s,4.1,2.63,3.3,.79,'업무 흐름 분석 1,000건',accent=True,size=18);txt(s,7.85,2.68,4.6,.7,'시각이 빠져 있어도\n업무 수행 기록은 보존',16)
node(s,4.1,4.22,3.3,1.08,'시각 완전 케이스 402건\n모든 이벤트가 1970년 기록',size=16);line(s,7.4,4.76,8.15,4.76,arrow=True)
node(s,8.15,4.22,4.4,1.08,'일반 경로 시간 분석 339건\n날짜 정확성은 별도 검증 필요',accent=True,size=16)
line(s,.75,5.61,12.55,5.61);txt(s,.75,5.77,11.8,.38,'개별 업무 시간은 start·complete 1:1 연결 3,631쌍 중 조건을 충족한 3,151쌍으로 계산',15)
note(s,2,'업무 흐름과 시간 계산은 필요한 데이터 조건이 다릅니다. 업무 기록 자체는1,000건 모두 보존했습니다. 시각이 빠지지 않고 연도가1970년으로 일치하는402건 중 일반 경로339건을 시간 분석에 사용했습니다. 이는 실제 날짜가 검증됐다는 뜻은 아닙니다. 개별 업무의 처리 시간은 별도로 선별한3,151쌍을 사용합니다.')
# 03
s=page('업무 흐름 분석','수리 요청은 일반·예외 경로로 나뉘고\n수리 업무도 여러 형태로 기록됐다','시간과 반복 업무를 분석하기 전에 케이스별 수행 업무와 종료 경로를 확인했다','8-2 / 10-3–10-5','일반 경로에서도 여러 수리 업무를 함께 수행하거나 같은 업무를 반복할 수 있다','구조 요약 / FirstContact1건 미기록 / ExternRepair 완료 시각 없음 / 종결 관련 두 업무의 순서는 고정하지 않음')
txt(s,.75,2.53,3.05,.38,'전체 요청 1,000건',18,bold=True)
ribbon(s,1.05,2.53,3.01,2.98,1.2*.927,'light');ribbon(s,1.05,2.53,3.01+1.2*.927,5.22,1.2*.073,'bar')
rect(s,.93,3.01,.12,1.2,'ink');rect(s,2.53,2.98,.12,1.2*.927,'accent');rect(s,2.53,5.22,.12,1.2*.073,'bar')
txt(s,2.8,3.0,1.15,.65,'일반\n927건',16,bold=True);txt(s,2.8,5.09,1.1,.44,'예외 73건',14,bold=True);txt(s,.75,4.42,1.9,.58,'FirstContact\n최초 접수',13)
line(s,2.59,5.31,2.59,5.59,'muted',1.5,arrow=True);node(s,.75,5.61,3.03,.52,'InformClientWrongPlace',size=12.5)
line(s,3.8,3.35,4.16,3.35,arrow=True)
for x,w,a,b in [(4.16,1.94,'MakeTicket','수리 요청 등록'),(6.55,2.35,'ArrangeSurvey','현장 조사 일정 조율'),(9.35,2.35,'Survey','현장 조사')]:task(s,x,2.98,w,a,b)
line(s,6.1,3.35,6.55,3.35,arrow=True);line(s,8.9,3.35,9.35,3.35,arrow=True)
txt(s,6.55,3.8,5.5,.27,'InformClientSurvey: 현장 조사 안내',11,color='muted')
line(s,10.525,3.73,10.525,4.12);line(s,5.225,4.12,11.025,4.12)
for x,a,b in [(4.05,'InternRepair','내부 수리'),(6.95,'ImmediateRepair','즉시 수리'),(9.85,'ExternRepair','외부 수리')]:
 line(s,x+1.175,4.12,x+1.175,4.3,arrow=True);task(s,x,4.3,2.35,a,b);line(s,x+1.175,5.05,x+1.175,5.27)
line(s,5.225,5.27,11.025,5.27);line(s,5.225,5.27,5.225,5.61,arrow=True)
node(s,4.05,5.61,2.35,.52,'RepairReady',size=13);line(s,6.4,5.87,6.72,5.87,arrow=True)
node(s,6.72,5.54,3.3,.66,'ReadyInformClient\nSendTicketToFinAdmin',size=12);line(s,10.02,5.87,10.3,5.87,arrow=True);node(s,10.3,5.61,2.0,.52,'TicketReady',size=13)
note(s,3,'전체 요청은 일반 경로927건과 잘못 접수된 요청의 예외 경로73건으로 나뉩니다. 일반 경로는 요청 등록과 조사 일정 조율, 현장 조사를 거쳐 수리 관련 업무로 이어집니다. 내부·즉시·외부 수리가 한 가지로만 기록되는 것은 아닙니다. 생키의 폭은927건과73건의 구성만 표현하며 오른쪽 연결선은 공통 업무 흐름의 요약입니다. 그래서 다음에는 복합 수리와 같은 업무의 반복 기록을 확인했습니다.')
# 04
s=page('반복 작업 분석','일부 요청에서 여러 수리 업무와\n같은 업무의 반복 기록이 관찰됐다','불필요한 반복이 있는지 살펴보기 위해 수리 업무의 종류와 시작·완료 기록을 비교했다','10-3 / 10-5 / 12-5','반복 기록만으로 재작업이나 불필요한 반복을 확정할 수는 없다\n이어서 전체 처리 시간 중 어느 구간에 시간이 집중되는지 살펴봤다','복합 수리=두 종류 이상 수리 업무 / 복합91건과 반복38건은 겹칠 수 있으므로 합산하지 않음')
txt(s,.75,2.57,6.5,.4,'일반 경로 927건의 수리 업무 구성',18,bold=True)
rect(s,.75,3.35,6.35*836/927,.64,'bar');rect(s,.75+6.35*836/927,3.35,6.35*91/927,.64,'accent')
txt(s,.75,4.12,3.7,.76,'단일 수리 836건\n90.2%',19);txt(s,4.65,4.12,2.75,.76,'복합 수리 91건\n9.8%',19)
line(s,7.65,2.62,7.65,5.02);txt(s,8.08,2.57,4.5,.4,'InternRepair 반복 기록',18,bold=True);txt(s,8.08,3.14,4.5,.72,'38건',35,font='head')
txt(s,8.08,4.03,4.48,.95,'InternRepair 시각이 모두 1970년인\n29건을 시간순으로 확인했고\n28건에서 완료 후 다시 시작',14)
line(s,.75,5.25,12.55,5.25);txt(s,.75,5.44,2.5,.61,'실제 기록 순서 예시\ncaseID 124',13,color='muted')
line(s,3.7,5.72,10.3,5.72,'bar')
for x,t,ev in [(3.7,'11:41','start'),(5.9,'15:21','complete'),(8.1,'19:45','start'),(10.3,'23:03','complete')]:
 dot(s,x,5.72);txt(s,x-.5,5.35,1.0,.28,t,13,align=PP_ALIGN.CENTER);txt(s,x-.6,5.86,1.2,.29,ev,12.5,align=PP_ALIGN.CENTER)
txt(s,11.2,5.45,1.3,.6,'동일 일자\n간격 비례 아님',11,color='muted')
note(s,4,'일반 경로927건 중91건에는 두 종류 이상의 수리 업무가 기록됐습니다. 별도로 InternRepair 반복 기록은38건입니다. 두 집단은 겹칠 수 있습니다. 반복38건 중 InternRepair의 모든 시각이1970년으로 기록된29건을 시간순으로 확인했고28건에서 완료 후 다시 시작하는 형태를 봤습니다. 하지만 정상적인 복수 수행인지 재작업인지는 현재 기록으로 알 수 없습니다.')
# 05
s=page('전체 처리 시간 분석','조사 일정 조율 완료 후 현장 조사 시작까지\n기록상 대기가 가장 큰 비중을 차지했다','업무를 수행하는 시간과 다음 업무를 기다리는 시간을 나누어 우선 살펴볼 구간을 찾았다','12-2-1 / 12-4-1','70.8% 전체가 줄일 수 있는 시간이라는 뜻은 아니다\n이 구간의 긴 대기가 일부 요청에만 나타나는지 분포를 확인했다','비중=구간 시간 합계/전체 시간 합계 / 반올림 합계100.1% / 위 흐름도의 간격은 시간 비례가 아님')
txt(s,.75,2.48,11.8,.34,'시간 분석 339건 / 기록상 경과 시간 합계 비중 / 전체 처리 시간 중앙값 77.3시간',14,bold=True)
events=[('FirstContact','complete'),('ArrangeSurvey','complete'),('Survey','start'),('RepairReady','complete'),('TicketReady','complete')]
for j,(a,b) in enumerate(events):node(s,.65+2.5*j,3.0,2.03,.74,a+'\n'+b,accent=j in [1,2],size=14)
for j,val in enumerate([1.9,70.8,22.5,4.9]):
 x=.65+2.5*j;line(s,x+2.03,3.37,x+2.5,3.37,'accent' if j==1 else 'bar',2.4 if j==1 else 1.3,arrow=True);txt(s,x+1.58,3.98,1.4,.42,f'{val:.1f}%',21,bold=True,align=PP_ALIGN.CENTER)
rect(s,.65,4.64,12.03,1.04,'ink',rad=True);rect(s,.9,4.83,11.53,.31,'bar',rad=True);rect(s,.9,4.83,11.53*.708,.31,'accent',rad=True);rect(s,.9+11.53*.708-.14,4.83,.14,.31,'accent');rect(s,.9+11.53*.708-.025,4.83,.05,.31,'ink')
txt(s,.9,5.28,8.2,.28,'ArrangeSurvey complete → Survey start  70.8%',14,color='white',bold=True);txt(s,9.5,5.28,2.93,.28,'나머지 세 구간',14,color='white',align=PP_ALIGN.RIGHT)
txt(s,.75,5.84,11.8,.28,'22.5% 구간에는 현장 조사(Survey)와 수리 관련 경과 시간이 함께 포함된다',13,color='muted')
note(s,5,'일반 경로 시간 표본339건의 전체 경과 시간을 네 구간으로 나눴습니다. 조사 일정 조율 완료부터 현장 조사 시작까지가 전체 시간 합계의70.8%였습니다. 이는 케이스별 비중의 평균이나 중앙값이 아니라 구간 시간 합계의 비중입니다. 정상적인 예약 대기일 수도 있으므로 이 비율을 절감 가능 시간으로 해석하지 않았습니다.')
# 06
s=page('주요 대기 구간 분석','현장 조사 시작 전 대기는\n절반 이상에서 하루를 초과했다','큰 시간 비중이 일부 긴 사례 때문인지 여러 요청에서 나타나는 현상인지 분포를 확인했다','12-3-1–12-3-3','조사 예정 시각이 없어 정상적인 예약 대기와 실제 일정 지연은 구분하기 어렵다\n다음으로 수리 유형에 따라서도 전체 처리 시간 차이가 있는지 살펴봤다','ArrangeSurvey complete→Survey start / 기록상 대기 / 24시간은 탐색 분석 경계이며 실제 SLA 기준이 아님')
txt(s,.75,2.48,7.35,.35,'조사 시작 전 대기 분포 / 339건 / 단위: 건',15,bold=True)
bar(s,.7,2.97,7.4,3.13,['0시간','0시간 초과–24시간 이하','24시간 초과–72시간 이하','72시간 초과–168시간 이하','168시간 초과'],[('케이스 수',[39,89,63,90,58],'bar')],120,accents={0:['bar','bar','accent','accent','accent']})
line(s,8.4,2.6,8.4,6.04)
for y,val,lab in [(2.57,'54.5시간','대기 시간 중앙값'),(3.8,'195.7시간','약 90%가 이 시간 이하(P90)'),(5.03,'62.2%','24시간 초과 211 / 339건')]:txt(s,8.82,y,3.74,.63,val,29,font='head');txt(s,8.82,y+.66,3.74,.37,lab,14.5)
note(s,6,'339건 중211건에서24시간을 넘는 기록상 대기가 관찰됐습니다. 중앙값54.5시간과 P90 195.7시간을 함께 보면 일부 극단값만의 현상은 아니라는 점을 확인할 수 있습니다. P90은 약90%가 해당 시간 이하인 경계이지 상위10%의 평균이 아닙니다. 예정된 조사 날짜가 없기 때문에 실제 일정 지연 여부는 아직 판단할 수 없습니다.')
# 07
s=page('수리 유형별 비교','E 유형은 전체 처리 시간과\n수리 시작 전 대기가 모두 길었다','전체 처리 시간의 차이가 수리 유형과 관련되는지 확인하고 같은 수리 업무끼리도 비교했다','14-1 / 14-4-1–14-4-2','같은 InternRepair 업무로 좁혀도 E 유형의 긴 대기가 유지됐다\n수리 유형이 지연의 원인이라는 결론은 아니며 접수 방식과 담당자 기록도 추가로 살펴봤다','전체339건·구간338건은 별도 표본 / 모든 시간은 기록상 경과 시간 / B·E·P의 의미를 임의 해석하지 않음')
txt(s,.75,2.47,7.2,.39,'케이스 전체 처리 시간 중앙값 / 339건 / 시간(h)',15,bold=True)
bar(s,.7,2.98,7.25,3.1,['B (46건)','E (143건)','P (150건)'],[('전체 처리 시간 중앙값',[11.6,160.4,71.4],'bar')],200,fmt='0.0',accents={0:['bar','accent','bar']})
line(s,8.3,2.6,8.3,6.05)
txt(s,8.68,2.48,3.9,.4,'수리 시작 전 대기 / 338건',16,bold=True);txt(s,8.68,3.04,3.9,.65,'E  25.2시간',29,font='head');txt(s,8.68,3.74,3.9,.34,'중앙값 / E 143건',14.5);txt(s,8.68,4.13,3.9,.32,'B 0.0h(46건) / P 0.0h(149건)',13.5)
line(s,8.68,4.62,12.55,4.62);txt(s,8.68,4.79,3.9,.52,'InternRepair 단독 1:1 연결\nE 유형 127건',15,bold=True)
txt(s,8.68,5.42,1.8,.43,'25.9h',23,font='head');txt(s,10.8,5.42,1.75,.43,'4.0h',23,font='head');txt(s,8.68,5.9,1.9,.28,'시작 전 대기',13);txt(s,10.8,5.9,1.75,.28,'실제 수리 업무',13)
note(s,7,'유형별 전체 처리 시간 중앙값은 B11.6시간, E160.4시간, P71.4시간으로 E가 가장 길었습니다. 다음으로 조사 완료 후 수리 시작 전 대기를 비교했습니다. 전체 시간339건과 달리 이 구간은338건이 비교 가능하며 E143건의 중앙값은25.2시간입니다. 같은 InternRepair 단독·시작완료1:1 조건으로 좁힌 E127건에서도 대기25.9시간과 실제 수리4.0시간이 확인됐습니다. 실제 수리4.0시간은 InternRepair start→complete이며 RepairReady까지의 경과 시간과 구분합니다.')
# 08 — fixed-width data table avoids native-chart category truncation
s=page('접수 채널과 담당자 분석','Web·Letter 접수에서 조사 일정 조율 업무의\n장시간 처리 사례가 집중됐다','이번에는 조사 일정 조율 업무 자체의 시간을 비교했다\n접수 채널과 시작·완료 담당자 기록에 따른 차이를 살펴봤다','12-1 / 13-3–13-4','Web은 중앙값부터 길었고 Letter는 짧은 중앙값에도 긴 처리 사례가 많았다\n채널별 처리 방식과 담당자가 다른 이유를 추가로 확인해야 한다','ArrangeSurvey start→complete 809건 / 앞선 조사 시작 전 대기와 다른 구간 / 담당자 차이를 인수인계·지연 원인으로 단정하지 않음')
rect(s,.75,2.63,7.1,.52,'ink')
for x,w,t in [(.9,1.25,'접수 채널'),(2.3,1.0,'건수'),(3.46,1.35,'중앙값(분)'),(5.05,2.58,'1시간 초과율')]:txt(s,x,2.68,w,.39,t,14,color='white',bold=True)
rows=[('Web',207,119,52.7),('Letter',175,5,48.6),('Phone',224,3,0),('Personal',203,4,0)]
for j,(channel,n,med,rate) in enumerate(rows):
 y=3.32+j*.57
 if j==0:rect(s,.75,y-.06,7.1,.54,'pale')
 txt(s,.9,y,1.3,.34,channel,16);txt(s,2.3,y,1.0,.34,str(n),16);txt(s,3.46,y,1.3,.34,str(med),16,bold=j==0)
 rect(s,5.05,y+.08,1.4,.15,'surface')
 if rate:rect(s,5.05,y+.08,1.4*rate/100,.15,'accent')
 txt(s,6.65,y,1.0,.34,f'{rate:.1f}%',16,align=PP_ALIGN.RIGHT);line(s,.85,y+.45,7.75,y+.45)
txt(s,.9,5.73,6.9,.36,'1시간 초과 194건은 모두 Web·Letter에서 발생',15.5,bold=True)
line(s,8.2,2.62,8.2,6.03);txt(s,8.6,2.53,3.95,.65,'시작·완료 담당자별 중앙값\nArrangeSurvey / 809건',16,bold=True)
for y,label,n,val in [(3.39,'동일 담당자','704건','4분'),(4.4,'서로 다른 담당자','105건','332분')]:
 txt(s,8.6,y,2.4,.33,label,15);txt(s,8.6,y+.38,2.0,.28,n,13,color='muted');txt(s,11.05,y,1.5,.66,val,26,font='head',align=PP_ALIGN.RIGHT)
txt(s,8.6,5.47,3.95,.63,'동일 담당자의 Web·Letter에서도\n1시간 초과 89건 관찰',14.5)
note(s,8,'이번 분석은 업무 사이의 대기가 아니라 ArrangeSurvey 업무 내부의 start→complete 시간입니다.809건에서1시간 초과194건이 모두 Web과Letter였습니다. 하지만 두 채널의 중앙값은119분과5분으로 다릅니다. 담당자 동일 집단은4분, 다른 집단은332분이었지만 동일 담당자의 두 채널에서도89건이1시간을 넘었습니다. 따라서 담당자 차이만으로 긴 시간을 설명하거나 실제 인수인계를 원인으로 단정할 수 없습니다. 막대는0~100% 기준이며 표의 원수치가 해석 기준입니다.')
# 09
s=page('개선 방안','긴 대기와 장시간 처리의 원인을 확인한 뒤\n개선안을 적용한다','관찰한 현상을 실제 개선으로 연결하기 위해 추가 기록과 가능한 운영 조치를 함께 정리했다','15-2–15-5','우선순위는 검토 순서이며 절감 효과가 입증된 순서가 아니다\n원인을 확인한 뒤 시범 적용하고 구간과 전체 처리 시간을 함께 비교한다','예정 시각·진행 상태·지연 사유는 신규 수집 항목 / 자재 부족·인력 과부하 등을 현재 결과만으로 단정하지 않음')
rect(s,.65,2.63,12.03,.56,'ink')
cols=[.8,1.6,5.32,9.03];widths=[.65,3.45,3.43,3.3]
for x,w,t in zip(cols,widths,['순위','관찰한 현상','추가로 기록할 정보','원인 확인 후 검토할 조치']):txt(s,x,2.67,w,.46,t,14.5,color='white',bold=True)
rows=[('1','조사 시작 전 대기\n전체 시간 합계의 70.8%','조사 예정 시각·배정 상태\n착수 상태와 지연 사유','예정 시각 초과 미착수 확인\n사유에 맞는 일정 조정'),('2','E 유형 수리 시작 전 대기\n중앙값 25.2시간','수리 예정 시각\n시작 전 진행 상태·지연 사유','반복되는 지연 원인 확인\n해당 절차와 일정 조정'),('3','Web·Letter 조사 일정 조율\n1시간 초과 사례 집중','채널별 진행 상태\n시작·완료 담당자가 다른 이유','실제 지연이 확인된\n세부 처리 절차 점검')]
for i,row in enumerate(rows):
 y=3.35+.9*i
 if i==0:rect(s,.65,y-.04,12.03,.83,'pale')
 for x,w,t in zip(cols,widths,row):txt(s,x,y+.05,w,.7,t,15 if x>1 else 22,bold=x<1,font='head' if x<1 else 'body')
 line(s,.75,y+.8,12.55,y+.8)
note(s,9,'첫째, 가장 큰 비중을 보인 조사 시작 전 대기는 예정 시각과 배정 상태를 기록해 실제 일정 초과 여부부터 확인합니다. 둘째, E유형은 수리 예정 시각과 시작 전 상태를 기록해 반복되는 대기 원인을 확인합니다. 셋째, Web·Letter는 세부 진행 상태와 담당자가 다른 이유를 남겨 처리 절차를 점검합니다. 우선순위는 검토 순서이며 기대 절감효과가 검증된 순서는 아닙니다.')
# 10
s=page('효과 검증과 분석 한계','실제 일정 지연이 줄었는지 확인하고\n전체 처리 시간으로 개선 효과를 판단한다','일부 업무가 빨라지는 데 그치지 않고 요청 전체의 처리 기간이 줄었는지 확인한다','9-4–9-5 / 15-5','현재 예정 시각이 없고 1970년 날짜의 정확성도 확인되지 않았다\n신뢰할 수 있는 새 기록으로 기준선을 세운 뒤 개선 전후를 비교해야 한다','실제 지연 원인 미확인 / ExternRepair 완료 시각 부재 / 계산식·조회 시점·집계 조건은 발표자 노트와 지표 정의 문서에 수록')
steps=[('추가 기록','예정·실제 시각과 지연 사유'),('원인 확인','예약 대기와 실제 지연 구분'),('시범 개선','확인된 원인에 맞는 조치'),('다시 측정','구간과 전체 처리 시간 비교')]
for i,(a,b) in enumerate(steps):
 x=.65+3.15*i;node(s,x,2.65,2.55,1.04,a+'\n'+b,accent=i==3,size=14)
 if i<3:line(s,x+2.55,3.17,x+3.15,3.17,'bar',1.4,arrow=True)
txt(s,.75,4.02,7.4,.4,'구간별로 확인할 변화',18,bold=True)
for y,t in [(4.57,'조사 예정 시각 초과 미착수율과 실제 착수 지연'),(5.02,'조사 시작 전 대기와 수리 시작 전 대기'),(5.47,'조사 일정 조율 업무의 1시간 초과율')]:txt(s,.75,y,7.4,.35,t,15.5)
rect(s,8.62,4.0,3.95,2.03,'ink');txt(s,8.9,4.18,3.38,.35,'최종 결과 지표',17,color='white',bold=True);txt(s,8.9,4.76,3.38,.84,'케이스 전체 처리 시간\n중앙값·P90',22,color='white',font='head');txt(s,8.9,5.68,3.38,.25,'수리 유형·접수 채널별 전후 비교',12.5,color='white')
technical=(W/'지표정의_계산기준_v006.md').read_text(encoding='utf-8')
note(s,10,'추가 기록으로 정상 예약 대기와 실제 지연을 구분하고 확인된 원인에 맞게 시범 개선합니다. 이후 조사와 수리 시작 전 대기, 조사 일정 조율의 장시간 처리율을 비교합니다. 최종 판단은 케이스 전체 처리 시간 중앙값과P90입니다. 수리 유형과접수 채널별로 비교하고 지연이 다른 구간으로 옮겨갔는지도 확인합니다. 기존 날짜의 정확성과예정시각에한계가있어 신뢰할수있는새기준선이필요합니다.\n\n'+technical)
# Save clean native objects without inherited Office shadows
for s in p.slides:
 for sh in s.shapes:
  for ef in sh._element.xpath('.//a:effectRef'):ef.set('idx','0')
  for sp in sh._element.xpath('./p:spPr'):
   for old in list(sp):
    if old.tag.endswith('}effectLst') or old.tag.endswith('}effectDag'):sp.remove(old)
   sp.append(OxmlElement('a:effectLst'))
p.save(out)
app=win32com.client.Dispatch('PowerPoint.Application');deck=app.Presentations.Open(str(out),False,False,False);overflow=[]
try:
 for sl in deck.Slides:
  for sh in sl.Shapes:
   if sh.HasTextFrame and sh.TextFrame.HasText:
    tr=sh.TextFrame.TextRange
    if tr.BoundHeight>sh.Height+2 or tr.BoundWidth>sh.Width+2:overflow.append({'page':sl.SlideIndex,'text':tr.Text,'height':sh.Height,'boundHeight':tr.BoundHeight,'width':sh.Width,'boundWidth':tr.BoundWidth})
 deck.Save();deck.SaveAs(str(pdf),32)
finally:deck.Close()
pp=Presentation(out);doc=fitz.open(pdf);norm=lambda t:''.join(t.split());mismatch=[];bounds=[]
for i,s in enumerate(pp.slides):
 for sh in s.shapes:
  if sh.has_text_frame and norm(sh.text) not in norm(doc[i].get_text()):mismatch.append([i+1,sh.text])
  if sh.left<0 or sh.top<0 or sh.left+sh.width>pp.slide_width+1000 or sh.top+sh.height>pp.slide_height+1000:bounds.append([i+1,sh.name])
 doc[i].get_pixmap(matrix=fitz.Matrix(1.5,1.5)).save(Q/f'page-{i+1:02}.png')
for start in [0,5]:
 sheet=fitz.open();pg=sheet.new_page(width=1600,height=1350)
 for j,i in enumerate(range(start,start+5)):pg.show_pdf_page(fitz.Rect(j%2*800,j//2*450,(j%2+1)*800,(j//2+1)*450),doc,i)
 pg.get_pixmap().save(Q/f'전체-{start+1:02}.png')
text='\n'.join(pg.get_text() for pg in doc)
assert '분기' not in text and '조회 시각' not in doc[9].get_text()
assert all(t in doc[7].get_text() for t in ['119','52.7%','48.6%','332분','704건','105건','89건'])
assert '339건' in doc[4].get_text()
assert len(doc)==10 and sum(len(pg.get_images()) for pg in doc)==0
assert not mismatch and not bounds,(mismatch,bounds)
assert all(sha(R/k)==h for k,h in masters.items())
report={'version':7,'pptx':str(out.relative_to(R)),'pdf':str(pdf.relative_to(R)),'pages':10,'changed_pages':list(range(1,11)),'native_charts_with_workbooks':2,'native_rate_bars':4,'raster_images_in_pdf':0,'text_mismatch':mismatch,'overflow':overflow,'shape_bounds':bounds,'fonts':sorted({z[3] for pg in doc for z in pg.get_fonts()}),'master_hashes_unchanged':masters,'visual_review':'pending','latest_second_review':'pending','content_basis':['원본 노트북 저장 출력','최종 제작 가이드','PPT 수정 체크리스트','분석스토리 수정 가이드','사용자 일상어·편집형 그래프·별도관리 지시'],'design_review':'v006의 일괄압축을 폐기 / v005의 글자크기와 여백체계 기반으로 내용별 재배치 / p08 고정열 표·도형막대로 중앙값 잘림 해결'}
(Q/'validation.json').write_text(json.dumps(report,ensure_ascii=False,indent=2),encoding='utf-8')
(W/'발표멘트_v007.md').write_text('# 발표 멘트 / v007\n\n'+'\n\n'.join(f'## {i+1:02}페이지\n\n'+t for i,t in enumerate(page_notes)),encoding='utf-8')
print(json.dumps({'version':7,'pages':10,'overflow':overflow,'mismatch':mismatch,'previews':str(Q)},ensure_ascii=False,indent=2))
