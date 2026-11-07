from pathlib import Path
import json,shutil
from pptx import Presentation
from pptx.util import Inches
import pymupdf as fitz
import win32com.client
W=Path(__file__).resolve().parent;Q=W/'검수_v002';draft=Q/'제작중보존';draft.mkdir(exist_ok=True)
out=W/'versions/repair2_채용과제_분석결과_v002.pptx';pdf=out.with_suffix('.pdf')
for f in [out,pdf]:
 target=draft/f.name
 assert not target.exists()
 shutil.copy2(f,target)
p=Presentation(out)
for sh in p.slides[9].shapes:
 if sh.has_chart:
  for plot in sh.chart.plots:
   plot.data_labels.number_format='0.0'
   plot.data_labels.number_format_is_linked=False
 if sh.has_text_frame and Inches(4.65)<sh.top<Inches(6.19):
  sh.top-=Inches(.1)
for sh in p.slides[12].shapes:
 if sh.has_text_frame and '케이스 전체 처리 시간' in sh.text and sh.left>Inches(8):
  sh.height=Inches(.72)
 if not sh.has_text_frame or not sh.text:
  if abs(sh.left/Inches(1)-8.45)<.02 and abs(sh.top/Inches(1)-2.5)<.02:sh.height=Inches(1.42)
p.save(out)
# Keep the reproducible generator consistent with the visual corrections.
code=(W/'revise_v002.py').read_text(encoding='utf-8-sig')
code=code.replace("],200,accents=", "],200,fmt='0.0',accents=")
code=code.replace("rect(s,8.45,2.5,4.1,1.28,'ink')","rect(s,8.45,2.5,4.1,1.42,'ink')")
code=code.replace("txt(s,8.65,3.06,3.7,.53,","txt(s,8.65,3.06,3.7,.72,")
code=code.replace("txt(s,8.4,4.67,","txt(s,8.4,4.57,").replace("txt(s,x,5.06,","txt(s,x,4.96,").replace("y=5.39+i*.27","y=5.29+i*.27")
(W/'revise_v002.py').write_text(code,encoding='utf-8')
app=win32com.client.Dispatch('PowerPoint.Application');deck=app.Presentations.Open(str(out),False,False,False);overflow=[]
try:
 for sl in deck.Slides:
  for sh in sl.Shapes:
   if sh.HasTextFrame and sh.TextFrame.HasText:
    tr=sh.TextFrame.TextRange
    if tr.BoundHeight>sh.Height+2 or tr.BoundWidth>sh.Width+2:overflow.append({'page':sl.SlideIndex,'text':tr.Text,'h':sh.Height,'bh':tr.BoundHeight,'w':sh.Width,'bw':tr.BoundWidth})
 deck.Save();deck.SaveAs(str(pdf),32)
finally:deck.Close()
d=fitz.open(pdf);pp=Presentation(out);norm=lambda t:''.join(t.split());mismatch=[]
for i,s in enumerate(pp.slides):
 for sh in s.shapes:
  if sh.has_text_frame and norm(sh.text) not in norm(d[i].get_text()):mismatch.append([i+1,sh.text])
for i in [9,12]:d[i].get_pixmap(matrix=fitz.Matrix(1.5,1.5)).save(Q/f'page-{i+1:02}.png')
for start in [8,12]:
 sheet=fitz.open();pg=sheet.new_page(width=1600,height=900)
 for j,i in enumerate(range(start,min(start+4,13))):pg.show_pdf_page(fitz.Rect(j%2*800,j//2*450,(j%2+1)*800,(j//2+1)*450),d,i)
 pg.get_pixmap().save(Q/f'sheet-{start+1:02}.png')
r=json.loads((Q/'validation.json').read_text(encoding='utf-8'));r['overflow']=overflow;r['text_mismatch']=mismatch
(Q/'validation.json').write_text(json.dumps(r,ensure_ascii=False,indent=2),encoding='utf-8')
assert all(x in d[9].get_text() for x in ['11.6','160.4','71.4','25건','127건','8건','25.9시간','0.9시간'])
print(json.dumps({'overflow':overflow,'text_mismatch':mismatch,'chart_decimal_labels':'pass'},ensure_ascii=False))
