from pathlib import Path
import shutil,json
from pptx import Presentation
from pptx.util import Inches,Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE_TYPE
import pymupdf as fitz
import win32com.client
W=Path(__file__).resolve().parent;Q=W/'검수_v002'
out=W/'versions/repair2_채용과제_분석결과_v002.pptx';pdf=out.with_suffix('.pdf')
for f in [out,pdf]:shutil.copy2(f,Q/'제작중보존'/('연결선보강전_'+f.name))
p=Presentation(out)
for sh in p.slides[4].shapes:
 if sh.shape_type==MSO_SHAPE_TYPE.LINE and Inches(2.8)<sh.top<Inches(6.1):
  sh.line.color.rgb=RGBColor.from_string('70838A');sh.line.width=Pt(1.5)
  if abs(sh.left/Inches(1)-2.59)<.02 and abs(sh.top/Inches(1)-5.31)<.02:sh.height=Inches(.18)
p.save(out)
app=win32com.client.Dispatch('PowerPoint.Application');deck=app.Presentations.Open(str(out),False,False,False);overflow=[]
try:
 for sl in deck.Slides:
  for sh in sl.Shapes:
   if sh.HasTextFrame and sh.TextFrame.HasText:
    tr=sh.TextFrame.TextRange
    if tr.BoundHeight>sh.Height+2 or tr.BoundWidth>sh.Width+2:overflow.append([sl.SlideIndex,tr.Text])
 deck.Save();deck.SaveAs(str(pdf),32)
finally:deck.Close()
d=fitz.open(pdf);pp=Presentation(out);norm=lambda t:''.join(t.split());mismatch=[]
for i,s in enumerate(pp.slides):
 for sh in s.shapes:
  if sh.has_text_frame and norm(sh.text) not in norm(d[i].get_text()):mismatch.append([i+1,sh.text])
d[4].get_pixmap(matrix=fitz.Matrix(1.5,1.5)).save(Q/'page-05.png')
r=json.loads((Q/'validation.json').read_text(encoding='utf-8'));r['overflow']=overflow;r['text_mismatch']=mismatch
(Q/'validation.json').write_text(json.dumps(r,ensure_ascii=False,indent=2),encoding='utf-8')
print(json.dumps({'overflow':overflow,'mismatch':mismatch},ensure_ascii=False))
