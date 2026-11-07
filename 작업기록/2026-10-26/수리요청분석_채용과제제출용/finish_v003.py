from pathlib import Path
import shutil
from pptx import Presentation
from pptx.util import Inches
W=Path(__file__).resolve().parent;V=W/'versions';Q=W/'검수_v003';out=V/'repair2_채용과제_분석결과_v003.pptx';pdf=out.with_suffix('.pdf')
draft=Q/'제작중보존';draft.mkdir(exist_ok=True)
for f in [out,pdf]:shutil.copy2(f,draft/f.name)
p=Presentation(out)
for i,s in enumerate(p.slides):
 for sh in s.shapes:
  if not sh.has_text_frame:continue
  if abs(sh.top/Inches(1)-6.46)<.02 and abs(sh.left/Inches(1)-1.6)<.02:
   sh.top=Inches(6.44);sh.height=Inches(.56)
  if i==2 and sh.text in ['13개','16종']:sh.height=Inches(.48)
  if i==3 and sh.text.startswith('업무 구성 1,000건'):sh.height=Inches(.56)
p.save(out)
code=(W/'revise_v003.py').read_text(encoding='utf-8-sig').replace("txt(s,1.6,6.46,6.42,.51,","txt(s,1.6,6.44,6.42,.56,").replace("txt(s,x+.2,2.62,1.25,.4,","txt(s,x+.2,2.62,1.25,.48,").replace("txt(s,8.4,5.32,4.03,.45,","txt(s,8.4,5.32,4.03,.56,")
(W/'revise_v003.py').write_text(code,encoding='utf-8')
import json
import pymupdf as fitz
import win32com.client
qa=(W/'revise_v002.py').read_text(encoding='utf-8-sig')
qa=qa[qa.index("app=win32com.client.Dispatch"):].replace("'version':2","'version':3")
exec(compile(qa,'qa_v003','exec'))
