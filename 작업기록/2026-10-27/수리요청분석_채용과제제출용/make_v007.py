from pathlib import Path
import shutil

W = Path(__file__).resolve().parent
code = (W / 'build_v006.py').read_text(encoding='utf-8')
code = code.replace("Q = W / '검수_v006'", "Q = W / '검수_v007'")
code = code.replace("out = V / 'repair2_채용과제_분석결과_v006.pptx'", "out = V / 'repair2_채용과제_분석결과_v007.pptx'")
code = code.replace("prev_ppt = Presentation(V / 'repair2_채용과제_분석결과_v005.pptx')", "prev_ppt = Presentation(V / 'repair2_채용과제_분석결과_v006.pptx')")
code = code.replace("'노현동 / 데이터 분석 채용 과제 / 2026.10.26'", "'노현동 / 데이터 분석 / 2026.10.26'")
code = code.replace("'DATA ANALYSIS ASSIGNMENT'", "'DATA ANALYSIS'")
code = code.replace("'version': 6,", "'version': 7,")
(W / 'build_v007.py').write_text(code, encoding='utf-8')
print('build_v007.py created')
