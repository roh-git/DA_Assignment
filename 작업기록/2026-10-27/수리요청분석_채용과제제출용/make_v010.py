from pathlib import Path

W = Path(__file__).resolve().parent
code = (W / 'build_v009.py').read_text(encoding='utf-8')
code = code.replace("Q = W / '검수_v009'", "Q = W / '검수_v010'")
code = code.replace("out = V / 'repair2_채용과제_분석결과_v009.pptx'", "out = V / 'repair2_채용과제_분석결과_v010.pptx'")
code = code.replace("prev_ppt = Presentation(V / 'repair2_채용과제_분석결과_v008.pptx')", "prev_ppt = Presentation(V / 'repair2_채용과제_분석결과_v009.pptx')")

# Remove all ' / 상세 보고서 ...쪽' from bottom footer source strings
code = code.replace("'1–5 / 8–9 / 상세 보고서 5–7쪽'", "'1–5 / 8–9'")
code = code.replace("'9-1–9-5 / 상세 보고서 5–7쪽'", "'9-1–9-5'")
code = code.replace("'12-1–12-2 / 상세 보고서 9쪽'", "'12-1–12-2'")
assert '상세 보고서' not in code, "상세 보고서 still present in slide code"

code = code.replace("'version': 9,", "'version': 10,")
(W / 'build_v010.py').write_text(code, encoding='utf-8')
print('build_v010.py created')
