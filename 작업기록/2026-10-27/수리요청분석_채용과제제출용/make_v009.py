from pathlib import Path
import shutil

W = Path(__file__).resolve().parent
S = W / 'sources'
tmp_img = Path(r'C:\Users\rhd67\AppData\Local\Temp\orca-paste-1793064066680-8a8096d5-fde9-4932-9258-c84863dea2ed.png')
if tmp_img.exists():
    shutil.copy2(tmp_img, S / 'v009_slide04_align_feedback.png')

code = (W / 'build_v008.py').read_text(encoding='utf-8')
code = code.replace("Q = W / '검수_v008'", "Q = W / '검수_v009'")
code = code.replace("out = V / 'repair2_채용과제_분석결과_v008.pptx'", "out = V / 'repair2_채용과제_분석결과_v009.pptx'")
code = code.replace("prev_ppt = Presentation(V / 'repair2_채용과제_분석결과_v007.pptx')", "prev_ppt = Presentation(V / 'repair2_채용과제_분석결과_v008.pptx')")

# 1. Slide 4 Title replacement
code = code.replace(
    "'날짜 이상을 확인하고 임의 보정 대신 분석 대상을 구분했다'",
    "'날짜 기록의 정합성을 검증하고 분석 대상을 구분했다'"
)

# 2. Slide 4 verification subtext replacement
code = code.replace(
    "'0~859행이 실제 업무 시각과 무관하게 일련번호처럼 하루씩 증가'",
    "'최초 860행에서 날짜가 행 순서대로 하루씩 증가하는 패턴 확인'"
)

# 3. Slide 4 badge & number center alignment
old_block = """for i, (v_badge, v_num, v_main, v_sub) in enumerate(ver_2):
    y = 3.02 + i * 1.24
    rect(s04, 0.85, y, 6.05, 1.10, c='white', line_c='rule', rad=True, rad_val=0.06)
    pill(s04, 1.02, y + 0.16, 1.28, 0.30, v_badge, fill='pale', text_c='accent', size=10.5, bold=True)
    txt(s04, 1.02, y + 0.50, 1.40, 0.48, v_num, size=25, color='accent', bold=True, font='head')
    txt(s04, 2.48, y + 0.14, 4.30, 0.42, v_main, size=13.0, color='ink', bold=True)
    txt(s04, 2.48, y + 0.56, 4.30, 0.42, v_sub, size=11.0, color='muted')"""

new_block = """for i, (v_badge, v_num, v_main, v_sub) in enumerate(ver_2):
    y = 3.02 + i * 1.24
    rect(s04, 0.85, y, 6.05, 1.10, c='white', line_c='rule', rad=True, rad_val=0.06)
    pill(s04, 1.02, y + 0.16, 1.32, 0.30, v_badge, fill='pale', text_c='accent', size=10.5, bold=True)
    txt(s04, 1.02, y + 0.50, 1.32, 0.46, v_num, size=25, color='accent', bold=True, font='head', align=PP_ALIGN.CENTER)
    txt(s04, 2.50, y + 0.16, 4.26, 0.38, v_main, size=13.0, color='ink', bold=True)
    txt(s04, 2.50, y + 0.56, 4.26, 0.38, v_sub, size=11.0, color='muted')"""

assert old_block in code, "old_block not found"
code = code.replace(old_block, new_block)
code = code.replace("'version': 8,", "'version': 9,")

(W / 'build_v009.py').write_text(code, encoding='utf-8')
print('build_v009.py created')
