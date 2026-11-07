from pathlib import Path

W = Path(__file__).resolve().parent
code = (W / 'build_v007.py').read_text(encoding='utf-8')
code = code.replace("Q = W / '검수_v007'", "Q = W / '검수_v008'")
code = code.replace("out = V / 'repair2_채용과제_분석결과_v007.pptx'", "out = V / 'repair2_채용과제_분석결과_v008.pptx'")
code = code.replace("prev_ppt = Presentation(V / 'repair2_채용과제_분석결과_v006.pptx')", "prev_ppt = Presentation(V / 'repair2_채용과제_분석결과_v007.pptx')")
code = code.replace(
    "'예정 시각과 지연 사유를 새로 기록하고\\n대기와 전체 완료 시간이 함께 줄었는지 검증한다'",
    "'예정 시각과 지연 사유를 추가로 기록하고\\n대기와 전체 완료 시간이 함께 줄었는지 검증한다'"
)
code = code.replace("'version': 7,", "'version': 8,")
(W / 'build_v008.py').write_text(code, encoding='utf-8')
print('build_v008.py created')
