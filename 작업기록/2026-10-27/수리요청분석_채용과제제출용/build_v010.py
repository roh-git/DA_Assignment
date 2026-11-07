from pathlib import Path
import json, re, shutil, hashlib
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

W = Path(__file__).resolve().parent
R = W.parents[2]
V = W / 'versions'
Q = W / '검수_v010'
S = W / 'sources'
for d in [V, Q, S]:
    d.mkdir(parents=True, exist_ok=True)

# Archive user-provided reference screenshots in sources/
for img_idx, tmp_path in enumerate([
    Path(r'C:\Users\rhd67\AppData\Local\Temp\orca-paste-1793061560012-fd3d3fba-60bb-4f8a-9f41-34d639b4484d.png'),
    Path(r'C:\Users\rhd67\AppData\Local\Temp\orca-paste-1793061592819-667d411d-8ecf-4daf-8a4b-d6a391adbee0.png')
], start=1):
    if tmp_path.exists():
        shutil.copy2(tmp_path, S / f'v006_user_feedback_{img_idx:02}.png')

out = V / 'repair2_채용과제_분석결과_v010.pptx'
pdf = out.with_suffix('.pdf')
if out.exists():
    out.unlink()
if pdf.exists():
    pdf.unlink()

# Preserve speaker notes from v005
prev_ppt = Presentation(V / 'repair2_채용과제_분석결과_v009.pptx')
prev_notes = [sl.notes_slide.notes_text_frame.text if sl.has_notes_slide else '' for sl in prev_ppt.slides]

C = {
    'bg': 'FFFFFF',
    'ink': '192327',
    'dark_card': '223036',
    'dark_line': '32454D',
    'muted': '5A6972',
    'accent': '087F75',
    'accent_hi': '12A598',
    'sage': '7FA9A2',
    'pale': 'E5F2EF',
    'pale_line': 'B8DAD3',
    'surface': 'F4F6F7',
    'rule': 'DCE3E6',
    'bar': '94A3AB',
    'white': 'FFFFFF'
}
F = {'head': 'Paperlogy 6 SemiBold', 'body': 'Pretendard'}

p = Presentation()
p.slide_width = Inches(13.333333)
p.slide_height = Inches(7.5)
p.core_properties.title = '수리 요청 프로세스 분석 및 개선 제안 — 채용 과제'
p.core_properties.author = 'Roh Hyundong'
p.core_properties.subject = '업무 흐름, 날짜 품질 검증, 기록상 대기 분석과 개선 제안'

def rgb(k):
    return RGBColor.from_string(C.get(k, k))

def strip_effects(sh):
    for ef in sh._element.xpath('.//a:effectRef'):
        ef.set('idx', '0')
    for sp in sh._element.xpath('./p:spPr'):
        for el in list(sp):
            if el.tag.endswith('}effectLst') or el.tag.endswith('}effectDag'):
                sp.remove(el)
        sp.append(OxmlElement('a:effectLst'))
    return sh

def rect(s, x, y, w, h, c='surface', line_c=None, rad=False, rad_val=0.04, line_w=0.75):
    shape_type = MSO_SHAPE.ROUNDED_RECTANGLE if rad else MSO_SHAPE.RECTANGLE
    sh = s.shapes.add_shape(shape_type, Inches(x), Inches(y), Inches(w), Inches(h))
    if c is None:
        sh.fill.background()
    else:
        sh.fill.solid()
        sh.fill.fore_color.rgb = rgb(c)
    if line_c:
        sh.line.color.rgb = rgb(line_c)
        sh.line.width = Pt(line_w)
    else:
        sh.line.fill.background()
    if rad:
        sh.adjustments[0] = rad_val
    return strip_effects(sh)

def txt(s, x, y, w, h, t, size=14, color='ink', bold=False, font='body', align=None, anchor=MSO_ANCHOR.MIDDLE, line_sp=1.15):
    sh = s.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = sh.text_frame
    tf.clear()
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    tf.vertical_anchor = anchor
    for i, ln in enumerate(t.split('\n')):
        par = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        par.space_before = Pt(0)
        par.space_after = Pt(0)
        par.line_spacing = line_sp
        if align is not None:
            par.alignment = align
        r = par.add_run()
        r.text = ln
        r.font.name = F[font]
        r.font.size = Pt(size)
        r.font.bold = bold
        r.font.color.rgb = rgb(color)
        rp = r._r.get_or_add_rPr()
        ea = OxmlElement('a:ea')
        ea.set('typeface', F[font])
        rp.append(ea)
    return strip_effects(sh)

def pill(s, x, y, w, h, t, fill='pale', text_c='accent', size=11, bold=True, rad_val=0.45, line_c=None):
    rect(s, x, y, w, h, c=fill, line_c=line_c, rad=True, rad_val=rad_val)
    return txt(s, x + 0.04, y, w - 0.08, h, t, size=size, color=text_c, bold=bold, align=PP_ALIGN.CENTER)

def line(s, x1, y1, x2, y2, color='rule', width=1.0, arrow=False):
    sh = s.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(x1), Inches(y1), Inches(x2), Inches(y2))
    sh.line.color.rgb = rgb(color)
    sh.line.width = Pt(width)
    if arrow:
        end = OxmlElement('a:tailEnd')
        end.set('type', 'triangle')
        sh.line._get_or_add_ln().append(end)
    return strip_effects(sh)

def dot(s, x, y, r=0.06, color='accent'):
    sh = s.shapes.add_shape(MSO_SHAPE.OVAL, Inches(x - r), Inches(y - r), Inches(2 * r), Inches(2 * r))
    sh.fill.solid()
    sh.fill.fore_color.rgb = rgb(color)
    sh.line.fill.background()
    return strip_effects(sh)

def ribbon(s, x1, x2, y1, y2, h, color):
    pts = []
    for j in range(31):
        u = j / 30
        f = 3 * u * u - 2 * u * u * u
        pts.append((Inches(x1 + (x2 - x1) * u), Inches(y1 + (y2 - y1) * f)))
    for j in range(30, -1, -1):
        u = j / 30
        f = 3 * u * u - 2 * u * u * u
        pts.append((Inches(x1 + (x2 - x1) * u), Inches(y1 + (y2 - y1) * f + h)))
    fb = s.shapes.build_freeform(int(pts[0][0]), int(pts[0][1]))
    fb.add_line_segments([(int(px), int(py)) for px, py in pts[1:]], close=True)
    sh = fb.convert_to_shape()
    sh.fill.solid()
    sh.fill.fore_color.rgb = rgb(color)
    sh.line.fill.background()
    return strip_effects(sh)

def bar(s, x, y, w, h, cats, series, maxval, fmt='0', legend=False, accents=None):
    data = CategoryChartData()
    data.categories = cats
    for label, values, color in series:
        data.add_series(label, values)
    ch = s.shapes.add_chart(XL_CHART_TYPE.BAR_CLUSTERED, Inches(x), Inches(y), Inches(w), Inches(h), data).chart
    ch.has_title = False
    ch.has_legend = legend
    if legend:
        ch.legend.position = XL_LEGEND_POSITION.BOTTOM
        ch.legend.include_in_layout = False
        ch.legend.font.name = F['body']
        ch.legend.font.size = Pt(12)
    plot = ch.plots[0]
    plot.gap_width = 65
    plot.has_data_labels = True
    dl = plot.data_labels
    dl.position = XL_LABEL_POSITION.OUTSIDE_END
    dl.font.name = F['body']
    dl.font.size = Pt(13)
    dl.font.color.rgb = rgb('ink')
    dl.number_format = fmt
    dl.number_format_is_linked = False
    ca = ch.category_axis
    ca.reverse_order = True
    ca.tick_labels.font.name = F['body']
    ca.tick_labels.font.size = Pt(13)
    ca.tick_labels.font.color.rgb = rgb('ink')
    ca.major_tick_mark = XL_TICK_MARK.NONE
    ca.minor_tick_mark = XL_TICK_MARK.NONE
    ca.format.line.fill.background()
    va = ch.value_axis
    va.minimum_scale = 0
    va.maximum_scale = maxval
    va.major_unit = maxval / 4
    va.tick_labels.font.name = F['body']
    va.tick_labels.font.size = Pt(11)
    va.tick_labels.font.color.rgb = rgb('muted')
    va.major_tick_mark = XL_TICK_MARK.NONE
    va.minor_tick_mark = XL_TICK_MARK.NONE
    va.has_major_gridlines = True
    va.major_gridlines.format.line.color.rgb = rgb('rule')
    va.major_gridlines.format.line.width = Pt(0.5)
    va.format.line.fill.background()
    va.tick_labels.number_format = '0'
    for k, (lab, vals, col) in enumerate(series):
        se = ch.series[k]
        se.format.fill.solid()
        se.format.fill.fore_color.rgb = rgb(col)
        se.format.line.fill.background()
        if accents and k in accents:
            for i, c_name in enumerate(accents[k]):
                se.points[i].format.fill.solid()
                se.points[i].format.fill.fore_color.rgb = rgb(c_name)
                se.points[i].format.line.fill.background()
    for e in ch._chartSpace.iter():
        if e.tag.endswith('}defRPr') or e.tag.endswith('}rPr'):
            ea = OxmlElement('a:ea')
            ea.set('typeface', F['body'])
            e.append(ea)
    return ch

def make_header_footer(idx, role, title, intent, source):
    s = p.slides.add_slide(p.slide_layouts[6])
    s.background.fill.solid()
    s.background.fill.fore_color.rgb = rgb('bg')
    pill(s, 0.65, 0.26, 2.25, 0.28, f'{idx:02} · {role}', fill='pale', text_c='accent', size=10.5, bold=True)
    pill(s, 11.78, 0.26, 0.90, 0.28, f'Page {idx:02}', fill='surface', text_c='muted', size=10, bold=True)
    if '\n' in title:
        txt(s, 0.65, 0.62, 12.03, 0.88, title, size=25.5, color='ink', bold=True, font='head', line_sp=1.12)
        txt(s, 0.65, 1.54, 12.03, 0.44, intent, size=14, color='muted', line_sp=1.12)
    else:
        txt(s, 0.65, 0.68, 12.03, 0.56, title, size=25.5, color='ink', bold=True, font='head', line_sp=1.12)
        txt(s, 0.65, 1.34, 12.03, 0.44, intent, size=14, color='muted', line_sp=1.12)
    line(s, 0.65, 2.06, 12.68, 2.06, color='rule', width=0.9)
    line(s, 0.65, 7.10, 12.68, 7.10, color='rule', width=0.6)
    txt(s, 0.65, 7.15, 11.2, 0.22, 'repair2 분석 / ' + source, size=9.5, color='muted')
    txt(s, 12.10, 7.14, 0.58, 0.24, f'{idx:02}', size=10, color='muted', align=PP_ALIGN.RIGHT)
    return s

def add_bottom_split(s, insight, caution, left_badge='해석', right_badge='해석 시 주의', left_w=7.35):
    right_x = 0.65 + left_w + 0.18
    right_w = 12.03 - left_w - 0.18
    rect(s, 0.65, 6.30, left_w, 0.66, c='pale', line_c='pale_line', rad=True, rad_val=0.10)
    pill(s, 0.82, 6.47, 0.86, 0.32, left_badge, fill='accent', text_c='white', size=11, bold=True, rad_val=0.35)
    txt(s, 1.80, 6.34, left_w - 1.28, 0.58, insight, size=13.5, color='ink', bold=True, line_sp=1.14)

    rect(s, right_x, 6.30, right_w, 0.66, c='surface', line_c='rule', rad=True, rad_val=0.10)
    pill(s, right_x + 0.16, 6.47, 0.98, 0.32, right_badge, fill='ink', text_c='white', size=10, bold=True, rad_val=0.35)
    txt(s, right_x + 1.24, 6.34, right_w - 1.36, 0.58, caution, size=11.5, color='muted', line_sp=1.14)

def add_bottom_single(s, badge_text, main_text, right_note='', tone='pale', note_w_override=None):
    bg_c = 'pale' if tone == 'pale' else 'surface'
    bd_c = 'pale_line' if tone == 'pale' else 'rule'
    b_fill = 'accent' if tone == 'pale' else 'ink'
    rect(s, 0.65, 6.34, 12.03, 0.60, c=bg_c, line_c=bd_c, rad=True, rad_val=0.12)
    pill(s, 0.82, 6.48, 1.15, 0.32, badge_text, fill=b_fill, text_c='white', size=10.5, bold=True, rad_val=0.35)
    if right_note:
        note_w = note_w_override if note_w_override else (2.75 if len(right_note) <= 24 else 3.90)
        main_w = 10.40 - note_w - 0.12
        txt(s, 2.10, 6.36, main_w, 0.56, main_text, size=12.4, color='ink', bold=True)
        txt(s, 12.50 - note_w, 6.36, note_w, 0.56, right_note, size=10.0, color='muted', align=PP_ALIGN.RIGHT, line_sp=1.14)
    else:
        txt(s, 2.10, 6.38, 10.40, 0.52, main_text, size=13.2, color='ink', bold=True)

# ==============================================================================
# SLIDE 01: 표지
# ==============================================================================
s01 = p.slides.add_slide(p.slide_layouts[6])
s01.background.fill.solid()
s01.background.fill.fore_color.rgb = rgb('ink')
rect(s01, 0, 0, 13.333333, 0.08, c='accent_hi')
pill(s01, 0.75, 0.48, 1.45, 0.32, '분석 개요', fill='dark_card', text_c='accent_hi', size=11, bold=True, line_c='dark_line')
txt(s01, 9.20, 0.50, 3.38, 0.28, 'DATA ANALYSIS', size=10.5, color='bar', bold=True, align=PP_ALIGN.RIGHT)
line(s01, 0.75, 0.96, 12.58, 0.96, color='dark_line', width=0.8)

txt(s01, 0.75, 1.38, 11.8, 1.65, '수리 요청 프로세스\n분석 및 개선 제안', size=40, color='white', font='head', line_sp=1.12)
txt(s01, 0.75, 3.22, 11.5, 0.78, '네덜란드 임대 주택 기관의 수리 요청 데이터\n업무 흐름과 기록상 대기를 분석해 개선 검토 대상을 제안', size=18, color='pale', line_sp=1.2)

cover_cards = [
    (0.75, 'SCOPE 01 · 전체 기록', '1,000', '전체 수리 요청 / 건', '13,262개 업무 이벤트 보존'),
    (4.76, 'SCOPE 02 · 일반 경로', '927', '일반 경로 / 건', '잘못 접수된 예외 73건 분리'),
    (8.77, 'SCOPE 03 · 시간 표본', '339', '시간 분석 대상 / 건', '모든 시각 존재 · 1970년 일반 경로')
]
for x, tag, num, label, sub in cover_cards:
    rect(s01, x, 4.35, 3.80, 2.05, c='dark_card', line_c='dark_line', rad=True, rad_val=0.06)
    pill(s01, x + 0.24, 4.55, 1.85, 0.28, tag, fill='ink', text_c='accent_hi', size=10, bold=True)
    txt(s01, x + 0.24, 4.92, 3.32, 0.68, num, size=36, color='white', font='head')
    txt(s01, x + 0.24, 5.64, 3.32, 0.30, label, size=14.5, color='white', bold=True)
    txt(s01, x + 0.24, 5.96, 3.32, 0.26, sub, size=11.5, color='bar')

line(s01, 0.75, 6.76, 12.58, 6.76, color='dark_line', width=0.7)
txt(s01, 0.75, 6.90, 11.8, 0.28, '노현동 / 데이터 분석 / 2026.10.26', size=11.5, color='bar')

# ==============================================================================
# SLIDE 02: 분석 목적과 검토 질문 (Q3·Q4 문구 명확화)
# ==============================================================================
s02 = make_header_footer(
    2, '분석 목적과 검토 질문',
    '수리 요청의 지연과 반복이\n어디에서 나타나는지 확인하고자 했다',
    '업무 수행 시간과 업무 사이 대기를 구분하고 업무 구성·수리 유형·접수 방식과의 관계를 살펴봤다',
    '분석 질문 재구성 / 노트북 7–15'
)
q_cards = [
    ('Q1', '시간 집중 구간', '시간이 길어지는 곳은\n어디인가', '개별 업무 시간과\n업무 사이 대기를 나눠 비교'),
    ('Q2', '반복 수행 여부', '동일한 수리 업무가\n반복되는가', '업무 종류, 수행 횟수,\n시작·완료 순서를 확인'),
    ('Q3', '유형·구성 비교', '수리 유형과 수행 업무에 따라\n처리 시간이 다른가', 'RepairType별 전체 시간과\n같은 업무 조건을 비교'),
    ('Q4', '채널·담당자 비교', '접수 채널과 담당자 기록에 따라\n처리 시간이 다른가', '접수 채널과 시작·완료\n담당자별 시간 분포를 비교')
]
for i, (code, tag, q_title, q_desc) in enumerate(q_cards):
    x = 0.65 + i * 3.06
    rect(s02, x, 2.24, 2.85, 3.04, c='surface', line_c='rule', rad=True, rad_val=0.05)
    rect(s02, x + 0.14, 2.40, 2.57, 0.46, c='pale', rad=True, rad_val=0.12)
    pill(s02, x + 0.24, 2.49, 0.48, 0.28, code, fill='accent', text_c='white', size=10.5, bold=True)
    txt(s02, x + 0.82, 2.47, 1.80, 0.32, tag, size=12, color='accent', bold=True)
    txt(s02, x + 0.12, 3.04, 2.61, 0.80, q_title, size=14.0, color='ink', bold=True, align=PP_ALIGN.CENTER, line_sp=1.18)
    line(s02, x + 0.24, 4.00, x + 2.61, 4.00, color='rule', width=0.8)
    txt(s02, x + 0.20, 4.16, 2.45, 0.90, q_desc, size=13.5, color='muted', align=PP_ALIGN.CENTER, line_sp=1.2)

rect(s02, 0.65, 5.44, 12.03, 0.72, c='ink', rad=True, rad_val=0.10)
pill(s02, 0.85, 5.62, 1.35, 0.36, '분석 진행 순서', fill='accent_hi', text_c='ink', size=11, bold=True)
txt(s02, 2.42, 5.56, 10.0, 0.48, '데이터 품질 확인   →   업무 구성·순서 확인   →   시간 비교   →   개선과 검증 지표 제안', size=15.5, color='white', bold=True)

add_bottom_single(
    s02, '질문 구성',
    '분석 질문은 수행 과정의 판단을 정리한 것이며 E 유형과 특정 채널 비교는 탐색 후 추가',
    tone='surface'
)

# ==============================================================================
# SLIDE 03: 데이터 개요와 전처리 (1,000건 선별 기준 문구 수정)
# ==============================================================================
s03 = make_header_footer(
    3, '데이터 개요와 전처리',
    '1,000건의 수리 요청과\n13,262개 업무 기록을 분석했다',
    '요청(caseID)과 업무(taskID)를 기준으로 기록을 정리하고 시간 계산에 필요한 표본을 선별',
    '1–5 / 8–9'
)
rect(s03, 0.65, 2.24, 3.85, 3.90, c='surface', line_c='rule', rad=True, rad_val=0.04)
rect(s03, 0.83, 2.40, 3.49, 0.46, c='pale', rad=True, rad_val=0.12)
txt(s03, 0.95, 2.46, 3.25, 0.34, '원본 데이터 구조와 결측 확인', size=13, color='accent', bold=True, align=PP_ALIGN.CENTER)

rect(s03, 0.83, 3.02, 3.49, 0.98, c='white', line_c='rule', rad=True, rad_val=0.08)
txt(s03, 1.02, 3.14, 1.30, 0.72, '13개', size=28, color='accent', bold=True, font='head')
txt(s03, 2.32, 3.20, 1.85, 0.62, '업무 종류\n접수·조사·수리·종결', size=13.5, color='ink', bold=True)

rect(s03, 0.83, 4.14, 3.49, 0.98, c='white', line_c='rule', rad=True, rad_val=0.08)
txt(s03, 1.02, 4.26, 1.30, 0.72, '16종', size=28, color='accent', bold=True, font='head')
txt(s03, 2.32, 4.30, 1.90, 0.64, '업무 수행 주체 /\n시스템 포함', size=13.5, color='ink', bold=True)

txt(s03, 0.83, 5.26, 3.52, 0.74, 'Date·Time 결합 후 시작·완료 시각 연결\n날짜·시각 결측 13행 / 수리 속성 결측은 예외 경로에서 확인', size=10.5, color='muted', line_sp=1.18)

rect(s03, 4.72, 2.24, 7.96, 0.46, c='ink', rad=True, rad_val=0.10)
txt(s03, 4.96, 2.32, 2.20, 0.30, '분석 대상', size=13.5, color='white', bold=True)
txt(s03, 7.25, 2.32, 1.45, 0.30, '규모', size=13.5, color='white', bold=True)
txt(s03, 8.85, 2.32, 3.65, 0.30, '선별 기준', size=13.5, color='white', bold=True)

sample_rows = [
    ('업무 구성·경로', '1,000건', '원본 이벤트 보존\n날짜 오류와 관계없이 수행된 업무와 경로 확인', 'surface', 'rule', 'ink'),
    ('개별 업무 시간', '3,151쌍', '시작·완료가 각 1회인 3,631쌍에서 선별\n모두 1970년 기록이며 완료 시각 ≥ 시작 시각', 'surface', 'rule', 'ink'),
    ('일반 경로 전체 시간', '339건', '모든 시각이 존재하고 1970년인 402건\n이 중 예외 경로 63건 제외 (일반 경로 927건의 36.6%)', 'pale', 'accent', 'accent')
]
for i, (a, b, c_text, bg_c, bd_c, num_c) in enumerate(sample_rows):
    y = 2.84 + i * 1.12
    rect(s03, 4.72, y, 7.96, 0.98, c=bg_c, line_c=bd_c, rad=True, rad_val=0.06, line_w=1.1 if i == 2 else 0.75)
    pill(s03, 4.92, y + 0.29, 2.05, 0.40, a, fill='white' if i < 2 else 'accent', text_c='ink' if i < 2 else 'white', size=13, bold=True, line_c='rule' if i < 2 else None)
    txt(s03, 7.15, y + 0.19, 1.45, 0.60, b, size=22, color=num_c, bold=True, font='head')
    txt(s03, 8.68, y + 0.12, 3.88, 0.74, c_text, size=12.5, color='ink', line_sp=1.16)

add_bottom_single(
    s03, '선별 원칙',
    '업무 흐름은 전체 1,000건으로 확인하고 시간은 조건을 충족한 표본에서만 탐색적으로 계산',
    right_note='1970년 실제 연도 미확인 · 339건 전체 대표성 보장 안 됨',
    tone='pale'
)

# ==============================================================================
# SLIDE 04: 날짜 검증과 처리 결정 — [1순위: 상단 제목·부제목 변환 + 임의 보정 보류 판단 근거 명확화]
# ==============================================================================
s04 = make_header_footer(
    4, '날짜 검증과 처리 결정',
    '날짜 기록의 정합성을 검증하고 분석 대상을 구분했다',
    '행 순서에 따른 날짜 증가와 업무 시작·완료 시간 역전을 확인해 날짜 기록의 신뢰성을 검토',
    '9-1–9-5'
)
# Left: 2 Core Verification Evidences (860건 & 226쌍)
rect(s04, 0.65, 2.24, 6.45, 3.90, c='surface', line_c='rule', rad=True, rad_val=0.04)
rect(s04, 0.85, 2.40, 6.05, 0.46, c='pale', rad=True, rad_val=0.12)
txt(s04, 1.02, 2.47, 5.70, 0.32, '확인된 날짜 이상 (핵심 검증 근거 2가지)', size=14.5, color='accent', bold=True)

ver_2 = [
    ('행 순서 검사', '860건', '2020~2022년 날짜가 CSV 행 순서대로 하루씩 증가', '최초 860행에서 날짜가 행 순서대로 하루씩 증가하는 패턴 확인'),
    ('시작·완료 검사', '226쌍', '시작보다 완료 시각이 먼저 기록된 업무가 존재', '비교 가능한 3,627쌍 중 시작·완료 연도 불일치로 시간 역전 발생')
]
for i, (v_badge, v_num, v_main, v_sub) in enumerate(ver_2):
    y = 3.02 + i * 1.24
    rect(s04, 0.85, y, 6.05, 1.10, c='white', line_c='rule', rad=True, rad_val=0.06)
    pill(s04, 1.02, y + 0.16, 1.32, 0.30, v_badge, fill='pale', text_c='accent', size=10.5, bold=True)
    txt(s04, 1.02, y + 0.50, 1.32, 0.46, v_num, size=25, color='accent', bold=True, font='head', align=PP_ALIGN.CENTER)
    txt(s04, 2.50, y + 0.16, 4.26, 0.38, v_main, size=13.0, color='ink', bold=True)
    txt(s04, 2.50, y + 0.56, 4.26, 0.38, v_sub, size=11.0, color='muted')

txt(s04, 0.85, 5.58, 6.05, 0.40, '연도별 이벤트: 1970년 12,383행 / 2020–2022년 860행 / 1905년 6행 / 날짜·시각 결측 13행', size=11.5, color='muted')

# Right: Judgment (임의 보정 보류 판단 근거) + Action (분석 대상 구분)
rect(s04, 7.32, 2.24, 5.36, 3.90, c='surface', line_c='rule', rad=True, rad_val=0.04)

# Top Box: 판단
rect(s04, 7.52, 2.40, 4.96, 1.66, c='ink', rad=True, rad_val=0.06)
pill(s04, 7.74, 2.58, 0.90, 0.30, '판단 근거', fill='accent_hi', text_c='ink', size=10.5, bold=True)
txt(s04, 8.78, 2.55, 3.50, 0.36, '연도 일괄 치환·행 삭제 보류', size=16.5, color='white', bold=True, font='head')
txt(s04, 7.74, 3.04, 4.52, 0.84, '실제 발생 날짜와 오류 수정 규칙을 확인할 수 없어\n임의로 연도를 치환하면 경과 시간이 왜곡될 수 있음', size=13.5, color='pale', bold=True, line_sp=1.22)

# Bottom Box: 처리
rect(s04, 7.52, 4.20, 4.96, 1.76, c='pale', line_c='accent', rad=True, rad_val=0.06, line_w=1.1)
pill(s04, 7.74, 4.36, 0.90, 0.30, '처리 결정', fill='accent', text_c='white', size=10.5, bold=True)
txt(s04, 8.78, 4.34, 3.56, 0.36, '원본은 유지하고 분석 목적에 따라 데이터 분리', size=13.5, color='ink', bold=True)

for j, (s_tag, s_val) in enumerate([('업무 흐름', '1,000건'), ('개별 업무 시간', '3,151쌍'), ('일반 경로 시간', '339건')]):
    sx = 7.72 + j * 1.56
    rect(s04, sx, 4.84, 1.44, 0.94, c='white', line_c='pale_line', rad=True, rad_val=0.08)
    txt(s04, sx + 0.06, 4.94, 1.32, 0.30, s_tag, size=11.5, color='muted', bold=True, align=PP_ALIGN.CENTER)
    txt(s04, sx + 0.06, 5.28, 1.32, 0.40, s_val, size=17.5, color='accent', bold=True, font='head', align=PP_ALIGN.CENTER)

add_bottom_single(
    s04, '해석 시 주의',
    '1970년 기록도 실제 날짜의 정확성이 검증된 것은 아님 (동일 연도 내 기록상 경과 시간만 탐색적으로 비교)',
    tone='surface'
)

# ==============================================================================
# SLIDE 05: 업무 흐름 분석 (하단 표현 명확화)
# ==============================================================================
s05 = make_header_footer(
    5, '업무 흐름 분석',
    '수리 요청은 일반·예외 경로로 나뉘고\n수리 업무도 여러 형태로 기록됐다',
    '시간과 반복 업무를 분석하기 전에 케이스별 수행 업무와 종료 경로를 확인했다',
    '8-2 / 10-3–10-5'
)
rect(s05, 0.65, 2.24, 3.30, 3.90, c='surface', line_c='rule', rad=True, rad_val=0.04)
rect(s05, 0.83, 2.38, 2.94, 0.38, c='pale', rad=True, rad_val=0.12)
txt(s05, 0.93, 2.42, 2.74, 0.30, '전체 요청 1,000건', size=14, color='accent', bold=True, align=PP_ALIGN.CENTER)

ribbon(s05, 1.05, 2.45, 2.95, 2.92, 1.15 * 0.927, 'pale_line')
ribbon(s05, 1.05, 2.45, 2.95 + 1.15 * 0.927, 4.88, 1.15 * 0.073, 'bar')
rect(s05, 0.93, 2.95, 0.12, 1.15, c='ink')
rect(s05, 2.45, 2.92, 0.12, 1.15 * 0.927, c='accent')
rect(s05, 2.45, 4.88, 0.12, 1.15 * 0.073, c='bar')

txt(s05, 2.66, 2.98, 1.15, 0.62, '일반\n927건', size=15.5, color='accent', bold=True)
txt(s05, 2.66, 4.74, 1.15, 0.44, '예외 73건', size=13.5, color='ink', bold=True)
txt(s05, 0.85, 4.22, 1.65, 0.52, 'FirstContact\n최초 접수', size=12.5, color='ink', bold=True)

line(s05, 2.51, 5.02, 2.51, 5.28, color='muted', width=1.4, arrow=True)
rect(s05, 0.83, 5.30, 2.94, 0.66, c='white', line_c='rule', rad=True, rad_val=0.08)
txt(s05, 0.90, 5.36, 2.80, 0.54, 'InformClientWrongPlace\n잘못 접수된 요청 안내', size=12, color='ink', bold=True, align=PP_ALIGN.CENTER)

rect(s05, 4.12, 2.24, 8.56, 3.90, c='surface', line_c='rule', rad=True, rad_val=0.04)
line(s05, 3.72, 3.04, 4.32, 3.04, color='accent', width=1.6, arrow=True)

def flow_card(s, x, y, w, h, en, ko, hi=False):
    rect(s, x, y, w, h, c='pale' if hi else 'white', line_c='accent' if hi else 'rule', rad=True, rad_val=0.08, line_w=1.1 if hi else 0.75)
    txt(s, x + 0.08, y + 0.06, w - 0.16, h * 0.46, en, size=12.5, color='accent' if hi else 'ink', bold=True, align=PP_ALIGN.CENTER)
    txt(s, x + 0.08, y + h * 0.48, w - 0.16, h * 0.44, ko, size=11.5, color='ink' if hi else 'muted', align=PP_ALIGN.CENTER)

flow_card(s05, 4.35, 2.66, 2.25, 0.74, 'MakeTicket', '수리 요청 등록')
line(s05, 6.60, 3.03, 7.12, 3.03, color='muted', width=1.4, arrow=True)
flow_card(s05, 7.12, 2.66, 2.45, 0.74, 'ArrangeSurvey', '현장 조사 일정 조율')
line(s05, 9.57, 3.03, 10.10, 3.03, color='muted', width=1.4, arrow=True)
flow_card(s05, 10.10, 2.66, 2.35, 0.74, 'Survey', '현장 조사')

pill(s05, 6.90, 3.48, 3.30, 0.25, 'InformClientSurvey: 현장 조사 안내', fill='white', text_c='muted', size=10, bold=False, line_c='rule')

line(s05, 11.27, 3.40, 11.27, 3.88, color='accent', width=1.3)
line(s05, 5.47, 3.88, 11.27, 3.88, color='accent', width=1.3)
for x, en, ko in [(4.35, 'InternRepair', '내부 수리'), (7.12, 'ImmediateRepair', '즉시 수리'), (10.10, 'ExternRepair', '외부 수리')]:
    w_box = 2.35 if x > 10 else (2.45 if x > 7 else 2.25)
    cx = x + w_box / 2
    line(s05, cx, 3.88, cx, 4.06, color='accent', width=1.3, arrow=True)
    flow_card(s05, x, 4.06, w_box, 0.70, en, ko, hi=True)
    line(s05, cx, 4.76, cx, 4.96, color='accent', width=1.3)
line(s05, 5.47, 4.96, 11.27, 4.96, color='accent', width=1.3)
pill(s05, 6.75, 4.83, 3.20, 0.25, '단일 또는 복합 가능 · 반복 수행 관찰', fill='pale', text_c='accent', size=10, bold=True, line_c='accent')
line(s05, 5.47, 4.96, 5.47, 5.26, color='accent', width=1.3, arrow=True)

flow_card(s05, 4.35, 5.26, 2.25, 0.70, 'RepairReady', '수리 완료 처리')
line(s05, 6.60, 5.61, 6.95, 5.61, color='muted', width=1.4, arrow=True)
rect(s05, 6.95, 5.26, 2.80, 0.70, c='white', line_c='rule', rad=True, rad_val=0.08)
txt(s05, 7.01, 5.30, 2.68, 0.62, '고객 안내·정산 전달\nReadyInformClient /\nSendTicketToFinAdmin', size=10.5, color='ink', bold=True, align=PP_ALIGN.CENTER, line_sp=1.10)
line(s05, 9.75, 5.61, 10.10, 5.61, color='muted', width=1.4, arrow=True)
flow_card(s05, 10.10, 5.26, 2.35, 0.70, 'TicketReady', '요청 종결')

add_bottom_single(
    s05, '구조 해석',
    '한 요청에서 여러 수리 업무나 동일 업무의 반복이 관찰됨',
    right_note='업무 흐름 구조 요약 · 종결 관련 두 업무의 순서는 고정 아님',
    tone='pale'
)

# ==============================================================================
# SLIDE 06: 반복 작업 분석 (부제목 간결화)
# ==============================================================================
s06 = make_header_footer(
    6, '반복 작업 분석',
    '일부 요청에서 여러 수리 업무와\n같은 업무의 반복 기록이 관찰됐다',
    '반복 작업 여부를 확인하기 위해 업무별 수행 횟수와 순서를 분석',
    '10-3 / 10-5 / 12-5'
)
rect(s06, 0.65, 2.24, 6.85, 2.46, c='surface', line_c='rule', rad=True, rad_val=0.04)
rect(s06, 0.85, 2.40, 6.45, 0.42, c='pale', rad=True, rad_val=0.12)
txt(s06, 1.00, 2.45, 6.15, 0.32, '일반 경로 927건의 수리 업무 구성', size=14.5, color='accent', bold=True)

bar_w = 6.45
w_single = bar_w * (836 / 927)
w_multi = bar_w - w_single
rect(s06, 0.85, 2.98, w_single, 0.48, c='bar')
rect(s06, 0.85 + w_single, 2.98, w_multi, 0.48, c='accent')

rect(s06, 0.85, 3.62, 3.12, 0.90, c='white', line_c='rule', rad=True, rad_val=0.08)
pill(s06, 1.00, 3.76, 0.90, 0.28, '단일 수리', fill='surface', text_c='muted', size=10.5, bold=True)
txt(s06, 2.02, 3.70, 1.80, 0.40, '836건', size=20, color='ink', bold=True, font='head')
txt(s06, 2.02, 4.12, 1.80, 0.28, '비중 90.2%', size=12.5, color='muted', bold=True)

rect(s06, 4.18, 3.62, 3.12, 0.90, c='pale', line_c='accent', rad=True, rad_val=0.08)
pill(s06, 4.33, 3.76, 0.90, 0.28, '복합 수리', fill='accent', text_c='white', size=10.5, bold=True)
txt(s06, 5.35, 3.70, 1.80, 0.40, '91건', size=20, color='accent', bold=True, font='head')
txt(s06, 5.35, 4.12, 1.80, 0.28, '비중 9.8%', size=12.5, color='accent', bold=True)

rect(s06, 7.72, 2.24, 4.96, 2.46, c='surface', line_c='rule', rad=True, rad_val=0.04)
rect(s06, 7.92, 2.40, 4.56, 1.02, c='pale', line_c='pale_line', rad=True, rad_val=0.08)
txt(s06, 8.12, 2.50, 2.50, 0.34, 'InternRepair 반복 기록', size=14.5, color='accent', bold=True)
txt(s06, 8.12, 2.86, 2.20, 0.42, '동일 수리 업무 2회 이상', size=12, color='muted')
txt(s06, 10.45, 2.52, 1.85, 0.78, '38건', size=32, color='accent', bold=True, font='head', align=PP_ALIGN.RIGHT)

rect(s06, 7.92, 3.56, 4.56, 0.96, c='white', line_c='rule', rad=True, rad_val=0.08)
pill(s06, 8.08, 3.86, 0.95, 0.32, '시간순 점검', fill='ink', text_c='white', size=10.5, bold=True)
txt(s06, 9.16, 3.64, 3.20, 0.80, 'InternRepair 시각이 모두 1970년인\n29건을 시간순으로 확인\n28건에서 완료 후 다시 시작', size=12.5, color='ink', line_sp=1.15)

rect(s06, 0.65, 4.86, 12.03, 1.28, c='surface', line_c='rule', rad=True, rad_val=0.05)
pill(s06, 0.85, 5.08, 2.15, 0.36, '실제 기록 순서 예시', fill='ink', text_c='white', size=11.5, bold=True)
txt(s06, 0.90, 5.52, 2.05, 0.38, 'caseID 124', size=14, color='accent', bold=True, align=PP_ALIGN.CENTER)

line(s06, 3.85, 5.48, 10.15, 5.48, color='accent', width=2.0)
for tx, t_str, ev_str, is_st in [(3.85, '11:41', 'start', True), (5.95, '15:21', 'complete', False), (8.05, '19:45', 'start', True), (10.15, '23:03', 'complete', False)]:
    dot(s06, tx, 5.48, r=0.08, color='accent' if is_st else 'ink')
    pill(s06, tx - 0.55, 5.00, 1.10, 0.32, t_str, fill='pale' if is_st else 'white', text_c='accent' if is_st else 'ink', size=12, bold=True, line_c='accent' if is_st else 'rule')
    txt(s06, tx - 0.65, 5.64, 1.30, 0.32, ev_str, size=12.5, color='ink', bold=True, align=PP_ALIGN.CENTER)

pill(s06, 11.05, 5.20, 1.42, 0.56, '동일 일자\n간격 비례 아님', fill='white', text_c='muted', size=10.5, bold=False, rad_val=0.20, line_c='rule')

add_bottom_single(
    s06, '해석 시 주의',
    '복합 수리 91건과 InternRepair 반복 38건은 두 집단이 겹칠 수 있어 합산 불가 · 반복 원인·재작업 여부는 미확인',
    tone='surface'
)

# ==============================================================================
# SLIDE 07: 개별 업무 시간과 전체 시간 비교 — [2순위: 하단 주석 문구 변경 및 자연스러운 2줄 배치]
# ==============================================================================
s07 = make_header_footer(
    7, '개별 업무 시간과 전체 시간 비교',
    '조율 업무가 1시간을 넘는 요청에서도\n전체 처리 시간 중앙값 차이는 작았다',
    'ArrangeSurvey(현장 조사 일정 조율)의 일부 긴 처리 기록을 확인한 뒤 케이스 전체 시간과의 관계를 비교했다',
    '12-1–12-2'
)
rect(s07, 0.65, 2.24, 5.86, 3.90, c='surface', line_c='rule', rad=True, rad_val=0.04)
rect(s07, 0.85, 2.40, 5.46, 0.52, c='pale', rad=True, rad_val=0.12)
txt(s07, 1.02, 2.48, 5.12, 0.36, '업무 시작→완료 / 809 업무쌍 / 단위: 분', size=15.5, color='accent', bold=True, align=PP_ALIGN.CENTER)

rect(s07, 0.85, 3.10, 2.64, 1.60, c='white', line_c='rule', rad=True, rad_val=0.06)
pill(s07, 1.05, 3.28, 1.05, 0.30, '중앙값', fill='surface', text_c='muted', size=11.5, bold=True)
txt(s07, 1.05, 3.74, 2.24, 0.74, '4분', size=32, color='accent', bold=True, font='head')

rect(s07, 3.67, 3.10, 2.64, 1.60, c='white', line_c='accent', rad=True, rad_val=0.06)
pill(s07, 3.87, 3.28, 1.05, 0.30, 'P90', fill='pale', text_c='accent', size=11.5, bold=True)
txt(s07, 3.87, 3.74, 2.24, 0.74, '404.8분', size=30, color='accent', bold=True, font='head')

rect(s07, 0.85, 4.88, 5.46, 1.06, c='white', line_c='rule', rad=True, rad_val=0.06)
pill(s07, 1.02, 5.24, 1.05, 0.32, '분포 기준', fill='ink', text_c='white', size=11, bold=True)
txt(s07, 2.22, 5.00, 3.92, 0.80, '1시간 초과 194건(24.0%)\nP90: 기록의 하위 90%와 상위 10%를 나누는 경계', size=13, color='ink', line_sp=1.2)

rect(s07, 6.82, 2.24, 5.86, 3.90, c='surface', line_c='rule', rad=True, rad_val=0.04)
rect(s07, 7.02, 2.40, 5.46, 0.52, c='pale', rad=True, rad_val=0.12)
txt(s07, 7.18, 2.48, 5.12, 0.36, '전체 경과 시간 중앙값 / 339 케이스 / 단위: 시간', size=15.5, color='accent', bold=True, align=PP_ALIGN.CENTER)

for i, (g_lab, val, bar_c, bd_c) in enumerate([
    ('조율 1시간 초과 / 75건', 78.2, 'accent', 'accent'),
    ('조율 1시간 이하 / 264건', 77.2, 'bar', 'rule')
]):
    y = 3.12 + i * 1.42
    rect(s07, 7.02, y, 5.46, 1.24, c='white', line_c=bd_c, rad=True, rad_val=0.06)
    txt(s07, 7.24, y + 0.16, 3.40, 0.34, g_lab, size=14.5, color='ink', bold=True)
    rect(s07, 7.24, y + 0.64, 3.55, 0.32, c='surface', rad=True, rad_val=0.25)
    rect(s07, 7.24, y + 0.64, 3.55 * (val / 90.0), 0.32, c=bar_c, rad=True, rad_val=0.25)
    txt(s07, 10.95, y + 0.38, 1.35, 0.62, f'{val}h', size=24, color='accent' if i == 0 else 'ink', bold=True, font='head', align=PP_ALIGN.RIGHT)

add_bottom_single(
    s07, '비교 판단',
    '개별 업무 시간만으로는 전체 지연을 설명하기 어려워 업무 사이 대기를 이어서 확인',
    right_note='업무 809쌍과 요청 339건은 분석 대상이 다르며\n중앙값 비교만으로 관계를 단정할 수 없음',
    tone='pale',
    note_w_override=4.20
)

# ==============================================================================
# SLIDE 08: 전체 처리 시간 분석
# ==============================================================================
s08 = make_header_footer(
    8, '전체 처리 시간 분석',
    '조사 일정 조율 완료 후 현장 조사 시작까지\n기록상 대기가 가장 큰 비중을 차지했다',
    '업무를 수행하는 시간과 다음 업무를 기다리는 시간을 나누어 우선 살펴볼 구간을 찾았다',
    '12-2-1 / 12-4-1'
)
rect(s08, 0.65, 2.24, 12.03, 2.36, c='surface', line_c='rule', rad=True, rad_val=0.04)
pill(s08, 0.85, 2.38, 6.45, 0.34, '시간 분석 339건 / 기록상 경과 시간 합계 비중 / 전체 처리 시간 중앙값 77.3시간', fill='white', text_c='ink', size=12.5, bold=True, line_c='rule')

events = [
    ('FirstContact', 'complete'),
    ('ArrangeSurvey', 'complete'),
    ('Survey', 'start'),
    ('RepairReady', 'complete'),
    ('TicketReady', 'complete')
]
for j, (ev_a, ev_b) in enumerate(events):
    ex = 0.85 + j * 2.42
    hi = j in (1, 2)
    rect(s08, ex, 2.88, 1.96, 0.76, c='pale' if hi else 'white', line_c='accent' if hi else 'rule', rad=True, rad_val=0.08, line_w=1.2 if hi else 0.75)
    txt(s08, ex + 0.06, 2.94, 1.84, 0.64, f'{ev_a}\n{ev_b}', size=13, color='accent' if hi else 'ink', bold=hi, align=PP_ALIGN.CENTER)

intervals = [('1.9%', '접수→조율 완료', False), ('70.8%', '조율 완료→조사 시작', True), ('22.5%', '조사 시작→수리 완료', False), ('4.9%', '수리 완료→종결', False)]
for j, (pct_str, int_lab, hi) in enumerate(intervals):
    ex = 0.85 + j * 2.42
    line(s08, ex + 1.96, 3.26, ex + 2.42, 3.26, color='accent' if hi else 'bar', width=2.4 if hi else 1.4, arrow=True)
    cx = ex + 2.19
    rect(s08, cx - 0.95, 3.76, 1.90, 0.70, c='accent' if hi else 'white', line_c='accent' if hi else 'rule', rad=True, rad_val=0.12)
    txt(s08, cx - 0.90, 3.79, 1.80, 0.36, pct_str, size=18 if hi else 16, color='white' if hi else 'ink', bold=True, font='head', align=PP_ALIGN.CENTER)
    txt(s08, cx - 0.90, 4.15, 1.80, 0.24, int_lab, size=10.5, color='pale' if hi else 'muted', bold=hi, align=PP_ALIGN.CENTER)

rect(s08, 0.65, 4.74, 12.03, 1.42, c='ink', rad=True, rad_val=0.06)
rect(s08, 0.92, 4.96, 11.49, 0.40, c='bar', rad=True, rad_val=0.25)
rect(s08, 0.92, 4.96, 11.49 * 0.708, 0.40, c='accent_hi', rad=True, rad_val=0.25)
txt(s08, 0.92, 5.46, 7.80, 0.30, 'ArrangeSurvey complete → Survey start  70.8%', size=14.5, color='white', bold=True)
txt(s08, 8.90, 5.46, 3.51, 0.30, '나머지 세 구간', size=13.5, color='pale', bold=True, align=PP_ALIGN.RIGHT)
txt(s08, 0.92, 5.78, 11.49, 0.26, '22.5% 구간에는 현장 조사(Survey)와 수리 관련 경과 시간이 함께 포함됨', size=12, color='bar')

add_bottom_single(
    s08, '해석 시 주의',
    '70.8%는 절감 가능 비율이 아니며 정상적인 예약 대기와 실제 운영 지연은 현재 기록만으로 구분 불가',
    tone='surface'
)

# ==============================================================================
# SLIDE 09: 주요 대기 구간 분석
# ==============================================================================
s09 = make_header_footer(
    9, '주요 대기 구간 분석',
    '현장 조사 시작 전 대기는\n절반 이상에서 하루를 초과했다',
    '큰 시간 비중이 일부 긴 사례 때문인지 여러 요청에서 나타나는 현상인지 분포를 확인했다',
    '12-3-1–12-3-3'
)
rect(s09, 0.65, 2.24, 7.65, 3.90, c='surface', line_c='rule', rad=True, rad_val=0.04)
rect(s09, 0.85, 2.38, 7.25, 0.42, c='pale', rad=True, rad_val=0.12)
txt(s09, 1.02, 2.43, 6.90, 0.32, '조사 시작 전 대기 분포 / 339건 / 단위: 건', size=14.5, color='accent', bold=True)

bar(
    s09, 0.80, 2.86, 7.30, 3.16,
    ['0시간', '0시간 초과–24시간 이하', '24시간 초과–72시간 이하', '72시간 초과–168시간 이하', '168시간 초과'],
    [('케이스 수', [39, 89, 63, 90, 58], 'bar')],
    120,
    accents={0: ['bar', 'bar', 'accent', 'accent', 'accent']}
)

kpi_09 = [
    ('대기 시간 중앙값', '54.5시간', 'surface', 'rule', 'ink'),
    ('약 90%가 이 시간 이하(P90)', '195.7시간', 'surface', 'rule', 'ink'),
    ('24시간 초과 211 / 339건', '62.2%', 'pale', 'accent', 'accent')
]
for i, (k_lab, k_val, bg_c, bd_c, val_c) in enumerate(kpi_09):
    y = 2.24 + i * 1.34
    rect(s09, 8.52, y, 4.16, 1.22, c=bg_c, line_c=bd_c, rad=True, rad_val=0.06, line_w=1.1 if i == 2 else 0.75)
    rect(s09, 8.70, y + 0.20, 0.08, 0.82, c='accent' if i == 2 else 'sage', rad=True, rad_val=0.45)
    txt(s09, 8.92, y + 0.14, 3.58, 0.52, k_val, size=27, color=val_c, bold=True, font='head')
    txt(s09, 8.92, y + 0.66, 3.58, 0.38, k_lab, size=13.5, color='ink', bold=True)

add_bottom_single(
    s09, '분포 해석',
    '긴 대기가 일부 극단값에 국한되지 않고 다수의 요청에서 나타나 수리 유형별 차이를 이어서 비교',
    right_note='24시간은 분포 확인용 기준이며 정해진 처리 기한이 아님',
    tone='pale',
    note_w_override=3.55
)

# ==============================================================================
# SLIDE 10: 수리 유형별 비교 (유지)
# ==============================================================================
s10 = make_header_footer(
    10, '수리 유형별 비교',
    'E 유형은 전체 처리 시간과\n수리 시작 전 대기가 모두 길었다',
    '수리 유형별 전체 처리 시간을 비교한 뒤 수리 시작 전 대기를 확인 / 업무 구성의 차이를 줄이기 위해 InternRepair(내부 수리)만 수행한 요청끼리 추가 비교',
    '14-1 / 14-4-1–14-4-2'
)
rect(s10, 0.65, 2.24, 7.10, 3.90, c='surface', line_c='rule', rad=True, rad_val=0.04)
rect(s10, 0.85, 2.38, 6.70, 0.42, c='pale', rad=True, rad_val=0.12)
txt(s10, 1.02, 2.43, 6.35, 0.32, '전체 처리 시간 중앙값 / 339건 / 시간', size=14.5, color='accent', bold=True)

bar(
    s10, 0.80, 2.86, 6.80, 3.16,
    ['B (46건)', 'E (143건)', 'P (150건)'],
    [('중앙값', [11.6, 160.4, 71.4], 'bar')],
    200, fmt='0.0',
    accents={0: ['bar', 'accent', 'bar']}
)

rect(s10, 7.95, 2.24, 4.73, 1.78, c='pale', line_c='accent', rad=True, rad_val=0.05, line_w=1.1)
pill(s10, 8.15, 2.38, 2.55, 0.32, 'E 유형 전체 · 수리 시작 전 대기', fill='accent', text_c='white', size=11, bold=True)
txt(s10, 10.75, 2.38, 1.75, 0.32, '비교 표본 338건', size=10.5, color='muted', bold=True, align=PP_ALIGN.RIGHT)
txt(s10, 8.15, 2.78, 4.35, 0.54, '143건 · 중앙값 25.2시간', size=23, color='accent', bold=True, font='head')
txt(s10, 8.15, 3.36, 4.42, 0.30, 'Survey complete → 첫 수리 start (B 46건 0.0h · P 149건 0.0h)', size=10.5, color='ink', bold=True)
txt(s10, 8.15, 3.66, 4.42, 0.26, '복합·단독 수리를 모두 포함한 E 유형 전체', size=10.5, color='muted')

rect(s10, 7.95, 4.14, 4.73, 2.00, c='surface', line_c='rule', rad=True, rad_val=0.05)
txt(s10, 8.15, 4.24, 4.35, 0.32, '같은 내부 수리(InternRepair)만 수행한 요청 비교', size=13.5, color='ink', bold=True)
txt(s10, 8.15, 4.54, 4.35, 0.22, '단독 수행 · 시작·완료 1:1 연결 160건', size=11, color='muted')
rect(s10, 8.15, 4.80, 4.33, 0.30, c='ink', rad=True, rad_val=0.10)
txt(s10, 8.32, 4.82, 0.80, 0.26, '유형', size=11.5, color='white', bold=True)
txt(s10, 9.30, 4.82, 1.05, 0.26, '건수', size=11.5, color='white', bold=True, align=PP_ALIGN.RIGHT)
txt(s10, 10.50, 4.82, 1.80, 0.26, '대기 중앙값', size=11.5, color='white', bold=True, align=PP_ALIGN.RIGHT)

for i, (t_code, t_cnt, t_med) in enumerate([('B', '25건', '0.0시간'), ('E', '127건', '25.9시간'), ('P', '8건', '0.9시간')]):
    ry = 5.14 + i * 0.31
    hi = i == 1
    rect(s10, 8.15, ry, 4.33, 0.28, c='pale' if hi else 'white', line_c='accent' if hi else 'rule', rad=True, rad_val=0.10)
    txt(s10, 8.32, ry + 0.02, 0.80, 0.24, t_code, size=13, color='accent' if hi else 'ink', bold=hi)
    txt(s10, 9.30, ry + 0.02, 1.05, 0.24, t_cnt, size=13, color='accent' if hi else 'ink', bold=hi, align=PP_ALIGN.RIGHT)
    txt(s10, 10.50, ry + 0.02, 1.80, 0.24, t_med, size=13, color='accent' if hi else 'ink', bold=hi, align=PP_ALIGN.RIGHT)

add_bottom_split(
    s10,
    '같은 내부 수리 업무로 범위를 좁혀도 E 유형에서 긴 대기가 관찰됨',
    '수리 유형이 지연의 원인이라는 뜻은 아니며\nP 유형은 표본이 8건으로 작음',
    left_w=7.10
)

# ==============================================================================
# SLIDE 11: 접수 채널과 담당자 분석 (하단 표현 수정)
# ==============================================================================
s11 = make_header_footer(
    11, '접수 채널과 담당자 분석',
    'Web·Letter 접수에서 조사 일정 조율 업무의\n장시간 처리 사례가 집중됐다',
    '이번에는 조사 일정 조율 업무 자체의 시간을 비교했다 / 접수 채널과 시작·완료 담당자 기록에 따른 차이를 살펴봤다',
    '12-1 / 13-3–13-4'
)
rect(s11, 0.65, 2.24, 7.35, 3.90, c='surface', line_c='rule', rad=True, rad_val=0.04)
rect(s11, 0.85, 2.40, 6.95, 0.44, c='ink', rad=True, rad_val=0.10)
for cx, cw, ct in [(1.02, 1.35, '접수 채널'), (2.45, 0.95, '건수'), (3.55, 1.30, '중앙값(분)'), (5.05, 2.55, '1시간 초과율')]:
    txt(s11, cx, 2.46, cw, 0.32, ct, size=13, color='white', bold=True)

ch_rows = [
    ('Web', 207, 119, 52.7, True),
    ('Letter', 175, 5, 48.6, True),
    ('Phone', 224, 3, 0.0, False),
    ('Personal', 203, 4, 0.0, False)
]
for j, (ch_name, ch_n, ch_med, ch_rate, hi) in enumerate(ch_rows):
    ry = 2.94 + j * 0.58
    rect(s11, 0.85, ry, 6.95, 0.50, c='pale' if j == 0 else 'white', line_c='accent' if j == 0 else 'rule', rad=True, rad_val=0.08)
    txt(s11, 1.02, ry + 0.09, 1.35, 0.32, ch_name, size=14.5, color='ink', bold=hi)
    txt(s11, 2.45, ry + 0.09, 0.95, 0.32, str(ch_n), size=14.5, color='ink')
    txt(s11, 3.55, ry + 0.09, 1.30, 0.32, str(ch_med), size=14.5, color='accent' if j == 0 else 'ink', bold=j == 0)
    rect(s11, 5.05, ry + 0.18, 1.45, 0.15, c='surface', rad=True, rad_val=0.30)
    if ch_rate > 0:
        rect(s11, 5.05, ry + 0.18, 1.45 * (ch_rate / 100.0), 0.15, c='accent', rad=True, rad_val=0.30)
    txt(s11, 6.60, ry + 0.09, 1.00, 0.32, f'{ch_rate:.1f}%', size=14.5, color='accent' if hi else 'muted', bold=hi, align=PP_ALIGN.RIGHT)

rect(s11, 0.85, 5.34, 6.95, 0.60, c='pale', line_c='pale_line', rad=True, rad_val=0.10)
pill(s11, 1.02, 5.49, 1.05, 0.30, '집중 현상', fill='accent', text_c='white', size=11, bold=True)
txt(s11, 2.22, 5.45, 5.40, 0.38, '1시간 초과 194건은 모두 Web·Letter에서 발생', size=14.5, color='ink', bold=True)

rect(s11, 8.22, 2.24, 4.46, 3.90, c='surface', line_c='rule', rad=True, rad_val=0.04)
rect(s11, 8.42, 2.38, 4.06, 0.60, c='pale', rad=True, rad_val=0.12)
txt(s11, 8.56, 2.40, 3.78, 0.56, '시작·완료 담당자별 중앙값\nArrangeSurvey / 809건', size=13, color='accent', bold=True, align=PP_ALIGN.CENTER, line_sp=1.12)

for i, (r_lab, r_cnt, r_val, hi) in enumerate([
    ('동일 담당자', '704건', '4분', False),
    ('서로 다른 담당자', '105건', '332분', True)
]):
    ry = 3.12 + i * 1.04
    rect(s11, 8.42, ry, 4.06, 0.92, c='pale' if hi else 'white', line_c='accent' if hi else 'rule', rad=True, rad_val=0.08)
    txt(s11, 8.64, ry + 0.14, 2.10, 0.32, r_lab, size=14.5, color='ink', bold=True)
    txt(s11, 8.64, ry + 0.50, 2.10, 0.28, r_cnt, size=12.5, color='muted')
    txt(s11, 10.65, ry + 0.14, 1.62, 0.62, r_val, size=25, color='accent' if hi else 'ink', bold=True, font='head', align=PP_ALIGN.RIGHT)

rect(s11, 8.42, 5.24, 4.06, 0.70, c='white', line_c='rule', rad=True, rad_val=0.08)
txt(s11, 8.58, 5.32, 3.74, 0.54, '동일 담당자의 Web·Letter에서도\n1시간 초과 89건 관찰', size=13, color='ink', bold=True, align=PP_ALIGN.CENTER)

add_bottom_single(
    s11, '해석 시 주의',
    '시작·완료 담당자 차이만으로 실제 업무 인수인계를 단정할 수 없음 · 채널별 처리 절차와 담당자 변경 사유를 추가 점검',
    tone='surface'
)

# ==============================================================================
# SLIDE 12: 개선 방안 — [3순위: 확인 후 실행할 조치 구체화]
# ==============================================================================
s12 = make_header_footer(
    12, '개선 방안',
    '예정 시각과 진행 상태를 기록하고\n실제 지연이 확인된 단계부터 개선한다',
    '기록상 긴 대기를 곧바로 불필요한 지연으로 보지 않고 확인할 정보와 실행할 조치를 연결',
    '15-2–15-5'
)
pill(s12, 0.65, 2.24, 3.60, 0.44, '01 · 개선 검토 대상', fill='ink', text_c='white', size=13.5, bold=True, rad_val=0.25)
pill(s12, 4.65, 2.24, 3.80, 0.44, '02 · 추가로 기록할 정보', fill='accent', text_c='white', size=13.5, bold=True, rad_val=0.25)
pill(s12, 8.85, 2.24, 3.83, 0.44, '03 · 확인 후 실행할 조치', fill='ink', text_c='white', size=13.5, bold=True, rad_val=0.25)

action_rows = [
    ('01', '조사 시작 전 대기\n전체 시간 합계의 70.8%', '조사 예정 시각·착수 상태\n예정 초과 사유', '예정 시각 초과 사례를 확인하고\n지연 사유에 따라 일정 재조정'),
    ('02', 'E 유형 수리 시작 전 대기\n중앙값 25.2시간', '조사 완료 이후의 진행 상태\n수리 예정 시각·지연 사유', '반복되는 지연 원인을 확인하고\n수리 착수 일정과 처리 절차 개선'),
    ('03', 'Web·Letter 일정 조율\n1시간 초과 194건', '채널별 업무 진행 상태\n시작·완료 담당자 기록 차이', '채널별 처리 과정과 담당자 기록을 점검하고\n확인된 원인에 맞춰 개선')
]
for i, (num_tag, col1, col2, col3) in enumerate(action_rows):
    ry = 2.84 + i * 1.12
    hi = i == 0
    rect(s12, 0.65, ry, 3.60, 0.98, c='pale' if hi else 'surface', line_c='accent' if hi else 'rule', rad=True, rad_val=0.08)
    pill(s12, 0.83, ry + 0.31, 0.48, 0.36, num_tag, fill='accent' if hi else 'ink', text_c='white', size=11.5, bold=True)
    txt(s12, 1.44, ry + 0.16, 2.68, 0.66, col1, size=14, color='ink', bold=True, line_sp=1.18)

    line(s12, 4.25, ry + 0.49, 4.65, ry + 0.49, color='accent', width=1.6, arrow=True)

    rect(s12, 4.65, ry, 3.80, 0.98, c='white', line_c='accent', rad=True, rad_val=0.08, line_w=1.1)
    txt(s12, 4.88, ry + 0.16, 3.34, 0.66, col2, size=14, color='ink', bold=True, align=PP_ALIGN.CENTER, line_sp=1.18)

    line(s12, 8.45, ry + 0.49, 8.85, ry + 0.49, color='accent', width=1.6, arrow=True)

    rect(s12, 8.85, ry, 3.83, 0.98, c='pale' if hi else 'surface', line_c='accent' if hi else 'rule', rad=True, rad_val=0.08)
    txt(s12, 9.02, ry + 0.14, 3.50, 0.70, col3, size=13.2, color='ink', line_sp=1.18)

add_bottom_single(
    s12, '적용 원칙',
    '추가 기록으로 실제 지연 원인을 확인한 뒤 필요한 조치를 시범 적용',
    right_note='신규 기록 수집이 선행 · 특정 원인을 미리 단정하지 않음',
    tone='pale'
)

# ==============================================================================
# SLIDE 13: 효과 검증과 분석 한계 — [4순위: 부제목 간결화 · 중복 제거 · 최종 결론 강조]
# ==============================================================================
s13 = make_header_footer(
    13, '효과 검증과 분석 한계',
    '예정 시각과 지연 사유를 추가로 기록하고\n대기와 전체 완료 시간이 함께 줄었는지 검증한다',
    '신뢰할 수 있는 시간 기록을 확보하고 수리 유형과 접수 채널을 구분하여 개선 전후의 처리 시간을 비교',
    '9-4–9-5 / 15-5 / 개선 후 측정 설계'
)
# ① 왼쪽: 개선 과제별 핵심 측정 지표
rect(s13, 0.65, 2.24, 7.15, 3.90, c='surface', line_c='rule', rad=True, rad_val=0.04)
rect(s13, 0.85, 2.40, 6.75, 0.46, c='pale', rad=True, rad_val=0.12)
txt(s13, 1.04, 2.47, 4.20, 0.32, '개선 과제별 핵심 측정 지표', size=15, color='accent', bold=True)
txt(s13, 5.20, 2.47, 2.20, 0.32, '개선 과제  →  핵심 측정 지표', size=11, color='muted', bold=True, align=PP_ALIGN.RIGHT)

m_rows = [
    ('현장 조사 일정', '조사 착수 지연 시간 · 예정 시각 초과 미착수율'),
    ('E 유형 수리 착수', '수리 시작 전 대기 · 예정 시각 초과 미착수율'),
    ('Web·Letter 일정 조율', 'ArrangeSurvey 1시간 초과율 · P90')
]
for i, (m_cat, m_val) in enumerate(m_rows):
    ry = 3.02 + i * 0.76
    rect(s13, 0.85, ry, 6.75, 0.64, c='white', line_c='rule', rad=True, rad_val=0.08)
    pill(s13, 1.02, ry + 0.14, 1.82, 0.36, m_cat, fill='pale', text_c='accent', size=12, bold=True)
    txt(s13, 3.00, ry + 0.12, 4.45, 0.40, m_val, size=15, color='ink', bold=True)

rect(s13, 0.85, 5.34, 6.75, 0.62, c='pale', line_c='accent', rad=True, rad_val=0.10)
pill(s13, 1.00, 5.49, 1.05, 0.32, '측정 전제', fill='accent', text_c='white', size=10.5, bold=True)
txt(s13, 2.16, 5.41, 5.32, 0.48, '신규 지표는 추가 데이터 수집 이후 계산 가능하며 개선은 아직 실행·검증되지 않음', size=11.8, color='ink', bold=True)

# ② 오른쪽: 최종 성과 지표 + 분석 한계 요약 (부제목과의 설명 중복 제거)
rect(s13, 8.02, 2.24, 4.66, 3.90, c='ink', rad=True, rad_val=0.05)
pill(s13, 8.32, 2.50, 1.55, 0.34, '최종 성과 지표', fill='accent_hi', text_c='ink', size=11.5, bold=True)
txt(s13, 8.32, 3.04, 4.06, 0.56, '수리 요청 전체 처리 시간', size=24, color='white', bold=True, font='head')
txt(s13, 8.32, 3.64, 4.06, 0.64, '중앙값 · P90', size=32, color='accent_hi', bold=True, font='head')
txt(s13, 8.32, 4.30, 4.06, 0.30, '개별 구간과 전체 완료 시간의 동시 감소를 확인', size=12, color='pale')

line(s13, 8.32, 4.70, 12.38, 4.70, color='dark_line', width=0.9)
rect(s13, 8.24, 4.86, 4.22, 1.08, c='dark_card', line_c='dark_line', rad=True, rad_val=0.08)
pill(s13, 8.40, 4.98, 0.98, 0.26, '분석 한계', fill='ink', text_c='accent_hi', size=9.5, bold=True, line_c='dark_line')
txt(s13, 8.40, 5.30, 3.92, 0.54, '날짜 정확성 미검증 · 예정 시각 미기록 · 외부 수리 완료 시각 없음\n시간 분석 대상 339건(일반 경로 927건의 36.6%)', size=10.5, color='pale', line_sp=1.18)

# ③ 하단 1줄: 최종 결론 강조
add_bottom_single(
    s13, '최종 결론',
    '현장 조사 시작 전 대기를 우선 검토 대상으로 선정하고 실제 지연 원인 확인 후 전체 처리 시간의 감소 여부로 개선 효과를 검증',
    tone='pale'
)

# Restore speaker notes on all 13 slides
for idx, sl in enumerate(p.slides):
    base_note = prev_notes[idx] if idx < len(prev_notes) else ''
    sl.notes_slide.notes_text_frame.text = base_note
    for sh in sl.shapes:
        strip_effects(sh)

p.save(out)

# Export to PDF via PowerPoint and verify geometry/text bounds + line wrap check
app = win32com.client.Dispatch('PowerPoint.Application')
deck = app.Presentations.Open(str(out), False, False, False)
overflow = []
try:
    for sl in deck.Slides:
        for sh in sl.Shapes:
            if sh.HasTextFrame and sh.TextFrame.HasText:
                tr = sh.TextFrame.TextRange
                expected_lines = tr.Text.count('\r') + 1
                actual_lines = tr.Lines().Count
                if tr.BoundHeight > sh.Height + 2 or tr.BoundWidth > sh.Width + 2 or actual_lines > expected_lines:
                    overflow.append({
                        'page': sl.SlideIndex,
                        'text': tr.Text,
                        'expected_lines': expected_lines,
                        'actual_lines': actual_lines,
                        'h': round(sh.Height, 2),
                        'bh': round(tr.BoundHeight, 2),
                        'w': round(sh.Width, 2),
                        'bw': round(tr.BoundWidth, 2)
                    })
    deck.Save()
    deck.SaveAs(str(pdf), 32)
finally:
    deck.Close()

pp = Presentation(out)
doc = fitz.open(pdf)
norm = lambda t: ''.join(t.split())
mismatch = []
bounds = []
for i, s in enumerate(pp.slides):
    pdf_txt = norm(doc[i].get_text())
    for sh in s.shapes:
        if sh.has_text_frame and norm(sh.text) not in pdf_txt:
            mismatch.append([i + 1, sh.text])
        if sh.left < 0 or sh.top < 0 or sh.left + sh.width > pp.slide_width + 1000 or sh.top + sh.height > pp.slide_height + 1000:
            bounds.append([i + 1, sh.name])
    doc[i].get_pixmap(matrix=fitz.Matrix(1.5, 1.5)).save(Q / f'page-{i+1:02}.png')

for start in range(0, 13, 4):
    sheet = fitz.open()
    pg = sheet.new_page(width=1600, height=900)
    for j, i in enumerate(range(start, min(start + 4, 13))):
        pg.show_pdf_page(fitz.Rect(j % 2 * 800, j // 2 * 450, (j % 2 + 1) * 800, (j // 2 + 1) * 450), doc, i)
    pg.get_pixmap().save(Q / f'sheet-{start+1:02}.png')

report = {
    'version': 10,
    'pages': len(doc),
    'overflow': overflow,
    'text_mismatch': mismatch,
    'shape_bounds': bounds,
    'raster_images': sum(len(pg.get_images()) for pg in doc),
    'native_charts': sum(sh.has_chart for sl in pp.slides for sh in sl.shapes),
    'fonts': sorted({x[3] for pg in doc for x in pg.get_fonts()}),
}
(Q / 'validation.json').write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding='utf-8')
print(json.dumps(report, ensure_ascii=False, indent=2))
