from pathlib import Path

W = Path(__file__).resolve().parent
code = (W / 'build_v010.py').read_text(encoding='utf-8')
code = code.replace("Q = W / '검수_v010'", "Q = W / '검수_v011'")
code = code.replace("out = V / 'repair2_채용과제_분석결과_v010.pptx'", "out = V / 'repair2_채용과제_분석결과_v011.pptx'")
code = code.replace("prev_ppt = Presentation(V / 'repair2_채용과제_분석결과_v009.pptx')", "prev_ppt = Presentation(V / 'repair2_채용과제_분석결과_v010.pptx')")

# Update make_header_footer prefix from 'repair2 분석 / ' to 'repair2_분석.ipynb  |  '
code = code.replace(
    "txt(s, 0.65, 7.15, 11.2, 0.22, 'repair2 분석 / ' + source, size=9.5, color='muted')",
    "txt(s, 0.65, 7.15, 11.4, 0.22, 'repair2_분석.ipynb  |  ' + source, size=9.2, color='muted')"
)

# Replace all 12 slide source strings with uniform notebook section locations
replacements = [
    (
        "'분석 질문 재구성 / 노트북 7–15'",
        "'7. 기본 분석 — EDA 이전 작업 ~ 15. 분석 결과 종합 및 개선안 도출'"
    ),
    (
        "'1–5 / 8–9'",
        "'1. 데이터 불러오기 ~ 5. 결측값 검토 · 7-1~7-3 · 8-2. 종료 후보 · 9-5. 시간 분석 대상 확정'"
    ),
    (
        "'9-1–9-5'",
        "'9. 날짜 검증 및 복원 (9-1. 날짜 이상 패턴 확인 · 9-2. 시간 정합성 점검 · 9-4. 처리 방향 결정 · 9-5. 대상 확정)'"
    ),
    (
        "'8-2 / 10-3–10-5'",
        "'8-2. 종료 업무 후보 확인 · 10. 프로세스 흐름 및 분기 분석 (10-3. 주요 경로 · 10-4. 분기 지점 · 10-5. 프로세스 맵)'"
    ),
    (
        "'10-3 / 10-5 / 12-5'",
        "'10-3. 주요 업무 경로 및 반복 수행 분석 (10-3-1~10-3-3) · 12-5. 반복 수리 및 복합 수리 비교'"
    ),
    (
        "'12-1–12-2'",
        "'12-1. 업무별 처리 시간 계산 및 분포 비교 (12-1-1~12-1-2) · 12-2-2. ArrangeSurvey 처리 시간별 전체 시간 비교'"
    ),
    (
        "'12-2-1 / 12-4-1'",
        "'12-2-1. 케이스별 전체 관찰 시간 분석 · 12-4-1. 전체 처리 시간을 구간별로 나누어 비교'"
    ),
    (
        "'12-3-1–12-3-3'",
        "'12-3. 업무 간 대기 시간 분석 (12-3-1. 업무 간 대기 시간 비교 · 12-3-2. ArrangeSurvey → Survey 대기 상세 분석)'"
    ),
    (
        "'14-1 / 14-4-1–14-4-2'",
        "'14. 수리 속성과 소요 시간의 관계 분석 (14-1. RepairType별 전체 시간 · 14-4-1~14-4-2. 수리 업무 구성과 대기 비교)'"
    ),
    (
        "'12-1 / 13-3–13-4'",
        "'12-1-4. 접수 채널별 처리 시간 비교 · 13. 담당자별 업무 분석 (13-3. 병목 연관성 · 13-4. 담당자 변경과 처리 시간)'"
    ),
    (
        "'15-2–15-5'",
        "'15. 분석 결과 종합 및 개선안 도출 (15-2. 조사 대기 개선 · 15-3. E 유형 대기 개선 · 15-4. 채널·담당자 개선 · 15-5)'"
    ),
    (
        "'9-4–9-5 / 15-5 / 개선 후 측정 설계'",
        "'9-4~9-5. 날짜 처리 및 시간 분석 대상 확정 · 15-2~15-5. 개선 과제 우선순위 및 최종 제안'"
    ),
]

for old_s, new_s in replacements:
    assert old_s in code, f"Not found: {old_s}"
    code = code.replace(old_s, new_s)

code = code.replace("'version': 10,", "'version': 11,")
(W / 'build_v011.py').write_text(code, encoding='utf-8')
print('build_v011.py created')
