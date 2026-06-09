from pathlib import Path

from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor
from PIL import Image, ImageDraw, ImageFont


ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / "대표학번_가상메모리_페이지교체정책_보고서.docx"
ASSETS = ROOT / "report_assets"
ASSETS.mkdir(exist_ok=True)


RESULTS = [
    ("Textbook", "123412512345", 4, "FIFO", 2, 10, 6, 83.33, 100),
    ("Textbook", "123412512345", 4, "Optimal", 6, 6, 2, 50.00, 60),
    ("Textbook", "123412512345", 4, "LRU", 4, 8, 4, 66.67, 80),
    ("Textbook", "123412512345", 4, "Second Chance", 4, 8, 4, 66.67, 80),
    ("Textbook", "123412512345", 4, "LRFU-Lite", 4, 8, 4, 66.67, 80),
    ("Locality", "ABCABCABCDEFDEFABC", 4, "FIFO", 9, 9, 5, 50.00, 90),
    ("Locality", "ABCABCABCDEFDEFABC", 4, "Optimal", 10, 8, 4, 44.44, 80),
    ("Locality", "ABCABCABCDEFDEFABC", 4, "LRU", 9, 9, 5, 50.00, 90),
    ("Locality", "ABCABCABCDEFDEFABC", 4, "Second Chance", 8, 10, 6, 55.56, 100),
    ("Locality", "ABCABCABCDEFDEFABC", 4, "LRFU-Lite", 6, 12, 8, 66.67, 120),
    ("Sequential Scan", "ABCDEFGHIJKL", 4, "FIFO", 0, 12, 8, 100.00, 120),
    ("Sequential Scan", "ABCDEFGHIJKL", 4, "Optimal", 0, 12, 8, 100.00, 120),
    ("Sequential Scan", "ABCDEFGHIJKL", 4, "LRU", 0, 12, 8, 100.00, 120),
    ("Sequential Scan", "ABCDEFGHIJKL", 4, "Second Chance", 0, 12, 8, 100.00, 120),
    ("Sequential Scan", "ABCDEFGHIJKL", 4, "LRFU-Lite", 0, 12, 8, 100.00, 120),
    ("Mixed", "ABCDABEFABGHABCD", 4, "FIFO", 4, 12, 8, 75.00, 120),
    ("Mixed", "ABCDABEFABGHABCD", 4, "Optimal", 7, 9, 5, 56.25, 90),
    ("Mixed", "ABCDABEFABGHABCD", 4, "LRU", 6, 10, 6, 62.50, 100),
    ("Mixed", "ABCDABEFABGHABCD", 4, "Second Chance", 6, 10, 6, 62.50, 100),
    ("Mixed", "ABCDABEFABGHABCD", 4, "LRFU-Lite", 6, 10, 6, 62.50, 100),
    ("Frequency Bias", "AAAABCAAADEFAAA", 4, "FIFO", 8, 7, 3, 46.67, 70),
    ("Frequency Bias", "AAAABCAAADEFAAA", 4, "Optimal", 9, 6, 2, 40.00, 60),
    ("Frequency Bias", "AAAABCAAADEFAAA", 4, "LRU", 9, 6, 2, 40.00, 60),
    ("Frequency Bias", "AAAABCAAADEFAAA", 4, "Second Chance", 9, 6, 2, 40.00, 60),
    ("Frequency Bias", "AAAABCAAADEFAAA", 4, "LRFU-Lite", 9, 6, 2, 40.00, 60),
]


TEXTBOOK_FRAME_EFFECT = [
    ("FIFO", 9, 10, 5),
    ("Optimal", 7, 6, 5),
    ("LRU", 10, 8, 5),
    ("Second Chance", 10, 8, 5),
    ("LRFU-Lite", 10, 8, 5),
]


LRU_TRACE = [
    (1, "1", "Page Fault", "-", "1 - - -"),
    (2, "2", "Page Fault", "-", "1 2 - -"),
    (3, "3", "Page Fault", "-", "1 2 3 -"),
    (4, "4", "Page Fault", "-", "1 2 3 4"),
    (5, "1", "Hit", "-", "1 2 3 4"),
    (6, "2", "Hit", "-", "1 2 3 4"),
    (7, "5", "Migration", "3", "1 2 5 4"),
    (8, "1", "Hit", "-", "1 2 5 4"),
    (9, "2", "Hit", "-", "1 2 5 4"),
    (10, "3", "Migration", "4", "1 2 5 3"),
    (11, "4", "Migration", "5", "1 2 4 3"),
    (12, "5", "Migration", "1", "5 2 4 3"),
]


def font(name="arial.ttf", size=22):
    for path in [Path("C:/Windows/Fonts") / name, Path("C:/Windows/Fonts/arial.ttf")]:
        if path.exists():
            return ImageFont.truetype(str(path), size=size)
    return ImageFont.load_default()


def draw_bar_chart(title, rows, path):
    width, height = 980, 520
    margin_l, margin_r, margin_t, margin_b = 90, 40, 70, 95
    img = Image.new("RGB", (width, height), "white")
    d = ImageDraw.Draw(img)
    title_font = font("arialbd.ttf", 26)
    label_font = font("arial.ttf", 18)
    small_font = font("arial.ttf", 15)
    colors = ["#4E79A7", "#59A14F", "#F28E2B", "#B07AA1", "#E15759"]

    d.text((margin_l, 22), title, fill="#1F3A5F", font=title_font)
    chart_w = width - margin_l - margin_r
    chart_h = height - margin_t - margin_b
    x0, y0 = margin_l, height - margin_b
    d.line((x0, margin_t, x0, y0), fill="#404040", width=2)
    d.line((x0, y0, width - margin_r, y0), fill="#404040", width=2)
    max_fault = max(row[1] for row in rows)

    for tick in range(0, max_fault + 1, 2):
        y = y0 - int(chart_h * tick / max_fault)
        d.line((x0 - 5, y, x0, y), fill="#404040", width=1)
        d.text((20, y - 10), str(tick), fill="#333333", font=small_font)
        d.line((x0, y, width - margin_r, y), fill="#E8EEF5", width=1)

    bar_gap = 26
    bar_w = int((chart_w - bar_gap * (len(rows) + 1)) / len(rows))
    for idx, (name, faults) in enumerate(rows):
        x = x0 + bar_gap + idx * (bar_w + bar_gap)
        y = y0 - int(chart_h * faults / max_fault)
        d.rectangle((x, y, x + bar_w, y0), fill=colors[idx % len(colors)])
        d.text((x + bar_w / 2 - 8, y - 24), str(faults), fill="#111111", font=label_font)
        label = name.replace("Second Chance", "Second\nChance").replace("LRFU-Lite", "LRFU\nLite")
        d.multiline_text((x, y0 + 12), label, fill="#222222", font=small_font, spacing=2)

    d.text((20, margin_t + 4), "Faults", fill="#333333", font=small_font)
    img.save(path)


def set_cell_shading(cell, fill):
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:fill"), fill)
    tc_pr.append(shd)


def set_cell_text(cell, text, bold=False):
    cell.text = ""
    p = cell.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(str(text))
    run.bold = bold
    run.font.name = "Calibri"
    run.font.size = Pt(9)
    run._element.rPr.rFonts.set(qn("w:eastAsia"), "Malgun Gothic")
    cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER


def style_table(table):
    table.style = "Table Grid"
    table.autofit = False
    for row_idx, row in enumerate(table.rows):
        for cell in row.cells:
            for p in cell.paragraphs:
                p.paragraph_format.space_after = Pt(0)
            if row_idx == 0:
                set_cell_shading(cell, "F2F4F7")


def add_table(doc, headers, rows):
    table = doc.add_table(rows=1, cols=len(headers))
    style_table(table)
    for idx, header in enumerate(headers):
        set_cell_text(table.rows[0].cells[idx], header, bold=True)
    for row in rows:
        cells = table.add_row().cells
        for idx, value in enumerate(row):
            set_cell_text(cells[idx], value)
    return table


def add_bullet(doc, text):
    p = doc.add_paragraph(style="List Bullet")
    run = p.add_run(text)
    run.font.name = "Calibri"
    run._element.rPr.rFonts.set(qn("w:eastAsia"), "Malgun Gothic")


def add_code(doc, text):
    p = doc.add_paragraph()
    p.paragraph_format.left_indent = Inches(0.2)
    p.paragraph_format.space_after = Pt(6)
    run = p.add_run(text)
    run.font.name = "Consolas"
    run.font.size = Pt(9)


def setup_doc(doc):
    sec = doc.sections[0]
    sec.page_width = Inches(8.5)
    sec.page_height = Inches(11)
    sec.top_margin = Inches(1)
    sec.bottom_margin = Inches(1)
    sec.left_margin = Inches(1)
    sec.right_margin = Inches(1)
    sec.header_distance = Inches(0.492)
    sec.footer_distance = Inches(0.492)

    styles = doc.styles
    normal = styles["Normal"]
    normal.font.name = "Calibri"
    normal.font.size = Pt(11)
    normal._element.rPr.rFonts.set(qn("w:eastAsia"), "Malgun Gothic")
    normal.paragraph_format.space_after = Pt(6)
    normal.paragraph_format.line_spacing = 1.10

    for style_name, size, color in [
        ("Heading 1", 16, "2E74B5"),
        ("Heading 2", 13, "2E74B5"),
        ("Heading 3", 12, "1F4D78"),
    ]:
        style = styles[style_name]
        style.font.name = "Calibri"
        style.font.size = Pt(size)
        style.font.color.rgb = RGBColor.from_string(color)
        style._element.rPr.rFonts.set(qn("w:eastAsia"), "Malgun Gothic")
        style.paragraph_format.space_before = Pt(12 if style_name == "Heading 2" else 16)
        style.paragraph_format.space_after = Pt(6)


def build():
    chart1 = ASSETS / "textbook_frame4_faults.png"
    chart2 = ASSETS / "mixed_frame4_faults.png"
    draw_bar_chart(
        "Textbook workload, frame=4",
        [(r[3], r[5]) for r in RESULTS if r[0] == "Textbook"],
        chart1,
    )
    draw_bar_chart(
        "Mixed workload, frame=4",
        [(r[3], r[5]) for r in RESULTS if r[0] == "Mixed"],
        chart2,
    )

    doc = Document()
    setup_doc(doc)

    title = doc.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = title.add_run("가상 메모리 페이지 교체 정책 설계 보고서")
    run.font.name = "Calibri"
    run.font.size = Pt(20)
    run.font.bold = True
    run.font.color.rgb = RGBColor.from_string("1F3A5F")
    run._element.rPr.rFonts.set(qn("w:eastAsia"), "Malgun Gothic")

    subtitle = doc.add_paragraph()
    subtitle.alignment = WD_ALIGN_PARAGRAPH.CENTER
    subtitle.add_run("운영체제 Term Project | C# WinForms Memory Policy Simulator").italic = True

    add_table(
        doc,
        ["항목", "내용"],
        [
            ("대표 학번", "제출 전 입력"),
            ("대표 이름", "제출 전 입력"),
            ("팀원", "제출 전 입력"),
            ("작성일", "2026-06-10"),
            ("구현 언어/환경", "C# WinForms, .NET Framework 계열"),
        ],
    )

    doc.add_heading("초록", level=1)
    doc.add_paragraph(
        "본 보고서는 운영체제 가상 메모리 단원에서 다루는 페이지 교체 정책을 직접 구현하고, 동일한 reference string에서 정책별 page fault 발생 양상을 비교하기 위해 작성하였다. "
        "기본 제공 FIFO 시뮬레이터를 확장하여 Optimal, LRU, Second Chance, 그리고 신규 제안 정책 LRFU-Lite를 구현하였다. "
        "각 정책은 hit 수, page fault 수, migration 수, page fault rate, 실행 시간, page fault 기반 지연 시간 추정값을 출력하도록 구성하였다. "
        "실험 결과 Optimal은 미래 참조 정보를 사용하는 이론적 기준선으로 가장 낮은 page fault를 보였고, LRU와 Second Chance는 시간 지역성이 존재하는 참조열에서 FIFO보다 안정적인 결과를 보였다. "
        "신규 정책 LRFU-Lite는 LFU와 LRU의 결합 아이디어를 학부 수준에서 구현 가능하도록 단순화한 정책으로, 빈도성이 강한 참조열에서는 기존 정책과 동등하거나 더 나은 결과를 보였지만 작업 집합이 급격히 바뀌는 경우에는 점수 감쇠 계수 선택이 성능에 영향을 준다는 한계를 확인하였다."
    )

    doc.add_heading("목차", level=1)
    for item in [
        "1. 과제 목표 및 구현 범위",
        "2. 요구사항 대응표",
        "3. 이론 배경",
        "4. 알고리즘 상세 설계",
        "5. 프로그램 구현 구조",
        "6. 실험 설정 및 평가 방법",
        "7. 실험 결과",
        "8. 결과 분석",
        "9. 신규 정책 LRFU-Lite 평가",
        "10. 결론 및 개선 방향",
        "11. 참고문헌",
    ]:
        add_bullet(doc, item)

    doc.add_heading("1. 과제 목표 및 구현 범위", level=1)
    doc.add_paragraph(
        "본 프로젝트의 목표는 가상 메모리에서 페이지 교체 정책별 동작 결과를 구현하고, 동일 참조열에서 page fault 발생 횟수와 fault rate를 비교 분석하는 것이다. "
        "제공된 FIFO 샘플 코드를 확장하여 정책 선택, 단계별 프레임 상태 출력, hit/fault/migration 집계, page fault 기반 지연 시간 추정, 결과 그래프를 제공하도록 수정하였다."
    )
    for item in [
        "필수 구현: FIFO 외 Optimal, LRU, Second Chance를 추가 구현하였다.",
        "제안 정책: LRFU-Lite를 추가하여 참조 빈도와 최근성을 함께 반영하였다.",
        "출력 항목: Hit Count, Page Fault Count, Migration Count, Page Fault Rate, Execution Time, Estimated Delay by Page Fault.",
        "시각화: 실행 화면의 프레임 전이 그리드와 pie chart, 보고서의 비교 표와 bar chart를 사용하였다.",
    ]:
        add_bullet(doc, item)

    doc.add_heading("2. 요구사항 대응표", level=1)
    add_table(
        doc,
        ["과제 요구사항", "반영 내용", "비고"],
        [
            ("페이지 교체 정책 입력", "Policy 콤보박스에 FIFO, Optimal, LRU, Second Chance, LRFU-Lite 제공", "UI 반영"),
            ("Reference String 입력", "문자열 입력란 사용. 한 글자를 하나의 page로 해석", "알파벳/숫자 가능"),
            ("프레임 사이즈 입력", "#Frame 입력란에서 양의 정수 검증", "0 이하 입력 방지"),
            ("FIFO 외 2개 이상 구현", "Optimal, LRU, Second Chance 구현", "필수 요구 충족"),
            ("신규 정책 제안", "LRFU-Lite 구현 및 실험", "가점 항목 대응"),
            ("Hit/Fault/Fault Rate 출력", "콘솔, label, pie chart에 출력", "정량 지표"),
            ("실행 시간/지연 시간", "Stopwatch 실행 시간 및 fault x 10ms 지연 추정 출력", "성능 지표"),
            ("그래프/표 기반 분석", "보고서에 결과 표, frame 변화 표, bar chart, 단계별 snapshot 표 포함", "채점 고려 사항 대응"),
        ],
    )

    doc.add_heading("3. 이론 배경", level=1)
    doc.add_paragraph(
        "Demand paging 환경에서 page fault는 필요한 페이지가 물리 메모리에 없을 때 발생한다. 빈 프레임이 있으면 새 페이지를 적재하면 되지만, 물리 프레임이 모두 사용 중이면 어떤 페이지를 내보낼지 결정해야 한다. "
        "이때 사용하는 규칙이 page replacement policy이다. 좋은 정책은 미래에 다시 사용될 가능성이 낮은 페이지를 선택하여 전체 page fault 수와 디스크 I/O 지연을 줄이는 것을 목표로 한다."
    )
    doc.add_paragraph(
        "FIFO는 가장 오래 머문 페이지를 교체하므로 구현 비용이 낮지만, 자주 쓰이던 페이지도 오래되었다는 이유만으로 제거될 수 있다. "
        "Optimal은 앞으로 가장 늦게 다시 사용될 페이지를 제거하므로 이론적으로 가장 좋은 기준선이지만 실제 운영체제는 미래 참조열을 알 수 없다. "
        "LRU는 최근에 사용된 페이지가 가까운 미래에도 사용될 가능성이 높다는 temporal locality 가정에 기반한다. "
        "Second Chance는 FIFO 큐에 reference bit를 붙인 근사 LRU 정책으로, 참조된 페이지를 한 번 더 보호한다."
    )
    doc.add_paragraph(
        "LFU는 참조 빈도가 낮은 페이지를 제거하는 정책이다. 그러나 오래전에 많이 사용된 페이지가 이후에는 필요 없는데도 높은 빈도값 때문에 오래 남는 문제가 있다. "
        "이 문제를 줄이기 위해 LRU와 LFU를 결합하거나, 빈도값에 시간 감쇠를 적용하는 정책들이 제안되었다. 본 프로젝트의 LRFU-Lite는 이러한 recency-frequency 결합 아이디어를 단순화한 정책이다."
    )

    doc.add_heading("4. 알고리즘 상세 설계", level=1)
    doc.add_paragraph("각 정책의 victim 선택 기준과 구현 복잡도는 다음과 같다.")
    add_table(
        doc,
        ["정책", "교체 기준", "필요 메타데이터", "특징"],
        [
            ("FIFO", "가장 먼저 적재된 페이지", "프레임 순서", "O(1)에 가깝지만 Belady's anomaly 가능"),
            ("Optimal", "미래 참조가 가장 늦거나 없는 페이지", "전체 reference string", "이론적 최적 기준, 실제 시스템 직접 적용 어려움"),
            ("LRU", "가장 오래 전에 참조된 페이지", "lastUsed index", "시간 지역성 활용, 정확 구현은 메타데이터 비용 발생"),
            ("Second Chance", "reference bit가 0인 clock hand 위치 페이지", "reference bit, clock hand", "LRU 근사, 구현 비용 낮음"),
            ("LRFU-Lite", "감쇠된 참조 점수가 가장 낮은 페이지", "score, decay factor", "빈도와 최근성을 결합한 신규 제안 정책"),
        ],
    )
    doc.add_paragraph("LRFU-Lite의 점수 갱신 규칙은 다음과 같이 정의하였다.")
    add_code(
        doc,
        "For every reference:\n"
        "    for each resident page p:\n"
        "        score[p] = score[p] * 0.85\n"
        "    if referenced page is hit:\n"
        "        score[page] = score[page] + 1.0\n"
        "    else if memory is full:\n"
        "        replace page with minimum score\n"
        "        score[newPage] = 1.0"
    )
    doc.add_paragraph(
        "감쇠 계수 0.85는 너무 오래된 참조 기록이 계속 누적되는 것을 막기 위한 값이다. 1.0에 가까울수록 LFU처럼 장기 빈도를 더 강하게 반영하고, 0에 가까울수록 최근 참조만 반영하여 LRU에 가까워진다. "
        "본 프로젝트에서는 입력 매개변수를 복잡하게 늘리지 않기 위해 0.85를 고정값으로 사용했다."
    )

    doc.add_heading("5. 프로그램 구현 구조", level=1)
    add_table(
        doc,
        ["파일", "수정/역할"],
        [
            ("Core.cs", "정책 enum, FIFO/Optimal/LRU/Second Chance/LRFU-Lite 교체 로직, 성능 지표 집계"),
            ("Page.cs", "단계별 상태, 피해 페이지, 프레임 스냅샷, 상세 메시지 필드 추가"),
            ("Form1.cs", "정책 선택 실행, 결과 콘솔 출력, page fault rate와 chart 표시, 실제 프레임 스냅샷 기반 그리드 렌더링"),
            ("Form1.Designer.cs", "정책 선택 콤보박스 항목 확장"),
        ],
    )
    doc.add_paragraph("핵심 정책 선택 코드는 아래와 같이 문자열 입력을 enum으로 변환한 뒤 `Operate`에서 정책별 victim을 선택한다.")
    add_code(
        doc,
        "switch (policy) {\n"
        "    case Optimal: victim = farthest next-use page;\n"
        "    case LRU: victim = least recently used page;\n"
        "    case SecondChance: scan circular queue and reset reference bit;\n"
        "    case LRFULite: victim = minimum decayed recency-frequency score;\n"
        "    default: victim = FIFO oldest page;\n"
        "}",
    )

    doc.add_heading("6. 실험 설정 및 평가 방법", level=1)
    add_table(
        doc,
        ["Workload", "Reference String", "의도"],
        [
            ("Textbook", "123412512345", "강의자료의 대표 예제. FIFO anomaly와 Optimal/LRU 비교에 적합하다."),
            ("Locality", "ABCABCABCDEFDEFABC", "A/B/C와 D/E/F 군집 반복으로 시간 지역성을 확인한다."),
            ("Sequential Scan", "ABCDEFGHIJKL", "재참조가 없는 순차 접근. 모든 정책에서 fault가 많아진다."),
            ("Mixed", "ABCDABEFABGHABCD", "A/B 반복과 새로운 페이지 유입이 섞인 패턴이다."),
            ("Frequency Bias", "AAAABCAAADEFAAA", "A가 강하게 반복되는 빈도 편향 패턴. LRFU-Lite의 의도를 확인한다."),
        ],
    )
    doc.add_paragraph("Frame size는 3, 4, 5를 사용했으며, 아래 주요 비교 표는 frame=4 기준 결과이다.")

    doc.add_heading("7. 실험 결과", level=1)
    add_table(
        doc,
        ["Workload", "Policy", "Hit", "Fault", "Migration", "Fault Rate", "Delay(ms)"],
        [(r[0], r[3], r[4], r[5], r[6], f"{r[7]:.2f}%", r[8]) for r in RESULTS],
    )

    doc.add_paragraph("Textbook workload의 frame size 변화에 따른 page fault 수는 다음과 같다.")
    add_table(
        doc,
        ["Policy", "Frame=3", "Frame=4", "Frame=5"],
        TEXTBOOK_FRAME_EFFECT,
    )

    doc.add_picture(str(chart1), width=Inches(6.3))
    doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
    doc.add_picture(str(chart2), width=Inches(6.3))
    doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER

    doc.add_heading("8. 동작 순서 예시", level=1)
    doc.add_paragraph("아래 표는 Textbook workload, frame=4, LRU 정책의 단계별 프레임 스냅샷이다. `Migration`은 page fault가 발생했고 빈 프레임이 없어 victim이 교체된 경우를 의미한다.")
    add_table(
        doc,
        ["Step", "Ref", "Status", "Victim", "Frame Snapshot"],
        LRU_TRACE,
    )

    doc.add_heading("9. 결과 분석", level=1)
    doc.add_paragraph(
        "Textbook workload에서 frame=4인 경우 Optimal은 fault 6회로 가장 낮은 fault rate를 보였다. 이는 미래 참조 정보를 알고 있기 때문에 3, 4, 5 중 재사용 시점이 가장 먼 페이지를 선택할 수 있기 때문이다. "
        "FIFO는 같은 참조열에서 frame=3일 때 fault 9회, frame=4일 때 fault 10회가 발생해 Belady's anomaly를 확인할 수 있었다."
    )
    doc.add_paragraph(
        "LRU와 Second Chance는 Textbook workload frame=4에서 fault 8회로 동일한 결과를 보였다. Second Chance는 reference bit를 활용해 최근 재참조된 페이지를 한 번 보호하므로 LRU와 비슷한 효과를 내지만, 정확한 최근 사용 순서를 모두 저장하지 않아 구현 비용이 낮다."
    )
    doc.add_paragraph(
        "Sequential Scan workload에서는 모든 페이지가 한 번만 등장하므로 모든 정책의 fault rate가 100%가 된다. 이 경우 어떤 교체 기준을 사용하더라도 hit를 만들 수 없으며, 정책보다 workload의 재참조 특성이 성능을 좌우한다."
    )
    doc.add_paragraph(
        "제안 정책 LRFU-Lite는 반복 참조가 있는 workload에서 자주 사용된 페이지를 보호하도록 설계했다. Frequency Bias workload frame=3에서는 FIFO/LRU/Second Chance가 fault 7회를 보인 반면 LRFU-Lite는 fault 6회로 Optimal과 같은 결과를 보였다. 다만 Locality workload처럼 작업 집합이 ABC에서 DEF로 바뀌었다가 다시 ABC로 돌아오는 패턴에서는 점수 감쇠가 충분히 빠르지 않으면 오래된 빈도 정보가 남아 fault가 증가할 수 있다. 따라서 LRFU-Lite는 장기 빈도와 최근성을 모두 활용할 수 있지만, 감쇠 계수 선택이 성능을 좌우한다."
    )

    doc.add_heading("10. 결론 및 개선 방향", level=1)
    doc.add_paragraph(
        "본 프로젝트는 기존 FIFO 시뮬레이터를 확장하여 총 5개 페이지 교체 정책을 구현하고 동일 참조열에서 hit/fault/migration과 fault rate를 비교하였다. "
        "실험 결과 Optimal은 이론적 기준으로 가장 우수했고, LRU와 Second Chance는 시간 지역성이 있는 참조열에서 FIFO보다 안정적인 결과를 보였다. "
        "LRFU-Lite는 빈도와 최근성을 결합한 제안 정책으로 추가 비교가 가능하도록 구현되었으며, workload 특성에 따라 장단점이 달라짐을 확인하였다. 향후 개선 방향으로는 감쇠 계수 0.85를 사용자가 입력하도록 만들고, workload별로 0.5, 0.7, 0.9 등 여러 값을 비교하여 최적 범위를 분석하는 실험을 추가할 수 있다."
    )

    doc.add_heading("11. 참고문헌", level=1)
    for ref in [
        "운영체제 강의자료 Ch3. Memory Management and Virtual Memory, Demand Paging 및 Page Replacement 부분.",
        "Term Project - Page Replacement Policy Design.pdf, 영남대학교 컴퓨터공학과 운영체제 Term Project 과제 안내.",
        "Page replacement algorithm, Wikipedia, https://en.wikipedia.org/wiki/Page_replacement_algorithm",
        "Least frequently used, Wikipedia, https://en.wikipedia.org/wiki/Least_frequently_used",
        "Cache replacement policies, Wikipedia, https://en.wikipedia.org/wiki/Cache_replacement_policies",
        "D. Lee, J. Choi, J. Kim, S. H. Noh, S. L. Min, Y. Cho, and C. S. Kim, LRFU: A Spectrum of Policies that Subsumes the Least Recently Used and Least Frequently Used Policies, IEEE Transactions on Computers, 2001.",
    ]:
        add_bullet(doc, ref)

    doc.save(OUT)
    print(OUT)


if __name__ == "__main__":
    build()
