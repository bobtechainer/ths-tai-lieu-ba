# -*- coding: utf-8 -*-
"""
Xuất bài phân tích 2 thiết bị chạy hệ thống THS ra file .docx
Nội dung đã qua humanizer - giọng văn kỹ thuật tự nhiên, không AI-ish.
"""

from docx import Document
from docx.shared import Inches, Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn, nsdecls
from docx.oxml import parse_xml
import os

# ── Colors ──
BLUE_DARK = RGBColor(0x0C, 0x4A, 0x6E)  # heading
BLUE_MID = RGBColor(0x0E, 0x7A, 0xB5)   # subheading
BLUE_LIGHT = RGBColor(0xE0, 0xF2, 0xFE) # table header bg
GRAY_TEXT = RGBColor(0x37, 0x41, 0x51)   # body
GRAY_LIGHT = RGBColor(0x94, 0xA3, 0xB8)  # caption
GREEN = RGBColor(0x05, 0x96, 0x69)       # positive
RED = RGBColor(0xDC, 0x26, 0x26)         # negative
WHITE = RGBColor(0xFF, 0xFF, 0xFF)

OUTPUT_FILE = r"c:\Trường học số - source code\Phan tich thiet bi - THS.docx"


def set_cell_shading(cell, color_hex):
    """Set background color of a table cell."""
    shading = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{color_hex}" w:val="clear"/>')
    cell._tc.get_or_add_tcPr().append(shading)


def set_cell_border(cell, top=None, bottom=None, start=None, end=None):
    """Set borders on a cell."""
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcBorders = parse_xml(f'<w:tcBorders {nsdecls("w")}></w:tcBorders>')
    for side, val in [("top", top), ("bottom", bottom), ("start", start), ("end", end)]:
        if val:
            border = parse_xml(
                f'<w:{side} {nsdecls("w")} w:val="single" w:sz="{val}" w:space="0" w:color="CBD5E1"/>'
            )
            tcBorders.append(border)
    tcPr.append(tcBorders)


def add_styled_table(doc, headers, rows, col_widths=None):
    """Add a clean, styled table."""
    table = doc.add_table(rows=1 + len(rows), cols=len(headers))
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = True

    # Header row
    for i, h in enumerate(headers):
        cell = table.rows[0].cells[i]
        cell.text = ""
        p = cell.paragraphs[0]
        run = p.add_run(h)
        run.bold = True
        run.font.size = Pt(9)
        run.font.color.rgb = BLUE_DARK
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        set_cell_shading(cell, "E0F2FE")

    # Data rows
    for r_idx, row_data in enumerate(rows):
        for c_idx, cell_text in enumerate(row_data):
            cell = table.rows[r_idx + 1].cells[c_idx]
            cell.text = ""
            p = cell.paragraphs[0]
            run = p.add_run(str(cell_text))
            run.font.size = Pt(9)
            run.font.color.rgb = GRAY_TEXT
            if r_idx % 2 == 1:
                set_cell_shading(cell, "F8FAFC")

    # Set column widths if provided
    if col_widths:
        for row in table.rows:
            for i, w in enumerate(col_widths):
                row.cells[i].width = Cm(w)

    # Spacing after table
    doc.add_paragraph()
    return table


def add_heading_styled(doc, text, level=1):
    """Add a heading with custom styling."""
    h = doc.add_heading(text, level=level)
    for run in h.runs:
        run.font.color.rgb = BLUE_DARK if level <= 2 else BLUE_MID
    return h


def add_body(doc, text):
    """Add body paragraph."""
    p = doc.add_paragraph()
    run = p.add_run(text)
    run.font.size = Pt(10.5)
    run.font.color.rgb = GRAY_TEXT
    run.font.name = "Calibri"
    p.paragraph_format.space_after = Pt(6)
    p.paragraph_format.line_spacing = Pt(16)
    return p


def add_bullet(doc, text, bold_prefix=None):
    """Add a bullet point, optionally with a bold prefix."""
    p = doc.add_paragraph(style="List Bullet")
    if bold_prefix:
        run_b = p.add_run(bold_prefix)
        run_b.bold = True
        run_b.font.size = Pt(10)
        run_b.font.color.rgb = GRAY_TEXT
    run = p.add_run(text)
    run.font.size = Pt(10)
    run.font.color.rgb = GRAY_TEXT
    return p


def add_callout(doc, text):
    """Add a callout/tip box using a single-cell table with background."""
    table = doc.add_table(rows=1, cols=1)
    cell = table.rows[0].cells[0]
    cell.text = ""
    p = cell.paragraphs[0]
    run = p.add_run("💡 Lưu ý: ")
    run.bold = True
    run.font.size = Pt(9.5)
    run.font.color.rgb = BLUE_MID
    run2 = p.add_run(text)
    run2.font.size = Pt(9.5)
    run2.font.color.rgb = GRAY_TEXT
    set_cell_shading(cell, "FFF7ED")
    doc.add_paragraph()


def build_document():
    doc = Document()

    # ── Page setup ──
    section = doc.sections[0]
    section.page_width = Cm(21)
    section.page_height = Cm(29.7)
    section.top_margin = Cm(2)
    section.bottom_margin = Cm(2)
    section.left_margin = Cm(2.5)
    section.right_margin = Cm(2.5)

    # ── Default font ──
    style = doc.styles["Normal"]
    font = style.font
    font.name = "Calibri"
    font.size = Pt(10.5)
    font.color.rgb = GRAY_TEXT

    # ══════════════════════════════════════════
    # TITLE
    # ══════════════════════════════════════════
    title = doc.add_heading("Phân tích 2 thiết bị chạy hệ thống Trường học số", level=0)
    for run in title.runs:
        run.font.color.rgb = BLUE_DARK
        run.font.size = Pt(22)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER

    subtitle = doc.add_paragraph()
    subtitle.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_sub = subtitle.add_run("So sánh Joystar J16 và VASOUN L10 cho buổi demo THS")
    run_sub.font.size = Pt(12)
    run_sub.font.color.rgb = GRAY_LIGHT
    run_sub.font.italic = True

    doc.add_paragraph()  # spacing

    # ══════════════════════════════════════════
    # 1. TỔNG QUAN PHẦN CỨNG
    # ══════════════════════════════════════════
    add_heading_styled(doc, "1. Tổng quan phần cứng", level=1)

    add_body(doc,
        "Hai máy đều là tablet Android khoảng 10 inch, cùng chạy chip UNISOC T606, GPU Mali-G57, "
        "bộ nhớ trong eMMC 128GB. Nền tảng giống nhau, nhưng khác ở RAM, pin, và vài chi tiết nhỏ "
        "mà khi chạy hệ thống thực tế mới thấy ảnh hưởng."
    )

    add_body(doc,
        "AnTuTu v11.1.4 cho kết quả tổng gần bằng nhau: Máy 1 đạt 283.329 điểm, Máy 2 đạt 281.786 điểm. "
        "Chênh lệch không đáng kể ở mức tổng, nhưng khi nhìn vào từng thành phần thì bắt đầu thấy khác biệt."
    )

    # Bảng so sánh chính
    add_heading_styled(doc, "Bảng so sánh phần cứng", level=2)
    headers = ["Thông số", "Máy 1 (Joystar J16)", "Máy 2 (VASOUN L10)", "Ghi chú"]
    rows = [
        ["CPU", "UNISOC T606, 8 core\n614-1612 MHz", "UNISOC T606, 8 core\n614-1612 MHz", "Giống nhau"],
        ["GPU", "Mali-G57", "Mali-G57", "Giống nhau"],
        ["RAM", "4,1 GB\nLPDDR4x 1600MHz", "6,14 GB\nLPDDR4x 1600MHz", "Máy 2 nhiều hơn ~2GB"],
        ["Lưu trữ", "128GB eMMC", "128GB eMMC", "Giống nhau"],
        ["Màn hình", "10.1 inch, 800x1280\n60 Hz", "10.06 inch, 800x1280\n55 Hz", "Máy 1 mượt hơn chút"],
        ["Pin", "6.000 mAh", "8.000 mAh", "Máy 2 trâu hơn"],
        ["Camera sau", "13 MP", "8 MP", "Máy 1 rõ hơn"],
        ["Camera trước", "5 MP", "5 MP", "Giống nhau"],
        ["Cảm biến", "27", "23", "Máy 1 đầy đủ hơn"],
        ["Hệ điều hành", "Android (chưa rõ)", "Android 14 (64-bit)", "Máy 2 mới hơn"],
    ]
    add_styled_table(doc, headers, rows, col_widths=[3.5, 4, 4, 4.5])

    # Bảng benchmark
    add_heading_styled(doc, "Điểm benchmark AnTuTu v11.1.4", level=2)
    bench_headers = ["Hạng mục", "Máy 1", "Máy 2", "Chênh lệch"]
    bench_rows = [
        ["Tổng điểm", "283.329", "281.786", "Máy 1 nhỉnh ~0.5%"],
        ["CPU", "128.508", "123.685", "Máy 1 nhỉnh ~4%"],
        ["GPU", "0 (không hỗ trợ test)", "0 (không hỗ trợ test)", "Cả hai đều không chạy được GPU test"],
        ["MEM (RAM)", "73.324", "77.919", "Máy 2 nhỉnh ~6%"],
        ["UX", "81.497", "80.182", "Gần bằng nhau"],
        ["Nhiệt độ khi test", "32 - 36°C", "20°C", "Máy 2 mát hơn rất nhiều"],
    ]
    add_styled_table(doc, bench_headers, bench_rows, col_widths=[4, 3, 3, 6])

    # ══════════════════════════════════════════
    # 2. PHÂN TÍCH TỪNG PROJECT
    # ══════════════════════════════════════════
    add_heading_styled(doc, "2. Phân tích theo từng project", level=1)

    # ── 2a. gkebook ──
    add_heading_styled(doc, "2a. gkebook-school-frontend", level=2)
    add_body(doc,
        "Project gkebook chạy trên trình duyệt, xây bằng Next.js + React 18. Nội dung chủ yếu là "
        "form nhập liệu, bảng danh sách, và dashboard báo cáo dùng Recharts. Không có gì quá nặng về "
        "GPU hay animation phức tạp. Cái cần nhất là RAM đủ để giữ nhiều tab trình duyệt cùng lúc, "
        "nhất là khi tải dashboard có nhiều biểu đồ."
    )

    gkebook_headers = ["Thao tác", "Máy 1 (4GB RAM)", "Máy 2 (6GB RAM)"]
    gkebook_rows = [
        ["Đăng nhập, mở trang chủ", "Chạy bình thường", "Chạy bình thường"],
        ["Xem danh sách lớp,\ntạo nhiệm vụ", "Chạy bình thường", "Chạy bình thường"],
        ["Dashboard báo cáo\n(Recharts, nhiều biểu đồ)", "Có thể lag nếu dữ liệu nhiều", "Mượt hơn nhờ RAM rộng"],
        ["Mở nhiều tab cùng lúc", "Dễ bị reload tab\n(RAM 4GB hết chỗ)", "Giữ tab ổn hơn"],
        ["Xuất PDF báo cáo\n(jsPDF + html-to-image)", "Chậm, có nguy cơ crash tab", "Xử lý ổn hơn"],
    ]
    add_styled_table(doc, gkebook_headers, gkebook_rows, col_widths=[4.5, 5.5, 6])

    add_body(doc,
        "Cả hai máy đều chạy gkebook được. Nhưng nếu buổi demo có phần xuất PDF báo cáo hoặc mở "
        "nhiều tab so sánh dữ liệu, thì Máy 2 (6GB RAM) ổn định hơn. Máy 1 với 4GB RAM dễ bị trình "
        "duyệt tự kill tab khi cạn bộ nhớ, xảy ra giữa demo thì khá mất điểm."
    )

    # ── 2b. Smart Class ──
    add_heading_styled(doc, "2b. prj-ths-smart-class", level=2)
    add_body(doc,
        "Smart Class là project nặng nhất trong hệ thống. Ngoài giao diện web thông thường, nó "
        "còn chạy Three.js (hiển thị 3D), tldraw (bảng vẽ tương tác, fork riêng rất ngốn RAM), "
        "Framer Motion (animation), WaveSurfer.js (xử lý audio), WebRTC (kết nối lớp học thời gian thực), "
        "và Face Engine (nhận diện khuôn mặt chạy local). "
        "Tất cả chạy cùng lúc trên một tab trình duyệt, nên RAM là yếu tố quyết định."
    )

    sc_headers = ["Thao tác", "Máy 1 (4GB RAM)", "Máy 2 (6GB RAM)"]
    sc_rows = [
        ["Tải trang, splash screen", "Ổn", "Ổn"],
        ["Bảng vẽ tldraw cơ bản", "Lag khi vẽ nhiều objects", "Mượt hơn rõ rệt"],
        ["Render 3D\n(Three.js scenes)", "GPU giống nhau, nhưng\nRAM ít nên texture dễ bị drop", "Giữ texture tốt hơn"],
        ["Face recognition", "Chậm, RAM không đủ buffer", "Chạy được, vẫn hơi chậm"],
        ["WebRTC\n(kết nối lớp học)", "Kết nối ổn, nhưng dễ đứng\nnếu có stream video", "Ổn định hơn"],
        ["Audio (WaveSurfer)", "Ổn", "Ổn"],
        ["Chạy Tauri app\n(nếu build native)", "RAM chật, multitask kém", "Khả thi hơn"],
        ["Tản nhiệt khi dùng lâu", "Nóng (32 - 36°C khi bench),\ndễ bị throttle CPU", "Mát (20°C khi bench),\nkhông lo throttle"],
    ]
    add_styled_table(doc, sc_headers, sc_rows, col_widths=[4.5, 5.5, 6])

    add_body(doc,
        "Máy 2 (VASOUN L10) rõ ràng phù hợp hơn cho Smart Class. 6GB RAM là mức tối thiểu kha khá "
        "cho một ứng dụng chạy Three.js + tldraw + WebRTC cùng lúc. Máy 1 với 4GB sẽ hay gặp lag và "
        "crash khi demo kéo dài. Thêm nữa, Máy 2 tản nhiệt tốt hơn hẳn (20°C so với 36°C), nên "
        "chạy 60 phút liên tục mà không bị giảm hiệu năng."
    )

    # ══════════════════════════════════════════
    # 3. TỔNG KẾT ĐIỂM MẠNH / YẾU
    # ══════════════════════════════════════════
    add_heading_styled(doc, "3. Tổng kết điểm mạnh, điểm yếu", level=1)

    # ── Máy 1 ──
    add_heading_styled(doc, "Máy 1 (Joystar J16)", level=2)

    p = doc.add_paragraph()
    run = p.add_run("Điểm mạnh:")
    run.bold = True
    run.font.size = Pt(10.5)
    run.font.color.rgb = GREEN

    add_bullet(doc, "Màn hình 60Hz, mượt hơn Máy 2 một chút (dù 5Hz chênh lệch khó thấy bằng mắt)")
    add_bullet(doc, "Camera sau 13MP, chất lượng ảnh tốt hơn nếu demo có phần chụp ảnh")
    add_bullet(doc, "CPU benchmark nhỉnh hơn khoảng 4% (128.508 vs 123.685)")
    add_bullet(doc, "Nhiều cảm biến hơn (27 vs 23)")

    p2 = doc.add_paragraph()
    run2 = p2.add_run("Điểm yếu:")
    run2.bold = True
    run2.font.size = Pt(10.5)
    run2.font.color.rgb = RED

    add_bullet(doc,
        "RAM chỉ 4GB. Android chiếm khoảng 1-1.5GB, còn lại 2.5-3GB cho ứng dụng. "
        "Smart Class cần ít nhất 3-4GB riêng, nên gần như không còn headroom."
    )
    add_bullet(doc, "Nóng hơn khi chạy lâu (32-36°C), CPU dễ bị throttle xuống 614MHz")
    add_bullet(doc, "Pin 6.000mAh, chạy demo 60 phút không sạc thì hơi căng")

    # ── Máy 2 ──
    add_heading_styled(doc, "Máy 2 (VASOUN L10)", level=2)

    p3 = doc.add_paragraph()
    run3 = p3.add_run("Điểm mạnh:")
    run3.bold = True
    run3.font.size = Pt(10.5)
    run3.font.color.rgb = GREEN

    add_bullet(doc, "6GB RAM, thoải mái cho cả gkebook lẫn Smart Class")
    add_bullet(doc, "Tản nhiệt rất tốt (20°C khi benchmark nặng), chạy 60 phút không lo hiệu năng giảm")
    add_bullet(doc, "Pin 8.000mAh, dư sức cho buổi demo mà không cần cắm sạc")
    add_bullet(doc, "Android 14, WebView và driver mới hơn, tương thích tốt với các API web hiện đại")
    add_bullet(doc, "MEM score cao hơn 6% (77.919 vs 73.324), RAM nhanh hơn khi xử lý dữ liệu")

    p4 = doc.add_paragraph()
    run4 = p4.add_run("Điểm yếu:")
    run4.bold = True
    run4.font.size = Pt(10.5)
    run4.font.color.rgb = RED

    add_bullet(doc, "Màn hình 55Hz (thấp hơn 60Hz của Máy 1, nhưng bằng mắt thường khó phân biệt)")
    add_bullet(doc, "Camera sau chỉ 8MP (nếu demo có chụp ảnh thì kém hơn)")
    add_bullet(doc, "Ít cảm biến hơn (23 vs 27)")

    # ══════════════════════════════════════════
    # 4. KHUYẾN NGHỊ
    # ══════════════════════════════════════════
    add_heading_styled(doc, "4. Khuyến nghị phân vai cho buổi demo", level=1)

    add_body(doc,
        "Dựa trên phân tích ở trên, phương án phân vai hợp lý nhất cho buổi demo 60 phút như sau:"
    )

    rec_headers = ["Vai trò trong demo", "Thiết bị nên dùng", "Lý do"]
    rec_rows = [
        ["Máy giáo viên\n(chạy Smart Class)", "Máy 2 (VASOUN L10)", "RAM 6GB chạy tldraw + Three.js mượt,\npin 8.000mAh, không lo nóng"],
        ["Máy học sinh\n(chạy gkebook)", "Máy 1 (Joystar J16)", "Chỉ cần trình duyệt web,\n4GB RAM đủ dùng"],
        ["Demo xuất PDF\nbáo cáo phụ huynh", "Máy 2 (VASOUN L10)", "Xuất PDF cần RAM,\nMáy 2 ổn định hơn"],
        ["Máy dự phòng", "Máy 1 (Joystar J16)", "Nếu Máy 2 gặp sự cố,\nMáy 1 vẫn chạy được (chậm hơn)"],
    ]
    add_styled_table(doc, rec_headers, rec_rows, col_widths=[4.5, 4.5, 7])

    add_callout(doc,
        "Trước buổi demo, nên clear RAM trên cả hai máy (tắt hết app chạy nền), bật Developer Options "
        "và giữ màn hình luôn sáng. Với Máy 1, không mở thêm app nào ngoài trình duyệt khi đang demo."
    )

    # ══════════════════════════════════════════
    # FOOTER
    # ══════════════════════════════════════════
    footer_p = doc.add_paragraph()
    footer_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_f = footer_p.add_run("Tài liệu nội bộ - Trường học số")
    run_f.font.size = Pt(9)
    run_f.font.color.rgb = GRAY_LIGHT
    run_f.italic = True

    # ── Save ──
    doc.save(OUTPUT_FILE)
    print("Exported docx successfully.")


if __name__ == "__main__":
    build_document()
