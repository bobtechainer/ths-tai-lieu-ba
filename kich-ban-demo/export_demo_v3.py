# -*- coding: utf-8 -*-
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from copy import copy

src = openpyxl.load_workbook(r'c:\Trường học số - source code\Kịch bản Demo THS - Mobifone Solution.xlsx')
wb = openpyxl.Workbook()

def copy_sheet(src_ws, dst_ws):
    for row in src_ws.iter_rows():
        for cell in row:
            nc = dst_ws.cell(row=cell.row, column=cell.column, value=cell.value)
            if cell.has_style:
                nc.font = copy(cell.font); nc.fill = copy(cell.fill)
                nc.alignment = copy(cell.alignment); nc.border = copy(cell.border)
                nc.number_format = cell.number_format
    for mc in src_ws.merged_cells.ranges:
        dst_ws.merge_cells(str(mc))
    for k, d in src_ws.column_dimensions.items():
        dst_ws.column_dimensions[k].width = d.width
    for k, d in src_ws.row_dimensions.items():
        dst_ws.row_dimensions[k].height = d.height

# Copy ALL 4 original sheets
for i, name in enumerate(src.sheetnames):
    if i == 0:
        ws = wb.active
        ws.title = name
    else:
        ws = wb.create_sheet(title=name)
    copy_sheet(src[name], ws)
    print(f"Copied: [{name}]")

# === NEW SHEET: BTVN + Bao cao ===
WHITE = "FFFFFF"; GRAY = "F2F2F2"; DARK = "BF6B00"; LIGHT = "FFF2CC"
thin_b = Border(left=Side("thin","BFBFBF"), right=Side("thin","BFBFBF"),
                top=Side("thin","BFBFBF"), bottom=Side("thin","BFBFBF"))
wrap_top = Alignment(wrap_text=True, vertical="top")
wrap_ctr = Alignment(wrap_text=True, vertical="center", horizontal="center")
HEADERS = ["STT", "Vai trò", "Cửa sổ thao tác", "Tính năng trải nghiệm",
           "Thao tác chi tiết", "Kết quả mong đợi", "Data demo gợi ý"]
COL_W = [5, 14, 14, 26, 50, 30, 30]

ROWS = [
    [1, "Giáo viên", "Giáo viên", "Gv tạo nhiệm vụ (bài tập về nhà) cho lớp",
     "B1: Vào lớp học --> tab \"Nhiệm vụ\" --> bấm \"Tạo nhiệm vụ\"\n"
     "B2: Bước 1/3, điền thông tin chung:\n"
     "- Nhập tiêu đề nhiệm vụ\n"
     "- Chọn lớp nhận bài (giao 1 hoặc nhiều lớp cùng lúc)\n"
     "- Chọn hạn nộp, tick \"Cho phép nộp muộn\" nếu cần\n"
     "B3: Bước 2/3, tạo nội dung bài tập:\n"
     "- Gv tự tạo quiz trắc nghiệm: bấm \"Thêm câu hỏi\" --> chọn dạng câu hỏi\n"
     "- Nhập đề, nhập đáp án, chọn đáp án đúng\n"
     "- Hoặc giao dạng tự luận cho hs upload file bài làm\n"
     "B4: Bước 3/3, xem lại --> bấm \"Giao nhiệm vụ\"",
     "Nhiệm vụ xuất hiện trong danh sách lớp.\nHS nhận thông báo bài tập mới.",
     "Lớp: Hóa học 10A1\nTiêu đề: \"Ôn tập Chương 6 - Tốc độ phản ứng\"\n"
     "Tạo 5 câu trắc nghiệm chọn 1 đáp án\nHạn nộp: 23:59 ngày hôm sau"],

    [2, "Học sinh", "Học sinh", "Hs nhận bài tập và làm bài trên hệ thống",
     "B1: Hs vào mục \"Nhiệm vụ\" hoặc vào lớp --> tab Nhiệm vụ\n"
     "B2: Thấy danh sách nhiệm vụ được giao kèm deadline\n"
     "B3: Bấm vào nhiệm vụ --> bấm \"Làm bài\"\n"
     "B4: Trả lời lần lượt các câu trắc nghiệm, hệ thống hiện thanh tiến độ\n"
     "B5: Bấm \"Nộp bài\"\n"
     "- Với quiz: hệ thống tự chấm ngay, hs biết kết quả luôn\n"
     "- Với tự luận: chờ gv chấm",
     "Hs thấy rõ deadline, tiến độ nộp bài.\nQuiz tự chấm ngay sau khi nộp, hs xem được đáp án đúng/sai từng câu.",
     "Hs demo: Nguyễn Văn A\nTrả lời 5 câu, đúng 4/5 = 8 điểm"],

    [3, "Giáo viên", "Giáo viên", "Gv xem tổng quan bài nộp và chấm điểm",
     "B1: Gv vào chi tiết nhiệm vụ đã giao\n"
     "B2: Hệ thống hiện Dashboard chấm điểm:\n"
     "- Tổng quan: số hs đã nộp / chưa nộp, tỷ lệ hoàn thành, độ chính xác trung bình\n"
     "- Phân bố điểm 4 mức (9-10, 7-8, 5-6, dưới 5)\n"
     "- Hs điểm cao nhất / thấp nhất\n"
     "B3: Tab \"Học sinh\": bảng danh sách từng hs, trạng thái, điểm, số câu đúng/sai\n"
     "B4: Click vào 1 hs --> trang chấm chi tiết:\n"
     "- Xem bài làm từng câu\n"
     "- Gv viết nhận xét (AI Tag gợi ý nhận xét tự động)\n"
     "- Bấm \"Hiển thị nhận xét\" để hs thấy đánh giá\n"
     "B5: Tab \"Danh sách câu hỏi\": % đúng từng câu, phân bố đáp án",
     "Gv có dashboard chấm điểm rõ ràng.\nPhân bố điểm 4 mức + top/bottom hs.\n"
     "AI Tag gợi ý nhận xét, tiết kiệm thời gian.\nPhân tích từng câu giúp gv biết hs yếu ở đâu.",
     "28/35 hs đã nộp = 80%\nĐộ chính xác TB: 72%\n"
     "Mức 9-10: 5 hs, 7-8: 12 hs\n5-6: 8 hs, dưới 5: 3 hs\n"
     "Top: Nguyễn Văn A (10đ)\nBottom: Trần Văn B (3đ)"],

    [4, "Học sinh", "Học sinh", "Hs xem kết quả, nhận xét từ gv và báo cáo cá nhân",
     "B1: Hs vào chi tiết nhiệm vụ đã nộp:\n"
     "- Xem điểm số, nhận xét của gv\n"
     "- Xem chi tiết từng câu đúng/sai, so sánh đáp án\n"
     "B2: Hs mở báo cáo cá nhân sau buổi học:\n"
     "- Điểm số, thứ hạng trong lớp, tỷ lệ trả lời đúng\n"
     "- Tab \"Nhận xét\": AI đánh giá điểm mạnh, điểm yếu, gợi ý cải thiện\n"
     "- Tab \"So sánh\": radar chart so với trung bình lớp\n"
     "- Tab \"Mức độ tập trung\": biểu đồ engagement theo timeline\n"
     "B3: Xem Dashboard thống kê tuần: số buổi tham gia, giờ học, tỷ lệ chuyên cần",
     "Hs nắm được mình đúng/sai ở đâu.\nAI gợi ý cá nhân hóa.\n"
     "Hs tự theo dõi tần suất học tập qua stat tuần.",
     "Điểm: 8/10, đúng 4/5 câu\nAI nhận xét: \"Trả lời tốt phần tốc độ phản ứng, cần ôn thêm phần xúc tác\"\n"
     "Stat tuần: 8/10 buổi, 6.5h, chuyên cần 80%"],

    [5, "Giáo viên", "Giáo viên", "Gv xem báo cáo lớp và so sánh giữa các lớp",
     "B1: Vào Dashboard Smart Class, xem stat tuần: số tiết dạy, tổng giờ giảng, tổng hs, tỷ lệ chuyên cần\n"
     "B2: Xem báo cáo buổi học chi tiết:\n"
     "- Biểu đồ mức độ tập trung hs theo timeline\n"
     "- Kết quả quiz trong buổi: % đúng từng câu\n"
     "- BXH hs theo accuracy + tốc độ\n"
     "B3: So sánh giữa các lớp cùng bài giảng",
     "Gv biết lúc nào hs mất tập trung, câu nào hs yếu.\nSo sánh được chất lượng giữa các lớp.",
     "12 tiết/tuần, 9.5 giờ giảng\n140 hs, chuyên cần 92%\nSo sánh 10A1 vs 10A2"],

    [6, "GV --> PHHS", "Giáo viên", "Gv gửi phiếu báo cáo học tập cho phụ huynh qua email",
     "B1: Gv vào lớp --> tab \"Báo cáo học sinh\"\n"
     "B2: Chọn hs cần gửi, chọn tuần báo cáo\n"
     "B3: Hệ thống hiện phiếu báo cáo gồm:\n"
     "- Biểu đồ chuyên cần (đi học / vắng)\n"
     "- Phân bố điểm nhiệm vụ\n"
     "- Tỷ lệ hoàn thành + danh sách nhiệm vụ chi tiết\n"
     "B4: Bấm \"Gửi cho phụ huynh\" --> form soạn email có tiêu đề + nội dung mẫu sẵn, gv sửa nhận xét nếu cần\n"
     "B5: Bấm \"Gửi ngay\" --> hệ thống tự xuất PDF, gửi email đến parent_emails\n"
     "Lưu ý: email phụ huynh cần nhập sẵn trong profile hs",
     "Phụ huynh nhận email kèm file PDF phiếu báo cáo.\nGv tiết kiệm thời gian nhờ form mẫu + PDF tự động.",
     "Hs: Nguyễn Văn A, lớp 10A1\nTuần: tuần 4 tháng 7\nChuyên cần: 4/5 buổi\n"
     "3 nhiệm vụ, điểm: 8, 9, 7\nEmail PHHS: phuhuynhdemo@gmail.com"],
]

ws_new = wb.create_sheet(title="P3 SAU GIỜ HỌC - BTVN & BÁO CÁO")
ws_new.sheet_properties.tabColor = DARK
for i, w in enumerate(COL_W, 1):
    ws_new.column_dimensions[get_column_letter(i)].width = w

ws_new.merge_cells("A1:G1")
c = ws_new.cell(1, 1, "Sau giờ học: Giao BTVN, chấm điểm, báo cáo cho GV / HS / Phụ huynh")
c.font = Font("Arial", 14, bold=True, color=WHITE)
c.fill = PatternFill("solid", fgColor=DARK)
c.alignment = Alignment(vertical="center")
ws_new.row_dimensions[1].height = 35

ws_new.merge_cells("A2:G2")
c = ws_new.cell(2, 1,
    "Mục tiêu: Gv giao BTVN sau buổi học, hs làm bài và nộp trên hệ thống, gv chấm điểm. "
    "Kết quả được phản hồi cho cả gv, hs và phụ huynh. Phụ huynh nhận email báo cáo kèm PDF phiếu học tập.")
c.font = Font("Arial", 11, italic=True, color="333333")
c.fill = PatternFill("solid", fgColor=LIGHT)
c.alignment = Alignment(wrap_text=True, vertical="center")
ws_new.row_dimensions[2].height = 42

hf = PatternFill("solid", fgColor=DARK)
for j, h in enumerate(HEADERS, 1):
    c = ws_new.cell(3, j, h)
    c.font = Font("Arial", 11, bold=True, color=WHITE)
    c.fill = hf; c.alignment = wrap_ctr; c.border = thin_b
ws_new.row_dimensions[3].height = 28

for i, rd in enumerate(ROWS):
    r = 4 + i
    stripe = PatternFill("solid", fgColor=GRAY if i % 2 else WHITE)
    for j, val in enumerate(rd, 1):
        c = ws_new.cell(r, j, val)
        c.font = Font("Arial", 10, bold=(j == 4))
        c.alignment = wrap_ctr if j <= 3 else wrap_top
        c.fill = stripe; c.border = thin_b
    lines = max(str(v).count("\n") + 1 for v in rd if v)
    ws_new.row_dimensions[r].height = max(20, min(lines * 15, 250))
ws_new.freeze_panes = "A4"

# Move new sheet after P2 (index 2) and before Admin (index 3)
# Current order: [Tài khoản, P1, P2, Admin, P3_NEW]
# Want: [Tài khoản, P1, P2, P3_NEW, Admin]
wb.move_sheet("P3 SAU GIỜ HỌC - BTVN & BÁO CÁO", offset=-1)

out = r'c:\Trường học số - source code\Kịch bản Demo THS - Nâng cao v3.xlsx'
wb.save(out)
print("OK: " + out)
print(f"Final sheets: {wb.sheetnames}")
