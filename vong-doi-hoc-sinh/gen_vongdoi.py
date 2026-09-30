# -*- coding: utf-8 -*-
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
import os

WB_PATH = os.path.join(os.path.dirname(__file__), "vong-doi-hoc-sinh-THS.xlsx")
wb = openpyxl.Workbook()

# === STYLES ===
hdr_font = Font(name="Arial", bold=True, size=11, color="FFFFFF")
hdr_fill = PatternFill(start_color="2E5A88", end_color="2E5A88", fill_type="solid")
sec_font = Font(name="Arial", bold=True, size=11, color="1A3A5C")
sec_fill = PatternFill(start_color="D6E4F0", end_color="D6E4F0", fill_type="solid")
grn_fill = PatternFill(start_color="4CAF50", end_color="4CAF50", fill_type="solid")
note_fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
body_font = Font(name="Arial", size=10)
wrap = Alignment(wrap_text=True, vertical="top")
thin = Border(left=Side("thin"), right=Side("thin"), top=Side("thin"), bottom=Side("thin"))

def hdr(ws, row, n, fill=None):
    for c in range(1, n+1):
        cl = ws.cell(row=row, column=c)
        cl.font = hdr_font; cl.fill = fill or hdr_fill; cl.alignment = wrap; cl.border = thin

def bdy(ws, row, n):
    for c in range(1, n+1):
        cl = ws.cell(row=row, column=c)
        cl.font = body_font; cl.alignment = wrap; cl.border = thin

def tbl(ws, heads, data, sr=1, fill=None):
    for i, h in enumerate(heads, 1):
        ws.cell(row=sr, column=i, value=h)
    hdr(ws, sr, len(heads), fill)
    for r, rd in enumerate(data, sr+1):
        for c, v in enumerate(rd, 1):
            ws.cell(row=r, column=c, value=v)
        bdy(ws, r, len(heads))
    return sr + len(data) + 1

def wcol(ws, widths):
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[chr(64+i)].width = w

# ==========================================
# SHEET 1: TỔNG QUAN VÒNG ĐỜI HỌC SINH
# ==========================================
ws1 = wb.active
ws1.title = "1. Tổng quan"
wcol(ws1, [5, 24, 55, 22, 20, 38])

tbl(ws1,
    ["STT", "Nghiệp vụ", "Mô tả", "Ai thực hiện", "Thời điểm", "Ghi chú"],
    [
        ["1", "Nhập học",
         "Học sinh mới vào trường lần đầu. Trường ghi nhận thông tin cá nhân, phụ huynh, rồi gán vào lớp và năm học.",
         "Phụ huynh + Văn phòng", "Đầu năm / Giữa năm",
         "Cần nhập đủ: họ tên, ngày sinh, giới tính, phụ huynh, khối lớp."],
        ["2", "Xếp lớp",
         "Phân học sinh vào lớp chủ nhiệm cụ thể. Có thể xếp thủ công hoặc tự động theo tiêu chí (sĩ số, giới tính, học lực).",
         "Hiệu trưởng / Văn phòng", "Đầu năm học",
         "Hiện tại THS chỉ cho thêm từng học sinh vào lớp, chưa xếp hàng loạt."],
        ["3", "Lên lớp",
         "Cuối năm, học sinh đạt yêu cầu sẽ được chuyển lên khối trên ở năm học mới. Ví dụ: học sinh lớp 6A năm 2024-2025 lên lớp 7 năm 2025-2026.",
         "Giáo viên + Hiệu trưởng", "Cuối năm học",
         "Tiểu học đánh giá theo năng lực (TT27/2020). THCS/THPT tính điểm trung bình + hạnh kiểm."],
        ["4", "Ở lại lớp (đúp)",
         "Học sinh chưa đạt yêu cầu phải học lại cùng khối ở năm mới. Ví dụ: em ở lớp 7 mà chưa đạt, năm sau vẫn học lớp 7.",
         "Giáo viên + Hiệu trưởng", "Cuối năm học",
         "Tiểu học hiện nay hạn chế cho đúp (chỉ khi phụ huynh đồng ý). THCS/THPT: thi lại rồi mới xét."],
        ["5", "Chuyển lớp",
         "Chuyển học sinh sang lớp khác trong cùng trường, cùng khối. Ví dụ: chuyển từ 8A sang 8B.",
         "Văn phòng / Hiệu trưởng", "Bất kỳ lúc nào",
         "Thường do cân bằng sĩ số, nguyện vọng phụ huynh, hoặc lý do đặc biệt."],
        ["6", "Chuyển trường (đi)",
         "Học sinh rời trường hiện tại để sang trường khác. Trường xuất hồ sơ, cấp giấy chuyển trường.",
         "Phụ huynh + Văn phòng", "Bất kỳ lúc nào",
         "Cần cập nhật trạng thái trên CSDL ngành. Học sinh vẫn còn trong hệ thống nhưng đánh dấu 'đã chuyển'."],
        ["7", "Chuyển trường (đến)",
         "Học sinh từ trường khác chuyển đến nhập học. Trường kiểm tra hồ sơ, xếp lớp.",
         "Phụ huynh + Văn phòng", "Bất kỳ lúc nào",
         "Cần phân biệt với nhập học mới: học sinh chuyển đến đã có hồ sơ ở trường cũ."],
        ["8", "Thôi học",
         "Học sinh ngừng học vĩnh viễn vì hoàn cảnh, kỷ luật, quá tuổi, hoặc lý do khác.",
         "Phụ huynh / Trường", "Bất kỳ lúc nào",
         "THCS là giáo dục bắt buộc nên trường phải vận động học sinh quay lại trước khi cho thôi học."],
        ["9", "Bảo lưu",
         "Tạm ngừng học (bệnh, hoàn cảnh gia đình), giữ quyền quay lại học tiếp khi hết thời hạn.",
         "Phụ huynh + Văn phòng", "Bất kỳ lúc nào",
         "Phổ biến ở THPT. Tiểu học và THCS ít gặp. Thường bảo lưu tối đa 1 năm."],
        ["10", "Tốt nghiệp",
         "Hoàn thành chương trình cấp học. Lớp 5 và lớp 9: xét tốt nghiệp. Lớp 12: thi tốt nghiệp THPT quốc gia.",
         "Trường + Sở/Phòng GD", "Cuối năm cuối cấp",
         "Sau tốt nghiệp, học sinh không còn xuất hiện ở năm học mới."],
        ["11", "Tái nhập học",
         "Học sinh đã thôi học hoặc hết thời hạn bảo lưu quay lại học tiếp.",
         "Phụ huynh + Văn phòng", "Đầu năm / Giữa năm",
         "Cần xác định em học tiếp từ khối nào, xếp vào lớp nào."],
    ])

# ==========================================
# SHEET 2: QUY TRÌNH THỰC TẾ TẠI VIỆT NAM
# ==========================================
ws2 = wb.create_sheet("2. Quy trình thực tế VN")
wcol(ws2, [5, 22, 16, 55, 32, 28])

tbl(ws2,
    ["STT", "Nghiệp vụ", "Cấp học", "Quy trình thực tế", "Giấy tờ cần thiết", "Ai duyệt"],
    [
        ["1", "Lên lớp", "Tiểu học",
         "Giáo viên đánh giá học sinh cuối năm theo 3 mức: Hoàn thành tốt, Hoàn thành, Chưa hoàn thành (theo Thông tư 27/2020). Giáo viên đề xuất danh sách lên lớp. Hiệu trưởng ký quyết định. Không có thi cuối kỳ theo kiểu chấm điểm ở tiểu học.",
         "Phiếu đánh giá cuối năm, Học bạ", "GVCN + Hiệu trưởng"],
        ["2", "Lên lớp", "THCS",
         "Dựa trên điểm trung bình các môn và xếp loại hạnh kiểm (Thông tư 22/2021). Học sinh đạt thì lên lớp. Nếu chưa đạt 1-2 môn, được thi lại trong hè. Thi lại đạt thì vẫn lên lớp.",
         "Bảng điểm, Học bạ, Biên bản họp xét", "GVCN + Hội đồng sư phạm"],
        ["3", "Lên lớp", "THPT",
         "Tương tự THCS. Điểm trung bình các môn từ 5.0 trở lên, không môn nào dưới 3.5. Hạnh kiểm từ Trung bình trở lên. Được thi lại 1 lần trong hè.",
         "Bảng điểm, Học bạ, Biên bản họp xét", "GVCN + Hội đồng sư phạm"],
        ["4", "Ở lại lớp", "Tiểu học",
         "Theo Thông tư 27/2020, chỉ cho ở lại khi học sinh chưa hoàn thành và phụ huynh đồng ý bằng văn bản. Xu hướng hiện tại là hạn chế tối đa việc cho đúp ở tiểu học, ưu tiên hỗ trợ thêm.",
         "Phiếu đánh giá, Đơn đồng ý của phụ huynh", "Hiệu trưởng"],
        ["5", "Ở lại lớp", "THCS, THPT",
         "Không đạt điều kiện lên lớp sau khi đã thi lại. Hoặc nghỉ học quá 45 buổi/năm không phép.",
         "Bảng điểm, Biên bản xét", "Hội đồng sư phạm"],
        ["6", "Chuyển trường", "Tất cả cấp",
         "Bước 1: Phụ huynh nộp đơn xin chuyển ở trường cũ. Bước 2: Trường cũ cấp giấy chuyển trường kèm hồ sơ (học bạ, giấy khai sinh). Bước 3: Phụ huynh mang hồ sơ đến trường mới. Bước 4: Trường mới xét nhận và xếp lớp. Bước 5: Cả hai trường cập nhật trên CSDL ngành.",
         "Đơn xin chuyển, Học bạ, Giấy chuyển trường, Giấy khai sinh (bản sao)", "Hiệu trưởng 2 trường"],
        ["7", "Thôi học", "THCS, THPT",
         "Phụ huynh làm đơn (hoặc trường lập biên bản nếu kỷ luật). Trường ra quyết định. Cập nhật trạng thái trên CSDL ngành. Với THCS (giáo dục bắt buộc), trường phải vận động học sinh quay lại trước.",
         "Đơn xin thôi học, Quyết định của trường", "Hiệu trưởng"],
        ["8", "Bảo lưu", "THPT",
         "Phụ huynh nộp đơn kèm lý do (sức khỏe, hoàn cảnh). Trường xét duyệt. Học sinh giữ quyền quay lại trong thời hạn (thường 1 năm). Khi quay lại, xếp vào lớp phù hợp với khối đang dở.",
         "Đơn xin bảo lưu, Giấy xác nhận lý do (bệnh viện, UBND...)", "Hiệu trưởng"],
        ["9", "Tốt nghiệp tiểu học", "Tiểu học",
         "Xét tốt nghiệp, không thi. Học sinh hoàn thành chương trình lớp 5 thì được công nhận hoàn thành chương trình tiểu học.",
         "Học bạ, Phiếu đánh giá lớp 5", "Hiệu trưởng"],
        ["10", "Tốt nghiệp THCS", "THCS",
         "Xét tốt nghiệp dựa trên điểm tổng kết lớp 9 và quá trình học tập. Không có kỳ thi riêng (bỏ từ năm 2006).",
         "Học bạ, Bảng điểm tổng kết", "Hội đồng xét tốt nghiệp"],
        ["11", "Tốt nghiệp THPT", "THPT",
         "Thi tốt nghiệp THPT quốc gia. Kết quả thi kết hợp với điểm học bạ để xét. Sở GD&ĐT cấp bằng.",
         "Học bạ, Phiếu đăng ký thi, Kết quả thi", "Sở GD&ĐT"],
    ])

wb.save(WB_PATH)
print("Sheet 1-2 OK")
