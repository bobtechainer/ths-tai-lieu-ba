# -*- coding: utf-8 -*-
"""Sheet Cách đọc file, chèn lên đầu workbook."""
from openpyxl.styles import Font, PatternFill, Alignment

TIEU_DE = Font(bold=True, size=12, color="1F3864")
THUONG = Alignment(vertical="top", wrap_text=True)


def them(wb, so_dat, so_khong, so_chua, tong_truong, so_co, so_dang, so_khong_tr):
    ws = wb.create_sheet("Cách đọc file", 0)
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 34
    ws.column_dimensions["B"].width = 108

    khoi = [
        ("Việc đã làm", None),
        ("Mục tiêu",
         "Kiểm xem trang quản lý trường học đã có đủ các thông tin mà CSDL ngành Đồng Nai "
         "yêu cầu khi đẩy dữ liệu lên hay chưa."),
        ("Phạm vi",
         "37 đầu mục do chị Thư phụ trách, tức số 39 đến 75 trong mục Tiếp nhận xử lý của "
         "checklist. Không xét các đầu mục của bạn Bách."),
        ("Tài liệu yêu cầu",
         "Tài liệu API tiếp nhận kết nối CSDL Đồng Nai ngày 12/08/2026. Với mỗi đầu mục, "
         "lấy đúng phần Request body, tức danh sách thông tin bắt buộc phải gửi lên."),
        ("Hệ thống được kiểm",
         "uat.truonghocsovietnam.vn, đăng nhập bằng tài khoản của Trường THCS Ngô Tất Tố. "
         "Đã mở từng màn hình và từng ô nhập để đối chiếu."),
        ("Khối lượng đã đối chiếu",
         "%d nội dung của cả 37 đầu mục, không gộp và không suy diễn giữa các cấp học. "
         "Mỗi cấp học được đọc riêng vì tài liệu yêu cầu khác nhau theo cấp." % tong_truong),
        ("", None),
        ("Kết quả", None),
        ("37 đầu mục",
         "Đạt %d. Không đạt %d. Chưa kết luận %d." % (so_dat, so_khong, so_chua)),
        ("%d nội dung chi tiết" % tong_truong,
         "Hệ thống đã có %d. Đã có nhưng lưu ở dạng khác %d. Chưa có %d."
         % (so_co, so_dang, so_khong_tr)),
        ("", None),
        ("Hai sheet trong file", None),
        ("Sheet Checklist",
         "Một dòng cho một đầu mục. Cột Kết luận và cột Ghi chú tóm tắt lý do."),
        ("Sheet Trace chi tiết",
         "Một dòng cho một nội dung mà ngành yêu cầu. Lọc theo cột STT để xem đúng đầu mục "
         "đang cần. Đây là phần để trả lời câu hỏi thiếu cụ thể cái gì."),
        ("", None),
        ("Cách hiểu các ô màu", None),
        ("Đạt (xanh)",
         "Hệ thống có đủ mọi nội dung mà đầu mục đó yêu cầu."),
        ("Không đạt (đỏ)",
         "Còn ít nhất một nội dung hệ thống chưa có chỗ để nhập."),
        ("Chưa kết luận (vàng)",
         "Chưa kiểm được vì chưa có tài khoản trường phù hợp."),
        ("Có",
         "Hệ thống có ô nhập tương ứng. Cột bên cạnh ghi rõ nằm ở màn hình nào."),
        ("Có, ở màn khác",
         "Thông tin có nhưng nằm ở màn hình khác, không nằm cùng chỗ với đầu mục. "
         "Khi làm chức năng đẩy dữ liệu thì phải lấy từ nhiều màn ghép lại."),
        ("Có, chưa đúng dạng",
         "Thông tin có nhưng lưu khác cách ngành yêu cầu. Ví dụ ngành cần chọn tỉnh, "
         "huyện, xã theo danh mục nhưng hệ thống chỉ có một ô để gõ chữ. "
         "Trường hợp này vẫn tính là có, nhưng khi đẩy lên phải quy đổi lại."),
        ("Không",
         "Hệ thống chưa có chỗ nào để nhập nội dung này."),
        ("", None),
        ("Điểm cần lưu ý khi đọc", None),
        ("Về cấp học",
         "Toàn bộ việc kiểm được làm trên tài khoản một trường THCS. Các màn hình đều dùng "
         "chung cho nhiều cấp và có cột Cấp học, nên kết luận được áp cho các cấp còn lại. "
         "Nếu cần chắc chắn tuyệt đối thì nên kiểm thêm bằng tài khoản trường mầm non, "
         "tiểu học, THPT và GDTX."),
        ("Về hai đầu mục còn treo",
         "Số 67 và 71 là bản dành cho GDTX. Hệ thống chỉ có màn cho tiểu học và cho "
         "THCS/THPT, chưa thấy màn riêng cho GDTX nên để Chưa kết luận."),
        ("Về mã định danh Bộ GD&ĐT",
         "Hồ sơ nhân sự có ô nhập mã này. Hồ sơ học sinh thì trang chi tiết có hiển thị "
         "nhưng form chỉnh sửa lại không có ô để nhập. Đây là điểm nên xử lý sớm vì thiếu "
         "mã này thì không ghép được dữ liệu học sinh với CSDL ngành."),
    ]

    r = 1
    for a, b in khoi:
        if b is None:
            ws.cell(r, 1, a).font = TIEU_DE
            ws.row_dimensions[r].height = 24 if a else 8
        else:
            ws.cell(r, 1, a).font = Font(bold=True)
            ws.cell(r, 1).alignment = THUONG
            ws.cell(r, 2, b).alignment = THUONG
            ws.row_dimensions[r].height = max(16, 15 * (len(b) // 95 + 1))
        r += 1
    ws.cell(1, 1).fill = PatternFill("solid", fgColor="FFFFFF")
    return ws
