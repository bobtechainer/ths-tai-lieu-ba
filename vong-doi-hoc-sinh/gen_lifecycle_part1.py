import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
import json, os

WB_PATH = os.path.join(os.path.dirname(__file__), "vong-doi-hoc-sinh-THS.xlsx")
wb = openpyxl.Workbook()

# === STYLES ===
hdr_font = Font(name="Arial", bold=True, size=12, color="FFFFFF")
hdr_fill = PatternFill(start_color="2E5A88", end_color="2E5A88", fill_type="solid")
sub_font = Font(name="Arial", bold=True, size=11, color="1A3A5C")
sub_fill = PatternFill(start_color="D6E4F0", end_color="D6E4F0", fill_type="solid")
body_font = Font(name="Arial", size=10)
wrap = Alignment(wrap_text=True, vertical="top")
thin_border = Border(
    left=Side(style="thin"), right=Side(style="thin"),
    top=Side(style="thin"), bottom=Side(style="thin")
)
note_fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")

def style_header(ws, row, cols):
    for c in range(1, cols+1):
        cell = ws.cell(row=row, column=c)
        cell.font = hdr_font; cell.fill = hdr_fill; cell.alignment = wrap; cell.border = thin_border

def style_row(ws, row, cols, fill=None):
    for c in range(1, cols+1):
        cell = ws.cell(row=row, column=c)
        cell.font = body_font; cell.alignment = wrap; cell.border = thin_border
        if fill: cell.fill = fill

def add_table(ws, headers, data, start_row=1):
    for i, h in enumerate(headers, 1):
        ws.cell(row=start_row, column=i, value=h)
    style_header(ws, start_row, len(headers))
    for r, row_data in enumerate(data, start_row+1):
        for c, val in enumerate(row_data, 1):
            ws.cell(row=r, column=c, value=val)
        style_row(ws, r, len(headers))
    return start_row + len(data) + 1

# =====================
# SHEET 1: TONG QUAN
# =====================
ws1 = wb.active
ws1.title = "1. Tong quan"
ws1.column_dimensions["A"].width = 5
ws1.column_dimensions["B"].width = 22
ws1.column_dimensions["C"].width = 50
ws1.column_dimensions["D"].width = 20
ws1.column_dimensions["E"].width = 20
ws1.column_dimensions["F"].width = 35

overview_headers = ["STT", "Nghiep vu", "Mo ta ngan", "Ai thuc hien", "Thoi diem", "Luu y"]
overview_data = [
    ["1", "Nhap hoc", "Hoc sinh moi vao truong lan dau, duoc gan vao lop va nam hoc cu the.", "Phu huynh + Van phong", "Dau nam hoc / Giua nam", "Can nhap du thong tin ca nhan, phu huynh, khoi lop."],
    ["2", "Xep lop", "Phan hoc sinh vao lop cu the (lop chu nhiem) trong nam hoc.", "Hieu truong / Van phong", "Dau nam hoc", "Co the xep thu cong hoac tu dong theo tieu chi."],
    ["3", "Len lop", "Hoc sinh dat yeu cau cuoi nam, duoc chuyen len khoi tren trong nam hoc moi.", "Giao vien + Hieu truong", "Cuoi nam hoc", "Tieu hoc: danh gia nang luc. THCS/THPT: diem trung binh + hanh kiem."],
    ["4", "O lai lop (Dup)", "Hoc sinh chua dat yeu cau, hoc lai cung khoi lop o nam hoc moi.", "Giao vien + Hieu truong", "Cuoi nam hoc", "Tieu hoc hien tai giam toi da viec cho dup (TT30/2014 -> TT27/2020)."],
    ["5", "Chuyen lop", "Chuyen hoc sinh sang lop khac trong cung truong, cung khoi.", "Van phong / Hieu truong", "Bat ky luc nao", "Thuong xay ra khi can can bang si so hoac theo nguyen vong."],
    ["6", "Chuyen truong (di)", "Hoc sinh roi truong hien tai de sang truong khac.", "Phu huynh + Van phong", "Bat ky luc nao", "Can xuat ho so, giay chuyen truong. CSDL nganh can cap nhat."],
    ["7", "Chuyen truong (den)", "Hoc sinh tu truong khac chuyen den nhap hoc.", "Phu huynh + Van phong", "Bat ky luc nao", "Can kiem tra ho so, xep lop. Truong cu phai da xuat du lieu."],
    ["8", "Thoi hoc", "Hoc sinh ngung hoc vinh vien (do hoan canh, ky luat, qua tuoi...).", "Phu huynh/Truong + Van phong", "Bat ky luc nao", "Can luu ly do. THCS: hiem gap vi giao duc bat buoc."],
    ["9", "Bao luu", "Tam ngung hoc (benh, hoan canh), giu quyen quay lai hoc tiep.", "Phu huynh + Van phong", "Bat ky luc nao", "Pho bien o THPT va dai hoc. Tieu hoc/THCS it gap."],
    ["10", "Tot nghiep", "Hoan thanh chuong trinh cap hoc (lop 5, lop 9, lop 12).", "Truong + So/Phong GD", "Cuoi nam cuoi cap", "Lop 5: xet tot nghiep. Lop 9: xet. Lop 12: thi THPT quoc gia."],
    ["11", "Tai nhap hoc", "Hoc sinh da thoi hoc hoac bao luu quay lai hoc tiep.", "Phu huynh + Van phong", "Dau nam / Giua nam", "Can xac dinh hoc tiep tu khoi nao, lop nao."],
]
add_table(ws1, overview_headers, overview_data)

# =====================
# SHEET 2: QUY TRINH THUC TE
# =====================
ws2 = wb.create_sheet("2. Quy trinh thuc te VN")
ws2.column_dimensions["A"].width = 5
ws2.column_dimensions["B"].width = 20
ws2.column_dimensions["C"].width = 18
ws2.column_dimensions["D"].width = 50
ws2.column_dimensions["E"].width = 30
ws2.column_dimensions["F"].width = 30

proc_headers = ["STT", "Nghiep vu", "Cap hoc", "Quy trinh thuc te", "Giay to can thiet", "Ai duyet"]
proc_data = [
    ["1", "Len lop - Tieu hoc", "Tieu hoc",
     "GV danh gia hoc sinh theo 3 muc: Hoan thanh tot, Hoan thanh, Chua hoan thanh (theo TT27/2020). Cuoi nam, GV de xuat danh sach len lop. Hieu truong ky quyet dinh.",
     "Phieu danh gia cuoi nam, Hoc ba", "GV chu nhiem + Hieu truong"],
    ["2", "Len lop - THCS", "THCS",
     "Dua tren diem trung binh cac mon va xep loai hanh kiem (theo TT22/2021). Hoc sinh dat yeu cau thi len lop. Truong hop can thi lai, hoc sinh thi lai cac mon chua dat roi xet lai.",
     "Bang diem, Hoc ba, Bien ban hop xet", "GV + HDSP truong"],
    ["3", "Len lop - THPT", "THPT",
     "Tuong tu THCS. Diem trung binh cac mon >= 5.0 va khong co mon nao duoi 3.5. Hanh kiem tu Trung binh tro len. Co the thi lai 1 lan.",
     "Bang diem, Hoc ba, Bien ban hop xet", "GV + HDSP truong"],
    ["4", "O lai lop - Tieu hoc", "Tieu hoc",
     "Theo TT27/2020, chi cho dup khi hoc sinh chua hoan thanh va phu huynh dong y. Xu huong hien tai la han che toi da viec cho dup o tieu hoc.",
     "Phieu danh gia, Don cua phu huynh", "Hieu truong"],
    ["5", "O lai lop - THCS/THPT", "THCS, THPT",
     "Hoc sinh khong dat dieu kien len lop sau khi thi lai. Hoac nghi hoc qua 45 buoi/nam (khong phep).",
     "Bang diem, Bien ban xet", "HDSP truong"],
    ["6", "Chuyen truong", "Tat ca",
     "Phu huynh nop don o truong cu. Truong cu cap giay chuyen truong + ho so. Phu huynh mang ho so den truong moi. Truong moi xet nhan va xep lop. Ca 2 truong cap nhat CSDL nganh.",
     "Don xin chuyen, Hoc ba, Giay chuyen truong, Giay khai sinh", "Hieu truong 2 truong"],
    ["7", "Thoi hoc", "THCS, THPT",
     "Phu huynh lam don (hoac truong lap bien ban). Truong ra quyet dinh. Cap nhat trang thai tren CSDL nganh. Tieu hoc va THCS (giao duc bat buoc) truong phai van dong hoc sinh quay lai.",
     "Don xin thoi hoc, Quyet dinh cua truong", "Hieu truong"],
    ["8", "Bao luu", "THPT",
     "Phu huynh lam don xin bao luu (ly do suc khoe, hoan canh). Truong xet duyet. Hoc sinh giu quyen quay lai hoc trong thoi han (thuong 1 nam). Khi quay lai, xep vao lop phu hop.",
     "Don xin bao luu, Giay xac nhan ly do", "Hieu truong"],
    ["9", "Tot nghiep tieu hoc", "Tieu hoc",
     "Xet tot nghiep. Hoc sinh hoan thanh chuong trinh lop 5 thi duoc cong nhan hoan thanh chuong trinh tieu hoc. Khong co thi.",
     "Hoc ba, Phieu danh gia lop 5", "Hieu truong"],
    ["10", "Tot nghiep THCS", "THCS",
     "Xet tot nghiep dua tren diem trung binh lop 9 va qua trinh hoc tap. Khong co ky thi rieng (tu 2006).",
     "Hoc ba, Bang diem tong ket", "Hoi dong xet tot nghiep"],
    ["11", "Tot nghiep THPT", "THPT",
     "Thi tot nghiep THPT quoc gia. Ket qua thi ket hop voi diem hoc ba de xet tot nghiep.",
     "Hoc ba, Phieu dang ky thi, Ket qua thi", "So GD&DT"],
]
add_table(ws2, proc_headers, proc_data)

# =====================
# SHEET 3: CSDL NGANH
# =====================
ws3 = wb.create_sheet("3. CSDL nganh Bo GD")
ws3.column_dimensions["A"].width = 5
ws3.column_dimensions["B"].width = 25
ws3.column_dimensions["C"].width = 55
ws3.column_dimensions["D"].width = 40

db_headers = ["STT", "Hang muc", "Mo ta", "Lien he voi THS"]
db_data = [
    ["1", "He thong CSDL nganh",
     "Bo GD&DT van hanh he thong tai csdl.moet.gov.vn. Chia theo cap: Mam non, Tieu hoc (C1), THCS (C2), THPT (C3), GDTX. Moi truong co tai khoan rieng de nhap lieu.",
     "THS can tuong thich voi cac truong du lieu ma CSDL nganh yeu cau, hoac ho tro xuat du lieu theo format cua Bo."],
    ["2", "Thong tin hoc sinh tren CSDL nganh",
     "Moi hoc sinh co ma dinh danh. Thong tin gom: ho ten, ngay sinh, gioi tinh, dan toc, dia chi, ho khau, doi tuong chinh sach, trang thai (dang hoc, chuyen di, thoi hoc, tot nghiep).",
     "Bang school_users hien tai chua co truong trang thai (status). Can bo sung."],
    ["3", "Quy trinh len lop tren CSDL nganh",
     "Cuoi nam hoc, truong nhap ket qua xep loai tung hoc sinh. He thong tu dong chuyen hoc sinh sang nam hoc moi o khoi lop tiep theo. Hoc sinh dup se o lai khoi cu.",
     "THS chua co tinh nang 'chuyen nam hoc'. Can xay dung quy trinh copy hoc sinh tu nam hoc cu sang nam hoc moi."],
    ["4", "Quy trinh chuyen truong tren CSDL nganh",
     "Truong cu 'xuat' hoc sinh (chuyen trang thai thanh 'da chuyen'). Truong moi 'nhap' hoc sinh bang ma dinh danh. Du lieu lien thong giua cac truong trong cung tinh qua he thong cua So GD.",
     "THS hien khong lien thong giua cac truong. Can thiet ke quy trinh 'xuat ho so so' va 'nhap ho so so' hoac API lien thong."],
    ["5", "Trang thai hoc sinh",
     "CSDL nganh quan ly cac trang thai: Dang hoc, Chuyen truong, Thoi hoc, Bao luu, Tot nghiep, Bi duoi hoc. Moi lan thay doi trang thai deu phai cap nhat.",
     "THS can them truong 'status' vao school_users va ghi log moi lan thay doi trang thai."],
    ["6", "Bao cao thong ke",
     "CSDL nganh tu dong tong hop so lieu: si so dau nam, cuoi nam, tang giam (chuyen den, chuyen di, thoi hoc, bo hoc...). Cac truong phai bao cao dinh ky.",
     "THS nen co man hinh dashboard thong ke tuong tu de truong doi chieu voi so lieu bao cao len Phong/So."],
]
add_table(ws3, db_headers, db_data)

# Save part 1
wb.save(WB_PATH)
print("Part 1 saved OK")
