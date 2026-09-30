# Map tab Details vào các tab module — Testing THS ver2

Ngày: 2026-08-26
File đích: `C:\Users\letru\Downloads\Testing THS ver2.xlsx` (ghi trực tiếp, backup `.bak`)

## Mục tiêu

Sheet đang là bản v1 copy sang. Lấy tab `Details` làm chuẩn, đưa Expected Result /
Actual Result / Trạng thái dev về đúng dòng của đúng tab module, và điền nốt cột
Priority còn thiếu.

## Hiện trạng đã đo

- 27 sheet: `Summary Report`, `Nhân sự`, 24 tab module, `Details`.
- Tab module dùng chung 13 cột: Test ID | Description | Pre-conditions | Test Data |
  Steps | Expected Result | Actual Result | Status | Bug ID | Priority | Tester |
  Date | Môi trường (M:Q = Chrome/Edge/Safari/Mobile/Tablet). Tổng 7.816 case.
- `Details` 590 dòng: Module | Week | Môi trường | Bug ID (thực ra là Test ID) |
  Description | Pre-conditions | Test Data | Steps | Expected Result | Actual Result |
  Dev PIC | Test PIC | Trạng thái dev | Trạng thái test NKG | Trạng thái test MDS | Ghi chú.

Bốn phát hiện chi phối thiết kế:

1. Khớp theo Test ID đơn thuần không an toàn. `TC_MAT_002` ở Details là "Bấm vào một
   học liệu để xem chi tiết", cùng ID ở tab Kho học liệu lại là "Bố cục trên điện thoại".
2. 1.213 Test ID trùng giữa hai tab khác nhau (`TC_QUIZ_*` ở cả hai tab Bộ đề,
   `TC_BOOK_003` ở cả Site trường - Bộ sách và THS - App). Bắt buộc chặn phạm vi theo Module.
3. Tab `Ngân hàng câu hỏi` dùng tiền tố `TC_BANK_*` còn Details dùng `TC_NHCH_*` —
   ID lệch hoàn toàn nhưng 45/45 case khớp được theo nội dung.
4. Tab `Smart class` có 48 nhóm dòng trùng y hệt 4 trường nội dung → khớp nội dung
   thuần sẽ nhập nhằng, phải phá thế bằng Test ID.

## Quyết định đã chốt với người dùng

| Vấn đề | Quyết định |
|---|---|
| Giá trị cột Status | Copy nguyên văn `Trạng thái dev` (Reject / Done on production / Ready test on staging / Open / Ready to test) |
| Expected Result lệch (109 dòng) | Ghi đè theo Details |
| Actual Result cả hai bên đều có | Details thắng |
| Ghi chú, Dev PIC, Week, Trạng thái test NKG/MDS | Không thêm cột, bỏ; tra ngược ở tab Details |
| 20 dòng khớp mờ | Thêm thành dòng mới cuối tab |
| 17 dòng không có tab đích | Tạo 2 tab mới + 2 dòng trong Summary Report |
| Thang Priority | High/Medium/Low, khảo sát code từng module trước khi chấm |
| File ra | Ghi trực tiếp file gốc, có backup |
| Tab `Trang tính3` (bản copy của tab phân quyền) | Ghi vào cả hai tab, không xoá gì |

## Thuật toán khớp

Điểm giống = 0,40×Description + 0,15×Pre-conditions + 0,10×Test Data + 0,35×Steps,
tính bằng `difflib.SequenceMatcher` trên chuỗi đã gộp khoảng trắng và hạ chữ thường.
Có bước lọc thô bằng Jaccard trên token để chỉ tính điểm đầy đủ cho 15 ứng viên tốt nhất.

Thứ tự quyết định:

1. Chặn phạm vi ứng viên theo bảng định tuyến Module → tab.
2. Nếu Test ID có trong tab đích **và** điểm ≥ 0,60 → ghi vào đúng dòng đó.
3. Nếu không, lấy dòng điểm cao nhất; nhận nếu ≥ 0,75. Khi điểm bằng nhau, ưu tiên
   ứng viên trùng Test ID.
4. Còn lại → thêm dòng mới ở cuối tab.

Kết quả chạy thật trên 595 lượt ghi (590 dòng Details, 5 dòng phân quyền nhân đôi):

| Nhóm | Số lượt |
|---|---|
| Ghi vào dòng có sẵn, điểm ≥ 0,99 | 536 |
| Ghi vào dòng có sẵn, điểm 0,90–0,99 | 20 |
| Ghi vào dòng có sẵn, cần soát mắt | 2 |
| Thêm dòng mới ở tab có sẵn | 20 |
| Thêm dòng mới ở 2 tab mới | 17 |

## Điền Priority

~1.439 ô trống + 21 ô ghi tiếng Việt cần chuẩn hoá.

Rubric:
- **High** — chặn luồng chính, mất/sai dữ liệu, sai phân quyền, sai điểm/kết quả thi,
  không đăng nhập được, mất bài đang soạn.
- **Medium** — nhánh phụ, kiểm tra dữ liệu nhập, phân trang, tìm kiếm/lọc, thông báo,
  trạng thái rỗng, điều hướng.
- **Low** — giao diện, bố cục điện thoại, chuỗi quá dài, tương phản màu, edge ít gặp.

Định hướng phân bố ~30/50/20, không ép cứng.

Repo tra cứu: prj-ths-smart-class, gkebook-school-frontend, prj-ths-question-bank,
prj-ths-shadcn-admin-lms, prj-ths-shadcn-admin, prj-gkebook-shadcn-frontend.

## An toàn khi ghi

- Backup `.bak` trước khi ghi.
- Mở workbook với `data_only=False` để giữ công thức Summary Report và ô gộp.
- Hai tab mới copy đúng schema 13 cột + hàng phụ Chrome/Edge/Safari/Mobile/Tablet.
- Thêm dòng Summary Report dạng `=COUNTA('<tab>'!B2:B965)`.
- Xuất file đối soát `bao-cao-map-ver2.xlsx`: dòng Details, tab đích, dòng đích,
  điểm khớp, đường khớp (ID/nội dung), cờ cần soát mắt.

---

## Kết quả thực tế sau khi chạy (2026-08-26)

### Phần 1 - Đổ dữ liệu

595 lượt ghi trên 590 dòng Details (5 dòng phân quyền đổ vào 2 tab):

| Nhóm | Số lượt |
|---|---|
| Ghi vào dòng có sẵn, điểm khớp >= 0,99 | 536 |
| Ghi vào dòng có sẵn, điểm 0,90-0,99 | 20 |
| Ghi vào dòng có sẵn, cần soát mắt | 2 |
| Thêm dòng mới ở tab có sẵn (5 Kho học liệu, 15 Bộ đề) | 20 |
| Thêm dòng mới ở 2 tab mới | 17 |

Không có dòng đích nào bị hai dòng Details ghi chồng lên nhau.

Tên tab mới phải cắt còn 31 ký tự theo giới hạn Excel:
`Admin Trường - Báo cáo giáo viê` (Summary Report vẫn ghi tên đầy đủ, giống
cách file đang làm với `Site trường - ThiKiểm tra` / `Site trường - Thi/Kiểm tra`).

### Phần 2 - Priority

1.476 ô đã điền + 21 ô chuẩn hoá từ "Cao"/"Trung bình".
Phân bố phần mới: High 30% / Medium 53% / Low 17%.
Phân bố toàn bộ 7.852 case sau khi điền: **High 30% / Medium 50% / Low 20%** —
khớp phân bố sẵn có của team.

Bốn lỗi thật gặp khi dựng bộ luật, đã sửa và ghi lại để khỏi lặp:

1. `unicodedata.normalize('NFD')` **không tách chữ đ/Đ** (đó là ký tự riêng, không
   phải chữ có dấu phụ). Phải `.replace('đ','d')` trước. Không có bước này thì mọi
   luật về "đăng nhập", "điểm danh", "đình chỉ", "đồng thời", "đặt mật khẩu" đều chết.
2. `\bkhoa\b` khớp cả "từ **khóa**" -> mọi case tìm kiếm bị đẩy lên High.
3. `\btra loi\b` khớp cả "kiểm **tra lỗi**" -> mọi case kiểm tra lỗi bị đẩy lên High.
4. `\bdien thoai\b` (ý là thiết bị) khớp cả "**số điện thoại**" -> các case validate
   số điện thoại bị đẩy xuống Low. Đã dùng lookbehind `(?<!so )`.

Ngoài ra: sửa bộ luật bằng cách thay chuỗi trong shell rất dễ biến `\b` thành ký tự
backspace `\x08`; module có hàm `selfcheck()` để bắt lỗi này.

### Việc còn treo (không nằm trong phạm vi, chưa sửa)

Công thức ở Summary Report dùng phạm vi `B2:B965`, trong khi hai tab Bộ đề có
1.093 và 1.216 case. Dashboard đang **đếm thiếu 409 case** (144 + 265). Đây là
vấn đề có từ trước, nhưng 15 dòng mới thêm vào tab `Site trường - Bộ đề` nằm ở
dòng 1096+ nên cũng không được đếm. Nới phạm vi lên `B2:B1300` là xong.

---

## Vòng 2 - dọn sạch toàn bộ file (2026-08-26)

Người dùng yêu cầu một bản không còn khiếm khuyết. Đã rà toàn workbook và sửa 8 nhóm:

| # | Khiếm khuyết | Cách sửa |
|---|---|---|
| 1 | 20 dòng thêm mới trùng mã với dòng có sẵn trong cùng tab (5 Kho học liệu, 15 Bộ đề) | Đặt lại mã theo số kế tiếp của cùng họ mã: `TC_MAT_147..151`, `TC_QUIZ_NAV_047`, `TC_QUIZ_TC_046..047`, `TC_QUIZ_TT_075`, `TC_QUIZ_TN_070`, `TC_QUIZ_KEO_077`, `TC_QUIZ_GD_072..074`, `TC_QUIZ_CT_093`, `TC_QUIZ_XT_078..082`. Bảng đối chiếu mã cũ - mã mới ở sheet 2 của file đối soát. |
| 2 | `Site trường - Lịch giảng dạy` dòng 3 có Description = "xZ" | Lấy lại từ Details: "Kiểm tra giao diện Thời khóa biểu" |
| 3 | `Site trường - Nhiệm vụ` dòng 286 là hàng header lặp giữa sheet, làm COUNTA đếm dư 1 | Xoá dòng |
| 4 | 2 tab mới thiếu ô gộp hàng tiêu đề | Nhân bản 13 ô gộp từ tab mẫu, thêm freeze panes A3 |
| 5 | Công thức Summary Report dùng `B2:B965` -> đếm thiếu 409 case ở hai tab Bộ đề | Nới lên `B2:B2000` cho cả 26 dòng |
| 6 | Cột Passed / Failed / Blocked / Untested / Pass Rate của Summary Report bỏ trống -> M3:M6 luôn ra 0 | Điền COUNTIF trên cột Status của từng tab, kèm chú giải cách quy đổi ngay dưới bảng |
| 7 | Tab `Trang tính3` là bản sao -> 21 case bị đếm hai lần trong ô tổng | Giữ nguyên dòng 25 và tab, chỉ đổi nhãn thành "(bản sao ... không tính vào tổng)" và đổi các ô tổng thành `SUM(C11:C24)+SUM(C26:C40)` |
| 8 | 4 ô Bug Summary trỏ tới sheet `Bug Log` không tồn tại -> #REF! | Dựng sheet `Bug Log` từ 214 case đang có Status = "Reject", cột A = Bug ID, cột E = Trạng thái (Open); nới phạm vi công thức lên 2000 và sửa lỗi `E1:E965` thành `E2:E2000` |

### Quy đổi Status sang Passed/Failed/Blocked (lựa chọn duy nhất mang tính suy diễn)

Cột Status của tab module vẫn giữ nguyên văn giá trị `Trạng thái dev` theo đúng quyết định
ban đầu. Riêng dashboard cần gộp nhóm, nên dùng bảng sau, đã ghi thành chú giải trong sheet:

- Passed = `Done on production`
- Failed = `Reject`
- Blocked = `Open`
- Untested = Total trừ ba mức trên (gồm `Ready test on staging`, `Ready to test`, ô trống)

### Kết quả dashboard sau khi sửa (tự tính lại)

Total 7.831 (đã bỏ 21 case của tab bản sao) / Passed 202 / Failed 213 / Blocked 5 /
Untested 7.411. Pass Rate 2,6% — đây là trạng thái thật, vì chỉ 604 trong 7.852 case
có kết quả test, phần còn lại chưa chạy.

### Đã kiểm

- Không còn mã test case trùng trong cùng tab.
- Không còn dòng rác, không còn ô Priority trống, Priority chỉ có High/Medium/Low.
- 26 dòng module ở Summary Report: COUNTA khớp đúng số case thật của từng tab.
- 60 file XML trong xlsx phân tích được, không lỗi; 0 công thức sai tên sheet hoặc thiếu dấu nháy.
- 6 ảnh biểu đồ và toàn bộ comment còn nguyên. 30 sheet (27 gốc + 2 tab mới + Bug Log).

### Còn lại cần người viết, không tự bù

7 case thiếu Steps hoặc Expected Result mà **tab Details cũng trống** nên không có gì để
bù: `TC_BANK_166`, `TC_BANK_206`, `TC_BANK_207`, `TC_BANK_230`, `TC_BANK_270`,
`TC_LIVECLASS_33`, `TC_LIVECLASS_34`. Liệt kê ở sheet 4 của file đối soát.

37 dòng thêm mới không có Tester và Date vì tab Details có cột `Test PIC` trống toàn bộ.
