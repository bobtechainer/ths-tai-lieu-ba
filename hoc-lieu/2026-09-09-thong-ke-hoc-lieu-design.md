# Báo cáo thống kê học liệu THS — thiết kế đã chốt

Ngày chốt thiết kế: 09/09/2026
Người yêu cầu: ba@techainer.com (NKG)
Đầu ra: một file `.xlsx`, 9 sheet
Người đọc: đội nội bộ NKG **và** một bên chuyên gia bên ngoài đánh giá/kiểm kê
Quy tắc viết: toàn bộ nhãn cột, tiêu đề và diễn giải bằng tiếng Việt phi kỹ thuật. Không tên bảng, không tên cột kỹ thuật, không mã trạng thái. Người không làm kỹ thuật phải hiểu mà không hỏi lại.

## Nguồn dữ liệu

Đọc trực tiếp cơ sở dữ liệu production qua Hasura: `https://api.truonghocsoquocgia.vn/v1/graphql`.
**Chỉ đọc.** Không mutation, không sửa, không xoá — kể cả để thử.
Mốc chốt số liệu ghi vào sheet Tổng quan; dữ liệu đang chạy thật (tổng số học liệu nhích 53.126 → 53.137 trong 25 phút khảo sát ngày 09/09/2026), nên hai lần chạy khác giờ sẽ ra số khác nhau vài đơn vị.

## Sáu quyết định đã chốt với người yêu cầu

1. **Phạm vi** — gồm cả hai loại học liệu, tách sheet riêng, cộng thêm một sheet gộp toàn bộ để tra chéo.
2. **Mục đích** — kiểm kê nội bộ và gửi chuyên gia ngoài đánh giá. Bỏ hết từ ngữ kỹ thuật.
3. **Con số "đang công khai"** — lấy **học liệu cộng đồng của nền tảng, không thuộc riêng trường nào**: 2.442. Không dùng 2.545 (có lẫn 103 học liệu gắn riêng cho trường), không dùng 50.965 (số trang admin đang hiện, không xét trạng thái nào).
4. **"Định dạng"** — hai cột riêng: "Loại học liệu" (nhãn hệ thống) và "Dạng tệp" (đuôi tệp thật).
5. **Cấu trúc** — đúng 9 sheet, chỉ bộ Kết nối tri thức có sheet chi tiết riêng.
6. **Đường dẫn** — trang học liệu trên web và trang sách. Không làm cột link tải tệp vì link tải có ký hạn, hết hạn sau 7 ngày.

## Cách đếm (định nghĩa khoá)

**Học liệu công khai của nền tảng = 2.442.** Lấy đúng bộ điều kiện mà trang Kho học liệu trên web đang dùng: chưa xoá, đã xuất bản, đặt công khai, đang bật hiển thị, chưa lưu trữ, và không gắn riêng cho trường nào. Bộ điều kiện này rút từ chính mã nguồn giao diện (`gkebook-school-frontend`, trang `/kho-hoc-lieu`), được hai lượt khảo sát độc lập xác nhận trùng nhau.

**Học liệu nằm trong sách.** Mỗi học liệu là một điểm chạm đặt trên một trang sách cụ thể. Đếm theo học liệu khác nhau, không đếm theo lượt gắn: bộ Kết nối tri thức có 25.972 lượt gắn nhưng chỉ 25.966 học liệu khác nhau (6 cái được gắn ở hai trang).

**Một bài "đã có học liệu"** = trong mục lục sách, mục đó (hoặc trang thuộc mục đó) có ít nhất một học liệu.

## Chuỗi liên kết dùng để dựng vị trí trong sách

`sách → trang (số trang, trỏ về một mục trong mục lục) → điểm chạm → học liệu`

Mục lục sách là một cây ba nhãn: chương / mục / bài. Cấp "tập" **không** có trong dữ liệu — mỗi tập là một quyển riêng, tên tập nằm trong tên sách. Nên cột vị trí chỉ có: Tên sách › Chương › Bài › Trang. Không dựng cấp "phần/tập" giả.

## Đường dẫn

- Học liệu: `https://truonghocsoquocgia.vn/hoc-lieu/<tên-rút-gọn-của-học-liệu>`
- Sách: `https://truonghocsoquocgia.vn/sach/<tên-rút-gọn-của-sách>`
- Mở sách đúng trang: `https://truonghocsoquocgia.vn/sach/<tên-rút-gọn>/doc#<số-trang>` — số sau dấu `#` là thứ tự trang trong quyển, đánh số từ 1.

## Chín sheet

| # | Sheet | Số dòng dự kiến | Nội dung |
|---|---|---:|---|
| 1 | Tổng quan | ~40 | Mốc chốt số liệu, ba con số lớn, bảng ba bộ sách kèm độ phủ, bảng loại học liệu công khai, các điểm cần lưu ý |
| 2 | Cách đếm | ~25 | Giải thích bằng lời thường vì sao cùng một câu hỏi ra 2.442 / 2.545 / 50.965 / 53.137 |
| 3 | KNTT – từng học liệu | 25.966 | Khối · Môn · Tên sách · Chương · Bài · Trang · Tên học liệu · Loại · Dạng tệp · Lượt xem · Link mở sách tại trang · Link học liệu |
| 4 | KNTT – theo bài | 4.333 | Mỗi bài một dòng: có mấy học liệu, gồm loại gì, hay còn trống |
| 5 | Thư viện – từng học liệu | 2.442 | Tên · Loại · Dạng tệp · Môn · Khối · Mức độ · Năm · Thẻ · Lượt xem · Link |
| 6 | Thư viện – bảng chéo | ~60 | Loại × khối và loại × môn |
| 7 | Bảng gộp toàn bộ | 53.137 | Mọi học liệu chưa xoá, có cột nói rõ đến từ đâu và đang ở trạng thái nào |
| 8 | Danh mục sách | 574 | Bộ · khối · môn · số trang · số học liệu · độ phủ bài · link đọc sách |
| 9 | Việc cần xử lý | ~10 nhóm | Các chỗ dữ liệu đang hở, mỗi nhóm một dòng kèm số lượng và giải thích |

## Những sự thật phải nói rõ trong báo cáo, không được im

Đây là các phát hiện có thể làm người đọc hiểu sai nếu không chú thích:

1. **Bộ SGK thống nhất trống 99,7%** — 204 quyển, 6.000 bài, chỉ 17 bài có học liệu.
2. **163 quyển Chân trời sáng tạo bị mất nhãn bộ sách** — nhãn bộ đã bị xoá mềm trong danh mục nhưng sách vẫn đang chạy. Báo cáo gọi lại đúng tên bộ và ghi chú việc này.
3. **Trang Báo cáo tổng quan trong admin đang hiện 50.965** trong khi số học liệu người dùng thật sự thấy được là 2.442. Chênh lệch này phải giải thích, không được im.
4. **726 học liệu của Kết nối tri thức nằm trên trang không rõ thuộc bài nào** (2,8%) — cột Chương/Bài sẽ trống, ghi rõ "chưa gắn vào mục lục" chứ không để trống suông.
5. **951 học liệu mang cờ đính kèm nhưng không gắn vào trang sách nào.**
6. **210 học liệu dán nhãn Kết nối tri thức nhưng chưa đặt lên trang nào.**
7. **Thẻ (hashtag) gần như toàn bộ do AI gắn** — 2.194/2.200 thẻ. Cột Thẻ phải nói rõ là do máy gắn, không phải người biên tập, và chỉ 575/3.586 học liệu thư viện có thẻ.
8. **685 học liệu công khai không có tệp tải về** — là bộ đề và học liệu tương tác, nội dung nằm trong hệ thống. Cột Dạng tệp ghi "không có tệp tải về", không để trống.
9. **Số ô "cấp học" chỉ phủ 2.460/3.586** học liệu thư viện; môn và khối phủ ~93%.
10. **Trường "từ khoá" trong hệ thống rỗng hoàn toàn** (0 bản ghi) — không dùng làm trục phân loại.

## Bẫy kỹ thuật khi lấy số (ghi để lần sau không mất thời gian)

- **Lọc bằng chuỗi tiếng Việt trả về sai.** `_ilike "%Kết nối%"` cho ra 0 dòng trong khi lọc bằng mã định danh cho ra 207. Mọi bộ lọc phải dùng mã định danh.
- Trường "công khai" cho phép để rỗng, nên điều kiện `= true` sẽ âm thầm bỏ các dòng rỗng. Đếm riêng nhóm rỗng.
- Chú thích số dòng trong file schema nằm **trước** bảng nó mô tả, không phải sau. Đọc theo bản năng là lệch một bảng.
- Bảng "định dạng học liệu" trong hệ thống rỗng 0 dòng; nhãn loại học liệu nằm ở bảng danh mục (11 loại đang bật).
- Không dùng bảng duyệt học liệu làm điều kiện (0 bản ghi, mã nguồn đã bỏ dùng).

## Kiểm tra trước khi giao

1. Đối chiếu lại mọi con số headline bằng một lượt truy vấn độc lập, cùng một mốc thời gian.
2. Dựng agent đóng vai **giáo viên** đọc file, không được giải thích trước: hỏi xem có hiểu từng sheet mà không cần hỏi lại không.
3. Dựng agent đóng vai **người không làm kỹ thuật** soát toàn bộ nhãn cột và câu diễn giải, chỉ ra chỗ còn mùi kỹ thuật.
4. Soát tính toàn vẹn file: số dòng từng sheet khớp số đã công bố, không cột nào rỗng toàn bộ, link bấm được.

## Đã giao — những gì khác với thiết kế ban đầu

Bản giao: `Kiểm kê học liệu THS - 09.09.2026.xlsx` (5,1 MB, 9 sheet). Ba vòng đọc độc lập (5 + 3 + 2 người đọc, trong đó luôn có một người đóng vai giáo viên chưa được giải thích gì) đã làm thay đổi mấy chỗ so với thiết kế:

1. **Thêm bảng thành phần "48.606 học liệu trong sách gồm những gì" lên ngay sheet Tổng quan.** 43.313 trong số đó (89%) là ảnh cắt từ chính trang sách. Không có bảng này thì con số tổng bị đọc thành khối lượng đã soạn.
2. **Mỗi tỉ lệ độ phủ tách thành hai cột**: "bài đã gắn thứ gì đó" (74,0% với Kết nối tri thức) và "bài có nội dung soạn riêng" (21,8%). Cột thứ hai mới là con số đưa đi họp. Áp cho cả sheet Tổng quan và sheet Danh mục sách — vì giáo viên mở sheet Danh mục sách trước, và ở đó Toán 4 hiện 100% trong khi tập hai không có bài nào có nội dung soạn riêng.
3. **Sheet 4 có bốn tình trạng** thay vì hai: có nội dung soạn riêng · chỉ có ảnh cắt trang sách · còn trống · không phải bài học (lời nói đầu, mục lục, phụ lục — trước đó bị đếm lẫn vào "còn trống").
4. **Bỏ nhãn "đã lưu trữ" cho học liệu trong sách.** Trang đọc sách lấy học liệu chỉ với điều kiện "chưa xoá", nên cờ lưu trữ không ẩn chúng. Thay bằng hai cột: "Người đọc sách có thấy" và "Có ở Kho học liệu". Kèm một dòng ở sheet 9 nhờ quản trị mở thử bằng máy chưa đăng nhập, vì quyền theo vai của Hasura không có trong repo nên không xác minh được từ mã nguồn.
5. **Bảng giải nghĩa gộp vào sheet 2** (không thêm sheet thứ 10, giữ đúng 9 sheet): giải nghĩa từng loại học liệu, từng trạng thái, từng tên cột khó và từng đuôi tệp.
6. **Thêm bảng chéo khối × môn** vào sheet 6 — giáo viên tiểu học không trả lời được câu "môn tôi khối tôi có mấy học liệu" nếu chỉ có hai bảng cũ.
7. **Cột "Cấp học" đổi thành "Mức độ"** — trường đó là Cơ bản/Trung bình/Nâng cao, không phải cấp học.
8. Đường dẫn chuyển sang công thức `HYPERLINK` (dùng dấu phẩy) vì hyperlink thường bị Excel cắt mất phần `#số trang`.

**Còn bỏ ngỏ, đã nói rõ trong file:** bộ Chân trời sáng tạo có 22.607 học liệu và 835 bài có nội dung soạn riêng nhưng không có sheet chi tiết theo bài (đúng phạm vi đã chốt); nếu bên thẩm định hỏi tới thì làm thêm một sheet giống sheet 4.

## Thu hẹp phạm vi (quyết định bổ sung, 09/09/2026)

Sau khi xem bản đầu, người yêu cầu chốt bỏ hai phần cho đỡ loạn:

1. **Bỏ bộ Chân trời sáng tạo** — 163 quyển, 22.607 học liệu.
2. **Bỏ học liệu do từng trường tự gắn riêng cho trường mình** — 514 học liệu.

Số liệu sau khi thu hẹp: 411 quyển sách (41.401 trang) · 25.999 học liệu trong sách · Kho học liệu 2.442 (không đổi) · bảng gộp 30.017 dòng · 20 việc ở sheet Việc cần xử lý.

Giữ lại một câu phạm vi ở đầu sheet Tổng quan nói rõ đã bỏ hai phần nào, để bên thẩm định ngoài không đọc thành làm thiếu.
