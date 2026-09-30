# Chia nhóm chuyên gia đọc và duyệt học liệu — bộ SGK thống nhất

Ngày: 21/09/2026
Sản phẩm: `Đề xuất chia nhóm chuyên gia học liệu - 21.09.2026.xlsx` (7 sheet, ở thư mục gốc)
Script dựng: `C:\Users\letru\AppData\Local\Temp\claude\ths-plan\build_nhom.py`
Nguồn số liệu: `D:\[NKG nội bộ] - Kế hoạch làm học liệu cho bộ SGK thống nhất.xlsx`, sheet 3

## Bối cảnh

Hiệp hội học liệu số quốc gia sẽ cử chuyên gia giúp nâng chất lượng học liệu của bộ SGK thống nhất.
Cần một đề xuất về cách chia nhóm để mang đi bàn với hiệp hội.

## Những điều đã chốt với user

| Câu hỏi | Chốt |
|---|---|
| Chuyên gia làm gì | Duyệt chất lượng sau khi đội NKG làm xong — không tự soạn, không định hướng trước |
| Cơ chế duyệt | Duyệt mẫu trước → chốt mẫu → đội nhân rộng → hiệp hội duyệt lô nhân rộng |
| Đơn vị của một mẫu | Theo **nhóm môn** |
| Trục ngoài | **Cấp học**, theo đúng gợi ý của hiệp hội |
| Số cấp | **2 cấp: Tiểu học (1–5) và Trung học (6–12)**. Ban đầu chia 3 cấp, user chốt lại: Toán 6 và Toán 12 vẫn là một người kiểm duyệt |
| Quy mô mỗi nhóm | 1–2 chuyên gia hiệp hội |

## Thiết kế

**Khung: 2 cấp × 6 nhóm môn = 12 ô. Mỗi ô là một nhóm và cũng là một mẫu phải chốt.**

Hai cấp: Tiểu học (1-5) · Trung học (6-12). Hệ quả phải liệu trước: cấp trung học gánh 7.771 học
liệu so với 3.584 của tiểu học, nặng gấp 2,2 lần. Cách chữa đã ghi trong file: các ô nặng của cấp
trung học thì xin hai chuyên gia thay vì một.

Sáu nhóm môn, gom theo **dạng học liệu** (vì mẫu là mẫu của một dạng học liệu), giữ nguyên tên qua
cả hai cấp nhưng ruột khác nhau theo chương trình 2018:

1. Ngữ văn – Tiếng Việt (2.254)
2. Toán (1.511)
3. Khoa học tự nhiên (1.254) — TN&XH + Khoa học | KHTN (6-9) + Lí, Hoá, Sinh (10-12)
4. Khoa học xã hội – Đạo đức (1.282)
5. Tin học – Công nghệ (1.304)
6. Nghệ thuật – Thể chất – Trải nghiệm (3.750, gồm cả Quốc phòng ở THPT)

**Chỗ gộp 2 cấp gây khó:** ô Khoa học tự nhiên và ô Khoa học xã hội của cấp trung học mỗi ô ôm hai
chặng khác nhau, vì chương trình 2018 gộp môn ở lớp dưới rồi tách dần: Tự nhiên và Xã hội (1-3) →
Khoa học + Lịch sử và Địa lí (4-5) → Khoa học tự nhiên + Lịch sử và Địa lí (6-9) → Vật lí, Hoá học,
Sinh học + Lịch sử, Địa lí, GD kinh tế và pháp luật (10-12). Người dạy Khoa học tự nhiên THCS không
thay được người dạy Vật lí THPT. Hai ô này vì thế nên xin hai chuyên gia mỗi ô.

**Ba ghế trong mỗi ô:** chuyên gia môn (hiệp hội cử) · thầy cô đang dạy đúng môn đúng cấp (NKG mời từ
trường) · người giữ chuẩn hình thức chung cho cả 12 ô (NKG, một người cho cả bộ).

**Mở nhóm theo sóng.** Lịch sản xuất vốn đi tuần tự từ lớp 1 lên lớp 12, nên chuyên gia vào lần lượt:
2 người trước tuần 1 (Ngữ văn + Toán tiểu học), +2 ở tuần 2 (Ngữ văn + Toán trung học), +2 ở tuần 6
(Khoa học tự nhiên cả hai cấp). Hết 10 tuần: **6 người**, không lần nào mời quá 2 người.
Sáu ô còn lại mở ở đợt sau.

**Bộ tiêu chí chấm mẫu** mượn bốn nhóm tiêu chuẩn của Thông tư 22/2026, viết lại cho học liệu rời:
đúng và hợp pháp · dạy được và dễ dùng · chạy được trên máy của trường · an toàn cho học sinh.
Nhóm 1 và 4 là cổng chặn, NKG lọc trước phần hình thức để tiết kiệm thời gian chuyên gia.

## Hai điều phải nói thẳng trong file

- Nhóm 1–2 người thì gọi là **chuyên gia cố vấn chốt mẫu**, không gọi là hội đồng thẩm định.
  Thông tư 33/2017 Điều 12 đòi tối thiểu 7 người, số lẻ, ≥1/3 là nhà giáo đang dạy; EdReports dùng
  5 người/nhóm; nghiên cứu về đồng thuận chuyên gia đòi tối thiểu 3.
- Học liệu rời **chưa có bộ tiêu chí quốc gia nào**. Thông tư 22/2026 chỉ phủ SGK điện tử chuyển thể
  từ SGK in đã duyệt; Thông tư 09/2021 Điều 7 với học liệu trong trường chỉ nói "tổ chuyên môn thông
  qua, hiệu trưởng phê duyệt", không kèm tiêu chí.

## Căn cứ

- Thông tư 33/2017/TT-BGDĐT Điều 12 — hội đồng lập theo từng môn ở từng cấp, ≥7 người, số lẻ, ≥1/3
  nhà giáo đang dạy. Tức mô hình pháp lý là "môn × cấp", không phải "cấp rồi mới tới môn".
- Thông tư 22/2026/TT-BGDĐT (6/4/2026, hiệu lực 22/5/2026) — bốn nhóm tiêu chuẩn cho SGK điện tử.
- Thông tư 09/2021/TT-BGDĐT Điều 7 — cổng duyệt học liệu trong trường, không kèm tiêu chí.
- Chương trình GDPT tổng thể kèm Thông tư 32/2018 — mạch gộp môn rồi tách dần.

## Chỗ còn là ước, chưa đo

- Thời gian mỗi ghế mất bao nhiêu buổi. Phải đo thật ở vòng chốt mẫu đầu tiên rồi sửa lại trước khi
  hứa với hiệp hội.
- Khối lượng từng tuần: quyển nào trải nhiều tuần thì chia đều cho các tuần đó.
- Bộ sách hiện không có quyển Tiếng Anh nào. Nếu thêm thì phải thêm một nhóm môn và tính lại.
