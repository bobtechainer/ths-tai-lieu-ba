# Tài liệu BA - Trường học số

Repo này gom các script và trang HTML mà Bob làm trong vai trò BA cho dự án Trường học số (THS). Script Python phần lớn đọc một file Excel hoặc Word ở máy, rồi ghi ra file Excel, Word hay PowerPoint để gửi cho khách, đối tác hoặc đội nội bộ. Trang HTML là bài nghiên cứu, sơ đồ hệ thống và kịch bản minh họa, mở thẳng bằng trình duyệt. Các file `.md` là bản thiết kế đã chốt của vài báo cáo Excel lớn.

Người làm: Bob (BobLe, bob@techainer.com), tháng 5–9/2026. Tình trạng: đã dùng thật. Các file đầu ra đã gửi đi, repo này là bản lưu để bàn giao.

## Mở hoặc chạy thế nào

- Trang HTML: mở đúp bằng trình duyệt, không cần cài gì. Vài trang tải phông từ Google Fonts nên có mạng thì chữ hiện đúng hơn.
- Script Python cần Python 3 và các thư viện `openpyxl`, `python-docx`, `python-pptx`. Riêng `csdl-dong-nai/_form.py` dùng thêm `playwright`.

```
pip install openpyxl python-docx python-pptx playwright
```

- Trên Windows, chạy với `PYTHONIOENCODING=utf-8` để màn hình dòng lệnh in được tiếng Việt.
- Đa số script trỏ cứng đường dẫn về thư mục gốc cũ `c:\Trường học số - source code\`. Script đọc file đầu vào thì file đó không có trong repo (xem mục "Không có trong repo").
- Chạy được ngay, không cần file đầu vào: `kich-ban-demo/export_demo_script.py`, `thiet-bi-man-hinh/export_device_analysis.py`, `lop-hoc-thong-minh/export_smartclass_en_pptx.py`, `module-ai/generate_excel_report.py` và bộ `vong-doi-hoc-sinh/gen_vongdoi*.py`. Nhớ sửa đường dẫn ghi file ở cuối script nếu máy khác thư mục.

## Trong repo có gì

"Đầu ra" là file script ghi ra ở máy gốc. Các file đó không có trong repo, trừ khi ghi rõ.

| File | Dùng để làm gì | Đầu ra | Ngày |
|---|---|---|---|
| `csdl-dong-nai/parse_api.py` | Bóc danh sách thông tin phải gửi của 37 đầu mục (số 39–75) từ bản chữ của tài liệu API CSDL Đồng Nai 12/08/2026 | `_rs/api_fields.json` | 24/08/2026 |
| `csdl-dong-nai/build.py` | Đối chiếu hồ sơ trường, hồ sơ lớp: ngành đòi thông tin gì, web quản lý trường đã có chưa, nằm ở màn nào | dùng chung, không tự ghi file | 24/08/2026 |
| `csdl-dong-nai/map2.py` | Phần còn lại của bảng đối chiếu: nhân sự, học sinh, phân công, môn học | dùng chung | 24/08/2026 |
| `csdl-dong-nai/core.py` | Gom hai bảng trên thành từng đầu mục kèm trạng thái | dùng chung | 24/08/2026 |
| `csdl-dong-nai/final.py` | Điền checklist gốc, giữ nguyên 5 cột và 38 dòng đầu của người khác. Đây là bản đã giao | `Checklist_RaSoat_API_CSDL_DongNai_CThu.xlsx` | 24/08/2026 |
| `csdl-dong-nai/gen.py`, `huongdan.py` | Bản trước của `final.py`: 2 sheet (checklist, dò từng trường) và thêm sheet "Cách đọc file" | cùng tên file trên, ghi đè | 24/08/2026 |
| `csdl-dong-nai/trace_p1.py` | Bảng dò từng trường phần 1, viết trước `build.py`. Không script nào gọi tới nữa | không có | 24/08/2026 |
| `csdl-dong-nai/_form.py` | Mở một form trên trang UAT bằng phiên đăng nhập đã lưu, ghi lại tab, nhãn, ô nhập và chụp ảnh | `_rs/<tên>.txt`, `_rs/<tên>.png` | 24/08/2026 |
| `phan-quyen/read_xlsx.py` | In file rà soát tính năng trùng với VNTIT ra màn hình để đọc | `xlsx_output.txt` | 31/07/2026 |
| `phan-quyen/read_perm.py` | In file phân quyền tính năng Sổ liên lạc điện tử của VNTIT ra màn hình | `perm_output.txt` | 31/07/2026 |
| `phan-quyen/pq_data.py` | Bảng chia việc theo mã chức năng: THS tự làm hay nhúng màn VNTIT, đặt ở menu nào, dev làm gì | dữ liệu cho script dưới | 31/07/2026 |
| `phan-quyen/gen_phanquyen.py` | Thêm 3 cột (bên phụ trách, vị trí menu, yêu cầu cho dev) vào file phân quyền của VNTIT | `YEU_CAU_PHAN_QUYEN_v3.xlsx` | 31/07/2026 |
| `kich-ban-demo/export_demo_script.py` | Kịch bản demo THS 5 phần: trước giờ học, trong giờ, bài về nhà, báo cáo, admin trường | `Kịch bản Demo THS - Nâng cao.xlsx` | 30/07/2026 |
| `kich-ban-demo/export_demo_v2.py` | Lấy file kịch bản gốc (Mobifone Solution), thêm phần bài về nhà và phần báo cáo | cùng tên file trên, ghi đè | 30/07/2026 |
| `kich-ban-demo/export_demo_v3.py` | Chép đủ 4 sheet gốc, thêm một sheet "Sau giờ học" | `Kịch bản Demo THS - Nâng cao v3.xlsx` | 30/07/2026 |
| `thiet-bi-man-hinh/export_device_analysis.py` | So hai máy tính bảng Joystar J16 và VASOUN L10 khi chạy THS, chia máy cho buổi demo | `Phan tich thiet bi - THS.docx` | 30/07/2026 |
| `thiet-bi-man-hinh/nghien-cuu-man-hinh-tuong-tac-philips.html` | Nghiên cứu đưa Lớp học thông minh lên màn hình tương tác Philips: chọn đời máy nào, cài app bằng cửa nào | trang HTML | 24/08/2026 |
| `thiet-bi-man-hinh/so-sanh-smart-class-vs-philips.html` | Đối chiếu từng hạng mục phần mềm có sẵn trên Philips với Lớp học thông minh | trang HTML | 24/08/2026 |
| `lop-hoc-thong-minh/trang-thai-lop-hoc.html` | Một tiết học theo mốc giờ: trạng thái lớp trên lịch và trạng thái dưới hệ thống lệch nhau lúc nào | trang HTML | 20/07/2026 |
| `lop-hoc-thong-minh/Smartclass 2.0 - EN translation.md` | Bản dịch tiếng Anh 19 slide giới thiệu Smartclass 2.0 | không có | 25/08/2026 |
| `lop-hoc-thong-minh/export_smartclass_en_pptx.py` | Dựng bản dịch trên thành PowerPoint, chữ sửa được sau khi đưa lên Canva | `Smart School - EN deck.pptx` | 25/08/2026 |
| `so-do-he-thong/so-do-he-thong-truong-hoc-so.html` | Sơ đồ hệ thống THS: 4 chặng, 20 khối, bấm khối để xem chi tiết, có nút VI/EN. Làm để đưa vào CV | trang HTML | 10/09/2026 |
| `so-do-he-thong/so-do-he-thong.html` | Bản làm lại kiểu làn: mỗi làn một nhóm người dùng, bấm khối thì sáng các đường nối | trang HTML | 10/09/2026 |
| `tuyen-sinh-ielts/ielts-xet-tuyen-dai-hoc-2027.html` | Bài "IELTS đáng mấy điểm" về xét tuyển đại học 2027 ở 18 trường Hà Nội, viết cho phụ huynh và học sinh lớp 11 | trang HTML | 15/09/2026 |
| `tuyen-sinh-ielts/ho-so-tuyen-sinh-18-truong.html` | Hồ sơ tuyển sinh từng trường trong 18 trường, IELTS được tính bao nhiêu ở mỗi nơi | trang HTML | 15/09/2026 |
| `tuyen-sinh-ielts/ielts-tuyen-sinh-2027-full.html` | Bản gộp bài đọc và hồ sơ 18 trường vào một trang | trang HTML | 17/09/2026 |
| `nghien-cuu/gamification-platforms.html` | Bảng hơn 40 nền tảng gamification giáo dục: năm, nước, vốn, doanh thu, tình trạng | trang HTML | 14/08/2026 |
| `nghien-cuu/nghien-cuu-5-app-nau-an.html` | So 5 ứng dụng nấu ăn (ReciMe, Paprika, Flavorish, Mealime, Eat This Much) | trang HTML | 18/08/2026 |
| `thiet-ke-figma/figma_audit.example.py` | Đếm module, tính năng, màn hình, component trong các file Figma của THS qua API Figma | `figma_audit_report.html`, `figma_audit_data.json` | 10/08/2026 |
| `vong-doi-hoc-sinh/gen_vongdoi.py`, `_p2.py`, `_p3.py` | Dựng tài liệu vòng đời học sinh (nhập học, xếp lớp, lên lớp, chuyển trường...), so với THS đang có. Chạy lần lượt 3 file | `vong-doi-hoc-sinh-THS.xlsx`, có trong repo | 14/09/2026 |
| `vong-doi-hoc-sinh/gen_lifecycle_part1..3.py` | Bản trước của bộ trên, tên sheet không dấu | cùng tên file, ghi đè | 14/09/2026 |
| `vong-doi-hoc-sinh/vong-doi-hoc-sinh-THS.xlsx` | Kết quả của bộ `gen_vongdoi`, 7 sheet, không có dữ liệu cá nhân | không có | 14/09/2026 |
| `module-ai/bao-cao-module-ai-gkebook-2026-06-03.md` | Đối chiếu module và tính năng AI giữa hai frontend `gkebook-school-frontend` và `prj-gkebook-shadcn-frontend` | không có | 03/06/2026 |
| `module-ai/generate_excel_report.py` | Bảng Excel của báo cáo trên: module nào có AI, có nên gộp, gợi ý AI tiếp | `bao-cao-module-ai-gkebook-2026-06-03.xlsx` | 03/06/2026 |
| `module-ai/export_ai_modules.py` | Thêm 4 cột về AI vào file danh sách module NKG, giọng kỹ thuật | `Danh sách Module hiện có NKG - Đề xuất AI - bo-sung-AI.xlsx` | 03/06/2026 |
| `module-ai/export_ai_modules_humanized.py` | Như trên, viết lại cho người không làm kỹ thuật | `Danh sách Module hiện có NKG - Đề xuất AI - humanized.xlsx` | 03/06/2026 |
| `mau-import-cau-hoi/vba-macro-cau-hoi.example.bas` | Macro cho sheet "Câu hỏi" của mẫu import: chọn loại câu ở cột B thì tự điền, khóa ô đáp án và gắn danh sách chọn đáp án đúng | dán vào file `mau-import-cau-hoi*.xlsm` | 28/05/2026 |
| `hoc-lieu/2026-09-09-thong-ke-hoc-lieu-design.md` | Thiết kế và kết quả báo cáo kiểm kê học liệu, số liệu đọc từ hệ thống thật (chỉ đọc) | `Kiểm kê học liệu THS - 09.09.2026.xlsx` | 09/09/2026 |
| `hoc-lieu/2026-09-17-ke-hoach-san-xuat-hoc-lieu-design.md` | Kế hoạch 10 tuần làm học liệu cho bộ SGK thống nhất: khối lượng, năng suất, số người | `Kế hoạch sản xuất học liệu SGK thống nhất - 17.09.2026.xlsx` | 17/09/2026 |
| `hoc-lieu/2026-09-21-chia-nhom-chuyen-gia-design.md` | Đề xuất chia nhóm chuyên gia hiệp hội duyệt học liệu: 2 cấp học × 6 nhóm môn | `Đề xuất chia nhóm chuyên gia học liệu - 21.09.2026.xlsx` | 21/09/2026 |
| `kiem-thu/2026-08-26-testcase-ver2-mapping-design.md` | Cách khớp tab Details vào các tab module của file testcase THS ver2, luật chấm Priority, các lỗi đã gặp | `Testing THS ver2.xlsx` (thư mục Downloads) | 26/08/2026 |

## Dữ liệu

Repo không có dữ liệu học sinh, giáo viên hay tài khoản thật. Tên người xuất hiện trong kịch bản demo là tên mẫu ("Nguyễn Văn A", "Trần Văn B", email `phuhuynhdemo@gmail.com`). Trang `trang-thai-lop-hoc.html` dùng tên bịa (An, Bình). Số liệu trong các bài IELTS, Philips, gamification lấy từ nguồn công khai, có ghi nguồn trong trang. Số liệu học liệu trong các file `hoc-lieu/*.md` là số tổng đếm từ hệ thống thật, không gắn với ai.

## Không có trong repo

Những thứ dưới đây cố ý để lại ở máy gốc, thư mục `c:\Trường học số - source code\`:

- File Excel, Word, PowerPoint, PDF đã giao (kịch bản demo, yêu cầu phân quyền, checklist Đồng Nai, kiểm kê học liệu, kế hoạch sản xuất, testcase...). Đây là tài liệu gửi khách và đối tác. Vài file có sheet tài khoản đăng nhập demo, vài file là tài liệu của VNTIT hoặc của Sở.
- Hai file tài khoản chuyên gia học liệu ngày 24/09/2026. Trong đó có tên và tài khoản của người thật.
- Bảng kết quả thi thử THPT môn Địa lý của trường THPT Sơn Tây. Đây là điểm của học sinh thật.
- Thư mục `_rs`. Trong đó có phiên đăng nhập trang UAT, ảnh chụp hồ sơ học sinh và bản chữ của tài liệu API.
- File dump và file trung gian (`*_output.txt`, `xlsx_full_dump.txt`, `details_output.txt`, `spreadsheet_data.json`, `humanized_data.json`, `figma_audit_data.json`, `tail.bin`...). Chúng chép lại nội dung các file Excel ở trên.
- Bản gốc `figma_audit.py` (có token Figma cá nhân) và `figma_audit_report.html` (có đường link thẳng tới file thiết kế). Repo chỉ giữ bản `.example`.
- Bản gốc `vba-macro-cau-hoi.bas` (có mật khẩu khóa sheet) và các file mẫu import câu hỏi `.xlsm`, `.xlsx`.
- `mk.py`, bản đầu của checklist Đồng Nai. File này ghi tên đăng nhập tài khoản UAT, và đã có `final.py` thay.
- Ảnh chụp màn hình cho báo cáo module AI, chụp từ hệ thống đang chạy.
- `ddg.html`, `o.html`, `l.html`, `h.html`: trang kết quả tìm kiếm và trang tra văn bản cào về, `h.html` rỗng.
- Script sinh file kiểm kê học liệu, kế hoạch sản xuất, chia nhóm chuyên gia, map testcase ver2 và bản sơ đồ kiểu làn. Chúng nằm ở thư mục tạm trên máy, không nằm trong thư mục dự án, có thể đã mất. Repo chỉ có bản thiết kế `.md`.

## Lưu ý cho người nhận

- Script giữ nguyên như lúc dùng, chưa sửa đường dẫn. Muốn chạy lại thì đặt file đầu vào đúng chỗ hoặc sửa đường dẫn ở đầu và cuối script.
- CSDL Đồng Nai: chạy `parse_api.py` trước, rồi `final.py`. Hai script này cần `_rs/api_full.txt` và `_rs/api_fields.json`, không có trong repo. Phạm vi chỉ là đầu mục 39–75. Đầu mục 1–38 do người khác phụ trách.
- `_form.py` cần phiên đăng nhập `_rs/auth.json` và hai file `fields.js`, `text.js`. Hai file js này không còn trên máy.
- `figma_audit.example.py`: đặt biến môi trường `FIGMA_TOKEN` bằng token của mình và điền mã file Figma vào chỗ `FILE_KEY_xx`.
- `vba-macro-cau-hoi.example.bas`: sửa `DOI_MAT_KHAU_KHOA_SHEET` thành đúng mật khẩu khóa sheet trong file mẫu.
- `so-do-he-thong.html` được sinh bằng script, không viết tay. Sửa tay dễ làm lệch chữ với đường nối.
- Bài IELTS chốt số liệu ngày 14–17/09/2026. Đề án tuyển sinh 2027 của các trường ra khoảng tháng 12/2026, lúc đó phải đối chiếu lại bảng quy đổi.
- Hai trang Philips chưa được kiểm trên máy thật. Kết luận kỹ thuật là suy ra từ thông số và tài liệu hãng.
- Kế hoạch sản xuất học liệu dựa trên năng suất 22,7 học liệu/người/tuần, đo từ đợt năm 2025. Hai đợt gần nhất chậm hơn nhiều, nên đo lại trước khi đem đi họp.
- Vài trang HTML có link tới trang claude.ai riêng của Bob. Người khác có thể không mở được link đó.
