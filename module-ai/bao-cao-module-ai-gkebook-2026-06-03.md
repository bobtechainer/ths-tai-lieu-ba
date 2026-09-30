# Báo cáo đối chiếu module và tính năng AI

Ngày rà soát: 2026-06-03  
Phạm vi rà soát:
- `gkebook-school-frontend`
- `prj-gkebook-shadcn-frontend`

## 1. Tình trạng mã nguồn trước khi rà

Đã kéo mã nguồn mới nhất từ nhánh `develop` của cả hai project.

- `gkebook-school-frontend`: đã chuyển sang `develop`, pull thành công, mã nguồn đang đồng bộ với `origin/develop`.
- `prj-gkebook-shadcn-frontend`: đã chuyển sang `develop`, pull thành công, mã nguồn đang đồng bộ với `origin/develop`.

---

## 2. Mục đích của báo cáo

Báo cáo này dùng để:

1. Liệt kê các module hiện có của từng project.
2. Chỉ ra những module đang trùng nhau giữa hai project và có thể gộp lại.
3. Xem mỗi module đã có tính năng AI hay chưa.
4. Mô tả ngắn gọn tính năng AI đó đang làm gì.
5. Gợi ý thêm những hướng AI phù hợp để phát triển tiếp.

---

## 3. Nhận xét tổng quan

Hai project này không phải là hai sản phẩm hoàn toàn tách rời.

Nếu nhìn vào route, cây thư mục, dịch vụ AI, chatbot, quiz studio và kho học liệu, có thể thấy hai bên đang chia sẻ khá nhiều ý tưởng và cùng đi trên một lõi nghiệp vụ học tập số.

Tuy vậy, mỗi project đang đi theo một hướng khác nhau:

### `gkebook-school-frontend`
Thiên về hệ thống dùng trong môi trường nhà trường, có thêm nhiều phần phục vụ quản lý dạy học như:
- lớp học
- điểm danh
- thời khóa biểu
- nhiệm vụ
- báo cáo học sinh
- quản trị nhà trường
- tích hợp lớp học trực tuyến, lớp học thông minh

### `prj-gkebook-shadcn-frontend`
Thiên về nền tảng học liệu và trải nghiệm học tập cá nhân hơn, nổi bật ở:
- bộ sách
- kho học liệu
- thư viện cá nhân
- hồ sơ luyện tập
- gói cước
- chatbot toàn cục
- tự học
- quiz studio
- một số tính năng AI “gần người học” hơn như tóm tắt tin tức, chuyển văn bản thành giọng nói

---

## 4. Danh sách module hiện có của từng project

# 4.1. Các module hiện có của `gkebook-school-frontend`

Dựa trên route, menu và cấu trúc thư mục, có thể gom thành các nhóm module sau.

## A. Nhóm tài nguyên học tập

1. **Bộ sách**
   - Danh sách sách
   - Xem chi tiết sách
   - Đọc sách

2. **Kho học liệu**
   - Danh sách học liệu
   - Xem thêm học liệu
   - Xem chi tiết học liệu
   - Chỉnh sửa học liệu
   - Duyệt học liệu

3. **Thư viện của tôi**
   - Học liệu đã tạo
   - Học liệu yêu thích
   - Học liệu đã tải lên

## B. Nhóm dạy học và lớp học

4. **Lớp học**
   - Danh sách lớp học
   - Chi tiết lớp học
   - Tạo nhiệm vụ cho lớp
   - Báo cáo học sinh theo lớp

5. **Điểm danh**
   - Điểm danh theo lịch
   - Điểm danh theo lớp
   - Màn hình điểm danh tổng quan

6. **Thời khóa biểu**
   - Lịch giảng dạy

7. **Nhiệm vụ học tập**
   - Danh sách nhiệm vụ
   - Chi tiết nhiệm vụ
   - Làm bài
   - Chấm điểm

8. **Tự học**
   - Danh sách nội dung tự học
   - Làm bài tự học
   - Xem kết quả tự học

## C. Nhóm kiểm tra, đánh giá

9. **Thi – kiểm tra (Phòng thi)**
   - Làm bài kiểm tra
   - Xem bài kiểm tra
   - Xem trước đề
   - Xem trước theo biểu mẫu
   - Thêm câu hỏi
   - Chỉnh sửa câu hỏi

10. **Ngân hàng câu hỏi**
    - Danh sách câu hỏi dùng cho kiểm tra, luyện tập

## D. Nhóm thông tin và thông báo

11. **Trang chủ**
12. **Tin tức**
13. **Bảng tin lớp học**
14. **Bảng tin nhà trường**
15. **Thông báo**
16. **Hồ sơ cá nhân**
17. **Đăng nhập**

## E. Nhóm mở rộng cho nhà trường

18. **Lớp học trực tuyến**
19. **Lớp học thông minh**
20. **Sổ liên lạc điện tử**
21. **Giáo án số**
22. **Sổ chủ nhiệm**
23. **Sổ đầu bài**
24. **Quản lý nề nếp**
25. **Quản lý bán trú**
26. **Thẻ và điểm danh**

---

# 4.2. Các module hiện có của `prj-gkebook-shadcn-frontend`

## A. Nhóm tài nguyên học tập

1. **Bộ sách**
   - Danh sách sách
   - Chi tiết sách
   - Đọc sách
   - Xem trước sách
   - Chỉnh sửa sách

2. **Kho học liệu**
   - Danh sách học liệu
   - Xem thêm học liệu
   - Xem chi tiết học liệu
   - Chỉnh sửa học liệu

3. **Thư viện của tôi**
   - Thư viện học liệu
   - Thư viện câu hỏi
   - Học liệu yêu thích
   - Học liệu đã tải lên

4. **Hồ sơ luyện tập**
   - Danh sách hồ sơ luyện tập
   - Chi tiết hồ sơ luyện tập

## B. Nhóm học và luyện tập

5. **Tự học**
6. **Ngân hàng câu hỏi**
7. **Xưởng tạo đề (Quiz Studio)**
   - Chọn kiểu tạo đề
   - Tạo đề từ kho học liệu
   - Tạo đề thủ công
   - Có lối vào cho “tạo bằng AI” nhưng chưa mở đầy đủ

8. **Làm bài / Xem trước / Luyện đề**
   - Biểu mẫu đề
   - Xem trước đề
   - Luyện đề
   - Học theo đề

## C. Nhóm hỗ trợ và người dùng

9. **Trang chủ**
10. **Tin tức**
11. **Câu hỏi thường gặp**
12. **Hồ sơ cá nhân**
13. **Báo cáo hoạt động cá nhân**
14. **Đăng nhập / Đăng ký / Quên mật khẩu / Kiểm tra SSO**

## D. Nhóm kinh doanh và tiện ích

15. **Gói cước**
16. **Chi tiết gói cước**
17. **Trang thẻ thông minh điện tử**
18. **Chính sách bảo mật**
19. **Điều khoản sử dụng**
20. **Màn hình xem học liệu dạng nhúng (web-view)**
21. **Bộ công cụ toàn cục**

---

## 5. Module chung giữa hai project và gợi ý gộp

### 5.1. Module chung (đều có ở cả 2 bên)

- **Trợ lý học tập AI / Chatbot toàn cục**
- **Kho học liệu**
- **Bộ sách / Đọc sách**
- **Ngân hàng câu hỏi**
- **Tự học**
- **Tin tức**
- **Thư viện của tôi**
- **Tạo đề / Xưởng tạo đề (Quiz Studio)** (mức độ khác nhau nhưng cùng hướng)

### 5.2. Nhóm rất giống nhau, nên tính chuyện gộp sớm

#### 1) Trợ lý học tập AI / Chatbot toàn cục
Cả hai bên đều có:
- chatbot nổi trên giao diện
- đổi chế độ theo ngữ cảnh màn hình
- tạo cuộc trò chuyện mới
- gửi câu hỏi theo luồng phản hồi stream
- gắn nhãn câu trả lời là nội dung do AI tạo

**Nên gộp trước**:
- kiểu dữ liệu AI
- dịch vụ gọi API AI
- xử lý stream phản hồi
- trạng thái chatbot
- hiển thị câu trả lời
- bộ xác định ngữ cảnh AI theo màn hình

#### 2) Kho học liệu
Hai bên đều có:
- danh sách học liệu
- chi tiết học liệu
- chỉnh sửa học liệu
- thư viện học liệu cá nhân

Cụm này phù hợp chuẩn hóa UI và logic dùng chung.

#### 3) Bộ sách / Đọc sách
Hai bên đều có:
- danh sách sách
- chi tiết sách
- đọc sách
- kết nối với chatbot theo ngữ cảnh sách

Có thể gộp phần “lõi đọc sách” và “AI theo ngữ cảnh sách” trước, sau đó mới tính UI đặc thù.

#### 4) Tạo đề / Xưởng tạo đề (Quiz Studio)
Hai bên đều có “lối vào tạo đề bằng AI” nhưng hiện đều đang ở trạng thái chưa mở trọn.
Có thể gom phần chọn nguồn tạo đề và các thao tác AI trong màn soạn câu hỏi.

### 5.3. Nhóm chưa nên gộp ngay

#### Project school (chưa nên gộp sớm)
- lớp học
- điểm danh
- thời khóa biểu
- nhiệm vụ
- báo cáo học sinh
- quản trị nhà trường

Lý do: gắn chặt với nghiệp vụ vận hành trường.

#### Project shadcn (chưa nên gộp sớm)
- gói cước
- câu hỏi thường gặp
- hồ sơ luyện tập
- bộ công cụ toàn cục
- web-view học liệu

Lý do: thiên về trải nghiệm cá nhân và kinh doanh.

---

## 6. Từng module đã có AI chưa?

# 6.1. AI trong `gkebook-school-frontend`

## A. Đã có AI rõ ràng

1) **Trợ lý học tập AI / Chatbot toàn cục**
- Có AI.
- Có các ngữ cảnh: chế độ tự do, chế độ sách số, chế độ đọc sách, chế độ học theo bài kiểm tra.

2) **Dịch vụ AI nền**
- Có AI.
- Có tạo hội thoại mới, gửi tin nhắn dạng stream, tạo tiêu đề hội thoại, trích từ khóa, chuyển markdown sang cấu trúc editor.

3) **Gắn nhãn nội dung do AI tạo**
- Có AI.
- UI gắn nhãn “Tạo bởi AI” cho câu trả lời.

4) **Đánh giá học liệu của giáo viên bằng AI**
- Có AI.
- AI được dùng để đánh giá mức độ tích cực tạo học liệu của giáo viên ở trang chủ.

## B. Đã đặt chỗ AI nhưng chưa mở hoàn toàn

5) **Tạo đề bằng AI**
- Có chỗ “AI Gen” và mô tả hỗ trợ tạo nội dung tự động, nhưng hiện còn khóa/đang phát triển.

6) **Các thao tác AI trong màn soạn câu hỏi**
- Có menu như: thêm câu hỏi tương tự, thêm giải thích, thêm tùy chọn.
- Hiện các mục này chưa bật.

## C. Chưa thấy AI rõ trong lần rà này

- lớp học
- điểm danh
- thời khóa biểu
- thông báo
- bảng tin
- hồ sơ cá nhân
- quản trị nhà trường

---

# 6.2. AI trong `prj-gkebook-shadcn-frontend`

## A. Đã có AI rõ ràng

1) **Trợ lý học tập AI / Chatbot toàn cục**
- Có AI.
- Có đổi chế độ theo ngữ cảnh: sách số, đọc sách, học theo bài kiểm tra, chế độ tự do.

2) **Chatbot ở màn ngoài / màn giới thiệu**
- Có AI.
- Có hook riêng cho chatbot ngoài một số ngữ cảnh chính.

3) **Tóm tắt tin tức bằng AI**
- Có AI.
- Hiển thị phần “Nội dung do AI tóm tắt”.

4) **Chuyển văn bản thành giọng nói**
- Có AI.
- Có hook gọi API chuyển chữ thành âm thanh.

5) **Phân tích tài liệu để tạo câu hỏi**
- Có AI.
- Có hook dùng backend AI để đọc tài liệu và parse câu hỏi.

6) **Hạn mức sử dụng trợ lý AI**
- Có AI.
- Có UI/logic theo dõi quota dùng chatbot.

## B. Đã đặt chỗ AI nhưng chưa mở hoàn toàn

7) **Tạo đề bằng AI**
- Có mô tả “tạo nội dung tự động bằng AI”, nhưng khi vào màn tạo đề vẫn đang báo phát triển.

---

## 7. Gợi ý AI thêm cho từng project

# 7.1. Gợi ý cho `gkebook-school-frontend`

1) **AI tóm tắt báo cáo học sinh cho giáo viên**
- Tóm tắt tình hình 7 ngày/30 ngày
- Nêu điểm đáng chú ý
- Gợi ý câu nhận xét gửi phụ huynh

2) **AI gợi ý giao nhiệm vụ theo lớp**
- Dựa trên kho học liệu và môn/lớp
- Gợi ý bài phù hợp, mức độ, thời gian

3) **AI hỗ trợ chấm tự luận**
- Gợi ý rubric, nhận xét
- Đánh dấu ý chính học sinh đã có/chưa có

4) **AI phát hiện tín hiệu cần để ý từ điểm danh và tiến độ học**
- Nghỉ tăng bất thường, nộp chậm liên tục…

5) **AI tạo đề từ học liệu của lớp đang dạy**
- Bật theo hướng “đọc tài liệu → sinh câu hỏi → tạo đề”.

# 7.2. Gợi ý cho `prj-gkebook-shadcn-frontend`

1) **AI huấn luyện viên học tập cho hồ sơ luyện tập**
- Nói điểm mạnh/yếu theo chủ đề
- Gợi ý 3 bước học tiếp theo

2) **AI giải thích đoạn đang đọc trong sách**
- Bôi đen đoạn → hỏi “giải thích dễ hơn/cho ví dụ/rút ý chính”.

3) **AI tạo thẻ ghi nhớ từ bài học**
- Ý chính, công thức, câu hỏi ôn tập.

4) **AI tạo bài luyện ngay sau khi đọc xong tài liệu**
- 5–10 câu theo mức độ.

5) **AI giọng nói cho hỗ trợ học**
- Đọc câu hỏi/đọc tóm tắt, nghe khi học trên di động.

---

## 8. Kết luận

1) Hai project trùng nhau nhiều nhất ở: **học liệu – sách – trợ lý AI – tạo đề – tự học**.
2) Nên gộp trước phần “lõi trợ lý AI” và “lõi tạo đề”, vì đây là nơi trùng lặp cao và dễ lệch hành vi nếu để tách lâu.
3) Các module “đặc thù nhà trường” (lớp học/điểm danh/nhiệm vụ…) nên để riêng thêm, sau đó mới tính chuẩn hóa dần.

---

## 9. Danh sách file tham chiếu chính (để team dev đối chiếu nhanh)

### `gkebook-school-frontend`
- `src/lib/constants/index.tsx`
- `src/components/Chatbot/GlobalChatbot.tsx`
- `src/components/Chatbot/Answer.tsx`
- `src/hooks/use-chat-with-assistant.tsx`
- `src/services/ai/index.ts`
- `src/mock-data/sidebar.tsx`
- `src/mock-data/quiz-studio/quiz-source.tsx`
- `src/components/Quiz/Create/SourceTypes/index.tsx`
- `src/app/(dashboard)/trang-chu/_components/Actions/getAiTeacherMaterialEvaluate.tsx`

### `prj-gkebook-shadcn-frontend`
- `src/constants/index.ts`
- `src/mock-data/sidebar.tsx`
- `src/components/global-chatbot/index.tsx`
- `src/components/chatbot/Answer.tsx`
- `src/hooks/use-chat-with-assistant.tsx`
- `src/hooks/use-chat-with-assistant-landing.tsx`
- `src/services/ai/index.ts`
- `src/hooks/use-question-banks-streamming.tsx`
- `src/hooks/use-text-to-speech.tsx`
- `src/app/(news)/tin-tuc/[slug]/components/NewsSummary.tsx`
- `src/components/quiz-studio/quiz-types/index.tsx`
