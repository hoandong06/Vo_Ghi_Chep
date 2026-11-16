# 01_Project_Proposal: Chatbot hướng dẫn sử dụng các hàm Excel cơ bản

-  Nhóm: 1
- Lớp: 75DCTT26
- Giảng viên: Nguyễn Văn Cường
- Giai đoạn 1:Phân tích và Dữ liệu
- Tuần 1: Khởi động dự án và phân tích yêu cầu

## 1. Giới thiệu
Người mới học Excel thường biết mình cần làm gì nhưng không nhớ hàm nào, cú pháp ra sao, hoặc không hiểu vì sao công thức báo lỗi. Đề tài xây dựng **chatbot hỏi đáp tiếng Việt** giúp người dùng hiểu và dùng đúng **32 hàm Excel cơ bản**, trả lời dựa trên kho tài liệu do nhóm soạn và kiểm chứng.

**Mục tiêu:** trong 10 tuần hoàn thành một chatbot chạy được từ đầu đến cuối, trả lời chính xác các câu hỏi thường gặp về 32 hàm cơ bản.

## 2. Bài toán và người dùng
- **Người dùng chính:** sinh viên và nhân viên văn phòng dùng Excel ở mức cơ bản.
- **Bài toán thực tế:** tra cứu hàm Excel mất thời gian, tài liệu dài hoặc bằng tiếng Anh, khó tìm đúng ví dụ.
- **Giá trị:** trả lời ngắn gọn, đúng cú pháp, có ví dụ chép thử được.

## 3. Phạm vi (trong 10 tuần)
**Trong phạm vi:** 32 hàm chia thành 5 nhóm. Với mỗi hàm chatbot trả lời 4 phần: *công dụng, cú pháp, 1 ví dụ, lỗi thường gặp*.

| Nhóm | Các hàm |
|---|---|
| Toán và thống kê | SUM, AVERAGE, MIN, MAX, COUNT, COUNTA, ROUND |
| Điều kiện | IF, IFS, AND, OR, IFERROR, COUNTIF, SUMIF, AVERAGEIF |
| Tra cứu | VLOOKUP, HLOOKUP, XLOOKUP, INDEX, MATCH |
| Văn bản | LEFT, RIGHT, MID, LEN, TRIM, CONCAT, UPPER, LOWER |
| Ngày giờ | TODAY, NOW, DATE, DATEDIF |

**Ngoài phạm vi (không làm):** PivotTable, biểu đồ, Power Query, VBA/macro, các hàm nâng cao (mảng động, LAMBDA...), viết công thức riêng cho bảng dữ liệu của người dùng. Câu hỏi thuộc nhóm này chatbot trả lời "chưa hỗ trợ".

## 4. Dữ liệu
- **Nguồn:** trang hỗ trợ của Microsoft, giáo trình Tin học văn phòng, các bài hướng dẫn công khai. Nhóm **tự soạn lại bằng lời của mình** và ghi nguồn tham khảo.
- **Cấu trúc:** mỗi hàm là một mục gồm *tên hàm, nhóm, công dụng, cú pháp, ví dụ, kết quả, lỗi thường gặp, phiên bản Excel yêu cầu*.
- **Quy mô:** 32 mục (khoảng 130 đoạn nội dung ngắn), một thành viên vai trò Data làm được trong vài tuần.
- **Kiểm chứng:** mọi ví dụ được chạy thử trong Excel trước khi đưa vào kho.
- **Lưu ý:** IFS cần Excel 2019 hoặc Microsoft 365; XLOOKUP cần Excel 2021 hoặc Microsoft 365; dấu phân cách đối số có thể là dấu phẩy hoặc chấm phẩy tùy cài đặt vùng.

## 5. Phân tích yêu cầu
### 5.1 Sơ đồ Use Case
![Sơ đồ Use Case](usecase_diagram.png)

| Mã | Use Case | Tác nhân | Mô tả | User Story |
|---|---|---|---|---|
| UC1 | Đặt câu hỏi về hàm Excel | Người dùng | Gõ câu hỏi tiếng Việt; hệ thống tìm trong kho tài liệu và trả lời. Câu hỏi ngoài phạm vi sẽ nhận thông báo "chưa hỗ trợ". | US1, US6 |
| UC2 | Xem cú pháp và ví dụ của hàm | Người dùng | Xem công dụng, cú pháp và 1 ví dụ có thể chép thử vào Excel. | US2 |
| UC3 | Tra cứu lỗi thường gặp | Người dùng | Hỏi vì sao công thức báo lỗi (#N/A, #VALUE!, #DIV/0!...) và cách sửa. | US3 |
| UC4 | Duyệt danh sách hàm theo nhóm | Người dùng | Xem 32 hàm chia theo 5 nhóm để chọn hàm phù hợp. | US4 |
| UC5 | So sánh hai hàm | Người dùng | So sánh các cặp dễ nhầm như COUNT/COUNTA, VLOOKUP/XLOOKUP. | US5 |
| UC6 | Cập nhật nội dung hàm | Quản trị viên | Thêm, sửa nội dung một hàm trong kho tài liệu mà không cần sửa code. | US7 |
| UC7 | Xem câu hỏi chưa trả lời được | Quản trị viên | Xem danh sách câu hỏi ngoài phạm vi hoặc bị chấm sai để bổ sung dữ liệu. | US8 |

### 5.2 User Stories
| Mã | Là một | Tôi muốn | Để |
|---|---|---|---|
| US1 | Người dùng | hỏi bằng tiếng Việt tự nhiên về một hàm | biết hàm đó làm gì mà không phải đọc tài liệu dài |
| US2 | Người dùng | xem cú pháp và một ví dụ cụ thể | chép thử ngay vào bảng tính của mình |
| US3 | Người dùng | hỏi vì sao công thức báo lỗi | sửa lỗi nhanh, không mất thời gian dò từng ô |
| US4 | Người dùng | duyệt danh sách hàm theo nhóm | tìm ra hàm phù hợp khi chưa biết tên hàm |
| US5 | Người dùng | so sánh hai hàm dễ nhầm | chọn đúng hàm cho bài toán của mình |
| US6 | Người dùng | được báo rõ khi câu hỏi nằm ngoài phạm vi | không bị chatbot trả lời sai hoặc bịa công thức |
| US7 | Quản trị viên | cập nhật nội dung của từng hàm | chatbot luôn trả lời đúng và đầy đủ |
| US8 | Quản trị viên | xem các câu hỏi chatbot chưa trả lời được | biết cần bổ sung dữ liệu nào ở các tuần sau |

## 6. Danh sách 15 câu hỏi mẫu đại diện
| STT | Nhóm | Câu hỏi | Ý chính của câu trả lời đúng |
|---|---|---|---|
| 1 | Toán | Hàm SUM dùng để làm gì? Cho ví dụ cộng vùng A1:A10. | Cộng các giá trị số; =SUM(A1:A10). |
| 2 | Toán | Cách tính điểm trung bình của cột B2:B30? | =AVERAGE(B2:B30); bỏ qua ô trống và ô chứa chữ. |
| 3 | Toán | Làm tròn số 3,14159 đến 2 chữ số thập phân như thế nào? | =ROUND(3.14159,2) cho kết quả 3.14. |
| 4 | Toán | COUNT và COUNTA khác nhau chỗ nào? | COUNT chỉ đếm ô chứa số; COUNTA đếm mọi ô không trống (cả chữ). |
| 5 | Điều kiện | Viết công thức xếp loại "Đạt" nếu điểm từ 5 trở lên, ngược lại "Không đạt". | =IF(A2>=5,"Đạt","Không đạt"). |
| 6 | Điều kiện | Khi nào nên dùng IFS thay cho IF lồng nhau? | Khi có nhiều mức điều kiện (Giỏi, Khá, Trung bình...); IFS viết gọn hơn. Cần Excel 2019 hoặc Microsoft 365. |
| 7 | Điều kiện | Đếm số ô trong C2:C50 có điểm lớn hơn 8. | =COUNTIF(C2:C50,">8"). |
| 8 | Điều kiện | Tính tổng doanh thu của riêng sản phẩm "Bút". | =SUMIF(vùng tên sản phẩm,"Bút",vùng doanh thu). |
| 9 | Tra cứu | VLOOKUP dùng như thế nào? Giải thích các đối số. | =VLOOKUP(giá trị cần tìm, bảng, số thứ tự cột, 0); 0 hoặc FALSE là khớp chính xác. |
| 10 | Tra cứu | Vì sao VLOOKUP báo lỗi #N/A? | Không tìm thấy giá trị: sai chính tả, thừa khoảng trắng, khác kiểu dữ liệu (số và chữ) hoặc thiếu khớp chính xác; có thể dùng TRIM hoặc IFERROR. |
| 11 | Tra cứu | XLOOKUP khác VLOOKUP thế nào? | XLOOKUP tra cứu được cả bên trái, mặc định khớp chính xác, không cần số thứ tự cột; chỉ có từ Excel 2021 và Microsoft 365. |
| 12 | Tra cứu | Cách kết hợp INDEX và MATCH để tra cứu? | =INDEX(cột kết quả, MATCH(giá trị cần tìm, cột dò, 0)). |
| 13 | Văn bản | Tách 2 ký tự đầu "SV" từ mã "SV001" trong ô A2. | =LEFT(A2,2) cho kết quả "SV". |
| 14 | Văn bản | Làm sao xóa khoảng trắng thừa trong họ tên? | =TRIM(A2): bỏ khoảng trắng đầu, cuối và giữ một dấu cách giữa các từ. |
| 15 | Ngày giờ | Tính số tuổi từ ngày sinh ở ô A2. | =DATEDIF(A2,TODAY(),"Y"). |

**Câu hỏi kiểm thử ngoài phạm vi** (chatbot phải trả lời "chưa hỗ trợ"):
- PivotTable là gì và tạo như thế nào?
- Viết macro VBA để tự động in báo cáo.
- Viết công thức riêng cho bảng dữ liệu tôi đang có.

## 7. Tổ chức nhóm
| Vai trò | Họ tên - MSSV | Nhiệm vụ chính |
|---|---|---|
| Leader | [điền tên] | Điều phối nhóm, chia việc, họp hàng tuần, theo dõi tiến độ, viết và nộp tài liệu, liên hệ giảng viên. |
| Data | [điền tên] | Gom và soạn lại nội dung 32 hàm theo mẫu 4 phần, chạy thử từng ví dụ trong Excel, mở rộng bộ câu hỏi mẫu. |
| AI | [điền tên] | Chọn mô hình, thiết kế prompt, xây cơ chế truy xuất từ kho tài liệu, giới hạn câu trả lời trong phạm vi dữ liệu. |
| Backend | [điền tên] | Xây API nhận câu hỏi và gọi AI, lưu câu hỏi chưa trả lời được, làm trang quản trị cập nhật nội dung. |
| Frontend/QA | [điền tên] | Làm giao diện chat và trang danh sách hàm; kiểm thử bằng bộ câu hỏi mẫu, ghi lỗi, đo độ chính xác. |

## 8. Kế hoạch 
| Tuần | Nội dung | Kết quả |
|---|---|---|
| 1 | Khởi động và phân tích yêu cầu | Chốt đề tài, Use Case, User Stories, 15 câu hỏi mẫu, chia vai trò |
| 2 | Thu thập dữ liệu | Gom tài liệu cho 32 hàm, chốt mẫu nội dung 4 phần |
| 3 | Chuẩn hóa dữ liệu | Soạn xong nội dung 32 hàm, chạy thử ví dụ, mở rộng lên 50 câu hỏi |
| 4 | Thiết kế và nguyên mẫu AI | Chọn mô hình, dựng bản chatbot đầu tiên trả lời từ kho tài liệu |
| 5 | Backend | API hỏi đáp, lưu lịch sử và câu hỏi ngoài phạm vi |
| 6 | Frontend | Giao diện chat, danh sách hàm theo nhóm |
| 7 | Tích hợp | Nối giao diện, API và AI; chạy được từ đầu đến cuối |
| 8 | Kiểm thử và đánh giá | Chạy bộ 50 câu hỏi, thống kê đúng sai, ghi lỗi |
| 9 | Cải thiện | Sửa lỗi, bổ sung dữ liệu từ câu hỏi chưa trả lời được |
| 10 | Hoàn thiện và báo cáo | Quay demo, viết báo cáo cuối, chuẩn bị bảo vệ |

## 9. Mục tiêu
- Đúng từ 80% trở lên trên bộ 15 câu hỏi mẫu ở bản chạy đầu tiên và trên bộ 50 câu ở Tuần 8 (đề xuất, nhóm có thể điều chỉnh).
- 100% ví dụ công thức trong kho tài liệu đã được chạy thử trong Excel.
- Câu hỏi ngoài phạm vi luôn nhận thông báo "chưa hỗ trợ", không bị trả lời bịa.

## 10. Rủi ro và xử lý
| Rủi ro | Cách xử lý |
|---|---|
| Chatbot trả lời sai hoặc bịa công thức | Chỉ trả lời dựa trên kho tài liệu nhóm soạn; mọi ví dụ đều được chạy thử trong Excel. |
| Phạm vi phình to giữa chừng | Khóa danh sách 32 hàm; muốn thêm hàm sau Tuần 2 phải họp nhóm quyết định. |
| Bản quyền tài liệu nguồn | Nhóm tự soạn lại bằng lời của mình, chỉ tham khảo và ghi rõ nguồn. |
| Khác biệt phiên bản và cài đặt Excel | Ghi chú hàm nào cần bản mới (IFS, XLOOKUP) và lưu ý dấu phân cách đối số có thể là dấu phẩy hoặc chấm phẩy tùy máy. |
| Thành viên chậm tiến độ | Họp nhóm hàng tuần, chia việc nhỏ theo tuần, Leader theo dõi bảng công việc. |


