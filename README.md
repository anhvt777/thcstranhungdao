# SchoolCollect · Quản lý thu học sinh

Ứng dụng web tĩnh để theo dõi sĩ số, khoản phải thu và kết quả đối soát. Giao diện tối ưu cho máy tính kế toán và hiệu trưởng, đồng thời dùng được trên điện thoại.

## Quyền riêng tư

- Đọc file ngay trong trình duyệt; không có API gửi file lên máy chủ và không dùng analytics/CDN.
- Học sinh, khoản thu, giao dịch và lịch sử được lưu trong IndexedDB của trình duyệt trên thiết bị đang sử dụng.
- Các máy không tự đồng bộ với nhau. Có thể chuyển dữ liệu bằng tệp sao lưu `.sctbackup` được mã hóa AES-GCM bằng mật khẩu.
- GitHub Pages chỉ lưu mã nguồn và giao diện. Không commit danh sách học sinh, sao kê hoặc tệp sao lưu có dữ liệu thật.
- Có thể dùng ngoại tuyến sau lần mở web đầu tiên khi service worker đã lưu giao diện.

## Nhập dữ liệu

- Hỗ trợ `.xlsx`, `.csv`, `.tsv` và `.txt`; trong Excel nhiều trang tính, tự chọn trang có tiêu đề phù hợp và nhiều dòng dữ liệu nhất. Không tải thư viện từ CDN.
- Danh sách theo mẫu trường có hai dòng cho mỗi học sinh. Web gộp theo “Mã học sinh”, lưu riêng mã “Mã HS theo Khoản nộp”, loại khoản BHYT/BHTT và số tiền; kiểm tra trùng mã và thiếu khoản trước khi lưu.
- Báo cáo ngân hàng: nhận diện “Mã khách hàng”, “Số hóa đơn”, “Ngày giao dịch”, “Tên khách hàng”, “Số tiền” và “Trạng thái giao dịch”. Mã khách hàng đuôi YT được ghép với BHYT, đuôi TT được ghép với BHTT. Trạng thái không thành công không được tính đã thu.
- Dịch vụ khác có thể xem chi tiết theo nội dung chuyển khoản như gửi xe, nước uống. Mã học sinh không tìm thấy sẽ hiện là chưa khớp.
- Một giao dịch chỉ được tính đã thu khi mã khách hàng/mã khoản, loại khoản, trạng thái thành công và số tiền khớp chính xác với một món phải thu chưa được ghép. Giao dịch sai số tiền, sai khoản, trùng món hoặc không tìm thấy học sinh không cộng vào số đã thu.
- Các mẫu CSV/XLSX cũ với mã học sinh và cột khoản thu vẫn được hỗ trợ qua bước ghép cột.
- Tạo mã QR: lưu một lần mã BIN, số tài khoản và tên chủ tài khoản của trường; sau đó tạo ảnh QR riêng cho từng món phải thu, lọc theo lớp và tải ZIP. Nội dung thanh toán được tạo thành ảnh QR và đóng gói ngay trong trình duyệt bằng thư viện cục bộ.
- QR mặc định chỉ tạo cho món chưa có giao dịch khớp chính xác. Sau khi cập nhật báo cáo thu, danh sách QR được tính lại. Hãy quét thử bằng ứng dụng ngân hàng trước khi gửi cho phụ huynh.

## Triển khai

Mở qua máy chủ web tĩnh HTTPS, chẳng hạn GitHub Pages. Tránh mở trực tiếp bằng `file://` vì trình duyệt có thể chặn IndexedDB và service worker. Với GitHub Pages của repository, đặt `index.html` ở thư mục gốc của nhánh được chọn làm nguồn Pages.

Trước khi dùng dữ liệu thật, hãy đối chiếu tên cột và cách ghi mã học sinh trong file gốc của trường/ngân hàng; kiểm tra tổng số dòng và tổng tiền với báo cáo nguồn.
