# Sổ thu học sinh

Ứng dụng web tĩnh, không cần tài khoản ChatGPT. Mã nguồn có thể được lưu trong GitHub và chạy trên GitHub Pages; danh sách học sinh, giao dịch, lịch sử và bản sao lưu **không được đưa vào repository**.

## Thiết kế dữ liệu

- Trang web đọc tệp ngay trong trình duyệt; không có API gửi tệp lên máy chủ và không dùng analytics/CDN.
- Dữ liệu ứng dụng nằm trong IndexedDB của hồ sơ trình duyệt hiện tại. Mỗi máy/hồ sơ trình duyệt có dữ liệu riêng.
- Có thể dùng ngoại tuyến sau lần mở web đầu tiên khi service worker đã lưu bộ giao diện.
- Có thể xuất bản sao lưu `.sctbackup` mã hóa AES-GCM bằng mật khẩu; mật khẩu không được lưu trong web.
- GitHub Pages chỉ nên chứa mã nguồn và tài sản giao diện. Không commit tệp học sinh, sao kê, bản sao lưu hoặc tệp cấu hình có dữ liệu thật.

## Nhập file hiện có

- Đọc tệp `.xlsx` (trang tính đầu tiên) trực tiếp trên máy bằng bộ đọc ZIP/XML tích hợp; không tải thư viện từ CDN.
- Đọc `.csv`, `.tsv`, `.txt`, hiển thị bản xem trước và cho phép ghép cột.
- Danh sách học sinh: ghép mã học sinh, họ tên, lớp, khoản phải thu; nhập lại cùng mã sẽ cập nhật học sinh đó.
- Báo cáo ngân hàng: ghép số tiền, mã học sinh, mã giao dịch, ngày, nội dung; giao dịch trùng mã tham chiếu được bỏ qua. Nếu không có mã học sinh, giao dịch được giữ ở trạng thái chưa khớp.

## Phần cần cấu hình sau khi nhận file mẫu

1. Cố định dòng tiêu đề, tên sheet, cột cần lấy và cách đọc ngày/tiền trong file trường.
2. Cố định cột định danh học sinh trong nội dung giao dịch ngân hàng và quy tắc ghép chính xác.
3. Xác định cách tính khi một học sinh có nhiều khoản thu, miễn giảm, hoàn tiền hoặc giao dịch đảo.
4. Kiểm thử bằng bản sao file đã ẩn thông tin nhạy cảm; đối chiếu tổng dòng và tổng tiền với báo cáo gốc.

Bản nền không có đăng nhập, phân quyền kế toán/hiệu trưởng hoặc đồng bộ tự động giữa các máy. Chỉ dùng dữ liệu thật sau khi nhà trường chốt quy trình bảo vệ máy, sao lưu, và kiểm tra đối soát trên file mẫu.

## Chạy thử

Mở bằng một máy chủ web tĩnh trên HTTPS (GitHub Pages phù hợp). Tránh mở trực tiếp bằng `file://`, vì trình duyệt có thể chặn IndexedDB/service worker. Với GitHub Pages của project repository, đặt `index.html` ở thư mục gốc của nhánh được chọn làm nguồn Pages.
