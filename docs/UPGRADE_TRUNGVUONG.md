# Kế hoạch nâng cấp THCS Trần Hưng Đạo theo Trưng Vương
Ngày rà soát: 2026-11-23

## Kiến trúc hiện tại (đã xác minh trong mã nguồn)
- THD: `app.js`, IndexedDB `so-thu-hoc-sinh-local`, schema version 1, mã lưu riêng.
- TV: IndexedDB `so-thu-hoc-sinh-trungvuong-local`, schema version 2, nhiều module mở rộng.
- Hai bản có cùng gốc SchoolCollect nhưng TV có các module giao diện, workflow, QR và bảo hiểm bổ sung.
- **Không thay thế toàn bộ app.js/index.html của THD bằng bản TV**: sẽ đổi schema, mapping, tên DB và có nguy cơ mất khả năng đối soát dữ liệu cũ.

## Phạm vi nâng cấp
| Nhóm | Tác vụ | Độ ưu tiên | Điều kiện |
|---|---|---|---|
| Quản lý khoản thu | Khoản thu đa kỳ theo trường/khối/lớp/HS, tùy chỉnh số tiền | P0 | Không đổi các khoản đang thu |
| Học sinh | Bổ sung/sửa/ngừng theo dõi, phát hiện trùng và chuyển lớp | P0 | Không thay mã định danh đang dùng |
| Thu tiền | Tiền mặt, xác nhận theo phiên, chứng từ, kiểm tra thu trùng | P0 | Có log và duyệt |
| Đối soát | Import BIDV Direct, ghi nhận một phần, xử lý sai lệch | P0 | Giữ quy tắc BHYT/BHTT |
| Thống kê | Theo lớp/khoản/kỳ/ngày, danh sách đã thu/chưa thu | P0 | Tổng chi tiết = tổng toàn trường |
| QR | QR một khoản/nhiều khoản, A4/Zalo, file hàng loạt | P1 | Kiểm thử VietQR/Bank QR thực tế |
| Xuất dữ liệu | Bảng kê BIDV, báo cáo BHYT/BHTT, Excel | P1 | Mẫu bảng kê xác thực |
| Phụ huynh | Bản nháp JSON, đồng bộ có phê duyệt | P2 | Backend an toàn, quyền truy cập |
| PWA | Cache/versioning, offline, nâng cấp và tự làm mới | P1 | Không mất IndexedDB |

## Thứ tự thực hiện
1. Chụp mẫu dữ liệu giả lập và xuất .sctbackup; đo tổng học sinh, số khoản, phải thu, đã thu, còn thiếu.
2. Tạo bản thử trong nhánh này, thử với dữ liệu giả lập và sao lưu mẫu. Giữ tên DB riêng trường.
3. Đưa cấu phần khoản thu đa kỳ + import/đối soát sang bản THD (viết adapter cho định danh mã HS cũ).
4. Bổ sung thu thủ quỹ, phiếu thu, báo cáo và xuất mẫu BIDV.
5. Bổ sung QR nhiều món và bản nháp đồng bộ phụ huynh.
6. Kiểm thử các trường hợp: 2 dòng BHYT/BHTT / hóa đơn đã hủy / chuyển lớp / thu một phần / trùng giao dịch / sai nội dung / khôi phục dữ liệu / trình duyệt mất mạng.
7. Sau khi xác nhận số liệu kiểm thử khớp và có bản sao lưu, mới merge vào main.

## Bất biến dữ liệu
- Không đổi `DB_NAME` hoặc loại bỏ store hiện có.
- Không ghi học sinh thật, sao kê, token hay bản backup lên GitHub.
- Tổng tiền phải thu - tiền đã nhận hợp lệ = còn phải thu.
- Một mã giao dịch ngân hàng chỉ ghi nhận thu một lần; giao dịch không rõ khoản giữ trạng thái chờ xác minh.
- Không tự ý phân bổ một giao dịch gộp sang nhiều món khi chưa xác định được mã tham chiếu.
- Chỉ công bố số đã thu lên cổng phụ huynh sau khi lô được duyệt.

## Tình trạng
Tài liệu tiền triển khai. Chưa đưa các chức năng mới vào mã chạy, chưa triển khai production.
