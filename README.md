# SGC | Báo Cáo Giao Nhận

Webapp quản lý báo cáo theo dõi giao nhận vật tư, máy móc thiết bị.

## Tính năng

- **Tab Đơn Giao**: Import file Excel Report_Orders_Đơn giao, xem & lọc dữ liệu
- **Tab Đơn Nhận**: Import file Excel Report_Orders_Đơn nhận, xem & lọc dữ liệu  
- **Cấu hình dữ liệu**: *(sắp ra mắt)*
- **Tổng hợp**: *(sắp ra mắt)*

## Cài đặt & Chạy

```bash
npm install
npm run dev
```

## Build & Deploy

```bash
npm run build
```

Sau đó deploy thư mục `dist/` lên Cloudflare Pages.

## Deploy lên Cloudflare Pages

1. Push code lên GitHub
2. Vào [Cloudflare Pages](https://pages.cloudflare.com/)
3. Connect repository
4. Build settings:
   - **Build command**: `npm run build`
   - **Build output directory**: `dist`
5. Deploy!

## Cấu trúc file Excel yêu cầu

File Excel cần có sheet **"Báo cáo theo dõi giao nhận"** với 36 cột:
- Dòng 1: Nhóm cột (Nhập kho, Xuất kho, Thông tin chung)
- Dòng 2: Sub-headers
- Dòng 3+: Dữ liệu

## Changelog — Bản sửa lỗi & tối ưu tốc độ (01/07/2026)

**1. Sửa lỗi tải trùng dữ liệu 2 lần gây chậm >30s (nguyên nhân chính)**
Trước đây `loadTableFromSupabase` tải song song mỗi trang 10.000 dòng, nhưng Supabase/PostgREST mặc định chỉ trả tối đa 1.000 dòng/request. Kết quả là mọi trang đều bị cắt còn 1.000 dòng, code phát hiện ra rồi **huỷ toàn bộ và tải lại tuần tự từ đầu** — tức tải dữ liệu 2 lần. Đã sửa `PAGE_SIZE` về đúng 1.000 để khớp giới hạn thật của server, tải song song thành công ngay lần đầu.

**2. Sửa lỗi "Smart Sync" so sánh sai dữ liệu với bảng >1.000 dòng**
Bước so sánh nhẹ (id + timestamp) để phát hiện thay đổi cũng bị giới hạn 1.000 dòng do thiếu phân trang, khiến các dòng ngoài 1.000 dòng đầu bị hiểu nhầm là "đã xoá trên server". Đã thêm phân trang đầy đủ cho bước này.

**3. Giảm spam lỗi 400 từ bảng `sgc_summary_configs` chưa tồn tại**
Nếu bảng `sgc_summary_configs` chưa được tạo trên Supabase, trước đây app sẽ gửi lại request thất bại mỗi lần chuyển tab / đồng bộ ngầm. Giờ chỉ thử 1 lần/phiên, chỉ kiểm tra lại khi bấm nút "Làm mới" thủ công.
→ **Bạn cần tạo bảng này** bằng cách mở app, tìm banner vàng "Thiếu bảng sgc_summary_configs trong Supabase", bấm copy SQL và chạy trong Supabase SQL Editor (tính năng này đã có sẵn trong app, không phải code mới).

### Lưu ý thêm (chưa sửa, để tránh rủi ro)
Bundle `xlsx` (thư viện đọc/xuất Excel) hiện nặng ~1MB và được tải ngay từ đầu dù chưa dùng đến, cũng góp phần làm chậm lần tải đầu tiên. Có thể tối ưu bằng cách chuyển sang `dynamic import()` chỉ tải khi người dùng bấm nút Import/Export Excel — nhưng việc này động vào ~26 vị trí sử dụng trong file xuất báo cáo, nên chưa tự ý sửa để tránh làm hỏng tính năng xuất Excel. Báo lại nếu bạn muốn mình làm tiếp phần này.
