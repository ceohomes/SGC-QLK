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
