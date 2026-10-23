# HRNest RecruitPlan - Landing page + form thu lead

- `index.html` - landing page (một file, mở thẳng bằng trình duyệt). Có section form `#dang-ky` trước footer; các nút CTA cuộn xuống form.
- `Code.gs` - Google Apps Script nhận dữ liệu form, ghi vào Google Sheet và tạo tab Dashboard.

## Form thu lead

Ô: Họ tên, SĐT (Zalo), Email công việc, Công ty, Quy mô nhân sự, Quan tâm (demo / dùng thử / báo giá), đồng ý liên hệ.
Dữ liệu gửi về Google Sheet của chủ trang. Nguồn truy cập lấy từ `utm_source`, `utm_medium`, `utm_campaign` trên link, ví dụ:

```
index.html?utm_source=facebook&utm_medium=social&utm_campaign=thang-11
```

## Cài đặt nhận dữ liệu

1. Tạo Google Sheet mới, đặt múi giờ (GMT+07:00) Hà Nội.
2. Extensions > Apps Script, dán toàn bộ `Code.gs`, chạy hàm `setup` một lần.
3. Deploy > New deployment > Web app (Execute as: Me, Who has access: Anyone), copy URL đuôi `/exec`.
4. Trong `index.html`, thay `[DÁN URL WEB APP ĐUÔI /exec VÀO ĐÂY]` bằng URL đó.

Sửa `Code.gs` sau khi deploy: Deploy > Manage deployments > bút chì > New version > Deploy.
