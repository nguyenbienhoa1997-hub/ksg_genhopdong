# VIBE-REQUIREMENT-SPEC.md
> Tool Gen Hợp Đồng Tự Động | VIBE MODE | v1.0

---

## 1. PROBLEM

| | |
|---|---|
| **Vấn đề bề mặt** | Tạo hợp đồng mất thời gian |
| **Vấn đề gốc** | Copy-paste thủ công + fill Excel tay → chậm, dễ sai sót |
| **User** | Team nội bộ ~5 người, cùng mạng LAN |
| **Hành vi hiện tại** | Mở Word, copy dữ liệu từ Excel, paste từng ô thủ công |

---

## 2. TECH STACK

```
✅ DÙNG
- Python 3.10+     → runtime
- Flask            → web framework, đơn giản, đủ dùng
- SQLite           → database, zero-config, đủ cho 5 user
- docxtpl          → fill biến vào Word template (Jinja2 syntax)
- docx2pdf         → convert Word → PDF qua MS Word COM (Windows)
- HTML/CSS/JS thuần → frontend, không build tool

❌ TRÁNH
- React / Vue / Next.js   → over-engineer
- PostgreSQL / MySQL       → quá nặng
- Docker                  → không cần thiết
- Redis / Celery           → overkill cho use case này
```

**Deploy:** 1 máy Windows host → team truy cập `http://192.168.x.x:5000`

---

## 3. UI & BRAND

| | |
|---|---|
| **Style** | Notion-like — clean, trắng, typography rõ, minimal |
| **Layout** | Sidebar trái cố định + content area phải |
| **Font** | System font (Inter / Segoe UI) |
| **Màu chính** | Trắng nền, xám nhạt border, đen text, xanh dương cho CTA |
| **Spacing** | Rộng rãi, nhiều whitespace |
| **Components** | Table rows có hover, button outline, tag/badge cho trạng thái |

---

## 4. FEATURES

### ⭐ MUST (MVP Core)

| ID | Tính năng | Mô tả |
|----|-----------|-------|
| F1 | Quản lý mẫu Word | Upload / xem / xóa mẫu hợp đồng `.docx` |
| F2 | Upload Excel | Nhận file `.xlsx`, parse header + data |
| F3 | Preview dữ liệu | Hiện bảng 5 dòng đầu + chọn cột mapping |
| F4 | Generate PDF | Mỗi dòng Excel → fill mẫu Word → xuất PDF |
| F5 | Danh sách hợp đồng | List tất cả hợp đồng đã tạo, sort theo ngày |
| F6 | Chi tiết hợp đồng | Màn hình detail: tab Thông tin + tab Tải PDF |

### SHOULD (nên có, làm sau MUST)

| ID | Tính năng | Mô tả |
|----|-----------|-------|
| S1 | Tìm kiếm | Search theo tên hợp đồng trong danh sách |
| S2 | Lọc theo mẫu | Filter danh sách theo template đã dùng |
| S3 | Download ZIP | Tải toàn bộ PDF của 1 lô về cùng lúc |

### LATER (không làm MVP)

- Xem trước PDF inline (iframe)
- Phân trang danh sách

---

## 5. NON-GOAL

```
❌ Login / xác thực / phân quyền user
❌ Chỉnh sửa nội dung hợp đồng trực tiếp trên web
❌ Gửi email / ký điện tử
❌ Multi-tenant / nhiều công ty
❌ Sync cloud / backup tự động
❌ Mobile responsive (chỉ cần desktop)
```

---

## 6. USER FLOW

```
[SETUP - 1 lần]
Vào tab "Mẫu hợp đồng"
→ Upload file .docx (đã có biến {{ TEN_BIEN }} bên trong)
→ Đặt tên mẫu
→ Lưu vào hệ thống

[GENERATE - mỗi lần cần tạo hợp đồng]
Vào tab "Tạo hợp đồng"
→ Upload file Excel
→ Hệ thống parse → hiện preview 5 dòng đầu
→ User chọn: cột nào là "tên hợp đồng", cột nào là "chọn mẫu"
→ Bấm "Tạo hợp đồng"
→ Hệ thống generate PDF từng dòng
→ Redirect sang Danh sách hợp đồng (hiện lô vừa tạo)

[VIEW & DOWNLOAD]
Danh sách hợp đồng
→ Click vào 1 hợp đồng
→ Màn hình chi tiết
  → Tab "Thông tin": hiện tất cả field từ Excel
  → Tab "Hợp đồng": nút Tải PDF về
```

---

## 7. UI SƠ BỘ

### Layout chung
```
┌─────────────────────────────────────────────┐
│  SIDEBAR (220px)     │  CONTENT AREA         │
│                      │                       │
│  📋 Hợp đồng         │  [Page content here]  │
│  ➕ Tạo mới          │                       │
│  📄 Quản lý mẫu      │                       │
│                      │                       │
└─────────────────────────────────────────────┘
```

### Screen 1 — Danh sách hợp đồng (`/`)
```
Danh sách hợp đồng                    [+ Tạo mới]

🔍 Tìm kiếm...          [Lọc theo mẫu ▼]

┌──────────────────────────────────────────────┐
│ Tên hợp đồng  │ Mẫu dùng │ Ngày tạo │       │
├──────────────────────────────────────────────┤
│ HĐ-Nguyen Van A│ Mẫu 1   │ 24/06/26 │  →    │
│ HĐ-Tran Thi B  │ Mẫu 2   │ 24/06/26 │  →    │
└──────────────────────────────────────────────┘
```
- Click bất kỳ dòng → vào Chi tiết

### Screen 2 — Tạo hợp đồng (`/generate`)
```
Tạo hợp đồng mới

[ Kéo thả hoặc chọn file Excel (.xlsx) ]

── Sau khi upload ──────────────────────────────
Preview dữ liệu (5 dòng đầu):
┌─────────────────────────────────────┐
│ Tên KH      │ Địa chỉ │ Giá trị    │
├─────────────────────────────────────┤
│ Nguyễn Văn A│ HN      │ 10,000,000 │
└─────────────────────────────────────┘

Cột tên hợp đồng:  [Chọn cột ▼]
Cột chọn mẫu:      [Chọn cột ▼]  ← Giá trị: tên mẫu đã setup

                              [Tạo hợp đồng →]
```

### Screen 3 — Chi tiết hợp đồng (`/contracts/:id`)
```
← Danh sách        HĐ-Nguyễn Văn A

[Thông tin]  [Hợp đồng]

── Tab Thông tin ────────────────────
Tên khách hàng:   Nguyễn Văn A
Địa chỉ:          Hà Nội
Giá trị HĐ:       10,000,000
Mẫu sử dụng:      Mẫu 1
Ngày tạo:         24/06/2026

── Tab Hợp đồng ─────────────────────
                  [⬇ Tải PDF về]
```

### Screen 4 — Quản lý mẫu (`/templates`)
```
Mẫu hợp đồng                [+ Upload mẫu mới]

┌────────────────────────────────────────┐
│ Tên mẫu  │ File          │ Ngày upload │     │
├────────────────────────────────────────┤
│ Mẫu 1    │ mau_hdmb.docx │ 20/06/26   │ 🗑  │
│ Mẫu 2    │ mau_hddv.docx │ 20/06/26   │ 🗑  │
└────────────────────────────────────────┘

ℹ️ Dùng {{ TEN_BIEN }} trong file Word để điền dữ liệu từ Excel
```

---

## 8. BEHAVIOR

### Generate flow
- Nếu cột "chọn mẫu" = tên mẫu đã setup → dùng mẫu đó
- Nếu không khớp tên mẫu nào → dùng mẫu đầu tiên + hiện warning
- Nếu dòng bị lỗi (biến không tìm thấy) → skip dòng đó + ghi log lỗi
- Sau generate xong → redirect về `/` với highlight lô vừa tạo

### Validation
- Excel phải có ít nhất 1 dòng dữ liệu (ngoài header)
- File upload chỉ nhận `.xlsx` và `.docx`
- Tên mẫu không được trùng nhau

### Error states
- Upload file sai định dạng → toast error đỏ
- Generate thất bại hoàn toàn → thông báo lỗi rõ ràng
- Generate có 1 số dòng lỗi → hiện "X/Y hợp đồng tạo thành công, N lỗi"

### Database
```
templates  (id, name, filename, file_path, created_at)
contracts  (id, name, template_id, template_name, pdf_path,
            row_data_json, batch_id, created_at)
batches    (id, excel_filename, contract_count, created_at)
```

---

## 9. ACCEPTANCE CHECKLIST

- [ ] Upload mẫu Word → lưu vào DB và disk
- [ ] Upload Excel 20 dòng → parse đúng header + data
- [ ] Preview hiện đúng 5 dòng đầu
- [ ] Generate 20 dòng → 20 PDF đúng dữ liệu, không sai field
- [ ] Hợp đồng vừa tạo xuất hiện trong danh sách
- [ ] Click vào hợp đồng → thấy đúng thông tin + tải được PDF
- [ ] 5 người vào cùng lúc không crash / không lỗi

---

## 10. SUCCESS STATE

```
✅ Upload Excel 20 dòng → generate 20 PDF đúng dữ liệu, không sai sót
✅ Cả team 5 người truy cập cùng lúc, không lỗi
✅ Tạo xong 1 lô hợp đồng trong < 5 phút (so với hàng tiếng copy-paste)
```

---

## 11. DECISIONS

- **Chọn mẫu:** Excel dùng **mã mẫu** (ví dụ: HD001) → phải khớp chính xác mã đã setup
- **Watermark + số trang:** Setup sẵn trong file Word template → hệ thống không xử lý
- **Xóa hợp đồng:** User xóa tay được (có nút xóa trong danh sách / chi tiết)

---

## 12. STATUS

```
✅ VIBE-1 Discovery   — Hoàn thành
✅ VIBE-2 Requirement — Hoàn thành
⏳ VIBE-3 Code        — Bước tiếp theo
```

Gõ /code hoặc "bắt đầu code" để tiếp.
