# PROGRESS — Tool Gen Hợp Đồng Tự Động

## FEATURES STATUS

| ID | Feature | Status |
|----|---------|--------|
| F1 | Quản lý mẫu Word | ✅ DONE |
| F2 | Upload Excel | ✅ DONE |
| F3 | Preview dữ liệu | ✅ DONE |
| F4 | Generate PDF | ✅ DONE |
| F5 | Danh sách hợp đồng | ✅ DONE |
| F6 | Chi tiết hợp đồng | ✅ DONE |
| S1 | Tìm kiếm | ⏳ TODO |
| S2 | Lọc theo mẫu | ⏳ TODO |

## ✅ DONE
- Scaffold: project structure, DB schema, base layout, CSS, routing

## 🔄 IN PROGRESS
- (none)

## ⏳ TODO
- Module 1: F1 — Quản lý mẫu
- Module 2: F2+F3 — Upload Excel + Preview
- Module 3: F4 — Generate PDF
- Module 4: F5+F6 — Danh sách + Chi tiết

## 🔧 TECHNICAL DECISIONS
- SQLite: zero-config, đủ cho 5 user đồng thời
- docxtpl: Jinja2 syntax `{{ TEN_BIEN }}` trong Word
- docx2pdf: dùng MS Word COM trên Windows
- host=0.0.0.0: team truy cập qua LAN
- HTML/CSS/JS thuần: không framework

## 🐛 BUGS FIXED
(none yet)

## ✅ AC STATUS
- [ ] Upload mẫu Word → lưu DB + disk
- [ ] Upload Excel 20 dòng → parse đúng
- [ ] Preview 5 dòng đầu đúng
- [ ] Generate 20 dòng → 20 PDF đúng dữ liệu
- [ ] Hợp đồng vừa tạo hiện trong danh sách
- [ ] Click hợp đồng → thông tin + tải PDF được
- [ ] 5 người cùng lúc không crash

## 🌟 SUCCESS STATE STATUS
- [ ] 20 PDF đúng, không sai sót
- [ ] 5 người đồng thời, không lỗi
- [ ] Xong 1 lô < 5 phút
