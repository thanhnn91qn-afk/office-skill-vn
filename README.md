# office-skill-vn

Skill cho **Cursor Agent** để **tạo và chỉnh sửa** văn bản Word (`.docx`) theo thể thức công văn.

## Yêu cầu

- **Python 3.10+**
- **Microsoft Word** để kiểm tra bản in

## Cài đặt

```powershell
git clone https://github.com/thanhnn91qn-afk/office-skill-vn.git
cd office-skill-vn
pip install -r requirements.txt
```

**Excel / PowerPoint (tài liệu trong `scripts/SKILL/`):** cài thêm `openpyxl`, `pandas`, hoặc `pptxgenjs` (npm) khi cần.

## Gắn skill vào Cursor

- Chép thư mục này vào `~/.cursor/skills/office-skill-vn/`
- Trong Cursor: `+` -> `Skills` -> chọn `office-skill-vn`

## Cấu trúc chính

- `SKILL.md`: Word NĐ 30/78 + tóm tắt Office mở rộng
- `scripts/office_skill_cli.py`: CLI công văn NĐ 30 (`rebuild`, `fix`, `legacy`)
- `scripts/SKILL/`: tài liệu tham chiếu Word tổng quát, Excel, PowerPoint, [pptxgenjs.md](scripts/SKILL/pptxgenjs.md) — xem [scripts/SKILL/README.md](scripts/SKILL/README.md)

## Sử dụng nhanh

### 1) Rebuild văn bản từ nguồn `.docx`

```powershell
python scripts/office_skill_cli.py rebuild --source "du_thao.docx" --output "ket_qua.docx"
```

### 2) Chỉ ép spacing về 0pt (không đổi nội dung/layout)

```powershell
python scripts/office_skill_cli.py fix "file_can_sua.docx" --spacing-only
```

### 3) Áp layout (chỉ dùng khi cần)

```powershell
python scripts/office_skill_cli.py fix "file_can_sua.docx" --apply-layout
```

## Ghi chú

- Mặc định `fix` không tự áp layout nếu không có `--apply-layout`.
- Khi chỉ cần sửa khoảng cách đoạn, dùng `--spacing-only` để an toàn nội dung.

## Repository

https://github.com/thanhnn91qn-afk/office-skill-vn
