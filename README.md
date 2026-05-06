# office-skill-vn

Skill cho **Cursor Agent** và công cụ dòng lệnh: **tạo / chỉnh** tài liệu Word (`.docx`) **theo mẫu quy định** — công văn / trả lời công văn **hai bảng** (tham chiếu pháp lý thường gặp: **NĐ 30/2020**; chuẩn thực tế là **file mẫu** hoặc output `rebuild` đúng bố cục) và **văn bản QPPL** theo **phụ lục / mẫu** (tham chiếu **NĐ 78/2025**). Kèm **`scripts/SKILL/`** — Word tổng quát, Excel, PowerPoint / PptxGenJS (bổ sung Claude), không bắt buộc cho CLI hai bảng.

## Yêu cầu

- **Python 3.10+**
- **`python-docx`** (cài qua `requirements.txt`)
- **Microsoft Word** (khuyến nghị) để kiểm tra bản in và bố cục

## Cài đặt

```powershell
git clone https://github.com/thanhnn91qn-afk/office-skill-vn.git
cd office-skill-vn
pip install -r requirements.txt
```

Trên Windows, nếu log Python lỗi encoding tiếng Việt, có thể đặt trước khi chạy:

```powershell
$env:PYTHONIOENCODING = "utf-8"
```

**Excel / PowerPoint:** chỉ cần khi làm theo tài liệu trong `scripts/SKILL/` — cài thêm `openpyxl`, `pandas` (Python) hoặc `pptxgenjs` (npm).

## Gắn skill vào Cursor

- Sao chép thư mục repo vào thư mục skill của Cursor, ví dụ:
  - Windows: `%USERPROFILE%\.cursor\skills\office-skill-vn\`
  - macOS / Linux: `~/.cursor/skills/office-skill-vn/`
- Trong Cursor: **Skills** → chọn **office-skill-vn** (file định nghĩa chính là **`SKILL.md`** ở gốc repo).

## Cấu trúc chính

| Đường dẫn | Nội dung |
|-----------|----------|
| `SKILL.md` | Hướng dẫn đầy đủ: **mẫu quy định** trước, NĐ 30/78 là tham chiếu; `python-docx`; Office mở rộng |
| `reference-van-ban-quy-pham-phap-luat.md` | Tham chiếu văn bản QPPL (mẫu phụ lục; NĐ 78) |
| `scripts/office_skill_cli.py` | CLI: `rebuild`, `fix`, `legacy` |
| `scripts/SKILL/` | Tài liệu chi tiết Word / Excel / PowerPoint — xem [scripts/SKILL/README.md](scripts/SKILL/README.md) |

**Lưu ý:** `Mau_cong_van_ND30_tai_ve.docx` là **mẫu đông lạnh** gắn skill (đối chiếu NĐ 30); đơn vị có thể dùng mẫu riêng. **`rebuild`** **không** đọc file mẫu — **tạo mới** `.docx` **hai bảng đúng bố cục mẫu quy định** từ `--source`.

## Sử dụng CLI nhanh

### 1) `rebuild` — tạo `.docx` hai bảng (mẫu quy định) từ file nguồn

Trích **thân bài** (và metadata nếu có) từ `du_thao.docx`, xuất `ket_qua.docx` có hai bảng layout, gạch ngang quốc hiệu / tiêu ngữ / tên cơ quan theo logic trong script. Nguồn phải có tiếng Việt đầy đủ dấu (script từ chối nếu nghi ngờ mất dấu).

```powershell
python scripts/office_skill_cli.py rebuild --source "du_thao.docx" --output "ket_qua.docx"
```

Tùy chọn: `--no-justify-body`, `--body-pt 13` hoặc `14` (mặc định 14).

### 2) `fix` — kiểm tra hoặc sửa file sẵn có

- **Chỉ kiểm tra** (mặc định): không chỉnh layout, **không ghi đè** file — chỉ báo nếu qua kiểm tra tiếng Việt.

```powershell
python scripts/office_skill_cli.py fix "file.docx"
```

- **Chỉ ép spacing** đoạn về 0 pt (trước/sau), có **lưu** file:

```powershell
python scripts/office_skill_cli.py fix "file.docx" --spacing-only
```

- **Áp layout** (bảng đầu, chữ ký, căn thân bài, gạch ngang header, …) — **chỉ khi cần**, có **lưu** file:

```powershell
python scripts/office_skill_cli.py fix "file.docx" --apply-layout
```

Kèm `--apply-layout` có thể dùng thêm: `--no-justify-body`, `--body-pt 13|14`, `--keep-empty-lines`.

### 3) `legacy` — không dùng cho công văn hai bảng đúng mẫu

Tạo kiểu xếp dòng căn giữa cũ; mặc định **bị từ chối** trừ khi thêm `--allow-legacy-stack`. Xem `SKILL.md` và `--help` của subcommand.

## Ghi chú an toàn

- `fix` không có `--spacing-only` và không có `--apply-layout`: file **không** bị sửa trên đĩa.
- `--spacing-only` và `--apply-layout` đều có cơ chế hạn chế làm tăng ký tự lỗi `?` / ``.
- Đóng file trong Word trước khi CLI ghi cùng đường dẫn để tránh lỗi khóa file.

## Repository

https://github.com/thanhnn91qn-afk/office-skill-vn
