# Bảng điều khiển: Skill tạo PowerPoint đẹp

Kho này chứa 5 Claude Code Skills trong `.claude/skills/` để tạo file thuyết trình `.pptx` thiết kế hiện đại, tỷ lệ 16:9, có chuyển động mượt bằng Morph, GIF động, biểu đồ xuất hiện từng bước và quy trình kiểm tra chất lượng.

## Cách Claude xử lý lệnh tắt

Khi tin nhắn của người dùng có một lệnh tắt dưới đây (ở bất kỳ vị trí nào trong câu), Claude phải gọi tool `Skill` với (các) skill tương ứng **trước khi** làm bất cứ việc gì khác, rồi làm theo hướng dẫn của skill đó. Phần chữ còn lại trong tin nhắn là yêu cầu cụ thể (chủ đề, số liệu, file cần sửa…).

| Lệnh tắt | Skill được gọi | Dùng khi |
|----------|----------------|----------|
| `@pptx` | Trọn bộ quy trình (xem bên dưới) | Tạo một deck `.pptx` hoàn chỉnh, đẹp, chuẩn bố cục 16:9 từ brief |
| `@morph` | `pptx-morph-motion` + `pptx-visual-design` | Tạo các cặp slide gắn tên `!!` để Morph biến hình mượt (bay, zoom/pan, thanh tiến trình, lật thẻ, cột mọc) |
| `@gif-pptx` | `pptx-animated-media` | Tự viết script render hoạt họa GIF (sóng, biểu đồ chạy số, luồng dữ liệu, hạt) và nhúng vào `.pptx` |
| `@chart-pptx` | `pptx-charts-diagrams` | Biểu đồ Bar/Line/Doughnut/Radar, sơ đồ khối, Progressive Reveal (xuất hiện từng bước) |
| `@audit-pptx` | `pptx-quality-audit` | Rà soát và sửa lỗi tràn chữ, lệch khung, lấn lề, tương phản, trùng tên Morph trong file `.pptx` |

Có thể ghép nhiều lệnh tắt, ví dụ `@morph @chart-pptx báo cáo doanh thu quý 3`. Khi đó gọi các skill theo thứ tự của quy trình `@pptx`.

Không có lệnh tắt nhưng người dùng muốn tạo hoặc sửa file `.pptx` trong kho này thì Claude vẫn dùng bộ skill này (ít nhất `pptx-visual-design` + `pptx-quality-audit`), không dùng skill PowerPoint chung khác.

## Quy trình `@pptx`

Gọi skill đúng thứ tự, mỗi bước chỉ nạp skill khi tới bước đó. Bỏ qua bước không cần cho brief (vd. không có số liệu thì bỏ bước 3).

1. **Hiểu brief**: chủ đề, khán giả, mục đích (thuyết phục / báo cáo / giảng dạy), số slide, trình chiếu hay gửi đọc, có cần Morph không. Thiếu thông tin quan trọng thì đề xuất giả định cụ thể rồi làm, không hỏi dồn.
2. **`pptx-visual-design`**: lập design plan trong câu trả lời (4–6 màu hex kèm vai trò, cặp font, danh sách slide → bố cục trong 8 bố cục, một "chữ ký" riêng của deck), đối chiếu bảng lối mòn AI, rồi mới code.
3. **`pptx-charts-diagrams`**: nếu có số liệu hoặc quy trình. Tiêu đề biểu đồ là kết luận; dùng Progressive Reveal cho biểu đồ/sơ đồ phức tạp.
4. **`pptx-morph-motion`**: viết kịch bản chuyển động (mỗi transition một câu), dựng mảng `ELEMENTS` với `states`, đặt tên `!!`, chèn Morph bằng `injectMorph`.
5. **`pptx-animated-media`**: chỉ khi cần chuyển động liên tục trong một slide. Tối đa một GIF làm điểm nhấn mỗi slide.
6. **`pptx-quality-audit`**: chạy `audit_pptx.py` + `render_preview.sh`, mở ảnh từng slide để soát, sửa đến khi 0 ERROR. Bắt buộc trước khi giao.
7. **Bàn giao**: đường dẫn file `.pptx`, tóm tắt kết quả kiểm tra, và nhắc người dùng cách bật/xem Morph (mục README "Bật hiệu ứng Morph").

## Quy ước chung

- Bố cục: `pptx.layout = "LAYOUT_16x9"` (10 × 5.625 inch), lề an toàn 0.5 inch, lưới 12 cột, gutter 0.2 inch.
- Chữ: Title 36–44pt, Heading 20–24pt, Body 14–16pt, không chữ < 12pt, tối đa 6 dòng chữ nội dung mỗi slide.
- Nội dung tiếng Việt dùng sentence case, font có đủ dấu (Georgia, Arial, Cambria, Calibri, Segoe UI, Times New Roman).
- Code sinh deck đặt trong `scripts/` hoặc `decks/<tên>/build.js`; file xuất ra `output/` (đã có trong `.gitignore`). Luôn sinh lại từ code thay vì vá file đã sinh.
- Điểm khởi đầu nhanh: `templates/morph-deck-generator.js` (4 slide Morph + Bento Grid). Sao chép và sửa, không viết lại từ đầu.
- Đặt tên phần tử: `!!<vai-trò>` cho phần tử Morph; chứa `deco`, `bg` hoặc `bleed` cho phần tử cố ý tràn mép (audit miễn kiểm tra lề).
- Mỗi slide có speaker notes (`slide.addNotes`), mỗi ảnh có `altText`.

## Khi các skill mâu thuẫn nhau

Theo thứ tự ưu tiên (cao xuống thấp):

1. Yêu cầu rõ ràng của người dùng trong brief.
2. Đọc được: tương phản, cỡ chữ tối thiểu 12pt, lề an toàn, không chữ đè chữ (`pptx-quality-audit`).
3. Định hướng thẩm mỹ của `pptx-visual-design` (bố cục, màu, chữ, tiết chế).
4. Mức độ chuyển động: `pptx-visual-design` và mục "Tiết chế" của `pptx-morph-motion` thắng khi muốn thêm hiệu ứng. Slide nhiều chữ để yên; mỗi transition một ý chính; mỗi slide tối đa một GIF.
5. Chi tiết kỹ thuật của từng skill chuyên biệt (charts, animated-media).

## Môi trường

- Node.js 18+: `npm install` một lần (cài `pptxgenjs`; `jszip` đi kèm).
- Python 3 + `pip install pillow matplotlib numpy` cho GIF; `pip install python-pptx` cho `fix_pptx.py`.
- LibreOffice (`soffice`) + poppler (`pdftoppm`) cho ảnh xem trước.
- Playwright (tùy chọn) cho HTML Canvas → GIF và SVG → PNG.
- LibreOffice không chạy Morph và chỉ hiện khung đầu của GIF: ảnh xem trước dùng để soát bố cục tĩnh; chuyển động phải xem bằng PowerPoint 365/2019+.
