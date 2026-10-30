---
name: pptx-quality-audit
description: Quy trình kiểm tra chất lượng file .pptx trước khi bàn giao - phát hiện chữ tràn khung (text overflow), khung chữ đè nhau (overlap), phần tử lấn lề an toàn (margin >= 0.5 inch) hoặc vượt mép slide, độ tương phản màu theo WCAG, cỡ chữ quá nhỏ, quá 6 dòng chữ, tên Morph "!!" trùng lặp, GIF quá nặng, ảnh méo; xuất ảnh từng slide bằng LibreOffice (soffice --headless --convert-to pdf) + pdftoppm để soát bằng mắt; tự sửa lỗi lề và tràn chữ. Dùng khi người dùng nhắc @audit-pptx, kiểm tra pptx, rà soát slide, sửa lỗi tràn chữ, lệch khung, QA PowerPoint, trước khi giao file .pptx.
---

# PPTX Quality Audit

Không file `.pptx` nào được giao trước khi qua quy trình này. Chạy được cho deck do `pptxgenjs` tạo lẫn file `.pptx` người dùng gửi tới.

Ba lớp kiểm tra, theo thứ tự:

1. **Tự động** (`scripts/audit_pptx.py`): đọc XML, tính hình học, ước tính tràn chữ, trộn màu tính tương phản. Nhanh, bắt được lỗi định lượng.
2. **Trực quan** (`scripts/render_preview.sh`): LibreOffice → PDF → PNG từng slide. Claude mở từng ảnh để soi những gì script không thấy.
3. **Sửa và lặp**: sửa code sinh deck (ưu tiên) hoặc dùng `scripts/fix_pptx.py` cho file có sẵn, rồi chạy lại 1 và 2 đến khi đạt.

## Bước 1: Kiểm tra tự động

```bash
S=.claude/skills/pptx-quality-audit/scripts
python3 $S/audit_pptx.py output/deck.pptx                 # báo cáo chữ
python3 $S/audit_pptx.py output/deck.pptx --morph-map     # + phần tử "!!" ghép cặp / hiện ra / biến mất
python3 $S/audit_pptx.py output/deck.pptx --json          # cho xử lý tiếp
python3 $S/audit_pptx.py output/deck.pptx --margin 0.5 --max-lines 6 --min-font 12 --quiet-info
```

Không cần thư viện ngoài. Mã thoát `1` nếu có `ERROR`, dùng được trong CI.

Mẫu kết quả:

```
output/deck.pptx: 8 slide, 10.0x5.625 in, Morph ở slide [2, 3, 4]
ERROR BOUNDS   slide 3 "box": Vượt mép slide (x=8.50, y=2.00, w=2.00, h=1.00)
ERROR CONTRAST slide 5 "!!hero-kpi": Tương phản 2.44:1 < 3.0:1 cho chữ 100pt "+38%"
WARN  OVERFLOW slide 2 "Text 1": Ước tính cần 1.80" chiều cao, khung chỉ có 1.50"
WARN  MARGIN   slide 6 "Text 2": Lấn lề an toàn 0.5": trái 0.10", dưới 0.22"
WARN  OVERLAP  slide 7 "Text 3": Đè lên khung chữ "Text 4" (0.85 in²)
INFO  MORPH    slide 2: 9 phần tử '!!' dùng chung với slide 1 nhưng slide 2 chưa bật Morph
```

### Các nhóm lỗi và cách sửa

| Mã | Phát hiện | Mức | Sửa trong code pptxgenjs |
|----|-----------|-----|--------------------------|
| `SIZE` | Slide không phải 16:9 | WARN | `pptx.layout = "LAYOUT_16x9"` |
| `BOUNDS` | Phần tử vượt mép slide | ERROR | Dời vào trong; nếu cố ý tràn mép, đặt tên có `bleed`/`bg`/`deco` |
| `MARGIN` | Phần tử trong slide nhưng cách mép < 0.5 inch | WARN | `x ≥ 0.5`, `y ≥ 0.5`, `x + w ≤ 9.5`, `y + h ≤ 5.125` |
| `OVERFLOW` | Ước tính chữ cần cao hơn khung, hoặc một từ dài hơn bề rộng khung | WARN / ERROR | Rút gọn chữ → tăng `w` → tăng `h` → giảm `fontSize` (không dưới thang chữ) → tách slide |
| `OVERLAP` | Hai khung chữ giao nhau > 2% diện tích | WARN | Tách tọa độ; nếu cố ý (chữ trong thẻ), đặt chữ vào shape bằng `addText(..., { shape })` |
| `LINES` | > 6 dòng chữ nội dung (đoạn ≥ 25 ký tự, < 20pt) | WARN | Tách slide, chuyển sang sơ đồ/bảng, đưa chi tiết vào speaker notes |
| `FONT` | Chữ < 12pt | WARN | Nâng lên ≥ 12pt; nhãn chart cũng ≥ 12pt |
| `CONTRAST` | Tỷ lệ tương phản < 4.5:1 (chữ thường) hoặc < 3:1 (≥ 18pt hoặc ≥ 14pt đậm) | WARN / ERROR (< 3:1) | Đổi màu chữ, tăng độ đặc của thẻ, dời trang trí khỏi vùng chữ |
| `MORPH` | Tên `!!` trùng trong một slide; có cặp `!!` nhưng chưa bật Morph | ERROR / INFO | Tên duy nhất; chạy `injectMorph` (xem `pptx-morph-motion`) |
| `MEDIA` | GIF > 8 MB; ảnh bị méo tỷ lệ | WARN | Render GIF nhỏ hơn; `h = w / (rộng px / cao px)` hoặc `sizing: { type: "cover" }` |

### Script tính toán thế nào

- **Lề**: đọc `a:off` / `a:ext` (EMU, 914400 EMU = 1 inch) của mọi shape, ảnh, chart, kể cả trong group. Phần tử `!!` nằm hẳn ngoài slide được bỏ qua (trạng thái bay vào/ra của Morph). Phần tử bị bóp rộng < 0.05 inch (thẻ lật) cũng bỏ qua.
- **Tràn chữ**: bề rộng khả dụng = `w − lIns − rIns`. Mỗi ký tự được gán bề rộng theo em (chữ thường 0.55, hoa/số 0.66, hẹp 0.3, rộng 0.85, khoảng trắng 0.28), ×1.06 nếu đậm. Ngắt dòng theo từ, cộng dòng cứng `<a:br/>`. Chiều cao cần = Σ dòng × cỡ chữ × 1.2 / 72. Báo khi cần > 108% chiều cao khả dụng. Đây là **ước tính thận trọng**, luôn xác nhận bằng ảnh ở bước 2.
- **Tương phản**: lấy màu nền slide, rồi trộn lần lượt (alpha blending) mọi shape nằm dưới tâm khung chữ theo thứ tự lớp, có xét `transparency` và hình elip. Màu chữ cũng được trộn theo `transparency` của chữ. Tính theo công thức độ sáng tương đối WCAG 2.x. Chữ nằm trên ảnh được báo `INFO` để soát bằng mắt.

## Bước 2: Xuất ảnh và soát bằng mắt

```bash
bash .claude/skills/pptx-quality-audit/scripts/render_preview.sh output/deck.pptx output/preview 110
# -> output/preview/deck.pdf, slide-01.png ... slide-NN.png, contact-sheet.png
```

Lệnh tương đương nếu cần chạy tay:

```bash
soffice --headless --convert-to pdf --outdir output/preview output/deck.pptx
pdftoppm -png -r 110 output/preview/deck.pdf output/preview/slide
```

Cài đặt: `sudo apt install libreoffice-impress poppler-utils` (Ubuntu/Debian), `brew install --cask libreoffice && brew install poppler` (macOS). Script dùng hồ sơ LibreOffice tạm (`-env:UserInstallation`) nên chạy song song không bị khóa.

Sau khi xuất, **mở `contact-sheet.png` để nhìn toàn deck, rồi mở từng `slide-XX.png`** (Claude dùng tool Read với file ảnh) và soát:

- [ ] Chữ bị cắt, tràn ra ngoài thẻ, đè lên chữ khác
- [ ] Từ mồ côi: dòng cuối tiêu đề chỉ có 1 chữ (sửa bằng `\n` chủ động hoặc đổi `w`)
- [ ] Các phần tử cùng hàng có thẳng mép trái/trên không
- [ ] Khoảng cách giữa các thẻ có đều không (cùng `GAP`)
- [ ] Phân cấp: nhìn 3 giây có biết đâu là điều quan trọng nhất không
- [ ] Màu nhấn chỉ ở một chỗ mỗi slide
- [ ] Hai slide liền nhau có bị trùng bố cục không cần thiết
- [ ] Ảnh có bị méo, mờ, vỡ hạt không
- [ ] Trang trí (quả cầu, khối màu) có đè lên vùng chữ không

### Giới hạn của bản xem trước

- **Font**: LibreOffice thay font thiếu bằng font khác, bề rộng chữ có thể lệch. Font tương thích số đo có sẵn trên Linux: Calibri → Carlito, Cambria → Caladea, Arial → Liberation Sans, Times New Roman → Liberation Serif. Georgia không có bản tương thích mặc định (cài `Gelasio` để gần đúng). Kiểm tra bằng `fc-list | grep -i <tên font>`. Nếu font thay thế **rộng hơn** bản gốc thì ảnh xem trước là trường hợp xấu nhất, an toàn để soát tràn chữ.
- **Morph**: không chạy. Mỗi slide hiện ở trạng thái tĩnh. Dùng `--morph-map` để soát cặp ghép.
- **GIF**: chỉ hiện khung đầu.
- **Chart**: LibreOffice vẽ lại chart, kiểu chữ/khoảng cách có thể khác PowerPoint một chút.

## Bước 3: Sửa

### Ưu tiên: sửa trong code sinh deck

Sửa tại nguồn (token, hàm bố cục, mảng `states`) rồi sinh lại file. Đừng vá file `.pptx` đã sinh nếu còn code.

Chèn bộ chặn vào code để lỗi không xảy ra ngay từ lúc sinh:

```js
const SAFE = { x0: 0.5, y0: 0.5, x1: 9.5, y1: 5.125 };

// Gọi trước mỗi addText/addShape/addImage nội dung (bỏ qua phần tử deco/bleed/bg)
function assertSafe(name, { x, y, w, h }) {
  const bad = [];
  if (x < SAFE.x0 - 1e-3) bad.push(`trái ${x.toFixed(2)}`);
  if (y < SAFE.y0 - 1e-3) bad.push(`trên ${y.toFixed(2)}`);
  if (x + w > SAFE.x1 + 1e-3) bad.push(`phải ${(x + w).toFixed(2)}`);
  if (y + h > SAFE.y1 + 1e-3) bad.push(`dưới ${(y + h).toFixed(2)}`);
  if (bad.length && !/deco|bleed|(^|[-_!])bg([-_]|$)/i.test(name)) {
    throw new Error(`"${name}" lấn lề an toàn: ${bad.join(", ")}`);
  }
}

// Ước tính số dòng (cùng hệ số với audit_pptx.py) để chọn cỡ chữ vừa khung
function charEm(ch) {
  if (/\s/.test(ch)) return 0.28;
  if (/[A-ZÀ-Ỹ0-9]/.test(ch)) return 0.66;
  if (/[il.,:;'|!"“”‘’()\[\]]/.test(ch)) return 0.3;
  if (/[mwMW%@]/.test(ch)) return 0.85;
  return 0.55;
}
function textWidthIn(str, pt, bold = false) {
  return [...str].reduce((s, ch) => s + charEm(ch), 0) * (pt / 72) * (bold ? 1.06 : 1);
}
function estimateLines(str, pt, widthIn, bold = false) {
  return str.split("\n").reduce((total, para) => {
    let lines = 1, cur = 0;
    for (const word of para.split(/\s+/).filter(Boolean)) {
      const ww = textWidthIn(word, pt, bold), sp = textWidthIn(" ", pt, bold);
      if (cur && cur + sp + ww > widthIn) { lines++; cur = ww; } else cur += (cur ? sp : 0) + ww;
    }
    return total + lines;
  }, 0);
}
// Cỡ chữ lớn nhất trong [minPt, maxPt] để chữ vừa khung (margin: 0)
function fitFontSize(str, { w, h }, { maxPt = 40, minPt = 28, bold = false, lineHeight = 1.2 } = {}) {
  for (let pt = maxPt; pt >= minPt; pt--) {
    if (estimateLines(str, pt, w, bold) * pt * lineHeight / 72 <= h) return pt;
  }
  return null; // không vừa ở cỡ tối thiểu: rút gọn chữ hoặc tách slide
}

// Dùng:
const box = { x: 0.5, y: 0.5, w: 9, h: 0.65 };
assertSafe("title", box);
const pt = fitFontSize("Tháng 9 gánh cả quý", box, { maxPt: 40, minPt: 36 });
if (!pt) throw new Error("Tiêu đề quá dài cho một dòng 36pt: viết ngắn lại");
slide.addText("Tháng 9 gánh cả quý", { ...box, fontSize: pt, margin: 0 });
```

Tính tương phản ngay trong code (WCAG 2.x):

```js
function luminance(hex) {
  const [r, g, b] = [0, 2, 4].map((i) => parseInt(hex.slice(i, i + 2), 16) / 255)
    .map((c) => (c <= 0.03928 ? c / 12.92 : ((c + 0.055) / 1.055) ** 2.4));
  return 0.2126 * r + 0.7152 * g + 0.0722 * b;
}
function contrast(fg, bg) {
  const [a, b] = [luminance(fg), luminance(bg)].sort((x, y) => y - x);
  return (a + 0.05) / (b + 0.05);
}
// Màu thẻ kính sau khi trộn: fill FFFFFF transparency 85 trên nền 16130F
function blend(top, bottom, alpha) {
  return [0, 2, 4].map((i) => {
    const t = parseInt(top.slice(i, i + 2), 16), b = parseInt(bottom.slice(i, i + 2), 16);
    return Math.round(t * alpha + b * (1 - alpha)).toString(16).padStart(2, "0");
  }).join("").toUpperCase();
}
const glassOnDark = blend("FFFFFF", "16130F", 0.15);
console.log(contrast("A89F90", glassOnDark).toFixed(2)); // cần ≥ 4.5 cho chữ 12–17pt thường
```

### File có sẵn (không có code nguồn): `fix_pptx.py`

```bash
pip install python-pptx
python3 $S/fix_pptx.py input.pptx output/input-fixed.pptx --margin 0.5 --min-font 12 --max-shrink 0.25
python3 $S/audit_pptx.py output/input-fixed.pptx
```

Script tự:
1. Dời phần tử lấn lề vào trong vùng an toàn (thu nhỏ nếu to hơn vùng nội dung).
2. Nâng chữ < 12pt lên 12pt.
3. Giảm cỡ chữ tràn khung từng 1pt, tối đa 25% để tiêu đề vẫn là tiêu đề; còn tràn thì nới chiều cao khung trong phạm vi lề, và báo `VẪN TRÀN` nếu không đủ chỗ.

Không đụng tới phần tử `deco`/`bg`/`bleed`, phần tử `!!` ngoài slide, phần tử bị xoay. Không bao giờ ghi đè file gốc. Lỗi tương phản, trùng tên Morph và chồng lấn cần người (hoặc Claude) quyết định, nên script chỉ báo, không tự sửa.

## Tiêu chí đạt

- `audit_pptx.py`: **0 ERROR**. Mọi WARN còn lại phải có lý do ghi trong báo cáo (vd. "OVERLAP cố ý: số liệu nằm trong thẻ").
- Ảnh xem trước: không có lỗi nào trong checklist bước 2.
- Deck có Morph: `--morph-map` khớp kịch bản chuyển động (đúng phần tử ghép cặp, không có phần tử biến mất ngoài ý muốn).
- File mở được bằng PowerPoint; nếu có Morph, ghi chú cho người dùng mở bằng PowerPoint 365/2019+.

## Mẫu báo cáo bàn giao

```
Kiểm tra output/deck.pptx (8 slide, 16:9)
- Tự động: 0 ERROR, 1 WARN (OVERLAP slide 4 cố ý: nhãn nằm trong thẻ)
- Trực quan: đã soát 8 ảnh, sửa từ mồ côi ở tiêu đề slide 6
- Đã sửa: slide 2 tiêu đề tràn khung (tăng h 1.5 → 1.9), slide 5 tương phản 2.4 → 6.8 (đổi màu chữ)
- Morph: slide 2–4, 8/8/19 phần tử ghép cặp
- Lưu ý: mở bằng PowerPoint 365/2019+ để thấy Morph; font Georgia/Arial
```
