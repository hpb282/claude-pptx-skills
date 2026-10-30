---
name: pptx-visual-design
description: Quy chuẩn mỹ thuật cho slide PowerPoint 16:9 (LAYOUT_16x9, 10 x 5.625 inch) tạo bằng pptxgenjs, chống lối mòn giao diện AI. Gồm 8 bố cục hiện đại có code mẫu (Bento Grid, Asymmetric Split, Big Number Callout, Process Timeline, Spotlight Comparison, Full-bleed Hero, Executive Dashboard, Quote Slide), lưới 12 cột, lề an toàn 0.5 inch, thang chữ Title 36-44pt / Heading 20-24pt / Body 14-16pt, tối đa 6 dòng chữ mỗi slide, glassmorphism giả lập bằng shape bán trong suốt. Dùng khi tạo deck .pptx, thiết kế slide, chọn màu/font cho bài thuyết trình, @pptx, @morph, slide đẹp, slide hiện đại, presentation design.
---

# PPTX Visual Design

Mục tiêu: mỗi slide trông như do một designer làm cho đúng nội dung đó, không phải template. Áp dụng cho mọi deck tạo bằng `pptxgenjs`. Đi kèm `pptx-morph-motion` khi cần chuyển động, `pptx-charts-diagrams` khi có số liệu, và luôn kết thúc bằng `pptx-quality-audit`.

## 1. Khung slide

```js
const pptx = new pptxgen();
pptx.layout = "LAYOUT_16x9"; // 10 x 5.625 inch. Không dùng LAYOUT_WIDE (13.33 x 7.5) trừ khi người dùng yêu cầu
```

| Thông số | Giá trị |
|----------|---------|
| Kích thước | 10 × 5.625 inch |
| Lề an toàn | 0.5 inch mọi phía → vùng nội dung 9 × 4.625 inch (x 0.5–9.5, y 0.5–5.125) |
| Lưới | 12 cột, khoảng cách (gutter) 0.2 inch, mỗi cột ≈ 0.567 inch |
| Vùng tiêu đề | y 0.5–1.15 |
| Vùng nội dung | y 1.35–5.125 |
| Tràn mép | Chỉ cho ảnh nền, khối màu, trang trí; đặt tên có `bleed`, `bg` hoặc `deco` |

Mọi `x, y, w, h` phải tính từ token, không gõ số rời rạc. Hai phần tử thẳng hàng phải có cùng `x` (hoặc cùng `x + w`) đến 2 chữ số thập phân.

## 2. Thang chữ

| Vai trò | Cỡ | Font | Ghi chú |
|---------|----|------|---------|
| Title (tiêu đề slide) | 36–44pt | display/serif | 1–2 dòng. Ngắt dòng chủ động bằng `\n` để không mồ côi một chữ |
| Heading (tiêu đề thẻ, ý chính) | 20–24pt | display hoặc body đậm | |
| Body | 14–16pt | body | `lineSpacingMultiple: 1.1–1.2` |
| Caption, nhãn, nguồn | 12pt | body | Không nhỏ hơn 12pt: khó đọc khi chiếu |
| Số lớn (Big Number) | 60–160pt | display | Ngoại lệ có chủ đích, mỗi slide tối đa 1 số |
| Eyebrow (nhãn nhỏ trên tiêu đề) | 12pt, IN HOA, đậm | body | `charSpacing: 2` |

Quy tắc:

- **Tối đa 6 dòng chữ nội dung mỗi slide** (câu, bullet). Tiêu đề, nhãn số liệu ngắn và nguồn không tính. Quá 6 dòng → tách slide hoặc chuyển thành sơ đồ/bảng.
- Tối đa 3 cỡ chữ trên một slide (không tính số lớn).
- Căn trái là mặc định. Căn giữa chỉ cho Quote, Big Number đứng một mình, hoặc nhãn dưới cột/bước.
- Dùng sentence case ("Doanh thu quý 3"), không Title Case Every Word. Nội dung tiếng Việt luôn sentence case.
- Đặt `margin: 0` cho text box để chữ thẳng mép với shape cùng `x`.
- Bullet: `bullet: { indent: 14 }` thay vì thụt lề mặc định quá rộng.

### Font an toàn cho tiếng Việt

pptxgenjs không nhúng font. Người nhận phải có font trên máy, nên chọn font hệ thống có đủ dấu tiếng Việt:

| Cặp (Heading / Body) | Có trên | Tính cách |
|----------------------|---------|-----------|
| Georgia / Arial | Windows, macOS | Biên tập, ấm |
| Cambria / Calibri | Máy có Office | Trang trọng, gọn |
| Times New Roman / Arial | Mọi nơi | Học thuật |
| Arial Black / Arial | Windows, macOS | Mạnh, quảng cáo |
| Segoe UI Semibold / Segoe UI | Windows | Hiện đại, công nghệ |

Nếu brief bắt buộc font thương hiệu (Inter, Be Vietnam Pro…), ghi rõ cho người dùng phải cài font đó, và tạo thêm bản PDF để gửi đi.

## 3. Màu

- Bảng màu **4–6 màu hex**: nền, bề mặt (thẻ), chữ chính, chữ phụ, **một** màu nhấn, tùy chọn một màu nhấn phụ cực nhạt.
- Màu nhấn chỉ dùng cho **một thứ quan trọng nhất** trên mỗi slide: con số chính, cột cao nhất, bước hiện tại, phương án đề xuất.
- Chữ phụ phải đạt tương phản ≥ 4.5:1 với nền (≥ 3:1 nếu ≥ 18pt hoặc ≥ 14pt đậm). `pptx-quality-audit` tính tự động.
- Lấy màu từ chủ đề: xưởng cà phê → nâu espresso, kem, đỏ son; bệnh viện nhi → xanh lá nhạt, vàng nắng, không phải xanh dương corporate; fintech → mực đậm, xanh mint. Đặt ra một câu lý do cho mỗi màu.

```js
const T = {
  c: {
    bg: "F4F1EA", surface: "FFFFFF", ink: "1B1A17", muted: "5F5A51",
    line: "D8D2C4", accent: "1F5C4A", accentSoft: "DCE8E2",
  },
};
```

## 4. Chống lối mòn "slide AI"

Trước khi viết code, kiểm tra plan với danh sách này. Gặp mục nào thì đổi.

| Lối mòn | Thay bằng |
|---------|-----------|
| Gradient tím–xanh, nền xanh navy + chữ trắng cho mọi chủ đề | Màu lấy từ chất liệu của chủ đề |
| Mọi slide là tiêu đề + 3–5 bullet | Xoay vòng 8 bố cục bên dưới; không 2 slide liền nhau cùng bố cục (trừ cặp Morph) |
| 3 thẻ bằng nhau, mỗi thẻ một icon + tiêu đề + 2 dòng | Bento lệch kích thước: ô quan trọng nhất to gấp 2–4 lần |
| Mọi thứ căn giữa | Căn trái theo lưới, bất đối xứng có chủ đích |
| Emoji làm bullet/icon | Số thứ tự `01 02 03`, chấm màu, hoặc không có gì |
| Shadow đậm, viền dày, bo góc lớn mọi nơi | Viền 0.75pt màu `line`, bo 0.12–0.16 inch, shadow chỉ cho phần tử nổi bật nhất |
| "Giới thiệu", "Tổng quan", "Kết luận" làm tiêu đề | Tiêu đề là một câu khẳng định: "Tháng 9 gánh cả quý" |
| Số không có đơn vị/so sánh | `+38% so với quý 2`, `2,1 ngày (năm ngoái 4,3)` |
| Ảnh stock bắt tay, bóng đèn, bánh răng | Ảnh thật của sản phẩm/người/địa điểm, hoặc không ảnh |
| Chữ đặt lên ảnh không có lớp phủ | Lớp phủ `transparency: 40–55` hoặc khối màu đặc |
| Dùng nhiều màu nhấn để "sinh động" | Một màu nhấn, phần còn lại trung tính |

## 5. Glassmorphism giả lập

PowerPoint không có backdrop-blur. Hiệu ứng kính được giả lập bằng shape bán trong suốt **đặt trên một nền có màu sắc** (quả cầu màu, ảnh, gradient). Trên nền phẳng một màu, kính chỉ là ô xám.

```js
// Nền tối có quả cầu màu phía sau
slide.background = { color: "16130F" };
slide.addShape(pptx.ShapeType.ellipse, { x: 6.4, y: -1.3, w: 4.6, h: 4.6, objectName: "deco-orb",
  fill: { color: "E2552B", transparency: 72 } });

// Thẻ kính
const GLASS = {
  fill: { color: "FFFFFF", transparency: 85 },              // 15% trắng
  line: { color: "FFFFFF", transparency: 70, width: 0.75 }, // viền mảnh 30% trắng: "cạnh kính"
  rectRadius: 0.16,
};
slide.addShape(pptx.ShapeType.roundRect, { x: 6.6, y: 3.15, w: 2.9, h: 1.6, ...GLASS });
```

Nền sáng: `fill: { color: "FFFFFF", transparency: 35 }`, `line: { color: "FFFFFF", transparency: 20, width: 0.75 }` cộng shadow rất nhẹ `{ type: "outer", color: "000000", opacity: 0.08, blur: 16, offset: 3, angle: 90 }`.

Quy tắc:
- Chữ trên kính phải kiểm tra tương phản với màu **sau khi trộn** (audit tự trộn các lớp).
- Bo góc đồng tâm: phần tử con trong thẻ có bán kính = bán kính thẻ − khoảng đệm; nếu khoảng đệm ≥ bán kính thì phần tử con dùng góc vuông.
- Không lồng kính trong kính.

## 6. Hệ thống 8 bố cục

| # | Bố cục | Dùng khi | Không dùng khi |
|---|--------|----------|----------------|
| 1 | **Bento Grid** | 3–5 chỉ số/ý có độ quan trọng khác nhau | Các ý ngang hàng tuyệt đối |
| 2 | **Asymmetric Split** | Một ý + một hình ảnh/minh họa mạnh | Không có hình đủ tốt |
| 3 | **Big Number Callout** | Một con số là thông điệp | Có 2+ số cùng quan trọng |
| 4 | **Process Timeline** | 3–6 bước, giai đoạn, lộ trình | Quan hệ không tuyến tính (dùng sơ đồ) |
| 5 | **Spotlight Comparison** | So sánh 2 phương án và có đề xuất | So sánh trung lập nhiều tiêu chí (dùng bảng) |
| 6 | **Full-bleed Hero** | Mở đầu, chuyển chương, khoảnh khắc cảm xúc | Slide nhiều thông tin |
| 7 | **Executive Dashboard** | Báo cáo định kỳ cho lãnh đạo: KPI + xu hướng + việc cần quyết | Khán giả cần hiểu từng con số (tách slide) |
| 8 | **Quote Slide** | Lời khách hàng, chuyên gia, câu chốt | Câu dài hơn 30 từ |

### Wireframe

```
1 Bento Grid                         2 Asymmetric Split
┌──────────────────────────────┐     ┌───────────────┬──────────────┐
│ Tiêu đề                      │     │ EYEBROW       │              │
│ ┌──────────┐ ┌────┐ ┌────┐   │     │ Tiêu đề 36pt  │   ẢNH TRÀN   │
│ │  +38%    │ │1240│ │96% │   │     │ 2–3 dòng      │   MÉP PHẢI   │
│ │          │ └────┘ └────┘   │     │               │              │
│ │          │ ┌──────────┐    │     │ Body ≤3 dòng  │              │
│ └──────────┘ └──────────┘    │     └───────────────┴──────────────┘
└──────────────────────────────┘
3 Big Number                         4 Process Timeline
┌──────────────────────────────┐     ┌──────────────────────────────┐
│ 2,1 ngày   (140pt)           │     │ Tiêu đề                      │
│ ──                           │     │  01      02      03      04  │
│ Câu khẳng định 22pt          │     │  ●───────●───────●-------○   │
│ Nguồn 12pt                   │     │ Bước    Bước    BƯỚC    Bước │
└──────────────────────────────┘     └──────────────────────────────┘
5 Spotlight Comparison               6 Full-bleed Hero
┌──────────────────────────────┐     ┌──────────────────────────────┐
│ Tiêu đề                      │     │░░░░░░░░ ẢNH + LỚP PHỦ ░░░░░░░│
│ ┌────────┐ ┏━━━━━━━━━━━━┓    │     │░░░░░░░░░░░░░░░░░░░░░░░░░░░░░░│
│ │ A (mờ) │ ┃ ĐỀ XUẤT    ┃    │     │ EYEBROW                      │
│ │        │ ┃ (nổi, viền)┃    │     │ Tiêu đề 44pt trắng           │
│ └────────┘ ┗━━━━━━━━━━━━┛    │     └──────────────────────────────┘
└──────────────────────────────┘
7 Executive Dashboard                8 Quote Slide
┌──────────────────────────────┐     ┌──────────────────────────────┐
│ Tiêu đề                      │     │ “                            │
│ [KPI] [KPI] [KPI] [KPI]      │     │    Câu trích 30pt nghiêng    │
│ ┌──────────────────┐ ┌─────┐ │     │    2–3 dòng                  │
│ │ biểu đồ xu hướng │ │Quyết│ │     │    ──                        │
│ └──────────────────┘ └─────┘ │     │    Tên · vai trò             │
└──────────────────────────────┘     └──────────────────────────────┘
```

### Code mẫu đầy đủ (đã kiểm tra: 0 lỗi với `pptx-quality-audit`)

```js
const pptxgen = require("pptxgenjs");

// 1) Token: đổi theo từng brief, KHÔNG dùng lại nguyên bộ này cho mọi deck
const T = {
  W: 10, H: 5.625, M: 0.5, GAP: 0.2,
  font: { head: "Georgia", body: "Arial" },
  size: { title: 40, heading: 22, body: 15, caption: 12 },
  c: {
    bg: "F4F1EA",      // giấy ngà
    surface: "FFFFFF",
    ink: "1B1A17",     // chữ chính
    muted: "5F5A51",   // chữ phụ (≥ 4.5:1 trên bg)
    line: "D8D2C4",
    accent: "1F5C4A",  // xanh rêu đậm: 1 màu nhấn duy nhất
    accentSoft: "DCE8E2",
  },
};
const CW = T.W - 2 * T.M; // 9 inch chiều rộng nội dung
const CONTENT_TOP = 1.35; // dưới tiêu đề
const CONTENT_BOTTOM = T.H - T.M; // 5.125

const text = (s, str, o) => s.addText(str, { margin: 0, valign: "top", fontFace: T.font.body,
  fontSize: T.size.body, color: T.c.ink, ...o });

function title(s, str, o = {}) {
  text(s, str, { x: T.M, y: T.M, w: CW, h: 0.65, fontFace: T.font.head, fontSize: 36, ...o });
}
function eyebrow(s, str, o = {}) {
  text(s, str.toUpperCase(), { x: T.M, y: T.M, w: CW, h: 0.25, fontSize: T.size.caption,
    bold: true, charSpacing: 2, color: T.c.accent, ...o });
}
function card(pptx, s, box, o = {}) {
  s.addShape(pptx.ShapeType.roundRect, { ...box, rectRadius: 0.12,
    fill: { color: T.c.surface }, line: { color: T.c.line, width: 0.75 }, ...o });
}

// Chia lưới cột: col(0, 4) = vị trí x/w của 4 cột đầu trên lưới 12 cột
function col(start, span, cols = 12) {
  const w1 = (CW - (cols - 1) * T.GAP) / cols;
  return { x: T.M + start * (w1 + T.GAP), w: span * w1 + (span - 1) * T.GAP };
}

// ---------------------------------------------------------------- 1. Bento Grid
function bentoGrid(pptx, { heading, tiles }) {
  const s = pptx.addSlide(); s.background = { color: T.c.bg };
  title(s, heading);
  const cols = 4, rows = 2;
  const cw = (CW - (cols - 1) * T.GAP) / cols;
  const rh = (CONTENT_BOTTOM - CONTENT_TOP - (rows - 1) * T.GAP) / rows;
  // tiles: [{ c, r, cs, rs, value, label, accent }]
  tiles.forEach((t) => {
    const box = { x: T.M + t.c * (cw + T.GAP), y: CONTENT_TOP + t.r * (rh + T.GAP),
      w: t.cs * cw + (t.cs - 1) * T.GAP, h: t.rs * rh + (t.rs - 1) * T.GAP };
    card(pptx, s, box, t.accent ? { fill: { color: T.c.accent }, line: { color: T.c.accent, width: 0.75 } } : {});
    const big = t.cs * t.rs >= 4;
    text(s, t.value, { x: box.x + 0.25, y: box.y + 0.22, w: box.w - 0.5, h: big ? 1.1 : 0.55,
      fontFace: T.font.head, fontSize: big ? 60 : 28, bold: true, color: t.accent ? "FFFFFF" : T.c.ink });
    text(s, t.label, { x: box.x + 0.25, y: box.y + (big ? 1.4 : 0.85), w: box.w - 0.5, h: 0.6,
      fontSize: big ? T.size.body : T.size.caption, color: t.accent ? "E6F0EB" : T.c.muted });
  });
  return s;
}

// ---------------------------------------------------------------- 2. Asymmetric Split
function asymmetricSplit(pptx, { kicker, heading, body, image }) {
  const s = pptx.addSlide(); s.background = { color: T.c.bg };
  const left = col(0, 5);
  eyebrow(s, kicker, { x: left.x, y: 0.9, w: left.w });
  text(s, heading, { x: left.x, y: 1.25, w: left.w, h: 1.9, fontFace: T.font.head, fontSize: 36 });
  text(s, body, { x: left.x, y: 3.35, w: left.w, h: 1.25, color: T.c.muted, lineSpacingMultiple: 1.15 });
  // Cột phải tràn mép (tên có "bleed" => audit miễn kiểm tra lề)
  const rx = col(6, 6).x;
  if (image) s.addImage({ path: image, x: rx, y: 0, w: T.W - rx, h: T.H, objectName: "bleed-image",
    sizing: { type: "cover", w: T.W - rx, h: T.H }, altText: heading });
  else s.addShape(pptx.ShapeType.rect, { x: rx, y: 0, w: T.W - rx, h: T.H, fill: { color: T.c.accent }, objectName: "bleed-panel" });
  return s;
}

// ---------------------------------------------------------------- 3. Big Number Callout
function bigNumber(pptx, { number, unit, claim, source }) {
  const s = pptx.addSlide(); s.background = { color: T.c.bg };
  text(s, [
    { text: number, options: { fontSize: 140, bold: true, color: T.c.accent } },
    { text: unit, options: { fontSize: 44, color: T.c.accent } },
  ], { x: T.M, y: 0.8, w: CW, h: 2.3, fontFace: T.font.head });
  s.addShape(pptx.ShapeType.rect, { x: T.M, y: 3.3, w: 0.6, h: 0.05, fill: { color: T.c.ink } });
  text(s, claim, { x: T.M, y: 3.5, w: 6.2, h: 0.8, fontFace: T.font.head, fontSize: T.size.heading });
  text(s, source, { x: T.M, y: 4.6, w: 6.2, h: 0.3, fontSize: T.size.caption, color: T.c.muted });
  return s;
}

// ---------------------------------------------------------------- 4. Process Timeline
function processTimeline(pptx, { heading, steps, current = 0 }) {
  const s = pptx.addSlide(); s.background = { color: T.c.bg };
  title(s, heading);
  const n = steps.length, y = 2.35;
  const stepW = CW / n;
  s.addShape(pptx.ShapeType.line, { x: T.M + stepW / 2, y, w: CW - stepW, h: 0, line: { color: T.c.line, width: 2 } });
  if (current > 0) s.addShape(pptx.ShapeType.line, { x: T.M + stepW / 2, y, w: stepW * current, h: 0,
    line: { color: T.c.accent, width: 2 } });
  steps.forEach((st, i) => {
    const cx = T.M + stepW * i + stepW / 2;
    const on = i <= current;
    s.addShape(pptx.ShapeType.ellipse, { x: cx - 0.16, y: y - 0.16, w: 0.32, h: 0.32,
      fill: { color: on ? T.c.accent : T.c.bg }, line: { color: on ? T.c.accent : T.c.line, width: 2 } });
    text(s, String(i + 1).padStart(2, "0"), { x: cx - stepW / 2 + 0.1, y: 1.55, w: stepW - 0.2, h: 0.35,
      align: "center", fontSize: T.size.caption, bold: true, color: on ? T.c.accent : T.c.muted });
    text(s, st.title, { x: cx - stepW / 2 + 0.1, y: 2.75, w: stepW - 0.2, h: 0.4, align: "center",
      fontFace: T.font.head, fontSize: 20, color: i === current ? T.c.ink : T.c.muted });
    text(s, st.desc, { x: cx - stepW / 2 + 0.15, y: 3.25, w: stepW - 0.3, h: 1.0, align: "center",
      fontSize: 14, color: T.c.muted });
  });
  return s;
}

// ---------------------------------------------------------------- 5. Spotlight Comparison
function spotlightComparison(pptx, { heading, left, right }) {
  // right = phương án được đề xuất: nổi bật; left = phương án còn lại: lùi về sau
  const s = pptx.addSlide(); s.background = { color: T.c.bg };
  title(s, heading);
  const a = col(0, 5), b = col(5, 7);
  const top = CONTENT_TOP, h = CONTENT_BOTTOM - top;
  card(pptx, s, { x: a.x, y: top + 0.25, w: a.w, h: h - 0.25 }, { fill: { color: T.c.bg }, line: { color: T.c.line, width: 0.75 } });
  card(pptx, s, { x: b.x, y: top, w: b.w, h }, { fill: { color: T.c.surface }, line: { color: T.c.accent, width: 1.5 },
    shadow: { type: "outer", color: "000000", opacity: 0.12, blur: 14, offset: 4, angle: 90 } });
  const col_ = (box, d, hot) => {
    text(s, d.label, { x: box.x + 0.3, y: box.y + 0.25, w: box.w - 0.6, h: 0.3, fontSize: T.size.caption, bold: true,
      charSpacing: 1.5, color: hot ? T.c.accent : T.c.muted });
    text(s, d.title, { x: box.x + 0.3, y: box.y + 0.6, w: box.w - 0.6, h: 0.45, fontFace: T.font.head,
      fontSize: hot ? 24 : 20, color: hot ? T.c.ink : T.c.muted });
    text(s, d.points.map((p) => ({ text: p, options: { bullet: { indent: 14 }, breakLine: true } })),
      { x: box.x + 0.3, y: box.y + 1.25, w: box.w - 0.6, h: box.h - 1.5, fontSize: 14,
        color: hot ? T.c.ink : T.c.muted, paraSpaceAfter: 6 });
  };
  col_({ x: a.x, y: top + 0.25, w: a.w, h: h - 0.25 }, left, false);
  col_({ x: b.x, y: top, w: b.w, h }, right, true);
  return s;
}

// ---------------------------------------------------------------- 6. Full-bleed Hero
function fullBleedHero(pptx, { image, kicker, heading }) {
  const s = pptx.addSlide(); s.background = { color: T.c.ink };
  if (image) s.addImage({ path: image, x: 0, y: 0, w: T.W, h: T.H, objectName: "bg-photo",
    sizing: { type: "cover", w: T.W, h: T.H }, altText: heading });
  // Lớp phủ tối dần từ dưới lên giữ chữ đọc được trên mọi ảnh
  s.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: T.W, h: T.H, objectName: "bg-scrim",
    fill: { color: "000000", transparency: 45 } });
  eyebrow(s, kicker, { y: 3.05, color: "FFFFFF" });
  text(s, heading, { x: T.M, y: 3.4, w: 7.5, h: 1.6, fontFace: T.font.head, fontSize: 44, color: "FFFFFF" });
  return s;
}

// ---------------------------------------------------------------- 7. Executive Dashboard
function executiveDashboard(pptx, { heading, kpis, chart, notes }) {
  const s = pptx.addSlide(); s.background = { color: T.c.bg };
  title(s, heading, { fontSize: 32, h: 0.55 });
  const n = kpis.length, kw = (CW - (n - 1) * T.GAP) / n, ky = 1.25, kh = 1.05;
  kpis.forEach((k, i) => {
    const x = T.M + i * (kw + T.GAP);
    card(pptx, s, { x, y: ky, w: kw, h: kh });
    text(s, k.label, { x: x + 0.2, y: ky + 0.15, w: kw - 1.1, h: 0.25, fontSize: T.size.caption, color: T.c.muted });
    text(s, k.delta, { x: x + kw - 0.9, y: ky + 0.15, w: 0.7, h: 0.25, align: "right",
      fontSize: T.size.caption, bold: true, color: k.delta.startsWith("-") ? "A1372B" : T.c.accent });
    text(s, k.value, { x: x + 0.2, y: ky + 0.45, w: kw - 0.4, h: 0.5, fontFace: T.font.head, fontSize: 26, bold: true });
  });
  const cy = ky + kh + T.GAP, chH = CONTENT_BOTTOM - cy;
  const left = col(0, 8), right = col(8, 4);
  card(pptx, s, { x: left.x, y: cy, w: left.w, h: chH });
  s.addChart(pptx.ChartType.line, chart, { x: left.x + 0.15, y: cy + 0.1, w: left.w - 0.3, h: chH - 0.2,
    chartColors: [T.c.accent, "B9B2A3"], lineSize: 2, lineDataSymbol: "none",
    catAxisLabelColor: T.c.muted, valAxisLabelColor: T.c.muted, catAxisLabelFontSize: 12, valAxisLabelFontSize: 12,
    valGridLine: { color: "ECE7DC", size: 0.5 }, catGridLine: { style: "none" },
    showLegend: true, legendPos: "t", legendFontSize: 12, legendColor: T.c.muted });
  card(pptx, s, { x: right.x, y: cy, w: right.w, h: chH }, { fill: { color: T.c.accentSoft }, line: { color: T.c.accentSoft, width: 0.75 } });
  text(s, "Cần quyết định", { x: right.x + 0.2, y: cy + 0.18, w: right.w - 0.4, h: 0.3, fontSize: T.size.caption, bold: true, color: T.c.accent });
  text(s, notes.map((t) => ({ text: t, options: { bullet: { indent: 14 }, breakLine: true } })),
    { x: right.x + 0.2, y: cy + 0.55, w: right.w - 0.4, h: chH - 0.7, fontSize: 14, paraSpaceAfter: 4 });
  return s;
}

// ---------------------------------------------------------------- 8. Quote Slide
function quoteSlide(pptx, { quote, who, role }) {
  const s = pptx.addSlide(); s.background = { color: T.c.accent };
  text(s, "“", { x: T.M, y: T.M, w: 0.8, h: 2.0, fontFace: T.font.head, fontSize: 120, color: "8FB8A8",
    objectName: "deco-quote-mark" });
  text(s, quote, { x: 1.3, y: 1.6, w: 7.4, h: 2.1, fontFace: T.font.head, fontSize: 30, italic: true,
    color: "FFFFFF", lineSpacingMultiple: 1.1 });
  s.addShape(pptx.ShapeType.rect, { x: 1.3, y: 3.95, w: 0.5, h: 0.04, fill: { color: "FFFFFF" } });
  text(s, [
    { text: who, options: { bold: true, breakLine: true } },
    { text: role, options: { color: "CFE3DA" } },
  ], { x: 1.3, y: 4.15, w: 6, h: 0.6, fontSize: 14, color: "FFFFFF" });
  return s;
}
```

Ví dụ gọi:

```js
const pptx = new pptxgen();
pptx.layout = "LAYOUT_16x9";

bentoGrid(pptx, { heading: "Quý 3 trong bốn con số", tiles: [
  { c: 0, r: 0, cs: 2, rs: 2, value: "+38%", label: "Doanh thu so với quý 2, kênh sỉ dẫn đầu", accent: true },
  { c: 2, r: 0, cs: 1, rs: 1, value: "1.240", label: "kg cà phê đã rang" },
  { c: 3, r: 0, cs: 1, rs: 1, value: "96%", label: "đơn giao đúng hẹn" },
  { c: 2, r: 1, cs: 2, rs: 1, value: "61%", label: "đơn mới đến từ quán cà phê" },
] });
asymmetricSplit(pptx, { kicker: "Chương 2", heading: "Rang theo mẻ nhỏ, giao trong 48 giờ",
  body: "Mỗi mẻ 12 kg, rang theo đơn. Hạt đến tay khách khi còn đủ 7–21 ngày nghỉ khí.", image: "assets/roastery.jpg" });
bigNumber(pptx, { number: "2,1", unit: " ngày", claim: "Thời gian giao trung bình, nhanh gấp đôi năm ngoái",
  source: "Nguồn: hệ thống đơn hàng nội bộ, 07–09/2026" });
processTimeline(pptx, { heading: "Từ hạt xanh đến ly cà phê", current: 2, steps: [
  { title: "Nhập hạt", desc: "Kiểm độ ẩm 10–12%" }, { title: "Rang", desc: "Mẻ 12 kg, ghi hồ sơ nhiệt" },
  { title: "Nghỉ khí", desc: "7 ngày trong túi van" }, { title: "Giao", desc: "Trong 48 giờ" } ] });
spotlightComparison(pptx, { heading: "Hai cách mở rộng công suất",
  left: { label: "PHƯƠNG ÁN A", title: "Thuê ca đêm", points: ["Không cần vốn đầu tư", "Chi phí nhân sự +35%", "Rủi ro chất lượng ca đêm"] },
  right: { label: "ĐỀ XUẤT", title: "Mua máy rang 30 kg", points: ["Hoàn vốn sau 14 tháng", "Công suất gấp 2,5 lần", "Giữ nguyên một ca làm việc"] } });
fullBleedHero(pptx, { image: "assets/dalat.jpg", kicker: "Kế hoạch 2027", heading: "Mở xưởng thứ hai\nở Đà Lạt" });
executiveDashboard(pptx, { heading: "Tổng quan vận hành tháng 9",
  kpis: [ { label: "Doanh thu", value: "664 tr", delta: "+28%" }, { label: "Đơn hàng", value: "1.932", delta: "+12%" },
          { label: "Tỷ lệ hoàn", value: "0,8%", delta: "-0,3 đ" }, { label: "Biên lãi", value: "31%", delta: "+2 đ" } ],
  chart: [ { name: "2026", labels: ["T4","T5","T6","T7","T8","T9"], values: [301, 334, 360, 412, 518, 664] },
           { name: "2025", labels: ["T4","T5","T6","T7","T8","T9"], values: [250, 262, 270, 288, 300, 322] } ],
  notes: ["Duyệt mua máy rang 30 kg", "Chốt nhà cung cấp túi van quý 4"] });
quoteSlide(pptx, { quote: "Từ khi đổi sang Mộc, khách quen hỏi tên hạt cà phê nhiều hơn hỏi giá.",
  who: "Chị Lan", role: "Chủ quán Góc Phố, Quận 3" });

await pptx.writeFile({ fileName: "output/deck.pptx" });
```

Đây là điểm xuất phát, không phải khuôn đúc: đổi token theo brief, đổi tỷ lệ bento, đổi vị trí số lớn. Giữ lại lưới, lề và thang chữ.

## 7. Quy trình thiết kế một deck

1. **Xác định**: khán giả, mục đích (thuyết phục / báo cáo / dạy), thời lượng, chiếu hay gửi đọc.
2. **Viết thông điệp từng slide trước** thành một câu khẳng định. Câu đó thường chính là tiêu đề.
3. **Lập design plan** (ghi trong câu trả lời trước khi code):
   - 4–6 màu hex kèm vai trò và lý do;
   - cặp font và vai trò;
   - danh sách slide → bố cục (không 2 slide liền nhau cùng bố cục, trừ cặp Morph);
   - một "chữ ký" riêng cho deck (vd: thanh tiến trình đỏ son, số thứ tự serif lớn, quả cầu trôi).
4. **Tự phản biện**: đối chiếu plan với bảng lối mòn ở mục 4. Thay ít nhất một lựa chọn "mặc định".
5. **Code** với token + hàm bố cục; ghi speaker notes (`slide.addNotes`) cho từng slide.
6. **Kiểm tra** bằng `pptx-quality-audit` (script + ảnh xem trước), sửa đến khi 0 ERROR và WARN đã được giải thích.

## 8. Checklist nhanh

- [ ] `LAYOUT_16x9`, mọi phần tử nội dung nằm trong lề 0.5 inch
- [ ] Title 36–44pt, Heading 20–24pt, Body 14–16pt, không chữ < 12pt
- [ ] ≤ 6 dòng chữ nội dung mỗi slide, ≤ 3 cỡ chữ
- [ ] Một màu nhấn, dùng cho đúng một thứ mỗi slide
- [ ] Tương phản chữ ≥ 4.5:1 (≥ 3:1 cho chữ lớn)
- [ ] Không 2 slide liền nhau cùng bố cục (trừ cặp Morph)
- [ ] Ảnh dùng `sizing: { type: "cover" }` hoặc tính `h = w / tỷ lệ ảnh`, không bóp méo
- [ ] Mọi ảnh có `altText`, mọi slide có speaker notes
