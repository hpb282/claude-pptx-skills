---
name: pptx-charts-diagrams
description: Vẽ biểu đồ chuyên nghiệp (Bar, Line, Doughnut, Radar) và sơ đồ khối (Architecture/Flow diagram) trong file .pptx bằng pptxgenjs và SVG-to-PNG. Kỹ thuật Progressive Reveal - tách một biểu đồ hoặc sơ đồ phức tạp thành chuỗi 3-4 slide liên tiếp, mỗi slide tô sáng một cột/nhánh/series và làm mờ phần còn lại để dẫn mắt người xem, kết hợp Morph. Dùng khi người dùng nhắc @chart-pptx, biểu đồ PowerPoint, chart trong slide, sơ đồ kiến trúc, flow diagram, data storytelling, xuất hiện từng bước, highlight dữ liệu.
---

# PPTX Charts & Diagrams

Một biểu đồ trên slide có đúng **một** việc: chứng minh câu tiêu đề. Nếu người xem phải tự tìm điểm đáng chú ý, biểu đồ đó chưa xong.

Dùng cùng `pptx-visual-design` (token màu, chữ, lưới), `pptx-morph-motion` (chuyển cảnh giữa các bước reveal) và kiểm tra bằng `pptx-quality-audit`.

## 1. Chọn loại biểu đồ

| Câu hỏi của slide | Loại | pptxgenjs |
|-------------------|------|-----------|
| So sánh giá trị giữa các nhóm | Bar (cột đứng / ngang) | `ChartType.bar`, `barDir: "col"` hoặc `"bar"` |
| Xu hướng theo thời gian | Line | `ChartType.line` |
| Tỷ trọng của một tổng (≤ 5 phần) | Doughnut | `ChartType.doughnut` |
| So sánh nhiều tiêu chí của 2–3 đối tượng | Radar | `ChartType.radar` |
| Một con số duy nhất | Không dùng chart: Big Number Callout | `pptx-visual-design` |
| > 7 nhóm, nhãn dài | Bar ngang, sắp xếp giảm dần | `barDir: "bar"` |
| Quan hệ, luồng, kiến trúc | Sơ đồ khối | shape + đường nối, hoặc SVG → PNG |

Không dùng: pie 3D, bar 3D, doughnut > 5 lát, radar > 3 series, line > 4 series, hai trục tung.

## 2. Quy tắc trình bày

- **Tiêu đề là kết luận**, không phải tên biểu đồ: "Tháng 9 cao nhất: 664 triệu", không phải "Doanh thu theo tháng".
- **Một màu nhấn** cho dữ liệu cần chú ý; phần còn lại màu `dim` trung tính.
- Xóa thứ không mang thông tin: khung chart, lưới dọc, trục tung khi đã có nhãn giá trị, legend khi chỉ có một series.
- Nhãn trục, nhãn dữ liệu, legend **≥ 12pt** (`catAxisLabelFontSize`, `dataLabelFontSize`, `legendFontSize`).
- Định dạng số kiểu Việt Nam trong chữ: `1.240`, `2,1`. Với `dataLabelFormatCode` dùng `"#,##0"` (PowerPoint hiển thị dấu phân cách theo locale máy chiếu).
- Ghi đơn vị và nguồn: "(triệu đồng)", "Nguồn: …" cỡ 12pt.
- Biểu đồ đặt trong vùng nội dung (y ≥ 1.35, trong lề 0.5 inch). Đặt `objectName` có `!!` để Morph ghép chart giữa các bước.

## 3. Code mẫu (đã kiểm tra: 0 lỗi với `pptx-quality-audit`)

```js
const pptxgen = require("pptxgenjs");

const T = {
  M: 0.5, W: 10, H: 5.625,
  font: { head: "Georgia", body: "Arial" },
  c: { bg: "F4F1EA", ink: "1B1A17", muted: "5F5A51", dim: "C9C3B6", line: "D8D2C4", accent: "1F5C4A", accent2: "C4622D" },
};
const text = (s, str, o) => s.addText(str, { margin: 0, valign: "top", fontFace: T.font.body, fontSize: 15, color: T.c.ink, ...o });
const title = (s, str) => text(s, str, { x: T.M, y: T.M, w: 9, h: 0.65, fontFace: T.font.head, fontSize: 36 });

// Kiểu chung cho mọi chart: bỏ khung, bỏ lưới dọc, chữ trục 12pt màu phụ
const BASE = {
  fontFace: T.font.body,
  catAxisLabelColor: T.c.muted, valAxisLabelColor: T.c.muted,
  catAxisLabelFontSize: 12, valAxisLabelFontSize: 12,
  catAxisLineShow: false, valAxisLineShow: false,
  valGridLine: { color: "E4DFD3", size: 0.5 }, catGridLine: { style: "none" },
  showLegend: false,
};

// ------------------------------------------------------------ Bar (cột)
function barChart(slide, pptx, { labels, values, focus = -1, box }) {
  // Một series + mảng chartColors nhiều màu => pptxgenjs tô màu TỪNG CỘT
  const colors = values.map((_, i) => (focus < 0 || i === focus ? T.c.accent : T.c.dim));
  slide.addChart(pptx.ChartType.bar, [{ name: "Doanh thu", labels, values }], {
    ...BASE, ...box, objectName: "!!chart-main",
    barDir: "col", barGapWidthPct: 55, chartColors: colors,
    showValue: true, dataLabelPosition: "outEnd", dataLabelColor: T.c.ink, dataLabelFontSize: 12,
    dataLabelFormatCode: "#,##0", valAxisHidden: true, valGridLine: { style: "none" },
  });
}

// ------------------------------------------------------------ Line (đường)
function lineChart(slide, pptx, { series, focus = -1, box }) {
  const palette = [T.c.accent, T.c.accent2, "6B7F99", "8C7B5E"];
  const colors = series.map((_, i) => (focus < 0 || i === focus ? palette[i % palette.length] : T.c.dim));
  slide.addChart(pptx.ChartType.line, series, {
    ...BASE, ...box, objectName: "!!chart-main",
    chartColors: colors, lineSize: 2.5, lineDataSymbol: "circle", lineDataSymbolSize: 6,
    showLegend: true, legendPos: "t", legendFontSize: 12, legendColor: T.c.muted,
  });
}

// ------------------------------------------------------------ Doughnut (bánh vòng)
function doughnutChart(slide, pptx, { labels, values, focus = 0, box, centerLabel }) {
  const colors = values.map((_, i) => (i === focus ? T.c.accent : T.c.dim));
  slide.addChart(pptx.ChartType.doughnut, [{ name: "Cơ cấu", labels, values }], {
    ...box, objectName: "!!chart-main", holeSize: 68, chartColors: colors,
    dataBorder: { pt: 1.5, color: T.c.bg }, showLegend: false, showPercent: false, showValue: false,
  });
  // Nhãn giữa lỗ: con số của phần đang tô sáng
  const pct = Math.round((values[focus] / values.reduce((a, b) => a + b, 0)) * 100);
  text(slide, [
    { text: `${pct}%`, options: { fontSize: 32, bold: true, color: T.c.accent, breakLine: true } },
    { text: centerLabel ?? labels[focus], options: { fontSize: 12, color: T.c.muted } },
  ], { x: box.x + box.w / 2 - 1, y: box.y + box.h / 2 - 0.45, w: 2, h: 0.9, align: "center",
    fontFace: T.font.head, objectName: "!!chart-center" });
}

// ------------------------------------------------------------ Radar (mạng nhện)
function radarChart(slide, pptx, { series, focus = -1, box }) {
  const palette = [T.c.accent, T.c.accent2, "6B7F99"];
  const colors = series.map((_, i) => (focus < 0 || i === focus ? palette[i % palette.length] : T.c.dim));
  slide.addChart(pptx.ChartType.radar, series, {
    ...box, objectName: "!!chart-main", radarStyle: "marker", chartColors: colors, lineSize: 2,
    catAxisLabelColor: T.c.muted, catAxisLabelFontSize: 12, valAxisHidden: true,
    valGridLine: { color: "E4DFD3", size: 0.5 },
    showLegend: true, legendPos: "b", legendFontSize: 12, legendColor: T.c.muted,
  });
}

// ------------------------------------------------------------ Progressive Reveal
// Một biểu đồ → chuỗi slide. Mỗi bước tô sáng 1 phần, làm mờ phần còn lại,
// tiêu đề đổi thành câu nói về đúng phần đang sáng.
function progressiveReveal(pptx, steps, draw) {
  steps.forEach((step, i) => {
    const s = pptx.addSlide();
    s.background = { color: T.c.bg };
    title(s, step.heading);
    draw(s, step.focus, i);
    // chấm chỉ bước: người xem biết còn bao nhiêu bước
    steps.forEach((_, k) => s.addShape(pptx.ShapeType.ellipse, {
      x: 9.5 - (steps.length - k) * 0.22, y: 4.98, w: 0.12, h: 0.12, objectName: `!!step-dot-${k}`,
      fill: { color: k === i ? T.c.accent : T.c.dim } }));
    if (step.note) s.addNotes(step.note);
  });
}

// ------------------------------------------------------------ Sơ đồ khối (Architecture / Flow)
// node: { id, label, x, y, w?, h? }  edge: [fromId, toId]
function diagram(slide, pptx, { nodes, edges, active = null }) {
  const byId = Object.fromEntries(nodes.map((n) => [n.id, { w: 1.7, h: 0.7, ...n }]));
  const on = (id) => !active || active.includes(id);

  edges.forEach(([a, b], k) => {
    const A = byId[a], B = byId[b];
    const lit = on(a) && on(b);
    const ln = { color: lit ? T.c.accent : T.c.dim, width: lit ? 2 : 1.25, endArrowType: "triangle" };
    const ax = A.x + A.w, ay = A.y + A.h / 2, bx = B.x, by = B.y + B.h / 2;
    if (Math.abs(ay - by) < 0.01) {
      slide.addShape(pptx.ShapeType.line, { x: ax, y: ay, w: bx - ax, h: 0, line: ln, objectName: `!!edge-${k}` });
    } else {
      // đường gấp khúc vuông góc: ngang → dọc → ngang (pptxgenjs chưa có elbow connector)
      const mx = ax + (bx - ax) / 2;
      const seg = { color: ln.color, width: ln.width };
      slide.addShape(pptx.ShapeType.line, { x: ax, y: ay, w: mx - ax, h: 0, line: seg, objectName: `!!edge-${k}-a` });
      slide.addShape(pptx.ShapeType.line, { x: mx, y: Math.min(ay, by), w: 0, h: Math.abs(by - ay), line: seg,
        objectName: `!!edge-${k}-b` });
      slide.addShape(pptx.ShapeType.line, { x: mx, y: by, w: bx - mx, h: 0, line: ln, objectName: `!!edge-${k}-c` });
    }
  });

  Object.values(byId).forEach((n) => {
    const lit = on(n.id);
    // Text gắn liền shape = 1 đối tượng => Morph ghép gọn, chữ không lệch khỏi khung
    slide.addText(n.label, {
      shape: pptx.ShapeType.roundRect, rectRadius: 0.1, x: n.x, y: n.y, w: n.w, h: n.h,
      objectName: `!!node-${n.id}`, align: "center", valign: "middle", margin: 4,
      fontFace: T.font.body, fontSize: 14, bold: lit,
      color: lit ? "FFFFFF" : T.c.muted,
      fill: { color: lit ? T.c.accent : "FFFFFF" },
      line: { color: lit ? T.c.accent : T.c.line, width: 0.75 },
    });
  });
}
```

### Ví dụ dùng

```js
const pptx = new pptxgen();
pptx.layout = "LAYOUT_16x9";
const months = ["T4", "T5", "T6", "T7", "T8", "T9"];
const rev = [301, 334, 360, 412, 518, 664];

// Progressive Reveal 3 bước trên một biểu đồ cột
progressiveReveal(pptx, [
  { heading: "Doanh thu tăng đều 6 tháng", focus: -1 },            // toàn cảnh, mọi cột cùng màu
  { heading: "Tháng 7: kênh sỉ bắt đầu chạy", focus: 3 },          // tô sáng cột T7
  { heading: "Tháng 9 cao nhất: 664 triệu", focus: 5, note: "Dừng ở đây, hỏi khán giả lý do" },
], (s, focus) => barChart(s, pptx, { labels: months, values: rev, focus,
  box: { x: 0.5, y: 1.35, w: 9, h: 3.35 } }));

// Line: tô sáng series năm nay, làm mờ năm ngoái
const s1 = pptx.addSlide();
lineChart(s1, pptx, { box: { x: 0.5, y: 1.35, w: 9, h: 3.5 }, focus: 0, series: [
  { name: "2026", labels: months, values: rev },
  { name: "2025", labels: months, values: [250, 262, 270, 288, 300, 322] },
] });

// Doughnut: một lát sáng + số phần trăm ở giữa lỗ
const s2 = pptx.addSlide();
doughnutChart(s2, pptx, { box: { x: 0.5, y: 1.3, w: 4.6, h: 3.8 },
  labels: ["Quán cà phê", "Bán lẻ", "Văn phòng"], values: [61, 27, 12], focus: 0 });

// Radar: đối tượng của mình sáng, đối thủ mờ
const s3 = pptx.addSlide();
const crit = ["Độ tươi", "Giá", "Giao hàng", "Đa dạng", "Hỗ trợ"];
radarChart(s3, pptx, { box: { x: 1.5, y: 1.3, w: 7, h: 3.8 }, focus: 0, series: [
  { name: "Mộc", labels: crit, values: [9, 5, 8, 6, 8] },
  { name: "Đối thủ", labels: crit, values: [6, 8, 6, 7, 5] },
] });
```

## 4. Progressive Reveal (xuất hiện từng bước)

Một biểu đồ phức tạp chiếu một lần thì người xem đọc lung tung. Tách thành **3–4 slide**, cùng bố cục, cùng tọa độ, chỉ khác phần được tô sáng:

| Bước | Nội dung | Tiêu đề |
|------|----------|---------|
| 1. Toàn cảnh | Mọi phần cùng màu (hoặc cùng màu mờ) | Bối cảnh chung |
| 2. Điểm thứ nhất | Tô sáng cột/nhánh A, làm mờ phần còn lại | Câu về A |
| 3. Điểm thứ hai | Tô sáng B | Câu về B |
| 4. Kết luận (tùy chọn) | Tô sáng phần quyết định + chú thích | "Vì vậy…" |

Quy tắc:

1. **Không xê dịch** biểu đồ giữa các bước: `box` giống hệt nhau, chỉ đổi màu/độ mờ. Mắt người xem đứng yên, chỉ điểm sáng di chuyển.
2. **Làm mờ bằng màu `dim`** (xám ấm gần nền) hoặc `transparency: 60–75` cho shape. Không ẩn hẳn phần còn lại: người xem cần bối cảnh.
3. **Tiêu đề đổi theo bước**, nói đúng về phần đang sáng.
4. **Chấm chỉ bước** (`!!step-dot-k`) góc dưới phải cho biết còn bao nhiêu bước.
5. **Chuyển cảnh**: bật Morph cho slide 2–4 (hàm `injectMorph` trong `pptx-morph-motion`). Sơ đồ dựng bằng shape có tên `!!node-*`, `!!edge-*` sẽ chuyển màu mượt; chart native cũng được Morph nội suy khi cùng loại và cùng `!!chart-main`. Nếu chart nhấp nháy khi Morph, đổi sang Fade 0.3 giây.
6. **Cột mọc lên** thay vì đổi màu: dựng cột bằng shape (xem `pptx-morph-motion` mục "Cột biểu đồ mọc lên").

### Reveal cho sơ đồ khối

```js
const nodes = [
  { id: "app",   label: "Ứng dụng đặt hàng", x: 0.5, y: 2.6 },
  { id: "api",   label: "API đơn hàng",      x: 2.9, y: 2.6 },
  { id: "queue", label: "Hàng đợi",          x: 5.3, y: 1.6 },
  { id: "roast", label: "Lịch rang",         x: 5.3, y: 3.6 },
  { id: "wh",    label: "Kho dữ liệu",       x: 7.8, y: 2.6 },
];
const edges = [["app", "api"], ["api", "queue"], ["api", "roast"], ["queue", "wh"], ["roast", "wh"]];

[
  { heading: "Đơn hàng đi qua 5 khối",              active: null },
  { heading: "Bước 1: khách đặt, API ghi nhận",     active: ["app", "api"] },
  { heading: "Bước 2: API xếp lịch rang",           active: ["api", "roast"] },
  { heading: "Bước 3: mọi sự kiện về kho dữ liệu",  active: ["queue", "roast", "wh"] },
].forEach((st) => {
  const s = pptx.addSlide();
  s.background = { color: T.c.bg };
  title(s, st.heading);
  diagram(s, pptx, { nodes, edges, active: st.active });
});
```

Nút sáng: nền màu nhấn, chữ trắng đậm. Nút mờ: nền trắng, viền `line`, chữ `muted`. Cạnh chỉ sáng khi cả hai đầu cùng sáng.

### Bố trí sơ đồ khối

- Luồng chính đi **trái → phải** (hoặc trên → dưới), không quá 6 nút trên một hàng.
- Nút cùng cấp có cùng `w`, `h` và thẳng hàng theo `x` hoặc `y`.
- Khoảng cách ngang giữa các nút ≥ 0.6 inch để mũi tên đọc được.
- Nhãn nút ≤ 3 từ, 14pt. Giải thích dài đưa vào speaker notes hoặc tiêu đề bước.
- Đường nối vuông góc (ngang → dọc → ngang) gọn hơn đường chéo. pptxgenjs chưa có elbow connector, nên hàm `diagram` vẽ 3 đoạn thẳng.

## 5. SVG → PNG cho sơ đồ phức tạp

Khi sơ đồ cần đường cong, icon, nhóm lồng nhau, hoặc đã có sẵn SVG (từ Mermaid, Figma, draw.io): vẽ SVG, chuyển sang PNG độ phân giải cao, rồi `addImage`. Mất khả năng sửa trong PowerPoint, đổi lại hình chính xác tuyệt đối.

Cách 1: Playwright (Chromium vẽ SVG giống trình duyệt, chữ tiếng Việt chuẩn)

```js
const fs = require("fs");
const { chromium } = require("playwright"); // npm i -D playwright && npx playwright install chromium

async function svgToPng(svgPath, pngPath, scale = 3) {
  const svg = fs.readFileSync(svgPath, "utf8");
  const [, , vw, vh] = svg.match(/viewBox="([^"]+)"/)[1].trim().split(/[\s,]+/).map(Number);
  const browser = await chromium.launch();
  const page = await browser.newPage({ viewport: { width: Math.ceil(vw), height: Math.ceil(vh) }, deviceScaleFactor: scale });
  await page.setContent(`<!doctype html><meta charset="utf-8"><body style="margin:0;background:transparent">${svg}</body>`);
  await page.locator("svg").first().screenshot({ path: pngPath, omitBackground: true }); // PNG nền trong suốt
  await browser.close();
  return { ratio: vh / vw };
}

const { ratio } = await svgToPng("diagrams/flow.svg", "media/flow.png");
slide.addImage({ path: "media/flow.png", x: 0.5, y: 1.4, w: 9, h: 9 * ratio, altText: "Sơ đồ luồng dữ liệu" });
```

Cách 2: `sharp` (nhẹ, không cần trình duyệt)

```js
const sharp = require("sharp"); // npm i sharp
await sharp(fs.readFileSync("diagrams/flow.svg"), { density: 288 }) // 288 dpi = 4x
  .png().toFile("media/flow.png");
```

Cách 3: Python `cairosvg.svg2png(url="flow.svg", write_to="flow.png", scale=3)` hoặc CLI `rsvg-convert -z 3 flow.svg -o flow.png`.

Quy tắc SVG cho slide:

- Luôn khai báo `viewBox`; tính `h = w * vh / vw` khi đặt ảnh để không méo (audit báo `MEDIA` nếu méo).
- Độ phân giải: ít nhất 2× kích thước hiển thị. Ảnh rộng 9 inch → ≥ 2700 px (scale 3 với viewBox 900).
- Dùng đúng mã màu token của deck, `font-family` là font có trên máy render (`Arial, sans-serif`).
- Chữ trong SVG ≥ 20px ở viewBox 900 px rộng (≈ 14pt khi đặt rộng 9 inch).
- Progressive Reveal với SVG: xuất một PNG cho mỗi bước (đổi `fill`/`opacity` theo bước), đặt cùng `x, y, w, h` và cùng `objectName: "!!diagram"` để Morph crossfade.

## 6. Checklist

- [ ] Tiêu đề là kết luận rút ra từ dữ liệu
- [ ] Một màu nhấn cho điểm chính, còn lại `dim`
- [ ] Không lưới dọc, không khung, không 3D, không legend thừa
- [ ] Nhãn ≥ 12pt, có đơn vị và nguồn
- [ ] Reveal: cùng `box` ở mọi bước, chỉ đổi phần sáng, có chấm chỉ bước
- [ ] Ảnh SVG→PNG đúng tỷ lệ, ≥ 2× độ phân giải, có `altText`
