#!/usr/bin/env node
/**
 * morph-deck-generator.js
 * ------------------------------------------------------------------
 * Deck mẫu 4 slide (16:9) minh họa Morph Transition + Bento Grid.
 *
 *   Slide 1  Full-bleed Hero       – tiêu đề lớn, thẻ kính nhỏ
 *   Slide 2  Big Number Callout    – thẻ kính phóng to (zoom), con số lớn
 *   Slide 3  Bento Grid            – thẻ co lại thành ô bento, thêm 3 ô mới
 *   Slide 4  Spotlight             – ô biểu đồ phóng to, cột mọc lên, thẻ lật
 *
 * Mọi phần tử dùng chung giữa các slide có objectName bắt đầu bằng "!!".
 * PowerPoint Morph ghép các phần tử cùng tên giữa slide N và N+1 rồi nội suy
 * x, y, w, h, rotate, màu và độ trong suốt => chuyển động mượt như video.
 *
 * Chạy:
 *   npm install            # cài pptxgenjs (đã khai báo trong package.json)
 *   node templates/morph-deck-generator.js [output.pptx] [--no-morph]
 *
 * Mặc định script tự chèn Morph transition vào XML của slide 2–4.
 * Dùng --no-morph nếu muốn tự bật trong PowerPoint (Transitions -> Morph).
 */

const fs = require("fs");
const path = require("path");
const pptxgen = require("pptxgenjs");
const JSZip = require("jszip"); // có sẵn: là dependency của pptxgenjs

// ------------------------------------------------------------------
// 1. Design tokens
// ------------------------------------------------------------------
const W = 10; // LAYOUT_16x9: 10 x 5.625 inch
const H = 5.625;
const M = 0.5; // lề an toàn

const C = {
  bg: "16130F", // espresso – nền
  paper: "F3EADB", // chữ chính
  muted: "A89F90", // chữ phụ (tương phản 7:1 trên nền)
  accent: "E2552B", // đỏ son – điểm nhấn duy nhất
  sage: "8FB39A", // màu phụ cho trang trí
};

// Font có sẵn trên Windows + macOS và hỗ trợ tiếng Việt.
const F = { head: "Georgia", body: "Arial" };

// Bề mặt kính (glassmorphism giả lập)
const GLASS = {
  fill: { color: "FFFFFF", transparency: 85 },
  line: { color: "FFFFFF", transparency: 70, width: 0.75 },
  rectRadius: 0.16,
};

const SLIDE_COUNT = 4;

// ------------------------------------------------------------------
// 2. Kịch bản chuyển động: mỗi phần tử khai báo 1 lần, kèm trạng thái
//    trên từng slide. null = không có mặt (Morph sẽ fade in/out).
// ------------------------------------------------------------------
const bar = (x, h, base, extra = {}) => ({ x, y: base - h, w: 1.0, h, ...extra });

const ELEMENTS = [
  // --- Trang trí nền: quả cầu mờ phóng to / trôi giữa các slide (Zoom/Pan)
  {
    name: "!!deco-orb-a",
    kind: "shape",
    shape: "ellipse",
    style: { fill: { color: C.accent, transparency: 72 } },
    states: [
      { x: 6.4, y: -1.3, w: 4.6, h: 4.6 },
      { x: 4.6, y: -2.6, w: 7.6, h: 7.6, fill: { color: C.accent, transparency: 86 } },
      { x: 7.6, y: 3.3, w: 3.4, h: 3.4 },
      { x: -1.2, y: 3.6, w: 2.8, h: 2.8, fill: { color: C.accent, transparency: 82 } },
    ],
  },
  {
    name: "!!deco-orb-b",
    kind: "shape",
    shape: "ellipse",
    style: { fill: { color: C.sage, transparency: 78 } },
    states: [
      { x: -1.1, y: 3.5, w: 3.2, h: 3.2 },
      { x: 7.9, y: 3.9, w: 2.4, h: 2.4 },
      { x: -1.4, y: -1.4, w: 3.0, h: 3.0 },
      { x: 7.4, y: -1.6, w: 3.6, h: 3.6 },
    ],
  },

  // --- Chữ
  {
    name: "!!eyebrow",
    kind: "text",
    text: "BÁO CÁO VẬN HÀNH · QUÝ 3/2026",
    style: { fontFace: F.body, fontSize: 12, color: C.muted, bold: true, charSpacing: 2 },
    states: [{ x: M, y: 0.5, w: 6, h: 0.3 }, null, null, null],
  },
  {
    name: "!!main-title",
    kind: "text",
    text: "Một quý rang chậm,\nbán nhanh hơn.",
    style: { fontFace: F.head, fontSize: 44, color: C.paper },
    states: [
      { x: M, y: 1.3, w: 7.6, h: 1.7 },
      // Big Number: tiêu đề lùi về vai trò heading để con số làm nhân vật chính
      { x: M, y: 0.5, w: 9, h: 0.45, fontSize: 20, text: "Một quý rang chậm, bán nhanh hơn." },
      null,
      null,
    ],
  },
  {
    name: "!!subtitle",
    kind: "text",
    text: "Xưởng rang Mộc · số liệu minh họa",
    style: { fontFace: F.body, fontSize: 16, color: C.muted },
    states: [{ x: M, y: 3.3, w: 5.6, h: 0.4 }, null, null, null],
  },
  {
    name: "!!section-title",
    kind: "text",
    text: "Ba con số kể lại cả quý",
    style: { fontFace: F.head, fontSize: 36, color: C.paper },
    states: [
      null,
      null,
      { x: M, y: 0.5, w: 9, h: 0.6 },
      { x: M, y: 0.5, w: 9, h: 0.6, text: "Tháng 9 gánh cả quý" },
    ],
  },

  // --- Thẻ kính chính: nhỏ -> phóng to -> ô bento -> ô nhỏ bên phải
  {
    name: "!!hero-card",
    kind: "shape",
    shape: "roundRect",
    style: { ...GLASS },
    states: [
      { x: 6.6, y: 3.15, w: 2.9, h: 1.6 },
      { x: M, y: 1.2, w: 9, h: 3.65 },
      { x: M, y: 1.3, w: 4.4, h: 3.55 },
      { x: 6.3, y: 1.3, w: 3.2, h: 1.1 },
    ],
  },
  {
    name: "!!hero-kpi",
    kind: "text",
    text: "+38%",
    style: { fontFace: F.head, fontSize: 36, color: C.accent, bold: true },
    states: [
      { x: 6.85, y: 3.3, w: 2.4, h: 0.7 },
      { x: 0.9, y: 1.5, w: 5.2, h: 1.6, fontSize: 100 },
      { x: 0.8, y: 1.6, w: 3.8, h: 1.0, fontSize: 60 },
      { x: 6.5, y: 1.45, w: 1.4, h: 0.8, fontSize: 28 },
    ],
  },
  {
    name: "!!hero-kpi-label",
    kind: "text",
    text: "doanh thu so với quý 2",
    style: { fontFace: F.body, fontSize: 12, color: C.paper },
    states: [
      { x: 6.85, y: 4.0, w: 2.4, h: 0.5 },
      { x: 0.9, y: 3.3, w: 5.2, h: 0.4, fontSize: 16 },
      { x: 0.8, y: 2.6, w: 3.8, h: 0.4, fontSize: 16 },
      { x: 7.95, y: 1.45, w: 1.4, h: 0.8 },
    ],
  },
  {
    name: "!!hero-note",
    kind: "text",
    text: "Kênh bán sỉ cho quán cà phê chiếm 61% đơn mới. Lô rang thử nghiệm bán hết trong 4 ngày.",
    style: { fontFace: F.body, fontSize: 14, color: C.muted },
    states: [
      null,
      { x: 6.3, y: 1.6, w: 2.9, h: 1.3, color: C.paper }, // nằm trên quả cầu cam: dùng màu sáng hơn
      { x: 0.8, y: 3.45, w: 3.8, h: 1.1 },
      null,
    ],
  },

  // --- Các ô bento phụ
  {
    name: "!!tile-2",
    kind: "shape",
    shape: "roundRect",
    style: { ...GLASS },
    states: [null, null, { x: 5.1, y: 1.3, w: 2.1, h: 1.65 }, { x: 6.3, y: 2.525, w: 3.2, h: 1.1 }],
  },
  {
    name: "!!tile-2-num",
    kind: "text",
    text: "1.240 kg",
    style: { fontFace: F.head, fontSize: 24, color: C.paper, bold: true },
    states: [null, null, { x: 5.3, y: 1.5, w: 1.75, h: 0.5 }, { x: 6.5, y: 2.675, w: 1.4, h: 0.8, fontSize: 20 }],
  },
  {
    name: "!!tile-2-label",
    kind: "text",
    text: "cà phê đã rang",
    style: { fontFace: F.body, fontSize: 12, color: C.muted },
    states: [null, null, { x: 5.3, y: 2.05, w: 1.75, h: 0.6 }, { x: 7.95, y: 2.675, w: 1.4, h: 0.8 }],
  },

  // Thẻ lật: mặt trước bóp về 0 ở tâm, mặt sau nở ra từ tâm (giả lập flip)
  {
    name: "!!tile-3",
    kind: "shape",
    shape: "roundRect",
    style: { ...GLASS },
    states: [
      null,
      null,
      { x: 7.4, y: 1.3, w: 2.1, h: 1.65 },
      { x: 7.89, y: 3.75, w: 0.02, h: 1.1, fill: { color: "FFFFFF", transparency: 100 }, line: { color: "FFFFFF", transparency: 100, width: 0.75 } },
    ],
  },
  {
    name: "!!tile-3-num",
    kind: "text",
    text: "96%",
    style: { fontFace: F.head, fontSize: 24, color: C.paper, bold: true },
    states: [null, null, { x: 7.6, y: 1.5, w: 1.75, h: 0.5 }, { x: 7.89, y: 3.9, w: 0.02, h: 0.5, transparency: 100 }],
  },
  {
    name: "!!tile-3-label",
    kind: "text",
    text: "đơn giao đúng hẹn",
    style: { fontFace: F.body, fontSize: 12, color: C.muted },
    states: [null, null, { x: 7.6, y: 2.05, w: 1.75, h: 0.6 }, null],
  },
  {
    name: "!!tile-3-back",
    kind: "shape",
    shape: "roundRect",
    style: { fill: { color: C.accent, transparency: 0 }, line: { color: C.accent, width: 0.75 }, rectRadius: 0.16 },
    states: [
      null,
      null,
      { x: 8.44, y: 1.3, w: 0.02, h: 1.65, fill: { color: C.accent, transparency: 100 }, line: { color: C.accent, transparency: 100, width: 0.75 } },
      { x: 6.3, y: 3.75, w: 3.2, h: 1.1 },
    ],
  },
  {
    name: "!!tile-3-back-num",
    kind: "text",
    text: "2,1 ngày",
    style: { fontFace: F.head, fontSize: 20, color: C.bg, bold: true },
    states: [null, null, null, { x: 6.5, y: 3.9, w: 1.4, h: 0.8 }],
  },
  {
    name: "!!tile-3-back-label",
    kind: "text",
    text: "thời gian giao trung bình",
    style: { fontFace: F.body, fontSize: 12, color: C.bg },
    states: [null, null, null, { x: 7.95, y: 3.9, w: 1.4, h: 0.8 }],
  },

  // --- Ô biểu đồ: cột mọc từ đáy (thanh tiến trình/biểu đồ chạy)
  {
    name: "!!tile-4",
    kind: "shape",
    shape: "roundRect",
    style: { ...GLASS },
    states: [null, null, { x: 5.1, y: 3.15, w: 4.4, h: 1.7 }, { x: M, y: 1.3, w: 5.6, h: 3.55 }],
  },
  {
    name: "!!tile-4-label",
    kind: "text",
    text: "Doanh thu theo tháng (triệu đồng)",
    style: { fontFace: F.body, fontSize: 12, color: C.muted },
    states: [null, null, { x: 5.3, y: 3.3, w: 4.0, h: 0.3 }, { x: 0.8, y: 1.5, w: 5.0, h: 0.35, fontSize: 14 }],
  },
  ...[
    { v: 412, label: "Tháng 7" },
    { v: 518, label: "Tháng 8" },
    { v: 664, label: "Tháng 9" },
  ].flatMap((d, i) => {
    const hSmall = 0.12 + i * 0.06;
    const hBig = (d.v / 664) * 2.0;
    const isPeak = i === 2;
    return [
      {
        name: `!!chart-bar-${i + 1}`,
        kind: "shape",
        shape: "rect",
        style: { fill: { color: isPeak ? C.accent : C.paper, transparency: isPeak ? 0 : 55 } },
        states: [null, null, bar(5.4 + i * 1.3, hSmall, 4.6), bar(1.0 + i * 1.65, hBig, 4.25)],
      },
      {
        name: `!!chart-val-${i + 1}`,
        kind: "text",
        text: String(d.v),
        style: { fontFace: F.body, fontSize: 14, color: isPeak ? C.accent : C.paper, bold: isPeak, align: "center" },
        states: [null, null, null, { x: 1.0 + i * 1.65, y: 4.25 - hBig - 0.35, w: 1.0, h: 0.3 }],
      },
      {
        name: `!!chart-cat-${i + 1}`,
        kind: "text",
        text: d.label,
        style: { fontFace: F.body, fontSize: 12, color: C.muted, align: "center" },
        states: [null, null, null, { x: 0.85 + i * 1.65, y: 4.32, w: 1.3, h: 0.3 }],
      },
    ];
  }),

  // --- Thanh tiến trình chạy dần qua 4 slide
  {
    name: "!!progress-track",
    kind: "shape",
    shape: "rect",
    style: { fill: { color: C.paper, transparency: 85 } },
    states: Array.from({ length: SLIDE_COUNT }, () => ({ x: M, y: 5.06, w: W - 2 * M, h: 0.04 })),
  },
  {
    name: "!!progress-fill",
    kind: "shape",
    shape: "rect",
    style: { fill: { color: C.accent } },
    states: Array.from({ length: SLIDE_COUNT }, (_, i) => ({ x: M, y: 5.06, w: ((W - 2 * M) * (i + 1)) / SLIDE_COUNT, h: 0.04 })),
  },
];

const NOTES = [
  "Hero: tiêu đề 44pt, thẻ kính nhỏ góc phải. Morph sang slide 2 sẽ phóng to thẻ này.",
  "Big Number Callout: !!hero-card phóng to thành panel, !!hero-kpi lên 100pt. Quả cầu nền phóng to tạo cảm giác camera zoom.",
  "Bento Grid: panel co lại thành ô lớn bên trái, 3 ô mới fade in. Cột biểu đồ đang ở mức thấp.",
  "Spotlight: ô biểu đồ phóng to, cột mọc lên theo số liệu, ô 96% lật sang mặt sau 2,1 ngày.",
];

// ------------------------------------------------------------------
// 3. Dựng deck
// ------------------------------------------------------------------
function buildDeck() {
  const pptx = new pptxgen();
  pptx.layout = "LAYOUT_16x9";
  pptx.title = "Morph demo – Xưởng rang Mộc";
  pptx.theme = { headFontFace: F.head, bodyFontFace: F.body };

  for (let i = 0; i < SLIDE_COUNT; i++) {
    const slide = pptx.addSlide();
    slide.background = { color: C.bg };
    slide.addNotes(NOTES[i]);

    // Thứ tự trong mảng ELEMENTS = thứ tự z (vẽ sau nằm trên)
    for (const el of ELEMENTS) {
      const st = el.states[i];
      if (!st) continue;
      const { text, ...geo } = st;
      const opts = { ...el.style, ...geo, objectName: el.name };

      if (el.kind === "shape") {
        slide.addShape(pptx.ShapeType[el.shape], opts);
      } else {
        slide.addText(text ?? el.text, { margin: 0, valign: "top", fit: "none", ...opts });
      }
    }
  }
  return pptx;
}

// ------------------------------------------------------------------
// 4. Chèn Morph transition vào XML (pptxgenjs chưa hỗ trợ transition)
// ------------------------------------------------------------------
function morphXml(durationMs) {
  return (
    '<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">' +
    '<mc:Choice xmlns:p159="http://schemas.microsoft.com/office/powerpoint/2015/09/main" ' +
    'xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" Requires="p159">' +
    `<p:transition spd="slow" p14:dur="${durationMs}"><p159:morph option="byObject"/></p:transition>` +
    "</mc:Choice>" +
    '<mc:Fallback><p:transition spd="slow"><p:fade/></p:transition></mc:Fallback>' +
    "</mc:AlternateContent>"
  );
}

async function injectMorph(buffer, { durationMs = 1200, fromSlide = 2 } = {}) {
  const zip = await JSZip.loadAsync(buffer);
  const slideFiles = Object.keys(zip.files).filter((f) => /^ppt\/slides\/slide\d+\.xml$/.test(f));
  for (const file of slideFiles) {
    const n = Number(file.match(/slide(\d+)\.xml$/)[1]);
    if (n < fromSlide) continue; // transition nằm ở slide ĐÍCH
    let xml = await zip.file(file).async("string");
    if (xml.includes("p159:morph")) continue;
    xml = xml.replace("</p:clrMapOvr>", `</p:clrMapOvr>${morphXml(durationMs)}`);
    zip.file(file, xml);
  }
  return zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE" });
}

// ------------------------------------------------------------------
// 5. Main
// ------------------------------------------------------------------
async function main() {
  const args = process.argv.slice(2);
  const noMorph = args.includes("--no-morph");
  const out = path.resolve(args.find((a) => !a.startsWith("--")) || "output/morph-demo.pptx");
  fs.mkdirSync(path.dirname(out), { recursive: true });

  const pptx = buildDeck();
  let buf = await pptx.write({ outputType: "nodebuffer" });
  if (!noMorph) buf = await injectMorph(buf, { durationMs: 1200 });
  fs.writeFileSync(out, buf);

  console.log(`Đã tạo ${out}`);
  console.log(noMorph ? "Mở PowerPoint: chọn slide 2–4 -> Transitions -> Morph." : "Morph đã được bật sẵn cho slide 2–4.");
}

if (require.main === module) {
  main().catch((err) => {
    console.error(err);
    process.exit(1);
  });
}

module.exports = { buildDeck, injectMorph, ELEMENTS, C, F };
