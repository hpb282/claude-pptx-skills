---
name: pptx-morph-motion
description: Thiết kế chuyển động mượt như video trong PowerPoint bằng Morph Transition (Biến hình) với pptxgenjs. Dựng các cặp slide trạng thái N -> N+1 dùng chung phần tử có objectName tiền tố "!!" (vd "!!hero-card", "!!main-title", "!!chart-bar-1"), đổi x/y/w/h/rotate/transparency để tạo hiệu ứng bay, zoom/pan, thanh tiến trình chạy, lật thẻ, cột biểu đồ mọc. Tự chèn Morph vào XML. Dùng khi người dùng nhắc @morph, Morph, biến hình, slide chuyển động, animation PowerPoint, hiệu ứng mượt, pptx động, smooth transition, magic move.
---

# PPTX Morph Motion

Morph là transition của PowerPoint (365, 2019+, PowerPoint for the web, Mac 2019+). Khi chuyển từ slide N sang N+1, PowerPoint tìm các phần tử "giống nhau" ở hai slide rồi nội suy vị trí, kích thước, góc xoay, màu và độ trong suốt. Kết quả là chuyển động 60fps mà file không chứa một dòng animation nào.

Skill này biến việc đó thành lập trình có kiểm soát: mỗi slide là một **trạng thái** (state), chuyển động là **khoảng cách giữa hai trạng thái**.

Đi kèm: `pptx-visual-design` (bố cục, màu, chữ). Kiểm tra cuối: `pptx-quality-audit`. Code mẫu chạy được: `templates/morph-deck-generator.js`.

## Quy tắc bắt buộc

1. **Đặt tên `!!` cho mọi phần tử cần Morph ghép cặp.** Tên bắt đầu bằng hai dấu chấm than buộc PowerPoint ghép theo tên thay vì tự đoán theo nội dung/vị trí.
   ```js
   slide.addShape(pptx.ShapeType.roundRect, { objectName: "!!hero-card", x: 6.6, y: 3.15, w: 2.9, h: 1.6 });
   slide.addText("+38%", { objectName: "!!hero-kpi", x: 6.85, y: 3.3, w: 2.4, h: 0.7, fontSize: 36 });
   ```
2. **Tên phải duy nhất trong một slide** và **giống hệt nhau** (phân biệt hoa thường) giữa hai slide liền kề. Hai phần tử cùng tên trên một slide làm Morph ghép sai: lỗi `MORPH` trong audit.
3. **Cùng loại đối tượng.** Shape ↔ shape (đổi được dạng hình: `rect` → `roundRect` → `ellipse`), text ↔ text, ảnh ↔ ảnh, chart ↔ chart cùng loại. Không ghép text box với shape.
4. **Transition đặt ở slide ĐÍCH.** Morph từ slide 2 sang 3 = bật Morph cho slide 3. Slide 1 không cần.
5. **Phần tử không có cặp** sẽ tự fade in (xuất hiện ở slide đích) hoặc fade out (chỉ có ở slide nguồn). Đây là cách làm phần tử "hiện ra" mà không cần animation.
6. **Mỗi bước chỉ đổi 1–3 thuộc tính chính.** Đổi cùng lúc vị trí, cỡ, màu, góc của 20 phần tử sẽ thành hỗn loạn. Một ý chính cho mỗi transition.
7. **Thời lượng 0.8–1.5 giây** (`p14:dur="1200"`). Ngắn hơn 0.6 giây nhìn giật, dài hơn 2 giây làm người xem sốt ruột.
8. Đặt `margin: 0` cho text box để tọa độ chữ khớp tọa độ khung, và luôn khai báo `x, y, w, h` bằng inch (slide 16:9 = 10 × 5.625).

## Quy ước đặt tên

| Nhóm | Mẫu tên | Ví dụ |
|------|---------|-------|
| Chữ | `!!<vai-trò>` | `!!main-title`, `!!subtitle`, `!!eyebrow`, `!!section-title` |
| Thẻ/khung | `!!<tên>-card`, `!!tile-<n>` | `!!hero-card`, `!!tile-2` |
| Nội dung trong thẻ | `!!<thẻ>-<phần>` | `!!tile-2-num`, `!!tile-2-label` |
| Biểu đồ dựng bằng shape | `!!chart-bar-<n>`, `!!chart-val-<n>`, `!!chart-cat-<n>` | `!!chart-bar-1` |
| Tiến trình | `!!progress-track`, `!!progress-fill` | |
| Trang trí | `!!deco-<tên>` | `!!deco-orb-a` (audit miễn kiểm tra lề) |

Chữ `deco`, `bg`, `bleed` trong tên báo cho `pptx-quality-audit` biết phần tử được phép tràn mép.

## Mẫu lập trình: khai báo phần tử một lần, liệt kê trạng thái theo slide

Viết từng slide bằng tay rất dễ lệch tên hoặc quên một phần tử. Hãy mô tả **mỗi phần tử một lần** với mảng `states`, phần tử thứ `i` là trạng thái ở slide `i`, `null` là vắng mặt.

```js
const pptxgen = require("pptxgenjs");

const C = { bg: "16130F", paper: "F3EADB", muted: "A89F90", accent: "E2552B" };
const GLASS = {
  fill: { color: "FFFFFF", transparency: 85 },
  line: { color: "FFFFFF", transparency: 70, width: 0.75 },
  rectRadius: 0.16,
};

const ELEMENTS = [
  {
    name: "!!hero-card", kind: "shape", shape: "roundRect", style: GLASS,
    states: [
      { x: 6.6, y: 3.15, w: 2.9, h: 1.6 },   // slide 1: thẻ nhỏ góc phải
      { x: 0.5, y: 1.2,  w: 9.0, h: 3.65 },  // slide 2: phóng to thành panel (zoom)
      { x: 0.5, y: 1.3,  w: 4.4, h: 3.55 },  // slide 3: co lại thành ô bento
    ],
  },
  {
    name: "!!hero-kpi", kind: "text", text: "+38%",
    style: { fontFace: "Georgia", fontSize: 36, bold: true, color: C.accent },
    states: [
      { x: 6.85, y: 3.3, w: 2.4, h: 0.7 },
      { x: 0.9,  y: 1.5, w: 5.2, h: 1.6, fontSize: 100 }, // trạng thái được ghi đè style
      { x: 0.8,  y: 1.6, w: 3.8, h: 1.0, fontSize: 60 },
    ],
  },
];

function buildDeck(slideCount) {
  const pptx = new pptxgen();
  pptx.layout = "LAYOUT_16x9";
  for (let i = 0; i < slideCount; i++) {
    const slide = pptx.addSlide();
    slide.background = { color: C.bg };
    for (const el of ELEMENTS) {            // thứ tự mảng = thứ tự lớp (z-order)
      const st = el.states[i];
      if (!st) continue;
      const { text, ...geo } = st;
      const opts = { ...el.style, ...geo, objectName: el.name };
      if (el.kind === "shape") slide.addShape(pptx.ShapeType[el.shape], opts);
      else slide.addText(text ?? el.text, { margin: 0, valign: "top", ...opts });
    }
  }
  return pptx;
}
```

Lợi ích: thứ tự lớp giống nhau ở mọi slide (Morph không bị đảo lớp giữa chừng), tên không thể lệch, và đọc mảng `states` là thấy ngay kịch bản chuyển động.

## Thư viện kỹ thuật chuyển động

Mỗi kỹ thuật dưới đây là một cặp trạng thái. Ghép chúng lại thành câu chuyện.

### 1. Bay vào / bay ra (fly-in)

Đặt trạng thái nguồn **ngoài khung slide**, trạng thái đích trong khung. Morph tạo chuyển động trượt vào.

```js
{ name: "!!quote", kind: "text", text: "Khách quay lại 3 lần/tháng",
  states: [
    { x: 10.4, y: 2.2, w: 6, h: 0.8 },  // slide 1: nằm ngoài mép phải (x > 10)
    { x: 2.0,  y: 2.2, w: 6, h: 0.8 },  // slide 2: trượt vào giữa
    { x: 2.0,  y: -1.2, w: 6, h: 0.8 }, // slide 3: bay lên thoát khỏi khung
  ] }
```

Phần tử ngoài khung không hiện khi trình chiếu, nhưng phải có tên `!!` (audit chỉ báo `INFO`, không lỗi).

### 2. Zoom / Pan (camera)

Morph không có camera, nhưng phóng to **toàn bộ cảnh** quanh một điểm cho cảm giác camera tiến vào. Dùng hàm biến đổi tọa độ:

```js
// Phóng to cả cụm phần tử quanh điểm (cx, cy) với hệ số k, rồi đưa điểm đó về (tx, ty).
function camera(state, { k = 2, cx = 5, cy = 2.8125, tx = 5, ty = 2.8125 } = {}) {
  return {
    ...state,
    x: tx + (state.x - cx) * k,
    y: ty + (state.y - cy) * k,
    w: state.w * k,
    h: state.h * k,
    ...(state.fontSize ? { fontSize: Math.round(state.fontSize * k) } : {}),
  };
}

const mapNodes = [
  { name: "!!node-hcm", x: 6.2, y: 3.4, w: 0.5, h: 0.5 },
  { name: "!!node-hn",  x: 4.1, y: 0.9, w: 0.5, h: 0.5 },
  { name: "!!node-dn",  x: 5.3, y: 2.1, w: 0.5, h: 0.5 },
];
// Slide 1: toàn cảnh. Slide 2: zoom x2.5 vào TP.HCM và kéo nó về giữa slide.
const slide1 = mapNodes;
const slide2 = mapNodes.map((n) => camera(n, { k: 2.5, cx: 6.45, cy: 3.65, tx: 5, ty: 2.8 }));
```

Pan (lia ngang) = cùng hệ số `k`, chỉ đổi `tx`/`ty`. Phần tử bị đẩy ra ngoài khung sau khi zoom là bình thường.

### 3. Thanh tiến trình chạy

Track cố định, fill tăng `w` theo số slide. Đặt cùng tên trên **mọi** slide.

```js
const N = 6; // tổng số slide trong chương
const track = { x: 0.5, y: 5.06, w: 9, h: 0.04 };
for (let i = 0; i < N; i++) {
  const s = pptx.addSlide();
  s.addShape(pptx.ShapeType.rect, { ...track, objectName: "!!progress-track",
    fill: { color: "F3EADB", transparency: 85 } });
  s.addShape(pptx.ShapeType.rect, { ...track, w: (9 * (i + 1)) / N, objectName: "!!progress-fill",
    fill: { color: "E2552B" } });
}
```

Biến thể: chấm tròn `!!progress-dot` trượt dọc `x`, hoặc vòng tiến trình bằng `pptx.ShapeType.blockArc` với `angleRange: [270, 270 + 360 * p]` (Morph nội suy được cung tròn).

### 4. Cột biểu đồ mọc lên

Biểu đồ native (`addChart`) Morph được nhưng chỉ khi hai chart cùng loại và khó kiểm soát. Muốn cột "mọc" chắc chắn, dựng cột bằng shape: giữ **đáy cố định**, đổi `y` và `h` cùng lúc.

```js
const base = 4.25;            // đường đáy (inch)
const maxH = 2.0;
const data = [412, 518, 664];
const peak = Math.max(...data);

data.forEach((v, i) => {
  const x = 1.0 + i * 1.65;
  const hSmall = 0.12;                // slide A: gần như phẳng
  const hBig = (v / peak) * maxH;     // slide B: đúng tỷ lệ số liệu
  slideA.addShape(pptx.ShapeType.rect, { objectName: `!!chart-bar-${i + 1}`,
    x, y: base - hSmall, w: 1.0, h: hSmall, fill: { color: "F3EADB", transparency: 55 } });
  slideB.addShape(pptx.ShapeType.rect, { objectName: `!!chart-bar-${i + 1}`,
    x, y: base - hBig, w: 1.0, h: hBig,
    fill: { color: v === peak ? "E2552B" : "F3EADB", transparency: v === peak ? 0 : 55 } });
});
```

Màu cũng được Morph nội suy, nên cột cao nhất "nóng" dần lên trong lúc mọc.

### 5. Lật thẻ (card flip)

Morph chỉ làm 2D, nên flip được giả lập bằng **bóp bề ngang về 0 tại tâm**:

- Mặt trước `!!card-front`: slide A rộng đầy đủ → slide B `w: 0.02`, `x` = tâm thẻ, `transparency: 100`.
- Mặt sau `!!card-back`: slide A `w: 0.02` tại tâm, `transparency: 100` → slide B rộng đầy đủ.

```js
const card = { x: 6.3, y: 3.75, w: 3.2, h: 1.1 };
const cx = card.x + card.w / 2;
const squeezed = { x: cx - 0.01, y: card.y, w: 0.02, h: card.h };
const hidden = { fill: { color: "FFFFFF", transparency: 100 }, line: { color: "FFFFFF", transparency: 100, width: 0.75 } };

// Slide A
slideA.addShape(pptx.ShapeType.roundRect, { objectName: "!!card-front", ...card, ...GLASS });
slideA.addShape(pptx.ShapeType.roundRect, { objectName: "!!card-back", ...squeezed, ...hidden, rectRadius: 0.16 });
// Slide B
slideB.addShape(pptx.ShapeType.roundRect, { objectName: "!!card-front", ...squeezed, ...hidden, rectRadius: 0.16 });
slideB.addShape(pptx.ShapeType.roundRect, { objectName: "!!card-back", ...card,
  fill: { color: "E2552B" }, line: { color: "E2552B", width: 0.75 }, rectRadius: 0.16 });
```

Muốn flip rõ hai pha (đóng hết rồi mới mở) thì thêm một slide trung gian nơi cả hai mặt đều bóp, với transition 0.5 giây mỗi bước. Chữ trên thẻ xử lý giống hệt: chữ mặt trước bóp + `transparency: 100`, chữ mặt sau mới xuất hiện ở slide B.

### 6. Xoay (rotate)

```js
{ name: "!!badge", kind: "shape", shape: "roundRect",
  states: [ { x: 7, y: 1, w: 1.6, h: 1.6, rotate: -12 }, { x: 7, y: 1, w: 1.6, h: 1.6, rotate: 168 } ] }
```

Giữ mỗi bước xoay ≤ 180° để hướng quay dễ đoán. Muốn quay nhiều vòng thì chia thành nhiều slide liên tiếp, mỗi slide cộng thêm tối đa 180°.

### 7. Hiện/mờ có kiểm soát (transparency)

Fade mặc định của phần tử không có cặp không chỉnh được tốc độ. Muốn chủ động, giữ phần tử ở cả hai slide và đổi `transparency`:

```js
// Shape: fill.transparency. Text: transparency của màu chữ.
slideA.addText("Bước 2", { objectName: "!!step-2", x: 4, y: 2, w: 3, h: 0.5, color: "F3EADB", transparency: 100 });
slideB.addText("Bước 2", { objectName: "!!step-2", x: 4, y: 2, w: 3, h: 0.5, color: "F3EADB", transparency: 0 });
```

Đây cũng là cách làm "tô sáng một phần, làm mờ phần còn lại" của `pptx-charts-diagrams`.

## Bật Morph tự động trong file

pptxgenjs chưa có API transition. Chèn XML sau khi tạo file (JSZip là dependency có sẵn của pptxgenjs):

```js
const JSZip = require("jszip");

function morphXml(durationMs = 1200) {
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
  for (const file of Object.keys(zip.files)) {
    const m = file.match(/^ppt\/slides\/slide(\d+)\.xml$/);
    if (!m || Number(m[1]) < fromSlide) continue;
    let xml = await zip.file(file).async("string");
    if (xml.includes("p159:morph")) continue;
    // p:transition phải đứng ngay sau p:clrMapOvr (trước p:timing) theo schema
    xml = xml.replace("</p:clrMapOvr>", `</p:clrMapOvr>${morphXml(durationMs)}`);
    zip.file(file, xml);
  }
  return zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE" });
}

// Dùng:
const buf = await pptx.write({ outputType: "nodebuffer" });
fs.writeFileSync("deck.pptx", await injectMorph(buf));
```

- `option="byObject"`: ghép theo đối tượng (mặc định, dùng với `!!`). `"byWord"` / `"byChar"`: Morph từng chữ khi nội dung chữ thay đổi.
- `mc:Fallback` với `p:fade` giúp Keynote, LibreOffice, PowerPoint 2016 trở xuống vẫn mở được và có transition mờ dần.
- Nếu chỉ muốn Morph ở một vài slide, lọc theo số slide trong vòng lặp.
- Bật tay vẫn được: chọn slide đích → **Transitions → Morph** → *Effect Options → Objects*.

## Quy trình làm một chuỗi Morph

1. **Viết kịch bản trước khi code.** Mỗi slide một dòng: "Slide 2: thẻ KPI phóng to chiếm khung, số lớn 100pt". Nếu không mô tả được thay đổi bằng một câu, chuyển động đó thừa.
2. **Liệt kê phần tử xuyên suốt** (ai đi qua nhiều slide) và đặt tên `!!` cho chúng.
3. **Dựng trạng thái cuối trước** (slide đẹp nhất), rồi suy ngược các trạng thái trước nó. Mỗi slide phải đẹp khi đứng yên, vì người xem có thể dừng ở bất kỳ slide nào và bản PDF chỉ có trạng thái tĩnh.
4. **Giữ thứ tự lớp ổn định**: phần tử nền thêm trước, chữ thêm sau, ở mọi slide.
5. **Chạy `pptx-quality-audit`**: kiểm tra trùng tên, tràn chữ, lề, tương phản ở từng trạng thái.
6. **Mở bằng PowerPoint thật** để xem Morph (LibreOffice không chạy Morph).

## Lỗi thường gặp

| Triệu chứng | Nguyên nhân | Cách sửa |
|-------------|-------------|----------|
| Phần tử mờ đi rồi hiện lại thay vì di chuyển | Tên khác nhau giữa hai slide, hoặc thiếu `!!` | So tên trong Selection Pane (Alt+F10), dùng mảng `states` |
| Hai thẻ đổi chỗ cho nhau kỳ lạ | Tên trùng trong cùng slide | Audit `MORPH` ERROR, đặt tên duy nhất |
| Chữ bị méo lúc chuyển | Text box đổi tỷ lệ w/h quá mạnh cùng lúc đổi cỡ chữ | Đổi `fontSize` và giữ tỷ lệ khung gần nhau |
| Phần tử nhảy lớp giữa chừng | Thứ tự `add*` khác nhau giữa các slide | Một mảng `ELEMENTS` chung cho mọi slide |
| Không có chuyển động khi mở bằng Keynote/Google Slides | Các ứng dụng này không hỗ trợ Morph | Giữ `mc:Fallback` fade, báo người dùng mở bằng PowerPoint |
| Chuyển động quá rối | Quá nhiều thuộc tính đổi cùng lúc | Mỗi transition chỉ một ý; phần tử phụ đứng yên |

## Tiết chế

- Không phải slide nào cũng cần Morph. Dùng cho: mở đầu, chuyển chương, tiết lộ con số, giải thích quy trình. Slide nhiều chữ đọc kỹ thì để yên.
- Trang trí (`!!deco-*`) chuyển động chậm và ít; nội dung chính mới là thứ di chuyển nhiều.
- Giữ 1 chuyển động "anh hùng" mỗi transition; phần còn lại là chuyển động phụ nhỏ (thanh tiến trình, quả cầu nền trôi).
