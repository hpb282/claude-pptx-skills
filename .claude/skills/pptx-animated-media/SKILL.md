---
name: pptx-animated-media
description: Tự viết script Python (matplotlib/Pillow) hoặc HTML Canvas (Playwright) để render hiệu ứng phức tạp - sóng lan truyền, biểu đồ động chạy số, sơ đồ luồng dữ liệu, hạt chuyển động - thành file GIF động độ phân giải cao lặp liền mạch, rồi nhúng vào slide .pptx bằng pptxgenjs slide.addImage({ path: "animation.gif", x, y, w, h }). Hiệu ứng tự chạy khi trình chiếu PowerPoint, không cần plugin hay macro. Dùng khi người dùng nhắc @gif-pptx, GIF động trong PowerPoint, animation nhúng slide, biểu đồ động, hiệu ứng hạt, sóng, data flow animation, slide có hình chuyển động.
---

# PPTX Animated Media

Morph (`pptx-morph-motion`) chỉ chuyển động **giữa** hai slide. Khi cần chuyển động **liên tục trong một slide** (sóng lan, gói tin chạy, hạt trôi, số đếm lên), hãy render thành GIF. PowerPoint, Keynote và Google Slides đều tự phát GIF động khi trình chiếu, không cần plugin, macro hay bấm nút.

## Khi nào dùng GIF, khi nào không

| Dùng GIF | Không dùng GIF |
|----------|----------------|
| Hiệu ứng nền/minh họa lặp 2–4 giây | Nội dung người xem phải đọc kỹ (GIF làm phân tâm) |
| Biểu đồ "chạy số" một lần khi mở slide | Dữ liệu cần sửa trong PowerPoint (dùng chart native) |
| Sơ đồ luồng có gói tin di chuyển | Video thật, ảnh chụp nhiều màu (GIF chỉ 256 màu → dùng MP4 với `slide.addMedia`) |
| Hạt, sóng, nhịp đập làm điểm nhấn | Clip > 10 giây hoặc > 1280 px rộng (file quá nặng) |

Mỗi slide tối đa **một** GIF động là nhân vật chính. Hai thứ cùng chuyển động sẽ tranh nhau sự chú ý.

## Quy trình

1. **Chọn vị trí và kích thước trên slide trước** (vd. `w: 9, h: 5.0625` inch), rồi suy ra pixel: tỷ lệ phải khớp. `960×540` cho khung 16:9 rộng tới 9 inch; `720×720` cho khung vuông.
2. **Lấy màu từ token của deck** (nền, chữ, màu nhấn). GIF chỉ có trong suốt 1-bit nên viền sẽ bị răng cưa: **vẽ nền GIF trùng màu nền slide** thay vì dùng nền trong suốt.
3. **Viết hàm `render(t)` với `t ∈ [0, 1)`**: mọi chuyển động là hàm tuần hoàn của `t` (sin/cos của `2πt`, `(t + offset) % 1`). Khung cuối nối liền khung đầu → lặp không giật.
4. **Render** bằng `scripts/render_gif.py` (Pillow/matplotlib) hoặc `scripts/canvas_to_gif.js` (HTML Canvas).
5. **Nhúng** bằng `slide.addImage`, `h = w × (cao px / rộng px)`.
6. **Kiểm tra** bằng `pptx-quality-audit` (cảnh báo GIF > 8 MB, ảnh méo) và mở thử chế độ trình chiếu.

## Thông số GIF chuẩn

| Thông số | Giá trị | Lý do |
|----------|---------|-------|
| Khung hình | 25 fps (40 ms/khung) | GIF tính thời gian theo 1/100 giây: 30 fps (33,3 ms) không biểu diễn được; < 20 ms bị nhiều trình phát hạ xuống 100 ms |
| Độ dài vòng | 2–4 giây | Đủ một nhịp, file nhẹ |
| Kích thước | 960×540 (16:9) | Nét khi chiếu, thường < 3 MB |
| Bảng màu | 1 bảng chung cho mọi khung, không dither | Màu phẳng không nhấp nháy, file nhỏ |
| Lặp | `loop=0` (vô hạn) cho hiệu ứng nền; không lặp cho biểu đồ chạy số | Biểu đồ dừng ở khung cuối = trạng thái đọc được |
| Dung lượng | < 3 MB lý tưởng, tối đa 8 MB | Deck nhẹ, mở nhanh, gửi email được |
| Khử răng cưa | Vẽ to 2× rồi thu nhỏ (`SS = 2`) | Pillow không khử răng cưa cho `ellipse`/`line` |

## Script có sẵn

### `scripts/render_gif.py` (Python: Pillow, matplotlib, numpy)

```bash
pip install pillow matplotlib numpy
S=.claude/skills/pptx-animated-media/scripts

python3 $S/render_gif.py ripple    --out media/ripple.gif                       # sóng lan truyền
python3 $S/render_gif.py countup   --out media/revenue.gif --values 412,518,664 --labels "Tháng 7,Tháng 8,Tháng 9"
python3 $S/render_gif.py flow      --out media/flow.gif    --labels "App,API,Hàng đợi,Kho dữ liệu"
python3 $S/render_gif.py particles --out media/particles.gif                    # hạt chuyển động
python3 $S/render_gif.py frames    --out media/canvas.gif  --input frames/      # ghép thư mục PNG

# Tùy chọn chung: --width 960 --height 540 --fps 25 --seconds 3
#                 --bg 16130F --fg F3EADB --accent E2552B --once
```

Đã chạy thử với kích thước mặc định: ripple ≈ 1,2 MB, flow ≈ 0,3 MB, particles ≈ 2,9 MB, countup ≈ 0,1 MB.

| Preset | Hiệu ứng | Kỹ thuật |
|--------|----------|----------|
| `ripple` | 4 vòng tròn nở ra từ tâm, mờ dần, lõi "thở" | `p = (t + k/4) % 1`, bán kính `ease_out_cubic(p)` |
| `countup` | Cột mọc lên so le, nhãn đếm từ 0, cột cao nhất màu nhấn | matplotlib vẽ từng khung, 70% thời gian mọc + 30% đứng yên, không lặp |
| `flow` | Các nút nối nhau, gói tin chạy dọc đường nối, viền nút sáng lệch pha | `ease_in_out(p)` cho vị trí, `sin(πp)` cho độ sáng |
| `particles` | 140 hạt trôi theo quỹ đạo Lissajous, hạt nhấn có quầng sáng | chu kỳ nguyên lần trong vòng lặp + `GaussianBlur` + `ImageChops.add` |
| `frames` | Ghép PNG bất kỳ thành GIF | dùng cho Canvas, Blender, After Effects xuất PNG |

### `scripts/canvas_to_gif.js` (HTML Canvas qua Playwright)

Dùng khi hiệu ứng dễ viết bằng Canvas 2D/WebGL hơn Python, hoặc cần font/hiệu ứng của trình duyệt.

```bash
npm i -D playwright && npx playwright install chromium
node .claude/skills/pptx-animated-media/scripts/canvas_to_gif.js \
  .claude/skills/pptx-animated-media/scripts/canvas-wave.html media/wave.gif --fps 25 --seconds 2
```

Hợp đồng của trang HTML (xem `scripts/canvas-wave.html`):

```html
<canvas id="stage" width="960" height="540"></canvas>
<script>
  const ctx = document.getElementById("stage").getContext("2d");
  // Vẽ khung tại t ∈ [0, 1). KHÔNG dùng requestAnimationFrame / Date.now():
  // script gọi hàm này cho từng khung để kết quả xác định, không rớt khung.
  window.renderFrame = (t) => {
    ctx.fillStyle = "#16130f";
    ctx.fillRect(0, 0, 960, 540);
    // ... vẽ theo t, mọi pha dùng 2 * Math.PI * t để lặp liền mạch
  };
</script>
```

## Viết hiệu ứng mới (khung sườn Python)

Khi người dùng mô tả một hiệu ứng chưa có preset, viết script theo khung này:

```python
#!/usr/bin/env python3
import math
from PIL import Image, ImageDraw

W, H, SS = 960, 540, 2          # SS: vẽ to gấp 2 để khử răng cưa
FPS, SECONDS = 25, 3
BG, FG, ACCENT = (0x16, 0x13, 0x0F), (0xF3, 0xEA, 0xDB), (0xE2, 0x55, 0x2B)

def mix(a, b, t):
    return tuple(round(x + (y - x) * t) for x, y in zip(a, b))

def render(t: float) -> Image.Image:
    """t ∈ [0, 1). Trả về một khung RGB kích thước W×H."""
    im = Image.new("RGB", (W * SS, H * SS), BG)
    d = ImageDraw.Draw(im)
    # Ví dụ: 12 chấm chạy vòng tròn, chấm đầu màu nhấn, các chấm sau mờ dần
    cx, cy, R = W * SS / 2, H * SS / 2, 160 * SS
    for k in range(12):
        a = 2 * math.pi * (t - k / 48)
        x, y = cx + R * math.cos(a), cy + R * math.sin(a)
        r = (14 - k) * SS
        d.ellipse([x - r, y - r, x + r, y + r], fill=mix(BG, ACCENT if k == 0 else FG, 1 - k / 12))
    return im.resize((W, H), Image.LANCZOS)

n = FPS * SECONDS
frames = [render(i / n) for i in range(n)]

# Bảng màu chung + không dither => màu phẳng ổn định
palette = frames[0].quantize(colors=255, method=Image.Quantize.MEDIANCUT)
pal = [f.quantize(palette=palette, dither=Image.Dither.NONE) for f in frames]
pal[0].save("media/orbit.gif", save_all=True, append_images=pal[1:],
            duration=1000 // FPS // 10 * 10, loop=0, optimize=True, disposal=1)
```

Với matplotlib (biểu đồ động), vẽ lại figure cho từng khung và chụp buffer:

```python
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import numpy as np
from PIL import Image

x = np.arange(12)                                       # 12 tháng
y = np.array([30, 33, 36, 41, 52, 66, 70, 74, 81, 88, 97, 112])

fig, ax = plt.subplots(figsize=(9.6, 5.4), dpi=100)   # 960×540 px
frames = []
for i in range(75):
    p = min(i / 52, 1)                                  # 70% thời gian chạy, 30% đứng yên
    ease = 1 - (1 - p) ** 3
    ax.clear()
    ax.set_facecolor("#16130F"); fig.patch.set_facecolor("#16130F")
    ax.plot(x[: max(2, int(len(x) * ease))], y[: max(2, int(len(x) * ease))], color="#E2552B", lw=3)
    ax.set_xlim(x.min(), x.max()); ax.set_ylim(0, y.max() * 1.15)   # cố định trục: không nhảy khung
    ax.axis("off")
    fig.canvas.draw()
    frames.append(Image.fromarray(np.asarray(fig.canvas.buffer_rgba())[..., :3].copy()))
```

Luôn **cố định giới hạn trục** (`set_xlim`, `set_ylim`) và `tight_layout` một cách nhất quán, nếu không biểu đồ sẽ "rung" giữa các khung.

## Nhúng vào slide

```js
const pptxgen = require("pptxgenjs");
const { execFileSync } = require("child_process");

// 1) Render GIF từ Node (hoặc chạy lệnh Python riêng)
execFileSync("python3", [
  ".claude/skills/pptx-animated-media/scripts/render_gif.py", "flow",
  "--out", "media/flow.gif", "--width", "960", "--height", "540",
  "--bg", "16130F", "--fg", "F3EADB", "--accent", "E2552B",
  "--labels", "Ứng dụng,API,Hàng đợi,Kho dữ liệu",
], { stdio: "inherit" });

// 2) Nhúng: giữ đúng tỷ lệ 960×540
const pptx = new pptxgen();
pptx.layout = "LAYOUT_16x9";
const slide = pptx.addSlide();
slide.background = { color: "16130F" };             // trùng nền GIF: không lộ mép ảnh
slide.addText("Mỗi đơn hàng đi qua 4 khối trong 2 giây", {
  x: 0.5, y: 0.5, w: 9, h: 0.6, fontFace: "Georgia", fontSize: 32, color: "F3EADB", margin: 0,
});
const w = 9;
slide.addImage({
  path: "media/flow.gif",
  x: 0.5, y: 1.25, w, h: w * (540 / 960),            // h = w × tỷ lệ px => không méo
  objectName: "!!gif-flow",                          // Morph có thể di chuyển/phóng GIF giữa các slide
  altText: "Sơ đồ động: đơn hàng đi từ ứng dụng qua API, hàng đợi đến kho dữ liệu",
});
await pptx.writeFile({ fileName: "output/flow-demo.pptx" });
```

Lưu ý:

- `h = 9 × 0.5625 = 5.06` vượt lề dưới nếu đặt ở `y: 1.25`. Khi cần cả tiêu đề, giảm `w` (vd. `w: 7.6` → `h: 4.275`) hoặc cho GIF tràn mép có chủ đích với `objectName` chứa `bleed`.
- Có thể đặt GIF làm nền tràn mép (`x: 0, y: 0, w: 10, h: 5.625`, `objectName: "bg-particles"`) rồi phủ chữ lên trên. Khi đó thêm lớp phủ tối `transparency: 40–55` giữa GIF và chữ.
- **Khung đầu tiên là khung "đại diện"**: thumbnail trong PowerPoint, bản in và bản xuất PDF chỉ hiện khung đầu. Với biểu đồ chạy số, khung đầu là cột rỗng. Nếu deck sẽ được xuất PDF, thêm một bản PNG tĩnh của khung cuối cho slide handout, hoặc ghi chú cho người dùng.
- GIF phát ngay khi slide hiện, kể cả trong chế độ soạn thảo. Không cần thiết lập animation hay trigger.

## Hiệu năng và dung lượng

- Giảm dung lượng theo thứ tự: ít khung hơn (2 giây thay vì 4) → kích thước nhỏ hơn (720×405) → ít màu hơn (`colors=128`) → vùng chuyển động nhỏ hơn (chỉ phần động là GIF, phần tĩnh vẽ bằng shape native).
- Nền phẳng nén tốt nhất; gradient lớn và nhiễu hạt nén kém (preset `particles` nặng nhất).
- Tổng dung lượng GIF của một deck nên < 20 MB.

## Checklist

- [ ] Tỷ lệ px của GIF = tỷ lệ `w/h` trên slide
- [ ] Nền GIF trùng màu nền slide (không dùng trong suốt)
- [ ] Lặp liền mạch: khung cuối nối khung đầu (trừ biểu đồ chạy số)
- [ ] 25 fps, 2–4 giây, < 3 MB (tối đa 8 MB)
- [ ] Màu lấy từ token deck, một màu nhấn
- [ ] `altText` mô tả nội dung động
- [ ] Mỗi slide tối đa một GIF làm điểm nhấn
- [ ] Chạy `pptx-quality-audit` và xem thử chế độ trình chiếu
