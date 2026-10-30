#!/usr/bin/env python3
"""
render_gif.py — render hiệu ứng động thành GIF lặp liền mạch để nhúng vào .pptx.

    pip install pillow matplotlib numpy

    python3 render_gif.py ripple    --out media/ripple.gif
    python3 render_gif.py countup   --out media/revenue.gif --values 412,518,664 --labels "T7,T8,T9"
    python3 render_gif.py flow      --out media/flow.gif    --labels "App,API,Kafka,Kho dữ liệu"
    python3 render_gif.py particles --out media/particles.gif
    python3 render_gif.py frames    --out media/canvas.gif  --input frames/ --fps 25

Tùy chọn chung:
    --width 960 --height 540   kích thước pixel (giữ đúng tỷ lệ khung đặt trên slide)
    --fps 25                   25 fps = 40 ms/khung (GIF tính theo 1/100 giây)
    --seconds 3                độ dài một vòng lặp
    --bg 16130F --fg F3EADB --accent E2552B   màu trùng với theme deck
    --once                     chạy một lần rồi dừng ở khung cuối (mặc định lặp vô hạn)

Mỗi preset trả về danh sách khung RGB; save_gif() lượng tử hóa về 1 bảng màu chung
(không dither) để màu phẳng không nhấp nháy và file nhẹ.
"""
from __future__ import annotations

import argparse
import math
import random
from pathlib import Path

from PIL import Image, ImageChops, ImageDraw, ImageFilter, ImageFont

SS = 2  # supersampling: vẽ to gấp 2 rồi thu nhỏ để khử răng cưa


def hex_rgb(h: str) -> tuple[int, int, int]:
    h = h.strip("#")
    return tuple(int(h[i : i + 2], 16) for i in (0, 2, 4))


def mix(a, b, t: float):
    return tuple(round(x + (y - x) * t) for x, y in zip(a, b))


def ease_out_cubic(t: float) -> float:
    return 1 - (1 - t) ** 3


def ease_in_out(t: float) -> float:
    return 0.5 - 0.5 * math.cos(math.pi * t)


def load_font(size: int, bold: bool = False):
    candidates = [
        "DejaVuSans-Bold.ttf" if bold else "DejaVuSans.ttf",
        "Arial Bold.ttf" if bold else "Arial.ttf",
        "arialbd.ttf" if bold else "arial.ttf",
        "/System/Library/Fonts/Supplemental/Arial.ttf",
        "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
    ]
    for c in candidates:
        try:
            return ImageFont.truetype(c, size)
        except OSError:
            continue
    return ImageFont.load_default()


# ----------------------------------------------------------------------------
# Lưu GIF
# ----------------------------------------------------------------------------
def save_gif(frames: list[Image.Image], out: str, fps: int, loop: bool = True, hold_last_ms: int = 0):
    out_path = Path(out)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    # Bảng màu chung lấy từ vài khung đại diện -> không nhấp nháy giữa các khung
    sample = Image.new("RGB", (frames[0].width, frames[0].height * 3))
    for i, idx in enumerate((0, len(frames) // 2, len(frames) - 1)):
        sample.paste(frames[idx], (0, frames[0].height * i))
    palette = sample.quantize(colors=255, method=Image.Quantize.MEDIANCUT)
    pal_frames = [f.quantize(palette=palette, dither=Image.Dither.NONE) for f in frames]

    frame_ms = max(20, round(1000 / fps / 10) * 10)  # bội số 10 ms, tối thiểu 20 ms
    durations = [frame_ms] * len(pal_frames)
    if hold_last_ms:
        durations[-1] = hold_last_ms
    kwargs = dict(save_all=True, append_images=pal_frames[1:], duration=durations, optimize=True, disposal=1)
    if loop:
        kwargs["loop"] = 0  # 0 = lặp vô hạn
    pal_frames[0].save(out_path, **kwargs)
    kb = out_path.stat().st_size / 1024
    print(f"{out_path}  {frames[0].width}x{frames[0].height}px  {len(frames)} khung  {frame_ms} ms/khung  {kb:.0f} KB")
    print(f"Nhúng vào slide: w = <inch>, h = w * {frames[0].height / frames[0].width:.4f}")


# ----------------------------------------------------------------------------
# Preset 1: sóng lan truyền (ripple) — vòng tròn nở ra và mờ dần, lặp liền mạch
# ----------------------------------------------------------------------------
def ripple(a) -> list[Image.Image]:
    W, H = a.width * SS, a.height * SS
    bg, fg, ac = hex_rgb(a.bg), hex_rgb(a.fg), hex_rgb(a.accent)
    n = a.fps * a.seconds
    rings = 4
    max_r = math.hypot(W, H) / 2
    frames = []
    for i in range(n):
        t = i / n  # [0, 1): khung cuối nối liền khung đầu
        im = Image.new("RGB", (W, H), bg)
        d = ImageDraw.Draw(im)
        for k in range(rings):
            p = (t + k / rings) % 1.0
            r = 30 * SS + ease_out_cubic(p) * max_r
            fade = (1 - p) ** 1.6
            col = mix(bg, ac if k % 2 == 0 else fg, fade * 0.9)
            w = max(1, round((6 - 4 * p) * SS))
            d.ellipse([W / 2 - r, H / 2 - r, W / 2 + r, H / 2 + r], outline=col, width=w)
        core = 22 * SS * (1 + 0.08 * math.sin(2 * math.pi * t))
        d.ellipse([W / 2 - core, H / 2 - core, W / 2 + core, H / 2 + core], fill=ac)
        frames.append(im.resize((a.width, a.height), Image.LANCZOS))
    return frames


# ----------------------------------------------------------------------------
# Preset 2: biểu đồ cột chạy số (matplotlib) — cột mọc lên, nhãn đếm từ 0
# ----------------------------------------------------------------------------
def countup(a) -> list[Image.Image]:
    import matplotlib

    matplotlib.use("Agg")
    import matplotlib.pyplot as plt
    import numpy as np

    values = [float(v) for v in a.values.split(",")]
    labels = a.labels.split(",") if a.labels else [str(i + 1) for i in range(len(values))]
    bg, fg, ac = "#" + a.bg, "#" + a.fg, "#" + a.accent
    muted = "#%02x%02x%02x" % mix(hex_rgb(a.fg), hex_rgb(a.bg), 0.55)
    peak = int(np.argmax(values))
    grow = int(a.fps * a.seconds * 0.7)  # 70% thời gian để mọc, 30% đứng yên
    n = int(a.fps * a.seconds)
    dpi = 100
    fig, ax = plt.subplots(figsize=(a.width / dpi, a.height / dpi), dpi=dpi)
    fig.patch.set_facecolor(bg)
    frames = []
    for i in range(n):
        ax.clear()
        ax.set_facecolor(bg)
        for s in ax.spines.values():
            s.set_visible(False)
        ax.tick_params(colors=muted, length=0, labelsize=14)
        ax.set_yticks([])
        ax.set_ylim(0, max(values) * 1.18)
        stagger = 0.12  # mỗi cột bắt đầu trễ hơn cột trước 12% thời gian mọc
        span = 1 - stagger * (len(values) - 1)
        for k, v in enumerate(values):
            p = min(max((i / grow - k * stagger) / span, 0), 1)
            cur = v * ease_out_cubic(p)
            color = ac if k == peak else fg
            ax.bar(k, cur, width=0.62, color=color, alpha=1 if k == peak else 0.45)
            ax.text(k, cur + max(values) * 0.03, f"{cur:,.0f}".replace(",", "."), ha="center", va="bottom",
                    color=ac if k == peak else fg, fontsize=16, fontweight="bold" if k == peak else "normal")
        ax.set_xticks(range(len(values)), labels)
        fig.tight_layout(pad=0.6)
        fig.canvas.draw()
        buf = np.asarray(fig.canvas.buffer_rgba())[..., :3]
        frames.append(Image.fromarray(buf.copy()))
    plt.close(fig)
    return frames


# ----------------------------------------------------------------------------
# Preset 3: sơ đồ luồng dữ liệu — các nút nối nhau, gói tin chạy dọc đường nối
# ----------------------------------------------------------------------------
def flow(a) -> list[Image.Image]:
    W, H = a.width * SS, a.height * SS
    bg, fg, ac = hex_rgb(a.bg), hex_rgb(a.fg), hex_rgb(a.accent)
    line_c = mix(bg, fg, 0.25)
    labels = (a.labels or "Nguồn,Xử lý,Lưu trữ,Báo cáo").split(",")
    n_nodes = len(labels)
    font = load_font(int(H * 0.045), bold=True)
    pad_x = W * 0.08
    node_w, node_h = (W - 2 * pad_x) / (n_nodes * 1.6), H * 0.2
    gap = (W - 2 * pad_x - node_w * n_nodes) / max(n_nodes - 1, 1)
    xs = [pad_x + k * (node_w + gap) for k in range(n_nodes)]
    cy = H / 2
    n = a.fps * a.seconds
    packets = 3  # số gói tin trên mỗi đoạn
    frames = []
    for i in range(n):
        t = i / n
        im = Image.new("RGB", (W, H), bg)
        d = ImageDraw.Draw(im)
        for k in range(n_nodes - 1):
            x0, x1 = xs[k] + node_w, xs[k + 1]
            d.line([(x0, cy), (x1, cy)], fill=line_c, width=3 * SS)
            for j in range(packets):
                p = (t + j / packets + k * 0.17) % 1.0
                px = x0 + (x1 - x0) * ease_in_out(p)
                alpha = math.sin(math.pi * p)  # hiện dần rồi tắt dần ở 2 đầu
                r = 6 * SS
                d.ellipse([px - r, cy - r, px + r, cy + r], fill=mix(bg, ac, alpha))
        for k, x in enumerate(xs):
            # nút "thở" lệch pha theo thứ tự => cảm giác dữ liệu đi qua từng nút
            glow = 0.5 + 0.5 * math.sin(2 * math.pi * (t - k / n_nodes))
            border = mix(line_c, ac, glow * 0.8)
            d.rounded_rectangle([x, cy - node_h / 2, x + node_w, cy + node_h / 2], radius=14 * SS,
                                fill=mix(bg, fg, 0.08), outline=border, width=3 * SS)
            tw = d.textlength(labels[k], font=font)
            d.text((x + node_w / 2 - tw / 2, cy - font.size * 0.6), labels[k], fill=fg, font=font)
        frames.append(im.resize((a.width, a.height), Image.LANCZOS))
    return frames


# ----------------------------------------------------------------------------
# Preset 4: hạt chuyển động — trường hạt trôi theo quỹ đạo tuần hoàn (lặp liền mạch)
# ----------------------------------------------------------------------------
def particles(a) -> list[Image.Image]:
    W, H = a.width * SS, a.height * SS
    bg, fg, ac = hex_rgb(a.bg), hex_rgb(a.fg), hex_rgb(a.accent)
    rnd = random.Random(7)
    count = 140
    pts = [
        dict(
            x=rnd.random() * W, y=rnd.random() * H,
            ax=rnd.uniform(20, 90) * SS, ay=rnd.uniform(10, 50) * SS,
            ph=rnd.random() * 2 * math.pi, k=rnd.choice((1, 1, 2)),
            r=rnd.uniform(1.5, 4.5) * SS, c=ac if rnd.random() < 0.18 else fg,
        )
        for _ in range(count)
    ]
    n = a.fps * a.seconds
    frames = []
    for i in range(n):
        t = 2 * math.pi * i / n  # mọi chuyển động có chu kỳ chia hết cho vòng lặp
        im = Image.new("RGB", (W, H), bg)
        glow = Image.new("RGB", (W, H), (0, 0, 0))
        d, g = ImageDraw.Draw(im), ImageDraw.Draw(glow)
        for p in pts:
            x = p["x"] + p["ax"] * math.sin(p["k"] * t + p["ph"])
            y = p["y"] + p["ay"] * math.sin(2 * p["k"] * t + p["ph"])
            tw = 0.55 + 0.45 * math.sin(p["k"] * t + p["ph"] * 2)
            r = p["r"]
            d.ellipse([x - r, y - r, x + r, y + r], fill=mix(bg, p["c"], tw))
            if p["c"] == ac:
                g.ellipse([x - r * 4, y - r * 4, x + r * 4, y + r * 4], fill=mix((0, 0, 0), ac, 0.5 * tw))
        glow = glow.filter(ImageFilter.GaussianBlur(8 * SS))
        im = ImageChops.add(im, glow)  # cộng sáng: hạt nhấn màu phát quầng
        frames.append(im.resize((a.width, a.height), Image.LANCZOS))
    return frames


def frames_dir(a) -> list[Image.Image]:
    """Ghép thư mục PNG (vd. do canvas_to_gif.js chụp) thành GIF."""
    files = sorted(Path(a.input).glob("*.png"))
    if not files:
        raise SystemExit(f"Không có file PNG trong {a.input}")
    return [Image.open(f).convert("RGB") for f in files]


PRESETS = {"ripple": ripple, "countup": countup, "flow": flow, "particles": particles, "frames": frames_dir}


def main():
    ap = argparse.ArgumentParser(description="Render GIF động cho slide PowerPoint")
    ap.add_argument("preset", choices=PRESETS)
    ap.add_argument("--out", required=True)
    ap.add_argument("--width", type=int, default=960)
    ap.add_argument("--height", type=int, default=540)
    ap.add_argument("--fps", type=int, default=25)
    ap.add_argument("--seconds", type=int, default=3)
    ap.add_argument("--bg", default="16130F")
    ap.add_argument("--fg", default="F3EADB")
    ap.add_argument("--accent", default="E2552B")
    ap.add_argument("--values", default="412,518,664")
    ap.add_argument("--labels", default="")
    ap.add_argument("--input", default="", help="thư mục PNG cho preset frames")
    ap.add_argument("--once", action="store_true", help="chạy một lần, dừng ở khung cuối")
    a = ap.parse_args()

    frames = PRESETS[a.preset](a)
    # countup đẹp nhất khi chạy 1 lần rồi đứng yên; các preset khác lặp liền mạch
    once = a.once or a.preset == "countup"
    save_gif(frames, a.out, a.fps, loop=not once, hold_last_ms=0)


if __name__ == "__main__":
    main()
