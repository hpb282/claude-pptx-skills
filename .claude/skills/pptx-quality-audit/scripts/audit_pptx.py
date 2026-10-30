#!/usr/bin/env python3
"""
audit_pptx.py — kiểm tra chất lượng file .pptx trước khi bàn giao.

Không cần thư viện ngoài (Pillow là tùy chọn, dùng để soi méo ảnh).

    python3 audit_pptx.py deck.pptx              # báo cáo dạng chữ
    python3 audit_pptx.py deck.pptx --json       # báo cáo JSON
    python3 audit_pptx.py deck.pptx --margin 0.5 --max-lines 6
    python3 audit_pptx.py deck.pptx --morph-map  # xem phần tử "!!" nào ghép cặp

Các nhóm lỗi:
    SIZE      tỷ lệ slide không phải 16:9
    BOUNDS    phần tử vượt ra ngoài slide
    MARGIN    phần tử lấn vào lề an toàn (mặc định 0.5 inch)
    OVERFLOW  ước tính chữ tràn khỏi khung
    OVERLAP   hai khung chữ đè lên nhau
    LINES     quá nhiều dòng chữ nội dung trên một slide
    FONT      cỡ chữ quá nhỏ
    CONTRAST  độ tương phản chữ/nền dưới chuẩn WCAG
    MORPH     tên "!!" trùng lặp, hoặc có cặp "!!" nhưng chưa bật Morph
    MEDIA     GIF quá nặng, ảnh bị méo tỷ lệ

Mã thoát: 0 nếu không có ERROR, 1 nếu có.

Quy ước tên được miễn kiểm tra lề/biên: tên chứa "deco", "bg" hoặc "bleed"
(ví dụ "!!deco-orb-a", "bg-photo", "full-bleed-hero").
"""
from __future__ import annotations

import argparse
import io
import json
import math
import posixpath
import re
import sys
import zipfile
import xml.etree.ElementTree as ET
from dataclasses import dataclass, field

EMU = 914400
NS = {
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
}
EXEMPT = re.compile(r"deco|(^|[-_!])bg([-_]|$)|bleed", re.I)
DEFAULT_SIZE = 18.0


# ----------------------------------------------------------------------------
# Mô hình dữ liệu
# ----------------------------------------------------------------------------
@dataclass
class Para:
    text: str
    size: float
    bold: bool
    color: str | None
    alpha: float


@dataclass
class Item:
    name: str
    kind: str  # sp | pic | graphicFrame | cxnSp
    x: float
    y: float
    w: float
    h: float
    rot: float = 0.0
    fill: str | None = None
    fill_alpha: float = 1.0
    paras: list[Para] = field(default_factory=list)
    insets: tuple[float, float, float, float] = (0.1, 0.05, 0.1, 0.05)  # l t r b
    autofit: bool = False
    image: str | None = None
    has_crop: bool = False
    geom: str = "rect"

    def contains(self, px: float, py: float) -> bool:
        if not (self.x <= px <= self.x + self.w and self.y <= py <= self.y + self.h):
            return False
        if self.geom == "ellipse":
            rx, ry = self.w / 2 or 1e-9, self.h / 2 or 1e-9
            return ((px - self.x - rx) / rx) ** 2 + ((py - self.y - ry) / ry) ** 2 <= 1
        return True

    @property
    def has_text(self) -> bool:
        return any(p.text.strip() for p in self.paras)

    @property
    def exempt(self) -> bool:
        return bool(EXEMPT.search(self.name))


@dataclass
class Finding:
    severity: str
    code: str
    slide: int
    name: str
    message: str


# ----------------------------------------------------------------------------
# Đọc XML
# ----------------------------------------------------------------------------
def q(tag: str) -> str:
    pre, local = tag.split(":")
    return f"{{{NS[pre]}}}{local}"


def read_rels(z: zipfile.ZipFile, part: str) -> dict[str, str]:
    base = posixpath.dirname(part)
    rels_path = posixpath.join(base, "_rels", posixpath.basename(part) + ".rels")
    if rels_path not in z.namelist():
        return {}
    root = ET.fromstring(z.read(rels_path))
    out = {}
    for r in root.findall("rel:Relationship", NS):
        target = r.get("Target", "")
        if r.get("TargetMode") == "External":
            continue
        out[r.get("Id")] = posixpath.normpath(posixpath.join(base, target))
    return out


def solid_color(el) -> tuple[str | None, float]:
    """Trả về (hex, alpha) của <a:solidFill> bên trong el, hoặc (None, 1)."""
    if el is None:
        return None, 1.0
    sf = el.find("a:solidFill", NS)
    if sf is None:
        return None, 1.0
    c = sf.find("a:srgbClr", NS)
    if c is None:
        c = sf.find("a:sysClr", NS)
        if c is not None:
            return c.get("lastClr"), 1.0
        return None, 1.0  # schemeClr: không giải được, bỏ qua
    alpha_el = c.find("a:alpha", NS)
    alpha = int(alpha_el.get("val")) / 100000 if alpha_el is not None else 1.0
    return c.get("val") or c.get("lastClr"), alpha


def parse_xfrm(xfrm, parent=None):
    if xfrm is None:
        return None
    off, ext = xfrm.find("a:off", NS), xfrm.find("a:ext", NS)
    if off is None or ext is None:
        return None
    x, y = int(off.get("x")) / EMU, int(off.get("y")) / EMU
    w, h = int(ext.get("cx")) / EMU, int(ext.get("cy")) / EMU
    if parent:
        x, y, w, h = parent(x, y, w, h)
    return x, y, w, h, int(xfrm.get("rot", "0")) / 60000


def parse_text(tx) -> tuple[list[Para], tuple, bool]:
    paras: list[Para] = []
    insets = (0.1, 0.05, 0.1, 0.05)
    autofit = False
    if tx is None:
        return paras, insets, autofit
    bp = tx.find("a:bodyPr", NS)
    if bp is not None:
        insets = tuple(
            int(bp.get(k, d)) / EMU
            for k, d in (("lIns", 91440), ("tIns", 45720), ("rIns", 91440), ("bIns", 45720))
        )
        autofit = bp.find("a:normAutofit", NS) is not None or bp.find("a:spAutoFit", NS) is not None
    for p in tx.findall("a:p", NS):
        # Tách đoạn theo <a:br/> để đếm đúng số dòng cứng
        chunks, cur = [], ""
        size, bold, color, alpha = None, False, None, 1.0
        for child in p:
            if child.tag == q("a:r") or child.tag == q("a:fld"):
                rpr = child.find("a:rPr", NS)
                t = child.find("a:t", NS)
                cur += t.text or "" if t is not None else ""
                if rpr is not None:
                    if rpr.get("sz"):
                        size = max(size or 0, int(rpr.get("sz")) / 100)
                    bold = bold or rpr.get("b") == "1"
                    c, a_ = solid_color(rpr)
                    if c:
                        color, alpha = c, a_
            elif child.tag == q("a:br"):
                chunks.append(cur)
                cur = ""
        if size is None:
            end = p.find("a:endParaRPr", NS)
            size = int(end.get("sz")) / 100 if end is not None and end.get("sz") else DEFAULT_SIZE
        chunks.append(cur)
        for ch in chunks:
            paras.append(Para(ch, size, bold, color, alpha))
    return paras, insets, autofit


def walk(tree, rels, items: list[Item], transform=None):
    for el in tree:
        tag = el.tag.split("}")[1]
        if tag == "grpSp":
            gx = el.find("p:grpSpPr/a:xfrm", NS)
            sub = transform
            if gx is not None:
                off, ext = gx.find("a:off", NS), gx.find("a:ext", NS)
                choff, chext = gx.find("a:chOff", NS), gx.find("a:chExt", NS)
                if None not in (off, ext, choff, chext):
                    ox, oy = int(off.get("x")) / EMU, int(off.get("y")) / EMU
                    ew, eh = int(ext.get("cx")) / EMU, int(ext.get("cy")) / EMU
                    cx, cy = int(choff.get("x")) / EMU, int(choff.get("y")) / EMU
                    cw, ch = int(chext.get("cx")) / EMU or 1, int(chext.get("cy")) / EMU or 1
                    sx, sy = ew / cw, eh / ch

                    def sub(x, y, w, h, _o=(ox, oy, cx, cy, sx, sy), _t=transform):
                        ox_, oy_, cx_, cy_, sx_, sy_ = _o
                        r = (ox_ + (x - cx_) * sx_, oy_ + (y - cy_) * sy_, w * sx_, h * sy_)
                        return _t(*r) if _t else r

            walk(el, rels, items, sub)
            continue
        if tag not in ("sp", "pic", "graphicFrame", "cxnSp"):
            continue
        cnv = el.find(".//p:cNvPr", NS)
        name = cnv.get("name", "") if cnv is not None else ""
        xfrm = el.find("p:spPr/a:xfrm", NS) if tag != "graphicFrame" else el.find("p:xfrm", NS)
        geo = parse_xfrm(xfrm, transform)
        if geo is None:
            continue  # placeholder kế thừa vị trí từ layout: bỏ qua
        it = Item(name, tag, *geo)
        sppr = el.find("p:spPr", NS)
        if sppr is not None and sppr.find("a:noFill", NS) is None:
            it.fill, it.fill_alpha = solid_color(sppr)
        pg = sppr.find("a:prstGeom", NS) if sppr is not None else None
        if pg is not None:
            it.geom = pg.get("prst", "rect")
        it.paras, it.insets, it.autofit = parse_text(el.find("p:txBody", NS))
        if tag == "pic":
            blip = el.find(".//a:blip", NS)
            if blip is not None:
                it.image = rels.get(blip.get(q("r:embed")))
            src = el.find(".//a:srcRect", NS)
            it.has_crop = src is not None and any(src.get(k) for k in ("l", "t", "r", "b"))
        items.append(it)


# ----------------------------------------------------------------------------
# Tiện ích hình học, chữ, màu
# ----------------------------------------------------------------------------
def intersect(a: Item, b: Item) -> float:
    dx = min(a.x + a.w, b.x + b.w) - max(a.x, b.x)
    dy = min(a.y + a.h, b.y + b.h) - max(a.y, b.y)
    return dx * dy if dx > 0 and dy > 0 else 0.0


def text_width_in(text: str, size: float, bold: bool) -> float:
    """Ước tính bề rộng chuỗi (inch). Hệ số 0.55 em: thận trọng cho cả serif lẫn sans."""
    em = 0.0
    for ch in text:
        if ch.isspace():
            em += 0.28
        elif ch.isupper() or ch.isdigit():
            em += 0.66
        elif ch in "il.,:;'|!\"“”‘’()[]":
            em += 0.3
        elif ch in "mwMW%@":
            em += 0.85
        else:
            em += 0.55
    return em * size / 72 * (1.06 if bold else 1.0)


def estimate_lines(p: Para, avail_w: float) -> int:
    if not p.text.strip():
        return 1
    words = p.text.split()
    lines, cur = 1, 0.0
    space = text_width_in(" ", p.size, p.bold)
    for wd in words:
        ww = text_width_in(wd, p.size, p.bold)
        if cur and cur + space + ww > avail_w:
            lines += 1
            cur = ww
        else:
            cur += (space if cur else 0) + ww
        while cur > avail_w > 0:  # một từ dài hơn cả khung
            lines += 1
            cur -= avail_w
    return lines


def hex_rgb(h: str) -> tuple[float, float, float]:
    h = h.strip("#")
    return tuple(int(h[i : i + 2], 16) / 255 for i in (0, 2, 4))


def blend(top: tuple, bottom: tuple, alpha: float) -> tuple:
    return tuple(t * alpha + b * (1 - alpha) for t, b in zip(top, bottom))


def luminance(rgb: tuple) -> float:
    def lin(c):
        return c / 12.92 if c <= 0.03928 else ((c + 0.055) / 1.055) ** 2.4

    r, g, b = (lin(c) for c in rgb)
    return 0.2126 * r + 0.7152 * g + 0.0722 * b


def contrast(c1: tuple, c2: tuple) -> float:
    l1, l2 = sorted((luminance(c1), luminance(c2)), reverse=True)
    return (l1 + 0.05) / (l2 + 0.05)


def image_size(data: bytes):
    if data[:8] == b"\x89PNG\r\n\x1a\n":
        return int.from_bytes(data[16:20], "big"), int.from_bytes(data[20:24], "big")
    if data[:6] in (b"GIF87a", b"GIF89a"):
        return int.from_bytes(data[6:8], "little"), int.from_bytes(data[8:10], "little")
    try:
        from PIL import Image  # type: ignore

        return Image.open(io.BytesIO(data)).size
    except Exception:
        return None


# ----------------------------------------------------------------------------
# Kiểm tra
# ----------------------------------------------------------------------------
def audit(path: str, margin: float, max_lines: int, min_font: float) -> tuple[list[Finding], dict]:
    z = zipfile.ZipFile(path)
    pres = ET.fromstring(z.read("ppt/presentation.xml"))
    sz = pres.find("p:sldSz", NS)
    W, H = int(sz.get("cx")) / EMU, int(sz.get("cy")) / EMU
    prels = read_rels(z, "ppt/presentation.xml")
    order = [prels[s.get(q("r:id"))] for s in pres.findall("p:sldIdLst/p:sldId", NS)]

    out: list[Finding] = []
    add = lambda sev, code, n, name, msg: out.append(Finding(sev, code, n, name, msg))

    if abs(W / H - 16 / 9) > 0.01:
        add("WARN", "SIZE", 0, "", f'Slide {W:.3f}x{H:.3f}" không phải 16:9 (cần 10x5.625")')

    names_per_slide: list[set[str]] = []
    morph_on: list[bool] = []

    for n, part in enumerate(order, start=1):
        raw = z.read(part)
        root = ET.fromstring(raw)
        rels = read_rels(z, part)
        morph_on.append(b"p159:morph" in raw or b":morph " in raw)

        bg_hex, _ = solid_color(root.find("p:cSld/p:bg/p:bgPr", NS))
        bg = hex_rgb(bg_hex or "FFFFFF")

        items: list[Item] = []
        walk(root.find("p:cSld/p:spTree", NS), rels, items)

        # --- MORPH: trùng tên
        seen: dict[str, int] = {}
        for it in items:
            if it.name.startswith("!!"):
                seen[it.name] = seen.get(it.name, 0) + 1
        for nm, cnt in seen.items():
            if cnt > 1:
                add("ERROR", "MORPH", n, nm, f"Tên Morph bị trùng {cnt} lần trên cùng slide — Morph sẽ ghép sai")
        names_per_slide.append(set(seen))

        body_lines = 0
        for idx, it in enumerate(items):
            r, b = it.x + it.w, it.y + it.h
            on_slide = r > 0 and b > 0 and it.x < W and it.y < H
            visible_text = it.has_text and any(p.alpha > 0.02 for p in it.paras)
            tiny = it.w < 0.05 or it.h < 0.05  # trạng thái "ẩn" của thẻ lật, v.v.

            # --- BOUNDS / MARGIN
            if not on_slide:
                if not it.name.startswith("!!") and not it.exempt:
                    add("WARN", "BOUNDS", n, it.name, "Phần tử nằm hẳn ngoài slide (không phải phần tử Morph)")
            elif not it.exempt:
                full_bleed = it.x <= 0.01 and r >= W - 0.01 and not it.has_text
                eps = 0.005
                if (it.x < -eps or it.y < -eps or r > W + eps or b > H + eps) and not full_bleed:
                    add("ERROR", "BOUNDS", n, it.name,
                        f'Vượt mép slide (x={it.x:.2f}, y={it.y:.2f}, w={it.w:.2f}, h={it.h:.2f})')
                elif not full_bleed and not tiny and (
                    it.x < margin - eps or it.y < margin - eps or r > W - margin + eps or b > H - margin + eps
                ):
                    side = []
                    if it.x < margin - eps: side.append(f"trái {it.x:.2f}\"")
                    if it.y < margin - eps: side.append(f"trên {it.y:.2f}\"")
                    if r > W - margin + eps: side.append(f"phải {W - r:.2f}\"")
                    if b > H - margin + eps: side.append(f"dưới {H - b:.2f}\"")
                    add("WARN", "MARGIN", n, it.name, f"Lấn lề an toàn {margin}\": " + ", ".join(side))

            if not visible_text or tiny:
                continue

            # --- OVERFLOW
            l, t, rr, bb = it.insets
            avail_w = max(it.w - l - rr, 0.01)
            avail_h = it.h - t - bb
            need_h = 0.0
            for p in it.paras:
                lines = estimate_lines(p, avail_w)
                need_h += lines * p.size * 1.2 / 72
                # Dòng nội dung = câu/bullet (>= 25 ký tự, < 20pt). Nhãn số liệu ngắn và tiêu đề không tính.
                if p.size < 20 and len(p.text.strip()) >= 25:
                    body_lines += lines
                longest = max((text_width_in(w_, p.size, p.bold) for w_ in p.text.split()), default=0)
                if longest > avail_w * 1.02:
                    add("ERROR", "OVERFLOW", n, it.name,
                        f'Một từ rộng {longest:.2f}" > khung {avail_w:.2f}" — sẽ bị bẻ giữa từ')
                if p.text.strip() and p.size < min_font:
                    add("WARN", "FONT", n, it.name, f"Cỡ chữ {p.size:g}pt < {min_font:g}pt: khó đọc khi trình chiếu")
            if need_h > avail_h * 1.08 + 0.02:
                sev = "INFO" if it.autofit else "WARN"
                add(sev, "OVERFLOW", n, it.name,
                    f'Ước tính cần {need_h:.2f}" chiều cao, khung chỉ có {avail_h:.2f}"'
                    + (" (đang bật autofit)" if it.autofit else ""))

            # --- OVERLAP giữa các khung chữ
            for other in items[idx + 1 :]:
                if not other.has_text or other.w < 0.05 or other.h < 0.05:
                    continue
                area = intersect(it, other)
                if area > 0.02 * min(it.w * it.h, other.w * other.h):
                    add("WARN", "OVERLAP", n, it.name, f'Đè lên khung chữ "{other.name}" ({area:.2f} in²)')

            # --- CONTRAST: ghép màu các lớp nằm dưới tâm khung chữ
            cx, cy = it.x + it.w / 2, it.y + min(it.h / 2, 0.3)
            under = bg
            over_image = False
            for low in items[:idx]:
                if low.contains(cx, cy):
                    if low.kind == "pic":
                        over_image = True
                    elif low.fill:
                        under = blend(hex_rgb(low.fill), under, low.fill_alpha)
            if it.fill:  # khung chữ có nền riêng
                under = blend(hex_rgb(it.fill), under, it.fill_alpha)
            for p in it.paras:
                if not p.text.strip():
                    continue
                fg = blend(hex_rgb(p.color or "000000"), under, p.alpha)
                ratio = contrast(fg, under)
                large = p.size >= 18 or (p.bold and p.size >= 14)
                need = 3.0 if large else 4.5
                if over_image:
                    add("INFO", "CONTRAST", n, it.name, "Chữ nằm trên ảnh: kiểm tra tương phản bằng mắt")
                    break
                if ratio < need:
                    sev = "ERROR" if ratio < 3.0 else "WARN"
                    add(sev, "CONTRAST", n, it.name,
                        f'Tương phản {ratio:.2f}:1 < {need}:1 cho chữ {p.size:g}pt "{p.text[:28]}"')
                    break

        if body_lines > max_lines:
            add("WARN", "LINES", n, "", f"Khoảng {body_lines} dòng chữ nội dung (> {max_lines}). Tách slide hoặc rút gọn")

        # --- MEDIA
        for it in items:
            if it.kind != "pic" or not it.image or it.image not in z.namelist():
                continue
            data = z.read(it.image)
            if it.image.lower().endswith(".gif") and len(data) > 8 * 1024 * 1024:
                add("WARN", "MEDIA", n, it.name, f"GIF nặng {len(data) / 1e6:.1f} MB (> 8 MB): giảm kích thước/khung hình")
            dims = image_size(data)
            if dims and not it.has_crop and it.h > 0:
                img_ratio, box_ratio = dims[0] / dims[1], it.w / it.h
                if abs(img_ratio / box_ratio - 1) > 0.03:
                    add("WARN", "MEDIA", n, it.name,
                        f"Ảnh bị méo: tỷ lệ gốc {img_ratio:.2f}, khung {box_ratio:.2f}. Tính h = w / {img_ratio:.3f}")

    # --- MORPH: có cặp "!!" giữa 2 slide liền nhau nhưng slide đích chưa bật Morph
    for i in range(1, len(order)):
        shared = names_per_slide[i - 1] & names_per_slide[i]
        if shared and not morph_on[i]:
            add("INFO", "MORPH", i + 1, "",
                f"{len(shared)} phần tử '!!' dùng chung với slide {i} nhưng slide {i + 1} chưa bật Morph")

    morph_map = []
    for i in range(1, len(order)):
        a, b = names_per_slide[i - 1], names_per_slide[i]
        morph_map.append({
            "transition": f"{i} -> {i + 1}",
            "morph": morph_on[i],
            "matched": sorted(a & b),
            "fade_in": sorted(b - a),
            "fade_out": sorted(a - b),
        })
    stats = {
        "file": path,
        "slides": len(order),
        "size_in": [round(W, 3), round(H, 3)],
        "morph_slides": [i + 1 for i, on in enumerate(morph_on) if on],
        "morph_map": morph_map,
    }
    return out, stats


def main() -> int:
    ap = argparse.ArgumentParser(description="Kiểm tra chất lượng file .pptx")
    ap.add_argument("pptx")
    ap.add_argument("--margin", type=float, default=0.5, help="lề an toàn tính bằng inch (mặc định 0.5)")
    ap.add_argument("--max-lines", type=int, default=6, help="số dòng chữ nội dung tối đa mỗi slide")
    ap.add_argument("--min-font", type=float, default=12, help="cỡ chữ nhỏ nhất (pt)")
    ap.add_argument("--json", action="store_true")
    ap.add_argument("--quiet-info", action="store_true", help="ẩn mục INFO")
    ap.add_argument("--morph-map", action="store_true", help="in danh sách phần tử '!!' ghép cặp giữa các slide")
    a = ap.parse_args()

    findings, stats = audit(a.pptx, a.margin, a.max_lines, a.min_font)
    if a.quiet_info:
        findings = [f for f in findings if f.severity != "INFO"]
    order = {"ERROR": 0, "WARN": 1, "INFO": 2}
    findings.sort(key=lambda f: (f.slide, order[f.severity], f.code))

    if a.json:
        print(json.dumps({"stats": stats, "findings": [f.__dict__ for f in findings]}, ensure_ascii=False, indent=2))
    else:
        print(f"{stats['file']}: {stats['slides']} slide, {stats['size_in'][0]}x{stats['size_in'][1]} in, "
              f"Morph ở slide {stats['morph_slides'] or 'không có'}")
        for f in findings:
            where = f"slide {f.slide}" if f.slide else "deck"
            nm = f' "{f.name}"' if f.name else ""
            print(f"{f.severity:<5} {f.code:<8} {where}{nm}: {f.message}")
        if a.morph_map:
            print("\nMorph map:")
            for m in stats["morph_map"]:
                flag = "Morph" if m["morph"] else "chưa bật Morph"
                print(f"  {m['transition']} ({flag}): {len(m['matched'])} ghép cặp")
                if m["fade_in"]:
                    print(f"      + hiện ra: {', '.join(m['fade_in'])}")
                if m["fade_out"]:
                    print(f"      - biến mất: {', '.join(m['fade_out'])}")
        counts = {s: sum(f.severity == s for f in findings) for s in order}
        print(f"\nTổng: {counts['ERROR']} ERROR, {counts['WARN']} WARN, {counts['INFO']} INFO")
    return 1 if any(f.severity == "ERROR" for f in findings) else 0


if __name__ == "__main__":
    sys.exit(main())
