#!/usr/bin/env bash
# render_preview.sh — xuất từng slide của file .pptx thành ảnh PNG để kiểm tra bằng mắt.
#
#   bash render_preview.sh deck.pptx [out_dir] [dpi]
#
# Kết quả:
#   out_dir/deck.pdf            bản PDF do LibreOffice dựng
#   out_dir/slide-01.png ...    mỗi slide một ảnh
#   out_dir/contact-sheet.png   lưới tất cả slide (cần python3 + Pillow)
#
# Lưu ý: LibreOffice không chạy Morph và chỉ hiện khung đầu của GIF.
# Font thiếu sẽ bị thay thế, nên cài đúng font của deck trước khi soi tràn chữ.
set -euo pipefail

PPTX="${1:?Cách dùng: render_preview.sh deck.pptx [out_dir] [dpi]}"
OUT="${2:-$(dirname "$PPTX")/preview}"
DPI="${3:-110}"

command -v soffice >/dev/null || { echo "Thiếu LibreOffice (soffice)"; exit 1; }
command -v pdftoppm >/dev/null || { echo "Thiếu poppler-utils (pdftoppm)"; exit 1; }

mkdir -p "$OUT"
NAME="$(basename "${PPTX%.*}")"
PROFILE="$(mktemp -d)"   # hồ sơ LibreOffice riêng: tránh lỗi khóa khi chạy song song
trap 'rm -rf "$PROFILE"' EXIT

soffice -env:UserInstallation="file://$PROFILE" --headless \
  --convert-to pdf --outdir "$OUT" "$PPTX" >/dev/null 2>&1

PDF="$OUT/$NAME.pdf"
[ -f "$PDF" ] || { echo "Chuyển PDF thất bại: $PPTX"; exit 1; }

rm -f "$OUT"/slide-*.png
pdftoppm -png -r "$DPI" "$PDF" "$OUT/slide"

# pdftoppm đặt tên slide-1.png hoặc slide-01.png tùy số trang; chuẩn hóa về 2 chữ số
for f in "$OUT"/slide-*.png; do
  n="$(basename "$f" .png | sed 's/^slide-0*//')"
  mv -n "$f" "$OUT/$(printf 'slide-%02d.png' "$n")" 2>/dev/null || true
done

python3 - "$OUT" <<'PY' || echo "(Bỏ qua contact sheet: cần Pillow)"
import sys, glob, os
from PIL import Image, ImageDraw
out = sys.argv[1]
files = sorted(glob.glob(os.path.join(out, "slide-*.png")))
if not files:
    sys.exit(0)
thumbs = [Image.open(f).convert("RGB") for f in files]
tw = 480
th = round(tw * thumbs[0].height / thumbs[0].width)
cols = 3 if len(thumbs) > 4 else 2
rows = -(-len(thumbs) // cols)
pad = 16
sheet = Image.new("RGB", (cols * (tw + pad) + pad, rows * (th + pad + 18) + pad), "white")
d = ImageDraw.Draw(sheet)
for i, im in enumerate(thumbs):
    x = pad + (i % cols) * (tw + pad)
    y = pad + (i // cols) * (th + pad + 18)
    sheet.paste(im.resize((tw, th)), (x, y + 18))
    d.text((x, y + 2), f"Slide {i + 1}", fill="black")
sheet.save(os.path.join(out, "contact-sheet.png"))
PY

echo "PDF:    $PDF"
echo "Ảnh:    $OUT/slide-XX.png ($(ls "$OUT"/slide-*.png | wc -l) slide)"
[ -f "$OUT/contact-sheet.png" ] && echo "Lưới:   $OUT/contact-sheet.png"
