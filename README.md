# claude-pptx-skills

Bộ 5 **Claude Code Skills** giúp Claude tạo file thuyết trình PowerPoint (`.pptx`) có thiết kế hiện đại, chuẩn 16:9 và chuyển động mượt như video, không cần plugin.

- Chuyển cảnh **Morph** (Biến hình): thẻ bay, phóng to/thu nhỏ, thanh tiến trình chạy, cột biểu đồ mọc lên, lật thẻ.
- **GIF động** nhúng thẳng vào slide: sóng lan truyền, biểu đồ chạy số, luồng dữ liệu, hạt chuyển động.
- **Biểu đồ và sơ đồ** xuất hiện từng bước (Progressive Reveal).
- **Kiểm tra tự động** trước khi giao: tràn chữ, đè chữ, lấn lề, tương phản màu, trùng tên Morph.

## 5 skill

| Skill | Chức năng | Lệnh tắt |
|-------|-----------|----------|
| [`pptx-morph-motion`](.claude/skills/pptx-morph-motion/SKILL.md) | Thiết kế chuỗi slide trạng thái dùng chung phần tử tên `!!` (vd. `!!hero-card`). Đổi `x`, `y`, `w`, `h`, `rotate`, `transparency` giữa các slide để Morph tạo hiệu ứng bay, zoom/pan, tiến trình, lật thẻ. Tự chèn Morph vào file. | `@morph` |
| [`pptx-visual-design`](.claude/skills/pptx-visual-design/SKILL.md) | Quy chuẩn thẩm mỹ cho slide 16:9 (10 × 5.625 inch), chống lối mòn "slide AI". 8 bố cục có code: Bento Grid, Asymmetric Split, Big Number Callout, Process Timeline, Spotlight Comparison, Full-bleed Hero, Executive Dashboard, Quote Slide. Thang chữ, tối đa 6 dòng/slide, glassmorphism giả lập. | `@pptx`, `@morph` |
| [`pptx-animated-media`](.claude/skills/pptx-animated-media/SKILL.md) | Viết script Python (Pillow/matplotlib) hoặc HTML Canvas để render hoạt họa thành GIF lặp liền mạch, rồi nhúng bằng `slide.addImage({ path: "animation.gif", ... })`. Tự chạy khi trình chiếu. | `@gif-pptx` |
| [`pptx-charts-diagrams`](.claude/skills/pptx-charts-diagrams/SKILL.md) | Biểu đồ Bar, Line, Doughnut, Radar và sơ đồ khối bằng pptxgenjs hoặc SVG → PNG. Tách một biểu đồ phức tạp thành 3–4 slide, mỗi slide tô sáng một phần và làm mờ phần còn lại. | `@chart-pptx` |
| [`pptx-quality-audit`](.claude/skills/pptx-quality-audit/SKILL.md) | Kiểm tra `.pptx` trước khi giao: tràn chữ, đè khung, lề ≥ 0.5 inch, tương phản WCAG, cỡ chữ, số dòng, tên Morph, GIF/ảnh. Xuất ảnh từng slide bằng LibreOffice + `pdftoppm` để soát bằng mắt. Tự sửa lỗi lề và tràn chữ. | `@audit-pptx` |

## Lệnh tắt

Gõ lệnh tắt ở bất kỳ đâu trong tin nhắn gửi Claude Code (đã khai báo trong [`CLAUDE.md`](CLAUDE.md)):

| Lệnh tắt | Claude sẽ làm |
|----------|---------------|
| `@pptx` | Chạy trọn quy trình: lập design plan → biểu đồ (nếu có số liệu) → Morph → GIF (nếu cần) → kiểm tra chất lượng → giao file `.pptx` 16:9 |
| `@morph` | Dùng `pptx-morph-motion` + `pptx-visual-design` để tạo các cặp slide gắn tên `!!` biến hình mượt |
| `@gif-pptx` | Dùng `pptx-animated-media` để tự tạo GIF/biểu đồ động và nhúng vào `.pptx` |
| `@chart-pptx` | Dùng `pptx-charts-diagrams` để tạo biểu đồ số liệu và sơ đồ xuất hiện từng bước |
| `@audit-pptx` | Dùng `pptx-quality-audit` để rà soát và sửa lỗi tràn chữ, lệch khung trong file `.pptx` |

Ví dụ:

```text
@pptx 10 slide giới thiệu xưởng rang cà phê Mộc cho nhà đầu tư, có số liệu doanh thu quý 3
@morph @chart-pptx báo cáo doanh thu 6 tháng, tô sáng tháng 9
@gif-pptx thêm hiệu ứng luồng dữ liệu App → API → Kho dữ liệu vào slide 4
@audit-pptx kiểm tra file output/deck.pptx
```

Trong Claude Code, gõ `@` sẽ mở gợi ý chọn file; cứ gõ tiếp `@pptx` rồi nhấn Space, Claude vẫn đọc được lệnh tắt. Cũng có thể gọi trực tiếp bằng slash command: `/pptx-morph-motion`, `/pptx-visual-design`, `/pptx-animated-media`, `/pptx-charts-diagrams`, `/pptx-quality-audit`.

## Bắt đầu nhanh

```bash
git clone https://github.com/hpb282/claude-pptx-skills.git
```

```bash
cd claude-pptx-skills && npm install
```

```bash
npm run demo
```

Lệnh trên chạy [`templates/morph-deck-generator.js`](templates/morph-deck-generator.js) và tạo `output/morph-demo.pptx`: 4 slide minh họa Morph (thẻ kính phóng to, ô bento sắp xếp lại, cột biểu đồ mọc, thẻ lật, thanh tiến trình chạy) với bố cục Bento Grid.

Kiểm tra file vừa tạo:

```bash
npm run audit
```

```bash
npm run preview
```

Dùng bộ skill trong dự án khác: sao chép thư mục `.claude/skills/` vào dự án đó (hoặc vào `~/.claude/skills/` để dùng ở mọi nơi), kèm `CLAUDE.md` nếu muốn dùng lệnh tắt.

## Bật hiệu ứng Morph trong PowerPoint

File do template tạo **đã bật sẵn Morph** cho slide 2 trở đi. Nếu bạn mở file và chưa thấy chuyển động (hoặc dùng `--no-morph`), chỉ cần **một bước**:

> Ở khung thumbnail bên trái, nhấn **Ctrl + A** (Mac: **Cmd + A**) để chọn tất cả slide → mở tab **Transitions** (Chuyển tiếp) → bấm **Morph** (Biến hình).

Sau đó nhấn **F5** để trình chiếu. Tùy chọn thêm: *Effect Options → Objects* (mặc định), *Duration* 1.00–1.50 giây.

Lưu ý:
- Morph cần **PowerPoint 365, PowerPoint 2019 trở lên** (Windows/Mac) hoặc PowerPoint trên web. Keynote, Google Slides, LibreOffice và PowerPoint 2016 trở xuống sẽ dùng hiệu ứng mờ dần (fade) thay thế.
- GIF động tự chạy trong mọi phiên bản PowerPoint, Keynote và Google Slides. Bản in và PDF chỉ hiện khung đầu tiên của GIF.

## Cấu trúc kho

```
claude-pptx-skills/
├── CLAUDE.md                          # bảng điều khiển + lệnh tắt cho Claude
├── README.md
├── package.json                       # pptxgenjs + script demo/audit/preview
├── templates/
│   └── morph-deck-generator.js        # 4 slide Morph + Bento Grid, tự chèn Morph
└── .claude/skills/
    ├── pptx-morph-motion/SKILL.md
    ├── pptx-visual-design/SKILL.md
    ├── pptx-animated-media/
    │   ├── SKILL.md
    │   └── scripts/
    │       ├── render_gif.py          # ripple, countup, flow, particles, frames
    │       ├── canvas_to_gif.js       # HTML Canvas → GIF qua Playwright
    │       └── canvas-wave.html       # trang Canvas mẫu
    ├── pptx-charts-diagrams/SKILL.md
    └── pptx-quality-audit/
        ├── SKILL.md
        └── scripts/
            ├── audit_pptx.py          # kiểm tra tự động (không cần thư viện ngoài)
            ├── render_preview.sh      # .pptx → PDF → PNG từng slide + lưới tổng
            └── fix_pptx.py            # tự sửa lề và tràn chữ (python-pptx)
```

## Yêu cầu

| Công cụ | Dùng cho | Cài đặt |
|---------|----------|---------|
| Node.js 18+ | Sinh `.pptx` bằng pptxgenjs | `npm install` |
| Python 3 | GIF, kiểm tra, sửa lỗi | `pip install pillow matplotlib numpy python-pptx` |
| LibreOffice + poppler | Ảnh xem trước từng slide | `sudo apt install libreoffice-impress poppler-utils` / `brew install --cask libreoffice && brew install poppler` |
| Playwright (tùy chọn) | HTML Canvas → GIF, SVG → PNG | `npm i -D playwright && npx playwright install chromium` |

## Câu hỏi thường gặp

**Vì sao phần tử mờ đi rồi hiện lại thay vì di chuyển?** Tên phần tử ở hai slide không khớp hoặc thiếu tiền tố `!!`. Chạy `python3 .claude/skills/pptx-quality-audit/scripts/audit_pptx.py output/deck.pptx --morph-map` để xem phần tử nào được ghép cặp.

**Font hiển thị khác trên máy khác?** pptxgenjs không nhúng font. Bộ skill mặc định dùng font hệ thống có đủ dấu tiếng Việt (Georgia, Arial…). Nếu dùng font riêng, máy trình chiếu cũng phải cài font đó.

**Số liệu trong template là thật?** Không. "Xưởng rang Mộc" và mọi con số trong template chỉ để minh họa.
