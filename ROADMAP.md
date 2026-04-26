# WallBrand Roadmap

Forward-looking scope for the wallpaper branding tool. Aimed at corporate branding, digital signage, and personal batch watermarking workflows.

## Planned Features

### Compositing
- Multi-layer pipeline (bg image + solid/gradient fill + logo + text + shape overlay) with per-layer blend modes (multiply, screen, overlay, soft light).
- Drop shadow, outer glow, inner bevel on logo and text layers.
- Smart scaling: percentage-of-height or DPI-aware sizing instead of raw pixels so 4K/8K wallpapers don't require rebuild.
- Safe-zone overlay that mirrors Windows lock-screen/taskbar reserved regions per resolution.

### Batch & Automation
- Folder-watch CLI mode: drop an image into an input folder and get a branded output written to a sibling folder.
- Template presets saved as JSON with embedded logo paths, export to recreate a campaign in one command.
- Multi-resolution export matrix (1080p, 1440p, 2160p, ultrawide, mobile portrait) from one source composition.
- ICC/color-profile-aware export so corporate brand color (e.g., Pantone → sRGB) matches the style guide.

### Targeting & Deployment
- Intune / GPO wallpaper deployment helper: drop branded image + generate `DesktopImagePath` / `LockScreenImagePath` policy ADMX snippet.
- Active Directory OU targeting: render different branding per OU and emit per-OU GPO snippets.
- Active Setup / Run-once script generator for one-time wallpaper push on new logons.
- Digital-signage `.bat` that cycles wallpapers every N minutes via scheduled task.

### GUI / UX
- Live side-by-side multi-monitor preview (each monitor with its own composition).
- Drag-to-reposition logo/text handles in the preview pane with snap-to-safe-zone guides.
- Undo/redo stack per composition session.
- Branding package export: zip containing JSON template + logo + fonts.

## Competitive Research
- **SnapComms / Netpresenter Corporate Wallpaper** — enterprise targeting per department/location + schedule rotation; WallBrand should match the scheduling + targeting story, at least via GPO snippets.
- **Watermarkly / Visual Watermark** — 900+ fonts, template reuse; WallBrand's edge is free + scriptable + no cloud upload, but should keep feature parity on text effects.
- **Canva watermark feature** — template marketplace is a moat; out of scope but ship 5-10 starter templates to anchor new users.

## Nice-to-Haves
- AI-assisted logo placement that detects low-entropy regions of the background and recommends the optical-center corner.
- Animated / video wallpaper export (MP4) for digital signage screens.
- Font embedding check that warns when a .ttf file isn't licensed for redistribution before baking into a signage template.
- Command-line module published as `Install-Module WallBrand` for CI pipelines.
- SVG-source logo support with runtime tint (swap brand color without re-exporting from the design tool).
- Grouped campaign mode: apply the same branding to a photo set (e.g., department-specific hero shots) with consistent crop/position.

## Open-Source Research (Round 2)

### Related OSS Projects
- https://github.com/WaGi-Coding/Simple-Batch-Image-Watermarker — Windows GUI, EXIF copyright, repeat-mode
- https://github.com/sandeshpoudel/batch_logo_overlay — fastest batch logo overlay reference
- https://github.com/applegrew/Batch-Watermarker — Java CLI with EXIF preservation
- https://github.com/mbgh/watermark-img — ImageMagick drop-target .bat reference
- https://github.com/jlengrand/batchWaterMarking — orientation-aware corner placement
- https://github.com/gppam/batch-img-watermark — ImageMagick minimal reference
- https://github.com/ImageMagick/ImageMagick — underlying engine for all CLI variants
- https://github.com/python-pillow/Pillow — Python-native compositing alternative

### Features to Borrow
- EXIF copyright tag write on output (Simple-Batch-Image-Watermarker, Batch-Watermarker)
- Repeat-mode: tile watermark from center so crops still show branding (Simple-Batch-Image-Watermarker)
- Orientation auto-detect: different corner for portrait vs. landscape (jlengrand)
- Explorer drag-target batch files beside the GUI for "drop folder → branded output" workflow (mbgh)
- Live preview panel with opacity/size sliders that render on sample image (Tkinter app variant)
- Per-file output-folder templating `{YYYY}\{MM}\{orig}_branded.{ext}`
- Presets as JSON: logo path + anchor + size + opacity + EXIF fields, shareable across ops (Axiom-style)
- Watermark scaling: auto-resize logo to N% of shortest edge instead of fixed pixels
- Metadata preservation: copy EXIF/XMP/IPTC from source to output (Batch-Watermarker)
- Multi-layer overlays: logo + date + text signature in one pass

### Patterns & Architectures Worth Studying
- Dual entry points: shared core module + separate GUI + CLI shims that both call it (already partial)
- Worker pool for batches: PowerShell runspaces or `ForEach-Object -Parallel` for .NET 7+ hosts
- Color-profile-aware compositing: convert to sRGB before overlay, restore profile on save (Pillow/ImageMagick convention)
- Signed presets: GPG-sign JSON preset files so corporate templates can't be tampered with
- Dry-run mode: enumerate inputs and print planned outputs without writing anything
