#!/usr/bin/env bash
# Regenerate the README slide gallery from real pptx-ts output.
#
#   1. build the two showcase decks (pnpm showcases:build)
#   2. build features.pptx (HTML table, text fit, image in shape) from features-deck.mjs
#   3. render every deck through LibreOffice to PDF, then to PNG
#   4. compose gallery.webp (transparent) plus gallery-light.jpg / gallery-dark.jpg,
#      and one cropped card per README feature under features/
#
# Needs: pnpm, node, soffice, pdftoppm, magick (with WebP support).
set -euo pipefail

HERE="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
ROOT="$(cd "$HERE/../../.." && pwd)"
WORK="$HERE/work"
SLIDES="$WORK/slides"
mkdir -p "$WORK" "$SLIDES" "$HERE/features"

# --- 1-2. decks ------------------------------------------------------------------------
if [[ "${SKIP_DECKS:-0}" != 1 ]]; then
	(cd "$ROOT" && pnpm showcases:build)
	node "$HERE/features-deck.mjs" "$WORK/features.pptx"
fi
OUT="$ROOT/.tmp/showcases"
cp "$OUT/Kestrel_Q3_Business_Review.pptx" "$WORK/kestrel.pptx"
cp "$OUT/Field_Notes_Four_Cities.pptx" "$WORK/field.pptx"

# --- 3. render -------------------------------------------------------------------------
# One page = one slide. 1600px wide is enough for both the grid and the feature cards.
for deck in kestrel field features; do
	soffice --headless --convert-to pdf --outdir "$WORK" "$WORK/$deck.pptx" >/dev/null 2>&1
	rm -f "$SLIDES/$deck"-*.png
	pdftoppm -scale-to-x 1600 -scale-to-y -1 -png "$WORK/$deck.pdf" "$SLIDES/$deck"
done
# pdftoppm pads the page number to the page count's width; normalise to no padding.
for f in "$SLIDES"/*-0*.png; do
	[[ -e "$f" ]] || continue
	b=$(basename "$f")
	mv "$f" "$SLIDES/$(echo "$b" | sed -E "s/-0+([0-9])/-\\1/")"
done

# --- 4. compose ------------------------------------------------------------------------
# card <in> <out> <width>: resize, round the corners, add a hairline border that reads on
# both white and #0d1117, and a soft drop shadow, on a transparent canvas.
card() {
	local in="$1" out="$2" w="$3"
	local r=$((w / 70 + 4))
	local h
	h=$(magick "$in" -resize "${w}x" -format '%h' info:)
	magick "$in" -resize "${w}x" -filter Lanczos \
		\( -size "${w}x${h}" xc:none -fill white -draw "roundrectangle 0,0,$((w - 1)),$((h - 1)),$r,$r" \) \
		-alpha set -compose DstIn -composite \
		\( -size "${w}x${h}" xc:none -fill none -stroke 'rgba(128,128,128,0.45)' -strokewidth 1.5 \
		-draw "roundrectangle 0.75,0.75,$((w - 1)).25,$((h - 1)).25,$r,$r" \) \
		-compose Over -composite \
		\( +clone -background 'rgba(0,0,0,1)' -shadow 38x9+0+5 \) +swap \
		-background none -compose Over -layers merge +repage "$out"
}

# The grid: dark and light slides in a checkerboard so no row reads as one block.
GRID=(
	kestrel-1 kestrel-5 field-2
	kestrel-4 field-6 kestrel-8
	field-4 kestrel-10 field-7
)
COLS=3 GAP=26 PAD=30 CANVAS_W=1600
TW=$(((CANVAS_W - 2 * PAD - (COLS - 1) * GAP) / COLS))
rm -f "$WORK"/card-*.png
i=0
for s in "${GRID[@]}"; do
	card "$SLIDES/$s.png" "$WORK/card-$i.png" "$TW"
	i=$((i + 1))
done
# Every card has the same size (slide + shadow), so place them on a fixed pitch.
CW=$(magick "$WORK/card-0.png" -format '%w' info:)
CH=$(magick "$WORK/card-0.png" -format '%h' info:)
TH=$(magick "$SLIDES/${GRID[0]}.png" -resize "${TW}x" -format '%h' info:)
OFF_X=$(((CW - TW) / 2))
ROWS=$(((${#GRID[@]} + COLS - 1) / COLS))
CANVAS_H=$((2 * PAD + ROWS * TH + (ROWS - 1) * GAP + 8))
args=()
for ((k = 0; k < ${#GRID[@]}; k++)); do
	cx=$((PAD + (k % COLS) * (TW + GAP) - OFF_X))
	cy=$((PAD + (k / COLS) * (TH + GAP) - OFF_X + 4))
	args+=("$WORK/card-$k.png" -geometry "+$cx+$cy" -composite)
done
magick -size "${CANVAS_W}x${CANVAS_H}" xc:none "${args[@]}" "$WORK/gallery-full.png"
magick "$WORK/gallery-full.png" -quality 84 -define webp:alpha-quality=90 -define webp:method=6 "$HERE/gallery.webp"
magick "$WORK/gallery-full.png" -background '#ffffff' -flatten -quality 82 -sampling-factor 4:2:0 -strip "$HERE/gallery-light.jpg"
magick "$WORK/gallery-full.png" -background '#0d1117' -flatten -quality 82 -sampling-factor 4:2:0 -strip "$HERE/gallery-dark.jpg"

# Feature cards, one per README bullet. 960px wide, transparent, same treatment.
fcard() { card "$SLIDES/$1.png" "$WORK/f.png" 960 && magick "$WORK/f.png" -quality 86 -define webp:method=6 "$HERE/features/$2.webp"; }
fcard kestrel-5 chart
fcard kestrel-8 table
fcard kestrel-4 kpi-group
fcard features-3 text-fit
fcard features-4 image-in-shape
fcard field-6 photo-scrim

# HTML table to slides: the two pages it produced, the second fanned behind the first.
card "$SLIDES/features-1.png" "$WORK/p1.png" 860
card "$SLIDES/features-2.png" "$WORK/p2.png" 860
magick -size 1000x600 xc:none \
	\( "$WORK/p2.png" -background none -rotate 3 \) -geometry +118+14 -composite \
	\( "$WORK/p1.png" \) -geometry +0+70 -composite \
	-trim +repage "$WORK/html.png"
magick "$WORK/html.png" -bordercolor none -border 12 -quality 86 -define webp:method=6 "$HERE/features/html-table-to-slides.webp"

ls -l "$HERE"/gallery* "$HERE"/features/
