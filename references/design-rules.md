# Deck Design Rules

These are working rules for generating presentation slides that look designed, not just assembled.
Units assume a 16:9 slide of 13.333 x 7.5 in (1280 x 720 px at 96 dpi, so 1 pt = 1.33 px).
Treat every number as a default with a sensible range. Break a rule only when you mean to, and only once per slide.

---

## 1. Core principles

1. **Scale contrast is the whole game.** The largest type on a slide is 4-8x body size on content slides and 10-20x on covers, sections and closings. If the ratio is under 3x, the slide reads as a document.
2. **One hero per slide.** Each slide gets exactly one dominant element (a headline, a numeral, an image or a chart). The next element down is at most 50% of its visual weight.
3. **Margins are sacred.** Keep a safe margin of at least 5% of slide width (0.65-0.75 in) on all four sides. Only full-bleed images, colour fields and deliberately cropped display type may cross it, and they must cross the edge completely. Never stop 2-10 px short of it.
4. **Everything snaps to one grid.** Use 12 columns with 0.2-0.25 in gutters. Every left edge on the slide sits on a column line. Count the distinct left edges: 3 or fewer is calm, 5 or more looks broken.
5. **Colour discipline.** One neutral ground, one ink colour and one accent. The accent covers no more than 10% of the area on light or dark slides, unless the accent *is* the ground (a colour-field slide). A second accent is allowed only as a tint or shade of the first.
6. **Use the accent for meaning, not decoration.** Spend it on the key word, the key figure, the active tier or the current step. If three things are accented, nothing is.
7. **Word budgets beat font shrinking.** Never drop body text below 12 pt to fit content. Cut the words or split the slide instead. Budgets: airy 0-15 words, light 15-50, detailed 50-120. More than 120 words is a handout.
8. **Hairlines over boxes.** Separate content with 0.5-1 pt rules and whitespace before you reach for filled cards. Use cards only when the items are true peers that need equal weight, such as tiers or modules.
9. **Repetition makes it a system.** A running header or footer (a label left, a section name centred or right, a page counter as "n / N") sits on the same baseline on every content slide. Covers, sections and closings may drop it.
10. **Numbers are display type.** Any figure that carries the point is set at 48-120 pt in the heading face, with a small all-caps or small-weight label under it. Never put a key number inside a sentence on a stat slide.
11. **Tension through cropping and offset.** Great slides break symmetry on purpose. Examples are a headline bleeding off one edge, an image straddling a colour seam, or a numeral overlapping a frame. Use at most one such gesture per slide.
12. **Two typefaces, maximum.** Use a display face (often condensed, heavy or editorial serif) plus a neutral sans for text. An optional mono or caps label style is a third *role*, not a third family.
13. **Images are full or framed, never floating.** A photo either bleeds to at least one edge or sits in a deliberate frame aligned to the grid, with a consistent radius. Small stock images in the middle of whitespace are always wrong.
14. **Mirror the open in the close.** The closing slide reuses the cover's composition, ground and type scale so the deck feels bookended.
15. **Group by proximity.** The space inside a group (label to value) is at most half the space between groups. If everything is equally spaced, nothing is grouped.

---

## 2. Measurements cheat-sheet (13.333 x 7.5 in)

**Frame**
| Item | Value |
|---|---|
| Side margins | 0.65-0.75 in (about 5-5.5% of width); airy editorial up to 1.0 in |
| Top margin to header baseline | 0.35-0.5 in |
| Header rule (optional) | 0.55-0.6 in from top, full margin-to-margin width |
| Footer baseline | 0.35-0.45 in from bottom |
| Title zone top | 1.0-1.3 in from top on content slides |
| Live content area | about 11.9 x 5.6 in |
| Grid | 12 columns, 0.2-0.25 in gutter. Common splits: 6/6, 5/7, 4/8, 3x4, 4x3 |
| Vertical rhythm | 0.125 in base unit; all gaps are multiples of it (0.25 / 0.5 / 0.75 / 1.0) |

**Type scale (pt)**
| Role | Size | Weight | Line-height | Tracking |
|---|---|---|---|---|
| Display (cover, section, closing, statement) | 96-200 | Heavy or condensed bold; light serif also works | 0.85-0.95 | -1% to -3% |
| Giant numeral (section index, hero stat) | 120-260 | Same as display | 0.8-0.9 | -2% |
| H1 (content-slide title) | 40-60 | Bold | 0.95-1.05 | -1% to -2% |
| H2 (card or column title) | 18-24 | Semibold | 1.1-1.2 | 0 |
| Stat figure | 48-96 (40-56 when 4+ in a row) | Display | 0.9 | -1% |
| Lead / pull quote | 22-32 | Regular or light | 1.25-1.35 | 0 |
| Body | 12-16 (14 default) | Regular | 1.4-1.55 | 0 |
| Label / eyebrow / meta | 8-11 | Medium-bold, often caps or mono | 1.2 | +8% to +15% (caps only) |
| Footer / page counter | 8-9 | Regular | 1.2 | +5% |

The ratios matter more than the absolute sizes: display / H1 is about 2.5-3.5x, H1 / body is about 3-4x, and body / label is about 1.4x.

**Lines, shapes, imagery**
| Item | Value |
|---|---|
| Hairline divider | 0.5-0.75 pt, ink at 15-25% opacity (or a mid-grey) |
| Structural rule (header, table top) | 1-1.5 pt, full ink |
| Accent dash under a figure | 2-3 pt thick, 0.3-0.4 in long |
| Corner radius | Pick one family: 0 (Swiss or brutal), 0.08-0.15 in (soft), or 0.25-0.35 in (friendly, pill-like cards) |
| Pills and tags | Height 0.25-0.3 in, radius = half the height, 1 pt outline or solid fill |
| Hard offset shadow (playful styles only) | 0.08-0.12 in, solid ink, no blur |
| Image to text ratio, image-led slide | 55-65% image / 35-45% text |
| Image to text ratio, text-led with image | 35-45% image |
| Full-bleed image | 100% of the slide; text sits on a 40-60% scrim or in a solid panel |
| Card padding | 0.25-0.35 in inside, never less than 0.2 in |
| Gap between cards | 0.15-0.25 in (equals the gutter) |

---

## 3. Archetypes

Content limits are given as **minimal / light / detailed**. Title = maximum words in the title, items = maximum count, item = maximum words per item.

### Cover
*Purpose:* set the voice in two seconds and name the subject.
*Patterns:* (a) a giant wordmark or title filling 60-90% of the width, in 1-3 stacked lines, with one line or word in the accent or set in italic or outline, plus a quiet meta row (date, author, edition) along the footer. (b) A 6/6 or 7/5 split: stacked display title on one side and a framed image or colour panel on the other. (c) Pure colour-field: two or three vertical bands or a gradient ground with the title spanning across them.
*Limits:* title 2-4 / 5 / 8. Subtitle 0 / 12 / 25 words. Meta items 0 / 3 / 4.
*Avoid:* a centred title at 44 pt in the middle of empty space; a logo bigger than the title; more than one supporting paragraph; a decorative stock photo behind small type.

### Section divider
*Purpose:* reset attention and mark position in the deck.
*Patterns:* (a) a huge index numeral (150-260 pt) in the accent, set against a title one-third its size, with ample void around it. (b) A single word or short phrase at display scale, with one tiny supporting label. (c) An asymmetric colour split, with the numeral on the coloured half and the title or illustration on the other.
*Limits:* title 1-3 / 5 / 6. Items 0 / 0 / 4 (a "what's in this part" list). Item 0 / 0 / 6.
*Avoid:* reusing the cover layout exactly; adding bullets; dropping the numeral so dividers stop reading as a sequence.

### Statement / manifesto
*Purpose:* make one claim land.
*Patterns:* (a) a huge 3-5 line headline, centred or left, that may sink off the bottom edge, flanked by one or two tiny side notes. (b) A large lead-size sentence (28-40 pt, light weight) with one phrase in bold or accent, plus a quiet attribution line. (c) A two-tone heading on the left (5 cols) and a lead statement on the right (6 cols).
*Limits:* statement 6 / 15 / 25 words. Side notes 0 / 1 / 2 at 25 words each.
*Avoid:* full paragraphs at lead size; more than one emphasised phrase; quotation marks when it is not a quote.

### Agenda / contents
*Purpose:* show the shape of the talk.
*Patterns:* (a) a title anchored bottom-left at display size, with a ruled numbered list in the right half (number, title, optional page number right-aligned). (b) A full-width table: index, title, one-line descriptor and a tag, each row separated by hairlines. (c) Numbered chips or dots in a column beside a framed note or image.
*Limits:* items 4 / 6 / 8. Item 3 / 5 / 12 words (with a descriptor).
*Avoid:* more than 8 rows; bullets instead of numbers; row heights that drift; numbers that do not match the section dividers.

### KPI / stats
*Purpose:* make numbers memorable.
*Patterns:* (a) a 2x2 or 1x4 grid of big figures separated by hairlines, each with an accent dash and a caps label, with the accent colour on one or two key figures only. (b) A single hero figure at 150-250 pt with a short explanation, often with an oversized unit or percent sign as a graphic element. (c) Figures stepping in tone or height (light to dark tiles, or rising bars behind them) so the row itself shows growth.
*Limits:* figures 1 / 3 / 4 (6 maximum in a 3x2 grid). Label 3 / 6 / 10 words. Title 3 / 6 / 8.
*Avoid:* different number formats in one row; labels longer than the figure is wide; every figure in the accent; mixing currency, percent and count with no visual grouping.

### Process / steps
*Purpose:* show ordered movement.
*Patterns:* (a) 3-5 equal columns, each with a large step numeral, a short title and 1-2 lines, joined by a thin rule with small node dots or arrows. (b) Numbered horizontal rows (index, title, description) stacked down the right side, with the title block on the left. (c) Connected cards with an icon tab on top, linked by arrows, at most one colour per step.
*Limits:* steps 3 / 4 / 5 (6 maximum). Step title 1-2 / 3 / 4 words. Step body 0 / 12 / 25 words.
*Avoid:* numbers out of order; arrows that do not align with the centres; uneven column widths; more than 5 colours.

### Timeline / roadmap
*Purpose:* place events in time.
*Patterns:* (a) a horizontal axis line with nodes and alternating labels above and below (zig-zag), with phase brackets grouping the nodes. (b) Large year or phase labels (60-120 pt, light or condensed) in 3-4 columns, each with a short description, the display title anchored in a corner. (c) Year tabs or badges on a hairline, stepping through tints of one hue.
*Limits:* nodes 3 / 5 / 8. Node label 2 / 4 / 6 words. Node body 0 / 10 / 18 words.
*Avoid:* unequal spacing for equal time intervals (or equal spacing for unequal ones, unless labelled); more than 8 nodes on one axis; a diagram too small to read from the back of the room.

### Comparison
*Purpose:* contrast two or more options and make the winner obvious.
*Patterns:* (a) two side-by-side panels: the "before" or "without" option in a muted neutral, the "after" or "with" option in the dark or accent ground, with matching rows of minus and check marks. (b) Columns with a shared row structure, separated by hairlines, the preferred column raised or inverted. (c) A hub layout: a few labelled points on one side joined by a bracket to a claim on the other.
*Limits:* options 2 / 2 / 4. Rows 3 / 4 / 6. Row 4 / 8 / 12 words.
*Avoid:* asymmetric row counts; both panels in the same weight; using the accent on the losing side.

### Cards / services / features
*Purpose:* present peer items with equal weight.
*Patterns:* (a) 3-4 ruled columns, each topped by a numeral (01, 02, 03) in the accent, a title and short copy, with no boxes. (b) A 2x3 or 4-up card grid where one card is inverted (dark or accent) to mark the hero item. (c) A bento grid mixing colour tiles and image tiles of different spans, title in one tile. (d) Full-width numbered rows: index, title, description, arrow.
*Limits:* items 3 / 4 / 6. Title 1-2 / 3 / 4 words. Body 0 / 15 / 30 words.
*Avoid:* icon soup (decorative icons on every card); cards of unequal height; text overflowing a coloured tile; 7 or more cards.

### Image + text
*Purpose:* let a picture carry the idea, with text explaining it.
*Patterns:* (a) a half or two-thirds bleed image on one side, with headline, short body and an optional mini-stat row on the other. (b) An image bleeding from the top edge with a large statement below or beside it, anchored by an index numeral. (c) An image framed on the grid with a caption block and a stat footer aligned to the image edge.
*Limits:* title 4 / 8 / 10. Body 0 / 30 / 70 words. Bullets 0 / 3 / 4.
*Avoid:* images with a margin on three sides and a bleed on one (pick one); text on a busy image without a scrim; mismatched image and text heights.

### Quote / testimonial
*Purpose:* borrow credibility from another voice.
*Patterns:* (a) a large quote (26-40 pt) on a solid ground, with an oversized quotation mark in the accent and a small avatar plus name and role below a hairline. (b) A half-bleed photo with the quote in light weight on the other half and tiny meta in the footer. (c) One hero testimonial card plus 2 smaller ruled quotes, with a rating or summary line in mono.
*Limits:* quote 12 / 30 / 50 words. Quotes per slide 1 / 1 / 3.
*Avoid:* italic body plus a quote mark plus a box (pick one device); quotes longer than 4 lines at lead size; missing attribution.

### Team / people
*Purpose:* put faces to the work.
*Patterns:* (a) a single-person profile: a large portrait (40-50% of width), a big name, role, 2-3 skill tags and an optional stat row. (b) 3-4 equal portrait columns with names and roles set inside or under each, all portraits treated the same way (same crop, tone and radius). (c) An expressive variant: an oversized italic or script word overlapping the portrait.
*Limits:* people 1 / 4 / 6. Bio 0 / 25 / 50 words. Role 2 / 4 / 6 words.
*Avoid:* mixed photo styles (some colour, some mono, different crops); names smaller than roles; more than 6 faces (use a grid of names instead).

### Chart / data
*Purpose:* show one finding, not a dataset.
*Patterns:* (a) a chart taking 55-65% of the area, with the title plus a one-line takeaway and 1-2 callout figures in a side column. (b) Bars bleeding off the bottom or side edge with large value labels on the bars and no axis. (c) A donut or pie in neutral greys with one segment in the accent, plus a ruled legend list of large percentages.
*Limits:* series 1 / 2 / 3. Categories 3 / 6 / 10. Takeaway 6 / 12 / 20 words.
*Avoid:* default chart styling (gridlines, 3D, legends far from the data); more than one accent colour; unlabelled axes when an axis is shown; decimals beyond what matters.

### Pricing / tiers
*Purpose:* make the choice easy and steer it.
*Patterns:* (a) three equal tier cards, the middle one raised, outlined or inverted, with a small "popular" tab and the price as display type with a small unit or period. (b) 3-4 columns separated only by hairlines, with the price in a large serif or condensed face and small cents. (c) Colour-stepped tiles (tints of one hue) with the price on top.
*Limits:* tiers 2 / 3 / 4. Features per tier 3 / 5 / 6. Feature 3 / 5 / 8 words.
*Avoid:* more than one highlighted tier; inconsistent price formats; feature lists of different lengths with no alignment; the call-to-action text larger than the price.

### Gallery / portfolio
*Purpose:* show work and let the images speak.
*Patterns:* (a) 2-3 large images bleeding to an edge, with a big index numeral and a narrow text column on the other side. (b) A masonry or bento grid with one image spanning two cells, captions as tiny labels inside a bottom gradient. (c) An even row of 4 frames with a caption bar, above two short lists.
*Limits:* images 1 / 4 / 6. Caption 0 / 4 / 8 words. Body 0 / 25 / 50 words.
*Avoid:* gutters that differ between images; mixed aspect ratios with no grid logic; titles covering faces or focal points.

### Closing
*Purpose:* end with intent and give the next step.
*Patterns:* (a) the cover composition repeated with a sign-off phrase at the same display scale. (b) A huge two-line sign-off (one line in the accent) with a contact row of 2-3 label-value pairs along the bottom. (c) Poster scale: the sign-off fills the slide edge to edge, with a single URL as the only other text.
*Limits:* sign-off 2 / 4 / 6 words. Contact items 0 / 3 / 4.
*Avoid:* a generic small "thank you" centred on white; a closing that ignores the cover; a full contact block in body size competing with the sign-off.

---

## 4. Deck-level rhythm

- **Alternate density.** Never place three detailed slides in a row. A good pattern is airy, light, detailed, light, airy (statement or section), then repeat. Aim for 30% airy, 45% light and 25% detailed across the deck.
- **Alternate ground.** Pick a base ground (light or dark) for about 70% of slides. Use the inverse ground or a full accent field for about 20-30%: dividers, statements, the hero stat and the close. A ground switch is a signal, so use it at structural moments, not at random.
- **Alternate the hero type.** Rotate the hero element between type, number, image and diagram. Two consecutive slides should not have the same layout unless they are a deliberate series (such as a run of steps or case studies).
- **Image budget.** In a 10-slide stretch, aim for 3-4 slides with meaningful imagery: 1-2 bleed or hero images and 1-2 framed or supporting images. A deck with no imagery needs stronger type and colour fields to compensate.
- **Section cadence.** Place a divider every 4-7 content slides. A 20-slide deck usually has 3-4 sections.
- **Opening sequence.** Cover, then agenda or statement, then the first content. The second slide should be the calmest in the deck.
- **Closing sequence.** Summary (stat or statement), then call to action or pricing, then the closing slide that mirrors the cover.
- **Consistency anchors.** Keep these identical across all content slides: the header and footer, the title position, the margin, the accent colour and the radius family. Vary composition, not the system.

---

## 5. RTL (Hebrew) notes

**Mirror**
- Reading direction and alignment: titles and body are right-aligned, and the layout's primary column moves to the right.
- Asymmetric splits: if the English layout puts the text on the left and the image on the right, swap them.
- Lists and process flows: step 1 is on the right and arrows point left. Timelines run right to left.
- Header and footer: the brand label goes on the right and the page counter on the left.
- Bullet markers, check marks and arrow glyphs appear on the right side of the text.

**Do not mirror**
- Numerals and number groups (2026, 98%, +35, phone numbers), which stay left to right inside the line. Wrap them in an LTR run when needed.
- Logos, wordmarks, product UI screenshots, code and URLs.
- Charts: keep the x-axis running left to right with time increasing to the right (this is the convention Hebrew readers expect for data), and keep the bar order and the value axis as they are. Mirror only the chart's title block and legend position.
- Photographs, unless a face or gaze directs attention off the slide. In that case flip the layout, not the photo.
- Icons with real-world direction (play, clock, chart-up arrows).

**Typography**
- Hebrew has no capitals. Labels and eyebrows that rely on caps plus tracking need another approach: use a heavier weight, the accent colour, or a small size with a rule or dash. Use +2-4% tracking at most, because wide tracking breaks Hebrew word shapes.
- Hebrew has no true italics. Do not oblique it. Create emphasis with weight, colour or a second face, never slant.
- Hebrew glyphs look smaller and denser at the same point size. Raise body text by 1-2 pt (16 instead of 14) and line-height by about 0.05-0.1.
- Display Hebrew sets tighter vertically. Line-height of 0.95-1.05 is safer than 0.85 for stacked display lines.
- Hebrew text runs about 10-20% shorter than English. Do not stretch it to fill English-sized boxes, and let the extra whitespace stand.
- Pair a Hebrew display face whose weight matches the Latin face when the two languages mix. Use one family that covers both scripts when possible.
- For mixed lines (Hebrew with English terms), check that punctuation and parentheses resolve on the correct side.

---

## 6. Self-review checklist (run on the rendered image)

1. **Hero:** can you name the single dominant element within one second? Is it at least 2x the visual weight of the next element?
2. **Scale ratio:** is the largest type at least 4x body size (10x on cover, section and closing)?
3. **Margins:** is all non-bleed content at least 0.65 in (about 62 px at 1280 wide) from every edge?
4. **Bleeds:** does every element that touches an edge cross it fully, with no 1-10 px slivers or near misses?
5. **Alignment:** count the distinct left edges (right edges in RTL). Are there 4 or fewer, all on grid columns?
6. **Accent area:** is the accent at most about 10% of the area (unless the slide is a colour field), and on the meaningful element?
7. **Colour count:** are there at most 3 non-image colours (ground, ink, accent), plus tints?
8. **Fonts:** are there at most 2 families visible, with consistent weights per role?
9. **Word count:** is the slide within its density budget? Is any text under 12 pt (body) or under 8 pt (meta)?
10. **Overflow:** is any text clipped, overlapping another element, running past its card or colliding with an image?
11. **Line lengths:** do body lines stay at 45-75 characters? Is there a lone widow word on the last line of any heading?
12. **Repeated items:** are cards, columns and rows equal in size, spacing and internal structure?
13. **Sequence:** are numbers and steps in order and formatted the same way (01 vs 1 vs 001)?
14. **Number styling:** are key figures set as display type with their labels directly below and aligned?
15. **Contrast:** does text on colour or on an image meet a contrast ratio of at least 4.5:1 for body and 3:1 for large type? Is there a scrim on photos?
16. **Images:** are images either bled or grid-framed, consistently cropped, not stretched and not pixelated?
17. **Running elements:** are the header, footer and page counter present (on content slides) and on the same baseline as on other slides?
18. **Hierarchy order:** does the eye path go hero, then supporting element, then detail, then meta, with nothing out of order?
19. **RTL (Hebrew only):** is the layout mirrored while numerals, logos and chart axes stay left to right? Is there no fake italic and no letter-spaced caps logic?
20. **Deck context:** does this slide differ in layout and density from its neighbours? Does the closing slide echo the cover?

