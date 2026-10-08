# Archetypes — fields and limits

Common optional fields on every slide: `kicker` (small accent label above the title), `section` (header label), `tone` (base | inverse | accent), `notes` (speaker notes), `overlays` (free-placed images), `chrome: false` (hide footer).

Word limits per density: **minimal / light / detailed**. "T" = title words, "I" = items, "w/i" = words per item.

| type | fields | limits (min / light / detailed) | notes |
|---|---|---|---|
| cover | title, subtitle?, kicker?, meta[]? (type variant: up to 3 lines opposite the subtitle, e.g. presenter, date, url), image?, variant type\|image\|split, dim? | T 2-6 · subtitle 0 / 12 / 20 | `\n` in title for a deliberate 2-line break. image variant = full-bleed photo with dim + gradient scrim behind the text. |
| statement | title, kicker?, text?, image?, dim?, width? | T 3-8 / 5-12 / 8-16 · text 0 / 20 / 35 | One sentence the room should remember. Photo optional. |
| section | title, number?, kicker?, text?, cutout?, cutout_x/y/w? (inches) | T 1-4 · text 0 / 15 / 25 | Giant numeral in accent; flips to inverse tone by default. A transparent cut-out can stand in front of the numeral. |
| agenda | title, text?, items[str \| {title, text?, meta?}] | I 3-5 / 4-7 / 4-8 · w/i 2-4 / 2-6 / 4-10 | meta = duration or time. |
| kpi | title, text?, metrics[{value, label}], highlight[] | 2-4 metrics · label 2-4 / 3-6 / 4-8 words | Values short ("41%", "18 days"). Highlight the one that matters. 4 metrics sit in one row (scoreboard), not a 2x2. |
| process | title, text?, steps[{title, text}], highlight[], note? | 3-5 steps · text 0 / 8-12 / 12-20 | Step titles = verbs. Base layout: big step numerals on a ruled line, the highlighted step in accent (default: first). |
| timeline | title, text?, events[{when, title, text?}], highlight[] | 3-8 events · ≤4 events gives big type, >4 zig-zags | Highlight the next milestone or the goal. ≤4 events: dates set as display numerals above the line, kept on one line (short `when`: "2024", "Today"). |
| comparison | title, text?, left{title, items[]}, right{title, items[]} | 3-6 items each · 3-8 words per item | Right side = the good side (filled panel). |
| cards | title, text?, cards[{title, text, n?}], highlight[] | 2-6 cards · text 0 / 10-15 / 15-25 | 3 or 6 look best; highlight one at most. One row (2-3 cards): tall tiles filling the slide, numeral on top, title + text on the bottom; cards get a soft shadow on light grounds. |
| image-text | title, image, kicker?, text?, bullets[]?, image_side left\|right, image_share? | text 0 / 20 / 40 · bullets ≤4 | Image full-height. Best for people, product, place. |
| quote | quote, name?, role?, image? \| cutout? | quote 6-25 words | Real quotes only; attribute honestly. |
| team | title, text?, people[{name, role, photo}] | 2-4 people | Same crop and light for all photos. |
| chart | title, text?, chart column\|bar\|line\|doughnut, categories[], series[{name, values[]}], highlight?, takeaway?, takeaway_label?, format?, labels?, clean? | ≤12 categories, ≤3 series | Native, editable chart. Bars start at 0. Single-series bar/column charts are `clean` by default (no axis/gridlines, big value labels); `clean: false` brings the axis back. `format` e.g. `0%`, `€#,##0`. Title = the insight. `takeaway` = a figure or 2-6 words (longer phrases drop to headline size and warn). |
| pricing | title, text?, plans[{name, price, unit?, period?, features[], cta?, badge?}], highlight[], points[]? | 1-3 plans · ≤6 features | One plan = argument on the left (title, text, up to 4 `points`), the offer on a full-height panel in the inverse ground on the right (display price, `cta` as a pill button). 2 plans = title column + two tall cards. 3 plans = full-width cards, the highlighted one raised; `badge` (e.g. "Most popular") adds a pill. Without any `cta`, cards hug their content and stand on the footer line. |
| gallery | title, text?, images[path \| {path, caption}], weights? | 2-3 images | weights e.g. [2,1] for a hero + side image. |
| closing | title, title2?, kicker?, contacts[[label, value]], note?, image?, variant split? | title 1-4 words each line | title2 renders in accent. One ask. With an image: default = full-bleed + dim; `variant: split` mirrors the split cover (photo right, ask left). |

### Long-text archetypes (use them for "detailed" decks, PDFs sent without a presenter, reports, proposals)
| type | fields | limits | notes |
|---|---|---|---|
| text | title, body (string with blank lines between paragraphs, or a list of paragraphs), kicker?, lead?, quote?, quote_by?, variant essay\|columns | body 80-220 words, 2-4 paragraphs | essay = title + lead on the side, one long column. columns = two columns; with a `quote` the text becomes one wide column and the quote sits beside it. |
| case | title, challenge, approach, result, metric?: {value, label}, *_label? (rename the three headings) | 30-70 words per part | Case study / pilot / story. Result column is marked in accent. |
| bullets | title, points: [{lead, text}], kicker?, text? | 3-6 points · text 12-35 words | Bold lead + full sentences. ≤3 points: columns headed by a display numeral; 4+ points: two columns of rows (`variant: list` forces rows). |
| faq | title, items: [{q, a}], kicker?, text? | 3-6 pairs · answers 15-45 words | Objections, board questions, onboarding questions. |

Body text in these never goes below 12pt; if a WARN appears, cut words or split the slide in two.

### Business structures (data, plans, proof)
| type | fields | limits | notes |
|---|---|---|---|
| table | title, columns[], rows[[cell]], kicker?, text?, highlight_col?, highlight_row?, total?, source?, takeaway?, takeaway_label?, wide? | 2-7 columns · 1-8 rows · cells 1-4 words | Native, editable PowerPoint table. ≤4 columns: the title column sits beside the table (with an optional `takeaway` figure); 5+ (or `wide`): full width under the title. Figures align on their last digit; Hebrew tables run right to left. Highlight one column OR one row. bold = header on the inverse band, highlighted column becomes an accent band; editorial = hairlines only; soft = rounded card, quiet zebra rows. `total: true` sets the last row as a total. |
| funnel | title, stages[{label, value, text?, conversion?}], highlight[], conversion? | 3-6 stages · label 2-5 words · text ≤8 words | Widest first. Numeric values set the tier widths (compressed so the last tier stays legible); step conversion between rows is computed unless given or `conversion: false`. The last stage takes the accent by default. |
| matrix | title, quadrants[4 × {title, text? \| items[], tag?}], x_axis?, y_axis?, highlight?, text?, variant base\|swot | quadrant title 2-6 words · text ≤20 words (or ≤4 items) | Order: top-left, top-right, bottom-left, bottom-right (mirrored in Hebrew). Axes are drawn as arrows with labels at the high end. `swot`: full-width 2x2 with giant S/W/O/T initials, default quadrant names localised; give `items`. |
| org | title, root{name, role, photo?}, reports[{name, role, photo?, children[{name, role}]}], highlight?, text? | 1-4 reports · ≤3 children each | Lead in an inverse card, reports on a bus line, teams on a spine. Initials stand in for missing photos. `highlight`: `[report]` or `[report, child]`. |
| roadmap | title, periods[], lanes[{name, text?, items[{title, from, to, highlight?}]}], now?, now_label? | 3-6 periods · 2-4 lanes · ≤6 items per lane, titles 1-4 words | `from`/`to` = period names (inclusive) or numbers (0-based start, end exclusive, fractions allowed). Overlapping items stack. Work that ended before `now` goes quiet; the `now` line + pill cut through. One `highlight` item. |
| device | title, image (screenshot) \| video, kicker?, text?, bullets[]?, variant laptop\|phone, image_side | text ≤30 words · bullets ≤4 | A code-drawn laptop or phone; the screen takes the image's own shape (never cropped). bold = device on an accent field off the edge; editorial = quiet panel; soft = device + shadow. Phone wants a portrait screenshot (about 9:19). |
| testimonials | title, quotes[{quote, name, role?, photo?}], featured?, kicker? | 2-4 quotes · 10-30 words each | Default: the first quote featured on a full-height panel, the rest stacked beside it. `featured: null` = equal columns. Real quotes only. |
| logos | title, logos[path], kicker?, text?, stat?{value, unit?, label}, mono? | 4-12 logos | Logos are trimmed and sized by optical weight, in an even grid. `mono` (default) recolours them to one ink so the wall reads as one surface (`false` keeps colours, or pass a hex). PNG with transparency preferred. |
| beforeafter | title, before{image \| title, text?, items[]?, caption?}, after{same}, labels?, caption? | items ≤5 per side | Two images meet on a seam with a slider handle and corner labels (Before/After, לפני/אחרי). Without images: the old state quiet and narrow, the new one on the inverse ground, an arrow on the seam. Photos are not mirrored here. |
| pyramid | title, tiers[{title, text?, n?}], highlight[] | 3-5 tiers · text ≤12 words | Apex first. One silhouette cut into tiers, each tier's story on a leader beside it; the apex takes the accent by default. |
| video | title, video, poster?, kicker?, text?, dim?, variant full\|framed, image_side?, meta? | title 2-8 words | `full`: the clip loops full-bleed behind dim + scrim, the line set large. `framed`: the clip in a frame, caption beside it with the running time. See SKILL.md §5 Video. |

`video` also works on cover (image variant), statement, closing, image-text, quote (bleed) and device: it takes the slide's main image slot.

## Choosing between similar archetypes
- Numbers that stand alone → **kpi**; numbers over time or by category → **chart**.
- Ordered steps → **process**; dated milestones → **timeline**; equal options/features → **cards**.
- Before/after, us/them, old/new → **comparison**.
- A person's words → **quote**; several customers → **testimonials**; your own claim → **statement**.
- Exact figures to compare across rows → **table**; one figure that shrinks stage by stage → **funnel**.
- Options judged on two criteria → **matrix**; strengths/weaknesses/opportunities/threats → **matrix** `swot`.
- Before/after with two photos or two states → **beforeafter**; a list of old vs new points → **comparison**.
- Dated milestones in one line → **timeline**; several workstreams over periods → **roadmap**.
- Layers that build on each other → **pyramid**; who reports to whom → **org**.
- A product screen → **device** (or image-text with `fit: contain`); client names → **logos**.

