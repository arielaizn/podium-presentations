# Animated HTML contract

Build the HTML from the same narrative and data as the PPTX. Preserve all material claims, slide order and notes where relevant; adapt the responsive layout rather than screenshotting slides. Text must remain selectable and screen-reader accessible.

Use one section per slide, with class `slide`, an accessible label, and `data-transition="fade|slide|zoom"`. Inside each, assign class `reveal` to text blocks and visuals and optional CSS custom property `--step` (0,1,2...) for staggered entrances. Include a single heading appropriate for the slide, semantic lists/tables, useful image alt text, and visible sources where necessary.

Add controls with IDs `prev-slide`, `next-slide`, optional `fullscreen`, and a status element `slide-status` with role=status and aria-live=polite. Include the bundled runtime CSS and deferred JavaScript using local relative paths. Customize colors, typography, layouts, illustrations and motion for the deck. The starter runtime deliberately does not choose slide content or one generic layout for every slide.

Set `<html lang="he" dir="rtl">` for Hebrew, or en/ltr for English. Wrap Latin tokens, currency figures and product names in `<bdi dir="ltr">` where needed. In RTL, the runtime uses left arrow to advance and right arrow to go back; controls and progress should communicate the same reading direction. Touch swipes follow that direction.

Use CSS logical properties for margins and positioning. Have no text overflow at desktop, narrow mobile and landscape mobile sizes. Dense mobile slides may scroll internally. Avoid excessive rotation, flashes or animations longer than about 700 ms. The runtime disables transitions and entrances for prefers-reduced-motion and printing.

Browser QA checklist:
- One slide initially active; navigation changes the visible slide and hash.
- Real text visibly enters, transitions animate, and repeated visits still work.
- Prev/next, keyboard, touch, Home/End and direct hashes behave correctly.
- No focus trap or offscreen active controls; hidden slides are not exposed to accessibility navigation.
- Sources and all media load. No unapproved third-party scripts, trackers or network uploads.
- All slides print as static pages; reduced-motion removes animations.
- Hebrew alignment, glyphs and mixed-direction values remain correct.
- Verify desktop and mobile screenshots after the final change, and check console errors.
