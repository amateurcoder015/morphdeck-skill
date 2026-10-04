# morphdeck spec reference

A spec is one JSON object:

```json
{
  "title": "Deck title (stored in file properties)",
  "theme": "midnight",
  "options": {"slides": 10, "motion": "normal", "aspect": "16:9", "tone": "exec"},
  "theme_overrides": {"accent": "FF3366"},
  "fonts": {"heading": "Unbounded", "body": "Unbounded"},
  "slides": [ { "layout": "title", "...": "..." } ]
}
```

Only `slides` is required. Every slide can also carry `"notes": "speaker notes"`.

## Options

| key | values | effect |
|---|---|---|
| `slides` | number | The target count. The build warns if the spec has a different number of slides. |
| `motion` | `calm` / `normal` / `dramatic` | Morph speed, delay between items, rise distance, zoom depth, letter-by-letter titles (off in `calm`), ring spin per slide, sway and float size, Ken Burns depth. |
| `aspect` | `16:9` / `16:10` / `4:3` | Slide size. Layouts reflow, and round shapes stay round. |
| `tone` | free text | Guides the writing only; the script ignores it. |

## Layouts

`*` = required. Image fields accept a local path, a URL or `stock:<search words>`.

| layout | fields | what it looks like / animates |
|---|---|---|
| `title` | `title`*, `kicker`, `subtitle`, `byline` | Big title built letter by letter, kicker in tracked caps, subtitle rises in. |
| `agenda` | `title`, `items` (≤ 7) | Numbered list on the right of a divider. If `items` is left out, it lists the deck's `section` titles. |
| `section` | `title`*, `number`, `subtitle` | Chapter divider. The number zooms in inside the ring. |
| `statement` | `text`*, `attribution` | One big sentence on a floating panel, revealed word by word. |
| `question` | `text`*, `kicker`, `subtitle` | Centred question inside the ring, revealed word by word. A pause slide. |
| `quote` | `text`*, `author`*, `role`, `image` | Large quote mark and quote text. The optional photo is cropped to a circle and framed by the ring. |
| `bullets` | `title`*, `points`* (≤ 6), `aside` or `image` | Points rise in with coloured dots. The right-hand column shows a short `aside` phrase or a photo. |
| `stat` | `value`*, `suffix`, `label`*, `note`, `context` | Huge number zooms in. `context` sits inside the ring. |
| `chart` | `title`*, `type`*, `categories`*, `series`*, `takeaway`, `caption`, `source`, `number_format`, `axis_format` | Native, editable PowerPoint chart in theme colours. Columns wipe up, bars and lines wipe across, and donut and pie charts wheel in. `takeaway` is a big 2–4 word conclusion on the side panel. |
| `cards` | `title`*, `cards`* (2–4 × `{title, text}`) | Numbered cards rise in one after another. |
| `people` | `title`*, `people`* (2–4 × `{name, role, text, image}`) | Circular photos, or initials when no photo is given, with name, role and a line of text. |
| `timeline` | `title`*, `steps`* (3–6 × `{label, text}`) | The accent bar morphs into the axis, and nodes build left to right. |
| `process` | `title`*, `steps`* (3–5 × `{label, text}`), `center` | Numbered nodes orbit the ring, matched to steps listed on the left. `center` is a word or two shown in the middle. Use for cycles and flywheels. |
| `compare` | `title`*, `left`*, `right`* (each `{title, points[]}`) | Two cards with a vertical divider. |
| `split` | `title`*, `kicker`, `text`, `image` or `aside` | Text on the left, a panel or photo on the right. Photos get a Ken Burns push-in. |
| `image` | `image`*, `title`*, `kicker`, `text` | Full-bleed photo with a theme-coloured gradient scrim, text over it, and a slow Ken Burns zoom. |
| `gallery` | `title`*, `images`* (2–4 × path or `{image, caption}`) | Rounded photo tiles that zoom in one after another, with captions. |
| `closing` | `title`*, `subtitle` | Centred final line, built letter by letter. |

### Chart details

- `type`: `column`, `bar`, `line`, `area`, `stacked`, `donut` or `pie`.
- `series`: `[{"name": "Revenue", "values": [1, 2, 3]}]`. Use several series for comparisons;
  a legend is added automatically. Donut and pie charts use the first series only.
- `number_format`: Excel format string for data labels, e.g. `"0.0"`, `"0\"%\""` or `"$#,##0"`.
  `axis_format` does the same for the value axis.
- A single-series column or bar chart shows value labels.
- The chart's data can be edited in PowerPoint (right-click, then Edit Data).

## How the motion works (for edits)

- Stage shapes on every slide: `!!glow_a`, `!!glow_b`, `!!panel`, `!!ring`,
  `!!disc`, `!!bar`, `!!progress`. Each layout has a pose for each one in `POSES`.
  Morph tweens those poses. The ring also turns by the motion profile's `spin` on every slide.
- Content shapes have ordinary names (`title`, `point0`, `card1`, `chart`, …). They are
  new on each slide, so Morph fades them, and their entrance animations play after
  the transition lands.
- Effects are "With Previous" with explicit delays. Open the Animation Pane to retime them.
- Motion profiles live in `MOTION` at the top of `build_deck.py`.

## Adding a layout

1. Add its stage pose to `POSES`, in design inches on a 13.333 × 7.5 canvas.
2. Add a `L_<name>(self, slide, s, a, st)` method. Place shapes with `self.shape`, `self.text`,
   `self.card` and `self.picture`, all in design inches. Call `a.add(kind, shape, delay_ms, dur_ms)`,
   where kind is one of `fade`, `rise`, `letters`, `zoom`, `wipe` (`dir=` `right`, `up`, `left` or `down`),
   `wheel` or `kenburns`. `rise` accepts `by_word=True`.
3. If the layout uses images, add it to `Deck.image_jobs()` so the images are fetched in parallel.
