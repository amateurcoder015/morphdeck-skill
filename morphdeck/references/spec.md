# morphdeck spec reference

A spec is one JSON object:

```json
{
  "title": "Deck title (stored in file properties)",
  "theme": "midnight",
  "theme_overrides": {"accent": "FF3366"},
  "fonts": {"heading": "Unbounded", "body": "Unbounded"},
  "slides": [ { "layout": "title", "...": "..." } ]
}
```

Only `slides` is required. Every slide can also carry `"notes": "speaker notes"`.

## Layouts

| layout | fields | what it looks like / animates |
|---|---|---|
| `title` | `title`*, `kicker`, `subtitle`, `byline` | Big bottom-anchored title that builds letter by letter, kicker in tracked caps, subtitle rises in. Glow and dashed ring on the right. |
| `section` | `title`*, `number` (e.g. `"02"`), `subtitle` | Chapter divider. The number zooms in inside the ring on the left, and the title builds letter by letter on the right. |
| `statement` | `text`*, `attribution` | One big sentence on a floating panel, revealed word by word. Use for the core idea or a quote. |
| `bullets` | `title`*, `points`* (≤ 6, best 3–5), `aside` | Points rise in one after another with coloured dots. An optional `aside` (2–4 words) sits in the right-hand column inside the ring. |
| `stat` | `value`*, `suffix`, `label`*, `note`, `context` | Huge number (suffix in the accent colour) zooms in. `context` is a short line shown inside the ring. Example: `"value": "6.5", "suffix": "B"`. |
| `cards` | `title`*, `cards`* (2–4 × `{title, text}`) | Numbered cards rise in one after another. |
| `timeline` | `title`*, `steps`* (3–6 × `{label, text}`) | The accent bar morphs into a long axis, and nodes, labels and text build left to right. Labels can be years or one-word stages. |
| `compare` | `title`*, `left`*, `right`* (each `{title, points[]}`) | Two cards split by a vertical divider. Use for before/after, myth/reality, or A vs B. |
| `split` | `title`*, `kicker`, `text`, `image` or `aside` | Text on the left, a panel on the right. `image` is a path relative to the spec and is cropped to fill the panel. Without an image, `aside` is a short phrase shown in the ring. |
| `closing` | `title`*, `subtitle` | Centred final line that builds letter by letter inside the ring. |

`*` = required.

## How the motion works (for edits)

- Stage shapes on every slide: `!!glow_a`, `!!glow_b`, `!!panel`, `!!ring`,
  `!!disc`, `!!bar`, `!!progress`. Each layout has its own pose for each one, set in
  `Deck.stage()`. Morph tweens those poses. The ring also turns 70° per slide, so
  it never sits still.
- Content shapes have ordinary names (`title`, `point0`, `card1`, …). They are new
  on each slide, so Morph fades them, and their entrance animations play after the
  transition lands.
- Effects are "With Previous" with explicit delays, which is how PowerPoint stores
  a staggered chain. Open the Animation Pane to retime them.
- The first slide has no transition. Every later slide uses Morph by object, with
  `tempo` seconds of duration (1.2–1.6 s depending on theme). A fade fallback is
  included for apps without Morph.

## Adding a layout

1. Add a pose for it in the dictionary in `Deck.stage()`.
2. Add a `L_<name>(self, slide, s, a, st)` method that places content and calls
   `a.add(kind, shape, delay_ms, dur_ms)`. The kinds are `fade`, `rise`, `letters`,
   `zoom`, `wipe`, `float` and `sway`. `rise` accepts `by_word=True`.
3. Add the name to `LAYOUTS`.
