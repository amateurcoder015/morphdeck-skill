---
name: morphdeck
description: Turn a topic title (optionally with a short brief) into a cinematic, animated PowerPoint .pptx that uses Morph transitions, staggered entrance animations, letter-by-letter titles and looping ambient motion, so the deck plays like one continuous film. Use when the user asks for a ppt, pptx, PowerPoint, slide deck or presentation about a topic and wants it to look modern, animated, fluid, cinematic or "not basic", or when they invoke /morphdeck.
---

# morphdeck

Builds a native `.pptx` whose motion lives inside PowerPoint itself:

- **Morph by object** between every slide. A persistent "stage" of shapes named
  with the `!!` prefix (glows, a dashed ring, a disc, a panel, an accent bar and a
  progress line) changes position, size and rotation on each layout. Morph glides
  them between those poses, so the deck feels like one camera move.
- **Entrance animations** on each slide's content: rise, fade, zoom, wipe, and titles
  that build letter by letter. They start automatically after the transition and
  are staggered with delays, so you never need extra clicks.
- **Ambient loops**: the disc floats up and down and the ring sways, with smooth
  start and end and auto-reverse.

Everything stays editable in PowerPoint's Animation Pane and Selection Pane.

## Workflow

### 1. Read the input
The user gives a **topic title** and sometimes a **brief**. If there is no brief,
write the content from your own knowledge. Keep facts accurate and do not invent
statistics. If a number is uncertain, use a layout that doesn't need one. If the
user mentions an audience or slide count, follow it. Otherwise aim for **8–12 slides**.

### 2. Pick a theme
Match the topic's mood. Run `python3 scripts/build_deck.py --list-themes` to see them:

| theme | look | use for |
|---|---|---|
| `midnight` | navy, violet and cyan glow | tech, AI, software, space, future |
| `emerald` | deep green, lime | finance, sustainability, nature, growth |
| `ember` | warm black, orange and amber | history, energy, sport, culture, bold stories |
| `aurum` | black and gold | luxury, leadership, law, premium, banking |
| `paper` | warm off-white, red and yellow | business, education, research, policy |
| `lagoon` | pale teal, coral | health, science, wellbeing, travel |

When the user names a brand colour, add `"theme_overrides": {"accent": "HEX"}`.

### 3. Write the storyline
Write it as a short film, not as a report:
- **Slide 1 `title`**: hook kicker, title, one-line promise.
- **Slide 2**: usually a `statement`, the big idea in one sentence.
- **Middle**: mix `section`, `bullets`, `stat`, `cards`, `timeline`, `compare`
  and `split`. **Never use the same layout twice in a row.** Morph looks best when
  consecutive layouts differ, because the stage moves further. Use a `section`
  divider every 3–4 slides in decks of 10 or more slides.
- **Last slide `closing`**: a memorable line plus a thank-you or call to action.
- Put talking points in `notes` (speaker notes), not on the slide.

Text budgets (the generator shrinks fonts to fit, but less text looks better):
titles ≤ 8 words, bullet points ≤ 12 words, up to 5 points, card text ≤ 18 words,
statement ≤ 25 words, timeline step text ≤ 14 words.

Full field reference for every layout: [references/spec.md](references/spec.md).
Worked example: [examples/black-holes.json](examples/black-holes.json).

### 4. Build
Write the spec to `<slug>.json` in the user's working directory, then:

```bash
python3 ~/.claude/skills/morphdeck/scripts/build_deck.py <slug>.json <slug>.pptx
```

Needs `python-pptx`. If the import fails, run `python3 -m pip install --user python-pptx`.

### 5. Check it
On macOS with PowerPoint installed, render static previews and **look at the contact sheet**:

```bash
python3 ~/.claude/skills/morphdeck/scripts/preview.py <slug>.pptx <scratch-dir>/preview
```

(Needs `pymupdf` and `pillow`. Add `--restart` if a font was just installed.)
Check for text that overflows its box, words that look cramped, or text that
collides with stage shapes. Fix problems by shortening the text in the spec,
then rebuild. Previews show final frames only. Morph and animation timing can
only be seen in PowerPoint's slideshow. Without PowerPoint, skip this step and
say so.

### 6. Deliver
Give the `.pptx` path. Tell the user:
- Play it as a slideshow in **PowerPoint 2019+ / Microsoft 365**. Keynote and Google
  Slides replace Morph with a fade.
- The **Unbounded** font must be installed (free on Google Fonts). Otherwise
  PowerPoint substitutes another font and the spacing changes.
- To restyle, edit the JSON and rebuild, or edit directly in PowerPoint. Keep the
  `!!` names on stage shapes if they duplicate slides, or Morph stops matching them.

## Tuning knobs
- `"fonts": {"heading": "...", "body": "..."}` swaps the fonts (default Unbounded for both).
- `theme_overrides` can set any theme key: `bg`, `surface`, `text`, `muted`,
  `accent`, `accent2`, `glow` (glow strength 0–1) and `tempo` (Morph duration in seconds).
- `--theme NAME` on the command line overrides the spec's theme. This is useful
  for showing the user two looks.
