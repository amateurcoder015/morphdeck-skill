---
name: morphdeck
description: Turn a topic title (optionally with a short brief) into a cinematic, animated PowerPoint .pptx that uses Morph transitions, staggered entrance animations, letter-by-letter titles, animated native charts, stock photos with Ken Burns zoom and looping ambient motion, so the deck plays like one continuous film. Supports slide count, motion intensity (calm/normal/dramatic), aspect ratio (16:9, 16:10, 4:3), tone/audience and six topic-matched themes. Use when the user asks for a ppt, pptx, PowerPoint, slide deck or presentation about a topic and wants it to look modern, animated, fluid, cinematic or "not basic", or when they invoke /morphdeck.
---

# morphdeck

Builds a native `.pptx` whose motion lives inside PowerPoint itself:

- **Morph by object** between every slide. A persistent "stage" of shapes named
  with the `!!` prefix (glows, a dashed ring, a disc, a panel, an accent bar and a
  progress line) changes position, size and rotation on each layout. Morph glides
  them between those poses, so the deck feels like one camera move.
- **Entrance animations** on each slide's content: rise, fade, zoom, wipe, wheel,
  and titles that build letter by letter or word by word. They start automatically
  after the transition and are staggered with delays.
- **Ambient loops**: the disc floats, the ring sways, and photos slowly push in
  (Ken Burns).

Everything stays editable in PowerPoint's Animation Pane and Selection Pane.

## Workflow

### 1. Read the input and the options
The user gives a **topic title** and sometimes a **brief**. They may also state options
in plain words or as `key=value` pairs, for example
`/morphdeck Black holes slides=8 motion=dramatic theme=midnight aspect=4:3 tone=kids images=yes`.

| option | values | default when not stated |
|---|---|---|
| `slides` | a number | 8–12, sized to how much the topic needs |
| `motion` | `calm`, `normal`, `dramatic` | `normal`. Use `calm` for corporate, medical or serious topics, and `dramatic` for pitches, launches, sport and storytelling |
| `theme` | see step 2 | picked from the topic |
| `aspect` | `16:9`, `16:10`, `4:3` | `16:9` |
| `tone` | free text: `exec`, `academic`, `kids`, `casual`, `persuasive`, … | inferred from the topic and audience |
| `images` | `yes`, `no`, or local paths and URLs the user provides | `yes` when photos would help the story, `no` for abstract topics |
| `brand` | hex colours | theme colours |

Hit the slide count exactly when the user gives one. Ask a question only when
something blocks you. Otherwise choose defaults and list them in the delivery message.

If there is no brief, write the content from your own knowledge. Keep facts
accurate. **Never invent statistics, chart data or quotes.** Use a chart or
`stat` only for figures you are confident of, and put the source in `source`
or `note`. Attribute a `quote` only to someone who really said it; otherwise
use `statement`.

### 2. Pick a theme
Run `python3 scripts/build_deck.py --list-themes` to see them:

| theme | look | use for |
|---|---|---|
| `midnight` | navy, violet and cyan glow | tech, AI, software, space, future |
| `emerald` | deep green, lime | finance, sustainability, nature, growth |
| `ember` | warm black, orange and amber | history, energy, sport, culture, bold stories |
| `aurum` | black and gold | luxury, leadership, law, premium, banking |
| `paper` | warm off-white, red and yellow | business, education, research, policy |
| `lagoon` | pale teal, coral | health, science, wellbeing, travel |

When the user gives brand colours, add `"theme_overrides": {"accent": "HEX", "accent2": "HEX"}`.

### 3. Write the storyline for the tone
Write it as a short film:
- **Slide 1 `title`**: hook kicker, title, one-line promise.
- **Slide 2**: a `statement` with the big idea, or an `agenda` for decks of 10 or more slides.
- **Middle**: mix the content layouts. **Never use the same layout twice in a row**,
  because the stage moves furthest between different layouts. Use a `section` every
  3–4 slides in long decks. A `question` makes a good pause before a turning point.
- **Last slide `closing`**.
- Put talking points in `notes` (speaker notes), written in the chosen tone.

How each tone changes the writing:
- **exec**: lead with the conclusion, use numbers, `stat` and `chart` slides, short noun phrases.
- **academic**: definitions, sources in `note`/`source`, `timeline` and `compare` slides.
- **kids**: simple words, one idea per slide, questions, `calm` or `normal` motion, bright themes.
- **persuasive** or pitch: problem → stakes → solution → proof → call to action, `dramatic` motion.

Text budgets (fonts shrink to fit, but less text looks better): titles ≤ 8 words,
up to 5 bullet points of ≤ 12 words each, card text ≤ 18 words, statement ≤ 25 words,
timeline and process step text ≤ 14 words.

Layout fields: [references/spec.md](references/spec.md). Examples:
[examples/black-holes.json](examples/black-holes.json) and
[examples/electric-vehicles.json](examples/electric-vehicles.json), which uses
charts, stock images, agenda, process and question slides.

### 4. Images
Image fields (`image` on `image`, `split`, `bullets` and `quote` slides, people
photos, and `images` on `gallery`) accept:
- `"photos/x.jpg"`: a local file, relative to the spec. Use this for anything the user supplies.
- `"https://…"`: downloaded once.
- `"stock:<search words>"`: a stock photo. The search uses Pexels if `PEXELS_API_KEY`
  is set, otherwise Openverse (Creative Commons, no key needed).

Write concrete, visual search words, such as `stock:wind turbines at sunset` rather than
`stock:renewable energy policy`. Credits are added to that slide's speaker notes
automatically. Downloads are cached in `images/` next to the spec. A new search takes
about 30–120 s; the build prints a warning and carries on if nothing is found.
Never use stock photos to show specific real people. Use initials (the default in
`people`) or photos the user supplies.

### 5. Build
Write the spec to `<slug>.json` in the user's working directory, with
`"options": {"slides": N, "motion": "...", "aspect": "...", "tone": "..."}`. Then run:

```bash
python3 ~/.claude/skills/morphdeck/scripts/build_deck.py <slug>.json <slug>.pptx
```

`--theme`, `--motion` and `--aspect` on the command line override the spec. This is
useful for showing the user two looks. Needs `python-pptx`; if the import fails,
run `python3 -m pip install --user python-pptx`.

### 6. Check it
On macOS with PowerPoint installed, render previews and **look at the contact sheet**:

```bash
python3 ~/.claude/skills/morphdeck/scripts/preview.py <slug>.pptx <scratch-dir>/preview
```

(Needs `pymupdf` and `pillow`. Add `--restart` if a font was just installed.)
Fix text that overflows, words that look cramped, text that collides with stage
shapes, and badly cropped photos by editing the spec and rebuilding. Previews show
final frames only, so Morph and timing can only be judged in PowerPoint. Without
PowerPoint, skip this step and say so.

### 7. Deliver
Give the `.pptx` path and the options you chose (theme, motion, aspect, slide
count, tone). Tell the user:
- Play it as a slideshow in **PowerPoint 2019+ / Microsoft 365**. Keynote and Google
  Slides replace Morph with a fade.
- The **Unbounded** font must be installed (free on Google Fonts).
- Stock photo credits are in the speaker notes. Keep them if the deck is shared publicly.
