# morphdeck: cinematic PowerPoint decks from a topic

A [Claude Code](https://claude.com/claude-code) skill. Give it a topic title, and optionally a short brief. It writes the storyline, picks a theme that fits the topic, and builds a **native `.pptx`** that moves like a film:

- **Morph transitions** on every slide. A shared stage of glows, a dashed ring, panels and accent bars glides, resizes and rotates between slides, so the deck plays as one continuous camera move.
- **Staggered entrance animations**: titles build letter by letter, statements word by word, and points, cards and timeline steps rise in one after another. They all start automatically, with no extra clicks.
- **Ambient motion loops**: a floating accent disc and a swaying ring, with smooth start and end and auto-reverse.
- **6 themes**, chosen to match the topic: `midnight`, `emerald`, `ember`, `aurum`, `paper` and `lagoon`.
- **10 layouts**: title, section, statement, bullets, stat, cards, timeline, compare, split (with optional image) and closing.

Everything stays editable in PowerPoint's Animation Pane and Selection Pane.

## ⚠️ Install the Unbounded font first

The decks use **[Unbounded](https://fonts.google.com/specimen/Unbounded)**, a free font from Google Fonts. Install it on **every machine that will open or present the deck**. On a Mac, double-click the `.ttf` file and choose *Install*. On Windows, right-click it and choose *Install for all users*. Then restart PowerPoint.

Without the font, PowerPoint substitutes another typeface. The slides still work, but the type looks generic and the spacing changes.

## Requirements

- **PowerPoint 2019, 2021 or Microsoft 365** (Mac or Windows) to present the deck. Morph is a PowerPoint feature. Keynote, Google Slides and older PowerPoint versions fall back to a plain fade.
- Python 3.9 or newer with `python-pptx`:
  ```bash
  python3 -m pip install --user python-pptx
  ```
- Optional, for previews on a Mac with PowerPoint installed: `python3 -m pip install --user pymupdf pillow`

## Install the skill

```bash
git clone https://github.com/amateurcoder015/claude-skill2.git
mkdir -p ~/.claude/skills
cp -r claude-skill2/morphdeck ~/.claude/skills/morphdeck
```

Restart Claude Code.

## Use it

```
/morphdeck Black holes
```

```
/morphdeck The 2008 financial crisis — for a college econ class, 10 slides,
focus on causes, the Lehman collapse, and what changed after.
```

Or ask in plain words, for example: "make me an animated ppt about photosynthesis". Claude writes a JSON spec, builds `<topic>.pptx` in your current folder, checks the layout, and tells you where the file is. Open it in PowerPoint and start the slideshow (**⌘⇧↩** on Mac, **F5** on Windows) to see the motion.

## Run the generator directly

```bash
python3 morphdeck/scripts/build_deck.py morphdeck/examples/black-holes.json black-holes.pptx
python3 morphdeck/scripts/build_deck.py morphdeck/examples/black-holes.json black-holes.pptx --theme ember
python3 morphdeck/scripts/build_deck.py --list-themes
```

The spec format for every layout is in [`morphdeck/references/spec.md`](morphdeck/references/spec.md).

## Layout

```
morphdeck/
├── SKILL.md               instructions Claude follows
├── references/spec.md     JSON spec + layout reference
├── examples/black-holes.json
└── scripts/
    ├── build_deck.py      spec → .pptx (Morph + animation XML)
    ├── themes.json        the 6 themes
    └── preview.py         .pptx → PNG contact sheet via PowerPoint (macOS)
```

## Limits

- Previews show each slide's final frame. To judge Morph and the animation timing, play the deck in PowerPoint.
- Text is sized to fit using estimates. Very long text shrinks to small sizes, so keep slides short.
- The generator does not create images. You can add your own to `split` slides with `"image": "photo.jpg"`.
