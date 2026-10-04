# morphdeck: cinematic PowerPoint decks from a topic

A [Claude Code](https://claude.com/claude-code) skill. Give it a topic title, and optionally a short brief. It writes the storyline, picks a theme that fits the topic, and builds a **native `.pptx`** that moves like a film:

- **Morph transitions** on every slide. A shared stage of glows, a dashed ring, panels and accent bars glides, resizes and rotates between slides, so the deck plays as one continuous camera move.
- **Staggered entrance animations**: titles build letter by letter, statements word by word, and points, cards and timeline steps rise in one after another. They all start automatically, with no extra clicks.
- **Ambient motion loops**: a floating accent disc, a swaying ring, and a slow Ken Burns push-in on photos.
- **Native animated charts**: column, bar, line, area, stacked, donut and pie charts in theme colours. The data stays editable in PowerPoint.
- **Images**: your own files, URLs, `stock:<search>` photos fetched automatically (Openverse, or Pexels with a free key), or `ai:<prompt>` images generated free with Cloudflare Workers AI (FLUX.1 schnell). Credits go into the speaker notes.
- **Clickable agenda**: agenda items jump to their sections, and each section links back, with Morph on every jump.
- **6 themes × 4 motifs**: colour themes (`midnight`, `emerald`, `ember`, `aurum`, `paper`, `lagoon`) combine with shape styles for the morphing stage: `orbit` (glowing orbs and rings), `prism` (Bauhaus squares and triangles), `swiss` (flat editorial blocks) and `flow` (ribbons and arcs).
- **19 layouts**: title, agenda, section, statement, question, quote, bullets, detail, stat, chart, cards, people, timeline, process, compare, split, image (full-bleed), gallery and closing.

## Customise it

| option | values | default |
|---|---|---|
| `slides` | any number | 8–12, depending on the topic |
| `motion` | `calm`, `normal`, `dramatic` | `normal` |
| `theme` | `midnight`, `emerald`, `ember`, `aurum`, `paper`, `lagoon` | picked from the topic |
| `motif` | `orbit`, `prism`, `swiss`, `flow` | the theme's motif |
| `aspect` | `16:9`, `16:10`, `4:3` | `16:9` |
| `tone` | `exec`, `academic`, `kids`, `casual`, `persuasive`, or anything else | inferred from the topic |
| `density` | `text` (content-heavy), `balanced`, `visual` (image-heavy) | `balanced` |
| `images` | `stock`, `ai`, `mixed`, `none`, or your own files and links | `stock` when photos help the story |
| `agenda` | `yes`, `no` (a clickable agenda) | `yes` for longer decks |
| `brand` | hex colours | theme colours |

`motion` changes Morph speed, the delay between items, how far things travel, zoom depth, letter-by-letter titles (off in `calm`), and how much the ring spins, sways and floats.

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
- Optional, for Pexels stock photos instead of Openverse: a free API key from [pexels.com/api](https://www.pexels.com/api/), set as `export PEXELS_API_KEY=...`
- Optional, for AI images: a free Cloudflare account (see below).

## Free AI images with Cloudflare

The `ai:` images use Cloudflare Workers AI's **FLUX.1 schnell** model. The free plan includes 10,000 neurons a day, and one image costs about 58, so you get **around 170 free images a day**. The allowance resets at 00:00 UTC.

1. Sign up at [dash.cloudflare.com](https://dash.cloudflare.com).
2. **Account ID**: on the dashboard, open *Workers & Pages* and copy the Account ID from the right sidebar.
3. **API token**: go to *My Profile → API Tokens → Create Token*, use the **Workers AI** template, and create the token.
4. Save both in `~/.config/morphdeck/.env`:
   ```
   CF_ACCOUNT_ID=your_account_id
   CF_API_TOKEN=your_token
   ```
5. Test it: `python3 morphdeck/scripts/images.py "ai:a lighthouse at dusk, cinematic photo" /tmp/test`

Without keys, `ai:` images fall back to a stock photo search, so decks still build. FLUX makes square 1024×1024 images, which the deck crops to each frame.

## Install the skill

```bash
git clone https://github.com/amateurcoder015/morphdeck-skill.git
mkdir -p ~/.claude/skills
cp -r morphdeck-skill/morphdeck ~/.claude/skills/morphdeck
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

```
/morphdeck Our Q3 results slides=8 motion=calm tone=exec theme=paper aspect=4:3 brand=#0052FF
```

Or ask in plain words, for example: "make me an animated ppt about photosynthesis". Claude writes a JSON spec, builds `<topic>.pptx` in your current folder, checks the layout, and tells you where the file is. Open it in PowerPoint and start the slideshow (**⌘⇧↩** on Mac, **F5** on Windows) to see the motion.

## Run the generator directly

```bash
python3 morphdeck/scripts/build_deck.py morphdeck/examples/black-holes.json black-holes.pptx
python3 morphdeck/scripts/build_deck.py morphdeck/examples/black-holes.json black-holes.pptx --theme ember --motion dramatic
python3 morphdeck/scripts/build_deck.py morphdeck/examples/electric-vehicles.json ev.pptx   # charts + stock photos
python3 morphdeck/scripts/images.py "wind turbines at sunset" /tmp/img                      # test a stock search
python3 morphdeck/scripts/images.py "ai:wind turbines at sunset, cinematic" /tmp/img       # test AI generation
python3 morphdeck/scripts/build_deck.py --list-themes
```

The spec format for every layout is in [`morphdeck/references/spec.md`](morphdeck/references/spec.md).

## Layout

```
morphdeck/
├── SKILL.md               instructions Claude follows
├── references/spec.md     JSON spec + layout reference
├── examples/              black-holes.json, electric-vehicles.json
└── scripts/
    ├── build_deck.py      spec → .pptx (Morph, animation XML, charts)
    ├── images.py          local / URL / stock / Cloudflare-AI image resolver with credits
    ├── themes.json        the 6 themes
    └── preview.py         .pptx → PNG contact sheet via PowerPoint (macOS)
```

## Limits

- Previews show each slide's final frame. To judge Morph and the animation timing, play the deck in PowerPoint.
- Text is sized to fit using estimates. Very long text shrinks to small sizes, so keep slides short.
- Stock photos come from Creative Commons and Pexels libraries. Quality varies, so check the preview and change the search words if a photo misses. The first search takes 30–120 s; results are cached in `images/` next to the spec.
- Vertical 9:16 decks are not supported yet. The layouts are built for landscape.
