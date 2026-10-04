# md2cv — how to use

Short and practical. For the project overview see [README.md](README.md).

## Basic

```bash
md2cv cv.md                       # -> cv.docx next to it
md2cv cv.md -o Name_CV.docx       # custom output name
md2cv *.md -d out/                # all inputs into out/
```

## Presets

Ready-made bundles of layout settings.

```bash
md2cv cv.md --preset compact       # Arial Narrow 9.2 — fits two stubborn pages
md2cv cv.md --preset comfortable   # 10 pt, more air — for a one-pager
md2cv cv.md --preset classic       # sober Word-blue accent, no rules
md2cv --list-presets               # show what's inside each preset
```

An explicit flag always wins over the preset, in any order:

```bash
md2cv cv.md --preset compact --accent '#8B0000'
```

## Manual tuning

```bash
# squeeze
md2cv cv.md --font "Arial Narrow" --body-size 9.2 --margin-top 1.0

# breathe
md2cv cv.md --body-size 10 --line-spacing 1.1 --margin-left 2 --margin-right 2

# recolor
md2cv cv.md --accent '#2F5496' --rule-size 0                              # blue, no rules
md2cv cv.md --accent '#000000' --muted '#000000' --rule-color '#000000'   # B/W for ATS
```

## Help

```bash
md2cv --help          # all flags with live defaults, plus examples
md2cv --list-presets  # preset contents
md2cv --version
```

## Flags at a glance

| What                       | Flag                                             | Default                    |
| -------------------------- | ------------------------------------------------ | -------------------------- |
| Font family                | `--font`                                         | `Calibri`                  |
| Body text size             | `--body-size`                                    | `9.5` pt                   |
| Name size                  | `--name-size`                                    | `20` pt                    |
| Tagline size               | `--tagline-size`                                 | `10.5` pt                  |
| Contacts and dates size    | `--contact-size`                                 | `9` pt                     |
| Section heading size       | `--section-size`                                 | `11` pt                    |
| Company size               | `--company-size`                                 | `11.5` pt                  |
| Role title size            | `--role-size`                                    | `10` pt                    |
| Project heading size       | `--project-size`                                 | `10.5` pt                  |
| Line height                | `--line-spacing`                                 | `1.0`                      |
| Page margins               | `--margin-top/bottom/left/right`                 | `1.1 / 1.1 / 1.5 / 1.5` cm |
| Heading color              | `--accent`                                       | `#1F3864`                  |
| Dates / contacts color     | `--muted`                                        | `#595959`                  |
| Rule color                 | `--rule-color`                                   | `#9AA7B4`                  |
| Rule thickness             | `--rule-size`                                    | `6` (⅛ pt); `0` disables   |
| Section spacing            | `--section-space-before / --section-space-after` | `9 / 2.5` pt               |
| Company spacing            | `--company-space-before`                         | `8` pt                     |
| Bullet spacing             | `--item-space-after`                             | `1` pt                     |
| Paragraph spacing          | `--para-space-after`                             | `1.5` pt                   |
| Bullet indent              | `--list-indent`                                  | `0.5` cm                   |
| Output path (single input) | `-o, --output`                                   | next to input              |
| Output directory (batch)   | `-d, --output-dir`                               | next to input              |
| Preset                     | `--preset`                                       | `default`                  |

`--help` always prints the current defaults, so trust it over this table if
they ever diverge.

## PDF

```bash
libreoffice --headless --convert-to pdf --outdir . cv.docx
```

Batch:

```bash
md2cv ~/cv/*.md -d ~/cv/docx
libreoffice --headless --convert-to pdf --outdir ~/cv/pdf ~/cv/docx/*.docx
```

LibreOffice substitutes `Calibri`/`Arial Narrow` with metric-compatible
`Carlito`/`Liberation Sans Narrow`; pagination usually holds. Install the
originals if you need exact fidelity.

## Supported Markdown shape

This is the shape `md2cv` was designed for. Anything outside it becomes a
plain paragraph: readable, but not specially typeset.

```markdown
# Full Name

**Tagline** — one line

City · phone · email

## Section Title

### Company Name

**Role Title** · Jan 2020 – Present

_Optional italic blurb._

_Stack:_ Node.js, PostgreSQL

- Bullet

#### Project: Sub-project

**Project-level Tech Lead** · Jan 2020 – Dec 2021

- Bullet under project
```

The `·` in the role line must be `U+00B7 MIDDLE DOT`, not a hyphen. Bullets
must start with `- ` (hyphen and space).

### Not recognised

These constructs are tolerated (they don't crash the converter) but they are
**not** specially typeset — they become plain paragraphs:

- Numbered lists (`1. item`)
- Bullets with `* ` or `+ ` — only `- ` is a bullet, by design, so that
  `*italic*` at the start of a paragraph doesn't turn into a bullet
- Tables (`| a | b |`)
- Fenced code blocks (` ``` `)
- Block quotes (`> text`)
- Clickable links — `[text](url)` stays as literal text
- Images
- Headings deeper than `####`
- YAML frontmatter
- Role/date separators other than `·`
