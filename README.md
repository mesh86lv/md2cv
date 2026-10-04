# md2cv

Typeset a specific structured-Markdown résumé shape into neatly laid-out DOCX.

[![Python](https://img.shields.io/badge/python-3.9%2B-blue.svg)](https://www.python.org/)
[![License: MIT](https://img.shields.io/badge/license-MIT-green.svg)](LICENSE)
[![Made with python-docx](https://img.shields.io/badge/built%20with-python--docx-blueviolet.svg)](https://python-docx.readthedocs.io/)

> **Scope.** `md2cv` is **not** a general Markdown→DOCX converter. It targets
> one specific résumé shape, documented in
> [USAGE.md](USAGE.md#supported-markdown-shape) and demonstrated in
> [`examples/sample.md`](examples/sample.md). Files that deviate from that
> shape still convert — unrecognised parts fall through as plain paragraphs —
> but the careful typography (right-aligned dates, hanging bullets, rules
> under sections) only kicks in for the supported patterns. If you need a
> universal converter, use [pandoc](https://pandoc.org/).

## Why

- **Compact by default.** 9.5 pt body, narrow margins, single line spacing —
  a two-page résumé that still reads comfortably.
- **Consistent.** Company names, roles, dates and projects all sit on a fixed
  grid instead of drifting per section.
- **Configurable.** Every knob (fonts, sizes, margins, colors, spacing) is a
  CLI flag, so you can trade space for air without editing the source.
- **NixOS-friendly.** Ships with a `shell.nix`; also installable via `pipx`.

## Install

### With pipx (any OS)

```bash
pipx install git+https://github.com/mesh86lv/md2cv.git
```

Or clone and install locally:

```bash
git clone https://github.com/mesh86lv/md2cv.git
cd md2cv
pipx install .
```

### From a Nix shell (NixOS, no install)

```bash
git clone https://github.com/mesh86lv/md2cv.git
cd md2cv
nix-shell
md2cv examples/sample.md
```

## Usage

```bash
md2cv cv.md                       # -> cv.docx next to it
md2cv cv.md --preset compact      # squeeze into two pages
md2cv ~/cv/*.md -d ~/cv/docx      # batch into a folder
md2cv examples/                   # all *.md from a directory
```

Full reference, every flag, sensible ranges, and more recipes:
**[USAGE.md](USAGE.md)**.

## Expected Markdown shape

```markdown
# Full Name

**Tagline** — one line

City · phone · email
GitHub: … · LinkedIn: …

## Section Title

### Company Name

**Role Title** · Jan 2020 – Present

_Optional italic blurb about the role._

_Stack:_ Node.js, PostgreSQL

- A bullet point

#### Project: Sub-project

**Project-level Tech Lead** · Jan 2020 – Dec 2021

- Bullet under the project
```

The parser is intentionally tolerant: anything that isn't a heading, bullet
or bold-inline line becomes a plain paragraph. See
[USAGE.md](USAGE.md#supported-markdown-shape) for what is and isn't
recognised.

## Development

```bash
git clone https://github.com/mesh86lv/md2cv.git
cd md2cv
nix-shell                  # or: python -m venv .venv && . .venv/bin/activate && pip install -e '.[dev]'
python md2cv.py examples/sample.md
```

## License

[MIT](LICENSE)
