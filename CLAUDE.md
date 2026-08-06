# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What this is

A static marketing website for **Shaqool**, the publisher of SCQF (the Saudi Construction Quality Framework). Four standalone, hand-written HTML pages. There is **no build system, no package manager, no dependencies, and no backend** — every page is a single self-contained `.html` file with its CSS in an inline `<style>` block and its behaviour in an inline `<script>` block.

- `index.html` — home / overview
- `scqf.html` — the SCQF standard
- `advisory.html` — Shaqool Advisory
- `register.html` — register-interest form

## Running / previewing

Open any `.html` file directly in a browser, or serve the folder statically:

```
python3 -m http.server 8000    # then visit http://localhost:8000/
```

There is nothing to build, lint, or test. The only external runtime dependency is Google Fonts (IBM Plex Sans / Sans Arabic / Mono, Noto Kufi Arabic), loaded via `<link>`; the site falls back to system fonts (SF Arabic / Geeza Pro / Noto / Tahoma) offline without shifting layout.

## Architecture

**Bilingual by construction.** The site is Arabic-first (default `dir="rtl"`, `lang="ar"`) with an English (LTR) alternative. Language is toggled by a single hidden checkbox `#lang` plus CSS sibling selectors (`#lang:checked ~ .app ...`) — **the switch works with JavaScript disabled** (this is deliberate: iOS previews downloaded HTML with JS off). Every visible string is authored as an adjacent pair of spans:

```html
<span class="ar">النص العربي</span><span class="en">English text</span>
```

`.en` is hidden by default; checking `#lang` hides `.ar` and shows `.en`, and also flips `direction`, font family, and heading styles for the whole `.app` wrapper. When you add or edit copy, you must supply **both** the `.ar` and `.en` sibling — never one alone.

**Inline JS is progressive enhancement only.** The `<script>` at the top sets `documentElement.className="js"` so reveal animations only ever hide content when JS is actually running. The tail `<script>` does two things: (1) `sync()` mirrors the checkbox into `html.lang` / `html.dir` and swaps `document.title`, and (2) an `IntersectionObserver` adds `.in` to `.rv` elements to fade them in on scroll (with a no-JS/no-observer fallback that just shows everything).

**The design system and boilerplate are duplicated across all four pages.** Each file carries its own full copy of: the `:root` CSS custom properties (colours like `--ink`, `--paper`, `--accent`, `--brass`; layout tokens; font-family variables), the language-switch CSS, the cross-platform hardening block, the header/nav, the footer, and both `<script>` blocks. There is no shared stylesheet or JS file. **A change to shared styling, the header/nav, the footer, or the language/reveal scripts must be applied to all four HTML files** to keep them consistent. The per-page `sync()` function is the one intentional difference: each page hard-codes its own bilingual `document.title` pair.

**Illustrations are inline SVG, not images.** The technical "plates" (e.g. `PL. 01`) are drawn as inline `<svg>`. See the `PHOTOGRAPHY NOTES` / `TECHNICAL NOTES` HTML comment at the bottom of `index.html` for how to swap an SVG plate for a real photo and the intended art direction.

## Conventions

- **Contact address:** `hello@shaqool.org`. Keep it consistent across pages if it ever changes (footers, and the "contact us directly" fallback links on `register.html`).
- **Register form delivery:** the site is static (GitHub Pages), so the form on `register.html` submits to **Web3Forms** (`action="https://api.web3forms.com/submit"`), which emails each enquiry to `hello@shaqool.org`. The public `access_key` hidden field must hold a real key for delivery to work (see the setup comment above the `<form>`). The tail `<script>` intercepts submit and posts via `fetch()` so it sends in place and shows an inline `.formmsg` acknowledgement — with a hidden `botcheck` honeypot for spam and a `redirect` field so no-JS submissions return to `?sent=1`. There is no `mailto:` form action anymore.
- **Scope wording is substantive, not filler.** Claims about what SCQF is and is not (e.g. "we do not replace or duplicate the Saudi Building Code", warranty/scope boundaries in "Part 100") are deliberate positioning. Treat copy edits to these sections as content decisions, not stylistic tweaks.
- Class-name idiom is terse and semantic (`.hd`, `.hero`, `.plate`, `.band`, `.fence`, `.rv`, `.mono`, `.eyebrow`). Match the existing naming and the compact, minified-ish CSS style already in the file rather than reformatting it.
