# CLAUDE.md

## Architecture

This is a single-file Python build script (`build_graph.py`) that produces a standalone HTML file (`index.html`). There is no frontend framework or bundler.

- `build_graph.py` — reads `results.xlsx`, builds a NetworkX graph, and uses pyvis to generate the HTML. All custom CSS and JavaScript lives in the `_ANCESTOR_JS` string constant near the bottom of the file. This string is injected into the pyvis-generated HTML by `_inject_ancestor_highlight()`.
- `results.xlsx` — source data with one sheet per wave/round plus a Group Summary sheet for final standings.
- `index.html`, `lib/`, `dist/` — build artifacts, gitignored.

## Important: where to edit JavaScript and CSS

**Do not edit `index.html` directly.** It is a generated file. All JS/CSS changes must be made inside the `_ANCESTOR_JS` string in `build_graph.py`, then rebuilt.

## Build

```bash
python build_graph.py        # produces index.html locally for testing
```

## Deployment

GitHub Actions (`.github/workflows/pages.yml`) builds on push to `main` and deploys `dist/index.html` to GitHub Pages.

## Excel sheet layout

Each wave sheet (0-based column indices):
- 1: Player 1 name
- 2: Player 1 finish time
- 3: Player 2 name
- 4: Player 2 finish time
- 5: Winner name
- 11: "Played" flag (row only counted when truthy)

Group Summary sheet:
- 3: Player name
- 8: Median-Buchholz
- 10: Head-to-Head
- 11: ΔTimes
