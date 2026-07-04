# mt-graph

Interactive tournament results graph. Reads match results from an Excel spreadsheet and produces a standalone HTML visualization using [pyvis](https://pyvis.readthedocs.io/) and [vis-network](https://visjs.github.io/vis-network/).

Edges point from loser → winner. Following arrows takes you up the dominance hierarchy.

## Usage

```bash
python build_graph.py                        # outputs index.html
python build_graph.py results.xlsx out.html  # custom input/output
```

Open the output file in a browser. No server required.

## Features

- Node size reflects win count; edge thickness reflects time margin
- Left sidebar lists players ranked by tiebreakers (wins → Median-Buchholz → H2H → ΔTimes)
- Week filter shows cumulative standings through any round
- Clicking a node highlights its full ancestry chain with depth-based coloring
- Hovering an ancestor while a node is selected traces the path between them
- Browser back/forward navigates selection history; shareable URLs encode player and week

## Setup

Requires Python 3.12+.

```bash
pip install -r requirements.txt
```

## Deployment

Pushes to `main` automatically build and deploy to GitHub Pages via the workflow in `.github/workflows/pages.yml`.
