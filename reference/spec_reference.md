# Spec reference

A spec is one JSON object. Comments (`//`), trailing commas and code fences are allowed; the script cleans them up.

```
{ "theme": {...optional...},
  "slides": [ <layered slide> or <connected slide>, ... ] }
```
A file with a single slide may skip `slides` and put the slide's fields at the top level.

## Layered slide (overviews, capability maps)

```
{ "title": "...", "subtitle": "...",                       // optional slide title row
  "rows": [ { "panels": [ { "title": "...", "w": 1, "align": "center", "blocks": [ ... ] } ] } ],
  "notes": "speaker notes (optional)" }
```
- `rows`: horizontal bands, top to bottom. Use at most 3 rows.
- `panels`: white cards across a row, left to right. Use at most 3 per row. `w` is a relative width (for example 0.9 / 2.6 / 0.9).
- A panel `title` becomes its heading. `align: "center"` suits the main architecture column. `style: "none"` removes the white card.

### Blocks (inside a panel, top to bottom)

| type | fields | notes |
|---|---|---|
| `heading` | `text`, `size` (15) | Panel header. |
| `subheading` | `text` | Small bold label. |
| `text` | `text` | 1 to 3 sentences. |
| `banner` | `title`, `subtitle` | The bold colored layer. Use **one per panel**, for the central idea. `size: "sm"` makes a compact version. |
| `cards` | `items`, `columns`, `layout` | Pastel tiles. `layout: "vertical"` (default; icon on top) or `"horizontal"` (icon left; good for outcomes). Up to 12 items. |
| `band` | `title`, `items` (with `label`), `layout` | A strip of icons with labels (users, sources, integrations). `layout: "inline"` puts icon and label side by side. |
| `bar` | `text` | One line of shared services: `"Shared: A · B · C"`. |
| `chip` | `text` | Small caption box. |
| `list` | `items`, `hued` | Icon + title + description rows (benefits, principles, use cases). `hued: true` colors the icons. |
| `flow` | `items` | Steps joined by arrows. 4 to 6 steps; put it in a full-width row. Number the titles ("1. Collect"). |

### Items (`cards`, `list`, `flow`)
`{"title": "...", "desc": "...", "icon": "...", "hue": "..."}`. Only `title` is required.
Hues: `blue green violet amber rose teal slate`. Leave `hue` out and the colors cycle automatically.

## Connected slide (system context, data flow, swimlane process)

```
{ "title": "...", "subtitle": "...",
  "lanes": ["Role A", "Role B"],                 // optional: swimlanes, one per role
  "nodes": [ {"title": "...", "desc": "...", "icon": "...", "shape": "card", "lane": "Role A"} ],
  "edges": [ {"from": "<node title>", "to": "<node title>", "label": "...", "dashed": false, "arrow": "end"} ],
  "zones": [ {"title": "Our Cloud", "nodes": ["<node title>", ...], "hue": "blue"} ] }
```
- **Node `shape`:** `card` (default), `banner` (the core system: bold color), `start` / `end` (rounded pill), `decision` (diamond; short question ending in "?").
- **Placement:** do **not** set `col`/`row`. The script places nodes left to right in flow order, keeps rework steps close to the step they return to, and keeps zones together. (Experts can set `col` and `row` on every node; partial placement is not recommended.)
- **Edges** are real PowerPoint connectors glued to the boxes. Options: `label` (1 to 2 words), `dashed` (rework, auth, async), `arrow`: `end` (default) / `both` / `none`. `from_side`/`to_side` (`top|left|bottom|right`) force where a line attaches.
- **Zones** are dashed boundary boxes drawn around the listed nodes. Nodes that don't belong are moved out automatically.

## Theme (optional, top level)

| key | meaning | default |
|---|---|---|
| `font` | Font for all text. Calibri and Arial fit most accurately. | Calibri |
| `accent`, `accent2` | Banner gradient colors | 2F6FD6, 2A63C4 |
| `arrow` | Line and arrow color | 6D8FC9 |
| `ink`, `ink2`, `ink3` | Text colors (titles, body, captions) | 2A3546, 4A5568, 6B7686 |
| `canvas` | Slide background | EEF2F7 |
| `hues` | Override a hue: `{"blue": ["tint", "border", "icon"]}` | (see engine) |
| `hue_order` | The order colors cycle in | blue, green, violet, amber, rose, teal |

## Size guide (the report is the final judge)

| text | target |
|---|---|
| Slide title | 3 to 7 words |
| Card / box title | 1 to 3 words, 18 to 22 characters or fewer |
| Description | 3 to 6 words; 40 characters or fewer (28 in connected boxes) |
| Band label | 1 or 2 words |
| Line label | 1 or 2 words (HTTPS, Yes, orders) |
| Banner subtitle | 12 words or fewer |
