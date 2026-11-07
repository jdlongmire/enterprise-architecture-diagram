# Reading the report

Every run prints a report and saves it as `<output>.report.json`. The JSON has the same content:
`result`, `output`, `preview`, `extras`, `normalized_spec`, `auto_fixes`, `issues` (each with `severity`, `path`, `problem`, `fix`), `next_actions`, `min_text_scale`.

`extras` lists the extra output files made when `--pdf`, `--jpg` or `--svg` was given, keyed by format. The printed report shows them as `PDF:`, `JPG:` and `SVG:` lines. An export problem (for example LibreOffice missing) appears as an `info` issue; it never changes the result, and the `.pptx` is still delivered.

With `--dry-run` the report has `"dry_run": true` and `output` is null: the layout and checks ran in full, but no file was kept. Use it to read the real per-box limits for a crowded spec before writing final text.

## Result

- **PASS**: no errors, and text is at 75% or more of its normal size. Deliver the file. Between 75% and 85% the run still passes but carries a warning that the text is small when projected; shorten the longest descriptions or accept it knowingly.
- **FIX**: the file was built but has errors. Do every item under NEXT ACTIONS and rerun.
- **INVALID**: nothing was built. The spec is unreadable or structurally wrong. Do the NEXT ACTIONS, or compare your file with a template.

## Paths

A path such as `slides[0].rows[1].panels[0].blocks[2].items[3].desc` points into the **clean spec** (`*.normalized.json`): slide 1, row 2, panel 1, block 3, item 4, field `desc`. Indexes start at 0. The message also quotes the current text, so you can find it by searching.

## Messages and what to do

| Message contains | Do this |
|---|---|
| `wraps to N lines (limit M)` | Shorten that exact text to the number of characters given. Keep the meaning. |
| `words too long for their box` | Replace that word with a shorter word or abbreviation. |
| `content is ... too tall` | Delete the least important item, shorten long descriptions, or move the last row to a new slide. |
| `text shrunk to N%` | Shorten the longest descriptions on that slide, or give the crowded panel a bigger `w`. |
| `text at N% of normal size (below 85%)` | Warning only, on a PASS. The text will look small projected. Shorten the longest descriptions or widen the crowded boxes, or accept it. |
| `boxes are too tall for their rows` | Shorten node descriptions, or use fewer lanes. |
| `same cell` | Delete `col` and `row` from every node (auto-layout), or move one node. |
| `passes through another box` | Delete `col`/`row` from all nodes, or remove that line. |
| `shares a box side with another line` | Warning only. Fine for a fork or merge. |
| `line label ... touches another line` | Warning only. Shorten the label, or drop it if the line is obvious. |
| `line crossings` | Warning only. Reorder nodes so connected nodes are listed next to each other. |
| `template placeholder left in` | Replace the `<<...>>` text with real content, or delete the item. |
| `has no title` / `has no text` | Add the missing text, or delete the item. |
| `unknown icon` | Use a word from `reference/icons.md`, or delete the `icon` field. |
| `not valid JSON` | Fix quotes, commas and brackets, or start again from a template. |

## Auto-fixes

The script repairs common slips and lists each one under AUTO-FIXES: wrong block names (`steps` becomes `flow`), field names (`description` becomes `desc`), icon words (`database` becomes `db`), lane name typos, missing node ids, node placement, and nodes sitting inside the wrong zone. You don't need to act on these. If an auto-fix chose badly (for example an odd icon), set that field explicitly in the clean spec.
