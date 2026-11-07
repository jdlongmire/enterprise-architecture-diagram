---
name: enterprise-architecture-diagram
description: Make editable PowerPoint diagrams (architecture overviews, capability maps, system context diagrams, swimlane process flows) in a clean pastel-card, line-icon style with live connectors. Use when the user asks for an architecture diagram, platform overview, capability map, system context, data flow, or process/swimlane diagram as a slide or .pptx.
---

# Enterprise Architecture Diagram

You write a small JSON **spec** that describes the content. A script turns it into an editable PowerPoint file and tells you exactly what to fix. **You never draw shapes or choose coordinates yourself.**

Everything you need is in this folder:
- `scripts/make_diagram.py` is the only command you run.
- `templates/` holds four fill-in specs. Always start from one.
- `examples/` holds finished specs with `.preview.png` images of the result.
- `reference/spec_reference.md` lists every field. `reference/icons.md` lists the icons. `reference/report_guide.md` explains every report message.

Requirements: Python 3 with `pip install python-pptx lxml cairosvg`. A preview PNG also needs LibreOffice (`soffice`) and `pdftoppm`. The diagram is built without them.

---

## Step 1: Choose the template

| The user wants... | Use template | Slide kind |
|---|---|---|
| An overview of a platform or system: layers, capabilities, benefits, a short process | `templates/architecture_overview.json` | layered |
| A grid of capabilities, services or functions | `templates/capability_map.json` | layered |
| Which systems talk to which, data flows, interfaces, trust boundaries | `templates/system_context.json` | connected (live connectors) |
| A process with roles, decisions, approvals, rework loops | `templates/swimlane_process.json` | connected (live connectors) |
| Several of these | Put each slide in one file: `{"slides": [ slide1, slide2 ]}` | mixed |

If the user asks for "connectors", "arrows between boxes", or lines that stay attached, use a **connected** template.

## Step 2: Fill the template

Copy the template to a new file (for example `spec.json`) and replace every `<<...>>` with real content from the user's request.

**Hard rules:**
1. Use only facts the user gave or clearly implied. Do not invent numbers, products or vendors.
2. **Titles:** 1 to 3 words, Title Case. For example "Data Lake" or "Mission Planning".
3. **Descriptions:** 3 to 6 words, no period. For example "Catalog and long-term storage".
4. **Length limits:** card title 18 characters or fewer, box title 22 or fewer, any description 40 or fewer (28 inside connected boxes), line labels 1 or 2 words.
5. **Item counts:** 4 to 6 list items, 4 to 12 cards, 4 to 6 flow steps, 6 to 14 boxes per connected slide, 3 to 5 lanes.
6. **Placeholders:** delete template items you have no content for. Never leave `<<...>>` in the file.
7. **Icons:** `icon` may be any plain word ("database", "security", "ai"), or you can leave it out.
8. **Connected slides:** do not add `col` or `row`, because the script places the boxes. List nodes in the order the work flows. In edges, refer to nodes by their exact title.
9. **Lanes:** every node in a swimlane process needs `"lane"` set to one of the lane names.
10. **Decisions:** a decision node is a short question ending in "?", and its outgoing edges are labeled "Yes" and "No".
11. **Dashed lines:** use `"dashed": true` for rework loops, authentication and optional or async flows.
12. **Core box:** in a system context, give the one core system `"shape": "banner"`.

## Step 3: Build

```
python3 scripts/make_diagram.py spec.json --out <Topic>_Diagram.pptx --png preview.png
```

The script prints a report. Read the first line:

| RESULT | Meaning | What you do |
|---|---|---|
| `PASS` | The file is good. | Go to Step 5. |
| `FIX` | The file was built but needs changes. | Go to Step 4. |
| `INVALID` | Nothing was built; the spec is broken. | Go to Step 4. |

## Step 4: Fix and rebuild (repeat up to 4 times)

1. Open the file named on the `CLEAN SPEC:` line (it ends in `.normalized.json`). **Edit that file from now on**, not your original.
2. Do **every** numbered item under `NEXT ACTIONS`, in order. Each item gives a path (like `slides[0].nodes[3].desc`), the problem, and the exact fix (for example "shorten it to at most 20 characters"). Count characters; do not guess.
3. Rebuild using the clean spec:
   ```
   python3 scripts/make_diagram.py spec.normalized.json --out <Topic>_Diagram.pptx --png preview.png
   ```
4. Stop when RESULT is `PASS`. `[WARNING]` items never block a PASS. Fix them only if that's easy.

If it is still `FIX` after 4 rounds, deliver the file anyway and tell the user which items remain.

**Common fixes:**
- *"wraps to N lines"*: shorten that exact text to the character count given. Keep its meaning ("Procedure and command set" becomes "Command set").
- *"content is too tall"* or *"text shrunk"*: delete the least important list item or card, or shorten the longest descriptions on that slide.
- *"same cell"*: delete `col` and `row` from **every** node so auto-layout handles placement.
- *"passes through another box"*: delete `col`/`row` from all nodes, or remove that line if it is minor.
- *"placeholder left in"*: replace the `<<...>>` text with real content, or delete the item.

## Step 5: Check and deliver

1. If you can view images, open each preview PNG listed in the report and confirm: no text runs outside a box, no line passes through a box, and the icons make sense. If something is wrong, fix the spec and rebuild.
2. Give the user the `.pptx` file. Tell them in one or two sentences what it shows, and that every box, label and line is editable in PowerPoint (lines stay attached when boxes are moved).
3. Do not paste the JSON spec into your answer unless the user asks for it.

---

## Worked example

User: *"Make a slide showing how a purchase request gets approved: employee submits, manager reviews, anything over $10k goes to finance, finance can reject it back to the employee, then procurement issues the PO."*

1. **Template:** `swimlane_process.json`, with lanes Employee, Manager, Finance and Procurement.
2. **Nodes (in order):** Submit Request (start), Review Request, Over $10k? (decision), Finance Review, Revise Request, Issue PO, Done (end).
3. **Edges:** Yes goes to Finance Review, No goes to Issue PO, and a dashed "Rejected" line goes to Revise Request. Revise Request has a dashed line back to Review Request.
4. **Build:** `RESULT: PASS`. Deliver.

The finished version is `examples/purchase_approval_swimlane.json`, and its picture is `examples/purchase_approval_swimlane-1.preview.png`.

## Brand colors and fonts (optional)

Add `"theme"` at the top level of the spec:
```json
"theme": {"font": "Arial", "accent": "1F3A68", "accent2": "172E54", "arrow": "4F6D9A"}
```
`accent` is the main banner color and `arrow` is the line color. Colors are 6-digit hex, with or without `#`. Other keys are listed in `reference/spec_reference.md`.
