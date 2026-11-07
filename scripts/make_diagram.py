#!/usr/bin/env python3
"""make_diagram.py: the one command to run.

  python3 scripts/make_diagram.py SPEC.json --out NAME.pptx [--png preview.png]

1. Reads SPEC.json leniently and writes SPEC.normalized.json (the clean spec; edit THAT file on later rounds).
2. Validates structure, builds an editable .pptx, renders a PNG preview if LibreOffice is installed.
3. Prints a report and writes NAME.report.json.

Exit codes: 0 = PASS (deliver it)  1 = FIX (file built, but follow NEXT ACTIONS and rerun)  2 = INVALID spec (nothing built)
"""
import argparse, json, os, sys, shutil

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import normalize as N          # noqa: E402
import archdiagram as A        # noqa: E402

LIMITS = {"cards": 12, "list": 7, "flow": 6, "band": 7, "nodes": 16, "edges": 24}


PLACEHOLDER = __import__("re").compile(r"<<|>>|\bREPLACE\b|\bTODO\b|\bTBD\b|\bXXX+\b|[Ll]orem ipsum")


def strings(o, path=""):
    if isinstance(o, dict):
        for k, v in o.items(): yield from strings(v, f"{path}.{k}" if path else k)
    elif isinstance(o, list):
        for i, v in enumerate(o): yield from strings(v, f"{path}[{i}]")
    elif isinstance(o, str): yield path, o


def validate(spec):
    errs, warns = [], []
    for path, text in strings(spec.get("slides", []), "slides"):
        if PLACEHOLDER.search(text):
            errs.append((path, f'template placeholder left in: "{text}"', "replace it with real content from the request, or delete that item"))
    if not spec.get("slides"): errs.append(("slides", "spec has no slides", "add at least one slide with rows or nodes"))
    for i, s in enumerate(spec.get("slides", [])):
        p = f"slides[{i}]"
        if "nodes" in s:
            if not s["nodes"]: errs.append((p, "connected slide has no nodes", "add nodes")); continue
            if len(s["nodes"]) > LIMITS["nodes"]:
                warns.append((p, f"{len(s['nodes'])} nodes on one slide", f"keep to {LIMITS['nodes']} or fewer; split into two slides"))
            if len(s.get("edges", [])) > LIMITS["edges"]:
                warns.append((p, f"{len(s['edges'])} lines on one slide", "drop minor lines or split the slide"))
            cells = {}
            for k, n in enumerate(s["nodes"]):
                if not n.get("title"): errs.append((f"{p}.nodes[{k}]", "node has no title", "give every node a 1 to 3 word title"))
                key = (n.get("col"), n.get("row"))
                if key in cells:
                    errs.append((f"{p}.nodes[{k}]", f'"{n["title"]}" and "{cells[key]}" are in the same cell col={key[0]} row={key[1]}',
                                 "give one of them a different col or row, or delete col/row from every node to auto-layout"))
                cells[key] = n.get("title")
            if s.get("lanes"):
                for k, n in enumerate(s["nodes"]):
                    if n.get("row", 0) >= len(s["lanes"]):
                        errs.append((f"{p}.nodes[{k}]", f'"{n["title"]}" is in row {n["row"]} but there are only {len(s["lanes"])} lanes', "set its lane to one of the lane names"))
            if not s.get("edges"): warns.append((p, "connected slide has no lines", "add edges, or use a layered slide instead"))
        else:
            for ri, r in enumerate(s.get("rows", [])):
                if not r.get("panels"): errs.append((f"{p}.rows[{ri}]", "row has no panels", "add a panel or delete the row"))
                if len(r.get("panels", [])) > 4: warns.append((f"{p}.rows[{ri}]", f"{len(r['panels'])} panels in one row", "use at most 3 panels per row"))
                for pi, pn in enumerate(r.get("panels", [])):
                    for bi, b in enumerate(pn["blocks"]):
                        bp = f"{p}.rows[{ri}].panels[{pi}].blocks[{bi}]"
                        t = b["type"]
                        if t in LIMITS and len(b.get("items", [])) > LIMITS[t]:
                            warns.append((bp, f"{t} has {len(b['items'])} items", f"keep to {LIMITS[t]} or fewer"))
                        if t in ("heading", "text", "bar", "chip") and not b.get("text"):
                            errs.append((bp, f"{t} block has no text", "add text or delete the block"))
                        for k, it in enumerate(b.get("items", [])):
                            if t != "band" and not it.get("title"): errs.append((f"{bp}.items[{k}]", "item has no title", "give it a 1 to 3 word title"))
                            if t == "band" and not it.get("label"): errs.append((f"{bp}.items[{k}]", "band item has no label", "add a 1 to 2 word label"))
    return errs, warns


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("spec"); ap.add_argument("--out", required=True); ap.add_argument("--png", default=None)
    a = ap.parse_args()
    raw = open(a.spec, encoding="utf-8").read()
    base = os.path.splitext(a.out)[0]
    report = {"result": None, "output": None, "preview": [], "normalized_spec": None, "auto_fixes": [], "issues": [], "next_actions": []}
    try:
        data = N.loads_tolerant(raw)
    except Exception as e:  # unreadable JSON
        report["result"] = "INVALID"
        report["next_actions"] = [f"The spec is not valid JSON ({e}). Rewrite it as one JSON object; check quotes, commas and brackets."]
        return finish(report, base, 2)
    try:
        spec, notes = N.normalize(data)
    except Exception as e:
        import traceback
        report["result"] = "INVALID"
        report["next_actions"] = [f"The spec structure could not be read ({type(e).__name__}: {e}). Compare it with a file in templates/ and match that shape.",
                                  "Internal detail: " + traceback.format_exc(limit=1).strip().splitlines()[-1]]
        return finish(report, base, 2)
    norm_path = os.path.splitext(a.spec)[0].replace(".normalized", "") + ".normalized.json"
    json.dump(spec, open(norm_path, "w"), indent=1, ensure_ascii=False)
    report["normalized_spec"] = norm_path; report["auto_fixes"] = notes
    errs, warns = validate(spec)
    for sev, lst in (("error", errs), ("warning", warns)):
        for path, prob, fix in lst:
            report["issues"].append({"severity": sev, "slide": None, "path": path, "problem": prob, "fix": fix})
    if errs:
        report["result"] = "INVALID"
        report["next_actions"] = action_list(report["issues"], norm_path)
        return finish(report, base, 2)
    try:
        d = A.build(spec, a.out)
    except SystemExit as e:
        report["result"] = "INVALID"; report["next_actions"] = [str(e)]
        return finish(report, base, 2)
    report["output"] = os.path.abspath(a.out)
    report["issues"] += d.issues
    if a.png:
        if shutil.which("soffice") and shutil.which("pdftoppm"):
            try:
                A.export_png(a.out, a.png)
                stem = os.path.splitext(os.path.basename(a.png))[0]; folder = os.path.dirname(os.path.abspath(a.png))
                report["preview"] = sorted(os.path.join(folder, f) for f in os.listdir(folder) if f.startswith(stem + "-") and f.endswith(".png"))
            except Exception as e:
                report["issues"].append({"severity": "info", "slide": None, "path": "", "problem": f"preview failed: {e}", "fix": "ignore; the .pptx is fine"})
        else:
            report["issues"].append({"severity": "info", "slide": None, "path": "", "problem": "no preview: LibreOffice (soffice) or pdftoppm not installed", "fix": "rely on this report"})
    n_err = sum(1 for i in report["issues"] if i["severity"] == "error")
    report["min_text_scale"] = round(d.min_scale, 2)
    ok = n_err == 0 and d.min_scale >= 0.75
    report["result"] = "PASS" if ok else "FIX"
    report["next_actions"] = action_list(report["issues"], norm_path) if not ok else []
    return finish(report, base, 0 if ok else 1)


def action_list(issues, norm_path):
    acts, seen = [], set()
    for sev in ("error", "warning"):
        for i in issues:
            if i["severity"] != sev: continue
            key = (i["path"], i["problem"])
            if key in seen: continue
            seen.add(key)
            where = f"{i['path']}: " if i["path"] else ""
            acts.append(f"[{sev.upper()}] {where}{i['problem']}. FIX: {i['fix']}.")
    if acts:
        acts.append(f"Edit {os.path.basename(norm_path)} (not the original spec), then rerun make_diagram.py on it.")
    return acts[:15]


def finish(report, base, code):
    json.dump(report, open(base + ".report.json", "w"), indent=1, ensure_ascii=False)
    print(f"RESULT: {report['result']}")
    if report.get("output"): print(f"OUTPUT: {report['output']}")
    for p in report.get("preview", []): print(f"PREVIEW: {p}")
    if report.get("normalized_spec"): print(f"CLEAN SPEC: {report['normalized_spec']}")
    if report.get("min_text_scale") is not None: print(f"TEXT SCALE: {int(report['min_text_scale']*100)}% (75% or more required)")
    if report["auto_fixes"]:
        print(f"AUTO-FIXES APPLIED ({len(report['auto_fixes'])}):")
        for n in report["auto_fixes"][:12]: print("  -", n)
        if len(report["auto_fixes"]) > 12: print(f"  ... {len(report['auto_fixes']) - 12} more in the report file")
    warn_only = [i for i in report["issues"] if i["severity"] == "warning"]
    if report["result"] == "PASS" and warn_only:
        print("NOTES (no action required):")
        for i in warn_only[:6]: print(f"  - {i['path']}: {i['problem']}")
    if report["next_actions"]:
        print("NEXT ACTIONS (do all of these, then rerun):")
        for k, act in enumerate(report["next_actions"], 1): print(f"  {k}. {act}")
    print(f"REPORT: {base}.report.json")
    return code


if __name__ == "__main__":
    sys.exit(main())
