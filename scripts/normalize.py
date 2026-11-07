"""normalize.py: turn a loosely written diagram spec into the strict form the engine expects.

Forgiving by design, so a model can write the spec naturally:
  * tolerant JSON: code fences, // and /* */ comments, trailing commas, smart quotes
  * alias field and type names (description->desc, name->title, steps->flow, source/target->from/to ...)
  * plain-word icons ("database", "security", "ai") mapped to the icon set; missing icons inferred from titles
  * node ids optional; edges may reference node titles
  * connected slides: col/row optional (auto-layout); lanes may be given per node as "lane": "<name>"
  * zones may list member node ids instead of grid spans
Every change is recorded in `notes` so the caller can report what was fixed.
"""
import json, re, os, difflib

HERE = os.path.dirname(os.path.abspath(__file__))
ICONS = json.load(open(os.path.join(HERE, "icons.json")))

# ---------------------------------------------------------------- tolerant JSON
def loads_tolerant(text):
    t = text.strip()
    m = re.search(r"```(?:json)?\s*(.*?)```", t, re.S)
    if m: t = m.group(1)
    t = t.replace("“", '"').replace("”", '"').replace("‘", "'").replace("’", "'")
    out, i, n, in_str = [], 0, len(t), False      # strip comments outside strings
    while i < n:
        c = t[i]
        if in_str:
            out.append(c)
            if c == "\\" and i + 1 < n: out.append(t[i + 1]); i += 2; continue
            if c == '"': in_str = False
            i += 1; continue
        if c == '"': in_str = True; out.append(c); i += 1; continue
        if t.startswith("//", i):
            j = t.find("\n", i); i = n if j < 0 else j; continue
        if t.startswith("/*", i):
            j = t.find("*/", i + 2); i = n if j < 0 else j + 2; continue
        out.append(c); i += 1
    t = "".join(out)
    t = re.sub(r",(\s*[}\]])", r"\1", t)          # trailing commas
    start = min([k for k in (t.find("{"), t.find("[")) if k >= 0] or [0])
    return json.loads(t[start:])

# ---------------------------------------------------------------- vocabularies
BLOCK_ALIASES = {
    "heading": ["heading", "header", "title", "h1", "h2", "section", "section_title"],
    "subheading": ["subheading", "subheader", "subtitle_label", "h3", "label"],
    "text": ["text", "paragraph", "body", "description", "desc", "lede", "intro", "summary", "subtitle"],
    "banner": ["banner", "layer", "core", "hero", "platform", "harness", "hub", "central"],
    "cards": ["cards", "card", "tiles", "tile", "capabilities", "capability", "grid", "boxes", "modules", "components", "outcomes"],
    "band": ["band", "strip", "actors", "users", "integrations", "icons", "icon_row", "row_of_icons", "personas", "sources"],
    "bar": ["bar", "services", "shared_services", "footer", "ribbon", "foundation"],
    "chip": ["chip", "caption", "tag", "pill", "note"],
    "list": ["list", "bullets", "benefits", "items", "features", "principles", "use_cases", "drivers"],
    "flow": ["flow", "steps", "process", "pipeline", "sequence", "lifecycle", "timeline", "stages"],
    "spacer": ["spacer", "gap", "space"],
}
BLOCK_MAP = {a: k for k, v in BLOCK_ALIASES.items() for a in v}
TITLE_KEYS = ["title", "name", "heading", "label", "header", "step", "capability"]
DESC_KEYS = ["desc", "description", "text", "subtitle", "detail", "details", "summary", "body", "caption", "sub"]
SHAPE_ALIASES = {"card": "card", "box": "card", "rect": "card", "rectangle": "card", "process": "card", "step": "card", "task": "card",
                 "component": "card", "service": "card", "system": "card",
                 "banner": "banner", "core": "banner", "hub": "banner", "primary": "banner", "main": "banner", "central": "banner",
                 "pill": "pill", "start": "pill", "end": "pill", "terminal": "pill", "terminator": "pill", "stop": "pill", "begin": "pill", "event": "pill",
                 "decision": "decision", "diamond": "decision", "gateway": "decision", "choice": "decision", "condition": "decision", "if": "decision"}
HUE_ALIASES = {"blue": "blue", "navy": "blue", "green": "green", "emerald": "green", "violet": "violet", "purple": "violet", "indigo": "violet",
               "amber": "amber", "orange": "amber", "yellow": "amber", "gold": "amber", "rose": "rose", "red": "rose", "pink": "rose",
               "teal": "teal", "cyan": "teal", "aqua": "teal", "slate": "slate", "gray": "slate", "grey": "slate", "neutral": "slate"}

ICON_SYNONYMS = {
    "db": "database data datastore storage store repository warehouse lake archive records sql table catalog",
    "shield": "security secure governance compliance protection risk policy assurance trust safety defense",
    "brain": "ai ml model models llm intelligence machine learning cognitive reasoning neural",
    "bot": "agent agents robot assistant copilot automation_agent chatbot",
    "wrench": "tool tools maintenance fix repair sustainment utility",
    "plug": "api apis integration integrations connector connectors adapter interface plugin",
    "cloud": "cloud saas hosting aws azure gcp",
    "server": "server servers infrastructure infra compute hardware datacenter on-prem onprem host station stations",
    "user": "user person customer individual operator analyst engineer persona human",
    "users": "users team people stakeholders community group customers staff workforce",
    "cog": "settings config configuration operations ops engine processing processor service setup conversion",
    "workflow": "workflow automation automate process processes procedure",
    "route": "orchestration orchestrate routing route plan planning journey path",
    "eye": "monitoring observability visibility telemetry watch audit review inspect oversight",
    "chart": "analytics metrics kpi reporting dashboard measure measurement statistics insight insights",
    "trending": "growth trend improvement scale scaling increase performance_trend forecast",
    "clock": "time latency speed schedule_time realtime duration deadline",
    "calendar": "schedule scheduling calendar planning_calendar date tasking",
    "file": "document documents doc report reports file files record spec specification artifact evidence",
    "mail": "email mail notification notify notice newsletter",
    "message": "chat message messaging communication conversation comment feedback",
    "network": "network topology hierarchy org structure interoperability architecture partner",
    "globe": "global internet web world public portal website external",
    "key": "identity auth authentication credentials pki keys sso iam login",
    "lock": "access lock authorization encryption privacy secret zero-trust guard guards",
    "bulb": "idea innovation insight_bulb creativity research discovery",
    "rocket": "launch rocket satellite spacecraft space deploy deployment",
    "plane": "aircraft plane flight aviation airborne drone uav",
    "wifi": "wireless radio rf signal link downlink uplink antenna comms communications connectivity",
    "cpu": "processor chip cpu compute_unit edge embedded hardware_processing fpga",
    "search": "search discovery query find lookup investigate analysis",
    "alert": "alert alarm warning incident issue anomaly threat",
    "check": "success done approve approved complete validated validation quality qa verify verified",
    "target": "goal target objective mission aim focus outcome",
    "dollar": "money cost finance financial budget revenue price funding",
    "code": "code development developer software devops programming app application",
    "terminal": "terminal cli console command shell scripting",
    "layers": "layers platform stack tiers fabric foundation layer",
    "box": "package module component product container deliverable",
    "building": "building organization enterprise company facility office business government",
    "branch": "version versioning branch configuration_management baseline release git",
    "link": "link linking connect connection relationship thread traceability",
    "refresh": "refresh sync synchronize loop retry iterate recovery update continuous",
    "filter": "filter triage sort screen prioritize classify",
    "flag": "flag milestone start begin initiative",
    "gauge": "gauge performance health readiness maturity status",
    "phone": "phone mobile device devices smartphone",
    "monitor": "monitor desktop screen display workstation console_screen",
    "laptop": "laptop power-user poweruser endpoint workstation_laptop",
    "sparkles": "sparkles magic generative genai new",
    "zap": "zap fast power event events trigger streaming stream real-time instant",
    "book": "book knowledge docs documentation library training learning guide",
    "sliders": "sliders tuning controls control parameters calibration adjust",
    "building": "building organization enterprise company facility office business government erp supplier vendor bank",
    "users": "users team people stakeholders community group customers staff workforce crm audience",
    "wifi": "wireless radio rf signal link downlink uplink antenna comms communications connectivity sensor sensors iot telemetry_feed",
    "filter": "filter triage sort screen prioritize classify clean cleanse cleansing validate_data dedupe",
    "file": "document documents doc report reports file files record spec specification artifact evidence request submit form order po invoice contract",
    "check": "success done approve approved approval complete validated validation quality qa verify verified finish finished",
    "refresh": "refresh sync synchronize loop retry iterate recovery update continuous revise rework edit amend",
    "dollar": "money cost finance financial budget revenue price funding payment payments billing card purchase spend",
    "alert": "alert alarm warning incident issue anomaly threat fraud risk_event exception",
    "db": "database data datastore storage store repository warehouse lake archive records sql table catalog ledger cache",
    "plug": "api apis integration integrations connector connectors adapter interface plugin gateway endpoint",
    "monitor": "monitor desktop screen display workstation dashboard_screen ui frontend room console workbench",
}
SYN = {}
for icon, words in ICON_SYNONYMS.items():
    if icon not in ICONS: continue
    for w in words.split(): SYN.setdefault(w.replace("_", " "), icon)
for name in ICONS: SYN[name] = name


def resolve_icon(name, context="", notes=None, where=""):
    """Map any icon word to a real icon name. Falls back to keywords in the item's own text, then 'box'."""
    def try_word(w):
        w = w.lower().strip().replace("_", " ").replace("-", " ")
        if not w: return None
        if w in SYN: return SYN[w]
        w2 = w.rstrip("s")
        if w2 in SYN: return SYN[w2]
        for part in w.split():
            if part in SYN: return SYN[part]
            if part.rstrip("s") in SYN: return SYN[part.rstrip("s")]
        close = difflib.get_close_matches(w, list(SYN), n=1, cutoff=0.82)
        return SYN[close[0]] if close else None
    if name:
        raw = str(name).lower().strip()
        if raw in ICONS: return raw
        got = try_word(str(name))
        if got:
            if got != name and notes is not None: notes.append(f"{where}: icon '{name}' -> '{got}'")
            return got
    title_part, _, rest = (context or "").partition("\x00")
    words = list(reversed(re.findall(r"[A-Za-z][A-Za-z\-]+", title_part))) + re.findall(r"[A-Za-z][A-Za-z\-]+", rest)
    GENERIC = {"service", "services", "system", "systems", "platform", "app", "application", "manager", "management",
               "center", "centre", "hub", "layer", "module", "unit", "engine", "suite", "solution", "function", "capability"}
    words = [w for w in words if w.lower() not in GENERIC] + [w for w in words if w.lower() in GENERIC]
    for word in words:
        got = SYN.get(word.lower()) or SYN.get(word.lower().rstrip("s"))
        if got:
            if notes is not None: notes.append(f"{where}: icon {'chosen' if not name else repr(name)+' unknown,'} -> '{got}' (from text)")
            return got
    if notes is not None: notes.append(f"{where}: icon {repr(name) if name else 'missing'} -> 'box' (default)")
    return "box"


def pick(d, keys, default=None):
    for k in keys:
        if k in d and d[k] not in (None, ""): return d[k]
    return default


def slug(s):
    return re.sub(r"[^a-z0-9]+", "_", str(s).lower()).strip("_") or "n"

# ---------------------------------------------------------------- items / blocks
def norm_item(it, kind, notes, where):
    if isinstance(it, str):
        title, _, desc = it.partition(":") if ":" in it else it.partition(" - ")
        it = {"title": title.strip(), "desc": desc.strip()} if desc.strip() else {"title": it.strip()}
    out = {}
    title = pick(it, TITLE_KEYS, "")
    desc = pick(it, [k for k in DESC_KEYS if k != "text"] + ["text"], None)
    if kind == "band":
        out["label"] = str(pick(it, ["label", "title", "name", "text"], ""))
        out["icon"] = resolve_icon(it.get("icon"), out["label"], notes, where)
        return out
    out["title"] = str(title)
    if desc and str(desc) != str(title): out["desc"] = str(desc)
    out["icon"] = resolve_icon(it.get("icon"), f"{title}\x00{desc or ''}", notes, where)
    h = it.get("hue") or it.get("color") or it.get("colour")
    if h:
        hh = HUE_ALIASES.get(str(h).lower())
        if hh: out["hue"] = hh
        else: notes.append(f"{where}: unknown hue '{h}' ignored")
    return out


def norm_block(b, notes, where):
    if isinstance(b, str): b = {"type": "text", "text": b}
    t = str(b.get("type") or b.get("kind") or b.get("block") or "").lower().replace("-", "_").replace(" ", "_")
    typ = BLOCK_MAP.get(t)
    if not typ:  # infer from fields
        if "items" in b or "cards" in b or "steps" in b:
            typ = "flow" if "steps" in b else "cards"
        elif "subtitle" in b and "title" in b: typ = "banner"
        else: typ = "text"
        notes.append(f"{where}: block type '{t or '?'}' -> '{typ}'")
    elif typ != t: notes.append(f"{where}: block type '{t}' -> '{typ}'")
    out = {"type": typ}
    if typ in ("heading", "subheading", "text", "bar", "chip"):
        out["text"] = str(pick(b, ["text", "title", "content", "value", "label", "body"], ""))
        if typ == "heading" and "size" in b: out["size"] = b["size"]
        if typ == "text" and "size" in b: out["size"] = b["size"]
        if "align" in b: out["align"] = b["align"]
        if typ == "bar" and isinstance(b.get("items"), list):
            out["text"] = (out["text"] + ": " if out["text"] else "") + " · ".join(str(x if isinstance(x, str) else pick(x, TITLE_KEYS, "")) for x in b["items"])
    elif typ == "banner":
        out["title"] = str(pick(b, ["title", "text", "name", "label"], ""))
        sub = pick(b, ["subtitle", "desc", "description", "caption"])
        if sub: out["subtitle"] = str(sub)
        if b.get("size") == "sm" or b.get("small"): out["size"] = "sm"
    elif typ == "spacer":
        out["h"] = float(b.get("h", b.get("height", 0.1)))
    else:
        items = b.get("items") or b.get("cards") or b.get("steps") or b.get("entries") or []
        out["items"] = [norm_item(it, typ, notes, f"{where}.items[{k}]") for k, it in enumerate(items)]
        title = pick(b, ["title", "heading", "label"])
        if typ == "band" and title: out["title"] = str(title)
        for k in ("layout", "columns", "hued"):
            if k in b: out[k] = b[k]
        if typ == "cards" and out.get("layout") not in (None, "vertical", "horizontal"):
            out["layout"] = "horizontal" if "hor" in str(out["layout"]) or "row" in str(out["layout"]) else "vertical"
        if typ == "band" and out.get("layout") not in (None, "stack", "inline"): out["layout"] = "inline"
        if typ in ("list", "cards", "flow", "band") and not out["items"]:
            notes.append(f"{where}: {typ} has no items (block dropped)"); return None
        if typ in ("cards", "list", "flow", "band") and title and typ != "band":
            return [{"type": "heading", "text": str(title)}, out]   # a titled list becomes heading + list
    return out


def norm_panel(p, notes, where):
    if isinstance(p, list): p = {"blocks": p}
    blocks = []
    for k, b in enumerate(p.get("blocks") or p.get("content") or p.get("items") or []):
        nb = norm_block(b, notes, f"{where}.blocks[{k}]")
        if nb is None: continue
        blocks.extend(nb if isinstance(nb, list) else [nb])
    if not blocks:
        title = pick(p, ["title", "heading"])
        blocks = [{"type": "heading", "text": str(title or "")}]
    elif pick(p, ["title", "heading"]) and blocks[0]["type"] != "heading":
        blocks.insert(0, {"type": "heading", "text": str(pick(p, ["title", "heading"]))})
    out = {"blocks": blocks}
    w = p.get("w", p.get("width", p.get("weight")))
    if w is not None:
        try: out["w"] = float(w)
        except (TypeError, ValueError): notes.append(f"{where}: width '{w}' ignored")
    if p.get("align") in ("center", "centre", "middle"): out["align"] = "center"
    if p.get("style") in ("none", "bare", "plain", "transparent"): out["style"] = "none"
    if "distribute" in p: out["distribute"] = bool(p["distribute"])
    return out


def norm_layered(s, notes, where):
    if "rows" not in s:
        if "panels" in s: s = dict(s, rows=[{"panels": s["panels"]}])
        elif "blocks" in s: s = dict(s, rows=[{"panels": [{"blocks": s["blocks"]}]}])
    rows = []
    title = pick(s, ["title"]); sub = pick(s, ["subtitle", "lede"])
    if title and not any(isinstance(r, dict) and any(isinstance(p, dict) and p.get("style") == "none" for p in r.get("panels", [])) for r in s["rows"]):
        blocks = [{"type": "heading", "text": str(title), "size": 22}]
        if sub: blocks.append({"type": "text", "text": str(sub), "size": 11})
        rows.append({"panels": [{"style": "none", "blocks": blocks}]})
    for ri, r in enumerate(s["rows"]):
        if isinstance(r, list): r = {"panels": r}
        if "panels" not in r and "blocks" in r: r = {"panels": [r]}
        nr = {"panels": [norm_panel(p, notes, f"{where}.rows[{ri}].panels[{pi}]") for pi, p in enumerate(r.get("panels", []))]}
        for k in ("weight", "min_h"):
            if k in r: nr[k] = r[k]
        rows.append(nr)
    out = {"rows": rows}
    for k in ("notes", "font_scale", "max_stretch", "canvas"):
        if k in s: out[k] = s[k]
    return out

# ---------------------------------------------------------------- connected slides
def norm_graph(s, notes, where):
    out = {k: s[k] for k in ("title", "subtitle", "notes", "node_w", "node_h", "canvas", "font_scale", "title_align") if k in s}
    lanes = s.get("lanes") or s.get("swimlanes")
    if lanes: out["lanes"] = [str(l if isinstance(l, str) else pick(l, TITLE_KEYS, "")) for l in lanes]
    nodes, ids, by_title = [], set(), {}
    for k, n in enumerate(s.get("nodes") or s.get("boxes") or s.get("components") or []):
        if isinstance(n, str): n = {"title": n}
        nw = {}
        nw["title"] = str(pick(n, TITLE_KEYS, f"Node {k+1}"))
        nid = str(n.get("id") or slug(nw["title"]))
        while nid in ids: nid += "_2"
        ids.add(nid); nw["id"] = nid
        by_title[nw["title"].lower()] = nid; by_title[slug(nw["title"])] = nid
        if n.get("id") is not None: by_title[str(n["id"]).lower()] = nid
        d = pick(n, [k2 for k2 in DESC_KEYS if k2 != "text"])
        if d: nw["desc"] = str(d)
        shape = SHAPE_ALIASES.get(str(n.get("shape") or n.get("kind") or n.get("type") or "card").lower(), "card")
        nw["shape"] = shape
        if shape in ("card", "banner"):
            nw["icon"] = resolve_icon(n.get("icon"), f"{nw['title']}\x00{d or ''}", notes, f"{where}.nodes[{k}]")
        h = n.get("hue") or n.get("color")
        if h and HUE_ALIASES.get(str(h).lower()): nw["hue"] = HUE_ALIASES[str(h).lower()]
        col = pick(n, ["col", "column", "x"]); row = pick(n, ["row", "y"])
        if col is not None: nw["col"] = int(col)
        if row is not None: nw["row"] = int(row)
        if col is not None and row is not None: nw["_fixed"] = True
        lane = n.get("lane") or n.get("swimlane") or n.get("role") or n.get("actor")
        if lane is not None and lanes:
            names = [l.lower() for l in out["lanes"]]
            if isinstance(lane, int): nw["row"] = lane
            elif str(lane).lower() in names: nw["row"] = names.index(str(lane).lower())
            else:
                m = difflib.get_close_matches(str(lane).lower(), names, n=1, cutoff=0.5)
                if m: nw["row"] = names.index(m[0]); notes.append(f"{where}.nodes[{k}]: lane '{lane}' -> '{out['lanes'][nw['row']]}'")
                else: notes.append(f"{where}.nodes[{k}]: lane '{lane}' not in lanes; placed automatically")
        if "w" in n: nw["w"] = float(n["w"])
        nodes.append(nw)
    edges = []
    for k, e in enumerate(s.get("edges") or s.get("links") or s.get("connections") or s.get("arrows") or []):
        if isinstance(e, (list, tuple)): e = {"from": e[0], "to": e[1], **({"label": e[2]} if len(e) > 2 else {})}
        a = pick(e, ["from", "source", "src", "start", "origin"]); b = pick(e, ["to", "target", "dst", "dest", "end"])
        ra = by_title.get(str(a).lower()) or by_title.get(slug(a)); rb = by_title.get(str(b).lower()) or by_title.get(slug(b))
        if not ra or not rb:
            notes.append(f"{where}.edges[{k}]: unknown node '{a if not ra else b}' (edge dropped)"); continue
        ne = {"from": ra, "to": rb}
        lab = pick(e, ["label", "text", "name", "protocol", "data"])
        if lab: ne["label"] = str(lab)
        style = str(e.get("style") or e.get("line") or "").lower()
        if e.get("dashed") or "dash" in style or "dot" in style or str(e.get("type", "")).lower() in ("async", "dependency", "dashed", "optional"):
            ne["dashed"] = True
        if style in ("curve", "curved"): ne["style"] = "curve"
        arrow = str(e.get("arrow") or e.get("direction") or "end").lower()
        ne["arrow"] = {"both": "both", "bidirectional": "both", "two-way": "both", "none": "none", "undirected": "none"}.get(arrow, "end")
        if e.get("bidirectional"): ne["arrow"] = "both"
        for k2 in ("from_side", "to_side", "color", "width"):
            if k2 in e: ne[k2] = e[k2]
        edges.append(ne)
    out["nodes"], out["edges"] = nodes, edges
    placed = [n for n in nodes if "col" in n and n.get("_fixed")]
    if placed and len(placed) < len(nodes):
        for n in nodes:
            n.pop("col", None); n.pop("_fixed", None)
            if not out.get("lanes"): n.pop("row", None)
        notes.append(f"{where}: ignored col/row on {len(placed)} of {len(nodes)} nodes (set them on all nodes or none); auto-layout used")
    elif any("col" in n and not n.get("_fixed") for n in nodes) and not out.get("lanes"):
        for n in nodes: n.pop("col", None); n.pop("row", None)
        notes.append(f"{where}: ignored partial col/row; auto-layout used")
    auto_layout(out, notes, where)
    zones = []
    for k, z in enumerate(s.get("zones") or s.get("groups") or s.get("boundaries") or []):
        nz = {"title": str(pick(z, TITLE_KEYS, "Zone"))}
        hz = HUE_ALIASES.get(str(z.get("hue") or z.get("color") or "slate").lower(), "slate"); nz["hue"] = hz
        members = z.get("nodes") or z.get("members") or z.get("contains")
        if members:
            mids = [by_title.get(str(m).lower()) or by_title.get(slug(m)) for m in members]
            mn = [n for n in out["nodes"] if n["id"] in mids]
            if not mn: notes.append(f"{where}.zones[{k}]: no matching nodes (zone dropped)"); continue
            nz["cols"] = [min(n["col"] for n in mn), max(n["col"] for n in mn)]
            nz["rows"] = [min(n["row"] for n in mn), max(n["row"] for n in mn)]
            intruders = [n for n in out["nodes"] if n["id"] not in mids and nz["cols"][0] <= n["col"] <= nz["cols"][1] and nz["rows"][0] <= n["row"] <= nz["rows"][1]]
            for n in intruders:
                if out.get("lanes") or n.get("_fixed"):
                    notes.append(f"{where}.zones[{k}]: WARNING zone '{nz['title']}' also covers '{n['title']}'; give it a col/row outside the zone")
                    continue
                n["row"] = nz["rows"][1] + 1   # drop it below the zone
                while any(m is not n and m["col"] == n["col"] and m["row"] == n["row"] for m in out["nodes"]): n["row"] += 1
                out["rows"] = max(out.get("rows", 0), n["row"] + 1)
                notes.append(f"{where}.zones[{k}]: moved '{n['title']}' out of zone '{nz['title']}'")
        else:
            nz["cols"] = list(z.get("cols", [0, 0])); nz["rows"] = list(z.get("rows", [0, 0]))
        zones.append(nz)
    if zones: out["zones"] = zones
    if "node_w" not in out and not out.get("lanes") and any("label" in e for e in edges): out["node_w"] = 0.74
    for n in out["nodes"]: n.pop("_fixed", None)
    return out


def auto_layout(g, notes, where):
    """Assign col/row to nodes that lack them: columns by longest path (back edges ignored),
    rows by lane or by barycenter ordering. Nodes with explicit col/row are kept."""
    nodes = g["nodes"]
    if all("col" in n and "row" in n for n in nodes):
        g.setdefault("cols", max(n["col"] for n in nodes) + 1)
        if not g.get("lanes"): g.setdefault("rows", max(n["row"] for n in nodes) + 1)
        return
    ids = [n["id"] for n in nodes]; idx = {i: k for k, i in enumerate(ids)}
    succ = {i: [] for i in ids}; pred = {i: [] for i in ids}
    # drop back edges with DFS (order = node listing order)
    state, back = {}, set()
    def dfs(u):
        state[u] = 1
        for e in g["edges"]:
            if e["from"] != u: continue
            v = e["to"]
            if state.get(v) == 1: back.add(id(e))
            elif v not in state: dfs(v)
        state[u] = 2
    for i in ids:
        if i not in state: dfs(i)
    for e in g["edges"]:
        if id(e) in back or e["from"] == e["to"]: continue
        succ[e["from"]].append(e["to"]); pred[e["to"]].append(e["from"])
    rank = {}
    def r(u, seen=()):
        if u in rank: return rank[u]
        rank[u] = 0 if not pred[u] else 1 + max(r(p) for p in pred[u])
        return rank[u]
    for i in ids: r(i)
    lanes = g.get("lanes")
    for n in nodes:
        if "col" not in n: n["col"] = rank[n["id"]]
    # feedback nodes (every outgoing line loops back): place them just after the step they return to
    back_targets = {}
    for e in g["edges"]:
        if id(e) in back: back_targets.setdefault(e["from"], []).append(e["to"])
    byid = {n["id"]: n for n in nodes}
    for nid, tg in back_targets.items():
        n = byid[nid]
        if n.get("_fixed") or any(e["from"] == nid and id(e) not in back for e in g["edges"]): continue
        desired = min(byid[t]["col"] for t in tg) + 1
        if desired >= n["col"]: continue
        for c in range(desired, n["col"]):
            clash = lanes and "row" in n and any(m is not n and m["col"] == c and m.get("row") == n["row"] for m in nodes)
            if not clash:
                n["col"] = c; break
    if lanes:
        for n in nodes:
            if "row" not in n:
                ps = [m for m in nodes if m["id"] in pred[n["id"]] and "row" in m]
                n["row"] = ps[0]["row"] if ps else 0
        # resolve collisions inside a lane by pushing later nodes right (and their descendants)
        changed = True
        while changed:
            changed = False
            seen = {}
            for n in sorted(nodes, key=lambda m: (m["col"], idx[m["id"]])):
                key = (n["col"], n["row"])
                if key in seen:
                    shift(n["id"], nodes, succ); changed = True; break
                seen[key] = n["id"]
    else:
        cols = {}
        for n in nodes: cols.setdefault(n["col"], []).append(n)
        rowpos = {}
        for c in sorted(cols):
            group = cols[c]
            def bary(n):
                ps = [rowpos[p] for p in pred[n["id"]] if p in rowpos]
                return (sum(ps) / len(ps)) if ps else idx[n["id"]] * 0.01
            group.sort(key=lambda n: (bary(n), idx[n["id"]]))
            used = set(m["row"] for m in group if "row" in m)
            nxt = 0
            for n in group:
                if "row" not in n:
                    want = int(round(bary(n))) if any(p in rowpos for p in pred[n["id"]]) else nxt
                    while want in used: want += 1
                    n["row"] = want; used.add(want); nxt = want + 1
                rowpos[n["id"]] = n["row"]
    g["cols"] = max(g.get("cols", 0), max(n["col"] for n in nodes) + 1)
    if not lanes: g["rows"] = max(g.get("rows", 0), max(n["row"] for n in nodes) + 1)
    notes.append(f"{where}: auto-layout placed {sum(1 for _ in nodes)} nodes on a {g['cols']} x {g.get('rows', len(lanes or []))} grid")


def shift(nid, nodes, succ, seen=None):
    seen = seen or set()
    if nid in seen: return
    seen.add(nid)
    for n in nodes:
        if n["id"] == nid: n["col"] += 1
    for v in succ.get(nid, []):
        a = next(n for n in nodes if n["id"] == nid); b = next(n for n in nodes if n["id"] == v)
        if b["col"] <= a["col"]: shift(v, nodes, succ, seen)

# ---------------------------------------------------------------- top level
def normalize(raw):
    notes = []
    spec = loads_tolerant(raw) if isinstance(raw, str) else raw
    if isinstance(spec, list): spec = {"slides": spec}
    slides = spec.get("slides") or spec.get("pages") or [spec]
    out = {}
    theme = spec.get("theme")
    if theme:
        th = {}
        for k, v in theme.items():
            if isinstance(v, str) and re.fullmatch(r"#?[0-9a-fA-F]{6}", v.strip()): v = v.strip().lstrip("#").upper()
            th[k] = v
        out["theme"] = th
    ns = []
    for i, s in enumerate(slides):
        w = f"slides[{i}]"
        if any(k in s for k in ("nodes", "boxes", "components")) and not isinstance(s.get("rows"), list):
            ns.append(norm_graph(s, notes, w))
        else:
            ns.append(norm_layered(s, notes, w))
    out["slides"] = ns
    return out, notes


if __name__ == "__main__":
    import sys
    spec, notes = normalize(open(sys.argv[1]).read())
    print(json.dumps(spec, indent=1)); [print("NOTE:", n, file=sys.stderr) for n in notes]
