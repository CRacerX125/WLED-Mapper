
import json
from pathlib import Path
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, colorchooser
import re
# ---------- Submodel keyword controls (editable & persisted) ----------
from pathlib import Path as _KPath

_SUB_KEYWORDS_DEFAULT = [
    "outline", "inside", "outside",
    "middle", "center",
    "eyes", "mouth", "teeth", "nose", "stem",
    "bow", "wings", "body",
    "left", "right", "hourglass",
    "legs", "leg",
]
_SUB_KW_PATH = _KPath(__file__).with_name("sub_keywords.txt")

def _load_sub_keywords():
    try:
        rows = _SUB_KW_PATH.read_text(encoding="utf-8").splitlines()
        kws = [r.strip().lower() for r in rows if r.strip()]
        if kws:
            return set(kws)
    except Exception:
        pass
    return set(k.lower() for k in _SUB_KEYWORDS_DEFAULT)

def _save_sub_keywords(kws):
    try:
        _SUB_KW_PATH.write_text("\n".join(sorted(kws)), encoding="utf-8")
    except Exception:
        pass

SUB_KEYWORDS = _load_sub_keywords()
# ---------------------------------------------------------------------


_missing = []
try:
    import pandas as pd
except Exception:
    _missing.append("pandas")
try:
    import xlsxwriter  # noqa: F401
except Exception:
    pass

# ---------------------- WLED metadata ----------------------

WLED_EFFECTS = [
    (0, "Solid"),
    (1, "Blink"),
    (2, "Breathe"),
    (3, "Wipe"),
    (4, "Wipe Random"),
    (5, "Random Colors"),
    (6, "Sweep"),
    (7, "Dynamic"),
    (8, "Colorloop"),
    (9, "Rainbow"),
    (10, "Scan"),
    (11, "Scan Dual"),
    (12, "Fade"),
    (13, "Theater"),
    (14, "Theater Rainbow"),
    (15, "Running"),
    (16, "Saw"),
    (17, "Twinkle"),
    (18, "Dissolve"),
    (19, "Dissolve Rnd"),
    (20, "Sparkle"),
    (21, "Sparkle Dark"),
    (22, "Sparkle+"),
    (23, "Strobe"),
    (24, "Strobe Rainbow"),
    (25, "Strobe Mega"),
    (26, "Blink Rainbow"),
    (27, "Android"),
    (28, "Chase"),
    (29, "Chase Random"),
    (30, "Chase Rainbow"),
    (31, "Chase Flash"),
    (32, "Chase Flash Rnd"),
    (33, "Rainbow Runner"),
    (34, "Colorful"),
    (35, "Traffic Light"),
    (36, "Sweep Random"),
    (37, "Chase 2"),
    (38, "Aurora"),
    (39, "Stream"),
    (40, "Scanner"),
    (41, "Lighthouse"),
    (42, "Fireworks"),
    (43, "Rain"),
    (44, "Tetrix"),
    (45, "Fire Flicker"),
    (46, "Gradient"),
    (47, "Loading"),
    (49, "Fairy"),
    (50, "Two Dots"),
    (51, "Fairytwinkle"),
    (52, "Running Dual"),
    (54, "Chase 3"),
    (55, "Tri Wipe"),
    (56, "Tri Fade"),
    (57, "Lightning"),
    (58, "ICU"),
    (59, "Multi Comet"),
    (60, "Scanner Dual"),
    (61, "Stream 2"),
    (62, "Oscillate"),
    (63, "Pride 2015"),
    (64, "Juggle"),
    (65, "Palette"),
    (66, "Fire 2012"),
    (67, "Colorwaves"),
    (68, "Bpm"),
    (69, "Fill Noise"),
    (70, "Noise 1"),
    (71, "Noise 2"),
    (72, "Noise 3"),
    (73, "Noise 4"),
    (74, "Colortwinkles"),
    (75, "Lake"),
    (76, "Meteor"),
    (77, "Meteor Smooth"),
    (78, "Railway"),
    (79, "Ripple"),
    (80, "Twinklefox"),
    (81, "Twinklecat"),
    (82, "Halloween Eyes"),
    (83, "Solid Pattern"),
    (84, "Solid Pattern Tri"),
    (85, "Spots"),
    (86, "Spots Fade"),
    (87, "Glitter"),
    (88, "Candle"),
    (89, "Fireworks Starburst"),
    (90, "Fireworks 1D"),
    (91, "Bouncing Balls"),
    (92, "Sinelon"),
    (93, "Sinelon Dual"),
    (94, "Sinelon Rainbow"),
    (95, "Popcorn"),
    (96, "Drip"),
    (97, "Plasma"),
    (98, "Percent"),
    (99, "Ripple Rainbow"),
    (100, "Heartbeat"),
    (101, "Pacifica"),
    (102, "Candle Multi"),
    (103, "Solid Glitter"),
    (104, "Sunrise"),
    (105, "Phased"),
    (106, "Twinkleup"),
    (107, "Noise Pal"),
    (108, "Sine"),
    (109, "Phased Noise"),
    (110, "Flow"),
    (111, "Chunchun"),
    (112, "Dancing Shadows"),
    (113, "Washing Machine"),
    (115, "Blends"),
    (116, "TV Simulator"),
    (117, "Dynamic Smooth"),
    (118, "Spaceships"),
    (119, "Crazy Bees"),
    (120, "Ghost Rider"),
    (121, "Blobs"),
    (122, "Scrolling Text"),
    (123, "Drift Rose"),
    (124, "Distortion Waves"),
    (125, "Soap"),
    (126, "Octopus"),
    (127, "Waving Cell"),
    (128, "Pixels"),
    (129, "Pixelwave"),
    (130, "Juggles"),
    (131, "Matripix"),
    (132, "Gravimeter"),
    (133, "Plasmoid"),
    (134, "Puddles"),
    (135, "Midnoise"),
    (136, "Noisemeter"),
    (137, "Freqwave"),
    (138, "Freqmatrix"),
    (139, "GEQ"),
    (140, "Waterfall"),
    (141, "Freqpixels"),
    (143, "Noisefire"),
    (144, "Puddlepeak"),
    (145, "Noisemove"),
    (146, "Noise2D"),
    (147, "Perlin Move"),
    (148, "Ripple Peak"),
    (149, "Firenoise"),
    (150, "Squared Swirl"),
    (152, "DNA"),
    (153, "Matrix"),
    (154, "Metaballs"),
    (155, "Freqmap"),
    (156, "Gravcenter"),
    (157, "Gravcentric"),
    (158, "Gravfreq"),
    (159, "DJ Light"),
    (160, "Funky Plank"),
    (162, "Pulser"),
    (163, "Blurz"),
    (164, "Drift"),
    (165, "Waverly"),
    (166, "Sun Radiation"),
    (167, "Colored Bursts"),
    (168, "Julia"),
    (172, "Game Of Life"),
    (173, "Tartan"),
    (174, "Polar Lights"),
    (175, "Swirl"),
    (176, "Lissajous"),
    (177, "Frizzles"),
    (178, "Plasma Ball"),
    (179, "Flow Stripe"),
    (180, "Hiphotic"),
    (181, "Sindots"),
    (182, "DNA Spiral"),
    (183, "Black Hole"),
    (184, "Wavesins"),
    (185, "Rocktaves"),
    (186, "Akemi")
]

WLED_PALETTES = [
    (0, "Default"),
    (1, "Random Cycle"),
    (2, "Color 1"),
    (3, "Colors 1&2"),
    (4, "Color Gradient"),
    (5, "Colors Only"),
    (6, "Party"),
    (7, "Cloud"),
    (8, "Lava"),
    (9, "Ocean"),
    (10, "Forest"),
    (11, "Rainbow"),
    (12, "Rainbow Bands"),
    (13, "Sunset"),
    (14, "Rivendell"),
    (15, "Breeze"),
    (16, "Red & Blue"),
    (17, "Yellowout"),
    (18, "Analogous"),
    (19, "Splash"),
    (20, "Pastel"),
    (21, "Sunset 2"),
    (22, "Beach"),
    (23, "Vintage"),
    (24, "Departure"),
    (25, "Landscape"),
    (26, "Beech"),
    (27, "Sherbet"),
    (28, "Hult"),
    (29, "Hult 64"),
    (30, "Drywet"),
    (31, "Jul"),
    (32, "Grintage"),
    (33, "Rewhi"),
    (34, "Tertiary"),
    (35, "Fire"),
    (36, "Icefire"),
    (37, "Cyane"),
    (38, "Light Pink"),
    (39, "Autumn"),
    (40, "Magenta"),
    (41, "Magred"),
    (42, "Yelmag"),
    (43, "Yelblu"),
    (44, "Orange & Teal"),
    (45, "Tiamat"),
    (46, "April Night"),
    (47, "Orangery"),
    (48, "C9"),
    (49, "Sakura"),
    (50, "Aurora"),
    (51, "Atlantica"),
    (52, "C9 2"),
    (53, "C9 New"),
    (54, "Temperature"),
    (55, "Aurora 2"),
    (56, "Retro Clown"),
    (57, "Candy"),
    (58, "Toxy Reaf"),
    (59, "Fairy Reaf"),
    (60, "Semi Blue"),
    (61, "Pink Candy"),
    (62, "Red Reaf"),
    (63, "Aqua Flash"),
    (64, "Yelblu Hot"),
    (65, "Lite Light"),
    (66, "Red Flash"),
    (67, "Blink Red"),
    (68, "Red Shift"),
    (69, "Red Tide"),
    (70, "Candy2")
]

EFFECT_NAME_TO_ID = {n: i for (i, n) in WLED_EFFECTS}
PALETTE_NAME_TO_ID = {n: i for (i, n) in WLED_PALETTES}
EFFECT_ID_TO_NAME = {i: n for (i, n) in WLED_EFFECTS}
PALETTE_ID_TO_NAME = {i: n for (i, n) in WLED_PALETTES}

# ---------------------- parsing / helpers ----------------------

def load_first_preset_segments(preset_path: Path):
    with preset_path.open("r", encoding="utf-8") as f:
        presets = json.load(f)
    for key, val in presets.items():
        if isinstance(val, dict) and "seg" in val:
            segs = val["seg"]
            _ = val.get("n", str(key))
            norm = []
            for s in segs:
                start = int(s.get("start", 0))
                stop = int(s.get("stop", start))
                nm = s.get("n", "")
                norm.append({"start": start, "stop": stop, "name": nm})
            norm.sort(key=lambda x: x["start"])
            return norm, _
    raise RuntimeError("No presets with 'seg' found.")

def _normalize_spaces(s: str) -> str:
    s = (s or "")
    s = s.replace("_", " ").replace("-", " ")
    s = s.replace("—", " ").replace("–", " ")
    s = re.sub(r"\s+", " ", s)
    return s.strip()

def _titlecase_preserve_digits(s: str) -> str:
    parts = s.split()
    out = []
    for p in parts:
        out.append(p if p.isdigit() else p[:1].upper() + p[1:].lower())
    return " ".join(out)

def split_prop_sub_hardcoded(segment_name: str, aliases: dict):
    seg = (segment_name or "").strip()
    if not seg:
        return "(no name)", "(all)"
    # explicit separators first
    for delim in (" — ", " – ", " - ", " : "):
        if delim in seg:
            base, sub = [t.strip() for t in seg.split(delim, 1)]
            # If base ends with Inside/Outside, fold that token into the sub name
            btoks = base.split()
            if len(btoks) >= 2 and btoks[-1].lower() in ("inside","outside"):
                base = " ".join(btoks[:-1]).strip()
                sub = f"{btoks[-1].capitalize()} {sub}".strip()
            base = aliases.get(base, base) if aliases else base
            return base, (sub or "(all)")
    # heuristic without explicit delimiter:
    parts = seg.split()
    if parts:
        # Case: "... Inside/Outside X"  -> base="...", sub="Inside X"
        if len(parts) >= 3 and parts[-2].lower() in ("inside","outside"):
            base = " ".join(parts[:-2]).strip()
            sub = f"{parts[-2].capitalize()} {parts[-1]}"
            base = aliases.get(base, base) if aliases else base
            return base, sub
        # Case: "... <keyword>" where keyword is in SUB_KEYWORDS
        last = parts[-1].lower()
        if last in SUB_KEYWORDS and len(parts) >= 2:
            base = " ".join(parts[:-1]).strip()
            base = aliases.get(base, base) if aliases else base
            sub = parts[-1]
            sub = (sub[:1].upper() + sub[1:]) if not sub.isdigit() else sub
            return base, sub
    base = aliases.get(seg, seg) if aliases else seg
    return base, "(all)"
def build_prop_structure(segments, aliases: dict):
    """
    Build props as *sequential runs* of the same base name.
    Example global order:
        House — Left, House — Middle, House — Right, Pumpkin — Outline, House — Trim
    becomes UI props:
        "House", "Pumpkin", "House (2)"
    Each run contains only the submodels that appear contiguously in that block.
    """
    from collections import defaultdict, OrderedDict
    # Normalize into a list of (start, stop, base_prop, sub, full_name) sorted by 'start'
    items = []
    for s in segments:
        base, sub = split_prop_sub_hardcoded(s["name"], aliases)
        items.append((int(s["start"]), int(s["stop"]), base, sub, s["name"]))
    items.sort(key=lambda t: t[0])  # by start

    run_counts = defaultdict(int)
    props_runs = []  # list of dicts in UI order

    last_base = None
    for start, stop, base, sub, full in items:
        if base != last_base:
            # starting a new run for this base
            run_counts[base] += 1
            run_idx = run_counts[base]
            display = base if run_idx == 1 else f"{base} ({run_idx})"
            props_runs.append({"display": display, "base": base, "run": run_idx,
                               "min_start": start, "segments": []})
        props_runs[-1]["segments"].append({
            "start": start, "stop": stop, "segment_name": full, "submodel": sub
        })
        # keep min_start minimal inside the run (for completeness)
        if start < props_runs[-1]["min_start"]:
            props_runs[-1]["min_start"] = start
        last_base = base

    # Construct the final props dict keyed by display name in sequential order
    props = OrderedDict((r["display"], {"segments": r["segments"], "min_start": r["min_start"]})
                        for r in props_runs)
    return props

def contiguous_ranges(sorted_ints):
    """Given a sorted list of unique ints, return list of [start, stop_exclusive] contiguous ranges (1-based input)."""
    if not sorted_ints:
        return []
    ranges = []
    start = prev = sorted_ints[0]
    for v in sorted_ints[1:]:
        if v == prev + 1:
            prev = v
            continue
        ranges.append([start, prev + 1])  # 1-based stop exclusive
        start = prev = v
    ranges.append([start, prev + 1])
    return ranges

# ---------------------- tooltip ----------------------

class TreeTooltip:
    def __init__(self, widget):
        self.widget = widget
        self.tip = None
        self.visible_for_iid = None

    def show(self, text, x_root, y_root):
        self.hide()
        self.tip = tk.Toplevel(self.widget)
        self.tip.wm_overrideredirect(True)
        self.tip.wm_geometry(f"+{x_root + 16}+{y_root + 16}")
        label = tk.Label(self.tip, text=text, justify="left", relief="solid", borderwidth=1, background="#FFF9E6")
        label.pack(ipadx=6, ipady=4)

    def hide(self):
        if self.tip is not None:
            try:
                self.tip.destroy()
            except Exception:
                pass
        self.tip = None
        self.visible_for_iid = None

# ---------------------- GUI ----------------------

    # --- Loaded-sequence star badge helpers ---
    def _strip_star_prefix(self, s):
        try:
            s = str(s)
        except Exception:
            return s
        for prefix in ("[LOADED] ", "★ ", "* "):
            if s.startswith(prefix):
                return s[len(prefix):]
        return s

def _apply_loaded_star_to_list(self, lb):
        """Show a leading '[LOADED] ' for the loaded sequence in the given listbox, without changing selection."""
        try:
            if not lb or not str(lb):
                return
        except Exception:
            return
        loaded = self._get_loaded_for_list(lb) if hasattr(self, "_get_loaded_for_list") else getattr(self, "_loaded_seq_name_editor", None)
        try:
            count = lb.size()
        except Exception:
            count = 0
        try:
            sel = lb.curselection()
        except Exception:
            sel = ()
        for i in range(count):
            try:
                t = lb.get(i)
            except Exception:
                continue
            base = self._strip_star_prefix(t)
            display = f"[LOADED] {base}" if (loaded and str(base) == str(loaded)) else base
            if display != t:
                try:
                    lb.delete(i)
                    lb.insert(i, display)
                except Exception:
                    pass
        try:
            if sel:
                lb.selection_clear(0, "end")
                for idx in sel:
                    lb.selection_set(idx)
        except Exception:
            pass

class GeneratorGUI(tk.Tk):

    def _ensure_props_tree_populated(self):
        """Populate left props list if currently empty; used when Map tab Load is pressed."""
        try:
            if not hasattr(self, "props_tree") or not hasattr(self, "prop_order"):
                return
            items = self.props_tree.get_children()
            if items:
                return
            for prop in getattr(self, "prop_order", []):
                data = self.props.get(prop, {"segments":[]})
                prop_leds = sum(s.get("stop",0) - s.get("start",0) for s in data.get("segments", []))
                iid = self.props_tree.insert("", "end", text=prop, values=(prop_leds, ""))
                if not hasattr(self, "_tree_iids_by_prop"): self._tree_iids_by_prop = {}
                self._tree_iids_by_prop[prop] = iid
        except Exception:
            pass

    def _split_prop_sub(self, nm: str):
        """
        Split a 'Prop — Sub' display name into (prop, sub).
        Adds heuristics so names like 'G1 Inside Wings' -> ('G1', 'Inside Wings').
        """
        try:
            s = (nm or "").strip()
            # Normalize unicode dashes to ASCII hyphen for detection, but prefer explicit em-dash split first
            if "—" in s:
                parts = [p.strip() for p in s.split("—", 1)]
                if len(parts) == 2:
                    return parts[0], parts[1]
            if " - " in s:
                parts = [p.strip() for p in s.split(" - ", 1)]
                if len(parts) == 2:
                    return parts[0], parts[1]
            if "-" in s:
                parts = [p.strip() for p in s.split("-", 1)]
                if len(parts) == 2:
                    return parts[0], parts[1]
            # Heuristic: "<PROP> Inside/Outside <rest>"
            toks = s.split()
            if len(toks) >= 3 and toks[1].lower() in ("inside", "outside"):
                prop = toks[0]
                sub = " ".join(toks[1:])
                return prop, sub
            # Fallback: if there is a space, take the first token as prop
            if " " in s:
                first, rest = s.split(" ", 1)
                return first.strip(), rest.strip()
            return s, None
        except Exception:
            return s, None
    def _render_loaded_badges(self):
        """Apply [LOADED] label logic to both lists without altering selections."""
        try:
            self._rewrite_list_with_badges(getattr(self, "seq_list", None))
        except Exception:
            pass
        try:
            self._rewrite_list_with_badges(getattr(self, "map_seq_list", None))
        except Exception:
            pass

    def _get_loaded_for_list(self, lb):
        """Return the loaded sequence name appropriate for the given listbox."""
        try:
            if lb is getattr(self, "seq_list", None):
                return getattr(self, "_loaded_seq_name_editor", None)
            if lb is getattr(self, "map_seq_list", None):
                return getattr(self, "_loaded_seq_name_map", None)
        except Exception:
            pass
        return getattr(self, "_loaded_seq_name_editor", None)

    def _display_label_for_seq_list(self, name: str):
        base = self._strip_badge_prefix(name) if hasattr(self, "_strip_badge_prefix") else name
        loaded = getattr(self, "_loaded_seq_name_editor", None)
        if loaded is not None and str(base) == str(loaded):
            return f"[LOADED] {base}"
        return base

    def _display_label_for_map_list(self, name: str):
        base = self._strip_badge_prefix(name) if hasattr(self, "_strip_badge_prefix") else name
        loaded = getattr(self, "_loaded_seq_name_map", None)
        if loaded is not None and str(base) == str(loaded):
            return f"[LOADED] {base}"
        return base

    # --- Loaded badge helpers ---

    def _strip_badge_prefix(self, s: str):

        try:

            s = str(s)

        except Exception:

            return s

        for prefix in ("[LOADED] ", "★ ", "* "):

            if s.startswith(prefix):

                return s[len(prefix):]

        return s



    def _display_label_for_name(self, name: str):

        base = self._strip_badge_prefix(name)

        loaded = self._get_loaded_for_list(lb) if hasattr(self, "_get_loaded_for_list") else getattr(self, "_loaded_seq_name_editor", None)

        if loaded is not None and str(base) == str(loaded):

            return f"[LOADED] {base}"

        return base



    def _rewrite_list_with_badges(self, lb):
        """Rewrite list items with [LOADED] prefix where appropriate, preserving selection."""
        try:
            if not lb or not str(lb):
                return
        except Exception:
            return
        try:
            sel = list(lb.curselection())
        except Exception:
            sel = []
        try:
            lb.delete(0, tk.END)
            for nm in (getattr(self, "seq_order", []) or []):
                if lb is getattr(self, "seq_list", None):
                    label = self._display_label_for_seq_list(nm)
                elif lb is getattr(self, "map_seq_list", None):
                    label = self._display_label_for_map_list(nm)
                else:
                    label = self._display_label_for_seq_list(nm)
                lb.insert(tk.END, label)
            if sel:
                for i in sel:
                    if 0 <= i < lb.size():
                        lb.selection_set(i)
        except Exception:
            pass


    def __init__(self):
        self._suppress_map_details = True
        super().__init__()
        self.title("WLED Generator v2.12.26")
        self.geometry("1280x880")
        self.segments = []
        self.props = {}          # prop -> {segments:[...], min_start:int}
        self.prop_order = []     # explicit prop ordering for UI
        self.aliases = {}
        self.manual_map = {}     # (prop, sub) -> list[int] (1-based GLOBAL LEDs)
        self.prop_tooltips = {}
        self.sub_tooltips = {}   # iid -> text for right pane tooltip
        self.suppressed = {}     # prop -> set(submodels) hidden autos
        self.sub_order = {}      # prop -> list[submodel] display order

        # Sequences: name -> {"bri": int, "include": set((prop, sub)),
        #                     "fx": int, "sx": int, "ix": int, "pal": int,
        #                     "col": [[r,g,b],[r2,g2,b2],[r3,g3,b3]]}
        self.sequences = {}
        self.seq_order = []

        self._build_ui()

    def _build_ui(self):
        self.notebook = ttk.Notebook(self)
        # Keep loaded info persistent across tab switches
        def _on_tab_change(_e=None):
            try:
                tab = self.notebook.tab(self.notebook.select(), "text")
            except Exception:
                tab = ""
            try:
                if "Sequences" in tab or "Presets" in tab:
                    # Repaint sequences editor and table from stored loaded name
                    if getattr(self, "_loaded_seq_name_editor", None):
                        try:
                            self._refresh_included_table_loaded()
                        except Exception:
                            pass
                elif "Mapping" in tab or "Submodels" in tab:
                    if getattr(self, "_loaded_seq_name_map", None):
                        try:
                            self._refresh_map_table_loaded()
                        except Exception:
                            pass
            except Exception:
                pass
        try:
            self.notebook.bind("<<NotebookTabChanged>>", _on_tab_change)
        except Exception:
            pass
        self.notebook.pack(fill="both", expand=True)

        # ---------- Tab: Mapping ----------
        map_tab = ttk.Frame(self.notebook)
        self.notebook.add(map_tab, text="Mapping / Submodels")

        pad = {"padx": 8, "pady": 6}
        top = ttk.Frame(map_tab); top.pack(fill="x")
        ttk.Label(top, text="Presets JSON:").grid(row=0, column=0, sticky="e", **pad)
        self.presets_var = tk.StringVar()
        ttk.Entry(top, textvariable=self.presets_var, width=60).grid(row=0, column=1, sticky="we", **pad)
        ttk.Button(top, text="Browse...", command=self.browse_presets).grid(row=0, column=2, **pad)
        ttk.Button(top, text="Submodel Keywords…", command=lambda: open_subkeyword_editor(self)).grid(row=0, column=5, **pad)

        mid = ttk.Frame(map_tab); mid.pack(fill="both", expand=True, padx=8, pady=8)
        mid.columnconfigure(0, weight=1); mid.columnconfigure(1, weight=1)
        mid.rowconfigure(0, weight=1)

        left = ttk.LabelFrame(mid, text="Props — select a prop to manage its submodels")
        left.grid(row=0, column=0, sticky="nsew", padx=6, pady=6)
        # --- Sequences (same look/feel as Sequences / Presets) ---
        seq_lf = ttk.LabelFrame(left, text="Sequences")
        seq_lf.pack(fill="x", padx=6, pady=(6, 0))
        seq_wrap = ttk.Frame(seq_lf); seq_wrap.pack(fill="x", padx=6, pady=6)
        self.map_seq_list = tk.Listbox(seq_wrap, height=6, exportselection=False, activestyle="none")
        self.map_seq_list.pack(side="left", fill="x", expand=True)
        # Persistent loaded highlight for Mapping list
        try:
            self.after(300, self._map_loaded_highlight_tick)
        except Exception:
            pass
        sby = ttk.Scrollbar(seq_wrap, orient="vertical", command=self.map_seq_list.yview)
        sby.pack(side="left", fill="y")
        self.map_seq_list.configure(yscrollcommand=sby.set)
        self.map_seq_list.bind("<<ListboxSelect>>", self._on_map_seq_select)
        self.map_seq_list.bind("<Button-1>", lambda e: self._handle_seq_list_click(e, "map"))
        try:
            self.map_seq_list.unbind("<Up>"); self.map_seq_list.unbind("<Down>")
        except Exception:
            pass
        self.map_seq_list.bind("<Up>",   lambda e: self._on_seq_arrow(e, "map"))
        self.map_seq_list.bind("<Down>", lambda e: self._on_seq_arrow(e, "map"))
        # Controls (reuse existing handlers)
        seq_btns = ttk.Frame(seq_lf); seq_btns.pack(fill="x", padx=6, pady=(0,6))
        # Mapping: explicit Load button (left of Add)
        self._btn_map_load = ttk.Button(seq_btns, text="Load", command=self.load_map_from_selection)
        self._btn_map_load.pack(side="left", padx=(0,6))
        ttk.Button(seq_btns, text="Add", command=lambda: (self._ensure_seq_selection_from_map(), self.add_sequence())).pack(side="left")
        ttk.Button(seq_btns, text="Duplicate", command=lambda: (self._ensure_seq_selection_from_map(), self.duplicate_sequence())).pack(side="left", padx=6)
        ttk.Button(seq_btns, text="Rename", command=lambda: (self._ensure_seq_selection_from_map(), self.rename_sequence())).pack(side="left", padx=6)
        ttk.Button(seq_btns, text="Delete", command=lambda: (self._ensure_seq_selection_from_map(), self.delete_sequence())).pack(side="left")
        ttk.Button(seq_btns, text="Move Up", command=lambda: (self._ensure_seq_selection_from_map(), self.move_sequence(-1))).pack(side="left", padx=6)
        ttk.Button(seq_btns, text="Move Down", command=lambda: (self._ensure_seq_selection_from_map(), self.move_sequence(+1))).pack(side="left")
        # Seed list if sequences are present
        try:
            self._sync_map_seq_list()
        except Exception:
            pass
        self.props_tree = ttk.Treeview(left, columns=("count","status"), show="tree headings", height=22, selectmode="browse")
        self.props_tree.heading("#0", text="Prop")
        self.props_tree.heading("count", text="LEDs")
        self.props_tree.heading("status", text="Validation")
        self.props_tree.column("#0", width=300)
        self.props_tree.column("count", width=80, anchor="e")
        self.props_tree.column("status", width=220, anchor="w")
        self.props_tree.pack(fill="both", expand=True, padx=6, pady=6)

        prop_actions = ttk.Frame(left); prop_actions.pack(fill="x", padx=6, pady=2)
        ttk.Button(prop_actions, text="Add prop", command=self.add_prop).pack(side="left", padx=4)
        ttk.Button(prop_actions, text="Delete prop", command=self.delete_prop).pack(side="left", padx=4)
        ttk.Button(prop_actions, text="▲ Move up", command=lambda: self.move_prop(-1)).pack(side="right", padx=4)
        ttk.Button(prop_actions, text="▼ Move down", command=lambda: self.move_prop(1)).pack(side="right", padx=4)

        right = ttk.LabelFrame(mid, text="Submodels (double-click to edit LEDs; use ▲/▼ to reorder)")
        right.grid(row=0, column=1, sticky="nsew", padx=6, pady=6)
        self.sub_tree = ttk.Treeview(right, columns=("sub","count","mode"), show="headings", height=18, selectmode="browse")
        self.sub_tree.heading("sub", text="Submodel")
        self.sub_tree.heading("count", text="LEDs")
        self.sub_tree.heading("mode", text="Source")
        self.sub_tree.column("sub", width=260, anchor="w")
        self.sub_tree.column("count", width=80, anchor="e")
        self.sub_tree.column("mode", width=120, anchor="center")
        self.sub_tree.pack(fill="both", expand=True, padx=6, pady=6)

        self.sub_tree.tag_configure("conflict", background="#FFECEC", foreground="#B00000")

        actions = ttk.Frame(right); actions.pack(fill="x", padx=6, pady=6)
        ttk.Button(actions, text="Add submodel", command=self.add_submodel).pack(side="left", padx=4)
        ttk.Button(actions, text="Delete submodel", command=self.delete_submodel).pack(side="left", padx=4)
        self.btn_up = ttk.Button(actions, text="▲ Move up", command=self.move_up)
        self.btn_up.pack(side="right", padx=4)
        self.btn_down = ttk.Button(actions, text="▼ Move down", command=self.move_down)
        self.btn_down.pack(side="right", padx=4)

        # tooltips
        self._prop_tooltip = TreeTooltip(self.props_tree)
        self._sub_tooltip = TreeTooltip(self.sub_tree)
        self.props_tree.bind("<Motion>", self._on_props_motion)
        self.props_tree.bind("<Leave>", lambda e: self._prop_tooltip.hide())
        self.sub_tree.bind("<Motion>", self._on_subs_motion)
        self.sub_tree.bind("<Leave>", lambda e: self._sub_tooltip.hide())

        bottom = ttk.Frame(map_tab); bottom.pack(fill="x", padx=8, pady=8)

        export_row = ttk.Frame(bottom)
        export_row.pack(anchor='center', pady=8)
        ttk.Button(bottom, text="Export ledmap.json", command=self.export_ledmap_json_clicked).pack(in_=export_row, side='left', padx=8)
        ttk.Button(bottom, text="Export Excel Detailed Breakdown", command=self.export_excel_clicked).pack(in_=export_row, side='left', padx=8)

        # status
        self.status = tk.Text(map_tab, height=7, state='disabled', cursor='arrow', takefocus=0); self.status.pack(fill="both", expand=False, padx=8, pady=8)

        # tags
        self._selected_prop = None
        self.props_tree.tag_configure("ok", foreground="green")
        self.props_tree.tag_configure("bad", foreground="red")
        self._update_move_buttons(False)

        # bindings
        self.props_tree.bind("<<TreeviewSelect>>", self.on_prop_select)
        self.sub_tree.bind("<Double-1>", self.on_sub_double_click)
        self.sub_tree.bind("<<TreeviewSelect>>", self.on_sub_select)

        # ---------- Tab: Sequences ----------
        seq_tab = ttk.Frame(self.notebook)
        self.notebook.add(seq_tab, text="Sequences / Presets")

        seq_pan = ttk.Frame(seq_tab)
        seq_pan.pack(fill="both", expand=True, padx=8, pady=8)
        seq_pan.columnconfigure(0, weight=2)
        seq_pan.columnconfigure(1, weight=3)
        seq_pan.rowconfigure(0, weight=1)
        seq_pan.rowconfigure(1, weight=2)

        # Sequences list
        left_seq = ttk.LabelFrame(seq_pan, text="Sequences")
        left_seq.grid(row=0, column=0, sticky="nsew", padx=6, pady=6)
        self.seq_list = tk.Listbox(left_seq, height=22, activestyle="none")
        self.seq_list.pack(side="left", fill="both", expand=True, padx=6, pady=6)
        # Keep the loaded highlight persistent across updates
        try:
            self.seq_list.bind("<<ListboxSelect>>", lambda e: self._highlight_loaded_sequence())
        except Exception:
            pass
        self.seq_list.bind("<<ListboxSelect>>", self.on_seq_select)
        self.seq_list.bind("<Button-1>", lambda e: self._handle_seq_list_click(e, "seq"))
        for _lb in (self.seq_list, getattr(self, 'map_seq_list', None)):
            try:
                if _lb is None: 
                    continue
                for _k in ("<Up>", "<Down>", "<Home>", "<End>", "<Prior>", "<Next>"):
                    _lb.bind(_k, lambda e: "break")
            except Exception:
                pass

        seq_btns = ttk.Frame(left_seq); seq_btns.pack(side="right", fill="y", padx=6, pady=6)
        ttk.Button(seq_btns, text="Load", command=self.load_editor_from_selection).pack(fill="x", pady=(0,4))
        ttk.Button(seq_btns, text="Add", command=self.add_sequence).pack(fill="x", pady=2)
        ttk.Button(seq_btns, text="Rename", command=self.rename_sequence).pack(fill="x", pady=2)
        ttk.Button(seq_btns, text="Delete", command=self.delete_sequence).pack(fill="x", pady=2)
        ttk.Button(seq_btns, text="▲", command=lambda: self.move_sequence(-1)).pack(fill="x", pady=8)
        ttk.Button(seq_btns, text="▼", command=lambda: self.move_sequence(1)).pack(fill="x")

        # Sequence editor
        right_seq = ttk.LabelFrame(seq_pan, text="Sequence Editor")
        right_seq.grid(row=0, column=1, sticky="nsew", padx=6, pady=6)

        self.right_seq = right_seq
        form = ttk.Frame(right_seq); form.pack(fill="x", padx=6, pady=6)
        ttk.Label(form, text="Name:").grid(row=0, column=0, sticky="e", padx=4, pady=4)
        self.seq_name_var = tk.StringVar()
        self.name_entry = ttk.Entry(form, textvariable=self.seq_name_var, width=28)
        try:
            self.name_entry.state(["readonly"])
        except Exception:
            self.name_entry.configure(state="readonly")
        self.name_entry.grid(row=0, column=1, sticky="w", padx=4, pady=4)

        ttk.Label(form, text="Brightness (1-255):").grid(row=1, column=0, sticky="e", padx=4, pady=4)
        self.seq_bri_var = tk.IntVar(value=128)
        self.bri_spin = ttk.Spinbox(form, from_=1, to=255, textvariable=self.seq_bri_var, width=8, wrap=False, command=self._on_brightness_spin); self.bri_spin.grid(row=1, column=1, sticky="w", padx=4, pady=4)

        # Effect options
        ef = ttk.Frame(right_seq); ef.pack(fill="x", padx=6, pady=2)
        ttk.Label(ef, text="Effect:").grid(row=0, column=0, sticky="e", padx=4, pady=4)
        self.fx_var = tk.StringVar(value="Solid")  # store display name OR numeric id
        self.fx_combo = ttk.Combobox(ef, textvariable=self.fx_var, values=[n for _, n in WLED_EFFECTS]+["(enter id)"], width=24)
        self.fx_combo.grid(row=0, column=1, sticky="w", padx=4, pady=4)
        ttk.Label(ef, text="Speed (0-255):").grid(row=0, column=2, sticky="e", padx=4, pady=4)
        self.sx_var = tk.IntVar(value=128)
        self.sx_spin = ttk.Spinbox(ef, from_=0, to=255, textvariable=self.sx_var, width=8, state="disabled")
        self.sx_spin.grid(row=0, column=3, sticky="w", padx=4, pady=4)
        ttk.Label(ef, text="Intensity (0-255):").grid(row=0, column=4, sticky="e", padx=4, pady=4)
        self.ix_var = tk.IntVar(value=128)
        ttk.Spinbox(ef, from_=0, to=255, textvariable=self.ix_var, width=8).grid(row=0, column=5, sticky="w", padx=4, pady=4)

        pf = ttk.Frame(right_seq); pf.pack(fill="x", padx=6, pady=2)
        ttk.Label(pf, text="Palette:").grid(row=0, column=0, sticky="e", padx=4, pady=4)
        self.pal_var = tk.StringVar(value="Default")
        self.pal_combo = ttk.Combobox(pf, textvariable=self.pal_var, values=[n for _, n in WLED_PALETTES]+["(enter id)"], width=24)
        self.pal_combo.grid(row=0, column=1, sticky="w", padx=4, pady=4)

        # When user chooses the special '(enter id)' option, explain how to type an ID
        try:
            self.fx_combo.bind('<<ComboboxSelected>>', self._on_fx_combo_selected)
            self.pal_combo.bind('<<ComboboxSelected>>', self._on_pal_combo_selected)
        except Exception:
            pass

        # Colors
        cf = ttk.LabelFrame(right_seq, text="Colors")
        cf.pack(fill="x", padx=6, pady=6)
        self.col_vars = [(tk.IntVar(value=255), tk.IntVar(value=160), tk.IntVar(value=0)),
                         (tk.IntVar(value=0), tk.IntVar(value=0), tk.IntVar(value=0)),
                         (tk.IntVar(value=0), tk.IntVar(value=0), tk.IntVar(value=0))]

        def mk_color_row(row, label):
            ttk.Label(cf, text=label).grid(row=row, column=0, sticky="e", padx=4, pady=4)
            r, g, b = self.col_vars[row-1]
            ent = ttk.Frame(cf)
            ent.grid(row=row, column=1, sticky="w", padx=4, pady=4)
            tk.Entry(ent, width=4, textvariable=r).pack(side="left")
            tk.Entry(ent, width=4, textvariable=g).pack(side="left", padx=2)
            tk.Entry(ent, width=4, textvariable=b).pack(side="left")
            def pick():
                initial = '#%02x%02x%02x' % (r.get(), g.get(), b.get())
                c = colorchooser.askcolor(color=initial, title=f"Pick {label}")
                if c and c[0]:
                    rr, gg, bb = map(int, c[0])
                    r.set(rr); g.set(gg); b.set(bb)
                    sw.configure(background='#%02x%02x%02x' % (rr, gg, bb))
            ttk.Button(cf, text="Pick…", command=pick).grid(row=row, column=2, sticky="w", padx=4)
            sw = tk.Label(cf, text="     ", relief="ridge")
            sw.grid(row=row, column=3, sticky="w", padx=4)
            sw.configure(background='#%02x%02x%02x' % (r.get(), g.get(), b.get()))

            # auto-refresh this row's swatch when any of its IntVars change
            def _refresh_swatch(*_a, sw=sw, r=r, g=g, b=b):
                try:
                    sw.configure(background='#%02x%02x%02x' % (int(r.get()), int(g.get()), int(b.get())))
                except Exception:
                    pass
            for _var in (r, g, b):
                try:
                    _var.trace_add('write', _refresh_swatch)
                except Exception:
                    try:
                        _var.trace('w', _refresh_swatch)
                    except Exception:
                        pass

        mk_color_row(1, "Color 1")
        mk_color_row(2, "Color 2")
        mk_color_row(3, "Color 3")

        # available submodels (all props)
        # Available submodels across all props (select to include in this sequence)
        self.avail_subs_lb = tk.Listbox(right_seq, selectmode="extended", height=12)
        self.avail_subs_lb.pack(fill="both", expand=True, padx=8, pady=6)
        self.avail_subs_lb.bind("<<ListboxSelect>>", self.on_avail_sub_select)
        self._avail_items = []
        self.avail_subs_lb.bind("<<ListboxSelect>>", self.on_avail_sub_select)
        self._avail_items = []
        # Hide the top Available list — using only the Included table now
        try:
            self.avail_subs_lb.pack_forget()
        except Exception:
            pass


        # Included submodels with per-segment settings
        
        # --- Read-only Edit Bar for Sequence Editor ---
        # Track which frames contain editor controls
        try:
            self._seq_editor_frames = [form, ef, pf, cf]
        except Exception:
            self._seq_editor_frames = []

        edit_bar = ttk.Frame(right_seq); edit_bar.pack(fill="x", padx=6, pady=4)
        self._btn_edit = ttk.Button(edit_bar, text="Edit", command=self._enter_seq_edit_mode); self._btn_edit.pack(side="left")
        self._btn_save = ttk.Button(edit_bar, text="Save", command=self._save_seq_edit_mode); self._btn_save.pack(side="left", padx=6)
        self._btn_cancel = ttk.Button(edit_bar, text="Cancel", command=self._cancel_seq_edit_mode); self._btn_cancel.pack(side="left")
        
        # Start read-only
        self._seq_edit_mode = False
        try:
            self._btn_save.pack_forget(); self._btn_cancel.pack_forget()
        except Exception:
            pass
        try:
            self.after(0, lambda: self.after(0, lambda: self._set_seq_editor_readonly(True)))
        except Exception:
            pass

        inc_frame = ttk.LabelFrame(seq_pan, text="Included submodels (per-segment settings)")
        inc_frame.grid(row=1, column=0, columnspan=2, sticky='nsew', padx=8, pady=8)
        self.seq_inc_tree = ttk.Treeview(inc_frame, columns=("name","count","bri","fx","sx","ix","pal","c1","c2","c3"), show="headings", height=8)
        for col, txt, w, anchor in [('name','Prop — Sub',280,'w'),('count','LEDs',60,'e'),('bri','Bri',44,'e'),('fx','FX',100,'w'),('sx','SX',44,'e'),('ix','IX',44,'e'),('pal','Pal',120,'w'),('c1','C1',96,'w'),('c2','C2',96,'w'),('c3','C3',96,'w')]:
            self.seq_inc_tree.heading(col, text=txt, anchor='center')
            self.seq_inc_tree.column(col, width=w, anchor=anchor)
        self.seq_inc_tree.pack(fill="both", expand=True, padx=6, pady=6)
        # Apply style & zebra striping for readability
        try:
            style = ttk.Style(self)
            style.configure('Seq.Treeview', rowheight=22)
            style.configure('Seq.Treeview.Heading', font=('Segoe UI', 10, 'bold'))
            self.seq_inc_tree.configure(style='Seq.Treeview')
            self.seq_inc_tree.tag_configure('even', background='#fafafa')
            self.seq_inc_tree.tag_configure('odd', background='#f0f3f7')
        except Exception:
            pass
        self.seq_inc_tree.bind("<<TreeviewSelect>>", self.on_seq_inc_select)

        inc_btns = ttk.Frame(inc_frame); inc_btns.pack_forget()
        ttk.Button(inc_btns, text="Remove selected from sequence", command=self.remove_selected_from_sequence).pack(side="right", padx=4)


        self.seq_summary = tk.StringVar(value="")

        seq_bottom = ttk.Frame(seq_tab); seq_bottom.pack(fill="x", padx=8, pady=8)
        ttk.Button(seq_bottom, text="Export presets.json", command=self.export_presets_json).pack(anchor="center", pady=8)
        bind_keyword_editor_shortcut(self)

    
    # ----- FX/Palette '(enter id)' helpers -----
    def _show_enter_id_hint(self, kind, min_id, max_id, example_text=None):
        msg = f"Type a numeric {kind} ID in this box.\nValid range: {min_id}–{max_id}. Not all numbers are used."
        if example_text:
            msg += f"\nExample: {example_text}"
        messagebox.showinfo(f"Enter {kind} ID", msg)

    def _on_fx_combo_selected(self, event=None):
        try:
            if str(self.fx_var.get()).strip() == '(enter id)':
                ids = [i for i, _ in WLED_EFFECTS]
                self._show_enter_id_hint('effect', min(ids), max(ids), '67 = Colorwaves')
                self.fx_combo.focus_set()
                try:
                    self.fx_combo.selection_range(0, tk.END)
                except Exception:
                    pass
        except Exception:
            pass

    def _on_pal_combo_selected(self, event=None):
        try:
            if str(self.pal_var.get()).strip() == '(enter id)':
                ids = [i for i, _ in WLED_PALETTES]
                self._show_enter_id_hint('palette', min(ids), max(ids), '11 = Rainbow')
                self.pal_combo.focus_set()
                try:
                    self.pal_combo.selection_range(0, tk.END)
                except Exception:
                    pass
        except Exception:
            pass

    # ---------------------- util ----------------------
    def _update_move_buttons(self, enabled: bool):
        state = "normal" if enabled else "disabled"
        self.btn_up.config(state=state)
        self.btn_down.config(state=state)

    def log(self, msg):
        # Normalize accidental literal backslash-n sequences and ensure a trailing newline
        s = str(msg).replace('\\n', '\n')
        if not s.endswith('\n'):
            s += '\n'
        self.status.insert(tk.END, s)
        self.status.see(tk.END)

    # ---------------------- files / parse ----------------------
    def _load_aliases_if_present(self, presets_path: Path):
        alias_path = presets_path.with_name("submodel_aliases.json")
        try:
            if alias_path.exists():
                with alias_path.open("r", encoding="utf-8") as f:
                    raw = json.load(f)
                self.aliases = {str(k).lower(): str(v) for k, v in raw.items() if isinstance(k, str)}
                self.log(f"Loaded submodel aliases: {len(self.aliases)}")
            else:
                self.aliases = {}
        except Exception as e:
            self.aliases = {}
            self.log(f"Alias load error: {e}")

    def browse_presets(self):
        p = filedialog.askopenfilename(title="Select presets JSON", filetypes=[("JSON","*.json"),("All","*.*")])
        if p:
            self.presets_var.set(p)
            try:
                # Auto-load immediately so user doesn't have to click Load
                self.build_clicked()
            except Exception as e:
                try:
                    self.log(f"Auto-load failed: {e}")
                except Exception:
                    pass


            except Exception:
                pass
        

    def build_clicked(self):
        p = self.presets_var.get().strip()
        if not p:
            messagebox.showerror("Error", "Pick a presets JSON."); return
        try:
            presets_path = Path(p)
            segments, _ = load_first_preset_segments(presets_path)
            self._load_aliases_if_present(presets_path)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to parse presets:\\n{e}"); return
        self.segments = segments
        self.props = build_prop_structure(self.segments, self.aliases)
        self.prop_order = list(self.props.keys())
        self.prop_tooltips.clear()
        # import any sequences/presets from the same JSON file
        try:
            self._import_sequences_from_json(Path(p))
        except Exception as ie:
            self.log(f"Sequence import note: {ie}")

        self.suppressed.clear()
        self.sub_order.clear()

        for i in self.props_tree.get_children(): self.props_tree.delete(i)
        self._tree_iids_by_prop = {}
        if not getattr(self, "_suppress_map_details", False):
            for prop in self.prop_order:
                data = self.props[prop]
                prop_leds = sum(s["stop"] - s["start"] for s in data["segments"])
                iid = self.props_tree.insert("", "end", text=prop, values=(prop_leds, "Checking..."))
                self._tree_iids_by_prop[prop] = iid
        else:
            # keep mapping UI blank until user clicks Load
            pass

        self.log(f"Loaded preset across {len(self.props)} props.")
        self.recompute_all_validations()

        self._selected_prop = None
        for i in self.sub_tree.get_children(): self.sub_tree.delete(i)
        self._update_move_buttons(False)

        # refresh sequences available list
        self.refresh_available_subs_list()
        try:
            self.seq_summary.set("")
        except Exception:
            pass
        try:
            self._clear_seq_editor_display()
        except Exception:
            pass
        try:
            self._clear_seq_inc_tree()
        except Exception:
            pass


    # ---------------------- computations ----------------------

            try:
                self.seq_summary.set("")
            except Exception:
                pass
            try:
                self._clear_seq_editor_display()
            except Exception:
                pass
            try:
                self._clear_seq_inc_tree()
            except Exception:
                pass



    def _prop_all_leds_global(self, prop):
        """All global LEDs assigned to this prop (from segments or manual_map if no segments)."""
        ids = set()
        data = self.props.get(prop, {})
        for s in data.get("segments", []):
            for g0 in range(s["start"], s["stop"]):
                ids.add(g0 + 1)
        # also include any manual_map LEDs for this prop (in case of props without auto segments)
        for (pp, sub), glist in self.manual_map.items():
            if pp == prop:
                for v in glist:
                    ids.add(v)
        return ids

    def _auto_sub_leds_global(self, prop, sub):
        out = []
        data = self.props.get(prop, {})
        for s in data.get("segments", []):
            if s["submodel"] == sub:
                for g0 in range(s["start"], s["stop"]):
                    out.append(g0 + 1)  # global LED is 1-based
        return out

    def _assign_segment_name(self, prop, gl, sub_pref=None):
        data = self.props.get(prop, {})
        for s in data.get("segments", []):
            if s["start"] <= gl-1 < s["stop"]:
                if (sub_pref is None) or (s["submodel"] == sub_pref):
                    return s["segment_name"]
        for s in data.get("segments", []):
            if s["start"] <= gl-1 < s["stop"]:
                return s["segment_name"]
        return f"{prop} manual"

    def _final_assignment_for_prop(self, prop):
        data = self.props.get(prop, {})
        prop_set = self._prop_all_leds_global(prop)

        per_sub = {}
        suppressed = self.suppressed.get(prop, set())

        # manual assignments (GLOBAL)
        for (pp, sub), glist in self.manual_map.items():
            if pp != prop:
                continue
            sset = per_sub.setdefault(sub, set())
            for gl in glist:
                if gl in prop_set:
                    sset.add(gl)

        # auto segments (GLOBAL) for subs without manual override and not suppressed
        for s in data.get("segments", []):
            sub = s["submodel"]
            if sub in suppressed:
                continue
            if (prop, sub) in self.manual_map:
                continue
            sset = per_sub.setdefault(sub, set())
            for g0 in range(s["start"], s["stop"]):
                sset.add(g0 + 1)

        # ensure empties exist for ordering
        for (pp, sub) in list(self.manual_map.keys()):
            if pp == prop:
                per_sub.setdefault(sub, set())
        for sub in self.sub_order.get(prop, []):
            per_sub.setdefault(sub, set())

        # duplicates (GLOBAL)
        led_to_subs = {}
        for sub, leds in per_sub.items():
            for gl in leds:
                led_to_subs.setdefault(gl, []).append(sub)
        dup_leds = {gl for gl, subs in led_to_subs.items() if len(subs) > 1}
        conflict_subs = {sub for sub, leds in per_sub.items() if any(gl in dup_leds for gl in leds)}

        # per-sub duplicate list (GLOBAL)
        per_sub_dup_list = {}
        for sub, sset in per_sub.items():
            ints = sorted([gl for gl in sset if gl in dup_leds])
            if ints:
                per_sub_dup_list[sub] = ints
        return per_sub, dup_leds, prop_set, conflict_subs, per_sub_dup_list

    def _build_rows_with_manual(self):
        rows = []
        for prop in self.prop_order:
            per_sub, _, prop_set, _, _ = self._final_assignment_for_prop(prop)
            order = self._build_current_sub_order(prop, list(per_sub.keys()))
            prop_led = 1
            for sub in order:
                sset = sorted(per_sub.get(sub, set()))
                sub_led = 1
                for gl in sset:
                    segname = self._assign_segment_name(prop, gl, sub_pref=sub)
                    rows.append({
                        "Prop": prop,
                        "Prop_LED": prop_led,   # local to prop (1-based)
                        "Submodel": sub,
                        "Sub_LED": sub_led,     # local to submodel (1-based)
                        "Global_LED": gl,       # include GLOBAL for reference
                        "Segment_name": segname
                    })
                    prop_led += 1
                    sub_led += 1
            covered = set().union(*per_sub.values()) if per_sub else set()
            remaining = sorted(set(prop_set) - covered)
            for gl in remaining:
                rows.append({
                    "Prop": prop,
                    "Prop_LED": prop_led,
                    "Submodel": "(unassigned)",
                    "Sub_LED": 0,
                    "Global_LED": gl,
                    "Segment_name": self._assign_segment_name(prop, gl)
                })
                prop_led += 1
        return rows

    def _build_current_sub_order(self, prop, present_subs):
        """
        Determine display order for submodels inside this *prop run*.
        - If we've stored a manual order, keep it (minus missing, plus any new at the end).
        - Otherwise, use the order submodels appear in this prop's segments (global order),
          not alphabetical.
        """
        present_subs = list(present_subs)
        saved = self.sub_order.get(prop)
        if not saved:
            # derive from this prop's run segments in appearance order
            seen = set()
            ordered = []
            for seg in self.props.get(prop, {}).get("segments", []):
                s = seg["submodel"]
                if s in present_subs and s not in seen:
                    seen.add(s); ordered.append(s)
            # append any stragglers (e.g., manual empty submodels)
            for s in present_subs:
                if s not in seen:
                    ordered.append(s)
            self.sub_order[prop] = ordered
            return ordered

        # reconcile saved order with current present set
        cur = [s for s in saved if s in present_subs]
        for s in present_subs:
            if s not in cur:
                cur.append(s)
        self.sub_order[prop] = cur
        return cur

     # ---------------------- Mapping tab view refresh ----------------------

    def refresh_current_prop_view(self):
        sel = self.props_tree.selection()
        if not sel:
            self._update_move_buttons(False)
            return
        prop = self.props_tree.item(sel[0])["text"]
        self._selected_prop = prop

        per_sub, dup_gl, prop_set, conflict_subs, per_sub_dup_gl = self._final_assignment_for_prop(prop)

        # Build local index mapping for this prop
        prop_order = sorted(prop_set)                # GLOBAL values ordered
        g2l = {gl: i+1 for i, gl in enumerate(prop_order)}  # GLOBAL->LOCAL(1-based)

        present = list(per_sub.keys())
        order = self._build_current_sub_order(prop, present)

        self.sub_tooltips.clear()
        for i in self.sub_tree.get_children(): self.sub_tree.delete(i)
        for sub in order:
            mode = "manual" if (prop, sub) in self.manual_map else ("auto (segments)" if sub not in self.suppressed.get(prop, set()) else "deleted")
            tags = ("conflict",) if sub in conflict_subs else ()
            iid = self.sub_tree.insert("", "end", values=(sub, len(per_sub[sub]), mode), tags=tags)
            if sub in per_sub_dup_gl:
                dup_vals_gl = per_sub_dup_gl[sub]
                dup_vals_local = [g2l[g] for g in dup_vals_gl if g in g2l]
                preview = ", ".join(map(str, dup_vals_local[:50]))
                if len(dup_vals_local) > 50:
                    preview += ", ..."
                self.sub_tooltips[iid] = f"Duplicate LEDs (local indices; {len(dup_vals_local)}):\\n{preview}"

        if dup_gl:
            prev_local = ", ".join(map(str, [g2l[g] for g in sorted(list(dup_gl))[:12] if g in g2l]))
            suffix = "..." if len(dup_gl) > 12 else ""
            self.log(f"[{prop}] Duplicated LED IDs across submodels (local indices): {len(dup_gl)} (e.g. {prev_local}{suffix})")

        self._set_prop_validation_display(prop, per_sub, dup_gl, prop_set)
        self._update_move_buttons(False)

    def on_prop_select(self, event=None):
        self.refresh_current_prop_view()
        # also update sequences selection list highlighting
        self.refresh_available_subs_list()

    def on_sub_select(self, event=None):
        if not self._selected_prop:
            self._update_move_buttons(False); return
        sel = self.sub_tree.selection()
        if not sel:
            self._update_move_buttons(False); return
        names = [self.sub_tree.item(i)["values"][0] for i in self.sub_tree.get_children()]
        selected_name = self.sub_tree.item(sel[0])["values"][0]
        idx = names.index(selected_name)
        can_up = idx > 0
        can_down = idx < len(names) - 1
        self.btn_up.config(state=("normal" if can_up else "disabled"))
        self.btn_down.config(state=("normal" if can_down else "disabled"))

    def _set_prop_validation_display(self, prop, per_sub, dup_gl, prop_set):
        union = set().union(*per_sub.values()) if per_sub else set()
        sum_counts = sum(len(sset) for sset in per_sub.values())
        missing = len(prop_set - union)
        extras = sum_counts - len(union)
        prop_total = len(prop_set)

        # Build local mapping for previews
        prop_order = sorted(prop_set)
        g2l = {gl: i+1 for i, gl in enumerate(prop_order)}

        if dup_gl or missing > 0 or extras > 0 or sum_counts != prop_total:
            status = "Needs attention"
            tag = "bad"
            parts = [f"Prop LEDs={prop_total}", f"Sum submodels={sum_counts}"]
            if missing > 0: parts.append(f"Missing={missing}")
            if extras > 0: parts.append(f"Overlaps={extras}")
            if dup_gl:
                parts.append(f"Duplicate LEDs={len(dup_gl)}")
                preview = ", ".join(map(str, [g2l[g] for g in sorted(list(dup_gl))[:12] if g in g2l]))
                if preview:
                    parts.append(f"Examples (local): {preview}{'...' if len(dup_gl) > 12 else ''}")
            tip = "\\n".join(parts)
        else:
            status = "OK (partitioned)"
            tag = "ok"
            tip = ""

        iid = self._tree_iids_by_prop.get(prop)
        if iid:
            vals = list(self.props_tree.item(iid)["values"])
            if len(vals) < 2:
                vals.append(status)
            else:
                vals[1] = status
            self.props_tree.item(iid, values=tuple(vals), tags=(tag,))
            self.prop_tooltips[iid] = tip

    def recompute_all_validations(self):
        for prop in self.prop_order:
            per_sub, dup_gl, prop_set, _, _ = self._final_assignment_for_prop(prop)
            self._set_prop_validation_display(prop, per_sub, dup_gl, prop_set)

    # ---------------------- Prop CRUD + reorder ----------------------
    def add_prop(self):
        dlg = tk.Toplevel(self); dlg.title("Add prop"); dlg.geometry("360x140")
        tk.Label(dlg, text="New prop name:").pack(pady=6)
        var = tk.StringVar()
        e = ttk.Entry(dlg, textvariable=var, width=32); e.pack(pady=6); e.focus_set()
        def do_add(event=None):
            name = var.get().strip()
            if not name:
                return
            if name in self.props:
                messagebox.showwarning("Exists", "A prop with that name already exists."); return
            self.props[name] = {"segments": [], "min_start": 10**9}
            self.prop_order.append(name)
            self._tree_iids_by_prop[name] = self.props_tree.insert("", "end", text=name, values=(0, "Checking..."))
            dlg.destroy()
            self.recompute_all_validations()
            self.refresh_available_subs_list()
            self._sync_map_seq_combo(); self._sync_map_seq_list(); self._sync_map_seq_list()
        ttk.Button(dlg, text="Add", command=do_add).pack(pady=6)
        dlg.bind("<Return>", do_add)
        dlg.transient(self); dlg.grab_set(); dlg.wait_window()

    def delete_prop(self):
        sel = self.props_tree.selection()
        if not sel:
            return
        prop = self.props_tree.item(sel[0])["text"]
        if messagebox.askyesno("Delete prop", f"Delete prop '{prop}'?\\n(This removes any manual submodels too.)"):
            # remove manual entries
            for (pp, sub) in list(self.manual_map.keys()):
                if pp == prop:
                    del self.manual_map[(pp, sub)]
            if prop in self.suppressed: del self.suppressed[prop]
            if prop in self.sub_order: del self.sub_order[prop]
            if prop in self.props: del self.props[prop]
            if prop in self.prop_order: del self.prop_order[prop]
            # remove from tree
            self.props_tree.delete(sel[0])
            self._selected_prop = None
            for i in self.sub_tree.get_children(): self.sub_tree.delete(i)
            self.recompute_all_validations()
            self.refresh_available_subs_list()

    def move_prop(self, delta):
        sel = self.props_tree.selection()
        if not sel:
            return
        prop = self.props_tree.item(sel[0])["text"]
        if prop not in self.prop_order:
            return
        idx = self.prop_order.index(prop)
        new_idx = idx + delta
        if new_idx < 0 or new_idx >= len(self.prop_order):
            return
        self.prop_order.pop(idx)
        self.prop_order.insert(new_idx, prop)
        # rebuild left tree to reflect new order
        for i in self.props_tree.get_children(): self.props_tree.delete(i)
        self._tree_iids_by_prop = {}
        for p in self.prop_order:
            data = self.props.get(p, {"segments":[]})
            prop_leds = sum(s["stop"] - s["start"] for s in data["segments"])
            iid = self.props_tree.insert("", "end", text=p, values=(prop_leds, ""))
            self._tree_iids_by_prop[p] = iid
        # reselect
        self.props_tree.selection_set(self._tree_iids_by_prop[prop])
        self.props_tree.see(self._tree_iids_by_prop[prop])

    # ---------------------- Submodel CRUD + reorder ----------------------

    def add_submodel(self):
        if not self._selected_prop:
            messagebox.showinfo("Add submodel", "Select a prop first.")
            return
        dlg = tk.Toplevel(self); dlg.title("New submodel"); dlg.geometry("360x180")
        tk.Label(dlg, text=f"Create a manual submodel for {self._selected_prop}:").pack(pady=6)
        name_var = tk.StringVar()
        e = ttk.Entry(dlg, textvariable=name_var, width=32); e.pack(pady=6); e.focus_set()

        info = tk.Label(dlg, text="This creates an empty manual submodel; double-click it to paste local LED indices.")
        info.pack(pady=4)

        def do_create(event=None):
            name = name_var.get().strip()
            if not name:
                return
            order = self.sub_order.setdefault(self._selected_prop, [])
            if name in order or (self._selected_prop, name) in self.manual_map:
                messagebox.showwarning("Exists", "That submodel already exists."); return
            self.manual_map[(self._selected_prop, name)] = []
            order.append(name)
            dlg.destroy()
            self.refresh_current_prop_view()
            # select the new row
            for iid in self.sub_tree.get_children():
                if self.sub_tree.item(iid)["values"][0] == name:
                    self.sub_tree.selection_set(iid)
                    self.sub_tree.see(iid)
                    break
            self.on_sub_select()
            self.refresh_available_subs_list()

        ttk.Button(dlg, text="Create", command=do_create).pack(pady=6)
        dlg.bind("<Return>", do_create)  # ENTER submits
        dlg.transient(self); dlg.grab_set(); dlg.wait_window()

    def delete_submodel(self):
        if not self._selected_prop:
            return
        sel = self.sub_tree.selection()
        if not sel:
            return
        sub = self.sub_tree.item(sel[0])["values"][0]
        prop = self._selected_prop
        if (prop, sub) in self.manual_map:
            del self.manual_map[(prop, sub)]
            if prop in self.sub_order and sub in self.sub_order[prop]:
                self.sub_order[prop].remove(sub)
            self.log(f"Deleted manual submodel: {prop} / {sub}")
        else:
            self.suppressed.setdefault(prop, set()).add(sub)
            if prop in self.sub_order and sub in self.sub_order[prop]:
                self.sub_order[prop].remove(sub)
            self.log(f"Suppressed auto submodel: {prop} / {sub}")
        self.refresh_current_prop_view()
        self._update_move_buttons(False)
        self.refresh_available_subs_list()

    def move_up(self):
        self._move_selected(delta=-1)

    def move_down(self):
        self._move_selected(delta=1)

    def _move_selected(self, delta: int):
        if not self._selected_prop:
            return
        sel = self.sub_tree.selection()
        if not sel:
            return
        prop = self._selected_prop
        names = [self.sub_tree.item(i)["values"][0] for i in self.sub_tree.get_children()]
        selected_name = self.sub_tree.item(sel[0])["values"][0]
        if selected_name not in names:
            return
        idx = names.index(selected_name)
        new_idx = idx + delta
        if new_idx < 0 or new_idx >= len(names):
            return
        order = self.sub_order.setdefault(prop, [])
        # sync saved order with shown (in case)
        if set(order) != set(names) or len(order) != len(names):
            order[:] = names
        # reorder
        order.pop(idx)
        order.insert(new_idx, selected_name)
        self.log(f"Reordered {prop}: moved '{selected_name}' -> position {new_idx+1}")
        # refresh and keep selection
        self.refresh_current_prop_view()
        for iid in self.sub_tree.get_children():
            if self.sub_tree.item(iid)["values"][0] == selected_name:
                self.sub_tree.selection_set(iid)
                self.sub_tree.see(iid)
                break
        self.on_sub_select()
        self.refresh_available_subs_list()

    # ---------------------- editor (LOCAL indexing UI) ----------------------

    def on_sub_double_click(self, event=None):
        if not self._selected_prop:
            messagebox.showinfo("Pick a prop", "Select a prop on the left first.")
            return
        sel = self.sub_tree.selection()
        if not sel:
            return
        sub = self.sub_tree.item(sel[0])["values"][0]
        prop = self._selected_prop

        
        # Use the current final assignment snapshot so defaults match the UI
        per_sub, _, prop_set, _, _ = self._final_assignment_for_prop(prop)
# Build mappings between GLOBAL and LOCAL for this prop
        prop_set = prop_set
        prop_order = sorted(prop_set)                     # GLOBAL ordered
        g2l = {gl: i+1 for i, gl in enumerate(prop_order)}
        l2g = {i+1: gl for i, gl in enumerate(prop_order)}

        # Determine current LEDs for this submodel (GLOBAL), then show LOCAL
        existing_gl = self.manual_map.get((prop, sub))
        default_gl = existing_gl if existing_gl is not None else sorted(list(per_sub.get(sub, set())))
        default_local = [g2l[g] for g in default_gl if g in g2l]

        dlg = tk.Toplevel(self)
        dlg.title(f"Edit LEDs for {prop} / {sub}")
        dlg.geometry("820x600")
        label_text = ("Paste LOCAL LED numbers for this prop (1-based), separated by commas, spaces, or newlines.\\n"
                      "You can also use ranges like 5-12. Example: 1,2,3,10-12")
        tk.Label(dlg, text=label_text).pack(pady=6)
        txt = tk.Text(dlg, height=22, wrap="word"); txt.pack(fill="both", expand=True, padx=8, pady=8)
        txt.insert("1.0", ", ".join(str(x) for x in default_local))
        txt.tag_configure("dup", background="#FFF59D")
        txt.tag_configure("bad", background="#FFCDD2")

        status_var = tk.StringVar(value=f"Count: {len(default_local)} (local to prop) — Duplicates: 0")
        tk.Label(dlg, textvariable=status_var).pack(pady=4)

        # For duplicate detection, compute other sub LEDs (GLOBAL) and convert to LOCAL
        def current_other_local_set():
            per_sub, _, _, _, _ = self._final_assignment_for_prop(prop)
            other_gl = set()
            for sname, sset in per_sub.items():
                if sname == sub:
                    continue
                other_gl |= sset
            return {g2l[g] for g in other_gl if g in g2l}

        def parse_tokens_with_spans():
            raw = txt.get("1.0", "end-1c")
            tokens = []
            for m in re.finditer(r"\d+\s*-\s*\d+|\d+", raw):
                tok = m.group(0)
                start = m.start(); end = m.end()
                start_idx = f"1.0+{start}c"
                end_idx = f"1.0+{end}c"
                if "-" in tok:
                    a_str, b_str = re.split(r"-", tok)
                    a = int(a_str.strip()); b = int(b_str.strip())
                    rng = list(range(min(a,b), max(a,b)+1))
                    tokens.append((tok, start_idx, end_idx, rng))
                else:
                    v = int(tok)
                    tokens.append((tok, start_idx, end_idx, [v]))
            return tokens

        def refresh_highlights_and_status():
            tokens = parse_tokens_with_spans()
            valid_local = []
            max_local = len(prop_order)
            for _, _, _, nums in tokens:
                for v in nums:
                    if 1 <= v <= max_local:
                        valid_local.append(v)
            seen = set()
            valid_local_unique = [v for v in valid_local if not (v in seen or seen.add(v))]
            other_local = current_other_local_set()
            dup_set = set(v for v in valid_local_unique if v in other_local)
            # Clear previous tags
            txt.tag_remove("dup", "1.0", "end")
            txt.tag_remove("bad", "1.0", "end")
            # Highlight duplicates vs other subs
            for _, sidx, eidx, nums in tokens:
                if any(v in dup_set for v in nums if 1 <= v <= max_local):
                    txt.tag_add("dup", sidx, eidx)
            # Highlight out-of-range (<=0 or > max_local)
            oor_count = 0
            for _, sidx, eidx, nums in tokens:
                if any(v < 1 or v > max_local for v in nums):
                    txt.tag_add("bad", sidx, eidx)
                    oor_count += len([v for v in nums if v < 1 or v > max_local])
            status_var.set(
                f"Count: {len(valid_local_unique)} (local to prop) — Duplicates vs other subs: {len(dup_set)} — OOR: {oor_count} (valid range 1–{max_local})"
            )

        def parse_leds_final_to_global():
            tokens = parse_tokens_with_spans()
            vals_local = []
            seen = set()
            max_local = len(prop_order)
            for _, _, _, nums in tokens:
                for v in nums:
                    if 1 <= v <= max_local and v not in seen:
                        vals_local.append(v); seen.add(v)
            # convert to GLOBAL list for storage
            return [l2g[v] for v in vals_local]

        txt.bind("<KeyRelease>", lambda e: refresh_highlights_and_status())
        refresh_highlights_and_status()

        btns = ttk.Frame(dlg); btns.pack(fill="x", pady=6)
        def save_and_close():
            # Prevent save if there are out-of-range tokens highlighted
            if txt.tag_ranges("bad"):
                messagebox.showerror("Out of range", "Some LED numbers are out of range for this prop.\nPlease fix highlights in red.")
                return
            vals_gl = parse_leds_final_to_global()
            self.manual_map[(prop, sub)] = vals_gl
            dlg.destroy()
            self.refresh_current_prop_view()
            self.recompute_all_validations()
            self.refresh_available_subs_list()

        def clear_and_close():
            if (prop, sub) in self.manual_map:
                del self.manual_map[(prop, sub)]
            dlg.destroy()
            self.refresh_current_prop_view()
            self.recompute_all_validations()
            self.refresh_available_subs_list()

        ttk.Button(btns, text="Save", command=save_and_close).pack(side="left", padx=6)
        ttk.Button(btns, text="Clear", command=clear_and_close).pack(side="left", padx=6)
        ttk.Button(btns, text="Cancel", command=dlg.destroy).pack(side="right", padx=6)
        dlg.bind("<Return>", lambda e: save_and_close())  # ENTER saves

        dlg.transient(self); dlg.grab_set(); dlg.wait_window()

    # ---------------------- LEDMAP Export ----------------------

    def export_ledmap_json_clicked(self):
        if not self.props:
            messagebox.showinfo("Info", "Load a presets JSON via Browse (it auto-loads)."); return
        # flatten rows using current prop + submodel ordering
        rows = self._build_rows_with_manual()
        # LED map maps new index -> existing physical index (0-based)
        # We'll include only rows that are assigned to some submodel, in the shown order
        mapping = [r["Global_LED"] - 1 for r in rows if r["Submodel"] != "(unassigned)"]
        if not mapping:
            messagebox.showwarning("Empty", "No assigned LEDs found to build a ledmap."); return
        out = filedialog.asksaveasfilename(
            title="Save ledmap.json",
            defaultextension=".json",
            filetypes=[("JSON","*.json")],
            initialfile="ledmap.json"
        )
        if not out:
            return
        try:
            with open(out, "w", encoding="utf-8") as f:
                json.dump({"map": mapping}, f, indent=2)
            messagebox.showinfo("Saved", f"ledmap.json saved to:\\n{out}")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save ledmap.json:\\n{e}")

    # ---------------------- Export: Excel ----------------------

    def export_excel_clicked(self):
        if "pandas" in _missing:
            messagebox.showwarning("Missing", "Install pandas + xlsxwriter to export Excel."); return
        if not self.props:
            messagebox.showinfo("Info", "Load a presets JSON via Browse (it auto-loads)."); return
        rows = self._build_rows_with_manual()
        out = filedialog.asksaveasfilename(
            title="Save Excel",
            defaultextension=".xlsx",
            filetypes=[("Excel","*.xlsx")],
            initialfile="LED_Map_Generator.xlsx"
        )
        if not out:
            return
        try:
            import pandas as pd
            df = pd.DataFrame(rows, columns=[
                "Prop","Prop_LED","Submodel","Sub_LED","Global_LED","Segment_name"
            ])
            with pd.ExcelWriter(out, engine="xlsxwriter") as writer:
                sheet = "LED Map Generator"
                df.to_excel(writer, index=False, sheet_name=sheet)
                wb = writer.book
                ws = writer.sheets[sheet]
                ws.set_column("A:A", 18)
                ws.set_column("B:B", 10)
                ws.set_column("C:C", 22)
                ws.set_column("D:D", 10)
                ws.set_column("E:E", 12)
                ws.set_column("F:F", 28)
                ws.freeze_panes(1, 0)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save Excel:\\n{e}")
            return
        messagebox.showinfo("Saved", f"Excel saved to:\\n{out}")

    def _clamp_brightness(self, v):
        try:
            v = int(v)
        except Exception:
            v = 1
        if v < 1: v = 1
        if v > 255: v = 255
        try:
            self.seq_bri_var.set(v)
        except Exception:
            try:
                self.seq_bri_var = tk.IntVar(value=v)
            except Exception:
                pass
        try:
            if hasattr(self, "bri_spin"):
                self.bri_spin.configure(from_=1, to=255, wrap=False)
        except Exception:
            pass
        return v

    def _on_brightness_spin(self):
        try:
            v = int(self.seq_bri_var.get())
        except Exception:
            v = 1
        self._clamp_brightness(v)

    # ---------------------- Sequences tab helpers ----------------------

    def _refresh_included_table_loaded(self):
        name = getattr(self, "_loaded_seq_name_editor", None)
        try:
            tree = self.seq_inc_tree
        except Exception:
            return
        if not name:
            try:
                tree.delete(*tree.get_children())
            except Exception:
                pass
            return
        cfg = self.sequences.get(name, {})
        include = cfg.get("include", set()) or []
        if isinstance(include, dict):
            include = list(include.keys())
        elif isinstance(include, set):
            include = list(include)
        rows = []
        for prop, sub in include:
            try:
                per_sub, _, _, _, _ = self._final_assignment_for_prop(prop)
                cnt = len(per_sub.get(sub, set()))
            except Exception:
                cnt = 0
            rows.append((f"{prop} — {sub}", cnt))
        rows.sort(key=lambda r: r[0].lower())
        try:
            for iid in tree.get_children():
                tree.delete(iid)
            bri = cfg.get("bri", 128)
            fx = cfg.get("fx", 0); sx = cfg.get("sx", 128); ix = cfg.get("ix", 128)
            pal = cfg.get("pal", 0)
            cols = cfg.get("col", [[255,160,0],[0,0,0],[0,0,0]])
            c1 = tuple(cols[0]) if len(cols)>0 else (0,0,0)
            c2 = tuple(cols[1]) if len(cols)>1 else (0,0,0)
            c3 = tuple(cols[2]) if len(cols)>2 else (0,0,0)
            for i, (nm, cnt) in enumerate(rows):
                tag = 'even' if (i % 2 == 0) else 'odd'
                prop, sub = self._split_prop_sub(nm)
                # pick per-sub override if available
                per = {}
                try:
                    if prop and sub:
                        per = cfg.get("per_sub", {}).get((prop, sub), {}) or {}
                except Exception:
                    per = {}
                _bri = int(per.get("bri", bri))
                _fx  = int(per.get("fx",  fx))
                _sx  = int(per.get("sx",  sx))
                _ix  = int(per.get("ix",  ix))
                _pal = int(per.get("pal", pal))
                _cols = per.get("col", cols)
                _c1 = tuple(_cols[0]) if len(_cols)>0 else c1
                _c2 = tuple(_cols[1]) if len(_cols)>1 else c2
                _c3 = tuple(_cols[2]) if len(_cols)>2 else c3

                tree.insert("", "end",
                            values=(nm, cnt, _bri,
                                    EFFECT_ID_TO_NAME.get(_fx, str(_fx)),
                                    _sx, _ix,
                                    PALETTE_ID_TO_NAME.get(_pal, str(_pal)),
                                    f"{_c1[0]},{_c1[1]},{_c1[2]}",
                                    f"{_c2[0]},{_c2[1]},{_c2[2]}",
                                    f"{_c3[0]},{_c3[1]},{_c3[2]}"),
                            tags=(tag,))
        except Exception:
            pass


    def _format_seq_summary(self, name, cfg):
        inc = cfg.get("include", set())
        bri = int(cfg.get("bri", 128))
        fx = cfg.get("fx", 0)
        sx = cfg.get("sx", 128)
        ix = cfg.get("ix", 128)
        pal = cfg.get("pal", 0)
        cols = cfg.get("col", [[255,160,0],[0,0,0],[0,0,0]])
        fx_name = EFFECT_ID_TO_NAME.get(fx, str(fx))
        pal_name = PALETTE_ID_TO_NAME.get(pal, str(pal))
        mix = cfg.get("mixed_fields", set())
        fx_part = f"fx={fx_name} ({fx})" + (" [mixed]" if "fx" in mix else "")
        sx_part = f"sx={sx}" + (" [mixed]" if "sx" in mix else "")
        ix_part = f"ix={ix}" + (" [mixed]" if "ix" in mix else "")
        pal_part = f"pal={pal_name} ({pal})" + (" [mixed]" if "pal" in mix else "")
        col_part = f"col={cols[0]}/{cols[1]}/{cols[2]}" + (" [mixed]" if "col" in mix else "")
        ov = cfg.get("per_sub", {})
        ov_cnt = len(ov)
        return (f"{name}: {len(inc)} submodels • overrides={ov_cnt} • bri={bri} • {fx_part} • "
                f"{sx_part} • {ix_part} • {pal_part} • {col_part}")

    def refresh_available_subs_list(self):
        """Populate the list of submodels across all props for sequences UI."""
        self.avail_subs_lb.delete(0, tk.END)
        items = []
        for prop in self.prop_order:
            per_sub, _, _, _, _ = self._final_assignment_for_prop(prop)
            order = self._build_current_sub_order(prop, list(per_sub.keys()))
            for sub in order:
                items.append((prop, sub, len(per_sub[sub])))
        for prop, sub, count in items:
            self.avail_subs_lb.insert(tk.END, f"{prop} — {sub}  ({count})")

        # if a sequence selected, reflect its selection
        sel_idx = self._current_seq_index()
        if sel_idx is not None:
            name = self.seq_order[sel_idx]
            cfg = self.sequences.get(name, {})
            inc = set(cfg.get("include", set()))
            if not inc and cfg.get("per_sub"):
                inc = set(cfg.get("per_sub", {}).keys())
                cfg["include"] = set(inc)
            # select indices matching inc
            idxs = []
            for i, (prop, sub, _) in enumerate(items):
                if (prop, sub) in inc:
                    idxs.append(i)
            for i in idxs:
                self.avail_subs_lb.selection_set(i)
        # (info box intentionally left blank until user clicks Load)

            try:
                self._render_loaded_badges()
            except Exception:
                pass

    def _import_sequences_from_json(self, presets_path: Path):
        """Parse presets.json and pre-populate sequences with per-segment overrides.
        Handles files where presets are stored under numeric-string keys: {"2": {...}, ...}
        and also a single-preset object with a top-level "seg".
        """
        with presets_path.open("r", encoding="utf-8") as f:
            raw = json.load(f)

        # Normalize into a dict of {name_or_id: preset_dict}
        presets = {}
        if isinstance(raw, dict) and ("seg" in raw or ("win" in raw and isinstance(raw.get("win"), dict) and "seg" in raw["win"])):
            # Single preset style (top-level seg or win.seg)
            src = raw.get("win") if isinstance(raw.get("win"), dict) and "seg" in raw.get("win", {}) else raw
            k = raw.get("n", raw.get("id", "Preset"))
            presets[str(k)] = src
        elif isinstance(raw, dict):
            # WLED style: numeric keys -> preset (seg may be under win)
            for k, v in raw.items():
                if isinstance(v, dict):
                    src = v.get("win") if isinstance(v.get("win"), dict) and "seg" in v.get("win", {}) else v
                    if isinstance(src, dict) and "seg" in src:
                        presets[str(v.get("n", k))] = src
        else:
            # Unknown format—bail gracefully
            self.log("Unrecognized presets.json structure (no 'seg' found).")
            return

        # Lookup to map (start, stop) to known name from first parsed segments
        # (used as fallback if a segment omits 'n')
        rng2name = {}
        # Also capture the first preset to inform name/range mapping in case
        # we need to build from ranges
        first_preset = next(iter(presets.values())) if presets else None
        if first_preset:
            for s in first_preset.get("seg", []):
                nm = s.get("n", "")
                rng2name[(int(s.get("start", 0)), int(s.get("stop", 0)))] = nm

        imported = []
        self.sequences.clear()
        self.seq_order.clear()

        def first_and_mixed(values, default=None):
            vals = [v for v in values if v is not None]
            if not vals:
                return default, False
            canon = [(json.dumps(v, sort_keys=True) if isinstance(v, (list, dict)) else str(v)) for v in vals]
            mixed = len(set(canon)) > 1
            return vals[0], mixed

        for pname, pval in presets.items():
            if not isinstance(pval, dict):
                continue
            if "seg" not in pval:
                continue
            name = pval.get("n", str(pname))
            segs = pval.get("seg", [])
            include = set()

            fx_vals, sx_vals, ix_vals, pal_vals, col_vals = [], [], [], [], []
            per_sub = {}

            for s in segs:
                
                seg_name = s.get("n") or rng2name.get((int(s.get("start", 0)), int(s.get("stop", 0))), "")
                base_prop, sub = split_prop_sub_hardcoded(seg_name, self.aliases)

                # Resolve to display prop (run-aware) by matching the (start, stop) to self.props
                s_start = int(s.get("start", 0)); s_stop = int(s.get("stop", 0))
                prop_display = None
                for disp, pdata in self.props.items():
                    for seginfo in pdata.get("segments", []):
                        if s_start == int(seginfo.get("start", 0)) and s_stop == int(seginfo.get("stop", 0)):
                            prop_display = disp
                            break
                    if prop_display:
                        break
                if prop_display is None:
                    prop_display = base_prop  # fallback

                include.add((prop_display, sub))


                fx_vals.append(s.get("fx"))
                sx_vals.append(s.get("sx"))
                ix_vals.append(s.get("ix"))
                pal_vals.append(s.get("pal"))
                col_vals.append(s.get("col"))

                # Store a full per-sub override including per-segment brightness and flags
                per_sub[(prop_display, sub)] = {
                    "bri": int(s.get("bri", pval.get("bri", 128))),
                    "fx": int(s.get("fx", 0)) if s.get("fx") is not None else 0,
                    "sx": int(s.get("sx", 128)) if s.get("sx") is not None else 128,
                    "ix": int(s.get("ix", 128)) if s.get("ix") is not None else 128,
                    "pal": int(s.get("pal", 0)) if s.get("pal") is not None else 0,
                    "col": s.get("col", [[255,160,0],[0,0,0],[0,0,0]]),
                    # optional segment flags we preserve on export if present
                    "rev": bool(s.get("rev", False)),
                    "grp": int(s.get("grp", 1)),
                    "spc": int(s.get("spc", 0)),
                    "of": int(s.get("of", 0)),
                }

            fx, fx_m = first_and_mixed(fx_vals, default=0)
            sx, sx_m = first_and_mixed(sx_vals, default=128)
            ix, ix_m = first_and_mixed(ix_vals, default=128)
            pal, pal_m = first_and_mixed(pal_vals, default=0)
            col, col_m = first_and_mixed(col_vals, default=[[255,160,0],[0,0,0],[0,0,0]])

            mixed_fields = set()
            if fx_m: mixed_fields.add("fx")
            if sx_m: mixed_fields.add("sx")
            if ix_m: mixed_fields.add("ix")
            if pal_m: mixed_fields.add("pal")
            if col_m: mixed_fields.add("col")

            self.sequences[name] = {
                "bri": int(pval.get("bri", 128)),
                "include": set(per_sub.keys()) if per_sub else include,
                "fx": int(fx) if fx is not None else 0,
                "sx": int(sx) if sx is not None else 128,
                "ix": int(ix) if ix is not None else 128,
                "pal": int(pal) if pal is not None else 0,
                "col": col if col is not None else [[255,160,0],[0,0,0],[0,0,0]],
                "mixed_fields": mixed_fields,
                "per_sub": per_sub
            }
            imported.append(name)

        # refresh sequences list UI
        self.seq_list.delete(0, tk.END)
        for nm in imported:
            self.seq_order.append(nm)
            self.seq_list.insert(tk.END, self._display_label_for_seq_list(nm))
        if imported:
            self.seq_list.selection_set(0)
            self.on_seq_select()
            self._sync_map_seq_combo(); self._sync_map_seq_list()
            self._apply_current_selection()
            self.log(f"Imported {len(imported)} sequences from presets.json")

    def _current_seq_index(self):
        sel = self.seq_list.curselection()
        if not sel:
            return None
        return sel[0]
    def _get_seq(self):
        idx = self._current_seq_index()
        if idx is not None:
            name = self.seq_order[idx]
            return name, self.sequences.get(name)
        # Fallback: current name if selection is empty
        name = getattr(self, '_current_seq_name', None)
        if name:
            return name, self.sequences.get(name)
        # Last resort: auto-select first if available
        if getattr(self, 'seq_order', None):
            try:
                if hasattr(self, 'seq_list') and self.seq_order:
                    self.seq_list.selection_clear(0, tk.END)
                    self.seq_list.selection_set(0)
                    self.seq_list.activate(0)
                    self._current_seq_name = self.seq_order[0]
                    try: (
                self.on_seq_select(None)
                if (which == 'seq' or not getattr(self, '_suppress_map_details', False))
                else (self._clear_seq_inc_tree() if hasattr(self, '_clear_seq_inc_tree') else None)
            )
                    except Exception: pass
                    return self.seq_order[0], self.sequences.get(self.seq_order[0])
            except Exception:
                pass
        return None, None


    def _on_seq_arrow(self, event, which="seq"):
        """Keyboard navigation for sequence lists (Up/Down)."""
        lb = getattr(self, "seq_list" if which == "seq" else "map_seq_list", None)
        try:
            import tkinter as tk
        except Exception:
            return "break"
        if not lb or not hasattr(self, "seq_order") or not self.seq_order:
            return "break"

        size = lb.size() if hasattr(lb, "size") else len(self.seq_order)
        if size <= 0:
            return "break"

        # Determine current index
        idx = None
        cur_name = getattr(self, "_current_seq_name", None)
        if cur_name in getattr(self, "seq_order", []):
            idx = self.seq_order.index(cur_name)
        if idx is None:
            try:
                sel = lb.curselection()
                idx = sel[0] if sel else 0
            except Exception:
                idx = 0

        # Compute new index
        if event.keysym == "Up":
            new = max(0, idx - 1)
        elif event.keysym == "Down":
            new = min(size - 1, idx + 1)
        else:
            return "break"

        # Apply selection & refresh
        if 0 <= new < len(self.seq_order):
            self._current_seq_name = self.seq_order[new]
            try:
                lb.selection_clear(0, tk.END)
                lb.selection_set(new)
                lb.activate(new)
                lb.see(new)
            except Exception:
                pass
            # Keep both lists in sync
            try:
                if which == "seq" and hasattr(self, "map_seq_list"):
                    self.map_seq_list.selection_clear(0, tk.END)
                    self.map_seq_list.selection_set(new)
                    self.map_seq_list.activate(new)
                    self.map_seq_list.see(new)
                elif which == "map" and hasattr(self, "seq_list"):
                    self.seq_list.selection_clear(0, tk.END)
                    self.seq_list.selection_set(new)
                    self.seq_list.activate(new)
                    self.seq_list.see(new)
            except Exception:
                pass
            try:
                self._apply_current_selection()
            except Exception:
                pass
            try:
                (
                self.on_seq_select(None)
                if (which == 'seq' or not getattr(self, '_suppress_map_details', False))
                else (self._clear_seq_inc_tree() if hasattr(self, '_clear_seq_inc_tree') else None)
            )
            except Exception:
                pass
        return "break"
    def _per_sub_settings(self, seq_cfg, prop, sub):
        s = self._lookup_per_sub(seq_cfg, prop, sub)
        if s is None:
            # fall back to sequence-level
            return {
                'fx': int(seq_cfg.get('fx',0)),
                'sx': int(seq_cfg.get('sx',128)),
                'ix': int(seq_cfg.get('ix',128)),
                'pal': int(seq_cfg.get('pal',0)),
                'col': seq_cfg.get('col', [[255,160,0],[0,0,0],[0,0,0]])
            }
        return s

    def refresh_included_tree(self):
        # Populate the included submodels grid with per-segment settings in PROPER (prop-run) ORDER
        self.seq_inc_tree.delete(*self.seq_inc_tree.get_children())
        row_idx = 0
        name, seq = self._get_seq()
        if not seq:
            return
        include = list(seq.get('include', set()))
        if not include and seq.get('per_sub'):
            include = list(seq.get('per_sub').keys())
        if not include:
            return
        # Group by prop
        by_prop = {}
        for prop, sub in include:
            by_prop.setdefault(prop, set()).add(sub)
        # Walk props in UI prop_order (already computed from global/run order)
        for prop in self.prop_order:
            subs = by_prop.get(prop)
            if not subs:
                continue
            # Order subs using the current prop run order (not alphabetical)
            ordered_subs = self._build_current_sub_order(prop, subs)
            for sub in ordered_subs:
                per = self._per_sub_settings(seq, prop, sub)
                # count LEDs for this submodel
                per_sub, _, _, _, _ = self._final_assignment_for_prop(prop)
                count = len(per_sub.get(sub, set()))
                fx_name = EFFECT_ID_TO_NAME.get(int(per.get('fx',0)), str(per.get('fx',0)))
                pal_name = PALETTE_ID_TO_NAME.get(int(per.get('pal',0)), str(per.get('pal',0)))
                c1, c2, c3 = per.get('col', [[255,160,0],[0,0,0],[0,0,0]])
                self.seq_inc_tree.insert('', 'end', values=(
                    f"{prop} — {sub}", count, int(per.get('bri', seq.get('bri',128))),
                    int(per.get('sx',128)), int(per.get('ix',128)), fx_name, pal_name,
                    f"{int(c1[0])},{int(c1[1])},{int(c1[2])}", f"{int(c2[0])},{int(c2[1])},{int(c2[2])}", f"{int(c3[0])},{int(c3[1])},{int(c3[2])}"
                ), tags=(('even' if (row_idx % 2 == 0) else 'odd'),))
                row_idx += 1

    def on_seq_inc_select(self, event=None):        # Remember last selected (prop, sub); require explicit Load to populate editor
        try:
            sel = self.seq_inc_tree.selection()
            if sel:
                label = self.seq_inc_tree.item(sel[0])['values'][0]
                if '—' in label:
                    prop, sub = [x.strip() for x in label.split('—', 1)]
                elif '-' in label:
                    prop, sub = [x.strip() for x in label.split('-', 1)]
                else:
                    prop = sub = None
                if prop and sub:
                    self._current_inc_target = (prop, sub)
                    try:
                        self.seq_summary.set(f"Selected: {prop} — {sub} (click Load)")
                    except Exception:
                        pass
        except Exception:
            pass

        # When selecting a row, load its settings into the editor controls
        sel = self.seq_inc_tree.selection()
        if not sel:
            return
        row = self.seq_inc_tree.item(sel[0])['values']
        name, seq = self._get_seq()
        if not seq:
            return
        # parse 'Prop — Sub'
        label = row[0]
        if '—' in label:
            prop, sub = [x.strip() for x in label.split('—', 1)]
        elif '-' in label:
            prop, sub = [x.strip() for x in label.split('-', 1)]
        else:
            return
        per = self._per_sub_settings(seq, prop, sub)
        self._current_inc_target = (prop, sub)
        # push into editor controls
        self.seq_bri_var.set(int(per.get('bri', seq.get('bri',128))))
        self.fx_var.set(EFFECT_ID_TO_NAME.get(int(per.get('fx',0)), str(per.get('fx',0))))
        self.sx_var.set(int(per.get('sx',128)))
        self.ix_var.set(int(per.get('ix',128)))
        self.pal_var.set(PALETTE_ID_TO_NAME.get(int(per.get('pal',0)), str(per.get('pal',0))))
        cols = per.get('col', [[255,160,0],[0,0,0],[0,0,0]])
        for (rv, gv, bv), val in zip(self.col_vars, cols):
            rv.set(int(val[0])); gv.set(int(val[1])); bv.set(int(val[2]))

    
    
    def apply_editor_to_selected(self):
        # Apply current editor settings as per-sub overrides to selected included rows
        name, seq = self._get_seq()
        if not seq:
            return
        per = seq.setdefault('per_sub', {})

        sel_iids = list(self.seq_inc_tree.selection()) if hasattr(self, 'seq_inc_tree') else []
        targets = []

        if sel_iids:
            for iid in sel_iids:
                try:
                    label = self.seq_inc_tree.item(iid)['values'][0]
                except Exception:
                    continue
                if '—' in label:
                    prop, sub = [x.strip() for x in label.split('—', 1)]
                elif '-' in label:
                    prop, sub = [x.strip() for x in label.split('-', 1)]
                else:
                    continue
                targets.append((prop, sub))
        else:
            tgt = getattr(self, '_current_inc_target', None)
            if tgt:
                targets.append(tgt)
            else:
                return

        fx_id = self._resolve_effect_id(self.fx_var.get())
        pal_id = self._resolve_palette_id(self.pal_var.get())
        cols = [[int(r.get()), int(g.get()), int(b.get())] for (r, g, b) in self.col_vars]

        for (prop, sub) in targets:
            per[(prop, sub)] = {
                'bri': int(self.seq_bri_var.get()),
                'fx': int(fx_id), 'sx': int(self.sx_var.get()), 'ix': int(self.ix_var.get()),
                'pal': int(pal_id), 'col': cols
            }

        last_label = f"{targets[-1][0]} — {targets[-1][1]}" if targets else None
        self.refresh_included_tree()
        if last_label:
            try:
                for iid in self.seq_inc_tree.get_children(''):
                    vals = self.seq_inc_tree.item(iid).get('values', [])
                    if vals and str(vals[0]) == last_label:
                        self.seq_inc_tree.selection_set(iid)
                        self.seq_inc_tree.focus(iid)
                        self.seq_inc_tree.see(iid)
                        break
            except Exception:
                pass
        try:
            self.seq_summary.set(self._format_seq_summary(name, seq))
        except Exception:
            pass

    def clear_overrides_selected(self):
        name, seq = self._get_seq()
        if not seq:
            return
        per = seq.setdefault('per_sub', {})
        sel = self.seq_inc_tree.selection()
        for iid in sel:
            label = self.seq_inc_tree.item(iid)['values'][0]
            if '—' in label:
                prop, sub = [x.strip() for x in label.split('—', 1)]
            elif '-' in label:
                prop, sub = [x.strip() for x in label.split('-', 1)]
            else:
                continue
            per.pop((prop, sub), None)
        self.refresh_included_tree()
        self.seq_summary.set(self._format_seq_summary(name, seq))

    def remove_selected_from_sequence(self):
        name, seq = self._get_seq()
        if not seq:
            return
        include = set(seq.get('include', set()))
        per = seq.setdefault('per_sub', {})
        sel = self.seq_inc_tree.selection()
        for iid in sel:
            label = self.seq_inc_tree.item(iid)['values'][0]
            if '—' in label:
                prop, sub = [x.strip() for x in label.split('—', 1)]
            elif '-' in label:
                prop, sub = [x.strip() for x in label.split('-', 1)]
            else:
                continue
            include.discard((prop, sub))
            per.pop((prop, sub), None)
        seq['include'] = include
        self.refresh_available_subs_list()
        self.refresh_included_tree()
        self.seq_summary.set(self._format_seq_summary(name, seq))

    def on_seq_select(self, event=None):

        """Selection in seq_list list: record pending only; don't clear or load."""
        try:
            sel = self.seq_list.curselection()
            if not sel:
                return
            idx = sel[0]
            if idx < 0 or idx >= len(self.seq_order):
                return
            self._pending_seq_name = self.seq_order[idx]
        except Exception:
            pass


    def add_sequence(self):
        base = "Preset"
        n = 1
        while f"{base} {n}" in self.sequences:
            n += 1
        name = f"{base} {n}"
        self.sequences[name] = {
            "bri":128, "include": set(),
            "fx": 0, "sx": 128, "ix": 128, "pal": 0,
            "col": [[255,160,0],[0,0,0],[0,0,0]],
            "per_sub": {}
        }
        self.seq_order.append(name)
        self.seq_list.insert(tk.END, self._display_label_for_seq_list(name))
        self.seq_list.selection_clear(0, tk.END)
        self.seq_list.selection_set(tk.END)
        self._current_seq_name = name
        self.on_seq_select()
        try:
            pass
        except Exception:
            pass
        try:
            self._apply_current_selection()
        except Exception:
            pass
        try:
            self._sync_map_seq_combo()
        except Exception:
            pass
        try:
            self._sync_map_seq_list()
            if hasattr(self, 'map_seq_list') and hasattr(self, 'seq_order'):
                idx = self.seq_order.index(name)
                self.map_seq_list.selection_clear(0, tk.END)
                self.map_seq_list.selection_set(idx)
                self.map_seq_list.activate(idx)
        except Exception:
            pass

    def rename_sequence(self):
        idx = self._current_seq_index()
        if idx is None:
            return
        old = self.seq_order[idx]
        dlg = tk.Toplevel(self); dlg.title("Rename sequence"); dlg.geometry("360x140")
        tk.Label(dlg, text="New name:").pack(pady=6)
        var = tk.StringVar(value=old)
        e = ttk.Entry(dlg, textvariable=var, width=32); e.pack(pady=6); e.focus_set()
        def do_rename(event=None):
            new = var.get().strip()
            if not new or new in self.sequences:
                return
            self.sequences[new] = self.sequences.pop(old)
            self.seq_order[idx] = new
            self.seq_list.delete(idx)
            self.seq_list.insert(idx, self._display_label_for_name(new))
            self.seq_list.selection_set(idx)
            dlg.destroy()
            self.on_seq_select()
        ttk.Button(dlg, text="Rename", command=do_rename).pack(pady=6)
        dlg.bind("<Return>", do_rename)
        dlg.transient(self); dlg.grab_set(); dlg.wait_window()

    def delete_sequence(self):
        idx = self._current_seq_index()
        if idx is None:
            return
        name = self.seq_order[idx]
        if messagebox.askyesno("Delete sequence", f"Delete '{name}'?"):
            del self.sequences[name]
            self.seq_order.pop(idx)
            self.seq_list.delete(idx)
            self.seq_name_var.set("")
            self.seq_bri_var.set(128)
            self.refresh_available_subs_list()

    def move_sequence(self, delta):
        idx = self._current_seq_index()
        if idx is None:
            return
        new = idx + delta
        if new < 0 or new >= len(self.seq_order):
            return
        self.seq_order[idx], self.seq_order[new] = self.seq_order[new], self.seq_order[idx]
        # refresh listbox
        self.seq_list.delete(0, tk.END)
        for n in self.seq_order:
            self.seq_list.insert(tk.END, self._display_label_for_seq_list(n))
        self.seq_list.selection_set(new)
        self.on_seq_select()

    # keep sequences model in sync with UI selections and fields
    def duplicate_sequence(self):
        """Create a copy of the currently selected sequence and insert it after the original."""
        import copy, tkinter as tk
        name, seq = self._get_seq()
        if not seq or not name:
            self.log("No sequence selected to duplicate.")
            return
        # Build a unique copy name
        base = f"{name} (copy)"
        new_name = base
        idx = 2
        while new_name in self.sequences:
            new_name = f"{base} {idx}"
            idx += 1
        # Deep copy the sequence structure
        new_seq = {}
        for k, v in seq.items():
            if k == 'include' and isinstance(v, set):
                new_seq[k] = set(v)
            elif k == 'include' and isinstance(v, list):
                new_seq[k] = list(v)
            elif k == 'per_sub' and isinstance(v, dict):
                new_seq[k] = copy.deepcopy(v)
            else:
                try:
                    new_seq[k] = copy.deepcopy(v)
                except Exception:
                    new_seq[k] = v
        # Insert right after the current one in order
        try:
            cur_idx = self.seq_order.index(name)
        except Exception:
            cur_idx = len(self.seq_order) - 1
        insert_pos = max(0, cur_idx + 1)
        self.seq_order.insert(insert_pos, new_name)
        self.sequences[new_name] = new_seq
        # Refresh both lists from seq_order
        try:
            self.seq_list.delete(0, tk.END)
            for nm in self.seq_order:
                self.seq_list.insert(tk.END, self._display_label_for_seq_list(nm))
        except Exception:
            pass
        try:
            if hasattr(self, 'map_seq_list'):
                self.map_seq_list.delete(0, tk.END)
                for nm in self.seq_order:
                    self.map_seq_list.insert(tk.END, self._display_label_for_map_list(nm))
        except Exception:
            pass
        # Make the copy current & highlight
        self._current_seq_name = new_name
        try:
            (
                self.on_seq_select(None)
                if (which == 'seq' or not getattr(self, '_suppress_map_details', False))
                else (self._clear_seq_inc_tree() if hasattr(self, '_clear_seq_inc_tree') else None)
            )
        except Exception:
            pass
        try:
            self._apply_current_selection()
        except Exception:
            pass
        self.log(f"Duplicated sequence '{name}' as '{new_name}'.")

    def _resolve_effect_id(self, fx_val: str) -> int:
        if fx_val in EFFECT_NAME_TO_ID:
            return EFFECT_NAME_TO_ID[fx_val]
        try:
            return int(fx_val)
        except Exception:
            return 0

    def _resolve_palette_id(self, pal_val: str) -> int:
        if pal_val in PALETTE_NAME_TO_ID:
            return PALETTE_NAME_TO_ID[pal_val]
        try:
            return int(pal_val)
        except Exception:
            return 0

    def _sync_sequence_from_ui(self):
        idx = self._current_seq_index()
        if idx is None:
            return
        name = self.seq_order[idx]
        new_name = self.seq_name_var.get().strip() or name
        if new_name != name and new_name not in self.sequences:
            # apply rename
            self.sequences[new_name] = self.sequences.pop(name)
            self.seq_order[idx] = new_name
            self.seq_list.delete(idx)
            self.seq_list.insert(idx, self._display_label_for_name(new_name))
            self.seq_list.selection_set(idx)
            name = new_name

        # included submodels:
        inc = set()
        items = []
        for prop in self.prop_order:
            per_sub, _, _, _, _ = self._final_assignment_for_prop(prop)
            order = self._build_current_sub_order(prop, list(per_sub.keys()))
            for sub in order:
                items.append((prop, sub))
        sel_idxs = self.avail_subs_lb.curselection()
        for i in sel_idxs:
            prop, sub = items[i]
            inc.add((prop, sub))

        # fx, sx, ix, palette, colors
        fx_id = self._resolve_effect_id(self.fx_var.get())
        pal_id = self._resolve_palette_id(self.pal_var.get())
        cols = [[int(r.get()), int(g.get()), int(b.get())] for (r, g, b) in self.col_vars]

        self.sequences[name]["include"] = inc
        # prune per_sub entries for removed items
        per = self.sequences[name].setdefault("per_sub", {})
        for k in list(per.keys()):
            if k not in inc:
                per.pop(k, None)
        self.sequences[name]["bri"] = int(self.seq_bri_var.get())
        self.sequences[name]["fx"] = int(fx_id)
        self.sequences[name]["sx"] = int(self.sx_var.get())
        self.sequences[name]["ix"] = int(self.ix_var.get())
        self.sequences[name]["pal"] = int(pal_id)
        self.sequences[name]["col"] = cols
        self.sequences[name]["mixed_fields"] = set()

        # update summary
        self.seq_summary.set(self._format_seq_summary(name, self.sequences[name]))

    # ---------------------- Export: presets.json ----------------------
    def export_presets_json(self):
        if not self.seq_order:
            messagebox.showinfo("No sequences", "Create at least one sequence first."); return
        # sync model
        self._sync_sequence_from_ui()

        # build presets
        presets = {}
        next_id = 1
        for name in self.seq_order:
            cfg = self.sequences.get(name, {"bri":128, "include": set()})
            included = cfg.get("include", set())
            bri = int(cfg.get("bri", 128))
            fx = int(cfg.get("fx", 0))
            sx = int(cfg.get("sx", 128))
            ix = int(cfg.get("ix", 128))
            pal = int(cfg.get("pal", 0))
            cols = cfg.get("col", [[255,160,0],[0,0,0],[0,0,0]])

            # gather segment ranges across included submodels
            segs = []
            for (prop, sub) in sorted(included):
                # get LED set (GLOBAL) for this sub
                per_sub_map, _, _, _, _ = self._final_assignment_for_prop(prop)
                leds = sorted(per_sub_map.get(sub, set()))
                # pick per-sub override if available, else fall back to sequence-level
                per = cfg.get("per_sub", {}).get((prop, sub), None)
                _fx = int(per.get("fx")) if per and "fx" in per else fx
                _sx = int(per.get("sx")) if per and "sx" in per else sx
                _ix = int(per.get("ix")) if per and "ix" in per else ix
                _pal = int(per.get("pal")) if per and "pal" in per else pal
                _cols = per.get("col") if per and "col" in per else cols
                # split into contiguous ranges (1-based), convert to 0-based start/stop(exclusive)
                for a, b in contiguous_ranges(leds):
                    start0 = a - 1
                    stop0 = b - 1
                    seg_obj = {
                        "start": start0,
                        "stop": stop0,
                        "n": f"{prop} — {sub}",
                        "bri": int(per.get("bri", cfg.get("bri", 128))) if per else int(cfg.get("bri", 128)), "fx": _fx, "sx": _sx, "ix": _ix, "pal": _pal,
                        "col": [_cols[0], _cols[1], _cols[2]]
                    }
                    # preserve extra flags if present
                    if per and isinstance(per, dict):
                        for k in ("rev","grp","spc","of"):
                            if k in per:
                                seg_obj[k] = per[k]
                    segs.append(seg_obj)
            if not segs:
                continue

            presets[str(next_id)] = {
                "n": name,
                "bri": bri,
                "on": True,
                "seg": segs
            }
            next_id += 1

        if not presets:
            messagebox.showwarning("Nothing to export", "No segments selected in any sequence."); return

        out = filedialog.asksaveasfilename(
            title="Save presets.json",
            defaultextension=".json",
            filetypes=[("JSON","*.json")],
            initialfile="presets.json"
        )
        if not out:
            return
        try:
            with open(out, "w", encoding="utf-8") as f:
                json.dump(presets, f, indent=2)
            messagebox.showinfo("Saved", f"presets.json saved to:\\n{out}")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save presets:\\n{e}")

    # ---------------------- tooltips ----------------------

    def _on_props_motion(self, event):
        rowid = self.props_tree.identify_row(event.y)
        if not rowid:
            self._prop_tooltip.hide()
            return
        tip = self.prop_tooltips.get(rowid, "")
        tags = self.props_tree.item(rowid, "tags")
        if "bad" in tags and tip:
            if self._prop_tooltip.visible_for_iid != rowid:
                self._prop_tooltip.show(tip, event.x_root, event.y_root)
                self._prop_tooltip.visible_for_iid = rowid
        else:
            self._prop_tooltip.hide()

    def _on_subs_motion(self, event):
        rowid = self.sub_tree.identify_row(event.y)
        if not rowid:
            self._sub_tooltip.hide()
            return
        tip = self.sub_tooltips.get(rowid, "")
        tags = self.sub_tree.item(rowid, "tags")
        if "conflict" in tags and tip:
            if self._sub_tooltip.visible_for_iid != rowid:
                self._sub_tooltip.show(tip, event.x_root, event.y_root)
                self._sub_tooltip.visible_for_iid = rowid
        else:
            self._sub_tooltip.hide()

# ---------------------- main ----------------------

    def on_avail_sub_select(self, event=None):
        """When a submodel is clicked in the Mapping tab, preview its settings in the editor.
        Robust even if no sequence is selected: falls back to current or first sequence."""
        # Identify the clicked (prop, sub)
        try:
            ixs = self.avail_subs_lb.curselection()
        except Exception:
            return
        if not ixs:
            return
        i = ixs[0]
        prop = sub = None
        items = getattr(self, "_avail_items", None)
        if items and 0 <= i < len(items):
            prop, sub, _ = items[i]
        else:
            try:
                label = self.avail_subs_lb.get(i)
                if "—" in label:
                    prop, sub = [x.strip() for x in label.split("—", 1)]
                elif "-" in label:
                    prop, sub = [x.strip() for x in label.split("-", 1)]
            except Exception:
                return
        if not prop or not sub:
            return
        # Resolve the current sequence (with fallbacks)
        name, seq = self._get_seq()
        if not seq:
            return
        # Use per-sub settings (if present) falling back to sequence-level
        per = self._per_sub_settings(seq, prop, sub)
        # Push into editor controls
        self.seq_bri_var.set(int(per.get('bri', seq.get('bri',128))))
        self.fx_var.set(EFFECT_ID_TO_NAME.get(int(per.get('fx',0)), str(per.get('fx',0))))
        self.sx_var.set(int(per.get('sx',128)))
        self.ix_var.set(int(per.get('ix',128)))
        c1, c2, c3 = per.get('col', [[255,160,0],[0,0,0],[0,0,0]])
        # Update RGB entries
        self._try_set_color_vars(c1, c2, c3)
        # Update color swatches
        try:
            self._update_color_swatch(1, c1)
            self._update_color_swatch(2, c2)
            self._update_color_swatch(3, c3)
        except Exception:
            pass
    def _norm_label(self, s: str) -> str:
        s = (s or "").lower()
        for ch in ["—", "–", "-", "_"]:
            s = s.replace(ch, " ")
        s = re.sub(r"\(\d+\)", "", s)
        s = s.replace("(all)", "all")
        s = re.sub(r"\s+", " ", s).strip()
        return s

    def _lookup_per_sub(self, seq_cfg, prop, sub):
        per = seq_cfg.get('per_sub', {})
        if not per:
            return None
        if (prop, sub) in per:
            return per[(prop, sub)]
        np, ns = self._norm_label(prop), self._norm_label(sub)
        for (pp, ss), val in per.items():
            if self._norm_label(pp) == np and self._norm_label(ss) == ns:
                return val
        for (pp, ss), val in per.items():
            if self._norm_label(pp) == np and (ns in self._norm_label(ss) or self._norm_label(ss) in ns):
                return val
        return None

    def on_avail_sub_select(self, event=None):
        """When a submodel is clicked in the Mapping tab, preview its settings in the editor.
        Robust even if no sequence is selected: falls back to current or first sequence."""
        # Identify the clicked (prop, sub)
        try:
            ixs = self.avail_subs_lb.curselection()
        except Exception:
            return
        if not ixs:
            return
        i = ixs[0]
        prop = sub = None
        items = getattr(self, "_avail_items", None)
        if items and 0 <= i < len(items):
            prop, sub, _ = items[i]
        else:
            try:
                label = self.avail_subs_lb.get(i)
                if "—" in label:
                    prop, sub = [x.strip() for x in label.split("—", 1)]
                elif "-" in label:
                    prop, sub = [x.strip() for x in label.split("-", 1)]
            except Exception:
                return
        if not prop or not sub:
            return
        # Resolve the current sequence (with fallbacks)
        name, seq = self._get_seq()
        if not seq:
            return
        # Use per-sub settings (if present) falling back to sequence-level
        per = self._per_sub_settings(seq, prop, sub)
        # Push into editor controls
        self.seq_bri_var.set(int(per.get('bri', seq.get('bri',128))))
        self.fx_var.set(EFFECT_ID_TO_NAME.get(int(per.get('fx',0)), str(per.get('fx',0))))
        self.sx_var.set(int(per.get('sx',128)))
        self.ix_var.set(int(per.get('ix',128)))
        c1, c2, c3 = per.get('col', [[255,160,0],[0,0,0],[0,0,0]])
        # Update RGB entries
        self._try_set_color_vars(c1, c2, c3)
        # Update color swatches
        try:
            self._update_color_swatch(1, c1)
            self._update_color_swatch(2, c2)
            self._update_color_swatch(3, c3)
        except Exception:
            pass
    # ===== Read-only / Edit mode for Sequence Editor =====

    def _set_seq_editor_readonly(self, readonly: bool):
        """Enable/disable editor widgets; ignore until the editor is built. Name stays readonly always."""
        try:
            if readonly:
                self._btn_edit.configure(state='normal')
                self._btn_save.configure(state='disabled')
                self._btn_cancel.configure(state='disabled')
            else:
                self._btn_edit.configure(state='disabled')
                self._btn_save.configure(state='normal')
                self._btn_cancel.configure(state='normal')
        except Exception:
            pass

        container = getattr(self, 'right_seq', None)
        if container is None or not str(container):
            return

        if not hasattr(self, '_editor_inputs'):
            self._editor_inputs = []
            def _collect(w):
                try:
                    cls = w.winfo_class()
                except Exception:
                    cls = ''
                is_input = cls in (
                    'TEntry','Entry','TCombobox','TSpinbox','Spinbox',
                    'TScale','Scale','TCheckbutton','Checkbutton','TRadiobutton','Radiobutton'
                )
                exclude = (
                    getattr(self,'_btn_edit',None),
                    getattr(self,'_btn_save',None),
                    getattr(self,'_btn_cancel',None),
                    getattr(self,'_btn_load',None),
                    getattr(self,'name_entry',None),
                )
                if is_input and (w not in exclude):
                    self._editor_inputs.append(w)
                for child in w.winfo_children():
                    _collect(child)
            try:
                _collect(container)
            except Exception:
                pass

        for w in getattr(self, '_editor_inputs', []):
            try:
                w.configure(state=('disabled' if readonly else 'normal'))
            except Exception:
                try:
                    if readonly: w.state(['disabled'])
                    else:        w.state(['!disabled'])
                except Exception:
                    pass

    def _clear_seq_editor_display(self):
        """Reset editor fields to a neutral state without enabling editing."""
        try:
            self.seq_name_var.set("")
        except Exception:
            pass
        try:
            self.seq_bri_var.set(0)
            self.sx_var.set(0)
            self.ix_var.set(0)
        except Exception:
            pass
        try:
            self.fx_var.set("")
        except Exception:
            pass
        try:
            self.pal_var.set("")
        except Exception:
            pass
        try:
            for (rv, gv, bv) in self.col_vars:
                rv.set(0); gv.set(0); bv.set(0)
        except Exception:
            pass
    def _enter_seq_edit_mode(self):
        self._seq_edit_mode = True
        # enable controls
        self._set_seq_editor_readonly(False)
        # toggle buttons
        try:
            self._btn_edit.pack_forget()
            self._btn_save.pack(side="left", padx=6)
            self._btn_cancel.pack(side="left")
        except Exception:
            pass

    def load_editor_from_selection(self):
        try:
            sel = self.seq_list.curselection()
            if not sel:
                return
            idx = sel[0]
            if idx < 0 or idx >= len(self.seq_order):
                return
            name = self.seq_order[idx]
            self._loaded_seq_name_editor = name

            
            try:
                self._render_badges_and_highlight()
            except Exception:
                pass
# Refresh list displays with [LOADED] prefix
            try:
                self._rewrite_list_with_badges(getattr(self, "seq_list", None))
                self._rewrite_list_with_badges(getattr(self, "map_seq_list", None))
            except Exception:
                pass

            try:
                self._render_loaded_badges()
            except Exception:
                pass

            try:
                lb = self.seq_list
                normal_bg = lb.cget("bg"); normal_fg = lb.cget("fg")
                hi_bg = "#ffe9a6"; hi_fg = normal_fg
                try:
                    count = lb.size()
                except Exception:
                    count = len(getattr(self, "seq_order", []) or [])
                for i in range(count):
                    try:
                        nm = lb.get(i)
                    except Exception:
                        continue
                    nm_base = self._strip_badge_prefix(nm) if hasattr(self, "_strip_badge_prefix") else nm
                    if str(nm_base) == str(name):
                        try: lb.itemconfig(i, background=hi_bg, foreground=hi_fg)
                        except Exception: pass
                    else:
                        try: lb.itemconfig(i, background=normal_bg, foreground=normal_fg)
                        except Exception: pass
            except Exception:
                pass

            cfg = self.sequences.get(name, {})
            try:
                self.name_var.set(str(name))
            except Exception:
                pass
            try:
                self.bri_var.set(int(cfg.get("bri",0)))
                self.bri_spin.set(int(cfg.get("bri",0)))
            except Exception:
                pass
            try:
                self.fx_var.set(int(cfg.get("fx",0)))
                self.sx_var.set(cfg.get("sx", 0))
                self.ix_var.set(int(cfg.get("ix",0)))
                self.pal_var.set(int(cfg.get("pal",0)))
            except Exception:
                pass
            try:
                cols = cfg.get("col", [(0,0,0),(0,0,0),(0,0,0)])
                r,g,b = cols[0] if len(cols)>0 else (0,0,0)
                self.color1_r.set(r); self.color1_g.set(g); self.color1_b.set(b)
                r,g,b = cols[1] if len(cols)>1 else (0,0,0)
                self.color2_r.set(r); self.color2_g.set(g); self.color2_b.set(b)
                r,g,b = cols[2] if len(cols)>2 else (0,0,0)
                self.color3_r.set(r); self.color3_g.set(g); self.color3_b.set(b)
            except Exception:
                pass
            try:
                self._highlight_loaded_sequence()
            except Exception:
                pass
            try:
                self._highlight_loaded_sequence()
            except Exception:
                pass
            try:
                self.seq_summary.set(self._format_seq_summary(name, cfg))
            except Exception:
                pass
            self._loaded_seq_name_editor = name

            # Refresh list displays with [LOADED] prefix
            try:
                self._rewrite_list_with_badges(getattr(self, "seq_list", None))
                self._rewrite_list_with_badges(getattr(self, "map_seq_list", None))
            except Exception:
                pass

            try:
                self._render_loaded_badges()
            except Exception:
                pass

            try:
                lb = self.seq_list
                normal_bg = lb.cget("bg"); normal_fg = lb.cget("fg")
                hi_bg = "#ffe9a6"; hi_fg = normal_fg
                try:
                    count = lb.size()
                except Exception:
                    count = len(getattr(self, "seq_order", []) or [])
                for i in range(count):
                    try:
                        nm = lb.get(i)
                    except Exception:
                        continue
                    nm_base = self._strip_badge_prefix(nm) if hasattr(self, "_strip_badge_prefix") else nm
                    if str(nm_base) == str(name):
                        try: lb.itemconfig(i, background=hi_bg, foreground=hi_fg)
                        except Exception: pass
                    else:
                        try: lb.itemconfig(i, background=normal_bg, foreground=normal_fg)
                        except Exception: pass
            except Exception:
                pass

            try:
                self.name_entry.configure(state="normal")
                self.seq_name_var.set(name or "")
                self.name_entry.configure(state="readonly")
            except Exception:
                pass
            self._clamp_brightness(cfg.get("bri", 128))
            fx = cfg.get("fx", 0)
            pal = cfg.get("pal", 0)
            self.fx_var.set(EFFECT_ID_TO_NAME.get(fx, str(fx)))
            self.pal_var.set(PALETTE_ID_TO_NAME.get(pal, str(pal)))
            self.sx_var.set(int(cfg.get("sx", 128)))
            self.ix_var.set(int(cfg.get("ix", 128)))
            cols = cfg.get("col", [[255,160,0],[0,0,0],[0,0,0]])
            for i, (r,g,b) in enumerate(cols[:3], start=1):
                try:
                    rv, gv, bv = self.col_vars[i-1]
                    rv.set(int(r)); gv.set(int(g)); bv.set(int(b))
                except Exception:
                    pass
            try:
                self._refresh_included_table_loaded()
            except Exception:
                pass
            try:
                self._set_seq_editor_readonly(True)
            except Exception:
                pass
        except Exception:
            pass
    def _save_seq_edit_mode(self):
        # Validate that Effect/Palette are not the placeholder
        try:
            if str(self.fx_var.get()).strip() == '(enter id)':
                self._on_fx_combo_selected()
                return
            if str(self.pal_var.get()).strip() == '(enter id)':
                self._on_pal_combo_selected()
                return
        except Exception:
            pass
        # Reuse existing apply logic
        try:
            self.apply_editor_to_selected()
        except Exception as e:
            try:
                self.log(f"Save failed: {e}")
            except Exception:
                pass
        self._seq_edit_mode = False
        self.after(0, lambda: self._set_seq_editor_readonly(True))
        try:
            self._btn_save.pack_forget(); self._btn_cancel.pack_forget()
            self._btn_edit.pack(side="left")
        except Exception:
            pass

    def _cancel_seq_edit_mode(self):
        # Reload current selection to revert any changes in the editor
        try:
            self.on_avail_sub_select(None)
        except Exception:
            pass
        self._seq_edit_mode = False
        self.after(0, lambda: self._set_seq_editor_readonly(True))
        try:
            self._btn_save.pack_forget(); self._btn_cancel.pack_forget()
            self._btn_edit.pack(side="left")
        except Exception:
            pass

    # ===== Mapping tab: sequence selector sync =====
    def _sync_map_seq_combo(self):
        try:
            if hasattr(self, 'map_seq_combo'):
                self.map_seq_combo['values'] = list(getattr(self, 'seq_order', []))
                cur = getattr(self, '_current_seq_name', '') or (self.seq_order[0] if getattr(self, 'seq_order', None) else '')
                if cur:
                    self.map_seq_combo.set(cur)
        except Exception:
            pass

    def _on_map_seq_change(self, event=None):
        name = self.map_seq_var.get().strip()
        if not name or name not in getattr(self, 'sequences', {}):
            return
        # Update current selection globally
        self._current_seq_name = name
        # Mirror selection in the Sequences tab listbox if present
        try:
            if hasattr(self, 'seq_order') and hasattr(self, 'seq_list'):
                idx = self.seq_order.index(name)
                self.seq_list.selection_clear(0, tk.END)
                self.seq_list.selection_set(idx)
                self.seq_list.activate(idx)
        except Exception:
            pass
        # Refresh sequence‑dependent views (editor + included grid)
        try:
            (
                self.on_seq_select(None)
                if (which == 'seq' or not getattr(self, '_suppress_map_details', False))
                else (self._clear_seq_inc_tree() if hasattr(self, '_clear_seq_inc_tree') else None)
            )
        except Exception:
            pass

    # ===== Mapping tab: sequence list mirroring the Sequences tab =====
    def _sync_map_seq_list(self):
        if not hasattr(self, 'map_seq_list'):
            return
        self.map_seq_list.delete(0, tk.END)
        for name in getattr(self, 'seq_order', []):
            self.map_seq_list.insert(tk.END, self._display_label_for_map_list(name))
        cur = getattr(self, '_current_seq_name', '')
        try:
            if cur and cur in self.seq_order:
                idx = self.seq_order.index(cur)
                self.map_seq_list.selection_clear(0, tk.END)
                self.map_seq_list.selection_set(idx)
                self.map_seq_list.activate(idx)
        except Exception:
            pass

        try:
            self._rewrite_list_with_badges(getattr(self, "map_seq_list", None))
        except Exception:
            pass

    def _ensure_seq_selection_from_map(self):
        if not hasattr(self, 'map_seq_list') or not hasattr(self, 'seq_list'):
            return
        sel = self.map_seq_list.curselection()
        if not sel:
            return
        idx = sel[0]
        self.seq_list.selection_clear(0, tk.END)
        self.seq_list.selection_set(idx)
        self.seq_list.activate(idx)

    def _on_map_seq_select(self, event=None):

        """Selection in map_seq_list list: record pending only; don't clear or load."""
        try:
            sel = self.map_seq_list.curselection()
            if not sel:
                return
            idx = sel[0]
            if idx < 0 or idx >= len(self.seq_order):
                return
            self._pending_seq_name = self.seq_order[idx]
        except Exception:
            pass



    
    def _highlight_loaded_sequence(self):
        """Ensure the loaded sequence is visibly selected in the Sequences list."""
        try:
            name = getattr(self, "_loaded_seq_name_editor", None)
            if not name or not hasattr(self, "seq_list") or not hasattr(self, "seq_order"):
                return
            if name in self.seq_order:
                idx = self.seq_order.index(name)
                try:
                    self.seq_list.selection_clear(0, tk.END)
                    self.seq_list.selection_set(idx)
                    self.seq_list.activate(idx)
                    self.seq_list.see(idx)
                except Exception:
                    pass
        except Exception:
            pass

    def _loaded_highlight_tick(self):
        """Keep the loaded highlight persistent even after list refreshes."""
        try:
            self._highlight_loaded_sequence()
        except Exception:
            pass
        try:
            self.after(800, self._loaded_highlight_tick)
        except Exception:
            pass
    def _highlight_map_loaded(self):

        lb = getattr(self, "map_seq_list", None)

        if not lb or not str(lb):

            return

        try:

            loaded = getattr(self, "_map_loaded_seq_name", None)

            normal_bg = lb.cget("bg"); normal_fg = lb.cget("fg")

            hi_bg = "#ffe9a6"; hi_fg = normal_fg

            for idx in range(lb.size()):

                nm = lb.get(idx)

                nm_base = self._strip_badge_prefix(nm) if hasattr(self, "_strip_badge_prefix") else nm
                if loaded and str(nm_base) == str(loaded):

                    lb.itemconfig(idx, background=hi_bg, foreground=hi_fg)

                else:

                    lb.itemconfig(idx, background=normal_bg, foreground=normal_fg)

        except Exception:

            pass


    def _map_loaded_highlight_tick(self):

        try:

            self._highlight_map_loaded()

        except Exception:

            pass

        try:

            self.after(300, self._map_loaded_highlight_tick)

        except Exception:

            pass


    def _clear_seq_inc_tree(self):

        try:

            if hasattr(self, "seq_inc_tree"):

                for i in self.seq_inc_tree.get_children():

                    self.seq_inc_tree.delete(i)

        except Exception:

            pass


    def load_map_from_selection(self):



        self._suppress_map_details = False
        self._ensure_props_tree_populated()
        """Explicit load for Mapping tab: set loaded name, populate editor/included, and highlight."""

        try:

            sel = self.map_seq_list.curselection()

            name = self.map_seq_list.get(sel[0]) if sel else None

        except Exception:

            name = None

        if not name:

            return

        # Remember as the mapping loaded sequence

        try:

            self._loaded_seq_name_map = name

            
            try:
                self._render_badges_and_highlight()
            except Exception:
                pass
# Refresh list displays with [LOADED] prefix
            try:
                self._rewrite_list_with_badges(getattr(self, "seq_list", None))
                self._rewrite_list_with_badges(getattr(self, "map_seq_list", None))
            except Exception:
                pass

            try:
                self._render_loaded_badges()
            except Exception:
                pass

            # sync main sequences selection to reuse logic

            self._ensure_seq_selection_from_map()

            try:
                self._highlight_map_loaded()
                self._highlight_loaded_sequence()
            except Exception:
                pass

            # call existing loader/editor logic

            self.load_editor_from_selection()

        except Exception:

            try:

                # fallback: call on_seq_select if loader not present

                (
                self.on_seq_select(None)
                if (which == 'seq' or not getattr(self, '_suppress_map_details', False))
                else (self._clear_seq_inc_tree() if hasattr(self, '_clear_seq_inc_tree') else None)
            )

            except Exception:

                pass

        # Refresh the map highlight

        try:

            self._highlight_map_loaded()

        except Exception:

            pass
    def _color_vars_ready(self):
        return all(hasattr(self, name) for name in (
            'r1_var','g1_var','b1_var','r2_var','g2_var','b2_var','r3_var','g3_var','b3_var'))

    def _try_set_color_vars(self, c1, c2, c3):
        """Safely set RGB IntVars if they exist; otherwise ignore (UI not built yet)."""
        try:
            if not self._color_vars_ready():
                return
            self._try_set_color_vars(c1, c2, c3)
            if hasattr(self, 'update_color_swatches'):
                try:
                    self.update_color_swatches()
                except Exception:
                    pass
        except Exception:
            pass

    # ---- Listbox click guard: only select when clicking on a row ----
    def _guard_listbox_click(self, event, which='seq'):
        try:
            lb = self.seq_list if which == 'seq' else self.map_seq_list
        except Exception:
            return 'break'
        try:
            if lb.size() == 0:
                return 'break'
            idx = lb.nearest(event.y)
            bbox = lb.bbox(idx)
            if not bbox:
                # Item not visible; treat as whitespace
                lb.selection_clear(0, 'end')
                return 'break'
            x, y, w, h = bbox
            # If the click is vertically outside the row's rectangle, ignore
            if event.y < y or event.y > y + h:
                lb.selection_clear(0, 'end')
                return 'break'
            # Otherwise, allow default behavior (the Listbox will select the row)
        except Exception:
            # Fail-closed: do not accidentally select
            try:
                lb.selection_clear(0, 'end')
            except Exception:
                pass
            return 'break'
        # Returning None lets Tk proceed with normal selection
        return None

    # ===== Robust selection control for sequence listboxes =====
    def _current_seq_index(self):
        try:
            if hasattr(self, 'seq_order') and getattr(self, '_current_seq_name', None) in self.seq_order:
                return self.seq_order.index(self._current_seq_name)
        except Exception:
            pass
        return None

    def _apply_current_selection(self):
        # Ensure only the active sequence is highlighted in both lists
        try:
            import tkinter as tk
        except Exception:
            return
        idx = self._current_seq_index()
        for lb_name in ('seq_list', 'map_seq_list'):
            lb = getattr(self, lb_name, None)
            if not lb:
                continue
            try:
                lb.selection_clear(0, tk.END)
                if idx is not None and 0 <= idx < lb.size():
                    lb.selection_set(idx)
                    lb.activate(idx)
                    try:
                        lb.see(idx)
                    except Exception:
                        pass
            except Exception:
                pass

    def _handle_seq_list_click(self, event, which='seq'):
        # Only change selection if the click is inside a row; otherwise revert to current
        lb = getattr(self, 'seq_list' if which == 'seq' else 'map_seq_list', None)
        if not lb:
            return 'break'
        try:
            if lb.size() == 0:
                return 'break'
            idx = lb.nearest(event.y)
            bbox = lb.bbox(idx)
            if not bbox:
                self._apply_current_selection()
                try:
                    lb.focus_set()
                except Exception:
                    pass
                return 'break'
            x, y, w, h = bbox
            if event.y < y or event.y > y + h:
                self._apply_current_selection()
                try:
                    lb.focus_set()
                except Exception:
                    pass
                return 'break'
            # Inside a row; if it's already current, just enforce highlight
            cur = self._current_seq_index()
            if cur == idx:
                self._apply_current_selection()
                try:
                    lb.focus_set()
                except Exception:
                    pass
                return 'break'
            # Switch to clicked sequence
            if hasattr(self, 'seq_order') and 0 <= idx < len(self.seq_order):
                self._current_seq_name = self.seq_order[idx]
                self._apply_current_selection()
                try:
                    (
                self.on_seq_select(None)
                if (which == 'seq' or not getattr(self, '_suppress_map_details', False))
                else (self._clear_seq_inc_tree() if hasattr(self, '_clear_seq_inc_tree') else None)
            )
                except Exception:
                    pass
                try:
                    lb.focus_set()
                except Exception:
                    pass
            return 'break'
        except Exception:
            # Fail-closed: keep current selection
            self._apply_current_selection()
            try:
                lb.focus_set()
            except Exception:
                pass
            return 'break'


def main():
    app = GeneratorGUI()
    app.mainloop()


# ===================== DEBUG SINGLE-FILE PATCH (v2.11.10) =====================
def __dbg_apply_patches_v2115():
    try:
        GUI = GeneratorGUI
    except NameError:
        return

    def _dbg(self, *a):
        msg = " ".join(map(str, a))
        try:
            print("[DBG]", msg)
            if hasattr(self, "_status_var"):
                self._status_var.set(msg[:160])
        except Exception:
            pass

    def _norm_label(self, s: str) -> str:
        s = (s or "").lower()
        for ch in ("—", "–", "-", "_"):
            s = s.replace(ch, " ")
        s = re.sub(r"\(\d+\)", "", s)
        s = s.replace("(all)", "all")
        s = re.sub(r"\s+", " ", s).strip()
        return s

    def _per_sub_settings(self, seq_cfg, prop, sub):
        per = (seq_cfg or {}).get("per_sub", {})
        _dbg(self, f"per_sub size={len(per)} lookup=({prop}, {sub})")
        if (prop, sub) in per:
            _dbg(self, "match: exact")
            return per[(prop, sub)]
        np, ns = _norm_label(self, prop), _norm_label(self, sub)
        for (pp, ss), val in per.items():
            if _norm_label(self, pp) == np and _norm_label(self, ss) == ns:
                _dbg(self, "match: normalized")
                return val
        for (pp, ss), val in per.items():
            if _norm_label(self, pp) == np and (ns in _norm_label(self, ss) or _norm_label(self, ss) in ns):
                _dbg(self, "match: partial")
                return val
        _dbg(self, "no per-sub override; using sequence defaults")
        return {}

    _orig_build = getattr(GUI, "_build_ui", None)
    def _build_ui_debug(self, *a, **kw):
        if _orig_build:
            _orig_build(self, *a, **kw)
        try:
            from tkinter import ttk
            if not hasattr(self, "_status_var"):
                self._status_var = tk.StringVar(value="")
                parent = getattr(self, "tab_sequences", None) or getattr(self, "seq_tab", None) or self
                try:
                    ttk.Label(parent, textvariable=self._status_var, foreground="#666").pack(fill="x", padx=8, pady=(0,4))
                except Exception:
                    pass
            try:
                if hasattr(self, "seq_order") and self.seq_order:
                    if hasattr(self, "seq_list"):
                        if not self.seq_list.curselection():
                            self.seq_list.selection_set(0)
                            self.seq_list.activate(0)
                            self._current_seq_name = self.seq_order[0]
                            try: (
                self.on_seq_select(None)
                if (which == 'seq' or not getattr(self, '_suppress_map_details', False))
                else (self._clear_seq_inc_tree() if hasattr(self, '_clear_seq_inc_tree') else None)
            )
                            except Exception: pass
                    else:
                        self._current_seq_name = self.seq_order[0]
            except Exception:
                pass
            try:
                if hasattr(self, "avail_subs_lb"):
                    self.avail_subs_lb.bind("<<ListboxSelect>>", self.on_avail_sub_select)
            except Exception:
                pass
        except Exception:
            pass
    GUI._build_ui = _build_ui_debug

    _orig_refresh = getattr(GUI, "refresh_available_subs_list", None)
    def refresh_available_subs_list_debug(self, *a, **kw):
        res = _orig_refresh(self, *a, **kw) if _orig_refresh else None
        try:
            items = []
            if hasattr(self, "avail_subs_lb"):
                import re as _re
                for i in range(self.avail_subs_lb.size()):
                    label = self.avail_subs_lb.get(i)
                    prop = sub = None; cnt = 0
                    if "—" in label:
                        prop, rest = [x.strip() for x in label.split("—", 1)]
                    elif "-" in label:
                        prop, rest = [x.strip() for x in label.split("-", 1)]
                    else:
                        continue
                    m = _re.search(r"(.*)\((\d+)\)\s*$", rest)
                    if m:
                        sub = m.group(1).strip(); cnt = int(m.group(2))
                    else:
                        sub = rest.strip()
                    if prop and sub:
                        items.append((prop, sub, cnt))
            self._avail_items = items
            _dbg(self, f"cached avail items: {len(items)}")
        except Exception as e:
            _dbg(self, "refresh_available_subs_list_debug:", e)
        return res
    if _orig_refresh:
        GUI.refresh_available_subs_list = refresh_available_subs_list_debug

    _orig_onseq = getattr(GUI, "on_seq_select", None)
    def on_seq_select_debug(self, *a, **kw):
        try:
            if hasattr(self, "seq_list") and hasattr(self, "seq_order"):
                idxs = self.seq_list.curselection()
                if idxs and 0 <= idxs[0] < len(self.seq_order):
                    self._current_seq_name = self.seq_order[idxs[0]]
                    _dbg(self, f"on_seq_select(): selected {self._current_seq_name}")
        except Exception as e:
            _dbg(self, "on_seq_select_debug error:", e)
        if _orig_onseq:
            return _orig_onseq(self, *a, **kw)
    if _orig_onseq:
        GUI.on_seq_select = on_seq_select_debug

    def on_avail_sub_select(self, event=None):
        """When a submodel is clicked in the Mapping tab, preview its settings in the editor.
        Robust even if no sequence is selected: falls back to current or first sequence."""
        # Identify the clicked (prop, sub)
        try:
            ixs = self.avail_subs_lb.curselection()
        except Exception:
            return
        if not ixs:
            return
        i = ixs[0]
        prop = sub = None
        items = getattr(self, "_avail_items", None)
        if items and 0 <= i < len(items):
            prop, sub, _ = items[i]
        else:
            try:
                label = self.avail_subs_lb.get(i)
                if "—" in label:
                    prop, sub = [x.strip() for x in label.split("—", 1)]
                elif "-" in label:
                    prop, sub = [x.strip() for x in label.split("-", 1)]
            except Exception:
                return
        if not prop or not sub:
            return
        # Resolve the current sequence (with fallbacks)
        name, seq = self._get_seq()
        if not seq:
            return
        # Use per-sub settings (if present) falling back to sequence-level
        per = self._per_sub_settings(seq, prop, sub)
        # Push into editor controls
        self.seq_bri_var.set(int(per.get('bri', seq.get('bri',128))))
        self.fx_var.set(EFFECT_ID_TO_NAME.get(int(per.get('fx',0)), str(per.get('fx',0))))
        self.sx_var.set(int(per.get('sx',128)))
        self.ix_var.set(int(per.get('ix',128)))
        c1, c2, c3 = per.get('col', [[255,160,0],[0,0,0],[0,0,0]])
        # Update RGB entries
        self._try_set_color_vars(c1, c2, c3)
        # Update color swatches
        try:
            self._update_color_swatch(1, c1)
            self._update_color_swatch(2, c2)
            self._update_color_swatch(3, c3)
        except Exception:
            pass
def __dbg_main_v2115():
    __dbg_apply_patches_v2115()
    try:
        main()
    except NameError:
        try:
            app = GeneratorGUI()
            app.mainloop()
        except Exception as e:
            print("DEBUG launcher error:", e)




# ---------- UI: Submodel Keywords Editor (top-level helpers) ----------
def open_subkeyword_editor(app):
    import tkinter as _tk
    from tkinter import ttk as _ttk, messagebox as _mb

    top = _tk.Toplevel(app)
    top.title("Submodel Keywords")
    top.transient(app)
    top.grab_set()

    _ttk.Label(top, text=("One keyword per line (case-insensitive).\n"
                          "When a segment name has no explicit delimiter, "
                          "the last word matching a keyword becomes the submodel.\n"
                          "Example: 'Pumpkin Outline' -> base: Pumpkin, sub: Outline")).pack(padx=10, pady=(10,6))

    txt = _tk.Text(top, width=36, height=16)
    txt.pack(padx=10, pady=6, fill="both", expand=True)
    txt.insert("1.0", "\n".join(sorted(SUB_KEYWORDS)))

    btns = _ttk.Frame(top); btns.pack(padx=10, pady=(6,10), fill="x")

    def _apply(save=False, reset=False):
        global SUB_KEYWORDS
        if reset:
            kws = [k.lower() for k in _SUB_KEYWORDS_DEFAULT]
        else:
            raw = txt.get("1.0", "end-1c")
            kws = [ln.strip().lower() for ln in raw.splitlines() if ln.strip()]
        SUB_KEYWORDS = set(kws)
        if save:
            _save_sub_keywords(kws)

        try:
            if getattr(app, "segments", None):
                app.props = build_prop_structure(app.segments, app.aliases)
                app.prop_order = list(app.props.keys())
                for i in app.props_tree.get_children():
                    app.props_tree.delete(i)
                app._tree_iids_by_prop = {}
                for prop in app.prop_order:
                    data = app.props[prop]
                    prop_leds = sum(s["stop"] - s["start"] for s in data["segments"])
                    iid = app.props_tree.insert("", "end", text=prop, values=(prop_leds, ""))
                    app._tree_iids_by_prop[prop] = iid
                app._selected_prop = None
                for i in app.sub_tree.get_children():
                    app.sub_tree.delete(i)
                app.suppressed.clear()
                app.sub_order.clear()
                app.recompute_all_validations()
                app.refresh_available_subs_list()
        except Exception as _e:
            try:
                app.log(f"Keyword change apply note: {_e}")
            except Exception:
                pass

        if save:
            _mb.showinfo("Submodel Keywords", "Saved and applied.")
        top.destroy()

    _ttk.Button(btns, text="Apply", command=lambda: _apply(save=False)).pack(side="left")
    _ttk.Button(btns, text="Save & Apply", command=lambda: _apply(save=True)).pack(side="left", padx=6)
    _ttk.Button(btns, text="Reset to Defaults", command=lambda: _apply(save=True, reset=True)).pack(side="left")
    _ttk.Button(btns, text="Cancel", command=top.destroy).pack(side="right")

def bind_keyword_editor_shortcut(app):
    # Ctrl+K to open
    try:
        app.bind_all("<Control-k>", lambda e: open_subkeyword_editor(app))
    except Exception:
        pass
# ---------------------------------------------------------------------



def _strip_star_prefix(self, s):
        try:
            s = str(s)
        except Exception:
            return s
        for prefix in ("[LOADED] ", "★ ", "* "):
            if s.startswith(prefix):
                return s[len(prefix):]
        return s

def _apply_loaded_star_to_list(self, lb):
        """Show a leading '[LOADED] ' for the loaded sequence in the given listbox, without changing selection."""
        try:
            if not lb or not str(lb):
                return
        except Exception:
            return
        loaded = self._get_loaded_for_list(lb) if hasattr(self, "_get_loaded_for_list") else getattr(self, "_loaded_seq_name_editor", None)
        try:
            count = lb.size()
        except Exception:
            count = 0
        try:
            sel = lb.curselection()
        except Exception:
            sel = ()
        for i in range(count):
            try:
                t = lb.get(i)
            except Exception:
                continue
            base = self._strip_star_prefix(t)
            display = f"[LOADED] {base}" if (loaded and str(base) == str(loaded)) else base
            if display != t:
                try:
                    lb.delete(i)
                    lb.insert(i, display)
                except Exception:
                    pass
        try:
            if sel:
                lb.selection_clear(0, "end")
                for idx in sel:
                    lb.selection_set(idx)
        except Exception:
            pass

def _render_loaded_badges(self):
    """Apply star to both sequence lists (Sequences and Mapping) without altering selections."""
    try:
        self._apply_loaded_star_to_list(getattr(self, "seq_list", None))
    except Exception:
        pass
    try:
        self._apply_loaded_star_to_list(getattr(self, "map_seq_list", None))
    except Exception:
        pass

if __name__ == "__main__":
    __dbg_apply_patches_v2115()
    try:
        main()
    except NameError:
        try:
            app = GeneratorGUI()
            app.mainloop()
        except Exception as e:
            print("DEBUG launcher error:", e)
# =================== END DEBUG SINGLE-FILE PATCH (v2.11.10) ===================