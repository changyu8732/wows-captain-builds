#!/usr/bin/env python3
"""
update_builds.py
----------------
Run this script whenever "World of Warships captain builds.docx" is updated.
It re-extracts all builds from the docx, preserves any existing skills data
already in captain_builds.html, regenerates the BUILDS JS array, updates the
HTML file, and optionally pushes to GitHub.

Usage:
    python3 update_builds.py            # update HTML only
    python3 update_builds.py --push     # update HTML + git commit + push
"""

import argparse
import json
import re
import subprocess
import sys
import zipfile
from pathlib import Path
from xml.etree import ElementTree

# ── Paths ─────────────────────────────────────────────────────────────────────
SCRIPT_DIR  = Path(__file__).parent
DOCX        = SCRIPT_DIR / "World of Warships captain builds.docx"
HTML        = SCRIPT_DIR / "captain_builds.html"

# ── H1 whitelist (doc has a misclassified H1 "Kléber, Mogador") ───────────────
REAL_H1 = {
    "Tech tree branches, analogous premiums and T10 ships",
    "Other premiums",
}

# ── Nation detection: (patterns, nation_code) ─────────────────────────────────
NATION_PATTERNS = [
    (["ijn ", "japan", "yamato", "shimakaze", "harugumo", "hayate", "zao", "yodo", "bungo",
      "azuma", "yoshino", "kitakami", "hosho", "shinano", "hakuryu", "mikasa", "yuubari",
      "maya", "tokachi", "chikuma", "daisen", "aki", "suzuya", "i-56", "yumihari", "ise"], "IJN"),
    (["usn ", "usa ", "united states", "gearing", "burrows", "forrest", "montana", "vermont",
      "des moines", "worcester", "austin", "salem", "essex", "midway", "fdr", "franklin",
      "atlanta", "flint", "san diego", "kidd", "black", "halford", "johnston", "velos",
      "benham", "sims", "ohio", "rhode island", "maine", "wisconsin", "cachalot", "balao",
      "gato", "archerfish", "independence", "yorktown", "langley", "ranger", "lexington",
      "saipan", "hornet", "vallejo", "hawaii", "congress", "alaska", "puerto rico",
      "illinois", "narai"], "USN"),
    (["vmf ", "soviet", "russian", "kremlin", "delny", "khabarovsk", "grozovoi", "slava",
      "petropavlovsk", "stalingrad", "nevsky", "smolensk", "komissar", "nakhimov",
      "ognevoi", "bagration", "molotov", "kirov", "mikoyan", "pozharsky", "navarin",
      "kozma", "neustrashimy", "k-1", "s-1", "l-20", "s-189", "provorny", "arkhangelsk",
      "sibir", "sevastopol", "moskva", "tashkent"], "VMF"),
    (["km ", "german", "hindenburg", "elbing", "z-52", "z-42", "z-31", "schlieffen",
      "preussen", "mecklenburg", "richthofen", "immelmann", "adalbert", "hildebrand",
      "georg", "lütjens", "u-69", "u-190", "u-2501", "u-4501", "karl von", "zf-6", "z-44",
      "hoche", "odin", "brandenburg", "anhalt", "weimar", "nurnberg", "mainz", "wiesbaden",
      "admiral schröder", "blucher", "bremen", "e. löwenhardt", "lowenhardt"], "KM"),
    (["rn ", "british", "uk", "daring", "druid", "conqueror", "st. vincent", "minotaur",
      "edgar", "plymouth", "goliath", "gibraltar", "indomitable", "audacious", "malta",
      "eagle", "hermes", "furious", "alliance", "thrasher", "undine", "sturdy",
      "repulse", "thunderer", "hampshire", "dido", "colossus", "gambia", "monmouth",
      "cambridge", "hull", "eskimo", "seal", "victoria", "laffey", "irresistible",
      "lugdunum", "bridgeport", "pioneer", "jupiter '42", "gallant", "theseus"], "RN"),
    (["fr ", "french", "kléber", "kleber", "henri", "marseille", "colbert", "cassard",
      "bourgogne", "republique", "carnot", "bayard", "surcouf", "marceau", "la pampa",
      "le havre", "béarn", "champagne", "picardie", "fr25", "roussillion"], "FR"),
    (["italian", "venezia", "napoli", "varese", "colombo", "ferrucio", "gonzaga",
      "regolo", "barbiano", "sicilia", "verdi", "messina", "aquila", "leone", "roma",
      "it ", "bixio"], "IT"),
    (["pan-eu", "eu ", "halland", "gdansk", "småland", "ragnar", "dalarna", "svea",
      "blyska", "blysk", "stord", "orkan", "friesland", "groningen", "jäger", "karl xiv",
      "niord", "thor", "pan eu", "błysk"], "EU"),
    (["pan-asia", "pa ", "yue yang", "lüshun", "pan asia", "fenyang", "loyang",
      "siliwangi", "nanning", "dalian", "tianjin", "anshan", "tashkent '39",
      "yimeng", "lanzhou", "wukong", "teng she", "taihang", "xin zhong", "huanghe"], "PA"),
    (["pan-america", "pam ", "libertad", "san martin", "atlântico", "nueve",
      "almirante grau", "rio de janeiro", "valparaiso", "comandante aguirre"], "PAm"),
    (["spain", "es ", "castilla", "álvaro", "bazán", "canarias", "elli"], "ES"),
    (["netherlands", "nl ", "goulden", "leeuw", "prins van", "de zeven", "utrecht",
      "tromp", "tonijn", "willem"], "NL"),
    (["commonwealth", "haida", "huron", "vampire", "cerberus", "brisbane", "hector",
      "incheon"], "Commonwealth"),
]

# Manual overrides for ships that pattern-matching can't reliably catch
NATION_OVERRIDES = {
    138: "RN",   # Jupiter '42 — uses curly apostrophe, breaks pattern match
}

# Battlecruiser line keywords (H3 text containing these → Battlecruiser class)
BC_KEYWORDS = ["bc", "battlecruiser", "bungo", "schlieffen", "st. vincent"]

# Captain recommendation patterns
CAPTAIN_PATTERNS = [
    r"(Suzuki|Ovechkin|Lütjens|Lutjens|Yamamoto|Kuznetsov|Halsey|"
    r"Honoré|Honore|Auboyneau|Swirski|Rong|Sa Zhenbing|Doe|Dunkirk|"
    r"Sansonetti|Cunningham|Cunnigham|Znamensky|Von Jütland|Jütland)\s+(?:is|are|captain|bros|brothers)",
    r"(?:recommended|Using|use)\s+(?:is\s+)?(Suzuki|Ovechkin|Lütjens|Yamamoto|"
    r"Kuznetsov|Halsey|Honoré|Honore|Auboyneau|Swirski|Sansonetti|Cunningham|Cunnigham)",
]


# ── Docx parsing ──────────────────────────────────────────────────────────────

def parse_docx(docx_path):
    """Return list of {text, has_image, style} dicts from the docx."""
    with zipfile.ZipFile(docx_path) as z:
        with z.open("word/document.xml") as f:
            tree = ElementTree.parse(f)
    ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
    lines = []
    for para in tree.getroot().findall(".//w:p", ns):
        text = "".join(
            node.text for node in para.findall(".//w:t", ns) if node.text
        ).strip()
        has_image = bool(para.findall(".//w:drawing", ns))
        pPr = para.find("w:pPr", ns)
        style = ""
        if pPr is not None:
            ps = pPr.find("w:pStyle", ns)
            if ps is not None:
                style = ps.get(
                    "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val", ""
                )
        lines.append({"text": text, "has_image": has_image, "style": style})
    return lines


def extract_builds(lines):
    """Walk the parsed lines and return a list of build dicts."""
    img_positions = [i for i, ln in enumerate(lines) if ln["has_image"]]

    current = {"h1": "", "h2": "", "h3": "", "h4": ""}
    builds = []

    for i, line in enumerate(lines):
        s = line["style"]
        t = line["text"]

        if s == "Heading1" and t in REAL_H1:
            current = {"h1": t, "h2": "", "h3": "", "h4": ""}
        elif s == "Heading2" and t:
            current.update({"h2": t, "h3": "", "h4": ""})
        elif s == "Heading3":
            if t:
                current.update({"h3": t, "h4": ""})
        elif s == "Heading4":
            current["h4"] = t

        if not line["has_image"]:
            continue

        seq_idx = img_positions.index(i)
        seq_num = seq_idx + 1

        # Collect notes until next image
        next_img = (
            img_positions[seq_idx + 1]
            if seq_idx + 1 < len(img_positions)
            else len(lines)
        )
        notes = [
            lines[j]["text"]
            for j in range(i + 1, next_img)
            if not lines[j]["has_image"]
            and not lines[j]["style"].startswith("Heading")
            and lines[j]["text"]
        ]

        # Build name
        build_name = current["h3"]
        if current["h4"]:
            build_name += f" \u2014 {current['h4']}"

        # Nation
        nation = detect_nation(seq_num, current["h3"], current["h2"])

        # Ship class
        ship_class = normalise_class(current["h2"])
        if ship_class == "Battleship":
            ship_class = maybe_battlecruiser(current["h3"], notes)

        # appliesTo
        applies_to = extract_applies_to(notes, current["h3"])

        # Captain recommendation
        captain = extract_captain(notes)

        # Description = first note (overview line)
        description = notes[0] if notes else ""

        builds.append(
            {
                "seq": seq_num,
                "nation": nation,
                "shipClass": ship_class,
                "buildName": build_name,
                "appliesTo": applies_to,
                "description": description,
                "notes": notes,
                "captainRecommendation": captain,
            }
        )

    return builds


# ── Helper functions ──────────────────────────────────────────────────────────

def detect_nation(seq_num, h3, h2):
    if seq_num in NATION_OVERRIDES:
        return NATION_OVERRIDES[seq_num]
    combined = (h3 + " " + h2).lower()
    for patterns, code in NATION_PATTERNS:
        if any(p in combined for p in patterns):
            return code
    return ""


def normalise_class(h2):
    mapping = {
        "Aircraft carriers": "Aircraft Carrier",
        "Submarines":        "Submarine",
        "Battleships":       "Battleship",
        "Cruisers":          "Cruiser",
        "Destroyers":        "Destroyer",
    }
    return mapping.get(h2, h2)


def maybe_battlecruiser(h3, notes):
    h3l = h3.lower()
    if any(kw in h3l for kw in BC_KEYWORDS):
        return "Battlecruiser"
    for note in notes:
        # Match 'battlecruiser' only when NOT immediately followed by 'accuracy'
        # (avoids false positives like 'battlecruiser accuracy' describing a BB)
        if re.search(r"battlecruiser\b(?!\s+accuracy)", note, re.IGNORECASE):
            return "Battlecruiser"
    return "Battleship"


def extract_applies_to(notes, h3):
    for note in notes:
        if note.lower().startswith("applies to"):
            ships_text = re.sub(r"^applies to:?\s*", "", note, flags=re.IGNORECASE).rstrip(".")
            return [s.strip() for s in re.split(r",\s*(?:and\s+)?|\s+and\s+", ships_text) if s.strip()]
    # Fallback: derive from H3
    clean = re.sub(r"\s*\(.*?\)\s*", "", h3).strip()
    if "/" in clean:
        return [s.strip() for s in clean.split("/")]
    if "," in clean:
        return [s.strip() for s in clean.split(",")]
    return [clean] if clean else []


def extract_captain(notes):
    for note in notes:
        for pat in CAPTAIN_PATTERNS:
            m = re.search(pat, note, re.IGNORECASE)
            if m:
                return m.group(1)
    return ""


# ── Skills extraction from existing HTML ─────────────────────────────────────

def extract_existing_skills(html_text):
    """Return {seq_num: raw_skills_js_text} from the current HTML."""
    m = re.search(r"const BUILDS\s*=\s*\[(.*?)\];\s*\n", html_text, re.DOTALL)
    if not m:
        return {}
    builds_text = m.group(1)

    # Split into individual top-level { } blocks
    blocks, depth, current = [], 0, ""
    for ch in builds_text:
        if ch == "{":
            depth += 1
            current += ch
        elif ch == "}":
            depth -= 1
            current += ch
            if depth == 0:
                blocks.append(current.strip())
                current = ""
        elif depth > 0:
            current += ch

    skills = {}
    for block in blocks:
        img_m = re.search(r'images:\s*\["build_images/seq_(\d+)\.png"\]', block)
        if not img_m:
            continue
        seq = int(img_m.group(1))
        skills_start = block.find("skills:")
        if skills_start == -1:
            continue
        bracket_start = block.index("[", skills_start)
        d, end = 0, bracket_start
        for k in range(bracket_start, len(block)):
            if block[k] == "[":
                d += 1
            elif block[k] == "]":
                d -= 1
                if d == 0:
                    end = k
                    break
        # Return only the content between [ and ] so empty arrays give ""
        # Only keep if it contains actual skill objects (has 'name:' key)
        content = block[bracket_start + 1 : end].strip()
        if content and "name:" in content:
            skills[seq] = content
    return skills


# ── JS generation ─────────────────────────────────────────────────────────────

def esc(s):
    """Escape a string for a JS double-quoted string."""
    return s.replace("\\", "\\\\").replace('"', '\\"').replace("\n", "\\n").replace("\r", "")


def format_array_of_strings(items):
    if not items:
        return "[]"
    escaped = [f'"{esc(s)}"' for s in items]
    if sum(len(e) for e in escaped) + len(escaped) * 2 < 80:
        return "[" + ", ".join(escaped) + "]"
    return "[\n" + ",\n".join(f"      {e}" for e in escaped) + "\n    ]"


def build_to_js(b, skills_map):
    seq = b["seq"]
    skills_js = skills_map.get(seq, "")
    skills_block = f"[\n{skills_js}\n    ]" if skills_js else "[]"
    notes_block = (
        "[\n" + ",\n".join(f'      "{esc(n)}"' for n in b["notes"]) + "\n    ]"
        if b["notes"]
        else "[]"
    )
    applies_block = format_array_of_strings(b["appliesTo"])

    return (
        f"  {{\n"
        f'    nation: "{esc(b["nation"])}",\n'
        f'    shipClass: "{esc(b["shipClass"])}",\n'
        f'    buildName: "{esc(b["buildName"])}",\n'
        f"    appliesTo: {applies_block},\n"
        f'    images: ["build_images/seq_{seq}.png"],\n'
        f'    description: "{esc(b["description"])}",\n'
        f"    skills: {skills_block},\n"
        f"    notes: {notes_block},\n"
        f'    captainRecommendation: "{esc(b["captainRecommendation"])}"\n'
        f"  }}"
    )


def generate_builds_js(builds, skills_map):
    entries = [build_to_js(b, skills_map) for b in builds]
    return "const BUILDS = [\n" + ",\n".join(entries) + "\n];"


# ── HTML update ───────────────────────────────────────────────────────────────

def replace_builds_in_html(html_text, new_builds_js):
    start_token = "const BUILDS = ["
    start = html_text.index(start_token)
    depth, end = 0, start
    for i in range(start, len(html_text)):
        if html_text[i] == "[":
            depth += 1
        elif html_text[i] == "]":
            depth -= 1
            if depth == 0:
                end = html_text.index(";", i + 1) + 1
                break
    return html_text[:start] + new_builds_js + html_text[end:]


# ── Git push ──────────────────────────────────────────────────────────────────

def git_push(message):
    cmds = [
        ["git", "add", "captain_builds.html"],
        ["git", "commit", "-m", message],
        ["git", "push"],
    ]
    for cmd in cmds:
        result = subprocess.run(cmd, cwd=SCRIPT_DIR, capture_output=True, text=True)
        print(result.stdout.strip())
        if result.returncode != 0:
            print(f"  stderr: {result.stderr.strip()}", file=sys.stderr)
            if cmd[1] == "commit" and "nothing to commit" in result.stdout:
                print("  (nothing changed, skipping push)")
                return
            if result.returncode != 0:
                print(f"  WARNING: command failed: {' '.join(cmd)}", file=sys.stderr)


# ── Main ──────────────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(description="Rebuild WoWS captain builds HTML from docx.")
    parser.add_argument("--push", action="store_true", help="git commit + push after updating HTML")
    parser.add_argument("--no-skills-preserve", action="store_true",
                        help="Don't carry over existing skills (use empty arrays for all builds)")
    args = parser.parse_args()

    print(f"Reading {DOCX.name} ...")
    if not DOCX.exists():
        sys.exit(f"ERROR: {DOCX} not found.")
    lines = parse_docx(DOCX)
    print(f"  Parsed {len(lines)} paragraphs")

    builds = extract_builds(lines)
    print(f"  Extracted {len(builds)} builds")

    # Report any unknown nations
    unknown = [b for b in builds if not b["nation"]]
    if unknown:
        print(f"\n  WARNING: {len(unknown)} builds have unknown nation:")
        for b in unknown:
            print(f"    seq_{b['seq']}: {b['buildName']}")
        print("  Add entries to NATION_OVERRIDES or NATION_PATTERNS in this script.\n")

    # Report class breakdown
    by_class = {}
    for b in builds:
        by_class[b["shipClass"]] = by_class.get(b["shipClass"], 0) + 1
    print("  Classes: " + ", ".join(f"{k}={v}" for k, v in sorted(by_class.items())))

    print(f"\nReading existing HTML to preserve skills ...")
    html_text = HTML.read_text(encoding="utf-8")
    skills_map = {} if args.no_skills_preserve else extract_existing_skills(html_text)
    print(f"  Found skills for {len(skills_map)} existing builds")

    print("Generating new BUILDS JS array ...")
    new_builds_js = generate_builds_js(builds, skills_map)

    # Quick JS sanity check via node if available
    try:
        import tempfile, os
        with tempfile.NamedTemporaryFile(mode="w", suffix=".js", delete=False) as tmp:
            tmp.write(new_builds_js + "\nprocess.stdout.write('OK:' + BUILDS.length);")
            tmp_path = tmp.name
        node_check = subprocess.run(
            ["node", tmp_path],
            capture_output=True, text=True, timeout=15
        )
        os.unlink(tmp_path)
        if node_check.returncode == 0 and node_check.stdout.startswith("OK:"):
            count = node_check.stdout.split(":")[1]
            print(f"  JS validation passed ({count} builds)")
        else:
            print(f"  WARNING: JS validation failed: {node_check.stderr.strip()[:200]}")
    except (FileNotFoundError, subprocess.TimeoutExpired):
        print("  (node.js not found, skipping JS validation)")

    print(f"Updating {HTML.name} ...")
    new_html = replace_builds_in_html(html_text, new_builds_js)
    HTML.write_text(new_html, encoding="utf-8")
    print(f"  Done — {len(new_html):,} chars written")

    if args.push:
        print("\nPushing to GitHub ...")
        git_push(f"Update builds from docx ({len(builds)} builds)")
        print("  Done — changes live on GitHub Pages within ~1 minute")
    else:
        print("\nDone. Run with --push to also commit and push to GitHub.")


if __name__ == "__main__":
    main()
