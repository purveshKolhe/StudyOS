import os
import json
import re
from datetime import datetime
from typing import Dict, Any, List

from flask import Flask, render_template, request, jsonify, send_from_directory, abort
from pptx import Presentation

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
TEMPLATE_DIR = os.path.join(BASE_DIR, "template")
TEMPLATE_PPTX = os.path.join(TEMPLATE_DIR, "template.pptx")
METADATA_JSON = os.path.join(TEMPLATE_DIR, "layout_metadata.json")
OUTPUT_DIR = os.path.join(BASE_DIR, "generated")

os.makedirs(OUTPUT_DIR, exist_ok=True)

app = Flask(__name__)


def load_metadata() -> Dict[str, Any]:
    with open(METADATA_JSON, "r", encoding="utf-8") as f:
        meta = json.load(f)
    # Precompute helpers
    layouts_by_name = {l["layout_name"].strip().lower(): l for l in meta.get("layouts", [])}
    must_have_names = [
        l["layout_name"] for l in meta.get("layouts", [])
        if "must have" in l.get("layout_description", "").lower()
    ]
    ignore_names = [
        l["layout_name"] for l in meta.get("layouts", [])
        if "ignore" in l.get("layout_description", "").lower()
    ]
    meta["_layouts_by_name"] = layouts_by_name
    meta["_must_have_names"] = must_have_names
    meta["_ignore_names"] = ignore_names
    return meta


METADATA = load_metadata()


def sentence_case(s: str) -> str:
    s = s.strip()
    if not s:
        return s
    return s[0].upper() + s[1:]


def apply_content_rule(text: str, rule_desc: str) -> str:
    d = rule_desc.lower()
    t = text or ""
    if "all-caps" in d or "all caps" in d:
        t = t.upper()
    elif "title case" in d:
        t = t.title()
    elif "sentence case" in d:
        # basic sentence case
        t = sentence_case(t)
    return t


def clip_text(text: str, max_chars: int | None) -> str:
    if max_chars and max_chars > 0 and len(text) > max_chars:
        return text[:max_chars]
    return text


def find_layout(prs: Presentation, layout_name: str):
    target = (layout_name or "").strip().lower()
    # exact match
    for l in prs.slide_layouts:
        if (l.name or "").strip().lower() == target:
            return l
    # relaxed contains
    for l in prs.slide_layouts:
        if target and target in (l.name or "").strip().lower():
            return l
    # fallback: first layout
    return prs.slide_layouts[0]


def fill_placeholders(slide, layout_name: str, items: Dict[str, str]):
    layout = METADATA["_layouts_by_name"].get(layout_name.strip().lower())
    rule_by_pid: Dict[int, Dict[str, Any]] = {}
    if layout:
        for ph in layout.get("placeholders", []):
            rule_by_pid[ph["id"]] = ph
    # Fill only matching placeholders by placeholder idx
    # Use slide.placeholders to iterate known placeholder shapes
    for shp in slide.placeholders:
        try:
            pid = shp.placeholder_format.idx
        except Exception:
            continue
        key = str(pid)
        if key in items and hasattr(shp, "text_frame"):
            text = items[key]
            rule_desc = rule_by_pid.get(pid, {}).get("content_description", "")
            maxchars = rule_by_pid.get(pid, {}).get("maxchars")
            text = apply_content_rule(text, rule_desc)
            text = clip_text(text, maxchars)
            tf = shp.text_frame
            tf.clear()
            tf.text = text


def add_thank_you_slide(prs: Presentation):
    # Add last layout unchanged as the final slide
    last_layout = prs.slide_layouts[len(prs.slide_layouts) - 1]
    prs.slides.add_slide(last_layout)


def enforce_plan_rules(plan: Dict[str, Any]) -> Dict[str, Any]:
    slides: List[Dict[str, Any]] = plan.get("slides", [])

    # Remove any slide using ignored layouts
    ignore_set = set([n.strip().lower() for n in METADATA["_ignore_names"]])
    keep: List[Dict[str, Any]] = []
    for s in slides:
        name = (s.get("layout_name") or "").strip().lower()
        if name not in ignore_set:
            keep.append(s)
    slides = keep

    # Ensure MUST HAVE layouts exist at least once
    present = set([(s.get("layout_name") or "").strip().lower() for s in slides])
    for must in METADATA["_must_have_names"]:
        key = must.strip().lower()
        if key not in present:
            # Minimal default content for must-have
            placeholders = {}
            layout = METADATA["_layouts_by_name"].get(key)
            if layout:
                for ph in layout.get("placeholders", []):
                    pid = str(ph["id"])
                    desc = ph.get("content_description", "")
                    if "title" in desc.lower() and "summary" not in desc.lower():
                        placeholders[pid] = "SUMMARY" if "synonym" in desc.lower() else "TOPIC"
                    else:
                        placeholders[pid] = "Auto-generated content about the topic."
            slides.append({
                "layout_name": must,
                "placeholders": placeholders
            })

    plan["slides"] = slides
    return plan


def safe_slug(s: str) -> str:
    s = s.strip().lower()
    s = re.sub(r"[^a-z0-9\-\_\s]", "", s)
    s = re.sub(r"\s+", "-", s)
    return s or "presentation"


def call_gemini_for_plan(topic: str, metadata: Dict[str, Any]) -> Dict[str, Any]:
    api_key = os.environ.get("GEMINI_API_KEY")
    if not api_key:
        return {}

    try:
        import google.generativeai as genai
        genai.configure(api_key=api_key)
        model = genai.GenerativeModel(model_name="gemini-2.5-flash-lite")

        system = (
            "You generate slide plans as JSON for a PowerPoint builder. "
            "Use the provided layout metadata to choose appropriate layouts and map placeholder IDs to text. "
            "Constraints: ALWAYS include all layouts whose description contains 'MUST HAVE'. "
            "NEVER use any layout whose description contains 'IGNORE'. "
            "Return JSON only with the exact schema. Do not include markdown."
        )
        schema_hint = {
            "slides": [
                {
                    "layout_name": "Blank",
                    "placeholders": {"10": "TITLE IN ALL CAPS", "11": "Title Case subtitle"}
                }
            ]
        }
        prompt = (
            f"TOPIC: {topic}\n\n" \
            f"METADATA (JSON):\n{json.dumps(metadata, ensure_ascii=False)}\n\n" \
            f"Output strictly as minified JSON matching this schema: {json.dumps(schema_hint)}\n"
        )
        res = model.generate_content([system, prompt])
        text = getattr(res, "text", None)
        if not text and hasattr(res, "candidates") and res.candidates:
            parts = []
            for c in res.candidates:
                if hasattr(c, "content") and hasattr(c.content, "parts"):
                    for p in c.content.parts:
                        if hasattr(p, "text"):
                            parts.append(p.text)
            text = "\n".join(parts)
        if not text:
            return {}
        # Extract JSON
        m = re.search(r"\{[\s\S]*\}$", text.strip())
        raw = m.group(0) if m else text
        plan = json.loads(raw)
        return plan if isinstance(plan, dict) else {}
    except Exception:
        return {}


def stub_plan(topic: str) -> Dict[str, Any]:
    # Minimal plan covering MUST HAVEs
    def pad(min_len: int, base: str) -> str:
        s = base
        while len(s) < min_len:
            s += " " + base
        return s[:max(min_len, len(base))]

    agenda_desc = (
        "This section introduces a key idea, provides context, and outlines what the student will learn in this part. "
        "It connects the topic to real-world relevance and sets clear expectations."
    )
    agenda_desc = pad(160, agenda_desc)

    summary_desc = (
        "This summary consolidates the main concepts, definitions, and relationships covered in the lesson. It clarifies the core idea, "
        "highlights the essential steps or properties, and reflects on misconceptions. The section also suggests how to practice and "
        "apply the knowledge with confidence in new situations."
    )
    # Ensure ~500 chars
    while len(summary_desc) < 520:
        summary_desc += " " + summary_desc.split(".")[0] + "."

    return {
        "slides": [
            {
                "layout_name": "Blank",
                "placeholders": {
                    "10": topic.upper(),
                    "11": "An Overview"
                }
            },
            {
                "layout_name": "2_Custom Layout",
                "placeholders": {
                    "10": "INTRODUCTION",
                    "12": "CORE IDEAS",
                    "11": "APPLICATIONS",
                    "13": "SUMMARY",
                    "14": agenda_desc,
                    "15": agenda_desc,
                    "16": agenda_desc,
                    "17": agenda_desc
                }
            },
            {
                "layout_name": "14_Custom Layout",
                "placeholders": {
                    "10": "SUMMARY",
                    "11": summary_desc
                }
            }
        ]
    }


def build_pptx_from_plan(topic: str, plan: Dict[str, Any]) -> str:
    prs = Presentation(TEMPLATE_PPTX)

    # Build slides from plan
    for slide_spec in plan.get("slides", []):
        layout_name = slide_spec.get("layout_name")
        placeholders = slide_spec.get("placeholders", {})
        if not layout_name:
            continue
        layout = find_layout(prs, layout_name)
        slide = prs.slides.add_slide(layout)
        try:
            fill_placeholders(slide, layout_name, placeholders)
        except Exception:
            # Continue even if a placeholder can't be filled
            pass

    # Append the Thank You slide as last layout unchanged
    try:
        add_thank_you_slide(prs)
    except Exception:
        pass

    slug = safe_slug(topic)
    fname = f"{slug}-{datetime.now().strftime('%Y%m%d-%H%M%S')}.pptx"
    out_path = os.path.join(OUTPUT_DIR, fname)
    prs.save(out_path)
    return fname


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/generate", methods=["POST"])
def generate():
    try:
        data = request.get_json(force=True)
        topic = (data.get("topic") or "").strip()
    except Exception:
        return jsonify({"error": "Invalid request"}), 400

    if not topic:
        return jsonify({"error": "Topic is required"}), 400

    plan = call_gemini_for_plan(topic, METADATA)
    if not plan:
        plan = stub_plan(topic)

    plan = enforce_plan_rules(plan)

    try:
        filename = build_pptx_from_plan(topic, plan)
    except FileNotFoundError:
        return jsonify({"error": "Template PPTX not found."}), 500
    except Exception as e:
        return jsonify({"error": f"Failed to build PPTX: {e}"}), 500

    return jsonify({
        "filename": filename,
        "download_url": f"/download/{filename}"
    })


@app.route("/download/<path:filename>")
def download(filename):
    if not os.path.exists(os.path.join(OUTPUT_DIR, filename)):
        abort(404)
    return send_from_directory(OUTPUT_DIR, filename, as_attachment=True, mimetype=(
        "application/vnd.openxmlformats-officedocument.presentationml.presentation"
    ))


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=True)
