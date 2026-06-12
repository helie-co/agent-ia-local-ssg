import re
import json


def parse_text(text):
    lines = text.strip().split("\n")
    deck = {
        "cover": {"deckTitle": "", "deckSubtitle": "", "date": "", "authors": ""},
        "slides": [],
    }

    current_slide = None
    current_type = None
    for line in lines:
        line = line.strip()
        if not line:
            continue

        if line.startswith("# "):
            deck["cover"]["deckTitle"] = line[2:].strip()
        elif line.startswith("## "):
            if current_slide:
                deck["slides"].append(current_slide)
            title = line[3:].strip()
            current_slide = {
                "title": title,
                "type": "contexte",
                "layout": "",
                "cards": [],
                "kpis": [],
                "lessons": [],
                "steps": [],
                "phases": [],
            }
        elif line.lower().startswith("type:"):
            if current_slide:
                current_slide["type"] = line.split(":", 1)[1].strip()
        elif line.lower().startswith("layout:"):
            if current_slide:
                current_slide["layout"] = line.split(":", 1)[1].strip()
        elif line.lower().startswith("template:"):
            pass
        elif line.lower().startswith("source:"):
            pass
        elif "carte" in line.lower() and ":" in line:
            if current_slide:
                parts = line.split(":", 1)[1].split("|", 1)
                title = parts[0].strip()
                text = parts[1].strip() if len(parts) > 1 else ""
                current_slide.setdefault("cards", []).append({
                    "title": title, "text": text, "icon": "lightbulb"
                })
        elif "kpi" in line.lower() and ":" in line:
            if current_slide:
                parts = line.split(":", 1)[1].split("|", 1)
                value = parts[0].strip()
                label = parts[1].strip() if len(parts) > 1 else ""
                current_slide.setdefault("kpis", []).append({
                    "value": value, "label": label
                })
        elif "lecon" in line.lower() and ":" in line:
            if current_slide:
                text = line.split(":", 1)[1].strip()
                current_slide.setdefault("lessons", []).append(text)
        elif "etape" in line.lower() and ":" in line:
            if current_slide:
                parts = line.split(":", 1)[1].split("|", 1)
                title = parts[0].strip()
                desc = parts[1].strip() if len(parts) > 1 else ""
                current_slide.setdefault("steps", []).append({
                    "title": title, "text": desc
                })
        elif "question" in line.lower() and ":" in line:
            if current_slide:
                current_slide["question"] = line.split(":", 1)[1].strip()
        elif "reponse" in line.lower() and ":" in line:
            if current_slide:
                current_slide["answer"] = line.split(":", 1)[1].strip()
        elif current_slide and line and not line.startswith("#"):
            current_slide.setdefault("notes", []).append(line)

    if current_slide:
        deck["slides"].append(current_slide)

    _auto_type(deck)
    return deck


def _auto_type(deck):
    for s in deck.get("slides", []):
        if s.get("type") != "contexte" or not s.get("layout"):
            continue
        title = s.get("title", "").lower()
        if s.get("kpis"):
            s["type"] = "chiffres"
        elif s.get("lessons"):
            s["type"] = "enseignements"
        elif s.get("steps"):
            s["type"] = "processus"
        elif s.get("question"):
            s["type"] = "question"
        elif s.get("answer"):
            s["type"] = "reponse"
        elif any(w in title for w in ["probleme", "problematique", "defi", "enjeu"]):
            s["type"] = "problematique"
        elif any(w in title for w in ["lecon", "enseign", "retour"]):
            s["type"] = "enseignements"
        elif any(w in title for w in ["process", "etape", "approche"]):
            s["type"] = "processus"
        elif any(w in title for w in ["role", "responsabilite"]):
            s["type"] = "roles"
        elif any(w in title for w in ["kpi", "chiffre", "indicateur", "mesure"]):
            s["type"] = "chiffres"
        elif any(w in title for w in ["conclusion", "bilan", "synthese"]):
            s["type"] = "bilan"


def parse_markdown_file(path):
    with open(path, "r", encoding="utf-8") as f:
        content = f.read()
    return parse_text(content)


def reduce_text(deck, max_words=12, max_lines=4):
    for s in deck.get("slides", []):
        for c in s.get("cards", []):
            text = c.get("text", "")
            words = text.split()
            if len(words) > max_words:
                c["text"] = " ".join(words[:max_words])
    return deck


def classify_request(text):
    text_lower = text.lower()

    if not text or len(text) < 5:
        return {"mode": "help"}

    if text.startswith("@"):
        return {"mode": "file", "path": text[1:].strip()}

    if text == "--help":
        return {"mode": "help"}

    if text == "--create-template":
        return {"mode": "create-template"}

    return {"mode": "generate", "text": text}
