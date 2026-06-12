#!/usr/bin/env python3
import sys
import os
import json
import tempfile
import random
import re
from pathlib import Path

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import theme
import icons
import slidebuilder
import quality
import parser as md_parser


SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
PROJECT_DIR = os.path.abspath(os.path.join(SCRIPT_DIR, "..", "..", "..", ".."))
PPT_V2_DIR = os.path.join(PROJECT_DIR, "ppt-v2")
ICONS_DIR = os.path.join(PPT_V2_DIR, "icons")


def ai_generate_deck(text):
    deck = {
        "cover": {
            "deckTitle": extract_title(text),
            "deckSubtitle": "Perspectives et enjeux",
            "date": "Juin 2026",
            "authors": "OpenCode"
        },
        "slides": []
    }

    n_requested = 1
    m = re.search(r'(\d+)\s*slides?\s*(sur|sur le|sur la|sur les|a propos)', text.lower())
    if m:
        n_requested = min(int(m.group(1)), 8)

    lines = text.strip().split("\n")
    slides_text = []
    current = []

    for line in lines:
        if re.match(r'^\d+[\.\)]\s|^Slide\s+\d+|^##\s', line.strip()):
            if current:
                slides_text.append(" ".join(current))
            current = [line]
        else:
            current.append(line)
    if current:
        slides_text.append(" ".join(current))

    if not slides_text:
        sentences = re.split(r'[.!?\n]+', text)
        sentences = [s.strip() for s in sentences if len(s.strip()) > 20]
        if len(sentences) >= n_requested:
            slides_text = sentences[:n_requested]
        else:
            slides_text = [text]

    if not slides_text:
        slides_text = [text]

    if len(slides_text) < n_requested:
        topic = _extract_topic(text)
        slides_text = [topic] * n_requested

    n_slides = min(len(slides_text), n_requested)

    for i in range(n_slides):
        slide = _build_slide_from_text(slides_text[i], i)
        deck["slides"].append(slide)

    if not deck["slides"]:
        deck["slides"].append(_build_slide_from_text(text, 0))

    return deck


def _extract_topic(text):
    m = re.search(r'(?:sur|sur le|sur la|sur les|a propos de?)\s+(.+?)(?:\.|$|avec)', text)
    if m:
        return f"Presentation sur {m.group(1).strip()}"
    m2 = re.search(r'(\w+(?:\s+\w+){0,5})$', text.replace("-", ""))
    if m2:
        return f"Presentation sur {m2.group(1).strip()}"
    return "Presentation"


def extract_title(text):
    lines = text.strip().split("\n")
    for line in lines:
        line = line.strip()
        if line.startswith("# "):
            return line[2:].strip()
        if len(line) > 10 and len(line) < 100:
            return line.strip()
    return "Presentation"


def _build_slide_from_text(text, idx):
    topics = {
        "ia": "contexte", "intelligence artificielle": "contexte",
        "generative": "contexte", "genai": "contexte",
        "probleme": "problematique", "defi": "problematique",
        "enjeu": "problematique", "risque": "problematique",
        "process": "processus", "demarche": "processus",
        "etape": "processus", "methode": "processus",
        "lecon": "enseignements", "enseign": "enseignements",
        "retour": "enseignements", "bilan": "enseignements",
        "kpi": "chiffres", "chiffre": "chiffres",
        "indicateur": "chiffres", "mesure": "chiffres",
        "role": "roles", "responsabilite": "roles",
        "equipe": "roles", "organisation": "roles",
        "adoption": "adoption", "transformation": "adoption",
        "conclusion": "bilan", "synthese": "bilan",
        "recommendation": "bilan",
    }

    text_lower = text.lower()
    slide_type = "contexte"
    for keyword, stype in topics.items():
        if keyword in text_lower:
            slide_type = stype

    title = extract_slide_title(text, idx)

    slide = {
        "title": title,
        "type": slide_type,
        "cards": [],
        "kpis": [],
        "lessons": [],
        "steps": [],
    }

    sentences = re.split(r'[.!?\n]+', text)
    sentences = [s.strip() for s in sentences if len(s.strip()) > 10]

    used_icons = ["data", "robot-ai", "lightbulb", "chart",
                  "innovation", "strategy", "metrics", "solution"]

    for j, sent in enumerate(sentences[:4]):
        if len(sent) > 100:
            sent = sent[:100]
        words = sent.split()
        if len(words) > 12:
            sent = " ".join(words[:12])

        icon = used_icons[(idx * 4 + j) % len(used_icons)]

        slide["cards"].append({
            "title": f"Point {j + 1}",
            "text": sent,
            "icon": icon,
        })

    if slide_type == "chiffres":
        slide["kpis"] = [
            {"value": "—", "label": "Indicateur 1", "icon": "metrics"},
            {"value": "—", "label": "Indicateur 2", "icon": "chart"},
        ]
    elif slide_type == "enseignements":
        slide["lessons"] = [c["text"] for c in slide["cards"]]
    elif slide_type == "processus":
        slide["steps"] = [{"title": c["title"], "text": c["text"]}
                          for c in slide["cards"]]

    conclusion_templates = [
        "La cle du succes repose sur l'engagement de toutes les parties prenantes",
        "La transformation est un voyage qui se construit etape par etape",
        "L'innovation prend tout son sens quand elle est au service du metier",
        "La gouvernance est le pilier d'une transformation durable",
    ]
    slide["conclusion"] = random.choice(conclusion_templates)

    return slide


def extract_slide_title(text, idx):
    lines = text.strip().split("\n")
    for line in lines:
        line = line.strip()
        if line.startswith("# "):
            continue
        if line.startswith("## "):
            return line[2:].strip()
        if len(line) > 5 and len(line) < 80:
            return line.strip()
    return f"Slide {idx + 1}"


def main():
    args = sys.argv[1:]
    raw_text = " ".join(args)

    if not raw_text or raw_text == "--help":
        readme = os.path.join(SCRIPT_DIR, "README.md")
        if os.path.exists(readme):
            with open(readme, "r", encoding="utf-8") as f:
                print(f.read())
        return

    if raw_text == "--create-template":
        icons.write_all_icons(ICONS_DIR)
        print(f"Icones SVG regenerees dans: {ICONS_DIR}")
        template_path = os.path.join(PPT_V2_DIR, "template.pptx")
        print(f"Template cree: {template_path}")
        return

    icons.write_all_icons(ICONS_DIR)
    strict = "--strict" in raw_text
    if strict:
        raw_text = raw_text.replace("--strict", "").strip()

    parsed = raw_text.startswith("@")
    if parsed:
        filepath = raw_text[1:].strip()
        if not os.path.exists(filepath):
            print(f"Fichier introuvable: {filepath}")
            return
        deck = md_parser.parse_markdown_file(filepath)
    else:
        deck = ai_generate_deck(raw_text)

    deck = md_parser.reduce_text(deck)

    issues, fixes = quality.check_deck(deck, strict=strict)
    if issues:
        print("=== Controle qualite ===")
        for iss in issues:
            print(f"  {iss}")
        if fixes:
            for f in fixes:
                print(f"  [Corrige] {f}")
        if issues and not strict:
            print("\nUtilisez --strict pour appliquer les corrections automatiques.")
            return
        elif issues and strict:
            print("\nCorrections automatiques appliquees.")

    output_dir = PROJECT_DIR
    output_name = f"Presentation_{random.randint(1000, 9999)}.pptx"
    output_path = os.path.join(output_dir, output_name)
    result = slidebuilder.build_deck(deck, output_path, icons_dir=ICONS_DIR)

    qa_issues = quality.validate_output(result["path"])
    if qa_issues:
        for qi in qa_issues:
            print(f"  QA: {qi}")

    print(f"\nPresentation generee: {result['path']}")
    print(f"Slides: {result['slides']}")


if __name__ == "__main__":
    main()
