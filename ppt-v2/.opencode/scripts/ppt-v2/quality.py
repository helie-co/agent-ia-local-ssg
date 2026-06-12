import re


def check_deck(deck, strict=False):
    issues = []
    auto_fixes = []

    cover = deck.get("cover", {})
    if not cover.get("deckTitle"):
        issues.append("ERREUR: cover.deckTitle absent")
        if strict:
            cover["deckTitle"] = "Presentation"
            auto_fixes.append("cover.deckTitle = 'Presentation'")

    slides = deck.get("slides", [])
    if not slides:
        issues.append("ERREUR: aucune slide definie")
        return issues, auto_fixes

    for i, s in enumerate(slides):
        si = i + 1

        if not s.get("title"):
            issues.append(f"Slide {si}: titre absent")
            if strict:
                s["title"] = f"Slide {si}"
                auto_fixes.append(f"Slide {si}: titre auto genere")

        if not s.get("type"):
            issues.append(f"Slide {si}: type absent (contexte, processus, etc.)")
            if strict:
                s["type"] = "contexte"
                auto_fixes.append(f"Slide {si}: type -> contexte")

        cards = s.get("cards", [])
        kpis = s.get("kpis", [])
        lessons = s.get("lessons", [])
        steps = s.get("steps", [])
        phases = s.get("phases", [])

        content_count = max(len(cards), len(kpis), len(lessons), len(steps), len(phases))

        if s.get("type") in ("contexte", "chiffres") and not kpis and not cards:
            issues.append(f"Slide {si} ({s['type']}) : aucun KPI defini")
            if strict:
                s["kpis"] = [{"value": "—", "label": "Indicateur"}]
                auto_fixes.append(f"Slide {si}: KPI ajoute par defaut")

        if s.get("type") == "enseignements" and not lessons and not cards:
            issues.append(f"Slide {si} (enseignements) : aucune lecon definie")
            if strict:
                s["lessons"] = ["Enseignement cle a documenter"]
                auto_fixes.append(f"Slide {si}: lecon ajoutee par defaut")

        if not s.get("conclusion"):
            issues.append(f"Slide {si}: conclusion absente")
            if strict:
                concl = _auto_conclusion(s.get("type", ""))
                s["conclusion"] = concl
                auto_fixes.append(f"Slide {si}: conclusion auto -> '{concl}'")

        if content_count > 4:
            issues.append(f"Slide {si}: {content_count} elements (>4 recommande)")
            if strict:
                if len(cards) > 4:
                    s["cards"] = cards[:4]
                    auto_fixes.append(f"Slide {si}: cards tronquees a 4")
                if len(kpis) > 4:
                    s["kpis"] = kpis[:4]
                    auto_fixes.append(f"Slide {si}: kpis tronques a 4")
                if len(lessons) > 5:
                    s["lessons"] = lessons[:5]
                    auto_fixes.append(f"Slide {si}: lecons tronquees a 5")

        for ci, c in enumerate(cards):
            title = c.get("title", "")
            text = c.get("text", "")
            word_count = len(text.split()) if text else 0
            if word_count > 12:
                issues.append(f"Slide {si}, carte {ci + 1}: {word_count} mots (>12)")
                if strict:
                    words = text.split()[:12]
                    s["cards"][ci]["text"] = " ".join(words)
                    auto_fixes.append(f"Slide {si}, carte {ci + 1}: texte reduit a 12 mots")

            if not c.get("icon"):
                issues.append(f"Slide {si}, carte {ci + 1}: icone absente")
                if strict:
                    s["cards"][ci]["icon"] = "lightbulb"
                    auto_fixes.append(f"Slide {si}, carte {ci + 1}: icone -> lightbulb")

    return issues, auto_fixes


def _auto_conclusion(slide_type):
    conclusions = {
        "contexte": "La comprehension du contexte est le point de depart de toute transformation.",
        "problematique": "La resolution du probleme passe par une approche structuree.",
        "probleme": "La resolution du probleme passe par une approche structuree.",
        "question": "La reponse apportee eclaire la prise de decision.",
        "reponse": "La mise en oeuvre est la cle du succes.",
        "processus": "La standardisation du processus garantit la reproductibilite.",
        "roles": "La mobilisation des equipes est la cle du succes.",
        "enseignements": "Le partage des enseignements accelere la maturation.",
        "chiffres": "La mesure de la performance pilote l'amelioration continue.",
        "bilan": "L'amelioration continue est le moteur de la transformation.",
        "adoption": "L'adoption reelle passe par l'appropriation par les utilisateurs.",
        "default": "La valeur se cree dans la duree par l'engagement collectif.",
    }
    return conclusions.get(slide_type.lower(), conclusions["default"])


def validate_output(path):
    import os
    if not os.path.exists(path):
        return [f"Fichier non trouve: {path}"]
    size_kb = os.path.getsize(path) / 1024
    if size_kb < 10:
        return [f"Fichier anormalement petit: {size_kb:.1f} KB"]
    return []
