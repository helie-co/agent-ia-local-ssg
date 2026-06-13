import re


class DesignScore:
    def __init__(self):
        self.scores = {}
        self.details = []

    def add(self, category, score, max_score, detail=""):
        self.scores[category] = {"score": score, "max": max_score}
        if detail:
            self.details.append(f"  {category}: {score}/{max_score} — {detail}")

    @property
    def total(self):
        return sum(v["score"] for v in self.scores.values())

    @property
    def max_total(self):
        return sum(v["max"] for v in self.scores.values())

    @property
    def percentage(self):
        if self.max_total == 0:
            return 0
        return round(self.total / self.max_total * 100)

    def report(self):
        lines = []
        lines.append(f"Design Score: {self.percentage}/100")
        for d in self.details:
            lines.append(d)
        return "\n".join(lines)

    def passed(self, threshold=90):
        return self.percentage >= threshold


def score_deck(deck):
    score = DesignScore()
    slides = deck.get("slides", [])

    if not slides:
        score.add("slides", 0, 10, "aucune slide")
        return score

    total_slides = len(slides) + 1
    cover = deck.get("cover", {})

    cover_title = bool(cover.get("deckTitle"))
    if cover_title:
        score.add("cover", 5, 5, "titre present")
    else:
        score.add("cover", 0, 5, "titre absent")

    slides_with_titles = sum(1 for s in slides if s.get("title"))
    title_score = min(round(slides_with_titles / max(total_slides - 1, 1) * 10), 10)
    score.add("titres", title_score, 10,
              f"{slides_with_titles}/{total_slides - 1} slides avec titre")

    slides_with_conclusion = sum(1 for s in slides if s.get("conclusion"))
    conc_score = min(round(slides_with_conclusion / max(total_slides - 1, 1) * 10), 10)
    score.add("conclusions", conc_score, 10,
              f"{slides_with_conclusion}/{total_slides - 1} slides avec conclusion")

    icon_count = 0
    card_count = 0
    for s in slides:
        for key in ("cards", "kpis", "steps", "phases"):
            for item in s.get(key, []):
                card_count += 1
                if item.get("icon"):
                    icon_count += 1
    if card_count > 0:
        icon_score = min(round(icon_count / card_count * 10), 10)
        score.add("icones", icon_score, 10,
                  f"{icon_count}/{card_count} elements avec icone")
    else:
        score.add("icones", 10, 10, "aucun element (score par defaut)")

    card_word_counts = []
    for s in slides:
        for c in s.get("cards", []):
            wc = len(c.get("text", "").split()) if c.get("text") else 0
            card_word_counts.append(wc)
    if card_word_counts:
        over_12 = sum(1 for w in card_word_counts if w > 12)
        wc_score = max(0, 10 - over_12)
        avg_wc = round(sum(card_word_counts) / len(card_word_counts)) if card_word_counts else 0
        score.add("texte", wc_score, 10,
                  f"{over_12} cartes >12 mots (moy: {avg_wc})")
    else:
        score.add("texte", 10, 10, "aucune carte")

    cards_per_slide = []
    for s in slides:
        count = 0
        for key in ("cards", "kpis", "lessons", "steps", "phases"):
            count += len(s.get(key, []))
        cards_per_slide.append(count)
    if cards_per_slide:
        overloaded = sum(1 for c in cards_per_slide if c > 4)
        card_score = max(0, 10 - overloaded * 2)
        score.add("densite", card_score, 10,
                  f"{overloaded} slides >4 elements")
    else:
        score.add("densite", 10, 10, "aucune slide avec elements")

    overflow_score = 10
    for s in slides:
        for c in s.get("cards", []):
            text = c.get("text", "")
            lines = text.count("\n") + 1 if text else 0
            if lines > 3:
                overflow_score = max(0, overflow_score - 2)
    score.add("overflow", overflow_score, 10,
              f"notes: {10 - overflow_score} depassements detectes" if overflow_score < 10 else "aucun depassement")

    has_type = sum(1 for s in slides if s.get("type"))
    type_score = min(round(has_type / max(total_slides - 1, 1) * 10), 10)
    score.add("typage", type_score, 10,
              f"{has_type}/{total_slides - 1} slides typees")

    has_layout = sum(1 for s in slides if s.get("layout"))
    if has_layout == 0:
        score.add("layout", 5, 5, "layouts automatiques (non explicites)")
    else:
        score.add("layout", 5, 5, f"{has_layout} layouts explicites")

    slides_with_subtitle = sum(1 for s in slides if s.get("subtitle"))
    sub_score = min(round(slides_with_subtitle / max(total_slides - 1, 1) * 5), 5)
    score.add("sous-titres", sub_score, 5,
              f"{slides_with_subtitle}/{total_slides - 1} sous-titres")

    slides_with_notes = sum(1 for s in slides
                            if s.get("mainMessage") or s.get("message"))
    msg_score = min(round(slides_with_notes / max(total_slides - 1, 1) * 5), 5)
    score.add("message", msg_score, 5,
              f"{slides_with_notes}/{total_slides - 1} messages cles")

    kpi_slides = sum(1 for s in slides if s.get("kpis"))
    if kpi_slides > 0:
        for s in slides:
            if s.get("kpis"):
                for k in s["kpis"]:
                    if not k.get("value") or k["value"] == "—":
                        score.add("kpi", 0, 5, "valeurs KPI manquantes")
                        break
                else:
                    continue
                break
        else:
            score.add("kpi", 5, 5, "tous les KPIs ont des valeurs")
    else:
        score.add("kpi", 5, 5, "pas de slide KPI")

    return score


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
            issues.append(f"Slide {si}: type absent")
            if strict:
                s["type"] = "contexte"
                auto_fixes.append(f"Slide {si}: type -> contexte")

        cards = s.get("cards", [])
        kpis = s.get("kpis", [])
        lessons = s.get("lessons", [])
        steps = s.get("steps", [])
        phases = s.get("phases", [])

        content_count = max(len(cards), len(kpis), len(lessons),
                            len(steps), len(phases))

        if not s.get("conclusion"):
            issues.append(f"Slide {si}: conclusion absente")
            if strict:
                s["conclusion"] = _auto_conclusion(s.get("type", ""))
                auto_fixes.append(f"Slide {si}: conclusion auto")

        if content_count > 4:
            issues.append(f"Slide {si}: {content_count} elements (>4)")
            if strict:
                for key in ("cards", "kpis", "lessons", "steps", "phases"):
                    if len(s.get(key, [])) > 4:
                        s[key] = s[key][:4]
                        auto_fixes.append(f"Slide {si}: {key} tronque a 4")

        for ci, c in enumerate(cards):
            text = c.get("text", "")
            word_count = len(text.split()) if text else 0
            if word_count > 12:
                issues.append(f"Slide {si}, carte {ci+1}: {word_count} mots")
                if strict:
                    words = text.split()[:12]
                    s["cards"][ci]["text"] = " ".join(words)
                    auto_fixes.append(f"Slide {si}, carte {ci+1}: reduit a 12 mots")
            if not c.get("icon"):
                issues.append(f"Slide {si}, carte {ci+1}: icone absente")
                if strict:
                    s["cards"][ci]["icon"] = "lightbulb"
                    auto_fixes.append(f"Slide {si}, carte {ci+1}: icone -> lightbulb")

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
