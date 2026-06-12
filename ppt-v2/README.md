# /ppt-v2

Commandes OpenCode pour generer des presentations PowerPoint professionnelles de niveau cabinet de conseil / direction de programme / Orange.

**Architecture : Python + python-pptx (plus de COM/PowerPoint).**

## Usage

```text
/ppt-v2 Genere 5 slides sur la data
/ppt-v2 @mon_deck.md
/ppt-v2 --strict Genere 3 slides
/ppt-v2 --help
```

| Flag | Effet |
|---|---|
| `--strict` | Corrections automatiques (texte, icones, conclusion) |
| `@fichier.md` | Genere depuis un fichier Markdown |
| `--help` | Affiche cette aide |

## Les 11 layouts

| Layout | Usage |
|---|---|
| `cover_orange` | Slide de titre : bande verticale orange, titre fort |
| `message_only` | Une phrase cle + chiffres en appui |
| `three_cards` | 3 cartes horizontales avec icones |
| `four_cards` | 4 cartes en grille 2x2 |
| `problem_solution` | Probleme / Solution avec fleche centrale |
| `process_horizontal` | 4-6 etapes avec ligne de progression |
| `kpi_context` | 3-4 chiffres cles en grand orange |
| `lessons_learned` | 5 enseignements numerotes |
| `role_focus` | Role central + responsabilites autour |
| `adoption_loop` | Boucle cyclique 4 phases |
| `closing` | Message de conclusion + 3 takeaways |

## Charte graphique

| Element | Valeur |
|---|---|
| Fond | Blanc #FFFFFF |
| Primaire | Orange #FF7900 |
| Texte titre | Gris #222222, Arial 30pt Gras |
| Texte carte | Gris #666666, Arial 12-14pt |
| Fond carte | Gris #F7F7F7 |
| Conclusion | Bandeau orange clair #FFF3E8 + barre orange |
| Footer | "C2 – Usage restreint" + page N / total |
| Marges | 1 pouce (2.54cm) de chaque cote |
| Taille slide | 1280 x 720 (16:9) |

## Pipeline de generation

```
Texte libre / Markdown / @fichier
  → parser.py       : analyse, decoupage, typage
  → quality.py       : validation + corrections --strict
  → slidebuilder.py  : construction python-pptx
  → layouts.py       : rendu des 11 layouts
  → Presentation.pptx
```

## Installer

```powershell
# Depuis le projet source
.\ppt-v2\install.ps1

# L'installeur verifie aussi python-pptx
```

Pre-requis : Python 3.12+ avec `python-pptx` installe.

```bash
pip install python-pptx
```

## Modules Python

```
.opencode/scripts/ppt-v2/
  generate.py     → CLI entry point
  theme.py        → charte graphique
  card.py         → composant carte
  icons.py        → 25 SVG lineaires
  footer.py       → footer C2
  conclusion.py   → bandeau orange
  layouts.py      → 11 builders
  slidebuilder.py → assembleur
  parser.py       → analyse texte/MD
  quality.py      → validation
```

## Exemple

```json
{
  "cover": {
    "deckTitle": "Transformation Digitale",
    "deckSubtitle": "REX et enseignements",
    "date": "Juin 2026",
    "authors": "Direction"
  },
  "slides": [
    {
      "title": "Contexte",
      "type": "contexte",
      "layout": "kpi_context",
      "kpis": [
        { "value": "85%", "label": "Adoption", "icon": "users" }
      ],
      "conclusion": "La comprehension du contexte est cle."
    }
  ]
}
```
