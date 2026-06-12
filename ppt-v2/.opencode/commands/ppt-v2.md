---
description: Genere une presentation PowerPoint professionnelle avec 11 layouts Orange / conseil, icones SVG, auto-conclusion et controle qualite.
agent: build
---

Tu aides a preparer puis generer une presentation PowerPoint professionnelle de niveau cabinet de conseil.

Travaille depuis la racine du projet et utilise :
- Le dossier `ppt-v2/` pour les fichiers intermediaires.
- `.opencode/scripts/ppt-v2/generate.py` pour generer le PowerPoint final.

Demande utilisateur :
```text
$ARGUMENTS
```

Modes supportes :
- Texte libre : `/ppt-v2 Genere 5 slides sur la data`
- Source Markdown : `/ppt-v2 @source.md`
- Mode strict : `/ppt-v2 --strict Genere 3 slides`
- Aide : `/ppt-v2 --help`

Pre-requis : Python 3.12+ avec `python-pptx` installe.

Regles strictes :
- Affiche chaque etape en gras : `**Étape X/9 - Nom**`.
- 9 etapes : 1 Analyse, 2 Type de chaque slide, 3 Template, 4 Markdown/JSON, 5 Validation, 6 Reduction texte, 7 Generation PPT, 8 Controle qualite, 9 Finalisation.
- Au premier tour, arrete a l'etape 5 apres validation du Markdown.
- Pour chaque slide, identifie le type : contexte, problematique, question, processus, roles, enseignements, chiffres cles, adoption, bilan.
- Genere un JSON intermediaire avant generation PPT.
- Applique la reduction automatique du texte (max 12 mots/carte).
- Ajoute systematiquement un message de conclusion en bas de slide.
- Qualite : verifie titre absent, conclusion absente, +4 cartes, texte trop long, icone manquante.

Architecture logicielle (Python / python-pptx) :
- `theme.py` : charte graphique (orange #FF7900, Arial, marges)
- `card.py` : composant carte homogene
- `icons.py` : 25 SVG lineaires
- `footer.py` : footer "C2 – Usage restreint"
- `conclusion.py` : bandeau orange conclusion
- `layouts.py` : 11 layouts (cover_orange, message_only, three_cards, four_cards, problem_solution, process_horizontal, kpi_context, lessons_learned, role_focus, adoption_loop, closing)
- `slidebuilder.py` : assembleur slide
- `parser.py` : analyse texte libre / Markdown
- `quality.py` : validation stricte
- `generate.py` : point d'entree CLI

Formats Markdown acceptes :
```markdown
# Titre presentation
Source: texte libre

## Slide 1 - Contexte
Type: contexte
KPI 1: 85% | Taux adoption
KPI 2: 500+ | Utilisateurs

## Slide 2 - Processus
Type: processus
Etape 1: Diagnostic | Audit des pratiques
Etape 2: Pilote | Deploiement rapide
```

Format JSON intermediaire :
```json
{
  "cover": { "deckTitle": "...", "deckSubtitle": "...", "date": "...", "authors": "..." },
  "slides": [
    {
      "title": "Titre clair",
      "type": "contexte",
      "layout": "kpi_context",
      "cards": [{ "title": "...", "text": "...", "icon": "data" }],
      "conclusion": "Phrase courte et forte."
    }
  ]
}
```

Commande de generation :
```powershell
$ErrorActionPreference = 'Stop'
$python = "C:\Users\jfhelie\AppData\Local\Programs\Python\Python312\python.exe"
$script = Join-Path (Get-Location) '.opencode\scripts\ppt-v2\generate.ps1'

# Generer le contenu avec le JSON intermediaire, puis appeler generate.py :
$rawArgs = '$ARGUMENTS'
$scriptPy = Join-Path (Get-Location) '.opencode\scripts\ppt-v2\generate.py'
& $python $scriptPy $rawArgs 2>&1
```
