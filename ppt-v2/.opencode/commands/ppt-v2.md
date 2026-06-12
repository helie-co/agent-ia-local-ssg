---
description: Genere une presentation PowerPoint professionnelle avec 11 layouts Orange / conseil, icones SVG, auto-conclusion et controle qualite.
agent: build
---

Tu aides a preparer puis generer une presentation PowerPoint avec des cartes visuelles, en respectant strictement un template `.pptx`.

Travaille depuis la racine du projet et utilise :
- Le dossier `ppt-v2/` pour les fichiers intermediaires Markdown/JSON.
- `.opencode/scripts/ppt-v2/New-ProfessionalTemplate.ps1` pour creer un template.
- `.opencode/scripts/ppt-v2/Get-ProfessionalLayouts.ps1` pour generer le catalogue de layouts.
- `.opencode/scripts/ppt-v2/Generate-ProfessionalPpt.ps1` pour generer le PowerPoint final.

Demande utilisateur :
```text
$ARGUMENTS
```

Modes supportes :
- Aide : `/ppt-v2 --help`
- Creer template : `/ppt-v2 --create-template`
- Texte libre : `/ppt-v2 Genere 5 slides sur la data`
- Source Markdown : `/ppt-v2 @source.md`

Mode aide : lit et affiche `.opencode/scripts/ppt-v2/README.md`.

Mode `--create-template` :
```powershell
$ErrorActionPreference = 'Stop'
$nom = Read-Host "Nom du template (sans .template.pptx)"
$templateDir = Join-Path (Get-Location) "ppt-v2"
New-Item -ItemType Directory -Force -Path $templateDir | Out-Null
$templatePath = Join-Path $templateDir "$nom.template.pptx"
powershell -NoProfile -ExecutionPolicy Bypass -File ".\.opencode\scripts\ppt-v2\New-ProfessionalTemplate.ps1" -OutputPath $templatePath
Write-Output "Template cree: $templatePath (11 layouts Orange)"
Write-Output "Ouvre dans PowerPoint pour personnaliser couleurs, polices et ajouter des icones."
```

Regles strictes :
- Affiche chaque etape en gras : `**Étape X/9 - Nom**`.
- 9 etapes : 1 Analyse, 2 Type de chaque slide, 3 Template, 4 Markdown, 5 Validation, 6 JSON, 7 Generation PPT, 8 Controle qualite, 9 Finalisation.
- Au premier tour, arrete a l'etape 5 apres validation du Markdown.
- Genere le catalogue layouts avec `Get-ProfessionalLayouts.ps1`.
- Cherche `*.template.pptx` dans `ppt-v2/` et `template-ppt/`.
- Pour chaque slide, identifie le type : contexte, problematique, question, reponse, processus, roles, enseignements, chiffres cles, bilan, conclusion.
- Applique la reduction automatique du texte (max 12 mots/carte, 4 lignes/bloc).
- Ajoute systematiquement un message de conclusion (phrase courte en bas de slide).
- Qualite : verifie texte trop long, >5 blocs, conclusion absente, coherence.

Formats Markdown :
```markdown
# Titre presentation
Template: ppt-v2/MonTemplate.template.pptx
Source: texte libre

## Slide 1 - Contexte
Type: contexte
Layout: context_kpi_slide
Titre: Contexte du projet
KPI 1: 85% | Taux adoption
KPI 2: 500+ | Utilisateurs

## Slide 2 - Question cle
Type: question
Layout: question_answer_slide
Titre: Question cle
Question: Comment industrialiser ?
Reponse: Une plateforme unique
Point 1: Approche progressive
Point 2: Gouvernance transverse

## Slide 3 - Processus
Type: processus
Layout: process_slide
Titre: Notre approche
Etape 1: Diagnostic | Audit des pratiques
Etape 2: Pilote | Deploiement rapide

## Slide 4 - Enseignements
Type: enseignements
Layout: lessons_slide
Titre: Enseignements cles
Lecon 1: Commencer petit, penser grand
Lecon 2: La gouvernance se co-construit
```

Format JSON (structure complete dans slides.example.json) :
```json
{
  "cover": { "deckTitle": "...", "deckSubtitle": "...", "date": "...", "authors": "..." },
  "slides": [
    { "type": "contexte", "layout": "context_kpi_slide",
      "placeholders": { "{{TITLE}}": "...", "{{SUBTITLE}}": "..." },
      "kpis": [{ "value": "85%", "label": "Taux adoption" }] },
    { "type": "enseignements", "layout": "lessons_slide",
      "lessons": ["Premier enseignement", "Second enseignement"] },
    { "type": "conclusion", "layout": "conclusion_slide",
      "mainMessage": "La transformation est un voyage",
      "takeaways": ["Point cle 1", "Point cle 2", "Point cle 3"] }
  ]
}
```

Commande normale :
```powershell
$script = Join-Path (Get-Location) '.opencode\scripts\ppt-v2\Generate-ProfessionalPpt.ps1'
$rawArgs = '$ARGUMENTS'
$output = & powershell -NoProfile -ExecutionPolicy Bypass -Command "& '$script' -RawArgs '$rawArgs'"
@($output) -join "`r`n"
```
