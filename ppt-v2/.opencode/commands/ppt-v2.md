---
description: Prepare puis genere une presentation PowerPoint avec des cartes visuelles (blocs autonomes en grille).
agent: build
---

Tu aides a preparer puis generer une presentation PowerPoint avec des cartes visuelles, en respectant strictement un template `.pptx`.

Travaille depuis la racine du projet et utilise :

- Le dossier courant du projet est la racine de travail.
- Le dossier `ppt-v2/` sert aux operations et fichiers intermediaires Markdown/JSON.
- `.opencode/scripts/ppt-v2/New-CardTemplate.ps1` pour creer un template cards.
- `.opencode/scripts/ppt-v2/Get-CardLayouts.ps1` pour generer le catalogue de layouts depuis le template.
- `.opencode/scripts/ppt-v2/Generate-CardPpt.ps1` pour generer le PowerPoint final avec cartes.

Demande utilisateur :

```text
$ARGUMENTS
```

Modes supportes :

- Aide : `/ppt-v2 --help`
- Installation dependances : `/ppt-v2 --install`
- Creer un template cards : `/ppt-v2 --create-template`
- Texte libre : `/ppt-v2 Genere 5 slides avec cartes sur les piliers de la data`
- Source Markdown : `/ppt-v2 @source.md`

Mode aide :

- Si la demande est exactement `/ppt-v2 --help`, ne lance pas le workflow.
- Lis et affiche le contenu complet de `.opencode/scripts/ppt-v2/README.md`.

Mode creation du template :

- Si la demande est `/ppt-v2 --create-template`, ne lance pas le workflow.
- Execute le bloc PowerShell ci-dessous.

Commande pour `/ppt-v2 --create-template` :

```powershell
$ErrorActionPreference = 'Stop'
$nom = Read-Host "Nom du template (sans .template.pptx)"
$templateDir = Join-Path (Get-Location) "ppt-v2"
New-Item -ItemType Directory -Force -Path $templateDir | Out-Null
$templatePath = Join-Path $templateDir "$nom.template.pptx"
powershell -NoProfile -ExecutionPolicy Bypass -File ".\.opencode\scripts\ppt-v2\New-CardTemplate.ps1" -OutputPath $templatePath
Write-Output "Template cree: $templatePath"
Write-Output "Tu peux maintenant l ouvrir dans PowerPoint pour personnaliser les couleurs et polices."
```

Regles strictes :

- Affiche toujours l'etape courante en gras avec le format `**Étape X/8 - Nom de l'étape**`.
- Utilise toujours ces 8 etapes :
  1. `**Étape 1/8 - Analyse de la demande**`
  2. `**Étape 2/8 - Recherche et choix du template**`
  3. `**Étape 3/8 - Préparation du Markdown**`
  4. `**Étape 4/8 - Validation utilisateur**`
  5. `**Étape 5/8 - Génération du JSON**`
  6. `**Étape 6/8 - Génération PowerPoint**`
  7. `**Étape 7/8 - Contrôle visuel et corrections**`
  8. `**Étape 8/8 - Finalisation et ouverture optionnelle**`
- Au premier tour, arrete-toi toujours a l'etape 4 apres avoir demande validation du Markdown.
- Genere le catalogue de layouts depuis le template avec `Get-CardLayouts.ps1` avant la generation PowerPoint.
- Si aucun template n'est fourni, cherche les fichiers `*.template.pptx` dans `ppt-v2/` et `template-ppt/`.

Format attendu pour `ppt-v2/deck.generated.md` :

```markdown
# Titre de la presentation

Template: chemin/vers/template.template.pptx
Source: texte libre

## Slide 1 - Cover
Layout: cover

Titre:
Sous-titre:
Date:
Auteurs:

## Slide 2 - Trois piliers
Layout: card_3

Titre: Les piliers de la data
Sous-titre: Nos axes stratégiques

Carte 1:
Titre: Gouvernance
- Cadre reglementaire
- Qualite des donnees

Carte 2:
Titre: Infrastructure
- Plateforme cloud
- Securite

Carte 3:
Titre: Culture
- Formation
- Accompagnement
```

Format JSON :

```json
{
  "cover": {
    "deckTitle": "...",
    "deckSubtitle": "...",
    "date": "...",
    "authors": "..."
  },
  "slides": [
    {
      "layout": "card_3",
      "placeholders": {
        "{{TITLE}}": "Les piliers",
        "{{SUBTITLE}}": "Nos axes"
      },
      "cards": [
        { "title": "Gouvernance", "body": ["Cadre", "Qualite"] },
        { "title": "Infrastructure", "body": ["Cloud", "Securite"] },
        { "title": "Culture", "body": ["Formation", "Accompagnement"] }
      ]
    }
  ]
}
```

Commande normale :

```powershell
$script = Join-Path (Get-Location) '.opencode\scripts\ppt-v2\Generate-CardPpt.ps1'
$rawArgs = '$ARGUMENTS'
$output = & powershell -NoProfile -ExecutionPolicy Bypass -Command "& '$script' -RawArgs '$rawArgs'"
@($output) -join "`r`n"
```
