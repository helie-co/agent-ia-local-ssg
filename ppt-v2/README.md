# /ppt-v2

Commande pour preparer puis generer une presentation PowerPoint avec des **cartes visuelles** (blocs autonomes en grille).

## Usage rapide

```text
/ppt-v2 --create-template           Creer un template cards
/ppt-v2 --help                      Afficher cette aide
/ppt-v2 Genere 3 slides cartes sur les piliers data
/ppt-v2 @deck.generated.md          Depuis un Markdown existant
```

## Installation

Depuis OpenCode Desktop, dans le projet cible, demander :

```text
Installer uniquement la commande projet /ppt-v2 depuis https://github.com/helie-co/agent-ia-local-ssg/tree/main/ppt-v2. OpenCode Desktop est deja installe, ne pas l installer. Ne pas utiliser git.
```

Apres installation, creer un template cards :

```text
/ppt-v2 --create-template
```

Puis redemarrer OpenCode Desktop.

## Cartes

Chaque **carte** est un bloc visuel autonome avec titre + contenu, arrange en grille :

- `card_2` : 2 cartes cote a cote
- `card_3` : 3 cartes en ligne
- `card_4` : grille 2x2
- `card_6` : grille 3x2

Les cartes sont generees a partir d'un template `.template.pptx` cree avec `--create-template`.

## Workflow

1. Analyse de la demande
2. Recherche du template
3. Preparation du Markdown avec cartes
4. Validation utilisateur
5. Generation du JSON
6. Generation PowerPoint
7. Controle visuel et corrections
8. Finalisation

## Formats

Markdown :

```markdown
## Slide 2 - Trois piliers
Layout: card_3

Titre: Les piliers
Sous-titre: Nos axes

Carte 1:
Titre: Gouvernance
- Reglementation
- Qualite

Carte 2:
Titre: Infrastructure
- Cloud
- Securite

Carte 3:
Titre: Culture
- Formation
- Changement
```

JSON :

```json
{
  "layout": "card_3",
  "placeholders": { "{{TITLE}}": "Les piliers" },
  "cards": [
    { "title": "Gouvernance", "body": ["Reglementation", "Qualite"] },
    { "title": "Infrastructure", "body": ["Cloud", "Securite"] },
    { "title": "Culture", "body": ["Formation", "Changement"] }
  ]
}
```

## Template

Le template est un fichier `.template.pptx` contenant :

- Slide 1 : **cover** (`{{DECK_TITLE}}`, `{{DECK_SUBTITLE}}`, `{{DATE}}`, `{{AUTHORS}}`)
- Slide 2 : **card_2** (`{{TITLE}}`, `{{CARD1_TITLE}}`, `{{CARD1_BODY}}`, `{{CARD2_TITLE}}`, `{{CARD2_BODY}}`)
- Slide 3 : **card_3** (idem avec 3 cartes)
- Slide 4 : **card_4** (4 cartes en 2x2)
- Slide 5 : **card_6** (6 cartes en 3x2)
- Slide 6 : **classic_bullets** (`{{TITLE}}`, `{{BODY}}`)

Ouvrir le template dans PowerPoint pour personnaliser les couleurs, polices et styles des cartes.

## Dossiers

- `ppt-v2/` : fichiers intermediaires Markdown, JSON et template
- Sortie finale : racine du projet
