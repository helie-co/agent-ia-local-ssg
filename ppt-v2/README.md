# /ppt-v2

Commande pour generer des presentations PowerPoint professionnelles avec 11 layouts Orange / conseil, icones SVG, auto-conclusion et controle qualite.

## Usage rapide

```text
/ppt-v2 --create-template           Creer un template (11 layouts)
/ppt-v2 --help                      Afficher cette aide
/ppt-v2 Genere 5 slides sur la data
/ppt-v2 @deck.md                    Depuis un Markdown existant
```

## Installation

Depuis OpenCode Desktop, dans le projet cible, demander :

```text
Installer uniquement la commande projet /ppt-v2 depuis https://github.com/helie-co/agent-ia-local-ssg/tree/main/ppt-v2. OpenCode Desktop est deja installe, ne pas l installer. Ne pas utiliser git.
```

Apres installation, creer un template :

```text
/ppt-v2 --create-template
```

Puis redemarrer OpenCode Desktop.

## Les 11 layouts

| Layout | Type de slide | Description |
|---|---|---|
| `title_slide` | Cover | Titre, sous-titre, date, auteur |
| `context_kpi_slide` | Contexte / chiffres cles | 4 KPI avec valeur + etiquette |
| `problem_slide` | Problematique | Probleme a gauche, impacts a droite |
| `question_answer_slide` | Question / reponse | Question, reponse cle, 3 points |
| `three_cards_slide` | 3 cartes | 3 blocs horizontaux avec titre + corps |
| `four_cards_slide` | 4 cartes | Grille 2x2 avec titre + corps |
| `process_slide` | Processus | 4 etapes numerotees avec fleches |
| `role_focus_slide` | Focus role | Nom, titre, 4 responsabilites |
| `lessons_slide` | Enseignements | 4 lecons numerotees |
| `adoption_loop_slide` | Boucle d'adoption | 4 phases en cycle |
| `conclusion_slide` | Bilan / conclusion | Message cle + 3 takeaways |

## Charte graphique

- Couleur principale : orange (`#FF7900`)
- Secondaires : gris fonce, gris clair, bleu discret
- Typographie : Calibri
- Coins arrondis, ombres legeres, espacement genereux
- Bandeau conclusion orange en bas de chaque slide

## Workflow

1. Analyse de la demande
2. Typage automatique de chaque slide (contexte, probleme, processus, etc.)
3. Recherche et detection du template
4. Preparation du Markdown enrichi
5. Validation utilisateur
6. Generation du JSON
7. Generation PowerPoint (avec reduction texte, auto-conclusion, selection layout)
8. Controle qualite automatique (longueur, blocs, conclusion, coherence)
9. Finalisation

## Fonctionnalites automatiques

- **Reduction texte** : max 12 mots par carte, 4 lignes par bloc
- **Conclusion auto** : phrase d'accroche generee selon le type de slide
- **Selection layout** : automatique selon le type de slide
- **Nettoyage placeholders** : les espaces iconographiques non renseignes sont vides
- **Icons SVG** : 25 icones lineaires dans `icons/` (users, data, IA, decision, etc.)

## Dossiers

- `ppt-v2/` : fichiers intermediaires Markdown, JSON et template
- Sortie finale : racine du projet
