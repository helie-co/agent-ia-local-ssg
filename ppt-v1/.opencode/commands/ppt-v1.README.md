# /ppt-v1

Commande pour preparer puis generer une presentation PowerPoint en respectant strictement un template `.pptx`.

## Usage rapide

```text
/ppt-v1 Genere 5 slides de presentation sur l'IA
/ppt-v1 @source.md
/ppt-v1 @photo.png
/ppt-v1 Reprends les slides de @source.pptx avec le template @template.pptx
/ppt-v1 --template chemin/vers/MonTemplate.template.pptx Genere 5 slides sur l'IA
/ppt-v1 --install
/ppt-v1 --help
```

## Modes supportes

- Texte libre : cree un Markdown de presentation a partir de la demande.
- Source Markdown ou texte : lit un fichier `.md` ou `.txt`, puis demande le niveau de reprise.
- Source image : lit une image `.png`, `.jpg`, `.jpeg`, `.bmp`, `.tif` ou `.tiff` avec OCR Tesseract.
- Reprise PPTX : extrait les textes, tableaux et textes OCR des images d'un `.pptx`.
- Template explicite : utilise le template fourni avec `--template`.
- Installation OCR : installe Tesseract et la langue francaise avec `/ppt-v1 --install`.

## Workflow

La commande travaille en deux phases.

Premiere phase : preparation Markdown.

- Analyse de la demande.
- Recherche ou validation du template.
- Extraction eventuelle de la source.
- Demande du niveau de reprise pour les sources fichier.
- Production de `ppt-v1/deck.generated.md`.
- Affichage du Markdown pour validation.

Deuxieme phase : generation PowerPoint apres validation explicite.

- Conversion vers `ppt-v1/slides.generated.json`.
- Generation du `.pptx` final a la racine du projet.
- Controle visuel avec export dans `ppt-v1/visual-check/`.
- Corrections puis regeneration si necessaire.

## Niveaux de reprise

Pour les sources `.pptx`, `.md`, `.txt` et images, la commande demande :

```text
Quel niveau de reprise veux-tu pour le contenu source ? Reponds 1, 2 ou 3 : 1 exact, 2 synthetique, 3 tres synthetique.
```

- `1 exact` : conserve le contenu au plus proche, nettoie et adapte aux layouts.
- `2 synthetique` : garde la structure et les messages cles, reduit les details.
- `3 tres synthetique` : produit une version courte orientee executif.

## Choix Des Layouts

- `/ppt-v1` choisit dynamiquement le layout de chaque slide avant d'ecrire le Markdown et avant de generer le JSON.
- Ordre de priorite : `section_header`, puis `two_columns_bullets`, puis `table_dynamic`, puis `classic_bullets`.
- `section_header` sert aux slides de chapitre ou transition courte.
- Les oppositions binaires comme `Avantages / Inconvenients`, `pour / contre`, `forces / limites`, `opportunites / risques` utilisent `two_columns_bullets`.
- `table_dynamic` est reserve aux tableaux, matrices, grilles ou comparaisons avec colonnes homogenes, sans plafond strict de lignes ou colonnes.
- Si le controle visuel montre un tableau illisible, coupe ou trop dense, reduire, scinder ou synthetiser le tableau puis regenerer.
- Ne pas utiliser `table_dynamic` pour une opposition binaire simple, sauf demande explicite de tableau par l'utilisateur.
- `classic_bullets` est le fallback pour les listes lineaires, recommandations ou contenus desequilibres.

## Dossiers et fichiers

- Racine projet : emplacement du `.pptx` final.
- `ppt-v1/` : fichiers intermediaires Markdown, JSON, OCR et controle visuel.
- Templates : peuvent etre dans la racine projet ou n'importe quel sous-dossier.
- Sources utilisateur : peuvent etre dans la racine projet ou n'importe quel sous-dossier.

Fichiers intermediaires principaux :

- `ppt-v1/deck.generated.md`
- `ppt-v1/slides.generated.json`
- `ppt-v1/deck.extracted.md`
- `ppt-v1/image.extracted.md`
- `ppt-v1/visual-check-report.json`

Si `ppt-v1/deck.generated.md` ou `ppt-v1/slides.generated.json` n'existe pas encore, la commande doit creer le fichier complet. Elle ne doit pas tenter de patch/update sur un fichier absent.

## OCR

`/ppt-v1 --install` installe :

- Tesseract OCR via `winget` package `tesseract-ocr.tesseract`.
- La langue francaise `fra.traineddata`.

OCR supporte :

- images directes : `.png`, `.jpg`, `.jpeg`, `.bmp`, `.tif`, `.tiff` ;
- images/photos/captures presentes dans un `.pptx`.

Si une photo est jointe mais qu'aucun chemin fichier n'est disponible, enregistre-la dans le repertoire courant du projet ou un sous-dossier, puis relance par exemple :

```text
/ppt-v1 @photo.png
/ppt-v1 @docs/photo.png
```

## Validation

La generation PowerPoint ne demarre jamais au premier tour. Elle demarre uniquement apres validation explicite du Markdown avec une reponse comme :

```text
OK
oui
ca marche
OK genere
```

## Cover

- La cover contient toujours une date.
- Si aucune date n'est fournie par l'utilisateur, utiliser la date du jour au format `D MOIS AAAA`, avec le mois en majuscules.
- Pour le 8 juin 2026, ecrire exactement `8 JUIN 2026`.
- Ne pas utiliser un format numerique comme `08/06/2026` sauf demande explicite.

## Footer

- `{{FOOTER_TITLE}}` est gere par defaut par le generateur.
- Sa valeur par defaut est le titre de la presentation (`cover.deckTitle`).
- Une slide peut fournir explicitement `{{FOOTER_TITLE}}` dans ses placeholders JSON pour remplacer cette valeur.

## Execution PowerShell

Les scripts `.ps1` doivent toujours etre lances avec :

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\.opencode\scripts\ppt-v1\Generate-Ppt.ps1 ...
```

Ne lance pas directement un script `.ps1` avec `& .\script.ps1`, car la politique d'execution Windows peut bloquer l'execution directe.

## Templates

Si aucun template n'est fourni, `/ppt-v1` cherche les fichiers :

```text
**/*.template.pptx
```

Si plusieurs templates sont trouves, la commande demande lequel utiliser.

Catalogue de layouts :

- Par defaut, `/ppt-v1` genere automatiquement le catalogue depuis le template choisi avec `.opencode/scripts/ppt-v1/Get-PptLayouts.ps1`.
- Le catalogue detecte est ecrit dans `ppt-v1/layouts.detected.json`.
- Ce catalogue ne contient que des layouts dont la slide source existe vraiment dans le template choisi.
- Les catalogues statiques `.opencode/scripts/ppt-v1/layouts.json` et `.opencode/scripts/ppt-v1/SopraSteriaNext.layouts.json` sont des fallbacks/debug uniquement.

## Auteur

Si aucun auteur n'est fourni par l'utilisateur, le champ `authors` reste vide dans le JSON. Le script de generation utilise alors l'auteur Office/Windows par defaut. Ne jamais mettre `OpenCode`, `Assistant`, `IA` ou le nom de l'outil comme auteur.

## Ouverture

Le fichier PowerPoint final n'est pas ouvert automatiquement. La commande demande d'abord confirmation, puis ouvre le fichier uniquement si l'utilisateur repond explicitement `OK`, `oui`, `ca marche`, `ouvre` ou equivalent.

## Sortie finale

Le fichier PowerPoint final est genere a la racine du projet avec un nom derive du titre de la presentation.

Le nom final ne doit pas contenir :

- `Generated`
- `generated`
- `genere`
- `généré`
