---
description: Prepare puis genere une presentation PowerPoint depuis texte, Markdown ou PPTX source.
agent: build
---

Tu aides a preparer puis generer une presentation PowerPoint en respectant strictement un template `.pptx`.

Travaille depuis la racine du projet et utilise :

- Le dossier courant du projet est la racine de travail.
- Les templates peuvent etre dans le dossier courant ou n'importe quel sous-dossier.
- `ppt-v1/` sert aux operations et fichiers intermediaires Markdown/JSON/OCR/controle visuel.
- Les sources utilisateur et templates peuvent rester dans le dossier courant du projet ou n'importe quel sous-dossier ; ne les deplace pas automatiquement.
- Le fichier PowerPoint final `.pptx` doit etre genere dans le dossier courant du projet, pas dans `ppt-v1/` et pas forcement dans le dossier du template.
- `.opencode/scripts/ppt-v1/Generate-Ppt.ps1` pour generer le PowerPoint.
- `.opencode/scripts/ppt-v1/Get-PptLayouts.ps1` pour generer automatiquement le catalogue de layouts depuis le template choisi.
- `.opencode/scripts/ppt-v1/Extract-PptContent.ps1` pour extraire le contenu d'un PowerPoint source, y compris le texte OCR des photos/images si Tesseract est installe.
- `.opencode/scripts/ppt-v1/Extract-ImageOcr.ps1` pour extraire le texte OCR d'une image source directe.
- `.opencode/scripts/ppt-v1/Test-PptVisual.ps1` pour verifier visuellement le PowerPoint genere.
- Par defaut, ne choisis pas un catalogue statique a la main : genere `ppt-v1/layouts.detected.json` depuis le template reellement choisi avec `Get-PptLayouts.ps1`.
- Les catalogues statiques `.opencode/scripts/ppt-v1/layouts.json` et `.opencode/scripts/ppt-v1/SopraSteriaNext.layouts.json` sont uniquement des fallbacks/debug si la detection automatique echoue.
- Avant generation, utilise le catalogue detecte `ppt-v1/layouts.detected.json`; il ne contient que des layouts dont le `sourceSlide` existe dans le template choisi.

Demande utilisateur :

```text
$ARGUMENTS
```

Modes supportes :

- Aide : `/ppt-v1 --help`
- Installation OCR : `/ppt-v1 --install`
- Texte libre : `/ppt-v1 Genere 5 slides de presentation sur l'IA`
- Source Markdown ou texte : `/ppt-v1 @source.md` ou `/ppt-v1 @docs/source.txt`
- Source image OCR : `/ppt-v1 @photo.png` ou `/ppt-v1 @docs/capture.jpg`
- Reprise PPTX : `/ppt-v1 Reprends les slides de @source.pptx avec le template @template.pptx`
- Template explicite : `/ppt-v1 --template chemin/vers/MonTemplate.template.pptx Genere 5 slides sur l'IA`
- Template Sopra : `/ppt-v1 --template template-ppt/SopraSteriaNext.generated-template.pptx Genere 3 slides sur l'IA`

Mode aide :

- Si la demande est exactement `/ppt-v1 --help`, ne lance pas le workflow PowerPoint en 8 etapes.
- Lis et affiche le contenu complet de `.opencode/scripts/ppt-v1/README.md`.
- Si le fichier README est introuvable, affiche une aide courte avec les modes supportes et indique que `.opencode/scripts/ppt-v1/README.md` est manquant.

Mode installation :

- Si la demande est exactement `/ppt-v1 --install`, ne lance pas le workflow PowerPoint en 8 etapes.
- Verifie que `tesseract` est disponible en ligne de commande.
- Si `tesseract` est introuvable, installe Tesseract OCR avec `winget` via le package `tesseract-ocr.tesseract`.
- Recharge le `PATH` de la session apres installation.
- Si le `PATH` n'est pas encore a jour, teste aussi `C:\Users\<user>\AppData\Local\Programs\Tesseract-OCR\tesseract.exe`, `C:\Program Files\Tesseract-OCR\tesseract.exe` et `C:\Program Files (x86)\Tesseract-OCR\tesseract.exe`.
- Verifie que la langue francaise OCR `fra.traineddata` est disponible dans le dossier `tessdata` de Tesseract.
- Si `fra.traineddata` est absent, telecharge-le depuis `https://raw.githubusercontent.com/tesseract-ocr/tessdata_fast/main/fra.traineddata` vers le dossier `tessdata` de Tesseract.
- Termine en affichant le chemin de `tesseract.exe`, le chemin de `fra.traineddata` et la version.
- Si `tesseract.exe` reste introuvable apres installation, indique de redemarrer OpenCode ou le terminal puis de relancer `/ppt-v1 --install`.

Commande a executer pour `/ppt-v1 --install` :

```powershell
$ErrorActionPreference = 'Stop'

$tesseractCmd = Get-Command tesseract -ErrorAction SilentlyContinue
$tesseractPath = if ($tesseractCmd) { $tesseractCmd.Source } else { $null }

if (-not $tesseractPath) {
  $candidatePaths = @(
    (Join-Path $env:LOCALAPPDATA 'Programs\Tesseract-OCR\tesseract.exe'),
    'C:\Program Files\Tesseract-OCR\tesseract.exe',
    'C:\Program Files (x86)\Tesseract-OCR\tesseract.exe'
  )

  foreach ($candidatePath in $candidatePaths) {
    if (-not $tesseractPath -and (Test-Path -LiteralPath $candidatePath)) {
      $tesseractPath = $candidatePath
    }
  }
}

if (-not $tesseractPath) {
  winget install -e --id tesseract-ocr.tesseract --accept-source-agreements --accept-package-agreements
  $env:Path = [System.Environment]::GetEnvironmentVariable('Path', 'Machine') + ';' + [System.Environment]::GetEnvironmentVariable('Path', 'User')

  $tesseractCmd = Get-Command tesseract -ErrorAction SilentlyContinue
  $tesseractPath = if ($tesseractCmd) { $tesseractCmd.Source } else { $null }

  $candidatePaths = @(
    (Join-Path $env:LOCALAPPDATA 'Programs\Tesseract-OCR\tesseract.exe'),
    'C:\Program Files\Tesseract-OCR\tesseract.exe',
    'C:\Program Files (x86)\Tesseract-OCR\tesseract.exe'
  )

  foreach ($candidatePath in $candidatePaths) {
    if (-not $tesseractPath -and (Test-Path -LiteralPath $candidatePath)) {
      $tesseractPath = $candidatePath
    }
  }
}

if (-not $tesseractPath) {
  throw 'tesseract.exe est introuvable apres installation. Redemarre OpenCode ou le terminal puis relance /ppt-v1 --install.'
}

$tessDataDirectory = Join-Path (Split-Path -Parent $tesseractPath) 'tessdata'
if (-not (Test-Path -LiteralPath $tessDataDirectory)) {
  New-Item -ItemType Directory -Path $tessDataDirectory | Out-Null
}

$fraDataPath = Join-Path $tessDataDirectory 'fra.traineddata'
if (-not (Test-Path -LiteralPath $fraDataPath)) {
  $fraDataUrl = 'https://raw.githubusercontent.com/tesseract-ocr/tessdata_fast/main/fra.traineddata'
  Invoke-WebRequest -Uri $fraDataUrl -OutFile $fraDataPath
}

if (-not (Test-Path -LiteralPath $fraDataPath)) {
  throw 'fra.traineddata est introuvable apres installation. Verifie la connexion internet puis relance /ppt-v1 --install.'
}

"tesseract=$tesseractPath"
"fra.traineddata=$fraDataPath"
& $tesseractPath --version
```

Regles strictes :

- Affiche toujours l'etape courante en gras avec le format `**Étape X/8 - Nom de l'étape**`.
- Ne passe jamais silencieusement d'une phase a l'autre : annonce chaque etape avec une phrase courte de statut.
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
- Apres validation explicite, reprends a l'etape 5.
- Ne genere jamais le fichier PowerPoint au premier tour.
- Commence toujours par produire et afficher le contenu sous forme Markdown.
- Demande toujours validation avant de lancer la generation PowerPoint.
- Ne lance `.opencode/scripts/ppt-v1/Generate-Ppt.ps1` que si l'utilisateur confirme explicitement avec `OK`, `oui`, `ça marche`, `OK genere`, `OK génère`, `genere le PPT` ou une formule equivalente sans ambiguite.
- Si l'utilisateur demande des modifications, modifie le Markdown et redemande validation.
- Si la demande utilise une source `.pptx`, `.md`, `.txt`, `.png`, `.jpg`, `.jpeg`, `.bmp`, `.tif` ou `.tiff`, demande toujours le niveau de reprise du contenu avant de produire `ppt-v1/deck.generated.md`.
- Demande exactement : `Quel niveau de reprise veux-tu pour le contenu source ? Réponds 1, 2 ou 3 : 1 exact, 2 synthétique, 3 très synthétique.`
- `1 exact` signifie : conserver le contenu au plus proche de la source, nettoyer seulement la forme, adapter aux layouts du template et scinder en plusieurs slides si le contenu est trop dense.
- `2 synthétique` signifie : garder la structure générale, réduire les détails secondaires, conserver les messages clés, chiffres importants, décisions et actions, et résumer les grands tableaux.
- `3 très synthétique` signifie : produire une version courte orientée exécutif, garder uniquement les idées essentielles, transformer les longs tableaux en enseignements clés, limiter les bullets et privilégier les messages clairs.
- Si l'utilisateur a deja donne explicitement ce niveau dans sa demande (`1`, `exact`, `2`, `synthétique`, `3`, `très synthétique`), applique-le sans redemander.
- Si aucun template n'est fourni, cherche d'abord dans le dossier courant du projet et tous ses sous-dossiers les fichiers `*.template.pptx`.
- Pour cette recherche, utilise de preference le pattern `**/*.template.pptx`.
- Si un seul template est trouve, propose-le comme choix recommande et demande confirmation. Les reponses `OK`, `oui`, `ça marche` ou equivalents valident ce template.
- Si plusieurs templates sont trouves, affiche la liste des chemins relatifs et demande lequel utiliser.
- Si aucun template n'est trouve, demande a l'utilisateur de fournir le chemin d'un template `.pptx`.
- Quand l'utilisateur demande un nombre de slides, ce nombre correspond toujours aux slides de contenu, hors couverture.
- Ajoute donc une cover en plus du nombre demandé : `2 slides` signifie 1 cover + 2 slides de contenu, soit 3 slides PowerPoint au total.
- Pour la cover, le titre doit etre court, impactant et adapte a la zone du template.
- Titre de cover : 45 caracteres maximum recommande.
- Si le titre de cover depasse 45 caracteres, raccourcis-le automatiquement et deplace les elements explicatifs dans le sous-titre.
- Ne cree pas de titre de cover avec plusieurs segments separes par virgules ou deux-points si cela rend le titre long.
- Exemple de correction cover : `IA générative : comprendre, cadrer, déployer` devient `IA générative`, avec le sous-titre `Comprendre, cadrer et déployer les premiers usages`.
- Pour la cover, renseigne toujours une date.
- Si l'utilisateur ne fournit pas explicitement de date, utilise la date du jour au format `D MOIS AAAA`, avec le mois en majuscules.
- Pour la date courante du 8 juin 2026, ecris exactement `8 JUIN 2026`.
- Utilise cette date dans `ppt-v1/deck.generated.md` champ `Date:` et dans `ppt-v1/slides.generated.json` champ `cover.date`.
- N'utilise pas de format numerique comme `08/06/2026` sauf demande explicite de l'utilisateur.
- Le style appartient au template PowerPoint.
- Le contenu appartient a `ppt-v1/deck.generated.md` puis `ppt-v1/slides.generated.json`.
- Avant d'ecrire un fichier intermediaire Markdown ou JSON, verifie que le dossier `ppt-v1/` existe et cree-le si necessaire.
- Si `ppt-v1/deck.generated.md` n'existe pas, cree le fichier Markdown complet ; ne tente pas de faire un patch/update sur un fichier absent.
- Si `ppt-v1/slides.generated.json` n'existe pas, cree le fichier JSON complet ; ne tente pas de faire un patch/update sur un fichier absent.
- Si le fichier Markdown ou JSON existe deja, tu peux le modifier par patch ciblé en preservant les changements utiles.
- Le fichier PowerPoint final doit porter un nom explicite derive du titre de la presentation et etre cree a la racine du projet.
- N'utilise pas `Generated`, `generated`, `généré` ou `genere` dans le nom du fichier final.
- Exemple : pour `IA générative en entreprise`, utilise `./IA_generative_en_entreprise.pptx`.
- Si le nom derive existe deja, ajoute un suffixe numerique court comme `_v2`, jamais le mot `generated`.
- Si la demande est en français, conserve les accents et la typographie française dans le Markdown, le JSON et le PowerPoint.
- Écris les fichiers Markdown et JSON en UTF-8.
- Pour la cover, si l'utilisateur ne fournit pas d'auteur, laisse le champ auteur vide dans le JSON afin que le script recupere l'auteur par defaut de l'ordinateur/Office.
- Ne mets jamais `OpenCode`, `Assistant`, `IA` ou le nom de l'outil comme auteur. Si l'auteur n'est pas explicitement fourni par l'utilisateur, laisse `authors` vide.
- Ne change pas les polices, couleurs, positions ou formes dans le script.
- Pour `Orange.template.pptx`, utilise une seule `cover`, puis choisis entre `classic_bullets` et `two_columns_bullets` pour chaque slide de contenu.
- La slide classique utilise `{{TITLE}}`, `{{SUBTITLE}}`, `{{BODY}}`.
- `{{BODY}}` contient des bullets multi-lignes, sans limite technique stricte.
- La slide a deux colonnes utilise `{{TITLE}}`, `{{SUBTITLE}}`, `{{LEFT_TITLE}}`, `{{LEFT_BODY}}`, `{{RIGHT_TITLE}}`, `{{RIGHT_BODY}}`.
- Gere `{{FOOTER_TITLE}}` par defaut sur tous les layouts qui le contiennent : sa valeur est le titre de la presentation (`cover.deckTitle`).
- Si une slide fournit explicitement `{{FOOTER_TITLE}}` dans ses placeholders JSON, utilise cette valeur au lieu du titre de presentation.
- Pour `two_columns_bullets`, genere toujours un titre court pour chaque colonne.
- Utilise `classic_bullets` quand le contenu est une liste unique, une sequence lineaire, des etapes, des recommandations ou un message simple.
- Utilise `two_columns_bullets` quand le contenu se separe naturellement en deux categories equilibrees : avantages / inconvenients, avantages / risques, pour / contre, forces / limites, opportunites / risques, problemes / reponses, constats / actions, avant / apres, metier / IT, benefices / conditions de succes, ce que l'IA permet / ce qu'elle impose.
- Si la demande contient une opposition binaire comme `Avantages / Inconvénients`, `avantages et inconvénients`, `pour / contre`, `pros / cons`, `forces / limites` ou `opportunités / risques`, choisis obligatoirement `two_columns_bullets` sauf si l'utilisateur demande explicitement un tableau, une matrice ou une grille.
- N'utilise pas `two_columns_bullets` si une colonne est nettement plus pauvre que l'autre ; dans ce cas, reformule ou reste en `classic_bullets`.
- Utilise `table_dynamic` quand le contenu est une comparaison structuree avec des colonnes homogenes : cas d'usage / valeur / risque, action / responsable / echeance / statut, risques / parades / priorite, criteres de choix ou matrice simple.
- N'utilise pas `table_dynamic` pour une opposition binaire simple comme avantages / inconvenients, pour / contre ou forces / limites ; utilise `two_columns_bullets`.
- Utilise `table_dynamic` uniquement si l'utilisateur demande explicitement un tableau, une matrice, une grille, une liste structuree en colonnes homogenes, ou si le contenu exige naturellement plusieurs colonnes homogenes.
- Pour `table_dynamic`, il n'y a pas de plafond strict de colonnes ou de lignes : genere le tableau demande et laisse la verification visuelle signaler les problemes de rendu.
- Pour `table_dynamic`, le tableau doit avoir la meme largeur que la zone de titre de la slide et etre aligne a gauche sur ce titre ; ne l'elargis pas sur toute la slide si le titre est plus etroit.
- Si le controle visuel ou les images exportees montrent un tableau illisible, coupe, trop dense ou avec des retours de ligne disgracieux, corrige le contenu en reduisant, scindant ou synthetisant le tableau, puis regenere.
- En reprise `.pptx`, conserve dans le Markdown les tableaux extraits par `Extract-PptContent.ps1` sous forme `Colonnes:` / `Lignes:`.
- En reprise `.pptx`, conserve aussi dans le Markdown le texte OCR extrait des photos, images et captures d'ecran sous forme `Texte OCR image:`.
- Si l'OCR indique `tesseract introuvable`, demande a l'utilisateur de lancer `/ppt-v1 --install`, puis relance l'extraction avant de preparer le Markdown final.
- Ignore les images decoratives, logos et pictogrammes quand aucun texte utile n'est detecte.
- Pour une source image directe (`.png`, `.jpg`, `.jpeg`, `.bmp`, `.tif`, `.tiff`), extrais d'abord son texte avec `.opencode/scripts/ppt-v1/Extract-ImageOcr.ps1` vers `./ppt-v1/image.extracted.md`.
- Si l'utilisateur joint une photo/image en piece jointe et qu'un chemin fichier est disponible dans la demande ou le contexte, traite-la comme une source image directe.
- Si l'utilisateur joint une photo/image mais qu'aucun chemin fichier exploitable n'est disponible, ne tente pas d'OCR directement depuis le chat : demande d'enregistrer l'image dans le repertoire courant du projet ou un sous-dossier, puis de relancer `/ppt-v1 @photo.png` ou `/ppt-v1 @chemin/vers/photo.png`.
- Pour une source image directe, utilise le texte OCR comme source de contenu : en mode `1 exact`, garde ce texte au plus proche ; en mode `2 synthétique`, regroupe et clarifie les informations utiles ; en mode `3 très synthétique`, transforme le contenu en messages clés.
- Si l'image directe ne produit aucun texte OCR exploitable, indique-le clairement et demande une image plus lisible ou un autre format source.
- Si un tableau extrait devient illisible dans `table_dynamic` apres generation, ne le laisse pas tel quel : resume-le, scinde-le en plusieurs slides ou transforme-le en messages clés selon le template.
- Si `ppt-v1/deck.extracted.md` contient `Note extraction: tableau source ... a resumer ou scinder avant generation`, traite cette note comme une instruction de preparation et ne la recopie pas dans le PowerPoint final.
- Si `ppt-v1/deck.extracted.md` contient des blocs `Texte OCR image:`, utilise-les comme contenu source issu de l'image : en mode `1 exact`, garde ce texte au plus proche ; en mode `2 synthétique`, garde uniquement les informations utiles ; en mode `3 très synthétique`, transforme-le en messages clés.
- Pour une source `.md` ou `.txt`, ne suppose pas que le fichier source est deja le Markdown final : applique d'abord le niveau de reprise choisi, puis produis `ppt-v1/deck.generated.md` dans le format attendu.
- En mode `1 exact`, preserve les tableaux Markdown source autant que possible, puis scinde ou adapte uniquement si le template ne peut pas les recevoir lisiblement.
- En mode `2 synthétique`, condense les tableaux Markdown ou PowerPoint en lignes principales, totaux, ecarts significatifs ou messages clés.
- En mode `3 très synthétique`, transforme les tableaux longs en 3 a 5 enseignements clés, sauf si un tableau court est indispensable a la comprehension.
- Pour `SopraSteriaNext.generated-template.pptx`, les layouts disponibles au premier lot sont `cover`, `section_header` et `classic_bullets` avec `.opencode/scripts/ppt-v1/SopraSteriaNext.layouts.json`.
- Pour `.opencode/scripts/ppt-v1/SopraSteriaNext.template.pptx`, n'utilise pas `.opencode/scripts/ppt-v1/SopraSteriaNext.layouts.json`; utilise `.opencode/scripts/ppt-v1/layouts.json` et evite tout layout dont la slide source est absente du template.
- Selectionne dynamiquement le layout de chaque slide avant d'ecrire `ppt-v1/deck.generated.md` et avant de convertir en JSON.
- Pour chaque slide de contenu, applique cet ordre de priorite strict : `section_header`, puis `two_columns_bullets`, puis `table_dynamic`, puis `classic_bullets`.
- Utilise `section_header` uniquement pour une slide de chapitre, transition ou separation courte : `#1`, `Partie 2`, `Conclusion`, `Synthese`, titre de section sans contenu detaille.
- Utilise `two_columns_bullets` des que le contenu contient exactement deux categories naturelles et suffisamment equilibrees, notamment avantages / inconvenients, pour / contre, forces / limites, opportunites / risques, problemes / reponses, avant / apres, metier / IT.
- Utilise `table_dynamic` uniquement si `two_columns_bullets` ne s'applique pas et si le contenu est explicitement tabulaire : tableau, matrice, grille, comparatif multi-criteres, ou colonnes homogenes.
- Utilise `classic_bullets` par defaut si aucune regle plus specifique ne s'applique, si le contenu est lineaire, ou si les categories sont desequilibrees.
- Si le layout choisi exige des placeholders specifiques, reformate le contenu pour les remplir correctement ; sinon choisis un layout plus simple compatible avec le contenu.
- Ne choisis jamais un layout dont le `sourceSlide` du catalogue est absent du template utilise.
- Maximum recommande : 90 caracteres pour un titre, 130 pour un sous-titre, 120 par bullet, 5 bullets par slide.
- Si une slide est trop dense, scinde-la ou reformule-la dans le Markdown avant generation.
- Apres generation, verifie toujours le rendu avec `Test-PptVisual.ps1` avant de conclure.
- Si le rapport signale un placeholder restant, un debordement ou une erreur visible, corrige le contenu dans `ppt-v1/deck.generated.md` et `ppt-v1/slides.generated.json`, regenere le PPT, puis relance la verification.
- Si le rapport signale `cover_title_too_long` ou si la cover semble visuellement mauvaise, corrige d'abord le contenu : raccourcis le titre de cover et deplace les details dans le sous-titre, puis regenere.
- Repete la boucle generation -> verification -> correction jusqu'a obtenir un rapport sans issue ou jusqu'a identifier un probleme de template non corrigeable par le contenu.
- Consulte les images exportees dans `ppt-v1/visual-check/` si le rapport signale un probleme, ou si le rendu semble douteux.

Workflow au premier tour :

1. Afficher `**Étape 1/8 - Analyse de la demande**`, puis identifier le mode : texte libre, fichier `.md/.txt`, image OCR, image jointe avec chemin fichier, ou reprise `.pptx`.
2. Afficher `**Étape 2/8 - Recherche et choix du template**`, puis identifier le template. S'il manque, rechercher `**/*.template.pptx` depuis la racine du projet, puis demander a l'utilisateur de choisir parmi les resultats trouves.
3. Apres choix du template, generer automatiquement le catalogue de layouts depuis ce template avec :

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\.opencode\scripts\ppt-v1\Get-PptLayouts.ps1 -TemplatePath "<template.pptx>" -OutputJson .\ppt-v1\layouts.detected.json
```

4. Utiliser ensuite `.\ppt-v1\layouts.detected.json` comme catalogue de layouts pour preparer le Markdown, convertir le JSON et generer le PowerPoint. Si la detection echoue, expliquer l'erreur et demander un template contenant des placeholders `{{...}}`, ou utiliser un catalogue statique uniquement comme fallback explicite.
5. Si la source est un `.pptx`, extraire d'abord son contenu, y compris le texte OCR des photos/images quand Tesseract est disponible, avec :

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\.opencode\scripts\ppt-v1\Extract-PptContent.ps1 -SourcePptx "<source.pptx>" -OutputMarkdown .\ppt-v1\deck.extracted.md
```

6. Si la source est une image `.png`, `.jpg`, `.jpeg`, `.bmp`, `.tif` ou `.tiff`, extraire d'abord son texte OCR avec :

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\.opencode\scripts\ppt-v1\Extract-ImageOcr.ps1 -SourceImage "<source-image>" -OutputMarkdown .\ppt-v1\image.extracted.md
```

7. Si la source est un `.pptx`, `.md`, `.txt`, `.png`, `.jpg`, `.jpeg`, `.bmp`, `.tif` ou `.tiff`, demander le niveau de reprise si l'utilisateur ne l'a pas deja fourni : `Quel niveau de reprise veux-tu pour le contenu source ? Réponds 1, 2 ou 3 : 1 exact, 2 synthétique, 3 très synthétique.`
8. Attendre la reponse avant de produire `ppt-v1/deck.generated.md`, sauf si le niveau est deja explicite dans la demande.
9. Afficher `**Étape 3/8 - Préparation du Markdown**`, puis choisir dynamiquement le layout de chaque slide parmi les layouts presents dans `ppt-v1/layouts.detected.json`, produire `ppt-v1/deck.generated.md` selon le niveau choisi, et indiquer le `Layout:` retenu dans chaque section de slide. Si le fichier n'existe pas, le creer completement au lieu de le patcher.
10. Afficher le Markdown complet dans la reponse.
11. Afficher `**Étape 4/8 - Validation utilisateur**`, puis demander : `Est-ce que ce Markdown est OK pour generer le PowerPoint ? Reponds "OK", "oui", "ça marche", "OK génère" ou indique les modifications.`
12. S'arreter. Ne pas generer le PowerPoint.

Workflow apres validation explicite :

1. Afficher `**Étape 5/8 - Génération du JSON**`, puis convertir `ppt-v1/deck.generated.md` en `ppt-v1/slides.generated.json` en respectant strictement le `Layout:` choisi pour chaque slide. Si le fichier JSON n'existe pas, le creer completement au lieu de le patcher.
2. Utiliser ce format JSON :

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
      "layout": "classic_bullets",
      "placeholders": {
        "{{TITLE}}": "...",
        "{{SUBTITLE}}": "...",
        "{{BODY}}": ["...", "...", "..."]
      }
    }
  ]
}
```

Exemple de slide a deux colonnes :

```json
{
  "layout": "two_columns_bullets",
  "placeholders": {
    "{{TITLE}}": "IA générative : bénéfices et vigilance",
    "{{SUBTITLE}}": "Accélérer les usages sans perdre le contrôle",
    "{{LEFT_TITLE}}": "Bénéfices",
    "{{LEFT_BODY}}": [
      "Produire rapidement des synthèses",
      "Accélérer la rédaction de contenus",
      "Explorer plusieurs pistes de solution"
    ],
    "{{RIGHT_TITLE}}": "Points de vigilance",
    "{{RIGHT_BODY}}": [
      "Vérifier les réponses produites",
      "Protéger les données sensibles",
      "Cadrer les usages autorisés"
    ]
  }
}
```

Exemple de slide tableau dynamique :

```json
{
  "layout": "table_dynamic",
  "placeholders": {
    "{{TITLE}}": "Prioriser les cas d’usage IA",
    "{{SUBTITLE}}": "Comparer valeur, risque et maturité",
    "{{CONCLUSION}}": "Prioriser les usages simples, fréquents et contrôlables."
  },
  "table": {
    "columns": ["Cas d’usage", "Valeur", "Risque", "Maturité"],
    "rows": [
      ["Synthèse documentaire", "Forte", "Faible", "Élevée"],
      ["Rédaction assistée", "Moyenne", "Faible", "Élevée"],
      ["Support utilisateur", "Forte", "Moyen", "Moyenne"]
    ]
  }
}
```

3. Afficher `**Étape 6/8 - Génération PowerPoint**`, puis choisir un nom de fichier final clair derive du titre de la presentation, sans `Generated` ou `généré`, a la racine du projet, et lancer depuis la racine du projet.
4. Lance toujours les scripts `.ps1` avec `powershell -NoProfile -ExecutionPolicy Bypass -File ...`. Ne lance jamais directement un script `.ps1` avec `& .\script.ps1`, car la politique d'execution Windows peut bloquer l'execution directe.

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\.opencode\scripts\ppt-v1\Generate-Ppt.ps1 -TemplatePath "<template.pptx>" -LayoutsJson .\ppt-v1\layouts.detected.json -InputJson .\ppt-v1\slides.generated.json -OutputPath .\<Nom_de_la_presentation>.pptx
```

5. Afficher `**Étape 7/8 - Contrôle visuel et corrections**`, puis lancer la verification visuelle :

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\.opencode\scripts\ppt-v1\Test-PptVisual.ps1 -PresentationPath .\<Nom_de_la_presentation>.pptx -ExportDirectory .\ppt-v1\visual-check -ReportPath .\ppt-v1\visual-check-report.json
```

6. Analyser la sortie : placeholders restants, alertes de rendu, pagination, densite et images exportees.
7. Si une alerte est corrigeable par le contenu, ajuste `ppt-v1/deck.generated.md`, regenere `ppt-v1/slides.generated.json`, puis relance generation et verification.
8. Ne conclus que lorsque `Issues=0`, ou explique clairement le probleme restant s'il vient du template.
9. Afficher `**Étape 8/8 - Finalisation et ouverture optionnelle**`, puis terminer avec le chemin du fichier genere et les alertes restantes eventuelles.
10. Demande ensuite : `Veux-tu que j'ouvre la présentation PowerPoint ?`
11. N'ouvre jamais le fichier automatiquement. Ouvre le fichier uniquement si l'utilisateur confirme avec `OK`, `oui`, `ça marche`, `ouvre`, `ouvre le fichier` ou une formule equivalente sans ambiguite.

Format attendu pour `ppt-v1/deck.generated.md` :

```markdown
# Titre de la presentation

Template: chemin/vers/template.template.pptx
Source: texte libre | source.md | source.pptx | image.png

## Slide 1 - Cover
Layout: cover

Titre:
Sous-titre:
Date:
Auteurs:

## Slide 2 - Titre court
Layout: classic_bullets

Titre:
Sous-titre:

Points:
- ...
- ...
- ...

## Slide 3 - Comparaison courte
Layout: two_columns_bullets

Titre:
Sous-titre:

Titre colonne gauche:
Colonne gauche:
- ...
- ...

Titre colonne droite:
Colonne droite:
- ...
- ...

## Slide 4 - Tableau court
Layout: table_dynamic

Titre:
Sous-titre:

Colonnes:
- Colonne 1
- Colonne 2
- Colonne 3

Lignes:
- Valeur 1 | Valeur 2 | Valeur 3
- Valeur 1 | Valeur 2 | Valeur 3

Conclusion:
Phrase de synthese courte.
```

Exemple : si l'utilisateur demande `genere 2 slides sur l'IA`, produire `Slide 1 - Cover`, `Slide 2 - ...` et `Slide 3 - ...`.
