# /rec

Commande pour enregistrement ecran/audio avec transcription live (whisper.cpp), fusion des chunks, et gestion des sessions.

## Usage rapide

```text
/rec                   Afficher l'aide
/rec help              Afficher l'aide
/rec video              Enregistrement ecran (non decoupe)
/rec video+             Enregistrement ecran + transcription directe
/rec audio              Enregistrement audio (non decoupe)
/rec audio+             Enregistrement audio + transcription directe
/rec stop               Arreter l'enregistrement et finaliser
/rec status             Afficher la progression des transcriptions
/rec transcribe --fichier F  Transcrire un fichier existant
/rec devices            Lister les peripheriques audio
/rec list               Lister les enregistrements
/rec recover [N|all]    Recuperer des chunks orphelins
/rec clean [--all]      Nettoyer les chunks deja fusionnes
/rec install            Lancer l'installation des dependances
/rec install-status     Afficher la progression de l'installation
/rec install --window   Lancer l'installation dans une fenetre PowerShell
```

## Installation

La commande `/rec` est installee dans le projet cible via le depot `agent-ia-local-ssg`.

Depuis OpenCode Desktop, dans le projet cible, demander :

```text
Installer uniquement la commande projet /rec depuis https://github.com/helie-co/agent-ia-local-ssg/tree/main/rec. OpenCode Desktop est deja installe, ne pas l installer. Ne pas utiliser git.
```

Cette etape installe uniquement la commande projet. Elle ne telecharge pas les dependances.

Apres installation, redemarrer OpenCode Desktop depuis ce projet pour charger `/rec`.

### Dependances

Les dependances sont installees separement avec `/rec install` :
- **ffmpeg** via winget (package `Gyan.FFmpeg`)
- **whisper-cli** (telecharge depuis GitHub releases)
- **Modele whisper** `ggml-small.bin` (telecharge depuis HuggingFace)

Dans OpenCode, `/rec install` lance l'installation en arriere-plan et retourne immediatement le PID et le chemin du log. La progression est visible avec `/rec install-status`.

Pour voir la progression en direct dans une fenetre separee, utiliser `/rec install --window`.

Apres le redemarrage d'OpenCode Desktop, lancer :

```text
/rec install
/rec install-status
```

Ce mode peut aussi etre relance plus tard pour reparer une installation incomplete.

### Stereo Mix

Pour enregistrer l'audio systeme (son des applications, musique, visio), le peripherique **Stereo Mix** doit etre active.

Si `/rec install` detecte que Stereo Mix est absent, active-le avec le chemin suivant :

Modifier les sons système > Enregistrement > Stéréo Mix > Bouton droit > Activer > OK

Verifier ensuite avec :

```text
/rec devices
```

## Commandes

### video / video+

- `video` : enregistrement ecran + audio, fichier unique `.mp4`.
- `video+` : idem avec decoupage en chunks et transcription en direct.

Les deux modes enregistrent l'ecran (capture GDI) et l'audio (micro + systeme si Stereo Mix actif).

Options :
- `--titre T` ou `--title T` : ajoute un libelle au fichier de sortie.

### audio / audio+

- `audio` : enregistrement audio seul, fichier unique `.mp3`.
- `audio+` : idem avec decoupage en chunks et transcription en direct.

L'audio est enregistre en mono 16kHz MP3.

### stop

Arrete l'enregistrement actif (ou un enregistrement specifique).

```text
/rec stop               Arrete le plus recent
/rec stop 2             Arrete la session N (cf /rec status)
/rec stop all           Arrete toutes les sessions
```

Pour les sessions avec `+` (chunked), la fusion et la transcription se lancent en arriere-plan.

### status

Affiche la liste des enregistrements actifs avec la progression de la transcription.

```text
/rec status
```

### transcribe

Transcrit un fichier audio/video avec whisper.cpp.

```text
/rec transcribe --fichier C:\enregistrement.mp3
/rec transcribe --langue en   # pour une autre langue
```

Par defaut, transcrit le dernier enregistrement termine.

### devices

Liste les peripheriques audio disponibles avec leur etat :
- `[MIC]` micro detecte automatiquement
- `[SYSTEM]` Stereo Mix detecte
- `[OTHER]` autre peripherique

```text
/rec devices
```

### list

Liste tous les enregistrements termines dans le dossier `recordings/`.

```text
/rec list
```

### recover

Recupere des chunks orphelins (chunks dont le fichier final n'a pas ete genere).

```text
/rec recover                  # Liste les groupes orphelins
/rec recover 1                # Recupere le groupe N
/rec recover all              # Recupere tous les groupes
```

### clean

Supprime les chunks deja fusionnes dans un fichier final.

```text
/rec clean                    # Nettoie le dossier recordings/
/rec clean --all              # Nettoie tous les sous-dossiers recordings/
```

## Dossiers et fichiers

- `recordings/` : enregistrements termines (`.mp3`, `.mp4`, `.txt`).
- `recordings/logs/` : journaux ffmpeg.
- `recordings/.rec_state.json` : etat des sessions actives.

## Transcription

- La transcription utilise whisper.cpp avec le modele `ggml-small.bin`.
- Langue par defaut : `fr` (francais). Changer avec `--langue en` ou `--language en`.
- Pendant un enregistrement `+`, la transcription est lancee en direct sur chaque chunk.
- Apres `stop`, la fusion et transcription se finalisent en arriere-plan.

## Notes

- Les sessions actives sont persistees dans `.rec_state.json` : elles survivent a un redemarrage d'OpenCode.
- Les enregistrements non decoupes (`video`, `audio`) ne produisent pas de transcription automatique. Utiliser `rec transcribe` apres coup.
- L'enregistrement `video` produit des fichiers volumineux. Privilegier `video+` pour les longues sessions.
