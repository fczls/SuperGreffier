# SuperGreffier

> Point de départ de lecture du projet. À lire en premier pour reprendre le contexte.

## En une phrase

Outil (app Streamlit + API Claude) qui transforme les **notes d'audience brutes** des bénévoles du MSDE en **récit d'audience structuré**, conforme à la charte rédactionnelle, avec anonymisation systématique des données personnelles avant tout envoi à l'IA.

- **Porteur :** Fabien Cazals — **Statut :** PRD v0.3 en révision, prototype fonctionnel
- **Objectif clé :** passer de ~8 h à < 30 min pour produire un récit, sans reformulation (fidélité verbatim au texte source).

## Le principe non négociable

**L'outil met en forme et décode, il ne rédige pas.** Chaque phrase en entrée reste, dans sa teneur, la phrase de sortie. L'IA a le droit de : corriger orthographe/typo/ponctuation, compléter un mot tronqué si le sens est non ambigu, développer les acronymes du glossaire, attribuer les prises de parole, restructurer la mise en page. Elle n'a **pas** le droit de reformuler, ajouter, supprimer ou interpréter — les passages douteux sont marqués `[passage incertain — à vérifier]`.

## Contenu du dossier

| Fichier | Rôle |
|---|---|
| `README.md` | Ce point d'entrée |
| `PRD_Outil_Notes_Audience_MSDE.md` | **Document de référence** — problème, objectifs, périmètre, exigences P0/P1/P2, jalons, questions ouvertes |
| `Charte rédaction harmonisée - Traitement récits.md` | Règles d'écriture MSDE (typo, majuscules, écriture inclusive) — source de vérité rédactionnelle |
| `Politique_RGPD_Outil_IA_MSDE.md` | Politique de protection des données (v0.1, à valider par la direction) |
| `Mise en page Texte - Boîte à Outils.md` | Guide manuel de mise en page Word (livret A4, calques) — **hors périmètre v1** de l'outil |
| `app/app.py` | Code de l'application Streamlit (793 lignes) |
| `app/glossaire.json` | Bibliothèque d'acronymes partagée (persistante) |
| `app/requirements.txt` | Dépendances : streamlit, anthropic, python-docx |
| `app/LANCER.md` | Guide d'installation et de lancement pour bénévoles non-techniques |

## Architecture de l'app (`app/app.py`)

Pipeline principal, dans l'ordre :

1. **Upload** d'un `.docx` de notes brutes → `extract_text()`
2. **Anonymisation** locale via table de correspondance → `apply_anonymization()` (les vrais noms ne quittent jamais le poste)
3. **Détection des acronymes inconnus** → `detect_unknown_acronyms()` (résolution interactive par l'utilisateur)
4. **Appel API Claude** avec le prompt système + glossaire → `build_prompt()` / `call_claude()`
5. **Résolution des passages incertains** → `extract_uncertain_passages()` / `render_raw_notes_component()`
6. **Export `.docx` structuré** → `generate_docx()`

Le **prompt système** (`PROMPT_SYSTEME`, en haut du fichier) encode la règle de fidélité et les règles typographiques MSDE. Le glossaire de base est dans `GLOSSAIRE_BASE`, fusionné avec `glossaire.json` (modifiable via l'UI, sidebar).

## Lancer l'outil

```bash
cd ~/Documents/Obsidian/msde/SuperGreffier/app
pip3 install -r requirements.txt   # une seule fois
python3 -m streamlit run app.py
```

Ouvre `http://localhost:8501`. Nécessite une **clé API Anthropic** (console.anthropic.com), à coller dans la sidebar. Coût ~0,05 € / document de 10 pages. Détails et dépannage : `app/LANCER.md`.

## RGPD — l'essentiel

Aucune donnée personnelle réelle n'est envoyée à l'API. L'anonymisation (pseudonymes, table de correspondance locale) a lieu **avant** l'appel. Anthropic ne traite que du texte pseudonymisé. Base légale : intérêt légitime de l'association (art. 6.1.f) ; audiences publiques (art. 306 et 400 CPP). Voir `Politique_RGPD_Outil_IA_MSDE.md`.

## Hors périmètre v1 (à ne pas confondre avec des bugs)

Transcription audio · mise en page livret finale (reste manuelle, cf. *Boîte à Outils*) · notes de bas de page (signalées, pas rédigées) · multilingue · archivage des récits · collaboration temps réel.

## Où reprendre — état et prochaines étapes

- **Questions ouvertes bloquantes** (cf. fin du PRD) : récits de référence few-shot à intégrer au prompt ? Politique RGPD à valider par la direction avant déploiement large ?
- **Jalons v1** : Prototype → Bibliothèque & résolution → Test sur 5 cas réels → Consolidation marqueurs `[incertain]` → Déploiement. Total estimé 5-6 semaines à temps partiel.
- **Calibrage du prompt** : à tester sur les notes de niveau 2 (compressées) et 3 (quasi-sténo), où le respect de la fidélité est critique. Voir l'annexe « Niveaux de dégradation des notes » du PRD.

## Glossaire de contexte

- **MSDE** : l'association porteuse (observation d'audiences liées aux mobilisations climatiques et sociales).
- **Scribe** : bénévole qui prend les notes en audience.
- **Récit d'audience** : le livret final publié, destiné aux avocats partenaires et au réseau militant.
