# SuperGreffier — Guide pour les testeur·euses

Merci de tester l'outil 🙏 Ce guide te fait passer de zéro à un premier récit d'audience. Compte **15 minutes** la première fois. Aucune compétence technique requise, juste suivre les étapes dans l'ordre.

L'outil transforme des **notes d'audience brutes** (fichier Word) en un **récit structuré**, conforme à la charte du MSDE. Il ne réécrit pas : il met en forme, corrige la typo, développe les acronymes et signale les passages douteux.

---

## Étape 1 — Récupérer l'outil

1. Va sur **https://github.com/fczls/SuperGreffier**
2. Bouton vert **« Code »** → **« Download ZIP »**
3. Décompresse le ZIP (double-clic). Tu obtiens un dossier `SuperGreffier`.

---

## Étape 2 — Installer (une seule fois)

1. Ouvre le **Terminal** : `Cmd + Espace`, tape `Terminal`, Entrée.
2. Vérifie que Python est là :
   ```
   python3 --version
   ```
   Si ça affiche une erreur → installe Python depuis [python.org/downloads](https://www.python.org/downloads/), puis reprends.
3. Place-toi dans le sous-dossier `app`. Ce dossier `app` se trouve **à l'intérieur du dossier `SuperGreffier`** que tu as décompressé à l'étape 1. Le plus simple : tape `cd ` (le mot `cd` suivi d'un espace), puis **glisse-dépose le dossier `app`** depuis le Finder dans la fenêtre Terminal — son chemin s'écrit tout seul — puis Entrée.
4. Installe les dépendances :
   ```
   pip3 install -r requirements.txt
   ```

---

## Étape 3 — Lancer l'outil

Dans le Terminal, toujours dans le dossier `app` :

```
python3 -m streamlit run app.py
```

L'outil s'ouvre **tout seul dans ton navigateur** (sinon, ouvre `http://localhost:8501`).
Pour l'arrêter : reviens au Terminal et fais `Ctrl + C`.

---

## Étape 4 — La clé API

L'outil a besoin d'une clé pour parler à l'IA. **Fabien te la fournit en privé** (par message direct) pour la phase de test.

- Colle-la dans le champ **« Clé API Anthropic »** de l'écran **« Tout est prêt ? »**, juste avant de lancer le traitement.
- **Ne la partage pas, ne la mets nulle part par écrit en public.** Elle est liée au compte du MSDE.

---

## Étape 5 — Traiter un premier document

1. Prépare tes notes dans un fichier Word (`.docx`).
2. **Anonymise avant tout** : remplace les vrais noms par des pseudonymes. Les vrais noms ne doivent jamais être envoyés à l'IA (voir la politique RGPD du MSDE).
3. Charge le fichier dans l'outil, laisse-le proposer les développements d'acronymes inconnus, puis récupère le `.docx` structuré en sortie.
4. Relis dans Word : l'outil marque `[passage incertain — à vérifier]` là où il n'est pas sûr, **surligné en jaune**. Ces passages demandent ton œil.

---

## En cas de problème

| Symptôme | Solution |
|---|---|
| `command not found: streamlit` | Utilise bien `python3 -m streamlit run app.py` |
| `ModuleNotFoundError` | Les dépendances ne sont pas installées → refais l'étape 2.4 |
| `AuthenticationError` | La clé API est absente ou mal collée |
| L'outil ne s'ouvre pas | Ouvre `http://localhost:8501` à la main dans ton navigateur |

---

## Ton retour est précieux

Note ce qui coince : passages mal décodés, acronymes manqués, mise en forme à revoir, étapes d'installation pénibles. Remonte tout à Fabien — c'est exactement ce qu'on cherche à améliorer avec cette phase de test.
