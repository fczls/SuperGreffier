# SuperGreffier — Guide de lancement

## Prérequis (une seule fois)

**1. Python** — vérifie qu'il est installé :
```
python3 --version
```
Si la commande ne marche pas → télécharge Python sur [python.org](https://www.python.org/downloads/) et installe-le.

**2. Ouvre un Terminal**
- Mac : `Cmd + Espace` → tape `Terminal` → Entrée

**3. Va dans le dossier de l'outil**

Tape `cd ` (avec un espace), puis glisse-dépose le dossier `app` (celui du dossier `SuperGreffier` décompressé) depuis le Finder dans le Terminal, et appuie sur Entrée.

> **Version Obsidian (Fabien) :** le dossier est dans le coffre, donc `cd ~/Documents/Obsidian/msde/SuperGreffier/app`.

**4. Installe les dépendances** (une seule fois)
```bash
pip3 install -r requirements.txt
```
Si `pip3` ne marche pas non plus :
```bash
python3 -m pip install -r requirements.txt
```

---

## Lancer l'outil

Depuis le dossier `app` :
```bash
python3 -m streamlit run app.py
```

L'outil s'ouvre **automatiquement dans ton navigateur** sur `http://localhost:8501`.

Pour l'arrêter : retourner dans le Terminal et appuyer sur `Ctrl + C`.

---

## Clé API

L'outil a besoin d'une clé API Anthropic.
- **Testeur·euses :** Fabien vous la transmet en privé. Ne la partagez pas.
- **Sinon :** crée un compte sur [console.anthropic.com](https://console.anthropic.com) et génère une clé API.

Colle-la dans le champ « Clé API Anthropic » de l'écran « Tout est prêt ? », juste avant de lancer le traitement.

**Coût estimé : ~0,05 € par document de 10 pages.**

---

## En cas de problème

| Symptôme | Solution |
|---|---|
| `command not found: streamlit` | Utiliser `python3 -m streamlit run app.py` à la place |
| `ModuleNotFoundError` | Idem — dépendances pas installées |
| `AuthenticationError` | Vérifie ta clé API |
| L'outil ne s'ouvre pas | Ouvre manuellement `http://localhost:8501` dans ton navigateur |
