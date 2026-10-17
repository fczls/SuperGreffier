"""
SuperGreffier — MSDE
Outil de traitement des notes d'audience
"""

import streamlit as st
import json
import re
import io
import time
import html as html_lib
from pathlib import Path

from docx import Document
from docx.shared import Pt, RGBColor
import anthropic

from energie import (
    estimate_energy, format_range, pc_minutes, format_duration,
    minutes_evitees, format_hm,
    SOURCES, DEMO_G_CO2, G_CO2_PC_PAR_HEURE,
)
from anonymisation import apply_anonymization, find_residual_names

# ─────────────────────────────────────────────────────────────────
# Configuration
# ─────────────────────────────────────────────────────────────────

GLOSSAIRE_PATH = Path(__file__).parent / "glossaire.json"

GLOSSAIRE_BASE = {
    "GAV": "garde à vue",
    "OP": "ordonnance pénale",
    "FDO": "forces de l'ordre",
    "DR": "Dernière Rénovation",
    "B2": "bulletin n°2 du casier judiciaire",
    "CC°": "Cour de cassation",
    "CEDH": "Cour européenne des droits de l'homme",
    "EN": "état de nécessité",
    "LE": "liberté d'expression",
    "RC": "réchauffement climatique",
    "AN": "Assemblée nationale",
    "LDB": "lanceur de balles de défense",
    "APP": "audience de présentation préalable",
    "ADEME": "Agence de l'environnement et de la maîtrise de l'énergie",
    "GIEC": "Groupe d'experts intergouvernemental sur l'évolution du climat",
    "CPP": "code de procédure pénale",
    "CP": "code pénal",
    "TJ": "tribunal judiciaire",
    "MP": "ministère public",
    "NV": "non-violent",
}

PROMPT_SYSTEME = """Tu es l'outil SuperGreffier du MSDE. Ta mission est de transformer des notes d'audience brutes en récit d'audience structuré, selon les règles rédactionnelles du MSDE.

## Règle fondamentale — fidélité absolue au texte source

Chaque mot, dans son ordre exact, doit être conservé tel quel. La phrase d'entrée est la phrase de sortie — tu n'es pas là pour améliorer le style, corriger la syntaxe ou rendre le texte plus élégant.

Tu es autorisé à :
- Corriger l'orthographe, les accents, la typographie et la ponctuation
- Compléter un mot manifestement tronqué ou phonétique lorsque le sens est NON AMBIGU et ÉVIDENT (ex : "trib judic" → "tribunal judiciaire")
- **Développer les acronymes selon le glossaire fourni** — c'est une obligation, pas une option. Chaque acronyme présent dans le glossaire DOIT être développé à sa première occurrence, même si cela modifie la forme du texte source. C'est la seule exception à la règle de fidélité verbatim.
- Attribuer une prise de parole à son locuteur lorsqu'il est identifiable
- Appliquer les règles typographiques MSDE ci-dessous

Tu N'ES PAS autorisé à :
- Changer l'ordre des mots d'une phrase
- Remplacer un mot par un synonyme
- Fusionner, scinder ou reformuler des phrases
- Ajouter des informations absentes des notes
- Supprimer des passages, même redondants ou répétitifs
- Réorganiser les paragraphes ou les prises de parole
- Marquer [passage incertain] pour un nom propre lisible et cohérent : un nom propre inconnu de toi n'est pas un passage incertain (c'est peut-être un pseudonyme ou un nom rare). Ne touche pas aux noms propres clairement écrits.

**Règle sur les noms propres corrompus :** si un mot est manifestement corrompu, phonétique ou illisible ET qu'il se résout vraisemblablement en un nom propre (nom de personne, nom de loi, nom d'organisation), NE JAMAIS deviner — toujours marquer [passage incertain — à vérifier]. Exemple : "duplomnb" pourrait être plusieurs noms différents → [passage incertain — à vérifier], pas "Dupont".

Exemple de ce qui est INTERDIT :
- Entrée : "Lyon, tribunal judiciaire. Celui-ci s'est refait une beauté."
- Sortie interdite : "Le tribunal judiciaire de Lyon vient d'être rénové."
- Sortie correcte : "Lyon, tribunal judiciaire. Celui-ci s'est refait une beauté."

## Structure de sortie

**Ordre chronologique intangible.** Ne déplace aucun paragraphe, aucune prise de parole. L'ordre des notes est l'ordre du récit. Les sections ci-dessous sont des étiquettes pour regrouper ce qui appartient naturellement ensemble — elles ne t'autorisent pas à réorganiser. Si un passage ne correspond pas à la section attendue, place-le là où il apparaît chronologiquement dans les notes.

Répartis le contenu dans ces sections sans en modifier l'ordre interne. Si une section manque d'informations, indique "(non renseigné dans les notes)".

---
**RÉSUMÉ D'AUDIENCE**
Tribunal :
Date :
Affaire :
Composition (présidente, procureur, avocates de la défense, greffière) :
Prévenu·es :
Faits reprochés :
Décision / date de délibéré :
---

# Avant l'audience
*[contexte, observations du scribe avant l'entrée de la présidente — narration en italique]*

# L'audience

## Ouverture et vérification des identités

## Déclarations des prévenu·es
*[une sous-section ### par personne]*

## Questions du tribunal et du parquet

## Réquisitions du parquet

## Plaidoiries de la défense

# La décision

## Règles typographiques MSDE

**Abréviations** : jamais Mme, M. — écrire en toutes lettres. "etc" → "*et cætera*" (italique).

**Institutions uniques** → majuscule : la Cour de cassation, l'Assemblée nationale, le Conseil d'État.
**Institutions locales** → minuscule : la cour d'appel de Paris, le tribunal judiciaire de Lyon.
**Institution = assemblée** → majuscule : le Tribunal doit statuer.

**Titres** → minuscule par choix MSDE : madame la procureure, monsieur le président.
Exception : le ministre de l'Intérieur (majuscule au domaine, minuscule à ministre).

**Madame/Monsieur** → minuscule à la 3e personne ("je ne vois pas madame Dupont") ; majuscule en adresse directe ("Comment allez-vous, Madame ?").

**Acronymes développés** : premier mot majuscule, reste minuscule, sauf noms propres.
Exemples : ONU → Organisation des nations unies ; CEDH → Cour européenne des droits de l'homme.

**Nombres** → en toutes lettres par défaut. Exceptions : dates (10 décembre 2026), grands nombres (2 956 €), données scientifiques, cas de lisibilité (18 km/h).

**Unités** → en lettres si nombre en lettres (six litres) ; en symbole si nombre en chiffres (800 €).

**Codes juridiques** → sans majuscule : le code pénal, le code de procédure pénale.

**État et République** → toujours majuscule.

## Écriture inclusive

**Citations directes** : aucune écriture inclusive — rapporter les paroles telles que prononcées.

**Narration** (parties rédigées, en italique) : écriture inclusive obligatoire.
- Privilégier tournures neutres : "ils et elles" plutôt que "iels"
- Noms épicènes quand possible pour éviter le point médian
- Point médian en dernier recours : prévenu·es, auteur·ices

## Format des prises de parole

**Nom ou titre en gras** suivi de deux points, puis le texte sur la même ligne.

Exemples attendus :
**La présidente :** Bonjour messieurs dames.
**Le procureur :** Je n'ai pas d'observations.
**L'avocate de la défense A.B. :** Madame la présidente, j'ai du mal à appréhender...
**Thomas Bertrand :** Vivre en société, c'est aussi se sentir responsable.

## Narration

Les observations et descriptions du scribe (ce qu'il voit, entend, ressent hors des prises de parole) doivent être en italique et rédigées avec écriture inclusive.

## Glossaire d'acronymes

Développe chaque acronyme à sa première occurrence sous la forme "terme (ACRONYME)". Les occurrences suivantes restent en acronyme.

{glossaire}
"""

# ─────────────────────────────────────────────────────────────────
# Fonctions utilitaires
# ─────────────────────────────────────────────────────────────────

def load_glossaire():
    if GLOSSAIRE_PATH.exists():
        return json.loads(GLOSSAIRE_PATH.read_text(encoding="utf-8"))
    save_glossaire(GLOSSAIRE_BASE)
    return GLOSSAIRE_BASE.copy()

def save_glossaire(glossaire):
    GLOSSAIRE_PATH.write_text(
        json.dumps(glossaire, ensure_ascii=False, indent=2),
        encoding="utf-8"
    )

def extract_text(docx_file):
    doc = Document(docx_file)
    return "\n".join(p.text for p in doc.paragraphs if p.text.strip())

def detect_unknown_acronyms(text, glossaire):
    """Détecte les séquences de 2 à 6 majuscules non présentes dans le glossaire."""
    EXCLUSIONS = {
        "IL", "ILS", "ELLE", "ELLES", "JE", "TU", "NOUS", "VOUS", "ON",
        "UN", "UNE", "DES", "LES", "AU", "AUX", "DU", "DE", "LA", "LE",
        "ET", "EN", "OUI", "NON", "MAIS", "CAR", "NI", "PAR", "SUR", "OR",
        "PAS", "EST", "SONT", "FAIT", "BIEN", "ALORS", "DONC", "VOUS",
        "DANS", "AVEC", "POUR", "QUE", "QUI", "TON", "MON", "SON",
    }
    found = set(re.findall(r'\b[A-Z]{2,6}°?\b', text))
    return sorted(a for a in found if a not in glossaire and a not in EXCLUSIONS)

def extract_uncertain_passages(text):
    """
    Extrait les passages incertains avec leur contexte (paragraphes voisins).
    Retourne une liste de dicts : line, context_before, context_after.
    """
    lines = text.split("\n")
    passages = []
    for i, line in enumerate(lines):
        if "[passage incertain" in line:
            before = "\n".join(l for l in lines[max(0, i - 3):i] if l.strip())
            after = "\n".join(l for l in lines[i + 1:min(len(lines), i + 4)] if l.strip())
            passages.append({
                "line_idx": i,
                "line": line,
                "context_before": before,
                "context_after": after,
            })
    return passages


def apply_uncertain_corrections(text, corrections_map, n_passages):
    """
    Remplace les occurrences de [passage incertain — à vérifier] par les
    corrections saisies (mapping correction_0…N → texte).
    Les passages sans correction conservent le marqueur.
    """
    marker = "[passage incertain — à vérifier]"
    parts = text.split(marker)
    result = parts[0]
    for i in range(len(parts) - 1):
        correction = (corrections_map.get(f"correction_{i}") or "").strip()
        result += correction if correction else marker
        result += parts[i + 1]
    return result


def build_prompt(glossaire):
    lines = [f"- {k} : {v}" for k, v in sorted(glossaire.items())]
    return PROMPT_SYSTEME.format(glossaire="\n".join(lines))

def call_claude(text, system_prompt, api_key):
    """Retourne (texte, (tokens_entrée, tokens_sortie))."""
    client = anthropic.Anthropic(api_key=api_key)
    with client.messages.stream(
        model="claude-sonnet-4-6",
        max_tokens=32000,
        system=system_prompt,
        messages=[{"role": "user", "content": text}],
    ) as stream:
        message = stream.get_final_message()
    result = "".join(b.text for b in message.content if b.type == "text")
    return result, (message.usage.input_tokens, message.usage.output_tokens)

def add_inline_markdown(paragraph, text):
    """
    Parse le markdown inline (**bold**, *italic*, combinaisons) et ajoute
    les runs correspondants au paragraphe docx fourni.
    """
    # Tokenise en segments : **...**  *...*  ou texte brut
    pattern = re.compile(r'(\*\*.*?\*\*|\*[^*]+?\*)')
    parts = pattern.split(text)
    for part in parts:
        if part.startswith("**") and part.endswith("**"):
            run = paragraph.add_run(part[2:-2])
            run.bold = True
        elif part.startswith("*") and part.endswith("*"):
            run = paragraph.add_run(part[1:-1])
            run.italic = True
        else:
            paragraph.add_run(part)


def generate_docx(text):
    """Convertit le markdown simple produit par Claude en .docx stylé."""
    doc = Document()

    for line in text.split("\n"):
        s = line.strip()
        if not s:
            doc.add_paragraph()
            continue

        if s.startswith("# "):
            doc.add_heading(s[2:], level=1)
        elif s.startswith("## "):
            doc.add_heading(s[3:], level=2)
        elif s.startswith("### "):
            doc.add_heading(s[4:], level=3)
        elif s == "---":
            doc.add_paragraph()
        elif re.match(r"^\*[^*].*[^*]\*$", s):
            # Narration en italique : *texte entier*
            p = doc.add_paragraph()
            p.add_run(s[1:-1]).italic = True
        elif re.match(r"^\*\*.+?\*\*\s*:", s):
            # Prise de parole : **Nom :** texte (potentiellement avec markdown inline)
            m = re.match(r"^\*\*(.+?)\*\*\s*:\s*(.*)", s)
            if m:
                p = doc.add_paragraph()
                p.add_run(m.group(1) + " : ").bold = True
                if m.group(2):
                    add_inline_markdown(p, m.group(2))
            else:
                p = doc.add_paragraph()
                add_inline_markdown(p, s)
        elif "[passage incertain" in s:
            # Passage incertain → texte souligné en orange
            p = doc.add_paragraph()
            run = p.add_run(s)
            run.underline = True
            run.font.color.rgb = RGBColor(0xC0, 0x50, 0x00)
        else:
            # Paragraphe général — parse le markdown inline
            p = doc.add_paragraph()
            add_inline_markdown(p, s)

    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf

def parse_anon_table(raw):
    """Parse le tableau d'anonymisation saisi par l'utilisateur."""
    table = {}
    for line in raw.strip().split("\n"):
        for sep in ["→", "->"]:
            if sep in line:
                parts = line.split(sep, 1)
                table[parts[0].strip()] = parts[1].strip()
                break
    return table


def build_demo_result(raw_text):
    """
    Récit fictif pour le mode démo : permet de parcourir tout le flux
    (passages incertains, notes suggérées, export) sans appel API.
    """
    lines = [l.strip() for l in raw_text.split("\n") if l.strip()]
    extrait = " ".join(lines[:2])[:200] if lines else "(notes vides)"
    return f"""---
**RÉSUMÉ D'AUDIENCE**
Tribunal : tribunal judiciaire de Lyon
Date : 8 juin 2026
Affaire : blocage de l'autoroute A6 — entrave à la circulation
Composition (présidente, procureur, avocates de la défense, greffière) : madame la présidente, madame la procureure, avocates A.B. et C.D., greffière
Prévenu·es : Thomas Bertrand, Sophie Duval
Faits reprochés : entrave à la circulation lors d'une action militante
Décision / date de délibéré : délibéré au 10 septembre 2026
---

# Avant l'audience
*La salle est presque pleine, une dizaine de personnes soutiennent les prévenu·es. Début des notes : {extrait}*

# L'audience

## Ouverture et vérification des identités

**La présidente :** Bonjour messieurs dames. Nous allons vérifier les identités.

**Thomas Bertrand :** Je suis ingénieur informaticien, j'ai toujours été écolo.

## Déclarations des prévenu·es

### Thomas Bertrand

**Thomas Bertrand :** Il m'a fallu beaucoup de temps pour me rendre compte qu'il y avait des moyens pour agir et j'ai alors [passage incertain — à vérifier] engagé dans Dernière Rénovation (DR).

*Il évoque ensuite la fresque du climat, ateliers [passage incertain — à vérifier] où on informe le grand public sur les causes du réchauffement climatique (RC).*

### Sophie Duval

**Sophie Duval :** Nous avions une banderole et une collègue a été interviewée par une journaliste pour évoquer l'action.

## Questions du tribunal et du parquet

**La présidente :** Pas d'autres questions ?

## Réquisitions du parquet

**La procureure :** Je requiers deux cents euros d'amende. Le rapport du GIEC [note suggérée] ne justifie pas l'entrave.

## Plaidoiries de la défense

**L'avocate de la défense A.B. :** Madame la présidente, j'invoque l'état de nécessité (EN) au titre de l'article [passage incertain — à vérifier] du code pénal.

# La décision

*Le tribunal renvoie la décision au délibéré. Ils et elles se donnent rendez-vous le 10 septembre.*
"""

# ─────────────────────────────────────────────────────────────────
# UI — thème visuel (inspiration « Relaxes pour le vivant » / MSDE)
# ─────────────────────────────────────────────────────────────────

CUSTOM_CSS = """
<style>
@import url('https://fonts.googleapis.com/css2?family=Fraunces:opsz,wght@9..144,500;9..144,600;9..144,700&family=Poppins:wght@400;500;600;700&display=swap');

/* ── Base ─────────────────────────────────────────── */
html, body, .stApp, button, input, textarea, select,
[class*="st-"]:not([data-testid="stIconMaterial"]):not(.material-symbols-rounded):not([data-testid="stExpanderToggleIcon"]) {
  font-family: 'Poppins', -apple-system, 'Segoe UI', sans-serif;
}
.stApp { background: #FBF8EF; }
h1, h2, h3 {
  font-family: 'Fraunces', Georgia, serif !important;
  color: #14201D !important;
  letter-spacing: -0.01em;
}
.block-container { padding-top: 2.4rem; max-width: 1080px; }
#MainMenu, footer { visibility: hidden; }
header[data-testid="stHeader"] { background: transparent; }

/* Restaurer la police des icônes Material (écrasée par la règle Poppins) */
[data-testid="stIconMaterial"], [data-testid="stExpanderToggleIcon"],
.material-symbols-rounded, .material-symbols-outlined, .material-icons {
  font-family: 'Material Symbols Rounded' !important;
  font-weight: normal !important;
}

/* Masquer les mentions « Press Enter to apply » (déroutantes pour les novices) */
[data-testid="InputInstructions"] { display: none !important; }

/* Clé API : champ texte simple (pas de type password → pas de gestionnaire
   de mots de passe), masqué visuellement par CSS */
.st-key-api_key_widget input,
input[placeholder="sk-ant-…"] {
  -webkit-text-security: disc;
}

/* ── En-tête ──────────────────────────────────────── */
.sg-kicker {
  display: inline-block; background: #14201D; color: #FFFFFF;
  padding: 3px 12px; font-size: 0.72rem; font-weight: 600;
  letter-spacing: 0.06em; text-transform: uppercase;
  transform: rotate(-0.6deg); margin-bottom: 10px;
}
.sg-title {
  font-family: 'Fraunces', Georgia, serif; font-size: 2.4rem; font-weight: 600;
  color: #14201D; margin: 0 0 4px 0; line-height: 1.1;
}
.sg-sub { color: #5C6660; font-size: 0.95rem; margin: 0 0 6px 0; max-width: 640px; }

/* ── Stepper : fil d'ariane cliquable ─────────────── */
div[class*="st-key-stepbtn_"] button {
  border-radius: 999px !important; font-size: 0.85rem !important;
  font-weight: 500 !important; padding: 4px 10px !important;
  border: 1px solid #E3DECB !important; background: #FFFFFF !important;
  color: #8A857A !important; box-shadow: none !important;
}
div[class*="st-key-stepbtn_"] button:disabled { opacity: 1 !important; cursor: default !important; }
div[class*="st-key-stepbtn_"][class*="_todo"] button { opacity: 0.55 !important; }
div[class*="st-key-stepbtn_"][class*="_done"] button {
  background: #DFF2E7 !important; color: #0E5B44 !important; border-color: #BFE3D0 !important;
  cursor: pointer !important;
}
div[class*="st-key-stepbtn_"][class*="_done"] button:hover:enabled {
  background: #CDEBD9 !important; transform: none;
}
div[class*="st-key-stepbtn_"][class*="_active"] button {
  background: #0C3A34 !important; color: #FFFFFF !important; border-color: #0C3A34 !important;
}

/* Bouton énergie (popover) : discret, aligné à droite */
.st-key-energie_pop button {
  border: none !important; background: transparent !important;
  color: #5C6660 !important; box-shadow: none !important;
  font-size: 0.85rem !important; float: right;
}
.st-key-energie_pop button p { white-space: nowrap !important; }
.st-key-energie_pop button:hover { color: #14201D !important; }
[data-testid="stPopoverBody"] { min-width: 420px; }

/* ── Stats (étape 4) : icône / valeur / libellé ───── */
.sg-stats { display: flex; gap: 40px; margin: 14px 0 6px 0; flex-wrap: wrap; }
.sg-stat {
  display: flex; flex-direction: column; align-items: center;
  text-align: center; gap: 7px; min-width: 110px;
}
.sg-stat .ic {
  width: 46px; height: 46px; border-radius: 50%;
  display: flex; align-items: center; justify-content: center;
  font-size: 1.3rem; font-weight: 700;
}
.sg-stat .ic.ok   { background: #DFF2E7; color: #0E5B44; }
.sg-stat .ic.warn { background: #FDE4D5; color: #9A4B00; }
.sg-stat .ic.info { background: #E7E4FA; color: #4A42B8; }
.sg-stat .ic.neutral { background: #F1EDE0; color: #5C6660; }
.sg-stat .v {
  font-family: 'Fraunces', Georgia, serif; font-size: 1.45rem;
  font-weight: 600; color: #14201D; line-height: 1;
}
.sg-stat .l { font-size: 0.85rem; color: #5C6660; max-width: 170px; line-height: 1.35; }

/* ── Chips ────────────────────────────────────────── */
.sg-chip {
  display: inline-flex; align-items: center; gap: 6px;
  padding: 4px 14px; border-radius: 999px;
  font-size: 0.82rem; font-weight: 500; margin: 2px 6px 2px 0;
}
.sg-chip.ok   { background: #DFF2E7; color: #0E5B44; border: 1px solid #BFE3D0; }
.sg-chip.warn { background: #FDE4D5; color: #9A4B00; border: 1px solid #F5CBAA; }
.sg-chip.info { background: #E7E4FA; color: #4A42B8; border: 1px solid #D2CDF2; }
.sg-chip.neutral { background: #F1EDE0; color: #5C6660; border: 1px solid #E3DECB; }

/* ── Récapitulatif (étape 3) ──────────────────────── */
.sg-recap { display: flex; flex-direction: column; gap: 12px; padding: 4px 0; }
.sg-recap .row { display: flex; align-items: baseline; gap: 12px; }
.sg-recap .ic {
  width: 22px; height: 22px; border-radius: 50%; flex-shrink: 0;
  display: inline-flex; align-items: center; justify-content: center;
  font-size: 0.72rem; font-weight: 700; transform: translateY(4px);
}
.sg-recap .ic.ok { background: #0E5B44; color: #FFFFFF; }
.sg-recap .ic.ko { background: #FDE4D5; color: #9A4B00; }
.sg-recap .lb { font-weight: 600; color: #14201D; min-width: 140px; }
.sg-recap .dt { color: #5C6660; font-size: 0.9rem; }

/* ── Boutons « sticker » ──────────────────────────── */
button[kind="primary"], [data-testid="stBaseButton-primary"] {
  background: #FFC933 !important; color: #14201D !important;
  border: 1.5px solid #14201D !important; border-radius: 6px !important;
  font-weight: 600 !important; box-shadow: 3px 3px 0 #14201D !important;
  transition: all .12s ease;
}
button[kind="primary"]:hover:enabled, [data-testid="stBaseButton-primary"]:hover:enabled {
  background: #FFD75E !important; transform: translate(-1px,-1px);
  box-shadow: 4px 4px 0 #14201D !important;
}
button[kind="primary"]:disabled, [data-testid="stBaseButton-primary"]:disabled {
  background: #F1EDE0 !important; color: #A8A395 !important;
  border-color: #C9C4B4 !important; box-shadow: none !important;
}
button[kind="secondary"], [data-testid="stBaseButton-secondary"] {
  background: #FFFFFF !important; color: #14201D !important;
  border: 1px solid #14201D !important; border-radius: 6px !important;
  font-weight: 500 !important;
}
[data-testid="stDownloadButton"] button {
  background: #FFC933 !important; color: #14201D !important;
  border: 1.5px solid #14201D !important; border-radius: 6px !important;
  font-weight: 600 !important; box-shadow: 3px 3px 0 #14201D !important;
}

/* ── Conteneurs / cartes ──────────────────────────── */
[data-testid="stMain"] div[data-testid="stVerticalBlockBorderWrapper"] {
  background: #FFFFFF; border: 1px solid #E7E2D0; border-radius: 12px;
}
[data-testid="stMain"] [data-testid="stExpander"] {
  background: #FFFFFF; border: 1px solid #E7E2D0; border-radius: 10px;
}
[data-testid="stMain"] [data-testid="stExpander"] details { border: none; }
[data-testid="stMetric"] {
  background: #FFFFFF; border: 1px solid #E7E2D0; border-radius: 10px;
  padding: 12px 16px;
}
[data-testid="stFileUploaderDropzone"] {
  background: #FFFFFF; border: 1.5px dashed #0C3A34; border-radius: 10px;
}

/* ── Écran de traitement ──────────────────────────── */
.sg-processing {
  background: #0C3A34; border-radius: 12px; padding: 56px 32px;
  text-align: center; color: #F4F1E6; margin: 8px 0;
}
.sg-processing h3 {
  color: #FFFFFF !important; font-family: 'Fraunces', Georgia, serif;
  margin: 20px 0 8px 0; font-size: 1.5rem;
}
.sg-processing p { color: #9DBBB0; font-size: 0.9rem; margin: 0; line-height: 1.6; }
.sg-dots { display: flex; gap: 10px; justify-content: center; }
.sg-dots span {
  width: 14px; height: 14px; border-radius: 50%; background: #FFC933;
  animation: sg-bounce 1.2s infinite ease-in-out;
}
.sg-dots span:nth-child(2) { animation-delay: 0.15s; background: #DFF2E7; }
.sg-dots span:nth-child(3) { animation-delay: 0.3s; }
@keyframes sg-bounce {
  0%, 80%, 100% { transform: scale(0.6); opacity: 0.5; }
  40% { transform: scale(1); opacity: 1; }
}

/* ── Comparaison énergie (popover) ───────────────── */
.sg-compare { background: #FBF8EF; border: 1px solid #E7E2D0; border-radius: 10px; padding: 14px 16px; }
.sg-compare .row { display: flex; justify-content: space-between; gap: 12px; padding: 4px 0; font-size: 0.88rem; }
.sg-compare .lb { color: #5C6660; }
.sg-compare .v { color: #14201D; font-weight: 500; text-align: right; }
.sg-compare .gain { margin-top: 10px; padding-top: 12px; border-top: 1px dashed #C9C4B4; text-align: center; }
.sg-compare .big { display: block; font-family: 'Fraunces', Georgia, serif; font-size: 2rem; font-weight: 600; color: #0E5B44; line-height: 1.1; }
.sg-compare .txt { display: block; font-size: 0.82rem; color: #5C6660; margin-top: 4px; }
.sg-small, .sg-small li { font-size: 0.74rem !important; color: #8A857A !important; line-height: 1.5; }
.sg-small a { color: #8A857A !important; }

/* ── Divers ───────────────────────────────────────── */
[data-testid="stMain"] textarea { background: #FDFCF7; border-radius: 8px; }
.sg-passage-nav {
  text-align: right; font-weight: 600; padding-top: 14px;
  color: #5C6660; font-size: 0.9rem;
}
</style>
"""

STEPS = ["Charger", "Préparer", "Lancer", "Relire & exporter"]


def render_stepper(current, nav_enabled=True):
    """Fil d'ariane cliquable : les étapes passées permettent de revenir en arrière."""
    cols = st.columns([1.15, 1.15, 1.05, 1.75, 1.0, 1.15])
    for i, name in enumerate(STEPS, 1):
        state = "done" if i < current else ("active" if i == current else "todo")
        label = ("✓ " if state == "done" else f"{i} · ") + name
        # L'étape 1 se rejoue via « Remplacer le fichier », pas via le fil d'ariane
        clickable = nav_enabled and state == "done" and i != 1
        with cols[i - 1]:
            if st.button(
                label,
                key=f"stepbtn_{i}_{state}",
                disabled=not clickable,
                use_container_width=True,
                help="Revenir à cette étape" if clickable else None,
            ):
                goto(i)
    with cols[5]:
        with st.container(key="energie_pop"):
            with st.popover("⚡ Énergie"):
                render_energy_panel()


def render_energy_panel():
    """Estimation CO₂e du dernier traitement + méthode de calcul."""
    usage = st.session_state.usage_tokens
    demo = st.session_state.demo_run and st.session_state.result_text
    g_range = None
    if demo:
        g_range = DEMO_G_CO2
    elif usage:
        g_range = estimate_energy(*usage)["g_co2"]

    if g_range:
        st.markdown(chip(format_range(*g_range, "g CO₂e"), "neutral"), unsafe_allow_html=True)
        if demo:
            st.caption("Valeurs de démonstration — aucun appel API n'a été fait.")
        else:
            st.caption(
                f"{usage[0]:,} tokens envoyés · {usage[1]:,} tokens générés".replace(",", " ")
            )
        m_bas, m_haut = pc_minutes(*g_range)
        st.number_input(
            "Temps de mise au propre sans l'outil (heures)",
            min_value=0.5, max_value=40.0, step=0.5, value=4.0, key="manual_hours",
        )
        h = st.session_state.manual_hours
        evitees = minutes_evitees(h, g_range[1])
        st.markdown(
            '<div class="sg-compare">'
            f'<div class="row"><span class="lb">Sans l\'outil</span>'
            f'<span class="v">{format_hm(h * 60)} d\'ordinateur allumé</span></div>'
            f'<div class="row"><span class="lb">Avec l\'outil</span>'
            f'<span class="v">au plus {format_duration(m_haut)} d\'équivalent '
            f'({format_range(*g_range, "g CO₂e")})</span></div>'
            f'<div class="gain"><span class="big">≥ {format_hm(evitees)}</span>'
            f'<span class="txt">d\'ordinateur allumé évitées en équivalent énergétique, '
            f'dans le pire cas</span></div>'
            '</div>',
            unsafe_allow_html=True,
        )
        st.markdown(
            f'<p class="sg-small">Base : un ordinateur allumé émet '
            f'{G_CO2_PC_PAR_HEURE["bas"]} à {G_CO2_PC_PAR_HEURE["haut"]} g CO₂e par heure. '
            "Le temps de relecture n'est pas compté.</p>",
            unsafe_allow_html=True,
        )
    else:
        st.info("L'estimation s'affiche après un traitement (API ou mode démo).")

    with st.expander("Méthode et sources"):
        st.markdown(
            '<p class="sg-small">Ordre de grandeur, pas une mesure : Anthropic ne publie '
            "aucune donnée d'énergie par requête. La fourchette part de chiffres publiés pour "
            "d'autres modèles (≈ 0,5 à 6 mWh par token généré, un token envoyé pesant dix fois "
            "moins) et d'une intensité du réseau électrique entre 55 et 400 g CO₂e/kWh selon le "
            "lieu du datacenter (inconnu). Fabrication du matériel et eau non comprises. "
            "La consommation d'un ordinateur allumé vient d'Ekwateur, qui ne cite pas sa "
            "source.</p>"
            '<ul class="sg-small">'
            + "".join(f'<li><a href="{u}" target="_blank">{t}</a></li>' for t, u in SOURCES)
            + "</ul>",
            unsafe_allow_html=True,
        )


def chip(label, kind="ok"):
    return f'<span class="sg-chip {kind}">{label}</span>'


def goto(step):
    st.session_state.step = step
    st.rerun()


# ─────────────────────────────────────────────────────────────────
# Interface
# ─────────────────────────────────────────────────────────────────

st.set_page_config(page_title="SuperGreffier — MSDE", page_icon="⚖️", layout="wide")
st.markdown(CUSTOM_CSS, unsafe_allow_html=True)

# Session state
for key, default in [
    ("glossaire", None),
    ("raw_text", ""),
    ("result_text", ""),
    ("uploaded_filename", ""),
    ("uncertain_passages", []),
    ("uncertain_idx", 0),
    ("step", 1),
    ("anon_raw", ""),
    ("corrections", {}),
    ("flash", ""),
    ("processing", False),
    ("process_error", ""),
    ("api_key_store", ""),
    ("demo_run", False),
    ("review_done", False),
    ("usage_tokens", None),
]:
    if key not in st.session_state:
        st.session_state[key] = default

if st.session_state.glossaire is None:
    st.session_state.glossaire = load_glossaire()

# La clé API vit dans un état persistant (le champ de saisie est à l'étape 3)
api_key = st.session_state.api_key_store.strip()

# ── En-tête + stepper ─────────────────────────────────────────────
st.markdown(
    '<div class="sg-kicker">MSDE · Outil de mise en forme</div>'
    '<h1 class="sg-title">SuperGreffier</h1>'
    '<p class="sg-sub">Des notes d\'audience brutes au récit structuré — fidèle au texte, '
    'anonymisé avant tout envoi à l\'IA.</p>',
    unsafe_allow_html=True,
)

step = st.session_state.step
render_stepper(step, nav_enabled=not st.session_state.processing)

# Message différé (affiché après un rerun)
if st.session_state.flash:
    st.success(st.session_state.flash)
    st.session_state.flash = ""

# ══════════════════════════════════════════════════════════════════
# Étape 1 — Charger
# ══════════════════════════════════════════════════════════════════
if step == 1:
    with st.container(border=True):
        st.markdown("### Vos notes d'audience")
        st.caption("Dépose le fichier `.docx` de notes brutes — la préparation démarre automatiquement.")
        uploaded = st.file_uploader(
            "Fichier .docx de notes brutes", type=["docx"], label_visibility="collapsed"
        )

        if uploaded:
            st.session_state.raw_text = extract_text(uploaded)
            st.session_state.uploaded_filename = uploaded.name
            st.session_state.result_text = ""
            st.session_state.uncertain_passages = []
            st.session_state.uncertain_idx = 0
            st.session_state.corrections = {}
            st.session_state.review_done = False
            goto(2)

# ══════════════════════════════════════════════════════════════════
# Étape 2 — Préparer (acronymes + anonymisation)
# ══════════════════════════════════════════════════════════════════
elif step == 2:

    # ── Fichier chargé : rappel + actions sur une seule ligne ─────
    n_lines = len(st.session_state.raw_text.split("\n"))
    f_col, a_col, r_col = st.columns([3.4, 1.0, 1.6])
    with f_col:
        st.markdown(
            chip(f"✓ {st.session_state.uploaded_filename} · {n_lines} lignes", "ok"),
            unsafe_allow_html=True,
        )
    with a_col:
        with st.popover("Aperçu", use_container_width=True):
            preview = st.session_state.raw_text
            st.text(preview[:3000] + ("…" if len(preview) > 3000 else ""))
    with r_col:
        if st.button("Remplacer le fichier", use_container_width=True):
            st.session_state.raw_text = ""
            st.session_state.uploaded_filename = ""
            st.session_state.result_text = ""
            goto(1)

    # ── Acronymes ─────────────────────────────────────────────────
    with st.container(border=True):
        st.markdown("### Acronymes")
        unknown = detect_unknown_acronyms(
            st.session_state.raw_text, st.session_state.glossaire
        )
        if not unknown:
            g_col1, g_col2 = st.columns([3.4, 2.6])
            with g_col1:
                st.markdown(
                    chip("✓ Tous les acronymes détectés sont dans le glossaire", "ok"),
                    unsafe_allow_html=True,
                )
        else:
            st.caption(
                f"{len(unknown)} acronyme(s) non reconnu(s) — renseigne leur signification "
                "(ou laisse vide pour les ignorer), puis enregistre-les."
            )
            cols = st.columns(3)
            for i, acro in enumerate(unknown):
                cols[i % 3].text_input(
                    acro, key=f"acro_{acro}", placeholder="Signification"
                )

            g_col1, g_col2 = st.columns([1.6, 4.4])
            with g_col1:
                if st.button("Ajouter au glossaire", use_container_width=True):
                    clicked_entries = {}
                    for acro in unknown:
                        val = (st.session_state.get(f"acro_{acro}") or "").strip()
                        if val:
                            clicked_entries[acro] = val
                    if clicked_entries:
                        st.session_state.glossaire.update(clicked_entries)
                        save_glossaire(st.session_state.glossaire)
                        st.session_state.flash = (
                            f"{len(clicked_entries)} entrée(s) ajoutée(s) au glossaire."
                        )
                        st.rerun()
                    else:
                        st.warning("Renseigne au moins une signification avant d'ajouter.")

        with g_col2:
            with st.popover(
                f"Glossaire complet · {len(st.session_state.glossaire)} entrées"
            ):
                raw_glossaire = "\n".join(
                    f"{k} : {v}" for k, v in sorted(st.session_state.glossaire.items())
                )
                edited = st.text_area(
                    "Format : SIGLE : terme (un par ligne)", value=raw_glossaire, height=280
                )
                if st.button("Sauvegarder le glossaire"):
                    new_g = {}
                    for line in edited.strip().split("\n"):
                        if ":" in line:
                            k, v = line.split(":", 1)
                            new_g[k.strip()] = v.strip()
                    st.session_state.glossaire = new_g
                    save_glossaire(new_g)
                    st.session_state.flash = "Glossaire sauvegardé."
                    st.rerun()
                st.caption("Partagé entre tous les documents, conservé d'une session à l'autre.")

    # ── Anonymisation ─────────────────────────────────────────────
    with st.container(border=True):
        st.markdown("### Anonymisation")
        st.caption(
            "Un remplacement par ligne, au format `Vrai Nom → Pseudonyme` — initiales "
            "pour les avocat·es (`L.K. → A.B.`), prénom et nom fictifs à consonance proche "
            "pour les prévenu·es. Les vrais noms ne quittent jamais ton poste."
        )

        if "anon_widget" not in st.session_state:
            st.session_state.anon_widget = st.session_state.anon_raw
        st.text_area(
            "Noms à anonymiser",
            key="anon_widget",
            placeholder="Malo Royen → Thomas Bertrand\nJessica Remaillon → Sophie Duval\nL.K. → A.B.",
            height=140,
            label_visibility="collapsed",
        )
        st.session_state.anon_raw = st.session_state.anon_widget

        anon_table = parse_anon_table(st.session_state.anon_raw)
        if anon_table:
            st.markdown(
                chip(f"✓ {len(anon_table)} remplacement(s) configuré(s)", "ok"),
                unsafe_allow_html=True,
            )
        else:
            st.caption("Tu peux laisser vide si les notes ne mentionnent aucun nom réel.")

    _, col_next = st.columns([5, 1.4])
    with col_next:
        if st.button("Continuer →", type="primary", use_container_width=True):
            goto(3)

# ══════════════════════════════════════════════════════════════════
# Étape 3 — Vérifier & lancer
# ══════════════════════════════════════════════════════════════════
elif step == 3:
    anon_table = parse_anon_table(st.session_state.anon_raw)
    n_anon = len(anon_table)
    n_gloss = len(st.session_state.glossaire)
    n_lines = len(st.session_state.raw_text.split("\n"))

    if st.session_state.processing:
        # ── Écran de traitement : rien d'autre à voir ni à faire ──
        st.markdown(
            '<div class="sg-processing">'
            '<div class="sg-dots"><span></span><span></span><span></span></div>'
            '<h3>Traitement en cours…</h3>'
            '<p>Mise en forme fidèle du récit, sans reformulation.<br>'
            '30 à 90 secondes — ne ferme pas cet onglet.</p>'
            '</div>',
            unsafe_allow_html=True,
        )

        # Appel bloquant : l'animation ci-dessus reste affichée pendant ce temps
        st.session_state.uncertain_passages = []
        st.session_state.uncertain_idx = 0
        st.session_state.corrections = {}

        error = None
        if st.session_state.demo_run:
            time.sleep(2.5)  # laisser le temps de voir l'écran de traitement
            st.session_state.result_text = build_demo_result(st.session_state.raw_text)
            st.session_state.usage_tokens = None
        else:
            anonymized = apply_anonymization(st.session_state.raw_text, anon_table)
            residual = find_residual_names(anonymized, anon_table)
            system_prompt = build_prompt(st.session_state.glossaire)
            if residual:
                error = (
                    "Envoi bloqué : ces noms réels sont encore présents après anonymisation : "
                    + ", ".join(residual)
                    + ". Ajoute-les explicitement à la table d'anonymisation (étape 2)."
                )
            try:
                if error:
                    raise RuntimeError(error)
                st.session_state.result_text, st.session_state.usage_tokens = call_claude(
                    anonymized, system_prompt, api_key
                )
            except RuntimeError:
                pass
            except anthropic.AuthenticationError:
                error = "Clé API invalide. Vérifie la clé saisie ci-dessous."
            except anthropic.APIConnectionError:
                error = "Impossible de contacter l'API. Vérifie ta connexion internet."
            except Exception as e:
                error = f"Erreur inattendue : {e}"

        st.session_state.processing = False
        if error:
            st.session_state.process_error = error
            st.rerun()
        else:
            goto(4)

    else:
        if st.session_state.process_error:
            st.error(st.session_state.process_error)
            st.session_state.process_error = ""

        with st.container(border=True):
            st.markdown("### Tout est prêt ?")
            rows = [
                ("ok", "Fichier", f"{st.session_state.uploaded_filename} · {n_lines} lignes"),
                ("ok", "Glossaire", f"{n_gloss} acronymes seront développés"),
                (
                    "ok" if n_anon else "ko",
                    "Anonymisation",
                    f"{n_anon} remplacement(s) appliqué(s) avant l'envoi" if n_anon
                    else "aucun remplacement configuré",
                ),
            ]
            html_rows = "".join(
                f'<div class="row"><span class="ic {k}">{"✓" if k == "ok" else "!"}</span>'
                f'<span class="lb">{label}</span><span class="dt">{detail}</span></div>'
                for k, label, detail in rows
            )
            st.markdown(f'<div class="sg-recap">{html_rows}</div>', unsafe_allow_html=True)

            st.write("")
            demo = st.checkbox(
                "Mode démo — parcourir l'outil avec un récit fictif, sans appel API",
                key="demo_toggle",
            )

            if not demo:
                st.markdown("**Clé API Anthropic**")
                if "api_key_widget" not in st.session_state:
                    st.session_state.api_key_widget = st.session_state.api_key_store
                st.text_input(
                    "Clé API Anthropic",
                    key="api_key_widget",
                    label_visibility="collapsed",
                    placeholder="sk-ant-…",
                    help="Obtenir une clé sur console.anthropic.com",
                    autocomplete="off",
                )
                st.session_state.api_key_store = st.session_state.api_key_widget
                api_key = st.session_state.api_key_store.strip()
                if not api_key:
                    st.markdown(
                        chip("Clé API requise pour lancer le traitement", "warn"),
                        unsafe_allow_html=True,
                    )

            st.write("")
            can_run = bool(st.session_state.raw_text and (demo or api_key))
            if st.button("Lancer le traitement", type="primary", disabled=not can_run):
                st.session_state.demo_run = demo
                st.session_state.review_done = False
                st.session_state.processing = True
                st.rerun()
            if demo:
                st.caption("Mode démo : récit fictif généré localement, aucun crédit consommé.")
            else:
                st.caption("30 à 90 secondes · ~0,05 € par document de 10 pages")

# ══════════════════════════════════════════════════════════════════
# Étape 4 — Relire & exporter
# ══════════════════════════════════════════════════════════════════
elif step == 4:
    result = st.session_state.result_text

    if not result:
        st.info("Aucun récit généré pour le moment. Reviens à l'étape « Lancer » via le fil d'ariane.")
    else:
        n_uncertain = result.count("[passage incertain")
        n_notes = result.count("[note suggérée]")

        # ── 1. Le récit : état d'un coup d'œil ────────────────────
        with st.container(border=True):
            st.markdown("### Votre récit est prêt")

            def stat(icon, value, label, kind):
                return (
                    f'<div class="sg-stat"><div class="ic {kind}">{icon}</div>'
                    f'<div class="v">{value}</div><div class="l">{label}</div></div>'
                )

            stats = []
            if n_uncertain:
                stats.append(stat("!", n_uncertain, "passage(s) incertain(s) à relire ci-dessous", "warn"))
            else:
                stats.append(stat("✓", "0", "passage incertain — récit fidèle aux notes", "ok"))
            if n_notes:
                stats.append(stat("✎", n_notes, "note(s) de bas de page suggérée(s)", "info"))
            if st.session_state.demo_run:
                stats.append(stat("⚙", "Démo", "récit fictif de démonstration", "neutral"))
            st.markdown(f'<div class="sg-stats">{"".join(stats)}</div>', unsafe_allow_html=True)

            if n_uncertain:
                st.caption(
                    "Les passages non corrigés resteront soulignés en orange dans le document final."
                )
            with st.expander("Aperçu du récit"):
                st.text_area(
                    "Aperçu du récit",
                    value=st.session_state.result_text,
                    height=420,
                    label_visibility="collapsed",
                )

        # ── 2. Révision des passages incertains (sens unique) ─────
        if n_uncertain > 0 and not st.session_state.review_done:
            with st.container(border=True):
                title_col, count_col = st.columns([4, 1.3])

                if not st.session_state.uncertain_passages:
                    st.session_state.uncertain_passages = extract_uncertain_passages(result)
                    st.session_state.uncertain_idx = 0

                passages = st.session_state.uncertain_passages
                n = len(passages)
                idx = st.session_state.uncertain_idx

                with title_col:
                    st.markdown("### Relire les passages incertains")
                with count_col:
                    st.markdown(
                        f"<div class='sg-passage-nav'>Passage {idx + 1} sur {n}</div>",
                        unsafe_allow_html=True,
                    )
                st.caption(
                    "Compare avec tes notes et saisis la correction, puis valide — "
                    "laisse vide pour conserver le marqueur."
                )

                def _save_current_correction():
                    val = st.session_state.get(f"corr_widget_{idx}", "")
                    st.session_state.corrections[idx] = val

                passage = passages[idx]

                # Contexte du passage — pleine largeur, plus lisible
                parts = []
                if passage["context_before"]:
                    parts.append(
                        f'<span style="color:#8A857A;font-size:0.85em">{html_lib.escape(passage["context_before"])}</span>'
                    )
                highlighted_line = html_lib.escape(passage["line"]).replace(
                    html_lib.escape("[passage incertain — à vérifier]"),
                    '<mark style="background:#FFC933;font-weight:600">[passage incertain — à vérifier]</mark>',
                )
                parts.append(highlighted_line)
                if passage["context_after"]:
                    parts.append(
                        f'<span style="color:#8A857A;font-size:0.85em">{html_lib.escape(passage["context_after"])}</span>'
                    )
                st.markdown(
                    '<div style="font-size:0.92em;line-height:1.65;padding:14px 16px;'
                    'background:#FDFCF7;border-radius:8px;border:1px solid #E7E2D0;color:#1B2B27">'
                    + "<br><br>".join(parts) + "</div>",
                    unsafe_allow_html=True,
                )
                st.write("")

                # Correction : champ + bouton côte à côte
                if f"corr_widget_{idx}" not in st.session_state:
                    st.session_state[f"corr_widget_{idx}"] = (
                        st.session_state.corrections.get(idx, "")
                    )
                corr_col, btn_col = st.columns([3.2, 1.2])
                with corr_col:
                    st.text_area(
                        "Remplacer le marqueur par :",
                        key=f"corr_widget_{idx}",
                        height=90,
                        placeholder="Saisis la correction — ou laisse vide pour conserver le marqueur",
                        label_visibility="collapsed",
                    )
                with btn_col:
                    is_last = idx == n - 1
                    if st.button("Valider" if not is_last else "Valider et terminer",
                                 key=f"btn_valider_{idx}", use_container_width=True,
                                 type="primary",
                                 help="Valider et passer au suivant (Ctrl+Entrée)"):
                        _save_current_correction()
                        if not is_last:
                            st.session_state.uncertain_idx += 1
                            st.rerun()
                        else:
                            mapping = {
                                f"correction_{i}": st.session_state.corrections.get(i, "")
                                for i in range(n)
                            }
                            st.session_state.result_text = apply_uncertain_corrections(
                                result, mapping, n
                            )
                            st.session_state.uncertain_passages = []
                            st.session_state.uncertain_idx = 0
                            st.session_state.corrections = {}
                            st.session_state.review_done = True
                            st.session_state.flash = "Relecture terminée — corrections appliquées."
                            st.rerun()

                # Ctrl+Entrée → clic sur le bouton Valider
                st.components.v1.html("""
<script>
(function() {
  function attach() {
    try {
      var tas = window.parent.document.querySelectorAll('textarea');
      tas.forEach(function(ta) {
        if (!ta._ctrlEnterBound) {
          ta._ctrlEnterBound = true;
          ta.addEventListener('keydown', function(e) {
            if ((e.ctrlKey || e.metaKey) && e.key === 'Enter') {
              e.preventDefault();
              var btns = window.parent.document.querySelectorAll('button');
              for (var i = 0; i < btns.length; i++) {
                if (btns[i].innerText.trim().indexOf('Valider') === 0) { btns[i].click(); return; }
              }
            }
          });
        }
      });
    } catch(err) {}
  }
  attach(); setTimeout(attach, 400); setTimeout(attach, 1200);
})();
</script>
""", height=0)

        # ── 3. Exporter / recommencer ──────────────────────────────
        st.write("")
        docx_buf = generate_docx(st.session_state.result_text)
        out_name = (
            st.session_state.uploaded_filename.replace(".docx", "_récit.docx")
            or "récit_audience.docx"
        )
        dl_col, new_col, _ = st.columns([2.2, 2.2, 2.6])
        with dl_col:
            st.download_button(
                label="Télécharger le récit (.docx)",
                data=docx_buf,
                file_name=out_name,
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                type="primary",
                use_container_width=True,
            )
        with new_col:
            if st.button("Traiter un nouveau document", use_container_width=True):
                st.session_state.raw_text = ""
                st.session_state.result_text = ""
                st.session_state.uploaded_filename = ""
                st.session_state.uncertain_passages = []
                st.session_state.uncertain_idx = 0
                st.session_state.corrections = {}
                st.session_state.anon_raw = ""
                st.session_state.demo_run = False
                st.session_state.review_done = False
                if "anon_widget" in st.session_state:
                    del st.session_state["anon_widget"]
                goto(1)
