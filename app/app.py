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
from docx.enum.text import WD_COLOR_INDEX
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

# Marqueur posé par l'IA sur les passages douteux : il reste dans le document
# exporté, surligné en jaune, pour relecture dans Word.
MARKER_RE = re.compile(r"(\[passage incertain[^\]]*\])")


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

def add_text_runs(paragraph, text, bold=False, italic=False):
    """Ajoute le texte au paragraphe ; les passages incertains sont surlignés en jaune."""
    for seg in MARKER_RE.split(text):
        if not seg:
            continue
        run = paragraph.add_run(seg)
        if bold:
            run.bold = True
        if italic:
            run.italic = True
        if MARKER_RE.fullmatch(seg):
            run.font.highlight_color = WD_COLOR_INDEX.YELLOW


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
            add_text_runs(paragraph, part[2:-2], bold=True)
        elif part.startswith("*") and part.endswith("*"):
            add_text_runs(paragraph, part[1:-1], italic=True)
        else:
            add_text_runs(paragraph, part)


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
            add_text_runs(p, s[1:-1], italic=True)
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
        else:
            # Paragraphe général — parse le markdown inline
            p = doc.add_paragraph()
            add_inline_markdown(p, s)

    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf

MOTS_PAR_PAGE = 450  # page Word A4 par défaut (Calibri 11, marges 2,54 cm), titres compris


def estimate_pages(text):
    """Nombre de pages estimé du .docx exporté. python-docx ne met pas en page :
    le nombre exact n'est connu qu'à l'ouverture dans Word, d'où une estimation au nombre de mots."""
    n_mots = len(re.sub(r"[*#]", " ", text).split())
    return max(1, -(-n_mots // MOTS_PAR_PAGE))


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

# Grain papier de la carte foncée (bruit SVG, aucune image externe)
GRAIN = (
    "url(\"data:image/svg+xml;utf8,<svg xmlns='http://www.w3.org/2000/svg' width='240' height='240'>"
    "<filter id='n'><feTurbulence type='fractalNoise' baseFrequency='0.8' numOctaves='3' stitchTiles='stitch'/>"
    "<feColorMatrix values='0 0 0 0 0.01  0 0 0 0 0.10  0 0 0 0 0.11  0 0 0 0.75 0'/></filter>"
    "<rect width='100%25' height='100%25' filter='url(%23n)'/></svg>\")"
)

CUSTOM_CSS = """
<style>
@font-face { font-family: 'Gelica'; font-weight: 400; src: url('/app/static/Gelica-Regular.otf') format('opentype'); }
@font-face { font-family: 'Gelica'; font-style: italic; font-weight: 400; src: url('/app/static/Gelica-Italic.otf') format('opentype'); }
@font-face { font-family: 'Poppins'; font-weight: 400; src: url('/app/static/Poppins-Regular.ttf') format('truetype'); }
@font-face { font-family: 'Poppins'; font-weight: 500; src: url('/app/static/Poppins-Medium.ttf') format('truetype'); }
@font-face { font-family: 'Poppins'; font-weight: 600; src: url('/app/static/Poppins-SemiBold.ttf') format('truetype'); }

/* ── Base ─────────────────────────────────────────── */
html, body, .stApp, button, input, textarea, select,
[class*="st-"]:not([data-testid="stIconMaterial"]):not(.material-symbols-rounded):not([data-testid="stExpanderToggleIcon"]) {
  font-family: 'Poppins', system-ui, -apple-system, 'Segoe UI', Arial, sans-serif;
}
.stApp { background: #FBF8EF; }
[data-testid="stMain"] { overflow-x: clip; }
h1, h2, h3 {
  font-family: 'Gelica', Georgia, serif !important;
  font-weight: 400 !important; color: #14201D !important; letter-spacing: -0.01em;
}
.block-container { padding-top: 2.2rem; padding-bottom: 0; max-width: 1280px; }
#MainMenu, footer, [data-testid="stToolbar"] { visibility: hidden; }
header[data-testid="stHeader"] { background: transparent; }

/* Restaurer la police des icônes Material (écrasée par la règle de police du corps) */
[data-testid="stIconMaterial"], [data-testid="stExpanderToggleIcon"],
.material-symbols-rounded, .material-symbols-outlined, .material-icons {
  font-family: 'Material Symbols Rounded' !important;
  font-weight: normal !important;
}

/* Masquer les mentions « Press Enter to apply » (déroutantes pour les novices) */
[data-testid="InputInstructions"] { display: none !important; }
/* Pas de lien d'ancre à côté des titres */
[data-testid="stHeaderActionElements"] { display: none !important; }
.st-key-scrolltop { position: absolute; height: 0; overflow: hidden; }

/* Clé API : champ texte simple (pas de type password → pas de gestionnaire
   de mots de passe), masqué visuellement par CSS */
.st-key-api_key_widget input { -webkit-text-security: disc; }

/* ── En-tête ──────────────────────────────────────── */
.sg-top { display: flex; justify-content: space-between; align-items: center; gap: 16px 32px; flex-wrap: wrap; }
.sg-kicker {
  background: #14201D; color: #FFFFFF; padding: 3px 10px;
  font-size: 0.72rem; font-weight: 600; letter-spacing: 0.08em; text-transform: uppercase;
}
.sg-tagline { color: #5C6660; font-size: 0.85rem; }
.sg-title {
  font-family: 'Gelica', Georgia, serif; font-weight: 400; color: #14201D;
  text-align: center; font-size: 3.2rem; line-height: 1.1; margin: 40px 0 0 0;
  animation: sg-fade .4s ease both;
}
.sg-title.big { font-size: 5rem; margin-top: 96px; }
/* En-tête compact : le titre ne garde pas sa marge de l'accueil ;
   tout le contenu en dessous remonte de 32 px */
[class*="st-key-header"] .sg-title { margin: 0; animation: none; }  /* pas de clignotement d'un écran à l'autre */
[class*="st-key-header"] { margin-bottom: -32px; }
.sg-lead {
  font-family: 'Gelica', Georgia, serif; font-size: 1.6rem !important; color: #14201D;
  text-align: center; margin: 56px 0 8px 0 !important;
}
.sg-done {
  font-family: 'Gelica', Georgia, serif; font-size: 2.3rem !important; color: #14201D;
  text-align: center; margin: 4px 0 16px 0 !important; animation: sg-fade .5s .1s ease both;
}
.sg-note { text-align: center; color: #2E3A36; font-size: 0.98rem !important; margin: 0 !important; }
@keyframes sg-appear { from { opacity: 0; } to { opacity: 1; } }
@keyframes sg-fade { from { opacity: 0; transform: translateY(8px); } to { opacity: 1; transform: none; } }

/* ── Pile de feuilles ─────────────────────────────── */
/* La pile est coupée net au bas de la feuille du dessus (overflow: hidden) :
   la page s'arrête là, on ne peut donc jamais défiler jusqu'au bas des feuillets.
   --side : marge latérale où dépassent les feuillets ; --w : largeur de la feuille. */
[class*="st-key-stack_"] {
  --side: min(120px, 8vw); --w: calc(100% - 2 * var(--side));
  position: relative; isolation: isolate; overflow: hidden; box-sizing: border-box;
  width: 100%; max-width: calc(780px + 2 * var(--side)); margin: 16px auto 0 auto;
  padding: 40px var(--side) 0 var(--side);
}
[class*="st-key-stack_"]::before, [class*="st-key-stack_"]::after {
  content: ""; position: absolute; z-index: 0; pointer-events: none; bottom: -400px;
}
[class*="st-key-stack_"]::before {
  top: 0; left: calc(var(--side) + var(--w) * 0.14); right: calc(var(--side) - var(--w) * 0.14);
  background: #E9E1CE; transform: rotate(5deg); transform-origin: top left;
  animation: sg-appear .5s ease both;
}
[class*="st-key-stack_"]::after {
  top: 152px; left: calc(var(--side) - var(--w) * 0.06); right: calc(var(--side) + var(--w) * 0.30);
  background: #E4DBC6; transform: rotate(-1.8deg); transform-origin: top left;
  animation: sg-appear .5s .05s ease both;
}
[class*="st-key-sheet_"] {
  position: relative; z-index: 1; background: #FFFFFF;
  padding: 56px 72px 80px 72px; min-height: calc(100vh - 200px); gap: 16px;
  animation: sg-sheet-in .45s cubic-bezier(.2,.7,.2,1) both;
}
/* Markdown : pas de marge négative ni de marge sous le dernier paragraphe,
   pour que les espacements ci-dessous soient exacts et les lignes bien centrées */
[data-testid="stMain"] [data-testid="stMarkdownContainer"] { margin-bottom: 0 !important; }
[class*="st-key-sheet_"] [data-testid="stMarkdownContainer"] > :last-child:not(h3),
[class*="st-key-sheet_"] [data-testid="stMarkdownContainer"] > :last-child > :last-child:not(h3) { margin-bottom: 0 !important; }
@keyframes sg-sheet-in {
  from { opacity: 0; transform: translateY(32px) rotate(-0.8deg); }
  to   { opacity: 1; transform: none; }
}
[class*="st-key-sheet_"] h3 { font-size: 2.1rem !important; padding: 0 !important; margin: 0 0 16px 0 !important; line-height: 1.2 !important; }
/* Titre « Acronymes » sur la même ligne que le glossaire : la marge passe sur la ligne */
.st-key-acro_head h3 { margin: 0 !important; }
.st-key-acro_head { margin-bottom: 16px; }
/* Boutons d'action principaux : 56 px au-dessus (40 px + 16 px d'écart) */
.st-key-btn_acro_next, .st-key-btn_anon_next,
[data-testid="stLayoutWrapper"]:has(> .st-key-launch_row) { margin-top: 40px; }
/* Anonymisation : 32 px au-dessus du champ */
.st-key-anon_widget { margin-top: 16px; }
[class*="st-key-sheet_"] p, [class*="st-key-sheet_"] li { font-size: 0.98rem; color: #2E3A36; line-height: 1.65; }
[class*="st-key-sheet_"] [data-testid="stWidgetLabel"] p { font-size: 0.95rem; color: #14201D; }

/* Accueil : pile vide, plus étroite */
.st-key-stack_home { max-width: calc(680px + 2 * var(--side)); margin-top: 80px; }
.st-key-sheet_home { background: #EFEADB; min-height: 100vh; padding: 0; }
.st-key-sheet_result { min-height: 100vh; padding: 0; }
/* Accueil et résultat : seul le haut de la pile dépasse en bas de l'écran, pas de défilement */
html:has(.st-key-stack_home, .st-key-stack_result), body:has(.st-key-stack_home, .st-key-stack_result),
.stApp:has(.st-key-stack_home, .st-key-stack_result),
.stApp:has(.st-key-stack_home, .st-key-stack_result) [data-testid="stMain"] { overflow: hidden !important; }

/* Feuille foncée à grain (traitement, aperçu du récit) */
.st-key-sheet_processing, .st-key-sheet_result {
  background-color: #06393B;
  background-image: __GRAIN__, radial-gradient(ellipse at 30% 20%, rgba(20,90,90,.35), transparent 55%);
  border: 2px solid #14201D; color: #EDE8DA;
}
.st-key-sheet_processing { justify-content: center; text-align: center; }
.sg-processing h3 { color: #FFFFFF !important; font-size: 1.9rem !important; margin: 26px 0 10px 0 !important; }
.sg-processing p { color: #9DBBB0 !important; margin: 0; }

/* ══ Transitions ═════════════════════════════════════════════════
   Courbes : arrivée « posée » avec un léger rebond (sg-land), sortie nette. */

/* 1 · Accueil → Acronymes : les feuilles du bas montent se poser une à une
   (feuille gauche, feuille droite, puis la feuille du dessus), le titre se range
   dans l'en-tête, puis le contenu de la feuille apparaît ligne à ligne. */
.st-key-stack_acronymes_intro::after  { animation: sg-rise-left 1s cubic-bezier(.2,1.1,.35,1) both; }
.st-key-stack_acronymes_intro::before { animation: sg-rise-right 1s .16s cubic-bezier(.2,1.1,.35,1) both; }
.st-key-sheet_acronymes_intro { animation: sg-rise-top 1.05s .32s cubic-bezier(.2,1.08,.35,1) both; }
.st-key-sheet_acronymes_intro > * { animation: sg-fade .55s ease both; }
.st-key-sheet_acronymes_intro > :nth-child(1) { animation-delay: 1.05s; }
.st-key-sheet_acronymes_intro > :nth-child(2) { animation-delay: 1.13s; }
.st-key-sheet_acronymes_intro > :nth-child(3) { animation-delay: 1.21s; }
.st-key-sheet_acronymes_intro > :nth-child(4) { animation-delay: 1.29s; }
.st-key-sheet_acronymes_intro > :nth-child(n+5) { animation-delay: 1.37s; }
@keyframes sg-rise-left {
  from { transform: translateY(80vh) rotate(-7deg); box-shadow: 0 50px 90px rgba(20,32,29,.22); }
  to   { transform: rotate(-1.8deg); box-shadow: 0 0 0 rgba(20,32,29,0); }
}
@keyframes sg-rise-right {
  from { transform: translateY(80vh) rotate(11deg); box-shadow: 0 50px 90px rgba(20,32,29,.22); }
  to   { transform: rotate(5deg); box-shadow: 0 0 0 rgba(20,32,29,0); }
}
@keyframes sg-rise-top {
  0%   { transform: translateY(80vh) rotate(2.5deg); background-color: #EFEADB;
         box-shadow: 0 60px 100px rgba(20,32,29,.25); }
  70%  { background-color: #EFEADB; }
  100% { transform: none; background-color: #FFFFFF; box-shadow: 0 0 0 rgba(20,32,29,0); }
}
.st-key-header_intro .sg-title { animation: sg-title-dock .95s cubic-bezier(.16,1,.3,1) both; }
.st-key-header_intro [data-testid="stButton"] { animation: sg-fade .5s .85s ease both; }
@keyframes sg-title-dock {
  from { transform: translateY(150px) scale(1.56); }
  to   { transform: none; }
}

/* 2 · Lancement : la feuille blanche se lève et découvre la page foncée,
   sur laquelle des notes griffonnées sont réécrites au propre, ligne à ligne. */
.st-key-stack_processing::before, .st-key-stack_processing::after { animation: none; }
.st-key-sheet_processing { animation: none; overflow: hidden; }
.st-key-sheet_processing::after {
  content: ""; position: absolute; inset: 0; z-index: 3; background: #FFFFFF; pointer-events: none;
  animation: sg-unveil .85s cubic-bezier(.75,0,.25,1) both;
}
@keyframes sg-unveil { from { clip-path: inset(0 0 0 0); } to { clip-path: inset(0 0 100% 0); } }
.sg-processing {
  position: relative; z-index: 2; margin: 0 auto; padding: 56px 96px;
  /* Les traits passent derrière le texte, floutés et assombris pour rester lisibles */
  background: rgba(6,57,59,.6); backdrop-filter: blur(5px); -webkit-backdrop-filter: blur(5px);
  -webkit-mask-image: radial-gradient(farthest-side, #000 62%, transparent);
  mask-image: radial-gradient(farthest-side, #000 62%, transparent);
  animation: sg-fade .6s .7s ease both;
}

.sg-ink { position: absolute; inset: 40px 56px; z-index: 1; pointer-events: none; }
/* Le calque d'écriture se positionne par rapport à la feuille, pas à son bloc Markdown */
.st-key-sheet_processing [data-testid="stElementContainer"],
.st-key-sheet_processing [data-testid="stMarkdown"],
.st-key-sheet_processing [data-testid="stMarkdownContainer"] { position: static !important; }
.sg-ink svg { width: 100%; height: 100%; }
/* Notes brutes : traits de crayon ondulés, qui s'effacent quand la ligne est tapée */
.sg-ink .scr { fill: none; stroke: #9DBBB0; stroke-width: 1.4; stroke-linecap: round; opacity: .35;
  animation: ink-scr 14s linear infinite; }
/* Version au propre, façon machine à écrire : chaque ligne est une suite de « mots »
   (tirets) que l'on découvre par à-coups, une frappe à la fois, ligne après ligne.
   Un léger filtre de déplacement donne le grain irrégulier de l'encre sur le papier. */
.sg-ink .w  { fill: none; stroke-linecap: butt; }
.sg-ink .h  { stroke: #EDE8DA; stroke-width: 9; }
.sg-ink .lb { stroke: #9DBBB0; stroke-width: 4.6; }
.sg-ink .t  { stroke: #EDE8DA; stroke-width: 4.2; opacity: .8; }
.sg-ink .sp { stroke: #FFFFFF; stroke-width: 5.6; }
.sg-ink .n  { stroke: #9DBBB0; stroke-width: 3.4; }
.sg-ink .line { animation: ink-line 14s linear infinite; }
.sg-ink .rv { transform-box: fill-box; transform-origin: left center; transform: scaleX(0);
  animation: ink-type 14s infinite; animation-timing-function: steps(var(--n), end); }
.sg-ink .hl { fill: #FFC933; opacity: .3; transform-box: fill-box; transform-origin: left center;
  transform: scaleX(0); animation: ink-hl 14s cubic-bezier(.5,0,.2,1) infinite; }
.sg-ink .scr, .sg-ink .line, .sg-ink .rv, .sg-ink .hl { animation-delay: calc(.9s + var(--i) * .5s); }
@keyframes ink-scr { 0%, 5% { opacity: .35; } 8.6% { opacity: 0; } 90% { opacity: 0; } 96%, 100% { opacity: .35; } }
@keyframes ink-type { 0%, 5% { transform: scaleX(0); } 8.6%, 100% { transform: scaleX(1); } }
@keyframes ink-line { 0%, 88% { opacity: 1; } 94%, 100% { opacity: 0; } }
@keyframes ink-hl {
  0%, 9% { transform: scaleX(0); opacity: .3; } 11.5% { transform: scaleX(1); }
  88% { transform: scaleX(1); opacity: .3; } 94%, 100% { transform: scaleX(1); opacity: 0; }
}

/* 3 · Résultat : la page foncée redescend se ranger en bas de l'écran,
   puis le message, les chiffres et les actions apparaissent. */
.st-key-stack_result { animation: sg-drop 1s cubic-bezier(.2,1.06,.35,1) both; }
.st-key-stack_result::before, .st-key-stack_result::after, .st-key-sheet_result { animation: none; }
@keyframes sg-drop { from { transform: translateY(-58vh); } to { transform: none; } }
.stApp:has(.st-key-stack_result) .sg-done { animation: sg-fade .6s .55s ease both; }
.stApp:has(.st-key-stack_result) .sg-note { animation: sg-fade .6s .68s ease both; }
.stApp:has(.st-key-stack_result) .sg-stat { animation: sg-fade .6s .8s ease both; }
.stApp:has(.st-key-stack_result) .sg-stat:nth-child(2) { animation-delay: .9s; }
.stApp:has(.st-key-stack_result) .sg-stat:nth-child(3) { animation-delay: 1s; }
.st-key-dl, .st-key-link_new { animation: sg-fade .6s 1.12s ease both; }
.st-key-energybar { animation: sg-bar-up .6s 1.3s cubic-bezier(.2,1,.3,1) both; }
@keyframes sg-bar-up { from { transform: translateY(100%); } to { transform: none; } }

@media (prefers-reduced-motion: reduce) {
  *, *::before, *::after {
    animation-duration: .01ms !important; animation-delay: 0s !important;
    animation-iteration-count: 1 !important;
  }
}


/* ── Boutons ──────────────────────────────────────── */
button[kind="primary"], [data-testid="stBaseButton-primary"] {
  background: #14201D !important; color: #FFFFFF !important; border: none !important;
  border-radius: 2px !important; padding: 12px 32px !important; min-height: 54px !important;
  box-shadow: none !important; transition: background .15s ease, transform .15s ease;
}
button[kind="primary"] p, [data-testid="stBaseButton-primary"] p {
  color: #FFFFFF !important; font-size: 1.05rem !important; font-weight: 400 !important;
}
button[kind="primary"]:hover:enabled, [data-testid="stBaseButton-primary"]:hover:enabled {
  background: #0C3A34 !important; transform: translateY(-1px);
}
button[kind="primary"]:disabled, [data-testid="stBaseButton-primary"]:disabled {
  background: #DCD6C6 !important; cursor: not-allowed !important;
}
button[kind="primary"]:disabled p { color: #8A857A !important; }

button[kind="secondary"], [data-testid="stBaseButton-secondary"],
[data-testid="stPopoverButton"] {
  background: #E2DCC8 !important; color: #14201D !important; border: none !important;
  border-radius: 3px !important; box-shadow: none !important; min-height: 0 !important;
  padding: 7px 16px !important;
}
button[kind="secondary"] p, [data-testid="stPopoverButton"] p { font-weight: 500 !important; font-size: 0.95rem !important; white-space: nowrap; }
button[kind="secondary"]:hover:enabled, [data-testid="stPopoverButton"]:hover { background: #D6CEB6 !important; }

/* Liens (boutons tertiaires) : Retour, Modifier, Traiter un nouveau document */
[data-testid="stBaseButton-tertiary"] {
  color: #0E5B44 !important; padding: 0 !important; min-height: 0 !important;
}
[data-testid="stBaseButton-tertiary"] p { font-weight: 500 !important; font-size: 0.95rem !important; color: inherit !important; }
[data-testid="stBaseButton-tertiary"]:hover { color: #14201D !important; }
.st-key-link_new { align-items: center; }
.st-key-link_new button { color: #14201D !important; }
.st-key-link_new button p { text-decoration: underline; text-underline-offset: 4px; }

/* ── Champs ───────────────────────────────────────── */
[data-baseweb="input"], [data-baseweb="base-input"], [data-baseweb="textarea"] {
  background: #EFEADC !important; border: none !important; border-radius: 6px !important;
}
[data-baseweb="input"]:focus-within, [data-baseweb="textarea"]:focus-within {
  box-shadow: inset 0 0 0 1.5px #0E5B44;
}
[data-baseweb="input"] input, [data-baseweb="textarea"] textarea {
  background: transparent !important; font-size: 0.98rem; color: #14201D;
}
[data-baseweb="input"] input { padding: 12px 16px !important; }
[data-baseweb="textarea"] textarea { padding: 14px 16px !important; line-height: 1.6; }
[data-baseweb="input"] input::placeholder, [data-baseweb="textarea"] textarea::placeholder { color: #5C6660; opacity: 1; }

/* ── Accueil : zone de dépôt horizontale ──────────── */
.st-key-uploader { max-width: 780px; width: 100%; margin: 0 auto; }
.st-key-uploader [data-testid="stFileUploaderDropzone"] {
  background: #FFFFFF; border: 1.5px dashed #9C9580; border-radius: 10px;
  padding: 16px 22px 16px 40px; transition: border-color .15s ease, background .15s ease;
}
.st-key-uploader [data-testid="stFileUploaderDropzone"]:hover { border-color: #0E5B44; background: #FDFCF7; }
.st-key-uploader [data-testid="stFileUploaderDropzoneInstructions"] > span,
.st-key-uploader [data-testid="stFileUploaderDropzoneInstructions"] > div > span { display: none; }
.st-key-uploader [data-testid="stFileUploaderDropzoneInstructions"] > div::before {
  content: "Dépose le fichier .docx ici"; font-size: 1.1rem; color: #14201D;
}
.st-key-uploader [data-testid="stFileUploaderDropzone"] button { font-size: 0 !important; }
.st-key-uploader [data-testid="stFileUploaderDropzone"] button::after {
  content: "Parcourir…"; font-size: 0.95rem; font-weight: 500;
}

/* ── Écran Acronymes ──────────────────────────────── */
.st-key-filebar {
  background: #EFEADC; border: 1.5px solid #BFE3D0; border-radius: 8px; padding: 12px;
  margin-bottom: 18px;
}
.sg-file, .sg-gloss {
  display: flex; align-items: center; gap: 12px; margin: 0;
  color: #0E5B44; font-weight: 500; font-size: 0.95rem; white-space: nowrap;
}
.sg-gloss { justify-content: flex-end; }
.st-key-filebar .sg-file { font-size: calc(0.95rem - 2px); padding-left: 10px; min-width: 0; }
/* Nom de fichier long : tronqué avec « … », l'icône garde sa taille */
.st-key-filebar .sg-file svg { flex-shrink: 0; }
.st-key-filebar .sg-file span { min-width: 0; overflow: hidden; text-overflow: ellipsis; }
.st-key-filebar [data-testid="stColumn"]:first-child { min-width: 0; overflow: hidden; }
.st-key-filebar button[kind] p { font-size: calc(0.95rem - 2px) !important; }
.st-key-filebar [data-testid="stHorizontalBlock"] { gap: 4px; }
.st-key-filebar [data-testid="stMarkdownContainer"] p { margin: 0 !important; }

/* ── Chips / pastilles ────────────────────────────── */
.sg-chip {
  display: inline-flex; align-items: center; gap: 6px;
  padding: 4px 14px; border-radius: 999px; font-size: 0.85rem; font-weight: 500;
}
.sg-chip.ok   { background: #DFF2E7; color: #0E5B44; border: 1px solid #BFE3D0; }
.sg-chip.warn { background: #FDE4D5; color: #9A4B00; border: 1px solid #F5CBAA; }
.sg-chip.neutral { background: #F1EDE0; color: #5C6660; border: 1px solid #E3DECB; }
.sg-pill {
  display: inline-block; padding: 3px 14px; border-radius: 999px; margin-left: 10px;
  border: 1px solid #E3DECB; background: #FBF8EF; font-size: 0.85rem; font-weight: 500;
  vertical-align: middle; animation: sg-fade .2s ease both;
}
.sg-keylabel { font-weight: 600; font-size: 1.15rem; color: #14201D; margin: 10px 0 0 0; }
.sg-hint { font-size: 0.82rem !important; color: #8A857A !important; margin: -6px 0 0 0; }
.sg-cost { text-align: right; color: #2E3A36; font-size: 0.98rem; margin: 0; }

/* ── Récapitulatif « Tout est prêt ? » ─────────────── */
.sg-recap-row { display: flex; align-items: center; gap: 16px; }
.sg-recap-row .ic {
  width: 30px; height: 30px; border-radius: 50%; flex-shrink: 0;
  display: inline-flex; align-items: center; justify-content: center;
  font-size: 0.85rem; font-weight: 700;
}
.sg-recap-row .ic.ok { background: #0E5B44; color: #FFFFFF; }
.sg-recap-row .ic.ko { background: #FDE4D5; color: #C0501F; }
.sg-recap-row .lb { font-weight: 600; color: #14201D; min-width: 150px; }
.sg-recap-row .dt { color: #2E3A36; }
.sg-error {
  background: #FDE4D5; border-left: 3px solid #C0501F; padding: 12px 16px;
  color: #6B3300; font-size: 0.92rem; border-radius: 4px;
}

/* ── Résultat : statistiques ──────────────────────── */
.sg-stats { display: flex; justify-content: center; align-items: stretch; margin: 36px 0 30px 0; flex-wrap: wrap; }
.sg-stat {
  display: flex; flex-direction: column; align-items: center; gap: 10px;
  text-align: center; width: 240px; padding: 0 20px; animation: sg-fade .5s ease both;
}
.sg-stat:nth-child(2) { animation-delay: .1s; } .sg-stat:nth-child(3) { animation-delay: .2s; }
.sg-stat + .sg-stat { border-left: 1.5px solid #D9D3C2; }
.sg-stat .ic { width: 60px; height: 60px; border-radius: 50%; display: flex; align-items: center; justify-content: center; }
.sg-stat .ic.warn { background: #FBDDD2; }
.sg-stat .ic.ok { background: #DFF2E7; }
.sg-stat .ic.info { background: #E5E2FB; }
.sg-stat .ic.neutral { background: #EFEBE0; }
.sg-stat .v { font-family: 'Gelica', Georgia, serif; font-size: 1.9rem; color: #14201D; line-height: 1.1; }
.sg-stat .l { font-size: 0.95rem; color: #5C6660; line-height: 1.4; }
.st-key-dl { max-width: 320px; width: 100%; margin: 0 auto; }
.st-key-dl button { width: 100%; }

/* ── Barre « Énergie consommée » (fixe, en bas) ───── */
.st-key-energybar {
  position: fixed; left: 0; right: 0; bottom: 0; z-index: 999;
  background: #14201D; padding: 4px 0; align-items: center;
}
.st-key-energybar [data-testid="stPopoverButton"] {
  background: transparent !important; color: #FFFFFF !important; padding: 10px 24px !important;
}
.st-key-energybar [data-testid="stPopoverButton"] p { color: #FFFFFF !important; font-weight: 400 !important; font-size: 1rem !important; }
.st-key-energybar button[data-testid="stPopoverButton"]:hover:enabled,
.st-key-energybar button[data-testid="stPopoverButton"]:focus,
.st-key-energybar button[data-testid="stPopoverButton"]:focus-visible,
.st-key-energybar button[data-testid="stPopoverButton"]:active,
.st-key-energybar button[data-testid="stPopoverButton"][aria-expanded="true"] {
  background: rgba(255, 255, 255, 0.05) !important; box-shadow: none !important; color: #FFFFFF !important;
}
[data-testid="stPopoverBody"] { min-width: 420px; }
/* Panneau Énergie : centré, posé sur la barre, ancré par le bas. Quand « Méthode et
   sources » s'ouvre, le panneau grandit vers le haut et reste entièrement visible. */
[data-testid="stPopoverBody"] {
  position: fixed !important; top: auto !important; right: auto !important;
  left: 50% !important; bottom: 56px !important; transform: translateX(-50%) !important;
  width: min(640px, calc(100vw - 32px)); max-height: calc(100vh - 56px - 16px);
  overflow-y: auto;
}

/* Comparaison énergie (panneau) */
.sg-compare { background: #FBF8EF; border: 1px solid #E7E2D0; border-radius: 10px; padding: 14px 16px; }
.sg-compare .row { display: flex; justify-content: space-between; gap: 12px; padding: 4px 0; font-size: 0.88rem; }
.sg-compare .lb { color: #5C6660; }
.sg-compare .v { color: #14201D; font-weight: 500; text-align: right; }
.sg-compare .gain { margin-top: 10px; padding-top: 12px; border-top: 1px dashed #C9C4B4; text-align: center; }
.sg-compare .big { display: block; font-family: 'Gelica', Georgia, serif; font-size: 2rem; font-weight: 400; color: #0E5B44; line-height: 1.1; }
.sg-compare .txt { display: block; font-size: 0.82rem; color: #5C6660; margin-top: 4px; }
.sg-small, .sg-small li { font-size: 0.74rem !important; color: #8A857A !important; line-height: 1.5; }
.sg-small a { color: #8A857A !important; }

/* Fenêtre d'aperçu des notes */
.sg-preview {
  white-space: pre-wrap; font-size: 0.9rem; line-height: 1.6; color: #2E3A36;
  background: #FBF8EF; border-radius: 8px; padding: 16px 18px; max-height: 60vh; overflow: auto;
}
</style>
""".replace("__GRAIN__", GRAIN)


def _svg(body, size=24):
    return (
        f'<svg width="{size}" height="{size}" viewBox="0 0 24 24" fill="none" stroke="currentColor" '
        f'stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">{body}</svg>'
    )

ICON_DOC = (
    '<svg width="26" height="26" viewBox="0 0 26 26" fill="none" stroke="#14201D" stroke-width="1.8" '
    'stroke-linejoin="round" aria-hidden="true"><path d="M6 3h9l5 5v6M6 3v10"/><path d="M15 3v5h5"/>'
    '<text x="2" y="23" font-size="8" font-weight="700" fill="#14201D" stroke="none" '
    'font-family="Arial">DOC</text></svg>'
)
ICON_BOOK = (
    '<svg width="28" height="28" viewBox="0 0 26 26" fill="none" stroke="#14201D" stroke-width="1.8" '
    'stroke-linejoin="round" aria-hidden="true"><path d="M2 5h8a3 3 0 0 1 3 3v13a2 2 0 0 0-2-2H2z"/>'
    '<path d="M24 5h-8a3 3 0 0 0-3 3v13a2 2 0 0 1 2-2h9z"/><path d="M16 9h5M16 12h5M16 15h5"/></svg>'
)
ICON_BINOCULARS = _svg(
    '<circle cx="6.5" cy="16" r="3.5"/><circle cx="17.5" cy="16" r="3.5"/>'
    '<path d="M3 16l2-9a2 2 0 0 1 4 0v9M21 16l-2-9a2 2 0 0 0-4 0v9M10 12h4"/>'
)
ICON_COLUMNS = _svg(
    '<rect x="4" y="4" width="6.5" height="16" rx="2"/><rect x="13.5" y="4" width="6.5" height="16" rx="2"/>'
    '<path d="M4 9h6.5M4 14h6.5M13.5 9h6.5M13.5 14h6.5"/>'
)
ICON_PAGES = _svg('<path d="M14 3H7a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h10a2 2 0 0 0 2-2V8z"/>'
                  '<path d="M14 3v5h5M9 13h6M9 17h6"/>')
ICON_CHECK = _svg('<path d="M5 12.5l4.5 4.5L19 7.5"/>')

def ink_svg():
    """Fond animé de l'écran de traitement : des notes griffonnées (traits ondulés) sont
    retapées au propre ligne à ligne, façon machine à écrire (titres, champs, prises de
    parole, narration) ; un passage incertain est surligné en jaune. Animé en CSS (.sg-ink)."""
    import random
    rnd = random.Random(7)  # même dessin à chaque fois
    # (type, longueur du préfixe, fin de ligne, passage incertain (début, fin) ou None)
    rows = [
        ("h", 0, 300, None), ("f", 70, 330, None), ("f", 70, 420, None),
        ("f", 70, 260, None), ("f", 70, 380, None), None,
        ("h", 0, 240, None), ("s", 96, 560, None), ("t", 0, 540, None), ("t", 0, 400, None),
        ("n", 0, 520, None), ("n", 0, 330, None),
        ("s", 124, 560, None), ("t", 0, 548, (210, 340)), ("t", 0, 450, None), None,
        ("h", 0, 210, None), ("s", 110, 560, None), ("t", 0, 300, None),
    ]

    def words(length, big):
        """Motif de tirets imitant des mots : « mot espace mot espace… »."""
        lo, hi, gap = (22, 64, 10) if big else (9, 46, 6)
        dash, total = [], 0
        while total < length:
            w = rnd.randint(lo, hi)
            dash += [w, gap]
            total += w + gap
        return " ".join(map(str, dash))

    x0, y, parts = 40, 52, []
    for i, row in enumerate(rows):
        if row is None:
            y += 26
            continue
        kind, pre, end, unsure = row
        gap = 14 if kind == "h" else 30
        var = f'style="--i:{i};--n:{max(6, (end - x0) // 8)}"'
        # Griffonnage de la note brute : ligne ondulée, légèrement de travers
        d, x = f"M{x0 + rnd.randint(-6, 10)} {y + rnd.uniform(-3, 3):.1f}", x0
        while x < end - 20:
            dx = rnd.randint(14, 24)
            x += dx
            d += f" q{dx // 2} {rnd.choice((-1, 1)) * rnd.uniform(3, 7):.1f} {dx} 0"
        parts.append(f'<path class="scr" {var} d="{d}" transform="rotate({rnd.uniform(-1.4, 1.4):.2f} {x0} {y})"/>')
        # Ligne tapée : révélée par un masque qui avance par à-coups
        parts.append(
            f'<clipPath id="sg-r{i}"><rect class="rv" {var} x="{x0 - 2}" y="{y - 10}" '
            f'width="{end - x0 + 4}" height="20"/></clipPath>'
        )
        line = []
        if unsure:
            line.append(f'<rect class="hl" {var} x="{unsure[0]}" y="{y - 9}" width="{unsure[1] - unsure[0]}" height="18" rx="2"/>')
        if pre:
            cls = "lb" if kind == "f" else "sp"
            line.append(f'<path class="w {cls}" d="M{x0} {y} H{x0 + pre}" stroke-dasharray="{words(pre, False)}"/>')
            line.append(f'<path class="w t" d="M{x0 + pre + 12} {y} H{end}" stroke-dasharray="{words(end - x0 - pre, False)}"/>')
        else:
            line.append(f'<path class="w {kind}" d="M{x0} {y} H{end}" stroke-dasharray="{words(end - x0, kind == "h")}"/>')
        parts.append(f'<g class="line" {var} clip-path="url(#sg-r{i})">' + "".join(line) + "</g>")
        y += gap
    return (
        '<div class="sg-ink" aria-hidden="true"><svg viewBox="0 0 600 640" '
        'preserveAspectRatio="xMidYMid meet" xmlns="http://www.w3.org/2000/svg">'
        '<defs><filter id="sg-rough" x="-2%" y="-20%" width="104%" height="140%">'
        '<feTurbulence type="fractalNoise" baseFrequency="0.9" numOctaves="2" seed="4"/>'
        '<feDisplacementMap in="SourceGraphic" scale="2.4"/></filter></defs>'
        '<g filter="url(#sg-rough)">' + "".join(parts) + "</g></svg></div>"
    )


TAGLINE = (
    "Des notes d'audience brutes au récit structuré. "
    "Fidèle au texte, et anonymisé avant tout envoi à l'IA."
)


# Écran précédent de chaque écran (le retour à l'accueil oublie le document)
BACK_TARGET = {1: 0, 2: 1, 3: 2}


def render_header(step):
    """Accueil : surtitre complet + accroche, puis le grand titre.
    Écrans suivants : en-tête compact sur une seule ligne, titre centré, avec à gauche
    « ‹ Retour » (étapes 2 à 4) ou « MSDE » (traitement, résultat)."""
    if step > 0:
        back = None if st.session_state.processing else BACK_TARGET.get(step)
        intro = step == 1 and st.session_state.intro
        with st.container(key="header_intro" if intro else "header"):
            left, mid, _ = st.columns([1, 1.4, 1], vertical_alignment="center")
        with left:
            if back is None:
                st.markdown('<span class="sg-kicker">MSDE</span>', unsafe_allow_html=True)
            elif st.button("‹ Retour", key="link_back", type="tertiary"):
                if back == 0:
                    reset_document()
                goto(back)
        with mid:
            st.markdown(
                '<div class="sg-title" role="heading" aria-level="1">SuperGreffier</div>',
                unsafe_allow_html=True,
            )
        return
    st.markdown(
        '<div class="sg-top"><span class="sg-kicker">MSDE · Automatisation de mise en forme</span>'
        f'<span class="sg-tagline">{TAGLINE}</span></div>'
        f'<div class="sg-title{" big" if step == 0 else ""}" role="heading" aria-level="1">'
        "SuperGreffier</div>",
        unsafe_allow_html=True,
    )


def sheet(name):
    """Feuille du dessus de la pile de papiers. Un nom par écran : la feuille
    est recréée à chaque changement d'écran, ce qui rejoue son animation d'entrée."""
    return st.container(key=f"stack_{name}").container(key=f"sheet_{name}")


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
    if step != 1:
        st.session_state.intro = False  # l'arrivée animée ne se joue qu'une fois
    st.session_state.step = step
    st.rerun()


def reset_document():
    """Oublie le document en cours (le glossaire et la clé API sont conservés)."""
    for key, value in [
        ("raw_text", ""), ("result_text", ""), ("uploaded_filename", ""),
        ("anon_raw", ""), ("demo_run", False), ("usage_tokens", None),
        ("process_error", ""), ("process_error_fix", False),
    ]:
        st.session_state[key] = value
    st.session_state.pop("anon_widget", None)
    st.session_state.upload_nonce += 1  # vide la zone de dépôt


@st.dialog("Aperçu des notes", width="large")
def dialog_preview():
    st.markdown(
        f'<div class="sg-preview">{html_lib.escape(st.session_state.raw_text)}</div>',
        unsafe_allow_html=True,
    )


@st.dialog("Glossaire", width="large")
def dialog_glossaire():
    st.caption(
        "Un acronyme par ligne, au format `SIGLE : signification`. "
        "Partagé entre tous les documents, conservé d'une session à l'autre."
    )
    raw_glossaire = "\n".join(
        f"{k} : {v}" for k, v in sorted(st.session_state.glossaire.items())
    )
    edited = st.text_area(
        "Glossaire", value=raw_glossaire, height=380, label_visibility="collapsed"
    )
    if st.button("Enregistrer le glossaire", type="primary"):
        new_g = {}
        for line in edited.strip().split("\n"):
            if ":" in line:
                k, v = line.split(":", 1)
                new_g[k.strip()] = v.strip()
        st.session_state.glossaire = new_g
        save_glossaire(new_g)
        st.session_state.flash = "Glossaire enregistré."
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
    ("step", 0),
    ("anon_raw", ""),
    ("flash", ""),
    ("processing", False),
    ("process_error", ""),
    ("process_error_fix", False),
    ("api_key_store", ""),
    ("demo_run", False),
    ("usage_tokens", None),
    ("upload_nonce", 0),
    ("intro", False),
]:
    if key not in st.session_state:
        st.session_state[key] = default

if st.session_state.glossaire is None:
    st.session_state.glossaire = load_glossaire()

# La clé API vit dans un état persistant (le champ de saisie est sur l'écran « Tout est prêt ? »)
api_key = st.session_state.api_key_store.strip()

step = st.session_state.step
render_header(step)

# Nouvel écran → retour en haut de page (Streamlit conserve sinon le défilement).
# Le contenu dépend de l'écran : le script ne se rejoue qu'au changement d'écran.
with st.container(key="scrolltop"):
    st.components.v1.html(
        f"<script>/* écran {step} */ var m = window.parent.document.querySelector('[data-testid=\"stMain\"]');"
        " if (m) m.scrollTo(0, 0); window.parent.scrollTo(0, 0);</script>",
        height=0,
    )

# Message différé (affiché après un rerun)
if st.session_state.flash:
    st.toast(st.session_state.flash)
    st.session_state.flash = ""

# ══════════════════════════════════════════════════════════════════
# Écran 0 — Accueil : dépôt du fichier
# ══════════════════════════════════════════════════════════════════
if step == 0:
    st.markdown('<p class="sg-lead">Dépose tes notes d\'audience :</p>', unsafe_allow_html=True)
    with st.container(key="uploader"):
        uploaded = st.file_uploader(
            "Fichier .docx de notes brutes", type=["docx"],
            label_visibility="collapsed", key=f"upload_{st.session_state.upload_nonce}",
        )
    with sheet("home"):
        st.markdown('<div aria-hidden="true"></div>', unsafe_allow_html=True)

    if uploaded:
        reset_document()
        st.session_state.raw_text = extract_text(uploaded)
        st.session_state.uploaded_filename = uploaded.name
        st.session_state.intro = True  # joue la montée des feuilles à l'arrivée
        goto(1)

# ══════════════════════════════════════════════════════════════════
# Écran 1 — Acronymes
# ══════════════════════════════════════════════════════════════════
elif step == 1:
    unknown = detect_unknown_acronyms(st.session_state.raw_text, st.session_state.glossaire)

    with sheet("acronymes_intro" if st.session_state.intro else "acronymes"):
        with st.container(key="filebar"):
            f_col, a_col, r_col = st.columns([4.4, 1.2, 1.5], vertical_alignment="center")
            with f_col:
                name = html_lib.escape(st.session_state.uploaded_filename)
                st.markdown(
                    f'<div class="sg-file">{ICON_DOC}<span title="{name}">{name}</span></div>',
                    unsafe_allow_html=True,
                )
            with a_col:
                if st.button("Aperçu", key="btn_preview", use_container_width=True):
                    dialog_preview()
            with r_col:
                if st.button("Remplacer", key="btn_replace", use_container_width=True):
                    reset_document()
                    goto(0)

        t_col, g_col, v_col = st.container(key="acro_head").columns(
            [2.6, 2.4, 0.7], vertical_alignment="center"
        )
        with t_col:
            st.markdown("### Acronymes")
        with g_col:
            st.markdown(
                f'<div class="sg-gloss">{ICON_BOOK}<span>Glossaire · '
                f'{len(st.session_state.glossaire)} entrées</span></div>',
                unsafe_allow_html=True,
            )
        with v_col:
            if st.button("Voir", key="btn_gloss", use_container_width=True):
                dialog_glossaire()

        if unknown:
            st.markdown(
                f"<p>{len(unknown)} acronyme(s) non reconnu(s) — renseigne leur signification "
                "(ou laisse vide pour les ignorer), puis enregistre-les.</p>",
                unsafe_allow_html=True,
            )
            cols = st.columns(2, gap="medium")
            for i, acro in enumerate(unknown):
                cols[i % 2].text_input(acro, key=f"acro_{acro}", placeholder="Signification")
            next_label = "Compléter le glossaire"
        else:
            st.markdown(
                chip("✓ Tous les acronymes des notes sont dans le glossaire", "ok"),
                unsafe_allow_html=True,
            )
            next_label = "Continuer"

        if st.button(next_label, type="primary", key="btn_acro_next"):
            entries = {
                acro: (st.session_state.get(f"acro_{acro}") or "").strip() for acro in unknown
            }
            entries = {k: v for k, v in entries.items() if v}
            if entries:
                st.session_state.glossaire.update(entries)
                save_glossaire(st.session_state.glossaire)
                st.session_state.flash = f"{len(entries)} entrée(s) ajoutée(s) au glossaire."
            goto(2)

# ══════════════════════════════════════════════════════════════════
# Écran 2 — Anonymisation
# ══════════════════════════════════════════════════════════════════
elif step == 2:
    with sheet("anonymisation"):
        st.markdown("### Anonymisation")
        st.markdown(
            "<p style='margin-bottom:0'>Crée une ligne par nom à remplacer, un nom fictif pour les "
            "témoins, des initiales pour les avocats. Format : Prénom Nom → Pseudo.</p>"
            "<ul style='margin:0'><li>Les vrais noms ne quittent jamais ton ordinateur.</li>"
            "<li>Laisse vide s'il n'y a pas de nom réel.</li></ul>",
            unsafe_allow_html=True,
        )

        if "anon_widget" not in st.session_state:
            st.session_state.anon_widget = st.session_state.anon_raw
        st.text_area(
            "Noms à anonymiser",
            key="anon_widget",
            placeholder="Malo Rayen → Thomas Bertrand\nL.K. → A.B.",
            height=120,
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
            st.markdown('<p class="sg-hint">Astuce : « -> » fonctionne aussi à la place de « → ».</p>',
                        unsafe_allow_html=True)

        if st.button("Continuer", type="primary", key="btn_anon_next"):
            goto(3)

# ══════════════════════════════════════════════════════════════════
# Écran 3 — Tout est prêt ? (puis traitement)
# ══════════════════════════════════════════════════════════════════
elif step == 3:
    anon_table = parse_anon_table(st.session_state.anon_raw)
    n_anon = len(anon_table)
    n_gloss = len(st.session_state.glossaire)
    n_lines = len(st.session_state.raw_text.split("\n"))

    if st.session_state.processing:
        # ── La feuille devient la carte foncée : rien d'autre à voir ni à faire ──
        with sheet("processing"):
            # Masque les éléments de l'écran précédent pendant l'appel bloquant
            st.markdown(
                '<style>[data-stale="true"] { display: none !important; }</style>'
                + ink_svg() +
                '<div class="sg-processing">'
                '<h3>Mise en forme en cours…</h3>'
                '<p>Récit fidèle aux notes, sans reformulation.<br>'
                '30 à 90 secondes — ne ferme pas cet onglet.</p>'
                '</div>',
                unsafe_allow_html=True,
            )

        # Appel bloquant : l'animation ci-dessus reste affichée pendant ce temps
        error, fix = None, False
        if st.session_state.demo_run:
            time.sleep(6)  # laisser le temps de voir l'animation d'écriture
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
                    + ". Ajoute-les explicitement à la table d'anonymisation."
                )
                fix = True
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
            st.session_state.process_error_fix = fix
            st.rerun()
        else:
            goto(4)

    else:
        with sheet("lancer"):
            st.markdown("### Tout est prêt ?")

            if st.session_state.process_error:
                st.markdown(
                    f'<div class="sg-error">{html_lib.escape(st.session_state.process_error)}</div>',
                    unsafe_allow_html=True,
                )
                if st.session_state.process_error_fix and st.button(
                    "Corriger l'anonymisation →", key="link_fix_anon", type="tertiary"
                ):
                    st.session_state.process_error = ""
                    goto(2)

            rows = [
                ("ok", "Fichier", f"{st.session_state.uploaded_filename} · {n_lines} lignes", None),
                ("ok", "Glossaire", f"{n_gloss} acronymes seront développés", 1),
                (
                    "ok" if n_anon else "ko",
                    "Anonymisation",
                    f"{n_anon} remplacement(s) appliqué(s) avant l'envoi" if n_anon
                    else "Aucun remplacement configuré",
                    2,
                ),
            ]
            for kind, label, detail, target in rows:
                r_col, m_col = st.columns([8, 1], vertical_alignment="center")
                with r_col:
                    st.markdown(
                        f'<div class="sg-recap-row"><span class="ic {kind}">'
                        f'{"✓" if kind == "ok" else "!"}</span>'
                        f'<span class="lb">{label}</span>'
                        f'<span class="dt">{html_lib.escape(detail)}</span></div>',
                        unsafe_allow_html=True,
                    )
                with m_col:
                    if target is not None and st.button(
                        "Modifier", key=f"link_edit_{label}", type="tertiary"
                    ):
                        st.session_state.process_error = ""
                        goto(target)

            st.write("")
            demo = st.checkbox("Mode démo — tester l'outil sans appel API", key="demo_toggle")

            if not demo:
                if "api_key_widget" not in st.session_state:
                    st.session_state.api_key_widget = st.session_state.api_key_store
                api_key = st.session_state.api_key_widget.strip()
                st.markdown(
                    '<div class="sg-keylabel">Clé API Anthropic'
                    + ("" if api_key else '<span class="sg-pill">Requise</span>')
                    + "</div>",
                    unsafe_allow_html=True,
                )
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

            b_col, c_col = st.container(key="launch_row").columns([1, 1.3], vertical_alignment="center")
            with b_col:
                can_run = bool(st.session_state.raw_text and (demo or api_key))
                if st.button("Lancer le traitement", type="primary", disabled=not can_run):
                    st.session_state.demo_run = demo
                    st.session_state.process_error = ""
                    st.session_state.processing = True
                    st.rerun()
            with c_col:
                cost = (
                    "Récit fictif, aucun crédit consommé" if demo
                    else "30 à 90 secondes · ~0,05 € pour 10 pages"
                )
                st.markdown(f'<p class="sg-cost">{cost}</p>', unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════
# Écran 4 — Résultat
# ══════════════════════════════════════════════════════════════════
elif step == 4:
    result = st.session_state.result_text

    if not result:
        goto(3)

    n_uncertain = len(MARKER_RE.findall(result))
    n_notes = result.count("[note suggérée]")

    st.markdown('<p class="sg-done">Ton document est prêt.</p>', unsafe_allow_html=True)
    st.markdown(
        '<p class="sg-note">'
        + ("Les passages incertains resteront surlignés en jaune dans le document."
           if n_uncertain else "Aucun passage incertain : le récit suit fidèlement les notes.")
        + "</p>",
        unsafe_allow_html=True,
    )

    def stat(icon, value, label, kind):
        return (
            f'<div class="sg-stat"><div class="ic {kind}">{icon}</div>'
            f'<div class="v">{value}</div><div class="l">{label}</div></div>'
        )

    n_pages = estimate_pages(result)
    stats = [
        stat(ICON_PAGES, f"≈ {n_pages}", "page(s) dans le document", "neutral"),
        stat(ICON_BINOCULARS, n_uncertain, "passage(s) incertain(s)", "warn")
        if n_uncertain else stat(ICON_CHECK, "0", "passage incertain", "ok")
    ]
    if n_notes:
        stats.append(stat(ICON_COLUMNS, n_notes, "note(s) de bas de page suggérée(s)", "info"))
    st.markdown(f'<div class="sg-stats">{"".join(stats)}</div>', unsafe_allow_html=True)

    out_name = (
        st.session_state.uploaded_filename.replace(".docx", "_récit.docx")
        or "récit_audience.docx"
    )
    with st.container(key="dl"):
        st.download_button(
            label="Télécharger",
            data=generate_docx(result),
            file_name=out_name,
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            type="primary",
            icon=":material/download:",
            use_container_width=True,
        )
    with st.container(key="link_new"):
        if st.button("Traiter un nouveau document", key="btn_new", type="tertiary"):
            reset_document()
            goto(0)

    with sheet("result"):
        st.markdown('<div aria-hidden="true"></div>', unsafe_allow_html=True)

    with st.container(key="energybar"):
        with st.popover("Énergie consommée"):
            render_energy_panel()
