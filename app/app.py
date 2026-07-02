"""
SuperGreffier — MSDE
Outil de traitement des notes d'audience
"""

import streamlit as st
import json
import re
import io
import html as html_lib
from pathlib import Path

from docx import Document
from docx.shared import Pt, RGBColor
import anthropic

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

def apply_anonymization(text, anon_table):
    # Construire une table étendue : ajouter les remplacements nom-seul
    # ex: "Malo Royen → Thomas Bertrand" génère aussi "Royen → Bertrand"
    extended = dict(anon_table)
    for real, pseudo in anon_table.items():
        real_parts = real.strip().split()
        pseudo_parts = pseudo.strip().split()
        if len(real_parts) >= 2 and len(pseudo_parts) >= 2:
            last_real = real_parts[-1]
            last_pseudo = pseudo_parts[-1]
            # N'ajouter que si le nom seul n'est pas déjà dans la table
            if last_real not in extended:
                extended[last_real] = last_pseudo

    # Remplacer du plus long au plus court pour éviter les collisions
    for real in sorted(extended.keys(), key=len, reverse=True):
        pseudo = extended[real]
        if real.strip() and pseudo.strip():
            text = text.replace(real, pseudo)
    return text

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


def apply_uncertain_corrections(text, session_state, n_passages):
    """
    Remplace les occurrences de [passage incertain — à vérifier] par les
    corrections saisies (lues depuis session_state via les clés correction_0…N).
    Les passages sans correction conservent le marqueur.
    """
    marker = "[passage incertain — à vérifier]"
    parts = text.split(marker)
    result = parts[0]
    for i in range(len(parts) - 1):
        correction = session_state.get(f"correction_{i}", "").strip()
        result += correction if correction else marker
        result += parts[i + 1]
    return result


def render_raw_notes_component(raw_text, passage_line_idx, total_processed_lines, window=15):
    """
    Notes brutes avec barre de recherche compacte intégrée (pill sombre).
    Barre : 🔍 [input]  N/M  ↑  ↓  ✕  — tout sur une ligne.
    """
    raw_lines = raw_text.split("\n")
    total_raw = max(1, len(raw_lines))
    ratio = passage_line_idx / max(1, total_processed_lines)
    center = int(ratio * total_raw)
    start = max(0, center - window)
    end = min(total_raw, center + window)

    raw_json = json.dumps(raw_text)
    default_win_json = json.dumps("\n".join(raw_lines[start:end]))
    header_json = json.dumps(f"lignes {start + 1}–{end} sur {total_raw}")

    return f"""
<style>
  #sbar {{display:flex;align-items:center;background:#1e1e2e;border-radius:20px;padding:3px 10px;flex-shrink:0}}
  #sinp {{flex:1;background:transparent;border:none;outline:none;color:#e0e0e0;font-size:0.85em;font-family:monospace;padding:2px 4px}}
  #sinp::placeholder {{color:#555}}
  .nb {{background:none;border:none;color:#888;cursor:pointer;font-size:1.1em;padding:0 3px;line-height:1}}
  .nb:hover {{color:#e0e0e0}}
  #ctr {{color:#aaa;font-size:0.8em;margin:0 6px;min-width:28px;text-align:right}}
</style>
<div style="font-family:monospace;height:320px;display:flex;flex-direction:column;gap:5px">
  <div id="sbar">
    <span style="color:#555;font-size:0.9em;margin-right:5px">🔍</span>
    <input id="sinp" type="text" placeholder="Rechercher dans les notes…" oninput="onS(this.value)" />
    <span id="ctr"></span>
    <button class="nb" onclick="mv(-1)" title="Précédent">↑</button>
    <button class="nb" onclick="mv(1)" title="Suivant">↓</button>
    <button class="nb" onclick="clr()" title="Effacer" style="padding-left:6px">✕</button>
  </div>
  <pre id="box" style="flex:1;overflow-y:auto;background:#f8f9fa;padding:10px;border-radius:4px;border:1px solid #ddd;white-space:pre-wrap;font-size:0.82em;margin:0;color:#1a1a1a"></pre>
</div>
<script>
  var FR={raw_json}, DW={default_win_json}, DH={header_json};
  var occs=[], cur=0;
  var box=document.getElementById('box'), ctr=document.getElementById('ctr');
  function esc(s){{return s.replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;')}}
  box.innerHTML='<span style="color:#888;font-size:0.8em">— '+DH+' —</span>\\n'+esc(DW);
  function updC(){{ctr.textContent=occs.length?(cur+1)+'/'+occs.length:''}}
  function mv(d){{
    if(!occs.length)return;
    occs[cur].style.outline='';
    cur=(cur+d+occs.length)%occs.length;
    occs[cur].style.outline='2px solid #e07000';
    occs[cur].scrollIntoView({{block:'center'}});
    updC();
  }}
  function onS(q){{
    if(!q.trim()){{box.innerHTML='<span style="color:#888;font-size:0.8em">— '+DH+' —</span>\\n'+esc(DW);occs=[];cur=0;ctr.textContent='';return;}}
    var lq=q.toLowerCase(),t=FR,lt=t.toLowerCase(),i=0,h='';
    while(i<t.length){{
      var p=lt.indexOf(lq,i);
      if(p===-1){{h+=esc(t.slice(i));break;}}
      h+=esc(t.slice(i,p))+'<mark class="occ" style="background:#FFE066;color:#1a1a1a">'+esc(t.slice(p,p+q.length))+'</mark>';
      i=p+q.length;
    }}
    box.innerHTML=h;
    occs=Array.from(box.querySelectorAll('.occ'));
    cur=0;
    if(occs.length){{occs[0].style.outline='2px solid #e07000';occs[0].scrollIntoView({{block:'center'}});}}
    updC();
  }}
  function clr(){{document.getElementById('sinp').value='';onS('');}}
</script>
"""


def build_prompt(glossaire):
    lines = [f"- {k} : {v}" for k, v in sorted(glossaire.items())]
    return PROMPT_SYSTEME.format(glossaire="\n".join(lines))

def call_claude(text, system_prompt, api_key):
    client = anthropic.Anthropic(api_key=api_key)
    with client.messages.stream(
        model="claude-sonnet-4-6",
        max_tokens=32000,
        system=system_prompt,
        messages=[{"role": "user", "content": text}],
    ) as stream:
        return stream.get_final_text()

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

# ─────────────────────────────────────────────────────────────────
# Interface
# ─────────────────────────────────────────────────────────────────

st.set_page_config(page_title="SuperGreffier — MSDE", page_icon="⚖️", layout="wide")
st.title("⚖️ SuperGreffier")
st.caption("Outil de traitement des notes d'audience — MSDE")

# Session state
for key, default in [
    ("glossaire", None),
    ("raw_text", ""),
    ("result_text", ""),
    ("uploaded_filename", ""),
    ("uncertain_passages", []),
    ("uncertain_idx", 0),
]:
    if key not in st.session_state:
        st.session_state[key] = default

if st.session_state.glossaire is None:
    st.session_state.glossaire = load_glossaire()

# ── Sidebar ───────────────────────────────────────────────────────
with st.sidebar:
    st.header("Configuration")
    api_key = st.text_input(
        "Clé API Anthropic",
        type="password",
        help="Obtenir une clé sur console.anthropic.com",
    )

    st.divider()
    st.header(f"Glossaire ({len(st.session_state.glossaire)} entrées)")

    with st.expander("Consulter / modifier"):
        raw_glossaire = "\n".join(
            f"{k} : {v}" for k, v in sorted(st.session_state.glossaire.items())
        )
        edited = st.text_area("Format : SIGLE : terme (un par ligne)", value=raw_glossaire, height=300)
        if st.button("💾 Sauvegarder"):
            new_g = {}
            for line in edited.strip().split("\n"):
                if ":" in line:
                    k, v = line.split(":", 1)
                    new_g[k.strip()] = v.strip()
            st.session_state.glossaire = new_g
            save_glossaire(new_g)
            st.success("Glossaire sauvegardé.")

    with st.expander("Ajouter une entrée"):
        c1, c2 = st.columns(2)
        new_k = c1.text_input("Sigle")
        new_v = c2.text_input("Terme")
        if st.button("Ajouter") and new_k and new_v:
            st.session_state.glossaire[new_k.strip()] = new_v.strip()
            save_glossaire(st.session_state.glossaire)
            st.success(f"'{new_k}' ajouté.")

# ── Onglets ───────────────────────────────────────────────────────
tab_main, tab_help = st.tabs(["Traiter une note", "Aide"])

with tab_main:

    # ── 1. Upload ─────────────────────────────────────────────────
    st.subheader("1 · Charger le fichier")
    uploaded = st.file_uploader("Fichier .docx de notes brutes", type=["docx"])

    if uploaded:
        if uploaded.name != st.session_state.uploaded_filename:
            # Nouveau fichier → reset complet
            st.session_state.raw_text = extract_text(uploaded)
            st.session_state.uploaded_filename = uploaded.name
            st.session_state.result_text = ""
            st.session_state.uncertain_passages = []
            st.session_state.uncertain_idx = 0
        elif not st.session_state.raw_text:
            # Même fichier mais texte pas encore extrait
            st.session_state.raw_text = extract_text(uploaded)
        with st.expander("Aperçu des notes brutes"):
            preview = st.session_state.raw_text
            st.text(preview[:3000] + ("…" if len(preview) > 3000 else ""))

    st.divider()

    # ── 2. Acronymes inconnus ─────────────────────────────────────
    st.subheader("2 · Acronymes inconnus")

    if st.session_state.raw_text:
        unknown = detect_unknown_acronyms(
            st.session_state.raw_text, st.session_state.glossaire
        )
        if not unknown:
            st.success("Tous les acronymes détectés sont dans le glossaire.")
        else:
            st.info(
                f"{len(unknown)} acronyme(s) non reconnu(s). "
                "Renseigne leur signification ou laisse vide pour ignorer."
            )
            new_entries = {}
            cols = st.columns(3)
            for i, acro in enumerate(unknown):
                term = cols[i % 3].text_input(acro, key=f"acro_{acro}", placeholder="Signification")
                if term.strip():
                    new_entries[acro] = term.strip()

            if st.button("Valider les acronymes") and new_entries:
                st.session_state.glossaire.update(new_entries)
                save_glossaire(st.session_state.glossaire)
                st.success(f"{len(new_entries)} entrée(s) ajoutée(s) au glossaire.")
    else:
        st.caption("Charge un fichier à l'étape 1.")

    st.divider()

    # ── 3. Anonymisation ──────────────────────────────────────────
    st.subheader("3 · Anonymisation")
    st.caption(
        "Liste les noms à remplacer, un par ligne. "
        "Format : `Vrai Nom → Pseudonyme`  |  Avocats : initiales (`L.K. → A.B.`)  |  "
        "Prévenu·es : prénom + nom fictif à consonance proche."
    )
    anon_input = st.text_area(
        "Noms à anonymiser",
        placeholder="Malo Royen → Thomas Bertrand\nJessica Remaillon → Sophie Duval\nL.K. → A.B.",
        height=140,
    )
    anon_table = parse_anon_table(anon_input)
    if anon_table:
        st.caption(f"{len(anon_table)} remplacement(s) configuré(s).")

    st.divider()

    # ── 4. Lancement ──────────────────────────────────────────────
    st.subheader("4 · Traitement")

    can_run = bool(st.session_state.raw_text and api_key)
    if not api_key:
        st.warning("Renseigne ta clé API dans la barre latérale.")
    if not st.session_state.raw_text:
        st.warning("Charge un fichier à l'étape 1.")

    if st.button("⚙️ Lancer le traitement", disabled=not can_run, type="primary"):
        with st.spinner("Traitement en cours… (30 à 90 secondes)"):
            # Reset révision passages incertains
            st.session_state.uncertain_passages = []
            st.session_state.uncertain_idx = 0
            for k in list(st.session_state.keys()):
                if k.startswith("correction_") or k.startswith("search_"):
                    del st.session_state[k]

            anonymized = apply_anonymization(st.session_state.raw_text, anon_table)
            system_prompt = build_prompt(st.session_state.glossaire)
            try:
                st.session_state.result_text = call_claude(anonymized, system_prompt, api_key)
            except anthropic.AuthenticationError:
                st.error("Clé API invalide. Vérifie la clé dans la barre latérale.")
            except anthropic.APIConnectionError:
                st.error("Impossible de contacter l'API. Vérifie ta connexion internet.")
            except Exception as e:
                st.error(f"Erreur inattendue : {e}")

    st.divider()

    # ── 5. Résultat ───────────────────────────────────────────────
    st.subheader("5 · Résultat")

    if st.session_state.result_text:
        n_uncertain = st.session_state.result_text.count("[passage incertain")
        n_notes = st.session_state.result_text.count("[note suggérée]")

        c1, c2 = st.columns(2)
        c1.metric("Passages incertains", n_uncertain)
        c2.metric("Notes suggérées", n_notes)

        if n_uncertain:
            st.warning(
                f"{n_uncertain} passage(s) marqué(s) [passage incertain — à vérifier]. "
                "Ils sont soulignés en orange dans le .docx."
            )

        st.text_area("Aperçu du récit", value=st.session_state.result_text, height=450)

        docx_buf = generate_docx(st.session_state.result_text)
        out_name = st.session_state.uploaded_filename.replace(".docx", "_récit.docx") or "récit_audience.docx"

        st.download_button(
            label="📥 Télécharger le récit (.docx)",
            data=docx_buf,
            file_name=out_name,
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            type="primary",
        )

    # ── 6. Révision passages incertains ──────────────────────────
    n_uncertain = st.session_state.result_text.count("[passage incertain") if st.session_state.result_text else 0
    if n_uncertain > 0:
        st.divider()
        st.subheader("6 · Révision des passages incertains")

        # Initialiser la liste si nécessaire
        if not st.session_state.uncertain_passages:
            st.session_state.uncertain_passages = extract_uncertain_passages(
                st.session_state.result_text
            )
            st.session_state.uncertain_idx = 0

        passages = st.session_state.uncertain_passages
        n = len(passages)
        idx = st.session_state.uncertain_idx

        # ── Navigation ────────────────────────────────────────────
        nav_c1, nav_c2, nav_c3 = st.columns([1, 3, 1])
        with nav_c1:
            if st.button("← Précédent", disabled=(idx == 0), key="btn_prev"):
                st.session_state.uncertain_idx -= 1
                st.rerun()
        with nav_c2:
            st.markdown(
                f"<div style='text-align:center;font-weight:600;padding-top:6px'>"
                f"Passage {idx + 1} sur {n}</div>",
                unsafe_allow_html=True,
            )
        with nav_c3:
            if st.button("Suivant →", disabled=(idx == n - 1), key="btn_next"):
                st.session_state.uncertain_idx += 1
                st.rerun()

        passage = passages[idx]

        # ── Triptych ──────────────────────────────────────────────
        col1, col2, col3 = st.columns([2, 2, 1.5])

        with col1:
            st.markdown("**Contexte — texte traité**")
            parts = []
            if passage["context_before"]:
                parts.append(
                    f'<span style="color:#888;font-size:0.85em">{html_lib.escape(passage["context_before"])}</span>'
                )
            highlighted_line = html_lib.escape(passage["line"]).replace(
                html_lib.escape("[passage incertain — à vérifier]"),
                '<mark style="background:#FFB347;font-weight:600">[passage incertain — à vérifier]</mark>',
            )
            parts.append(highlighted_line)
            if passage["context_after"]:
                parts.append(
                    f'<span style="color:#888;font-size:0.85em">{html_lib.escape(passage["context_after"])}</span>'
                )
            st.markdown(
                '<div style="font-size:0.88em;line-height:1.6;padding:10px;'
                'background:#fafafa;border-radius:4px;border:1px solid #eee;color:#1a1a1a">'
                + "<br><br>".join(parts) + "</div>",
                unsafe_allow_html=True,
            )

        with col2:
            st.markdown("**Notes brutes originales**")
            total_processed_lines = len(st.session_state.result_text.split("\n"))
            st.components.v1.html(
                render_raw_notes_component(
                    st.session_state.raw_text,
                    passage["line_idx"],
                    total_processed_lines,
                ),
                height=345,
            )

        with col3:
            st.markdown("**Correction**")
            st.text_area(
                "Remplacer le marqueur par :",
                key=f"correction_{idx}",
                height=115,
                placeholder="Laisser vide pour conserver [passage incertain]",
                label_visibility="collapsed",
            )
            if st.button("✅ Valider", key=f"btn_valider_{idx}", use_container_width=True,
                         help="Valider et passer au suivant (Ctrl+Entrée)"):
                if idx < n - 1:
                    st.session_state.uncertain_idx += 1
                    st.rerun()
                else:
                    corrected = apply_uncertain_corrections(
                        st.session_state.result_text, st.session_state, n
                    )
                    st.session_state.result_text = corrected
                    st.session_state.uncertain_passages = []
                    st.session_state.uncertain_idx = 0
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
                if (btns[i].innerText.trim() === '✅ Valider') { btns[i].click(); return; }
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

        st.divider()
        if st.button("✅ Appliquer toutes les corrections", type="primary"):
            corrected = apply_uncertain_corrections(
                st.session_state.result_text, st.session_state, n
            )
            st.session_state.result_text = corrected
            st.session_state.uncertain_passages = []
            st.session_state.uncertain_idx = 0
            st.rerun()

    if st.session_state.result_text:
        st.divider()
        if st.button("🔄 Traiter un nouveau document"):
            st.session_state.raw_text = ""
            st.session_state.result_text = ""
            st.session_state.uploaded_filename = ""
            st.rerun()
    else:
        st.caption("Le récit apparaîtra ici après le traitement.")

with tab_help:
    st.markdown("""
## Comment utiliser SuperGreffier

### Prérequis

Une clé API Anthropic — à obtenir sur [console.anthropic.com](https://console.anthropic.com).
Coût estimé : **~0,05 € par document de 10 pages**.

### Étapes

| Étape | Action |
|-------|--------|
| **1** | Charge ton fichier `.docx` de notes brutes |
| **2** | L'outil détecte les acronymes absents du glossaire. Renseigne leur signification ou laisse vide pour ignorer |
| **3** | Liste les noms à anonymiser au format `Vrai Nom → Pseudonyme` |
| **4** | Clique sur **Lancer le traitement** |
| **5** | Relis l'aperçu, télécharge le `.docx` |

### Marqueurs dans le document de sortie

- `[passage incertain — à vérifier]` — passage trop dégradé pour être interprété avec certitude. **À corriger manuellement.** Affiché en orange souligné dans le .docx.
- `[note suggérée]` — passage qui pourrait bénéficier d'une note de bas de page (référence juridique, personnalité, événement).

### Anonymisation

- **Avocats** : remplacer par les initiales (`L.K.`)
- **Prévenu·es** : choisir un prénom et nom fictif à consonance proche du nom original
- **Adresses** : remplacer manuellement par la région dans les notes avant d'uploader si besoin

### Confidentialité

Seul le texte **après anonymisation** est envoyé à Anthropic. Les vrais noms ne quittent jamais ton poste.

### Glossaire

Partagé entre tous les documents. Sauvegardé automatiquement à chaque ajout. Visible et modifiable dans la barre latérale.
    """)
