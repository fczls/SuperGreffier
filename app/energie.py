"""
Estimation de l'énergie (et du CO₂) consommée par un traitement, à partir des
tokens facturés par l'API Claude.

⚠ Ordre de grandeur, pas une mesure : Anthropic ne publie aucune donnée
d'énergie par requête ni par token. Les valeurs ci-dessous sont dérivées de
chiffres par requête publiés ailleurs (voir SOURCES) et d'hypothèses explicites.
On présente donc toujours une fourchette basse / haute.
"""

# Énergie par token de SORTIE, en Wh.
#  - bas  : 0,24 Wh / ~500 tokens ≈ 0,5 mWh (médiane Gemini publiée par Google, 2025 ;
#           ordre de grandeur voisin de 0,34 Wh/requête annoncé par OpenAI).
#  - haut : ~3 Wh / ~500 tokens ≈ 6 mWh (estimation indirecte citée par Bon Pote,
#           et estimations tierces pour Claude Sonnet : 0,8 à 5,5 Wh selon la longueur).
WH_PAR_TOKEN_SORTIE = {"bas": 0.0005, "haut": 0.006}

# Hypothèse : un token d'ENTRÉE (préremplissage, parallélisable) coûte ~10× moins
# qu'un token de sortie (génération séquentielle).
RATIO_ENTREE_SORTIE = 0.1

# Intensité carbone de l'électricité, en gCO₂e/kWh (selon la localisation inconnue du
# datacenter) : ordre de grandeur réseau français vs réseau américain (rapport ≥ 10,
# cf. Bon Pote).
G_CO2_PAR_KWH = {"bas": 55, "haut": 400}

# Émissions d'un ordinateur allumé, en gCO₂e par heure (40 à 80 g, John Buckley, cité par
# Ekwateur sans lien vers la source) : sert à exprimer le résultat en « temps d'ordinateur allumé ».
G_CO2_PC_PAR_HEURE = {"bas": 40, "haut": 80}

# Valeurs affichées en mode démo (gCO₂e, bas / haut), sans appel API.
DEMO_G_CO2 = (0.32, 28.0)

SOURCES = [
    ("Google, médiane d'un prompt Gemini (0,24 Wh), via Sustainability by Numbers",
     "https://www.sustainabilitybynumbers.com/p/ai-footprint-august-2025"),
    ("Bon Pote — Intelligence artificielle : le vrai coût environnemental",
     "https://bonpote.com/intelligence-artificielle-le-vrai-cout-environnemental-de-la-course-a-lia/"),
    ("Ekwateur — Empreinte carbone d'une recherche Google (ordinateur allumé : 40 à 80 g/h)",
     "https://ekwateur.fr/blog/enjeux-environnementaux/empreinte-carbone-recherche-google/"),
    ("Simon Couch — Électricité des agents de code",
     "https://simonpcouch.com/blog/2026-01-20-cc-impact/"),
]


def estimate_energy(input_tokens, output_tokens):
    """Retourne {'wh': (bas, haut), 'g_co2': (bas, haut)}."""
    eq_tokens = max(output_tokens, 0) + RATIO_ENTREE_SORTIE * max(input_tokens, 0)
    wh_bas = eq_tokens * WH_PAR_TOKEN_SORTIE["bas"]
    wh_haut = eq_tokens * WH_PAR_TOKEN_SORTIE["haut"]
    return {
        "wh": (wh_bas, wh_haut),
        "g_co2": (
            wh_bas / 1000 * G_CO2_PAR_KWH["bas"],
            wh_haut / 1000 * G_CO2_PAR_KWH["haut"],
        ),
    }


def format_range(low, high, unit):
    def fmt(x):
        return f"{x:.0f}" if x >= 10 else (f"{x:.1f}" if x >= 1 else f"{x:.2f}")
    return f"{fmt(low)} à {fmt(high)} {unit}"


def pc_minutes(g_low, g_high):
    """Durée d'ordinateur allumé (minutes) émettant autant de CO₂e : bornes les plus prudentes."""
    return (
        g_low / G_CO2_PC_PAR_HEURE["haut"] * 60,
        g_high / G_CO2_PC_PAR_HEURE["bas"] * 60,
    )


def format_duration(minutes):
    if minutes < 1:
        return f"{minutes * 60:.0f} s"
    if minutes < 120:
        return f"{minutes:.0f} min"
    return f"{minutes / 60:.1f} h"


def minutes_evitees(heures_manuelles, g_haut):
    """Minutes d'ordinateur allumé évitées au minimum : temps manuel moins le pire cas de l'outil."""
    _, pire = pc_minutes(0, g_haut)
    return max(heures_manuelles * 60 - pire, 0)


def format_hm(minutes):
    h, m = divmod(round(minutes), 60)
    return f"{h} h {m:02d}" if h else f"{m} min"
