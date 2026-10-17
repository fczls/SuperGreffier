"""
Anonymisation locale des notes d'audience (avant tout envoi à l'API).

Règles :
- remplacement en une seule passe (pas de chaînage A→B puis B→C) ;
- mots entiers uniquement (« Royen » ne touche pas « Royenne ») ;
- insensible à la casse (ROYEN, royen) ;
- prénom seul et nom seul déduits des entrées « Prénom Nom → Prénom' Nom' ».
"""

import re

MIN_TOKEN_LEN = 3


def _derive_entries(anon_table):
    """Table étendue : entrées explicites + prénom/nom seuls déduits."""
    explicit = {k.strip(): v.strip() for k, v in anon_table.items() if k.strip() and v.strip()}
    explicit_keys = {k.casefold() for k in explicit}

    candidates = {}  # casefold(token) -> set de (token, pseudo_token)
    for real, pseudo in explicit.items():
        rp, pp = real.split(), pseudo.split()
        if len(rp) < 2 or len(pp) < 2:
            continue
        for r_tok, p_tok in ((rp[0], pp[0]), (rp[-1], pp[-1])):
            if len(r_tok) >= MIN_TOKEN_LEN and r_tok.casefold() not in explicit_keys:
                candidates.setdefault(r_tok.casefold(), set()).add((r_tok, p_tok))

    extended = dict(explicit)
    for variants in candidates.values():
        # Deux personnes partageant ce mot avec des pseudos différents : ambigu, on ne devine pas
        if len({p for _, p in variants}) == 1:
            r_tok, p_tok = sorted(variants)[0]
            extended[r_tok] = p_tok
    return extended


def _match_case(matched, pseudo):
    if len(matched) > 1 and matched.isupper() and not pseudo.isupper():
        return pseudo.upper()
    return pseudo


def apply_anonymization(text, anon_table):
    extended = _derive_entries(anon_table)
    if not extended:
        return text

    lookup = {k.casefold(): v for k, v in extended.items()}
    keys = sorted(extended, key=len, reverse=True)
    pattern = re.compile(
        r"(?<!\w)(?:" + "|".join(re.escape(k) for k in keys) + r")(?!\w)",
        re.IGNORECASE,
    )
    return pattern.sub(lambda m: _match_case(m.group(0), lookup[m.group(0).casefold()]), text)


def find_residual_names(anonymized_text, anon_table):
    """Vrais noms (entrées complètes ou mots isolés) encore présents après anonymisation."""
    tokens = set()
    for real in anon_table:
        real = real.strip()
        if not real:
            continue
        tokens.add(real)
        parts = real.split()
        if len(parts) >= 2:
            tokens.update(p for p in parts if len(p) >= MIN_TOKEN_LEN)
    residual = []
    for tok in sorted(tokens):
        if re.search(r"(?<!\w)" + re.escape(tok) + r"(?!\w)", anonymized_text, re.IGNORECASE):
            residual.append(tok)
    return residual
