# PRD — Outil de traitement des notes d'audience MSDE

**Version** 0.3  
**Date** 11 juin 2026  
**Auteur** Fabien Cazals  
**Statut** En révision

---

## Problème

Les bénévoles du MSDE assistent aux audiences judiciaires et prennent des notes en temps réel. Ces notes sont rédigées dans l'urgence : orthographe phonétique, acronymes non développés, phrases incomplètes, absence totale de mise en page. La qualité se dégrade au fil de l'audience — les déclarations des prévenus, prises tôt, restent lisibles ; les plaidoiries et réquisitions, prises tard dans une fatigue croissante, deviennent quasi-inchiffrables.

Transformer ces notes en un récit d'audience conforme à la charte rédactionnelle du MSDE demande aujourd'hui environ une journée de travail. Ce délai ralentit la communication vers les avocats partenaires et les militants, réduit la réactivité du MSDE et épuise les bénévoles sur une tâche de reformatage plutôt que d'analyse.

---

## Objectifs

1. Réduire le temps de production d'un récit d'audience de ~8 heures à moins de 30 minutes.
2. Produire un document structuré, lisible et **factuellement fidèle**, conforme à la charte rédactionnelle du MSDE, utilisable sans relecture lourde dans au moins 80 % des cas.
3. Générer un récit unique destiné aux avocats partenaires, lisible également par les militants — sans reformulation ni adaptation de fond.
4. Anonymiser systématiquement les données personnelles selon les règles MSDE avant tout envoi à l'API.
5. Maintenir une bibliothèque d'acronymes partagée, visible et modifiable par les utilisateurs dès la v1.
6. Rendre l'outil accessible à des bénévoles non-techniques, sans formation requise.

---

## Hors périmètre (v1)

- **Transcription audio** : l'outil ne traite que du texte déjà saisi.
- **Mise en page livret finale** : la mise en page livret (format A4 plié, styles Word, couverture InDesign/Canva) reste manuelle. L'outil produit un `.docx` propre et structuré que le relecteur met ensuite en page dans le template Word MSDE.
- **Notes de bas de page** : l'outil peut signaler les passages qui nécessitent une note, mais la rédaction et l'insertion des notes de bas de page restent manuelles en v1.
- **Intégration avec un logiciel de gestion de dossiers** : export manuel uniquement.
- **Support multilingue** : français uniquement.
- **Historique des récits traités** : pas de base de données d'archives. Seule la bibliothèque d'acronymes est persistante.
- **Collaboration en temps réel** : un seul utilisateur à la fois.

---

## Principe directeur — fidélité absolue au contenu

**L'outil n'est pas un outil de rédaction. C'est un outil de mise en forme et de décodage.**

La règle fondamentale : **chaque phrase en entrée doit rester, dans sa teneur, la phrase de sortie.** L'outil est autorisé à :

- Corriger l'orthographe, la typographie et la ponctuation.
- Compléter des mots tronqués ou phonétiques lorsque le sens est non ambigu.
- Remplacer un acronyme par sa forme développée (selon la bibliothèque active).
- Attribuer une prise de parole à son locuteur lorsque ce dernier est identifiable.
- Appliquer les règles typographiques et rédactionnelles de la charte MSDE (majuscules, nombres, écriture inclusive dans les parties narratives).
- Restructurer la mise en page (sections, titres, paragraphes, styles).

L'outil **n'est pas autorisé à** :

- Reformuler une phrase pour la rendre plus élégante.
- Ajouter des informations absentes des notes.
- Supprimer des passages, même redondants.
- Interpréter un passage ambigu en choisissant une formulation définitive — ces passages sont marqués `[passage incertain — à vérifier]`.

---

## Charte rédactionnelle MSDE — règles à appliquer

Cette section décrit les règles issues de la *Charte rédaction harmonisée* du MSDE que l'outil doit appliquer automatiquement.

### Typographie

**Abréviations**
L'outil n'introduit jamais d'abréviations (Mme, M., etc.) qui ne figurent pas dans les notes. On transcrit un langage parlé : seules les abréviations prononcées sont conservées. Le terme « etc » est remplacé par « *et cætera* » (en italique, nom latin).

**Majuscules — institutions**
- Institutions uniques en France : majuscule — *la Cour de cassation, le Conseil d'État.*
- Institutions locales ou génériques : minuscule — *la cour d'appel de Paris, le tribunal judiciaire de Lyon.*
- Quand l'institution désigne une assemblée de personnes : majuscule — *le Tribunal doit statuer.*

**Majuscules — titres honorifiques**
Par choix politique du MSDE, les titres honorifiques ne prennent pas de majuscule : *madame la procureure, monsieur le président, le gouvernement.*
Exception : *le ministre de l'Intérieur* (majuscule au domaine, pas au titre).

**Majuscules — Madame / Monsieur**
- Minuscule quand on parle de la personne à la troisième personne : *je ne vois pas madame Vani.*
- Majuscule quand on s'adresse directement à la personne : *Comment avez-vous trouvé ce livre, Madame Lapi ?*

**Majuscules — acronymes développés**
Premier mot en majuscule, reste en minuscule, sauf noms propres inclus.
- ONU → Organisation des nations unies
- CEDH → Cour européenne des droits de l'homme (noms propres = majuscule)
- TotalEnergies → conservé tel quel (nom d'entreprise)

**Majuscules — éléments fixes**
- *État* et *République* toujours avec majuscule.
- Les codes juridiques sans majuscule : *le code pénal, le code de procédure pénale.*

**Nombres et chiffres**
Règle générale : en toutes lettres.
Exceptions (en chiffres) :
- Dates : *10 décembre 2026.*
- Nombres longs ou peu lisibles : *2 956 €* (espace après les milliers).
- Citations de données scientifiques ou accumulation de chiffres.
- Cas de lisibilité évidente : *18 km/h* plutôt que *dix-huit kilomètres par heure.*

**Unités de mesure**
- Nombre en lettres → unité en lettres : *six litres.*
- Nombre en chiffres → unité en symbole : *800 €, 1 250 L* (jamais *500 euros*).

**Divers**
- CO₂ : avec le 2 en indice (lettre O, pas le chiffre 0).
- Séparateurs de section : ***** (cinq étoiles), police Libre Baskerville, centré.

### Écriture inclusive

**Citations directes** : pas d'écriture inclusive. On rapporte les paroles telles que prononcées, généralement avec le masculin pluriel.

**Narration et texte rédigé par le MSDE** : écriture inclusive obligatoire, selon les principes suivants :
- Privilégier les tournures neutres accessibles : *ils et elles* plutôt que *iels.*
- Utiliser des noms et adjectifs épicènes quand c'est possible pour éviter le point médian.
- Point médian utilisé en dernier recours, sous la forme : *prévenu·es, auteur·ices.*
- Éviter les formes coupées abusives : préférer *agriculteurs et agricultrices* à *agriculteur·ices.*

### Attribution des prises de parole

Format : **Nom ou titre en gras**, suivi de deux points, puis le texte de la prise de parole.

Formes attendues :
- **La présidente :** ou **Le président :**
- **Le procureur :** ou **La procureure :**
- **L'avocat de la défense A.B. :** (initiales Prénom.Nom)
- **L'avocate de la partie civile F.B. :**
- **Prénomanonymisé Nomanonymisé :** (pour les prévenu·es — voir section Anonymisation)
- **La témoin :** (quand un·e témoin prend la parole — non identifié·e par son nom)

### Notes de bas de page

Les notes de bas de page servent à contextualiser pour un lecteur sans culture juridique ou politique (référence à un autre procès, à une personnalité, à un concept juridique, à un événement politique). Elles ont leur propre format :

- Textes juridiques : *Art. xxx-xx Texte législatif.* (ex : *Art. 322-4-1 C. pénal*)
- Lois : *Type de texte n° numéro du Date Titre (Source Date).*
- Jurisprudence Cour de cassation : *Crim., date, n°xxxx.*

Un dictionnaire des notes déjà rédigées est disponible à [msde.frama.space/f/26191](https://msde.frama.space/f/26191).

**En v1, l'outil ne génère pas de notes de bas de page.** Il peut signaler les passages qui semblent appeler une note (référence à une jurisprudence, à une personnalité, à un événement) avec le marqueur `[note suggérée]`. La rédaction et l'insertion restent manuelles.

---

## Anonymisation — règles précises

L'anonymisation est **systématique et automatique**, appliquée avant tout envoi à l'API. Elle suit les règles spécifiques du MSDE, plus nuancées qu'une simple substitution par identifiant neutre.

### Ce qui n'est pas anonymisé

- Noms des entreprises parties civiles.
- Noms des personnalités politiques.
- Noms des magistrats, procureurs, greffiers (acteurs institutionnels publics).

### Règles par type d'acteur

**Avocat·es**
Remplacé par ses initiales : *L'avocate de la défense A.B.* ou *L'avocat de la partie civile F.B.*
Initiale du prénom en premier, initiale du nom de famille en second.

**Prévenu·es**
Remplacé par un **faux prénom et nom à consonance proche de l'original**. Ce pseudonyme est cohérent tout au long du document. Lors des prises de parole, on note le pseudonyme complet : *Amandine Crochet : je ne voulais pas me laisser faire.*

**Témoins**
Même règle que les prévenu·es pour les références dans le texte. Lors des prises de parole, on note : *La témoin :* (sans nom).

**Autres informations identifiantes**
- Adresse : remplacée par la région ou grande zone géographique (*Île-de-France*, *région lyonnaise*).
- Métier trop précis : remplacé par une description générique (*pour la mairie d'une grande ville* au lieu de *mairie de Paris*).
- Ville d'habitation : supprimée.

### Mécanisme technique

**Étape 1 — Détection et proposition de pseudonymes (avant envoi)**

L'outil détecte les noms propres des prévenu·es, témoins et avocat·es dans le document. Pour chaque personne identifiée, il propose :
- Pour les prévenu·es et témoins : un pseudonyme à consonance proche (suggestion automatique).
- Pour les avocat·es : les initiales extraites automatiquement.

Ces propositions sont présentées à l'utilisateur dans une interface de validation avant envoi. L'utilisateur peut **accepter, modifier ou remplacer** chaque pseudonyme proposé — ce qui lui permet de calibrer lui-même le niveau d'anonymat et de choisir des noms qui lui semblent appropriés.

Il peut également ajuster manuellement les remplacements d'adresses et d'informations professionnelles.

**Étape 2 — Envoi à l'API**

Le texte envoyé à l'API contient uniquement les pseudonymes validés. Les vrais noms ne quittent jamais le poste local. La table de correspondance original → pseudonyme est stockée localement et n'est jamais transmise.

**Étape 3 — Document final**

Les pseudonymes validés *sont* les noms du document final. Les vrais noms ne réapparaissent pas dans le récit publié.

---

## Gestion des acronymes — variabilité inter-scribes

Chaque scribe a ses propres conventions. Un même terme peut être noté différemment selon les auteurs.

### Pipeline en deux étapes

**Étape 1 — Résolution avant traitement**

1. L'outil analyse le document et extrait tous les tokens ressemblant à des acronymes absents de la bibliothèque partagée.
2. Ces tokens sont présentés au scribe : *« RC » — quel terme désigne cet acronyme dans ce document ?*
3. Le scribe renseigne la définition ou indique que c'est déjà dans la bibliothèque sous une autre notation.
4. Les nouvelles définitions sont proposées pour enrichissement de la bibliothèque partagée.

**Étape 2 — Traitement par l'API**

Le document est envoyé avec la bibliothèque complète (partagée + résolutions de session) intégrée dans le prompt système.

### Bibliothèque d'acronymes partagée

Stockée localement en JSON, accessible depuis l'interface sous forme de tableau éditable (colonnes : Acronyme / Terme développé / Actions). Persistante entre les sessions. Enrichissable après chaque traitement.

Lexique de base (issu du document exemple) :

| Acronyme | Terme développé |
|---|---|
| GAV | garde à vue |
| OP | ordonnance pénale |
| FDO | forces de l'ordre |
| DR | Dernière Rénovation |
| B2 | bulletin n°2 du casier judiciaire |
| CC° | Cour de cassation |
| CEDH | Cour européenne des droits de l'homme |
| EN | état de nécessité |
| LE | liberté d'expression |
| RC | réchauffement climatique |
| AN | Assemblée nationale |
| LDB | lanceur de balles de défense |
| APP | audience de présentation préalable |
| ADEME | Agence de l'environnement et de la maîtrise de l'énergie |
| GIEC | Groupe d'experts intergouvernemental sur l'évolution du climat |

---

## Structure du récit produit

L'outil produit un `.docx` structuré en sections fixes, prêt à être copié-collé dans le template Word MSDE pour la mise en page livret finale.

### Sections

1. **Résumé d'audience** (tableau) : tribunal, date, affaire, composition du tribunal, prévenu·es, chef d'inculpation, décision ou date de délibéré.
2. **Avant l'audience** : contexte, observations du scribe avant l'entrée de la juge.
3. **L'audience**
   - Ouverture et vérification des identités
   - Déclarations des prévenu·es (une sous-section par personne)
   - Questions du tribunal et du parquet
   - Réquisitions du parquet
   - Plaidoiries de la défense
4. **La décision**

### Styles de paragraphes dans le `.docx`

Le document de sortie applique 4 styles distincts correspondant aux calques du template Word MSDE :

| Style | Usage | Rendu |
|---|---|---|
| `Titre-récit` | Titres de section (Avant l'audience, L'audience…) | Titre bleu |
| `Narration` | Paragraphes descriptifs rédigés par le scribe | Italique |
| `Dialogue-entrée` | Première ligne d'une prise de parole (**Nom :** texte) | Normal, nom en gras |
| `Dialogue-suite` | Suite de la prise de parole si saut de ligne dans le même discours | Normal |

Les notes de bas de page (style `NoteBasDePage`) sont hors scope v1.

---

## Utilisateurs cibles

**Scribe (utilisateur principal)**
Présent·e en salle d'audience, prend des notes en temps réel. Niveau technique variable. Souhaite produire un document propre rapidement après l'audience. Peut avoir ses propres conventions d'abréviation.

**Coordinateur·rice (utilisateur secondaire)**
Relit, valide et diffuse les récits. Gère la bibliothèque d'acronymes partagée. Peut lancer le traitement si le scribe n'a pas pu le faire.

**Avocat·e partenaire (consommateur)**
Reçoit le récit finalisé. A besoin de précision factuelle et d'une structure claire des échanges juridiques.

**Militant·e / réseau (consommateur)**
Lit le même récit. Profite du développement des acronymes et de la narration en écriture inclusive pour suivre sans culture juridique.

---

## Parcours utilisateur

1. Le scribe ouvre l'application web depuis son navigateur (en local).
2. Il charge son fichier `.docx` de notes brutes.
3. **Résolution des acronymes** : l'outil détecte les acronymes inconnus et les soumet au scribe. Le scribe peut aussi consulter et modifier la bibliothèque à cette étape.
4. Le scribe confirme et clique sur « Traiter ».
5. L'outil anonymise automatiquement (génération des pseudonymes, remplacement des informations identifiantes).
6. L'outil envoie le texte anonymisé à l'API Claude avec la bibliothèque et les règles de la charte.
7. L'API produit le récit structuré.
8. Le résultat s'affiche dans l'interface pour relecture rapide, avec les passages incertains et les suggestions de notes en évidence.
9. Le scribe télécharge le `.docx` finalisé.
10. L'outil propose d'enrichir la bibliothèque avec les nouveaux acronymes résolus.

---

## Exigences

### Must-have (P0)

**Résolution interactive des acronymes inconnus**
- Détection des tokens inconnus avant traitement.
- Interface de saisie des définitions manquantes.
- Critère d'acceptation : aucun acronyme non résolu n'est envoyé à l'API sans information de l'utilisateur.

**Bibliothèque d'acronymes persistante et éditable**
- Tableau éditable dans l'interface (ajout, modification, suppression).
- Persistance locale (JSON) entre les sessions.
- Proposition d'enrichissement après chaque traitement.
- Critère d'acceptation : les modifications persistent entre deux sessions sans manipulation de fichier.

**Anonymisation avec validation utilisateur**
- L'outil détecte les noms propres et propose automatiquement des remplacements (initiales pour les avocats, pseudonymes à consonance proche pour les prévenu·es et témoins).
- L'utilisateur valide, modifie ou remplace chaque proposition avant l'envoi — il garde le contrôle total sur les pseudonymes choisis.
- Adresses → zone géographique large (ex. *région lyonnaise*), proposée et modifiable.
- Informations professionnelles précises → description générique, proposée et modifiable.
- Critère d'acceptation : aucune donnée identifiante réelle ne figure dans le texte envoyé à l'API, vérifiable via log debug. L'utilisateur ne peut pas déclencher le traitement sans avoir validé l'écran d'anonymisation.

**Application de la charte rédactionnelle**
Le prompt système encode les règles suivantes et le modèle les applique :
- Majuscules : institutions, titres, Madame/Monsieur, acronymes développés.
- Nombres en toutes lettres (avec exceptions).
- Remplacement de « etc » par « *et cætera* ».
- Écriture inclusive dans les parties narratives, pas dans les citations.
- Attribution des prises de parole au bon format (**Nom :** texte).
- Critère d'acceptation : sur le document exemple, toutes les règles listées ci-dessus sont appliquées correctement sur un échantillon de 20 paragraphes.

**Fidélité et signalement des incertitudes**
- Passages trop dégradés → marqueur `[passage incertain — à vérifier]`.
- Passages pouvant nécessiter une note de bas de page → marqueur `[note suggérée]`.
- Critère d'acceptation : aucune phrase de sortie ne contient une information absente de la phrase d'entrée correspondante.

**Structure fixe du récit et styles Word**
- 4 styles de paragraphes appliqués dans le `.docx` de sortie.
- Toutes les sections sont présentes, même vides.
- Critère d'acceptation : le fichier s'ouvre sans erreur dans Word et LibreOffice avec les styles visibles.

**Prévisualisation dans le navigateur**
- Aperçu du récit avec passages incertains et suggestions de notes mis en évidence.

### Nice-to-have (P1)

**Mode debug / transparence API**
- Affichage du texte exact envoyé à l'API (post-anonymisation) pour vérification.

**Export du résumé d'audience seul**
- Possibilité d'exporter uniquement le tableau de résumé (section 1) pour envoi rapide.

### Futur / v2 (P2)

- Notes de bas de page : détection automatique + intégration avec le dictionnaire MSDE.
- Transcription audio → notes brutes → récit.
- Modèle local (Ollama) pour éliminer tout transfert de données.
- Historique des récits avec recherche par date / affaire.
- Intégration avec le template Word MSDE pour mise en page automatique.

---

## Métriques de succès

| Métrique | Cible | Méthode |
|---|---|---|
| Temps de production d'un récit | < 30 min (vs ~8h) | Suivi sur les 10 premiers récits |
| Taux d'utilisation | 100 % des bénévoles actifs à J+30 | Fichiers traités |
| Taux de document utilisable sans relecture lourde | ≥ 80 % | Évaluation qualitative coordinateurs |
| Conformité à la charte rédactionnelle | ≥ 90 % des règles appliquées correctement | Audit sur 20 paragraphes |
| Taux de passages `[incertain]` | < 10 % des paragraphes | Comptage automatique |
| Délai audience → diffusion récit | < 4h (vs ~48h) | Suivi à J+60 |

---

## Contraintes et risques

**Confidentialité**
Le mécanisme d'anonymisation pré-API garantit que les données identifiantes ne quittent pas le poste. Risque résiduel : détection imparfaite des noms propres peu courants. Le scribe peut vérifier la liste des éléments anonymisés avant envoi (affichée dans l'interface).

**RGPD**
Le MSDE doit formaliser sa politique d'utilisation d'APIs IA avant déploiement large. Bloquant pour le passage à l'échelle.

**Fidélité factuelle**
Le modèle peut mal interpréter les passages de niveau 3 (quasi-sténo). Le marqueur `[passage incertain]` est le principal filet de sécurité. Relecture humaine obligatoire avant diffusion.

**Coût API**
~0,05 € par document de 10 pages avec Claude Sonnet. Non bloquant à faible volume.

---

## Questions ouvertes

| Question | Responsable | Bloquante ? |
|---|---|---|
| Existe-t-il des récits finalisés de référence (few-shot) à intégrer dans le prompt ? | Coordinateurs MSDE | Oui — détermine la qualité du prompt |
| Le MSDE a-t-il une politique RGPD formalisée pour l'usage d'APIs IA ? | Direction MSDE | Oui — avant déploiement large |
| Comment générer des pseudonymes à consonance proche ? Règle linguistique ou appel API séparé ? | À définir | Non — peut être itéré |
| Faut-il des profils de bibliothèque d'acronymes par scribe, ou une bibliothèque commune suffit ? | Coordinateurs MSDE | Non — la résolution interactive par session compense |
| Le dictionnaire de notes de bas de page (frama.space) est-il exportable / intégrable en v2 ? | Coordinateurs MSDE | Non — v2 |

---

## Jalons proposés (v1)

| Phase | Contenu | Durée estimée |
|---|---|---|
| **Phase 1 — Prototype** | Upload `.docx`, anonymisation, traitement API, export `.docx` structuré | 1 semaine |
| **Phase 2 — Bibliothèque & résolution** | Interface bibliothèque, résolution interactive des acronymes inconnus | 1 semaine |
| **Phase 3 — Test sur cas réels** | Test sur 5 notes réelles, ajustement prompt, intégration des règles de la charte | 1-2 semaines |
| **Phase 4 — Consolidation** | Marqueurs `[incertain]` / `[note suggérée]`, prévisualisation, mode debug, documentation | 1 semaine |
| **Phase 5 — Déploiement** | Installation postes bénévoles, formation 30 min, collecte métriques | 1 semaine |

Durée totale estimée v1 : **5 à 6 semaines** à rythme partiel.

---

## Annexe — Niveaux de dégradation des notes (document exemple)

**Niveau 1 — Notes claires** (déclarations Malo Royen, Jessica Remaillon) : phrases complètes, fautes mineures. Traitement simple.

**Niveau 2 — Notes compressées** (Christian Farel, Céléstin) : mots omis, orthographe approximative, phrases inachevées. Complétion prudente requise.

**Niveau 3 — Notes quasi-inchiffrables** (réquisitions procureur, plaidoiries avocates) : quasi-sténo. Exemples : *"pck on a aanticipé les désgarémaents"*, *"cnsq de RC"*, *"ENjoit ) prendre en compte"*. Ces passages doivent être marqués `[passage incertain — à vérifier]` si la reconstitution n'est pas certaine.

Le prompt système doit être calibré sur les niveaux 2 et 3. Le respect du principe de fidélité est critique ici : mieux vaut un `[passage incertain]` qu'une reformulation inventée.
