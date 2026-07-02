# Politique de protection des données — Outil IA de traitement des récits d'audience

**Version** 0.1 — Ébauche  
**Date** juin 2026  
**Statut** À valider par la direction du MSDE

---

## Pourquoi ce document existe

Le MSDE utilise un outil d'intelligence artificielle (API Claude d'Anthropic) pour traiter les notes d'audience brutes et produire des récits structurés. Cet outil traite des informations relatives à des personnes physiques identifiables (prévenu·es, témoins, avocat·es). Le Règlement Général sur la Protection des Données (RGPD, UE 2016/679) impose aux associations comme le MSDE de documenter formellement comment ces données sont traitées, protégées et pour quelle finalité.

Ce document décrit la politique du MSDE sur ce point spécifique.

---

## Ce que le MSDE traite comme données personnelles

Dans le cadre de l'observation des audiences judiciaires, les scribes du MSDE collectent :

- Noms et prénoms des prévenu·es et témoins
- Dates et lieux de naissance
- Adresses de domicile
- Situations professionnelles
- Revenus déclarés à l'audience
- Antécédents judiciaires (casier)
- Éléments de vie personnelle exprimés à l'audience

Ces données sont saisies dans des notes brutes (fichiers Word) par les scribes pendant ou après l'audience.

---

## Finalité du traitement

**Finalité principale :** produire un récit d'audience fidèle et structuré, destiné à :
- Informer les avocats partenaires du MSDE sur le déroulement de la procédure.
- Informer le réseau militant et le public sur les procès observés.
- Constituer une mémoire associative des procès liés aux mobilisations climatiques et sociales.

**Base légale (art. 6 RGPD) :** intérêt légitime de l'association (art. 6.1.f), dans le cadre de sa mission d'observation judiciaire et de défense des droits des personnes poursuivies pour fait de militantisme.

> Note : les audiences judiciaires sont publiques par principe (art. 306 et 400 CPP). Les informations recueillies en audience le sont dans le cadre d'une activité légale d'observation. Le MSDE n'accède pas à des dossiers judiciaires confidentiels.

---

## Mesure centrale de protection : anonymisation avant envoi à l'IA

**Aucune donnée personnelle réelle n'est transmise à l'API Anthropic.**

Avant tout envoi à l'outil IA, les données personnelles sont remplacées par des pseudonymes choisis par le scribe (noms fictifs à consonance proche pour les prévenu·es et témoins, initiales pour les avocat·es). La table de correspondance nom réel / pseudonyme reste sur le poste local du scribe et n'est jamais transmise.

En conséquence, Anthropic ne traite **jamais** de données personnelles au sens du RGPD dans le cadre de cet usage. Anthropic intervient uniquement sur des textes pseudonymisés, ce qui le soustrait à la qualification de sous-traitant de données personnelles pour ce traitement.

Le document final (le récit publié) contient uniquement les pseudonymes, conformément aux règles d'anonymisation du MSDE.

---

## Qui accède aux données

| Acteur | Accès | Données concernées |
|---|---|---|
| Scribe MSDE | Plein accès | Notes brutes avec données réelles |
| Coordinateur·rice MSDE | Accès au récit pseudonymisé | Pseudonymes uniquement |
| Avocat·es partenaires | Accès au récit pseudonymisé | Pseudonymes uniquement |
| Réseau militant / public | Accès au récit publié | Pseudonymes uniquement |
| Anthropic (API) | Accès au texte pseudonymisé uniquement | Pseudonymes uniquement — pas de données réelles |

---

## Durée de conservation

| Support | Durée | Justification |
|---|---|---|
| Notes brutes (fichier Word) | Supprimées après production du récit finalisé | Données sensibles non nécessaires à la conservation |
| Table de correspondance pseudonymes | Supprimée à la fin de la session de traitement | Donnée temporaire, plus nécessaire après validation |
| Récit finalisé (pseudonymisé) | Conservation longue durée | Mission mémorielle de l'association |

> En pratique : le scribe est invité à supprimer le fichier de notes brutes après validation du récit. L'outil peut rappeler cette recommandation en fin de session.

---

## Droits des personnes concernées

Toute personne dont les données ont été collectées (prévenu·e, témoin) dispose des droits suivants au titre du RGPD :

- **Droit d'accès** (art. 15) : obtenir confirmation que des données la concernant sont traitées et en obtenir copie.
- **Droit de rectification** (art. 16) : faire corriger des données inexactes.
- **Droit à l'effacement** (art. 17) : demander la suppression, sous conditions.
- **Droit d'opposition** (art. 21) : s'opposer au traitement fondé sur l'intérêt légitime.

**Contact pour exercer ces droits :** [adresse email de contact du MSDE à compléter]

**Délai de réponse :** 1 mois à compter de la réception de la demande (art. 12 RGPD).

> Note pratique : dans la mesure où les récits publiés utilisent des pseudonymes, la possibilité d'identifier une personne dans un récit publié est limitée. Le risque d'exercice effectif de ces droits sur les récits est donc faible. En revanche, les notes brutes (si conservées) contiennent des données réelles et doivent être traitées avec soin.

---

## Mesures de sécurité

- Les notes brutes sont stockées uniquement sur le poste local du scribe (pas de stockage cloud sans chiffrement).
- L'outil IA tourne en local (pas de serveur partagé en v1).
- L'anonymisation est une étape obligatoire et non contournable avant tout envoi à l'API.
- La bibliothèque d'acronymes (partagée entre scribes) ne contient aucune donnée personnelle.
- Aucun log des textes traités n'est conservé côté Anthropic au-delà de 30 jours (politique Anthropic à vérifier et à documenter).

---

## Ce que le MSDE doit encore formaliser

- [ ] Désigner un référent protection des données au sein de l'association (DPO informel, pas obligatoire pour une association de cette taille mais recommandé).
- [ ] Inscrire ce traitement dans le **registre des activités de traitement** de l'association (obligatoire, art. 30 RGPD).
- [ ] Informer les personnes concernées de l'existence de ce traitement (information lors de l'observation de l'audience, ou mention dans les récits publiés).
- [ ] Vérifier la politique de rétention des données d'Anthropic et s'assurer qu'elle est compatible avec les engagements du MSDE.
- [ ] Mettre à jour les statuts ou le règlement intérieur de l'association pour mentionner l'activité d'observation judiciaire et de publication de récits.

---

## Limites de ce document

Cette ébauche n'est pas un avis juridique. Elle constitue une base de travail pour formaliser la politique du MSDE. Avant tout déploiement à grande échelle de l'outil IA, il est recommandé de faire relire ce document par un·e juriste spécialisé·e en protection des données, idéalement via le réseau d'avocats partenaires du MSDE.
