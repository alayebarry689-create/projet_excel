# CongoTopFashion — Analyse de la base de données commerciale

**Projet de portfolio en analyse de données | Diagnostic d'entreprise et recommandations à la direction**

> Analyse complète d'une chaîne de distribution de mode opérant en République Démocratique du Congo : contrôle qualité des données, statistiques descriptives, segmentation client, classification produits, prévision de ventes et plan de recommandations chiffrées.

---

## 1. Contexte du projet

CongoTopFashion est une entreprise fictive (données simulées à des fins pédagogiques) de distribution de vêtements et d'accessoires de mode, active en République Démocratique du Congo à travers un réseau de boutiques physiques et plusieurs canaux numériques (site e-commerce, Facebook, Instagram, WhatsApp Business, vente terrain).

L'objectif de ce projet était de reconstituer, à partir d'une base de données relationnelle brute, un diagnostic complet de la performance de l'entreprise — ventes, produits, clients, logistique, marketing, stocks — et d'en tirer un plan d'action opérationnel chiffré, présentable à un comité de direction.

**Périmètre de la base analysée :**

| Élément | Valeur |
|---|---|
| Tables relationnelles | 24 |
| Lignes total (toutes tables) | 100 351 |
| Commandes | 15 000 |
| Clients | 2 000 |
| Références produits | 700 |
| Boutiques | 12 (11 physiques + 1 plateforme e-commerce) |
| Période couverte | 2021 – 2025 (5 exercices complets) |

---

## 2. Objectifs de l'analyse

1. **Valider la fiabilité de la base** avant toute exploitation (doublons, valeurs manquantes, intégrité référentielle, cohérence arithmétique).
2. **Mesurer la performance réelle** de l'entreprise (chiffre d'affaires, marge, rentabilité) et **corriger les biais** du reporting existant.
3. **Comprendre la structure de la valeur** : quels canaux, boutiques, produits et clients génèrent réellement le résultat.
4. **Diagnostiquer les zones de risque** : stocks, objectifs commerciaux, attribution marketing, logistique.
5. **Projeter l'activité 2026** par une méthode statistique reproductible.
6. **Formuler des recommandations priorisées**, avec impact estimé et échéance, exploitables immédiatement par la direction.

---

## 3. Contenu du dossier `Livrables`

| Fichier | Description | Format |
|---|---|---|
| [`CongoTopFashion_Analyse.xlsx`](CongoTopFashion_Analyse.xlsx) | Classeur d'analyse complet : 10 onglets, tableaux croisés, formules de calcul, 12 graphiques natifs Excel. Document de travail source de toutes les analyses. | Excel |
| [`CongoTopFashion_Rapport_Analyse.docx`](CongoTopFashion_Rapport_Analyse.docx) | Rapport rédigé détaillant la méthodologie, les résultats et les recommandations, avec commentaires d'interprétation. | Word |
| [`CongoTopFashion_Presentation_Direction.pptx`](CongoTopFashion_Presentation_Direction.pptx) | Support de présentation de 17 slides pour un comité de direction : chiffres clés, constats, recommandations et synthèse des impacts. | PowerPoint |
| [`CongoTopFashion_Dashboard.html`](CongoTopFashion_Dashboard.html) | Tableau de bord interactif (HTML/CSS/JS autonome, ouvrable dans n'importe quel navigateur sans installation) reprenant l'ensemble des indicateurs et graphiques sous forme narrative. | Web |
| `README.md` | Ce document. | Markdown |

Les quatre livrables présentent le **même contenu analytique** sous des formats adaptés à des usages différents : le classeur Excel pour l'exploration et l'audit des calculs, le rapport Word pour une lecture détaillée, la présentation PowerPoint pour un comité de direction, et le tableau de bord HTML pour une consultation rapide et visuelle.

---

## 4. Méthodologie

### 4.1 Contrôle et nettoyage des données

Avant toute analyse, la base a été auditée table par table :

- **Doublons** : recherche de lignes dupliquées et d'identifiants dupliqués sur les 19 tables transactionnelles → **aucun doublon détecté**.
- **Intégrité référentielle** : vérification des 18 relations de clé étrangère (ex. `Commandes.ID_Client → Clients`, `Détails_Commandes.ID_Produit → Produits`) → **zéro référence orpheline**.
- **Valeurs aberrantes** : contrôle des montants négatifs, marges négatives, quantités invalides, âges hors plage, délais de livraison négatifs → **aucune anomalie**.
- **Cohérence arithmétique** : vérification que `Montant Brut − Remise = Montant Net TTC` sur les 15 000 commandes.
- **Valeurs manquantes** : identifiées et qualifiées comme **structurelles** plutôt que comme défauts de saisie (ex. 8 607 commandes sans vendeur = ventes en ligne sans intervention humaine ; 931 livraisons sans date d'expédition = commandes annulées).
- **Traitements appliqués** : suppression des espaces superflus dans les champs texte, conversion explicite des colonnes de date.

### 4.2 Méthodes statistiques et d'analyse employées

| Méthode | Application |
|---|---|
| Marge pondérée vs moyenne des pourcentages | Correction d'un biais de calcul dans le reporting existant (33,45 % affiché vs 28,76 % réel — écart de 4,7 points) |
| Classification ABC / loi de Pareto | Segmentation des 700 références produits selon leur contribution au chiffre d'affaires |
| Segmentation RFM (Récence, Fréquence, Montant) | Segmentation des 2 000 clients en 7 profils comportementaux |
| Calcul de la valeur vie client (CLV) | Comparaison de la valeur générée par les clients VIP et non-VIP |
| Analyse du retour sur investissement publicitaire (ROAS) | Évaluation de la rentabilité par canal marketing et détection d'un biais d'attribution |
| Régression avec correction de saisonnalité | Projection mensuelle du chiffre d'affaires 2026 sur la base de 60 mois d'historique |
| Analyse de couverture de stock et rotation | Diagnostic du niveau de stock par rapport au rythme réel des ventes |
| Analyse des écarts d'inventaire | Comparaison stock théorique vs comptages physiques sur 500 inventaires |

---

## 5. Structure du classeur Excel (`CongoTopFashion_Analyse.xlsx`)

Le classeur est organisé en **10 onglets séquentiels**, du diagnostic global au plan d'action :

1. **Synthèse Direction** — KPI globaux, évolution annuelle 2021-2025, six points d'alerte prioritaires.
2. **Qualité des données** — résultats des contrôles de doublons, intégrité référentielle, valeurs aberrantes et manquantes.
3. **Ventes** — répartition du chiffre d'affaires par canal, boutique, catégorie de produit et province.
4. **Produits** — classification ABC, top 15 produits, top 10 marques.
5. **Clients** — segmentation RFM, indicateurs de valeur client (CLV, écart VIP), statut d'activité, répartition géographique et canaux d'acquisition.
6. **Logistique-Retours** — performance par mode de livraison et transporteur, motifs de retour, retours par catégorie.
7. **Marketing** — performance par canal (budget, ROAS, coût par conversion), impact des promotions sur la marge.
8. **Stocks-Achats** — indicateurs de couverture de stock, écarts d'inventaire, top 10 fournisseurs.
9. **Objectifs-Prévision** — taux de réalisation des objectifs par boutique, projection mensuelle 2026.
10. **Recommandations** — 12 actions priorisées avec constat chiffré, nature de l'impact, impact estimé et échéance.

---

## 6. Principaux résultats

### 6.1 Chiffres clés (cumul 2021-2025)

| Indicateur | Valeur |
|---|---|
| Chiffre d'affaires cumulé | 2,39 M$ |
| Bénéfice net cumulé | 686 k$ |
| Marge nette réelle (pondérée) | 28,8 % |
| Panier moyen | 170 $ |
| Taux de retour | 13,3 % |
| Taux d'annulation | 6,2 % |
| Croissance 2025 vs 2024 | +16,4 % (contre +51,5 % en 2022) |

### 6.2 Six constats majeurs

1. **Surstock critique** — 8,02 M$ de stock valorisé pour 274 k$ de coût des ventes annuel, soit **29 années de couverture** ; 4,77 M$ immobilisés sur les seules lignes en surstock.
2. **Objectifs commerciaux inopérants** — aucun des 500 objectifs fixés n'a été atteint en 5 ans (taux de réalisation moyen : 14,2 %) ; l'objectif annuel par boutique (376 767 $) représente près de huit fois le réalisé (45 016 $).
3. **Marge nette surestimée dans le reporting actuel** — 33,45 % affichés (moyenne des pourcentages par commande) contre 28,76 % réels (marge pondérée), soit un écart de 4,7 points qui fausse les décisions de tarification.
4. **Attribution marketing incohérente** — les campagnes déclarent 3,25 M$ de chiffre d'affaires généré, soit 136 % du chiffre d'affaires réel de l'entreprise ; 55 campagnes affichent un ROAS inférieur à 1 (192 944 $ sans retour).
5. **Croissance en décélération** — le rythme de croissance annuel est passé de +51 % (2022) à +16 % (2025), divisé par trois en quatre ans.
6. **Écarts d'inventaire généralisés** — 87,8 % des comptages d'inventaire présentent un écart avec le stock théorique, pour une perte brute de 13 644 $.

### 6.3 Segmentation client (RFM)

| Segment | Clients | % de la base | % du CA |
|---|---|---|---|
| Champions | 341 | 17,1 % | 49,2 % |
| Clients Fidèles | 350 | 17,5 % | 28,1 % |
| Clients Réguliers | 356 | 17,8 % | 9,7 % |
| Clients Perdus | 387 | 19,4 % | 5,4 % |
| Clients à Risque | 139 | 7,0 % | 4,8 % |
| Nouveaux Clients | 170 | 8,5 % | 2,8 % |
| Sans Achat | 257 | 12,9 % | 0,0 % |

Les 341 clients **Champions** (17 % de la base) génèrent à eux seuls **49,2 % du chiffre d'affaires**, avec une valeur vie moyenne de 6 200 $ contre 843 $ pour les clients non-VIP (écart de 7,4x).

### 6.4 Classification produits (ABC / Pareto)

| Classe | Références | % du catalogue | % du CA |
|---|---|---|---|
| A | 254 | 36,4 % | 79,9 % |
| B | 204 | 29,3 % | 15,0 % |
| C | 239 | 34,3 % | 5,0 % |

### 6.5 Projection 2026

Régression sur 60 mois d'historique avec correction saisonnière : chiffre d'affaires prévisionnel de **874 k$ pour 2026, soit +23,3 % vs 2025**, avec un pic attendu en décembre (113 204 $).

---

## 7. Plan de recommandations

12 recommandations opérationnelles, classées par priorité, chacune assortie d'un constat chiffré, d'un impact estimé et d'une échéance :

| # | Recommandation | Impact estimé | Échéance |
|---|---|---|---|
| 1 | Déstocker et geler les réapprovisionnements | 2 à 3 M$ de trésorerie | Immédiat |
| 2 | Refonder le système d'objectifs commerciaux | Non chiffrable | Immédiat |
| 3 | Corriger le calcul de la marge dans le reporting | 4,7 points | Immédiat |
| 4 | Réduire les retours liés aux problèmes de taille | 18 à 24 k$/an | 3 mois |
| 5 | Auditer et réallouer le budget marketing | 190 à 400 k$ | 3 mois |
| 6 | Concentrer l'effort de fidélisation sur les Champions | 1,17 M$ protégés | 6 mois |
| 7 | Réactiver les 631 clients perdus et inactifs | 80 à 150 k$/an | 6 mois |
| 8 | Améliorer la fiabilité de la livraison à domicile | +1,1 point /5 satisfaction | 6 mois |
| 9 | Réviser la politique promotionnelle | 3,9 points de marge | 3 mois |
| 10 | Renforcer le contrôle des inventaires | 10 à 14 k$/an | 6 mois |
| 11 | Développer le e-commerce et réduire la dépendance à Kinshasa | Réduction du risque | 12 mois |
| 12 | Homogénéiser la performance des vendeurs | 60 à 100 k$/an | 12 mois |

**Potentiel d'amélioration du résultat annuel** (hors effet de trésorerie ponctuel) : entre **394 000 $ et 724 000 $**, à comparer à un bénéfice annuel 2025 de 203 096 $ — soit une possibilité de **plus que doubler la rentabilité** de l'entreprise via la mise en œuvre des recommandations prioritaires.

---

## 8. Compétences mises en œuvre

Ce projet illustre les compétences suivantes, pertinentes pour un portfolio d'analyste de données / analyste business :

- **Modélisation et audit de bases de données relationnelles** (24 tables, contrôle systématique de l'intégrité et de la qualité).
- **Excel avancé** : tableaux croisés dynamiques, formules conditionnelles multi-critères, calculs pondérés, mise en forme professionnelle, graphiques natifs.
- **Statistiques appliquées** : segmentation RFM, classification ABC/Pareto, régression avec saisonnalité, calcul de la valeur vie client.
- **Détection et correction de biais analytiques** : identification d'un biais de calcul de marge et d'un biais d'attribution marketing dans un reporting existant.
- **Storytelling et communication data** : traduction d'une analyse technique en un rapport, une présentation de direction et un tableau de bord, adaptés chacun à leur audience.
- **Recommandations orientées décision** : chaque conclusion est reliée à une action concrète, un impact chiffré et un responsable/délai.

---

## 9. Comment utiliser ces livrables

- **Pour explorer les calculs en détail** → ouvrir [`CongoTopFashion_Analyse.xlsx`](CongoTopFashion_Analyse.xlsx) et naviguer entre les 10 onglets.
- **Pour une lecture complète et argumentée** → consulter [`CongoTopFashion_Rapport_Analyse.docx`](CongoTopFashion_Rapport_Analyse.docx).
- **Pour une présentation orale à un comité de direction** → utiliser [`CongoTopFashion_Presentation_Direction.pptx`](CongoTopFashion_Presentation_Direction.pptx) (17 slides, prêt à projeter).
- **Pour une consultation rapide et interactive** → ouvrir [`CongoTopFashion_Dashboard.html`](CongoTopFashion_Dashboard.html) dans un navigateur (aucune installation requise).

---

## 10. Précisions et limites

- Les données utilisées sont **simulées** (jeu de données fictif à visée pédagogique) ; les montants, noms de boutiques et de fournisseurs ne correspondent à aucune entreprise réelle.
- Les montants sont exprimés en dollars américains (USD), devise de référence utilisée dans le jeu de données source.
- La projection 2026 constitue un **scénario de référence** (prolongation de la tendance historique) et n'intègre pas l'effet des recommandations proposées dans la section 7.

---

*Dernière mise à jour : 29 septembre 2026*
