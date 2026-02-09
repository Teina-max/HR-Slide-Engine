# Catalogue des Layouts — HR Slide Engine

## Layouts textuels

### 1. `add_title_slide(prs, title, subtitle, notes)`

**Usage** : Première slide de la présentation.
**Visuel** : Fond navy, titre blanc centré (36pt bold), ligne orange horizontale, sous-titre gris clair.
**Quand l'utiliser** : Toujours en slide 1.

### 2. `add_agenda_slide(prs, items, title, notes)`

**Usage** : Plan / sommaire de la présentation.
**Visuel** : Fond blanc, titre navy, numéros orange (01, 02...) alignés à gauche, items en corps.
**Quand l'utiliser** : Slide 2, après le titre. 3 à 6 items idéalement.

### 3. `add_section_slide(prs, title, subtitle, notes)`

**Usage** : Transition entre les parties.
**Visuel** : Fond blanc, barre navy verticale à gauche, titre large navy, sous-titre gris optionnel.
**Quand l'utiliser** : Avant chaque nouvelle partie (ex: "1. Contexte et enjeux").

### 4. `add_bullets_slide(prs, title, bullets, notes)`

**Usage** : Liste de points / arguments.
**Visuel** : Fond blanc, titre navy avec soulignement orange, bullets orange avec texte gris.
**Quand l'utiliser** : Pour lister des arguments, constats, actions. 3 à 6 bullets.

### 5. `add_two_columns_slide(prs, title, left_title, left_items, right_title, right_items, notes)`

**Usage** : Comparaison, avant/après, deux approches.
**Visuel** : Fond blanc, titre navy, deux colonnes avec sous-titres navy et bullets orange, séparateur gris vertical.
**Quand l'utiliser** : Comparaisons, avantages/inconvénients, court terme/long terme.

### 6. `add_key_stat_slide(prs, stat, description, notes)`

**Usage** : Chiffre clé impactant.
**Visuel** : Fond blanc, gros chiffre orange centré (72pt bold), description gris en dessous.
**Quand l'utiliser** : Pour marquer les esprits avec un KPI, un %, un montant. 1 stat par slide.

### 7. `add_quote_slide(prs, quote, author, notes)`

**Usage** : Citation académique ou d'expert.
**Visuel** : Fond gris clair, guillemet décoratif orange (120pt), citation navy, auteur gris.
**Quand l'utiliser** : Pour ancrer un propos dans un cadre théorique. Parfait pour les définitions.

### 8. `add_conclusion_slide(prs, title, points, notes)`

**Usage** : Slide de conclusion / synthèse.
**Visuel** : Bannière navy en haut avec titre blanc, checkmarks orange avec points de synthèse.
**Quand l'utiliser** : Avant-dernière ou dernière slide. 3 à 5 points de synthèse.

---

## Layouts visuels

### 9. `add_process_flow_slide(prs, title, steps, notes)`

**Usage** : Processus en étapes séquentielles.
**Visuel** : Chevrons colorés connectés horizontalement, numéros dans des cercles au-dessus, description en dessous.
**Quand l'utiliser** : Processus GPEC, parcours collaborateur, workflow RH, cycle de recrutement. 3 à 6 étapes.

### 10. `add_timeline_slide(prs, title, milestones, notes)`

**Usage** : Frise chronologique.
**Visuel** : Ligne horizontale navy, points orange, dates et descriptions alternées au-dessus/en-dessous.
**Paramètres** : `milestones` = liste de tuples `(date_label, description)`.
**Quand l'utiliser** : Évolution législative, phasage plan d'action, historique entreprise. 3 à 6 jalons.

### 11. `add_matrix_slide(prs, title, top_left, top_right, bottom_left, bottom_right, x_label, y_label, notes)`

**Usage** : Matrice 2x2 à quatre quadrants.
**Visuel** : Quatre rectangles arrondis colorés (bleu, orange, vert, rouge), bullets dans chaque quadrant, labels d'axes optionnels.
**Paramètres** : Chaque quadrant est un dict `{"title": "...", "items": ["..."]}`.
**Quand l'utiliser** : SWOT, matrice compétences/criticité, urgence/importance, BCG. Idéal pour les analyses croisées.

### 12. `add_pyramid_slide(prs, title, levels, notes)`

**Usage** : Pyramide hiérarchique.
**Visuel** : Barres horizontales colorées qui rétrécissent vers le haut, texte blanc centré dans chaque niveau.
**Paramètres** : `levels` = liste du sommet (plus étroit) vers la base (plus large).
**Quand l'utiliser** : Maslow, niveaux de compétences, maturité GPEC, hiérarchie organisationnelle. 3 à 6 niveaux.

### 13. `add_bar_chart_slide(prs, title, categories, values, notes)`

**Usage** : Graphique en barres verticales.
**Visuel** : Barres orange, axes gris, grille légère. Graphique natif PowerPoint (éditable).
**Quand l'utiliser** : Comparaisons chiffrées, turnover par département, évolution budget, KPIs par année.

### 14. `add_pie_chart_slide(prs, title, categories, values, notes)`

**Usage** : Graphique camembert.
**Visuel** : Segments multicolores, pourcentages affichés, légende en bas. Graphique natif PowerPoint (éditable).
**Quand l'utiliser** : Répartition effectifs, budget formation par type, structure contrats, composition équipe.

### 15. `add_icon_cards_slide(prs, title, cards, notes)`

**Usage** : Dashboard de KPIs / métriques.
**Visuel** : Grille de cartes avec barre de couleur en haut, gros chiffre coloré, label descriptif en dessous.
**Paramètres** : `cards` = liste de dicts `{"value": "78%", "label": "Satisfaction"}`. Max 6 cartes.
**Quand l'utiliser** : Tableau de bord RH, bilan social en un coup d'œil, KPIs clés, résultats d'enquête.

---

## Règles de sélection automatique

| Contenu détecté                    | Layout recommandé   |
|------------------------------------|---------------------|
| Introduction, titre principal      | `title`             |
| Plan, sommaire, agenda             | `agenda`            |
| Transition de partie               | `section`           |
| Liste d'arguments ou constats      | `bullets`           |
| Comparaison, deux catégories       | `two_columns`       |
| Chiffre clé, statistique unique    | `key_stat`          |
| Citation, définition académique    | `quote`             |
| Synthèse, conclusion, à retenir    | `conclusion`        |
| Processus, étapes, workflow        | `process_flow`      |
| Chronologie, évolution, phasage    | `timeline`          |
| SWOT, matrice 2x2, analyse croisée| `matrix`            |
| Hiérarchie, niveaux, pyramide      | `pyramid`           |
| Données chiffrées comparatives     | `bar_chart`         |
| Répartition, proportions           | `pie_chart`         |
| Dashboard, KPIs multiples          | `icon_cards`        |
