# Catalogue des Layouts — HR Slide Engine

## 1. `add_title_slide(prs, title, subtitle, notes)`

**Usage** : Première slide de la présentation.
**Visuel** : Fond navy, titre blanc centré (36pt bold), ligne orange horizontale, sous-titre gris clair.
**Quand l'utiliser** : Toujours en slide 1.

## 2. `add_agenda_slide(prs, items, title, notes)`

**Usage** : Plan / sommaire de la présentation.
**Visuel** : Fond blanc, titre navy, numéros orange (01, 02...) alignés à gauche, items en corps.
**Quand l'utiliser** : Slide 2, après le titre. 3 à 6 items idéalement.

## 3. `add_section_slide(prs, title, subtitle, notes)`

**Usage** : Transition entre les parties.
**Visuel** : Fond blanc, barre navy verticale à gauche, titre large navy, sous-titre gris optionnel.
**Quand l'utiliser** : Avant chaque nouvelle partie (ex: "1. Contexte et enjeux").

## 4. `add_bullets_slide(prs, title, bullets, notes)`

**Usage** : Liste de points / arguments.
**Visuel** : Fond blanc, titre navy avec soulignement orange, bullets orange avec texte gris.
**Quand l'utiliser** : Pour lister des arguments, constats, actions. 3 à 6 bullets.

## 5. `add_two_columns_slide(prs, title, left_title, left_items, right_title, right_items, notes)`

**Usage** : Comparaison, avant/après, deux approches.
**Visuel** : Fond blanc, titre navy, deux colonnes avec sous-titres navy et bullets orange, séparateur gris vertical.
**Quand l'utiliser** : Comparaisons, avantages/inconvénients, court terme/long terme.

## 6. `add_key_stat_slide(prs, stat, description, notes)`

**Usage** : Chiffre clé impactant.
**Visuel** : Fond blanc, gros chiffre orange centré (72pt bold), description gris en dessous.
**Quand l'utiliser** : Pour marquer les esprits avec un KPI, un %, un montant. 1 stat par slide.

## 7. `add_quote_slide(prs, quote, author, notes)`

**Usage** : Citation académique ou d'expert.
**Visuel** : Fond gris clair, guillemet décoratif orange (120pt), citation navy, auteur gris.
**Quand l'utiliser** : Pour ancrer un propos dans un cadre théorique. Parfait pour les définitions.

## 8. `add_conclusion_slide(prs, title, points, notes)`

**Usage** : Slide de conclusion / synthèse.
**Visuel** : Bannière navy en haut avec titre blanc, checkmarks orange avec points de synthèse.
**Quand l'utiliser** : Avant-dernière ou dernière slide. 3 à 5 points de synthèse.

---

## Règles de sélection automatique

| Contenu détecté                    | Layout recommandé  |
|------------------------------------|--------------------|
| Introduction, titre principal      | `title`            |
| Plan, sommaire, agenda             | `agenda`           |
| Transition de partie               | `section`          |
| Liste d'arguments ou constats      | `bullets`          |
| Comparaison, deux catégories       | `two_columns`      |
| Chiffre clé, statistique, KPI      | `key_stat`         |
| Citation, définition académique    | `quote`            |
| Synthèse, conclusion, à retenir    | `conclusion`       |
