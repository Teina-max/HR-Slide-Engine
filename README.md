# HR Slide Engine

Skill Claude Code qui transforme du texte brut en présentations PowerPoint professionnelles pour le Master RH.

## Installation

```bash
# 1. Cloner le repo
git clone https://github.com/YOUR_USERNAME/HR-Slide-Engine.git
cd HR-Slide-Engine

# 2. Installer les dépendances Python
pip install python-pptx

# 3. Installer le skill Claude Code
python install.py
```

## Usage

Dans Claude Code :

```
/pptx-master-rh La GPEC comme levier stratégique de transformation RH
```

Ou avec du texte plus détaillé :

```
/pptx-master-rh

La Qualité de Vie au Travail dans le secteur hospitalier.
Contexte : 45% des soignants déclarent un épuisement professionnel (HAS, 2024).
Cadre théorique : Karasek (1979), Siegrist (1996), Clot (2010).
Diagnostic : turnover de 18%, absentéisme en hausse de 12%.
Préconisations : espaces de discussion, aménagement du temps, formation managériale.
```

Le skill génère automatiquement un fichier `.pptx` dans le répertoire courant.

## Design System

| Couleur | Hex | Usage |
|---------|-----|-------|
| Navy | `#1B2A4A` | Titres, fonds |
| Gray | `#6B7280` | Corps de texte |
| Orange | `#E87C3E` | Accents, bullets |
| White | `#FFFFFF` | Fonds, texte sur navy |

**Typographie** : Calibri (Bold 28pt / Regular 20pt / Light 16pt)
**Format** : 16:9

## 8 Layouts disponibles

1. **Title** — Fond navy, texte blanc centré
2. **Agenda** — Numéros orange, items listés
3. **Section** — Barre navy à gauche, transition
4. **Bullets** — Bullets orange, texte gris
5. **Two Columns** — Comparaison côte à côte
6. **Key Stat** — Gros chiffre orange centré
7. **Quote** — Fond gris, guillemet décoratif
8. **Conclusion** — Bannière navy, checkmarks orange

## Structure du projet

```
HR-Slide-Engine/
├── slide_engine/          # Module Python
│   ├── design.py          # Constantes design
│   ├── engine.py          # Fonctions core
│   └── layouts.py         # 8 fonctions de layout
├── skill/                 # Skill Claude Code
│   ├── SKILL.md           # Pipeline 3 passes
│   └── references/        # Docs de référence
├── tests/                 # Tests pytest
├── install.py             # Script d'installation
└── requirements.txt
```

## Développement

```bash
# Lancer les tests
python -m pytest tests/ -v

# Générer un .pptx de test
python -c "
from slide_engine import *
prs = create_presentation()
add_title_slide(prs, 'Test', 'Sous-titre')
add_bullets_slide(prs, 'Points', ['A', 'B', 'C'])
save_presentation(prs, 'test_output')
"
```

## Licence

MIT
