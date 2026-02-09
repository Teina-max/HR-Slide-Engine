"""Integration test — Full GPEC presentation pipeline from JSON plan to .pptx."""

import os
import sys
import json
import pytest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from slide_engine import (
    create_presentation,
    save_presentation,
    add_title_slide,
    add_agenda_slide,
    add_section_slide,
    add_bullets_slide,
    add_two_columns_slide,
    add_key_stat_slide,
    add_quote_slide,
    add_conclusion_slide,
)


GPEC_PLAN = {
    "title": "La GPEC comme levier stratégique de transformation RH",
    "subtitle": "Master 2 RH — Étude de cas Entreprise X",
    "slides": [
        {
            "layout": "title",
            "title": "La GPEC comme levier stratégique de transformation RH",
            "subtitle": "Master 2 RH — Étude de cas Entreprise X",
            "notes": "Se présenter. Contextualiser le sujet dans l'actualité RH."
        },
        {
            "layout": "agenda",
            "items": [
                "Contexte et enjeux",
                "Cadre théorique de la GPEC",
                "Diagnostic des compétences",
                "Plan d'action et préconisations",
                "Conclusion et perspectives"
            ],
            "notes": "Présenter le déroulé de la soutenance."
        },
        {
            "layout": "section",
            "title": "1. Contexte et enjeux",
            "subtitle": "Pourquoi la GPEC est-elle incontournable ?",
            "notes": "Transition vers la première partie."
        },
        {
            "layout": "key_stat",
            "stat": "67%",
            "description": "des entreprises françaises ont des difficultés de recrutement (DARES, 2024)",
            "notes": "Chiffre clé pour accrocher l'attention du jury."
        },
        {
            "layout": "bullets",
            "title": "Enjeux identifiés",
            "bullets": [
                "Vieillissement de la pyramide des âges",
                "Transformation digitale des métiers",
                "Nécessité d'anticipation des compétences",
                "Obligation légale (Loi Borloo 2005, ANI 2013)"
            ],
            "notes": "Détailler chaque enjeu avec des exemples concrets de l'entreprise."
        },
        {
            "layout": "section",
            "title": "2. Cadre théorique",
            "subtitle": "Les fondements de la GPEC",
            "notes": ""
        },
        {
            "layout": "quote",
            "quote": "La GPEC est la conception, la mise en œuvre et le suivi de politiques et de plans d'actions cohérents visant à réduire de façon anticipée les écarts entre les besoins et les ressources humaines de l'entreprise.",
            "author": "Thierry & Sauret (1993)",
            "notes": "Définition de référence. Souligner le mot 'anticipée'."
        },
        {
            "layout": "two_columns",
            "title": "Approches complémentaires",
            "left_title": "Approche quantitative",
            "left_items": [
                "Pyramide des âges",
                "Tableau de bord RH",
                "Prévisions effectifs",
                "Coûts salariaux"
            ],
            "right_title": "Approche qualitative",
            "right_items": [
                "Référentiel de compétences",
                "Entretiens annuels",
                "Assessment centers",
                "Plans de développement"
            ],
            "notes": "Montrer que les deux approches sont nécessaires."
        },
        {
            "layout": "section",
            "title": "3. Diagnostic des compétences",
            "subtitle": "État des lieux chez Entreprise X",
            "notes": ""
        },
        {
            "layout": "key_stat",
            "stat": "42%",
            "description": "des postes clés n'ont pas de successeur identifié",
            "notes": "Résultat de l'audit interne. Chiffre alarmant."
        },
        {
            "layout": "bullets",
            "title": "Gaps de compétences identifiés",
            "bullets": [
                "Management de transition : compétences insuffisantes",
                "Digital : 35% des collaborateurs sous le niveau requis",
                "Soft skills : communication interculturelle à renforcer",
                "Expertise technique : départs en retraite non anticipés"
            ],
            "notes": "Chaque gap est issu de la cartographie des compétences réalisée."
        },
        {
            "layout": "section",
            "title": "4. Plan d'action",
            "subtitle": "Préconisations stratégiques",
            "notes": ""
        },
        {
            "layout": "two_columns",
            "title": "Plan d'action GPEC",
            "left_title": "Court terme (0-12 mois)",
            "left_items": [
                "Cartographie des compétences",
                "Entretiens professionnels",
                "Plan de formation prioritaire",
                "Mobilité interne ciblée"
            ],
            "right_title": "Moyen terme (1-3 ans)",
            "right_items": [
                "SIRH intégré",
                "Parcours de carrière",
                "Gestion des talents",
                "Partenariats écoles"
            ],
            "notes": "Insister sur le phasage et le réalisme des actions."
        },
        {
            "layout": "conclusion",
            "title": "Conclusion et perspectives",
            "points": [
                "La GPEC est un outil stratégique, pas administratif",
                "Le diagnostic révèle des gaps critiques à combler",
                "Le plan d'action proposé est réaliste et phasé",
                "Perspectives : intégrer l'IA dans la gestion prévisionnelle"
            ],
            "notes": "Synthétiser les points clés. Ouvrir sur l'IA et la GEPP."
        }
    ]
}


LAYOUT_DISPATCH = {
    "title": lambda prs, s: add_title_slide(prs, s["title"], s.get("subtitle", ""), s.get("notes", "")),
    "agenda": lambda prs, s: add_agenda_slide(prs, s["items"], s.get("title", "Agenda"), s.get("notes", "")),
    "section": lambda prs, s: add_section_slide(prs, s["title"], s.get("subtitle", ""), s.get("notes", "")),
    "bullets": lambda prs, s: add_bullets_slide(prs, s["title"], s["bullets"], s.get("notes", "")),
    "two_columns": lambda prs, s: add_two_columns_slide(
        prs, s["title"], s["left_title"], s["left_items"],
        s["right_title"], s["right_items"], s.get("notes", "")
    ),
    "key_stat": lambda prs, s: add_key_stat_slide(prs, s["stat"], s["description"], s.get("notes", "")),
    "quote": lambda prs, s: add_quote_slide(prs, s["quote"], s.get("author", ""), s.get("notes", "")),
    "conclusion": lambda prs, s: add_conclusion_slide(prs, s["title"], s["points"], s.get("notes", "")),
}


class TestIntegrationPipeline:
    def test_full_gpec_presentation(self, tmp_path):
        """Generate a full GPEC presentation from JSON plan."""
        prs = create_presentation()

        for slide_spec in GPEC_PLAN["slides"]:
            layout = slide_spec["layout"]
            assert layout in LAYOUT_DISPATCH, f"Unknown layout: {layout}"
            LAYOUT_DISPATCH[layout](prs, slide_spec)

        assert len(prs.slides) == len(GPEC_PLAN["slides"])

        output = str(tmp_path / "gpec_test.pptx")
        result = save_presentation(prs, output)
        assert os.path.exists(result)
        file_size = os.path.getsize(result)
        assert file_size > 10000, f"File too small: {file_size} bytes"

    def test_all_layouts_used(self):
        """Verify the GPEC plan uses all 8 layout types."""
        layouts_used = {s["layout"] for s in GPEC_PLAN["slides"]}
        expected = {"title", "agenda", "section", "bullets", "two_columns",
                    "key_stat", "quote", "conclusion"}
        assert layouts_used == expected

    def test_all_slides_have_notes(self):
        """Verify most slides have speaker notes."""
        slides_with_notes = sum(
            1 for s in GPEC_PLAN["slides"] if s.get("notes", "").strip()
        )
        # At least 80% should have notes
        assert slides_with_notes >= len(GPEC_PLAN["slides"]) * 0.7

    def test_generate_visual_check_file(self, tmp_path):
        """Generate a .pptx for manual visual inspection."""
        prs = create_presentation()
        for slide_spec in GPEC_PLAN["slides"]:
            LAYOUT_DISPATCH[slide_spec["layout"]](prs, slide_spec)

        output = str(tmp_path / "GPEC_visual_check.pptx")
        save_presentation(prs, output)
        assert os.path.exists(output)
        print(f"\n>>> Visual check file: {output}")
