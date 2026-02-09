"""Unit tests for slide_engine — each layout individually."""

import os
import sys
import pytest

# Add parent directory to path
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


@pytest.fixture
def prs():
    return create_presentation()


class TestCreatePresentation:
    def test_creates_presentation(self):
        prs = create_presentation()
        assert prs is not None
        assert len(prs.slides) == 0

    def test_slide_dimensions_16_9(self):
        from pptx.util import Inches
        prs = create_presentation()
        assert prs.slide_width == Inches(13.333)
        assert prs.slide_height == Inches(7.5)


class TestSavePresentation:
    def test_save_adds_extension(self, prs, tmp_path):
        filepath = str(tmp_path / "test")
        result = save_presentation(prs, filepath)
        assert result.endswith(".pptx")
        assert os.path.exists(result)

    def test_save_keeps_extension(self, prs, tmp_path):
        filepath = str(tmp_path / "test.pptx")
        result = save_presentation(prs, filepath)
        assert result == filepath
        assert os.path.exists(result)


class TestTitleSlide:
    def test_basic(self, prs):
        slide = add_title_slide(prs, "Mon Titre", "Sous-titre")
        assert len(prs.slides) == 1
        texts = [shape.text for shape in slide.shapes if shape.has_text_frame]
        assert "Mon Titre" in texts

    def test_without_subtitle(self, prs):
        slide = add_title_slide(prs, "Titre seul")
        assert len(prs.slides) == 1

    def test_with_notes(self, prs):
        slide = add_title_slide(prs, "Titre", notes="Notes du présentateur")
        notes_text = slide.notes_slide.notes_text_frame.text
        assert "Notes du présentateur" in notes_text

    def test_unicode(self, prs):
        slide = add_title_slide(prs, "Données RH : évaluation & développement")
        texts = [shape.text for shape in slide.shapes if shape.has_text_frame]
        assert any("évaluation" in t for t in texts)


class TestAgendaSlide:
    def test_basic(self, prs):
        items = ["Introduction", "Diagnostic", "Préconisations"]
        slide = add_agenda_slide(prs, items)
        assert len(prs.slides) == 1

    def test_custom_title(self, prs):
        slide = add_agenda_slide(prs, ["A", "B"], title="Plan")
        texts = [shape.text for shape in slide.shapes if shape.has_text_frame]
        assert "Plan" in texts

    def test_with_notes(self, prs):
        slide = add_agenda_slide(prs, ["A"], notes="Rappeler le contexte")
        assert "Rappeler le contexte" in slide.notes_slide.notes_text_frame.text


class TestSectionSlide:
    def test_basic(self, prs):
        slide = add_section_slide(prs, "Partie I")
        assert len(prs.slides) == 1

    def test_with_subtitle(self, prs):
        slide = add_section_slide(prs, "Partie I", "Cadre théorique")
        texts = [shape.text for shape in slide.shapes if shape.has_text_frame]
        assert "Cadre théorique" in texts


class TestBulletsSlide:
    def test_basic(self, prs):
        slide = add_bullets_slide(prs, "Points clés", ["Point A", "Point B", "Point C"])
        assert len(prs.slides) == 1

    def test_unicode_bullets(self, prs):
        slide = add_bullets_slide(prs, "Résumé", ["Première étape", "Deuxième étape"])
        assert len(prs.slides) == 1


class TestTwoColumnsSlide:
    def test_basic(self, prs):
        slide = add_two_columns_slide(
            prs,
            title="Comparaison",
            left_title="Avant",
            left_items=["Item A", "Item B"],
            right_title="Après",
            right_items=["Item C", "Item D"],
        )
        assert len(prs.slides) == 1

    def test_with_notes(self, prs):
        slide = add_two_columns_slide(
            prs,
            title="Test",
            left_title="L", left_items=["a"],
            right_title="R", right_items=["b"],
            notes="Comparer les deux approches",
        )
        assert "Comparer" in slide.notes_slide.notes_text_frame.text


class TestKeyStatSlide:
    def test_basic(self, prs):
        slide = add_key_stat_slide(prs, "78%", "des salariés satisfaits")
        assert len(prs.slides) == 1

    def test_stat_text(self, prs):
        slide = add_key_stat_slide(prs, "3.2M€", "Budget formation annuel")
        texts = [shape.text for shape in slide.shapes if shape.has_text_frame]
        assert "3.2M€" in texts


class TestQuoteSlide:
    def test_basic(self, prs):
        slide = add_quote_slide(prs, "La GRH est un levier stratégique.")
        assert len(prs.slides) == 1

    def test_with_author(self, prs):
        slide = add_quote_slide(
            prs,
            "Le capital humain est la première richesse.",
            author="Jean-Marie Peretti",
        )
        texts = [shape.text for shape in slide.shapes if shape.has_text_frame]
        assert any("Peretti" in t for t in texts)


class TestConclusionSlide:
    def test_basic(self, prs):
        slide = add_conclusion_slide(
            prs,
            "Conclusion",
            ["Mettre en place un SIRH", "Former les managers", "Suivre les KPIs"],
        )
        assert len(prs.slides) == 1

    def test_with_notes(self, prs):
        slide = add_conclusion_slide(
            prs, "À retenir", ["Point 1"], notes="Ouvrir sur les perspectives"
        )
        assert "perspectives" in slide.notes_slide.notes_text_frame.text


class TestSpeakerNotes:
    def test_empty_notes_no_crash(self, prs):
        slide = add_title_slide(prs, "Test", notes="")
        # Should not crash with empty notes

    def test_long_notes(self, prs):
        long_notes = "Note très longue. " * 100
        slide = add_title_slide(prs, "Test", notes=long_notes)
        assert len(slide.notes_slide.notes_text_frame.text) > 100

    def test_unicode_notes(self, prs):
        slide = add_bullets_slide(
            prs, "Test", ["A"],
            notes="Référence : Thévenet (2015), « Fonctions RH » — 4ème édition"
        )
        assert "Thévenet" in slide.notes_slide.notes_text_frame.text
