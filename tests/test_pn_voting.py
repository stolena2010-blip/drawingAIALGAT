"""Tests for pn_voting.vote_best_pn — Strategy 3 title-block rejection."""
import pytest
from src.services.extraction.pn_voting import vote_best_pn


class TestVoteBestPnTitleBlock:
    """Ensure title-block text like 'Eng.Mgr' is rejected in favour of Vision."""

    def test_eng_mgr_rejected_vision_wins(self):
        """Regression: Y00529D-00 was overridden by pdfplumber 'Eng.Mgr'."""
        best, source = vote_best_pn(
            vision_pn='Y00529D-00',
            pdfplumber_pn='Eng.Mgr',
            tesseract_pn='',
            filename='unY00529D-00_RevA-_A0_S2_190',
        )
        assert best == 'Y00529D-00'
        assert 'vision' in source

    def test_checker_rejected(self):
        best, source = vote_best_pn(
            vision_pn='ABC12345',
            pdfplumber_pn='Checker',
            tesseract_pn='',
            filename='ABC12345_Rev01.pdf',
        )
        assert best == 'ABC12345'

    def test_approved_rejected(self):
        best, source = vote_best_pn(
            vision_pn='XY99001',
            pdfplumber_pn='Approved',
            tesseract_pn='',
            filename='XY99001.pdf',
        )
        assert best == 'XY99001'

    def test_drawn_by_dot_rejected(self):
        """Title block 'Drawn.By' with dots — purely alpha after stripping."""
        best, source = vote_best_pn(
            vision_pn='PN-123456',
            pdfplumber_pn='Drawn.By',
            tesseract_pn='',
            filename='PN-123456_A0.pdf',
        )
        assert best == 'PN-123456'

    def test_valid_pdfplumber_still_preferred(self):
        """Legit pdfplumber part number with digits should still win."""
        best, source = vote_best_pn(
            vision_pn='WRONG999',
            pdfplumber_pn='AB123456',
            tesseract_pn='',
            filename='AB123456_Rev01.pdf',
        )
        assert best == 'AB123456'
        assert 'pdfplumber' in source or 'filename' in source
