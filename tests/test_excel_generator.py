"""Tests per excel_generator.py."""

import tempfile
import os
from pathlib import Path


def test_ensure_output_dir():
    from excel_generator import ensure_output_dir
    
    with tempfile.TemporaryDirectory() as tmpdir:
        # Patch temporaneo
        import excel_generator
        old_dir = excel_generator.OUTPUT_DIR
        excel_generator.OUTPUT_DIR = Path(tmpdir) / "output"
        
        path = ensure_output_dir()
        
        excel_generator.OUTPUT_DIR = old_dir
        
        assert path.exists()
        assert path.is_dir()


def test_get_template_path_milano():
    from excel_generator import get_template_path
    
    path = get_template_path("milano")
    assert "template-mi" in str(path)


def test_get_template_path_liguria():
    from excel_generator import get_template_path
    
    path = get_template_path("liguria")
    assert "template-lig" in str(path)


def test_get_template_path_invalid():
    from excel_generator import get_template_path
    
    try:
        get_template_path("roma")
        assert False, "Dovrebbe sollevare ValueError"
    except ValueError:
        pass