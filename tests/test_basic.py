"""Basic tests for Read_Excels project"""
import pytest
from pathlib import Path
import sys

# Add parent directory to path
sys.path.insert(0, str(Path(__file__).parent.parent))


def test_imports():
    """Test that main modules can be imported"""
    try:
        import utils.read_data
        import utils.patterns
        import utils.output_excel
        assert True
    except ImportError as e:
        pytest.fail(f"Failed to import modules: {e}")


def test_test_data_exists():
    """Test that test_data directory and files exist"""
    test_data_path = Path(__file__).parent / "test_data"
    assert test_data_path.exists(), "test_data directory not found"

    # Check for sample company data
    company_path = test_data_path / "sample_company"
    assert company_path.exists(), "sample_company directory not found"

    # Check for sample firefighter data
    firefighter_path = test_data_path / "sample_firefighter_survey"
    assert firefighter_path.exists(), "sample_firefighter_survey directory not found"


def test_requirements_file_exists():
    """Test that requirements.txt exists"""
    req_path = Path(__file__).parent.parent / "requirements.txt"
    assert req_path.exists(), "requirements.txt not found"


if __name__ == "__main__":
    pytest.main([__file__, "-v"])
