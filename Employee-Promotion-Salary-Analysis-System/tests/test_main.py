import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

from main import (
    check_promotion,
    validate_employee_id,
    validate_email,
    extract_id_number,
)


def test_employee_id():
    assert validate_employee_id("EMP 101")
    assert not validate_employee_id("101")


def test_email():
    assert validate_email("employee@example.com")
    assert not validate_email("employee")


def test_extract_id():
    assert extract_id_number("EMP 101") == "101"


def test_promotion_eligibility():
    assert check_promotion(3, 65, 80, 3, "Completed") == "Eligible"
    assert check_promotion(2, 65, 80, 3, "Completed") == "Not Eligible"
