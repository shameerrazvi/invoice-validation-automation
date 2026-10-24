import numpy as np
import pytest

from src.abn import is_valid_abn, random_invalid_abn, random_valid_abn


@pytest.mark.parametrize("abn", ["51 824 753 556", "51824753556", " 51-824-753-556 "])
def test_known_valid_abn(abn):
    # ABN of the Australian Taxation Office
    assert is_valid_abn(abn)


@pytest.mark.parametrize("abn", ["51 824 753 557", "1234567890", "", None, "nan", "01 824 753 556"])
def test_invalid_abn(abn):
    assert not is_valid_abn(abn)


def test_generated_abns_pass_and_fail_as_expected():
    rng = np.random.default_rng(0)
    assert all(is_valid_abn(random_valid_abn(rng)) for _ in range(500))
    assert not any(is_valid_abn(random_invalid_abn(rng)) for _ in range(500))
