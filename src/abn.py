"""Australian Business Number (ABN) checksum validation.

Algorithm published by the Australian Business Register:
  1. Subtract 1 from the first digit.
  2. Multiply each digit by its weight (10, 1, 3, 5, ... 19).
  3. Sum the results. The ABN is valid if the total divides evenly by 89.
"""
import numpy as np

WEIGHTS = (10, 1, 3, 5, 7, 9, 11, 13, 15, 17, 19)


def clean_abn(value) -> str:
    if value is None:
        return ""
    text = str(value).strip()
    if text.lower() in ("", "nan", "none"):
        return ""
    return "".join(ch for ch in text if ch.isdigit())


def is_valid_abn(value) -> bool:
    digits = clean_abn(value)
    if len(digits) != 11 or digits[0] == "0":
        return False
    nums = [int(d) for d in digits]
    nums[0] -= 1
    return sum(n * w for n, w in zip(nums, WEIGHTS)) % 89 == 0


def format_abn(digits: str) -> str:
    return f"{digits[:2]} {digits[2:5]} {digits[5:8]} {digits[8:]}"


def random_valid_abn(rng: np.random.Generator) -> str:
    """Pick nine random trailing digits, then solve for the two leading digits."""
    tail = rng.integers(0, 10, size=9)
    tail_sum = sum(int(d) * w for d, w in zip(tail, WEIGHTS[2:]))
    for lead in range(10, 100):
        d1, d2 = divmod(lead, 10)
        if ((d1 - 1) * 10 + d2 + tail_sum) % 89 == 0:
            return format_abn(f"{lead}" + "".join(str(int(d)) for d in tail))
    raise RuntimeError("no valid leading digits found")  # unreachable


def random_invalid_abn(rng: np.random.Generator) -> str:
    """A valid-looking ABN with the final digit altered so the checksum fails."""
    digits = clean_abn(random_valid_abn(rng))
    last = (int(digits[-1]) + int(rng.integers(1, 10))) % 10
    return format_abn(digits[:-1] + str(last))
