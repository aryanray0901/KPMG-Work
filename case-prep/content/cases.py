"""Combines all mock case content into a single CASES list."""
from .cases_part1 import CASES as CASES_1
from .cases_part2 import CASES_PART2

CASES = CASES_1 + CASES_PART2


def get_case(slug):
    for c in CASES:
        if c["slug"] == slug:
            return c
    return None


def cases_by_firm(firm_slug):
    return [c for c in CASES if c["firm_style"] == firm_slug]
