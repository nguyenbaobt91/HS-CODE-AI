# -*- coding: utf-8 -*-
"""Shared Statutory Rules Evaluator (Python) reading statutory_rules.json."""
import json
import os
import re

_BASE_DIR = os.path.dirname(os.path.abspath(__file__))
_RULES_PATH = os.path.join(_BASE_DIR, "statutory_rules.json")
_RULES_CACHE = None


def load_statutory_rules():
    global _RULES_CACHE
    if _RULES_CACHE is None:
        with open(_RULES_PATH, "r", encoding="utf-8") as f:
            _RULES_CACHE = json.load(f).get("rules", [])
    return _RULES_CACHE


def _matches_cond(cond, norm, hn):
    for pat in cond.get("hn_any_regex", []) or []:
        if re.search(pat, hn, re.I):
            break
    else:
        if cond.get("hn_any_regex"):
            return False

    for pat in cond.get("hn_none_regex", []) or []:
        if re.search(pat, hn, re.I):
            return False

    any_pats = cond.get("any_regex", []) or []
    if any_pats and not any(re.search(pat, norm, re.I) for pat in any_pats):
        return False

    for pat in cond.get("all_regex", []) or []:
        if not re.search(pat, norm, re.I):
            return False

    for pat in cond.get("none_regex", []) or []:
        if re.search(pat, norm, re.I):
            return False

    return True


def _format_result(rule, res):
    out = dict(res)
    out.setdefault("missing_attributes", [])
    out.setdefault("assumptions", "")
    out.setdefault("reserve_code", "")
    out["legal_basis"] = rule.get("legal_basis", "")
    out["rule_id"] = rule.get("id", "")
    return out


def match_statutory_rule(hn, norm):
    """Evaluate shared declarative rules in order; returns dict or None."""
    for rule in load_statutory_rules():
        if not _matches_cond(rule, norm, hn):
            continue
        for br in rule.get("branches", []) or []:
            if _matches_cond(br, norm, hn):
                return _format_result(rule, br["result"])
        if "default_result" in rule:
            return _format_result(rule, rule["default_result"])
    return None
