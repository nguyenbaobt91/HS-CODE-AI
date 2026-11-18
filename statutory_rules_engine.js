// Shared Statutory Rules Evaluator (Browser JS) reading window.STATUTORY_RULES
(function (global) {
    function testListAny(pats, text) {
        if (!pats || pats.length === 0) return true;
        return pats.some(p => new RegExp(p, 'i').test(text));
    }
    function testListAll(pats, text) {
        if (!pats || pats.length === 0) return true;
        return pats.every(p => new RegExp(p, 'i').test(text));
    }
    function testListNone(pats, text) {
        if (!pats || pats.length === 0) return true;
        return !pats.some(p => new RegExp(p, 'i').test(text));
    }
    function matchesCond(cond, norm, hn) {
        if (cond.hn_any_regex && cond.hn_any_regex.length > 0 && !testListAny(cond.hn_any_regex, hn)) return false;
        if (cond.hn_none_regex && cond.hn_none_regex.length > 0 && !testListNone(cond.hn_none_regex, hn)) return false;
        if (!testListAny(cond.any_regex, norm)) return false;
        if (!testListAll(cond.all_regex, norm)) return false;
        if (!testListNone(cond.none_regex, norm)) return false;
        return true;
    }
    function formatResult(rule, res) {
        return Object.assign({
            missing_attributes: [],
            assumptions: '',
            reserve_code: '',
            legal_basis: rule.legal_basis || '',
            rule_id: rule.id || ''
        }, res);
    }
    global.StatutoryRulesEngine = {
        match: function (hn, norm) {
            const rules = (global.STATUTORY_RULES && global.STATUTORY_RULES.rules) || [];
            for (const rule of rules) {
                if (!matchesCond(rule, norm, hn)) continue;
                for (const br of (rule.branches || [])) {
                    if (matchesCond(br, norm, hn)) {
                        return formatResult(rule, br.result);
                    }
                }
                if (rule.default_result) {
                    return formatResult(rule, rule.default_result);
                }
            }
            return null;
        }
    };
})(window);
