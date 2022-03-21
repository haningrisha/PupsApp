from reporter import common
from reporter import code_matching
from reporter import utils

SEPARATORS = (' ', '"', '_', ';', '\.', ',', ':')

gazprom_code = {
    common.AccessTypes.PK: [["Согаз Межрегионгаз Прямой доступ ПК-КДЦ", "Межрегионгаз Прямой доступ ПК-КДЦ", 4000971]],
    common.AccessTypes.SMT: [["Согаз Межрегионгаз Прямой доступ ПК-КДЦ", "Межрегионгаз Прямой доступ ПК-КДЦ", 4000971]],
}
abrosia_code = {
    common.AccessTypes.PK: [["АБС СТОМАТОЛОГИЯ Прямой доступ", "ДС№9 к 0618RP137 АБРОССИЯ", 1102]],
    common.AccessTypes.SMT: [["АБС СТОМАТОЛОГИЯ Прямой доступ", "ДС№9 к 0618RP137 АБРОССИЯ", 1102]],
}
nevskoe_code = {
    common.AccessTypes.PK: [["Невское ПКБ СОГАЗ", "ДС№10 к 0618RP137", 4001439]],
    common.AccessTypes.SMT: [["Невское ПКБ СОГАЗ", "ДС№10 к 0618RP137", 4001439]],
}
mariinsky_code = {
    common.AccessTypes.PK: [
        ["Мариинский театр Прямой доступ ПК", "0618RР137 Мариинский театр Прямой доступ ПК", 4000996]],
    common.AccessTypes.SMT: [
        ["Мариинский театр Прямой доступ Согаз КДЦ", "0618RВ138 Мариинский театр Прямой доступ КДЦ", 4000996]],
}
arktika_code = {
    common.AccessTypes.PK: [["СОГАЗ Прямой доступ Арктика ПК", "0618RP137 Арктика ПК", 1102]],
    common.AccessTypes.SMT: [["СОГАЗ Прямой доступ Арктика КДЦ", "0618RВ138 Арктика КДЦ", 1102]],
}
ctss_code = {
    common.AccessTypes.PK: [["ЦТСС прямой контроль", "0618RP137 (ЦТСС)", 4001294]],
    common.AccessTypes.SMT: [["ЦТСС прямой контроль", "0618RP137 (ЦТСС)", 4001294]],
}
ambulatory_code = {
    common.AccessTypes.PK: [["СОГАЗ ПРЯМОЙ ДОСТУП ПК", "0618RР137 ПРЯМОЙ ДОСТУП ПК", 4000971]],
    common.AccessTypes.SMT: [["Прямой доступ СОГАЗ КДЦ", "0618RВ138 Прямой доступ СОГАЗ КДЦ", 4000971]],
}
stomatology_code = {
    common.AccessTypes.PK: [
        ["СОГАЗ СТОМАТОЛОГИЯ Прямой доступ ПК", "0618RР137 СТОМАТОЛОГИЯ ПРЯМОЙ ДОСТУП ПК", 4000971]],
    common.AccessTypes.SMT: [
        ["СОГАЗ СТОМАТОЛОГИЯ Прямой доступ ПК", "0618RР137 СТОМАТОЛОГИЯ ПРЯМОЙ ДОСТУП ПК", 4000971]],
}


def get_sogaz_code_matcher() -> code_matching.CodeMatcher:
    gazprom_filter = code_matching.FilterWorkplaceExact('"ГАЗПРОМ МЕЖРЕГИОНГАЗ САНКТ-ПЕТЕРБУРГ" ООО')
    abrosia_filter = code_matching.FilterIntersection({'"аброссия"', 'аброссия', 'аб'}, SEPARATORS)
    nevskoe_filter = code_matching.FilterInclude('невское', SEPARATORS)
    mariinsky_filter = code_matching.FilterInclude('мариинский', SEPARATORS)
    arktika_filter = code_matching.FilterInclude('арктика', SEPARATORS)
    ctss_filter = code_matching.FilterIntersection({'"цтсс"', 'цтсс'}, SEPARATORS)
    ambulatory_filter = code_matching.FilterIntersection({'амб', 'пнд', 'амбулаторный'}, SEPARATORS)
    stomatology_filter = code_matching.FilterInclude('стоматология', SEPARATORS)

    gazprom_handler = code_matching.CodeHandler(gazprom_code)
    gazprom_handler.add_filter(gazprom_filter)
    abrosia_handler = code_matching.CodeHandler(abrosia_code)
    abrosia_handler.add_filter(abrosia_filter)
    nevskoe_handler = code_matching.CodeHandler(nevskoe_code)
    nevskoe_handler.add_filter(nevskoe_filter)
    mariinsky_handler = code_matching.CodeHandler(mariinsky_code)
    mariinsky_handler.add_filter(mariinsky_filter)
    arktika_handler = code_matching.CodeHandler(arktika_code)
    arktika_handler.add_filter(arktika_filter)
    ctss_handler = code_matching.CodeHandler(ctss_code)
    ctss_handler.add_filter(ctss_filter)
    ambulatory_stomatology_handler = code_matching.CodeHandler(
        utils.combine_codes(ambulatory_code, stomatology_code)
    )
    ambulatory_stomatology_handler.add_filter(ambulatory_filter)
    ambulatory_stomatology_handler.add_filter(stomatology_filter)
    ambulatory_handler = code_matching.CodeHandler(ambulatory_code)
    ambulatory_handler.add_filter(ambulatory_filter)
    stomatology_handler = code_matching.CodeHandler(stomatology_code)
    stomatology_handler.add_filter(stomatology_filter)

    code_matcher = code_matching.CodeMatcher()
    code_matcher.add_handler(gazprom_handler)
    code_matcher.add_handler(abrosia_handler)
    code_matcher.add_handler(nevskoe_handler)
    code_matcher.add_handler(mariinsky_handler)
    code_matcher.add_handler(arktika_handler)
    code_matcher.add_handler(ctss_handler)
    code_matcher.add_handler(ambulatory_stomatology_handler)
    code_matcher.add_handler(ambulatory_handler)
    code_matcher.add_handler(stomatology_handler)
    return code_matcher
