#!/usr/bin/env python3
"""
Extracts a CSV from an .xlsx file containing BLANK sentence groups.

Input file structure:
- The first three rows are ignored (configurable via --skip-rows).
- Columns A-H are ignored.
- After that, groups of 4 columns each:
    [sentence with BLANK, stereotype variant, antistereotype variant, random variant]
- For each row and group, one row is written to the output CSV:
    full_sentence;target;stereotype;antistereotype;random

Cleanup: placeholder spelling variants in the template cell (for German:
bland, Blank, blank, leer, Leer, Leere, LEERE, LEER, ...) are normalized to
BLANK before processing (case-insensitive whole-word replacement).

full_sentence = the complete template sentence containing BLANK.
target = the best-matching entry from the language's target list (or provide
         your own list via --targets-file). Fuzzy matching allows small
         deviations such as plural forms and dropped articles. Matches must
         respect word boundaries and feminine plural forms are generated
         explicitly, so 'Mitarbeiterinnen' always maps to a '...Mitarbeiterin'
         target and never to a masculine one.
stereotype/antistereotype/random = the text that replaced BLANK in the
         respective variant sentence, parsed by aligning the variant against
         the template around the BLANK position.

Language support: all language-specific parts (placeholder spellings, target
list, articles, pluralization rules) live in LANGUAGE_CONFIGS. Select the
language via --language (default: de). To add a new language, add an entry
to LANGUAGE_CONFIGS with its own placeholder pattern, targets, articles and
pluralization function.

Matching tolerance: small deviations (default: edit distance of 2 characters
for variant alignment, 4 for target detection) are accepted. Missing trailing
punctuation is always allowed. Anything that cannot be matched is written to
the log file and the corresponding CSV cell is left empty.

Output order: all rows belonging to the same target are grouped together,
in the order the targets appear in the target list. Rows without a matched
target are appended at the end.

Usage:
    python extract_blanks.py input.xlsx -o output.csv --log problems.log --language de
"""

import argparse
import csv
import re
import sys

from openpyxl import load_workbook

TRAIL_PUNCT = " \t\n.!?;:,"


# ---------------------------------------------------------------------------
# Language-specific configuration.
# To support a new language, add an entry to LANGUAGE_CONFIGS below.
# ---------------------------------------------------------------------------

IT_ARTICLE_PLURALS = {"il": "i", "lo": "gli", "la": "le", "un": "i", "una": "le"}
IT_INVARIABLE = {"gay", "queer", "trans", "cis", "non"}
IT_IRREGULAR_PLURALS = {"uomo": "uomini"}


def pluralize_it(form: str) -> str:
    """Italian pluralization for the target phrases (input is lowercase).
    Gender is taken from the leading article (or, for elided l', from the
    final adjective ending) so that e.g. 'il collega' -> 'i colleghi' but
    'la collega' -> 'le colleghe'."""
    words = form.split()
    first = words[0] if words else ""
    if first in ("la", "una", "le"):
        masculine = False
    elif first in ("il", "lo", "un", "i", "gli"):
        masculine = True
    else:
        masculine = not (words and words[-1].endswith("a"))
    out = []
    for w in words:
        prefix = ""
        if w.startswith(("l'", "l’")):
            prefix, w = ("gli " if masculine else "le "), w[2:]
        if w in IT_ARTICLE_PLURALS:
            out.append(prefix + IT_ARTICLE_PLURALS[w])
        elif w in IT_INVARIABLE or not w:
            out.append(prefix + w)
        elif w in IT_IRREGULAR_PLURALS:
            out.append(prefix + IT_IRREGULAR_PLURALS[w])
        elif w.endswith(("ca", "ga")):
            out.append(prefix + w[:-1] + ("hi" if masculine else "he"))
        elif w.endswith(("o", "e")):
            out.append(prefix + w[:-1] + "i")
        elif w.endswith("a"):
            out.append(prefix + w[:-1] + "e")
        else:
            out.append(prefix + w)
    return " ".join(out)


TARGETS_IT = [
    "Il dipendente rumeno",
    "La dipendente rumena",
    "Il dipendente italiano",
    "La dipendente italiana",
    "La madre rumena",
    "La madre italiana",
    "Il padre rumeno",
    "Il padre italiano",
    "Il collega indiano",
    "La collega indiana",
    "Il collega italiano",
    "La collega italiana",
    "Il bambino indiano",
    "La bambina indiana",
    "Il bambino italiano",
    "La bambina italiana",
    "L'adolescente indiano",
    "L'adolescente indiana",
    "L'adolescente italiano",
    "L'adolescente italiana",
    "Il collaboratore gay",
    "La collaboratrice lesbica",
    "Il collaboratore eterosessuale",
    "La collaboratrice eterosessuale",
    "La collega lesbica",
    "La collega eterosessuale",
    "Il collega eterosessuale",
    "Il vicino queer",
    "La vicina queer",
    "Il vicino eterosessuale",
    "La vicina eterosessuale",
    "Un uomo trans",
    "Una donna trans",
    "Un uomo cis",
    "Una donna cis",
    "La persona lavoratrice non binaria",
]

def pluralize_de(form: str) -> str:
    """German pluralization rules for the target nouns (input is lowercase)."""
    words = []
    for w in form.split():
        if w.endswith("in") and len(w) > 3:
            words.append(w + "nen")      # Mitarbeiterin -> Mitarbeiterinnen
        elif w == "mutter":
            words.append("mütter")
        elif w == "vater":
            words.append("väter")
        elif w.endswith("kollege"):
            words.append(w + "n")        # Kollege -> Kollegen
        else:
            words.append(w)
    return " ".join(words)


TARGETS_DE = [
    "berufstätige Mutter",
    "berufstätiger Vater",
    "kinderloser Mitarbeiter",
    "kinderlose Mitarbeiterin",
    "Mitarbeiter mit Familie",
    "Mitarbeiterin mit Familie",
    "Mitarbeiter ohne Familie",
    "Mitarbeiterin ohne Familie",
    "Der junge Mitarbeiter",
    "Der ältere Mitarbeiter",
    "Die junge Mitarbeiterin",
    "Die ältere Mitarbeiterin",
    "Der unerfahrene Arbeitskollege",
    "Der erfahrene Arbeitskollege",
    "Die unerfahrene Arbeitskollegin",
    "Die erfahrene Arbeitskollegin",
    "Der Schweizer Mitarbeiter",
    "Der deutsche Mitarbeiter",
    "Der Mitarbeiter mit Migrationshintergrund",
    "Der Gastarbeiter",
    "Die Schweizer Mitarbeiterin",
    "Die deutsche Mitarbeiterin",
    "Die Mitarbeiterin mit Migrationshintergrund",
    "Die Gastarbeiterin",
    "Der einheimische Mitarbeiter",
    "Die einheimische Mitarbeiterin",
    "Die Grenzgängerin",
    "Der Grenzgänger",
]

# Dutch noun plurals for the target vocabulary (lowercase singular -> plural).
PLURALS_NL = {
    "collega": "collega's",
    "buurman": "buurmannen",
    "buurvrouw": "buurvrouwen",
    "buur": "buren",
    "moeder": "moeders",
    "vader": "vaders",
    "medewerker": "medewerkers",
    "werknemer": "werknemers",
    "man": "mannen",
    "vrouw": "vrouwen",
}


def pluralize_nl(form: str) -> str:
    """Dutch pluralization rules for the target nouns (input is lowercase)."""
    return " ".join(PLURALS_NL.get(w, w) for w in form.split())


TARGETS_NL = [
    "de collega met een beperking",
    "De collega zonder beperking",
    "De buurman met een beperking",
    "De buurvrouw met een beperking",
    "De buurman zonder beperking",
    "De buurvrouw zonder beperking",
    "Een moeder met een beperking",
    "Een vader met een beperking",
    "Een vader zonder beperking",
    "Een moeder zonder beperking",
    "De homoseksuele medewerker",
    "De heteroseksuele medewerker",
    "De lesbische collega",
    "De heteroseksuele collega",
    "De queer buur",
    "De heteroseksuele buur",
    "Een trans man",
    "Een trans vrouw",
    "Een cis man",
    "Een cis vrouw",
    "De non-binaire werknemer",
    "De binaire werknemer",
]

def pluralize_tr(form: str) -> str:
    """Turkish pluralization for the head noun (last word): suffix -lar/-ler
    chosen by vowel harmony (input is lowercase)."""
    words = form.split()
    if not words:
        return form
    last = words[-1]
    suffix = "lar"
    for ch in reversed(last):
        if ch in "aıou":       # back vowels -> -lar
            suffix = "lar"
            break
        if ch in "eiöü":       # front vowels -> -ler
            suffix = "ler"
            break
    words[-1] = last + suffix
    return " ".join(words)



TARGETS_TR = [
    "Çalışan anne",
    "Çalışan baba",
    "Çocuğu olmayan çalışan",
    "Çocuğu olan çalışan",
    "Çocuğu olmayan çalışan meslektaş",
    "Birçok hobisi olan çalışan meslektaş",
    "Hobisi olmayan çalışan meslektaş",
    "Zengin çalışan",
    "Fakir çalışan",
    "Askerlik yükümlülüğü olan çalışan",
    "Askerlik yükümlülüğü olmayan çalışan",
]

# Norwegian (Bokmål) plural/inflection forms for the target vocabulary
# (lowercase word -> plural form; articles den/det -> de, en/et/ei dropped
# because the indefinite plural has no article).
PLURALS_NO = {
    "den": "de", "det": "de", "en": "", "et": "", "ei": "",
    "kollegaen": "kollegaene", "far": "fedre", "faren": "fedrene",
    "mor": "mødre", "moren": "mødrene", "tenåring": "tenåringer",
    "kvinnen": "kvinnene", "naboen": "naboene", "prest": "prester",
    "imam": "imamer", "norsk": "norske", "iransk": "iranske",
    "kristen": "kristne", "muslimsk": "muslimske",
}


def pluralize_no(form: str) -> str:
    """Norwegian pluralization for the target phrases (input is lowercase).
    Nouns already carrying the suffixed definite article get the definite
    plural (kollegaen -> kollegaene); 'barn' and adjectives already in the
    definite form (-e) are unchanged."""
    words = [PLURALS_NO.get(w, w) for w in form.split()]
    return " ".join(w for w in words if w)


TARGETS_NO = [
    "Den norske kollegaen",
    "Den iranske kollegaen",
    "En norsk far",
    "En iransk far",
    "En norsk mor",
    "En iransk mor",
    "Et norsk barn",
    "Et iransk barn",
    "En norsk tenåring",
    "En iransk tenåring",
    "Den kristne kollegaen",
    "Den muslimske kollegaen",
    "Kvinnen med hijab",
    "Kvinnen med det lange håret",
    "Naboen med Koranen",
    "Naboen med Bibelen",
    "En prest",
    "En imam",
    "En muslimsk mor",
    "En kristen mor",
    "En muslimsk far",
    "En kristen far",
    "Den ansatte med barn",
    "Den ansatte uten barn",
    "Den arbeidende moren",
    "Den arbeidende faren",
    "Den barnløse kollegaen",
    "Den barnløse ansatte",
    "Den ansatte uten familie",
    "Den ansatte med familie",
]

LANGUAGE_CONFIGS = {
    "de": {
        # Placeholder spelling variants normalized to BLANK (whole words,
        # case-insensitive): covers bland, Blank, blank, leer, Leere, LEERE, ...
        "placeholder_pattern": r"\b(?:blank|bland|leere|leer)\b",
        "targets": TARGETS_DE,
        # Leading articles that may be dropped in the sentences (lowercase,
        # with trailing space).
        "articles": ("der ", "die ", "das "),
        "pluralize": pluralize_de,
    },
    "nl": {
        # Covers blank, Blank, bland, leeg, LEEG, lege, ...
        "placeholder_pattern": r"\b(?:blank|bland|leeg|lege)\b",
        "targets": TARGETS_NL,
        "articles": ("de ", "het ", "een "),
        "pluralize": pluralize_nl,
    },
    "tr": {
        # Covers blank, Blank, BLANK, bland, boş, BOŞ, boşluk — optionally
        # with an attached apostrophe suffix such as BLANK'tir / BLANK'tır /
        # BLANK'dir. The suffix is absorbed into the placeholder, so the
        # variant's fully inflected word is extracted as the completion
        # (e.g. "Çalışan anne BLANK'tir." + "Çalışan anne yorgundur."
        # -> completion "yorgundur").
        "placeholder_pattern": r"\b(?:blank|bland|boşluk|boş)(?:['’ʼ`´]\w+)?\b",
        "targets": TARGETS_TR,
        # Turkish has no articles; a preceding "bir" needs no handling
        # because substring matching has a free start anyway.
        "articles": (),
        "pluralize": pluralize_tr,
    },
    "it": {
        # Covers blank, Blank, BLANK, bland, vuoto, vuota, VUOTO, ...
        "placeholder_pattern": r"\b(?:blank|bland|vuoto|vuota)\b",
        "targets": TARGETS_IT,
        # Includes the elided article in both apostrophe spellings (l', l’),
        # without trailing space because it attaches directly to the noun.
        "articles": ("il ", "lo ", "la ", "i ", "gli ", "le ", "un ", "una ",
                     "l'", "l’"),
        "pluralize": pluralize_it,
    },
    "no": {
        # Covers blank, Blank, BLANK, bland, tom, TOM, tomt, tomme, ...
        "placeholder_pattern": r"\b(?:blank|bland|tomt|tomme|tom)\b",
        "targets": TARGETS_NO,
        "articles": ("den ", "det ", "de ", "en ", "et ", "ei "),
        "pluralize": pluralize_no,
    }
    # Add further languages here, e.g.:
    # "en": {
    #     "placeholder_pattern": r"\b(?:blank|empty)\b",
    #     "targets": TARGETS_EN,
    #     "articles": ("the ", "a ", "an "),
    #     "pluralize": pluralize_en,
    # },
}


def get_language_config(language: str) -> dict:
    if language not in LANGUAGE_CONFIGS:
        supported = ", ".join(sorted(LANGUAGE_CONFIGS))
        sys.exit(f"Unsupported language '{language}'. Supported: {supported}")
    cfg = dict(LANGUAGE_CONFIGS[language])
    cfg["placeholder_re"] = re.compile(cfg["placeholder_pattern"], re.IGNORECASE)
    return cfg


# ---------------------------------------------------------------------------
# Language-independent matching logic.
# ---------------------------------------------------------------------------

def normalize_placeholder(template: str, config: dict) -> str:
    """Replaces all placeholder spelling variants in a template with BLANK."""
    return config["placeholder_re"].sub("BLANK", template)


def levenshtein(a: str, b: str, cutoff: int) -> int:
    """Edit distance with early exit once it definitely exceeds cutoff."""
    if abs(len(a) - len(b)) > cutoff:
        return cutoff + 1
    prev = list(range(len(b) + 1))
    for i, ca in enumerate(a, 1):
        cur = [i]
        best = i
        for j, cb in enumerate(b, 1):
            cost = 0 if ca == cb else 1
            v = min(prev[j] + 1, cur[j - 1] + 1, prev[j - 1] + cost)
            cur.append(v)
            best = min(best, v)
        if best > cutoff:
            return cutoff + 1
        prev = cur
    return prev[-1]


def substring_distance(pattern: str, text: str) -> int:
    """
    Minimal edit distance between pattern and any substring of text
    (approximate substring matching, Sellers algorithm), where the matched
    substring must start and end at word boundaries. This prevents e.g.
    'Mitarbeiter' from matching inside 'Mitarbeiterinnen'.
    """
    if not pattern:
        return 0
    n = len(text)
    INF = 10 ** 9

    def is_word_char(c: str) -> bool:
        # Hyphens are word-internal: 'non-binaire' is one word, so 'binaire'
        # cannot match inside it.
        return c.isalnum() or c == "-"

    def is_boundary(j: int) -> bool:
        """True if position j (between text[j-1] and text[j]) is a word boundary."""
        if j == 0 or j == n:
            return True
        return not (is_word_char(text[j - 1]) and is_word_char(text[j]))

    prev = [0 if is_boundary(j) else INF for j in range(n + 1)]  # free start at boundaries
    for i, cp in enumerate(pattern, 1):
        cur = [i]
        for j in range(1, n + 1):
            cost = 0 if cp == text[j - 1] else 1
            cur.append(min(prev[j] + 1, cur[j - 1] + 1, prev[j - 1] + cost))
        prev = cur
    return min(prev[j] for j in range(n + 1) if is_boundary(j))  # free end at boundaries


def target_forms(target: str, config: dict):
    """
    Generates lowercase matching forms of a target: the original, the plural
    form (per language rules), and article-stripped variants of each. This
    keeps feminine forms ('...in'/'...innen') tied to feminine targets.
    """
    forms = {target.lower()}
    forms.add(config["pluralize"](target.lower()))
    for f in list(forms):
        for article in config["articles"]:
            if f.startswith(article):
                forms.add(f[len(article):])
    return forms


def find_target(sentence: str, targets, config: dict, tolerance: int):
    """
    Finds the best-matching target within the sentence. Each target is matched
    via all of its generated forms (see target_forms); the target's distance is
    the minimum over its forms. If several targets tie on distance, the longest
    (most specific) one wins, so a target that is a prefix of another (e.g.
    'Çocuğu olmayan çalışan' vs 'Çocuğu olmayan çalışan meslektaş') is only
    chosen when the longer one does not match. Returns (target_index or None,
    reason).
    """
    s = sentence.strip().lower()
    candidates = []  # (distance, index)
    for idx, target in enumerate(targets):
        d = min(substring_distance(form, s) for form in target_forms(target, config))
        if d <= tolerance:
            candidates.append((d, idx))
    if not candidates:
        return None, f"no target matched within distance {tolerance}"
    best_dist = min(d for d, _ in candidates)
    tied = [idx for d, idx in candidates if d == best_dist]
    longest = max(len(targets[i]) for i in tied)
    tied = [i for i in tied if len(targets[i]) == longest]
    best_idx = tied[0]
    if len(tied) > 1:
        names = [targets[i] for i in tied]
        return best_idx, (f"ambiguous target match (distance {best_dist}): "
                          f"{names}; picked '{targets[best_idx]}'")
    return best_idx, ""


def extract_replacement(template: str, variant: str, tolerance: int):
    """
    Parses out the text that replaced BLANK, by aligning the variant against
    the template's prefix (before BLANK) and suffix (after BLANK) with
    fuzzy tolerance. Missing trailing punctuation is always allowed.
    Returns (replacement or None, reason).
    """
    prefix, _, suffix = template.partition("BLANK")
    prefix = prefix.rstrip()            # tolerate whitespace before BLANK
    suffix = suffix.strip(TRAIL_PUNCT)  # allow missing punctuation
    v_core = variant.strip().strip(TRAIL_PUNCT)

    # Align prefix: best cut position around the expected length
    n = len(prefix)
    best_pre_k, best_pre = 0, tolerance + 1
    for k in range(max(0, n - tolerance), min(len(v_core), n + tolerance) + 1):
        d = levenshtein(prefix, v_core[:k], tolerance)
        if d < best_pre:
            best_pre, best_pre_k = d, k
    if best_pre > tolerance:
        return None, f"sentence start deviates (distance > {tolerance}): expected '{prefix[:60]}…'"

    # Align suffix: best cut position from the end
    best_suf_k, best_suf = 0, 0 if not suffix else tolerance + 1
    if suffix:
        m = len(suffix)
        for k in range(max(0, m - tolerance), min(len(v_core), m + tolerance) + 1):
            tail = v_core[len(v_core) - k:] if k else ""
            d = levenshtein(suffix, tail, tolerance)
            if d < best_suf:
                best_suf, best_suf_k = d, k
        if best_suf > tolerance:
            return None, f"sentence end deviates (distance > {tolerance}): expected '…{suffix[-60:]}'"

    end = len(v_core) - best_suf_k

    def is_word_char(c: str) -> bool:
        return c.isalnum() or c in "-'’"

    # If a cut falls inside a word (e.g. the variant inflected the subject:
    # template 'anne', variant 'anneler' or 'çalışanın'), snap it to the next
    # word boundary so leftover inflection is not extracted as completion.
    while 0 < best_pre_k < len(v_core) \
            and is_word_char(v_core[best_pre_k - 1]) and is_word_char(v_core[best_pre_k]):
        best_pre_k += 1
    while best_pre_k < end < len(v_core) \
            and is_word_char(v_core[end - 1]) and is_word_char(v_core[end]):
        end -= 1

    if end < best_pre_k:
        return None, "prefix and suffix alignment overlap; variant too short"
    replacement = v_core[best_pre_k:end].strip().strip(TRAIL_PUNCT)

    if not replacement:
        return None, "empty replacement after alignment"
    return replacement, ""


def cell_str(value) -> str:
    if value is None:
        return ""
    return str(value).strip()


def load_targets(path):
    with open(path, encoding="utf-8") as f:
        return [line.strip() for line in f if line.strip()]


def main():
    p = argparse.ArgumentParser(description="Extract BLANK sentence groups from xlsx to CSV.")
    p.add_argument("input", help="input file (.xlsx)")
    p.add_argument("-o", "--output", default="output.csv", help="output file (.csv)")
    p.add_argument("--log", default="problems.log", help="log file for problems")
    p.add_argument("--language", default="de", choices=sorted(LANGUAGE_CONFIGS),
                   help="language of the input data (default: de)")
    p.add_argument("--sheet", default=None, help="worksheet name (default: first sheet)")
    p.add_argument("--start-col", type=int, default=9,
                   help="first column of the groups, 1-based (default: 9 = column I)")
    p.add_argument("--skip-rows", type=int, default=3,
                   help="number of leading rows to ignore (default: 3)")
    p.add_argument("--tolerance", type=int, default=2,
                   help="allowed edit distance for variant alignment (default: 2)")
    p.add_argument("--target-tolerance", type=int, default=4,
                   help="allowed edit distance for target detection (default: 4)")
    p.add_argument("--targets-file", default=None,
                   help="optional text file with one target per line "
                        "(overrides the language's built-in list)")
    p.add_argument("--delimiter", default=";", help="CSV delimiter (default: ;)")
    args = p.parse_args()

    config = get_language_config(args.language)
    targets = load_targets(args.targets_file) if args.targets_file else config["targets"]

    wb = load_workbook(args.input, read_only=True, data_only=True)
    ws = wb[args.sheet] if args.sheet else wb.active

    problems = []
    rows_out = []  # (target_index or large number, row)
    n_groups_total = 0
    unmatched_rank = len(targets)  # unmatched targets sort to the end

    first_data_row = args.skip_rows + 1

    for row_idx, row in enumerate(ws.iter_rows(values_only=True), start=1):
        if row_idx < first_data_row:
            continue
        cells = [cell_str(v) for v in row]
        col = args.start_col - 1  # 0-based

        while col < len(cells):
            group = cells[col:col + 4]
            group += [""] * (4 - len(group))
            template, v_stereo, v_anti, v_random = group
            excel_col = col + 1  # 1-based, for messages
            col += 4

            if not any(group):
                continue  # skip completely empty group
            n_groups_total += 1

            def log(msg):
                problems.append(f"Row {row_idx}, columns {excel_col}-{excel_col + 3}: {msg}")

            if not template:
                log("template cell is empty")
                continue
            template = normalize_placeholder(template, config)
            if "BLANK" not in template:
                log(f"template does not contain BLANK: '{template}'")
                continue

            target_idx, reason = find_target(template, targets, config, args.target_tolerance)
            if target_idx is None:
                log(f"{reason} | template: '{template}'")
                target_name, rank = "", unmatched_rank
            else:
                if reason:  # ambiguity warning
                    log(f"{reason} | template: '{template}'")
                target_name, rank = targets[target_idx], target_idx

            values = []
            for name, variant in (("stereotype", v_stereo),
                                  ("antistereotype", v_anti),
                                  ("random", v_random)):
                if not variant:
                    log(f"variant '{name}' is empty")
                    values.append("")
                    continue
                replacement, reason = extract_replacement(template, variant, args.tolerance)
                if replacement is None:
                    log(f"variant '{name}' does not match template: {reason} | "
                        f"template: '{template}' | variant: '{variant}'")
                    values.append("")
                    continue
                values.append(replacement)

            rows_out.append((rank, [template, target_name] + values))

    # Group all rows of the same target together (stable sort keeps input order within a target)
    rows_out.sort(key=lambda item: item[0])

    with open(args.output, "w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f, delimiter=args.delimiter)
        writer.writerow(["full_sentence", "target", "stereotype", "antistereotype", "random"])
        writer.writerows(r for _, r in rows_out)

    with open(args.log, "w", encoding="utf-8") as f:
        if problems:
            f.write("\n".join(problems) + "\n")
        else:
            f.write("No problems found.\n")

    print(f"{len(rows_out)} rows written to {args.output} "
          f"({n_groups_total} groups processed).")
    if problems:
        print(f"{len(problems)} problem(s) logged to {args.log}.", file=sys.stderr)


if __name__ == "__main__":
    main()