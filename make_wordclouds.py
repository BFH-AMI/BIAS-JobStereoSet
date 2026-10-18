#!/usr/bin/env python3
"""
Creates one word cloud per target from the stereotype completions.

Input: the cleaned CSV produced by the extraction/validation pipeline with
columns full_sentence;target;stereotype;antistereotype;random.

For every target, all stereotype completions are collected, multi-word
completions are split into single words, stopwords and punctuation are
removed, and a word cloud weighted by word frequency is rendered to a PNG
file (one file per target, named after the target). A frequency summary of
the words behind each cloud is written to a text file alongside the images.

Language support: the stopword list is language-specific and lives in
LANGUAGE_CONFIGS. By default the language is detected automatically from the
sentences via stopword statistics; use --language de/nl to force it. To add a new
language, add an entry with its stopword set.

Requirements:
    pip install wordcloud matplotlib

Usage:
    python make_wordclouds.py output_corrected.csv -d wordclouds   # language is auto-detected
"""

import argparse
import csv
import re
import sys
from collections import Counter, defaultdict
from pathlib import Path

from wordcloud import WordCloud


# ---------------------------------------------------------------------------
# Language-specific configuration.
# To support a new language, add an entry to LANGUAGE_CONFIGS below.
# ---------------------------------------------------------------------------

STOPWORDS_DE = {
    "aber", "als", "also", "am", "an", "auch", "auf", "aus", "bei", "beim",
    "bin", "bis", "bisschen", "da", "dabei", "dann", "das", "dass", "dem",
    "den", "der", "des", "die", "diese", "diesem", "diesen", "dieser",
    "dieses", "doch", "dort", "durch", "ein", "eine", "einem", "einen",
    "einer", "eines", "er", "es", "etwas", "für", "ganz", "gegen", "haben",
    "hat", "hier", "ihm", "ihn", "ihr", "ihre", "ihrem", "ihren", "ihrer",
    "im", "in", "ins", "ist", "ja", "kann", "kein", "keine", "keinem",
    "keinen", "keiner", "man", "mehr", "mit", "nach", "nicht", "noch", "nur",
    "ob", "oder", "ohne", "sehr", "sein", "seine", "seinem", "seinen",
    "seiner", "sich", "sie", "sind", "so", "über", "um", "und", "uns",
    "unter", "viel", "vom", "von", "vor", "war", "waren", "was", "wenn",
    "werden", "wie", "wird", "wurde", "zu", "zum", "zur",
}

STOPWORDS_NL = {
    "aan", "al", "alle", "alles", "als", "altijd", "bij", "daar", "dan",
    "dat", "de", "der", "deze", "die", "dit", "door", "dus", "een", "en",
    "er", "erg", "geen", "haar", "had", "hadden", "heb", "hebben", "heeft",
    "heel", "hem", "hen", "het", "hier", "hij", "hoe", "hun", "iets", "ik",
    "in", "is", "ja", "je", "jij", "kan", "kunnen", "maar", "me", "meer",
    "met", "mij", "mijn", "moet", "moeten", "naar", "niet", "niets", "nog",
    "nu", "of", "om", "onder", "ons", "onze", "ook", "op", "over", "te",
    "tegen", "toch", "tot", "u", "uit", "van", "veel", "voor", "waar",
    "wanneer", "waren", "was", "wat", "we", "weer", "wel", "werd", "werden",
    "wie", "wij", "wil", "willen", "wordt", "worden", "zal", "ze", "zeer",
    "zich", "zij", "zijn", "zo", "zonder", "zou", "zouden", "zullen",
}

LANGUAGE_CONFIGS = {
    "de": {"stopwords": STOPWORDS_DE},
    "nl": {"stopwords": STOPWORDS_NL},
    # Add further languages here, e.g.:
    # "en": {"stopwords": STOPWORDS_EN},
}

WORD_RE = re.compile(r"[^\W\d_]+(?:-[^\W\d_]+)*", re.UNICODE)  # words incl. hyphenated


def detect_language(sentences: list) -> str:
    """
    Detects the language by comparing stopword hit rates: the sentences are
    tokenized and, for each configured language, the fraction of tokens found
    in its stopword list is computed; the language with the highest fraction
    wins. Function words are frequent and language-typical (de/het/een/met vs.
    der/die/mit/ist), so this is reliable even for small inputs.
    """
    tokens = [w.lower() for s in sentences for w in WORD_RE.findall(s)]
    if not tokens:
        sys.exit("Cannot detect language: no text found; use --language explicitly.")
    scores = {
        lang: sum(t in cfg["stopwords"] for t in tokens) / len(tokens)
        for lang, cfg in LANGUAGE_CONFIGS.items()
    }
    ranked = sorted(scores.items(), key=lambda kv: kv[1], reverse=True)
    best_lang, best = ranked[0]
    runner_up = ranked[1][1] if len(ranked) > 1 else 0.0
    if best < 0.05 or best < runner_up * 1.2:
        sys.exit(f"Language detection is not confident (scores: "
                 f"{ {k: round(v, 3) for k, v in scores.items()} }); "
                 f"use --language explicitly.")
    print(f"Detected language: {best_lang} "
          f"(stopword hit rates: { {k: round(v, 3) for k, v in scores.items()} })")
    return best_lang


def tokenize(text: str, stopwords: set) -> list:
    """Splits a completion into lowercase words, dropping stopwords and punctuation."""
    return [w.lower() for w in WORD_RE.findall(text) if w.lower() not in stopwords]


def safe_filename(name: str) -> str:
    """Turns a target label into a safe file name."""
    return re.sub(r"[^\w\-]+", "_", name).strip("_") or "unknown"


def main():
    p = argparse.ArgumentParser(description="Create per-target word clouds of stereotype words.")
    p.add_argument("input", help="input file (.csv with the cleaned data)")
    p.add_argument("-d", "--output-dir", default="wordclouds",
                   help="directory for the PNG files (default: wordclouds)")
    p.add_argument("--language", default="auto",
                   choices=["auto"] + sorted(LANGUAGE_CONFIGS),
                   help="language of the data, selects the stopword list; "
                        "'auto' detects it from the sentences (default: auto)")
    p.add_argument("--column", default="stereotype",
                   help="which completion column to visualize (default: stereotype)")
    p.add_argument("--width", type=int, default=1200, help="image width in px (default: 1200)")
    p.add_argument("--height", type=int, default=800, help="image height in px (default: 800)")
    p.add_argument("--delimiter", default=";", help="CSV delimiter (default: ;)")
    args = p.parse_args()

    # Read all rows first, so language detection can look at the sentences.
    with open(args.input, newline="", encoding="utf-8") as f:
        reader = csv.DictReader(f, delimiter=args.delimiter)
        if not reader.fieldnames or args.column not in reader.fieldnames \
                or "target" not in reader.fieldnames:
            sys.exit(f"Input CSV must contain 'target' and '{args.column}' columns "
                     f"(found: {reader.fieldnames})")
        rows = list(reader)

    if args.language == "auto":
        args.language = detect_language(
            [(r.get("full_sentence") or "").strip() for r in rows])
    stopwords = LANGUAGE_CONFIGS[args.language]["stopwords"]

    # Collect word frequencies per target.
    freqs = defaultdict(Counter)
    for row in rows:
        target = (row.get("target") or "").strip()
        completion = (row.get(args.column) or "").strip()
        if not target or not completion:
            continue
        freqs[target].update(tokenize(completion, stopwords))

    if not freqs:
        sys.exit("No data found to visualize.")

    out_dir = Path(args.output_dir)
    out_dir.mkdir(parents=True, exist_ok=True)

    summary_lines = []
    for target in sorted(freqs):
        counter = freqs[target]
        if not counter:
            print(f"Skipping '{target}': no words left after stopword removal.")
            continue
        wc = WordCloud(width=args.width, height=args.height,
                       background_color="white", collocations=False)
        wc.generate_from_frequencies(counter)
        path = out_dir / f"{safe_filename(target)}.png"
        wc.to_file(str(path))
        top = ", ".join(f"{w} ({n})" for w, n in counter.most_common(15))
        summary_lines.append(f"{target} [{sum(counter.values())} words, "
                             f"{len(counter)} unique] -> {path.name}\n    top: {top}")
        print(f"Wrote {path} ({len(counter)} unique words).")

    summary_path = out_dir / "word_frequencies.txt"
    with open(summary_path, "w", encoding="utf-8") as f:
        f.write("\n".join(summary_lines) + "\n")
    print(f"Summary written to {summary_path}.")


if __name__ == "__main__":
    main()