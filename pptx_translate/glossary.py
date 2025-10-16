import csv
from typing import List, Optional, Tuple


def read_glossary(path: Optional[str]) -> List[Tuple[str, str]]:
    pairs: List[Tuple[str, str]] = []
    if not path:
        return pairs
    with open(path, newline="", encoding="utf-8-sig") as f:
        reader = csv.reader(f)
        for row in reader:
            if len(row) >= 2:
                src, tgt = row[0].strip(), row[1].strip()
                if src:
                    pairs.append((src, tgt))
    return pairs


def apply_glossary(text: str, glossary: List[Tuple[str, str]]) -> str:
    for src, tgt in glossary:
        if src:
            text = text.replace(src, tgt)
    return text
