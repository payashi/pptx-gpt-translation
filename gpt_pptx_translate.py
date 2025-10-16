#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Translate PPTX with GPT + python-pptx.

- Batches by slide and uses neighbor/summary context to improve coherence.
- Optional glossary enforcement (post-pass string replacement).
- Optional speaker notes and masters/layouts.
- Supports OpenAI Chat Completions API (legacy) and Responses API backends.

Requires:
  pip install openai python-pptx tqdm tenacity python-dotenv
  export OPENAI_API_KEY=...

Example:
  python gpt_pptx_translate.py input.pptx output_ja.pptx --target JA --source EN --notes --masters --glossary glossary.csv --model gpt-4o-mini --strategy neighbor
"""

import argparse
import sys
from typing import Optional, Type

from dotenv import load_dotenv

from pptx_translate.pipeline import TranslationResult, translate_presentation
from pptx_translate.translators import (
    BaseTranslator,
    ChatGPTTranslator,
    ResponsesGPTTranslator,
)


def build_arg_parser() -> argparse.ArgumentParser:
    ap = argparse.ArgumentParser(
        description="Translate PPTX slides with GPT + python-pptx."
    )
    ap.add_argument("input", help="Input .pptx path")
    ap.add_argument("output", help="Output .pptx path")
    ap.add_argument(
        "--model",
        default="gpt-4o-mini",
        help="OpenAI model (e.g., gpt-4o, gpt-4o-mini, gpt-5-mini)",
    )
    ap.add_argument("--source", default=None, help="Source language (e.g., EN)")
    ap.add_argument("--target", required=True, help="Target language (e.g., JA)")
    ap.add_argument("--notes", action="store_true", help="Include speaker notes")
    ap.add_argument(
        "--masters", action="store_true", help="Include slide masters/layouts"
    )
    ap.add_argument(
        "--glossary", default=None, help="CSV file with source,target terms"
    )
    ap.add_argument(
        "--strategy",
        default="neighbor",
        choices=["neighbor", "title-only", "deck"],
        help="Context strategy",
    )
    ap.add_argument(
        "--temperature",
        type=float,
        default=None,
        help="Sampling temperature (omit to use the model's default)",
    )
    ap.add_argument(
        "--dry_run", action="store_true", help="Preview changes without saving"
    )
    ap.add_argument("--log", default=None, help="File path to write translation log")
    ap.add_argument(
        "--slides",
        default=None,
        help="Comma-separated slide numbers or ranges to translate (e.g., 1,3-5)",
    )
    ap.add_argument(
        "--api",
        default="chat",
        choices=["chat", "responses"],
        help="OpenAI API backend to use (chat=completions, responses=new Responses API)",
    )
    return ap


def translator_class(api_mode: str) -> Type[BaseTranslator]:
    if api_mode == "responses":
        return ResponsesGPTTranslator
    return ChatGPTTranslator


def run_cli(argv: Optional[list[str]] = None) -> int:
    load_dotenv()
    parser = build_arg_parser()
    args = parser.parse_args(argv)

    translator_cls = translator_class(args.api)
    try:
        translator = translator_cls(
            model=args.model,
            source=args.source,
            target=args.target,
            temperature=args.temperature,
        )
    except Exception as exc:
        print(f"[ERROR] Failed to initialize translator: {exc}", file=sys.stderr)
        return 1

    try:
        result: TranslationResult = translate_presentation(
            translator=translator,
            input_path=args.input,
            output_path=args.output,
            include_notes=args.notes,
            include_masters=args.masters,
            glossary_path=args.glossary,
            strategy=args.strategy,
            dry_run=args.dry_run,
            log_path=args.log,
            slide_spec=args.slides,
        )
    except KeyboardInterrupt:
        print("\n[INFO] Translation interrupted by user.", file=sys.stderr)
        return 130
    except Exception as exc:
        print(f"[ERROR] Translation failed: {exc}", file=sys.stderr)
        return 1

    for warn in result.warnings:
        print(warn)

    if result.skipped_all and not args.dry_run:
        print("[INFO] No content translated (check --slides filter).")

    return 0


def main():
    sys.exit(run_cli())


if __name__ == "__main__":
    main()
