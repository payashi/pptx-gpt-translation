import json
import os
from typing import List, Optional

from tenacity import (
    retry,
    retry_if_exception_type,
    stop_after_attempt,
    wait_exponential,
)

from .base import BaseTranslator

try:
    from openai import OpenAI
except Exception:
    OpenAI = None


class ResponsesGPTTranslator(BaseTranslator):
    def __init__(
        self,
        model: str,
        source: Optional[str],
        target: str,
        temperature: Optional[float] = None,
    ):
        super().__init__(model, source, target, temperature)
        if OpenAI is None:
            raise RuntimeError("OpenAI SDK not available. Install `openai` >= 1.37.0.")
        api_key = os.environ.get("OPENAI_API_KEY")
        if not api_key:
            raise RuntimeError("Set OPENAI_API_KEY environment variable.")
        self.client = OpenAI(api_key=api_key)

    @retry(
        reraise=True,
        stop=stop_after_attempt(4),
        wait=wait_exponential(multiplier=1, min=1, max=15),
        retry=retry_if_exception_type(Exception),
    )
    def translate(self, texts: List[str], context: str = "") -> List[str]:
        if not texts:
            return []

        delimiter = "\n\n<<<SPLIT>>>\n\n"
        joined = delimiter.join(texts)

        sys = (
            "You are a professional translator. "
            f"Translate from {self.source or 'the source language'} to {self.target}. "
            "Preserve technical terms, numbers, math, and code blocks. "
            "Keep bullet-like brevity for short lines; keep paragraph flow for long text. "
            "While faithfully retaining every keyword from the source (device names, terminology, etc.), craft the translation so it stays natural and slide-ready: concise and preferably ending in nouns. "
            "If a line is a source attribution, copy it verbatim without translating. "
            "Do NOT add extra commentary. Return only the translations joined by the same delimiter."
        )

        if context:
            sys += f" Use this context for disambiguation: {context[:4000]}"

        messages = [
            {
                "role": "system",
                "content": [{"type": "text", "text": sys}],
            },
            {
                "role": "user",
                "content": [{"type": "text", "text": joined}],
            },
        ]

        request_args = {
            "model": self.model,
            "input": messages,
            "response_format": {
                "type": "json_schema",
                "json_schema": {
                    "name": "translation_response",
                    "schema": {
                        "type": "object",
                        "properties": {
                            "translations": {
                                "type": "array",
                                "items": {"type": "string"},
                            }
                        },
                        "required": ["translations"],
                        "additionalProperties": False,
                    },
                },
            },
        }
        if self.temperature is not None:
            request_args["temperature"] = self.temperature

        resp = self.client.responses.create(**request_args)

        translations: Optional[List[str]] = None

        # Try structured output first
        for output in getattr(resp, "output", []) or []:
            for content in getattr(output, "content", []) or []:
                ctype = getattr(content, "type", None)
                if ctype == "json_schema":
                    data = getattr(content, "json", None)
                    if isinstance(data, dict) and "translations" in data:
                        translations = data["translations"]
                elif ctype == "text":
                    text_val = getattr(content, "text", None)
                    if text_val:
                        try:
                            data = json.loads(text_val)
                            if isinstance(data, dict) and "translations" in data:
                                translations = data["translations"]
                        except json.JSONDecodeError:
                            pass

        if translations is None:
            raw = getattr(resp, "output_text", None)
            if raw:
                try:
                    data = json.loads(raw)
                    if isinstance(data, dict) and "translations" in data:
                        translations = data["translations"]
                except json.JSONDecodeError:
                    pass

        if translations is None:
            # Fall back to delimiter-based parsing
            raw_text = getattr(resp, "output_text", "")
            parts = (raw_text or "").split(delimiter)
            if len(parts) != len(texts):
                return (raw_text or "").splitlines()[: len(texts)] + [""] * (
                    len(texts) - len((raw_text or "").splitlines())
                )
            return parts

        if len(translations) != len(texts):
            translations = translations[: len(texts)] + [""] * (
                len(texts) - len(translations)
            )

        return translations
