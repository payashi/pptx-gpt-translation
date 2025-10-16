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


class ChatGPTTranslator(BaseTranslator):
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
            "While faithfully retaining every keyword from the source (device names, terminology, etc.), craft the translation so it stays natural and slide-ready: concise, nominal in tone, and preferably ending in nouns even if it requires light rephrasing. "
            "Adjust sentence structure so each translation reads as natural slide text in the target language, keeping it concise, polished, and ending in noun phrasing rather than a literal English rendering. "
            "For example, rephrase \"Cooling requirements for the FTQC device can be met with commercially available cryoplants\" as the noun phrase \"Fulfillment of FTQC device cooling requirements via commercially available cryoplants\". "
            "Leave English personal names exactly as written in English; do not translate or transliterate them. "
            "If a line is a source attribution, copy it verbatim without translating. "
            "Do NOT add extra commentary. Return only the translations joined by the same delimiter."
        )

        if context:
            sys += f" Use this context for disambiguation: {context[:4000]}"

        request_args = {
            "model": self.model,
            "messages": [
                {"role": "system", "content": sys},
                {"role": "user", "content": joined},
            ],
        }
        if self.temperature is not None:
            request_args["temperature"] = self.temperature

        resp = self.client.chat.completions.create(**request_args)
        out = resp.choices[0].message.content or ""
        parts = out.split(delimiter)
        if len(parts) != len(texts):
            return out.splitlines()[: len(texts)] + [""] * (
                len(texts) - len(out.splitlines())
            )
        return parts
