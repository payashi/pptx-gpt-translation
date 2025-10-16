import inspect
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
        self._response_format_mode = self._detect_response_format_support()
        if self._response_format_mode is None:
            version = getattr(OpenAI, "__version__", None)
            if version is None:
                try:
                    import openai  # type: ignore

                    version = getattr(openai, "__version__", None)
                except Exception:
                    version = None
            raise RuntimeError(
                "The installed `openai` package does not expose structured output parameters for the Responses API. "
                "Install a release that supports either the `response_format` or `text.format` JSON schema configuration "
                "(for example `pip install \"openai>=1.48,<2.0\"`)."
                + (f" Detected version: {version}." if version else "")
            )

    def _detect_response_format_support(self) -> Optional[str]:
        try:
            signature = inspect.signature(self.client.responses.create)
        except (TypeError, ValueError, AttributeError):
            return None
        params = signature.parameters
        if "response_format" in params:
            return "response_format"
        if "text" in params:
            return "text"
        return None

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
            "Return the final answer strictly as JSON with a single field `translations`, which is an array of strings whose length matches the number of provided segments."
        )

        if context:
            sys += f" Use this context for disambiguation: {context[:4000]}"

        schema_body = {
            "type": "object",
            "properties": {
                "translations": {
                    "type": "array",
                    "items": {"type": "string"},
                }
            },
            "required": ["translations"],
            "additionalProperties": False,
        }

        request_args = {
            "model": self.model,
            "input": [
                {
                    "role": "system",
                    "content": [{"type": "input_text", "text": sys}],
                },
                {
                    "role": "user",
                    "content": [{"type": "input_text", "text": joined}],
                },
            ],
        }

        if self._response_format_mode == "response_format":
            request_args["response_format"] = {
                "type": "json_schema",
                "json_schema": {
                    "name": "translation_response",
                    "schema": schema_body,
                },
            }
        elif self._response_format_mode == "text":
            request_args["text"] = {
                "format": {
                    "type": "json_schema",
                    "name": "translation_response",
                    "schema": schema_body,
                }
            }
        else:
            raise RuntimeError("Structured output not supported by the current OpenAI client.")

        if self.temperature is not None:
            request_args["temperature"] = self.temperature

        resp = self.client.responses.create(**request_args)

        # Aggregate all text outputs; Responses API may return multiple pieces.
        collected_text: List[str] = []
        for output in getattr(resp, "output", []) or []:
            for content in getattr(output, "content", []) or []:
                ctype = getattr(content, "type", None)
                if ctype in ("json_schema", "output_json"):
                    data = getattr(content, "json", None)
                    if isinstance(data, dict):
                        translations = data.get("translations")
                        if isinstance(translations, list):
                            translations = [str(x) for x in translations]
                            if len(translations) != len(texts):
                                if len(translations) < len(texts):
                                    translations = translations + [""] * (
                                        len(texts) - len(translations)
                                    )
                                else:
                                    translations = translations[: len(texts)]
                            return translations
                if ctype == "output_text":
                    text_val = getattr(content, "text", None)
                    if text_val:
                        collected_text.append(text_val)

        if not collected_text:
            raw_text = getattr(resp, "output_text", "")
            if raw_text:
                collected_text.append(raw_text)

        translations: Optional[List[str]] = None
        for chunk in collected_text:
            try:
                data = json.loads(chunk)
            except json.JSONDecodeError:
                continue
            if isinstance(data, dict) and isinstance(data.get("translations"), list):
                translations = [str(x) for x in data["translations"]]
                break

        if translations is None:
            # Fall back to delimiter-based parsing
            fallback_text = "\n".join(collected_text)
            parts = fallback_text.split(delimiter)
            if len(parts) != len(texts):
                lines = fallback_text.splitlines()
                return lines[: len(texts)] + [""] * (len(texts) - len(lines))
            return parts

        if len(translations) != len(texts):
            if len(translations) < len(texts):
                translations = translations + [""] * (len(texts) - len(translations))
            else:
                translations = translations[: len(texts)]

        return translations
