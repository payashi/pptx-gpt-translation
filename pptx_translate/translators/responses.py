import inspect
import json
import logging
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

logger = logging.getLogger(__name__)


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
                '(for example `pip install "openai>=1.48,<2.0"`).'
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

        payload = json.dumps({"segments": texts}, ensure_ascii=False)

        sys = (
            "You are a professional translator. "
            f"Translate from {self.source or 'the source language'} to {self.target}. "
            "Preserve technical terms, numbers, math, and code blocks. "
            "Keep bullet-like brevity for short lines; keep paragraph flow for long text. "
            "Adjust sentence structure so each translation reads as natural slide text in the target language, keeping it concise, polished, and ending in noun phrasing rather than a literal English rendering. "
            'For example, rephrase "Cooling requirements for the FTQC device can be met with commercially available cryoplants" as the noun phrase "Fulfillment of FTQC device cooling requirements via commercially available cryoplants". '
            "Leave English personal names exactly as written in English; do not translate or transliterate them. "
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
                    "content": [
                        {
                            "type": "input_text",
                            "text": (
                                "Translate each entry in the `segments` array contained in the "
                                "following JSON payload. Return only the JSON response requested "
                                "in the system instructions."
                            ),
                        },
                        {"type": "input_text", "text": payload},
                    ],
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
            raise RuntimeError(
                "Structured output not supported by the current OpenAI client."
            )

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
                        if translations is None:
                            logger.warning(
                                "Responses API JSON payload missing `translations` field; retrying request. Payload: %s",
                                str(data)[:400],
                            )
                            raise ValueError("Missing `translations` in response JSON")
                        if not isinstance(translations, list):
                            logger.warning(
                                "Responses API returned `translations` of type %s; expected list. Retrying request. Payload: %s",
                                type(translations).__name__,
                                str(data)[:400],
                            )
                            raise ValueError(
                                "Invalid `translations` type in response JSON"
                            )
                        if len(translations) != len(texts):
                            logger.warning(
                                "Responses API returned %s translations; expected %s. Retrying request. Payload: %s",
                                len(translations),
                                len(texts),
                                str(data)[:400],
                            )
                            raise ValueError(
                                "Mismatched translations count in response JSON"
                            )
                        return [str(x) for x in translations]
                if ctype == "output_text":
                    text_val = getattr(content, "text", None)
                    if text_val:
                        collected_text.append(text_val)

        if not collected_text:
            raw_text = getattr(resp, "output_text", "")
            if raw_text:
                collected_text.append(raw_text)

        fallback_text = "\n".join(collected_text).strip()
        errors: List[str] = []
        for chunk in collected_text:
            try:
                data = json.loads(chunk)
            except json.JSONDecodeError as exc:
                errors.append(f"JSON decode error: {exc}")
                continue
            if not isinstance(data, dict):
                errors.append(f"Response JSON type {type(data).__name__}")
                continue
            translations = data.get("translations")
            if translations is None:
                errors.append("Missing `translations` field")
                continue
            if not isinstance(translations, list):
                logger.warning(
                    "Responses API returned `translations` of type %s; expected list. Retrying request.",
                    type(translations).__name__,
                )
                raise ValueError("Invalid `translations` type in response JSON")
            if len(translations) != len(texts):
                logger.warning(
                    "Responses API returned %s translations; expected %s. Retrying request. Raw output: %s",
                    len(translations),
                    len(texts),
                    chunk[:400],
                )
                raise ValueError("Mismatched translations count in response JSON")
            return [str(x) for x in translations]

        if fallback_text:
            logger.warning(
                "Responses API returned non-conforming payload; retrying request. Reasons: %s. Raw output: %s",
                "; ".join(errors) if errors else "unknown",
                fallback_text[:400],
            )
        else:
            logger.warning("Responses API returned empty payload; retrying request.")

        raise ValueError("Responses API did not return valid translations JSON")
