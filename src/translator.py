from __future__ import annotations

import json

import anthropic
from loguru import logger

from src.config import settings
from src.llm.schema.translation import TranslationBatchResponse, TranslationItem
from src.llm.structured import structured_call
from src.models import (
    DocumentExtraction,
    FormField,
    TranslatedDocument,
    TranslatedField,
    TranslatedPage,
)

_SYSTEM = """\
You are an expert multilingual translator specialising in medical and insurance documents.

Rules:
- Translate the provided field labels (keys) and values into the target language.
- Keep numbers, dates, alphanumeric codes, IDs, and medical abbreviations unchanged.
- Keep [REDACTED] and blank/empty values unchanged.
- Do NOT translate proper nouns: person names, clinic names, insurance company names.
- Use standard medical terminology in the target language.
"""

_PROMPT_TEMPLATE = """\
Translate every field below into {language}.

Input (JSON array):
{fields_json}

Fill the 'items' list with one entry per input field, in the same order:
  translated_key: the translated label
  translated_value: the translated value (keep original if it is a code, number, or proper noun)
"""

_BATCH_SIZE = 20


class Translator:
    def __init__(self) -> None:
        self.client = anthropic.Anthropic(api_key=settings.anthropic_api_key)
        self.model = settings.anthropic_model
        self.language = settings.target_language
        logger.info(f"Translator initialised (model={self.model}, lang={self.language})")

    # ── Public ────────────────────────────────────────────────────────────────

    def translate_document(self, extraction: DocumentExtraction) -> TranslatedDocument:
        logger.info(f"=== TRANSLATION START → {self.language} ===")
        translated_pages: list[TranslatedPage] = []

        for page in extraction.pages:
            logger.info(
                f"Translating page {page.page_number}/{extraction.total_pages} "
                f"({len(page.fields)} fields) …"
            )
            t_fields = self._translate_fields(page.fields)
            translated_pages.append(
                TranslatedPage(
                    page_number=page.page_number,
                    width_px=page.width_px,
                    height_px=page.height_px,
                    translated_fields=t_fields,
                )
            )
            logger.info(f"  → page {page.page_number} done ({len(t_fields)} fields)")

        total = sum(len(p.translated_fields) for p in translated_pages)
        logger.info(f"=== TRANSLATION DONE: {total} fields → {self.language} ===")

        return TranslatedDocument(
            original=extraction,
            pages=translated_pages,
            target_language=self.language,
        )

    # ── Private ───────────────────────────────────────────────────────────────

    def _translate_fields(self, fields: list[FormField]) -> list[TranslatedField]:
        if not fields:
            return []

        results: list[TranslatedField] = []
        for batch_start in range(0, len(fields), _BATCH_SIZE):
            batch = fields[batch_start : batch_start + _BATCH_SIZE]
            logger.debug(
                f"  batch {batch_start // _BATCH_SIZE + 1}: "
                f"fields {batch_start}–{batch_start + len(batch) - 1}"
            )
            results.extend(self._translate_batch(batch))
        return results

    def _translate_batch(self, fields: list[FormField]) -> list[TranslatedField]:
        payload = [
            {"key": f.key, "value": f.value, "field_type": f.field_type}
            for f in fields
        ]
        prompt = _PROMPT_TEMPLATE.format(
            language=self.language,
            fields_json=json.dumps(payload, ensure_ascii=False, indent=2),
        )

        try:
            result = structured_call(
                self.client,
                self.model,
                messages=[{"role": "user", "content": prompt}],
                schema=TranslationBatchResponse,
                system=_SYSTEM,
                max_tokens=4096,
            )
            translations = result.items
        except anthropic.APIError as exc:
            logger.error(f"API error during translation batch: {exc}")
            return self._fallback(fields)
        except Exception as exc:
            logger.error(f"Error during translation batch: {exc}")
            return self._fallback(fields)

        if len(translations) != len(fields):
            logger.warning(
                f"Translation count mismatch: expected {len(fields)}, "
                f"got {len(translations)} — padding with originals"
            )
            while len(translations) < len(fields):
                i = len(translations)
                translations.append(
                    TranslationItem(
                        translated_key=fields[i].key,
                        translated_value=fields[i].value,
                    )
                )

        _PROTECTED = {"", "[REDACTED]", "[BLANK]", "[redacted]", "[blank]"}
        output: list[TranslatedField] = []
        for field, t in zip(fields, translations):
            translated_value = t.translated_value or field.value
            if field.value.strip() in _PROTECTED:
                translated_value = field.value
            output.append(TranslatedField(
                original=field,
                translated_key=t.translated_key or field.key,
                translated_value=translated_value,
            ))
        return output

    @staticmethod
    def _fallback(fields: list[FormField]) -> list[TranslatedField]:
        logger.warning("Using original text as fallback for this batch")
        return [
            TranslatedField(original=f, translated_key=f.key, translated_value=f.value)
            for f in fields
        ]
