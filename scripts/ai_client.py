"""
OpenAI Responses API helper for JSON-only extraction tasks.
"""
import json
import os
import re
import time

from openai import APIConnectionError, APIError, APITimeoutError, BadRequestError, OpenAI, RateLimitError


DEFAULT_MODEL = "gpt-5"
MAX_RETRIES = 3
RETRY_BASE_DELAY = 1.5
MAX_JSON_TOKENS = 12000


def get_model() -> str:
    return os.environ.get("OPENAI_MODEL", DEFAULT_MODEL)


def get_fallback_model() -> str:
    return os.environ.get("OPENAI_FALLBACK_MODEL", "")


def strip_json_fence(text: str) -> str:
    text = (text or "").strip()
    if text.startswith("```"):
        text = re.sub(r"^```(?:json)?\s*", "", text, flags=re.IGNORECASE)
        text = re.sub(r"\s*```$", "", text)
    match = re.search(r"\{[\s\S]*\}", text)
    return match.group(0) if match else text


def repair_common_json_errors(text: str) -> str:
    """Repair small JSON formatting mistakes that sometimes appear in model output."""
    repaired = text.strip()
    repaired = re.sub(r",\s*([}\]])", r"\1", repaired)
    repaired = re.sub(r"([}\]])\s*\n\s*([{[])", r"\1,\n\2", repaired)
    repaired = re.sub(
        r'([}\]"0-9]|true|false|null)\s*\n\s*("([^"\\]|\\.)+"\s*:)',
        r"\1,\n\2",
        repaired,
    )
    return repaired


def parse_json_text(text: str) -> dict:
    cleaned = strip_json_fence(text)
    try:
        return json.loads(cleaned)
    except json.JSONDecodeError:
        repaired = repair_common_json_errors(cleaned)
        if repaired != cleaned:
            return json.loads(repaired)
        raise


def extract_response_text(response) -> str:
    """Return text from Responses API output, including fallbacks for empty output_text."""
    text = getattr(response, "output_text", "") or ""
    if text.strip():
        return text

    parts = []
    for item in getattr(response, "output", []) or []:
        for content in getattr(item, "content", []) or []:
            content_text = getattr(content, "text", None)
            if content_text:
                parts.append(content_text)
    return "\n".join(parts)


def summarize_response_state(response) -> str:
    status = getattr(response, "status", None)
    incomplete = getattr(response, "incomplete_details", None)
    error = getattr(response, "error", None)
    output_types = []
    for item in getattr(response, "output", []) or []:
        output_types.append(str(getattr(item, "type", type(item).__name__)))
    return (
        f"status={status}, incomplete_details={incomplete}, "
        f"error={error}, output_types={output_types}"
    )


def is_transient_error(error: Exception) -> bool:
    return isinstance(error, (APIConnectionError, APIError, APITimeoutError, RateLimitError))


def is_capacity_error(error: Exception) -> bool:
    message = str(error).lower()
    return "capacity" in message or "overloaded" in message or "temporarily unavailable" in message


def create_json_response(api_key: str, prompt: str, instructions: str, max_output_tokens: int = 4000) -> dict:
    client = OpenAI(api_key=api_key)
    model = get_model()
    fallback_model = get_fallback_model()
    output_tokens = min(max(max_output_tokens, 4000), MAX_JSON_TOKENS)
    original_prompt = prompt

    for attempt in range(MAX_RETRIES):
        try:
            response = client.responses.create(
                model=model,
                instructions=instructions,
                input=[{"role": "user", "content": prompt}],
                text={"format": {"type": "json_object"}},
                max_output_tokens=output_tokens,
            )
            text = extract_response_text(response)
            if not text.strip():
                state = summarize_response_state(response)
                if attempt < MAX_RETRIES - 1 and output_tokens < MAX_JSON_TOKENS:
                    output_tokens = min(output_tokens * 2, MAX_JSON_TOKENS)
                    time.sleep(RETRY_BASE_DELAY * (2**attempt))
                    continue
                raise ValueError(
                    "AI 응답이 비어 있습니다. 출력 토큰 제한 또는 응답 중단 가능성이 있습니다. "
                    f"{state}"
                )

            try:
                return parse_json_text(text)
            except json.JSONDecodeError as error:
                if attempt < MAX_RETRIES - 1:
                    prompt = (
                        "The previous answer was invalid JSON. Return the same result again as one valid JSON object only. "
                        "Do not omit commas between object properties or array items. Do not add markdown fences.\n\n"
                        "Original task:\n"
                        f"{original_prompt}\n\n"
                        "Invalid previous answer excerpt:\n"
                        f"{text[:6000]}"
                    )
                    time.sleep(RETRY_BASE_DELAY * (2**attempt))
                    continue
                preview = text[:1000]
                raise ValueError(
                    f"AI 응답을 JSON으로 파싱할 수 없습니다: {error}\n"
                    f"응답 앞부분: {preview!r}"
                ) from error
        except BadRequestError as error:
            if fallback_model and is_capacity_error(error) and model != fallback_model:
                model = fallback_model
                prompt = original_prompt
                continue
            raise
        except Exception as error:
            if fallback_model and is_capacity_error(error) and model != fallback_model:
                model = fallback_model
                prompt = original_prompt
                continue
            if not is_transient_error(error) or attempt == MAX_RETRIES - 1:
                raise
            time.sleep(RETRY_BASE_DELAY * (2**attempt))

    raise RuntimeError("OpenAI API 호출에 실패했습니다.")
