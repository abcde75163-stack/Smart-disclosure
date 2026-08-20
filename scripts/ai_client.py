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


def get_model() -> str:
    return os.environ.get("OPENAI_MODEL", DEFAULT_MODEL)


def get_fallback_model() -> str:
    return os.environ.get("OPENAI_FALLBACK_MODEL", "")


def strip_json_fence(text: str) -> str:
    text = text.strip()
    if text.startswith("```"):
        text = re.sub(r"^```(?:json)?\s*", "", text, flags=re.IGNORECASE)
        text = re.sub(r"\s*```$", "", text)
    match = re.search(r"\{[\s\S]*\}", text)
    return match.group(0) if match else text


def is_transient_error(error: Exception) -> bool:
    return isinstance(error, (APIConnectionError, APIError, APITimeoutError, RateLimitError))


def is_capacity_error(error: Exception) -> bool:
    message = str(error).lower()
    return "capacity" in message or "overloaded" in message or "temporarily unavailable" in message


def create_json_response(api_key: str, prompt: str, instructions: str, max_output_tokens: int = 4000) -> dict:
    client = OpenAI(api_key=api_key)
    model = get_model()
    fallback_model = get_fallback_model()

    for attempt in range(MAX_RETRIES):
        try:
            response = client.responses.create(
                model=model,
                instructions=instructions,
                input=[{"role": "user", "content": prompt}],
                text={"format": {"type": "json_object"}},
                max_output_tokens=max_output_tokens,
            )
            return json.loads(strip_json_fence(response.output_text))
        except BadRequestError as error:
            if fallback_model and is_capacity_error(error) and model != fallback_model:
                model = fallback_model
                continue
            raise
        except json.JSONDecodeError as error:
            raise ValueError(f"AI 응답을 JSON으로 파싱할 수 없습니다: {error}") from error
        except Exception as error:
            if fallback_model and is_capacity_error(error) and model != fallback_model:
                model = fallback_model
                continue
            if not is_transient_error(error) or attempt == MAX_RETRIES - 1:
                raise
            time.sleep(RETRY_BASE_DELAY * (2**attempt))

    raise RuntimeError("OpenAI API 호출에 실패했습니다.")
