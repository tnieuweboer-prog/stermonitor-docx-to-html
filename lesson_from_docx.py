import os
import json
import time
from openai import OpenAI, RateLimitError, APIError


def call_openai_json(prompt: str, model: str = "gpt-4o-mini", max_retries: int = 5) -> dict:
    """
    Roept OpenAI aan en houdt rekening met rate limits.
    - bij RateLimitError: wacht en probeer opnieuw
    - bij andere API-fouten: probeer een paar keer
    - geeft JSON terug
    """
    api_key = os.getenv("OPENAI_API_KEY")
    if not api_key:
        raise RuntimeError("OPENAI_API_KEY ontbreekt.")

    client = OpenAI(api_key=api_key)

    # exponentiële backoff: 1s, 2s, 4s, 8s, ...
    delay = 1.0

    for attempt in range(1, max_retries + 1):
        try:
            resp = client.chat.completions.create(
                model=model,
                messages=[{"role": "user", "content": prompt}],
                response_format={"type": "json_object"},
            )
            # gelukt → JSON parsen
            return json.loads(resp.choices[0].message.content)

        except RateLimitError:
            if attempt == max_retries:
                raise RuntimeError("OpenAI rate limit bleef terugkomen, probeer later opnieuw.")
            time.sleep(delay)
            delay *= 2  # langer wachten bij volgende poging

        except APIError as e:
            # andere tijdelijke API-fout → ook even wachten
            if attempt == max_retries:
                raise RuntimeError(f"OpenAI API-fout bleef terugkomen: {e}")
            time.sleep(delay)
            delay *= 2

    # zou hier niet moeten komen
    raise RuntimeError("Onbekende fout bij OpenAI.")

