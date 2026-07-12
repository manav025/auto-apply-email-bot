"""
Shared AI client -- wraps Groq's free-tier API (OpenAI-compatible).

Get a free API key at https://console.groq.com/keys (no credit card required).
Groq's free tier gives generous daily request/token limits on open models
like Llama 3.3, which is what this project uses by default.
"""

import os
from groq import Groq

GROQ_API_KEY = os.getenv("GROQ_API_KEY")
MODEL = os.getenv("GROQ_MODEL", "llama-3.3-70b-versatile")

_client = None


def get_client() -> Groq:
    global _client
    if not GROQ_API_KEY:
        raise RuntimeError("Missing GROQ_API_KEY in .env. Get a free key at "
                            "https://console.groq.com/keys")
    if _client is None:
        _client = Groq(api_key=GROQ_API_KEY)
    return _client


def call_ai(prompt: str, max_tokens: int = 1000, temperature: float = 0.6) -> str:
    """Send a single-turn prompt to the AI and return the text response."""
    client = get_client()
    resp = client.chat.completions.create(
        model=MODEL,
        max_tokens=max_tokens,
        temperature=temperature,
        messages=[{"role": "user", "content": prompt}],
    )
    return resp.choices[0].message.content.strip()
