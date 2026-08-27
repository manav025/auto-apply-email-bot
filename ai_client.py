"""
Shared AI client -- wraps AgentRouter's OpenAI-compatible API for Claude Opus 4.8.
AgentRouter is a third-party gateway (not an official Anthropic product) that
proxies requests to Claude and other models behind one API key.

Set these in your .env:
  AGENTROUTER_API_KEY=sk-...                     (required)
  AGENTROUTER_BASE_URL=https://agentrouter.org/v1 (check your dashboard/docs --
                                                    some setups use co.agentrouter.org
                                                    or a different path)
  AGENTROUTER_MODEL=claude-opus-4-8               (or whatever model ID your
                                                    AgentRouter account lists)
"""
import os
from openai import OpenAI

AGENTROUTER_API_KEY = os.getenv("AGENTROUTER_API_KEY")
AGENTROUTER_BASE_URL = os.getenv("AGENTROUTER_BASE_URL", "https://agentrouter.org/v1")
MODEL = os.getenv("AGENTROUTER_MODEL", "claude-opus-4-8")

_client = None


def get_client() -> OpenAI:
    global _client
    if not AGENTROUTER_API_KEY:
        raise RuntimeError("Missing AGENTROUTER_API_KEY in .env. Get one from "
                            "your AgentRouter dashboard.")
    if _client is None:
        _client = OpenAI(api_key=AGENTROUTER_API_KEY, base_url=AGENTROUTER_BASE_URL)
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
