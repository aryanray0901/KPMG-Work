"""Create or replace Deck Refresh's local OpenAI environment file."""

from __future__ import annotations

from getpass import getpass
from pathlib import Path


ROOT = Path(__file__).resolve().parent
ENV_PATH = ROOT / ".env"


def main() -> int:
    print("Deck Refresh AI setup")
    print("Paste a current OpenAI API key. Input is hidden while you type.")
    key = getpass("OpenAI API key: ").strip()
    if not key:
        print("No key entered. Nothing was changed.")
        return 1
    if not key.startswith("sk-"):
        print("That does not look like an OpenAI API key. Nothing was changed.")
        return 1

    ENV_PATH.write_text(
        "OPENAI_API_KEY=" + key + "\n"
        "OPENAI_MODEL=gpt-5\n"
        "OPENAI_REASONING_EFFORT=minimal\n",
        encoding="utf-8",
    )
    try:
        ENV_PATH.chmod(0o600)
    except OSError:
        pass
    print("API key saved locally in .env. It is not added to exported ZIP files.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
