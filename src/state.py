"""Pure helpers for preventing stale experiment results in the UI."""
from __future__ import annotations
import hashlib
import json


def settings_fingerprint(settings: dict) -> str:
    encoded = json.dumps(settings, sort_keys=True, ensure_ascii=False, default=str).encode("utf-8")
    return hashlib.sha256(encoded).hexdigest()


def should_invalidate(previous_fingerprint: str | None, current_fingerprint: str) -> bool:
    return previous_fingerprint is not None and previous_fingerprint != current_fingerprint
