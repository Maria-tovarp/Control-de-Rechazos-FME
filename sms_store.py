from __future__ import annotations

import json
from datetime import datetime
from pathlib import Path
from typing import Dict, Iterable

STORE_PATH = Path(__file__).with_name("sms_fme.json")


def load_sms_status() -> Dict[str, dict]:
    if not STORE_PATH.exists():
        return {}
    try:
        data = json.loads(STORE_PATH.read_text(encoding="utf-8"))
        return data if isinstance(data, dict) else {}
    except Exception:
        return {}


def mark_sms_sent(keys: Iterable[str]) -> None:
    data = load_sms_status()
    now = datetime.now().isoformat(timespec="seconds")
    for key in keys:
        if key:
            data[str(key)] = {"enviado": True, "fecha_envio": now}
    STORE_PATH.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")


def mark_sms_pending(keys: Iterable[str]) -> None:
    data = load_sms_status()
    changed = False
    for key in keys:
        if str(key) in data:
            del data[str(key)]
            changed = True
    if changed:
        STORE_PATH.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
