from __future__ import annotations

import json
from pathlib import Path
from typing import Dict

STORE_PATH = Path(__file__).with_name("feedback_ows.json")


def load_feedbacks() -> Dict[str, dict]:
    if not STORE_PATH.exists():
        return {}
    try:
        data = json.loads(STORE_PATH.read_text(encoding="utf-8"))
        return data if isinstance(data, dict) else {}
    except Exception:
        return {}


def save_feedback(
    key: str,
    fecha: str,
    estado: str,
    feedback: str,
    *,
    task_id: str = "",
    assignment_id: str = "",
    site_id: str = "",
    fme_id: str = "",
    coordinador: str = "",
) -> None:
    data = load_feedbacks()
    data[str(key)] = {
        "task_id": task_id,
        "assignment_id": assignment_id,
        "site_id": site_id,
        "fme_id": fme_id,
        "coordinador": coordinador,
        "fecha": fecha,
        "estado": estado,
        "feedback": feedback,
    }
    STORE_PATH.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")


def delete_feedback(key: str) -> None:
    data = load_feedbacks()
    if str(key) in data:
        del data[str(key)]
        STORE_PATH.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
