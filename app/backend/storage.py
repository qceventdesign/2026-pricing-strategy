"""JSON file-based storage for estimates."""

from __future__ import annotations
import json
import uuid
from datetime import datetime, timezone
from pathlib import Path

DATA_DIR = Path(__file__).resolve().parent.parent / "data"
ESTIMATES_FILE = DATA_DIR / "estimates.json"


def _ensure_data_dir():
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    if not ESTIMATES_FILE.exists():
        ESTIMATES_FILE.write_text("[]")


def _read_all() -> list[dict]:
    _ensure_data_dir()
    return json.loads(ESTIMATES_FILE.read_text())


def _write_all(estimates: list[dict]):
    _ensure_data_dir()
    ESTIMATES_FILE.write_text(json.dumps(estimates, indent=2, default=str))


def create_estimate(data: dict) -> dict:
    estimates = _read_all()
    estimate = {
        "id": str(uuid.uuid4()),
        "created_at": datetime.now(timezone.utc).isoformat(),
        "updated_at": datetime.now(timezone.utc).isoformat(),
        **data,
    }
    estimates.append(estimate)
    _write_all(estimates)
    return estimate


def get_all_estimates() -> list[dict]:
    return _read_all()


def get_estimate(estimate_id: str) -> dict | None:
    for est in _read_all():
        if est["id"] == estimate_id:
            return est
    return None


def update_estimate(estimate_id: str, updates: dict) -> dict | None:
    estimates = _read_all()
    for i, est in enumerate(estimates):
        if est["id"] == estimate_id:
            for k, v in updates.items():
                if v is not None:
                    est[k] = v
            est["updated_at"] = datetime.now(timezone.utc).isoformat()
            estimates[i] = est
            _write_all(estimates)
            return est
    return None


def delete_estimate(estimate_id: str) -> bool:
    estimates = _read_all()
    new_list = [e for e in estimates if e["id"] != estimate_id]
    if len(new_list) == len(estimates):
        return False
    _write_all(new_list)
    return True
