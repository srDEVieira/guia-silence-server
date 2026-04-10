from __future__ import annotations

import json
import os
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

from fastapi import FastAPI, Header, HTTPException
from fastapi.responses import HTMLResponse

app = FastAPI(title="Guia License Server", version="1.1.0")

DB_PATH = Path("devices.json")
ADMIN_TOKEN = os.getenv("ADMIN_TOKEN", "").strip()


def now_iso() -> str:
    return datetime.now(timezone.utc).isoformat()


def load_db() -> dict[str, Any]:
    if not DB_PATH.exists():
        return {"devices": {}}
    return json.loads(DB_PATH.read_text(encoding="utf-8"))


def save_db(data: dict[str, Any]) -> None:
    DB_PATH.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")


def require_admin(token: str | None) -> None:
    if not ADMIN_TOKEN:
        raise HTTPException(status_code=500, detail="ADMIN_TOKEN não configurado no servidor.")
    if token != ADMIN_TOKEN:
        raise HTTPException(status_code=401, detail="Token admin inválido.")


@app.get("/")
def root() -> dict[str, Any]:
    return {"ok": True, "service": "license-server", "status": "online"}


@app.get("/health")
def health() -> dict[str, str]:
    return {"status": "healthy"}


@app.post("/register")
def register(payload: dict[str, Any]) -> dict[str, Any]:
    device_id = str(payload.get("device_id", "")).strip()
    machine_name = str(payload.get("machine_name", "")).strip()
    user_name = str(payload.get("user_name", "")).strip()

    if not device_id:
        raise HTTPException(status_code=400, detail="device_id obrigatório.")

    db = load_db()
    devices = db["devices"]

    current = devices.get(device_id, {})
    blocked = bool(current.get("blocked", False))

    devices[device_id] = {
        "device_id": device_id,
        "machine_name": machine_name,
        "user_name": user_name,
        "blocked": blocked,
        "first_seen": current.get("first_seen", now_iso()),
        "last_seen": now_iso(),
    }
    save_db(db)

    return {"ok": True, "device_id": device_id, "blocked": blocked}


@app.get("/admin/devices")
def admin_devices(x_admin_token: str | None = Header(default=None)) -> dict[str, Any]:
    require_admin(x_admin_token)
    db = load_db()
    return {"ok": True, "devices": list(db["devices"].values())}


@app.post("/admin/block/{device_id}")
def admin_block(device_id: str, x_admin_token: str | None = Header(default=None)) -> dict[str, Any]:
    require_admin(x_admin_token)
    db = load_db()
    if device_id not in db["devices"]:
        raise HTTPException(status_code=404, detail="Dispositivo não encontrado.")
    db["devices"][device_id]["blocked"] = True
    db["devices"][device_id]["last_seen"] = now_iso()
    save_db(db)
    return {"ok": True, "device_id": device_id, "blocked": True}


@app.post("/admin/unblock/{device_id}")
def admin_unblock(device_id: str, x_admin_token: str | None = Header(default=None)) -> dict[str, Any]:
    require_admin(x_admin_token)
    db = load_db()
    if device_id not in db["devices"]:
        raise HTTPException(status_code=404, detail="Dispositivo não encontrado.")
    db["devices"][device_id]["blocked"] = False
    db["devices"][device_id]["last_seen"] = now_iso()
    save_db(db)
    return {"ok": True, "device_id": device_id, "blocked": False}


@app.get("/admin", response_class=HTMLResponse)
def admin_page() -> str:
    return """
    <html><body style="font-family:Segoe UI;padding:24px">
      <h2>Painel Admin</h2>
      <p>Use os endpoints /admin/devices, /admin/block/{id}, /admin/unblock/{id} com header <b>X-Admin-Token</b>.</p>
    </body></html>
    """
