from __future__ import annotations

import json
import os
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

from fastapi import FastAPI, Header, HTTPException
from fastapi.responses import HTMLResponse

try:
    import psycopg
except Exception:  # pragma: no cover
    psycopg = None

app = FastAPI(title="Guia License Server", version="1.2.0")

DB_PATH = Path("devices.json")
ADMIN_TOKEN = os.getenv("ADMIN_TOKEN", "").strip()
DATABASE_URL = os.getenv("DATABASE_URL", "").strip()


def now_iso() -> str:
    return datetime.now(timezone.utc).isoformat()


def normalized_database_url() -> str:
    if not DATABASE_URL:
        return ""
    if DATABASE_URL.startswith("postgres://"):
        return "postgresql://" + DATABASE_URL[len("postgres://") :]
    return DATABASE_URL


def using_postgres() -> bool:
    return bool(normalized_database_url())


def get_pg_connection():
    db_url = normalized_database_url()
    if not db_url:
        raise RuntimeError("DATABASE_URL nao configurada.")
    if psycopg is None:
        raise RuntimeError("psycopg nao instalado.")
    return psycopg.connect(db_url)


def init_postgres() -> None:
    with get_pg_connection() as conn:
        with conn.cursor() as cur:
            cur.execute(
                """
                CREATE TABLE IF NOT EXISTS devices (
                    device_id TEXT PRIMARY KEY,
                    machine_name TEXT NOT NULL DEFAULT '',
                    user_name TEXT NOT NULL DEFAULT '',
                    blocked BOOLEAN NOT NULL DEFAULT FALSE,
                    first_seen TIMESTAMPTZ NOT NULL DEFAULT NOW(),
                    last_seen TIMESTAMPTZ NOT NULL DEFAULT NOW()
                )
                """
            )
        conn.commit()


def load_db() -> dict[str, Any]:
    if not DB_PATH.exists():
        return {"devices": {}}
    return json.loads(DB_PATH.read_text(encoding="utf-8"))


def save_db(data: dict[str, Any]) -> None:
    DB_PATH.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")


def require_admin(token: str | None) -> None:
    if not ADMIN_TOKEN:
        raise HTTPException(status_code=500, detail="ADMIN_TOKEN nao configurado no servidor.")
    if token != ADMIN_TOKEN:
        raise HTTPException(status_code=401, detail="Token admin invalido.")


@app.get("/")
def root() -> dict[str, Any]:
    return {"ok": True, "service": "license-server", "status": "online"}


@app.get("/health")
def health() -> dict[str, str]:
    storage = "postgres" if using_postgres() else "json"
    return {"status": "healthy", "storage": storage}


@app.post("/register")
def register(payload: dict[str, Any]) -> dict[str, Any]:
    device_id = str(payload.get("device_id", "")).strip()
    machine_name = str(payload.get("machine_name", "")).strip()
    user_name = str(payload.get("user_name", "")).strip()

    if not device_id:
        raise HTTPException(status_code=400, detail="device_id obrigatorio.")

    if using_postgres():
        with get_pg_connection() as conn:
            with conn.cursor() as cur:
                cur.execute(
                    """
                    INSERT INTO devices (device_id, machine_name, user_name)
                    VALUES (%s, %s, %s)
                    ON CONFLICT (device_id)
                    DO UPDATE SET
                        machine_name = EXCLUDED.machine_name,
                        user_name = EXCLUDED.user_name,
                        last_seen = NOW()
                    RETURNING blocked
                    """,
                    (device_id, machine_name, user_name),
                )
                row = cur.fetchone()
            conn.commit()

        blocked = bool(row[0]) if row else False
        return {"ok": True, "device_id": device_id, "blocked": blocked}

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

    if using_postgres():
        with get_pg_connection() as conn:
            with conn.cursor() as cur:
                cur.execute(
                    """
                    SELECT device_id, machine_name, user_name, blocked, first_seen, last_seen
                    FROM devices
                    ORDER BY last_seen DESC
                    """
                )
                rows = cur.fetchall()

        devices = [
            {
                "device_id": r[0],
                "machine_name": r[1],
                "user_name": r[2],
                "blocked": bool(r[3]),
                "first_seen": r[4].isoformat() if r[4] else None,
                "last_seen": r[5].isoformat() if r[5] else None,
            }
            for r in rows
        ]
        return {"ok": True, "devices": devices}

    db = load_db()
    return {"ok": True, "devices": list(db["devices"].values())}


@app.post("/admin/block/{device_id}")
def admin_block(device_id: str, x_admin_token: str | None = Header(default=None)) -> dict[str, Any]:
    require_admin(x_admin_token)

    if using_postgres():
        with get_pg_connection() as conn:
            with conn.cursor() as cur:
                cur.execute(
                    """
                    UPDATE devices
                    SET blocked = TRUE, last_seen = NOW()
                    WHERE device_id = %s
                    RETURNING device_id
                    """,
                    (device_id,),
                )
                row = cur.fetchone()
            conn.commit()

        if not row:
            raise HTTPException(status_code=404, detail="Dispositivo nao encontrado.")
        return {"ok": True, "device_id": device_id, "blocked": True}

    db = load_db()
    if device_id not in db["devices"]:
        raise HTTPException(status_code=404, detail="Dispositivo nao encontrado.")
    db["devices"][device_id]["blocked"] = True
    db["devices"][device_id]["last_seen"] = now_iso()
    save_db(db)
    return {"ok": True, "device_id": device_id, "blocked": True}


@app.post("/admin/unblock/{device_id}")
def admin_unblock(device_id: str, x_admin_token: str | None = Header(default=None)) -> dict[str, Any]:
    require_admin(x_admin_token)

    if using_postgres():
        with get_pg_connection() as conn:
            with conn.cursor() as cur:
                cur.execute(
                    """
                    UPDATE devices
                    SET blocked = FALSE, last_seen = NOW()
                    WHERE device_id = %s
                    RETURNING device_id
                    """,
                    (device_id,),
                )
                row = cur.fetchone()
            conn.commit()

        if not row:
            raise HTTPException(status_code=404, detail="Dispositivo nao encontrado.")
        return {"ok": True, "device_id": device_id, "blocked": False}

    db = load_db()
    if device_id not in db["devices"]:
        raise HTTPException(status_code=404, detail="Dispositivo nao encontrado.")
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


if using_postgres():
    init_postgres()
