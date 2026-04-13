from __future__ import annotations

import json
import os
import uuid
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

from fastapi import FastAPI, Header, HTTPException
from fastapi.responses import HTMLResponse

try:
    import psycopg
except Exception:  # pragma: no cover
    psycopg = None

app = FastAPI(title="Guia License Server", version="1.3.0")

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
            cur.execute(
                """
                CREATE TABLE IF NOT EXISTS profiles (
                    profile_id TEXT PRIMARY KEY,
                    display_name TEXT NOT NULL,
                    hero_bg_url TEXT NOT NULL DEFAULT '',
                    accent_color TEXT NOT NULL DEFAULT '',
                    active BOOLEAN NOT NULL DEFAULT TRUE,
                    sort_order INTEGER NOT NULL DEFAULT 0,
                    updated_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
                )
                """
            )
        conn.commit()

    ensure_default_profiles()


def load_db() -> dict[str, Any]:
    if not DB_PATH.exists():
        return {"devices": {}, "profiles": []}
    data = json.loads(DB_PATH.read_text(encoding="utf-8"))
    if "devices" not in data:
        data["devices"] = {}
    if "profiles" not in data:
        data["profiles"] = []
    return data


def save_db(data: dict[str, Any]) -> None:
    DB_PATH.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")


def require_admin(token: str | None) -> None:
    if not ADMIN_TOKEN:
        raise HTTPException(status_code=500, detail="ADMIN_TOKEN nao configurado no servidor.")
    if token != ADMIN_TOKEN:
        raise HTTPException(status_code=401, detail="Token admin invalido.")


def normalize_profile(payload: dict[str, Any], *, require_name: bool) -> dict[str, Any]:
    display_name = str(payload.get("display_name", "")).strip()
    if require_name and not display_name:
        raise HTTPException(status_code=400, detail="display_name obrigatorio.")
    hero_bg_url = str(payload.get("hero_bg_url", "")).strip()
    accent_color = str(payload.get("accent_color", "")).strip()
    active = bool(payload.get("active", True))
    sort_order = int(payload.get("sort_order", 0) or 0)

    return {
        "display_name": display_name,
        "hero_bg_url": hero_bg_url,
        "accent_color": accent_color,
        "active": active,
        "sort_order": sort_order,
    }


def ensure_default_profiles() -> None:
    defaults = [
        {
            "profile_id": "padrao",
            "display_name": "Padrão",
            "hero_bg_url": "",
            "accent_color": "",
            "active": True,
            "sort_order": 1,
        }
    ]

    if using_postgres():
        with get_pg_connection() as conn:
            with conn.cursor() as cur:
                for profile in defaults:
                    cur.execute(
                        """
                        INSERT INTO profiles (profile_id, display_name, hero_bg_url, accent_color, active, sort_order)
                        VALUES (%s, %s, %s, %s, %s, %s)
                        ON CONFLICT (profile_id) DO NOTHING
                        """,
                        (
                            profile["profile_id"],
                            profile["display_name"],
                            profile["hero_bg_url"],
                            profile["accent_color"],
                            profile["active"],
                            profile["sort_order"],
                        ),
                    )
            conn.commit()
        return

    db = load_db()
    profiles = db.get("profiles", [])
    if not profiles:
        db["profiles"] = defaults
        save_db(db)


def get_profiles_data(*, include_inactive: bool) -> list[dict[str, Any]]:
    if using_postgres():
        with get_pg_connection() as conn:
            with conn.cursor() as cur:
                if include_inactive:
                    cur.execute(
                        """
                        SELECT profile_id, display_name, hero_bg_url, accent_color, active, sort_order, updated_at
                        FROM profiles
                        ORDER BY sort_order ASC, display_name ASC
                        """
                    )
                else:
                    cur.execute(
                        """
                        SELECT profile_id, display_name, hero_bg_url, accent_color, active, sort_order, updated_at
                        FROM profiles
                        WHERE active = TRUE
                        ORDER BY sort_order ASC, display_name ASC
                        """
                    )
                rows = cur.fetchall()
        return [
            {
                "profile_id": r[0],
                "display_name": r[1],
                "hero_bg_url": r[2] or "",
                "accent_color": r[3] or "",
                "active": bool(r[4]),
                "sort_order": int(r[5] or 0),
                "updated_at": r[6].isoformat() if r[6] else None,
            }
            for r in rows
        ]

    db = load_db()
    profiles = db.get("profiles", [])
    if include_inactive:
        return profiles
    return [p for p in profiles if bool(p.get("active", True))]


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


@app.get("/profiles")
def public_profiles() -> dict[str, Any]:
    return {"ok": True, "profiles": get_profiles_data(include_inactive=False)}


@app.get("/admin/profiles")
def admin_profiles(x_admin_token: str | None = Header(default=None)) -> dict[str, Any]:
    require_admin(x_admin_token)
    return {"ok": True, "profiles": get_profiles_data(include_inactive=True)}


@app.post("/admin/profiles")
def admin_create_profile(payload: dict[str, Any], x_admin_token: str | None = Header(default=None)) -> dict[str, Any]:
    require_admin(x_admin_token)
    profile = normalize_profile(payload, require_name=True)
    profile_id = str(payload.get("profile_id", "")).strip() or str(uuid.uuid4())[:8]

    if using_postgres():
        with get_pg_connection() as conn:
            with conn.cursor() as cur:
                cur.execute(
                    """
                    INSERT INTO profiles (profile_id, display_name, hero_bg_url, accent_color, active, sort_order, updated_at)
                    VALUES (%s, %s, %s, %s, %s, %s, NOW())
                    """,
                    (
                        profile_id,
                        profile["display_name"],
                        profile["hero_bg_url"],
                        profile["accent_color"],
                        profile["active"],
                        profile["sort_order"],
                    ),
                )
            conn.commit()
        return {"ok": True, "profile_id": profile_id}

    db = load_db()
    profiles = db["profiles"]
    if any(str(p.get("profile_id", "")).strip() == profile_id for p in profiles):
        raise HTTPException(status_code=409, detail="profile_id ja existe.")
    profiles.append(
        {
            "profile_id": profile_id,
            **profile,
            "updated_at": now_iso(),
        }
    )
    save_db(db)
    return {"ok": True, "profile_id": profile_id}


@app.put("/admin/profiles/{profile_id}")
def admin_update_profile(profile_id: str, payload: dict[str, Any], x_admin_token: str | None = Header(default=None)) -> dict[str, Any]:
    require_admin(x_admin_token)
    updates = normalize_profile(payload, require_name=False)

    if using_postgres():
        with get_pg_connection() as conn:
            with conn.cursor() as cur:
                cur.execute(
                    """
                    UPDATE profiles
                    SET
                        display_name = COALESCE(NULLIF(%s, ''), display_name),
                        hero_bg_url = %s,
                        accent_color = %s,
                        active = %s,
                        sort_order = %s,
                        updated_at = NOW()
                    WHERE profile_id = %s
                    RETURNING profile_id
                    """,
                    (
                        updates["display_name"],
                        updates["hero_bg_url"],
                        updates["accent_color"],
                        updates["active"],
                        updates["sort_order"],
                        profile_id,
                    ),
                )
                row = cur.fetchone()
            conn.commit()
        if not row:
            raise HTTPException(status_code=404, detail="Perfil nao encontrado.")
        return {"ok": True, "profile_id": profile_id}

    db = load_db()
    profiles = db["profiles"]
    for p in profiles:
        if str(p.get("profile_id", "")).strip() == profile_id:
            if updates["display_name"]:
                p["display_name"] = updates["display_name"]
            p["hero_bg_url"] = updates["hero_bg_url"]
            p["accent_color"] = updates["accent_color"]
            p["active"] = updates["active"]
            p["sort_order"] = updates["sort_order"]
            p["updated_at"] = now_iso()
            save_db(db)
            return {"ok": True, "profile_id": profile_id}
    raise HTTPException(status_code=404, detail="Perfil nao encontrado.")


@app.delete("/admin/profiles/{profile_id}")
def admin_delete_profile(profile_id: str, x_admin_token: str | None = Header(default=None)) -> dict[str, Any]:
    require_admin(x_admin_token)

    if using_postgres():
        with get_pg_connection() as conn:
            with conn.cursor() as cur:
                cur.execute("DELETE FROM profiles WHERE profile_id = %s RETURNING profile_id", (profile_id,))
                row = cur.fetchone()
            conn.commit()
        if not row:
            raise HTTPException(status_code=404, detail="Perfil nao encontrado.")
        return {"ok": True, "profile_id": profile_id}

    db = load_db()
    profiles = db["profiles"]
    original_len = len(profiles)
    db["profiles"] = [p for p in profiles if str(p.get("profile_id", "")).strip() != profile_id]
    if len(db["profiles"]) == original_len:
        raise HTTPException(status_code=404, detail="Perfil nao encontrado.")
    save_db(db)
    return {"ok": True, "profile_id": profile_id}


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
      <p>Use endpoints admin com header <b>X-Admin-Token</b>.</p>
    </body></html>
    """


if using_postgres():
    init_postgres()
else:
    ensure_default_profiles()
