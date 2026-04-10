from __future__ import annotations

from fastapi import FastAPI

app = FastAPI(title="Guia License Server", version="1.0.0")


@app.get("/")
def root() -> dict[str, object]:
    return {
        "ok": True,
        "service": "license-server",
        "status": "online",
    }


@app.get("/health")
def health() -> dict[str, str]:
    return {"status": "healthy"}

