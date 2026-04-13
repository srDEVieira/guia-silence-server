from __future__ import annotations

from datetime import datetime
import getpass
import hashlib
import importlib
import json
import os
from pathlib import Path
import platform
import queue
import shutil
import tempfile
import threading
import time
from typing import Any
from urllib import error, request


class GuideAppApi:
    MAX_ITEM_ROWS = 21
    PATRIMONY_LENGTH = 5
    TEMP_PRINT_RETENTION_SECONDS = 20
    LICENSE_RECHECK_SECONDS = 8
    LICENSE_BACKGROUND_POLL_SECONDS = 6
    LICENSE_TIMEOUT_SECONDS = 3
    INVENTORY_SYNC_SECONDS = 120
    INVENTORY_META_TIMEOUT_SECONDS = 5

    def __init__(self, base_dir: Path | None = None) -> None:
        self.base_dir = (base_dir or Path.cwd()).resolve()
        self.delivery_template = self.base_dir / "Guia solano Entrega.docx"
        self.receipt_template = self.base_dir / "Guia solano Recebimento2.docx"
        self.inventory_base = self.base_dir / "base moveis.xls"

        self.inventory_lookup: dict[str, str] = {}
        self._inventory_ready = False
        self._inventory_lock = threading.Lock()
        self._inventory_version_path = self.base_dir / ".inventory_version.json"
        self._inventory_local_version = self._load_inventory_local_version()

        self._printers_cache: list[str] = []
        self._default_printer_cache = ""
        self._printer_lock = threading.Lock()

        self._docx_tools_module = None
        self._docx_tools_lock = threading.Lock()

        self.license_server_url = self._load_license_server_url()
        self._license_lock = threading.Lock()
        self._license_blocked = False
        self._license_connected = False
        self._license_message = "Licenca remota nao configurada."
        self._license_checked_at = 0.0
        self._license_device = self._build_device_metadata()
        self._license_cache_path = self.base_dir / ".license_cache.json"
        self._load_license_cache()

        self._print_queue: queue.Queue[dict[str, Any]] = queue.Queue()
        self._print_worker = threading.Thread(target=self._process_print_queue, daemon=True)
        self._print_worker.start()

        threading.Thread(target=self._load_inventory_lookup_background, daemon=True).start()
        threading.Thread(target=self._inventory_sync_loop, daemon=True).start()
        threading.Thread(target=self._refresh_printer_cache, daemon=True).start()
        threading.Thread(target=self._warmup_word_automation_background, daemon=True).start()
        threading.Thread(target=self._refresh_license_status, kwargs={"force": True}, daemon=True).start()
        threading.Thread(target=self._license_poll_loop, daemon=True).start()

    def _docx_tools(self):
        with self._docx_tools_lock:
            if self._docx_tools_module is None:
                self._docx_tools_module = importlib.import_module("src.docx_tools")
            return self._docx_tools_module

    def get_initial_state(self) -> dict[str, Any]:
        self._refresh_license_status(force=False)

        with self._printer_lock:
            printers = list(self._printers_cache)
            default_printer = self._default_printer_cache
        with self._inventory_lock:
            inventory_count = len(self.inventory_lookup)
            inventory_ready = self._inventory_ready
        with self._license_lock:
            license_blocked = self._license_blocked
            license_connected = self._license_connected
            license_message = self._license_message

        return {
            "ok": True,
            "maxItemRows": self.MAX_ITEM_ROWS,
            "patrimonyLength": self.PATRIMONY_LENGTH,
            "templates": {
                "delivery": self.delivery_template.name,
                "receipt": self.receipt_template.name,
            },
            "inventoryBaseFound": self.inventory_base.is_file(),
            "inventoryCount": inventory_count,
            "inventoryReady": inventory_ready,
            "printers": printers,
            "defaultPrinter": default_printer,
            "licenseBlocked": license_blocked,
            "licenseConnected": license_connected,
            "licenseMessage": license_message,
            "deviceId": self._license_device["device_id"],
        }

    def refresh_printers(self) -> dict[str, Any]:
        self._refresh_license_status(force=False)
        self._refresh_printer_cache()

        with self._printer_lock:
            printers = list(self._printers_cache)
            default_printer = self._default_printer_cache
        with self._inventory_lock:
            inventory_count = len(self.inventory_lookup)
            inventory_ready = self._inventory_ready
        with self._license_lock:
            license_blocked = self._license_blocked
            license_connected = self._license_connected
            license_message = self._license_message

        return {
            "ok": True,
            "printers": printers,
            "defaultPrinter": default_printer,
            "inventoryBaseFound": self.inventory_base.is_file(),
            "inventoryCount": inventory_count,
            "inventoryReady": inventory_ready,
            "licenseBlocked": license_blocked,
            "licenseConnected": license_connected,
            "licenseMessage": license_message,
        }

    def get_license_status(self) -> dict[str, Any]:
        self._refresh_license_status(force=False)
        with self._license_lock:
            return {
                "ok": True,
                "licenseBlocked": self._license_blocked,
                "licenseConnected": self._license_connected,
                "licenseMessage": self._license_message,
            }

    def lookup_item(self, patrimony: str) -> dict[str, Any]:
        key = self._docx_tools().normalize_patrimony(patrimony)
        return {
            "ok": True,
            "description": self.inventory_lookup.get(key, ""),
        }

    def get_profiles(self) -> dict[str, Any]:
        if not self.license_server_url:
            return {"ok": True, "profiles": []}

        endpoint = f"{self.license_server_url}/profiles"
        req = request.Request(endpoint, method="GET")
        try:
            with request.urlopen(req, timeout=self.LICENSE_TIMEOUT_SECONDS) as resp:
                raw = resp.read().decode("utf-8")
                data = json.loads(raw or "{}")
            profiles = data.get("profiles", [])
            if not isinstance(profiles, list):
                profiles = []
            return {"ok": True, "profiles": profiles}
        except Exception:
            return {"ok": False, "profiles": []}

    def print_guides(self, payload: dict[str, Any]) -> dict[str, Any]:
        try:
            self._refresh_license_status(force=True)
            with self._license_lock:
                if self._license_blocked:
                    raise ValueError(
                        self._license_message
                        or "Uso bloqueado remotamente. Entre em contato com o administrador."
                    )

            normalized = self._normalize_payload(payload)
            self._print_queue.put(normalized)
            return {
                "ok": True,
                "message": "Impressao enviada para a fila.",
            }
        except Exception as exc:
            return {
                "ok": False,
                "error": str(exc),
            }

    def _process_print_queue(self) -> None:
        while True:
            payload = self._print_queue.get()
            try:
                self._print_guides_from_payload(payload)
            except Exception:
                pass
            finally:
                self._print_queue.task_done()

    def _load_inventory_lookup(self) -> dict[str, str]:
        try:
            return self._docx_tools().load_inventory_lookup(self.inventory_base)
        except Exception:
            return {}

    def _load_inventory_lookup_background(self) -> None:
        self._sync_inventory_from_remote(force=True)
        lookup = self._load_inventory_lookup()
        with self._inventory_lock:
            self.inventory_lookup = lookup
            self._inventory_ready = True

    def _refresh_printer_cache(self) -> None:
        tools = self._docx_tools()
        printers = tools.list_printers()
        default_printer = tools.get_default_printer_name()
        with self._printer_lock:
            self._printers_cache = printers
            self._default_printer_cache = default_printer

    def _warmup_word_automation_background(self) -> None:
        try:
            self._docx_tools().warmup_word_automation()
        except Exception:
            pass

    def _inventory_sync_loop(self) -> None:
        while True:
            try:
                self._sync_inventory_from_remote(force=False)
            except Exception:
                pass
            time.sleep(self.INVENTORY_SYNC_SECONDS)

    def _license_poll_loop(self) -> None:
        while True:
            try:
                self._refresh_license_status(force=False)
            except Exception:
                pass
            time.sleep(self.LICENSE_BACKGROUND_POLL_SECONDS)

    def _normalize_payload(self, payload: dict[str, Any]) -> dict[str, Any]:
        mode = str(payload.get("mode", "ambos")).strip() or "ambos"
        if mode not in {"entrega", "recebimento", "ambos"}:
            raise ValueError("Modo de geracao invalido.")

        printer_name = str(payload.get("printerName", "")).strip()
        if not printer_name:
            raise ValueError("Selecione uma impressora.")

        copies_raw = payload.get("copies", 1)
        try:
            copies = int(copies_raw)
        except (TypeError, ValueError):
            raise ValueError("Informe um numero inteiro de copias.")
        if copies < 1:
            raise ValueError("A quantidade de copias deve ser pelo menos 1.")

        allow_modifiable_guides = bool(payload.get("allowModifiableGuides"))

        delivery_receiver_unit = str(payload.get("deliveryReceiverUnit", "")).strip()
        receipt_sender_unit = str(payload.get("receiptSenderUnit", "")).strip()
        delivery_room = str(payload.get("deliveryRoom", "")).strip()
        receipt_room = str(payload.get("receiptRoom", "")).strip()

        delivery_items = self._normalize_items(
            payload.get("deliveryItems", []),
            "entrega",
            allow_modifiable_guides=allow_modifiable_guides,
        )
        receipt_items = self._normalize_items(
            payload.get("receiptItems", []),
            "recebimento",
            allow_modifiable_guides=allow_modifiable_guides,
        )

        if mode in {"entrega", "ambos"}:
            if not self.delivery_template.is_file():
                raise ValueError("Modelo de entrega nao encontrado.")
            if not delivery_receiver_unit:
                raise ValueError("Informe a UA Receptora da guia de entrega.")
            if delivery_items and not delivery_room:
                raise ValueError("Informe a sala da guia de entrega.")
            if not delivery_items:
                raise ValueError("Preencha pelo menos um item na entrega.")

        if mode in {"recebimento", "ambos"}:
            if not self.receipt_template.is_file():
                raise ValueError("Modelo de recebimento nao encontrado.")
            if not receipt_sender_unit:
                raise ValueError("Informe a UA Remetente da guia de recebimento.")
            if receipt_items and not receipt_room:
                raise ValueError("Informe a sala da guia de recebimento.")
            if not receipt_items:
                raise ValueError("Preencha pelo menos um item no recebimento.")

        return {
            "mode": mode,
            "printer_name": printer_name,
            "copies": copies,
            "allow_modifiable_guides": allow_modifiable_guides,
            "delivery_receiver_unit": delivery_receiver_unit,
            "receipt_sender_unit": receipt_sender_unit,
            "delivery_room": delivery_room,
            "receipt_room": receipt_room,
            "delivery_items": delivery_items,
            "receipt_items": receipt_items,
        }

    def _normalize_items(
        self,
        items: Any,
        label: str,
        *,
        allow_modifiable_guides: bool = False,
    ) -> list[tuple[str, str]]:
        normalized: list[tuple[str, str]] = []
        source_items = items if isinstance(items, list) else []

        for item in source_items:
            if not isinstance(item, dict):
                continue

            patrimony = str(item.get("patrimony", "")).strip()
            description = str(item.get("description", "")).strip()

            if not patrimony and not description:
                continue
            if not description:
                raise ValueError(f"Cada linha de {label} precisa ter descricao.")
            if not patrimony and not allow_modifiable_guides:
                raise ValueError(
                    f"Cada linha de {label} precisa ter patrimonio, ou ative a guia modificavel."
                )

            normalized.append((patrimony, description))

        if len(normalized) > self.MAX_ITEM_ROWS:
            raise ValueError(f"A guia de {label} aceita no maximo {self.MAX_ITEM_ROWS} linhas.")

        return normalized

    def _schedule_temp_cleanup(self, working_dir: Path) -> None:
        def _cleanup() -> None:
            try:
                shutil.rmtree(working_dir, ignore_errors=True)
            except Exception:
                pass

        timer = threading.Timer(self.TEMP_PRINT_RETENTION_SECONDS, _cleanup)
        timer.daemon = True
        timer.start()

    def _print_guides_from_payload(self, payload: dict[str, Any]) -> None:
        timestamp = datetime.now().strftime("%Y%m%d-%H%M%S")
        working_dir = Path(tempfile.mkdtemp(prefix="guias-impressao-"))
        generated_files: list[Path] = []

        if payload["mode"] in {"entrega", "ambos"}:
            delivery_output = working_dir / f"Guia-Entrega-{timestamp}.docx"
            self._docx_tools().generate_delivery_document(
                template_path=self.delivery_template,
                output_path=delivery_output,
                receiver_unit=payload["delivery_receiver_unit"],
                room_number=payload["delivery_room"],
                items=payload["delivery_items"],
            )
            generated_files.append(delivery_output)

        if payload["mode"] in {"recebimento", "ambos"}:
            receipt_output = working_dir / f"Guia-Recebimento-{timestamp}.docx"
            self._docx_tools().generate_receipt_document(
                template_path=self.receipt_template,
                output_path=receipt_output,
                sender_unit=payload["receipt_sender_unit"],
                room_number=payload["receipt_room"],
                items=payload["receipt_items"],
            )
            generated_files.append(receipt_output)

        try:
            self._docx_tools().print_docx_batch(
                generated_files,
                payload["printer_name"],
                copies=payload["copies"],
            )
        except Exception:
            shutil.rmtree(working_dir, ignore_errors=True)
            raise

        self._schedule_temp_cleanup(working_dir)

    def _load_license_server_url(self) -> str:
        env_url = os.getenv("LICENSE_SERVER_URL", "").strip()
        if env_url:
            return env_url.rstrip("/")

        config_path = self.base_dir / "license_config.json"
        if not config_path.is_file():
            return ""
        try:
            data = json.loads(config_path.read_text(encoding="utf-8"))
            return str(data.get("server_url", "")).strip().rstrip("/")
        except Exception:
            return ""

    def _build_device_metadata(self) -> dict[str, str]:
        machine_name = platform.node().strip() or os.getenv("COMPUTERNAME", "").strip() or "unknown-machine"
        user_name = getpass.getuser().strip() or os.getenv("USERNAME", "").strip() or "unknown-user"
        source = f"{machine_name}|{user_name}".encode("utf-8", errors="ignore")
        device_id = hashlib.sha256(source).hexdigest()[:32]
        return {
            "device_id": device_id,
            "machine_name": machine_name,
            "user_name": user_name,
        }

    def _load_inventory_local_version(self) -> str:
        if not self._inventory_version_path.is_file():
            return ""
        try:
            data = json.loads(self._inventory_version_path.read_text(encoding="utf-8"))
        except Exception:
            return ""
        return str(data.get("version", "")).strip()

    def _save_inventory_local_version(self, version: str) -> None:
        payload = {"version": version.strip()}
        try:
            self._inventory_version_path.write_text(
                json.dumps(payload, ensure_ascii=False, indent=2),
                encoding="utf-8",
            )
            self._inventory_local_version = version.strip()
        except Exception:
            pass

    def _sync_inventory_from_remote(self, *, force: bool) -> None:
        if not self.license_server_url:
            return

        meta_url = f"{self.license_server_url}/inventory/meta"
        req = request.Request(meta_url, method="GET")
        try:
            with request.urlopen(req, timeout=self.INVENTORY_META_TIMEOUT_SECONDS) as resp:
                raw = resp.read().decode("utf-8")
                data = json.loads(raw or "{}")
        except Exception:
            return

        remote_version = str(data.get("version", "")).strip()
        remote_url = str(data.get("url", "")).strip()
        if not remote_url or not remote_version:
            return
        if not force and remote_version == self._inventory_local_version:
            return

        temp_target = self.base_dir / "base_moveis_tmp_download.xls"
        try:
            with request.urlopen(remote_url, timeout=self.INVENTORY_META_TIMEOUT_SECONDS + 5) as resp:
                content = resp.read()
            if not content:
                return
            temp_target.write_bytes(content)
            # Replace atomically only after successful download.
            temp_target.replace(self.inventory_base)
        except Exception:
            try:
                if temp_target.exists():
                    temp_target.unlink()
            except Exception:
                pass
            return

        lookup = self._load_inventory_lookup()
        with self._inventory_lock:
            self.inventory_lookup = lookup
            self._inventory_ready = True

        if remote_version:
            self._save_inventory_local_version(remote_version)

    def _load_license_cache(self) -> None:
        if not self._license_cache_path.is_file():
            return
        try:
            data = json.loads(self._license_cache_path.read_text(encoding="utf-8"))
        except Exception:
            return
        with self._license_lock:
            self._license_blocked = bool(data.get("blocked", False))
            self._license_connected = bool(data.get("connected", False))
            self._license_message = str(data.get("message", self._license_message))
            self._license_checked_at = float(data.get("checked_at", 0.0))

    def _save_license_cache(self) -> None:
        with self._license_lock:
            payload = {
                "blocked": self._license_blocked,
                "connected": self._license_connected,
                "message": self._license_message,
                "checked_at": self._license_checked_at,
            }
        try:
            self._license_cache_path.write_text(
                json.dumps(payload, ensure_ascii=False, indent=2),
                encoding="utf-8",
            )
        except Exception:
            pass

    def _refresh_license_status(self, *, force: bool = False) -> None:
        now = time.time()
        with self._license_lock:
            if not force and (now - self._license_checked_at) < self.LICENSE_RECHECK_SECONDS:
                return

        if not self.license_server_url:
            with self._license_lock:
                self._license_connected = False
                self._license_blocked = False
                self._license_message = "Licenca remota nao configurada."
                self._license_checked_at = now
            self._save_license_cache()
            return

        endpoint = f"{self.license_server_url}/register"
        payload_bytes = json.dumps(self._license_device).encode("utf-8")
        req = request.Request(
            endpoint,
            data=payload_bytes,
            headers={"Content-Type": "application/json"},
            method="POST",
        )

        try:
            with request.urlopen(req, timeout=self.LICENSE_TIMEOUT_SECONDS) as resp:
                raw = resp.read().decode("utf-8")
                data = json.loads(raw or "{}")
            blocked = bool(data.get("blocked", False))
            block_reason = str(data.get("block_reason", "")).strip()
            with self._license_lock:
                self._license_connected = True
                self._license_blocked = blocked
                self._license_message = (
                    (f"Licenca bloqueada: {block_reason}" if block_reason else "Licenca bloqueada pelo administrador.")
                    if blocked
                    else "Licenca validada."
                )
                self._license_checked_at = now
            self._save_license_cache()
        except (error.URLError, TimeoutError, ValueError, json.JSONDecodeError):
            with self._license_lock:
                was_blocked = self._license_blocked
                self._license_connected = False
                self._license_blocked = was_blocked
                self._license_message = (
                    "Licenca bloqueada (cache local)."
                    if was_blocked
                    else "Servidor de licenca indisponivel."
                )
                self._license_checked_at = now
            self._save_license_cache()
