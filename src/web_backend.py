from __future__ import annotations

from datetime import datetime
import importlib
from pathlib import Path
import queue
import shutil
import tempfile
import threading
from typing import Any


class GuideAppApi:
    MAX_ITEM_ROWS = 21
    PATRIMONY_LENGTH = 5
    TEMP_PRINT_RETENTION_SECONDS = 20

    def __init__(self, base_dir: Path | None = None) -> None:
        self.base_dir = (base_dir or Path.cwd()).resolve()
        self.delivery_template = self.base_dir / "Guia solano Entrega.docx"
        self.receipt_template = self.base_dir / "Guia solano Recebimento2.docx"
        self.inventory_base = self.base_dir / "base moveis.xls"
        self.inventory_lookup: dict[str, str] = {}
        self._inventory_ready = False
        self._inventory_lock = threading.Lock()
        self._printers_cache: list[str] = []
        self._default_printer_cache = ""
        self._printer_lock = threading.Lock()
        self._docx_tools_module = None
        self._docx_tools_lock = threading.Lock()
        self._print_queue: queue.Queue[dict[str, Any]] = queue.Queue()
        self._print_worker = threading.Thread(target=self._process_print_queue, daemon=True)
        self._print_worker.start()
        threading.Thread(target=self._load_inventory_lookup_background, daemon=True).start()
        threading.Thread(target=self._refresh_printer_cache, daemon=True).start()
        threading.Thread(target=self._warmup_word_automation_background, daemon=True).start()

    def _docx_tools(self):
        with self._docx_tools_lock:
            if self._docx_tools_module is None:
                self._docx_tools_module = importlib.import_module("src.docx_tools")
            return self._docx_tools_module

    def get_initial_state(self) -> dict[str, Any]:
        with self._printer_lock:
            printers = list(self._printers_cache)
            default_printer = self._default_printer_cache
        with self._inventory_lock:
            inventory_count = len(self.inventory_lookup)
            inventory_ready = self._inventory_ready

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
        }

    def refresh_printers(self) -> dict[str, Any]:
        self._refresh_printer_cache()
        with self._printer_lock:
            printers = list(self._printers_cache)
            default_printer = self._default_printer_cache
        with self._inventory_lock:
            inventory_count = len(self.inventory_lookup)
            inventory_ready = self._inventory_ready
        return {
            "ok": True,
            "printers": printers,
            "defaultPrinter": default_printer,
            "inventoryBaseFound": self.inventory_base.is_file(),
            "inventoryCount": inventory_count,
            "inventoryReady": inventory_ready,
        }

    def lookup_item(self, patrimony: str) -> dict[str, Any]:
        key = self._docx_tools().normalize_patrimony(patrimony)
        return {
            "ok": True,
            "description": self.inventory_lookup.get(key, ""),
        }

    def print_guides(self, payload: dict[str, Any]) -> dict[str, Any]:
        try:
            normalized = self._normalize_payload(payload)
            self._print_queue.put(normalized)
            return {
                "ok": True,
                "message": "Impressão enviada para a fila.",
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

    def _normalize_payload(self, payload: dict[str, Any]) -> dict[str, Any]:
        mode = str(payload.get("mode", "ambos")).strip() or "ambos"
        if mode not in {"entrega", "recebimento", "ambos"}:
            raise ValueError("Modo de geração inválido.")

        printer_name = str(payload.get("printerName", "")).strip()
        if not printer_name:
            raise ValueError("Selecione uma impressora.")

        copies_raw = payload.get("copies", 1)
        try:
            copies = int(copies_raw)
        except (TypeError, ValueError):
            raise ValueError("Informe um número inteiro de cópias.")
        if copies < 1:
            raise ValueError("A quantidade de cópias deve ser pelo menos 1.")

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
                raise ValueError("Modelo de entrega não encontrado.")
            if not delivery_receiver_unit:
                raise ValueError("Informe a UA Receptora da guia de entrega.")
            if delivery_items and not delivery_room:
                raise ValueError("Informe a sala da guia de entrega.")
            if not delivery_items:
                raise ValueError("Preencha pelo menos um item na entrega.")

        if mode in {"recebimento", "ambos"}:
            if not self.receipt_template.is_file():
                raise ValueError("Modelo de recebimento não encontrado.")
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
                raise ValueError(f"Cada linha de {label} precisa ter descrição.")
            if not patrimony and not allow_modifiable_guides:
                raise ValueError(
                    f"Cada linha de {label} precisa ter patrimônio, ou ative a guia modificável."
                )

            normalized.append((patrimony, description))

        if len(normalized) > self.MAX_ITEM_ROWS:
            raise ValueError(f"A guia de {label} aceita no máximo {self.MAX_ITEM_ROWS} linhas.")

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
