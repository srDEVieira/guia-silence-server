from __future__ import annotations

from datetime import datetime
from pathlib import Path
import tempfile
import threading
import tkinter as tk
from tkinter import messagebox, ttk

try:
    from src.docx_tools import (
        generate_delivery_document,
        generate_receipt_document,
        get_default_printer_name,
        list_printers,
        load_inventory_lookup,
        normalize_patrimony,
        print_docx_batch,
        warmup_word_automation,
    )
except ModuleNotFoundError:
    from docx_tools import (
        generate_delivery_document,
        generate_receipt_document,
        get_default_printer_name,
        list_printers,
        load_inventory_lookup,
        normalize_patrimony,
        print_docx_batch,
        warmup_word_automation,
    )


class DocumentGeneratorApp(tk.Tk):
    MAX_ITEM_ROWS = 21
    BG_COLOR = "#e9eef4"
    SURFACE_COLOR = "#f6f8fb"
    PANEL_COLOR = "#ffffff"
    ACCENT_COLOR = "#0f4c81"
    ACCENT_SOFT = "#e6f0f8"
    TEXT_COLOR = "#16324f"
    MUTED_TEXT = "#607284"
    BORDER_COLOR = "#d5dde6"
    PREVIEW_BG = "#d7e0e8"

    def __init__(self) -> None:
        super().__init__()
        self.title("Guia patrimonial")
        self.geometry("1380x860")
        self.minsize(1180, 760)
        self.configure(bg=self.BG_COLOR, padx=18, pady=16)

        base_dir = Path.cwd()

        self.delivery_template = tk.StringVar(value=str(base_dir / "Guia solano Entrega.docx"))
        self.receipt_template = tk.StringVar(value=str(base_dir / "Guia solano Recebimento2.docx"))
        self.inventory_base_path = base_dir / "base moveis.xls"
        self.document_type = tk.StringVar(value="ambos")
        self.delivery_receiver_unit = tk.StringVar()
        self.receipt_sender_unit = tk.StringVar()
        self.delivery_room = tk.StringVar()
        self.receipt_room = tk.StringVar()
        self.selected_printer = tk.StringVar()
        self.print_copies = tk.StringVar(value="1")
        self.status_text = tk.StringVar(
            value="Selecione a guia, preencha os campos e confira a pagina na pre-visualizacao."
        )

        self.delivery_item_rows: list[tuple[tk.StringVar, tk.StringVar, ttk.Frame]] = []
        self.receipt_item_rows: list[tuple[tk.StringVar, tk.StringVar, ttk.Frame]] = []
        self.delivery_file_label: ttk.Label | None = None
        self.receipt_file_label: ttk.Label | None = None
        self.delivery_fields_frame: ttk.LabelFrame | None = None
        self.receipt_fields_frame: ttk.LabelFrame | None = None
        self.delivery_items_frame: ttk.LabelFrame | None = None
        self.receipt_items_frame: ttk.LabelFrame | None = None
        self.delivery_items_container: ttk.Frame | None = None
        self.receipt_items_container: ttk.Frame | None = None
        self.delivery_items_canvas: tk.Canvas | None = None
        self.receipt_items_canvas: tk.Canvas | None = None
        self.delivery_items_window_id: int | None = None
        self.receipt_items_window_id: int | None = None
        self.preview_canvas: tk.Canvas | None = None
        self.controls_canvas: tk.Canvas | None = None
        self.controls_window_id: int | None = None
        self.printer_combo: ttk.Combobox | None = None
        self.printer_status_label: ttk.Label | None = None
        self.inventory_lookup: dict[str, str] = {}
        self.header_glow_canvas: tk.Canvas | None = None
        self._header_glow_step = 0

        self._configure_theme()
        self._bind_live_preview()
        self._build_ui()
        threading.Thread(target=warmup_word_automation, daemon=True).start()
        self._load_inventory_lookup()
        self._load_printers()
        self._bind_mousewheel_support()
        for _ in range(8):
            self._add_item_row("delivery")
            self._add_item_row("receipt")
        self._update_mode_fields()
        self._refresh_preview()
        self._animate_header_glow()

    def _configure_theme(self) -> None:
        style = ttk.Style(self)
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass

        style.configure(".", background=self.BG_COLOR, foreground=self.TEXT_COLOR, font=("Segoe UI", 10))
        style.configure("App.TFrame", background=self.BG_COLOR)
        style.configure("Card.TFrame", background=self.PANEL_COLOR)
        style.configure(
            "Card.TLabelframe",
            background=self.PANEL_COLOR,
            borderwidth=1,
            relief="solid",
            bordercolor=self.BORDER_COLOR,
            lightcolor=self.BORDER_COLOR,
            darkcolor=self.BORDER_COLOR,
        )
        style.configure(
            "Card.TLabelframe.Label",
            background=self.PANEL_COLOR,
            foreground=self.TEXT_COLOR,
            font=("Segoe UI Semibold", 10),
        )
        style.configure("SectionTitle.TLabel", background=self.PANEL_COLOR, foreground=self.TEXT_COLOR, font=("Segoe UI Semibold", 10))
        style.configure("Body.TLabel", background=self.PANEL_COLOR, foreground=self.TEXT_COLOR, font=("Segoe UI", 10))
        style.configure("Muted.TLabel", background=self.PANEL_COLOR, foreground=self.MUTED_TEXT, font=("Segoe UI", 9))
        style.configure("Status.TLabel", background=self.TEXT_COLOR, foreground="#f8fafc", font=("Segoe UI Semibold", 9))
        style.configure("Accent.TButton", font=("Segoe UI Semibold", 10), padding=(14, 10), foreground="#ffffff", background=self.ACCENT_COLOR, borderwidth=0)
        style.map("Accent.TButton", background=[("active", "#135e9e"), ("pressed", "#0b3d68")])
        style.configure("Soft.TButton", font=("Segoe UI Semibold", 10), padding=(12, 9), foreground=self.ACCENT_COLOR, background=self.ACCENT_SOFT, borderwidth=0)
        style.map("Soft.TButton", background=[("active", "#d9e8f5"), ("pressed", "#c7ddee")])
        style.configure("Field.TEntry", fieldbackground="#f8fbff", background="#f8fbff", foreground=self.TEXT_COLOR, bordercolor="#c7d2e0", lightcolor="#c7d2e0", darkcolor="#c7d2e0", padding=6)
        style.configure("Field.TCombobox", fieldbackground="#f8fbff", background="#f8fbff", foreground=self.TEXT_COLOR, bordercolor="#c7d2e0", lightcolor="#c7d2e0", darkcolor="#c7d2e0", padding=6, arrowsize=14)
        style.map("Field.TCombobox", fieldbackground=[("readonly", "#f8fbff")])
        style.configure("Mode.TRadiobutton", background=self.PANEL_COLOR, foreground=self.TEXT_COLOR, font=("Segoe UI", 10))
        style.map("Mode.TRadiobutton", foreground=[("active", self.ACCENT_COLOR)])

    def _build_ui(self) -> None:
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)

        header = tk.Frame(self, bg=self.TEXT_COLOR, highlightthickness=0, bd=0)
        header.grid(row=0, column=0, sticky="ew", pady=(0, 14))
        header.columnconfigure(0, weight=1)

        content = tk.Frame(header, bg=self.TEXT_COLOR, padx=24, pady=22)
        content.grid(row=0, column=0, sticky="ew")
        content.columnconfigure(0, weight=1)
        tk.Label(
            content,
            text="Central de Guias Patrimoniais",
            bg=self.TEXT_COLOR,
            fg="#ffffff",
            font=("Segoe UI Semibold", 22),
        ).grid(row=0, column=0, sticky="w")
        tk.Label(
            content,
            text="Fluxo interno para preenchimento, conferencia visual e impressao direta.",
            bg=self.TEXT_COLOR,
            fg="#d6e3ef",
            font=("Segoe UI", 10),
        ).grid(row=1, column=0, sticky="w", pady=(4, 0))
        badges = tk.Frame(content, bg=self.TEXT_COLOR)
        badges.grid(row=0, column=1, rowspan=2, sticky="e", padx=(12, 0))
        for idx, text in enumerate(["Preview", "Base automatica", "Impressao"]):
            tk.Label(
                badges,
                text=text,
                bg="#27496d",
                fg="#f8fafc",
                padx=12,
                pady=5,
                font=("Segoe UI", 9),
            ).grid(row=0, column=idx, padx=(0 if idx == 0 else 8, 0))

        self.header_glow_canvas = tk.Canvas(header, height=18, bg=self.TEXT_COLOR, highlightthickness=0, bd=0)
        self.header_glow_canvas.grid(row=1, column=0, sticky="ew")

        body = ttk.Frame(self, style="App.TFrame")
        body.grid(row=1, column=0, sticky="nsew")
        body.columnconfigure(0, weight=0)
        body.columnconfigure(1, weight=1)
        body.rowconfigure(0, weight=1)

        self._build_controls_panel(body)
        self._build_preview_panel(body)

        footer = tk.Frame(self, bg=self.TEXT_COLOR, padx=16, pady=12)
        footer.grid(row=2, column=0, sticky="ew", pady=(10, 0))
        footer.columnconfigure(0, weight=1)

        ttk.Label(footer, textvariable=self.status_text, style="Status.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Button(footer, text="Limpar itens", command=self._clear_items, style="Soft.TButton").grid(row=0, column=1, padx=(12, 0))
        ttk.Button(footer, text="Imprimir", command=self._print_documents, style="Accent.TButton").grid(
            row=0, column=2, padx=(10, 0)
        )

    def _build_controls_panel(self, parent: ttk.Frame) -> None:
        wrapper = ttk.Frame(parent, width=470, style="Card.TFrame")
        wrapper.grid(row=0, column=0, sticky="nsew", padx=(0, 14))
        wrapper.grid_propagate(False)
        wrapper.columnconfigure(0, weight=1)
        wrapper.rowconfigure(0, weight=1)

        canvas = tk.Canvas(wrapper, highlightthickness=0, bg=self.PANEL_COLOR)
        scrollbar = ttk.Scrollbar(wrapper, orient="vertical", command=canvas.yview)
        panel = ttk.Frame(canvas, style="Card.TFrame", padding=14)
        panel.columnconfigure(0, weight=1)

        panel.bind("<Configure>", lambda _e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.bind("<Configure>", self._resize_controls_panel)

        self.controls_window_id = canvas.create_window((0, 0), window=panel, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.grid(row=0, column=0, sticky="nsew")
        scrollbar.grid(row=0, column=1, sticky="ns")

        self.controls_canvas = canvas

        self._build_template_section(panel).grid(row=0, column=0, sticky="ew", pady=(0, 10))
        self._build_printer_section(panel).grid(row=1, column=0, sticky="ew", pady=(0, 10))
        self._build_type_section(panel).grid(row=2, column=0, sticky="ew", pady=(0, 10))
        self._build_people_section(panel).grid(row=3, column=0, sticky="ew", pady=(0, 10))
        self._build_items_section(panel).grid(row=4, column=0, sticky="nsew")

    def _build_preview_panel(self, parent: ttk.Frame) -> None:
        wrapper = ttk.LabelFrame(parent, text="Pagina", style="Card.TLabelframe", padding=10)
        wrapper.grid(row=0, column=1, sticky="nsew")
        wrapper.columnconfigure(0, weight=1)
        wrapper.rowconfigure(0, weight=1)

        canvas = tk.Canvas(wrapper, bg=self.PREVIEW_BG, highlightthickness=0)
        scrollbar = ttk.Scrollbar(wrapper, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.grid(row=0, column=0, sticky="nsew")
        scrollbar.grid(row=0, column=1, sticky="ns")

        self.preview_canvas = canvas
        self.preview_canvas.bind("<Configure>", lambda _event: self._refresh_preview())

    def _build_template_section(self, parent: ttk.Frame) -> ttk.LabelFrame:
        frame = ttk.LabelFrame(parent, text="Modelos", style="Card.TLabelframe", padding=12)
        frame.columnconfigure(0, weight=1)

        ttk.Label(frame, text="Os modelos sao carregados automaticamente da pasta do projeto.", style="Muted.TLabel").grid(
            row=0, column=0, sticky="w", pady=(0, 6)
        )
        ttk.Label(frame, text=f"Entrega: {self._display_name(self.delivery_template.get())}", style="Body.TLabel").grid(
            row=1, column=0, sticky="w", pady=2
        )
        ttk.Label(frame, text=f"Recebimento: {self._display_name(self.receipt_template.get())}", style="Body.TLabel").grid(
            row=2, column=0, sticky="w", pady=2
        )
        base_status = "Base de moveis: carregada" if self.inventory_base_path.is_file() else "Base de moveis: nao encontrada"
        ttk.Label(frame, text=base_status, style="Muted.TLabel").grid(
            row=3, column=0, sticky="w", pady=(4, 0)
        )
        return frame

    def _build_printer_section(self, parent: ttk.Frame) -> ttk.LabelFrame:
        frame = ttk.LabelFrame(parent, text="Impressao", style="Card.TLabelframe", padding=12)
        frame.columnconfigure(0, weight=1)
        frame.columnconfigure(1, weight=0)

        self.printer_combo = ttk.Combobox(frame, textvariable=self.selected_printer, style="Field.TCombobox")
        self.printer_combo.grid(row=0, column=0, sticky="ew", padx=(0, 8), pady=(0, 8))
        ttk.Label(frame, text="Copias", style="SectionTitle.TLabel").grid(row=0, column=1, sticky="w", padx=(0, 8), pady=(0, 8))
        ttk.Entry(frame, textvariable=self.print_copies, width=6, style="Field.TEntry").grid(
            row=0, column=2, sticky="w", pady=(0, 8)
        )
        self.printer_status_label = ttk.Label(frame, text="", style="Muted.TLabel")
        self.printer_status_label.grid(row=1, column=0, columnspan=3, sticky="w", pady=(0, 8))
        ttk.Button(frame, text="Atualizar impressoras", command=self._load_printers, style="Soft.TButton").grid(
            row=2, column=0, columnspan=3, sticky="w"
        )
        return frame

    def _build_type_section(self, parent: ttk.Frame) -> ttk.LabelFrame:
        frame = ttk.LabelFrame(parent, text="Fluxo", style="Card.TLabelframe", padding=12)

        ttk.Radiobutton(
            frame,
            text="Somente entrega",
            value="entrega",
            variable=self.document_type,
            command=self._update_mode_fields,
            style="Mode.TRadiobutton",
        ).pack(anchor="w", pady=(0, 6))
        ttk.Radiobutton(
            frame,
            text="Somente recebimento",
            value="recebimento",
            variable=self.document_type,
            command=self._update_mode_fields,
            style="Mode.TRadiobutton",
        ).pack(anchor="w", pady=6)
        ttk.Radiobutton(
            frame,
            text="Gerar os dois",
            value="ambos",
            variable=self.document_type,
            command=self._update_mode_fields,
            style="Mode.TRadiobutton",
        ).pack(anchor="w", pady=(6, 0))
        return frame

    def _build_people_section(self, parent: ttk.Frame) -> ttk.Frame:
        frame = ttk.Frame(parent, style="Card.TFrame")
        frame.columnconfigure(0, weight=1)

        self.delivery_fields_frame = ttk.LabelFrame(frame, text="Entrega", style="Card.TLabelframe", padding=12)
        self.delivery_fields_frame.columnconfigure(1, weight=1)
        ttk.Label(self.delivery_fields_frame, text="UA Receptora", style="SectionTitle.TLabel").grid(
            row=0, column=0, sticky="w", pady=(0, 6)
        )
        ttk.Entry(self.delivery_fields_frame, textvariable=self.delivery_receiver_unit, style="Field.TEntry").grid(
            row=0, column=1, sticky="ew", pady=(0, 6)
        )
        ttk.Label(self.delivery_fields_frame, text="Sala", style="SectionTitle.TLabel").grid(
            row=1, column=0, sticky="w", pady=(0, 0)
        )
        ttk.Entry(self.delivery_fields_frame, textvariable=self.delivery_room, style="Field.TEntry").grid(
            row=1, column=1, sticky="ew"
        )

        self.receipt_fields_frame = ttk.LabelFrame(frame, text="Recebimento", style="Card.TLabelframe", padding=12)
        self.receipt_fields_frame.columnconfigure(1, weight=1)
        ttk.Label(self.receipt_fields_frame, text="UA Remetente", style="SectionTitle.TLabel").grid(
            row=0, column=0, sticky="w", pady=(0, 6)
        )
        ttk.Entry(self.receipt_fields_frame, textvariable=self.receipt_sender_unit, style="Field.TEntry").grid(
            row=0, column=1, sticky="ew", pady=(0, 6)
        )
        ttk.Label(self.receipt_fields_frame, text="Sala", style="SectionTitle.TLabel").grid(
            row=1, column=0, sticky="w"
        )
        ttk.Entry(self.receipt_fields_frame, textvariable=self.receipt_room, style="Field.TEntry").grid(
            row=1, column=1, sticky="ew"
        )
        return frame

    def _build_items_section(self, parent: ttk.Frame) -> ttk.Frame:
        frame = ttk.Frame(parent, style="Card.TFrame")
        frame.columnconfigure(0, weight=1)

        self.delivery_items_frame = self._build_single_items_frame(frame, "Itens da entrega", "delivery")
        self.receipt_items_frame = self._build_single_items_frame(frame, "Itens do recebimento", "receipt")
        return frame

    def _build_single_items_frame(self, parent: ttk.Frame, title: str, key: str) -> ttk.LabelFrame:
        frame = ttk.LabelFrame(parent, text=title, style="Card.TLabelframe", padding=12)
        frame.columnconfigure(0, weight=1)

        toolbar = ttk.Frame(frame, style="Card.TFrame")
        toolbar.grid(row=0, column=0, sticky="ew", pady=(0, 10))
        ttk.Button(toolbar, text="Adicionar linha", command=lambda: self._add_item_row(key), style="Soft.TButton").pack(side="left")
        ttk.Button(toolbar, text="Remover ultima", command=lambda: self._remove_item_row(key), style="Soft.TButton").pack(side="left", padx=(8, 0))

        header = ttk.Frame(frame, style="Card.TFrame")
        header.grid(row=1, column=0, sticky="ew", pady=(0, 8))
        header.columnconfigure(0, weight=2)
        header.columnconfigure(1, weight=3)
        ttk.Label(header, text="Numero de patrimonio", style="SectionTitle.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Label(header, text="Descricao do item", style="SectionTitle.TLabel").grid(row=0, column=1, sticky="w", padx=(12, 0))

        items_frame = ttk.Frame(frame, style="Card.TFrame")
        items_frame.grid(row=2, column=0, sticky="ew")
        items_frame.columnconfigure(0, weight=1)

        if key == "delivery":
            self.delivery_items_container = items_frame
            self.delivery_items_canvas = None
            self.delivery_items_window_id = None
        else:
            self.receipt_items_container = items_frame
            self.receipt_items_canvas = None
            self.receipt_items_window_id = None

        return frame

    def _display_name(self, path_text: str) -> str:
        path = Path(path_text)
        return path.name if path.name else "Nenhum arquivo selecionado"

    def _load_printers(self) -> None:
        printers = list_printers()
        default_printer = get_default_printer_name()

        if self.printer_combo is not None:
            self.printer_combo["values"] = printers

        if self.printer_status_label is not None:
            if printers:
                self.printer_status_label.config(text=f"Encontradas: {', '.join(printers)}")
            else:
                self.printer_status_label.config(
                    text="Nenhuma impressora foi listada automaticamente. Voce pode digitar o nome manualmente."
                )

        if default_printer and default_printer in printers:
            self.selected_printer.set(default_printer)
        elif printers and not self.selected_printer.get():
            self.selected_printer.set(printers[0])
        elif not printers:
            self.selected_printer.set("")
            self.status_text.set("Nenhuma impressora encontrada pelo Windows para este app.")

    def _load_inventory_lookup(self) -> None:
        if not self.inventory_base_path.is_file():
            return

        try:
            self.inventory_lookup = load_inventory_lookup(self.inventory_base_path)
            if self.inventory_lookup:
                self.status_text.set(
                    f"Base de moveis carregada com {len(self.inventory_lookup)} chapas para autopreenchimento."
                )
        except Exception:
            self.inventory_lookup = {}

    def _resize_items_container(self, event, key: str) -> None:
        if key == "delivery" and self.delivery_items_canvas is not None and self.delivery_items_window_id is not None:
            self.delivery_items_canvas.itemconfigure(self.delivery_items_window_id, width=event.width)
            self.delivery_items_canvas.configure(scrollregion=self.delivery_items_canvas.bbox("all"))
        if key == "receipt" and self.receipt_items_canvas is not None and self.receipt_items_window_id is not None:
            self.receipt_items_canvas.itemconfigure(self.receipt_items_window_id, width=event.width)
            self.receipt_items_canvas.configure(scrollregion=self.receipt_items_canvas.bbox("all"))

    def _resize_controls_panel(self, event) -> None:
        if self.controls_canvas is not None and self.controls_window_id is not None:
            self.controls_canvas.itemconfigure(self.controls_window_id, width=event.width)

    def _bind_mousewheel_support(self) -> None:
        self.bind_all("<MouseWheel>", self._on_mousewheel, add="+")
        self.bind_all("<Button-4>", self._on_mousewheel_linux, add="+")
        self.bind_all("<Button-5>", self._on_mousewheel_linux, add="+")

    def _on_mousewheel(self, event) -> None:
        widget = self.winfo_containing(event.x_root, event.y_root)
        target = self._find_scrollable_canvas(widget)
        if target is None:
            return

        delta = -1 * int(event.delta / 120) if event.delta else 0
        if delta:
            target.yview_scroll(delta, "units")

    def _on_mousewheel_linux(self, event) -> None:
        widget = self.winfo_containing(event.x_root, event.y_root)
        target = self._find_scrollable_canvas(widget)
        if target is None:
            return

        if event.num == 4:
            target.yview_scroll(-1, "units")
        elif event.num == 5:
            target.yview_scroll(1, "units")

    def _find_scrollable_canvas(self, widget) -> tk.Canvas | None:
        while widget is not None:
            if widget in {
                self.controls_canvas,
                self.preview_canvas,
                self.delivery_items_canvas,
                self.receipt_items_canvas,
            }:
                return widget
            widget = widget.master
        return self.controls_canvas

    def _update_mode_fields(self) -> None:
        if not all(
            [
                self.delivery_fields_frame,
                self.receipt_fields_frame,
                self.delivery_items_frame,
                self.receipt_items_frame,
            ]
        ):
            return

        self.delivery_fields_frame.grid_forget()
        self.receipt_fields_frame.grid_forget()
        self.delivery_items_frame.grid_forget()
        self.receipt_items_frame.grid_forget()

        mode = self.document_type.get()
        row = 0
        if mode in {"entrega", "ambos"}:
            self.delivery_fields_frame.grid(row=row, column=0, sticky="ew", pady=(0, 10))
            row += 1
        if mode in {"recebimento", "ambos"}:
            self.receipt_fields_frame.grid(row=row, column=0, sticky="ew")

        if mode == "entrega":
            self.delivery_items_frame.grid(row=0, column=0, sticky="nsew")
        elif mode == "recebimento":
            self.receipt_items_frame.grid(row=0, column=0, sticky="nsew")
        else:
            self.delivery_items_frame.grid(row=0, column=0, sticky="nsew", pady=(0, 10))
            self.receipt_items_frame.grid(row=1, column=0, sticky="nsew")

        self._refresh_preview()

    def _add_item_row(self, key: str) -> None:
        container = self.delivery_items_container if key == "delivery" else self.receipt_items_container
        rows = self.delivery_item_rows if key == "delivery" else self.receipt_item_rows
        if container is None:
            return
        if len(rows) >= self.MAX_ITEM_ROWS:
            messagebox.showinfo(
                "Limite de itens",
                f"Cada guia aceita no maximo {self.MAX_ITEM_ROWS} linhas de itens.",
            )
            self.status_text.set(
                f"Limite de {self.MAX_ITEM_ROWS} linhas atingido em {'entrega' if key == 'delivery' else 'recebimento'}."
            )
            return

        patrimony_var = tk.StringVar()
        description_var = tk.StringVar()

        patrimony_var.trace_add(
            "write",
            lambda *_args, p=patrimony_var, d=description_var: self._on_patrimony_change(p, d),
        )
        description_var.trace_add("write", self._on_preview_source_change)

        row_frame = ttk.Frame(container)
        row_frame.pack(fill="x", pady=3)
        row_frame.columnconfigure(0, weight=2)
        row_frame.columnconfigure(1, weight=3)

        ttk.Entry(row_frame, textvariable=patrimony_var, style="Field.TEntry").grid(row=0, column=0, sticky="ew")
        ttk.Entry(row_frame, textvariable=description_var, style="Field.TEntry").grid(
            row=0, column=1, sticky="ew", padx=(12, 0)
        )

        rows.append((patrimony_var, description_var, row_frame))
        self.status_text.set(
            f"{len(rows)}/{self.MAX_ITEM_ROWS} linhas em {'entrega' if key == 'delivery' else 'recebimento'}."
        )
        self._refresh_preview()

    def _on_patrimony_change(self, patrimony_var: tk.StringVar, description_var: tk.StringVar) -> None:
        patrimony_key = normalize_patrimony(patrimony_var.get())
        description = self.inventory_lookup.get(patrimony_key)

        if description and description_var.get().strip() != description:
            description_var.set(description)

        self._refresh_preview()

    def _remove_item_row(self, key: str) -> None:
        rows = self.delivery_item_rows if key == "delivery" else self.receipt_item_rows
        if not rows:
            return

        scroll_offset = None
        if self.controls_canvas is not None:
            self.update_idletasks()
            bbox = self.controls_canvas.bbox("all")
            viewport_height = max(self.controls_canvas.winfo_height(), 1)
            if bbox is not None:
                content_height = max(bbox[3] - bbox[1], 1)
                max_scroll = max(content_height - viewport_height, 0)
                scroll_offset = self.controls_canvas.yview()[0] * max_scroll

        _, _, row_frame = rows.pop()
        row_frame.destroy()
        self.status_text.set(
            f"{len(rows)}/{self.MAX_ITEM_ROWS} linhas em {'entrega' if key == 'delivery' else 'recebimento'}."
        )
        if self.controls_canvas is not None and scroll_offset is not None:
            self.after_idle(lambda offset=scroll_offset: self._restore_controls_scroll_offset(offset))
        self._refresh_preview()

    def _restore_controls_scroll_offset(self, scroll_offset: float) -> None:
        if self.controls_canvas is None:
            return

        self.update_idletasks()
        bbox = self.controls_canvas.bbox("all")
        viewport_height = max(self.controls_canvas.winfo_height(), 1)
        if bbox is None:
            return

        content_height = max(bbox[3] - bbox[1], 1)
        max_scroll = max(content_height - viewport_height, 0)
        if max_scroll <= 0:
            self.controls_canvas.yview_moveto(0.0)
            return

        target_fraction = min(max(scroll_offset / max_scroll, 0.0), 1.0)
        self.controls_canvas.yview_moveto(target_fraction)

    def _clear_items(self) -> None:
        for rows in [self.delivery_item_rows, self.receipt_item_rows]:
            for patrimony_var, description_var, _ in rows:
                patrimony_var.set("")
                description_var.set("")
        self.status_text.set("Linhas de itens limpas.")
        self._refresh_preview()

    def _print_documents(self) -> None:
        self._generate_documents_core(print_after=True)

    def _generate_documents_core(self, *, print_after: bool) -> None:
        mode = self.document_type.get()
        delivery_path = Path(self.delivery_template.get().strip())
        receipt_path = Path(self.receipt_template.get().strip())
        delivery_receiver_unit = self.delivery_receiver_unit.get().strip()
        receipt_sender_unit = self.receipt_sender_unit.get().strip()
        delivery_room = self.delivery_room.get().strip()
        receipt_room = self.receipt_room.get().strip()
        printer_name = self.selected_printer.get().strip()
        copies_text = self.print_copies.get().strip()

        if mode in {"entrega", "ambos"} and not delivery_path.is_file():
            messagebox.showerror("Arquivo invalido", "Selecione a guia de entrega .docx.")
            return

        if mode in {"recebimento", "ambos"} and not receipt_path.is_file():
            messagebox.showerror("Arquivo invalido", "Selecione a guia de recebimento .docx.")
            return

        if mode in {"entrega", "ambos"} and not delivery_receiver_unit:
            messagebox.showerror("UA Receptora vazia", "Informe a UA Receptora da guia de entrega.")
            return

        if mode in {"recebimento", "ambos"} and not receipt_sender_unit:
            messagebox.showerror("UA Remetente vazia", "Informe a UA Remetente da guia de recebimento.")
            return

        if mode in {"entrega", "ambos"} and not delivery_room:
            messagebox.showerror("Sala vazia", "Informe o numero da sala da guia de entrega.")
            return

        if mode in {"recebimento", "ambos"} and not receipt_room:
            messagebox.showerror("Sala vazia", "Informe o numero da sala da guia de recebimento.")
            return

        if print_after and not printer_name:
            messagebox.showerror("Impressora vazia", "Selecione uma impressora antes de imprimir.")
            return

        try:
            copies = int(copies_text)
        except ValueError:
            messagebox.showerror("Copias invalidas", "Informe um numero inteiro de copias.")
            return

        if copies < 1:
            messagebox.showerror("Copias invalidas", "A quantidade de copias deve ser pelo menos 1.")
            return

        try:
            delivery_items = self._collect_items(self.delivery_item_rows)
            receipt_items = self._collect_items(self.receipt_item_rows)
        except ValueError as exc:
            messagebox.showerror("Itens invalidos", str(exc))
            return

        if mode in {"entrega", "ambos"} and not delivery_items:
            messagebox.showerror("Itens vazios", "Preencha pelo menos um item na entrega.")
            return

        if mode in {"recebimento", "ambos"} and not receipt_items:
            messagebox.showerror("Itens vazios", "Preencha pelo menos um item no recebimento.")
            return

        timestamp = datetime.now().strftime("%Y%m%d-%H%M%S")
        keep_files = not print_after
        output_dir = Path.cwd() / "saida" if keep_files else None
        temp_dir_context = tempfile.TemporaryDirectory(prefix="guias-impressao-") if not keep_files else None

        try:
            generated_files: list[Path] = []
            working_dir = Path(temp_dir_context.name) if temp_dir_context is not None else output_dir
            assert working_dir is not None

            if mode in {"entrega", "ambos"}:
                delivery_output = working_dir / f"Guia-Entrega-{timestamp}.docx"
                generate_delivery_document(
                    template_path=delivery_path,
                    output_path=delivery_output,
                    receiver_unit=delivery_receiver_unit,
                    room_number=delivery_room,
                    items=delivery_items,
                )
                generated_files.append(delivery_output)

            if mode in {"recebimento", "ambos"}:
                receipt_output = working_dir / f"Guia-Recebimento-{timestamp}.docx"
                generate_receipt_document(
                    template_path=receipt_path,
                    output_path=receipt_output,
                    sender_unit=receipt_sender_unit,
                    room_number=receipt_room,
                    items=receipt_items,
                )
                generated_files.append(receipt_output)

            if print_after:
                print_docx_batch(generated_files, printer_name, copies=copies)
        except Exception as exc:
            if temp_dir_context is not None:
                temp_dir_context.cleanup()
            messagebox.showerror("Erro ao gerar", f"Nao foi possivel gerar os documentos.\n\n{exc}")
            return

        if print_after:
            action_text = "enviado(s) para impressao"
        else:
            action_text = "gerado(s)"
        if keep_files:
            names = "\n".join(path.name for path in generated_files)
            self.status_text.set(f"Documento(s) {action_text} em: {output_dir}")
            messagebox.showinfo(
                "Concluido",
                f"Arquivos {action_text} com sucesso.\n\nPasta: {output_dir}\n\n{names}",
            )
        else:
            if temp_dir_context is not None:
                temp_dir_context.cleanup()
            self.status_text.set(f"Documento(s) {action_text} sem salvar na pasta.")
            messagebox.showinfo(
                "Concluido",
                f"Arquivos {action_text} com sucesso.\n\nNada foi salvo na pasta do projeto.",
            )

    def _collect_items(self, rows: list[tuple[tk.StringVar, tk.StringVar, ttk.Frame]]) -> list[tuple[str, str]]:
        items: list[tuple[str, str]] = []
        for patrimony_var, description_var, _ in rows:
            patrimony = patrimony_var.get().strip()
            description = description_var.get().strip()

            if not patrimony and not description:
                continue

            if not patrimony or not description:
                raise ValueError("Cada linha precisa ter numero de patrimonio e descricao.")

            items.append((patrimony, description))

        return items

    def _bind_live_preview(self) -> None:
        for variable in [
            self.document_type,
            self.delivery_template,
            self.receipt_template,
            self.delivery_receiver_unit,
            self.receipt_sender_unit,
            self.delivery_room,
            self.receipt_room,
        ]:
            variable.trace_add("write", self._on_preview_source_change)

    def _on_preview_source_change(self, *_args) -> None:
        self._refresh_preview()

    def _animate_header_glow(self) -> None:
        if self.header_glow_canvas is None:
            return

        canvas = self.header_glow_canvas
        width = max(canvas.winfo_width(), 640)
        height = max(canvas.winfo_height(), 42)

        canvas.delete("all")
        canvas.create_rectangle(0, 0, width, height, fill=self.TEXT_COLOR, outline="")
        canvas.create_line(0, height - 1, width, height - 1, fill="#35597b", width=1)
        beam_width = 220
        lead_x = (self._header_glow_step % (width + beam_width)) - beam_width
        canvas.create_rectangle(lead_x, 0, lead_x + beam_width, height, fill="#2b5c88", outline="")
        canvas.create_rectangle(lead_x + 38, 2, lead_x + beam_width - 40, height - 2, fill="#8fb6d8", outline="")
        canvas.create_rectangle(lead_x + 78, 4, lead_x + beam_width - 82, height - 4, fill="#d7e8f4", outline="")
        self._header_glow_step += 8
        self.after(40, self._animate_header_glow)

    def _refresh_preview(self) -> None:
        if self.preview_canvas is None:
            return

        canvas = self.preview_canvas
        canvas.delete("all")
        canvas_width = max(canvas.winfo_width(), 720)
        canvas_height = max(canvas.winfo_height(), 720)
        canvas.create_rectangle(0, 0, canvas_width, canvas_height, fill=self.PREVIEW_BG, outline="")
        canvas.create_oval(-180, -120, 280, 260, fill="#e3eaf1", outline="")
        canvas.create_oval(
            canvas_width - 220,
            canvas_height - 180,
            canvas_width + 100,
            canvas_height + 120,
            fill="#dde5ed",
            outline="",
        )

        page_width = min(840, canvas_width - 50)
        page_width = max(page_width, 560)
        x = (canvas_width - page_width) / 2
        y = 18

        mode = self.document_type.get()
        if mode in {"entrega", "ambos"}:
            page_height = self._draw_guide_page(
                x=x,
                y=y,
                width=page_width,
                title="GUIA DE TRANSFERENCIA DE BENS PATRIMONIAIS - MOVEIS",
                top_left_label="UA Remetente:",
                top_left_value="DIVISAO DE PATRIMONIO",
                top_right_label="Sala:",
                top_right_value="ESTOQUE",
                bottom_left_label="UA Receptora:",
                bottom_left_value=self.delivery_receiver_unit.get().strip(),
                bottom_right_label="Sala:",
                bottom_right_value=self.delivery_room.get().strip(),
                item_rows=self._collect_partial_items(self.delivery_item_rows),
                signature_side="left",
            )
            y += page_height + 28

        if mode in {"recebimento", "ambos"}:
            page_height = self._draw_guide_page(
                x=x,
                y=y,
                width=page_width,
                title="GUIA DE TRANSFERENCIA DE BENS PATRIMONIAIS - MOVEIS",
                top_left_label="UA Remetente:",
                top_left_value=self.receipt_sender_unit.get().strip(),
                top_right_label="Sala:",
                top_right_value=self.receipt_room.get().strip(),
                bottom_left_label="UA Receptora:",
                bottom_left_value="DIVISAO DE PATRIMONIO",
                bottom_right_label="Sala:",
                bottom_right_value="ESTOQUE",
                item_rows=self._collect_partial_items(self.receipt_item_rows),
                signature_side="right",
            )
            y += page_height + 28

        canvas.configure(scrollregion=(0, 0, canvas_width, y))

    def _draw_guide_page(
        self,
        *,
        x: float,
        y: float,
        width: float,
        title: str,
        top_left_label: str,
        top_left_value: str,
        top_right_label: str,
        top_right_value: str,
        bottom_left_label: str,
        bottom_left_value: str,
        bottom_right_label: str,
        bottom_right_value: str,
        item_rows: list[tuple[str, str]],
        signature_side: str,
    ) -> float:
        assert self.preview_canvas is not None
        canvas = self.preview_canvas

        page_height = width * 1.43
        for offset, color in [(16, "#bfccda"), (10, "#ccd7e3"), (5, "#d6e0ea")]:
            canvas.create_rectangle(
                x + offset,
                y + offset,
                x + width + offset,
                y + page_height + offset,
                fill=color,
                outline="",
            )
        canvas.create_rectangle(x, y, x + width, y + page_height, fill="#ffffff", outline="#d2dbe6")

        self._draw_page_pattern(x, y, width)
        self._draw_brand(x + 28, y + 16)

        title_left = x + 34
        title_top = y + 78
        title_right = x + width - 28
        title_bottom = title_top + 28
        canvas.create_rectangle(title_left, title_top, title_right, title_bottom, outline="#111827", width=1)
        canvas.create_text(
            title_left + 6,
            title_top + 14,
            text=title,
            anchor="w",
            font=("Bahnschrift SemiBold", 13),
            fill="#111827",
        )

        margin = 30
        table_x = x + margin
        table_y = y + 126
        table_width = width - (margin * 2)

        info_heights = [24, 24]
        info_cols = [1602 / 10036, 6056 / 10036, 639 / 10036, 1739 / 10036]
        self._draw_table(
            table_x,
            table_y,
            table_width,
            info_heights,
            info_cols,
            [
                [top_left_label, top_left_value or "", top_right_label, top_right_value or ""],
                [bottom_left_label, bottom_left_value or "", bottom_right_label, bottom_right_value or ""],
            ],
            font=("Arial", 8),
            bold_cells={(0, 0), (0, 2), (1, 0), (1, 2)},
        )

        items_y = table_y + sum(info_heights) + 16
        table_bottom = self._draw_items_table(
            x=table_x,
            y=items_y,
            width=table_width,
            row_count=22,
            item_rows=item_rows,
        )

        self._draw_signature_boxes(
            x=table_x,
            y=table_bottom + 18,
            width=table_width,
            signature_side=signature_side,
        )
        canvas.create_text(
            x + width / 2,
            y + page_height - 16,
            text="PATRIMONIO: SALA T16 - TERREO - RAMAIS: 6188/6524/6528",
            font=("Segoe UI", 6),
            fill="#7d7d7d",
        )
        return page_height

    def _draw_page_pattern(self, x: float, y: float, width: float) -> None:
        assert self.preview_canvas is not None
        canvas = self.preview_canvas
        right = x + width - 8
        top = y - 10
        for radius in range(24, 180, 22):
            canvas.create_arc(
                right - radius * 2,
                top,
                right,
                top + radius * 2,
                start=92,
                extent=118,
                style="arc",
                outline="#e5e7eb",
                width=1,
            )

    def _draw_brand(self, x: float, y: float) -> None:
        assert self.preview_canvas is not None
        canvas = self.preview_canvas
        canvas.create_text(x, y, text="Alesp", anchor="nw", font=("Bahnschrift SemiBold", 24), fill="#5d2a6c")
        canvas.create_text(
            x + 2,
            y + 31,
            text="ASSEMBLEIA LEGISLATIVA\nDO ESTADO DE SAO PAULO",
            anchor="nw",
            font=("Segoe UI", 5, "bold"),
            fill="#6c6c6c",
        )
        for offset, color in [(0, "#26a9e0"), (8, "#39b54a"), (16, "#26a9e0"), (24, "#39b54a"), (32, "#26a9e0")]:
            canvas.create_arc(
                x + 42 + offset,
                y - 12 - offset / 4,
                x + 126 + offset,
                y + 42 + offset / 4,
                start=298,
                extent=118,
                style="arc",
                outline=color,
                width=2,
            )

    def _draw_signature_boxes(self, *, x: float, y: float, width: float, signature_side: str) -> None:
        assert self.preview_canvas is not None
        canvas = self.preview_canvas

        gap = 18
        box_width = (width - gap) / 2
        box_height = 126
        left_x = x
        right_x = x + box_width + gap

        canvas.create_rectangle(left_x, y, left_x + box_width, y + box_height, outline="#222222", width=1)
        canvas.create_rectangle(right_x, y, right_x + box_width, y + box_height, outline="#222222", width=1)

        today = datetime.now().strftime("%d/%m/%Y")
        left_box = {
            "title": f"Data de envio: {today}",
            "matricula": "Matricula: 31356",
            "nome": "Nome: Felipe Solano Silva Lyra",
            "role": "Assinatura - Remetente",
            "signed": signature_side == "left",
        }
        right_box = {
            "title": "Data de recebimento: ____/____/______",
            "matricula": "Matricula: ______________________",
            "nome": "Nome: __________________________",
            "role": "Assinatura - Receptor",
            "signed": signature_side == "right",
        }

        if signature_side == "right":
            left_box = {
                "title": "Data de envio: ____/____/______",
                "matricula": "Matricula: ______________________",
                "nome": "Nome: __________________________",
                "role": "Assinatura - Remetente",
                "signed": False,
            }
            right_box = {
                "title": f"Data de recebimento: {today}",
                "matricula": "Matricula: 31356",
                "nome": "Nome: Felipe Solano Silva Lyra",
                "role": "Assinatura - Receptor",
                "signed": True,
            }

        self._draw_signature_box_contents(x=left_x, y=y, width=box_width, **left_box)
        self._draw_signature_box_contents(x=right_x, y=y, width=box_width, **right_box)

    def _draw_signature_box_contents(
        self,
        *,
        x: float,
        y: float,
        width: float,
        title: str,
        matricula: str,
        nome: str,
        role: str,
        signed: bool,
    ) -> None:
        assert self.preview_canvas is not None
        canvas = self.preview_canvas

        canvas.create_text(x + 8, y + 16, text=title, anchor="nw", font=("Segoe UI", 8))
        canvas.create_text(x + 8, y + 42, text=matricula, anchor="nw", font=("Segoe UI", 8))
        canvas.create_text(x + 8, y + 68, text=nome, anchor="nw", font=("Segoe UI", 8))
        canvas.create_line(x + 10, y + 112, x + width - 10, y + 112, fill="#222222")
        canvas.create_text(x + width / 2, y + 114, text=role, anchor="n", font=("Segoe UI", 7))

        if signed:
            canvas.create_text(
                x + width / 2,
                y + 84,
                text="Felipe Solano",
                font=("Segoe Script", 20),
                fill="#6f6f6f",
            )

    def _draw_items_table(
        self,
        *,
        x: float,
        y: float,
        width: float,
        row_count: int,
        item_rows: list[tuple[str, str]],
    ) -> float:
        assert self.preview_canvas is not None
        canvas = self.preview_canvas

        patrimony_width = width * (1866 / (1866 + 8206))
        description_width = width - patrimony_width
        header_height = 22
        row_height = 22

        x0 = x
        x1 = x0 + patrimony_width
        x2 = x + width

        for start, end in [(x0, x1), (x1, x2)]:
            canvas.create_rectangle(start, y, end, y + header_height, outline="#222222", width=1)

        canvas.create_text((x0 + x1) / 2, y + header_height / 2, text="Nº Patrimônio", font=("Arial", 8, "bold"))
        canvas.create_text((x1 + x2) / 2, y + header_height / 2, text="Descrição do bem", font=("Arial", 8, "bold"))

        current_y = y + header_height
        for patrimony, description in self._collect_preview_rows(item_rows, row_count):
            canvas.create_rectangle(x0, current_y, x1, current_y + row_height, outline="#222222", width=1)
            canvas.create_rectangle(x1, current_y, x2, current_y + row_height, outline="#222222", width=1)

            if patrimony:
                canvas.create_text(
                    x0 + 5,
                    current_y + row_height / 2,
                    text=patrimony,
                    anchor="w",
                    font=("Arial", 7),
                    width=max(patrimony_width - 10, 10),
                )
            if description:
                canvas.create_text(
                    x1 + 5,
                    current_y + row_height / 2,
                    text=description,
                    anchor="w",
                    font=("Arial", 7),
                    width=max(description_width - 10, 10),
                )
            current_y += row_height

        return current_y

    def _draw_table(
        self,
        x: float,
        y: float,
        width: float,
        row_heights: list[float],
        col_widths: list[float],
        data: list[list[str]],
        *,
        font: tuple[str, int] | tuple[str, int, str],
        bold_cells: set[tuple[int, int]],
    ) -> None:
        assert self.preview_canvas is not None
        canvas = self.preview_canvas

        absolute_widths = [width * part for part in col_widths]
        current_y = y
        for row_index, row_height in enumerate(row_heights):
            current_x = x
            for col_index, cell_width in enumerate(absolute_widths):
                canvas.create_rectangle(
                    current_x,
                    current_y,
                    current_x + cell_width,
                    current_y + row_height,
                    outline="#222222",
                    width=1,
                )
                text_value = data[row_index][col_index]
                cell_font = font
                if (row_index, col_index) in bold_cells and len(font) == 2:
                    cell_font = (font[0], font[1], "bold")
                if text_value:
                    canvas.create_text(
                        current_x + 5,
                        current_y + row_height / 2,
                        text=text_value,
                        anchor="w",
                        font=cell_font,
                        width=max(cell_width - 10, 10),
                    )
                current_x += cell_width
            current_y += row_height

    def _collect_preview_rows(self, items: list[tuple[str, str]], limit: int) -> list[tuple[str, str]]:
        rows = list(items[:limit])
        while len(rows) < limit:
            rows.append(("", ""))
        return rows

    def _collect_partial_items(self, rows: list[tuple[tk.StringVar, tk.StringVar, ttk.Frame]]) -> list[tuple[str, str]]:
        items: list[tuple[str, str]] = []
        for patrimony_var, description_var, _ in rows:
            patrimony = patrimony_var.get().strip()
            description = description_var.get().strip()
            if not patrimony and not description:
                continue
            items.append((patrimony, description))
        return items
