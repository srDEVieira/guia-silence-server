from __future__ import annotations

import json
import os
import sys
import urllib.error
import urllib.parse
import urllib.request
from datetime import datetime
from pathlib import Path
import tkinter as tk
from tkinter import messagebox, simpledialog, ttk


ROOT_DIR = Path(__file__).resolve().parent
CONFIG_PATH = ROOT_DIR / "admin_panel_config.json"


def _ensure_tk_runtime() -> None:
    if os.name != "nt":
        return
    if os.getenv("TCL_LIBRARY") and os.getenv("TK_LIBRARY"):
        return

    candidates = [
        Path(sys.base_prefix) / "tcl",
        Path(sys.executable).resolve().parent.parent / "tcl",
        Path.home() / "AppData" / "Local" / "Programs" / "Python" / "Python313" / "tcl",
    ]
    for base in candidates:
        tcl_dir = base / "tcl8.6"
        tk_dir = base / "tk8.6"
        if tcl_dir.is_dir() and tk_dir.is_dir():
            os.environ.setdefault("TCL_LIBRARY", str(tcl_dir))
            os.environ.setdefault("TK_LIBRARY", str(tk_dir))
            return


class AdminPanelApp:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("Painel Admin - Guia Patrimonial")
        self.root.geometry("1120x700")
        self.root.minsize(980, 620)

        self.base_url_var = tk.StringVar()
        self.token_var = tk.StringVar()
        self.status_var = tk.StringVar(value="Status: nao testado")
        self.total_var = tk.StringVar(value="Dispositivos: 0 | Perfis: 0")

        self._build_ui()
        self._load_config()
        self._refresh_health()
        self._refresh_devices()
        self._refresh_profiles()

    def _build_ui(self) -> None:
        self.root.configure(bg="#0f172a")

        top = tk.Frame(self.root, bg="#0f172a", padx=14, pady=12)
        top.pack(fill="x")

        title = tk.Label(top, text="Painel Admin Remoto", bg="#0f172a", fg="#e2e8f0", font=("Segoe UI", 16, "bold"))
        title.grid(row=0, column=0, columnspan=6, sticky="w", pady=(0, 10))

        tk.Label(top, text="Servidor", bg="#0f172a", fg="#94a3b8", font=("Segoe UI", 9, "bold")).grid(row=1, column=0, sticky="w")
        tk.Entry(top, textvariable=self.base_url_var, width=54).grid(row=2, column=0, columnspan=3, sticky="we", padx=(0, 8))

        tk.Label(top, text="Token Admin", bg="#0f172a", fg="#94a3b8", font=("Segoe UI", 9, "bold")).grid(row=1, column=3, sticky="w")
        tk.Entry(top, textvariable=self.token_var, show="*", width=36).grid(row=2, column=3, sticky="we", padx=(0, 8))

        ttk.Button(top, text="Salvar", command=self._save_config).grid(row=2, column=4, padx=(0, 8))
        ttk.Button(top, text="Testar", command=self._refresh_health).grid(row=2, column=5)

        info = tk.Frame(self.root, bg="#0f172a", padx=14, pady=0)
        info.pack(fill="x")
        self.status_label = tk.Label(info, textvariable=self.status_var, bg="#0f172a", fg="#22c55e", font=("Segoe UI", 10, "bold"))
        self.status_label.pack(side="left")
        tk.Label(info, textvariable=self.total_var, bg="#0f172a", fg="#93c5fd", font=("Segoe UI", 10, "bold")).pack(side="right")

        notebook_wrap = tk.Frame(self.root, bg="#0f172a", padx=14, pady=12)
        notebook_wrap.pack(fill="both", expand=True)
        self.notebook = ttk.Notebook(notebook_wrap)
        self.notebook.pack(fill="both", expand=True)

        self.tab_devices = tk.Frame(self.notebook, bg="#f8fafc")
        self.tab_profiles = tk.Frame(self.notebook, bg="#f8fafc")
        self.notebook.add(self.tab_devices, text="Dispositivos")
        self.notebook.add(self.tab_profiles, text="Perfis")

        self._build_devices_tab()
        self._build_profiles_tab()

    def _build_devices_tab(self) -> None:
        actions = tk.Frame(self.tab_devices, bg="#f8fafc", padx=12, pady=10)
        actions.pack(fill="x")
        ttk.Button(actions, text="Atualizar", command=self._refresh_devices).pack(side="left")
        ttk.Button(actions, text="Bloquear", command=self._block_selected).pack(side="left", padx=8)
        ttk.Button(actions, text="Desbloquear", command=self._unblock_selected).pack(side="left")
        ttk.Button(actions, text="Exportar CSV", command=self._export_devices_csv).pack(side="right")

        table_wrap = tk.Frame(self.tab_devices, bg="#f8fafc", padx=12, pady=(0, 12))
        table_wrap.pack(fill="both", expand=True)

        columns = ("device_id", "machine_name", "user_name", "blocked", "last_seen")
        self.devices_tree = ttk.Treeview(table_wrap, columns=columns, show="headings")
        for col, text, width in (
            ("device_id", "Device ID", 280),
            ("machine_name", "Maquina", 180),
            ("user_name", "Usuario", 150),
            ("blocked", "Bloqueado", 100),
            ("last_seen", "Ultimo acesso", 220),
        ):
            self.devices_tree.heading(col, text=text)
            self.devices_tree.column(col, width=width, anchor="w")
        self.devices_tree.column("blocked", anchor="center")

        ys = ttk.Scrollbar(table_wrap, orient="vertical", command=self.devices_tree.yview)
        xs = ttk.Scrollbar(table_wrap, orient="horizontal", command=self.devices_tree.xview)
        self.devices_tree.configure(yscrollcommand=ys.set, xscrollcommand=xs.set)
        self.devices_tree.grid(row=0, column=0, sticky="nsew")
        ys.grid(row=0, column=1, sticky="ns")
        xs.grid(row=1, column=0, sticky="ew")
        table_wrap.grid_rowconfigure(0, weight=1)
        table_wrap.grid_columnconfigure(0, weight=1)

    def _build_profiles_tab(self) -> None:
        actions = tk.Frame(self.tab_profiles, bg="#f8fafc", padx=12, pady=10)
        actions.pack(fill="x")
        ttk.Button(actions, text="Atualizar", command=self._refresh_profiles).pack(side="left")
        ttk.Button(actions, text="Adicionar", command=self._add_profile).pack(side="left", padx=8)
        ttk.Button(actions, text="Editar", command=self._edit_profile).pack(side="left")
        ttk.Button(actions, text="Excluir", command=self._delete_profile).pack(side="left", padx=8)

        table_wrap = tk.Frame(self.tab_profiles, bg="#f8fafc", padx=12, pady=(0, 12))
        table_wrap.pack(fill="both", expand=True)

        columns = ("profile_id", "display_name", "active", "sort_order", "accent_color", "hero_bg_url")
        self.profiles_tree = ttk.Treeview(table_wrap, columns=columns, show="headings")
        for col, text, width in (
            ("profile_id", "ID", 120),
            ("display_name", "Nome", 170),
            ("active", "Ativo", 80),
            ("sort_order", "Ordem", 80),
            ("accent_color", "Cor", 120),
            ("hero_bg_url", "Plano de fundo (URL)", 420),
        ):
            self.profiles_tree.heading(col, text=text)
            self.profiles_tree.column(col, width=width, anchor="w")
        self.profiles_tree.column("active", anchor="center")
        self.profiles_tree.column("sort_order", anchor="center")

        ys = ttk.Scrollbar(table_wrap, orient="vertical", command=self.profiles_tree.yview)
        xs = ttk.Scrollbar(table_wrap, orient="horizontal", command=self.profiles_tree.xview)
        self.profiles_tree.configure(yscrollcommand=ys.set, xscrollcommand=xs.set)
        self.profiles_tree.grid(row=0, column=0, sticky="nsew")
        ys.grid(row=0, column=1, sticky="ns")
        xs.grid(row=1, column=0, sticky="ew")
        table_wrap.grid_rowconfigure(0, weight=1)
        table_wrap.grid_columnconfigure(0, weight=1)

    def _set_status(self, text: str, ok: bool | None = None) -> None:
        self.status_var.set(f"Status: {text}")
        color = "#93c5fd"
        if ok is True:
            color = "#22c55e"
        elif ok is False:
            color = "#ef4444"
        self.status_label.configure(fg=color)

    def _load_config(self) -> None:
        if not CONFIG_PATH.exists():
            self.base_url_var.set("https://guia-silence-server.onrender.com")
            self.token_var.set("")
            return
        try:
            data = json.loads(CONFIG_PATH.read_text(encoding="utf-8"))
            self.base_url_var.set(str(data.get("base_url", "")).strip())
            self.token_var.set(str(data.get("admin_token", "")).strip())
        except Exception:
            self.base_url_var.set("https://guia-silence-server.onrender.com")
            self.token_var.set("")

    def _save_config(self) -> None:
        payload = {
            "base_url": self.base_url_var.get().strip().rstrip("/"),
            "admin_token": self.token_var.get().strip(),
        }
        CONFIG_PATH.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
        self._set_status("configuracao salva", True)

    def _headers(self) -> dict[str, str]:
        token = self.token_var.get().strip()
        return {"X-Admin-Token": token} if token else {}

    def _base_url(self) -> str:
        return self.base_url_var.get().strip().rstrip("/")

    def _request(self, method: str, path: str, body: dict | None = None, *, admin: bool = True) -> dict:
        base = self._base_url()
        if not base:
            raise RuntimeError("Informe a URL do servidor.")
        url = f"{base}{path}"
        headers: dict[str, str] = self._headers() if admin else {}
        data = None
        if body is not None:
            headers["Content-Type"] = "application/json"
            data = json.dumps(body).encode("utf-8")
        req = urllib.request.Request(url=url, data=data, headers=headers, method=method)
        try:
            with urllib.request.urlopen(req, timeout=8) as resp:
                raw = resp.read().decode("utf-8")
        except urllib.error.HTTPError as exc:
            detail = exc.read().decode("utf-8", errors="ignore")
            raise RuntimeError(f"HTTP {exc.code}: {detail or exc.reason}") from exc
        except urllib.error.URLError as exc:
            raise RuntimeError(f"Falha de conexao: {exc.reason}") from exc
        if not raw.strip():
            return {}
        try:
            return json.loads(raw)
        except json.JSONDecodeError:
            return {}

    def _refresh_health(self) -> None:
        try:
            data = self._request("GET", "/health", admin=False)
            self._set_status(f"online (storage: {data.get('storage', 'desconhecido')})", True)
        except Exception as exc:
            self._set_status(str(exc), False)

    def _refresh_devices(self) -> None:
        try:
            data = self._request("GET", "/admin/devices")
            devices = data.get("devices", []) or []
            self._populate_devices(devices)
            self._set_status("dispositivos atualizados", True)
        except Exception as exc:
            self._populate_devices([])
            self._set_status(str(exc), False)
        self._update_totals()

    def _refresh_profiles(self) -> None:
        try:
            data = self._request("GET", "/admin/profiles")
            profiles = data.get("profiles", []) or []
            self._populate_profiles(profiles)
            self._set_status("perfis atualizados", True)
        except Exception as exc:
            self._populate_profiles([])
            self._set_status(str(exc), False)
        self._update_totals()

    def _populate_devices(self, devices: list[dict]) -> None:
        for iid in self.devices_tree.get_children():
            self.devices_tree.delete(iid)
        for d in sorted(devices, key=lambda x: str(x.get("last_seen", "")), reverse=True):
            blocked = bool(d.get("blocked", False))
            self.devices_tree.insert(
                "",
                "end",
                values=(
                    d.get("device_id", ""),
                    d.get("machine_name", ""),
                    d.get("user_name", ""),
                    "Sim" if blocked else "Nao",
                    d.get("last_seen", ""),
                ),
                tags=("blocked",) if blocked else ("active",),
            )
        self.devices_tree.tag_configure("blocked", foreground="#ef4444")
        self.devices_tree.tag_configure("active", foreground="#15803d")

    def _populate_profiles(self, profiles: list[dict]) -> None:
        for iid in self.profiles_tree.get_children():
            self.profiles_tree.delete(iid)
        for p in sorted(profiles, key=lambda x: (int(x.get("sort_order", 0) or 0), str(x.get("display_name", "")))):
            self.profiles_tree.insert(
                "",
                "end",
                values=(
                    p.get("profile_id", ""),
                    p.get("display_name", ""),
                    "Sim" if bool(p.get("active", True)) else "Nao",
                    p.get("sort_order", 0),
                    p.get("accent_color", ""),
                    p.get("hero_bg_url", ""),
                ),
            )

    def _update_totals(self) -> None:
        self.total_var.set(
            f"Dispositivos: {len(self.devices_tree.get_children())} | Perfis: {len(self.profiles_tree.get_children())}"
        )

    def _selected_device_id(self) -> str:
        selected = self.devices_tree.selection()
        if not selected:
            return ""
        return str(self.devices_tree.item(selected[0], "values")[0])

    def _change_block(self, block: bool) -> None:
        device_id = self._selected_device_id()
        if not device_id:
            messagebox.showwarning("Painel Admin", "Selecione um dispositivo.")
            return
        action = "bloquear" if block else "desbloquear"
        if not messagebox.askyesno("Confirmacao", f"Deseja {action}?\n\n{device_id}"):
            return
        path = f"/admin/block/{urllib.parse.quote(device_id)}" if block else f"/admin/unblock/{urllib.parse.quote(device_id)}"
        try:
            self._request("POST", path)
            self._refresh_devices()
            self._set_status(f"dispositivo {action}ado", True)
        except Exception as exc:
            self._set_status(str(exc), False)

    def _block_selected(self) -> None:
        self._change_block(True)

    def _unblock_selected(self) -> None:
        self._change_block(False)

    def _export_devices_csv(self) -> None:
        rows = [self.devices_tree.item(iid, "values") for iid in self.devices_tree.get_children()]
        if not rows:
            messagebox.showinfo("Painel Admin", "Nenhum dispositivo para exportar.")
            return
        export_dir = ROOT_DIR / "admin_exports"
        export_dir.mkdir(parents=True, exist_ok=True)
        target = export_dir / f"devices-{datetime.now().strftime('%Y%m%d-%H%M%S')}.csv"
        lines = ["device_id,machine_name,user_name,blocked,last_seen\n"]
        for values in rows:
            safe = [str(v).replace('"', '""') for v in values]
            lines.append(",".join([f'"{v}"' for v in safe]) + "\n")
        target.write_text("".join(lines), encoding="utf-8")
        messagebox.showinfo("Painel Admin", f"Exportado:\n{target}")

    def _selected_profile(self) -> tuple[str, str, str, str, str, str] | None:
        selected = self.profiles_tree.selection()
        if not selected:
            return None
        return tuple(self.profiles_tree.item(selected[0], "values"))

    def _profile_dialog(self, title: str, current: dict[str, Any] | None = None) -> dict[str, Any] | None:
        current = current or {}
        name = simpledialog.askstring(title, "Nome do perfil:", initialvalue=str(current.get("display_name", "")), parent=self.root)
        if name is None:
            return None
        if not name.strip():
            messagebox.showwarning("Painel Admin", "Nome e obrigatorio.")
            return None
        bg_url = simpledialog.askstring(
            title,
            "URL do plano de fundo (opcional):",
            initialvalue=str(current.get("hero_bg_url", "")),
            parent=self.root,
        )
        if bg_url is None:
            return None
        accent = simpledialog.askstring(
            title,
            "Cor de destaque (hex, opcional. ex: #114c78):",
            initialvalue=str(current.get("accent_color", "")),
            parent=self.root,
        )
        if accent is None:
            return None
        sort_order_raw = simpledialog.askstring(
            title,
            "Ordem (numero inteiro):",
            initialvalue=str(current.get("sort_order", 0)),
            parent=self.root,
        )
        if sort_order_raw is None:
            return None
        try:
            sort_order = int(sort_order_raw or 0)
        except ValueError:
            sort_order = 0
        active = messagebox.askyesno(title, "Perfil ativo?")
        return {
            "display_name": name.strip(),
            "hero_bg_url": (bg_url or "").strip(),
            "accent_color": (accent or "").strip(),
            "sort_order": sort_order,
            "active": active,
        }

    def _add_profile(self) -> None:
        payload = self._profile_dialog("Novo perfil")
        if payload is None:
            return
        try:
            self._request("POST", "/admin/profiles", body=payload)
            self._refresh_profiles()
            self._set_status("perfil criado", True)
        except Exception as exc:
            self._set_status(str(exc), False)

    def _edit_profile(self) -> None:
        selected = self._selected_profile()
        if not selected:
            messagebox.showwarning("Painel Admin", "Selecione um perfil.")
            return
        profile_id, name, active, sort_order, accent_color, hero_bg_url = selected
        payload = self._profile_dialog(
            "Editar perfil",
            {
                "display_name": name,
                "hero_bg_url": hero_bg_url,
                "accent_color": accent_color,
                "sort_order": sort_order,
                "active": active == "Sim",
            },
        )
        if payload is None:
            return
        try:
            self._request("PUT", f"/admin/profiles/{urllib.parse.quote(profile_id)}", body=payload)
            self._refresh_profiles()
            self._set_status("perfil atualizado", True)
        except Exception as exc:
            self._set_status(str(exc), False)

    def _delete_profile(self) -> None:
        selected = self._selected_profile()
        if not selected:
            messagebox.showwarning("Painel Admin", "Selecione um perfil.")
            return
        profile_id, name, *_ = selected
        if not messagebox.askyesno("Excluir perfil", f"Excluir perfil '{name}' ({profile_id})?"):
            return
        try:
            self._request("DELETE", f"/admin/profiles/{urllib.parse.quote(profile_id)}")
            self._refresh_profiles()
            self._set_status("perfil excluido", True)
        except Exception as exc:
            self._set_status(str(exc), False)


def main() -> None:
    _ensure_tk_runtime()
    root = tk.Tk()
    style = ttk.Style(root)
    try:
        style.theme_use("clam")
    except tk.TclError:
        pass
    AdminPanelApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
