"""
Generador de Plantilla RIPS - interfaz grafica de escritorio.

Toma los archivos planos AF0/US0/AP0/AD0 (y opcionalmente CT0 y el Excel
de pago moderador) y llena la plantilla oficial RIPS (.xlsm) en las hojas
transaccion, usuarios y consultas, sin modificar su estructura. Tambien
puede convertir la plantilla resultante a JSON usando el conversor oficial
de SISPRO (fevrips.sispro.gov.co), sin necesidad de navegador, y generar
el archivo plano de importacion (World Office).
"""
import json
import os
import shutil
import sys
import threading
import traceback
from datetime import date, datetime
from tkinter import (
    Tk, Toplevel, Canvas, Frame, Label, Button, Entry, StringVar, filedialog, messagebox, END,
)
from tkinter import ttk
from tkinter.scrolledtext import ScrolledText

from tkcalendar import DateEntry

import rips_core


def resource_path(relative_path):
    """Ubica un recurso tanto en desarrollo como empaquetado con PyInstaller (--onefile)."""
    base_path = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base_path, relative_path)


TEMPLATE_PATH = resource_path(os.path.join("resources", "plantilla_fev-rips_v_1_0.xlsm"))
NITS_CLIENTES_PATH = resource_path(os.path.join("resources", "nits_clientes.json"))
ICON_PATH = resource_path(os.path.join("resources", "logo.ico"))


def _load_nits_clientes():
    """Carga {nit: nombre} desde el listado de clientes. Si falta el archivo,
    la app sigue funcionando pero sin desplegable (solo NIT manual)."""
    try:
        with open(NITS_CLIENTES_PATH, encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}


NITS_CLIENTES = _load_nits_clientes()
NOMBRES_A_NIT = {nombre: nit for nit, nombre in NITS_CLIENTES.items()}
NOMBRES_CLIENTES_ORDENADOS = sorted(NOMBRES_A_NIT.keys(), key=str.upper)

# ---------------------------------------------------------------------------
# Paleta: rojo + gris, colores corporativos MICRODIAG
# ---------------------------------------------------------------------------
PRIMARIO = "#A32026"
PRIMARIO_OSCURO = "#6E1518"
PRIMARIO_SUAVE = "#E4E4E5"
ACENTO = "#58595B"
ACENTO_CLARO = "#808285"
FONDO = "#F2F2F3"
TARJETA = "#FFFFFF"
BLANCO = "#FFFFFF"
GRIS_CLARO = "#EAEAEA"
GRIS_TEXTO = "#6B6B6B"
TEXTO = "#2B2B2B"
VERDE_OK = "#22A559"
ROJO_WARN = "#A32026"

# ---------------------------------------------------------------------------
# Licencia / periodo de prueba
# ---------------------------------------------------------------------------
LICENSE_CODE = "540-149-RIPS"
TRIAL_DEADLINE = datetime(2026, 9, 15, 7, 0, 0)
LICENSE_DIR = os.path.join(os.environ.get("APPDATA") or os.path.expanduser("~"), "GeneradorRIPS")
LICENSE_FILE = os.path.join(LICENSE_DIR, "license.ok")


def _licencia_activa():
    return os.path.exists(LICENSE_FILE)


def _guardar_licencia():
    os.makedirs(LICENSE_DIR, exist_ok=True)
    with open(LICENSE_FILE, "w", encoding="utf-8") as f:
        f.write("OK")


def _pedir_licencia(root):
    """Modal bloqueante: pide el codigo de licencia. Devuelve True si se activo
    correctamente (y guarda la licencia para no volver a pedirla), False si el
    usuario cerro la ventana sin activarla."""
    resultado = {"ok": False}

    win = Toplevel(root)
    win.title("Licencia requerida")
    win.configure(bg=TARJETA)
    win.resizable(False, False)
    win.protocol("WM_DELETE_WINDOW", lambda: win.destroy())

    Label(
        win, text="Periodo de prueba finalizado", bg=TARJETA, fg=PRIMARIO_OSCURO,
        font=("Segoe UI", 13, "bold"),
    ).pack(padx=36, pady=(26, 6))
    Label(
        win, text="Ingresa el codigo de licencia para continuar usando el programa.",
        bg=TARJETA, fg=GRIS_TEXTO, font=("Segoe UI", 9), wraplength=320, justify="center",
    ).pack(padx=36, pady=(0, 16))

    codigo_var = StringVar()
    entry = Entry(
        win, textvariable=codigo_var, font=("Segoe UI", 12), justify="center",
        width=20, relief="flat", highlightthickness=1,
        highlightbackground=PRIMARIO, highlightcolor=PRIMARIO,
    )
    entry.pack(padx=36, pady=(0, 6))

    error_var = StringVar(value="")
    Label(
        win, textvariable=error_var, bg=TARJETA, fg=ROJO_WARN, font=("Segoe UI", 9),
    ).pack(pady=(0, 4))

    def _validar():
        if codigo_var.get().strip().upper() == LICENSE_CODE:
            _guardar_licencia()
            resultado["ok"] = True
            win.destroy()
        else:
            error_var.set("Codigo incorrecto. Intenta de nuevo.")
            codigo_var.set("")

    entry.bind("<Return>", lambda e: _validar())

    botones = Frame(win, bg=TARJETA)
    botones.pack(pady=(6, 24))
    Button(
        botones, text="Activar", command=_validar,
        bg=ACENTO, fg=BLANCO, activebackground=ACENTO_CLARO, activeforeground=BLANCO,
        relief="flat", font=("Segoe UI", 10, "bold"), padx=20, pady=7, cursor="hand2",
    ).pack(side="left", padx=6)
    Button(
        botones, text="Salir", command=win.destroy,
        bg=GRIS_CLARO, fg=GRIS_TEXTO, relief="flat",
        font=("Segoe UI", 10), padx=20, pady=7, cursor="hand2",
    ).pack(side="left", padx=6)

    win.update_idletasks()
    w, h = win.winfo_width(), win.winfo_height()
    ws, hs = win.winfo_screenwidth(), win.winfo_screenheight()
    win.geometry(f"+{(ws - w) // 2}+{(hs - h) // 2}")

    win.transient(root)
    win.lift()
    win.attributes("-topmost", True)
    win.after(200, lambda: win.attributes("-topmost", False))
    win.focus_force()
    win.grab_set()
    entry.focus_set()
    root.wait_window(win)
    return resultado["ok"]


class RipsApp(Frame):
    def __init__(self, root, on_volver=None):
        super().__init__(root, bg=FONDO)
        self.root = root
        self.on_volver = on_volver
        self._destroyed = False
        self._licencia_job = None

        self.paths = {
            "af0": StringVar(),
            "us0": StringVar(),
            "ap0": StringVar(),
            "ad0": StringVar(),
            "ct0": StringVar(),
            "pago": StringVar(),
            "xml_factura": StringVar(),
            "output": StringVar(),
        }
        self.txt_status_vars = {}
        self.prefijo_var = StringVar()
        self.consecutivo_var = StringVar()
        self.tipo_usuario_var = StringVar(value=list(rips_core.TIPO_USUARIO_OPCIONES.keys())[0])
        self.empresa_var = StringVar()
        self.nit_cliente_var = StringVar()
        self.nit_resuelto_var = StringVar(value="Selecciona una empresa o escribe el NIT (9 digitos)")
        self.date_entry_factura = None
        self.last_output_path = None
        self.last_json_path = None

        self._build_header()
        self._build_footer()
        self._build_main()

    # ------------------------------------------------------------------
    def _build_header(self):
        header = Frame(self, bg=PRIMARIO, height=84)
        header.pack(fill="x", side="top")
        header.pack_propagate(False)

        if self.on_volver is not None:
            Button(
                header, text="◀ Menú principal", command=self.on_volver,
                bg=PRIMARIO_OSCURO, fg=BLANCO, activebackground="#8C262B", activeforeground=BLANCO,
                relief="flat", font=("Segoe UI", 9, "bold"), padx=12, pady=6, cursor="hand2",
            ).pack(side="left", padx=(24, 12), pady=10)

        Label(
            header, text="🧾 GENERADOR DE RIPS ELECTRONICOS Y XML - LABORATORIO ESPECIALIZADO MICRODIAG",
            bg=PRIMARIO, fg=BLANCO,
            font=("Segoe UI", 14, "bold"),
        ).pack(side="left", padx=(0 if self.on_volver is not None else 24), pady=10)

        self.licencia_var = StringVar()
        self.lbl_licencia = Label(
            header, textvariable=self.licencia_var, bg=PRIMARIO, fg=BLANCO,
            font=("Segoe UI", 10, "bold"),
        )
        self.lbl_licencia.pack(side="right", padx=24)
        self._actualizar_licencia()

        underline = Frame(self, bg=ACENTO, height=4)
        underline.pack(fill="x")

    def _actualizar_licencia(self):
        if self._destroyed:
            return
        if _licencia_activa():
            self.licencia_var.set("✓ Licencia activa")
            self.lbl_licencia.configure(fg="#C9FFDC")
            return

        restante = TRIAL_DEADLINE - datetime.now()
        if restante.total_seconds() <= 0:
            self.licencia_var.set("⚠ Periodo de prueba vencido")
            self.lbl_licencia.configure(fg=ACENTO_CLARO)
            return

        dias = restante.days
        horas, resto = divmod(restante.seconds, 3600)
        minutos, segundos = divmod(resto, 60)
        if dias > 0:
            texto = f"⏳ Prueba vence en {dias}d {horas}h {minutos}m"
        else:
            texto = f"⏳ Prueba vence en {horas}h {minutos}m {segundos}s"
        self.licencia_var.set(texto)
        self.lbl_licencia.configure(fg=BLANCO)
        self._licencia_job = self.root.after(1000, self._actualizar_licencia)

    def on_navigate_away(self):
        """Llamado por el controlador de vistas antes de destruir este frame,
        para cancelar temporizadores y bindings globales pendientes (si no,
        siguen disparando sobre widgets ya destruidos)."""
        self._destroyed = True
        if self._licencia_job is not None:
            try:
                self.root.after_cancel(self._licencia_job)
            except Exception:
                pass
        try:
            self.root.unbind_all("<MouseWheel>")
        except Exception:
            pass

    # ------------------------------------------------------------------
    def _build_main(self):
        outer = Frame(self, bg=FONDO)
        outer.pack(fill="both", expand=True, padx=16, pady=14)

        self._build_sidebar(outer)

        card = Frame(outer, bg=TARJETA, bd=0, highlightbackground="#E0DADA", highlightthickness=1)
        card.pack(side="left", fill="both", expand=True, padx=(0, 14))

        # Registro de proceso: fijo abajo, siempre visible (se empaqueta primero).
        log_wrap = Frame(card, bg=TARJETA)
        log_wrap.pack(side="bottom", fill="x", padx=20, pady=(0, 16))
        self._section_title(log_wrap, "📋 Registro de proceso")
        self.log = ScrolledText(
            log_wrap, height=9, bg="#2A2A2A", fg="#EDEDED",
            insertbackground=BLANCO, font=("Consolas", 9),
            relief="flat", padx=10, pady=8,
        )
        self.log.pack(fill="x", pady=(2, 0))
        self.log.configure(state="disabled")

        # Formulario: area con scroll, ocupa el resto del espacio disponible.
        scroll_area = Frame(card, bg=TARJETA)
        scroll_area.pack(side="top", fill="both", expand=True)
        self._build_scrollable_form(scroll_area)

    def _build_scrollable_form(self, parent):
        canvas = Canvas(parent, bg=TARJETA, highlightthickness=0)
        scrollbar = ttk.Scrollbar(parent, orient="vertical", command=canvas.yview)
        content = Frame(canvas, bg=TARJETA)

        content.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all")),
        )
        canvas_window = canvas.create_window((0, 0), window=content, anchor="nw")
        canvas.bind(
            "<Configure>",
            lambda e: canvas.itemconfigure(canvas_window, width=e.width),
        )
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.pack(side="left", fill="both", expand=True, padx=(20, 0), pady=18)
        scrollbar.pack(side="right", fill="y", pady=18)

        def _on_mousewheel(event):
            canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

        canvas.bind_all("<MouseWheel>", _on_mousewheel)

        self._build_form(content)

    # ------------------------------------------------------------------
    def _build_sidebar(self, parent):
        sidebar = Frame(parent, bg=PRIMARIO_OSCURO, width=230)
        sidebar.pack(side="right", fill="y")
        sidebar.pack_propagate(False)

        Label(
            sidebar, text="ACCIONES", bg=PRIMARIO_OSCURO, fg=BLANCO,
            font=("Segoe UI", 10, "bold"),
        ).pack(anchor="w", padx=18, pady=(20, 8))

        self.btn_procesar = Button(
            sidebar, text="1. Generar plantilla", command=self.on_procesar,
            bg=ACENTO, fg=BLANCO,
            activebackground=ACENTO_CLARO, activeforeground=BLANCO,
            font=("Segoe UI", 11, "bold"), relief="flat",
            padx=10, pady=12, cursor="hand2", anchor="w", justify="left",
            wraplength=190,
        )
        self.btn_procesar.pack(fill="x", padx=18, pady=(0, 6))

        self.status_var = StringVar(value="")
        Label(
            sidebar, textvariable=self.status_var, bg=PRIMARIO_OSCURO, fg=PRIMARIO_SUAVE,
            font=("Segoe UI", 9), wraplength=190, justify="left", anchor="w",
        ).pack(fill="x", padx=18, pady=(0, 18))

        separador = Frame(sidebar, bg=PRIMARIO, height=1)
        separador.pack(fill="x", padx=18, pady=(0, 18))

        self.btn_json = Button(
            sidebar, text="2. Convertir a JSON\n(SISPRO)", command=self.on_convertir_json,
            bg=PRIMARIO, fg=BLANCO,
            activebackground="#8C262B", activeforeground=BLANCO,
            font=("Segoe UI", 11, "bold"), relief="flat",
            padx=10, pady=12, cursor="hand2", anchor="w", justify="left",
            wraplength=190, state="disabled",
        )
        self.btn_json.pack(fill="x", padx=18, pady=(0, 6))

        self.status_json_var = StringVar(value="Genera la plantilla primero.")
        Label(
            sidebar, textvariable=self.status_json_var, bg=PRIMARIO_OSCURO, fg=PRIMARIO_SUAVE,
            font=("Segoe UI", 9), wraplength=190, justify="left", anchor="w",
        ).pack(fill="x", padx=18, pady=(0, 18))

        separador2 = Frame(sidebar, bg=PRIMARIO, height=1)
        separador2.pack(fill="x", padx=18, pady=(0, 18))

        self.btn_xml = Button(
            sidebar, text="3. Ajustar XML", command=self.on_ajustar_xml,
            bg=ACENTO, fg=BLANCO,
            activebackground=ACENTO_CLARO, activeforeground=BLANCO,
            font=("Segoe UI", 11, "bold"), relief="flat",
            padx=10, pady=12, cursor="hand2", anchor="w", justify="left",
            wraplength=190,
        )
        self.btn_xml.pack(fill="x", padx=18, pady=(0, 6))

        self.status_xml_var = StringVar(value="Sube el XML de la factura electronica.")
        Label(
            sidebar, textvariable=self.status_xml_var, bg=PRIMARIO_OSCURO, fg=PRIMARIO_SUAVE,
            font=("Segoe UI", 9), wraplength=190, justify="left", anchor="w",
        ).pack(fill="x", padx=18, pady=(0, 18))

    # ------------------------------------------------------------------
    def _build_form(self, content):
        self._card_group(
            content, "🧾  Datos de la factura",
            lambda body: self._seccion_facturacion(body),
        )

        fila_split = Frame(content, bg=FONDO)
        fila_split.pack(fill="x", padx=(0, 16), pady=(0, 4))
        col_izq = Frame(fila_split, bg=FONDO)
        col_izq.pack(side="left", fill="both", expand=True, padx=(0, 8))
        col_der = Frame(fila_split, bg=FONDO)
        col_der.pack(side="left", fill="both", expand=True, padx=(8, 0))

        self._card_group(
            col_izq, "📂  Archivos RIPS (TXT)",
            lambda body: self._seccion_archivos_rips(body),
            ultima=True, right_pad=0,
        )
        self._card_group(
            col_der, "💳  Pagos compartidos",
            lambda body: self._seccion_pago_y_salida(body),
            ultima=True, right_pad=0,
        )

    def _card_group(self, parent, titulo, builder, ultima=False, right_pad=16):
        card = Frame(parent, bg=PRIMARIO_SUAVE)
        card.pack(fill="x", padx=(0, right_pad), pady=(0, 14 if not ultima else 4))

        Label(
            card, text=titulo, bg=PRIMARIO_SUAVE, fg=PRIMARIO_OSCURO,
            font=("Segoe UI", 12, "bold"), anchor="w",
        ).pack(fill="x", padx=16, pady=(12, 6))

        body = Frame(card, bg=TARJETA)
        body.pack(fill="x", padx=12, pady=(0, 12))
        inner = Frame(body, bg=TARJETA)
        inner.pack(fill="x", padx=8, pady=10)
        builder(inner)

    # ------------------------------------------------------------------
    def _seccion_facturacion(self, body):
        LABEL_W = 13

        # Orden pedido: Fecha, Prefijo, Consecutivo, Empresa (Cliente), Tipo de usuario.
        fila1 = Frame(body, bg=TARJETA)
        fila1.pack(fill="x", pady=(2, 10))

        Label(
            fila1, text="Fecha", bg=TARJETA, fg="#333333",
            font=("Segoe UI", 9), width=8, anchor="w",
        ).pack(side="left")
        self.date_entry_factura = DateEntry(
            fila1, date_pattern="dd/mm/yyyy", font=("Segoe UI", 10), width=12,
            background=PRIMARIO, foreground=BLANCO, borderwidth=1,
        )
        self.date_entry_factura.pack(side="left", padx=(0, 24))

        Label(
            fila1, text="Prefijo", bg=TARJETA, fg="#333333",
            font=("Segoe UI", 9), width=8, anchor="w",
        ).pack(side="left")
        Entry(
            fila1, textvariable=self.prefijo_var, font=("Segoe UI", 10),
            width=10, relief="flat", highlightthickness=1,
            highlightbackground=PRIMARIO, highlightcolor=PRIMARIO,
        ).pack(side="left", padx=(0, 24))

        Label(
            fila1, text="Consecutivo", bg=TARJETA, fg="#333333",
            font=("Segoe UI", 9), width=12, anchor="w",
        ).pack(side="left")
        Entry(
            fila1, textvariable=self.consecutivo_var, font=("Segoe UI", 10),
            width=14, relief="flat", highlightthickness=1,
            highlightbackground=PRIMARIO, highlightcolor=PRIMARIO,
        ).pack(side="left")

        self.prefijo_var.trace_add("write", lambda *_: self._suggest_output())
        self.consecutivo_var.trace_add("write", lambda *_: self._suggest_output())

        fila2 = Frame(body, bg=TARJETA)
        fila2.pack(fill="x", pady=(2, 2))
        Label(
            fila2, text="Empresa (cliente)", bg=TARJETA, fg="#333333",
            font=("Segoe UI", 9), width=LABEL_W, anchor="w",
        ).pack(side="left")
        self.combo_empresa = ttk.Combobox(
            fila2, textvariable=self.empresa_var,
            values=NOMBRES_CLIENTES_ORDENADOS,
            font=("Segoe UI", 10), width=50,
        )
        self.combo_empresa.pack(side="left")
        self.combo_empresa.bind("<KeyRelease>", self._filtrar_empresas)
        self.combo_empresa.bind("<<ComboboxSelected>>", self._resolver_nit_cliente)
        self.combo_empresa.bind("<FocusOut>", self._resolver_nit_cliente)

        fila2_hint = Frame(body, bg=TARJETA)
        fila2_hint.pack(fill="x", pady=(2, 10))
        Label(
            fila2_hint, text="", bg=TARJETA, width=LABEL_W,
        ).pack(side="left")
        Label(
            fila2_hint, textvariable=self.nit_resuelto_var, bg=TARJETA, fg=GRIS_TEXTO,
            font=("Segoe UI", 8, "italic"), anchor="w",
        ).pack(side="left")
        Label(
            fila2_hint,
            text="  (si tu cliente no aparece, escribe directamente su NIT de 9 digitos)",
            bg=TARJETA, fg=GRIS_TEXTO, font=("Segoe UI", 8), anchor="w",
        ).pack(side="left")

        fila3 = Frame(body, bg=TARJETA)
        fila3.pack(fill="x", pady=(2, 2))
        Label(
            fila3, text="Tipo de usuario", bg=TARJETA, fg="#333333",
            font=("Segoe UI", 9), width=LABEL_W, anchor="w",
        ).pack(side="left")
        ttk.Combobox(
            fila3, textvariable=self.tipo_usuario_var,
            values=list(rips_core.TIPO_USUARIO_OPCIONES.keys()),
            state="readonly", font=("Segoe UI", 10), width=22,
        ).pack(side="left")

    def _seccion_archivos_rips(self, body):
        row_auto = Frame(body, bg=TARJETA)
        row_auto.pack(fill="x", pady=(2, 8))
        Button(
            row_auto, text="Cargar RIPS TXT",
            command=self._browse_todos_txt,
            bg=ACENTO, fg=BLANCO, activebackground=ACENTO_CLARO,
            activeforeground=BLANCO, relief="flat",
            font=("Segoe UI", 9, "bold"), padx=12, pady=6, cursor="hand2",
        ).pack(side="left")
        Label(
            row_auto,
            text="  selecciona AF0/US0/AP0/AD0/CT0 juntos: se identifican por el prefijo del nombre",
            bg=TARJETA, fg=GRIS_TEXTO, font=("Segoe UI", 8),
        ).pack(side="left")

        listas = Frame(body, bg=TARJETA)
        listas.pack(fill="x", pady=(4, 4))

        col_obligatorios = Frame(listas, bg=TARJETA)
        col_obligatorios.pack(side="left", anchor="n", padx=(0, 40))
        self._subtitle(col_obligatorios, "Obligatorios")
        self._build_txt_checklist(col_obligatorios, ["af0", "us0", "ap0", "ad0"])

        col_opcional = Frame(listas, bg=TARJETA)
        col_opcional.pack(side="left", anchor="n")
        self._subtitle(col_opcional, "Opcional (valida cantidades)")
        self._build_txt_checklist(col_opcional, ["ct0"])

    def _seccion_pago_y_salida(self, body):
        self._subtitle(body, "Copagos / cuotas moderadoras / pagos compartidos")
        row4 = Frame(body, bg=TARJETA)
        row4.pack(fill="x", pady=(2, 2))
        self._file_row(
            row4, "Reporte órdenes por empresa:", "pago",
            [("Excel", "*.xls *.xlsx"), ("Todos", "*.*")], optional=True, wide=True,
        )

        self._subtitle(body, "Factura electrónica (XML)")
        row5 = Frame(body, bg=TARJETA)
        row5.pack(fill="x", pady=(2, 2))
        self._file_row(
            row5, "XML de la factura:", "xml_factura",
            [("XML", "*.xml"), ("Todos", "*.*")], optional=True, wide=True,
        )

    # ------------------------------------------------------------------
    def _section_title(self, parent, text):
        Label(
            parent, text=text, bg=parent["bg"], fg=PRIMARIO_OSCURO,
            font=("Segoe UI", 10, "bold"),
        ).pack(fill="x", pady=(0, 4))

    def _subtitle(self, parent, text):
        Label(
            parent, text=text, bg=TARJETA, fg=TEXTO,
            font=("Segoe UI", 9, "bold"),
        ).pack(fill="x", pady=(4, 2))

    def _file_row(self, parent, label, key, filetypes, optional=False, wide=False):
        row = Frame(parent, bg=TARJETA)
        row.pack(fill="x", pady=3)

        tag = " (opcional)" if optional else ""
        Label(
            row, text=f"{label}{tag}", bg=TARJETA, fg="#333333",
            font=("Segoe UI", 9), width=26 if not wide else 30, anchor="w",
        ).pack(side="left")

        entry = Label(
            row, textvariable=self.paths[key], bg=GRIS_CLARO, fg="#222222",
            font=("Segoe UI", 9), anchor="w", padx=8, pady=5,
            relief="flat", bd=1,
        )
        entry.pack(side="left", fill="x", expand=True, padx=8)

        Button(
            row, text="Examinar...", command=lambda: self._browse(key, filetypes),
            bg=PRIMARIO, fg=BLANCO, activebackground=PRIMARIO_OSCURO,
            activeforeground=BLANCO, relief="flat", font=("Segoe UI", 9),
            padx=10, cursor="hand2",
        ).pack(side="left")

        if optional:
            Button(
                row, text="Quitar", command=lambda: self.paths[key].set(""),
                bg=GRIS_CLARO, fg=GRIS_TEXTO, relief="flat",
                font=("Segoe UI", 9), padx=8, cursor="hand2",
            ).pack(side="left", padx=(6, 0))

    TXT_LABELS = {"af0": "AF0", "us0": "US0", "ap0": "AP0", "ad0": "AD0", "ct0": "CT0"}

    def _build_txt_checklist(self, parent, keys):
        for key in keys:
            row = Frame(parent, bg=TARJETA)
            row.pack(fill="x", pady=2, anchor="w")
            Label(
                row, text=self.TXT_LABELS[key], bg=TARJETA, fg="#333333",
                font=("Segoe UI", 9, "bold"), width=6, anchor="w",
            ).pack(side="left")
            var = self.txt_status_vars.setdefault(key, StringVar())
            Label(
                row, textvariable=var, bg=TARJETA, fg=GRIS_TEXTO,
                font=("Segoe UI", 9), anchor="w",
            ).pack(side="left", padx=(6, 0))
            self._refresh_txt_status(key)

    def _refresh_txt_status(self, key):
        var = self.txt_status_vars.get(key)
        if var is None:
            return
        path = self.paths[key].get()
        var.set(f"✓ {os.path.basename(path)}" if path else "no cargado")

    def _numero_factura(self):
        return f"{self.prefijo_var.get().strip()}{self.consecutivo_var.get().strip()}"

    def _filtrar_empresas(self, event=None):
        if event is not None and event.keysym in ("Up", "Down", "Return", "Escape"):
            return
        texto = self.empresa_var.get().strip().upper()
        if not texto:
            self.combo_empresa["values"] = NOMBRES_CLIENTES_ORDENADOS
        else:
            self.combo_empresa["values"] = [
                n for n in NOMBRES_CLIENTES_ORDENADOS if texto in n.upper()
            ]
        self._resolver_nit_cliente()

    def _resolver_nit_cliente(self, event=None):
        texto = self.empresa_var.get().strip()
        if texto in NOMBRES_A_NIT:
            self.nit_cliente_var.set(NOMBRES_A_NIT[texto])
            self.nit_resuelto_var.set(f"NIT: {NOMBRES_A_NIT[texto]}")
        elif texto.isdigit() and len(texto) == 9:
            self.nit_cliente_var.set(texto)
            nombre = NITS_CLIENTES.get(texto)
            self.nit_resuelto_var.set(f"NIT: {texto}" + (f" ({nombre})" if nombre else " (no esta en el listado)"))
        else:
            self.nit_cliente_var.set("")
            self.nit_resuelto_var.set("Selecciona una empresa de la lista o escribe el NIT (9 digitos)")

    # ------------------------------------------------------------------
    def _build_footer(self):
        footer = Frame(self, bg=PRIMARIO_OSCURO, height=28)
        footer.pack(fill="x", side="bottom")
        footer.pack_propagate(False)
        Label(
            footer, text="Creado por: Christian Duque - Contador Público - 3134380447",
            bg=PRIMARIO_OSCURO, fg=ACENTO_CLARO, font=("Segoe UI", 8),
        ).pack(side="right", padx=14)

    # ------------------------------------------------------------------
    def _browse(self, key, filetypes):
        path = filedialog.askopenfilename(title="Selecciona el archivo", filetypes=filetypes)
        if path:
            self.paths[key].set(path)
            if key == "af0":
                self._suggest_output()

    PREFIJOS_TXT = {"AF": "af0", "US": "us0", "AP": "ap0", "AD": "ad0", "CT": "ct0"}

    def _browse_todos_txt(self):
        rutas = filedialog.askopenfilenames(
            title="Selecciona todos los TXT (AF0, US0, AP0, AD0, CT0)",
            filetypes=[("Texto", "*.txt"), ("Todos", "*.*")],
        )
        if not rutas:
            return

        asignados = {}
        duplicados = []
        no_reconocidos = []
        for ruta in rutas:
            nombre = os.path.basename(ruta).upper()
            prefijo = next((p for p in self.PREFIJOS_TXT if nombre.startswith(p)), None)
            if prefijo is None:
                no_reconocidos.append(os.path.basename(ruta))
                continue
            key = self.PREFIJOS_TXT[prefijo]
            if key in asignados:
                duplicados.append(os.path.basename(ruta))
                continue
            asignados[key] = ruta

        for key, ruta in asignados.items():
            self.paths[key].set(ruta)
            self._refresh_txt_status(key)

        if "af0" in asignados:
            self._suggest_output()

        resumen = "\n".join(f"- {k.upper()}: {os.path.basename(v)}" for k, v in asignados.items())
        mensaje = f"Archivos asignados:\n{resumen}" if resumen else "No se reconocio ningun archivo."
        if duplicados:
            mensaje += "\n\nIgnorados (prefijo repetido): " + ", ".join(duplicados)
        if no_reconocidos:
            mensaje += "\n\nNo reconocidos (no empiezan por AF/US/AP/AD/CT): " + ", ".join(no_reconocidos)
        messagebox.showinfo("Carga automatica de TXT", mensaje)

    def _suggest_output(self):
        numero = self._numero_factura()
        af0_path = self.paths["af0"].get()
        if not numero or not af0_path:
            return
        try:
            carpeta = os.path.dirname(af0_path)
            nombre = f"RIPS_{numero}.xlsm"
            self.paths["output"].set(os.path.join(carpeta, nombre))
        except Exception:
            pass


    # ------------------------------------------------------------------
    def _log(self, text):
        self.log.configure(state="normal")
        self.log.insert(END, text + "\n")
        self.log.configure(state="disabled")
        self.log.see(END)

    def _clear_log(self):
        self.log.configure(state="normal")
        self.log.delete("1.0", END)
        self.log.configure(state="disabled")

    # ------------------------------------------------------------------
    # 1. Generar plantilla
    # ------------------------------------------------------------------
    def on_procesar(self):
        if not os.path.exists(TEMPLATE_PATH):
            messagebox.showerror(
                "Falta la plantilla del programa",
                f"No se encontro la plantilla oficial en:\n{TEMPLATE_PATH}",
            )
            return
        required = ["af0", "us0", "ap0", "ad0"]
        faltantes = [k for k in required if not self.paths[k].get().strip()]
        if faltantes:
            messagebox.showerror(
                "Faltan archivos",
                "Debes cargar: " + ", ".join(faltantes),
            )
            return
        if not self.prefijo_var.get().strip() or not self.consecutivo_var.get().strip():
            messagebox.showerror(
                "Falta el numero de factura",
                "Debes ingresar el prefijo y el consecutivo de la factura (ej. FEDI / 38841).",
            )
            return

        self._suggest_output()
        if not self.paths["output"].get().strip():
            messagebox.showerror(
                "Error interno",
                "No se pudo determinar la ruta de salida de la plantilla.",
            )
            return

        self.btn_procesar.configure(state="disabled")
        self.btn_json.configure(state="disabled")
        self.status_var.set("Procesando...")
        self.status_json_var.set("Genera la plantilla primero.")
        self._clear_log()
        self._log("Iniciando procesamiento...")

        thread = threading.Thread(target=self._run_proceso, daemon=True)
        thread.start()

    def _run_proceso(self):
        try:
            resumen, warnings = rips_core.procesar(
                af0_path=self.paths["af0"].get(),
                us0_path=self.paths["us0"].get(),
                ap0_path=self.paths["ap0"].get(),
                ad0_path=self.paths["ad0"].get(),
                template_path=TEMPLATE_PATH,
                output_path=self.paths["output"].get(),
                num_factura=self._numero_factura(),
                tipo_usuario=rips_core.TIPO_USUARIO_OPCIONES[self.tipo_usuario_var.get()],
                ct0_path=self.paths["ct0"].get() or None,
                pago_path=self.paths["pago"].get() or None,
                today=date.today(),
            )
        except rips_core.RipsError as exc:
            self.root.after(0, self._on_error, str(exc))
            return
        except Exception:
            self.root.after(0, self._on_error, traceback.format_exc())
            return

        self.root.after(0, self._on_success, resumen, warnings)

    def _on_error(self, mensaje):
        if self._destroyed:
            return
        self._log("ERROR: " + mensaje)
        self.status_var.set("Error")
        self.btn_procesar.configure(state="normal")
        messagebox.showerror("Error al generar la plantilla", mensaje)

    def _on_success(self, resumen, warnings):
        if self._destroyed:
            return
        self._log(f"Factura: {resumen['numFactura']}")
        self._log(f"NIT obligado: {resumen['numDocumentoIdObligado']}")
        self._log(f"Usuarios cargados: {resumen['totalUsuarios']}")
        self._log(f"Registros de consultas cargados: {resumen['totalConsultas']}")
        self._log(f"Suma vrServicio: {resumen['sumaVrServicio']}")
        self._log(f"Valor total AD0: {resumen['valorTotalAD0']}")
        self._log(
            "Cuadre AD0 vs vrServicio: "
            + ("OK" if resumen["cuadreOk"] else "NO CUADRA - revisar")
        )
        if warnings:
            self._log("")
            self._log("Advertencias:")
            for w in warnings:
                self._log(" - " + w)
        self._log("")
        self._log(f"Plantilla generada en: {resumen['outputPath']}")

        self.status_var.set("Listo" if resumen["cuadreOk"] else "Listo, con advertencias")
        self.btn_procesar.configure(state="normal")

        self.last_output_path = resumen["outputPath"]
        self.btn_json.configure(state="normal")
        self.status_json_var.set("Listo para convertir.")

        messagebox.showinfo("Listo", "Plantilla generada con exito.")

    # ------------------------------------------------------------------
    # 2. Convertir a JSON en SISPRO
    # ------------------------------------------------------------------
    def on_convertir_json(self):
        if not self.last_output_path or not os.path.exists(self.last_output_path):
            messagebox.showerror(
                "No hay plantilla generada",
                "Primero genera la plantilla con el boton 1.",
            )
            return

        self.btn_json.configure(state="disabled")
        self.status_json_var.set("Subiendo a SISPRO...")
        self._log("")
        self._log("Subiendo la plantilla al convertidor oficial de SISPRO...")

        thread = threading.Thread(target=self._run_convertir_json, daemon=True)
        thread.start()

    def _run_convertir_json(self):
        try:
            output_json = rips_core.convertir_json_sispro(self.last_output_path)
        except rips_core.RipsError as exc:
            self.root.after(0, self._on_json_error, str(exc))
            return
        except Exception:
            self.root.after(0, self._on_json_error, traceback.format_exc())
            return

        self.root.after(0, self._on_json_success, output_json)

    def _on_json_error(self, mensaje):
        if self._destroyed:
            return
        self._log("ERROR (SISPRO): " + mensaje)
        self.status_json_var.set("Error al convertir.")
        self.btn_json.configure(state="normal")
        messagebox.showerror("Error al convertir a JSON", mensaje)

    def _on_json_success(self, output_json):
        if self._destroyed:
            return
        self.last_json_path = output_json
        self._log(f"JSON generado por SISPRO guardado en: {output_json}")
        self.status_json_var.set("JSON listo.")
        self.btn_json.configure(state="normal")
        messagebox.showinfo(
            "JSON generado",
            f"El JSON se genero y se guardo en:\n{output_json}",
        )

    # ------------------------------------------------------------------
    # 3. Ajustar XML de la factura electronica
    # ------------------------------------------------------------------
    def on_ajustar_xml(self):
        xml_path = self.paths["xml_factura"].get().strip()
        if not xml_path or not os.path.exists(xml_path):
            messagebox.showerror(
                "Falta el XML",
                "Debes cargar el XML de la factura electronica.",
            )
            return

        prefijo = self.prefijo_var.get().strip()
        consecutivo = self.consecutivo_var.get().strip()
        if not prefijo or not consecutivo:
            messagebox.showerror(
                "Falta el numero de factura",
                "Debes ingresar el prefijo y el consecutivo de la factura.",
            )
            return

        self.btn_xml.configure(state="disabled")
        self.status_xml_var.set("Ajustando XML...")
        self._log("")
        self._log("Ajustando el XML de la factura electronica...")

        thread = threading.Thread(
            target=self._run_ajustar_xml, args=(xml_path, prefijo, consecutivo), daemon=True,
        )
        thread.start()

    def _run_ajustar_xml(self, xml_path, prefijo, consecutivo):
        try:
            tipo_usuario = rips_core.TIPO_USUARIO_OPCIONES[self.tipo_usuario_var.get()]

            base_dir = None
            salida = self.paths["output"].get().strip()
            if salida:
                base_dir = os.path.dirname(salida)
            if not base_dir:
                base_dir = os.path.dirname(xml_path)

            carpeta = rips_core.carpeta_salida_rips(base_dir, prefijo, consecutivo)

            nombre_xml = os.path.splitext(os.path.basename(xml_path))[0] + "_ajustado.xml"
            xml_destino = os.path.join(carpeta, nombre_xml)
            rips_core.ajustar_xml_factura(xml_path, tipo_usuario, xml_destino)

            json_copiado = None
            if self.last_json_path and os.path.exists(self.last_json_path):
                json_destino = os.path.join(carpeta, os.path.basename(self.last_json_path))
                if os.path.abspath(json_destino) != os.path.abspath(self.last_json_path):
                    if os.path.exists(json_destino):
                        os.remove(json_destino)
                    shutil.move(self.last_json_path, json_destino)
                    self.last_json_path = json_destino
                json_copiado = json_destino
        except rips_core.RipsError as exc:
            self.root.after(0, self._on_xml_error, str(exc))
            return
        except Exception:
            self.root.after(0, self._on_xml_error, traceback.format_exc())
            return

        self.root.after(0, self._on_xml_success, carpeta, xml_destino, json_copiado)

    def _on_xml_error(self, mensaje):
        if self._destroyed:
            return
        self._log("ERROR (XML): " + mensaje)
        self.status_xml_var.set("Error al ajustar el XML.")
        self.btn_xml.configure(state="normal")
        messagebox.showerror("Error al ajustar el XML", mensaje)

    def _on_xml_success(self, carpeta, xml_destino, json_copiado):
        if self._destroyed:
            return
        self._log(f"XML ajustado guardado en: {xml_destino}")
        if json_copiado:
            self._log(f"JSON movido a: {json_copiado}")
        else:
            self._log("Nota: aun no se ha generado el JSON (boton 2); solo se guardo el XML ajustado.")
        self.status_xml_var.set("XML ajustado.")
        self.btn_xml.configure(state="normal")
        messagebox.showinfo(
            "XML ajustado",
            f"El XML se ajusto y se guardo en la carpeta:\n{carpeta}",
        )


# ---------------------------------------------------------------------------
# Menu principal: elegir entre convertir una plantilla ya diligenciada a
# JSON, o crear la plantilla desde los archivos planos (interfaz actual).
# ---------------------------------------------------------------------------

class MainMenuFrame(Frame):
    def __init__(self, root, on_json, on_rips):
        super().__init__(root, bg=FONDO)
        self.root = root

        header = Frame(self, bg=PRIMARIO, height=84)
        header.pack(fill="x", side="top")
        header.pack_propagate(False)
        Label(
            header, text="🧾 GENERADOR DE RIPS ELECTRONICOS Y XML - LABORATORIO ESPECIALIZADO MICRODIAG",
            bg=PRIMARIO, fg=BLANCO, font=("Segoe UI", 14, "bold"),
        ).pack(side="left", padx=24, pady=10)
        Frame(self, bg=ACENTO, height=4).pack(fill="x")

        footer = Frame(self, bg=PRIMARIO_OSCURO, height=28)
        footer.pack(fill="x", side="bottom")
        footer.pack_propagate(False)
        Label(
            footer, text="Creado por: Christian Duque - Contador Público - 3134380447",
            bg=PRIMARIO_OSCURO, fg=ACENTO_CLARO, font=("Segoe UI", 8),
        ).pack(side="right", padx=14)

        center = Frame(self, bg=FONDO)
        center.pack(fill="both", expand=True)

        Label(
            center, text="¿Qué deseas hacer?", bg=FONDO, fg=TEXTO,
            font=("Segoe UI", 17, "bold"),
        ).pack(pady=(70, 32))

        botones = Frame(center, bg=FONDO)
        botones.pack()

        self._tarjeta_menu(
            botones,
            "🔄  Convertir a JSON\nla plantilla RIPS MINSALUD",
            "Descarga la plantilla oficial de SISPRO, carga la que ya tengas diligenciada "
            "y conviértela a JSON con el conversor oficial.",
            on_json,
        ).pack(side="left", padx=20)

        self._tarjeta_menu(
            botones,
            "🧾  Crear plantilla RIPS MINSALUD\ndesde archivos planos",
            "Genera la plantilla oficial a partir de los archivos AF0/US0/AP0/AD0/CT0, "
            "y opcionalmente conviértela a JSON y ajusta el XML de la factura.",
            on_rips,
        ).pack(side="left", padx=20)

    def _tarjeta_menu(self, parent, titulo, descripcion, comando):
        card = Frame(
            parent, bg=TARJETA, bd=0, highlightbackground="#E0DADA", highlightthickness=1,
            width=340, height=230,
        )
        card.pack_propagate(False)

        Label(
            card, text=titulo, bg=TARJETA, fg=PRIMARIO_OSCURO,
            font=("Segoe UI", 13, "bold"), justify="center",
        ).pack(pady=(30, 10), padx=16)
        Label(
            card, text=descripcion, bg=TARJETA, fg=GRIS_TEXTO, font=("Segoe UI", 9),
            wraplength=280, justify="center",
        ).pack(padx=20, pady=(0, 20))
        Button(
            card, text="Seleccionar", command=comando,
            bg=PRIMARIO, fg=BLANCO, activebackground=PRIMARIO_OSCURO, activeforeground=BLANCO,
            relief="flat", font=("Segoe UI", 10, "bold"), padx=22, pady=9, cursor="hand2",
        ).pack(pady=(0, 20))
        return card


# ---------------------------------------------------------------------------
# Sub-interfaz: convertir una plantilla RIPS ya diligenciada a JSON (SISPRO),
# sin pasar por la generacion desde archivos planos.
# ---------------------------------------------------------------------------

class JsonOnlyFrame(Frame):
    def __init__(self, root, on_volver):
        super().__init__(root, bg=FONDO)
        self.root = root
        self.on_volver = on_volver
        self._destroyed = False

        self.plantilla_path = StringVar()

        self._build_header()
        self._build_footer()
        self._build_body()

    def on_navigate_away(self):
        self._destroyed = True

    # ------------------------------------------------------------------
    def _build_header(self):
        header = Frame(self, bg=PRIMARIO, height=84)
        header.pack(fill="x", side="top")
        header.pack_propagate(False)

        Button(
            header, text="◀ Menú principal", command=self.on_volver,
            bg=PRIMARIO_OSCURO, fg=BLANCO, activebackground="#8C262B", activeforeground=BLANCO,
            relief="flat", font=("Segoe UI", 9, "bold"), padx=12, pady=6, cursor="hand2",
        ).pack(side="left", padx=(24, 12), pady=10)

        Label(
            header, text="🔄 CONVERTIR PLANTILLA RIPS A JSON (SISPRO)",
            bg=PRIMARIO, fg=BLANCO, font=("Segoe UI", 14, "bold"),
        ).pack(side="left", pady=10)

        Frame(self, bg=ACENTO, height=4).pack(fill="x")

    def _build_footer(self):
        footer = Frame(self, bg=PRIMARIO_OSCURO, height=28)
        footer.pack(fill="x", side="bottom")
        footer.pack_propagate(False)
        Label(
            footer, text="Creado por: Christian Duque - Contador Público - 3134380447",
            bg=PRIMARIO_OSCURO, fg=ACENTO_CLARO, font=("Segoe UI", 8),
        ).pack(side="right", padx=14)

    def _build_body(self):
        outer = Frame(self, bg=FONDO)
        outer.pack(fill="both", expand=True, padx=24, pady=20)

        card = Frame(outer, bg=TARJETA, bd=0, highlightbackground="#E0DADA", highlightthickness=1)
        card.pack(fill="both", expand=True)

        inner = Frame(card, bg=TARJETA)
        inner.pack(fill="x", padx=30, pady=24)

        Label(
            inner, text="1. Descarga la plantilla oficial (si aún no la tienes)",
            bg=TARJETA, fg=PRIMARIO_OSCURO, font=("Segoe UI", 11, "bold"), anchor="w",
        ).pack(fill="x", pady=(0, 6))
        Label(
            inner,
            text="Descarga directamente el archivo oficial de la plantilla RIPS MINSALUD vigente desde SISPRO.",
            bg=TARJETA, fg=GRIS_TEXTO, font=("Segoe UI", 9), anchor="w", justify="left", wraplength=760,
        ).pack(fill="x", pady=(0, 10))
        self.btn_descargar = Button(
            inner, text="⬇️  Descargar plantilla desde SISPRO", command=self._on_descargar_plantilla,
            bg=ACENTO, fg=BLANCO, activebackground=ACENTO_CLARO, activeforeground=BLANCO,
            relief="flat", font=("Segoe UI", 10, "bold"), padx=14, pady=8, cursor="hand2",
        )
        self.btn_descargar.pack(anchor="w", pady=(0, 24))

        Label(
            inner, text="2. Carga la plantilla ya diligenciada", bg=TARJETA, fg=PRIMARIO_OSCURO,
            font=("Segoe UI", 11, "bold"), anchor="w",
        ).pack(fill="x", pady=(0, 6))

        fila = Frame(inner, bg=TARJETA)
        fila.pack(fill="x", pady=(0, 4))
        entry = Label(
            fila, textvariable=self.plantilla_path, bg=GRIS_CLARO, fg="#222222",
            font=("Segoe UI", 9), anchor="w", padx=8, pady=6, relief="flat", bd=1,
        )
        entry.pack(side="left", fill="x", expand=True, padx=(0, 8))
        Button(
            fila, text="Examinar...", command=self._elegir_plantilla,
            bg=PRIMARIO, fg=BLANCO, activebackground=PRIMARIO_OSCURO, activeforeground=BLANCO,
            relief="flat", font=("Segoe UI", 9), padx=10, cursor="hand2",
        ).pack(side="left")

        Label(
            inner, text="3. Convierte la plantilla a JSON", bg=TARJETA, fg=PRIMARIO_OSCURO,
            font=("Segoe UI", 11, "bold"), anchor="w",
        ).pack(fill="x", pady=(24, 6))

        self.btn_convertir = Button(
            inner, text="Convertir plantilla a JSON", command=self._on_convertir,
            bg=PRIMARIO, fg=BLANCO, activebackground="#8C262B", activeforeground=BLANCO,
            relief="flat", font=("Segoe UI", 11, "bold"), padx=16, pady=10, cursor="hand2",
        )
        self.btn_convertir.pack(anchor="w")

        self.status_var = StringVar(value="")
        Label(
            inner, textvariable=self.status_var, bg=TARJETA, fg=GRIS_TEXTO,
            font=("Segoe UI", 9), anchor="w",
        ).pack(fill="x", pady=(6, 20))

        log_wrap = Frame(card, bg=TARJETA)
        log_wrap.pack(side="bottom", fill="x", padx=30, pady=(0, 24))
        Label(
            log_wrap, text="📋 Registro de proceso", bg=TARJETA, fg=PRIMARIO_OSCURO,
            font=("Segoe UI", 10, "bold"),
        ).pack(fill="x", pady=(0, 4))
        self.log = ScrolledText(
            log_wrap, height=10, bg="#2A2A2A", fg="#EDEDED",
            insertbackground=BLANCO, font=("Consolas", 9), relief="flat", padx=10, pady=8,
        )
        self.log.pack(fill="x")
        self.log.configure(state="disabled")

    # ------------------------------------------------------------------
    def _log(self, text):
        self.log.configure(state="normal")
        self.log.insert(END, text + "\n")
        self.log.configure(state="disabled")
        self.log.see(END)

    def _on_descargar_plantilla(self):
        destino = filedialog.asksaveasfilename(
            title="Guardar plantilla RIPS MINSALUD",
            defaultextension=".xlsm",
            initialfile="plantilla_fev-rips_v_1_0.xlsm",
            filetypes=[("Excel habilitado para macros", "*.xlsm"), ("Todos", "*.*")],
        )
        if not destino:
            return

        self.btn_descargar.configure(state="disabled")
        self.status_var.set("Descargando plantilla desde SISPRO...")
        self._log("Descargando la plantilla oficial desde SISPRO...")

        thread = threading.Thread(target=self._run_descargar_plantilla, args=(destino,), daemon=True)
        thread.start()

    def _run_descargar_plantilla(self, destino):
        try:
            rips_core.descargar_plantilla_sispro(destino)
        except rips_core.RipsError as exc:
            self.root.after(0, self._on_descarga_error, str(exc))
            return
        except Exception:
            self.root.after(0, self._on_descarga_error, traceback.format_exc())
            return

        self.root.after(0, self._on_descarga_success, destino)

    def _on_descarga_error(self, mensaje):
        if self._destroyed:
            return
        self._log("ERROR (descarga): " + mensaje)
        self.status_var.set("Error al descargar la plantilla.")
        self.btn_descargar.configure(state="normal")
        messagebox.showerror("Error al descargar la plantilla", mensaje)

    def _on_descarga_success(self, destino):
        if self._destroyed:
            return
        self._log(f"Plantilla descargada en: {destino}")
        self.status_var.set("Plantilla descargada.")
        self.btn_descargar.configure(state="normal")
        self.plantilla_path.set(destino)
        messagebox.showinfo("Plantilla descargada", f"La plantilla se guardo en:\n{destino}")

    def _elegir_plantilla(self):
        path = filedialog.askopenfilename(
            title="Selecciona la plantilla RIPS diligenciada",
            filetypes=[("Excel", "*.xlsm *.xlsx"), ("Todos", "*.*")],
        )
        if path:
            self.plantilla_path.set(path)

    def _on_convertir(self):
        path = self.plantilla_path.get().strip()
        if not path or not os.path.exists(path):
            messagebox.showerror(
                "Falta la plantilla",
                "Debes cargar la plantilla RIPS que quieres convertir.",
            )
            return

        self.btn_convertir.configure(state="disabled")
        self.status_var.set("Subiendo a SISPRO...")
        self._log("Subiendo la plantilla al convertidor oficial de SISPRO...")

        thread = threading.Thread(target=self._run_convertir, args=(path,), daemon=True)
        thread.start()

    def _run_convertir(self, path):
        try:
            output_json = rips_core.convertir_json_sispro(path)
        except rips_core.RipsError as exc:
            self.root.after(0, self._on_error, str(exc))
            return
        except Exception:
            self.root.after(0, self._on_error, traceback.format_exc())
            return

        self.root.after(0, self._on_success, output_json)

    def _on_error(self, mensaje):
        if self._destroyed:
            return
        self._log("ERROR: " + mensaje)
        self.status_var.set("Error al convertir.")
        self.btn_convertir.configure(state="normal")
        messagebox.showerror("Error al convertir a JSON", mensaje)

    def _on_success(self, output_json):
        if self._destroyed:
            return
        self._log(f"JSON generado por SISPRO guardado en: {output_json}")
        self.status_var.set("JSON listo.")
        self.btn_convertir.configure(state="normal")
        messagebox.showinfo(
            "JSON generado",
            f"El JSON se genero y se guardo en:\n{output_json}",
        )


# ---------------------------------------------------------------------------
# Controlador: intercambia el frame visible (menu / json / generador) sin
# crear una ventana nueva, y limpia temporizadores/bindings al navegar.
# ---------------------------------------------------------------------------

class AppController:
    def __init__(self, root):
        self.root = root
        self.view = None
        self.show_menu()

    def _swap(self, factory):
        anterior = self.view
        self.view = factory()
        self.view.pack(fill="both", expand=True)
        if anterior is not None:
            on_navigate_away = getattr(anterior, "on_navigate_away", None)
            if callable(on_navigate_away):
                on_navigate_away()
            anterior.destroy()

    def show_menu(self):
        self._swap(lambda: MainMenuFrame(self.root, on_json=self.show_json, on_rips=self.show_rips))

    def show_json(self):
        self._swap(lambda: JsonOnlyFrame(self.root, on_volver=self.show_menu))

    def show_rips(self):
        self._swap(lambda: RipsApp(self.root, on_volver=self.show_menu))


def main():
    root = Tk()
    # No se usa root.withdraw() aqui: en Windows, un Toplevel modal
    # (transient + grab_set) sobre una ventana raiz retirada (withdraw)
    # a veces no llega a mostrarse en pantalla (queda como proceso vivo
    # pero invisible, sin icono en la barra de tareas). Por eso la raiz
    # se deja visible (pequena) mientras se resuelve la licencia.
    root.geometry("1x1+0+0")
    root.title("GENERADOR DE RIPS ELECTRONICOS Y XML - LABORATORIO ESPECIALIZADO MICRODIAG")
    try:
        root.iconbitmap(ICON_PATH)
    except Exception:
        pass
    root.update()

    if datetime.now() > TRIAL_DEADLINE and not _licencia_activa():
        activada = _pedir_licencia(root)
        if not activada:
            try:
                root.destroy()
            except Exception:
                pass
            return

    root.configure(bg=FONDO)
    root.geometry("1280x860")
    root.minsize(1180, 760)
    try:
        root.state("zoomed")
    except Exception:
        root.attributes("-zoomed", True)

    AppController(root)

    # Windows a veces no le da el foco a una ventana nueva lanzada desde una
    # consola/otro proceso: el programa abre bien, pero queda detras de la
    # terminal u otras ventanas sin que el usuario lo note. Se fuerza a que
    # pase al frente al iniciar.
    root.lift()
    root.attributes("-topmost", True)
    root.after(200, lambda: root.attributes("-topmost", False))
    root.focus_force()

    root.mainloop()


if __name__ == "__main__":
    main()
