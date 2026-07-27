#!/usr/bin/env python3
"""
comparar_docx_gui.py
---------------------
Interfaz gráfica para comparar hasta 7 documentos de Word (.docx) y generar
un reporte .docx con los párrafos que difieren entre ellos.

REQUISITOS:
    pip install python-docx
    pip install tkinterdnd2      (opcional, habilita arrastrar y soltar)

EJECUTAR:
    python comparar_docx_gui.py

Si "tkinterdnd2" no está instalado la app funciona igual, solo sin
arrastrar y soltar: se usa el botón "Agregar archivos...".
"""

import os
import sys
import queue
import threading
import subprocess
from pathlib import Path
from difflib import SequenceMatcher
from lxml import etree

import tkinter as tk
from tkinter import ttk, filedialog, messagebox

from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn

try:
    from tkinterdnd2 import TkinterDnD, DND_FILES
    DND_DISPONIBLE = True
except ImportError:
    DND_DISPONIBLE = False

MAX_DOCS = 7
PALABRAS_POR_PAGINA_RESPALDO = 400

# Mensajes especiales que viajan por la cola
_MSG_LOG    = "LOG"
_MSG_OK     = "OK"
_MSG_ERROR  = "ERROR"


# ==========================================================================
# LÓGICA DE COMPARACIÓN
# ==========================================================================

def _reparar_docx(ruta_docx):
    """
    Devuelve un BytesIO con el .docx limpio de referencias a 'NULL'.

    Algunos documentos contienen:
      a) Entradas ZIP con nombre 'NULL' o que incluyen 'NULL' en su ruta.
      b) Referencias a esas entradas en los archivos .rels (relaciones XML).

    python-docx falla con "There is no item named 'word/NULL' in the archive"
    cuando intenta cargar un recurso referenciado desde las relaciones pero
    cuyo archivo no existe o tiene nombre nulo.  Esta función:
      1. Omite del ZIP las entradas con nombre NULL.
      2. Elimina de todos los archivos .rels las líneas <Relationship> que
         apunten a 'NULL' (Target="NULL" o Target que termina en /NULL).
    """
    import zipfile, io, re

    # Patrón que casa con <Relationship ... Target="...NULL..."/>
    RE_REL_NULL = re.compile(
        rb'<Relationship\b[^>]*\bTarget=["\'][^"\']*NULL[^"\']*["\'][^>]*/?>',
        re.IGNORECASE,
    )

    buf = io.BytesIO()
    with zipfile.ZipFile(ruta_docx, 'r') as zin, \
         zipfile.ZipFile(buf, 'w', zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            nombre = item.filename
            # 1. Omitir entradas cuyo nombre sea NULL o contenga /NULL
            if not nombre:
                continue
            partes = nombre.replace('\\', '/').split('/')
            if any(p.upper() == 'NULL' for p in partes):
                continue
            try:
                datos = zin.read(nombre)
            except Exception:
                continue   # entrada ilegible → se omite

            # 2. En archivos de relaciones, borrar las referencias a NULL
            if nombre.endswith('.rels'):
                datos = RE_REL_NULL.sub(b'', datos)

            zout.writestr(item, datos)
    buf.seek(0)
    return buf


def extraer_parrafos(ruta_docx):
    # Intentar abrir directo; si falla por entradas NULL corruptas, reparar primero
    try:
        doc = Document(ruta_docx)
    except Exception as e:
        msg = str(e).lower()
        if "null" in msg or "no item named" in msg or "there is no item" in msg:
            doc = Document(_reparar_docx(ruta_docx))
        else:
            raise

    parrafos = []
    pagina_actual = 1
    tiene_marcas_reales = False
    num_parrafo = 0

    for p in doc.paragraphs:
        p_xml = p._p
        saltos  = p_xml.findall('.//' + qn('w:br') + "[@" + qn('w:type') + "='page']")
        renders = p_xml.findall('.//' + qn('w:lastRenderedPageBreak'))

        if renders:
            tiene_marcas_reales = True
            pagina_actual += len(renders)
        if saltos:
            tiene_marcas_reales = True
            pagina_actual += len(saltos)

        texto = p.text.strip()
        if texto:
            num_parrafo += 1
            parrafos.append({"num": num_parrafo, "pagina": pagina_actual, "texto": texto})

    if not tiene_marcas_reales:
        acumulado = 0
        for item in parrafos:
            acumulado += len(item["texto"].split())
            item["pagina"] = max(1, (acumulado - 1) // PALABRAS_POR_PAGINA_RESPALDO + 1)

    return parrafos, tiene_marcas_reales


def calcular_diferencias(lista_parrafos_por_doc):
    n = len(lista_parrafos_por_doc)
    textos_base = [item["texto"] for item in lista_parrafos_por_doc[0]]

    registros = []
    for doc_idx in range(1, n):
        textos_otro = [item["texto"] for item in lista_parrafos_por_doc[doc_idx]]
        sm = SequenceMatcher(None, textos_base, textos_otro, autojunk=False)
        for tag, i1, i2, j1, j2 in sm.get_opcodes():
            if tag != "equal":
                registros.append((i1, i2, doc_idx, j1, j2))

    if not registros:
        return []

    rangos_base = sorted(set((r[0], r[1]) for r in registros))
    bloques_base = []
    cur_ini, cur_fin = rangos_base[0]
    for ini, fin in rangos_base[1:]:
        if ini <= cur_fin:
            cur_fin = max(cur_fin, fin)
        else:
            bloques_base.append((cur_ini, cur_fin))
            cur_ini, cur_fin = ini, fin
    bloques_base.append((cur_ini, cur_fin))

    resultado = []
    for b_ini, b_fin in bloques_base:
        filas = []
        if b_ini < b_fin:
            for k in range(b_ini, b_fin):
                filas.append((0, lista_parrafos_por_doc[0][k]))
        else:
            filas.append((0, None))

        for doc_idx in range(1, n):
            sub  = [r for r in registros if r[2] == doc_idx and r[0] < b_fin and r[1] > b_ini]
            sub += [r for r in registros if r[2] == doc_idx and r[0] == r[1] and b_ini <= r[0] <= b_fin]
            if not sub:
                continue
            j1 = min(r[3] for r in sub)
            j2 = max(r[4] for r in sub)
            if j1 < j2:
                for k in range(j1, j2):
                    filas.append((doc_idx, lista_parrafos_por_doc[doc_idx][k]))
            else:
                filas.append((doc_idx, None))

        resultado.append(filas)

    return resultado


def _sombrear_celda(celda, color_hex):
    tcPr = celda._tc.get_or_add_tcPr()
    shd = tcPr.makeelement(qn('w:shd'), {
        qn('w:val'): 'clear', qn('w:color'): 'auto', qn('w:fill'): color_hex,
    })
    tcPr.append(shd)


def _highlight_run(run, color="yellow"):
    """
    Aplica resaltado de color (highlight) a un run usando el elemento OOXML
    w:highlight.  'color' debe ser un valor válido de OOXML: 'yellow', 'cyan',
    'green', 'magenta', 'blue', 'red', 'darkBlue', etc.
    """
    rPr = run._r.get_or_add_rPr()
    # Elimina cualquier highlight previo para evitar duplicados
    for existing in rPr.findall(qn('w:highlight')):
        rPr.remove(existing)
    hl = etree.SubElement(rPr, qn('w:highlight'))
    hl.set(qn('w:val'), color)


def diff_palabras(texto_base, texto_otro):
    """
    Compara dos cadenas palabra a palabra con SequenceMatcher.

    Devuelve una lista de tuplas (palabra, es_diferente) para cada texto:
        - tokens_base : lista de (palabra_o_espacio, es_diferente)
        - tokens_otro : lista de (palabra_o_espacio, es_diferente)

    Las palabras marcadas como es_diferente=True son las que no están
    presentes (o cambiaron) respecto al otro texto.
    """
    import re

    def tokenizar(texto):
        """Separa el texto en palabras y espacios, preservando separadores."""
        return re.findall(r'\S+|\s+', texto) if texto else []

    palabras_base = tokenizar(texto_base or "")
    palabras_otro = tokenizar(texto_otro or "")

    sm = SequenceMatcher(None, palabras_base, palabras_otro, autojunk=False)

    tokens_base = []
    tokens_otro = []

    for tag, i1, i2, j1, j2 in sm.get_opcodes():
        if tag == 'equal':
            for t in palabras_base[i1:i2]:
                tokens_base.append((t, False))
            for t in palabras_otro[j1:j2]:
                tokens_otro.append((t, False))
        elif tag == 'replace':
            for t in palabras_base[i1:i2]:
                tokens_base.append((t, True))
            for t in palabras_otro[j1:j2]:
                tokens_otro.append((t, True))
        elif tag == 'delete':
            for t in palabras_base[i1:i2]:
                tokens_base.append((t, True))
        elif tag == 'insert':
            for t in palabras_otro[j1:j2]:
                tokens_otro.append((t, True))

    return tokens_base, tokens_otro


def _escribir_celda_con_diff(parrafo, tokens):
    """
    Escribe los tokens (lista de (texto, es_diferente)) en el párrafo
    de una celda, usando runs separados.  Los tokens marcados como
    diferentes reciben fondo amarillo (highlight).
    """
    for texto, diferente in tokens:
        if not texto:  # saltar tokens vacíos
            continue
        run = parrafo.add_run(texto)
        if diferente:
            _highlight_run(run, "yellow")


def construir_reporte(rutas, bloques, marcas_reales_por_doc, ruta_salida):
    out = Document()
    style = out.styles["Normal"]
    style.font.name = "Calibri"
    style.font.size = Pt(10.5)

    out.add_heading("Comparación de documentos Word", level=1)
    out.add_paragraph(
        f"Se compararon {len(rutas)} documentos, usando "
        f"«{Path(rutas[0]).name}» como documento base. "
        "La tabla siguiente muestra únicamente los párrafos que difieren "
        "entre documentos, agrupados por bloque de diferencia."
    )

    if not all(marcas_reales_por_doc):
        nota = out.add_paragraph()
        run = nota.add_run(
            "Nota: uno o más documentos no conservaban marcas reales de "
            "paginación de Word, por lo que su número de página es una "
            "estimación y no la paginación exacta al imprimir."
        )
        run.italic = True
        run.font.size = Pt(9)

    if not bloques:
        out.add_paragraph("No se encontraron diferencias entre los documentos.")
        out.save(ruta_salida)
        return

    tabla = out.add_table(rows=1, cols=4)
    tabla.style = "Light Grid Accent 1"
    tabla.alignment = WD_TABLE_ALIGNMENT.CENTER

    anchos      = [Cm(2.0), Cm(3.2), Cm(2.3), Cm(9.5)]
    encabezados = ["Bloque", "Documento", "Pág. (aprox.)", "Texto del párrafo"]
    hdr_cells   = tabla.rows[0].cells
    for i, texto in enumerate(encabezados):
        hdr_cells[i].width = anchos[i]
        run = hdr_cells[i].paragraphs[0].add_run(texto)
        run.bold = True
        run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
        _sombrear_celda(hdr_cells[i], "2E5395")

    for num_bloque, filas in enumerate(bloques, start=1):
        # Recolectar items indexados por doc_idx para poder comparar
        items_por_doc = {}   # doc_idx -> item (o None)
        for doc_idx, item in filas:
            items_por_doc[doc_idx] = item

        # Texto base (doc 0) para referencia en el diff
        item_base = items_por_doc.get(0)
        texto_base = item_base["texto"] if item_base else ""

        for doc_idx, item in filas:
            row_cells = tabla.add_row().cells
            for i, ancho in enumerate(anchos):
                row_cells[i].width = ancho

            row_cells[0].paragraphs[0].add_run(str(num_bloque))
            row_cells[1].paragraphs[0].add_run(
                f"Doc. {doc_idx + 1}: {Path(rutas[doc_idx]).name}"
            )

            if item is None:
                row_cells[2].paragraphs[0].add_run("—")
                row_cells[3].paragraphs[0].add_run(
                    "(sin texto correspondiente / párrafo ausente)"
                ).italic = True
            else:
                row_cells[2].paragraphs[0].add_run(
                    f"{item['pagina']} (párr. {item['num']})"
                )
                texto_este = item["texto"]

                if doc_idx == 0:
                    # Fila base: resaltar palabras que difieren del primer doc que cambia
                    primer_otro_idx = next(
                        (d for d, it in items_por_doc.items() if d != 0 and it is not None),
                        None,
                    )
                    if primer_otro_idx is not None:
                        texto_ref = items_por_doc[primer_otro_idx]["texto"]
                        tokens_base, _ = diff_palabras(texto_este, texto_ref)
                    else:
                        tokens_base = [(texto_este, False)]
                    _escribir_celda_con_diff(
                        row_cells[3].paragraphs[0], tokens_base
                    )
                else:
                    # Fila de otro doc: comparar contra la base
                    _, tokens_otro = diff_palabras(texto_base, texto_este)
                    _escribir_celda_con_diff(
                        row_cells[3].paragraphs[0], tokens_otro
                    )

    tabla.autofit = False
    out.save(ruta_salida)


# ==========================================================================
# TRABAJO EN SEGUNDO PLANO  (sin tocar tkinter)
# ==========================================================================

def _comparar_en_hilo(archivos, salida, cola):
    """Corre en un hilo; se comunica con la UI solo a través de 'cola'."""
    try:
        listas, marcas = [], []
        for ruta in archivos:
            parrafos, tiene_marcas = extraer_parrafos(ruta)
            listas.append(parrafos)
            marcas.append(tiene_marcas)
            aviso = "" if tiene_marcas else "  [paginación estimada]"
            cola.put((_MSG_LOG, f"  ✓ {Path(ruta).name}: {len(parrafos)} párrafos{aviso}"))

        cola.put((_MSG_LOG, "Comparando párrafos…"))
        bloques = calcular_diferencias(listas)
        cola.put((_MSG_LOG, f"  Bloques de diferencia encontrados: {len(bloques)}"))

        cola.put((_MSG_LOG, "Generando reporte…"))
        construir_reporte(archivos, bloques, marcas, salida)
        cola.put((_MSG_OK, salida))
    except Exception as exc:
        cola.put((_MSG_ERROR, str(exc)))


# ==========================================================================
# INTERFAZ GRÁFICA
# ==========================================================================

class App:
    POLL_MS = 100   # cada cuántos ms se revisa la cola (funciona bien en macOS)

    def __init__(self, root):
        self.root = root
        self.root.title("Comparador de documentos Word")
        self.root.geometry("740x580")
        self.root.minsize(620, 480)

        self.archivos = []
        self.ruta_salida = tk.StringVar(value=str(Path.cwd() / "comparacion.docx"))
        self.ultima_ruta_generada = None
        self._cola = queue.Queue()

        self._construir_widgets()
        self._log("Listo. Agrega documentos .docx y presiona «Comparar y generar reporte».")

    # ---------------------------------------------------------------- UI --

    def _construir_widgets(self):
        PAD = {"padx": 10, "pady": 6}

        ttk.Label(
            self.root,
            text="Comparador de documentos Word (.docx)",
            font=("Helvetica", 13, "bold"),
        ).pack(anchor="w", **PAD)

        ttk.Label(
            self.root,
            text="El primer documento de la lista es el BASE contra el que se comparan los demás.",
        ).pack(anchor="w", padx=10)

        # --- Zona drag-and-drop / lista ---
        zona = ttk.LabelFrame(self.root, text="Documentos (2 – 7)")
        zona.pack(fill="both", expand=True, **PAD)

        hint = ("Arrastra archivos .docx aquí  ·  o usa «Agregar archivos…»"
                if DND_DISPONIBLE else
                "Usa «Agregar archivos…» para cargar tus documentos .docx")
        self.zona_drop = tk.Label(
            zona, text=hint, relief="groove", bd=2,
            bg="#eef2f7", fg="#444", height=2, font=("Helvetica", 11),
        )
        self.zona_drop.pack(fill="x", padx=8, pady=(8, 4))

        if DND_DISPONIBLE:
            self.zona_drop.drop_target_register(DND_FILES)
            self.zona_drop.dnd_bind("<<Drop>>", self._al_soltar_archivos)
            self.zona_drop.bind("<Enter>", lambda _: self.zona_drop.config(bg="#d6e4f5"))
            self.zona_drop.bind("<Leave>", lambda _: self.zona_drop.config(bg="#eef2f7"))

        lista_frame = ttk.Frame(zona)
        lista_frame.pack(fill="both", expand=True, padx=8, pady=4)

        self.lista = tk.Listbox(lista_frame, selectmode="browse", height=7,
                                font=("Courier", 11))
        sb = ttk.Scrollbar(lista_frame, orient="vertical", command=self.lista.yview)
        self.lista.config(yscrollcommand=sb.set)
        self.lista.pack(side="left", fill="both", expand=True)
        sb.pack(side="left", fill="y")

        bf = ttk.Frame(zona)
        bf.pack(fill="x", padx=8, pady=(0, 8))
        for txt, cmd in [
            ("Agregar archivos…", self._agregar_archivos),
            ("Subir ↑",           lambda: self._mover(-1)),
            ("Bajar ↓",           lambda: self._mover(1)),
            ("Quitar",            self._quitar_seleccionado),
            ("Limpiar lista",     self._limpiar_lista),
        ]:
            ttk.Button(bf, text=txt, command=cmd).pack(side="left", padx=2)

        # --- Salida ---
        sf = ttk.LabelFrame(self.root, text="Archivo de salida")
        sf.pack(fill="x", **PAD)
        ttk.Entry(sf, textvariable=self.ruta_salida).pack(
            side="left", fill="x", expand=True, padx=(8, 4), pady=8)
        ttk.Button(sf, text="Examinar…", command=self._elegir_salida).pack(
            side="left", padx=(0, 8))

        # --- Botones de acción ---
        af = ttk.Frame(self.root)
        af.pack(fill="x", **PAD)

        self.boton_comparar = ttk.Button(
            af, text="⚙  Comparar y generar reporte",
            command=self._ejecutar_comparacion)
        self.boton_comparar.pack(side="left", padx=(0, 6))

        self.boton_abrir = ttk.Button(
            af, text="📄  Abrir reporte",
            command=self._abrir_reporte, state="disabled")
        self.boton_abrir.pack(side="left")

        # --- Log de estado ---
        ef = ttk.LabelFrame(self.root, text="Estado")
        ef.pack(fill="both", expand=True, **PAD)

        self.texto_estado = tk.Text(
            ef, height=6, state="disabled", wrap="word",
            font=("Courier", 10), bg="#f8f8f8")
        sb2 = ttk.Scrollbar(ef, orient="vertical", command=self.texto_estado.yview)
        self.texto_estado.config(yscrollcommand=sb2.set)
        self.texto_estado.pack(side="left", fill="both", expand=True, padx=(8, 0), pady=8)
        sb2.pack(side="left", fill="y", padx=(0, 8), pady=8)

        self.barra = ttk.Progressbar(self.root, mode="indeterminate")
        self.barra.pack(fill="x", padx=10, pady=(0, 10))

    # ----------------------------------------------------------- helpers --

    def _log(self, msg):
        self.texto_estado.config(state="normal")
        self.texto_estado.insert("end", msg + "\n")
        self.texto_estado.see("end")
        self.texto_estado.config(state="disabled")

    def _refrescar_lista(self):
        self.lista.delete(0, "end")
        for i, ruta in enumerate(self.archivos, 1):
            etiq = "BASE" if i == 1 else f"#{i}  "
            self.lista.insert("end", f"[{etiq}]  {Path(ruta).name}")

    def _agregar_rutas(self, rutas):
        agregados = rechazados_ext = rechazados_cupo = 0
        for ruta in rutas:
            ruta = str(ruta).strip()
            if not ruta:
                continue
            if not ruta.lower().endswith(".docx"):
                rechazados_ext += 1
                continue
            if len(self.archivos) >= MAX_DOCS:
                rechazados_cupo += 1
                continue
            if ruta not in self.archivos:
                self.archivos.append(ruta)
                agregados += 1
        self._refrescar_lista()
        if agregados:
            self._log(f"Se agregaron {agregados} archivo(s). Total: {len(self.archivos)}.")
        if rechazados_ext:
            messagebox.showwarning("Archivos ignorados",
                f"{rechazados_ext} archivo(s) no son .docx y se ignoraron.")
        if rechazados_cupo:
            messagebox.showwarning("Máximo alcanzado",
                f"Solo se admiten {MAX_DOCS} documentos. "
                f"{rechazados_cupo} archivo(s) no se agregaron.")

    def _al_soltar_archivos(self, event):
        self._agregar_rutas(self.root.tk.splitlist(event.data))

    def _agregar_archivos(self):
        rutas = filedialog.askopenfilenames(
            title="Selecciona documentos .docx",
            filetypes=[("Documentos Word", "*.docx")])
        if rutas:
            self._agregar_rutas(rutas)

    def _quitar_seleccionado(self):
        sel = self.lista.curselection()
        if sel:
            del self.archivos[sel[0]]
            self._refrescar_lista()

    def _mover(self, delta):
        sel = self.lista.curselection()
        if not sel:
            return
        i = sel[0]
        j = i + delta
        if 0 <= j < len(self.archivos):
            self.archivos[i], self.archivos[j] = self.archivos[j], self.archivos[i]
            self._refrescar_lista()
            self.lista.selection_set(j)

    def _limpiar_lista(self):
        self.archivos.clear()
        self._refrescar_lista()

    def _elegir_salida(self):
        ruta = filedialog.asksaveasfilename(
            title="Guardar reporte como",
            defaultextension=".docx",
            filetypes=[("Documento Word", "*.docx")],
            initialfile="comparacion.docx")
        if ruta:
            self.ruta_salida.set(ruta)

    # ---------------------------------------------------------- acciones --

    def _ejecutar_comparacion(self):
        if len(self.archivos) < 2:
            messagebox.showerror("Faltan documentos",
                "Se necesitan al menos 2 documentos para comparar.")
            return
        salida = self.ruta_salida.get().strip()
        if not salida:
            messagebox.showerror("Sin ruta de salida",
                "Indica dónde guardar el reporte.")
            return

        self.boton_comparar.config(state="disabled")
        self.boton_abrir.config(state="disabled")
        self.barra.start(10)
        self._log("─" * 40)
        self._log(f"Iniciando comparación de {len(self.archivos)} documentos…")

        threading.Thread(
            target=_comparar_en_hilo,
            args=(list(self.archivos), salida, self._cola),
            daemon=True,
        ).start()

        # Empezar a drenar la cola desde el hilo principal
        self.root.after(self.POLL_MS, self._drenar_cola)

    def _drenar_cola(self):
        """Llamado por el mainloop; lee todos los mensajes disponibles y
        programa otra lectura si el hilo todavía está trabajando."""
        terminado = False
        try:
            while True:
                tipo, datos = self._cola.get_nowait()
                if tipo == _MSG_LOG:
                    self._log(datos)
                elif tipo == _MSG_OK:
                    self.ultima_ruta_generada = datos
                    self._log(f"✅ Reporte guardado en:\n   {datos}")
                    self._al_terminar_ok()
                    terminado = True
                elif tipo == _MSG_ERROR:
                    self._log(f"❌ ERROR: {datos}")
                    self._al_terminar_error(datos)
                    terminado = True
        except queue.Empty:
            pass

        if not terminado:
            self.root.after(self.POLL_MS, self._drenar_cola)

    def _al_terminar_ok(self):
        self.barra.stop()
        self.boton_comparar.config(state="normal")
        self.boton_abrir.config(state="normal")
        messagebox.showinfo("Listo", "El reporte se generó correctamente.")

    def _al_terminar_error(self, msg):
        self.barra.stop()
        self.boton_comparar.config(state="normal")
        messagebox.showerror("Error al comparar", msg)

    def _abrir_reporte(self):
        if not self.ultima_ruta_generada:
            return
        try:
            if sys.platform == "darwin":
                subprocess.run(["open", self.ultima_ruta_generada], check=False)
            elif sys.platform.startswith("win"):
                os.startfile(self.ultima_ruta_generada)
            else:
                subprocess.run(["xdg-open", self.ultima_ruta_generada], check=False)
        except Exception as exc:
            messagebox.showerror("No se pudo abrir", str(exc))


# ==========================================================================
# PUNTO DE ENTRADA
# ==========================================================================

def main():
    if DND_DISPONIBLE:
        root = TkinterDnD.Tk()
    else:
        root = tk.Tk()
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
