# aplicacion_presentismo.py
import os
import json
import pandas as pd
import datetime
from dateutil import parser
from rapidfuzz import fuzz
from PyQt6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QLabel,
    QFileDialog, QTableWidget, QTableWidgetItem, QMessageBox,
    QComboBox, QDialog, QListWidget, QLineEdit, QCheckBox, QSplitter,
    QSizePolicy, QFormLayout, QSpinBox, QCalendarWidget
)
from PyQt6.QtCore import Qt, QTimer
from PyQt6.QtGui import QColor

# Archivos persistentes
SECTORES_FILE = "sectores.json"
HORARIOS_FILE = "horarios.json"
FECHAS_ELIMINAR_FILE = "fechas_eliminar.json"


class AplicacionPresentismo(QWidget):
    """
    Pestaña Presentismo:
    - carga múltiples excels (.xls/.xlsx/.csv)
    - detecta columna Nombre y fecha/hora
    - separa Fecha y Hora
    - asigna Sector (automático + manual)
    - editor visual de sectores y horarios (por sector/persona/día, con overrides por fecha)
    - filtros por sector, fechas a excluir, primer fichaje, retrasos
    - exporta Excel con:
        * resumen de presentismo (por empleado)
        * detalle de días tarde / inasistencias
        * datos procesados a nivel día (Entrada/Salida/EstadoDia)
    """

    def __init__(self):
        super().__init__()
        self.setWindowTitle("📅 Presentismo — Control de sectores y horarios")
        self.resize(1400, 800)

        # DataFrames
        self.df = pd.DataFrame()          # Data procesada actual (fichadas crudas)
        self.df_original = pd.DataFrame() # Raw concatenado
        self.display_df = pd.DataFrame()  # Lo que se muestra (subset o filtrado)

        # Cargas persistentes
        self.sectores = self._load_json(SECTORES_FILE, default=self._default_sectores())
        self.horarios = self._load_json(HORARIOS_FILE, default={})

        # Fechas a eliminar
        self.fechas_a_eliminar = self._load_json(FECHAS_ELIMINAR_FILE, default=[])

        # UI
        self._init_ui()

        # Poblar lista de fechas a eliminar
        for fecha in self.fechas_a_eliminar:
            self.list_fechas_eliminar.addItem(fecha)

    # -------------------------
    # Utilidades persistencia
    # -------------------------
    def _default_sectores(self):
        return {
            "Choferes/Varios": ["Acosta Cristian Gustavo", "Flores Lucas Sebastian"],
            "Chofer": [
                "Aracuyu Miguel", "Barbagallo Daniel", "Barboza Carlos",
                "Cabrera Mario", "Campopiano Rube", "Cantero Julio",
                "Franco Juan Carlos", "Gomez Hernan Ariel", "Guerra Carlos",
                "Manetto Jose", "Marchano Cristian", "Meo Pablo",
                "Perez Mauro", "Suarez Bringas Cristian"
            ],
            "Ayudante de Reparto": [
                "Alvarez Carlos", "Barbosa Ivan", "Bianchini Leandro", "Bianchini Lucas",
                "Cardenas Gonzales Carlos", "Carrozzo Cesar", "Esteban Matias",
                "Gonzalez Gonzalo", "Gonzalez Matias", "Graneros Cristian",
                "Lopez Marcelo", "Martin Ivan Ezequiel", "Medina Miguel",
                "Meo Claudio", "Monsalve Cuicas Richard", "Pereyra Fernando Ariel",
                "Rodriguez Ramiro", "Tranfo Elias"
            ],
            "Carga y Descarga Guido": [
                "Diaz Pablo", "Farias Yamil", "Gabillot Pablo Daniel",
                "Morandi Joel Alan", "Sosa Fernando"
            ],
            "Almacen": [
                "Amarilla Agnelli Lautaro", "Cornejo Daniel", "Graño Lautaro",
                "Lamarque Joaquin", "Lopez Brian Emanuel", "Moliterno Lautaro"
            ]
        }

    def _load_json(self, path, default=None):
        if os.path.exists(path):
            try:
                with open(path, "r", encoding="utf-8") as f:
                    return json.load(f)
            except Exception:
                return default if default is not None else {}
        return default if default is not None else {}

    def _save_json(self, path, data):
        try:
            with open(path, "w", encoding="utf-8") as f:
                json.dump(data, f, indent=2, ensure_ascii=False)
        except Exception as e:
            raise

    # -------------------------
    # UI principal
    # -------------------------
    def _init_ui(self):
        root = QHBoxLayout(self)

        # columna principal (tabla + controles)
        main_col = QVBoxLayout()

        # Título
        title = QLabel("📅 Presentismo — Carga y asignación de sectores")
        title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        main_col.addWidget(title)

        # Botones fila superior
        fila0 = QHBoxLayout()
        self.btn_cargar = QPushButton("📂 Cargar Excel(s)")
        self.btn_cargar.clicked.connect(self.cargar_varios_excels)
        fila0.addWidget(self.btn_cargar)

        self.btn_export = QPushButton("💾 Exportar resultado (Excel)")
        self.btn_export.clicked.connect(self.exportar_resultado)
        fila0.addWidget(self.btn_export)

        self.btn_export_filtrado = QPushButton("💾 Exportar vista actual")
        self.btn_export_filtrado.clicked.connect(self.exportar_vista_actual)
        fila0.addWidget(self.btn_export_filtrado)

        self.btn_export_no = QPushButton("📤 Exportar no-asignados (CSV)")
        self.btn_export_no.clicked.connect(self.exportar_no_asignados_csv)
        fila0.addWidget(self.btn_export_no)

        main_col.addLayout(fila0)

        # Botones secundarios
        fila1 = QHBoxLayout()
        self.btn_asign_manual = QPushButton("✏️ Asignar sector (manual, visual)")
        self.btn_asign_manual.clicked.connect(self.abrir_asignar_sector_visual)
        fila1.addWidget(self.btn_asign_manual)

        self.btn_editar_sectores = QPushButton("⚙️ Editar sectores")
        self.btn_editar_sectores.clicked.connect(self.abrir_editor_sectores)
        fila1.addWidget(self.btn_editar_sectores)

        self.btn_editar_horarios = QPushButton("⏰ Editar horarios")
        self.btn_editar_horarios.clicked.connect(self.abrir_editor_horarios)
        fila1.addWidget(self.btn_editar_horarios)

        main_col.addLayout(fila1)

        # Tabla principal (muestra fichadas crudas)
        self.tabla = QTableWidget()
        self.tabla.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
        self.tabla.setColumnCount(7)
        self.tabla.setHorizontalHeaderLabels(
            ["Nombre", "Fecha/Hora (raw)", "Fecha", "Hora", "Sector", "Retraso", "Horario Ref"]
        )
        main_col.addWidget(self.tabla)

        root.addLayout(main_col, 4)

        # panel derecho: filtros y opciones
        right_col = QVBoxLayout()
        lbl = QLabel("🧰 Panel de filtros y opciones")
        lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)
        right_col.addWidget(lbl)

        self.chk_primero = QCheckBox("Primer fichaje del día (vista)")
        right_col.addWidget(self.chk_primero)

        self.chk_retrasos = QCheckBox("Detectar retrasos (usar horarios configurados)")
        right_col.addWidget(self.chk_retrasos)

        right_col.addSpacing(6)
        right_col.addWidget(QLabel("Umbral fuzzy asignación (min score 0-100):"))
        self.spin_umbral = QSpinBox()
        self.spin_umbral.setRange(40, 100)
        self.spin_umbral.setValue(85)
        right_col.addWidget(self.spin_umbral)

        btn_apply = QPushButton("✅ Aplicar filtros y recalcular (vista)")
        btn_apply.clicked.connect(self.aplicar_filtros)
        right_col.addWidget(btn_apply)

        right_col.addSpacing(8)
        self.btn_ver_sin = QPushButton("👀 Ver personas sin sector")
        self.btn_ver_sin.clicked.connect(self.mostrar_sin_sector)
        right_col.addWidget(self.btn_ver_sin)

        self.btn_ver_todos = QPushButton("🔁 Ver todos")
        self.btn_ver_todos.clicked.connect(self._mostrar_todos)
        right_col.addWidget(self.btn_ver_todos)

        right_col.addSpacing(8)
        right_col.addWidget(QLabel("🏢 Filtrar por sectores:"))
        self.list_sectores_filtro = QListWidget()
        self.list_sectores_filtro.setSelectionMode(QListWidget.SelectionMode.MultiSelection)
        self.list_sectores_filtro.setMaximumHeight(100)
        # Poblar con sectores ordenados
        for sec in sorted(self.sectores.keys()):
            self.list_sectores_filtro.addItem(sec)
        right_col.addWidget(self.list_sectores_filtro)

        right_col.addSpacing(8)
        right_col.addWidget(QLabel("📅 Fechas a eliminar (no considerar en presentismo):"))
        self.list_fechas_eliminar = QListWidget()
        self.list_fechas_eliminar.setMaximumHeight(100)
        right_col.addWidget(self.list_fechas_eliminar)

        self.calendario = QCalendarWidget()
        self.calendario.setMaximumHeight(200)
        right_col.addWidget(self.calendario)

        fila_fechas = QHBoxLayout()
        btn_agregar_fecha = QPushButton("➕ Agregar fecha")
        btn_agregar_fecha.clicked.connect(self.agregar_fecha_eliminar)
        fila_fechas.addWidget(btn_agregar_fecha)

        btn_eliminar_fecha = QPushButton("🗑️ Eliminar fecha")
        btn_eliminar_fecha.clicked.connect(self.eliminar_fecha_eliminar)
        fila_fechas.addWidget(btn_eliminar_fecha)
        right_col.addLayout(fila_fechas)

        right_col.addStretch()
        root.addLayout(right_col, 1)

        # timer placeholder
        self._auto_refresh_timer = QTimer()
        self._auto_refresh_timer.setInterval(800)
        self._auto_refresh_timer.timeout.connect(lambda: None)

    # -------------------------
    # Métodos para gestionar fechas a eliminar
    # -------------------------
    def agregar_fecha_eliminar(self):
        selected_date = self.calendario.selectedDate().toString("yyyy-MM-dd")
        if selected_date not in self.fechas_a_eliminar:
            self.fechas_a_eliminar.append(selected_date)
            self.list_fechas_eliminar.addItem(selected_date)
            self._save_json(FECHAS_ELIMINAR_FILE, self.fechas_a_eliminar)
            QMessageBox.information(self, "Agregado", f"Fecha {selected_date} agregada a eliminar.")

    def eliminar_fecha_eliminar(self):
        current_item = self.list_fechas_eliminar.currentItem()
        if current_item:
            fecha = current_item.text()
            if fecha in self.fechas_a_eliminar:
                self.fechas_a_eliminar.remove(fecha)
                self.list_fechas_eliminar.takeItem(self.list_fechas_eliminar.row(current_item))
                self._save_json(FECHAS_ELIMINAR_FILE, self.fechas_a_eliminar)
                QMessageBox.information(self, "Eliminado", f"Fecha {fecha} eliminada.")
        else:
            QMessageBox.warning(self, "Atención", "Seleccioná una fecha para eliminar.")

    # -------------------------
    # Leer múltiples Excel
    # -------------------------
    def cargar_varios_excels(self):
        archivos, _ = QFileDialog.getOpenFileNames(
            self, "Seleccionar archivos Excel", "", "Archivos Excel (*.xlsx *.xls *.csv)"
        )
        if not archivos:
            return

        dfs = []
        errores = []
        for archivo in archivos:
            try:
                ext = os.path.splitext(archivo)[1].lower()
                if ext == ".xls":
                    df_temp = pd.read_excel(archivo, engine="xlrd")
                elif ext == ".csv":
                    df_temp = pd.read_csv(archivo, encoding="latin1")
                else:
                    df_temp = pd.read_excel(archivo, engine="openpyxl")
                if not df_temp.empty:
                    dfs.append(df_temp)
            except Exception as e:
                errores.append(f"{os.path.basename(archivo)}: {e}")

        if errores:
            QMessageBox.warning(
                self,
                "Aviso lectura archivos",
                "Algunos archivos no se pudieron leer:\n" + "\n".join(errores)
            )

        if not dfs:
            QMessageBox.information(self, "Sin datos", "No se cargaron datos válidos.")
            return

        # concatenar todo
        df_all = pd.concat(dfs, ignore_index=True, sort=False)
        self.df_original = df_all.copy()

        # detectar columna Nombre
        cols = [c for c in df_all.columns]
        col_nombre = None
        for c in cols:
            if "nombre" == str(c).strip().lower() or "name" == str(c).strip().lower():
                col_nombre = c
                break
        if col_nombre is None:
            for c in cols:
                cl = str(c).lower()
                if "nom" in cl or "name" in cl:
                    col_nombre = c
                    break
        if col_nombre is None:
            QMessageBox.critical(self, "Error", "No se pudo identificar la columna 'Nombre'. Revisa tus archivos.")
            return

        # detectar columna Fecha/Hora
        col_fecha = None
        for c in cols:
            cl = str(c).lower()
            if "fecha/hora" in cl or "fecha hora" in cl or "fecha_hora" in cl or cl.strip() == "fecha":
                col_fecha = c
                break
        if col_fecha is None:
            for c in cols:
                cl = str(c).lower()
                if "marc" in cl or "marc." in cl or "marca" in cl or ("date" in cl and "time" in cl):
                    col_fecha = c
                    break
        if col_fecha is None:
            if len(cols) >= 2:
                col_fecha = cols[1]
            else:
                QMessageBox.critical(self, "Error", "No se pudo identificar la columna de Fecha/Hora.")
                return

        df_min = pd.DataFrame()
        df_min["Nombre"] = df_all[col_nombre]
        df_min["FechaHora_raw"] = df_all[col_fecha].astype(str).fillna("")

        # parseo robusto de fecha/hora
        def parse_seguro(val_raw):
            if val_raw is None:
                return None
            s = str(val_raw).strip()
            if s == "" or s.lower() in ("nan", "none", "na", "n/a"):
                return None
            s = s.replace("\xa0", " ").replace("\t", " ").replace("\r", " ").replace("\n", " ")
            s = " ".join(s.split())
            formatos = [
                "%d/%m/%Y %H:%M:%S", "%d/%m/%Y %H:%M", "%d/%m/%Y %I:%M:%S %p", "%d/%m/%Y %I:%M %p",
                "%m/%d/%Y %H:%M:%S", "%m/%d/%Y %I:%M:%S %p", "%Y-%m-%d %H:%M:%S",
                "%d-%m-%Y %H:%M:%S", "%d-%m-%Y %I:%M:%S %p"
            ]
            for fmt in formatos:
                try:
                    dt = pd.to_datetime(s, format=fmt, dayfirst=True, errors="raise")
                    return dt
                except Exception:
                    continue
            try:
                dt = parser.parse(s, dayfirst=True, fuzzy=True)
                return pd.to_datetime(dt)
            except Exception:
                try:
                    dt = parser.parse(s, dayfirst=False, fuzzy=True)
                    return pd.to_datetime(dt)
                except Exception:
                    return None

        df_min["FechaHora_parsed"] = df_min["FechaHora_raw"].apply(parse_seguro)
        df_min["Fecha"] = df_min["FechaHora_parsed"].dt.date
        df_min["Hora"] = df_min["FechaHora_parsed"].dt.time
        df_min["Hora_text"] = df_min.apply(
            lambda r: "" if pd.notna(r["Hora"]) else r["FechaHora_raw"], axis=1
        )

        # sector automático
        df_min["Sector"] = df_min["Nombre"].apply(self._auto_asign_sector_with_threshold)
        df_min["MatchedName"] = ""
        df_min["MatchScore"] = 0.0

        self.df = df_min.reset_index(drop=True)
        self.display_df = self.df.copy()
        self._actualizar_tabla(self.display_df)

        QMessageBox.information(self, "Listo", f"Cargados {len(self.df)} registros (raw).")

    # -------------------------
    # Asignación automática de sector
    # -------------------------
    def _auto_asign_sector_with_threshold(self, nombre):
        umbral = int(self.spin_umbral.value()) if hasattr(self, "spin_umbral") else 85
        return self._asignar_sector(nombre, umbral=umbral, guardar_matchinfo=False)

    def _asignar_sector(self, nombre, umbral=85, guardar_matchinfo=True):
        if not isinstance(nombre, str) or not nombre.strip():
            return "No asignado"
        nombre_proc = self._normalize_text(nombre)
        mejor_score = 0.0
        mejor_sector = "No asignado"
        mejor_persona = ""
        for sector, lista in self.sectores.items():
            for persona in lista:
                persona_proc = self._normalize_text(persona)
                s1 = fuzz.token_set_ratio(nombre_proc, persona_proc)
                s2 = fuzz.token_sort_ratio(nombre_proc, persona_proc)
                s3 = fuzz.partial_ratio(nombre_proc, persona_proc)
                score = 0.5 * s1 + 0.3 * s2 + 0.2 * s3
                tokens_n = set(nombre_proc.split())
                tokens_p = set(persona_proc.split())
                tokens_shared = len(tokens_n & tokens_p)
                if tokens_shared >= 1:
                    score += tokens_shared * 6
                if score > mejor_score:
                    mejor_score = score
                    mejor_sector = sector
                    mejor_persona = persona
        if mejor_score >= umbral:
            return mejor_sector
        return "No asignado"

    def _normalize_text(self, s):
        if s is None:
            return ""
        s = str(s).lower()
        for ch in [",", ".", ";", "-", "_", "(", ")", "/", "\"", "'"]:
            s = s.replace(ch, " ")
        s = " ".join(s.split())
        return s.strip()

    # -------------------------
    # Tabla (vista cruda)
    # -------------------------
    def _actualizar_tabla(self, df=None):
        if df is None:
            df = self.df
        if df is None or df.empty:
            self.tabla.clearContents()
            self.tabla.setRowCount(0)
            return

        if "Delay" not in df.columns:
            self.tabla.clearContents()

        self.display_df = df.reset_index(drop=True)
        self.tabla.setRowCount(len(self.display_df))
        self.tabla.setColumnCount(7)
        self.tabla.setHorizontalHeaderLabels(
            ["Nombre", "Fecha/Hora (raw)", "Fecha", "Hora", "Sector", "Retraso", "Horario Ref"]
        )

        for i, row in self.display_df.iterrows():
            item_nom = QTableWidgetItem(str(row["Nombre"]))
            self.tabla.setItem(i, 0, item_nom)

            item_raw = QTableWidgetItem(str(row.get("FechaHora_raw", "")))
            self.tabla.setItem(i, 1, item_raw)

            fecha_display = "" if pd.isna(row.get("Fecha")) or row.get("Fecha") is None else str(row.get("Fecha"))
            item_fecha = QTableWidgetItem(fecha_display)
            self.tabla.setItem(i, 2, item_fecha)

            hora_val = row.get("Hora")
            if pd.notna(hora_val) and hora_val is not None:
                hora_display = hora_val.strftime("%H:%M:%S") if hasattr(hora_val, "strftime") else str(hora_val)
            else:
                hora_display = row.get("Hora_text", "") if "Hora_text" in row.index else ""
            item_hora = QTableWidgetItem(hora_display)
            self.tabla.setItem(i, 3, item_hora)

            item_sector = QTableWidgetItem(str(row.get("Sector", "")))
            if str(row.get("Sector", "")).strip().lower() in ("no asignado", "", "none"):
                item_sector.setBackground(QColor(255, 255, 200))
            self.tabla.setItem(i, 4, item_sector)

            delay = row.get("Delay", "")
            delay_text = ""
            is_late = False
            is_early = False
            if isinstance(delay, pd.Timedelta) or hasattr(delay, "total_seconds"):
                try:
                    total_s = int(delay.total_seconds())
                    if total_s < 0:
                        total_s = abs(total_s)
                        h = total_s // 3600
                        m = (total_s % 3600) // 60
                        s = total_s % 60
                        delay_text = f"Early: {h:02d}:{m:02d}:{s:02d}"
                        is_early = True
                    elif total_s >= 60:
                        h = total_s // 3600
                        m = (total_s % 3600) // 60
                        s = total_s % 60
                        delay_text = f"{h:02d}:{m:02d}:{s:02d}"
                        is_late = True
                    else:
                        delay_text = "00:00:00"
                except Exception:
                    delay_text = str(delay)
            elif isinstance(delay, str) and delay:
                delay_text = delay
            else:
                delay_text = ""

            item_delay = QTableWidgetItem(delay_text)
            if is_late:
                for col_idx in range(7):
                    it = self.tabla.item(i, col_idx) or QTableWidgetItem("")
                    it.setBackground(QColor(255, 255, 150))
                    self.tabla.setItem(i, col_idx, it)
            elif is_early:
                for col_idx in range(7):
                    it = self.tabla.item(i, col_idx) or QTableWidgetItem("")
                    it.setBackground(QColor(150, 255, 150))
                    self.tabla.setItem(i, col_idx, it)
            self.tabla.setItem(i, 5, item_delay)

            horario_ref = str(row.get("Horario_Referencia", ""))
            item_horario_ref = QTableWidgetItem(horario_ref)
            self.tabla.setItem(i, 6, item_horario_ref)

        self.tabla.resizeColumnsToContents()

    # -------------------------
    # Mostrar solo sin sector / todos
    # -------------------------
    def mostrar_sin_sector(self):
        if self.df.empty:
            QMessageBox.information(self, "Info", "No hay datos cargados.")
            return
        df_sin = self.df[self.df["Sector"] == "No asignado"]
        if df_sin.empty:
            QMessageBox.information(self, "Info", "No hay personas sin sector.")
            return
        self._actualizar_tabla(df_sin)

    def _mostrar_todos(self):
        self._actualizar_tabla(self.df)

    # -------------------------
    # Ventana visual para asignar sector
    # -------------------------
    def abrir_asignar_sector_visual(self):
        if self.df.empty:
            QMessageBox.warning(self, "Atención", "Primero cargá datos.")
            return

        dlg = QDialog(self)
        dlg.setWindowTitle("Asignar sector visual")
        dlg.resize(900, 500)
        root = QHBoxLayout(dlg)

        left = QVBoxLayout()
        left.addWidget(QLabel("Personas (seleccionar una o varias):"))
        list_personas = QListWidget()
        nombres = sorted(self.df["Nombre"].dropna().unique().tolist())
        list_personas.addItems(nombres)
        list_personas.setSelectionMode(QListWidget.SelectionMode.ExtendedSelection)
        left.addWidget(list_personas)
        root.addLayout(left, 2)

        right = QVBoxLayout()
        right.addWidget(QLabel("Sectores:"))
        list_sectores = QListWidget()
        list_sectores.addItems(sorted(self.sectores.keys()))
        right.addWidget(list_sectores)

        right.addWidget(QLabel("Acciones:"))
        btn_asign = QPushButton("Asignar sector seleccionado a las personas")
        right.addWidget(btn_asign)

        btn_asign_and_add = QPushButton("Asignar y agregar persona(s) al sector (persistir)")
        right.addWidget(btn_asign_and_add)

        btn_close = QPushButton("Cerrar")
        right.addWidget(btn_close)

        root.addLayout(right, 1)

        def do_assign(add_to_sector=False):
            sel_personas = [it.text() for it in list_personas.selectedItems()]
            sel_sector_item = list_sectores.currentItem()
            if not sel_personas:
                QMessageBox.warning(dlg, "Atención", "Seleccioná al menos una persona.")
                return
            if not sel_sector_item:
                QMessageBox.warning(dlg, "Atención", "Seleccioná un sector.")
                return
            sector = sel_sector_item.text()
            for p in sel_personas:
                self.df.loc[self.df["Nombre"] == p, "Sector"] = sector
                if add_to_sector:
                    lst = self.sectores.get(sector, [])
                    if not any(self._normalize_text(x) == self._normalize_text(p) for x in lst):
                        lst.append(p)
                        self.sectores[sector] = sorted(lst)
            if add_to_sector:
                self._save_json(SECTORES_FILE, self.sectores)
            self._actualizar_tabla(self.df)
            QMessageBox.information(dlg, "Asignado", f"Asigné {len(sel_personas)} persona(s) al sector '{sector}'.")

        btn_asign.clicked.connect(lambda: do_assign(add_to_sector=False))
        btn_asign_and_add.clicked.connect(lambda: do_assign(add_to_sector=True))
        btn_close.clicked.connect(dlg.accept)
        dlg.exec()

    # -------------------------
    # Editor de sectores
    # -------------------------
    def abrir_editor_sectores(self):
        dlg = QDialog(self)
        dlg.setWindowTitle("Editor de sectores")
        dlg.resize(1000, 600)
        root = QHBoxLayout(dlg)

        left = QVBoxLayout()
        left.addWidget(QLabel("Sectores:"))
        lw_sect = QListWidget()
        for s in sorted(self.sectores.keys()):
            lw_sect.addItem(s)
        left.addWidget(lw_sect)
        inp_new_sector = QLineEdit()
        inp_new_sector.setPlaceholderText("Nuevo sector...")
        left.addWidget(inp_new_sector)
        btn_add_sector = QPushButton("➕ Agregar sector")
        btn_del_sector = QPushButton("🗑️ Eliminar sector seleccionado")
        left.addWidget(btn_add_sector)
        left.addWidget(btn_del_sector)
        root.addLayout(left, 1)

        right = QVBoxLayout()
        right.addWidget(QLabel("Empleados del sector seleccionado:"))
        lw_emps = QListWidget()
        right.addWidget(lw_emps)
        inp_new_emp = QLineEdit()
        inp_new_emp.setPlaceholderText("Agregar empleado (apellido y nombre)...")
        right.addWidget(inp_new_emp)
        btn_add_emp = QPushButton("➕ Agregar empleado")
        btn_del_emp = QPushButton("🗑️ Eliminar empleado seleccionado")
        right.addWidget(btn_add_emp)
        right.addWidget(btn_del_emp)
        btn_guardar = QPushButton("Guardar y cerrar")
        right.addWidget(btn_guardar)
        root.addLayout(right, 2)

        def refresh_emps():
            lw_emps.clear()
            cur = lw_sect.currentItem()
            if not cur:
                return
            sec = cur.text()
            for e in sorted(self.sectores.get(sec, [])):
                lw_emps.addItem(e)

        def add_sector():
            name = inp_new_sector.text().strip()
            if not name:
                return
            if name in self.sectores:
                QMessageBox.warning(dlg, "Aviso", "Ese sector ya existe.")
                return
            self.sectores[name] = []
            lw_sect.addItem(name)
            inp_new_sector.clear()

        def del_sector():
            cur = lw_sect.currentItem()
            if not cur:
                return
            sec = cur.text()
            if QMessageBox.question(
                dlg, "Confirmar", f"Eliminar sector '{sec}' y sus empleados?"
            ) == QMessageBox.StandardButton.Yes:
                self.sectores.pop(sec, None)
                lw_sect.takeItem(lw_sect.row(cur))
                lw_emps.clear()

        def add_emp():
            cur = lw_sect.currentItem()
            if not cur:
                QMessageBox.warning(dlg, "Aviso", "Seleccioná un sector.")
                return
            sec = cur.text()
            name = inp_new_emp.text().strip()
            if not name:
                return
            lst = self.sectores.get(sec, [])
            if not any(self._normalize_text(x) == self._normalize_text(name) for x in lst):
                lst.append(name)
                self.sectores[sec] = sorted(lst)
                refresh_emps()
                inp_new_emp.clear()
            else:
                QMessageBox.information(dlg, "Info", "El empleado ya está en el sector.")

        def del_emp():
            cur_sec = lw_sect.currentItem()
            cur_emp = lw_emps.currentItem()
            if not cur_sec or not cur_emp:
                return
            sec = cur_sec.text()
            emp = cur_emp.text()
            if emp in self.sectores.get(sec, []):
                self.sectores[sec].remove(emp)
                refresh_emps()

        def save_and_close():
            try:
                self._save_json(SECTORES_FILE, self.sectores)
                QMessageBox.information(dlg, "Guardado", "Sectores guardados en sectores.json")
                dlg.accept()
            except Exception as e:
                QMessageBox.critical(dlg, "Error guardado", str(e))

        lw_sect.itemClicked.connect(lambda _: refresh_emps())
        btn_add_sector.clicked.connect(add_sector)
        btn_del_sector.clicked.connect(del_sector)
        btn_add_emp.clicked.connect(add_emp)
        btn_del_emp.clicked.connect(del_emp)
        btn_guardar.clicked.connect(save_and_close)

        dlg.exec()

    # -------------------------
    # Editor de horarios (por sector/persona/día y overrides)
    # -------------------------
    class HorariosDialog(QDialog):
        def __init__(self, parent):
            super().__init__(parent)
            self.parent = parent

        def closeEvent(self, event):
            try:
                self.parent._save_json(HORARIOS_FILE, self.parent.horarios)
                self.parent._save_json(SECTORES_FILE, self.parent.sectores)
            except Exception as e:
                QMessageBox.warning(self, "Error guardado", str(e))
            event.accept()

    def abrir_editor_horarios(self):
        dlg = self.HorariosDialog(self)
        dlg.setWindowTitle("Editor de horarios (por sector y persona, con días y overrides)")
        dlg.resize(1400, 800)
        root = QHBoxLayout(dlg)

        left = QVBoxLayout()
        left.addWidget(QLabel("Sectores:"))
        lw_sect = QListWidget()
        for s in sorted(self.sectores.keys()):
            lw_sect.addItem(s)
        left.addWidget(lw_sect)
        inp_new_sector = QLineEdit()
        inp_new_sector.setPlaceholderText("Nuevo sector...")
        left.addWidget(inp_new_sector)
        btn_add_sector = QPushButton("➕ Agregar sector")
        left.addWidget(btn_add_sector)
        root.addLayout(left, 1)

        right = QVBoxLayout()
        right.addWidget(QLabel("Horario del sector por día (HH:MM) — deja vacío si no aplica:"))

        dias = ["lunes", "martes", "miércoles", "jueves", "viernes", "sábado", "domingo"]
        sector_form = QFormLayout()
        sector_inputs = {}
        for dia in dias:
            inp = QLineEdit()
            inp.setPlaceholderText(f"{dia.capitalize()} HH:MM")
            sector_form.addRow(f"{dia.capitalize()}:", inp)
            sector_inputs[dia] = inp
        right.addLayout(sector_form)

        btn_set_sector_hora = QPushButton("Guardar horario sector")
        right.addWidget(btn_set_sector_hora)

        right.addWidget(QLabel("Empleados del sector (seleccione para editar su horario por día):"))
        lw_emps = QListWidget()
        right.addWidget(lw_emps)

        right.addWidget(QLabel("Horario empleado seleccionado por día (HH:MM):"))
        emp_form = QFormLayout()
        emp_inputs = {}
        for dia in dias:
            inp = QLineEdit()
            inp.setPlaceholderText(f"{dia.capitalize()} HH:MM")
            emp_form.addRow(f"{dia.capitalize()}:", inp)
            emp_inputs[dia] = inp
        right.addLayout(emp_form)

        btn_set_emp_hora = QPushButton("Guardar horario empleado")
        btn_remove_emp_hora = QPushButton("Eliminar horario de empleado")
        right.addWidget(btn_set_emp_hora)
        right.addWidget(btn_remove_emp_hora)

        right.addWidget(QLabel("Overrides por fecha específica (para cambios temporales):"))
        override_form = QFormLayout()
        self.override_date = QLineEdit()
        self.override_date.setPlaceholderText("Fecha (YYYY-MM-DD)")
        override_form.addRow("Fecha:", self.override_date)
        self.override_emp = QComboBox()
        override_form.addRow("Empleado:", self.override_emp)
        self.override_time = QLineEdit()
        self.override_time.setPlaceholderText("Horario (HH:MM)")
        override_form.addRow("Horario:", self.override_time)
        right.addLayout(override_form)

        btn_add_override = QPushButton("Agregar override")
        btn_remove_override = QPushButton("Eliminar override seleccionado")
        right.addWidget(btn_add_override)
        right.addWidget(btn_remove_override)

        self.lw_overrides = QListWidget()
        right.addWidget(QLabel("Overrides actuales:"))
        right.addWidget(self.lw_overrides)

        btn_close = QPushButton("Cerrar y guardar todo")
        right.addStretch()
        right.addWidget(btn_close)
        root.addLayout(right, 2)

        def refresh_overrides():
            self.lw_overrides.clear()
            overrides = self.horarios.get("__overrides__", {})
            for fecha, emps in overrides.items():
                for emp, time in emps.items():
                    self.lw_overrides.addItem(f"{fecha} - {emp}: {time}")

        def refresh_emps_and_sector():
            lw_emps.clear()
            for dia, inp in sector_inputs.items():
                inp.setText("")
            cur = lw_sect.currentItem()
            if not cur:
                return
            sec = cur.text()
            if sec in self.horarios:
                sec_entry = self.horarios[sec]
                if isinstance(sec_entry, dict) and "__sector__" in sec_entry:
                    sector_horas = sec_entry["__sector__"]
                    if isinstance(sector_horas, dict):
                        for dia in dias:
                            sector_inputs[dia].setText(sector_horas.get(dia, ""))
                    elif isinstance(sector_horas, str):
                        for dia in dias:
                            sector_inputs[dia].setText(sector_horas)
            miembros = sorted(self.sectores.get(sec, []))
            for m in miembros:
                lw_emps.addItem(m)
            self.override_emp.clear()
            self.override_emp.addItems(miembros)
            refresh_overrides()

        def add_sector_here():
            s = inp_new_sector.text().strip()
            if not s:
                return
            if s in self.sectores:
                QMessageBox.information(dlg, "Info", "Sector ya existe.")
                return
            self.sectores[s] = []
            if s not in self.horarios:
                self.horarios[s] = {"__sector__": {}}
            lw_sect.addItem(s)
            inp_new_sector.clear()

        def set_sector_hora():
            cur = lw_sect.currentItem()
            if not cur:
                QMessageBox.warning(dlg, "Aviso", "Seleccioná un sector.")
                return
            sec = cur.text()
            horas = {}
            for dia in dias:
                h = sector_inputs[dia].text().strip()
                if h:
                    horas[dia] = h
            if not horas:
                if sec in self.horarios and "__sector__" in self.horarios[sec]:
                    del self.horarios[sec]["__sector__"]
            else:
                self.horarios.setdefault(sec, {})
                self.horarios[sec]["__sector__"] = horas
            QMessageBox.information(dlg, "Guardado", f"Horario sector {sec} guardado.")

        def set_emp_hora():
            cur_sec = lw_sect.currentItem()
            cur_emp = lw_emps.currentItem()
            if not cur_sec or not cur_emp:
                QMessageBox.warning(dlg, "Aviso", "Seleccioná sector y empleado.")
                return
            sec = cur_sec.text()
            emp = cur_emp.text()
            horas = {}
            for dia in dias:
                h = emp_inputs[dia].text().strip()
                if h:
                    horas[dia] = h
            if not horas:
                QMessageBox.warning(dlg, "Aviso", "Ingresá al menos un horario.")
                return
            self.horarios.setdefault(sec, {})
            self.horarios[sec][emp] = horas
            QMessageBox.information(dlg, "Guardado", f"Horario {emp} en {sec} guardado.")

        def remove_emp_hora():
            cur_sec = lw_sect.currentItem()
            cur_emp = lw_emps.currentItem()
            if not cur_sec or not cur_emp:
                return
            sec = cur_sec.text()
            emp = cur_emp.text()
            if sec in self.horarios and emp in self.horarios[sec]:
                del self.horarios[sec][emp]
                QMessageBox.information(dlg, "Eliminado", f"Horario de {emp} en {sec} eliminado.")

        def on_emp_click():
            for dia, inp in emp_inputs.items():
                inp.setText("")
            cur_sec = lw_sect.currentItem()
            cur_emp = lw_emps.currentItem()
            if not cur_sec or not cur_emp:
                return
            sec = cur_sec.text()
            emp = cur_emp.text()
            if sec in self.horarios and emp in self.horarios[sec]:
                emp_horas = self.horarios[sec][emp]
                if isinstance(emp_horas, dict):
                    for dia in dias:
                        emp_inputs[dia].setText(emp_horas.get(dia, ""))
                elif isinstance(emp_horas, str):
                    for dia in dias:
                        emp_inputs[dia].setText(emp_horas)

        def add_override():
            fecha = self.override_date.text().strip()
            emp = self.override_emp.currentText()
            time_str = self.override_time.text().strip()
            if not fecha or not emp or not time_str:
                QMessageBox.warning(dlg, "Aviso", "Completá todos los campos.")
                return
            self.horarios.setdefault("__overrides__", {})
            self.horarios["__overrides__"].setdefault(fecha, {})
            self.horarios["__overrides__"][fecha][emp] = time_str
            refresh_overrides()
            QMessageBox.information(dlg, "Agregado", f"Override agregado para {emp} en {fecha}: {time_str}")

        def remove_override():
            cur = self.lw_overrides.currentItem()
            if not cur:
                QMessageBox.warning(dlg, "Aviso", "Seleccioná un override para eliminar.")
                return
            text = cur.text()
            parts = text.split(" - ")
            if len(parts) != 2:
                return
            fecha = parts[0]
            emp_time = parts[1].split(": ")
            if len(emp_time) != 2:
                return
            emp = emp_time[0]
            if (
                "__overrides__" in self.horarios
                and fecha in self.horarios["__overrides__"]
                and emp in self.horarios["__overrides__"][fecha]
            ):
                del self.horarios["__overrides__"][fecha][emp]
                if not self.horarios["__overrides__"][fecha]:
                    del self.horarios["__overrides__"][fecha]
                refresh_overrides()
                QMessageBox.information(dlg, "Eliminado", f"Override eliminado para {emp} en {fecha}")

        lw_sect.itemClicked.connect(lambda _: refresh_emps_and_sector())
        lw_emps.itemClicked.connect(lambda _: on_emp_click())
        btn_add_sector.clicked.connect(add_sector_here)
        btn_set_sector_hora.clicked.connect(set_sector_hora)
        btn_set_emp_hora.clicked.connect(set_emp_hora)
        btn_remove_emp_hora.clicked.connect(remove_emp_hora)
        btn_add_override.clicked.connect(add_override)
        btn_remove_override.clicked.connect(remove_override)

        def close_and_save():
            try:
                self._save_json(SECTORES_FILE, self.sectores)
                self._save_json(HORARIOS_FILE, self.horarios)
            except Exception as e:
                QMessageBox.warning(dlg, "Error guardado", str(e))
            dlg.accept()

        btn_close.clicked.connect(close_and_save)
        dlg.exec()

    # -------------------------
    # Helper horario referencia (para resumen diario)
    # -------------------------
    def _get_ref_time_for(self, nombre, sector, fecha, fecha_parsed_example=None):
        """
        Devuelve el horario de referencia (string HH:MM) para una persona/sector/fecha
        usando horarios.json (con overrides y horario por día de la semana).
        """
        dias_semana = ["lunes", "martes", "miércoles", "jueves", "viernes", "sábado", "domingo"]
        fecha_str = str(fecha)

        # Overrides por fecha específica
        if "__overrides__" in self.horarios:
            ov = self.horarios["__overrides__"]
            if fecha_str in ov and nombre in ov[fecha_str]:
                return ov[fecha_str][nombre]

        # Determinar día de la semana
        if isinstance(fecha_parsed_example, (pd.Timestamp, datetime.datetime)):
            dt = fecha_parsed_example
        else:
            try:
                dt = pd.Timestamp(fecha)
            except Exception:
                return None
        dia_idx = dt.weekday()
        dia = dias_semana[dia_idx]

        if sector in self.horarios:
            sec_entry = self.horarios[sector]
            if isinstance(sec_entry, dict):
                # Prioridad persona
                if nombre in sec_entry:
                    persona_horas = sec_entry[nombre]
                    if isinstance(persona_horas, dict) and dia in persona_horas:
                        return persona_horas[dia]
                    elif isinstance(persona_horas, str):
                        return persona_horas
                # Luego horario del sector
                if "__sector__" in sec_entry:
                    sector_horas = sec_entry["__sector__"]
                    if isinstance(sector_horas, dict) and dia in sector_horas:
                        return sector_horas[dia]
                    elif isinstance(sector_horas, str):
                        return sector_horas
            elif isinstance(sec_entry, str):
                return sec_entry
        return None

    # -------------------------
    # Filtro: primer fichaje del dia (vista)
    # -------------------------
    def _filter_primer_fichaje(self, df):
        if df.empty:
            return df
        df2 = df.copy()
        df2["_sort_dt"] = df2["FechaHora_parsed"].apply(
            lambda x: pd.Timestamp.max if pd.isna(x) else x
        )
        df2 = df2.sort_values(["Nombre", "_sort_dt"], ascending=[True, True])
        df2["_Fecha_group"] = df2["Fecha"].astype(str)
        firsts = df2.groupby(["Nombre", "_Fecha_group"], as_index=False).first()
        firsts = firsts.drop(columns=["_sort_dt", "_Fecha_group"], errors="ignore")
        return firsts.reset_index(drop=True)

    # -------------------------
    # Filtro: retrasos por fichaje (vista)
    # -------------------------
    def _filter_retrasos(self, df):
        if df.empty:
            return df
        df2 = df.copy()
        df2["Delay"] = pd.NaT
        df2["Horario_Referencia"] = ""
        dias_semana = ["lunes", "martes", "miércoles", "jueves", "viernes", "sábado", "domingo"]

        def get_ref_time(row):
            sec = row.get("Sector", "")
            nombre = row.get("Nombre", "")
            fecha = row.get("Fecha")
            fecha_parsed = row.get("FechaHora_parsed")
            return self._get_ref_time_for(nombre, sec, fecha, fecha_parsed)

        for idx, row in df2.iterrows():
            ref = get_ref_time(row)
            df2.at[idx, "Horario_Referencia"] = ref if ref else ""
            if not ref:
                df2.at[idx, "Delay"] = pd.NaT
                continue

            hora_real = None
            if pd.notna(row.get("FechaHora_parsed")):
                hora_real = row["FechaHora_parsed"].time()
            else:
                ht = row.get("Hora_text", "")
                if ht:
                    try:
                        dt = parser.parse(ht, dayfirst=True, fuzzy=True)
                        hora_real = dt.time()
                    except Exception:
                        hora_real = None

            try:
                ref_time = parser.parse(ref, default=pd.Timestamp("1900-01-01")).time()
            except Exception:
                try:
                    ref_time = pd.to_datetime(ref, format="%H:%M").time()
                except Exception:
                    ref_time = None

            if hora_real and ref_time:
                dt_real = pd.Timestamp.combine(pd.Timestamp("1900-01-01").date(), hora_real)
                dt_ref = pd.Timestamp.combine(pd.Timestamp("1900-01-01").date(), ref_time)
                delay = dt_real - dt_ref
                df2.at[idx, "Delay"] = delay
            else:
                df2.at[idx, "Delay"] = pd.NaT

        return df2

    # -------------------------
    # Filtro: eliminar fechas específicas
    # -------------------------
    def _filter_eliminar_fechas(self, df):
        if df.empty or not self.fechas_a_eliminar:
            return df
        fechas_set = set(self.fechas_a_eliminar)
        return df[~df["Fecha"].astype(str).isin(fechas_set)]

    # -------------------------
    # APLICAR FILTROS (solo para la vista)
    # -------------------------
    def aplicar_filtros(self):
        if self.df.empty:
            QMessageBox.information(self, "Info", "Primero cargá datos.")
            return
        df_work = self.df.copy()

        if self.chk_primero.isChecked():
            df_work = self._filter_primer_fichaje(df_work)

        if self.chk_retrasos.isChecked():
            df_work = self._filter_retrasos(df_work)

        df_work = self._filter_eliminar_fechas(df_work)

        selected_sectors = [item.text() for item in self.list_sectores_filtro.selectedItems()]
        if selected_sectors:
            df_work = df_work[df_work["Sector"].isin(selected_sectors)]

        self._actualizar_tabla(df_work)

    # -------------------------
    # RESUMEN DIARIO PARA PRESENTISMO
    # -------------------------
    def _calcular_presentismo_diario(self, df_base):
        """
        A partir de todas las fichadas, genera un DF a nivel día/persona:
        - Entrada
        - Salida
        - Delay (sólo de la entrada)
        - EstadoDia (OK, FALTA SALIDA, INASISTENCIA, etc.)
        - EsInasistencia (bool)
        Regla clave:
        - Si el PRIMER fichaje del día tiene más de 4 horas de diferencia con el horario
          de entrada configurado, se toma como SALIDA y se marca INASISTENCIA (falta entrada).
        """
        if df_base.empty:
            return pd.DataFrame()

        registros = []
        for (nombre, fecha), grupo in df_base.groupby(["Nombre", "Fecha"], dropna=False):
            if pd.isna(fecha):
                continue

            grupo_ord = grupo.sort_values("FechaHora_parsed")
            # sector: primer sector no vacío
            sector = "No asignado"
            for sec_val in grupo_ord["Sector"]:
                if isinstance(sec_val, str) and sec_val.strip():
                    sector = sec_val
                    break

            # timestamp de ejemplo para weekday
            ejemplo_dt = None
            for v in grupo_ord["FechaHora_parsed"]:
                if pd.notna(v):
                    ejemplo_dt = v
                    break

            ref_str = self._get_ref_time_for(nombre, sector, fecha, ejemplo_dt)

            # lista de horas reales
            times = []
            for v in grupo_ord["FechaHora_parsed"]:
                if pd.notna(v):
                    times.append(v.time())

            entrada_time = None
            salida_time = None
            delay_td = pd.NaT
            estado = ""
            es_inasistencia = False

            if not times:
                # Sin fichaje parseable: ausencia total (si aparece, es raro)
                estado = "SIN FICHAJE VÁLIDO"
                es_inasistencia = True
            else:
                # Intentamos parsear horario ref
                if ref_str:
                    try:
                        ref_time = parser.parse(ref_str, default=pd.Timestamp("1900-01-01")).time()
                    except Exception:
                        try:
                            ref_time = pd.to_datetime(ref_str, format="%H:%M").time()
                        except Exception:
                            ref_time = None
                else:
                    ref_time = None

                if ref_time:
                    earliest = times[0]
                    dt_ref = pd.Timestamp.combine(pd.Timestamp("1900-01-01"), ref_time)
                    dt_ear = pd.Timestamp.combine(pd.Timestamp("1900-01-01"), earliest)
                    diff = dt_ear - dt_ref

                    # REGLA DE LAS 4 HORAS:
                    # si la primera fichada está > 4h después del horario de entrada,
                    # se considera que NO fichó entrada (solo salida) => inasistencia.
                    if diff > pd.Timedelta(hours=4):
                        entrada_time = None
                        # Por simplicidad tomamos la primera fichada como salida declarada
                        salida_time = earliest
                        estado = "INASISTENCIA (solo salida, falta entrada)"
                        es_inasistencia = True
                        delay_td = pd.NaT
                    else:
                        # Consideramos earliest como ENTRADA
                        entrada_time = earliest
                        if len(times) > 1:
                            salida_time = times[-1]
                            estado = "OK"
                        else:
                            salida_time = None
                            estado = "FALTA SALIDA (solo entrada)"

                        # Delay sólo de la ENTRADA
                        delay_td = dt_ear - dt_ref
                else:
                    # No hay horario de referencia → no podemos evaluar la regla de 4h
                    earliest = times[0]
                    entrada_time = earliest
                    if len(times) > 1:
                        salida_time = times[-1]
                        estado = "OK (sin horario ref)"
                    else:
                        salida_time = None
                        estado = "FALTA SALIDA (sin horario ref)"
                    delay_td = pd.NaT

            registros.append({
                "Nombre": nombre,
                "Fecha": fecha,
                "Sector": sector,
                "Entrada": entrada_time.strftime("%H:%M:%S") if entrada_time else "",
                "Salida": salida_time.strftime("%H:%M:%S") if salida_time else "",
                "Horario_Referencia": ref_str or "",
                "Delay": delay_td,
                "EstadoDia": estado,
                "EsInasistencia": es_inasistencia,
            })

        df_dias = pd.DataFrame(registros)
        return df_dias

    # -------------------------
    # Exportar vista actual (tal cual se ve)
    # -------------------------
    def exportar_vista_actual(self):
        if self.display_df.empty:
            QMessageBox.information(self, "Info", "No hay datos para exportar.")
            return
        path, _ = QFileDialog.getSaveFileName(
            self, "Guardar vista actual (Excel)", "", "Archivos Excel (*.xlsx)"
        )
        if not path:
            return
        try:
            self.display_df.to_excel(path, index=False)
            QMessageBox.information(self, "Guardado", f"Vista guardada en:\n{path}")
        except Exception as e:
            QMessageBox.critical(self, "Error guardado", str(e))

    # -------------------------
    # EXPORTAR RESULTADO (Presentismo)
    # -------------------------
    def exportar_resultado(self):
        if self.df.empty:
            QMessageBox.information(self, "Info", "No hay datos procesados para exportar.")
            return

        # Tomamos SIEMPRE todos los registros crudos, pero:
        # - filtramos fechas a eliminar
        # - filtramos por sectores seleccionados (si los hay)
        df_base = self.df.copy()
        df_base = self._filter_eliminar_fechas(df_base)

        selected_sectors = [item.text() for item in self.list_sectores_filtro.selectedItems()]
        if selected_sectors:
            df_base = df_base[df_base["Sector"].isin(selected_sectors)]

        if df_base.empty:
            QMessageBox.information(self, "Info", "No hay datos para exportar después de filtros.")
            return

        # Calcular resumen diario con la lógica de entrada/salida/inasistencia
        df_daily = self._calcular_presentismo_diario(df_base)
        if df_daily.empty:
            QMessageBox.information(self, "Info", "No se generaron registros diarios.")
            return

        path, _ = QFileDialog.getSaveFileName(
            self, "Guardar resultado (Excel)", "", "Archivos Excel (*.xlsx)"
        )
        if not path:
            return

        try:
            # Título del informe
            fecha_max = df_daily['Fecha'].max()
            if pd.notna(fecha_max):
                try:
                    fecha_max_ts = pd.to_datetime(fecha_max)
                    month = fecha_max_ts.strftime('%B').upper()
                    year = fecha_max_ts.year
                except Exception:
                    month = 'MES'
                    year = ''
            else:
                month = 'MES'
                year = ''

            sectores_unicos = df_daily['Sector'].dropna().unique()
            if len(sectores_unicos) == 1 and sectores_unicos[0] != "No asignado":
                sector = sectores_unicos[0].upper()
            else:
                sector = 'TODOS'
            title = f"SECTOR {sector} {month} {year}".strip()

            with pd.ExcelWriter(path, engine='openpyxl') as writer:
                # ==============================
                # Hoja 1: Resumen de Presentismo
                # ==============================
                all_employees = set()
                for sec in sectores_unicos:
                    if sec in self.sectores:
                        all_employees.update(self.sectores[sec])

                resumen_data = []
                for emp in sorted(all_employees):
                    emp_data = df_daily[df_daily["Nombre"] == emp]

                    if emp_data.empty:
                            # No tiene ningún fichaje en el período (puede ser vacaciones/licencia, etc.)
                            resumen_data.append({
                            "Nombre": emp,
                            "Dias_tarde_total": 0,
                            "Total_tiempo_tarde": "00:00:00",
                            "Cobrar Presentismo": "",
                            "Aclaraciones": "SIN REGISTROS DE FICHAJES EN EL PERÍODO (vacaciones/licencia o sin fichar)."
                            })
                            continue


                    # Días con retraso (Delay >= 1 minuto)
                    all_positive_delays = emp_data[emp_data["Delay"] >= pd.Timedelta(minutes=1)]

                    # Normales (<= 3 horas)
                    normal_delays = all_positive_delays[
                        all_positive_delays["Delay"] <= pd.Timedelta(hours=3)
                    ]

                    # Extremos (> 3 horas)
                    extreme_delays = all_positive_delays[
                        all_positive_delays["Delay"] > pd.Timedelta(hours=3)
                    ]
                    num_extreme = len(extreme_delays)

                    # Días marcados como inasistencia (por falta de entrada, etc.)
                    dias_inasistencia = emp_data[emp_data["EsInasistencia"] == True]
                    num_inasist = len(dias_inasistencia)

                    # Días totales que afectan presentismo
                    total_late_days = len(normal_delays) + num_extreme + num_inasist

                    if len(normal_delays) == 0 and num_extreme == 0 and num_inasist == 0:
                        row = {
                            "Dias_tarde_total": 0,
                            "Dias_tarde_menor_igual_15min": 0,
                            "Dias_tarde_16_60": 0,
                            "Dias_tarde_mayor_60": 0,
                            "Total_tiempo_tarde": pd.Timedelta(0),
                        }
                        cobrar = "Si 100%"
                        aclaracion = "Nunca llegó tarde ni tuvo inasistencias"
                    else:
                        if normal_delays.empty:
                            row = {
                                "Dias_tarde_total": total_late_days,
                                "Dias_tarde_menor_igual_15min": 0,
                                "Dias_tarde_16_60": 0,
                                "Dias_tarde_mayor_60": 0,
                                "Total_tiempo_tarde": pd.Timedelta(0),
                            }
                        else:
                            group = normal_delays
                            row = {
                                "Dias_tarde_total": total_late_days,
                                "Dias_tarde_menor_igual_15min": (
                                    group["Delay"] <= pd.Timedelta(minutes=15)
                                ).sum(),
                                "Dias_tarde_16_60": (
                                    (group["Delay"] > pd.Timedelta(minutes=15))
                                    & (group["Delay"] <= pd.Timedelta(minutes=60))
                                ).sum(),
                                "Dias_tarde_mayor_60": (
                                    group["Delay"] > pd.Timedelta(minutes=60)
                                ).sum(),
                                "Total_tiempo_tarde": group["Delay"].sum(),
                            }

                    delays_16_60 = row.get("Dias_tarde_16_60", 0)
                    delays_over_60 = row.get("Dias_tarde_mayor_60", 0)
                    num_delays_over_15 = delays_16_60 + delays_over_60 + num_extreme

                    # Lógica de cobro de presentismo
                    if total_late_days == 0:
                        cobrar = "Si 100%"
                        aclaracion = "Nunca llegó tarde ni tuvo inasistencias"
                    elif total_late_days == 1:
                        if num_delays_over_15 == 0 and num_inasist == 0:
                            cobrar = "Si 100%"
                            aclaracion = "1 llegada tarde ≤ 15 min"
                        elif num_inasist > 0:
                            cobrar = "Si 50%"
                            aclaracion = "1 inasistencia/falta de ingreso"
                        else:
                            cobrar = "Si 100%"
                            aclaracion = "1 llegada tarde > 15 min"
                    elif total_late_days == 2:
                        if num_inasist >= 2:
                            cobrar = "No"
                            aclaracion = "2 inasistencias/faltas de ingreso"
                        elif num_inasist == 1:
                            if num_delays_over_15 == 0:
                                cobrar = "Si 50%"
                                aclaracion = "1 inasistencia y 1 llegada tarde ≤15 min"
                            else:
                                cobrar = "No"
                                aclaracion = "1 inasistencia y 1 llegada tarde >15 min"
                        else:
                            if num_delays_over_15 == 0:
                                cobrar = "Si 100%"
                                aclaracion = "2 llegadas tarde, todas ≤15 min"
                            elif num_delays_over_15 == 1:
                                cobrar = "Si 100%"
                                aclaracion = "2 llegadas tarde, una >15 min y una ≤15 min"
                            else:
                                cobrar = "Si 50%"
                                aclaracion = "2 llegadas tarde, ambas >15 min"
                    else:
                        cobrar = "No"
                        aclaracion = f"{total_late_days} días entre tardanzas e inasistencias"
                        if num_inasist > 0:
                            aclaracion += f" (incluye {num_inasist} inasistencia(s))."

                    row["Cobrar Presentismo"] = cobrar
                    row["Aclaraciones"] = aclaracion

                    total_td = row["Total_tiempo_tarde"]
                    if isinstance(total_td, pd.Timedelta):
                        total_str = (
                            f"{int(total_td.total_seconds() // 3600):02d}:"
                            f"{int((total_td.total_seconds() % 3600) // 60):02d}:"
                            f"{int(total_td.total_seconds() % 60):02d}"
                        )
                    else:
                        total_str = "00:00:00"
                    row["Total_tiempo_tarde"] = total_str
                    row["Nombre"] = emp
                    resumen_data.append(row)

                df_resumen = pd.DataFrame(resumen_data)
                df_resumen = df_resumen[["Nombre", "Dias_tarde_total", "Total_tiempo_tarde",
                                         "Cobrar Presentismo", "Aclaraciones"]]

                if not df_resumen.empty:
                    empleados_con_datos = df_resumen[df_resumen["Dias_tarde_total"].astype(int) > 0]
                    if not empleados_con_datos.empty:
                        total_dias_tarde = empleados_con_datos["Dias_tarde_total"].astype(int).sum()
                        tiempos_validos = empleados_con_datos[
                            empleados_con_datos["Total_tiempo_tarde"] != "00:00:00"
                        ]
                        if not tiempos_validos.empty:
                            total_tiempo_tarde = pd.to_timedelta(
                                tiempos_validos["Total_tiempo_tarde"].apply(lambda x: f"0 days {x}")
                            ).sum()
                            total_tiempo_str = (
                                f"{int(total_tiempo_tarde.total_seconds() // 3600):02d}:"
                                f"{int((total_tiempo_tarde.total_seconds() % 3600) // 60):02d}:"
                                f"{int(total_tiempo_tarde.total_seconds() % 60):02d}"
                            )
                        else:
                            total_tiempo_str = "00:00:00"
                        fila_total = pd.DataFrame([{
                            "Nombre": "TOTAL",
                            "Dias_tarde_total": str(total_dias_tarde),
                            "Total_tiempo_tarde": total_tiempo_str,
                            "Cobrar Presentismo": "",
                            "Aclaraciones": ""
                        }])
                        df_resumen = pd.concat([df_resumen, fila_total], ignore_index=True)

                df_resumen = df_resumen.astype(str)
                pd.DataFrame([title]).to_excel(
                    writer, sheet_name='Resumen Retrasos', index=False, header=False
                )
                df_resumen.to_excel(
                    writer, sheet_name='Resumen Retrasos', index=False, startrow=1
                )

                # ==============================
                # Hoja 2: Días tarde / inasistencias (detalle)
                # ==============================
                df_detalle = df_daily.copy()

                # Filtrar:
                # - días con Delay >= 1 minuto (tardanzas)
                # - o días marcados como inasistencia
                mask_tarde = df_detalle["Delay"] >= pd.Timedelta(minutes=1)
                mask_inasist = df_detalle["EsInasistencia"] == True
                df_detalle = df_detalle[mask_tarde | mask_inasist]

                if not df_detalle.empty:
                    def minutos_tarde(row):
                        d = row["Delay"]
                        if pd.notna(d) and d != pd.NaT and d.total_seconds() > 0:
                            return int(d.total_seconds() // 60)
                        return 0

                    df_detalle["Minutos Tarde"] = df_detalle.apply(minutos_tarde, axis=1)
                    df_detalle["Delay"] = df_detalle["Delay"].apply(
                        lambda x: "" if pd.isna(x) or x == pd.NaT
                        else f"{int(x.total_seconds() // 3600):02d}:"
                             f"{int((x.total_seconds() % 3600) // 60):02d}:"
                             f"{int(x.total_seconds() % 60):02d}"
                    )
                    df_detalle = df_detalle[[
                        "Nombre", "Fecha", "Sector", "Entrada", "Salida",
                        "Horario_Referencia", "Delay", "Minutos Tarde", "EstadoDia"
                    ]]
                    df_title_detalle = pd.DataFrame([title])
                    df_headers_detalle = pd.DataFrame([df_detalle.columns.tolist()])
                    df_combined_detalle = pd.concat(
                        [df_title_detalle, df_headers_detalle, df_detalle], ignore_index=True
                    )
                    df_combined_detalle.to_excel(
                        writer, sheet_name='Dias Tarde', index=False, header=False
                    )
                else:
                    df_title_detalle = pd.DataFrame([title])
                    df_message_detalle = pd.DataFrame({"Mensaje": ["No hay registros de días tarde/inasistencias"]})
                    df_combined_detalle = pd.concat(
                        [df_title_detalle, df_message_detalle], ignore_index=True
                    )
                    df_combined_detalle.to_excel(
                        writer, sheet_name='Dias Tarde', index=False, header=False
                    )

                # ==============================
                # Hoja 3: Datos Procesados (resumen diario por persona)
                # ==============================
                df_procesados = df_daily.copy()
                df_title_procesados = pd.DataFrame([title])
                df_headers_procesados = pd.DataFrame([df_procesados.columns.tolist()])
                df_combined_procesados = pd.concat(
                    [df_title_procesados, df_headers_procesados, df_procesados],
                    ignore_index=True
                )
                df_combined_procesados.to_excel(
                    writer, sheet_name='Datos Procesados', index=False, header=False
                )

            QMessageBox.information(
                self,
                "Guardado",
                f"Archivo guardado en:\n{path}\n"
                f"Incluye hojas 'Resumen Retrasos', 'Dias Tarde' y 'Datos Procesados' "
                f"con título '{title}'."
            )
        except Exception as e:
            QMessageBox.critical(self, "Error guardado", str(e))

    # -------------------------
    # Exportar no-asignados
    # -------------------------
    def exportar_no_asignados_csv(self):
        sin = self.df[self.df["Sector"] == "No asignado"]
        if sin.empty:
            QMessageBox.information(self, "Información", "No hay registros sin sector.")
            return
        path, _ = QFileDialog.getSaveFileName(
            self, "Exportar no-asignados (CSV)", "", "Archivos CSV (*.csv)"
        )
        if not path:
            return
        try:
            sin.to_csv(path, index=False, encoding="utf-8-sig")
            QMessageBox.information(self, "Exportado", f"No-asignados exportados a:\n{path}")
        except Exception as e:
            QMessageBox.critical(self, "Error export", str(e))
