# Correcci-n-Orbital-Integrada---Versi-n-Profesional-v3.2.py
## ⚡ EJECUCIÓN  1. Copiar el archivo `Corrección Orbital Integrada - Versión Profesional v3.2.py` 2. Guardar con extensión `.py` 3. Instalar dependencias: `pip install -r requirements.txt` 4. Ejecutar: `python "Corrección Orbital Integrada - Versión Profesional v3.2.py"`

"""
Modelo de Corrección Orbital Integrada - Versión Profesional v3.2
π_corr = π · (1 + Σδ_i) / (1 + ΣΔ_j)
Variables integradas: δ_masa, δ_solar, δ_eclipse
Exportación: TXT, PDF, DOCX
Botón Estimar δ corregido con ventana de confirmación
"""

import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import math
import json
import os
from datetime import datetime
import zipfile

# Constantes
RADIO_TIERRA = 6371.0
VELOCIDAD_ORBITAL_LEO = 7.7
PI = math.pi

BASE_MISIONES = {
    "GOCE (260 km)": {
        "deltas": {"δ_drag": 0.0015, "δ_SRP": 0.0003, "δ_therm": 0.0002,
                   "δ_att": 0.0002, "δ_mag": 0.0001, "δ_model": 0.0001,
                   "δ_masa": 0.0, "δ_solar": 0.0, "δ_eclipse": 0.0},
        "deltas_mayus": {"Δ_geo": -0.0004, "Δ_shape": 0.0001, "Δ_bias": 0.0000},
        "fisicos": {"altitud": 260.0, "radio_tierra": 6371.0, "velocidad_orbital": 7.8,
                    "masa_satelite": 1077.0, "area_efectiva": 1.0, "coeficiente_arrastre": 2.2,
                    "densidad_atmosferica": 8e-12, "presion_radiacion": 4.56e-6,
                    "factor_reflectividad": 1.2, "fraccion_solar": 0.6,
                    "masa_inicial_kg": 1077.0, "masa_final_kg": 872.0,
                    "indice_f107": 70.0, "f107_referencia": 70.0,
                    "inclinacion_orbital": 96.7, "raan_inicial": 0.0},
        "consumo_diario_kg": 0.001, "costo_mision_millones": 350, "duracion_mision_anos": 4
    },
    "CHAMP (454 km)": {
        "deltas": {"δ_drag": 0.0010, "δ_SRP": 0.0003, "δ_therm": 0.0002,
                   "δ_att": 0.0001, "δ_mag": 0.0002, "δ_model": 0.0001,
                   "δ_masa": 0.0, "δ_solar": 0.0, "δ_eclipse": 0.0},
        "deltas_mayus": {"Δ_geo": -0.0003, "Δ_shape": 0.0001, "Δ_bias": 0.0000},
        "fisicos": {"altitud": 454.0, "radio_tierra": 6371.0, "velocidad_orbital": 7.63,
                    "masa_satelite": 522.0, "area_efectiva": 1.8, "coeficiente_arrastre": 2.3,
                    "densidad_atmosferica": 1.2e-12, "presion_radiacion": 4.56e-6,
                    "factor_reflectividad": 1.2, "fraccion_solar": 0.6,
                    "masa_inicial_kg": 522.0, "masa_final_kg": 480.0,
                    "indice_f107": 100.0, "f107_referencia": 100.0,
                    "inclinacion_orbital": 87.3, "raan_inicial": 0.0},
        "consumo_diario_kg": 0.0002, "costo_mision_millones": 150, "duracion_mision_anos": 5
    },
    "GRACE (500 km)": {
        "deltas": {"δ_drag": 0.0008, "δ_SRP": 0.0002, "δ_therm": 0.0001,
                   "δ_att": 0.0002, "δ_mag": 0.0001, "δ_model": 0.0001,
                   "δ_masa": 0.0, "δ_solar": 0.0, "δ_eclipse": 0.0},
        "deltas_mayus": {"Δ_geo": -0.0003, "Δ_shape": 0.0001, "Δ_bias": 0.0000},
        "fisicos": {"altitud": 500.0, "radio_tierra": 6371.0, "velocidad_orbital": 7.6,
                    "masa_satelite": 487.0, "area_efectiva": 1.5, "coeficiente_arrastre": 2.2,
                    "densidad_atmosferica": 1e-12, "presion_radiacion": 4.56e-6,
                    "factor_reflectividad": 1.3, "fraccion_solar": 0.6,
                    "masa_inicial_kg": 487.0, "masa_final_kg": 432.0,
                    "indice_f107": 100.0, "f107_referencia": 100.0,
                    "inclinacion_orbital": 89.0, "raan_inicial": 0.0},
        "consumo_diario_kg": 0.0001, "costo_mision_millones": 200, "duracion_mision_anos": 5
    },
    "Sentinel-1 (693 km)": {
        "deltas": {"δ_drag": 0.0003, "δ_SRP": 0.0006, "δ_therm": 0.0002,
                   "δ_att": 0.0002, "δ_mag": 0.0001, "δ_model": 0.0001,
                   "δ_masa": 0.0, "δ_solar": 0.0, "δ_eclipse": 0.0001},
        "deltas_mayus": {"Δ_geo": -0.0004, "Δ_shape": 0.0001, "Δ_bias": 0.0000},
        "fisicos": {"altitud": 693.0, "radio_tierra": 6371.0, "velocidad_orbital": 7.5,
                    "masa_satelite": 2300.0, "area_efectiva": 20.0, "coeficiente_arrastre": 2.5,
                    "densidad_atmosferica": 1e-13, "presion_radiacion": 4.56e-6,
                    "factor_reflectividad": 1.3, "fraccion_solar": 0.7,
                    "masa_inicial_kg": 2300.0, "masa_final_kg": 2200.0,
                    "indice_f107": 80.0, "f107_referencia": 80.0,
                    "inclinacion_orbital": 98.2, "raan_inicial": 0.0},
        "consumo_diario_kg": 0.0001, "costo_mision_millones": 280, "duracion_mision_anos": 7
    }
}


class CorrectorOrbitalPro:
    """Motor de cálculo v3.2"""
    
    def __init__(self):
        self.pi = PI
        self.variables_delta = {}
        self.variables_delta_mayus = {}
        self.parametros_fisicos = {}
        self.consumo_diario_kg = 0.001
        self.costo_mision_millones = 0
        self.duracion_mision_anos = 1
        self._inicializar()
    
    def _inicializar(self):
        self.variables_delta = {
            "δ_drag": {"nombre": "Arrastre atmosférico", "valor": 0.0},
            "δ_SRP": {"nombre": "Presión de radiación solar", "valor": 0.0},
            "δ_therm": {"nombre": "Efecto termoelástico", "valor": 0.0},
            "δ_att": {"nombre": "Variaciones por actitud", "valor": 0.0},
            "δ_mag": {"nombre": "Perturbaciones magnéticas", "valor": 0.0},
            "δ_model": {"nombre": "Sesgo residual del modelo", "valor": 0.0},
            "δ_masa": {"nombre": "Variación de masa (consumo)", "valor": 0.0},
            "δ_solar": {"nombre": "Actividad solar (F10.7)", "valor": 0.0},
            "δ_eclipse": {"nombre": "Fracción real de eclipse", "valor": 0.0}
        }
        self.variables_delta_mayus = {
            "Δ_geo": {"nombre": "Corrección geométrica J2", "valor": 0.0},
            "Δ_shape": {"nombre": "Cambios estructurales", "valor": 0.0},
            "Δ_bias": {"nombre": "Sesgo geométrico residual", "valor": 0.0}
        }
        self.parametros_fisicos = {
            "altitud": 420.0, "radio_tierra": RADIO_TIERRA, "velocidad_orbital": VELOCIDAD_ORBITAL_LEO,
            "masa_satelite": 300.0, "area_efectiva": 6.0, "coeficiente_arrastre": 2.2,
            "densidad_atmosferica": 4e-12, "presion_radiacion": 4.56e-6,
            "factor_reflectividad": 1.2, "fraccion_solar": 0.6,
            "masa_inicial_kg": 300.0, "masa_final_kg": 280.0,
            "indice_f107": 70.0, "f107_referencia": 70.0,
            "inclinacion_orbital": 90.0, "raan_inicial": 0.0
        }
    
    def cargar_mision(self, nombre):
        if nombre not in BASE_MISIONES:
            return False
        d = BASE_MISIONES[nombre]
        for k, v in d["deltas"].items():
            if k in self.variables_delta:
                self.variables_delta[k]["valor"] = v
        for k, v in d["deltas_mayus"].items():
            if k in self.variables_delta_mayus:
                self.variables_delta_mayus[k]["valor"] = v
        for k, v in d["fisicos"].items():
            self.parametros_fisicos[k] = v
        self.consumo_diario_kg = d.get("consumo_diario_kg", 0.001)
        self.costo_mision_millones = d.get("costo_mision_millones", 0)
        self.duracion_mision_anos = d.get("duracion_mision_anos", 1)
        return True
    
    def estimar_delta_masa(self):
        """Calcula delta_masa. Se activa cuando masa_actual < masa_inicial."""
        m_inicial = self.parametros_fisicos.get("masa_inicial_kg", 0.0)
        m_actual = self.parametros_fisicos.get("masa_satelite", 0.0)
        
        if m_inicial <= 0 or m_actual <= 0:
            return 0.0
        if m_actual >= m_inicial:
            return 0.0
        
        fraccion = (m_inicial - m_actual) / m_inicial
        return fraccion * 0.01
    
    def estimar_delta_solar(self):
        p = self.parametros_fisicos
        f107 = p.get("indice_f107", 70.0)
        ref = p.get("f107_referencia", 70.0)
        if ref > 0:
            return ((f107 / ref) - 1) * 0.0005
        return 0.0
    
    def estimar_delta_eclipse(self):
        p = self.parametros_fisicos
        real = p.get("fraccion_solar", 0.6)
        media = 0.6
        return (real - media) * 0.001
    
    def calcular_factor(self):
        sd = sum(v["valor"] for v in self.variables_delta.values())
        sD = sum(v["valor"] for v in self.variables_delta_mayus.values())
        den = 1 + sD
        if den == 0:
            return None, sd, sD
        return (1 + sd) / den, sd, sD
    
    def calcular_pi_corregido(self):
        res = self.calcular_factor()
        if res is None or res[0] is None:
            return None
        return self.pi * res[0]
    
    def calcular_ahorro(self, dias_por_ano=365):
        res = self.calcular_factor()
        if res is None or res[0] is None:
            return None
        _, sd, sD = res
        magnitud = abs(sd) + abs(sD)
        ref = 0.0024
        escala = magnitud / ref if ref != 0 else 0
        
        amin = self.consumo_diario_kg * 0.10 * escala
        amax = self.consumo_diario_kg * 0.20 * escala
        inc = 0.30
        amin_c = amin * (1 - inc)
        amax_c = amax * (1 + inc)
        aanual_min = amin_c * dias_por_ano
        aanual_max = amax_c * dias_por_ano
        
        dmin = int(aanual_min / max(self.consumo_diario_kg, 0.0001))
        dmax = int(aanual_max / max(self.consumo_diario_kg, 0.0001))
        
        cd = 0
        vmin = 0
        vmax = 0
        if self.costo_mision_millones > 0 and self.duracion_mision_anos > 0:
            cd = self.costo_mision_millones / (self.duracion_mision_anos * 365)
            vmin = dmin * cd
            vmax = dmax * cd
        
        return {"ahorro_diario_min": amin_c, "ahorro_diario_max": amax_c,
                "ahorro_anual_min": aanual_min, "ahorro_anual_max": aanual_max,
                "dias_min": dmin, "dias_max": dmax, "magnitud": magnitud,
                "inc_pct": inc*100, "costo_diario": cd, "valor_min": vmin, "valor_max": vmax}
    
    def guardar_configuracion(self, archivo):
        pi_corr = self.calcular_pi_corregido()
        ahorro = self.calcular_ahorro()
        res = self.calcular_factor()
        datos = {
            "software": "Modelo de Corrección Orbital Integrada v3.2",
            "fecha": datetime.now().isoformat(),
            "deltas": {k: v["valor"] for k, v in self.variables_delta.items()},
            "deltas_mayus": {k: v["valor"] for k, v in self.variables_delta_mayus.items()},
            "parametros_fisicos": self.parametros_fisicos,
            "consumo_diario_kg": self.consumo_diario_kg,
            "resultados": {
                "pi_tradicional": self.pi,
                "pi_corregido": pi_corr,
                "factor_correccion": res[0] if res else None,
                "suma_deltas": res[1] if res else None,
                "suma_deltas_mayus": res[2] if res else None,
                "ahorro_propelente": ahorro
            }
        }
        with open(archivo, 'w', encoding='utf-8') as f:
            json.dump(datos, f, indent=2, ensure_ascii=False)
        return True
    
    def cargar_configuracion(self, archivo):
        with open(archivo, 'r', encoding='utf-8') as f:
            datos = json.load(f)
        for k, v in datos.get("deltas", {}).items():
            if k in self.variables_delta:
                self.variables_delta[k]["valor"] = v
        for k, v in datos.get("deltas_mayus", {}).items():
            if k in self.variables_delta_mayus:
                self.variables_delta_mayus[k]["valor"] = v
        for k, v in datos.get("parametros_fisicos", {}).items():
            if k in self.parametros_fisicos:
                self.parametros_fisicos[k] = v
        self.consumo_diario_kg = datos.get("consumo_diario_kg", 0.001)
        return True


class InterfazMejorada:
    """Interfaz gráfica profesional v3.2"""
    
    def __init__(self, ventana):
        self.ventana = ventana
        self.ventana.title("Modelo de Corrección Orbital Integrada v3.2")
        self.ventana.configure(bg="#0d1117")
        
        sw = self.ventana.winfo_screenwidth()
        sh = self.ventana.winfo_screenheight()
        self.ventana.geometry(f"{int(sw*0.8)}x{int(sh*0.8)}")
        self.ventana.minsize(850, 550)
        
        self.modelo = CorrectorOrbitalPro()
        self.vars_e = {}
        self.vars_f = {}
        self.vars_i = {}
        
        self.c = {
            "bg": "#0d1117", "bg2": "#161b22", "bg3": "#1c2333",
            "btn": "#2d3748", "btn_ok": "#238636", "btn_info": "#1f6feb",
            "btn_warn": "#9e6a03", "borde": "#30363d",
            "txt_tit": "#58a6ff", "txt_sub": "#c9d1d9", "txt_dim": "#8b949e",
            "verde": "#3fb950", "rojo": "#f85149", "ambar": "#d2991d",
            "entrada": "#0d1117"
        }
        
        self._construir()
        self._aplicar_estilos_ttk()
    
    def _aplicar_estilos_ttk(self):
        style = ttk.Style()
        style.theme_use('clam')
        style.configure("TNotebook", background=self.c["bg"], borderwidth=0)
        style.configure("TNotebook.Tab", background=self.c["btn"], foreground=self.c["txt_sub"],
                        padding=[12, 4], font=("Segoe UI", 9, "bold"))
        style.map("TNotebook.Tab", background=[("selected", self.c["btn_ok"])],
                  foreground=[("selected", "white")])
    
    def _construir(self):
        c = self.c
        main = tk.Frame(self.ventana, bg=c["bg"])
        main.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        hdr = tk.Frame(main, bg=c["bg"])
        hdr.pack(fill=tk.X, pady=(0, 8))
        tk.Label(hdr, text="Modelo de Corrección Orbital Integrada v3.2", font=("Segoe UI", 15, "bold"),
                 fg=c["txt_tit"], bg=c["bg"]).pack(anchor=tk.W)
        tk.Label(hdr, text="π_corr = π · (1 + Σδ_i) / (1 + ΣΔ_j)  |  δ_masa · δ_solar · δ_eclipse",
                 font=("Consolas", 9), fg=c["ambar"], bg=c["bg"]).pack(anchor=tk.W)
        
        self.nb = ttk.Notebook(main)
        self.nb.pack(fill=tk.BOTH, expand=True, pady=4)
        
        self._pestana_deltas()
        self._pestana_deltas_mayus()
        self._pestana_fisicos()
        self._pestana_misiones()
        self._pestana_graficos()
        self._pestana_resultados()
        
        inf = tk.Frame(main, bg=c["bg"])
        inf.pack(fill=tk.X, pady=(8, 0))
        
        rap = tk.Frame(inf, bg=c["bg2"], relief=tk.RIDGE, bd=1)
        rap.pack(fill=tk.X, pady=(0, 6), ipady=5)
        self.lbl_res = tk.StringVar(value="π_corr = --")
        tk.Label(rap, textvariable=self.lbl_res, font=("Consolas", 14, "bold"),
                 fg=c["verde"], bg=c["bg2"]).pack()
        self.lbl_info = tk.StringVar(value="v3.2 · Interfaz mejorada · Botón Estimar δ corregido")
        tk.Label(rap, textvariable=self.lbl_info, font=("Segoe UI", 8),
                 fg=c["txt_dim"], bg=c["bg2"]).pack()
        
        btn_frame = tk.Frame(inf, bg=c["bg"])
        btn_frame.pack(fill=tk.X)
        
        botones = [
            ("▶  CALCULAR", self._calcular, c["btn_ok"], 12, "bold"),
            ("📄 PDF", self._exportar_pdf, c["btn_warn"], 9, "normal"),
            ("📝 Word", self._exportar_docx, c["btn_info"], 9, "normal"),
            ("📊 TXT", self._exportar_txt, c["btn"], 9, "normal"),
            ("💾 Guardar", self._guardar, c["btn"], 9, "normal"),
            ("📂 Cargar", self._cargar, c["btn"], 9, "normal"),
            ("🔬 Estimar δ", self._estimar_nuevos, c["btn_info"], 9, "normal"),
            ("🔄 Limpiar", self._limpiar, c["btn"], 9, "normal"),
        ]
        for texto, cmd, bg, size, peso in reversed(botones):
            tk.Button(btn_frame, text=texto, command=cmd, bg=bg, fg="white",
                      font=("Segoe UI", size, peso), padx=10, pady=4, cursor="hand2",
                      relief=tk.FLAT, activebackground=bg).pack(side=tk.RIGHT, padx=2)
        
        self.sts = tk.StringVar(value="Listo · v3.2 · Botón Estimar δ corregido")
        tk.Label(inf, textvariable=self.sts, font=("Segoe UI", 7),
                 fg=c["txt_dim"], bg=c["bg"], anchor=tk.W).pack(fill=tk.X, pady=(6, 0))
    
    def _crear_tabla(self, parent, variables):
        c = self.c
        canvas = tk.Canvas(parent, bg=c["bg"], highlightthickness=0)
        scroll = tk.Scrollbar(parent, orient="vertical", command=canvas.yview)
        frame = tk.Frame(canvas, bg=c["bg"])
        frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=frame, anchor="nw")
        canvas.configure(yscrollcommand=scroll.set)
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scroll.pack(side=tk.RIGHT, fill=tk.Y)
        
        headers = ["Variable", "Nombre", "Valor"]
        widths = [12, 24, 10]
        for j, (h, w) in enumerate(zip(headers, widths)):
            tk.Label(frame, text=h, font=("Segoe UI", 9, "bold"), bg=c["borde"],
                     fg="white", relief=tk.RIDGE, padx=6, pady=4, width=w).grid(row=0, column=j, sticky="nsew", padx=1, pady=1)
        
        for i, (clave, datos) in enumerate(variables.items(), 1):
            bg = c["bg2"] if i % 2 == 0 else c["bg3"]
            es_nueva = clave in ["δ_masa", "δ_solar", "δ_eclipse"]
            color_var = c["ambar"] if es_nueva else c["txt_tit"]
            
            tk.Label(frame, text=clave, font=("Consolas", 9, "bold"), bg=bg,
                     fg=color_var, relief=tk.RIDGE, padx=6, pady=3).grid(row=i, column=0, sticky="nsew", padx=1, pady=1)
            tk.Label(frame, text=datos["nombre"], font=("Segoe UI", 9), bg=bg,
                     fg=c["txt_sub"], relief=tk.RIDGE, padx=6, pady=3).grid(row=i, column=1, sticky="nsew", padx=1, pady=1)
            
            var = tk.StringVar(value=str(datos["valor"]))
            self.vars_e[clave] = var
            tk.Entry(frame, textvariable=var, font=("Consolas", 9), bg=c["entrada"],
                     fg=c["verde"], insertbackground=c["verde"], relief=tk.SUNKEN,
                     width=10, justify=tk.CENTER).grid(row=i, column=2, sticky="nsew", padx=1, pady=1)
        frame.grid_columnconfigure(1, weight=1)
    
    def _pestana_deltas(self):
        p = tk.Frame(self.nb, bg=self.c["bg"])
        self.nb.add(p, text="  Perímetro (δ)  ")
        self._crear_tabla(p, self.modelo.variables_delta)
    
    def _pestana_deltas_mayus(self):
        p = tk.Frame(self.nb, bg=self.c["bg"])
        self.nb.add(p, text="  Diámetro (Δ)  ")
        self._crear_tabla(p, self.modelo.variables_delta_mayus)
    
    def _pestana_fisicos(self):
        c = self.c
        p = tk.Frame(self.nb, bg=c["bg"])
        self.nb.add(p, text="  Parámetros Físicos  ")
        
        canvas = tk.Canvas(p, bg=c["bg"], highlightthickness=0)
        scroll = tk.Scrollbar(p, orient="vertical", command=canvas.yview)
        frame = tk.Frame(canvas, bg=c["bg"])
        frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=frame, anchor="nw")
        canvas.configure(yscrollcommand=scroll.set)
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scroll.pack(side=tk.RIGHT, fill=tk.Y)
        
        self._seccion(frame, "Parámetros Físicos del Satélite", [
            ("Altitud orbital (km):", "altitud", "420.0"),
            ("Radio terrestre (km):", "radio_tierra", "6371.0"),
            ("Velocidad orbital (km/s):", "velocidad_orbital", "7.7"),
            ("Masa actual del satélite (kg):", "masa_satelite", "300.0"),
            ("Área efectiva (m²):", "area_efectiva", "6.0"),
            ("Coeficiente de arrastre (CD):", "coeficiente_arrastre", "2.2"),
            ("Densidad atmosférica (kg/m³):", "densidad_atmosferica", "4e-12"),
            ("Presión radiación solar (N/m²):", "presion_radiacion", "4.56e-6"),
            ("Factor de reflectividad (Qpr):", "factor_reflectividad", "1.2"),
            ("Fracción solar (fSol):", "fraccion_solar", "0.6"),
        ], c["txt_tit"])
        
        self._seccion(frame, "Parámetros para δ_masa, δ_solar, δ_eclipse", [
            ("Masa inicial del satélite (kg):", "masa_inicial_kg", "300.0"),
            ("Masa final esperada (kg):", "masa_final_kg", "280.0"),
            ("Índice F10.7 actual:", "indice_f107", "70.0"),
            ("Índice F10.7 referencia:", "f107_referencia", "70.0"),
            ("Inclinación orbital (grados):", "inclinacion_orbital", "90.0"),
        ], c["ambar"])
        
        self._seccion(frame, "Información de Misión", [
            ("Consumo diario propelente (kg):", "consumo_diario", "0.001"),
            ("Costo misión (millones USD):", "costo_mision", "350"),
            ("Duración prevista (años):", "duracion_mision", "4"),
        ], c["txt_sub"])
    
    def _seccion(self, parent, titulo, items, color):
        c = self.c
        frm = tk.Frame(parent, bg=c["bg2"], relief=tk.RIDGE, bd=1)
        frm.pack(fill=tk.X, padx=8, pady=6, ipady=6)
        tk.Label(frm, text=titulo, font=("Segoe UI", 10, "bold"), fg=color, bg=c["bg2"]).pack(anchor=tk.W, padx=8, pady=(6, 4))
        
        for etiqueta, clave, defecto in items:
            fila = tk.Frame(frm, bg=c["bg2"])
            fila.pack(fill=tk.X, padx=8, pady=1)
            tk.Label(fila, text=etiqueta, font=("Segoe UI", 9), fg=c["txt_dim"], bg=c["bg2"],
                     width=30, anchor=tk.W).pack(side=tk.LEFT)
            var = tk.StringVar(value=defecto)
            if clave in ["consumo_diario", "costo_mision", "duracion_mision"]:
                self.vars_i[clave] = var
            else:
                self.vars_f[clave] = var
            tk.Entry(fila, textvariable=var, font=("Consolas", 9), bg=c["entrada"],
                     fg=c["verde"], insertbackground=c["verde"], relief=tk.SUNKEN, width=14).pack(side=tk.LEFT, padx=6)
    
    def _pestana_misiones(self):
        c = self.c
        p = tk.Frame(self.nb, bg=c["bg"])
        self.nb.add(p, text="  Misiones  ")
        
        frm = tk.Frame(p, bg=c["bg"])
        frm.pack(fill=tk.BOTH, expand=True, padx=15, pady=15)
        
        tk.Label(frm, text="Base de Datos de Misiones Reales", font=("Segoe UI", 11, "bold"),
                 fg=c["txt_tit"], bg=c["bg"]).pack(pady=(0, 10))
        
        self.lista_misiones = tk.Listbox(frm, font=("Segoe UI", 10), bg=c["bg2"],
                                          fg=c["txt_sub"], selectbackground=c["btn_ok"],
                                          selectforeground="white", relief=tk.SUNKEN, height=8)
        self.lista_misiones.pack(fill=tk.BOTH, expand=True, pady=4)
        for nombre in BASE_MISIONES:
            self.lista_misiones.insert(tk.END, nombre)
        
        tk.Button(frm, text="Cargar misión seleccionada", command=self._cargar_mision,
                  bg=c["btn_ok"], fg="white", font=("Segoe UI", 10, "bold"),
                  padx=15, pady=5, cursor="hand2", relief=tk.FLAT).pack(pady=10)
    
    def _pestana_graficos(self):
        c = self.c
        p = tk.Frame(self.nb, bg=c["bg"])
        self.nb.add(p, text="  Proyección  ")
        
        self.canvas_graf = tk.Canvas(p, bg=c["bg2"], highlightthickness=0)
        self.canvas_graf.pack(fill=tk.BOTH, expand=True, padx=6, pady=6)
        
        self.txt_graf = tk.Text(p, font=("Consolas", 9), bg=c["bg2"], fg=c["verde"],
                                 relief=tk.SUNKEN, padx=8, pady=8, height=8, wrap=tk.WORD)
        self.txt_graf.pack(fill=tk.X, padx=6, pady=4)
        self.txt_graf.insert(tk.END, "Presione CALCULAR para generar la proyección.\n")
        self.txt_graf.configure(state=tk.DISABLED)
    
    def _pestana_resultados(self):
        c = self.c
        p = tk.Frame(self.nb, bg=c["bg"])
        self.nb.add(p, text="  Resultados  ")
        
        self.txt_res = tk.Text(p, font=("Consolas", 9), bg=c["entrada"], fg=c["verde"],
                                insertbackground=c["verde"], wrap=tk.WORD, relief=tk.SUNKEN, padx=10, pady=10)
        self.txt_res.pack(fill=tk.BOTH, expand=True, padx=6, pady=6)
        self.txt_res.insert(tk.END, "╔══════════════════════════════════════════╗\n")
        self.txt_res.insert(tk.END, "║  MODELO DE CORRECCIÓN ORBITAL v3.2      ║\n")
        self.txt_res.insert(tk.END, "║  π_corr = π · (1+Σδ) / (1+ΣΔ)          ║\n")
        self.txt_res.insert(tk.END, "╚══════════════════════════════════════════╝\n\n")
        self.txt_res.insert(tk.END, "▶ Cargue una misión o ingrese valores y presione CALCULAR.\n")
    
    def _sincronizar(self):
        try:
            for clave, var in self.vars_e.items():
                val = float(var.get())
                if clave in self.modelo.variables_delta:
                    self.modelo.variables_delta[clave]["valor"] = val
                elif clave in self.modelo.variables_delta_mayus:
                    self.modelo.variables_delta_mayus[clave]["valor"] = val
            for clave, var in self.vars_f.items():
                txt = var.get()
                if 'e' in txt.lower():
                    partes = txt.lower().split('e')
                    val = float(partes[0]) * (10 ** int(partes[1]))
                else:
                    val = float(txt)
                if clave in self.modelo.parametros_fisicos:
                    self.modelo.parametros_fisicos[clave] = val
            if "consumo_diario" in self.vars_i:
                self.modelo.consumo_diario_kg = float(self.vars_i["consumo_diario"].get())
            if "costo_mision" in self.vars_i:
                self.modelo.costo_mision_millones = float(self.vars_i["costo_mision"].get())
            if "duracion_mision" in self.vars_i:
                self.modelo.duracion_mision_anos = float(self.vars_i["duracion_mision"].get())
            return True
        except ValueError as e:
            messagebox.showerror("Error", f"Valor inválido. Use punto decimal.\n{str(e)}")
            return False
    
    def _estimar_nuevos(self):
        """Estima δ_masa, δ_solar, δ_eclipse desde parámetros físicos.
        Ahora sincroniza correctamente y muestra ventana de confirmación."""
        
        # Sincronizar parámetros físicos desde la interfaz hacia el modelo
        for clave, var in self.vars_f.items():
            try:
                txt = var.get()
                if 'e' in txt.lower():
                    partes = txt.lower().split('e')
                    val = float(partes[0]) * (10 ** int(partes[1]))
                else:
                    val = float(txt)
                if clave in self.modelo.parametros_fisicos:
                    self.modelo.parametros_fisicos[clave] = val
            except ValueError:
                pass
        
        # Calcular los tres nuevos δ
        dm = self.modelo.estimar_delta_masa()
        ds = self.modelo.estimar_delta_solar()
        de = self.modelo.estimar_delta_eclipse()
        
        # Actualizar el modelo
        self.modelo.variables_delta["δ_masa"]["valor"] = dm
        self.modelo.variables_delta["δ_solar"]["valor"] = ds
        self.modelo.variables_delta["δ_eclipse"]["valor"] = de
        
        # Actualizar la interfaz gráfica (tabla de δ)
        if "δ_masa" in self.vars_e:
            self.vars_e["δ_masa"].set(f"{dm:.6f}")
        if "δ_solar" in self.vars_e:
            self.vars_e["δ_solar"].set(f"{ds:.6f}")
        if "δ_eclipse" in self.vars_e:
            self.vars_e["δ_eclipse"].set(f"{de:.6f}")
        
        # Mostrar ventana de confirmación
        m_ini = self.modelo.parametros_fisicos.get("masa_inicial_kg", "N/D")
        m_act = self.modelo.parametros_fisicos.get("masa_satelite", "N/D")
        f107 = self.modelo.parametros_fisicos.get("indice_f107", "N/D")
        fref = self.modelo.parametros_fisicos.get("f107_referencia", "N/D")
        fsol = self.modelo.parametros_fisicos.get("fraccion_solar", "N/D")
        
        mensaje = f"δ_masa = {dm:.6f}\n"
        mensaje += f"  (Masa inicial: {m_ini} kg → Masa actual: {m_act} kg)\n\n"
        mensaje += f"δ_solar = {ds:.6f}\n"
        mensaje += f"  (F10.7 actual: {f107} / referencia: {fref})\n\n"
        mensaje += f"δ_eclipse = {de:.6f}\n"
        mensaje += f"  (Fracción solar: {fsol})\n\n"
        mensaje += "Valores actualizados en la pestaña Perímetro (δ)."
        
        messagebox.showinfo("Estimación de δ completada", mensaje)
        self.sts.set(f"δ_masa={dm:.6f} | δ_solar={ds:.6f} | δ_eclipse={de:.6f}")
    
    def _cargar_mision(self):
        sel = self.lista_misiones.curselection()
        if not sel:
            messagebox.showinfo("Seleccione", "Elija una misión de la lista.")
            return
        nombre = self.lista_misiones.get(sel[0])
        if self.modelo.cargar_mision(nombre):
            for clave, var in self.vars_e.items():
                if clave in self.modelo.variables_delta:
                    var.set(str(self.modelo.variables_delta[clave]["valor"]))
                elif clave in self.modelo.variables_delta_mayus:
                    var.set(str(self.modelo.variables_delta_mayus[clave]["valor"]))
            for clave, var in self.vars_f.items():
                if clave in self.modelo.parametros_fisicos:
                    var.set(str(self.modelo.parametros_fisicos[clave]))
            self.vars_i["consumo_diario"].set(str(self.modelo.consumo_diario_kg))
            self.vars_i["costo_mision"].set(str(self.modelo.costo_mision_millones))
            self.vars_i["duracion_mision"].set(str(self.modelo.duracion_mision_anos))
            self.sts.set(f"Misión '{nombre}' cargada")
            self._calcular()
    
    def _calcular(self):
        if not self._sincronizar():
            return
        
        res = self.modelo.calcular_factor()
        if res is None or res[0] is None:
            messagebox.showerror("Error", "División por cero.")
            return
        factor, sd, sD = res
        
        pi_corr = self.modelo.calcular_pi_corregido()
        if pi_corr is None:
            return
        
        pi_trad = self.modelo.pi
        dif = pi_corr - pi_trad
        mej = abs(dif) / pi_trad * 100
        
        self.lbl_res.set(f"π_corr = {pi_corr:.6f} | π_trad = {pi_trad:.6f} | Mejora = {mej:.4f}%")
        self.lbl_info.set(f"Factor = {factor:.6f} | Σδ = {sd:.6f} | ΣΔ = {sD:.6f}")
        
        aho = self.modelo.calcular_ahorro()
        if aho is None:
            return
        
        t = self.txt_res
        t.delete(1.0, tk.END)
        t.insert(tk.END, "╔══════════════════════════════════════════════════╗\n")
        t.insert(tk.END, "║     RESULTADOS DEL MODELO DE CORRECCIÓN v3.2    ║\n")
        t.insert(tk.END, "╚══════════════════════════════════════════════════╝\n\n")
        t.insert(tk.END, f"▶ FÓRMULA: π_corr = {pi_trad:.6f} · (1+{sd:.6f})/(1+{sD:.6f})\n\n")
        t.insert(tk.END, f"▶ VARIABLES DEL PERÍMETRO (δ): Σδ = {sd:+.6f}\n")
        for clave, datos in self.modelo.variables_delta.items():
            if abs(datos["valor"]) > 0.000001:
                marca = " [NUEVA]" if clave in ["δ_masa", "δ_solar", "δ_eclipse"] else ""
                t.insert(tk.END, f"   {clave}: {datos['valor']:+.6f}{marca}\n")
        t.insert(tk.END, f"\n▶ VARIABLES DEL DIÁMETRO (Δ): ΣΔ = {sD:+.6f}\n")
        for clave, datos in self.modelo.variables_delta_mayus.items():
            if abs(datos["valor"]) > 0.000001:
                t.insert(tk.END, f"   {clave}: {datos['valor']:+.6f}\n")
        t.insert(tk.END, f"\n▶ RESULTADOS FINALES:\n")
        t.insert(tk.END, f"   Factor de corrección: {factor:.6f}\n")
        t.insert(tk.END, f"   π tradicional:        {pi_trad:.6f}\n")
        t.insert(tk.END, f"   π corregido:          {pi_corr:.6f}\n")
        t.insert(tk.END, f"   Diferencia:           {dif:+.6f}\n")
        t.insert(tk.END, f"   Mejora relativa:      {mej:.4f}%\n")
        t.insert(tk.END, f"\n▶ AHORRO DE PROPELENTE (±{aho['inc_pct']:.0f}% incertidumbre):\n")
        t.insert(tk.END, f"   Consumo diario:       {self.modelo.consumo_diario_kg:.6f} kg\n")
        t.insert(tk.END, f"   Ahorro diario (mín):  {aho['ahorro_diario_min']:.6f} kg\n")
        t.insert(tk.END, f"   Ahorro diario (máx):  {aho['ahorro_diario_max']:.6f} kg\n")
        t.insert(tk.END, f"   Ahorro anual (mín):   {aho['ahorro_anual_min']:.4f} kg\n")
        t.insert(tk.END, f"   Ahorro anual (máx):   {aho['ahorro_anual_max']:.4f} kg\n")
        t.insert(tk.END, f"   Días extra (mín):     {aho['dias_min']} días\n")
        t.insert(tk.END, f"   Días extra (máx):     {aho['dias_max']} días\n")
        if aho['valor_max'] > 0:
            t.insert(tk.END, f"\n▶ VALOR ECONÓMICO ESTIMADO:\n")
            t.insert(tk.END, f"   Costo diario misión:  ${aho['costo_diario']:.2f}M\n")
            t.insert(tk.END, f"   Valor generado (mín): ${aho['valor_min']:.1f}M\n")
            t.insert(tk.END, f"   Valor generado (máx): ${aho['valor_max']:.1f}M\n")
        t.insert(tk.END, f"\n▶ PROYECCIÓN ACUMULADA (4 años):\n")
        for ano in range(1, 5):
            t.insert(tk.END, f"   Año {ano}: {aho['dias_min']*ano}-{aho['dias_max']*ano} días extra\n")
        
        self.sts.set("✅ Cálculo completado")
        self._actualizar_grafico(aho)
        self.nb.select(5)
    
    def _actualizar_grafico(self, aho):
        self.txt_graf.configure(state=tk.NORMAL)
        self.txt_graf.delete(1.0, tk.END)
        self.txt_graf.insert(tk.END, "PROYECCIÓN ACUMULADA DE DÍAS EXTRA DE MISIÓN\n" + "="*50 + "\n\n")
        max_dias = aho['dias_max'] * 4
        if max_dias > 0:
            escala = 40 / max_dias
            for ano in range(1, 5):
                dmin = aho['dias_min'] * ano
                dmax = aho['dias_max'] * ano
                bmin = int(dmin * escala)
                bmax = int(dmax * escala)
                self.txt_graf.insert(tk.END, f"Año {ano}: " + "█"*bmin + "░"*(bmax-bmin) + f"  {dmin}-{dmax} días\n")
            self.txt_graf.insert(tk.END, "\n" + "─"*50 + "\n█ Conservador | ░ Adicional\n")
        self.txt_graf.configure(state=tk.DISABLED)
        
        self.canvas_graf.delete("all")
        w = self.canvas_graf.winfo_width()
        h = self.canvas_graf.winfo_height()
        if w < 100 or h < 100 or max_dias <= 0:
            return
        mx, my = 70, 35
        ag = w - mx - 30
        al = h - my - 30
        self.canvas_graf.create_line(mx, h-my, w-15, h-my, fill="#30363d", width=2)
        self.canvas_graf.create_line(mx, h-my, mx, 15, fill="#30363d", width=2)
        ex = ag / max_dias
        for ano in range(1, 5):
            dmin = aho['dias_min'] * ano
            dmax = aho['dias_max'] * ano
            xmin = mx + dmin * ex
            xmax = mx + dmax * ex
            yb = h - my
            ab = 30
            self.canvas_graf.create_rectangle(mx, yb-ab*ano+8, xmin, yb-ab*(ano-1)-3, fill="#238636", outline="")
            self.canvas_graf.create_rectangle(xmin, yb-ab*ano+8, xmax, yb-ab*(ano-1)-3, fill="#3fb950", outline="", stipple="gray25")
            self.canvas_graf.create_text(mx-45, yb-ab*ano+ab/2, text=f"Año {ano}", fill="#8b949e", font=("Segoe UI", 8), anchor=tk.E)
            self.canvas_graf.create_text(xmax+8, yb-ab*ano+ab/2, text=f"{dmin}-{dmax} d", fill="#c9d1d9", font=("Segoe UI", 7), anchor=tk.W)
        self.canvas_graf.create_text(w/2, 12, text="Días Extra de Misión - Proyección Acumulada", fill="#58a6ff", font=("Segoe UI", 9, "bold"))
    
    def _generar_reporte(self):
        if not self._sincronizar():
            return None
        res = self.modelo.calcular_factor()
        if res is None:
            return None
        factor, sd, sD = res
        pi_corr = self.modelo.calcular_pi_corregido()
        aho = self.modelo.calcular_ahorro()
        pi_trad = self.modelo.pi
        dif = pi_corr - pi_trad if pi_corr else 0
        mej = abs(dif) / pi_trad * 100 if pi_trad else 0
        
        lineas = []
        lineas.append("="*60)
        lineas.append("INFORME DE CORRECCIÓN ORBITAL INTEGRADA v3.2")
        lineas.append(f"Fecha: {datetime.now().strftime('%Y-%m-%d %H:%M')}")
        lineas.append("="*60)
        lineas.append("")
        lineas.append(f"Factor de corrección: {factor:.6f}")
        lineas.append(f"π tradicional:        {pi_trad:.6f}")
        lineas.append(f"π corregido:          {pi_corr:.6f}")
        lineas.append(f"Diferencia:           {dif:+.6f}")
        lineas.append(f"Mejora relativa:      {mej:.4f}%")
        lineas.append("")
        lineas.append("VARIABLES DEL PERÍMETRO (δ):")
        for clave, datos in self.modelo.variables_delta.items():
            if abs(datos["valor"]) > 0.000001:
                lineas.append(f"  {clave}: {datos['valor']:+.6f}")
        lineas.append(f"  Σδ = {sd:+.6f}")
        lineas.append("")
        lineas.append("VARIABLES DEL DIÁMETRO (Δ):")
        for clave, datos in self.modelo.variables_delta_mayus.items():
            if abs(datos["valor"]) > 0.000001:
                lineas.append(f"  {clave}: {datos['valor']:+.6f}")
        lineas.append(f"  ΣΔ = {sD:+.6f}")
        lineas.append("")
        lineas.append("AHORRO DE PROPELENTE:")
        lineas.append(f"  Consumo diario:       {self.modelo.consumo_diario_kg:.6f} kg")
        lineas.append(f"  Ahorro diario (mín):  {aho['ahorro_diario_min']:.6f} kg")
        lineas.append(f"  Ahorro diario (máx):  {aho['ahorro_diario_max']:.6f} kg")
        lineas.append(f"  Ahorro anual (mín):   {aho['ahorro_anual_min']:.4f} kg")
        lineas.append(f"  Ahorro anual (máx):   {aho['ahorro_anual_max']:.4f} kg")
        lineas.append(f"  Días extra (mín):     {aho['dias_min']} días")
        lineas.append(f"  Días extra (máx):     {aho['dias_max']} días")
        lineas.append("")
        lineas.append("PROYECCIÓN ACUMULADA (4 años):")
        for ano in range(1, 5):
            lineas.append(f"  Año {ano}: {aho['dias_min']*ano}-{aho['dias_max']*ano} días extra")
        if aho['valor_max'] > 0:
            lineas.append("")
            lineas.append("VALOR ECONÓMICO:")
            lineas.append(f"  Costo diario:  ${aho['costo_diario']:.2f}M")
            lineas.append(f"  Valor (mín):   ${aho['valor_min']:.1f}M")
            lineas.append(f"  Valor (máx):   ${aho['valor_max']:.1f}M")
        lineas.append("")
        lineas.append("="*60)
        return "\n".join(lineas), pi_corr, factor, aho
    
    def _exportar_txt(self):
        reporte = self._generar_reporte()
        if reporte is None:
            return
        texto, _, _, _ = reporte
        archivo = filedialog.asksaveasfilename(defaultextension=".txt", filetypes=[("Texto", "*.txt")])
        if archivo:
            with open(archivo, 'w', encoding='utf-8') as f:
                f.write(texto)
            messagebox.showinfo("Éxito", "Informe TXT exportado.")
    
    def _exportar_pdf(self):
        reporte = self._generar_reporte()
        if reporte is None:
            return
        texto, _, _, _ = reporte
        archivo = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF", "*.pdf")])
        if archivo:
            try:
                with open(archivo, 'w', encoding='utf-8') as f:
                    f.write("%PDF-1.4\n")
                    f.write("1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj\n")
                    f.write("2 0 obj\n<< /Type /Pages /Kids [3 0 R] /Count 1 >>\nendobj\n")
                    f.write("3 0 obj\n<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Contents 4 0 R /Resources << /Font << /F1 5 0 R >> >> >>\nendobj\n")
                    f.write("4 0 obj\n<< /Length 6 0 R >>\nstream\nBT\n/F1 10 Tf\n")
                    for linea in texto.split('\n'):
                        f.write(f"({linea}) Tj\nT*\n")
                    f.write("ET\nendstream\nendobj\n")
                    f.write("5 0 obj\n<< /Type /Font /Subtype /Type1 /BaseFont /Courier >>\nendobj\n")
                    f.write("6 0 obj\n0\nendobj\n")
                    f.write("xref\n0 7\n0000000000 65535 f \n0000000009 00000 n \n0000000058 00000 n \n0000000115 00000 n \n0000000285 00000 n \n0000000400 00000 n \n0000000456 00000 n \n")
                    f.write("trailer\n<< /Size 7 /Root 1 0 R >>\nstartxref\n492\n%%EOF\n")
                messagebox.showinfo("Éxito", "Informe PDF exportado.")
            except Exception as e:
                messagebox.showerror("Error", f"No se pudo crear PDF:\n{str(e)}")
    
    def _exportar_docx(self):
        reporte = self._generar_reporte()
        if reporte is None:
            return
        texto, pi_corr, factor, aho = reporte
        archivo = filedialog.asksaveasfilename(defaultextension=".docx", filetypes=[("Word", "*.docx")])
        if archivo:
            try:
                doc_body = ""
                for linea in texto.split('\n'):
                    doc_body += f"<w:p><w:r><w:t>{linea}</w:t></w:r></w:p>\n"
                
                with zipfile.ZipFile(archivo, 'w', zipfile.ZIP_DEFLATED) as z:
                    z.writestr('[Content_Types].xml', '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>')
                    z.writestr('_rels/.rels', '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')
                    z.writestr('word/document.xml', f'<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>{doc_body}</w:body></w:document>')
                
                messagebox.showinfo("Éxito", "Informe DOCX exportado.\nCompatible con Word y LibreOffice.")
            except Exception as e:
                messagebox.showerror("Error", f"No se pudo crear DOCX:\n{str(e)}")
    
    def _guardar(self):
        self._sincronizar()
        archivo = filedialog.asksaveasfilename(defaultextension=".json", filetypes=[("JSON", "*.json")])
        if archivo:
            self.modelo.guardar_configuracion(archivo)
            self.sts.set(f"Configuración guardada: {os.path.basename(archivo)}")
    
    def _cargar(self):
        archivo = filedialog.askopenfilename(filetypes=[("JSON", "*.json")])
        if archivo:
            self.modelo.cargar_configuracion(archivo)
            for clave, var in self.vars_e.items():
                if clave in self.modelo.variables_delta:
                    var.set(str(self.modelo.variables_delta[clave]["valor"]))
                elif clave in self.modelo.variables_delta_mayus:
                    var.set(str(self.modelo.variables_delta_mayus[clave]["valor"]))
            for clave, var in self.vars_f.items():
                if clave in self.modelo.parametros_fisicos:
                    var.set(str(self.modelo.parametros_fisicos[clave]))
            self.vars_i["consumo_diario"].set(str(self.modelo.consumo_diario_kg))
            self.sts.set(f"Configuración cargada: {os.path.basename(archivo)}")
            self._calcular()
    
    def _limpiar(self):
        for clave in self.modelo.variables_delta:
            self.modelo.variables_delta[clave]["valor"] = 0.0
            if clave in self.vars_e:
                self.vars_e[clave].set("0.0")
        for clave in self.modelo.variables_delta_mayus:
            self.modelo.variables_delta_mayus[clave]["valor"] = 0.0
            if clave in self.vars_e:
                self.vars_e[clave].set("0.0")
        self.lbl_res.set("π_corr = --")
        self.lbl_info.set("v3.2 · Variables reiniciadas")
        self.txt_res.delete(1.0, tk.END)
        self.txt_res.insert(tk.END, "▶ Variables reiniciadas.\n")
        self.txt_graf.configure(state=tk.NORMAL)
        self.txt_graf.delete(1.0, tk.END)
        self.txt_graf.insert(tk.END, "Presione CALCULAR.\n")
        self.txt_graf.configure(state=tk.DISABLED)
        self.canvas_graf.delete("all")
        self.sts.set("🔄 Variables reiniciadas")


def main():
    ventana = tk.Tk()
    InterfazMejorada(ventana)
    ventana.mainloop()

if __name__ == "__main__":
    main()

