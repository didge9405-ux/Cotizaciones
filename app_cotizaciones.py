import os
import re
from docx import Document

ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")

try:
    from docx2pdf import convert
    PDF_DISPONIBLE = True
except ImportError:
    PDF_DISPONIBLE = False

def numero_a_letras(num):
    unidades = ("", "un", "dos", "tres", "cuatro", "cinco", "seis", "siete", "ocho", "nueve",
                "diez", "once", "doce", "trece", "catorce", "quince", "dieciséis", "diecisiete", "dieciocho", "diecinueve", "veinte")
    decenas = ("", "", "veinti", "treinta", "cuarenta", "cincuenta", "sesenta", "setenta", "ochenta", "noventa")
    centenas = ("", "ciento", "doscientos", "trescientos", "cuatrocientos", "quinientos", "seiscientos", "setecientos", "ochocientos", "novecientos")

    if num == 0: return "cero pesos M.N."
    def convertir_grupo(n):
        if n == 100: return "cien"
        res = ""
        c = n // 100
        d = (n % 100) // 10
        u = n % 10
        if c > 0: res += centenas[c] + " "
        resto = n % 100
        if resto <= 20 and resto > 0: res += unidades[resto] + " "
        elif resto > 20:
            if d == 2: res += "veinti" + unidades[u] + " "
            else:
                res += decenas[d]
                if u > 0: res += " y " + unidades[u]
                res += " "
        return res.strip()

    entero = int(num)
    if entero == 0: return "cero pesos M.N."
    letras = ""
    miles = entero // 1000
    resto_miles = entero % 1000
    if miles > 0: letras += "mil " if miles == 1 else convertir_grupo(miles) + " mil "
    if resto_miles > 0: letras += convertir_grupo(resto_miles) + " "
    return letras.strip().capitalize() + " M.N."

class AppCotizacionesOffline(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Generador de Cotizaciones - Procesamiento Local (Sin Internet) - YAVIS")
        
        self.attributes('-fullscreen', True)
        self.bind("<Escape>", lambda e: self.attributes('-fullscreen', not self.attributes('-fullscreen')))

        self.scrollable_frame = ctk.CTkScrollableFrame(self, width=850, height=700)
        self.scrollable_frame.pack(padx=40, pady=30, fill="both", expand=True)

        self.lbl_titulo = ctk.CTkLabel(
            self.scrollable_frame, 
            text="⚡ Cotizador Local: Procesamiento Inteligente de Texto (100% Offline)", 
            font=ctk.CTkFont(size=22, weight="bold")
        )
        self.lbl_titulo.pack(pady=(0, 15))

        # --- SECCIÓN DE ENTRADA DE TEXTO LIBRE LOCAL ---
        self.crear_label("Escribe o pega tu solicitud de corrido (Ej: 'Cotización de Bimbo, folio 88, área de mantenimiento, 5 motores, precio 130000'):")
        self.text_prompt_ia = ctk.CTkTextbox(self.scrollable_frame, height=70, width=780, font=ctk.CTkFont(size=12))
        self.text_prompt_ia.pack(pady=(0, 10))
        self.text_prompt_ia.insert("1.0", "Cotización para Bimbo, área de mantenimiento, folio 88, 5 piezas de motores con costo unitario de 130000")

        self.btn_procesar_local = ctk.CTkButton(
            self.scrollable_frame, 
            text="✨ Rellenar Formulario Automáticamente (Sin Internet)", 
            fg_color="#2980b9", 
            hover_color="#1f618d", 
            font=ctk.CTkFont(size=14, weight="bold"),
            command=self.procesar_texto_local
        )
        self.btn_procesar_local.pack(pady=(0, 20), fill="x", ipady=8)

        # Campos del formulario
        self.crear_campo("Nombre del archivo de salida:", "COT_153_GRUPO_GUSI", "entry_archivo")
        self.crear_campo("Folio:", "153", "entry_folio")
        self.crear_campo("Fecha:", "15 de septiembre del 2026", "entry_fecha")
        self.crear_campo("Empresa / Cliente:", "GRUPO GUSI", "entry_empresa")
        self.crear_campo("Área:", "TRANSPORTES", "entry_area")
        self.crear_campo("Componente:", "EDIFICIO", "entry_componente")

        self.crear_label("Alcances / Observaciones:")
        self.text_alcances = ctk.CTkTextbox(self.scrollable_frame, height=80, width=780, font=ctk.CTkFont(size=12))
        self.text_alcances.pack(pady=(0, 15))

        self.crear_campo("Cantidad de la tabla:", "2", "entry_cantidad")

        self.crear_label("Descripción dentro de la tabla:")
        self.text_desc = ctk.CTkTextbox(self.scrollable_frame, height=80, width=780, font=ctk.CTkFont(size=12))
        self.text_desc.pack(pady=(0, 15))

        self.crear_campo("Precio Unitario:", "$23,000.00 pesos", "entry_pu")
        self.crear_campo("Subtotal (número para cálculos de IVA):", "23000", "entry_subtotal")

        self.crear_label("Formato de salida:")
        self.var_formato = ctk.StringVar(value="docx")
        self.frame_radio = ctk.CTkFrame(self.scrollable_frame, fg_color="transparent")
        self.frame_radio.pack(anchor="w", pady=(0, 20))
        ctk.CTkRadioButton(self.frame_radio, text="Word (.docx)", variable=self.var_formato, value="docx", font=ctk.CTkFont(size=13)).pack(side="left", padx=(0, 30))
        ctk.CTkRadioButton(self.frame_radio, text="PDF (.pdf)", variable=self.var_formato, value="pdf", font=ctk.CTkFont(size=13)).pack(side="left")

        self.btn_generar = ctk.CTkButton(
            self.scrollable_frame, 
            text="🚀 Generar Cotización Conservando Logos", 
            fg_color="#2ecc71", 
            hover_color="#27ae60", 
            font=ctk.CTkFont(size=16, weight="bold"), 
            command=self.generar_cotizacion
        )
        self.btn_generar.pack(pady=15, fill="x", ipady=10)

    def crear_label(self, texto):
        lbl = ctk.CTkLabel(self.scrollable_frame, text=texto, font=ctk.CTkFont(size=13, weight="bold"))
        lbl.pack(anchor="w", pady=(2, 2))

    def crear_campo(self, label_text, default_val, attr_name):
        self.crear_label(label_text)
        entry = ctk.CTkEntry(self.scrollable_frame, width=780, height=32, font=ctk.CTkFont(size=12))
        entry.pack(anchor="w", pady=(0, 12))
        entry.insert(0, default_val)
        setattr(self, attr_name, entry)

    # ================= PROCESAMIENTO 100% LOCAL SIN INTERNET =================
    def procesar_texto_local(self):
        texto = self.text_prompt_ia.get("1.0", "end-1c").strip()
        if not texto:
            messagebox.showwarning("Campo vacío", "Por favor, escribe o pega una instrucción.")
            return

        # Extracción inteligente mediante patrones de texto (Regex)
        datos = {}
        
        # Buscar folio
        match_folio = re.search(r'(?:folio|número|num)\s*[:#]?\s*(\d+)', texto, re.IGNORECASE)
        if match_folio:
            datos['folio'] = match_folio.group(1)

        # Buscar empresa (ej. "para Bimbo", "de Bimbo", "cliente Bimbo")
        match_empresa = re.search(r'(?:para|de|cliente|empresa)\s+([A-ZÁÉÍÓÚa-záéíóú0-9\s]+?)(?:,|\s+área|\s+folio|\s+componente|\s+con|\s+precio|\s+costo|$)', texto)
        if match_empresa:
            datos['empresa'] = match_empresa.group(1).strip().upper()

        # Buscar área
        match_area = re.search(r'área\s+(?:de\s+)?([A-ZÁÉÍÓÚa-záéíóú0-9\s]+?)(?:,|\s+folio|\s+componente|\s+con|\s+precio|\s+costo|$)', texto, re.IGNORECASE)
        if match_area:
            datos['area'] = match_area.group(1).strip().upper()

        # Buscar cantidad y descripción (ej. "5 piezas de motores")
        match_cant_desc = re.search(r'(\d+)\s+(?:piezas?|unidades?|equipos?|motores?)\s+(?:de\s+)?([A-ZÁÉÍÓÚa-záéíóú0-9\s]+)', texto, re.IGNORECASE)
        if match_cant_desc:
            datos['cantidad'] = match_cant_desc.group(1)
            datos['descripcion'] = match_cant_desc.group(0).strip()
        else:
            datos['cantidad'] = "1"
            datos['descripcion'] = texto

        # Buscar precio / subtotal (ej. "precio unitario 130000", "costo 12000")
        match_precio = re.search(r'(?:precio|costo|subtotal|unitario)\s*(?:unitario)?\s*(?:de)?\s*\$?([\d,]+\.?\d*)', texto, re.IGNORECASE)
        if match_precio:
            num_limpio = match_precio.group(1).replace(",", "")
            datos['subtotal'] = float(num_limpio)
        else:
            datos['subtotal'] = 0.0

        # Aplicar los datos extraídos en la interfaz
        if 'folio' in datos:
            self.entry_folio.delete(0, "end")
            self.entry_folio.insert(0, datos['folio'])
            emp_val = datos.get('empresa', 'CLIENTE')
            self.entry_archivo.delete(0, "end")
            self.entry_archivo.insert(0, f"COT_{datos['folio']}_{emp_val.replace(' ', '_')}")
        if 'empresa' in datos:
            self.entry_empresa.delete(0, "end")
            self.entry_empresa.insert(0, datos['empresa'])
        if 'area' in datos:
            self.entry_area.delete(0, "end")
            self.entry_area.insert(0, datos['area'])
        if 'cantidad' in datos:
            self.entry_cantidad.delete(0, "end")
            self.entry_cantidad.insert(0, datos['cantidad'])
        if 'descripcion' in datos:
            self.text_desc.delete("1.0", "end")
            self.text_desc.insert("1.0", datos['descripcion'])
        if 'subtotal' in datos and datos['subtotal'] > 0:
            sub = datos['subtotal']
            self.entry_subtotal.delete(0, "end")
            self.entry_subtotal.insert(0, str(sub))
            self.entry_pu.delete(0, "end")
            self.entry_pu.insert(0, f"${sub:,.2f} pesos")

        messagebox.showinfo("✨ Procesado Local Exitoso", "¡Formulario rellenado de forma local sin usar internet!")

    # ================= GENERACIÓN DE DOCUMENTO PROTEGIENDO LOGOS =================
    def generar_cotizacion(self):
        nombre_archivo = self.entry_archivo.get().strip()
        folio = self.entry_folio.get()
        fecha = self.entry_fecha.get()
        empresa = self.entry_empresa.get()
        area = self.entry_area.get()
        componente = self.entry_componente.get()
        alcances = self.text_alcances.get("1.0", "end-1c").strip()
        cantidad = self.entry_cantidad.get()
        descripcion = self.text_desc.get("1.0", "end-1c").strip()
        precio_unitario = self.entry_pu.get()
        formato = self.var_formato.get()

        try:
            subtotal = float(self.entry_subtotal.get())
        except ValueError:
            messagebox.showerror("Error", "El subtotal debe ser un número válido.")
            return

        ruta_plantilla = 'plantilla_base.docx'
        if not os.path.exists(ruta_plantilla):
            messagebox.showerror("Error", f"No se encontró el archivo '{ruta_plantilla}'.")
            return

        try:
            doc = Document(ruta_plantilla)
            iva = subtotal * 0.16
            total = subtotal + iva
            precio_letra = numero_a_letras(total)

            reemplazos = {
                '[FOLIO]': str(folio),
                '[FECHA]': str(fecha),
                '[EMPRESA]': str(empresa),
                '[AREA]': str(area),
                '[COMPONENTE]': str(componente),
                '[ALCANCES]': str(alcances),
                '[CANTIDAD]': str(cantidad),
                '[DESCRIPCION]': str(descripcion),
                '[PRECIO_UNITARIO]': str(precio_unitario),
                '[PRECIO_TOTAL]': f"${subtotal:,.2f} pesos",
                '[SUBTOTAL]': f"${subtotal:,.2f}",
                '[IVA]': f"${iva:,.2f}",
                '[TOTAL]': f"${total:,.2f}",
                '[PRECIO_LETRA]': precio_letra
            }

            for parrafo in doc.paragraphs:
                for clave, valor in reemplazos.items():
                    if clave in parrafo.text:
                        for run in parrafo.runs:
                            if clave in run.text:
                                run.text = run.text.replace(clave, valor)

            for tabla in doc.tables:
                for fila in tabla.rows:
                    for celda in fila.cells:
                        for clave, valor in reemplazos.items():
                            if clave in celda.text:
                                for parrafo in celda.paragraphs:
                                    for run in parrafo.runs:
                                        if clave in run.text:
                                            run.text = run.text.replace(clave, valor)

            os.makedirs('cotizaciones_generadas', exist_ok=True)
            if not nombre_archivo.endswith(".docx"): nombre_archivo += ".docx"
            ruta_docx = os.path.join('cotizaciones_generadas', nombre_archivo)
            doc.save(ruta_docx)

            if formato == "pdf":
                if not PDF_DISPONIBLE:
                    messagebox.showwarning("Aviso", "docx2pdf no disponible. Se guardó como Word.")
                    ruta_final = ruta_docx
                else:
                    ruta_pdf = ruta_docx.replace(".docx", ".pdf")
                    convert(ruta_docx, ruta_pdf)
                    ruta_final = ruta_pdf
            else:
                ruta_final = ruta_docx

            messagebox.showinfo("¡Éxito!", f"Cotización creada conservando los logos:\n{ruta_final}")

        except Exception as e:
            messagebox.showerror("Error", str(e))

if __name__ == "__main__":
    app = AppCotizacionesOffline()
    app.mainloop()
