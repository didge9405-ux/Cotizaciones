import os
import json
import streamlit as st
from docx import Document
from google import genai

# Configuración de la página web
st.set_page_config(
    page_title="Generador de Cotizaciones IA - YAVIS",
    page_icon="⚡",
    layout="centered"
)

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

st.title("⚡ Cotizador IA - YAVIS (Móvil & PC)")
st.write("Escribe tu solicitud de corrido o rellena los campos para generar la cotización conservando tus logos.")

# --- SECCIÓN DE ENTRADA DE TEXTO LIBRE PARA LA IA ---
st.markdown("### 🤖 Procesamiento Inteligente de Texto")
texto_prompt = st.text_area(
    "Escribe o pega tu solicitud (Ej: 'Cotización para Bimbo, área de mantenimiento, folio 88, 5 motores, precio unitario 130000'):",
    value="Cotización para Bimbo, área de mantenimiento, folio 88, 5 piezas de motores con costo unitario de 130000"
)

# Inicializar variables en session_state si no existen
if 'datos_ia' not in st.session_state:
    st.session_state['datos_ia'] = {
        "folio": "153",
        "empresa": "GRUPO GUSI",
        "area": "TRANSPORTES",
        "componente": "EDIFICIO",
        "alcances": "",
        "cantidad": "2",
        "descripcion": "Servicio de mantenimiento general",
        "subtotal": 23000.0
    }

if st.button("✨ Rellenar Formulario Automáticamente con IA", type="primary", use_container_width=True):
    if not texto_prompt.strip():
        st.warning("Por favor, escribe una instrucción en la caja de texto.")
    else:
        try:
            # Nota: Asegúrate de configurar tu API Key de Gemini en los Secrets de Streamlit Cloud
            api_key = st.secrets.get("GEMINI_API_KEY", os.environ.get("GEMINI_API_KEY", ""))
            client = genai.Client(api_key=api_key) if api_key else genai.Client()
            
            prompt = f"""
            Analiza la siguiente instrucción de una cotización industrial: "{texto_prompt}"
            Extrae y clasifica la información en formato JSON estricto con las siguientes llaves exactas:
            - folio (texto o número)
            - empresa (nombre del cliente)
            - area (departamento o área)
            - componente (elemento o equipo)
            - alcances (descripción general de alcances u observaciones)
            - cantidad (cantidad numérica para la tabla)
            - descripcion (descripción del producto/servicio en la tabla)
            - subtotal (número puro sin símbolos de moneda, correspondiente al total de la partida o subtotal)
            
            Devuelve ÚNICAMENTE el JSON válido, sin texto adicional ni bloques markdown de código.
            """
            
            response = client.models.generate_content(model='gemini-2.5-flash', contents=prompt)
            raw_text = response.text.strip()
            
            if raw_text.startswith("```json"):
                raw_text = raw_text[7:]
            if raw_text.endswith("```"):
                raw_text = raw_text[:-3]
            
            datos_nuevos = json.loads(raw_text.strip())
            
            # Actualizar session_state
            for k, v in datos_nuevos.items():
                if k in st.session_state['datos_ia']:
                    st.session_state['datos_ia'][k] = v
                    
            st.success("¡Formulario rellenado automáticamente por la IA con éxito!")
            st.rerun()
            
        except Exception as e:
            st.error(f"Error al conectar con la IA: {str(e)}")

st.divider()

# --- FORMULARIO EDITABLE ---
st.markdown("### 📋 Detalles de la Cotización")

d = st.session_state['datos_ia']

folio = st.text_input("Folio:", value=str(d.get("folio", "153")))
empresa = st.text_input("Empresa / Cliente:", value=str(d.get("empresa", "GRUPO GUSI")))
area = st.text_input("Área:", value=str(d.get("area", "TRANSPORTES")))
componente = st.text_input("Componente:", value=str(d.get("componente", "EDIFICIO")))
alcances = st.text_area("Alcances / Observaciones:", value=str(d.get("alcances", "")))
cantidad = st.text_input("Cantidad de la tabla:", value=str(d.get("cantidad", "1")))
descripcion = st.text_area("Descripción dentro de la tabla:", value=str(d.get("descripcion", "")))

try:
    sub_default = float(d.get("subtotal", 23000.0))
except:
    sub_default = 23000.0

subtotal = st.number_input("Subtotal (número para cálculos):", value=sub_default, step=100.0)
precio_unitario = f"${subtotal:,.2f} pesos"
nombre_archivo = f"COT_{folio}_{str(empresa).replace(' ', '_')}.docx"

st.divider()

# --- GENERACIÓN DE DOCUMENTO ---
if st.button("🚀 Generar y Descargar Cotización Word", type="primary", use_container_width=True):
    ruta_plantilla = 'plantilla_base.docx'
    if not os.path.exists(ruta_plantilla):
        st.error(f"No se encontró el archivo '{ruta_plantilla}' en el servidor.")
    else:
        try:
            doc = Document(ruta_plantilla)
            iva = subtotal * 0.16
            total = subtotal + iva
            precio_letra = numero_a_letras(total)

            reemplazos = {
                '[FOLIO]': str(folio),
                '[FECHA]': "15 de septiembre del 2026",
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
            ruta_docx = os.path.join('cotizaciones_generadas', nombre_archivo)
            doc.save(ruta_docx)

            st.success("¡Cotización generada correctamente!")

            # Botón de descarga directa para el celular o PC
            with open(ruta_docx, "rb") as file:
                st.download_button(
                    label="📥 Descargar Archivo Word (.docx)",
                    data=file,
                    file_name=nombre_archivo,
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True
                )

        except Exception as e:
            st.error(f"Error al generar el documento: {str(e)}")
