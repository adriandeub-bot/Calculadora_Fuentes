import pandas as pd
import re
import os
from difflib import SequenceMatcher

# CONFIGURACIÓN
try:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))
except NameError:
    BASE_DIR = os.getcwd()
RUTA_EXCEL = os.path.join(BASE_DIR, "FuentesAlimentacion.xlsx")

# FUNCIONES DE UTILIDAD
def _norm(s: str) -> str:
    s = (s or "").strip().lower()
    rep = {"á":"a","é":"e","í":"i","ó":"o","ú":"u","ü":"u","ñ":"n"}
    for k,v in rep.items():
        s = s.replace(k,v)
    return s

def _try_float(x):
    try:
        return float(str(x).replace(",", ".").strip())
    except:
        return None

# Normalización avanzada para búsqueda inteligente
def _norm_avanzada(s: str) -> str:
    s = (s or "").strip().lower()
    rep = {"á":"a","é":"e","í":"i","ó":"o","ú":"u","ü":"u","ñ":"n"}
    for k,v in rep.items():
        s = s.replace(k,v)
    
    # Eliminar palabras de conexión comunes
    palabras_conexion = {
        'de', 'del', 'la', 'el', 'y', 'en', 'a', 'para', 'por', 'con', 'sin', 
        'sobre', 'bajo', 'entre', 'hacia', 'desde', 'hasta', 'mediante', 'según',
        'como', 'que', 'cuando', 'donde', 'cual', 'quien', 'cuyo', 'cuyas', 'cuyos',
        'unas', 'unos', 'una', 'un', 'lo', 'los', 'las', 'al', 'se', 'su', 'sus',
        'este', 'esta', 'estos', 'estas', 'ese', 'esa', 'esos', 'esas', 'aquel',
        'aquella', 'aquellos', 'aquellas', 'otro', 'otra', 'otros', 'otras',
        'mismo', 'misma', 'mismos', 'mismas', 'todo', 'toda', 'todos', 'todas',
        'cada', 'cualquier', 'cualesquiera', 'varios', 'varias', 'ambos', 'ambas',
        'etc', 'etcétera', 'entre otros', 'entre otras', 'para que', 'de la', 'de los',
        'de las', 'en la', 'en el', 'a la', 'al', 'del', 'y las', 'y los', 'y la', 'y el'
    }
    
    # Eliminar caracteres especiales y dividir en palabras
    s = re.sub(r'[^\w\s]', ' ', s)
    palabras = re.findall(r'\b[a-z0-9]+\b', s)
    
    # Filtrar palabras de conexión y palabras muy cortas sin significado
    palabras_filtradas = [p for p in palabras if p not in palabras_conexion and len(p) > 2]
    
    return ' '.join(palabras_filtradas)

# Función para calcular similitud entre cadenas
def _calcular_similitud(a: str, b: str) -> float:
    if not a or not b:
        return 0.0
    return SequenceMatcher(None, a, b).ratio()

# Función para buscar coincidencias con umbral de similitud
def _buscar_coincidencias(texto: str, busqueda: str, umbral=0.7) -> bool:
    if not texto or not busqueda:
        return False
    
    texto_norm = _norm_avanzada(texto)
    busqueda_norm = _norm_avanzada(busqueda)
    
    if not texto_norm or not busqueda_norm:
        return False
    
    # MEJORA: Dividir ambos textos en términos individuales
    def dividir_terminos(texto):
        # Usar los mismos separadores que en la API
        separadores = [',', ';', '/', '|', ' y ', ' e ']
        texto_para_dividir = texto
        for sep in separadores:
            texto_para_dividir = texto_para_dividir.replace(sep, ',')
        return [t.strip() for t in texto_para_dividir.split(',') if t.strip()]
    
    terminos_texto = dividir_terminos(texto_norm)
    terminos_busqueda = dividir_terminos(busqueda_norm)
    
    # Buscar si algún término de búsqueda coincide con algún término del texto
    for termino_b in terminos_busqueda:
        for termino_t in terminos_texto:
            # Si el término de búsqueda está contenido en el término del texto
            if termino_b in termino_t:
                return True
            # Calcular similitud entre términos individuales
            similitud = _calcular_similitud(termino_b, termino_t)
            if similitud >= umbral:
                return True
    
    # También verificar coincidencia completa por si acaso
    if busqueda_norm in texto_norm:
        return True
    
    similitud_completa = _calcular_similitud(texto_norm, busqueda_norm)
    return similitud_completa >= umbral

# CARGA DEL CATÁLOGO (sin expansión de voltajes)
def cargar_catalogo_expandido(ruta_excel, hoja="Sheet4", columna_expandir=None):
    try:
        df = pd.read_excel(ruta_excel, sheet_name=hoja, dtype=str)
    except FileNotFoundError:
        print(f"[ERROR] No se pudo encontrar el archivo: '{ruta_excel}'")
        return pd.DataFrame()
    except Exception as e:
        print(f"[ERROR] No se pudo leer el archivo '{ruta_excel}': {e}")
        return pd.DataFrame()

    # Normalización de datos
    df.columns = (
        df.columns.str.strip()
        .str.lower()
        .str.replace(r"[\s\-/]+", "_", regex=True)
        .str.replace(r"[()]", "", regex=True)
    )

    # --- Renombrado automático según coincidencias comunes ---
    renombres = {}
    for c in df.columns:
        if re.search(r"volt", c):
            renombres[c] = "voltaje_v"
        elif re.search(r"corr", c):
            renombres[c] = "corriente_a"
        elif re.search(r"pote|w", c):
            renombres[c] = "potencia_w"
        elif re.search(r"fuente|modelo|parte", c):
            renombres[c] = "fuente"
        elif re.search(r"entrada|salida|acdc|dcdc", c):
            renombres[c] = "entrada_salida"
        elif re.search(r"uso|aplic", c):
            renombres[c] = "usos"

    df.rename(columns=renombres, inplace=True)

    print("[DEBUG] Columnas después de renombrar:", df.columns.tolist())
    
    # NO expandir voltajes - mantener la columna original como está
    # Esto preserva tanto rangos (5-12) como múltiples valores (12,24,48)
    
    return df

# Función para verificar voltaje (maneja rangos, múltiples valores y valores simples)
def _verificar_voltaje(voltaje_catalogo, voltaje_buscado, tolerancia=0.2):
    """
    Verifica si el voltaje buscado coincide con el voltaje del catálogo.
    Maneja:
    - Valores simples: "12" 
    - Rangos: "5-12"
    - Múltiples valores: "12,24,48"
    """
    if pd.isna(voltaje_catalogo) or not voltaje_catalogo:
        return False
    
    voltaje_str = str(voltaje_catalogo).strip()
    
    # Caso 1: Es un rango (ej: "5-12")
    if '-' in voltaje_str and not voltaje_str.startswith('-'):
        try:
            min_v, max_v = map(float, voltaje_str.split('-'))
            return min_v <= voltaje_buscado <= max_v
        except:
            pass
    
    # Caso 2: Son múltiples valores (ej: "12,24,48")
    if ',' in voltaje_str:
        try:
            valores = [float(v.strip()) for v in voltaje_str.split(',')]
            for valor in valores:
                if (valor * (1 - tolerancia)) <= voltaje_buscado <= (valor * (1 + tolerancia)):
                    return True
            return False
        except:
            pass
    
    # Caso 3: Es un valor único
    try:
        valor = float(voltaje_str)
        return (valor * (1 - tolerancia)) <= voltaje_buscado <= (valor * (1 + tolerancia))
    except:
        return False

# NUEVA FUNCIÓN PARA FILTRADO CON UN SOLO PARÁMETRO
def filtrar_por_un_parametro(cat: pd.DataFrame, potencia: float, voltaje: float, corriente: float, aplicacion: str, tipo_entrada_salida: str):
    """
    Filtra el catálogo cuando solo se proporciona un parámetro numérico (voltaje, corriente o potencia).
    """
    datos = cat.copy()
    if datos.empty:
        return pd.DataFrame()

    # Contar cuántos parámetros numéricos se proporcionaron
    parametros_numericos = sum(1 for x in [voltaje, corriente, potencia] if x > 0)
    
    # Si solo hay un parámetro numérico, aplicamos filtro más flexible
    if parametros_numericos == 1:
        print("[INFO] Modo de búsqueda simple: filtrando por un solo parámetro numérico")
        
        if voltaje > 0:
            print(f"[INFO] Filtrando por voltaje: {voltaje}V")
            if 'voltaje_v' in datos.columns:
                mask = datos['voltaje_v'].apply(
                    lambda x: _verificar_voltaje(x, voltaje, 0.3)
                )
                datos = datos[mask]
                print(f"[INFO] Después de filtrar por voltaje: {len(datos)} fuentes")
            
        elif corriente > 0:
            print(f"[INFO] Filtrando por corriente: {corriente}A")
            if 'corriente_a' in datos.columns:
                # Rango más amplio para búsqueda simple
                datos = datos[pd.to_numeric(datos["corriente_a"], errors='coerce')
                              .between(corriente * 0.7, corriente * 1.3).fillna(False)]
                print(f"[INFO] Después de filtrar por corriente: {len(datos)} fuentes")
            
        elif potencia > 0:
            print(f"[INFO] Filtrando por potencia: {potencia}W")
            if 'potencia_w' in datos.columns:
                # Rango más amplio para búsqueda simple
                datos = datos[pd.to_numeric(datos["potencia_w"], errors='coerce')
                              .between(potencia * 0.7, potencia * 1.3).fillna(False)]
                print(f"[INFO] Después de filtrar por potencia: {len(datos)} fuentes")

    # Filtros de texto (siempre se aplican)
    if tipo_entrada_salida:
        if 'entrada_salida' in datos.columns:
            datos['entrada_salida'] = datos['entrada_salida'].astype(str)
            te_norm = _norm(tipo_entrada_salida)
            datos = datos[datos['entrada_salida'].str.lower().str.contains(te_norm, na=False)]
            print(f"[INFO] Después de filtrar por tipo entrada/salida: {len(datos)} fuentes")
    
    if aplicacion:
        if 'usos' in datos.columns:
            datos['usos'] = datos['usos'].astype(str)
            mask = datos['usos'].apply(
                lambda x: _buscar_coincidencias(x, aplicacion, 0.6)  # Usar la nueva función con umbral 60%
            )
            datos = datos[mask]
            print(f"[INFO] Después de filtrar por aplicación '{aplicacion}': {len(datos)} fuentes")

    return datos

# FUNCIÓN DE CÁLCULO MODIFICADA
def calcular(cat: pd.DataFrame, potencia: float, voltaje: float, corriente: float, aplicacion: str, tipo_salidas: str, tipo_entrada_salida: str = None):
    datos = cat.copy()
    if datos.empty:
        return pd.DataFrame()

    print(f"[DEBUG] Parámetros recibidos: V={voltaje}, A={corriente}, W={potencia}, App={aplicacion}, Tipo={tipo_entrada_salida}")

    # --- Verificar si es un caso de un solo parámetro ---
    parametros_numericos = sum(1 for x in [voltaje, corriente, potencia] if x > 0)
    
    if parametros_numericos == 1:
        # Usar el nuevo filtro para un solo parámetro
        datos = filtrar_por_un_parametro(cat, potencia, voltaje, corriente, aplicacion, tipo_entrada_salida)
    else:
        # --- (Filtro de fuente variable o fijas) ---
        if tipo_salidas:
            if 'variable' in datos.columns:
                col_variable_num = pd.to_numeric(datos['variable'], errors='coerce')
                
                if tipo_salidas == "1":
                    print("[INFO] Filtrando por Salida Variable (1)")
                    datos = datos[col_variable_num == 1]
                    
                elif tipo_salidas == "0":
                    print("[INFO] Filtrando por Salida Fija (0 o Nulo)")
                    datos = datos[(col_variable_num == 0) | (col_variable_num.isnull())]
            else:
                print("[WARN] No se encontró la columna 'variable' (para fijo/variable).")
    
        # --- (Filtros de texto) ---
        if tipo_entrada_salida:
            if 'entrada_salida' in datos.columns:
                datos['entrada_salida'] = datos['entrada_salida'].astype(str)
                te_norm = _norm(tipo_entrada_salida)
                datos = datos[datos['entrada_salida'].str.lower().str.contains(te_norm, na=False)]
        if aplicacion:
            if 'usos' in datos.columns:
                datos['usos'] = datos['usos'].astype(str)
                mask = datos['usos'].apply(
                    lambda x: _buscar_coincidencias(x, aplicacion, 0.6)  # Usar la nueva función con umbral 60%
                )
                datos = datos[mask]
                print(f"[INFO] Filtro aplicación '{aplicacion}': {len(datos)} fuentes después del filtro")
        if tipo_entrada_salida == "AC-DC": # Caso para fuente AC-DC
            #Filtros numéricos AC-DC
            if voltaje > 0 and 'voltaje_v' in datos.columns:
                mask = datos['voltaje_v'].apply(
                    lambda x: _verificar_voltaje(x, voltaje, 0.2)
                )
                datos = datos[mask]

            if corriente > 0:
                if 'corriente_a' in datos.columns:
                    datos = datos[pd.to_numeric(datos["corriente_a"], errors='coerce').between(corriente * 0.8, corriente * 1.2).fillna(False)]
            
            if potencia > 0:
                if 'potencia_w' in datos.columns:
                    datos = datos[pd.to_numeric(datos["potencia_w"], errors='coerce').between(potencia * 0.8, potencia * 1.2).fillna(False)]
        elif tipo_entrada_salida =="DC-AC": # Caso para fuente DC-AC
            #Filtros numéricos DC-AC
            if voltaje > 0 and 'voltaje_v' in datos.columns:
                mask = datos['voltaje_v'].apply(
                    lambda x: _verificar_voltaje(x, voltaje, 0.2)
                )
                datos = datos[mask]

            if corriente > 0 and 'corriente_a' in datos.columns:
                corriente_series = pd.to_numeric(datos["corriente_a"], errors='coerce')
                mask_corriente = (
                    (corriente_series >= corriente)
                ).fillna(False)
                datos = datos[mask_corriente]
            
            if potencia > 0 and 'potencia_w' in datos.columns:
                potencia_series = pd.to_numeric(datos["potencia_w"], errors='coerce')
                mask_potencia = (
                    (potencia_series == potencia) | 
                    (potencia_series >= potencia * 1.2)
                ).fillna(False)
                datos = datos[mask_potencia]
    
    print(f"[DEBUG] Fuentes después de todos los filtros: {len(datos)}")
    
    if datos.empty:
        return pd.DataFrame()

    # --- (Ordenamiento) ---
    sort_cols = []
    if 'potencia_w' in datos.columns:
        sort_cols.append('potencia_w')
    if 'corriente_a' in datos.columns:
        sort_cols.append('corriente_a')
    if sort_cols:
        res = datos.sort_values(by=sort_cols, ascending=True)
    else:
        res = datos
    
    # --- LÓGICA DE LIMPIEZA FINAL ---
    cols_to_drop = [c for c in res.columns if c.startswith('unnamed:')]
    res = res.drop(columns=cols_to_drop, errors="ignore")

    return res.reset_index(drop=True)

# INTERFAZ DE USUARIO

def main():
    print("=== CALCULADORA DE FUENTES ===")
    print("Si no sabe algún dato, déjelo en blanco y presione Enter.")
    print("NOTA: Ahora puede buscar solo con un parámetro (ej: solo 200W de potencia)")

    entrada_salida = input("Tipo de entrada/salida (ej. AC-DC): ").strip()
    
    # Caso para fuente AC-DC
    if entrada_salida == "AC-DC":
        respuesta_variable = input("¿Busca una salida Variable? (Si / No): ").strip()
    
        respuesta_norm = _norm(respuesta_variable)
        tipo_salidas = ""
        if respuesta_norm == "si":
            tipo_salidas = "1"
        elif respuesta_norm == "no":
            tipo_salidas = "0"

        aplicacion = input("Aplicación / Uso (ej. 'iluminación LED'): ").strip()
        voltaje = _try_float(input("Voltaje de salida (V DC, ej. 3/9/12/24/48): ")) or 0
        corriente = _try_float(input("Corriente de salida (A): ")) or 0
        potencia = _try_float(input("Potencia total de la carga (W): ")) or 0

        # Verificar que al menos un parámetro numérico esté presente
        if all(x == 0 for x in [voltaje, corriente, potencia]):
            print("\n[ERROR] Debe ingresar al menos un parámetro numérico (voltaje, corriente o potencia)")
            return

        # Cálculo automático solo si hay múltiples parámetros
        parametros_numericos = sum(1 for x in [voltaje, corriente, potencia] if x > 0)
        if parametros_numericos >= 2:
            try:
                if voltaje > 0 and corriente > 0 and potencia == 0:
                    potencia = voltaje * corriente
                    print(f"-> Potencia calculada: {potencia:.2f} W")
                elif potencia > 0 and voltaje > 0 and corriente == 0:
                    corriente = potencia / voltaje
                    print(f"-> Corriente calculada: {corriente:.2f} A")
                elif potencia > 0 and corriente > 0 and voltaje == 0:
                    voltaje = potencia / corriente
                    print(f"-> Voltaje calculado: {voltaje:.2f} V")
            except ZeroDivisionError:
                print("Error: No se puede dividir por cero.")
                return
        
    # Caso para fuentes DC-AC
    elif entrada_salida == "DC-AC":
        tipo_salidas = "0"  # Salida fija por defecto

        aplicacion = input("Aplicación / Uso (ej. electrodomésticos, dispositivos'): ").strip()
        voltaje = _try_float(input("Voltaje de salida (V AC, ej. 100, 110, 115, 120): ")) or 0
        corriente = _try_float(input("Corriente de salida (A): ")) or 0
        potencia = _try_float(input("Potencia total de la carga a alimentar (W): ")) or 0

        # Verificar que al menos un parámetro numérico esté presente
        if all(x == 0 for x in [voltaje, corriente, potencia]):
            print("\n[ERROR] Debe ingresar al menos un parámetro numérico (voltaje, corriente o potencia)")
            return

        # Cálculo automático solo si hay múltiples parámetros
        parametros_numericos = sum(1 for x in [voltaje, corriente, potencia] if x > 0)
        if parametros_numericos >= 2:
            try:
                if voltaje > 0 and corriente > 0 and potencia == 0:
                    potencia = voltaje * corriente
                    print(f"-> Potencia calculada: {potencia:.2f} W")
                elif potencia > 0 and voltaje > 0 and corriente == 0:
                    corriente = potencia / voltaje
                    print(f"-> Corriente calculada: {corriente:.2f} A")
                elif potencia > 0 and corriente > 0 and voltaje == 0:
                    voltaje = potencia / corriente
                    print(f"-> Voltaje calculado: {voltaje:.2f} V")
            except ZeroDivisionError:
                print("Error: No se puede dividir por cero.")
                return
    else:
        # Para otros tipos de entrada/salida
        tipo_salidas = ""
        aplicacion = input("Aplicación / Uso (ej. 'iluminación LED'): ").strip()
        voltaje = _try_float(input("Voltaje de salida: ")) or 0
        corriente = _try_float(input("Corriente de salida (A): ")) or 0
        potencia = _try_float(input("Potencia total (W): ")) or 0

        # Verificar que al menos un parámetro numérico esté presente
        if all(x == 0 for x in [voltaje, corriente, potencia]):
            print("\n[ERROR] Debe ingresar al menos un parámetro numérico (voltaje, corriente o potencia)")
            return

    # --- Carga de catálogo ---
    cat = cargar_catalogo_expandido(RUTA_EXCEL, hoja="Sheet4", columna_expandir="voltaje_v")
    if cat.empty:
        print("No se pudieron cargar datos del catálogo. Revise la ruta o el formato del Excel.")
        return

    print("[DEBUG] Columnas disponibles en main:", cat.columns.tolist())

    # --- Filtrado ---
    res = calcular(cat, potencia=potencia, voltaje=voltaje, corriente=corriente, 
                   aplicacion=aplicacion, tipo_entrada_salida=entrada_salida, tipo_salidas=tipo_salidas)

    # --- Impresión de resultados ---
    if res.empty:
        print("\nNo se encontraron coincidencias para los criterios especificados.")
    else:
        print(f"\n=== RESULTADOS ENCONTRADOS ({len(res)} fuentes) ===")
        print(res.head(25).to_string(index=False))

        salida_archivo = "recomendaciones_fuentes.xlsx"
        try:
            with pd.ExcelWriter(salida_archivo, engine="openpyxl") as w:
                res.to_excel(w, index=False, sheet_name="Resultados")
            print(f"\nArchivo de resultados guardado como: {salida_archivo}")
        except Exception as e:
            print(f"\nNo se pudo guardar el archivo Excel de salida: {e}")

if __name__ == "__main__":
    main()