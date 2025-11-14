import pandas as pd
import re
import os
from difflib import SequenceMatcher

# CONFIGURACIÓN (sin cambios)
try:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))
except NameError:
    BASE_DIR = os.getcwd()
RUTA_EXCEL = os.path.join(BASE_DIR, "FuentesAlimentacion.xlsx")

# FUNCIONES DE UTILIDAD (sin cambios)
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

# Normalización avanzada para búsqueda inteligente (sin cambios)
def _norm_avanzada(s: str) -> str:
    s = (s or "").strip().lower()
    rep = {"á":"a","é":"e","í":"i","ó":"o","ú":"u","ü":"u","ñ":"n"}
    for k,v in rep.items():
        s = s.replace(k,v)
    
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
    
    s = re.sub(r'[^\w\s]', ' ', s)
    palabras = re.findall(r'\b[a-z0-9]+\b', s)
    
    palabras_filtradas = [p for p in palabras if p not in palabras_conexion and len(p) > 2]
    
    return ' '.join(palabras_filtradas)

# Agregar esta función en calc9.py

def extraer_voltaje_entrada_inversor(nombre_inversor):
    """
    Extrae el voltaje de entrada del inversor desde el nombre del modelo.
    Ejemplo: NTS-450-112UN -> 12V (descartando el primer dígito '1')
    """
    import re
    
    # Buscar patrones de voltaje en el nombre del inversor
    patrones = [
        r'-(\d{3})U',           # Ejemplo: -112UN, -224UN, -448UN
        r'-(\d{3})[A-Z]*$',     # Ejemplo: -112, -224, -448
        r'(\d{3})V',            # Ejemplo: 112V, 224V, 448V
    ]
    
    for patron in patrones:
        match = re.search(patron, nombre_inversor)
        if match:
            codigo_voltaje = match.group(1)
            
            if codigo_voltaje.isdigit() and len(codigo_voltaje) == 3:
                # Descarta el primer dígito y toma los últimos dos para el voltaje
                voltaje_str = codigo_voltaje[1:]  # Tomar últimos 2 dígitos
                try:
                    voltaje = int(voltaje_str)
                    return voltaje
                except ValueError:
                    continue
    
    # Si no encuentra patrón de 3 dígitos, buscar voltaje directo
    patron_directo = r'-(\d{2})V'
    match_directo = re.search(patron_directo, nombre_inversor)
    if match_directo:
        try:
            return int(match_directo.group(1))
        except ValueError:
            pass
    
    return None

# Función para recomendar baterías para inversores
def recomendar_baterias_para_inversor(potencia_w, horas_uso, ruta_baterias="DuracionBateriasAG.xlsx"):
    """
    Calcula y recomienda baterías para un inversor en base a la potencia y autonomía deseada.
    SOLO recomienda baterías que cumplan con la autonomía mínima requerida.
    """
    import pandas as pd
    import numpy as np

    try:
        df_bat = pd.read_excel(ruta_baterias, sheet_name="Baterias")
    except Exception as e:
        return pd.DataFrame({'Error': [f"No se pudo cargar el catálogo de baterías: {e}"]})

    # Normalizar nombres de columnas
    df_bat.columns = [c.strip() for c in df_bat.columns]

    # Función mejorada para calcular Wh
    def obtener_wh(fila):
        # Primero intentar obtener Wh directamente
        try:
            wh_cols = [col for col in fila.index if 'wh' in col.lower()]
            if wh_cols:
                wh_val = fila[wh_cols[0]]
                if pd.notna(wh_val):
                    wh = float(str(wh_val).replace(',', '.').strip())
                    if 10 <= wh <= 50000:  # Rango razonable para Wh
                        return wh
        except:
            pass
        
        # Si no se puede obtener Wh directamente, calcular desde V * Ah
        try:
            volt_cols = [col for col in fila.index if 'volt' in col.lower() and 'v' in col.lower()]
            ah_cols = [col for col in fila.index if 'ah' in col.lower() or 'corr' in col.lower()]
            
            if volt_cols and ah_cols:
                volt = float(str(fila[volt_cols[0]]).replace(',', '.').strip())
                ah = float(str(fila[ah_cols[0]]).replace(',', '.').strip())
                wh_calculado = volt * ah
                if 10 <= wh_calculado <= 50000:  # Rango razonable
                    return wh_calculado
        except:
            pass
        
        return np.nan

    # Calcular Wh para todas las baterías
    df_bat['Capacidad_Wh_Calculada'] = df_bat.apply(obtener_wh, axis=1)
    
    # Eliminar baterías sin capacidad calculada
    df_bat = df_bat[df_bat['Capacidad_Wh_Calculada'].notna()]
    
    if df_bat.empty:
        return pd.DataFrame({'Info': ['No se encontraron baterías con capacidad válida']})

    # Calcular autonomía REAL (Wh / W = horas)
    df_bat['Autonomia_Horas'] = df_bat['Capacidad_Wh_Calculada'] / potencia_w
    df_bat['Autonomia_Horas'] = df_bat['Autonomia_Horas'].round(2)

    # FILTRAR SOLO baterías que CUMPLAN con la autonomía mínima
    df_filtrado = df_bat[df_bat['Autonomia_Horas'] >= horas_uso].copy()
    
    if df_filtrado.empty:
        # Si no hay baterías que cumplan, buscar la que más se acerque
        df_bat['Diferencia'] = abs(df_bat['Autonomia_Horas'] - horas_uso)
        df_filtrado = df_bat.nsmallest(1, 'Diferencia')
    else:
        # Ordenar por autonomía (más cercana a la solicitada primero) y limitar a 7
        df_filtrado['Diferencia'] = abs(df_filtrado['Autonomia_Horas'] - horas_uso)
        df_filtrado = df_filtrado.sort_values('Diferencia').head(7)

    # Preparar columnas de resultado
    cols_disponibles = []
    for col in ['No. de parte', 'Modelo', 'Tipo', 'Voltaje (V)', 'Corriente (Ah)']:
        if col in df_filtrado.columns:
            cols_disponibles.append(col)
    
    cols_disponibles.extend(['Capacidad_Wh_Calculada', 'Autonomia_Horas'])
    
    df_result = df_filtrado[cols_disponibles].rename(columns={
        'Capacidad_Wh_Calculada': 'Capacidad (Wh)',
        'Autonomia_Horas': 'Autonomía (h)'
    })

    return df_result.reset_index(drop=True)

# NUEVA FUNCIÓN PARA CALCULAR ARREGLOS DE BATERÍAS
def calcular_arreglos_baterias(potencia_w, horas_uso, ruta_baterias="DuracionBateriasAG.xlsx", voltaje_inversor=None):
    """
    Calcula arreglos de baterías en serie/paralelo para alcanzar la autonomía requerida.
    Excluye baterías de Óxido de Plata, Níquel y Alcalinas.
    Limita la autonomía máxima a 12 horas.
    Excluye arreglos con solo una batería total (1S × 1P).
    """
    import pandas as pd
    import numpy as np

    try:
        df_bat = pd.read_excel(ruta_baterias, sheet_name="Baterias")
    except Exception as e:
        return pd.DataFrame({'Error': [f"No se pudo cargar el catálogo de baterías: {e}"]})

    # Normalizar nombres de columnas
    df_bat.columns = [c.strip() for c in df_bat.columns]

    # Tipos de baterías a EXCLUIR
    tipos_excluir = ['óxido de plata', 'oxido de plata', 'plata', 'níquel', 'nickel', 'níquel-cadmio', 
                    'nickel-cadmium', 'alcalina', 'alkaline', 'nicd', 'nimh']
    
    # Filtrar excluyendo los tipos no deseados
    if 'Tipo' in df_bat.columns:
        mask_excluir = df_bat['Tipo'].astype(str).str.lower().apply(
            lambda x: any(tipo in x.lower() for tipo in tipos_excluir)
        )
        df_bat = df_bat[~mask_excluir]
    
    # Función para obtener voltaje
    def obtener_voltaje(fila):
        try:
            volt_cols = [col for col in fila.index if 'volt' in col.lower() and 'v' in col.lower()]
            if volt_cols:
                volt_val = fila[volt_cols[0]]
                if pd.notna(volt_val):
                    volt = float(str(volt_val).replace(',', '.').strip())
                    if 1 <= volt <= 48:  # Rango razonable para voltaje de batería
                        return volt
        except:
            pass
        return np.nan

    # Función para obtener capacidad Ah
    def obtener_ah(fila):
        try:
            ah_cols = [col for col in fila.index if 'ah' in col.lower() or 'corr' in col.lower()]
            if ah_cols:
                ah_val = fila[ah_cols[0]]
                if pd.notna(ah_val):
                    ah = float(str(ah_val).replace(',', '.').strip())
                    if 0.1 <= ah <= 1000:  # Rango razonable para Ah
                        return ah
        except:
            pass
        return np.nan

    # Calcular valores para todas las baterías
    df_bat['Voltaje_V'] = df_bat.apply(obtener_voltaje, axis=1)
    df_bat['Capacidad_Ah'] = df_bat.apply(obtener_ah, axis=1)
    df_bat['Capacidad_Wh'] = df_bat['Voltaje_V'] * df_bat['Capacidad_Ah']
    
    # Eliminar baterías sin valores válidos
    df_bat = df_bat[(df_bat['Voltaje_V'].notna()) & (df_bat['Capacidad_Ah'].notna())]
    
    if df_bat.empty:
        return pd.DataFrame({'Info': ['No se encontraron baterías válidas para arreglos']})

    # Calcular energía total requerida (Wh)
    energia_requerida = potencia_w * horas_uso
    
    # Limitar autonomía máxima a 12 horas
    horas_maximas = min(horas_uso, 12)
    energia_maxima = potencia_w * horas_maximas
    
    resultados_arreglos = []
    
    for _, bateria in df_bat.iterrows():
        voltaje_bat = bateria['Voltaje_V']
        capacidad_ah_bat = bateria['Capacidad_Ah']
        capacidad_wh_bat = bateria['Capacidad_Wh']
        modelo = bateria.get('No. de parte', bateria.get('Modelo', 'N/A'))
        tipo = bateria.get('Tipo', 'N/A')
        
        # Si se especifica voltaje del inversor, usar ese como referencia
        voltaje_target = voltaje_inversor if voltaje_inversor else voltaje_bat
        
        # Calcular número de baterías en serie para alcanzar el voltaje target
        if voltaje_target > 0 and voltaje_bat > 0:
            baterias_serie = max(1, round(voltaje_target / voltaje_bat))
        else:
            baterias_serie = 1
        
        # Calcular energía por conjunto serie
        energia_por_serie = capacidad_wh_bat * baterias_serie
        
        # Calcular número de conjuntos en paralelo para alcanzar la energía requerida
        if energia_por_serie > 0:
            conjuntos_paralelo = max(1, int(np.ceil(energia_maxima / energia_por_serie)))
        else:
            conjuntos_paralelo = 1
        
        # Calcular total de baterías y autonomía real
        total_baterias = baterias_serie * conjuntos_paralelo
        energia_total = capacidad_wh_bat * total_baterias
        autonomia_real = energia_total / potencia_w
        
        # FILTRAR: Excluir arreglos con solo una batería total (1S × 1P)
        if total_baterias <= 1:
            continue
        
        # Solo considerar arreglos que cumplan con al menos el 80% de la autonomía requerida
        if autonomia_real >= horas_uso * 0.8:
            resultados_arreglos.append({
                'Modelo': modelo,
                'Tipo': tipo,
                'Voltaje_Bateria': voltaje_bat,
                'Capacidad_Ah': capacidad_ah_bat,
                'Capacidad_Wh': capacidad_wh_bat,
                'Baterias_Serie': baterias_serie,
                'Conjuntos_Paralelo': conjuntos_paralelo,
                'Total_Baterias': total_baterias,
                'Energia_Total_Wh': energia_total,
                'Autonomia_Real_h': round(autonomia_real, 2),
                'Voltaje_Sistema': voltaje_bat * baterias_serie
            })
    
    if not resultados_arreglos:
        return pd.DataFrame({'Info': ['No se encontraron arreglos válidos que cumplan con la autonomía requerida (mínimo 2 baterías)']})
    
    # Ordenar por total de baterías (menor primero) y limitar a 5 mejores opciones
    df_resultados = pd.DataFrame(resultados_arreglos)
    df_resultados = df_resultados.sort_values('Total_Baterias').head(5)
    
    return df_resultados.reset_index(drop=True)

# Función para calcular similitud entre cadenas (sin cambios)
def _calcular_similitud(a: str, b: str) -> float:
    if not a or not b:
        return 0.0
    return SequenceMatcher(None, a, b).ratio()

# Función para buscar coincidencias con umbral de similitud (sin cambios)
def _buscar_coincidencias(texto: str, busqueda: str, umbral=0.7) -> bool:
    if not texto or not busqueda:
        return False
    
    texto_norm = _norm_avanzada(texto)
    busqueda_norm = _norm_avanzada(busqueda)
    
    if not texto_norm or not busqueda_norm:
        return False
    
    def dividir_terminos(texto):
        separadores = [',', ';', '/', '|', ' y ', ' e ']
        texto_para_dividir = texto
        for sep in separadores:
            texto_para_dividir = texto_para_dividir.replace(sep, ',')
        return [t.strip() for t in texto_para_dividir.split(',') if t.strip()]
    
    terminos_texto = dividir_terminos(texto_norm)
    terminos_busqueda = dividir_terminos(busqueda_norm)
    
    for termino_b in terminos_busqueda:
        for termino_t in terminos_texto:
            if termino_b in termino_t:
                return True
            similitud = _calcular_similitud(termino_b, termino_t)
            if similitud >= umbral:
                return True
    
    if busqueda_norm in texto_norm:
        return True
    
    similitud_completa = _calcular_similitud(texto_norm, busqueda_norm)
    return similitud_completa >= umbral

# CARGA DEL CATÁLOGO - MEJORADA PARA DETECTAR COLUMNAS DE VARIABLE/FIJO
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

    # --- Renombrado automático MEJORADO ---
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
        # NUEVO: Detectar columnas relacionadas con variable/fijo
        elif re.search(r"variable|var|tipo.*salida|salida.*tipo", c):
            renombres[c] = "variable"
        elif re.search(r"fijo|fija|fixed", c):
            renombres[c] = "fijo"

    df.rename(columns=renombres, inplace=True)

    print("[DEBUG] Columnas después de renombrar:", df.columns.tolist())
    
    # NUEVO: Si no existe columna 'variable', intentar crearla a partir de otras columnas
    if 'variable' not in df.columns:
        print("[INFO] No se encontró columna 'variable'. Buscando en otras columnas...")
        # Buscar en toda la DataFrame indicadores de variable/fijo
        for col in df.columns:
            if any(word in str(df[col].iloc[0]).lower() for word in ['variable', 'var', 'ajustable', 'adjustable']) if len(df) > 0 else False:
                print(f"[INFO] Detectada columna con datos variables: {col}")
                df['variable'] = '1'
                break
            elif any(word in str(df[col].iloc[0]).lower() for word in ['fijo', 'fija', 'fixed']) if len(df) > 0 else False:
                print(f"[INFO] Detectada columna con datos fijos: {col}")
                df['variable'] = '0'
                break
    
    return df

# Función para verificar voltaje (sin cambios)
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

# NUEVA FUNCIÓN PARA DETECTAR SI UNA FUENTE ES VARIABLE
def _es_fuente_variable(fuente_data):
    """
    Detecta si una fuente es variable basándose en múltiples criterios
    """
    if pd.isna(fuente_data) or not fuente_data:
        return False
    
    texto = str(fuente_data).lower()
    
    # Patrones que indican fuente variable
    patrones_variable = [
        r'\bvariable\b',
        r'\bvar\b',
        r'\bajustable\b',
        r'\badjustable\b',
        r'\bregulable\b',
        r'\bvariable\s*/\s*fija',
        r'\brango\b',
        r'\bvoltaje\s*ajustable',
        r'\bsalida\s*variable'
    ]
    
    for patron in patrones_variable:
        if re.search(patron, texto):
            return True
    
    return False

# NUEVA FUNCIÓN PARA DETECTAR SI UNA FUENTE ES FIJA
def _es_fuente_fija(fuente_data):
    """
    Detecta si una fuente es fija basándose en múltiples criterios
    """
    if pd.isna(fuente_data) or not fuente_data:
        return False
    
    texto = str(fuente_data).lower()
    
    # Patrones que indican fuente fija
    patrones_fija = [
        r'\bfijo\b',
        r'\bfija\b',
        r'\bfixed\b',
        r'\best[áa]ndar\b',
        r'\bvoltaje\s*fijo',
        r'\bsalida\s*fija'
    ]
    
    for patron in patrones_fija:
        if re.search(patron, texto):
            return True
    
    return False

# FUNCIÓN DE FILTRADO MEJORADA PARA VARIABLE/FIJO
def filtrar_por_un_parametro(cat: pd.DataFrame, potencia: float, voltaje: float, corriente: float, aplicacion: str, tipo_entrada_salida: str, tipo_salidas: str):
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
                datos = datos[pd.to_numeric(datos["corriente_a"], errors='coerce')
                              .between(corriente * 0.7, corriente * 1.3).fillna(False)]
                print(f"[INFO] Después de filtrar por corriente: {len(datos)} fuentes")
            
        elif potencia > 0:
            print(f"[INFO] Filtrando por potencia: {potencia}W")
            if 'potencia_w' in datos.columns:
                datos = datos[pd.to_numeric(datos["potencia_w"], errors='coerce')
                              .between(potencia * 0.7, potencia * 1.3).fillna(False)]
                print(f"[INFO] Después de filtrar por potencia: {len(datos)} fuentes")

    # MEJORA: Aplicar filtro de variable/fijo también en búsqueda simple
    if tipo_salidas:
        print(f"[INFO] Aplicando filtro de salida: tipo_salidas={tipo_salidas}")
        datos = _aplicar_filtro_variable_fijo(datos, tipo_salidas)

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
                lambda x: _buscar_coincidencias(x, aplicacion, 0.6)
            )
            datos = datos[mask]
            print(f"[INFO] Después de filtrar por aplicación '{aplicacion}': {len(datos)} fuentes")

    return datos

# NUEVA FUNCIÓN PARA APLICAR FILTRO VARIABLE/FIJO
def _aplicar_filtro_variable_fijo(datos: pd.DataFrame, tipo_salidas: str):
    """
    Aplica filtro de variable/fijo usando múltiples estrategias
    """
    if datos.empty:
        return datos
        
    print(f"[DEBUG] Aplicando filtro variable/fijo. Tipo: {tipo_salidas}")
    print(f"[DEBUG] Columnas disponibles: {datos.columns.tolist()}")
    
    # Estrategia 1: Usar columna 'variable' si existe
    if 'variable' in datos.columns:
        print("[DEBUG] Usando columna 'variable' para filtrado")
        datos['variable'] = datos['variable'].astype(str)
        
        if tipo_salidas == "1":  # Variable
            # Buscar '1', 'si', 'true', o patrones que indiquen variable
            mask = (
                datos['variable'].str.contains(r'1|si|true|variable', case=False, na=False) |
                datos['variable'].apply(_es_fuente_variable)
            )
            datos_filtrados = datos[mask]
            print(f"[INFO] Fuentes variables encontradas: {len(datos_filtrados)}")
            return datos_filtrados
            
        elif tipo_salidas == "0":  # Fijo
            # Buscar '0', 'no', 'false', o patrones que indiquen fijo, o valores nulos
            mask = (
                datos['variable'].str.contains(r'0|no|false|fijo|fija', case=False, na=False) |
                datos['variable'].apply(_es_fuente_fija) |
                datos['variable'].isna()
            )
            datos_filtrados = datos[mask]
            print(f"[INFO] Fuentes fijas encontradas: {len(datos_filtrados)}")
            return datos_filtrados
    
    # Estrategia 2: Buscar en todas las columnas textuales
    print("[DEBUG] Buscando en columnas textuales para determinar variable/fijo")
    columnas_texto = ['fuente', 'entrada_salida', 'usos', 'descripcion', 'notas']
    columnas_disponibles = [col for col in columnas_texto if col in datos.columns]
    
    if columnas_disponibles:
        print(f"[DEBUG] Buscando en columnas: {columnas_disponibles}")
        
        def es_variable_fila(fila):
            for col in columnas_disponibles:
                if pd.notna(fila[col]) and _es_fuente_variable(fila[col]):
                    return True
            return False
        
        def es_fija_fila(fila):
            for col in columnas_disponibles:
                if pd.notna(fila[col]) and _es_fuente_fija(fila[col]):
                    return True
            return False
        
        if tipo_salidas == "1":  # Variable
            mask = datos.apply(es_variable_fila, axis=1)
            datos_filtrados = datos[mask]
            print(f"[INFO] Fuentes variables (por análisis textual): {len(datos_filtrados)}")
            return datos_filtrados
            
        elif tipo_salidas == "0":  # Fijo
            mask = datos.apply(es_fija_fila, axis=1)
            # Incluir también filas que no son variables
            mask_no_variable = ~datos.apply(es_variable_fila, axis=1)
            datos_filtrados = datos[mask | mask_no_variable]
            print(f"[INFO] Fuentes fijas (por análisis textual): {len(datos_filtrados)}")
            return datos_filtrados
    
    print("[WARNING] No se pudo aplicar filtro variable/fijo - columna no encontrada")
    return datos

# FUNCIÓN DE CÁLCULO MODIFICADA - MEJORADA PARA VARIABLE/FIJO
def calcular(cat: pd.DataFrame, potencia: float, voltaje: float, corriente: float, aplicacion: str, tipo_salidas: str, tipo_entrada_salida: str = None):
    datos = cat.copy()
    if datos.empty:
        return pd.DataFrame()

    print(f"[DEBUG] Parámetros recibidos: V={voltaje}, A={corriente}, W={potencia}, App={aplicacion}, Tipo={tipo_entrada_salida}, Salidas={tipo_salidas}")

    # --- Verificar si es un caso de un solo parámetro ---
    parametros_numericos = sum(1 for x in [voltaje, corriente, potencia] if x > 0)
    
    if parametros_numericos == 1:
        # Usar el nuevo filtro para un solo parámetro
        datos = filtrar_por_un_parametro(cat, potencia, voltaje, corriente, aplicacion, tipo_entrada_salida, tipo_salidas)
    else:
        # --- Aplicar filtro de variable/fijo PRIMERO ---
        if tipo_salidas:
            datos = _aplicar_filtro_variable_fijo(datos, tipo_salidas)
            print(f"[INFO] Después de filtro variable/fijo: {len(datos)} fuentes")
    
        # --- (Filtros de texto) ---
        if tipo_entrada_salida:
            if 'entrada_salida' in datos.columns:
                datos['entrada_salida'] = datos['entrada_salida'].astype(str)
                te_norm = _norm(tipo_entrada_salida)
                datos = datos[datos['entrada_salida'].str.lower().str.contains(te_norm, na=False)]
                print(f"[INFO] Después de filtrar por entrada/salida: {len(datos)} fuentes")
                
        if aplicacion:
            if 'usos' in datos.columns:
                datos['usos'] = datos['usos'].astype(str)
                mask = datos['usos'].apply(
                    lambda x: _buscar_coincidencias(x, aplicacion, 0.6)
                )
                datos = datos[mask]
                print(f"[INFO] Filtro aplicación '{aplicacion}': {len(datos)} fuentes después del filtro")
                
        # --- Filtros numéricos ---
        if tipo_entrada_salida == "AC-DC":
            if voltaje > 0 and 'voltaje_v' in datos.columns:
                mask = datos['voltaje_v'].apply(
                    lambda x: _verificar_voltaje(x, voltaje, 0.2)
                )
                datos = datos[mask]
                print(f"[INFO] Después de filtrar por voltaje AC-DC: {len(datos)} fuentes")

            if corriente > 0 and 'corriente_a' in datos.columns:
                datos = datos[pd.to_numeric(datos["corriente_a"], errors='coerce').between(corriente * 0.8, corriente * 1.2).fillna(False)]
                print(f"[INFO] Después de filtrar por corriente AC-DC: {len(datos)} fuentes")
            
            if potencia > 0 and 'potencia_w' in datos.columns:
                datos = datos[pd.to_numeric(datos["potencia_w"], errors='coerce').between(potencia * 0.8, potencia * 1.2).fillna(False)]
                print(f"[INFO] Después de filtrar por potencia AC-DC: {len(datos)} fuentes")
                
        elif tipo_entrada_salida == "DC-AC":
            if voltaje > 0 and 'voltaje_v' in datos.columns:
                mask = datos['voltaje_v'].apply(
                    lambda x: _verificar_voltaje(x, voltaje, 0.2)
                )
                datos = datos[mask]
                print(f"[INFO] Después de filtrar por voltaje DC-AC: {len(datos)} fuentes")

            if corriente > 0 and 'corriente_a' in datos.columns:
                corriente_series = pd.to_numeric(datos["corriente_a"], errors='coerce')
                mask_corriente = (corriente_series >= corriente).fillna(False)
                datos = datos[mask_corriente]
                print(f"[INFO] Después de filtrar por corriente DC-AC: {len(datos)} fuentes")
            
            if potencia > 0 and 'potencia_w' in datos.columns:
                potencia_series = pd.to_numeric(datos["potencia_w"], errors='coerce')
                mask_potencia = ((potencia_series == potencia) | (potencia_series >= potencia * 1.2)).fillna(False)
                datos = datos[mask_potencia]
                print(f"[INFO] Después de filtrar por potencia DC-AC: {len(datos)} fuentes")
    
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

# ---------------------------------------------------
# NUEVA FUNCIÓN: cargar_catalogo_baterias
# ---------------------------------------------------
def cargar_catalogo_baterias(ruta_excel_baterias, hoja="Baterias"):
    """
    Carga y normaliza la hoja de baterías del Excel.
    Devuelve DataFrame con columnas mínimas: ['Modelo', 'Tipo', 'Voltaje', 'Corriente (Ah)', 'Capacidad_Ah', 'Capacidad de la Batería (Wh)']
    """
    try:
        df = pd.read_excel(ruta_excel_baterias, sheet_name=hoja, dtype=str)
    except FileNotFoundError:
        print(f"[ERROR] No se pudo encontrar el archivo de baterías: '{ruta_excel_baterias}'")
        return pd.DataFrame()
    except Exception as e:
        print(f"[ERROR] No se pudo leer el archivo de baterías '{ruta_excel_baterias}': {e}")
        return pd.DataFrame()
    
    # Mostrar columnas detectadas
    print("[INFO] Columnas originales de baterías:", df.columns.tolist())

    # Normalizar nombres simples (no forzamos mayúsculas/minúsculas para mantener nombres originales)
    # Intentar mapear columnas comunes
    col_map = {}
    cols_lower = {c.lower(): c for c in df.columns}

    # Mapeo manual si existen variantes
    for key in cols_lower:
        if 'no' in key and 'parte' in key:
            col_map[cols_lower[key]] = 'No. de parte'
        elif 'model' in key or 'parte' in key:
            col_map[cols_lower[key]] = 'No. de parte'
        elif 'volt' in key and ('v' in key or 'volt' in key):
            col_map[cols_lower[key]] = 'Voltaje'
        elif 'corr' in key or 'ah' in key:
            col_map[cols_lower[key]] = 'Corriente (Ah)'
        elif 'capaci' in key and 'wh' in key:
            col_map[cols_lower[key]] = 'Capacidad de la Batería (Wh)'
        elif 'capaci' in key and 'ah' in key:
            col_map[cols_lower[key]] = 'Capacidad_Ah'
        elif 'tipo' in key:
            col_map[cols_lower[key]] = 'Tipo'
        elif 'uso' in key:
            col_map[cols_lower[key]] = 'Uso'

    df = df.rename(columns=col_map)

    # Asegurar columnas esperadas
    if 'No. de parte' not in df.columns and 'Modelo' not in df.columns:
        # crear columna Modelo basada en la primera columna
        first_col = df.columns[0]
        df['No. de parte'] = df[first_col].astype(str)

    # Crear columna Capacidad_Ah si no existe, priorizando 'Corriente (Ah)'
    if 'Capacidad_Ah' not in df.columns:
        if 'Corriente (Ah)' in df.columns and 'Voltaje' in df.columns:
            # Convertir a numéricas
            df['Corriente (Ah)'] = pd.to_numeric(df['Corriente (Ah)'], errors='coerce')
            df['Voltaje'] = pd.to_numeric(df['Voltaje'], errors='coerce')
            # Calcular Wh si falta
            if 'Capacidad de la Batería (Wh)' not in df.columns or df['Capacidad de la Batería (Wh)'].isna().all():
                try:
                    df['Capacidad de la Batería (Wh)'] = (df['Corriente (Ah)'] * df['Voltaje']).where(df['Corriente (Ah)'].notna(), None)
                except:
                    pass
            df['Capacidad_Ah'] = df['Corriente (Ah)']
        else:
            # intentar inferir Capacidad_Ah desde columna que tenga 'Wh' y 'Voltaje'
            if 'Capacidad de la Batería (Wh)' in df.columns and 'Voltaje' in df.columns:
                df['Voltaje'] = pd.to_numeric(df['Voltaje'], errors='coerce')
                df['Capacidad de la Batería (Wh)'] = pd.to_numeric(df['Capacidad de la Batería (Wh)'], errors='coerce')
                df['Capacidad_Ah'] = (df['Capacidad de la Batería (Wh)'] / df['Voltaje']).where(df['Voltaje'].notna(), None)
            else:
                df['Capacidad_Ah'] = None

    # Normalizar Voltaje numérico donde se pueda
    if 'Voltaje' in df.columns:
        df['Voltaje'] = pd.to_numeric(df['Voltaje'], errors='coerce')

    # Devolver DataFrame con columnas seleccionadas
    cols_out = []
    for c in ['Tipo', 'Uso', 'No. de parte', 'Voltaje', 'Corriente (Ah)', 'Capacidad_Ah', 'Capacidad de la Batería (Wh)']:
        if c in df.columns:
            cols_out.append(c)
    df_out = df[cols_out].copy()
    # Renombrar 'No. de parte' a 'No. de parte' (o Model) conservado
    return df_out.reset_index(drop=True)

# (rest of helpers remain unchanged; they are above and below)
# ... (the rest of the file already included previously)
# (end of file)
