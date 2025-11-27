# app_fuentes.py (completo y corregido)
from flask import Flask, render_template, request, jsonify
import pandas as pd
import logging
import os
from calc9 import (
    cargar_catalogo_expandido,
    calcular,
    _try_float,
    _norm,
    RUTA_EXCEL,
    recomendar_baterias_para_inversor,
    cargar_catalogo_baterias,
)

# Configurar logging
logging.basicConfig(level=logging.DEBUG)
logger = logging.getLogger(__name__)

app = Flask(__name__)

# Ruta del archivo de baterías (en la misma carpeta del proyecto)
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
RUTA_BATERIAS = os.path.join(BASE_DIR, "DuracionBateriasAG.xlsx")
RUTA_KITS_LED = os.path.join(BASE_DIR, "KITS ILUMINA-FUENTES.xlsx")
HOJA_BATERIAS = "Baterias"

@app.route('/')
def index():
    """Página principal"""
    return render_template('index2.html')

@app.route('/buscar', methods=['POST'])
def buscar_fuentes():
    """Endpoint para buscar fuentes y (opcional) kits LED"""
    try:
        data = request.get_json()
        logger.info(f"Datos recibidos: {data}")

        # Obtener parámetros
        tipo_entrada_salida = data.get('tipo_entrada_salida', '').strip()
        salida_variable = data.get('salida_variable', '').strip()
        aplicacion = data.get('aplicacion', '').strip()
        voltaje_str = data.get('voltaje', '').strip()
        corriente_str = data.get('corriente', '').strip()
        potencia_str = data.get('potencia', '').strip()
        horas_autonomia_str = data.get('horas_autonomia', '').strip() if data.get('horas_autonomia') is not None else ''
        tira_led = data.get('tira_led', '').strip()

        # Convertir valores numéricos
        voltaje = _try_float(voltaje_str) or 0
        corriente = _try_float(corriente_str) or 0
        potencia = _try_float(potencia_str) or 0
        horas_autonomia = _try_float(horas_autonomia_str) or 0

        # Determinar tipo de salida (para AC-DC)
        tipo_salidas = ""
        if tipo_entrada_salida == "AC-DC" and salida_variable:
            if salida_variable.lower() == 'si':
                tipo_salidas = "1"
            elif salida_variable.lower() == 'no':
                tipo_salidas = "0"
        elif tipo_entrada_salida == "DC-AC":
            tipo_salidas = "0"  # Salida fija por defecto para DC-AC

        # Si se especificó una tira LED, cargar sus parámetros
        kit_info = None
        if tira_led:
            try:
                df_tiras = pd.read_excel(RUTA_KITS_LED, sheet_name="Hoja 2")
                df_tiras.columns = [col.strip() for col in df_tiras.columns]
                
                tira_info = df_tiras[df_tiras['Modelo'] == tira_led].iloc[0]
                
                # Sobrescribir parámetros con los de la tira LED
                voltaje = float(tira_info.get('Voltaje_requerido_V', voltaje))
                corriente = float(tira_info.get('Corriente_requerida_A', 0))
                potencia = voltaje * corriente
                
                # Obtener información del conector y jack estándar
                kit_info = {
                    'modelo_tira': tira_led,
                    'voltaje_tira': voltaje,
                    'corriente_tira': corriente,
                    'jack': str(tira_info.get('JACK', '')),
                    'conector': str(tira_info.get('Conector', '')),
                    'potencia_tira': potencia
                }
                
                logger.info(f"Tira LED configurada: {kit_info}")
                
            except Exception as e:
                logger.error(f"Error procesando tira LED: {e}")

        # Validaciones
        if not tipo_entrada_salida:
            return jsonify({'success': False, 'error': 'Debe seleccionar el tipo de entrada/salida'})

        if all(x == 0 for x in [voltaje, corriente, potencia]):
            return jsonify({'success': False, 'error': 'Debe ingresar al menos un parámetro numérico (voltaje, corriente o potencia)'})

        # Cargar catálogo de fuentes
        cat = cargar_catalogo_expandido(RUTA_EXCEL, hoja="Sheet4", columna_expandir="voltaje_v")
        if cat.empty:
            return jsonify({'success': False, 'error': 'No se pudo cargar el catálogo'})

        logger.debug(f"Catálogo cargado. Columnas: {cat.columns.tolist()} - Total registros: {len(cat)}")

        # Calcular fuentes usando la función calcular
        resultados = calcular(
            cat=cat,
            potencia=potencia,
            voltaje=voltaje,
            corriente=corriente,
            aplicacion=aplicacion,
            tipo_salidas=tipo_salidas,
            tipo_entrada_salida=tipo_entrada_salida
        )

        # LIMITAR a máximo 7 fuentes
        resultados = resultados.head(7)

        # Preparar respuesta JSON
        fuentes = []

        for _, fuente in resultados.iterrows():
            fuente_data = {
                'modelo': fuente.get('fuente', 'N/A'),
                'tipo_entrada_salida': fuente.get('entrada_salida', 'N/A'),
                'aplicaciones': fuente.get('usos', 'N/A'),
            }

            # Obtener valores numéricos de la fuente
            potencia_val = _try_float(fuente.get('potencia_w', 0)) or 0
            
            # USAR VALORES ORIGINALES DEL EXCEL - NO CALCULAR
            if 'corriente_a' in fuente and pd.notna(fuente['corriente_a']):
                try:
                    corriente_val = _try_float(fuente['corriente_a'])
                    if corriente_val is not None:
                        fuente_data['corriente'] = corriente_val
                    else:
                        fuente_data['corriente'] = str(fuente['corriente_a']).strip()
                except:
                    fuente_data['corriente'] = str(fuente['corriente_a']).strip()
            else:
                fuente_data['corriente'] = 0

            if 'potencia_w' in fuente and pd.notna(fuente['potencia_w']):
                try:
                    potencia_val = _try_float(fuente['potencia_w'])
                    if potencia_val is not None:
                        fuente_data['potencia'] = potencia_val
                    else:
                        fuente_data['potencia'] = str(fuente['potencia_w']).strip()
                except:
                    fuente_data['potencia'] = str(fuente['potencia_w']).strip()
            else:
                fuente_data['potencia'] = 0

            # Manejar voltajes - mostrar el string original
            if 'voltaje_v' in fuente and pd.notna(fuente['voltaje_v']):
                fuente_data['voltaje'] = str(fuente['voltaje_v']).strip()
            else:
                fuente_data['voltaje'] = 'N/A'

            # 🔋 Si es inversor (DC-AC) y se solicitaron horas de autonomía, calcular baterías
            fuente_data['baterias'] = []
            fuente_data['arreglos_baterias'] = []
            
            if tipo_entrada_salida == "DC-AC" and horas_autonomia > 0 and fuente_data['potencia'] > 0:
                try:
                    logger.info(f"Buscando baterías para inversor '{fuente_data['modelo']}' - {fuente_data['potencia']}W por {horas_autonomia}h")
                    
                    # Baterías individuales
                    recomendaciones = recomendar_baterias_para_inversor(
                        potencia_w=float(fuente_data['potencia']),
                        horas_uso=float(horas_autonomia),
                        ruta_baterias=RUTA_BATERIAS
                    )
                    
                    if not recomendaciones.empty and 'Error' not in recomendaciones.columns and 'Info' not in recomendaciones.columns:
                        for _, bat in recomendaciones.iterrows():
                            bat_model = bat.get('No. de parte') or bat.get('Modelo', 'N/A')
                            bat_tipo = bat.get('Tipo', 'N/A')
                            bat_volt = bat.get('Voltaje (V)', 'N/A')
                            bat_ah = bat.get('Corriente (Ah)', bat.get('Capacidad_Ah', 'N/A'))
                            bat_wh = bat.get('Capacidad (Wh)', 'N/A')
                            bat_autonomia = bat.get('Autonomía (h)', bat.get('Autonomia_Horas', 'N/A'))
                            
                            fuente_data['baterias'].append({
                                'modelo': str(bat_model),
                                'tipo': str(bat_tipo),
                                'voltaje': bat_volt,
                                'corriente_ah': bat_ah,
                                'capacidad_wh': bat_wh,
                                'autonomia_horas': bat_autonomia
                            })
                    
                    # Arreglos de baterías
                    from calc9 import calcular_arreglos_baterias
                    
                    # Intentar extraer voltaje del inversor del nombre del modelo
                    voltaje_inversor = None
                    if 'voltaje_v' in fuente and pd.notna(fuente['voltaje_v']):
                        try:
                            voltaje_str = str(fuente['voltaje_v'])
                            if '-' in voltaje_str:
                                # Es un rango, tomar el valor máximo
                                voltaje_inversor = max([float(x.strip()) for x in voltaje_str.split('-')])
                            else:
                                voltaje_inversor = float(voltaje_str)
                        except:
                            pass
                    
                    arreglos = calcular_arreglos_baterias(
                        potencia_w=float(fuente_data['potencia']),
                        horas_uso=float(horas_autonomia),
                        ruta_baterias=RUTA_BATERIAS,
                        voltaje_inversor=voltaje_inversor
                    )
                    
                    if not arreglos.empty and 'Error' not in arreglos.columns and 'Info' not in arreglos.columns:
                        for _, arr in arreglos.iterrows():
                            fuente_data['arreglos_baterias'].append({
                                'modelo': arr.get('Modelo', 'N/A'),
                                'tipo': arr.get('Tipo', 'N/A'),
                                'voltaje_bateria': arr.get('Voltaje_Bateria', 'N/A'),
                                'capacidad_ah': arr.get('Capacidad_Ah', 'N/A'),
                                'baterias_serie': int(arr.get('Baterias_Serie', 1)),
                                'conjuntos_paralelo': int(arr.get('Conjuntos_Paralelo', 1)),
                                'total_baterias': int(arr.get('Total_Baterias', 1)),
                                'voltaje_sistema': arr.get('Voltaje_Sistema', 'N/A'),
                                'autonomia_horas': arr.get('Autonomia_Real_h', 'N/A')
                            })
                    
                except Exception as e:
                    logger.error(f"Error calculando baterías para {fuente_data['modelo']}: {e}")
                    fuente_data['baterias'].append({
                        'modelo': 'Error en cálculo',
                        'tipo': f'Error: {str(e)}',
                        'voltaje': 'N/A',
                        'corriente_ah': 'N/A', 
                        'capacidad_wh': 'N/A',
                        'autonomia_horas': 'N/A'
                    })

            fuentes.append(fuente_data)

        # Buscar kits recomendados si hay tira LED seleccionada - MOSTRAR TODOS LOS KITS
        kits_recomendados = []
        if tira_led:
            try:
                logger.info(f"Buscando kits para: {tira_led}")
                
                # Cargar Hoja 2 para compatibilidad
                df_compatibilidad = pd.read_excel(RUTA_KITS_LED, sheet_name="Hoja 2")
                df_compatibilidad.columns = [col.strip() for col in df_compatibilidad.columns]
                
                # Filtrar fuentes compatibles para esta tira LED
                fuentes_compatibles = df_compatibilidad[df_compatibilidad['Modelo'] == tira_led]
                
                logger.info(f"Fuentes compatibles encontradas: {len(fuentes_compatibles)}")
                
                # MOSTRAR TODOS LOS KITS DISPONIBLES
                for _, fuente_row in fuentes_compatibles.iterrows():
                    if pd.notna(fuente_row.get('Fuente Compatible')):
                        fuente_modelo = str(fuente_row['Fuente Compatible'])
                        corriente_fuente = float(fuente_row.get('Corriente de la fuente (A)', 0))
                        conector = str(fuente_row.get('Conector', '221-412'))
                        jack = str(fuente_row.get('JACK', 'DC5.1-JACK'))
                        cantidad_tiras = int(fuente_row.get('Cantidad de Tiras', 1))  # NUEVO: Obtener cantidad de tiras
                        
                        # Obtener información de la fuente desde el catálogo principal
                        fuente_info_catalogo = None
                        if not cat.empty and 'fuente' in cat.columns:
                            fuente_match = cat[cat['fuente'].astype(str) == fuente_modelo]
                            if not fuente_match.empty:
                                fuente_info_catalogo = fuente_match.iloc[0]
                        
                        # Crear información del kit - VERSIÓN MEJORADA CON CANTIDAD DE TIRAS
                        kit_data = {
                            'nombre_kit': f"Kit {tira_led} + {fuente_modelo}",
                            'fuente_compatible': fuente_modelo,
                            'corriente_fuente': corriente_fuente,
                            'conector_recomendado': conector,
                            'jack_recomendado': jack,
                            'cantidad_tiras': cantidad_tiras,  # NUEVO: Incluir cantidad de tiras
                            'voltaje_fuente': str(fuente_info_catalogo.get('voltaje_v', '12V')) if fuente_info_catalogo is not None else '12V',
                            'potencia_fuente': str(fuente_info_catalogo.get('potencia_w', f'{corriente_fuente * 12}W')) if fuente_info_catalogo is not None else f'{corriente_fuente * 12}W'
                        }
                        
                        kits_recomendados.append(kit_data)
                        logger.info(f"Kit creado: {kit_data['nombre_kit']} - Soporta {cantidad_tiras} tiras")
                
                logger.info(f"Kits finales para {tira_led}: {len(kits_recomendados)}")
                
            except Exception as e:
                logger.error(f"Error buscando kits LED: {e}")
                import traceback
                logger.error(traceback.format_exc())

        return jsonify({
            'success': True,
            'resultados': fuentes,
            'total': len(fuentes),
            'horas_autonomia': horas_autonomia,
            'tira_led_seleccionada': tira_led if tira_led else None,
            'kits_recomendados': kits_recomendados
        })

    except Exception as e:
        logger.error(f"Error en la búsqueda: {str(e)}")
        return jsonify({
            'success': False,
            'error': f'Error en la búsqueda: {str(e)}'
        })

@app.route('/tipos-entrada-salida')
def obtener_tipos_entrada_salida():
    """Obtener tipos de entrada/salida disponibles"""
    return jsonify({
        'success': True,
        'tipos': ['AC-DC', 'DC-AC', 'DC-DC']
    })

@app.route('/aplicaciones')
def obtener_aplicaciones():
    """Obtener TODAS las aplicaciones disponibles (sin filtro)"""
    try:
        cat = cargar_catalogo_expandido(RUTA_EXCEL, hoja="Sheet4")
        if cat.empty or 'usos' not in cat.columns:
            return jsonify({'success': True, 'aplicaciones': []})
        
        todas_aplicaciones = set()
        for uso in cat['usos']:
            if pd.notna(uso):
                aplicaciones = [app.strip() for app in str(uso).split(',')]
                todas_aplicaciones.update(aplicaciones)
        
        aplicaciones_ordenadas = sorted(list(todas_aplicaciones))
        return jsonify({'success': True, 'aplicaciones': aplicaciones_ordenadas})
        
    except Exception as e:
        logger.error(f"Error obteniendo aplicaciones: {e}")
        return jsonify({'success': True, 'aplicaciones': []})

@app.route('/aplicaciones/<tipo_fuente>')
def obtener_aplicaciones_por_tipo(tipo_fuente):
    """Obtener aplicaciones filtradas por tipo de fuente y salida variable"""
    try:
        salida_variable = request.args.get('salida_variable', '').strip().lower()
        cat = cargar_catalogo_expandido(RUTA_EXCEL, hoja="Sheet4")
        if cat.empty or 'usos' not in cat.columns or 'entrada_salida' not in cat.columns:
            return jsonify({'success': True, 'aplicaciones': []})
        
        # Filtrar por tipo de fuente
        cat_filtrado = cat[cat['entrada_salida'].astype(str).str.contains(tipo_fuente, case=False, na=False)]
        
        # Filtrar adicionalmente por salida variable si se especifica
        if salida_variable and 'variable' in cat_filtrado.columns:
            if salida_variable == 'si':
                cat_filtrado = cat_filtrado[pd.to_numeric(cat_filtrado['variable'], errors='coerce') == 1]
            elif salida_variable == 'no':
                cat_filtrado = cat_filtrado[(pd.to_numeric(cat_filtrado['variable'], errors='coerce') == 0) | (cat_filtrado['variable'].isnull())]
        
        todas_aplicaciones = set()
        for uso in cat_filtrado['usos']:
            if pd.notna(uso):
                aplicaciones = [app.strip() for app in str(uso).split(',')]
                todas_aplicaciones.update(aplicaciones)
        
        aplicaciones_ordenadas = sorted(list(todas_aplicaciones))
        logger.info(f"Aplicaciones para {tipo_fuente} (salida_variable={salida_variable}): {len(aplicaciones_ordenadas)} encontradas")
        return jsonify({'success': True, 'aplicaciones': aplicaciones_ordenadas})
        
    except Exception as e:
        logger.error(f"Error obteniendo aplicaciones para {tipo_fuente}: {e}")
        return jsonify({'success': True, 'aplicaciones': []})

@app.route('/tiras-led')
def obtener_tiras_led():
    """Obtener todas las tiras LED disponibles desde el Excel"""
    try:
        # Cargar el archivo Excel de kits de iluminación - HOJA 2
        df_tiras = pd.read_excel(RUTA_KITS_LED, sheet_name="Hoja 2")
        
        # Limpiar y normalizar los datos
        df_tiras.columns = [col.strip() for col in df_tiras.columns]
        
        # Filtrar solo filas que tienen modelo
        df_tiras = df_tiras[df_tiras['Modelo'].notna()]
        
        # Agrupar por modelo de tira LED
        modelos_unicos = df_tiras['Modelo'].dropna().unique()
        
        tiras = []
        for modelo in modelos_unicos:
            # Obtener la primera fila de cada modelo para información básica
            tira_data = df_tiras[df_tiras['Modelo'] == modelo].iloc[0]
            
            # Manejar valores NaN
            jack_val = tira_data.get('JACK', '')
            conector_val = tira_data.get('Conector', '')
            
            if pd.isna(jack_val):
                jack_val = ''
            else:
                jack_val = str(jack_val)
                
            if pd.isna(conector_val):
                conector_val = ''
            else:
                conector_val = str(conector_val)
            
            tira_info = {
                'modelo': str(modelo),
                'voltaje': float(tira_data.get('Voltaje_requerido_V', 0)),
                'corriente': float(tira_data.get('Corriente_requerida_A', 0)),
                'jack': jack_val,
                'conector': conector_val
            }
            tiras.append(tira_info)
        
        logger.info(f"Se cargaron {len(tiras)} modelos de tiras LED desde Hoja 2")
        return jsonify({'success': True, 'tiras_led': tiras})
        
    except Exception as e:
        logger.error(f"Error obteniendo tiras LED: {e}")
        return jsonify({'success': False, 'error': str(e)})

@app.route('/kit-led/<modelo_tira>')
def obtener_kit_led(modelo_tira):
    """Obtener información completa del kit para una tira LED específica"""
    try:
        # Cargar el archivo Excel - HOJA 2
        df_tiras = pd.read_excel(RUTA_KITS_LED, sheet_name="Hoja 2")
        
        # Limpiar y normalizar los datos
        df_tiras.columns = [col.strip() for col in df_tiras.columns]
        
        # Filtrar por modelo de tira LED
        tiras_filtradas = df_tiras[df_tiras['Modelo'] == modelo_tira]
        
        if tiras_filtradas.empty:
            return jsonify({'success': False, 'error': 'Modelo de tira LED no encontrado'})
        
        # Obtener información básica de la tira
        tira_info = tiras_filtradas.iloc[0]
        info_basica = {
            'modelo': str(modelo_tira),
            'voltaje': float(tira_info.get('Voltaje_requerido_V', 0)),
            'corriente': float(tira_info.get('Corriente_requerida_A', 0)),
            'jack': str(tira_info.get('JACK', '')),
            'conector': str(tira_info.get('Conector', ''))
        }
        
        # Obtener fuentes compatibles
        fuentes_compatibles = []
        for _, tira in tiras_filtradas.iterrows():
            if pd.notna(tira.get('Fuente Compatible')):
                fuente_info = {
                    'modelo_fuente': str(tira.get('Fuente Compatible', '')),
                    'corriente_fuente': float(tira.get('Corriente de la fuente (A)', 0)),
                    'max_tiras': int(tira.get('Cantidad de Tiras', 1))
                }
                fuentes_compatibles.append(fuente_info)
        
        return jsonify({
            'success': True, 
            'tira_led': info_basica,
            'fuentes_compatibles': fuentes_compatibles
        })
        
    except Exception as e:
        logger.error(f"Error obteniendo kit LED: {e}")
        return jsonify({'success': False, 'error': str(e)})

@app.route('/debug')
def debug():
    """Endpoint de debug"""
    try:
        cat = cargar_catalogo_expandido(RUTA_EXCEL, hoja="Sheet4")
        
        info = {
            'archivo_existe': os.path.exists(RUTA_EXCEL),
            'catalogo_cargado': not cat.empty,
            'total_fuentes': len(cat) if not cat.empty else 0,
            'columnas': cat.columns.tolist() if not cat.empty else [],
            'ruta_excel': RUTA_EXCEL,
            'ruta_baterias': RUTA_BATERIAS,
            'ruta_kits_led': RUTA_KITS_LED,
            'archivo_kits_existe': os.path.exists(RUTA_KITS_LED),
        }
        
        return jsonify(info)
    except Exception as e:
        return jsonify({'error': str(e)})

if __name__ == '__main__':
    logger.info("🚀 Iniciando servidor de Fuentes de Alimentación (con integración de baterías)...")
    app.run(debug=False, host='0.0.0.0', port=5000)
