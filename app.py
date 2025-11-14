# app_fuentes.py (actualizado)
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
HOJA_BATERIAS = "Baterias"

@app.route('/')
def index():
    """Página principal"""
    return render_template('index.html')

@app.route('/buscar', methods=['POST'])
def buscar_fuentes():
    """Endpoint para buscar fuentes y (opcional) baterías para inversores"""
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

        return jsonify({
            'success': True,
            'resultados': fuentes,
            'total': len(fuentes),
            'horas_autonomia': horas_autonomia
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
        }
        
        return jsonify(info)
    except Exception as e:
        return jsonify({'error': str(e)})

if __name__ == '__main__':
    logger.info("🚀 Iniciando servidor de Fuentes de Alimentación (con integración de baterías)...")
    app.run(debug=False, host='0.0.0.0', port=5000)
