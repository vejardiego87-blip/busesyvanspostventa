"""Persistencia de informes técnicos en Google Sheets (solo textos)."""
import json
import re
from datetime import datetime
from io import BytesIO
from threading import RLock
from zoneinfo import ZoneInfo

import gspread
import streamlit as st
from google.oauth2.service_account import Credentials

LOCK = RLock()  # Protege hilos de una instancia, no distintas instancias de Streamlit.
COLUMNAS = ['numero','fecha','modelo','vin','patente','marca','sistema','motivo','entidades','analisis','responsable','cargo','correo','firma','creado_en','actualizado_en','estado']
CONTADORES = ['prefijo','ultimo_numero']


def _ahora():
    return datetime.now(ZoneInfo('America/Santiago')).strftime('%Y-%m-%d %H:%M:%S')


@st.cache_resource(show_spinner=False)
def _hojas():
    if 'gcp_service_account' in st.secrets:
        credenciales = dict(st.secrets['gcp_service_account'])
    elif 'gcp' in st.secrets and 'credentials_json' in st.secrets['gcp']:
        credenciales = json.loads(st.secrets['gcp']['credentials_json'])
    else:
        raise RuntimeError('Faltan credenciales [gcp_service_account] o [gcp].credentials_json en Secrets.')
    credenciales['private_key'] = credenciales['private_key'].replace('\\n', '\n')
    scopes = ['https://www.googleapis.com/auth/spreadsheets']
    cliente = gspread.authorize(Credentials.from_service_account_info(credenciales, scopes=scopes))
    libro = cliente.open_by_key(st.secrets['google_sheets']['spreadsheet_id'])
    informes = libro.worksheet('Informes')
    correlativos = libro.worksheet('Correlativos')
    return informes, correlativos


def _preparar():
    informes, correlativos = _hojas()
    encabezados = informes.row_values(1)
    if not encabezados:
        informes.update(range_name='A1:Q1', values=[COLUMNAS])
    elif encabezados != COLUMNAS:
        # No sobrescribir columnas existentes, porque podría perderse información.
        raise RuntimeError('Los encabezados de la pestaña Informes no coinciden con los requeridos. Usa el botón "Copiar encabezados" de las instrucciones y reemplaza únicamente la fila 1 después de revisar si existen datos.')
    encabezados_c = correlativos.row_values(1)
    if not encabezados_c:
        correlativos.update(range_name='A1:B1', values=[CONTADORES])
    elif [x.lower() for x in encabezados_c[:2]] not in ([x.lower() for x in CONTADORES], ['prefijo','último número'], ['prefijo','ultimo_numero']):
        raise RuntimeError('La hoja Correlativos debe comenzar con prefijo y ultimo_numero.')
    return informes, correlativos


def listar_informes(incluir_eliminados=False):
    informes, _ = _preparar()
    filas = informes.get_all_records(expected_headers=COLUMNAS, numericise_ignore=['all'])
    filas = [{k: str(r.get(k, '') or '') for k in COLUMNAS} for r in filas if r.get('numero')]
    if not incluir_eliminados:
        filas = [r for r in filas if r['estado'].lower() != 'eliminado']
    return list(reversed(filas))


def _siguiente(prefijo, informes, correlativos):
    valores = correlativos.get_all_values()
    fila_contador = next((i for i, fila in enumerate(valores[1:], 2) if fila and fila[0].strip().upper() == prefijo), None)
    valor = 0
    if fila_contador:
        try:
            valor = int(valores[fila_contador-1][1])
        except (ValueError, IndexError):
            pass
    # Considerar también números de informes eliminados para no reutilizarlos.
    for fila in informes.get_all_records(numericise_ignore=['all']):
        numero = str(fila.get('numero', ''))
        if re.fullmatch(rf'{re.escape(prefijo)}-\d+', numero):
            valor = max(valor, int(numero.split('-')[-1]))
    return valor + 1, fila_contador


def guardar_informe(datos, fotos, prefijo, firma, generar_informe):
    """Devuelve ({documento, nombre_archivo, numero}, nuevo).

    El documento Word se genera en memoria; nunca se sube a Google Sheets.
    """
    with LOCK:
        informes, correlativos = _preparar()
        registros = listar_informes(incluir_eliminados=True)
        existente = next((r for r in registros if r['firma'] == firma and r['estado'].lower() != 'eliminado'), None)
        if existente:
            numero, nuevo = existente['numero'], False
        else:
            siguiente, fila_contador = _siguiente(prefijo, informes, correlativos)
            numero, nuevo = f'{prefijo}-{siguiente:04d}', True
        datos_doc = dict(datos, numero=numero)
        # Contrato del generador original: generar_informe(datos, fotos).
        resultado = generar_informe(datos_doc, fotos)
        if isinstance(resultado, dict):
            documento = resultado.get('documento')
            nombre_archivo = resultado.get('nombre_archivo', f'Informe_{numero}.docx')
        elif isinstance(resultado, (tuple, list)) and len(resultado) >= 2:
            documento, nombre_archivo = resultado[:2]
        else:
            documento, nombre_archivo = resultado, f'Informe_{numero}.docx'
        if hasattr(documento, 'getvalue'):
            documento = documento.getvalue()
        elif hasattr(documento, 'read'):
            documento = documento.read()
        if not isinstance(documento, bytes):
            raise TypeError('generar_informe debe devolver bytes, un BytesIO, un dict o una tupla con el Word.')
        if nuevo:
            momento = _ahora()
            registro = dict(datos, numero=numero, marca='FOTON', firma=firma, creado_en=momento, actualizado_en=momento, estado='Activo')
            informes.append_row([str(registro.get(k, '')) for k in COLUMNAS], value_input_option='RAW')
            # Contador se actualiza después del guardado, y el próximo cálculo también revisa el historial.
            if fila_contador:
                correlativos.update_cell(fila_contador, 2, siguiente)
            else:
                correlativos.append_row([prefijo, siguiente], value_input_option='RAW')
        return {'documento': documento, 'nombre_archivo': nombre_archivo, 'numero': numero}, nuevo


def actualizar_informe(numero, cambios):
    permitidos = {'fecha','modelo','vin','patente','sistema','motivo','entidades','analisis'}
    with LOCK:
        hoja, _ = _preparar()
        filas = hoja.get_all_values()
        for indice, fila in enumerate(filas[1:], 2):
            if fila and fila[0] == numero:
                actual = dict(zip(COLUMNAS, fila + [''] * len(COLUMNAS)))
                if actual['estado'].lower() == 'eliminado':
                    raise ValueError('No se puede editar un informe eliminado.')
                actual.update({k: str(v) for k, v in cambios.items() if k in permitidos})
                actual['actualizado_en'] = _ahora()
                # Firma anterior no debe impedir guardar una versión corregida.
                actual['firma'] = ''
                hoja.update(range_name=f'A{indice}:Q{indice}', values=[[actual[k] for k in COLUMNAS]])
                return
        raise ValueError('No se encontró el informe.')


def eliminar_informe(numero):
    """Eliminación lógica: conserva textos y correlativo para auditoría."""
    with LOCK:
        hoja, _ = _preparar()
        filas = hoja.get_all_values()
        for indice, fila in enumerate(filas[1:], 2):
            if fila and fila[0] == numero:
                hoja.update(range_name=f'Q{indice}', values=[['Eliminado']])
                hoja.update(range_name=f'P{indice}', values=[[_ahora()]])
                return
        raise ValueError('No se encontró el informe.')


def respaldo_base():
    """Respaldo CSV (solo textos); no es un archivo SQLite."""
    import csv
    from io import StringIO
    salida = StringIO()
    escritor = csv.DictWriter(salida, fieldnames=COLUMNAS)
    escritor.writeheader()
    for fila in listar_informes(incluir_eliminados=True):
        escritor.writerow({k: "'" + v if v.lstrip().startswith(('=', '+', '-', '@')) else v for k, v in fila.items()})
    return salida.getvalue().encode('utf-8-sig')