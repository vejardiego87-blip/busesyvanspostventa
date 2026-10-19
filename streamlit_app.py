from datetime import datetime
from zoneinfo import ZoneInfo
from hashlib import sha256
import re
import csv
from io import StringIO, BytesIO
from PIL import Image
import streamlit as st
from generador import generar_informe
from base_datos import guardar_informe, listar_informes, respaldo_base, actualizar_informe, eliminar_informe

MODELOS = ['U9', 'U10', 'U12', 'D9', 'D10']
RESPONSABLES = {
    'Alejandro Vasquez Andrews': {'prefijo': 'AV', 'cargo': 'Jefe Asistencia Técnica Buses y Vans (iveco)', 'correo': 'avasqueza@andesmotor.cl'},
    'Ignacio Gutiérrez Gaete': {'prefijo': 'IG', 'cargo': 'Supervisor de Flotas Buses', 'correo': 'ignacio.gutierrez@andesmotor.cl'},
    'Patricio Gormaz Pardo': {'prefijo': 'PG', 'cargo': 'Ingeniero de Operaciones y Desarrollo Técnico', 'correo': 'pgormaz@andesmotor.cl'},
    'Dimas Toro Echeverria': {'prefijo': 'DT', 'cargo': 'Ingeniero de Operaciones y Desarrollo Técnico', 'correo': 'dimas.toro@andesmotor.cl'},
    'David Esparza': {'prefijo': 'DE', 'cargo': 'Supervisor de Flotas Buses', 'correo': 'desparza@andesmotor.cl'},
    'Danilo Flores Burgos': {'prefijo': 'DF', 'cargo': 'Ingeniero de Operaciones y Desarrollo Técnico', 'correo': 'danilo.flores@andesmotor.cl'},
}

st.set_page_config(page_title='Informes técnicos Foton', page_icon='🚌', layout='centered')
st.title('Informes técnicos Foton')
st.caption('Postventa · Andes Motor · Plantilla Word')
st.write('Completa los antecedentes, agrega las fotos y descarga el informe editable.')
st.info('Marca fija: FOTON. Solo se guardan datos y textos en Google Sheets; nunca las fotografías ni el Word.')

a,b = st.columns(2)
modelo = a.selectbox('Modelo *', ['Selecciona un modelo'] + MODELOS)
fecha = b.date_input('Fecha de evaluación', datetime.now(ZoneInfo('America/Santiago')).date())
vin = a.text_input('VIN', max_chars=17).strip().upper()
patente = b.text_input('Patente').strip().upper()
sistema = st.text_input('Sistema afectado *', placeholder='Ej.: puertas, carrocería, motor o baterías')
st.subheader('Motivo de admisión')
motivo = st.text_area('Síntoma o reclamo informado *', height=110)
st.subheader('Entidades')
participantes = st.text_area('Cliente, taller y otras personas participantes', placeholder='Cliente / operador:\nTaller / sucursal:')
entidades = participantes.strip()
st.subheader('Análisis')
analisis = st.text_area('Inspección, diagnóstico, intervención y resultado *', height=190, placeholder='Inspección realizada:\nDiagnóstico confirmado o pendiente:\nIntervención:\nResultado y pendientes:')
st.subheader('Evidencia')
archivos = st.file_uploader('Fotografías JPG o PNG (máximo 12; 10 MB por foto)', type=['jpg','jpeg','png'], accept_multiple_files=True)
fotos=[]
errores=[]
if len(archivos)>12: errores.append('Adjunta un máximo de 12 fotografías.')
for i,f in enumerate(archivos[:12]):
    contenido = f.getvalue()
    if len(contenido)>10*1024*1024:
        errores.append(f'{f.name}: supera 10 MB.'); continue
    try:
        with Image.open(BytesIO(contenido)) as imagen:
            imagen.verify()
    except Exception:
        errores.append(f'{f.name}: no es una imagen válida.'); continue
    with st.expander(f'Fotografía {i+1}: {f.name}', expanded=True):
        st.image(contenido, width=280)
        desc = st.text_input('Descripción del hallazgo', key=f'foto_{i}_{sha256(contenido).hexdigest()[:16]}')
        fotos.append((contenido,desc))
st.subheader('Responsable del informe')
responsable = st.selectbox('Responsable *', ['Selecciona un responsable'] + list(RESPONSABLES))
perfil = RESPONSABLES.get(responsable)
if perfil:
    st.text(f"{perfil['cargo']}\n{perfil['correo']}")
    st.caption(f"Correlativo {perfil['prefijo']}-0001, {perfil['prefijo']}-0002… Se asigna al guardar.")
datos=dict(modelo=modelo,fecha=fecha.strftime('%d-%m-%Y'),vin=vin,patente=patente,sistema=sistema,motivo=motivo,entidades=entidades,analisis=analisis,responsable=responsable,cargo=perfil['cargo'] if perfil else '',correo=perfil['correo'] if perfil else '')
firma=sha256(repr((datos,[(sha256(x).hexdigest(),y) for x,y in fotos])).encode()).hexdigest()
confirmado=st.checkbox('Revisé los antecedentes técnicos y las descripciones de las fotos.')
if st.button('Generar Word y registrar correlativo', type='primary'):
    st.session_state.pop('informe', None)
    if modelo not in MODELOS: errores.append('Selecciona un modelo Foton.')
    if perfil is None: errores.append('Selecciona al responsable del informe.')
    if not all(x.strip() for x in (sistema,motivo,analisis)): errores.append('Completa los campos obligatorios (*).')
    if not vin and not patente: errores.append('Ingresa al menos VIN o patente.')
    if vin and (len(vin)!=17 or not re.fullmatch('[A-HJ-NPR-Z0-9]{17}',vin)): errores.append('Revisa el VIN: debe tener 17 caracteres y no contener I, O ni Q.')
    if any(not d.strip() for _,d in fotos): errores.append('Agrega la descripción de cada fotografía.')
    if not confirmado: errores.append('Confirma la revisión de los antecedentes.')
    if errores:
        for e in errores: st.error(e)
    else:
        try:
            resultado, nuevo = guardar_informe(datos, fotos, perfil['prefijo'], firma, generar_informe)
            st.session_state.informe=(firma,resultado['documento'],resultado['nombre_archivo'],resultado['numero'])
            if not nuevo: st.info('Este contenido ya tiene correlativo. Se generó el Word con el mismo número.')
        except Exception as exc:
            st.error(f'No se pudo guardar el informe. Detalle: {exc}')
if 'informe' in st.session_state:
    firma_guardada,documento,nombre,numero=st.session_state.informe
    if firma_guardada==firma and confirmado and not errores:
        st.success(f'Informe {numero} registrado. Descarga el Word antes de cerrar la sesión.')
        st.download_button('Descargar Word', documento, nombre, mime='application/vnd.openxmlformats-officedocument.wordprocessingml.document')
    else:
        st.warning('Los datos cambiaron. Genera nuevamente el informe para descargar la versión actualizada.')

st.divider()
st.subheader('Historial de informes')
st.caption('Base Google Sheets · solo datos y textos. No se almacenan Word ni fotografías.')
try:
    registros = listar_informes()
    filtro = st.text_input('Buscar por número, patente, VIN, modelo o responsable')
    visibles = [r for r in registros if not filtro or filtro.casefold() in ' '.join(str(v) for v in r.values()).casefold()]
    columnas_visibles=['numero','fecha','modelo','vin','patente','sistema','responsable','estado']
    st.dataframe([{k:r.get(k,'') for k in columnas_visibles} for r in visibles], use_container_width=True, hide_index=True)
    if registros:
        salida = StringIO()
        escritor = csv.DictWriter(salida, fieldnames=list(registros[0]))
        escritor.writeheader()
        for registro in registros:
            escritor.writerow({k: ("'"+str(v) if str(v).lstrip().startswith(('=','+','-','@')) else v) for k,v in registro.items()})
        st.download_button('Exportar historial CSV', salida.getvalue().encode('utf-8-sig'), 'historial_foton.csv', 'text/csv')
        st.subheader('Consultar o corregir un informe')
        elegido = st.selectbox('Seleccionar número de informe', ['Selecciona un informe'] + [r['numero'] for r in visibles])
        if elegido != 'Selecciona un informe':
            registro = next(r for r in visibles if r['numero'] == elegido)
            st.caption(f"Responsable: {registro['responsable']} · Registrado: {registro['creado_en']}")
            with st.form(f'editar_{elegido}'):
                cambios={}
                for campo in ['fecha','modelo','vin','patente','sistema','motivo','entidades','analisis']:
                    cambios[campo]=st.text_area(campo.capitalize(), value=registro.get(campo,''), height=120 if campo in ('motivo','entidades','analisis') else 68)
                guardar=st.form_submit_button('Guardar correcciones')
            if guardar:
                actualizar_informe(elegido, cambios)
                st.success('Informe actualizado en Google Sheets.')
                st.rerun()
            with st.expander('Eliminar informe incorrecto'):
                st.warning('La eliminación será lógica: el registro no aparecerá en el historial activo, pero se conservará para auditoría y su número no se reutilizará.')
                confirmar=st.checkbox(f'Confirmo que deseo eliminar {elegido}', key=f'confirmar_{elegido}')
                if st.button('Eliminar informe', disabled=not confirmar):
                    eliminar_informe(elegido)
                    st.success('Informe marcado como eliminado.')
                    st.rerun()
    if st.button('Preparar respaldo completo CSV'):
        st.session_state.respaldo = respaldo_base()
    if 'respaldo' in st.session_state:
        st.download_button('Descargar respaldo de textos', st.session_state.respaldo, 'respaldo_informes_foton.csv', 'text/csv')
except Exception as exc:
    st.error(f'No se pudo abrir el historial: {exc}')