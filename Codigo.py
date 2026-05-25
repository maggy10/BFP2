import streamlit as st
import pandas as pd
import numpy as np
import io

# 1. Configuración de la página de Streamlit
st.set_page_config(
    page_title="Balance Financiero Proyectado",
    layout="centered"
)

# Renderizar logo y título con el nuevo formato 2026
try:
    st.image("logo3.png", width="stretch")
except:
    pass

st.title("Balance Financiero Proyectado")

# 2. Función optimizada para cargar archivos pesados con caché
@st.cache_data
def cargar_datos(archivo_subido):
    return pd.read_excel(archivo_subido, sheet_name=0, engine='openpyxl')

# 3. Interfaz de usuario
archivo = st.file_uploader("Sube tu archivo de Excel", type=["xlsx"])
boton = st.button('Genera el BFP')

# 4. Ejecución de la lógica
if archivo is not None:
    try:
        df_original = cargar_datos(archivo)
        st.success("¡Archivo cargado con éxito en memoria!")
        st.write("Vista previa de los datos originales:")
        st.dataframe(df_original.head(5), width="stretch")
    except Exception as e:
        st.error(f"Hubo un error al procesar el archivo: {e}")
        st.stop()

    # Si el usuario presiona el botón, procesamos toda la información financiera
    if boton:
        with st.spinner("Procesando datos y generando balances... Por favor espera."):
            try:
                df = df_original.copy()

                # Convertimos la columna a fechas reales
                df['FECHA'] = pd.to_datetime(df['FECHA_CONEX'], format='%d/%m/%Y', errors='coerce')
                df['ID'] = df['SUCURSAL_CVE'].astype(str) + df['UNIDAD_OP_CVE'].astype(str) + df['FOLIO_DEF'].astype(str)
                
                # Filtros iniciales de limpieza
                df_limpio1 = df.dropna(subset=['FECHA'])
                df_limpio2 = df_limpio1[df_limpio1['ESTATUS'] == 'Conectado']
                df_limpio3 = df_limpio2[df_limpio2['TIPO_CEGAP'] != 'Cegap Nac. de Proveedores Descontados']

                base = df_limpio3[~((df_limpio3['FOLIO_DEF'].between(49999,79999)) &
                                    ((df_limpio3['TIPO_CEGAP'] == 'Cegap de Erogacion') |
                                     (df_limpio3['TIPO_CEGAP'] == 'Cegap de Erogacion de Proveedores (Mercancias)')) &
                                    (df_limpio3['TIPO_PAGO'] == 'Cegap de registro'))]

                # --- SECCIÓN: EGRESOS ---
                gasto = base[(base['CAPITULO'] != 7000000)]
                gasto = gasto[(gasto['PARTIDA'] != 43101001)]
                gasto = gasto[(gasto['PARTIDA'] != 43701001)]
                gasto = gasto[(gasto['PARTIDA'] != 39908031)]
                gasto = gasto[(gasto['PARTIDA'] != 39908046)]
                gasto = gasto[~(((gasto['PARTIDAOF'] == 39908) | (gasto['PARTIDAOF'] == 39909)) &
                                (gasto['FUENTE_FIN'] != 'Rec. Propios') &
                                (gasto['UNIDAD_OP_NOM'] != 'OFICINAS CENTRALES  '))]

                # Gasto transversal
                gasto_trans = gasto[(((gasto['PROG_PRES'] == 'M001') & (gasto['PARTIDAOF'] != 33903)) |
                                      (gasto['PROG_PRES'] == 'O001') |
                                      (gasto['PROG_PRES'] == 'P021')) &
                                    ~(((gasto['PARTIDAOF'] == 39908) | (gasto['PARTIDAOF'] == 39909)))]
                gasto_transversal = ((gasto_trans.groupby('CAPITULO')['IMPORTE'].sum()) / 1000000).reset_index()
                gasto_transversal = gasto_transversal.rename(columns={'CAPITULO': 'Capítulo', 'IMPORTE': 'Gasto Transversal'})

                # Gasto Acopio
                gasto_acopiop = gasto[(gasto['PROG_PRES'] == 'S290') & 
                                      (gasto['FUENTE_FIN'] == 'Rec. Propios') &
                                      ~(((gasto['PARTIDAOF'] == 39908) | (gasto['PARTIDAOF'] == 39909)))]
                gasto_acopiof = gasto[(gasto['PROG_PRES'] == 'S290') & (gasto['FUENTE_FIN'] != 'Rec. Propios')]

                gap = ((gasto_acopiop.groupby('CAPITULO')['IMPORTE'].sum()) / 1000000).reset_index().set_index('CAPITULO')
                gaf = ((gasto_acopiof['IMPORTE'].sum()) / 1000000)
                gaf2 = pd.DataFrame({'CAPITULO': [40000000], 'IMPORTE': [gaf]}).set_index('CAPITULO')
                ga = pd.concat([gap, gaf2], axis=0)
                gasto_acopio = ga.reset_index().rename(columns={'CAPITULO': 'Capítulo', 'IMPORTE': 'Acopio'})

                # Gasto Par
                gasto_par_base = gasto[((gasto['PROG_PRES'] == 'S053') & 
                                        (gasto['FUENTE_FIN'] == 'Rec. Propios') &
                                        ~gasto['AREA_AFECT'].isin([952, 953, 954, 955, 957])) &
                                       ~gasto['PARTIDAOF'].isin([39908, 39909])]
                gpp = gasto_par_base.groupby('CAPITULO')['IMPORTE'].sum()

                gasto_p = gasto[(gasto['PROG_PRES'] == 'M001') & (gasto['PARTIDAOF'] == 33903)]
                gpp2 = gasto_p.groupby('CAPITULO')['IMPORTE'].sum()
                gpp3 = (pd.concat([gpp, gpp2], axis=0)).reset_index()
                gpp4 = (gpp3.groupby('CAPITULO')['IMPORTE'].sum()).reset_index()

                gpf = gasto[((gasto['PROG_PRES'] == 'S053') & 
                             (gasto['FUENTE_FIN'] == 'Rec. Fiscales') &
                             (~gasto['CAPITULO'].isin([40000000, 10000000]))) &
                            ~gasto['AREA_AFECT'].isin([952, 953, 954, 955, 957])]

                gpf2 = gasto[gasto['ID'].isin(gpf['ID'])]['IMPORTE'].sum()
                gpf3 = pd.DataFrame({'CAPITULO': [40000000], 'IMPORTE': [gpf2]})
                gpf3.index.name = 'CAPITULO'

                gpar = pd.concat([gpp4, gpf3], axis=0)
                gasto_par = ((gpar.groupby('CAPITULO')['IMPORTE'].sum()) / 1000000).reset_index()
                gasto_par = gasto_par.rename(columns={'CAPITULO': 'Capítulo', 'IMPORTE': 'PAR'})

                # Gasto Transformación
                gasto_tp1 = gasto[((gasto['PROG_PRES'] == 'S053') & 
                                    (gasto['FUENTE_FIN'] == 'Rec. Propios') &
                                    gasto['AREA_AFECT'].isin([952, 953, 954, 955, 957])) &
                                   ~gasto['PARTIDAOF'].isin([39908, 39909])]

                gasto_tp2 = gasto[(gasto['PROG_PRES'] == 'W001') & 
                                   gasto['PARTIDA'].isin([39909003, 39909005]) &
                                   gasto['AREA_AFECT'].isin([952, 953, 954, 955, 957])]

                gasto_tp = pd.concat([gasto_tp1, gasto_tp2], axis=0)
                g_tp = ((gasto_tp.groupby('CAPITULO')['IMPORTE'].sum()) / 1000000).reset_index()

                gasto_tf = gasto[((gasto['PROG_PRES'] == 'S053') & (gasto['FUENTE_FIN'] == 'Rec. Fiscales') & (gasto['CAPITULO'] == 40000000)) |
                                 (((gasto['PROG_PRES'] == 'S053') & (gasto['FUENTE_FIN'] == 'Rec. Fiscales') & (gasto['CAPITULO'] != 40000000)) &
                                  gasto['AREA_AFECT'].isin([952, 953, 954, 955, 957]))]

                g_tf = (gasto_tf['IMPORTE'].sum()) / 1000000
                g_tf2 = pd.DataFrame({'CAPITULO': [40000000], 'IMPORTE': [g_tf]})
                g_tf2.index.name = 'CAPITULO'

                g_tpf = pd.concat([g_tp, g_tf2], axis=0)
                gasto_transformacion = ((g_tpf.groupby('CAPITULO')['IMPORTE'].sum())).reset_index()
                gasto_transformacion = gasto_transformacion.rename(columns={'CAPITULO': 'Capítulo', 'IMPORTE': 'Transformación'})

                # Gasto Maíz es la Raíz
                gmr1 = gasto[(gasto['PROG_PRES'] == 'S053') & (gasto['FUENTE_FIN'] == 'Rec. Fiscales') & (gasto['CAPITULO'] == 10000000)]
                gastom = ((gmr1.groupby('CAPITULO')['IMPORTE'].sum()) / 1000000).reset_index()

                gmr = gasto[(gasto['PROG_PRES'] == 'S053') & (gasto['AREA_AFECT'] == 990) & (gasto['FUENTE_FIN'] == 'Rec. Fiscales')]
                gastom2 = ((gmr.groupby('CAPITULO')['IMPORTE'].sum()) / 1000000).reset_index()

                g_tpf_maiz = pd.concat([gastom, gastom2], axis=0)
                gasto_maiz = g_tpf_maiz.rename(columns={'CAPITULO': 'Capítulo', 'IMPORTE': 'Maíz es la Raíz'})

                # --- CÁLCULO DE FUENTES ANTES DE FORMATO ---
                p = gpp4['IMPORTE'].sum() / 1000000 if not gpp4.empty else 0
                tr = g_tp['IMPORTE'].sum() if not g_tp.empty else 0
                ta = gasto_transversal['Gasto Transversal'].sum() if not gasto_transversal.empty else 0
                a = gap['IMPORTE'].sum() if not gap.empty else 0

                pf_val = gpf3['IMPORTE'].sum() / 1000000
                trf = g_tf2['IMPORTE'].sum()
                af = gaf2['IMPORTE'].sum()
                m = gasto_maiz['Maíz es la Raíz'].sum() if not gasto_maiz.empty else 0

                # Gasto integrado (Egresos)
                gasto_cap = pd.DataFrame({
                    'Capítulo': [10000000, 20000000, 30000000, 40000000, 50000000],
                    'Concepto': ['Servicios Personales', 'Materiales y Suministros', 'Servicios Generales', 'Subsidios y Transferencias', 'Inversión']
                }).set_index('Capítulo')

                e1 = pd.merge(gasto_cap, gasto_acopio, on='Capítulo', how='outer')
                e2 = pd.merge(e1, gasto_par, on='Capítulo', how='outer')
                e3 = pd.merge(e2, gasto_transformacion, on='Capítulo', how='outer')
                e4 = pd.merge(e3, gasto_maiz, on='Capítulo', how='outer')
                e5 = pd.merge(e4, gasto_transversal, on='Capítulo', how='outer')
                egresos = e5.fillna(0)

                # --- SECCIÓN: INGRESOS ---
                pf = base[(base['PARTIDAOF'] == 72310) & (base['FUENTE_FIN'] == 'Rec. Propios')]
                pf2 = ((pf.groupby('CAPITULO')['IMPORTE'].sum()) / 1000000).abs().reset_index()
                pf_val_ing = pf2['IMPORTE'].sum() if not pf2.empty else 0
                productos_financieros = pd.DataFrame({'Concepto': ['Productos Financieros'], 'Gasto Transversal': [pf_val_ing], 'PAR': [0], 'Acopio': [0], 'Transformación': [0], 'Maíz es la Raíz': [0]})

                otros = base[(base['PARTIDAOF'] == 72320) & (base['FUENTE_FIN'] == 'Rec. Propios')]
                otros2 = ((otros.groupby('CAPITULO')['IMPORTE'].sum()) / 1000000).abs().reset_index()
                otros_val = otros2['IMPORTE'].sum() if not otros2.empty else 0
                otros_ingresos = pd.DataFrame({'Concepto': ['Otros'], 'PAR': [otros_val], 'Acopio': [0], 'Transformación': [0], 'Maíz es la Raíz': [0], 'Gasto Transversal': [0]})

                ventas = base[(base['PARTIDAOF'] == 72210) & (base['FUENTE_FIN'] == 'Rec. Propios')]
                ventas2 = abs(((ventas['IMPORTE'].sum()) / 1000000))
                ventas_bienes = pd.DataFrame({'Concepto': ['Venta de Bienes'], 'Acopio': [494.37], 'PAR': [ventas2 - 494.37], 'Transformación': [0], 'Maíz es la Raíz': [0], 'Gasto Transversal': [0]})

                ingresos = pd.DataFrame({
                    'Concepto': ['Venta de Bienes', 'Productos Financieros', 'Otros', 'Subsidios y Transferencias'],
                    'Acopio': [ventas_bienes['Acopio'].sum(), productos_financieros['Acopio'].sum(), otros_ingresos['Acopio'].sum(), 0],
                    'PAR': [ventas_bienes['PAR'].sum(), productos_financieros['PAR'].sum(), otros_ingresos['PAR'].sum(), 0],
                    'Transformación': [ventas_bienes['Transformación'].sum(), productos_financieros['Transformación'].sum(), otros_ingresos['Transformación'].sum(), 0],
                    'Maíz es la Raíz': [ventas_bienes['Maíz es la Raíz'].sum(), productos_financieros['Maíz es la Raíz'].sum(), otros_ingresos['Maíz es la Raíz'].sum(), 0],
                    'Gasto Transversal': [ventas_bienes['Gasto Transversal'].sum(), productos_financieros['Gasto Transversal'].sum(), otros_ingresos['Gasto Transversal'].sum(), 0]
                })

                # --- CONSOLIDACIÓN FINAL ---
                columns_order = ['Acopio', 'PAR', 'Transformación', 'Maíz es la Raíz', 'Gasto Transversal']
                
                ig = pd.DataFrame({'Concepto': ['Ingresos'], 'Acopio': [''], 'PAR': [''], 'Transformación': [''], 'Maíz es la Raíz': [''], 'Gasto Transversal': ['']}).set_index('Concepto')
                eg = pd.DataFrame({'Concepto': ['Egresos'], 'Acopio': [''], 'PAR': [''], 'Transformación': [''], 'Maíz es la Raíz': [''], 'Gasto Transversal': ['']}).set_index('Concepto')

                ingresos_as_idx = ingresos.set_index('Concepto')
                egeresos2 = egresos.drop('Capítulo', axis=1).set_index('Concepto')

                for col in columns_order:
                    ingresos_as_idx[col] = ingresos_as_idx[col].map('{:,.3f}'.format)
                    egeresos2[col] = egeresos2[col].map('{:,.3f}'.format)

                resultado = pd.concat([ig, ingresos_as_idx, eg, egeresos2], axis=0)

                fuente = pd.DataFrame({
                    'Concepto': ['Fuente de Financiamiento', 'Propios', 'Fiscales'],
                    'Acopio': ['', '{:,.3f}'.format(a), '{:,.3f}'.format(af)],
                    'PAR': ['', '{:,.3f}'.format(p), '{:,.3f}'.format(pf_val)],
                    'Transformación': ['', '{:,.3f}'.format(tr), '{:,.3f}'.format(trf)],
                    'Maíz es la Raíz': ['', '0.000', '{:,.3f}'.format(m)],
                    'Gasto Transversal': ['', '{:,.3f}'.format(ta), '0.000']
                }).set_index('Concepto')

                tabla = pd.concat([resultado, fuente], axis=0).reset_index()
                tabla = tabla.rename(columns={'index': 'Concepto'})

                # Despliegue en Streamlit usando width="stretch" para cumplir reglas 2026
                st.subheader("📊 Balance Financiero Consolidado (Millones)")
                st.dataframe(tabla, width="stretch")

                # Exportación a Excel
                buffer = io.BytesIO()
                with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                    tabla.to_excel(writer, index=False, sheet_name='Balance')

                archivo_excel = buffer.getvalue()

                st.download_button(
                    label="Descargar Tabla en Excel ⬇️ ",
                    data=archivo_excel,
                    file_name='Balance_Financiero_Proyectado.xlsx',
                    mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                )
            except Exception as processing_error:
                st.error(f"Error en los cálculos financieros: {processing_error}")


#Para actualizar
#streamlit run Codigo.py
