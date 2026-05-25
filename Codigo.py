import streamlit as st
import pandas as pd
import numpy as np
import io

# Configuración de la página de Streamlit
import os
try:
    import openpyxl
except ImportError:
    os.system('pip install openpyxl')

st.set_page_config(
    page_title="Balance Financiero Proyectado",
    layout="centered")


st.image("logo.png", use_container_width=True)


st.title("Balance Financiero Proyectado")

# 1. Cargar archivo
@st.cache_data
def cargar_datos(archivo_subido):
    # Usamos 'openpyxl' explícitamente como motor de lectura
    return pd.read_excel(archivo_subido, sheet_name=0, engine='openpyxl')


archivo = st.file_uploader("Sube tu archivo de Excel", type=["xlsx"])
boton = st.button('Genera el BFP')

# 1. Cargar archivo
if archivo is not None:
      st.info("Cargando archivo pesado... Esto puede tomar un momento.")
      try:
          # Cargar los datos desde el archivo subido
          df = cargar_datos(archivo)
        
          # 'datos' ya es un DataFrame, no necesitas hacer df = pd.DataFrame(datos)
          df_m = datos.head(5)
        
          st.success("¡Archivo cargado con éxito!")
          st.write(df_m)
      except Exception as e:
          st.error(f"Hubo un error al procesar el archivo: {e}")

  #2 Correr codigo
      if boton is True:

          # Convertimos la columna a fechas reales
          df['FECHA'] = pd.to_datetime(df['FECHA_CONEX'], format='%d/%m/%Y', errors='coerce')
          df['ID'] = df['SUCURSAL_CVE'].astype(str) + df['UNIDAD_OP_CVE'].astype(str) + df['FOLIO_DEF'].astype(str)
          #Quita sin fecha de conexión
          df_limpio1 = df.dropna(subset=['FECHA'])
          df_limpio2 = df_limpio1[df_limpio1['ESTATUS'] == 'Conectado']
          df_limpio3 = df_limpio2[df_limpio2['TIPO_CEGAP'] != 'Cegap Nac. de Proveedores Descontados']

          duplicado = df_limpio3[(df_limpio3['FOLIO_DEF'].between(49999,79999))&
                       ((df_limpio3['TIPO_CEGAP'] == 'Cegap de Erogacion')|
                        (df_limpio3['TIPO_CEGAP'] == 'Cegap de Erogacion de Proveedores (Mercancias)'))&
                       (df_limpio3['TIPO_PAGO'] == 'Cegap de registro')]

          base = df_limpio3[~((df_limpio3['FOLIO_DEF'].between(49999,79999))&
                    ((df_limpio3['TIPO_CEGAP'] == 'Cegap de Erogacion')|
                    (df_limpio3['TIPO_CEGAP'] == 'Cegap de Erogacion de Proveedores (Mercancias)'))&
                    (df_limpio3['TIPO_PAGO'] == 'Cegap de registro'))]

          #Gasto

          gasto = base[(base['CAPITULO'] != 7000000)]
          gasto = gasto[(gasto['PARTIDA'] != 43101001)]
          gasto = gasto[(gasto['PARTIDA'] != 43701001)]
          gasto = gasto[(gasto['PARTIDA'] != 39908031)]
          gasto = gasto[(gasto['PARTIDA'] != 39908046)]
          gasto = gasto[~(((gasto['PARTIDAOF']== 39908)|
                (gasto['PARTIDAOF']== 39909))&
                (gasto['FUENTE_FIN']!= 'Rec. Propios') &
                (gasto['UNIDAD_OP_NOM']!= 'OFICINAS CENTRALES  '))]


          #Gasto transversal
          gasto_trans = gasto[(((gasto['PROG_PRES']=='M001')&
                    (gasto['PARTIDAOF']!= 33903))|
                    (gasto['PROG_PRES']== 'O001')|
                    (gasto['PROG_PRES']== 'P021'))&
                    ~(((gasto['PARTIDAOF']== 39908)|
                    (gasto['PARTIDAOF']== 39909)))]
          gasto_transversal = ((gasto_trans.groupby('CAPITULO')['IMPORTE'].sum())/1000000).reset_index()

          gasto_transversal = gasto_transversal.rename(columns={'CAPITULO': 'Capítulo','IMPORTE': 'Gasto Transversal'})

          #Gasto Acopio
          gasto_acopiop = gasto[(gasto['PROG_PRES']== 'S290')
                     & (gasto['FUENTE_FIN']== 'Rec. Propios')&
                    ~(((gasto['PARTIDAOF']== 39908)|
                    (gasto['PARTIDAOF']== 39909)))]

          gasto_acopiof = gasto[(gasto['PROG_PRES']== 'S290')&
                      (gasto['FUENTE_FIN']!= 'Rec. Propios')]


          gap = ((gasto_acopiop.groupby('CAPITULO')['IMPORTE'].sum())/1000000).reset_index().set_index('CAPITULO')


          gaf = ((gasto_acopiof['IMPORTE'].sum())/1000000)


          gaf2 = pd.DataFrame({'CAPITULO': [40000000],'IMPORTE': [gaf]}).set_index('CAPITULO')


          ga = pd.concat([gap, gaf2], axis=0)


          gasto_acopio = ga.reset_index()

          gasto_acopio = gasto_acopio.rename(columns={'CAPITULO': 'Capítulo','IMPORTE': 'Acopio'})


          #Gasto Par
          gasto_par = gasto[((gasto['PROG_PRES']== 'S053')
                     & (gasto['FUENTE_FIN']== 'Rec. Propios') &
                  ~ ((gasto['AREA_AFECT']== 952)|
                    (gasto['AREA_AFECT']== 953)|
                    (gasto['AREA_AFECT']== 954)|
                    (gasto['AREA_AFECT']== 955)|
                    (gasto['AREA_AFECT']== 957)))&
                    ~(((gasto['PARTIDAOF']== 39908)|
                    (gasto['PARTIDAOF']== 39909)))]

          gpp = gasto_par.groupby('CAPITULO')['IMPORTE'].sum()

          gasto_p = gasto[(gasto['PROG_PRES']== 'M001')&
                (gasto['PARTIDAOF']== 33903)]

          gpp2 = gasto_p.groupby('CAPITULO')['IMPORTE'].sum()

          gpp3 = (pd.concat([gpp, gpp2], axis=0)).reset_index()

          gpp4 = (gpp3.groupby('CAPITULO')['IMPORTE'].sum()).reset_index()


          gpf = gasto[((gasto['PROG_PRES'] == 'S053')
             & (gasto['FUENTE_FIN'] == 'Rec. Fiscales') &
              ((gasto['CAPITULO'] != 40000000)&
              (gasto['CAPITULO'] != 10000000)))&
               ~ ((gasto['AREA_AFECT']== 952)|
                    (gasto['AREA_AFECT']== 953)|
                    (gasto['AREA_AFECT']== 954)|
                    (gasto['AREA_AFECT']== 955)|
                    (gasto['AREA_AFECT']== 957))]

          gpf2 = gasto[gasto['ID'].isin(gpf['ID'])]

          gpf2 = gpf2['IMPORTE'].sum()

          gpf3 = pd.DataFrame({'CAPITULO': [40000000], 'IMPORTE': [gpf2]})
          gpf3.index.name = 'CAPITULO'

          gpar = (pd.concat([gpp4, gpf3], axis=0))

          gasto_par = ((gpar.groupby('CAPITULO')['IMPORTE'].sum())/1000000).reset_index()

          gasto_par = gasto_par.rename(columns={'CAPITULO': 'Capítulo','IMPORTE': 'PAR'})


          #Gasto Transformación
          gasto_tp1 = gasto[((gasto['PROG_PRES']== 'S053') &
                 (gasto['FUENTE_FIN']== 'Rec. Propios') &
                   ((gasto['AREA_AFECT']== 952)|
                    (gasto['AREA_AFECT']== 953)|
                    (gasto['AREA_AFECT']== 954)|
                    (gasto['AREA_AFECT']== 955)|
                    (gasto['AREA_AFECT']== 957)))&
                    ~((gasto['PARTIDAOF']== 39908)|
                   (gasto['PARTIDAOF']== 39909))]

          gasto_tp2 = gasto[(gasto['PROG_PRES']== 'W001') &
                 (((gasto['PARTIDA']== 39909003)|
                   (gasto['PARTIDA']== 39909005)))&
                   ((gasto['AREA_AFECT']== 952)|
                    (gasto['AREA_AFECT']== 953)|
                    (gasto['AREA_AFECT']== 954)|
                    (gasto['AREA_AFECT']== 955)|
                    (gasto['AREA_AFECT']== 957))]

          gasto_tp = (pd.concat([gasto_tp1, gasto_tp2], axis=0))

          g_tp = ((gasto_tp.groupby('CAPITULO')['IMPORTE'].sum())/1000000).reset_index()


          gasto_tf = gasto[((gasto['PROG_PRES'] == 'S053')&
                        (gasto['FUENTE_FIN'] == 'Rec. Fiscales') &
                        (gasto['CAPITULO'] == 40000000))|
                       (((gasto['PROG_PRES'] == 'S053') &
                        (gasto['FUENTE_FIN'] == 'Rec. Fiscales') &
                        (gasto['CAPITULO'] != 40000000)) &
                        ((gasto['AREA_AFECT']== 952)|
                         (gasto['AREA_AFECT']== 953)|
                         (gasto['AREA_AFECT']== 954)|
                         (gasto['AREA_AFECT']== 955)|
                         (gasto['AREA_AFECT']== 957)))]

          g_tf = (gasto_tf['IMPORTE'].sum())/1000000

          g_tf2 = pd.DataFrame({'CAPITULO': [40000000], 'IMPORTE': [g_tf]})
          g_tf2.index.name = 'CAPITULO'

          g_tpf = (pd.concat([g_tp, g_tf2], axis=0))

          gasto_transformacion = ((g_tpf.groupby('CAPITULO')['IMPORTE'].sum())).reset_index()

          gasto_transformacion = gasto_transformacion.rename(columns={'CAPITULO': 'Capítulo','IMPORTE': 'Transformación'})


          #Gasto Maíz es la Raíz
          gmr1 = gasto[(gasto['PROG_PRES'] == 'S053')
             & (gasto['FUENTE_FIN'] == 'Rec. Fiscales') &
              (gasto['CAPITULO'] == 10000000)]

          gastom = ((gmr1.groupby('CAPITULO')['IMPORTE'].sum())/1000000).reset_index()

          gmr = gasto[(gasto['PROG_PRES'] == 'S053') &
             (gasto['AREA_AFECT']== 990) &
            (gasto['FUENTE_FIN'] == 'Rec. Fiscales')]
          gastom2 = ((gmr.groupby('CAPITULO')['IMPORTE'].sum())/1000000).reset_index()

          g_tpf = (pd.concat([gastom, gastom2], axis=0))

          gasto_maiz = g_tpf.rename(columns={'CAPITULO': 'Capítulo','IMPORTE': 'Maíz es la Raíz'})


         #Gasto integrado
          gasto_cap = (pd.DataFrame({'Capítulo':[10000000, 20000000, 30000000, 40000000, 50000000],'Concepto': ['Servicios Personales', 'Materiales y Suministros', 'Servicios Generales', 'Subsidios y Transferencias', 'Inversión']})).set_index('Capítulo')

          e1 = pd.merge(gasto_cap, gasto_acopio, on='Capítulo', how='outer')
          e2 = pd.merge(e1, gasto_par, on='Capítulo', how='outer')
          e3 = pd.merge(e2, gasto_transformacion, on='Capítulo', how='outer')
          e4 = pd.merge(e3, gasto_maiz, on='Capítulo', how='outer')
          e5 = pd.merge(e4, gasto_transversal, on='Capítulo', how='outer')


          egresos = e5.fillna(0)

          egresos['Acopio'] = egresos['Acopio'].map('{:,.3f}'.format)
          egresos['PAR'] = egresos['PAR'].map('{:,.3f}'.format)
          egresos['Transformación'] = egresos['Transformación'].map('{:,.3f}'.format)
          egresos['Maíz es la Raíz'] = egresos['Maíz es la Raíz'].map('{:,.3f}'.format)
          egresos['Gasto Transversal'] = egresos['Gasto Transversal'].map('{:,.3f}'.format)

          egeresos2 = egresos.drop('Capítulo', axis=1)


          #Ingresos

          pf = base[(base['PARTIDAOF'] == 72310) &
          (base['FUENTE_FIN'] == 'Rec. Propios')]

          #Productos Financieros
          pf2 = ((pf.groupby('CAPITULO')['IMPORTE'].sum())/1000000).abs().reset_index()

          productos_financieros = pd.DataFrame({'Concepto': 'Productos Financieros', 'Gasto Transversal': pf2['IMPORTE'], 'PAR':[0], 'Acopio': [0],'Transformación':[0], 'Maíz es la Raíz': [0]})


          #Otros ingresos
          otros = base[(base['PARTIDAOF'] == 72320) &
          (base['FUENTE_FIN'] == 'Rec. Propios')]

          otros2 = ((otros.groupby('CAPITULO')['IMPORTE'].sum())/1000000).abs().reset_index()

          otros_ingresos = pd.DataFrame({'Concepto': 'Otros', 'PAR': otros2['IMPORTE'], 'Acopio': [0],'Transformación':[0], 'Maíz es la Raíz': [0], 'Gasto Transversal':[0]})



          #Ventas
          ventas = base[(base['PARTIDAOF']==72210) &
          (base['FUENTE_FIN'] == 'Rec. Propios')]

          ventas2 = abs(((ventas['IMPORTE'].sum())/1000000))


          ventas_bienes = pd.DataFrame({'Concepto': ['Venta de Bienes'], 'Acopio': [494.37], 'PAR': [ventas2 - 494.37],'Transformación':[0], 'Maíz es la Raíz': [0], 'Gasto Transversal':[0]})


          #Integración de ingresos
          ingresos = pd.DataFrame({'Concepto':['Venta de Bienes', 'Productos Financieros', 'Otros', 'Subsidios y Transferencias'],
                             'Acopio': [ventas_bienes['Acopio'].sum(), productos_financieros['Acopio'].sum(), otros_ingresos['Acopio'].sum(), 0],
                             'PAR': [ventas_bienes['PAR'].sum(), productos_financieros['PAR'].sum(), otros_ingresos['PAR'].sum(), 0],
                             'Transformación': [ventas_bienes['Transformación'].sum(), productos_financieros['Transformación'].sum(), otros_ingresos['Transformación'].sum(), 0],
                             'Maíz es la Raíz': [ventas_bienes['Maíz es la Raíz'].sum(), productos_financieros['Maíz es la Raíz'].sum(), otros_ingresos['Maíz es la Raíz'].sum(), 0],
                             'Gasto Transversal': [ventas_bienes['Gasto Transversal'].sum(), productos_financieros['Gasto Transversal'].sum(), otros_ingresos['Gasto Transversal'].sum(), 0]
                          })


          ingresos['Acopio'] = ingresos['Acopio'].map('{:,.3f}'.format)
          ingresos['PAR'] = ingresos['PAR'].map('{:,.3f}'.format)
          ingresos['Transformación'] = ingresos['Transformación'].map('{:,.3f}'.format)
          ingresos['Maíz es la Raíz'] = ingresos['Maíz es la Raíz'].map('{:,.3f}'.format)
          ingresos['Gasto Transversal'] = ingresos['Gasto Transversal'].map('{:,.3f}'.format)



          #Balance compelto

          ig = pd.DataFrame({'Concepto': ['Acopio', 'PAR', 'Transformación', 'Maíz es la Raíz', 'Gasto Transversal'], 'Ingresos': ['', '', '', '', '']})
          ig = ig.set_index('Concepto')
          ig = ig.T

          eg = pd.DataFrame({'Concepto': ['Acopio', 'PAR', 'Transformación', 'Maíz es la Raíz', 'Gasto Transversal'], 'Egresos': ['', '', '', '', '']})
          eg = eg.set_index('Concepto')
          eg = eg.T

          ingresos = ingresos.set_index('Concepto')
          egeresos2 = egeresos2.set_index('Concepto')

          resultado = pd.concat([ig, ingresos,eg,egeresos2], axis=0)

          #Fuentes de financiamiento
          p = gpp4['IMPORTE'].sum()/1000000
          tr= g_tp['IMPORTE'].sum()
          ta = gasto_transversal['Gasto Transversal'].sum()
          a = gap['IMPORTE'].sum()

          pf = gpf3['IMPORTE'].sum()/1000000
          trf =g_tf2['IMPORTE'].sum()
          af = gaf2['IMPORTE'].sum()
          m = gasto_maiz['Maíz es la Raíz'].sum()

          fuente = pd.DataFrame({'Concepto': ['Acopio', 'PAR', 'Transformación', 'Maíz es la Raíz', 'Gasto Transversal'], 'Fuente de Financiamiento': ['', '', '', '', ''], 'Propios': [a, p, tr, 0, ta], 'Fiscales': [af, pf, trf, m, 0] })

          fuente = fuente.set_index('Concepto')

          fuente['Propios'] = fuente['Propios'].map('{:,.3f}'.format)
          fuente['Fiscales'] = fuente['Fiscales'].map('{:,.3f}'.format)
          fuente = fuente.T


          #Tabla final

          tabla = pd.concat([resultado, fuente], axis=0)
          tabla = tabla.reset_index()
          tabla = tabla.rename(columns={'index': 'Concepto'})

          st.table(tabla, border='horizontal', hide_index=False, hide_header=False)

          buffer = io.BytesIO()

          # Creamos el archivo Excel en memoria usando openpyxl
          with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
            tabla.to_excel(writer, index=False, sheet_name='Balance')

          # Obtenemos los datos binarios
          archivo_excel = buffer.getvalue()

          # Mostramos el botón de descarga SOLO cuando la tabla ya ha sido generada
          st.download_button(
            label="Descargar Tabla en Excel ⬇️ ",
            data=archivo_excel,
            file_name='Balance Financiero Proyectado.xlsx',
            mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            )



#Para actualizar
#streamlit run Codigo.py
