from Utils.tools import Tools, CustomException
from Utils.querys import Querys
import pandas as pd
import base64
from io import BytesIO
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
import requests
import os
from dotenv import load_dotenv

load_dotenv()

class Graph:
    """
    Clase para gestionar el envío de correos mediante Microsoft Graph API.
    """

    def __init__(self, db):
        self.db = db
        self.tools = Tools()
        self.querys = Querys(self.db)
        
        # Credenciales de Microsoft Graph
        self.client_id = os.getenv('MICROSOFT_CLIENT_ID')
        self.client_secret = os.getenv('MICROSOFT_CLIENT_SECRET')
        self.tenant_id = os.getenv('MICROSOFT_TENANT_ID')
        self.graph_url = os.getenv('MICROSOFT_URL_GRAPH')
        self.auth_url = os.getenv('MICROSOFT_URL')

    # Función para obtener token de autenticación
    def obtener_token_graph(self):
        """
        Obtiene el token de acceso de Microsoft Graph API.
        
        Returns:
            str: Token de acceso
        """
        try:
            url = f"{self.auth_url}{self.tenant_id}/oauth2/v2.0/token"
            
            headers = {
                'Content-Type': 'application/x-www-form-urlencoded'
            }
            
            data = {
                'client_id': self.client_id,
                'client_secret': self.client_secret,
                'scope': 'https://graph.microsoft.com/.default',
                'grant_type': 'client_credentials'
            }
            
            response = requests.post(url, headers=headers, data=data)
            
            if response.status_code == 200:
                return response.json().get('access_token')
            else:
                raise CustomException(f"Error al obtener token: {response.text}")
                
        except Exception as e:
            raise CustomException(f"Error en autenticación Graph: {str(e)}")

    # Función para generar HTML con tablas agrupadas
    def generar_html_tablas(self, datos_dian, datos_dms):
        """
        Genera HTML con tablas agrupadas por tipo de documento.
        
        Args:
            datos_dian: Datos procesados de DIAN
            datos_dms: Datos procesados de DMS
            
        Returns:
            tuple: (html, total_valor_dian, total_valor_dms, df_dian_completo, df_dms_completo)
        """
        try:
            total_valor_dian = 0
            total_valor_dms = 0
            df_dian_completo = None
            df_dms_completo = None
            
            html = """
            <html>
            <head>
                <style>
                    body {
                        font-family: Arial, sans-serif;
                        padding: 20px;
                    }
                    h2 {
                        color: #2c3e50;
                        border-bottom: 3px solid #3498db;
                        padding-bottom: 10px;
                        margin-top: 30px;
                    }
                    table {
                        width: 100%;
                        border-collapse: collapse;
                        margin: 20px 0;
                        box-shadow: 0 2px 8px rgba(0,0,0,0.1);
                    }
                    th {
                        background-color: #00B050;
                        color: white;
                        padding: 12px;
                        text-align: left;
                        font-weight: bold;
                    }
                    td {
                        padding: 10px;
                        border-bottom: 1px solid #ddd;
                    }
                    tr:nth-child(even) {
                        background-color: #f8f9fa;
                    }
                    tr:hover {
                        background-color: #e3f2fd;
                    }
                    .total-row {
                        background-color: #e8f5e9 !important;
                        font-weight: bold;
                    }
                    .number {
                        text-align: right;
                    }
                </style>
            </head>
            <body>
                <h1 style="color: #2c3e50;">Resumen de Facturación Electrónica</h1>
            """
            
            # Tabla DIAN
            if datos_dian and datos_dian.get('registros'):
                df_dian_completo = pd.DataFrame(datos_dian['registros'])
                
                # Agrupar por Tipo de documento, contar y sumar Saldo2
                agrupado_dian = df_dian_completo.groupby('Tipo de documento').agg(
                    Valor=('Saldo2', 'sum'),
                    Registros=('Saldo2', 'count')
                ).reset_index()
                agrupado_dian.columns = ['Tipo de documento', 'Valor', 'N° de registros']
                
                # Calcular totales
                total_valor_dian = agrupado_dian['Valor'].sum()
                total_registros_dian = agrupado_dian['N° de registros'].sum()
                
                html += """
                <h2>DIAN FACTURACION ELECTRONICA</h2>
                <table>
                    <thead>
                        <tr>
                            <th>Tipo de documento</th>
                            <th class="number">N° de registros</th>
                            <th class="number">Valor</th>
                        </tr>
                    </thead>
                    <tbody>
                """
                
                for _, row in agrupado_dian.iterrows():
                    html += f"""
                        <tr>
                            <td>{row['Tipo de documento']}</td>
                            <td class="number">{int(row['N° de registros'])}</td>
                            <td class="number">{row['Valor']:,.2f}</td>
                        </tr>
                    """
                
                html += f"""
                        <tr class="total-row">
                            <td><strong>Total general</strong></td>
                            <td class="number"><strong>{int(total_registros_dian)}</strong></td>
                            <td class="number"><strong>{total_valor_dian:,.2f}</strong></td>
                        </tr>
                    </tbody>
                </table>
                """
            
            # Tabla DMS
            if datos_dms and datos_dms.get('registros'):
                df_dms_completo = pd.DataFrame(datos_dms['registros'])
                
                # Mapear códigos a nombres descriptivos
                mapeo_tipos = {
                    'FC': 'Factura electrónica',
                    'DV': 'Nota de crédito electrónica'
                }
                
                # Aplicar mapeo
                df_dms_completo['Tipo de documento'] = df_dms_completo['Tipo Docto.'].map(mapeo_tipos).fillna(df_dms_completo['Tipo Docto.'])
                
                # Agrupar por Tipo de documento, contar y sumar Saldo2
                agrupado_dms = df_dms_completo.groupby('Tipo de documento').agg(
                    Valor=('Saldo2', 'sum'),
                    Registros=('Saldo2', 'count')
                ).reset_index()
                agrupado_dms.columns = ['Tipo de documento', 'Valor', 'N° de registros']
                
                # Calcular totales
                total_valor_dms = agrupado_dms['Valor'].sum()
                total_registros_dms = agrupado_dms['N° de registros'].sum()
                
                html += """
                <h2>FACTURACION ELECTRONICA DMS CONTABLE</h2>
                <table>
                    <thead>
                        <tr>
                            <th>Tipo de documento</th>
                            <th class="number">N° de registros</th>
                            <th class="number">Valor</th>
                        </tr>
                    </thead>
                    <tbody>
                """
                
                for _, row in agrupado_dms.iterrows():
                    html += f"""
                        <tr>
                            <td>{row['Tipo de documento']}</td>
                            <td class="number">{int(row['N° de registros'])}</td>
                            <td class="number">{row['Valor']:,.2f}</td>
                        </tr>
                    """
                
                html += f"""
                        <tr class="total-row">
                            <td><strong>Total general</strong></td>
                            <td class="number"><strong>{int(total_registros_dms)}</strong></td>
                            <td class="number"><strong>{total_valor_dms:,.2f}</strong></td>
                        </tr>
                    </tbody>
                </table>
                """
            
            html += """
            </body>
            </html>
            """
            
            return html, total_valor_dian, total_valor_dms, df_dian_completo, df_dms_completo
            
        except Exception as e:
            raise CustomException(f"Error al generar HTML: {str(e)}")

    # Función para generar Excel de DIAN
    def generar_excel_adjunto(self, df, nombre_origen):
        """
        Genera un archivo Excel con formato para adjuntar al correo.
        
        Args:
            df: DataFrame con los datos
            nombre_origen: "DIAN" o "DMS"
            
        Returns:
            str: Archivo Excel en base64
        """
        try:
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df.to_excel(writer, index=False, sheet_name=f'Datos {nombre_origen}')
                
                # Obtener el workbook y la hoja activa
                workbook = writer.book
                worksheet = writer.sheets[f'Datos {nombre_origen}']
                
                # Definir estilos para el encabezado
                header_fill = PatternFill(start_color='00B050', end_color='00B050', fill_type='solid')
                header_font = Font(name='Calibri', size=11, bold=True, color='FFFFFF')
                header_alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
                border_side = Side(style='thin', color='000000')
                header_border = Border(left=border_side, right=border_side, top=border_side, bottom=border_side)
                
                # Aplicar estilos a la fila de encabezado
                for cell in worksheet[1]:
                    cell.fill = header_fill
                    cell.font = header_font
                    cell.alignment = header_alignment
                    cell.border = header_border
                
                # Ajustar el ancho de las columnas automáticamente
                for column in worksheet.columns:
                    max_length = 0
                    column_letter = column[0].column_letter
                    for cell in column:
                        try:
                            if len(str(cell.value)) > max_length:
                                max_length = len(str(cell.value))
                        except:
                            pass
                    adjusted_width = min(max_length + 2, 50)
                    worksheet.column_dimensions[column_letter].width = adjusted_width
            
            output.seek(0)
            return base64.b64encode(output.read()).decode('utf-8')
            
        except Exception as e:
            raise CustomException(f"Error al generar Excel de {nombre_origen}: {str(e)}")

    # Función para enviar correo
    def enviar_correo_reporte(self, data: dict):
        """
        Envía correo con el reporte de facturación electrónica.
        
        Args:
            data (dict): Datos opcionales (por ahora vacío)
        """
        try:
            # Obtener token de autenticación
            token = self.obtener_token_graph()
            
            # Obtener últimos datos procesados
            datos = self.querys.obtener_ultimos_datos_procesados()
            
            if not datos.get('dian') and not datos.get('dms'):
                raise CustomException("No hay datos procesados disponibles para enviar")
            
            # Generar HTML del correo y obtener totales
            datos_dian = datos.get('dian', {}).get('datos') if datos.get('dian') else None
            datos_dms = datos.get('dms', {}).get('datos') if datos.get('dms') else None
            
            html_body, total_dian, total_dms, df_dian, df_dms = self.generar_html_tablas(datos_dian, datos_dms)
            
            # Verificar si los totales son diferentes
            totales_diferentes = False
            adjuntos = []
            
            if total_dian != 0 and total_dms != 0:
                # Comparar con una tolerancia para evitar problemas de precisión de punto flotante
                if abs(total_dian - total_dms) > 0.01:
                    totales_diferentes = True
                    
                    # Generar archivos Excel
                    if df_dian is not None:
                        excel_dian_base64 = self.generar_excel_adjunto(df_dian, "DIAN")
                        adjuntos.append({
                            "nombre": "Datos_DIAN.xlsx",
                            "contenido": excel_dian_base64
                        })
                    
                    if df_dms is not None:
                        excel_dms_base64 = self.generar_excel_adjunto(df_dms, "DMS")
                        adjuntos.append({
                            "nombre": "Datos_DMS.xlsx",
                            "contenido": excel_dms_base64
                        })
            
            # Preparar el asunto del correo
            asunto = "Resumen de Facturación Electrónica - DIAN y DMS"
            if totales_diferentes:
                asunto += " ⚠️ DIFERENCIA DETECTADA"
            
            # Preparar el correo
            email_data = {
                "message": {
                    "subject": asunto,
                    "body": {
                        "contentType": "HTML",
                        "content": html_body
                    },
                    "toRecipients": [
                        {
                            "emailAddress": {
                                "address": "sistemas@avantika.com.co"
                            }
                        }
                    ],
                    "ccRecipients": [
                        {
                            "emailAddress": {
                                "address": "auxiliartic@avantika.com.co"
                            }
                        }
                    ]
                },
                "saveToSentItems": "true"
            }
            
            # Agregar adjuntos si existen
            if adjuntos:
                email_data["message"]["attachments"] = []
                for adjunto in adjuntos:
                    email_data["message"]["attachments"].append({
                        "@odata.type": "#microsoft.graph.fileAttachment",
                        "name": adjunto["nombre"],
                        "contentType": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        "contentBytes": adjunto["contenido"]
                    })
            
            # Enviar correo usando Graph API
            url = f"{self.graph_url}sistemas@avantika.com.co/sendMail"
            
            headers = {
                'Authorization': f'Bearer {token}',
                'Content-Type': 'application/json'
            }
            
            response = requests.post(url, json=email_data, headers=headers)
            
            if response.status_code == 202:
                mensaje_respuesta = "Correo enviado exitosamente"
                if totales_diferentes:
                    mensaje_respuesta += " con archivos Excel adjuntos (diferencia detectada)"
                
                return self.tools.output(200, mensaje_respuesta, {
                    "destinatarios": ["sistemas@avantika.com.co"],
                    "copia": ["auxiliartic@avantika.com.co"],
                    "tiene_datos_dian": datos.get('dian') is not None,
                    "tiene_datos_dms": datos.get('dms') is not None,
                    "totales_diferentes": totales_diferentes,
                    "total_dian": float(total_dian),
                    "total_dms": float(total_dms),
                    "archivos_adjuntos": len(adjuntos)
                })
            else:
                raise CustomException(f"Error al enviar correo: {response.status_code} - {response.text}")
            
        except CustomException as e:
            raise e
        except Exception as e:
            print(f"Error al enviar correo: {e}")
            raise CustomException(f"Error al enviar correo: {str(e)}")
