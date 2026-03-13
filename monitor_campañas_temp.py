#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Sistema de Monitoreo de Campañas - Envío de Alertas por Email
Analiza un archivo Excel desde un link y envía alertas según el estado de piezas restantes
"""

import pandas as pd
import smtplib
import requests
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
import logging
from datetime import datetime
import os
from typing import Dict, List, Tuple
import time

# Configuración de logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('monitor_campañas.log'),
        logging.StreamHandler()
    ]
)

class MonitorCampañas:
    def __init__(self, excel_url: str, config_email: Dict):
        """
        Inicializa el monitor de campañas
        
        Args:
            excel_url: URL del archivo Excel
            config_email: Configuración del servidor de email
        """
        self.excel_url = excel_url
        self.config_email = config_email
        self.df_consumo = None
        self.campañas_correos = {}  # Mapeo campaña -> correo
        
    def descargar_excel(self) -> str:
        """
        Descarga el archivo Excel desde la URL
        
        Returns:
            str: Ruta del archivo descargado
        """
        try:
            logging.info(f"Descargando archivo Excel desde: {self.excel_url}")
            
            # Convertir URL de SharePoint si es necesario
            url_descarga = self.convertir_sharepoint_url(self.excel_url)
            
            # Realizar la petición HTTP
            headers = {
                'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'
            }
            
            response = requests.get(url_descarga, headers=headers, timeout=30, allow_redirects=True)
            response.raise_for_status()
            
            # Guardar el archivo temporalmente
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            filename = f"consumo_temp_{timestamp}.xlsx"
            
            with open(filename, 'wb') as f:
                f.write(response.content)
            
            logging.info(f"Archivo descargado exitosamente: {filename}")
            return filename
            
        except requests.RequestException as e:
            logging.error(f"Error al descargar el archivo: {e}")
            raise
        except Exception as e:
            logging.error(f"Error inesperado al descargar: {e}")
            raise
    
    def convertir_sharepoint_url(self, url: str) -> str:
        """
        Convierte una URL de SharePoint para descarga directa
        
        Args:
            url: URL original de SharePoint
            
        Returns:
            str: URL convertida para descarga
        """
        try:
            if 'sharepoint.com' in url and 'Doc.aspx' in url:
                # Extraer el sourcedoc ID
                import re
                match = re.search(r'sourcedoc=%7B([^%]+)%7D', url)
                if match:
                    doc_id = match.group(1)
                    # Convertir a URL de descarga directa
                    base_url = url.split('/_layouts')[0]
                    download_url = f"{base_url}/_layouts/15/download.aspx?SourceUrl={url}"
                    logging.info(f"URL convertida para SharePoint: {download_url}")
                    return download_url
            
            # Si no es SharePoint o no se puede convertir, devolver la URL original
            return url
            
        except Exception as e:
            logging.warning(f"Error al convertir URL de SharePoint, usando original: {e}")
            return url
    
    def leer_datos_consumo(self, archivo_excel: str) -> pd.DataFrame:
        """
        Lee los datos del nuevo archivo consolidado de horas soporte
        
        Args:
            archivo_excel: Ruta del archivo Excel
            
        Returns:
            pd.DataFrame: DataFrame con los datos de soporte y diseño
        """
        try:
            logging.info("Leyendo datos del archivo consolidado de horas soporte")
            
            # Leer el archivo Excel con header=0 para mantener los nombres originales
            df = pd.read_excel(archivo_excel, header=0)
            
            # Mapear las columnas basándose en la estructura conocida del Excel
            column_mapping = {
                'SEPTIEMBRE': 'CLIENTE',
                'Unnamed: 1': 'AREA_CAMPAÑA',
                'Horas Soporte Merpes': 'HORAS_MERPES_ASIGNADAS',
                'Unnamed: 3': 'HORAS_MERPES_CONSUMIDAS',
                'Unnamed: 4': 'HORAS_MERPES_DISPONIBLES',
                'Horas Soporte Zoftinium': 'HORAS_ZOFTINIUM_ASIGNADAS', 
                'Unnamed: 6': 'HORAS_ZOFTINIUM_CONSUMIDAS',
                'Unnamed: 7': 'HORAS_ZOFTINIUM_DISPONIBLES',
                'Piezas diseño': 'PIEZAS_DISEÑO_ASIGNADAS',
                'Unnamed: 9': 'PIEZAS_DISEÑO_CONSUMIDAS',
                'Unnamed: 10': 'PIEZAS_DISEÑO_DISPONIBLES',
                                'Correos': 'CORREOS'  # Esta columna ya tiene el nombre correcto
            }
            
            # Aplicar el mapeo de columnas solo para las que existen
            existing_mappings = {old_col: new_col for old_col, new_col in column_mapping.items() if old_col in df.columns}
            df = df.rename(columns=existing_mappings)
            logging.info(f"Columnas mapeadas: {existing_mappings}")
            
            # Filtrar la fila de encabezados (fila 0 dice "Cliente", "Asignadas", etc)
            df = df[df['CLIENTE'] != 'Cliente']
            df = df[df['CLIENTE'].notna()]
            
            # Limpiar y convertir datos
            df['CLIENTE'] = df['CLIENTE'].astype(str).str.strip()
            
            # Convertir columnas numéricas
            columnas_numericas = ['HORAS_MERPES_ASIGNADAS', 'HORAS_MERPES_CONSUMIDAS', 
                                'HORAS_ZOFTINIUM_ASIGNADAS', 'HORAS_ZOFTINIUM_CONSUMIDAS',
                                'PIEZAS_DISEÑO_ASIGNADAS', 'PIEZAS_DISEÑO_CONSUMIDAS']
            
            for col in columnas_numericas:
                if col in df.columns:
                    df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
                    logging.info(f"Columna '{col}' convertida a numérica")
                else:
                    df[col] = 0
                    logging.info(f"Columna '{col}' creada con valor por defecto: 0")
            
            # Calcular disponibles si no se mapearon correctamente
            if 'HORAS_MERPES_DISPONIBLES' not in df.columns:
                df['HORAS_MERPES_DISPONIBLES'] = df['HORAS_MERPES_ASIGNADAS'] - df['HORAS_MERPES_CONSUMIDAS']
            if 'HORAS_ZOFTINIUM_DISPONIBLES' not in df.columns:  
                df['HORAS_ZOFTINIUM_DISPONIBLES'] = df['HORAS_ZOFTINIUM_ASIGNADAS'] - df['HORAS_ZOFTINIUM_CONSUMIDAS']
            if 'PIEZAS_DISEÑO_DISPONIBLES' not in df.columns:
                df['PIEZAS_DISEÑO_DISPONIBLES'] = df['PIEZAS_DISEÑO_ASIGNADAS'] - df['PIEZAS_DISEÑO_CONSUMIDAS']
            
            # Filtrar filas válidas
            df = df[df['CLIENTE'] != '']
            
            logging.info(f"Datos leídos exitosamente. Total clientes: {len(df)}")
            return df
            
        except Exception as e:
            logging.error(f"Error al leer los datos de consumo: {e}")
            raise
    
    def cargar_mapeo_correos(self, archivo_mapeo: str = None):
        """
        Carga el mapeo de correos desde el mismo DataFrame (columna CORREOS)
        
        Args:
            archivo_mapeo: No se usa en la nueva versión
        """
        try:
            if self.df_consumo is not None:
                # Extraer correos del mismo DataFrame
                filas_con_correos = self.df_consumo[self.df_consumo['CORREOS'].notna()]
                
                for _, row in filas_con_correos.iterrows():
                    cliente = str(row['CLIENTE']).strip()
                    correos_str = str(row['CORREOS']).strip()
                    
                    # Skip if cliente is empty, NaN, or just numbers/invalid
                    if not cliente or cliente in ['0', 'nan', 'NaN', ''] or cliente.isdigit():
                        logging.warning(f"Saltando cliente inválido: '{cliente}'")
                        continue
                    
                    # Si hay múltiples correos separados por ; o ,
                    correos_lista = [email.strip() for email in correos_str.replace(';', ',').split(',') if email.strip()]
                    
                    if correos_lista:
                        self.campañas_correos[cliente] = correos_lista
                
                logging.info(f"Mapeo de correos cargado: {len(self.campañas_correos)} clientes con correos definidos")
                
                # Mostrar el mapeo encontrado
                for cliente, correos in self.campañas_correos.items():
                    logging.info(f"  • {cliente}: {', '.join(correos)}")
                    
            else:
                logging.warning("No hay datos cargados para extraer mapeo de correos")
                
        except Exception as e:
            logging.error(f"Error al cargar mapeo de correos: {e}")
    
    def crear_archivo_mapeo_ejemplo(self, archivo_mapeo: str):
        """
        Crea un archivo de ejemplo para el mapeo de correos
        """
        try:
            # Obtener campañas únicas del DataFrame actual
            if self.df_consumo is not None:
                campañas = self.df_consumo['CAMPAÑA'].unique()
            else:
                campañas = ['CAMPAÑA_EJEMPLO_1', 'CAMPAÑA_EJEMPLO_2']
            
            df_ejemplo = pd.DataFrame({
                'CAMPAÑA': campañas,
                'EMAIL': [f'responsable{i+1}@empresa.com' for i in range(len(campañas))]
            })
            
            df_ejemplo.to_excel(archivo_mapeo, index=False)
            logging.info(f"Archivo de mapeo de ejemplo creado: {archivo_mapeo}")
            
        except Exception as e:
            logging.error(f"Error al crear archivo de mapeo de ejemplo: {e}")
    
    def analizar_estado_campañas(self) -> List[Dict]:
        """
        Genera reportes mensuales para cada cliente con correos definidos
        
        Returns:
            List[Dict]: Lista de reportes a enviar
        """
        reportes = []
        
        # Solo procesar clientes que tienen correos definidos
        for cliente, correos_lista in self.campañas_correos.items():
            # Buscar datos del cliente
            datos_cliente = self.df_consumo[self.df_consumo['CLIENTE'] == cliente]
            
            if not datos_cliente.empty:
                # Sumar todos los valores para el cliente (puede tener múltiples filas)
                merpes_asignadas = datos_cliente['HORAS_MERPES_ASIGNADAS'].sum()
                merpes_consumidas = datos_cliente['HORAS_MERPES_CONSUMIDAS'].sum() 
                merpes_disponibles = datos_cliente['HORAS_MERPES_DISPONIBLES'].sum()
                
                zoftinium_asignadas = datos_cliente['HORAS_ZOFTINIUM_ASIGNADAS'].sum()
                zoftinium_consumidas = datos_cliente['HORAS_ZOFTINIUM_CONSUMIDAS'].sum()
                zoftinium_disponibles = datos_cliente['HORAS_ZOFTINIUM_DISPONIBLES'].sum()
                
                diseño_asignadas = datos_cliente['PIEZAS_DISEÑO_ASIGNADAS'].sum()
                diseño_consumidas = datos_cliente['PIEZAS_DISEÑO_CONSUMIDAS'].sum() 
                diseño_disponibles = datos_cliente['PIEZAS_DISEÑO_DISPONIBLES'].sum()
                
                # Calcular porcentajes reales
                porcentaje_merpes = (merpes_consumidas / merpes_asignadas * 100) if merpes_asignadas > 0 else 0
                porcentaje_zoftinium = (zoftinium_consumidas / zoftinium_asignadas * 100) if zoftinium_asignadas > 0 else 0
                porcentaje_diseño = (diseño_consumidas / diseño_asignadas * 100) if diseño_asignadas > 0 else 0
                
                # Crear reporte para este cliente
                area_campaña = datos_cliente['AREA_CAMPAÑA'].iloc[0] if 'AREA_CAMPAÑA' in datos_cliente.columns and not datos_cliente['AREA_CAMPAÑA'].isna().all() else 'Desarrollo Web'
                
                reporte = {
                    'cliente': cliente,
                    'area_campaña': area_campaña,
                    'correos_destino': correos_lista,
                    
                    # Datos de Merpes
                    'merpes_asignadas': merpes_asignadas,
                    'merpes_consumidas': merpes_consumidas, 
                    'merpes_disponibles': merpes_disponibles,
                    'porcentaje_merpes': porcentaje_merpes,
                    
                    # Datos de Zoftinium  
                    'zoftinium_asignadas': zoftinium_asignadas,
                    'zoftinium_consumidas': zoftinium_consumidas,
                    'zoftinium_disponibles': zoftinium_disponibles, 
                    'porcentaje_zoftinium': porcentaje_zoftinium,
                    
                    # Datos de Diseño
                    'diseño_asignadas': diseño_asignadas,
                    'diseño_consumidas': diseño_consumidas,
                    'diseño_disponibles': diseño_disponibles,
                    'porcentaje_diseño': porcentaje_diseño
                }
                reportes.append(reporte)
                
                logging.info(f"Reporte generado para {cliente}: Merpes {merpes_consumidas}/{merpes_asignadas}h ({porcentaje_merpes:.1f}%), Zoftinium {zoftinium_consumidas}/{zoftinium_asignadas}h ({porcentaje_zoftinium:.1f}%), Diseño {diseño_consumidas}/{diseño_asignadas} piezas ({porcentaje_diseño:.1f}%)")
        
        logging.info(f"Se generaron {len(reportes)} reportes mensuales")
        return reportes
    
    def calcular_porcentaje_consumo(self, valor_actual: float, valor_meta: float) -> float:
        """
        Calcula el porcentaje de consumo basado en una meta
        
        Args:
            valor_actual: Valor actual consumido
            valor_meta: Valor meta/límite
            
        Returns:
            float: Porcentaje de consumo
        """
        if valor_meta == 0:
            return 0
        return min((valor_actual / valor_meta) * 100, 100)  # Máximo 100%
    
    def determinar_tipo_alerta(self, piezas_restantes: int) -> str:
        """
        Determina el tipo de alerta según las piezas restantes
        
        Args:
            piezas_restantes: Número de piezas restantes
            
        Returns:
            str: Tipo de alerta ('critica', 'advertencia', None)
        """
        if piezas_restantes == 0:
            return 'critica'
        elif piezas_restantes <= 4:
            return 'advertencia'
        else:
            return None
    
    def generar_contenido_email(self, reporte: Dict) -> Tuple[str, str]:
        """
        Genera el contenido del email usando Jinja2 con personalización completa
        """
        from jinja2 import Template
        import base64
        import os
        
        cliente_nombre = reporte['cliente']
        area_campaña = reporte.get('area_campaña', 'Desarrollo Web')
        
        # Asunto del email
        asunto = f"📊 Reporte Mensual - {cliente_nombre}"
        
        # Función para convertir imágenes a base64
        def imagen_to_base64(ruta_imagen):
            try:
                if os.path.exists(ruta_imagen):
                    with open(ruta_imagen, "rb") as img_file:
                        return base64.b64encode(img_file.read()).decode()
                else:
                    logging.warning(f"Imagen no encontrada: {ruta_imagen}")
                    return ""
            except Exception as e:
                logging.error(f"Error al convertir imagen {ruta_imagen}: {e}")
                return ""
        
        # Preparar datos de imágenes
        imagenes = {
            'encabezado': imagen_to_base64("ENCABEZADO.png"),
            'footer': imagen_to_base64("FOOTER.png"),
            'titulo': imagen_to_base64("TITULO.png"),
            'icono1': imagen_to_base64("ICONO1.png"),
            'icono2': imagen_to_base64("ICONO2.png")
        }
        
        # Preparar datos del cliente para el template
        cliente_data = {
            'nombre': cliente_nombre,
            'area_campaña': area_campaña,
            'merpes': {
                'asignadas': reporte.get('merpes_asignadas', 0),
                'consumidas': reporte.get('merpes_consumidas', 0),
                'disponibles': reporte.get('merpes_disponibles', 0),
                'porcentaje': round(reporte.get('porcentaje_merpes', 0), 1)
            },
            'zoftinium': {
                'asignadas': reporte.get('zoftinium_asignadas', 0),
                'consumidas': reporte.get('zoftinium_consumidas', 0),
                'disponibles': reporte.get('zoftinium_disponibles', 0),
                'porcentaje': round(reporte.get('porcentaje_zoftinium', 0), 1),
                'estado': self._obtener_estado_alerta(reporte.get('zoftinium_disponibles', 0), 'horas')
            },
            'diseño': {
                'asignadas': reporte.get('diseño_asignadas', 0),
                'consumidas': reporte.get('diseño_consumidas', 0),
                'disponibles': reporte.get('diseño_disponibles', 0),
                'porcentaje': round(reporte.get('porcentaje_diseño', 0), 1),
                'estado': self._obtener_estado_alerta(reporte.get('diseño_disponibles', 0), 'piezas')
            },
            'recomendaciones': self._generar_recomendaciones(reporte)
        }
        
        # Cargar y renderizar template
        try:
            with open('email_template.html', 'r', encoding='utf-8') as f:
                template_content = f.read()
            
            template = Template(template_content)
            cuerpo_html = template.render(cliente=cliente_data, imagenes=imagenes)
            
            logging.info(f"✅ Email personalizado generado para {cliente_nombre}")
            
        except FileNotFoundError:
            logging.warning("❌ Template Jinja2 no encontrado. Usando fallback básico.")
            
            # Fallback simple si no encuentra el template
            cuerpo_html = f"""
            <html>
            <body>
                <h2>Reporte Mensual - {cliente_nombre}</h2>
                <p>Merpes: {cliente_data['merpes']['consumidas']}/{cliente_data['merpes']['asignadas']}h ({cliente_data['merpes']['porcentaje']}%)</p>
                <p>Zoftinium: {cliente_data['zoftinium']['consumidas']}/{cliente_data['zoftinium']['asignadas']}h ({cliente_data['zoftinium']['porcentaje']}%)</p>
                <p>Diseño: {cliente_data['diseño']['consumidas']}/{cliente_data['diseño']['asignadas']} piezas ({cliente_data['diseño']['porcentaje']}%)</p>
            </body>
            </html>
            """
        
        return asunto, cuerpo_html
    
    def _obtener_estado_alerta(self, disponibles: float, tipo: str) -> str:
        """Determina el estado de alerta según los recursos disponibles"""
        if tipo == 'horas':
            if disponibles <= 0:
                return 'critica'
            elif disponibles <= 2:
                return 'advertencia'
        elif tipo == 'piezas':
            if disponibles <= 0:
                return 'critica'
            elif disponibles <= 4:
                return 'advertencia'
        return ''
    
    def _generar_recomendaciones(self, reporte: Dict) -> list:
        """Genera recomendaciones personalizadas según el estado del cliente"""
        recomendaciones = []
        
        # Recomendaciones para Merpes
        porcentaje_merpes = reporte.get('porcentaje_merpes', 0)
        if porcentaje_merpes > 80:
            recomendaciones.append("Considera solicitar horas adicionales de soporte Merpes para el próximo mes.")
        
        # Recomendaciones para Zoftinium
        porcentaje_zoftinium = reporte.get('porcentaje_zoftinium', 0)
        disponibles_zoftinium = reporte.get('zoftinium_disponibles', 0)
        if porcentaje_zoftinium > 100:
            recomendaciones.append("⚠️ Has excedido las horas de Zoftinium asignadas. Revisa tu plan.")
        elif disponibles_zoftinium <= 2:
            recomendaciones.append("Quedan pocas horas de Zoftinium disponibles este mes.")
        
        # Recomendaciones para Diseño
        disponibles_diseño = reporte.get('diseño_disponibles', 0)
        if disponibles_diseño <= 4:
            recomendaciones.append("Quedan pocas piezas de diseño disponibles para este mes.")
        
        # Recomendación general
        recomendaciones.append("Verifica que los datos que registra Kupaa sean los correctos, así mitigas el consumo de soporte.")
        
        return recomendaciones
<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <style>
        body {{
            margin: 0;
            padding: 0;
            font-family: Arial, sans-serif;
            background-color: #f0f0f0;
        }}
        .email-container {{
            max-width: 600px;
            margin: 0 auto;
            background-color: white;
            box-shadow: 0 0 20px rgba(0,0,0,0.1);
        }}
        .header-image {{
            width: 100%;
            display: block;
            border-radius: 0;
        }}
        .content {{
            padding: 20px;
        }}
        .titulo-section {{
            text-align: center;
            background: linear-gradient(135deg, #2c3e50, #34495e);
            color: white;
            padding: 20px;
            border-radius: 15px;
            margin: 20px 0;
        }}
        .titulo-imagen {{
            max-width: 400px;
            width: 100%;
            height: auto;
            margin-bottom: 15px;
        }}
        .cliente-campaña {{
            background-color: #1abc9c;
            color: white;
            padding: 12px 25px;
            border-radius: 25px;
            font-size: 16px;
            font-weight: bold;
            display: inline-block;
            margin: 10px 0;
        }}
        .consumo-title {{
            text-align: center;
            font-size: 16px;
            color: #555;
            margin: 30px 0;
            font-weight: bold;
        }}
        .porcentajes-row {{
            display: flex;
            justify-content: space-around;
            margin: 30px 0;
            flex-wrap: wrap;
        }}
        .porcentaje-card {{
            text-align: center;
            margin: 10px;
            flex: 1;
            min-width: 120px;
        }}
        .porcentaje-label {{
            font-size: 11px;
            color: #666;
            text-transform: uppercase;
            margin-bottom: 8px;
            line-height: 1.2;
        }}
        .porcentaje-circle {{
            width: 80px;
            height: 80px;
            border-radius: 50%;
            display: flex;
            align-items: center;
            justify-content: center;
            margin: 0 auto 10px;
            font-size: 18px;
            font-weight: bold;
            color: white;
            border: 3px solid rgba(255,255,255,0.3);
        }}
        .diseño {{ background-color: #FF6B35; }}
        .zoftinium {{ background-color: #4A90E2; }}
        .merpes {{ background-color: #7ED321; }}
        .detalles-table {{
            background: linear-gradient(135deg, #2c3e50, #34495e);
            border-radius: 15px;
            padding: 20px;
            margin: 30px 0;
            color: white;
        }}
        .table-header {{
            display: flex;
            justify-content: center;
            align-items: center;
            margin-bottom: 20px;
        }}
        .table-icon {{
            width: 50px;
            height: 50px;
            margin-right: 0px;
        }}
        .table-content {{
            display: grid;
            grid-template-columns: 1fr 1fr 1fr;
            gap: 15px;
            text-align: center;
        }}
        .table-column {{
            background-color: rgba(255,255,255,0.1);
            border-radius: 10px;
            padding: 15px 8px;
        }}
        .column-header {{
            font-size: 11px;
            font-weight: bold;
            margin-bottom: 15px;
            text-transform: uppercase;
            color: #1abc9c;
        }}
        .column-item {{
            padding: 6px 0;
            font-size: 12px;
            border-bottom: 1px solid rgba(255,255,255,0.1);
        }}
        .column-item:last-child {{
            border-bottom: none;
        }}
        .recomendaciones {{
            margin: 30px 0;
        }}
        .recomendaciones-title {{
            display: flex;
            align-items: center;
            font-size: 16px;
            font-weight: bold;
            color: #2c3e50;
            margin-bottom: 15px;
        }}
        .recomendaciones-icon {{
            width: 30px;
            height: 30px;
            margin-right: 10px;
        }}
        .recomendaciones-list {{
            list-style: none;
            padding: 0;
        }}
        .recomendaciones-list li {{
            padding: 10px 0;
            color: #555;
            font-size: 13px;
            border-left: 3px solid #1abc9c;
            padding-left: 15px;
            margin-bottom: 8px;
            background-color: #f8f9fa;
            border-radius: 0 5px 5px 0;
            line-height: 1.4;
        }}
        .footer-section {{
            background-color: #2c3e50;
            padding: 20px;
            text-align: center;
        }}
        .contactanos {{
            color: white;
            font-size: 18px;
            font-weight: bold;
            margin-bottom: 15px;
        }}
        .iconos-contacto {{
            display: flex;
            justify-content: center;
            gap: 15px;
            flex-wrap: wrap;
            margin-bottom: 20px;
        }}
        .icono-contacto {{
            width: 50px;
            height: 50px;
            border-radius: 50%;
            background-color: #34495e;
            padding: 10px;
        }}
        .footer-image {{
            width: 100%;
            display: block;
            border-radius: 0;
        }}
        @media (max-width: 600px) {{
            .porcentajes-row {{
                flex-direction: column;
                align-items: center;
            }}
            .table-content {{
                grid-template-columns: 1fr;
                gap: 10px;
            }}
        }}
    </style>
</head>
<body>
    <div class="email-container">
        <!-- Encabezado con imagen real -->
        <img src="data:image/png;base64,{encabezado_b64}" alt="Encabezado Merpes - Kupaa" class="header-image">
        
        <div class="content">
            <!-- Título con imagen real -->
            <div class="titulo-section">
                <img src="data:image/png;base64,{titulo_b64}" alt="Conoce el estado de tus requerimientos" class="titulo-imagen">
                <div class="cliente-campaña">{cliente} - {area_campaña}</div>
            </div>
            
            <!-- Consumo del mes -->
            <div class="consumo-title">CONSUMO DEL MES EN SOPORTE Y DISEÑO</div>
            
            <!-- Porcentajes -->
            <div class="porcentajes-row">
                <div class="porcentaje-card">
                    <div class="porcentaje-label">Piezas<br>Diseño</div>
                    <div class="porcentaje-circle diseño">{reporte['porcentaje_diseño']:.0f}%</div>
                </div>
                <div class="porcentaje-card">
                    <div class="porcentaje-label">Soporte<br>Zoftinium</div>
                    <div class="porcentaje-circle zoftinium">{reporte['porcentaje_zoftinium']:.0f}%</div>
                </div>
                <div class="porcentaje-card">
                    <div class="porcentaje-label">Soporte Grupo<br>Merpes</div>
                    <div class="porcentaje-circle merpes">{reporte['porcentaje_merpes']:.0f}%</div>
                </div>
            </div>
            
            <!-- Tabla de detalles con icono real -->
            <div class="detalles-table">
                <div class="table-header">
                    <img src="data:image/png;base64,{icono1_b64}" alt="Icono estadísticas" class="table-icon">
                </div>
                
                <div class="table-content">
                    <div class="table-column">
                        <div class="column-header">Piezas Diseño</div>
                        <div class="column-item">Piezas meta: {reporte['diseño_asignadas']:.0f}</div>
                        <div class="column-item">Piezas consumidas: {reporte['diseño_consumidas']:.0f}</div>
                        <div class="column-item">Piezas disponibles: {reporte['diseño_disponibles']:.0f}</div>
                    </div>
                    
                    <div class="table-column">
                        <div class="column-header">Soporte Zoftinium</div>
                        <div class="column-item">Horas meta: {reporte['zoftinium_asignadas']:.0f}</div>
                        <div class="column-item">Horas consumidas: {reporte['zoftinium_consumidas']:.1f}</div>
                        <div class="column-item">Horas disponibles: {reporte['zoftinium_disponibles']:.1f}</div>
                    </div>
                    
                    <div class="table-column">
                        <div class="column-header">Soporte Grupo Merpes</div>
                        <div class="column-item">Horas meta: {reporte['merpes_asignadas']:.0f}</div>
                        <div class="column-item">Horas consumidas: {reporte['merpes_consumidas']:.1f}</div>
                        <div class="column-item">Horas disponibles: {reporte['merpes_disponibles']:.1f}</div>
                    </div>
                </div>
            </div>
            
            <!-- Recomendaciones con icono real -->
            <div class="recomendaciones">
                <div class="recomendaciones-title">
                    <img src="data:image/png;base64,{icono2_b64}" alt="Recomendaciones" class="recomendaciones-icon">
                    Recomendaciones
                </div>
                <ul class="recomendaciones-list">
                    <li>Coordina con el equipo la optimización de las horas disponibles mes a mes.</li>
                    <li>Planifica los próximos requerimientos para que no superes la capacidad disponible.</li>
                    <li>Utiliza los recursos disponibles para que tu cliente tenga un mejor servicio.</li>
                    <li>No mal gastes las piezas gráficas, escales el mayor provecho.</li>
                    <li>Verifica que los datos que registra Kupaa sean los correctos, así mitigas el consumo de soporte.</li>
                </ul>
            </div>
        </div>
        
        <!-- Footer con iconos de contacto reales -->
        <div class="footer-section">
            <div class="contactanos">CONTÁCTANOS</div>
            <div class="iconos-contacto">
                <img src="data:image/png;base64,{icono1_b64}" alt="Contacto" class="icono-contacto">
                <img src="data:image/png;base64,{icono2_b64}" alt="Contacto" class="icono-contacto">
                <img src="data:image/png;base64,{icono1_b64}" alt="Contacto" class="icono-contacto">
                <img src="data:image/png;base64,{icono2_b64}" alt="Contacto" class="icono-contacto">
            </div>
        </div>
        
        <!-- Footer con imagen real -->
        <img src="data:image/png;base64,{footer_b64}" alt="Footer" class="footer-image">
    </div>
</body>
</html>
        """
        
        return asunto, cuerpo_html

    def enviar_email(self, correos_destino: List[str], asunto: str, cuerpo_html: str) -> bool:
        merpes_disponibles = reporte['merpes_disponibles']
        porcentaje_merpes = reporte['porcentaje_merpes']
        
        # Datos de Zoftinium
        zoftinium_asignadas = reporte['zoftinium_asignadas']
        zoftinium_consumidas = reporte['zoftinium_consumidas'] 
        zoftinium_disponibles = reporte['zoftinium_disponibles']
        porcentaje_zoftinium = reporte['porcentaje_zoftinium']
        
        # Datos de Diseño
        diseño_asignadas = reporte['diseño_asignadas']
        diseño_consumidas = reporte['diseño_consumidas']
        diseño_disponibles = reporte['diseño_disponibles'] 
        porcentaje_diseño = reporte['porcentaje_diseño']
        
        # Asunto del email
        asunto = f"� Reporte Mensual de Consumo - {cliente}"
        
        # Cuerpo HTML del email - Siguiendo el diseño de la imagen
        cuerpo_html = f"""
        <!DOCTYPE html>
        <html>
        <head>
            <meta charset="UTF-8">
            <style>
                body {{ font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; margin: 0; padding: 20px; background-color: #f5f7fa; }}
                .container {{ max-width: 800px; margin: 0 auto; background-color: white; border-radius: 15px; overflow: hidden; box-shadow: 0 8px 32px rgba(0,0,0,0.1); }}
                .header {{ background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); color: white; padding: 30px; text-align: center; }}
                .header h1 {{ margin: 0; font-size: 2.2em; font-weight: 300; }}
                .header h2 {{ margin: 10px 0 0 0; font-size: 1.4em; opacity: 0.9; }}
                .content {{ padding: 40px; }}
                .title {{ text-align: center; margin-bottom: 40px; }}
                .title h1 {{ color: #2c3e50; font-size: 2em; margin: 0; }}
                .title p {{ color: #7f8c8d; font-size: 1.1em; margin: 10px 0; }}
                
                .metrics-container {{ display: flex; justify-content: space-around; gap: 20px; margin: 40px 0; }}
                .metric-card {{ 
                    background: white;
                    border-radius: 20px;
                    padding: 30px 20px;
                    text-align: center;
                    box-shadow: 0 4px 20px rgba(0,0,0,0.08);
                    border: 2px solid #e8f4fd;
                    flex: 1;
                    min-width: 200px;
                }}
                
                .metric-card h3 {{
                    color: #2c3e50;
                    font-size: 0.9em;
                    text-transform: uppercase;
                    letter-spacing: 1px;
                    margin: 0 0 20px 0;
                    font-weight: 600;
                }}
                
                .percentage {{
                    font-size: 3.5em;
                    font-weight: bold;
                    margin: 10px 0;
                    line-height: 1;
                }}
                
                .merpes {{ color: #3498db; }}
                .zoftinium {{ color: #e74c3c; }}
                .diseño {{ color: #f39c12; }}
                
                .value-detail {{
                    color: #7f8c8d;
                    font-size: 1.1em;
                    margin-top: 15px;
                    font-weight: 500;
                }}
                
                .progress-ring {{
                    width: 120px;
                    height: 120px;
                    margin: 20px auto;
                    position: relative;
                }}
                
                .progress-circle {{
                    width: 100%;
                    height: 100%;
                    border-radius: 50%;
                    background: conic-gradient(var(--color) 0deg, var(--color) var(--degrees), #e0e0e0 var(--degrees), #e0e0e0 360deg);
                    display: flex;
                    align-items: center;
                    justify-content: center;
                }}
                
                .progress-inner {{
                    width: 80px;
                    height: 80px;
                    background: white;
                    border-radius: 50%;
                    display: flex;
                    align-items: center;
                    justify-content: center;
                    font-weight: bold;
                    font-size: 1.2em;
                    color: var(--color);
                }}
                
                .summary {{ background-color: #f8f9fa; padding: 30px; border-radius: 10px; margin: 30px 0; }}
                .summary h3 {{ color: #2c3e50; margin-bottom: 20px; }}
                .summary-item {{ display: flex; justify-content: space-between; margin: 15px 0; padding: 10px 0; border-bottom: 1px solid #e9ecef; }}
                .summary-item:last-child {{ border-bottom: none; }}
                
                .footer {{ background-color: #2c3e50; color: white; padding: 30px; text-align: center; }}
                .footer p {{ margin: 5px 0; opacity: 0.8; }}
                
                @media (max-width: 600px) {{
                    .metrics-container {{ flex-direction: column; }}
                    .metric-card {{ margin-bottom: 20px; }}
                }}
            </style>
        </head>
        <body>
            <div class="container">
                <div class="header">
                    <h1>Reporte Mensual</h1>
                    <h2>{cliente}</h2>
                </div>
                
                <div class="content">
                    <div class="title">
                        <h1>CONSUMO DEL MES EN SOPORTE Y DISEÑO</h1>
                        <p>Resumen de actividades - {datetime.now().strftime('%B %Y')}</p>
                    </div>
                    
                    <div class="metrics-container">
                        <div class="metric-card">
                            <h3>Horas Soporte Merpes</h3>
                            <div class="percentage merpes">{porcentaje_merpes:.0f}%</div>
                            <div class="progress-ring">
                                <div class="progress-circle" style="--color: #3498db; --degrees: {porcentaje_merpes * 3.6}deg;">
                                    <div class="progress-inner" style="--color: #3498db;">{merpes_consumidas}</div>
                                </div>
                            </div>
                            <div class="value-detail">{merpes_consumidas} horas utilizadas</div>
                        </div>
                        
                        <div class="metric-card">
                            <h3>Horas Soporte Zoftinium</h3>
                            <div class="percentage zoftinium">{porcentaje_zoftinium:.0f}%</div>
                            <div class="progress-ring">
                                <div class="progress-circle" style="--color: #e74c3c; --degrees: {porcentaje_zoftinium * 3.6}deg;">
                                    <div class="progress-inner" style="--color: #e74c3c;">{zoftinium_consumidas}</div>
                                </div>
                            </div>
                            <div class="value-detail">{zoftinium_consumidas} horas utilizadas</div>
                        </div>
                        
                        <div class="metric-card">
                            <h3>Piezas de Diseño</h3>
                            <div class="percentage diseño">{porcentaje_diseño:.0f}%</div>
                            <div class="progress-ring">
                                <div class="progress-circle" style="--color: #f39c12; --degrees: {porcentaje_diseño * 3.6}deg;">
                                    <div class="progress-inner" style="--color: #f39c12;">{diseño_consumidas}</div>
                                </div>
                            </div>
                            <div class="value-detail">{diseño_consumidas} piezas realizadas</div>
                        </div>
                    </div>
                    
                    <div class="summary">
                        <h3>📈 Resumen Detallado</h3>
                        <div class="summary-item">
                            <span><strong>Total Horas Soporte Merpes:</strong></span>
                            <span>{merpes_consumidas} horas</span>
                        </div>
                        <div class="summary-item">
                            <span><strong>Total Horas Soporte Zoftinium:</strong></span>
                            <span>{zoftinium_consumidas} horas</span>
                        </div>
                        <div class="summary-item">
                            <span><strong>Total Piezas de Diseño:</strong></span>
                            <span>{diseño_consumidas} piezas</span>
                        </div>
                        <div class="summary-item">
                            <span><strong>Total Horas de Soporte:</strong></span>
                            <span><strong>{merpes_consumidas + zoftinium_consumidas} horas</strong></span>
                        </div>
                    </div>
                    
                    <div style="background-color: #e8f6f3; padding: 20px; border-left: 4px solid #27ae60; border-radius: 5px; margin: 30px 0;">
                        <h4>💡 Información Importante:</h4>
                        <ul>
                            <li>Los datos reflejan el consumo acumulado del mes actual</li>
                            <li>Las horas de soporte incluyen desarrollo, mantenimiento y resolución de incidencias</li>
                            <li>Las piezas de diseño corresponden a elementos gráficos creados y entregados</li>
                            <li>Para consultas adicionales, contacte al equipo de desarrollo</li>
                        </ul>
                    </div>
                </div>
                
                <div class="footer">
                    <p><strong>Sistema de Reportes Automático</strong></p>
                    <p>Generado automáticamente el {datetime.now().strftime('%d/%m/%Y a las %H:%M:%S')}</p>
                    <p>Grupo Merpes - Desarrollo Web</p>
                </div>
            </div>
        </body>
        </html>
        """
        
        return asunto, cuerpo_html
    
    def enviar_email(self, correos_destino: List[str], asunto: str, cuerpo_html: str) -> bool:
        """
        Envía un email de reporte a múltiples destinatarios
        
        Args:
            correos_destino: Lista de emails destinatarios
            asunto: Asunto del email
            cuerpo_html: Cuerpo del email en HTML
            
        Returns:
            bool: True si se envió exitosamente
        """
        try:
            # Crear mensaje
            msg = MIMEMultipart('alternative')
            msg['Subject'] = asunto
            msg['From'] = self.config_email['email']
            msg['To'] = ', '.join(correos_destino)  # Múltiples destinatarios
            
            # Agregar cuerpo HTML
            part_html = MIMEText(cuerpo_html, 'html', 'utf-8')
            msg.attach(part_html)
            
            # Conectar al servidor SMTP
            with smtplib.SMTP(self.config_email['smtp_server'], self.config_email['puerto']) as server:
                server.starttls()
                server.login(self.config_email['email'], self.config_email['password'])
                server.send_message(msg)
            
            destinatarios_str = ', '.join(correos_destino)
            logging.info(f"Email enviado exitosamente a: {destinatarios_str}")
            return True
            
        except Exception as e:
            logging.error(f"Error al enviar email a {correos_destino}: {e}")
            return False
    
    def procesar_alertas(self, reportes: List[Dict]) -> Dict:
        """
        Procesa y envía todos los reportes mensuales
        
        Args:
            reportes: Lista de reportes a procesar
            
        Returns:
            Dict: Estadísticas de envío
        """
        estadisticas = {
            'total_reportes': len(reportes),
            'enviados_exitosamente': 0,
            'errores': 0,
            'clientes_procesados': []
        }
        
        for reporte in reportes:
            try:
                # Generar contenido del email
                asunto, cuerpo_html = self.generar_contenido_email(reporte)
                
                # Enviar email
                if self.enviar_email(reporte['correos_destino'], asunto, cuerpo_html):
                    estadisticas['enviados_exitosamente'] += 1
                    estadisticas['clientes_procesados'].append(reporte['cliente'])
                else:
                    estadisticas['errores'] += 1
                
                # Pausa entre envíos para evitar spam
                time.sleep(1)
                
            except Exception as e:
                logging.error(f"Error al procesar reporte para {reporte['cliente']}: {e}")
                estadisticas['errores'] += 1
        
        return estadisticas
    
    def ejecutar_monitoreo(self):
        """
        Ejecuta el proceso completo de monitoreo
        """
        try:
            logging.info("=== INICIANDO MONITOREO DE CAMPAÑAS ===")
            
            # 1. Descargar archivo Excel
            archivo_excel = self.descargar_excel()
            
            # 2. Leer datos de consumo
            self.df_consumo = self.leer_datos_consumo(archivo_excel)
            
            # 3. Cargar mapeo de correos
            self.cargar_mapeo_correos()
            
            # 4. Analizar y generar reportes
            reportes = self.analizar_estado_campañas()
            
            if not reportes:
                logging.info("No se encontraron clientes con correos definidos para generar reportes")
                return
            
            # 5. Procesar y enviar reportes
            estadisticas = self.procesar_alertas(reportes)
            
            # 6. Mostrar resumen
            logging.info("=== RESUMEN DE EJECUCIÓN ===")
            logging.info(f"Total de reportes generados: {estadisticas['total_reportes']}")
            logging.info(f"Emails enviados exitosamente: {estadisticas['enviados_exitosamente']}")
            logging.info(f"Errores en envío: {estadisticas['errores']}")
            logging.info(f"Clientes procesados: {', '.join(estadisticas['clientes_procesados'])}")
            
            # 7. Limpiar archivo temporal
            try:
                os.remove(archivo_excel)
                logging.info("Archivo temporal eliminado")
            except:
                pass
                
        except Exception as e:
            logging.error(f"Error en la ejecución del monitoreo: {e}")
            raise

def main():
    """
    Función principal - Configuración y ejecución
    """
    # Configuración del archivo Excel
    EXCEL_URL = "https://tu-link-del-archivo-excel.com/archivo.xlsx"
    
    # Configuración del servidor de email
    CONFIG_EMAIL = {
        'smtp_server': 'smtp.gmail.com',  # Para Gmail
        'puerto': 587,
        'email': 'tu-email@gmail.com',
        'password': 'tu-contraseña-de-aplicación'  # Usar contraseña de aplicación, no la normal
    }
    
    # Crear instancia del monitor
    monitor = MonitorCampañas(EXCEL_URL, CONFIG_EMAIL)
    
    # Ejecutar monitoreo
    monitor.ejecutar_monitoreo()

if __name__ == "__main__":
    main()
