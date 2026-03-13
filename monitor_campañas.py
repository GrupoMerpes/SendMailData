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
            df = pd.read_excel(archivo_excel, header=1)
            
            # Mapear las columnas basándose en la estructura real del Excel
            column_mapping = {
                'DICIEMBRE': 'Cliente',
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
                'Correos': 'CORREOS'
            }
            
            # Aplicar el mapeo de columnas solo para las que existen
            existing_mappings = {old_col: new_col for old_col, new_col in column_mapping.items() if old_col in df.columns}
            df = df.rename(columns=existing_mappings)
            logging.info(f"Columnas mapeadas: {existing_mappings}")
            
            # Filtrar la fila de encabezados (fila 0 dice "Cliente", "Asignadas", etc)
            df = df[df['Cliente'] != 'Cliente']
            df = df[df['Cliente'].notna()]
            
            # Limpiar y convertir datos
            df['Cliente'] = df['Cliente'].astype(str).str.strip()
            
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
            print(df.columns.tolist())

            df = df[df['Cliente'] != '']
            
            logging.info(f"Datos leídos exitosamente. Total Clientes: {len(df)}")
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
                print('OE CARE MONDAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA', self.df_consumo.columns.tolist())
                filas_con_correos = self.df_consumo[self.df_consumo['Correos'].notna()]
                print("Filas con correos encontrados:")
                print(filas_con_correos)

                for _, row in filas_con_correos.iterrows():
                    Cliente = str(row['Cliente']).strip()
                    correos_str = str(row['CORREOS']).strip()
                    
                    # Skip if Cliente is empty, NaN, or just numbers/invalid
                    if not Cliente or Cliente in ['0', 'nan', 'NaN', ''] or Cliente.isdigit():
                        logging.warning(f"Saltando Cliente inválido: '{Cliente}'")
                        continue
                    
                    # Si hay múltiples correos separados por ; o ,
                    correos_lista = [email.strip() for email in correos_str.replace(';', ',').split(',') if email.strip()]
                    
                    if correos_lista:
                        self.campañas_correos[Cliente] = correos_lista
                
                logging.info(f"Mapeo de correos cargado: {len(self.campañas_correos)} Clientes con correos definidos")
                
                # Mostrar el mapeo encontrado
                for Cliente, correos in self.campañas_correos.items():
                    logging.info(f"  • {Cliente}: {', '.join(correos)}")
                    
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
        Genera reportes mensuales para cada Cliente con correos definidos
        
        Returns:
            List[Dict]: Lista de reportes a enviar
        """
        reportes = []
        
        # Solo procesar Clientes que tienen correos definidos
        for Cliente, correos_lista in self.campañas_correos.items():
            # Buscar datos del Cliente
            datos_Cliente = self.df_consumo[self.df_consumo['Cliente'] == Cliente]
            
            if not datos_Cliente.empty:
                # Sumar todos los valores para el Cliente (puede tener múltiples filas)
                merpes_asignadas = datos_Cliente['HORAS_MERPES_ASIGNADAS'].sum()
                merpes_consumidas = datos_Cliente['HORAS_MERPES_CONSUMIDAS'].sum() 
                merpes_disponibles = datos_Cliente['HORAS_MERPES_DISPONIBLES'].sum()
                
                zoftinium_asignadas = datos_Cliente['HORAS_ZOFTINIUM_ASIGNADAS'].sum()
                zoftinium_consumidas = datos_Cliente['HORAS_ZOFTINIUM_CONSUMIDAS'].sum()
                zoftinium_disponibles = datos_Cliente['HORAS_ZOFTINIUM_DISPONIBLES'].sum()
                
                diseño_asignadas = datos_Cliente['PIEZAS_DISEÑO_ASIGNADAS'].sum()
                diseño_consumidas = datos_Cliente['PIEZAS_DISEÑO_CONSUMIDAS'].sum() 
                diseño_disponibles = datos_Cliente['PIEZAS_DISEÑO_DISPONIBLES'].sum()
                
                # Calcular porcentajes reales
                porcentaje_merpes = (merpes_consumidas / merpes_asignadas * 100) if merpes_asignadas > 0 else 0
                porcentaje_zoftinium = (zoftinium_consumidas / zoftinium_asignadas * 100) if zoftinium_asignadas > 0 else 0
                porcentaje_diseño = (diseño_consumidas / diseño_asignadas * 100) if diseño_asignadas > 0 else 0
                
                # Crear reporte para este Cliente
                area_campaña = datos_Cliente['AREA_CAMPAÑA'].iloc[0] if 'AREA_CAMPAÑA' in datos_Cliente.columns and not datos_Cliente['AREA_CAMPAÑA'].isna().all() else 'Desarrollo Web'
                
                reporte = {
                    'Cliente': Cliente,
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
                
                logging.info(f"Reporte generado para {Cliente}: Merpes {merpes_consumidas}/{merpes_asignadas}h ({porcentaje_merpes:.1f}%), Zoftinium {zoftinium_consumidas}/{zoftinium_asignadas}h ({porcentaje_zoftinium:.1f}%), Diseño {diseño_consumidas}/{diseño_asignadas} piezas ({porcentaje_diseño:.1f}%)")
        
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
        Genera el contenido del email con máxima compatibilidad para email clients
        """
        from datetime import datetime
        import base64
        import os
        
        Cliente = reporte['Cliente']
        area_campaña = reporte.get('area_campaña', 'Desarrollo Web')
        
        # Asunto del email
        asunto = f"Reporte mensual soporte y diseño + {Cliente}"
        
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
        
        # Convertir todas las imágenes a base64 para evitar que desaparezcan
        encabezado_b64 = imagen_to_base64("ENCABEZADO.png")
        footer_b64 = imagen_to_base64("FOOTER.png") 
        icono1_b64 = imagen_to_base64("ICONO1.png")
        icono2_b64 = imagen_to_base64("ICONO2.png")
        recomendaciones_b64 = imagen_to_base64("MAIL-REPORTE_06.png")
        
        # HTML ultra-compatible usando técnicas híbridas para máxima compatibilidad
        cuerpo_html = f"""
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1.0" />
    <title>Reporte Mensual - {Cliente}</title>
    <!--[if gte mso 9]>
    <xml>
        <o:OfficeDocumentSettings>
            <o:AllowPNG/>
            <o:PixelsPerInch>96</o:PixelsPerInch>
        </o:OfficeDocumentSettings>
    </xml>
    <![endif]-->
    <style type="text/css">
        /* Reset para email */
        body, table, td, p, a, li {{ margin: 0; padding: 0; border: 0; }}
        table {{ border-collapse: collapse !important; mso-table-lspace: 0pt; mso-table-rspace: 0pt; }}
        img {{ border: 0; height: auto; line-height: 100%; outline: none; text-decoration: none; -ms-interpolation-mode: bicubic; }}
        
        /* Clases específicas */
        .email-body {{ margin: 0 !important; padding: 0 !important; background-color: #f0f0f0; }}
        .email-container {{ width: 600px !important; max-width: 600px !important; }}
        
        /* Responsive para móvil */
        @media only screen and (max-width: 600px) {{
            .email-container {{ width: 100% !important; max-width: 100% !important; }}
            .mobile-stack {{ display: block !important; width: 100% !important; }}
        }}
        
        /* Outlook específico */
        <!--[if mso]>
        table {{ border-collapse: collapse; border-spacing: 0; }}
        .fallback-font {{ font-family: Arial, sans-serif !important; }}
        <![endif]-->
    </style>
</head>
<body class="email-body" style="margin: 0; padding: 0; background-color: #f0f0f0;">
    
    <!-- Wrapper para centrar todo el contenido -->
    <!--[if mso | IE]>
    <table role="presentation" border="0" cellpadding="0" cellspacing="0" width="100%" style="background-color: #f0f0f0;">
        <tr>
            <td align="center">
    <![endif]-->
    
    <div align="center" style="background-color: #f0f0f0; padding: 20px 0;">
        
        <!-- Contenedor principal con ancho fijo -->
        <table role="presentation" class="email-container" cellspacing="0" cellpadding="0" border="0" style="width: 600px; max-width: 600px; margin: 0 auto; background-color: #ffffff; border-collapse: collapse;">
            
            <!-- HEADER IMAGE -->
            <tr>
                <td align="center" style="padding: 0; margin: 0; line-height: 0; font-size: 0;">
                    <table role="presentation" cellspacing="0" cellpadding="0" border="0" style="border-collapse: collapse;">
                        <tr>
                            <td style="padding: 0; margin: 0; line-height: 0; font-size: 0;">
                                <img src="data:image/png;base64,{encabezado_b64}" alt="Header Merpes-Kupaa" style="display: block; width: 600px; height: auto; border: 0; outline: none; text-decoration: none;" width="600" />
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
            
            <!-- HERO SECTION -->
            <tr>
                <td style="padding: 30px 40px; text-align: center; background-color: #ffffff;">
                    
                    <!-- Título superior -->
                    <table role="presentation" width="100%" cellspacing="0" cellpadding="0" border="0" style="border-collapse: collapse;">
                        <tr>
                            <td style="font-family: Arial, sans-serif; font-size: 14px; color: black; text-align: center; padding-bottom: 10px; font-weight: normal;">
                                CONOCE EL ESTADO DE
                            </td>
                        </tr>
                        <tr>
                            <td style="font-family: Arial, sans-serif; font-size: 24px; color: #2d3748; text-align: center; padding-bottom: 20px; font-weight: bold;">
                                TUS REQUERIMIENTOS
                            </td>
                        </tr>
                    </table>
                    
                    <!-- Badge de campaña - Centrado -->
                    <div align="center">
                        <table role="presentation" cellspacing="0" cellpadding="0" border="0" style="border-collapse: collapse; margin: 0 auto;">
                            <tr>
                                <td style="background-color: #33ffa7; padding: 15px 25px; text-align: center;">
                                    <table role="presentation" cellspacing="0" cellpadding="0" border="0" style="border-collapse: collapse;">
                                        <tr>
                                            <td style="font-family: Arial, sans-serif; font-size: 10px; color: #000000; text-align: center; font-weight: bold; padding-bottom: 3px;">
                                                CAMPAÑA ACTIVA
                                            </td>
                                        </tr>
                                        <tr>
                                            <td style="font-family: Arial, sans-serif; font-size: 14px; color: #000000; text-align: center; font-weight: bold;">
                                                {Cliente} - {area_campaña}
                                            </td>
                                        </tr>
                                    </table>
                                </td>
                            </tr>
                        </table>
                    </div>
                    
                </td>
            </tr>
            
            <!-- CONTENT SECTION -->
            <tr>
                <td style="padding: 20px 30px; background-color: #ffffff;">
                    
                    <!-- Título de consumo -->
                    <table role="presentation" width="100%" cellspacing="0" cellpadding="0" border="0" style="border-collapse: collapse;">
                        <tr>
                            <td style="font-family: Arial, sans-serif; font-size: 16px; color: #555555; text-align: center; padding: 20px 0; font-weight: bold;">
                                CONSUMO DEL MES EN SOPORTE Y DISEÑO <br>
                                CONSOLIDADO HASTA X ENERO
                            </td>
                        </tr>
                    </table>
                    
                    <!-- CARDS ROW - Centradas perfectamente -->
                    <table role="presentation" cellspacing="0" cellpadding="0" border="0" align="center" style="border-collapse: collapse;">
                        <tr>
                            <!-- Card 1: Piezas de Diseño -->
                            <td style="padding-right: 15px; vertical-align: top;">
                                <table role="presentation" cellspacing="0" cellpadding="0" border="0" style="width: 160px; border-collapse: collapse; background-color: #f8f9fa; border: 2px solid #e9ecef;">
                                    <tr>
                                        <td style="padding: 20px 10px; text-align: center; height: 90px;">
                                            <table role="presentation" width="100%" cellspacing="0" cellpadding="0" border="0" style="border-collapse: collapse;">
                                                <tr>
                                                    <td style="font-family: Arial, sans-serif; font-size: 10px; color: #6c757d; text-align: center; font-weight: bold; padding-bottom: 15px;">
                                                        PIEZAS DE DISEÑO
                                                    </td>
                                                </tr>
                                                <tr>
                                                    <td style="font-family: Arial, sans-serif; font-size: 28px; color: #495057; text-align: center; font-weight: bold;">
                                                        {reporte['porcentaje_diseño']:.0f}%
                                                    </td>
                                                </tr>
                                            </table>
                                        </td>
                                    </tr>
                                </table>
                            </td>
                            
                            <!-- Card 2: Soporte Zoftinium -->
                            <td style="padding-right: 15px; vertical-align: top;">
                                <table role="presentation" cellspacing="0" cellpadding="0" border="0" style="width: 160px; border-collapse: collapse; background-color: #f8f9fa; border: 2px solid #e9ecef;">
                                    <tr>
                                        <td style="padding: 20px 10px; text-align: center; height: 90px;">
                                            <table role="presentation" width="100%" cellspacing="0" cellpadding="0" border="0" style="border-collapse: collapse;">
                                                <tr>
                                                    <td style="font-family: Arial, sans-serif; font-size: 10px; color: #6c757d; text-align: center; font-weight: bold; padding-bottom: 5px;">
                                                        SOPORTE KUPAA
                                                    </td>
                                                </tr>
                                                <tr>
                                                    <td style="font-family: Arial, sans-serif; font-size: 10px; color: #6c757d; text-align: center; font-weight: bold; padding-bottom: 10px;">
                                                        ZOFTINIUM
                                                    </td>
                                                </tr>
                                                <tr>
                                                    <td style="font-family: Arial, sans-serif; font-size: 28px; color: #495057; text-align: center; font-weight: bold;">
                                                        {reporte['porcentaje_zoftinium']:.0f}%
                                                    </td>
                                                </tr>
                                            </table>
                                        </td>
                                    </tr>
                                </table>
                            </td>
                            
                            <!-- Card 3: Equipo Merpes -->
                            <td style="vertical-align: top;">
                                <table role="presentation" cellspacing="0" cellpadding="0" border="0" style="width: 160px; border-collapse: collapse; background-color: #f8f9fa; border: 2px solid #e9ecef;">
                                    <tr>
                                        <td style="padding: 20px 10px; text-align: center; height: 90px;">
                                            <table role="presentation" width="100%" cellspacing="0" cellpadding="0" border="0" style="border-collapse: collapse;">
                                                <tr>
                                                    <td style="font-family: Arial, sans-serif; font-size: 10px; color: #6c757d; text-align: center; font-weight: bold; padding-bottom: 5px;">
                                                        SOPORTE KUPAA
                                                    </td>
                                                </tr>
                                                <tr>
                                                    <td style="font-family: Arial, sans-serif; font-size: 10px; color: #6c757d; text-align: center; font-weight: bold; padding-bottom: 10px;">
                                                        EQUIPO MERPES
                                                    </td>
                                                </tr>
                                                <tr>
                                                    <td style="font-family: Arial, sans-serif; font-size: 28px; color: #495057; text-align: center; font-weight: bold;">
                                                        {reporte['porcentaje_merpes']:.0f}%
                                                    </td>
                                                </tr>
                                            </table>
                                        </td>
                                    </tr>
                                </table>
                            </td>
                        </tr>
                    </table>
                    
                </td>
            </tr>
            
            <!-- SPACER -->
            <tr>
                <td style="height: 20px; line-height: 20px; font-size: 20px;">&nbsp;</td>
            </tr>            
            <!-- DATA TABLE -->
            <tr>
                <td style="padding: 0 30px 20px;">
                    
                    <!-- Tabla con ancho fijo de 540px para evitar estiramientos -->
                    <table role="presentation" cellspacing="0" cellpadding="0" border="0" style="width: 540px; border-collapse: collapse; background-color: #2d2c2c; margin: 0 auto;">
                        
                        <!-- Header con icono -->
                        <tr>
                            <td colspan="3" style="background-color: #2d2c2c; padding: 15px; text-align: center; border-bottom: 1px solid #444444;">
                                <img src="data:image/png;base64,{icono1_b64}" alt="Estadísticas" style="display: block; width: 24px; height: 24px; margin: 0 auto; border: 0;" width="24" height="24" />
                            </td>
                        </tr>
                        
                        <!-- Headers de columnas -->
                        <tr style="background-color: #3a3939;">
                            <td style="width: 180px; font-family: Arial, sans-serif; font-size: 10px; font-weight: bold; text-transform: uppercase; color: #ffffff; text-align: center; padding: 12px 8px; border-right: 1px solid #444444;">
                                PIEZAS DISEÑO
                            </td>
                            <td style="width: 180px; font-family: Arial, sans-serif; font-size: 10px; font-weight: bold; text-transform: uppercase; color: #ffffff; text-align: center; padding: 12px 8px; border-right: 1px solid #444444;">
                                SOPORTE ZOFTINIUM
                            </td>
                            <td style="width: 180px; font-family: Arial, sans-serif; font-size: 10px; font-weight: bold; text-transform: uppercase; color: #ffffff; text-align: center; padding: 12px 8px;">
                                SOPORTE GRUPO MERPES
                            </td>
                        </tr>
                        
                        <!-- Fila 1: Metas -->
                        <tr>
                            <td style="font-family: Arial, sans-serif; font-size: 11px; color: #ffffff; text-align: center; padding: 10px 8px; border-right: 1px solid #444444; border-bottom: 1px solid #444444; background-color: #2d2c2c;">
                                Piezas mensuales: {reporte['diseño_asignadas']:.0f}
                            </td>
                            <td style="font-family: Arial, sans-serif; font-size: 11px; color: #ffffff; text-align: center; padding: 10px 8px; border-right: 1px solid #444444; border-bottom: 1px solid #444444; background-color: #2d2c2c;">
                                Horas mensuales: {reporte['zoftinium_asignadas']:.0f}
                            </td>
                            <td style="font-family: Arial, sans-serif; font-size: 11px; color: #ffffff; text-align: center; padding: 10px 8px; border-bottom: 1px solid #444444; background-color: #2d2c2c;">
                                Horas mensuales: {reporte['merpes_asignadas']:.0f}
                            </td>
                        </tr>
                        
                        <!-- Fila 2: Consumidas -->
                        <tr>
                            <td style="font-family: Arial, sans-serif; font-size: 11px; color: #ffffff; text-align: center; padding: 10px 8px; border-right: 1px solid #444444; border-bottom: 1px solid #444444; background-color: #2d2c2c;">
                                Piezas consumidas: {reporte['diseño_consumidas']:.0f}
                            </td>
                            <td style="font-family: Arial, sans-serif; font-size: 11px; color: #ffffff; text-align: center; padding: 10px 8px; border-right: 1px solid #444444; border-bottom: 1px solid #444444; background-color: #2d2c2c;">
                                Horas consumidas: {reporte['zoftinium_consumidas']:.0f}
                            </td>
                            <td style="font-family: Arial, sans-serif; font-size: 11px; color: #ffffff; text-align: center; padding: 10px 8px; border-bottom: 1px solid #444444; background-color: #2d2c2c;">
                                Horas consumidas: {reporte['merpes_consumidas']:.0f}
                            </td>
                        </tr>
                        
                        <!-- Fila 3: Disponibles -->
                        <tr>
                            <td style="font-family: Arial, sans-serif; font-size: 11px; color: #ffffff; text-align: center; padding: 10px 8px; border-right: 1px solid #444444; border-bottom: 1px solid #444444; background-color: #2d2c2c;">
                                Piezas disponibles: {reporte['diseño_disponibles']:.0f}
                            </td>
                            <td style="font-family: Arial, sans-serif; font-size: 11px; color: #ffffff; text-align: center; padding: 10px 8px; border-right: 1px solid #444444; border-bottom: 1px solid #444444; background-color: #2d2c2c;">
                                Horas disponibles: {reporte['zoftinium_disponibles']:.0f}
                            </td>
                            <td style="font-family: Arial, sans-serif; font-size: 11px; color: #ffffff; text-align: center; padding: 10px 8px; border-bottom: 1px solid #444444; background-color: #2d2c2c;">
                                Horas disponibles: {reporte['merpes_disponibles']:.0f}
                            </td>
                        </tr>

                    </table>
                    
                </td>
            </tr>
            
            <!-- SPACER -->
            <tr>
                <td style="height: 30px; line-height: 30px; font-size: 30px;">&nbsp;</td>
            </tr>
            
            <!-- RECOMMENDATIONS SECTION -->
            <tr>
                <td style="padding: 0 30px 30px;">
                    
                    <!-- Título con líneas decorativas -->
                    <table role="presentation" width="100%" cellspacing="0" cellpadding="0" border="0" style="border-collapse: collapse; margin-bottom: 20px;">
                        <tr>
                            <td width="35%" style="border-bottom: 2px solid #cccccc; height: 1px;"></td>
                            <td style="text-align: center; padding: 0 15px;">
                                <table role="presentation" cellspacing="0" cellpadding="0" border="0" align="center" style="border-collapse: collapse;">
                                    <tr>
                                        <td style="padding-right: 8px;">
                                            <img src="data:image/png;base64,{icono2_b64}" alt="Recomendaciones" style="display: block; width: 18px; height: 18px; border: 0;" width="18" height="18" />
                                        </td>
                                        <td style="font-family: Arial, sans-serif; font-size: 14px; font-weight: bold; color: #333333;">
                                            Recomendaciones
                                        </td>
                                    </tr>
                                </table>
                            </td>
                            <td width="35%" style="border-bottom: 2px solid #cccccc; height: 1px;"></td>
                        </tr>
                    </table>
                    
                    <!-- Lista de recomendaciones -->
                    <table role="presentation" width="100%" cellspacing="0" cellpadding="0" border="0" style="border-collapse: collapse;">
                        <tr>
                            <td style="padding: 0 20px;">
                                
                                <table role="presentation" width="100%" cellspacing="0" cellpadding="0" border="0" style="border-collapse: collapse;">
                                    <tr>
                                        <td style="padding: 6px 0;">
                                            <table role="presentation" cellspacing="0" cellpadding="0" border="0" style="border-collapse: collapse;">
                                                <tr>
                                                    <td width="15" style="font-family: Arial, sans-serif; font-size: 12px; font-weight: bold; color: #333333; padding-right: 8px;">▶</td>
                                                    <td style="font-family: Arial, sans-serif; font-size: 13px; color: #555555; line-height: 18px;">
                                                        Coordina con el equipo la optimización de las horas disponibles mes a mes.
                                                    </td>
                                                </tr>
                                            </table>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td style="padding: 6px 0;">
                                            <table role="presentation" cellspacing="0" cellpadding="0" border="0" style="border-collapse: collapse;">
                                                <tr>
                                                    <td width="15" style="font-family: Arial, sans-serif; font-size: 12px; font-weight: bold; color: #333333; padding-right: 8px;">▶</td>
                                                    <td style="font-family: Arial, sans-serif; font-size: 13px; color: #555555; line-height: 18px;">
                                                        Planifica los próximos requerimientos para que no superes la capacidad disponible.
                                                    </td>
                                                </tr>
                                            </table>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td style="padding: 6px 0;">
                                            <table role="presentation" cellspacing="0" cellpadding="0" border="0" style="border-collapse: collapse;">
                                                <tr>
                                                    <td width="15" style="font-family: Arial, sans-serif; font-size: 12px; font-weight: bold; color: #333333; padding-right: 8px;">▶</td>
                                                    <td style="font-family: Arial, sans-serif; font-size: 13px; color: #555555; line-height: 18px;">
                                                        Utiliza los recursos disponibles para que tu Cliente cuente con un excelente servicio.
                                                    </td>
                                                </tr>
                                            </table>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td style="padding: 6px 0;">
                                            <table role="presentation" cellspacing="0" cellpadding="0" border="0" style="border-collapse: collapse;">
                                                <tr>
                                                    <td width="15" style="font-family: Arial, sans-serif; font-size: 12px; font-weight: bold; color: #333333; padding-right: 8px;">▶</td>
                                                    <td style="font-family: Arial, sans-serif; font-size: 13px; color: #555555; line-height: 18px;">
                                                        No mal gastes las piezas gráficas, sácales el mayor provecho.
                                                    </td>
                                                </tr>
                                            </table>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td style="padding: 6px 0;">
                                            <table role="presentation" cellspacing="0" cellpadding="0" border="0" style="border-collapse: collapse;">
                                                <tr>
                                                    <td width="15" style="font-family: Arial, sans-serif; font-size: 12px; font-weight: bold; color: #333333; padding-right: 8px;">▶</td>
                                                    <td style="font-family: Arial, sans-serif; font-size: 13px; color: #555555; line-height: 18px;">
                                                        Verifica que los cargues de data que hagas a Kupaa sean los correctos, así mitigas el consumo de soporte.
                                                    </td>
                                                </tr>
                                            </table>
                                        </td>
                                    </tr>
                                </table>
                                
                            </td>
                        </tr>
                    </table>
                    
                </td>
            </tr>
            
            <!-- SPACER antes del footer -->
            <tr>
                <td style="height: 20px; line-height: 20px; font-size: 20px;">&nbsp;</td>
            </tr>
            
            
            <!-- SPACER antes del footer -->
            <tr>
                <td style="height: 30px; line-height: 30px; font-size: 30px;">&nbsp;</td>
            </tr>
            
            <!-- FOOTER IMAGE - Con atributos específicos para evitar cortes al actualizar -->
            <tr>
                <td align="center" style="padding: 0; margin: 0; line-height: 0; font-size: 0;">
                    <table role="presentation" cellspacing="0" cellpadding="0" border="0" style="border-collapse: collapse;">
                        <tr>
                            <td style="padding: 0; margin: 0; line-height: 0; font-size: 0;">
                                <img src="data:image/png;base64,{footer_b64}" alt="Footer Merpes-Kupaa" 
                                     style="display: block; width: 600px; height: auto; border: 0; outline: none; text-decoration: none; max-width: none !important; min-width: 600px;" 
                                     width="600" height="" border="0" />
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
            
            <!-- SPACER después del footer -->
            <tr>
                <td style="height: 30px; line-height: 30px; font-size: 30px;">&nbsp;</td>
            </tr>
            
            <!-- SPACER después del footer -->
            <tr>
                <td style="height: 20px; line-height: 20px; font-size: 20px;">&nbsp;</td>
            </tr>
            
        </table>
        <!-- End Email Container -->
        
    </div>
    
    <!--[if mso | IE]>
            </td>
        </tr>
    </table>
    <![endif]-->
    
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
            'Clientes_procesados': []
        }
        
        for reporte in reportes:
            try:
                # Generar contenido del email
                asunto, cuerpo_html = self.generar_contenido_email(reporte)
                
                # Enviar email
                if self.enviar_email(reporte['correos_destino'], asunto, cuerpo_html):
                    estadisticas['enviados_exitosamente'] += 1
                    estadisticas['Clientes_procesados'].append(reporte['Cliente'])
                else:
                    estadisticas['errores'] += 1
                
                # Pausa entre envíos para evitar spam
                time.sleep(1)
                
            except Exception as e:
                logging.error(f"Error al procesar reporte para {reporte['Cliente']}: {e}")
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
                logging.info("No se encontraron Clientes con correos definidos para generar reportes")
                return
            
            # 5. Procesar y enviar reportes
            estadisticas = self.procesar_alertas(reportes)
            
            # 6. Mostrar resumen
            logging.info("=== RESUMEN DE EJECUCIÓN ===")
            logging.info(f"Total de reportes generados: {estadisticas['total_reportes']}")
            logging.info(f"Emails enviados exitosamente: {estadisticas['enviados_exitosamente']}")
            logging.info(f"Errores en envío: {estadisticas['errores']}")
            logging.info(f"Clientes procesados: {', '.join(estadisticas['Clientes_procesados'])}")
            
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
