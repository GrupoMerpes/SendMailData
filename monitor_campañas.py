#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import pandas as pd
import smtplib
import requests
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import logging
from datetime import datetime
import os
from typing import Dict, List
import re

# Configuración de logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[logging.FileHandler('monitor_campañas.log'), logging.StreamHandler()]
)

class MonitorCampañas:
    def __init__(self, excel_url: str, config_email: Dict):
        self.excel_url = excel_url
        self.config_email = config_email
        self.df_consumo = None
        self.campañas_correos = {}

    def descargar_excel(self) -> str:
        if os.path.exists(self.excel_url): return self.excel_url
        try:
            logging.info(f"Descargando archivo...")
            response = requests.get(self.excel_url, headers={'User-Agent': 'Mozilla/5.0'}, timeout=30)
            response.raise_for_status()
            filename = f"consumo_temp.xlsx"
            with open(filename, 'wb') as f: f.write(response.content)
            return filename
        except Exception as e:
            logging.error(f"Error descarga: {e}"); raise

    def leer_datos_consumo(self, archivo_excel: str) -> pd.DataFrame:
        try:
            df = pd.read_excel(archivo_excel, header=1)
            mapeo = {
                'Asignadas': 'H_M_A', 'Consumidas': 'H_M_C',
                'Asignadas.1': 'H_Z_A', 'Consumidas.1': 'H_Z_C',
                'Asignadas.2': 'P_D_A', 'Consumidas.2': 'P_D_C',
                'Unnamed: 11': 'Correos', 'Correos': 'Correos'
            }
            df = df.rename(columns=mapeo)
            df = df[df['Cliente'].notna() & (df['Cliente'].astype(str).str.upper() != 'CLIENTE')]
            
            # Convertir a números
            for col in ['H_M_A', 'H_M_C', 'H_Z_A', 'H_Z_C', 'P_D_A', 'P_D_C']:
                if col in df.columns: df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
            
            self.df_consumo = df
            return df
        except Exception as e:
            logging.error(f"Error procesando datos: {e}"); raise

    def cargar_mapeo_correos(self):
        if self.df_consumo is None: return
        col_correo = next((c for c in self.df_consumo.columns if 'CORREO' in c.upper()), None)
        if not col_correo: return
        
        for _, row in self.df_consumo[self.df_consumo[col_correo].notna()].iterrows():
            cliente = str(row['Cliente']).strip()
            mail_str = str(row[col_correo]).strip()
            mails = [m.strip() for m in re.split('[,;]', mail_str) if '@' in m]
            if mails: self.campañas_correos[cliente] = mails

    def enviar_email_real(self, destinatarios: List[str], cliente: str, datos: pd.Series):
        """ESTA ES LA FUNCIÓN QUE FALTABA"""
        try:
            msg = MIMEMultipart()
            msg['From'] = self.config_email['email']
            msg['To'] = ", ".join(destinatarios)
            msg['Subject'] = f"Reporte de Consumo - {cliente}"

            # Cuerpo simple (puedes meter tu HTML aquí después)
            cuerpo = f"""
            Hola, este es el reporte para {cliente}:
            - Merpes: {datos['H_M_C']} / {datos['H_M_A']} horas.
            - Zoftinium: {datos['H_Z_C']} / {datos['H_Z_A']} horas.
            - Diseño: {datos['P_D_C']} / {datos['P_D_A']} piezas.
            """
            msg.attach(MIMEText(cuerpo, 'plain'))

            # CONEXIÓN AL SERVIDOR
            server = smtplib.SMTP(self.config_email['smtp_server'], self.config_email['puerto'])
            server.starttls()
            server.login(self.config_email['email'], self.config_email['password'])
            server.send_message(msg)
            server.quit()
            return True
        except Exception as e:
            logging.error(f"Fallo envío a {cliente}: {e}")
            return False

    def analizar_y_enviar(self):
        if not self.campañas_correos: return
        for cliente, correos in self.campañas_correos.items():
            fila = self.df_consumo[self.df_consumo['Cliente'] == cliente].iloc[0]
            logging.info(f"Intentando enviar email real a {cliente}...")
            
            # AQUÍ SE LLAMA A LA FUNCIÓN DE ENVÍO
            exito = self.enviar_email_real(correos, cliente, fila)
            if exito:
                logging.info(f"✅ Email enviado correctamente a {cliente}")

    def ejecutar_monitoreo(self):
        archivo = self.descargar_excel()
        self.leer_datos_consumo(archivo)
        self.cargar_mapeo_correos()
        self.analizar_y_enviar()
        if "temp" in archivo: os.remove(archivo)

if __name__ == "__main__":
    from config import EMAIL_CONFIG, EXCEL_URL
    monitor = MonitorCampañas(EXCEL_URL, EMAIL_CONFIG)
    monitor.ejecutar_monitoreo()