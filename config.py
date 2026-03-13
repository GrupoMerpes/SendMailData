# Configuración del Monitor de Campañas
# Edita este archivo con tus datos específicos

# URL del archivo Excel (reemplaza con tu link real)
EXCEL_URL = "https://merpes2-my.sharepoint.com/:x:/r/personal/jefediseno_grupomerpes_com/_layouts/15/Doc.aspx?sourcedoc=%7B851A7042-13C1-4E10-B670-36E76EAA40DD%7D&file=Trafico%20dise%25u00f1o%20final%20-%202025.xlsx&action=default&mobileredirect=true&isSPOFile=1&ovuser=14d2c19f-8f2c-4895-b4f9-0de0e93f3f0c%2Cdesarrolloweb%40grupomerpes.com&clickparams=eyJBcHBOYW1lIjoiVGVhbXMtRGVza3RvcCIsIkFwcFZlcnNpb24iOiI0OS8yNTA4MTUwMDcxNyIsIkhhc0ZlZGVyYXRlZFVzZXIiOmZhbHNlfQ%3D%3D"

# Configuración del servidor de email
EMAIL_CONFIG = {
    # Para Gmail
    'smtp_server': 'smtp.gmail.com',
    'puerto': 587,
    'email': 'gmerpesdesarrollo2@gmail.com',
    'password': 'uhqd msrp wdkg xdan',  # Usar contraseña de aplicación de Google
    
    # Para Outlook/Hotmail
    # 'smtp_server': 'smtp-mail.outlook.com',
    # 'puerto': 587,
    # 'email': 'tu-email@outlook.com',
    # 'password': 'tu-contraseña',
    
    # Para Yahoo
    # 'smtp_server': 'smtp.mail.yahoo.com',
    # 'puerto': 587,
    # 'email': 'tu-email@yahoo.com',
    # 'password': 'tu-contraseña-de-aplicación',
}

# Configuración de alertas
ALERTAS_CONFIG = {
    # Número de piezas restantes para alerta de advertencia
    'umbral_advertencia': 4,
    
    # Número de piezas restantes para alerta crítica (normalmente 0)
    'umbral_critico': 0,
    
    # Intervalo entre envíos de emails (segundos)
    'pausa_entre_envios': 1,
    
    # Archivo donde se mapean las campañas con los emails
    'archivo_mapeo': 'mapeo_campañas_correos.xlsx'
}

# Configuración de logging
LOGGING_CONFIG = {
    'nivel': 'INFO',  # DEBUG, INFO, WARNING, ERROR
    'archivo_log': 'monitor_campañas.log',
    'formato': '%(asctime)s - %(levelname)s - %(message)s'
}
