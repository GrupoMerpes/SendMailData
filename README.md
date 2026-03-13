# 🚀 Sistema de Monitoreo de Campañas

Sistema automatizado que monitorea un archivo Excel de campañas y envía alertas por email cuando las piezas restantes están bajas o agotadas.

## 📋 Características

- ✅ Lee archivos Excel desde un link/URL
- ✅ Procesa la pestaña "Consumo" automáticamente
- ✅ Envía alertas personalizadas por email
- ✅ Diferentes tipos de alerta (crítica y advertencia)
- ✅ Emails HTML con estadísticas detalladas
- ✅ Logging completo de todas las operaciones
- ✅ Mapeo flexible de campañas a correos electrónicos

## 🎯 Tipos de Alertas

### 🚨 Alerta Crítica (0 piezas restantes)
- **Asunto**: "🚨 CRÍTICO: Campaña X - Sin piezas restantes"
- **Mensaje**: No se pueden realizar más cambios
- **Color**: Rojo

### ⚠️ Alerta de Advertencia (≤4 piezas restantes)
- **Asunto**: "⚠️ ADVERTENCIA: Campaña X - Pocas piezas restantes"
- **Mensaje**: Piense bien los cambios que va a realizar
- **Color**: Naranja

## 🛠️ Instalación

### 1. Instalar Python
Asegúrate de tener Python 3.7 o superior instalado.

### 2. Instalar dependencias
```bash
pip install -r requirements.txt
```

### 3. Configurar el sistema

#### A. Editar config.py
```python
# Cambiar la URL del Excel
EXCEL_URL = "https://tu-link-real-del-excel.com/archivo.xlsx"

# Configurar email (ejemplo para Gmail)
EMAIL_CONFIG = {
    'smtp_server': 'smtp.gmail.com',
    'puerto': 587,
    'email': 'tu-email@gmail.com',
    'password': 'tu-contraseña-de-aplicación'
}
```

#### B. Crear mapeo de campañas
```bash
python crear_mapeo.py
```
Esto crea `mapeo_campañas_correos.xlsx` que debes editar con tus datos reales.

## 📧 Configuración de Email

### Para Gmail:
1. Activa la verificación en 2 pasos
2. Genera una "Contraseña de aplicación"
3. Usa esa contraseña en el config.py

### Para Outlook:
```python
EMAIL_CONFIG = {
    'smtp_server': 'smtp-mail.outlook.com',
    'puerto': 587,
    'email': 'tu-email@outlook.com',
    'password': 'tu-contraseña'
}
```

### Para Yahoo:
```python
EMAIL_CONFIG = {
    'smtp_server': 'smtp.mail.yahoo.com',
    'puerto': 587,
    'email': 'tu-email@yahoo.com',
    'password': 'tu-contraseña-de-aplicación'
}
```

## 🚀 Uso

### Ejecución Manual
```bash
python ejecutar_monitor.py
```

### Automatización con Cron (Linux/Mac)
```bash
# Ejecutar cada hora
0 * * * * cd /ruta/al/proyecto && python ejecutar_monitor.py

# Ejecutar cada 30 minutos
*/30 * * * * cd /ruta/al/proyecto && python ejecutar_monitor.py
```

### Automatización con Programador de Tareas (Windows)
1. Abrir "Programador de tareas"
2. Crear tarea básica
3. Configurar trigger (cada hora, diario, etc.)
4. Acción: Iniciar programa
5. Programa: `python.exe`
6. Argumentos: `ejecutar_monitor.py`
7. Directorio: Ruta de tu proyecto

## 📊 Estructura del Excel

El Excel debe tener una pestaña llamada **"Consumo"** con estas columnas:

| CAMPAÑA | TOTAL PIEZAS ASIGNADAS | PIEZAS CONSUMIDAS | PIEZAS RESTANTES |
|---------|------------------------|-------------------|------------------|
| CAMPAÑA_1 | 1000 | 996 | 4 |
| CAMPAÑA_2 | 500 | 500 | 0 |
| CAMPAÑA_3 | 750 | 650 | 100 |

## 📁 Archivos del Proyecto

```
envio-correos/
├── monitor_campañas.py          # Clase principal del monitor
├── ejecutar_monitor.py          # Script principal de ejecución
├── config.py                    # Configuración del sistema
├── crear_mapeo.py              # Script para crear mapeo de correos
├── requirements.txt            # Dependencias de Python
├── README.md                   # Este archivo
├── mapeo_campañas_correos.xlsx # Mapeo campaña -> email (creado automáticamente)
└── monitor_campañas.log        # Log de ejecuciones
```

## 🔧 Personalización

### Cambiar umbrales de alerta
En `config.py`:
```python
ALERTAS_CONFIG = {
    'umbral_advertencia': 10,  # Alerta cuando quedan ≤10 piezas
    'umbral_critico': 0,       # Alerta crítica cuando quedan 0 piezas
}
```

### Modificar diseño del email
Edita la función `generar_contenido_email()` en `monitor_campañas.py`

## 📝 Logs

Todos los eventos se registran en `monitor_campañas.log`:
- Descargas de archivos
- Envíos de emails
- Errores y excepciones
- Estadísticas de ejecución

## 🆘 Solución de Problemas

### Error: "No se puede descargar el Excel"
- Verifica que la URL sea correcta y accesible
- Algunos links de Google Drive/OneDrive requieren permisos especiales

### Error: "Error al enviar email"
- Verifica la configuración SMTP
- Para Gmail, usa contraseña de aplicación, no tu contraseña normal
- Revisa que el puerto y servidor sean correctos

### Error: "Columnas faltantes"
- Verifica que la pestaña se llame exactamente "Consumo"
- Asegúrate de que las columnas tengan los nombres exactos esperados

### Error: "No se encontró email para campaña"
- Actualiza el archivo `mapeo_campañas_correos.xlsx`
- Verifica que los nombres de campaña coincidan exactamente

## 🔄 Automatización Avanzada

### Script Batch para Windows
Crea `ejecutar_monitor.bat`:
```batch
@echo off
cd "C:\ruta\a\tu\proyecto"
python ejecutar_monitor.py
pause
```

### Notificaciones adicionales
Puedes agregar notificaciones por:
- Slack
- Telegram
- WhatsApp Business API
- SMS

## 📈 Mejoras Futuras

- [ ] Interface web para configuración
- [ ] Dashboard con gráficos
- [ ] Integración con bases de datos
- [ ] Notificaciones push móviles
- [ ] Reportes automáticos por PDF

## 🤝 Soporte

Si tienes problemas:
1. Revisa el archivo `monitor_campañas.log`
2. Verifica la configuración en `config.py`
3. Asegúrate de que el Excel tenga el formato correcto

---

**¡Listo para monitorear tus campañas! 🎉**
