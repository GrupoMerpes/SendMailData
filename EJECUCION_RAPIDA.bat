@echo off
chcp 65001 >nul
title Monitor de Campañas - Ejecución Rápida

echo.
echo 🚀 MONITOR DE CAMPAÑAS - GRUPO MERPES
echo =====================================
echo.

cd /d "%~dp0"

REM Crear y ejecutar script temporal
(
echo from monitor_campañas import MonitorCampañas
echo from config import EMAIL_CONFIG
echo import os
echo class MonitorLocal^(MonitorCampañas^):
echo     def __init__^(self, archivo_local, config_email^): 
echo         self.archivo_local = archivo_local
echo         self.config_email = config_email
echo         self.df_consumo = None
echo         self.campañas_correos = {}
echo     def descargar_excel^(self^): return self.archivo_local
echo archivos = [f for f in os.listdir^("datos"^) if f.endswith^(".xlsx"^)] if os.path.exists^("datos"^) else []
echo if archivos: MonitorLocal^(f"datos/{archivos[0]}", EMAIL_CONFIG^).ejecutar_monitoreo^(^)
echo else: print^("❌ No se encontró archivo Excel en carpeta datos/"^)
) > temp.py

python temp.py
del temp.py 2>nul

if errorlevel 1 (
    echo.
    echo ❌ Error en la ejecución. Usa EJECUTAR_MONITOR.bat para diagnóstico completo.
) else (
    echo.
    echo ✅ Proceso completado exitosamente
)

echo.
pause
