#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
🚀 EJECUTOR DE CONSOLA - Monitor de Campañas
============================================
Script simple para ejecutar desde terminal/consola
"""

import os
import sys
from monitor_campañas import MonitorCampañas
from config import EMAIL_CONFIG

class MonitorLocal(MonitorCampañas):
    def __init__(self, archivo_local, config_email):
        self.archivo_local = archivo_local
        self.config_email = config_email
        self.df_consumo = None
        self.campañas_correos = {}
    
    def descargar_excel(self):
        return self.archivo_local

def main():
    print("🚀 Monitor de Campañas - Ejecución desde Consola")
    print("=" * 50)
    
    # Buscar archivos Excel en carpeta datos (como funcionaba antes)
    archivos_excel = []
    
    # Buscar en carpeta datos
    if os.path.exists('datos'):
        for archivo in os.listdir('datos'):
            if archivo.endswith(('.xlsx', '.xls')) and not archivo.startswith('mapeo_'):
                archivos_excel.append(os.path.join('datos', archivo))
    
    # Si no hay en datos, buscar en directorio actual
    if not archivos_excel:
        for archivo in os.listdir('.'):
            if archivo.endswith(('.xlsx', '.xls')) and not archivo.startswith('mapeo_'):
                archivos_excel.append(archivo)
    
    if not archivos_excel:
        print("❌ No se encontraron archivos Excel")
        print("💡 Copia tu archivo Excel a la carpeta 'datos/' o al directorio actual")
        return False
    
    # Usar el primer archivo encontrado
    archivo_seleccionado = archivos_excel[0]
    print(f"📊 Procesando: {os.path.basename(archivo_seleccionado)}")
    
    # VERIFICAR CORREOS EN EL NUEVO ARCHIVO
    print("\n🔍 VALIDANDO CORREOS EN TU NUEVO ARCHIVO...")
    print("-" * 60)
    
    try:
        import pandas as pd
        df = pd.read_excel(archivo_seleccionado, header=0)
        
        print(f"📊 Archivo leído: {len(df)} filas, {len(df.columns)} columnas")
        
        # Mostrar estructura del archivo
        print("\n📋 ESTRUCTURA DEL ARCHIVO:")
        print("Primeras columnas:")
        for i, col in enumerate(df.columns[:5]):
            print(f"  {i}: '{col}'")
        
        print("Últimas columnas:")
        for i, col in enumerate(df.columns[-5:]):
            pos = len(df.columns) - 5 + i
            print(f"  {pos}: '{col}'")
        
        # Buscar TODOS los correos en TODAS las columnas
        print(f"\n🔍 BUSCANDO CORREOS EN TODAS LAS COLUMNAS...")
        correos_por_columna = {}
        
        for col_name in df.columns:
            correos_en_columna = []
            for idx, valor in enumerate(df[col_name]):
                if pd.notna(valor) and '@' in str(valor) and str(valor) != 'nan':
                    cliente = df.iloc[idx, 0] if pd.notna(df.iloc[idx, 0]) else f"Fila {idx}"
                    correos_en_columna.append((cliente, str(valor)))
            
            if correos_en_columna:
                correos_por_columna[col_name] = correos_en_columna
        
        print(f"\n📧 RESULTADO DE LA BÚSQUEDA:")
        print("=" * 60)
        
        if correos_por_columna:
            total_correos = sum(len(correos) for correos in correos_por_columna.values())
            print(f"✅ SE ENCONTRARON {total_correos} CORREOS EN {len(correos_por_columna)} COLUMNAS:")
            
            for col_name, correos_lista in correos_por_columna.items():
                print(f"\n🔹 COLUMNA: '{col_name}' ({len(correos_lista)} correos)")
                for cliente, correo in correos_lista:
                    print(f"   • {cliente}: {correo}")
        else:
            print("❌ NO SE ENCONTRARON CORREOS EN EL ARCHIVO")
            print("💡 El archivo no tiene ninguna columna con correos (@)")
            
        print(f"\n" + "=" * 60)
        print("¿Quieres continuar ejecutando el sistema con estos correos?")
        respuesta = input("Escribe 's' para continuar o cualquier otra tecla para salir: ").lower()
        
        if respuesta != 's':
            print("⏹️ Ejecución cancelada por el usuario")
            return False
    
    except Exception as e:
        print(f"❌ Error al validar el archivo: {e}")
        import traceback
        traceback.print_exc()
        return False
    
    print("\n🚀 EJECUTANDO SISTEMA CON LOS CORREOS ENCONTRADOS...")
    print("=" * 60)
    
    try:
        # Ejecutar monitor
        monitor = MonitorLocal(archivo_seleccionado, EMAIL_CONFIG)
        monitor.ejecutar_monitoreo()
        print("✅ Proceso completado exitosamente")
        return True
        
    except Exception as e:
        print(f"❌ Error: {e}")
        print("📄 Revisa los logs para más detalles")
        return False

if __name__ == "__main__":
    main()
