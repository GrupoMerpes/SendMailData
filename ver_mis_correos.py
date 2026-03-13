#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Script para ver EXACTAMENTE qué correos tienes en tu archivo Excel
"""

import pandas as pd
import os

def mostrar_correos_reales():
    archivo = "CONSOLIDADO HORAS SOPORTE - DESARROLLO.xlsx"
    
    print(f"🔍 REVISANDO CORREOS EN: {archivo}")
    print("=" * 80)
    
    if not os.path.exists(archivo):
        print(f"❌ Archivo no encontrado: {archivo}")
        print("\n📁 Archivos disponibles:")
        for f in os.listdir('.'):
            if f.endswith('.xlsx'):
                print(f"  - {f}")
        return
    
    try:
        # Leer archivo
        df = pd.read_excel(archivo, header=0)
        
        print(f"📋 Todas las columnas ({len(df.columns)}):")
        for i, col in enumerate(df.columns):
            print(f"  {i:2d}: {col}")
        
        # Buscar columna de correos
        print(f"\n🔍 BUSCANDO COLUMNA DE CORREOS:")
        correos_col = None
        
        for col in df.columns:
            if 'correo' in str(col).lower() or str(col).lower() == 'unnamed: 14':
                correos_col = col
                print(f"✓ Columna de correos encontrada: '{col}'")
                break
        
        if not correos_col:
            print("❌ No encontré columna de correos, usando la última columna")
            correos_col = df.columns[-1]
        
        print(f"\n📧 CORREOS EXACTOS EN TU ARCHIVO (columna: {correos_col}):")
        print("-" * 80)
        
        # Mostrar todos los correos reales
        contador = 0
        for idx, row in df.iterrows():
            cliente = str(row.iloc[0]) if pd.notna(row.iloc[0]) else f"Fila {idx}"
            correos = row[correos_col]
            
            if pd.notna(correos) and '@' in str(correos):
                contador += 1
                print(f"{contador:2d}. {cliente}: {correos}")
        
        if contador == 0:
            print("❌ No se encontraron correos con @ en ninguna fila")
        
        print(f"\n📊 RESUMEN: {contador} clientes con correos definidos")
        
    except Exception as e:
        print(f"❌ Error: {e}")

if __name__ == "__main__":
    mostrar_correos_reales()