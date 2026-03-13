#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Script para ver exactamente qué correos está leyendo el sistema
"""

import pandas as pd
import os
import sys

def main():
    print("🔍 VERIFICANDO CORREOS EN EL ARCHIVO")
    print("=" * 50)
    
    # Buscar archivos Excel en carpeta datos (igual que ejecutar.py)
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
        return
    
    archivo_seleccionado = archivos_excel[0]
    print(f"📊 Leyendo archivo: {archivo_seleccionado}")
    print("-" * 50)
    
    try:
        # Leer el archivo Excel igual que el sistema
        df = pd.read_excel(archivo_seleccionado, header=0)
        
        print(f"📋 Total de columnas: {len(df.columns)}")
        print(f"📋 Total de filas: {len(df)}")
        print()
        
        # Mapeo de columnas (igual que monitor_campañas.py)
        column_mapping = {
            'SEPTIEMBRE': 'CLIENTE',
            'Unnamed: 1': 'AREA_CAMPAÑA',
            'Horas Soporte Merpes': 'HORAS_MERPES_ASIGNADAS',
            'Unnamed: 3': 'HORAS_MERPES_CONSUMIDAS',
            'Unnamed: 4': 'HORAS_MERPES_DISPONIBLES',
            'Horas Soporte Zoftinium': 'HORAS_ZOFTINIUM_ASIGNADAS', 
            'Unnamed: 7': 'HORAS_ZOFTINIUM_CONSUMIDAS',
            'Unnamed: 8': 'HORAS_ZOFTINIUM_DISPONIBLES',
            'Piezas diseño': 'PIEZAS_DISEÑO_ASIGNADAS',
            'Unnamed: 11': 'PIEZAS_DISEÑO_CONSUMIDAS',
            'Unnamed: 12': 'PIEZAS_DISEÑO_DISPONIBLES',
            'Unnamed: 14': 'CORREOS'
        }
        
        # Aplicar mapeo
        existing_mappings = {old_col: new_col for old_col, new_col in column_mapping.items() if old_col in df.columns}
        df = df.rename(columns=existing_mappings)
        
        print("✅ COLUMNAS MAPEADAS:")
        for old_col, new_col in existing_mappings.items():
            print(f"   {old_col} -> {new_col}")
        print()
        
        # Filtrar encabezados
        df = df[df['CLIENTE'] != 'Cliente']
        df = df[df['CLIENTE'].notna()]
        
        print("📧 CORREOS ENCONTRADOS EN EL ARCHIVO:")
        print("=" * 80)
        
        if 'CORREOS' in df.columns:
            # Buscar filas con correos
            filas_con_correos = df[df['CORREOS'].notna()]
            
            if len(filas_con_correos) > 0:
                print(f"✅ Se encontraron {len(filas_con_correos)} clientes con correos:")
                print()
                
                for index, row in filas_con_correos.iterrows():
                    cliente = str(row['CLIENTE']).strip()
                    correos = str(row['CORREOS']).strip()
                    area = row.get('AREA_CAMPAÑA', 'N/A')
                    
                    if '@' in correos and correos not in ['nan', 'NaN']:
                        print(f"🔹 {cliente}")
                        print(f"   Área: {area}")
                        print(f"   Correos: {correos}")
                        print()
            else:
                print("❌ No se encontraron filas con correos")
        else:
            print("❌ No se encontró la columna CORREOS")
            print("Columnas disponibles:")
            for col in df.columns:
                print(f"   - {col}")
        
    except Exception as e:
        print(f"❌ Error al leer el archivo: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    main()