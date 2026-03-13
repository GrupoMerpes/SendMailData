import pandas as pd

def diagnosticar_correos():
    archivo = 'datos/CONSOLIDADO HORAS SOPORTE - DESARROLLO.xlsx'
    
    try:
        print("🔍 DIAGNÓSTICO DE CORREOS")
        print("=" * 50)
        
        # Leer archivo con encabezados en la fila 2 (índice 1)
        df = pd.read_excel(archivo, header=1)
        
        print(f"📊 Total filas: {len(df)}")
        print(f"📊 Total columnas: {len(df.columns)}")
        
        print(f"\n📋 TODAS LAS COLUMNAS:")
        for i, col in enumerate(df.columns):
            print(f"  {i:2d}: '{col}'")
        
        print(f"\n🔍 BUSCANDO COLUMNA DE CORREOS:")
        correo_cols = []
        for col in df.columns:
            if 'correo' in str(col).lower():
                correo_cols.append(col)
                print(f"  ✅ Encontrada: '{col}'")
        
        if not correo_cols:
            print("  ❌ No se encontraron columnas de correos")
            return
        
        # Analizar datos en la columna de correos
        col_correos = correo_cols[0]
        print(f"\n📧 ANALIZANDO COLUMNA: '{col_correos}'")
        
        # Filtrar filas que no son encabezados
        df_filtrado = df[df[df.columns[0]] != 'Cliente'].copy()
        
        print(f"📊 Filas después de filtrar encabezados: {len(df_filtrado)}")
        
        # Buscar correos no vacíos
        correos_no_vacios = df_filtrado[df_filtrado[col_correos].notna()]
        
        print(f"📧 Filas con correos definidos: {len(correos_no_vacios)}")
        
        if len(correos_no_vacios) > 0:
            print(f"\n📋 CLIENTES CON CORREOS:")
            for i, row in correos_no_vacios.iterrows():
                cliente = row[df.columns[0]]
                correo = row[col_correos]
                print(f"  • {cliente}: {correo}")
        else:
            print(f"\n❌ NO HAY CLIENTES CON CORREOS DEFINIDOS")
            print(f"\n🔍 MUESTRA DE LA COLUMNA DE CORREOS:")
            for i in range(min(10, len(df_filtrado))):
                cliente = df_filtrado.iloc[i][df.columns[0]]
                correo = df_filtrado.iloc[i][col_correos]
                print(f"  Fila {i}: {cliente} -> '{correo}' (tipo: {type(correo)})")
        
        return True
        
    except Exception as e:
        print(f"❌ Error: {e}")
        return False

if __name__ == "__main__":
    diagnosticar_correos()