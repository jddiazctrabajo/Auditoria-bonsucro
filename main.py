#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Sistema de Trazabilidad de Fertilización
Procesador Batch - Procesa archivos desde carpeta

Uso:
    python main.py
"""

from pathlib import Path
import pandas as pd
from core.procesador import Procesador

# ============================
# CONFIGURACIÓN DE RUTAS
# ============================

# Carpeta raíz del proyecto
CARPETA_PROYECTO = Path(__file__).parent

# Carpetas de datos
CARPETA_ENTRADA = CARPETA_PROYECTO / "data" / "entrada"
CARPETA_SALIDA = CARPETA_PROYECTO / "data" / "salida"

# Archivos de configuración
RUTA_MAESTRO = CARPETA_PROYECTO / "config" / "maestro.xlsx"
RUTA_PRONTUARIO = CARPETA_PROYECTO / "config" / "prontuario.xls"

# ============================
# VALIDACIONES
# ============================

def validar_configuracion():
    """Valida que todos los archivos necesarios existan."""
    errores = []
    
    if not RUTA_MAESTRO.exists():
        errores.append(f"❌ Maestro no encontrado: {RUTA_MAESTRO}")
    
    if not RUTA_PRONTUARIO.exists():
        errores.append(f"❌ Prontuario no encontrado: {RUTA_PRONTUARIO}")
    
    if not CARPETA_ENTRADA.exists():
        errores.append(f"❌ Carpeta entrada no existe: {CARPETA_ENTRADA}")
    
    if errores:
        print("⚠️  Problemas de configuración:")
        for error in errores:
            print(f"   {error}")
        return False
    
    return True


# ============================
# EJECUTAR
# ============================

if __name__ == "__main__":
    
    print("\n" + "="*60)
    print("Sistema de Trazabilidad de Fertilización")
    print("Modo: Procesamiento Batch")
    print("="*60 + "\n")
    
    # Validar configuración
    if not validar_configuracion():
        print("\n⚠️  Por favor, verifica tu configuración.")
        exit(1)
    
    try:
        # Crear procesador
        print(f"📂 Entrada: {CARPETA_ENTRADA}")
        print(f"📂 Salida: {CARPETA_SALIDA}")
        print(f"📋 Maestro: {RUTA_MAESTRO.name}")
        print(f"📋 Prontuario: {RUTA_PRONTUARIO.name}\n")
        
        print("⏳ Procesando archivos...\n")
        
        procesador = Procesador(
            ruta_maestro=str(RUTA_MAESTRO),
            ruta_prontuario=str(RUTA_PRONTUARIO)
        )
        
        # Procesar
        df_consolidado, errores_df = procesador.procesar(
            carpeta=str(CARPETA_ENTRADA)
        )
        
        # Resultados
        print("\n" + "="*60)
        print("✅ PROCESO COMPLETADO EXITOSAMENTE")
        print("="*60)
        print(f"📊 Registros procesados: {len(df_consolidado)}")
        print(f"⚠️  Errores encontrados: {len(errores_df) if isinstance(errores_df, list) else len(errores_df)}")
        print("="*60 + "\n")
        
        # Guardar archivo consolidado
        if not df_consolidado.empty:
            archivo_salida = CARPETA_SALIDA / f"Consolidado_{pd.Timestamp.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
            df_consolidado.to_excel(archivo_salida, index=False)
            print(f"💾 Archivo guardado: {archivo_salida}")
        
        print("\n✅ Proceso terminado correctamente.")
        
    except Exception as e:
        print(f"\n❌ ERROR: {str(e)}")
        print("\nPor favor, verifica:")
        print("  1. La estructura de los archivos de entrada")
        print("  2. Que el maestro y prontuario existan")
        print("  3. Los logs para más información")
        exit(1)