# Guía Rápida de Inicio 🚀

## Instalación Rápida (Windows)

### Opción 1: Con el script de instalación (Recomendado)

1. Descarga el proyecto
2. Abre una terminal en la carpeta del proyecto
3. Ejecuta:
```batch
install.bat
```

Esto creará un entorno virtual e instalará todas las dependencias automáticamente.

### Opción 2: Manual

1. Abre una terminal en la carpeta del proyecto
2. Crea un entorno virtual:
```batch
python -m venv venv
```

3. Activa el entorno:
```batch
venv\Scripts\activate
```

4. Instala dependencias:
```batch
pip install -r requirements.txt
```

---

## Ejecutar la Aplicación

### Opción A: Con el script de ejecución

Haz doble clic en `run.bat` o desde terminal:
```batch
run.bat
```

### Opción B: Manual

1. Activa el entorno virtual:
```batch
venv\Scripts\activate
```

2. Ejecuta Streamlit:
```batch
streamlit run app.py
```

3. La aplicación se abrirá en tu navegador en `http://localhost:8501`

---

## Uso de la Aplicación

### Modo 1: Subir Archivos

1. Selecciona "Subir archivos" en el panel lateral
2. Carga tus archivos Excel (.xlsx o .xlsm)
3. Los archivos deben tener una hoja llamada "FORMATO FERT"
4. Espera a que se procesen los datos
5. Visualiza los resultados y descarga el consolidado

### Modo 2: Procesar desde Carpeta

1. Copia tus archivos Excel en la carpeta `data/entrada/`
2. Selecciona "Carpeta del proyecto" en el panel lateral
3. Haz clic en "Procesar archivos"
4. Espera a que terminen de procesar
5. Descarga los resultados

### Procesamiento Batch

Para procesar múltiples archivos desde línea de comandos sin UI:

1. Coloca los archivos en `data/entrada/`
2. Ejecuta desde terminal:
```batch
python main.py
```

Los archivos procesados se guardarán en `data/salida/`

---

## Formato de Archivos

Los archivos Excel de entrada deben tener:

- **Nombre de hoja**: `FORMATO FERT`
- **Columnas requeridas**:
  - HACIENDA
  - SUERTE
  - ÁREA
  - PRODUCTO
  - CANTIDAD
  - UNIDAD
  - DOSIS X HA
  - UNIDADES - N
  - UNIDADES - P
  - UNIDADES - K
  - MENORES

---

## Archivos de Configuración

Dos archivos son necesarios en `config/`:

1. **maestro.xlsx**: Base de datos de aporte de fertilizantes
   - Hojas: "APORTE FERTILIZANTES", "APORTE HERBICIDAS"

2. **prontuario.xls**: Base de datos de referencia

---

## Troubleshooting

### "Comando no encontrado"
- Asegúrate de estar en la carpeta correcta
- En Windows, usa `python` no `python3`
- Verifica que Python esté instalado: `python --version`

### "ModuleNotFoundError"
- Activa el entorno virtual: `venv\Scripts\activate`
- Instala nuevamente: `pip install -r requirements.txt`

### "No se encontraron archivos"
- Verifica que los archivos estén en `data/entrada/`
- Asegúrate que sean `.xlsx` o `.xlsm`

### La app no se abre
- Intenta: `streamlit run app.py --logger.level=debug`
- Verifica que el puerto 8501 esté disponible

---

## Notas Importantes

- El sistema genera un archivo consolidado con todos los datos procesados
- Los errores se reportan en un archivo separado
- Los archivos temporales se limpian automáticamente
- Todos los datos se procesan en local (sin conexión a internet)

---

## ¿Necesitas ayuda?

1. Consulta el [README.md](README.md) para información completa
2. Revisa el [CONTRIBUTING.md](CONTRIBUTING.md) para contribuir
3. Abre un Issue en GitHub si encuentras problemas

---

**¡Listo para comenzar!** 🌿
