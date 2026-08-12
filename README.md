# Sistema de Trazabilidad de Fertilización 🌿

Un sistema completo para procesar, validar y consolidar datos de fertilización agrícola con interfaz Streamlit.

## 🎯 Características

- **Carga de datos**: Importa archivos Excel con información de fertilización
- **Homologación**: Normaliza nombres de campos y valores
- **Complementación**: Enriquece datos con información del prontuario y maestro
- **Cálculos**: Calcula nutrientes (N, P, K, S) y unidades
- **Validación**: Valida datos y reporta errores
- **Consolidación**: Genera reportes consolidados con trazabilidad

## 📋 Requisitos

- Python 3.8+
- pip o conda

## 🚀 Instalación

1. Clona el repositorio:
```bash
git clone https://github.com/tu-usuario/Proyecto_Fertilizacion.git
cd Proyecto_Fertilizacion
```

2. Crea un entorno virtual (recomendado):
```bash
python -m venv venv

# En Windows
venv\Scripts\activate
# En macOS/Linux
source venv/bin/activate
```

3. Instala las dependencias:
```bash
pip install -r requirements.txt
```

## 📁 Estructura del Proyecto

```
Proyecto_Fertilizacion/
├── app.py                      # Aplicación Streamlit principal
├── main.py                     # Script para procesamiento batch
├── requirements.txt            # Dependencias del proyecto
├── README.md                   # Este archivo
├── .gitignore                  # Archivos a ignorar en git
│
├── config/                     # Archivos de configuración
│   ├── maestro.xlsx           # Base de datos maestro (fertilizantes)
│   └── prontuario.xls         # Base de datos prontuario (herbicidas)
│
├── core/                       # Lógica principal
│   ├── procesador.py          # Orquestador principal
│   ├── homologacion.py        # Normalización de datos
│   ├── maestro.py             # Consultas al maestro
│   ├── prontuario.py          # Consultas al prontuario
│   ├── calculos.py            # Cálculos de nutrientes
│   ├── validaciones.py        # Validaciones de datos
│   ├── normalizador_campos.py # Normalización de campos (en desarrollo)
│   └── reportes.py            # Generación de reportes (en desarrollo)
│
├── services/                   # Servicios utilitarios
│   ├── excel_reader.py        # Lectura de archivos Excel
│   ├── file_manager.py        # Gestión de archivos
│   ├── logger.py              # Sistema de logging
│   └── config_reader.py       # Lectura de configuración
│
├── models/                     # Modelos de datos (dataclasses)
│   ├── archivo_entrada.py     # Modelo de archivo de entrada
│   ├── configuracion.py       # Modelo de configuración
│   ├── error.py               # Modelo de errores
│   ├── proceso.py             # Modelo de proceso
│   └── registro.py            # Modelo de registro
│
├── data/                       # Datos de la aplicación
│   ├── entrada/               # Archivos de entrada
│   ├── salida/                # Archivos procesados
│   └── temporal/              # Archivos temporales
│
├── logs/                       # Archivos de log
├── assets/                     # Recursos estáticos
└── ui/                         # Componentes UI (en desarrollo)
```

## 🎮 Uso

### Opción 1: Aplicación Streamlit (Interactiva)

```bash
streamlit run app.py
```

Luego abre tu navegador en `http://localhost:8501`

1. Carga los archivos Excel de fertilización
2. Verifica que el maestro y prontuario estén disponibles
3. Revisa los datos homologados y validados
4. Descarga el reporte consolidado

### Opción 2: Procesamiento Batch (Lote)

Para procesar varios archivos desde línea de comandos:

```bash
python main.py
```

**Nota**: Configura las rutas de entrada y salida en `main.py` según tu entorno.

## 📊 Formatos Esperados

### Archivo de Entrada
Debe contener una hoja llamada `FORMATO FERT` con las siguientes columnas:

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

### Maestro (config/maestro.xlsx)
- Hojas: `APORTE FERTILIZANTES`, `APORTE HERBICIDAS`
- Contiene información de nutrientes de cada producto

### Prontuario (config/prontuario.xls)
- Base de datos de referencia para validación
- Normalización de nombres de productos

## ⚙️ Configuración

La mayoría de configuraciones se encuentran en:
- `services/config_reader.py` - Lectura de archivos de configuración
- `core/procesador.py` - Parámetros del procesamiento

Las rutas de datos se pueden configurar en:
- Variables de entorno (`.env`)
- Parámetros en `main.py` para batch
- Interfaz de Streamlit para la versión interactiva

## 🧪 Pruebas

Para ejecutar pruebas (si las hay):

```bash
pytest tests/
```

**Nota**: El directorio de pruebas está en desarrollo.

## 📝 Flujo de Procesamiento

1. **Lectura**: Lee archivos Excel de la carpeta de entrada
2. **Homologación**: Normaliza nombres de campos y productos
3. **Prontuario**: Completa datos desde base de datos de prontuario
4. **Maestro**: Completa información nutricional desde maestro
5. **Cálculos**: Calcula unidades de nutrientes (N, P, K, S)
6. **Validación**: Valida datos y recopila errores
7. **Consolidación**: Genera DataFrame final con todos los datos

## 🐛 Troubleshooting

### "No se encontraron archivos Excel"
- Verifica que los archivos estén en la carpeta de entrada
- Asegúrate que sean `.xlsx` o `.xlsm`

### "Error al leer maestro o prontuario"
- Verifica que existan en `config/maestro.xlsx` y `config/prontuario.xls`
- Comprueba que las hojas tengan los nombres correctos

### Errores de validación
- Revisa el reporte de errores en la app de Streamlit
- Verifica que los datos cumplan con el formato esperado

## 📧 Contacto

**Autor**: jddiazc

## 📄 Licencia

Este proyecto está bajo licencia MIT. Ver archivo LICENSE para más detalles.

## 🔄 Historial de Versiones

- **v1.0.0** - Versión inicial con procesamiento básico
  - Lectura de archivos
  - Homologación
  - Cálculos
  - Validación
  - Interfaz Streamlit

## 🤝 Contribuciones

Las contribuciones son bienvenidas. Por favor:
1. Fork el repositorio
2. Crea una rama para tu feature (`git checkout -b feature/AmazingFeature`)
3. Commit tus cambios (`git commit -m 'Add some AmazingFeature'`)
4. Push a la rama (`git push origin feature/AmazingFeature`)
5. Abre un Pull Request

## 📚 Documentación Adicional

- Ver `docs/` para documentación detallada
- Ver comentarios en el código para funciones específicas

---

**Última actualización**: Agosto 2024
