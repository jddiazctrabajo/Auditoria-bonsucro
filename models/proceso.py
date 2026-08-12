from dataclasses import dataclass, field
from datetime import datetime


@dataclass
class Proceso:

    fecha_inicio: datetime = field(default_factory=datetime.now)

    fecha_fin: datetime = None

    usuario: str = ""

    periodo: str = ""

    archivos_procesados: int = 0

    registros: int = 0

    errores: int = 0

    advertencias: int = 0