from dataclasses import dataclass
from datetime import datetime


@dataclass
class ErrorSistema:

    id_registro: str

    archivo: str

    hoja: str

    fila: int

    modulo: str

    funcion: str

    tipo: str

    mensaje: str

    fecha: datetime