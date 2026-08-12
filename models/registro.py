from dataclasses import dataclass, field


@dataclass
class Registro:

    id_registro: str

    archivo: str

    hoja: str

    fila_original: int

    datos: dict

    resultados: list = field(default_factory=list)

    procesado: bool = False