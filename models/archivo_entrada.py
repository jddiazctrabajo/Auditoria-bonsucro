from dataclasses import dataclass
import pandas as pd


@dataclass
class ArchivoEntrada:

    razon_social: str

    hacienda: str

    ruta: str

    nombre_archivo: str

    dataframe: pd.DataFrame