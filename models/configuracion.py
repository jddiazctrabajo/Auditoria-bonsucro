from dataclasses import dataclass
import pandas as pd


@dataclass
class Configuracion:

    columnas: pd.DataFrame

    productos: pd.DataFrame

    unidades: pd.DataFrame

    operarios: pd.DataFrame

    haciendas: pd.DataFrame

    razones_sociales: pd.DataFrame

    tipos_aplicacion: pd.DataFrame

    parametros: pd.DataFrame