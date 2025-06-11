import pandas as pd
import openpyxl
from datetime import datetime, timedelta

class ProcesadorPiezas:
    def _init_(self, archivo_csv, fecha_inicio, fecha_fin):
        self.archivo_csv = archivo_csv
        self.fecha_inicio = fecha_inicio
        self.fecha_fin = fecha_fin
        self.df = None
        self.resultado = None

    def cargar_datos(self):
        self.df = pd.read_csv(self.archivo_csv)
        self.df["fecha"] = pd.to_datetime(self.df["fecha"], errors="coerce").dt.tz_localize(None)
        self.df["M5_C"] = pd.to_numeric(self.df["M5_C"], errors="coerce").fillna(0).astype(int)
        self.df.sort_values(by="fecha", ascending=True, inplace=True)