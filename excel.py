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

    def filtrar_intervalo(self):
        self.df = self.df[(self.df["fecha"] >= self.fecha_inicio) & (self.df["fecha"] <= self.fecha_fin)]

    def calcular_piezas(self):
        self.df["Piezas_Hechas"] = self.df["M5_C"].diff().fillna(0).astype(int).abs()

    def agrupar_por_intervalos(self):
        self.df["Fecha Inicio"] = self.df["fecha"].apply(
            lambda x: self.fecha_inicio + timedelta(minutes=((x.minute // 5) * 5))
        )
        self.df["Fecha Fin"] = self.df["Fecha Inicio"] + timedelta(minutes=4)
        self.resultado = self.df.groupby(["Fecha Inicio", "Fecha Fin"])["Piezas_Hechas"].sum().reset_index()


