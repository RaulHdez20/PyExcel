# raul_procesamiento.py
from datetime import datetime
from procesador import ProcesadorPiezas

archivo = "M5_C.08.csv"
fecha_inicio = datetime(2025, 6, 8, 21, 0)
fecha_fin = datetime(2025, 6, 8, 23, 0)

procesador = ProcesadorPiezas(archivo, fecha_inicio, fecha_fin)
procesador.cargar_datos()
procesador.filtrar_intervalo()