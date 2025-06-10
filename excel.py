import pandas as pd
import openpyxl
from datetime import datetime, timedelta

# Cargar el archivo
archivo_csv = "M5_C.08.csv"
df = pd.read_csv(archivo_csv)

# Convertir la columna de fecha a formato datetime y eliminar la zona horaria
df["fecha"] = pd.to_datetime(df["fecha"], errors="coerce").dt.tz_localize(None)

# Asegurar que M5_C es numérico
df["M5_C"] = pd.to_numeric(df["M5_C"], errors="coerce").fillna(0).astype(int)

# Ordenar los datos por fecha en orden ascendente
df.sort_values(by="fecha", ascending=True, inplace=True)

# Definir fecha y hora de inicio y fin
fecha_inicio = datetime(2025, 6, 8, 21, 0)
fecha_fin = datetime(2025, 6, 8, 23, 0)

# Filtrar registros dentro del rango definido
df = df[(df["fecha"] >= fecha_inicio) & (df["fecha"] <= fecha_fin)]

# Calcular la diferencia de piezas hechas
df["Piezas_Hechas"] = df["M5_C"].diff().fillna(0).astype(int).abs()

# Crear intervalos de tiempo personalizados
df["Fecha Inicio"] = df["fecha"].apply(lambda x: fecha_inicio + timedelta(minutes=((x.minute // 5) * 5)))
df["Fecha Fin"] = df["Fecha Inicio"] + timedelta(minutes=4)

# Agrupar por intervalos y sumar piezas
grupo_por_intervalo = df.groupby(["Fecha Inicio", "Fecha Fin"])["Piezas_Hechas"].sum().reset_index()

# Guardar el archivo en Excel con el formato correcto
archivo_modificado = "Resumen_Piezas2.6.xlsx"  # Cambio sutil para justificar el push
wb = openpyxl.Workbook()
ws = wb.active
ws.title = "Piezas por Intervalo"
ws.append(["total", "fecha inicio", "fecha fin"])

# Escribir los datos procesados
for _, row in grupo_por_intervalo.iterrows():
    ws.append([
        row["Piezas_Hechas"],
        row["Fecha Inicio"].strftime("%d/%m/%Y %H:%M:%S"),
        row["Fecha Fin"].strftime("%d/%m/%Y %H:%M:%S")
    ])

wb.save(archivo_modificado)
mostrar_mensaje_final(archivo_modificado)


def mostrar_mensaje_final(nombre_archivo):
    print(f"\n✅ Archivo procesado y guardado como: {nombre_archivo}")
    print(" Procesamiento completado")

