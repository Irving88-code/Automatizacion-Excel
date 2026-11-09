"""
Genera archivos Excel de ejemplo para probar el pipeline.
Ejecuta este script primero antes de reporte_consolidado.py
"""
import pandas as pd
import numpy as np
from pathlib import Path

np.random.seed(42)

VENDEDORES = ["Ana López", "Carlos Méndez", "Diana Ruiz", "Eduardo Vega", "Fernanda Cruz"]
PRODUCTOS  = ["Laptop", "Monitor", "Teclado", "Mouse", "Auriculares", "Webcam"]
REGIONES   = ["Norte", "Sur", "Centro", "Este", "Oeste"]

def mes(año, mes_num, n=80):
    fechas = pd.date_range(f"{año}-{mes_num:02d}-01",
                           periods=n, freq="D").normalize()
    fechas = fechas[fechas.month == mes_num][:n]
    if len(fechas) < n:
        fechas = fechas.tolist() + [fechas[-1]] * (n - len(fechas))
        fechas = pd.DatetimeIndex(fechas)

    return pd.DataFrame({
        "Fecha":      np.random.choice(fechas, n),
        "Vendedor":   np.random.choice(VENDEDORES, n),
        "Producto":   np.random.choice(PRODUCTOS, n),
        "Region":     np.random.choice(REGIONES, n),
        "Cantidad":   np.random.randint(1, 20, n),
        "PrecioUni":  np.round(np.random.uniform(50, 2000, n), 2),
        # Introducir suciedad intencional
        "Descuento":  np.where(np.random.random(n) < 0.1, np.nan,
                               np.round(np.random.uniform(0, 0.3, n), 2)),
        "Estado":     np.random.choice(
                          ["Completado", "completado", "COMPLETADO",
                           "Cancelado", "Pendiente", None], n,
                          p=[0.4, 0.1, 0.1, 0.15, 0.2, 0.05]),
    })

carpeta = Path("datos_entrada")
carpeta.mkdir(exist_ok=True)

for i, (año, mes_num) in enumerate([(2024, 10), (2024, 11), (2024, 12)], 1):
    df = mes(año, mes_num)
    path = carpeta / f"ventas_{año}_{mes_num:02d}.xlsx"
    df.to_excel(path, index=False)
    print(f"Creado: {path}  ({len(df)} filas)")

print("\nArchivos de ejemplo listos. Ahora ejecuta: reporte_consolidado.py")
