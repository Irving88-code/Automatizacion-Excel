"""
Automatización de Reportes Excel
---------------------------------
Pipeline completo: unir múltiples .xlsx → limpiar datos → reporte con gráficas.

Uso rápido:
    python reporte_consolidado.py
    python reporte_consolidado.py --carpeta mis_datos/ --salida informe.xlsx
    python reporte_consolidado.py --titulo "Ventas Q4 2024"
"""
import argparse
import sys
from pathlib import Path
from datetime import datetime

import pandas as pd
import numpy as np
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.chart import BarChart, LineChart, PieChart, Reference
from openpyxl.utils import get_column_letter
from openpyxl.utils.dataframe import dataframe_to_rows

# ─────────────────────────────────────────────────────────────
# CONFIGURACIÓN — ajusta estos valores para cada cliente
# ─────────────────────────────────────────────────────────────

CONFIG = {
    # Paleta de colores (hex sin #)
    "color_primario":   "1F3864",   # azul oscuro — encabezados principales
    "color_secundario": "2E75B6",   # azul medio  — encabezados de tabla
    "color_acento1":    "70AD47",   # verde
    "color_acento2":    "ED7D31",   # naranja
    "color_fila_par":   "F2F2F2",   # gris claro para filas alternas

    # Columnas esperadas en los archivos de entrada
    "col_fecha":      "Fecha",
    "col_cantidad":   "Cantidad",
    "col_precio":     "PrecioUni",
    "col_descuento":  "Descuento",
    "col_estado":     "Estado",
    "cols_texto":     ["Vendedor", "Producto", "Region"],

    # Nombre del título del reporte (se puede sobreescribir con --titulo)
    "titulo_reporte": "REPORTE DE VENTAS CONSOLIDADO",
}

# ─────────────────────────────────────────────────────────────
# 1. CARGA Y UNIÓN
# ─────────────────────────────────────────────────────────────

def cargar_excels(carpeta: Path, patron: str = "*.xlsx") -> pd.DataFrame:
    """Lee todos los archivos que coincidan con 'patron' y los une en un solo DataFrame."""
    archivos = sorted(carpeta.glob(patron))
    if not archivos:
        sys.exit(f"[ERROR] No se encontraron archivos '{patron}' en '{carpeta}'")

    partes = []
    for f in archivos:
        try:
            df = pd.read_excel(f)
            df["_archivo"] = f.name
            partes.append(df)
            print(f"  ✔  {f.name}  →  {len(df):,} filas")
        except Exception as exc:
            print(f"  ✗  {f.name}: {exc}")

    if not partes:
        sys.exit("[ERROR] Ningún archivo pudo cargarse correctamente.")

    combinado = pd.concat(partes, ignore_index=True)
    print(f"\n  Total: {len(combinado):,} filas de {len(partes)} archivos\n")
    return combinado


# ─────────────────────────────────────────────────────────────
# 2. LIMPIEZA Y AUDITORÍA
# ─────────────────────────────────────────────────────────────

def limpiar(df: pd.DataFrame) -> tuple[pd.DataFrame, dict, list[dict]]:
    """
    Limpia el DataFrame y devuelve:
      - df limpio
      - stats: resumen numérico de la limpieza
      - log:   registro detallado de cada acción (para la hoja de auditoría)
    """
    stats: dict = {}
    log: list[dict] = []
    original = len(df)

    def _registrar(accion, afectadas, detalle=""):
        log.append({"Acción": accion, "Filas afectadas": afectadas, "Detalle": detalle})

    # — Duplicados exactos —
    antes = len(df)
    df = df.drop_duplicates()
    eliminados = antes - len(df)
    stats["duplicados_eliminados"] = eliminados
    _registrar("Eliminar duplicados exactos", eliminados,
               "drop_duplicates() sobre todas las columnas")

    # — Normalizar texto —
    col_estado = CONFIG["col_estado"]
    if col_estado in df.columns:
        nulos_e = df[col_estado].isna().sum()
        df[col_estado] = df[col_estado].str.strip().str.title().fillna("Desconocido")
        _registrar("Normalizar Estado", int(nulos_e),
                   "strip + title case; nulos → 'Desconocido'")

    for col in CONFIG["cols_texto"]:
        if col in df.columns:
            df[col] = df[col].str.strip().str.title()
    _registrar("Normalizar columnas de texto", 0,
               f"strip + title case en: {CONFIG['cols_texto']}")

    # — Fechas —
    col_fecha = CONFIG["col_fecha"]
    if col_fecha in df.columns:
        df[col_fecha] = pd.to_datetime(df[col_fecha], errors="coerce")
        nulos_f = int(df[col_fecha].isna().sum())
        df = df.dropna(subset=[col_fecha])
        stats["filas_sin_fecha"] = nulos_f
        _registrar("Eliminar filas sin fecha válida", nulos_f,
                   "to_datetime(errors='coerce') + dropna")

    # — Descuento —
    col_desc = CONFIG["col_descuento"]
    if col_desc in df.columns:
        nulos_d = int(df[col_desc].isna().sum())
        df[col_desc] = df[col_desc].fillna(0).clip(0, 1)
        stats["descuentos_nulos_rellenos"] = nulos_d
        _registrar("Rellenar descuentos nulos", nulos_d,
                   "fillna(0) + clip(0, 1)")

    # — Valores numéricos negativos o nulos —
    for col in [CONFIG["col_cantidad"], CONFIG["col_precio"]]:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce")
            invalidos = int((df[col].fillna(0) <= 0).sum())
            df = df[df[col].fillna(0) > 0]
            _registrar(f"Eliminar filas con {col} ≤ 0", invalidos,
                       "to_numeric(errors='coerce') + filtro > 0")

    # — Columna calculada —
    cols_calc = {CONFIG["col_cantidad"], CONFIG["col_precio"], CONFIG["col_descuento"]}
    if cols_calc.issubset(df.columns):
        df["Venta"] = (
            df[CONFIG["col_cantidad"]]
            * df[CONFIG["col_precio"]]
            * (1 - df[CONFIG["col_descuento"]])
        ).round(2)
        _registrar("Calcular columna Venta", len(df),
                   "Cantidad × PrecioUni × (1 − Descuento)")

    # — Columnas de tiempo —
    if col_fecha in df.columns:
        df["Mes"]    = df[col_fecha].dt.to_period("M").astype(str)
        df["Semana"] = df[col_fecha].dt.isocalendar().week.astype(int)
        _registrar("Agregar columnas Mes y Semana", len(df),
                   "dt.to_period('M') / dt.isocalendar().week")

    stats["filas_originales"] = original
    stats["filas_finales"]    = len(df)
    return df, stats, log


# ─────────────────────────────────────────────────────────────
# 3. AGREGACIONES
# ─────────────────────────────────────────────────────────────

def agregar(df: pd.DataFrame) -> dict:
    """Genera tablas resumen por diferentes dimensiones."""
    agg: dict = {}
    if "Venta" not in df.columns:
        return agg

    agg["por_mes"] = (
        df.groupby("Mes")["Venta"]
        .agg(Total="sum", Transacciones="count", Promedio="mean")
        .reset_index().sort_values("Mes")
    )

    agg["por_vendedor"] = (
        df.groupby("Vendedor")["Venta"]
        .agg(Total="sum", Transacciones="count")
        .reset_index().sort_values("Total", ascending=False)
    )

    agg["por_producto"] = (
        df.groupby("Producto")["Venta"]
        .agg(Total="sum", Unidades="count")
        .reset_index().sort_values("Total", ascending=False)
    )

    if "Region" in df.columns:
        agg["por_region"] = (
            df.groupby("Region")["Venta"]
            .sum().reset_index()
            .rename(columns={"Venta": "Total"})
            .sort_values("Total", ascending=False)
        )

    if CONFIG["col_estado"] in df.columns:
        agg["por_estado"] = (
            df[CONFIG["col_estado"]].value_counts()
            .reset_index()
            .rename(columns={"index": "Estado", CONFIG["col_estado"]: "Cantidad",
                              "count": "Cantidad"})
        )

    return agg


# ─────────────────────────────────────────────────────────────
# 4. HELPERS DE ESTILO
# ─────────────────────────────────────────────────────────────

def _borde():
    lado = Side(style="thin", color="BFBFBF")
    return Border(left=lado, right=lado, top=lado, bottom=lado)


def _fill(hex_color: str) -> PatternFill:
    return PatternFill("solid", fgColor=hex_color)


def aplicar_encabezado(ws, fila: int, n_cols: int, color: str = None):
    color = color or CONFIG["color_primario"]
    for c in range(1, n_cols + 1):
        cell = ws.cell(row=fila, column=c)
        cell.fill = _fill(color)
        cell.font = Font(color="FFFFFF", bold=True, size=11)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = _borde()


def escribir_tabla(ws, df: pd.DataFrame, fila_inicio: int,
                   col_inicio: int = 1, cols_dinero: set = None):
    """Escribe df en ws a partir de fila_inicio con estilo de tabla."""
    cols_dinero = cols_dinero or set()

    for ci, col in enumerate(df.columns, col_inicio):
        ws.cell(row=fila_inicio, column=ci, value=col)
    aplicar_encabezado(ws, fila_inicio, len(df.columns) + col_inicio - 1,
                       CONFIG["color_secundario"])

    for ri, row in enumerate(df.itertuples(index=False), fila_inicio + 1):
        bg = CONFIG["color_fila_par"] if ri % 2 == 0 else "FFFFFF"
        for ci, val in enumerate(row, col_inicio):
            cell = ws.cell(row=ri, column=ci, value=val)
            cell.fill = _fill(bg)
            cell.border = _borde()
            cell.alignment = Alignment(vertical="center")
            nombre_col = df.columns[ci - col_inicio]
            if nombre_col in cols_dinero or any(
                p in nombre_col.lower() for p in ("total", "venta", "promedio")
            ):
                cell.number_format = '"$"#,##0.00'


def autoajustar(ws):
    for col in ws.columns:
        max_len = max((len(str(c.value)) for c in col if c.value), default=10)
        ws.column_dimensions[get_column_letter(col[0].column)].width = min(max_len + 4, 40)


# ─────────────────────────────────────────────────────────────
# 5. HOJAS DEL REPORTE
# ─────────────────────────────────────────────────────────────

def hoja_resumen(wb: Workbook, df: pd.DataFrame, stats: dict,
                 agg: dict, titulo: str):
    ws = wb.create_sheet("Resumen Ejecutivo")

    # Banda de título
    ws.merge_cells("A1:F1")
    ws["A1"] = titulo
    ws["A1"].font = Font(bold=True, size=16, color="FFFFFF")
    ws["A1"].fill = _fill(CONFIG["color_primario"])
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 38

    ws["A2"] = f"Generado: {datetime.now().strftime('%d/%m/%Y  %H:%M')}"
    ws["A2"].font = Font(italic=True, color="808080", size=9)

    # Subtítulo KPIs
    ws.merge_cells("A4:F4")
    ws["A4"] = "INDICADORES CLAVE DE DESEMPEÑO"
    ws["A4"].font = Font(bold=True, size=12, color=CONFIG["color_primario"])
    ws["A4"].alignment = Alignment(horizontal="left")

    kpis = [
        ("Registros procesados",   f"{stats.get('filas_finales', 0):,}"),
        ("Duplicados eliminados",  f"{stats.get('duplicados_eliminados', 0):,}"),
        ("Venta total",
         f"${df['Venta'].sum():,.2f}" if "Venta" in df.columns else "—"),
        ("Ticket promedio",
         f"${df['Venta'].mean():,.2f}" if "Venta" in df.columns else "—"),
        ("Mejor vendedor",
         agg["por_vendedor"].iloc[0]["Vendedor"] if "por_vendedor" in agg else "—"),
        ("Producto líder",
         agg["por_producto"].iloc[0]["Producto"] if "por_producto" in agg else "—"),
    ]
    colores_kpi = [
        CONFIG["color_secundario"], CONFIG["color_acento2"],
        CONFIG["color_acento1"],    CONFIG["color_secundario"],
        CONFIG["color_primario"],   CONFIG["color_acento1"],
    ]

    for i, ((label, val), color) in enumerate(zip(kpis, colores_kpi)):
        col_ws = (i % 3) * 2 + 1
        row_ws = 5 + (i // 3) * 4

        ws.merge_cells(start_row=row_ws, start_column=col_ws,
                       end_row=row_ws, end_column=col_ws + 1)
        ws.merge_cells(start_row=row_ws + 1, start_column=col_ws,
                       end_row=row_ws + 2, end_column=col_ws + 1)

        lbl = ws.cell(row=row_ws, column=col_ws, value=label)
        lbl.fill = _fill(color)
        lbl.font = Font(bold=True, color="FFFFFF", size=9)
        lbl.alignment = Alignment(horizontal="center", vertical="center")
        lbl.border = _borde()

        val_c = ws.cell(row=row_ws + 1, column=col_ws, value=val)
        val_c.font = Font(bold=True, size=14, color=CONFIG["color_primario"])
        val_c.alignment = Alignment(horizontal="center", vertical="center")
        val_c.border = _borde()
        ws.row_dimensions[row_ws + 1].height = 28

    autoajustar(ws)


def hoja_datos(wb: Workbook, df: pd.DataFrame):
    ws = wb.create_sheet("Datos Limpios")
    df_out = df.drop(columns=["_archivo"], errors="ignore")

    for r in dataframe_to_rows(df_out, index=False, header=True):
        ws.append(r)

    aplicar_encabezado(ws, 1, len(df_out.columns))
    for ri in range(2, len(df_out) + 2):
        bg = CONFIG["color_fila_par"] if ri % 2 == 0 else "FFFFFF"
        for ci in range(1, len(df_out.columns) + 1):
            cell = ws.cell(ri, ci)
            cell.fill = _fill(bg)
            cell.border = _borde()
            if df_out.columns[ci - 1] == "Venta":
                cell.number_format = '"$"#,##0.00'

    ws.auto_filter.ref = ws.dimensions
    ws.freeze_panes = "A2"
    autoajustar(ws)


def hoja_ventas_mes(wb: Workbook, agg: dict):
    if "por_mes" not in agg:
        return
    ws = wb.create_sheet("Ventas por Mes")
    df = agg["por_mes"]

    ws["A1"] = "Ventas Mensuales"
    ws["A1"].font = Font(bold=True, size=14, color=CONFIG["color_primario"])

    escribir_tabla(ws, df, fila_inicio=3, cols_dinero={"Total", "Promedio"})
    autoajustar(ws)
    n = len(df)

    # Barras — ventas totales
    bc = BarChart()
    bc.type, bc.title = "col", "Ventas Totales por Mes"
    bc.y_axis.title, bc.x_axis.title = "Total ($)", "Mes"
    bc.style, bc.width, bc.height = 10, 20, 14
    datos = Reference(ws, min_col=2, min_row=3, max_row=3 + n)
    cats  = Reference(ws, min_col=1, min_row=4, max_row=3 + n)
    bc.add_data(datos, titles_from_data=True)
    bc.set_categories(cats)
    bc.series[0].graphicalProperties.solidFill = CONFIG["color_secundario"]
    ws.add_chart(bc, f"A{n + 7}")

    # Líneas — transacciones
    lc = LineChart()
    lc.title = "Número de Transacciones por Mes"
    lc.y_axis.title, lc.x_axis.title = "Transacciones", "Mes"
    lc.style, lc.width, lc.height = 10, 20, 14
    datos2 = Reference(ws, min_col=3, min_row=3, max_row=3 + n)
    lc.add_data(datos2, titles_from_data=True)
    lc.set_categories(cats)
    lc.series[0].graphicalProperties.line.solidFill = CONFIG["color_acento1"]
    ws.add_chart(lc, f"L{n + 7}")


def hoja_vendedores(wb: Workbook, agg: dict):
    if "por_vendedor" not in agg:
        return
    ws = wb.create_sheet("Por Vendedor")
    df = agg["por_vendedor"]

    ws["A1"] = "Ranking de Vendedores"
    ws["A1"].font = Font(bold=True, size=14, color=CONFIG["color_primario"])

    escribir_tabla(ws, df, fila_inicio=3, cols_dinero={"Total"})
    autoajustar(ws)
    n = len(df)

    bc = BarChart()
    bc.type, bc.title = "bar", "Ventas por Vendedor"
    bc.y_axis.title, bc.x_axis.title = "Vendedor", "Total ($)"
    bc.style, bc.width, bc.height = 10, 22, 14
    datos = Reference(ws, min_col=2, min_row=3, max_row=3 + n)
    cats  = Reference(ws, min_col=1, min_row=4, max_row=3 + n)
    bc.add_data(datos, titles_from_data=True)
    bc.set_categories(cats)
    bc.series[0].graphicalProperties.solidFill = CONFIG["color_acento2"]
    ws.add_chart(bc, "E3")


def hoja_productos(wb: Workbook, agg: dict):
    if "por_producto" not in agg:
        return
    ws = wb.create_sheet("Por Producto")
    df = agg["por_producto"]

    ws["A1"] = "Ventas por Producto"
    ws["A1"].font = Font(bold=True, size=14, color=CONFIG["color_primario"])

    escribir_tabla(ws, df, fila_inicio=3, cols_dinero={"Total"})
    autoajustar(ws)
    n = len(df)

    pie = PieChart()
    pie.title  = "Distribución de Ventas por Producto"
    pie.style, pie.width, pie.height = 10, 20, 16
    datos = Reference(ws, min_col=2, min_row=3, max_row=3 + n)
    cats  = Reference(ws, min_col=1, min_row=4, max_row=3 + n)
    pie.add_data(datos, titles_from_data=True)
    pie.set_categories(cats)
    ws.add_chart(pie, "E3")


def hoja_region(wb: Workbook, agg: dict):
    if "por_region" not in agg:
        return
    ws = wb.create_sheet("Por Región")
    df = agg["por_region"]

    ws["A1"] = "Ventas por Región"
    ws["A1"].font = Font(bold=True, size=14, color=CONFIG["color_primario"])

    escribir_tabla(ws, df, fila_inicio=3, cols_dinero={"Total"})
    autoajustar(ws)
    n = len(df)

    bc = BarChart()
    bc.type, bc.title = "col", "Ventas por Región"
    bc.y_axis.title, bc.x_axis.title = "Total ($)", "Región"
    bc.style, bc.width, bc.height = 10, 18, 14
    datos = Reference(ws, min_col=2, min_row=3, max_row=3 + n)
    cats  = Reference(ws, min_col=1, min_row=4, max_row=3 + n)
    bc.add_data(datos, titles_from_data=True)
    bc.set_categories(cats)
    bc.series[0].graphicalProperties.solidFill = CONFIG["color_acento1"]
    ws.add_chart(bc, "D3")


def hoja_auditoria(wb: Workbook, stats: dict, log: list[dict]):
    """Hoja de auditoría: documenta cada transformación aplicada a los datos."""
    ws = wb.create_sheet("Auditoría Limpieza")

    ws.merge_cells("A1:D1")
    ws["A1"] = "REGISTRO DE LIMPIEZA DE DATOS"
    ws["A1"].font = Font(bold=True, size=14, color="FFFFFF")
    ws["A1"].fill = _fill(CONFIG["color_primario"])
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 30

    # Resumen general
    ws["A3"] = "Resumen general"
    ws["A3"].font = Font(bold=True, color=CONFIG["color_primario"])
    resumen = [
        ("Filas originales",   stats.get("filas_originales", 0)),
        ("Filas después de limpieza", stats.get("filas_finales", 0)),
        ("Duplicados eliminados", stats.get("duplicados_eliminados", 0)),
        ("Filas sin fecha descartadas", stats.get("filas_sin_fecha", 0)),
        ("Descuentos nulos rellenos", stats.get("descuentos_nulos_rellenos", 0)),
    ]
    for i, (k, v) in enumerate(resumen, 4):
        ws.cell(i, 1, k).font = Font(bold=True)
        ws.cell(i, 2, v)

    # Log detallado
    if log:
        log_df = pd.DataFrame(log)
        fila_log = len(resumen) + 6
        ws.cell(fila_log, 1, "Detalle de transformaciones").font = Font(
            bold=True, color=CONFIG["color_primario"])
        escribir_tabla(ws, log_df, fila_inicio=fila_log + 1)

    autoajustar(ws)


# ─────────────────────────────────────────────────────────────
# 6. MAIN
# ─────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(
        description="Genera un reporte Excel consolidado a partir de múltiples archivos.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Ejemplos:
  python reporte_consolidado.py
  python reporte_consolidado.py --carpeta datos/ --salida output/reporte_q4.xlsx
  python reporte_consolidado.py --titulo "Reporte Trimestral Q1 2025"
        """,
    )
    parser.add_argument("--carpeta", default="datos_entrada",
                        help="Carpeta con los .xlsx de entrada  [defecto: datos_entrada]")
    parser.add_argument("--salida",  default="reporte_ventas.xlsx",
                        help="Archivo Excel de salida           [defecto: reporte_ventas.xlsx]")
    parser.add_argument("--titulo",  default=CONFIG["titulo_reporte"],
                        help="Título del reporte")
    parser.add_argument("--patron",  default="*.xlsx",
                        help="Patrón glob de archivos a incluir [defecto: *.xlsx]")
    args = parser.parse_args()

    carpeta = Path(args.carpeta)
    salida  = Path(args.salida)
    salida.parent.mkdir(parents=True, exist_ok=True)

    sep = "─" * 55
    print(f"\n{sep}")
    print("   PIPELINE  Excel → Limpieza → Reporte")
    print(f"{sep}\n")

    print("[1/4] Cargando archivos...")
    df_raw = cargar_excels(carpeta, args.patron)

    print("[2/4] Limpiando datos...")
    df, stats, log = limpiar(df_raw)
    print(f"  Filas originales : {stats['filas_originales']:,}")
    print(f"  Filas finales    : {stats['filas_finales']:,}")
    print(f"  Duplicados       : {stats['duplicados_eliminados']:,}")

    print("\n[3/4] Calculando métricas...")
    agg = agregar(df)
    for nombre in agg:
        print(f"  ✔  {nombre}")

    print("\n[4/4] Generando reporte Excel...")
    wb = Workbook()
    wb.remove(wb.active)

    hoja_resumen(wb, df, stats, agg, args.titulo)
    hoja_datos(wb, df)
    hoja_ventas_mes(wb, agg)
    hoja_vendedores(wb, agg)
    hoja_productos(wb, agg)
    hoja_region(wb, agg)
    hoja_auditoria(wb, stats, log)

    wb.save(salida)
    print(f"\n{sep}")
    print(f"  Reporte listo: {salida.resolve()}")
    print(f"{sep}\n")


if __name__ == "__main__":
    main()
