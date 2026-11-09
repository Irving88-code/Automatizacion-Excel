# Automatización de Reportes Excel con Python

> **Pipeline completo:** une múltiples archivos Excel, limpia los datos y genera un reporte profesional con gráficas — todo automatizado con un solo comando.

---

## ¿Qué hace este proyecto?

Muchas empresas acumulan datos en docenas de archivos Excel separados por mes, sucursal o vendedor. Revisarlos y consolidarlos a mano consume horas y es propenso a errores. Este script resuelve ese problema en segundos:

```
datos_entrada/
  ventas_oct.xlsx
  ventas_nov.xlsx      ──→  reporte_ventas.xlsx  (6 hojas + 5 gráficas)
  ventas_dic.xlsx
```

---

## Funcionalidades

### Carga y unión
- Lee todos los `.xlsx` de una carpeta automáticamente
- Une los archivos en un solo DataFrame con trazabilidad (columna `_archivo`)
- Tolera archivos corruptos sin detener el proceso

### Limpieza de datos
| Problema detectado | Acción aplicada |
|---|---|
| Duplicados exactos | Eliminados |
| Texto inconsistente (`COMPLETADO`, `completado`) | Normalizado a `Title Case` |
| Fechas malformadas | Descartadas con registro |
| Descuentos nulos | Rellenos con `0` |
| Precios o cantidades negativas | Eliminados |

Cada transformación queda registrada en la **hoja de auditoría** del reporte.

### Reporte Excel (7 hojas)
| Hoja | Contenido |
|---|---|
| **Resumen Ejecutivo** | 6 KPIs: venta total, ticket promedio, mejor vendedor, producto líder |
| **Datos Limpios** | Tabla completa con filtros y filas congeladas |
| **Ventas por Mes** | Tabla + gráfica de barras + gráfica de líneas |
| **Por Vendedor** | Ranking + gráfica de barras horizontales |
| **Por Producto** | Tabla + gráfica de pastel |
| **Por Región** | Tabla + gráfica de columnas |
| **Auditoría Limpieza** | Log de cada transformación aplicada |

---

## Tecnologías

- **Python 3.11+**
- **pandas** — carga, limpieza y agregaciones
- **openpyxl** — escritura de Excel, estilos, gráficas nativas
- **numpy** — generación de datos de ejemplo

---

## Instalación

```bash
# 1. Clonar o descargar el repositorio
# 2. Instalar dependencias
pip install -r requirements.txt
```

---

## Uso

### Opción A — datos de ejemplo (demo rápida)
```bash
python generar_datos_ejemplo.py   # crea 3 Excel con datos ficticios
python reporte_consolidado.py     # genera reporte_ventas.xlsx
```

### Opción B — con tus propios archivos
```bash
python reporte_consolidado.py --carpeta ruta/a/tus/datos --salida mi_reporte.xlsx
```

### Opciones disponibles
```
--carpeta   Carpeta con archivos .xlsx de entrada  [defecto: datos_entrada]
--salida    Nombre del archivo de salida           [defecto: reporte_ventas.xlsx]
--titulo    Título del reporte                     [defecto: REPORTE DE VENTAS CONSOLIDADO]
--patron    Patrón glob de archivos                [defecto: *.xlsx]
```

---

## Personalización para tu empresa

El bloque `CONFIG` al inicio del script centraliza todos los parámetros ajustables:

```python
CONFIG = {
    "color_primario":   "1F3864",   # color de encabezados (tu paleta de marca)
    "color_acento1":    "70AD47",
    "col_fecha":        "Fecha",    # nombres de columnas de tu archivo real
    "col_cantidad":     "Cantidad",
    "cols_texto":       ["Vendedor", "Producto", "Region"],
    ...
}
```

No es necesario tocar el resto del código.

---

## Estructura del proyecto

```
📁 Automatizacion de Excel/
├── reporte_consolidado.py      # Script principal
├── generar_datos_ejemplo.py    # Genera datos ficticios para demo
├── requirements.txt
├── README.md
└── 📁 datos_entrada/           # Coloca aquí tus archivos .xlsx
```

---

## Casos de uso

- Consolidación de reportes de ventas por sucursal / mes / vendedor
- Limpieza de exports de CRM o ERP antes de un análisis
- Generación automática de informes periódicos (cron / tarea programada)
- Entrega de reportes formateados a clientes sin que usen Python

---

## Autor

**Irving Torres** — Automatización de procesos con Python  

