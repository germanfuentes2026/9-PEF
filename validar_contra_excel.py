"""
Compara los resultados del motor Python con los valores guardados por Excel.

Uso:
    python tools/validar_contra_excel.py data/PEF_X.xlsx [otro.xlsx ...]

El archivo tiene que haberse guardado con Excel (para que tenga los valores calculados).
"""
import os
import sys

import numpy as np
import openpyxl

sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))
from pef_model import cargar_inputs_excel, calcular  # noqa: E402

FILAS = [  # (nombre, hoja, fila, primera columna, fila en resultados)
    ("Peajes", "FLUJO", 73, 3, ("flujo", "Peajes")),
    ("Total ingresos", "FLUJO", 22, 3, ("flujo", "TOTAL INGRESOS")),
    ("Servicio de la deuda", "FLUJO", 29, 3, ("flujo", "Servicio de la deuda")),
    ("Impuestos", "FLUJO", 30, 3, ("flujo", "Impuestos")),
    ("Total CAPEX", "FLUJO", 39, 3, ("flujo", "TOTAL CAPEX")),
    ("Flujo neto", "FLUJO", 45, 3, ("flujo", "FLUJO NETO")),
    ("IVA a pagar", "IMPUESTOS", 41, 4, ("impuestos", "IVA a pagar")),
    ("Amortización", "IMPUESTOS", 48, 4, ("impuestos", "Amortización")),
    ("Ganancias", "IMPUESTOS", 66, 4, ("impuestos", "Impuesto a las Ganancias")),
]


def validar(ruta):
    res = calcular(cargar_inputs_excel(ruta))
    wb = openpyxl.load_workbook(ruta, data_only=True)
    print(f"\n{os.path.basename(ruta)}")
    for nombre, hoja, fila, c0, (tabla, etiqueta) in FILAS:
        ws = wb[hoja]
        excel = np.array([v if isinstance(v := ws.cell(fila, c0 + i).value, (int, float)) else 0 for i in range(20)], float)
        py = getattr(res, tabla).loc[etiqueta].values
        dif = np.max(np.abs(excel - py))
        print(f"  {nombre:22s} diferencia máxima: $ {dif:,.2f}  {'OK' if dif < 1 else 'REVISAR'}")
    C = wb["CONTROL"]
    k = res.kpis
    for lab, celda, clave in [("TIR", "C2", "TIR (con IVA, incluye préstamo)"), ("TIR modificada", "C3", "TIR modificada"),
                              ("VAN", "C4", "VAN"), ("Payback", "C6", "Payback (años)")]:
        print(f"  {lab:22s} Excel: {C[celda].value}  ·  Python: {k[clave]}")


if __name__ == "__main__":
    for r in sys.argv[1:] or [os.path.join(os.path.dirname(__file__), "..", "data", "PEF_X.xlsx")]:
        validar(r)
