import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill


def exportar_a_excel(df, ruta_salida):
    # Crear una hoja principal
    with pd.ExcelWriter(ruta_salida, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Datos procesados", index=False)

        # Crear hoja de resumen
        resumen = df.groupby("Nombre")["Horas Extra"].sum().reset_index()
        resumen["Horas Extra"] = pd.to_timedelta(resumen["Horas Extra"]).dt.total_seconds() / 3600
        resumen["Horas Extra"] = resumen["Horas Extra"].apply(lambda x: f"{int(x)}h {int((x % 1)*60)}m")
        resumen.to_excel(writer, sheet_name="Total Horas Extra", index=False)

    # Agregar filas amarillas por persona
    wb = load_workbook(ruta_salida)
    ws = wb["Datos procesados"]
    fill = PatternFill(start_color="FFF59D", end_color="FFF59D", fill_type="solid")

    nombre_actual = None
    fila_total = len(ws["A"]) + 2

    for i in range(2, ws.max_row + 1):
        nombre = ws[f"A{i}"].value
        if nombre != nombre_actual:
            if nombre_actual is not None:
                ws.append([f"Total de horas extras trabajadas por {nombre_actual}"])
                for celda in ws[ws.max_row]:
                    celda.fill = fill
            nombre_actual = nombre

    wb.save(ruta_salida)
