import pandas as pd


def write_excel_file(data):
    df = pd.DataFrame(data.items(), columns=["Ruta", "Valor de la Celda"])
    df.to_excel("PyExcelUtilGui.xlsx", index=False)
