
import pandas as pd
import os

archivo = "LaTiendaCom (1).xls"

print("Carpeta actual:", os.getcwd())

if not os.path.exists(archivo):
    print("\nNo se encuentra el archivo:", archivo)
    print("Archivos disponibles en esta carpeta:")
    for nombre in os.listdir():
        print("-", nombre)
else:
    try:
        excel = pd.ExcelFile(archivo)

        print("\nEXCEL ENCONTRADO")
        print("Hojas:", excel.sheet_names)

        for hoja in excel.sheet_names:
            datos = pd.read_excel(
                archivo,
                sheet_name=hoja
            )

            print("\n" + "=" * 50)
            print("HOJA:", hoja)
            print("Columnas:", datos.columns.tolist())
            print("Dimensiones:", datos.shape)
            print("\nPrimeros cinco registros:")
            print(datos.head().to_string(index=False))

        print("\nRevision terminada.")

    except Exception as error:
        print("\nError al leer el Excel:", error)
