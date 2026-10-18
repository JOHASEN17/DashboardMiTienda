
import pandas as pd
import os

# ==========================================
# UBICACION DEL ARCHIVO EXCEL
# ==========================================

archivo = "LaTiendaCom (1).xls"

# Comprobar si existe el archivo
if not os.path.exists(archivo):
    print("ERROR: No se encontro el archivo Excel.")
    print("Coloca el Excel en la misma carpeta del programa.")

else:
    try:
        # Leer los nombres de las hojas
        excel = pd.ExcelFile(archivo)

        print("ARCHIVO EXCEL ENCONTRADO")
        print("Hojas disponibles:", excel.sheet_names)

        # ==========================================
        # REVISAR CADA HOJA
        # ==========================================

        for hoja in excel.sheet_names:
            print("\n" + "=" * 50)
            print("HOJA:", hoja)
            print("=" * 50)

            datos = pd.read_excel(
                archivo,
                sheet_name=hoja
            )

            print("\nColumnas:")
            print(datos.columns.tolist())

            print("\nCantidad de filas y columnas:")
            print(datos.shape)

            print("\nPrimeros cinco registros:")
            print(datos.head().to_string(index=False))

        print("\nRevision del Excel terminada.")

    except Exception as error:
        print("No se pudo leer el archivo.")
        print("Detalle:", error)