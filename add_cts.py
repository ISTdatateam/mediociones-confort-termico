import pandas as pd
from db.mysql_utils import MySQLDatabaseManager
import openpyxl


def poblar_centros_trabajo():
    """
    Lee un archivo Excel (centros_trabajo.xlsx) y carga sus datos en la tabla centros_trabajo.
    Se asume que el archivo tiene las columnas en el mismo orden que la tabla:
      - La primera columna es bigint.
      - Las columnas restantes son varchar.
    """
    db_manager = MySQLDatabaseManager()

    try:
        # Lee el archivo Excel. Ajusta el nombre del archivo si es necesario.
        df = pd.read_excel("DB-LOCALES-SMU-HIGIENE.xlsx")

        # Recorre cada fila del DataFrame y la inserta en la tabla.
        # Se asume que el archivo tiene exactamente 7 columnas.
        for index, row in df.iterrows():
            # Convertimos los valores de la fila a los tipos adecuados:
            # - La primera columna se convierte a entero (bigint).
            # - Las demás se convierten a cadena (varchar).
            valores = (
                int(row[0]),  # Primera columna (bigint)
                str(row[1]),  # Segunda columna (varchar)
                str(row[2]),  # Tercera columna (varchar)
                str(row[3]),  # Cuarta columna (varchar)
                str(row[4]),  # Quinta columna (varchar)
                str(row[5]),  # Sexta columna (varchar)
                str(row[6]),  # Séptima columna (varchar)
                str(row[7]),  # Séptima columna (varchar)
                int(row[8])  # Séptima columna (varchar)
            )

            # Si prefieres no especificar los nombres de columnas y que se inserte
            # en el mismo orden en que están definidas en la tabla, puedes usar:
            query = "INSERT INTO centros_trabajo (cuv, rut, razon_social, rut2, nombre_ct, direccion_ct, comuna_ct, region_ct, region_num_ct) VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s)"
            db_manager.cursor.execute(query, valores)

        # Confirma los cambios
        db_manager.connection.commit()
        print("Datos importados exitosamente en la tabla centros_trabajo.")

    except Exception as e:
        print(f"Error al importar datos: {e}")

    finally:
        db_manager.close()


if __name__ == "__main__":
    poblar_centros_trabajo()
