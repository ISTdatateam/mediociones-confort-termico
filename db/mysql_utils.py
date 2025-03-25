import mysql.connector
from mysql.connector import Error
from dotenv import load_dotenv
import os

# Cargar variables de entorno
load_dotenv()

class MySQLDatabaseManager:
    def __init__(self):
        """Inicializa la conexión a la base de datos MySQL."""
        self.connection = None
        self.cursor = None
        try:
            self.connection = mysql.connector.connect(
                host=os.getenv("DB_HOST"),
                user=os.getenv("DB_USERNAME"),
                password=os.getenv("DB_PASS"),
                database=os.getenv("DB_NAME"),
                port=int(os.getenv("DB_PORT", 3306)),
            )
            self.cursor = self.connection.cursor(dictionary=True)
        except Error as e:
            print(f"Error al conectar a MySQL: {e}")
            self.connection = None
            self.cursor = None

    def fetch_one(self, query, params=None):
        """Ejecuta una consulta SQL y retorna un solo resultado."""
        if not self.cursor:
            print("Error: No hay conexión a MySQL.")
            return None
        try:
            self.cursor.execute(query, params)
            return self.cursor.fetchone()
        except Error as e:
            print(f"Error al ejecutar la consulta: {e}")
            return None

    def close(self):
        """Cierra la conexión a MySQL."""
        if self.cursor:
            self.cursor.close()
        if self.connection:
            self.connection.close()
