import bcrypt
from db.mysql_utils import MySQLDatabaseManager

def poblar_usuarios():
    """
    Pobla la tabla de usuarios con RUT, nombre y contraseñas hasheadas.
    """
    db_manager = MySQLDatabaseManager()
    usuarios = [
        {
            "email": "francisca.lobos@ist.cl",
            "pass": "password123",
            "name": "Francisca Lobos",
            "type": 1,
        }
    ]
    try:
        for usuario in usuarios:
            # Generar el hash de la contraseña
            hashed_password = bcrypt.hashpw(usuario["pass"].encode(), bcrypt.gensalt())

            # Depuración: Imprimir el hash generado
            print(f"Contraseña original: {usuario['pass']}")
            print(f"Hash generado: {hashed_password.decode()}")

            # Insertar el usuario en la base de datos
            query = """
                INSERT INTO usuarios (email, pass, name, type)
                VALUES (%s, %s, %s, %s)
            """
            db_manager.cursor.execute(
                query,
                (usuario["email"], hashed_password.decode(), usuario["name"], usuario["type"])
            )

        # Confirmar cambios
        db_manager.connection.commit()
        print("Usuarios insertados correctamente.")
    except Exception as e:
        print(f"Error al insertar usuarios: {e}")
    finally:
        db_manager.close()

if __name__ == "__main__":
    poblar_usuarios()
