from database import SessionLocal
from models import Usuario
import traceback

print("Intentando conectar a la base de datos...")
db = SessionLocal()

try:
    nuevo_usuario = Usuario(email="test_firme@astillero.com", password_hash="123456")
    db.add(nuevo_usuario)
    db.commit()
    print("¡ÉXITO MANO! El usuario se guardó en la tabla.")
except Exception as e:
    print("\n" + "="*40)
    print("AQUÍ ESTÁ EL BENDITO ERROR:")
    print("="*40)
    traceback.print_exc()
finally:
    db.close()