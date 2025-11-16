import sqlite3
from openpyxl import Workbook
from datetime import datetime
import sys

DB_NAME = "afis.db"

# --- Funciones auxiliares ---
def ejecutar_sql(q, p=(), fetch=False):
    with sqlite3.connect(DB_NAME) as conn:
        cur = conn.cursor()
        cur.execute(q, p)
        if fetch: return cur.fetchall()
        conn.commit()

def crear_tabla():
    ejecutar_sql("""
        CREATE TABLE IF NOT EXISTS Asistencias (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            nombre TEXT NOT NULL,
            matricula TEXT NOT NULL,
            categoria TEXT NOT NULL,
            fecha TEXT NOT NULL
        )
    """)

def semestre_actual():
    hoy = datetime.now()
    return hoy.year, 1 if hoy.month <= 6 else 2

# --- Registrar asistencia (máx. 2 por semestre) ---
def registrar_asistencia():
    print("\n=== REGISTRO DE ASISTENCIA ===")
    matricula = input("Matrícula: ").strip().upper()
    if not matricula:
        return print(" La matrícula es obligatoria.")

    # Obtener nombre si ya existe
    alumno = ejecutar_sql("SELECT nombre FROM Asistencias WHERE matricula=? LIMIT 1", (matricula,), True)
    if alumno:
        nombre = alumno[0][0]
        print(f"Alumno detectado: {nombre}")
    else:
        nombre = input("Nombre: ").strip()
        if not nombre:
            return print(" El nombre es obligatorio.")

    # Validar límite de 2 por semestre
    año, sem = semestre_actual()
    f_ini, f_fin = (f"{año}-01-01", f"{año}-06-30") if sem == 1 else (f"{año}-07-01", f"{año}-12-31")
    total = ejecutar_sql("SELECT COUNT(*) FROM Asistencias WHERE matricula=? AND fecha BETWEEN ? AND ?", 
                         (matricula, f_ini, f_fin), True)[0][0]
    if total >= 2:
        return print(f" Ya tienes 2 AFIS registrados este semestre ({año}-{sem}).")

    print(f"Llevas {total}/2 AFIS este semestre.")

    categorias = [
        "CULTURAL", "ARTISTICA", "RESPONSABILIDAD SOCIAL", "ACADEMICA", "INNOVACION Y EMPRENDIMIENTO"
    ]
    for i, c in enumerate(categorias, 1):
        print(f"{i}. {c}")
    try:
        categoria = categorias[int(input("Categoría (número): ")) - 1]
    except:
        return print(" Opción inválida.")

    fecha = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    ejecutar_sql("INSERT INTO Asistencias (nombre, matricula, categoria, fecha) VALUES (?, ?, ?, ?)",
                 (nombre, matricula, categoria, fecha))
    print(" Asistencia registrada correctamente.")

# --- Consultar ---
def consultar_asistencias():
    registros = ejecutar_sql("SELECT * FROM Asistencias ORDER BY fecha DESC", fetch=True)
    if not registros: return print("No hay registros.")
    print("\n=== LISTADO DE ASISTENCIAS ===")
    for r in registros:
        print(f"{r[0]} | {r[1]} | {r[2]} | {r[3]} | {r[4]}")

# --- Modificar ---
def modificar_asistencia():
    try: idr = int(input("ID a modificar: "))
    except: return print("ID inválido.")
    n_nombre = input("Nuevo nombre (Enter = igual): ").strip()
    n_mat = input("Nueva matrícula (Enter = igual): ").strip()
    n_cat = input("Nueva categoría (Enter = igual): ").strip()
    if n_nombre: ejecutar_sql("UPDATE Asistencias SET nombre=? WHERE id=?", (n_nombre, idr))
    if n_mat: ejecutar_sql("UPDATE Asistencias SET matricula=? WHERE id=?", (n_mat, idr))
    if n_cat: ejecutar_sql("UPDATE Asistencias SET categoria=? WHERE id=?", (n_cat, idr))
    print(" Registro actualizado.")

# --- Eliminar ---
def eliminar_asistencia():
    try: idr = int(input("ID a eliminar: "))
    except: return print("ID inválido.")
    reg = ejecutar_sql("SELECT * FROM Asistencias WHERE id=?", (idr,), True)
    if not reg: return print("No se encontró el ID.")
    if input(f"¿Eliminar asistencia de {reg[0][1]}? (S/N): ").upper() == "S":
        ejecutar_sql("DELETE FROM Asistencias WHERE id=?", (idr,))
        print("✅ Eliminado.")

# --- Reiniciar ---
def reiniciar_ids():
    if input(" ¿Borrar todo y reiniciar IDs? (S/N): ").upper() != "S":
        return print("Cancelado.")
    ejecutar_sql("DELETE FROM Asistencias")
    ejecutar_sql("DELETE FROM sqlite_sequence WHERE name='Asistencias'")
    print(" Registros borrados y contador reiniciado.")

# --- Exportar ---
def exportar_excel():
    registros = ejecutar_sql("SELECT * FROM Asistencias ORDER BY fecha", fetch=True)
    if not registros: return print("No hay registros para exportar.")
    wb = Workbook()
    hoja = wb.active
    hoja.title = "Asistencias AFIS"
    hoja.append(["ID", "Nombre", "Matrícula", "Categoría", "Fecha"])
    for fila in registros: hoja.append(fila)
    archivo = f"Registro_AFIS_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.xlsx"
    wb.save(archivo)
    print(f" Exportado a '{archivo}'")

# --- Menú principal ---
def menu():
    crear_tabla()
    opciones = {
        "1": registrar_asistencia,
        "2": consultar_asistencias,
        "3": modificar_asistencia,
        "4": eliminar_asistencia,
        "5": exportar_excel,
        "6": reiniciar_ids,
        "7": lambda: sys.exit(" Saliendo...")
    }
    while True:
        print("""
====== SISTEMA DE ASISTENCIAS AFIS ======
1. Registrar asistencia
2. Consultar asistencias
3. Modificar registro
4. Eliminar registro
5. Exportar a Excel
6. Reiniciar registros
7. Salir
""")
        (opciones.get(input("Elige una opción: ").strip()) or (lambda: print("Opción inválida.")))()

menu()

