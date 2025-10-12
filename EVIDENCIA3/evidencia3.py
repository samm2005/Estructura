import random as rd
import sys
import sqlite3
from sqlite3 import Error
from datetime import datetime, timedelta
from openpyxl import Workbook 


DB_NAME = "35.db"

# --- Crear tablas ---
def crear_tablas():
    try:
        with sqlite3.connect(DB_NAME) as conn:
            micursor = conn.cursor()
            micursor.execute("""
                CREATE TABLE IF NOT EXISTS Usuarios (
                    clave INTEGER PRIMARY KEY,
                    nombre TEXT NOT NULL
                );
            """)
            micursor.execute("""
                CREATE TABLE IF NOT EXISTS Reservaciones (
                    folio INTEGER PRIMARY KEY,
                    nombre TEXT NOT NULL,
                    horario TEXT NOT NULL,
                    fecha TEXT
                );
            """)
            micursor.execute("""
                CREATE TABLE IF NOT EXISTS Salas (
                    clave INTEGER PRIMARY KEY,
                    nombre TEXT NOT NULL,
                    capacidad INTEGER NOT NULL
                );
            """)
            conn.commit()
            print("Tablas creadas o verificadas correctamente.")
    except Error as e:
        print(f"Error en la creación de tablas: {e}")

# --- Opción 1: Registrar una reservación ---
def registrar_reservacion():
    while True:
        try:
            valor_clave = int(input("¿Cuál es tu clave de cliente?: ").strip())
        except ValueError:
            print("Por favor ingresa un número válido.")
            continue

        try:
            with sqlite3.connect(DB_NAME) as conn:
                micursor = conn.cursor()
                micursor.execute("SELECT nombre FROM Usuarios WHERE clave = ?", (valor_clave,))
                registro = micursor.fetchone()
                if registro:
                    print(f"Cliente encontrado: {valor_clave}\t{registro[0]}")
                else:
                    print(f"No se encontró un cliente con la clave {valor_clave}.")
                    continue
        except Exception as e:
            print(f"Error al consultar cliente: {e}")
            continue

        nombre_reserva = input("Ingrese el nombre de la reservación (Escribe SALIR para regresar al menú): ").strip()
        if nombre_reserva.upper() == 'SALIR':
            return

        horario = input("¿Cuál es el horario que quieres? (M, V, N): ").strip().upper()
        if horario not in ('M', 'V', 'N'):
            print("Horario inválido. Usa M, V o N.")
            continue

        fecha_ingresada = input("Ingresa la fecha de reservación en este formato dd/mm/aaaa: ").strip()
        try:
            fecha_dt = datetime.strptime(fecha_ingresada, '%d/%m/%Y')
        except ValueError:
            print("Formato de fecha no válido. Usa dd/mm/aaaa")
            continue

        fecha_minima = datetime.now() + timedelta(days=2)
        if fecha_dt < fecha_minima:
            print("La reservación debe ser con al menos 2 días de anticipación.")
            continue

        fecha_str = fecha_dt.strftime("%Y-%m-%d")  # guardamos como texto YYYY-MM-DD

        try:
            with sqlite3.connect(DB_NAME) as conn:
                micursor = conn.cursor()
                # generar folio único
                while True:
                    folio = rd.randint(1, 9999)
                    micursor.execute("SELECT 1 FROM Reservaciones WHERE folio = ?", (folio,))
                    if not micursor.fetchone():
                        break
                micursor.execute(
                    "INSERT INTO Reservaciones (folio, nombre, horario, fecha) VALUES (?, ?, ?, ?)",
                    (folio, nombre_reserva, horario, fecha_str)
                )
                conn.commit()
                print(f"¡Reservación realizada con éxito! Folio: {folio}")
        except Exception as e:
            print(f"Surgió una falla al insertar reservación: {e}")
        return

# --- Opción 2: Modificar descripción de la reservación ---
def modificar_descripcion():
    nombre_buscar = input("¿Cuál es el nombre de tu reservación?: ").strip()
    try:
        with sqlite3.connect(DB_NAME) as conn:
            micursor = conn.cursor()
            micursor.execute("SELECT folio, nombre, horario, fecha FROM Reservaciones WHERE nombre = ?", (nombre_buscar,))
            registros = micursor.fetchall()
            if not registros:
                print(f"No se encontró una reservación con el nombre {nombre_buscar}")
                return
            print("Folio\tNombre\tHorario\tFecha")
            folios = []
            for folio, nombre, horario, fecha in registros:
                fecha_mostrar = datetime.strptime(fecha, "%Y-%m-%d").strftime("%d/%m/%Y") if fecha else ""
                print(f"{folio}\t{nombre}\t{horario}\t{fecha_mostrar}")
                folios.append(folio)
    except Exception as e:
        print(f"Error al consultar reservaciones: {e}")
        return

    try:
        folio_elegido = int(input("Ingresa el folio que deseas modificar (o 0 para cancelar): "))
    except ValueError:
        print("Folio inválido.")
        return
    if folio_elegido == 0:
        return
    if folio_elegido not in folios:
        print("El folio indicado no coincide con los resultados mostrados.")
        return

    nuevo_nombre = input("¿A qué nombre lo quieres cambiar?: ").strip()
    if not nuevo_nombre:
        print("El nombre no puede quedar vacío.")
        return

    try:
        with sqlite3.connect(DB_NAME) as conn:
            micursor = conn.cursor()
            micursor.execute("UPDATE Reservaciones SET nombre = ? WHERE folio = ?", (nuevo_nombre, folio_elegido))
            conn.commit()
            print("Modificación realizada con éxito.")
    except Exception as e:
        print(f"Surgió una falla al actualizar: {e}")




# --- Opción 3: Consultar si una fecha está disponible ---
def consulta_fecha():
    fecha_consultar = input("Dime una fecha (dd/mm/aaaa): ").strip()
    try:
        fecha_dt = datetime.strptime(fecha_consultar, "%d/%m/%Y")
        fecha_str = fecha_dt.strftime("%Y-%m-%d")
    except ValueError:
        print("Formato inválido, usa dd/mm/aaaa")
        return

    try:
        with sqlite3.connect(DB_NAME) as conn:
            micursor = conn.cursor()
            micursor.execute("SELECT COUNT(*) FROM Reservaciones WHERE fecha = ?", (fecha_str,))
            count = micursor.fetchone()[0]
            if count > 0:
                print(f"La fecha {fecha_dt.strftime('%d/%m/%Y')} NO está disponible.")
            else:
                print(f"La fecha {fecha_dt.strftime('%d/%m/%Y')} está disponible :D")
    except Exception as e:
        print(f"Error al consultar fecha: {e}")




# --- Opción 4: Reporte de reservaciones por fecha ---
def reporte_reservaciones_fecha():
    fecha_consultar = input("Dime una fecha (dd/mm/aaaa): ").strip()
    try:
        fecha_dt = datetime.strptime(fecha_consultar, "%d/%m/%Y")
        fecha_str = fecha_dt.strftime("%Y-%m-%d")
    except ValueError:
        print("Formato inválido. Usa dd/mm/aaaa")
        return

    try:
        with sqlite3.connect(DB_NAME) as conn:
            micursor = conn.cursor()
            micursor.execute("SELECT folio, nombre, horario, fecha FROM Reservaciones WHERE fecha = ?", (fecha_str,))
            registros = micursor.fetchall()
            if registros:
                print(f"Reservaciones para la fecha {fecha_dt.strftime('%d/%m/%Y')}:")
                for folio, nombre, horario, fecha in registros:
                    fecha_mostrar = datetime.strptime(fecha, "%Y-%m-%d").strftime("%d/%m/%Y") if fecha else ""
                    print(f"Folio: {folio}, Nombre: {nombre}, Horario: {horario}, Fecha: {fecha_mostrar}")
            else:
                print("No hay reservaciones para esa fecha.")
    except Exception as e:
        print(f"Error: {e}")




# --- Opción 5: Registrar sala ---
def registrar_sala():
    while True:
        sala = input("¿Cómo se va a llamar la sala? (escribe SALIR para regresar al menú): ").strip()
        if sala.upper() == "SALIR":
            return
        if not sala:
            print("El nombre de la sala no puede estar vacío.")
            continue
        try:
            capacidad = int(input("¿Cuál va a ser la capacidad de la sala?: ").strip())
        except ValueError:
            print("Capacidad inválida, debe ser un número.")
            continue

        try:
            with sqlite3.connect(DB_NAME) as conn:
                micursor = conn.cursor()
                while True:
                    clave = rd.randint(1, 9999)
                    micursor.execute("SELECT 1 FROM Salas WHERE clave = ?", (clave,))
                    if not micursor.fetchone():
                        break
                micursor.execute("INSERT INTO Salas (clave, nombre, capacidad) VALUES (?, ?, ?)",
                                 (clave, sala, capacidad))
                conn.commit()
                print("Sala registrada.")
                print(f"Tu clave de sala es: {clave}")
        except Exception as e:
            print(f"Surgió una falla: {e}")
        return
    



# --- Opción 6: Registrar cliente ---
def registrar_cliente():
    while True:
        usuario = input("Ingresa el nombre del usuario: ").strip()
        if not usuario:
            print("El usuario no puede estar vacío.")
            continue
        try:
            with sqlite3.connect(DB_NAME) as conn:
                micursor = conn.cursor()
                while True:
                    clave = rd.randint(1, 9999)
                    micursor.execute("SELECT 1 FROM Usuarios WHERE clave = ?", (clave,))
                    if not micursor.fetchone():
                        break
                micursor.execute("INSERT INTO Usuarios (clave, nombre) VALUES (?, ?)", (clave, usuario))
                conn.commit()
                print("Usuario registrado.")
                print(f"Tu clave de usuario es: {clave}")
        except Exception as e:
            print(f"Surgió una falla: {e}")
        return
    



# --- Opción 7: Eliminar una reservación ---
def eliminar_reservacion():
    try:
        folio = int(input("Introduce el folio de la reservación que deseas eliminar: ").strip())
    except ValueError:
        print("Debes ingresar un número válido.")
        return

    try:
        with sqlite3.connect(DB_NAME) as conn:
            micursor = conn.cursor()
            micursor.execute("SELECT folio, nombre, horario, fecha FROM Reservaciones WHERE folio = ?", (folio,))
            registro = micursor.fetchone()
            if not registro:
                print(f"No existe una reservación con folio {folio}.")
                return
            fol, nombre, horario, fecha = registro
            fecha_mostrar = datetime.strptime(fecha, "%Y-%m-%d").strftime("%d/%m/%Y") if fecha else ""
            print("Se encontró la siguiente reservación:")
            print(f"Folio: {fol}, Nombre: {nombre}, Horario: {horario}, Fecha: {fecha_mostrar}")
            confirmar = input("¿Deseas eliminarla? (S/N): ").strip().upper()
            if confirmar == 'S':
                micursor.execute("DELETE FROM Reservaciones WHERE folio = ?", (folio,))
                conn.commit()
                print("Reservación eliminada correctamente.")
            else:
                print("Operación cancelada.")
    except Exception as e:
        print(f"Surgió una falla: {e}")




# --- Opción 8: Exportar a Excel ---
def exportar_a_excel():
    try:
        with sqlite3.connect(DB_NAME) as conn:
            micursor = conn.cursor()
            micursor.execute("SELECT folio, nombre, horario, fecha FROM Reservaciones ORDER BY fecha")
            reservaciones = micursor.fetchall()

        workbook = Workbook()
        hoja = workbook.active
        hoja.title = "RESERVACIONES"
        encabezados = ["Folio", "Nombre", "Horario", "Fecha"]
        hoja.append(encabezados)

        for folio, nombre, horario, fecha in reservaciones:
            # convertir fecha a dd/mm/aaaa para el reporte
            if fecha:
                try:
                    fecha_mostrar = datetime.strptime(fecha, "%Y-%m-%d").strftime("%d/%m/%Y")
                except Exception:
                    fecha_mostrar = fecha
            else:
                fecha_mostrar = ""
            hoja.append([folio, nombre, horario, fecha_mostrar])

        archivo = "Reporte_reservaciones.xlsx"
        workbook.save(archivo)
        print(f"Base de datos exportada a Excel como {archivo}")
    except Exception as e:
        print(f"Surgió una falla al exportar a Excel: {e}")



# --- Menú principal ---
def menu():
    crear_tablas()
    while True:
        print("\n........... M E N U ..............")
        print("1. Registrar reservación")
        print("2. Modificar reservación")
        print("3. Consultar disponibilidad por fecha")
        print("4. Reporte de reservación por fecha")
        print("5. Registrar sala")
        print("6. Registrar cliente")
        print("7. Eliminar reservación")
        print("8. Exportar base de datos a Excel")
        print("9. SALIR")
        try:
            opcion = int(input("Seleccione una opción (1-9): ").strip())
            if opcion == 1:
                registrar_reservacion()
            elif opcion == 2:
                modificar_descripcion()
            elif opcion == 3:
                consulta_fecha()
            elif opcion == 4:
                reporte_reservaciones_fecha()
            elif opcion == 5:
                registrar_sala()
            elif opcion == 6:
                registrar_cliente()
            elif opcion == 7:
                eliminar_reservacion()
            elif opcion == 8:
                exportar_a_excel()
            elif opcion == 9:
                print("Saliendo...")
                sys.exit(0)
            else:
                print("Opción no válida.")
        except ValueError:
            print("Entrada inválida, ingrese un número del 1 al 9")

menu()
