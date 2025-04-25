# -*- coding: utf-8 -*-
import mysql.connector
import json
from openpyxl import Workbook
from datetime import datetime, timedelta
from dateutil.relativedelta import relativedelta
import os

# Leer configuración desde el archivo JSON
def cargar_configuracion(ruta_archivo="date.json"):
    """
    Carga la configuración de la base de datos desde un archivo JSON.
    """
    try:
        with open(ruta_archivo, "r") as config_file:
            return json.load(config_file)
    except FileNotFoundError:
        raise FileNotFoundError(f"El archivo {ruta_archivo} no existe.")
    except json.JSONDecodeError:
        raise ValueError(f"El archivo {ruta_archivo} no tiene un formato JSON válido.")

# Calcular el rango de fechas
def obtener_rango_fechas(periodo):
    """
    Calcula el rango de fechas basado en el periodo especificado.
    """
    hoy = datetime.now()
    if periodo == "semanal":
        inicio = hoy - timedelta(days=7)
    elif periodo == "mensual":
        inicio = hoy - relativedelta(months=1)
    elif periodo == "anual":
        inicio = hoy - relativedelta(years=1)
    else:
        raise ValueError("Periodo no válido. Usa 'semanal', 'mensual' o 'anual'.")
    return inicio.strftime('%Y-%m-%d'), hoy.strftime('%Y-%m-%d')

# Generar el reporte en Excel
def generar_reporte(periodo, conexion, config):
    """
    Genera un reporte en formato Excel basado en el periodo especificado.
    """
    inicio, fin = obtener_rango_fechas(periodo)
    print(f"Generando reporte para el periodo: {periodo} ({inicio} a {fin})")

    cursor = conexion.cursor()

    # Crear el archivo Excel
    workbook = Workbook()
    workbook.remove(workbook.active)  # Eliminar la hoja predeterminada

    # Lista de consultas (sin modificar las consultas SQL existentes)
    consultas = [
        {
            "titulo": f"Clientes más valiosos ({inicio} a {fin})",
            "query": f"""
                SELECT Clientes.Nombre, SUM(precio_total) AS total_compras
                FROM Ventas
                JOIN Clientes ON Ventas.cliente = Clientes.id_cliente
                WHERE fecha BETWEEN '{inicio}' AND '{fin}'
                GROUP BY Clientes.Nombre
                ORDER BY total_compras DESC
                LIMIT 5;
            """
        },
        {
            "titulo": f"Productividad de los empleados ({inicio} a {fin})",
            "query": f"""
                SELECT Empleados.Nombre, SUM(precio_total) AS total_ventas
                FROM Ventas
                JOIN Empleados ON Ventas.empleado = Empleados.id_empleado
                WHERE fecha BETWEEN '{inicio}' AND '{fin}'
                GROUP BY Empleados.Nombre
                ORDER BY total_ventas DESC
                LIMIT 3;
            """
        },
        {
            "titulo": f"Productos más populares ({inicio} a {fin})",
            "query": f"""
                SELECT Productos.producto, SUM(cantidad_producto) AS total_vendido
                FROM Ventas
                JOIN Productos ON Ventas.producto = Productos.id_producto
                WHERE fecha BETWEEN '{inicio}' AND '{fin}'
                GROUP BY Productos.producto
                ORDER BY total_vendido DESC
                LIMIT 5;
            """
        },
        {
            "titulo": f"Análisis de proveedores ({inicio} a {fin})",
            "query": f"""
                SELECT Proveedores.Nombre, Productos.producto, SUM(cantidad_producto) AS total_vendido
                FROM Ventas
                JOIN Productos ON Ventas.producto = Productos.id_producto
                JOIN Proveedores ON Productos.proveedor = Proveedores.id_proveedor
                WHERE fecha BETWEEN '{inicio}' AND '{fin}'
                GROUP BY Proveedores.Nombre, Productos.producto
                ORDER BY total_vendido DESC
                LIMIT 3;
            """
        },
        {
            "titulo": f"Temporada más alta ({inicio} a {fin})",
            "query": f"""
                SELECT DATE_FORMAT(fecha, '%Y-%m') AS mes, SUM(precio_total) AS total_ventas
                FROM Ventas
                WHERE fecha IS NOT NULL AND precio_total IS NOT NULL 
                AND fecha BETWEEN '{inicio}' AND '{fin}'
                GROUP BY mes
                ORDER BY total_ventas DESC
                LIMIT 1;
            """
        },
        {
            "titulo": f"Fidelidad de clientes ({inicio} a {fin})",
            "query": f"""
                SELECT Clientes.Nombre, COUNT(Ventas.id_venta) AS total_compras, AVG(precio_total) AS promedio_gasto
                FROM Ventas
                JOIN Clientes ON Ventas.cliente = Clientes.id_cliente
                WHERE fecha BETWEEN '{inicio}' AND '{fin}'
                GROUP BY Clientes.Nombre
                ORDER BY total_compras DESC
                LIMIT 5;
            """
        }
    ]

    # Ejecutar las consultas y agregar los resultados al archivo Excel
    for consulta in consultas:
        cursor.execute(consulta["query"])
        resultados = cursor.fetchall()
        columnas = [desc[0] for desc in cursor.description]

        # Crear una hoja para cada consulta
        hoja = workbook.create_sheet(title=consulta["titulo"][:31])  # Limitar el título a 31 caracteres
        hoja.append(columnas)  # Agregar encabezados

        for fila in resultados:
            hoja.append(fila)  # Agregar filas de resultados

    # Leer la carpeta de salida desde el archivo JSON
    output_folder = config.get("output_folder", os.getcwd())  # Usa la carpeta actual si no se especifica

    # Verificar si la carpeta existe
    if not os.path.exists(output_folder):
        print(f"La carpeta {output_folder} no existe. Usando la carpeta actual.")
        output_folder = os.getcwd()

    # Guardar el archivo con un nombre dinámico
    nombre_archivo = f"Resultados_{periodo}_{datetime.now().strftime('%Y%m%d')}.xlsx"
    ruta_archivo = os.path.join(output_folder, nombre_archivo)
    workbook.save(ruta_archivo)
    print(f"Reporte guardado como: {ruta_archivo}")

    # Cerrar el cursor
    cursor.close()

# Función principal
def main():
    """
    Función principal para generar reportes en Excel.
    """
    try:
        # Cargar configuración
        config = cargar_configuracion()

        # Conectar a la base de datos
        conexion = mysql.connector.connect(
            host=config["host"],
            user=config["user"],
            password=config["password"],
            database=config["database"],
            port=config["port"]
        )

        # Generar reportes para diferentes periodos
        for periodo in ["semanal", "mensual", "anual"]:
            generar_reporte(periodo, conexion, config)

        # Cerrar la conexión
        conexion.close()
        print("Conexión a la base de datos cerrada.")
    except Exception as e:
        print(f"Error: {e}")

if __name__ == "__main__":
    main()