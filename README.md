# Generación Automática de Reportes en Word y Excel

Este proyecto genera reportes en formato **Word** y **Excel** basados en consultas SQL. Los reportes se generan para periodos específicos: **semanal**, **mensual** y **anual**. Está diseñado para ser flexible y fácil de usar, permitiendo a los usuarios automatizar la generación de reportes mediante herramientas como el **Programador de Tareas** en Windows o **cron jobs** en Linux.

---

## 🚀 Funcionalidades

- Generación de reportes en formato **Word** (`.docx`) y **Excel** (`.xlsx`).
- Reportes basados en consultas SQL personalizables.
- Soporte para periodos de tiempo:
  - **Semanal**
  - **Mensual**
  - **Anual**
- Automatización mediante herramientas del sistema operativo.
- Configuración sencilla a través de un archivo JSON.

---

## 📋 Requisitos Previos

Antes de usar este proyecto, asegúrate de cumplir con los siguientes requisitos:

1. **Python 3** instalado en tu sistema.

2. Las siguientes bibliotecas de Python instaladas:
   - `mysql-connector-python`
   - `python-docx`
   - `openpyxl`
   - `python-dateutil`

3. Una base de datos **MySQL** configurada con las tablas necesarias.
4. Un archivo `date.json` con la configuración de conexión a la base de datos.

---

## ⚙️ Configuración

### 1. Configurar el Archivo `date.json`

Crea un archivo llamado `date.json` en el mismo directorio que los scripts. Este archivo debe contener la configuración de conexión a tu base de datos MySQL. Ejemplo:

```json
{
    "host": "localhost",
    "user": "tu_usuario",
    "password": "tu_contraseña",
    "database": "nombre_de_tu_base_de_datos",
    "port": 3306
    "output_folder": "Ruta_de_la_carpeta_en_donde_se_guardaran_los_archivos"
}

      output_folder: Especifica la carpeta donde se guardarán los reportes generados. Si no se especifica o la carpeta no existe, los reportes se guardarán en el directorio actual.

2. Crear la Base de Datos y Tablas

Si no tienes experiencia con SQL o necesitas configurar la base de datos desde cero, este proyecto incluye un archivo llamado database.mysql que contiene los scripts necesarios para crear la base de datos, las tablas y algunos datos de ejemplo.

Pasos para usar el archivo database.mysql:

  1. Abre tu cliente MySQL (como MySQL Workbench o la terminal de MySQL).
  2. Ejecuta el archivo database.mysql para crear la base de datos y las tablas necesarias:
       mysql -u tu_usuario -p < ruta/del/archivo/database.mysql

  3. Esto creará automáticamente:
      - La base de datos.
      -Tablas como Empleados, Clientes, Proveedores, Productos, Ventas, Pedidos e Inventario.
      -Algunos datos de ejemplo para comenzar.

¿Qué contiene el archivo database.mysql?
  - Creación de la base de datos: Define la base de datos que se usará para los reportes.
  - Creación de tablas: Incluye tablas como Empleados, Clientes, Proveedores, etc., con sus relaciones.
  - Inserción de datos de ejemplo: Proporciona datos iniciales para probar el proyecto.
  - Ejemplo de carga de datos desde un archivo CSV: Muestra cómo cargar datos masivos en una tabla.


⚠️RECUERDA: Este archivo es solo un ejemplo y puedes modificarlo
  según tus necesidades. La estructura de las tablas y los datos de ejemplo son personalizables. 
  No olvides ajustar las consultas SQL en los archivos 'word.py' y 'excel.py' para que coincidan con la estructura de tu base de datos. ⚠️

🛠️ Uso del Proyecto
1. Probar los Scripts Manualmente

  1. Abre una terminal o consola.

  2. Ejecutar el Script para Generar Reportes
  Utiliza el siguiente comando para generar un reporte específico:
  python word.py

  3. Ejecuta el script excel.py para generar reportes en formato Excel:
  python excel.py


2. Automatizar la Ejecución

Para que los reportes se generen automáticamente cada semana, mes y año, configura el Programador de Tareas en Windows o un cron job en Linux.

Usar el Programador de Tareas en Windows
 1. Abre el Programador de Tareas (Win + S y busca "Programador de Tareas").
 2. Haz clic en Crear tarea.
 3. Configura la tarea:
    General:
        - Asigna un nombre a la tarea (por ejemplo, "Generar Reportes").
        - Marca la opción Ejecutar con los privilegios más altos.
    Desencadenadores:
        - Haz clic en Nuevo y selecciona la frecuencia (semanal, 
          mensual, etc.).
         Configura la fecha y hora de inicio.
    Acciones:
        - Haz clic en Nuevo y selecciona Iniciar un programa.
        - En el campo Programa/script, escribe:
          python

        - En el campo Agregar argumentos, escribe la ruta completa del
          script que deseas ejecutar. Por ejemplo:
              c:\Users\jose\OneDrive\Desktop\sql\Mi_proyecto\archivos\word.py

    Guarda la tarea y verifica que se ejecute correctamente.


Usar un Cron Job en Linux
  1. Abre el archivo de configuración de cron:
     crontab -e

  2. Agrega una línea para programar la ejecución del script. Ejemplo:

      Para ejecutar el script semanalmente (lunes a las 9:00 AM):
      0 9 * * 1 python3 /ruta/completa/a/word.py

      Para ejecutarlo mensualmente (día 1 de cada mes a las 9:00 AM):
      0 9 1 * * python3 /ruta/completa/a/word.py

      Para ejecutarlo anualmente (1 de enero a las 9:00 AM):
      0 9 1 1 * python3 /ruta/completa/a/word.py

  3. Guarda y cierra el archivo.



📂 Estructura del Proyecto

Mi_proyecto/
│
├── archivos/
│   ├── word.py          # Genera reportes en formato Word
│   ├── excel.py         # Genera reportes en formato Excel
│   ├── date.json        # Configuración de conexión a la base de datos
│   └── README.md        # Documentación del proyecto
│
└── sql/
    └── script.sql       # Script SQL para crear las tablas necesarias
 

 📝 Notas Importantes
Asegúrate de que los scripts tengan acceso a la base de datos configurada en date.json.
Si encuentras algún error, verifica los mensajes en la consola para identificar el problema.
Los reportes generados se guardarán en el mismo directorio que los scripts.

🤝 Contribuciones
Si deseas contribuir a este proyecto, siéntete libre de hacer un fork y enviar un pull request. ¡Toda ayuda es bienvenida!

📧 Contacto
Si tienes preguntas o necesitas ayuda, no dudes en contactarme en yolic.luna.ps@gmail.com.



¡Gracias por usar este proyecto! 🎉
