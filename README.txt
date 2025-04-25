===========================================
GENERACIÓN AUTOMÁTICA DE REPORTES
===========================================

Este proyecto genera reportes en formato Word y Excel basados en consultas SQL. Los reportes se generan para periodos específicos: semanal, mensual y anual.

A continuación, se explican los pasos para configurar y usar este proyecto.

-------------------------------------------
1. REQUISITOS PREVIOS
-------------------------------------------
Antes de usar este proyecto, asegúrate de cumplir con los siguientes requisitos:
- Python 3 instalado en tu sistema.
- Las siguientes bibliotecas de Python instaladas:
  - mysql-connector-python
  - python-docx
  - openpyxl
  - python-dateutil
- Una base de datos MySQL configurada con las tablas necesarias.
- Un archivo `date.json` con la configuración de conexión a la base de datos.

-------------------------------------------
2. CONFIGURAR EL ARCHIVO `date.json`
-------------------------------------------
Crea un archivo llamado `date.json` en el mismo directorio que los scripts. Este archivo debe contener la configuración de conexión a tu base de datos MySQL. Ejemplo:

{
    "host": "localhost",
    "user": "tu_usuario",
    "password": "tu_contraseña",
    "database": "nombre_de_tu_base_de_datos",
    "port": 3306
}

-------------------------------------------
3. PROBAR LOS SCRIPTS
-------------------------------------------
Antes de automatizar la ejecución, prueba los scripts manualmente para asegurarte de que funcionan correctamente.

1. Abre una terminal o consola.
2. Ejecuta el script `word.py` para generar reportes en formato Word:
3. Ejecuta el script `excel.py` para generar reportes en formato Excel:
4. Verifica que los archivos generados se guarden en el mismo directorio que los scripts. Los nombres de los archivos incluirán el periodo y la fecha de generación, por ejemplo:
- `Resultados_semanal_20250421.docx`
- `Resultados_mensual_20250421.xlsx`

-------------------------------------------
4. AUTOMATIZAR LA EJECUCIÓN
-------------------------------------------
Para que los reportes se generen automáticamente cada semana, mes y año, configura el Programador de Tareas en Windows o un cron job en Linux.

-------------------------------------------
4.1. USAR EL PROGRAMADOR DE TAREAS EN WINDOWS
-------------------------------------------
1. Abre el Programador de Tareas (Win + S y busca "Programador de Tareas").
2. Haz clic en "Crear tarea".
3. Configura la tarea:
- **General**:
  - Asigna un nombre a la tarea (por ejemplo, "Generar Reportes").
  - Marca la opción "Ejecutar con los privilegios más altos".
- **Desencadenadores**:
  - Haz clic en "Nuevo" y selecciona la frecuencia (semanal, mensual, etc.).
  - Configura la fecha y hora de inicio.
- **Acciones**:
  - Haz clic en "Nuevo" y selecciona "Iniciar un programa".
  - En el campo "Programa/script", escribe:
    ```
    python
    ```
  - En el campo "Agregar argumentos", escribe la ruta completa del script que deseas ejecutar. Por ejemplo:
    ```
    [word.py](http://_vscodecontentref_/1)
    ```
- Guarda la tarea y verifica que se ejecute correctamente.

-------------------------------------------
4.2. USAR UN CRON JOB EN LINUX
-------------------------------------------
1. Abre el archivo de configuración de cron:
2. Agrega una línea para programar la ejecución del script. Ejemplo:
- Para ejecutar el script semanalmente (lunes a las 9:00 AM):
  ```
  0 9 * * 1 python3 /ruta/completa/a/word.py
  ```
- Para ejecutarlo mensualmente (día 1 de cada mes a las 9:00 AM):
  ```
  0 9 1 * * python3 /ruta/completa/a/word.py
  ```
3. Guarda y cierra el archivo.

-------------------------------------------
5. NOTAS IMPORTANTES
-------------------------------------------
- Asegúrate de que los scripts tengan acceso a la base de datos configurada en `date.json`.
- Si encuentras algún error, verifica los mensajes en la consola para identificar el problema.
- Los reportes generados se guardarán en el mismo directorio que los scripts.

-------------------------------------------
¡Gracias por usar este proyecto!
-------------------------------------------
