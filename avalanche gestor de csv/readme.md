# Gestor de CSV - Clientes y Emails

## Descripción
Esta aplicación permite importar datos desde archivos CSV, clasificarlos en categorías (clientes, contactos comerciales y emails no válidos) y almacenarlos en una base de datos SQLite. La aplicación mantiene un historial de todos los datos importados.

## Características
- Importación de datos desde archivos CSV
- Clasificación automática de contactos en tres categorías:
  - Clientes: Contactos personales
  - Comerciales: Contactos empresariales o profesionales
  - No válidos: Emails con formato incorrecto
- Validación automática de direcciones de email
- Almacenamiento persistente en base de datos SQLite
- Manejo inteligente de emails duplicados durante la importación
- Posibilidad de reclasificar contactos entre las categorías
- Eliminación de contactos (individual o múltiple)
- Interfaz gráfica intuitiva con colores pastel
- Funcionalidad de exportación de datos por categoría o completa

## Requisitos
- Python 3.x
- Módulos: tkinter, sqlite3, csv, re, os, datetime, PIL, ttkbootstrap

## Uso
1. Ejecute el script `csv_manager.py`
2. Haga clic en "Seleccionar archivo CSV" para importar datos
   - Active o desactive "Clasificación automática" según prefiera
   - La aplicación gestionará automáticamente los emails duplicados
3. Vea los datos clasificados en las pestañas correspondientes
4. Use el botón "Reclasificar Contacto" para mover contactos entre categorías
5. Use el botón "Eliminar Contacto(s)" para eliminar registros
6. Use los botones de exportación para guardar los datos en archivos CSV

## Estructura del CSV
El programa espera un archivo CSV con al menos dos columnas: una para nombres y otra para emails.
Si existe una columna con encabezado que contenga la palabra "empresa", "compañía" o "company",
se usará para ayudar a clasificar los contactos como comerciales.

## Lógica de clasificación automática
- Emails con dominios personales comunes (como gmail.com, hotmail.com, etc.) se clasifican como clientes
- Emails con dominios empresariales se clasifican como contactos comerciales
- Contactos que tienen una empresa asociada se clasifican como comerciales
