# Sistema de Envío Masivo de Emails

Una aplicación de escritorio desarrollada en Python para el envío masivo de emails personalizados.

## Características

- Interfaz gráfica intuitiva y moderna usando ttkbootstrap
- Soporte para importación de destinatarios mediante CSV
- Personalización de mensajes usando variables
- Sistema de logging en tiempo real
- Configuración SMTP flexible
- Diseño responsive con scroll vertical

## Requisitos

- Python 3.x
- Bibliotecas requeridas:
  - tkinter
  - ttkbootstrap
  - smtplib
  - email

Para instalar las dependencias:
```bash
pip install ttkbootstrap
```

## Estructura del CSV

El archivo CSV debe contener al menos una de estas columnas:
- email/Email/correo: La dirección de email del destinatario
- nombre/Nombre: El nombre del destinatario (opcional)

Ejemplo:
```csv
email,nombre
usuario@ejemplo.com,Juan Pérez
otro@ejemplo.com,María García
```

## Variables Disponibles

En el cuerpo del email puedes usar las siguientes variables:
- `{email}`: Se reemplazará con la dirección de email del destinatario
- `{nombre}`: Se reemplazará con el nombre del destinatario (si está disponible en el CSV)

## Configuración SMTP

Necesitarás los siguientes datos de tu servidor SMTP:
1. Dirección del servidor SMTP
2. Puerto SMTP (generalmente 587 para TLS)
3. Usuario SMTP
4. Contraseña SMTP
5. Dirección de email del remitente

### Ejemplo para Gmail

```
Servidor SMTP: smtp.gmail.com
Puerto: 587
Usuario: tu_email@gmail.com
Contraseña: tu_contraseña_de_aplicación
```

**Nota**: Para Gmail, necesitarás usar una "Contraseña de aplicación" específica.

## Uso

1. Ejecuta el script:
```bash
python envioemail.py
```

2. Configura los datos del servidor SMTP
3. Escribe el asunto del email
4. Añade los destinatarios (mediante CSV o manualmente)
5. Escribe el cuerpo del mensaje
6. Haz clic en "Enviar Emails"

## Consideraciones de Seguridad

- Las contraseñas se muestran ocultas en la interfaz
- Se recomienda usar conexiones SMTP con TLS
- No se almacenan datos sensibles de forma permanente

## Solución de Problemas

### Errores Comunes

1. **Error de conexión SMTP**
   - Verifica la configuración del servidor
   - Comprueba tu conexión a internet
   - Asegúrate de que el puerto no está bloqueado

2. **Error al importar CSV**
   - Verifica que el archivo está en formato UTF-8
   - Comprueba que las columnas tienen los nombres correctos

## Autor

Desarrollado por [franHR](https://pcprogramacion.es/)

## Licencia

Copyright © 2025 [pcprogramacion.es](https://pcprogramacion.es/)
Todos los derechos reservados.
