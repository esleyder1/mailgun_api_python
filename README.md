# Aplicación de Notificación de Eventos Próximos
Esta es una aplicación en Python que permite leer un cronograma en formato Excel y determinar si un evento está próximo. En caso de que el evento esté próximo, se enviará una notificación por correo electrónico al destinatario correspondiente.

La aplicación se ejecuta automáticamente una vez al día, a la misma hora, y compara la fecha del cronograma con la fecha actual del equipo para determinar si es necesario enviar una notificación. A continuación, se presenta una guía para configurar y utilizar la aplicación.

### Requisitos
Python 3.6 o superior
Paquetes de Python:

### Instalación
Abre una terminal o línea de comandos.
Navega hasta el directorio raíz de la aplicación.
Ejecuta el siguiente comando para instalar los paquetes requeridos:

```
pip install -r requirements.txt
```
Esto instalará todos los paquetes necesarios para ejecutar la aplicación.

### Uso
Para ejecutar la aplicación, sigue estos pasos:

Abre una terminal o línea de comandos.
Navega hasta el directorio raíz de la aplicación.
Ejecuta el siguiente comando:
```
python main.py
```

# Ruta al archivo Excel del cronograma
CRONOGRAMA_ARCHIVO = "ruta/al/cronograma.xlsx"

# Detalles de la API de Mailgun
MAILGUN_API_KEY = "tu-api-key"
MAILGUN_DOMAIN = "tu-dominio-mailgun.com"
CORREO_EMISOR = "correo@emisor.com"
En CRONOGRAMA_ARCHIVO, proporciona la ruta completa al archivo Excel que contiene el cronograma.
En MAILGUN_API_KEY, ingresa tu clave de API proporcionada por Mailgun.
En MAILGUN_DOMAIN, especifica tu dominio de Mailgun.
En CORREO_EMISOR, ingresa la dirección de correo electrónico del remitente desde la cual se enviarán las notificaciones.
Asegúrate de tener conexión a Internet y acceso a una cuenta de Mailgun válida.

### Contribución
Si deseas contribuir a este proyecto, siéntete libre de hacerlo. Puedes abrir problemas (issues) para informar errores o sugerir mejoras. Además, las solicitudes de extracción (pull requests) son bienvenidas.

¡Gracias por usar nuestra aplicación! Si tienes alguna pregunta o necesitas asistencia, no dudes en ponerte en contacto con nosotros.
