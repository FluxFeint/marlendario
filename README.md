# Marlendario

### Descripción

Este proyecto es una aplicación de escritorio desarrollada en Python con Tkinter que permite a los usuarios seleccionar su horario laboral desde un archivo PDF, elegir una semana específica y sincronizar estos horarios con Google Calendar automáticamente. La aplicación ofrece una interfaz gráfica fácil de usar para seleccionar archivos, identificar IDs únicos dentro del PDF, y manejar la autenticación y eventos en Google Calendar.

Esta aplicación resuelve un caso muy específico y para que funcione la tabla con los horarios el pdf deberá tener esta estructura:

![Estructura de nuestro PDF](https://raw.githubusercontent.com/FluxFeint/marlendario/main/recursos/estructuraReadMe.png)

### Características
- **Selección de Archivo PDF:** Permite a los usuarios cargar un archivo PDF desde su sistema para extraer horarios.
- **Extracción de IDs**: Analiza el PDF para obtener identificadores únicos asociados con horarios.
- **Selección de Usuario**: Los usuarios pueden seleccionar su ID correspondiente para filtrar sus horarios específicos.
- **Integración con Google Calendar**: Sincroniza automáticamente los horarios seleccionados con el calendario del usuario en Google Calendar.
- **Interfaz Gráfica**: Provee una interfaz gráfica intuitiva para una mejor experiencia de usuario.

### Tecnologías Utilizadas
- Python 3
- Tkinter para la interfaz gráfica de usuario.
- pdfplumber para la extracción de datos de archivos PDF.
- API de Google Calendar para la gestión de eventos en el calendario.

### Configuración del Proyecto
#### Pre-requisitos
- Python 3.8 o superior.
- Acceso a Internet para la autenticación de Google Calendar y la manipulación de eventos.

### Instalación
Instalación

- Clona el repositorio:

`git clone https://github.com/FluxFeint/marlendario.git`

- Instala las dependencias necesarias:

`pip install -r requirements.txt`

- Necesitaras un archivo con credenciales con permisos para la api de Google Calendars que podras obtener aquí: https://console.cloud.google.com/projectselector2/apis/dashboard


### Uso
Para ejecutar la aplicación, navega al directorio del proyecto y ejecuta:
`python main.py`
