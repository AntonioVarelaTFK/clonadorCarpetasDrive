# 🧭 Clonador Carpetas Drive

Este script de Google Apps Script permite copiar de forma recursiva el contenido de una carpeta compartida de Drive en tu unidad personal, manteniendo un registro de los elementos copiados, permitiendo reanudar la copia en caso de interrupción (aún en pruebas) y añadiendo opciones de verificación, limpieza y compresión.

## 📘 Contexto

Tenemos una carpeta en Google Drive que contiene archivos y carpetas y más archivos y más carpetas… y queremos hacer una copia en otro Drive de otra cuenta que no pertenece a nuestro dominio.  
Hasta ahora las soluciones pasaban por herramientas externas, descargar y volver a subir, o hacerlo manualmente. Esta herramienta permite automatizar ese proceso **sin depender de software de terceros** y con una interfaz web sencilla.

## 🔧 Funcionalidades principales

- ✅ Copia recursiva de todos los archivos y subcarpetas.
- 🔄 Reanudación automática si se interrumpe la copia.
- 🔍 Verificación de integridad entre origen y destino.
- 🧹 Limpieza opcional del destino (elimina lo que ya no está en el origen).
- 🗂️ Registro JSON de cada operación (copia y compresión).
- 💾 Compresión en ZIP o múltiples ZIPs de hasta 100 MB cada uno:
  - Mantiene la estructura de carpetas en el fichero zip.
  - Convierte documentos de Google a formatos editables:
    - Google Docs → `.docx`
    - Google Sheets → `.xlsx`
    - Google Slides → `.pptx`
    - Google Forms no son exportables, pero se incluye un acceso directo (`.url`) a su ubicación copiada.
- 📊 Estimación previa del tamaño total y número de ZIPs antes de comprimir.
- 🖥️ Interfaz clara, en HTML con botones de acción.

## ▶ Cómo instalarlo

1. Crea un nuevo proyecto en Google Apps Script desde [script.google.com](https://script.google.com/).
2. Asígnale un nombre.
3. Copia los archivos `Code.gs` y `FormularioDestino.html` de este proyecto.
4. Prueba la implementación desde **Implementar** > **Implementación de prueba**.
5. Cuando estés satisfecho con los resultados, crea una nueva implementación desde **Implementar** > **Nueva implementación**.
6. Copia y guarda la **url** creada.

> ⚠️ Requiere autorización para acceder a tu Google Drive.

## ✅ Instrucciones de uso

![Interfaz](interfaz.png)

1. Abre el enlace de la app publicada.
2. Introduce el **ID de la carpeta origen**.
3. Indica el **nombre que tendrá la copia** en tu Drive.
4. Pulsa **Ejecutar copia** para iniciar la clonación.
5. Usa **Verificar integridad** o **Limpiar destino** según lo necesites.
6. Tras verificar correctamente, se habilitará la opción de **💾 Comprimir en ZIP(s)**.

> ⚠️ Requiere autorización para acceder a tu Google Drive.


## 📦 ¿Qué pasa con los archivos ZIP?

- Se guardan en una carpeta automática llamada **Backups ZIP**, junto a la carpeta destino.
- Los ZIP mantienen la estructura original y se nombran como `backup_<nombre>_part1_of_3.zip`, etc.
- Si un archivo no se puede convertir, se incluye en su formato por defecto, que suele ser una conversión a **pdf**.
- Se genera un archivo de log detallado con la información de cada ZIP.

## 📊 Comparativa de soluciones de copia de seguridad en Google Drive

| Solución                  | Ventajas principales                                                                                      | Inconvenientes                                                                                      | Precio      | Requiere instalación |
|---------------------------|-----------------------------------------------------------------------------------------------------------|------------------------------------------------------------------------------------------------------|-------------|-----------------------|
| **Clonador Carpetas Drive** | - Gratuito y personalizable<br>- Registro detallado y verificación<br>- Compresión con estructura de carpetas<br>- Conversión de documentos Google | - Limitado por el tiempo de ejecución de Apps Script<br>- No permite programar copias periódicas   | Gratuito    | No                   |
| **Kernel GDrive Backup**  | - Soporte multiusuario<br>- Migración incremental<br>- Filtros por fecha y carpetas                       | - De pago<br>- Requiere configuración técnica                                                       | De pago     | Sí                   |
| **Backupify**             | - Copias completas de Google Workspace<br>- Gestión en la nube<br>- Restauración eficiente                | - Precio elevado<br>- Orientado a empresas                                                          | De pago     | No                   |
| **SpinBackup**            | - Copia automatizada<br>- Restauración granular<br>- Protección frente a ransomware                        | - De pago<br>- Configuración avanzada                                                               | De pago     | No                   |
| **Rclone**                | - Muy flexible<br>- Compatible con múltiples nubes<br>- Automatizable y con cifrado opcional               | - Requiere uso por línea de comandos<br>- Instalación y configuración complejas                     | Gratuito    | Sí                   |

## 📂 Estructura del proyecto

- `Code.gs`: lógica principal del script.
- `FormularioDestino.html`: interfaz de usuario.
- `README.md`: este archivo.
- `ejemplo registro.json`: ejemplo del progreso de copia.
- `ejemplo log backup zip.json`: ejemplo de log de compresión.
- `interfaz.png`: captura de pantalla de la interfaz.
- `manual de uso.pdf`: guía detallada de instalación y uso.

## 📝 Licencia

Este proyecto se distribuye bajo la licencia MIT.  
Puedes modificarlo, adaptarlo y reutilizarlo libremente citando al autor original si lo deseas.

## ✍️ Autor

Antonio Varela · [https://antoniovarela.es](https://antoniovarela.es/)
