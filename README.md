# Copiador de carpetas compartidas en Google Drive

Este script de Google Apps Script permite copiar de forma recursiva el contenido de una carpeta compartida de Drive en tu unidad personal, manteniendo un registro de los elementos copiados y permitiendo reanudar la copia en caso de interrupción.

## 🔧 Funcionalidades principales

- Copia recursiva de carpetas y archivos.
- Detección de cambios en tamaño o fecha de modificación.
- Registro detallado de las copias realizadas.
- Verificación de integridad entre origen y destino.
- Limpieza de elementos obsoletos en el destino.

## ▶ Cómo usarlo

1. Crea un nuevo proyecto en Google Apps Script y copia los archivos `code.gs` y `FormularioDestino.html`.
2. Despliega el proyecto como aplicación web (con permisos de edición).
3. Introduce el ID de la carpeta compartida y un nombre para la carpeta destino.
4. Usa los botones para lanzar la copia, verificar cambios o limpiar el destino.

> ⚠️ Requiere autorización para acceder a tu Google Drive.

## 📂 Estructura del proyecto

- `code.gs`: lógica principal del script.
- `FormularioDestino.html`: interfaz de usuario.
- `appsscript.json`: configuración del proyecto.
- `LICENCIA`: licencia del repositorio.
- `README.md`: este archivo.

## 📝 Licencia

Este proyecto se distribuye bajo la licencia MIT.