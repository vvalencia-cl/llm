# LMM - Combinación de Correspondencia (Word + Excel → PDF)

Este proyecto es una aplicación de escritorio para Windows desarrollada en **C# WinForms** que automatiza la generación de archivos PDF a partir de una plantilla de Word (.docx) y un origen de datos en Excel (.xlsx).

## Funcionalidades

- **Generación masiva de PDFs**: Crea un archivo PDF individual por cada fila de datos en una hoja de Excel.
- **Mapeo Inteligente de Campos**: Vincula automáticamente los `MERGEFIELD` de la plantilla de Word con las columnas de Excel, con soporte para normalización de nombres (espacios, guiones, caracteres especiales).
- **Configuración de Nombre de Archivo**: Permite definir el nombre de los PDFs generados utilizando prefijos, sufijos y hasta tres campos dinámicos de las columnas de Excel.
- **Vista Previa en Tiempo Real**: Muestra una vista previa del nombre del archivo resultante según la configuración elegida.
- **Control de Carpeta de Salida**:
    - Opción para **borrar el contenido** de la carpeta de salida antes de iniciar el proceso.
    - Botón para **abrir la carpeta de salida** directamente una vez finalizada la combinación.
- **Procesamiento en Segundo Plano**: Ejecuta la automatización de Word en un hilo separado (STA) para mantener la interfaz de usuario fluida, permitiendo la cancelación en cualquier momento.
- **Registro de Actividad (Log)**: Muestra el progreso detallado y errores específicos por fila, permitiendo copiar el log al portapapeles.

## Requisitos

- **Microsoft Word**: Es necesario tener instalado Word en el sistema, ya que la aplicación utiliza Interop para la exportación precisa a PDF.
- **Windows**: Compatible con .NET 10.0 en Windows.

## Publicación

Para generar el ejecutable de la aplicación, el proyecto incluye un script de automatización.

1. Abra una terminal en la raíz del proyecto.
2. Ejecute el siguiente comando:
   ```cmd
   publish-win-x64.cmd
   ```
3. El ejecutable y sus dependencias se generarán en la carpeta:
   `LMM\bin\Release\net10.0-windows\win-x64\publish`

---

*Desarrollado para facilitar la creación de documentos personalizados de forma rápida y eficiente.*
