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

## Publicación y Release

Para generar el ejecutable de la aplicación, el proyecto incluye un script de automatización que crea un binario único (Self-contained, Single File) para Windows x64.

### Generar ejecutable
1. Abra una terminal en la raíz del proyecto.
2. Ejecute el siguiente comando:
   ```cmd
   publish-win-x64.cmd
   ```
3. El ejecutable se generará en: `LMM\bin\Release\net10.0-windows\win-x64\publish\LMM_1.0.0.exe` (el nombre incluirá la versión actual).

### Proceso de Release recomendado
1. **Incrementar versión**: Actualice la versión en `LMM.csproj` (ej. `<Version>1.1.0</Version>`).
2. **Pruebas**: Ejecute `dotnet test` para asegurar que todo funciona.
3. **Publicar**: Ejecute el script de publicación.
4. **Etiquetar**: Cree un tag en Git (ej. `git tag -a v1.1.0 -m "Release v1.1.0"`) y súbalo al repositorio.

---

## Git Flow

Este proyecto es totalmente compatible con el modelo de ramificación **Git Flow**. Se recomienda su uso para mantener un desarrollo organizado.

### Configuración inicial
Si tiene Git Flow instalado, inicialícelo en el proyecto:
```bash
git flow init
```
*(Se recomienda usar `master` como rama de producción y `develop` como rama de desarrollo).*

### Flujo de trabajo común
- **Nuevas funcionalidades**: 
  ```bash
  git flow feature start nombre-de-la-feature
  # ... desarrollar ...
  git flow feature finish nombre-de-la-feature
  ```
- **Preparar un release**:
  ```bash
  git flow release start 1.1.0
  # ... ajustes finales, actualizar versión en .csproj ...
  git flow release finish 1.1.0
  ```
- **Correcciones urgentes (Hotfix)**:
  ```bash
  git flow hotfix start fix-error-grave
  # ... corregir ...
  git flow hotfix finish fix-error-grave
  ```

---

## Pruebas Unitarias

El proyecto incluye un conjunto de pruebas unitarias para validar la lógica de negocio en el espacio de nombres `LMM.Application`.

Para ejecutar las pruebas:

1. Abra una terminal en la raíz del proyecto.
2. Ejecute el siguiente comando:
   ```cmd
   dotnet test
   ```

---

*Desarrollado para facilitar la creación de documentos personalizados de forma rápida y eficiente.*
