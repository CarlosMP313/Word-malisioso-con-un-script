# Incrustar y Ejecutar un Script desde un Documento de Word usando Macros

Este tutorial muestra cómo incrustar y ejecutar un script desde un documento de Word utilizando macros. Esta técnica debe ser utilizada con cuidado y solo en entornos controlados debido a los riesgos de seguridad asociados.

## Habilitar la Pestaña Desarrollador en Word

1. Abre Microsoft Word.
2. Haz clic en "Archivo" > "Opciones".
3. Selecciona "Personalizar cinta de opciones".
4. Marca la casilla "Desarrollador".
5. Haz clic en "Aceptar".

## Abrir el Editor de Visual Basic (VBA)

1. Ve a la pestaña "Desarrollador".
2. Haz clic en "Visual Basic" en el grupo de "Código".

## Crear y Escribir un Script en un Nuevo Módulo

1. En el Editor de VBA, haz clic en "Insertar" > "Módulo".
2. Escribe el siguiente script en la ventana del módulo:

    ```vba
    Sub EjecutarScript()
        ' Este script muestra un mensaje de saludo
        MsgBox "Hola, este es un script ejecutado desde Word."
    End Sub
    ```

## Guardar el Documento Habilitado para Macros

1. Guarda el documento como **.docm**:
   - Ve a "Archivo" > "Guardar como".
   - Selecciona la ubicación para guardar.
   - En "Tipo", elige "Documento de Word habilitado para macros (*.docm)".
   - Guarda el archivo.

## Ejecutar la Macro Manualmente

1. Ve a la pestaña "Desarrollador".
2. Haz clic en "Macros".
3. Selecciona `EjecutarScript`.
4. Haz clic en "Ejecutar".

## Ejecutar una Macro Automáticamente al Abrir el Documento

1. En el Editor de VBA, haz doble clic en "ThisDocument" en la ventana del proyecto.
2. Escribe el siguiente código:

    ```vba
    Private Sub Document_Open()
        ' Llama a la macro EjecutarScript cuando se abre el documento
        Call EjecutarScript
    End Sub
    ```

## Ejemplo Avanzado: Ejecutar un Comando del Sistema

Aquí tienes un ejemplo más avanzado que ejecuta un comando del sistema (abre la calculadora):

```vba
Sub EjecutarComandoSistema()
    ' Ejecuta la calculadora usando el comando shell
    Dim comando As String
    comando = "calc.exe"
    Shell comando, vbNormalFocus
End Sub
```

## Código VBA para Mostrar Mensaje de Compromiso Múltiples Veces

Este script abre el símbolo del sistema y muestra el mensaje "Tu sistema ha sido comprometido" cinco veces.

```vba
Sub MostrarMensajeDeCompromisoMultiplesVeces()
    ' Ejecuta el símbolo del sistema y muestra el mensaje 5 veces
    Dim comando As String
    Dim i As Integer
    
    For i = 1 To 5
        comando = "cmd.exe /c echo Tu sistema ha sido comprometido & pause"
        Shell comando, vbNormalFocus
    Next i
End Sub
```

## Código VBA para Ejecutar un Comando como Administrador

Este script utiliza la API de Windows `ShellExecute` para ejecutar el símbolo del sistema como administrador y mostrar el mensaje "Tu sistema ha sido comprometido".

```vba
Private Declare PtrSafe Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" ( _
    ByVal hwnd As LongPtr, _
    ByVal lpOperation As String, _
    ByVal lpFile As String, _
    ByVal lpParameters As String, _
    ByVal lpDirectory As String, _
    ByVal nShowCmd As Long) As Long

Sub EjecutarComoAdministrador()
    Dim resultado As Long
    Dim comando As String
    comando = "cmd.exe /c echo Tu sistema ha sido comprometido & pause"
    
    ' Ejecutar cmd.exe como administrador
    resultado = ShellExecute(0, "runas", "cmd.exe", "/c echo Tu sistema ha sido comprometido & pause", vbNullString, 1)
    
    If resultado <= 32 Then
        MsgBox "Error al intentar ejecutar como administrador."
    End If
End Sub
```
## Exención de Responsabilidad

La información y los archivos adjuntos proporcionados en este repositorio se ofrecen únicamente con fines educativos y de prueba. **No asumo ninguna responsabilidad por el uso indebido o malintencionado de los contenidos aquí presentes**. Los scripts y ejemplos de código son ilustrativos y deben ser utilizados únicamente en entornos controlados y seguros.

**Advertencia:** La ejecución de macros o scripts incrustados en documentos puede presentar riesgos de seguridad. Se recomienda encarecidamente no utilizar estos ejemplos en entornos de producción o en sistemas que contengan información sensible.

El uso de este material es bajo tu propio riesgo. No se proporciona garantía alguna sobre la precisión, seguridad o idoneidad de los contenidos para cualquier propósito específico. **Al utilizar este repositorio, aceptas que no me hago responsable de ninguna consecuencia, daño o pérdida que pueda resultar del uso de la información y archivos aquí compartidos**.


