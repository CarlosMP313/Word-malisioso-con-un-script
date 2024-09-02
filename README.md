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

