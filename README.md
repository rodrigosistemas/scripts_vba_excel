# Macros útiles para el desarrollo de tareas en Excel

## Filtro automático
```vbnet
Sub filtrar()
    filtro = "*" & Sheets("hoja").TextBox1.Text & "*"
    Range("").AutoFilter field:=2, Criteria1:=filtro
End Sub

Sub mostrar_hojas()
    Dim wks As Worksheet
    For Each wks In ActiveWorkbook.Worksheets
        wks.Visible = xlSheetVisible
    Next wks
End Sub

Sub OcultarHojaEspecifica()
    ' Verifica si la hoja existe y luego la oculta
    Dim hoja As Worksheet
    On Error Resume Next ' Evita errores si la hoja no existe
    Set hoja = ThisWorkbook.Sheets("NombreDeLaHoja") ' Reemplaza con el nombre de tu hoja
    On Error GoTo 0 ' Restablece el manejo normal de errores

    ' Si la hoja existe, la oculta
    If Not hoja Is Nothing Then
        hoja.Visible = xlSheetHidden ' Oculta la hoja (también puedes usar xlSheetVeryHidden para ocultarla de forma más estricta)
    Else
        MsgBox "La hoja especificada no existe."
    End If
End Sub
