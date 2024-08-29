# Macros útiles para el desarrollo de tareas en excel

## Filtro automático
Sub filtrar()
	filtro = "*" & Sheets("Clientes").TextBox1.Text & "*"
	Range("A7").AutoFilter field:=2, Criteria1:=filtro
End Sub

## Mostrar hojas muy ocultas
Sub mostrar_hojas()
    Dim wks As Worksheet
    For Each wks In ActiveWorkbook.Worksheets
        wks.Visible = xlSheetVisible
    Next wks
End Sub
