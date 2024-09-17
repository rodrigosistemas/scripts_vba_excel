# Macros útiles para el desarrollo de tareas en Excel

## Filtro automático
```vbnet
Sub filtrar()
    filtro = "*" & Sheets("hoja").TextBox1.Text & "*"
    Range("").AutoFilter field:=2, Criteria1:=filtro
End Sub
