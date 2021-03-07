Attribute VB_Name = "variable_objeto"
Option Explicit

Sub variables_objeto()

'Las variables objeto son aquellas cuyos datos no son de tipo primitivo (integer, string, boolean... etc) sino de tipo celda, rango, libro, hoja...

Dim celda As Range

'Podemos utilizar una variable objeto para simplificar el código al acceder a objetos
'Imaginemos que queremos introducir el valor 124 en B5, poner el texto en negrita y el fondo en amarillo

'Worksheets(1).Range("B5").Value = 124
'Worksheets(1).Range("B5").Font.Bold = True
'Worksheets(1).Range("B5").Interior.Color = vbYellow

'Esto puede simplificarse si asignamos la ruta al objeto B5 a una variable

Set celda = Worksheets(1).Range("B5")

'Y ahora utilizamos este objeto para modificar sus atributos

'celda.Value = 124
'celda.Font.Bold = True
'celda.Interior.Color = vbYellow

'Incluso puede quedar mucho más simplificado utilizando un with.. end with

With celda
    .Value = 124
    .Font.Bold = True
    .Interior.Color = vbYellow
End With

End Sub

Sub formato_condicional()

'Crear un procedimiento que cambie el color de las celdas del rango B1 a B100 a verde si el valor de la celda es superior a 1000

Dim rango As Range

Set rango = Worksheets(1).Range("B1:B100")

Dim celda As Range

For Each celda In rango
    If celda.Value > 1000 Then
        celda.Interior.Color = vbGreen
    End If
Next celda

End Sub
