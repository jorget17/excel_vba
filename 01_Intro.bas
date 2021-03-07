Attribute VB_Name = "Intro"
Sub insertar()

'Insertar valores en celdas
Range("B5") = 250

End Sub

Sub borrar_contenido()

'Borra el contenido (el valor) de la celda
Range("B5").ClearContents

End Sub

Sub borrar_todo()

'Borrar el contenido y el formato de la celda
Range("B5").Clear

End Sub

Sub destacar()

'Diferentes formas de elegir colores

'Colores básicos con vbColor
Range("B5").Interior.Color = vbGreen

'Utilizando valores rgb
Range("B5").Interior.Color = RGB(0, 255, 0)

End Sub
