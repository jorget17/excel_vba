Attribute VB_Name = "Matrices"
'Definición de matriz: grupo de elementos del mismo tipo que tiene un nombre común.

'Para que todas la matrices empiecen en un indice distinto del 0 (por defecto) se puede usar (no recomendado)
'Option Base 1

Option Explicit

Sub declaracion_Matrices()

Dim rango1(4) As Integer

'Insertar valores en los índices de la matriz:

rango1(0) = 5
rango1(1) = 7
rango1(2) = 4
rango1(3) = 6
rango1(4) = 9

MsgBox rango1(3)

Dim rango2(0 To 4) As Integer

rango2(0) = 5
rango2(1) = 7
rango2(2) = 4
rango2(3) = 6
rango2(4) = 9

'Utilizar la ventana inmediato para depurar

Debug.Print rango2(4)

End Sub

Sub mostrar_valores()

'Vamos a almacenar en una matriz rango3 los valores del rango D1:D6 de la Hoja1

Dim rango3(5) As Integer

'Declaramos otra variable de tipo objeto (en este caso rango)
Dim celda As Range

'Declaramos una variable contandor
Dim indice As Integer
'Al no iniciar la variable indice se le asigna el valor por defecto (0)

'Obtenemos el rango que nos interesa
Range("D1").Select
'Selecciona todas las celdas adyacentes a la indicada (D1)
Selection.CurrentRegion.Select


'Creamos un for each loop para obtener los valores

For Each celda In Selection
    rango3(indice) = celda.Value
    indice = indice + 1
Next celda

'Mostramos los valores

Dim i As Integer

For i = 0 To 5 Step 1
    Debug.Print rango3(i)
Next i

End Sub
