Attribute VB_Name = "Matrices_bidimensionales"
Option Explicit

Sub declaracion_matrices()

Dim matriz1(3, 4) As Integer

Dim x As Integer, y As Integer

For x = 0 To 3 Step 1
    For y = 0 To 4 Step 1
        matriz1(x, y) = Math.Round(Math.Rnd * 100)
        Debug.Print "("; x & ", " & y & ") --> " & matriz1(x, y)
    Next y
Next x
        
End Sub
