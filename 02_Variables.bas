Attribute VB_Name = "Variables"
'Obligar la declaraci�n de variables mediante el comando
'Option Explicit

'Dim nombre As String
'Public nombre As String
'Public Const valor_constante As Byte = 7
Public Const euro As Double = 166.386


Sub tipo_datos()

'byte -- n�meros de 0 a 255
'integer -- n�meros enteros -32768 a 32768
'long -- n�meros enteros grandes
'single -- n�meros decimales (parte decimal corta)
'double -- n�meros decimales (parte decimal larga)
'decimal -- n�meros decimales (parte decimal extremadamente larga)
'boolean -- verdadero (true) o falso (false)
'currency -- datos de tipo moneda
'object -- objetos
'string -- cadena de caracteres
'variant -- sin especificar

End Sub

Sub declaracion_variables()

'declaracion de variables

Dim nombre As String

nombre = "Marta"

Dim edad As Byte

edad = 48

'Declaraci�n de variables m�ltiples
'Dim nombre as string, edad as byte

'Asignar valores m�ltiples
'nombre = "Marta": edad = 48

'Usar variables en un procedimiento
MsgBox "La usuaria " & nombre & " tiene " & edad & " a�os."

End Sub

Sub scope_variables()

'Las variables se pueden declarar en diferentes �mbitos (scopes)

'1. Local a nivel de procedimiento: dentro del procedimiento en el que se declara

'Declarando la variable dentro del procedimiento

'2. Local a nivel de m�dulo: dentro del m�dulo en el que se declara

'Declarando la variable antes de los procedimientos (ver arriba) e inici�ndolas en el momento de su uso (en cualquier procedimiento del m�dulo).
'Tambi�n podemos llamar al procedimiento que incluye la declaraci�n de la variable que queremos utilizar.

'Call declaracion_variables
'Esto permite utilizar las variables declaradas en el procedimiento anterior (declaracion_variables). Podemos reasignar valores y utilizarlas en este m�dulo.

'3. P�blica: visible desde cualquier m�dulo del proyecto.

'Sustituyendo Dim por Public al declarar la variable (ver arriba)

End Sub

Sub declaracion_constantes()

'Se suelen declarar a nivel de m�dulo o p�blicas (ver arriba).
'Es obligatorio declararla y asignarla al mismo tiempo.
'No puede cambiar de valor.
'El scope funciona igual que con las variables.

'Const valor_constante As Byte = 7

End Sub

Sub salario_euros()

'Vamos a hacer un programa que utilice una constante de valor 166,386 (1 euro en pesetas) y convierta un valor en euros a pesetas.
'Declaramos una constante p�blica de tipo Double euro = 166,386 (ver arriba)

Dim salario As Currency

salario = 1984

salario = salario * euro

MsgBox "El salario en pesetas es " & salario

End Sub
