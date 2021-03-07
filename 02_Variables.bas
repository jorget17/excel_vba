Attribute VB_Name = "Variables"
'Obligar la declaración de variables mediante el comando
'Option Explicit

'Dim nombre As String
'Public nombre As String
'Public Const valor_constante As Byte = 7
Public Const euro As Double = 166.386


Sub tipo_datos()

'byte -- números de 0 a 255
'integer -- números enteros -32768 a 32768
'long -- números enteros grandes
'single -- números decimales (parte decimal corta)
'double -- números decimales (parte decimal larga)
'decimal -- números decimales (parte decimal extremadamente larga)
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

'Declaración de variables múltiples
'Dim nombre as string, edad as byte

'Asignar valores múltiples
'nombre = "Marta": edad = 48

'Usar variables en un procedimiento
MsgBox "La usuaria " & nombre & " tiene " & edad & " años."

End Sub

Sub scope_variables()

'Las variables se pueden declarar en diferentes ámbitos (scopes)

'1. Local a nivel de procedimiento: dentro del procedimiento en el que se declara

'Declarando la variable dentro del procedimiento

'2. Local a nivel de módulo: dentro del módulo en el que se declara

'Declarando la variable antes de los procedimientos (ver arriba) e iniciándolas en el momento de su uso (en cualquier procedimiento del módulo).
'También podemos llamar al procedimiento que incluye la declaración de la variable que queremos utilizar.

'Call declaracion_variables
'Esto permite utilizar las variables declaradas en el procedimiento anterior (declaracion_variables). Podemos reasignar valores y utilizarlas en este módulo.

'3. Pública: visible desde cualquier módulo del proyecto.

'Sustituyendo Dim por Public al declarar la variable (ver arriba)

End Sub

Sub declaracion_constantes()

'Se suelen declarar a nivel de módulo o públicas (ver arriba).
'Es obligatorio declararla y asignarla al mismo tiempo.
'No puede cambiar de valor.
'El scope funciona igual que con las variables.

'Const valor_constante As Byte = 7

End Sub

Sub salario_euros()

'Vamos a hacer un programa que utilice una constante de valor 166,386 (1 euro en pesetas) y convierta un valor en euros a pesetas.
'Declaramos una constante pública de tipo Double euro = 166,386 (ver arriba)

Dim salario As Currency

salario = 1984

salario = salario * euro

MsgBox "El salario en pesetas es " & salario

End Sub
