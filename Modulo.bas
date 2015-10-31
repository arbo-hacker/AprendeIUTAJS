Attribute VB_Name = "Modulo"
Public RS As Recordset 'Variable usada para abrir las tablas
Public Sql As Command 'Variable usada junto con Rs para poder utilizar los registros guardados en las tablas
Public Conexion As Connection 'Variable usada para abrir la base de datos
Public azar As Integer  'Variable usada para guardar el número para generar las preguntas al azar
Public Respuesta_correcta As String 'En esta variable se guarda la respuesta correcta a la pregunta generada al azar
Public Preguntas As Integer 'Variable usada para guardar el numero de preguntas que ya se han generado para saber si el usario ya gano
Public Tiempo As Single 'Aqui guardo el tiempo que las etiquetas se van a poner del color verde
Public Cantidad_preguntas As Single 'Aqui guardo el numero de preguntas que se van genrando para saber cuando llegue a 50 y gane el juegador, y para saber de que tabla hay que sacar las preguntas
Public tabla As String 'Aqui guardo cual fue la ultima tabla que se uso al genrar la pregunta
Public Comodin_Fifty As Boolean
Public Comodin_Audiencia As Boolean
Public Comodin As Boolean
Public Nombre_jugador As String
Public Quien As String
Public Cantidades1 As Single, Cantidades2 As Single, Cantidades3 As Single 'Variable usada para almacenar el número de preguntas guardadas en la base de datos
Public Porcentaje(4) As Integer
Public mayor(4) As Integer
Public Pregunta_Actual As String
Public Llamada As Boolean
Public video As String
Public Sub AbrirBD()
Set Sql = New Command
Set Conexion = New Connection
Set RS = New Recordset

On Error GoTo ErrorConexion 'En caso de error hacer lo que dice en la etiqueta ErrorConexion
    'Conexion.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\millonario.mdb" & ";Persist Security Info=False;Jet OLEDB:Database Password=millonario" 'Abro la base de datos
    Conexion.Open "millonario" '"Provider=VFPOLEDB.1;Data Source=" & App.Path & "\data\millonario.dbc;Password='';Collating Sequence=SPANISH"
    Sql.ActiveConnection = Conexion 'Le doy los datos de la conexion a la variable sql
    
ErrorConexion: 'En caso de error hacer lo que dice a continuacion
If Err.Number <> 0 Then 'Si el numero del error es diferente a 0 entonces
    MsgBox "Ha ocurrido el error número " & Err.Number & vbCrLf & Err.Description & vbCrLf & "Consulte con el administrador del programa", vbCritical, "Error" 'Se da un mensaje informando sobre el error
    End 'Se cierra el programa
End If
End Sub
Public Sub Espera(Segundos As Single) ' FUNCION ESPERAR CON SEGUNDOS
  Dim ComienzoSeg As Single
  Dim FinSeg As Single
  ComienzoSeg = Timer
  FinSeg = ComienzoSeg + Segundos
  Do While FinSeg > Timer
      DoEvents
      If ComienzoSeg > Timer Then
          FinSeg = FinSeg - 24 * 60 * 60
      End If
  Loop
End Sub
Public Function Añadir(ParamArray Elementos() As Variant) As String
        If RS.State = 1 Then RS.Close 'Si rs esta abierto, es decir, acaba de ser usado lo cierro para poder usarlo de nuevo
        RS.Open "select * from " & Elementos(0), Conexion, adOpenDynamic, adLockOptimistic 'Abro la tabla que esta escrita en el primer valor del array de elementos
        Dim cadena As String 'Aqui guardo los nombres de todos los campos
        Dim Campo As ADODB.Field 'Esta variable se usa para sacar los nombres de los campos
            For Each Campo In RS.Fields 'Recorrer todos los nombres de los campos en la tabla y darle el nombre a la variable campo
                cadena = cadena & Campo.Name & "," 'Aqui armo los nombres de los campos. Cadena es igual a cadena mas el nombre del campo
            Next 'Cierro el ciclo del for
            
Dim Elem As String 'Aqui construyo los valores que se le daran a los campos
Dim i As Integer 'Declaro la variable que me dara el valor de elementos

        For i = 1 To UBound(Elementos) 'Hacer desde el segundo elemento del array Elementos hasta el ultimo valor de Elementos
            If Mid(Elementos(i), 1, 1) = "#" Then 'Si la primera letra de lo que vale elementos es # entonces es un valor numerico
                If Elem = "" Then 'Si elem es igual a blanco entonces
                    Elem = Mid(Elementos(i), 2, Len(Elementos(i))) & ""  'Elem es igual a lo que hay en elementos a partir de la segunda letra
                Else 'De lo contrario
                    Elem = Elem & "," & Mid(Elementos(i), 2, Len(Elementos(i))) & "" 'Elem es igual a elem mas lo que hay en elementos a partir de la segunda letra
                End If 'Cierro el if
            Else 'De lo contrario elementos es un string
                If Elem = "" Then 'Si elem es igual a blanco entonces
                    Elem = "'" & Elementos(i) & "'" 'Elem es igual a lo que hay en la variable elementos
                Else 'De lo contrario
                    Elem = Elem & ",'" & Elementos(i) & "'" 'Elem es igual a elem mas lo que hay en la variable elementos
                End If 'Cierro el if
            End If 'Cierro el if
        Next 'Cierro el ciclo del for
Añadir = "insert into " & Elementos(0) & "(" & Mid$(cadena, 1, Len(cadena) - 1) & ") values(" & Elem & ")" 'Aqui contruyo la cadena sql
End Function
Public Function ObtieneCodigo(tabla As String)
Dim cuantos As Integer
Sql.CommandText = "select max(cod) autonumerico from " & tabla 'Cuento el numero de registro guardados en la tabla
Set RS = Sql.Execute 'Para saber la cantidad de preguntas que hay
If RS.EOF Then
    cuantos = 0
Else
    cuantos = RS!autonumerico
End If
ObtieneCodigo = cuantos + 1
End Function
Public Function Eliminar(ParamArray Elem() As Variant) As String
    If Mid(Elem(2), 1, 1) = "#" Then
        Elem(2) = Mid(Elem(2), 2, Len(Elem(2)))
    Else
        Elem(2) = "'" & Elem(2) & "'"
    End If
On Error GoTo elem3
        If Elem(3) <> "" Then
            Eliminar = "delete from " & Elem(0) & " where " & Elem(1) & "=" & Elem(2) & " " & Elem(3)
        End If
elem3:
        If Err.Number = 9 Then
            Eliminar = "delete from " & Elem(0) & " where " & Elem(1) & "=" & Elem(2)
        End If
End Function
Public Function Modificar(ParamArray Elementos() As Variant) As String
Dim Elem As String
Dim Campo As Field
Dim i As Integer
Dim Campo_clave As String
If RS.State = 1 Then RS.Close 'Si la tabla esta abierta la cierro
RS.Open "Select * from " & Elementos(0), Conexion, adOpenDynamic 'Abro la tabla cuyo nombre es igual a elementos(0)
    For Each Campo In RS.Fields 'Con este ciclo obtengo los nombres de todos los campos de una tabla
        If i >= 1 Then 'Si i es mayor o igual que 1 entonces
            If Mid(Elementos(i), 1, 1) = "#" Then 'Si la primera letra de lo que vale elementos es # entonces es un valor numerico
                If Elem = "" Then 'Si elem es igual a blanco entonces
                    Elem = Campo.Name & "=" & Mid(Elementos(i), 2, Len(Elementos(i)))  'Elem es igual a lo que hay en elementos a partir de la segunda letra
                Else 'De lo contrario
                    Elem = Elem & ", " & Campo.Name & "=" & Mid(Elementos(i), 2, Len(Elementos(i))) 'Elem es igual a elem mas lo que hay en elementos a partir de la segunda letra
                End If 'Cierro el if
            Else 'De lo contrario elementos es un string
                If Elem = "" Then 'Si elem es igual a blanco entonces
                    Elem = Campo.Name & "='" & Elementos(i) & "'"  'Elem es igual a lo que hay en la variable elementos
                Else 'De lo contrario
                    Elem = Elem & ", " & Campo.Name & "=" & "'" & Elementos(i) & "'" 'Elem es igual a elem mas lo que hay en la variable elementos
                End If 'Cierro el if
            End If 'Cierro el if
        Else 'De lo contrario
        Campo_clave = Campo.Name 'Consigo el nombre del campo clave para poder modificar un registro
        End If 'Cierro el if
        i = i + 1
    Next 'Cierro el ciclo del for
Dim Cual As String 'Determina si el valor del campo_clave es numerico o un srting
            If Mid(Elementos(i), 1, 1) = "#" Then 'Si la primera letra de lo que vale elementos es # entonces es un valor numerico
                Cual = Mid(Elementos(i), 2, Len(Elementos(i))) 'Cual es igual a lo que vale elementos a partir de la segunda letra
            Else 'De lo contrario
                Cual = "'" & Elementos(i) & "'" 'Elementos es un valor string y por eso Cual es igual a lo que vale elementos encerrado entre apostrofes
            End If 'Cierro el If
On Error GoTo Ubound_Elementos 'Si ocurre un error ir a donde dice Ubound_Elementos
                    If Elementos(i + 1) <> "" Then 'Si elementos es diferente de blanco entonces
                        Modificar = "Update " & Elementos(0) & " set " & Elem & " where " & Campo_clave & "=" & Cual & " " & Elementos(i + 1) 'Armo la cadena sql y lo concateno con el ultimo valor de elementos
                    End If 'Cierro el If
Ubound_Elementos: 'Si da error salta hasta aqui
                    If Err.Number = 9 Then 'Si ocurre el error numero 9 entonces
                        Modificar = "Update " & Elementos(0) & " set " & Elem & " where " & Campo_clave & "=" & Cual 'Armo la cadena sql
                    End If 'Cierro el If
            
End Function
Public Function Valores_Preguntas(NroPregunta As Single, Retirarse As Boolean) As String
Select Case NroPregunta
    Case 1 'Preguntas Easy
        Valores_Preguntas = "0,15"
    Case 2
        Valores_Preguntas = "0,3"
    Case 3
        Valores_Preguntas = "0,45"
    Case 4
        Valores_Preguntas = "0,6"
    Case 5
        Valores_Preguntas = "0,75"
    Case 6
        Valores_Preguntas = "0,9"
    Case 7
        Valores_Preguntas = "1,05"
    Case 8
        Valores_Preguntas = "1,2"
    Case 9
        Valores_Preguntas = "1,35"
    Case 10
        Valores_Preguntas = "1,5"
    Case 11
        Valores_Preguntas = "1,65"
    Case 12
        Valores_Preguntas = "1,8"
    Case 13
        Valores_Preguntas = "1,95"
    Case 14
        Valores_Preguntas = "2,1"
    Case 15
        Valores_Preguntas = "2,25"
    Case 16
        Valores_Preguntas = "2,4"
    Case 17
        Valores_Preguntas = "2,55"
    Case 18
        Valores_Preguntas = "2,7"
    Case 19
        Valores_Preguntas = "2,85"
    Case 20
        Valores_Preguntas = "3"
    Case 21 'Preguntas Medium
        Valores_Preguntas = "3,4"
    Case 22
        Valores_Preguntas = "3,8"
    Case 23
        Valores_Preguntas = "4,2"
    Case 24
        Valores_Preguntas = "4,6"
    Case 25
        Valores_Preguntas = "5"
    Case 26
        Valores_Preguntas = "5,4"
    Case 27
        Valores_Preguntas = "5,8"
    Case 28
        Valores_Preguntas = "6,2"
    Case 29
        Valores_Preguntas = "6,6"
    Case 30
        Valores_Preguntas = "7"
    Case 31
        Valores_Preguntas = "7,4"
    Case 32
        Valores_Preguntas = "7,8"
    Case 33
        Valores_Preguntas = "8,2"
    Case 34
        Valores_Preguntas = "8,6"
    Case 35
        Valores_Preguntas = "9"
    Case 36
        Valores_Preguntas = "9,4"
    Case 37
        Valores_Preguntas = "9,8"
    Case 38
        Valores_Preguntas = "10,2"
    Case 39
        Valores_Preguntas = "10,6"
    Case 40
        Valores_Preguntas = "11"
    Case 41 'Preguntas Hard
        Valores_Preguntas = "11,5"
    Case 42
        Valores_Preguntas = "12"
    Case 43
        Valores_Preguntas = "13"
    Case 44
        Valores_Preguntas = "14"
    Case 45
        Valores_Preguntas = "15"
    Case 46
        Valores_Preguntas = "16"
    Case 47
        Valores_Preguntas = "17"
    Case 48
        Valores_Preguntas = "18"
    Case 49
        Valores_Preguntas = "19"
    Case 50
        Valores_Preguntas = "20"
End Select
If NroPregunta = 0 Then
    Valores_Preguntas = "0"
    Exit Function
End If
If Retirarse = False Then
    Select Case NroPregunta
        Case 37 To 38
            Valores_Preguntas = "10"
        Case 39 To 40
            Valores_Preguntas = "11"
        Case 41
            Valores_Preguntas = "12"
        Case 42 To 50
            'Valores_Preguntas = "300000"
        Case Else
            Valores_Preguntas = "0"
    End Select
End If
End Function




