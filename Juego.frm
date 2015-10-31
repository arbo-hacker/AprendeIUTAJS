VERSION 5.00
Object = "{F5BE8BC2-7DE6-11D0-91FE-00C04FD701A5}#2.0#0"; "agentctl.dll"
Begin VB.Form FrmJuego 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ventana principal"
   ClientHeight    =   7185
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13875
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MouseIcon       =   "Juego.frx":0000
   MousePointer    =   99  'Custom
   Moveable        =   0   'False
   Picture         =   "Juego.frx":0CCA
   ScaleHeight     =   7185
   ScaleWidth      =   13875
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer TmResCorrecta 
      Enabled         =   0   'False
      Interval        =   150
      Left            =   120
      Top             =   120
   End
   Begin VB.Image ImgInfPreguntas 
      Height          =   735
      Left            =   9960
      MouseIcon       =   "Juego.frx":4A325
      MousePointer    =   99  'Custom
      Top             =   2280
      Width           =   855
   End
   Begin VB.Image ImgLlamarAmigo 
      Height          =   495
      Left            =   9840
      MouseIcon       =   "Juego.frx":4ABEF
      MousePointer    =   99  'Custom
      Top             =   1200
      Width           =   495
   End
   Begin VB.Image ImgAudiencia 
      Height          =   495
      Left            =   3000
      MouseIcon       =   "Juego.frx":4B4B9
      MousePointer    =   99  'Custom
      Top             =   3000
      Width           =   615
   End
   Begin VB.Image ImgFifty 
      Height          =   495
      Left            =   6600
      MouseIcon       =   "Juego.frx":4BD83
      MousePointer    =   99  'Custom
      ToolTipText     =   "Ayuda a mis respuestas"
      Top             =   1080
      Width           =   615
   End
   Begin AgentObjectsCtl.Agent Agent1 
      Left            =   120
      Top             =   840
   End
   Begin VB.Label LblRespuestaD 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   975
      Left            =   10200
      MouseIcon       =   "Juego.frx":4C64D
      MousePointer    =   99  'Custom
      TabIndex        =   4
      Top             =   5280
      Width           =   3135
   End
   Begin VB.Label LblRespuestaC 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   855
      Left            =   2640
      MouseIcon       =   "Juego.frx":4CF17
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   5280
      Width           =   2655
   End
   Begin VB.Label LblRespuestaB 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   735
      Left            =   9240
      MouseIcon       =   "Juego.frx":4D7E1
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   4080
      Width           =   1815
   End
   Begin VB.Label LblRespuestaA 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   735
      Left            =   3960
      MouseIcon       =   "Juego.frx":4E0AB
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   4080
      Width           =   1935
   End
   Begin VB.Label LblPregunta 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1575
      Left            =   4320
      TabIndex        =   0
      Tag             =   "algo"
      Top             =   1680
      Width           =   4935
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "D."
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   495
      Left            =   9960
      MouseIcon       =   "Juego.frx":4E975
      MousePointer    =   99  'Custom
      TabIndex        =   8
      Top             =   5280
      Width           =   735
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "B."
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   9000
      MouseIcon       =   "Juego.frx":4F23F
      MousePointer    =   99  'Custom
      TabIndex        =   7
      Top             =   4080
      Width           =   495
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "A."
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   3720
      MouseIcon       =   "Juego.frx":4FB09
      MousePointer    =   99  'Custom
      TabIndex        =   6
      Top             =   4080
      Width           =   495
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "C."
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   2400
      MouseIcon       =   "Juego.frx":503D3
      MousePointer    =   99  'Custom
      TabIndex        =   5
      Top             =   5280
      Width           =   375
   End
End
Attribute VB_Name = "FrmJuego"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Form_Load()
'Top_left Quien
Cantidad_preguntas = 1
Randomize 'Llamo esta funcion para que se generen numeros al azar
Call AbrirBD 'Habro la base de datos
Call Cantidades
Agent1.Characters.Load "james", App.Path & "\config.dat" 'Busco la ruta del agente
Agent1.Characters("james").Show 'Cargo el agente para que aparezcca
Agent1.Characters("james").Top = 400 '150 'Acomodo al agente en
Agent1.Characters("james").Left = 600 '170 'el lugar mas indicado
Call LiberaPreguntas
Call Generar_Pregunta 'Genero las preguntas y respuestas en el formulario
End Sub
Private Function IIFResultado(valor As Integer)
IIFResultado = IIf(valor < 0, 0, valor)
End Function
Private Function AudienciaRespuestaCorrecta(porcentaje1 As Integer, porcentaje2 As Integer, porcentaje3 As Integer, porcentaje4 As Integer) As Integer

AudienciaRespuestaCorrecta = 1
End Function
Private Sub Audiencia()
Dim IntAzar As Integer, IntPorcentaje As Single, IntRespuestaCorrecta As Single, _
Labels As Control, IntOtraRespuesta As Single, i As Integer

Dim porcentaje1 As Integer, porcentaje2 As Integer, porcentaje3 As Integer, porcentaje4 As Integer

On Error GoTo Repetir

porcentaje1 = Int(Rnd(Time) * 100) + 1
porcentaje2 = IIFResultado(Int(Rnd(Time) * (100 - (porcentaje1))))
porcentaje3 = IIFResultado(Int(Rnd(Time) * (100 - ((porcentaje1) + porcentaje2))))
porcentaje4 = IIFResultado(100 - porcentaje3 - porcentaje2 - porcentaje1)



If Comodin_Audiencia = False Then
    If Comodin = True Then
        IntRespuestaCorrecta = IIFResultado(Int(Rnd() * 20) + 65)
        IntPorcentaje = IIFResultado(100 - IntRespuestaCorrecta)
        IntAzar = Int(Rnd() * 20) + 1
        
        Select Case Respuesta_correcta
            Case LblRespuestaA.Caption
                LblRespuestaA.Tag = IntRespuestaCorrecta
                Porcentaje(1) = IntRespuestaCorrecta
            Case LblRespuestaB.Caption
                LblRespuestaB.Tag = IntRespuestaCorrecta
                Porcentaje(2) = IntRespuestaCorrecta
            Case LblRespuestaC.Caption
                LblRespuestaC.Tag = IntRespuestaCorrecta
                Porcentaje(3) = IntRespuestaCorrecta
            Case LblRespuestaD.Caption
                LblRespuestaD.Tag = IntRespuestaCorrecta
                Porcentaje(4) = IntRespuestaCorrecta
        End Select
        Dim Porcen As Integer
        For Each Labels In Me
            If UCase(TypeName(Labels)) = "LABEL" Then
                If Labels.Visible = True Then
                    If Labels.Tag = "" Then
                        Labels.Tag = IntPorcentaje
                        Select Case Labels.Name
                            Case LblRespuestaA.Name
                                Porcentaje(1) = IntPorcentaje
                            Case LblRespuestaA.Name
                                Porcentaje(2) = IntPorcentaje
                            Case LblRespuestaA.Name
                                Porcentaje(3) = IntPorcentaje
                            Case LblRespuestaA.Name
                                Porcentaje(4) = IntPorcentaje
                        End Select
                    End If
                End If
            End If
        Next

        Select Case IntAzar
            Case 5 To 10
                For i = 1 To 4
                    If Porcentaje(i) <> 0 Then
                        Porcentaje(i) = 50
                    End If
                Next
        End Select
        LblRespuestaA.Tag = ""
        LblRespuestaB.Tag = ""
        LblRespuestaC.Tag = ""
        LblRespuestaD.Tag = ""
        Comodin_Audiencia = True
        FrmAudiencia.Show 1
    Else
    
        IntRespuestaCorrecta = IIFResultado(Int(Rnd() * 40) + 36)
        IntPorcentaje = IIFResultado(Int(Rnd() * (100 - (IntRespuestaCorrecta - 5))) + 1)
        IntOtraRespuesta = IIFResultado(Int(Rnd() * (100 - ((IntRespuestaCorrecta) + IntPorcentaje))) + 2)
        IntAzar = Int(Rnd() * 20) + 1
        
        Select Case Respuesta_correcta
            Case LblRespuestaA.Caption
                LblRespuestaA.Tag = IntRespuestaCorrecta
            Case LblRespuestaB.Caption
                LblRespuestaB.Tag = IntRespuestaCorrecta
            Case LblRespuestaC.Caption
                LblRespuestaC.Tag = IntRespuestaCorrecta
            Case LblRespuestaD.Caption
                LblRespuestaD.Tag = IntRespuestaCorrecta
        End Select
        
        For Each Labels In Me
            If UCase(TypeName(Labels)) = "LABEL" Then
                If Labels.Tag = "" Then
                    i = i + 1
                    Select Case i
                        Case 1
                            Labels.Tag = IntPorcentaje
                        Case 2
                            Labels.Tag = IntOtraRespuesta
                        Case 3
                            Labels.Tag = IIFResultado(100 - (IntRespuestaCorrecta + IntPorcentaje + IntOtraRespuesta))
                    End Select
                End If
            End If
        Next
        If (LblRespuestaA.Tag = LblRespuestaB.Tag) Or (LblRespuestaA.Tag = LblRespuestaC.Tag) Or (LblRespuestaA.Tag = LblRespuestaD.Tag) Or (LblRespuestaB.Tag = LblRespuestaC.Tag) Or (LblRespuestaB.Tag = LblRespuestaD.Tag) Or (LblRespuestaC.Tag = LblRespuestaD.Tag) Then
            Call Audiencia
            LblRespuestaA.Tag = ""
            LblRespuestaB.Tag = ""
            LblRespuestaC.Tag = ""
            LblRespuestaD.Tag = ""
            Exit Sub
        End If
        Select Case IntAzar
            Case 11 To 20
                Porcentaje(1) = LblRespuestaA.Tag
                Porcentaje(2) = LblRespuestaB.Tag
                Porcentaje(3) = LblRespuestaC.Tag
                Porcentaje(4) = LblRespuestaD.Tag
            Case 5 To 6
                Porcentaje(1) = 25
                Porcentaje(2) = 25
                Porcentaje(3) = 25
                Porcentaje(4) = 25
            Case Else
                Call Porcentaje_Erroneo
        End Select
        LblRespuestaA.Tag = ""
        LblRespuestaB.Tag = ""
        LblRespuestaC.Tag = ""
        LblRespuestaD.Tag = ""
        FrmAudiencia.Show 1
        Comodin_Audiencia = True
    End If
End If
Repetir:
If Err.Number <> 0 Then
    DoEvents
    Call Audiencia
End If
End Sub
Private Sub Porcentaje_Erroneo()
Porcentaje(1) = Int(Rnd() * 30) + 1
Porcentaje(2) = Int(Rnd() * 35) + 1
Porcentaje(3) = Int(Rnd() * 28) + 1
Porcentaje(4) = 100 - (Porcentaje(2) + Porcentaje(1) + Porcentaje(3))
If (Porcentaje(1) = Porcentaje(2)) Or (Porcentaje(1) = Porcentaje(3)) Or (Porcentaje(1) = Porcentaje(4)) Or (Porcentaje(2) = Porcentaje(3)) Or (Porcentaje(2) = Porcentaje(4)) Or (Porcentaje(3) = Porcentaje(4)) Then
    Call Porcentaje_Erroneo
    Exit Sub
End If
End Sub

Private Sub Cantidades()
Sql.CommandText = "select count(cod) as cantidad from Preguntas" 'Cuento el numero de registro guardados en la tabla
Set RS = Sql.Execute 'Para saber la cantidad de preguntas que hay
Cantidades1 = RS!Cantidad 'Y luego guardo este dato en la variable cantidad
Sql.CommandText = "select count(cod) as cantidad from Preguntas2" 'Cuento el numero de registro guardados en la tabla
Set RS = Sql.Execute 'Para saber la cantidad de preguntas que hay
Cantidades2 = RS!Cantidad 'Y luego guardo este dato en la variable cantidad
Sql.CommandText = "select count(cod) as cantidad from Preguntas3" 'Cuento el numero de registro guardados en la tabla
Set RS = Sql.Execute 'Para saber la cantidad de preguntas que hay
Cantidades3 = RS!Cantidad 'Y luego guardo este dato en la variable cantidad
End Sub
Private Sub Top_left(jugador As String)
Dim y1 As Single, y2 As Single, y3 As Single, y4 As Single, y5 As Single, y6 As Single, y7 As Single, y8 As Single

Select Case jugador
    Case "cacheton"
        y1 = 5600
        y2 = 5600
        y3 = 5600
        y4 = 5580
        y5 = 6210
        y6 = 6220
        y7 = 6210
        y8 = 6210
    Case "frenton"
        y1 = 5570
        y2 = 5570
        y3 = 5560
        y4 = 5560
        y5 = 6180
        y6 = 6190
        y7 = 6180
        y8 = 6180
    Case "negra"
        y1 = 5580
        y2 = 5580
        y3 = 5580
        y4 = 5560
        y5 = 6200
        y6 = 6210
        y7 = 6210
        y8 = 6200
    Case "narizon"
        y1 = 5610
        y2 = 5610
        y3 = 5610
        y4 = 5600
        y5 = 6220
        y6 = 6240
        y7 = 6220
        y8 = 6220
End Select

Av.Top = y1
An.Top = y2
Bv.Top = y3
Bn.Top = y4
Cv.Top = y5
Cn.Top = y6
Dv.Top = y7
Dn.Top = y8
End Sub
Private Sub LiberaPreguntas()
Sql.CommandText = "update preguntas set lista='N'" 'Con esta cadena sql actualizo todos
Set RS = Sql.Execute 'los registros del campo lista para que obtengan el valor "N"
Sql.CommandText = "update preguntas2 set lista='N'" 'Con esta cadena sql actualizo todos
Set RS = Sql.Execute 'los registros del campo lista para que obtengan el valor "N"
Sql.CommandText = "update preguntas3 set lista='N'" 'Con esta cadena sql actualizo todos
Set RS = Sql.Execute 'los registros del campo lista para que obtengan el valor "N"
End Sub
Private Sub Generar_Pregunta()
    azar = 0
        If Cantidad_preguntas <= 20 Then
            Do While azar > Cantidades1 Or azar = 0 'Hacer mientras la variable azar sea mayor que la variable cantidad o mientras sea igual a 0
                azar = Int(Rnd() * (Cantidades1 + 1)) 'Genera un numero del 1 al 100 al azar
                If azar <= Cantidades1 And azar >= 1 Then 'Si la variabl azar es mayor o igual a lo que hay en la variable cantidad y si la variable azar es mayor  o igual a 1 entonces
                    Sql.CommandText = "Select * from Preguntas where cod=" & azar  'Busco en la base de datos
                    Set RS = Sql.Execute 'la pregunta cuyo codigo sera el valor de la variable azar
                    DoEvents
                    If RS.EOF = True Then
                        Call Generar_Pregunta 'llamo a la funcion generar preguna
                        Exit Sub 'y salgo de este procedimiento
                    End If
                    If RS!lista = "S" Then 'Si el campo lista es igual a S entonces
                        Call Generar_Pregunta 'llamo a la funcion generar preguna
                        Exit Sub 'y salgo de este procedimiento
                    End If
                        Sql.CommandText = "update preguntas set lista='S' where cod=" & azar 'Con esta funcion cada vez que se genera una pregunta se modifica el campo lista y se escribe S
                        Set RS = Sql.Execute 'Para que la pregunta no se vuelva a generar, es decir, que no se repita
                        Sql.CommandText = "Select * from Preguntas where cod=" & azar 'Busco en la base de datos
                        Set RS = Sql.Execute 'la pregunta cuyo codigo sera l valor de la variable azar
                        LblPregunta.Caption = RS!Preguntas 'Aqui colocare la pregunta al azar
                        LblRespuestaA.Caption = RS!respuesta1 '
                        LblRespuestaB.Caption = RS!respuesta2 'Aqui colocare las respuestas 1,2
                        LblRespuestaC.Caption = RS!respuesta3 '3 y 4 respectivamente; para que
                        LblRespuestaD.Caption = RS!respuesta4 'el usario seleccione la correcta
                        Respuesta_correcta = RS!Correcta 'Aqui guardo cual es la respuesta correcta
                        Cantidad_preguntas = Cantidad_preguntas + 1
                        Agent1.Characters("james").Speak LblPregunta.Caption 'Le digo al agente que diga la pregunta
                End If
            Loop
        ElseIf Cantidad_preguntas >= 21 And Cantidad_preguntas < 41 Then
            Do While azar > Cantidades2 Or azar = 0 'Hacer mientras la variable azar sea mayor que la variable cantidad o mientras sea igual a 0
                azar = Int(Rnd() * (Cantidades2 + 1)) 'Genera un numero del 1 al 100 al azar
                If azar <= Cantidades2 And azar >= 1 Then 'Si la variabl azar es mayor o igual a lo que hay en la variable cantidad y si la variable azar es mayor  o igual a 1 entonces
                    Sql.CommandText = "Select * from Preguntas2 where cod=" & azar  'Busco en la base de datos
                    Set RS = Sql.Execute 'la pregunta cuyo codigo sera el valor de la variable azar
                    DoEvents
                    If RS!lista = "S" Then 'Si el campo lista es igual a S entonces
                        Call Generar_Pregunta 'llamo a la funcion generar preguna
                        Exit Sub 'y salgo de este procedimiento
                    End If
                        Sql.CommandText = "update preguntas2 set lista='S' where cod=" & azar 'Con esta funcion cada vez que se genera una pregunta se modifica el campo lista y se escribe S
                        Set RS = Sql.Execute 'Para que la pregunta no se vuelva a generar, es decir, que no se repita
                        Sql.CommandText = "Select * from Preguntas2 where cod=" & azar 'Busco en la base de datos
                        Set RS = Sql.Execute 'la pregunta cuyo codigo sera l valor de la variable azar
                        LblPregunta.Caption = RS!Preguntas 'Aqui colocare la pregunta al azar
                        LblRespuestaA.Caption = RS!respuesta1 '
                        LblRespuestaB.Caption = RS!respuesta2 'Aqui colocare las respuestas 1,2
                        LblRespuestaC.Caption = RS!respuesta3 '3 y 4 respectivamente; para que
                        LblRespuestaD.Caption = RS!respuesta4 'el usario seleccione la correcta
                        Respuesta_correcta = RS!Correcta 'Aqui guardo cual es la respuesta correcta
                        Cantidad_preguntas = Cantidad_preguntas + 1
                        Agent1.Characters("james").Speak LblPregunta.Caption 'Le digo al agente que diga la pregunta
                End If
            Loop
        ElseIf Cantidad_preguntas >= 41 And Cantidad_preguntas < 51 Then
            Do While azar > Cantidades3 Or azar = 0 'Hacer mientras la variable azar sea mayor que la variable cantidad o mientras sea igual a 0
                azar = Int(Rnd() * (Cantidades3 + 1)) 'Genera un numero del 1 al 100 al azar
                If azar <= Cantidades3 And azar >= 1 Then 'Si la variabl azar es mayor o igual a lo que hay en la variable cantidad y si la variable azar es mayor  o igual a 1 entonces
                    Sql.CommandText = "Select * from Preguntas3 where cod=" & azar  'Busco en la base de datos
                    Set RS = Sql.Execute 'la pregunta cuyo codigo sera el valor de la variable azar
                    DoEvents
                    If RS!lista = "S" Then 'Si el campo lista es igual a S entonces
                        Call Generar_Pregunta 'llamo a la funcion generar preguna
                        Exit Sub 'y salgo de este procedimiento
                    End If
                        Sql.CommandText = "update preguntas3 set lista='S' where cod=" & azar 'Con esta funcion cada vez que se genera una pregunta se modifica el campo lista y se escribe S
                        Set RS = Sql.Execute 'Para que la pregunta no se vuelva a generar, es decir, que no se repita
                        Sql.CommandText = "Select * from Preguntas3 where cod=" & azar 'Busco en la base de datos
                        Set RS = Sql.Execute 'la pregunta cuyo codigo sera l valor de la variable azar
                        LblPregunta.Caption = RS!Preguntas 'Aqui colocare la pregunta al azar
                        LblRespuestaA.Caption = RS!respuesta1 '
                        LblRespuestaB.Caption = RS!respuesta2 'Aqui colocare las respuestas 1,2
                        LblRespuestaC.Caption = RS!respuesta3 '3 y 4 respectivamente; para que
                        LblRespuestaD.Caption = RS!respuesta4 'el usario seleccione la correcta
                        Respuesta_correcta = RS!Correcta 'Aqui guardo cual es la respuesta correcta
                        Cantidad_preguntas = Cantidad_preguntas + 1
                        Agent1.Characters("james").Speak LblPregunta.Caption 'Le digo al agente que diga la pregunta
                End If
            Loop
        ElseIf Cantidad_preguntas = 51 Then
            Dim Valor_Pregunta  As String
            Valor_Pregunta = Valores_Preguntas(Cantidad_preguntas - 1, False)
            Sql.CommandText = Añadir("records", Nombre_jugador, "#" & Valor_Pregunta)
            Set RS = Sql.Execute
            Unload Me
            Load FrmFinal
            FrmFinal.Show
        End If
    'Tabla = Trim(Mid(Sql.CommandText, 14, 11))
    Call Mostrar_lbl
    Comodin = False
    Pregunta_Actual = Cantidad_preguntas - 1
End Sub
Private Sub Verde_invisible()
'Av.Visible = False '
'Bv.Visible = False 'Oculto los fondos verdes
'Cv.Visible = False 'de las respuestas
'Dv.Visible = False '
End Sub
Private Sub Correcta(respuesta As String)
Dim Valor_Pregunta As String
If respuesta = Respuesta_correcta Then 'Si la variable respuesta es igual a la respuesta correcta entonces
    'TmResCorrecta.Enabled = True 'Activo el reloj
    Call Agente("Correcto")
    Call Espera(Tiempo)  'Espero dos segundos
    'TmResCorrecta.Enabled = False 'Desactivo el reloj
    Call Verde_invisible 'Desaparezco el fondo verde
    'Call Agente("Pregunta")
    'Espera Tiempo
    Call Generar_Pregunta 'Genero una nueva pregunta
Else 'De lo contrario
    Valor_Pregunta = Valores_Preguntas(Cantidad_preguntas - 2, False)
    If Valor_Pregunta <> "" And Valor_Pregunta <> "0" Then
        Sql.CommandText = Añadir("records", Nombre_jugador, "#" & Valor_Pregunta)
        Set RS = Sql.Execute
    End If
    'TmResCorrecta.Enabled = True 'Activo el reloj
    Call Agente("Incorrecto")
    Agent1.Characters("james").Hide
    Call Espera(Tiempo) 'Espero dos segundos
    'TmResCorrecta.Enabled = False 'Desactivo el reloj
    Call Verde_invisible 'Desaparezco el fondo verde
    Unload Me 'Cierro el formulario
    Load FrmPerdiste 'Cargo el evento load del formulario
    FrmPerdiste.Show 'Hago que aparezca el formulario
End If
End Sub
Private Function Agente(Tipo_respuesta As String) As String
Dim Suerte As Integer 'Variable que se usa para generar un número al azar
Suerte = Int(Rnd() * 10) 'Aqui se genera el número al azar
    Select Case Tipo_respuesta 'Si la variable tipo_respuuesta
        Case "Correcto" 'Es igual a Correcto
            Select Case Suerte '
                Case 1 'Con este codigo
                    Agente = "Muy Bien" ' lo que hago es preguntar el valor
                Case 2 'al azar que va a tomar la
                    Agente = "Correcto" 'variable suerte si obtiene
                Case 3 'el numero 1, 2, 3,4 o 9 entonces guardo el
                    Agente = "Eres muy buen competidor" 'comentario
                Case 4 'en la funcion, es decir, en Agente
                    Agente = "Te las sabes todas más 1" '
                Case 9 '
                    Agente = "Nunca antes se había visto esto" '
            End Select '
        Case "Incorrecto" 'Es igual a incorrecto
            Select Case Suerte '
                Case 1 'Con este codigo
                    Agente = "Haz sido un excelente jugador" 'lo que hago es
                Case 2 'preguntar el valor al azar que va a tomar la
                    Agente = "Suerte la próxima vez" 'variable suerte si obtiene
                Case 3 'el numero 1, 2, 3,4 o 9 entonces guardo el
                    Agente = "No debes ser mal perdedor" 'comentario
                Case 4 'en la funcion, es decir, en Agente
                    Agente = "Debes analizar mejor la pregunta" '
                Case 9 '
                    Agente = "No vuelvas a jugar, no vale la pena" '
            End Select '
        Case "Pregunta"
            Agente = "Tienes " & Valores_Preguntas(Cantidad_preguntas, True) & " de bolívares"
    End Select
    If Agente <> "" Then 'Si la funcion agente tiene algun valor guardado entonces
        Agent1.Characters("james").Speak Agente 'Le digo al agente james que diga el comentario
        'Tiempo = 2
        If Tipo_respuesta = "Incorrecto" Then
            Tiempo = 3
        Else
            Tiempo = 2
        End If
    Else
        Tiempo = 2
    End If
End Function
Private Sub Salir()
Comodin_Fifty = False
Comodin_Audiencia = False
Agent1.Characters.Unload "james" 'Descargo el agente james
End Sub
Private Sub Numero_preguntas()
Sql.CommandText = "select count(cod) as cantidad from Preguntas" 'Cuento el numero de registro guardados en la tabla
Set RS = Sql.Execute 'Para saber la cantidad de preguntas que hay
Cantidad = RS!Cantidad 'Y luego guardo este dato en la variable cantidad
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
'An.Visible = False 'Cuando el mouse este situado
'Bn.Visible = False 'en el formulario y no en las respuestas
'Cn.Visible = False 'el color de la respuestas
'Dn.Visible = False 'sera naranja
End Sub
Private Sub Form_Unload(Cancel As Integer)
Call Salir 'llamo este procedimiento para que las preguntas y respuestas puedan volver a generarse
End Sub
Private Function Ocultar_lbl()
LblRespuestaA.Visible = False
LblRespuestaB.Visible = False
LblRespuestaC.Visible = False
LblRespuestaD.Visible = False
End Function
Private Function Mostrar_lbl()
    LblRespuestaA.Visible = True
    LblRespuestaB.Visible = True
    LblRespuestaC.Visible = True
    LblRespuestaD.Visible = True
End Function
Private Function Fifty()
Dim ctrl As Control
Dim var As String
If Comodin_Fifty = False Then
    Call Ocultar_lbl
    Select Case Respuesta_correcta
        Case LblRespuestaA.Caption
            LblRespuestaA.Visible = True
        Case LblRespuestaB
            LblRespuestaB.Visible = True
        Case LblRespuestaC
            LblRespuestaC.Visible = True
        Case LblRespuestaD
            LblRespuestaD.Visible = True
    End Select
    
    For Each ctrl In Me
        If UCase(TypeName(ctrl)) = "LABEL" Then
            If ctrl.Visible = False Then
                var = ctrl.Name
            End If
        End If
    Next
    
    Select Case var
        Case LblRespuestaA.Name
            LblRespuestaA.Visible = True
        Case LblRespuestaB.Name
            LblRespuestaB.Visible = True
        Case LblRespuestaC.Name
            LblRespuestaC.Visible = True
        Case LblRespuestaD.Name
            LblRespuestaD.Visible = True
    End Select
    Comodin_Fifty = True
    Comodin = True
End If
End Function

Private Sub Image1_Click()

End Sub

Private Sub ImgAudiencia_Click()
Call Audiencia
End Sub

Private Sub ImgFifty_Click()
Call Fifty
End Sub

Private Sub ImgInfPreguntas_Click()
Agent1.Characters("james").Top = 2222
Agent1.Characters("james").Left = 2222
FrmInfPreguntas.Show 1
Agent1.Characters("james").Top = 400 'Acomodo al agente en
Agent1.Characters("james").Left = 600 'el lugar mas indicado
End Sub

Private Sub ImgLlamarAmigo_Click()
If Llamada = False Then
    Agent1.Characters("james").Top = 2222
    Agent1.Characters("james").Left = 2222
    FrmLlamarAmigo.Show 1
    Agent1.Characters("james").Top = 400 'Acomodo al agente en
    Agent1.Characters("james").Left = 600 'el lugar mas indicado
    Llamada = True
End If
End Sub

Private Sub Label1_Click()
LblRespuestaC_Click
End Sub

Private Sub Label2_Click()
LblRespuestaA_Click
End Sub

Private Sub Label3_Click()
LblRespuestaB_Click
End Sub

Private Sub Label4_Click()
LblRespuestaD_Click
End Sub

Private Sub LblPregunta_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
'An.Visible = False 'Cuando el mouse este situado
'Bn.Visible = False 'en el formulario y no en las respuestas
'Cn.Visible = False 'el color de la respuestas
'Dn.Visible = False 'sera naranja
End Sub
Private Sub LblRespuestaA_Click()
Call Correcta(LblRespuestaA.Caption) 'Llamo al procedimiento para saber si la respuesta A es la correcta
End Sub
Private Sub LblRespuestaA_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
'An.Visible = True 'Cuando el mouse este encima de la respuesta A el fondo se pondra anaranjado
'Bn.Visible = False '
'Cn.Visible = False 'Oculto los otros fondos que esten del color naranja
'Dn.Visible = False '
End Sub
Private Sub LblRespuestaB_Click()
Call Correcta(LblRespuestaB.Caption) 'Llamo al procedimiento para saber si la respuesta B es la correcta
End Sub
Private Sub LblRespuestaB_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
'An.Visible = False '
'Bn.Visible = True 'Cuando el mouse este encima de la respuesta B el fondo se pondra anaranjado
'Cn.Visible = False 'Oculto los otros fondos que esten del color naranja
'Dn.Visible = False '
End Sub
Private Sub LblRespuestaC_Click()
Call Correcta(LblRespuestaC.Caption) 'Llamo al procedimiento para saber si la respuesta C es la correcta
End Sub
Private Sub LblRespuestac_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
'An.Visible = False '
'Bn.Visible = False 'Oculto los otros fondos que esten del color naranja
'Cn.Visible = True 'Cuando el mouse este encima de la respuesta C el fondo
'Dn.Visible = False 'se pondra anaranjado
End Sub
Private Sub LblRespuestaD_Click()
Call Correcta(LblRespuestaD.Caption) 'Llamo al procedimiento para saber si la respuesta D es la correcta
End Sub
Private Sub LblRespuestad_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
'An.Visible = False '
'Bn.Visible = False 'Oculto los otros fondos que esten del color naranja
'Cn.Visible = False '
'Dn.Visible = True 'Cuando el mouse este encima de la respuesta D el fondo se pondra anaranjado
End Sub
Private Sub TmResCorrecta_Timer()
If LblRespuestaA.Caption = Respuesta_correcta Then 'Si la respuesta A es la respuesta correcta entonces
    If Av.Visible = False Then 'Si no se ve el fondo verde entonces
        Av.Visible = True 'La respuesta tiene el fondo verde
    Else 'De lo contrario
        Av.Visible = False 'La respuesta tiene su fondo normal (El negro)
    End If
End If
    If LblRespuestaB.Caption = Respuesta_correcta Then 'Si la respuesta B es la respuesta correcta entonces
        If Bv.Visible = False Then 'Si no se ve el fondo verde entonces
            Bv.Visible = True 'La respuesta tiene el fondo verde
        Else 'De lo contrario
            Bv.Visible = False 'La respuesta tiene su fondo normal (El negro)
        End If
    End If
        If LblRespuestaC.Caption = Respuesta_correcta Then 'Si la respuesta C es la respuesta correcta entonces
            If Cv.Visible = False Then 'Si no se ve el fondo verde entonces
                Cv.Visible = True 'La respuesta tiene el fondo verde
            Else 'De lo contrario
                Cv.Visible = False 'La respuesta tiene su fondo normal (El negro)
            End If
        End If
            If LblRespuestaD.Caption = Respuesta_correcta Then 'Si la respuesta D es la respuesta correcta entonces
                If Dv.Visible = False Then 'Si no se ve el fondo verde entonces
                    Dv.Visible = True 'La respuesta tiene el fondo verde
                Else 'De lo contrario
                    Dv.Visible = False 'La respuesta tiene su fondo normal (El negro)
                End If
            End If
End Sub
