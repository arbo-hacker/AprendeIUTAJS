VERSION 5.00
Object = "{F5BE8BC2-7DE6-11D0-91FE-00C04FD701A5}#2.0#0"; "agentctl.dll"
Begin VB.Form FrmLlamarAmigo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Llamar a un amigo"
   ClientHeight    =   8970
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13920
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MouseIcon       =   "LlamarAmigo.frx":0000
   MousePointer    =   99  'Custom
   Picture         =   "LlamarAmigo.frx":0CCA
   ScaleHeight     =   8970
   ScaleWidth      =   13920
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer TmTiempo 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   0
      Top             =   840
   End
   Begin VB.Timer TmRespuesta 
      Enabled         =   0   'False
      Interval        =   1500
      Left            =   0
      Top             =   480
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   11880
      MouseIcon       =   "LlamarAmigo.frx":592E3
      MousePointer    =   99  'Custom
      Top             =   6000
      Width           =   975
   End
   Begin VB.Label LblTiempo 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   10320
      TabIndex        =   1
      Top             =   5760
      Width           =   615
   End
   Begin VB.Label LblRespuesta 
      BackStyle       =   0  'Transparent
      Caption         =   "La respuesta correcta es la opción: "
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   2295
      Left            =   360
      TabIndex        =   0
      Top             =   3960
      Visible         =   0   'False
      Width           =   3615
   End
   Begin AgentObjectsCtl.Agent Agent1 
      Left            =   0
      Top             =   0
   End
   Begin VB.Image ImgWoman 
      Height          =   1095
      Left            =   7080
      MouseIcon       =   "LlamarAmigo.frx":59BAD
      MousePointer    =   99  'Custom
      Top             =   4800
      Width           =   855
   End
   Begin VB.Image ImgMan 
      Height          =   1095
      Left            =   7920
      MouseIcon       =   "LlamarAmigo.frx":5A477
      MousePointer    =   99  'Custom
      Top             =   4800
      Width           =   735
   End
   Begin VB.Image ImgOzzar 
      Height          =   1095
      Left            =   5640
      MouseIcon       =   "LlamarAmigo.frx":5AD41
      MousePointer    =   99  'Custom
      Top             =   4680
      Width           =   735
   End
   Begin VB.Image ImgMerlin 
      Height          =   1095
      Left            =   6360
      MouseIcon       =   "LlamarAmigo.frx":5B60B
      MousePointer    =   99  'Custom
      Stretch         =   -1  'True
      Top             =   4680
      Width           =   615
   End
End
Attribute VB_Name = "FrmLlamarAmigo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Agente_listo As Boolean
Public Agent As String
Private OpcionRes As String * 1
Public respuestaMala As Boolean


Private Sub Agent1_BalloonHide(ByVal CharacterID As String)
Call TerminaDeHablar
End Sub

Private Sub Form_Load()
respuestaMala = False
End Sub

Private Sub Image1_Click()
Unload Me
End Sub

Private Sub ImgMan_Click()
Call LlamarX("comodin4", "4")
End Sub
Private Sub ImgMerlin_Click()
Call LlamarX("comodin2", "2")
End Sub
Private Sub LlamarX(Agente As String, Archivo As String)
If Agente_listo = False Then
    Me.Picture = LoadPicture(App.Path & "\configuracion\" & Agente & ".jpg")
    Agent1.Characters.Load Agente, App.Path & "\configuracion\config" & Archivo & ".dat" 'Busco la ruta del agente
    Agent1.Characters(Agente).Show 'Cargo el agente para que aparezcca
    Agent1.Characters(Agente).Top = 470 'Acomodo al agente en
    Agent1.Characters(Agente).Left = 715 'el lugar mas indicado
    Agent = Agente
    TmTiempo.Enabled = True
    TmRespuesta.Enabled = True
    Agente_listo = True
End If
End Sub
Private Sub ImgOzzar_Click()
Call LlamarX("comodin1", "1")
End Sub
Private Sub ImgWoman_Click()
Call LlamarX("comodin3", "3")
End Sub
Private Sub respuesta()
Dim azar As Integer, respuestas As Integer, respuesta As String, i As Integer, Res As String, ctrl As Control, ValorRes As String
respuestas = Int(Rnd() * 10) + 1

Select Case respuestas
    Case 1
        respuesta = "Esa Pregunta es Muy Fácil,¡Yo Me la sabía!; ¡¡¡Espera un Momento ya me Acorde!!! Yo Creo que la Respuesta Correcta es "
    Case 2
        respuesta = "No Estoy Muy Seguro... Está Muy Dificil la Pregunta. Lo que Tengo entendido hasta ahora según Mis conocimientos es que la respuesta correcta es "
    Case 3
        respuesta = "Esa pregunta la he escuchado en otra parte, pero no recuerdo Bien cuál es la respuesta correcta porque yo no manejo muy bien esta área; pero creo que es "
    Case 4
        respuesta = "Por Favor me llamas solo para fastidiarme; no puedes responder esa pregunta tan fácil; pero como soy tu amigo te ayudaré. Sin duda alguna la respuesta es "
    Case 5
        respuesta = "Ehhhhhh realmente no quisiera ponerte nervioso pero de verdad no sé nada sobre ese tema, pero si me pides mi opinión creo que la respuesta correcta es "
    Case 6
        respuesta = "Hola ¿Como estás? tenías tiempo sin llamarme. Tengo unos chismes que contarte, no te imaginas... ahhhhhh verdad primero tu pregunta, yo creo aunque no estoy seguro que es "
    Case 7
        respuesta = "Veo que necesitas de Mí ayuda, Gracias por llamarme. Mis conocimientos son tan amplios que no te puedes imaginar que tan inteligente soy; por ello facilmente te digo que la respuesta correcta a tu pregunta es "
    Case 8
        respuesta = "Realmente no te puedo decir nada porque no me sé la Respuesta; sin embargo analizando bien la Pregunta la respuesta correcta puede ser "
    Case 9
        respuesta = "Gracias por llamarme pero tengo que consultar la enciclopedia porque tengo una duda..................... ¡Ya la encontré!, la respuesta es "
    Case 10
        respuesta = "La respuesta a tu pregunta es Muy obvia, debías haber pensado antes de llamarme porque te hubieras ahorrado el comodín, pero de todas maneras la respuesta es "
End Select
    For Each ctrl In FrmJuego
        If TypeName(ctrl) = "Label" Then
            ValorRes = ctrl.Caption
            Select Case UCase(ctrl.Name)
                Case UCase("LblRespuestaa")
                    If ValorRes = Respuesta_correcta Then
                        OpcionRes = "A"
                        Exit For
                    End If
                Case UCase("LblRespuestab")
                    If ValorRes = Respuesta_correcta Then
                        OpcionRes = "B"
                        Exit For
                    End If
                Case UCase("LblRespuestac")
                    If ValorRes = Respuesta_correcta Then
                        OpcionRes = "C"
                        Exit For
                    End If
                Case UCase("LblRespuestad")
                    If ValorRes = Respuesta_correcta Then
                        OpcionRes = "D"
                        Exit For
                    End If
            End Select
        End If
    Next
azar = Int(Rnd() * 20) + 1
Select Case azar
    Case 1 To 7, 15 To 20
         Res = respuesta & Respuesta_correcta
    Case 8, 13
        respuestaMala = True
        i = Int(Rnd() * 5) + 1
        Select Case i
            Case 1
                Res = "En este Momento No Me Encuentro, Deja Tu Mensaje despúes del Tono, Gracias piiiiiiiiiiiiiiiiiiiiiii..."
            Case 2
                Res = "Dejame Buscar la Respuesta porque en realidad tengo Muchas dudas; Puede ser La A o la B o la C o la D... Es cuestión de Buscar en los 3000 libros que tengo; confía en Mí, Yo de 20 preguntas fallo 15 pero siempre respondo Bien... Ehhhh aquí está, a no; no es me equivoque... seguiré Buscando"
            Case 3
                Res = "Hola, esa pregunta es fàcil, bueno creo; es que a veces tengo un problema de memoria y se me olvidan las preguntas. Pero la respuesta correcta a tu Pregunta es..... es.... Ya va Espera. ¡Se me olvido!. repiteme la pregunta para ver si me acuerdo; Ahhhhh ya lo tengo.... Espera un segundo ¿que ès èste Programa?"
            Case 4
                Res = "En este Momento No Me Encuentro, Deja Tu Mensaje despúes del Tono, Gracias piiiiiiiiiiiiiiiiiiiiiii..."
            Case 5
                Res = "La respuesta a Tu pregunta no tiene Sentido, porque si te pones analizarla; la respuesta no se encuentra en las cuatro opciones que te dan, ¿no serà que en el Juego se equivocaron?. Yo creo que no està en ninguna de las cuatro opciones o ¿serà que yo estoy equivocado? pero de todas maneras dejame buscar en encarta..."
        End Select
    Case Else
        For Each ctrl In FrmJuego
            If TypeName(ctrl) = "Label" Then
                ValorRes = ctrl.Caption
                Select Case UCase(ctrl.Name)
                    Case UCase("LblRespuestaa")
                        If ValorRes <> Respuesta_correcta Then
                            OpcionRes = "A"
                            Exit For
                        End If
                    Case UCase("LblRespuestab")
                        If ValorRes <> Respuesta_correcta Then
                            OpcionRes = "B"
                            Exit For
                        End If
                    Case UCase("LblRespuestac")
                        If ValorRes <> Respuesta_correcta Then
                            OpcionRes = "C"
                            Exit For
                        End If
                    Case UCase("LblRespuestad")
                        If ValorRes <> Respuesta_correcta Then
                            OpcionRes = "D"
                            Exit For
                        End If
                End Select
            End If
        Next
        Res = respuesta & ValorRes
End Select

Agent1.Characters(Agent).Speak Res
End Sub
Private Sub TmRespuesta_Timer()
Call respuesta
TmRespuesta.Enabled = False
End Sub
Private Sub TerminaDeHablar()
TmTiempo.Enabled = False
LblRespuesta.Visible = True
Agent1.Characters(Agent).Hide
LblRespuesta.Caption = IIf(respuestaMala, "", LblRespuesta.Caption & OpcionRes)
Espera 2
Unload Me
End Sub
Private Sub TmTiempo_Timer()
If LblTiempo = "00:20" Then
   Call TerminaDeHablar
   Exit Sub
End If
If LblTiempo = "" Then
    LblTiempo = "00:01"
Else
    LblTiempo = Mid(LblTiempo.Caption, 1, 4) & Mid(LblTiempo.Caption, 4, Len(LblTiempo.Caption)) + 1
    If Len(LblTiempo.Caption) = 6 Then
        LblTiempo.Caption = Mid(LblTiempo.Caption, 1, 3) & Mid(LblTiempo.Caption, 5, Len(LblTiempo.Caption))
    End If
End If
End Sub
