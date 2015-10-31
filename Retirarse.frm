VERSION 5.00
Begin VB.Form FrmRetirarse 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Retirarse"
   ClientHeight    =   6000
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4815
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MouseIcon       =   "Retirarse.frx":0000
   MousePointer    =   99  'Custom
   Picture         =   "Retirarse.frx":0CCA
   ScaleHeight     =   6000
   ScaleWidth      =   4815
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer TmSiNo 
      Enabled         =   0   'False
      Interval        =   150
      Left            =   0
      Top             =   0
   End
   Begin VB.Image ImgSalirSi 
      Height          =   495
      Left            =   2160
      MouseIcon       =   "Retirarse.frx":FCB6
      MousePointer    =   99  'Custom
      Top             =   4320
      Width           =   2415
   End
   Begin VB.Image ImgSalirNo 
      Height          =   495
      Left            =   315
      MouseIcon       =   "Retirarse.frx":10580
      MousePointer    =   99  'Custom
      Top             =   5250
      Width           =   2415
   End
   Begin VB.Image ImgSiV 
      Height          =   645
      Left            =   1995
      Picture         =   "Retirarse.frx":10E4A
      Top             =   4240
      Visible         =   0   'False
      Width           =   2700
   End
   Begin VB.Image ImgSiN 
      Height          =   645
      Left            =   1995
      MouseIcon       =   "Retirarse.frx":117A9
      MousePointer    =   99  'Custom
      Picture         =   "Retirarse.frx":12073
      Top             =   4240
      Visible         =   0   'False
      Width           =   2700
   End
   Begin VB.Image ImgNoV 
      Height          =   660
      Left            =   200
      Picture         =   "Retirarse.frx":12A8E
      Top             =   5110
      Visible         =   0   'False
      Width           =   2820
   End
   Begin VB.Image ImgNoN 
      Height          =   660
      Left            =   200
      MouseIcon       =   "Retirarse.frx":177C1
      MousePointer    =   99  'Custom
      Picture         =   "Retirarse.frx":1808B
      Top             =   5110
      Visible         =   0   'False
      Width           =   2820
   End
End
Attribute VB_Name = "FrmRetirarse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private var As Integer
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
ImgSiN.Visible = False
ImgNoN.Visible = False
End Sub

Private Sub ImgSalirNo_Click()
var = 2
TmSiNo.Enabled = True
Espera 1.3
TmSiNo.Enabled = False
Unload Me
End Sub

Private Sub ImgSalirNo_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
ImgNoN.Visible = True
End Sub
Private Sub Retirarse()
Dim Valor_Pregunta As String
Valor_Pregunta = Valores_Preguntas(Cantidad_preguntas - 2, True)
If Valor_Pregunta > 0 Then
    Sql.CommandText = Añadir("records", Nombre_jugador, "#" & Valor_Pregunta)
    Set RS = Sql.Execute
End If
FrmJuego.Agent1.Characters("James").Speak "Has sido un excelente jugador"
Espera (2)
Unload Me
Unload FrmJuego
Unload FrmInfPreguntas
FrmPrincipal.Show
End Sub
Private Sub ImgSalirSi_Click()
var = 1
TmSiNo.Enabled = True
Espera 1.3
TmSiNo.Enabled = False
If Pregunta_Actual = 1 Then
    FrmJuego.Agent1.Characters("James").Speak "Todavia no has ganado nada"
    Unload Me
Else
    Call Retirarse
End If
End Sub

Private Sub ImgSalirSi_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
ImgSiN.Visible = True
End Sub

Private Sub TmSiNo_Timer()
If var = 1 Then
    If ImgSiV.Visible = False Then
        ImgSiV.Visible = True
    Else
        ImgSiV.Visible = False
    End If
ElseIf var = 2 Then
    If ImgNoV.Visible = False Then
        ImgNoV.Visible = True
    Else
        ImgNoV.Visible = False
    End If
End If
End Sub

