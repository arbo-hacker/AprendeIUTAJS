VERSION 5.00
Begin VB.Form FrmCreditos 
   BackColor       =   &H80000008&
   BorderStyle     =   0  'None
   ClientHeight    =   8055
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10770
   LinkTopic       =   "Form1"
   ScaleHeight     =   8055
   ScaleWidth      =   10770
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer TmCambiar 
      Interval        =   5000
      Left            =   240
      Top             =   240
   End
   Begin VB.Image Image3 
      Height          =   2775
      Left            =   7200
      Top             =   3240
      Width           =   3615
   End
   Begin VB.Image Image2 
      Height          =   6255
      Left            =   6480
      Stretch         =   -1  'True
      Top             =   1320
      Width           =   4095
   End
   Begin VB.Image Image1 
      Height          =   5415
      Left            =   240
      Stretch         =   -1  'True
      Top             =   1440
      Width           =   3615
   End
   Begin VB.Label Lblquienes 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   360
      Index           =   2
      Left            =   5145
      TabIndex        =   3
      Top             =   3960
      Width           =   105
   End
   Begin VB.Label Lblquienes 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   360
      Index           =   1
      Left            =   5175
      TabIndex        =   2
      Top             =   2400
      Width           =   105
   End
   Begin VB.Label Lblquienes 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   360
      Index           =   3
      Left            =   5175
      TabIndex        =   1
      Top             =   5400
      Width           =   105
   End
   Begin VB.Label LblTitulo 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   765
      Left            =   5280
      TabIndex        =   0
      Top             =   720
      Width           =   135
   End
End
Attribute VB_Name = "FrmCreditos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" _
    (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long

Public Tiempo As Integer

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then
    Unload Me
    Load FrmPrincipal
    FrmPrincipal.Show
End If
End Sub

Private Sub Form_Load()


Const SND_ASYNC = &H1     'modo asíncrono. La función retorna una vez iniciada la música (sonido en background).
Const SND_LOOP = &H8      'La música seguirá sonando repetidamente hasta
                          'que la función sndPlaySound sea llamada de nuevo con un valor nulo para NombreWav (NULL).


'Para tocar un WAV de forma repetitiva, lo llamas así:
Call sndPlaySound(App.Path & "\configuracion\musica.wav", SND_ASYNC + SND_LOOP)

'Para detener lo que se esté tocando


End Sub

Public Sub Creditos(Titulo As String, ParamArray Quienes() As Variant)
Dim i As Integer, o As Integer
LblTitulo.Caption = Titulo


For i = 1 To UBound(Quienes) + 1
Lblquienes(i).Caption = Quienes(i - 1)
Next
End Sub

Private Sub TmCambiar_Timer()
DoEvents
Select Case Tiempo
    Case 0
        Call Creditos("Producido por", "Alejandro Barreto")
        Image1.Picture = LoadPicture(App.Path & "\Configuracion\" & "CruzFire.jpg")
        Image3.Picture = LoadPicture(App.Path & "\Configuracion\" & "Notihackers.jpg")
    Case 1
        Call Creditos("Diseñado por", "Alejandro Barreto")
        Image1.Picture = LoadPicture(App.Path & "\Configuracion\" & "lost-angel-couv.jpg")
        Image2.Picture = LoadPicture("")
        Image3.Visible = False
    Case 2
        Call Creditos("Programado por", "Alejandro Barreto")
        Image1.Picture = LoadPicture(App.Path & "\Configuracion\" & "AngelLost.jpg")
        Image2.Picture = LoadPicture(App.Path & "\Configuracion\" & "Green_biohazard.jpg")
    Case 3
        Call Creditos("Agradecimientos a", "Mi mamá" & vbCrLf & vbCrLf & "Mi Esposa" & vbCrLf & "Mariangela", "Mi Tutor" & vbCrLf & "Danny Robles" & vbCrLf & vbCrLf & "Mi coordinador" & vbCrLf & "de escuela" & vbCrLf & "Wilmen Gonzalez")
        Image1.Picture = LoadPicture(App.Path & "\Configuracion\" & "AngelAbadon.jpg")
        Image2.Picture = LoadPicture(App.Path & "\Configuracion\" & "bIoHaZaRd.JPG")
        Call sndPlaySound(ByVal App.Path & "\configuracion\musica.wav", 0)
    Case 4
        TmCambiar.Enabled = False
        Unload Me
        Load FrmPrincipal
        FrmPrincipal.Show
End Select
Tiempo = Tiempo + 1
End Sub
