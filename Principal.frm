VERSION 5.00
Begin VB.Form FrmPrincipal 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   6405
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9645
   LinkTopic       =   "Form1"
   MouseIcon       =   "Principal.frx":0000
   MousePointer    =   99  'Custom
   Picture         =   "Principal.frx":0CCA
   ScaleHeight     =   6405
   ScaleWidth      =   9645
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer TmVerde 
      Enabled         =   0   'False
      Interval        =   150
      Left            =   360
      Top             =   240
   End
   Begin VB.Image ImgCreditosV 
      Height          =   585
      Left            =   1875
      Picture         =   "Principal.frx":A08F
      Top             =   4320
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.Image ImgRecordsV 
      Height          =   540
      Left            =   1875
      Picture         =   "Principal.frx":F47D
      Top             =   3405
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.Image ImgRecordsN 
      Height          =   540
      Left            =   1875
      Picture         =   "Principal.frx":14756
      Top             =   3405
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.Image ImgConcursoV 
      Height          =   540
      Left            =   1875
      Picture         =   "Principal.frx":19C20
      Top             =   2445
      Visible         =   0   'False
      Width           =   2805
   End
   Begin VB.Image ImgConcursoN 
      Height          =   540
      Left            =   1880
      Picture         =   "Principal.frx":1F11C
      Top             =   2450
      Visible         =   0   'False
      Width           =   2805
   End
   Begin VB.Image ImgSalirV 
      Height          =   570
      Left            =   1920
      Picture         =   "Principal.frx":24770
      Top             =   5200
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.Image ImgSalirN 
      Height          =   570
      Left            =   1920
      Picture         =   "Principal.frx":28FE8
      Top             =   5200
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.Image ImgCreditosN 
      Height          =   585
      Left            =   1875
      Picture         =   "Principal.frx":2D99F
      Top             =   4320
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.Image ImgJugar 
      Height          =   540
      Left            =   1875
      MouseIcon       =   "Principal.frx":32E51
      MousePointer    =   99  'Custom
      Picture         =   "Principal.frx":3371B
      Top             =   2445
      Width           =   2805
   End
   Begin VB.Image ImgRecords 
      Height          =   540
      Left            =   1875
      MouseIcon       =   "Principal.frx":37E1C
      MousePointer    =   99  'Custom
      Picture         =   "Principal.frx":386E6
      Top             =   3405
      Width           =   2775
   End
   Begin VB.Image ImgCreditos 
      Height          =   585
      Left            =   1875
      MouseIcon       =   "Principal.frx":3CB1B
      MousePointer    =   99  'Custom
      Picture         =   "Principal.frx":3D3E5
      Top             =   4320
      Width           =   2925
   End
   Begin VB.Image ImgSalir 
      Height          =   570
      Left            =   1920
      MouseIcon       =   "Principal.frx":3F538
      MousePointer    =   99  'Custom
      Picture         =   "Principal.frx":3FE02
      Top             =   5200
      Width           =   2775
   End
End
Attribute VB_Name = "FrmPrincipal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public var As Integer


Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
ImgRecordsN.Visible = False
ImgConcursoN.Visible = False
ImgSalirN.Visible = False
ImgCreditosN.Visible = False
End Sub




Private Sub ImgConcursoN_Click()
ImgJugar_Click
End Sub

Private Sub ImgConcursoV_Click()
ImgJugar_Click
End Sub

Private Sub ImgCreditos_Click()
var = 4
TmVerde.Enabled = True
Espera 1
TmVerde.Enabled = False
Unload Me
Load FrmPreguntas
FrmPreguntas.Show

End Sub

Private Sub ImgCreditos_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
ImgCreditosN.Visible = True
End Sub

Private Sub ImgCreditosN_Click()
ImgCreditos_Click
End Sub

Private Sub ImgCreditosV_Click()
ImgCreditos_Click
End Sub

Private Sub ImgJugar_Click()
var = 1
TmVerde.Enabled = True
Espera 1
TmVerde.Enabled = False
Unload Me
Load FrmNombre
FrmNombre.Show
'Unload Me 'Descargo este formulario
'Load FrmJuego 'Cargo el formulario FrmJuego
'FrmJuego.Show 'Hago que aparezca el formulario FrmJuego
End Sub

Private Sub ImgJugar_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
ImgConcursoN.Visible = True
End Sub

Private Sub ImgRecords_Click()
var = 2
TmVerde.Enabled = True
Espera 1
TmVerde.Enabled = False
Unload Me 'Descargo este formulario
Load FrmRecords 'Cargo el formulario FrmRecords
FrmRecords.Show 'Hago que aparezca el formulario FrmRecords
End Sub

Private Sub ImgRecords_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
ImgRecordsN.Visible = True
End Sub

Private Sub ImgRecordsN_Click()
ImgRecords_Click
End Sub

Private Sub ImgRecordsV_Click()
ImgRecords_Click
End Sub

Private Sub ImgSalir_Click()
var = 3
TmVerde.Enabled = True
Espera 1
TmVerde.Enabled = False
FrmSalir.Show 1
End Sub

Private Sub ImgSalir_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
ImgSalirN.Visible = True
End Sub

Private Sub ImgSalirN_Click()
ImgSalir_Click
End Sub

Private Sub ImgSalirV_Click()
ImgSalir_Click
End Sub

Private Sub TmVerde_Timer()
Select Case var
    Case 1
        If ImgConcursoV.Visible = False Then
            ImgConcursoV.Visible = True
        Else
            ImgConcursoV.Visible = False
        End If
    Case 2
        If ImgRecordsV.Visible = False Then
            ImgRecordsV.Visible = True
        Else
            ImgRecordsV.Visible = False
        End If
    Case 3
        If ImgSalirV.Visible = False Then
            ImgSalirV.Visible = True
        Else
            ImgSalirV.Visible = False
        End If
    Case 4
        If ImgCreditosV.Visible = False Then
            ImgCreditosV.Visible = True
        Else
            ImgCreditosV.Visible = False
        End If
    Case 5
        If ImgAyudaV.Visible = False Then
            ImgAyudaV.Visible = True
        Else
            ImgAyudaV.Visible = False
        End If
End Select
    
End Sub

