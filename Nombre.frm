VERSION 5.00
Begin VB.Form FrmNombre 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Nombre"
   ClientHeight    =   7500
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10920
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MouseIcon       =   "Nombre.frx":0000
   MousePointer    =   99  'Custom
   Picture         =   "Nombre.frx":0CCA
   ScaleHeight     =   7500
   ScaleWidth      =   10920
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1200
      TabIndex        =   0
      Top             =   7800
      Width           =   375
   End
   Begin VB.TextBox TxtNombre 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   4815
      Left            =   7080
      TabIndex        =   1
      Text            =   "Introduzca su nombre"
      ToolTipText     =   "Escribe Tu nombre y presiona la tecla enter"
      Top             =   1800
      Width           =   3255
   End
End
Attribute VB_Name = "FrmNombre"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub TxtNombre_GotFocus()
If TxtNombre.Text = "Introduzca su nombre" Then
    TxtNombre.ForeColor = &H0&
    TxtNombre.Text = ""
End If
End Sub

Private Sub TxtNombre_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 And TxtNombre.Text <> "" Then
    Nombre_jugador = TxtNombre.Text
    Unload Me
    Load FrmJuego
    'FrmJuego.Picture = LoadPicture(App.Path & "\configuracion\millonario " & Quien & ".jpg")
    FrmJuego.Show
End If
End Sub

Private Sub TxtNombre_LostFocus()
If TxtNombre.Text = "" Then
    TxtNombre.ForeColor = &H808080
    TxtNombre.Text = "Introduzca su nombre"
'ElseIf TxtNombre.Text = "Introduzca su nombre" Then
Else
    TxtNombre.ForeColor = &H0&
End If
End Sub
