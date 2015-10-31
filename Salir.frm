VERSION 5.00
Begin VB.Form FrmSalir 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Salir"
   ClientHeight    =   6000
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4785
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "Salir.frx":0000
   MousePointer    =   99  'Custom
   Moveable        =   0   'False
   Picture         =   "Salir.frx":0CCA
   ScaleHeight     =   6000
   ScaleWidth      =   4785
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer TmSiNo 
      Enabled         =   0   'False
      Interval        =   150
      Left            =   240
      Top             =   240
   End
   Begin VB.Image ImgSalirNo 
      Height          =   495
      Left            =   360
      MouseIcon       =   "Salir.frx":10326
      MousePointer    =   99  'Custom
      Top             =   5160
      Width           =   2415
   End
   Begin VB.Image ImgSalirSi 
      Height          =   495
      Left            =   2160
      MouseIcon       =   "Salir.frx":10BF0
      MousePointer    =   99  'Custom
      Top             =   4320
      Width           =   2415
   End
   Begin VB.Image ImgSiV 
      Height          =   645
      Left            =   1965
      Picture         =   "Salir.frx":114BA
      Top             =   4230
      Visible         =   0   'False
      Width           =   2700
   End
   Begin VB.Image ImgSiN 
      Height          =   645
      Left            =   1965
      Picture         =   "Salir.frx":11E19
      Top             =   4230
      Visible         =   0   'False
      Width           =   2700
   End
   Begin VB.Image ImgNoV 
      Height          =   660
      Left            =   170
      Picture         =   "Salir.frx":12834
      Top             =   5100
      Visible         =   0   'False
      Width           =   2820
   End
   Begin VB.Image ImgNoN 
      Height          =   660
      Left            =   170
      Picture         =   "Salir.frx":17567
      Top             =   5100
      Visible         =   0   'False
      Width           =   2820
   End
End
Attribute VB_Name = "FrmSalir"
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

Private Sub ImgSalirSi_Click()
var = 1
TmSiNo.Enabled = True
Espera 1.3
TmSiNo.Enabled = False
End
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
