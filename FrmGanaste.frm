VERSION 5.00
Begin VB.Form FrmGanaste 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ganaste"
   ClientHeight    =   3570
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5910
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "FrmGanaste.frx":0000
   ScaleHeight     =   3570
   ScaleWidth      =   5910
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer TmSalir 
      Interval        =   3000
      Left            =   480
      Top             =   360
   End
End
Attribute VB_Name = "FrmGanaste"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
TmSalir.Enabled = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
TmSalir.Enabled = False
Load FrmPrincipal
FrmPrincipal.Show
End Sub

Private Sub TmSalir_Timer()
Unload Me
End Sub

