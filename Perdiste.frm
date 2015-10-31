VERSION 5.00
Object = "{F5BE8BC2-7DE6-11D0-91FE-00C04FD701A5}#2.0#0"; "agentctl.dll"
Begin VB.Form FrmPerdiste 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Perdiste"
   ClientHeight    =   6990
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11955
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "Perdiste.frx":0000
   MousePointer    =   99  'Custom
   Picture         =   "Perdiste.frx":0CCA
   ScaleHeight     =   6990
   ScaleWidth      =   11955
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer TmSalir 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   120
      Top             =   120
   End
   Begin AgentObjectsCtl.Agent Agent1 
      Left            =   480
      Top             =   2640
   End
End
Attribute VB_Name = "FrmPerdiste"
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
