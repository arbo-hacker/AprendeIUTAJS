VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form FrmFinal 
   ClientHeight    =   1785
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   2445
   LinkTopic       =   "Form1"
   ScaleHeight     =   1785
   ScaleWidth      =   2445
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   21000
      Left            =   240
      Top             =   360
   End
   Begin WMPLibCtl.WindowsMediaPlayer WindowsMediaPlayer1 
      Height          =   1800
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   2580
      URL             =   "C:\Documents and Settings\Administrador\Escritorio\Millonario\Configuracion\Final.Vid"
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
      autoStart       =   -1  'True
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   50
      mute            =   0   'False
      uiMode          =   "none"
      stretchToFit    =   0   'False
      windowlessVideo =   -1  'True
      enabled         =   0   'False
      enableContextMenu=   -1  'True
      fullScreen      =   -1  'True
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   4551
      _cy             =   3175
   End
End
Attribute VB_Name = "FrmFinal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
WindowsMediaPlayer1.URL = App.Path & "\configuracion\final.vid"
End Sub

Private Sub Timer1_Timer()
WindowsMediaPlayer1.Close
Unload Me
FrmCreditos.Show
End Sub

