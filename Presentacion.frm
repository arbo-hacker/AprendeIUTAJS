VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form FrmPresentacion 
   Caption         =   "Form1"
   ClientHeight    =   7185
   ClientLeft      =   8820
   ClientTop       =   5700
   ClientWidth     =   7290
   LinkTopic       =   "Form1"
   ScaleHeight     =   7185
   ScaleWidth      =   7290
   Visible         =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   21000
      Left            =   0
      Top             =   120
   End
   Begin WMPLibCtl.WindowsMediaPlayer WindowsMediaPlayer1 
      Height          =   7200
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   7380
      URL             =   "C:\Documents and Settings\Administrador\Escritorio\Millonario\Configuracion\Presentacion.vid"
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
      _cx             =   13018
      _cy             =   12700
   End
End
Attribute VB_Name = "FrmPresentacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
WindowsMediaPlayer1.URL = App.Path & "\configuracion\presentacion.vid"
End Sub

Private Sub Timer1_Timer()
WindowsMediaPlayer1.Close
Unload Me
FrmPrincipal.Show
End Sub
