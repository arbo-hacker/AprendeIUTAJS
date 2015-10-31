VERSION 5.00
Begin VB.Form FrmAyuda 
   BorderStyle     =   0  'None
   Caption         =   "Ayuda"
   ClientHeight    =   8310
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11265
   LinkTopic       =   "Form1"
   MouseIcon       =   "Ayuda.frx":0000
   MousePointer    =   99  'Custom
   Picture         =   "Ayuda.frx":0CCA
   ScaleHeight     =   8310
   ScaleWidth      =   11265
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Height          =   735
      Left            =   360
      MouseIcon       =   "Ayuda.frx":1D83C
      MousePointer    =   99  'Custom
      TabIndex        =   8
      Top             =   5040
      Width           =   1695
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Height          =   495
      Left            =   480
      MouseIcon       =   "Ayuda.frx":1E106
      MousePointer    =   99  'Custom
      TabIndex        =   7
      Top             =   4440
      Width           =   1215
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Height          =   735
      Left            =   120
      MouseIcon       =   "Ayuda.frx":1E9D0
      MousePointer    =   99  'Custom
      TabIndex        =   6
      Top             =   7320
      Width           =   1335
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "<--Seleccione Algún Tema en la Lista de Ayuda-->"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   2295
      Left            =   3360
      TabIndex        =   5
      Top             =   3120
      Width           =   6975
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Height          =   615
      Left            =   240
      MouseIcon       =   "Ayuda.frx":1F29A
      MousePointer    =   99  'Custom
      TabIndex        =   4
      Top             =   5640
      Width           =   1815
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Height          =   495
      Left            =   480
      MouseIcon       =   "Ayuda.frx":1FB64
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   3960
      Width           =   1215
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Height          =   495
      Left            =   360
      MouseIcon       =   "Ayuda.frx":2042E
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   3360
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   240
      MouseIcon       =   "Ayuda.frx":20CF8
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   6255
      Left            =   2760
      Stretch         =   -1  'True
      Top             =   1320
      Width           =   8055
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   120
      MouseIcon       =   "Ayuda.frx":215C2
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   2160
      Width           =   2055
   End
End
Attribute VB_Name = "FrmAyuda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Label1_Click()
Label6.Caption = ""
Image1.Picture = LoadPicture(App.Path & "\configuracion\" & "ayuda-Principal.jpg")

End Sub

Private Sub Label2_Click()
Label6.Caption = ""
Image1.Picture = LoadPicture(App.Path & "\configuracion\" & "ayuda-ElJuego.jpg")
End Sub

Private Sub Label3_Click()
Label6.Caption = ""
Image1.Picture = LoadPicture(App.Path & "\configuracion\" & "ayuda-El 50%.jpg")
End Sub

Private Sub Label4_Click()
Label6.Caption = ""
Image1.Picture = LoadPicture(App.Path & "\configuracion\" & "ayuda-Audiencia.jpg")
End Sub

Private Sub Label5_Click()
Label6.Caption = ""
Image1.Picture = LoadPicture(App.Path & "\configuracion\" & "ayuda-Records.jpg")
End Sub

Private Sub Label7_Click()
Unload Me
Load FrmPrincipal
FrmPrincipal.Show
End Sub

Private Sub Label8_Click()
Label6.Caption = ""
Image1.Picture = LoadPicture(App.Path & "\configuracion\" & "ayuda-Llamada.jpg")
End Sub

Private Sub Label9_Click()
Label6.Caption = ""
Image1.Picture = LoadPicture(App.Path & "\configuracion\" & "ayuda-MenúPreg.jpg")
End Sub
