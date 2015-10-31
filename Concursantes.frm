VERSION 5.00
Begin VB.Form FrmConcursantes 
   Caption         =   "Concursantes"
   ClientHeight    =   7125
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9585
   LinkTopic       =   "Form1"
   MouseIcon       =   "Concursantes.frx":0000
   MousePointer    =   99  'Custom
   Picture         =   "Concursantes.frx":0CCA
   ScaleHeight     =   7125
   ScaleWidth      =   9585
   StartUpPosition =   2  'CenterScreen
   Begin VB.Image ImgNarizon 
      Height          =   3255
      Left            =   8280
      MouseIcon       =   "Concursantes.frx":F627
      MousePointer    =   99  'Custom
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Image ImgNegra 
      Height          =   2415
      Left            =   6960
      MouseIcon       =   "Concursantes.frx":FEF1
      MousePointer    =   99  'Custom
      Top             =   2760
      Width           =   1095
   End
   Begin VB.Image ImgFrenton 
      Height          =   3255
      Left            =   120
      MouseIcon       =   "Concursantes.frx":107BB
      MousePointer    =   99  'Custom
      Top             =   2040
      Width           =   1095
   End
   Begin VB.Image ImgCacheton 
      Height          =   3495
      Left            =   1560
      MouseIcon       =   "Concursantes.frx":11085
      MousePointer    =   99  'Custom
      Top             =   1800
      Width           =   1095
   End
End
Attribute VB_Name = "FrmConcursantes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Me.Picture = LoadPicture(App.Path & "\configuracion\Concursantes.jpg")
End Sub

Private Sub ImgCacheton_Click()
Quien = "cacheton"
Unload Me
Load FrmNombre
'FrmNombre.Picture = LoadPicture(App.Path & "\configuracion\nombre cacheton.jpg")
FrmNombre.Show
End Sub

Private Sub ImgCacheton_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Me.Picture = LoadPicture(App.Path & "\configuracion\Concursante cacheton.jpg")

End Sub

Private Sub ImgFrenton_Click()
Quien = "frenton"
Unload Me
Load FrmNombre
FrmNombre.Picture = LoadPicture(App.Path & "\configuracion\nombre frenton.jpg")
FrmNombre.Show
End Sub

Private Sub ImgFrenton_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Me.Picture = LoadPicture(App.Path & "\configuracion\Concursante frenton.jpg")

End Sub

Private Sub ImgNarizon_Click()
Quien = "narizon"
Unload Me
Load FrmNombre
FrmNombre.Picture = LoadPicture(App.Path & "\configuracion\nombre narizon.jpg")
FrmNombre.Show
End Sub

Private Sub ImgNarizon_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Me.Picture = LoadPicture(App.Path & "\configuracion\Concursante narizon.jpg")

End Sub

Private Sub ImgNegra_Click()
Quien = "negra"
Unload Me
Load FrmNombre
FrmNombre.Picture = LoadPicture(App.Path & "\configuracion\Nombre negra.jpg")
FrmNombre.Show
End Sub

Private Sub ImgNegra_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Me.Picture = LoadPicture(App.Path & "\configuracion\Concursante negra.jpg")
End Sub
