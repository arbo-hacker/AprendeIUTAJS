VERSION 5.00
Begin VB.Form FrmAudiencia 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Audiencia"
   ClientHeight    =   3420
   ClientLeft      =   13470
   ClientTop       =   3255
   ClientWidth     =   2685
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "Audiencia.frx":0000
   MousePointer    =   99  'Custom
   Picture         =   "Audiencia.frx":0CCA
   ScaleHeight     =   3420
   ScaleWidth      =   2685
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "Audiencia.frx":1828F
      Left            =   7320
      List            =   "Audiencia.frx":18291
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   360
      Visible         =   0   'False
      Width           =   390
   End
   Begin VB.Image ImgVolver 
      Height          =   1095
      Left            =   8160
      Top             =   5760
      Width           =   1455
   End
   Begin VB.Image ImgDiagrama50 
      Height          =   1155
      Index           =   2
      Left            =   3240
      Picture         =   "Audiencia.frx":18293
      Top             =   1800
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Image ImgDiagrama50 
      Height          =   1155
      Index           =   1
      Left            =   2880
      Picture         =   "Audiencia.frx":1A243
      Top             =   1800
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Image ImgDiagrama25 
      Height          =   675
      Index           =   4
      Left            =   4320
      Picture         =   "Audiencia.frx":1C1F3
      Top             =   2280
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Image ImgDiagrama25 
      Height          =   675
      Index           =   3
      Left            =   3960
      Picture         =   "Audiencia.frx":1E12D
      Top             =   2280
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Image ImgDiagrama25 
      Height          =   675
      Index           =   2
      Left            =   3600
      Picture         =   "Audiencia.frx":20067
      Top             =   2280
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Image ImgDiagrama25 
      Height          =   675
      Index           =   1
      Left            =   4680
      Picture         =   "Audiencia.frx":21FA1
      Top             =   2280
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Image ImgDiagrama3 
      Height          =   1155
      Left            =   6120
      Picture         =   "Audiencia.frx":23EDB
      Top             =   1800
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Image ImgDiagrama2 
      Height          =   465
      Left            =   5760
      Picture         =   "Audiencia.frx":25E8B
      Top             =   2520
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Image ImgDiagrama1 
      Height          =   285
      Left            =   5400
      Picture         =   "Audiencia.frx":27DB8
      Top             =   2640
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Image ImgDiagrama4 
      Height          =   1680
      Left            =   5040
      Picture         =   "Audiencia.frx":29CE1
      Top             =   1320
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Label LblPorcentaje 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   285
      Index           =   4
      Left            =   8280
      TabIndex        =   3
      Top             =   1875
      Width           =   90
   End
   Begin VB.Label LblPorcentaje 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   285
      Index           =   3
      Left            =   7700
      TabIndex        =   2
      Top             =   1800
      Width           =   90
   End
   Begin VB.Label LblPorcentaje 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   285
      Index           =   2
      Left            =   7065
      TabIndex        =   1
      Top             =   1800
      Width           =   90
   End
   Begin VB.Label LblPorcentaje 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   285
      Index           =   1
      Left            =   6495
      TabIndex        =   0
      Top             =   1800
      Width           =   90
   End
End
Attribute VB_Name = "FrmAudiencia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then
    Unload Me
End If
End Sub

Private Sub Form_Load()
Dim i As Integer, y As Integer, Porcen(2) As Integer, x As Integer, z As Integer, _
PorcenB As Boolean
If Comodin = True Then
    For i = 1 To 4
        If Porcentaje(i) <> 0 Then
            y = y + 1
            Porcen(y) = Porcentaje(i)
            If y = 1 Then x = i Else z = i
        End If
    Next
    y = 0
    LblPorcentaje(x).Caption = Porcen(1) & "%"
    LblPorcentaje(z).Caption = Porcen(2) & "%"
    If Porcen(1) = Porcen(2) Then
        For i = 1 To 4
            If LblPorcentaje(i).Caption <> "" Then
                y = y + 1
                ImgDiagrama50(y).Left = LblPorcentaje(i).Left
            End If
        Next
    y = 0
    ImgDiagrama50(1).Visible = True
    ImgDiagrama50(2).Visible = True
    Else
        ImgDiagrama4.Left = LblPorcentaje(z).Left
        ImgDiagrama2.Left = LblPorcentaje(x).Left

        ImgDiagrama2.Visible = True
        ImgDiagrama4.Visible = True
    End If
Else
    For i = 1 To 4
        If Len(Trim(Str(Porcentaje(i)))) = 1 Then
            LblPorcentaje(i).Caption = "0" & Trim(Porcentaje(i)) & "%"
        Else
            LblPorcentaje(i).Caption = Trim(Porcentaje(i)) & "%"
        End If
    Next
    If Porcentaje(1) = Porcentaje(2) And Porcentaje(3) = Porcentaje(4) Then
        ImgDiagrama25(1).Left = 360
        ImgDiagrama25(2).Left = 960
        ImgDiagrama25(3).Left = 1560
        ImgDiagrama25(4).Left = 2160
        For i = 1 To 4
            ImgDiagrama25(i).Visible = True
        Next
        Exit Sub
    End If
    
    
    With Combo1
        For i = 1 To 4
            If Len(Trim(Str(Porcentaje(i)))) = 1 Then
                .AddItem "0" & Trim(Str(Porcentaje(i)))
            Else
                .AddItem Trim(Str(Porcentaje(i)))
            End If
        Next
        For i = 0 To 3
            y = y + 1
            mayor(y) = .List(i)
        Next
        .ListIndex = 0
    End With
    
    Select Case mayor(1)
        Case Porcentaje(1)
            ImgDiagrama1.Left = 360
        Case Porcentaje(2)
            ImgDiagrama1.Left = 960
        Case Porcentaje(3)
            ImgDiagrama1.Left = 1560
        Case Porcentaje(4)
            ImgDiagrama1.Left = 2160
    End Select
    Select Case mayor(2)
        Case Porcentaje(1)
            ImgDiagrama2.Left = 360
        Case Porcentaje(2)
            ImgDiagrama2.Left = 960
        Case Porcentaje(3)
            ImgDiagrama2.Left = 1560
        Case Porcentaje(4)
            ImgDiagrama2.Left = 2160
    End Select
    Select Case mayor(3)
        Case Porcentaje(1)
            ImgDiagrama3.Left = 360
        Case Porcentaje(2)
            ImgDiagrama3.Left = 960
        Case Porcentaje(3)
            ImgDiagrama3.Left = 1560
        Case Porcentaje(4)
            ImgDiagrama3.Left = 2160
    End Select
    Select Case mayor(4)
        Case Porcentaje(1)
            ImgDiagrama4.Left = 360
        Case Porcentaje(2)
            ImgDiagrama4.Left = 960
        Case Porcentaje(3)
            ImgDiagrama4.Left = 1560
        Case Porcentaje(4)
            ImgDiagrama4.Left = 2160
    End Select
    ImgDiagrama1.Visible = True
    ImgDiagrama2.Visible = True
    ImgDiagrama3.Visible = True
    ImgDiagrama4.Visible = True
End If
End Sub
Private Sub ImgVolver_Click()
Unload Me
End Sub
