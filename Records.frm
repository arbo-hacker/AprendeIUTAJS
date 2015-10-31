VERSION 5.00
Begin VB.Form FrmRecords 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Records"
   ClientHeight    =   7200
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9630
   LinkTopic       =   "Form1"
   MouseIcon       =   "Records.frx":0000
   MousePointer    =   99  'Custom
   Picture         =   "Records.frx":0CCA
   ScaleHeight     =   7200
   ScaleWidth      =   9630
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label LblRecord 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   5
      Left            =   5920
      TabIndex        =   11
      Top             =   5250
      Width           =   2175
   End
   Begin VB.Label LblRecord 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   4
      Left            =   5920
      TabIndex        =   10
      Top             =   4600
      Width           =   2175
   End
   Begin VB.Label LblRecord 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   3
      Left            =   5920
      TabIndex        =   9
      Top             =   4000
      Width           =   2175
   End
   Begin VB.Label LblRecord 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   6
      Left            =   5920
      TabIndex        =   8
      Top             =   5880
      Width           =   2175
   End
   Begin VB.Label LblRecord 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   2
      Left            =   5920
      TabIndex        =   7
      Top             =   3360
      Width           =   2175
   End
   Begin VB.Label LblRecord 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   1
      Left            =   5920
      TabIndex        =   6
      Top             =   2760
      Width           =   2175
   End
   Begin VB.Label LblJugador 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   2
      Left            =   650
      TabIndex        =   5
      Top             =   3360
      Width           =   3855
   End
   Begin VB.Label LblJugador 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   3
      Left            =   650
      TabIndex        =   4
      Top             =   4000
      Width           =   3855
   End
   Begin VB.Label LblJugador 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   4
      Left            =   650
      TabIndex        =   3
      Top             =   4600
      Width           =   3855
   End
   Begin VB.Label LblJugador 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   5
      Left            =   650
      TabIndex        =   2
      Top             =   5250
      Width           =   3855
   End
   Begin VB.Label LblJugador 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   6
      Left            =   650
      TabIndex        =   1
      Top             =   5880
      Width           =   3855
   End
   Begin VB.Label LblJugador 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   1
      Left            =   650
      TabIndex        =   0
      Top             =   2760
      Width           =   3855
   End
   Begin VB.Image ImgVolver 
      Height          =   855
      Left            =   7080
      MouseIcon       =   "Records.frx":1D877
      MousePointer    =   99  'Custom
      Top             =   6360
      Width           =   2535
   End
End
Attribute VB_Name = "FrmRecords"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Call AbrirBD
Call Generar_records
End Sub

Private Sub Form_Unload(Cancel As Integer)
Load FrmPrincipal 'Cargo el formulario FrmPrincipal
FrmPrincipal.Show 'Hago que aparezca el formulario FrmPrincipal
End Sub

Private Sub ImgVolver_Click()
Unload Me 'Descargo este formulario
End Sub
Private Sub Generar_records()
Dim i As Integer

Sql.CommandText = "Select * from records order by record desc"
Set RS = Sql.Execute

If RS.EOF = False Then
    RS.MoveFirst
    Do While Not RS.EOF
        If i <= 5 Then
            i = i + 1
            LblJugador(i).Caption = RS!nombre
            LblRecord(i).Caption = Format(RS!Record, "###,###,###")
        End If
        RS.MoveNext
    Loop
End If
End Sub
