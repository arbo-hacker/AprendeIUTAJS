VERSION 5.00
Begin VB.Form FrmInfPreguntas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Información Preguntas"
   ClientHeight    =   7230
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9690
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MouseIcon       =   "InfPreguntas.frx":0000
   MousePointer    =   99  'Custom
   Picture         =   "InfPreguntas.frx":0CCA
   ScaleHeight     =   7230
   ScaleWidth      =   9690
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer TmCambiar 
      Interval        =   500
      Left            =   0
      Top             =   120
   End
   Begin VB.Image Image1 
      Height          =   735
      Left            =   7440
      MouseIcon       =   "InfPreguntas.frx":20B17
      MousePointer    =   99  'Custom
      Top             =   6600
      Width           =   2055
   End
   Begin VB.Image ImgRetirarse 
      Height          =   615
      Left            =   960
      MouseIcon       =   "InfPreguntas.frx":213E1
      MousePointer    =   99  'Custom
      Top             =   6360
      Width           =   1935
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Index           =   50
      Left            =   6495
      TabIndex        =   49
      Top             =   4995
      Width           =   75
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Index           =   49
      Left            =   6495
      TabIndex        =   48
      Top             =   4725
      Width           =   75
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Index           =   48
      Left            =   6495
      TabIndex        =   47
      Top             =   4455
      Width           =   75
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Index           =   47
      Left            =   6495
      TabIndex        =   46
      Top             =   4185
      Width           =   75
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Index           =   46
      Left            =   6495
      TabIndex        =   45
      Top             =   3915
      Width           =   75
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Index           =   45
      Left            =   6495
      TabIndex        =   44
      Top             =   3645
      Width           =   75
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Index           =   44
      Left            =   6495
      TabIndex        =   43
      Top             =   3375
      Width           =   75
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Index           =   43
      Left            =   6495
      TabIndex        =   42
      Top             =   3120
      Width           =   75
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Index           =   42
      Left            =   6495
      TabIndex        =   41
      Top             =   2850
      Width           =   75
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Index           =   41
      Left            =   6495
      TabIndex        =   40
      Top             =   2580
      Width           =   75
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Index           =   40
      Left            =   6495
      TabIndex        =   39
      Top             =   2310
      Width           =   75
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Index           =   39
      Left            =   6495
      TabIndex        =   38
      Top             =   2040
      Width           =   75
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Index           =   38
      Left            =   6495
      TabIndex        =   37
      Top             =   1770
      Width           =   75
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Index           =   37
      Left            =   6495
      TabIndex        =   36
      Top             =   1485
      Width           =   75
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Index           =   36
      Left            =   6495
      TabIndex        =   35
      Top             =   1215
      Width           =   75
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Index           =   35
      Left            =   6495
      TabIndex        =   34
      Top             =   945
      Width           =   75
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Index           =   34
      Left            =   4200
      TabIndex        =   33
      Top             =   5145
      Width           =   75
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Index           =   33
      Left            =   4200
      TabIndex        =   32
      Top             =   4875
      Width           =   75
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Index           =   32
      Left            =   4200
      TabIndex        =   31
      Top             =   4590
      Width           =   75
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Index           =   31
      Left            =   4200
      TabIndex        =   30
      Top             =   4320
      Width           =   75
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Index           =   30
      Left            =   4200
      TabIndex        =   29
      Top             =   4065
      Width           =   75
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Index           =   18
      Left            =   4200
      TabIndex        =   28
      Top             =   810
      Width           =   75
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Index           =   19
      Left            =   4200
      TabIndex        =   27
      Top             =   1065
      Width           =   75
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Index           =   20
      Left            =   4200
      TabIndex        =   26
      Top             =   1350
      Width           =   75
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Index           =   21
      Left            =   4200
      TabIndex        =   25
      Top             =   1635
      Width           =   75
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Index           =   22
      Left            =   4200
      TabIndex        =   24
      Top             =   1905
      Width           =   75
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Index           =   23
      Left            =   4200
      TabIndex        =   23
      Top             =   2160
      Width           =   75
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Index           =   24
      Left            =   4200
      TabIndex        =   22
      Top             =   2445
      Width           =   75
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Index           =   25
      Left            =   4200
      TabIndex        =   21
      Top             =   2715
      Width           =   75
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Index           =   26
      Left            =   4200
      TabIndex        =   20
      Top             =   2985
      Width           =   75
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Index           =   27
      Left            =   4200
      TabIndex        =   19
      Top             =   3240
      Width           =   75
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Index           =   28
      Left            =   4200
      TabIndex        =   18
      Top             =   3510
      Width           =   75
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Index           =   29
      Left            =   4200
      TabIndex        =   17
      Top             =   3795
      Width           =   75
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Index           =   17
      Left            =   1815
      TabIndex        =   16
      Top             =   5145
      Width           =   75
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Index           =   16
      Left            =   1815
      TabIndex        =   15
      Top             =   4875
      Width           =   75
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Index           =   15
      Left            =   1815
      TabIndex        =   14
      Top             =   4590
      Width           =   75
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Index           =   14
      Left            =   1815
      TabIndex        =   13
      Top             =   4320
      Width           =   75
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Index           =   13
      Left            =   1815
      TabIndex        =   12
      Top             =   4065
      Width           =   75
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "90000000"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Index           =   12
      Left            =   1815
      TabIndex        =   11
      Top             =   3795
      Width           =   1320
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "82500000"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Index           =   11
      Left            =   1815
      TabIndex        =   10
      Top             =   3510
      Width           =   1320
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "75000000"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Index           =   10
      Left            =   1815
      TabIndex        =   9
      Top             =   3240
      Width           =   1320
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "67500000"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Index           =   9
      Left            =   1815
      TabIndex        =   8
      Top             =   2985
      Width           =   1320
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "60000000"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Index           =   8
      Left            =   1815
      TabIndex        =   7
      Top             =   2715
      Width           =   1320
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "52500000"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Index           =   7
      Left            =   1815
      TabIndex        =   6
      Top             =   2445
      Width           =   1320
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "45000000"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Index           =   6
      Left            =   1815
      TabIndex        =   5
      Top             =   2160
      Width           =   1320
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "37500000"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Index           =   5
      Left            =   1815
      TabIndex        =   4
      Top             =   1905
      Width           =   1320
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "30000000"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Index           =   4
      Left            =   1815
      TabIndex        =   3
      Top             =   1635
      Width           =   1320
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "22500000"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Index           =   3
      Left            =   1815
      TabIndex        =   2
      Top             =   1350
      Width           =   1320
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "15000000"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Index           =   2
      Left            =   1800
      TabIndex        =   1
      Top             =   1065
      Width           =   1320
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "7000000 Bs"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Index           =   1
      Left            =   1815
      TabIndex        =   0
      Top             =   810
      Width           =   1560
   End
End
Attribute VB_Name = "FrmInfPreguntas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Dim i As Integer
For i = 1 To 50
    Select Case i
        Case 1 'Preguntas Easy
            Label1(i).Caption = "0,15"
        Case 2
            Label1(i).Caption = "0,3"
        Case 3
            Label1(i).Caption = "0,45"
        Case 4
            Label1(i).Caption = "0,6"
        Case 5
            Label1(i).Caption = "0,75"
        Case 6
            Label1(i).Caption = "0,9"
        Case 7
            Label1(i).Caption = "1,05"
        Case 8
            Label1(i).Caption = "1,2"
        Case 9
            Label1(i).Caption = "1,35"
        Case 10
            Label1(i).Caption = "1,5"
        Case 11
            Label1(i).Caption = "1,65"
        Case 12
            Label1(i).Caption = "1,8"
        Case 13
            Label1(i).Caption = "1,95"
        Case 14
            Label1(i).Caption = "2,1"
        Case 15
            Label1(i).Caption = "2,25"
        Case 16
            Label1(i).Caption = "2,4"
        Case 17
            Label1(i).Caption = "2,55"
        Case 18
            Label1(i).Caption = "2,7"
        Case 19
            Label1(i).Caption = "2,85"
        Case 20
            Label1(i).Caption = "3"
        Case 21 'Preguntas Medium
            Label1(i).Caption = "3,4"
        Case 22
            Label1(i).Caption = "3,8"
        Case 23
            Label1(i).Caption = "4,2"
        Case 24
            Label1(i).Caption = "4,6"
        Case 25
            Label1(i).Caption = "5"
        Case 26
            Label1(i).Caption = "5,4"
        Case 27
            Label1(i).Caption = "5,8"
        Case 28
            Label1(i).Caption = "6,2"
        Case 29
            Label1(i).Caption = "6,6"
        Case 30
            Label1(i).Caption = "7"
        Case 31
            Label1(i).Caption = "7,4"
        Case 32
            Label1(i).Caption = "7,8"
        Case 33
            Label1(i).Caption = "8,2"
        Case 34
            Label1(i).Caption = "8,6"
        Case 35
            Label1(i).Caption = "9"
        Case 36
            Label1(i).Caption = "9,4"
        Case 37
            Label1(i).Caption = "9,8"
        Case 38
            Label1(i).Caption = "10,2"
        Case 39
            Label1(i).Caption = "10,6"
        Case 40
            Label1(i).Caption = "11"
        Case 41 'Preguntas Hard
            Label1(i).Caption = "11,5"
        Case 42
            Label1(i).Caption = "12"
        Case 43
            Label1(i).Caption = "13"
        Case 44
            Label1(i).Caption = "14"
        Case 45
            Label1(i).Caption = "15"
        Case 46
            Label1(i).Caption = "16"
        Case 47
            Label1(i).Caption = "17"
        Case 48
            Label1(i).Caption = "18"
        Case 49
            Label1(i).Caption = "19"
        Case 50
            Label1(i).Caption = "20"
    End Select
Next

End Sub


Private Sub Image1_Click()
Unload Me
End Sub

Private Sub ImgRetirarse_Click()
FrmRetirarse.Show 1
End Sub

Private Sub TmCambiar_Timer()
If Label1(Pregunta_Actual).ForeColor = vbWhite Then
    Label1(Pregunta_Actual).ForeColor = vbRed
Else
    Label1(Pregunta_Actual).ForeColor = vbWhite
End If
End Sub
