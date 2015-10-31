VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmPreguntas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Preguntas"
   ClientHeight    =   8130
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12870
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8130
   ScaleWidth      =   12870
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox TxtCorrecta 
      Height          =   405
      Left            =   360
      MaxLength       =   50
      TabIndex        =   18
      Top             =   3480
      Width           =   12135
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Volver"
      Height          =   375
      Left            =   10920
      TabIndex        =   17
      Top             =   7560
      Width           =   1575
   End
   Begin VB.CommandButton CmdLimpiar 
      Caption         =   "&Limpiar"
      Height          =   375
      Left            =   9120
      TabIndex        =   16
      Top             =   7560
      Width           =   1575
   End
   Begin VB.CommandButton CmdEliminar 
      Caption         =   "&Eliminar Pregunta"
      Height          =   375
      Left            =   7320
      TabIndex        =   15
      Top             =   7560
      Width           =   1575
   End
   Begin VB.CommandButton CmdModificar 
      Caption         =   "&Modificar Pregunta"
      Height          =   375
      Left            =   5520
      TabIndex        =   14
      Top             =   7560
      Width           =   1575
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   3375
      Left            =   360
      TabIndex        =   12
      Top             =   4080
      Width           =   12135
      _ExtentX        =   21405
      _ExtentY        =   5953
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.TextBox TxtRespuesta4 
      Height          =   405
      Left            =   6480
      MaxLength       =   50
      TabIndex        =   11
      Top             =   2760
      Width           =   6015
   End
   Begin VB.TextBox TxtRespuesta3 
      Height          =   405
      Left            =   360
      MaxLength       =   50
      TabIndex        =   10
      Top             =   2760
      Width           =   6015
   End
   Begin VB.TextBox TxtRespuesta2 
      Height          =   405
      Left            =   6480
      MaxLength       =   50
      TabIndex        =   7
      Top             =   2040
      Width           =   6015
   End
   Begin VB.TextBox TxtRespuesta1 
      Height          =   405
      Left            =   360
      MaxLength       =   50
      TabIndex        =   6
      Top             =   2040
      Width           =   6015
   End
   Begin VB.TextBox TxtPregunta 
      Height          =   615
      Left            =   360
      MaxLength       =   135
      TabIndex        =   2
      Top             =   1080
      Width           =   12135
   End
   Begin VB.ComboBox CmbTabla 
      Height          =   315
      ItemData        =   "FrmPreguntas.frx":0000
      Left            =   960
      List            =   "FrmPreguntas.frx":000D
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   360
      Width           =   1695
   End
   Begin VB.CommandButton CmdAgregar 
      Caption         =   "&Agregar Pregunta"
      Height          =   375
      Left            =   3720
      TabIndex        =   13
      Top             =   7560
      Width           =   1575
   End
   Begin VB.Label Label7 
      Caption         =   "Respuesta Correcta"
      Height          =   375
      Left            =   360
      TabIndex        =   19
      Top             =   3240
      Width           =   2535
   End
   Begin VB.Label Label6 
      Caption         =   "Respuesta 4"
      Height          =   375
      Left            =   6480
      TabIndex        =   9
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "Respuesta 3"
      Height          =   375
      Left            =   360
      TabIndex        =   8
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "Respuesta 2"
      Height          =   375
      Left            =   6480
      TabIndex        =   5
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Respuesta 1"
      Height          =   375
      Left            =   360
      TabIndex        =   4
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Pregunta"
      Height          =   375
      Left            =   360
      TabIndex        =   3
      Top             =   840
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Tabla"
      Height          =   255
      Left            =   360
      TabIndex        =   1
      Top             =   360
      Width           =   1095
   End
End
Attribute VB_Name = "FrmPreguntas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cod As Integer
Private Sub CmbTabla_Click()
Call LlenarLista
End Sub

Private Sub CmdAgregar_Click()
If preguntaValida Then
    Sql.CommandText = Modulo.Añadir(CmbTabla.Text, "#" & Modulo.ObtieneCodigo(CmbTabla.Text), TxtPregunta.Text, TxtRespuesta1.Text, TxtRespuesta2.Text, TxtRespuesta3.Text, TxtRespuesta4.Text, TxtCorrecta.Text, "N")
    Set RS = Sql.Execute
    MsgBox "Se agrego correctamente la pregunta '" & TxtPregunta.Text & "'", vbQuestion, "AprendeIUTAJS"
    Call CmdLimpiar_Click
    Call LlenarLista
End If
End Sub
Private Function preguntaValida() As Boolean
preguntaValida = False
If TxtPregunta.Text = "" Then
    MsgBox "Debe escribir una pregunta", vbQuestion, "AprendeIUTAJS"
    TxtPregunta.SetFocus
    Exit Function
End If
If TxtRespuesta1.Text = "" Then
    MsgBox "Debe escribir una respuesta", vbQuestion, "AprendeIUTAJS"
    TxtRespuesta1.SetFocus
    Exit Function
End If
If TxtRespuesta2.Text = "" Then
    MsgBox "Debe escribir una respuesta", vbQuestion, "AprendeIUTAJS"
    TxtRespuesta2.SetFocus
    Exit Function
End If
If TxtRespuesta3.Text = "" Then
    MsgBox "Debe escribir una respuesta", vbQuestion, "AprendeIUTAJS"
    TxtRespuesta3.SetFocus
    Exit Function
End If
If TxtRespuesta4.Text = "" Then
    MsgBox "Debe escribir una respuesta", vbQuestion, "AprendeIUTAJS"
    TxtRespuesta4.SetFocus
    Exit Function
End If
Select Case TxtCorrecta.Text
    Case ""
        MsgBox "Debe escribir la respuesta correcta", vbQuestion, "AprendeIUTAJS"
        TxtCorrecta.SetFocus
        Exit Function
    Case TxtRespuesta1.Text, TxtRespuesta2.Text, TxtRespuesta3.Text, TxtRespuesta4.Text
    Case Else
        MsgBox "La respuesta correcta debe coincidir con alguna de las respuestas anteriores", vbQuestion, "AprendeIUTAJS"
        TxtCorrecta.SetFocus
        Exit Function
End Select
preguntaValida = True
End Function

Private Sub CmdEliminar_Click()
If cod = 0 Then
    MsgBox "Debe seleccionar una pregunta de la lista", vbQuestion, "AprendeIUTAJS"
Else
    Sql.CommandText = "select cod from " & CmbTabla.Text 'Cuento el numero de registro guardados en la tabla
    Set RS = Sql.Execute 'Para saber la cantidad de preguntas que hay
    If RS.EOF Then
        MsgBox "La pregunta que desea eliminar no existe", vbQuestion, "AprendeIUTAJS"
    Else
        Sql.CommandText = Modulo.Eliminar(CmbTabla.Text, "#" & cod)
        Set RS = Sql.Execute
        MsgBox "Pregunta '" & TxtPregunta.Text & "' eliminada correctamente", vbQuestion, "AprendeIUTAJS"
        Call CmdLimpiar
        Call LlenarLista
    End If
End If
End Sub

Private Sub CmdLimpiar_Click()
cod = 0
TxtCorrecta.Text = ""
TxtPregunta.Text = ""
TxtRespuesta1.Text = ""
TxtRespuesta2.Text = ""
TxtRespuesta3.Text = ""
TxtRespuesta4.Text = ""
End Sub

Private Sub CmdModificar_Click()
If cod = 0 Then
    MsgBox "Debe seleccionar una pregunta de la lista", vbQuestion, "AprendeIUTAJS"
Else
    Sql.CommandText = "select cod from " & CmbTabla.Text
    Set RS = Sql.Execute
    If RS.EOF Then
        MsgBox "La pregunta que desea modificar no existe", vbQuestion, "AprendeIUTAJS"
    Else
        Sql.CommandText = Modulo.Modificar(CmbTabla.Text, TxtPregunta.Text, TxtRespuesta1.Text, TxtRespuesta2.Text, TxtRespuesta3.Text, TxtRespuesta4.Text, TxtCorrecta.Text, "N", "#" & cod)
        Set RS = Sql.Execute
        MsgBox "Pregunta '" & TxtPregunta.Text & "' modificada correctamente", vbQuestion, "AprendeIUTAJS"
        Call CmdLimpiar_Click
        Call LlenarLista
    End If
End If
End Sub

Private Sub Command3_Click()
Unload Me
End Sub

Private Sub Form_Load()
Call AbrirBD 'Habro la base de datos
Call DefinirLista
Call LlenarLista
cod = 0
End Sub
Private Sub DefinirLista()
With ListView1
    .View = lvwReport
    .GridLines = True
    .LabelEdit = lvwManual
    .ColumnHeaders.Add , , "Codigo"
    .ColumnHeaders.Add , , "Pregunta"
    .ColumnHeaders.Add , , "Respuesta1"
    .ColumnHeaders.Add , , "Respuesta2"
    .ColumnHeaders.Add , , "Respuesta3"
    .ColumnHeaders.Add , , "Respuesta4"
    .ColumnHeaders.Add , , "Respuesta Correcta"
End With
End Sub
Private Sub LlenarLista()
'Call cabecera
If CmbTabla.ListIndex >= 0 Then
    Sql.CommandText = "select * from " & CmbTabla.Text & " order by cod"
    Set RS = Sql.Execute
    ListView1.ListItems.Clear
    With RS
        If .EOF = True Then
        Exit Sub
    End If
    'Mostramos los datos ordenados por codigo
    
    .MoveFirst
    Do While Not .EOF
        Set TLI = ListView1.ListItems.Add(, , RS!cod)
        TLI.SubItems(1) = RS!Preguntas
        TLI.SubItems(2) = RS!respuesta1
        TLI.SubItems(3) = RS!respuesta2
        TLI.SubItems(4) = RS!respuesta3
        TLI.SubItems(5) = RS!respuesta4
        TLI.SubItems(6) = RS!Correcta
    .MoveNext
    Loop
    End With
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
Load FrmPrincipal
FrmPrincipal.Show
End Sub

Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo Lista
Sql.CommandText = "select * from " & CmbTabla.Text & " where cod=" & ListView1.selectedItem
Set RS = Sql.Execute
If Not RS.EOF Then
    cod = RS!cod
    TxtPregunta.Text = Trim(RS!Preguntas)
    TxtRespuesta1.Text = Trim(RS!respuesta1)
    TxtRespuesta2.Text = Trim(RS!respuesta2)
    TxtRespuesta3.Text = Trim(RS!respuesta3)
    TxtRespuesta4.Text = Trim(RS!respuesta4)
    TxtCorrecta.Text = Trim(RS!Correcta)
End If
Lista:
'If Err.Number <> 0 Then
'MsgBox "No hay nada para seleccionar", vbCritical, "Error"
'End If

End Sub
