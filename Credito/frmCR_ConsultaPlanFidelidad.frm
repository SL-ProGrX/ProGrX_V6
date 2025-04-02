VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCR_ConsultaPlanFidelidad 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   6750
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   10065
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6750
   ScaleWidth      =   10065
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton OptSorteo 
      Caption         =   "Provinciales"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   1
      Left            =   3360
      TabIndex        =   7
      Top             =   960
      Width           =   1455
   End
   Begin VB.OptionButton OptSorteo 
      Caption         =   "Generales"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   0
      Left            =   2040
      TabIndex        =   6
      Top             =   960
      Value           =   -1  'True
      Width           =   1215
   End
   Begin MSComctlLib.ListView lsw 
      Height          =   4455
      Left            =   240
      TabIndex        =   2
      Top             =   1680
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   7858
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      FlatScrollBar   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      HotTracking     =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Números"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.ListView lswSorteos 
      Height          =   4455
      Left            =   3240
      TabIndex        =   4
      Top             =   1680
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   7858
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FlatScrollBar   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      HotTracking     =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Sorteo"
         Object.Width           =   1658
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Text            =   "Fecha"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "Número Ganador"
         Object.Width           =   4304
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Premio"
         Object.Width           =   6068
      EndProperty
   End
   Begin VB.Label Label2 
      Caption         =   "Tipos de Sorteos .:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   8
      Top             =   960
      Width           =   1695
   End
   Begin VB.Label lblPremio 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   6240
      Width           =   9615
   End
   Begin VB.Label Label2 
      Caption         =   "Sorteos realizados y números favorecidos..:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   3240
      TabIndex        =   3
      Top             =   1320
      Width           =   3495
   End
   Begin VB.Label Label2 
      Caption         =   "Números asignados (Acciones)..:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   1
      Top             =   1320
      Width           =   2775
   End
   Begin VB.Label lblNombre 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   9615
   End
End
Attribute VB_Name = "frmCR_ConsultaPlanFidelidad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub sbAcciones()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListItem, vTipo As String

On Error GoTo vError

Select Case True
   Case OptSorteo.Item(0).Value 'General
      vTipo = "G"
   Case OptSorteo.Item(1).Value 'Provincial
      vTipo = "P"
End Select

lsw.ListItems.Clear
lswSorteos.ListItems.Clear
lblPremio.Caption = "Lo sentimos no ha sido favorecido!"

strSQL = "exec spAFI_SorteoAccionesConsulta '" & lblNombre.Tag & "','" & vTipo & "'"
Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
  Set itmX = lsw.ListItems.Add(, , rs!Numero)
  rs.MoveNext
Loop
rs.Close


strSQL = "exec spAFI_SorteoNumerosFavorecidos '" & vTipo & "'"
Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
  Set itmX = lswSorteos.ListItems.Add(, , rs!cod_Sorteo)
      itmX.SubItems(1) = Format(rs!Fecha_Sorteo, "dd/mm/yyyy")
      itmX.SubItems(2) = rs!Numero
      itmX.SubItems(3) = rs!Premio
      
      If Trim(rs!Cedula) = Trim(lblNombre.Tag) Then
         itmX.Bold = True
         itmX.ForeColor = vbBlue
         
         lblPremio.Caption = "Felicidades! Premio.: " & rs!Premio
      End If
  rs.MoveNext
Loop
rs.Close

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub Form_Load()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError


strSQL = "select cedula,nombre from socios where cedula = '" & GLOBALES.gTag & "'"
Call OpenRecordSet(rs, strSQL)
If Not rs.EOF And Not rs.BOF Then
  lblNombre.Caption = Trim(rs!Cedula) & " - " & rs!Nombre
  lblNombre.Tag = rs!Cedula
End If
rs.Close

Call sbAcciones

Exit Sub

vError:


End Sub

Private Sub OptSorteo_Click(Index As Integer)
  Call sbAcciones
End Sub
