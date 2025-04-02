VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Begin VB.Form frmFSL_ModificaPorcentaje 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   7320
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   7080
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7320
   ScaleWidth      =   7080
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin FPSpreadADO.fpSpread vgCreditos 
      Height          =   5535
      Left            =   480
      TabIndex        =   1
      Top             =   1440
      Width           =   5895
      _Version        =   524288
      _ExtentX        =   10398
      _ExtentY        =   9763
      _StockProps     =   64
      BackColorStyle  =   1
      BorderStyle     =   0
      EditEnterAction =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   4
      ScrollBars      =   2
      SpreadDesigner  =   "frmFSL_ModificaPorcentaje.frx":0000
      VScrollSpecialType=   2
      AppearanceStyle =   1
   End
   Begin VB.Label lblAsociado 
      Height          =   255
      Left            =   600
      TabIndex        =   4
      Top             =   960
      Width           =   6015
   End
   Begin VB.Label lblExpediente 
      Height          =   255
      Left            =   1560
      TabIndex        =   3
      Top             =   600
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Expediente:"
      Height          =   255
      Left            =   600
      TabIndex        =   2
      Top             =   600
      Width           =   855
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      MouseIcon       =   "frmFSL_ModificaPorcentaje.frx":06D5
      Picture         =   "frmFSL_ModificaPorcentaje.frx":15837
      Top             =   0
      Width           =   480
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Modifica Porcentaje por aplicar en FOSOL"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1320
      TabIndex        =   0
      Top             =   120
      Width           =   4815
   End
End
Attribute VB_Name = "frmFSL_ModificaPorcentaje"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSQL As String
Dim rs As New ADODB.Recordset
Dim vExpediente As Long, vOperacion As Long
Dim vPorcentaje As Double

Private Sub Form_Activate()
 vModulo = 22
End Sub

Private Sub Form_Load()
  vModulo = 22
  vExpediente = GLOBALES.gTag3
  lblExpediente.Caption = GLOBALES.gTag3
  lblAsociado.Caption = GLOBALES.gTag & " " & GLOBALES.gTag2
  Call sbCargaGrid
End Sub

Private Sub sbModificaPorcentaje(ByVal vOperacion As Long, ByVal vPorcentaje As Double)
  strSQL = "Update FSL_EXPEDIENTES_DETALLE set PORCENTAJE = " & vPorcentaje & " " _
         & " where COD_EXPEDIENTE = " & vExpediente & " and ID_SOLICITUD=" & vOperacion & ""
  glogon.Conection.Execute strSQL
End Sub

Private Sub sbCargaGrid()
Dim i As Integer
On Error GoTo vError

With vgCreditos
.MaxRows = 1
.Row = .MaxRows
For i = 1 To .MaxCols
  .Col = i
  .Text = ""
Next i

strSQL = "Select ID_SOLICITUD,TOTAL_DEUDA_P,PORCENTAJE from FSL_EXPEDIENTES_DETALLE" _
       & " where COD_EXPEDIENTE = " & vExpediente & ""
rs.Open strSQL, glogon.Conection, adOpenForwardOnly

Do While Not rs.EOF

 .Col = 1
 .Text = Format(rs!ID_SOLICITUD, "standard")
 .Col = 2
 .Text = Format(rs!TOTAL_DEUDA_P, "standard")
 .Col = 3
 .Value = Format(rs!Porcentaje, "standard")
 
 .MaxRows = .MaxRows + 1
 .Row = .MaxRows
 rs.MoveNext
Loop
rs.Close
.Row = .MaxRows
End With

Exit Sub
vError:
   MsgBox Err.Description, vbExclamation

End Sub

Private Sub vgCreditos_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo vError
With vgCreditos
.Row = .ActiveRow
If .ActiveCol = 4 And KeyCode = vbKeyReturn Then
   .Col = 1
   vOperacion = .Text
   .Col = 4
   If .Text <> "" Then
     vPorcentaje = .Text
   End If
   
   Call sbModificaPorcentaje(vOperacion, vPorcentaje)
   
'       strSQL = "exec spSifRegistraTags '" & txtCedula.Text & "','" & txtExpediente.Text & "', " _
'              & " '','A04','" & glogon.Usuario & "','" & txtObservaciones.Text & "'," & txtExpediente.Text & ""
'
'       glogon.Conection.Execute strSQL
   
End If
End With

Exit Sub
vError:
   MsgBox Err.Description, vbExclamation

End Sub
