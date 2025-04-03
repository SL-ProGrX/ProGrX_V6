VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Begin VB.Form frmCajas_Parametros 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Parámetros Cajas"
   ClientHeight    =   6855
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8145
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6855
   ScaleWidth      =   8145
   Begin FPSpreadADO.fpSpread vGrid 
      Height          =   5412
      Left            =   120
      TabIndex        =   0
      Top             =   1320
      Width           =   7932
      _Version        =   524288
      _ExtentX        =   13991
      _ExtentY        =   9546
      _StockProps     =   64
      BackColorStyle  =   1
      BorderStyle     =   0
      ColHeaderDisplay=   0
      EditEnterAction =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   3
      ScrollBars      =   2
      SpreadDesigner  =   "frmCajas_Parametros.frx":0000
      VScrollSpecialType=   2
      AppearanceStyle =   1
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Parámetros de Cajas"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   1680
      TabIndex        =   1
      Top             =   360
      Width           =   6255
   End
   Begin VB.Image imgBanner 
      Height          =   1215
      Left            =   0
      Top             =   0
      Width           =   10815
   End
End
Attribute VB_Name = "frmCajas_Parametros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Activate()
vModulo = 5
End Sub


Private Sub Form_Load()
Dim strSQL As String

vModulo = 5


Set imgBanner.Picture = frmContenedor.imgBanner_Mantenimiento.Picture

Call Formularios(Me)
Call RefrescaTags(Me)

On Error GoTo vError

strSQL = "exec spCajas_Parametros"
Call ConectionExecute(strSQL)

strSQL = "select cod_parametro,descripcion,valor from cajas_parametros where visible = 'S'" _
      & " order by cod_parametro"
Call sbCargaGrid(vGrid, 3, strSQL)

vGrid.MaxRows = vGrid.MaxRows - 1

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
  
End Sub


Private Function fxGuardar() As Long
Dim strSQL As String, rs As New ADODB.Recordset
Dim vValor As String
'Guarda la información de la linea
'si es Insert devuelve el codigo, sino devuelve 0

On Error GoTo vError

fxGuardar = 0
  vGrid.Row = vGrid.ActiveRow
vGrid.col = 1


vGrid.col = 3
vValor = vGrid.Text

strSQL = "update cajas_parametros set valor = '" & vGrid.Text & "'"
vGrid.col = 1
strSQL = strSQL & " where cod_parametro = '" & vGrid.Text & "'"
Call ConectionExecute(strSQL)

Call Bitacora("Modifica", "Parámetros de Cajas (Parámetro .:" & vGrid.Text & ")  Valor .:" & vValor)

fxGuardar = 1

Exit Function

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Function

Private Sub vGrid_KeyDown(KeyCode As Integer, Shift As Integer)
Dim i As Integer

If vGrid.ActiveCol = vGrid.MaxCols And (KeyCode = 13 Or KeyCode = vbKeyTab) Then
  i = fxGuardar
  If i = 0 Then Exit Sub
  vGrid.Row = vGrid.ActiveRow
End If

End Sub

