VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Begin VB.Form frmCR_ComitesParametros 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Parámetros de Comités de Resolución"
   ClientHeight    =   6180
   ClientLeft      =   48
   ClientTop       =   348
   ClientWidth     =   9648
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6180
   ScaleWidth      =   9648
   Begin FPSpreadADO.fpSpread vGrid 
      Height          =   4812
      Left            =   120
      TabIndex        =   0
      Top             =   1200
      Width           =   9372
      _Version        =   524288
      _ExtentX        =   16531
      _ExtentY        =   8488
      _StockProps     =   64
      BackColorStyle  =   1
      BorderStyle     =   0
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
      MaxCols         =   486
      MaxRows         =   498
      ScrollBarExtMode=   -1  'True
      ScrollBars      =   2
      SpreadDesigner  =   "frmCR_ComitesParametros.frx":0000
      VScrollSpecialType=   2
      AppearanceStyle =   1
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Parámetros de Comité de Resolución"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   16.2
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   480
      Index           =   2
      Left            =   1920
      TabIndex        =   1
      Top             =   300
      Width           =   7332
   End
   Begin VB.Image imgBanner 
      Height          =   1092
      Left            =   0
      Top             =   0
      Width           =   10812
   End
End
Attribute VB_Name = "frmCR_ComitesParametros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
Dim strSQL As String

vModulo = 3

vGrid.AppearanceStyle = fxGridStyle
Set imgBanner.Picture = frmContenedor.imgBanner_Mantenimiento.Picture

'Activar
Call Formularios(Me)
Call RefrescaTags(Me)
strSQL = "select cod_parametro, descripcion, valor from crd_comites_parametros order by cod_parametro"

Call sbCargaGrid(vGrid, 3, strSQL)

vGrid.MaxRows = vGrid.MaxRows - 1
End Sub


Private Function fxGuardar() As Long
Dim strSQL As String, rs As New ADODB.Recordset
'Guarda la información de la linea
'si es Insert devuelve el codigo, sino devuelve 0

On Error GoTo vError

fxGuardar = 0
  vGrid.Row = vGrid.ActiveRow
vGrid.col = 1


vGrid.col = 3
strSQL = "update CRD_COMITES_PARAMETROS set valor = '" & vGrid.Text & "'"
vGrid.col = 1
strSQL = strSQL & " where COD_PARAMETRO = '" & vGrid.Text & "'"
Call ConectionExecute(strSQL)

'TODO Activar

Call Bitacora("Modifica", "Parametro de comites de aprobación cod: " & vGrid.Text)

fxGuardar = 1

Exit Function

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Function

Private Sub vGrid_Click(ByVal col As Long, ByVal Row As Long)
vGrid.ToolTipText = ""
vGrid.col = 1
vGrid.Row = Row
'If (vGrid.Text = "001") Or (vGrid.Text = "002") Or (vGrid.Text = "003") _
'   Or (vGrid.Text = "005") Or (vGrid.Text = "006") Or (vGrid.Text = "007") Then
'   vGrid.Col = Col
'    vGrid.ToolTipText = "Presione F4 para consultar"
'End If

End Sub

Private Sub vGrid_KeyDown(KeyCode As Integer, Shift As Integer)
Dim i As Integer

If vGrid.ActiveCol = vGrid.MaxCols And (KeyCode = 13 Or KeyCode = vbKeyTab) Then
  i = fxGuardar
  If i = 0 Then Exit Sub
  vGrid.Row = vGrid.ActiveRow
End If

'If KeyCode = vbKeyF4 Then
'    vGrid.Row = vGrid.ActiveRow
'    vGrid.Col = 1
'    If (vGrid.Text = "001") Or (vGrid.Text = "002") Or (vGrid.Text = "003") Or (vGrid.Text = "009") Or (vGrid.Text = "010") Then
'        gBusquedas.Resultado = ""
'        Call sbgCntCuentaConsulta("D")
'        If gBusquedas.Resultado = "" Then Exit Sub
'        vGrid.Col = 3
'        vGrid.Row = vGrid.ActiveRow
'        vGrid.Text = gBusquedas.Resultado
'        i = fxGuardar
'        If i = 0 Then Exit Sub
'            vGrid.Row = vGrid.ActiveRow
'    ElseIf (vGrid.Text = "005") Or (vGrid.Text = "006") Or (vGrid.Text = "007") Or (vGrid.Text = "009") Then
'        gBusquedas.Resultado = ""
'        Call sbBusqueda(vGrid.ActiveRow, 3)
'    End If
'End If

End Sub

Private Sub sbBusqueda(ByVal fil As Integer, ByVal col As Integer)

On Error GoTo vError
Dim i As Integer
    gBusquedas.Convertir = "N"

    gBusquedas.Consulta = "SELECT COD_PARAMETRO, DESCRIPCION FROM "
    gBusquedas.Columna = "DESCRIPCION"
    gBusquedas.Orden = "DESCRIPCION"
    frmBusquedas.Show vbModal
    If gBusquedas.Resultado = "" Then Exit Sub
    vGrid.col = 3
    vGrid.Row = fil
    vGrid.Text = Trim(gBusquedas.Resultado)
    i = fxGuardar
    If i = 0 Then Exit Sub
    vGrid.Row = vGrid.ActiveRow
          
    gBusquedas.Consulta = ""
    gBusquedas.Columna = ""
    gBusquedas.Orden = ""
    gBusquedas.Resultado = ""
  
Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub





