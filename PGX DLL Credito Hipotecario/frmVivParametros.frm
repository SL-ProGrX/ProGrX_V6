VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Begin VB.Form frmVivParametros 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Parámetros Generales"
   ClientHeight    =   6672
   ClientLeft      =   48
   ClientTop       =   432
   ClientWidth     =   9228
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6672
   ScaleWidth      =   9228
   StartUpPosition =   2  'CenterScreen
   Begin FPSpreadADO.fpSpread vGrid 
      Height          =   5052
      Left            =   240
      TabIndex        =   1
      Top             =   1320
      Width           =   8892
      _Version        =   524288
      _ExtentX        =   15685
      _ExtentY        =   8911
      _StockProps     =   64
      BackColorStyle  =   1
      BorderStyle     =   0
      DisplayRowHeaders=   0   'False
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
      SpreadDesigner  =   "frmVivParametros.frx":0000
      VScrollSpecialType=   2
      AppearanceStyle =   1
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Parámetros de Crédito Hipotecario"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   13.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   492
      Left            =   1800
      TabIndex        =   0
      Top             =   360
      Width           =   6252
   End
   Begin VB.Image imgBanner 
      Height          =   1095
      Left            =   0
      Top             =   0
      Width           =   11655
   End
End
Attribute VB_Name = "frmVivParametros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Activate()
vModulo = 3
End Sub


Private Sub Form_Load()
Dim strSQL As String

vModulo = 3

vGrid.AppearanceStyle = fxGridStyle
imgBanner.Picture = frmContenedor.imgBanner_Mantenimiento.Picture


strSQL = "exec spCRDVivParametros"
Call ConectionExecute(strSQL)

strSQL = "select CodigoParametro,descripcion,valor from ViviendaParametros" _
                  & " order by CodigoParametro"

Call sbCargaGrid(vGrid, 3, strSQL)
'Call sbgCntParametros

vGrid.MaxRows = vGrid.MaxRows - 1

Call Formularios(Me)
Call RefrescaTags(Me)

End Sub


Private Function fxGuardar() As Long
Dim strSQL As String, rs As New ADODB.Recordset
'Guarda la información de la linea
'si es Insert devuelve el codigo, sino devuelve 0

On Error GoTo vError

fxGuardar = 0
  vGrid.Row = vGrid.ActiveRow
vGrid.Col = 1


vGrid.Col = 3
strSQL = "update ViviendaParametros set valor = '" & vGrid.Text & "'"
vGrid.Col = 1
strSQL = strSQL & " where CodigoParametro = '" & vGrid.Text & "'"
Call ConectionExecute(strSQL)

'TODO Activar

Call Bitacora("Modifica", "Parametro de vivienda cod: " & vGrid.Text)

fxGuardar = 1

Exit Function

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Function


Private Sub vGrid_Click(ByVal Col As Long, ByVal Row As Long)
vGrid.ToolTipText = ""
vGrid.Col = 1
vGrid.Row = Row
If (vGrid.Text = "001") Or (vGrid.Text = "002") Or (vGrid.Text = "003") _
   Or (vGrid.Text = "005") Or (vGrid.Text = "006") Or (vGrid.Text = "007") Then
   vGrid.Col = Col
    vGrid.ToolTipText = "Presione F4 para consultar"
End If

End Sub

Private Sub vGrid_KeyDown(KeyCode As Integer, Shift As Integer)
Dim i As Integer

If vGrid.ActiveCol = vGrid.MaxCols And (KeyCode = 13 Or KeyCode = vbKeyTab) Then
  i = fxGuardar
  If i = 0 Then Exit Sub
  vGrid.Row = vGrid.ActiveRow
End If

If KeyCode = vbKeyF4 Then
    vGrid.Row = vGrid.ActiveRow
    vGrid.Col = 1
    If (vGrid.Text = "001") Or (vGrid.Text = "002") Or (vGrid.Text = "003") Or (vGrid.Text = "009") Or (vGrid.Text = "010") Then
        gBusquedas.Resultado = ""
        Call sbgCntCuentaConsulta("D")
        If gBusquedas.Resultado = "" Then Exit Sub
        vGrid.Col = 3
        vGrid.Row = vGrid.ActiveRow
        vGrid.Text = gBusquedas.Resultado
        i = fxGuardar
        If i = 0 Then Exit Sub
            vGrid.Row = vGrid.ActiveRow
    ElseIf (vGrid.Text = "005") Or (vGrid.Text = "006") Or (vGrid.Text = "007") Or (vGrid.Text = "009") Then
        gBusquedas.Resultado = ""
        Call sbBusqueda(vGrid.ActiveRow, 3)
    End If
End If

End Sub

Private Sub sbBusqueda(ByVal fil As Integer, ByVal Col As Integer)

On Error GoTo vError
Dim i As Integer
    gBusquedas.Convertir = "N"

    gBusquedas.Consulta = "SELECT Codigo, Descripcion FROM ViviendaTiposDesembolsos"
    gBusquedas.Columna = "Descripcion"
    gBusquedas.Orden = "Descripcion"
    frmBusquedas.Show vbModal
    If gBusquedas.Resultado = "" Then Exit Sub
    vGrid.Col = 3
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



