VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Begin VB.Form frmCR_Instituciones 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Instituciones Registras para Deducciones x Planilla"
   ClientHeight    =   5070
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7410
   Icon            =   "frmCR_Instituciones.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5070
   ScaleWidth      =   7410
   Begin FPSpread.vaSpread vGrid 
      Height          =   4935
      Left            =   0
      TabIndex        =   0
      Top             =   60
      Width           =   7335
      _Version        =   393216
      _ExtentX        =   12938
      _ExtentY        =   8705
      _StockProps     =   64
      BorderStyle     =   0
      DisplayRowHeaders=   0   'False
      EditEnterAction =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   498
      ScrollBars      =   2
      SpreadDesigner  =   "frmCR_Instituciones.frx":030A
      VisibleCols     =   500
      VisibleRows     =   500
   End
End
Attribute VB_Name = "frmCR_Instituciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vCodigo As String


Private Sub sbCargaGridLocal(vGrid As vaSpread, vGridMaxCol As Integer, strSQL As String)
Dim rs As New ADODB.Recordset, i As Integer

Me.MousePointer = vbHourglass

vGrid.MaxCols = vGridMaxCol
vGrid.MaxRows = 1

vGrid.Row = vGrid.MaxRows

rs.CursorLocation = adUseServer
rs.Open strSQL, glogon.Conection, adOpenForwardOnly

Do While Not rs.EOF
  vGrid.Row = vGrid.MaxRows
  For i = 1 To vGrid.MaxCols
    vGrid.Col = i
    Select Case i
      Case 3
        vGrid.Text = CStr(rs.Fields(i - 1).Value) ' fxgCntCuentaFormato(True, CStr(rs.Fields(i - 1).Value))
      Case Else
        vGrid.Text = CStr(rs.Fields(i - 1).Value)
    End Select
 
  Next i
  vGrid.MaxRows = vGrid.MaxRows + 1
  rs.MoveNext
Loop

rs.Close

Me.MousePointer = vbDefault

End Sub


Private Sub Form_Load()
Dim strSQL As String


strSQL = "select cod_institucion,descripcion,cod_cuenta from instituciones" _
       & " order by cod_institucion"
Call sbCargaGridLocal(vGrid, 3, strSQL)

Call Formularios(Me)
Call RefrescaTags(Me)

' vGrid.Enabled = cmdModifica.Enabled

End Sub


Private Function fxGuardar() As Long
Dim strSQL As String, rs As New ADODB.Recordset
'Guarda la información de la linea
'si es Insert devuelve el codigo, sino devuelve 0

On Error GoTo vError

fxGuardar = 0
vGrid.Row = vGrid.ActiveRow
vGrid.Col = 1


If vGrid.Text = "" Then
   vGrid.Col = 2
   strSQL = "insert instituciones(descripcion,cod_cuenta)" _
          & " values('" & vGrid.Text & "','"
   vGrid.Col = 3
   strSQL = strSQL & vGrid.Text & "')"
   
   glogon.Conection.Execute strSQL
   
   vGrid.Col = 1
   strSQL = "select max(cod_institucion) as Ultimo from instituciones"
   rs.Open strSQL, glogon.Conection, adOpenStatic
     vGrid.Text = rs!ultimo
   rs.Close
   
   Call Bitacora("Registra", "Institucion - Cod: " & vGrid.Text)
   
 Else 'Actualizar
    vGrid.Col = 2
    strSQL = "update instituciones set descripcion = '" & vGrid.Text
    vGrid.Col = 3
    strSQL = strSQL & "','" & vGrid.Text & "'"
    vGrid.Col = 1
    strSQL = " where cod_institucion = " & vGrid.Text
   
    glogon.Conection.Execute strSQL
    
    Call Bitacora("Modifica", "Institucion Cod: " & vGrid.Text)
    
End If


vGrid.Col = 1
fxGuardar = vGrid.Text

Exit Function
   
vError:
 MsgBox Err.Description, vbCritical
End Function

Private Sub vGrid_KeyDown(KeyCode As Integer, Shift As Integer)
Dim i As Long

If vGrid.ActiveCol = vGrid.MaxCols And (KeyCode = vbKeyReturn Or KeyCode = vbKeyTab) Then
  i = fxGuardar
  vGrid.Row = vGrid.ActiveRow
  vGrid.Col = 1
  If vGrid.MaxRows <= vGrid.ActiveRow Then
    vGrid.MaxRows = vGrid.MaxRows + 1
    vGrid.Row = vGrid.MaxRows
  End If
End If


If KeyCode = vbKeyF4 And vGrid.ActiveCol = 3 Then
   gBusquedas.Columna = "Descripcion"
   gBusquedas.Orden = "descripcion"
   gBusquedas.Consulta = "select Cod_Cuenta,Descripcion from cuentas"
   gBusquedas.Filtro = " and acepta_movimientos = 'S'"
   frmBusquedas.Show vbModal
   
   vGrid.Col = vGrid.ActiveCol
   vGrid.Row = vGrid.ActiveRow
   vGrid.Text = gBusquedas.Resultado
End If

If vGrid.ActiveCol = 1 And (KeyCode = vbKeyReturn Or KeyCode = vbKeyTab) Then
  vGrid.Col = vGrid.ActiveCol
  vGrid.Row = vGrid.ActiveRow
  vGrid.Text = vGrid.Text
End If

End Sub


