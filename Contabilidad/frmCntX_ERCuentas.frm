VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Begin VB.Form frmCntX_ERCuentas 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Especificación Saldos finales de Inventario para Cierres Periodicos"
   ClientHeight    =   5475
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   9870
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5475
   ScaleWidth      =   9870
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin FPSpreadADO.fpSpread vGrid 
      Height          =   5055
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   9615
      _Version        =   524288
      _ExtentX        =   16960
      _ExtentY        =   8916
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
      MaxCols         =   495
      ScrollBars      =   2
      SpreadDesigner  =   "frmCntX_ERCuentas.frx":0000
      VScrollSpecial  =   -1  'True
      VScrollSpecialType=   2
      AppearanceStyle =   1
   End
End
Attribute VB_Name = "frmCntX_ERCuentas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub sbCargaGridLocal(vGrid As Object, vGridMaxCol As Integer, strSQL As String)
Dim rs As New ADODB.Recordset, i As Integer

Screen.MousePointer = vbHourglass

On Error GoTo vError

vGrid.MaxCols = vGridMaxCol
vGrid.MaxRows = 1
vGrid.Row = vGrid.MaxRows
'For i = 1 To vGrid.MaxCols
' vGrid.Col = i
' vGrid.Text = ""
'Next i

Call OpenRecordSet(rs, strSQL, 0)
Do While Not rs.EOF
  vGrid.Row = vGrid.MaxRows
  For i = 1 To vGrid.MaxCols
    Select Case i
      Case 3
        vGrid.Col = i
        vGrid.Text = fxCntX_CuentaFormato(True, CStr(rs.Fields(i - 1).Value))
      Case Else
        vGrid.Col = i
        vGrid.Text = CStr(rs.Fields(i - 1).Value)
     End Select
  Next i
  vGrid.MaxRows = vGrid.MaxRows + 1
  rs.MoveNext
Loop
rs.Close

vError:

Screen.MousePointer = vbDefault

End Sub


Private Sub Form_Load()
Dim strSQL As String, rs As New ADODB.Recordset


vGrid.AppearanceStyle = fxGridStyle


'strSQL = "select inicio_anio,inicio_mes,corte_anio,corte_mes,descripcion" _
'       & ",cuenta_ganper,cuenta_utilidad,cuenta_imprenta,impuesto_renta from cierres where cod_contabilidad = " & gCntX_Parametros.CodigoConta

strSQL = "Select I.anio,I.mes,I.cod_cuenta,C.descripcion,I.saldo_final" _
       & " from CntX_Inv_Periodico I inner join CntX_Cuentas C on I.cod_cuenta = C.cod_cuenta" _
       & " and I.cod_contabilidad = C.cod_contabilidad" _
       & " where I.cod_contabilidad = " & gCntX_Parametros.CodigoConta
Call sbCargaGridLocal(vGrid, 5, strSQL)

'Call Formularios(Me)
'Call RefrescaTags(Me)
'If tlb.Buttons(1).Enabled = False Then vGrid.Enabled = False

End Sub

Private Function fxVerificaCuentasLocal(strCuenta As String) As Boolean
Dim rsX As New ADODB.Recordset, strSQL As String

fxVerificaCuentasLocal = False

strSQL = "select isnull(count(*),0) as Total " _
       & " from CntX_Cuentas C inner join CntX_Tipos_Cuentas T on C.tipo_cuenta = T.tipo_cuenta " _
       & " and C.cod_contabilidad = T.cod_contabilidad" _
       & " where C.cod_contabilidad = " & gCntX_Parametros.CodigoConta & " and C.cod_cuenta = '" & fxCntX_CuentaFormato(False, strCuenta) _
       & "' and T.clasificacion = 'A'" 'and C.acepta_movimientos = 'S'
Call OpenRecordSet(rsX, strSQL, 0)
If rsX!Total = 1 Then fxVerificaCuentasLocal = True
rsX.Close
End Function

Private Function fxGuardar() As Long
Dim strSQL As String, rs As New ADODB.Recordset
'Guarda la información de la linea
'si es Insert devuelve el codigo, sino devuelve 0

On Error GoTo vError

fxGuardar = 0
vGrid.Row = vGrid.ActiveRow

vGrid.Col = 3
If Not fxVerificaCuentasLocal(fxCntX_CuentaFormato(False, vGrid.Text)) Then
  Me.MousePointer = vbDefault
  MsgBox "La cuenta especificada no es de Activo, o No Acepta Movimientos o no Existe...", vbExclamation
  Exit Function
End If

vGrid.Col = 1
strSQL = "select isnull(count(*),0) as Existe" _
      & " from CntX_Inv_Periodico where cod_contabilidad =" & gCntX_Parametros.CodigoConta _
       & " and anio = " & vGrid.Text
vGrid.Col = 2
strSQL = strSQL & " and mes = " & vGrid.Text
vGrid.Col = 3
strSQL = strSQL & " and cod_cuenta = " & fxCntX_CuentaFormato(False, vGrid.Text)

Call OpenRecordSet(rs, strSQL, 0)

If rs!Existe = 0 Then
    vGrid.Col = 1

    strSQL = "insert into CntX_Inv_Periodico(cod_contabilidad,anio,mes,cod_cuenta,saldo_final)" _
           & " values(" & gCntX_Parametros.CodigoConta & ","
    vGrid.Col = 1
    strSQL = strSQL & vGrid.Text & ","
    vGrid.Col = 2
    strSQL = strSQL & vGrid.Text & ",'"
    vGrid.Col = 3
    strSQL = strSQL & fxCntX_CuentaFormato(False, vGrid.Text) & "',"
    vGrid.Col = 5
    strSQL = strSQL & CCur(vGrid.Text) & ")"
    
    Call ConectionExecute(strSQL, 0)
  
    vGrid.Col = 3
    Call Bitacora("Registra", "Saldo Periodico Cuenta : " & fxCntX_CuentaFormato(True, vGrid.Text) & " Conta." & gCntX_Parametros.CodigoConta)
  
   Else 'Actualizar
       
       vGrid.Col = 5
       strSQL = "update CntX_Inv_Periodico set saldo_final = " & CCur(vGrid.Text)
       vGrid.Col = 1
       strSQL = strSQL & " where anio = " & vGrid.Text & " and mes = "
       vGrid.Col = 2
       strSQL = strSQL & vGrid.Text & " and cod_cuenta = '"
       vGrid.Col = 3
       strSQL = strSQL & fxCntX_CuentaFormato(False, vGrid.Text) & "' and cod_contabilidad = " _
              & gCntX_Parametros.CodigoConta
      
      Call ConectionExecute(strSQL, 0)
    
      vGrid.Col = 3
      Call Bitacora("Modifica", "Saldo Periodico Cuenta : " & fxCntX_CuentaFormato(True, vGrid.Text) & " Conta." & gCntX_Parametros.CodigoConta)
       
 End If

 rs.Close
 vGrid.Col = 1
 fxGuardar = 1
 
Exit Function

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Function

Private Sub sbBorrar()
Dim i As Integer, strSQL As String

i = MsgBox("Esta Seguro que desea borrar este registro..." _
           & vbCrLf & vbCrLf, vbYesNo)
If i = vbYes Then
   vGrid.Row = vGrid.ActiveRow
   vGrid.Col = 1
   strSQL = "delete CntX_Inv_Periodico where cod_contabilidad = " & gCntX_Parametros.CodigoConta _
          & " and anio = " & vGrid.Text
   vGrid.Col = 2
   strSQL = strSQL & " and mes = " & vGrid.Text
   vGrid.Col = 3
   strSQL = strSQL & " and cod_cuenta = '" & fxCntX_CuentaFormato(False, vGrid.Text) & "'"
   Call ConectionExecute(strSQL, 0)
End If

strSQL = "Select I.anio,I.mes,I.cod_cuenta,C.descripcion,I.saldo_final" _
       & " from CntX_Inv_Periodico I inner join CntX_Cuentas C on I.cod_cuenta = C.cod_cuenta" _
       & " and I.cod_contabilidad = C.cod_contabilidad" _
       & " where I.cod_contabilidad = " & gCntX_Parametros.CodigoConta
Call sbCargaGridLocal(vGrid, 5, strSQL)


End Sub

Private Sub vGrid_KeyDown(KeyCode As Integer, Shift As Integer)
Dim i As Variant, vDescripcion As String
'MsgBox "Columna : " & vGrid.Col
'MsgBox "Columna Activa: " & vGrid.ActiveCol
'MsgBox "Fila : " & vGrid.Row
'MsgBox "Fila Activa: " & vGrid.ActiveRow

'Guarda Linea
If vGrid.ActiveCol = vGrid.MaxCols And (KeyCode = 13 Or KeyCode = vbKeyTab) Then
  i = fxGuardar
  vGrid.Row = vGrid.ActiveRow
  vGrid.Col = 1
  If vGrid.MaxRows <= vGrid.ActiveRow Then
    vGrid.MaxRows = vGrid.MaxRows + 1
    vGrid.Row = vGrid.MaxRows
  End If
End If

'Consulta CntX_Cuentas
If KeyCode = vbKeyF4 And (vGrid.ActiveCol = 3) Then
  gBusquedas.Columna = "cod_cuenta"
  gBusquedas.Orden = "cod_cuenta"
  gBusquedas.Filtro = " and C.cod_contabilidad = " _
             & gCntX_Parametros.CodigoConta & " and T.clasificacion = 'A'"
  gBusquedas.Consulta = "select cod_cuenta,C.descripcion from CntX_Cuentas C" _
             & " inner join CntX_Tipos_Cuentas T on C.cod_contabilidad = T.cod_contabilidad" _
             & " and C.tipo_cuenta = T.tipo_cuenta"
  frmBusquedas.Show vbModal
  vGrid.Col = vGrid.ActiveCol
  vGrid.Row = vGrid.ActiveRow
  vGrid.Text = gBusquedas.Resultado
End If

'Da formato a las CntX_Cuentas
If (KeyCode = 13 Or KeyCode = vbKeyTab) And vGrid.ActiveCol < vGrid.MaxCols Then
    vGrid.Col = vGrid.ActiveCol
    vGrid.Row = vGrid.ActiveRow
    If vGrid.ActiveCol = 3 Then
         vDescripcion = fxCntX_Cuenta("D", fxCntX_CuentaFormato(False, vGrid.Text))
         vGrid.Text = fxCntX_CuentaFormato(True, vGrid.Text)
         vGrid.Col = 4
         vGrid.Row = vGrid.ActiveRow
         vGrid.Text = vDescripcion
    End If
End If


If KeyCode = vbKeyDelete Then
  Call sbBorrar
End If

End Sub







