VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Begin VB.Form frmCntX_Cierres 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Cierres Fiscales"
   ClientHeight    =   5736
   ClientLeft      =   48
   ClientTop       =   300
   ClientWidth     =   12600
   HelpContextID   =   5
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5736
   ScaleWidth      =   12600
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin FPSpreadADO.fpSpread vGrid 
      Height          =   4332
      Left            =   120
      TabIndex        =   0
      Top             =   1200
      Width           =   12372
      _Version        =   524288
      _ExtentX        =   21823
      _ExtentY        =   7641
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
      MaxCols         =   494
      ScrollBars      =   2
      SpreadDesigner  =   "frmCntX_Cierres.frx":0000
      VScrollSpecial  =   -1  'True
      VScrollSpecialType=   2
      AppearanceStyle =   1
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Definición de Cierres Fiscales"
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
      Height          =   372
      Index           =   1
      Left            =   2040
      TabIndex        =   1
      Top             =   360
      Width           =   7692
   End
   Begin VB.Image imgBanner 
      Appearance      =   0  'Flat
      Height          =   972
      Left            =   0
      Top             =   0
      Width           =   12852
   End
End
Attribute VB_Name = "frmCntX_Cierres"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub sbCargaGridLocal(vGrid As Object, vGridMaxCol As Integer, strSQL As String)
Dim rs As New ADODB.Recordset, i As Integer

Me.MousePointer = vbHourglass

On Error GoTo vError

vGrid.MaxCols = vGridMaxCol
vGrid.MaxRows = 1
vGrid.Row = vGrid.MaxRows

Call OpenRecordSet(rs, strSQL, 0)
Do While Not rs.EOF
  vGrid.Row = vGrid.MaxRows
  For i = 1 To vGrid.MaxCols
    Select Case i
      Case 7, 8, 9
        vGrid.Col = i
        vGrid.Text = fxCntX_CuentaFormato(True, CStr(rs.Fields(i - 1).Value))
      Case Else
        vGrid.Col = i
        vGrid.Text = rs.Fields(i - 1).Value
     End Select
  Next i
  vGrid.MaxRows = vGrid.MaxRows + 1
  rs.MoveNext
Loop
rs.Close

vError:

Me.MousePointer = vbDefault

End Sub


Private Sub Form_Activate()
vModulo = 20

End Sub

Private Sub Form_Load()
Dim strSQL As String

vModulo = 20

vGrid.AppearanceStyle = fxGridStyle

Set imgBanner.Picture = frmContenedor.imgBanner_Mantenimiento.Picture


strSQL = "select id_cierre,inicio_anio,inicio_mes,corte_anio,corte_mes,descripcion" _
       & ",cuenta_ganper,cuenta_utilidad,cuenta_imprenta,impuesto_renta,activo" _
       & " from CntX_Cierres " _
       & " where COD_CONTABILIDAD = " & gCntX_Parametros.CodigoConta _
       & " Order by inicio_anio desc,inicio_mes desc"
Call sbCargaGridLocal(vGrid, 11, strSQL)

Call Formularios(Me)
Call RefrescaTags(Me)


End Sub

Private Function fxVerificaCuentas(pCuenta As String) As Boolean
Dim rsX As New ADODB.Recordset, strSQL As String

fxVerificaCuentas = False

strSQL = "select isnull(count(*),0) as Total from Cntx_Cuentas where COD_CONTABILIDAD = " _
       & gCntX_Parametros.CodigoConta & " and cod_cuenta = '" & fxCntX_CuentaFormato(False, pCuenta) & "' and acepta_movimientos = 1"
Call OpenRecordSet(rsX, strSQL, 0)
If rsX!Total = 1 Then fxVerificaCuentas = True
rsX.Close
End Function

Private Function fxVerificaDatos() As Boolean
Dim lngAnioInicio As Long, lngAnioCorte As Long
Dim iMesInicio As Integer, iMesCorte As Long

On Error GoTo vError

fxVerificaDatos = True

vGrid.Row = vGrid.ActiveRow
'1. Verificar Si existen Las cuentas en el catalogo
'2. Verificar si el periodo no choca con otros

vGrid.Col = 7
fxVerificaDatos = fxVerificaCuentas(vGrid.Text)

If fxVerificaDatos Then
    vGrid.Col = 8
    fxVerificaDatos = fxVerificaCuentas(vGrid.Text)
End If

If fxVerificaDatos Then
    vGrid.Col = 9
    fxVerificaDatos = fxVerificaCuentas(vGrid.Text)
End If

If fxVerificaDatos Then
    vGrid.Col = 2
    lngAnioInicio = vGrid.Text
    vGrid.Col = 3
    iMesInicio = vGrid.Text
    
    vGrid.Col = 4
    lngAnioCorte = vGrid.Text
    vGrid.Col = 5
    iMesCorte = vGrid.Text
End If



Exit Function

vError:
  fxVerificaDatos = False
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Function

Private Function fxGuardar() As Long
Dim strSQL As String, rs As New ADODB.Recordset
'Guarda la información de la linea
'si es Insert devuelve el codigo, sino devuelve 0

On Error GoTo vError

fxGuardar = 0
vGrid.Row = vGrid.ActiveRow
vGrid.Col = 1

If Not fxVerificaDatos Then
  MsgBox "La información especificada no es válida, verifiquela...", vbCritical
  fxGuardar = 1
  Exit Function
End If

vGrid.Col = 1
If vGrid.Text = "" Then
    vGrid.Col = 1

    strSQL = "select isnull(max(id_Cierre),0) + 1 as Consecutivo from CntX_Cierres where cod_contabilidad = " & gCntX_Parametros.CodigoConta
    Call OpenRecordSet(rs, strSQL, 0)
    vGrid.Text = rs!Consecutivo
    rs.Close

    strSQL = "insert into CntX_Cierres(id_cierre,COD_CONTABILIDAD,inicio_anio,inicio_mes,corte_anio,corte_mes" _
           & ",descripcion,cuenta_ganper,cuenta_utilidad,cuenta_imprenta,impuesto_renta,activo)" _
           & " values(" & vGrid.Text & "," & gCntX_Parametros.CodigoConta & ","
    vGrid.Col = 2
    strSQL = strSQL & vGrid.Text & ","
    vGrid.Col = 3
    strSQL = strSQL & vGrid.Text & ","
    vGrid.Col = 4
    strSQL = strSQL & vGrid.Text & ","
    vGrid.Col = 5
    strSQL = strSQL & vGrid.Text & ",'"
    vGrid.Col = 6
    strSQL = strSQL & vGrid.Text & "','"
    vGrid.Col = 7
    strSQL = strSQL & fxCntX_CuentaFormato(False, vGrid.Text) & "','"
    vGrid.Col = 8
    strSQL = strSQL & fxCntX_CuentaFormato(False, vGrid.Text) & "','"
    vGrid.Col = 9
    strSQL = strSQL & fxCntX_CuentaFormato(False, vGrid.Text) & "',"
    vGrid.Col = 10
    If Not IsNumeric(vGrid.Text) Then
      strSQL = strSQL & "0,"
    Else
        strSQL = strSQL & vGrid.Text & ","
    End If
    vGrid.Col = 11
    strSQL = strSQL & vGrid.Value & ")"
    
    
    Call ConectionExecute(strSQL, 0)
  
    vGrid.Col = 1
    Call Bitacora("Registra", "Cierre Fiscal : " & vGrid.Text & " Conta." & gCntX_Parametros.CodigoConta)
  
   Else 'Actualizar
       
       vGrid.Col = 6
       strSQL = "update CntX_Cierres set descripcion = '" & vGrid.Text & "',inicio_anio = "
       vGrid.Col = 2
       strSQL = strSQL & vGrid.Text & ",inicio_mes = "
       vGrid.Col = 3
       strSQL = strSQL & vGrid.Text & ",corte_anio = "
       vGrid.Col = 4
       strSQL = strSQL & vGrid.Text & ",corte_mes = "
       vGrid.Col = 5
       strSQL = strSQL & vGrid.Text & ",cuenta_ganper = '"
       vGrid.Col = 7
       strSQL = strSQL & fxCntX_CuentaFormato(False, vGrid.Text) & "',cuenta_utilidad = '"
       vGrid.Col = 8
       strSQL = strSQL & fxCntX_CuentaFormato(False, vGrid.Text) & "',cuenta_imprenta = '"
       vGrid.Col = 9
       strSQL = strSQL & fxCntX_CuentaFormato(False, vGrid.Text) & "',impuesto_renta = "
       vGrid.Col = 10
       If Not IsNumeric(vGrid.Text) Then
         strSQL = strSQL & "0,activo = "
       Else
         strSQL = strSQL & vGrid.Text & ",activo = "
       End If
       
       vGrid.Col = 11
       strSQL = strSQL & vGrid.Value
       vGrid.Col = 1
       strSQL = strSQL & " where id_cierre = " & vGrid.Text & " and cod_Contabilidad = " & gCntX_Parametros.CodigoConta
      
      Call ConectionExecute(strSQL, 0)
    
      vGrid.Col = 1
      vGrid.Col = 2
      Call Bitacora("Actualiza", "Cierre Fiscal : " & vGrid.Text & " Conta." & gCntX_Parametros.CodigoConta)
       
 End If

 vGrid.Col = 1
 fxGuardar = 1
 
Exit Function

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Function



Private Sub vGrid_KeyDown(KeyCode As Integer, Shift As Integer)
Dim i As Integer, strSQL As String

On Error GoTo vError

'Guarda Linea
If vGrid.ActiveCol = vGrid.MaxCols And (KeyCode = vbKeyReturn Or KeyCode = vbKeyTab) Then
  i = fxGuardar
  vGrid.Row = vGrid.ActiveRow
  vGrid.Col = 1
  If vGrid.MaxRows <= vGrid.ActiveRow Then
    vGrid.MaxRows = vGrid.MaxRows + 1
    vGrid.Row = vGrid.MaxRows
  End If
End If

'Consulta Cuentas
If KeyCode = vbKeyF4 And (vGrid.ActiveCol = 7 Or vGrid.ActiveCol = 8 Or vGrid.ActiveCol = 9) Then
'  gBusquedas.Columna = "cod_cuenta"
'  gBusquedas.Orden = "cod_cuenta"
'  gBusquedas.Filtro = " and acepta_movimientos = 1 and COD_CONTABILIDAD = " & gCntX_Parametros.CodigoConta
'  gBusquedas.Consulta = "select cod_cuenta, descripcion from CntX_Cuentas"
'  frmBusquedas.Show vbModal
  frmCntX_ConsultaCuentas.Show vbModal
  vGrid.Col = vGrid.ActiveCol
  vGrid.Row = vGrid.ActiveRow
  vGrid.Text = gCuenta
End If

'Da formato a las cuentas
If (KeyCode = vbKeyReturn Or KeyCode = vbKeyTab) And vGrid.ActiveCol < vGrid.MaxCols Then
    vGrid.Col = vGrid.ActiveCol
    vGrid.Row = vGrid.ActiveRow
    
    Select Case vGrid.ActiveCol
      Case 7, 8, 9 'Cuenta
        vGrid.Text = fxCntX_CuentaFormato(True, vGrid.Text)
    End Select
End If


'Reporte
If KeyCode = vbKeyF5 Then
    Call sbCntX_Reportes_Catalogos("Cierres")
End If

'Inserta Linea
If KeyCode = vbKeyInsert Then
    vGrid.MaxRows = vGrid.MaxRows + 1
    vGrid.InsertRows vGrid.ActiveRow, 1
    vGrid.Row = vGrid.ActiveRow
End If

'Borrar una linea
If KeyCode = vbKeyDelete Then
     i = MsgBox("Esta Seguro que desea borrar este registro", vbYesNo)
     If i = vbYes Then
        
        vGrid.Row = vGrid.ActiveRow
        vGrid.Col = 1
        
        vGrid.Row = vGrid.ActiveRow
        vGrid.Col = 1
        
        
        strSQL = "delete CntX_Cierres where COD_CONTABILIDAD = " & gCntX_Parametros.CodigoConta _
               & " and Activo = 1 and id_cierre = " & vGrid.Text
        Call ConectionExecute(strSQL, 0)
        strSQL = vGrid.Text
        vGrid.Col = 2
        
        Call Bitacora("Elimina", "Cierre Fiscal : " & vGrid.Text & " Conta." & gCntX_Parametros.CodigoConta)
        
        vGrid.DeleteRows vGrid.ActiveRow, 1
        vGrid.MaxRows = vGrid.MaxRows - 1
        If vGrid.MaxRows = 0 Then vGrid.MaxRows = 1
     
     End If
End If

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

