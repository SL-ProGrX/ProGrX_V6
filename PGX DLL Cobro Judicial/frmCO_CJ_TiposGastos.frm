VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Begin VB.Form frmCO_CJ_TiposGastos 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Cobros: Tipos de Gastos"
   ClientHeight    =   6480
   ClientLeft      =   48
   ClientTop       =   312
   ClientWidth     =   10332
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6480
   ScaleWidth      =   10332
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin FPSpreadADO.fpSpread vGrid 
      Height          =   4932
      Left            =   240
      TabIndex        =   0
      Top             =   1320
      Width           =   9852
      _Version        =   524288
      _ExtentX        =   17378
      _ExtentY        =   8700
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
      MaxCols         =   497
      ScrollBars      =   2
      SpreadDesigner  =   "frmCO_CJ_TiposGastos.frx":0000
      VScrollSpecialType=   2
      AppearanceStyle =   1
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Tipos de Gastos del Proceso de Cobro Judicial"
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
      Height          =   612
      Index           =   1
      Left            =   1920
      TabIndex        =   1
      Top             =   360
      Width           =   8292
   End
   Begin VB.Image imgBanner 
      Height          =   1215
      Left            =   0
      Top             =   0
      Width           =   10815
   End
End
Attribute VB_Name = "frmCO_CJ_TiposGastos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Activate()
vModulo = 6
End Sub

Private Sub Form_Load()
Dim strSQL As String

vModulo = 6
vGrid.AppearanceStyle = fxGridStyle
Set imgBanner.Picture = frmContenedor.imgBanner_Mantenimiento.Picture

Call Formularios(Me)
Call RefrescaTags(Me)


strSQL = "select T.*,C.descripcion as CtaDesc" _
       & " from Cbr_Cj_Tipos_Gastos T left join CntX_cuentas C on T.cod_cuenta = C.cod_cuenta and C.cod_contabilidad = " & GLOBALES.gEnlace _
       & " order by T.Tipo_Gasto"
Call sbCargaGridLocal(vGrid, 6, strSQL)


End Sub


Private Sub sbCargaGridLocal(vGrid As Object, vGridMaxCol As Integer, strSQL As String)
Dim rs As New ADODB.Recordset, i As Integer, strResultado As String

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
     Case 1
        vGrid.Text = CStr(rs!Tipo_Gasto)
     Case 2
        vGrid.Text = CStr(rs!Descripcion)
     Case 3 'Monto Default
        vGrid.Text = CStr(rs!Monto)
     
     Case 4 'Cuenta Contable
        vGrid.Text = fxgCntCuentaFormato(True, rs!cod_cuenta)
        vGrid.CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
        vGrid.CellNote = rs!CtaDesc & ""
        vGrid.TextTip = TextTipFixed

     Case 5 'Aplica Desembolsos
        vGrid.Value = rs!Aplica_Desembolso
     
     Case 6 'Activida + Record
        vGrid.Value = rs!activo
        vGrid.CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
        vGrid.CellNote = "Registro: " & rs!Registro_Usuario & vbCrLf & "Fecha: " & rs!Registro_Fecha
        vGrid.TextTip = TextTipFixed

    End Select
  
  Next i
  
  vGrid.MaxRows = vGrid.MaxRows + 1
  
  rs.MoveNext

Loop

rs.Close

Me.MousePointer = vbDefault

End Sub


Private Function fxGuardar() As Long
Dim strSQL As String, rs As New ADODB.Recordset
Dim vCuenta As String

On Error GoTo vError

vGrid.Col = 1
fxGuardar = 0

vGrid.Row = vGrid.ActiveRow
vGrid.Col = 4
vCuenta = fxgCntCuentaFormato(False, vGrid.Text)


If Not fxgCntCuentaValida(vCuenta) Then
   MsgBox "No se puede guardar el registro porque la Cuenta Contable no es válida!", vbExclamation
   Exit Function
End If
vGrid.Col = 1


strSQL = "select isnull(count(*),0) as Existe from Cbr_Cj_Tipos_Gastos" _
       & " where Tipo_Gasto = '" & vGrid.Text & "'"
Call OpenRecordSet(rs, strSQL)

If rs!Existe = 0 Then 'Insertar
  
  strSQL = "insert into Cbr_Cj_Tipos_Gastos(Tipo_Gasto,descripcion,monto,cod_cuenta,Aplica_Desembolso,Activo,Registro_Usuario,Registro_Fecha) values('" & vGrid.Text
  vGrid.Col = 2
  strSQL = strSQL & "','" & vGrid.Text & "',"
  vGrid.Col = 3
  strSQL = strSQL & CCur(vGrid.Text) & ",'" & vCuenta & "',"
  vGrid.Col = 5
  strSQL = strSQL & vGrid.Value & ","
  vGrid.Col = 6
  strSQL = strSQL & vGrid.Value & ",'" & glogon.Usuario & "',dbo.MyGetdate())"
  
  Call ConectionExecute(strSQL)

  vGrid.Col = 1
  
  Call Bitacora("Registra", "Tipo de Gasto: " & vGrid.Text)
  
Else 'Actualizar

    vGrid.Col = 2
    strSQL = "update Cbr_Cj_Tipos_Gastos set descripcion = '" & vGrid.Text & "'" _
           & ",cod_cuenta = '" & vCuenta & "',Monto = "
    vGrid.Col = 3
    strSQL = strSQL & CCur(vGrid.Text) & ", Aplica_Desembolso = "
    vGrid.Col = 5
    strSQL = strSQL & vGrid.Value & ", Activo = "
    vGrid.Col = 6
    strSQL = strSQL & vGrid.Value & " where Tipo_Gasto = '"
    vGrid.Col = 1
    strSQL = strSQL & vGrid.Text & "'"
    Call ConectionExecute(strSQL)
 
   Call Bitacora("Modifica", "Tipo de Gasto: " & vGrid.Text)
 
End If

rs.Close
fxGuardar = 1

Exit Function

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Function



Private Sub vGrid_KeyDown(KeyCode As Integer, Shift As Integer)
Dim i As Long, strSQL As String

If vGrid.ActiveCol = vGrid.MaxCols And (KeyCode = vbKeyReturn Or KeyCode = vbKeyTab) Then
  i = fxGuardar
  vGrid.Row = vGrid.ActiveRow
  vGrid.Col = 1
  If vGrid.MaxRows <= vGrid.ActiveRow Then
    vGrid.MaxRows = vGrid.MaxRows + 1
    vGrid.Row = vGrid.MaxRows
  End If
End If

'Formato de Cuenta Contable
If vGrid.ActiveCol = 4 And (KeyCode = vbKeyReturn Or KeyCode = vbKeyTab) Then
  vGrid.Col = vGrid.ActiveCol
  vGrid.Row = vGrid.ActiveRow
  vGrid.Text = fxgCntCuentaFormato(True, vGrid.Text)
End If

'Consulta Cuentas Contables
If vGrid.ActiveCol = 4 And KeyCode = vbKeyF4 Then
  frmCntX_ConsultaCuentas.Show vbModal
  vGrid.Col = vGrid.ActiveCol
  vGrid.Row = vGrid.ActiveRow
  vGrid.Text = gCuenta
End If

'Inserta Linea
If KeyCode = vbKeyInsert Then
    vGrid.MaxRows = vGrid.MaxRows + 1
    vGrid.InsertRows vGrid.ActiveRow, 1
    vGrid.Row = vGrid.ActiveRow
End If


'Borrar Linea
If KeyCode = vbKeyDelete Then
     i = MsgBox("Esta Seguro que desea borrar este registro", vbYesNo)
     If i = vbYes Then
        vGrid.Row = vGrid.ActiveRow
        vGrid.Col = 1
        strSQL = "delete Cbr_Cj_Tipos_Gastos where Tipo_Gasto = '" & vGrid.Text & "'"
        Call ConectionExecute(strSQL)
        
        strSQL = vGrid.Text
        vGrid.Col = 1
        Call Bitacora("Elimina", "Tipo de Gasto: " & vGrid.Text)
                
        vGrid.DeleteRows vGrid.ActiveRow, 1
        If vGrid.MaxRows > 1 Then vGrid.MaxRows = vGrid.MaxRows - 1
        vGrid.Row = vGrid.ActiveRow
     End If
End If

End Sub

