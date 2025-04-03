VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Begin VB.Form frmPosFormaPago 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Formas de Pago"
   ClientHeight    =   5595
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11745
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5595
   ScaleWidth      =   11745
   Begin MSComctlLib.Toolbar tlb 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11745
      _ExtentX        =   20717
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   5
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "nuevo"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "borrar"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "reportes"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "ayuda"
         EndProperty
      EndProperty
   End
   Begin FPSpreadADO.fpSpread vGrid 
      Height          =   4812
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   11412
      _Version        =   524288
      _ExtentX        =   20130
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
      MaxCols         =   484
      ScrollBars      =   2
      SpreadDesigner  =   "frmPosFormaPago.frx":0000
      VScrollSpecial  =   -1  'True
      VScrollSpecialType=   2
      AppearanceStyle =   1
   End
End
Attribute VB_Name = "frmPosFormaPago"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Activate()
vModulo = 33
End Sub

Private Sub Form_Load()
Dim strSQL As String

vModulo = 33
vGrid.AppearanceStyle = fxGridStyle

Call sbToolBarIconos(tlb)

Call Formularios(Me)
Call RefrescaTags(Me)


strSQL = "select cod_forma_pago,descripcion,clasificacion,cod_cuenta,prioridad,pide_banco,pide_cuenta,pide_documento" _
       & " From pv_formas_pago"
Call sbCargaGridLocal(vGrid, 8, strSQL)

If tlb.Buttons(1).Enabled = False Then vGrid.Enabled = False

End Sub


Private Sub sbCargaGridLocal(vGrid As Object, vGridMaxCol As Integer, strSQL As String)
Dim rs As New ADODB.Recordset, i As Integer, strResultado As String

Me.MousePointer = vbHourglass

vGrid.MaxCols = vGridMaxCol
vGrid.MaxRows = 1

vGrid.Row = vGrid.MaxRows

Call OpenRecordSet(rs, strSQL, 0)


Do While Not rs.EOF
  vGrid.Row = vGrid.MaxRows
  
  For i = 1 To vGrid.MaxCols
    vGrid.col = i
    Select Case i
     Case 1
        vGrid.Text = CStr(rs!Cod_Forma_Pago)
     Case 2
        vGrid.Text = CStr(rs!Descripcion)
     Case 3
        Select Case Trim(rs!Clasificacion)
          Case "01"
             vGrid.Text = "01 - Efectivo"
          Case "02"
             vGrid.Text = "02 - CxC Externa"
          Case "03"
             vGrid.Text = "03 - CxC Interna"
          Case "04"
             vGrid.Text = "04 - Fondos"
        End Select
     Case 4
        vGrid.Text = fxgCntCuentaFormato(True, rs!cod_cuenta)
        vGrid.CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
        vGrid.CellNote = fxgCntCuentaDesc(rs!cod_cuenta)
        vGrid.TextTip = TextTipFixed
     
     Case 5
        vGrid.Text = CStr(rs!prioridad)
     
     Case 6, 7, 8
        vGrid.Value = rs.Fields(i - 1)
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
Dim vCuenta As String, vCodigo As Integer

On Error GoTo vError

vGrid.col = 1
fxGuardar = 0

vGrid.Row = vGrid.ActiveRow
vGrid.col = 4
vCuenta = fxgCntCuentaFormato(False, vGrid.Text)

vGrid.col = 1

If vGrid.Text = "" Then 'Insertar

   strSQL = "select isnull(max(cod_forma_pago),0) as ultimo from pv_formas_pago"
   Call OpenRecordSet(rs, strSQL)
     vCodigo = rs!ultimo + 1
   rs.Close

  vGrid.col = 2
  
  strSQL = "insert into pv_formas_pago(cod_forma_pago,descripcion,clasificacion,cod_cuenta" _
          & ",prioridad,pide_banco,pide_cuenta,pide_documento)" _
         & "  values(" & vCodigo & ",'" & UCase(vGrid.Text) & "','"
  vGrid.col = 3
  strSQL = strSQL & Mid(vGrid.Text, 1, 2) & "','" & vCuenta & "','"
  vGrid.col = 5
  strSQL = strSQL & vGrid.Text & "',"
  vGrid.col = 6
  strSQL = strSQL & vGrid.Value & ","
  vGrid.col = 7
  strSQL = strSQL & vGrid.Value & ","
  vGrid.col = 8
  strSQL = strSQL & vGrid.Value & ")"
  
  Call ConectionExecute(strSQL)

  vGrid.col = 1
  vGrid.Text = CStr(vCodigo)
    
  Call Bitacora("Registra", "Forma de Pago : " & vCodigo)
    
  
Else 'Actualizar

    vGrid.col = 2
    strSQL = "update pv_formas_pago set descripcion = '" & UCase(vGrid.Text) & "'" _
           & ",cod_cuenta = '" & vCuenta & "',clasificacion = '"
    vGrid.col = 3
    strSQL = strSQL & Mid(vGrid.Text, 1, 2) & "',prioridad = '"
    vGrid.col = 5
    strSQL = strSQL & vGrid.Text & "',pide_banco = "
    vGrid.col = 6
    strSQL = strSQL & vGrid.Value & ",pide_cuenta = "
    vGrid.col = 7
    strSQL = strSQL & vGrid.Value & ",pide_documento = "
    vGrid.col = 8
    strSQL = strSQL & vGrid.Value
    
    vGrid.col = 1
    strSQL = strSQL & " where cod_forma_pago = " & vGrid.Text
    Call ConectionExecute(strSQL)
   
    Call Bitacora("Modifica", "Forma de Pago : " & vGrid.Text)
   
End If

fxGuardar = 1

Exit Function

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Function


Private Sub tlb_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim i As Integer, strSQL As String

On Error Resume Next

Select Case UCase(Button.Key)
  Case "NUEVO"
    vGrid.Row = vGrid.MaxRows
    vGrid.col = 1
    If vGrid.Text <> "" Then vGrid.MaxRows = vGrid.MaxRows + 1
  
  Case "BORRAR"
     i = MsgBox("Esta Seguro que desea borrar este registro...", vbYesNo)
     If i = vbYes Then
        vGrid.Row = vGrid.ActiveRow
        vGrid.col = 1
        strSQL = "delete pv_formas_pago where cod_forma_pago = " & vGrid.Text
        Call ConectionExecute(strSQL)
                
        Call Bitacora("Elimina", "Forma de Pago Cod: " & vGrid.Text)
        
        strSQL = "select cod_forma_pago,descripcion,clasificacion,cod_cuenta,prioridad,pide_banco,pide_cuenta,pide_documento" _
               & " From pv_formas_pago"
        Call sbCargaGridLocal(vGrid, 8, strSQL)
     End If
  
  Case "REPORTES"
 '    Call sbInvReportes("TiposEST", "Tipos de Entradas/Salidas/Traslado", "Listado", "")
  
  Case "AYUDA"
        frmContenedor.CD.HelpContext = Me.HelpContextID
        frmContenedor.CD.ShowHelp


End Select

End Sub


Private Sub vGrid_KeyDown(KeyCode As Integer, Shift As Integer)
Dim i As Long
'MsgBox "Columna : " & vGrid.Col
'MsgBox "Columna Activa: " & vGrid.ActiveCol
'MsgBox "Fila : " & vGrid.Row
'MsgBox "Fila Activa: " & vGrid.ActiveRow

If vGrid.ActiveCol = vGrid.MaxCols And (KeyCode = vbKeyReturn Or KeyCode = vbKeyTab) Then
  i = fxGuardar
  vGrid.Row = vGrid.ActiveRow
  vGrid.col = 1
  If vGrid.MaxRows <= vGrid.ActiveRow Then
    vGrid.MaxRows = vGrid.MaxRows + 1
    vGrid.Row = vGrid.MaxRows
  End If
End If

'Formato de Cuenta Contable
If vGrid.ActiveCol = 4 And (KeyCode = vbKeyReturn Or KeyCode = vbKeyTab) Then
  vGrid.col = vGrid.ActiveCol
  vGrid.Row = vGrid.ActiveRow
  vGrid.Text = fxgCntCuentaFormato(True, vGrid.Text)
End If

'Consulta Cuentas Contables
If vGrid.ActiveCol = 4 And KeyCode = vbKeyF4 Then
  frmCntX_ConsultaCuentas.Show vbModal
  vGrid.col = vGrid.ActiveCol
  vGrid.Row = vGrid.ActiveRow
  vGrid.Text = gCuenta
End If

'Inserta Linea
If KeyCode = vbKeyInsert Then
    vGrid.MaxRows = vGrid.MaxRows + 1
    vGrid.InsertRows vGrid.ActiveRow, 1
    vGrid.Row = vGrid.ActiveRow
End If

End Sub


