VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Begin VB.Form frmInvTipoES 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tipos de Movimientos"
   ClientHeight    =   5355
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10485
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5355
   ScaleWidth      =   10485
   Begin MSComctlLib.Toolbar tlb 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10485
      _ExtentX        =   18494
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
      Left            =   0
      TabIndex        =   1
      Top             =   480
      Width           =   10452
      _Version        =   524288
      _ExtentX        =   18436
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
      ScrollBars      =   2
      SpreadDesigner  =   "frmInvTiposES.frx":0000
      VScrollSpecial  =   -1  'True
      VScrollSpecialType=   2
      AppearanceStyle =   1
   End
End
Attribute VB_Name = "frmInvTipoES"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Activate()
vModulo = 32
End Sub

Private Sub Form_Load()
Dim strSQL As String

vModulo = 32
vGrid.AppearanceStyle = fxGridStyle

Call sbToolBarIconos(tlb)

Call Formularios(Me)
Call RefrescaTags(Me)


strSQL = "select T.cod_entsal,T.descripcion,T.tipo,T.cod_cuenta,T.activo,C.descripcion as CtaDesc" _
       & " from pv_entrada_salida T left join CntX_cuentas C on T.cod_cuenta = C.cod_cuenta and C.cod_contabilidad = " & GLOBALES.gEnlace _
       & " order by T.cod_entsal"
Call sbCargaGridLocal(vGrid, 5, strSQL)

If tlb.Buttons(1).Enabled = False Then vGrid.Enabled = False

End Sub


Private Sub sbCargaGridLocal(vGrid As Object, vGridMaxCol As Integer, strSQL As String)
Dim rs As New ADODB.Recordset, i As Integer, strResultado As String

Me.MousePointer = vbHourglass

vGrid.MaxCols = vGridMaxCol
vGrid.MaxRows = 1

vGrid.Row = vGrid.MaxRows

rs.CursorLocation = adUseServer
Call OpenRecordSet(rs, strSQL, 0)


Do While Not rs.EOF
  vGrid.Row = vGrid.MaxRows
  
  For i = 1 To vGrid.MaxCols
    vGrid.col = i
    Select Case i
     Case 1
        vGrid.Text = CStr(rs!COD_ENTSAL)
     Case 2
        vGrid.Text = CStr(rs!Descripcion)
     Case 3
        Select Case rs!Tipo
          Case "E"
             vGrid.Text = "Entradas"
          Case "S"
             vGrid.Text = "Salidas"
          Case "T"
             vGrid.Text = "Traslados"
          Case "R"
             vGrid.Text = "Requisiciones"
        End Select
     Case 4
        vGrid.Text = fxgCntCuentaFormato(True, rs!cod_cuenta)
        vGrid.CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
        vGrid.CellNote = rs!CtaDesc & ""
        vGrid.TextTip = TextTipFixed

     Case 5
        vGrid.Value = rs!activo
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

vGrid.col = 1
fxGuardar = 0

vGrid.Row = vGrid.ActiveRow
vGrid.col = 4
vCuenta = fxgCntCuentaFormato(False, vGrid.Text)

vGrid.col = 1


strSQL = "select isnull(count(*),0) as Existe from pv_entrada_salida" _
       & " where cod_entsal = '" & vGrid.Text & "'"
Call OpenRecordSet(rs, strSQL)

If rs!Existe = 0 Then 'Insertar
  
  strSQL = "insert into pv_entrada_salida(cod_entsal,descripcion,cod_cuenta,tipo,activo) values('" & UCase(vGrid.Text)
  vGrid.col = 2
  strSQL = strSQL & "','" & UCase(vGrid.Text) & "','" & vCuenta & "','"
  vGrid.col = 3
  strSQL = strSQL & Mid(vGrid.Text, 1, 1) & "',"
  vGrid.col = 5
  strSQL = strSQL & vGrid.Value & ")"
  
  Call ConectionExecute(strSQL)

  vGrid.col = 1
  
  Call Bitacora("Registra", "Tipo de E/S/T Cod: " & vGrid.Text)
  
Else 'Actualizar

    vGrid.col = 2
    strSQL = "update pv_entrada_salida set descripcion = '" & UCase(vGrid.Text) & "'" _
           & ",cod_cuenta = '" & vCuenta & "',tipo = '"
    vGrid.col = 3
    strSQL = strSQL & Mid(vGrid.Text, 1, 1) & "', Activo = "
    vGrid.col = 5
    strSQL = strSQL & vGrid.Value & " where cod_entsal = '"
    
    vGrid.col = 1
    strSQL = strSQL & vGrid.Text & "'"
    Call ConectionExecute(strSQL)
 
   Call Bitacora("Modifica", "Tipo de E/S/T Cod: " & vGrid.Text)
 
End If

rs.Close
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
        strSQL = "delete pv_entrada_salida where cod_entsal = '" & vGrid.Text & "'"
        Call ConectionExecute(strSQL)
                
        Call Bitacora("Elimina", "Tipo de E/S/T Cod: " & vGrid.Text)
        
        strSQL = "select T.cod_entsal,T.descripcion,T.tipo,T.cod_cuenta,T.activo,C.descripcion as CtaDesc" _
               & " from pv_entrada_salida T left join CntX_cuentas C on T.cod_cuenta = C.cod_cuenta and C.cod_contabilidad = " & GLOBALES.gEnlace _
               & " order by T.cod_entsal"
        Call sbCargaGridLocal(vGrid, 5, strSQL)
     End If
  
  Case "REPORTES"
     Call sbInvReportes("TiposEST", "Tipos de Entradas/Salidas/Traslado", "Listado", "")
  
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





