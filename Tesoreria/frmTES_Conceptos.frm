VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Begin VB.Form frmTES_Conceptos 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Conceptos Bancarios"
   ClientHeight    =   8025
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12855
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8025
   ScaleWidth      =   12855
   Begin MSComctlLib.Toolbar tlb 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   12855
      _ExtentX        =   22675
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
      Height          =   7215
      Left            =   240
      TabIndex        =   1
      Top             =   600
      Width           =   12495
      _Version        =   524288
      _ExtentX        =   22040
      _ExtentY        =   12726
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
      MaxCols         =   487
      ScrollBars      =   2
      SpreadDesigner  =   "frmTES_Conceptos.frx":0000
      VScrollSpecial  =   -1  'True
      VScrollSpecialType=   2
      AppearanceStyle =   1
   End
End
Attribute VB_Name = "frmTES_Conceptos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Activate()
vModulo = 9
End Sub

Private Sub Form_Load()
Dim strSQL As String

vModulo = 9
vGrid.AppearanceStyle = fxGridStyle

Call sbToolBarIconos(tlb, False)

Call Formularios(Me)
Call RefrescaTags(Me)

If tlb.Buttons(1).Enabled = False Then vGrid.Enabled = False

strSQL = "select cod_concepto,descripcion,activo,cod_cuenta_Mask, AUTO_REGISTRO, DP_TRAMITE_APL from vTes_conceptos" _
      & " order by cod_concepto"
Call sbCargaGrid(vGrid, 6, strSQL)


End Sub


Private Function fxGuardar() As Long
Dim strSQL As String, rs As New ADODB.Recordset
'Guarda la información de la linea
'si es Insert devuelve el codigo, sino devuelve 0

On Error GoTo vError

fxGuardar = 0
vGrid.Row = vGrid.ActiveRow
vGrid.col = 1

strSQL = "select isnull(count(*),0) as Existe from tes_conceptos " _
       & " where cod_concepto = '" & vGrid.Text & "'"
Call OpenRecordSet(rs, strSQL)

vGrid.col = 4

If Not fxgCntCuentaValida(fxgCntCuentaFormato(False, vGrid.Text)) Then
  MsgBox "Cuenta Contable no es válida...", vbCritical
  Exit Function
End If

If rs!Existe = 0 Then 'Insertar
  If Trim(vGrid.Text) = "" Then Exit Function
  
  vGrid.col = 1
  strSQL = "insert into tes_conceptos(cod_concepto,descripcion,estado,cod_cuenta " _
         & ", AUTO_REGISTRO, DP_TRAMITE_APL, REGISTRO_FECHA, REGISTRO_USUARIO) values('" _
         & vGrid.Text & "', '"
  vGrid.col = 2
  strSQL = strSQL & UCase(vGrid.Text) & "', '"
  vGrid.col = 3
  strSQL = strSQL & IIf((vGrid.Value = 0), "I", "A") & "', '"
  vGrid.col = 4
  strSQL = strSQL & fxgCntCuentaFormato(False, vGrid.Text, 0) & "', "
  vGrid.col = 5
  strSQL = strSQL & vGrid.Value & ", "
  vGrid.col = 6
  strSQL = strSQL & vGrid.Value & ", dbo.myGetdate(), '" & glogon.Usuario & "')"
  Call ConectionExecute(strSQL)

  vGrid.col = 1
  Call Bitacora("Registra", "Concepto Desembolso : " & vGrid.Text)

Else 'Actualizar

 vGrid.col = 2
 strSQL = "update tes_conceptos set descripcion = '" & vGrid.Text & "',estado = '"
 vGrid.col = 3
 strSQL = strSQL & IIf((vGrid.Value = 0), "I", "A") & "',cod_cuenta = '"
 vGrid.col = 4
 strSQL = strSQL & fxgCntCuentaFormato(False, vGrid.Text, 0) & "', AUTO_REGISTRO = "
 vGrid.col = 5
 strSQL = strSQL & vGrid.Value & ", DP_TRAMITE_APL = "
  vGrid.col = 6
 strSQL = strSQL & vGrid.Value & ", MODIFICA_FECHA = dbo.myGetdate(), MODIFICA_USUARIO = '" _
    & glogon.Usuario & "' Where cod_concepto = '"
 vGrid.col = 1
 strSQL = strSQL & vGrid.Text & "'"
 
 Call ConectionExecute(strSQL)

 vGrid.col = 1
 Call Bitacora("Modifica", "Concepto Desembolso : " & vGrid.Text)

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
    vGrid.MaxRows = vGrid.MaxRows + 1

  Case "BORRAR"
     i = MsgBox("Esta Seguro que desea borrar este registro", vbYesNo)
     If i = 6 Then
        vGrid.Row = vGrid.ActiveRow
        vGrid.col = 1
        strSQL = "delete tes_conceptos where cod_concepto = '" & vGrid.Text & "'"
        Call ConectionExecute(strSQL)
        strSQL = vGrid.Text
        vGrid.col = 1
        Call Bitacora("Elimina", "Concepto Desembolso : " & vGrid.Text)
        
        vGrid.col = 1
        strSQL = "select cod_concepto,descripcion,activo,cod_cuenta_Mask, AUTO_REGISTRO, DP_TRAMITE_APL from vTes_conceptos" _
              & " order by cod_concepto"
        Call sbCargaGrid(vGrid, 6, strSQL)
     End If
  
  Case "REPORTES"
'     Call sbReportes("Caracteristicas", Me)

  Case "AYUDA"
        frmContenedor.CD.HelpContext = Me.HelpContextID
        frmContenedor.CD.ShowHelp

End Select

End Sub


Private Sub vGrid_KeyDown(KeyCode As Integer, Shift As Integer)
Dim i As Long

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
  vGrid.CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
  vGrid.CellNote = fxgCntCuentaDesc(fxgCntCuentaFormato(False, vGrid.Text))
  vGrid.TextTip = TextTipFixed
End If

'Consulta Cuentas Contables
If vGrid.ActiveCol = 4 And KeyCode = vbKeyF4 Then
   Call sbgCntCuentaConsulta("D")
   vGrid.col = vGrid.ActiveCol
   vGrid.Row = vGrid.ActiveRow
   vGrid.Text = gBusquedas.Resultado
End If

'Inserta Linea
If KeyCode = vbKeyInsert Then
    vGrid.MaxRows = vGrid.MaxRows + 1
    vGrid.InsertRows vGrid.ActiveRow, 1
    vGrid.Row = vGrid.ActiveRow
End If


End Sub









