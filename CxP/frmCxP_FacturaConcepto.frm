VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.controls.v22.1.0.ocx"
Begin VB.Form frmCxPFacturaConcepto 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Conceptos de Facturas"
   ClientHeight    =   4905
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8640
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4905
   ScaleWidth      =   8640
   Begin XtremeSuiteControls.CheckBox chkActivo 
      Height          =   255
      Left            =   2160
      TabIndex        =   11
      Top             =   840
      Width           =   2055
      _Version        =   1441793
      _ExtentX        =   3625
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Concepto Activo?"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
      Appearance      =   17
   End
   Begin VB.TextBox txtTotal 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   6720
      Locked          =   -1  'True
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   4080
      Width           =   855
   End
   Begin MSComctlLib.Toolbar tlb 
      Align           =   1  'Align Top
      Height          =   330
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   8640
      _ExtentX        =   15240
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Style           =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   9
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "nuevo"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "editar"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "borrar"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "guardar"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "deshacer"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "consultar"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "reportes"
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "ayuda"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar sBar 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   5
      Top             =   4590
      Width           =   8640
      _ExtentX        =   15240
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   6068
            MinWidth        =   6068
            Object.ToolTipText     =   "Usuario"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   6068
            MinWidth        =   6068
            Object.ToolTipText     =   "Fecha y Hora de la Creación"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   3528
            MinWidth        =   3528
            Text            =   "Conceptos"
            TextSave        =   "Conceptos"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin FPSpreadADO.fpSpread vGrid 
      Height          =   2052
      Left            =   120
      TabIndex        =   6
      Top             =   1920
      Width           =   8532
      _Version        =   524288
      _ExtentX        =   15050
      _ExtentY        =   3620
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
      SpreadDesigner  =   "frmCxP_FacturaConcepto.frx":0000
      VScrollSpecial  =   -1  'True
      VScrollSpecialType=   2
      AppearanceStyle =   1
   End
   Begin MSComCtl2.FlatScrollBar FlatScrollBar 
      Height          =   255
      Left            =   8040
      TabIndex        =   7
      Top             =   480
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   450
      _Version        =   393216
      Arrows          =   65536
      Orientation     =   1638401
   End
   Begin XtremeSuiteControls.FlatEdit txtCodigo 
      Height          =   315
      Left            =   1080
      TabIndex        =   8
      Top             =   480
      Width           =   1095
      _Version        =   1441793
      _ExtentX        =   1926
      _ExtentY        =   550
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   2
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtDescripcion 
      Height          =   315
      Left            =   2160
      TabIndex        =   9
      Top             =   480
      Width           =   5775
      _Version        =   1441793
      _ExtentX        =   10186
      _ExtentY        =   556
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtNotas 
      Height          =   675
      Left            =   1080
      TabIndex        =   10
      Top             =   1200
      Width           =   6855
      _Version        =   1441793
      _ExtentX        =   12091
      _ExtentY        =   1191
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MultiLine       =   -1  'True
      ScrollBars      =   2
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin VB.Label Label5 
      Caption         =   "Concepto"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   480
      Width           =   1005
   End
   Begin VB.Label Label8 
      Caption         =   "Total:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   5880
      TabIndex        =   3
      Top             =   4080
      Width           =   765
   End
   Begin VB.Label Label2 
      Caption         =   "Nota"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Width           =   945
   End
End
Attribute VB_Name = "frmCxPFacturaConcepto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vEdita As Boolean, vBusca As Integer
Dim vScroll As Boolean

Private Sub sbLimpiezaParcial(iCodigo As Integer)
    vGrid.MaxRows = 0
    vGrid.MaxRows = 1
    txtNotas = ""
End Sub

Private Sub FlatScrollBar_Change()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

If vScroll Then
    strSQL = "select Top 1 cod_concepto from cxp_facConceptos"
           
    If FlatScrollBar.Value = 1 Then
       strSQL = strSQL & " where cod_concepto > '" & txtCodigo.Text & "' order by cod_concepto asc"
    Else
       strSQL = strSQL & " where cod_concepto < '" & txtCodigo.Text & "' order by cod_concepto desc"
    End If
    
    Call OpenRecordSet(rs, strSQL)
    If Not rs.EOF And Not rs.BOF Then
      Call sbConsulta(rs!cod_Concepto)
    End If
    rs.Close
End If

vScroll = False
FlatScrollBar.Value = 0
vScroll = True

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub Form_Load()

Set Me.Icon = frmContenedor.Icon

Call sbToolBarIconos(tlb)
 

vGrid.AppearanceStyle = fxGridStyle

 vScroll = False
 FlatScrollBar.Value = 0
 vScroll = True

vEdita = False
Call sbToolBar(tlb, "activo")
Call sbLimpiaPantalla

Call Formularios(Me)
Call RefrescaTags(Me)
 
End Sub

Private Function fxVerificaAsiento() As Boolean
Dim rsX As New ADODB.Recordset, strSQL As String
Dim vMensaje As String, lng As Long

'Verificar Periodo
'Tipo de Asiento
'Cuentas (En el Detalle)

fxVerificaAsiento = True
vMensaje = ""

If CCur(txtTotal) <> 100 Then vMensaje = vMensaje & vbCrLf & "- El porcentaje de las cuentas no esta blanceado%"


For lng = 1 To vGrid.MaxRows
 vGrid.Row = lng
 vGrid.col = 1
 If vGrid.Text <> "" Then
   vGrid.col = 2
   If vGrid.Text = "" Then
      vGrid.col = 1
      vMensaje = vMensaje & vbCrLf & "- Cuenta " & vGrid.Text & " No Existe"
   End If
 End If
Next lng

If Len(vMensaje) > 0 Then
  fxVerificaAsiento = False
  MsgBox vMensaje, vbCritical
End If

End Function


Private Sub sbLimpiaPantalla()
vBusca = 1

txtCodigo = ""
txtDescripcion = ""

txtTotal = 0

txtNotas = ""

chkActivo.Value = vbChecked

vGrid.MaxRows = 0
vGrid.MaxRows = 1
vGrid.MaxCols = 4

sBar.Panels(1) = ""
sBar.Panels(2) = ""

End Sub


Private Sub tlb_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim strSQL As String

Select Case UCase(Button.Key)
    Case "INSERTAR", "NUEVO"
      vEdita = False
      Call sbLimpiaPantalla
      Call sbToolBar(tlb, "edicion")
    
      txtCodigo.SetFocus
    
    Case "MODIFICAR", "EDITAR"
        vEdita = True
        txtDescripcion.SetFocus
        Call sbToolBar(tlb, "edicion")
    
    Case "BORRAR"
        Call sbBorrar
      
    Case "GUARDAR", "SALVAR"
      Call sbGuardar
    
    Case "DESHACER"
        Call sbLimpiaPantalla
        Call sbToolBar(tlb, "nuevo")
        vEdita = True
    
    Case "CONSULTAR"
       Select Case vBusca
         Case 3, 4 'Codigo o Descripcion
            If vBusca = 3 Then
                gBusquedas.Columna = "cod_concepto"
                gBusquedas.Orden = "cod_concepto"
            Else
                gBusquedas.Columna = "Descripcion"
                gBusquedas.Orden = "Descripcion"
            End If
            gBusquedas.Consulta = "select cod_concepto,descripcion from cxp_FacConceptos"
            gBusquedas.Filtro = ""
            frmBusquedas.Show vbModal
            txtCodigo = gBusquedas.Resultado
            txtDescripcion = gBusquedas.Resultado2
            txtCodigo.SetFocus
       
       End Select

    Case "REPORTES"
      
'      strSQL = "{ASIENTOS.COD_EMPRESA} = " & vParametros.CodigoEmpresa _
'             & " AND {ASIENTOS.TIPO_ASIENTO} = '" & txtCAsiento & "' AND " _
'             & " {ASIENTOS.NUM_ASIENTO} = '" & txtNAsiento & "'"
'
'      Call sbReportes("ASIENTO", strSQL)
    
    Case "AYUDA"
        frmContenedor.CD.HelpContext = Me.HelpContextID
        frmContenedor.CD.ShowHelp
    
    Case "CERRAR"
      Unload Me
End Select

End Sub


Private Sub sbCargaGridLocal(vGrid As Object, vGridMaxCol As Integer, strSQL As String)
Dim rs As New ADODB.Recordset, i As Integer

Me.MousePointer = vbHourglass

vGrid.MaxCols = vGridMaxCol
vGrid.MaxRows = 1

vGrid.Row = vGrid.MaxRows

rs.CursorLocation = adUseServer
Call OpenRecordSet(rs, strSQL)

Do While Not rs.EOF
  vGrid.Row = vGrid.MaxRows
  For i = 1 To vGrid.MaxCols
    vGrid.col = i
    Select Case i
      Case 1
        vGrid.Text = fxgCntCuentaFormato(True, CStr(rs.Fields(i - 1).Value))
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


Private Sub sbConsulta(vCodigo As String)
Dim rs As New ADODB.Recordset, strSQL As String

On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = "select * from cxp_facConceptos where cod_concepto = '" & vCodigo & "'"
       
Call OpenRecordSet(rs, strSQL)

If Not rs.BOF And Not rs.EOF Then
  Call sbToolBar(tlb, "activo")
  vEdita = True
 
  'llenar datos en pantalla
  
  txtCodigo = rs!cod_Concepto
  txtDescripcion = rs!Descripcion & ""
  
  txtNotas = rs!Notas
 
  sBar.Panels(1) = rs!Usuario
  sBar.Panels(2) = rs!fecha
  
strSQL = "select A.cod_cuenta,B.descripcion,porcentaje,Idx" _
       & " from cxp_FacCuentas A inner join CntX_cuentas B on A.cod_cuenta = B.cod_cuenta" _
       & " where B.cod_contabilidad = " & GLOBALES.gEnlace _
       & " and A.cod_concepto = " & rs!cod_Concepto _
       & " order by IDx"
   
  Call sbCargaGridLocal(vGrid, 4, strSQL)
 
  Call sbSumaDebitosCreditos

End If

rs.Close
Me.MousePointer = vbDefault

Exit Sub
vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub sbGuardar()
Dim strSQL As String, rs As New ADODB.Recordset, lng As Long

On Error GoTo vError

If fxVerificaAsiento Then
    
    If vEdita Then
      
      strSQL = "update cxp_FacConceptos set descripcion = '" & UCase(txtDescripcion) _
             & "',notas = '" & Trim(txtNotas) _
             & "',Activo = " & chkActivo.Value & " where cod_concepto = '" & txtCodigo & "'"
      Call ConectionExecute(strSQL)
     
      strSQL = "delete cxp_FacCuentas where cod_concepto = '" & txtCodigo & "'"
      Call ConectionExecute(strSQL)
    
      For lng = 1 To vGrid.MaxRows
        vGrid.Row = lng
        vGrid.col = 1
        If vGrid.Text <> "" Then
            strSQL = "insert into cxp_FacCuentas(cod_concepto" _
                   & ",Idx,cod_cuenta,porcentaje) values('" & txtCodigo _
                   & "'," & lng & ",'"
            vGrid.Row = lng
            vGrid.col = 1
            strSQL = strSQL & fxgCntCuentaFormato(False, vGrid.Text) & "',"
            vGrid.col = 3
            strSQL = strSQL & CCur(IIf((vGrid.Text = ""), 0, vGrid.Text)) & ")"
            Call ConectionExecute(strSQL)
              
         End If 'vgrid.Text <> ""
       Next lng
    
'      Call Bitacora("Modifica", "Plantilla Diferido : " & txtCodigo & " Emp" & vParametros.CodigoEmpresa)
    
    
    Else 'Inserta
      
       strSQL = "insert into cxp_FacConceptos(cod_concepto,descripcion,notas" _
              & ",usuario,fecha,activo) values('" & txtCodigo & "','" & UCase(txtDescripcion) _
              & "','" & txtNotas & "','" & glogon.Usuario & "',dbo.MyGetdate()," & chkActivo.Value & ")"
       Call ConectionExecute(strSQL)
      
      strSQL = "delete cxp_FacCuentas where cod_concepto = '" & txtCodigo & "'"
      Call ConectionExecute(strSQL)
    
      For lng = 1 To vGrid.MaxRows
        vGrid.Row = lng
        vGrid.col = 1
        If vGrid.Text <> "" Then
            strSQL = "insert into cxp_FacCuentas(cod_concepto" _
                   & ",Idx,cod_cuenta,porcentaje) values('" & txtCodigo _
                   & "'," & lng & ",'"
            vGrid.Row = lng
            vGrid.col = 1
            strSQL = strSQL & fxgCntCuentaFormato(False, vGrid.Text) & "',"
            vGrid.col = 3
            strSQL = strSQL & CCur(IIf((vGrid.Text = ""), 0, vGrid.Text)) & ")"
            Call ConectionExecute(strSQL)
              
         End If 'vgrid.Text <> ""
       Next lng
       
'       Call Bitacora("Registra", "Plantilla Diferido : " & txtCodigo & " Emp" & vParametros.CodigoEmpresa)
        
    End If 'Si Inserta o Actualiza

        Call sbToolBar(tlb, "activo")
        Call sbConsulta(txtCodigo)
        
        vEdita = True
        
        MsgBox "Información guardada satisfactoriamente...", vbInformation


End If 'Verificacion del Asiento

 Call RefrescaTags(Me)

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
 Resume

End Sub

Private Sub sbBorrar()
Dim i As Integer, strSQL As String
On Error GoTo vError
i = MsgBox("Esta Seguro que desea borrar este registro", vbYesNo)
If i = vbYes Then
  strSQL = "delete cxp_FacCuentas where cod_concepto = '" & txtCodigo & "'"
  Call ConectionExecute(strSQL)
  
  strSQL = "delete cxp_FacConceptos where cod_concepto = '" & txtCodigo & "'"
  Call ConectionExecute(strSQL)
  

'  Call Bitacora("Elimina", "Plantilla Diferidos : " & txtCodigo & " EMP:" _
                  & vParametros.CodigoEmpresa)

  Call sbLimpiaPantalla
  Call sbToolBar(tlb, "nuevo")
End If

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub txtCodigo_GotFocus()
vBusca = 3
End Sub

Private Sub txtDescripcion_GotFocus()
vBusca = 4
End Sub

Private Sub txtDescripcion_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtNotas.SetFocus
End Sub

Private Sub txtCodigo_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo vError
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then
 Call sbConsulta(txtCodigo)
 txtDescripcion.SetFocus
End If
If KeyCode = vbKeyF4 Then Call tlb_ButtonClick(tlb.Buttons(7))
Exit Sub
vError:
  Call sbLimpiaPantalla
End Sub


Private Sub sbSumaDebitosCreditos()
Dim x As Long, curValor As Currency

  txtTotal = 0
   
  For x = 1 To vGrid.MaxRows
    vGrid.Row = x
    vGrid.col = 1
    If vGrid.Text <> "" Then
      vGrid.col = 3
      txtTotal = CCur(txtTotal) + CCur(IIf(vGrid.Text = "", 0, vGrid.Text))
    End If 'vGrid.text <> ""
  Next x
  txtTotal = Format(txtTotal, "Standard")

End Sub

Private Sub txtNotas_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then vGrid.SetFocus
End Sub

Private Sub vGrid_KeyDown(KeyCode As Integer, Shift As Integer)
Dim i As Variant, lng As Long, vTemp(7) As Variant, x As Integer

If KeyCode = vbKeyDelete Then
  
  vGrid.Row = vGrid.ActiveRow
  vGrid.col = vGrid.MaxCols
  If vGrid.Text <> "" Then 'Existe en la Base de datos
    'Preguntar y si la respuesta es afirmativa eliminar de la Base de datos
  
  
  End If
  
  For lng = vGrid.ActiveRow To vGrid.MaxRows
     vGrid.Row = lng + 1
     For x = 1 To vGrid.MaxCols
        vGrid.col = x
        vTemp(x) = vGrid.Text
     Next x
     
     vGrid.Row = lng
     For x = 1 To vGrid.MaxCols
       vGrid.col = x
       vGrid.Text = vTemp(x)
     Next x
  Next lng
  vGrid.MaxRows = vGrid.MaxRows - 1
  If vGrid.MaxRows = 0 Then vGrid.MaxRows = 1
  
  Call sbSumaDebitosCreditos
  
  
End If

If KeyCode = vbKeyF4 And vGrid.ActiveCol = 1 Then
  frmCntX_ConsultaCuentas.Show vbModal
  vGrid.col = vGrid.ActiveCol
  vGrid.Row = vGrid.ActiveRow
  vGrid.Text = gCuenta
End If

If (KeyCode = vbKeyReturn Or KeyCode = vbKeyTab) Then
    vGrid.col = vGrid.ActiveCol
    vGrid.Row = vGrid.ActiveRow
    
    Select Case vGrid.ActiveCol
      Case 1 'Cuenta
        vGrid.Text = fxgCntCuentaFormato(True, vGrid.Text)
        i = fxgCntCuentaFormato(False, vGrid.Text)
        If fxgCntCuentaValida(CStr(i)) Then
          vGrid.col = 2
          vGrid.Text = fxSIFCCodigos("D", fxgCntCuentaFormato(False, CStr(i)), "cuentas")
        Else
          MsgBox "Cuenta no es válida : " & vbCrLf & " - No Existe o No Acepta Movimientos" _
                 & vbCrLf & " - VERIFIQUE O MODIFIQUE EN EL CATALAGO DE CUENTAS", vbCritical
        End If
        
      Case 3 'Debe
        If Val(vGrid.Text) > 0 Then
            vGrid.col = vGrid.ActiveCol + 1
            vGrid.Row = vGrid.ActiveRow
            vGrid.Text = 0
        
            Call sbSumaDebitosCreditos
            
        End If
      
      Case vGrid.MaxCols  'Nueva linea
        If vGrid.MaxRows = vGrid.Row Then
            vGrid.MaxRows = vGrid.MaxRows + 1
        End If
    End Select
End If


If KeyCode = vbKeyInsert Then
    vGrid.MaxRows = vGrid.MaxRows + 1
    vGrid.InsertRows vGrid.ActiveRow, 1
End If

End Sub


