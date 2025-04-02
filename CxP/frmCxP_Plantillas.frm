VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#20.3#0"; "Codejock.Controls.v20.3.0.ocx"
Begin VB.Form frmCxPPlantillas 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Plantillas para Facturas"
   ClientHeight    =   6420
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10500
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6420
   ScaleWidth      =   10500
   Begin XtremeSuiteControls.CheckBox chkActivo 
      Height          =   255
      Left            =   2640
      TabIndex        =   7
      Top             =   1080
      Width           =   1935
      _Version        =   1310723
      _ExtentX        =   3413
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Plantilla Activa?"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
      Appearance      =   17
   End
   Begin MSComctlLib.Toolbar tlb 
      Align           =   1  'Align Top
      Height          =   330
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10500
      _ExtentX        =   18521
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
   Begin FPSpreadADO.fpSpread vGrid 
      Height          =   2895
      Left            =   120
      TabIndex        =   1
      Top             =   2400
      Width           =   10215
      _Version        =   524288
      _ExtentX        =   18018
      _ExtentY        =   5106
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
      SpreadDesigner  =   "frmCxP_Plantillas.frx":0000
      VScrollSpecial  =   -1  'True
      VScrollSpecialType=   2
      AppearanceStyle =   1
   End
   Begin MSComCtl2.FlatScrollBar FlatScrollBar 
      Height          =   255
      Left            =   9360
      TabIndex        =   2
      Top             =   600
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   450
      _Version        =   393216
      Arrows          =   65536
      Orientation     =   1638401
   End
   Begin MSComctlLib.StatusBar sBar 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   6
      Top             =   6105
      Width           =   10500
      _ExtentX        =   18521
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   6068
            MinWidth        =   6068
            Object.ToolTipText     =   "Usuario"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   6068
            MinWidth        =   6068
            Object.ToolTipText     =   "Fecha y Hora de la Creación"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   3528
            MinWidth        =   3528
            Text            =   "Plantillas"
            TextSave        =   "Plantillas"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin XtremeSuiteControls.FlatEdit txtCodigo 
      Height          =   315
      Left            =   1320
      TabIndex        =   8
      Top             =   600
      Width           =   1335
      _Version        =   1310723
      _ExtentX        =   2355
      _ExtentY        =   556
      _StockProps     =   77
      ForeColor       =   0
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
      Left            =   2640
      TabIndex        =   9
      Top             =   600
      Width           =   6615
      _Version        =   1310723
      _ExtentX        =   11668
      _ExtentY        =   556
      _StockProps     =   77
      ForeColor       =   0
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
      Left            =   1320
      TabIndex        =   10
      Top             =   1440
      Width           =   7935
      _Version        =   1310723
      _ExtentX        =   13996
      _ExtentY        =   1191
      _StockProps     =   77
      ForeColor       =   0
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
   Begin XtremeSuiteControls.FlatEdit txtTotal 
      Height          =   315
      Left            =   8880
      TabIndex        =   11
      Top             =   5520
      Width           =   1095
      _Version        =   1310723
      _ExtentX        =   1931
      _ExtentY        =   556
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   16777152
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16777152
      Alignment       =   1
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Notas"
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
      Left            =   240
      TabIndex        =   5
      Top             =   1440
      Width           =   945
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Total:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   7800
      TabIndex        =   4
      Top             =   5520
      Width           =   765
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Plantilla"
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
      Left            =   240
      TabIndex        =   3
      Top             =   600
      Width           =   1005
   End
End
Attribute VB_Name = "frmCxPPlantillas"
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
    strSQL = "select Top 1 cod_plantilla from CxP_Plantillas"
           
    If FlatScrollBar.Value = 1 Then
       strSQL = strSQL & " where cod_plantilla > '" & txtCodigo.Text & "' order by cod_plantilla asc"
    Else
       strSQL = strSQL & " where cod_plantilla < '" & txtCodigo.Text & "' order by cod_plantilla desc"
    End If
    
    Call OpenRecordSet(rs, strSQL)
    If Not rs.EOF And Not rs.BOF Then
      Call sbConsulta(rs!cod_plantilla)
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

Private Sub Form_Activate()
vModulo = 30
End Sub

Private Sub Form_Load()

vModulo = 30

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
Dim vMensaje As String, lng As Long

vMensaje = ""

If CCur(txtTotal) <> 100 Then vMensaje = vMensaje & vbCrLf & "- El porcentaje de las cuentas no esta blanceado%"


For lng = 1 To vGrid.MaxRows
 vGrid.Row = lng
 vGrid.col = 1
 If vGrid.Text <> "" Then
   vGrid.col = 1
   If Not fxgCntCuentaValida(fxgCntCuentaFormato(False, vGrid.Text)) Then
      vMensaje = vMensaje & vbCrLf & "Línea " & lng & ": La cuenta no es válida."
   End If
   
   vGrid.col = 3
   If vGrid.Text <> "" Then
     If fxgCntUnidad(vGrid.Text) = "" Then
          vMensaje = vMensaje & vbCrLf & "Línea " & lng & ": La unidad no es válida."
     End If
   Else
      vMensaje = vMensaje & vbCrLf & "Línea " & lng & ": No se especificó Unidad Contable."
   End If
   
   vGrid.col = 4
   If vGrid.Text <> "" Then
     If fxgCntCentroCostos(vGrid.Text) = "" Then
          vMensaje = vMensaje & vbCrLf & "Línea " & lng & ": Centro de Costo no es válido."
     End If
   End If
   
   vGrid.col = 5
   If Not IsNumeric(vGrid.Text) Then
          vMensaje = vMensaje & vbCrLf & "Línea " & lng & ": El % indicado no es válido."
   End If
 End If
Next lng

If Len(vMensaje) > 0 Then
  fxVerificaAsiento = False
  MsgBox vMensaje, vbCritical
Else
  fxVerificaAsiento = True
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
vGrid.MaxCols = 5

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
                gBusquedas.Columna = "cod_plantilla"
                gBusquedas.Orden = "cod_plantilla"
            Else
                gBusquedas.Columna = "Descripcion"
                gBusquedas.Orden = "Descripcion"
            End If
            gBusquedas.Consulta = "select cod_plantilla,descripcion from CxP_Plantillas"
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

strSQL = "select * from CxP_Plantillas where cod_plantilla = '" & vCodigo & "'"
       
Call OpenRecordSet(rs, strSQL)

If Not rs.BOF And Not rs.EOF Then
  Call sbToolBar(tlb, "activo")
  vEdita = True
 
  'llenar datos en pantalla
  
  txtCodigo = rs!cod_plantilla
  txtDescripcion = rs!Descripcion & ""
  
  txtNotas = rs!Notas
 
  sBar.Panels(1) = rs!Registro_Usuario
  sBar.Panels(2) = rs!Registro_Fecha
  
strSQL = "select A.cod_cuenta,B.descripcion,A.cod_unidad,A.cod_Centro_Costo,A.porcentaje" _
       & " from CxP_Plantillas_Asiento A inner join CntX_cuentas B on A.cod_cuenta = B.cod_cuenta" _
       & " and A.cod_contabilidad = B.cod_contabilidad" _
       & " where B.cod_contabilidad = " & GLOBALES.gEnlace _
       & " and A.cod_plantilla = " & rs!cod_plantilla _
       & " order by Linea"
   
  Call sbCargaGridLocal(vGrid, 5, strSQL)
 
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


If Len(txtCodigo.Text) = 0 Then
    MsgBox "Indique el Código de la Plantilla y el Nombre", vbExclamation
    Exit Sub
End If

If Not fxVerificaAsiento Then
   Exit Sub
End If
    
On Error GoTo vError
    
If vEdita Then
 
 strSQL = "update CxP_Plantillas set descripcion = '" & Trim(txtDescripcion.Text) _
        & "',notas = '" & Trim(txtNotas.Text) _
        & "',Activo = " & chkActivo.Value & " where cod_plantilla = '" & txtCodigo & "'"
 Call ConectionExecute(strSQL)

 Call Bitacora("Modifica", "Plantilla CxP : " & txtCodigo)


Else 'Inserta
  strSQL = "insert into CxP_Plantillas(cod_plantilla,descripcion,notas" _
         & ",Registro_Usuario,registro_fecha,activo) values('" & txtCodigo.Text & "','" & Trim(txtDescripcion.Text) _
         & "','" & txtNotas & "','" & glogon.Usuario & "',dbo.MyGetdate()," & chkActivo.Value & ")"
  Call ConectionExecute(strSQL)
 
  Call Bitacora("Registra", "Plantilla CxP : " & txtCodigo)
   
End If 'Si Inserta o Actualiza

'Guarda el Detalle del Asiento
strSQL = "delete CxP_Plantillas_Asiento where cod_plantilla = '" & txtCodigo & "'"
Call ConectionExecute(strSQL)

For lng = 1 To vGrid.MaxRows
  vGrid.Row = lng
  vGrid.col = 1
  If vGrid.Text <> "" Then
      strSQL = "insert into CxP_Plantillas_Asiento(Linea,cod_plantilla,cod_cuenta,cod_contabilidad,cod_divisa,cod_unidad" _
             & ",cod_centro_costo,porcentaje) values(" & lng & ",'" & txtCodigo & "','"
      vGrid.col = 1
      strSQL = strSQL & fxgCntCuentaFormato(False, vGrid.Text) & "'," & GLOBALES.gEnlace & ",'COL','"
      vGrid.col = 3
      strSQL = strSQL & UCase(Trim(vGrid.Text)) & "','"
      vGrid.col = 4
      strSQL = strSQL & UCase(Trim(vGrid.Text)) & "',"
      vGrid.col = 5
      strSQL = strSQL & CCur(IIf((vGrid.Text = ""), 0, vGrid.Text)) & ")"
      Call ConectionExecute(strSQL)
   End If 'vgrid.Text <> ""
 Next lng


Call sbToolBar(tlb, "activo")
Call sbConsulta(txtCodigo)

vEdita = True

MsgBox "Información guardada satisfactoriamente...", vbInformation

Call RefrescaTags(Me)

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub sbBorrar()
Dim i As Integer, strSQL As String

On Error GoTo vError

i = MsgBox("Esta Seguro que desea borrar este registro", vbYesNo)

If i = vbYes Then
  strSQL = "delete CxP_Plantillas_Asiento where cod_plantilla = '" & txtCodigo & "'"
  Call ConectionExecute(strSQL)
  
  strSQL = "delete CxP_Plantillas where cod_plantilla = '" & txtCodigo & "'"
  Call ConectionExecute(strSQL)
  

  Call Bitacora("Elimina", "Plantilla CxP : " & txtCodigo)

  MsgBox "Plantilla de CxP Eliminada Satisfactoriamente...!", vbInformation

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
      vGrid.col = 5
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
  vGrid.DeleteRows vGrid.ActiveRow, 1
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

'Consulta unidad
If KeyCode = vbKeyF4 And vGrid.ActiveCol = 3 Then
  gBusquedas.Resultado = ""
  gBusquedas.Resultado2 = ""
  gBusquedas.Columna = "descripcion"
  gBusquedas.Orden = "descripcion"
  gBusquedas.Filtro = " and Activa = 1 and cod_contabilidad = " & GLOBALES.gEnlace
  gBusquedas.Consulta = "select cod_unidad,descripcion from CntX_unidades"
  frmBusquedas.Show vbModal
    
  vGrid.col = vGrid.ActiveCol
  vGrid.Row = vGrid.ActiveRow
  vGrid.Text = gBusquedas.Resultado
  
  vGrid.CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
  vGrid.CellNote = gBusquedas.Resultado2
  vGrid.TextTip = TextTipFixed
  
End If



'Consulta Centro de Costo
If KeyCode = vbKeyF4 And vGrid.ActiveCol = 4 Then
  
  vGrid.col = 2
  gBusquedas.Resultado = ""
  gBusquedas.Resultado2 = ""
  gBusquedas.Filtro = " and C.cod_Contabilidad = " & GLOBALES.gEnlace
  gBusquedas.Columna = "descripcion"
  gBusquedas.Orden = "descripcion"
  gBusquedas.Consulta = "select C.COD_CENTRO_COSTO,C.descripcion" _
                      & " from CNTX_CENTRO_COSTOS C inner join CNTX_UNIDADES_CC A on C.COD_CENTRO_COSTO = A.COD_CENTRO_COSTO" _
                      & " and C.cod_contabilidad = A.cod_Contabilidad" _
                      & " and A.cod_unidad = '" & vGrid.Text & "'"
  frmBusquedas.Show vbModal
    
  vGrid.col = vGrid.ActiveCol
  vGrid.Row = vGrid.ActiveRow
  vGrid.Text = gBusquedas.Resultado
  
  vGrid.CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
  vGrid.CellNote = gBusquedas.Resultado2
  vGrid.TextTip = TextTipFixed
  
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
          vGrid.Text = fxgCntCuentaDesc(CStr(i))
        Else
          MsgBox "Cuenta no es válida : " & vbCrLf & " - No Existe o No Acepta Movimientos" _
                 & vbCrLf & " - VERIFIQUE O MODIFIQUE EN EL CATALAGO DE CUENTAS", vbCritical
        End If
      
      Case vGrid.MaxCols  'Nueva linea
         If vGrid.MaxRows = vGrid.ActiveRow Then
            vGrid.MaxRows = vGrid.MaxRows + 1
         End If
         Call sbSumaDebitosCreditos
    
    End Select
End If


If KeyCode = vbKeyInsert Then
    vGrid.MaxRows = vGrid.MaxRows + 1
    vGrid.InsertRows vGrid.ActiveRow, 1
End If



End Sub




