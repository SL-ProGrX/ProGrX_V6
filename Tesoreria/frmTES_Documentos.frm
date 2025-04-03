VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.0#0"; "Codejock.Controls.v22.0.0.ocx"
Begin VB.Form frmTES_Documentos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Documentos"
   ClientHeight    =   4635
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8655
   Icon            =   "frmTES_Documentos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4635
   ScaleWidth      =   8655
   Begin XtremeSuiteControls.GroupBox GroupBox1 
      Height          =   3492
      Left            =   120
      TabIndex        =   5
      Top             =   960
      Width           =   8412
      _Version        =   1441792
      _ExtentX        =   14838
      _ExtentY        =   6159
      _StockProps     =   79
      Caption         =   "Enlace Contable"
      ForeColor       =   8421504
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
      BorderStyle     =   1
      Begin MSComCtl2.FlatScrollBar FlatScrollBarT 
         Height          =   252
         Left            =   7680
         TabIndex        =   6
         Top             =   480
         Width           =   492
         _ExtentX        =   873
         _ExtentY        =   450
         _Version        =   393216
         Arrows          =   65536
         Orientation     =   1638401
      End
      Begin XtremeSuiteControls.FlatEdit txtTipoAsiento 
         Height          =   312
         Left            =   1320
         TabIndex        =   9
         Top             =   480
         Width           =   1572
         _Version        =   1441792
         _ExtentX        =   2773
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
         Appearance      =   2
      End
      Begin XtremeSuiteControls.FlatEdit txtTipoAsientoDesc 
         Height          =   312
         Left            =   2880
         TabIndex        =   10
         Top             =   480
         Width           =   4692
         _Version        =   1441792
         _ExtentX        =   8276
         _ExtentY        =   550
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
         Locked          =   -1  'True
         Appearance      =   2
      End
      Begin XtremeSuiteControls.ComboBox cboMov 
         Height          =   312
         Left            =   5040
         TabIndex        =   11
         Top             =   840
         Width           =   2532
         _Version        =   1441792
         _ExtentX        =   4471
         _ExtentY        =   582
         _StockProps     =   77
         ForeColor       =   1973790
         BackColor       =   16185078
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16185078
         Style           =   2
         Appearance      =   16
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.FlatEdit txtMascara 
         Height          =   312
         Left            =   3000
         TabIndex        =   13
         Top             =   2520
         Width           =   1572
         _Version        =   1441792
         _ExtentX        =   2773
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
         Appearance      =   2
      End
      Begin XtremeSuiteControls.CheckBox chkGenera 
         Height          =   252
         Left            =   1320
         TabIndex        =   14
         Top             =   1440
         Width           =   4332
         _Version        =   1441792
         _ExtentX        =   7641
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Genera Asientos a la Contabilidad"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Transparent     =   -1  'True
         UseVisualStyle  =   -1  'True
         Appearance      =   16
      End
      Begin XtremeSuiteControls.CheckBox chkAsTransac 
         Height          =   252
         Left            =   1920
         TabIndex        =   15
         Top             =   1800
         Width           =   3612
         _Version        =   1441792
         _ExtentX        =   6371
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Genera Asientos x Transacción"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Transparent     =   -1  'True
         UseVisualStyle  =   -1  'True
         Appearance      =   16
      End
      Begin XtremeSuiteControls.CheckBox chkAsFormato 
         Height          =   252
         Left            =   1920
         TabIndex        =   16
         Top             =   2160
         Width           =   4692
         _Version        =   1441792
         _ExtentX        =   8276
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Establece Formato al Número de Asiento"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Transparent     =   -1  'True
         UseVisualStyle  =   -1  'True
         Appearance      =   16
      End
      Begin XtremeSuiteControls.CheckBox chkAsIDBanco 
         Height          =   372
         Left            =   3000
         TabIndex        =   17
         Top             =   2880
         Width           =   5292
         _Version        =   1441792
         _ExtentX        =   9334
         _ExtentY        =   656
         _StockProps     =   79
         Caption         =   "Incluir el [ID] de la Cuenta Bancaria en el Formato"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Transparent     =   -1  'True
         UseVisualStyle  =   -1  'True
         Appearance      =   16
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Mascara para el Formato"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   0
         Left            =   4680
         TabIndex        =   12
         ToolTipText     =   "Cuenta de Inventarios para la Bodega"
         Top             =   2520
         Width           =   2532
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Afectación Contable"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   1
         Left            =   2160
         TabIndex        =   8
         Top             =   840
         Width           =   2412
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo Asiento"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   1
         Left            =   0
         TabIndex        =   7
         ToolTipText     =   "Cuenta de Inventarios para la Bodega"
         Top             =   480
         Width           =   1332
      End
   End
   Begin MSComctlLib.Toolbar tlb 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
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
   Begin MSComCtl2.FlatScrollBar FlatScrollBar 
      Height          =   255
      Left            =   7800
      TabIndex        =   2
      Top             =   480
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   450
      _Version        =   393216
      Arrows          =   65536
      Orientation     =   1638401
   End
   Begin XtremeSuiteControls.FlatEdit txtCodigo 
      Height          =   312
      Left            =   1440
      TabIndex        =   3
      Top             =   480
      Width           =   1572
      _Version        =   1441792
      _ExtentX        =   2773
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
      Appearance      =   2
   End
   Begin XtremeSuiteControls.FlatEdit txtNombre 
      Height          =   312
      Left            =   3000
      TabIndex        =   4
      Top             =   480
      Width           =   4692
      _Version        =   1441792
      _ExtentX        =   8276
      _ExtentY        =   550
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
      Appearance      =   2
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Documento"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   1332
   End
End
Attribute VB_Name = "frmTES_Documentos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vEdita As Boolean, vCodigo As String
Dim vScroll As Boolean, vScrollT As Boolean

Private Sub chkAsFormato_Click()
If chkAsFormato.Value = vbChecked Then
   txtMascara.Enabled = True
   chkAsIDBanco.Enabled = True
Else
   txtMascara.Enabled = False
   chkAsIDBanco.Enabled = False
End If
End Sub

Private Sub chkGenera_Click()
If chkGenera.Value = vbChecked Then
  chkAsFormato.Enabled = True
  chkAsTransac.Enabled = True
Else
  chkAsFormato.Enabled = False
  chkAsTransac.Enabled = False
End If

Call chkAsFormato_Click

End Sub

Private Sub FlatScrollBar_Change()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

If vScroll Then
    strSQL = "select Tipo from tes_tipos_doc"
    
    If FlatScrollBar.Value = 1 Then
       strSQL = strSQL & " where Tipo > '" & txtCodigo.Text & "' order by Tipo asc"
    Else
       strSQL = strSQL & " where Tipo < '" & txtCodigo.Text & "' order by Tipo desc"
    End If
    
    Call OpenRecordSet(rs, strSQL)
    If Not rs.EOF And Not rs.BOF Then
      txtCodigo.Text = rs!Tipo
      Call sbConsulta(txtCodigo)
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

Private Sub FlatScrollBarT_Change()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

If vScrollT Then
    strSQL = "select Tipo_Asiento,descripcion from CNTX_TIPOS_ASIENTOS"
    
    If FlatScrollBarT.Value = 1 Then
       strSQL = strSQL & " where Tipo_Asiento > '" & txtTipoAsiento.Text & "' and cod_contabilidad = " _
              & GLOBALES.gEnlace & " and activo = 1 order by Tipo_Asiento asc"
    Else
       strSQL = strSQL & " where Tipo_Asiento < '" & txtTipoAsiento.Text & "' and cod_contabilidad = " _
              & GLOBALES.gEnlace & " and activo = 1 order by Tipo_Asiento desc"
    End If
    
    Call OpenRecordSet(rs, strSQL)
    If Not rs.EOF And Not rs.BOF Then
            txtTipoAsiento.Text = rs!Tipo_Asiento
            txtTipoAsientoDesc.Text = rs!Descripcion
    
    End If
    rs.Close
End If

vScrollT = False
FlatScrollBarT.Value = 0
vScrollT = True

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub Form_Activate()
 vModulo = 9
End Sub

Private Sub Form_Load()

On Error GoTo vError
 
 vModulo = 9
 
 vScroll = False
 FlatScrollBar.Value = 0
 vScroll = True
 
 vScrollT = False
 FlatScrollBarT.Value = 0
 vScrollT = True
 
 

cboMov.Clear
cboMov.AddItem "Debita"
cboMov.AddItem "Acredita"
cboMov.Text = "Debita"

 
 vEdita = True
 
 Call sbToolBarIconos(tlb, False)
 Call sbToolBar(tlb, "nuevo")
 Call sbLimpiaPantalla

 Call Formularios(Me)
 Call RefrescaTags(Me)
 
Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbExclamation
  
End Sub

Private Sub sbLimpiaPantalla()
vCodigo = ""
txtCodigo = ""

cboMov.Text = "Debita"

txtNombre = ""

txtTipoAsiento = ""
txtTipoAsientoDesc.Text = ""

chkGenera.Value = vbChecked

chkAsFormato.Value = vbChecked
chkAsIDBanco.Value = vbChecked
chkAsTransac.Value = vbChecked

txtMascara.Text = "00000000"

End Sub



Private Sub tlb_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim strSQL As String

Select Case UCase(Button.Key)
    Case "INSERTAR", "NUEVO"
      vEdita = False
      Call sbLimpiaPantalla
      txtCodigo.SetFocus
      Call sbToolBar(tlb, "edicion")
    Case "MODIFICAR", "EDITAR"
      vEdita = True
      txtCodigo.SetFocus
      Call sbToolBar(tlb, "edicion")
    Case "BORRAR"
      Call sbBorrar
    Case "GUARDAR", "SALVAR"
     If fxValida Then Call sbGuardar
    Case "DESHACER"
      Call sbToolBar(tlb, "activo")
      If vCodigo = "" Then
        Call sbLimpiaPantalla
        Call sbToolBar(tlb, "nuevo")
        vEdita = True
      Else
        Call sbConsulta(vCodigo)
      End If
      
    Case "CONSULTAR"
       gBusquedas.Columna = "descripcion"
       gBusquedas.Orden = "descripcion"
       gBusquedas.Consulta = "select Tipo,descripcion from tes_tipos_doc"
       frmBusquedas.Show vbModal
       txtCodigo.SetFocus
       txtCodigo = gBusquedas.Resultado
       txtNombre.SetFocus
    
    Case "REPORTES"
    
    Case "AYUDA"
        frmContenedor.CD.HelpContext = Me.HelpContextID
        frmContenedor.CD.ShowHelp
   
End Select

End Sub

Private Sub sbConsulta(xCodigo As String)
Dim rs As New ADODB.Recordset, strSQL As String

On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = "select * from tes_tipos_doc where Tipo = '" & xCodigo & "'"
Call OpenRecordSet(rs, strSQL)

If Not rs.BOF And Not rs.EOF Then
  Call sbToolBar(tlb, "activo")
  vEdita = True
  
  vCodigo = rs!Tipo
  txtCodigo = rs!Tipo
 
  txtNombre = rs!Descripcion & ""
    
  If rs!Movimiento = "D" Then
    cboMov.Text = "Debita"
  Else
    cboMov.Text = "Acredita"
  End If
     
  txtTipoAsiento = rs!Tipo_Asiento
  txtTipoAsientoDesc.Text = fxgCntTipoAsientoDesc(rs!Tipo_Asiento)
  
  chkGenera.Value = rs!generacion
  chkAsFormato.Value = rs!asiento_formato
  chkAsTransac.Value = rs!asiento_transac
  chkAsIDBanco.Value = rs!asiento_banco
  
  txtMascara.Text = Trim(rs!asiento_mascara)
  
  Call chkGenera_Click

Else
  MsgBox "No se encontró registro verifique...", vbInformation
End If

rs.Close
Me.MousePointer = vbDefault
Call RefrescaTags(Me)

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Function fxValida() As Boolean
Dim vMensaje As String

vMensaje = ""
fxValida = True

'Validar Cuentas Aqui

If txtNombre = "" Then vMensaje = vMensaje & vbCrLf & " - Nombre del Documento no es válido ..."


If Len(vMensaje) > 0 Then
  fxValida = False
  MsgBox vMensaje, vbCritical
End If

End Function

Private Sub sbGuardar()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

If vEdita Then
  strSQL = "update tes_tipos_doc set descripcion = '" & UCase(Trim(txtNombre)) & "'" _
         & ",movimiento = '" & Mid(cboMov.Text, 1, 1) & "',Generacion = " & chkGenera.Value _
         & ",Tipo_asiento = '" & txtTipoAsiento & "',asiento_transac = " & chkAsTransac.Value _
         & ",asiento_formato = " & chkAsFormato.Value & ",asiento_banco = " & chkAsIDBanco.Value _
         & ",asiento_mascara = '" & txtMascara.Text & "' where Tipo = '" & vCodigo & "'"
  Call ConectionExecute(strSQL)
  Call Bitacora("Modifica", "Tipo de Documento : " & vCodigo)

Else
  vCodigo = txtCodigo
   
   strSQL = "insert into tes_tipos_doc(Tipo,descripcion,movimiento,generacion,tipo_asiento" _
          & ",asiento_transac,asiento_banco,asiento_formato,asiento_mascara)" _
          & " values('" & vCodigo & "','" & UCase(txtNombre) & "','" & Mid(cboMov.Text, 1, 1) & "'," _
          & chkGenera.Value & ",'" & txtTipoAsiento & "'," & chkAsTransac.Value & "," & chkAsIDBanco.Value _
          & "," & chkAsFormato.Value & ",'" & Trim(txtMascara.Text) & "')"
   Call ConectionExecute(strSQL)
    
   Call Bitacora("Registra", "Tipo de Documento: " & vCodigo)
 
End If

MsgBox "Información guardada satisfactoriamente...", vbInformation
Call sbToolBar(tlb, "activo")

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
  strSQL = "delete tes_tipos_doc where Tipo = '" & vCodigo & "'"
  Call ConectionExecute(strSQL)
  
  Call Bitacora("Elimina", "Tipo de Documento : " & vCodigo)
  Call sbLimpiaPantalla
  Call sbToolBar(tlb, "nuevo")
  Call RefrescaTags(Me)
End If

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub tlb_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
'Select Case ButtonMenu.Key
'  Case "LisTipo de Documentos"
'     Call sbReportesInv("Tipo de Documentos", "Tipo de DocumentoS", "Listado", "")
'  Case "InvTipo de Documentos"
'     Call sbReportesInv("InvTipo de Documentos", "Tipo de DocumentoS", "Inventario", "")
'End Select

End Sub

Private Sub txtCodigo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then
  If txtCodigo <> "" And vEdita Then Call sbConsulta(txtCodigo)
  txtNombre.SetFocus
End If

If KeyCode = vbKeyF4 Then
  gBusquedas.Convertir = "N"
  gBusquedas.Columna = "Tipo"
  gBusquedas.Orden = "Tipo"
  gBusquedas.Consulta = "select Tipo,descripcion from tes_tipos_doc"
  gBusquedas.Filtro = ""
  frmBusquedas.Show vbModal
  txtCodigo = gBusquedas.Resultado
  If txtCodigo <> "" Then Call sbConsulta(gBusquedas.Resultado)
End If

End Sub


Private Sub txtCodigo_LostFocus()
'txtNombre = fxgConCodigos("D", txtCodigo, "Tipo de Documentos")
End Sub


Private Sub txtTipoAsiento_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then cboMov.SetFocus
If KeyCode = vbKeyF4 Then
  Call sbgCntTipoAsientoConsulta("D")
  txtTipoAsiento = gBusquedas.Resultado
  txtTipoAsientoDesc.Text = gBusquedas.Resultado2
End If
End Sub

Private Sub txtTipoAsiento_LostFocus()
txtTipoAsientoDesc.Text = fxgCntTipoAsientoDesc(txtTipoAsiento)
End Sub


Private Sub txtNombre_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtTipoAsiento.SetFocus
If KeyCode = vbKeyF4 Then
  gBusquedas.Convertir = "N"
  gBusquedas.Columna = "descripcion"
  gBusquedas.Orden = "descripcion"
  gBusquedas.Consulta = "select Tipo,descripcion from tes_tipos_doc"
  gBusquedas.Filtro = ""
  frmBusquedas.Show vbModal
  txtCodigo = gBusquedas.Resultado
  If txtCodigo <> "" Then Call sbConsulta(gBusquedas.Resultado)
End If

End Sub




