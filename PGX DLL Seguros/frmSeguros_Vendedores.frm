VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.0#0"; "Codejock.Controls.v22.0.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.0#0"; "Codejock.ShortcutBar.v22.0.0.ocx"
Begin VB.Form frmSeguros_Vendedores 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Registro de Vendedores"
   ClientHeight    =   8115
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   10920
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8115
   ScaleWidth      =   10920
   ShowInTaskbar   =   0   'False
   Begin XtremeSuiteControls.ListView lswCuentas 
      Height          =   1575
      Left            =   240
      TabIndex        =   8
      Top             =   2880
      Width           =   10455
      _Version        =   1441792
      _ExtentX        =   18441
      _ExtentY        =   2778
      _StockProps     =   77
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
      View            =   3
      Appearance      =   17
      UseVisualStyle  =   0   'False
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   9120
      Top             =   600
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSeguros_Vendedores.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSeguros_Vendedores.frx":3492
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSeguros_Vendedores.frx":6924
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSeguros_Vendedores.frx":6A42
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin FPSpreadADO.fpSpread vGrid 
      Height          =   3135
      Left            =   240
      TabIndex        =   1
      Top             =   4920
      Width           =   10455
      _Version        =   524288
      _ExtentX        =   18441
      _ExtentY        =   5530
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
      SpreadDesigner  =   "frmSeguros_Vendedores.frx":6B79
      VScrollSpecialType=   2
      AppearanceStyle =   1
   End
   Begin MSComctlLib.Toolbar tlb 
      Height          =   330
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   8310
      _ExtentX        =   14658
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
   Begin MSComCtl2.FlatScrollBar FlatScrollBar 
      Height          =   255
      Left            =   10320
      TabIndex        =   4
      Top             =   600
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   450
      _Version        =   393216
      Arrows          =   65536
      Orientation     =   1638401
   End
   Begin XtremeSuiteControls.FlatEdit txtNombre 
      Height          =   315
      Left            =   4440
      TabIndex        =   5
      Top             =   600
      Width           =   5775
      _Version        =   1441792
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
   Begin XtremeSuiteControls.FlatEdit txtCodigo 
      Height          =   315
      Left            =   2400
      TabIndex        =   6
      Top             =   600
      Width           =   2055
      _Version        =   1441792
      _ExtentX        =   3619
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
   Begin XtremeSuiteControls.ComboBox cboComercializadora 
      Height          =   330
      Left            =   2400
      TabIndex        =   7
      Top             =   960
      Width           =   6735
      _Version        =   1441792
      _ExtentX        =   11880
      _ExtentY        =   582
      _StockProps     =   77
      ForeColor       =   1973790
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Style           =   2
      Appearance      =   6
      UseVisualStyle  =   0   'False
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.CheckBox chkActivo 
      Height          =   255
      Left            =   4680
      TabIndex        =   9
      Top             =   1680
      Width           =   1575
      _Version        =   1441792
      _ExtentX        =   2778
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Activo?"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   16
   End
   Begin XtremeSuiteControls.ComboBox cboBancos 
      Height          =   330
      Left            =   2400
      TabIndex        =   10
      Top             =   2040
      Width           =   6495
      _Version        =   1441792
      _ExtentX        =   11456
      _ExtentY        =   582
      _StockProps     =   77
      ForeColor       =   1973790
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Style           =   2
      Appearance      =   6
      UseVisualStyle  =   0   'False
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.PushButton btnCuentas 
      Height          =   375
      Left            =   9000
      TabIndex        =   11
      Top             =   2040
      Width           =   1695
      _Version        =   1441792
      _ExtentX        =   2984
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Cuentas Bancarias"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
      Appearance      =   17
   End
   Begin XtremeSuiteControls.ComboBox cboTipo 
      Height          =   315
      Left            =   2400
      TabIndex        =   12
      Top             =   2400
      Width           =   2055
      _Version        =   1441792
      _ExtentX        =   3625
      _ExtentY        =   582
      _StockProps     =   77
      ForeColor       =   1973790
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Style           =   2
      Appearance      =   6
      UseVisualStyle  =   0   'False
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.CheckBox chkComision 
      Height          =   255
      Left            =   2400
      TabIndex        =   13
      Top             =   1320
      Width           =   1935
      _Version        =   1441792
      _ExtentX        =   3408
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Aplica Comisión?"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   16
   End
   Begin XtremeSuiteControls.PushButton btnCuentasRefresh 
      Height          =   375
      Left            =   9000
      TabIndex        =   14
      Top             =   2400
      Width           =   1695
      _Version        =   1441792
      _ExtentX        =   2984
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Refrescar"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
      Appearance      =   17
   End
   Begin XtremeSuiteControls.FlatEdit txtIdentificacion 
      Height          =   315
      Left            =   2400
      TabIndex        =   15
      Top             =   1680
      Width           =   2055
      _Version        =   1441792
      _ExtentX        =   3619
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
      Alignment       =   2
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.CheckBox chkComisonInformativa 
      Height          =   255
      Left            =   4680
      TabIndex        =   20
      Top             =   1320
      Width           =   4815
      _Version        =   1441792
      _ExtentX        =   8493
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Comisión es informativa? (No aplica desembolso)"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   16
   End
   Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
      Height          =   375
      Left            =   240
      TabIndex        =   19
      Top             =   4560
      Width           =   10455
      _Version        =   1441792
      _ExtentX        =   18441
      _ExtentY        =   661
      _StockProps     =   14
      Caption         =   "Productos y Comisiones"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SubItemCaption  =   -1  'True
      Alignment       =   1
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo de Pago"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   5
      Left            =   240
      TabIndex        =   18
      Top             =   2400
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Identificación"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   13
      Left            =   240
      TabIndex        =   17
      Top             =   1680
      Width           =   1575
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Cuenta/Desembolsos"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   16
      Top             =   2040
      Width           =   2415
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Comercializadora"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   2
      Top             =   960
      Width           =   1815
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      Caption         =   "Vendedor"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Index           =   14
      Left            =   240
      TabIndex        =   0
      Top             =   600
      Width           =   2055
   End
End
Attribute VB_Name = "frmSeguros_Vendedores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vEdita As Boolean, vCodigo As String
Dim vScroll As Boolean, vPaso As Boolean


Private Sub btnCuentas_Click()
If txtCodigo.Text = "" Then
   MsgBox "Consulte una comercializadora Primero...", vbExclamation
   Exit Sub
End If

GLOBALES.gTag = Trim(txtIdentificacion.Text)
GLOBALES.gTag2 = "CdS"

frmCC_Cuentas_Bancarias.Show vbModal

Call sbCuentas_Load
End Sub



Private Sub sbCuentas_Load()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem

On Error GoTo vError

lswCuentas.ListItems.Clear
    strSQL = "select rtrim(B.Descripcion) as 'Banco'" _
           & ",case when C.tipo = 'A' then 'Ahorros' else 'Corriente' end as 'TipoDesc'" _
           & ",C.cod_Divisa,C.CUENTA_INTERNA, C.CUENTA_INTERBANCA, C.ACTIVA, C.DESTINO, C.REGISTRO_FECHA , C.REGISTRO_USUARIO" _
           & " from SYS_CUENTAS_BANCARIAS C inner join TES_BANCOS_GRUPOS B on C.cod_banco = B.cod_grupo" _
           & " where C.Identificacion = '" & Trim(txtIdentificacion.Text) & "' and C.Modulo = 'CdS'"
    
    Call OpenRecordSet(rs, strSQL)
    Do While Not rs.EOF
       Set itmX = lswCuentas.ListItems.Add(, , rs!CUENTA_INTERNA)
           itmX.SubItems(1) = Trim(rs!Banco)
           itmX.SubItems(2) = rs!TipoDesc
           itmX.SubItems(3) = rs!cod_Divisa
           itmX.SubItems(4) = IIf(rs!CUENTA_INTERBANCA = 1, "Sí", "No")
           itmX.SubItems(5) = rs!Destino & ""
           itmX.SubItems(6) = IIf(rs!Activa = 1, "Activa", "Cerrada")
           itmX.SubItems(7) = rs!Registro_Fecha & ""
           itmX.SubItems(8) = rs!Registro_Usuario & ""
     
       rs.MoveNext
    Loop
    rs.Close

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub



Private Sub btnCuentasRefresh_Click()
Dim strSQL As String

vPaso = True
    strSQL = "exec spCrd_SGT_Bancos '" & glogon.Usuario & "'"
    Call sbCbo_Llena_New(cboBancos, strSQL, False, True)
vPaso = False

Call sbCuentas_Load

End Sub



Private Sub cboBancos_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then cboTipo.SetFocus
End Sub

Private Sub cboTipo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then btnCuentas.SetFocus
End Sub

Private Sub FlatScrollBar_Change()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

If txtCodigo.Text = "" Or Not IsNumeric(txtCodigo.Text) Then
   txtCodigo.Text = 0
End If


If vScroll Then
    strSQL = "select Top 1 cod_vendedor from SEGUROS_Vendedores"
    
    If FlatScrollBar.Value = 1 Then
       strSQL = strSQL & " where cod_vendedor > " & txtCodigo.Text & " order by cod_vendedor asc"
    Else
       strSQL = strSQL & " where cod_vendedor < " & txtCodigo.Text & " order by cod_vendedor desc"
    End If
    
    Call OpenRecordSet(rs, strSQL)
    If Not rs.EOF And Not rs.BOF Then
      txtCodigo.Text = rs!cod_vendedor
      Call txtCodigo_LostFocus
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
vModulo = 17
End Sub

Private Sub Form_Load()
Dim strSQL As String

On Error GoTo vError

 vModulo = 17

 vEdita = True
 Call sbToolBarIconos(tlb)
 Call sbToolBar(tlb, "nuevo")
 
 
 vScroll = False
   FlatScrollBar.Value = 0
 vScroll = True
 

vEdita = False

 With lswCuentas.ColumnHeaders
    .Add 1, , "Cuenta", 2500
    .Add 2, , "Banco", 3500
    .Add 3, , "Tipo", 1100, vbCenter
    .Add 4, , "Divisa", 1100, vbCenter
    .Add 5, , "Interbanca", 1100, vbCenter
    .Add 6, , "Destino", 1100, vbCenter
    .Add 7, , "Activa", 1100, vbCenter
    .Add 8, , "Fecha", 2500
    .Add 9, , "Usuario", 2500
 End With

Call btnCuentasRefresh_Click

     
 strSQL = "select rtrim(cod_comercializadora) as 'IDx', rtrim(nombre) as 'ItmX' from seguros_comercializadoras where activo = 1"
 Call sbCbo_Llena_New(cboComercializadora, strSQL, False, True)

 strSQL = "select rtrim(Descripcion) as 'ItmX', Id_Banco as 'IdX' from Tes_Bancos where estado = 'A'"
 Call sbCbo_Llena_New(cboBancos, strSQL, False, True)

With cboTipo
    .Clear
    .AddItem "CK - Cheque"
    .AddItem "TE - Transferencia"
End With


 Call sbLimpiaPantalla

 Call Formularios(Me)
 Call RefrescaTags(Me)

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbExclamation

End Sub

Private Sub sbLimpiaPantalla()
Dim strSQL As String

vCodigo = 0
txtCodigo = ""

txtIdentificacion.Text = ""
txtNombre.Text = ""
cboTipo.Text = "TE - Transferencia"

chkActivo.Value = vbChecked

chkComision.Value = vbUnchecked
chkComisonInformativa.Value = vbUnchecked

vGrid.MaxRows = 0

lswCuentas.ListItems.Clear

End Sub



Private Sub tlb_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim strSQL As String

Select Case UCase(Button.Key)
    Case "InsertAR", "NUEVO"
      vEdita = False
      Call sbLimpiaPantalla
      txtNombre.SetFocus
      Call sbToolBar(tlb, "edicion")
    Case "MODIFICAR", "EDITAR"
      vEdita = True
      txtNombre.SetFocus
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
       gBusquedas.Columna = "nombre"
       gBusquedas.Orden = "nombre"
       gBusquedas.Consulta = "select cod_vendedor,nombre from SEGUROS_Vendedores"
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

Private Sub sbConsulta(pCodigo As String)
Dim rs As New ADODB.Recordset, strSQL As String


On Error GoTo vError

If Not fxSIFValidaCadena(pCodigo) Then
   Exit Sub
End If

Me.MousePointer = vbHourglass

strSQL = "select V.*,rtrim(B.descripcion) as BancoDesc, rtrim(C.nombre) as 'Comercializadora'" _
       & " from SEGUROS_Vendedores V inner join Tes_Bancos B on V.cod_Banco = B.id_Banco" _
       & " inner join Seguros_Comercializadoras C on V.cod_Comercializadora = C.cod_Comercializadora" _
       & " where V.cod_vendedor = " & pCodigo
Call OpenRecordSet(rs, strSQL)

If Not rs.BOF And Not rs.EOF Then
  Call sbToolBar(tlb, "activo")
  vEdita = True

  vCodigo = rs!cod_vendedor
  txtCodigo.Text = rs!cod_vendedor
  
  txtIdentificacion.Text = rs!Cedula & ""
  txtNombre = rs!Nombre & ""
  
  chkActivo.Value = rs!Activo
  chkComision.Value = rs!Comision_Aplica
  chkComisonInformativa.Value = rs!comision_Informativa
    
  
  If rs!Tipo_Emision = "CK" Then
     cboTipo.Text = "CK - Cheque"
  Else
     cboTipo.Text = "TE - Transferencia"
  End If
  
  Call sbCboAsignaDato(cboBancos, rs!BancoDesc, True, rs!Cod_Banco)
  Call sbCboAsignaDato(cboComercializadora, rs!Comercializadora, True, rs!cod_Comercializadora)
  
  Call sbCuentas_Load
  
  strSQL = "exec spSeguros_ComisionesTabla_Consulta 'V', '" & vCodigo & "'"
  Call sbCargaGrid(vGrid, 7, strSQL)
  If vGrid.MaxRows > 0 Then
      vGrid.MaxRows = vGrid.MaxRows - 1
  End If
Else
  MsgBox "No se encontró registro verifique...", vbInformation
End If

rs.Close

Call RefrescaTags(Me)

Me.MousePointer = vbDefault

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical


End Sub


Private Function fxValida() As Boolean
Dim strSQL As String, rs As New ADODB.Recordset
Dim vMensaje As String

vMensaje = ""
fxValida = True

If txtNombre = "" Then vMensaje = vMensaje & vbCrLf & " - Nombre no es válido ..."

strSQL = "select count(*) as 'Existe' from SEGUROS_Vendedores" _
        & " where cedula = '" & txtIdentificacion.Text & "' and cod_vendedor <> " & vCodigo
Call OpenRecordSet(rs, strSQL)
If rs!Existe > 0 Then
    vMensaje = vMensaje & vbCrLf & " - El número de identificacion ya esta siendo utilizado por otro Abogado (Revise!) ..."
End If
rs.Close
 

If Len(vMensaje) > 0 Then
  fxValida = False
  MsgBox vMensaje, vbCritical
End If

End Function

Private Sub sbGuardar()
Dim strSQL As String, rs As New ADODB.Recordset
Dim vBufete As String

On Error GoTo vError

If vEdita Then
  strSQL = "update SEGUROS_Vendedores set nombre = '" & Trim(txtNombre.Text) & "',cedula = '" & txtIdentificacion.Text & "',Activo = " & chkActivo.Value _
         & ", cod_Banco = " & cboBancos.ItemData(cboBancos.ListIndex) & ", Tipo_Emision = '" & SIFGlobal.fxCodText(cboTipo.Text) & "', Cuenta_Bancaria = ''" _
         & ", cod_comercializadora = '" & cboComercializadora.ItemData(cboComercializadora.ListIndex) & "'" _
         & ", comision_aplica = " & chkComision.Value & ", comision_Informativa = " & chkComisonInformativa.Value _
         & " where cod_vendedor = " & vCodigo
  Call ConectionExecute(strSQL)
  
  Call Bitacora("Modifica", "SEGUROS Vendedores:  " & vCodigo)

Else
   'Extraer el Ultimo
   strSQL = "select isnull(max(cod_vendedor),0) as Ultimo from SEGUROS_Vendedores"
   Call OpenRecordSet(rs, strSQL)
     txtCodigo.Text = rs!ultimo + 1
   rs.Close
   vCodigo = txtCodigo.Text
   
   strSQL = "Insert into SEGUROS_Vendedores(cod_vendedor,cedula, nombre,cod_Banco,tipo_Emision,Cuenta_Bancaria,comision_Aplica" _
          & ",comision_Informativa,Activo,registro_fecha,registro_usuario,cod_Comercializadora)" _
          & " values(" & vCodigo & ",'" & txtIdentificacion.Text & "','" & txtNombre.Text & "'," & cboBancos.ItemData(cboBancos.ListIndex) & ",'" _
          & SIFGlobal.fxCodText(cboTipo.Text) & "',''," & chkComision.Value & "," & chkComisonInformativa.Value & "," _
          & chkActivo.Value & ",dbo.MyGetdate(),'" & glogon.Usuario & "','" & cboComercializadora.ItemData(cboComercializadora.ListIndex) & "')"
   Call ConectionExecute(strSQL)
   

   Call Bitacora("Registra", "SEGUROS Vendedores:  " & vCodigo)

End If

MsgBox "Información guardada satisfactoriamente...", vbInformation

Call sbConsulta(txtCodigo)

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical


End Sub

Private Sub sbBorrar()
Dim i As Integer, strSQL As String

On Error GoTo vError

i = MsgBox("Esta Seguro que desea borrar este registro", vbYesNo)

If i = vbYes Then
  strSQL = "delete SEGUROS_Vendedores where cod_vendedor = " & vCodigo
  Call ConectionExecute(strSQL)

  Call Bitacora("Elimina", "SEGUROS Vendedores:  " & vCodigo)
  Call sbLimpiaPantalla
  Call sbToolBar(tlb, "nuevo")
  Call RefrescaTags(Me)

End If

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub txtCodigo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtNombre.SetFocus

If KeyCode = vbKeyF4 Then
  gBusquedas.Convertir = "N"
  gBusquedas.Columna = "cod_vendedor"
  gBusquedas.Orden = "cod_vendedor"
  gBusquedas.Consulta = "select cod_vendedor,nombre from SEGUROS_Vendedores"
  gBusquedas.Filtro = ""
  frmBusquedas.Show vbModal
  txtCodigo = gBusquedas.Resultado
  If txtCodigo <> "" Then Call sbConsulta(gBusquedas.Resultado)
End If

End Sub


Private Sub txtCodigo_LostFocus()

If IsNumeric(txtCodigo.Text) Then
  Call sbConsulta(txtCodigo.Text)
End If

End Sub

Private Sub txtIdentificacion_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then cboBancos.SetFocus
End Sub

Private Sub txtNombre_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtIdentificacion.SetFocus

If KeyCode = vbKeyF4 Then
  gBusquedas.Convertir = "N"
  gBusquedas.Columna = "nombre"
  gBusquedas.Orden = "nombre"
  gBusquedas.Consulta = "select cod_vendedor,nombre from SEGUROS_Vendedores"
  gBusquedas.Filtro = ""
  frmBusquedas.Show vbModal
  txtCodigo = gBusquedas.Resultado
  If txtCodigo <> "" Then Call sbConsulta(gBusquedas.Resultado)
End If

End Sub



Private Function fxGuardar() As Long
Dim strSQL As String
'Guarda la información de la linea
'si es Insert devuelve el codigo, sino devuelve 0

Dim pAseguradora As String, pProducto As String
Dim pCtaInicio As Integer, pCtaCorte As Integer, pComisionVenta As Currency, pComisionCta As Currency


On Error GoTo vError

fxGuardar = 0

vGrid.Row = vGrid.ActiveRow
vGrid.Col = 1
pAseguradora = vGrid.Text
vGrid.Col = 2
pProducto = vGrid.Text
vGrid.Col = 4
pComisionVenta = CCur(vGrid.Text)
vGrid.Col = 5
pComisionCta = CCur(vGrid.Text)
vGrid.Col = 6
pCtaInicio = CInt(vGrid.Text)
vGrid.Col = 7
pCtaCorte = CInt(vGrid.Text)
 

 strSQL = "exec spSeguros_ComisionesTabla_Actualiza 'V','" & vCodigo & "','" & pAseguradora & "','" & pProducto & "'," & pComisionVenta _
        & "," & pComisionCta & "," & pCtaInicio & "," & pCtaCorte & ",'" & glogon.Usuario & "'"
 Call ConectionExecute(strSQL)

 Call Bitacora("Modifica", "Comisión (V:" & vCodigo & "  A:" & pAseguradora & "  P:" & pProducto & ")")

fxGuardar = 1

Exit Function

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Function



Private Sub vGrid_KeyDown(KeyCode As Integer, Shift As Integer)
Dim i As Integer, strSQL As String

On Error GoTo vError

If vGrid.ActiveCol = vGrid.MaxCols And (KeyCode = vbKeyReturn Or KeyCode = vbKeyTab) Then
  i = fxGuardar
  If i = 0 Then Exit Sub
  vGrid.Row = vGrid.ActiveRow
  If vGrid.MaxRows <= vGrid.ActiveRow Then
    vGrid.MaxRows = vGrid.MaxRows + 1
    vGrid.Row = vGrid.MaxRows
  End If
End If

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub




