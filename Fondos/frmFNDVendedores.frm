VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#20.3#0"; "Codejock.Controls.v20.3.0.ocx"
Begin VB.Form frmFNDVendedores 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Vendedores de Planes"
   ClientHeight    =   6045
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9315
   Icon            =   "frmFNDVendedores.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6045
   ScaleWidth      =   9315
   Begin XtremeSuiteControls.TabControl tcMain 
      Height          =   4932
      Left            =   0
      TabIndex        =   0
      Top             =   960
      Width           =   9132
      _Version        =   1310723
      _ExtentX        =   16108
      _ExtentY        =   8700
      _StockProps     =   68
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   4
      Color           =   32
      ItemCount       =   1
      Item(0).Caption =   "Datos Generales"
      Item(0).ControlCount=   17
      Item(0).Control(0)=   "Label12(5)"
      Item(0).Control(1)=   "Label12(4)"
      Item(0).Control(2)=   "Label12(3)"
      Item(0).Control(3)=   "lswCuentas"
      Item(0).Control(4)=   "chkComision"
      Item(0).Control(5)=   "btnCuentas"
      Item(0).Control(6)=   "cboTipoPago"
      Item(0).Control(7)=   "Label3(13)"
      Item(0).Control(8)=   "Label3(11)"
      Item(0).Control(9)=   "txtPorcentaje"
      Item(0).Control(10)=   "txtCedula"
      Item(0).Control(11)=   "txtNombre"
      Item(0).Control(12)=   "Label3(2)"
      Item(0).Control(13)=   "Label3(0)"
      Item(0).Control(14)=   "cboBancos"
      Item(0).Control(15)=   "txtMinimoContratos"
      Item(0).Control(16)=   "Label1(1)"
      Begin XtremeSuiteControls.ListView lswCuentas 
         Height          =   1572
         Left            =   360
         TabIndex        =   1
         Top             =   3276
         Width           =   8532
         _Version        =   1310723
         _ExtentX        =   15049
         _ExtentY        =   2773
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
         Appearance      =   16
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.CheckBox chkComision 
         Height          =   252
         Left            =   1560
         TabIndex        =   2
         Top             =   1080
         Width           =   1932
         _Version        =   1310723
         _ExtentX        =   3408
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Aplica comisión?"
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
      Begin XtremeSuiteControls.PushButton btnCuentas 
         Height          =   315
         Left            =   7200
         TabIndex        =   3
         Top             =   2880
         Width           =   1692
         _Version        =   1310723
         _ExtentX        =   2984
         _ExtentY        =   556
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
         Appearance      =   16
      End
      Begin XtremeSuiteControls.ComboBox cboTipoPago 
         Height          =   312
         Left            =   360
         TabIndex        =   4
         Top             =   2880
         Width           =   1932
         _Version        =   1310723
         _ExtentX        =   3413
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
         UseVisualStyle  =   0   'False
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.FlatEdit txtNombre 
         Height          =   312
         Left            =   3360
         TabIndex        =   5
         Top             =   600
         Width           =   5412
         _Version        =   1310723
         _ExtentX        =   9546
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
      Begin XtremeSuiteControls.ComboBox cboBancos 
         Height          =   312
         Left            =   2280
         TabIndex        =   6
         Top             =   2880
         Width           =   4812
         _Version        =   1310723
         _ExtentX        =   8493
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
         UseVisualStyle  =   0   'False
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.FlatEdit txtCedula 
         Height          =   312
         Left            =   1560
         TabIndex        =   7
         Top             =   600
         Width           =   1812
         _Version        =   1310723
         _ExtentX        =   3196
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
         Appearance      =   2
      End
      Begin XtremeSuiteControls.FlatEdit txtPorcentaje 
         Height          =   312
         Left            =   1560
         TabIndex        =   18
         Top             =   1560
         Width           =   972
         _Version        =   1310723
         _ExtentX        =   1714
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
         Appearance      =   2
      End
      Begin XtremeSuiteControls.FlatEdit txtMinimoContratos 
         Height          =   312
         Left            =   1560
         TabIndex        =   19
         Top             =   1920
         Width           =   972
         _Version        =   1310723
         _ExtentX        =   1714
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
         Appearance      =   2
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Contratos"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   312
         Index           =   5
         Left            =   480
         TabIndex        =   23
         ToolTipText     =   "Cuenta de Inventarios para la Bodega"
         Top             =   1920
         Width           =   852
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Meta: Cantidad de Contratos mínimos para reconocimiento de comisión"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   492
         Index           =   1
         Left            =   2640
         TabIndex        =   22
         Top             =   1920
         Width           =   4212
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "% Comisión"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   312
         Index           =   4
         Left            =   2640
         TabIndex        =   21
         ToolTipText     =   "Cuenta de Inventarios para la Bodega"
         Top             =   1560
         Width           =   1572
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Porcentaje"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   312
         Index           =   3
         Left            =   480
         TabIndex        =   20
         ToolTipText     =   "Cuenta de Inventarios para la Bodega"
         Top             =   1560
         Width           =   1092
      End
      Begin XtremeSuiteControls.Label Label3 
         Height          =   264
         Index           =   13
         Left            =   360
         TabIndex        =   11
         Top             =   2640
         Width           =   1932
         _Version        =   1310723
         _ExtentX        =   3408
         _ExtentY        =   466
         _StockProps     =   79
         Caption         =   "Emitir"
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
         WordWrap        =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label3 
         Height          =   252
         Index           =   11
         Left            =   2280
         TabIndex        =   10
         Top             =   2640
         Width           =   1932
         _Version        =   1310723
         _ExtentX        =   3408
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Cuenta/Desembolso"
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
         WordWrap        =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label3 
         Height          =   252
         Index           =   2
         Left            =   3360
         TabIndex        =   9
         Top             =   360
         Width           =   1332
         _Version        =   1310723
         _ExtentX        =   2350
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Nombre"
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
         WordWrap        =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label3 
         Height          =   252
         Index           =   0
         Left            =   0
         TabIndex        =   8
         Top             =   600
         Width           =   1332
         _Version        =   1310723
         _ExtentX        =   2350
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Identificación"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Transparent     =   -1  'True
         WordWrap        =   -1  'True
      End
   End
   Begin MSComCtl2.FlatScrollBar FlatScrollBar 
      Height          =   252
      Left            =   3480
      TabIndex        =   12
      Top             =   360
      Width           =   492
      _ExtentX        =   873
      _ExtentY        =   450
      _Version        =   393216
      Arrows          =   65536
      Orientation     =   1638401
   End
   Begin XtremeSuiteControls.FlatEdit txtCodigo 
      Height          =   312
      Left            =   1560
      TabIndex        =   13
      Top             =   360
      Width           =   1812
      _Version        =   1310723
      _ExtentX        =   3196
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
   Begin XtremeSuiteControls.ComboBox cboEstado 
      Height          =   312
      Left            =   6240
      TabIndex        =   14
      Top             =   360
      Width           =   1932
      _Version        =   1310723
      _ExtentX        =   3413
      _ExtentY        =   582
      _StockProps     =   77
      ForeColor       =   1973790
      BackColor       =   16185078
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16185078
      Style           =   2
      Appearance      =   16
      UseVisualStyle  =   0   'False
   End
   Begin MSComctlLib.Toolbar tlb 
      Height          =   990
      Left            =   1560
      TabIndex        =   15
      Top             =   0
      Width           =   2430
      _ExtentX        =   4286
      _ExtentY        =   1746
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
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Vendedor Id:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   312
      Index           =   0
      Left            =   240
      TabIndex        =   17
      Top             =   360
      Width           =   1092
   End
   Begin XtremeSuiteControls.Label Label3 
      Height          =   252
      Index           =   3
      Left            =   4680
      TabIndex        =   16
      Top             =   360
      Width           =   1332
      _Version        =   1310723
      _ExtentX        =   2350
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Estado"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
      Transparent     =   -1  'True
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmFNDVendedores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vEdita As Boolean, vCodigo As String, vScroll As Boolean

Private Sub sbCuentas_Load()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem

On Error GoTo vError

lswCuentas.ListItems.Clear
If vCodigo <> "" Then
    strSQL = "select rtrim(B.Descripcion) as 'Banco'" _
           & ",case when C.tipo = 'A' then 'Ahorros' else 'Corriente' end as 'TipoDesc'" _
           & ",C.cod_Divisa,C.CUENTA_INTERNA, C.CUENTA_INTERBANCA, C.ACTIVA, C.DESTINO, C.REGISTRO_FECHA , C.REGISTRO_USUARIO" _
           & " from SYS_CUENTAS_BANCARIAS C inner join TES_BANCOS_GRUPOS B on C.cod_banco = B.cod_grupo" _
           & " where C.Identificacion = '" & Trim(txtCedula.Text) & "'"
    
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
End If


Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub btnCuentas_Click()

GLOBALES.gTag = Trim(txtCedula.Text)
GLOBALES.gTag2 = "FND"

frmCC_Cuentas_Bancarias.Show vbModal

Call sbCuentas_Load
End Sub

Private Sub cboBancos_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab And btnCuentas.Enabled Then btnCuentas.SetFocus
End Sub



Private Sub FlatScrollBar_Change()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

If vScroll Then
    strSQL = "select Top 1 cod_Vendedor from fnd_vendedores"
    
    If FlatScrollBar.Value = 1 Then
       strSQL = strSQL & " where cod_Vendedor > '" & txtCodigo.Text & "' order by cod_Vendedor asc"
    Else
       strSQL = strSQL & " where cod_Vendedor < '" & txtCodigo.Text & "' order by cod_Vendedor desc"
    End If
    
    Call OpenRecordSet(rs, strSQL)
    If Not rs.EOF And Not rs.BOF Then
      Call sbConsulta(rs!cod_Vendedor)
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
 vModulo = 18
End Sub

Private Sub Form_Load()
Dim strSQL As String

On Error GoTo vError
 
 vModulo = 18
 
 vScroll = False
    FlatScrollBar.Value = 0
 vScroll = True
 
cboTipoPago.Clear
cboTipoPago.AddItem fxTipoDocumento("CK")
cboTipoPago.AddItem fxTipoDocumento("TE")

  
lswCuentas.ColumnHeaders.Add 1, , "Cuenta", 2500
lswCuentas.ColumnHeaders.Add 2, , "Banco", 3500
lswCuentas.ColumnHeaders.Add 3, , "Tipo", 1100, vbCenter
lswCuentas.ColumnHeaders.Add 4, , "Divisa", 1100, vbCenter
lswCuentas.ColumnHeaders.Add 5, , "Interbanca", 1100, vbCenter
lswCuentas.ColumnHeaders.Add 6, , "Destino", 1100, vbCenter
lswCuentas.ColumnHeaders.Add 7, , "Activa", 1100, vbCenter
lswCuentas.ColumnHeaders.Add 8, , "Fecha", 2500
lswCuentas.ColumnHeaders.Add 9, , "Usuario", 2500

strSQL = "exec spCrd_SGT_Bancos '" & glogon.Usuario & "'"
Call sbCbo_Llena_New(cboBancos, strSQL, False, True)
 
cboEstado.Clear
cboEstado.AddItem "Activo"
cboEstado.AddItem "Inactivo"

 vEdita = True
 Call sbToolBarIconos(tlb)
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


txtNombre.Text = ""
txtCedula.Text = ""
chkComision.Value = vbChecked


cboEstado.Text = "Activo"
cboTipoPago.Text = fxTipoDocumento("TE")

chkComision.Value = vbUnchecked


txtPorcentaje.Text = "0"
txtMinimoContratos.Text = "30"

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
        gBusquedas.Convertir = "N"
        gBusquedas.Col1Name = "Vendedor Id"
        gBusquedas.Col2Name = "Persona Id"
        gBusquedas.Col3Name = "Nombre"
        gBusquedas.Columna = "Cedula"
        gBusquedas.Orden = "Cedula"
        gBusquedas.Consulta = "select cod_Vendedor,Cedula,Nombre from fnd_vendedores"
        gBusquedas.Filtro = ""
        frmBusquedas.Show vbModal
        txtCodigo = gBusquedas.Resultado
        If txtCodigo <> "" Then Call sbConsulta(gBusquedas.Resultado)
    
    Case "REPORTES"
    
    Case "AYUDA"
       ' frmContenedor.CD.HelpContext = Me.HelpContextID
       ' frmContenedor.CD.ShowHelp
   
End Select

End Sub

Private Sub sbConsulta(pCodigo As String)
Dim rs As New ADODB.Recordset, strSQL As String

On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = "select V.*, B.Descripcion as 'BancoDesc' " _
       & " from fnd_vendedores V inner join tes_Bancos B on V.Cod_Banco = B.id_Banco" _
       & " where V.COD_VENDEDOR = '" & pCodigo & "'"
Call OpenRecordSet(rs, strSQL)

If Not rs.BOF And Not rs.EOF Then
  Call sbToolBar(tlb, "activo")
  vEdita = True
  
  vCodigo = rs!cod_Vendedor
  txtCodigo = rs!cod_Vendedor
  
  txtCedula.Text = rs!Cedula & ""
  txtNombre.Text = rs!Nombre & ""
  
  chkComision.Value = rs!aplica_comision
  
  txtPorcentaje.Text = Format(rs!porc_comision, "Standard")
  txtMinimoContratos.Text = CStr(rs!Minimo)
  

    If rs!Estado = "A" Then
      cboEstado.Text = "Activo"
    Else
      cboEstado.Text = "Inactivo"
    End If
    
    cboTipoPago.Text = fxTipoDocumento(rs!TIPO_PAGO)
    
    Call sbCboAsignaDato(cboBancos, Trim(rs!BancoDesc), True, rs!cod_banco)

    Call sbCuentas_Load
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

On Error GoTo vError

vMensaje = ""
fxValida = True

If txtNombre = "" Then vMensaje = vMensaje & vbCrLf & " - Nombre no es válido ..."
If txtCedula = "" Then vMensaje = vMensaje & vbCrLf & " - Número de Identificación no es válida ..."
If CCur(txtPorcentaje) > 100 Or CCur(txtPorcentaje) < 0 Then vMensaje = vMensaje & vbCrLf & " - Porcentaje de Comisión no es válida ..."
If CInt(txtMinimoContratos) < 0 Then vMensaje = vMensaje & vbCrLf & " - Tiempo de Resolución no es válida ..."

If Len(vMensaje) > 0 Then
  fxValida = False
  MsgBox vMensaje, vbCritical
End If

Exit Function

vError:
  fxValida = False
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
 
End Function

Private Sub sbGuardar()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

If vEdita Then
  strSQL = "update fnd_vendedores set nombre = '" & Trim(txtNombre) & "'" _
         & ",cedula = '" & txtCedula & "',estado = '" & Mid(cboEstado.Text, 1, 1) _
         & "',aplica_comision = " & chkComision.Value & ",cod_banco = " & cboBancos.ItemData(cboBancos.ListIndex) _
         & ",Tipo_Pago = '" & fxTipoDocumento(cboTipoPago.Text) _
         & "',porc_comision = " & CCur(txtPorcentaje) _
         & ",Minimo = " & CLng(txtMinimoContratos.Text) _
         & " where cod_Vendedor = '" & vCodigo & "'"
  Call ConectionExecute(strSQL)
  
  Call Bitacora("Modifica", "Vendedor de Planes de Ahorros: " & vCodigo)

Else
  vCodigo = txtCodigo
   
   strSQL = "insert into fnd_vendedores(cod_vendedor,nombre,cedula,estado,aplica_comision" _
          & ",cod_banco,Tipo_Pago,Minimo,porc_comision,Tipo_Id)" _
          & " values('" & vCodigo & "','" & txtNombre & "','" & txtCedula & "','" _
          & Mid(cboEstado.Text, 1, 1) & "'," & chkComision.Value _
          & "," & cboBancos.ItemData(cboBancos.ListIndex) & ",'" & fxTipoDocumento(cboTipoPago.Text) _
          & "'," & CLng(txtMinimoContratos.Text) & "," & CCur(txtPorcentaje.Text) & ",1)"
   
   Call ConectionExecute(strSQL)
    
   Call Bitacora("Registra", "Vendedor de Planes de Ahorros: " & vCodigo)
 
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
  strSQL = "delete fnd_vendedores where cod_vendedor = '" & vCodigo & "'"
  Call ConectionExecute(strSQL)
  
  Call Bitacora("Elimina", "Vendedor de Planes de Ahorros: " & vCodigo)
  Call sbLimpiaPantalla
  Call sbToolBar(tlb, "nuevo")
  Call RefrescaTags(Me)
End If

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub txtCodigo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then
  If txtCodigo <> "" And vEdita Then Call sbConsulta(txtCodigo)
  txtCedula.SetFocus
End If

If KeyCode = vbKeyF4 Then
  gBusquedas.Convertir = "N"
  gBusquedas.Col1Name = "Vendedor Id"
  gBusquedas.Col2Name = "Persona Id"
  gBusquedas.Col3Name = "Nombre"
  gBusquedas.Columna = "Cedula"
  gBusquedas.Orden = "Cedula"
  gBusquedas.Consulta = "select cod_Vendedor,Cedula,Nombre from fnd_vendedores"
  gBusquedas.Filtro = ""
  frmBusquedas.Show vbModal
  txtCodigo = gBusquedas.Resultado
  If txtCodigo <> "" Then Call sbConsulta(gBusquedas.Resultado)
End If
End Sub

Private Sub txtCedula_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtNombre.SetFocus
If KeyCode = vbKeyF4 Then
  gBusquedas.Convertir = "N"
  gBusquedas.Col1Name = "Vendedor Id"
  gBusquedas.Col2Name = "Persona Id"
  gBusquedas.Col3Name = "Nombre"
  gBusquedas.Columna = "Cedula"
  gBusquedas.Orden = "Cedula"
  gBusquedas.Consulta = "select cod_Vendedor,Cedula,Nombre from fnd_vendedores"
  gBusquedas.Filtro = ""
  frmBusquedas.Show vbModal
  txtCodigo = gBusquedas.Resultado
  If txtCodigo <> "" Then Call sbConsulta(gBusquedas.Resultado)
End If
End Sub


Private Sub txtPorcentaje_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtMinimoContratos.SetFocus
End Sub

Private Sub txtMinimoContratos_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then cboTipoPago.SetFocus
End Sub

Private Sub txtNombre_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then chkComision.SetFocus
If KeyCode = vbKeyF4 Then
  gBusquedas.Convertir = "N"
  gBusquedas.Col1Name = "Vendedor Id"
  gBusquedas.Col2Name = "Persona Id"
  gBusquedas.Col3Name = "Nombre"
  gBusquedas.Columna = "nombre"
  gBusquedas.Orden = "nombre"
  gBusquedas.Consulta = "select cod_Vendedor,Cedula,Nombre from fnd_vendedores"
  gBusquedas.Filtro = ""
  frmBusquedas.Show vbModal
  txtCodigo = gBusquedas.Resultado
  If txtCodigo <> "" Then Call sbConsulta(gBusquedas.Resultado)
End If
End Sub

