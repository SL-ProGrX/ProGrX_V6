VERSION 5.00
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "ComCt332.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmPosPedidos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pedidos"
   ClientHeight    =   6315
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10155
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6315
   ScaleWidth      =   10155
   Begin VB.CheckBox chkPlantilla 
      Caption         =   "Plantilla ?"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2760
      TabIndex        =   26
      Top             =   480
      Width           =   1215
   End
   Begin VB.TextBox txtFecha 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   7680
      Locked          =   -1  'True
      TabIndex        =   15
      Top             =   480
      Width           =   2415
   End
   Begin VB.TextBox txtCodigo 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1080
      TabIndex        =   14
      Top             =   480
      Width           =   1575
   End
   Begin VB.TextBox txtSubTotal 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   8040
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   13
      Top             =   4875
      Width           =   1812
   End
   Begin VB.TextBox txtDescuento 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   315
      Left            =   8040
      TabIndex        =   12
      Top             =   5235
      Width           =   1812
   End
   Begin VB.TextBox txtImpuestos 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   8040
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   11
      Top             =   5565
      Width           =   1812
   End
   Begin VB.TextBox txtTotal 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   8040
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   10
      Top             =   5925
      Width           =   1812
   End
   Begin VB.ComboBox cboPrecio 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   300
      Left            =   1080
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   960
      Width           =   3735
   End
   Begin VB.ComboBox cboAgente 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   300
      Left            =   6000
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   960
      Width           =   4095
   End
   Begin VB.TextBox txtCedula 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1080
      TabIndex        =   7
      Top             =   1800
      Width           =   1935
   End
   Begin VB.TextBox txtNombre 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3000
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   1800
      Width           =   6855
   End
   Begin VB.ComboBox cboFormaPago 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   300
      ItemData        =   "frmPosPedidos.frx":0000
      Left            =   1080
      List            =   "frmPosPedidos.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   1320
      Width           =   3735
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   8640
      Top             =   960
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
            Picture         =   "frmPosPedidos.frx":0004
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPosPedidos.frx":031E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPosPedidos.frx":0BF8
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPosPedidos.frx":0F12
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComCtl2.DTPicker dtpFecha 
      Height          =   312
      Left            =   5040
      TabIndex        =   0
      Top             =   480
      Width           =   1572
      _ExtentX        =   2778
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   294649859
      CurrentDate     =   37788
   End
   Begin ComCtl3.CoolBar CoolBar 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   10152
      _ExtentX        =   17912
      _ExtentY        =   635
      _CBWidth        =   10155
      _CBHeight       =   360
      _Version        =   "6.7.9839"
      Child1          =   "tlb"
      MinHeight1      =   270
      Width1          =   4485
      NewRow1         =   0   'False
      Child2          =   "tlbAux"
      MinHeight2      =   315
      Width2          =   1995
      NewRow2         =   0   'False
      Child3          =   "tlbAux02"
      MinHeight3      =   315
      Width3          =   2730
      NewRow3         =   0   'False
      Begin MSComctlLib.Toolbar tlb 
         Height          =   264
         Left            =   132
         TabIndex        =   4
         Top             =   48
         Width           =   4332
         _ExtentX        =   7646
         _ExtentY        =   476
         ButtonWidth     =   487
         ButtonHeight    =   466
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
               Style           =   5
               BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                  NumButtonMenus  =   2
                  BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "repBoleta"
                     Text            =   "Boleta (Factura)"
                  EndProperty
                  BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "repListadoGeneral"
                     Text            =   "Listado General"
                  EndProperty
               EndProperty
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "ayuda"
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.Toolbar tlbAux 
         Height          =   312
         Left            =   4644
         TabIndex        =   3
         Top             =   24
         Width           =   1836
         _ExtentX        =   3228
         _ExtentY        =   556
         ButtonWidth     =   2350
         ButtonHeight    =   550
         Style           =   1
         TextAlignment   =   1
         ImageList       =   "ImgAux01"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   2
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Ficha Cliente"
               Key             =   "Cliente"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
               Caption         =   "Ficha Crédito"
               Key             =   "Credito"
               ImageIndex      =   3
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.Toolbar tlbAux02 
         Height          =   312
         Left            =   6660
         TabIndex        =   2
         Top             =   24
         Width           =   3420
         _ExtentX        =   6033
         _ExtentY        =   556
         ButtonWidth     =   1884
         ButtonHeight    =   550
         Style           =   1
         TextAlignment   =   1
         ImageList       =   "ImageList1"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   1
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Proforma"
               Key             =   "ProForma"
               Object.ToolTipText     =   "Imprimir como Proforma"
               ImageIndex      =   4
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.ImageList ImgAux01 
      Left            =   5760
      Top             =   4680
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPosPedidos.frx":122C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPosPedidos.frx":1B08
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPosPedidos.frx":1E24
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPosPedidos.frx":2148
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPosPedidos.frx":2464
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin FPSpreadADO.fpSpread vGrid 
      Height          =   2292
      Left            =   0
      TabIndex        =   28
      Top             =   2280
      Width           =   10092
      _Version        =   524288
      _ExtentX        =   17801
      _ExtentY        =   4043
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
      MaxCols         =   488
      ScrollBars      =   2
      SpreadDesigner  =   "frmPosPedidos.frx":2780
      VScrollSpecial  =   -1  'True
      VScrollSpecialType=   2
      AppearanceStyle =   1
   End
   Begin VB.Label Label1 
      Caption         =   "Vence"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   1
      Left            =   4320
      TabIndex        =   27
      Top             =   480
      Width           =   732
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha"
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
      Index           =   4
      Left            =   6840
      TabIndex        =   25
      Top             =   480
      Width           =   735
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "# Pedido"
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
      Left            =   240
      TabIndex        =   24
      Top             =   480
      Width           =   732
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Sub Total"
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
      Index           =   6
      Left            =   6720
      TabIndex        =   23
      Top             =   4872
      Width           =   972
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "(-) Descuento"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   252
      Index           =   7
      Left            =   6720
      TabIndex        =   22
      Top             =   5232
      Width           =   1332
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "(+) Impuestos"
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
      Index           =   8
      Left            =   6720
      TabIndex        =   21
      Top             =   5568
      Width           =   1212
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Total"
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
      Index           =   9
      Left            =   6720
      TabIndex        =   20
      Top             =   5928
      Width           =   972
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Precio"
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
      Index           =   3
      Left            =   240
      TabIndex        =   19
      Top             =   960
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Agente"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   10
      Left            =   5040
      TabIndex        =   18
      Top             =   960
      Width           =   735
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Cliente"
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
      Index           =   12
      Left            =   240
      TabIndex        =   17
      Top             =   1800
      Width           =   735
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Pago"
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
      Index           =   14
      Left            =   240
      TabIndex        =   16
      Top             =   1320
      Width           =   735
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00FFFFFF&
      X1              =   10080
      X2              =   0
      Y1              =   1680
      Y2              =   1680
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   10080
      X2              =   0
      Y1              =   840
      Y2              =   840
   End
End
Attribute VB_Name = "frmPosPedidos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vEdita As Boolean, vCodigo As String

Private Sub cboAgente_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then cboPrecio.SetFocus
End Sub

Private Sub cboFormaPago_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtCedula.SetFocus
End Sub

Private Sub cboPrecio_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then cboFormaPago.SetFocus
End Sub

Private Sub Form_Activate()
vModulo = 33
End Sub

Private Sub Form_Load()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

vModulo = 33
vGrid.AppearanceStyle = fxGridStyle
 
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
Dim strSQL As String, rs As New ADODB.Recordset
Dim i As Integer

Me.MousePointer = vbHourglass

On Error GoTo vError

vCodigo = 0
txtCodigo = ""

dtpFecha.Value = Format(fxFechaServidor, "yyyy/mm/dd hh:mm:ss")
txtFecha = Format(dtpFecha.Value, "yyyy/mm/dd hh:mm:ss")

txtNombre = ""
txtCedula = ""

Call sbPosCombosCarga("Agentes", cboAgente, " where estado = 'A'")
Call sbPosCombosCarga("Precios", cboPrecio)
Call sbPosCombosCarga("FormaPago", cboFormaPago)

vGrid.MaxRows = 1
vGrid.MaxCols = 7
For i = 1 To vGrid.MaxCols
  vGrid.col = i
  vGrid.Text = ""
Next

txtSubTotal = 0
txtDescuento = 0
txtImpuestos = 0
txtTotal = 0

txtCodigo.Enabled = True
txtCodigo.SetFocus
Me.MousePointer = vbDefault

Exit Sub

vError:
  Me.MousePointer = vbDefault
'  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub


Private Sub tlb_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim strSQL As String

Select Case UCase(Button.Key)
    Case "INSERTAR", "NUEVO"
      vEdita = False
      Call sbLimpiaPantalla
      Call sbToolBar(tlb, "edicion")
      
      txtCodigo.Enabled = False
      
    Case "MODIFICAR", "EDITAR"
      vEdita = True
      txtCedula.SetFocus
      Call sbToolBar(tlb, "edicion")
    Case "BORRAR"
      Call sbBorrar
    Case "GUARDAR", "SALVAR"
     If fxValida Then Call sbGuardar
    Case "DESHACER"
      Call sbToolBar(tlb, "activo")
      If vCodigo = 0 Then
        Call sbLimpiaPantalla
        Call sbToolBar(tlb, "nuevo")
        vEdita = True
      Else
        Call sbConsulta(vCodigo)
      End If

    Case "CONSULTAR"
'       gBusquedas.Columna = "descripcion"
'       gBusquedas.Orden = "descripcion"
'       gBusquedas.Consulta = "select cod_proveedor,descripcion from cxp_proveedores"
'       frmBusquedas.Show vbModal
'       txtCodigo.SetFocus
'       txtCodigo = IIf((gBusquedas.Resultado = ""), 0, gBusquedas.Resultado)
'       txtNombre.SetFocus

    Case "REPORTES"

    Case "AYUDA"
        frmContenedor.CD.HelpContext = Me.HelpContextID
        frmContenedor.CD.ShowHelp

End Select

End Sub

Private Sub sbConsulta(vCodigo As String)
Dim rs As New ADODB.Recordset, strSQL As String

On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = "select F.cod_pedido,F.fecha,F.vence,F.sub_total,F.descuento,F.imp_ventas,F.total,F.plantilla" _
       & ",F.cedula,C.nombre,(rtrim(F.cod_precio) + ' - ' + P.descripcion) as Precio" _
       & ",(rtrim(F.cod_agente) + ' - ' + A.nombre) as Agente" _
       & ",(rtrim(CONVERT(char, F.cod_forma_pago))+ ' - ' + R.descripcion) as FormaPago" _
       & " from pv_Pedidos F inner join pv_clientes C on F.cedula = C.cedula" _
       & " inner join pv_tipos_precios P on F.cod_precio = P.cod_precio" _
       & " inner join pv_agentes A on F.cod_agente = A.cod_agente" _
       & " inner join pv_formas_pago R on F.cod_forma_pago = R.cod_forma_pago" _
       & " where cod_pedido = '" & vCodigo & "'"
Call OpenRecordSet(rs, strSQL)

If Not rs.BOF And Not rs.EOF Then
  Call sbToolBar(tlb, "activo")
  vEdita = True
  vCodigo = rs!cod_pedido
  txtCodigo = rs!cod_pedido
      
  txtCedula = rs!Cedula
  txtNombre = rs!Nombre
  
  chkPlantilla.Value = rs!plantilla
     
  Call sbPosCombosCarga("Agentes", cboAgente)
  cboAgente.Text = Trim(rs!Agente)
  cboFormaPago.Text = Trim(rs!FormaPago)
  cboPrecio.Text = Trim(rs!Precio)
    
  txtFecha = Format(rs!fecha, "yyyy/mm/dd hh:mm:ss")
  
  dtpFecha.Value = rs!Vence
  
  txtSubTotal = Format(rs!sub_Total, "Standard")
  txtDescuento = Format(rs!descuento, "Standard")
  txtImpuestos = Format(rs!imp_ventas, "Standard")
  txtTotal = Format(rs!Total, "Standard")

  strSQL = "select D.cod_producto,P.descripcion,D.cantidad,D.cod_bodega,D.precio,D.imp_ventas," _
         & "(D.cantidad * D.precio) + (D.cantidad * D.precio * (D.imp_ventas / 100)) as Total" _
         & " from pv_pedidos_detalle D inner join pv_productos P on D.cod_producto = P.cod_producto" _
         & " where D.cod_pedido = '" & rs!cod_pedido & "' order by D.Linea"
  Call sbCargaGrid(vGrid, 7, strSQL)
  
  Call sbCalculaTotales
  
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


'Validar Cliente, que exista
strSQL = "select isnull(count(*),0) as Existe from pv_clientes" _
       & " where cedula = '" & txtCedula & "'"
Call OpenRecordSet(rs, strSQL)
If rs!Existe = 0 Then
    vMensaje = vMensaje & vbCrLf & " - No Existe registro del cliente especificado ..."
End If
rs.Close

'Validar que se haya facturado almenos un articulo
   
Select Case vGrid.MaxRows
  Case 1 'Solo hay una linea, verifica si existe un producto en ella
    vGrid.col = 2
    vGrid.Row = 1
    If Trim(vGrid.Text) = "" Then
       vMensaje = vMensaje & vbCrLf & " - No hay productos / articulos o servicios en el detalle ..."
    End If
  
  Case 0 'No hay linea de detalle
       vMensaje = vMensaje & vbCrLf & " - No hay productos / articulos o servicios en el detalle..."
End Select

If Len(vMensaje) > 0 Then
  fxValida = False
  MsgBox vMensaje, vbCritical
End If


End Function


Private Sub sbGuardar()
Dim strSQL As String, i As Integer, curCantidad As Currency
Dim vCodPro As String, vCodBodega As String, vFecha As Date
Dim curPrecio As Currency, curImpVentas As Currency, curImpConsumo As Currency
Dim vImprime As Boolean

On Error GoTo vError


'Solo se puede Insertar y no Editar
'01 - Guardar el registro en las Entradas y Afectar Inventarios

If vEdita Then
   MsgBox "No se puede editar un pedido Guardado...", vbInformation
   Exit Sub
End If

vCodigo = fxSIFCConsecutivos("pedidos")
txtCodigo = vCodigo

vFecha = fxFechaServidor

glogon.Conection.BeginTrans

strSQL = "insert pv_pedidos(cod_pedido,cedula,cod_agente,cod_precio,cod_forma_pago,fecha,vence" _
       & ",sub_total,descuento,imp_ventas,imp_consumo,total,plantilla)" _
       & " values('" & vCodigo & "','" & txtCedula & "','" & fxCodigoCbo(cboAgente) _
       & "','" & fxCodigoCbo(cboPrecio) & "'," & fxCodigoCbo(cboFormaPago) _
       & ",'" & Format(vFecha, "yyyy/mm/dd hh:mm:ss") & "','" & Format(dtpFecha.Value, "yyyy/mm/dd") _
       & "'," & CCur(txtSubTotal) & "," & CCur(txtDescuento) & "," & CCur(txtImpuestos) _
       & ",0," & CCur(txtTotal) & "," & chkPlantilla.Value & ")"
Call ConectionExecute(strSQL)

Call Bitacora("Registra", "Pedido N.: " & vCodigo)

txtCodigo.Enabled = True

'Guardar Detalle de la Factura y Registra Inventario
strSQL = "delete pv_pedidos_detalle" _
         & " where cod_pedido = '" & vCodigo & "'"
Call ConectionExecute(strSQL)

For i = 1 To vGrid.MaxRows
  vGrid.Row = i
  
  vGrid.col = 3
  curCantidad = CCur(IIf((vGrid.Text = ""), 0, vGrid.Text))
  
  vGrid.col = 1
  
  If vGrid.Text <> "" And curCantidad > 0 Then
    
    vGrid.col = 1
    vCodPro = Trim(vGrid.Text)
    strSQL = "insert pv_pedidos_detalle(linea,cod_pedido,cod_producto,cantidad,cod_bodega" _
           & ",precio,imp_ventas,imp_consumo) values(" & i & ",'" & vCodigo & "','" _
           & vGrid.Text & "'," & curCantidad & ",'"
    vGrid.col = 4
    vCodBodega = Trim(vGrid.Text)
    strSQL = strSQL & vGrid.Text & "',"
    vGrid.col = 5
    curPrecio = CCur(vGrid.Text)
    strSQL = strSQL & CCur(vGrid.Text) & ","
    vGrid.col = 6
    curImpVentas = CCur(vGrid.Text)
    curImpConsumo = 0
    strSQL = strSQL & CCur(vGrid.Text) & ",0)"
    Call ConectionExecute(strSQL)
  End If
Next i

'Guarda Transaccion
glogon.Conection.CommitTrans

Call sbToolBar(tlb, "activo")
Call RefrescaTags(Me)

MsgBox "Información guardada satisfactoriamente...", vbInformation

Exit Sub

vError:
 glogon.Conection.RollbackTrans
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub sbBorrar()
Dim i As Integer, strSQL As String

On Error GoTo vError

i = MsgBox("Esta Seguro que desea borrar este registro", vbYesNo)

If i = vbYes Then
   'no se pueden Ejecutar Borrados en Ordenes
'  strSQL = "delete cxp_proveedores where cod_proveedor = " & vCodigo
'  Call ConectionExecute(strSQL)

'  Call Bitacora("Elimina", "ER ESPECIAL : " & vCodigo & " EMP: " & vParametros.CodigoEmpresa)
  Call sbLimpiaPantalla
  Call sbToolBar(tlb, "nuevo")
  Call RefrescaTags(Me)
End If

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub tlb_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
Select Case ButtonMenu.Key
  Case "repBoleta"
     'Preguntar por el Formato si es Boucher o Clasico
  Case "repListado"
End Select
End Sub

Private Sub tlbAux_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim strSQL As String, rs As New ADODB.Recordset

Select Case Button.Key
  Case "Cliente"
    Call MuestraForms(frmPosFichaCliente)
End Select
End Sub

Private Sub tlbAux02_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim vSQL As String

vSQL = "{PV_PEDIDOS.cod_pedido} = '" & txtCodigo & "'"

Select Case Button.Key
  Case "ProForma"
   Call sbPosReportes("PROFORMAS", "FACTURA PROFORMA", "NUMERO : " & Format(txtCodigo, "000000"), vSQL)
End Select
End Sub

Private Sub txtCedula_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtNombre.SetFocus

If KeyCode = vbKeyF4 Then
  gBusquedas.Convertir = "N"
  gBusquedas.Columna = "cedula"
  gBusquedas.Orden = "cedula"
  gBusquedas.Consulta = "select cedula,nombre from pv_clientes"
  gBusquedas.Filtro = ""
  frmBusquedas.Show vbModal
  txtCedula = gBusquedas.Resultado
  txtNombre = gBusquedas.Resultado2
End If

End Sub

Private Sub txtCedula_LostFocus()
'Verifica el Enlace con SIFA
Call sbXFichaCliente(txtCedula)
txtNombre = fxSIFCCodigos("D", txtCedula, "clientes")
End Sub


Private Sub txtCodigo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then dtpFecha.SetFocus

If KeyCode = vbKeyF4 Then
  gBusquedas.Convertir = "N"
  gBusquedas.Columna = "P.cod_pedido"
  gBusquedas.Orden = "P.cod_pedido"
  gBusquedas.Consulta = "select P.cod_pedido,P.cedula,C.nombre from pv_facturacion P " _
            & " inner join pv_clientes C on P.cedula = C.cedula"
  gBusquedas.Filtro = ""
  frmBusquedas.Show vbModal
  txtCodigo = gBusquedas.Resultado
  If txtCodigo <> "" Then Call sbConsulta(CLng(gBusquedas.Resultado))
End If

End Sub

Private Sub txtCodigo_LostFocus()
If txtCodigo <> "" And vEdita Then Call sbConsulta(txtCodigo)
End Sub

Private Sub txtDescuento_GotFocus()
On Error GoTo vError
txtDescuento = CCur(txtDescuento)
Exit Sub
vError:
  MsgBox "Información del Descuento no es válida...", vbCritical
End Sub

Private Sub txtDescuento_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtTotal.SetFocus
End Sub

Private Sub txtDescuento_LostFocus()
On Error GoTo vError

txtDescuento = Format(CCur(txtDescuento), "Standard")
txtTotal = Format(CCur(txtSubTotal) + CCur(txtImpuestos) - CCur(txtDescuento), "Standard")

Exit Sub
vError:
  MsgBox "Información del Descuento no es válida...", vbCritical
End Sub


Private Sub sbCalculaTotales()
Dim curSubTotal As Currency, curIV As Currency
Dim curTmpPrecio As Currency, curTmpIV As Currency, curTmpCant As Currency
Dim i As Integer, lng As Long

'**********************************************    OJO
'Revisar esta formula por la situacion del descuento, si es antes o despues del
'impuesto de ventas, por ahora está despues del impuesto

curSubTotal = 0
curIV = 0

For lng = 1 To vGrid.MaxRows
 vGrid.Row = lng
 vGrid.col = 3
 If vGrid.Text <> "" Then
    curTmpCant = CCur(vGrid.Text)
    vGrid.col = 5
    curTmpPrecio = CCur(vGrid.Text)
    vGrid.col = 6
    curTmpIV = CCur(vGrid.Text)
    
    vGrid.col = 7
    vGrid.Text = (curTmpCant * curTmpPrecio) + ((curTmpCant * curTmpPrecio) * curTmpIV / 100)
     
     
    curSubTotal = curSubTotal + (curTmpCant * curTmpPrecio)
    curIV = curIV + ((curTmpCant * curTmpPrecio) * (curTmpIV / 100))
 End If
Next lng

txtSubTotal = Format(curSubTotal, "Standard")
txtImpuestos = Format(curIV, "Standard")
txtTotal = Format(curSubTotal + curIV - CCur(txtDescuento), "Standard")

End Sub

Private Sub sbConsultaArticulo(fila As Long, Columna As Integer, vCriterio As String)
Dim strSQL As String, rs As New ADODB.Recordset, vPaso As Boolean
Dim vBodega As String

'Busquedas
'1. Por Codigo del Articulo
'2. Por Codigo de Barras
'3. Por Codigo del Fabricante
vPaso = False

vGrid.Row = fila
vGrid.col = 5
vGrid.Lock = False

vGrid.Row = fila
vGrid.col = 6
vGrid.Lock = False

vGrid.Row = fila
vGrid.col = 7
vGrid.Lock = True

vBodega = ""

strSQL = "select cod_producto,descripcion,precio_regular,impuesto_ventas from pv_productos" _
       & " where cod_producto = '" & vCriterio & "'"
Call OpenRecordSet(rs, strSQL)
If Not rs.EOF And Not rs.BOF Then vPaso = True

If Not vPaso Then
  rs.Close
  strSQL = "select cod_producto,descripcion,precio_regular,impuesto_ventas from pv_productos" _
         & " where cod_barras = '" & vCriterio & "'"
  Call OpenRecordSet(rs, strSQL)
  If Not rs.EOF And Not rs.BOF Then vPaso = True
End If

If Not vPaso Then
  rs.Close
  strSQL = "select cod_producto,descripcion,precio_regular,impuesto_ventas from pv_productos" _
         & " where cod_fabricante = '" & vCriterio & "'"
  Call OpenRecordSet(rs, strSQL)
  If Not rs.EOF And Not rs.BOF Then vPaso = True
End If

If Not vPaso Then
  MsgBox "No se encontró el Articulo en la Base de Datos...", vbExclamation
Else
 
 If vGrid.MaxRows > 2 Then
  vGrid.Row = fila - 1
  vGrid.col = 4
  vBodega = vGrid.Text
 End If
  
  vGrid.Row = fila
  vGrid.col = 1
  vGrid.Text = rs!Cod_Producto
  vGrid.col = 2
  vGrid.Text = rs!Descripcion
  vGrid.col = 3
  vGrid.Text = 1
    
  vGrid.col = 4
  vBodega = vGrid.Text
  
  vGrid.col = 5
  vGrid.Text = CStr(rs!precio_regular)
  vGrid.col = 6
  vGrid.Text = CStr(rs!impuesto_ventas)
  
  'Verificar si existe precio especificado en combo y si es asi cambiarlo
  strSQL = "select monto from pv_producto_precios where cod_producto = '" _
         & rs!Cod_Producto & "' and cod_precio = '" & fxCodigoCbo(cboPrecio) & "'"
  rs.Close
  rs.CursorLocation = adUseServer
  Call OpenRecordSet(rs, strSQL)
  If Not rs.EOF And Not rs.BOF Then
    vGrid.col = 5
    vGrid.Text = CStr(rs!Monto)
  End If
  
  Call vGrid_KeyPress(vbKeyReturn)
  
End If
rs.Close

'Si la caja no puede modificar los precios, bloquea las columnas de precios e Impuestos
'If Not gCajas.ModPrecios Then
    vGrid.Row = fila
    vGrid.col = 5
    vGrid.Lock = True
    
    vGrid.Row = fila
    vGrid.col = 6
    vGrid.Lock = True

    vGrid.Row = fila
    vGrid.col = 7
    vGrid.Lock = True
'End If

End Sub


Private Sub txtNombre_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then vGrid.SetFocus

If KeyCode = vbKeyF4 Then
  gBusquedas.Convertir = "N"
  gBusquedas.Columna = "nombre"
  gBusquedas.Orden = "nombre"
  gBusquedas.Consulta = "select cedula,nombre from pv_clientes"
  gBusquedas.Filtro = ""
  frmBusquedas.Show vbModal
  txtCedula = gBusquedas.Resultado
  txtNombre = gBusquedas.Resultado2
End If

End Sub

Private Sub vGrid_KeyDown(KeyCode As Integer, Shift As Integer)
Dim i As Variant, lng As Long, vTemp(7) As Variant, x As Integer

'Abrir Nueva Linea
If vGrid.ActiveCol = vGrid.MaxCols And (KeyCode = vbKeyReturn Or KeyCode = vbKeyTab) Then
  vGrid.Row = vGrid.ActiveRow
  If vGrid.MaxRows <= vGrid.ActiveRow Then
    vGrid.MaxRows = vGrid.MaxRows + 1
    vGrid.Row = vGrid.MaxRows
    Call sbCalculaTotales
  End If
End If

'Consulta Articulo
If vGrid.ActiveCol = 1 And KeyCode = vbKeyReturn Then
  vGrid.col = vGrid.ActiveCol
  vGrid.Row = vGrid.ActiveRow
  Call sbConsultaArticulo(vGrid.ActiveRow, vGrid.ActiveCol, vGrid.Text)
End If

'Consular Bodegas
If vGrid.ActiveCol = 4 And KeyCode = vbKeyF4 Then
   gBusquedas.Columna = "cod_bodega"
   gBusquedas.Orden = "cod_bodega"
   gBusquedas.Consulta = "select cod_bodega,descripcion from pv_bodegas"
   gBusquedas.Filtro = " and permite_salidas = 1"
   frmBusquedas.Show vbModal
   vGrid.Row = vGrid.ActiveRow
   vGrid.col = 4
   vGrid.Text = gBusquedas.Resultado
End If

'Consular Articulo
If vGrid.ActiveCol = 1 And KeyCode = vbKeyF4 Then
   frmBusquedaArticulos.Show vbModal
   vGrid.Row = vGrid.ActiveRow
   vGrid.col = 1
   vGrid.Text = gBusquedas.Resultado
End If


'Borrar una linea
If KeyCode = vbKeyDelete Then
  vGrid.Row = vGrid.ActiveRow
  vGrid.col = 7
  For lng = vGrid.ActiveRow To vGrid.MaxRows
     vGrid.Row = lng + 1
     For x = 1 To 7
        vGrid.col = x
        vTemp(x) = vGrid.Text
     Next x

     vGrid.Row = lng
     For x = 1 To 7
       vGrid.col = x
       vGrid.Text = vTemp(x)
     Next x
  Next lng
  vGrid.MaxRows = vGrid.MaxRows - 1
  If vGrid.MaxRows = 0 Then vGrid.MaxRows = 1
  Call sbCalculaTotales
End If

'Inserta Linea
If KeyCode = vbKeyInsert Then
    vGrid.MaxRows = vGrid.MaxRows + 1
    vGrid.InsertRows vGrid.ActiveRow, 1
    vGrid.Row = vGrid.ActiveRow
End If


End Sub



Private Sub vGrid_KeyPress(KeyAscii As Integer)
Dim curCantidad As Currency, curPrecio As Currency, curIV As Currency

On Error GoTo vError
'Calcula Total
Select Case vGrid.ActiveCol
  Case 3, 5, 6
    vGrid.Row = vGrid.ActiveRow
    vGrid.col = 3
    curCantidad = CCur(vGrid.Text)
    vGrid.col = 5
    curPrecio = CCur(vGrid.Text)
    vGrid.col = 6
    curIV = CCur(vGrid.Text)
    vGrid.col = 7
    vGrid.Text = (curPrecio * curCantidad) + ((curPrecio * curCantidad) * (curIV / 100))
   Call sbCalculaTotales
  Case Else 'No Aplica
End Select
vError:
End Sub





