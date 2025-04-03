VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Begin VB.Form frmInsVendedores 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Registro de Vendedores"
   ClientHeight    =   3825
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   9915
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3825
   ScaleWidth      =   9915
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtNombre 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2400
      MaxLength       =   38
      TabIndex        =   1
      Top             =   480
      Width           =   6495
   End
   Begin VB.TextBox txtCodigo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      DataField       =   "e"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1200
      TabIndex        =   0
      Top             =   480
      Width           =   1215
   End
   Begin TabDlg.SSTab ssTab 
      Height          =   2655
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   4683
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "General"
      TabPicture(0)   =   "frmInsVendedores.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1(13)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label1(1)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label1(0)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label1(2)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label1(3)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Line1"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label1(4)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label1(5)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label1(6)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "txtIdentificacion"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "chkActivo"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "chkComision"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "txtComisionVentaPorc"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "txtComisionCtaPorc"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "txtComisionCtaInicio"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "txtComisionCtaCorte"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "cboTipo"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "cboBancos"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "txtCtaBanco"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).ControlCount=   19
      TabCaption(1)   =   "Pólizas Activas"
      TabPicture(1)   =   "frmInsVendedores.frx":0121
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lsw"
      Tab(1).Control(1)=   "Line2(2)"
      Tab(1).ControlCount=   2
      Begin VB.TextBox txtCtaBanco 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   5520
         MaxLength       =   30
         TabIndex        =   22
         Top             =   1200
         Width           =   2535
      End
      Begin VB.ComboBox cboBancos 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "frmInsVendedores.frx":01DA
         Left            =   1800
         List            =   "frmInsVendedores.frx":01E4
         Style           =   2  'Dropdown List
         TabIndex        =   21
         Top             =   840
         Width           =   6255
      End
      Begin VB.ComboBox cboTipo 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "frmInsVendedores.frx":01FA
         Left            =   1800
         List            =   "frmInsVendedores.frx":0204
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Top             =   1200
         Width           =   1695
      End
      Begin VB.TextBox txtComisionCtaCorte 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   7320
         TabIndex        =   18
         Top             =   2040
         Width           =   735
      End
      Begin VB.TextBox txtComisionCtaInicio 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6600
         TabIndex        =   16
         Top             =   2040
         Width           =   735
      End
      Begin VB.TextBox txtComisionCtaPorc 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4200
         TabIndex        =   14
         Top             =   2040
         Width           =   735
      End
      Begin VB.TextBox txtComisionVentaPorc 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1800
         TabIndex        =   13
         Top             =   2040
         Width           =   735
      End
      Begin VB.CheckBox chkComision 
         Appearance      =   0  'Flat
         Caption         =   "Aplica Comisión ?"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   360
         TabIndex        =   12
         Top             =   1680
         Width           =   1695
      End
      Begin VB.CheckBox chkActivo 
         Appearance      =   0  'Flat
         Caption         =   "Activo ?"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   4080
         TabIndex        =   4
         Top             =   480
         Width           =   1215
      End
      Begin VB.TextBox txtIdentificacion 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1800
         TabIndex        =   3
         Top             =   480
         Width           =   2055
      End
      Begin MSComctlLib.ListView lsw 
         Height          =   2055
         Left            =   -74760
         TabIndex        =   5
         Top             =   420
         Width           =   9135
         _ExtentX        =   16113
         _ExtentY        =   3625
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         HotTracking     =   -1  'True
         HoverSelection  =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   7
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Id. Trámite"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Operación"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   2
            Text            =   "Línea"
            Object.Width           =   1658
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   3
            Text            =   "Fecha"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   4
            Text            =   "Dias Transc."
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Text            =   "T.Deuda"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Abogado"
            Object.Width           =   6068
         EndProperty
      End
      Begin VB.Label Label1 
         Caption         =   "Cuenta Bancaria"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   6
         Left            =   3840
         TabIndex        =   25
         Top             =   1200
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Tipo de Pago"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   5
         Left            =   360
         TabIndex        =   24
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Banco"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   360
         TabIndex        =   23
         Top             =   840
         Width           =   1095
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         X1              =   360
         X2              =   8880
         Y1              =   1800
         Y2              =   1800
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Inicio / Corte"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   6720
         TabIndex        =   19
         Top             =   1800
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Rango de Cuotas para Comisión"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   2
         Left            =   5160
         TabIndex        =   17
         Top             =   2040
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "(%) Comisión s/Cuotas "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   0
         Left            =   2760
         TabIndex        =   15
         Top             =   2040
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "(%) Comisión de Venta"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   1
         Left            =   360
         TabIndex        =   11
         Top             =   2040
         Width           =   1335
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FFFFFF&
         Index           =   2
         X1              =   -65520
         X2              =   -74760
         Y1              =   1680
         Y2              =   1680
      End
      Begin VB.Label Label1 
         Caption         =   "Identificación"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   13
         Left            =   360
         TabIndex        =   6
         Top             =   480
         Width           =   1095
      End
   End
   Begin MSComCtl2.FlatScrollBar FlatScrollBar 
      Height          =   255
      Left            =   9000
      TabIndex        =   7
      Top             =   480
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   450
      _Version        =   393216
      Arrows          =   65536
      Orientation     =   1572865
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
            Picture         =   "frmInsVendedores.frx":021A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInsVendedores.frx":36AC
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInsVendedores.frx":6B3E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInsVendedores.frx":6C5C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin ComCtl3.CoolBar CoolBarX 
      Align           =   1  'Align Top
      Height          =   390
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   9915
      _ExtentX        =   17489
      _ExtentY        =   688
      BandCount       =   2
      _CBWidth        =   9915
      _CBHeight       =   390
      _Version        =   "6.7.9782"
      Child1          =   "tlb"
      MinWidth1       =   1800
      MinHeight1      =   330
      Width1          =   1800
      NewRow1         =   0   'False
      Child2          =   "tlbAux"
      MinHeight2      =   330
      Width2          =   1110
      NewRow2         =   0   'False
      Begin MSComctlLib.Toolbar tlb 
         Height          =   330
         Left            =   165
         TabIndex        =   9
         Top             =   30
         Width           =   8520
         _ExtentX        =   15028
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
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
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
      Left            =   120
      TabIndex        =   10
      Top             =   480
      Width           =   975
   End
End
Attribute VB_Name = "frmInsVendedores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vEdita As Boolean, vCodigo As String
Dim vScroll As Boolean, vPaso As Boolean


Private Sub cboBancos_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then cboTipo.SetFocus
End Sub

Private Sub cboTipo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtCtaBanco.SetFocus
End Sub

Private Sub FlatScrollBar_Change()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

If txtCodigo.Text = "" Or Not IsNumeric(txtCodigo.Text) Then
   txtCodigo.Text = 0
End If


If vScroll Then
    strSQL = "select Top 1 cod_vendedor from INS_Vendedores"
    
    If FlatScrollBar.Value = 1 Then
       strSQL = strSQL & " where cod_vendedor > " & txtCodigo.Text & " order by cod_vendedor asc"
    Else
       strSQL = strSQL & " where cod_vendedor < " & txtCodigo.Text & " order by cod_vendedor desc"
    End If
    
    rs.Open strSQL, glogon.Conection, adOpenStatic
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
  MsgBox Err.Description, vbCritical

End Sub

Private Sub Form_Activate()
vModulo = 17
End Sub

Private Sub Form_Load()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

 vModulo = 17

 vEdita = True
 Call sbToolBarIconos(tlb)
 Call sbToolBar(tlb, "nuevo")
 
 
 vScroll = False
   FlatScrollBar.Value = 0
 vScroll = True
 
SSTab.Tab = 0

vEdita = False


 Call sbLimpiaPantalla

 Call Formularios(Me)
 Call RefrescaTags(Me)

Exit Sub

vError:
  MsgBox Err.Description, vbExclamation

End Sub

Private Sub sbLimpiaPantalla()
Dim strSQL As String

vCodigo = 0
txtCodigo = ""

strSQL = "select rtrim(Descripcion) as ItmX, Id_Banco as IdX from Tes_Bancos where estado = 'A'"
Call sbLlenaCbo(cboBancos, strSQL, False, True)

txtIdentificacion.Text = ""
txtNombre.Text = ""

chkActivo.Value = vbChecked

With cboTipo
    .Clear
    .AddItem "CK - Cheque"
    .AddItem "TE - Transferencia"
    .Text = "TE - Transferencia"
End With

txtCtaBanco.Text = "0"


chkComision.Value = vbUnchecked
txtComisionCtaCorte.Text = 0
txtComisionCtaInicio.Text = 0
txtComisionCtaPorc.Text = 0
txtComisionVentaPorc.Text = 0


SSTab.Tab = 0
SSTab.TabEnabled(1) = False

End Sub




Private Sub ssTab_Click(PreviousTab As Integer)
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListItem, curMonto As Currency, curSaldo As Currency

If vCodigo = "" Then
  SSTab.Tab = 0
  Exit Sub
End If

Me.MousePointer = vbHourglass


vPaso = True
Select Case SSTab.Tab
   Case 1 'Operaciones
      
'       vPaso = True
'       lswOperaciones.ListItems.Clear
'       curMonto = 0
'       curSaldo = 0
'
'       strSQL = "exec spCxCPersonasCuentas '" & txtCodigo.Text & "','A'"
'       rs.Open strSQL, glogon.Conection, adOpenStatic
'       Do While Not rs.EOF
'         Set itmX = lswOperaciones.ListItems.Add(, , rs!Operacion)
'             itmX.SubItems(1) = rs!Num_Documento
'             itmX.SubItems(2) = Format(rs!Activa_Fecha, "dd/mm/yyyy")
'             itmX.SubItems(3) = Format(rs!Fecha_Vencimiento, "dd/mm/yyyy")
'             itmX.SubItems(4) = Format(rs!Fecha_Pago, "dd/mm/yyyy")
'             itmX.SubItems(5) = Format(rs!Monto, "Standard")
'             itmX.SubItems(6) = Format(rs!Saldo, "Standard")
'             itmX.SubItems(7) = rs!Estado
'             itmX.SubItems(8) = rs!ConceptoDesc
'             itmX.SubItems(9) = rs!OficinaDesc
'             itmX.SubItems(10) = rs!Nombre_Pagador
'
'             curMonto = curMonto + rs!Monto
'             curSaldo = curSaldo + rs!Saldo
'
'          rs.MoveNext
'       Loop
'       rs.Close
'         Set itmX = lswOperaciones.ListItems.Add(, , "---")
'             itmX.SubItems(5) = "-----------"
'             itmX.SubItems(6) = "-----------"
'         Set itmX = lswOperaciones.ListItems.Add(, , lswOperaciones.ListItems.Count - 1)
'             itmX.SubItems(5) = Format(curMonto, "Standard")
'             itmX.SubItems(6) = Format(curSaldo, "Standard")
'
       
       
       vPaso = False
   
End Select

vPaso = False

Me.MousePointer = vbDefault

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox Err.Description, vbCritical

End Sub

Private Sub tlb_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim strSQL As String

Select Case UCase(Button.Key)
    Case "INSERTAR", "NUEVO"
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
       gBusquedas.Consulta = "select cod_vendedor,nombre from INS_Vendedores"
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

strSQL = "select V.*,rtrim(B.descripcion) as BancoDesc" _
       & " from INS_Vendedores V inner join Tes_Bancos B on V.cod_Banco = B.id_Banco" _
       & " where V.cod_vendedor = " & pCodigo
rs.Open strSQL, glogon.Conection, adOpenStatic

If Not rs.BOF And Not rs.EOF Then
  Call sbToolBar(tlb, "activo")
  vEdita = True

  vCodigo = rs!cod_vendedor
  txtCodigo.Text = rs!cod_vendedor
  
  txtIdentificacion.Text = rs!Cedula & ""
  txtNombre = rs!Nombre & ""
  
  chkActivo.Value = rs!Activo
  chkComision.Value = rs!Comision_Aplica
    
  txtComisionCtaCorte.Text = rs!comision_cuota_corte
  txtComisionCtaInicio.Text = rs!comision_cuota_inicio
  
  txtComisionCtaPorc.Text = rs!comision_porc_Cuotas
  txtComisionVentaPorc.Text = rs!comision_porc_venta
  
  If rs!Tipo_Emision = "CK" Then
     cboTipo.Text = "CK - Cheque"
  Else
     cboTipo.Text = "TE - Transferencia"
  End If
  
  txtCtaBanco.Text = rs!Cuenta_Bancaria & ""
 
  cboBancos.Text = rs!BancoDesc
 
  SSTab.Tab = 0
  SSTab.TabEnabled(1) = True

Else
  MsgBox "No se encontró registro verifique...", vbInformation
End If

rs.Close

Call RefrescaTags(Me)

Me.MousePointer = vbDefault

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox Err.Description, vbCritical


End Sub


Private Function fxValida() As Boolean
Dim strSQL As String, rs As New ADODB.Recordset
Dim vMensaje As String

vMensaje = ""
fxValida = True

If txtNombre = "" Then vMensaje = vMensaje & vbCrLf & " - Nombre no es válido ..."

strSQL = "select count(*) as 'Existe' from INS_Vendedores" _
        & " where cedula = '" & txtIdentificacion.Text & "' and cod_vendedor <> " & vCodigo
rs.Open strSQL, glogon.Conection, adOpenStatic
If rs!existe > 0 Then
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
  strSQL = "update INS_Vendedores set nombre = '" & Trim(txtNombre.Text) & "',cedula = '" & txtIdentificacion.Text & "',Activo = " & chkActivo.Value _
         & ", cod_Banco = " & cboBancos.ItemData(cboBancos.ListIndex) & ", Tipo_Emision = '" & SIFGlobal.fxSIFCodText(cboTipo.Text) & "', Cuenta_Bancaria = '" & txtCtaBanco.Text _
         & "',Comision_Aplica = " & chkComision.Value & ",comision_porc_venta = " & CCur(txtComisionVentaPorc.Text) & ",comision_porc_Cuotas = " & CCur(txtComisionCtaPorc.Text) _
         & ", comision_cuota_inicio = " & txtComisionCtaInicio.Text & ", comision_cuota_corte = " & txtComisionCtaCorte.Text _
         & " where cod_vendedor = " & vCodigo
  glogon.Conection.Execute strSQL
  
  Call Bitacora("Modifica", "INS Vendedores:  " & vCodigo)

Else
   'Extraer el Ultimo
   strSQL = "select coalesce(max(cod_vendedor),0) as Ultimo from INS_Vendedores"
   rs.Open strSQL, glogon.Conection, adOpenStatic
     txtCodigo.Text = rs!ultimo + 1
   rs.Close
   vCodigo = txtCodigo.Text
   
   strSQL = "insert into INS_Vendedores(cod_vendedor,cedula, nombre,cod_Banco,tipo_Emision,Cuenta_Bancaria,comision_Aplica,comision_porc_venta" _
          & ",comision_porc_Cuotas,comision_cuota_inicio,comision_cuota_corte,Activo,registro_fecha,registro_usuario)" _
          & " values(" & vCodigo & ",'" & txtIdentificacion.Text & "','" & txtNombre.Text & "'," & cboBancos.ItemData(cboBancos.ListIndex) & ",'" _
          & SIFGlobal.fxSIFCodText(cboTipo.Text) & "','" & txtCtaBanco.Text & "'," & chkComision.Value & "," & CCur(txtComisionVentaPorc.Text) _
          & "," & CCur(txtComisionCtaPorc.Text) & "," & txtComisionCtaInicio.Text & "," & txtComisionCtaCorte.Text _
          & "," & chkActivo.Value & ",dbo.MyGetdate(),'" & glogon.Usuario & "')"
   glogon.Conection.Execute strSQL

   Call Bitacora("Registra", "INS Vendedores:  " & vCodigo)

End If

SSTab.TabEnabled(1) = True

Call sbToolBar(tlb, "activo")

Call RefrescaTags(Me)

MsgBox "Información guardada satisfactoriamente...", vbInformation

Exit Sub

vError:
 MsgBox Err.Description, vbCritical


End Sub

Private Sub sbBorrar()
Dim i As Integer, strSQL As String

On Error GoTo vError

i = MsgBox("Esta Seguro que desea borrar este registro", vbYesNo)

If i = vbYes Then
  strSQL = "delete INS_Vendedores where cod_vendedor = " & vCodigo
  glogon.Conection.Execute strSQL

  Call Bitacora("Elimina", "INS Vendedores:  " & vCodigo)
  Call sbLimpiaPantalla
  Call sbToolBar(tlb, "nuevo")
  Call RefrescaTags(Me)

End If

Exit Sub

vError:
 MsgBox Err.Description, vbCritical

End Sub


Private Sub txtCodigo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtNombre.SetFocus

If KeyCode = vbKeyF4 Then
  gBusquedas.Convertir = "N"
  gBusquedas.Columna = "cod_vendedor"
  gBusquedas.Orden = "cod_vendedor"
  gBusquedas.Consulta = "select cod_vendedor,nombre from INS_Vendedores"
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


Private Sub txtComisionCtaInicio_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtComisionCtaCorte.SetFocus
End Sub

Private Sub txtComisionCtaPorc_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtComisionCtaInicio.SetFocus
End Sub

Private Sub txtComisionVentaPorc_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtComisionCtaPorc.SetFocus
End Sub

Private Sub txtCtaBanco_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtComisionVentaPorc.SetFocus
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
  gBusquedas.Consulta = "select cod_vendedor,nombre from INS_Vendedores"
  gBusquedas.Filtro = ""
  frmBusquedas.Show vbModal
  txtCodigo = gBusquedas.Resultado
  If txtCodigo <> "" Then Call sbConsulta(gBusquedas.Resultado)
End If

End Sub

