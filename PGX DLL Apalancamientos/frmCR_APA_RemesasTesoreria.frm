VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.Ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmCR_APA_RemesasTesoreria 
   Caption         =   "Administración de Pagarés"
   ClientHeight    =   7290
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10395
   Icon            =   "frmCR_APA_RemesasTesoreria.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7290
   ScaleWidth      =   10395
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   9240
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_APA_RemesasTesoreria.frx":6852
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_APA_RemesasTesoreria.frx":1D214
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_APA_RemesasTesoreria.frx":33BD6
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_APA_RemesasTesoreria.frx":48D48
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_APA_RemesasTesoreria.frx":5DEBA
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_APA_RemesasTesoreria.frx":6471C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_APA_RemesasTesoreria.frx":6AF7E
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_APA_RemesasTesoreria.frx":717E0
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin TabDlg.SSTab ssTab 
      Height          =   6255
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   10215
      _ExtentX        =   18018
      _ExtentY        =   11033
      _Version        =   393216
      Style           =   1
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      ForeColor       =   16711680
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Remesas"
      TabPicture(0)   =   "frmCR_APA_RemesasTesoreria.frx":78042
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1(5)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label1(6)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Line1(5)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Line1(4)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Line1(3)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Line1(2)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Line1(0)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Line1(6)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Line1(7)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Line2"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Line1(10)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label2(0)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Label2(9)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Label2(5)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Label2(4)"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Label2(3)"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Label2(2)"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Label2(6)"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Label2(1)"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "tlb"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "lswRemesas"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "dtpCorte"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "dtpInicio"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "txtTotal"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "txtEstado"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "txtFecha"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "txtUsuario"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "txtRemesa"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "txtNotas"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).ControlCount=   29
      TabCaption(1)   =   "Cargado"
      TabPicture(1)   =   "frmCR_APA_RemesasTesoreria.frx":7805E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "txtCargaTotal"
      Tab(1).Control(1)=   "cboCarga"
      Tab(1).Control(2)=   "chkCarga"
      Tab(1).Control(3)=   "lswCarga"
      Tab(1).Control(4)=   "tlbCarga"
      Tab(1).Control(5)=   "Label2(8)"
      Tab(1).Control(6)=   "Label3(0)"
      Tab(1).Control(7)=   "Line1(8)"
      Tab(1).Control(8)=   "Label2(7)"
      Tab(1).ControlCount=   9
      TabCaption(2)   =   "Pago"
      TabPicture(2)   =   "frmCR_APA_RemesasTesoreria.frx":7807A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "txtPagoTotal"
      Tab(2).Control(1)=   "cboPago"
      Tab(2).Control(2)=   "tlbPago"
      Tab(2).Control(3)=   "lswPago"
      Tab(2).Control(4)=   "Label2(11)"
      Tab(2).Control(5)=   "Label3(1)"
      Tab(2).Control(6)=   "Label2(10)"
      Tab(2).Control(7)=   "Line1(9)"
      Tab(2).ControlCount=   8
      TabCaption(3)   =   "Reportes"
      TabPicture(3)   =   "frmCR_APA_RemesasTesoreria.frx":78096
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "chkRemesaInd"
      Tab(3).Control(1)=   "opt(1)"
      Tab(3).Control(2)=   "opt(0)"
      Tab(3).Control(3)=   "txtRepRemesas"
      Tab(3).Control(4)=   "cmdReporte"
      Tab(3).Control(5)=   "lswRep"
      Tab(3).Control(6)=   "lblRemesa"
      Tab(3).Control(7)=   "Label16(2)"
      Tab(3).Control(8)=   "Line1(1)"
      Tab(3).Control(9)=   "Label16(4)"
      Tab(3).ControlCount=   10
      Begin VB.TextBox txtCargaTotal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   -67440
         Locked          =   -1  'True
         TabIndex        =   16
         Top             =   5760
         Width           =   2535
      End
      Begin VB.TextBox txtNotas 
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
         ForeColor       =   &H00C00000&
         Height          =   675
         Left            =   3120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   15
         Top             =   3360
         Width           =   6735
      End
      Begin VB.ComboBox cboCarga 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   330
         Left            =   -71880
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   480
         Width           =   6975
      End
      Begin VB.CheckBox chkCarga 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Todos"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   210
         Left            =   -74880
         TabIndex        =   13
         Top             =   1215
         Width           =   1455
      End
      Begin VB.TextBox txtRemesa 
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
         Left            =   3120
         TabIndex        =   12
         Top             =   480
         Width           =   2655
      End
      Begin VB.TextBox txtUsuario 
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
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   3120
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   1920
         Width           =   2655
      End
      Begin VB.TextBox txtFecha 
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
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   3120
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   2280
         Width           =   2655
      End
      Begin VB.TextBox txtEstado 
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
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   3120
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   2640
         Width           =   2655
      End
      Begin VB.TextBox txtTotal 
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
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   3120
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   3000
         Width           =   2655
      End
      Begin VB.CheckBox chkRemesaInd 
         Appearance      =   0  'Flat
         Caption         =   "Indicar Remesa"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   -66600
         TabIndex        =   7
         Top             =   4920
         Width           =   1455
      End
      Begin VB.OptionButton opt 
         Appearance      =   0  'Flat
         Caption         =   "Detalle Agrupado de Remesa"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   -74640
         TabIndex        =   6
         Top             =   3840
         Width           =   2655
      End
      Begin VB.OptionButton opt 
         Appearance      =   0  'Flat
         Caption         =   "Detalle de Remesa"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   -74640
         TabIndex        =   5
         Top             =   3480
         Value           =   -1  'True
         Width           =   2055
      End
      Begin VB.TextBox txtRepRemesas 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   -65640
         TabIndex        =   4
         Text            =   "15"
         Top             =   3000
         Width           =   615
      End
      Begin VB.CommandButton cmdReporte 
         Caption         =   "&Reporte"
         Height          =   375
         Left            =   -66480
         TabIndex        =   3
         Top             =   5640
         Width           =   1455
      End
      Begin VB.TextBox txtPagoTotal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   -67440
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   5760
         Width           =   2535
      End
      Begin VB.ComboBox cboPago 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   330
         Left            =   -71880
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   480
         Width           =   6975
      End
      Begin MSComCtl2.DTPicker dtpInicio 
         Height          =   315
         Left            =   3120
         TabIndex        =   17
         Top             =   1560
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   295370755
         CurrentDate     =   36278
      End
      Begin MSComCtl2.DTPicker dtpCorte 
         Height          =   315
         Left            =   4440
         TabIndex        =   18
         Top             =   1560
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   295370755
         CurrentDate     =   36278
      End
      Begin MSComctlLib.ListView lswRemesas 
         Height          =   2055
         Left            =   3120
         TabIndex        =   19
         Top             =   4080
         Width           =   6735
         _ExtentX        =   11880
         _ExtentY        =   3625
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         HotTracking     =   -1  'True
         HoverSelection  =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Remesa"
            Object.Width           =   2187
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Usuario"
            Object.Width           =   4304
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Fecha"
            Object.Width           =   5715
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Inicio"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Corte"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Notas"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComctlLib.Toolbar tlb 
         Height          =   570
         Left            =   6000
         TabIndex        =   20
         Top             =   960
         Width           =   3465
         _ExtentX        =   6112
         _ExtentY        =   1005
         ButtonWidth     =   1111
         ButtonHeight    =   1005
         AllowCustomize  =   0   'False
         Style           =   1
         ImageList       =   "ImageList1"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   5
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Nuevo"
               Key             =   "nuevo"
               ImageIndex      =   6
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Borrar"
               Key             =   "borrar"
               ImageIndex      =   7
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Object.Visible         =   0   'False
               Caption         =   "Reportes"
               Key             =   "reportes"
               ImageIndex      =   8
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "ayuda"
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ListView lswCarga 
         Height          =   4215
         Left            =   -74880
         TabIndex        =   21
         Top             =   1440
         Width           =   9975
         _ExtentX        =   17595
         _ExtentY        =   7435
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         HotTracking     =   -1  'True
         HoverSelection  =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   9
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Nº Pago"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   1
            Text            =   "Documento"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Monto"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   3
            Text            =   "Fecha Pago"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   4
            Text            =   "Operación"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Acreedor"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Autorizado"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   7
            Text            =   "Usuario"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "Cod Acreedor"
            Object.Width           =   2646
         EndProperty
      End
      Begin MSComctlLib.Toolbar tlbCarga 
         Height          =   330
         Left            =   -71880
         TabIndex        =   22
         Top             =   840
         Width           =   4545
         _ExtentX        =   8017
         _ExtentY        =   582
         ButtonWidth     =   1905
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         Style           =   1
         TextAlignment   =   1
         ImageList       =   "ImageList1"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   5
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&Buscar"
               Key             =   "buscar"
               Object.ToolTipText     =   "Buscar Casos para Comision"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&Cargar"
               Key             =   "cargar"
               Object.ToolTipText     =   "cargar datos"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Cerrar"
               Key             =   "cerrar"
               Object.ToolTipText     =   "Cierra Remesa"
               ImageIndex      =   3
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ListView lswRep 
         Height          =   2175
         Left            =   -74760
         TabIndex        =   23
         Top             =   720
         Width           =   9735
         _ExtentX        =   17171
         _ExtentY        =   3836
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         HotTracking     =   -1  'True
         HoverSelection  =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Remesa"
            Object.Width           =   2187
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Usuario"
            Object.Width           =   4304
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Fecha"
            Object.Width           =   5715
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Inicio"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Corte"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Notas"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComctlLib.Toolbar tlbPago 
         Height          =   330
         Left            =   -71880
         TabIndex        =   24
         Top             =   840
         Width           =   2625
         _ExtentX        =   4630
         _ExtentY        =   582
         ButtonWidth     =   1905
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         Style           =   1
         TextAlignment   =   1
         ImageList       =   "ImageList1"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   3
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&Buscar"
               Key             =   "buscar"
               Object.ToolTipText     =   "Buscar Casos para Pago"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&Pago"
               Key             =   "pago"
               Object.ToolTipText     =   "Crear Desembolso"
               ImageIndex      =   5
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ListView lswPago 
         Height          =   4215
         Left            =   -74880
         TabIndex        =   25
         Top             =   1440
         Width           =   9975
         _ExtentX        =   17595
         _ExtentY        =   7435
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         HotTracking     =   -1  'True
         HoverSelection  =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   8
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Nº Pago "
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   1
            Text            =   "Documento"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Monto"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   3
            Text            =   "Fecha Pago"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   4
            Text            =   "Operación"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Acreedor"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   6
            Text            =   "Usuario"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Cod Acreedor"
            Object.Width           =   2646
         EndProperty
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Remesa"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   44
         Top             =   480
         Width           =   2895
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Fechas"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   43
         Top             =   1320
         Width           =   2895
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Usuario"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   42
         Top             =   1920
         Width           =   2895
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Fecha"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   41
         Top             =   2280
         Width           =   2895
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Estado"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   40
         Top             =   2640
         Width           =   2895
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Total"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   39
         Top             =   3000
         Width           =   2895
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Notas"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   9
         Left            =   120
         TabIndex        =   38
         Top             =   3360
         Width           =   2895
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Lista de Remesas"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   37
         Top             =   4080
         Width           =   2895
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Remesa"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   11
         Left            =   -74880
         TabIndex        =   36
         Top             =   480
         Width           =   2895
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Remesa"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   8
         Left            =   -74880
         TabIndex        =   35
         Top             =   480
         Width           =   2895
      End
      Begin VB.Label Label3 
         Caption         =   "Total...:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   -68280
         TabIndex        =   34
         Top             =   5760
         Width           =   855
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   10
         X1              =   120
         X2              =   3000
         Y1              =   3600
         Y2              =   3600
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   8
         X1              =   -74880
         X2              =   -72000
         Y1              =   720
         Y2              =   720
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Seleccione los Retiros / Liquidaciones Disponibles"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   7
         Left            =   -74880
         TabIndex        =   33
         Top             =   1200
         Width           =   9975
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FFFFFF&
         X1              =   0
         X2              =   9840
         Y1              =   840
         Y2              =   840
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   7
         X1              =   120
         X2              =   3000
         Y1              =   1560
         Y2              =   1560
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   6
         X1              =   120
         X2              =   3000
         Y1              =   4320
         Y2              =   4320
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   0
         X1              =   120
         X2              =   3000
         Y1              =   720
         Y2              =   720
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   2
         X1              =   120
         X2              =   3000
         Y1              =   2160
         Y2              =   2160
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   3
         X1              =   120
         X2              =   3000
         Y1              =   2520
         Y2              =   2520
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   4
         X1              =   120
         X2              =   3000
         Y1              =   2880
         Y2              =   2880
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   5
         X1              =   120
         X2              =   3000
         Y1              =   3240
         Y2              =   3240
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Corte"
         ForeColor       =   &H00FF0000&
         Height          =   315
         Index           =   6
         Left            =   4440
         TabIndex        =   32
         Top             =   1320
         Width           =   1335
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Inicio"
         ForeColor       =   &H00FF0000&
         Height          =   315
         Index           =   5
         Left            =   3120
         TabIndex        =   31
         Top             =   1320
         Width           =   1335
      End
      Begin VB.Label lblRemesa 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   300
         Left            =   -74760
         TabIndex        =   30
         Top             =   3000
         Width           =   5295
      End
      Begin VB.Label Label16 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Seleccione la Remesa que Desea Visualizar"
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
         Height          =   300
         Index           =   2
         Left            =   -74760
         TabIndex        =   29
         Top             =   480
         Width           =   9735
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   1
         X1              =   -74760
         X2              =   -65040
         Y1              =   5520
         Y2              =   5520
      End
      Begin VB.Label Label16 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Remesas - visualizar últimas"
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
         Height          =   300
         Index           =   4
         Left            =   -69960
         TabIndex        =   28
         Top             =   3000
         Width           =   4935
      End
      Begin VB.Label Label3 
         Caption         =   "Total...:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   -68280
         TabIndex        =   27
         Top             =   5760
         Width           =   855
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   " Desembolsos pendientes de Pago por Tesorería"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   10
         Left            =   -74880
         TabIndex        =   26
         Top             =   1200
         Width           =   9975
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   9
         X1              =   -74880
         X2              =   -72000
         Y1              =   720
         Y2              =   720
      End
   End
   Begin MSComctlLib.ProgressBar PrgBar 
      Align           =   2  'Align Bottom
      Height          =   150
      Left            =   0
      TabIndex        =   45
      Top             =   7140
      Width           =   10395
      _ExtentX        =   18336
      _ExtentY        =   265
      _Version        =   393216
      Appearance      =   0
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   360
      Picture         =   "frmCR_APA_RemesasTesoreria.frx":780B2
      Top             =   120
      Width           =   480
   End
   Begin VB.Label Label1 
      Caption         =   "Remesas de Pago a Tesorería"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   1200
      TabIndex        =   46
      Top             =   240
      Width           =   6255
   End
End
Attribute VB_Name = "frmCR_APA_RemesasTesoreria"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim itmX As ListItem, vPaso As Boolean
Dim strSQL As String, rs As New ADODB.Recordset


Private Sub cboCarga_Click()
Dim vFechaInicio As Date, vFechaCorte As Date

lswCarga.ListItems.Clear

If vPaso Then Exit Sub
If cboCarga.ListCount <= 0 Then Exit Sub

vPaso = True

strSQL = "select FECHA_INICIO,FECHA_CORTE from CRD_APA_REMESASTESORERIA where REMESA = " & cboCarga.ItemData(cboCarga.ListIndex)
Call OpenRecordSet(rs, strSQL)
  vFechaInicio = rs!FECHA_INICIO
  vFechaCorte = rs!Fecha_Corte
rs.Close

End Sub



Private Sub cboPago_Click()

If vPaso Then Exit Sub
If cboPago.ListCount <= 0 Then Exit Sub

lswPago.ListItems.Clear
txtPagoTotal.Text = 0

End Sub

Private Sub chkCarga_Click()
Dim i As Integer, curTotal As Currency


For i = 1 To lswCarga.ListItems.Count
  lswCarga.ListItems.Item(i).Checked = chkCarga.Value
  
   If chkCarga.Value = vbChecked Then
       curTotal = curTotal + CCur(lswCarga.ListItems.Item(i).SubItems(2))
   End If
  
Next i

txtCargaTotal.Text = Format(curTotal, "Standard")

End Sub

Private Sub cmdReporte_Click()
Dim vTitulo As String, vSubTitulo As String, vFiltro As String
Dim strSQL As String, vTipoUser As String, xRemesa As String

On Error GoTo vError

If lblRemesa.Tag = "" Then Exit Sub

Me.MousePointer = vbHourglass


vSubTitulo = ""
vFiltro = ""
strSQL = ""


With frmContenedor.Crt
 .Reset
 .WindowShowGroupTree = True
 .WindowShowPrintSetupBtn = True
 .WindowShowRefreshBtn = True
 .WindowShowSearchBtn = True
 .WindowState = crptMaximized
 .WindowTitle = "Reportes del Módulo Administración Pagarés"

 .Connect = glogon.ConectRPT

If chkRemesaInd.Value = vbChecked Then
   xRemesa = InputBox("Indique la Remesa que desea consultar", "Remesas de Tesorería")
  If Len(Trim(xRemesa)) = 0 Then xRemesa = "0"
  lblRemesa.Tag = xRemesa
End If



 Select Case True
  Case opt.Item(0).Value 'Detalle Remesa
     .ReportFileName = SIFGlobal.fxPathReportes("Acreedores_RemesaTesDetalle.rpt")
     vSubTitulo = "REMESA : " & lblRemesa.Tag & " LISTADO : DETALLADO"
  Case opt.Item(1).Value 'Detalle Agrupado Remesa
     .ReportFileName = SIFGlobal.fxPathReportes("Acreedores_RemesaTesDetalleAgrp.rpt")
     vSubTitulo = "REMESA : " & lblRemesa.Tag & " LISTADO : DETALLADO AGRUPADO"
 End Select
 
 .Formulas(0) = "fxFecha='FECHA: " & Format(fxFechaServidor, "dd/mm/yyyy") & "'"
 .Formulas(1) = "fxEmpresa='" & GLOBALES.gstrNombreEmpresa & "'"
 .Formulas(2) = "fxUsuario='USER: " & glogon.Usuario & "'"
 .Formulas(3) = "fxTitulo='REMESA DE PAGO ADMIN PAGARÉS'"
 .Formulas(4) = "fxSubTitulo='" & vSubTitulo & "'"
 .Formulas(5) = "fxFiltro='" & vFiltro & "'"
 .SelectionFormula = "{CRD_APA_REMESASTESORERIA.Remesa} = " & lblRemesa.Tag
' .PrintReport
 .Action = 1

End With

Me.MousePointer = vbDefault
Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub



Private Sub Form_Load()

    vModulo = 14 'Modulo de Credito

    Me.Icon = Me.Picture

'Inicializa Barra
'Call sbToolBarIconos(tlb, False)
'Inicializa Seguridad

    ssTab.Tab = 0
' Call Formularios(Me)
    
    Call RefrescaTags(Me)
    
    If GLOBALES.gEnlace = 0 Then
        Call sbgCntParametros
    End If
    
    '' Carga nombre de la ternimal
    If Len(glogon.Maquina) = 0 Then
        Call sbMaquina
    End If
    
    Call sbLimpia
 
End Sub

Private Sub sbConsulta(vRemesa As Long)

Call sbLimpia
  
strSQL = "select * from CRD_APA_REMESASTESORERIA where remesa = " & vRemesa
Call OpenRecordSet(rs, strSQL)
If Not rs.BOF And Not rs.EOF Then

'  Call sbToolBar(tlb, "Activo")
  
  txtRemesa = rs!REMESA
  txtUsuario = rs!REGISTRO_USUARIO
  txtFecha = rs!REGISTRO_FECHA
  
  Select Case rs!Estado
    Case "A"
      txtEstado = "Remesa Abierta"
    Case "C"
      txtEstado = "Remesa Cerrada"
    Case "X"
      txtEstado = "Remesa Cargando"
    Case "P"
      txtEstado = "Remesa en Pagada"
  End Select
  
  dtpInicio.Value = rs!FECHA_INICIO
  dtpCorte.Value = rs!Fecha_Corte
  
  txtNotas.Text = rs!Notas
  

  With glogon
    .strSQL = "select isnull(sum(monto),0) as Total from CRD_APA_PAGOS" _
            & " where TESORERIA_REMESA = " & vRemesa & ""
    If .Recordset.State = 1 Then
       .Recordset.Close
    End If
    .Recordset.Open .strSQL, .Conection, adOpenStatic
    txtTotal.Text = Format(.Recordset!total, "Standard")
    .Recordset.Close
  End With
  
End If
rs.Close

End Sub


Private Sub lswCarga_ItemCheck(ByVal Item As MSComctlLib.ListItem)
Dim curTotal As Currency

If Trim(txtCargaTotal.Text) = "" Then txtCargaTotal.Text = 0

curTotal = CCur(txtCargaTotal.Text)

If Item.Checked Then

    If Not fxValidaParametrosTesoreria(Item) Then
        MsgBox "Faltan por definir parámetros en el acreedor"
        Item.Checked = False
        Exit Sub
    End If
    
    curTotal = curTotal + CCur(Item.SubItems(2))
Else
   curTotal = curTotal - CCur(Item.SubItems(2))
End If

txtCargaTotal.Text = Format(curTotal, "Standard")

End Sub

Private Function fxValidaParametrosTesoreria(ByVal Item As MSComctlLib.ListItem) As Boolean
    Dim TipoPago As String
    fxValidaParametrosTesoreria = True
    
    'Consulta el Tipo de Pago
    strSQL = "select FORMA_PAGO from CRD_APA_PAGOS " _
            & " where NPAGO = " & Item.Text _
            & " and OPERACION = " & pc(Trim(Item.SubItems(4))) _
            & " and COD_ACREEDOR = " & pc(Trim(Item.SubItems(8)))
            
    Call OpenRecordSet(rs, strSQL)
    If Not rs.EOF Then
        TipoPago = rs.Fields(0)
    End If
    rs.Close
    
    'Consulta el Cuentas
    strSQL = "select COD_CUENTA,COD_CUENTA_TRANSITORIA,COD_CUENTA_GASTOS,COD_CUENTA_CARGOS,COD_CUENTA_COMISION, " _
            & " BANCO_CK,BANCO_DC from CRD_APA_ACREEDORES where COD_ACREEDOR = " & pc(Trim(Item.SubItems(8)))
            
    Call OpenRecordSet(rs, strSQL)
    If Not rs.EOF Then
        If IsNull(rs.Fields(0)) Or rs.Fields(0) = Empty Then
            fxValidaParametrosTesoreria = False
            rs.Close
            Exit Function
        End If
        If IsNull(rs.Fields(1)) Or rs.Fields(1) = Empty Then
            fxValidaParametrosTesoreria = False
            rs.Close
            Exit Function
        End If
        If IsNull(rs.Fields(2)) Or rs.Fields(2) = Empty Then
            fxValidaParametrosTesoreria = False
            rs.Close
            Exit Function
        End If
        If IsNull(rs.Fields(3)) Or rs.Fields(3) = Empty Then
            fxValidaParametrosTesoreria = False
            rs.Close
            Exit Function
        End If
        If IsNull(rs.Fields(4)) Or rs.Fields(4) = Empty Then
            fxValidaParametrosTesoreria = False
            rs.Close
            Exit Function
        End If
        If TipoPago = "CK" Then
            If IsNull(rs.Fields(5)) Or rs.Fields(5) = Empty Then
                fxValidaParametrosTesoreria = False
                rs.Close
                Exit Function
            End If
        End If
        If TipoPago = "DC" Then
            If IsNull(rs.Fields(6)) Or rs.Fields(6) = Empty Then
                fxValidaParametrosTesoreria = False
                rs.Close
                Exit Function
            End If
        End If
    End If
    rs.Close

End Function


Private Sub lswCarga_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo vError
    
    lswCarga.SortKey = ColumnHeader.Index - 1
    
    If (lswCarga.SortOrder = lvwAscending) Then
        lswCarga.SortOrder = lvwDescending
    Else
        lswCarga.SortOrder = lvwAscending
    End If
    
    lswCarga.Sorted = True
    Exit Sub

vError:
   MsgBox "Ocurrió un error al ordenar los datos de la columna seleccionada.", vbCritical

End Sub

Private Sub lswRemesas_Click()
If lswRemesas.ListItems.Count <= 0 Then Exit Sub
Call sbConsulta(lswRemesas.SelectedItem)
End Sub



Private Sub lswRep_Click()
If lswRep.ListItems.Count <= 0 Then Exit Sub

lblRemesa.Caption = lswRep.SelectedItem.Text & " ¦ " & lswRep.SelectedItem.SubItems(1) _
            & " ¦ " & lswRep.SelectedItem.SubItems(2)
lblRemesa.Tag = lswRep.SelectedItem

End Sub

Private Sub ssTab_Click(PreviousTab As Integer)
 Call sbLimpia
End Sub


Private Sub tlb_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim i As Integer

On Error GoTo vError

Select Case UCase(Button.Key)
  Case "NUEVO"
     
    If txtRemesa.Text = "" Then
     
            strSQL = "select isnull(max(remesa),0) + 1 as Ultimo from CRD_APA_REMESASTESORERIA"
            Call OpenRecordSet(rs, strSQL)
            
                strSQL = "insert CRD_APA_REMESASTESORERIA(REMESA,REGISTRO_USUARIO,REGISTRO_FECHA,ESTADO,FECHA_INICIO,FECHA_CORTE,NOTAS)" _
                       & " values(" & rs!Ultimo & ",'" & glogon.Usuario & "',dbo.MyGetdate(),'A','" & Format(dtpInicio.Value, "yyyy/mm/dd") _
                       & "','" & Format(dtpCorte.Value, "yyyy/mm/dd") & "','" & txtNotas.Text & "')"
                Call ConectionExecute(strSQL)
                
                txtRemesa = rs!Ultimo
            rs.Close
            Call Bitacora("Registra", "Remesa de administración pagarés: " & txtRemesa)
    
    Else
        If txtEstado.Text <> "Remesa Cerrada" Then
                    
            strSQL = "update CRD_APA_REMESASTESORERIA set REGISTRO_USUARIO = '" & glogon.Usuario & "',FECHA_INICIO = '" & Format(dtpInicio.Value, "yyyy/mm/dd") _
                   & "',FECHA_CORTE = '" & Format(dtpCorte.Value, "yyyy/mm/dd") & "',NOTAS = '" & txtNotas.Text _
                   & "' where remesa = " & txtRemesa
             Call ConectionExecute(strSQL)
            Call Bitacora("Modifica", "Remesa de administración pagarés: " & txtRemesa)
        Else
            MsgBox "No se puede Modifica la remesa, porque esta ya fue cerrada...", vbExclamation
        End If
    End If
    
    Call sbLimpia
    
    
  Case "BORRAR"
     i = MsgBox("Esta Seguro que desea borrar este registro", vbYesNo)
     If i = vbYes Then
        If txtEstado.Text = "Remesa Abierta" Then
            strSQL = "delete CRD_APA_REMESASTESORERIA where remesa = " & txtRemesa
            Call ConectionExecute(strSQL)
            
            strSQL = "update CRD_APA_PAGOS set TESORERIA_REMESA = null  where TesoreriaRemesa = " & txtRemesa
            Call ConectionExecute(strSQL)
            
            'Call Bitacora("Elimina", "Remesa de administración pagarés: " & txtRemesa)
         End If
       
        Call sbLimpia
     End If

  Case "AYUDA"
'        frmContenedor.CD.HelpContext = Me.HelpContextID
'        frmContenedor.CD.ShowHelp

End Select

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub



Private Sub sbLimpia()

Select Case ssTab.Tab
  Case 0 'Remesas
     txtEstado = ""
     txtFecha = ""
     txtTotal = 0
     txtUsuario = ""
     txtRemesa = ""
     txtNotas.Text = ""
     
     dtpInicio.Value = fxFechaServidor
     dtpCorte.Value = dtpInicio.Value
     
     
     strSQL = "select TOP 50 * from CRD_APA_REMESASTESORERIA order by REGISTRO_FECHA desc"
     lswRemesas.ListItems.Clear

     Call OpenRecordSet(rs, strSQL)
     Do While Not rs.EOF
       With lswRemesas.ListItems
            Set itmX = .Add(, , rs!REMESA)
                itmX.SubItems(1) = rs!REGISTRO_USUARIO
                itmX.SubItems(2) = rs!REGISTRO_FECHA
                itmX.SubItems(3) = rs!FECHA_INICIO
                itmX.SubItems(4) = rs!Fecha_Corte
                itmX.SubItems(5) = rs!Notas
       
       End With
       rs.MoveNext
     Loop
     rs.Close
     
  Case 1 'Carga
    'Solo busca las remesas que se encuentran Abiertas
    
    vPaso = True
    
    cboCarga.Clear
    lswCarga.ListItems.Clear
    chkCarga.Value = vbUnchecked
        
    strSQL = "select * from CRD_APA_REMESASTESORERIA where ESTADO in('A','X') order by REGISTRO_FECHA desc"
    Call OpenRecordSet(rs, strSQL)
    Do While Not rs.EOF
      cboCarga.AddItem (Format(rs!REMESA, "0000") & "..." & Trim(rs!REGISTRO_USUARIO) & "..." & rs!REGISTRO_FECHA)
      cboCarga.ItemData(cboCarga.NewIndex) = rs!REMESA
      rs.MoveNext
    Loop
    If rs.RecordCount > 0 Then
       rs.MoveFirst
       cboCarga.Text = (Format(rs!REMESA, "0000") & "..." & Trim(rs!REGISTRO_USUARIO) & "..." & rs!REGISTRO_FECHA)
    End If
    
    rs.Close

    vPaso = False
    Call cboCarga_Click

  Case 2 'Pago
    vPaso = True
    
    cboPago.Clear
    lswPago.ListItems.Clear
        
    strSQL = "select * from CRD_APA_REMESASTESORERIA where ESTADO in('C') order by REGISTRO_FECHA desc"
    Call OpenRecordSet(rs, strSQL)
    Do While Not rs.EOF
      cboPago.AddItem (Format(rs!REMESA, "0000") & "..." & Trim(rs!REGISTRO_USUARIO) & "..." & rs!REGISTRO_FECHA)
      cboPago.ItemData(cboPago.NewIndex) = rs!REMESA
      rs.MoveNext
    Loop
    If rs.RecordCount > 0 Then
       rs.MoveFirst
       cboPago.Text = (Format(rs!REMESA, "0000") & "..." & Trim(rs!REGISTRO_USUARIO) & "..." & rs!REGISTRO_FECHA)
    End If
    
    rs.Close

    vPaso = False
    Call cboPago_Click


  Case 3 'Reportes
     strSQL = "select TOP " & txtRepRemesas.Text & " * from CRD_APA_REMESASTESORERIA order by REGISTRO_FECHA desc"
     lswRep.ListItems.Clear

     Call OpenRecordSet(rs, strSQL)
     Do While Not rs.EOF
       With lswRep.ListItems
            Set itmX = .Add(, , rs!REMESA)
                itmX.SubItems(1) = rs!REGISTRO_USUARIO
                itmX.SubItems(2) = rs!REGISTRO_FECHA
                itmX.SubItems(3) = rs!FECHA_INICIO
                itmX.SubItems(4) = rs!Fecha_Corte
                itmX.SubItems(5) = rs!Notas
       
       End With
       rs.MoveNext
     Loop
     rs.Close
     
 End Select

End Sub



Private Sub sbCargaBuscar()
Dim vFechaInicio As Date, vFechaCorte As Date
Dim curTotal As Currency

Me.MousePointer = vbHourglass

On Error GoTo vError

lswCarga.ListItems.Clear
curTotal = 0



strSQL = "select FECHA_INICIO,FECHA_CORTE from CRD_APA_REMESASTESORERIA where REMESA = " & cboCarga.ItemData(cboCarga.ListIndex)
Call OpenRecordSet(rs, strSQL)
  vFechaInicio = rs!FECHA_INICIO
  vFechaCorte = rs!Fecha_Corte
rs.Close


strSQL = "select P.NPAGO,P.DOCUMENTO,P.MONTO,P.FECHA_PAGO,P.OPERACION, A.DESCRIPCION, P.DETALLE_USUARIO, P.COD_ACREEDOR, isnull(AU.NOMBRE,'') as NOMBRE " _
       & " From CRD_APA_PAGOS P inner join CRD_APA_ACREEDORES A on A.COD_ACREEDOR = P.COD_ACREEDOR " _
       & " left join CRD_APA_AUTORIZADOSCK AU on AU.CEDULA = P.CEDULA_AUTORIZADO " _
       & " where P.TESORERIA_REMESA is null" _
       & " and P.FECHA_PAGO between '" & Format(vFechaInicio, "yyyy/mm/dd") & " 00:00:00' and '" & Format(vFechaCorte, "yyyy/mm/dd") _
       & " 23:59:59'"
       
Call OpenRecordSet(rs, strSQL)

PrgBar.Max = rs.RecordCount + 1
PrgBar.Value = 1
PrgBar.Visible = True

Do While Not rs.EOF
 Set itmX = lswCarga.ListItems.Add(, , rs!nPago)
     itmX.SubItems(1) = rs!Documento
     itmX.SubItems(2) = Format(rs!Monto, "Standard")
     itmX.SubItems(3) = rs!Fecha_Pago
     itmX.SubItems(4) = rs!Operacion
     itmX.SubItems(5) = rs!Descripcion
     itmX.SubItems(6) = rs!Nombre
     itmX.SubItems(7) = rs!Detalle_Usuario
     itmX.SubItems(8) = rs!Cod_Acreedor
     
     itmX.Checked = chkCarga.Value
     
     If itmX.Checked Then
        curTotal = curTotal + CCur(itmX.SubItems(2))
     End If
     
 rs.MoveNext
 
 PrgBar.Value = PrgBar.Value + 1
 
Loop
rs.Close

PrgBar.Visible = False

txtCargaTotal.Text = Format(curTotal, "Standard")

Me.MousePointer = vbDefault

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
 lswCarga.ListItems.Clear

End Sub

Private Sub sbCerrar()
Dim i As Integer, vCasos As Integer
Dim vFecha As Date


On Error GoTo vError


'Valida el Estado de la Remesa
strSQL = "select count(*) as Existe from CRD_APA_REMESASTESORERIA" _
       & " where REMESA = " & cboCarga.ItemData(cboCarga.ListIndex) _
       & " and ESTADO in('A','X')"
Call OpenRecordSet(rs, strSQL)
 If rs!Existe = 0 Then
    rs.Close
    MsgBox "La Remesa actual; ya se encuentra cerrada...", vbExclamation
    Exit Sub
 End If
rs.Close

'Actualiza el Estado de la Remesa como cerrada
strSQL = "update CRD_APA_REMESASTESORERIA set ESTADO = 'C'" _
       & " where REMESA = " & cboCarga.ItemData(cboCarga.ListIndex)
Call ConectionExecute(strSQL)

 
Call Bitacora("CERRAR", "Remesa de administración pagarés: " & cboCarga.ItemData(cboCarga.ListIndex))


MsgBox "Remesa Cerrada Satisfactoriamente...", vbInformation
Call sbLimpia

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
 lswCarga.ListItems.Clear

End Sub

Private Sub sbCarga()
Dim i As Integer, vCasos As Integer
Dim vFecha As Date

On Error GoTo vError

'Valida el Estado de la Remesa
strSQL = "select count(*) as Existe from CRD_APA_REMESASTESORERIA" _
       & " where REMESA = " & cboCarga.ItemData(cboCarga.ListIndex) _
       & " and ESTADO in('A','X') "
Call OpenRecordSet(rs, strSQL)
 If rs!Existe = 0 Then
    rs.Close
    MsgBox "La Remesa actual; ya se encuentra cerrada...", vbExclamation
    Exit Sub
 End If
rs.Close



Me.MousePointer = vbHourglass

vFecha = fxFechaServidor

'Calcula los casos a procesar
vCasos = 1
For i = 1 To lswCarga.ListItems.Count
 If lswCarga.ListItems.Item(i).Checked Then
    vCasos = vCasos + 1
 End If
Next i

PrgBar.Max = vCasos
PrgBar.Value = 1
PrgBar.Visible = True


With lswCarga.ListItems

For i = 1 To .Count
 If .Item(i).Checked Then
 
    strSQL = "update CRD_APA_PAGOS set TESORERIA_REMESA = " & cboCarga.ItemData(cboCarga.ListIndex) _
            & " where NPAGO = " & .Item(i).Text _
            & " and OPERACION = " & pc(Trim(.Item(i).SubItems(4))) _
            & " and COD_ACREEDOR = " & pc(Trim(.Item(i).SubItems(8)))
        
     Call ConectionExecute(strSQL)
   
    PrgBar.Value = PrgBar.Value + 1
  End If
Next i
 
If vCasos > 0 Then
    'Actualiza el Estado de la Remesa como cerrada
    strSQL = "update CRD_APA_REMESASTESORERIA set ESTADO = 'X'" _
           & " where REMESA = " & cboCarga.ItemData(cboCarga.ListIndex)
    Call ConectionExecute(strSQL)
    
    Call Bitacora("CARGA", "Remesa de administración pagarés: " & cboCarga.ItemData(cboCarga.ListIndex))
End If

End With

PrgBar.Visible = False

Me.MousePointer = vbDefault


MsgBox "Proceso Realizado Satisfactoriamente...", vbInformation
Call sbCargaBuscar

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
 lswCarga.ListItems.Clear

End Sub


Private Sub tlbCarga_ButtonClick(ByVal Button As MSComctlLib.Button)

If cboCarga.ListCount = 0 Then Exit Sub

Select Case Button.Key
  Case "buscar"
    Call sbCargaBuscar
  
  Case "cargar"
    If lswCarga.ListItems.Count = 0 Then Exit Sub
    Call sbCarga
  
  Case "cerrar"
    Call sbCerrar
End Select

End Sub



Private Sub tlbPago_ButtonClick(ByVal Button As MSComctlLib.Button)
If cboPago.ListCount = 0 Then Exit Sub

Select Case Button.Key
  Case "buscar"
    Call sbPagoBuscar
  
  Case "pago"
    If lswPago.ListItems.Count = 0 Then Exit Sub
    Call sbPago

End Select
End Sub

Private Sub txtRemesa_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn And IsNumeric(txtRemesa) Then
   Call sbConsulta(txtRemesa)
End If
End Sub




Private Sub sbPagoBuscar()
Dim curTotal As Currency

Me.MousePointer = vbHourglass

On Error GoTo vError

lswPago.ListItems.Clear
curTotal = 0

strSQL = "select P.NPAGO,P.DOCUMENTO,P.MONTO,P.FECHA_PAGO,P.OPERACION, A.DESCRIPCION, P.DETALLE_USUARIO, P.COD_ACREEDOR " _
       & " From CRD_APA_PAGOS P inner join CRD_APA_ACREEDORES A on A.COD_ACREEDOR = P.COD_ACREEDOR " _
       & " where P.TESORERIA_FECHA is null and P.TESORERIA_REMESA = " & cboPago.ItemData(cboPago.ListIndex)

Call OpenRecordSet(rs, strSQL)

PrgBar.Max = rs.RecordCount + 1
PrgBar.Value = 1
PrgBar.Visible = True

Do While Not rs.EOF
 Set itmX = lswPago.ListItems.Add(, , rs!nPago)
     itmX.SubItems(1) = rs!Documento
     itmX.SubItems(2) = Format(rs!Monto, "Standard")
     itmX.SubItems(3) = rs!Fecha_Pago
     itmX.SubItems(4) = rs!Operacion
     itmX.SubItems(5) = rs!Descripcion
     itmX.SubItems(6) = rs!Detalle_Usuario
     itmX.SubItems(7) = rs!Cod_Acreedor
     curTotal = curTotal + CCur(itmX.SubItems(2))
 
 rs.MoveNext
 PrgBar.Value = PrgBar.Value + 1
Loop
rs.Close

PrgBar.Visible = False

txtPagoTotal.Text = Format(curTotal, "Standard")

Me.MousePointer = vbDefault

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
 lswPago.ListItems.Clear

End Sub

Private Sub sbPago()
Dim i As Integer, vCasos As Integer
Dim vFecha As Date


On Error GoTo vError


'Valida el Estado de la Remesa
strSQL = "select count(*) as Existe from CRD_APA_REMESASTESORERIA" _
       & " where REMESA = " & cboPago.ItemData(cboPago.ListIndex) _
       & " and ESTADO in('C') "
Call OpenRecordSet(rs, strSQL)
 If rs!Existe = 0 Then
    rs.Close
    MsgBox "La Remesa actual; ya se encuentra en procesada...", vbExclamation
    Exit Sub
 End If
rs.Close

'Actualiza el Estado de la Remesa como Cola de Pago / Al finalizar Revisa si ya fue Totalmente Pagada
'strSQL = "update CRD_APA_REMESASTESORERIA set ESTADO = 'P'" _
'       & " where REMESA = " & cboPago.ItemData(cboPago.ListIndex)
'Call ConectionExecute(strSQL)

Me.MousePointer = vbHourglass

vFecha = fxFechaServidor

PrgBar.Max = lswPago.ListItems.Count + 1
PrgBar.Value = 1
PrgBar.Visible = True


With lswPago.ListItems

For i = 1 To .Count
     
     strSQL = "exec spCRDAPATesoreriaPago " & cboPago.ItemData(cboPago.ListIndex) & ",'" & Trim(.Item(i).SubItems(7)) _
            & "','" & Trim(.Item(i).SubItems(4)) & "'," & Trim(.Item(i).Text) _
            & ",'" & glogon.Usuario & "','" & Format(vFecha, "yyyymmdd hh:mm:ss") & "'"
            
     Call ConectionExecute(strSQL)


'     Call Bitacora("Aplica", "Desembolso de Vivienda a Tesoreria Remesa:" & cboPago.ItemData(cboPago.ListIndex) _
'                    & " IdDesem:" & .Item(i).Text)

    PrgBar.Value = PrgBar.Value + 1
Next i

End With

PrgBar.Visible = False

Me.MousePointer = vbDefault


MsgBox "Proceso Realizado Satisfactoriamente...", vbInformation
Call sbLimpia

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
 lswPago.ListItems.Clear

End Sub



