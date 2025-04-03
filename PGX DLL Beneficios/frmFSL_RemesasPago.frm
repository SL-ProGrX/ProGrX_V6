VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.Ocx"
Begin VB.Form frmFSL_RemesasPago 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "FOSOL: Remesas de Pago"
   ClientHeight    =   7455
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10470
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "frmFSL_RemesasPago.frx":0000
   ScaleHeight     =   7455
   ScaleWidth      =   10470
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ProgressBar PrgBar 
      Align           =   2  'Align Bottom
      Height          =   150
      Left            =   0
      TabIndex        =   0
      Top             =   7305
      Width           =   10470
      _ExtentX        =   18468
      _ExtentY        =   265
      _Version        =   393216
      Appearance      =   0
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   8520
      Top             =   240
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
            Picture         =   "frmFSL_RemesasPago.frx":6852
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFSL_RemesasPago.frx":1D214
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFSL_RemesasPago.frx":33BD6
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFSL_RemesasPago.frx":48D48
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFSL_RemesasPago.frx":5DEBA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgToolbarIcons 
      Left            =   9240
      Top             =   240
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFSL_RemesasPago.frx":7487C
            Key             =   "New"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFSL_RemesasPago.frx":7498E
            Key             =   "Properties"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFSL_RemesasPago.frx":74AA0
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFSL_RemesasPago.frx":74BB2
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFSL_RemesasPago.frx":74CC4
            Key             =   "Undo"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFSL_RemesasPago.frx":74DD6
            Key             =   "Find"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFSL_RemesasPago.frx":74EE8
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFSL_RemesasPago.frx":74FFA
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFSL_RemesasPago.frx":7510C
            Key             =   "Open"
         EndProperty
      EndProperty
   End
   Begin TabDlg.SSTab ssTab 
      Height          =   6255
      Left            =   120
      TabIndex        =   1
      Top             =   960
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
      TabCaption(0)   =   "Remesa"
      TabPicture(0)   =   "frmFSL_RemesasPago.frx":7521E
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label2(8)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label2(2)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label2(3)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label2(4)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label2(20)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label2(19)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label2(6)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label1(5)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label1(6)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Line1(10)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Line1(15)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Line1(2)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Line1(3)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Line1(4)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Line1(1)"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Line1(12)"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "dtpCorte"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "dtpInicio"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "tlb"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "lswRemesas"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "fraReporte"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "txtNotas"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "txtRemesa"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "txtUsuario"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "txtFecha"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "txtEstado"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).ControlCount=   26
      TabCaption(1)   =   "Cargar"
      TabPicture(1)   =   "frmFSL_RemesasPago.frx":7523A
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label2(22)"
      Tab(1).Control(1)=   "Line1(5)"
      Tab(1).Control(2)=   "Label2(21)"
      Tab(1).Control(3)=   "Line1(18)"
      Tab(1).Control(4)=   "Label3(2)"
      Tab(1).Control(5)=   "lswCarga"
      Tab(1).Control(6)=   "tlbCarga"
      Tab(1).Control(7)=   "chkCarga"
      Tab(1).Control(8)=   "cboCarga"
      Tab(1).Control(9)=   "txtCargaTotal"
      Tab(1).ControlCount=   10
      TabCaption(2)   =   "Traslado / Pago"
      TabPicture(2)   =   "frmFSL_RemesasPago.frx":75256
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label2(24)"
      Tab(2).Control(1)=   "Label3(4)"
      Tab(2).Control(2)=   "Label2(16)"
      Tab(2).Control(3)=   "Line1(11)"
      Tab(2).Control(4)=   "tlbTraslado"
      Tab(2).Control(5)=   "lswTraslado"
      Tab(2).Control(6)=   "cboTraslado"
      Tab(2).Control(7)=   "txtPagoTotal"
      Tab(2).ControlCount=   8
      TabCaption(3)   =   "Reportes"
      TabPicture(3)   =   "frmFSL_RemesasPago.frx":75272
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Label16(2)"
      Tab(3).Control(1)=   "Line1(9)"
      Tab(3).Control(2)=   "Label16(4)"
      Tab(3).Control(3)=   "lblRemesa"
      Tab(3).Control(4)=   "lswRep"
      Tab(3).Control(5)=   "cmdReporte"
      Tab(3).Control(6)=   "chkRemesaInd"
      Tab(3).Control(7)=   "txtRepRemesas"
      Tab(3).ControlCount=   8
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
         TabIndex        =   22
         Top             =   5880
         Width           =   2535
      End
      Begin VB.ComboBox cboTraslado 
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
         TabIndex        =   21
         Top             =   600
         Width           =   6975
      End
      Begin VB.TextBox txtEstado 
         Appearance      =   0  'Flat
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   3120
         Locked          =   -1  'True
         TabIndex        =   20
         Top             =   2400
         Width           =   2655
      End
      Begin VB.TextBox txtFecha 
         Appearance      =   0  'Flat
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   3120
         Locked          =   -1  'True
         TabIndex        =   19
         Top             =   2040
         Width           =   2655
      End
      Begin VB.TextBox txtUsuario 
         Appearance      =   0  'Flat
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   3120
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   1680
         Width           =   2655
      End
      Begin VB.TextBox txtRemesa 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   3120
         TabIndex        =   17
         Top             =   480
         Width           =   2655
      End
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
         Top             =   5880
         Width           =   2535
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
         TabIndex        =   15
         Top             =   480
         Width           =   6975
      End
      Begin VB.CheckBox chkCarga 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Todos"
         ForeColor       =   &H00C00000&
         Height          =   210
         Left            =   -74880
         TabIndex        =   14
         Top             =   1455
         Width           =   1455
      End
      Begin VB.TextBox txtRepRemesas 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   -65640
         TabIndex        =   13
         Text            =   "15"
         Top             =   3120
         Width           =   615
      End
      Begin VB.CheckBox chkRemesaInd 
         Appearance      =   0  'Flat
         Caption         =   "Indicar Remesa"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   -66600
         TabIndex        =   12
         Top             =   5280
         Width           =   1695
      End
      Begin VB.CommandButton cmdReporte 
         Caption         =   "&Reporte"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -66600
         TabIndex        =   11
         Top             =   5760
         Width           =   1455
      End
      Begin VB.TextBox txtNotas 
         Appearance      =   0  'Flat
         ForeColor       =   &H00C00000&
         Height          =   675
         Left            =   3120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   10
         Top             =   2880
         Width           =   6975
      End
      Begin VB.Frame fraReporte 
         Caption         =   "Pendientes"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1815
         Left            =   6000
         TabIndex        =   3
         Top             =   960
         Visible         =   0   'False
         Width           =   3975
         Begin VB.CheckBox chkRepFechas 
            Appearance      =   0  'Flat
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   3480
            TabIndex        =   5
            ToolTipText     =   "Todas las Fechas"
            Top             =   480
            Width           =   195
         End
         Begin VB.ComboBox cboRepOficina 
            Height          =   330
            Left            =   960
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   840
            Width           =   2415
         End
         Begin MSComCtl2.DTPicker dtpRepInicio 
            Height          =   315
            Left            =   960
            TabIndex        =   6
            Top             =   480
            Width           =   1215
            _ExtentX        =   2143
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
            Format          =   541786115
            CurrentDate     =   36278
         End
         Begin MSComCtl2.DTPicker dtpRepCorte 
            Height          =   315
            Left            =   2160
            TabIndex        =   7
            Top             =   480
            Width           =   1215
            _ExtentX        =   2143
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
            Format          =   541786115
            CurrentDate     =   36278
         End
         Begin VB.Image imgRepRefresca 
            Height          =   240
            Left            =   3480
            Picture         =   "frmFSL_RemesasPago.frx":7528E
            ToolTipText     =   "Actualizar Oficinas"
            Top             =   840
            Width           =   240
         End
         Begin VB.Label Label7 
            Caption         =   "Fechas"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   9
            Top             =   480
            Width           =   735
         End
         Begin VB.Label Label7 
            Caption         =   "Oficina"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   8
            Top             =   840
            Width           =   735
         End
         Begin VB.Image imgCancelar 
            Height          =   240
            Left            =   3480
            Picture         =   "frmFSL_RemesasPago.frx":753A7
            Top             =   1440
            Width           =   240
         End
         Begin VB.Image imgReporte 
            Height          =   240
            Left            =   3120
            Picture         =   "frmFSL_RemesasPago.frx":754CE
            Top             =   1440
            Width           =   240
         End
      End
      Begin MSComctlLib.ListView lswTraslado 
         Height          =   3975
         Left            =   -74880
         TabIndex        =   2
         Top             =   1800
         Width           =   9975
         _ExtentX        =   17595
         _ExtentY        =   7011
         View            =   3
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         AllowReorder    =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         HotTracking     =   -1  'True
         HoverSelection  =   -1  'True
         _Version        =   393217
         SmallIcons      =   "ImageList1"
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
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Expediente"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   1
            Text            =   "Cédula"
            Object.Width           =   4304
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Nombre"
            Object.Width           =   7832
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Monto"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComctlLib.Toolbar tlbTraslado 
         Height          =   330
         Left            =   -72000
         TabIndex        =   23
         Top             =   1080
         Width           =   2625
         _ExtentX        =   4630
         _ExtentY        =   582
         ButtonWidth     =   2170
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
               Caption         =   "&Traslado"
               Key             =   "traslado"
               Object.ToolTipText     =   "Traslado de Operaciones a Tesoreria"
               ImageIndex      =   3
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ListView lswRemesas 
         Height          =   2535
         Left            =   3120
         TabIndex        =   24
         Top             =   3600
         Width           =   6975
         _ExtentX        =   12303
         _ExtentY        =   4471
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
         NumItems        =   7
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
            Text            =   "Estado"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Inicio"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Corte"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Notas"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComctlLib.Toolbar tlb 
         Height          =   330
         Left            =   6000
         TabIndex        =   25
         Top             =   480
         Width           =   2265
         _ExtentX        =   3995
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         Style           =   1
         ImageList       =   "imgToolbarIcons"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   5
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "nuevo"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "borrar"
               ImageIndex      =   3
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "reportes"
               ImageIndex      =   7
               Style           =   5
               BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                  NumButtonMenus  =   3
                  BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "pendientes"
                     Text            =   "Operaciones Pendientes"
                  EndProperty
                  BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Text            =   "-"
                  EndProperty
                  BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "trasladadas"
                     Text            =   "Operaciones Trasladadas"
                  EndProperty
               EndProperty
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "ayuda"
               ImageIndex      =   8
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.Toolbar tlbCarga 
         Height          =   330
         Left            =   -71880
         TabIndex        =   26
         Top             =   960
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
               ImageIndex      =   5
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ListView lswRep 
         Height          =   2175
         Left            =   -74760
         TabIndex        =   27
         Top             =   840
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
      Begin MSComCtl2.DTPicker dtpInicio 
         Height          =   315
         Left            =   3120
         TabIndex        =   28
         Top             =   1200
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
         Format          =   541786115
         CurrentDate     =   36278
      End
      Begin MSComCtl2.DTPicker dtpCorte 
         Height          =   315
         Left            =   4440
         TabIndex        =   29
         Top             =   1200
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
         Format          =   541786115
         CurrentDate     =   36278
      End
      Begin MSComctlLib.ListView lswCarga 
         Height          =   4095
         Left            =   -74880
         TabIndex        =   30
         Top             =   1680
         Width           =   9975
         _ExtentX        =   17595
         _ExtentY        =   7223
         View            =   3
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         AllowReorder    =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         HotTracking     =   -1  'True
         HoverSelection  =   -1  'True
         _Version        =   393217
         SmallIcons      =   "ImageList1"
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
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Expediente"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Cédula"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   2
            Text            =   "Nombre"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Presenta"
            Object.Width           =   7832
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Monto"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   12
         X1              =   120
         X2              =   3000
         Y1              =   1200
         Y2              =   1200
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   1
         X1              =   120
         X2              =   3000
         Y1              =   720
         Y2              =   720
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
         TabIndex        =   59
         Top             =   3120
         Width           =   5295
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
         TabIndex        =   58
         Top             =   5760
         Width           =   855
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   8
         X1              =   -74880
         X2              =   -72000
         Y1              =   1080
         Y2              =   1080
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Seleccione los promotores o comités para Pago por Tesorería"
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
         Left            =   -74880
         TabIndex        =   57
         Top             =   1800
         Width           =   9975
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   7
         X1              =   -74880
         X2              =   -72000
         Y1              =   720
         Y2              =   720
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Seleccione los promotores o comités a Generar Comisión"
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
         TabIndex        =   56
         Top             =   1320
         Width           =   9975
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   6
         X1              =   -74880
         X2              =   -72000
         Y1              =   720
         Y2              =   720
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   4
         X1              =   120
         X2              =   3000
         Y1              =   2280
         Y2              =   2280
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   3
         X1              =   120
         X2              =   3000
         Y1              =   1920
         Y2              =   1920
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   2
         X1              =   120
         X2              =   3000
         Y1              =   3120
         Y2              =   3120
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   15
         X1              =   120
         X2              =   3000
         Y1              =   3840
         Y2              =   3840
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
         Index           =   10
         Left            =   -74880
         TabIndex        =   55
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
         Index           =   11
         Left            =   -74880
         TabIndex        =   54
         Top             =   480
         Width           =   2895
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Banco"
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
         Index           =   12
         Left            =   -74880
         TabIndex        =   53
         Top             =   840
         Width           =   2895
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Reporte de Comisiones de Afiliación"
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
         Index           =   13
         Left            =   -74760
         TabIndex        =   52
         Top             =   480
         Width           =   3615
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
         Index           =   14
         Left            =   -74760
         TabIndex        =   51
         Top             =   2040
         Width           =   3615
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Promotor / Comité"
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
         Index           =   15
         Left            =   -74760
         TabIndex        =   50
         Top             =   2880
         Width           =   3615
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Banco"
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
         Index           =   17
         Left            =   -74760
         TabIndex        =   49
         Top             =   3720
         Width           =   3615
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
         Index           =   18
         Left            =   -74760
         TabIndex        =   48
         Top             =   4560
         Width           =   3615
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
         Index           =   2
         Left            =   -68280
         TabIndex        =   47
         Top             =   5880
         Width           =   855
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   18
         X1              =   -74880
         X2              =   -72000
         Y1              =   360
         Y2              =   360
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Seleccione los Gastos Pendientes"
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
         Index           =   21
         Left            =   -74880
         TabIndex        =   46
         Top             =   1440
         Width           =   9975
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
         TabIndex        =   45
         Top             =   3120
         Width           =   4935
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   9
         X1              =   -74760
         X2              =   -65040
         Y1              =   5160
         Y2              =   5160
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
         TabIndex        =   44
         Top             =   600
         Width           =   9735
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   11
         X1              =   -74880
         X2              =   -72000
         Y1              =   360
         Y2              =   360
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Lista de Operaciones Pendientes a Trasladar"
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
         Index           =   16
         Left            =   -74880
         TabIndex        =   43
         Top             =   1560
         Width           =   9975
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
         Index           =   4
         Left            =   -68400
         TabIndex        =   42
         Top             =   5880
         Width           =   855
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   10
         X1              =   120
         X2              =   3000
         Y1              =   2640
         Y2              =   2640
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
         TabIndex        =   41
         Top             =   960
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
         TabIndex        =   40
         Top             =   960
         Width           =   1335
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   5
         X1              =   -74880
         X2              =   -72000
         Y1              =   720
         Y2              =   720
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
         Index           =   6
         Left            =   120
         TabIndex        =   39
         Top             =   480
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
         Index           =   19
         Left            =   120
         TabIndex        =   38
         Top             =   2880
         Width           =   2895
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Fechas de Corte"
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
         Index           =   20
         Left            =   120
         TabIndex        =   37
         Top             =   960
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
         TabIndex        =   36
         Top             =   2400
         Width           =   2895
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Registro"
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
         TabIndex        =   35
         Top             =   2040
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
         TabIndex        =   34
         Top             =   1680
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
         Index           =   8
         Left            =   120
         TabIndex        =   33
         Top             =   3600
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
         Index           =   22
         Left            =   -74880
         TabIndex        =   32
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
         Index           =   24
         Left            =   -74880
         TabIndex        =   31
         Top             =   600
         Width           =   2895
      End
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   0
      X1              =   120
      X2              =   10560
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Label Label6 
      Caption         =   "Traspaso a Tesorería para Desembolso"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1200
      TabIndex        =   60
      Top             =   240
      Width           =   7575
   End
End
Attribute VB_Name = "frmFSL_RemesasPago"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim itmX As ListItem, vPaso As Boolean
Dim mRequiereAutorizacion As Boolean
Dim mUnidad As String, mConcepto As String

Private Sub cboCarga_Click()
Dim strSQL As String, rsTmp As New ADODB.Recordset
Dim vFechaInicio As Date, vFechaCorte As Date


On Error GoTo vError

lswCarga.ListItems.Clear
If cboCarga.ListCount <= 0 Then Exit Sub

Me.MousePointer = vbHourglass

strSQL = "select fecha_inicio,fecha_corte from FSL_REMESAS_TESORERIA where TESORERIA_REMESA = " & cboCarga.ItemData(cboCarga.ListIndex)
rsTmp.Open strSQL, glogon.Conection, adOpenStatic
  vFechaInicio = rsTmp!Fecha_Inicio
  vFechaCorte = rsTmp!Fecha_Corte
rsTmp.Close

Me.MousePointer = vbDefault

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical


End Sub


Private Sub sbConsulta(pRemesa As Long)
Dim strSQL As String, rs As New ADODB.Recordset

Call sbLimpia
  
strSQL = "select * from FSL_REMESAS_TESORERIA where TESORERIA_REMESA = " & pRemesa
Call OpenRecordSet(rs, strSQL)
If Not rs.BOF And Not rs.EOF Then
  txtRemesa = rs!TESORERIA_REMESA
  txtUsuario = rs!Registro_Usuario
  txtFecha = rs!registro_Fecha
  
  Select Case rs!Estado
    Case "A"
      txtEstado = "Remesa Abierta"
    Case "C"
      txtEstado = "Remesa Cerrada"
    Case "T"
      txtEstado = "Remesa Trasladada"
  End Select
  
  dtpInicio.Value = rs!Fecha_Inicio
  dtpCorte.Value = rs!Fecha_Corte
  
  txtNotas.Text = rs!notas
  
End If
rs.Close

End Sub




Private Sub chkRepFechas_Click()
If chkRepFechas.Value = vbChecked Then
  dtpRepInicio.Enabled = False
Else
  dtpRepInicio.Enabled = True
End If

dtpRepCorte.Enabled = dtpRepInicio.Enabled

End Sub

Private Sub cmdReporte_Click()
Dim vTitulo As String, vSubTitulo As String, vFiltro As String
Dim vTipoUser As String, xRemesa As String

On Error GoTo vError

If lblRemesa.Tag = "" Then Exit Sub

Me.MousePointer = vbHourglass


vSubTitulo = ""
vFiltro = ""

With frmContenedor.Crt
 .Reset
 .WindowShowGroupTree = True
 .WindowShowPrintSetupBtn = True
 .WindowShowRefreshBtn = True
 .WindowShowSearchBtn = True
 .WindowState = crptMaximized
 .WindowTitle = "Reportes del Módulo de FOSOL"

 .Connect = glogon.ConectRPT

If chkRemesaInd.Value = vbChecked Then
   xRemesa = InputBox("Indique la Remesa que desea consultar", "Remesas de Traslado a Tesoreria")
  If Len(Trim(xRemesa)) = 0 Then xRemesa = "0"
  lblRemesa.Tag = xRemesa
End If
'
' Select Case True
'  Case opt.Item(0).Value 'Pendiente Detalle Remesa
     .ReportFileName = SIFGlobal.fxPathReportes("FSL_RemesaTesDetalle.rpt")
     vSubTitulo = "REMESA : " & lblRemesa.Tag & " LISTADO : DETALLADO"
'  Case opt.Item(1).Value 'Traslado Detalle Agrupado Remesa
'     .ReportFileName = SIFGlobal.fxPathReportes("CxC_RemesaTESDetalleAgrp.rpt")
'     vSubTitulo = "REMESA : " & lblRemesa.Tag & " LISTADO : DETALLADO AGRUPADO"
' End Select
'
 .Formulas(0) = "fxFecha='FECHA: " & Format(fxFechaServidor, "dd/mm/yyyy") & "'"
 .Formulas(1) = "fxEmpresa='" & GLOBALES.gstrNombreEmpresa & "'"
 .Formulas(2) = "fxUsuario='USER: " & glogon.Usuario & "'"
 .Formulas(3) = "fxTitulo='REMESA TRASLADO A TESORERIA : Cbr J.'"
 .Formulas(4) = "fxSubTitulo='" & vSubTitulo & "'"
 .SelectionFormula = "{FSL_REMESAS_TESORERIA.TESORERIA_REMESA} = " & lblRemesa.Tag
 .PrintReport

End With

Me.MousePointer = vbDefault
Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub Form_Activate()
vModulo = 7
End Sub

Private Sub imgCancelar_Click()
fraReporte.Visible = False
End Sub

Private Sub imgReporte_Click()

Select Case fraReporte.Caption
  Case "Pendientes"
    Call sbReportePendientes
  Case "Trasladadas"
    Call sbReporteEnviadas
End Select

End Sub

Private Sub imgRepRefresca_Click()
Dim strSQL As String, rs As New ADODB.Recordset
Dim vFechaInicio As Date, vFechaCorte As Date

Me.MousePointer = vbHourglass

On Error GoTo vError

 
If chkRepFechas.Value = vbChecked Then
  vFechaInicio = "1900/01/01"
  vFechaCorte = fxFechaServidor
Else
  vFechaInicio = dtpRepInicio.Value
  vFechaCorte = dtpRepCorte.Value
End If


Me.MousePointer = vbDefault

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub lswRep_Click()
If lswRep.ListItems.Count <= 0 Then Exit Sub

lblRemesa.Caption = lswRep.SelectedItem.Text & " ¦ " & lswRep.SelectedItem.SubItems(1) _
            & " ¦ " & lswRep.SelectedItem.SubItems(2)
lblRemesa.Tag = lswRep.SelectedItem

End Sub

Private Sub lswTraslado_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo vError
    
    lswTraslado.SortKey = ColumnHeader.Index - 1
    
    If (lswTraslado.SortOrder = lvwAscending) Then
        lswTraslado.SortOrder = lvwDescending
    Else
        lswTraslado.SortOrder = lvwAscending
    End If
    
    lswTraslado.Sorted = True
    Exit Sub

vError:
   MsgBox "Ocurrió un error al ordenar los datos de la columna seleccionada.", vbCritical

End Sub


Private Sub cboTraslado_Click()
    lswTraslado.ListItems.Clear
    txtPagoTotal.Text = 0
End Sub

Private Sub chkCarga_Click()
Dim i As Integer, curTotal As Currency


For i = 1 To lswCarga.ListItems.Count
  
   If lswCarga.ListItems.Item(i).Checked Then
       curTotal = curTotal + CCur(lswCarga.ListItems.Item(i).SubItems(4))
   End If
  
Next i

txtCargaTotal.Text = Format(curTotal, "Standard")

End Sub


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

Private Sub lswCarga_ItemCheck(ByVal Item As MSComctlLib.ListItem)
Dim curTotal As Currency

If Trim(txtCargaTotal.Text) = "" Then txtCargaTotal.Text = 0

curTotal = CCur(txtCargaTotal.Text)

If Item.Checked Then
   curTotal = curTotal + CCur(Item.SubItems(8))
Else
   curTotal = curTotal - CCur(Item.SubItems(8))
End If

txtCargaTotal.Text = Format(curTotal, "Standard")

End Sub

Private Sub lswRemesas_Click()
    If lswRemesas.ListItems.Count <= 0 Then Exit Sub
    Call sbConsulta(lswRemesas.SelectedItem)
End Sub


Private Sub ssTab_Click(PreviousTab As Integer)
 Call sbLimpia
End Sub

Private Sub sbReporteRemesas()
Dim vSubTitulo As String, vFiltro As String
Dim strSQL As String

On Error GoTo vError

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
 .WindowTitle = "Reportes del Módulo de FOSOL"
 
 .Connect = glogon.ConectRPT
  
 
 .Formulas(0) = "fxFecha='FECHA: " & Format(fxFechaServidor, "dd/mm/yyyy") & "'"
 .Formulas(1) = "fxEmpresa='" & GLOBALES.gstrNombreEmpresa & "'"
 .Formulas(2) = "fxUsuario='USER: " & glogon.Usuario & "'"
 
 .ReportFileName = SIFGlobal.fxPathReportes("CBR_CJ_ListadoRemesa.rpt")
 .PrintReport

End With

Me.MousePointer = vbDefault
Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub tlb_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim strSQL As String, rs As New ADODB.Recordset
Dim i As Integer

On Error GoTo vError

Select Case UCase(Button.Key)
  Case "NUEVO"
     
    If txtRemesa.Text = "" Then
     
            strSQL = "select coalesce(max(TESORERIA_REMESA),0) + 1 as Ultimo from FSL_REMESAS_TESORERIA"
            Call OpenRecordSet(rs, strSQL)
                strSQL = "insert FSL_REMESAS_TESORERIA(TESORERIA_REMESA,registro_usuario,registro_fecha,estado,fecha_inicio,fecha_corte,notas) values(" & rs!Ultimo _
                       & ",'" & glogon.Usuario & "',getdate(),'A','" & Format(dtpInicio.Value, "yyyy/mm/dd") _
                       & "','" & Format(dtpCorte.Value, "yyyy/mm/dd") & "','" & txtNotas.Text & "')"
                Call ConectionExecute(strSQL)
                
                txtRemesa = rs!Ultimo
            rs.Close
            Call Bitacora("Registra", "Remesa FOSOL Traslado a Tesoreria : " & txtRemesa)
    
    Else
        If txtEstado.Text <> "Remesa Cerrada" Then
                    
            strSQL = "update FSL_REMESAS_TESORERIA set registro_usuario = '" & glogon.Usuario & "',fecha_inicio = '" & Format(dtpInicio.Value, "yyyy/mm/dd") _
                   & "',fecha_corte = '" & Format(dtpCorte.Value, "yyyy/mm/dd") & "',notas = '" & txtNotas.Text _
                   & "' where TESORERIA_REMESA = " & txtRemesa
             Call ConectionExecute(strSQL)
             
            Call Bitacora("Modifica", "Remesa FOSOL Traslado a Tesoreria : " & txtRemesa)
        Else
            MsgBox "No se puede Modifica la remesa, porque esta ya fue cerrada...", vbExclamation
        End If
    End If
    
    Call sbLimpia
    
    
  Case "BORRAR"
     i = MsgBox("Esta Seguro que desea borrar este registro", vbYesNo)
     If i = vbYes Then
        
        If txtEstado.Text = "Remesa Abierta" Then
            
            strSQL = "delete FSL_REMESAS_TESORERIA where TESORERIA_REMESA = " & txtRemesa.Text
            Call ConectionExecute(strSQL)
            
            
            Call Bitacora("Elimina", "Remesa FOSOL Traslado a Tesoreria : " & txtRemesa)
         End If
       
        Call sbLimpia
     End If
  
  Case "REPORTES"
     
         Call sbReporteRemesas
     
  Case "AYUDA"
'        frmContenedor.CD.HelpContext = Me.HelpContextID
'        frmContenedor.CD.ShowHelp

End Select

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub sbLimpia()
Dim strSQL As String, rs As New ADODB.Recordset

Me.MousePointer = vbHourglass

Select Case ssTab.Tab
  Case 0 'Remesas
     txtEstado = ""
     txtFecha = ""
     txtUsuario = ""
     txtRemesa = ""
     
    dtpInicio.Value = fxFechaServidor
    dtpCorte.Value = dtpInicio.Value
    
    dtpRepInicio.Value = dtpInicio.Value
    dtpRepCorte.Value = dtpInicio.Value
    
    txtNotas.Text = ""
     
     strSQL = "select TOP 150 * from FSL_REMESAS_TESORERIA order by registro_fecha desc"
     lswRemesas.ListItems.Clear
     Call OpenRecordSet(rs, strSQL)
     Do While Not rs.EOF
       With lswRemesas.ListItems
            Set itmX = .Add(, , rs!TESORERIA_REMESA)
                itmX.SubItems(1) = rs!Registro_Usuario
                itmX.SubItems(2) = rs!registro_Fecha
                
                Select Case rs!Estado
                  Case "A"
                     itmX.SubItems(3) = "Remesa Abierta"
                  Case "C"
                     itmX.SubItems(3) = "Remesa Cerrada"
                  Case "T"
                     itmX.SubItems(3) = "Remesa Trasladada"
                End Select
                
                itmX.SubItems(4) = Format(rs!Fecha_Inicio, "dd/mm/yyyy")
                itmX.SubItems(5) = Format(rs!Fecha_Corte, "dd/mm/yyyy")
                itmX.SubItems(6) = rs!notas
                
                
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
        
    strSQL = "select * from FSL_REMESAS_TESORERIA where estado = 'A' order by registro_fecha desc"
    Call OpenRecordSet(rs, strSQL)
    Do While Not rs.EOF
      cboCarga.AddItem (Format(rs!TESORERIA_REMESA, "0000") & "..." & Trim(rs!Registro_Usuario) & "..." _
            & rs!registro_Fecha & " I:" & Format(rs!Fecha_Inicio, "dd/mm/yyyy") & " C:" & Format(rs!Fecha_Corte, "dd/mm/yyyy"))
      cboCarga.ItemData(cboCarga.NewIndex) = rs!TESORERIA_REMESA
      rs.MoveNext
    Loop
    If rs.RecordCount > 0 Then
       rs.MoveFirst
       cboCarga.Text = (Format(rs!TESORERIA_REMESA, "0000") & "..." & Trim(rs!Registro_Usuario) & "..." _
            & rs!registro_Fecha & " I:" & Format(rs!Fecha_Inicio, "dd/mm/yyyy") & " C:" & Format(rs!Fecha_Corte, "dd/mm/yyyy"))
    End If
    
    rs.Close

    vPaso = False
    Call cboCarga_Click
   
    
  Case 2 'Traslado
    vPaso = True
    
    cboTraslado.Clear

    lswTraslado.ListItems.Clear
    txtPagoTotal.Text = 0
        
        
    strSQL = "select * from FSL_REMESAS_TESORERIA where estado = 'C' order by REGISTRO_fecha desc"
    Call OpenRecordSet(rs, strSQL)
    Do While Not rs.EOF
      cboTraslado.AddItem (Format(rs!TESORERIA_REMESA, "0000") & "..." & Trim(rs!Registro_Usuario) & "..." _
            & rs!registro_Fecha & " I:" & Format(rs!Fecha_Inicio, "dd/mm/yyyy") & " C:" & Format(rs!Fecha_Corte, "dd/mm/yyyy"))
      cboTraslado.ItemData(cboTraslado.NewIndex) = rs!TESORERIA_REMESA
      rs.MoveNext
    Loop
    If rs.RecordCount > 0 Then
       rs.MoveFirst
       cboTraslado.Text = (Format(rs!TESORERIA_REMESA, "0000") & "..." & Trim(rs!Registro_Usuario) & "..." _
            & rs!registro_Fecha & " I:" & Format(rs!Fecha_Inicio, "dd/mm/yyyy") & " C:" & Format(rs!Fecha_Corte, "dd/mm/yyyy"))
    End If
    
    rs.Close

    vPaso = False
    Call cboTraslado_Click

  
  Case 3 'Reportes
     strSQL = "select TOP " & txtRepRemesas.Text & " * from FSL_REMESAS_TESORERIA order by registro_fecha desc"
     lswRep.ListItems.Clear

     Call OpenRecordSet(rs, strSQL)
     Do While Not rs.EOF
       With lswRep.ListItems
            Set itmX = .Add(, , rs!TESORERIA_REMESA)
                itmX.SubItems(1) = rs!Registro_Usuario
                itmX.SubItems(2) = rs!registro_Fecha
                itmX.SubItems(3) = rs!Fecha_Inicio
                itmX.SubItems(4) = rs!Fecha_Corte
                itmX.SubItems(5) = rs!notas
       
       End With
       rs.MoveNext
     Loop
     rs.Close

    
End Select


Me.MousePointer = vbDefault

End Sub




Private Sub tlb_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)


Select Case ButtonMenu.Key
  Case "pendientes"
'    Call sbReportePendientes
     fraReporte.Visible = True
     fraReporte.Caption = "Pendientes"
  Case "trasladadas"
'    Call sbReporteEnviadas
     fraReporte.Visible = True
     fraReporte.Caption = "Trasladadas"
End Select

End Sub

Private Sub tlbTraslado_ButtonClick(ByVal Button As MSComctlLib.Button)

If cboTraslado.ListCount = 0 Then Exit Sub

Select Case Button.Key
  Case "buscar"
    Call sbTrasladoBuscar
  
  Case "traslado"
    Call sbTraslado

End Select

End Sub


Private Sub sbCargaBuscar()
Dim strSQL As String, rs As New ADODB.Recordset
Dim vFechaInicio As Date, vFechaCorte As Date
Dim curTotal As Currency

Me.MousePointer = vbHourglass

On Error GoTo vError

lswCarga.ListItems.Clear
curTotal = 0


strSQL = "select fecha_inicio,fecha_corte from FSL_REMESAS_TESORERIA where TESORERIA_REMESA = " & cboCarga.ItemData(cboCarga.ListIndex)
Call OpenRecordSet(rs, strSQL)
  vFechaInicio = rs!Fecha_Inicio
  vFechaCorte = rs!Fecha_Corte
rs.Close


strSQL = "Select E.COD_EXPEDIENTE,E.CEDULA,S.NOMBRE, E.TOTAL_SOBRANTE, E.PRESENTA_CEDULA, E.PRESENTA_NOMBRE" _
        & " from FSL_EXPEDIENTES E inner join SOCIOS S on E.CEDULA = S.CEDULA" _
        & " Where E.RESOLUCION_FECHA between '" & Format(vFechaInicio, "yyyy/mm/dd") & " 00:00:00'" _
        & " and '" & Format(vFechaCorte, "yyyy/mm/dd") & " 23:59:59' and E.TESORERIA_REMESA is null " _
        & " and E.Tipo_Desembolso = 'T' and E.Estado = 'X' and E.TOTAL_SOBRANTE > 0" _
        & " order by E.CEDULA,S.NOMBRE"

rs.CursorLocation = adUseServer
Call OpenRecordSet(rs, strSQL)

PrgBar.Max = rs.RecordCount + 1
PrgBar.Value = 1
PrgBar.Visible = True

With lswCarga
 .ListItems.Clear
 Do While Not rs.EOF
   Set itmX = .ListItems.Add(, , rs!COD_EXPEDIENTE)
       itmX.SubItems(1) = rs!Cedula
       itmX.SubItems(2) = rs!Nombre
       itmX.SubItems(3) = rs!PRESENTA_NOMBRE
       itmX.SubItems(4) = Format(rs!TOTAL_SOBRANTE, "Standard")
       
       itmX.Checked = vbChecked
       If itmX.Checked Then
            curTotal = curTotal + CCur(itmX.SubItems(4))
       End If
        
        rs.MoveNext
        
        PrgBar.Value = PrgBar.Value + 1
 Loop
End With

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


Private Sub sbTrasladoBuscar()
Dim strSQL As String, rs As New ADODB.Recordset
Dim vFechaInicio As Date, vFechaCorte As Date
Dim curTotal As Currency

Me.MousePointer = vbHourglass

On Error GoTo vError

lswTraslado.ListItems.Clear
curTotal = 0


strSQL = "select fecha_inicio,fecha_corte from FSL_REMESAS_TESORERIA where TESORERIA_REMESA = " & cboTraslado.ItemData(cboTraslado.ListIndex)
Call OpenRecordSet(rs, strSQL)
  vFechaInicio = rs!Fecha_Inicio
  vFechaCorte = rs!Fecha_Corte
rs.Close

strSQL = "Select E.COD_EXPEDIENTE,E.CEDULA,S.NOMBRE, E.TOTAL_SOBRANTE, E.PRESENTA_CEDULA, E.PRESENTA_NOMBRE" _
        & " from FSL_EXPEDIENTES E inner join SOCIOS S on E.CEDULA = S.CEDULA" _
        & " Where E.RESOLUCION_FECHA between '" & Format(vFechaInicio, "yyyy/mm/dd") & " 00:00:00'" _
        & " and '" & Format(vFechaCorte, "yyyy/mm/dd") & " 23:59:59' and E.TESORERIA_REMESA = " & cboTraslado.ItemData(cboTraslado.ListIndex) _
        & " and E.Tipo_Desembolso = 'T' and E.Estado = 'X' and E.TOTAL_SOBRANTE > 0 and isnull(E.Tesoreria_Solicitud,0) = 0" _
        & " order by E.CEDULA,S.NOMBRE"

rs.CursorLocation = adUseServer
Call OpenRecordSet(rs, strSQL)

PrgBar.Max = rs.RecordCount + 1
PrgBar.Value = 1
PrgBar.Visible = True

With lswTraslado
 .ListItems.Clear
 Do While Not rs.EOF
Set itmX = .ListItems.Add(, , rs!COD_EXPEDIENTE)
       itmX.SubItems(1) = rs!Cedula
       itmX.SubItems(2) = rs!Nombre
       itmX.SubItems(3) = Format(rs!TOTAL_SOBRANTE, "Standard")
       
       itmX.Checked = vbChecked
       If itmX.Checked Then
            curTotal = curTotal + CCur(itmX.SubItems(3))
       End If
       
       rs.MoveNext
       PrgBar.Value = PrgBar.Value + 1
 Loop

End With

rs.Close

PrgBar.Visible = False

txtPagoTotal.Text = Format(curTotal, "Standard")

Me.MousePointer = vbDefault

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
 lswTraslado.ListItems.Clear

End Sub



Private Sub sbCerrar()
Dim strSQL As String, rs As New ADODB.Recordset
Dim i As Integer, vCasos As Integer
Dim vFecha As Date

On Error GoTo vError


'Valida el Estado de la Remesa
strSQL = "select count(*) as Existe from FSL_REMESAS_TESORERIA" _
       & " where TESORERIA_REMESA = " & cboCarga.ItemData(cboCarga.ListIndex) _
       & " and estado = 'A'"
Call OpenRecordSet(rs, strSQL)
 If rs!Existe = 0 Then
    rs.Close
    MsgBox "La Remesa actual; ya se encuentra cerrada...", vbExclamation
    Exit Sub
 End If
rs.Close

'Actualiza el Estado de la Remesa como cerrada
strSQL = "update FSL_REMESAS_TESORERIA set estado = 'C'" _
       & " where TESORERIA_REMESA = " & cboCarga.ItemData(cboCarga.ListIndex)
Call ConectionExecute(strSQL)

 
Call Bitacora("Aplica", "Cierra Remesa Traslado a Tesoreria : " & cboCarga.ItemData(cboCarga.ListIndex))


MsgBox "Remesa Cerrada Satisfactoriamente...", vbInformation
Call sbLimpia

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
 lswCarga.ListItems.Clear

End Sub

Private Sub sbCarga()
Dim strSQL As String, rs As New ADODB.Recordset
Dim i As Integer, vCasos As Integer
Dim vFecha As Date

On Error GoTo vError


'Valida el Estado de la Remesa
strSQL = "select count(*) as Existe from FSL_REMESAS_TESORERIA" _
       & " where TESORERIA_REMESA = " & cboCarga.ItemData(cboCarga.ListIndex) _
       & " and estado = 'A'"
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
 
     strSQL = "update FSL_EXPEDIENTES set TESORERIA_REMESA = " & cboCarga.ItemData(cboCarga.ListIndex) _
            & " where COD_EXPEDIENTE = " & .Item(i).Text
     Call ConectionExecute(strSQL)
   
    PrgBar.Value = PrgBar.Value + 1
  End If
Next i
 
If vCasos > 0 Then
    Call Bitacora("Aplica", "Carga Remesa Traslado a Tesoreria : " & cboCarga.ItemData(cboCarga.ListIndex))
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



Private Sub txtRemesa_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn And IsNumeric(txtRemesa) Then
   Call sbConsulta(txtRemesa)
End If

End Sub


Private Function fxMaestroTesoreria(vTipoDocumento As String, vBanco As Integer, vMonto As Currency, vCodigo As String _
                              , vBeneficiario As String, vOP As Long, vDetalle1 As String, vReferencia As Long _
                              , vDetalle2 As String, vCuenta As String, vFecha As Date, vUnidad As String _
                              , vToken As String) As Long                                  'Regresa el NSOLICITUD
Dim rsX As New ADODB.Recordset, strSQL As String
Dim lngSol As Long

strSQL = "insert Tes_Transacciones(cod_concepto,cod_unidad,id_banco,tipo,codigo,beneficiario,monto,fecha_solicitud,estado,estadoi" _
       & ",modulo,submodulo,cta_ahorros,detalle1,detalle2,referencia,op,genera,actualiza,user_solicita,autoriza,user_autoriza" _
       & ",fecha_autorizacion,ID_TOKEN ,REMESA_TIPO, REMESA_ID)" _
       & " values('" & mConcepto & "','" & vUnidad & "'," & vBanco & ",'" & vTipoDocumento & "','" & vCodigo & "','" & vBeneficiario & "'," & vMonto _
       & ",'" & Format(vFecha, "yyyy/mm/dd") & "','P','P','CC','C','" & vCuenta _
       & "','" & vDetalle1 & "','" & vDetalle2 & "'," & vReferencia & "," & vOP & ",'S','S','" & glogon.Usuario & "'"
       
If UCase(vTipoDocumento) = "CK" Then
   strSQL = strSQL & ",'S','" & glogon.Usuario & "',getdate(),'" & vToken & "','FSL'," & cboTraslado.ItemData(cboTraslado.ListIndex) & ")"
Else
   strSQL = strSQL & ",'N',null,null,'" & vToken & "','FSL'," & cboTraslado.ItemData(cboTraslado.ListIndex) & ")"
End If
Call ConectionExecute(strSQL)

strSQL = "select max(nsolicitud) as Solicitud from Tes_Transacciones"
rsX.Open strSQL, glogon.Conection, adOpenStatic
 strSQL = "select * from Tes_Transacciones where nsolicitud = " & rsX!solicitud
rsX.Close

lngSol = 0

rsX.Open strSQL, glogon.Conection, adOpenStatic
If Trim(rsX!Codigo) = Trim(vCodigo) Then lngSol = rsX!NSolicitud
rsX.Close

If lngSol = 0 Then
  strSQL = "select max(nsolicitud) as Solicitud from Tes_Transacciones where codigo ='" & vCodigo _
         & "'"
  rsX.CursorLocation = adUseServer
  rsX.Open strSQL, glogon.Conection, adOpenStatic
  lngSol = rsX!solicitud
  rsX.Close
End If

fxMaestroTesoreria = lngSol

End Function



Private Sub sbCreaDetalle(vSolicitud As Long, vCtaConta As String, vMonto As Currency, vDH As String, vLinea As Integer, vUnidad As String)
Dim strSQL As String

strSQL = "insert Tes_Trans_Asiento(nsolicitud,cuenta_contable,monto,debehaber,linea,cod_unidad) values(" _
       & vSolicitud & ",'" & Trim(vCtaConta) & "'," & vMonto & ",'" & vDH _
       & "'," & vLinea & ",'" & vUnidad & "')"
Call ConectionExecute(strSQL)

End Sub

Private Function fxCtaBanco(pBanco As Integer) As String
Dim rsX As New ADODB.Recordset, strSQL As String

strSQL = "select CTACONTA from Tes_Bancos where id_banco =" & pBanco
rsX.Open strSQL, glogon.Conection, adOpenStatic
If rsX.EOF And rsX.BOF Then
 fxCtaBanco = ""
Else
 fxCtaBanco = rsX!ctaConta
End If
rsX.Close
End Function


Private Function fxCtaPuente(pCodigo As String) As String
Dim rsX As New ADODB.Recordset, strSQL As String

strSQL = "select Cod_Cuenta_Salida from CxC_Conceptos where cod_concepto ='" & pCodigo & "'"
rsX.Open strSQL, glogon.Conection, adOpenStatic
If rsX.EOF And rsX.BOF Then
     fxCtaPuente = ""
Else
     fxCtaPuente = rsX!cod_Cuenta_Salida
End If

rsX.Close

End Function


Private Sub sbCreaDesembolsos(vReferencia As Long, vOP As Long, vFecha As Date, vTipo As String, vBanco As Integer)
Dim rsTemp As New ADODB.Recordset, lngSolicitud As Long
'
'strSQL = "select * from desembolsos where retener = 0 and Operacion = " & vOP
'
'With rsTemp
' .CursorLocation = adUseServer
' .Open strSQL, glogon.Conection, adOpenStatic
' Do While Not .EOF
'     lngSolicitud = fxMaestroTesoreria(vTipo, vBanco, !Monto, !id_desembolso _
'                   , !Concepto, !Operacion, !Operacion, vReferencia, !cod_concepto, "0", vFecha)
'     Call sbCreaDetalle(lngSolicitud, fxCtaBanco(vBanco), !Monto, "H", 1)
'     Call sbCreaDetalle(lngSolicitud, !cuenta_conta, !Monto, "D", 2)
'
'     strSQL = "update desembolsos set tdocumento = '" & vTipo & "',Emitir_Tipo_Banco = " & vBanco & ",nsolicitud = " & lngSolicitud _
'            & " where id_desembolso = " & !id_desembolso
'     Call ConectionExecute(strSQL)
'  .MoveNext
' Loop
' .Close
'End With

End Sub

Private Sub sbTraslado()
Dim strSQL As String, rs As New ADODB.Recordset
Dim lngSolicitud As Long, vFecha As Date
Dim vTipo As String, vBanco As Integer
Dim i As Integer, vCasos As Integer
Dim vCuenta As String, vCuentaAhorros As String
Dim vToken As String

Me.MousePointer = vbHourglass

On Error GoTo vError

vCasos = 0
vFecha = fxFechaServidor

vCuenta = fxFSL_Parametros("01")

mConcepto = fxFSL_Parametros("05")
mUnidad = fxFSL_Parametros("07")


strSQL = "select top 1 id_token from tes_tokens where estado = 'A' order by registro_fecha "
Call OpenRecordSet(rs, strSQL)
If Not rs.EOF Then
  vToken = rs!ID_TOKEN
Else
  vToken = fxTesToken
End If
rs.Close

With lswTraslado.ListItems

For i = 1 To .Count
 If .Item(i).Checked Then
    
    strSQL = "select Top 1 * from cuentas_Ahorros where Tipo =  1 and cedula = '" & .Item(i).SubItems(1) _
            & "' order by Prioridad"
    Call OpenRecordSet(rs, strSQL)
    If Not rs.BOF And Not rs.EOF Then
      vTipo = "TE"
      vCuentaAhorros = Trim(rs!Cuenta)
      vBanco = rs!id_Banco
    Else
      vBanco = fxFSL_Parametros("04")
      vTipo = "CK"
      vCuentaAhorros = ""
    End If
    rs.Close
    
    lngSolicitud = fxMaestroTesoreria(vTipo, vBanco, .Item(i).SubItems(3), .Item(i).SubItems(1) _
                   , .Item(i).SubItems(2), 0, "FOSOL", 0 _
                   , "Exp.: " & .Item(i).Text, vCuentaAhorros, vFecha, mUnidad, vToken)
                   
    'Mata el Pasivo de la Nota de Debito de la Formalizacion contra Tes_Bancos
    Call sbCreaDetalle(lngSolicitud, fxTraeCuentaBanco(vBanco), .Item(i).SubItems(3), "H", 1, mUnidad)
    Call sbCreaDetalle(lngSolicitud, vCuenta, .Item(i).SubItems(3), "D", 2, mUnidad)

    'Actualiza Campo Tesoreria
    strSQL = "update FSL_EXPEDIENTES set Tesoreria_Solicitud = " & lngSolicitud & ", Tesoreria_Fecha = getdate()" _
           & ", Tesoreria_Usuario = '" & glogon.Usuario & "'" _
           & " where TESORERIA_REMESA = " & cboTraslado.ItemData(cboTraslado.ListIndex) _
           & "  and cod_expediente = " & .Item(i).Text
    Call ConectionExecute(strSQL)
 
    'Actualiza Bitacora
    Call Bitacora("Registra", "Traspaso a Tesoreria - Expediente:" & .Item(i).Text)
 
   
   ' PrgBar.Value = PrgBar.Value + 1
    vCasos = vCasos + 1
  End If
Next i
 
If vCasos > 0 Then
    Call Bitacora("Aplica", "Carga Remesa Traslado a Tesoreria : " & cboTraslado.ItemData(cboTraslado.ListIndex))
    'Actualiza y Carga Remesa
    strSQL = "update FSL_REMESAS_TESORERIA SET Estado = 'T'" _
           & "  Where TESORERIA_REMESA = " & cboTraslado.ItemData(cboTraslado.ListIndex)
End If

End With


Call ConectionExecute(strSQL)

Call sbLimpia


Me.MousePointer = vbDefault

PrgBar.Visible = False

MsgBox "Traslado a Tesoreria realizado satisfactoriamente...", vbInformation

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub


Private Sub sbReportePendientes()
Dim strSQL As String, rs As New ADODB.Recordset
Dim strTitulo As String
Dim strRuta As String, strInicio As String, strFinal As String
Dim strFiltro As String

On Error GoTo vError

Me.MousePointer = vbHourglass

strTitulo = "Honorarios pendientes de Traslado a Tesorería"


strRuta = SIFGlobal.fxPathReportes("CBR_CJ_GastoPenEnviar.rpt")
strInicio = "Date(" & Format(dtpInicio.Value, "yyyy,mm,dd") & ")"
strFinal = "Date(" & Format(dtpCorte.Value, "yyyy,mm,dd") & ")"

With frmContenedor.Crt
     .Reset
     .WindowShowGroupTree = True
     .WindowShowPrintSetupBtn = True
     .WindowShowRefreshBtn = True
     .WindowShowSearchBtn = True
     .WindowState = crptMaximized
     
     .Connect = glogon.ConectRPT
     
     .WindowTitle = "Honorarios a trasladar a Tesorería"
     
    .ReportFileName = strRuta
    .Formulas(1) = "empresa='" & GLOBALES.gstrNombreEmpresa & "'"
    .Formulas(2) = "fecha='" & Format(fxFechaServidor, "dd/mm/yyyy") & "'"
    .Formulas(3) = "Titulo='" & strTitulo & "'"
    
  
    If chkRepFechas.Value = vbUnchecked Then
        strSQL = "  cdate({FSL_EXPEDIENTES.Registro_Fecha}) in Date(" & Format(dtpRepInicio.Value, "yyyy,mm,dd")
        strSQL = strSQL & ") to Date (" & Format(dtpRepCorte.Value, "yyyy,mm,dd") & ")"
        strFiltro = "Desde " & Format(dtpRepInicio.Value, "dd/mm/yyyy") & " Hasta " & Format(dtpRepCorte.Value, "dd/mm/yyyy")
    Else
         strFiltro = "Todas las fechas "
    End If
    
    
    If cboRepOficina.Text <> "TODOS" Then
       strSQL = strSQL & " AND {REG_CREDITOS.COD_OFICINA_F} = '" & SIFGlobal.fxCodText(cboRepOficina.Text) & "'"
       
       strFiltro = strFiltro & " /OFICINA " & SIFGlobal.fxCodText(cboRepOficina.Text)
    Else
       strFiltro = strFiltro & "Todas las Oficinas"
    End If
    
    If strSQL = "" Then
      strSQL = "ISNULL({FSL_EXPEDIENTES.TESORERIA_NUMERO})"
    Else
      strSQL = strSQL & " AND ISNULL({FSL_EXPEDIENTES.TESORERIA_NUMERO})"
    End If
    .Formulas(4) = "Filtro='" & strFiltro & "'"
    
    
    .SelectionFormula = strSQL
    
    '.SubreportToChange = "subCkDesembolsos"
    '.SelectionFormula = "{DESEMBOLSOS.Operacion} = {?Pm-CxC_Cuentas.Operacion}"
    
    .PrintReport


End With

 Me.MousePointer = vbDefault

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub sbReporteEnviadas()
Dim strSQL As String, rs As New ADODB.Recordset
Dim strFiltro As String

On Error GoTo vError

Me.MousePointer = vbHourglass


With frmContenedor.Crt
 .Reset
 .WindowShowGroupTree = True
 .WindowShowPrintSetupBtn = True
 .WindowShowRefreshBtn = True
 .WindowShowSearchBtn = True
 .WindowState = crptMaximized
 .WindowTitle = "HONORARIOS ENVIADOS A TESORERIA"

 .Connect = glogon.ConectRPT

 .ReportFileName = SIFGlobal.fxPathReportes("CBR_CJ_GastoTrasladadas.rpt")
 .Formulas(1) = "Empresa='" & GLOBALES.gstrNombreEmpresa & "'"
 .Formulas(2) = "Fecha='" & Format(fxFechaServidor, "dd/mm/yyyy") & "'"
 .Formulas(3) = "Titulo='Desembolsos Solicitados en Tesorería'"
 .Formulas(4) = "Usuario='" & glogon.Usuario & "'"
 strFiltro = "INICIO : " & Format(dtpRepInicio.Value, "dd/mm/yyyy") & " CORTE : " & Format(dtpRepCorte.Value, "dd/mm/yyyy")
 
 strSQL = "{FSL_EXPEDIENTES.tesoreria_fecha} in date(" & Format(dtpRepInicio.Value, "yyyy,mm,dd") & ") to date(" _
       & Format(dtpRepCorte.Value, "yyyy,mm,dd") & ")"
    
 If cboRepOficina.Text <> "TODOS" Then
    strSQL = strSQL & " AND {REG_CREDITOS.COD_OFICINA_F} = '" & SIFGlobal.fxCodText(cboRepOficina.Text) & "'"
    strFiltro = strFiltro & " /OFICINA " & SIFGlobal.fxCodText(cboRepOficina.Text)
 Else
    strFiltro = strFiltro & "Todas las Oficinas"
 End If

 .Formulas(5) = "filtro='" & strFiltro & "'"

 .SelectionFormula = strSQL
 .Action = 1

End With

Me.MousePointer = vbDefault

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub



Private Sub Form_Load()
Dim strSQL As String

vModulo = 7

 ssTab.Tab = 0
 
 Call Formularios(Me)
 Call RefrescaTags(Me)
 
 Call sbLimpia
 
 
End Sub


Private Function fxTraeCuentaBanco(vBanco As Integer) As String
Dim strSQL As String, rs As New ADODB.Recordset


strSQL = "select ctaconta from tes_bancos where id_banco = " & vBanco & " "
Call OpenRecordSet(rs, strSQL)

If Not rs.EOF Then
  fxTraeCuentaBanco = rs!ctaConta
Else
  fxTraeCuentaBanco = "0"
End If
rs.Close

End Function



