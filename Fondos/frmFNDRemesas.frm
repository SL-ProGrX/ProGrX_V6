VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.Ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmFNDRemesas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Planes : Remesas"
   ClientHeight    =   7428
   ClientLeft      =   48
   ClientTop       =   432
   ClientWidth     =   10464
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   Picture         =   "frmFNDRemesas.frx":0000
   ScaleHeight     =   7428
   ScaleWidth      =   10464
   Begin TabDlg.SSTab ssTab 
      Height          =   6255
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   10215
      _ExtentX        =   18013
      _ExtentY        =   11028
      _Version        =   393216
      Style           =   1
      Tabs            =   4
      Tab             =   1
      TabsPerRow      =   4
      TabHeight       =   520
      ForeColor       =   16711680
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Remesas"
      TabPicture(0)   =   "frmFNDRemesas.frx":6852
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "txtNotas"
      Tab(0).Control(1)=   "txtRemesa"
      Tab(0).Control(2)=   "txtUsuario"
      Tab(0).Control(3)=   "txtFecha"
      Tab(0).Control(4)=   "txtEstado"
      Tab(0).Control(5)=   "txtTotal"
      Tab(0).Control(6)=   "dtpInicio"
      Tab(0).Control(7)=   "dtpCorte"
      Tab(0).Control(8)=   "lswRemesas"
      Tab(0).Control(9)=   "tlb"
      Tab(0).Control(10)=   "Line1(10)"
      Tab(0).Control(11)=   "Line2"
      Tab(0).Control(12)=   "Line1(7)"
      Tab(0).Control(13)=   "Line1(6)"
      Tab(0).Control(14)=   "Line1(0)"
      Tab(0).Control(15)=   "Line1(2)"
      Tab(0).Control(16)=   "Line1(3)"
      Tab(0).Control(17)=   "Line1(4)"
      Tab(0).Control(18)=   "Line1(5)"
      Tab(0).Control(19)=   "Label1(6)"
      Tab(0).Control(20)=   "Label1(5)"
      Tab(0).Control(21)=   "Label2(0)"
      Tab(0).Control(22)=   "Label2(5)"
      Tab(0).Control(23)=   "Label2(4)"
      Tab(0).Control(24)=   "Label2(3)"
      Tab(0).Control(25)=   "Label2(2)"
      Tab(0).Control(26)=   "Label2(6)"
      Tab(0).Control(27)=   "Label2(1)"
      Tab(0).Control(28)=   "Label2(9)"
      Tab(0).ControlCount=   29
      TabCaption(1)   =   "Cargado"
      TabPicture(1)   =   "frmFNDRemesas.frx":686E
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label2(10)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label2(8)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label2(7)"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Line1(8)"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Line1(9)"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Label3"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "tlbCarga"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "lswCarga"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "chkCarga"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "cboCarga"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "cboBanco"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "txtCargaTotal"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).ControlCount=   12
      TabCaption(2)   =   "Reportes"
      TabPicture(2)   =   "frmFNDRemesas.frx":688A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "cmdReporte"
      Tab(2).Control(1)=   "txtRepRemesas"
      Tab(2).Control(2)=   "opt(0)"
      Tab(2).Control(3)=   "opt(1)"
      Tab(2).Control(4)=   "chkRemesaInd"
      Tab(2).Control(5)=   "lswRep"
      Tab(2).Control(6)=   "Label16(4)"
      Tab(2).Control(7)=   "Line1(1)"
      Tab(2).Control(8)=   "Label16(2)"
      Tab(2).Control(9)=   "lblRemesa"
      Tab(2).ControlCount=   10
      TabCaption(3)   =   "Consultas"
      TabPicture(3)   =   "frmFNDRemesas.frx":68A6
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "txtRetiro"
      Tab(3).Control(1)=   "txtConRemesa"
      Tab(3).Control(2)=   "Label1(1)"
      Tab(3).Control(3)=   "Line9(1)"
      Tab(3).Control(4)=   "Label1(2)"
      Tab(3).Control(5)=   "Label16(3)"
      Tab(3).ControlCount=   6
      Begin VB.TextBox txtCargaTotal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   7560
         Locked          =   -1  'True
         TabIndex        =   44
         Top             =   5760
         Width           =   2535
      End
      Begin VB.TextBox txtNotas 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   675
         Left            =   -71880
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   43
         Top             =   3360
         Width           =   6735
      End
      Begin VB.ComboBox cboBanco 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   300
         Left            =   3120
         Style           =   2  'Dropdown List
         TabIndex        =   38
         Top             =   840
         Width           =   6975
      End
      Begin VB.ComboBox cboCarga 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   300
         Left            =   3120
         Style           =   2  'Dropdown List
         TabIndex        =   33
         Top             =   480
         Width           =   6975
      End
      Begin VB.CheckBox chkCarga 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Todos"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   210
         Left            =   120
         TabIndex        =   32
         Top             =   1695
         Width           =   1455
      End
      Begin VB.TextBox txtRemesa 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -71880
         TabIndex        =   22
         Top             =   480
         Width           =   2655
      End
      Begin VB.TextBox txtUsuario 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   -71880
         Locked          =   -1  'True
         TabIndex        =   21
         Top             =   1920
         Width           =   2655
      End
      Begin VB.TextBox txtFecha 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   -71880
         Locked          =   -1  'True
         TabIndex        =   20
         Top             =   2280
         Width           =   2655
      End
      Begin VB.TextBox txtEstado 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   -71880
         Locked          =   -1  'True
         TabIndex        =   19
         Top             =   2640
         Width           =   2655
      End
      Begin VB.TextBox txtTotal 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   -71880
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   3000
         Width           =   2655
      End
      Begin VB.CommandButton cmdReporte 
         Caption         =   "&Reporte"
         Height          =   375
         Left            =   -66480
         TabIndex        =   14
         Top             =   5640
         Width           =   1455
      End
      Begin VB.TextBox txtRepRemesas 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   -65640
         TabIndex        =   13
         Text            =   "15"
         Top             =   3000
         Width           =   615
      End
      Begin VB.OptionButton opt 
         Caption         =   "Detalle de Remesa"
         Height          =   255
         Index           =   0
         Left            =   -74640
         TabIndex        =   12
         Top             =   3480
         Value           =   -1  'True
         Width           =   2055
      End
      Begin VB.OptionButton opt 
         Caption         =   "Detalle Agrupado de Remesa"
         Height          =   255
         Index           =   1
         Left            =   -74640
         TabIndex        =   11
         Top             =   3840
         Width           =   2655
      End
      Begin VB.CheckBox chkRemesaInd 
         Appearance      =   0  'Flat
         Caption         =   "Indicar Remesa"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   -66600
         TabIndex        =   10
         Top             =   5160
         Width           =   1455
      End
      Begin VB.TextBox txtRetiro 
         Height          =   315
         Left            =   -72840
         TabIndex        =   6
         Top             =   840
         Width           =   2055
      End
      Begin VB.TextBox txtConRemesa 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   4635
         Left            =   -72840
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   5
         Top             =   1200
         Width           =   7575
      End
      Begin MSComCtl2.DTPicker dtpInicio 
         Height          =   315
         Left            =   -71880
         TabIndex        =   1
         Top             =   1560
         Width           =   1335
         _ExtentX        =   2350
         _ExtentY        =   550
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   194838531
         CurrentDate     =   36278
      End
      Begin MSComCtl2.DTPicker dtpCorte 
         Height          =   315
         Left            =   -70560
         TabIndex        =   2
         Top             =   1560
         Width           =   1335
         _ExtentX        =   2350
         _ExtentY        =   550
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   194838531
         CurrentDate     =   36278
      End
      Begin MSComctlLib.ListView lswRemesas 
         Height          =   2055
         Left            =   -71880
         TabIndex        =   23
         Top             =   4080
         Width           =   6735
         _ExtentX        =   11875
         _ExtentY        =   3620
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
            Size            =   8.4
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
         Height          =   264
         Left            =   -69000
         TabIndex        =   24
         Top             =   960
         Width           =   1908
         _ExtentX        =   3366
         _ExtentY        =   466
         ButtonWidth     =   487
         ButtonHeight    =   466
         AllowCustomize  =   0   'False
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
      Begin MSComctlLib.ListView lswCarga 
         Height          =   3735
         Left            =   120
         TabIndex        =   34
         Top             =   1920
         Width           =   9975
         _ExtentX        =   17590
         _ExtentY        =   6583
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
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   11
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "ID"
            Object.Width           =   1658
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Plan"
            Object.Width           =   1658
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Contrato"
            Object.Width           =   1658
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Cédula"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Nombre"
            Object.Width           =   6068
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Text            =   "Monto"
            Object.Width           =   3246
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   6
            Text            =   "Tipo"
            Object.Width           =   1482
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Cuenta"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "Fecha"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "Usuario"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Text            =   "OP"
            Object.Width           =   776
         EndProperty
      End
      Begin MSComctlLib.Toolbar tlbCarga 
         Height          =   312
         Left            =   3120
         TabIndex        =   35
         Top             =   1200
         Width           =   4548
         _ExtentX        =   8022
         _ExtentY        =   550
         ButtonWidth     =   1757
         ButtonHeight    =   550
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
         TabIndex        =   46
         Top             =   720
         Width           =   9735
         _ExtentX        =   17166
         _ExtentY        =   3831
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
            Size            =   8.4
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
      Begin VB.Label Label3 
         Caption         =   "Total...:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6720
         TabIndex        =   45
         Top             =   5760
         Width           =   855
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   10
         X1              =   -74880
         X2              =   -72000
         Y1              =   3600
         Y2              =   3600
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   9
         X1              =   120
         X2              =   3000
         Y1              =   1080
         Y2              =   1080
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   8
         X1              =   120
         X2              =   3000
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
            Size            =   8.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   7
         Left            =   120
         TabIndex        =   36
         Top             =   1680
         Width           =   9975
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FFFFFF&
         X1              =   -75000
         X2              =   -65160
         Y1              =   840
         Y2              =   840
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   7
         X1              =   -74880
         X2              =   -72000
         Y1              =   1560
         Y2              =   1560
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   6
         X1              =   -74880
         X2              =   -72000
         Y1              =   4320
         Y2              =   4320
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   0
         X1              =   -74880
         X2              =   -72000
         Y1              =   720
         Y2              =   720
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   2
         X1              =   -74880
         X2              =   -72000
         Y1              =   2160
         Y2              =   2160
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   3
         X1              =   -74880
         X2              =   -72000
         Y1              =   2520
         Y2              =   2520
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   4
         X1              =   -74880
         X2              =   -72000
         Y1              =   2880
         Y2              =   2880
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   5
         X1              =   -74880
         X2              =   -72000
         Y1              =   3240
         Y2              =   3240
      End
      Begin VB.Label Label16 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Remesas - visualizar últimas"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.4
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
         TabIndex        =   17
         Top             =   3000
         Width           =   4935
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
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Seleccione la Remesa que Desea Visualizar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.4
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
         TabIndex        =   16
         Top             =   480
         Width           =   9735
      End
      Begin VB.Label lblRemesa 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   300
         Left            =   -74760
         TabIndex        =   15
         Top             =   3000
         Width           =   5295
      End
      Begin VB.Label Label1 
         Caption         =   "# Retiro"
         Height          =   255
         Index           =   1
         Left            =   -74040
         TabIndex        =   9
         Top             =   840
         Width           =   1095
      End
      Begin VB.Line Line9 
         BorderColor     =   &H00FFFFFF&
         Index           =   1
         X1              =   -74880
         X2              =   -65160
         Y1              =   765
         Y2              =   765
      End
      Begin VB.Label Label1 
         Caption         =   "Remesa"
         Height          =   255
         Index           =   2
         Left            =   -74040
         TabIndex        =   7
         Top             =   1200
         Width           =   1095
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
         Left            =   -70560
         TabIndex        =   4
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
         Left            =   -71880
         TabIndex        =   3
         Top             =   1320
         Width           =   1335
      End
      Begin VB.Label Label16 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Consultas"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   300
         Index           =   3
         Left            =   -74880
         TabIndex        =   8
         Top             =   480
         Width           =   2055
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Lista de Remesas"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   -74880
         TabIndex        =   30
         Top             =   4080
         Width           =   2895
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Total"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   5
         Left            =   -74880
         TabIndex        =   25
         Top             =   3000
         Width           =   2895
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Estado"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   4
         Left            =   -74880
         TabIndex        =   26
         Top             =   2640
         Width           =   2895
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Fecha"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   3
         Left            =   -74880
         TabIndex        =   27
         Top             =   2280
         Width           =   2895
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Usuario"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   2
         Left            =   -74880
         TabIndex        =   28
         Top             =   1920
         Width           =   2895
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Fechas"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   6
         Left            =   -74880
         TabIndex        =   31
         Top             =   1320
         Width           =   2895
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Remesa"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   -74880
         TabIndex        =   29
         Top             =   480
         Width           =   2895
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Remesa"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.4
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
         TabIndex        =   37
         Top             =   480
         Width           =   2895
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Banco"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   10
         Left            =   120
         TabIndex        =   39
         Top             =   840
         Width           =   2895
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Notas"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.4
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
         TabIndex        =   42
         Top             =   3360
         Width           =   2895
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   9240
      Top             =   0
      _ExtentX        =   995
      _ExtentY        =   995
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFNDRemesas.frx":68C2
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFNDRemesas.frx":1D284
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFNDRemesas.frx":33C46
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFNDRemesas.frx":48DB8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ProgressBar PrgBar 
      Align           =   2  'Align Bottom
      Height          =   156
      Left            =   0
      TabIndex        =   40
      Top             =   7272
      Width           =   10464
      _ExtentX        =   18457
      _ExtentY        =   275
      _Version        =   393216
      Appearance      =   0
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Remesas de Retiros y Liquidaciones de Planes"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Index           =   0
      Left            =   1080
      TabIndex        =   41
      Top             =   240
      Width           =   8532
   End
End
Attribute VB_Name = "frmFNDRemesas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim itmX As ListItem, vPaso As Boolean
Dim strSQL As String, rs As New ADODB.Recordset


Private Sub cboBanco_Click()
  lswCarga.ListItems.Clear
End Sub

Private Sub cboCarga_Click()
Dim vFechaInicio As Date, vFechaCorte As Date

lswCarga.ListItems.Clear

If vPaso Then Exit Sub
If cboCarga.ListCount <= 0 Then Exit Sub

vPaso = True
cboBanco.Clear


strSQL = "select fecha_inicio,fecha_corte from fnd_remesas where remesa = " & cboCarga.ItemData(cboCarga.ListIndex)
Call OpenRecordSet(rs, strSQL)
  vFechaInicio = rs!FECHA_INICIO
  vFechaCorte = rs!fecha_corte
rs.Close
'Seleccionar Tes_Bancos


cboBanco.AddItem "[TODOS LOS Tes_Bancos]"

strSQL = "select B.id_banco,B.descripcion" _
       & " from Fnd_Liquidacion L inner join Tes_Bancos B on L.cod_Banco = B.id_banco" _
       & " Where L.fecha between '" & Format(vFechaInicio, "yyyy/mm/dd") & " 00:00:00' and '" & Format(vFechaCorte, "yyyy/mm/dd") _
       & " 23:59:59' and L.consec not in (select consec from fnd_remesa_asg)" _
       & " group by B.id_banco,B.descripcion"
Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
  cboBanco.AddItem (Format(rs!id_Banco, "0000") & "..." & Trim(rs!Descripcion))
  cboBanco.ItemData(cboBanco.NewIndex) = rs!id_Banco
  rs.MoveNext
Loop
If rs.RecordCount > 0 Then
   rs.MoveFirst
   cboBanco.Text = (Format(rs!id_Banco, "0000") & "..." & Trim(rs!Descripcion))
End If
rs.Close

vPaso = False
Call cboBanco_Click

End Sub


Private Sub chkCarga_Click()
Dim i As Integer, curTotal As Currency


For i = 1 To lswCarga.ListItems.Count
  lswCarga.ListItems.Item(i).Checked = chkCarga.Value
  
   If chkCarga.Value = vbChecked Then
       curTotal = curTotal + CCur(lswCarga.ListItems.Item(i).SubItems(5))
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
 .WindowTitle = "Reportes del Módulo de Planes de Ahorro"

 .Connect = glogon.ConectRPT

If chkRemesaInd.Value = vbChecked Then
   xRemesa = InputBox("Indique la Remesa que desea consultar", "Remesas de Planes")
  If Len(Trim(xRemesa)) = 0 Then xRemesa = "0"
  lblRemesa.Tag = xRemesa
End If

 Select Case True
  Case opt.Item(0).Value 'Detalle Remesa
     .ReportFileName = SIFGlobal.fxPathReportes("Fondos_RemesasDetalle.rpt")
     vSubTitulo = "REMESA : " & lblRemesa.Tag & " LISTADO : DETALLADO"
  Case opt.Item(1).Value 'Detalle Agrupado Remesa
     .ReportFileName = SIFGlobal.fxPathReportes("Fondos_RemesasDetalleAgrupado.rpt")
     vSubTitulo = "REMESA : " & lblRemesa.Tag & " LISTADO : DETALLADO AGRUPADO"
 End Select
 
 .Formulas(0) = "fxFecha='FECHA: " & Format(fxFechaServidor, "dd/mm/yyyy") & "'"
 .Formulas(1) = "fxEmpresa='" & GLOBALES.gstrNombreEmpresa & "'"
 .Formulas(2) = "fxUsuario='USER: " & glogon.Usuario & "'"
 .Formulas(3) = "fxTitulo='REMESA DE PLANES : RETIROS Y LIQUIDACIONES'"
 .Formulas(4) = "fxSubTitulo='" & vSubTitulo & "'"
 .Formulas(5) = "fxFiltro='" & vFiltro & "'"
 .SelectionFormula = "{FND_REMESAS.REMESA} = " & lblRemesa.Tag
 .PrintReport

End With

Me.MousePointer = vbDefault
Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub Form_Activate()
 vModulo = 18
End Sub

Private Sub Form_Load()
 
vModulo = 18

Me.Icon = Me.Picture

 ssTab.Tab = 0
 Call sbToolBarIconos(tlb, False)
 
 Call Formularios(Me)
 Call RefrescaTags(Me)
 
 Call sbLimpia
 
End Sub


Private Sub sbConsulta(vRemesa As Long)

Call sbLimpia
  
strSQL = "select * from fnd_remesas where remesa = " & vRemesa
Call OpenRecordSet(rs, strSQL)
If Not rs.BOF And Not rs.EOF Then
  txtRemesa = rs!remesa
  txtUsuario = rs!Usuario
  txtFecha = rs!fecha
  
  Select Case rs!Estado
    Case "A"
      txtEstado = "Remesa Abierta"
    Case "C"
      txtEstado = "Remesa Cerrada"
    Case "P"
      txtEstado = "Remesa en Proceso"
  End Select
  
  dtpInicio.Value = rs!FECHA_INICIO
  dtpCorte.Value = rs!fecha_corte
  
  txtNotas.Text = rs!notas
  
  With glogon
    .strSQL = "select isnull(sum(aportes_liq + rendi_liq),0) as Total from fnd_liquidacion" _
            & " where consec in (select consec from fnd_remesa_asg where remesa = " & vRemesa & ")"
    .Recordset.Open .strSQL, .Conection, adOpenStatic
    txtTotal.Text = Format(.Recordset!Total, "Standard")
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
   curTotal = curTotal + CCur(Item.SubItems(5))
Else
   curTotal = curTotal - CCur(Item.SubItems(5))
End If

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

Private Sub SSTab_Click(PreviousTab As Integer)
 Call sbLimpia
End Sub

Private Sub sbReporte()
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
 .WindowTitle = "Reportes del Módulo de Planes de Ahorro"
 
 .Connect = glogon.ConectRPT
  
 
 .Formulas(0) = "fxFecha='FECHA: " & Format(fxFechaServidor, "dd/mm/yyyy") & "'"
 .Formulas(1) = "fxEmpresa='" & GLOBALES.gstrNombreEmpresa & "'"
 .Formulas(2) = "fxUsuario='USER: " & glogon.Usuario & "'"
 .Formulas(3) = "fxSubTitulo='" & vSubTitulo & "'"
 .Formulas(4) = "fxFiltro='" & vFiltro & "'"
 
 .ReportFileName = SIFGlobal.fxPathReportes("Fondos_Remesas.rpt")
 .PrintReport

End With

Me.MousePointer = vbDefault
Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub tlb_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim i As Integer

On Error GoTo vError

Select Case UCase(Button.Key)
  Case "NUEVO"
     
    If txtRemesa.Text = "" Then
     
            strSQL = "select isnull(max(remesa),0) + 1 as Ultimo from fnd_remesas"
            Call OpenRecordSet(rs, strSQL)
                strSQL = "insert fnd_remesas(remesa,usuario,fecha,estado,fecha_inicio,fecha_corte,notas) values(" & rs!ultimo _
                       & ",'" & glogon.Usuario & "',dbo.MyGetdate(),'A','" & Format(dtpInicio.Value, "yyyy/mm/dd") _
                       & "','" & Format(dtpCorte.Value, "yyyy/mm/dd") & "','" & txtNotas.Text & "')"
                Call ConectionExecute(strSQL)
                
                txtRemesa = rs!ultimo
            rs.Close
            Call Bitacora("Registra", "Remesa de Planes de Ahorro : " & txtRemesa)
    
    Else
        If txtEstado.Text <> "Remesa Cerrada" Then
                    
            strSQL = "update fnd_remesas set usuario = '" & glogon.Usuario & "',fecha_inicio = '" & Format(dtpInicio.Value, "yyyy/mm/dd") _
                   & "',fecha_corte = '" & Format(dtpCorte.Value, "yyyy/mm/dd") & "',notas = '" & txtNotas.Text _
                   & "' where remesa = " & txtRemesa
             Call ConectionExecute(strSQL)
            Call Bitacora("Modifica", "Remesa de Planes de Ahorro : " & txtRemesa)
        Else
            MsgBox "No se puede Modifica la remesa, porque esta ya fue cerrada...", vbExclamation
        End If
    End If
    
    Call sbLimpia
    
    
  Case "BORRAR"
     i = MsgBox("Esta Seguro que desea borrar este registro", vbYesNo)
     If i = vbYes Then
        If txtEstado.Text = "Remesa Abierta" Then
            strSQL = "delete fnd_remesas where remesa = " & txtRemesa
            Call ConectionExecute(strSQL)
            
            Call Bitacora("Elimina", "Remesa de Planes de Ahorro  : " & txtRemesa)
         End If
       
        Call sbLimpia
     End If
  
  Case "REPORTES"
     Call sbReporte

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
     
     
     strSQL = "select TOP 50 * from fnd_remesas order by fecha desc"
     lswRemesas.ListItems.Clear

     Call OpenRecordSet(rs, strSQL)
     Do While Not rs.EOF
       With lswRemesas.ListItems
            Set itmX = .Add(, , rs!remesa)
                itmX.SubItems(1) = rs!Usuario
                itmX.SubItems(2) = rs!fecha
                itmX.SubItems(3) = rs!FECHA_INICIO
                itmX.SubItems(4) = rs!fecha_corte
                itmX.SubItems(5) = rs!notas
       
       End With
       rs.MoveNext
     Loop
     rs.Close
     
  Case 1 'Carga
    'Solo busca las remesas que se encuentran Abiertas
    
    vPaso = True
    
    cboCarga.Clear
    cboBanco.Clear
    lswCarga.ListItems.Clear
    chkCarga.Value = vbUnchecked
        
    strSQL = "select * from fnd_remesas where estado in('A','P') order by fecha desc"
    Call OpenRecordSet(rs, strSQL)
    Do While Not rs.EOF
      cboCarga.AddItem (Format(rs!remesa, "0000") & "..." & Trim(rs!Usuario) & "..." & rs!fecha)
      cboCarga.ItemData(cboCarga.NewIndex) = rs!remesa
      rs.MoveNext
    Loop
    If rs.RecordCount > 0 Then
       rs.MoveFirst
       cboCarga.Text = (Format(rs!remesa, "0000") & "..." & Trim(rs!Usuario) & "..." & rs!fecha)
    End If
    
    rs.Close

    vPaso = False
    Call cboCarga_Click

  Case 2 'Reportes
     strSQL = "select TOP " & txtRepRemesas.Text & " * from fnd_remesas order by fecha desc"
     lswRep.ListItems.Clear

     Call OpenRecordSet(rs, strSQL)
     Do While Not rs.EOF
       With lswRep.ListItems
            Set itmX = .Add(, , rs!remesa)
                itmX.SubItems(1) = rs!Usuario
                itmX.SubItems(2) = rs!fecha
                itmX.SubItems(3) = rs!FECHA_INICIO
                itmX.SubItems(4) = rs!fecha_corte
                itmX.SubItems(5) = rs!notas
       
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



strSQL = "select fecha_inicio,fecha_corte from fnd_remesas where remesa = " & cboCarga.ItemData(cboCarga.ListIndex)
Call OpenRecordSet(rs, strSQL)
  vFechaInicio = rs!FECHA_INICIO
  vFechaCorte = rs!fecha_corte
rs.Close

'Seleccionar Tes_Bancos

strSQL = "select L.consec,S.cedula,S.nombre,L.cod_plan,L.cod_operadora,L.cod_contrato,(L.aportes_liq+L.rendi_liq) as Monto" _
       & ",L.Fecha,L.usuario,L.cta_ahorros,L.tipo" _
       & " from fnd_liquidacion L inner join fnd_contratos C on L.cod_operadora = C.cod_operadora" _
       & " and L.cod_plan = C.cod_plan and L.cod_contrato = C.cod_contrato" _
       & " inner join Socios S on C.cedula = S.cedula"

If cboBanco.Text = "[TODOS LOS Tes_Bancos]" Then
   strSQL = strSQL & " where L.fecha between '" & Format(vFechaInicio, "yyyy/mm/dd") & " 00:00:00' and '" & Format(vFechaCorte, "yyyy/mm/dd") _
           & " 23:59:59' and L.traspaso_tesoreria is not null and L.consec not in(select consec from fnd_remesa_asg)"
Else
   strSQL = strSQL & " where L.cod_banco = " & cboBanco.ItemData(cboBanco.ListIndex) _
           & " and L.fecha between '" & Format(vFechaInicio, "yyyy/mm/dd") & " 00:00:00' and '" & Format(vFechaCorte, "yyyy/mm/dd") _
           & " 23:59:59' and L.traspaso_tesoreria is not null and L.consec not in(select consec from fnd_remesa_asg)"
End If
Call OpenRecordSet(rs, strSQL)

prgBar.Max = rs.RecordCount + 1
prgBar.Value = 1
prgBar.Visible = True

Do While Not rs.EOF
 Set itmX = lswCarga.ListItems.Add(, , rs!consec)
     itmX.SubItems(1) = rs!cod_Plan
     itmX.SubItems(2) = rs!COD_CONTRATO
     itmX.SubItems(3) = rs!Cedula
     itmX.SubItems(4) = rs!Nombre
     itmX.SubItems(5) = Format(rs!Monto, "Standard")
     itmX.SubItems(6) = Trim(rs!Tipo)
     itmX.SubItems(7) = Trim(rs!Cta_Ahorros)
     itmX.SubItems(8) = rs!fecha
     itmX.SubItems(9) = rs!Usuario
     itmX.SubItems(10) = rs!cod_Operadora
     
     itmX.Checked = chkCarga.Value
     
     If itmX.Checked Then
        curTotal = curTotal + CCur(itmX.SubItems(5))
     End If
     
 rs.MoveNext
 
 prgBar.Value = prgBar.Value + 1
 
Loop
rs.Close

prgBar.Visible = False

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
strSQL = "select count(*) as Existe from fnd_remesas" _
       & " where remesa = " & cboCarga.ItemData(cboCarga.ListIndex) _
       & " and estado in('A','P')"
Call OpenRecordSet(rs, strSQL)
 If rs!Existe = 0 Then
    rs.Close
    MsgBox "La Remesa actual; ya se encuentra cerrada...", vbExclamation
    Exit Sub
 End If
rs.Close

'Actualiza el Estado de la Remesa como cerrada
strSQL = "update fnd_remesas set estado = 'C'" _
       & " where remesa = " & cboCarga.ItemData(cboCarga.ListIndex)
Call ConectionExecute(strSQL)

 
Call Bitacora("Genera", "Remesa de Fondos Remesa : " & cboCarga.ItemData(cboCarga.ListIndex))


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
strSQL = "select count(*) as Existe from fnd_remesas" _
       & " where remesa = " & cboCarga.ItemData(cboCarga.ListIndex) _
       & " and estado in('A','P') "
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

prgBar.Max = vCasos
prgBar.Value = 1
prgBar.Visible = True


With lswCarga.ListItems

For i = 1 To .Count
 If .Item(i).Checked Then
 
     strSQL = "insert fnd_remesa_asg(remesa,consec) values(" & cboCarga.ItemData(cboCarga.ListIndex) _
             & "," & .Item(i).Text & ")"
     Call ConectionExecute(strSQL)
   
    prgBar.Value = prgBar.Value + 1
  End If
Next i
 
If vCasos > 0 Then
    'Actualiza el Estado de la Remesa como cerrada
    strSQL = "update fnd_remesas set estado = 'P'" _
           & " where remesa = " & cboCarga.ItemData(cboCarga.ListIndex)
    Call ConectionExecute(strSQL)
    Call Bitacora("Genera", "Remesa de Fondos : " & cboCarga.ItemData(cboCarga.ListIndex))
End If

End With

prgBar.Visible = False

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
If cboBanco.ListCount = 0 Then Exit Sub

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

Private Sub tlbReporte_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim vSubTitulo As String, vFiltro As String
Dim strSQL As String, vTipoUser As String

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
 .WindowTitle = "Reportes del Sub.Módulo de Comisiones de Afiliación"
 
 .Connect = glogon.ConectRPT
  
' If chkRepFechas.Value = vbUnchecked Then
'    If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
'    Select Case Mid(cboRepBase.Text, 1, 1)
'      Case "R" 'Fecha de Creación de la Remesa
'        strSQL = strSQL & "{AFI_COMISIONES.FECHA}"
'        vSubTitulo = "Generadas entre " & Format(dtpRepInicio.Value, "dd/mm/yyyy") & " y " & Format(dtpRepCorte.Value, "dd/mm/yyyy")
'      Case "P" 'Fecha de Traslado a Tesoreria
'        strSQL = strSQL & "{AFI_COMISION_PAGO.TRASLADO_FECHA}"
'        vSubTitulo = "Pagadas entre " & Format(dtpRepInicio.Value, "dd/mm/yyyy") & " y " & Format(dtpRepCorte.Value, "dd/mm/yyyy")
'    End Select
'    strSQL = strSQL & " in Date(" & Format(dtpRepInicio.Value, "yyyy,mm,dd") & ") to date(" _
'           & Format(dtpRepCorte.Value, "yyyy,mm,dd") & ")"
' Else
'   vSubTitulo = "Historico"
' End If
'
' If chkRepRemesas.Value = vbUnchecked Then
'   If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
'   strSQL = strSQL & "{AFI_COMISION_PAGO.COD_COMISION} = " & txtRepRemesa.Tag
'   vFiltro = vFiltro & "/ REMESA : " & txtRepRemesa.Text
' Else
'   vFiltro = vFiltro & "/ TODAS LAS REMESAS"
' End If
'
'
' If chkRepPromotor.Value = vbUnchecked Then
'   If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
'   strSQL = strSQL & "{AFI_COMISION_PAGO.ID_PROMOTOR} = " & txtRepPromotor.Tag
'   vFiltro = vFiltro & "/ PROMOTOR : " & txtRepPromotor.Text
' Else
'   vFiltro = vFiltro & "/ TODOS LOS PROMOTORES"
' End If
'
' If chkRepBancos.Value = vbUnchecked Then
'   If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
'   strSQL = strSQL & "{AFI_COMISION_PAGO.COD_BANCO} = " & txtRepBanco.Tag
'   vFiltro = vFiltro & "/ BANCO : " & txtRepBanco.Text
' Else
'   vFiltro = vFiltro & "/ TODOS LOS Tes_Bancos"
' End If
'
'
' If chkRepUsuario.Value = vbUnchecked Then
'   If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
'
'    Select Case Mid(cboRepBase.Text, 1, 1)
'      Case "R" 'Fecha de Creación de la Remesa
'            strSQL = strSQL & "{AFI_COMISIONES.USUARIO} = '" & txtRepUsuario.Text & "'"
'            vFiltro = vFiltro & "/ USUARIO : " & txtRepUsuario.Text
'      Case "P" 'Fecha de Traslado a Tesoreria
'            strSQL = strSQL & "{AFI_COMISION_PAGO.TRASLADO_USER} = '" & txtRepUsuario.Text & "'"
'            vFiltro = vFiltro & "/ USUARIO : " & txtRepUsuario.Text
'    End Select
'
' Else
'   vFiltro = vFiltro & "/ TODOS LOS USUARIOS"
' End If
'
'
'If chkRepSinComision.Value = vbUnchecked Then
'   If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
'
'   strSQL = strSQL & "{AFI_COMISION_PAGO.MONTO} > 0"
'    vFiltro = vFiltro & "/ SOLO CASOS CON MONTO > 0"
'End If
'
'
' .Formulas(0) = "fxFecha='FECHA: " & Format(fxFechaServidor, "dd/mm/yyyy") & "'"
' .Formulas(1) = "fxEmpresa='" & GLOBALES.gstrNombreEmpresa & "'"
' .Formulas(2) = "fxUsuario='USER: " & glogon.Usuario & "'"
' .Formulas(3) = "fxSubTitulo='" & vSubTitulo & "'"
' .Formulas(4) = "fxFiltro='" & vFiltro & "'"
'
' If cboRepTipo.Text = "Detalle" Then
'   Select Case True
'     Case optReportes.Item(0).Value 'Listado General
'         .ReportFileName = SIFGlobal.fxPathReportes("Personas_ComisionListadoGeneral.rpt"
'     Case optReportes.Item(1).Value 'Agrupado x Promotor
'         .ReportFileName = SIFGlobal.fxPathReportes("Personas_ComisionAgrpPromotor.rpt"
'     Case optReportes.Item(2).Value 'Agrupado x Usuario
'         .ReportFileName = SIFGlobal.fxPathReportes("Personas_ComisionAgrpUsuario.rpt"
'     Case optReportes.Item(3).Value 'Agrupado x Banco
'         .ReportFileName = SIFGlobal.fxPathReportes("Personas_ComisionAgrpBanco.rpt"
'     Case optReportes.Item(4).Value 'Tesoreria
'         .ReportFileName = SIFGlobal.fxPathReportes("Personas_ComisionTesoreria.rpt"
'    End Select
' Else
'   Select Case True
'     Case optReportes.Item(0).Value 'Listado General
'         .ReportFileName = SIFGlobal.fxPathReportes("Personas_ComisionListadoGeneralRsm.rpt"
'     Case optReportes.Item(1).Value 'Agrupado x Promotor
'         .ReportFileName = SIFGlobal.fxPathReportes("Personas_ComisionAgrpPromotorRsm.rpt"
'     Case optReportes.Item(2).Value 'Agrupado x Usuario
'         .ReportFileName = SIFGlobal.fxPathReportes("Personas_ComisionAgrpUsuarioRsm.rpt"
'     Case optReportes.Item(3).Value 'Agrupado x Banco
'         .ReportFileName = SIFGlobal.fxPathReportes("Personas_ComisionAgrpBancoRsm.rpt"
'     Case optReportes.Item(4).Value 'Tesoreria
'         .ReportFileName = SIFGlobal.fxPathReportes("Personas_ComisionTesoreriaRsm.rpt"
'    End Select
' End If
 
 .SelectionFormula = strSQL
 .PrintReport

End With

Me.MousePointer = vbDefault
Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub txtRemesa_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn And IsNumeric(txtRemesa) Then
   Call sbConsulta(txtRemesa)
End If
End Sub


Private Sub txtRetiro_Change()
 txtConRemesa.Text = ""
End Sub


Private Sub sbConsultaRetiro()

On Error GoTo vError


strSQL = "select A.* from fnd_remesas A inner join fnd_remesa_asg X on A.remesa = X.remesa where consec = " & txtRetiro.Text
Call OpenRecordSet(rs, strSQL)
If rs.BOF Or rs.EOF Then
 txtConRemesa.Text = "** No se encontró retiro/liq. en las remesas registradas **"
Else
 txtConRemesa.Text = "Remesa   " & vbTab & " ...:" & rs!remesa & vbCrLf
 txtConRemesa.Text = txtConRemesa & "Fecha   " & vbTab & " ...:" & rs!fecha & vbCrLf
 txtConRemesa.Text = txtConRemesa & "Usuario  " & vbTab & " ...:" & rs!Usuario
End If
rs.Close

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
 txtConRemesa.Text = ""

End Sub

Private Sub txtRetiro_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then Call sbConsultaRetiro
End Sub
