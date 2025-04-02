VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmCR_RetencionEnFondos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Retenciones en Fondos Liquidados"
   ClientHeight    =   7455
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   10515
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   Picture         =   "frmCR_RentencionEnFondos.frx":0000
   ScaleHeight     =   7455
   ScaleWidth      =   10515
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   9360
      Top             =   120
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
            Picture         =   "frmCR_RentencionEnFondos.frx":0332
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_RentencionEnFondos.frx":16CF4
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_RentencionEnFondos.frx":2D6B6
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_RentencionEnFondos.frx":42828
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ProgressBar PrgBar 
      Align           =   2  'Align Bottom
      Height          =   150
      Left            =   0
      TabIndex        =   1
      Top             =   7305
      Width           =   10515
      _ExtentX        =   18547
      _ExtentY        =   265
      _Version        =   393216
      Appearance      =   0
   End
   Begin TabDlg.SSTab ssTab 
      Height          =   6255
      Left            =   120
      TabIndex        =   2
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
      TabCaption(0)   =   "Remesas"
      TabPicture(0)   =   "frmCR_RentencionEnFondos.frx":5799A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label2(1)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label2(0)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label2(5)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label2(4)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label2(3)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label2(2)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label2(6)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label2(9)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Line1(10)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Line2"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Line1(7)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Line1(6)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Line1(0)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Line1(2)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Line1(3)"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Line1(4)"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Line1(5)"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Label1(6)"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Label1(5)"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "Label2(10)"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "Label2(11)"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "Line1(9)"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "Line1(11)"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "tlb"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "lswRemesas"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "dtpCorte"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "dtpInicio"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "txtNotas"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "txtRemesa"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "txtUsuario"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "txtFecha"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "txtEstado"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "txtTotal"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "cboCliente"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "cboPlan"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).ControlCount=   35
      TabCaption(1)   =   "Cargado"
      TabPicture(1)   =   "frmCR_RentencionEnFondos.frx":579B6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label2(12)"
      Tab(1).Control(1)=   "Label2(8)"
      Tab(1).Control(2)=   "Label3"
      Tab(1).Control(3)=   "Line1(8)"
      Tab(1).Control(4)=   "Label2(7)"
      Tab(1).Control(5)=   "Line1(12)"
      Tab(1).Control(6)=   "Line1(13)"
      Tab(1).Control(7)=   "tlbCarga"
      Tab(1).Control(8)=   "lswCarga"
      Tab(1).Control(9)=   "txtCargaTotal"
      Tab(1).Control(10)=   "cboCarga"
      Tab(1).Control(11)=   "chkCarga"
      Tab(1).Control(12)=   "cboInstitucion"
      Tab(1).Control(13)=   "chkMostrarNoAsignado"
      Tab(1).ControlCount=   14
      TabCaption(2)   =   "Reportes"
      TabPicture(2)   =   "frmCR_RentencionEnFondos.frx":579D2
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "lblRemesa"
      Tab(2).Control(1)=   "Label16(4)"
      Tab(2).Control(2)=   "Line1(1)"
      Tab(2).Control(3)=   "Label16(2)"
      Tab(2).Control(4)=   "cmdArchivo"
      Tab(2).Control(5)=   "lswRep"
      Tab(2).Control(6)=   "cmdReporte"
      Tab(2).Control(7)=   "txtRepRemesas"
      Tab(2).Control(8)=   "opt(0)"
      Tab(2).Control(9)=   "opt(1)"
      Tab(2).Control(10)=   "chkRemesaInd"
      Tab(2).Control(11)=   "chkArchivoAddInst"
      Tab(2).Control(12)=   "chkArchivoSepInst"
      Tab(2).Control(13)=   "opt(2)"
      Tab(2).ControlCount=   14
      TabCaption(3)   =   "Consultas"
      TabPicture(3)   =   "frmCR_RentencionEnFondos.frx":579EE
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "SSTabY"
      Tab(3).ControlCount=   1
      Begin TabDlg.SSTab SSTabY 
         Height          =   5535
         Left            =   -74760
         TabIndex        =   51
         Top             =   600
         Width           =   9735
         _ExtentX        =   17171
         _ExtentY        =   9763
         _Version        =   393216
         Style           =   1
         Tabs            =   2
         TabsPerRow      =   2
         TabHeight       =   520
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "# Retiro"
         TabPicture(0)   =   "frmCR_RentencionEnFondos.frx":57A0A
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Label1(2)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "Label1(1)"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "txtConRemesa"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "txtRetiro"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).ControlCount=   4
         TabCaption(1)   =   "Persona"
         TabPicture(1)   =   "frmCR_RentencionEnFondos.frx":57A26
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Label4(0)"
         Tab(1).Control(1)=   "Label4(1)"
         Tab(1).Control(2)=   "lswConsulta"
         Tab(1).Control(3)=   "txtCedula"
         Tab(1).Control(4)=   "txtNombre"
         Tab(1).ControlCount=   5
         Begin VB.TextBox txtNombre 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFC0&
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
            Left            =   -72240
            Locked          =   -1  'True
            TabIndex        =   58
            Top             =   840
            Width           =   5655
         End
         Begin VB.TextBox txtCedula 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
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
            Left            =   -73800
            MaxLength       =   15
            TabIndex        =   57
            Top             =   840
            Width           =   1575
         End
         Begin VB.TextBox txtRetiro 
            Height          =   315
            Left            =   1320
            TabIndex        =   53
            Top             =   480
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
            Height          =   4515
            Left            =   1320
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   52
            Top             =   840
            Width           =   7575
         End
         Begin MSComctlLib.ListView lswConsulta 
            Height          =   3975
            Left            =   -74880
            TabIndex        =   56
            Top             =   1320
            Width           =   9495
            _ExtentX        =   16748
            _ExtentY        =   7011
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
            NumItems        =   13
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Remesa"
               Object.Width           =   2187
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Usuario"
               Object.Width           =   3775
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
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   6
               Text            =   "Plan"
               Object.Width           =   2011
            EndProperty
            BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   7
               Text            =   "Cliente"
               Object.Width           =   2011
            EndProperty
            BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   8
               Text            =   "Institución"
               Object.Width           =   2011
            EndProperty
            BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   9
               Text            =   "# Retiro"
               Object.Width           =   2011
            EndProperty
            BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   10
               Text            =   "# Contrato"
               Object.Width           =   2011
            EndProperty
            BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   11
               Text            =   "Plan"
               Object.Width           =   2011
            EndProperty
            BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   12
               Text            =   "Monto"
               Object.Width           =   2540
            EndProperty
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Nombre"
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
            Height          =   375
            Index           =   1
            Left            =   -72240
            TabIndex        =   60
            Top             =   600
            Width           =   5655
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Cédula"
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
            Height          =   375
            Index           =   0
            Left            =   -73800
            TabIndex        =   59
            Top             =   600
            Width           =   1575
         End
         Begin VB.Label Label1 
            Caption         =   "# Retiro"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   55
            Top             =   480
            Width           =   1095
         End
         Begin VB.Label Label1 
            Caption         =   "Remesa"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   54
            Top             =   840
            Width           =   1095
         End
      End
      Begin VB.CheckBox chkMostrarNoAsignado 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Mostrar Casos no Asignados"
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
         Height          =   225
         Left            =   -74880
         TabIndex        =   50
         Top             =   1320
         Width           =   2895
      End
      Begin VB.OptionButton opt 
         Caption         =   "Resumen de Remesa"
         Height          =   255
         Index           =   2
         Left            =   -74640
         TabIndex        =   49
         Top             =   4200
         Width           =   2655
      End
      Begin VB.CheckBox chkArchivoSepInst 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "Crear archivos separados por cada Institución"
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
         Left            =   -72600
         TabIndex        =   48
         Top             =   5280
         Value           =   1  'Checked
         Width           =   3615
      End
      Begin VB.CheckBox chkArchivoAddInst 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "Incluir la Institución en el Archivo"
         Enabled         =   0   'False
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
         Left            =   -71760
         TabIndex        =   47
         Top             =   5640
         Width           =   2775
      End
      Begin VB.ComboBox cboInstitucion 
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
         Height          =   300
         Left            =   -71880
         Style           =   2  'Dropdown List
         TabIndex        =   44
         Top             =   840
         Width           =   6975
      End
      Begin VB.ComboBox cboPlan 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   8760
         Style           =   2  'Dropdown List
         TabIndex        =   43
         Top             =   2280
         Width           =   1095
      End
      Begin VB.ComboBox cboCliente 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   8760
         TabIndex        =   42
         Text            =   "cboCliente"
         Top             =   1920
         Width           =   1095
      End
      Begin VB.CheckBox chkRemesaInd 
         Appearance      =   0  'Flat
         Caption         =   "Indicar Remesa"
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
         Left            =   -66600
         TabIndex        =   16
         Top             =   4440
         Width           =   1455
      End
      Begin VB.OptionButton opt 
         Caption         =   "Detalle Agrupado de Remesa"
         Height          =   255
         Index           =   1
         Left            =   -74640
         TabIndex        =   15
         Top             =   3840
         Width           =   2655
      End
      Begin VB.OptionButton opt 
         Caption         =   "Detalle de Remesa"
         Height          =   255
         Index           =   0
         Left            =   -74640
         TabIndex        =   14
         Top             =   3480
         Value           =   -1  'True
         Width           =   2055
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
      Begin VB.CommandButton cmdReporte 
         Caption         =   "&Reporte"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   -66480
         Picture         =   "frmCR_RentencionEnFondos.frx":57A42
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   5160
         Width           =   1455
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
         TabIndex        =   11
         Top             =   3000
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
         TabIndex        =   10
         Top             =   2640
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
         TabIndex        =   9
         Top             =   2280
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
         TabIndex        =   8
         Top             =   1920
         Width           =   2655
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
         TabIndex        =   7
         Top             =   480
         Width           =   2655
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
         TabIndex        =   6
         Top             =   1815
         Width           =   1455
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
         Height          =   300
         Left            =   -71880
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   480
         Width           =   6975
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
         TabIndex        =   4
         Top             =   3360
         Width           =   6735
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
         TabIndex        =   3
         Top             =   5760
         Width           =   2535
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
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   111345667
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
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   111345667
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
         NumItems        =   8
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
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Plan"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Cliente"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComctlLib.Toolbar tlb 
         Height          =   264
         Left            =   6000
         TabIndex        =   20
         Top             =   960
         Width           =   1908
         _ExtentX        =   3360
         _ExtentY        =   476
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
         Height          =   3615
         Left            =   -74880
         TabIndex        =   21
         Top             =   2040
         Width           =   9975
         _ExtentX        =   17595
         _ExtentY        =   6376
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
         Height          =   330
         Left            =   -71880
         TabIndex        =   22
         Top             =   1320
         Width           =   4545
         _ExtentX        =   8017
         _ExtentY        =   556
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
         NumItems        =   8
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
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Plan"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Cliente"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.CommandButton cmdArchivo 
         Caption         =   "Generar &Archivo"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   -68520
         Picture         =   "frmCR_RentencionEnFondos.frx":57BB5
         Style           =   1  'Graphical
         TabIndex        =   46
         Top             =   5160
         Width           =   2055
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   13
         X1              =   -74880
         X2              =   -72000
         Y1              =   1560
         Y2              =   1560
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   12
         X1              =   -74880
         X2              =   -72000
         Y1              =   1080
         Y2              =   1080
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   11
         X1              =   6000
         X2              =   8880
         Y1              =   2160
         Y2              =   2160
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   9
         X1              =   6000
         X2              =   8880
         Y1              =   2520
         Y2              =   2520
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Código de Plan (Fondo)"
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
         Left            =   6000
         TabIndex        =   41
         Top             =   2280
         Width           =   2655
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Código de Cliente"
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
         Left            =   6000
         TabIndex        =   40
         Top             =   1920
         Width           =   2655
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
         TabIndex        =   30
         Top             =   1320
         Width           =   1335
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
         TabIndex        =   29
         Top             =   1320
         Width           =   1335
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
         TabIndex        =   27
         Top             =   480
         Width           =   9735
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   1
         X1              =   -74760
         X2              =   -65040
         Y1              =   4800
         Y2              =   4800
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
         TabIndex        =   26
         Top             =   3000
         Width           =   4935
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   5
         X1              =   120
         X2              =   3000
         Y1              =   3240
         Y2              =   3240
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
         Index           =   3
         X1              =   120
         X2              =   3000
         Y1              =   2520
         Y2              =   2520
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
         Index           =   0
         X1              =   120
         X2              =   3000
         Y1              =   720
         Y2              =   720
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
         Index           =   7
         X1              =   120
         X2              =   3000
         Y1              =   1560
         Y2              =   1560
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FFFFFF&
         X1              =   0
         X2              =   9840
         Y1              =   840
         Y2              =   840
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
         TabIndex        =   25
         Top             =   1800
         Width           =   9975
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   8
         X1              =   -74880
         X2              =   -72000
         Y1              =   720
         Y2              =   720
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   10
         X1              =   120
         X2              =   3000
         Y1              =   3600
         Y2              =   3600
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
         Left            =   -68280
         TabIndex        =   24
         Top             =   5760
         Width           =   855
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
         TabIndex        =   38
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
         Index           =   9
         Left            =   120
         TabIndex        =   39
         Top             =   3360
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
         TabIndex        =   36
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
         TabIndex        =   35
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
         TabIndex        =   34
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
         TabIndex        =   33
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
         TabIndex        =   32
         Top             =   3000
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
         TabIndex        =   31
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
         Index           =   1
         Left            =   120
         TabIndex        =   37
         Top             =   480
         Width           =   2895
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Instituciones"
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
         TabIndex        =   45
         Top             =   840
         Width           =   2895
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
         TabIndex        =   28
         Top             =   3000
         Width           =   5295
      End
   End
   Begin VB.Label Label1 
      Caption         =   "Remesas de Liquidaciones de Fondos de Clientes"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   960
      TabIndex        =   0
      Top             =   240
      Width           =   8295
   End
End
Attribute VB_Name = "frmCR_RetencionEnFondos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim itmX As ListItem, vPaso As Boolean
Dim vMensaje As String

Private Sub cboCarga_Click()
Dim strSQL As String, rs As New ADODB.Recordset
Dim vFechaInicio As Date, vFechaCorte As Date
Dim vPlan As String, vCliente As String

lswCarga.ListItems.Clear

If vPaso Then Exit Sub
If cboCarga.ListCount <= 0 Then Exit Sub

vPaso = True
cboInstitucion.Clear

Me.MousePointer = vbHourglass

strSQL = "select * from fnd_convenios_remesa where remesa = " & cboCarga.ItemData(cboCarga.ListIndex)
Call OpenRecordSet(rs, strSQL)
  vFechaInicio = rs!FECHA_INICIO
  vFechaCorte = rs!fecha_corte
  vPlan = rs!cod_Plan
  vCliente = rs!cod_cliente
rs.Close

'Seleccionar Instituciones
strSQL = "select I.cod_institucion,I.descripcion" _
       & " from Fnd_Liquidacion L inner join Fnd_contratos C on L.cod_operadora = C.cod_operadora" _
       & " and L.cod_plan = C.cod_plan and L.cod_contrato = C.cod_contrato" _
       & " inner join Socios S on C.cedula = S.cedula" _
       & " inner join instituciones I on S.cod_institucion = I.cod_institucion" _
       & " Where L.fecha between '" & Format(vFechaInicio, "yyyy/mm/dd") & " 00:00:00' and '" & Format(vFechaCorte, "yyyy/mm/dd") _
       & " 23:59:59' and L.consec not in (select consec from fnd_convenios_remesa_asg)" _
       & " and L.cod_plan = '" & vPlan & "'"
       
If chkMostrarNoAsignado.Value = vbChecked Then
  strSQL = strSQL & " and C.cedula not in(select cedula from reg_creditos where estado in('A','C'))" _
         & " group by I.cod_institucion,I.descripcion"
Else
  strSQL = strSQL & " and C.cedula in(select cedula from reg_creditos where estado in('A','C')" _
         & " and codigo = '" & vCliente & "')" _
         & " group by I.cod_institucion,I.descripcion"
End If

Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
  cboInstitucion.AddItem (Format(rs!cod_institucion, "0000") & "..." & Trim(rs!Descripcion))
  cboInstitucion.ItemData(cboInstitucion.NewIndex) = rs!cod_institucion
  rs.MoveNext
Loop
If rs.RecordCount > 0 Then
   rs.MoveFirst
   cboInstitucion.Text = (Format(rs!cod_institucion, "0000") & "..." & Trim(rs!Descripcion))
End If
rs.Close

'cboInstitucion.AddItem "[TODAS]"
'cboInstitucion.Text = "[TODAS]"


Me.MousePointer = vbDefault

vPaso = False
Call cboInstitucion_Click


End Sub


Private Sub cboInstitucion_Click()
  lswCarga.ListItems.Clear
End Sub

Private Sub chkArchivoSepInst_Click()

chkArchivoAddInst.Value = vbUnchecked

If chkArchivoSepInst.Value = vbChecked Then
   chkArchivoAddInst.Enabled = False
Else
   chkArchivoAddInst.Enabled = True
End If

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


Private Sub sbArchivoInstSeparado(pRemesa As Long, pInstitucion As Long)
Dim strSQL As String, rs As New ADODB.Recordset
Dim vCadena As String, vTempo As String, vInstDescripcion As String
Dim vFile As String, vArchivo As String, vRuta As String
Dim fnFile


fnFile = FreeFile

strSQL = "select descripcion from instituciones where cod_institucion = " & pInstitucion
Call OpenRecordSet(rs, strSQL)
  vInstDescripcion = Trim(rs!Descripcion)
rs.Close

strSQL = "select cod_cliente,cod_plan from FND_CONVENIOS_REMESA where remesa = " & pRemesa
Call OpenRecordSet(rs, strSQL)
  vArchivo = Format(pRemesa, "0000") & " " & Trim(rs!cod_cliente) & " I." _
           & Format(pInstitucion, "00") & " " & vInstDescripcion & ".txt"
rs.Close



'Crea Directorios

On Error Resume Next

MkDir SIFGlobal.DirectorioDeResultados
MkDir SIFGlobal.DirectorioDeResultados & "\RemesaFndConvenios"

vRuta = SIFGlobal.DirectorioDeResultados & "\RemesaFndConvenios"

vTempo = vRuta & "\" & vArchivo

vFile = Dir(vTempo, vbArchive)

If vFile = vArchivo Then  'El archivo existe
 Kill vTempo
End If


On Error GoTo vError

strSQL = "select S.cedula,S.nombre,R.cod_institucion,I.descripcion as Institucion,L.aportes_liq as Monto" _
       & " from FND_CONVENIOS_REMESA_ASG R inner join FND_Liquidacion L on R.Consec = L.consec" _
       & " inner join instituciones I on R.cod_institucion = I.cod_institucion" _
       & " inner join fnd_contratos C on L.cod_operadora = C.cod_operadora" _
       & " and L.cod_plan = C.cod_plan and L.cod_contrato = C.cod_contrato" _
       & " inner join Socios S on C.cedula = S.cedula" _
       & " where R.remesa = " & pRemesa & " and R.cod_institucion = " & pInstitucion _
       & " order by S.cedula"
       
rs.CursorLocation = adUseServer
Call OpenRecordSet(rs, strSQL)

Open vTempo For Output As #fnFile  ' Create file name.

Do While Not rs.EOF
  vCadena = SIFGlobal.fxStringRelleno(rs!Cedula, "D", " ", 15)
  vCadena = vCadena & SIFGlobal.fxStringRelleno(rs!Nombre, "D", " ", 50)
  vCadena = vCadena & Format(rs!Monto, "000000000.00")
  Print #fnFile, vCadena
  rs.MoveNext
Loop
rs.Close

Close #fnFile
  
If vMensaje = "" Then vMensaje = "Archivos Creados : " & vbCrLf
  
vMensaje = vMensaje & vbCrLf & "--> " & vTempo


Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub sbArchivoGeneral(pRemesa As Long)
Dim strSQL As String, rs As New ADODB.Recordset
Dim vCadena As String, vTempo As String
Dim vFile As String, vArchivo As String, vRuta As String
Dim fnFile


fnFile = FreeFile

strSQL = "select cod_cliente,cod_plan from FND_CONVENIOS_REMESA where remesa = " & pRemesa
Call OpenRecordSet(rs, strSQL)
  vArchivo = Format(pRemesa, "0000") & " " & Trim(rs!cod_cliente) & " [General].txt"
rs.Close


'Crea Directorios

On Error Resume Next

MkDir SIFGlobal.DirectorioDeResultados
MkDir SIFGlobal.DirectorioDeResultados & "\RemesaFndConvenios"

vRuta = SIFGlobal.DirectorioDeResultados & "\RemesaFndConvenios"

vTempo = vRuta & "\" & vArchivo

vFile = Dir(vTempo, vbArchive)

If vFile = vArchivo Then  'El archivo existe
 Kill vTempo
End If


On Error GoTo vError

strSQL = "select S.cedula,S.nombre,R.cod_institucion,I.descripcion as Institucion,L.aportes_liq as Monto" _
       & " from FND_CONVENIOS_REMESA_ASG R inner join FND_Liquidacion L on R.Consec = L.consec" _
       & " inner join instituciones I on R.cod_institucion = I.cod_institucion" _
       & " inner join fnd_contratos C on L.cod_operadora = C.cod_operadora" _
       & " and L.cod_plan = C.cod_plan and L.cod_contrato = C.cod_contrato" _
       & " inner join Socios S on C.cedula = S.cedula" _
       & " where R.remesa = " & pRemesa _
       & " order by R.cod_institucion,S.cedula"
Call OpenRecordSet(rs, strSQL)

Open vTempo For Output As #fnFile  ' Create file name.

Do While Not rs.EOF
  vCadena = SIFGlobal.fxStringRelleno(rs!Cedula, "D", " ", 15)
  vCadena = vCadena & SIFGlobal.fxStringRelleno(rs!Nombre, "D", " ", 50)
  vCadena = vCadena & Format(rs!Monto, "000000000.00")
  
  If chkArchivoAddInst.Value = vbChecked Then
      vCadena = vCadena & "     " & Format(rs!cod_institucion, "000") & "_" & SIFGlobal.fxStringRelleno(rs!Institucion, "D", " ", 30)
  End If
    
  Print #fnFile, vCadena
  rs.MoveNext
Loop
rs.Close

Close #fnFile
  
vMensaje = "Archivo creado satisfactoriamente en : " & vTempo
  
Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
  
End Sub


Private Sub chkMostrarNoAsignado_Click()
Call cboCarga_Click
End Sub

Private Sub cmdArchivo_Click()
Dim strSQL As String, rs As New ADODB.Recordset
Dim vRemesa As Long, xRemesa As String

On Error GoTo vError

If chkRemesaInd.Value = vbChecked Then
   xRemesa = InputBox("Indique la Remesa que desea consultar", "Remesas de Liquidaciones")
  If Len(Trim(xRemesa)) = 0 Then xRemesa = "0"
  lblRemesa.Tag = xRemesa
End If

If lblRemesa.Tag = "" Then Exit Sub

Me.MousePointer = vbHourglass

vRemesa = CLng(lblRemesa.Tag)

vMensaje = ""

If chkArchivoSepInst.Value = vbChecked Then
    strSQL = "select cod_institucion,remesa from FND_CONVENIOS_REMESA_ASG" _
           & " where remesa = " & vRemesa _
           & " group by cod_institucion,remesa"
    Call OpenRecordSet(rs, strSQL)
    Do While Not rs.EOF
      Call sbArchivoInstSeparado(rs!remesa, rs!cod_institucion)
      rs.MoveNext
    Loop
    rs.Close
Else
    Call sbArchivoGeneral(vRemesa)
End If

Me.MousePointer = vbDefault

MsgBox vMensaje, vbInformation

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

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
 .WindowTitle = "Reportes del Módulo"

 .Connect = glogon.ConectRPT

If chkRemesaInd.Value = vbChecked Then
   xRemesa = InputBox("Indique la Remesa que desea consultar", "Remesas de Planes")
  If Len(Trim(xRemesa)) = 0 Then xRemesa = "0"
  lblRemesa.Tag = xRemesa
End If

 Select Case True
  Case opt.Item(0).Value 'Detalle Remesa
     .ReportFileName = SIFGlobal.fxPathReportes("Fondos_RemesasConvenioDetalle.rpt")
     vSubTitulo = "REMESA : " & lblRemesa.Tag & " LISTADO : DETALLADO"
  Case opt.Item(1).Value 'Detalle Agrupado Remesa
     .ReportFileName = SIFGlobal.fxPathReportes("Fondos_RemesasConvenioDetalleAgrupado.rpt")
     vSubTitulo = "REMESA : " & lblRemesa.Tag & " LISTADO : DETALLADO AGRUPADO"
  Case opt.Item(2).Value 'Resumen Remesa
     .ReportFileName = SIFGlobal.fxPathReportes("Fondos_RemesasConvenioResumen.rpt")
     vSubTitulo = "REMESA : " & lblRemesa.Tag & " LISTADO : RESUMEN"
 End Select
 
 .Formulas(0) = "fxFecha='FECHA: " & Format(fxFechaServidor, "dd/mm/yyyy") & "'"
 .Formulas(1) = "fxEmpresa='" & GLOBALES.gstrNombreEmpresa & "'"
 .Formulas(2) = "fxUsuario='USER: " & glogon.Usuario & "'"
 .Formulas(3) = "fxTitulo='REMESA DE LIQUIDACIONES DE FONDOS DE CONVENIOS'"
 .Formulas(4) = "fxSubTitulo='" & vSubTitulo & "'"
 .Formulas(5) = "fxFiltro='" & vFiltro & "'"
 .SelectionFormula = "{fnd_convenios_remesa.REMESA} = " & lblRemesa.Tag
 .PrintReport
  
End With

Me.MousePointer = vbDefault
Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub Form_Activate()
 vModulo = 3
End Sub

Private Sub Form_Load()
Dim strSQL As String, rs As New ADODB.Recordset
 
vModulo = 3

strSQL = "select rtrim(cod_plan) as ItemX from fnd_planes"
Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
 cboPlan.AddItem rs!itemx
 rs.MoveNext
Loop
If rs.RecordCount > 0 Then
   rs.MoveFirst
   cboPlan.Text = rs!itemx
End If
rs.Close

strSQL = "select rtrim(codigo) as ItemX from catalogo where retencion = 'S'"
Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
 cboCliente.AddItem rs!itemx
 rs.MoveNext
Loop
If rs.RecordCount > 0 Then
   rs.MoveFirst
   cboCliente.Text = rs!itemx
End If
rs.Close


'Me.Icon = Me.Picture

 ssTab.Tab = 0
 Call sbToolBarIconos(tlb, False)
 
 Call Formularios(Me)
 Call RefrescaTags(Me)
 
 Call sbLimpia
 
 
'Valor por Omision
On Error Resume Next
cboPlan.Text = "INCR"
 
End Sub


Private Sub sbConsulta(vRemesa As Long)
Dim strSQL As String, rs As New ADODB.Recordset

Call sbLimpia
  
strSQL = "select * from fnd_convenios_remesa where remesa = " & vRemesa
Call OpenRecordSet(rs, strSQL)
If Not rs.BOF And Not rs.EOF Then
  txtRemesa = rs!remesa
  txtUsuario = rs!Usuario
  txtFecha = rs!fecha
  
  cboCliente.Text = Trim(rs!cod_cliente)
  cboPlan.Text = Trim(rs!cod_Plan)
  
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
            & " where consec in (select consec from fnd_convenios_remesa_asg where remesa = " & vRemesa & ")"
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
Dim strSQL As String, rs As New ADODB.Recordset
Dim i As Integer

On Error GoTo vError

Select Case UCase(Button.Key)
  Case "NUEVO"
     
    If txtRemesa.Text = "" Then
     
            strSQL = "select isnull(max(remesa),0) + 1 as Ultimo from fnd_convenios_remesa"
            Call OpenRecordSet(rs, strSQL)
                strSQL = "insert fnd_convenios_remesa(remesa,usuario,fecha,estado,fecha_inicio,fecha_corte,notas,cod_plan,cod_cliente) values(" & rs!ultimo _
                       & ",'" & glogon.Usuario & "',dbo.MyGetdate(),'A','" & Format(dtpInicio.Value, "yyyy/mm/dd") _
                       & "','" & Format(dtpCorte.Value, "yyyy/mm/dd") & "','" & txtNotas.Text & "','" & cboPlan.Text & "','" & cboCliente.Text & "')"
                Call ConectionExecute(strSQL)
                
                txtRemesa = rs!ultimo
            rs.Close
            Call Bitacora("Registra", "Remesa de Liquidaciones de Planes de Clientes : " & txtRemesa)
    
    Else
        If txtEstado.Text <> "Remesa Cerrada" Then
                    
            strSQL = "update fnd_convenios_remesa set usuario = '" & glogon.Usuario & "',fecha_inicio = '" & Format(dtpInicio.Value, "yyyy/mm/dd") _
                   & "',fecha_corte = '" & Format(dtpCorte.Value, "yyyy/mm/dd") & "',notas = '" & txtNotas.Text _
                   & "',cod_plan = '" & cboPlan.Text & "',cod_cliente = '" & cboCliente.Text _
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
            strSQL = "delete fnd_convenios_remesa where remesa = " & txtRemesa
            Call ConectionExecute(strSQL)
            
            Call Bitacora("Elimina", "Remesa de Liquidaciones de Planes de Clientes : " & txtRemesa)
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
Dim strSQL As String, rs As New ADODB.Recordset

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
     
     
     strSQL = "select TOP 50 * from fnd_convenios_remesa order by fecha desc"
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
                itmX.SubItems(6) = rs!cod_Plan
                itmX.SubItems(7) = rs!cod_cliente
       
       End With
       rs.MoveNext
     Loop
     rs.Close
     
  Case 1 'Carga
    'Solo busca las remesas que se encuentran Abiertas
    
    vPaso = True
    
    cboCarga.Clear
    cboInstitucion.Clear
    lswCarga.ListItems.Clear
    chkCarga.Value = vbUnchecked
        
    strSQL = "select * from fnd_convenios_remesa where estado in('A','P') order by fecha desc"
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
     strSQL = "select TOP " & txtRepRemesas.Text & " * from fnd_convenios_remesa order by fecha desc"
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
                itmX.SubItems(6) = rs!cod_Plan
                itmX.SubItems(7) = rs!cod_cliente
       
       End With
       rs.MoveNext
     Loop
     rs.Close
     
 End Select

End Sub



Private Sub sbCargaBuscar()
Dim strSQL As String, rs As New ADODB.Recordset
Dim vFechaInicio As Date, vFechaCorte As Date
Dim vPlan  As String, vCliente As String
Dim curTotal As Currency

Me.MousePointer = vbHourglass

On Error GoTo vError

lswCarga.ListItems.Clear
curTotal = 0


strSQL = "select * from fnd_convenios_remesa where remesa = " & cboCarga.ItemData(cboCarga.ListIndex)
Call OpenRecordSet(rs, strSQL)
  vFechaInicio = rs!FECHA_INICIO
  vFechaCorte = rs!fecha_corte
  vPlan = rs!cod_Plan
  vCliente = rs!cod_cliente
rs.Close


'Seleccionar Casos
strSQL = "select L.consec,S.cedula,S.nombre,L.cod_plan,L.cod_operadora,L.cod_contrato,(L.aportes_liq+L.rendi_liq) as Monto" _
       & ",L.Fecha,L.usuario,L.cta_ahorros,L.tipo" _
       & " from Fnd_Liquidacion L inner join Fnd_contratos C on L.cod_operadora = C.cod_operadora" _
       & " and L.cod_plan = C.cod_plan and L.cod_contrato = C.cod_contrato" _
       & " inner join Socios S on C.cedula = S.cedula" _
       & " inner join instituciones I on S.cod_institucion = I.cod_institucion" _
       & " Where L.fecha between '" & Format(vFechaInicio, "yyyy/mm/dd") & " 00:00:00' and '" & Format(vFechaCorte, "yyyy/mm/dd") _
       & " 23:59:59' and L.consec not in (select consec from fnd_convenios_remesa_asg)" _
       & " and L.cod_plan = '" & vPlan & "'"
If chkMostrarNoAsignado.Value = vbChecked Then
  strSQL = strSQL & " and C.cedula not in(select cedula from reg_creditos where estado in('A','C'))" _
       & " and S.cod_institucion = " & cboInstitucion.ItemData(cboInstitucion.ListIndex)

Else
  strSQL = strSQL & " and C.cedula in(select cedula from reg_creditos where estado in('A','C')" _
       & " and codigo = '" & vCliente & "') and S.cod_institucion = " & cboInstitucion.ItemData(cboInstitucion.ListIndex)

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
     itmX.SubItems(10) = rs!Cod_Operadora
     
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
Dim strSQL As String, rs As New ADODB.Recordset
Dim i As Integer, vCasos As Integer
Dim vFecha As Date


On Error GoTo vError


'Valida el Estado de la Remesa
strSQL = "select count(*) as Existe from fnd_convenios_remesa" _
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
strSQL = "update fnd_convenios_remesa set estado = 'C'" _
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
Dim strSQL As String, rs As New ADODB.Recordset
Dim i As Integer, vCasos As Integer
Dim vFecha As Date


On Error GoTo vError


'Valida el Estado de la Remesa
strSQL = "select count(*) as Existe from fnd_convenios_remesa" _
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
 
     strSQL = "insert fnd_convenios_remesa_asg(remesa,consec,cod_institucion) values(" & cboCarga.ItemData(cboCarga.ListIndex) _
             & "," & .Item(i).Text & "," & cboInstitucion.ItemData(cboInstitucion.ListIndex) & ")"
     Call ConectionExecute(strSQL)
   
    prgBar.Value = prgBar.Value + 1
  End If
Next i
 
If vCasos > 0 Then
    'Actualiza el Estado de la Remesa como cerrada
    strSQL = "update fnd_convenios_remesa set estado = 'P'" _
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

Select Case Button.Key
  Case "buscar"
    

    If cboInstitucion.ListCount = 0 Then Exit Sub
    Call sbCargaBuscar
  
  Case "cargar"
    If cboInstitucion.ListCount = 0 Then Exit Sub
    If lswCarga.ListItems.Count = 0 Then Exit Sub
    Call sbCarga
  
  Case "cerrar"
    Call sbCerrar
End Select

End Sub


Private Sub sbConsultaCedula(pCedula As String)
Dim strSQL As String, rs As New ADODB.Recordset

strSQL = "select R.*,A.consec,A.cod_institucion,C.cod_Contrato,C.cod_plan,(L.Aportes_Liq + L.Rendi_Liq) as Monto,L.consec" _
       & " from fnd_convenios_remesa R inner join FND_CONVENIOS_REMESA_ASG A on R.remesa = A.remesa" _
       & " inner join fnd_liquidacion L on A.consec = L.consec" _
       & " inner join fnd_contratos C on L.cod_operadora = C.cod_operadora and L.cod_plan = C.cod_plan" _
       & " and L.cod_Contrato = C.cod_Contrato" _
       & " where C.cedula = '" & txtCedula.Text & "' order by R.fecha desc"

lswConsulta.ListItems.Clear

Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
  With lswConsulta.ListItems
       Set itmX = .Add(, , rs!remesa)
           itmX.SubItems(1) = rs!Usuario
           itmX.SubItems(2) = rs!fecha
           itmX.SubItems(3) = rs!FECHA_INICIO
           itmX.SubItems(4) = rs!fecha_corte
           itmX.SubItems(5) = rs!notas
           itmX.SubItems(6) = rs!cod_Plan
           itmX.SubItems(7) = rs!cod_cliente
           itmX.SubItems(8) = rs!cod_institucion
           itmX.SubItems(9) = rs!consec
           itmX.SubItems(10) = rs!COD_CONTRATO
           itmX.SubItems(11) = rs!cod_Plan
           itmX.SubItems(12) = Format(rs!Monto, "Standard")
  End With
  rs.MoveNext
Loop
rs.Close
     
End Sub


Private Sub txtCedula_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyF4 Then
   gBusquedas.Convertir = "N"
   gBusquedas.Columna = "Cedula"
   gBusquedas.Orden = "Cedula"
   gBusquedas.Consulta = "select Cedula,Nombre from Socios"
   frmBusquedas.Show vbModal
   
   If Trim(gBusquedas.Resultado) <> "" Then
      txtCedula = Trim(gBusquedas.Resultado)
      txtNombre = Trim(gBusquedas.Resultado2)
   End If
   txtCedula_LostFocus
   
   gBusquedas.Resultado = ""
   gBusquedas.Resultado2 = ""
End If
End Sub

Private Sub txtCedula_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
  Case 48 To 57, 8
  Case vbKeyReturn
    txtCedula_LostFocus
  Case Else
    KeyAscii = 0
End Select
End Sub

Private Sub txtCedula_LostFocus()
If Trim(txtCedula) = "" Then
   txtNombre = ""
   lswConsulta.ListItems.Clear
Else
 txtNombre.Text = fxNombre(txtCedula.Text)
 Call sbConsultaCedula(txtCedula.Text)
End If
End Sub


Private Sub txtNombre_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then
   gBusquedas.Convertir = "N"
   gBusquedas.Columna = "Nombre"
   gBusquedas.Orden = "Nombre"
   gBusquedas.Consulta = "select Cedula,Nombre from Socios"
   frmBusquedas.Show vbModal
   
   If Trim(gBusquedas.Resultado) <> "" Then
      txtCedula = Trim(gBusquedas.Resultado)
      txtNombre = Trim(gBusquedas.Resultado2)
   End If
   txtCedula_LostFocus
   
   gBusquedas.Resultado = ""
   gBusquedas.Resultado2 = ""
End If
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
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError


strSQL = "select A.* from fnd_convenios_remesa A inner join fnd_convenios_remesa_asg X on A.remesa = X.remesa where consec = " & txtRetiro.Text
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


