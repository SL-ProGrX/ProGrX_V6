VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.Controls.v24.0.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.ShortcutBar.v24.0.0.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.Ocx"
Begin VB.Form frmAF_CD_RemesasComites 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Remesas: Pago de Actividades Comites y Delegados"
   ClientHeight    =   7710
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10440
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7710
   ScaleWidth      =   10440
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   9765
      Top             =   105
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAF_CD_RemesasComites.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAF_CD_RemesasComites.frx":15172
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAF_CD_RemesasComites.frx":2A2E4
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAF_CD_RemesasComites.frx":3046E
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAF_CD_RemesasComites.frx":36CD0
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ProgressBar PrgBar 
      Align           =   2  'Align Bottom
      Height          =   150
      Left            =   0
      TabIndex        =   11
      Top             =   7560
      Width           =   10440
      _ExtentX        =   18415
      _ExtentY        =   265
      _Version        =   393216
      Appearance      =   0
   End
   Begin TabDlg.SSTab ssTab 
      Height          =   6255
      Left            =   120
      TabIndex        =   0
      Top             =   1200
      Width           =   10215
      _ExtentX        =   18018
      _ExtentY        =   11033
      _Version        =   393216
      Style           =   1
      Tabs            =   5
      TabsPerRow      =   5
      TabHeight       =   520
      ForeColor       =   16711680
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Remesas"
      TabPicture(0)   =   "frmAF_CD_RemesasComites.frx":4BE42
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Labe9(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Labe9(1)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Labe9(2)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Labe9(3)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Labe9(4)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Labe9(5)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "ShortcutCaption1"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "txtTotal"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "txtUsuario"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "txtEstado"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "txtFecha"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "txtRemesa"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "txtNotas"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "dtpFechaCorte"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "dtpFechaInicio"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "tlb"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "lswRemesas"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).ControlCount=   17
      TabCaption(1)   =   "Cargado"
      TabPicture(1)   =   "frmAF_CD_RemesasComites.frx":4BE5E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label2(10)"
      Tab(1).Control(1)=   "Label2(8)"
      Tab(1).Control(2)=   "Label2(7)"
      Tab(1).Control(3)=   "Line1(8)"
      Tab(1).Control(4)=   "Line1(9)"
      Tab(1).Control(5)=   "Label3"
      Tab(1).Control(6)=   "tlbCarga"
      Tab(1).Control(7)=   "lswCarga"
      Tab(1).Control(8)=   "chkCarga"
      Tab(1).Control(9)=   "cboCarga"
      Tab(1).Control(10)=   "cboBanco"
      Tab(1).Control(11)=   "txtCargaTotal"
      Tab(1).ControlCount=   12
      TabCaption(2)   =   "Reportes"
      TabPicture(2)   =   "frmAF_CD_RemesasComites.frx":4BE7A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label16(4)"
      Tab(2).Control(1)=   "Label16(2)"
      Tab(2).Control(2)=   "lblRemesa"
      Tab(2).Control(3)=   "opt(1)"
      Tab(2).Control(4)=   "opt(0)"
      Tab(2).Control(5)=   "txtRepRemesas"
      Tab(2).Control(6)=   "dtpRepCorte"
      Tab(2).Control(7)=   "dtpRepInicio"
      Tab(2).Control(8)=   "Frame1"
      Tab(2).Control(9)=   "cmdReporte"
      Tab(2).Control(10)=   "lswRep"
      Tab(2).ControlCount=   11
      TabCaption(3)   =   "Consultas"
      TabPicture(3)   =   "frmAF_CD_RemesasComites.frx":4BE96
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Label16(3)"
      Tab(3).Control(1)=   "Line9(1)"
      Tab(3).Control(2)=   "Label5"
      Tab(3).Control(3)=   "Label7"
      Tab(3).Control(4)=   "Label8"
      Tab(3).Control(5)=   "lblComite"
      Tab(3).Control(6)=   "txtComite"
      Tab(3).Control(7)=   "dtpConCorte"
      Tab(3).Control(8)=   "dtpConInicio"
      Tab(3).Control(9)=   "lswRemCD"
      Tab(3).Control(10)=   "PrgFecRem"
      Tab(3).ControlCount=   11
      TabCaption(4)   =   "Enviar Remesas a Tesoreria"
      TabPicture(4)   =   "frmAF_CD_RemesasComites.frx":4BEB2
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Label4"
      Tab(4).Control(1)=   "LblReme"
      Tab(4).Control(2)=   "LblRotuloR"
      Tab(4).Control(3)=   "Lbl_NRem"
      Tab(4).Control(4)=   "Label6"
      Tab(4).Control(5)=   "cmdAplicar"
      Tab(4).Control(6)=   "lswRegistroR"
      Tab(4).Control(7)=   "PrgEnvio"
      Tab(4).Control(8)=   "lswOperaciones"
      Tab(4).ControlCount=   9
      Begin XtremeSuiteControls.ListView lswRep 
         Height          =   2055
         Left            =   -74760
         TabIndex        =   59
         Top             =   840
         Width           =   9735
         _Version        =   1572864
         _ExtentX        =   17171
         _ExtentY        =   3625
         _StockProps     =   77
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
         FullRowSelect   =   -1  'True
         Appearance      =   16
         ShowBorder      =   0   'False
      End
      Begin XtremeSuiteControls.ListView lswRemesas 
         Height          =   2175
         Left            =   120
         TabIndex        =   57
         Top             =   3960
         Width           =   9975
         _Version        =   1572864
         _ExtentX        =   17595
         _ExtentY        =   3836
         _StockProps     =   77
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
         FullRowSelect   =   -1  'True
         Appearance      =   16
         ShowBorder      =   0   'False
      End
      Begin XtremeSuiteControls.PushButton cmdReporte 
         Height          =   495
         Left            =   -66600
         TabIndex        =   38
         Top             =   5640
         Width           =   1575
         _Version        =   1572864
         _ExtentX        =   2778
         _ExtentY        =   873
         _StockProps     =   79
         Caption         =   "Reporte"
         BackColor       =   -2147483633
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
         Appearance      =   17
         Picture         =   "frmAF_CD_RemesasComites.frx":4BECE
         ImageAlignment  =   4
      End
      Begin VB.Frame Frame1 
         Caption         =   "Estado de la Remesa"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2175
         Left            =   -67800
         TabIndex        =   28
         Top             =   3360
         Width           =   2775
         Begin VB.OptionButton optRemesaEstado 
            Appearance      =   0  'Flat
            Caption         =   "Trasladadas"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   3
            Left            =   240
            TabIndex        =   33
            Top             =   1440
            Width           =   2295
         End
         Begin VB.OptionButton optRemesaEstado 
            Appearance      =   0  'Flat
            Caption         =   "Cerradas"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   2
            Left            =   240
            TabIndex        =   32
            Top             =   1080
            Width           =   2295
         End
         Begin VB.OptionButton optRemesaEstado 
            Appearance      =   0  'Flat
            Caption         =   "Abiertas"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   31
            Top             =   720
            Width           =   2295
         End
         Begin VB.OptionButton optRemesaEstado 
            Appearance      =   0  'Flat
            Caption         =   "Todos"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   30
            Top             =   360
            Value           =   -1  'True
            Width           =   2295
         End
         Begin VB.CheckBox chkRemesaInd 
            Appearance      =   0  'Flat
            Caption         =   "Indicar Remesa"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   720
            TabIndex        =   29
            Top             =   1800
            Width           =   1695
         End
      End
      Begin MSComctlLib.ProgressBar PrgFecRem 
         Height          =   225
         Left            =   -72225
         TabIndex        =   26
         Top             =   5625
         Width           =   7080
         _ExtentX        =   12488
         _ExtentY        =   397
         _Version        =   393216
         Appearance      =   0
      End
      Begin MSComctlLib.ListView lswRemCD 
         Height          =   4050
         Left            =   -72240
         TabIndex        =   25
         Top             =   1560
         Width           =   7080
         _ExtentX        =   12488
         _ExtentY        =   7144
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "No.Remesa"
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   1
            Text            =   "No. Operaci�n"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   2
            Text            =   "No.Solicitud"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Monto"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   4
            Text            =   "Fecha"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Comit�"
            Object.Width           =   7832
         EndProperty
      End
      Begin MSComctlLib.ListView lswOperaciones 
         Height          =   1740
         Left            =   -74760
         TabIndex        =   16
         Top             =   3180
         Width           =   9780
         _ExtentX        =   17251
         _ExtentY        =   3069
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FlatScrollBar   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   8
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "No. Operacion"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Cedula"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Nombre"
            Object.Width           =   6174
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Cuenta"
            Object.Width           =   2822
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Monto"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Id.Banco"
            Object.Width           =   882
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Banco"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Comite"
            Object.Width           =   6068
         EndProperty
      End
      Begin MSComctlLib.ProgressBar PrgEnvio 
         Height          =   135
         Left            =   -74760
         TabIndex        =   19
         Top             =   4980
         Visible         =   0   'False
         Width           =   9795
         _ExtentX        =   17277
         _ExtentY        =   238
         _Version        =   393216
         Appearance      =   0
         Scrolling       =   1
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
         TabIndex        =   13
         Top             =   5760
         Width           =   2535
      End
      Begin VB.ComboBox cboBanco 
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
         TabIndex        =   9
         Top             =   840
         Width           =   6975
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
         TabIndex        =   4
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
         TabIndex        =   3
         Top             =   1920
         Width           =   1455
      End
      Begin MSComctlLib.Toolbar tlb 
         Height          =   330
         Left            =   4800
         TabIndex        =   2
         Top             =   600
         Width           =   1905
         _ExtentX        =   3360
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
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
         Height          =   3510
         Left            =   -74880
         TabIndex        =   5
         Top             =   2145
         Width           =   9975
         _ExtentX        =   17595
         _ExtentY        =   6191
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
            Text            =   "Op"
            Object.Width           =   1058
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Up"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Comite"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Cedula"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "Asociado"
            Object.Width           =   6174
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   5
            Text            =   "Cuenta"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Monto"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Usuario"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "Tipo"
            Object.Width           =   1058
         EndProperty
      End
      Begin MSComctlLib.Toolbar tlbCarga 
         Height          =   570
         Left            =   -71895
         TabIndex        =   6
         Top             =   1200
         Width           =   5835
         _ExtentX        =   10292
         _ExtentY        =   1005
         ButtonWidth     =   2328
         ButtonHeight    =   1005
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
               Object.ToolTipText     =   "Buscar Operaciones"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&Cargar"
               Key             =   "cargar"
               Object.ToolTipText     =   "Carga Operaciones "
               ImageIndex      =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Cerrar"
               Key             =   "cerrar"
               Object.ToolTipText     =   "Cierra Remesa"
               ImageIndex      =   4
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ListView lswRegistroR 
         Height          =   1995
         Left            =   -74760
         TabIndex        =   18
         ToolTipText     =   "Si desea ver las operaciones de la remesa precione doble click derecho"
         Top             =   840
         Width           =   9765
         _ExtentX        =   17224
         _ExtentY        =   3519
         View            =   3
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   7
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Remesa"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Fecha Proceso"
            Object.Width           =   4057
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Fecha Inicio"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Fecha Corte"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Usuario"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Text            =   "Casos"
            Object.Width           =   2011
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   6
            Text            =   "Monto"
            Object.Width           =   2540
         EndProperty
      End
      Begin XtremeSuiteControls.DateTimePicker dtpFechaInicio 
         Height          =   330
         Left            =   1800
         TabIndex        =   34
         Top             =   1320
         Width           =   1335
         _Version        =   1572864
         _ExtentX        =   2355
         _ExtentY        =   582
         _StockProps     =   68
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   3
      End
      Begin XtremeSuiteControls.DateTimePicker dtpFechaCorte 
         Height          =   330
         Left            =   3120
         TabIndex        =   35
         Top             =   1320
         Width           =   1335
         _Version        =   1572864
         _ExtentX        =   2355
         _ExtentY        =   582
         _StockProps     =   68
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   3
      End
      Begin XtremeSuiteControls.DateTimePicker dtpRepInicio 
         Height          =   330
         Left            =   -74160
         TabIndex        =   36
         Top             =   4320
         Width           =   1335
         _Version        =   1572864
         _ExtentX        =   2355
         _ExtentY        =   582
         _StockProps     =   68
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   3
      End
      Begin XtremeSuiteControls.DateTimePicker dtpRepCorte 
         Height          =   330
         Left            =   -72840
         TabIndex        =   37
         Top             =   4320
         Width           =   1335
         _Version        =   1572864
         _ExtentX        =   2355
         _ExtentY        =   582
         _StockProps     =   68
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   3
      End
      Begin XtremeSuiteControls.DateTimePicker dtpConInicio 
         Height          =   330
         Left            =   -73680
         TabIndex        =   39
         Top             =   1440
         Width           =   1335
         _Version        =   1572864
         _ExtentX        =   2355
         _ExtentY        =   582
         _StockProps     =   68
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   3
      End
      Begin XtremeSuiteControls.DateTimePicker dtpConCorte 
         Height          =   330
         Left            =   -73680
         TabIndex        =   40
         Top             =   1800
         Width           =   1335
         _Version        =   1572864
         _ExtentX        =   2355
         _ExtentY        =   582
         _StockProps     =   68
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   3
      End
      Begin XtremeSuiteControls.PushButton cmdAplicar 
         Height          =   495
         Left            =   -66600
         TabIndex        =   41
         Top             =   5520
         Width           =   1575
         _Version        =   1572864
         _ExtentX        =   2778
         _ExtentY        =   873
         _StockProps     =   79
         Caption         =   "Aplicar"
         BackColor       =   -2147483633
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
         Appearance      =   17
         Picture         =   "frmAF_CD_RemesasComites.frx":4C5D5
         ImageAlignment  =   4
      End
      Begin XtremeSuiteControls.FlatEdit txtNotas 
         Height          =   615
         Left            =   1800
         TabIndex        =   42
         Top             =   2880
         Width           =   8055
         _Version        =   1572864
         _ExtentX        =   14208
         _ExtentY        =   1085
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
      Begin XtremeSuiteControls.FlatEdit txtRemesa 
         Height          =   435
         Left            =   1800
         TabIndex        =   44
         Top             =   600
         Width           =   2655
         _Version        =   1572864
         _ExtentX        =   4683
         _ExtentY        =   767
         _StockProps     =   77
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   12
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
      Begin XtremeSuiteControls.FlatEdit txtFecha 
         Height          =   330
         Left            =   1800
         TabIndex        =   49
         Top             =   2040
         Width           =   2655
         _Version        =   1572864
         _ExtentX        =   4683
         _ExtentY        =   582
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
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtEstado 
         Height          =   330
         Left            =   1800
         TabIndex        =   50
         Top             =   1680
         Width           =   2655
         _Version        =   1572864
         _ExtentX        =   4683
         _ExtentY        =   582
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
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtUsuario 
         Height          =   330
         Left            =   4440
         TabIndex        =   51
         Top             =   2040
         Width           =   2655
         _Version        =   1572864
         _ExtentX        =   4683
         _ExtentY        =   582
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
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtTotal 
         Height          =   330
         Left            =   1800
         TabIndex        =   56
         Top             =   2520
         Width           =   2655
         _Version        =   1572864
         _ExtentX        =   4683
         _ExtentY        =   582
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
      Begin XtremeSuiteControls.FlatEdit txtRepRemesas 
         Height          =   330
         Left            =   -65880
         TabIndex        =   60
         Top             =   3000
         Width           =   855
         _Version        =   1572864
         _ExtentX        =   1508
         _ExtentY        =   582
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
         Text            =   "15"
         Alignment       =   2
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.RadioButton opt 
         Height          =   255
         Index           =   0
         Left            =   -74640
         TabIndex        =   64
         Top             =   3480
         Width           =   3855
         _Version        =   1572864
         _ExtentX        =   6800
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Detalle de Remesa"
         BackColor       =   -2147483633
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
         Value           =   -1  'True
      End
      Begin XtremeSuiteControls.RadioButton opt 
         Height          =   255
         Index           =   1
         Left            =   -74640
         TabIndex        =   43
         Top             =   3840
         Width           =   3855
         _Version        =   1572864
         _ExtentX        =   6800
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Remesa por Fecha"
         BackColor       =   -2147483633
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
      Begin XtremeSuiteControls.FlatEdit txtComite 
         Height          =   435
         Left            =   -73680
         TabIndex        =   45
         Top             =   2160
         Width           =   1335
         _Version        =   1572864
         _ExtentX        =   2355
         _ExtentY        =   767
         _StockProps     =   77
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   12
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
      Begin XtremeShortcutBar.ShortcutCaption lblComite 
         Height          =   375
         Left            =   -72240
         TabIndex        =   48
         Top             =   1200
         Width           =   7095
         _Version        =   1572864
         _ExtentX        =   12515
         _ExtentY        =   661
         _StockProps     =   14
         Caption         =   "Comit�"
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
      End
      Begin VB.Label lblRemesa 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   330
         Left            =   -74760
         TabIndex        =   63
         Top             =   3000
         Width           =   5175
      End
      Begin VB.Label Label16 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         Caption         =   "Seleccione la Remesa que Desea Visualizar"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Index           =   2
         Left            =   -74760
         TabIndex        =   62
         Top             =   480
         Width           =   9735
      End
      Begin VB.Label Label16 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         Caption         =   "Remesas - visualizar �ltimas"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Index           =   4
         Left            =   -69600
         TabIndex        =   61
         Top             =   3000
         Width           =   3735
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
         Height          =   375
         Left            =   120
         TabIndex        =   58
         Top             =   3600
         Width           =   9975
         _Version        =   1572864
         _ExtentX        =   17595
         _ExtentY        =   661
         _StockProps     =   14
         Caption         =   "Lista de Remesas:"
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
      End
      Begin XtremeSuiteControls.Label Labe9 
         Height          =   255
         Index           =   5
         Left            =   360
         TabIndex        =   55
         Top             =   2880
         Width           =   1215
         _Version        =   1572864
         _ExtentX        =   2143
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Notas:"
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label Labe9 
         Height          =   255
         Index           =   4
         Left            =   360
         TabIndex        =   54
         Top             =   2520
         Width           =   1215
         _Version        =   1572864
         _ExtentX        =   2143
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Total:"
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label Labe9 
         Height          =   255
         Index           =   3
         Left            =   360
         TabIndex        =   53
         Top             =   2040
         Width           =   1215
         _Version        =   1572864
         _ExtentX        =   2143
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Registro:"
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label Labe9 
         Height          =   255
         Index           =   2
         Left            =   360
         TabIndex        =   52
         Top             =   1680
         Width           =   1215
         _Version        =   1572864
         _ExtentX        =   2143
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Estado:"
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label Labe9 
         Height          =   255
         Index           =   1
         Left            =   360
         TabIndex        =   47
         Top             =   1320
         Width           =   1215
         _Version        =   1572864
         _ExtentX        =   2143
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Fechas:"
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label Labe9 
         Height          =   375
         Index           =   0
         Left            =   360
         TabIndex        =   46
         Top             =   600
         Width           =   1215
         _Version        =   1572864
         _ExtentX        =   2138
         _ExtentY        =   656
         _StockProps     =   79
         Caption         =   "Remesa Id:"
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Transparent     =   -1  'True
      End
      Begin VB.Label Label6 
         Caption         =   "Operaciones de Remesa"
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
         Left            =   -74760
         TabIndex        =   27
         Top             =   2955
         Width           =   2655
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Corte"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   -74805
         TabIndex        =   24
         Top             =   1800
         Width           =   510
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Inicio"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -74805
         TabIndex        =   23
         Top             =   1350
         Width           =   810
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Comit� (U.P)"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   -74805
         TabIndex        =   22
         Top             =   2265
         Width           =   1080
      End
      Begin VB.Label Lbl_NRem 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   -72675
         TabIndex        =   21
         Top             =   5265
         Visible         =   0   'False
         Width           =   225
      End
      Begin VB.Label LblRotuloR 
         Caption         =   "Procesando Remesa No."
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   -74760
         TabIndex        =   20
         Top             =   5265
         Visible         =   0   'False
         Width           =   1800
      End
      Begin VB.Label LblReme 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   -72660
         TabIndex        =   17
         Top             =   435
         Visible         =   0   'False
         Width           =   1035
      End
      Begin VB.Label Label4 
         Caption         =   "Registro de Remesas"
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
         Left            =   -74775
         TabIndex        =   15
         Top             =   450
         Width           =   1860
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
         TabIndex        =   14
         Top             =   5760
         Width           =   855
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   9
         X1              =   -74880
         X2              =   -72000
         Y1              =   1080
         Y2              =   1080
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
         Caption         =   "Seleccion de Remesas"
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
         TabIndex        =   7
         Top             =   1905
         Width           =   9975
      End
      Begin VB.Line Line9 
         BorderColor     =   &H00FFFFFF&
         Index           =   1
         X1              =   -74880
         X2              =   -65160
         Y1              =   765
         Y2              =   765
      End
      Begin VB.Label Label16 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Consultas de Remesas de Comites y Delgados"
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
         Index           =   3
         Left            =   -74880
         TabIndex        =   1
         Top             =   480
         Width           =   3975
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
         TabIndex        =   8
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
         Index           =   10
         Left            =   -74880
         TabIndex        =   10
         Top             =   840
         Width           =   2895
      End
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Remesas de Comites y Delegados"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   420
      Index           =   0
      Left            =   1080
      TabIndex        =   12
      Top             =   240
      Width           =   4005
   End
   Begin VB.Image imgBanner 
      Appearance      =   0  'Flat
      Height          =   972
      Left            =   0
      Top             =   0
      Width           =   12852
   End
End
Attribute VB_Name = "frmAF_CD_RemesasComites"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim itmX As ListItem, vPaso As Boolean
Dim strSQL As String, rs As New ADODB.Recordset
Dim vRemesa As Boolean
Dim vFechaInicio As Date, vFechaCorte As Date


Private Function fxConsecRemesaDetalle() As Long
Dim strSQL As String, rs As New ADODB.Recordset

strSQL = "Select coalesce(Max(cod_rem),0) as Consecutivo from afi_cd_remesas_tes_detalle "
Call OpenRecordSet(rs, strSQL)
    fxConsecRemesaDetalle = rs!Consecutivo + 1
rs.Close

End Function


Private Sub sbEnvio()
 
On Error GoTo vError
 
Me.MousePointer = vbHourglass
 
lswRegistroR.ListItems.Clear
 
strSQL = "select R.cod_remesa,R.Fecha,R.Fecha_Inicio,R.Fecha_Corte,R.Usuario" _
       & " ,SUM(C.monto) as 'Monto', COUNT(*) as 'Casos'" _
       & " from afi_cd_remesas_tes R inner join AFI_CD_CUENTAS C on R.COD_REMESA = C.COD_REMESA" _
       & "  where R.estado = 'C'" _
       & " group by R.cod_remesa,R.Fecha,R.Fecha_Inicio,R.Fecha_Corte,R.Usuario"
Call OpenRecordSet(rs, strSQL)
  
Do While Not rs.EOF
    Set itmX = lswRegistroR.ListItems.Add(, , rs!cod_remesa)
        itmX.SubItems(1) = rs!Fecha
        itmX.SubItems(2) = rs!FECHA_INICIO
        itmX.SubItems(3) = rs!Fecha_CORTE
        itmX.SubItems(4) = rs!Usuario
        itmX.SubItems(5) = rs!Casos
        itmX.SubItems(6) = Format(rs!Monto, "Standard")
    rs.MoveNext
Loop
rs.Close
Me.MousePointer = vbDefault

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox Err.Description, vbCritical


End Sub


Sub sbLlamaReporte()

Dim vSubTitulo As String

'On Error GoTo vError
Me.MousePointer = vbHourglass

vSubTitulo = lswRemCD.SelectedItem
'vFiltro = ""
strSQL = ""

With frmContenedor.Crt
 .Reset
 .WindowShowGroupTree = True
 .WindowShowPrintSetupBtn = True
 .WindowShowRefreshBtn = True
 .WindowShowSearchBtn = True
 .WindowState = crptMaximized
 .WindowTitle = "Reportes del Remesas de Comites y Delegados"
 
 .Connect = glogon.ConectRPT
 
 .ReportFileName = SIFGlobal.fxPathReportes("Comites_RemesasCD1.rpt")
 
 .SelectionFormula = "{afi_cd_cuentas.remesa} = " & lswRemCD.SelectedItem & ""
 
 .Formulas(0) = "fxFecha='FECHA: " & Format(fxFechaServidor, "dd/mm/yyyy") & "'"
 .Formulas(1) = "fxEmpresa='" & GLOBALES.gstrNombreEmpresa & "'"
 .Formulas(2) = "fxUsuario='USER: " & glogon.Usuario & "'"
 .Formulas(3) = "fxTitulo='REMESA DE COMITES Y DELEGADOS: PAGO DE ACTIVIDADES'"
 .Formulas(4) = "fxSubTitulo='" & vSubTitulo & "'"
 .PrintReport

End With

Me.MousePointer = vbDefault
'Exit Sub

'vError:
 'Me.MousePointer = vbDefault
 'MsgBox Err.Description, vbCritical

End Sub

Sub sbOperaciones()

strSQL = "select * from afi_cd_cuentas where remesa = '" & lswRegistroR.SelectedItem.Selected & "'"
         rs.Open , glogon.Conection, adOpenStatic
         
End Sub


Private Sub sbRemesaComite()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListItem

Me.MousePointer = vbHourglass

strSQL = "select C.cod_Remesa,C.noperacion,D.nsolicitud,C.tesoreria_fecha,rtrim(P.cod_Comite + ' - ' + P.Descripcion) as Comite,D.monto,P.cod_Comite" _
         & " from AFI_CD_COMITES P inner join afi_cd_cuentas C on P.cod_comite = C.cod_comite " _
         & " inner join afi_cd_remesas_tes_detalle D on C.cod_remesa = D.cod_remesa and C.noperacion = D.noperacion" _
         & " where C.estado in('T','L') and C.cod_comite like '%" & txtComite.Text & "%'" _
         & " and C.tesoreria_fecha between '" & Format(dtpConInicio.Value, "yyyy/mm/dd") & " 00:00:00'" _
         & " and '" & Format(dtpConCorte.Value, "yyyymmdd") & " 23:59:59'"
Call OpenRecordSet(rs, strSQL)
 
lswRemCD.ListItems.Clear
 
If Not rs.EOF And Not rs.BOF Then
  lblComite.Caption = rs!Comite
  PrgFecRem.Max = rs.RecordCount
End If
    
Do While Not rs.EOF
    Set itmX = lswRemCD.ListItems.Add(, , rs!cod_remesa)
        itmX.SubItems(1) = rs!Noperacion
        itmX.SubItems(2) = IIf(Not IsNull(rs!NSolicitud), rs!NSolicitud, 0)
        itmX.SubItems(3) = IIf(Not IsNull(rs!Monto), Format(rs!Monto, "Standard"), 0)
        itmX.SubItems(4) = rs!TESORERIA_FECHA
        itmX.SubItems(5) = rs!Comite
    rs.MoveNext
    PrgFecRem.Value = PrgFecRem.Value + 1
Loop
rs.Close

PrgFecRem.Value = 0

Me.MousePointer = vbDefault

End Sub

Private Sub cboBanco_Click()
  lswCarga.ListItems.Clear
End Sub

Private Sub cboCarga_Click()

Dim vFechaInicio As Date
Dim vFechaCorte As Date

lswCarga.ListItems.Clear

If vPaso Then Exit Sub
If cboCarga.ListCount <= 0 Then Exit Sub

vPaso = True

cboBanco.Clear


strSQL = "select fecha_inicio,fecha_corte from afi_cd_remesas_tes where " _
         & "cod_remesa = " & cboCarga.ItemData(cboCarga.ListIndex)
         Call OpenRecordSet(rs, strSQL)
         vFechaInicio = rs!FECHA_INICIO
         vFechaCorte = rs!Fecha_CORTE
         rs.Close

'Poner en true la variable vRemesa
'Seleccionar Bancos

strSQL = "select B.id_banco,B.descripcion " _
         & "from afi_cd_cuentas C inner join bancos B " _
         & "on C.id_banco = B.id_banco where " _
         & "C.registro_fecha between '" & Format(vFechaInicio, "yyyymmdd") & " 00:00:00' " _
         & "and '" & Format(vFechaCorte, "yyyymmdd") & " 23:59:59' " _
         & "group by B.id_banco,B.descripcion"
         
         Call OpenRecordSet(rs, strSQL)

Do While Not rs.EOF
  cboBanco.AddItem (Format(rs!ID_BANCO, "0000") & "..." & Trim(rs!Descripcion))
  cboBanco.ItemData(cboBanco.NewIndex) = rs!ID_BANCO
  rs.MoveNext
Loop
If rs.RecordCount > 0 Then
   rs.MoveFirst
   cboBanco.Text = (Format(rs!ID_BANCO, "0000") & "..." & Trim(rs!Descripcion))
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
       curTotal = curTotal + CCur(lswCarga.ListItems.Item(i).SubItems(6))
   End If
  
Next i

txtCargaTotal.Text = Format(curTotal, "Standard")

End Sub



Private Function fxCreaToquen() As String
Dim strSQL As String, rs As New ADODB.Recordset
Dim strToken As String

strToken = Format(fxFechaServidor, "yyyy.mm.dd")


strSQL = "select  isnull(COUNT(id_token),0)+ 1 as 'consec'  from tes_tokens where id_token like('" & strToken & "%')"
Call OpenRecordSet(rs, strSQL)

strToken = strToken & "." & rs!Consec

rs.Close

strSQL = "insert tes_tokens(id_token,registro_fecha,registro_usuario,estado)" _
      & "values('" & strToken & "',dbo.MyGetdate(),'" & glogon.Usuario & "','A') "
Call ConectionExecute(strSQL)

fxCreaToquen = strToken

End Function



Private Sub cmdAplicar_Click()
Dim i As Long, vRemesa As Long
Dim vToken As String

LblRotuloR.Visible = True
Lbl_NRem.Visible = True
PrgEnvio.Visible = True
 
 If lswRegistroR.ListItems.Count = 0 Then
   MsgBox "No hay remesas para procesar", vbInformation, "Informaci�n"
   Exit Sub
 End If

vToken = ""

On Error GoTo vError
 
 
strSQL = "select top 1 id_token from tes_tokens where estado = 'A' order by registro_fecha "
Call OpenRecordSet(rs, strSQL)
If Not rs.EOF Then
  vToken = rs!ID_TOKEN
Else
  vToken = fxCreaToquen
End If
rs.Close
 

For i = 1 To lswRegistroR.ListItems.Count
    
  If lswRegistroR.ListItems.Item(i).Checked Then
    
    vRemesa = lswRegistroR.ListItems.Item(i)
    Lbl_NRem.Caption = vRemesa

    strSQL = "exec spAFI_CD_Remesa_Desembolso " & vRemesa & ", '" & vToken & "', '" & glogon.Usuario & "'"
    Call ConectionExecute(strSQL)
          
    End If 'lswRegistroR.ListItems.Item(i).Checked

Next i 'For i = 1 To lswRegistroR.ListItems.Count

PrgEnvio.Value = 0
PrgEnvio.Visible = False

lswOperaciones.ListItems.Clear

MsgBox "Env�o de Remesas a Tesorer�a realizado satisfactoriamente!", vbInformation, "Informaci�n"

LblRotuloR.Visible = False
Lbl_NRem.Visible = False

Call sbEnvio
 
Exit Sub

vError:
    MsgBox Err.Description, vbCritical
    
End Sub

Private Sub cmdCerrar_Click()
' Lblrotulo.Visible = False
 LblReme.Visible = False
 'cmdCerrar.Visible = False
' x.Enabled = True
 lswOperaciones.Visible = False
End Sub

Private Sub cmdReporte_Click()
Dim vTitulo As String, vSubTitulo As String, vFiltro As String
Dim strSQL As String, vTipoUser As String, xRemesa As String

On Error GoTo vError

If lblRemesa.Tag = "" Then Exit Sub

Me.MousePointer = vbHourglass


vSubTitulo = lblRemesa.Caption
vFiltro = ""
strSQL = ""


With frmContenedor.Crt
 .Reset
 .WindowShowGroupTree = True
 .WindowShowPrintSetupBtn = True
 .WindowShowRefreshBtn = True
 .WindowShowSearchBtn = True
 .WindowState = crptMaximized
 .WindowTitle = "Reportes del Remesas de Comites y Delegados"
 
 .Connect = glogon.ConectRPT
 

If chkRemesaInd.Value = vbChecked Then
   xRemesa = InputBox("Indique la Remesa que desea consultar", "Remesas de Comites")
  If Len(Trim(xRemesa)) = 0 Then xRemesa = "0"
  lblRemesa.Tag = xRemesa
End If

Select Case True
  Case opt.Item(0).Value 'Detalle Remesa
     .ReportFileName = SIFGlobal.fxPathReportes("Comites_Remesas.rpt")
     .SelectionFormula = "{afi_cd_remesas_tes_detalle.cod_remesa} = " & lblRemesa.Tag & ""
  Case opt.Item(1).Value 'Detalle Remesa por fechas
     .ReportFileName = SIFGlobal.fxPathReportes("Comites_Remesas.rpt")
     vSubTitulo = "REMESA : " & lblRemesa.Tag & " LISTADO : REMESAS POR FECHAS (" _
                & Format(dtpRepInicio.Value, "dd/mm/yyyy") & " -> " & Format(dtpRepCorte.Value, "dd/mm/yyyy")
        
     strSQL = strSQL & "cdate({AFI_CD_REMESAS_TES.FECHA_INICIO}) in Date(" & Format(dtpRepInicio.Value, "yyyy,mm,dd")
     strSQL = strSQL & ") to Date (" & Format(dtpRepCorte.Value, "yyyy,mm,dd") & ")"
      .SelectionFormula = strSQL
         


 End Select
 
 .Formulas(0) = "fxFecha='Fecha: " & Format(fxFechaServidor, "dd/mm/yyyy") & "'"
 .Formulas(1) = "fxEmpresa='" & GLOBALES.gstrNombreEmpresa & "'"
 .Formulas(2) = "fxUsuario='Usuario: " & glogon.Usuario & "'"
 .Formulas(3) = "fxTitulo='REMESA DE COMITES Y DELEGADOS: PAGO DE ACTIVIDADES'"
 .Formulas(4) = "fxSubTitulo='" & vSubTitulo & "'"
 
 .Action = 1

End With

Me.MousePointer = vbDefault
Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox Err.Description, vbCritical

End Sub

Private Sub dtpConCorte_Change()
 Call sbRemesaComite
End Sub


Private Sub dtpConInicio_Change()
 Call sbRemesaComite
End Sub




Private Sub cmdReportex_Click()
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
 .WindowTitle = "Reportes del Sub.M�dulo de Comisiones de Afiliaci�n"
 
 .Connect = glogon.ConectRPT
  
' If chkRepFechas.Value = vbUnchecked Then
'    If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
'    Select Case Mid(cboRepBase.Text, 1, 1)
'      Case "R" 'Fecha de Creaci�n de la Remesa
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
'   vFiltro = vFiltro & "/ TODOS LOS BANCOS"
' End If
'
'
' If chkRepUsuario.Value = vbUnchecked Then
'   If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
'
'    Select Case Mid(cboRepBase.Text, 1, 1)
'      Case "R" 'Fecha de Creaci�n de la Remesa
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
'         .ReportFileName = App.Path & "\Reportes\AfiComisionListadoGeneral.rpt"
'     Case optReportes.Item(1).Value 'Agrupado x Promotor
'         .ReportFileName = App.Path & "\Reportes\AfiComisionAgrpPromotor.rpt"
'     Case optReportes.Item(2).Value 'Agrupado x Usuario
'         .ReportFileName = App.Path & "\Reportes\AfiComisionAgrpUsuario.rpt"
'     Case optReportes.Item(3).Value 'Agrupado x Banco
'         .ReportFileName = App.Path & "\Reportes\AfiComisionAgrpBanco.rpt"
'     Case optReportes.Item(4).Value 'Tesoreria
'         .ReportFileName = App.Path & "\Reportes\AfiComisionTesoreria.rpt"
'    End Select
' Else
'   Select Case True
'     Case optReportes.Item(0).Value 'Listado General
'         .ReportFileName = App.Path & "\Reportes\AfiComisionListadoGeneralRsm.rpt"
'     Case optReportes.Item(1).Value 'Agrupado x Promotor
'         .ReportFileName = App.Path & "\Reportes\AfiComisionAgrpPromotorRsm.rpt"
'     Case optReportes.Item(2).Value 'Agrupado x Usuario
'         .ReportFileName = App.Path & "\Reportes\AfiComisionAgrpUsuarioRsm.rpt"
'     Case optReportes.Item(3).Value 'Agrupado x Banco
'         .ReportFileName = App.Path & "\Reportes\AfiComisionAgrpBancoRsm.rpt"
'     Case optReportes.Item(4).Value 'Tesoreria
'         .ReportFileName = App.Path & "\Reportes\AfiComisionTesoreriaRsm.rpt"
'    End Select
' End If
 
 .SelectionFormula = strSQL
 .PrintReport

End With

Me.MousePointer = vbDefault
Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox Err.Description, vbCritical

End Sub

Private Sub Form_Activate()
 vModulo = 40
End Sub

Private Sub Form_Load()
 
 vModulo = 40
'vRemesa = False

 Set imgBanner.Picture = frmContenedor.imgBanner_01.Picture
 
 ssTab.Tab = 0
 
 dtpRepCorte.Value = fxFechaServidor
 dtpRepInicio.Value = DateAdd("d", -30, dtpRepCorte.Value)
 
 dtpConCorte.Value = dtpRepCorte.Value
 dtpConInicio.Value = dtpRepInicio.Value
 
  With lswRemesas.ColumnHeaders
    .Clear
    .Add , , "Remesa Id", 1400
    .Add , , "Usuario", 1800
    .Add , , "Fecha", 2100
    .Add , , "Estado", 1400
    .Add , , "Inicio", 1400
    .Add , , "Corte", 1400
    .Add , , "Notas", 3400
 End With
 
  
 With lswRep.ColumnHeaders
    .Clear
    .Add , , "Remesa Id", 1400
    .Add , , "Usuario", 1800
    .Add , , "Fecha", 2100
    .Add , , "Inicio", 1400
    .Add , , "Corte", 1400
    .Add , , "Notas", 3400
 End With

 
 Call sbToolBarIconos(tlb, False)
 
 Call Formularios(Me)
 Call RefrescaTags(Me)
 
 Call sbLimpia

dtpConInicio.Value = fxFechaServidor
dtpConCorte.Value = dtpConInicio.Value

End Sub


Private Sub sbConsulta(vRemesa As Long, Optional vTodo As Boolean = True)

Call sbLimpia(vTodo)
  
strSQL = "select T.cod_remesa,T.usuario,T.fecha,T.fecha_corte,T.notas,T.estado,T.fecha_inicio,isnull(sum(D.monto),0) as Total" _
         & " from afi_cd_remesas_tes T left join afi_cd_remesas_tes_detalle D  on T.cod_remesa = D.cod_remesa " _
         & " Where T.cod_remesa = " & vRemesa _
         & " group by T.cod_remesa,T.usuario,T.fecha,T.fecha_corte,T.notas,T.estado,T.fecha_inicio"
Call OpenRecordSet(rs, strSQL)

If Not rs.BOF And Not rs.EOF Then
  txtRemesa.Text = rs!cod_remesa
  txtUsuario.Text = rs!Usuario
  txtFecha.Text = rs!Fecha
  
  
  Select Case rs!Estado
    Case "A"
      txtEstado.Text = "Remesa Abierta"
    Case "C"
      txtEstado.Text = "Remesa Cerrada"
    Case "P"
      txtEstado.Text = "Remesa en Proceso"
    Case "T"
      txtEstado.Text = "Remesa Transferida a Tesorer�a"
  End Select
  
  
  dtpFechaInicio.Value = rs!FECHA_INICIO
  dtpFechaCorte.Value = rs!Fecha_CORTE
  txtNotas.Text = rs!NOTAS
  
  While Not rs.EOF
     txtTotal.Text = Format(CCur(txtTotal.Text) + CCur(rs!Total), "Standard")
  rs.MoveNext
  Wend
  
  
End If
rs.Close


End Sub

Private Sub lswCarga_ItemCheck(ByVal Item As MSComctlLib.ListItem)
Dim curTotal As Currency

If Trim(txtCargaTotal.Text) = "" Then txtCargaTotal.Text = 0

curTotal = CCur(txtCargaTotal.Text)

If Item.Checked Then
   curTotal = curTotal + CCur(Item.SubItems(6))
Else
   curTotal = txtCargaTotal.Text - CCur(Item.SubItems(6))
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
   MsgBox "Ocurri� un error al ordenar los datos de la columna seleccionada.", vbCritical

End Sub

Private Sub LswRegistroR_DblClick()

Dim i As Integer
Dim vSubTotal As Currency

'Lblrotulo.Visible = True
LblReme.Visible = True
'CmdCerrar.Visible = True
lswOperaciones.Visible = True
'x.Enabled = False

If lswRegistroR.SelectedItem.Selected = True Then
  
  strSQL = "select C.noperacion,C.cedula,S.nombre,C.cuenta,C.Monto" _
         & ",C.id_banco,B.descripcion as 'Banco',rtrim(Com.Cod_Comite) + ' - ' + Com.Descripcion as 'Comite'" _
         & " from afi_cd_cuentas C inner join  Tes_bancos B on C.id_banco = B.id_banco " _
         & " inner join socios S on C.cedula = S.cedula " _
         & " inner join Afi_Cd_Comites Com on C.cod_Comite = Com.cod_comite" _
         & " Where C.cod_remesa = " & lswRegistroR.SelectedItem
  Call OpenRecordSet(rs, strSQL)
  
  lswOperaciones.ListItems.Clear
    
  Do While Not rs.EOF
       Set itmX = lswOperaciones.ListItems.Add(, , rs!Noperacion)
           itmX.SubItems(1) = rs!Cedula
           itmX.SubItems(2) = rs!Nombre
           itmX.SubItems(3) = rs!Cuenta
           itmX.SubItems(4) = Format(rs!Monto, "Standard")
           itmX.SubItems(5) = rs!ID_BANCO
           itmX.SubItems(6) = rs!Banco
           itmX.SubItems(7) = rs!Comite
   
    rs.MoveNext
  Loop
  rs.Close
 
 
' For i = 1 To lswOperaciones.ListItems.Count
'   vSubTotal = 0
'   strSQL = "select C.cod_remesa,j.monto,C.noperacion from afi_cd_cuentas_actividades j inner join afi_cd_cuentas C " _
'            & "on J.noperacion = C.noperacion " _
'            & "where c.noperacion = '" & lswOperaciones.ListItems.Item(i) & "' " _
'            & "and cod_remesa = '" & lswRegistroR.SelectedItem & "'"
'             call OpenRecordSet(rs, strsql)
'             While Not rs.EOF
'               vSubTotal = vSubTotal + rs!Monto
'               lswOperaciones.ListItems.Item(i).SubItems(4) = Format(vSubTotal, "Standard")
'               rs.MoveNext
'             Wend
'   rs.Close
' Next i
  
  LblReme.Caption = ""
  LblReme.Caption = lswRegistroR.SelectedItem
End If
End Sub


Private Sub LswRemCD_DblClick()
 Call sbLlamaReporte
End Sub

Private Sub LswRemCD_KeyPress(KeyAscii As Integer)

Select Case KeyAscii
      Case 48 To 57, 8
      Case 13
        Call sbLlamaReporte
      Case Else
       KeyAscii = 0
End Select
End Sub

Private Sub lswRemesas_ColumnClick(ByVal ColumnHeader As XtremeSuiteControls.ListViewColumnHeader)
 lswRemesas.SortKey = ColumnHeader.Index - 1
  If lswRemesas.SortOrder = 0 Then lswRemesas.SortOrder = 1 Else lswRemesas.SortOrder = 0
  lswRemesas.Sorted = True
End Sub

Private Sub lswRemesas_ItemClick(ByVal Item As XtremeSuiteControls.ListViewItem)
    If lswRemesas.ListItems.Count <= 0 Then Exit Sub
    Call sbConsulta(Item.Text)
End Sub

Private Sub lswRep_ColumnClick(ByVal ColumnHeader As XtremeSuiteControls.ListViewColumnHeader)
 lswRep.SortKey = ColumnHeader.Index - 1
  If lswRep.SortOrder = 0 Then lswRep.SortOrder = 1 Else lswRep.SortOrder = 0
  lswRep.Sorted = True
End Sub

Private Sub lswRep_ItemClick(ByVal Item As XtremeSuiteControls.ListViewItem)
If lswRep.ListItems.Count <= 0 Then Exit Sub

lblRemesa.Caption = Item.Text & " � " & Item.SubItems(1) _
            & " � " & Item.SubItems(2)
lblRemesa.Tag = Item.Text
End Sub

Private Sub optRemesaEstado_Click(Index As Integer)
    Call sbLimpia
End Sub

Private Sub ssTab_Click(PreviousTab As Integer)
  
 Call sbLimpia
 If ssTab.Tab = 4 Then
    Call sbEnvio
 End If

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
 .WindowTitle = "Reportes del M�dulo de Planes de Ahorro"
 
 .Connect = glogon.ConectRPT
  
 
 .Formulas(0) = "fxFecha='FECHA: " & Format(fxFechaServidor, "dd/mm/yyyy") & "'"
 .Formulas(1) = "fxEmpresa='" & GLOBALES.gstrNombreEmpresa & "'"
 .Formulas(2) = "fxUsuario='USER: " & glogon.Usuario & "'"
 .Formulas(3) = "fxSubTitulo='" & vSubTitulo & "'"
 .Formulas(4) = "fxFiltro='" & vFiltro & "'"
 
 .ReportFileName = SIFGlobal.fxPathReportes("FndRemesas.rpt")
 .PrintReport

End With

Me.MousePointer = vbDefault
Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox Err.Description, vbCritical

End Sub

Private Sub tlb_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim i As Integer

On Error GoTo vError

Select Case UCase(Button.Key)
  Case "NUEVO"
     
    If txtRemesa.Text = "" Then
     
            strSQL = "select coalesce(max(cod_remesa),0) + 1 as Ultimo from afi_cd_remesas_tes"
            Call OpenRecordSet(rs, strSQL)
                strSQL = "insert afi_cd_remesas_tes(cod_remesa,usuario,fecha,estado,fecha_inicio,fecha_corte,notas) " _
                        & "values(" & rs!ultimo _
                       & ",'" & glogon.Usuario & "',getdate(),'A','" & Format(dtpFechaInicio.Value, "yyyymmdd") _
                       & "','" & Format(dtpFechaCorte.Value, "yyyymmdd") & "','" & txtNotas.Text & "')"
                       Call ConectionExecute(strSQL)
                
                txtRemesa = rs!ultimo
            rs.Close
            Call Bitacora("Registra", "Remesa de Comites y Delegados : " & txtRemesa)
    Else
        If txtEstado.Text <> "Remesa Cerrada" Then
                    
            strSQL = "update afi_cd_remesas_tes set usuario = '" & glogon.Usuario & "', " _
                   & "fecha_inicio = '" & Format(dtpFechaInicio.Value, "yyyymmdd") _
                   & "',fecha_corte = '" & Format(dtpFechaCorte.Value, "yyyymmdd") & "', " _
                   & "notas = '" & txtNotas.Text _
                   & "' where cod_remesa = " & txtRemesa
             Call ConectionExecute(strSQL)
            
            Call Bitacora("Modifica", "Remesa de Comites y Delegados : " & txtRemesa)
        
        Else
            MsgBox "No se puede Modifica la remesa, porque esta ya fue cerrada...", vbExclamation
        End If
    End If
    
    Call sbLimpia
    
    
  Case "BORRAR"
     i = MsgBox("Esta Seguro que desea borrar este registro", vbYesNo)
     If i = vbYes Then
        If txtEstado.Text = "Remesa Abierta" Then
            strSQL = "delete afi_cd_remesas_tes where cod_remesa = " & txtRemesa
            Call ConectionExecute(strSQL)
            
             Call Bitacora("Elimina", "Remesa de Comites y Delegados : " & txtRemesa)
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
 MsgBox Err.Description, vbCritical

End Sub



Private Sub sbLimpia(Optional vTodo As Boolean = True)

Dim itmX As ListViewItem


On Error GoTo vError

Me.MousePointer = vbHourglass


Select Case ssTab.Tab
  
  Case 0 'Remesas
     txtEstado.Text = ""
     txtFecha.Text = ""
     txtTotal.Text = "0.00"
     txtUsuario.Text = ""
     txtRemesa.Text = ""
     txtNotas.Text = ""
     
     dtpFechaInicio.Value = fxFechaServidor
     dtpFechaCorte.Value = dtpFechaInicio.Value
     
     If vTodo Then
             strSQL = "select TOP 50 * from afi_cd_remesas_tes order by fecha desc"
             lswRemesas.ListItems.Clear
        
             Call OpenRecordSet(rs, strSQL)
             Do While Not rs.EOF
               With lswRemesas.ListItems
                    Set itmX = .Add(, , rs!cod_remesa)
                        itmX.SubItems(1) = rs!Usuario
                        itmX.SubItems(2) = rs!Fecha
                        itmX.SubItems(3) = rs!FECHA_INICIO
                        itmX.SubItems(4) = rs!Fecha_CORTE
                        itmX.SubItems(5) = rs!NOTAS
               
               End With
               rs.MoveNext
             Loop
             rs.Close
     End If
  Case 1 'Carga
    'Solo busca las remesas que se encuentran Abiertas
    
    vPaso = True
    
    cboCarga.Clear
    cboBanco.Clear
    lswCarga.ListItems.Clear
    chkCarga.Value = vbUnchecked
        
    strSQL = "select * from afi_cd_remesas_tes where estado in('A','P') order by fecha desc"
    Call OpenRecordSet(rs, strSQL)
    Do While Not rs.EOF
      cboCarga.AddItem (Format(rs!cod_remesa, "0000") & "..." & Trim(rs!Usuario) & "..." & rs!Fecha)
      cboCarga.ItemData(cboCarga.NewIndex) = rs!cod_remesa
      rs.MoveNext
    Loop
    If rs.RecordCount > 0 Then
       rs.MoveFirst
       cboCarga.Text = (Format(rs!cod_remesa, "0000") & "..." & Trim(rs!Usuario) & "..." & rs!Fecha)
    End If
    
    rs.Close

    vPaso = False
    Call cboCarga_Click

  Case 2 'Reportes
     strSQL = "select TOP " & txtRepRemesas.Text & " * from afi_cd_remesas_tes"
     
     Select Case True
        Case optRemesaEstado.Item(1).Value 'Abiertas
            strSQL = strSQL & " where Estado = 'A'"
        Case optRemesaEstado.Item(2).Value 'Cerradas
            strSQL = strSQL & " where Estado = 'C'"
        Case optRemesaEstado.Item(3).Value 'Trasladadas
            strSQL = strSQL & " where Estado = 'T'"
        Case Else
          'Todos
     End Select
     
     strSQL = strSQL & " order by fecha desc"
     lswRep.ListItems.Clear

     Call OpenRecordSet(rs, strSQL)
     Do While Not rs.EOF
       With lswRep.ListItems
            Set itmX = .Add(, , rs!cod_remesa)
                itmX.SubItems(1) = rs!Usuario
                itmX.SubItems(2) = rs!Fecha
                itmX.SubItems(3) = rs!FECHA_INICIO
                itmX.SubItems(4) = rs!Fecha_CORTE
                itmX.SubItems(5) = rs!NOTAS
       
       End With
       rs.MoveNext
     Loop
     rs.Close
     
 End Select

Me.MousePointer = vbDefault

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub



Private Sub sbCargaBuscar()
Dim i As Integer
Dim curTotal As Currency, vSubTotal As Currency
Dim vOperacion As Integer

Me.MousePointer = vbHourglass

On Error GoTo vError

lswCarga.ListItems.Clear
curTotal = 0
txtCargaTotal.Text = 0

strSQL = "select fecha_inicio,fecha_corte from afi_cd_remesas_tes " _
         & "where cod_remesa = " & cboCarga.ItemData(cboCarga.ListIndex) & ""
         Call OpenRecordSet(rs, strSQL)
         vFechaInicio = rs!FECHA_INICIO
         vFechaCorte = rs!Fecha_CORTE
rs.Close


'Seleccionar Bancos

strSQL = "select C.noperacion,C.cod_comite,P.descripcion,C.cedula, " _
         & "S.nombre as Asociado,C.cuenta,C.registro_usuario,C.tipo " _
         & "from uprogramatica P inner join afi_cd_cuentas C " _
         & "on P.codigo = C.cod_comite inner join Socios S " _
         & "on C.cedula = S.cedula " _
         & "Where C.id_banco = " & cboBanco.ItemData(cboBanco.ListIndex) & " " _
         & "and C.registro_fecha between " _
         & "'" & Format(vFechaInicio, "yyyymmdd") & " 00:00:00' " _
         & "and '" & Format(vFechaCorte, "yyyymmdd") & " 23:59:59' " _
         & "and C.estado = 'A'"
         Call OpenRecordSet(rs, strSQL)


PrgBar.Max = rs.RecordCount + 1
PrgBar.Value = 1
PrgBar.Visible = True

Do While Not rs.EOF
 
 Set itmX = lswCarga.ListItems.Add(, , rs!Noperacion)
     itmX.SubItems(1) = rs!cod_comite
     itmX.SubItems(2) = rs!Descripcion
     itmX.SubItems(3) = rs!Cedula
     itmX.SubItems(4) = rs!asociado
     itmX.SubItems(5) = rs!Cuenta
     itmX.SubItems(7) = rs!Registro_Usuario
     itmX.SubItems(8) = Trim(rs!Tipo)
     itmX.Checked = chkCarga.Value
     vOperacion = rs!Noperacion
     
  
' 'Validar Remesa - asignar el numero de remesa a las operaciones pendiente.
'
'  strSQL = " update afi_cd_cuentas set remesa = " & cboCarga.ItemData(cboCarga.ListIndex) & " " _
'             & "Where noperacion = " & vOperacion & ""
'             CALL ConectionExecute(strSQL)

 rs.MoveNext
 PrgBar.Value = PrgBar.Value + 1
Loop
rs.Close

For i = 1 To lswCarga.ListItems.Count
   vSubTotal = 0
   strSQL = "select * from afi_cd_cuentas_actividades where noperacion = '" & lswCarga.ListItems.Item(i) & "'"
             rs.Open strSQL, glogon.Conection, adOpenForwardOnly
             While Not rs.EOF
               vSubTotal = vSubTotal + rs!Monto
               lswCarga.ListItems.Item(i).SubItems(6) = Format(vSubTotal, "Standard")
               rs.MoveNext
             Wend
   rs.Close
Next i

'curTotal = curTotal + vSubTotal
PrgBar.Visible = False
'txtCargaTotal.Text = Format(curTotal, "Standard")

Me.MousePointer = vbDefault

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox Err.Description, vbCritical
 lswCarga.ListItems.Clear

End Sub

Private Sub sbCerrar()
Dim i As Integer, vCasos As Integer
Dim vFecha As Date

On Error GoTo vError

'Valida el Estado de la Remesa

strSQL = "select estado from afi_cd_remesas_tes" _
       & " where cod_remesa = " & cboCarga.ItemData(cboCarga.ListIndex) _
       & " and estado = 'C'"
Call OpenRecordSet(rs, strSQL)
 
 If Not rs.EOF Then
    rs.Close
    MsgBox "La Remesa actual; ya se encuentra cerrada...", vbExclamation
    Exit Sub
 End If
rs.Close

'Actualiza el Estado de la Remesa como cerrada

strSQL = "update afi_cd_remesas_tes set estado = 'C'" _
       & " where cod_remesa = " & cboCarga.ItemData(cboCarga.ListIndex)
Call ConectionExecute(strSQL)


Call Bitacora("Aplica", "CIERRE -> Remesa Comites y Delegados : " & cboCarga.ItemData(cboCarga.ListIndex))


'Actualiza las operaci�n de la remesa en proceso para ponerla como cerrada

strSQL = "update afi_cd_cuentas set estado = 'C' " _
         & "where cod_remesa = " & cboCarga.ItemData(cboCarga.ListIndex) & ""
         Call ConectionExecute(strSQL)

MsgBox "Remesa Cerrada Satisfactoriamente...", vbInformation
Call sbLimpia

Exit Sub
vError:
 Me.MousePointer = vbDefault
 MsgBox Err.Description, vbCritical
 lswCarga.ListItems.Clear

End Sub

Private Sub sbCarga()
Dim i As Integer, vCasos As Integer


On Error GoTo vError

'Valida el Estado de la Remesa
strSQL = "select estado from afi_cd_remesas_tes" _
       & " where cod_remesa = " & cboCarga.ItemData(cboCarga.ListIndex) _
       & " and estado = 'T'"
Call OpenRecordSet(rs, strSQL)
 
 If Not rs.EOF Then
    rs.Close
    MsgBox "La Remesa actual; ya se encuentra cerrada...", vbExclamation
    Exit Sub
 End If

Me.MousePointer = vbHourglass


strSQL = "select cod_remesa from afi_cd_cuentas where cod_remesa = " & cboCarga.ItemData(cboCarga.ListIndex) _
         & " and id_banco = " & cboBanco.ItemData(cboBanco.ListIndex)
Call OpenRecordSet(rs, strSQL)


'Validar Estado de la operacion - Pone la operaci�n en proceso
'Calcula los casos a procesar

Dim pNotas As String

pNotas = "Cargado en Remesa: " & cboCarga.ItemData(cboCarga.ListIndex)
strSQL = ""


'spAFI_CD_Cuenta_Remesa(@Operacion int, @RemesaId int, @Estado varchar(10), @Usuario varchar(30), @Notas varchar(300) = '')
For i = 1 To lswCarga.ListItems.Count
 
 If lswCarga.ListItems.Item(i).Checked Then
    
    strSQL = strSQL & Space(10) & "exec spAFI_CD_Cuenta_Remesa " & lswCarga.ListItems.Item(i) & ", " & cboCarga.ItemData(cboCarga.ListIndex) _
           & ", 'C', '" & glogon.Usuario & "', '" & pNotas & "'"
    If Len(strSQL) > 20000 Then
       Call ConectionExecute(strSQL)
       strSQL = ""
    End If
    
    vCasos = vCasos + 1
 End If

Next i

If Len(strSQL) > 0 Then
   Call ConectionExecute(strSQL)
End If

If vCasos > 0 Then
    
 'Actualiza el Estado de la Remesa en Proceso
   
    strSQL = "update afi_cd_remesas_tes set estado = 'P'" _
           & " where cod_remesa = " & cboCarga.ItemData(cboCarga.ListIndex)
           Call ConectionExecute(strSQL)
    Call Bitacora("Genera", "Remesa de Comites y Delegados : " & cboCarga.ItemData(cboCarga.ListIndex))

End If

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


Private Sub txtComite_KeyPress(KeyAscii As Integer)

Select Case KeyAscii
      Case 48 To 57, 8
      Case 13
        Call sbRemesaComite
      Case Else
       KeyAscii = 0
    End Select

End Sub


Private Sub txtRemesa_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn And IsNumeric(txtRemesa) Then
   Call sbConsulta(txtRemesa)
End If
End Sub

Private Function fxCDParametros(vParametro) As String
Dim rsX As New ADODB.Recordset

On Error GoTo vError

With glogon
 .strSQL = "select valor from AFI_CD_PARAMETROS where cod_parametro = '" & vParametro & "'"
 Call OpenRecordSet(rsX, .strSQL)
   fxCDParametros = rsX!Valor
 rsX.Close

End With

Exit Function

vError:
 MsgBox Err.Description, vbCritical

End Function

Private Sub txtRepRemesas_Change()
Call sbLimpia
End Sub


