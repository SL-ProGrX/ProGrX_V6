VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.Ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.0#0"; "Codejock.Controls.v22.0.0.ocx"
Begin VB.Form frmSeguros_RemesaPago 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Seguros: Remesa de Pagos"
   ClientHeight    =   7665
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   10425
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7665
   ScaleWidth      =   10425
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   8640
      Top             =   0
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
            Picture         =   "frmSeguros_RemesaPago.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSeguros_RemesaPago.frx":169C2
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSeguros_RemesaPago.frx":2D384
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSeguros_RemesaPago.frx":424F6
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSeguros_RemesaPago.frx":57668
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgToolbarIcons 
      Left            =   9360
      Top             =   0
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
            Picture         =   "frmSeguros_RemesaPago.frx":6E02A
            Key             =   "New"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSeguros_RemesaPago.frx":6E13C
            Key             =   "Properties"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSeguros_RemesaPago.frx":6E24E
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSeguros_RemesaPago.frx":6E360
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSeguros_RemesaPago.frx":6E472
            Key             =   "Undo"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSeguros_RemesaPago.frx":6E584
            Key             =   "Find"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSeguros_RemesaPago.frx":6E696
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSeguros_RemesaPago.frx":6E7A8
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSeguros_RemesaPago.frx":6E8BA
            Key             =   "Open"
         EndProperty
      EndProperty
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
      Tabs            =   4
      TabsPerRow      =   4
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
      TabCaption(0)   =   "Remesa"
      TabPicture(0)   =   "frmSeguros_RemesaPago.frx":6E9CC
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label2(20)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label2(8)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label2(2)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label2(3)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label2(4)"
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
      Tab(0).Control(10)=   "Line1(16)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Line1(15)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Line1(1)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Line1(2)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Line1(3)"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Line1(4)"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Line2(1)"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Label2(1)"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Line1(12)"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "cboTipo"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "dtpCorte"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "dtpInicio"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "tlb"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "lswRemesas"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "txtNotas"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "txtRemesa"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "txtUsuario"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "txtFecha"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "txtEstado"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).ControlCount=   29
      TabCaption(1)   =   "Cargar"
      TabPicture(1)   =   "frmSeguros_RemesaPago.frx":6E9E8
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cboVendedor"
      Tab(1).Control(1)=   "txtCargaTotal"
      Tab(1).Control(2)=   "cboCarga"
      Tab(1).Control(3)=   "chkCarga"
      Tab(1).Control(4)=   "tlbCarga"
      Tab(1).Control(5)=   "lswCarga"
      Tab(1).Control(6)=   "Label3(2)"
      Tab(1).Control(7)=   "Line1(18)"
      Tab(1).Control(8)=   "Label2(21)"
      Tab(1).Control(9)=   "Line1(5)"
      Tab(1).Control(10)=   "Label2(0)"
      Tab(1).Control(11)=   "Label2(22)"
      Tab(1).ControlCount=   12
      TabCaption(2)   =   "Traslado (Pago)"
      TabPicture(2)   =   "frmSeguros_RemesaPago.frx":6EA04
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "chkTrasladoAgrupar"
      Tab(2).Control(1)=   "txtPagoTotal"
      Tab(2).Control(2)=   "cboTraslado"
      Tab(2).Control(3)=   "tlbTraslado"
      Tab(2).Control(4)=   "lswTraslado"
      Tab(2).Control(5)=   "Line1(11)"
      Tab(2).Control(6)=   "Label2(16)"
      Tab(2).Control(7)=   "Label3(4)"
      Tab(2).Control(8)=   "Label2(24)"
      Tab(2).ControlCount=   9
      TabCaption(3)   =   "Reportes"
      TabPicture(3)   =   "frmSeguros_RemesaPago.frx":6EA20
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "cmdReporte"
      Tab(3).Control(1)=   "Frame3"
      Tab(3).Control(2)=   "Frame2"
      Tab(3).Control(3)=   "Frame1"
      Tab(3).Control(4)=   "txtRepRemesas"
      Tab(3).Control(5)=   "lswRep"
      Tab(3).Control(6)=   "Label16(4)"
      Tab(3).Control(7)=   "lblRemesa"
      Tab(3).Control(8)=   "Label16(2)"
      Tab(3).ControlCount=   9
      Begin XtremeSuiteControls.PushButton cmdReporte 
         Height          =   615
         Left            =   -66360
         TabIndex        =   71
         Top             =   5400
         Width           =   1335
         _Version        =   1441792
         _ExtentX        =   2355
         _ExtentY        =   1085
         _StockProps     =   79
         Caption         =   "Reporte"
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
         Picture         =   "frmSeguros_RemesaPago.frx":6EA3C
         ImageAlignment  =   4
      End
      Begin VB.Frame Frame3 
         Caption         =   "Tipo de Remesa"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   -74760
         TabIndex        =   63
         Top             =   480
         Width           =   9735
         Begin VB.OptionButton optTipo 
            Appearance      =   0  'Flat
            Caption         =   "Aseguradora (Pagos)"
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
            Height          =   570
            Index           =   1
            Left            =   7680
            TabIndex        =   65
            Top             =   180
            Width           =   1695
         End
         Begin VB.OptionButton optTipo 
            Appearance      =   0  'Flat
            Caption         =   "Vendedores (Comisiones)"
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
            Height          =   570
            Index           =   0
            Left            =   5640
            TabIndex        =   64
            Top             =   180
            Value           =   -1  'True
            Width           =   1695
         End
      End
      Begin VB.Frame Frame2 
         Height          =   1935
         Left            =   -74760
         TabIndex        =   60
         Top             =   4080
         Width           =   4095
         Begin VB.OptionButton opt 
            Appearance      =   0  'Flat
            Caption         =   "(Traslado) Detalle Agrupado de Remesa"
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
            Left            =   120
            TabIndex        =   62
            Top             =   600
            Width           =   3615
         End
         Begin VB.OptionButton opt 
            Appearance      =   0  'Flat
            Caption         =   "(Pendientes) Detalle de Remesa"
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
            Left            =   120
            TabIndex        =   61
            Top             =   240
            Value           =   -1  'True
            Width           =   3855
         End
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
         Height          =   1935
         Left            =   -70560
         TabIndex        =   54
         Top             =   4080
         Width           =   3855
         Begin VB.CheckBox chkRemesaInd 
            Appearance      =   0  'Flat
            Caption         =   "Indicar Remesa"
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
            Height          =   255
            Left            =   2040
            TabIndex        =   59
            Top             =   360
            Width           =   1455
         End
         Begin VB.OptionButton optRemesaEstado 
            Appearance      =   0  'Flat
            Caption         =   "Todos"
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
            Left            =   240
            TabIndex        =   58
            Top             =   360
            Value           =   -1  'True
            Width           =   2055
         End
         Begin VB.OptionButton optRemesaEstado 
            Appearance      =   0  'Flat
            Caption         =   "Abiertas"
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
            Left            =   240
            TabIndex        =   57
            Top             =   720
            Width           =   2055
         End
         Begin VB.OptionButton optRemesaEstado 
            Appearance      =   0  'Flat
            Caption         =   "Cerradas"
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
            Index           =   2
            Left            =   240
            TabIndex        =   56
            Top             =   1080
            Width           =   2055
         End
         Begin VB.OptionButton optRemesaEstado 
            Appearance      =   0  'Flat
            Caption         =   "Trasladadas"
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
            Index           =   3
            Left            =   240
            TabIndex        =   55
            Top             =   1440
            Width           =   2055
         End
      End
      Begin VB.ComboBox cboVendedor 
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
         TabIndex        =   52
         Top             =   720
         Width           =   6975
      End
      Begin VB.CheckBox chkTrasladoAgrupar 
         Appearance      =   0  'Flat
         Caption         =   "Agrupar por Fecha de Vencimiento ?"
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
         Height          =   375
         Left            =   -68880
         TabIndex        =   12
         Top             =   960
         Width           =   3855
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
         TabIndex        =   11
         Top             =   5760
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
         TabIndex        =   10
         Top             =   480
         Width           =   6975
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
         Top             =   2400
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
         TabIndex        =   8
         Top             =   2040
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
         TabIndex        =   7
         Top             =   1680
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
         TabIndex        =   6
         Top             =   480
         Width           =   1455
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
         TabIndex        =   5
         Top             =   5760
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
         TabIndex        =   4
         Top             =   360
         Width           =   6975
      End
      Begin VB.CheckBox chkCarga 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
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
         Top             =   1575
         Value           =   1  'Checked
         Width           =   1455
      End
      Begin VB.TextBox txtRepRemesas 
         Alignment       =   2  'Center
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
         Left            =   -65640
         TabIndex        =   2
         Text            =   "15"
         Top             =   3600
         Width           =   615
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
         TabIndex        =   1
         Top             =   2880
         Width           =   6975
      End
      Begin MSComctlLib.Toolbar tlbTraslado 
         Height          =   330
         Left            =   -72000
         TabIndex        =   13
         Top             =   960
         Width           =   2625
         _ExtentX        =   4630
         _ExtentY        =   582
         ButtonWidth     =   2143
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
         TabIndex        =   14
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
         BorderStyle     =   1
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
         Left            =   8160
         TabIndex        =   15
         Top             =   960
         Width           =   1665
         _ExtentX        =   2937
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
         TabIndex        =   16
         Top             =   1080
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
         Height          =   1935
         Left            =   -74760
         TabIndex        =   17
         Top             =   1680
         Width           =   9735
         _ExtentX        =   17171
         _ExtentY        =   3413
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         HotTracking     =   -1  'True
         HoverSelection  =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
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
            Text            =   "Tipo"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComctlLib.ListView lswCarga 
         Height          =   3855
         Left            =   -74880
         TabIndex        =   18
         Top             =   1800
         Width           =   9975
         _ExtentX        =   17595
         _ExtentY        =   6800
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
         BorderStyle     =   1
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
         NumItems        =   12
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "# P�liza"
            Object.Width           =   3775
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   1
            Text            =   "No. Cuota"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Cuota"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Comisi�n Interna"
            Object.Width           =   3422
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "Comisi�n Vendedor"
            Object.Width           =   3422
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Pago "
            Object.Width           =   6068
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   6
            Text            =   "Aprobado"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   7
            Text            =   "A Girar"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   8
            Text            =   "# Desem"
            Object.Width           =   2011
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   9
            Text            =   "Desembolsos"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   10
            Text            =   "Total"
            Object.Width           =   3246
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   11
            Text            =   "Linea"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComctlLib.ListView lswTraslado 
         Height          =   3975
         Left            =   -74880
         TabIndex        =   19
         Top             =   1680
         Width           =   9975
         _ExtentX        =   17595
         _ExtentY        =   7011
         View            =   3
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         HotTracking     =   -1  'True
         HoverSelection  =   -1  'True
         _Version        =   393217
         SmallIcons      =   "ImageList1"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
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
            Text            =   "#Operaci�n"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   1
            Text            =   "L�nea"
            Object.Width           =   1658
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "C�dula"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Nombre"
            Object.Width           =   6068
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "Aprobado"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Text            =   "A Girar"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   6
            Text            =   "# Desem"
            Object.Width           =   2011
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   7
            Text            =   "Desembolsos"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   8
            Text            =   "Total"
            Object.Width           =   3246
         EndProperty
      End
      Begin XtremeSuiteControls.DateTimePicker dtpInicio 
         Height          =   330
         Left            =   3120
         TabIndex        =   67
         Top             =   1320
         Width           =   1335
         _Version        =   1441792
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
      Begin XtremeSuiteControls.DateTimePicker dtpCorte 
         Height          =   330
         Left            =   4440
         TabIndex        =   68
         Top             =   1320
         Width           =   1335
         _Version        =   1441792
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
      Begin XtremeSuiteControls.ComboBox cboTipo 
         Height          =   330
         Left            =   7320
         TabIndex        =   70
         Top             =   480
         Width           =   2655
         _Version        =   1441792
         _ExtentX        =   4683
         _ExtentY        =   582
         _StockProps     =   77
         ForeColor       =   0
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
      Begin VB.Label Label16 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Remesas - visualizar �ltimas"
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
         Index           =   4
         Left            =   -69360
         TabIndex        =   35
         Top             =   3600
         Width           =   3735
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   12
         X1              =   4800
         X2              =   6960
         Y1              =   720
         Y2              =   720
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Tipo de Remesa?"
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
         Left            =   4800
         TabIndex        =   53
         Top             =   480
         Width           =   2175
      End
      Begin VB.Label lblRemesa 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
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
         Height          =   315
         Left            =   -74760
         TabIndex        =   49
         Top             =   3600
         Width           =   5415
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
         TabIndex        =   48
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
         Caption         =   "Seleccione los promotores o comit�s para Pago por Tesorer�a"
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
         TabIndex        =   47
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
         Caption         =   "Seleccione los promotores o comit�s a Generar Comisi�n"
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
         TabIndex        =   46
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
      Begin VB.Line Line2 
         BorderColor     =   &H00FFFFFF&
         Index           =   1
         X1              =   120
         X2              =   9960
         Y1              =   840
         Y2              =   840
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   4
         X1              =   120
         X2              =   3000
         Y1              =   2640
         Y2              =   2640
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   3
         X1              =   120
         X2              =   3000
         Y1              =   2280
         Y2              =   2280
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   2
         X1              =   120
         X2              =   3000
         Y1              =   1920
         Y2              =   1920
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   1
         X1              =   120
         X2              =   3000
         Y1              =   360
         Y2              =   360
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
         TabIndex        =   45
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
         TabIndex        =   44
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
         TabIndex        =   43
         Top             =   840
         Width           =   2895
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Reporte de Comisiones de Afiliaci�n"
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
         TabIndex        =   42
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
         TabIndex        =   41
         Top             =   2040
         Width           =   3615
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Promotor / Comit�"
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
         TabIndex        =   40
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
         TabIndex        =   39
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
         TabIndex        =   38
         Top             =   4560
         Width           =   3615
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   16
         X1              =   120
         X2              =   3000
         Y1              =   720
         Y2              =   720
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
         TabIndex        =   37
         Top             =   5760
         Width           =   855
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   18
         X1              =   -74880
         X2              =   -72000
         Y1              =   600
         Y2              =   600
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Seleccione los Casos Pendientes"
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
         TabIndex        =   36
         Top             =   1560
         Width           =   9975
      End
      Begin VB.Label Label16 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
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
         TabIndex        =   34
         Top             =   1440
         Width           =   9735
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   11
         X1              =   -74880
         X2              =   -72000
         Y1              =   720
         Y2              =   720
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
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
         TabIndex        =   33
         Top             =   1440
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
         TabIndex        =   32
         Top             =   5760
         Width           =   855
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   10
         X1              =   120
         X2              =   3000
         Y1              =   3120
         Y2              =   3120
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Corte"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         Index           =   6
         Left            =   4440
         TabIndex        =   31
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Inicio"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         Index           =   5
         Left            =   3120
         TabIndex        =   30
         Top             =   960
         Width           =   1335
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   5
         X1              =   -74880
         X2              =   -72000
         Y1              =   960
         Y2              =   960
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
         TabIndex        =   29
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
         TabIndex        =   28
         Top             =   2880
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
         TabIndex        =   26
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
         TabIndex        =   25
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
         TabIndex        =   24
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
         TabIndex        =   23
         Top             =   3600
         Width           =   2895
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
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
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   -74880
         TabIndex        =   22
         Top             =   720
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
         TabIndex        =   21
         Top             =   360
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
         TabIndex        =   20
         Top             =   480
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
         TabIndex        =   27
         Top             =   960
         Width           =   2895
      End
   End
   Begin MSComctlLib.ProgressBar PrgBar 
      Align           =   2  'Align Bottom
      Height          =   135
      Left            =   0
      TabIndex        =   50
      Top             =   7530
      Visible         =   0   'False
      Width           =   10425
      _ExtentX        =   18389
      _ExtentY        =   238
      _Version        =   393216
      Appearance      =   0
   End
   Begin XtremeSuiteControls.ComboBox cboAseguradora 
      Height          =   330
      Left            =   4080
      TabIndex        =   69
      Top             =   600
      Width           =   6255
      _Version        =   1441792
      _ExtentX        =   11033
      _ExtentY        =   582
      _StockProps     =   77
      ForeColor       =   0
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
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Aseguradora"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2880
      TabIndex        =   66
      Top             =   600
      Width           =   1335
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Remesas para Desembolso de Comisiones / Aseguradora"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   2880
      TabIndex        =   51
      Top             =   120
      Width           =   7455
   End
   Begin VB.Image imgBanner 
      Height          =   1095
      Left            =   0
      Top             =   0
      Width           =   11655
   End
End
Attribute VB_Name = "frmSeguros_RemesaPago"
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
Dim vTipo As String

On Error GoTo vError

lswCarga.ListItems.Clear
If cboCarga.ListCount <= 0 Then Exit Sub

Me.MousePointer = vbHourglass

strSQL = "select fecha_inicio,fecha_corte,Tipo from SEGUROS_REMESAS where cod_remesa = " & cboCarga.ItemData(cboCarga.ListIndex)
rsTmp.Open strSQL, glogon.Conection, adOpenStatic
  vFechaInicio = rsTmp!FECHA_INICIO
  vFechaCorte = rsTmp!FECHA_CORTE
  vTipo = Trim(rsTmp!Tipo)
rsTmp.Close


vPaso = True

lswCarga.ColumnHeaders.Clear
lswCarga.ColumnHeaders.Add , , "No. Seguro", 2780
lswCarga.ColumnHeaders.Add , , "Aseguradora", 1100, vbCenter
lswCarga.ColumnHeaders.Add , , "No. Cta", 990, vbCenter
lswCarga.ColumnHeaders.Add , , "L�nea", 990, vbCenter
lswCarga.ColumnHeaders.Add , , "Cuota", 1290, vbRightJustify
lswCarga.ColumnHeaders.Add , , "Comision", 1290, vbRightJustify

If vTipo = "V" Then
    lswCarga.ColumnHeaders.Add , , "C�dula", 1290
    lswCarga.ColumnHeaders.Add , , "T.Seguro", 1290, vbCenter
    lswCarga.ColumnHeaders.Add , , "Vendedor", 3290
           
   cboVendedor.Enabled = True
           
    strSQL = "select Ven.cod_Vendedor as Idx, Ven.Nombre as Itmx" _
           & " from SEGUROS_REGISTRO Pol inner join SEGUROS_Pagos Pag on Pol.num_poliza = Pag.num_poliza" _
           & " inner join SEGUROS_Vendedores Ven on Pol.cod_Vendedor = Ven.cod_Vendedor" _
           & " Where Pag.COMISION_VENDEDOR_ESTADO = 'A' and Pag.Cod_Remesa_Comision is null" _
           & " and Pag.Fecha_Vence between '" & Format(vFechaInicio, "yyyy/mm/dd") & " 00:00:00'" _
           & " and '" & Format(vFechaCorte, "yyyy/mm/dd") & " 23:59:59'" _
           & " group by Ven.cod_Vendedor, Ven.Nombre" _
           & " order by Ven.Nombre"
    
    Call sbLlenaCbo(cboVendedor, strSQL, True, True)
Else
   cboVendedor.Clear
   cboVendedor.Enabled = False

    lswCarga.ColumnHeaders.Add , , "Mtn.Pago", 1290, vbRightJustify
    lswCarga.ColumnHeaders.Add , , "C�dula", 1290
    lswCarga.ColumnHeaders.Add , , "T.Seguro", 1290, vbCenter
    lswCarga.ColumnHeaders.Add , , "Vendedor", 3290

End If
vPaso = False

Me.MousePointer = vbDefault

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical


End Sub


Private Sub sbConsulta(pRemesa As Long)
Dim strSQL As String, rs As New ADODB.Recordset

Call sbLimpia
  
strSQL = "select * from SEGUROS_REMESAS where cod_remesa = " & pRemesa
Call OpenRecordSet(rs, strSQL)
If Not rs.BOF And Not rs.EOF Then
  txtRemesa = rs!cod_remesa
  txtUsuario = rs!Usuario
  txtFecha = rs!fecha
  
  Select Case rs!estado
    Case "A"
      txtEstado = "Remesa Abierta"
    Case "C"
      txtEstado = "Remesa Cerrada"
    Case "T"
      txtEstado = "Remesa Trasladada"
  End Select
  
  dtpInicio.Value = rs!FECHA_INICIO
  dtpCorte.Value = rs!FECHA_CORTE
  
  txtNotas.Text = rs!notas
  
End If
rs.Close

End Sub


Private Sub cboTipo_Click()
Dim strSQL As String, rs As New ADODB.Recordset
     
If vPaso Then Exit Sub
     
On Error GoTo vError
     
Me.MousePointer = vbHourglass
     
     
txtEstado = ""
txtFecha = ""
txtUsuario = ""
txtRemesa = ""
txtNotas.Text = ""

lswRemesas.ListItems.Clear


strSQL = "select isnull(max(Fecha_Corte),dbo.MyGetdate()) as 'Inicio', dbo.MyGetdate() as 'Corte' from SEGUROS_REMESAS where Tipo = '" & Mid(cboTipo.Text, 1, 1) & "'"
Call OpenRecordSet(rs, strSQL)
    dtpInicio.Value = rs!Inicio
    dtpCorte.Value = rs!Corte
rs.Close

strSQL = "select TOP 150 * from SEGUROS_REMESAS where Tipo = '" & Mid(cboTipo.Text, 1, 1) & "' order by fecha desc"
Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
  With lswRemesas.ListItems
       Set itmX = .Add(, , rs!cod_remesa)
           itmX.SubItems(1) = rs!Usuario
           itmX.SubItems(2) = rs!fecha
           
           Select Case rs!estado
             Case "A"
                itmX.SubItems(3) = "Remesa Abierta"
             Case "C"
                itmX.SubItems(3) = "Remesa Cerrada"
             Case "T"
                itmX.SubItems(3) = "Remesa Trasladada"
           End Select
           
           itmX.SubItems(4) = Format(rs!FECHA_INICIO, "dd/mm/yyyy")
           itmX.SubItems(5) = Format(rs!FECHA_CORTE, "dd/mm/yyyy")
           itmX.SubItems(6) = rs!notas
           
  End With
  rs.MoveNext
Loop
rs.Close

Me.MousePointer = vbDefault

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical


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
 .WindowTitle = "M�dulo de Comercializaci�n de Seguros"

 .Connect = glogon.ConectRPT

If chkRemesaInd.Value = vbChecked Then
   xRemesa = InputBox("Indique la Remesa que desea consultar", "Remesas de Traslado a Bancos")
  If Len(Trim(xRemesa)) = 0 Then xRemesa = "0"
  lblRemesa.Tag = xRemesa
End If


If optTipo.Item(0).Value Then
  'Vendedores (Comisiones)
   .Formulas(3) = "fxTitulo='REMESA TRASLADO A TESORERIA : PAGO COMISIONES'"
    Select Case True
     Case opt.Item(0).Value 'Pendiente Detalle Remesa
        .ReportFileName = SIFGlobal.fxPathReportes("SEGUROS_RemesaVenDetalle.rpt")
        vSubTitulo = "REMESA : " & lblRemesa.Tag & " LISTADO : DETALLADO"
     Case opt.Item(1).Value 'Traslado Detalle Agrupado Remesa
        .ReportFileName = SIFGlobal.fxPathReportes("SEGUROS_RemesaVenDetalleAgrp.rpt")
        vSubTitulo = "REMESA : " & lblRemesa.Tag & " LISTADO : DETALLADO AGRUPADO"
    End Select
 Else
   'Aseguradora (Pagos)
   .Formulas(3) = "fxTitulo='REMESA TRASLADO A TESORERIA : PAGO ASEGURADORA'"
    
    Select Case True
     Case opt.Item(0).Value 'Pendiente Detalle Remesa
        .ReportFileName = SIFGlobal.fxPathReportes("SEGUROS_RemesaAseDetalle.rpt")
        vSubTitulo = "REMESA : " & lblRemesa.Tag & " LISTADO : DETALLADO"
     Case opt.Item(1).Value 'Traslado Detalle Agrupado Remesa
        .ReportFileName = SIFGlobal.fxPathReportes("SEGUROS_RemesaAseDetalleAgrp.rpt")
        vSubTitulo = "REMESA : " & lblRemesa.Tag & " LISTADO : DETALLADO AGRUPADO"
    End Select
   
   
 End If
 .Formulas(0) = "fxFecha='FECHA: " & Format(fxFechaServidor, "dd/mm/yyyy") & "'"
 .Formulas(1) = "fxEmpresa='" & GLOBALES.gstrNombreEmpresa & "'"
 .Formulas(2) = "fxUsuario='USER: " & glogon.Usuario & "'"
 .Formulas(4) = "fxSubTitulo='" & vSubTitulo & "'"
 .SelectionFormula = "{SEGUROS_REMESAS.cod_remesa} = " & lblRemesa.Tag
 .PrintReport

End With

Me.MousePointer = vbDefault
Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub Form_Activate()
vModulo = 17
End Sub

Private Sub lswRep_Click()
If lswRep.ListItems.Count <= 0 Then Exit Sub

lblRemesa.Caption = lswRep.SelectedItem.Text & " � " & lswRep.SelectedItem.SubItems(1) _
            & " � " & lswRep.SelectedItem.SubItems(2)
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
   MsgBox "Ocurri� un error al ordenar los datos de la columna seleccionada.", vbCritical

End Sub


Private Sub cboTraslado_Click()
    lswTraslado.ListItems.Clear
    txtPagoTotal.Text = 0
End Sub

Private Sub chkCarga_Click()
Dim i As Integer, curTotal As Currency
Dim Col As Integer, iCasos As Integer

If cboVendedor.Enabled Then
  Col = 4
Else
  Col = 5
End If


iCasos = 0
For i = 1 To lswCarga.ListItems.Count
  lswCarga.ListItems.Item(i).Checked = chkCarga.Value
  
   If chkCarga.Value = vbChecked Then
       curTotal = curTotal + CCur(lswCarga.ListItems.Item(i).SubItems(Col))
       iCasos = iCasos + 1
   End If
  
Next i

txtCargaTotal.Text = Format(curTotal, "Standard")
txtCargaTotal.ToolTipText = "Casos ..: " & iCasos

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

Private Sub lswCarga_ItemCheck(ByVal Item As MSComctlLib.ListItem)
Dim curTotal As Currency

If Trim(txtCargaTotal.Text) = "" Then txtCargaTotal.Text = 0

curTotal = CCur(txtCargaTotal.Text)

If Item.Checked Then
   curTotal = curTotal + CCur(Item.SubItems(9))
Else
   curTotal = curTotal - CCur(Item.SubItems(9))
End If

txtCargaTotal.Text = Format(curTotal, "Standard")

End Sub

Private Sub lswRemesas_Click()
    If lswRemesas.ListItems.Count <= 0 Then Exit Sub
    Call sbConsulta(lswRemesas.SelectedItem)
End Sub


Private Sub optRemesaEstado_Click(Index As Integer)
Call sbLimpia
End Sub

Private Sub optTipo_Click(Index As Integer)
Call sbLimpia
End Sub

Private Sub ssTab_Click(PreviousTab As Integer)
 Call sbLimpia
End Sub

Private Sub sbReporteRemesas(pRemesa As Long)
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
 .WindowTitle = "Reportes del M�dulo de Comercializaci�n de Seguros"
 
 .Connect = glogon.ConectRPT
  
 
 .Formulas(0) = "fxFecha='FECHA: " & Format(fxFechaServidor, "dd/mm/yyyy") & "'"
 .Formulas(1) = "fxEmpresa='" & GLOBALES.gstrNombreEmpresa & "'"
 .Formulas(2) = "fxUsuario='USER: " & glogon.Usuario & "'"
 .Formulas(3) = "fxSubTitulo='" & vSubTitulo & "'"
 .Formulas(4) = "fxFiltro='" & vFiltro & "'"
 
 .ReportFileName = SIFGlobal.fxPathReportes("AfiComisionRemesas.rpt")
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
     
            strSQL = "select isnull(max(cod_remesa),0) + 1 as Ultimo from SEGUROS_REMESAS"
            Call OpenRecordSet(rs, strSQL)
                strSQL = "Insert SEGUROS_REMESAS(cod_remesa,usuario,fecha,estado,fecha_inicio,fecha_corte,notas,Tipo) values(" & rs!ultimo _
                       & ",'" & glogon.Usuario & "',dbo.MyGetdate(),'A','" & Format(dtpInicio.Value, "yyyy/mm/dd") _
                       & "','" & Format(dtpCorte.Value, "yyyy/mm/dd") & "','" & txtNotas.Text & "','" & Mid(cboTipo.Text, 1, 1) & "')"
                Call ConectionExecute(strSQL)
                
                txtRemesa = rs!ultimo
            rs.Close
            Call Bitacora("Registra", "Remesa de Seguros (Abre): " & txtRemesa)
    
    Else
        If txtEstado.Text <> "Remesa Cerrada" Then
                    
            strSQL = "update SEGUROS_REMESAS set usuario = '" & glogon.Usuario & "',fecha_inicio = '" & Format(dtpInicio.Value, "yyyy/mm/dd") _
                   & "',fecha_corte = '" & Format(dtpCorte.Value, "yyyy/mm/dd") & "',notas = '" & txtNotas.Text _
                   & "' where cod_remesa = " & txtRemesa
             Call ConectionExecute(strSQL)
             
            Call Bitacora("Modifica", "Remesa de Seguros (Abre): " & txtRemesa)
        Else
            MsgBox "No se puede Modifica la remesa, porque esta ya fue cerrada...", vbExclamation
        End If
    End If
    
    Call sbLimpia
    
    
  Case "BORRAR"
     i = MsgBox("Esta Seguro que desea borrar este registro", vbYesNo)
     If i = vbYes Then
        
        If txtEstado.Text = "Remesa Abierta" Then
            strSQL = "update SEGUROS_REMESAS set cod_Remesa = Null where cod_Remesa = " & txtRemesa.Text
            Call ConectionExecute(strSQL)
            
            strSQL = "update SEGUROS_REMESAS set Cod_Remesa_Comision = Null where Cod_Remesa_Comision = " & txtRemesa.Text
            Call ConectionExecute(strSQL)
            
            Call Bitacora("Elimina", "Remesa de Seguros (Abre): " & txtRemesa)
         End If
       
        Call sbLimpia
     End If
  
  Case "REPORTES"
     If IsNumeric(txtRemesa) Then
         Call sbReporteRemesas(txtRemesa)
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
Dim strSQL As String, rs As New ADODB.Recordset

Me.MousePointer = vbHourglass

Select Case ssTab.Tab
  Case 0 'Remesas
    Call cboTipo_Click
    
     
  Case 1 'Carga
    'Solo busca las remesas que se encuentran Abiertas
    
    vPaso = True
    
    cboCarga.Clear

    lswCarga.ListItems.Clear
    chkCarga.Value = vbUnchecked
        
    strSQL = "select * from SEGUROS_REMESAS where estado = 'A' order by fecha desc"
    Call OpenRecordSet(rs, strSQL)
    Do While Not rs.EOF
      cboCarga.AddItem (Format(rs!cod_remesa, "0000") & ".." & Trim(rs!Tipo) & ".." & Trim(rs!Usuario) & "..." _
            & rs!fecha & " I:" & Format(rs!FECHA_INICIO, "dd/mm/yyyy") & " C:" & Format(rs!FECHA_CORTE, "dd/mm/yyyy"))
      cboCarga.ItemData(cboCarga.NewIndex) = rs!cod_remesa
      rs.MoveNext
    Loop
    If rs.RecordCount > 0 Then
       rs.MoveFirst
       cboCarga.Text = (Format(rs!cod_remesa, "0000") & ".." & Trim(rs!Tipo) & ".." & Trim(rs!Usuario) & "..." _
            & rs!fecha & " I:" & Format(rs!FECHA_INICIO, "dd/mm/yyyy") & " C:" & Format(rs!FECHA_CORTE, "dd/mm/yyyy"))
    End If
    
    rs.Close

    vPaso = False
    Call cboCarga_Click
   
    
  Case 2 'Traslado
    vPaso = True
    
    cboTraslado.Clear

    lswTraslado.ListItems.Clear
    txtPagoTotal.Text = 0
        
    lswTraslado.ColumnHeaders.Clear
    lswTraslado.ColumnHeaders.Add , , "No. P�liza", 2780
    lswTraslado.ColumnHeaders.Add , , "Aseguradora", 1100, vbCenter
    lswTraslado.ColumnHeaders.Add , , "No. Cta", 990, vbCenter
    lswTraslado.ColumnHeaders.Add , , "Cuota", 1290, vbRightJustify
    lswTraslado.ColumnHeaders.Add , , "Comision", 1290, vbRightJustify

    
    strSQL = "select * from SEGUROS_REMESAS where Estado = 'C'  order by fecha desc"
    Call OpenRecordSet(rs, strSQL)
    Do While Not rs.EOF
      cboTraslado.AddItem (Format(rs!cod_remesa, "0000") & ".." & Trim(rs!Tipo) & ".." & Trim(rs!Usuario) & "..." _
            & rs!fecha & " I:" & Format(rs!FECHA_INICIO, "dd/mm/yyyy") & " C:" & Format(rs!FECHA_CORTE, "dd/mm/yyyy"))
      cboTraslado.ItemData(cboTraslado.NewIndex) = rs!cod_remesa
      rs.MoveNext
    Loop
    If rs.RecordCount > 0 Then
       rs.MoveFirst
       cboTraslado.Text = (Format(rs!cod_remesa, "0000") & ".." & Trim(rs!Tipo) & ".." & Trim(rs!Usuario) & "..." _
            & rs!fecha & " I:" & Format(rs!FECHA_INICIO, "dd/mm/yyyy") & " C:" & Format(rs!FECHA_CORTE, "dd/mm/yyyy"))
    End If
    
    rs.Close

    vPaso = False
    Call cboTraslado_Click

  
  Case 3 'Reportes
     strSQL = "select TOP " & txtRepRemesas.Text & " * from SEGUROS_REMESAS"
     
     
     Select Case True
        Case optRemesaEstado.Item(1).Value 'Abiertas
            strSQL = strSQL & " where Estado = 'A'"
        Case optRemesaEstado.Item(2).Value 'Cerradas
            strSQL = strSQL & " where Estado = 'C'"
        Case optRemesaEstado.Item(3).Value 'Trasladadas
            strSQL = strSQL & " where Estado = 'T'"
        Case Else
            strSQL = strSQL & " where Estado in('A','C','T')"
    End Select
     
     Select Case True
        Case optTipo.Item(0).Value 'Vendedores
            strSQL = strSQL & " and Tipo = 'V'"
        Case optTipo.Item(1).Value 'Aseguradora
            strSQL = strSQL & " and Tipo = 'A'"
     End Select
     
    strSQL = strSQL & " order by fecha desc "
     
     lswRep.ListItems.Clear

     Call OpenRecordSet(rs, strSQL)
     Do While Not rs.EOF
       With lswRep.ListItems
            Set itmX = .Add(, , rs!cod_remesa)
                itmX.SubItems(1) = rs!Usuario
                itmX.SubItems(2) = rs!fecha
                itmX.SubItems(3) = rs!FECHA_INICIO
                itmX.SubItems(4) = rs!FECHA_CORTE
                itmX.SubItems(5) = rs!notas
                itmX.SubItems(6) = rs!Tipo
       
       End With
       rs.MoveNext
     Loop
     rs.Close
 End Select


Me.MousePointer = vbDefault

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
Dim vTipo As String
Dim curTotal As Currency

Me.MousePointer = vbHourglass

On Error GoTo vError

lswCarga.ListItems.Clear
curTotal = 0


strSQL = "select fecha_inicio,fecha_corte,Tipo from SEGUROS_REMESAS where cod_remesa = " & cboCarga.ItemData(cboCarga.ListIndex)
Call OpenRecordSet(rs, strSQL)
  vFechaInicio = rs!FECHA_INICIO
  vFechaCorte = rs!FECHA_CORTE
  vTipo = Trim(rs!Tipo)
rs.Close

strSQL = "select Pag.Cod_Aseguradora, Pag.Num_Poliza,Pag.Num_Cuota,Pag.Monto,Pol.Cedula,Pol.COD_PRODUCTO,Ven.Nombre as 'Vendedor'" _
       & ", Pag.Monto_Pago,Pag.Comision_Interna,Pag.Comision_Vendedor,Pag.Linea" _
       & " from SEGUROS_REGISTRO Pol" _
       & " inner join SEGUROS_Pagos Pag on Pol.num_poliza = Pag.num_poliza and Pol.cod_Aseguradora = Pag.Cod_Aseguradora" _
       & " inner join SEGUROS_Vendedores Ven on Pol.cod_Vendedor = Ven.cod_Vendedor"
    
If vTipo = "V" Then

    strSQL = strSQL & " Where Pag.COMISION_VENDEDOR_ESTADO = 'A' and Pag.Cod_Remesa_Comision is null" _
           & " and Pag.Fecha_Vence between '" & Format(vFechaInicio, "yyyy/mm/dd") & " 00:00:00'" _
           & " and '" & Format(vFechaCorte, "yyyy/mm/dd") & " 23:59:59'" _

    If cboVendedor.Text <> "TODOS" Then
       strSQL = strSQL & " and Ven.Cod_Vendedor = " & cboVendedor.ItemData(cboVendedor.ListIndex)
    End If
    
    If cboAseguradora.Text <> "TODOS" Then
       strSQL = strSQL & " and Pol.cod_Aseguradora = '" & cboAseguradora.ItemData(cboAseguradora.ListIndex) & "'"
    End If
    
    strSQL = strSQL & " order by Pol.cod_Vendedor,Pag.Cod_Aseguradora, Pag.Num_Poliza,Pag.Num_Cuota"
    
    Call OpenRecordSet(rs, strSQL)
    
    PrgBar.Max = rs.RecordCount + 1
    PrgBar.Value = 1
    PrgBar.Visible = True
    
    With lswCarga
     .ListItems.Clear
     Do While Not rs.EOF
       Set itmX = .ListItems.Add(, , rs!Num_Poliza)
           itmX.SubItems(1) = rs!cod_Aseguradora
           itmX.SubItems(2) = rs!Num_Cuota
           itmX.SubItems(3) = rs!Linea
           itmX.SubItems(4) = Format(rs!Monto, "Standard")
           itmX.SubItems(5) = Format(rs!Comision_Vendedor, "Standard")
           itmX.SubItems(6) = rs!Cedula
           itmX.SubItems(7) = rs!COD_PRODUCTO
           itmX.SubItems(8) = rs!Vendedor
           
           itmX.Checked = chkCarga.Value
             
           If itmX.Checked Then
                curTotal = curTotal + rs!Comision_Vendedor
           End If
            
            rs.MoveNext
            
            PrgBar.Value = PrgBar.Value + 1
     Loop
    End With
    
    rs.Close
    
Else
   
    strSQL = strSQL & " Where Pag.ESTADO_PAGO = 'A' and Pag.Cod_Remesa is null" _
           & " and Pag.Fecha_Vence between '" & Format(vFechaInicio, "yyyy/mm/dd") & " 00:00:00'" _
           & " and '" & Format(vFechaCorte, "yyyy/mm/dd") & " 23:59:59'" _

    If cboAseguradora.Text <> "TODOS" Then
       strSQL = strSQL & " and Pol.cod_Aseguradora = '" & cboAseguradora.ItemData(cboAseguradora.ListIndex) & "'"
    End If
    
    strSQL = strSQL & " order by Pag.Cod_Aseguradora, Pag.Num_Poliza,Pag.Num_Cuota"
    
    Call OpenRecordSet(rs, strSQL)
    
    PrgBar.Max = rs.RecordCount + 1
    PrgBar.Value = 1
    PrgBar.Visible = True
    
    With lswCarga
     .ListItems.Clear
     Do While Not rs.EOF
       Set itmX = .ListItems.Add(, , rs!Num_Poliza)
           itmX.SubItems(1) = rs!cod_Aseguradora
           itmX.SubItems(2) = rs!Num_Cuota
           itmX.SubItems(3) = rs!Linea
           itmX.SubItems(4) = Format(rs!Monto, "Standard")
           itmX.SubItems(5) = Format(rs!Comision_Interna, "Standard")
           itmX.SubItems(6) = Format(rs!Monto_Pago, "Standard")
           itmX.SubItems(7) = rs!Cedula
           itmX.SubItems(8) = rs!COD_PRODUCTO
           itmX.SubItems(9) = rs!Vendedor
           itmX.Checked = chkCarga.Value
             
           If itmX.Checked Then
                curTotal = curTotal + rs!Monto_Pago
           End If
            
            rs.MoveNext
            
            PrgBar.Value = PrgBar.Value + 1
     Loop
    End With
    
    rs.Close

End If


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
Dim vTipo As String, curTotal As Currency

On Error GoTo vError

lswTraslado.ListItems.Clear

If cboTraslado.ListCount <= 0 Then Exit Sub

Me.MousePointer = vbHourglass

curTotal = 0

strSQL = "select fecha_inicio,fecha_corte,Tipo from SEGUROS_REMESAS where cod_remesa = " & cboTraslado.ItemData(cboTraslado.ListIndex)
Call OpenRecordSet(rs, strSQL)
  vFechaInicio = rs!FECHA_INICIO
  vFechaCorte = rs!FECHA_CORTE
  vTipo = Trim(rs!Tipo)
rs.Close

strSQL = "select Pag.cod_Aseguradora, Pag.Num_Poliza,Pag.Num_Cuota,Pag.Monto,Pol.Cedula,Pol.COD_PRODUCTO" _
       & ",Ven.Nombre as 'Vendedor', Pol.cod_Comercializadora, Com.Descripcion as 'Comercializadora'" _
       & ", Pag.Monto_Pago,Pag.isnull(Comision_Interna,0) as 'ComInterna',isnull(Pag.Comision_Comercializa,0) as 'ComComercializa'" _
       & ", case when Comision_Vendedor_Info = 1 then 0 else isnull(Pag.Comision_Vendedor,0) end as 'ComVendedor'" _
       & ", Pag.Linea"


With lswTraslado
    
.ColumnHeaders.Clear
.ColumnHeaders.Add , , "Aseguradora", 1100, vbCenter
.ColumnHeaders.Add , , "No. P�liza", 2780
.ColumnHeaders.Add , , "No. Cta", 990, vbCenter
.ColumnHeaders.Add , , "Cuota", 1290, vbRightJustify
.ColumnHeaders.Add , , "Mtn.Pago", 1290, vbRightJustify
.ColumnHeaders.Add , , "Beneficiario", 1290, vbRightJustify
.ColumnHeaders.Add , , "Com.Interna", 1290, vbRightJustify
.ColumnHeaders.Add , , "C�dula", 1290
.ColumnHeaders.Add , , "T.Seguro", 1290, vbCenter
.ColumnHeaders.Add , , "Vendedor", 3290
.ColumnHeaders.Add , , "Comercializa", 3290

 Select Case vTipo
    
    Case "A" 'Aseguradora
    
    strSQL = strSQL & ",Pag.Monto_Pago as 'Monto_Ref',Aseg.Descripcion as 'Beneficiario'" _
       & " from SEGUROS_REGISTRO Pol" _
       & " inner join SEGUROS_Pagos Pag on Pol.num_poliza = Pag.num_poliza and Pol.Cod_Aseguradora = Pag.Cod_Aseguradora" _
       & " inner join SEGUROS_Vendedores Ven on Pol.cod_Vendedor = Ven.cod_Vendedor" _
       & " inner join SEGUROS_COMERCIALIZADORAS Com on Pol.cod_Comercializadora = Com.cod_Comercializadora " _
       & " inner join SEGUROS_ASEGURADORAS Aseg on Pol.cod_aseguradora = Aseg.Cod_Aseguradora" _
       & " Where Pag.Monto_Pago > 0 and Pag.ESTADO_PAGO = 'G' and Pag.Cod_Remesa = " & cboTraslado.ItemData(cboTraslado.ListIndex) _
               & " and Pag.Fecha_Vence between '" & Format(vFechaInicio, "yyyy/mm/dd") & " 00:00:00'" _
               & " and '" & Format(vFechaCorte, "yyyy/mm/dd") & " 23:59:59'" _
               & " order by Pag.Cod_Aseguradora,Pag.Num_Poliza,Pag.Num_Cuota"
        
        Call OpenRecordSet(rs, strSQL)
        
        PrgBar.Max = rs.RecordCount + 1
        PrgBar.Value = 1
        PrgBar.Visible = True
        
         .ListItems.Clear
         Do While Not rs.EOF
           Set itmX = .ListItems.Add(, , rs!cod_Aseguradora)
               itmX.SubItems(1) = rs!Num_Poliza
               itmX.SubItems(2) = rs!Num_Cuota
               itmX.SubItems(3) = Format(rs!Monto, "Standard")
               itmX.SubItems(4) = Format(rs!Monto_Pago, "Standard")
               itmX.SubItems(5) = rs!Beneficiario
               itmX.SubItems(6) = Format(rs!Comision_Interna, "Standard")
               itmX.SubItems(7) = rs!Cedula
               itmX.SubItems(8) = rs!COD_PRODUCTO
               itmX.SubItems(9) = rs!Vendedor
               itmX.SubItems(10) = rs!Comercializadora
               itmX.Tag = rs!Linea
               curTotal = curTotal + rs!Monto_Ref
                
                rs.MoveNext
                
                PrgBar.Value = PrgBar.Value + 1
         Loop
        
        rs.Close

    
    
    Case "V" 'Vendedor

            .ColumnHeaders.Add , , "C�dula", 1290
            .ColumnHeaders.Add , , "T.Seguro", 1290, vbCenter
            .ColumnHeaders.Add , , "Vendedor", 3290
        
            strSQL = strSQL & " Where Pag.COMISION_VENDEDOR_ESTADO = 'G' and Pag.Cod_Remesa_Comision = " & cboTraslado.ItemData(cboTraslado.ListIndex) _
                   & " and Pag.Fecha_Vence between '" & Format(vFechaInicio, "yyyy/mm/dd") & " 00:00:00'" _
                   & " and '" & Format(vFechaCorte, "yyyy/mm/dd") & " 23:59:59'" _
                   & " order by Pol.cod_Vendedor,Pag.Cod_Aseguradora, Pag.Num_Poliza,Pag.Num_Cuota"
            
            Call OpenRecordSet(rs, strSQL)
            
            PrgBar.Max = rs.RecordCount + 1
            PrgBar.Value = 1
            PrgBar.Visible = True
            
        
             .ListItems.Clear
             Do While Not rs.EOF
               Set itmX = .ListItems.Add(, , rs!Num_Poliza)
                   itmX.SubItems(1) = rs!cod_Aseguradora
                   itmX.SubItems(2) = rs!Num_Cuota
                   itmX.SubItems(3) = rs!Linea
                   itmX.SubItems(4) = Format(rs!Monto, "Standard")
                   itmX.SubItems(5) = Format(rs!Comision_Vendedor, "Standard")
                   itmX.SubItems(6) = rs!Cedula
                   itmX.SubItems(7) = rs!COD_PRODUCTO
                   itmX.SubItems(8) = rs!Vendedor
                     
                   curTotal = curTotal + rs!Comision_Vendedor
                   
                    rs.MoveNext
                    
                    PrgBar.Value = PrgBar.Value + 1
             Loop
            
            rs.Close



    Case "C" 'Comercializadora
    
        .ColumnHeaders.Add , , "Mtn.Pago", 1290, vbRightJustify
        .ColumnHeaders.Add , , "C�dula", 1290
        .ColumnHeaders.Add , , "T.Seguro", 1290, vbCenter
        .ColumnHeaders.Add , , "Vendedor", 3290
    
    
        strSQL = strSQL & " Where Pag.ESTADO_PAGO = 'G' and Pag.Cod_Remesa = " & cboTraslado.ItemData(cboTraslado.ListIndex) _
               & " and Pag.Fecha_Vence between '" & Format(vFechaInicio, "yyyy/mm/dd") & " 00:00:00'" _
               & " and '" & Format(vFechaCorte, "yyyy/mm/dd") & " 23:59:59'" _
               & " order by Pag.Cod_Aseguradora,Pag.Num_Poliza,Pag.Num_Cuota"
        
        Call OpenRecordSet(rs, strSQL)
        
        PrgBar.Max = rs.RecordCount + 1
        PrgBar.Value = 1
        PrgBar.Visible = True
        
         .ListItems.Clear
         Do While Not rs.EOF
           Set itmX = .ListItems.Add(, , rs!Num_Poliza)
               itmX.SubItems(1) = rs!cod_Aseguradora
               itmX.SubItems(2) = rs!Num_Cuota
               itmX.SubItems(3) = rs!Linea
               itmX.SubItems(4) = Format(rs!Monto, "Standard")
               itmX.SubItems(5) = Format(rs!Comision_Interna, "Standard")
               itmX.SubItems(6) = Format(rs!Monto_Pago, "Standard")
               itmX.SubItems(7) = rs!Cedula
               itmX.SubItems(8) = rs!COD_PRODUCTO
               itmX.SubItems(9) = rs!Vendedor
                 
               curTotal = curTotal + rs!Monto_Pago
                
                rs.MoveNext
                
                PrgBar.Value = PrgBar.Value + 1
         Loop
        
        rs.Close

End Select


End With

PrgBar.Visible = False

txtPagoTotal.Text = Format(curTotal, "Standard")
txtPagoTotal.ToolTipText = "Casos ..: " & PrgBar.Max - 1

Me.MousePointer = vbDefault

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

  txtPagoTotal.Text = Format(curTotal, "Standard")
  lswTraslado.ListItems.Clear
  
End Sub



Private Sub sbCerrar()
Dim strSQL As String, rs As New ADODB.Recordset
Dim i As Integer, vCasos As Integer
Dim vFecha As Date

On Error GoTo vError


'Valida el Estado de la Remesa
strSQL = "select count(*) as Existe from SEGUROS_REMESAS" _
       & " where cod_remesa = " & cboCarga.ItemData(cboCarga.ListIndex) _
       & " and estado = 'A'"
Call OpenRecordSet(rs, strSQL)
 If rs!Existe = 0 Then
    rs.Close
    MsgBox "La Remesa actual; ya se encuentra cerrada...", vbExclamation
    Exit Sub
 End If
rs.Close

'Actualiza el Estado de la Remesa como cerrada
'Actualiza datos de Pagos y Comisiones en el Maestro de Polizas
strSQL = "exec spSeguros_RemesaCierre " & cboCarga.ItemData(cboCarga.ListIndex)
Call ConectionExecute(strSQL)

 
Call Bitacora("Aplica", "SEGUROS Remesa (Cierra) : " & cboCarga.ItemData(cboCarga.ListIndex))


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
strSQL = "select count(*) as Existe from SEGUROS_REMESAS" _
       & " where cod_remesa = " & cboCarga.ItemData(cboCarga.ListIndex) _
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
     
     If cboVendedor.Enabled Then
        strSQL = "update SEGUROS_PAGOS Set cod_remesa_comision = " & cboCarga.ItemData(cboCarga.ListIndex) _
               & " where NUM_POLIZA = '" & .Item(i).Text & "' and COD_ASEGURADORA = '" & .Item(i).SubItems(1) _
               & "' AND NUM_CUOTA = " & .Item(i).SubItems(2) _
               & "   AND LINEA = " & .Item(i).SubItems(3)
     Else
        strSQL = "update SEGUROS_PAGOS set cod_remesa = " & cboCarga.ItemData(cboCarga.ListIndex) _
               & " where NUM_POLIZA = '" & .Item(i).Text & "' and COD_ASEGURADORA = '" & .Item(i).SubItems(1) _
               & "' AND NUM_CUOTA = " & .Item(i).SubItems(2) _
               & "   AND LINEA = " & .Item(i).SubItems(3)
     End If
     
     Call ConectionExecute(strSQL)
   
    PrgBar.Value = PrgBar.Value + 1
  End If
Next i
 
If vCasos > 0 Then
    Call Bitacora("Aplica", "Remesa de Seguros (Carga): " & cboCarga.ItemData(cboCarga.ListIndex))
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
                              , vDetalle2 As String, vCuenta As String, vFecha As Date, vUnidad As String) As Long                                 'Regresa el NSOLICITUD
Dim rsX As New ADODB.Recordset, strSQL As String
Dim lngSol As Long

strSQL = "Insert Tes_Transacciones(cod_concepto,cod_unidad,id_banco,tipo,codigo,beneficiario,monto,fecha_solicitud,estado,estadoi" _
       & ",modulo,submodulo,cta_ahorros,detalle1,detalle2,referencia,op,genera,actualiza,user_solicita,autoriza,user_autoriza,fecha_autorizacion)" _
       & " values('" & mConcepto & "','" & vUnidad & "'," & vBanco & ",'" & vTipoDocumento & "','" & vCodigo & "','" & vBeneficiario & "'," & vMonto _
       & ",'" & Format(vFecha, "yyyy/mm/dd") & "','P','P','CC','C','" & vCuenta _
       & "','" & vDetalle1 & "','" & vDetalle2 & "'," & vReferencia & "," & vOP & ",'S','S','" & glogon.Usuario & "'"
       
If UCase(vTipoDocumento) = "CK" Then
   strSQL = strSQL & ",'S','" & glogon.Usuario & "',dbo.MyGetdate())"
Else
   strSQL = strSQL & ",'N',null,null)"
End If
Call ConectionExecute(strSQL)

strSQL = "select max(nsolicitud) as Solicitud from Tes_Transacciones"
Call OpenRecordSet(rsX, strSQL)
 strSQL = "select * from Tes_Transacciones where nsolicitud = " & rsX!solicitud
rsX.Close

lngSol = 0

Call OpenRecordSet(rsX, strSQL)
If Trim(rsX!Codigo) = Trim(vCodigo) Then lngSol = rsX!NSolicitud
rsX.Close

If lngSol = 0 Then
  strSQL = "select max(nsolicitud) as Solicitud from Tes_Transacciones where codigo ='" & vCodigo _
         & "'"
  Call OpenRecordSet(rsX, strSQL)
  lngSol = rsX!solicitud
  rsX.Close
End If

fxMaestroTesoreria = lngSol

End Function



Private Sub sbCreaDetalle(vSolicitud As Long, vCtaConta As String, vMonto As Currency, vDH As String, vLinea As Integer, vUnidad As String)
Dim strSQL As String

strSQL = "Insert Tes_Trans_Asiento(nsolicitud,cuenta_contable,monto,debehaber,linea,cod_unidad) values(" _
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
Dim vTipo As String, Cuenta(4) As String

Dim vNombre As String, vCedula As String, vEmite As String, vBanco As Integer, vCtaBanco As String

Me.MousePointer = vbHourglass

On Error GoTo vError

vFecha = fxFechaServidor

strSQL = "select Tipo from SEGUROS_remesas where cod_remesa = " & cboTraslado.ItemData(cboTraslado.ListIndex)
Call OpenRecordSet(rs, strSQL)
vTipo = rs!Tipo

Cuenta(0) = fxSeguros_Parametro("10")  'Transitoria
Cuenta(1) = fxSeguros_Parametro("08")  'Ingreso por Comision
Cuenta(2) = fxSeguros_Parametro("09")  'Gasto por Comision


'Esta cuenta Existe del lado de la aseguradora no como par�metro
'Cuenta(3) = fxSeguros_Parametro("07")  'CxP Aseguradora


If rs!Tipo = "V" Then
 mConcepto = fxSeguros_Parametro("11")
Else
 mConcepto = fxSeguros_Parametro("12")
End If
rs.Close



strSQL = "select cod_unidad from sif_Oficinas where oficina_omision = 1"
Call OpenRecordSet(rs, strSQL)
mUnidad = Trim(rs!Cod_Unidad)
rs.Close


Select Case vTipo
    Case "V" 'Vendedor
       strSQL = "select Ven.CEDULA,Ven.NOMBRE, Ven.Cod_Banco,Ven.Tipo_Emision,Ven.Cuenta_Bancaria" _
              & ",SUM(Pag.COMISION_VENDEDOR) as 'Monto'" _
              & "  from SEGUROS_REGISTRO Pol" _
              & "  inner join SEGUROS_PAGOS Pag on Pol.NUM_POLIZA = Pag.NUM_POLIZA and Pol.Cod_Aseguradora = Pag.Cod_Aseguradora" _
              & "  inner join SEGUROS_VENDEDORES Ven on Pol.COD_VENDEDOR = Ven.COD_VENDEDOR" _
              & " Where Pag.COD_REMESA_Comision = " & cboTraslado.ItemData(cboTraslado.ListIndex) & " And Pag.Tesoreria_Solicitud_Comision Is Null" _
              & " group by Ven.CEDULA,Ven.NOMBRE, Ven.Cod_Banco,Ven.Tipo_Emision,Ven.Cuenta_Bancaria"

    
    Case "A" 'Aseguradora
        
        If chkTrasladoAgrupar.Value = vbChecked Then
            strSQL = "select Aseg.Cedula_Juridica as 'CEDULA',Aseg.Nombre as 'NOMBRE', Aseg.Cod_Banco as 'Cod_Banco'" _
                   & ",Aseg.Tipo_Emision as 'Tipo_Emision', Aseg.Cuenta_Cliente  as 'Cuenta_Bancaria',Aseg.Cod_Cuenta" _
                   & ",Pag.Fecha_Vence,SUM(Pag.Monto_Pago) as 'Monto', sum(Pag.Comision_Interna) as 'Comision', sum(Pag.Monto) as 'Total'" _
                   & "  from SEGUROS_REGISTRO Pol" _
                   & " inner join SEGUROS_PAGOS Pag on Pol.NUM_POLIZA = Pag.NUM_POLIZA and Pol.Cod_Aseguradora = Pag.Cod_Aseguradora" _
                   & " inner join SEGUROS_ASEGURADORAS Aseg on Pol.Cod_Aseguradora = Aseg.Cod_Aseguradora" _
                   & " Where Pag.COD_REMESA = " & cboTraslado.ItemData(cboTraslado.ListIndex) & " And Pag.Tesoreria_Solicitud Is Null" _
                   & " group by Pag.Fecha_Vence,Aseg.Cedula_Juridica,Aseg.Nombre,Aseg.Cod_Banco,Aseg.Tipo_Emision,Aseg.Cuenta_Cliente,Aseg.Cod_Cuenta"
        Else
            strSQL = "select Aseg.Cedula_Juridica as 'CEDULA',Aseg.Nombre as 'NOMBRE', Aseg.Cod_Banco as 'Cod_Banco'" _
                   & ",Aseg.Tipo_Emision as 'Tipo_Emision', Aseg.Cuenta_Cliente  as 'Cuenta_Bancaria',Aseg.Cod_Cuenta" _
                   & ",SUM(Pag.Monto_Pago) as 'Monto', sum(Pag.Comision_Interna) as 'Comision', sum(Pag.Monto) as 'Total'" _
                   & "  from SEGUROS_REGISTRO Pol" _
                   & " inner join SEGUROS_PAGOS Pag on Pol.NUM_POLIZA = Pag.NUM_POLIZA and Pol.Cod_Aseguradora = Pag.Cod_Aseguradora" _
                   & " inner join SEGUROS_ASEGURADORAS Aseg on Pol.Cod_Aseguradora = Aseg.Cod_Aseguradora" _
                   & " Where Pag.COD_REMESA = " & cboTraslado.ItemData(cboTraslado.ListIndex) & " And Pag.Tesoreria_Solicitud Is Null" _
                   & " group by Aseg.Cedula_Juridica,Aseg.Nombre,Aseg.Cod_Banco,Aseg.Tipo_Emision,Aseg.Cuenta_Cliente,Aseg.Cod_Cuenta"
        End If
End Select
Call OpenRecordSet(rs, strSQL)

PrgBar.Max = rs.RecordCount + 1
PrgBar.Value = 1
PrgBar.Visible = True


Do While Not rs.EOF

 'Graba y Devuelve el registro Maestro en Tesoreria
 
 If rs!Monto > 0 And (rs!Tipo_Emision = "CK" Or rs!Tipo_Emision = "TE" Or rs!Tipo_Emision = "ND") Then
 
    lngSolicitud = fxMaestroTesoreria(rs!Tipo_Emision, rs!Cod_Banco, rs!Monto, Trim(rs!Cedula) _
                   , rs!Nombre, 0, "Remesa.:" & cboTraslado.ItemData(cboTraslado.ListIndex), 0, mConcepto _
                   , rs!Cuenta_Bancaria, vFecha, mUnidad)
                   
                   
    'Asiento
    If vTipo = "V" Then
        Call sbCreaDetalle(lngSolicitud, fxCtaBanco(rs!Cod_Banco), rs!Monto, "H", 1, mUnidad)
        Call sbCreaDetalle(lngSolicitud, Cuenta(2), rs!Monto, "D", 2, mUnidad) 'Cuenta de Gasto Vendedores
    Else
        Call sbCreaDetalle(lngSolicitud, fxCtaBanco(rs!Cod_Banco), rs!Monto, "H", 1, mUnidad) 'Bancos
        Call sbCreaDetalle(lngSolicitud, rs!cod_cuenta, rs!Monto, "D", 2, mUnidad) 'CxP Aseguradora
        Call sbCreaDetalle(lngSolicitud, Cuenta(0), rs!Total, "D", 3, mUnidad) 'Transitoria
        Call sbCreaDetalle(lngSolicitud, Cuenta(1), rs!Comision, "H", 4, mUnidad) 'Ingresos
        Call sbCreaDetalle(lngSolicitud, rs!cod_cuenta, rs!Monto, "H", 5, mUnidad) 'CxP Aseguradora
    End If

 Else 'Monto a Girar > 0
   
   lngSolicitud = 0
 
 End If
  
 'Actualiza Campo Tesoreria
 If vTipo = "V" Then
        strSQL = "update Pag set Pag.Tesoreria_Solicitud_Comision = " & lngSolicitud & ",Pag.Comision_Vendedor_Estado = 'T'" _
               & " from SEGUROS_REGISTRO Pol" _
               & " inner join SEGUROS_Pagos Pag on Pol.num_poliza = Pag.Num_Poliza and Pol.Cod_Aseguradora = Pag.Cod_Aseguradora" _
               & " inner join SEGUROS_vendedores Ven on Pol.cod_Vendedor = Ven.cod_Vendedor" _
               & " Where Pag.cod_Remesa_Comision = " & cboTraslado.ItemData(cboTraslado.ListIndex) _
               & " and Ven.cedula = '" & rs!Cedula & "'"
 Else
        If chkTrasladoAgrupar.Value = vbChecked Then
           strSQL = "update SEGUROS_pagos set Tesoreria_Solicitud = " & lngSolicitud & ",Estado_Pago = 'T'" _
                  & " Where cod_Remesa = " & cboTraslado.ItemData(cboTraslado.ListIndex) _
                  & " and Fecha_Vence = '" & Format(rs!Fecha_Vence, "yyyy/mm/dd") & "'"
        Else
           strSQL = "update SEGUROS_pagos set Tesoreria_Solicitud = " & lngSolicitud & ",Estado_Pago = 'T'" _
                  & " Where cod_Remesa = " & cboTraslado.ItemData(cboTraslado.ListIndex)
        End If
 End If
 Call ConectionExecute(strSQL)
 
 If PrgBar.Max > PrgBar.Value Then PrgBar.Value = PrgBar.Value + 1
 rs.MoveNext
 
Loop
rs.Close




'Actualiza y Carga Remesa
strSQL = "update SEGUROS_REMESAS SET Estado = 'T',tesoreria_Fecha = dbo.MyGetdate(), Tesoreria_Usuario = '" & glogon.Usuario _
       & "'  Where cod_remesa = " & cboTraslado.ItemData(cboTraslado.ListIndex)
Call ConectionExecute(strSQL)


 'Actualiza Bitacora
 Call Bitacora("Registra", "Remesa de Seguros (Traslado): " & cboTraslado.ItemData(cboTraslado.ListIndex))

Call sbLimpia


Me.MousePointer = vbDefault

PrgBar.Visible = False

MsgBox "Operaciones Enviadas a Tesoreria Satisfactoriamente...", vbInformation

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub sbReportePendientes()
Dim strSQL As String, rs As New ADODB.Recordset

Dim strTitulo As String
Dim strRuta As String, strInicio As String, strFinal As String


On Error GoTo vError

Me.MousePointer = vbHourglass

strTitulo = "Operaciones pendientes de Traslado a Tesorer�a"


strRuta = SIFGlobal.fxPathReportes("CxC_Tesoreria_Envio.rpt")
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
     
     .WindowTitle = "Solicitudes a trasladar a Tesorer�a"
     
    .ReportFileName = strRuta
    .Formulas(1) = "empresa='" & GLOBALES.gstrNombreEmpresa & "'"
    .Formulas(2) = "fecha='" & Format(fxFechaServidor, "dd/mm/yyyy") & "'"
    .Formulas(3) = "Titulo='" & strTitulo & "'"
    
    strSQL = "{CxC_Cuentas.Autoriza_Estado} = 'F'"

    
    strSQL = strSQL & " and ISNULL({CxC_Cuentas.TESORERIA}) AND {CxC_Cuentas.ESTADO}='A'"
    
    .SelectionFormula = strSQL
    
    .SubreportToChange = "subCkDesembolsos"
    .SelectionFormula = "{DESEMBOLSOS.Operacion} = {?Pm-CxC_Cuentas.Operacion}"
    
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

On Error GoTo vError

Me.MousePointer = vbHourglass


With frmContenedor.Crt
 .Reset
 .WindowShowGroupTree = True
 .WindowShowPrintSetupBtn = True
 .WindowShowRefreshBtn = True
 .WindowShowSearchBtn = True
 .WindowState = crptMaximized
 .WindowTitle = "OPERACIONES ENVIADAS A TESORERIA"

 .Connect = glogon.ConectRPT

.ReportFileName = SIFGlobal.fxPathReportes("CxC_Tesoreria_Envio_Rec.rpt")
.Formulas(1) = "fxEmpresa='" & GLOBALES.gstrNombreEmpresa & "'"
.Formulas(2) = "fxFecha='" & Format(fxFechaServidor, "dd/mm/yyyy") & "'"
.Formulas(3) = "fxTitulo='Desembolsos Solicitados en Tesorer�a'"
.Formulas(4) = "fxUsuario='" & glogon.Usuario & "'"
'.Formulas(5) = "fxSubTitulo='INICIO : " & Format(dtpRepInicio.Value, "dd/mm/yyyy") & " CORTE : " & Format(dtpRepCorte.Value, "dd/mm/yyyy") & "'"
'
'strSQL = "{TES_TRANSACCIONES.FECHA_SOLICITUD} in date(" & Format(dtpRepInicio.Value, "yyyy,mm,dd") & ") to date(" _
'    & Format(dtpRepCorte.Value, "yyyy,mm,dd") & ") and {TES_TRANSACCIONES.MODULO} ='CC'"

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


vModulo = 17

Set imgBanner.Picture = frmContenedor.imgBanner_Tramites.Picture

strSQL = "select cod_aseguradora as 'IdX', rtrim( nombre ) as 'ItmX' from seguros_Aseguradoras"
Call sbCbo_Llena_New(cboAseguradora, strSQL, True, True)

 ssTab.Tab = 0
 
     vPaso = True
        cboTipo.Clear
        cboTipo.AddItem "Aseguradora"
        cboTipo.AddItem "Vendedores"
        cboTipo.AddItem "Comercializadoras"
        cboTipo.Text = "Aseguradora"
    vPaso = False

 
 Call Formularios(Me)
 Call RefrescaTags(Me)
 
 Call sbLimpia
End Sub

