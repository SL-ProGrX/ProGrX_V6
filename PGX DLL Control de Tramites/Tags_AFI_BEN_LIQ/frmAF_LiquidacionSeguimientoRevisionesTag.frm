VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.Ocx"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Begin VB.Form frmAF_LiquidacionRevision 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Revisión de Liquidaciones Obrero/Patronales"
   ClientHeight    =   7830
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11475
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "frmAF_LiquidacionSeguimientoRevisionesTag.frx":0000
   ScaleHeight     =   7830
   ScaleWidth      =   11475
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer TimerX 
      Interval        =   20
      Left            =   10440
      Top             =   120
   End
   Begin VB.TextBox txtId 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9960
      Locked          =   -1  'True
      TabIndex        =   44
      Top             =   600
      Width           =   1335
   End
   Begin VB.TextBox txtCedula 
      Alignment       =   2  'Center
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
      Left            =   2160
      TabIndex        =   11
      Top             =   600
      Width           =   1935
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   11400
      Top             =   1200
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAF_LiquidacionSeguimientoRevisionesTag.frx":6852
            Key             =   "IMG1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAF_LiquidacionSeguimientoRevisionesTag.frx":D0B4
            Key             =   "IMG2"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAF_LiquidacionSeguimientoRevisionesTag.frx":13916
            Key             =   "IMG3"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAF_LiquidacionSeguimientoRevisionesTag.frx":1A178
            Key             =   "IMG4"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAF_LiquidacionSeguimientoRevisionesTag.frx":209DA
            Key             =   "IMG5"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAF_LiquidacionSeguimientoRevisionesTag.frx":2723C
            Key             =   "IMG6"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgSemaforos 
      Left            =   11160
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
            Picture         =   "frmAF_LiquidacionSeguimientoRevisionesTag.frx":2DA9E
            Key             =   "verde"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAF_LiquidacionSeguimientoRevisionesTag.frx":2DBBC
            Key             =   "amarillo"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAF_LiquidacionSeguimientoRevisionesTag.frx":2DCE2
            Key             =   "rojo"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAF_LiquidacionSeguimientoRevisionesTag.frx":2DE0C
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAF_LiquidacionSeguimientoRevisionesTag.frx":2DF1E
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAF_LiquidacionSeguimientoRevisionesTag.frx":2E035
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAF_LiquidacionSeguimientoRevisionesTag.frx":2E136
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAF_LiquidacionSeguimientoRevisionesTag.frx":2E26D
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAF_LiquidacionSeguimientoRevisionesTag.frx":2E382
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList3 
      Left            =   9000
      Top             =   0
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
            Picture         =   "frmAF_LiquidacionSeguimientoRevisionesTag.frx":2E4A6
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAF_LiquidacionSeguimientoRevisionesTag.frx":34D08
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAF_LiquidacionSeguimientoRevisionesTag.frx":3B56A
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAF_LiquidacionSeguimientoRevisionesTag.frx":3B684
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6615
      Left            =   240
      TabIndex        =   0
      Top             =   1080
      Width           =   11055
      _ExtentX        =   19500
      _ExtentY        =   11668
      _Version        =   393216
      Style           =   1
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Liquidaciones"
      TabPicture(0)   =   "frmAF_LiquidacionSeguimientoRevisionesTag.frx":3B7A2
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "tlbRefresh"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "imgRefresh"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "vGrid"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Detalle"
      TabPicture(1)   =   "frmAF_LiquidacionSeguimientoRevisionesTag.frx":3B7BE
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lswDetalle"
      Tab(1).Control(1)=   "lblCuenta"
      Tab(1).Control(2)=   "Label13(1)"
      Tab(1).Control(3)=   "Label28"
      Tab(1).Control(4)=   "lblCausa"
      Tab(1).Control(5)=   "Label25"
      Tab(1).Control(6)=   "lblBanco"
      Tab(1).Control(7)=   "Label21"
      Tab(1).Control(8)=   "lblEmitir"
      Tab(1).Control(9)=   "Line1"
      Tab(1).Control(10)=   "Label23"
      Tab(1).Control(11)=   "Label22"
      Tab(1).Control(12)=   "lblFechaLiq"
      Tab(1).Control(13)=   "lblAccion"
      Tab(1).Control(14)=   "Label19"
      Tab(1).Control(15)=   "lblRige"
      Tab(1).Control(16)=   "Label17"
      Tab(1).Control(17)=   "lblRetenido"
      Tab(1).Control(18)=   "Label15"
      Tab(1).Control(19)=   "lblGirado"
      Tab(1).Control(20)=   "Label13(0)"
      Tab(1).Control(21)=   "lblTotalLiq"
      Tab(1).Control(22)=   "Label11"
      Tab(1).Control(23)=   "Label10"
      Tab(1).Control(24)=   "lblCapitaliza"
      Tab(1).Control(25)=   "Label7"
      Tab(1).Control(26)=   "lblAporte"
      Tab(1).Control(27)=   "lblAhorro"
      Tab(1).Control(28)=   "Label3"
      Tab(1).Control(29)=   "Label4"
      Tab(1).Control(30)=   "lblTipoLiq"
      Tab(1).Control(31)=   "Label1"
      Tab(1).Control(32)=   "lblBoleta"
      Tab(1).ControlCount=   33
      TabCaption(2)   =   "Seguimiento"
      TabPicture(2)   =   "frmAF_LiquidacionSeguimientoRevisionesTag.frx":3B7DA
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "vGridSeguimiento"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Revisión"
      TabPicture(3)   =   "frmAF_LiquidacionSeguimientoRevisionesTag.frx":3B7F6
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "txtObservacion"
      Tab(3).Control(1)=   "cboEtiquetas"
      Tab(3).Control(2)=   "tlbAplicar"
      Tab(3).Control(3)=   "lswErrores"
      Tab(3).Control(4)=   "Label8(1)"
      Tab(3).Control(5)=   "Label2(0)"
      Tab(3).Control(6)=   "Label27"
      Tab(3).ControlCount=   7
      Begin VB.TextBox txtObservacion 
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
         Height          =   1575
         Left            =   -73320
         MaxLength       =   995
         MultiLine       =   -1  'True
         TabIndex        =   5
         Top             =   4080
         Width           =   9135
      End
      Begin VB.ComboBox cboEtiquetas 
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
         ItemData        =   "frmAF_LiquidacionSeguimientoRevisionesTag.frx":3B812
         Left            =   -73320
         List            =   "frmAF_LiquidacionSeguimientoRevisionesTag.frx":3B814
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   720
         Width           =   5295
      End
      Begin FPSpreadADO.fpSpread vGrid 
         Height          =   5535
         Left            =   240
         TabIndex        =   2
         Top             =   960
         Width           =   10575
         _Version        =   524288
         _ExtentX        =   18653
         _ExtentY        =   9763
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
         MaxCols         =   7
         ScrollBarExtMode=   -1  'True
         SpreadDesigner  =   "frmAF_LiquidacionSeguimientoRevisionesTag.frx":3B816
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin FPSpreadADO.fpSpread vGridSeguimiento 
         Height          =   5775
         Left            =   -74760
         TabIndex        =   3
         Top             =   600
         Width           =   10575
         _Version        =   524288
         _ExtentX        =   18653
         _ExtentY        =   10186
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
         MaxCols         =   487
         ScrollBarExtMode=   -1  'True
         SpreadDesigner  =   "frmAF_LiquidacionSeguimientoRevisionesTag.frx":3C33C
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin MSComctlLib.Toolbar tlbAplicar 
         Height          =   570
         Left            =   -65520
         TabIndex        =   6
         Top             =   5880
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   1005
         ButtonWidth     =   2117
         ButtonHeight    =   1005
         Style           =   1
         TextAlignment   =   1
         ImageList       =   "ImageList1"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   1
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Aplicar"
               Key             =   "Aplicar"
               Object.ToolTipText     =   "Aplicar Etiqueta"
               ImageIndex      =   1
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ListView lswErrores 
         Height          =   2655
         Left            =   -73320
         TabIndex        =   7
         Top             =   1200
         Width           =   9135
         _ExtentX        =   16113
         _ExtentY        =   4683
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
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Código"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Descripción"
            Object.Width           =   8819
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Aplicado"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Mensaje"
            Object.Width           =   8819
         EndProperty
      End
      Begin MSComctlLib.ListView lswDetalle 
         Height          =   2775
         Left            =   -74760
         TabIndex        =   36
         Top             =   3600
         Width           =   10695
         _ExtentX        =   18865
         _ExtentY        =   4895
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
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Operación"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Código"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Abono"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Saldo Inicial"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Saldo Resultante"
            Object.Width           =   3528
         EndProperty
      End
      Begin MSComctlLib.ImageList imgRefresh 
         Left            =   8760
         Top             =   360
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   1
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmAF_LiquidacionSeguimientoRevisionesTag.frx":3C934
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.Toolbar tlbRefresh 
         Height          =   336
         Left            =   9480
         TabIndex        =   45
         Top             =   480
         Width           =   1212
         _ExtentX        =   2143
         _ExtentY        =   582
         ButtonWidth     =   1984
         ButtonHeight    =   582
         Style           =   1
         TextAlignment   =   1
         ImageList       =   "imgRefresh"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   1
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Refrescar"
               Key             =   "Refrescar"
               Object.ToolTipText     =   "Volver a cargar la información"
               ImageIndex      =   1
            EndProperty
         EndProperty
         MousePointer    =   1
      End
      Begin VB.Label lblCuenta 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
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
         Height          =   315
         Left            =   -66360
         TabIndex        =   47
         Top             =   2520
         Width           =   2295
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Cuenta"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   1
         Left            =   -66360
         TabIndex        =   46
         Top             =   2280
         Width           =   2295
      End
      Begin VB.Label Label28 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Causa"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   -74640
         TabIndex        =   43
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label lblCausa 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
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
         Height          =   315
         Left            =   -73440
         TabIndex        =   42
         Top             =   1080
         Width           =   7095
      End
      Begin VB.Label Label25 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
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
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   -71760
         TabIndex        =   41
         Top             =   2520
         Width           =   1215
      End
      Begin VB.Label lblBanco 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
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
         Height          =   315
         Left            =   -70560
         TabIndex        =   40
         Top             =   2520
         Width           =   4215
      End
      Begin VB.Label Label21 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Emitir"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   -74640
         TabIndex        =   39
         Top             =   2520
         Width           =   1215
      End
      Begin VB.Label lblEmitir 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
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
         Height          =   315
         Left            =   -73440
         TabIndex        =   38
         Top             =   2520
         Width           =   1455
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         X1              =   -75000
         X2              =   -63960
         Y1              =   3000
         Y2              =   3000
      End
      Begin VB.Label Label23 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Detalle Apliación Abonos "
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
         Height          =   210
         Left            =   -74760
         TabIndex        =   37
         Top             =   3120
         Width           =   2115
      End
      Begin VB.Label Label22 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Fecha Liq."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   -69000
         TabIndex        =   35
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label lblFechaLiq 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
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
         Height          =   315
         Left            =   -67800
         TabIndex        =   34
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label lblAccion 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
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
         Height          =   315
         Left            =   -70560
         TabIndex        =   33
         Top             =   2160
         Width           =   1455
      End
      Begin VB.Label Label19 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "N° Acción"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   -71760
         TabIndex        =   32
         Top             =   2160
         Width           =   1215
      End
      Begin VB.Label lblRige 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
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
         Height          =   315
         Left            =   -73440
         TabIndex        =   31
         Top             =   2160
         Width           =   1455
      End
      Begin VB.Label Label17 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Rige desde"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   -74640
         TabIndex        =   30
         Top             =   2160
         Width           =   1215
      End
      Begin VB.Label lblRetenido 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
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
         Height          =   315
         Left            =   -70560
         TabIndex        =   29
         Top             =   1800
         Width           =   1455
      End
      Begin VB.Label Label15 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Retenido"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   -71760
         TabIndex        =   28
         Top             =   1800
         Width           =   1215
      End
      Begin VB.Label lblGirado 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
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
         Height          =   315
         Left            =   -67800
         TabIndex        =   27
         Top             =   1800
         Width           =   1455
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Girado"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   0
         Left            =   -69000
         TabIndex        =   26
         Top             =   1800
         Width           =   1215
      End
      Begin VB.Label lblTotalLiq 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
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
         Height          =   315
         Left            =   -73440
         TabIndex        =   25
         Top             =   1800
         Width           =   1455
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Total Liq."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   -74640
         TabIndex        =   24
         Top             =   1800
         Width           =   1215
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Capitalización"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   -69000
         TabIndex        =   23
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label lblCapitaliza 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
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
         Height          =   315
         Left            =   -67800
         TabIndex        =   22
         Top             =   1440
         Width           =   1455
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Aporte"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   -71760
         TabIndex        =   21
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label lblAporte 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
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
         Height          =   315
         Left            =   -70560
         TabIndex        =   20
         Top             =   1440
         Width           =   1455
      End
      Begin VB.Label lblAhorro 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
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
         Height          =   315
         Left            =   -73440
         TabIndex        =   19
         Top             =   1440
         Width           =   1455
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Ahorro"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   -74640
         TabIndex        =   18
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Tipo Liq."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   -71760
         TabIndex        =   17
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label lblTipoLiq 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
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
         Height          =   315
         Left            =   -70560
         TabIndex        =   16
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Boleta N°"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   -74640
         TabIndex        =   15
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label lblBoleta 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
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
         Height          =   315
         Left            =   -73440
         TabIndex        =   14
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label Label8 
         Caption         =   "Observación"
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
         Index           =   1
         Left            =   -74760
         TabIndex        =   10
         Top             =   4080
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "Etiqueta"
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
         Index           =   0
         Left            =   -74760
         TabIndex        =   9
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label27 
         Caption         =   "Omisiones"
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
         Left            =   -74760
         TabIndex        =   8
         Top             =   1200
         Width           =   855
      End
   End
   Begin VB.Label LblOperacion 
      Caption         =   "Cedula"
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
      Left            =   1200
      TabIndex        =   13
      Top             =   600
      Width           =   975
   End
   Begin VB.Label lblNombre 
      Alignment       =   2  'Center
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
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
      Left            =   4080
      TabIndex        =   12
      ToolTipText     =   "Nombre"
      Top             =   600
      Width           =   5895
   End
   Begin VB.Image Image2 
      Height          =   360
      Left            =   1800
      Picture         =   "frmAF_LiquidacionSeguimientoRevisionesTag.frx":3CA59
      Top             =   120
      Width           =   360
   End
   Begin VB.Label lblNombreUsuario 
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
      Left            =   2280
      TabIndex        =   1
      Top             =   120
      Width           =   3735
   End
End
Attribute VB_Name = "frmAF_LiquidacionRevision"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vPaso As Boolean

Private Sub cboEtiquetas_Click()
If vPaso Then Exit Sub
Call sbCargarObservacion
End Sub

Private Sub Form_Activate()
vModulo = 8
End Sub

Private Sub Form_Load()
vModulo = 8

SSTab1.Tab = 0

lblNombreUsuario.Caption = glogon.Usuario



Call Formularios(Me)
Call RefrescaTags(Me)

End Sub


Private Sub sbCargarListaLiquidaciones(Optional ByVal strCedula As String = Empty)
' Carga Lista de afiliaciones
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo error
Me.MousePointer = vbHourglass

If strCedula = "" Then
    strSQL = "select Top 3000 'B',L.cedula, S.nombre,L.usuario,L.cod_remesa,R.usuario,L.consec from Liquidacion L" _
            & " inner join Socios S on L.Cedula = S.cedula left join AFI_REMESAS_LIQ R on L.cod_remesa = R.cod_remesa" _
            & " Where L.ANALISTA_REVISION  Is Null"
Else
  strSQL = "select  'B', L.cedula, S.nombre,L.usuario,L.cod_remesa,R.usuario,L.consec from Liquidacion L" _
            & " inner join Socios S on L.Cedula = S.cedula left join AFI_REMESAS_LIQ R on L.cod_remesa = R.cod_remesa" _
            & " Where L.ANALISTA_REVISION  Is Null and L.cedula = '" & strCedula & "'"
End If


vPaso = True
Call sbCargaGrid(vGrid, 7, strSQL)
vGrid.MaxRows = vGrid.MaxRows - 1
vPaso = False

Me.MousePointer = vbDefault

Exit Sub
    
error:
    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub sbCargarGridSeguimiento(pCedula As String, pDocumento As String)
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

vGridSeguimiento.MaxCols = 4
vGridSeguimiento.MaxRows = 0

If pCedula = Empty Or pCedula = "" Then Exit Sub

Me.MousePointer = vbHourglass

strSQL = "select T.DESCRIPCION, OT.NOTAS, OT.REGISTRO_FECHA, OT.REGISTRO_USUARIO" _
       & " from SIF_CONTROL_TAGS OT inner join SIF_TAGS T on OT.TAG_CODIGO = T.TAG_CODIGO" _
       & " where OT.codigo = '" & pCedula & "' and OT.cod_Modulo = 'LIQ' and OT.Documento = '" & pDocumento & "'"
Call OpenRecordSet(rs, strSQL)

Do While Not rs.EOF
    vGridSeguimiento.MaxRows = vGridSeguimiento.MaxRows + 1
    vGridSeguimiento.Row = vGridSeguimiento.MaxRows
  
    vGridSeguimiento.Col = 1
    vGridSeguimiento.Text = rs!Descripcion
    vGridSeguimiento.TextTip = TextTipFixed
    vGridSeguimiento.TextTipDelay = 1000
    vGridSeguimiento.CellNote = "Usuario: " & rs!registro_usuario & "[" & rs!Registro_Fecha & "]"
            
    vGridSeguimiento.Col = 2
    vGridSeguimiento.Value = IIf(IsNull(rs!notas), "", rs!notas)
    
    vGridSeguimiento.Col = 3
    vGridSeguimiento.Value = IIf(IsNull(rs!Registro_Fecha), "", rs!Registro_Fecha)
    
    vGridSeguimiento.Col = 4
    vGridSeguimiento.Value = IIf(IsNull(rs!registro_usuario), "", rs!registro_usuario)
    
    vGridSeguimiento.RowHeight(vGridSeguimiento.Row) = vGridSeguimiento.MaxTextRowHeight(vGridSeguimiento.Row)
    rs.MoveNext
Loop
rs.Close

Me.MousePointer = vbDefault

Exit Sub

vError:
    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub sbCargarListaErrores()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListItem

If txtCedula = Empty Then
    Exit Sub
End If

With lswErrores
 .ListItems.Clear
 
 strSQL = "select E.ID_ERROR,E.DESCRIPCION,ER.ID_ERROR as asignado, ISNULL(ER.APLICADO,'N') AS APLICADO, E.MENSAJE, ER.LINEA_ERR" _
        & " from sif_Omisiones E left join SIF_OMISIONESG ER on E.ID_ERROR = ER.ID_ERROR" _
        & " and ER.cedula = '" & txtCedula.Text & "' and ER.Modulo = 'LIQ' and ER.Codigo = '" & txtCedula.Text _
        & "' and ER.Documento = '" & txtId.Text & "'" _
        & " where E.ACTIVO = '1'  and E.ID_ERROR in(select ID_ERROR from SIF_OMISIONES_MODULOS where cod_modulo = 'LIQ') " _
        & " order by E.ID_ERROR"
        
 rs.Open strSQL, glogon.Conection, adOpenForwardOnly
 Do While Not rs.EOF
  Set itmX = .ListItems.Add(, , rs!ID_ERROR)
      itmX.SubItems(1) = rs!Descripcion
      If Not IsNull(rs!Asignado) Then
         itmX.Checked = vbChecked
         itmX.ForeColor = vbBlue
         itmX.Tag = rs!LINEA_ERR
      End If
      itmX.SubItems(2) = rs!APLICADO
      itmX.SubItems(3) = rs!Mensaje
  rs.MoveNext
 Loop
 rs.Close
End With
End Sub

Private Sub sbCargarCombosEtiquetas()
Dim strSQL As String

On Error GoTo vError

    
    strSQL = "SELECT CT.TAG_CODIGO + ' - ' +  rtrim(CT.DESCRIPCION) as 'ItmX'" _
            & " FROM SIF_TAGS CT INNER JOIN SIF_TAGS_GRUPOS CTG ON CT.TAG_CODIGO = CTG.TAG_CODIGO" _
            & " INNER JOIN SIF_GRPUSERS CGU ON CTG.COD_GRUPO = CGU.COD_GRUPO" _
            & " WHERE CT.ACTIVO = 1 AND CGU.USUARIO = '" & glogon.Usuario _
            & "' and  CT.TAG_CODIGO in(select TAG_CODIGO from SIF_TAGS_MODULOS where cod_modulo = 'LIQ')" _
            & " order by CT.TAG_CODIGO"
    vPaso = True
    Call sbLlenaCbo(cboEtiquetas, strSQL, False, False)
    vPaso = False
    Call cboEtiquetas_Click
    
    Exit Sub
vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub sbCargarObservacion()
Dim strSQL As String, rs As New ADODB.Recordset, i As Integer

On Error GoTo vError
    
    strSQL = "select ISNULL(MENSAJE,'') from SIF_TAGS_AVISOS where TAG_CODIGO = '" & SIFGlobal.fxCodText(cboEtiquetas.Text) & "'"
    Call OpenRecordSet(rs, strSQL)
    
    If Not rs.EOF Then
        txtObservacion = rs.Fields(0) & vbNewLine
    Else
        txtObservacion = Empty
    End If
    
    For i = 1 To lswErrores.ListItems.Count
        If lswErrores.ListItems(i).Checked = True Then
            If lswErrores.ListItems(i).SubItems(2) = "N" Then
                If txtObservacion = Empty Then
                    txtObservacion.Text = "-" & lswErrores.ListItems(i).SubItems(3)
                Else
                    txtObservacion.Text = txtObservacion.Text & vbNewLine & "-" & lswErrores.ListItems(i).SubItems(3)
                End If
            End If
        End If
    Next
    
    Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub



Private Sub lswErrores_ItemCheck(ByVal Item As MSComctlLib.ListItem)
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError
    
    If Item.SubItems(2) = "S" Then
        Item.Checked = True
        If MsgBox("El error ya fué aplicado desea agregar únicamente la nota", vbOKCancel) = vbOK Then
            If txtObservacion = Empty Then
              txtObservacion.Text = " - " & Item.SubItems(1)
            Else
              txtObservacion.Text = txtObservacion.Text & vbCrLf & " - " & Item.SubItems(1)
            End If
        End If
        Exit Sub
    End If
    
    If Item.Checked Then
    
      strSQL = "insert SIF_OMISIONESG (cedula,ID_ERROR,MODULO,CODIGO,DOCUMENTO,REGISTRO_FECHA,REGISTRO_USUARIO) values('" & txtCedula.Text _
             & "'," & Item.Text & ",'LIQ','" & txtCedula.Text & "','" & txtId.Text & "',dbo.MyGetdate(),'" & glogon.Usuario & "')"
      Call ConectionExecute(strSQL)
             
      strSQL = "select max(LINEA_ERR) as 'Linea' from SIF_OMISIONESG where codigo = '" & txtCedula.Text & "' and Documento = '" & txtId & "' and ID_ERROR = " & Item.Text
      Call OpenRecordSet(rs, strSQL)
          Item.Tag = rs!Linea
      rs.Close
      
      If txtObservacion = Empty Then
        txtObservacion.Text = " - " & Item.SubItems(1)
      Else
        txtObservacion.Text = txtObservacion.Text & vbCrLf & " - " & Item.SubItems(1)
      End If
      
    Else
      strSQL = "delete SIF_OMISIONESG where LINEA_ERR = " & Item.Tag
      Call ConectionExecute(strSQL)
      Item.Tag = ""

      Call sbCargarObservacion
    End If
    
    Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
Select Case SSTab1.Tab
   Case 2
     Call sbCargarGridSeguimiento(txtCedula.Text, txtId.Text)
   Case 3
     Call sbCargarListaErrores
     Call sbCargarCombosEtiquetas
End Select
End Sub


Private Sub TimerX_Timer()
TimerX.Interval = 0
TimerX.Enabled = False

Call sbCargarListaLiquidaciones

End Sub

Private Sub tlbAplicar_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo vError
    Me.MousePointer = vbHourglass
    
    If Trim(cboEtiquetas.Text) = Empty Then
        MsgBox "Debe seleccionar la etiqueta que desea plicar"
        Me.MousePointer = vbDefault
        Exit Sub
    End If

    If MsgBox("Está seguro que sea aplicar la etiqueta!", vbExclamation + vbYesNo) = vbNo Then
        Me.MousePointer = vbDefault
        Exit Sub
    End If
      Call sbSIFRegistraTags(txtCedula.Text, SIFGlobal.fxCodText(cboEtiquetas.Text), txtObservacion, txtId.Text, "LIQ", txtId.Text)
   
    
    Call sbAplicarErrores
    Call sbCargarListaLiquidaciones
    txtCedula.SetFocus
    SSTab1.Tab = 0
    
    
    Me.MousePointer = vbDefault
    Exit Sub
vError:
    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub sbAplicarErrores()
'' Procedimiento para colocar los errores ingresados en aplicados
Dim Linea As String, strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError
   
    strSQL = "update SIF_OMISIONESG SET APLICADO = 'S' WHERE cedula = '" & txtCedula.Text _
           & "' AND MODULO = 'LIQ' AND CODIGO = '" & txtCedula.Text & "' AND DOCUMENTO = '" & txtId.Text & "'"
    Call ConectionExecute(strSQL)

    Exit Sub
vError:
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub sbCargaDetalle(vConsec As Long)
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListItem

On Error GoTo vError

Call sbLimpiaDatos

strSQL = "Select L.Consec,L.cedula,S.nombre, L.AHORRO_LIQ,L.APORTE_LIQ,L.CAPITALIZADO_LIQ,L.TOTALBRUTO,L.TNETO" _
       & " ,L.RETENIDO,L.AC_BOLETA,L.AC_FECHA,L.FECLIQ,L.TDOCUMENTO , T.DESCRIPCION as 'Banco'" _
       & " ,R.DESCRIPCION as 'Causa',case when estadoactliq = 'A' then 'Ren. Asociación'" _
       & " Else 'Ren. Partronal' end as 'Tipo', isnull(Cta.Cuenta,'') as 'Cuenta'" _
       & " from liquidacion L inner join Socios S on L.cedula = S.cedula" _
       & " inner join  TES_BANCOS  T  on L.cod_banco  = T.ID_BANCO" _
       & " inner join causas_renuncias R on L.ID_CAUSA = R.ID_CAUSA" _
       & " Left join  CUENTAS_AHORROS Cta on L.cod_banco  = Cta.ID_BANCO  and L.CEDULA = Cta.CEDULA " _
       & " Where L.CONSEC = " & vConsec

Call OpenRecordSet(rs, strSQL)
If Not rs.EOF Then
 
   txtCedula.Text = rs!Cedula
   lblNombre.Caption = rs!Nombre
   txtId.Text = rs!Consec
   
   lblBoleta.Caption = vConsec
   lblAhorro.Caption = Format(rs!AHORRO_LIQ, "Standard")
   lblAporte.Caption = Format(rs!APORTE_LIQ, "Standard")
   lblCapitaliza.Caption = Format(rs!CAPITALIZADO_LIQ, "Standard")
   lblTotalLiq.Caption = Format(rs!TOTALBRUTO, "Standard")
   lblGirado.Caption = Format(rs!TNETO, "Standard")
   lblRetenido.Caption = Format(rs!RETENIDO, "Standard")
   lblAccion.Caption = IIf(IsNull(rs!AC_BOLETA), "No Aplica", rs!AC_BOLETA)
   lblFechaLiq.Caption = rs!fecLiq
   lblRige.Caption = IIf(IsNull(rs!AC_FECHA), "No Aplica", Format(rs!AC_FECHA, "dd/mm/yyyy"))
   lblEmitir.Caption = rs!TDOCUMENTO
   lblBanco.Caption = rs!Banco
   lblCausa.Caption = rs!Causa
   lblTipoLiq.Caption = rs!Tipo
   lblCuenta.Caption = rs!Cuenta
End If
rs.Close

strSQL = "select ID_SOLICITUD,CODIGO,LIQ_ABONO as 'Abono',LIQ_SALDO as 'Saldo'," _
        & " (LIQ_SALDO -LIQ_AMORTIZA) as 'Resultante'" _
       & " from LIQUIDA_DETALLE where CONSEC = " & vConsec & ""
rs.Open strSQL, glogon.Conection, adOpenForwardOnly
With lswDetalle
 Do While Not rs.EOF
  Set itmX = .ListItems.Add(, , rs!id_Solicitud)
      itmX.SubItems(1) = rs!Codigo
      itmX.SubItems(2) = Format(rs!abono, "Standard")
      itmX.SubItems(3) = Format(rs!Saldo, "Standard")
      itmX.SubItems(4) = Format(rs!Resultante, "Standard")
 rs.MoveNext
 Loop
End With
 rs.Close

Exit Sub

vError:
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub sbLimpiaDatos()
    lswDetalle.ListItems.Clear
    lblBoleta.Caption = ""
    lblAhorro.Caption = ""
    lblAporte.Caption = ""
    lblCapitaliza.Caption = ""
    lblTotalLiq.Caption = ""
    lblGirado.Caption = ""
    lblRetenido.Caption = ""
    lblAccion.Caption = ""
    lblFechaLiq.Caption = ""
    lblRige.Caption = ""
    lblEmitir.Caption = ""
    lblBanco.Caption = ""
    lblCausa.Caption = ""
    lblTipoLiq.Caption = ""
End Sub

Private Sub tlbRefresh_ButtonClick(ByVal Button As MSComctlLib.Button)
Call TimerX_Timer
End Sub

Private Sub txtCedula_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn And Trim(txtCedula.Text) <> "" Then Call sbCargarListaLiquidaciones(txtCedula.Text)
End Sub

Private Sub vGrid_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)

On Error GoTo vError

If vPaso Then Exit Sub

vGrid.Row = Row
vGrid.Col = 7
SSTab1.Tab = 1

Call sbCargaDetalle(vGrid.Text)

vError:

End Sub
