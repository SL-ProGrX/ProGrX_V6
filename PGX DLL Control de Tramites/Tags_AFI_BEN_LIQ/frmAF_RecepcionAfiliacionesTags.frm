VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.Ocx"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.controls.v22.1.0.ocx"
Begin VB.Form frmAF_RecepcionAfiliacionesTags 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Recepción de Afiliaciones"
   ClientHeight    =   8145
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11985
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8145
   ScaleWidth      =   11985
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab 
      Height          =   7935
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   13996
      _Version        =   393216
      Style           =   1
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Recepción - Devolución"
      TabPicture(0)   =   "frmAF_RecepcionAfiliacionesTags.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Image1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "LblCedula"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label3"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label4"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "PrgBar"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "tlbAplicar"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "lswCedula"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "cmdAgregar"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "txtCedula"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "optRecepcion"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "optDevolucion"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "cboBoleta"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).ControlCount=   12
      TabCaption(1)   =   "Pendientes"
      TabPicture(1)   =   "frmAF_RecepcionAfiliacionesTags.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "optPendDevolucion"
      Tab(1).Control(1)=   "optPendRecepcion"
      Tab(1).Control(2)=   "vGrid"
      Tab(1).Control(3)=   "tlbBuscar"
      Tab(1).Control(4)=   "Image2"
      Tab(1).Control(5)=   "Label2"
      Tab(1).ControlCount=   6
      TabCaption(2)   =   "Consultas"
      TabPicture(2)   =   "frmAF_RecepcionAfiliacionesTags.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "btnExport"
      Tab(2).Control(1)=   "vGridConsulta"
      Tab(2).Control(2)=   "dtpInicio"
      Tab(2).Control(3)=   "dtpCorte"
      Tab(2).Control(4)=   "txtBuscaUsuario"
      Tab(2).Control(5)=   "txtBuscaCedula"
      Tab(2).Control(6)=   "txtBuscaBoleta"
      Tab(2).Control(7)=   "Label1(3)"
      Tab(2).Control(8)=   "Label1(2)"
      Tab(2).Control(9)=   "Label1(1)"
      Tab(2).Control(10)=   "Label1(0)"
      Tab(2).ControlCount=   11
      Begin XtremeSuiteControls.PushButton btnExport 
         Height          =   375
         Left            =   -64440
         TabIndex        =   27
         Top             =   960
         Width           =   495
         _Version        =   1441793
         _ExtentX        =   873
         _ExtentY        =   661
         _StockProps     =   79
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         Picture         =   "frmAF_RecepcionAfiliacionesTags.frx":0054
      End
      Begin VB.ComboBox cboBoleta 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   3120
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   840
         Width           =   3495
      End
      Begin VB.OptionButton optPendDevolucion 
         Caption         =   "Devolución"
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
         Left            =   -70200
         TabIndex        =   11
         Top             =   600
         Width           =   1335
      End
      Begin VB.OptionButton optPendRecepcion 
         Caption         =   "Recepción"
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
         Left            =   -71640
         TabIndex        =   10
         Top             =   600
         Width           =   1335
      End
      Begin VB.OptionButton optDevolucion 
         Caption         =   "Devolución"
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
         Left            =   9120
         TabIndex        =   6
         Top             =   840
         Width           =   1335
      End
      Begin VB.OptionButton optRecepcion 
         Caption         =   "Recepción"
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
         Left            =   7800
         TabIndex        =   4
         Top             =   840
         Value           =   -1  'True
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
         Height          =   345
         Left            =   840
         TabIndex        =   2
         Top             =   840
         Width           =   2295
      End
      Begin VB.CommandButton cmdAgregar 
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6720
         TabIndex        =   1
         Top             =   840
         Width           =   495
      End
      Begin MSComctlLib.ListView lswCedula 
         Height          =   5775
         Left            =   240
         TabIndex        =   7
         Top             =   1320
         Width           =   11295
         _ExtentX        =   19923
         _ExtentY        =   10186
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
            Text            =   "Cédula"
            Object.Width           =   2999
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Nombre"
            Object.Width           =   8114
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Oficina"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Código"
            Object.Width           =   2999
         EndProperty
      End
      Begin MSComctlLib.Toolbar tlbAplicar 
         Height          =   570
         Left            =   8400
         TabIndex        =   8
         Top             =   7200
         Width           =   3225
         _ExtentX        =   5689
         _ExtentY        =   1005
         ButtonWidth     =   2461
         ButtonHeight    =   1005
         Style           =   1
         TextAlignment   =   1
         ImageList       =   "ImageList1"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   3
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Aplicar"
               Key             =   "Aplicar"
               Object.ToolTipText     =   "Aplicar Etiqueta"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Opciones"
               Key             =   "Opciones"
               Object.ToolTipText     =   "Opciones"
               ImageIndex      =   2
               Style           =   5
               BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                  NumButtonMenus  =   2
                  BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "Limpiar"
                     Text            =   "Limpiar"
                  EndProperty
                  BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "Eliminar"
                     Text            =   "Eliminar Cédula"
                  EndProperty
               EndProperty
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ProgressBar PrgBar 
         Height          =   315
         Left            =   240
         TabIndex        =   9
         Top             =   7320
         Visible         =   0   'False
         Width           =   7935
         _ExtentX        =   13996
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
      End
      Begin FPSpreadADO.fpSpread vGrid 
         Height          =   6495
         Left            =   -74880
         TabIndex        =   13
         Top             =   1200
         Width           =   11415
         _Version        =   524288
         _ExtentX        =   20135
         _ExtentY        =   11456
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
         MaxCols         =   6
         ScrollBarExtMode=   -1  'True
         SpreadDesigner  =   "frmAF_RecepcionAfiliacionesTags.frx":01BE
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin MSComctlLib.Toolbar tlbBuscar 
         Height          =   330
         Left            =   -68760
         TabIndex        =   14
         Top             =   600
         Width           =   2625
         _ExtentX        =   4630
         _ExtentY        =   582
         ButtonWidth     =   1931
         ButtonHeight    =   582
         Style           =   1
         TextAlignment   =   1
         ImageList       =   "ImageList3"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   3
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Buscar"
               Key             =   "Buscar"
               Object.ToolTipText     =   "Buscar"
               ImageIndex      =   5
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Imprimir"
               Key             =   "Imprimir"
               ImageIndex      =   6
            EndProperty
         EndProperty
      End
      Begin FPSpreadADO.fpSpread vGridConsulta 
         Height          =   6255
         Left            =   -74760
         TabIndex        =   15
         Top             =   1440
         Width           =   11055
         _Version        =   524288
         _ExtentX        =   19500
         _ExtentY        =   11033
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
         MaxCols         =   486
         ScrollBarExtMode=   -1  'True
         SpreadDesigner  =   "frmAF_RecepcionAfiliacionesTags.frx":08CD
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin XtremeSuiteControls.DateTimePicker dtpInicio 
         Height          =   330
         Left            =   -74160
         TabIndex        =   18
         Top             =   960
         Width           =   1575
         _Version        =   1441793
         _ExtentX        =   2778
         _ExtentY        =   582
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
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   3
      End
      Begin XtremeSuiteControls.DateTimePicker dtpCorte 
         Height          =   330
         Left            =   -72600
         TabIndex        =   19
         Top             =   960
         Width           =   1575
         _Version        =   1441793
         _ExtentX        =   2778
         _ExtentY        =   582
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
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   3
      End
      Begin XtremeSuiteControls.FlatEdit txtBuscaUsuario 
         Height          =   330
         Left            =   -67080
         TabIndex        =   20
         Top             =   960
         Width           =   1815
         _Version        =   1441793
         _ExtentX        =   3201
         _ExtentY        =   582
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtBuscaCedula 
         Height          =   330
         Left            =   -70920
         TabIndex        =   21
         Top             =   960
         Width           =   1815
         _Version        =   1441793
         _ExtentX        =   3201
         _ExtentY        =   582
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtBuscaBoleta 
         Height          =   330
         Left            =   -69000
         TabIndex        =   25
         Top             =   960
         Width           =   1815
         _Version        =   1441793
         _ExtentX        =   3201
         _ExtentY        =   582
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.Label Label1 
         Height          =   255
         Index           =   3
         Left            =   -68880
         TabIndex        =   26
         Top             =   600
         Width           =   1575
         _Version        =   1441793
         _ExtentX        =   2778
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Boleta Id"
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
         Transparent     =   -1  'True
         WordWrap        =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label1 
         Height          =   255
         Index           =   2
         Left            =   -70800
         TabIndex        =   24
         Top             =   600
         Width           =   1575
         _Version        =   1441793
         _ExtentX        =   2778
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Cédula"
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
         Transparent     =   -1  'True
         WordWrap        =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label1 
         Height          =   255
         Index           =   1
         Left            =   -66960
         TabIndex        =   23
         Top             =   600
         Width           =   1575
         _Version        =   1441793
         _ExtentX        =   2778
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Usuario"
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
         Transparent     =   -1  'True
         WordWrap        =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label1 
         Height          =   255
         Index           =   0
         Left            =   -74160
         TabIndex        =   22
         Top             =   600
         Width           =   2295
         _Version        =   1441793
         _ExtentX        =   4048
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Fechas de Registro"
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
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label4 
         Caption         =   "Boleta"
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
         Left            =   3240
         TabIndex        =   16
         Top             =   600
         Width           =   975
      End
      Begin VB.Image Image2 
         Height          =   480
         Left            =   -74760
         Picture         =   "frmAF_RecepcionAfiliacionesTags.frx":0E8F
         Top             =   480
         Width           =   480
      End
      Begin VB.Label Label2 
         Caption         =   "Afilaiciones  Pendientes de:"
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
         Left            =   -74040
         TabIndex        =   12
         Top             =   720
         Width           =   2295
      End
      Begin VB.Label Label3 
         Caption         =   "Tipo Movimiento"
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
         Left            =   7800
         TabIndex        =   5
         Top             =   600
         Width           =   1695
      End
      Begin VB.Label LblCedula 
         Caption         =   "Cédula"
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
         Left            =   840
         TabIndex        =   3
         Top             =   600
         Width           =   975
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   240
         Picture         =   "frmAF_RecepcionAfiliacionesTags.frx":10A0
         Top             =   600
         Width           =   480
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAF_RecepcionAfiliacionesTags.frx":1292
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAF_RecepcionAfiliacionesTags.frx":7AF4
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAF_RecepcionAfiliacionesTags.frx":E356
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList3 
      Left            =   0
      Top             =   720
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAF_RecepcionAfiliacionesTags.frx":14BB8
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAF_RecepcionAfiliacionesTags.frx":1B41A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAF_RecepcionAfiliacionesTags.frx":21C7C
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAF_RecepcionAfiliacionesTags.frx":21D96
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAF_RecepcionAfiliacionesTags.frx":21EB4
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAF_RecepcionAfiliacionesTags.frx":28716
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmAF_RecepcionAfiliacionesTags"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim itmX As ListItem
Dim mTagRecepcion As String, mTagDevolucion As String
Dim mTagRecepcionDev As String

Private Sub sbParametrosTags()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError
    
    'Elimina Afiliaciones Pendientes de Recibir Antiguas
    strSQL = "exec spAFI_Afiliaciones_Duplicados_Elimina"
    Call ConectionExecute(strSQL)
    
    ' Busca el parámetro del tag de recepción
    strSQL = "select isnull(valor,'') from SIF_PARAMETROS where cod_parametro = '10'"
    Call OpenRecordSet(rs, strSQL)
    If Not rs.EOF Then
        mTagRecepcion = rs.Fields(0)
    Else
        MsgBox "Falta agregar el parámetro 10 en la base de datos"
    End If
    rs.Close
    
    If Not mTagRecepcion = Empty Then
    
        strSQL = "select COUNT(*) FROM sif_tags where TAG_CODIGO = '" & mTagRecepcion & "'"
        Call OpenRecordSet(rs, strSQL)
        If rs.Fields(0) = 0 Then
            mTagRecepcion = Empty
            MsgBox "El código de tag definido el los parámetros para la revisión no existe"
        End If
        rs.Close
        
    End If
    
    '' Busca el parámetro del tag de devolución
    strSQL = "select isnull(valor,'') from SIF_PARAMETROS where cod_parametro = '11'"
    Call OpenRecordSet(rs, strSQL)
    If Not rs.EOF Then
        mTagDevolucion = rs.Fields(0)
    Else
        MsgBox "Falta agregar el parámetro 11 en la base de datos"
    End If
    rs.Close
    
    If Not mTagDevolucion = Empty Then
    
        strSQL = "select COUNT(*) FROM sif_tags where TAG_CODIGO = '" & mTagDevolucion & "'"
        Call OpenRecordSet(rs, strSQL)
        If rs.Fields(0) = 0 Then
            mTagRecepcion = Empty
            MsgBox "El código de tag definido el los parámetros para la revisión no existe"
        End If
        rs.Close
        
    End If
    
    
    strSQL = "select isnull(valor,'') from SIF_PARAMETROS where cod_parametro = '12'"
    Call OpenRecordSet(rs, strSQL)
    If Not rs.EOF Then
        mTagRecepcionDev = rs.Fields(0)
    Else
        MsgBox "Falta agregar el parámetro 12 en la base de datos"
    End If
    rs.Close
    
    If Not mTagRecepcionDev = Empty Then
    
        strSQL = "select COUNT(*) FROM sif_tags where TAG_CODIGO = '" & mTagRecepcionDev & "'"
        Call OpenRecordSet(rs, strSQL)
        If rs.Fields(0) = 0 Then
            mTagRecepcionDev = Empty
            MsgBox "El código de tag definido el los parámetros para la revisión no existe"
        End If
        rs.Close
        
    End If
    
    
    
    Exit Sub
vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub sbCargaInformacion()
Dim rs As New ADODB.Recordset, strSQL As String
Dim vBoleta As String

On Error GoTo vError

    If Not IsNumeric(txtCedula) Or cboBoleta.ListCount = 0 Then
        Exit Sub
    End If
    
    
    If fxValidaNoDuplicados = True Then
        MsgBox "La Cedula se ya fue digitada"
        txtCedula.Text = Empty
        txtCedula.SetFocus
        Exit Sub
    End If
    
    'Valida no agregar en forma mismo tag en forma consecutiva

    vBoleta = SIFGlobal.fxCodText(cboBoleta.Text)

    If optRecepcion.Value Then
        strSQL = "SELECT dbo.fxSIFValidaTagRev(" & Trim(txtCedula) & ",'" & Trim(mTagRecepcion) _
                & "','" & Trim(mTagDevolucion) & "','AFI','" & vBoleta & "',NULL)"
        Call OpenRecordSet(rs, strSQL)
         If rs.Fields(0) = 2 Then
            MsgBox "No es posible registrar en forma consecutiva dos recepciones en la cedula " & txtCedula.Text
            txtCedula.Text = Empty
            rs.Close
            Exit Sub
        End If
    Else
        strSQL = "SELECT dbo.fxSIFValidaTagRev(" & Trim(txtCedula) & ",'" & Trim(mTagRecepcion) _
                & "','" & Trim(mTagDevolucion) & "','AFI','" & vBoleta & "',NULL)"
        Call OpenRecordSet(rs, strSQL)
         If rs.Fields(0) = 3 Then
            MsgBox "No es posible registrar en forma consecutiva dos devoluciones en la cedula " & txtCedula.Text
            txtCedula.Text = Empty
            rs.Close
            Exit Sub
        
        End If
    End If
   rs.Close
    strSQL = "SELECT dbo.fxSIFValidaTagRev(" & Trim(txtCedula) & ",'" & Trim(mTagRecepcion) & "','" & Trim(mTagDevolucion) _
            & "','AFI','" & vBoleta & "','" & Trim(mTagRecepcionDev) & "')"
    Call OpenRecordSet(rs, strSQL)
    If rs.Fields(0) = 4 Then
       MsgBox "No es posible registrar una recepción sin aplicar la devolución en la cedula " & txtCedula.Text
       txtCedula.Text = Empty
       rs.Close
        Exit Sub
    End If
    rs.Close
    
strSQL = "SELECT I.CEDULA,S.nombre,I.CONSEC ,isnull(O.DESCRIPCION,'') as DESCRIPCION, S.EstadoActual" _
       & " FROM AFI_INGRESOS I  inner join SOCIOS S on I.CEDULA = S.CEDULA" _
       & " LEFT JOIN SIF_OFICINAS O ON I.COD_OFICINA = O.COD_OFICINA" _
       & " WHERE I.CEDULA = '" & txtCedula & "' and I.consec = " & vBoleta


    Call OpenRecordSet(rs, strSQL)
    
    If Not rs.EOF Then
        If rs!EstadoActual = "S" Then
            Set itmX = lswCedula.ListItems.Add(, , rs!Cedula)
               itmX.SubItems(1) = rs!Nombre
               itmX.SubItems(2) = rs!Descripcion
               itmX.SubItems(3) = rs!Consec
         End If
    End If
    rs.Close

    txtCedula.Text = Empty
    txtCedula.SetFocus
    cboBoleta.Clear

    Exit Sub
    
vError:
        MsgBox fxSys_Error_Handler(Err.Description)

End Sub

Private Sub sbAplicarRecepcionDevolucion()
Dim i As Integer, strSQL As String

On Error GoTo vError

If MsgBox("Está seguro que sea aplicar estas etiquetas", vbExclamation + vbYesNo) = vbNo Then
    Exit Sub
End If

If optRecepcion.Value = True Then
    If mTagRecepcion = Empty Then
        MsgBox "No se puede realizar el proceso no está definido la etiqueta de recepción"
        Exit Sub
    End If
Else
    If mTagDevolucion = Empty Then
        MsgBox "No se puede realizar el proceso no está definido la etiqueta de devolución"
        Exit Sub
    End If
End If

Me.MousePointer = vbHourglass

PrgBar.Max = lswCedula.ListItems.Count + 1
PrgBar.Value = 1
PrgBar.Visible = True


With lswCedula.ListItems

For i = 1 To .Count

    If optRecepcion.Value = True Then
    
        Call sbSIFRegistraTags(.Item(i).Text, mTagRecepcion, "Recibida la documentación de la afiliación", .Item(i).SubItems(3), "AFI", .Item(i).Text, .Item(i).SubItems(3))
       
    Else
        Call sbSIFRegistraTags(.Item(i).Text, mTagDevolucion, "Devolución la documentación de la afiliación", .Item(i).SubItems(3), "AFI", .Item(i).Text, .Item(i).SubItems(3))
    End If

    PrgBar.Value = PrgBar.Value + 1
Next i

.Clear

End With

PrgBar.Visible = False

Me.MousePointer = vbDefault


MsgBox "Proceso concluido con éxito...", vbInformation

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical


End Sub

Private Function fxValidaNoDuplicados() As Boolean
Dim i As Integer

    fxValidaNoDuplicados = False

    For i = 1 To lswCedula.ListItems.Count

        If Trim(lswCedula.ListItems(i).Text) = Trim(txtCedula.Text) And Trim(lswCedula.ListItems(i).SubItems(3)) = SIFGlobal.fxCodText(cboBoleta.Text) Then
            fxValidaNoDuplicados = True
        End If
        
    Next i

End Function

Private Sub cboUsuario_Change()

End Sub

Private Sub cmdAgregar_Click()
    Call sbCargaInformacion
End Sub

Private Sub sbLimpiarDatos(ByVal Todo As Boolean)

    If Todo = True Then
        txtCedula.Text = Empty
    End If
    
End Sub

Private Sub dtpCorte_CallbackKeyDown(ByVal KeyCode As Integer, ByVal Shift As Integer, ByVal CallbackField As String, CallbackDate As Date)

End Sub


Private Sub Form_Activate()
vModulo = 8
End Sub

Private Sub Form_Load()
 vModulo = 8
 Call Formularios(Me)
    
    SSTab.Tab = 0
    Call sbParametrosTags
    dtpInicio.Value = fxFechaServidor
    dtpCorte.Value = dtpInicio.Value
    vGrid.MaxRows = 0
    

 Call RefrescaTags(Me)
End Sub

Private Sub lswCedula_DblClick()
    If lswCedula.ListItems.Count > 0 Then
        If lswCedula.SelectedItem.Index > 0 Then
            If MsgBox("Desea eliminar el cédula " & lswCedula.SelectedItem, vbYesNo) = vbYes Then
                lswCedula.ListItems.Remove (lswCedula.SelectedItem.Index)
            End If
        End If
    End If
End Sub



Private Sub optDevolucion_Click()
    lswCedula.ListItems.Clear
End Sub



Private Sub optRecepcion_Click()
    lswCedula.ListItems.Clear
End Sub

Private Sub ssTab_Click(PreviousTab As Integer)
    Select Case SSTab.Tab
    Case 0
        optRecepcion.Value = True
    Case 1
        optPendRecepcion.Value = True
    Case 2
        vGridConsulta.MaxRows = 0
        vGridConsulta.MaxCols = 4
        txtBuscaCedula.Text = Empty
        txtBuscaCedula.SetFocus
        
         
        
'        Call sbCargarUsuarios
    End Select
End Sub

Private Sub tlbAplicar_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case UCase(Button.Key)
    Case "APLICAR"
        Call sbAplicarRecepcionDevolucion
    End Select
End Sub

Private Sub tlbAplicar_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    Select Case UCase(ButtonMenu.Key)
    Case "LIMPIAR"
        lswCedula.ListItems.Clear
    Case "ELIMINAR"
        If lswCedula.ListItems.Count > 0 Then
            If lswCedula.SelectedItem.Index <> 0 Then
                lswCedula.ListItems.Remove (lswCedula.SelectedItem.Index)
            End If
        End If
    End Select
End Sub

Private Sub sbCargarListaSolicitudes()

Dim strSQL As String, BancosSeleccionados As String, Estado As String
    
On Error GoTo error
    
    
    Me.MousePointer = vbHourglass
    vGrid.SetFocus
    dtpInicio.Refresh
    dtpCorte.Refresh
    
    If optPendRecepcion.Value = True Then
       strSQL = "SELECT L.CONSEC, L.CEDULA,S.nombre,isnull(O.DESCRIPCION,'') as DESCRIPCION,L.Fecha,L.USUARIO" _
                & " FROM Afi_ingresos L  inner join SOCIOS S on L.CEDULA = S.CEDULA " _
                & " LEFT JOIN SIF_OFICINAS O ON L.COD_OFICINA = O.COD_OFICINA" _
                & " Where isnull(L.ANALISTA_RECEPCION,0) = 0 and S.EstadoActual in('S','A','P')"
    Else
        strSQL = "SELECT L.CONSEC, L.CEDULA,S.nombre,isnull(O.DESCRIPCION,'') as DESCRIPCION,L.Fecha,L.USUARIO" _
                & " FROM Afi_ingresos L  inner join SOCIOS S on L.CEDULA = S.CEDULA " _
                & " LEFT JOIN SIF_OFICINAS O ON L.COD_OFICINA = O.COD_OFICINA" _
                & " Where isnull(L.ANALISTA_RECEPCION,0) = 2 and S.EstadoActual in('S','A','P')"
    End If
   
                      
    Call sbCargaGrid(vGrid, 6, strSQL)
    vGrid.MaxRows = vGrid.MaxRows - 1
    
    Me.MousePointer = vbDefault
    Exit Sub
error:
    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical
  
    
End Sub


Private Sub tlbBuscar_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case UCase(Button.Key)
    Case "BUSCAR"
        Call sbCargarListaSolicitudes
    Case "IMPRIMIR"
       
End Select
End Sub





Private Sub tlbReportes_ButtonClick(ByVal Button As MSComctlLib.Button)

End Sub



Private Sub txtBuscaBoleta_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    Call sbCargarGridConsulta
End If
End Sub

Private Sub txtBuscaCedula_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    Call sbCargarGridConsulta
End If
End Sub


Private Sub txtBuscaUsuario_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    Call sbCargarGridConsulta
End If

End Sub

Private Sub txtCedula_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call sbCargarBoleta
    End If
End Sub

Private Sub sbCargarBoleta()
Dim strSQL As String, rs As New ADODB.Recordset
Dim vValor As Integer

On Error GoTo vError


Me.MousePointer = vbHourglass

Select Case True
  Case optRecepcion.Value
    vValor = 0
  Case optDevolucion.Value
    vValor = 1
End Select

strSQL = "SELECT Consec, Usuario, Fecha_Ingreso" _
       & " FROM AFI_INGRESOS " _
       & " WHERE CEDULA = '" & txtCedula & "' and isnull(Analista_Recepcion,0) = " & vValor

cboBoleta.Clear

Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
  cboBoleta.AddItem rs!Consec & " - " & rs!Usuario & " (" & Format(rs!Fecha_Ingreso, "dd/mm/yyyy") & ")"
  cboBoleta.Text = rs!Consec & " - " & rs!Usuario & " (" & Format(rs!Fecha_Ingreso, "dd/mm/yyyy") & ")"
  
  rs.MoveNext
Loop
rs.Close

Me.MousePointer = vbDefault


If cboBoleta.ListCount = 0 Then
   MsgBox "No se encuentran registros pendientes...", vbExclamation
End If

If cboBoleta.ListCount = 1 Then
 Call sbCargaInformacion
End If



Exit Sub

vError:
    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub sbCargarGridConsulta()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

vGridConsulta.MaxCols = 4
vGridConsulta.MaxRows = 0
    

Me.MousePointer = vbHourglass

strSQL = "select T.DESCRIPCION, CT.NOTAS, CT.REGISTRO_FECHA, CT.REGISTRO_USUARIO, CT.Documento" _
       & " from SIF_CONTROL_TAGS CT inner join SIF_TAGS T on CT.TAG_CODIGO = T.TAG_CODIGO" _
       & " where CT.cod_Modulo = 'AFI' and CT.codigo like '%" & txtBuscaCedula.Text & "%'"

If txtBuscaBoleta.Text <> "" Then
  strSQL = strSQL & " and CT.documento = '%" & txtBuscaBoleta.Text & "%'"
End If
 




Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
    vGridConsulta.MaxRows = vGridConsulta.MaxRows + 1
    vGridConsulta.Row = vGridConsulta.MaxRows
  
    vGridConsulta.Col = 1
    vGridConsulta.Text = rs!Descripcion
    
    vGridConsulta.Col = 2
    vGridConsulta.Value = IIf(IsNull(rs!REGISTRO_FECHA), "", rs!REGISTRO_FECHA)
    
    vGridConsulta.Col = 3
    vGridConsulta.Value = IIf(IsNull(rs!REGISTRO_USUARIO), "", rs!REGISTRO_USUARIO)
    
    vGridConsulta.Col = 4
    vGridConsulta.Value = IIf(IsNull(rs!Documento), "", rs!Documento)
    
    vGridConsulta.RowHeight(vGridConsulta.Row) = vGridConsulta.MaxTextRowHeight(vGridConsulta.Row)
    rs.MoveNext
Loop
rs.Close

Me.MousePointer = vbDefault

Exit Sub

vError:
    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical
    
End Sub


'Private Sub sbCargarUsuarios()
'Dim strSQL As String
'
'On Error GoTo vError
'    Me.MousePointer = vbHourglass
'
'    strSQL = "SELECT UPPER(NOMBRE) as ItmX from USUARIOS WHERE ESTADO = 'A'"
'
'    Call sbLlenaCbo(cboUsuario, strSQL, True)
'
'    Me.MousePointer = vbDefault
'    Exit Sub
'vError:
'    Me.MousePointer = vbDefault
'    MsgBox fxSys_Error_Handler(Err.Description), vbCritical
'End Sub



Private Function fxIngresoConsec(vCedula)
Dim strSQL As String
End Function


