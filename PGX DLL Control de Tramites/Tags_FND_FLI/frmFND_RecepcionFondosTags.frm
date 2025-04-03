VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.Ocx"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.Controls.v24.0.0.ocx"
Begin VB.Form frmFNDRecepcionFondosTags 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Fondos.: Recepción / Devolución"
   ClientHeight    =   8130
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   12015
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8130
   ScaleWidth      =   12015
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
      Tab             =   2
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
      TabCaption(0)   =   "Recepción - Devolución"
      TabPicture(0)   =   "frmFND_RecepcionFondosTags.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "txtContratoB"
      Tab(0).Control(1)=   "cmdAgregar"
      Tab(0).Control(2)=   "txtCodigo"
      Tab(0).Control(3)=   "optDevolucion"
      Tab(0).Control(4)=   "optRecepcion"
      Tab(0).Control(5)=   "txtCedula"
      Tab(0).Control(6)=   "lswCedula"
      Tab(0).Control(7)=   "tlbAplicar"
      Tab(0).Control(8)=   "PrgBar"
      Tab(0).Control(9)=   "Label10"
      Tab(0).Control(10)=   "Label4"
      Tab(0).Control(11)=   "Label3"
      Tab(0).Control(12)=   "LblCedula"
      Tab(0).Control(13)=   "Image1"
      Tab(0).ControlCount=   14
      TabCaption(1)   =   "Pendientes"
      TabPicture(1)   =   "frmFND_RecepcionFondosTags.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "optPendDevolucion"
      Tab(1).Control(1)=   "optPendRecepcion"
      Tab(1).Control(2)=   "vGrid"
      Tab(1).Control(3)=   "tlbBuscar"
      Tab(1).Control(4)=   "Image2"
      Tab(1).Control(5)=   "Label2(1)"
      Tab(1).ControlCount=   6
      TabCaption(2)   =   "Consultas"
      TabPicture(2)   =   "frmFND_RecepcionFondosTags.frx":0038
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "Label1(14)"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Label6"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Image4"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "Image3"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "Label8"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "Label9"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "dtpFFin"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).Control(7)=   "dtpFInicio"
      Tab(2).Control(7).Enabled=   0   'False
      Tab(2).Control(8)=   "vGridConsulta"
      Tab(2).Control(8).Enabled=   0   'False
      Tab(2).Control(9)=   "tlbReportes"
      Tab(2).Control(9).Enabled=   0   'False
      Tab(2).Control(10)=   "cboUsuario"
      Tab(2).Control(10).Enabled=   0   'False
      Tab(2).Control(11)=   "txtContrato"
      Tab(2).Control(11).Enabled=   0   'False
      Tab(2).Control(12)=   "txtPlan"
      Tab(2).Control(12).Enabled=   0   'False
      Tab(2).ControlCount=   13
      Begin VB.TextBox txtContratoB 
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
         Height          =   315
         Left            =   -70560
         MaxLength       =   10
         TabIndex        =   26
         Top             =   840
         Width           =   1575
      End
      Begin VB.TextBox txtPlan 
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
         Height          =   315
         Left            =   960
         MaxLength       =   10
         TabIndex        =   24
         Top             =   2040
         Width           =   1575
      End
      Begin VB.TextBox txtContrato 
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
         Height          =   315
         Left            =   2640
         MaxLength       =   10
         TabIndex        =   22
         Top             =   2040
         Width           =   1575
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
         Height          =   315
         Left            =   -68760
         TabIndex        =   21
         Top             =   840
         Width           =   495
      End
      Begin VB.TextBox txtCodigo 
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
         Height          =   315
         Left            =   -72240
         MaxLength       =   10
         TabIndex        =   19
         Top             =   840
         Width           =   1575
      End
      Begin VB.ComboBox cboUsuario 
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
         ForeColor       =   &H00FF0000&
         Height          =   330
         Left            =   4680
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   960
         Width           =   2895
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
         TabIndex        =   10
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
         TabIndex        =   9
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
         Height          =   315
         Left            =   -66480
         TabIndex        =   3
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
         Height          =   315
         Left            =   -67800
         TabIndex        =   2
         Top             =   840
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.TextBox txtCedula 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   -74160
         TabIndex        =   1
         Top             =   840
         Width           =   1815
      End
      Begin MSComctlLib.ListView lswCedula 
         Height          =   5895
         Left            =   -74880
         TabIndex        =   4
         Top             =   1320
         Width           =   11535
         _ExtentX        =   20346
         _ExtentY        =   10398
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
         NumItems        =   7
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Cédula"
            Object.Width           =   2540
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
            Text            =   "Estado"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Operadora"
            Object.Width           =   2999
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Plan"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Contrato"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSComctlLib.Toolbar tlbAplicar 
         Height          =   570
         Left            =   -66480
         TabIndex        =   5
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
         Left            =   -74880
         TabIndex        =   6
         Top             =   7320
         Visible         =   0   'False
         Width           =   8295
         _ExtentX        =   14631
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
      End
      Begin FPSpreadADO.fpSpread vGrid 
         Height          =   6495
         Left            =   -74880
         TabIndex        =   11
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
         MaxCols         =   486
         ScrollBarExtMode=   -1  'True
         SpreadDesigner  =   "frmFND_RecepcionFondosTags.frx":0054
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin MSComctlLib.Toolbar tlbBuscar 
         Height          =   330
         Left            =   -66240
         TabIndex        =   12
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
      Begin MSComctlLib.Toolbar tlbReportes 
         Height          =   330
         Left            =   9360
         TabIndex        =   15
         Top             =   840
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   582
         ButtonWidth     =   1931
         ButtonHeight    =   582
         Style           =   1
         TextAlignment   =   1
         ImageList       =   "ImageList3"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   1
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Imprimir"
               Key             =   "Imprimir"
               ImageIndex      =   6
               Style           =   5
               BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                  NumButtonMenus  =   2
                  BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "Recepcion"
                     Text            =   "Recepción"
                  EndProperty
                  BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "Devolucion"
                     Text            =   "Devolución"
                  EndProperty
               EndProperty
            EndProperty
         EndProperty
      End
      Begin FPSpreadADO.fpSpread vGridConsulta 
         Height          =   4815
         Left            =   360
         TabIndex        =   16
         Top             =   2640
         Width           =   10935
         _Version        =   524288
         _ExtentX        =   19288
         _ExtentY        =   8493
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
         SpreadDesigner  =   "frmFND_RecepcionFondosTags.frx":07D2
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin XtremeSuiteControls.DateTimePicker dtpFInicio 
         Height          =   330
         Left            =   960
         TabIndex        =   28
         Top             =   960
         Width           =   1455
         _Version        =   1572864
         _ExtentX        =   2566
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
      Begin XtremeSuiteControls.DateTimePicker dtpFFin 
         Height          =   330
         Left            =   2400
         TabIndex        =   29
         Top             =   960
         Width           =   1455
         _Version        =   1572864
         _ExtentX        =   2566
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
      Begin VB.Label Label10 
         Caption         =   "Contrato"
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
         Left            =   -70560
         TabIndex        =   27
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label9 
         Caption         =   "Plan"
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
         Left            =   960
         TabIndex        =   25
         Top             =   1800
         Width           =   975
      End
      Begin VB.Label Label8 
         Caption         =   "Contrato"
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
         Left            =   2640
         TabIndex        =   23
         Top             =   1800
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "Plan"
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
         Left            =   -72240
         TabIndex        =   20
         Top             =   600
         Width           =   975
      End
      Begin VB.Image Image3 
         Height          =   480
         Left            =   240
         Picture         =   "frmFND_RecepcionFondosTags.frx":0D4F
         Top             =   1440
         Width           =   480
      End
      Begin VB.Image Image4 
         Height          =   480
         Left            =   240
         Picture         =   "frmFND_RecepcionFondosTags.frx":0F68
         Top             =   600
         Width           =   480
      End
      Begin VB.Label Label6 
         Caption         =   "Reportes / Fechas"
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
         Left            =   960
         TabIndex        =   18
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Usuario"
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
         Index           =   14
         Left            =   4680
         TabIndex        =   17
         Top             =   720
         Width           =   735
      End
      Begin VB.Image Image2 
         Height          =   480
         Left            =   -74760
         Picture         =   "frmFND_RecepcionFondosTags.frx":1172
         Top             =   600
         Width           =   480
      End
      Begin VB.Label Label2 
         Caption         =   "Liquidaciones  Pendientes de:"
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
         Left            =   -74040
         TabIndex        =   13
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
         Left            =   -67200
         TabIndex        =   8
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
         Left            =   -74160
         TabIndex        =   7
         Top             =   600
         Width           =   975
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   -74760
         Picture         =   "frmFND_RecepcionFondosTags.frx":1383
         Top             =   720
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
            Picture         =   "frmFND_RecepcionFondosTags.frx":1575
            Key             =   "IMG1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFND_RecepcionFondosTags.frx":7DD7
            Key             =   "IMG2"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFND_RecepcionFondosTags.frx":E639
            Key             =   "IMG3"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList3 
      Left            =   0
      Top             =   600
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
            Picture         =   "frmFND_RecepcionFondosTags.frx":14E9B
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFND_RecepcionFondosTags.frx":1B6FD
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFND_RecepcionFondosTags.frx":21F5F
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFND_RecepcionFondosTags.frx":22079
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFND_RecepcionFondosTags.frx":22197
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFND_RecepcionFondosTags.frx":289F9
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmFNDRecepcionFondosTags"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim itmX As ListItem
Dim mTagRecepcion As String, mTagDevolucion As String
Dim mTagRecepcionDev As String
Dim strMensaje As String

Private Sub sbParametrosTags()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError
    
    '' Busca el parámetro del tag de recepción
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

On Error GoTo vError
    

If optDevolucion.Value Then
 strSQL = "SELECT F.CEDULA,S.nombre,F.COD_PLAN,F.COD_OPERADORA,F.COD_CONTRATO" _
        & ",isnull(O.DESCRIPCION,'') as DESCRIPCION,F.estado" _
        & " FROM FND_CONTRATOS F  inner join SOCIOS S on F.CEDULA = S.CEDULA" _
        & " LEFT JOIN SIF_OFICINAS O ON F.COD_OFICINA = O.COD_OFICINA" _
        & " WHERE F.cod_plan = '" & txtCodigo.Text & "' and cod_Contrato = " & txtContratoB.Text & " and F.analista_recepcion = 1"
End If

If optRecepcion.Value Then
 strSQL = "SELECT F.CEDULA,S.nombre,F.COD_PLAN,F.COD_OPERADORA,F.COD_CONTRATO" _
        & ",isnull(O.DESCRIPCION,'') as DESCRIPCION,F.estado" _
        & " FROM FND_CONTRATOS F  inner join SOCIOS S on F.CEDULA = S.CEDULA" _
        & " LEFT JOIN SIF_OFICINAS O ON F.COD_OFICINA = O.COD_OFICINA" _
        & " WHERE F.cod_plan = '" & txtCodigo.Text & "' and cod_Contrato = " & txtContratoB.Text & " and isnull(F.analista_recepcion,0) = 0"
End If

Call OpenRecordSet(rs, strSQL)

Do While Not rs.EOF
    Set itmX = lswCedula.ListItems.Add(, , rs!Cedula)
     itmX.SubItems(1) = rs!Nombre
     
     Select Case rs!Estado
        Case "A"
         itmX.SubItems(2) = "Activa"
        Case "L"
         itmX.SubItems(2) = "Liquidada"
     End Select
                           
     itmX.SubItems(3) = rs!Descripcion
     itmX.SubItems(4) = rs!Cod_Operadora
     itmX.SubItems(5) = rs!cod_Plan
     itmX.SubItems(6) = rs!cod_contrato
 rs.MoveNext
Loop
    
    rs.Close

    txtCedula.Text = Empty
    txtCedula.SetFocus

    Exit Sub
    
vError:
        MsgBox fxSys_Error_Handler(Err.Description), vbCritical

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
    If .Item(i).Checked Then
        If optRecepcion.Value = True Then
        
            Call sbSIFRegistraTags(.Item(i).SubItems(5), mTagRecepcion, "Recibida la documentación del contrato ", .Item(i).SubItems(6), "FND" _
                        , .Item(i).SubItems(5), .Item(i).SubItems(6))
            
        Else
            Call sbSIFRegistraTags(.Item(i).SubItems(5), mTagDevolucion, "Devolución la documentación del contrato", .Item(i).SubItems(6), "FND" _
                        , .Item(i).SubItems(5), .Item(i).SubItems(6))
        End If
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

        If Trim(lswCedula.ListItems(i).Text) = Trim(txtCedula.Text) And lswCedula.ListItems(i).ListSubItems(6).Text = Trim(txtCodigo.Text) _
        And Trim(lswCedula.ListItems(i).ListSubItems(7).Text) = Trim(txtContratoB.Text) Then
            fxValidaNoDuplicados = True
        End If
        
    Next i

End Function

Private Sub cmdAgregar_Click()
    Call sbCargaInformacion
End Sub

Private Sub sbLimpiarDatos(ByVal Todo As Boolean)

    If Todo = True Then
        txtCedula.Text = Empty
    End If
    
End Sub

Private Sub Form_Activate()
vModulo = 8
End Sub

Private Sub Form_Load()
 vModulo = 8
 Call Formularios(Me)
    
    SSTab.Tab = 0
    Call sbParametrosTags
    dtpFInicio.Value = fxFechaServidor
    dtpFFin.Value = dtpFInicio.Value
    vGrid.MaxRows = 0
    

 Call RefrescaTags(Me)
End Sub

Private Sub lswCedula_DblClick()
    If lswCedula.ListItems.Count > 0 Then
        If lswCedula.SelectedItem.Index > 0 Then
            If MsgBox("Desea eliminar la cédula " & lswCedula.SelectedItem, vbYesNo) = vbYes Then
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
        vGridConsulta.MaxCols = 3

        txtPlan.SetFocus
        
        Call sbCargarUsuarios
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

Private Sub sbCargarListaSolicitudes(Optional ByVal Num_Operacion As String = Empty)
' Carga Lista de operaciones
    Dim strSQL As String, BancosSeleccionados As String
    
On Error GoTo error
    'Consulta la lista de las Operaciones
    
    Me.MousePointer = vbHourglass
    vGrid.SetFocus
    dtpFInicio.Refresh
    dtpFFin.Refresh
    
 If optPendRecepcion.Value = True Then

    strSQL = "SELECT F.CEDULA,S.nombre,isnull(O.DESCRIPCION,'') as DESCRIPCION,F.FECHA_INICIO,F.USUARIO" _
            & ",F.COD_OPERADORA,F.COD_PLAN,F.COD_CONTRATO" _
            & " FROM FND_CONTRATOS F  inner join SOCIOS S on F.CEDULA = S.CEDULA " _
            & " LEFT JOIN SIF_OFICINAS O ON F.COD_OFICINA = O.COD_OFICINA" _
            & " Where isnull(F.ANALISTA_RECEPCION,0) = 0"
            
Else
    strSQL = "SELECT F.CEDULA,S.nombre,isnull(O.DESCRIPCION,'') as DESCRIPCION,F.FECHA_INICIO,F.USUARIO" _
            & ",F.COD_OPERADORA,F.COD_PLAN,F.COD_CONTRATO" _
            & " FROM FND_CONTRATOS F  inner join SOCIOS S on F.CEDULA = S.CEDULA " _
            & " LEFT JOIN SIF_OFICINAS O ON F.COD_OFICINA = O.COD_OFICINA" _
            & " Where isnull(F.ANALISTA_RECEPCION,0) = 2"
End If
            
            
                      
    Call sbCargaGrid(vGrid, 8, strSQL)
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





Private Sub txtCedula_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call sbCargaInformacion
    End If
End Sub



Private Sub sbCargarGridConsulta()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError


Me.MousePointer = vbHourglass

vGridConsulta.MaxCols = 3
vGridConsulta.MaxRows = 0

strSQL = "select T.DESCRIPCION, CT.NOTAS, CT.REGISTRO_FECHA, CT.REGISTRO_USUARIO" _
       & " from SIF_CONTROL_TAGS CT inner join SIF_TAGS T on CT.TAG_CODIGO = T.TAG_CODIGO" _
       & " where CT.cod_modulo = 'FND' and CT.codigo = '" & txtPlan.Text & "' and CT.documento = '" & txtContrato.Text & "'"
Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
    vGridConsulta.MaxRows = vGridConsulta.MaxRows + 1
    vGridConsulta.Row = vGridConsulta.MaxRows
  
    vGridConsulta.Col = 1
    vGridConsulta.Text = rs!Descripcion
    
    vGridConsulta.Col = 2
    vGridConsulta.Value = IIf(IsNull(rs!Registro_Fecha), "", rs!Registro_Fecha)
    
    vGridConsulta.Col = 3
    vGridConsulta.Value = IIf(IsNull(rs!registro_usuario), "", rs!registro_usuario)
    
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


Private Sub sbCargarUsuarios()
Dim strSQL As String

On Error GoTo vError
    Me.MousePointer = vbHourglass

    strSQL = "SELECT UPPER(NOMBRE) as ItmX from USUARIOS WHERE ESTADO = 'A'"
    
    Call sbLlenaCbo(cboUsuario, strSQL, True)

    Me.MousePointer = vbDefault
    Exit Sub
vError:
    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub txtCedulaBuscar_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call sbCargarGridConsulta
    End If
End Sub


Private Sub txtCodigo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 And txtCodigo.Locked = False Then
   gBusquedas.Convertir = "N"
   gBusquedas.Columna = "cod_plan"
   gBusquedas.Orden = "cod_plan"
   gBusquedas.Filtro = " And Cod_operadora=1"
   gBusquedas.Consulta = "select cod_plan,descripcion from fnd_planes"
   frmBusquedas.Show vbModal
   
   If Trim(gBusquedas.Resultado) <> "" Then
      txtCodigo = Trim(gBusquedas.Resultado)

   End If
   gBusquedas.Resultado = ""
   gBusquedas.Resultado2 = ""
End If
If KeyCode = vbKeyReturn Then
        Call sbCargaInformacion
End If
    

End Sub



Private Sub txtContrato_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call sbCargarGridConsulta
    End If
End Sub

Private Sub txtContratoB_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 And txtCodigo.Locked = False Then
   
   gBusquedas.Convertir = "N"
   gBusquedas.Columna = "cod_contrato"
   gBusquedas.Orden = "cod_contrato"
   gBusquedas.Filtro = " And cod_plan = '" & txtCodigo.Text & "' and cedula = '" & txtCedula.Text _
                     & "' and Estado ='A'"
   gBusquedas.Consulta = "select cod_contrato,cod_operadora from fnd_contratos"
   frmBusquedas.Show vbModal
   
   If Trim(gBusquedas.Resultado) <> "" Then
      txtContratoB = Trim(gBusquedas.Resultado)
   End If
   gBusquedas.Resultado = ""
   gBusquedas.Resultado2 = ""
End If



If KeyCode = vbKeyReturn Then
   Call sbCargaInformacion
End If
End Sub

Private Function fxValidaExisteFondo(vCedula As String, vContrato As Long) As Boolean
Dim strSQL As String, rs As New ADODB.Recordset


fxValidaExisteFondo = False
'Valida no agregar en forma mismo tag en forma consecutiva
If optRecepcion.Value Then
        strSQL = "SELECT dbo.fxSIFValidaTagRev('" & Trim(txtCedula) & "','" & Trim(mTagRecepcion) & "','" & Trim(mTagDevolucion) & "','FND','" & vContrato & "',NULL)"
        Call OpenRecordSet(rs, strSQL)
         If rs.Fields(0) = 2 Then
            fxValidaExisteFondo = True
            rs.Close
            strMensaje = "No es posible registrar en forma consecutiva dos recepciones.."
            Exit Function
        End If
    Else
        strSQL = "SELECT dbo.fxSIFValidaTagRev(" & Trim(txtCedula) & ",'" & Trim(mTagRecepcion) & "','" & Trim(mTagDevolucion) & "','FND','" & vContrato & "',NULL)"
        Call OpenRecordSet(rs, strSQL)
         If rs.Fields(0) = 3 Then
             fxValidaExisteFondo = True
            rs.Close
            strMensaje = "No es posible registrar en forma consecutiva dos devoluciones..."
            Exit Function
        
        End If
    End If
   
   rs.Close
   
   strSQL = "SELECT dbo.fxSIFValidaTagRev(" & Trim(txtCedula) & ",'" & Trim(mTagRecepcion) & "','" & Trim(mTagDevolucion) & "','FND','" & vContrato _
                                              & "','" & Trim(mTagRecepcionDev) & "')"
    Call OpenRecordSet(rs, strSQL)
    If rs.Fields(0) = 4 Then
       fxValidaExisteFondo = True
       rs.Close
       strMensaje = "No es posible registrar una recepción sin aplicar la devolución..."
        Exit Function
    End If
    rs.Close


End Function

Private Sub txtPlan_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call sbCargarGridConsulta
    End If
End Sub
