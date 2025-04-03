VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.Ocx"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.Controls.v24.0.0.ocx"
Begin VB.Form frmFNDRecepcionLiqFondosTags 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Recepción de Liquidaciones Fondos"
   ClientHeight    =   8175
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   12030
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8175
   ScaleWidth      =   12030
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
      TabPicture(0)   =   "frmFNDRecepcionLiqFondosTags.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "txtCodigo"
      Tab(0).Control(1)=   "cmdAgregar"
      Tab(0).Control(2)=   "optRecepcion"
      Tab(0).Control(3)=   "optDevolucion"
      Tab(0).Control(4)=   "lswLiq"
      Tab(0).Control(5)=   "tlbAplicar"
      Tab(0).Control(6)=   "PrgBar"
      Tab(0).Control(7)=   "Image1"
      Tab(0).Control(8)=   "LblCedula"
      Tab(0).Control(9)=   "Label3"
      Tab(0).ControlCount=   10
      TabCaption(1)   =   "Pendientes"
      TabPicture(1)   =   "frmFNDRecepcionLiqFondosTags.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "txtUsuario"
      Tab(1).Control(1)=   "cboCasos"
      Tab(1).Control(2)=   "optPendRecepcion"
      Tab(1).Control(3)=   "optPendDevolucion"
      Tab(1).Control(4)=   "tlbBuscar"
      Tab(1).Control(5)=   "vGrid"
      Tab(1).Control(6)=   "Label2(2)"
      Tab(1).Control(7)=   "Label2(1)"
      Tab(1).Control(8)=   "Label2(0)"
      Tab(1).Control(9)=   "Image2"
      Tab(1).ControlCount=   10
      TabCaption(2)   =   "Consultas"
      TabPicture(2)   =   "frmFNDRecepcionLiqFondosTags.frx":0038
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "Image4"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Label6"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Label1(14)"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "Image3"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "Label7"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "dtpFFin"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "dtpFInicio"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).Control(7)=   "vGridConsulta"
      Tab(2).Control(7).Enabled=   0   'False
      Tab(2).Control(8)=   "tlbReportes"
      Tab(2).Control(8).Enabled=   0   'False
      Tab(2).Control(9)=   "cboUsuario"
      Tab(2).Control(9).Enabled=   0   'False
      Tab(2).Control(10)=   "txtCodigoBuscar"
      Tab(2).Control(10).Enabled=   0   'False
      Tab(2).ControlCount=   11
      Begin VB.TextBox txtUsuario 
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
         Left            =   -68280
         TabIndex        =   25
         Top             =   720
         Width           =   1935
      End
      Begin VB.ComboBox cboCasos 
         BackColor       =   &H00C0FFFF&
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
         Left            =   -71160
         Style           =   2  'Dropdown List
         TabIndex        =   22
         Top             =   720
         Width           =   2895
      End
      Begin VB.TextBox txtCodigoBuscar 
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
         Left            =   840
         TabIndex        =   19
         Top             =   1920
         Width           =   2775
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
         Left            =   4560
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   960
         Width           =   2895
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
         Height          =   255
         Left            =   -72720
         TabIndex        =   11
         Top             =   480
         Width           =   1335
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
         Height          =   255
         Left            =   -72720
         TabIndex        =   10
         Top             =   840
         Width           =   1335
      End
      Begin VB.TextBox txtCodigo 
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
         Left            =   -74160
         TabIndex        =   6
         Top             =   840
         Width           =   2775
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
         Left            =   -71160
         TabIndex        =   3
         Top             =   840
         Width           =   495
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
         Left            =   -69720
         TabIndex        =   2
         Top             =   840
         Value           =   -1  'True
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
         Left            =   -68400
         TabIndex        =   1
         Top             =   840
         Width           =   1335
      End
      Begin MSComctlLib.ListView lswLiq 
         Height          =   5775
         Left            =   -74760
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
         Left            =   -66600
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
                     Text            =   "Eliminar"
                  EndProperty
               EndProperty
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ProgressBar PrgBar 
         Height          =   315
         Left            =   -74760
         TabIndex        =   9
         Top             =   7320
         Visible         =   0   'False
         Width           =   7935
         _ExtentX        =   13996
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
      End
      Begin MSComctlLib.Toolbar tlbBuscar 
         Height          =   330
         Left            =   -66120
         TabIndex        =   12
         Top             =   720
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
      Begin FPSpreadADO.fpSpread vGrid 
         Height          =   6495
         Left            =   -74880
         TabIndex        =   14
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
         MaxCols         =   8
         ScrollBarExtMode=   -1  'True
         SpreadDesigner  =   "frmFNDRecepcionLiqFondosTags.frx":0054
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin MSComctlLib.Toolbar tlbReportes 
         Height          =   330
         Left            =   9240
         TabIndex        =   16
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
         Left            =   240
         TabIndex        =   21
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
         SpreadDesigner  =   "frmFNDRecepcionLiqFondosTags.frx":07C7
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin XtremeSuiteControls.DateTimePicker dtpFInicio 
         Height          =   330
         Left            =   840
         TabIndex        =   26
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
         Left            =   2280
         TabIndex        =   27
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
      Begin VB.Label Label2 
         Caption         =   "Usuario:"
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
         Index           =   2
         Left            =   -68280
         TabIndex        =   24
         Top             =   480
         Width           =   1695
      End
      Begin VB.Label Label2 
         Caption         =   "Tipo de Casos:"
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
         Index           =   1
         Left            =   -71160
         TabIndex        =   23
         Top             =   480
         Width           =   1695
      End
      Begin VB.Label Label7 
         Caption         =   "No.  Boleta"
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
         TabIndex        =   20
         Top             =   1560
         Width           =   1935
      End
      Begin VB.Image Image3 
         Height          =   480
         Left            =   240
         Picture         =   "frmFNDRecepcionLiqFondosTags.frx":0D66
         Top             =   1440
         Width           =   480
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
         Left            =   4560
         TabIndex        =   18
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label6 
         Caption         =   "Reportes"
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
         TabIndex        =   17
         Top             =   720
         Width           =   975
      End
      Begin VB.Image Image4 
         Height          =   480
         Left            =   240
         Picture         =   "frmFNDRecepcionLiqFondosTags.frx":0F7F
         Top             =   480
         Width           =   480
      End
      Begin VB.Label Label2 
         Caption         =   "Liquidaciones Pendientes de:"
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
         Left            =   -74040
         TabIndex        =   13
         Top             =   480
         Width           =   1695
      End
      Begin VB.Image Image2 
         Height          =   480
         Left            =   -74760
         Picture         =   "frmFNDRecepcionLiqFondosTags.frx":1189
         Top             =   480
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   -74760
         Picture         =   "frmFNDRecepcionLiqFondosTags.frx":139A
         Top             =   600
         Width           =   480
      End
      Begin VB.Label LblCedula 
         AutoSize        =   -1  'True
         Caption         =   "No. Boleta"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   -74160
         TabIndex        =   5
         Top             =   600
         Width           =   735
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
         Left            =   -69720
         TabIndex        =   4
         Top             =   600
         Width           =   1695
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
            Picture         =   "frmFNDRecepcionLiqFondosTags.frx":158C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFNDRecepcionLiqFondosTags.frx":7DEE
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFNDRecepcionLiqFondosTags.frx":E650
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
            Picture         =   "frmFNDRecepcionLiqFondosTags.frx":14EB2
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFNDRecepcionLiqFondosTags.frx":1B714
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFNDRecepcionLiqFondosTags.frx":21F76
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFNDRecepcionLiqFondosTags.frx":22090
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFNDRecepcionLiqFondosTags.frx":221AE
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFNDRecepcionLiqFondosTags.frx":28A10
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmFNDRecepcionLiqFondosTags"
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

    If Not IsNumeric(txtCodigo) Then
        Exit Sub
    End If
    
    If fxValidaNoDuplicados Then
        MsgBox "La liquidación  ya fue digitada"
        txtCodigo.Text = Empty
        txtCodigo.SetFocus
        Exit Sub
    End If
    
    'Valida no agregar en forma mismo tag en forma consecutiva

    If optRecepcion.Value Then
        strSQL = "SELECT dbo.fxSIFValidaTagRev(" & Trim(txtCodigo.Text) & ",'" & Trim(mTagRecepcion) & "','" & Trim(mTagDevolucion) & "','FLQ',NULL,NULL)"
        Call OpenRecordSet(rs, strSQL)
         If rs.Fields(0) = 2 Then
            MsgBox "No es posible registrar en forma consecutiva dos recepciones en la Boleta .: " & txtCodigo.Text
            txtCodigo.Text = Empty
            rs.Close
            Exit Sub
        End If
    Else
        strSQL = "SELECT dbo.fxSIFValidaTagRev(" & Trim(txtCodigo) & ",'" & Trim(mTagRecepcion) & "','" & Trim(mTagDevolucion) & "','FLQ',NULL,NULL)"
        Call OpenRecordSet(rs, strSQL)
         If rs.Fields(0) = 3 Then
            MsgBox "No es posible registrar en forma consecutiva dos devoluciones en la Boleta .:" & txtCodigo.Text
            txtCodigo.Text = Empty
            rs.Close
            Exit Sub
        
        End If
    End If
   rs.Close
    strSQL = "SELECT dbo.fxSIFValidaTagRev(" & Trim(txtCodigo) & ",'" & Trim(mTagRecepcion) & "','" & Trim(mTagDevolucion) & "','FLQ',NULL,'" & Trim(mTagRecepcionDev) & "')"
    Call OpenRecordSet(rs, strSQL)
    If rs.Fields(0) = 4 Then
       MsgBox "No es posible registrar una recepción sin aplicar la devolución en la Boleta .:" & txtCodigo.Text
       txtCodigo.Text = Empty
       rs.Close
        Exit Sub
    End If
    rs.Close
    
    strSQL = "SELECT F.CEDULA,S.nombre,L.CONSEC ,isnull(O.DESCRIPCION,'') as DESCRIPCION" _
           & " FROM fnd_liquidacion L  inner join FND_CONTRATOS F" _
           & " on L.COD_PLAN = F.COD_PLAN and L.COD_CONTRATO = F.COD_CONTRATO" _
           & " AND L.COD_OPERADORA = F.COD_OPERADORA" _
           & " inner join SOCIOS S on F.CEDULA = S.CEDULA" _
           & " LEFT JOIN SIF_OFICINAS O ON L.COD_OFICINA = O.COD_OFICINA" _
           & " WHERE L.CONSEC = '" & txtCodigo & "'"


    Call OpenRecordSet(rs, strSQL)
    
    If Not rs.EOF Then
         Set itmX = lswLiq.ListItems.Add(, , rs!Cedula)
        itmX.SubItems(1) = rs!Nombre
        itmX.SubItems(2) = rs!Descripcion
        itmX.SubItems(3) = rs!Consec
    End If
    rs.Close

    txtCodigo.Text = Empty
    txtCodigo.SetFocus

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

PrgBar.Max = lswLiq.ListItems.Count + 1
PrgBar.Value = 1
PrgBar.Visible = True


With lswLiq.ListItems

For i = 1 To .Count

    If optRecepcion.Value = True Then
    
        Call sbSIFRegistraTags(.Item(i).Text, mTagRecepcion, "Recibida la documentación de la Liquidación", .Item(i).SubItems(3), "FLQ" _
                    , .Item(i).SubItems(3))
        
    Else
        Call sbSIFRegistraTags(.Item(i).Text, mTagDevolucion, "Devolución la documentación de la Liquidación", .Item(i).SubItems(3), "FLQ" _
                    , .Item(i).SubItems(3))
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

    For i = 1 To lswLiq.ListItems.Count

        If Trim(lswLiq.ListItems(i).SubItems(3)) = Trim(txtCodigo.Text) Then
            fxValidaNoDuplicados = True
        End If
        
    Next i

End Function

Private Sub cmdAgregar_Click()
    Call sbCargaInformacion
End Sub

Private Sub sbLimpiarDatos(ByVal Todo As Boolean)

    If Todo = True Then
        txtCodigo.Text = Empty
    End If
    
End Sub

Private Sub Form_Activate()
vModulo = 8
End Sub

Private Sub Form_Load()
 vModulo = 8
 Call Formularios(Me)
    
    SSTab.Tab = 0
    cboCasos.AddItem "Todos"
    cboCasos.AddItem "Desembolso (Con..)"
    cboCasos.AddItem "Retenidos"
    cboCasos.Text = "Todos"
    
    Call sbParametrosTags
    dtpFInicio.Value = fxFechaServidor
    dtpFFin.Value = dtpFInicio.Value
    vGrid.MaxRows = 0
    

 Call RefrescaTags(Me)
End Sub

Private Sub lswLiq_DblClick()
    If lswLiq.ListItems.Count > 0 Then
        If lswLiq.SelectedItem.Index > 0 Then
            If MsgBox("Desea eliminar el cédula " & lswLiq.SelectedItem, vbYesNo) = vbYes Then
                lswLiq.ListItems.Remove (lswLiq.SelectedItem.Index)
            End If
        End If
    End If
End Sub



Private Sub optDevolucion_Click()
    lswLiq.ListItems.Clear
End Sub

Private Sub optRecepcion_Click()
    lswLiq.ListItems.Clear
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
        txtCodigoBuscar.Text = Empty
        txtCodigoBuscar.SetFocus
        
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
        lswLiq.ListItems.Clear
    Case "ELIMINAR"
        If lswLiq.ListItems.Count > 0 Then
            If lswLiq.SelectedItem.Index <> 0 Then
                lswLiq.ListItems.Remove (lswLiq.SelectedItem.Index)
            End If
        End If
    End Select
End Sub

Private Sub sbCargarListaSolicitudes()

Dim strSQL As String, BancosSeleccionados As String, Estado As String
    
On Error GoTo vError
    
    
Me.MousePointer = vbHourglass

vGrid.SetFocus

dtpFInicio.Refresh
dtpFFin.Refresh

If optPendRecepcion.Value = True Then
    strSQL = "SELECT L.CONSEC,L.COD_PLAN,L.COD_CONTRATO, F.CEDULA,S.nombre,isnull(O.DESCRIPCION,'') as DESCRIPCION,L.Fecha,L.USUARIO" _
           & " FROM fnd_liquidacion L  inner join FND_CONTRATOS F on L.COD_PLAN = F.COD_PLAN and L.COD_CONTRATO = F.COD_CONTRATO" _
           & " AND L.COD_OPERADORA = F.COD_OPERADORA" _
           & " inner join SOCIOS S on F.CEDULA = S.CEDULA" _
           & " LEFT JOIN SIF_OFICINAS O ON L.COD_OFICINA = O.COD_OFICINA" _
           & " Where isnull(L.ANALISTA_RECEPCION,0) = 0"
   
Else
    strSQL = "SELECT L.CONSEC,L.COD_PLAN,L.COD_CONTRATO,F.CEDULA,S.nombre,isnull(O.DESCRIPCION,'') as DESCRIPCION,L.Fecha,L.USUARIO" _
           & " FROM fnd_liquidacion L  inner join FND_CONTRATOS F on L.COD_PLAN = F.COD_PLAN and L.COD_CONTRATO = F.COD_CONTRATO" _
           & " AND L.COD_OPERADORA = F.COD_OPERADORA" _
           & " inner join SOCIOS S on F.CEDULA = S.CEDULA" _
           & " LEFT JOIN SIF_OFICINAS O ON L.COD_OFICINA = O.COD_OFICINA" _
           & " Where isnull(L.ANALISTA_RECEPCION,0) = 2"
End If



Select Case Mid(cboCasos.Text, 1, 1)
  Case "T"
  Case "D"
     strSQL = strSQL & " and L.Retencion_Codigo is null"
  Case "R"
     strSQL = strSQL & " and L.Retencion_Codigo is not null"
End Select

If Len(Trim(txtUsuario.Text)) > 0 Then
   strSQL = strSQL & " and L.Usuario like '%" & Trim(txtUsuario.Text) & "%'"
End If


Call sbCargaGrid(vGrid, 8, strSQL)
vGrid.MaxRows = vGrid.MaxRows - 1

Me.MousePointer = vbDefault

Exit Sub

vError:
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

Private Sub txtCodigo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call sbCargaInformacion
    End If
End Sub

Private Sub sbCargarGridConsulta()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError
            
vGridConsulta.MaxCols = 3
vGridConsulta.MaxRows = 0

If txtCodigoBuscar.Text = Empty Then Exit Sub

Me.MousePointer = vbHourglass

strSQL = "select T.DESCRIPCION, CT.NOTAS, CT.REGISTRO_FECHA, CT.REGISTRO_USUARIO" _
      & "  from SIF_CONTROL_TAGS CT inner join SIF_TAGS T on CT.TAG_CODIGO = T.TAG_CODIGO" _
      & " where CT.documento = '" & txtCodigoBuscar & "' and CT.COD_MODULO = 'FLQ'"
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

Private Sub txtCodigoBuscar_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call sbCargarGridConsulta
    End If
End Sub
