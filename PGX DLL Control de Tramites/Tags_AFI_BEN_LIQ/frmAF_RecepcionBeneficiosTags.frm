VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.Ocx"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.Controls.v24.0.0.ocx"
Begin VB.Form frmAF_RecepcionBeneficiosTags 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Recepción de Beneficios "
   ClientHeight    =   8190
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   12015
   BeginProperty Font 
      Name            =   "Arial Narrow"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8190
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
      TabPicture(0)   =   "frmAF_RecepcionBeneficiosTags.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Image1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "LblBeneficio"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label3"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label4"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "PrgBar"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "tlbAplicar"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "lswBeneficio"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "cmdAgregar"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "optDevolucion"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "txtBeneficio"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "optRecepcion"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "txtCodigo"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).ControlCount=   12
      TabCaption(1)   =   "Pendientes"
      TabPicture(1)   =   "frmAF_RecepcionBeneficiosTags.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "optPendRecepcion"
      Tab(1).Control(1)=   "optPendDevolucion"
      Tab(1).Control(2)=   "vGrid"
      Tab(1).Control(3)=   "tlbBuscar"
      Tab(1).Control(4)=   "Label2"
      Tab(1).Control(5)=   "Image2"
      Tab(1).ControlCount=   6
      TabCaption(2)   =   "Consultas"
      TabPicture(2)   =   "frmAF_RecepcionBeneficiosTags.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "txtCodigoBuscar"
      Tab(2).Control(1)=   "txtBeneficioBuscar"
      Tab(2).Control(2)=   "cboUsuario"
      Tab(2).Control(3)=   "tlbReportes"
      Tab(2).Control(4)=   "vGridConsulta"
      Tab(2).Control(5)=   "dtpFInicio"
      Tab(2).Control(6)=   "dtpFFin"
      Tab(2).Control(7)=   "Label8"
      Tab(2).Control(8)=   "Label7"
      Tab(2).Control(9)=   "Image3"
      Tab(2).Control(10)=   "Label1(14)"
      Tab(2).Control(11)=   "Label6"
      Tab(2).Control(12)=   "Image4"
      Tab(2).ControlCount=   13
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
         Height          =   355
         Left            =   -72720
         TabIndex        =   24
         Top             =   1920
         Width           =   1455
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
         Left            =   2400
         TabIndex        =   22
         Top             =   840
         Width           =   1455
      End
      Begin VB.TextBox txtBeneficioBuscar 
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
         Height          =   355
         Left            =   -74160
         TabIndex        =   19
         Top             =   1920
         Width           =   1455
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
         Left            =   -71160
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
         Height          =   495
         Left            =   -71640
         TabIndex        =   11
         Top             =   600
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
         Height          =   495
         Left            =   -70200
         TabIndex        =   10
         Top             =   600
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
         Left            =   5040
         TabIndex        =   7
         Top             =   840
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.TextBox txtBeneficio 
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
         TabIndex        =   6
         Top             =   840
         Width           =   1455
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
         Left            =   6360
         TabIndex        =   2
         Top             =   840
         Width           =   1335
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
         Left            =   4080
         TabIndex        =   1
         Top             =   840
         Width           =   570
      End
      Begin MSComctlLib.ListView lswBeneficio 
         Height          =   5775
         Left            =   240
         TabIndex        =   3
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
         NumItems        =   5
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
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Beneficio"
            Object.Width           =   2540
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
               ImageKey        =   "IMG1"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Opciones"
               Key             =   "Opciones"
               Object.ToolTipText     =   "Opciones"
               ImageKey        =   "IMG2"
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
         TabIndex        =   12
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
         MaxCols         =   7
         ScrollBarExtMode=   -1  'True
         SpreadDesigner  =   "frmAF_RecepcionBeneficiosTags.frx":0054
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin MSComctlLib.Toolbar tlbBuscar 
         Height          =   330
         Left            =   -66240
         TabIndex        =   13
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
         Left            =   -65760
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
         Left            =   -74760
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
         SpreadDesigner  =   "frmAF_RecepcionBeneficiosTags.frx":0756
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin XtremeSuiteControls.DateTimePicker dtpFInicio 
         Height          =   330
         Left            =   -74160
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
         Left            =   -72720
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
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         Caption         =   "Código"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -72600
         TabIndex        =   25
         Top             =   1560
         Width           =   1455
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Caption         =   "Código"
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
         Left            =   2400
         TabIndex        =   23
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Caption         =   "Beneficio"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74160
         TabIndex        =   20
         Top             =   1560
         Width           =   1455
      End
      Begin VB.Image Image3 
         Height          =   480
         Left            =   -74760
         Picture         =   "frmAF_RecepcionBeneficiosTags.frx":0CD3
         Top             =   1440
         Width           =   480
      End
      Begin VB.Label Label1 
         Caption         =   "Usuario"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   14
         Left            =   -71160
         TabIndex        =   18
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label6 
         Caption         =   "Reportes"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74160
         TabIndex        =   17
         Top             =   720
         Width           =   975
      End
      Begin VB.Image Image4 
         Height          =   480
         Left            =   -74760
         Picture         =   "frmAF_RecepcionBeneficiosTags.frx":0EEC
         Top             =   480
         Width           =   480
      End
      Begin VB.Label Label2 
         Caption         =   "Beneficios Pendientes de:"
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
         TabIndex        =   14
         Top             =   720
         Width           =   2295
      End
      Begin VB.Image Image2 
         Height          =   480
         Left            =   -74760
         Picture         =   "frmAF_RecepcionBeneficiosTags.frx":10F6
         Top             =   480
         Width           =   480
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
         Left            =   5040
         TabIndex        =   5
         Top             =   600
         Width           =   1695
      End
      Begin VB.Label LblBeneficio 
         Alignment       =   2  'Center
         Caption         =   "Beneficio"
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
         TabIndex        =   4
         Top             =   600
         Width           =   1455
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   240
         Picture         =   "frmAF_RecepcionBeneficiosTags.frx":1307
         Top             =   480
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
            Picture         =   "frmAF_RecepcionBeneficiosTags.frx":14F9
            Key             =   "IMG1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAF_RecepcionBeneficiosTags.frx":7D5B
            Key             =   "IMG2"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAF_RecepcionBeneficiosTags.frx":E5BD
            Key             =   "IMG3"
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
            Picture         =   "frmAF_RecepcionBeneficiosTags.frx":14E1F
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAF_RecepcionBeneficiosTags.frx":1B681
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAF_RecepcionBeneficiosTags.frx":21EE3
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAF_RecepcionBeneficiosTags.frx":21FFD
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAF_RecepcionBeneficiosTags.frx":2211B
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAF_RecepcionBeneficiosTags.frx":2897D
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmAF_RecepcionBeneficiosTags"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim itmX As ListItem
Dim mTagRecepcion As String, mTagDevolucion As String
Dim mTagRecepcionDev As String
Dim mCodigo As String, mBeneficio As String

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
Dim vIdCompuesto As String

On Error GoTo vError
  

        If Trim(txtBeneficio.Text) = Empty Or Trim(txtCodigo.Text) = Empty Then
            Exit Sub
        End If
    
    mCodigo = Trim(txtCodigo.Text)
    mBeneficio = Trim(txtBeneficio)

    vIdCompuesto = mBeneficio & "." & mBeneficio
    
'    mCodigo = SIFGlobal.fxCodText(txtBeneficio.Text)
'    mBeneficio = Trim(Mid(txtBeneficio.Text, InStr(1, txtBeneficio.Text, "-") + 1, Len(txtBeneficio.Text)))
     
    If fxValidaNoDuplicados = True Then
        MsgBox "La Cedula se ya fue digitada"
        txtBeneficio.Text = Empty
        txtBeneficio.SetFocus
        Exit Sub
    End If
    
    'Valida no agregar en forma mismo tag en forma consecutiva

    If optRecepcion.Value Then
        strSQL = "SELECT dbo.fxSIFValidaTagRev('" & Trim(vIdCompuesto) & "','" & Trim(mTagRecepcion) & "','" & Trim(mTagDevolucion) & "','BEN',  '" & Trim(mBeneficio) & "',NULL)"
        Call OpenRecordSet(rs, strSQL)
        
         If rs.Fields(0) = 2 Then
            MsgBox "No es posible registrar en forma consecutiva dos recepciones del Beneficio ..: " & vIdCompuesto
            txtBeneficio.Text = Empty
            rs.Close
            Exit Sub
        End If
    Else
        strSQL = "SELECT dbo.fxSIFValidaTagRev('" & Trim(vIdCompuesto) & "','" & Trim(mTagRecepcion) & "','" & Trim(mTagDevolucion) & "','BEN','" & Trim(mBeneficio) & "',NULL)"
        Call OpenRecordSet(rs, strSQL)
        
         If rs.Fields(0) = 3 Then
            MsgBox "No es posible registrar en forma consecutiva dos devoluciones del Beneficio ..: " & vIdCompuesto
            txtBeneficio.Text = Empty
            rs.Close
            Exit Sub
        
        End If
    End If
   rs.Close
   
    strSQL = "SELECT dbo.fxSIFValidaTagRev('" & Trim(vIdCompuesto) & "','" & Trim(mTagRecepcion) & "','" & Trim(mTagDevolucion) & "','BEN','" & Trim(mBeneficio) _
           & "','" & Trim(mTagRecepcionDev) & "')"
           
    Call OpenRecordSet(rs, strSQL)
    
    If rs.Fields(0) = 4 Then
       MsgBox "No es posible registrar una recepción sin aplicar la devolución del Beneficio ..: " & vIdCompuesto
       txtBeneficio.Text = Empty
       rs.Close
        Exit Sub
    End If
    rs.Close
    
  strSQL = "SELECT B.CEDULA,S.nombre,B.CONSEC,B.COD_BENEFICIO ,isnull(O.DESCRIPCION,'') as DESCRIPCION" _
       & " FROM AFI_BENE_OTORGA B  inner join SOCIOS S on B.CEDULA = S.CEDULA" _
       & " LEFT JOIN SIF_OFICINAS O ON B.COD_OFICINA = O.COD_OFICINA" _
       & " WHERE B.CONSEC = '" & mCodigo & "' and B.COD_BENEFICIO = '" & mBeneficio & "'"
 


    Call OpenRecordSet(rs, strSQL)
    
    If Not rs.EOF Then
         Set itmX = lswBeneficio.ListItems.Add(, , rs!Cedula)
        itmX.SubItems(1) = rs!Nombre
        itmX.SubItems(2) = rs!Descripcion
        itmX.SubItems(3) = rs!Consec
        itmX.SubItems(4) = rs!cod_beneficio
        
    End If
    rs.Close

    txtBeneficio.Text = Empty
    txtBeneficio.SetFocus

    Exit Sub
    
vError:
        MsgBox fxSys_Error_Handler(Err.Description)

End Sub

Private Sub sbAplicarRecepcionDevolucion()
Dim i As Integer, strSQL As String
Dim vIdCompuesto As String

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

PrgBar.Max = lswBeneficio.ListItems.Count + 1
PrgBar.Value = 1
PrgBar.Visible = True


With lswBeneficio.ListItems

For i = 1 To .Count

   
    If optRecepcion.Value = True Then
        
        Call sbSIFRegistraTags(Trim(.Item(i).SubItems(4)), mTagRecepcion, "Recibida la documentación del Beneficio ", Trim(.Item(i).SubItems(3)), "BEN" _
                        , Trim(.Item(i).SubItems(4)), Trim(.Item(i).SubItems(3)), Trim(.Item(i).Text))
        
    Else
        
        Call sbSIFRegistraTags(Trim(.Item(i).SubItems(4)), mTagDevolucion, "Devolución de la documentación del Beneficio ", Trim(.Item(i).SubItems(3)), "BEN" _
                        , Trim(.Item(i).SubItems(4)), Trim(.Item(i).SubItems(3)), Trim(.Item(i).Text))
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

    For i = 1 To lswBeneficio.ListItems.Count

        If Trim(lswBeneficio.ListItems(i).Text) = Trim(txtBeneficio.Text) Then
            fxValidaNoDuplicados = True
        End If
        
    Next i

End Function

Private Sub cmdAgregar_Click()
    Call sbCargaInformacion
End Sub

Private Sub sbLimpiarDatos(ByVal Todo As Boolean)

    If Todo = True Then
        txtBeneficio.Text = Empty
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

Private Sub lswBeneficio_DblClick()
    If lswBeneficio.ListItems.Count > 0 Then
        If lswBeneficio.SelectedItem.Index > 0 Then
            If MsgBox("Desea eliminar el Beneficio ..: " & lswBeneficio.SelectedItem, vbYesNo) = vbYes Then
                lswBeneficio.ListItems.Remove (lswBeneficio.SelectedItem.Index)
            End If
        End If
    End If
End Sub

Private Sub optDevolucion_Click()
    lswBeneficio.ListItems.Clear
End Sub

Private Sub optRecepcion_Click()
    lswBeneficio.ListItems.Clear
End Sub

Private Sub ssTab_Click(PreviousTab As Integer)
    Select Case SSTab.Tab
    Case 0
        optRecepcion.Value = True
    Case 1
        optPendRecepcion.Value = True
        vGrid.MaxRows = 0
    Case 2
        vGridConsulta.MaxRows = 0
        vGridConsulta.MaxCols = 3
        txtBeneficioBuscar.Text = Empty
        txtBeneficioBuscar.SetFocus
        
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
        lswBeneficio.ListItems.Clear
    Case "ELIMINAR"
        If lswBeneficio.ListItems.Count > 0 Then
            If lswBeneficio.SelectedItem.Index <> 0 Then
                lswBeneficio.ListItems.Remove (lswBeneficio.SelectedItem.Index)
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
      strSQL = "SELECT B.CONSEC,B.COD_BENEFICIO,B.CEDULA,S.nombre,isnull(O.DESCRIPCION,'') as DESCRIPCION,B.REGISTRA_FECHA,B.REGISTRA_USER" _
             & " FROM AFI_BENE_OTORGA B  inner join SOCIOS S on B.CEDULA = S.CEDULA" _
             & " LEFT JOIN SIF_OFICINAS O ON B.COD_OFICINA = O.COD_OFICINA" _
             & " Where isnull(B.ANALISTA_RECEPCION,0) = 0"
              
    Else
      strSQL = "SELECT B.CONSEC,B.COD_BENEFICIO,B.CEDULA,S.nombre,isnull(O.DESCRIPCION,'') as DESCRIPCION,B.REGISTRA_FECHA,B.REGISTRA_USER" _
             & " FROM AFI_BENE_OTORGA B  inner join SOCIOS S on B.CEDULA = S.CEDULA" _
             & " LEFT JOIN SIF_OFICINAS O ON B.COD_OFICINA = O.COD_OFICINA" _
            & " Where isnull(B.ANALISTA_RECEPCION,0) = 2"
                
    End If
   
                      
    Call sbCargaGrid(vGrid, 7, strSQL)
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
Private Sub txtBeneficio_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
         txtCodigo.SetFocus
    End If

If KeyCode = vbKeyF4 Then
  gBusquedas.Columna = "Cod_Beneficio"
  gBusquedas.Orden = "Cod_Beneficio"
  gBusquedas.Consulta = "select Cod_Beneficio,Descripcion from AFI_BENEFICIOS"
  gBusquedas.Filtro = " and estado = 'A'"
  frmBusquedas.Show vbModal
  txtBeneficio.Text = gBusquedas.Resultado
  
  If Trim(txtBeneficio.Text) <> "" Then txtCodigo.SetFocus
End If

End Sub

Private Sub sbCargarGridConsulta()
Dim strSQL As String, rs As New ADODB.Recordset
Dim vIdCompuesto As String

On Error GoTo vError

If Trim(txtBeneficioBuscar.Text) = Empty Or Trim(txtCodigoBuscar.Text) = Empty Then Exit Sub
    
Me.MousePointer = vbHourglass
    
mCodigo = Trim(txtCodigoBuscar.Text)
mBeneficio = Trim(txtBeneficioBuscar.Text)
vIdCompuesto = mBeneficio & "." & mCodigo

vGridConsulta.MaxCols = 3
vGridConsulta.MaxRows = 0

strSQL = "select T.DESCRIPCION, CT.NOTAS, CT.REGISTRO_FECHA, CT.REGISTRO_USUARIO" _
      & " from SIF_CONTROL_TAGS CT inner join SIF_TAGS T on CT.TAG_CODIGO = T.TAG_CODIGO" _
      & " where CT.codigo = '" & mBeneficio & "' and CT.cod_modulo = 'BEN' and CT.Documento = '" & mCodigo & "'"

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

Private Sub txtBeneficioBuscar_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then txtCodigoBuscar.SetFocus
        
    
    If KeyCode = vbKeyF4 Then
      gBusquedas.Columna = "Cod_Beneficio"
      gBusquedas.Orden = "Cod_Beneficio"
      gBusquedas.Consulta = "select Cod_Beneficio,Descripcion from AFI_BENEFICIOS"
      gBusquedas.Filtro = " and estado = 'A'"
      frmBusquedas.Show vbModal
      txtBeneficioBuscar.Text = gBusquedas.Resultado
      
      If Trim(txtBeneficioBuscar.Text) <> "" Then txtCodigoBuscar.SetFocus
    End If


End Sub

Private Sub txtCodigo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call sbCargaInformacion
    End If
End Sub


Private Sub txtCodigoBuscar_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call sbCargarGridConsulta
    End If
End Sub
