VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.Controls.v24.0.0.ocx"
Begin VB.Form frmSIF_RecepcionNdNc 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Documentos: Recepción"
   ClientHeight    =   9405
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   12210
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
   MinButton       =   0   'False
   ScaleHeight     =   9405
   ScaleWidth      =   12210
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin XtremeSuiteControls.TabControl tcMain 
      Height          =   7575
      Left            =   120
      TabIndex        =   4
      Top             =   1800
      Width           =   12015
      _Version        =   1572864
      _ExtentX        =   21193
      _ExtentY        =   13361
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
      Appearance      =   4
      Color           =   32
      ItemCount       =   3
      Item(0).Caption =   "Recepción o Devolución"
      Item(0).ControlCount=   7
      Item(0).Control(0)=   "lsw"
      Item(0).Control(1)=   "Label3(1)"
      Item(0).Control(2)=   "txtCodigo"
      Item(0).Control(3)=   "cmdAgregar"
      Item(0).Control(4)=   "rbMovimiento(0)"
      Item(0).Control(5)=   "rbMovimiento(1)"
      Item(0).Control(6)=   "btnAplicar"
      Item(1).Caption =   "Pendientes"
      Item(1).ControlCount=   3
      Item(1).Control(0)=   "vGrid"
      Item(1).Control(1)=   "btnPendientes(0)"
      Item(1).Control(2)=   "btnPendientes(1)"
      Item(2).Caption =   "Consulta"
      Item(2).ControlCount=   9
      Item(2).Control(0)=   "dtpFInicio"
      Item(2).Control(1)=   "dtpFFin"
      Item(2).Control(2)=   "txtCodigoBuscar"
      Item(2).Control(3)=   "cboUsuario"
      Item(2).Control(4)=   "Label3(2)"
      Item(2).Control(5)=   "Label3(3)"
      Item(2).Control(6)=   "Label3(4)"
      Item(2).Control(7)=   "vGridConsulta"
      Item(2).Control(8)=   "btnConsulta"
      Begin XtremeSuiteControls.ListView lsw 
         Height          =   6495
         Left            =   120
         TabIndex        =   5
         Top             =   960
         Width           =   11775
         _Version        =   1572864
         _ExtentX        =   20770
         _ExtentY        =   11456
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
         Checkboxes      =   -1  'True
         View            =   3
         Appearance      =   17
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.RadioButton rbMovimiento 
         Height          =   255
         Index           =   0
         Left            =   5280
         TabIndex        =   9
         Top             =   480
         Width           =   1455
         _Version        =   1572864
         _ExtentX        =   2566
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Recepción"
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
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         Value           =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit txtCodigo 
         Height          =   345
         Left            =   1560
         TabIndex        =   7
         Top             =   480
         Width           =   2295
         _Version        =   1572864
         _ExtentX        =   4048
         _ExtentY        =   609
         _StockProps     =   77
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
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
      Begin XtremeSuiteControls.PushButton cmdAgregar 
         Height          =   375
         Left            =   3960
         TabIndex        =   8
         Top             =   480
         Width           =   615
         _Version        =   1572864
         _ExtentX        =   1085
         _ExtentY        =   661
         _StockProps     =   79
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
         Appearance      =   21
         Picture         =   "frmSIF_RecepcionNdNc.frx":0000
      End
      Begin XtremeSuiteControls.RadioButton rbMovimiento 
         Height          =   255
         Index           =   1
         Left            =   6840
         TabIndex        =   10
         Top             =   480
         Width           =   1455
         _Version        =   1572864
         _ExtentX        =   2566
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Devolución"
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
         UseVisualStyle  =   -1  'True
         Appearance      =   17
      End
      Begin XtremeSuiteControls.PushButton btnAplicar 
         Height          =   375
         Left            =   10560
         TabIndex        =   12
         Top             =   480
         Width           =   1335
         _Version        =   1572864
         _ExtentX        =   2355
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Aplicar"
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
         UseVisualStyle  =   -1  'True
         Appearance      =   21
         Picture         =   "frmSIF_RecepcionNdNc.frx":0720
      End
      Begin FPSpreadADO.fpSpread vGrid 
         Height          =   6495
         Left            =   -69880
         TabIndex        =   13
         Top             =   960
         Visible         =   0   'False
         Width           =   11775
         _Version        =   524288
         _ExtentX        =   20770
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
         SpreadDesigner  =   "frmSIF_RecepcionNdNc.frx":0E47
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin XtremeSuiteControls.PushButton btnPendientes 
         Height          =   375
         Index           =   0
         Left            =   -60760
         TabIndex        =   14
         Top             =   480
         Visible         =   0   'False
         Width           =   1335
         _Version        =   1572864
         _ExtentX        =   2355
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Buscar"
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
         UseVisualStyle  =   -1  'True
         Appearance      =   21
         Picture         =   "frmSIF_RecepcionNdNc.frx":1661
      End
      Begin XtremeSuiteControls.PushButton btnPendientes 
         Height          =   375
         Index           =   1
         Left            =   -59440
         TabIndex        =   15
         Top             =   480
         Visible         =   0   'False
         Width           =   1335
         _Version        =   1572864
         _ExtentX        =   2355
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Exportar"
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
         UseVisualStyle  =   -1  'True
         Appearance      =   21
         Picture         =   "frmSIF_RecepcionNdNc.frx":1D61
      End
      Begin XtremeSuiteControls.DateTimePicker dtpFInicio 
         Height          =   330
         Left            =   -68920
         TabIndex        =   16
         Top             =   720
         Visible         =   0   'False
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
         Left            =   -67480
         TabIndex        =   17
         Top             =   720
         Visible         =   0   'False
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
      Begin XtremeSuiteControls.FlatEdit txtCodigoBuscar 
         Height          =   330
         Left            =   -63520
         TabIndex        =   18
         Top             =   720
         Visible         =   0   'False
         Width           =   2295
         _Version        =   1572864
         _ExtentX        =   4048
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin FPSpreadADO.fpSpread vGridConsulta 
         Height          =   6135
         Left            =   -69400
         TabIndex        =   23
         Top             =   1320
         Visible         =   0   'False
         Width           =   10695
         _Version        =   524288
         _ExtentX        =   18865
         _ExtentY        =   10821
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
         ScrollBars      =   2
         SpreadDesigner  =   "frmSIF_RecepcionNdNc.frx":1ECB
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin XtremeSuiteControls.PushButton btnConsulta 
         Height          =   375
         Left            =   -60280
         TabIndex        =   2
         Top             =   720
         Visible         =   0   'False
         Width           =   1335
         _Version        =   1572864
         _ExtentX        =   2355
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Buscar"
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
         UseVisualStyle  =   -1  'True
         Appearance      =   21
         Picture         =   "frmSIF_RecepcionNdNc.frx":246C
      End
      Begin XtremeSuiteControls.ComboBox cboUsuario 
         Height          =   330
         Left            =   -65920
         TabIndex        =   19
         Top             =   720
         Visible         =   0   'False
         Width           =   2295
         _Version        =   1572864
         _ExtentX        =   4048
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
      Begin XtremeSuiteControls.Label Label3 
         Height          =   255
         Index           =   4
         Left            =   -68920
         TabIndex        =   22
         Top             =   480
         Visible         =   0   'False
         Width           =   1575
         _Version        =   1572864
         _ExtentX        =   2778
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
         WordWrap        =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label3 
         Height          =   255
         Index           =   3
         Left            =   -65920
         TabIndex        =   21
         Top             =   480
         Visible         =   0   'False
         Width           =   1575
         _Version        =   1572864
         _ExtentX        =   2778
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Usuario:"
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
         WordWrap        =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label3 
         Height          =   255
         Index           =   2
         Left            =   -63520
         TabIndex        =   20
         Top             =   480
         Visible         =   0   'False
         Width           =   1575
         _Version        =   1572864
         _ExtentX        =   2778
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "No. Documento:"
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
         WordWrap        =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label3 
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   6
         Top             =   480
         Width           =   1575
         _Version        =   1572864
         _ExtentX        =   2778
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "No. Documento:"
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
         WordWrap        =   -1  'True
      End
   End
   Begin XtremeSuiteControls.ComboBox cboTipodoc 
      Height          =   345
      Left            =   1800
      TabIndex        =   1
      Top             =   1320
      Width           =   3975
      _Version        =   1572864
      _ExtentX        =   7011
      _ExtentY        =   609
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9.75
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
   Begin XtremeSuiteControls.ProgressBar PrgBar 
      Height          =   135
      Left            =   120
      TabIndex        =   11
      Top             =   1680
      Visible         =   0   'False
      Width           =   12015
      _Version        =   1572864
      _ExtentX        =   21193
      _ExtentY        =   238
      _StockProps     =   93
      BackColor       =   -2147483633
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   8520
      Top             =   1680
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
            Picture         =   "frmSIF_RecepcionNdNc.frx":2B6C
            Key             =   "IMG1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSIF_RecepcionNdNc.frx":93CE
            Key             =   "IMG2"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSIF_RecepcionNdNc.frx":FC30
            Key             =   "IMG3"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList3 
      Left            =   8640
      Top             =   2280
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
            Picture         =   "frmSIF_RecepcionNdNc.frx":16492
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSIF_RecepcionNdNc.frx":1CCF4
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSIF_RecepcionNdNc.frx":23556
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSIF_RecepcionNdNc.frx":23670
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSIF_RecepcionNdNc.frx":2378E
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSIF_RecepcionNdNc.frx":29FF0
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin XtremeSuiteControls.Label Label3 
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   3
      Top             =   1320
      Width           =   1575
      _Version        =   1572864
      _ExtentX        =   2778
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Tipo Documento"
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
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Recepción de Documentos"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   480
      Index           =   3
      Left            =   1560
      TabIndex        =   0
      Top             =   360
      Width           =   6252
   End
   Begin VB.Image imgBanner 
      Height          =   1092
      Left            =   0
      Top             =   0
      Width           =   13332
   End
End
Attribute VB_Name = "frmSIF_RecepcionNdNc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem
Dim vPaso As Boolean

Dim mTagRecepcion As String, mTagDevolucion As String
Dim mTagRecepcionDev As String

Dim mCodigo As String, mTipoDoc As String


Private Sub sbParametrosTags()

On Error GoTo vError
    
    '' Busca el parámetro del tag de recepción
    strSQL = "select isnull(valor,'') as 'Valor'from SIF_PARAMETROS where cod_parametro = '10'"
    Call OpenRecordSet(rs, strSQL)
    If Not rs.EOF Then
        mTagRecepcion = rs!Valor
    Else
        MsgBox "Falta agregar el parámetro 10 en la base de datos"
    End If
    rs.Close
    
    If Not mTagRecepcion = Empty Then
    
        strSQL = "select COUNT(*) as 'Existe' FROM sif_tags where TAG_CODIGO = '" & mTagRecepcion & "'"
        Call OpenRecordSet(rs, strSQL)
        If rs!Existe = 0 Then
            mTagRecepcion = Empty
            MsgBox "El código de tag definido el los parámetros para la revisión no existe"
        End If
        rs.Close
        
    End If
    
    '' Busca el parámetro del tag de devolución
    strSQL = "select isnull(valor,'') as 'Valor' from SIF_PARAMETROS where cod_parametro = '11'"
    Call OpenRecordSet(rs, strSQL)
    If Not rs.EOF Then
        mTagDevolucion = rs!Valor
    Else
        MsgBox "Falta agregar el parámetro 11 en la base de datos"
    End If
    rs.Close
    
    If Not mTagDevolucion = Empty Then
    
        strSQL = "select COUNT(*) as 'Existe' FROM sif_tags where TAG_CODIGO = '" & mTagDevolucion & "'"
        Call OpenRecordSet(rs, strSQL)
        If rs!Existe = 0 Then
            mTagRecepcion = Empty
            MsgBox "El código de tag definido el los parámetros para la revisión no existe"
        End If
        rs.Close
        
    End If
    
    strSQL = "select isnull(valor,'') as 'Valor' from SIF_PARAMETROS where cod_parametro = '12'"
    Call OpenRecordSet(rs, strSQL)
    If Not rs.EOF Then
        mTagRecepcionDev = rs!Valor
    Else
        MsgBox "Falta agregar el parámetro 12 en la base de datos"
    End If
    rs.Close
    
    If Not mTagRecepcionDev = Empty Then
    
        strSQL = "select COUNT(*) as 'Existe' FROM sif_tags where TAG_CODIGO = '" & mTagRecepcionDev & "'"
        Call OpenRecordSet(rs, strSQL)
        If rs!Existe = 0 Then
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


On Error GoTo vError

mTipoDoc = cboTipodoc.ItemData(cboTipodoc.ListIndex)

If rbMovimiento(0).Value Then
    strSQL = "Select Top 100 T.COD_TRANSACCION,T.TIPO_DOCUMENTO,T.CLIENTE_IDENTIFICACION, T.CLIENTE_NOMBRE,T.REGISTRO_USUARIO,T.REGISTRO_FECHA" _
            & " from SIF_TRANSACCIONES T " _
            & " where T.TIPO_DOCUMENTO = '" & mTipoDoc & "' and T.ANALISTA_REVISION is null" _
            & "   and isnull(T.ANALISTA_RECEPCION,0) = 0 order by T.REGISTRO_FECHA desc"
Else
    strSQL = "Select Top 100 T.COD_TRANSACCION,T.TIPO_DOCUMENTO,T.CLIENTE_IDENTIFICACION,T.CLIENTE_NOMBRE,T.REGISTRO_USUARIO,T.REGISTRO_FECHA" _
            & " from SIF_TRANSACCIONES T " _
            & " where T.TIPO_DOCUMENTO = '" & mTipoDoc & "' and T.ANALISTA_REVISION is null" _
            & "   and isnull(T.ANALISTA_RECEPCION,0) = 1 order by T.REGISTRO_FECHA desc"

End If
       
Call OpenRecordSet(rs, strSQL)

lsw.ListItems.Clear

Do While Not rs.EOF
    Set itmX = lsw.ListItems.Add(, , rs!Cod_Transaccion)
    itmX.SubItems(1) = rs!TIPO_DOCUMENTO
    itmX.SubItems(2) = RTrim(rs!CLIENTE_IDENTIFICACION)
    itmX.SubItems(3) = rs!CLIENTE_NOMBRE
    itmX.SubItems(4) = rs!REGISTRO_USUARIO
    itmX.SubItems(5) = Format(rs!REGISTRO_FECHA, "yyyy-mm-dd")
  rs.MoveNext
Loop

rs.Close


Exit Sub
    
vError:
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub sbAplicar()
Dim i As Long

On Error GoTo vError

If MsgBox("Está seguro que sea aplicar estas etiquetas", vbExclamation + vbYesNo) = vbNo Then
    Exit Sub
End If

If rbMovimiento(0).Value = True Then
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

PrgBar.Max = lsw.ListItems.Count + 1
PrgBar.Value = 1
PrgBar.Visible = True


With lsw.ListItems

For i = 1 To .Count

    If .Item(i).Checked Then
            If rbMovimiento(0).Value Then
            
                Call sbSIFRegistraTags(.Item(i).SubItems(1), mTagRecepcion, "Recibida la documentación de la liquidación", .Item(i).Text, "DOC" _
                                    , .Item(i).SubItems(1), .Item(i).Text)
                
                
            Else
                Call sbSIFRegistraTags(.Item(i).SubItems(1), mTagDevolucion, "Devolución la documentación de la liquidación", .Item(i).Text, "DOC" _
                                     , .Item(i).SubItems(1), .Item(i).Text)
            End If
    End If
    
    PrgBar.Value = PrgBar.Value + 1
Next i

.Clear

End With

PrgBar.Visible = False

Me.MousePointer = vbDefault


MsgBox "Proceso concluído con éxito!", vbInformation

Call sbCargaInformacion

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical


End Sub

Private Function fxValidaNoDuplicados() As Boolean
Dim i As Integer

    fxValidaNoDuplicados = False

    For i = 1 To lsw.ListItems.Count

        If lsw.ListItems(i).Text = Trim(mCodigo) And lsw.ListItems(i).SubItems(1) = Trim(mTipoDoc) Then
            fxValidaNoDuplicados = True
        End If
        
    Next i

End Function

Private Sub btnAplicar_Click()
  Call sbAplicar
End Sub

Private Sub btnConsulta_Click()
    Call sbConsulta_Seguimiento
End Sub

Private Sub btnPendientes_Click(Index As Integer)

Select Case Index
    Case 0 'Buscar
        Call sbPendientes

    Case 1 'Exportar
        Dim vHeaders As vGridHeaders
            vHeaders.Columnas = 6
            vHeaders.Headers(1) = "No. Transaccion"
            vHeaders.Headers(2) = "Tipo Doc."
            vHeaders.Headers(3) = "Identificación"
            vHeaders.Headers(4) = "Nombre"
            vHeaders.Headers(5) = "Fec.Registro"
            vHeaders.Headers(6) = "Usr.Registro"
        
         Call sbSIFGridExportar(vGrid, vHeaders, "ProGrX_Recepcion_Docs")

End Select

End Sub

Private Sub cboTipodoc_Click()

If vPaso Then Exit Sub

Select Case tcMain.SelectedItem
    Case 0 'Recepcion
        Call sbCargaInformacion
        
    Case 1 'Pendientes
        Call sbPendientes
        
    Case 2 'Consulta
        If cboUsuario.ListCount = 0 Then
            Call sbUsuarios_Load
        End If
        vGridConsulta.MaxRows = 0
End Select

End Sub

Private Sub cmdAgregar_Click()
If vPaso Then Exit Sub
If txtCodigo.Text = "" Then Exit Sub
If cboTipodoc.ListCount < 0 Then Exit Sub
    
Dim pTipoDoc As String
    
On Error GoTo vError
    
pTipoDoc = cboTipodoc.ItemData(cboTipodoc.ListIndex)
    

'Valida no agregar en forma mismo tag en forma consecutiva
If rbMovimiento(0).Value Then
   strSQL = "SELECT dbo.fxSIFValidaTagRev('" & Trim(pTipoDoc) & "','" & Trim(mTagRecepcion) & "','" & Trim(mTagDevolucion) & "','DOC','" & txtCodigo.Text & "',NULL) as 'Existe'"
    Call OpenRecordSet(rs, strSQL)
     If rs!Existe = 2 Then
         MsgBox "No es posible registrar en forma consecutiva dos recepciones del documento " & txtCodigo.Text, vbInformation
        txtCodigo.Text = ""
        rs.Close
        Exit Sub
    End If
Else
    strSQL = "SELECT dbo.fxSIFValidaTagRev(" & Trim(pTipoDoc) & ",'" & Trim(mTagDevolucion) & "','" & Trim(mTagRecepcion) & "','DOC','" & txtCodigo.Text & "',NULL) as 'Existe'"
    Call OpenRecordSet(rs, strSQL)
     If rs!Existe = 3 Then
        MsgBox "No es posible registrar en forma consecutiva dos devoluciones del documento " & txtCodigo.Text
        txtCodigo.Text = ""
        rs.Close
        Exit Sub
    
    End If
End If

strSQL = "SELECT dbo.fxSIFValidaTagRev('" & Trim(pTipoDoc) & "','" & Trim(mTagDevolucion) & "','" & Trim(mTagRecepcion) & "','DOC','" & txtCodigo.Text _
        & "','" & Trim(mTagRecepcionDev) & "') as 'Existe'"
Call OpenRecordSet(rs, strSQL)
If rs!Existe = 4 Then
   MsgBox "No es posible registrar una recepción sin aplicar la devolución del documento " & txtCodigo.Text
   txtCodigo.Text = ""
   rs.Close
   Exit Sub
End If

'Aplica el Movimiento
If rbMovimiento(0).Value Then

    Call sbSIFRegistraTags(pTipoDoc, mTagRecepcion, "Recibida la documentación de la liquidación", txtCodigo.Text, "DOC" _
                        , pTipoDoc, txtCodigo.Text)
    
    
Else
    Call sbSIFRegistraTags(pTipoDoc, mTagDevolucion, "Devolución la documentación de la liquidación", txtCodigo.Text, "DOC" _
                          , pTipoDoc, txtCodigo.Text)
End If


Call sbCargaInformacion

Exit Sub

vError:
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical


End Sub

Private Sub Form_Activate()
vModulo = 8
End Sub

Private Sub Form_Load()

vModulo = 8
   

Set imgBanner.Picture = frmContenedor.imgBanner_Mantenimiento.Picture

With lsw.ColumnHeaders
    .Clear
    .Add , , "Código", 1800
    .Add , , "Tipo", 1800
    .Add , , "Identificación", 1800, vbCenter
    .Add , , "Nombre", 4500
    .Add , , "Usuario", 2800, vbCenter
    .Add , , "Fecha", 1800
End With

tcMain.Item(0).Selected = True

Call sbParametrosTags

dtpFInicio.Value = fxFechaServidor
dtpFFin.Value = dtpFInicio.Value

vGrid.MaxRows = 0

vPaso = True
    strSQL = "select rtrim(Tipo_Documento) as IdX, rtrim(Descripcion) as 'Itmx'" _
           & " from SIF_Documentos" _
           & " where Tipo_documento in('NC','ND','FND','FNC','CA', 'CD.Liq', 'BEAC', 'CBJ', 'FSL', 'REA', 'RH', 'TCP', 'TRFA', 'TCP', 'THCJ', 'TRA', 'THAV')" _
           & " order by Descripcion"
    Call sbCbo_Llena_New(cboTipodoc, strSQL, False, True)
vPaso = False

Call Formularios(Me)
Call RefrescaTags(Me)

Call cboTipodoc_Click

End Sub



Private Sub sbPendientes()
    
On Error GoTo vError
    
    Me.MousePointer = vbHourglass
    
    strSQL = "Select Top 300 T.COD_TRANSACCION,T.TIPO_DOCUMENTO,T.CLIENTE_IDENTIFICACION,T.CLIENTE_NOMBRE,T.REGISTRO_USUARIO,T.REGISTRO_FECHA" _
           & " from SIF_TRANSACCIONES T" _
           & " where T.TIPO_DOCUMENTO = '" & cboTipodoc.ItemData(cboTipodoc.ListIndex) _
           & "' and isnull(T.ANALISTA_RECEPCION,0) = 0 order by T.REGISTRO_FECHA desc"
    Call sbCargaGrid(vGrid, 6, strSQL)
    vGrid.MaxRows = vGrid.MaxRows - 1
    
    Me.MousePointer = vbDefault

Exit Sub

vError:
    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical
  
    
End Sub


Private Sub lsw_ColumnClick(ByVal ColumnHeader As XtremeSuiteControls.ListViewColumnHeader)
 lsw.SortKey = ColumnHeader.Index - 1
  If lsw.SortOrder = 0 Then lsw.SortOrder = 1 Else lsw.SortOrder = 0
  lsw.Sorted = True
End Sub

Private Sub rbMovimiento_Click(Index As Integer)
Call sbCargaInformacion
End Sub

Private Sub tcMain_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)

Select Case Item.Index
    Case 0 'Recepcion
        Call sbCargaInformacion
        
    Case 1 'Pendientes
        Call sbPendientes
        
    Case 2 'Consulta
        If cboUsuario.ListCount = 0 Then
            Call sbUsuarios_Load
        End If
        vGridConsulta.MaxRows = 0
    
End Select
End Sub

Private Sub sbConsulta_Seguimiento()

On Error GoTo vError

    If txtCodigoBuscar.Text = Empty Then Exit Sub
    mCodigo = Trim(txtCodigoBuscar.Text)
    mTipoDoc = cboTipodoc.ItemData(cboTipodoc.ListIndex)
    
    Me.MousePointer = vbHourglass
    
    strSQL = "select T.DESCRIPCION, CT.NOTAS, CT.REGISTRO_FECHA, CT.REGISTRO_USUARIO" _
           & " from SIF_CONTROL_TAGS CT inner join SIF_TAGS T on CT.TAG_CODIGO = T.TAG_CODIGO" _
           & " where CT.codigo = '" & mTipoDoc & "' and cod_modulo = 'DOC' and documento = '" & mCodigo & "' order by CT.REGISTRO_FECHA desc"
            
    vGridConsulta.MaxCols = 3
    vGridConsulta.MaxRows = 0

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


Private Sub sbUsuarios_Load()

On Error GoTo vError
    
    Me.MousePointer = vbHourglass

    strSQL = "SELECT UPPER(NOMBRE) as 'ItmX', Nombre as 'IdX' from USUARIOS WHERE ESTADO = 'A'"
    
    Call sbCbo_Llena_New(cboUsuario, strSQL, True, False)

    Me.MousePointer = vbDefault
    
Exit Sub

vError:
    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub
