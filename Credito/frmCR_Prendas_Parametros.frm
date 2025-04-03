VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.Controls.v24.0.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.ShortcutBar.v24.0.0.ocx"
Begin VB.Form frmCR_Prendas_Parametros 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Parámetros para Prendas"
   ClientHeight    =   7635
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   13515
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7635
   ScaleWidth      =   13515
   Begin XtremeSuiteControls.TabControl tcMain 
      Height          =   6255
      Left            =   0
      TabIndex        =   1
      Top             =   1320
      Width           =   13455
      _Version        =   1572864
      _ExtentX        =   23733
      _ExtentY        =   11033
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
      ItemCount       =   4
      Item(0).Caption =   "Catálogos Generales"
      Item(0).ControlCount=   3
      Item(0).Control(0)=   "Label1(3)"
      Item(0).Control(1)=   "cboCatalogo"
      Item(0).Control(2)=   "vGrid"
      Item(1).Caption =   "Coberturas Pólizas"
      Item(1).ControlCount=   1
      Item(1).Control(0)=   "gCoberturas"
      Item(2).Caption =   "Comercializadoras"
      Item(2).ControlCount=   1
      Item(2).Control(0)=   "tcComercializa"
      Item(3).Caption =   "Unidades"
      Item(3).ControlCount=   1
      Item(3).Control(0)=   "gUds"
      Begin XtremeSuiteControls.TabControl tcComercializa 
         Height          =   5655
         Left            =   -70000
         TabIndex        =   6
         Top             =   360
         Visible         =   0   'False
         Width           =   13455
         _Version        =   1572864
         _ExtentX        =   23733
         _ExtentY        =   9975
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
         ItemCount       =   2
         Item(0).Caption =   "Lista"
         Item(0).ControlCount=   3
         Item(0).Control(0)=   "lsw"
         Item(0).Control(1)=   "txtComercializaFiltro"
         Item(0).Control(2)=   "ShortcutCaption1"
         Item(1).Caption =   "Mantenimiento"
         Item(1).ControlCount=   17
         Item(1).Control(0)=   "txtCodigo"
         Item(1).Control(1)=   "txtNombre"
         Item(1).Control(2)=   "Label1(0)"
         Item(1).Control(3)=   "cboTipoId"
         Item(1).Control(4)=   "txtCedJur"
         Item(1).Control(5)=   "Label18(3)"
         Item(1).Control(6)=   "lswCuentas"
         Item(1).Control(7)=   "btnCuentas"
         Item(1).Control(8)=   "cboBancos"
         Item(1).Control(9)=   "Label3(0)"
         Item(1).Control(10)=   "Label3(1)"
         Item(1).Control(11)=   "txtCorreo"
         Item(1).Control(12)=   "btnComercializa(0)"
         Item(1).Control(13)=   "btnComercializa(1)"
         Item(1).Control(14)=   "btnComercializa(2)"
         Item(1).Control(15)=   "btnComercializa(3)"
         Item(1).Control(16)=   "chkActivo"
         Begin XtremeSuiteControls.ListView lsw 
            Height          =   4455
            Left            =   120
            TabIndex        =   23
            Top             =   1200
            Width           =   13215
            _Version        =   1572864
            _ExtentX        =   23310
            _ExtentY        =   7858
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
            Appearance      =   17
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.ListView lswCuentas 
            Height          =   1935
            Left            =   -68200
            TabIndex        =   13
            Top             =   2565
            Visible         =   0   'False
            Width           =   8655
            _Version        =   1572864
            _ExtentX        =   15266
            _ExtentY        =   3413
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
            Appearance      =   17
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.CheckBox chkActivo 
            Height          =   375
            Left            =   -60520
            TabIndex        =   27
            Top             =   1080
            Visible         =   0   'False
            Width           =   975
            _Version        =   1572864
            _ExtentX        =   1720
            _ExtentY        =   661
            _StockProps     =   79
            Caption         =   "Activo ?"
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
            Appearance      =   21
            Alignment       =   1
         End
         Begin XtremeSuiteControls.PushButton btnComercializa 
            Height          =   375
            Index           =   0
            Left            =   -64000
            TabIndex        =   19
            Top             =   4680
            Visible         =   0   'False
            Width           =   1455
            _Version        =   1572864
            _ExtentX        =   2566
            _ExtentY        =   661
            _StockProps     =   79
            Caption         =   "Nuevo"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            UseVisualStyle  =   -1  'True
            Appearance      =   17
            Picture         =   "frmCR_Prendas_Parametros.frx":0000
         End
         Begin XtremeSuiteControls.FlatEdit txtCodigo 
            Height          =   330
            Left            =   -67000
            TabIndex        =   7
            Top             =   600
            Visible         =   0   'False
            Width           =   1095
            _Version        =   1572864
            _ExtentX        =   1926
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
            Alignment       =   2
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtNombre 
            Height          =   330
            Left            =   -65920
            TabIndex        =   8
            Top             =   600
            Visible         =   0   'False
            Width           =   6375
            _Version        =   1572864
            _ExtentX        =   11239
            _ExtentY        =   582
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
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.ComboBox cboTipoId 
            Height          =   315
            Left            =   -67000
            TabIndex        =   10
            Top             =   1080
            Visible         =   0   'False
            Width           =   1935
            _Version        =   1572864
            _ExtentX        =   3413
            _ExtentY        =   582
            _StockProps     =   77
            ForeColor       =   1973790
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
         Begin XtremeSuiteControls.FlatEdit txtCedJur 
            Height          =   330
            Left            =   -65080
            TabIndex        =   11
            Top             =   1080
            Visible         =   0   'False
            Width           =   1935
            _Version        =   1572864
            _ExtentX        =   3408
            _ExtentY        =   582
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
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.PushButton btnCuentas 
            Height          =   375
            Left            =   -61240
            TabIndex        =   14
            Tag             =   "1"
            Top             =   2040
            Visible         =   0   'False
            Width           =   1695
            _Version        =   1572864
            _ExtentX        =   2984
            _ExtentY        =   656
            _StockProps     =   79
            Caption         =   "Cuentas Bancarias"
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
         End
         Begin XtremeSuiteControls.ComboBox cboBancos 
            Height          =   330
            Left            =   -67000
            TabIndex        =   15
            Top             =   2085
            Visible         =   0   'False
            Width           =   5295
            _Version        =   1572864
            _ExtentX        =   9340
            _ExtentY        =   582
            _StockProps     =   77
            ForeColor       =   1973790
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
         Begin XtremeSuiteControls.FlatEdit txtCorreo 
            Height          =   330
            Left            =   -67000
            TabIndex        =   18
            Top             =   1680
            Visible         =   0   'False
            Width           =   7455
            _Version        =   1572864
            _ExtentX        =   13150
            _ExtentY        =   582
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
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.PushButton btnComercializa 
            Height          =   375
            Index           =   1
            Left            =   -62560
            TabIndex        =   20
            Top             =   4680
            Visible         =   0   'False
            Width           =   495
            _Version        =   1572864
            _ExtentX        =   873
            _ExtentY        =   661
            _StockProps     =   79
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            UseVisualStyle  =   -1  'True
            Appearance      =   17
            Picture         =   "frmCR_Prendas_Parametros.frx":0632
         End
         Begin XtremeSuiteControls.PushButton btnComercializa 
            Height          =   375
            Index           =   2
            Left            =   -62080
            TabIndex        =   21
            Top             =   4680
            Visible         =   0   'False
            Width           =   495
            _Version        =   1572864
            _ExtentX        =   873
            _ExtentY        =   661
            _StockProps     =   79
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            UseVisualStyle  =   -1  'True
            Appearance      =   17
            Picture         =   "frmCR_Prendas_Parametros.frx":0D63
         End
         Begin XtremeSuiteControls.PushButton btnComercializa 
            Height          =   375
            Index           =   3
            Left            =   -61480
            TabIndex        =   22
            Top             =   4680
            Visible         =   0   'False
            Width           =   495
            _Version        =   1572864
            _ExtentX        =   873
            _ExtentY        =   661
            _StockProps     =   79
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            UseVisualStyle  =   -1  'True
            Appearance      =   17
            Picture         =   "frmCR_Prendas_Parametros.frx":1463
         End
         Begin XtremeSuiteControls.FlatEdit txtComercializaFiltro 
            Height          =   330
            Left            =   120
            TabIndex        =   24
            Top             =   840
            Width           =   13215
            _Version        =   1572864
            _ExtentX        =   23310
            _ExtentY        =   582
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
            Alignment       =   2
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
            Height          =   375
            Left            =   120
            TabIndex        =   25
            Top             =   480
            Width           =   13215
            _Version        =   1572864
            _ExtentX        =   23310
            _ExtentY        =   661
            _StockProps     =   14
            Caption         =   "Lista de Comercializadoras (Agencias, PYMES, Personas Físicas, Otros)"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Alignment       =   1
         End
         Begin VB.Label Label3 
            Caption         =   "Correo"
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
            Index           =   1
            Left            =   -68200
            TabIndex        =   17
            Top             =   1680
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.Label Label3 
            Caption         =   "Cuenta"
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
            Index           =   0
            Left            =   -68200
            TabIndex        =   16
            Top             =   2085
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.Label Label18 
            Caption         =   "Identificación"
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
            Index           =   3
            Left            =   -68200
            TabIndex        =   12
            Top             =   1080
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            BackStyle       =   0  'Transparent
            Caption         =   "Comercializa"
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
            Height          =   315
            Index           =   0
            Left            =   -68200
            TabIndex        =   9
            Top             =   600
            Visible         =   0   'False
            Width           =   1095
         End
      End
      Begin XtremeSuiteControls.ComboBox cboCatalogo 
         Height          =   330
         Left            =   3840
         TabIndex        =   2
         Top             =   480
         Width           =   5055
         _Version        =   1572864
         _ExtentX        =   8916
         _ExtentY        =   582
         _StockProps     =   77
         ForeColor       =   1973790
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
      Begin FPSpreadADO.fpSpread vGrid 
         Height          =   5175
         Left            =   240
         TabIndex        =   4
         Top             =   960
         Width           =   12735
         _Version        =   524288
         _ExtentX        =   22463
         _ExtentY        =   9128
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
         MaxCols         =   5
         ScrollBars      =   2
         SpreadDesigner  =   "frmCR_Prendas_Parametros.frx":1A07
         VScrollSpecial  =   -1  'True
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin FPSpreadADO.fpSpread gUds 
         Height          =   5535
         Left            =   -70000
         TabIndex        =   5
         Top             =   480
         Visible         =   0   'False
         Width           =   13335
         _Version        =   524288
         _ExtentX        =   23521
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
         MaxCols         =   8
         ScrollBars      =   2
         SpreadDesigner  =   "frmCR_Prendas_Parametros.frx":20A0
         VScrollSpecial  =   -1  'True
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin FPSpreadADO.fpSpread gCoberturas 
         Height          =   5655
         Left            =   -69880
         TabIndex        =   26
         Top             =   480
         Visible         =   0   'False
         Width           =   13095
         _Version        =   524288
         _ExtentX        =   23098
         _ExtentY        =   9975
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
         SpreadDesigner  =   "frmCR_Prendas_Parametros.frx":4986
         VScrollSpecial  =   -1  'True
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Catálogo"
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
         Index           =   3
         Left            =   2640
         TabIndex        =   3
         Top             =   480
         Width           =   1095
      End
   End
   Begin XtremeSuiteControls.Label Label2 
      Height          =   495
      Left            =   1800
      TabIndex        =   0
      Top             =   360
      Width           =   6735
      _Version        =   1572864
      _ExtentX        =   11874
      _ExtentY        =   868
      _StockProps     =   79
      Caption         =   "Catálogos para Garantías Prendarias"
      ForeColor       =   16777215
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Transparent     =   -1  'True
   End
   Begin VB.Image imgBanner 
      Appearance      =   0  'Flat
      Height          =   1215
      Left            =   0
      Top             =   0
      Width           =   14055
   End
End
Attribute VB_Name = "frmCR_Prendas_Parametros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem
Dim vPaso As Boolean


Private Sub sbCuentas_Load()

On Error GoTo vError

lswCuentas.ListItems.Clear
If txtCodigo.Text > "0" Then
    strSQL = "select rtrim(B.Descripcion) as 'Banco'" _
           & ",case when C.tipo = 'A' then 'Ahorros' else 'Corriente' end as 'TipoDesc'" _
           & ",C.cod_Divisa,C.CUENTA_INTERNA, C.CUENTA_INTERBANCA, C.ACTIVA, C.DESTINO, C.REGISTRO_FECHA , C.REGISTRO_USUARIO" _
           & " from SYS_CUENTAS_BANCARIAS C inner join TES_BANCOS_GRUPOS B on C.cod_banco = B.cod_grupo" _
           & " where C.Identificacion = '" & Trim(txtCedJur.Text) & "'"
    
    Call OpenRecordSet(rs, strSQL)
    Do While Not rs.EOF
       Set itmX = lswCuentas.ListItems.Add(, , rs!CUENTA_INTERNA)
           itmX.SubItems(1) = Trim(rs!Banco)
           itmX.SubItems(2) = rs!TipoDesc
           itmX.SubItems(3) = rs!cod_Divisa
           itmX.SubItems(4) = IIf(rs!CUENTA_INTERBANCA = 1, "Sí", "No")
           itmX.SubItems(5) = rs!Destino & ""
           itmX.SubItems(6) = IIf(rs!ACTIVA = 1, "Activa", "Cerrada")
           itmX.SubItems(7) = rs!REGISTRO_FECHA & ""
           itmX.SubItems(8) = rs!REGISTRO_USUARIO & ""
     
       rs.MoveNext
    Loop
    rs.Close
End If


Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub btnComercializa_Click(Index As Integer)
Select Case Index
    Case 0 'Nuevo
        Call sbComercializa_Limpia
    Case 1 'Guarda
        Call sbComercializa_Guarda
    Case 2 'Refresca
        If txtCodigo.Text <> "" Then
            Call sbComercializa_Consulta(txtCodigo.Text)
        End If
    Case 3 'Elimina
        If txtCodigo.Text <> "" Then
            Call sbComercializa_Borra
        End If

End Select
End Sub

Private Sub btnCuentas_Click()

If Trim(txtCedJur.Text) = "" Then
   Exit Sub
End If

GLOBALES.gTag = Trim(txtCedJur.Text)
GLOBALES.gTag2 = "CRD"

frmCC_Cuentas_Bancarias.Show vbModal

Call sbCuentas_Load

End Sub

Private Sub cboCatalogo_Click()
If vPaso Then Exit Sub

Call sbConsulta

End Sub

Private Sub Form_Activate()
vModulo = 3
End Sub

Private Sub sbUnidades_Load()
Dim strSQL As String

If vPaso Then Exit Sub

strSQL = "select ID_UNIDAD, DESCRIPCION, PESO_APL, CAPACIDAD_APL, CILINDRAJE_APL, ACTIVA, REGISTRO_FECHA, REGISTRO_USUARIO" _
       & " From crd_prendas_uds"
Call sbCargaGrid(gUds, 8, strSQL)

End Sub


Private Sub sbCobertura_Load()
Dim strSQL As String

If vPaso Then Exit Sub

strSQL = "select ID_COBERTURA, COD_POLIZA, COD_COBERTURA, COBERTURA, DESCRIPCION, ACTIVA, REGISTRO_FECHA, REGISTRO_USUARIO" _
       & " From CRD_PRENDAS_POLIZAS_COBERTURAS"

Call sbCargaGrid(gCoberturas, 8, strSQL)

End Sub

Private Sub sbConsulta()
Dim strSQL As String

If vPaso Then Exit Sub

strSQL = "exec spCrd_Prendas_Cat_List '" & cboCatalogo.ItemData(cboCatalogo.ListIndex) & "'"
Call sbCargaGrid(vGrid, 5, strSQL)

End Sub

Private Sub Form_Load()

vModulo = 3

vPaso = True

tcMain.Item(0).Selected = True

vGrid.AppearanceStyle = fxGridStyle
gUds.AppearanceStyle = fxGridStyle

Set imgBanner.Picture = frmContenedor.imgBanner_Mantenimiento.Picture

cboCatalogo.AddItem "Combustible"
cboCatalogo.ItemData(cboCatalogo.ListCount - 1) = "Cob"
cboCatalogo.AddItem "Marcas"
cboCatalogo.ItemData(cboCatalogo.ListCount - 1) = "Mar"
cboCatalogo.AddItem "Modelos"
cboCatalogo.ItemData(cboCatalogo.ListCount - 1) = "Mod"
cboCatalogo.AddItem "Presentación"
cboCatalogo.ItemData(cboCatalogo.ListCount - 1) = "Pre"
cboCatalogo.AddItem "Extras"
cboCatalogo.ItemData(cboCatalogo.ListCount - 1) = "Ext"
cboCatalogo.AddItem "Aseguradoras"
cboCatalogo.ItemData(cboCatalogo.ListCount - 1) = "Ase"

cboCatalogo.Text = "Combustible"

vPaso = False

 
lsw.ColumnHeaders.Clear
lsw.ColumnHeaders.Add , , "Id", 1200
lsw.ColumnHeaders.Add , , "Nombre", 3500
lsw.ColumnHeaders.Add , , "Activo", 2100, vbCenter
lsw.ColumnHeaders.Add , , "Usuario", 2100, vbCenter
lsw.ColumnHeaders.Add , , "Fecha", 2100, vbCenter
 
lswCuentas.ColumnHeaders.Add 1, , "Cuenta", 2500
lswCuentas.ColumnHeaders.Add 2, , "Banco", 3500
lswCuentas.ColumnHeaders.Add 3, , "Tipo", 1100, vbCenter
lswCuentas.ColumnHeaders.Add 4, , "Divisa", 1100, vbCenter
lswCuentas.ColumnHeaders.Add 5, , "Interbanca", 1100, vbCenter
lswCuentas.ColumnHeaders.Add 6, , "Destino", 1100, vbCenter
lswCuentas.ColumnHeaders.Add 7, , "Activa", 1100, vbCenter
lswCuentas.ColumnHeaders.Add 8, , "Fecha", 2500
lswCuentas.ColumnHeaders.Add 9, , "Usuario", 2500


'Carga Tipos de Identificacion
vPaso = True
strSQL = "select TIPO_ID as Idx, rtrim(Descripcion) as ItmX from AFI_TIPOS_IDS" _
       & " order by Tipo_Id"
    Call sbCbo_Llena_New(cboTipoId, strSQL, False, True)
vPaso = False

strSQL = "exec spCxP_Bancos_Autorizados"
Call sbCbo_Llena_New(cboBancos, strSQL, False, True)


Call sbConsulta

Call Formularios(Me)
Call RefrescaTags(Me)

End Sub


Private Function fxGuardar() As Long
Dim pCodigo As String, pDescripcion As String, pActivo As Integer
Dim pMovimiento As String

On Error GoTo vError

fxGuardar = 0

vGrid.Row = vGrid.ActiveRow
vGrid.Col = 1
pCodigo = vGrid.Text
vGrid.Col = 2
pDescripcion = vGrid.Text
vGrid.Col = 3
pActivo = vGrid.Value

If pCodigo = "" Then
   pCodigo = "0"
End If

strSQL = "exec spCrd_Prendas_Cat_Parametros_Add '" & cboCatalogo.ItemData(cboCatalogo.ListIndex) _
       & "', '" & pCodigo & "', '" & pDescripcion & "', " & pActivo & ", '" & glogon.Usuario & "', 'A'"
Call OpenRecordSet(rs, strSQL)
    vGrid.Col = 1
    vGrid.Text = rs!Codigo
    pMovimiento = rs!Movimiento
rs.Close

Call Bitacora(pMovimiento, "Prendas Cat_" & cboCatalogo.Text & " Id: " & vGrid.Text)

fxGuardar = 1

Exit Function

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Function

Private Function fxGuardar_Coberturas() As Long
Dim pCodigo As String, pDescripcion As String, pActivo As Integer
Dim pPoliza As String, pCoberturaId As String, pCobertura As String

Dim pMovimiento As String

On Error GoTo vError

fxGuardar_Coberturas = 0
 
gCoberturas.Row = gCoberturas.ActiveRow

gCoberturas.Col = 1
pCodigo = gCoberturas.Text

gCoberturas.Col = 2
pPoliza = gCoberturas.Text

gCoberturas.Col = 3
pCoberturaId = gCoberturas.Text

gCoberturas.Col = 4
pCobertura = gCoberturas.Text

gCoberturas.Col = 5
pDescripcion = gCoberturas.Text

gCoberturas.Col = 6
pActivo = gCoberturas.Value

If pCodigo = "" Then
   pCodigo = "0"
End If

strSQL = "exec spCrd_Prendas_Cat_Coberturas_Add '" & pCodigo & "', '" & pPoliza & "', '" & pCoberturaId & "', '" & pCobertura _
       & "', '" & pDescripcion & "', " & pActivo & ", '" & glogon.Usuario & "', 'A'"
Call OpenRecordSet(rs, strSQL)
    gCoberturas.Col = 1
    gCoberturas.Text = rs!Codigo
    pMovimiento = rs!Movimiento
rs.Close

Call Bitacora(pMovimiento, "Prendas Cat_Coberturas Polizas Id: " & gCoberturas.Text)

fxGuardar_Coberturas = 1

Exit Function

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Function


Private Function fxGuardar_Unidades() As Long
Dim pCodigo As String, pDescripcion As String, pActivo As Integer
Dim pPeso As Integer, pCapacidad As Integer, pCilindraje As Integer

Dim pMovimiento As String

On Error GoTo vError

fxGuardar_Unidades = 0

gUds.Row = gUds.ActiveRow

gUds.Col = 1
pCodigo = gUds.Text

If pCodigo = "" Then
    MsgBox "Debe de indicar un código para la unidad!", vbExclamation
    Exit Function
End If

gUds.Col = 2
pDescripcion = gUds.Value

gUds.Col = 3
pPeso = gUds.Text

gUds.Col = 4
pCapacidad = gUds.Value

gUds.Col = 5
pCilindraje = gUds.Value

gUds.Col = 6
pActivo = gUds.Value

                       
strSQL = "exec spCrd_Prendas_Cat_Unidades_Add '" & pCodigo & "', '" & pDescripcion & "', " & pPeso & ", " & pCapacidad & ", " & pCilindraje & ", " _
        & pActivo & ", '" & glogon.Usuario & "', 'A'"
Call OpenRecordSet(rs, strSQL)
    gUds.Col = 1
    gUds.Text = rs!Codigo
    pMovimiento = rs!Movimiento
rs.Close

Call Bitacora(pMovimiento, "Prendas Cat_Unidades Id: " & gUds.Text)

fxGuardar_Unidades = 1

Exit Function

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Function




Private Sub gCoberturas_KeyDown(KeyCode As Integer, Shift As Integer)
Dim i As Integer, strSQL As String


If gCoberturas.ActiveCol = 6 And (KeyCode = vbKeyReturn Or KeyCode = vbKeyTab) Then
  i = fxGuardar_Coberturas
  If i = 0 Then Exit Sub
  gCoberturas.Row = gCoberturas.ActiveRow
  If gCoberturas.MaxRows <= gCoberturas.ActiveRow Then
    gCoberturas.MaxRows = gCoberturas.MaxRows + 1
    gCoberturas.Row = gCoberturas.MaxRows
  End If
End If

'Inserta Linea
If KeyCode = vbKeyInsert Then
    gCoberturas.MaxRows = gCoberturas.MaxRows + 1
    gCoberturas.InsertRows gCoberturas.ActiveRow, 1
    gCoberturas.Row = gCoberturas.ActiveRow
End If


If KeyCode = vbKeyF4 And gCoberturas.ActiveCol = 2 Then
    gBusquedas.Columna = "Codigo"
    gBusquedas.Orden = "Codigo"
    gBusquedas.Consulta = "select Codigo, Descripcion from Catalogo"
    gBusquedas.Filtro = " and Poliza = 'S'"
    
    gBusquedas.Col1Name = "Código"
    gBusquedas.Col2Name = "Descripción"

    frmBusquedas.Show vbModal
    
    If gBusquedas.Resultado <> "" Then
       gCoberturas.Row = gCoberturas.ActiveRow
       gCoberturas.Col = 2
       gCoberturas.Text = gBusquedas.Resultado
    End If

End If

'Borrar una linea
If KeyCode = vbKeyDelete Then
     i = MsgBox("Esta Seguro que desea borrar este registro", vbYesNo)
     If i = vbYes Then

        gCoberturas.Row = gCoberturas.ActiveRow
        gCoberturas.Col = 1
               
        strSQL = "exec spCrd_Prendas_Cat_Coberturas_Add '" & gCoberturas.Text & "', '', '', '', '', 0, '" & glogon.Usuario & "', 'E'"
        Call ConectionExecute(strSQL)
        
        Call Bitacora("Elimina", "Prendas Cat_Coberturas Polizas Id: " & gCoberturas.Text)
        
        Call sbCobertura_Load
     
     End If
End If


End Sub



Private Sub gUds_KeyDown(KeyCode As Integer, Shift As Integer)
Dim i As Integer, strSQL As String

If gUds.ActiveCol = 6 And (KeyCode = vbKeyReturn Or KeyCode = vbKeyTab) Then
  i = fxGuardar_Unidades
  If i = 0 Then Exit Sub
  gUds.Row = gUds.ActiveRow
  If gUds.MaxRows <= gUds.ActiveRow Then
    gUds.MaxRows = gUds.MaxRows + 1
    gUds.Row = gUds.MaxRows
  End If
End If

'Inserta Linea
If KeyCode = vbKeyInsert Then
    gUds.MaxRows = gUds.MaxRows + 1
    gUds.InsertRows gUds.ActiveRow, 1
    gUds.Row = gUds.ActiveRow
End If

'Borrar una linea
If KeyCode = vbKeyDelete Then
     i = MsgBox("Esta Seguro que desea borrar este registro", vbYesNo)
     If i = vbYes Then

        gUds.Row = gUds.ActiveRow
        gUds.Col = 1
        
        strSQL = "exec spCrd_Prendas_Cat_Unidades_Add '" & gUds.Text & "', '', 0, 0, 0, 0, '" & glogon.Usuario & "', 'E'"
        Call ConectionExecute(strSQL)
        
        Call Bitacora("Elimina", "Prendas Cat_Unidades Id: " & gUds.Text)
        
        Call sbUnidades_Load
     
     End If
End If


End Sub

Private Sub lsw_DblClick()
tcComercializa.Item(1).Selected = True
txtCodigo.Text = lsw.SelectedItem.Text
Call sbComercializa_Consulta(txtCodigo.Text)
End Sub

Private Sub tcComercializa_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
    If Item.Selected = 0 Then
        Call sbComercializa_Load
    End If
End Sub

Private Sub tcMain_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)

Select Case Item.Index
    Case 0 'Catalogos
        Call sbConsulta
    
    Case 1 'Coberturas
        Call sbCobertura_Load
        
    Case 2 'Comercializa
        tcComercializa.Item(0).Selected = True
        Call sbComercializa_Load
        
    Case 3 'Unidades
        Call sbUnidades_Load
End Select

End Sub

Private Sub sbComercializa_Load()

On Error GoTo vError

Me.MousePointer = vbHourglass

txtComercializaFiltro.Text = fxSysCleanTxtInject(txtComercializaFiltro.Text)

strSQL = "select ID_COMERCIO, DESCRIPCION, case when ACTIVA = 1 then 'Sí' else 'No' end as 'Activa', REGISTRO_FECHA, REGISTRO_USUARIO " _
       & "  From CRD_PRENDAS_COMERCIA" _
       & " Where DESCRIPCION like '%" & txtComercializaFiltro.Text & "%'" _
       & " order by Activa desc, DESCRIPCION"

lsw.ListItems.Clear
Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
  Set itmX = lsw.ListItems.Add(, , rs!ID_COMERCIO)
      itmX.SubItems(1) = rs!Descripcion
      itmX.SubItems(2) = rs!ACTIVA
      itmX.SubItems(3) = rs!REGISTRO_USUARIO & ""
      itmX.SubItems(4) = rs!REGISTRO_FECHA & ""
  rs.MoveNext
Loop
rs.Close

Me.MousePointer = vbDefault
Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
 
End Sub


Private Sub sbComercializa_Limpia()
   txtCodigo.Text = "0"
   txtNombre.Text = ""
   chkActivo.Value = xtpChecked
   txtCorreo.Text = ""
   
   txtCedJur.Text = ""
   lswCuentas.ListItems.Clear
End Sub

Private Sub sbComercializa_Consulta(pComerciaId As Long)

On Error GoTo vError

Me.MousePointer = vbHourglass

Call sbComercializa_Limpia

strSQL = "exec spCrd_Prendas_Cat_Comercializa_Consulta " & pComerciaId
Call OpenRecordSet(rs, strSQL)
If Not rs.EOF Then
   txtCodigo.Text = rs!ID_COMERCIO
   txtNombre.Text = rs!Descripcion
   chkActivo.Value = rs!ACTIVA
   txtCorreo.Text = rs!correo
   
   txtCedJur.Text = Trim(rs!Cedula)
   
   Call sbCboAsignaDato(cboTipoId, rs!Tipo_Id_Desc, True, rs!Tipo_Id)
   Call sbCboAsignaDato(cboBancos, rs!Banco_Desc, True, rs!Id_Banco)
End If
rs.Close

If txtCodigo.Text <> "0" Then
    Call sbCuentas_Load
End If

Me.MousePointer = vbDefault
Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
 
End Sub


Private Sub sbComercializa_Borra()

On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = "exec spCrd_Prendas_Cat_Comercializa_Add " & txtCodigo.Text & ", " & cboTipoId.ItemData(cboTipoId.ListIndex) _
       & ", '" & txtCedJur.Text & "', '" & txtNombre.Text & "', " & chkActivo.Value & ", " & cboBancos.ItemData(cboBancos.ListIndex) _
       & ", '" & txtCorreo.Text & "', '" & glogon.Usuario & "', 'E'"
Call OpenRecordSet(rs, strSQL)
If rs!Pass = 1 Then
   txtCodigo.Text = rs!Codigo
   Call Bitacora(rs!Movimiento, "Prendas> Comercializador Id: " & txtCodigo.Text)
   
   MsgBox "Se ha eliminado el Comercializador Id: " & txtCodigo.Text, vbInformation
   Call sbComercializa_Limpia
Else
    MsgBox rs!Mensaje, vbExclamation
End If

If txtCodigo.Text <> "0" Then
    Call sbCuentas_Load
End If

Me.MousePointer = vbDefault
Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
 
End Sub

Private Function fxComercializa_Valida() As Boolean
Dim rs As New ADODB.Recordset, i As Integer

Dim IdLargo As Integer

Dim vMensaje As String

vMensaje = ""
  
'Actualiza el Parametro de Validacion y Luego lo Aplica
strSQL = "select LARGO_MINIMO from AFI_TIPOS_IDS Where TIPO_ID = " & cboTipoId.ItemData(cboTipoId.ListIndex)
Call OpenRecordSet(rs, strSQL)
If Not rs.EOF And Not rs.BOF Then
    IdLargo = rs!Largo_Minimo
End If
rs.Close


If Len(Trim(txtCedJur.Text)) <> IdLargo Then vMensaje = vMensaje & " - Número de Identidad no es válido, se espera que sea de: " & txtCedJur _
        & " caracteres, verifique...!" & vbCrLf

If Len(Trim(txtCedJur.Text)) > 20 Then vMensaje = vMensaje & " - Número de Identidad no es válido, verifique...!" & vbCrLf

If Not fxEmail_Valida(txtCorreo.Text) Then
    vMensaje = vMensaje & " - El Email no es válido!" & vbCrLf
End If

If Trim(txtNombre.Text) = "" Then vMensaje = vMensaje & " - Indique el nombre comercial" & vbCrLf

If Len(vMensaje) = 0 Then
  fxComercializa_Valida = True
Else
  fxComercializa_Valida = False
  MsgBox vMensaje, vbExclamation

End If
End Function



Private Sub sbComercializa_Guarda()

On Error GoTo vError

If Not fxComercializa_Valida() Then
   Exit Sub
End If

Me.MousePointer = vbHourglass

strSQL = "exec spCrd_Prendas_Cat_Comercializa_Add " & txtCodigo.Text & ", " & cboTipoId.ItemData(cboTipoId.ListIndex) _
       & ", '" & txtCedJur.Text & "', '" & txtNombre.Text & "', " & chkActivo.Value & ", " & cboBancos.ItemData(cboBancos.ListIndex) _
       & ", '" & txtCorreo.Text & "', '" & glogon.Usuario & "', 'A'"
Call OpenRecordSet(rs, strSQL)
If rs!Pass = 1 Then
   txtCodigo.Text = rs!Codigo
   Call Bitacora(rs!Movimiento, "Prendas> Comercializador Id: " & txtCodigo.Text)

   MsgBox "Se ha " & rs!Movimiento & " el Comercializador Id: " & txtCodigo.Text, vbInformation

Else
    MsgBox rs!Mensaje, vbExclamation
End If

If txtCodigo.Text <> "0" Then
    Call sbCuentas_Load
End If

Me.MousePointer = vbDefault
Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
 
End Sub



Private Sub txtCodigo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtNombre.SetFocus

If KeyCode = vbKeyF4 Then
  gBusquedas.Col1Name = "Id. Comercio"
  gBusquedas.Col2Name = "Id. Real"
  gBusquedas.Col3Name = "Nombre"
  gBusquedas.Convertir = "N"
  gBusquedas.Columna = "Id_Comercio"
  gBusquedas.Orden = "Id_Comercio"
  gBusquedas.Consulta = "select Id_Comercio, Cedula, Descripcion from crd_Prendas_Comercia"
  gBusquedas.Filtro = ""
  frmBusquedas.Show vbModal
  txtCodigo = gBusquedas.Resultado
  If txtCodigo <> "" Then Call sbComercializa_Consulta(CLng(gBusquedas.Resultado))
End If

End Sub


Private Sub txtComercializaFiltro_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    Call sbComercializa_Load
End If
End Sub

Private Sub vGrid_KeyDown(KeyCode As Integer, Shift As Integer)
Dim i As Integer, strSQL As String


If vGrid.ActiveCol = 3 And (KeyCode = vbKeyReturn Or KeyCode = vbKeyTab) Then
  i = fxGuardar
  If i = 0 Then Exit Sub
  vGrid.Row = vGrid.ActiveRow
  If vGrid.MaxRows <= vGrid.ActiveRow Then
    vGrid.MaxRows = vGrid.MaxRows + 1
    vGrid.Row = vGrid.MaxRows
  End If
End If

'Inserta Linea
If KeyCode = vbKeyInsert Then
    vGrid.MaxRows = vGrid.MaxRows + 1
    vGrid.InsertRows vGrid.ActiveRow, 1
    vGrid.Row = vGrid.ActiveRow
End If

'Borrar una linea
If KeyCode = vbKeyDelete Then
     i = MsgBox("Esta Seguro que desea borrar este registro", vbYesNo)
     If i = vbYes Then

        vGrid.Row = vGrid.ActiveRow
        vGrid.Col = 1
        
        strSQL = "exec spCrd_Prendas_Cat_Parametros_Add '" & cboCatalogo.ItemData(cboCatalogo.ListIndex) _
               & "', '" & vGrid.Text & "', '', 0, '" & glogon.Usuario & "', 'E'"
        Call ConectionExecute(strSQL)
        
        Call Bitacora("Elimina", "Prendas Cat_" & cboCatalogo.Text & " Id: " & vGrid.Text)
        
        Call sbConsulta
     
     End If
End If


End Sub



