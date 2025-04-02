VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.controls.v22.1.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.shortcutbar.v22.1.0.ocx"
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
      _Version        =   1441793
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
      SelectedItem    =   3
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
         _Version        =   1441793
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
         Item(1).ControlCount=   18
         Item(1).Control(0)=   "txtCodigo"
         Item(1).Control(1)=   "txtNombre"
         Item(1).Control(2)=   "Label1(0)"
         Item(1).Control(3)=   "cboTipoId"
         Item(1).Control(4)=   "txtCedJur"
         Item(1).Control(5)=   "txtCodAlter"
         Item(1).Control(6)=   "Label18(3)"
         Item(1).Control(7)=   "Label14(1)"
         Item(1).Control(8)=   "lswCuentas"
         Item(1).Control(9)=   "btnCuentas"
         Item(1).Control(10)=   "cboBancos"
         Item(1).Control(11)=   "Label3(0)"
         Item(1).Control(12)=   "Label3(1)"
         Item(1).Control(13)=   "txtCorreo"
         Item(1).Control(14)=   "btnComercializa(0)"
         Item(1).Control(15)=   "btnComercializa(1)"
         Item(1).Control(16)=   "btnComercializa(2)"
         Item(1).Control(17)=   "btnComercializa(3)"
         Begin XtremeSuiteControls.ListView lswCuentas 
            Height          =   1695
            Left            =   -68200
            TabIndex        =   15
            Top             =   2445
            Visible         =   0   'False
            Width           =   8655
            _Version        =   1441793
            _ExtentX        =   15261
            _ExtentY        =   2984
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
         Begin XtremeSuiteControls.ListView lsw 
            Height          =   4455
            Left            =   120
            TabIndex        =   25
            Top             =   1200
            Width           =   13215
            _Version        =   1441793
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
            Appearance      =   17
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.PushButton btnComercializa 
            Height          =   375
            Index           =   0
            Left            =   -64000
            TabIndex        =   21
            Top             =   4440
            Visible         =   0   'False
            Width           =   1455
            _Version        =   1441793
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
            _Version        =   1441793
            _ExtentX        =   1926
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
         Begin XtremeSuiteControls.FlatEdit txtNombre 
            Height          =   330
            Left            =   -65920
            TabIndex        =   8
            Top             =   600
            Visible         =   0   'False
            Width           =   6375
            _Version        =   1441793
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
            _Version        =   1441793
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
            _Version        =   1441793
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
         Begin XtremeSuiteControls.FlatEdit txtCodAlter 
            Height          =   330
            Left            =   -61720
            TabIndex        =   12
            Top             =   1080
            Visible         =   0   'False
            Width           =   2175
            _Version        =   1441793
            _ExtentX        =   3836
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
            TabIndex        =   16
            Tag             =   "1"
            Top             =   2040
            Visible         =   0   'False
            Width           =   1695
            _Version        =   1441793
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
            TabIndex        =   17
            Top             =   2085
            Visible         =   0   'False
            Width           =   5295
            _Version        =   1441793
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
            TabIndex        =   20
            Top             =   1680
            Visible         =   0   'False
            Width           =   7455
            _Version        =   1441793
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
            TabIndex        =   22
            Top             =   4440
            Visible         =   0   'False
            Width           =   495
            _Version        =   1441793
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
            TabIndex        =   23
            Top             =   4440
            Visible         =   0   'False
            Width           =   495
            _Version        =   1441793
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
            TabIndex        =   24
            Top             =   4440
            Visible         =   0   'False
            Width           =   495
            _Version        =   1441793
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
            TabIndex        =   26
            Top             =   840
            Width           =   13215
            _Version        =   1441793
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
            TabIndex        =   27
            Top             =   480
            Width           =   13215
            _Version        =   1441793
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
            TabIndex        =   19
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
            TabIndex        =   18
            Top             =   2085
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.Label Label14 
            Caption         =   "ID Alterno"
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
            Left            =   -62920
            TabIndex        =   14
            Top             =   1080
            Visible         =   0   'False
            Width           =   975
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
            TabIndex        =   13
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
         Left            =   -66160
         TabIndex        =   2
         Top             =   480
         Visible         =   0   'False
         Width           =   5055
         _Version        =   1441793
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
         Left            =   -69760
         TabIndex        =   4
         Top             =   960
         Visible         =   0   'False
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
         Left            =   0
         TabIndex        =   5
         Top             =   480
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
         SpreadDesigner  =   "frmCR_Prendas_Parametros.frx":2091
         VScrollSpecial  =   -1  'True
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin FPSpreadADO.fpSpread gCoberturas 
         Height          =   5655
         Left            =   -69880
         TabIndex        =   28
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
         ScrollBars      =   2
         SpreadDesigner  =   "frmCR_Prendas_Parametros.frx":4968
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
         Left            =   -67360
         TabIndex        =   3
         Top             =   480
         Visible         =   0   'False
         Width           =   1095
      End
   End
   Begin XtremeSuiteControls.Label Label2 
      Height          =   495
      Left            =   1800
      TabIndex        =   0
      Top             =   360
      Width           =   6735
      _Version        =   1441793
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

strSQL = "select ID_COBERTURA, COD_POLIZA, COD_COBERTURA, COBERTURA, DESCRIPCION, ACTIVA " _
       & " From CRD_PRENDAS_POLIZAS_COBERTURAS"

Call sbCargaGrid(gCoberturas, 6, strSQL)
strSQL = "select ID_UNIDAD, DESCRIPCION, PESO_APL, CAPACIDAD_APL, CILINDRAJE_APL, ACTIVA, REGISTRO_FECHA, REGISTRO_USUARIO" _
       & " From crd_prendas_uds"

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
       & " Where TIPO_PERSONERIA = 'F' order by Tipo_Id"
    Call sbCbo_Llena_New(cboTipoId, strSQL, False, True)
vPaso = False



strSQL = "exec spCxP_Bancos_Autorizados"
Call sbCbo_Llena_New(cboBancos, strSQL, False, True)




Call sbConsulta

Call Formularios(Me)
Call RefrescaTags(Me)

End Sub


Private Function fxGuardar() As Long
Dim strSQL As String, rs As New ADODB.Recordset
'Guarda la información de la linea
'si es Insert devuelve el codigo, sino devuelve 0

On Error GoTo vError

fxGuardar = 0
vGrid.Row = vGrid.ActiveRow
vGrid.col = 1
'
'strSQL = "select isnull(count(*),0) as Existe from CRD_PRENDAS_TIPOS " _
'       & " where TIPO_PRENDA = '" & vGrid.Text & "'"
'Call OpenRecordSet(rs, strSQL)
'
'If rs!Existe = 0 Then 'Insertar
'  If Trim(vGrid.Text) = "" Then Exit Function
'
'  strSQL = "insert into CRD_PRENDAS_TIPOS(TIPO_PRENDA,DESCRIPCION, PORC_COBERTURA" _
'         & ", ACTIVA, REGISTRO_USUARIO, REGISTRO_FECHA) values('" _
'         & UCase(vGrid.Text) & "','"
'  vGrid.col = 2
'  strSQL = strSQL & vGrid.Text & "',"
'  vGrid.col = 3
'  strSQL = strSQL & CCur(vGrid.Text) & ","
'  vGrid.col = 4
'  strSQL = strSQL & vGrid.Value & ",'" & glogon.Usuario & "',dbo.Mygetdate())"
'
'  Call ConectionExecute(strSQL)
'
'  vGrid.col = 1
'  Call Bitacora("Registra", "Tipo de Prenda: " & vGrid.Text)
'
'Else 'Actualizar
'
' vGrid.col = 2
' strSQL = "update CRD_PRENDAS_TIPOS set descripcion = '" & vGrid.Text & "', PORC_COBERTURA = "
' vGrid.col = 3
' strSQL = strSQL & CCur(vGrid.Text) & ", ACTIVA = "
' vGrid.col = 4
' strSQL = strSQL & vGrid.Value & " where TIPO_PRENDA = '"
' vGrid.col = 1
' strSQL = strSQL & vGrid.Text & "'"
' Call ConectionExecute(strSQL)
'
' vGrid.col = 1
' Call Bitacora("Modifica", "Tipo de Prenda: " & vGrid.Text)
'
'End If
'rs.Close

fxGuardar = 1

Exit Function

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Function



Private Sub lsw_DblClick()
tcComercializa.Item(1).Selected = True
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
  Set itmX = lsw.ListItems.Add(, , rs!ID_Comercio)
      itmX.SubItems(1) = rs!Descripcion
      itmX.SubItems(2) = rs!Activa
      itmX.SubItems(3) = rs!Registro_Usuario & ""
      itmX.SubItems(4) = rs!Registro_Fecha & ""
  rs.MoveNext
Loop
rs.Close

Me.MousePointer = vbDefault
Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
 
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
        
'        vGrid.Row = vGrid.ActiveRow
'        vGrid.col = 1
'        strSQL = "delete CRD_PRENDAS_TIPOS where TIPO_PRENDA = '" & vGrid.Text & "'"
'        Call ConectionExecute(strSQL)
'        strSQL = vGrid.Text
'        vGrid.col = 1
'        Call Bitacora("Elimina", "Tipo de Prenda: " & vGrid.Text)
        
        Call sbConsulta
     
     End If
End If


End Sub



