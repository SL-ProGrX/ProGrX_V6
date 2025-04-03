VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "Codejock.Controls.v22.1.0.ocx"
Begin VB.Form frmSYS_Mercadito_Informes 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Informes de Ventas"
   ClientHeight    =   10515
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   16620
   LinkTopic       =   "Form1"
   ScaleHeight     =   10515
   ScaleWidth      =   16620
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin XtremeSuiteControls.FlatEdit txtClienteId 
      Height          =   330
      Left            =   240
      TabIndex        =   17
      Top             =   1920
      Width           =   2655
      _Version        =   1441793
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
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.ProgressBar prgBar 
      Height          =   135
      Left            =   14160
      TabIndex        =   10
      Top             =   1800
      Visible         =   0   'False
      Width           =   2055
      _Version        =   1441793
      _ExtentX        =   3625
      _ExtentY        =   238
      _StockProps     =   93
   End
   Begin XtremeSuiteControls.TabControl tcMain 
      Height          =   7935
      Left            =   240
      TabIndex        =   7
      Top             =   2400
      Width           =   16095
      _Version        =   1441793
      _ExtentX        =   28390
      _ExtentY        =   13996
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
      SelectedItem    =   2
      Item(0).Caption =   "Resumen"
      Item(0).ControlCount=   1
      Item(0).Control(0)=   "gResumen"
      Item(1).Caption =   "Detalle"
      Item(1).ControlCount=   1
      Item(1).Control(0)=   "gDetalle"
      Item(2).Caption =   "Inventario"
      Item(2).ControlCount=   6
      Item(2).Control(0)=   "gInventario"
      Item(2).Control(1)=   "Label2(8)"
      Item(2).Control(2)=   "cboEstado"
      Item(2).Control(3)=   "Label2(9)"
      Item(2).Control(4)=   "txtExistencia"
      Item(2).Control(5)=   "cboExistencia"
      Begin FPSpreadADO.fpSpread gResumen 
         Height          =   7455
         Left            =   -70000
         TabIndex        =   8
         Top             =   360
         Visible         =   0   'False
         Width           =   15975
         _Version        =   524288
         _ExtentX        =   28178
         _ExtentY        =   13150
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
         MaxCols         =   9
         SpreadDesigner  =   "frmSYS_Mercadito_Informes.frx":0000
         VScrollSpecial  =   -1  'True
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin FPSpreadADO.fpSpread gDetalle 
         Height          =   7455
         Left            =   -70000
         TabIndex        =   9
         Top             =   360
         Visible         =   0   'False
         Width           =   15975
         _Version        =   524288
         _ExtentX        =   28178
         _ExtentY        =   13150
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
         MaxCols         =   18
         SpreadDesigner  =   "frmSYS_Mercadito_Informes.frx":0798
         VScrollSpecial  =   -1  'True
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin FPSpreadADO.fpSpread gInventario 
         Height          =   6735
         Left            =   0
         TabIndex        =   26
         Top             =   960
         Width           =   15975
         _Version        =   524288
         _ExtentX        =   28178
         _ExtentY        =   11880
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
         MaxCols         =   11
         SpreadDesigner  =   "frmSYS_Mercadito_Informes.frx":120A
         VScrollSpecial  =   -1  'True
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin XtremeSuiteControls.ComboBox cboEstado 
         Height          =   330
         Left            =   1560
         TabIndex        =   28
         Top             =   480
         Width           =   1455
         _Version        =   1441793
         _ExtentX        =   2566
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
      Begin XtremeSuiteControls.FlatEdit txtExistencia 
         Height          =   330
         Left            =   5160
         TabIndex        =   30
         Top             =   480
         Width           =   975
         _Version        =   1441793
         _ExtentX        =   1720
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
         Text            =   "0"
         Alignment       =   2
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.ComboBox cboExistencia 
         Height          =   330
         Left            =   4200
         TabIndex        =   31
         Top             =   480
         Width           =   975
         _Version        =   1441793
         _ExtentX        =   1720
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
      Begin XtremeSuiteControls.Label Label2 
         Height          =   255
         Index           =   9
         Left            =   3240
         TabIndex        =   29
         Top             =   480
         Width           =   975
         _Version        =   1441793
         _ExtentX        =   1714
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Existencia"
         ForeColor       =   8388608
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   4
         Transparent     =   -1  'True
         WordWrap        =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label2 
         Height          =   255
         Index           =   8
         Left            =   600
         TabIndex        =   27
         Top             =   480
         Width           =   975
         _Version        =   1441793
         _ExtentX        =   1714
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Productos"
         ForeColor       =   8388608
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   4
         Transparent     =   -1  'True
         WordWrap        =   -1  'True
      End
   End
   Begin XtremeSuiteControls.ComboBox cboProveedor 
      Height          =   330
      Left            =   3000
      TabIndex        =   0
      Top             =   1320
      Width           =   4095
      _Version        =   1441793
      _ExtentX        =   7223
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
   Begin XtremeSuiteControls.PushButton btnBuscar 
      Height          =   375
      Left            =   14160
      TabIndex        =   1
      Top             =   1920
      Width           =   495
      _Version        =   1441793
      _ExtentX        =   868
      _ExtentY        =   656
      _StockProps     =   79
      BackColor       =   -2147483633
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Picture         =   "frmSYS_Mercadito_Informes.frx":1A40
   End
   Begin XtremeSuiteControls.PushButton btnExportar 
      Height          =   375
      Index           =   1
      Left            =   14640
      TabIndex        =   2
      Top             =   1920
      Width           =   1575
      _Version        =   1441793
      _ExtentX        =   2778
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Exportar"
      BackColor       =   -2147483633
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
      Picture         =   "frmSYS_Mercadito_Informes.frx":2140
   End
   Begin XtremeSuiteControls.DateTimePicker dtpInicio 
      Height          =   330
      Left            =   240
      TabIndex        =   5
      Top             =   1320
      Width           =   1335
      _Version        =   1441793
      _ExtentX        =   2355
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
      Left            =   1560
      TabIndex        =   6
      Top             =   1320
      Width           =   1335
      _Version        =   1441793
      _ExtentX        =   2355
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
   Begin XtremeSuiteControls.FlatEdit txtClienteNombre 
      Height          =   330
      Left            =   3000
      TabIndex        =   12
      Top             =   1920
      Width           =   4095
      _Version        =   1441793
      _ExtentX        =   7223
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
   Begin XtremeSuiteControls.FlatEdit txtAutorizacion 
      Height          =   330
      Left            =   7800
      TabIndex        =   15
      Top             =   1920
      Width           =   1935
      _Version        =   1441793
      _ExtentX        =   3413
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
   Begin XtremeSuiteControls.FlatEdit txtMarca 
      Height          =   330
      Left            =   9960
      TabIndex        =   18
      Top             =   1920
      Width           =   1935
      _Version        =   1441793
      _ExtentX        =   3413
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
   Begin XtremeSuiteControls.FlatEdit txtModelo 
      Height          =   330
      Left            =   12000
      TabIndex        =   20
      Top             =   1920
      Width           =   1935
      _Version        =   1441793
      _ExtentX        =   3413
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
   Begin XtremeSuiteControls.PushButton btnCboFilter 
      Height          =   330
      Index           =   0
      Left            =   7200
      TabIndex        =   22
      ToolTipText     =   "Proveedores Utilizados"
      Top             =   1320
      Width           =   375
      _Version        =   1441793
      _ExtentX        =   661
      _ExtentY        =   582
      _StockProps     =   79
      BackColor       =   -2147483633
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Picture         =   "frmSYS_Mercadito_Informes.frx":2A11
   End
   Begin XtremeSuiteControls.ComboBox cboCategoria 
      Height          =   330
      Left            =   7800
      TabIndex        =   24
      Top             =   1320
      Width           =   4095
      _Version        =   1441793
      _ExtentX        =   7223
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
   Begin XtremeSuiteControls.PushButton btnCboFilter 
      Height          =   330
      Index           =   1
      Left            =   12000
      TabIndex        =   25
      ToolTipText     =   "Categorias Utilizadas"
      Top             =   1320
      Width           =   375
      _Version        =   1441793
      _ExtentX        =   661
      _ExtentY        =   582
      _StockProps     =   79
      BackColor       =   -2147483633
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Picture         =   "frmSYS_Mercadito_Informes.frx":3119
   End
   Begin XtremeSuiteControls.Label lblLoading 
      Height          =   255
      Left            =   14160
      TabIndex        =   32
      Top             =   1560
      Visible         =   0   'False
      Width           =   2055
      _Version        =   1441793
      _ExtentX        =   3625
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Cargando...Espere!"
      ForeColor       =   16777215
      BackColor       =   16761024
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   6
      WordWrap        =   -1  'True
   End
   Begin XtremeSuiteControls.Label Label2 
      Height          =   255
      Index           =   7
      Left            =   3000
      TabIndex        =   23
      Top             =   1080
      Width           =   4095
      _Version        =   1441793
      _ExtentX        =   7223
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Proveedores"
      ForeColor       =   8388608
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   6
      Transparent     =   -1  'True
      WordWrap        =   -1  'True
   End
   Begin XtremeSuiteControls.Label Label2 
      Height          =   255
      Index           =   6
      Left            =   12120
      TabIndex        =   21
      Top             =   1680
      Width           =   1695
      _Version        =   1441793
      _ExtentX        =   2990
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Modelo"
      ForeColor       =   8388608
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   6
      Transparent     =   -1  'True
      WordWrap        =   -1  'True
   End
   Begin XtremeSuiteControls.Label Label2 
      Height          =   255
      Index           =   5
      Left            =   10080
      TabIndex        =   19
      Top             =   1680
      Width           =   1695
      _Version        =   1441793
      _ExtentX        =   2990
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Marca"
      ForeColor       =   8388608
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   6
      Transparent     =   -1  'True
      WordWrap        =   -1  'True
   End
   Begin XtremeSuiteControls.Label Label2 
      Height          =   255
      Index           =   4
      Left            =   360
      TabIndex        =   16
      Top             =   1680
      Width           =   2415
      _Version        =   1441793
      _ExtentX        =   4260
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Cliente Id"
      ForeColor       =   8388608
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   6
      Transparent     =   -1  'True
      WordWrap        =   -1  'True
   End
   Begin XtremeSuiteControls.Label Label2 
      Height          =   255
      Index           =   0
      Left            =   7920
      TabIndex        =   14
      Top             =   1680
      Width           =   1815
      _Version        =   1441793
      _ExtentX        =   3201
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "No Autorización"
      ForeColor       =   8388608
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   6
      Transparent     =   -1  'True
      WordWrap        =   -1  'True
   End
   Begin XtremeSuiteControls.Label Label2 
      Height          =   255
      Index           =   3
      Left            =   9240
      TabIndex        =   13
      Top             =   1080
      Width           =   1095
      _Version        =   1441793
      _ExtentX        =   1931
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Categoria"
      ForeColor       =   8388608
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   6
      Transparent     =   -1  'True
      WordWrap        =   -1  'True
   End
   Begin XtremeSuiteControls.Label Label2 
      Height          =   255
      Index           =   2
      Left            =   3120
      TabIndex        =   11
      Top             =   1680
      Width           =   3975
      _Version        =   1441793
      _ExtentX        =   7011
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Cliente Nombre"
      ForeColor       =   8388608
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   6
      Transparent     =   -1  'True
      WordWrap        =   -1  'True
   End
   Begin XtremeSuiteControls.Label Label2 
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   4
      Top             =   1080
      Width           =   975
      _Version        =   1441793
      _ExtentX        =   1714
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Fechas"
      ForeColor       =   8388608
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   4
      Transparent     =   -1  'True
      WordWrap        =   -1  'True
   End
   Begin XtremeSuiteControls.Label Label1 
      Height          =   615
      Left            =   3240
      TabIndex        =   3
      Top             =   120
      Width           =   4575
      _Version        =   1441793
      _ExtentX        =   8064
      _ExtentY        =   1080
      _StockProps     =   79
      Caption         =   "Informe de Ventas"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Transparent     =   -1  'True
      WordWrap        =   -1  'True
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      Height          =   1050
      Left            =   0
      Picture         =   "frmSYS_Mercadito_Informes.frx":3821
      Top             =   0
      Width           =   3000
   End
End
Attribute VB_Name = "frmSYS_Mercadito_Informes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim db As New ADODB.Connection
Dim strSQL As String, rs As New ADODB.Recordset
Dim pEmpresaId As Long, vPaso As Boolean


Private Sub sbLimpia()
On Error GoTo vError

    gResumen.MaxRows = 0
    gDetalle.MaxRows = 0

vError:
End Sub







Public Sub sbCargaGrid_Local(vGrid As Object, vGridMaxCol As Integer, pSQL As String, Optional vBorra As Boolean = True)
Dim rsGrid As New ADODB.Recordset, i As Integer

On Error GoTo vErrorLoad

If vBorra Then
    vGrid.MaxCols = vGridMaxCol
    vGrid.MaxRows = 1
    vGrid.Row = vGrid.MaxRows
    For i = 1 To vGrid.MaxCols
     vGrid.Col = i
     vGrid.Text = ""
    Next i
End If



lblLoading.Visible = True
prgBar.Visible = True

DoEvents

rsGrid.CursorLocation = adUseClient
rsGrid.Open strSQL, db, adOpenStatic, adLockReadOnly

   
prgBar.Max = rsGrid.RecordCount + 1
  
vGrid.MaxRows = 1
Do While Not rsGrid.EOF
  vGrid.Row = vGrid.MaxRows
  For i = 1 To vGrid.MaxCols
  
    vGrid.Col = i
    Select Case vGrid.CellType
        Case CellTypeDate
            vGrid.Text = Format(IIf(IsNull(rsGrid.Fields(i - 1).Value), "", rsGrid.Fields(i - 1)), "dd/mm/yyyy")
        Case Else
            vGrid.Text = CStr(IIf(IsNull(rsGrid.Fields(i - 1).Value), "", rsGrid.Fields(i - 1)))
    End Select
  

  Next i
  vGrid.MaxRows = vGrid.MaxRows + 1
  
  prgBar.Value = vGrid.MaxRows
  rsGrid.MoveNext
Loop
rsGrid.Close

prgBar.Visible = False
lblLoading.Visible = False

Exit Sub

vErrorLoad:
  MsgBox fxSys_Error_Handler(Err.Description), vbExclamation
  
End Sub

Private Sub sbResumen_Load()
On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = "select IdVenta, Fecha, Autorizacion, Proveedor , Identificacion, Nombre" _
       & " , SubTotal, Descuento, Total" _
       & " From vVenta_Informe_Resumen" _
       & " Where idEmpresa = " & pEmpresaId _
       & " and Fecha between '" & Format(dtpInicio.Value, "yyyy-MM-dd") _
       & " 00:00:00' and '" & Format(dtpCorte.Value, "yyyy-MM-dd") _
       & " 23:59:59'" _
  
If cboProveedor.Text <> "TODOS" Then
  strSQL = strSQL & " and IdProveedor = " & cboProveedor.ItemData(cboProveedor.ListIndex)
End If

If Len(txtAutorizacion.Text) > 0 Then
    strSQL = strSQL & " and Autorizacion like '%" & txtAutorizacion.Text & "%'"
End If

If Len(txtClienteId.Text) > 0 Then
    strSQL = strSQL & " and Identificacion like '%" & txtClienteId.Text & "%'"
End If

If Len(txtClienteNombre.Text) > 0 Then
    strSQL = strSQL & " and Nombre like '%" & txtClienteNombre.Text & "%'"
End If


Call sbCargaGrid_Local(gResumen, gResumen.MaxCols, strSQL)

Me.MousePointer = vbDefault

Exit Sub

vError:
    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub sbDetalle_Load()

On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = "select IdVenta, Fecha, Autorizacion, Proveedor , Identificacion, Nombre, Categoria" _
       & ", Producto, Modelo, Marca, Codigo, Cantidad, SubTotal, Descuento, Precio, PrecioRack, RetiroDias, RetiroMaximo" _
       & " From vVenta_Informe_Detallado" _
       & " Where idEmpresa = " & pEmpresaId _
       & " and Fecha between '" & Format(dtpInicio.Value, "yyyy-MM-dd") _
       & " 00:00:00' and '" & Format(dtpCorte.Value, "yyyy-MM-dd") _
       & " 23:59:59'"
       
If cboProveedor.Text <> "TODOS" Then
  strSQL = strSQL & " and IdProveedor = " & cboProveedor.ItemData(cboProveedor.ListIndex)
End If

If cboCategoria.Text <> "TODOS" Then
    strSQL = strSQL & " and IdCategoria = " & cboCategoria.ItemData(cboCategoria.ListIndex)
End If

If Len(txtAutorizacion.Text) > 0 Then
    strSQL = strSQL & " and Autorizacion like '%" & txtAutorizacion.Text & "%'"
End If

If Len(txtClienteId.Text) > 0 Then
    strSQL = strSQL & " and Identificacion like '%" & txtClienteId.Text & "%'"
End If

If Len(txtClienteNombre.Text) > 0 Then
    strSQL = strSQL & " and Nombre like '%" & txtClienteNombre.Text & "%'"
End If


If Len(txtModelo.Text) > 0 Then
    strSQL = strSQL & " and Modelo like '%" & txtModelo.Text & "%'"
End If

If Len(txtMarca.Text) > 0 Then
    strSQL = strSQL & " and Marca like '%" & txtMarca.Text & "%'"
End If

Call sbCargaGrid_Local(gDetalle, gDetalle.MaxCols, strSQL)


Me.MousePointer = vbDefault

Exit Sub

vError:
    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub sbInventario_Load()

On Error GoTo vError

Me.MousePointer = vbHourglass


' --Created_2023-10-12 [PBN]
' create view vProveedorProductosExistencias
' as
'  select P.IdProveedor, P.IdTipoIdentificacion, P.Identificacion
'        , P.Descripcion as 'Proveedor', Cat.IdCategoria, Cat.Descripcion as 'Categoria'
'        , REPLACE(REPLACE(REPLACE(convert(nvarchar(max),Pd.Descripcion ),CHaR(10),' ') ,CHaR(13),' ') ,'  ',' ')  as 'Producto'
'        , Pd.Codigo,  Pd.Cantidad,  Pd.Modelo , Pd.Marca , Pd.RetiroDias , Pd.RetiroMaximo , Pd.Estado
'        , Pe.IdEmpresa
'  from  Proveedor P
'        inner join Producto Pd on P.IdProveedor = Pd.IdProveedor
'        inner join Categoria Cat on Pd.IdCategoria = Cat.IdCategoria
'        inner join ProveedoresEmpresas Pe on Pe.IdProveedor = P.IdProveedor

strSQL = "select Proveedor , Categoria, Codigo, Producto, Modelo, Marca, Cantidad, Precio, PrecioRack, RetiroDias, RetiroMaximo" _
       & " From vProveedorProductosExistencias" _
       & " Where idEmpresa = " & pEmpresaId _

If cboProveedor.Text <> "TODOS" Then
  strSQL = strSQL & " and IdProveedor = " & cboProveedor.ItemData(cboProveedor.ListIndex)
End If

If cboCategoria.Text <> "TODOS" Then
    strSQL = strSQL & " and IdCategoria = " & cboCategoria.ItemData(cboCategoria.ListIndex)
End If

If Len(txtModelo.Text) > 0 Then
    strSQL = strSQL & " and Modelo like '%" & txtModelo.Text & "%'"
End If

If Len(txtMarca.Text) > 0 Then
    strSQL = strSQL & " and Marca like '%" & txtMarca.Text & "%'"
End If

Select Case Mid(cboEstado.Text, 1, 1)
    Case "A"
        strSQL = strSQL & " and Estado = 1"
    Case "I"
        strSQL = strSQL & " and Estado = 0"
    Case "T"
End Select

If Not IsNumeric(txtExistencia.Text) Then
    txtExistencia.Text = "0"
End If

strSQL = strSQL & " and Cantidad " & cboExistencia.Text & " " & txtExistencia.Text

strSQL = strSQL & " order by Proveedor, Producto"

Call sbCargaGrid_Local(gInventario, gInventario.MaxCols, strSQL)


Me.MousePointer = vbDefault

Exit Sub

vError:
    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub btnBuscar_Click()

If vPaso Then Exit Sub

On Error GoTo vError

txtClienteId.Text = fxSysCleanTxtInject(txtClienteId.Text)
txtClienteNombre.Text = fxSysCleanTxtInject(txtClienteNombre.Text)
txtAutorizacion.Text = fxSysCleanTxtInject(txtAutorizacion.Text)

txtMarca.Text = fxSysCleanTxtInject(txtMarca.Text)
txtModelo.Text = fxSysCleanTxtInject(txtModelo.Text)


Select Case tcMain.SelectedItem
    Case 0 'Resumen
        Call sbResumen_Load
    Case 1 'Detalle
        Call sbDetalle_Load
    Case 2 'Inventario
        Call sbInventario_Load
End Select

vError:

End Sub

Private Sub btnCboFilter_Click(Index As Integer)

Select Case Index
    Case 0 'Proveedore
        Call sbCbo_Load(cboProveedor, "P", 0)
    Case 1 'Categorias
        Call sbCbo_Load(cboCategoria, "C", 0)
End Select

End Sub

Private Sub btnExportar_Click(Index As Integer)

Dim vHeaders As vGridHeaders

On Error GoTo vError




Select Case tcMain.SelectedItem
    Case 0

            vHeaders.Columnas = 9
            vHeaders.Headers(1) = "Id Venta"
            vHeaders.Headers(2) = "Fecha"
            vHeaders.Headers(3) = "Autorización"
            vHeaders.Headers(4) = "Proveedor"
            vHeaders.Headers(5) = "Identificación"
            vHeaders.Headers(6) = "Nombre"
            vHeaders.Headers(7) = "Sub Total"
            vHeaders.Headers(8) = "Descuento"
            vHeaders.Headers(9) = "Total"
        
         Call sbSIFGridExportar(gResumen, vHeaders, "ProGrX_Mercadito_Ventas_Resumen")
    
    Case 1

            vHeaders.Columnas = 18
            vHeaders.Headers(1) = "Id Venta"
            vHeaders.Headers(2) = "Fecha"
            vHeaders.Headers(3) = "Autorización"
            vHeaders.Headers(4) = "Proveedor"
            vHeaders.Headers(5) = "Identificación"
            vHeaders.Headers(6) = "Nombre"
            vHeaders.Headers(7) = "Categoría"
            vHeaders.Headers(8) = "Producto"
            vHeaders.Headers(9) = "Modelo"
            vHeaders.Headers(10) = "Marca"
            vHeaders.Headers(11) = "Código"
            vHeaders.Headers(12) = "Cantidad"
            vHeaders.Headers(13) = "SubTotal"
            vHeaders.Headers(14) = "Descuento"
            vHeaders.Headers(15) = "Precio"
            vHeaders.Headers(16) = "Precio Rack"
            vHeaders.Headers(17) = "Retiro Días"
            vHeaders.Headers(18) = "Retiro Máximo"
            
         Call sbSIFGridExportar(gDetalle, vHeaders, "ProGrX_Mercadito_Ventas_Detalle")

    Case 2

            vHeaders.Columnas = 11
            vHeaders.Headers(1) = "Proveedor"
            vHeaders.Headers(2) = "Categoría"
            vHeaders.Headers(3) = "Código"
            vHeaders.Headers(4) = "Producto"
            vHeaders.Headers(5) = "Modelo"
            vHeaders.Headers(6) = "Marca"
            vHeaders.Headers(7) = "Existencia"
            vHeaders.Headers(8) = "Precio"
            vHeaders.Headers(9) = "Precio Rack"
            vHeaders.Headers(10) = "Retiro Días"
            vHeaders.Headers(11) = "Retiro Máximo"
            
        
         Call sbSIFGridExportar(gInventario, vHeaders, "ProGrX_Mercadito_Inventario")

End Select



Exit Sub

vError:
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub sbCbo_Load(cbo As XtremeSuiteControls.ComboBox, pTipo As String, Optional pInicial As Integer = 0)

On Error GoTo vError

cbo.Clear

Select Case True

 Case pTipo = "P" And pInicial = 1
    strSQL = "select P.IdProveedor as 'IdX', P.Descripcion as 'ItmX'" _
        & "   from Proveedor P" _
        & "    inner join ProveedoresEmpresas Pe on P.IdProveedor = Pe.IdProveedor" _
        & "    inner join Empresa E on Pe.IdEmpresa = Pe.IdEmpresa" _
        & "  Where E.IdEmpresa = " & pEmpresaId _
        & "  Order by P.Descripcion"
        
 Case pTipo = "C" And pInicial = 1
    strSQL = "select IdCategoria as 'IdX', Descripcion as 'ItmX'" _
           & "  From Categoria  Where idEmpresa = " & pEmpresaId _
           & "  order by Descripcion"
    
 Case pTipo = "P" And pInicial = 0
    strSQL = "select IdProveedor as 'IdX', Proveedor as 'ItmX'" _
           & "  From vVenta_Informe_Detallado" _
           & "  Where idEmpresa = " & pEmpresaId _
           & "    and Fecha between '" & Format(dtpInicio.Value, "yyyy-mm-dd") & " 00:00:00' and '" & Format(dtpCorte.Value, "yyyy-mm-dd") & " 23:59:59'" _
           & " group by  IdProveedor, Proveedor" _
           & " order by Proveedor"
    
 Case pTipo = "C" And pInicial = 0
    strSQL = "select IdCategoria as 'IdX', Categoria as 'ItmX'" _
           & "  From vVenta_Informe_Detallado" _
           & "  Where idEmpresa = " & pEmpresaId _
           & "    and Fecha between '" & Format(dtpInicio.Value, "yyyy-mm-dd") & " 00:00:00' and '" & Format(dtpCorte.Value, "yyyy-mm-dd") & " 23:59:59'" _
           & " group by IdCategoria, Categoria" _
           & " order by Categoria"
End Select
    
rs.Open strSQL, db, adOpenStatic
Do While Not rs.EOF
 cbo.AddItem rs!itmX & ""
 cbo.ItemData(cbo.ListCount - 1) = CStr(rs!IdX)
 rs.MoveNext
Loop
rs.Close

cbo.AddItem "TODOS"
cbo.ItemData(cbo.ListCount - 1) = "TODOS"
cbo.Text = "TODOS"

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub Form_Load()

On Error GoTo vError

'Temporal
pEmpresaId = 2


vPaso = True

gResumen.MaxCols = 9
gDetalle.MaxCols = 18

strSQL = "select PORTAL_ID from sif_Empresa"
Call OpenRecordSet(rs, strSQL)
    pEmpresaId = rs!Portal_Id
rs.Close


'Establece Conexion
strSQL = "PROVIDER=MSDASQL;Driver={SQL Server};Server=progrx.centralus.cloudapp.azure.com" _
       & ";Database=ElMercadito;APP=PGX_APL_Access;tcp:progrx.centralus.cloudapp.azure.com" _
       & "," & SIFGlobal.PuertosDisponibles & ";"
       
db.ConnectionString = strSQL
db.Open , "31M3rcadit0", "#S0n+oFl*v3M4t3w1/*"


strSQL = "select idEmpresa from Empresa where ClienteAPL = " & pEmpresaId
rs.Open strSQL, db, adOpenStatic
If Not rs.EOF And Not rs.BOF Then
   pEmpresaId = rs!idEmpresa
Else
    pEmpresaId = -1
End If
rs.Close

Call sbCbo_Load(cboProveedor, "P", 1)
Call sbCbo_Load(cboCategoria, "C", 1)

cboEstado.Clear
cboEstado.AddItem "Activos"
cboEstado.AddItem "Inactivos"
cboEstado.AddItem "TODOS"
cboEstado.Text = "Activos"

cboExistencia.AddItem " > "
cboExistencia.AddItem " >= "
cboExistencia.AddItem " = "
cboExistencia.AddItem " < "
cboExistencia.AddItem " <= "
cboExistencia.Text = " >= "


dtpCorte.Value = fxFechaServidor
dtpInicio.Value = DateAdd("d", -7, dtpCorte.Value)

Call sbLimpia

tcMain.Item(0).Selected = True

vPaso = False

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub Form_Resize()
On Error Resume Next

tcMain.Width = Me.Width - (tcMain.Left + 150)
tcMain.Height = Me.Height - (tcMain.Top + 250)

gResumen.Width = tcMain.Width - 50
gResumen.Height = tcMain.Height - 400

gDetalle.Width = gResumen.Width
gDetalle.Height = gResumen.Height

gInventario.Width = gResumen.Width
gInventario.Height = gResumen.Height - gInventario.Top

End Sub




Private Sub tcMain_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
Call btnBuscar_Click
End Sub
