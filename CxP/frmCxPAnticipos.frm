VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.controls.v22.1.0.ocx"
Begin VB.Form frmCxPAnticipos 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Registro de Adelantos a Proveedores"
   ClientHeight    =   8175
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10500
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8175
   ScaleWidth      =   10500
   Begin VB.Timer TimerX 
      Interval        =   10
      Left            =   0
      Top             =   0
   End
   Begin XtremeSuiteControls.TabControl tcMain 
      Height          =   6252
      Left            =   120
      TabIndex        =   2
      Top             =   1800
      Width           =   10332
      _Version        =   1441793
      _ExtentX        =   18224
      _ExtentY        =   11028
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
      Item(0).Caption =   "Registro de Adelanto"
      Item(0).ControlCount=   24
      Item(0).Control(0)=   "dtpCobroAnticipo"
      Item(0).Control(1)=   "cmdNuevo"
      Item(0).Control(2)=   "txtAnticipo"
      Item(0).Control(3)=   "txtCargoCod"
      Item(0).Control(4)=   "txtCargoDesc"
      Item(0).Control(5)=   "txtNotas"
      Item(0).Control(6)=   "txtDocumento"
      Item(0).Control(7)=   "opt(0)"
      Item(0).Control(8)=   "opt(1)"
      Item(0).Control(9)=   "txtOrden"
      Item(0).Control(10)=   "txtDisponible"
      Item(0).Control(11)=   "txtMonto"
      Item(0).Control(12)=   "vGridCargos"
      Item(0).Control(13)=   "Label3(9)"
      Item(0).Control(14)=   "Label3(7)"
      Item(0).Control(15)=   "Label3(5)"
      Item(0).Control(16)=   "Label3(1)"
      Item(0).Control(17)=   "Label3(6)"
      Item(0).Control(18)=   "Label3(0)"
      Item(0).Control(19)=   "Label3(2)"
      Item(0).Control(20)=   "Label3(3)"
      Item(0).Control(21)=   "Label3(4)"
      Item(0).Control(22)=   "cmdGuardar"
      Item(0).Control(23)=   "Label3(8)"
      Item(1).Caption =   "Historial"
      Item(1).ControlCount=   2
      Item(1).Control(0)=   "GroupBox1"
      Item(1).Control(1)=   "GroupBox2"
      Begin XtremeSuiteControls.GroupBox GroupBox1 
         Height          =   2892
         Left            =   -69880
         TabIndex        =   18
         Top             =   480
         Visible         =   0   'False
         Width           =   10212
         _Version        =   1441793
         _ExtentX        =   18013
         _ExtentY        =   5101
         _StockProps     =   79
         Caption         =   "Adelantos Registrados"
         ForeColor       =   4210752
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
         Appearance      =   16
         BorderStyle     =   1
         Begin XtremeSuiteControls.ListView lsw 
            Height          =   2412
            Left            =   120
            TabIndex        =   19
            Top             =   360
            Width           =   9972
            _Version        =   1441793
            _ExtentX        =   17590
            _ExtentY        =   4254
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
            Appearance      =   16
         End
      End
      Begin VB.OptionButton opt 
         Appearance      =   0  'Flat
         BackColor       =   &H00C00000&
         Caption         =   "Cargo directo"
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
         Height          =   315
         Index           =   1
         Left            =   6000
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   600
         Width           =   2772
      End
      Begin VB.OptionButton opt 
         Appearance      =   0  'Flat
         BackColor       =   &H00C00000&
         Caption         =   "Cargo flotante"
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
         Height          =   315
         Index           =   0
         Left            =   3240
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   600
         Value           =   -1  'True
         Width           =   2772
      End
      Begin XtremeSuiteControls.PushButton cmdNuevo 
         Height          =   612
         Left            =   6720
         TabIndex        =   3
         Top             =   5400
         Width           =   1332
         _Version        =   1441793
         _ExtentX        =   2350
         _ExtentY        =   1080
         _StockProps     =   79
         Caption         =   "Nuevo"
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
         Picture         =   "frmCxPAnticipos.frx":0000
      End
      Begin FPSpreadADO.fpSpread vGridCargos 
         Height          =   2292
         Left            =   4560
         TabIndex        =   6
         Top             =   2760
         Width           =   4812
         _Version        =   524288
         _ExtentX        =   8488
         _ExtentY        =   4043
         _StockProps     =   64
         BackColorStyle  =   1
         BorderStyle     =   0
         DisplayRowHeaders=   0   'False
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
         MaxCols         =   484
         ScrollBars      =   2
         SpreadDesigner  =   "frmCxPAnticipos.frx":07B9
         VScrollSpecial  =   -1  'True
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin XtremeSuiteControls.PushButton cmdGuardar 
         Height          =   612
         Left            =   8040
         TabIndex        =   16
         Top             =   5400
         Width           =   1332
         _Version        =   1441793
         _ExtentX        =   2350
         _ExtentY        =   1080
         _StockProps     =   79
         Caption         =   "Guardar"
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
         Picture         =   "frmCxPAnticipos.frx":0D39
      End
      Begin XtremeSuiteControls.GroupBox GroupBox2 
         Height          =   2892
         Left            =   -69880
         TabIndex        =   20
         Top             =   3360
         Visible         =   0   'False
         Width           =   10212
         _Version        =   1441793
         _ExtentX        =   18013
         _ExtentY        =   5101
         _StockProps     =   79
         Caption         =   "Historial de Pagos"
         ForeColor       =   4210752
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
         Appearance      =   16
         BorderStyle     =   1
         Begin XtremeSuiteControls.ListView lswPago 
            Height          =   2412
            Left            =   120
            TabIndex        =   21
            Top             =   360
            Width           =   9972
            _Version        =   1441793
            _ExtentX        =   17590
            _ExtentY        =   4254
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
            Appearance      =   16
         End
      End
      Begin XtremeSuiteControls.FlatEdit txtCargoCod 
         Height          =   312
         Left            =   1800
         TabIndex        =   22
         Top             =   1080
         Width           =   1452
         _Version        =   1441793
         _ExtentX        =   2561
         _ExtentY        =   550
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
      Begin XtremeSuiteControls.FlatEdit txtCargoDesc 
         Height          =   312
         Left            =   3240
         TabIndex        =   23
         Top             =   1080
         Width           =   6132
         _Version        =   1441793
         _ExtentX        =   10816
         _ExtentY        =   550
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
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtNotas 
         Height          =   792
         Left            =   1800
         TabIndex        =   26
         Top             =   1440
         Width           =   7572
         _Version        =   1441793
         _ExtentX        =   13356
         _ExtentY        =   1397
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
         MultiLine       =   -1  'True
         ScrollBars      =   2
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtDocumento 
         Height          =   312
         Left            =   2760
         TabIndex        =   27
         Top             =   2760
         Width           =   1692
         _Version        =   1441793
         _ExtentX        =   2984
         _ExtentY        =   550
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
      Begin XtremeSuiteControls.FlatEdit txtOrden 
         Height          =   312
         Left            =   2760
         TabIndex        =   28
         Top             =   3120
         Width           =   1692
         _Version        =   1441793
         _ExtentX        =   2984
         _ExtentY        =   550
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
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtAnticipo 
         Height          =   312
         Left            =   2760
         TabIndex        =   29
         Top             =   3480
         Width           =   1692
         _Version        =   1441793
         _ExtentX        =   2984
         _ExtentY        =   550
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
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtMonto 
         Height          =   312
         Left            =   2760
         TabIndex        =   31
         Top             =   4320
         Width           =   1692
         _Version        =   1441793
         _ExtentX        =   2984
         _ExtentY        =   550
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
         Alignment       =   1
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtDisponible 
         Height          =   312
         Left            =   2760
         TabIndex        =   30
         Top             =   3840
         Width           =   1692
         _Version        =   1441793
         _ExtentX        =   2984
         _ExtentY        =   550
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
         Alignment       =   1
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.DateTimePicker dtpCobroAnticipo 
         Height          =   312
         Left            =   2760
         TabIndex        =   32
         Top             =   4800
         Width           =   1692
         _Version        =   1441793
         _ExtentX        =   2984
         _ExtentY        =   556
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
      Begin VB.Label Label3 
         Caption         =   "Otras Rebajos Directos:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   8
         Left            =   4560
         TabIndex        =   17
         Top             =   2400
         Width           =   2172
      End
      Begin VB.Label Label3 
         Caption         =   "Monto del Adelanto"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   4
         Left            =   360
         TabIndex        =   15
         Top             =   4320
         Width           =   3252
      End
      Begin VB.Label Label3 
         Caption         =   "Monto Disponible en la Orden "
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   3
         Left            =   360
         TabIndex        =   14
         Top             =   3960
         Width           =   3252
      End
      Begin VB.Label Label3 
         Caption         =   "No. de Orden de Compra"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   2
         Left            =   360
         TabIndex        =   13
         Top             =   3120
         Width           =   3132
      End
      Begin VB.Label Label3 
         Caption         =   "No. Documento o Referencia"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   0
         Left            =   360
         TabIndex        =   12
         Top             =   2760
         Width           =   3132
      End
      Begin VB.Label Label3 
         Caption         =   "Detalle"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Index           =   6
         Left            =   960
         TabIndex        =   11
         Top             =   1440
         Width           =   732
      End
      Begin VB.Label Label3 
         Caption         =   "Cargo"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   1
         Left            =   960
         TabIndex        =   10
         Top             =   1080
         Width           =   732
      End
      Begin VB.Label Label3 
         Caption         =   "No. Adelanto"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   5
         Left            =   360
         TabIndex        =   9
         Top             =   3480
         Width           =   1932
      End
      Begin VB.Label Label3 
         Caption         =   "Tipo de Cargo"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   7
         Left            =   1800
         TabIndex        =   8
         Top             =   600
         Width           =   1332
      End
      Begin VB.Label Label3 
         Caption         =   "Cobro de Adelanto en próximos pagos a partir de:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   9
         Left            =   360
         TabIndex        =   7
         Top             =   4680
         Width           =   2415
      End
   End
   Begin XtremeSuiteControls.FlatEdit txtCodigo 
      Height          =   312
      Left            =   1920
      TabIndex        =   24
      ToolTipText     =   "Presione F4 para Consultar"
      Top             =   1320
      Width           =   1452
      _Version        =   1441793
      _ExtentX        =   2561
      _ExtentY        =   550
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
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtNombre 
      Height          =   312
      Left            =   3360
      TabIndex        =   25
      Top             =   1320
      Width           =   6132
      _Version        =   1441793
      _ExtentX        =   10816
      _ExtentY        =   550
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
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Registro de Adelantos"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   612
      Left            =   1680
      TabIndex        =   1
      Top             =   360
      Width           =   6972
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Proveedor"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   360
      TabIndex        =   0
      Top             =   1320
      Width           =   1452
   End
   Begin VB.Image imgBanner 
      Height          =   1092
      Left            =   0
      Top             =   0
      Width           =   10812
   End
End
Attribute VB_Name = "frmCxPAnticipos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vTipoCambio  As Currency, vPaso As Boolean


Private Function fxVerifica() As Boolean
Dim vMensaje As String

On Error GoTo vError

vMensaje = ""
fxVerifica = False

If Not IsNumeric(txtMonto) Then
 vMensaje = vMensaje & vbCrLf & " - El monto no es válido..."
Else
  If CCur(txtMonto) < 0 Then vMensaje = vMensaje & vbCrLf & " - El monto no es válido..."
End If

If txtCodigo = "" Then vMensaje = vMensaje & vbCrLf & " - El Proveedor no es válido..."
If txtCargoDesc = "" Then vMensaje = vMensaje & vbCrLf & " - El Cargo no es válido..."

If opt.Item(1).Value Then
   If txtOrden = "" Then vMensaje = vMensaje & vbCrLf & " - No se especificó la Orden de Compra..."
   If txtMonto > txtDisponible Then vMensaje = vMensaje & vbCrLf & " - El monto es mayor al disponible de la Orden de Compra..."
End If

If Len(vMensaje) > 0 Then
  fxVerifica = False
  MsgBox vMensaje, vbExclamation
Else
  fxVerifica = True
End If


Exit Function
vError:
End Function




Private Sub cmdGuardar_Click()
Dim strSQL As String, rs As New ADODB.Recordset
Dim vDocumento As String, vID As Integer
Dim curCargos As Currency, i As Integer, y As Integer
Dim vDivisa As String


'Guardar el Registro del Anticipo
'Programar Cargo al Proveedor
'Programar Pago del Anticipo

If Not fxVerifica Then Exit Sub

On Error GoTo vError

vDivisa = "COL"
vTipoCambio = 1

Me.MousePointer = vbHourglass

With vGridCargos
 curCargos = 0
 For i = 1 To .MaxRows
   .col = 3
   .Row = i
   curCargos = curCargos + CCur(.Text)
 Next i
End With


If curCargos > CCur(txtMonto.Text) Then
 Me.MousePointer = vbDefault
 MsgBox " - Los Cargos Son Mayores que el Monto del anticipo", vbExclamation
 Exit Sub
End If


'spCxP_Anticipos(@Proveedor int, @CargoCod varchar(10), @Monto dec(16,2), @Divisa varchar(10), @Documento varchar(30)
'                           ,@Notas varchar(500),  @Orden varchar(30), @Usuario varchar(30))
strSQL = "exec spCxP_Anticipos " & txtCodigo.Text & ",'" & txtCargoCod.Text & "'," & CCur(txtMonto.Text) & ",'" & vDivisa & "'" _
       & ",'" & txtDocumento.Text & "','" & txtNotas.Text & "','" & glogon.Usuario & "','" & Format(dtpCobroAnticipo.Value, "yyyy/mm/dd") & "'"
Call ConectionExecute(strSQL)



'Procedimiento 1 de la Aplicación de Cargos
strSQL = ""
If curCargos > 0 Then
 With vGridCargos
   For y = 1 To .MaxRows
      .Row = y
      .col = 3
      curCargos = CCur(.Text)
      If curCargos > 0 Then
           .col = 1
           i = 1
           strSQL = strSQL & Space(10) & "insert cxp_PagoProvCargos(Npago,Cod_factura,cod_proveedor,cod_cargo,monto,registro_fecha,registro_usuario" _
                  & ",cod_divisa,tipo_cambio,tipo_cargo,tipo_proceso)" _
                  & " values(" & i & ",'" & txtAnticipo.Text & "'," & txtCodigo.Text _
                  & ",'" & Trim(.Text) & "'," & curCargos & ",dbo.MyGetdate(),'" & glogon.Usuario & "','" & vDivisa _
                  & "'," & vTipoCambio & ",'M','D')"
      End If
   Next y
 End With
End If 'Chk y CurCargos > 0

'Procesa Todos los Cargos
If Len(strSQL) > 0 Then
    Call ConectionExecute(strSQL)
End If


tcMain.Item(1).Selected = True


Me.MousePointer = vbDefault
MsgBox "Anticipo Programado...", vbInformation


Exit Sub
vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub sbProgCargosIni()
Dim strSQL As String

strSQL = "select cod_Cargo,descripcion,0 as Monto " _
       & " from cxp_cargos where Activo = 1"
Call sbCargaGrid(vGridCargos, 3, strSQL)
'Limpia espacios vacios
vGridCargos.MaxRows = vGridCargos.MaxRows - 1

End Sub



Private Sub cmdNuevo_Click()
Dim strSQL As String, rs As New ADODB.Recordset


strSQL = "select (isnull(max(IDX),0) + 1) as Consecutivo from cxp_anticipos"

If Len(txtCodigo) = 0 Then
  strSQL = strSQL & " where cod_proveedor = 0"
Else
  strSQL = strSQL & " where cod_proveedor = " & txtCodigo
End If

Call OpenRecordSet(rs, strSQL)
'txtAnticipo = "ANT" & Format(rs!Consecutivo, "0000000")
txtAnticipo = "ANT." & Trim(txtCodigo.Text) & "." & Format(rs!Consecutivo, "00000")


rs.Close

Call sbProgCargosIni

'txtCodigo = ""
'txtNombre = ""
txtCargoCod = ""
txtCargoDesc = ""

txtDocumento = ""
txtNotas = ""
txtOrden = ""
txtDisponible = "0.00"
txtMonto = 0

opt.Item(0).Value = True
opt.Item(0).ForeColor = vbWhite
opt.Item(0).BackColor = vbBlue



Call opt_Click(0)

txtCodigo.SetFocus

Call RefrescaTags(Me)


End Sub

Private Sub Form_Activate()
vModulo = 30
End Sub

Private Sub Form_Load()

vModulo = 30

tcMain.Item(0).Selected = True

vTipoCambio = 1

Set imgBanner.Picture = frmContenedor.imgBanner_Mantenimiento.Picture


With lsw.ColumnHeaders
    .Clear
    .Add , , "Anticipo", 1500
    .Add , , "Fecha", 1800, vbCenter
    .Add , , "Monto", 1500, vbRightJustify
    .Add , , "Saldo [CxC]", 1500, vbRightJustify
    .Add , , "Usuario", 1500, vbCenter
    .Add , , "Documento", 1500, vbCenter
    .Add , , "No.Orden", 1500, vbCenter
    .Add , , "Flotante", 1500, vbCenter
    .Add , , "Nota", 3000
    .Add , , "Cargo", 1500, vbRightJustify
    .Add , , "Pago", 1500, vbRightJustify
    .Add , , "Vencimiento", 1400, vbCenter
    .Add , , "Cobra a Partir", 1800, vbCenter
End With

With lswPago.ColumnHeaders
    .Clear
    .Add , , "Anticipo", 1500
    .Add , , "Fecha", 1500, vbCenter
    .Add , , "Monto", 1500, vbRightJustify
    .Add , , "Factura", 2000, vbCenter
    .Add , , "Pago [Id]", 1400, vbCenter
    .Add , , "Usuario", 1500, vbCenter
    .Add , , "Divisa", 1200, vbCenter
    .Add , , "T.C.", 1200, vbRightJustify
End With


dtpCobroAnticipo.Value = fxFechaServidor
dtpCobroAnticipo.MinDate = dtpCobroAnticipo.Value
dtpCobroAnticipo.MaxDate = DateAdd("d", 45, dtpCobroAnticipo.Value)

Call Formularios(Me)
Call RefrescaTags(Me)

End Sub



Private Sub lsw_ItemClick(ByVal Item As XtremeSuiteControls.ListViewItem)

If vPaso Then Exit Sub

Dim itmX As ListViewItem, strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = "select A.ANTICIPOS, P.COD_PROVEEDOR, P.COD_FACTURA, P.REGISTRO_FECHA, REGISTRO_USUARIO, P.MONTO , P.COD_DIVISA, P.TIPO_CAMBIO, P.NPAGO " _
       & " from CXP_ANTICIPOS A inner join CXP_PAGOPROVCARGOS P on A.ID_CARGO = P.[ID]" _
       & " AND A.COD_PROVEEDOR = P.COD_PROVEEDOR" _
       & " where A.COD_PROVEEDOR = " & txtCodigo.Text & " and A.ANTICIPOS = '" & Item.Text & "'"
Call OpenRecordSet(rs, strSQL)
If glogon.error Then Exit Sub

lswPago.ListItems.Clear
Do While Not rs.EOF
  Set itmX = lswPago.ListItems.Add(, , rs!Anticipos)
      itmX.SubItems(1) = Format(rs!Registro_Fecha, "dd/mm/yyyy")
      itmX.SubItems(2) = Format(rs!Monto, "Standard")
      itmX.SubItems(3) = rs!cod_Factura
      itmX.SubItems(4) = rs!Npago
      itmX.SubItems(5) = rs!Registro_Usuario
      itmX.SubItems(6) = rs!cod_Divisa
      itmX.SubItems(7) = rs!TIPO_CAMBIO
      
      
  rs.MoveNext
Loop
rs.Close

Me.MousePointer = vbDefault
Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub opt_Click(Index As Integer)
Select Case True
  Case opt.Item(0).Value
     txtOrden.BackColor = txtAnticipo.BackColor
     txtOrden.Enabled = False
     txtOrden.Locked = True
     
     opt.Item(1).BackColor = opt.Item(0).BackColor
     opt.Item(1).ForeColor = opt.Item(0).ForeColor
     
     opt.Item(0).BackColor = vbWhite
     opt.Item(0).ForeColor = vbBlue
     
     
  Case opt.Item(1).Value
     txtOrden.BackColor = vbWhite
     txtOrden.Enabled = True
     txtOrden.Locked = True
     txtOrden.SetFocus

     opt.Item(0).BackColor = opt.Item(1).BackColor
     opt.Item(0).ForeColor = opt.Item(1).ForeColor
     
     opt.Item(1).BackColor = vbWhite
     opt.Item(1).ForeColor = vbBlue


End Select

End Sub

Private Sub sbTab_Load()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem

On Error GoTo vError

lsw.ListItems.Clear
lswPago.ListItems.Clear

If Len(txtCodigo) = 0 Then Exit Sub

If tcMain.Item(0).Selected Then
  cmdNuevo_Click
Else

  strSQL = "select A.*,P.tesoreria,P.fecha_vencimiento,C.descripcion as Cargo" _
         & ",dbo.fxCxP_CargoFlotanteSaldoCorte(A.cod_Proveedor,A.ID_Cargo, dbo.MyGetdate()) as 'Saldo'" _
         & " from cxp_anticipos A left join cxp_pagoProv P" _
         & " on A.cod_proveedor = P.cod_proveedor and A.Anticipos = P.cod_factura" _
         & " inner join CxP_Cargos C on A.cod_cargo = C.cod_cargo" _
         & " where A.cod_proveedor = " & txtCodigo & " order by Fecha desc"
  Call OpenRecordSet(rs, strSQL, 0)
  Do While Not rs.EOF
   Set itmX = lsw.ListItems.Add(, , rs!Anticipos)
       itmX.Tag = rs!IdX
       itmX.SubItems(1) = Format(rs!fecha, "dd/mm/yyyy")
       itmX.SubItems(2) = Format(rs!Monto, "Standard")
       itmX.SubItems(3) = Format(rs!Saldo, "Standard")
       itmX.SubItems(4) = rs!Usuario
       itmX.SubItems(5) = rs!Documento
       itmX.SubItems(6) = rs!cod_orden & ""
       itmX.SubItems(7) = IIf(IsNull(rs!Id_Cargo), "NO", "SI {" & rs!Id_Cargo & "}")
       itmX.SubItems(8) = rs!Notas
       itmX.SubItems(9) = rs!Cargo
       itmX.SubItems(10) = IIf(IsNull(rs!tesoreria), "NO", "SI {" & rs!tesoreria & "}")
       itmX.SubItems(11) = Format(rs!Fecha_Vencimiento, "dd/mm/yyyy")
       itmX.SubItems(12) = Format(rs!Fecha_Cobro_Anticipo & "", "dd/mm/yyyy")
   rs.MoveNext
  Loop
  rs.Close
End If

vError:

End Sub

Private Sub tcMain_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
Call sbTab_Load
End Sub

Private Sub TimerX_Timer()
TimerX.Interval = 0
Call cmdNuevo_Click
End Sub


Private Sub txtCargoCod_LostFocus()
txtCargoDesc = fxSIFCCodigos("D", txtCargoCod, "CargosProv")
End Sub

Private Sub txtCodigo_Change()
 cmdNuevo_Click
End Sub

Private Sub txtCodigo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtNombre.SetFocus

If KeyCode = vbKeyF4 Then
  gBusquedas.Convertir = "S"
  gBusquedas.Columna = "cod_proveedor"
  gBusquedas.Orden = "cod_proveedor"
  gBusquedas.Consulta = "select cod_proveedor,descripcion from cxp_proveedores"
  gBusquedas.Filtro = " and estado = 'A'"
  frmBusquedas.Show vbModal
  txtCodigo = gBusquedas.Resultado
  txtNombre = gBusquedas.Resultado2
End If
End Sub


Private Sub txtCodigo_LostFocus()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

If IsNumeric(txtCodigo.Text) Then
    strSQL = "select dbo.fxCntXTipoCambio(" & GLOBALES.gEnlace & ",COD_DIVISA,dbo.MyGetdate(),'V') as 'TipoCambio'" _
           & " From CxP_PRoveedores where cod_proveedor = " & txtCodigo.Text
    
    Call OpenRecordSet(rs, strSQL)
        vTipoCambio = rs!TipoCambio
    rs.Close
Else
  vTipoCambio = 1
End If

Exit Sub

vError:
  vTipoCambio = 1

End Sub

Private Sub txtDocumento_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtMonto.SetFocus
End Sub

Private Sub txtMonto_GotFocus()
On Error GoTo vError
txtMonto = CCur(txtMonto)
vError:
End Sub

Private Sub txtMonto_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then cmdGuardar.SetFocus
End Sub

Private Sub txtMonto_LostFocus()
On Error GoTo vError
txtMonto = Format(CCur(txtMonto), "Standard")
vError:
End Sub

Private Sub txtNombre_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtCargoCod.SetFocus
On Error GoTo vError
If KeyCode = vbKeyF4 Then
  gBusquedas.Convertir = "S"
  gBusquedas.Columna = "descripcion"
  gBusquedas.Orden = "descripcion"
  gBusquedas.Consulta = "select cod_proveedor,descripcion from cxp_proveedores"
  gBusquedas.Filtro = " and estado = 'A'"
  frmBusquedas.Show vbModal
  txtCodigo = gBusquedas.Resultado
  txtNombre = gBusquedas.Resultado2
  txtCargoCod.SetFocus
End If
vError:
End Sub

Private Sub txtCargoCod_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtCargoDesc.SetFocus

If KeyCode = vbKeyF4 Then
  gBusquedas.Convertir = "N"
  gBusquedas.Columna = "cod_cargo"
  gBusquedas.Orden = "cod_cargo"
  gBusquedas.Consulta = "select cod_cargo,descripcion from cxp_cargos"
  gBusquedas.Filtro = ""
  frmBusquedas.Show vbModal
  txtCargoCod = gBusquedas.Resultado
  txtCargoDesc = gBusquedas.Resultado2
End If
End Sub


Private Sub txtCargoDesc_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtDocumento.SetFocus

If KeyCode = vbKeyF4 Then
  gBusquedas.Convertir = "N"
  gBusquedas.Columna = "descripcion"
  gBusquedas.Orden = "descripcion"
  gBusquedas.Consulta = "select cod_cargo,descripcion from cxp_cargos"
  gBusquedas.Filtro = ""
  frmBusquedas.Show vbModal
  txtCargoCod = gBusquedas.Resultado
  txtCargoDesc = gBusquedas.Resultado2
End If

End Sub

Private Sub txtNotas_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then
  If opt.Item(0).Value Then
    txtMonto.SetFocus
  Else
    txtOrden.SetFocus
  End If
End If
End Sub

