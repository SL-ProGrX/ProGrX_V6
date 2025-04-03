VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#19.1#0"; "Codejock.Controls.v19.1.0.ocx"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Begin VB.Form frmSYS_APL_Control_Pagos 
   Caption         =   "APL: Control de Pagos (Triangulación)"
   ClientHeight    =   8304
   ClientLeft      =   48
   ClientTop       =   396
   ClientWidth     =   13320
   LinkTopic       =   "Form1"
   ScaleHeight     =   8304
   ScaleWidth      =   13320
   StartUpPosition =   3  'Windows Default
   Begin XtremeSuiteControls.TabControl tcMain 
      Height          =   6732
      Left            =   0
      TabIndex        =   0
      Top             =   1680
      Width           =   13092
      _Version        =   1245185
      _ExtentX        =   23093
      _ExtentY        =   11874
      _StockProps     =   68
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   4
      Color           =   32
      ItemCount       =   4
      Item(0).Caption =   "Pendientes de Formalizar"
      Item(0).ControlCount=   6
      Item(0).Control(0)=   "Label1(0)"
      Item(0).Control(1)=   "dtpInicio"
      Item(0).Control(2)=   "dtpCorte"
      Item(0).Control(3)=   "btnConsulta(0)"
      Item(0).Control(4)=   "btnConsulta(1)"
      Item(0).Control(5)=   "gPendientes"
      Item(1).Caption =   "Remesa de Cobros"
      Item(1).ControlCount=   7
      Item(1).Control(0)=   "Label1(2)"
      Item(1).Control(1)=   "btnCobros(0)"
      Item(1).Control(2)=   "btnCobros(1)"
      Item(1).Control(3)=   "btnCobros(2)"
      Item(1).Control(4)=   "gCobros"
      Item(1).Control(5)=   "dtpCobros_Inicio"
      Item(1).Control(6)=   "dtpCobros_Corte"
      Item(2).Caption =   "Cancelación"
      Item(2).ControlCount=   4
      Item(2).Control(0)=   "gbCancela"
      Item(2).Control(1)=   "Label1(4)"
      Item(2).Control(2)=   "cboRemesa"
      Item(2).Control(3)=   "lswCancela"
      Item(3).Caption =   "Pago a Pymes"
      Item(3).ControlCount=   5
      Item(3).Control(0)=   "lswPago_Facturas"
      Item(3).Control(1)=   "lswPago_Pyme"
      Item(3).Control(2)=   "Label1(10)"
      Item(3).Control(3)=   "Label1(11)"
      Item(3).Control(4)=   "btnCancela(1)"
      Begin XtremeSuiteControls.ListView lswCancela 
         Height          =   5652
         Left            =   -65440
         TabIndex        =   22
         Top             =   840
         Visible         =   0   'False
         Width           =   8532
         _Version        =   1245185
         _ExtentX        =   15049
         _ExtentY        =   9970
         _StockProps     =   77
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
         Checkboxes      =   -1  'True
         View            =   3
         FullRowSelect   =   -1  'True
         Appearance      =   16
      End
      Begin XtremeSuiteControls.ListView lswPago_Facturas 
         Height          =   5292
         Left            =   -64960
         TabIndex        =   34
         Top             =   720
         Visible         =   0   'False
         Width           =   7932
         _Version        =   1245185
         _ExtentX        =   13991
         _ExtentY        =   9334
         _StockProps     =   77
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
         View            =   3
         FullRowSelect   =   -1  'True
         Appearance      =   16
      End
      Begin XtremeSuiteControls.ListView lswPago_Pyme 
         Height          =   5292
         Left            =   -70000
         TabIndex        =   35
         Top             =   720
         Visible         =   0   'False
         Width           =   4932
         _Version        =   1245185
         _ExtentX        =   8700
         _ExtentY        =   9334
         _StockProps     =   77
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
         Checkboxes      =   -1  'True
         View            =   3
         FullRowSelect   =   -1  'True
         Appearance      =   16
      End
      Begin XtremeSuiteControls.GroupBox gbCancela 
         Height          =   6012
         Left            =   -69880
         TabIndex        =   17
         Top             =   480
         Visible         =   0   'False
         Width           =   4332
         _Version        =   1245185
         _ExtentX        =   7641
         _ExtentY        =   10604
         _StockProps     =   79
         Caption         =   "Datos de la Cancelación"
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
         Begin XtremeSuiteControls.FlatEdit FlatEdit2 
            Height          =   312
            Left            =   1920
            TabIndex        =   32
            Top             =   4680
            Width           =   2292
            _Version        =   1245185
            _ExtentX        =   4043
            _ExtentY        =   556
            _StockProps     =   77
            ForeColor       =   0
            BackColor       =   -2147483643
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
            Locked          =   -1  'True
            Appearance      =   2
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit FlatEdit1 
            Height          =   312
            Left            =   1920
            TabIndex        =   30
            Top             =   4320
            Width           =   2292
            _Version        =   1245185
            _ExtentX        =   4043
            _ExtentY        =   556
            _StockProps     =   77
            ForeColor       =   0
            BackColor       =   -2147483643
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
            Locked          =   -1  'True
            Appearance      =   2
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.DateTimePicker dtpCancela 
            Height          =   312
            Left            =   1920
            TabIndex        =   19
            Top             =   600
            Width           =   1332
            _Version        =   1245185
            _ExtentX        =   2350
            _ExtentY        =   550
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
         Begin XtremeSuiteControls.FlatEdit txtCancela_Monto 
            Height          =   312
            Left            =   1920
            TabIndex        =   26
            Top             =   1320
            Width           =   2292
            _Version        =   1245185
            _ExtentX        =   4043
            _ExtentY        =   556
            _StockProps     =   77
            ForeColor       =   0
            BackColor       =   -2147483643
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
            Appearance      =   2
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtCancela_Doc 
            Height          =   315
            Left            =   1920
            TabIndex        =   24
            Top             =   960
            Width           =   2292
            _Version        =   1245185
            _ExtentX        =   4043
            _ExtentY        =   556
            _StockProps     =   77
            ForeColor       =   0
            BackColor       =   -2147483643
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
            Appearance      =   2
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtCancela_Notas 
            Height          =   2232
            Left            =   240
            TabIndex        =   28
            Top             =   1920
            Width           =   3972
            _Version        =   1245185
            _ExtentX        =   7006
            _ExtentY        =   3937
            _StockProps     =   77
            ForeColor       =   0
            BackColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MultiLine       =   -1  'True
            ScrollBars      =   2
            Appearance      =   2
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.PushButton btnCancela 
            Height          =   432
            Index           =   0
            Left            =   1920
            TabIndex        =   33
            Top             =   5400
            Width           =   2292
            _Version        =   1245185
            _ExtentX        =   4043
            _ExtentY        =   762
            _StockProps     =   79
            Caption         =   "Aplicar Cancelación"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Appearance      =   6
         End
         Begin XtremeSuiteControls.Label Label1 
            Height          =   252
            Index           =   9
            Left            =   240
            TabIndex        =   31
            Top             =   4680
            Width           =   1812
            _Version        =   1245185
            _ExtentX        =   3196
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Diferencia : "
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Transparent     =   -1  'True
         End
         Begin XtremeSuiteControls.Label Label1 
            Height          =   252
            Index           =   8
            Left            =   240
            TabIndex        =   29
            Top             =   4320
            Width           =   1812
            _Version        =   1245185
            _ExtentX        =   3196
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Monto Detallado: "
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Transparent     =   -1  'True
         End
         Begin XtremeSuiteControls.Label Label1 
            Height          =   252
            Index           =   7
            Left            =   240
            TabIndex        =   27
            Top             =   1680
            Width           =   1812
            _Version        =   1245185
            _ExtentX        =   3196
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Notas : "
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Transparent     =   -1  'True
         End
         Begin XtremeSuiteControls.Label Label1 
            Height          =   252
            Index           =   6
            Left            =   240
            TabIndex        =   25
            Top             =   1320
            Width           =   1812
            _Version        =   1245185
            _ExtentX        =   3196
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Monto : "
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Transparent     =   -1  'True
         End
         Begin XtremeSuiteControls.Label Label1 
            Height          =   252
            Index           =   5
            Left            =   240
            TabIndex        =   23
            Top             =   960
            Width           =   1812
            _Version        =   1245185
            _ExtentX        =   3196
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Num. Documento : "
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Transparent     =   -1  'True
         End
         Begin XtremeSuiteControls.Label Label1 
            Height          =   252
            Index           =   3
            Left            =   240
            TabIndex        =   18
            Top             =   600
            Width           =   1812
            _Version        =   1245185
            _ExtentX        =   3196
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Fecha de Pago: "
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Transparent     =   -1  'True
         End
      End
      Begin XtremeSuiteControls.DateTimePicker dtpCobros_Corte 
         Height          =   312
         Left            =   -66520
         TabIndex        =   13
         Top             =   480
         Visible         =   0   'False
         Width           =   1332
         _Version        =   1245185
         _ExtentX        =   2350
         _ExtentY        =   550
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
      Begin XtremeSuiteControls.DateTimePicker dtpCobros_Inicio 
         Height          =   312
         Left            =   -67840
         TabIndex        =   12
         Top             =   480
         Visible         =   0   'False
         Width           =   1332
         _Version        =   1245185
         _ExtentX        =   2350
         _ExtentY        =   550
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
      Begin XtremeSuiteControls.PushButton btnConsulta 
         Height          =   312
         Index           =   0
         Left            =   4920
         TabIndex        =   5
         Top             =   480
         Width           =   492
         _Version        =   1245185
         _ExtentX        =   868
         _ExtentY        =   550
         _StockProps     =   79
         Caption         =   "..."
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   16
      End
      Begin FPSpreadADO.fpSpread gPendientes 
         Height          =   5532
         Left            =   120
         TabIndex        =   1
         Top             =   1080
         Width           =   12732
         _Version        =   524288
         _ExtentX        =   22458
         _ExtentY        =   9758
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
         MaxCols         =   502
         SpreadDesigner  =   "frmSYS_APL_Control_Pagos.frx":0000
         VScrollSpecial  =   -1  'True
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin XtremeSuiteControls.DateTimePicker dtpInicio 
         Height          =   312
         Left            =   2160
         TabIndex        =   3
         Top             =   480
         Width           =   1332
         _Version        =   1245185
         _ExtentX        =   2350
         _ExtentY        =   550
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
         Height          =   312
         Left            =   3480
         TabIndex        =   4
         Top             =   480
         Width           =   1332
         _Version        =   1245185
         _ExtentX        =   2350
         _ExtentY        =   550
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
      Begin XtremeSuiteControls.PushButton btnConsulta 
         Height          =   312
         Index           =   1
         Left            =   5400
         TabIndex        =   6
         Top             =   480
         Width           =   1092
         _Version        =   1245185
         _ExtentX        =   1926
         _ExtentY        =   550
         _StockProps     =   79
         Caption         =   "Exportar"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   16
      End
      Begin FPSpreadADO.fpSpread gCobros 
         Height          =   5532
         Left            =   -69880
         TabIndex        =   10
         Top             =   1080
         Visible         =   0   'False
         Width           =   12732
         _Version        =   524288
         _ExtentX        =   22458
         _ExtentY        =   9758
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
         MaxCols         =   19
         SpreadDesigner  =   "frmSYS_APL_Control_Pagos.frx":0A69
         VScrollSpecial  =   -1  'True
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin XtremeSuiteControls.PushButton btnCobros 
         Height          =   312
         Index           =   0
         Left            =   -65080
         TabIndex        =   14
         Top             =   480
         Visible         =   0   'False
         Width           =   492
         _Version        =   1245185
         _ExtentX        =   868
         _ExtentY        =   550
         _StockProps     =   79
         Caption         =   "..."
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   16
      End
      Begin XtremeSuiteControls.PushButton btnCobros 
         Height          =   312
         Index           =   1
         Left            =   -64600
         TabIndex        =   15
         Top             =   480
         Visible         =   0   'False
         Width           =   1092
         _Version        =   1245185
         _ExtentX        =   1926
         _ExtentY        =   550
         _StockProps     =   79
         Caption         =   "Exportar"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   16
      End
      Begin XtremeSuiteControls.PushButton btnCobros 
         Height          =   312
         Index           =   2
         Left            =   -63160
         TabIndex        =   16
         Top             =   480
         Visible         =   0   'False
         Width           =   2412
         _Version        =   1245185
         _ExtentX        =   4254
         _ExtentY        =   550
         _StockProps     =   79
         Caption         =   "Crear Remesa de Cobros"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   6
      End
      Begin XtremeSuiteControls.ComboBox cboRemesa 
         Height          =   312
         Left            =   -62320
         TabIndex        =   21
         Top             =   480
         Visible         =   0   'False
         Width           =   5412
         _Version        =   1245185
         _ExtentX        =   9546
         _ExtentY        =   550
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Style           =   2
         Appearance      =   2
         UseVisualStyle  =   0   'False
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.PushButton btnCancela 
         Height          =   432
         Index           =   1
         Left            =   -67360
         TabIndex        =   38
         Top             =   6120
         Visible         =   0   'False
         Width           =   2292
         _Version        =   1245185
         _ExtentX        =   4043
         _ExtentY        =   762
         _StockProps     =   79
         Caption         =   "Aplicar Pagos a Pymes"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   6
      End
      Begin XtremeSuiteControls.Label Label1 
         Height          =   252
         Index           =   11
         Left            =   -64960
         TabIndex        =   37
         Top             =   480
         Visible         =   0   'False
         Width           =   3972
         _Version        =   1245185
         _ExtentX        =   7006
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Facturas Pendientes de Pago: "
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
      End
      Begin XtremeSuiteControls.Label Label1 
         Height          =   252
         Index           =   10
         Left            =   -70000
         TabIndex        =   36
         Top             =   480
         Visible         =   0   'False
         Width           =   852
         _Version        =   1245185
         _ExtentX        =   1503
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Pyme: "
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
      End
      Begin XtremeSuiteControls.Label Label1 
         Height          =   252
         Index           =   4
         Left            =   -63280
         TabIndex        =   20
         Top             =   480
         Visible         =   0   'False
         Width           =   852
         _Version        =   1245185
         _ExtentX        =   1503
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Remesa: "
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
      End
      Begin XtremeSuiteControls.Label Label1 
         Height          =   252
         Index           =   2
         Left            =   -69640
         TabIndex        =   11
         Top             =   480
         Visible         =   0   'False
         Width           =   1812
         _Version        =   1245185
         _ExtentX        =   3196
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Formalizadas: "
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
      End
      Begin XtremeSuiteControls.Label Label1 
         Height          =   252
         Index           =   0
         Left            =   360
         TabIndex        =   2
         Top             =   480
         Width           =   1812
         _Version        =   1245185
         _ExtentX        =   3196
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Fechas de Registro: "
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
      End
   End
   Begin XtremeSuiteControls.ComboBox cboDominio 
      Height          =   312
      Left            =   1200
      TabIndex        =   9
      Top             =   1080
      Width           =   5412
      _Version        =   1245185
      _ExtentX        =   9546
      _ExtentY        =   550
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Style           =   2
      Appearance      =   2
      UseVisualStyle  =   0   'False
      Text            =   "ComboBox1"
   End
   Begin XtremeSuiteControls.Label Label1 
      Height          =   252
      Index           =   1
      Left            =   240
      TabIndex        =   8
      Top             =   1080
      Width           =   852
      _Version        =   1245185
      _ExtentX        =   1503
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Dominio: "
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
   End
   Begin XtremeSuiteControls.Label Label2 
      Height          =   612
      Left            =   2640
      TabIndex        =   7
      Top             =   240
      Width           =   6732
      _Version        =   1245185
      _ExtentX        =   11874
      _ExtentY        =   1080
      _StockProps     =   79
      Caption         =   "Control de Cobros  (Clientes) + Pagos a Pymes"
      ForeColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   16.2
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Transparent     =   -1  'True
   End
   Begin VB.Image imgBanner 
      Height          =   972
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   14412
   End
End
Attribute VB_Name = "frmSYS_APL_Control_Pagos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vPaso As Boolean

Private Sub Form_Load()
Dim strSQL As String

vModulo = 38

Me.imgBanner.Picture = frmContenedor.imgBanner_01.Picture

With lswCancela.ColumnHeaders
    .Clear
    .Add , , "Id Solicitud", 1200
    .Add , , "No. Factura", 1800
    .Add , , "Monto", 1200, vbRightJustify
    .Add , , "Pyme", 1800
    .Add , , "Documento", 1200, vbCenter
    .Add , , "Formalizada", 1800, vbCenter
End With


With lswPago_Facturas.ColumnHeaders
    .Clear
    .Add , , "Id Solicitud", 1200
    .Add , , "No. Factura", 1800
    .Add , , "Monto", 1200, vbRightJustify
    .Add , , "Documento", 1200, vbCenter
    .Add , , "Formalizada", 1800, vbCenter
End With

With lswPago_Pyme.ColumnHeaders
    .Clear
    .Add , , "Pyme", 2100
    .Add , , "Monto", 1400, vbRightJustify
End With

dtpInicio.Value = fxFechaServidor
dtpCorte.Value = dtpInicio.Value
dtpCobros_Inicio.Value = dtpInicio.Value
dtpCobros_Corte.Value = dtpInicio.Value
dtpCancela.Value = dtpInicio.Value

gPendientes.MaxRows = 1
gCobros.MaxRows = 1



vPaso = True

strSQL = "exec spAPL_Dominios_Vinculados '" & gAPL.APL_Dominio & "','V'"
Call sbCbo_Llena_New(cboDominio, strSQL, False, True)

vPaso = False
tcMain.Item(0).Selected = True

End Sub
