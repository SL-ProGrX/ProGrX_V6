VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#20.2#0"; "Codejock.Controls.v20.2.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#20.2#0"; "Codejock.ShortcutBar.v20.2.0.ocx"
Begin VB.Form frmFndReservas 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Fondos: Reservas"
   ClientHeight    =   8445
   ClientLeft      =   30
   ClientTop       =   390
   ClientWidth     =   11070
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8445
   ScaleWidth      =   11070
   Begin VB.Timer TimerX 
      Interval        =   5
      Left            =   9720
      Top             =   240
   End
   Begin XtremeSuiteControls.TabControl tcMain 
      Height          =   6975
      Left            =   120
      TabIndex        =   1
      Top             =   1320
      Width           =   10815
      _Version        =   1310722
      _ExtentX        =   19076
      _ExtentY        =   12303
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
      Item(0).Caption =   "Reservas"
      Item(0).ControlCount=   9
      Item(0).Control(0)=   "vGrid"
      Item(0).Control(1)=   "lswCtas"
      Item(0).Control(2)=   "txtCuentaCod"
      Item(0).Control(3)=   "txtCuentaDesc"
      Item(0).Control(4)=   "btnBarra(0)"
      Item(0).Control(5)=   "btnBarra(2)"
      Item(0).Control(6)=   "btnBarra(1)"
      Item(0).Control(7)=   "scBarra"
      Item(0).Control(8)=   "Label3(3)"
      Item(1).Caption =   "Contenido"
      Item(1).ControlCount=   4
      Item(1).Control(0)=   "cbo"
      Item(1).Control(1)=   "Label2(0)"
      Item(1).Control(2)=   "lsw"
      Item(1).Control(3)=   "gbGarantias"
      Item(2).Caption =   "Consultas"
      Item(2).ControlCount=   12
      Item(2).Control(0)=   "Label2(1)"
      Item(2).Control(1)=   "lswConsulta"
      Item(2).Control(2)=   "Label2(7)"
      Item(2).Control(3)=   "btnConsulta"
      Item(2).Control(4)=   "dtpInicio"
      Item(2).Control(5)=   "dtpCorte"
      Item(2).Control(6)=   "txtReserva"
      Item(2).Control(7)=   "cboConsulta"
      Item(2).Control(8)=   "lswDetalle"
      Item(2).Control(9)=   "lblSaldoContable"
      Item(2).Control(10)=   "btnExportar(0)"
      Item(2).Control(11)=   "btnExportar(1)"
      Begin XtremeSuiteControls.ListView lswCtas 
         Height          =   1815
         Left            =   360
         TabIndex        =   30
         Top             =   4920
         Width           =   10095
         _Version        =   1310722
         _ExtentX        =   17806
         _ExtentY        =   3201
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
         ShowBorder      =   0   'False
      End
      Begin XtremeSuiteControls.ListView lswDetalle 
         Height          =   2655
         Left            =   -69880
         TabIndex        =   27
         Top             =   4200
         Visible         =   0   'False
         Width           =   10170
         _Version        =   1310722
         _ExtentX        =   17939
         _ExtentY        =   4683
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
      Begin XtremeSuiteControls.ListView lswConsulta 
         Height          =   2895
         Left            =   -69880
         TabIndex        =   18
         Top             =   1200
         Visible         =   0   'False
         Width           =   10170
         _Version        =   1310722
         _ExtentX        =   17939
         _ExtentY        =   5106
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
      Begin XtremeSuiteControls.ListView lsw 
         Height          =   3615
         Left            =   -69880
         TabIndex        =   2
         Top             =   1200
         Visible         =   0   'False
         Width           =   10530
         _Version        =   1310722
         _ExtentX        =   18574
         _ExtentY        =   6376
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
      Begin XtremeSuiteControls.PushButton btnConsulta 
         Height          =   372
         Left            =   -61120
         TabIndex        =   22
         Top             =   720
         Visible         =   0   'False
         Width           =   972
         _Version        =   1310722
         _ExtentX        =   1714
         _ExtentY        =   656
         _StockProps     =   79
         Caption         =   "Consultar"
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   16
      End
      Begin FPSpreadADO.fpSpread vGrid 
         Height          =   3255
         Left            =   120
         TabIndex        =   3
         Top             =   480
         Width           =   10695
         _Version        =   524288
         _ExtentX        =   18865
         _ExtentY        =   5741
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
         MaxCols         =   495
         ScrollBars      =   2
         SpreadDesigner  =   "frmFndReservas.frx":0000
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin XtremeSuiteControls.GroupBox gbGarantias 
         Height          =   1935
         Left            =   -69160
         TabIndex        =   4
         Top             =   4920
         Visible         =   0   'False
         Width           =   8415
         _Version        =   1310722
         _ExtentX        =   14838
         _ExtentY        =   3408
         _StockProps     =   79
         Caption         =   "Registro"
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
         Begin XtremeSuiteControls.CheckBox chkPatrimonio 
            Height          =   312
            Left            =   6600
            TabIndex        =   5
            Top             =   1080
            Width           =   1692
            _Version        =   1310722
            _ExtentX        =   2984
            _ExtentY        =   556
            _StockProps     =   79
            Caption         =   "Patrimonio?   "
            BackColor       =   -2147483633
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   7.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Enabled         =   0   'False
            TextAlignment   =   1
            Appearance      =   16
            Alignment       =   1
         End
         Begin XtremeSuiteControls.PushButton btnMov 
            Height          =   312
            Index           =   0
            Left            =   2760
            TabIndex        =   6
            Top             =   1200
            Width           =   372
            _Version        =   1310722
            _ExtentX        =   656
            _ExtentY        =   556
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
            FlatStyle       =   -1  'True
            Appearance      =   16
            Picture         =   "frmFndReservas.frx":06D4
         End
         Begin XtremeSuiteControls.ComboBox cboOperadora 
            Height          =   312
            Left            =   1680
            TabIndex        =   7
            Top             =   360
            Width           =   6612
            _Version        =   1310722
            _ExtentX        =   11668
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
            Appearance      =   2
            UseVisualStyle  =   0   'False
            Text            =   "ComboBox1"
         End
         Begin XtremeSuiteControls.FlatEdit txtPlan 
            Height          =   312
            Left            =   1680
            TabIndex        =   8
            Top             =   720
            Width           =   972
            _Version        =   1310722
            _ExtentX        =   1714
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
            Appearance      =   2
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtPorcentaje 
            Height          =   312
            Left            =   1680
            TabIndex        =   9
            Top             =   1200
            Width           =   972
            _Version        =   1310722
            _ExtentX        =   1714
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
            Appearance      =   2
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtPlanDesc 
            Height          =   312
            Left            =   2640
            TabIndex        =   10
            Top             =   720
            Width           =   5652
            _Version        =   1310722
            _ExtentX        =   9970
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
            Appearance      =   2
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.PushButton btnMov 
            Height          =   312
            Index           =   1
            Left            =   3120
            TabIndex        =   11
            Top             =   1200
            Width           =   372
            _Version        =   1310722
            _ExtentX        =   656
            _ExtentY        =   556
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
            FlatStyle       =   -1  'True
            Appearance      =   16
            Picture         =   "frmFndReservas.frx":0DF4
         End
         Begin XtremeSuiteControls.FlatEdit txtLinea 
            Height          =   312
            Left            =   5520
            TabIndex        =   12
            Top             =   1200
            Visible         =   0   'False
            Width           =   972
            _Version        =   1310722
            _ExtentX        =   1714
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
            Appearance      =   2
            UseVisualStyle  =   0   'False
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Plan"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Index           =   2
            Left            =   600
            TabIndex        =   15
            Top             =   720
            Width           =   852
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Porcentaje"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   372
            Index           =   3
            Left            =   600
            TabIndex        =   14
            Top             =   1200
            Width           =   972
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Operadora"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Index           =   4
            Left            =   600
            TabIndex        =   13
            Top             =   360
            Width           =   972
         End
      End
      Begin XtremeSuiteControls.ComboBox cbo 
         Height          =   312
         Left            =   -69880
         TabIndex        =   16
         Top             =   720
         Visible         =   0   'False
         Width           =   4572
         _Version        =   1310722
         _ExtentX        =   8070
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
         Appearance      =   2
         UseVisualStyle  =   0   'False
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.ComboBox cboConsulta 
         Height          =   312
         Left            =   -69880
         TabIndex        =   19
         Top             =   720
         Visible         =   0   'False
         Width           =   4092
         _Version        =   1310722
         _ExtentX        =   7223
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
         Appearance      =   2
         UseVisualStyle  =   0   'False
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.DateTimePicker dtpInicio 
         Height          =   312
         Left            =   -65800
         TabIndex        =   23
         Top             =   720
         Visible         =   0   'False
         Width           =   1332
         _Version        =   1310722
         _ExtentX        =   2350
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
      Begin XtremeSuiteControls.DateTimePicker dtpCorte 
         Height          =   312
         Left            =   -64480
         TabIndex        =   24
         Top             =   720
         Visible         =   0   'False
         Width           =   1332
         _Version        =   1310722
         _ExtentX        =   2350
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
      Begin XtremeSuiteControls.FlatEdit txtReserva 
         Height          =   312
         Left            =   -63160
         TabIndex        =   25
         Top             =   720
         Visible         =   0   'False
         Width           =   1932
         _Version        =   1310722
         _ExtentX        =   3408
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
         Text            =   "0"
         Alignment       =   1
         Locked          =   -1  'True
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.PushButton btnExportar 
         Height          =   375
         Index           =   0
         Left            =   -59680
         TabIndex        =   28
         ToolTipText     =   "Exportar a Excel"
         Top             =   1200
         Visible         =   0   'False
         Width           =   375
         _Version        =   1310722
         _ExtentX        =   661
         _ExtentY        =   661
         _StockProps     =   79
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   16
         Picture         =   "frmFndReservas.frx":1398
      End
      Begin XtremeSuiteControls.PushButton btnExportar 
         Height          =   375
         Index           =   1
         Left            =   -59680
         TabIndex        =   29
         ToolTipText     =   "Exportar a Excel"
         Top             =   4200
         Visible         =   0   'False
         Width           =   375
         _Version        =   1310722
         _ExtentX        =   661
         _ExtentY        =   661
         _StockProps     =   79
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   16
         Picture         =   "frmFndReservas.frx":1C69
      End
      Begin XtremeSuiteControls.FlatEdit txtCuentaCod 
         Height          =   315
         Left            =   1320
         TabIndex        =   31
         Top             =   4440
         Width           =   2175
         _Version        =   1310722
         _ExtentX        =   3831
         _ExtentY        =   556
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
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtCuentaDesc 
         Height          =   315
         Left            =   3480
         TabIndex        =   32
         Top             =   4440
         Width           =   6975
         _Version        =   1310722
         _ExtentX        =   12303
         _ExtentY        =   556
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
         Locked          =   -1  'True
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.PushButton btnBarra 
         Height          =   315
         Index           =   0
         Left            =   8640
         TabIndex        =   33
         ToolTipText     =   "Nuevo"
         Top             =   3960
         Width           =   1095
         _Version        =   1310722
         _ExtentX        =   1926
         _ExtentY        =   550
         _StockProps     =   79
         Caption         =   "Nuevo"
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   6
         Picture         =   "frmFndReservas.frx":253A
         ImageAlignment  =   4
      End
      Begin XtremeSuiteControls.PushButton btnBarra 
         Height          =   315
         Index           =   2
         Left            =   10080
         TabIndex        =   34
         ToolTipText     =   "Eliminar"
         Top             =   3960
         Width           =   375
         _Version        =   1310722
         _ExtentX        =   656
         _ExtentY        =   550
         _StockProps     =   79
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   6
         Picture         =   "frmFndReservas.frx":2B6C
         ImageAlignment  =   6
      End
      Begin XtremeSuiteControls.PushButton btnBarra 
         Height          =   315
         Index           =   1
         Left            =   9720
         TabIndex        =   35
         ToolTipText     =   "Guardar"
         Top             =   3960
         Width           =   375
         _Version        =   1310722
         _ExtentX        =   656
         _ExtentY        =   550
         _StockProps     =   79
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   6
         Picture         =   "frmFndReservas.frx":3110
         ImageAlignment  =   6
      End
      Begin XtremeSuiteControls.Label Label3 
         Height          =   255
         Index           =   3
         Left            =   360
         TabIndex        =   37
         Top             =   4440
         Width           =   1215
         _Version        =   1310722
         _ExtentX        =   2138
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Cuenta"
         BackColor       =   -2147483633
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
      Begin XtremeShortcutBar.ShortcutCaption scBarra 
         Height          =   495
         Left            =   120
         TabIndex        =   36
         Top             =   3840
         Width           =   10455
         _Version        =   1310722
         _ExtentX        =   18441
         _ExtentY        =   873
         _StockProps     =   14
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SubItemCaption  =   -1  'True
         VisualTheme     =   3
      End
      Begin VB.Label lblSaldoContable 
         BackStyle       =   0  'Transparent
         Caption         =   "Contabilidad:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   -63160
         TabIndex        =   26
         ToolTipText     =   "Saldo Contable"
         Top             =   480
         Visible         =   0   'False
         Width           =   3012
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Rango de Fechas:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Index           =   7
         Left            =   -65800
         TabIndex        =   21
         Top             =   480
         Visible         =   0   'False
         Width           =   1932
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Reserva:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Index           =   1
         Left            =   -69880
         TabIndex        =   20
         Top             =   480
         Visible         =   0   'False
         Width           =   1212
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Reserva:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Index           =   0
         Left            =   -69880
         TabIndex        =   17
         Top             =   480
         Visible         =   0   'False
         Width           =   1212
      End
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Definición y Monitoreo de Reservas"
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
      Height          =   492
      Left            =   1680
      TabIndex        =   0
      Top             =   360
      Width           =   7932
   End
   Begin VB.Image imgBanner 
      Height          =   1215
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   11295
   End
End
Attribute VB_Name = "frmFndReservas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vPaso As Boolean


Private Sub sbReserva_Corte(pReserva As String, pInicio As String, pCorte As String _
                    , Optional pTipo As String = "R")

Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem

On Error GoTo vError

Me.MousePointer = vbHourglass

vPaso = True

If pTipo = "R" Then
    lswConsulta.ListItems.Clear
    lswDetalle.ListItems.Clear
    
    lblSaldoContable.Caption = "Contabilidad:"
    txtReserva.Text = Format(0, "Standard")
    
    strSQL = "exec spFnd_Reserva_Cortes '" & pReserva & "','" & pInicio & "','" _
            & pCorte & "','" & pTipo & "'"
    Call OpenRecordSet(rs, strSQL)
    Do While Not rs.EOF
     Set itmX = lswConsulta.ListItems.Add(, , rs!Corte)
         itmX.SubItems(1) = rs!cod_reserva
         itmX.SubItems(2) = Format(rs!Base, "Standard")
         itmX.SubItems(3) = Format(rs!Porcentaje, "Standard")
         itmX.SubItems(4) = Format(rs!Monto_Reserva, "Standard")
         itmX.SubItems(5) = Format(rs!SALDO_CONTABLE, "Standard")
         itmX.SubItems(6) = Format(rs!SALDO_CONTABLE - rs!Monto_Reserva, "Standard")
         
     rs.MoveNext
    Loop
    rs.Close

Else

    
    lswDetalle.ListItems.Clear
    strSQL = "exec spFnd_Reserva_Cortes '" & pReserva & "','" & pInicio & "','" _
            & pCorte & "','" & pTipo & "'"
    Call OpenRecordSet(rs, strSQL)
    Do While Not rs.EOF
     Set itmX = lswDetalle.ListItems.Add(, , rs!Corte)
         itmX.SubItems(1) = rs!cod_reserva
         itmX.SubItems(2) = rs!cod_Plan
         itmX.SubItems(3) = rs!Plan_Desc
         itmX.SubItems(4) = Format(rs!Base, "Standard")
         itmX.SubItems(5) = Format(rs!Porcentaje, "Standard")
         itmX.SubItems(6) = Format(rs!Monto_Reserva, "Standard")
     rs.MoveNext
    Loop
    rs.Close
    
    lblSaldoContable.Caption = "Contabilidad: " & Format(pInicio, "dd/MM/yyyy")
    strSQL = "exec spFnd_Reserva_Cuenta_Saldo  '" & pReserva & "','" & Format(pInicio, "yyyy/MM/dd") & "'"
    
    Call OpenRecordSet(rs, strSQL)
    If Not glogon.error Then
        txtReserva.Text = Format(rs!SALDO_CONTABLE, "Standard")
    Else
        txtReserva.Text = Format(0, "Standard")
    End If
    rs.Close

End If

vPaso = False

Me.MousePointer = vbDefault
Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub btnBarra_Click(Index As Integer)
Dim strSQL As String

If scBarra.Tag = "" Then Exit Sub

Select Case Index
 Case 0 'Nuevo
    txtCuentaCod.Tag = ""
    txtCuentaCod.Text = ""
    txtCuentaDesc.Text = ""
    
 Case 1 'Actualiza
    If txtCuentaCod.Text = "" Then
      MsgBox "Selecciones Primero la cuenta que desea eliminar...", vbExclamation
    Else
      strSQL = "exec spFnd_Reserva_Cuentas_Registro '" & scBarra.Tag & "','" & txtCuentaCod.Text & "','" & glogon.Usuario & "','A'"
      
      Call ConectionExecute(strSQL, 0)
      
      Call sbCuentas(scBarra.Tag)
    End If
  
 Case 2 'Borrar
    If txtCuentaCod.Text = "" Then
      MsgBox "Selecciones Primero la cuenta que desea eliminar...", vbExclamation
    Else
      strSQL = "exec spFnd_Reserva_Cuentas_Registro '" & scBarra.Tag & "','" & txtCuentaCod.Text & "','" & glogon.Usuario & "','E'"
      
      Call ConectionExecute(strSQL, 0)
      
      Call sbCuentas(scBarra.Tag)
    End If
End Select
End Sub

Private Sub btnConsulta_Click()
If cboConsulta.ListCount = 0 Then Exit Sub

Call sbReserva_Corte(cboConsulta.ItemData(cboConsulta.ListIndex), Format(dtpInicio.Value, "yyyy/mm/dd") _
        , Format(dtpCorte.Value, "yyyy/mm/dd"), "R")

End Sub

Private Sub btnExportar_Click(Index As Integer)
On Error GoTo vError

Me.MousePointer = vbHourglass

Select Case Index
Case 0
    Call Excel_Exportar_Lsw(lswConsulta)
Case 1
    Call Excel_Exportar_Lsw(lswDetalle)
End Select


Me.MousePointer = vbDefault

Exit Sub

vError:
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub btnMov_Click(Index As Integer)
Dim strSQL As String, strBitacora As String
Dim pOperadora As Integer, pPlan As String, pReserva As String, pEstado As String

On Error GoTo vError


If Not IsNumeric(txtPorcentaje.Text) Then
   MsgBox "Porcentaje no es válido", vbExclamation
   Exit Sub
Else
  If CCur(txtPorcentaje.Text) < 0 Or CCur(txtPorcentaje.Text) > 999 Then
     MsgBox "Porcentaje no es válido", vbExclamation
     Exit Sub
  End If
End If

pOperadora = cboOperadora.ItemData(cboOperadora.ListIndex)
pPlan = txtPlan.Text
pReserva = cbo.ItemData(cbo.ListIndex)

strBitacora = "Reserva de Fondos, Linea: " & txtLinea.Text & ", Reserva: " & pReserva _
         & ", Plan: " & pPlan & " Porcentaje: " & CCur(txtPorcentaje.Text)

strSQL = "exec spFnd_Reserva_Contenido_Registro '" & pReserva & "'," & txtLinea & "," & chkPatrimonio.Value _
       & "," & pOperadora & ",'" & pPlan & "'," & CCur(txtPorcentaje.Text) & ",'" & glogon.Usuario & "'"

Select Case Index
 Case 0 'Agregar / Modificar
       
    strSQL = strSQL & ",'A'"
    Call ConectionExecute(strSQL)
       
    Call Bitacora("Registra", strBitacora)
    
 
 Case 1 'Elimina

    strSQL = strSQL & ",'E'"
    Call ConectionExecute(strSQL)
    
    Call Bitacora("Elimina", strBitacora)
    
End Select

Call sbConsulta

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub cbo_Click()

If vPaso Then Exit Sub

If cbo.ListCount = 0 Then Exit Sub
If cbo.Text = "" Then Exit Sub

Call sbConsulta

End Sub


Private Sub sbConsulta()
Dim strSQL As String, rs  As New ADODB.Recordset
Dim vGarantia As String, vEstado As String
Dim itmX As ListViewItem

If vPaso Then Exit Sub

On Error GoTo vError

Me.MousePointer = vbHourglass

lsw.ListItems.Clear

If cbo.ListCount = 0 Then Exit Sub
If cbo.Text = "" Then Exit Sub

txtPlan.Text = ""
txtPlanDesc.Text = ""
txtPorcentaje.Text = Format(0, "Standard")
txtLinea.Text = "0"
chkPatrimonio.Value = xtpUnchecked


strSQL = "exec spFnd_Reserva_Contenido_Consulta '" & cbo.ItemData(cbo.ListIndex) & "'"
Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
  Set itmX = lsw.ListItems.Add(, , rs!Linea_Id)
      itmX.SubItems(1) = rs!cod_Operadora
      itmX.SubItems(2) = rs!cod_Plan
      itmX.SubItems(3) = rs!Descripcion
      itmX.SubItems(4) = Format(rs!Porcentaje, "Standard")
      itmX.SubItems(5) = IIf(rs!Patrimonio = 1, "Sí", "No")

  rs.MoveNext
Loop
rs.Close

Me.MousePointer = vbDefault

Exit Sub


vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
  
End Sub



Private Sub cboConsulta_Click()
If vPaso Then Exit Sub
If cboConsulta.ListCount = 0 Then Exit Sub
Call btnConsulta_Click
End Sub

Private Sub Form_Activate()
vModulo = 18

End Sub

Private Sub lsw_ColumnClick(ByVal ColumnHeader As XtremeSuiteControls.ListViewColumnHeader)
 lsw.SortKey = ColumnHeader.Index - 1
  If lsw.SortOrder = 0 Then lsw.SortOrder = 1 Else lsw.SortOrder = 0
  lsw.Sorted = True
End Sub

Private Sub lsw_ItemClick(ByVal Item As XtremeSuiteControls.ListViewItem)
If vPaso Then Exit Sub

txtLinea.Text = Item.Text
txtPlan.Text = Item.SubItems(2)
txtPlanDesc.Text = Item.SubItems(3)
txtPorcentaje.Text = Item.SubItems(4)

chkPatrimonio.Value = IIf(Mid(Item.SubItems(5), 1, 1) = "S", xtpChecked, xtpUnchecked)
End Sub


Private Sub lswConsulta_ItemClick(ByVal Item As XtremeSuiteControls.ListViewItem)

If vPaso Then Exit Sub

Call sbReserva_Corte(Item.SubItems(1), Item.Text _
        , Item.Text, "D")
End Sub



Private Sub lswCtas_ItemClick(ByVal Item As XtremeSuiteControls.ListViewItem)
If lswCtas.ListItems.Count > 0 Then
 txtCuentaCod.Text = Item.Text
 txtCuentaDesc.Text = Item.SubItems(1)
End If
End Sub


Private Sub tcMain_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
Dim strSQL As String


Select Case Item.Index

    Case 1 'Contenido
        vPaso = True
    
        strSQL = "select COD_RESERVA as 'Idx',Descripcion as 'ItmX'" _
               & " from FND_RESERVAS"
        
        Call sbCbo_Llena_New(cbo, strSQL, False, True)
        vPaso = False
    
        Call cbo_Click
    Case 2 'Consultas

        vPaso = True
    
        strSQL = "select COD_RESERVA as 'Idx',Descripcion as 'ItmX'" _
               & " from FND_RESERVAS"
        
        Call sbCbo_Llena_New(cboConsulta, strSQL, False, True)
        vPaso = False
    
End Select


End Sub

Private Sub TimerX_Timer()
TimerX.Interval = 0
TimerX.Enabled = False

Dim strSQL As String

vPaso = True

tcMain.Item(0).Selected = True

strSQL = "select *" _
       & " from vFND_RESERVAS"
Call sbCargaGridLocal(vGrid, 6, strSQL)

strSQL = "select rtrim(cod_Operadora) as 'IdX', rtrim(descripcion) as ItmX" _
         & " from  fnd_Operadoras"
Call sbCbo_Llena_New(cboOperadora, strSQL, False, True)

vPaso = False

End Sub


Private Sub sbCargaGridLocal(vGrid As Object, vGridMaxCol As Integer, strSQL As String)
Dim rs As New ADODB.Recordset, i As Integer, strResultado As String

Me.MousePointer = vbHourglass

vGrid.MaxCols = vGridMaxCol
vGrid.MaxRows = 1

vGrid.Row = vGrid.MaxRows
Call OpenRecordSet(rs, strSQL, 0)

vPaso = True
Do While Not rs.EOF
  vGrid.Row = vGrid.MaxRows
  
  For i = 1 To vGrid.MaxCols
    vGrid.Col = i
    Select Case i
     Case 1
        vGrid.Text = rs!cod_reserva
     Case 2
        vGrid.Text = rs!Descripcion
     Case 3
        vGrid.Text = rs!Cta_Reserva & ""
        vGrid.CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
        vGrid.CellNote = rs!Cta_Reserva_Desc & ""
        vGrid.TextTip = TextTipFixed
     Case 4
        vGrid.Text = rs!Cta_Transitoria & ""
        vGrid.CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
        vGrid.CellNote = rs!Cta_Transitoria_Desc & ""
        vGrid.TextTip = TextTipFixed
     Case 5
        vGrid.Value = rs!Activa
    End Select
  
  Next i
  
  vGrid.MaxRows = vGrid.MaxRows + 1
  
  rs.MoveNext

Loop

rs.Close

vPaso = False

Me.MousePointer = vbDefault

End Sub



Private Sub Form_Load()

vModulo = 18

vGrid.AppearanceStyle = fxGridStyle
Set imgBanner.Picture = frmContenedor.imgBanner_Mantenimiento.Picture

tcMain.Item(0).Selected = True

With lswCtas.ColumnHeaders
    .Clear
    .Add , , "Cuenta", 3200
    .Add , , "Descripción", lswCtas.Width - (3400)
End With


With lsw.ColumnHeaders
    .Clear
    .Add , , "[Id]", 600
    .Add , , "[Op]", 600, vbCenter
    .Add , , "Plan", 900, vbCenter
    .Add , , "Descripción", 3000
    .Add , , "Porcentaje", 1200, vbRightJustify
    .Add , , "Patrimonio", 1100, vbCenter
End With

With lswConsulta.ColumnHeaders
    .Clear
    .Add , , "Corte", 1800, vbCenter
    .Add , , "Reserva", 1100, vbCenter
    .Add , , "Base", 2100, vbRightJustify
    .Add , , "Porcentaje", 1100, vbRightJustify
    .Add , , "Monto", 2100, vbRightJustify
    .Add , , "Contabilidad", 2100, vbRightJustify
    .Add , , "Pendiente", 2100, vbRightJustify
End With

With lswDetalle.ColumnHeaders
    .Clear
    .Add , , "Corte", 1800, vbCenter
    .Add , , "Reserva", 1100, vbCenter
    .Add , , "Plan", 900, vbCenter
    .Add , , "Descripción", 3000
    .Add , , "Base", 2100, vbRightJustify
    .Add , , "Porcentaje", 1100, vbRightJustify
    .Add , , "Monto", 2100, vbRightJustify
End With


dtpCorte.Value = Date
dtpInicio.Value = DateAdd("d", (Day(dtpCorte.Value) - 1) * -1, dtpCorte.Value)

Call Formularios(Me)
Call RefrescaTags(Me)

btnMov(0).Enabled = vGrid.Enabled
btnMov(1).Enabled = vGrid.Enabled

End Sub


Private Function fxValida() As Boolean
Dim strSQL As String, rs As New ADODB.Recordset
Dim vMensaje As String

fxValida = True

vMensaje = ""


If Len(vMensaje) > 0 Then
   fxValida = False
   MsgBox vMensaje, vbExclamation
End If

End Function

Private Function fxGuardar() As Long
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

If Not fxValida Then
   fxGuardar = 0
   Exit Function
End If

vGrid.Row = vGrid.ActiveRow
vGrid.Col = 1

strSQL = "select isnull(count(*),0) as Existe from FND_RESERVAS where COD_RESERVA = '" & vGrid.Text & "'"
Call OpenRecordSet(rs, strSQL)

If rs!Existe = 0 Then   'Insertar
  
  vGrid.Col = 1
  strSQL = "insert FND_RESERVAS(COD_RESERVA,descripcion,cod_cuenta, cod_cuenta_Tra,activa) values('" & vGrid.Text & "','"
  vGrid.Col = 2
  strSQL = strSQL & vGrid.Text & "','"
  vGrid.Col = 3
  strSQL = strSQL & fxgCntCuentaFormato(False, vGrid.Text, 0) & "','"
  vGrid.Col = 4
  strSQL = strSQL & fxgCntCuentaFormato(False, vGrid.Text, 0) & "',"
  vGrid.Col = 5
  strSQL = strSQL & vGrid.Value & ")"
  Call ConectionExecute(strSQL)
  
  
  vGrid.Col = 1
  
  Call Bitacora("Registra", "Reserva de Fondos: " & vGrid.Text)
  
  fxGuardar = 1
  
Else 'Actualizar

    vGrid.Col = 2
    strSQL = "update FND_RESERVAS set descripcion = '" & vGrid.Text & "', cod_cuenta = '"
    vGrid.Col = 3
    strSQL = strSQL & fxgCntCuentaFormato(False, vGrid.Text, 0) & "', cod_cuenta_tra = '"
    vGrid.Col = 4
    strSQL = strSQL & fxgCntCuentaFormato(False, vGrid.Text, 0) & "', Activa = "
    vGrid.Col = 5
    strSQL = strSQL & vGrid.Value & " where COD_RESERVA = '"
    vGrid.Col = 1
    strSQL = strSQL & vGrid.Text & "'"
    
    Call ConectionExecute(strSQL)
    
    fxGuardar = 1
    
    Call Bitacora("Modifica", "Reserva de Fondos: " & vGrid.Text)
 
End If

Exit Function

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical


End Function

Private Sub sbCuentas(pReserva As String)
Dim strSQL As String, rs  As New ADODB.Recordset
Dim vGarantia As String, vEstado As String
Dim itmX As ListViewItem

If vPaso Then Exit Sub

On Error GoTo vError

Me.MousePointer = vbHourglass

lswCtas.ListItems.Clear

strSQL = "exec spFnd_Reserva_Cuentas_Consulta '" & pReserva & "'"
Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
  Set itmX = lswCtas.ListItems.Add(, , rs!Cod_Cuenta_Mask)
      itmX.SubItems(1) = rs!Descripcion

  rs.MoveNext
Loop
rs.Close

Me.MousePointer = vbDefault

Exit Sub


vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub


Private Sub txtCuentaCod_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtCuentaDesc.SetFocus
If KeyCode = vbKeyF4 Then
   frmCntX_ConsultaCuentas.Show vbModal
   txtCuentaCod = gCuenta
   txtCuentaDesc = fxCntX_Cuenta("D", fxCntX_CuentaFormato(False, gCuenta))
End If
End Sub

Private Sub txtCuentaCod_LostFocus()
txtCuentaCod = fxCntX_CuentaFormato(True, txtCuentaCod)
txtCuentaDesc = fxCntX_Cuenta("D", fxCntX_CuentaFormato(False, txtCuentaCod))
End Sub

Private Sub txtCuentaDesc_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then
   frmCntX_ConsultaCuentas.Show vbModal
   txtCuentaCod = gCuenta
   txtCuentaDesc = fxCntX_Cuenta("D", fxCntX_CuentaFormato(False, gCuenta))
End If
End Sub

Private Sub vGrid_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
If vPaso Then Exit Sub

If Col = 6 Then
   With vGrid
    .Row = Row
    .Col = 1
    scBarra.Tag = .Text
    
    .Col = 2
    scBarra.Caption = .Text
    
    Call sbCuentas(scBarra.Tag)
    
   End With
End If

End Sub

Private Sub vGrid_KeyDown(KeyCode As Integer, Shift As Integer)
Dim i As Long, strSQL As String

On Error GoTo vError

If vGrid.ActiveCol = vGrid.MaxCols And (KeyCode = vbKeyReturn Or KeyCode = vbKeyTab) Then
  i = fxGuardar
  If i > 0 Then
        vGrid.Row = vGrid.ActiveRow
        vGrid.Col = 1
        If vGrid.MaxRows <= vGrid.ActiveRow Then
          vGrid.MaxRows = vGrid.MaxRows + 1
          vGrid.Row = vGrid.MaxRows
        End If
  End If 'Actualiza o Inserta
End If

'Formato de Cuenta Contable
If (vGrid.ActiveCol = 3 Or vGrid.ActiveCol = 4) And (KeyCode = vbKeyReturn Or KeyCode = vbKeyTab) Then
  vGrid.Col = vGrid.ActiveCol
  vGrid.Row = vGrid.ActiveRow
  vGrid.Text = fxgCntCuentaFormato(True, vGrid.Text)
End If

'Consulta Cuentas Contables
If (vGrid.ActiveCol = 3 Or vGrid.ActiveCol = 4) And KeyCode = vbKeyF4 Then
  frmCntX_ConsultaCuentas.Show vbModal
  vGrid.Col = vGrid.ActiveCol
  vGrid.Row = vGrid.ActiveRow
  vGrid.Text = gCuenta
End If



'Inserta Linea
If KeyCode = vbKeyInsert Then
    vGrid.MaxRows = vGrid.MaxRows + 1
    vGrid.InsertRows vGrid.ActiveRow, 1
    vGrid.Row = vGrid.ActiveRow
End If


'Borrar una linea
If KeyCode = vbKeyDelete Then

        vGrid.Row = vGrid.ActiveRow
        vGrid.Col = 1

       If vGrid.Text = "" Then Exit Sub

     i = MsgBox("Esta Seguro que desea borrar este registro", vbYesNo)
     If i = vbYes Then
        
        vGrid.Col = 1
        strSQL = "delete FND_RESERVAS where COD_RESERVA = '" & vGrid.Text & "'"
        Call ConectionExecute(strSQL)
        
        
        Call Bitacora("Elimina", "Reserva de Fondos: " & vGrid.Text)
        
        vGrid.DeleteRows vGrid.ActiveRow, 1
        vGrid.MaxRows = vGrid.MaxRows - 1
        If vGrid.MaxRows = 0 Then vGrid.MaxRows = 1
        
     End If
End If

Exit Sub

vError:

  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


