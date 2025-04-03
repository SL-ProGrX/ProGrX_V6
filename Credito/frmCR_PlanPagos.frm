VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.Controls.v24.0.0.ocx"
Begin VB.Form frmCR_PlanPagos 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Plan de Pagos"
   ClientHeight    =   7140
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   18120
   LinkTopic       =   "Form1"
   ScaleHeight     =   7140
   ScaleWidth      =   18120
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin XtremeSuiteControls.GroupBox gbRevision 
      Height          =   3615
      Left            =   10920
      TabIndex        =   48
      Top             =   1680
      Visible         =   0   'False
      Width           =   5535
      _Version        =   1572864
      _ExtentX        =   9758
      _ExtentY        =   6371
      _StockProps     =   79
      BackColor       =   16777215
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Begin XtremeSuiteControls.GroupBox gbRev_CtaManual 
         Height          =   1572
         Left            =   120
         TabIndex        =   52
         Top             =   1080
         Visible         =   0   'False
         Width           =   5292
         _Version        =   1572864
         _ExtentX        =   9334
         _ExtentY        =   2773
         _StockProps     =   79
         BackColor       =   16777215
         UseVisualStyle  =   -1  'True
         BorderStyle     =   2
         Begin XtremeSuiteControls.FlatEdit txtRev_Cuota 
            Height          =   312
            Left            =   2160
            TabIndex        =   53
            Top             =   240
            Width           =   1572
            _Version        =   1572864
            _ExtentX        =   2773
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
            Alignment       =   1
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtRev_PlazoExt 
            Height          =   312
            Left            =   2160
            TabIndex        =   54
            Top             =   720
            Width           =   1572
            _Version        =   1572864
            _ExtentX        =   2773
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
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.CheckBox chkRev_PlazoAumentoAuto 
            Height          =   252
            Left            =   480
            TabIndex        =   58
            Top             =   1200
            Width           =   3252
            _Version        =   1572864
            _ExtentX        =   5736
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Aumentar Plazo automáticamente?"
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
            Appearance      =   17
            Alignment       =   1
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Cuota del Crédito"
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
            Index           =   10
            Left            =   480
            TabIndex        =   57
            Top             =   240
            Width           =   1692
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Extender Plazo"
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
            Index           =   11
            Left            =   480
            TabIndex        =   56
            Top             =   720
            Width           =   1692
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "No. de cuotas"
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
            Index           =   12
            Left            =   3840
            TabIndex        =   55
            Top             =   720
            Width           =   1332
         End
      End
      Begin XtremeSuiteControls.CheckBox chkRev_CtaDerivada 
         Height          =   252
         Left            =   240
         TabIndex        =   49
         Top             =   360
         Width           =   3612
         _Version        =   1572864
         _ExtentX        =   6371
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Considerar Cuota con abono parcial"
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
         Appearance      =   17
         Value           =   1
         Alignment       =   1
      End
      Begin XtremeSuiteControls.PushButton btnRevision_Aplicar 
         Height          =   624
         Left            =   1320
         TabIndex        =   50
         Top             =   2760
         Width           =   2532
         _Version        =   1572864
         _ExtentX        =   4466
         _ExtentY        =   1101
         _StockProps     =   79
         Caption         =   "Aplicar Revisión + Ajustes"
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
         Appearance      =   17
         Picture         =   "frmCR_PlanPagos.frx":0000
      End
      Begin XtremeSuiteControls.CheckBox chkRev_CtaAjusta 
         Height          =   252
         Left            =   240
         TabIndex        =   51
         Top             =   840
         Width           =   3612
         _Version        =   1572864
         _ExtentX        =   6371
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Ajustar Cuota Manualmente"
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
         Appearance      =   17
         Alignment       =   1
      End
   End
   Begin XtremeSuiteControls.GroupBox gbTotales 
      Height          =   615
      Left            =   120
      TabIndex        =   24
      Top             =   1680
      Width           =   17175
      _Version        =   1572864
      _ExtentX        =   30295
      _ExtentY        =   1085
      _StockProps     =   79
      BackColor       =   16777215
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      BorderStyle     =   1
      Begin XtremeSuiteControls.FlatEdit txtPlazo 
         Height          =   315
         Left            =   15960
         TabIndex        =   36
         Top             =   240
         Width           =   735
         _Version        =   1572864
         _ExtentX        =   1291
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
         Alignment       =   1
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtCuota 
         Height          =   312
         Left            =   6480
         TabIndex        =   34
         Top             =   240
         Width           =   1572
         _Version        =   1572864
         _ExtentX        =   2773
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
         Alignment       =   1
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtMonto 
         Height          =   312
         Left            =   1440
         TabIndex        =   29
         Top             =   240
         Width           =   1572
         _Version        =   1572864
         _ExtentX        =   2773
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
         Alignment       =   1
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtSaldo 
         Height          =   312
         Left            =   3960
         TabIndex        =   30
         Top             =   240
         Width           =   1572
         _Version        =   1572864
         _ExtentX        =   2773
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
         Alignment       =   1
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtPriDeduc 
         Height          =   315
         Left            =   9000
         TabIndex        =   31
         Top             =   240
         Width           =   1575
         _Version        =   1572864
         _ExtentX        =   2773
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
      Begin XtremeSuiteControls.FlatEdit txtTasa 
         Height          =   315
         Left            =   14400
         TabIndex        =   32
         Top             =   240
         Width           =   735
         _Version        =   1572864
         _ExtentX        =   1291
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
         Alignment       =   1
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtFecUltCta 
         Height          =   315
         Left            =   11880
         TabIndex        =   60
         Top             =   240
         Width           =   1575
         _Version        =   1572864
         _ExtentX        =   2773
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
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Ult.Mov.Cta"
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
         Index           =   13
         Left            =   10800
         TabIndex        =   61
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Plazo"
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
         Index           =   9
         Left            =   15240
         TabIndex        =   35
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Cuota"
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
         Index           =   8
         Left            =   5760
         TabIndex        =   33
         Top             =   240
         Width           =   852
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Monto"
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
         Left            =   0
         TabIndex        =   28
         Top             =   240
         Width           =   1332
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Saldo"
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
         Index           =   3
         Left            =   3240
         TabIndex        =   27
         Top             =   240
         Width           =   852
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Pri.Ded."
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
         Index           =   6
         Left            =   8040
         TabIndex        =   26
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Tasa"
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
         Index           =   7
         Left            =   13680
         TabIndex        =   25
         Top             =   240
         Width           =   735
      End
   End
   Begin XtremeSuiteControls.GroupBox gbAcciones 
      Height          =   852
      Left            =   10680
      TabIndex        =   15
      Top             =   900
      Width           =   6612
      _Version        =   1572864
      _ExtentX        =   11663
      _ExtentY        =   1503
      _StockProps     =   79
      Caption         =   "Acciones:"
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      BorderStyle     =   2
      Begin XtremeSuiteControls.GroupBox GroupBox3 
         Height          =   735
         Left            =   4560
         TabIndex        =   44
         Top             =   0
         Width           =   1935
         _Version        =   1572864
         _ExtentX        =   3408
         _ExtentY        =   1291
         _StockProps     =   79
         BackColor       =   16777215
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         Begin XtremeSuiteControls.RadioButton rbExport 
            Height          =   252
            Index           =   0
            Left            =   840
            TabIndex        =   46
            Top             =   144
            Width           =   972
            _Version        =   1572864
            _ExtentX        =   1714
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Excel"
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   7.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            UseVisualStyle  =   -1  'True
            Appearance      =   17
            Value           =   -1  'True
         End
         Begin XtremeSuiteControls.PushButton btnExportar 
            Height          =   420
            Left            =   120
            TabIndex        =   45
            Top             =   216
            Width           =   492
            _Version        =   1572864
            _ExtentX        =   868
            _ExtentY        =   741
            _StockProps     =   79
            BackColor       =   16777215
            FlatStyle       =   -1  'True
            UseVisualStyle  =   -1  'True
            Appearance      =   17
            Picture         =   "frmCR_PlanPagos.frx":09C3
         End
         Begin XtremeSuiteControls.RadioButton rbExport 
            Height          =   252
            Index           =   1
            Left            =   840
            TabIndex        =   47
            Top             =   396
            Width           =   972
            _Version        =   1572864
            _ExtentX        =   1714
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "HTML"
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   7.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            UseVisualStyle  =   -1  'True
            Appearance      =   17
         End
      End
      Begin XtremeSuiteControls.GroupBox gbPrinter 
         Height          =   735
         Left            =   1560
         TabIndex        =   40
         Top             =   0
         Width           =   3015
         _Version        =   1572864
         _ExtentX        =   5318
         _ExtentY        =   1296
         _StockProps     =   79
         BackColor       =   16777215
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         Begin XtremeSuiteControls.PushButton btnImpresora 
            Height          =   540
            Left            =   120
            TabIndex        =   41
            Top             =   120
            Width           =   492
            _Version        =   1572864
            _ExtentX        =   868
            _ExtentY        =   952
            _StockProps     =   79
            BackColor       =   16777215
            FlatStyle       =   -1  'True
            UseVisualStyle  =   -1  'True
            Appearance      =   17
            Picture         =   "frmCR_PlanPagos.frx":0FC2
         End
         Begin XtremeSuiteControls.RadioButton rbPrinter 
            Height          =   255
            Index           =   0
            Left            =   1560
            TabIndex        =   42
            Top             =   120
            Width           =   1455
            _Version        =   1572864
            _ExtentX        =   2566
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Plan de Pagos"
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   7.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            UseVisualStyle  =   -1  'True
            Appearance      =   17
            Value           =   -1  'True
         End
         Begin XtremeSuiteControls.RadioButton rbPrinter 
            Height          =   255
            Index           =   1
            Left            =   1560
            TabIndex        =   43
            Top             =   360
            Width           =   1455
            _Version        =   1572864
            _ExtentX        =   2566
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Movimientos"
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   7.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            UseVisualStyle  =   -1  'True
            Appearance      =   17
         End
         Begin XtremeSuiteControls.PushButton btnEmail 
            Height          =   540
            Left            =   720
            TabIndex        =   59
            Top             =   120
            Width           =   495
            _Version        =   1572864
            _ExtentX        =   868
            _ExtentY        =   952
            _StockProps     =   79
            BackColor       =   16777215
            FlatStyle       =   -1  'True
            UseVisualStyle  =   -1  'True
            Appearance      =   17
            Picture         =   "frmCR_PlanPagos.frx":177E
         End
      End
      Begin XtremeSuiteControls.PushButton btnRefrescar 
         Height          =   620
         Left            =   240
         TabIndex        =   16
         Top             =   100
         Width           =   612
         _Version        =   1572864
         _ExtentX        =   1080
         _ExtentY        =   1094
         _StockProps     =   79
         BackColor       =   16777215
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         Picture         =   "frmCR_PlanPagos.frx":1F9B
      End
      Begin XtremeSuiteControls.PushButton btnRevisar 
         Height          =   620
         Left            =   840
         TabIndex        =   17
         Top             =   100
         Width           =   612
         _Version        =   1572864
         _ExtentX        =   1080
         _ExtentY        =   1094
         _StockProps     =   79
         BackColor       =   16777215
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         Picture         =   "frmCR_PlanPagos.frx":2928
      End
   End
   Begin XtremeSuiteControls.GroupBox gbDetalle 
      Height          =   2532
      Left            =   120
      TabIndex        =   9
      Top             =   4440
      Width           =   12372
      _Version        =   1572864
      _ExtentX        =   21823
      _ExtentY        =   4466
      _StockProps     =   79
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
      Appearance      =   17
      BorderStyle     =   1
      Begin XtremeSuiteControls.TabControl tcMain 
         Height          =   2172
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   12132
         _Version        =   1572864
         _ExtentX        =   21399
         _ExtentY        =   3831
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
         ItemCount       =   5
         SelectedItem    =   4
         Item(0).Caption =   "Cargos"
         Item(0).ControlCount=   1
         Item(0).Control(0)=   "lsw"
         Item(1).Caption =   "Pólizas"
         Item(1).ControlCount=   1
         Item(1).Control(0)=   "lswPolizas"
         Item(2).Caption =   "Documentos"
         Item(2).ControlCount=   4
         Item(2).Control(0)=   "lswValores"
         Item(2).Control(1)=   "lswDocumentos"
         Item(2).Control(2)=   "chkTodos"
         Item(2).Control(3)=   "chkValores"
         Item(3).Caption =   "Ajustes"
         Item(3).ControlCount=   1
         Item(3).Control(0)=   "lswAjustes"
         Item(4).Caption =   "Activación"
         Item(4).ControlCount=   3
         Item(4).Control(0)=   "Label3"
         Item(4).Control(1)=   "dtpActivacion"
         Item(4).Control(2)=   "btnActivar"
         Begin XtremeSuiteControls.ListView lswValores 
            Height          =   1780
            Left            =   -63880
            TabIndex        =   13
            Top             =   360
            Visible         =   0   'False
            Width           =   6012
            _Version        =   1572864
            _ExtentX        =   10604
            _ExtentY        =   3140
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
         Begin XtremeSuiteControls.ListView lswPolizas 
            Height          =   1780
            Left            =   -70000
            TabIndex        =   12
            Top             =   360
            Visible         =   0   'False
            Width           =   12132
            _Version        =   1572864
            _ExtentX        =   21399
            _ExtentY        =   3140
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
         Begin XtremeSuiteControls.ListView lsw 
            Height          =   1780
            Left            =   -70000
            TabIndex        =   11
            Top             =   360
            Visible         =   0   'False
            Width           =   12135
            _Version        =   1572864
            _ExtentX        =   21405
            _ExtentY        =   3140
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
         Begin XtremeSuiteControls.ListView lswDocumentos 
            Height          =   1780
            Left            =   -68680
            TabIndex        =   14
            Top             =   360
            Visible         =   0   'False
            Width           =   4692
            _Version        =   1572864
            _ExtentX        =   8276
            _ExtentY        =   3140
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
         Begin XtremeSuiteControls.ListView lswAjustes 
            Height          =   1780
            Left            =   -70000
            TabIndex        =   37
            Top             =   360
            Visible         =   0   'False
            Width           =   12132
            _Version        =   1572864
            _ExtentX        =   21399
            _ExtentY        =   3140
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
         Begin XtremeSuiteControls.CheckBox chkTodos 
            Height          =   972
            Left            =   -69880
            TabIndex        =   38
            Top             =   480
            Visible         =   0   'False
            Width           =   1092
            _Version        =   1572864
            _ExtentX        =   1926
            _ExtentY        =   1714
            _StockProps     =   79
            Caption         =   "Mostrar todos los documentos registrados?"
            BackColor       =   -2147483633
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   7.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            UseVisualStyle  =   -1  'True
            TextAlignment   =   5
            Appearance      =   16
         End
         Begin XtremeSuiteControls.CheckBox chkValores 
            Height          =   612
            Left            =   -69880
            TabIndex        =   39
            Top             =   1560
            Visible         =   0   'False
            Width           =   1092
            _Version        =   1572864
            _ExtentX        =   1926
            _ExtentY        =   1080
            _StockProps     =   79
            Caption         =   "Mostrar Valores"
            BackColor       =   -2147483633
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   7.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            UseVisualStyle  =   -1  'True
            TextAlignment   =   5
            Appearance      =   16
         End
         Begin XtremeSuiteControls.DateTimePicker dtpActivacion 
            Height          =   330
            Left            =   2160
            TabIndex        =   63
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
         Begin XtremeSuiteControls.PushButton btnActivar 
            Height          =   615
            Left            =   3840
            TabIndex        =   64
            Top             =   960
            Width           =   1935
            _Version        =   1572864
            _ExtentX        =   3413
            _ExtentY        =   1085
            _StockProps     =   79
            Caption         =   "Activar Cuota"
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
            Appearance      =   17
            Picture         =   "frmCR_PlanPagos.frx":3106
         End
         Begin XtremeSuiteControls.Label Label3 
            Height          =   495
            Left            =   240
            TabIndex        =   62
            Top             =   840
            Width           =   1695
            _Version        =   1572864
            _ExtentX        =   2990
            _ExtentY        =   873
            _StockProps     =   79
            Caption         =   "Activación de Cuota que vence el:"
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
   End
   Begin MSComctlLib.StatusBar StatusBarX 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   7
      Top             =   6885
      Width           =   18120
      _ExtentX        =   31962
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   7
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Text            =   "Cuotas."
            TextSave        =   "Cuotas."
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   4304
            MinWidth        =   4304
            Text            =   "Inicio"
            TextSave        =   "Inicio"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   4304
            MinWidth        =   4304
            Text            =   "Corte"
            TextSave        =   "Corte"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Text            =   "Dias"
            TextSave        =   "Dias"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   6068
            MinWidth        =   6068
            Text            =   "Intereses"
            TextSave        =   "Intereses"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   4304
            MinWidth        =   4304
            Text            =   "Cargos"
            TextSave        =   "Cargos"
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   4304
            MinWidth        =   4304
            Text            =   "Dias Mora"
            TextSave        =   "Dias Mora"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Timer TimerX 
      Interval        =   20
      Left            =   6840
      Top             =   360
   End
   Begin FPSpreadADO.fpSpread vGrid 
      Height          =   2055
      Left            =   120
      TabIndex        =   1
      Top             =   2400
      Width           =   12255
      _Version        =   524288
      _ExtentX        =   21616
      _ExtentY        =   3625
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
      MaxCols         =   34
      SpreadDesigner  =   "frmCR_PlanPagos.frx":38E4
      VScrollSpecial  =   -1  'True
      VScrollSpecialType=   2
      AppearanceStyle =   1
   End
   Begin XtremeSuiteControls.FlatEdit txtDiaPago 
      Height          =   312
      Left            =   9000
      TabIndex        =   18
      Top             =   1320
      Width           =   1572
      _Version        =   1572864
      _ExtentX        =   2773
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
   Begin XtremeSuiteControls.FlatEdit txtFactorCalculo 
      Height          =   312
      Left            =   9000
      TabIndex        =   19
      Top             =   960
      Width           =   1572
      _Version        =   1572864
      _ExtentX        =   2773
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
   Begin XtremeSuiteControls.FlatEdit txtCedula 
      Height          =   312
      Left            =   1560
      TabIndex        =   20
      Top             =   960
      Width           =   1572
      _Version        =   1572864
      _ExtentX        =   2773
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
   Begin XtremeSuiteControls.FlatEdit txtLinea 
      Height          =   312
      Left            =   1560
      TabIndex        =   21
      Top             =   1320
      Width           =   1572
      _Version        =   1572864
      _ExtentX        =   2773
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
      Left            =   3120
      TabIndex        =   22
      Top             =   960
      Width           =   4332
      _Version        =   1572864
      _ExtentX        =   7641
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
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtLineaDesc 
      Height          =   312
      Left            =   3120
      TabIndex        =   23
      Top             =   1320
      Width           =   4332
      _Version        =   1572864
      _ExtentX        =   7641
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
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin VB.Label lblOficina 
      BackStyle       =   0  'Transparent
      Caption         =   "Oficina...."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   960
      TabIndex        =   8
      Top             =   520
      Width           =   7095
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Día de Pago"
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
      Index           =   5
      Left            =   7560
      TabIndex        =   6
      Top             =   1320
      Width           =   1335
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Factor Cálculo"
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
      Index           =   4
      Left            =   7560
      TabIndex        =   5
      Top             =   960
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Línea"
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
      Index           =   1
      Left            =   120
      TabIndex        =   4
      Top             =   1320
      Width           =   1332
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Identificación"
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
      Index           =   0
      Left            =   120
      TabIndex        =   3
      Top             =   960
      Width           =   1332
   End
   Begin VB.Label lblOperacion 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "1234567890"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   8040
      TabIndex        =   2
      Top             =   240
      Width           =   2535
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Plan de Pagos"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   360
      Index           =   0
      Left            =   960
      TabIndex        =   0
      Top             =   120
      Width           =   3855
   End
   Begin VB.Image imgBanner 
      Height          =   852
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   16572
   End
End
Attribute VB_Name = "frmCR_PlanPagos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem

Dim vPaso As Boolean, pLinea As Currency



Private Sub btnActivar_Click()
Dim pProceso As Currency, pPass As Boolean

On Error GoTo vError

pPass = False

If txtPlazo.Text = "999" Then
    strSQL = "select dbo.fxSIFDateTimeToProceso('" & Format(dtpActivacion.Value, "yyyy-mm-dd") & "') as 'Proceso'"
    Call OpenRecordSet(rs, strSQL)
     pProceso = rs!Proceso
    rs.Close
    
    strSQL = "exec dbo.spCrdPlanPagosActivaRetenciones_Operacion " & lblOperacion.Caption & ", " & pProceso
    Call ConectionExecute(strSQL)

    pPass = True
    
Else
    strSQL = "select Id_Seq from Crd_Operacion_Plan_Pagos where Id_Solicitud = " & lblOperacion.Caption & " and Estado = 'P' and Fecha_Corte <= '" _
             & Format(dtpActivacion.Value, "yyyy-mm-dd") & "' order by Id_Seq asc"
    Call OpenRecordSet(rs, strSQL)
    Do While Not rs.EOF
      strSQL = "exec spCrdPlanPagosActivaCuota " & lblOperacion.Caption & ", " & rs!Id_seq
      Call ConectionExecute(strSQL)
      
      pPass = True
      
      rs.MoveNext
    Loop
    rs.Close

End If

If pPass Then
    MsgBox "Activación de Cuotas Procesada!", vbInformation
    Call Bitacora("Activa", "Cuotas Operacion: " & lblOperacion.Caption & " al " & Format(dtpActivacion.Value, "yyyy-mm-dd"))
End If

Call sbInicializa

Exit Sub

vError:
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub btnEmail_Click()
Dim i As Integer, vDetalle As String

On Error GoTo vError

 i = MsgBox("Desea Enviar al Correo el Estado de Cuenta de la Operación?", vbYesNo)
 If i = vbNo Then
        Exit Sub
 End If
  

 strSQL = "exec spSys_Notifica_Credito_Estado_Operacion " & lblOperacion.Caption & ",'" & glogon.Usuario & "'"
 Call ConectionExecute(strSQL)

 vDetalle = "Notificación Email de Estado de la Operación: " & lblOperacion.Caption

 Call Bitacora("Aplica", vDetalle)
 
 MsgBox "Estado de Cuenta de la Operacion enviado al correo de la persona!", vbInformation
 
Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub btnExportar_Click()
Dim vHeaders As vGridHeaders

vHeaders.Columnas = vGrid.MaxCols
vHeaders.Headers(1) = ""
vHeaders.Headers(2) = "Linea"
vHeaders.Headers(3) = "No. Cuota"
vHeaders.Headers(4) = "Proceso"
vHeaders.Headers(5) = "Fec.Inicio"
vHeaders.Headers(6) = "Fec.Corte"
vHeaders.Headers(7) = "Fec.Pago"
vHeaders.Headers(8) = "Tasa"
vHeaders.Headers(9) = "Plazo"
vHeaders.Headers(10) = "Cuota"
vHeaders.Headers(11) = "IVA"
vHeaders.Headers(12) = "Cargos"
vHeaders.Headers(13) = "Póliza"
vHeaders.Headers(14) = "Int.Cor."
vHeaders.Headers(15) = "Int.Mor."
vHeaders.Headers(16) = "Amortización"
vHeaders.Headers(17) = "Saldo Ant."
vHeaders.Headers(18) = "Saldo Actual"
vHeaders.Headers(19) = "Dias"
vHeaders.Headers(20) = "Estado"
vHeaders.Headers(21) = "Dias Atraso"
vHeaders.Headers(22) = "Mov.Fecha"
vHeaders.Headers(23) = "Mov.Total"
vHeaders.Headers(24) = "Mov.IVA"
vHeaders.Headers(25) = "Mov.Cargos"
vHeaders.Headers(26) = "Mov.Póliza"
vHeaders.Headers(27) = "Mov.Int.Cor."
vHeaders.Headers(28) = "Mov.Int.Mor."
vHeaders.Headers(29) = "Mov.Amortización"
vHeaders.Headers(30) = "Usr.Caja"
vHeaders.Headers(31) = "Tipo.Doc."
vHeaders.Headers(32) = "No. Documento"
vHeaders.Headers(33) = ""
vHeaders.Headers(34) = "Concepto"

Select Case True
  Case rbExport.Item(0).Value    'EXCEL
      Call sbSIFGridExportar(vGrid, vHeaders, "Plan_Pagos_Op" & lblOperacion.Caption)
  Case rbExport.Item(1).Value    'HTML
      Call sbSIFGridExportar(vGrid, vHeaders, "Plan_Pagos_Op" & lblOperacion.Caption, "HTML")
End Select

End Sub

Private Sub btnImpresora_Click()

Select Case True
  Case rbPrinter.Item(0).Value  'Plan
      Call sbReportes("Plan")
  Case rbPrinter.Item(1).Value  'Movimientos
      Call sbReportes("Movimientos")
End Select

End Sub

Private Sub btnRefrescar_Click()
       Call sbInicializa
End Sub

Private Sub btnRevisar_Click()
If gbRevision.Visible Then
   gbRevision.Visible = False
Else
      
   chkRev_CtaDerivada.Value = xtpChecked
   
   chkRev_CtaAjusta.Value = xtpUnchecked
   chkRev_CtaAjusta_Click
   
   txtRev_Cuota.Text = txtCuota.Text
   txtRev_PlazoExt.Text = 0
   
   chkRev_PlazoAumentoAuto.Value = xtpUnchecked
   
   gbRevision.Visible = True
End If
End Sub

Private Sub btnRevision_Aplicar_Click()
       Call sbRevisarPlan
End Sub

Private Sub chkRev_CtaAjusta_Click()

gbRev_CtaManual.Visible = IIf((chkRev_CtaAjusta.Value = xtpChecked), True, False)

End Sub

Private Sub chkTodos_Click()
Call sbLoad_DT_Documentos
End Sub

Private Sub chkValores_Click()
Call Form_Resize
End Sub

Private Sub Form_Load()
vModulo = 3

Set imgBanner.Picture = frmContenedor.imgBanner_01.Picture
vGrid.AppearanceStyle = fxGridStyle

lblOperacion.Caption = Operacion.OperacionConsulta
pLinea = 0

With lsw.ColumnHeaders
  .Clear
  .Add , , "Linea", 1000
  .Add , , "Seq", 1000
  .Add , , "Detalle", 3500
  .Add , , "Monto", 1800, vbRightJustify
  .Add , , "Saldo", 1800, vbRightJustify
  .Add , , "Unidad", 1200, vbCenter
  .Add , , "C.Costo", 1200, vbCenter
  .Add , , "Cuenta", 2400
  .Add , , "Abono", 1800, vbRightJustify
  .Add , , "Pendiente", 1800, vbRightJustify
  .Add , , "Descripción", 3500
End With

With lswAjustes.ColumnHeaders
  .Clear
  .Add , , "Fecha", 2200
  .Add , , "Usuario", 2100, vbCenter
  .Add , , "Movimiento", 2100, vbCenter
  .Add , , "Detalle", 2500
  .Add , , "Notas", 4500
End With

With lswPolizas.ColumnHeaders
  .Clear
  .Add , , "Seq", 1000
  .Add , , "Código", 1500, vbCenter
  .Add , , "Descripción", 3500
  .Add , , "No. Póliza", 1500, vbCenter
  .Add , , "No. Contrato", 1500, vbCenter
  .Add , , "Monto", 1800, vbRightJustify
  .Add , , "Movimiento", 1800, vbRightJustify
  .Add , , "Saldo", 1800, vbRightJustify
  .Add , , "Fecha", 1800
  .Add , , "Usuario", 1800, vbCenter
  .Add , , "Aseguradora", 3500
  .Add , , "Cuenta", 2500, vbCenter
  .Add , , "Cuenta Desc...", 3500
End With

With lswDocumentos.ColumnHeaders
  .Clear
  .Add , , "Tipo Doc.", 2500
  .Add , , "No. Transacción", 2100
  .Add , , "Documento", 2500
  .Add , , "Concepto", 2500
  .Add , , "Monto Doc.", 1800, vbRightJustify
  .Add , , "Monto Apl.", 1800, vbRightJustify
  .Add , , "Int.Cor.", 1800, vbRightJustify
  .Add , , "Int.Mor.", 1800, vbRightJustify
  .Add , , "Cargos", 1800, vbRightJustify
  .Add , , "Pólizas", 1800, vbRightJustify
  .Add , , "Principal", 1800, vbRightJustify
  .Add , , "Fecha", 2100
  .Add , , "Usuario", 2100, vbCenter
End With


With lswValores.ColumnHeaders
  .Clear
  .Add , , "Tipo Doc.", 1000
  .Add , , "No. Transacción", 2100
  .Add , , "Tipo Valor", 2500
  .Add , , "No. Documento", 2100, vbCenter
  .Add , , "Monto Documento", 1800, vbRightJustify
  .Add , , "Monto Aplicado", 1800, vbRightJustify
  .Add , , "Divisa", 1100, vbCenter
  .Add , , "Tipo Cambio", 1800, vbRightJustify
  .Add , , "Saldo Favor Id", 1800, vbCenter
  .Add , , "Fecha", 2100
  .Add , , "Usuario", 2100, vbCenter
  .Add , , "Referencias", 3500
  .Add , , "Cuenta", 2500, vbCenter
  .Add , , "Cuenta Desc...", 3500
End With

Call Formularios(Me)

btnActivar.Tag = btnRevisar.Tag

Call RefrescaTags(Me)

End Sub


Private Sub sbInicializa()

On Error GoTo vError

pLinea = 0
gbDetalle.Caption = ""

tcMain.Item(0).Selected = True
lsw.ListItems.Clear


strSQL = "select S.cedula,S.nombre,R.codigo,C.descripcion,R.montoApr,R.Saldo,R.cuota,R.Interesv,R.plazo, R.cuota" _
       & ", R.Dia_Pago, R.Base_Calculo, R.PriDeduc, R.FecUlt, R.int as TasaO, Ofi.descripcion as 'OficinaX'" _
       & ", CONVERT(varchar, dbo.fxCrd_Operacion_Cta_Ultimo_Corte(R.ID_SOLICITUD)  ,23)  as 'CtaFechaUltCorte'" _
       & " from Socios S inner join Reg_creditos R on S.cedula = R.cedula" _
       & " inner join catalogo C on R.codigo = C.codigo" _
       & " left join SIF_Oficinas Ofi on R.cod_oficina_r = Ofi.cod_oficina" _
       & " where R.id_solicitud = " & lblOperacion.Caption
Call OpenRecordSet(rs, strSQL)

txtCedula.Text = rs!Cedula
txtNombre.Text = rs!Nombre

txtLinea.Text = rs!Codigo
txtLineaDesc.Text = rs!Descripcion

lblOficina.Caption = rs!OficinaX & ""

txtMonto.Text = Format(rs!montoapr, "Standard")
txtSaldo.Text = Format(rs!Saldo, "Standard")
txtCuota.Text = Format(rs!Cuota, "Standard")

txtCuota.Text = Format(rs!Cuota, "Standard")

txtPlazo.Text = CStr(rs!Plazo)

txtTasa.Text = Format(rs!interesv, "Standard")
txtTasa.ToolTipText = "Tasa Original: " & Format(rs!TasaO, "Standard")

txtDiaPago.Text = rs!dia_pago
If rs!dia_pago = 32 Then
   txtDiaPago.Text = "Ultimo Día del Mes"
End If

If rs!Base_Calculo = "06" Then
    txtPrideduc.Text = Format(rs!PriDeduc, "####-##.0")
Else
    txtPrideduc.Text = Format(rs!PriDeduc, "####-##")
End If

txtFecUltCta.Text = rs!CtaFechaUltCorte

dtpActivacion.Value = DateAdd("d", 1, rs!CtaFechaUltCorte)
dtpActivacion.MinDate = DateAdd("d", 1, rs!CtaFechaUltCorte)
dtpActivacion.MaxDate = DateAdd("d", 32, rs!CtaFechaUltCorte)

txtFactorCalculo.Text = fxCrd_Factor_Calculo(rs!Base_Calculo)

rs.Close

strSQL = "select 0 as 'Sep1', TP.Id_Seq, TP.num_cuota, TP.Fecha_Proceso, TP.Fecha_Inicio, TP.Fecha_Corte, TP.Fecha_Pago, TP.Tasa, TP.Plazo, TP.Cuota, isnull(Tp.IVA,0) as 'IVA', TP.Cargos" _
       & ", TP.Poliza, TP.IntCor, TP.IntMor, TP.Principal, TP.Saldo_Anterior, TP.Saldo_Actual" _
       & ", TP.Dias_Calculo,case TP.Estado when 'A' then 'Activa' when 'P' then 'Pendiente'" _
       & " when 'C' then 'Cancelada' when 'N' then 'Anulada' end as 'Estado', Mov.Mora_Dias" _
       & ", Mov.Mov_Fecha, Mov.Mov_Monto, isnull(Mov.Mov_IVA,0) as 'Mov_IVA', Mov.Mov_Cargos, Mov.Mov_Poliza, Mov.Mov_IntCor, Mov.Mov_IntMor, Mov.Mov_Principal, Mov.Cod_Caja + '/' + Mov.Mov_Usuario" _
       & ", isnull(Mov.Tipo_Documento,TP.Tipo_Documento) as 'Tipo_Documento' , isnull(Mov.Num_Comprobante,TP.Num_Comprobante) as 'Num_Comprobante',0 as 'Sep2',Con.Descripcion as 'Concepto'" _
       & " from crd_operacion_plan_pagos TP left join CRD_OPERACION_TRANSAC Mov on TP.Id_Seq = Mov.Id_Seq and TP.id_solicitud = Mov.Id_solicitud" _
       & " left join SIF_Conceptos Con on isnull(Mov.cod_concepto,TP.cod_concepto) = Con.cod_Concepto" _
       & " where TP.id_solicitud = " & lblOperacion.Caption _
       & " order by TP.Id_Seq"

vPaso = True
    Call sbCargaGridFps7(vGrid, 34, strSQL, False)
vPaso = False

strSQL = "select max(num_cuota) as Cuotas, sum(IntCor + IntMor) as Intereses, Sum(Cargos) as Cargos" _
       & ", sum(Dias_Calculo) as Dias, min(Fecha_Pago) as Inicio, max(Fecha_Pago) as Corte, Sum(Mora_Dias) as MoraDias" _
       & " from crd_operacion_plan_pagos" _
       & " where id_solicitud = " & lblOperacion.Caption
Call OpenRecordSet(rs, strSQL)
    StatusBarX.Panels.Item(1).Text = "Cuotas..: " & rs!Cuotas
    StatusBarX.Panels.Item(2).Text = "Inicio..: " & Format(rs!Inicio, "dd/mm/yyyy")
    StatusBarX.Panels.Item(3).Text = "Corte..: " & Format(rs!Corte, "dd/mm/yyyy")
    StatusBarX.Panels.Item(4).Text = "Días..: " & rs!Dias
    StatusBarX.Panels.Item(5).Text = "Intereses..: " & Format(rs!Intereses, "Standard")
    StatusBarX.Panels.Item(6).Text = "Cargos ..:  " & Format(rs!Cargos, "Standard")
    StatusBarX.Panels.Item(7).Text = "Mora días..: " & rs!MoraDias
rs.Close

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub Form_Resize()
On Error Resume Next

vGrid.Width = Me.Width - 450

imgBanner.Width = Me.Width

vGrid.Height = Me.Height - vGrid.top - 580 - StatusBarX.Height - gbDetalle.Height

gbDetalle.top = vGrid.top + vGrid.Height + 50
gbDetalle.Left = vGrid.Left
gbDetalle.Width = vGrid.Width

gbTotales.Width = vGrid.Width

tcMain.Width = gbDetalle.Width - 100

lsw.Width = gbDetalle.Width - 100
lswPolizas.Width = gbDetalle.Width - 100
lswAjustes.Width = gbDetalle.Width - 100


If chkValores.Value = xtpChecked Then
    lswValores.Visible = True
    lswDocumentos.Width = (gbDetalle.Width - (chkTodos.Width + 340)) / 2
    lswValores.Width = lswDocumentos.Width
    lswValores.Left = lswDocumentos.Left + lswDocumentos.Width + 50

Else
    lswValores.Visible = False
    lswDocumentos.Width = (gbDetalle.Width - (chkTodos.Width + 340))
End If



End Sub

Private Sub lswDocumentos_ItemClick(ByVal Item As XtremeSuiteControls.ListViewItem)

Call sbLoad_DT_Valores(Item.Tag, Item.SubItems(1))

End Sub


Private Sub tcMain_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)

On Error GoTo vErrorMain

Select Case Item.Index
  Case 0 'Cargos
        Call sbLoad_DT_Cargos
  Case 1 'Polizas
        Call sbLoad_DT_Polizas
  Case 2 'Documentos
        Call sbLoad_DT_Documentos
  Case 3 'Ajustes
        Call sbLoad_DT_Ajustes(lblOperacion.Caption)
End Select

vErrorMain:

End Sub

Private Sub TimerX_Timer()
TimerX.Interval = 0
Call sbInicializa
End Sub

Private Sub sbRevisarPlan()
Dim i As Integer, vDetalle As String

On Error GoTo vError

 i = MsgBox("Desea revisar el Plan de Pagos de esta Operación?", vbYesNo)
 If i = vbNo Then
        Exit Sub
 End If
  
 If fxCrd_Factor_Calculo(txtFactorCalculo.Text) = "06" Then
    If (CCur(txtRev_Cuota.Text) < CCur(txtSaldo.Text) * CCur(txtTasa.Text) / 2400) And chkRev_CtaAjusta.Value = xtpChecked Then
           MsgBox "La Cuota Manual no es válida porque es menor al cobro de intereses mínimo!", vbExclamation
           Exit Sub
    End If
 
 Else
    If (CCur(txtRev_Cuota.Text) < CCur(txtSaldo.Text) * CCur(txtTasa.Text) / 1200) And chkRev_CtaAjusta.Value = xtpChecked Then
           MsgBox "La Cuota Manual no es válida porque es menor al cobro de intereses mínimo!", vbExclamation
           Exit Sub
    End If
 End If

 strSQL = "exec spCrdPlanPagosRevision " & lblOperacion.Caption & ",'" & glogon.Usuario & "',1"

 vDetalle = "Revisión de Plan de Pago, Operacion: " & lblOperacion.Caption
 
 If chkRev_CtaAjusta.Value = xtpChecked Then
    vDetalle = vDetalle & ", Aj.Cta.Manual: " & txtRev_Cuota.Text
    strSQL = strSQL & "," & CCur(txtRev_Cuota.Text)
 Else
    strSQL = strSQL & ",0"
 End If
  
 If CLng(txtRev_PlazoExt.Text) >= 0 Then
    strSQL = strSQL & "," & CLng(txtRev_PlazoExt.Text)
    vDetalle = vDetalle & ", Ext.Plazo: " & CLng(txtRev_PlazoExt.Text)
 Else
    strSQL = strSQL & ",0"
 End If
 
 strSQL = strSQL & "," & chkRev_CtaDerivada.Value & "," & chkRev_PlazoAumentoAuto.Value
 
 Call ConectionExecute(strSQL)


 vDetalle = vDetalle & ", Ext.Plazo: " & CLng(txtRev_PlazoExt.Text) _
          & ", Cta.Deriv: " & chkRev_CtaDerivada.Value _
          & ", Plazo Aumenta: " & chkRev_PlazoAumentoAuto.Value

 Call Bitacora("Aplica", vDetalle)
 
 MsgBox "Plan de Pagos Revisado Satisfactoriamente!", vbInformation
 
 
 gbRevision.Visible = False
 Call sbInicializa

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub tlbX_ButtonClick(ByVal Button As MSComctlLib.Button)

Select Case Button.Key
  Case "Refrescar"
       Call sbInicializa
  Case "Revisar"
       Call sbRevisarPlan
  Case "Reporte"
  Case "Exportar"
End Select
End Sub

Private Sub sbReportes(pTipo As String)
Dim vFecha As Date

Me.MousePointer = vbHourglass

vFecha = fxFechaServidor

With frmContenedor.Crt
 .Reset
 .WindowShowGroupTree = False
 .WindowShowPrintSetupBtn = True
 .WindowShowRefreshBtn = True
 .WindowShowSearchBtn = True
 .WindowState = crptMaximized
 .WindowTitle = "Reportes del Módulo de Créditos"

 .Connect = glogon.ConectRPT
                
    Select Case pTipo
      Case "Plan"
         .ReportFileName = SIFGlobal.fxPathReportes("Credito_PlanPagos.rpt")
      Case "Movimientos"
         .ReportFileName = SIFGlobal.fxPathReportes("Credito_PlanPagosMov.rpt")
      Case "Estudio"
         .ReportFileName = SIFGlobal.fxPathReportes("Credito_PlanPagosEstudio.rpt")
    End Select

 .Formulas(0) = "fxFecha='FECHA: " & Format(vFecha, "dd/mm/yyyy  hh:mm:ss") & "'"
 .Formulas(1) = "fxEmpresa='" & GLOBALES.gstrNombreEmpresa & "'"
 .Formulas(2) = "fxUsuario='USER: " & glogon.Usuario & "'"
 .Formulas(3) = "fxOficina='" & GLOBALES.gOficina & "'"
 
 .SelectionFormula = "{REG_CREDITOS.ID_SOLICITUD} = " & lblOperacion.Caption
 
 .SubreportToChange = "sbCorte"
 .StoredProcParam(0) = lblOperacion.Caption
 .StoredProcParam(1) = Format(vFecha, "yyyy/mm/dd")
 
 
 
 .PrintReport

End With

Me.MousePointer = vbDefault


End Sub


Private Sub sbLoad_DT_Cargos()

On Error GoTo vError
 
Me.MousePointer = vbHourglass

strSQL = "exec spCrd_Operacion_Consulta_Cargos " & CLng(lblOperacion.Caption) & "," & pLinea
 
lsw.ListItems.Clear
Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
  Set itmX = lsw.ListItems.Add(, , rs!Linea)
      itmX.SubItems(1) = rs!Id_seq
      itmX.SubItems(2) = rs!Detalle
      itmX.SubItems(3) = Format(rs!Monto, "Standard")
      itmX.SubItems(4) = Format(rs!Mov_Saldo, "Standard")
      itmX.SubItems(5) = rs!Cod_Unidad
      itmX.SubItems(6) = rs!Cod_Centro_Costo
      itmX.SubItems(7) = rs!Cod_Cuenta_Mask
      itmX.SubItems(8) = Format(rs!Abono, "Standard")
      itmX.SubItems(9) = Format(rs!Monto - rs!Abono, "Standard")
      itmX.SubItems(10) = rs!Descripcion
  rs.MoveNext
Loop
rs.Close

Me.MousePointer = vbDefault

If lsw.ListItems.Count = 0 Then
  tcMain.Item(1).Selected = True
End If

Exit Sub

vError:
  Me.MousePointer = vbDefault
 
End Sub


Private Sub sbLoad_DT_Polizas()

On Error GoTo vError
 
Me.MousePointer = vbHourglass

strSQL = "exec spCrd_Operacion_Consulta_Polizas " & CLng(lblOperacion.Caption) & "," & pLinea

lswPolizas.ListItems.Clear
Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
  Set itmX = lswPolizas.ListItems.Add(, , rs!Id_seq)
      itmX.SubItems(1) = rs!cod_poliza
      itmX.SubItems(2) = rs!Descripcion
      itmX.SubItems(3) = rs!Num_Poliza
      itmX.SubItems(4) = rs!Num_Contrato
      itmX.SubItems(5) = Format(rs!Monto, "Standard")
      itmX.SubItems(6) = Format(rs!Mov_Monto, "Standard")
      itmX.SubItems(7) = Format(rs!Mov_Saldo, "Standard")
      itmX.SubItems(8) = rs!REGISTRO_FECHA
      itmX.SubItems(9) = rs!REGISTRO_USUARIO
      itmX.SubItems(10) = rs!ASEGURADORA_NOMBRE
      itmX.SubItems(11) = rs!Cod_Cuenta_Mask & ""
      itmX.SubItems(12) = rs!Cuenta_Desc & ""
  rs.MoveNext
Loop
rs.Close

Me.MousePointer = vbDefault

Exit Sub

vError:
  Me.MousePointer = vbDefault
 
End Sub


Private Sub sbLoad_DT_Documentos()

On Error GoTo vError
 
Me.MousePointer = vbHourglass

If chkTodos.Value = xtpChecked Then
    strSQL = "exec spCrd_Operacion_Consulta_Documento " & CLng(lblOperacion.Caption) & ",0"
Else
    strSQL = "exec spCrd_Operacion_Consulta_Documento " & CLng(lblOperacion.Caption) & "," & pLinea
End If
 
lswDocumentos.ListItems.Clear
lswValores.ListItems.Clear

Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
  Set itmX = lswDocumentos.ListItems.Add(, , rs!Documento_Desc)
      itmX.SubItems(1) = rs!Cod_Transaccion
      itmX.SubItems(2) = rs!Documento
      itmX.SubItems(3) = rs!Concepto_Desc
      itmX.SubItems(4) = Format(rs!Monto, "Standard")
      itmX.SubItems(5) = Format(rs!Mov_Monto, "Standard")
      itmX.SubItems(6) = Format(rs!Mov_IntCor, "Standard")
      itmX.SubItems(7) = Format(rs!Mov_IntMor, "Standard")
      itmX.SubItems(8) = Format(rs!Mov_Cargos, "Standard")
      itmX.SubItems(9) = Format(rs!Mov_Polizas, "Standard")
      itmX.SubItems(10) = Format(rs!Mov_Principal, "Standard")
      itmX.SubItems(11) = rs!REGISTRO_FECHA
      itmX.SubItems(12) = rs!REGISTRO_USUARIO
      itmX.Tag = rs!TIPO_DOCUMENTO
  rs.MoveNext
Loop
rs.Close

Me.MousePointer = vbDefault

Exit Sub

vError:
  Me.MousePointer = vbDefault
 
End Sub


Private Sub sbLoad_DT_Valores(vTipoDoc As String, vTransaccion As String)

On Error GoTo vError
 
Me.MousePointer = vbHourglass

strSQL = "exec spSys_Documento_Valores '" & vTipoDoc & "','" & vTransaccion & "'"
lswValores.ListItems.Clear

Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
  Set itmX = lswValores.ListItems.Add(, , rs!TIPO_DOCUMENTO)
      itmX.SubItems(1) = rs!Cod_Transaccion
      itmX.SubItems(2) = rs!FORMA_PAGO_DESC
      itmX.SubItems(3) = rs!Num_Referencia
      itmX.SubItems(4) = Format(rs!Monto_Doc, "Standard")
      itmX.SubItems(5) = Format(rs!Monto, "Standard")
      itmX.SubItems(6) = rs!cod_Divisa
      itmX.SubItems(7) = Format(rs!TIPO_CAMBIO, "Standard")
      itmX.SubItems(8) = rs!Saldo_Favor_Id
      itmX.SubItems(9) = rs!Registra_Fecha
      itmX.SubItems(10) = rs!Registra_Usuario
      itmX.SubItems(11) = rs!Referencias & ""
      itmX.SubItems(12) = rs!Cod_Cuenta_Mask & ""
      itmX.SubItems(13) = rs!Cuenta_Desc & ""
  rs.MoveNext
Loop
rs.Close

Me.MousePointer = vbDefault

Exit Sub

vError:
  Me.MousePointer = vbDefault
 
End Sub


Private Sub sbLoad_DT_Ajustes(vOperacion As Long)

On Error GoTo vError
 
Me.MousePointer = vbHourglass

strSQL = "exec spCrd_Operacion_Consulta_Ajustes " & lblOperacion.Caption
lswAjustes.ListItems.Clear

Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
  Set itmX = lswAjustes.ListItems.Add(, , rs!fecha)
      itmX.SubItems(1) = rs!Usuario
      itmX.SubItems(2) = rs!Movimiento_Desc
      itmX.SubItems(3) = rs!Detalle
      itmX.SubItems(4) = rs!Notas & ""
      
  rs.MoveNext
Loop
rs.Close

Me.MousePointer = vbDefault

Exit Sub

vError:
  Me.MousePointer = vbDefault
 
End Sub




Private Sub txtRev_Cuota_GotFocus()
On Error GoTo vError

txtRev_Cuota.Text = CCur(txtRev_Cuota.Text)

vError:

End Sub

Private Sub txtRev_Cuota_LostFocus()
On Error GoTo vError

txtRev_Cuota.Text = Format(CCur(txtRev_Cuota.Text), "Standard")

vError:
End Sub

Private Sub vGrid_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
If vPaso Then Exit Sub

Dim pTipoDoc As String, pNumDoc As String


vGrid.Row = Row
vGrid.Col = 2
If Not IsNumeric(vGrid.Text) Then
    Exit Sub
End If
pLinea = vGrid.Text

If Col = 33 Then
  vGrid.Col = 31
  pTipoDoc = vGrid.Text
  vGrid.Col = 32
  pNumDoc = vGrid.Text
  
  Call sbImprimeRecibo(pNumDoc, pTipoDoc)
End If

If Col = 1 Then
  chkTodos.Value = xtpUnchecked
  gbDetalle.Caption = "Línea: " & pLinea
  If tcMain.SelectedItem = 0 Then
      Call sbLoad_DT_Cargos
  Else
      tcMain.Item(0).Selected = True
  End If
End If


End Sub
