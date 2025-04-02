VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.controls.v22.1.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.shortcutbar.v22.1.0.ocx"
Begin VB.Form frmCajas_Crd_AbonosCtP 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Cajas...Abonos"
   ClientHeight    =   8130
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   9810
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8130
   ScaleWidth      =   9810
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraCuotas 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Cuotas Activas"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   2292
      Left            =   0
      TabIndex        =   1
      Top             =   720
      Width           =   9975
      Begin XtremeSuiteControls.ListView lsw 
         Height          =   1575
         Left            =   120
         TabIndex        =   3
         Top             =   600
         Width           =   9615
         _Version        =   1441793
         _ExtentX        =   16960
         _ExtentY        =   2778
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
         FullRowSelect   =   -1  'True
         Appearance      =   16
      End
      Begin XtremeSuiteControls.PushButton btnAjustes 
         Height          =   288
         Left            =   8040
         TabIndex        =   5
         Top             =   240
         Width           =   1572
         _Version        =   1441793
         _ExtentX        =   2773
         _ExtentY        =   508
         _StockProps     =   79
         Caption         =   "Ajustes"
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
         Appearance      =   6
      End
      Begin XtremeSuiteControls.CheckBox chkMarcaTodas 
         Height          =   252
         Left            =   360
         TabIndex        =   4
         Top             =   240
         Width           =   1812
         _Version        =   1441793
         _ExtentX        =   3196
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Marcar Todas"
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
         Appearance      =   16
      End
   End
   Begin XtremeSuiteControls.GroupBox GroupBox1 
      Height          =   1092
      Left            =   120
      TabIndex        =   57
      Top             =   1680
      Width           =   9492
      _Version        =   1441793
      _ExtentX        =   16743
      _ExtentY        =   1926
      _StockProps     =   79
      Caption         =   "Estado"
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
      Begin VB.Label lblFecUltMovR 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   252
         Left            =   7800
         TabIndex        =   74
         Top             =   720
         Width           =   1212
      End
      Begin VB.Label lblCuotaR 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   252
         Left            =   6360
         TabIndex        =   73
         Top             =   720
         Width           =   1452
      End
      Begin VB.Label lblInteresR 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   252
         Left            =   4800
         TabIndex        =   72
         Top             =   720
         Width           =   1572
      End
      Begin VB.Label lblAmortizaR 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   252
         Left            =   3240
         TabIndex        =   71
         Top             =   720
         Width           =   1572
      End
      Begin VB.Label lblSaldoR 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   252
         Left            =   1320
         TabIndex        =   70
         Top             =   720
         Width           =   1932
      End
      Begin VB.Label lblFecUltMov 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   252
         Left            =   7800
         TabIndex        =   69
         Top             =   480
         Width           =   1212
      End
      Begin VB.Label lblCuota 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   252
         Left            =   6360
         TabIndex        =   68
         Top             =   480
         Width           =   1452
      End
      Begin VB.Label lblInteres 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   252
         Left            =   4800
         TabIndex        =   67
         Top             =   480
         Width           =   1572
      End
      Begin VB.Label lblAmortiza 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   252
         Left            =   3240
         TabIndex        =   66
         Top             =   480
         Width           =   1572
      End
      Begin VB.Label lblSaldo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   252
         Left            =   1320
         TabIndex        =   65
         Top             =   480
         Width           =   1932
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Ult.Mov."
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   252
         Left            =   7800
         TabIndex        =   64
         ToolTipText     =   "Si es menor a la fecha de proceso se Utiliza la Fecha de Proceso"
         Top             =   240
         Width           =   1212
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Cuota"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   252
         Left            =   6360
         TabIndex        =   63
         Top             =   240
         Width           =   1452
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Intereses"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   252
         Left            =   4800
         TabIndex        =   62
         Top             =   240
         Width           =   1572
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Amortización"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   252
         Index           =   0
         Left            =   3240
         TabIndex        =   61
         Top             =   240
         Width           =   1572
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Saldo"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   252
         Left            =   1320
         TabIndex        =   60
         Top             =   240
         Width           =   1932
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Nuevo"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   252
         Index           =   1
         Left            =   120
         TabIndex        =   59
         Top             =   720
         Width           =   1212
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Actual"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   252
         Index           =   0
         Left            =   120
         TabIndex        =   58
         Top             =   480
         Width           =   1212
      End
   End
   Begin XtremeSuiteControls.GroupBox fraDatosAbono 
      Height          =   2172
      Left            =   120
      TabIndex        =   34
      Top             =   3960
      Width           =   9492
      _Version        =   1441793
      _ExtentX        =   16743
      _ExtentY        =   3831
      _StockProps     =   79
      Caption         =   "Abono:"
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
      Begin MSComCtl2.FlatScrollBar FlatScrollBar 
         Height          =   252
         Left            =   2760
         TabIndex        =   35
         Top             =   240
         Width           =   492
         _ExtentX        =   873
         _ExtentY        =   450
         _Version        =   393216
         Arrows          =   65536
         Orientation     =   1638401
      End
      Begin XtremeSuiteControls.CheckBox chkRecalculaCuota 
         Height          =   252
         Left            =   7080
         TabIndex        =   36
         Top             =   1320
         Width           =   1932
         _Version        =   1441793
         _ExtentX        =   3408
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Recalcular Cuota?"
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
         Appearance      =   16
         Alignment       =   1
      End
      Begin XtremeSuiteControls.ComboBox cboDiferenciaApl 
         Height          =   312
         Left            =   6840
         TabIndex        =   37
         Top             =   840
         Width           =   2172
         _Version        =   1441793
         _ExtentX        =   3836
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
      Begin XtremeSuiteControls.FlatEdit txtCuotas 
         Height          =   312
         Left            =   1680
         TabIndex        =   48
         Top             =   240
         Width           =   852
         _Version        =   1441793
         _ExtentX        =   1503
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
         Text            =   "1"
         Alignment       =   2
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit txtDatosAmortiza 
         Height          =   312
         Left            =   1680
         TabIndex        =   49
         Top             =   600
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
         Text            =   "0"
         Alignment       =   1
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtDatosInteres 
         Height          =   312
         Left            =   1680
         TabIndex        =   50
         Top             =   960
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
         Text            =   "0"
         Alignment       =   1
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtDatosCargos 
         Height          =   312
         Left            =   1680
         TabIndex        =   51
         Top             =   1320
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
         Text            =   "0"
         Alignment       =   1
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtDatosAnticipo 
         Height          =   312
         Left            =   1680
         TabIndex        =   52
         Top             =   1680
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
         Text            =   "0"
         Alignment       =   1
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtPolizas 
         Height          =   312
         Left            =   5040
         TabIndex        =   53
         Top             =   600
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
         Text            =   "0"
         Alignment       =   1
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtDiferencia 
         Height          =   312
         Left            =   5040
         TabIndex        =   54
         Top             =   1680
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
         Text            =   "0"
         Alignment       =   1
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtTotalCancela 
         Height          =   312
         Left            =   5040
         TabIndex        =   55
         Top             =   960
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
         Text            =   "0"
         Alignment       =   1
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtTotalPagar 
         Height          =   312
         Left            =   5040
         TabIndex        =   56
         Top             =   1320
         Width           =   1692
         _Version        =   1441793
         _ExtentX        =   2984
         _ExtentY        =   550
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   14737632
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   "0"
         BackColor       =   14737632
         Alignment       =   1
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin VB.Label Label26 
         BackStyle       =   0  'Transparent
         Caption         =   "Ctas.Pólizas..:"
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
         Left            =   3720
         TabIndex        =   47
         Top             =   600
         Width           =   1212
      End
      Begin VB.Label Label26 
         BackStyle       =   0  'Transparent
         Caption         =   "C. Pago Anticipado"
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
         Index           =   1
         Left            =   240
         TabIndex        =   46
         Top             =   1680
         Width           =   1455
      End
      Begin VB.Label Label27 
         BackStyle       =   0  'Transparent
         Caption         =   "Diferencia ...:"
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
         Left            =   3720
         TabIndex        =   45
         Top             =   1680
         Width           =   972
      End
      Begin VB.Label Label26 
         BackStyle       =   0  'Transparent
         Caption         =   "Cargos ..:"
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
         Index           =   2
         Left            =   240
         TabIndex        =   44
         Top             =   1320
         Width           =   1335
      End
      Begin VB.Label Label24 
         BackStyle       =   0  'Transparent
         Caption         =   "Compromiso"
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
         Left            =   3720
         TabIndex        =   43
         Top             =   960
         Width           =   1212
      End
      Begin VB.Label Label24 
         BackStyle       =   0  'Transparent
         Caption         =   "Total a Pagar"
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
         Left            =   3720
         TabIndex        =   42
         Top             =   1320
         Width           =   1212
      End
      Begin VB.Label Label25 
         BackStyle       =   0  'Transparent
         Caption         =   "Amortización ..:"
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
         Left            =   240
         TabIndex        =   41
         Top             =   600
         Width           =   1572
      End
      Begin VB.Label Label26 
         BackStyle       =   0  'Transparent
         Caption         =   "Intereses ..:"
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
         Left            =   240
         TabIndex        =   40
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label27 
         BackStyle       =   0  'Transparent
         Caption         =   "No. Cuotas"
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
         Left            =   240
         TabIndex        =   39
         Top             =   240
         Width           =   1572
      End
      Begin VB.Label Label23 
         BackStyle       =   0  'Transparent
         Caption         =   "Aplicar diferencias como..:"
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
         Left            =   6840
         TabIndex        =   38
         Top             =   600
         Width           =   2052
      End
   End
   Begin VB.Timer TimerX 
      Interval        =   10
      Left            =   9480
      Top             =   0
   End
   Begin MSComctlLib.StatusBar StatusBarX 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   7875
      Width           =   9810
      _ExtentX        =   17304
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   7832
            MinWidth        =   7832
            Object.ToolTipText     =   "Oficina"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   4304
            MinWidth        =   4304
            Object.ToolTipText     =   "Linea"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   4304
            MinWidth        =   4304
            Object.ToolTipText     =   "Recurso"
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
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   9360
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCajas_Crd_AbonosCtP.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCajas_Crd_AbonosCtP.frx":6862
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCajas_Crd_AbonosCtP.frx":703A
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCajas_Crd_AbonosCtP.frx":7807
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin XtremeSuiteControls.GroupBox fraFormaPago 
      Height          =   1572
      Left            =   120
      TabIndex        =   6
      Top             =   6240
      Width           =   9492
      _Version        =   1441793
      _ExtentX        =   16743
      _ExtentY        =   2773
      _StockProps     =   79
      BackColor       =   16777215
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      BorderStyle     =   1
      Begin XtremeSuiteControls.ComboBox cboTipoDoc 
         Height          =   312
         Left            =   1200
         TabIndex        =   7
         Top             =   240
         Width           =   2772
         _Version        =   1441793
         _ExtentX        =   4895
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
      Begin XtremeSuiteControls.FlatEdit txtTotalCajas 
         Height          =   312
         Left            =   4920
         TabIndex        =   8
         Top             =   240
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
         Alignment       =   1
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtNotas 
         Height          =   792
         Left            =   1200
         TabIndex        =   9
         Top             =   600
         Width           =   5412
         _Version        =   1441793
         _ExtentX        =   9546
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
         MultiLine       =   -1  'True
         ScrollBars      =   2
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.PushButton btnCajas 
         Height          =   792
         Index           =   0
         Left            =   6720
         TabIndex        =   10
         Top             =   600
         Width           =   852
         _Version        =   1441793
         _ExtentX        =   1503
         _ExtentY        =   1397
         _StockProps     =   79
         Caption         =   "Pago"
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
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         Picture         =   "frmCajas_Crd_AbonosCtP.frx":80D3
         TextImageRelation=   1
      End
      Begin XtremeSuiteControls.PushButton btnCajas 
         Height          =   792
         Index           =   1
         Left            =   7680
         TabIndex        =   11
         Top             =   600
         Width           =   852
         _Version        =   1441793
         _ExtentX        =   1503
         _ExtentY        =   1397
         _StockProps     =   79
         Caption         =   "Aplicar"
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
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         Picture         =   "frmCajas_Crd_AbonosCtP.frx":8580
         TextImageRelation=   1
      End
      Begin XtremeSuiteControls.PushButton btnCajas 
         Height          =   792
         Index           =   2
         Left            =   8520
         TabIndex        =   12
         Top             =   600
         Width           =   972
         _Version        =   1441793
         _ExtentX        =   1714
         _ExtentY        =   1397
         _StockProps     =   79
         Caption         =   "Cancelar"
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
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         Picture         =   "frmCajas_Crd_AbonosCtP.frx":8D58
         TextImageRelation=   1
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Documento ..:"
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
         TabIndex        =   15
         Top             =   240
         Width           =   1452
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Notas ..:"
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
         Left            =   120
         TabIndex        =   14
         Top             =   600
         Width           =   1452
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Total ..:"
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
         Left            =   4080
         TabIndex        =   13
         Top             =   240
         Width           =   1092
      End
   End
   Begin XtremeSuiteControls.GroupBox fraAbono 
      Height          =   972
      Left            =   120
      TabIndex        =   16
      Top             =   2880
      Width           =   9492
      _Version        =   1441793
      _ExtentX        =   16743
      _ExtentY        =   1714
      _StockProps     =   79
      BackColor       =   16777215
      Appearance      =   16
      BorderStyle     =   1
      Begin XtremeSuiteControls.DateTimePicker dtpFechaCancelacion 
         Height          =   312
         Left            =   6360
         TabIndex        =   17
         Top             =   600
         Width           =   1332
         _Version        =   1441793
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
      Begin XtremeSuiteControls.PushButton optAbono 
         Height          =   372
         Index           =   0
         Left            =   1560
         TabIndex        =   18
         Top             =   120
         Width           =   1572
         _Version        =   1441793
         _ExtentX        =   2773
         _ExtentY        =   656
         _StockProps     =   79
         Caption         =   "Ordinario"
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
         Appearance      =   6
      End
      Begin XtremeSuiteControls.PushButton optAbono 
         Height          =   372
         Index           =   1
         Left            =   3120
         TabIndex        =   19
         Top             =   120
         Width           =   1572
         _Version        =   1441793
         _ExtentX        =   2773
         _ExtentY        =   656
         _StockProps     =   79
         Caption         =   "Extraordinario"
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
         Appearance      =   6
      End
      Begin XtremeSuiteControls.PushButton optAbono 
         Height          =   372
         Index           =   2
         Left            =   4800
         TabIndex        =   20
         Top             =   120
         Width           =   1572
         _Version        =   1441793
         _ExtentX        =   2773
         _ExtentY        =   656
         _StockProps     =   79
         Caption         =   "Cancelación"
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
         Appearance      =   6
      End
      Begin XtremeSuiteControls.PushButton optAbono 
         Height          =   372
         Index           =   3
         Left            =   6480
         TabIndex        =   23
         Top             =   120
         Width           =   1572
         _Version        =   1441793
         _ExtentX        =   2773
         _ExtentY        =   656
         _StockProps     =   79
         Caption         =   "Adelanto"
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
         Appearance      =   6
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption2 
         Height          =   372
         Left            =   0
         TabIndex        =   22
         Top             =   120
         Width           =   9492
         _Version        =   1441793
         _ExtentX        =   16743
         _ExtentY        =   656
         _StockProps     =   14
         Caption         =   "Tipo de Abono:"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   8.93
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label lblFechaCancelacion 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha de Abono (Real) por parte del cliente para cancelación...:"
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
         Left            =   1560
         TabIndex        =   21
         Top             =   600
         Width           =   4692
      End
   End
   Begin XtremeSuiteControls.FlatEdit txtOperacion 
      Height          =   372
      Left            =   2040
      TabIndex        =   24
      Top             =   120
      Width           =   2052
      _Version        =   1441793
      _ExtentX        =   3619
      _ExtentY        =   656
      _StockProps     =   77
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   12
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
   Begin XtremeSuiteControls.FlatEdit txtProceso 
      Height          =   372
      Left            =   4080
      TabIndex        =   25
      Top             =   120
      Width           =   2052
      _Version        =   1441793
      _ExtentX        =   3619
      _ExtentY        =   656
      _StockProps     =   77
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   12
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
   Begin XtremeSuiteControls.PushButton cmdAplicar 
      Height          =   375
      Left            =   7800
      TabIndex        =   26
      Top             =   120
      Visible         =   0   'False
      Width           =   1095
      _Version        =   1441793
      _ExtentX        =   1931
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Aplicar"
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
      UseVisualStyle  =   -1  'True
      Appearance      =   17
   End
   Begin XtremeSuiteControls.FlatEdit txtNombre 
      Height          =   315
      Left            =   3720
      TabIndex        =   27
      Top             =   840
      Width           =   5415
      _Version        =   1441793
      _ExtentX        =   9551
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
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtDescripcion 
      Height          =   315
      Left            =   3720
      TabIndex        =   28
      Top             =   1200
      Width           =   4815
      _Version        =   1441793
      _ExtentX        =   8493
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
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtCedula 
      Height          =   315
      Left            =   1800
      TabIndex        =   29
      Top             =   840
      Width           =   1935
      _Version        =   1441793
      _ExtentX        =   3413
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
      Alignment       =   2
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtCodigo 
      Height          =   315
      Left            =   1800
      TabIndex        =   30
      Top             =   1200
      Width           =   1935
      _Version        =   1441793
      _ExtentX        =   3413
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
      Alignment       =   2
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtOpex 
      Height          =   315
      Left            =   8520
      TabIndex        =   31
      Top             =   1200
      Width           =   615
      _Version        =   1441793
      _ExtentX        =   1080
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
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Línea"
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
      Left            =   240
      TabIndex        =   33
      Top             =   1200
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
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
      Index           =   1
      Left            =   240
      TabIndex        =   32
      Top             =   840
      Width           =   1575
   End
   Begin VB.Image imgDocumento 
      Height          =   240
      Left            =   6240
      Picture         =   "frmCajas_Crd_AbonosCtP.frx":9525
      ToolTipText     =   "Confección del Documento en Caso de Error!"
      Top             =   240
      Width           =   240
   End
   Begin VB.Label Label15 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Operación"
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
      Height          =   312
      Left            =   480
      TabIndex        =   2
      Top             =   120
      Width           =   1332
   End
   Begin VB.Image imgBanner 
      Height          =   732
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12012
   End
End
Attribute VB_Name = "frmCajas_Crd_AbonosCtP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vOperacion As Long, vCuotasDeducidas As Integer, vCuotasDirectas As Integer
Dim vInteres As Currency, vPlazo As Integer, vSaldoMes As Currency, vUltimoRecibo As Long
Dim vRetencion As Boolean, vBaseCalculo As String, vPrideduc As Long, vAnticipoPorc As Currency, vAnticipoMeses As Integer
Dim vDiasActivo As Long, vFechaHoy As Date, vScroll As Boolean
Dim pCharRelleno As String, mControl As Currency


Private Sub btnAjustes_Click()
  GLOBALES.gTag = txtOperacion.Text
  frmCR_MoraCargosAjustes.Show vbModal
  
  'Verifica si recibio modificaciones, en cuyo caso procede a actualizar datos en pantalla
  If GLOBALES.gTag2 = 1 Then
    Call sbConsultaOperacion
  End If
End Sub

Private Sub btnCajas_Click(Index As Integer)
On Error GoTo vError


'Posicion Final
Call txtTotalPagar_LostFocus

Select Case Index
  Case 2 'Cancelar
     Call sbConsultaOperacion
     
  Case 0 'Desgloce
        If Not IsNumeric(txtTotalPagar.Text) Then txtTotalPagar.Text = 0
        If Not ModuloCajas.mConceptoValida Then
           MsgBox "Esta caja no está autorizada para registrar movimientos a este línea de crédito", vbExclamation
           Exit Sub
        End If
        
        ModuloCajas.mTotalAplicar = CCur(txtTotalPagar.Text)
        
        If ModuloCajas.mTotalAplicar = 0 Then
            MsgBox "No se ha especificado ningún monto a detallar?", vbExclamation
            Exit Sub
        End If
        
        ModuloCajas.mServicio = "Abonos a Operación de Crédito"
        
        Call sbFormsCall("frmCajas_DetallePago", vbModal, 0, 0, False, Me)
        
        txtTotalCajas.Text = Format(ModuloCajas.mTotalDetallado, "Standard")
        
        
        If txtTotalCajas.Text <> txtTotalPagar.Text Then
           txtTotalCajas.BackColor = vbRed
        Else
           txtTotalCajas.BackColor = vbWhite
        End If

  Case 1   'Aplicar
    Call CmdAplicar_Click
End Select

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub cboDiferenciaApl_Click()

If Not cboDiferenciaApl.Enabled Then Exit Sub

If cboDiferenciaApl.Text = "Abono Extraordinario" Then
   chkRecalculaCuota.Enabled = True
Else
   chkRecalculaCuota.Enabled = False
End If

End Sub


Private Sub chkMarcaTodas_Click()
Dim i As Integer

For i = 1 To lsw.ListItems.Count
  lsw.ListItems.Item(i).Checked = chkMarcaTodas.Value
Next

If lsw.ListItems.Count > 0 Then
    Call lsw_ItemCheck(lsw.ListItems.Item(1))
End If

End Sub

Private Sub chkRecalculaCuota_Click()

If vRetencion Then
   chkRecalculaCuota.Value = vbUnchecked
   MsgBox "Las retenciones no se pueden Ajustar para Recálculos, verifique...", vbExclamation
   Exit Sub
End If

Call txtTotalPagar_Change

End Sub

Private Sub sbAbono()
Dim strSQL As String, rs As New ADODB.Recordset
Dim vNumDoc As String, vCuenta As String
Dim vTipoDoc As String, vFecha As Date
Dim i As Integer, vExtraOrdinario As Boolean


Me.MousePointer = vbHourglass

On Error GoTo vError


vFecha = fxFechaServidor
vExtraOrdinario = False

vTipoDoc = cboTipoDoc.ItemData(cboTipoDoc.ListIndex)

vNumDoc = fxDocumentoConsecutivo(vTipoDoc)


'Inicia Transaccion
glogon.Conection.BeginTrans

 
Select Case True
  Case optAbono(0).Checked  'Abono Ordinario
  
        If Not cboDiferenciaApl.Enabled Then
            strSQL = "exec spCrdPlanPagoAbonoOrdinario " & vOperacion & ",'CRD001','" & glogon.Usuario & "','" & vTipoDoc _
                   & "','" & vNumDoc & "'," & CCur(txtTotalCajas.Text) & ",'" & Format(vFecha, "yyyy/mm/dd") & "',''"
            Call ConectionExecute(strSQL)
        Else
            strSQL = "exec spCrdPlanPagoAbonoOrdinario " & vOperacion & ",'CRD001','" & glogon.Usuario & "','" & vTipoDoc _
                   & "','" & vNumDoc & "'," & CCur(txtTotalCancela.Text) & ",'" & Format(vFecha, "yyyy/mm/dd") & "',''"
            Call ConectionExecute(strSQL)
           
           Select Case cboDiferenciaApl.Text
             Case "Adelanto de Cuota"
                strSQL = "exec spCrdPlanPagoAbonoOrdinario " & vOperacion & ",'CRD004','" & glogon.Usuario & "','" & vTipoDoc _
                       & "','" & vNumDoc & "'," & Abs(CCur(txtDiferencia.Text)) & ",'" _
                       & Format(vFecha, "yyyy/mm/dd") & "',''"
                Call ConectionExecute(strSQL)

             Case "Abono Extraordinario"
                'Calcula Datos del Abono Extraordinario (Dias, Intereses, Cargos, Principal)
                strSQL = "exec spCrdPlanPagosInfoExtraordinario " & vOperacion & "," & Abs(CCur(txtDiferencia.Text)) & ",'" & Format(dtpFechaCancelacion.Value, "yyyy/mm/dd") & "'"
                Call OpenRecordSet(rs, strSQL)
                    'Aplica Cargo por Anticipo
                    If rs!Cargos > 0 Then
                       strSQL = "exec spCrdOperacionCargoAdd " & vOperacion & "," & rs!Cargos & ",'" & GLOBALES.gOficinaUnidad _
                              & "','" & GLOBALES.gOficinaCentroCosto & "','Pago Anticipado','" & glogon.Usuario & "','CA','','',0"
                       Call ConectionExecute(strSQL)
                    End If
                    'Aplica Abono Extraordinario
                    strSQL = "exec spCrdPlanPagoAbonoEC " & vOperacion & ",'CRD002','" & glogon.Usuario & "','" & vTipoDoc _
                           & "','" & vNumDoc & "'," & rs!Dias & "," & rs!Intereses & "," & rs!Principal _
                           & "," & rs!Cargos & ",'" & Format(vFecha, "yyyy/mm/dd") & "',''," & chkRecalculaCuota.Value
                    Call ConectionExecute(strSQL)
                rs.Close
                
                vExtraOrdinario = True
           End Select
        
        End If
  
  Case optAbono(1).Checked 'Abono Extraordinario
        'Elimina Cuotas Activas, Registra Abono y Recalcula Plan de Pagos
        'Se Supone que solo queda una cuota activa para poder realizar un ab. extraordinario
        
        If CCur(txtDatosAnticipo.Text) > 0 Then
           strSQL = "exec spCrdOperacionCargoAdd " & vOperacion & "," & CCur(txtDatosAnticipo.Text) & ",'" & GLOBALES.gOficinaUnidad _
                  & "','" & GLOBALES.gOficinaCentroCosto & "','Pago Anticipado','" & glogon.Usuario & "','CA','','',0"
           Call ConectionExecute(strSQL)
        End If
        
        strSQL = "exec spCrdPlanPagoAbonoEC " & vOperacion & ",'CRD002','" & glogon.Usuario & "','" & vTipoDoc _
               & "','" & vNumDoc & "'," & vDiasActivo & "," & CCur(txtDatosInteres.Text) & "," & CCur(txtDatosAmortiza.Text) _
               & ",0,'" & Format(vFecha, "yyyy/mm/dd") & "',''," & chkRecalculaCuota.Value
        Call ConectionExecute(strSQL)

        vExtraOrdinario = True
        
  Case optAbono(2).Checked 'Cancelacion
        'Actualiza el estado de la morosidad
'        strSQL = "exec spCrdPlanPagosMoraActualizaOp " & vOperacion & ",'" & Format(vFecha, "yyyy/mm/dd") & "'"
        
        strSQL = "exec spCrdPlanPagosMoraActualizaOp " & vOperacion & ",'" & Format(dtpFechaCancelacion.Value, "yyyy/mm/dd") & "'"
        Call ConectionExecute(strSQL)
        
        If CCur(txtDatosAnticipo.Text) > 0 Then
           strSQL = "exec spCrdOperacionCargoAdd " & vOperacion & "," & CCur(txtDatosAnticipo.Text) & ",'" & GLOBALES.gOficinaUnidad _
                  & "','" & GLOBALES.gOficinaCentroCosto & "','Cancelacion Anticipada','" & glogon.Usuario & "','CA','','',0"
           Call ConectionExecute(strSQL)
        End If
'        strSQL = "exec spCrdPlanPagoAbonoCancelacion " & vOperacion & ",'CRD003','" & glogon.Usuario & "','" & vTipoDoc _
'               & "','" & vNumDoc & "'," & CCur(txtTotalCajas.Text) & ",'" & Format(vFecha, "yyyy/mm/dd") & "',''"
        
        strSQL = "exec spCrdPlanPagoAbonoCancelacion " & vOperacion & ",'CRD003','" & glogon.Usuario & "','" & vTipoDoc _
               & "','" & vNumDoc & "'," & CCur(txtTotalCajas.Text) & ",'" & Format(dtpFechaCancelacion.Value, "yyyy/mm/dd") & "',''"
        Call ConectionExecute(strSQL)
  
  
  Case optAbono(3).Checked 'Adelanto de Cuotas
       'Activa Nuevas Cuotas y luego las abona
'        strSQL = "exec spCrdPlanPagoAbonoOrdinario " & vOperacion & ",'CRD004','" & glogon.Usuario & "','" & vTipoDoc _
'               & "','" & vNumDoc & "'," & CCur(txtTotalCajas.Text) & ",'" & Format(vFecha, "yyyy/mm/dd") & "',''"
'        Call ConectionExecute(strSQL)
       

        If Not cboDiferenciaApl.Enabled Then
            strSQL = "exec spCrdPlanPagoAbonoOrdinario " & vOperacion & ",'CRD004','" & glogon.Usuario & "','" & vTipoDoc _
                   & "','" & vNumDoc & "'," & CCur(txtTotalCajas.Text) & ",'" & Format(vFecha, "yyyy/mm/dd") & "',''"
            Call ConectionExecute(strSQL)
        Else
            strSQL = "exec spCrdPlanPagoAbonoOrdinario " & vOperacion & ",'CRD004','" & glogon.Usuario & "','" & vTipoDoc _
                   & "','" & vNumDoc & "'," & CCur(txtTotalCancela.Text) & ",'" & Format(vFecha, "yyyy/mm/dd") & "',''"
            Call ConectionExecute(strSQL)
           
           Select Case cboDiferenciaApl.Text
             Case "Adelanto de Cuota"
                strSQL = "exec spCrdPlanPagoAbonoOrdinario " & vOperacion & ",'CRD004','" & glogon.Usuario & "','" & vTipoDoc _
                       & "','" & vNumDoc & "'," & Abs(CCur(txtDiferencia.Text)) & ",'" _
                       & Format(vFecha, "yyyy/mm/dd") & "',''"
                Call ConectionExecute(strSQL)

             Case "Abono Extraordinario"
                'Calcula Datos del Abono Extraordinario (Dias, Intereses, Cargos, Principal)
                strSQL = "exec spCrdPlanPagosInfoExtraordinario " & vOperacion & "," & Abs(CCur(txtDiferencia.Text)) _
                       & ",'" & Format(dtpFechaCancelacion.Value, "yyyy/mm/dd") & "'"
                Call OpenRecordSet(rs, strSQL)
                    'Aplica Cargo por Anticipo
                    If rs!Cargos > 0 Then
                       strSQL = "exec spCrdOperacionCargoAdd " & vOperacion & "," & rs!Cargos & ",'" & GLOBALES.gOficinaUnidad _
                              & "','" & GLOBALES.gOficinaCentroCosto & "','Pago Anticipado','" & glogon.Usuario & "','CA','','',0"
                       Call ConectionExecute(strSQL)
                    End If
                    'Aplica Abono Extraordinario
                    strSQL = "exec spCrdPlanPagoAbonoEC " & vOperacion & ",'CRD002','" & glogon.Usuario & "','" & vTipoDoc _
                           & "','" & vNumDoc & "'," & rs!Dias & "," & rs!Intereses & "," & rs!Principal _
                           & "," & rs!Cargos & ",'" & Format(vFecha, "yyyy/mm/dd") & "',''," & chkRecalculaCuota.Value
                    Call ConectionExecute(strSQL)
                rs.Close
                
                vExtraOrdinario = True
           End Select
        
        End If



End Select


'Cierra Transaccion
glogon.Conection.CommitTrans

'Indica si debe reprocesar el Plan de Pagos por registro de Abonos Extraordinario
If vExtraOrdinario Then
        strSQL = "exec spCrdPlanPagos " & vOperacion _
               & Space(10) & "exec spCrdPlanPagosActivaCuota " & vOperacion & ", 0" _
               & Space(10) & "exec spCrdPlanPagosMoraActualizaOp " & vOperacion
        Call ConectionExecute(strSQL)
End If

'Genera el Comprobante
Select Case True
  Case optAbono(0).Checked 'Abono Ordinario
      Call Bitacora("Registra", "Abono Ordinario a la Operacion : " & vOperacion)
      Call sbDocumentoAbono("ABONO ORDINARIO", vTipoDoc, CStr(vNumDoc), "CRD001", vCuenta)
  Case optAbono(1).Checked 'Abono Extraordinario
      Call Bitacora("Registra", "Abono ExtraOrd. " & IIf((chkRecalculaCuota.Value = 1), "Con Recal.", "Sin Recal") & " a la Op.: " & vOperacion)
      Call sbDocumentoAbono("ABONO EXTRAORDINARIO", vTipoDoc, CStr(vNumDoc), "CRD002", vCuenta)
  Case optAbono(2).Checked 'Abono De Cancelacion
      Call Bitacora("Registra", "Cancelación de la Operacion : " & vOperacion)
      Call sbDocumentoAbono("CANCELACION DE DEUDA", vTipoDoc, CStr(vNumDoc), "CRD003", vCuenta)
  Case optAbono(3).Checked 'Adelanto de Cuotas
      Call Bitacora("Registra", "Adelanto de Cuotas de la Operacion : " & vOperacion)
      Call sbDocumentoAbono("ADELANTO DE CUOTAS", vTipoDoc, CStr(vNumDoc), "CRD004", vCuenta)
End Select


'IMPRIMIR RECIBO
Call sbImprimeRecibo(vNumDoc, vTipoDoc)

Me.MousePointer = vbDefault

strSQL = " - Abono aplicado, con : " & cboTipoDoc.Text & " ...No.: " & vNumDoc & vbCrLf _
       & " - Desea Realizar Otra Transacción a esta Operación ?"

i = MsgBox(strSQL, vbYesNo)
If i = vbYes Then
    Call sbConsultaOperacion
    txtTotalCajas.Text = 0
Else
    Unload Me
End If

Me.MousePointer = vbDefault

Exit Sub

vError:
 Me.MousePointer = vbDefault
 glogon.Conection.RollbackTrans
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Function fxVerifica() As Boolean
Dim strSQL As String, rs As New ADODB.Recordset
Dim vControl As Currency, vCajasMov As Boolean
Dim vMensaje As String, i As Integer

vMensaje = ""

Call sbSIFCleanTxtInject(txtNotas)

strSQL = "select dbo.fxCrd_Operacion_Control(" & txtOperacion.Text & ") as 'Control'" _
       & ", dbo.fxCrd_Operacion_Movimientos_Cajas_Acepta(" & txtOperacion.Text & ") as 'CajaMov' "
Call OpenRecordSet(rs, strSQL)
  vControl = rs!Control
  vCajasMov = IIf((rs!CajaMov = 1), True, False)
rs.Close

'Control de Cambios
If vControl <> mControl Then
      vMensaje = vMensaje & "- Esta Operación ha sido cambiada por otro proceso, vuelva a consultarla!" & vbCrLf
End If

If Not vCajasMov Then
      vMensaje = vMensaje & "- Esta Operación -> No Permite Movimientos en Cajas! Puede que sea recaudo de ahorros o porque el código de linea no lo admite, revise!" & vbCrLf
End If

'Verifica el proceso
If txtProceso.Tag = "J" Then
   If Not fxCajasAbonosCbrJud(ModuloCajas.mCaja, ModuloCajas.mUsuario) Then
      vMensaje = vMensaje & "- Esta CAJA no cuenta con permisos para realizar abonos a Creditos en Cobro Judicial, verifique..." & vbCrLf
   End If
End If

'Verifica que la diferencia del Monto a Cancelar no supere el Saldo
If CCur(txtDiferencia.Text) < 0 Then
 If CCur(lblSaldoR.Caption) + CCur(txtDiferencia.Text) < 0 Then
      vMensaje = vMensaje & "- La diferencia supera el saldo!, verifique..." & vbCrLf
 End If
End If

'Verificar Congelamiento
If fxgCongelamiento(txtCedula, "per_abono_cajas") Then
  vMensaje = vMensaje & "- Esta Persona se encuentra CONGELADA, verifique..." & vbCrLf
End If

If vOperacion = 0 Then
  vMensaje = vMensaje & "- Número de Operacion no es válido..." & vbCrLf
End If
 
 
'Verifica Saldo Actual
If Not fxCrdSaldoVerifica(vOperacion, CCur(lblSaldo.Caption)) Then
   vMensaje = vMensaje & "- Esta Operación ha sido modificada, actualice los datos nuevamente antes de realizar el abono..." & vbCrLf
End If
 
If Not vRetencion Then
    If CCur(txtDatosAmortiza) > CCur(lblSaldo.Caption) Then
       vMensaje = vMensaje & "- La Amortización es mayor al Saldo Actual..." & vbCrLf
    End If
Else
    If vPlazo < 999 Then
        If CCur(txtDatosAmortiza.Text) > CCur(lblSaldo.Caption) Then
            vMensaje = vMensaje & "- La Amortización es mayor que el Remanente a Recaudar : " & lblSaldo.Caption & vbCrLf
         End If
    Else
    
'        If CCur(txtDatosAmortiza.Text) > ((CCur(lblCuota.Caption) * vPlazo) - CCur(lblAmortiza.Caption)) Then
'            vMensaje = vMensaje & "- La Amortización es mayor que el Remanente a Recaudar : " _
'                  & ((CCur(lblCuota.Caption) * vPlazo) - CCur(lblAmortiza.Caption)) & vbCrLf
'         End If
    End If
End If

If Not IsNumeric(txtTotalPagar.Text) Then
  vMensaje = vMensaje & "- El total a pagar no es un dato válido...verifique...!" & vbCrLf
Else
 If CCur(txtTotalPagar.Text) <= 0 Then
      vMensaje = vMensaje & "- El total a pagar no es un dato válido...verifique...!" & vbCrLf
 End If
End If

 If CCur(txtTotalCajas.Text) <= 0 Then
      vMensaje = vMensaje & "- Los valores Recibidos en Cajas no son válidos...verifique...!" & vbCrLf
 End If

 If Len(vMensaje) = 0 Then
    If CCur(txtTotalCajas.Text) <> CCur(txtTotalPagar.Text) Then
         vMensaje = vMensaje & "- Los valores Recibidos en Cajas son diferentes al monto a Pagar establecido para el Abono...!" & vbCrLf
    End If
 End If

 'Validacion para Abonos Extraordinarios
 If optAbono(1).Checked And Len(vMensaje) = 0 Then
    If CCur(txtDiferencia.Text) <> 0 Then
       vMensaje = vMensaje & "- El Monto detallado en formas de pago no cubre el compromiso de pago. SOLUCION: Copie el Monto detallado en el monto del abono extraordinario...!" & vbCrLf
    End If
 End If


 'Validacion para Cancelacion de Deudas
 If optAbono(2).Checked And Len(vMensaje) = 0 Then
    If CCur(txtDiferencia.Text) <> 0 Then
       vMensaje = vMensaje & "- El Monto detallado en formas de pago no cubre el compromiso de cancelación...!" & vbCrLf
    End If
 End If


'Abono Ordinario (Verificar Secuencia de Check's)
If optAbono.Item(0).Checked Then
 For i = 1 To lsw.ListItems.Count
   If i = 1 And Not lsw.ListItems.Item(i).Checked Then
      vMensaje = vMensaje & "- No se ha especificado un orden válido de aplicación de cuotas...!" & vbCrLf
      Exit For
   End If
   
   If i > 1 Then
        If lsw.ListItems.Item(i).Checked And Not lsw.ListItems.Item(i - 1).Checked Then
               vMensaje = vMensaje & "- No se ha especificado un orden válido de aplicación de cuotas...!" & vbCrLf
               Exit For
        End If
   End If
 Next
End If

If fxCajasAperturaEstado = "C" Then
   vMensaje = vMensaje & "- La apertura ..:" & ModuloCajas.mApertura & " de esta caja ha sido cerrada!" & vbCrLf
End If

If Len(vMensaje) = 0 Then
  fxVerifica = True
Else
  fxVerifica = False
  MsgBox vMensaje, vbExclamation
End If

End Function



Private Sub sbReporte(vTitulo As String)
If vOperacion = 0 Then Exit Sub

Me.MousePointer = vbHourglass

With frmContenedor.Crt
 .Reset
 .WindowShowGroupTree = True
 .WindowShowPrintSetupBtn = True
 .WindowShowRefreshBtn = True
 .WindowShowSearchBtn = True
 .WindowState = crptMaximized
 .WindowTitle = "Módulo de Crédito"
 
 .ReportFileName = SIFGlobal.fxPathReportes("Credito_BoletaAbono.rpt")
 
 .Formulas(1) = "empresa='" & GLOBALES.gstrNombreEmpresa & "'"
 .Formulas(2) = "fecha='" & Format(fxFechaServidor, "dd/mm/yyyy") & "'"
 .Formulas(3) = "usuario='" & glogon.Usuario & "'"
 If optAbono(0).Checked = True Then
  .Formulas(4) = "tipo_abono='ABONO ORDINARIO : CUOTAS: " & txtCuotas & "'"
 Else
  .Formulas(4) = "tipo_abono='ABONO EXTRAORDINARIO'"
 End If
 .Formulas(5) = "saldo_actual='" & lblSaldo.Caption & "'"
 .Formulas(6) = "amortizacion='" & Me.lblAmortiza.Caption & "'"
 .Formulas(7) = "interesc='" & lblInteres.Caption & "'"
 .Formulas(8) = "fecult='" & Format(lblFecUltMov.Caption, "####-##") & "'"

 .Formulas(9) = "saldo_res='" & lblSaldoR.Caption & "'"
 .Formulas(10) = "amortizacion_res='" & Me.lblAmortizaR.Caption & "'"
 .Formulas(11) = "interesc_res='" & lblInteresR.Caption & "'"
 .Formulas(12) = "fecult_res='" & Format(lblFecUltMovR.Caption, "####-##") & "'"
 
 .Formulas(13) = "abono_amortizacion='" & txtDatosAmortiza & "'"
 .Formulas(14) = "abono_interes='" & txtDatosInteres.Text & "'"
 .Formulas(15) = "abono_total='" & txtTotalPagar.Text & "'"
 
 .Formulas(16) = "titulo='" & vTitulo & "'"
 .Formulas(17) = "operacion='" & vOperacion & "'"
 .Formulas(18) = "cedula='" & txtCedula & " - " & txtNombre & "'"
 .Formulas(19) = "codigo='" & txtCodigo & " - " & txtDescripcion.Text & "'"
 
 .PrintReport
End With
Me.MousePointer = vbDefault

End Sub

Private Sub cmdReporte_Click()

Call sbReporte("ABONO A REALIZAR")

End Sub

Private Sub CmdAplicar_Click()
Dim iRespuesta As Integer

If Not fxVerifica Then Exit Sub

 iRespuesta = MsgBox("Esta seguro de realizar el abono a esta Operación " & vOperacion, vbYesNo)
 If iRespuesta = vbYes Then
  
  Call sbAbono

 Else 'Respuesta
  
  MsgBox "Transacción Cancelada...", vbInformation
 
 End If

End Sub

Private Sub dtpFechaCancelacion_Change()

If dtpFechaCancelacion.Enabled Then
   'Refresca información base para Cancelación y/o Abonos Extraordinarios
   Select Case True
      Case optAbono.Item(1).Checked 'Abono Extraordinario
            Call optAbono_Click(1)
      Case optAbono.Item(2).Checked 'Cancelación
            Call optAbono_Click(2)
   End Select
End If

End Sub

Private Sub FlatScrollBar_Change()
Dim vNumCuota As Integer

On Error GoTo vError


vNumCuota = txtCuotas.Text

If vScroll Then
    If FlatScrollBar.Value = 1 Then
       vNumCuota = vNumCuota + 1
    Else
       vNumCuota = vNumCuota - 1
    End If
End If

If vNumCuota <= 0 Then vNumCuota = 1

txtCuotas.Text = vNumCuota

vScroll = False
    FlatScrollBar.Value = 0
vScroll = True

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Function fxValidaCriterio(pCadena As String) As Boolean
Dim vResultado As Boolean, vMensaje As String

pCadena = UCase(pCadena)

vResultado = True
If InStr(1, pCadena, "SELECT") > 0 And vResultado Then vResultado = False
If InStr(1, pCadena, "DELETE") > 0 And vResultado Then vResultado = False
If InStr(1, pCadena, "UPDATE") > 0 And vResultado Then vResultado = False
If InStr(1, pCadena, "INSERT") > 0 And vResultado Then vResultado = False
If InStr(1, pCadena, "EXEC") > 0 And vResultado Then vResultado = False
If InStr(1, pCadena, "DROP") > 0 And vResultado Then vResultado = False
If InStr(1, pCadena, "CREATE") > 0 And vResultado Then vResultado = False
If InStr(1, pCadena, "ALTER") > 0 And vResultado Then vResultado = False
If InStr(1, pCadena, "sp_") > 0 And vResultado Then vResultado = False
If InStr(1, pCadena, "'") > 0 And vResultado Then vResultado = False


If Not vResultado Then
 'Registrar en Log de Seguridad todo el criterio
 MsgBox "!Error: El criterio de busqueda contiene información o datos que pueden afectar potencialmente la integridad de la información..!", vbExclamation
End If

fxValidaCriterio = vResultado

End Function


Private Sub sbCajaInicial()
Dim strSQL As String

'Paso 1: Si la Caja no está abierta (Llamar pantalla de login de Caja)
If ModuloCajas.mApertura = 0 Or ModuloCajas.mApertura = Empty Or ModuloCajas.mUsuario <> glogon.Usuario Then
   Call sbFormsCall("frmCajas_Acceso", vbModal, , , False, Me)
End If

'Paso 2: Si despues del Login de Caja permanece sin Apertura Salir
If ModuloCajas.mApertura = 0 Or ModuloCajas.mApertura = Empty Then
   MsgBox "No se ha indicado ninguna caja con Apertura disponible?", vbExclamation
   Unload Me
   Exit Sub
End If

pCharRelleno = fxCajasParametros("05")

Me.Caption = "Abonos a Créditos    ¦ Caja .: " & ModuloCajas.mCaja _
           & "   Apertura .: " & ModuloCajas.mApertura & "     Usuario.: " & ModuloCajas.mUsuario

txtTotalCajas.Text = 0
txtNotas.Text = ""
strSQL = "select rTrim(C.tipo_documento) as 'IdX', rtrim(D.Descripcion) as 'itmX'" _
       & " from SIF_DOCUMENTOS D inner join CAJAS_DOCUMENTOS C on D.TIPO_DOCUMENTO = C.TIPO_DOCUMENTO " _
       & " Where C.cod_caja =  '" & ModuloCajas.mCaja & "' and D.Tipo_Movimiento in('A','C')" _
       & " order by C.tipo_documento"
Call sbCbo_Llena_New(cboTipoDoc, strSQL, False, True)


ModuloCajas.mServicio = "Abonos a Operaciones de Crédito"

If IsNumeric(ModuloCajas.mRef_01) Then
    txtOperacion.Text = ModuloCajas.mRef_01
    vOperacion = txtOperacion.Text
    Call sbConsultaOperacion
End If


If ModuloCajas.mSesionId = 0 Then
   Call sbFormsCall("frmCajas_Sesion", vbModal, , , False, Me)
   If ModuloCajas.mSesionId = 0 Then
        MsgBox "No se ha iniciado ninguna sesión de Cliente para esta caja!", vbExclamation
        Unload Me
        Exit Sub
   End If
End If



End Sub

Private Sub Form_Activate()
 vModulo = 5
End Sub

Private Sub Form_Load()
Dim iDias As Integer

 vModulo = 5
 vOperacion = 0
 
 vFechaHoy = fxFechaServidor
 iDias = fxCrdParametro("32")
 
Set imgBanner.Picture = frmContenedor.imgBanner_01.Picture
 
 vScroll = False
 FlatScrollBar.Value = 0
 vScroll = True

cboDiferenciaApl.Clear
cboDiferenciaApl.AddItem "Abono Extraordinario"
cboDiferenciaApl.AddItem "Adelanto de Cuota"


With lsw.ColumnHeaders
    .Add , , "No.Cuota", 1000, vbCenter
    .Add , , "Proceso", 900, vbCenter
    .Add , , "Fec.Pago", 1300, vbCenter
    .Add , , "Cuota", 1300, vbRightJustify
    .Add , , "Estado", 1150
    .Add , , "Int.Cor.", 1300, vbRightJustify
    .Add , , "Int.Mor.", 1300, vbRightJustify
    .Add , , "Principal", 1300, vbRightJustify
    .Add , , "Cargos", 1200, vbRightJustify
    .Add , , "Pólizas", 1200, vbRightJustify
    .Add , , "Dias.Cor.", 1100, vbCenter
    .Add , , "Dias.Mor.", 1100, vbCenter
    .Add , , "Corte Cta", 1400, vbCenter
End With

Call sbLimpiaDatos

dtpFechaCancelacion.Value = vFechaHoy
dtpFechaCancelacion.MinDate = DateAdd("d", (iDias * -1), dtpFechaCancelacion.Value)
dtpFechaCancelacion.MaxDate = dtpFechaCancelacion.Value
 
 Call Formularios(Me)
 Call RefrescaTags(Me)
 
 btnCajas.Item(1).Enabled = cmdAplicar.Enabled

End Sub

Private Sub sbConsultaOperacion()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem

Me.MousePointer = vbHourglass

Call sbLimpiaDatos
 
strSQL = "select R.id_solicitud,R.saldo, R.saldo - isnull(V.amortiza,0) As Saldo_mes,R.proceso, isnull(R.cod_Divisa,'COL') as 'Divisa'" _
       & ",R.interesv,R.int,R.plazo,R.interesc,R.amortiza,R.fecult,R.Prideduc" _
       & ",R.opex,R.cuota,R.codigo,R.cedula,R.cuotas_planilla,R.cuotas_directas, datediff(m,R.fechaforp,dbo.MyGetdate()) as 'Meses'" _
       & ",S.nombre,C.descripcion,C.retencion,C.poliza,R.fechaforp,C.PORC_CARGO_CANCELACION,C.ANTICIPO_MESES,R.Base_Calculo" _
       & ",dbo.fxCrdPlanPagosDiasActivo(" & vOperacion & ") as 'DiasActivo', dbo.fxCrdOperacionTagReg(R.id_solicitud,'S15') as 'AutPagoAnt'" _
       & ",C.descripcion as 'LineaDesc',Ofi.descripcion as 'OficinaDesc',Pre.Descripcion as 'RecursoDesc',dbo.MyGetdate() as 'FechaServer'" _
       & ",dbo.fxCajas_Valida_Auxiliar('" & ModuloCajas.mCaja & "','CRD',R.Codigo) as 'Caja_Valida_Concepto'" _
       & ", dbo.fxCrd_Operacion_Control(R.id_solicitud) as 'Control'" _
       & " from reg_creditos R inner join Catalogo C on R.codigo = C.codigo " _
       & " inner join Socios S on R.cedula = S.cedula" _
       & " left join Sif_Oficinas Ofi on R.cod_Oficina_R = Ofi.cod_Oficina" _
       & " left join CATALOGO_GRUPOS Pre on R.cod_grupo = Pre.cod_grupo" _
       & " left join vista_morosidad V on R.id_solicitud = V.id_solicitud" _
       & " where R.estado = 'A' and R.saldo > 0" _
       & " and R.ID_SOLICITUD = " & vOperacion
       
Call OpenRecordSet(rs, strSQL)

If Not rs.EOF And Not rs.BOF Then
    txtCedula = rs!Cedula
    txtNombre = rs!Nombre
    txtCodigo = rs!Codigo
    
    mControl = rs!Control
    
    ModuloCajas.mClienteId = Trim(rs!Cedula)
    ModuloCajas.mCliente = Trim(rs!Nombre)
    ModuloCajas.mTiquete = Trim(rs!Codigo) & "." & rs!Id_Solicitud & "." & Format(Time, "HH:mm:ss")
    
    ModuloCajas.mDivisa = RTrim(rs!Divisa)
    ModuloCajas.mConceptoValida = IIf((rs!Caja_Valida_Concepto > 0), True, False)
    
    ModuloCajas.mTotalDetallado = 0
    txtTotalCajas.Text = 0
    
    
    
  vBaseCalculo = Trim(rs!Base_Calculo)
  vPrideduc = rs!PriDeduc
  vOperacion = rs!Id_Solicitud
  vPlazo = rs!Plazo
  vDiasActivo = rs!DiasActivo
  
  'Indica si Aplica Cargo por Cancelacion Anticipada y no se encuentra autorizado debe de cobrarse
  If rs!Meses <= rs!ANTICIPO_MESES And rs!AutPagoAnt = 0 Then
     vAnticipoPorc = rs!PORC_CARGO_CANCELACION / 100
  Else
     vAnticipoPorc = 0
  End If
  
  vInteres = IIf(IsNull(rs!interesv), rs!Int, rs!interesv)
  If IsNull(rs!saldo_mes) Then
    vSaldoMes = rs!Saldo
    strSQL = "update reg_creditos set saldo_mes = saldo where id_solicitud = " & rs!Id_Solicitud
    Call ConectionExecute(strSQL)
  Else
    If rs!saldo_mes = 0 Then
        vSaldoMes = rs!Saldo
        strSQL = "update reg_creditos set saldo_mes = saldo where id_solicitud = " & rs!Id_Solicitud
        Call ConectionExecute(strSQL)
    Else
       vSaldoMes = rs!saldo_mes
    End If
  
  End If
  
  vCuotasDeducidas = IIf(IsNull(rs!cuotas_planilla), 0, rs!cuotas_planilla)
  vCuotasDirectas = IIf(IsNull(rs!cuotas_directas), 0, rs!cuotas_directas)
     lblAmortiza.Caption = Format(rs!Amortiza, "Standard")
     lblAmortizaR.Caption = 0
     lblCuota = Format(rs!Cuota, "Standard")
     lblCuotaR.Caption = 0
     txtDatosAmortiza = 0
     txtDatosInteres.Text = 0
     lblFecUltMov.Caption = IIf(IsNull(rs!FecUlt), fxFechaProcesoAnterior(GLOBALES.glngFechaCR), rs!FecUlt)
    If CLng(lblFecUltMov.Caption) < GLOBALES.glngFechaCR Then
       lblFecUltMov.Caption = fxFechaProcesoAnterior(GLOBALES.glngFechaCR)
    End If
     lblFecUltMovR.Caption = 0
     lblInteres.Caption = Format(rs!interesc, "Standard")
     lblInteresR.Caption = 0
     txtOpex.Text = IIf((rs!opex = 1), "OPEX", "")
    
     lblSaldo.Tag = rs!FechaForp
     lblSaldo.Caption = Format(rs!Saldo, "Standard")
     lblSaldoR.Caption = 0
    
     txtCuotas = 0
     txtOperacion = rs!Id_Solicitud

     fraAbono.Enabled = True
     fraDatosAbono.Enabled = False

    
    txtProceso.Tag = rs!Proceso
    Select Case rs!Proceso
      Case "N"
        txtProceso.Text = "Normal"
      Case "T"
        txtProceso.Text = "Traspaso Deuda"
      Case "J"
        txtProceso.Text = "Cobro Judicial"
      Case "I"
        txtProceso.Text = "Incobrable"
    End Select
    
    
    txtDescripcion.Text = rs!DESCRIPCION
    
    txtDatosAnticipo.ToolTipText = "% de Comision : " & vAnticipoPorc
    txtDatosAnticipo.Tag = vAnticipoPorc
    
   
    
    If rs!retencion = "S" Or rs!Poliza = "S" Then
      vRetencion = True
    Else
      vRetencion = False
    End If
        
    'Barra de Estado
   
    StatusBarX.Panels.Item(1).Text = rs!OficinaDesc & ""
    StatusBarX.Panels.Item(2).Text = rs!LineaDesc & ""
    StatusBarX.Panels.Item(3).Text = rs!RecursoDesc & ""
        
        
       
    'Consulta Cuotas Activas
    strSQL = "select * from CRD_OPERACION_TRANSAC" _
           & " where estado = 'A' and id_solicitud = " & rs!Id_Solicitud _
           & "  and Fecha_Inicio < dbo.MyGetdate()" _
           & " order by ID_SEQ asc"
      
    rs.Close
    Call OpenRecordSet(rs, strSQL)
    lsw.ListItems.Clear
    Do While Not rs.EOF
      Set itmX = lsw.ListItems.Add(, , rs!Num_Cuota)
          itmX.SubItems(1) = Format(rs!Fecha_Proceso, "####-##")
          itmX.SubItems(2) = Format(rs!Fecha_Pago, "dd/mm/yyyy")
          itmX.SubItems(3) = Format(rs!Cuota, "Standard")
          itmX.SubItems(4) = IIf((rs!Mora_Dias > 0), "En Mora", "Al Día")
          itmX.SubItems(5) = Format(rs!IntCor, "Standard")
          itmX.SubItems(6) = Format(rs!IntMor, "Standard")
          itmX.SubItems(7) = Format(rs!Principal, "Standard")
          itmX.SubItems(8) = Format(rs!Cargos, "Standard")
          itmX.SubItems(9) = Format(rs!Poliza, "Standard")
          itmX.SubItems(10) = rs!Dias_calculo
          itmX.SubItems(11) = rs!Mora_Dias
          itmX.SubItems(12) = Format(rs!Fecha_Corte, "yyyy/mm/dd")
          
          itmX.Tag = rs!Id_seq
      rs.MoveNext
    Loop
    
    
    'Activacion de Tipos de Abonos
    
    Select Case lsw.ListItems.Count
      Case Is <= 0
            optAbono.Item(0).Enabled = False 'Ordinario
            optAbono.Item(1).Enabled = True 'Extraordinario
            optAbono.Item(2).Enabled = True 'Cancelacion
            optAbono.Item(3).Enabled = True 'Adelantos
            Call optAbono_Click(1)
      
      Case Is = 1
            optAbono.Item(0).Enabled = True 'Ordinario
            optAbono.Item(1).Enabled = True 'Extraordinario
            optAbono.Item(2).Enabled = True 'Cancelacion
            optAbono.Item(3).Enabled = True 'Adelantos
            Call optAbono_Click(0)
      
      Case Is > 1
            optAbono.Item(0).Enabled = True 'Ordinario
            optAbono.Item(1).Enabled = False 'Extraordinario
            optAbono.Item(2).Enabled = True 'Cancelacion
            optAbono.Item(3).Enabled = False 'Adelantos
            Call optAbono_Click(0)
    End Select
    

Else
 
 vOperacion = 0
 vPlazo = 0
 vInteres = 0
 vSaldoMes = 0
 MsgBox "No se Encontró operación para abonos,puede que se encuentre cancelada ", vbInformation

End If
rs.Close

Me.MousePointer = vbDefault

End Sub

Private Sub sbLimpiaDatos()
 
 txtTotalCajas.Text = 0
 
 txtDatosAnticipo.Text = 0
 lblAmortiza.Caption = 0
 lblAmortizaR.Caption = 0
 lblCuota = 0
 lblCuotaR.Caption = 0
 txtDatosAmortiza = 0
 txtDatosInteres.Text = 0
 lblFecUltMov.Caption = 0
 lblFecUltMovR.Caption = 0
 lblInteres.Caption = 0
 lblInteresR.Caption = 0
 txtDescripcion.Text = ""
 txtOpex.Text = ""
 lblSaldo.Caption = 0
 lblSaldoR.Caption = 0
 
 txtPolizas.Text = 0
 
 txtCedula = ""
 txtCodigo = ""
 txtCuotas = 0
 txtNombre = ""
 txtOperacion = ""


 cboDiferenciaApl.Text = "Abono Extraordinario"
 cboDiferenciaApl.Enabled = False
 txtTotalPagar.Text = 0
 txtTotalCancela.Text = 0
 
 txtProceso.Tag = ""
 txtProceso.Text = ""
 
 fraAbono.Enabled = False
 fraDatosAbono.Enabled = False
 
 fraCuotas.Visible = False
 lsw.ListItems.Clear
 
 chkRecalculaCuota.Value = vbUnchecked
 
 
dtpFechaCancelacion.Enabled = False
lblFechaCancelacion.Enabled = False
dtpFechaCancelacion.Value = vFechaHoy

 
 
End Sub

Private Sub sbBusqueda()

On Error GoTo vError

gBusquedas.Convertir = "N"
gBusquedas.Consulta = "Select R.id_solicitud as Operacion,R.Codigo,S.Cedula,S.Nombre,C.Descripcion" _
          & " from REG_CREDITOS R inner join SOCIOS S on R.cedula = S.cedula" _
          & " inner join Catalogo C on R.codigo = C.codigo"
gBusquedas.Columna = "R.CEDULA"
gBusquedas.Orden = "R.CEDULA"
gBusquedas.Filtro = " AND R.ESTADO = 'A'"

frmBusquedas.Show vbModal

txtOperacion = Trim(gBusquedas.Resultado)
vOperacion = txtOperacion

gBusquedas.Consulta = ""
gBusquedas.Columna = ""
gBusquedas.Orden = ""
gBusquedas.Resultado = ""
gBusquedas.Filtro = ""

Call sbConsultaOperacion

Me.MousePointer = vbDefault

Exit Sub

vError:
    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub sbCargaOperacionCodCed(vCedula As String, vCodigo As String)
Dim strSQL As String, rs As New ADODB.Recordset

strSQL = "select R.id_solicitud,R.saldo,R.saldo_mes,R.interesv,R.int,R.plazo,R.interesc,R.amortiza,R.fecult" _
       & ",R.opex,R.cuota,R.codigo,R.cedula,R.cuotas_planilla,R.cuotas_directas,C.retencion,C.poliza " _
       & "from reg_creditos R inner join Catalogo C on R.codigo = C.codigo " _
       & "where R.estado = 'A' and R.proceso <> 'N' and R.saldo > 0 " _
       & "and R.cedula = '" & txtCedula & "' and R.codigo = '" & txtCodigo & "'"
rs.CursorLocation = adUseServer
Call OpenRecordSet(rs, strSQL)
If Not rs.EOF And Not rs.BOF Then
  vOperacion = rs!Id_Solicitud
  vPlazo = rs!Plazo
  vInteres = IIf(IsNull(rs!interesv), rs!Int, rs!interesv)
  vSaldoMes = IIf(IsNull(rs!saldo_mes), rs!Saldo, rs!saldo_mes)
  vCuotasDeducidas = IIf(IsNull(rs!cuotas_planilla), 0, rs!cuotas_planilla)
  vCuotasDirectas = IIf(IsNull(rs!cuotas_directas), 0, rs!cuotas_directas)
     lblAmortiza.Caption = Format(rs!Amortiza, "Standard")
     lblAmortizaR.Caption = 0
     lblCuota = Format(rs!Cuota, "Standard")
     lblCuotaR.Caption = 0
     txtDatosAmortiza = 0
     txtDatosInteres.Text = 0
    
     lblFecUltMov.Caption = IIf(IsNull(rs!FecUlt), fxFechaProcesoAnterior(GLOBALES.glngFechaCR), rs!FecUlt)
    If CLng(lblFecUltMov.Caption) < GLOBALES.glngFechaCR Then
       lblFecUltMov.Caption = fxFechaProcesoAnterior(GLOBALES.glngFechaCR)
    End If
     lblFecUltMovR.Caption = 0
     lblInteres.Caption = Format(rs!interesc, "Standard")
     lblInteresR.Caption = 0
     txtOpex.Text = IIf((rs!opex = 1), "OPEX", "")
     lblSaldo.Caption = Format(vSaldoMes, "Standard")
     lblSaldoR.Caption = 0
     txtCuotas = 0
     txtOperacion = rs!Id_Solicitud

     fraAbono.Enabled = True
     fraDatosAbono.Enabled = False
    
    optAbono(0).Enabled = True
    optAbono(1).Enabled = True
    
    
    If rs!retencion = "S" Or rs!Poliza = "S" Then
      vRetencion = True
    Else
      vRetencion = False
    End If
        
    If Not vRetencion Then
        Select Case True
         Case optAbono(0).Checked
           Call optAbono_Click(0)
         Case optAbono(1).Checked
           Call optAbono_Click(1)
        End Select
    Else
           Call optAbono_Click(0)
           optAbono(1).Enabled = False
    End If
    
    
Else
 
 vOperacion = 0
 vPlazo = 0
 vInteres = 0
 vSaldoMes = 0
 MsgBox "No se Encontrarón operaciones para abonos con esta cédula y código", vbInformation
End If
rs.Close

End Sub

Private Sub imgDocumento_Click()
  frmCR_AbonosComprobante.Show vbModal
End Sub




Private Sub lsw_ItemCheck(ByVal Item As XtremeSuiteControls.ListViewItem)
Dim curInteres As Currency, curPrincipal As Currency, curCargos As Currency, curPolizas As Currency
Dim i As Integer

curInteres = 0
curPrincipal = 0
curCargos = 0
curPolizas = 0

With lsw.ListItems
  For i = 1 To .Count
    If .Item(i).Checked Then
       curInteres = curInteres + CCur(.Item(i).SubItems(5)) + CCur(.Item(i).SubItems(6))
       curPrincipal = curPrincipal + CCur(.Item(i).SubItems(7))
       curCargos = curCargos + CCur(.Item(i).SubItems(8))
       curPolizas = curPolizas + CCur(.Item(i).SubItems(9))
    End If
  Next i
End With

txtDatosAmortiza.Text = Format(curPrincipal, "Standard")
txtDatosInteres.Text = Format(curInteres, "Standard")
txtDatosCargos.Text = Format(curCargos, "Standard")
txtPolizas.Text = Format(curPolizas, "Standard")
txtDatosAnticipo.Text = 0

txtTotalPagar.Text = Format(curPrincipal + curInteres + curCargos + curPolizas, "Standard")
txtTotalCancela.Text = txtTotalPagar.Text

End Sub

Private Sub tblDesgloce_ButtonClick(ByVal Button As MSComctlLib.Button)
If Not IsNumeric(txtTotalPagar.Text) Then txtTotalPagar.Text = 0

ModuloCajas.mTotalAplicar = CCur(txtTotalPagar.Text)

If ModuloCajas.mTotalAplicar = 0 Then
    MsgBox "No se ha especificado ningún monto a detallar?", vbExclamation
    Exit Sub
End If

ModuloCajas.mServicio = "Abonos a Operación de Crédito"

Call sbFormsCall("frmCajas_DetallePago", vbModal, 0, 0, False, Me)

txtTotalCajas.Text = Format(ModuloCajas.mTotalDetallado, "Standard")


If txtTotalCajas.Text <> txtTotalPagar.Text Then
   txtTotalCajas.BackColor = vbRed
Else
   txtTotalCajas.BackColor = vbWhite
End If

End Sub

Private Sub TimerX_Timer()
TimerX.Enabled = False
TimerX.Interval = 0

Call sbCajaInicial

If ModuloCajas.mApertura = 0 Or ModuloCajas.mApertura = Empty Then
   Unload Me
   Exit Sub
End If

Call optAbono_Click(0)

End Sub



Private Sub txtTotalPagar_Change()
Dim strSQL As String, rs As New ADODB.Recordset
Dim ProcesosTmp As Currency, lngFecha As Currency, iPlazoRst As Integer, curCuota As Currency

On Error Resume Next

If chkRecalculaCuota.Value = vbChecked Then
  
    ' strSQL = "select plazo + DATEDIFF(mm,  dbo.MyGetdate(), CONVERT(DATETIME, substring(convert(varchar(6), prideduc), 1,4) + '/' + substring(convert(varchar(6), prideduc), 5,2) + '/28' )) as PlazoFaltante" _
    '       & " from reg_creditos where id_solicitud = " & txtOperacion
    ' Call OpenRecordSet(rs, strSQL)
    '    lblCuotaR.Caption = fxCalcula_Cuota(CDbl(lblSaldoR.Caption), rs!PlazoFaltante, vInteres)
    ' rs.Close
       lngFecha = lblFecUltMov.Caption
       If lngFecha < vPrideduc Then lngFecha = vPrideduc
      
       ProcesosTmp = vPrideduc
       iPlazoRst = 1
        Do While ProcesosTmp < lngFecha
          ProcesosTmp = fxFechaProcesoSiguiente(ProcesosTmp)
          iPlazoRst = iPlazoRst + 1
        Loop
       iPlazoRst = vPlazo - iPlazoRst
       curCuota = fxCalcula_Cuota(CDbl(lblSaldoR.Caption), iPlazoRst, vInteres)
       lblCuotaR.Caption = Format(curCuota, "Standard")
Else
  lblCuotaR.Caption = lblCuota.Caption
End If

End Sub

Private Sub optAbono_Click(Index As Integer)
Dim strSQL As String, rs As New ADODB.Recordset
Dim curInteres As Currency, curIntMor As Currency, curPrincipal As Currency, curCargos As Currency
Dim vFecha As Date, vProceso As Long, i As Integer

On Error GoTo vErrorCarga

Me.MousePointer = vbHourglass

fraCuotas.Visible = False
fraDatosAbono.Enabled = True

chkRecalculaCuota.Enabled = False
chkRecalculaCuota.Value = vbUnchecked

'&H00C0FFC0&
txtTotalPagar.BackColor = &HC0FFC0

txtTotalPagar.Locked = True

txtCuotas.Enabled = False
FlatScrollBar.Enabled = txtCuotas.Enabled

For i = 1 To lsw.ListItems.Count
  lsw.ListItems.Item(i).Checked = False
Next

For i = 0 To 3
  If i = Index Then
      optAbono.Item(i).Checked = True
  Else
      optAbono.Item(i).Checked = False
  End If
Next i


dtpFechaCancelacion.Enabled = False
lblFechaCancelacion.Enabled = False

Select Case Index

 Case 0 'Ordinario
   txtDatosCargos.Text = 0
   txtDatosInteres.Text = 0
   txtDatosAnticipo.Text = 0
   txtPolizas.Text = 0
   txtDatosAmortiza = 0
      
   txtTotalCancela.Text = 0
   txtTotalPagar.Text = 0
   
   txtCuotas.Text = 0 'Inicializa
   
   fraCuotas.Visible = True
   
   txtTotalPagar.BackColor = vbWhite
   txtTotalPagar.Locked = False
   txtTotalPagar.SetFocus
   
 Case 1 'Extraordinario
   txtCuotas = 0
   
   lblFechaCancelacion.Caption = "Fecha de Abono (Real) por parte del cliente para Ab.Extraordinario:"
   dtpFechaCancelacion.Enabled = True
   lblFechaCancelacion.Enabled = True
   
   txtDatosInteres.Text = 0
   txtDatosAnticipo.Text = 0
   txtDatosCargos.Text = 0
   txtPolizas.Text = 0
   
  
   txtDatosAmortiza.Text = 0
   txtTotalCancela.Text = 0
   
   txtTotalPagar.BackColor = vbWhite
   txtTotalPagar.Locked = False
   txtTotalPagar.SetFocus
   
   chkRecalculaCuota.Enabled = True
   strSQL = "select dbo.fxCrdPlanPagosDiasActivoFecha( " & txtOperacion.Text & ", '" & Format(dtpFechaCancelacion.Value, "yyyy/mm/dd") & "') as 'Dias'"
   Call OpenRecordSet(rs, strSQL)
     vDiasActivo = rs!Dias
   rs.Close
   
Case 2 'Cancelación
   
   txtDatosAmortiza.Text = 0
  
   lblFechaCancelacion.Caption = "Fecha de Abono (Real) por parte del cliente para cancelación...:"
   dtpFechaCancelacion.Enabled = True
   lblFechaCancelacion.Enabled = True
   
   strSQL = "exec spCrdPlanPagosInfoCancelacion " & txtOperacion.Text & ", '" & Format(dtpFechaCancelacion.Value, "yyyy/mm/dd") & "'"
      
      
   
   Call OpenRecordSet(rs, strSQL)
    txtDatosAmortiza.Text = Format(rs!Principal, "Standard")
    txtDatosInteres.Text = Format(rs!IntCor + rs!IntMor, "Standard")
    txtDatosCargos.Text = Format(rs!Cargos, "Standard")
    txtDatosAnticipo.Text = Format(rs!CargoAnticipo, "Standard")
    txtPolizas.Text = Format(rs!Poliza, "Standard")
    txtTotalPagar.Text = Format(rs!Principal + rs!IntCor + rs!IntMor + rs!Cargos + rs!CargoAnticipo + rs!Poliza, "Standard")
    txtTotalCancela.Text = txtTotalPagar.Text
   rs.Close
   
   If vRetencion Then
      txtDatosAnticipo.Text = "0.00"
   End If
   


 Case 3 'Adelantos
   txtDatosAnticipo.Text = 0
   txtDatosCargos.Text = 0
   txtDatosInteres.Text = 0
   txtPolizas.Text = 0
   txtDatosAmortiza.Text = 0
   
   txtCuotas.Enabled = True
   FlatScrollBar.Enabled = txtCuotas.Enabled
   
   txtCuotas.Text = 0 'Inicializa
   txtCuotas.Text = 1 'Inicializa
   txtCuotas.SetFocus
   
   txtTotalPagar.BackColor = vbWhite
   txtTotalPagar.Locked = False
   txtTotalPagar.SetFocus

End Select

vErrorCarga:

Call RefrescaTags(Me)


Me.MousePointer = vbDefault


End Sub

Private Sub txtCedula_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then Call sbBusqueda
End Sub

Private Sub txtCedula_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Or KeyAscii = vbKeyTab Then
  txtNombre = fxNombre(txtCedula)
  If txtCodigo <> "" Then Call sbCargaOperacionCodCed(txtCedula, txtCodigo)
  txtCodigo.SetFocus
End If
End Sub

Private Sub txtCodigo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then Call sbBusqueda
End Sub

Private Sub txtCodigo_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Or KeyAscii = vbKeyTab Then
  txtCodigo = UCase(txtCodigo)
  txtDescripcion.Text = fxDescribeCodigo(txtCodigo)
  If txtCedula <> "" Then Call sbCargaOperacionCodCed(txtCedula, txtCodigo)
  txtOperacion.SetFocus
End If

End Sub

Private Sub sbCuotaChangeAnterior()
Dim curSaldo As Currency, curAmortiza As Currency, curInteres As Currency
Dim curTmpAmortiza As Currency, curTmpInteres As Currency, i As Integer
Dim lngFecha As Currency, lngCuotas As Long, lngCuotaMaxima As Long


Dim iDias As Integer, vFecha As Date, curCuota As Currency, iPlazoRst As Integer, ProcesosTmp As Currency

On Error Resume Next

If txtCuotas = "" Then
 lngCuotas = 0
Else
 lngCuotas = txtCuotas
End If

lngFecha = CCur(lblFecUltMov.Caption)

If Not vRetencion Then
    curSaldo = vSaldoMes
Else
  'En las retenciones hay que proyectar el saldo del mes
  curSaldo = ((CCur(lblCuota.Caption) * vPlazo) - CCur(lblAmortiza.Caption))
End If

curAmortiza = 0
curInteres = 0
curCuota = lblCuota.Caption


If lngFecha < vPrideduc Then lngFecha = vPrideduc

If vBaseCalculo = "01" Then
    For i = 1 To lngCuotas
    '360 / 360
        If curSaldo > 0 Then
          lngCuotaMaxima = i
          curTmpInteres = (curSaldo * vInteres) / 1200
          curTmpAmortiza = CCur(lblCuota.Caption) - curTmpInteres
          
          curAmortiza = curAmortiza + curTmpAmortiza
          curInteres = curInteres + curTmpInteres
          
          curSaldo = curSaldo - curTmpAmortiza
          lngFecha = fxFechaProcesoSiguiente(lngFecha)
        
        End If
        
        If curSaldo < 0 Then
           curAmortiza = curAmortiza + curSaldo
           curSaldo = 0
        End If
     Next i
 
 Else
   '365 / 360
   
       'Calcula el Plazo Restante
       ProcesosTmp = vPrideduc
       iPlazoRst = 0
        Do While ProcesosTmp < lngFecha
          ProcesosTmp = fxFechaProcesoSiguiente(ProcesosTmp)
          iPlazoRst = iPlazoRst + 1
        Loop
       iPlazoRst = vPlazo - iPlazoRst
       
       'Saca el formato fecha del ultimo movimiento para calculo de dias
       vFecha = Mid(CStr(lngFecha), 1, 4) & "/" & Right(CStr(lngFecha), 2) & "/01"
       
       For i = 1 To lngCuotas
          lngCuotaMaxima = i

          If iPlazoRst = 1 Or iPlazoRst = vPlazo Then
            iDias = 30
          Else
            iDias = fxMesDias(Month(vFecha), Year(vFecha))
          End If
        
          curTmpInteres = curSaldo * (vInteres / 100) * iDias / 360
          curTmpAmortiza = curCuota - curTmpInteres
          
          curAmortiza = curAmortiza + curTmpAmortiza
          curInteres = curInteres + curTmpInteres
          
          curSaldo = curSaldo - curTmpAmortiza
          lngFecha = fxFechaProcesoSiguiente(lngFecha)
          vFecha = DateAdd("m", 1, vFecha)
          
          iPlazoRst = iPlazoRst - 1
          curCuota = fxCalcula_Cuota(CDbl(curSaldo), iPlazoRst, vInteres)
          
       Next i
    
   
 End If 'Base

txtDatosInteres.Text = Format(curInteres, "Standard")
txtDatosAmortiza = Format(curAmortiza, "Standard")
lblFecUltMovR.Caption = lngFecha
lblCuotaR.Caption = Format(curCuota, "Standard")

If Not vRetencion Then 'El proceso nuevo de retenciones no toca los saldos
    lblSaldoR.Caption = Format(CCur(lblSaldo.Caption) - curAmortiza, "Standard")
End If

lblAmortizaR.Caption = Format(CCur(lblAmortiza.Caption) + curAmortiza, "Standard")
lblInteresR.Caption = Format(CCur(lblInteres.Caption) + curInteres, "Standard")

If lngCuotas > lngCuotaMaxima Then txtCuotas = lngCuotaMaxima


End Sub

Private Sub txtCuotas_Change()
Dim strSQL As String, rs As New ADODB.Recordset
Dim lngCuotas As Long

If vOperacion = 0 Then Exit Sub

On Error GoTo vError

If Not IsNumeric(txtCuotas.Text) Then
 lngCuotas = 1
Else
 lngCuotas = txtCuotas.Text
End If

If lngCuotas <= 0 Then lngCuotas = 1

strSQL = "select isnull(max(id_Seq),0) as 'SeqX', isnull(sum(IntCor + IntMor),0) as 'IntCor', isnull(sum(Principal),0) as 'Principal'" _
       & ",isnull(min(Saldo_Actual),0) as 'Saldo', isnull(max(Fecha_Proceso),0) as 'Fecha_Proceso', isnull(sum(Poliza),0) as 'Poliza'" _
       & " from CRD_OPERACION_PLAN_PAGOS where id_solicitud = " & vOperacion _
       & " and Id_Seq in(select Top " & lngCuotas & " Id_Seq from CRD_OPERACION_PLAN_PAGOS" _
       & " where estado in('A','P') and id_solicitud = " & vOperacion & " and num_cuota > 0  order by num_cuota)"
Call OpenRecordSet(rs, strSQL)
    txtDatosInteres.Text = Format(rs!IntCor, "Standard")
    txtPolizas.Text = Format(rs!Poliza, "Standard")
    lblFecUltMovR.Caption = rs!Fecha_Proceso
    
    If Not vRetencion Then 'El proceso nuevo de retenciones no toca los saldos
        lblSaldoR.Caption = Format(CCur(lblSaldo.Caption) - rs!Principal, "Standard")
    End If
    
    
    lblAmortizaR.Caption = Format(CCur(lblAmortiza.Caption) + rs!Principal, "Standard")
    lblInteresR.Caption = Format(CCur(lblInteres.Caption) + rs!IntCor, "Standard")

    'Se pone de ultimo porque activa otro sub
    txtDatosAmortiza.Text = Format(rs!Principal, "Standard")

strSQL = "select cuota from CRD_OPERACION_PLAN_PAGOS where id_seq = " & rs!SeqX & " and id_solicitud = " & vOperacion
rs.Close

Call OpenRecordSet(rs, strSQL)
If Not rs.EOF And Not rs.BOF Then
    lblCuotaR.Caption = Format(rs!Cuota, "Standard")
End If
rs.Close


Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical


End Sub

Private Sub txtDatosAmortiza_Change()
On Error Resume Next

If Not vRetencion Then
    lblSaldoR.Caption = Format(CCur(lblSaldo.Caption) - CCur(txtDatosAmortiza), "Standard")
Else
    lblSaldoR.Caption = lblCuota.Caption
End If
lblAmortizaR.Caption = Format(CCur(lblAmortiza.Caption) + CCur(txtDatosAmortiza), "Standard")
lblInteresR.Caption = Format(CCur(lblInteres.Caption) + CCur(txtDatosInteres), "Standard")


txtTotalPagar.Text = Format(CCur(txtDatosAmortiza) + CCur(txtDatosInteres.Text) + CCur(txtPolizas.Text) _
                + CCur(txtDatosAnticipo.Text) + CCur(txtDatosCargos.Text), "Standard")
txtTotalCancela.Text = txtTotalPagar.Text
txtDiferencia.Text = "0.00"


End Sub

Private Sub txtDatosAmortiza_GotFocus()
On Error Resume Next
txtDatosAmortiza = CCur(txtDatosAmortiza)
End Sub

Private Sub txtDatosAmortiza_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Or KeyAscii = vbKeyTab Then
 txtDatosAmortiza = Format(txtDatosAmortiza, "Standard")
 cboTipoDoc.SetFocus
End If
End Sub

Private Sub txtDatosAmortiza_LostFocus()
On Error Resume Next
txtDatosAmortiza = Format(txtDatosAmortiza, "Standard")
End Sub

Public Sub sbConsultaExterna(xOpTemp As Long)
 txtOperacion = xOpTemp
 Call txtOperacion_KeyPress(vbKeyReturn)
End Sub

Private Sub txtOperacion_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then Call sbBusqueda
End Sub

Private Sub txtOperacion_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = vbKeyReturn Or KeyAscii = vbKeyTab Then
 vOperacion = txtOperacion
 Call sbConsultaOperacion
End If
End Sub


Private Sub sbDocumentoAbono(pTipoAbono As String, pTipoDoc As String, pNumDoc As String _
                                , pConcepto As String, pCuenta As String)
Dim rs As New ADODB.Recordset, strSQL As String, strLinea(11) As String
Dim strCliente As String, vCuenta As String
Dim rsTmp As New ADODB.Recordset, vCuentaPoliza As String, pTipoCambio As Currency
Dim curIntC As Currency, curIntM As Currency, curCargo As Currency, curAmortiza As Currency, curPoliza As Currency

vCuenta = pCuenta

pTipoCambio = fxCajasTipoCambio(ModuloCajas.mDivisa)

'Cuentas
strSQL = "exec spCrdOperacionCtas " & txtOperacion.Text
Call OpenRecordSet(rs, strSQL)


strSQL = "exec spCrdDocumentoAfectacion '" & fxTipoASENumero(pTipoDoc) & "','" & pNumDoc & "','R'"
Call OpenRecordSet(rsTmp, strSQL, 0)
If rsTmp.EOF And rsTmp.BOF Then
  curIntC = 0
  curIntM = 0
  curAmortiza = 0
  curCargo = 0
  curPoliza = 0
Else
  curIntC = rsTmp!IntCor
  curIntM = rsTmp!IntMor
  curAmortiza = rsTmp!Principal
  curCargo = rsTmp!Cargos
  curPoliza = rsTmp!Polizas
End If
rsTmp.Close



strLinea(1) = "Saldo Anterior    ..: " & SIFGlobal.fxStringRelleno(lblSaldo.Caption, "I", pCharRelleno, 15) '
strLinea(2) = "Saldo Actual      ..: " & SIFGlobal.fxStringRelleno(Format(CCur(lblSaldo.Caption) - curAmortiza, "Standard"), "I", pCharRelleno, 15) '
strLinea(3) = "Interes Corriente ..: " & SIFGlobal.fxStringRelleno(Format(curIntC, "Standard"), "I", pCharRelleno, 15) '
strLinea(4) = "Interes Atrasado  ..: " & SIFGlobal.fxStringRelleno(Format(curIntM, "Standard"), "I", pCharRelleno, 15) '
strLinea(5) = "Amortización      ..: " & SIFGlobal.fxStringRelleno(Format(curAmortiza, "Standard"), "I", pCharRelleno, 15) '
strLinea(6) = "Cargos Totales    ..: " & SIFGlobal.fxStringRelleno(Format(curCargo, "Standard"), "I", pCharRelleno, 15) '
strLinea(7) = "Pólizas           ..: " & SIFGlobal.fxStringRelleno(Format(curPoliza, "Standard"), "I", pCharRelleno, 15) '
strLinea(8) = "Operacion/Línea   ..: " & "Op.:" & txtOperacion.Text & " L.:" & txtCodigo & " Ret.:" & IIf(vRetencion, "SI", "NO")

If cboDiferenciaApl.Enabled Then
    strLinea(9) = "Aplica Diferencia ..: " & cboDiferenciaApl.Text
Else
    strLinea(9) = "Descripción       ..: " & txtDescripcion.Text

End If

'Lineas de Comprobante
strSQL = "exec spCrdOperacionFechaProxPago " & txtOperacion.Text
Call OpenRecordSet(rsTmp, strSQL, 0)
  If Not IsNull(rsTmp!Fecha_Pago) Then
       strLinea(10) = "Prox.Pago..:" & Format(rsTmp!Fecha_Pago, "dd/mm/yyyy") & " Cta.(" & rsTmp!Num_Cuota & ") " & Format(rsTmp!Cuota, "Standard")
  Else
       strLinea(10) = "Prox.Pago..: >> <<"
  End If
  strLinea(10) = "Notas: " & rsTmp!Notas & ""
rsTmp.Close
      

If dtpFechaCancelacion.Enabled Then
   strLinea(11) = "Fecha Real Abono  ..: " & Format(dtpFechaCancelacion.Value, "dd/mm/yyyy")
End If


'Registro del Comprobante
strSQL = "insert SIF_TRANSACCIONES(COD_TRANSACCION,TIPO_DOCUMENTO,REGISTRO_FECHA,REGISTRO_USUARIO,Cliente_IDENTIFICACION,CLIENTE_NOMBRE" _
         & ",cod_concepto,monto,estado,Referencia_01,Referencia_02,Referencia_03,cod_oficina" _
         & ",linea1,linea2,linea3,linea4,linea5,linea6,linea7,linea8,linea9,linea10,linea11,detalle,documento,cod_caja,cod_Apertura)" _
         & " values('" & pNumDoc & "','" & pTipoDoc & "',dbo.MyGetdate(),'" & glogon.Usuario & "','" & Trim(txtCedula.Text) _
         & "','" & Trim(txtNombre.Text) & "','" & pConcepto & "'," & curIntC + curIntM + curAmortiza + curCargo + curPoliza & ",'P','" & txtOperacion.Text _
         & "','" & txtCodigo.Text & "','" & vAseDocDeposito & "','" & GLOBALES.gOficinaTitular & "','" & strLinea(1) & "','" _
         & strLinea(2) & "','" & strLinea(3) & "','" & strLinea(4) & "','" _
         & strLinea(5) & "','" & strLinea(6) & "','" & strLinea(7) & "','" _
         & strLinea(8) & "','" & strLinea(9) & "','" & strLinea(10) & "','" & strLinea(11) & "','" _
         & txtNotas.Text & "','" & vAseDocDeposito & "','" & ModuloCajas.mCaja & "'," & ModuloCajas.mApertura & ")"
' Call ConectionExecute(strSQL)
 
 'ASIENTO
 If curIntC <> 0 Then
   strSQL = strSQL & Space(10) & "exec spSIFDocsAsiento '" & pTipoDoc & "','" & pNumDoc & "'," & curIntC * fxSys_Tipo_Cambio_Apl(pTipoCambio) & ",'C','" & rs!cod_Divisa _
          & "'," & pTipoCambio & "," & GLOBALES.gEnlace & ",'" & rs!Cod_Unidad & "','" & rs!Cod_Centro_Costo & "','" & rs!ctaintc _
          & "','" & rs!Id_Solicitud & "','" & rs!Codigo & "','" & vAseDocDeposito & "'"
'   Call ConectionExecute(strSQL)
 End If
 
 If curIntM <> 0 Then
   strSQL = strSQL & Space(10) & "exec spSIFDocsAsiento '" & pTipoDoc & "','" & pNumDoc & "'," & curIntM * fxSys_Tipo_Cambio_Apl(pTipoCambio) & ",'C','" & rs!cod_Divisa _
          & "'," & pTipoCambio & "," & GLOBALES.gEnlace & ",'" & rs!Cod_Unidad & "','" & rs!Cod_Centro_Costo & "','" & rs!ctaintm _
          & "','" & rs!Id_Solicitud & "','" & rs!Codigo & "','" & vAseDocDeposito & "'"
'   Call ConectionExecute(strSQL)
 End If
 
 If curCargo <> 0 Then
 'Detallar Cargos
   glogon.strSQL = "exec spCrdDocumentoAfectacionCargos '" & pTipoDoc & "','" & pNumDoc & "'"
   Call OpenRecordSet(rsTmp, glogon.strSQL, 0)
   Do While Not rsTmp.EOF
         strSQL = strSQL & Space(10) & "exec spSIFDocsAsiento '" & pTipoDoc & "','" & pNumDoc & "'," & IIf(IsNull(rsTmp!Mov_Monto), curCargo, rsTmp!Mov_Monto * fxSys_Tipo_Cambio_Apl(pTipoCambio)) & ",'C','" & rs!cod_Divisa _
                & "'," & pTipoCambio & "," & GLOBALES.gEnlace & ",'" & rsTmp!Cod_Unidad & "','" & rsTmp!Cod_Centro_Costo & "','" & rsTmp!cod_cuenta _
                & "','" & rsTmp!Id_Solicitud & "','" & rsTmp!Codigo & "','" & vAseDocDeposito & "'"
'         Call ConectionExecute(strSQL)
         rsTmp.MoveNext
   Loop
   rsTmp.Close
 End If
 
 If curPoliza > 0 Then
  
 'Detallar Poliza
   glogon.strSQL = "exec spCrdDocumentoAfectacionPolizas '" & pTipoDoc & "','" & pNumDoc & "'"
   Call OpenRecordSet(rsTmp, glogon.strSQL, 0)
   Do While Not rsTmp.EOF
         strSQL = strSQL & Space(10) & "exec spSIFDocsAsiento '" & pTipoDoc & "','" & pNumDoc & "'," & rsTmp!Mov_Monto * fxSys_Tipo_Cambio_Apl(pTipoCambio) & ",'C','" & rs!cod_Divisa _
                & "'," & pTipoCambio & "," & GLOBALES.gEnlace & ",'" & rs!Cod_Unidad & "','" & rs!Cod_Centro_Costo & "','" & rsTmp!cod_cuenta _
                & "','" & rs!Id_Solicitud & "','" & rs!Codigo & "','" & vAseDocDeposito & "'"
'         Call ConectionExecute(strSQL)
         rsTmp.MoveNext
   Loop
   rsTmp.Close
   
 End If
 
 If curAmortiza <> 0 Then
   strSQL = strSQL & Space(10) & "exec spSIFDocsAsiento '" & pTipoDoc & "','" & pNumDoc & "'," & curAmortiza * fxSys_Tipo_Cambio_Apl(pTipoCambio) & ",'C','" & rs!cod_Divisa _
          & "'," & pTipoCambio & "," & GLOBALES.gEnlace & ",'" & rs!Cod_Unidad & "','" & rs!Cod_Centro_Costo & "','" & rs!ctaamortiza _
          & "','" & rs!Id_Solicitud & "','" & rs!Codigo & "','" & vAseDocDeposito & "'"
'   Call ConectionExecute(strSQL)
 End If
 
 If curIntC + curIntM + curPoliza + curCargo + curAmortiza <> 0 Then
     'Procesa Formas de Pago (Registro Final / Asiento de Pago)
      strSQL = strSQL & Space(10) & "exec spCajas_DesglocePagosDocFinal '" & ModuloCajas.mCaja & "'," & ModuloCajas.mApertura & ",'" & ModuloCajas.mTiquete _
              & "','" & ModuloCajas.mUsuario & "','" & pTipoDoc & "','" & pNumDoc & "','" & ModuloCajas.mUnidad _
              & "','" & rs!Id_Solicitud & "','" & rs!Codigo & "'"
'      Call ConectionExecute(strSQL)
 End If
 
rs.Close


'Aplica en una sola Llamada
Call ConectionExecute(strSQL)


End Sub


Private Sub txtTotalPagar_GotFocus()

On Error GoTo vError
 txtTotalPagar.Text = CCur(txtTotalPagar.Text)
vError:

End Sub


Private Sub txtTotalPagar_LostFocus()
Dim vFecha As Date, vProceso As Long
Dim curInteres As Currency, curAmortiza As Currency, curAnticipo As Currency
Dim i As Integer, vChecks As Boolean, iPlazo As Integer
 
On Error GoTo vError
 
'ExtraOrdinario
If optAbono.Item(1).Checked = True Then
   'Cobra intereses desde el ultimo corte
    txtTotalCancela.Text = Format(txtTotalPagar.Text, "Standard")
    curInteres = (CCur(txtTotalPagar.Text) * vInteres / 36000) * vDiasActivo
    curAnticipo = CCur(txtTotalPagar.Text) * vAnticipoPorc
   'Se re-calculan intereses para ajustar y relacionar segun porcion amortizada
   'Previamente sobre el monto a cancelar
   
   If curInteres + curAnticipo > 0 Then
      'Hacer 10 aproximaciones
      For i = 1 To 10
            curAmortiza = CCur(txtTotalPagar.Text) - (curInteres + curAnticipo)
            curInteres = (curAmortiza * vInteres / 36000) * vDiasActivo
      Next i
   End If
   
   txtDatosInteres.Text = Format(curInteres, "Standard")
   txtDatosAnticipo.Text = Format(curAnticipo, "Standard")
   txtDatosAmortiza.Text = Format(CCur(txtTotalPagar.Text) - (curInteres + curAnticipo), "Standard")
End If


txtTotalPagar.Text = Format(CCur(txtTotalPagar.Text), "Standard")

txtDiferencia.Text = Format(CCur(txtTotalCancela.Text) - CCur(txtTotalPagar.Text), "Standard")

cboDiferenciaApl.Enabled = False

If CCur(txtDiferencia.Text) < 0 Then
    Select Case True
      Case optAbono.Item(0).Checked 'Abono Ordinario
           cboDiferenciaApl.Enabled = False
           'Verifica el Plazo sea menor que la ultima cuota marcada y que se hayan marcado todas con corte igual o menor a la fecha actual
           vChecks = True
           For i = 1 To lsw.ListItems.Count
             If Not lsw.ListItems.Item(i).Checked And DateDiff("d", CDate(lsw.ListItems.Item(i).SubItems(11)), vFechaHoy) >= 0 Then
               vChecks = False
             End If
           Next i
           
           'Verifica el Ultimo Plazo
           If vChecks Then
              If CCur(lblSaldoR.Caption) > CCur(txtDiferencia.Text) Then
                  cboDiferenciaApl.Enabled = True
                  Call cboDiferenciaApl_Click
              End If
           End If
           
      Case optAbono.Item(1).Checked 'Abono ExtraOrdinario
           cboDiferenciaApl.Enabled = False
      
      Case optAbono.Item(2).Checked 'Cancelacion
           cboDiferenciaApl.Enabled = False
      
      Case optAbono.Item(3).Checked 'Adelanto
           'Verifica si el Saldo Resultante del Credito es mayor igual a la diferencia.
              If CCur(lblSaldoR.Caption) > CCur(txtDiferencia.Text) Then
                  cboDiferenciaApl.Enabled = True
              End If
    
    End Select
End If

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub




