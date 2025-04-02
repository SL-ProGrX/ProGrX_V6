VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#19.3#0"; "Codejock.Controls.v19.3.0.ocx"
Begin VB.Form frmCO_ReversionCobroJudicial 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Reversión de Cobro Judicial"
   ClientHeight    =   6432
   ClientLeft      =   48
   ClientTop       =   312
   ClientWidth     =   11304
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6432
   ScaleWidth      =   11304
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.PushButton btnPrincipal 
      Height          =   612
      Index           =   0
      Left            =   6480
      TabIndex        =   0
      Top             =   5640
      Width           =   2412
      _Version        =   1245187
      _ExtentX        =   4254
      _ExtentY        =   1080
      _StockProps     =   79
      Caption         =   "Reversar Traslado"
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
      Picture         =   "frmCO_ReversionCobroJudicial.frx":0000
   End
   Begin XtremeSuiteControls.PushButton btnPrincipal 
      Height          =   612
      Index           =   1
      Left            =   8880
      TabIndex        =   1
      ToolTipText     =   "Cerrar Ventana"
      Top             =   5640
      Width           =   612
      _Version        =   1245187
      _ExtentX        =   1080
      _ExtentY        =   1080
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
      Appearance      =   16
      Picture         =   "frmCO_ReversionCobroJudicial.frx":098D
   End
   Begin XtremeSuiteControls.GroupBox gbDeuda 
      Height          =   2172
      Left            =   240
      TabIndex        =   3
      Top             =   1440
      Width           =   10812
      _Version        =   1245187
      _ExtentX        =   19071
      _ExtentY        =   3831
      _StockProps     =   79
      Caption         =   "Deuda"
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   16
      BorderStyle     =   1
      Begin XtremeSuiteControls.DateTimePicker dtpCalculoIntCorte 
         Height          =   312
         Left            =   2400
         TabIndex        =   4
         Top             =   4560
         Width           =   1212
         _Version        =   1245187
         _ExtentX        =   2138
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
      Begin XtremeSuiteControls.FlatEdit txtInteresesCorte 
         Height          =   312
         Left            =   2880
         TabIndex        =   5
         Top             =   720
         Width           =   1812
         _Version        =   1245187
         _ExtentX        =   3196
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
         Alignment       =   1
         Locked          =   -1  'True
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtSaldo 
         Height          =   312
         Left            =   2880
         TabIndex        =   6
         Top             =   360
         Width           =   1812
         _Version        =   1245187
         _ExtentX        =   3196
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
         Alignment       =   1
         Locked          =   -1  'True
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtIntereses 
         Height          =   312
         Left            =   7440
         TabIndex        =   7
         Top             =   360
         Width           =   1812
         _Version        =   1245187
         _ExtentX        =   3196
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
         Alignment       =   1
         Locked          =   -1  'True
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtPoliza 
         Height          =   312
         Left            =   7440
         TabIndex        =   8
         Top             =   1440
         Width           =   1812
         _Version        =   1245187
         _ExtentX        =   3196
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
         Alignment       =   1
         Locked          =   -1  'True
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtAmortizacion 
         Height          =   312
         Left            =   7440
         TabIndex        =   9
         Top             =   720
         Width           =   1812
         _Version        =   1245187
         _ExtentX        =   3196
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
         Alignment       =   1
         Locked          =   -1  'True
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtCargos 
         Height          =   312
         Left            =   7440
         TabIndex        =   10
         Top             =   1080
         Width           =   1812
         _Version        =   1245187
         _ExtentX        =   3196
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
         Alignment       =   1
         Locked          =   -1  'True
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtCbrIntereses 
         Height          =   312
         Left            =   2400
         TabIndex        =   11
         Top             =   4080
         Width           =   1812
         _Version        =   1245187
         _ExtentX        =   3196
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
         Alignment       =   1
         Locked          =   -1  'True
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtTotalAtrasado 
         Height          =   312
         Left            =   7440
         TabIndex        =   12
         Top             =   1800
         Width           =   1812
         _Version        =   1245187
         _ExtentX        =   3196
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
         Alignment       =   1
         Locked          =   -1  'True
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtTotal 
         Height          =   312
         Left            =   2880
         TabIndex        =   37
         Top             =   1800
         Width           =   1812
         _Version        =   1245187
         _ExtentX        =   3196
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
         Alignment       =   1
         Locked          =   -1  'True
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtTasaOriginal 
         Height          =   312
         Left            =   3840
         TabIndex        =   39
         ToolTipText     =   "Original"
         Top             =   1440
         Width           =   852
         _Version        =   1245187
         _ExtentX        =   1503
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
         Appearance      =   2
         UseVisualStyle  =   0   'False
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit txtPlazoOriginal 
         Height          =   312
         Left            =   3840
         TabIndex        =   40
         ToolTipText     =   "Original"
         Top             =   1080
         Width           =   852
         _Version        =   1245187
         _ExtentX        =   1503
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
         Appearance      =   2
         UseVisualStyle  =   0   'False
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit txtPlazo 
         Height          =   312
         Left            =   2880
         TabIndex        =   41
         Top             =   1080
         Width           =   972
         _Version        =   1245187
         _ExtentX        =   1714
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
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtTasa 
         Height          =   312
         Left            =   2880
         TabIndex        =   42
         Top             =   1440
         Width           =   972
         _Version        =   1245187
         _ExtentX        =   1714
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
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtDivisa 
         Height          =   312
         Left            =   9240
         TabIndex        =   13
         Top             =   1800
         Width           =   612
         _Version        =   1245187
         _ExtentX        =   1080
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
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin VB.Label lblTasa 
         BackStyle       =   0  'Transparent
         Caption         =   "..."
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
         Left            =   4800
         TabIndex        =   45
         Top             =   1440
         Visible         =   0   'False
         Width           =   2652
      End
      Begin VB.Label Label19 
         BackStyle       =   0  'Transparent
         Caption         =   "Plazo"
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
         Left            =   960
         TabIndex        =   44
         Top             =   1080
         Width           =   1212
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Tasa"
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
         Left            =   960
         TabIndex        =   43
         Top             =   1440
         Width           =   1692
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Total Reversión"
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
         Left            =   960
         TabIndex        =   38
         Top             =   1800
         Width           =   1692
      End
      Begin VB.Image imgCalculoInt 
         Height          =   252
         Index           =   1
         Left            =   3960
         Picture         =   "frmCO_ReversionCobroJudicial.frx":115A
         Stretch         =   -1  'True
         Top             =   4560
         Width           =   252
      End
      Begin VB.Image imgCalculoInt 
         Height          =   252
         Index           =   0
         Left            =   3600
         Picture         =   "frmCO_ReversionCobroJudicial.frx":1906
         Stretch         =   -1  'True
         Top             =   4560
         Width           =   252
      End
      Begin VB.Label Label2 
         Caption         =   "Intereses a Hoy"
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
         Index           =   18
         Left            =   480
         TabIndex        =   22
         Top             =   4080
         Width           =   1212
      End
      Begin VB.Label Label2 
         Caption         =   "Corte Intereses"
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
         Index           =   21
         Left            =   480
         TabIndex        =   21
         Top             =   4572
         Width           =   1812
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Total (Atrasado)"
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
         Index           =   17
         Left            =   5520
         TabIndex        =   20
         Top             =   1800
         Width           =   1932
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Principal atrasado"
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
         Index           =   11
         Left            =   5520
         TabIndex        =   19
         Top             =   720
         Width           =   1692
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Cargos registrados"
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
         Index           =   16
         Left            =   5520
         TabIndex        =   18
         Top             =   1080
         Width           =   1692
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Pólizas atrasadas"
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
         Index           =   19
         Left            =   5520
         TabIndex        =   17
         Top             =   1440
         Width           =   1692
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Pendiente/Vencido"
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
         Left            =   960
         TabIndex        =   16
         Top             =   720
         Width           =   1692
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Interes Pendientes"
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
         Index           =   22
         Left            =   5520
         TabIndex        =   15
         Top             =   360
         Width           =   1692
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
         Index           =   1
         Left            =   960
         TabIndex        =   14
         Top             =   360
         Width           =   1332
      End
   End
   Begin XtremeSuiteControls.FlatEdit txtOperacion 
      Height          =   432
      Left            =   2880
      TabIndex        =   23
      Top             =   120
      Width           =   1812
      _Version        =   1245187
      _ExtentX        =   3196
      _ExtentY        =   762
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
      Locked          =   -1  'True
      Appearance      =   2
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtProceso 
      Height          =   432
      Left            =   4680
      TabIndex        =   24
      Top             =   120
      Width           =   2052
      _Version        =   1245187
      _ExtentX        =   3619
      _ExtentY        =   762
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
      Locked          =   -1  'True
      Appearance      =   2
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtOpex 
      Height          =   432
      Left            =   6720
      TabIndex        =   25
      Top             =   120
      Width           =   1092
      _Version        =   1245187
      _ExtentX        =   1926
      _ExtentY        =   762
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
      Locked          =   -1  'True
      Appearance      =   2
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtCodigo 
      Height          =   312
      Left            =   2880
      TabIndex        =   26
      Top             =   600
      Width           =   1812
      _Version        =   1245187
      _ExtentX        =   3196
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
      Appearance      =   2
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtCedula 
      Height          =   312
      Left            =   2880
      TabIndex        =   27
      Top             =   960
      Width           =   1812
      _Version        =   1245187
      _ExtentX        =   3196
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
      Appearance      =   2
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtNombre 
      Height          =   312
      Left            =   4680
      TabIndex        =   28
      Top             =   960
      Width           =   6012
      _Version        =   1245187
      _ExtentX        =   10604
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
      Appearance      =   2
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtDescripcion 
      Height          =   312
      Left            =   4680
      TabIndex        =   29
      Top             =   600
      Width           =   6012
      _Version        =   1245187
      _ExtentX        =   10604
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
      Appearance      =   2
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.GroupBox gbTraslado 
      Height          =   1692
      Left            =   240
      TabIndex        =   33
      Top             =   3720
      Width           =   10932
      _Version        =   1245187
      _ExtentX        =   19283
      _ExtentY        =   2984
      _StockProps     =   79
      Caption         =   "Plan de Reversión de deuda"
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   16
      BorderStyle     =   1
      Begin XtremeSuiteControls.FlatEdit txtHonorarios 
         Height          =   312
         Left            =   2880
         TabIndex        =   35
         Top             =   480
         Width           =   1812
         _Version        =   1245187
         _ExtentX        =   3196
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
         Alignment       =   1
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtNotas 
         Height          =   792
         Left            =   2880
         TabIndex        =   36
         Top             =   840
         Width           =   6372
         _Version        =   1245187
         _ExtentX        =   11239
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
         Appearance      =   2
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Notas"
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
         Left            =   960
         TabIndex        =   2
         Top             =   840
         Width           =   852
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Honorarios"
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
         Left            =   960
         TabIndex        =   34
         Top             =   480
         Width           =   1332
      End
   End
   Begin VB.Label Label1 
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
      ForeColor       =   &H00FFFFFF&
      Height          =   252
      Index           =   0
      Left            =   1440
      TabIndex        =   32
      Top             =   600
      Width           =   1572
   End
   Begin VB.Label Label1 
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
      ForeColor       =   &H00FFFFFF&
      Height          =   252
      Index           =   3
      Left            =   1440
      TabIndex        =   31
      Top             =   960
      Width           =   1332
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Operación"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   10.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   252
      Index           =   11
      Left            =   1440
      TabIndex        =   30
      Top             =   120
      Width           =   1572
   End
   Begin VB.Image imgBanner 
      Height          =   1332
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12612
   End
End
Attribute VB_Name = "frmCO_ReversionCobroJudicial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mOperacion As Long, mcurIntCor As Currency, mcurIntMor As Currency, mCurPrincipal As Currency
Dim pTipoDoc As String, pDocumento As String
Dim mTramite As Long

Private Sub btnPrincipal_Click(Index As Integer)
Dim i As Byte


Call sbSIFCleanTxtInject(txtNotas)


Select Case Index
  
  Case 0 'Reversar
           i = MsgBox("Esta Seguro de realizar la reversión de Cobro Judicial de la deuda?", vbYesNo)
           If i = vbYes Then
              Call sbReversaCobroJudicial
           End If
 
  Case 1 'Cerrar
    Unload Me

End Select
 
End Sub

Private Sub Form_Load()

vModulo = 4

Set imgBanner.Picture = frmContenedor.imgBanner_01.Picture

mOperacion = GLOBALES.gTag

Call sbConsulta
 
End Sub

Private Sub sbConsulta()
Dim strSQL As String, rs As New ADODB.Recordset
Dim i As Integer

On Error GoTo vError

txtOperacion = mOperacion

strSQL = "SELECT COD_TRAMITE FROM CBR_CJ_TRAMITE where ID_SOLICITUD  = " & mOperacion
Call OpenRecordSet(rs, strSQL)

If Not rs.EOF Then
    mTramite = rs!COD_TRAMITE
Else
    mTramite = 0
End If

rs.Close



'Se supone que si entra en esta ventana es porque esta previo validado en CO_PRINCIPAL
'Activa toda la barra

'Consulta de Parametros
txtPlazo.Text = CStr(fxCBRPlazoRestante(mOperacion))


'Consulta el estado de la operación
If GLOBALES.SysPlanPagos = 1 Then
    strSQL = "select R.cedula,S.nombre,R.saldo,R.proceso,R.Interesv as Tasa,R.plazo,R.Int as TasaOriginal" _
           & ",isnull(sum(V.intCor + V.intMor),0) as Intereses,R.codigo,C.descripcion,R.Opex" _
           & ",isnull(sum(V.intCor),0) as MoraIntC,isnull(sum(V.intMor),0) as MoraIntM,isnull(sum(V.Principal),0) as MoraAmortiza" _
           & ",isnull(sum(V.Cargos),0) as 'Cargos', isnull(sum(V.Poliza),0) as 'Poliza',R.COD_DIVISA" _
           & " from Socios S inner join reg_creditos R on S.cedula = R.cedula inner join Catalogo C on R.codigo = C.codigo" _
           & " left join crd_operacion_Transac V on R.id_solicitud = V.id_solicitud and V.estado = 'A'" _
           & " Where R.id_solicitud = " & mOperacion _
           & " Group by R.cedula,S.nombre,R.saldo,R.proceso,R.Interesv,R.plazo,R.Int,R.codigo,C.descripcion,R.Opex,R.COD_DIVISA"
Else
    strSQL = "select R.cedula,S.nombre,R.saldo,R.proceso,R.Interesv as Tasa,R.plazo,R.Int as TasaOriginal" _
           & ",isnull(sum(V.intc + V.intm),0) as Intereses,R.codigo,C.descripcion,R.Opex" _
           & ",isnull(sum(V.intc),0) as MoraIntC,isnull(sum(V.intm),0) as MoraIntM,isnull(sum(V.amortiza),0) as MoraAmortiza" _
           & ",isnull(sum(V.Cargo),0) as 'Cargos', 0 as 'Poliza',R.COD_DIVISA" _
           & " from Socios S inner join reg_creditos R on S.cedula = R.cedula inner join Catalogo C on R.codigo = C.codigo" _
           & " left join morosidad V on R.id_solicitud = V.id_solicitud and V.estado = 'A'" _
           & " Where R.id_solicitud = " & mOperacion _
           & " Group by R.cedula,S.nombre,R.saldo,R.proceso,R.Interesv,R.plazo,R.Int,R.codigo,C.descripcion,R.Opex,R.COD_DIVISA"
End If
       
       
Call OpenRecordSet(rs, strSQL)
  
  
  mCurPrincipal = rs!MoraAmortiza
  
  txtHonorarios.Text = fxHonorarios(txtOperacion.Text)
  
  txtCedula.Text = rs!Cedula
  txtNombre.Text = rs!Nombre
  txtDivisa.Text = rs!cod_Divisa & ""
  
  txtTasa.Text = Format(rs!Tasa, "Standard")
  
  txtSaldo.Text = Format(rs!Saldo, "Standard")
  txtInteresesCorte.Text = Format(rs!Intereses, "Standard")
  
  txtCargos.Text = Format(rs!Cargos, "Standard")
  txtPoliza.Text = Format(rs!Poliza, "Standard")
  
  txtOperacion.Tag = rs!opex
  
  If rs!opex = 1 Then
    txtOpex.Text = "Sí"
  Else
    txtOpex.Text = "No"
  End If
   
    btnPrincipal.Item(0).Enabled = False
    Select Case rs!Proceso
      Case "N"
            txtProceso.Text = "Normal"
            
      Case "T" 'Traspaso
            txtProceso.Text = "Traslado"
      
      Case "J", "C" 'Cobro Judicial
            btnPrincipal.Item(0).Enabled = True
            txtProceso.Text = "Cobro Judicial"
      
      Case Else
            txtProceso.Text = "Incobrable"
    End Select
    
    
   txtCodigo.Text = rs!Codigo
   txtDescripcion.Text = rs!Descripcion
       
   txtTasaOriginal.Text = Format(rs!TasaOriginal, "Standard")
   txtPlazoOriginal.Text = CStr(rs!Plazo)
       
rs.Close

strSQL = "exec spCbrCobroJudicialInteresesHoy " & mOperacion
Call OpenRecordSet(rs, strSQL)
    mcurIntCor = rs!RegIntCor
    mcurIntMor = rs!RegIntMor
    mCurPrincipal = rs!RegPrincipal
    
    txtIntereses.Text = Format(rs!RegIntCor + rs!RegIntMor, "Standard")
    txtAmortizacion.Text = Format(rs!RegPrincipal, "Standard")
    
    txtTotalAtrasado.Text = Format(rs!RegIntCor + rs!RegIntMor + rs!RegPrincipal _
                            + CCur(txtCargos.Text) + CCur(txtPoliza.Text), "Standard")
    txtTotal.Text = Format(rs!RegIntCor + rs!RegIntMor + CCur(txtSaldo.Text) _
                            + CCur(txtCargos.Text) + CCur(txtPoliza.Text) + CCur(txtHonorarios.Text), "Standard")
rs.Close

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
 

End Sub


Private Function fxVerificar() As Boolean
Dim strSQL As String, rs As New ADODB.Recordset
Dim vMensaje As String

On Error GoTo vError

vMensaje = ""

txtPlazo.Text = CStr(CLng(txtPlazo.Text))

If CLng(txtPlazo.Text) > 300 Or CLng(txtPlazo.Text) < 1 Then
   vMensaje = vMensaje & " - El plazo es incorrecto verifique..." & vbCrLf
End If

If CCur(txtHonorarios.Text) < 0 Then
   vMensaje = vMensaje & " - Monto de Honorarios no es válido, verifique!!!" & vbCrLf
End If

If CCur(txtTasa.Text) > 100 Or CLng(txtTasa.Text) < 0 Then
   vMensaje = vMensaje & " - La Tasa es incorrecta verifique..." & vbCrLf
End If

If CCur(txtTotal.Text) = 0 Then
   vMensaje = vMensaje & " - No existe un monto a reversar..." & vbCrLf
End If

If Len(Trim(txtNotas)) = 0 Then
   vMensaje = vMensaje & " - Especifique una Nota para la reversion..." & vbCrLf
End If


'Verifica que la Operacion se encuentre en Proceso Judicial / Para evitar accidentes
strSQL = "select isnull(count(*),0) as Existe from reg_creditos where proceso in('J','C') and id_solicitud= '" & txtOperacion.Text & "'"
Call OpenRecordSet(rs, strSQL)
If rs!Existe = 0 Then
   vMensaje = vMensaje & " - La Operación no se encuentra en PROCESO DE COBRO JUDICIAL para realizar la reversión..." & vbCrLf
End If
rs.Close


If Len(vMensaje) > 0 Then
  fxVerificar = False
  MsgBox vMensaje, vbExclamation
Else
  fxVerificar = True
End If

Exit Function

vError:
  fxVerificar = False
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical


End Function

Private Sub sbDocumento(pTipoDocum As String, pNumDoc As Long, pConcepto As String, pCuenta As String, pDetalle As String _
                      , curIntC As Currency, curIntM As Currency, curCargo As Currency, curAmortiza As Currency, curPoliza As Currency)
Dim rs As New ADODB.Recordset, strSQL As String, strLinea(11) As String
Dim strCliente As String


'Cuentas
strSQL = "exec spCrdOperacionCtas " & txtOperacion.Text
Call OpenRecordSet(rs, strSQL)


strCliente = Trim(txtCedula.Text) & " - " & Trim(txtNombre.Text)
strCliente = Mid(strCliente, 1, 45)


strLinea(1) = "Saldo Anterior    " & Format(rs!Saldo - curAmortiza, "Standard")
strLinea(2) = "Interes Corriente " & Format(curIntC * -1, "Standard")
strLinea(3) = "Interes Moratorio " & Format(curIntM * -1, "Standard")
strLinea(4) = "Cargos            " & Format(curCargo * -1, "Standard")
strLinea(5) = "Amortización      " & Format(curAmortiza * -1, "Standard")
strLinea(6) = "Saldo Actual      " & Format(rs!Saldo, "Standard")
strLinea(7) = "Operación         " & txtOperacion.Text
strLinea(8) = "Línea             " & txtCodigo.Text
strLinea(9) = "Proc.Retencion    " & "NO"
strLinea(10) = "Usuario           " & glogon.Usuario
strLinea(11) = "Póliza            " & Format(curPoliza * -1, "Standard")


    strSQL = "insert SIF_TRANSACCIONES(COD_TRANSACCION,TIPO_DOCUMENTO,REGISTRO_FECHA,REGISTRO_USUARIO,Cliente_IDENTIFICACION,CLIENTE_NOMBRE" _
            & ",cod_concepto,monto,estado,Referencia_01,Referencia_02,Referencia_03,cod_oficina" _
            & ",linea1,linea2,linea3,linea4,linea5,linea6,linea7,linea8,linea9,linea10,detalle,documento,linea11)" _
            & " values('" & pNumDoc & "','" & pTipoDocum & "',dbo.MyGetdate(),'" & glogon.Usuario & "','" & Trim(txtCedula.Text) _
            & "','" & Trim(txtNombre.Text) & "','" & pConcepto & "'," & curIntC + curIntM + curAmortiza + curCargo + curPoliza & ",'P','" & txtOperacion.Text _
            & "','" & txtCodigo.Text & "','" & vAseDocDeposito & "','" & GLOBALES.gOficinaTitular & "','" & strLinea(1) & "','" _
            & strLinea(2) & "','" & strLinea(3) & "','" & strLinea(4) & "','" _
            & strLinea(5) & "','" & strLinea(6) & "','" & strLinea(7) & "','" _
            & strLinea(8) & "','" & strLinea(9) & "','" & strLinea(10) & "','" _
            & vAseDocDetalle & "','" & vAseDocDeposito & "','" & strLinea(11) & "')"
    Call ConectionExecute(strSQL)
    
    'ASIENTO
    If curIntC + curIntM + curAmortiza + curCargo + curPoliza > 0 Then
      strSQL = "exec spSIFDocsAsiento '" & pTipoDocum & "','" & pNumDoc & "'," & curIntC + curIntM + curCargo + curAmortiza + curPoliza & ",'C','" & rs!cod_Divisa _
             & "',1," & GLOBALES.gEnlace & ",'" & rs!cod_unidad & "','" & rs!cod_centro_costo & "','" & pCuenta _
             & "','" & rs!id_solicitud & "','" & rs!Codigo & "','" & vAseDocDeposito & "'"
      Call ConectionExecute(strSQL)
    End If
    
    
    If curIntC > 0 Then
      strSQL = "exec spSIFDocsAsiento '" & pTipoDocum & "','" & pNumDoc & "'," & curIntC & ",'D','" & rs!cod_Divisa _
             & "',1," & GLOBALES.gEnlace & ",'" & rs!cod_unidad & "','" & rs!cod_centro_costo & "','" & rs!ctaintc _
             & "','" & rs!id_solicitud & "','" & rs!Codigo & "','" & vAseDocDeposito & "'"
      Call ConectionExecute(strSQL)
    End If
    
    If curIntM > 0 Then
      strSQL = "exec spSIFDocsAsiento '" & pTipoDocum & "','" & pNumDoc & "'," & curIntM & ",'D','" & rs!cod_Divisa _
             & "',1," & GLOBALES.gEnlace & ",'" & rs!cod_unidad & "','" & rs!cod_centro_costo & "','" & rs!ctaintm _
             & "','" & rs!id_solicitud & "','" & rs!Codigo & "','" & vAseDocDeposito & "'"
      Call ConectionExecute(strSQL)
    End If
    
    If curCargo > 0 Then
      strSQL = "exec spSIFDocsAsiento '" & pTipoDocum & "','" & pNumDoc & "'," & curCargo & ",'D','" & rs!cod_Divisa _
             & "',1," & GLOBALES.gEnlace & ",'" & rs!cod_unidad & "','" & rs!cod_centro_costo & "','" & rs!CtaCargos _
             & "','" & rs!id_solicitud & "','" & rs!Codigo & "','" & vAseDocDeposito & "'"
      Call ConectionExecute(strSQL)
    End If
    
    
    If curAmortiza > 0 Then
      strSQL = "exec spSIFDocsAsiento '" & pTipoDocum & "','" & pNumDoc & "'," & curAmortiza & ",'D','" & rs!cod_Divisa _
             & "',1," & GLOBALES.gEnlace & ",'" & rs!cod_unidad & "','" & rs!cod_centro_costo & "','" & rs!ctaamortiza _
             & "','" & rs!id_solicitud & "','" & rs!Codigo & "','" & vAseDocDeposito & "'"
      Call ConectionExecute(strSQL)
    End If

rs.Close

End Sub




Private Sub txtHonorarios_Change()

On Error GoTo vError
If IsNumeric(txtHonorarios.Text) Then
    txtTotal.Text = Format(CCur(txtSaldo.Text) + CCur(txtIntereses.Text) + CCur(txtCargos.Text) + CCur(txtPoliza.Text) + CCur(txtHonorarios.Text), "Standard")
    txtTotalAtrasado.Text = Format(CCur(txtAmortizacion.Text) + CCur(txtIntereses.Text) + CCur(txtCargos.Text) + CCur(txtPoliza.Text), "Standard")
Else
    MsgBox "Solamente puede ingresar números..."
    txtHonorarios.Text = fxHonorarios(txtOperacion.Text)
End If
vError:

End Sub

Private Sub txtHonorarios_GotFocus()

On Error GoTo vError
  
  txtHonorarios.Text = CCur(txtHonorarios.Text)
  
vError:

End Sub

Private Sub txtHonorarios_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtNotas.SetFocus
End Sub

Private Sub txtHonorarios_LostFocus()
On Error GoTo vError
  
  txtHonorarios.Text = Format(CCur(txtHonorarios.Text), "Standard")
  
vError:
End Sub

Private Sub sbHonorariosAplica()
Dim strSQL As String, vCuenta As String, lngRecibo As Long
Dim vFecha As Date, vTipo As String, vTipoDoc As String


'Cuenta de Honorarios
If fxCBR_CJ_Parametros("07") = "S" Then
  vCuenta = fxCBR_CJ_Parametros("06")
Else
  vCuenta = fxCBRParametro("15")
End If



vFecha = fxFechaServidor
  
If GLOBALES.SysPlanPagos = 1 Then
   'Ingresa los honorarios como un cargo a la cuenta
    strSQL = "exec spCrdOperacionCargoAdd " & txtOperacion.Text & "," & CCur(txtHonorarios.Text) & ",'" & GLOBALES.gOficinaUnidad _
           & "','" & GLOBALES.gOficinaCentroCosto & "','Honorarios de Reversión de Cobro Judicial','" & glogon.Usuario _
           & "','CO','" & Trim(vCuenta) & "','',0"
    Call ConectionExecute(strSQL)

Else
    'Configuracion del Documento
    vTipo = "ND"
    lngRecibo = fxDocumentoConsecutivo(vTipo)
    vTipoDoc = "ND"
    
 
    
    strSQL = "update reg_creditos set estado = 'A', amortiza = amortiza - " & CCur(txtHonorarios.Text) _
           & ", saldo = saldo + " & CCur(txtHonorarios.Text) _
           & ", saldo_mes = saldo_mes + " & CCur(txtHonorarios) _
           & " where id_solicitud = " & txtOperacion.Text
    Call ConectionExecute(strSQL)
      
    strSQL = "insert creditos_dt(codigo,id_solicitud,cuota,abono,intcp,amortiza,fechas,fechap,estado" _
           & ",tcon,ncon) values('" & txtCodigo.Text & "'," & txtOperacion.Text & ",0,0,0," _
           & CCur(txtHonorarios.Text) & ",'" & Format(vFecha, "yyyy/mm/dd") & "'" _
           & "," & GLOBALES.glngFechaCR & ",'N','" & vTipoDoc & "','" & IIf((lngRecibo = 0), "null", lngRecibo) & "')"
    Call ConectionExecute(strSQL)
      
    Call sbDocumento(vTipo, lngRecibo, "CBR009", vCuenta, "COBRO DE HONORARIOS", 0, 0, 0, CCur(txtHonorarios.Text), 0)
       
    If lngRecibo > 0 Then Call sbImprimeRecibo(lngRecibo, "ND")
End If
 
 
End Sub


Private Sub sbCobroJudicialRevCntDocv2()
'-------------------------------------------------------------------------------------------
'OBJETIVO      : Ejecuta el Cobro Judicial a una Operación
'REFERENCIAS   : FxFechaServidor - (Devuelve la Fecha del servidor)
'                Bitacora - (Registra el movimiento realizado)
'OBSERVACIONES : Genera Asiento
'-------------------------------------------------------------------------------------------

Dim rs As New ADODB.Recordset, strSQL As String, strCuentas As String
Dim strObservacion As String, vFecha As Date, strLinea(11) As String
Dim vOficina As String, vUnidad As String, vDivisa  As String, vCuenta As String, vCuentaCbr As String
Dim vCentroCosto As String, pConcepto As String, vTipoCambio As Currency

Me.MousePointer = vbHourglass

On Error GoTo vError


'Extrae la Cuenta de Cobro Judicial y la Fecha
strSQL = "select CtaCAmort as 'Cuenta',dbo.MyGetdate() as Fecha from catalogo where codigo = '" & txtCodigo.Text & "'"
Call OpenRecordSet(rs, strSQL)
 vCuentaCbr = Trim(rs!Cuenta)
 vFecha = rs!fecha
rs.Close

'Extrae configuración Contable de la Operación
strSQL = "exec spCrdOperacionCtas " & txtOperacion.Text
Call OpenRecordSet(rs, strSQL)
 vCuenta = Trim(rs!ctaamortiza)
 vOficina = Trim(rs!cod_oficina_r)
 vUnidad = Trim(rs!cod_unidad)
 vDivisa = Trim(rs!cod_Divisa)
 vCentroCosto = ""
 vTipoCambio = rs!TipoCambio
rs.Close

'Otro parámetros contables
pTipoDoc = "CBJ"
pDocumento = ""
pConcepto = "CBR005"

vAseDocCuenta = ""
vAseDocDeposito = ""
vAseDocDetalle = strObservacion

pDocumento = fxDocumentoConsecutivo(pTipoDoc)


'Lineas de Comprobante
strLinea(1) = "Saldo Actual      " & txtSaldo.Text
strLinea(2) = "Interes Corriente " & Format(mcurIntCor, "Standard")
strLinea(3) = "Interes Atrasado  " & Format(mcurIntMor, "Standard")
strLinea(4) = "Amortización Atra." & Format(mCurPrincipal, "Standard")
strLinea(5) = "Cargos Regist.    " & txtCargos.Text
strLinea(6) = "Divisa.: " & vDivisa & "/ Tipo Cambio.:" & vTipoCambio
strLinea(7) = "Operacion/Línea   " & "Op.:" & txtOperacion.Text & " Lí.:" & txtCodigo.Text
strLinea(8) = Mid(Trim(txtDescripcion.Text), 1, 30)
strLinea(9) = ""
strLinea(10) = Mid("Notas: " & txtNotas.Text, 1, 30)
strLinea(11) = "Póliza Atradada  " & txtPoliza.Text
 

'Registro del Comprobante
strSQL = "insert SIF_TRANSACCIONES(COD_TRANSACCION,TIPO_DOCUMENTO,REGISTRO_FECHA,REGISTRO_USUARIO,Cliente_IDENTIFICACION,CLIENTE_NOMBRE" _
         & ",cod_concepto,monto,estado,Referencia_01,Referencia_02,Referencia_03,cod_oficina" _
         & ",linea1,linea2,linea3,linea4,linea5,linea6,linea7,linea8,linea9,linea10,linea11,detalle,documento)" _
         & " values('" & pDocumento & "','" & pTipoDoc & "',dbo.MyGetdate(),'" & glogon.Usuario & "','" & Trim(txtCedula.Text) _
         & "','" & Trim(txtNombre.Text) & "','" & pConcepto & "'," & CCur(txtSaldo.Text) & ",'P','" & txtOperacion.Text _
         & "','" & txtCodigo.Text & "','" & vAseDocDeposito & "','" & GLOBALES.gOficinaTitular & "','" & strLinea(1) & "','" _
         & strLinea(2) & "','" & strLinea(3) & "','" & strLinea(4) & "','" _
         & strLinea(5) & "','" & strLinea(6) & "','" & strLinea(7) & "','" _
         & strLinea(8) & "','" & strLinea(9) & "','" & strLinea(10) & "','" _
         & strLinea(11) & "','" & vAseDocDetalle & "','" & vAseDocDeposito & "')"
Call ConectionExecute(strSQL)

'ASIENTO

If CCur(txtSaldo.Text) > 0 Then
  strSQL = "exec spSIFDocsAsiento '" & pTipoDoc & "','" & pDocumento & "'," & CCur(txtSaldo.Text) * vTipoCambio & ",'D','" & vDivisa _
         & "'," & vTipoCambio & "," & GLOBALES.gEnlace & ",'" & vUnidad & "','" & vCentroCosto & "','" & vCuenta _
         & "','" & txtOperacion.Text & "','" & txtCodigo.Text & "','" & vAseDocDeposito & "'"
  Call ConectionExecute(strSQL)

  strSQL = "exec spSIFDocsAsiento '" & pTipoDoc & "','" & pDocumento & "'," & CCur(txtSaldo.Text) * vTipoCambio & ",'C','" & vDivisa _
         & "'," & vTipoCambio & "," & GLOBALES.gEnlace & ",'" & vUnidad & "','" & vCentroCosto & "','" & vCuentaCbr _
         & "','" & txtOperacion.Text & "','" & txtCodigo.Text & "','" & vAseDocDeposito & "'"
  Call ConectionExecute(strSQL)
End If
         
Me.MousePointer = vbDefault

'Control de Documentos v2
Call sbImprimeRecibo(pDocumento, pTipoDoc)


Exit Sub

vError:
    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical


End Sub


Private Sub sbReversaCobroJudicial()
'-------------------------------------------------------------------------------------------
'OBJETIVO      : Reversión de un envio a Cobro Judicial
'REFERENCIAS   : fxFechaServior - (Devuelve fecha del servidor)
'                Bitacora - (Registra el Movimiento efectuado)
'OBSERVACIONES : Ver Readecuacion con Cambio de Operacion (Para Ajustar Nuevos Montos)
'                Genera Asiento
'-------------------------------------------------------------------------------------------

Dim strSQL As String

On Error GoTo vError

 If Not fxVerificar Then
    Exit Sub
 End If

Me.MousePointer = vbHourglass
    
    'Reversión del Cobro Judicial + Ajuste de Existentes y Nuevas Cuotas Morosas
If GLOBALES.SysPlanPagos = 1 Then
    'Pone la Operacion en Estado Normal
    strSQL = "exec spCbrReversaCobroJudicialPlanPagos " & txtOperacion.Text & ",'" & glogon.Usuario & "'"
    Call ConectionExecute(strSQL)
Else
    strSQL = "exec spCbrReversaCobroJudicial " & txtOperacion.Text & ",'" & glogon.Usuario & "'"
    Call ConectionExecute(strSQL)
End If

'Comprobante
Call sbCobroJudicialRevCntDocv2

'Registro en Bitacora General
Call Bitacora("Reversa", "Cobro Judicial a la Operación:" & txtOperacion)
    
'Registro Historial y Expediente
Call sbCBRRegTransac("06", txtCedula, txtOperacion, txtNotas.Text, CCur(txtSaldo), mcurIntCor, mcurIntMor _
                    , CCur(txtCargos.Text), CCur(txtPoliza.Text), CCur(txtAmortizacion.Text), pTipoDoc, pDocumento)

'Si Existen Ajustes en Honorarios / Aplicar Nota de Debito a la Cuenta Reversada
If CCur(txtHonorarios.Text) > 0 Then
   Call sbHonorariosAplica
End If

'inserta proceso de reversion
If mTramite > 0 Then
    strSQL = "INSERT CBR_CJ_TRAMITE_PROCESO(NUM_LINEA,COD_TRAMITE,COD_PROCESO,NOTAS,APLICA_CIERRE_SENTENCIA," _
           & " REGISTRO_FECHA ,REGISTRO_USUARIO) Values(" & fxLineaTramite(mTramite) & "," & mTramite & "," _
           & " '" & fxCBR_CJ_Parametros("09") & "','Aplica Reversión Cobro Judicial',1,dbo.MyGetdate(),'" & glogon.Usuario & "')"
    Call ConectionExecute(strSQL)
    
    'Actualiza el monto y la fecha de la sentencia
    strSQL = "Update CBR_CJ_TRAMITE set SENTENCIA_FECHA = dbo.MyGetdate(),SENTENCIA_MONTO = 0 where COD_TRAMITE = " & mTramite & ""
    Call ConectionExecute(strSQL)
End If




Me.MousePointer = vbDefault

If GLOBALES.SysDocVersion = 1 Then
    MsgBox "- La operación fue reversada a estado NORMAL " & vbCrLf & vbCrLf _
         & "- Se generó Asiento (RCBR" & txtOperacion & ")", vbInformation
Else
    MsgBox "- La operación fue reversada a estado NORMAL " & vbCrLf & vbCrLf _
         & "- Se generó Nota de Cobro: " & pTipoDoc & "-" & pDocumento, vbInformation
End If



Call sbConsulta

Exit Sub

vError:
    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Function fxHonorarios(vOperacion) As Currency
Dim strSQL As String, rs As New ADODB.Recordset

strSQL = "SELECT isnull(sum(monto),0) as 'monto' " _
        & " FROM CBR_CJ_TRAMITE_GASTOS where TESORERIA_NUMERO is not null and cod_tramite = " & mTramite & ""

Call OpenRecordSet(rs, strSQL)

fxHonorarios = rs!Monto

rs.Close

End Function

Private Function fxLineaTramite(vTramite As Long) As Integer
Dim strSQL As String, rs As New ADODB.Recordset

strSQL = "select isnull(max(num_linea),0)+1 as 'Linea'" _
        & " from cbr_cj_tramite_proceso where cod_tramite =" & vTramite & ""
Call OpenRecordSet(rs, strSQL)
fxLineaTramite = rs!Linea
rs.Close
End Function

