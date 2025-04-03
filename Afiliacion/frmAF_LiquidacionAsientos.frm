VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.0#0"; "Codejock.Controls.v22.0.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.0#0"; "Codejock.ShortcutBar.v22.0.0.ocx"
Begin VB.Form frmAF_LiquidacionAsientos 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Liquidaciones: Traspaso a Tesoreria para Desembolsos"
   ClientHeight    =   7410
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15705
   Icon            =   "frmAF_LiquidacionAsientos.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7410
   ScaleWidth      =   15705
   WindowState     =   2  'Maximized
   Begin XtremeSuiteControls.GroupBox fraFiltros 
      Height          =   3135
      Left            =   7200
      TabIndex        =   29
      Top             =   1920
      Visible         =   0   'False
      Width           =   6135
      _Version        =   1441792
      _ExtentX        =   10821
      _ExtentY        =   5530
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      Begin XtremeSuiteControls.PushButton btnFiltros 
         Height          =   375
         Index           =   0
         Left            =   3360
         TabIndex        =   30
         Top             =   2520
         Width           =   1215
         _Version        =   1441792
         _ExtentX        =   2138
         _ExtentY        =   656
         _StockProps     =   79
         Caption         =   "Buscar"
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
         UseVisualStyle  =   -1  'True
         Appearance      =   17
      End
      Begin XtremeSuiteControls.ComboBox cboBanco 
         Height          =   315
         Left            =   1320
         TabIndex        =   31
         Top             =   600
         Width           =   4455
         _Version        =   1441792
         _ExtentX        =   7858
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
      Begin XtremeSuiteControls.ComboBox cboOficina 
         Height          =   315
         Left            =   1320
         TabIndex        =   32
         Top             =   1080
         Width           =   4455
         _Version        =   1441792
         _ExtentX        =   7858
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
      Begin XtremeSuiteControls.ComboBox cboUsuarios 
         Height          =   315
         Left            =   1320
         TabIndex        =   33
         Top             =   1560
         Width           =   4455
         _Version        =   1441792
         _ExtentX        =   7858
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
      Begin XtremeSuiteControls.PushButton btnFiltros 
         Height          =   375
         Index           =   1
         Left            =   4560
         TabIndex        =   34
         Top             =   2520
         Width           =   1215
         _Version        =   1441792
         _ExtentX        =   2138
         _ExtentY        =   656
         _StockProps     =   79
         Caption         =   "Refrescar"
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
         UseVisualStyle  =   -1  'True
         Appearance      =   17
      End
      Begin XtremeSuiteControls.CheckBox chkMarcas 
         Height          =   375
         Left            =   1320
         TabIndex        =   35
         Top             =   2640
         Visible         =   0   'False
         Width           =   1335
         _Version        =   1441792
         _ExtentX        =   2350
         _ExtentY        =   656
         _StockProps     =   79
         Caption         =   "Solo Casos Marcados"
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
      End
      Begin XtremeSuiteControls.ComboBox cboTokenConsulta 
         Height          =   315
         Left            =   1320
         TabIndex        =   36
         Top             =   2040
         Width           =   4455
         _Version        =   1441792
         _ExtentX        =   7858
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
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
         Height          =   375
         Index           =   0
         Left            =   0
         TabIndex        =   41
         Top             =   0
         Width           =   6135
         _Version        =   1441792
         _ExtentX        =   10821
         _ExtentY        =   661
         _StockProps     =   14
         Caption         =   "Filtros adicionalaes:"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   10.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SubItemCaption  =   -1  'True
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Token"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   4
         Left            =   240
         TabIndex        =   40
         Top             =   2040
         Width           =   975
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Usuarios"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   240
         TabIndex        =   39
         Top             =   1560
         Width           =   975
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Oficina"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   240
         TabIndex        =   38
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Banco"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   37
         Top             =   600
         Width           =   975
      End
   End
   Begin XtremeSuiteControls.GroupBox gbReportes 
      Height          =   2895
      Left            =   1440
      TabIndex        =   21
      Top             =   1920
      Visible         =   0   'False
      Width           =   5655
      _Version        =   1441792
      _ExtentX        =   9975
      _ExtentY        =   5106
      _StockProps     =   79
      BackColor       =   16777215
      Appearance      =   16
      Begin XtremeSuiteControls.PushButton btnInforme 
         Height          =   495
         Index           =   0
         Left            =   3600
         TabIndex        =   22
         Top             =   2040
         Width           =   1455
         _Version        =   1441792
         _ExtentX        =   2566
         _ExtentY        =   873
         _StockProps     =   79
         Caption         =   "Informe"
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
         Picture         =   "frmAF_LiquidacionAsientos.frx":6852
      End
      Begin XtremeSuiteControls.RadioButton rbInformes 
         Height          =   255
         Index           =   0
         Left            =   600
         TabIndex        =   23
         Top             =   600
         Width           =   4335
         _Version        =   1441792
         _ExtentX        =   7646
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Resumen"
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
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         Value           =   -1  'True
      End
      Begin XtremeSuiteControls.RadioButton rbInformes 
         Height          =   255
         Index           =   1
         Left            =   600
         TabIndex        =   24
         Top             =   960
         Width           =   4335
         _Version        =   1441792
         _ExtentX        =   7646
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Detalle"
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
         UseVisualStyle  =   -1  'True
         Appearance      =   17
      End
      Begin XtremeSuiteControls.RadioButton rbInformes 
         Height          =   255
         Index           =   2
         Left            =   600
         TabIndex        =   25
         Top             =   1320
         Width           =   4335
         _Version        =   1441792
         _ExtentX        =   7646
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Pendientes"
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
         UseVisualStyle  =   -1  'True
         Appearance      =   17
      End
      Begin XtremeSuiteControls.PushButton btnInforme 
         Height          =   495
         Index           =   1
         Left            =   5040
         TabIndex        =   26
         Top             =   2040
         Width           =   495
         _Version        =   1441792
         _ExtentX        =   873
         _ExtentY        =   873
         _StockProps     =   79
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
         Picture         =   "frmAF_LiquidacionAsientos.frx":6F59
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
         Height          =   375
         Index           =   1
         Left            =   0
         TabIndex        =   27
         Top             =   0
         Width           =   5895
         _Version        =   1441792
         _ExtentX        =   10398
         _ExtentY        =   661
         _StockProps     =   14
         Caption         =   "Informes:"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   10.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SubItemCaption  =   -1  'True
      End
   End
   Begin FPSpreadADO.fpSpread vGrid 
      Height          =   5055
      Left            =   -960
      TabIndex        =   0
      Top             =   1920
      Width           =   16935
      _Version        =   524288
      _ExtentX        =   29871
      _ExtentY        =   8916
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
      MaxCols         =   15
      SpreadDesigner  =   "frmAF_LiquidacionAsientos.frx":766F
      VScrollSpecial  =   -1  'True
      VScrollSpecialType=   2
      AppearanceStyle =   1
   End
   Begin XtremeSuiteControls.ComboBox cboEstado 
      Height          =   312
      Left            =   2160
      TabIndex        =   7
      Top             =   480
      Width           =   1932
      _Version        =   1441792
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
   Begin XtremeSuiteControls.ComboBox cboAccion 
      Height          =   312
      Left            =   2160
      TabIndex        =   8
      Top             =   120
      Width           =   1932
      _Version        =   1441792
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
   Begin XtremeSuiteControls.ComboBox cboTipoRenuncia 
      Height          =   312
      Left            =   2160
      TabIndex        =   9
      Top             =   840
      Width           =   1932
      _Version        =   1441792
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
   Begin XtremeSuiteControls.ComboBox cboTipo 
      Height          =   312
      Left            =   5760
      TabIndex        =   10
      Top             =   120
      Width           =   3972
      _Version        =   1441792
      _ExtentX        =   7011
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
   Begin XtremeSuiteControls.DateTimePicker dtpDesde 
      Height          =   312
      Left            =   5760
      TabIndex        =   11
      Top             =   480
      Width           =   1332
      _Version        =   1441792
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
   Begin XtremeSuiteControls.DateTimePicker dtpHasta 
      Height          =   312
      Left            =   7080
      TabIndex        =   12
      Top             =   480
      Width           =   1332
      _Version        =   1441792
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
   Begin XtremeSuiteControls.CheckBox chkTodos 
      Height          =   252
      Left            =   120
      TabIndex        =   6
      Top             =   1440
      Width           =   972
      _Version        =   1441792
      _ExtentX        =   1714
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Marcar"
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
   End
   Begin XtremeSuiteControls.ComboBox cboToken 
      Height          =   330
      Left            =   5760
      TabIndex        =   13
      Top             =   840
      Width           =   2655
      _Version        =   1441792
      _ExtentX        =   4683
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
   Begin XtremeSuiteControls.PushButton btnTokenNew 
      Height          =   315
      Index           =   2
      Left            =   8520
      TabIndex        =   14
      Top             =   840
      Width           =   855
      _Version        =   1441792
      _ExtentX        =   1508
      _ExtentY        =   556
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
   End
   Begin XtremeSuiteControls.ProgressBar prgBar 
      Height          =   135
      Left            =   0
      TabIndex        =   16
      Top             =   7200
      Width           =   1935
      _Version        =   1441792
      _ExtentX        =   3413
      _ExtentY        =   238
      _StockProps     =   93
      BackColor       =   -2147483633
      Value           =   5
   End
   Begin XtremeSuiteControls.PushButton btnBarra 
      Height          =   375
      Index           =   0
      Left            =   2160
      TabIndex        =   17
      Top             =   1320
      Width           =   1215
      _Version        =   1441792
      _ExtentX        =   2138
      _ExtentY        =   656
      _StockProps     =   79
      Caption         =   "Buscar"
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
      TextAlignment   =   1
      Appearance      =   14
      Picture         =   "frmAF_LiquidacionAsientos.frx":807C
      ImageAlignment  =   4
   End
   Begin XtremeSuiteControls.PushButton btnBarra 
      Height          =   375
      Index           =   1
      Left            =   4680
      TabIndex        =   18
      Top             =   1320
      Width           =   1215
      _Version        =   1441792
      _ExtentX        =   2138
      _ExtentY        =   656
      _StockProps     =   79
      Caption         =   "Generar"
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
      TextAlignment   =   1
      Appearance      =   14
      Picture         =   "frmAF_LiquidacionAsientos.frx":877C
      ImageAlignment  =   4
   End
   Begin XtremeSuiteControls.PushButton btnBarra 
      Height          =   375
      Index           =   2
      Left            =   5880
      TabIndex        =   19
      Top             =   1320
      Width           =   1215
      _Version        =   1441792
      _ExtentX        =   2138
      _ExtentY        =   656
      _StockProps     =   79
      Caption         =   "Informes"
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
      TextAlignment   =   1
      Appearance      =   14
      Picture         =   "frmAF_LiquidacionAsientos.frx":8EA3
      ImageAlignment  =   4
   End
   Begin XtremeSuiteControls.PushButton btnBarra 
      Height          =   375
      Index           =   3
      Left            =   3360
      TabIndex        =   20
      Top             =   1320
      Width           =   1215
      _Version        =   1441792
      _ExtentX        =   2138
      _ExtentY        =   656
      _StockProps     =   79
      Caption         =   "Exportar"
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
      TextAlignment   =   1
      Appearance      =   14
      Picture         =   "frmAF_LiquidacionAsientos.frx":95AA
      ImageAlignment  =   4
   End
   Begin XtremeSuiteControls.ProgressBar ProgressBarX 
      Height          =   135
      Left            =   2160
      TabIndex        =   28
      Top             =   1680
      Visible         =   0   'False
      Width           =   2415
      _Version        =   1441792
      _ExtentX        =   4260
      _ExtentY        =   238
      _StockProps     =   93
      BackColor       =   -2147483633
      Scrolling       =   1
   End
   Begin XtremeSuiteControls.CheckBox chkFiltros 
      Height          =   375
      Left            =   7320
      TabIndex        =   42
      Top             =   1320
      Width           =   1335
      _Version        =   1441792
      _ExtentX        =   2355
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "+ Filtros"
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
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Token"
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
      Height          =   375
      Index           =   4
      Left            =   4920
      TabIndex        =   15
      Top             =   840
      Width           =   615
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo"
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
      Height          =   375
      Index           =   3
      Left            =   1320
      TabIndex        =   5
      Top             =   840
      Width           =   615
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Corte"
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
      Height          =   372
      Index           =   0
      Left            =   4920
      TabIndex        =   4
      Top             =   480
      Width           =   612
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Acción"
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
      Height          =   375
      Index           =   1
      Left            =   1320
      TabIndex        =   3
      Top             =   120
      Width           =   615
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Estado"
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
      Height          =   375
      Index           =   2
      Left            =   1320
      TabIndex        =   2
      Top             =   480
      Width           =   615
   End
   Begin VB.Label lblTipoCaso 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo de Caso"
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
      Height          =   372
      Left            =   4320
      TabIndex        =   1
      Top             =   120
      Width           =   1332
   End
   Begin VB.Image imgBanner 
      Height          =   1212
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   11892
   End
End
Attribute VB_Name = "frmAF_LiquidacionAsientos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vPaso As Boolean, vDuplicado As Boolean


Private Sub sbExportar()

Dim vHeaders As vGridHeaders

On Error GoTo vError

Me.MousePointer = vbHourglass

ProgressBarX.Visible = True
            
            vHeaders.Columnas = 15
            vHeaders.Headers(1) = "..."
            vHeaders.Headers(2) = "No.Liq"
            vHeaders.Headers(3) = "Identificación"
            vHeaders.Headers(4) = "Nombre"
            vHeaders.Headers(5) = "Total Neto"
            vHeaders.Headers(6) = "Id. Banco"
            vHeaders.Headers(7) = "Emite?"
            vHeaders.Headers(8) = "Estado"
            vHeaders.Headers(9) = "Cuenta"
            vHeaders.Headers(10) = "Fecha"
            vHeaders.Headers(11) = "Usuario"
            vHeaders.Headers(12) = "Banco"
            vHeaders.Headers(13) = "Duplicado?"
            vHeaders.Headers(14) = "Divisa"
            vHeaders.Headers(15) = "Token"
        
         Call sbSIFGridExportar(vGrid, vHeaders, "ProGrX_Liquidaciones_Traslado_Bancos")
    

ProgressBarX.Visible = False

Me.MousePointer = vbDefault

Exit Sub

vError:
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub btnBarra_Click(Index As Integer)

Select Case Index
  Case 0 'Buscar
        Call sbBuscar

  Case 1 'Generar
        Call sbGenerar

   Case 2 'Informes
   
   gbReportes.top = vGrid.top
   gbReportes.Left = 1440
   gbReportes.Visible = IIf((gbReportes.Visible = True), False, True)
   
   
   Case 3 'Exportar
    Call sbExportar
End Select
End Sub

Private Sub btnFiltros_Click(Index As Integer)
Select Case Index
  Case 0  'Buscar
    Call sbBuscar
    fraFiltros.Visible = False
    chkFiltros.Value = vbUnchecked
  
  Case 1 'Refrescar
    Call sbFiltros
    
End Select

End Sub

Private Sub btnInforme_Click(Index As Integer)
Select Case Index
Case 0 'Informe
    Select Case True
        Case rbInformes.Item(0).Value
            Call sbInformes("Resumen")
        Case rbInformes.Item(1).Value
            Call sbInformes("Detalle")
        Case rbInformes.Item(2).Value
            Call sbInformes("Pendientes")
    End Select

Case 1 'Cerrar
    gbReportes.Visible = False

End Select

End Sub

Private Sub btnTokenNew_Click(Index As Integer)
Dim strSQL As String

On Error GoTo vError

strSQL = "exec spTes_Token_New '" & glogon.Usuario & "'"
Call ConectionExecute(strSQL)

Call sbTokens_Load

Exit Sub

vError:
End Sub

Private Sub cboAccion_Click()
Dim strSQL As String

If vPaso Then Exit Sub

vGrid.MaxRows = 0

vPaso = True

If Mid(cboAccion.Text, 1, 1) = "D" Then
  'D = Desembolsar
  strSQL = "select id_banco as Idx, rtrim(descripcion) + '  ' + rtrim(Cta) as ItmX from Tes_Bancos where estado = 'A'"
  Call sbCbo_Llena_New(cboTipo, strSQL, True, True)
  lblTipoCaso.Caption = "Bancos...:"
Else
  'R = Retener
  lblTipoCaso.Caption = "Retener por.:"
  
  strSQL = "select RTRIM(RETENCION_CODIGO) as 'IdX', RTRIM(RETENCION_CODIGO) + ' - ' + rtrim(descripcion) + ' [' + rtrim(COD_CUENTA) + ']' as ItmX" _
         & " from FND_RETENCION_CONCEPTOS where ACTIVO = 1"
  Call sbCbo_Llena_New(cboTipo, strSQL, False)

End If

vPaso = False

End Sub


Private Sub cboEstado_Click()

If vPaso Then Exit Sub

vGrid.MaxRows = 0

End Sub


Private Sub cboTipo_Click()
If vPaso Then Exit Sub

If Mid(cboAccion.Text, 1, 1) = "D" Then
    vGrid.MaxRows = 0
End If

End Sub


Private Sub sbFiltros()
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'OBJETIVO:      Llena combos con filtros adicionales segun los rangos de solicitud base    .
'REFERENCIAS:   Ninguna.
'OBSERVACIONES: Ninguna.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim strSQL As String, pRevision As String

If dtpHasta.Value < dtpDesde.Value Then
   MsgBox "Verifique el Rango de Fechas", vbInformation, "Error"
   Exit Sub
End If

Me.MousePointer = vbHourglass

pRevision = fxSIFParametros("15")

'Cargado de Bancos
strSQL = "Select L.cod_banco as Idx,isnull(B.descripcion,'Sin Banco') as Itmx" _
        & " From Liquidacion L left join Tes_Bancos B on L.cod_Banco = B.id_Banco" _
        & " Where L.FecLiq between '" & Format(dtpDesde.Value, "yyyy/mm/dd") & " 00:00:00' and '" _
        & Format(dtpHasta.Value, "yyyy/mm/dd") & " 23:59:59'"
        
        
If pRevision = "S" Then
   strSQL = strSQL & " and L.Analista_Revision = 'S'"
End If

If Mid(cboAccion.Text, 1, 1) = "D" And cboTipo.Text <> "TODOS" Then
  strSQL = strSQL & " And L.cod_Banco = " & cboTipo.ItemData(cboTipo.ListIndex)
End If

If Mid(cboEstado.Text, 1, 1) = "P" Then
  strSQL = strSQL & " And L.EstadoAsiento='P'"
Else
  strSQL = strSQL & " And L.EstadoAsiento='G'"
End If

If Mid(cboTipoRenuncia.Text, 1, 1) <> "T" Then
        If Mid(cboTipoRenuncia.Text, 1, 1) = "A" Then
          strSQL = strSQL & " And L.ESTADOACTLIQ='A'"
        Else
          strSQL = strSQL & " And L.ESTADOACTLIQ='P'"
        End If
End If

strSQL = strSQL & " Group by L.cod_banco,B.descripcion"

Call sbCbo_Llena_New(cboBanco, strSQL, True, True)
        

'Cargado de Usuarios
strSQL = "Select L.USUARIO as 'idX', L.USUARIO as 'Itmx'" _
        & " From Liquidacion L" _
        & " Where L.FecLiq between '" & Format(dtpDesde.Value, "yyyy/mm/dd") & " 00:00:00' and '" _
        & Format(dtpHasta.Value, "yyyy/mm/dd") & " 23:59:59'"
        
If pRevision = "S" Then
   strSQL = strSQL & " and L.Analista_Revision = 'S'"
End If

If Mid(cboAccion.Text, 1, 1) = "D" And cboTipo.Text <> "TODOS" Then
  strSQL = strSQL & " And L.cod_Banco = " & cboTipo.ItemData(cboTipo.ListIndex)
End If

If Mid(cboEstado.Text, 1, 1) = "P" Then
  strSQL = strSQL & " And L.EstadoAsiento='P'"
Else
  strSQL = strSQL & " And L.EstadoAsiento='G'"
End If

If Mid(cboTipoRenuncia.Text, 1, 1) <> "T" Then
        If Mid(cboTipoRenuncia.Text, 1, 1) = "A" Then
          strSQL = strSQL & " And L.ESTADOACTLIQ='A'"
        Else
          strSQL = strSQL & " And L.ESTADOACTLIQ='P'"
        End If
End If

strSQL = strSQL & " Group by L.usuario"
Call sbCbo_Llena_New(cboUsuarios, strSQL, True, True)




'Cargado de Token
strSQL = "Select ISNULL(L.ID_TOKEN,'') as 'Itmx', ISNULL(L.ID_TOKEN,'') as 'idX'" _
        & " From Liquidacion L" _
        & " Where L.FecLiq between '" & Format(dtpDesde.Value, "yyyy/mm/dd") & " 00:00:00' and '" _
        & Format(dtpHasta.Value, "yyyy/mm/dd") & " 23:59:59'"
        
If pRevision = "S" Then
   strSQL = strSQL & " and L.Analista_Revision = 'S'"
End If

If Mid(cboAccion.Text, 1, 1) = "D" And cboTipo.Text <> "TODOS" Then
  strSQL = strSQL & " And L.cod_Banco = " & cboTipo.ItemData(cboTipo.ListIndex)
End If

If Mid(cboEstado.Text, 1, 1) = "P" Then
  strSQL = strSQL & " And L.EstadoAsiento='P'"
Else
  strSQL = strSQL & " And L.EstadoAsiento='G'"
End If

If Mid(cboTipoRenuncia.Text, 1, 1) <> "T" Then
        If Mid(cboTipoRenuncia.Text, 1, 1) = "A" Then
          strSQL = strSQL & " And L.ESTADOACTLIQ='A'"
        Else
          strSQL = strSQL & " And L.ESTADOACTLIQ='P'"
        End If
End If

strSQL = strSQL & " Group by ISNULL(L.ID_TOKEN,'')"
Call sbCbo_Llena_New(cboTokenConsulta, strSQL, True, True)


'Cargado de Oficinas
strSQL = "Select rtrim(L.cod_Oficina) as 'idX',  isnull(O.descripcion,'') as 'Itmx'" _
        & " From Liquidacion L left join SIF_Oficinas O on L.cod_oficina = O.cod_oficina" _
        & " Where L.FecLiq between '" & Format(dtpDesde.Value, "yyyy/mm/dd") & " 00:00:00' and '" _
        & Format(dtpHasta.Value, "yyyy/mm/dd") & " 23:59:59'"

If pRevision = "S" Then
   strSQL = strSQL & " and L.Analista_Revision = 'S'"
End If

If Mid(cboAccion.Text, 1, 1) = "D" And cboTipo.Text <> "TODOS" Then
  strSQL = strSQL & " And L.cod_Banco = " & cboTipo.ItemData(cboTipo.ListIndex)
End If

If Mid(cboEstado.Text, 1, 1) = "P" Then
  strSQL = strSQL & " And L.EstadoAsiento='P'"
Else
  strSQL = strSQL & " And L.EstadoAsiento='G'"
End If

If Mid(cboTipoRenuncia.Text, 1, 1) <> "T" Then
        If Mid(cboTipoRenuncia.Text, 1, 1) = "A" Then
          strSQL = strSQL & " And L.ESTADOACTLIQ='A'"
        Else
          strSQL = strSQL & " And L.ESTADOACTLIQ='P'"
        End If
End If

strSQL = strSQL & " Group by L.cod_Oficina,O.descripcion"

Call sbCbo_Llena_New(cboOficina, strSQL, True, True)


Me.MousePointer = vbDefault


End Sub



Private Sub chkFiltros_Click()
If chkFiltros.Value = vbChecked Then
   fraFiltros.Visible = True
   fraFiltros.top = vGrid.top
   fraFiltros.Left = chkFiltros.Left
   Call sbFiltros
Else
   fraFiltros.Visible = False
End If
End Sub

Private Sub chkTodos_Click()
Dim i As Long

For i = 1 To vGrid.MaxRows
 vGrid.Row = i
 vGrid.col = 1
 vGrid.Value = chkTodos.Value
Next i

End Sub


Private Sub sbBuscar()
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'OBJETIVO:      Busca Liquidaciones pendientes o generadas con ubicacion en tesoreria.
'REFERENCIAS:   Ninguna.
'OBSERVACIONES: Ninguna.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim strSQL As String

If dtpHasta.Value < dtpDesde.Value Then
   MsgBox "Verifique el Rango de Fechas", vbInformation, "Error"
   Exit Sub
End If

Me.MousePointer = vbHourglass

strSQL = "Select " & chkTodos.Value & " as 'valor',L.consec,S.cedula,S.nombre,L.TNeto,L.cod_banco,L.TDocumento" _
       & ",case when L.EstadoActLiq = 'A' then 'Ren.Asociación' when  L.EstadoActLiq = 'P' then 'Ren.Patronal' end as 'Tipo'" _
       & ",isnull(L.cta_ahorros,0) as Cuenta,L.FecLiq,L.usuario,B.Descripcion" _
       & ",dbo.fxTesSupervisa(L.cedula,S.nombre,L.TNeto,0,'L') as 'Duplicado',TES_SUPERVISION_FECHA" _
       & ", isnull(B.Cod_Divisa,'') as 'Cod_Divisa', L.Id_Token" _
       & " from Liquidacion L inner join Socios S on L.cedula = S.cedula" _
       & " left join Tes_Bancos B on L.cod_Banco = B.id_Banco" _
       & " where L.FecLiq between '" & Format(dtpDesde, "yyyy/mm/dd") _
       & " 00:00:00' and '" & Format(dtpHasta, "yyyy/mm/dd") & " 23:59:59' and L.Ubicacion='T'" _
       & " and L.Estado = 'P'"


If fxSIFParametros("15") = "S" Then
   strSQL = strSQL & " and L.Analista_Revision = 'S'"
End If

If Mid(cboAccion.Text, 1, 1) = "D" And cboTipo.Text <> "TODOS" Then
  strSQL = strSQL & " And L.cod_Banco = " & cboTipo.ItemData(cboTipo.ListIndex)
End If

If Mid(cboEstado.Text, 1, 1) = "P" Then
  strSQL = strSQL & " And L.EstadoAsiento='P'"
Else
  strSQL = strSQL & " And L.EstadoAsiento='G'"
End If

If Mid(cboTipoRenuncia.Text, 1, 1) <> "T" Then
        If Mid(cboTipoRenuncia.Text, 1, 1) = "A" Then
          strSQL = strSQL & " And L.ESTADOACTLIQ='A'"
        Else
          strSQL = strSQL & " And L.ESTADOACTLIQ='P'"
        End If
End If



If chkFiltros.Value = vbChecked Then
    If cboBanco.Text <> "TODOS" Then
      strSQL = strSQL & " And L.cod_banco = " & cboBanco.ItemData(cboBanco.ListIndex)
    End If

    If cboOficina.Text <> "TODOS" Then
      strSQL = strSQL & " And L.cod_oficina = '" & cboOficina.ItemData(cboOficina.ListIndex) & "'"
    End If

    If cboUsuarios.Text <> "TODOS" Then
      strSQL = strSQL & " And L.usuario like '" & cboUsuarios.Text & "%'"
    End If

    If cboTokenConsulta.Text <> "TODOS" Then
      strSQL = strSQL & " And isnull(L.ID_Token,'') like '" & cboTokenConsulta.Text & "%'"
    End If

End If


strSQL = strSQL & " ORDER BY L.consec"

Call sbCargaGridLocal(15, strSQL)

If vGrid.MaxRows > 0 Then vGrid.MaxRows = vGrid.MaxRows - 1

Me.MousePointer = vbDefault

End Sub

Private Sub sbTesoreria(Row As Long, vFecha As String, pToken As String)
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'OBJETIVO:      Genera el Asiento de la liquidacion en el modulo de Tesoreria, junto con su
'               detalle y cambia el estado del asiento de la liquidacion a generado.
'REFERENCIAS:   fxFechaServidor - (Devuelve la fecha del servidor)
'OBSERVACIONES: Verifica que el monto al Debe este equilibrado con el monto al Haber.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim strSQL As String, rs As New ADODB.Recordset
Dim curMonto As Currency, vTipoCambio As Currency, vDivisa As String
Dim curDebitos As Currency, curCreditos As Currency
Dim lngSolicitud As Long

Dim oDivisa As String, oTipoCambio As Currency, oMonto As Currency


curDebitos = 0
curCreditos = 0
vTipoCambio = 1



vGrid.Row = Row
vGrid.col = 14
vDivisa = vGrid.Text 'Divisa del Banco


vGrid.col = 2

strSQL = "select COD_DIVISA, isnull(TIPO_CAMBIO,1) as 'TIPO_CAMBIO' FROM LIQUIDACION WHERE CONSEC = " & vGrid.Text
Call OpenRecordSet(rs, strSQL)
 oDivisa = rs!cod_Divisa
 oTipoCambio = rs!Tipo_Cambio
rs.Close

 'Control de Documentos v2
    strSQL = "Select isnull(SUM(MONTO),0) as MONTO From SIF_Transacciones_Asiento Where  Tipo_Documento = 'LIQ' and cod_transaccion ='" & vGrid.Text & "' And Tipo_Movimiento ='D'"
    Call OpenRecordSet(rs, strSQL)
    If Not rs.EOF And Not rs.BOF Then
         curDebitos = rs!Monto
    End If
    rs.Close
    
    strSQL = "Select isnull(SUM(MONTO),0) as MONTO From SIF_Transacciones_Asiento Where  Tipo_Documento = 'LIQ' and cod_transaccion ='" & vGrid.Text & "' And Tipo_Movimiento ='C'"
    Call OpenRecordSet(rs, strSQL)
    If Not rs.EOF And Not rs.BOF Then
         curCreditos = rs!Monto
    End If
    rs.Close



If curCreditos <> curDebitos Then
   MsgBox "La liquidación No." & vGrid.Text & " -> No se Emite a Tesoreria porque se encuentra desbalanceada.!", vbExclamation, "Desbalance"
   Exit Sub
Else
   If curCreditos = 0 Or curDebitos = 0 Then
        MsgBox "La liquidación No." & vGrid.Text & " -> No se Emite a Tesoreria el asiento contable presenta problemas o no existe.!", vbExclamation, "Revisar Asiento!"
        Exit Sub
   End If
End If


'Detalle del Monto a Desembolsar
vGrid.col = 5 'Monto a Liquidar

curMonto = CCur(vGrid.Text)


'TODO: Revisar la Conversion en Multi Divisa

vTipoCambio = oTipoCambio
oMonto = curMonto

curMonto = curMonto / fxSys_Tipo_Cambio_Apl(vTipoCambio)


strSQL = "Insert Tes_Transacciones(ID_Banco,Tipo,Codigo,Beneficiario,Monto,Fecha_Solicitud," _
       & "Estado,EstadoI,Modulo,Cta_Ahorros,"

vGrid.col = 6
strSQL = strSQL & "Detalle1,Detalle2,Detalle3,Detalle4,Detalle5,SubModulo,Actualiza,cod_unidad,cod_concepto,user_solicita" _
               & ",ID_TOKEN ,REMESA_TIPO, REMESA_ID, COD_DIVISA, TIPO_CAMBIO, COD_APP)" _
               & " Values(" & vGrid.Text & ",'"
        
vGrid.col = 7
strSQL = strSQL & vGrid.Text & "','"
        
vGrid.col = 3
strSQL = strSQL & vGrid.Text & "','"
       
vGrid.col = 4
strSQL = strSQL & vGrid.Text & "'," & curMonto & ",'" & Format(vFecha, "yyyy/mm/dd") & "','P','P','CC',"

vGrid.col = 9
strSQL = strSQL & "'" & Trim(vGrid.Text) & "',"
 
 vGrid.col = 2
 strSQL = strSQL & "'LIQ. DE PERSONA-AFILIACION'," _
        & "'#Liq: " & vGrid.Text & "',"
        
 vGrid.col = 8
 strSQL = strSQL & "'Tipo: " & vGrid.Text & "',"
 vGrid.col = 10
 strSQL = strSQL & "'Fecha: " & vGrid.Text & "',"
 vGrid.col = 11
 strSQL = strSQL & "'Usuario: " & vGrid.Text & "','A','S','" & GLOBALES.gOficinaUnidad & "','GEN','" & glogon.Usuario _
                 & "', '" & pToken & "','LIQ', 0, '" & vDivisa & "'," & vTipoCambio & ", 'ProGrX')"
 'Registra
 Call ConectionExecute(strSQL)


vGrid.col = 3
strSQL = "Select Max(NSolicitud) as Solicitud " _
       & " from Tes_Transacciones Where Codigo = '" & vGrid.Text _
       & "' and Fecha_Solicitud = '" & Format(vFecha, "yyyy/mm/dd") & "'"
   
Call OpenRecordSet(rs, strSQL)
If Not rs.EOF And Not rs.BOF Then
   lngSolicitud = rs!solicitud
End If
rs.Close

'-Asiento
curMonto = curMonto * fxSys_Tipo_Cambio_Apl(vTipoCambio)

vGrid.col = 6
strSQL = "Select CTACONTA From Tes_Bancos where id_banco=" & vGrid.Text
Call OpenRecordSet(rs, strSQL)
If Not rs.EOF And Not rs.BOF Then
     strSQL = "Insert Into Tes_Trans_Asiento(NSolicitud,Cuenta_Contable,Monto,DebeHaber,Linea,COD_UNIDAD,Tipo_Cambio,cod_Divisa) Values(" _
            & lngSolicitud & ",'" & rs!ctaConta & "'," & curMonto & ",'H',1,'" & GLOBALES.gOficinaUnidad & "'," & vTipoCambio & ",'" & vDivisa & "')"
     Call ConectionExecute(strSQL)
End If
rs.Close

strSQL = "Select CTA_LIQPAS From Par_AfAH Where cod_Divisa = '" & oDivisa & "'"
Call OpenRecordSet(rs, strSQL)
If Not rs.EOF And Not rs.BOF Then
    strSQL = "Insert Into Tes_Trans_Asiento(NSolicitud,Cuenta_Contable,Monto,DebeHaber,Linea,cod_unidad,Tipo_Cambio,cod_Divisa) Values(" _
           & lngSolicitud & ",'" & Trim(rs!cta_liqpas) & "'," & curMonto & ",'D',2,'" & GLOBALES.gOficinaUnidad & "'," & oTipoCambio & ",'" & oDivisa & "')"
    Call ConectionExecute(strSQL)
End If
rs.Close

vGrid.col = 2
strSQL = "Update Liquidacion set EstadoAsiento='G',Fecha_Traspaso=dbo.MyGetdate(),Traspaso_Usuario = '" _
       & glogon.Usuario & "',ID_TOKEN = '" & pToken & "', Tesoreria_Solicitud = " & lngSolicitud _
       & " Where consec=" & vGrid.Text
Call ConectionExecute(strSQL)


End Sub


Private Sub sbGenerar()
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'OBJETIVO:      Genera a Tesoreria las liquidaciones.
'REFERENCIAS:   AsientoLiquidacionTesoreria - (Genera el Asiento de la liquidacion en el
'               modulo de Tesoreria)
'               ProcedimientoErrores - (Registra error en caso de que ocurra uno dentro del
'               Procedimiento)
'OBSERVACIONES: Ninguna.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim strSQL As String, vRetencion As String
Dim i As Long, vFecha As String
Dim iDuplicado As Integer
Dim vToken As String
Dim rs As New ADODB.Recordset


On Error GoTo vError



If vGrid.MaxRows = 0 Then Exit Sub
If Mid(cboEstado.Text, 1, 1) = "G" Then Exit Sub

Me.MousePointer = vbHourglass

vToken = cboToken.ItemData(cboToken.ListIndex)

If Mid(cboAccion.Text, 1, 1) = "R" Then
  vRetencion = cboTipo.ItemData(cboTipo.ListIndex)
Else
  vRetencion = ""
End If

vFecha = Format(fxFechaServidor, "yyyy/mm/dd")
Prgbar.Max = vGrid.MaxRows


For i = 1 To vGrid.MaxRows
  vGrid.Row = i
  vGrid.col = 13
  iDuplicado = vGrid.Text
  
  vGrid.col = 1
 
  If vGrid.Value = vbChecked And iDuplicado = 0 Then
     
     If Mid(cboAccion.Text, 1, 1) = "D" Then
             Call sbTesoreria(i, vFecha, vToken)
     Else
        'Retener
        vGrid.col = 2
        strSQL = "Update Liquidacion set Fecha_Traspaso= dbo.MyGetdate(),EstadoAsiento = 'G',NDocumento= '0',Tdocumento = 'RT'" _
               & ",Tesoreria_Solicitud = 0, Traspaso_Usuario = '" & glogon.Usuario & "'" _
               & " Where Consec = " & vGrid.Text
        Call ConectionExecute(strSQL)
     
     End If
   
     
  End If
 
 Prgbar.Value = i
 
Next i

Prgbar.Value = 0

Me.MousePointer = vbDefault
MsgBox "Proceso Finalizado...", vbExclamation

Call sbBuscar


Exit Sub
vError:
   MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub dtpDesde_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
   dtpHasta.SetFocus
End If
End Sub
Private Sub sbTokens_Load()
Dim strSQL As String

On Error GoTo vError

strSQL = "exec spTes_Token_Consulta '', 'A','" & glogon.Usuario & "'"
Call sbCbo_Llena_New(cboToken, strSQL, False, True)

Exit Sub

vError:

End Sub

Private Sub Form_Load()
Dim strSQL As String

vModulo = 1 'Afiliación

dtpDesde.Value = fxFechaServidor
dtpHasta.Value = dtpDesde.Value

Set imgBanner.Picture = frmContenedor.imgBanner_01.Picture

vGrid.AppearanceStyle = fxGridStyle

vPaso = True

    cboAccion.Clear
    cboAccion.AddItem "Desembolsar"
    cboAccion.AddItem "Retener"
    cboAccion.Text = "Desembolsar"
    
    cboEstado.Clear
    cboEstado.AddItem "Pendiente"
    cboEstado.AddItem "Generado"
    cboEstado.Text = "Pendiente"
    
    cboTipoRenuncia.Clear
    cboTipoRenuncia.AddItem "Todas"
    cboTipoRenuncia.AddItem "Asociación"
    cboTipoRenuncia.AddItem "Patronal"
    cboTipoRenuncia.Text = "Todas"

vPaso = False

Call cboAccion_Click

Call sbTokens_Load

Call Formularios(Me)
Call RefrescaTags(Me)

End Sub


Private Sub Form_Resize()
On Error Resume Next

imgBanner.Width = Me.Width
vGrid.Width = Me.Width - 150
vGrid.Height = Me.Height - (1680 + 750)

Prgbar.Width = Me.Width
Prgbar.top = vGrid.top + vGrid.Height + 20

End Sub

Private Sub tlbAcciones_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
  Case "Buscar"
    Call sbBuscar
  Case "Generar"
    Call sbGenerar
End Select
End Sub

Private Sub sbInformes(pInforme As String)
Dim strSQL As String

Me.MousePointer = vbHourglass

With frmContenedor.Crt
 .Reset
 .WindowShowGroupTree = True
 .WindowShowPrintSetupBtn = True
 .WindowShowRefreshBtn = True
 .WindowShowSearchBtn = True
 .WindowState = crptMaximized
 .WindowTitle = "Reportes del Módulo de Personas"
 .Formulas(0) = "Fecha='" & Format(fxFechaServidor, "dd/mm/yyyy") & "'"
 .Formulas(1) = "Empresa='" & GLOBALES.gstrNombreEmpresa & "'"

 .Connect = glogon.ConectRPT

    Select Case pInforme
      Case "Resumen"
        .ReportFileName = SIFGlobal.fxPathReportes("Personas_LiquidacionTesoreriaResumen.rpt")
        strSQL = "{LIQUIDACION.ESTADO} = 'P' AND {LIQUIDACION.UBICACION} = 'T' AND {LIQUIDACION.FECHA_TRASPASO} in Date(" & Format(dtpDesde.Value, "yyyy,mm,dd") _
               & ") to Date(" & Format(dtpHasta.Value, "yyyy,mm,dd") & ")"
      Case "Detalle"
        .ReportFileName = SIFGlobal.fxPathReportes("Personas_LiquidacionTesoreriaDetalle.rpt")
        
        
        
        strSQL = "{LIQUIDACION.ESTADO} = 'P' AND {LIQUIDACION.UBICACION} = 'T' AND {LIQUIDACION.FECHA_TRASPASO} in Date(" & Format(dtpDesde.Value, "yyyy,mm,dd") _
               & ") to Date(" & Format(dtpHasta.Value, "yyyy,mm,dd") & ")"
            
      Case "Pendientes"
        .ReportFileName = SIFGlobal.fxPathReportes("Personas_LiquidacionTesoreriaPendientes.rpt")
        strSQL = "{LIQUIDACION.ESTADO} = 'P' AND {LIQUIDACION.UBICACION} = 'T' AND {LIQUIDACION.FECLIQ} in Date(" & Format(dtpDesde.Value, "yyyy,mm,dd") _
               & ") to Date(" & Format(dtpHasta.Value, "yyyy,mm,dd") & ") and {LIQUIDACION.ESTADOASIENTO} = 'P'"
 End Select

    If Mid(cboTipoRenuncia, 1, 1) <> "T" Then
        strSQL = strSQL & " AND {LIQUIDACION.ESTADOACTLIQ} = '" & Mid(cboTipoRenuncia, 1, 1) & "'"
    End If

  .Formulas(2) = "SubTitulo='Del  " & Format(dtpDesde, "dd/mm/yyyy") & "  Al  " & Format(dtpHasta, "dd/mm/yyyy") & "'"
  
  
  
    If cboBanco.Text <> "TODOS" Then
      strSQL = strSQL & " And {LIQUIDACION.COD_BANCO} = " & cboBanco.ItemData(cboBanco.ListIndex)
    End If

    If cboOficina.Text <> "TODOS" Then
      strSQL = strSQL & " And {LIQUIDACION.COD_OFICINA} = '" & cboOficina.ItemData(cboOficina.ListIndex) & "'"
    End If

    If cboUsuarios.Text <> "TODOS" Then
      strSQL = strSQL & " And {LIQUIDACION.USUARIO} = '" & cboUsuarios.Text & "'"
    End If

 
    If cboTokenConsulta.Text <> "TODOS" And cboTokenConsulta.Text <> "" Then
      strSQL = strSQL & " And {LIQUIDACION.ID_TOKEN} = '" & cboTokenConsulta.Text & "'"
    End If
  
  
  .SelectionFormula = strSQL
  .PrintReport

End With

Me.MousePointer = vbDefault

End Sub

Private Sub sbCargaGridLocal(vGridMaxCol As Integer, strSQL As String)
Dim rs As New ADODB.Recordset, i As Integer, strResultado As String
Dim strLista As String

Me.MousePointer = vbHourglass

vGrid.MaxCols = vGridMaxCol
vGrid.MaxRows = 1

vGrid.Row = vGrid.MaxRows
vDuplicado = False
strLista = ""

Call OpenRecordSet(rs, strSQL, 0)
 

Do While Not rs.EOF
  vGrid.Row = vGrid.MaxRows
  
  For i = 1 To vGrid.MaxCols
    vGrid.col = i
    Select Case i
     Case 1
        vGrid.Text = CStr(rs!Valor)
     Case 2
        vGrid.Text = CStr(rs!consec)
     Case 3
        vGrid.Text = CStr(rs!Cedula)
     Case 4
        vGrid.Text = CStr(rs!Nombre)
     Case 5
        vGrid.Text = Format(rs!TNETO, "Standard")
     Case 6
        vGrid.Text = CStr(rs!cod_banco)
     Case 7
            vGrid.Text = CStr(rs!TDOCUMENTO)
     Case 8
        vGrid.Text = CStr(rs!Tipo)
     Case 9
        vGrid.Value = CStr(rs!Cuenta)
     Case 10
        vGrid.Value = CStr(Format(rs!fecLiq, "dd/mm/yyyy"))
     
     Case 11
        vGrid.Value = CStr(rs!Usuario)
    Case 12
        vGrid.Value = CStr(rs!Descripcion)
    Case 13
         If rs!Duplicado = 1 And IsNull(rs!TES_SUPERVISION_FECHA) Then
             vDuplicado = True
             vGrid.col = 2
             vGrid.Row = vGrid.ActiveRow
             vGrid.BackColor = vbRed
             strLista = strLista & "Cédula " & rs!Cedula & " " & rs!Nombre & " " & "LIQ. " & rs!consec & vbCrLf
         Else
            vGrid.ForeColor = vbBlack
        End If
       vGrid.Value = IIf(vDuplicado = True, CStr(rs!Duplicado), CStr(0))
    
     Case 14
        vGrid.Text = Trim(rs!cod_Divisa)
    
     Case 15
        vGrid.Text = Trim(rs!id_token & "")
    
 
    End Select
  
  Next i
  
  vGrid.MaxRows = vGrid.MaxRows + 1
  
  rs.MoveNext
Loop

If vDuplicado Then
      MsgBox "Estas liquidaciones necesitan autorización para ser trasladadas ya que cuentan" _
          & "con una transacción por un monto igual en Tesorería " & vbCrLf & vbCrLf & strLista, vbCritical
End If
rs.Close

Me.MousePointer = vbDefault
End Sub


