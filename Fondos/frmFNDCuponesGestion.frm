VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "Codejock.Controls.v22.1.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "Codejock.ShortcutBar.v22.1.0.ocx"
Begin VB.Form frmFNDCuponesGestion 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Fondos: Gestión de Cupones"
   ClientHeight    =   9780
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14220
   Icon            =   "frmFNDCuponesGestion.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9780
   ScaleWidth      =   14220
   WindowState     =   2  'Maximized
   Begin XtremeSuiteControls.PushButton btnAccion 
      Height          =   495
      Index           =   0
      Left            =   7440
      TabIndex        =   20
      Top             =   1680
      Width           =   1215
      _Version        =   1441793
      _ExtentX        =   2143
      _ExtentY        =   873
      _StockProps     =   79
      Caption         =   "Buscar"
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
      Picture         =   "frmFNDCuponesGestion.frx":000C
      ImageAlignment  =   4
   End
   Begin XtremeSuiteControls.CheckBox chkFechas 
      Height          =   252
      Left            =   6840
      TabIndex        =   16
      Top             =   960
      Width           =   1692
      _Version        =   1441793
      _ExtentX        =   2984
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Todas"
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
   Begin MSComCtl2.FlatScrollBar FlatScrollBar 
      Height          =   252
      Left            =   8880
      TabIndex        =   0
      Top             =   480
      Width           =   492
      _ExtentX        =   873
      _ExtentY        =   450
      _Version        =   393216
      Arrows          =   65536
      Orientation     =   1638401
   End
   Begin FPSpreadADO.fpSpread vGrid 
      Height          =   3015
      Left            =   0
      TabIndex        =   4
      Top             =   2880
      Width           =   14295
      _Version        =   524288
      _ExtentX        =   25215
      _ExtentY        =   5318
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
      MaxCols         =   17
      SpreadDesigner  =   "frmFNDCuponesGestion.frx":070C
      VScrollSpecial  =   -1  'True
      VScrollSpecialType=   2
      AppearanceStyle =   1
   End
   Begin XtremeSuiteControls.ComboBox cboOperadora 
      Height          =   312
      Left            =   2640
      TabIndex        =   7
      Top             =   120
      Width           =   6132
      _Version        =   1441793
      _ExtentX        =   10821
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
   Begin XtremeSuiteControls.FlatEdit txtCodigo 
      Height          =   312
      Left            =   2640
      TabIndex        =   8
      Top             =   480
      Width           =   1212
      _Version        =   1441793
      _ExtentX        =   2138
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
   Begin XtremeSuiteControls.FlatEdit txtDescripcion 
      Height          =   312
      Left            =   3840
      TabIndex        =   9
      Top             =   480
      Width           =   4932
      _Version        =   1441793
      _ExtentX        =   8700
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
   Begin XtremeSuiteControls.ComboBox cboRetencion 
      Height          =   315
      Left            =   960
      TabIndex        =   10
      Top             =   1680
      Width           =   6135
      _Version        =   1441793
      _ExtentX        =   10821
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
   Begin XtremeSuiteControls.ComboBox cboBanco 
      Height          =   312
      Left            =   3480
      TabIndex        =   11
      Top             =   1680
      Width           =   3612
      _Version        =   1441793
      _ExtentX        =   6376
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
   Begin XtremeSuiteControls.ComboBox cboProceso 
      Height          =   312
      Left            =   960
      TabIndex        =   12
      Top             =   1320
      Width           =   1452
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
   Begin XtremeSuiteControls.ComboBox cboTipoDocumento 
      Height          =   312
      Left            =   960
      TabIndex        =   13
      Top             =   1680
      Width           =   1452
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
   Begin XtremeSuiteControls.DateTimePicker dtpInicio 
      Height          =   312
      Left            =   3840
      TabIndex        =   14
      Top             =   960
      Width           =   1452
      _Version        =   1441793
      _ExtentX        =   2561
      _ExtentY        =   550
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
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   3
   End
   Begin XtremeSuiteControls.DateTimePicker dtpCorte 
      Height          =   312
      Left            =   5280
      TabIndex        =   15
      Top             =   960
      Width           =   1452
      _Version        =   1441793
      _ExtentX        =   2561
      _ExtentY        =   550
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
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   3
   End
   Begin XtremeSuiteControls.CheckBox chkTesoreria 
      Height          =   252
      Left            =   2880
      TabIndex        =   17
      Top             =   2040
      Width           =   4212
      _Version        =   1441793
      _ExtentX        =   7429
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "&Enviar a Tesorería Automáticamente"
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
   Begin XtremeSuiteControls.CheckBox chkMarcas 
      Height          =   200
      Left            =   600
      TabIndex        =   18
      Top             =   2460
      Width           =   200
      _Version        =   1441793
      _ExtentX        =   353
      _ExtentY        =   353
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
      Transparent     =   -1  'True
      UseVisualStyle  =   -1  'True
      Appearance      =   16
      Alignment       =   1
   End
   Begin XtremeSuiteControls.PushButton btnAccion 
      Height          =   495
      Index           =   1
      Left            =   8640
      TabIndex        =   21
      Top             =   1680
      Width           =   1215
      _Version        =   1441793
      _ExtentX        =   2143
      _ExtentY        =   873
      _StockProps     =   79
      Caption         =   "Exportar"
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
      Picture         =   "frmFNDCuponesGestion.frx":12E9
      ImageAlignment  =   4
   End
   Begin XtremeSuiteControls.GroupBox gbResumen 
      Height          =   1575
      Left            =   0
      TabIndex        =   23
      Top             =   6000
      Width           =   11295
      _Version        =   1441793
      _ExtentX        =   19923
      _ExtentY        =   2778
      _StockProps     =   79
      Caption         =   "Resumen"
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
      Begin XtremeSuiteControls.PushButton cmdLiquidar 
         Height          =   615
         Left            =   9720
         TabIndex        =   24
         Top             =   480
         Width           =   1335
         _Version        =   1441793
         _ExtentX        =   2355
         _ExtentY        =   1085
         _StockProps     =   79
         Caption         =   "Procesar"
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
         Picture         =   "frmFNDCuponesGestion.frx":1BBA
         ImageAlignment  =   4
      End
      Begin XtremeSuiteControls.FlatEdit txtCasos 
         Height          =   312
         Left            =   1920
         TabIndex        =   25
         Top             =   480
         Width           =   1572
         _Version        =   1441793
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
      Begin XtremeSuiteControls.FlatEdit txtAportes 
         Height          =   312
         Left            =   1920
         TabIndex        =   26
         Top             =   840
         Width           =   1572
         _Version        =   1441793
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
      Begin XtremeSuiteControls.FlatEdit txtRendimientos 
         Height          =   312
         Left            =   1920
         TabIndex        =   27
         Top             =   1200
         Width           =   1572
         _Version        =   1441793
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
      Begin XtremeSuiteControls.FlatEdit txtTotal 
         Height          =   312
         Left            =   4800
         TabIndex        =   28
         Top             =   480
         Width           =   2052
         _Version        =   1441793
         _ExtentX        =   3619
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
      Begin XtremeSuiteControls.FlatEdit txtISR 
         Height          =   312
         Left            =   4800
         TabIndex        =   29
         ToolTipText     =   "Monto en Multa para aplicación masiva por persona"
         Top             =   840
         Width           =   2052
         _Version        =   1441793
         _ExtentX        =   3619
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
      Begin XtremeSuiteControls.FlatEdit txtNeto 
         Height          =   315
         Left            =   4800
         TabIndex        =   37
         Top             =   1200
         Width           =   2055
         _Version        =   1441793
         _ExtentX        =   3619
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
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Neto"
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
         Height          =   255
         Index           =   4
         Left            =   4080
         TabIndex        =   22
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "(-) ISR"
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
         Index           =   5
         Left            =   4080
         TabIndex        =   35
         Top             =   840
         Width           =   972
      End
      Begin VB.Image imgCalcula 
         Height          =   480
         Left            =   6960
         Picture         =   "frmFNDCuponesGestion.frx":22E1
         ToolTipText     =   "Calcular Ejecución"
         Top             =   480
         Width           =   480
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Total"
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
         Index           =   3
         Left            =   4080
         TabIndex        =   34
         Top             =   480
         Width           =   972
      End
      Begin VB.Label lbl 
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
         Height          =   255
         Left            =   7560
         TabIndex        =   33
         Top             =   1200
         Width           =   6375
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Rendimientos"
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
         Index           =   2
         Left            =   600
         TabIndex        =   32
         Top             =   1200
         Width           =   1212
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Principal"
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
         Height          =   255
         Index           =   1
         Left            =   600
         TabIndex        =   31
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Casos"
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
         Index           =   0
         Left            =   600
         TabIndex        =   30
         Top             =   480
         Width           =   1092
      End
   End
   Begin MSComctlLib.ProgressBar prgBar 
      Align           =   2  'Align Bottom
      Height          =   135
      Left            =   0
      TabIndex        =   36
      Top             =   9645
      Visible         =   0   'False
      Width           =   14220
      _ExtentX        =   25083
      _ExtentY        =   238
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin XtremeShortcutBar.ShortcutCaption lblTitulo 
      Height          =   372
      Left            =   0
      TabIndex        =   19
      Top             =   2400
      Width           =   11292
      _Version        =   1441793
      _ExtentX        =   19918
      _ExtentY        =   656
      _StockProps     =   14
      Caption         =   "Lista de Cupones pendientes de liquidación"
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
      Alignment       =   1
   End
   Begin VB.Label lblBanco 
      BackColor       =   &H00FFFFFF&
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
      Height          =   330
      Left            =   2520
      TabIndex        =   5
      Top             =   1680
      Width           =   975
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Cupones que vencen entre...:"
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
      Left            =   960
      TabIndex        =   3
      Top             =   960
      Width           =   2772
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
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
      ForeColor       =   &H00FFFFFF&
      Height          =   312
      Left            =   1560
      TabIndex        =   2
      Top             =   480
      Width           =   1212
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
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
      ForeColor       =   &H00FFFFFF&
      Height          =   312
      Index           =   0
      Left            =   1560
      TabIndex        =   1
      Top             =   120
      Width           =   1212
   End
   Begin VB.Label lblConcepto 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Concepto"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5400
      TabIndex        =   6
      Top             =   1440
      Width           =   1695
   End
   Begin VB.Image imgBanner 
      Height          =   852
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   15732
   End
End
Attribute VB_Name = "frmFNDCuponesGestion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSQL As String, rs As New ADODB.Recordset, itmX As ListItem
Dim vScroll As Boolean, vPaso As Boolean
Dim vCuentaRet As String, vCuentaRetencion As String

Private Sub sbConsultaPlan(vPlan As String)


On Error GoTo vError

Me.MousePointer = vbHourglass

Dim pTipoPago As String, pBancoId As Long

pTipoPago = fxTipoDocumento(cboTipoDocumento.Text)
pBancoId = cboBanco.ItemData(cboBanco.ListIndex)


If Mid(cboProceso.Text, 1, 1) = "D" And cboBanco.Text = "TODOS" Then
   pBancoId = 0
End If


Select Case pTipoPago
    Case "RC", "FD"
        pBancoId = 0
End Select

If Mid(cboProceso.Text, 1, 1) = "R" Then
  pBancoId = 0
  pTipoPago = "RT"
End If



If chkFechas.Value = vbUnchecked Then
    strSQL = "exec spFndCDPCuponesConsultaVencimiento '" & Format(dtpInicio.Value, "yyyy/mm/dd") & " 00:00:00','" _
           & Format(dtpCorte.Value, "yyyy/mm/dd") & " 23:59:59'," & cboOperadora.ItemData(cboOperadora.ListIndex)
Else
    strSQL = "exec spFndCDPCuponesConsultaVencimiento '1900/01/01 00:00:00','2199/12/31 23:59:59'," & cboOperadora.ItemData(cboOperadora.ListIndex)
End If

If txtCodigo.Text <> "" Then
  strSQL = strSQL & ",'" & txtCodigo.Text & "'"
Else
  strSQL = strSQL & ",''"
End If

strSQL = strSQL & ",'" & Mid(cboProceso.Text, 1, 1) & "','" & pTipoPago & "'," & pBancoId


vPaso = True

vGrid.MaxRows = 0


Call OpenRecordSet(rs, strSQL)

Do While Not rs.EOF
 vGrid.MaxRows = vGrid.MaxRows + 1
 vGrid.Row = vGrid.MaxRows
 
 vGrid.col = 1
 vGrid.Value = chkMarcas.Value
 vGrid.col = 2
 vGrid.Text = rs!Cupon_Id
 
 vGrid.col = 3
 vGrid.Text = Format(rs!Fecha_Vence, "dd/mm/yyyy")
 vGrid.col = 4
 vGrid.Text = rs!DiasTransc
 
 vGrid.col = 5
 vGrid.Text = rs!cod_Plan
 vGrid.col = 6
 vGrid.Text = rs!COD_CONTRATO
 
 vGrid.col = 7
 vGrid.Text = rs!Cedula
 vGrid.col = 8
 vGrid.Text = rs!Nombre
 vGrid.col = 9
 vGrid.Text = Format(rs!Principal, "standard")
 vGrid.col = 10
 vGrid.Text = Format(rs!Rendimiento, "standard")
 vGrid.col = 11
 vGrid.Text = Format(rs!Total, "standard")

 vGrid.col = 12
 vGrid.Text = Format(rs!ISR, "standard")

 vGrid.col = 13
 vGrid.Text = Format(rs!Neto, "standard")


 vGrid.col = 14
 vGrid.Text = rs!ID_BANCO
 vGrid.col = 15
 vGrid.Text = rs!TIPO_PAGO

 vGrid.col = 16
 vGrid.Text = rs!Cuenta_Ahorros & ""
 vGrid.col = 17
 vGrid.Text = rs!BancoDesc


 rs.MoveNext
Loop
rs.Close

vPaso = False

'If vGrid.MaxRows > 2 Then
' vGrid.MaxRows = vGrid.MaxRows - 1
'End If
Me.MousePointer = vbDefault

Call imgCalcula_Click

Exit Sub

vError:
    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub



Private Sub sbExportar()

Dim vHeaders As vGridHeaders
    vHeaders.Columnas = 17
    vHeaders.Headers(1) = "Check"
    vHeaders.Headers(2) = "Cupón Id"
    vHeaders.Headers(3) = "Fecha Vence"
    vHeaders.Headers(4) = "Dias al Vencimiento"
    vHeaders.Headers(5) = "Plan"
    vHeaders.Headers(6) = "No. Contrato"
    vHeaders.Headers(7) = "Identificación"
    vHeaders.Headers(8) = "Nombre"
    
    vHeaders.Headers(9) = "Cp.Principal"
    vHeaders.Headers(10) = "Cp.Rendimiento"
    vHeaders.Headers(11) = "Cp.Total"
    vHeaders.Headers(12) = "Cp.ISR"
    vHeaders.Headers(13) = "Cp.Mnt.Pagar"
    
    vHeaders.Headers(14) = "Cta.Id."
    vHeaders.Headers(15) = "Tipo Pago"
    vHeaders.Headers(16) = "Cta.Persona"
    vHeaders.Headers(17) = "Banco Desc."

 Call sbSIFGridExportar(vGrid, vHeaders, "ProGrX_Cupones_Liquidacion_" & Trim(txtCodigo.Text))

End Sub


Private Sub btnAccion_Click(Index As Integer)

Select Case Index
  Case 0  'Buscar
        Call sbConsultaPlan(Trim(txtCodigo.Text))
  Case 1  'Exportar
        Call sbExportar
End Select

End Sub

Private Sub cboBanco_Click()
vGrid.MaxRows = 0
Call imgCalcula_Click
End Sub

Private Sub cboOperadora_Click()
If vPaso Then Exit Sub

Call txtCodigo_LostFocus
If Trim(txtCodigo) <> "" Then Call sbConsultaPlan(Trim(txtCodigo))
End Sub


Private Sub cboOperadora_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtCodigo.SetFocus
End Sub


Private Function fxCuentaPlan(pTipo As String) As String
Dim strSQL As String, rs As New ADODB.Recordset

fxCuentaPlan = ""

If pTipo = "P" Then
 'Cuenta del Plan : Aportes
  strSQL = "Select Cuenta_Conta as CuentaX from Fnd_Planes Where Cod_Operadora=" & cboOperadora.ItemData(cboOperadora.ListIndex) _
         & " and Cod_Plan='" & txtCodigo & "'"
Else
 'Cuenta del Plan : Rendimiento
  strSQL = "Select Cuenta_Rendimiento as CuentaX from Fnd_Planes Where Cod_Operadora=" & cboOperadora.ItemData(cboOperadora.ListIndex) _
         & " and Cod_Plan='" & txtCodigo & "'"

End If

Call OpenRecordSet(rs, strSQL)
If Not rs.EOF And Not rs.BOF Then
    fxCuentaPlan = Trim(rs!CuentaX)
End If
rs.Close

End Function


Private Sub cboProceso_Click()

If vPaso Then Exit Sub

vGrid.MaxRows = 0
Call imgCalcula_Click

lblBanco.Visible = False
cboBanco.Visible = False
chkTesoreria.Visible = False

cboTipoDocumento.Visible = False
lblConcepto.Visible = False
cboRetencion.Visible = False


If Mid(cboProceso.Text, 1, 1) = "D" Then
    cboTipoDocumento.Visible = True
    
    Select Case fxTipoDocumento(cboTipoDocumento.Text)
        Case "TE", "CK"
            chkTesoreria.Visible = True
            lblBanco.Visible = True
            cboBanco.Visible = True
        Case "RC", "FD"
            chkTesoreria.Visible = False
            lblBanco.Visible = False
            cboBanco.Visible = False
    End Select
    
Else
    lblConcepto.Visible = True
    cboRetencion.Visible = True
End If

End Sub

Private Sub cboRetencion_Click()

If vPaso Then Exit Sub
If cboRetencion.ListCount <= 0 Then Exit Sub

Dim strSQL As String, rs As New ADODB.Recordset

strSQL = "select cod_cuenta from FND_RETENCION_CONCEPTOS where Retencion_Codigo = '" & cboRetencion.ItemData(cboRetencion.ListIndex) & "'"
Call OpenRecordSet(rs, strSQL)
  vCuentaRetencion = Trim(rs!cod_cuenta)
rs.Close

End Sub

Private Sub cboTipoDocumento_Click()
vGrid.MaxRows = 0
Call imgCalcula_Click

Select Case fxTipoDocumento(cboTipoDocumento.Text)
    Case "TE", "CK"
        chkTesoreria.Visible = True
        lblBanco.Visible = True
        cboBanco.Visible = True
    Case "RC", "FD"
        chkTesoreria.Visible = False
        lblBanco.Visible = False
        cboBanco.Visible = False
End Select
    
End Sub

Private Sub chkFechas_Click()

If chkFechas.Value = vbChecked Then
   dtpInicio.Enabled = False
Else
   dtpInicio.Enabled = True
End If

dtpCorte.Enabled = dtpInicio.Enabled

End Sub


Private Sub chkMarcas_Click()
Dim lng As Long

For lng = 1 To vGrid.MaxRows
 vGrid.Row = lng
 vGrid.col = 1
 vGrid.Value = chkMarcas.Value
Next lng


Call imgCalcula_Click

End Sub


Private Sub sbLiquidar()
Dim lng As Long, vFecha As Date, vCuenta As String
Dim vConcepto As String, vTipoDoc As String, vDocRef As String


strSQL = MsgBox("Confirma la liquidación de Cupones seleccionados?", vbExclamation + vbYesNo)
If strSQL = vbNo Then Exit Sub

On Error GoTo vError

Me.MousePointer = vbHourglass

lbl.Caption = "Procesando..."
lbl.Refresh

PrgBar.Max = vGrid.MaxRows + 1
PrgBar.Value = 1

PrgBar.Visible = True

Dim vCuponId As Integer
Dim vTipoPago As String, vBancoId As Long


vTipoPago = fxTipoDocumento(cboTipoDocumento.Text)
vBancoId = cboBanco.ItemData(cboBanco.ListIndex)

For lng = 1 To vGrid.MaxRows
 vGrid.Row = lng
 vGrid.col = 1
 If vGrid.Value = vbChecked Then
    lbl.Caption = "Procesando Contrato : " & vGrid.Text
    lbl.Refresh
    
    vGrid.col = 5 'Plan
    strSQL = "exec spFndCDPCuponesLiquida " & cboOperadora.ItemData(cboOperadora.ListIndex) & ",'" & vGrid.Text & "',"
    vGrid.col = 6 'Contrato
    strSQL = strSQL & vGrid.Text & ","
    vGrid.col = 2 'CuponId
    strSQL = strSQL & vGrid.Text & ",'" & glogon.Usuario & "','" & Mid(cboProceso.Text, 1, 1) & "'"
    vCuponId = vGrid.Text
    
    If Mid(cboProceso.Text, 1, 1) = "R" Then
        strSQL = strSQL & ",'" & cboRetencion.ItemData(cboRetencion.ListIndex) & "','OT',0,'',0,'Liquidación de Cupón: " & vCuponId & "','" & glogon.AppName & "'"
    Else
        vGrid.col = 16
        strSQL = strSQL & ",'','" & vTipoPago _
               & "'," & vBancoId & ",'" & vGrid.Text & "'," & chkTesoreria.Value & ",'Liquidación de Cupón: " & vCuponId & "','" & glogon.AppName & "'"
    
    End If
    
    
    
    Call ConectionExecute(strSQL)
    
 End If

 PrgBar.Value = PrgBar.Value + 1

Next lng


PrgBar.Visible = False
lbl.Caption = ""

Me.MousePointer = vbDefault
MsgBox "Cupones Liquidados Satisfactoriamente!", vbInformation

Call sbConsultaPlan(txtCodigo.Text)

Exit Sub


vError:
    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub cmdLiquidar_Click()
Call sbLiquidar
End Sub

Private Sub FlatScrollBar_Change()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

If vScroll Then
    strSQL = "select Top 1 cod_plan,descripcion from fnd_planes" _
           & " where cod_operadora = " & cboOperadora.ItemData(cboOperadora.ListIndex) & " and Tipo_CDP = 1"

    If FlatScrollBar.Value = 1 Then
       strSQL = strSQL & " and cod_plan > '" & txtCodigo & "' order by cod_plan asc"
    Else
       strSQL = strSQL & " and cod_plan < '" & txtCodigo & "' order by cod_plan desc"
    End If

    Call OpenRecordSet(rs, strSQL)
    If Not rs.EOF And Not rs.BOF Then
      txtCodigo.Text = rs!cod_Plan
      txtDescripcion.Text = rs!Descripcion
    End If
    rs.Close
End If

vScroll = False
FlatScrollBar.Value = 0
vScroll = True

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub Form_Activate()
vModulo = 18 'Fondo de Inversion

End Sub





Private Sub Form_Load()

vModulo = 18 'Fondo de Inversion

vPaso = True


Set imgBanner.Picture = frmContenedor.imgBanner_01.Picture

vGrid.MaxRows = 0
vGrid.MaxCols = 17
vGrid.AppearanceStyle = fxGridStyle

dtpInicio.Value = fxFechaServidor
dtpCorte.Value = dtpInicio.Value
chkFechas.Value = vbChecked

cboProceso.Clear
cboProceso.AddItem "Desembolsar"
cboProceso.AddItem "Retener"


cboTipoDocumento.Clear
cboTipoDocumento.AddItem fxTipoDocumento("TE")
cboTipoDocumento.AddItem fxTipoDocumento("CK")
cboTipoDocumento.AddItem fxTipoDocumento("RC")


If fxFnd_Fondos_Transitorios Then
    cboTipoDocumento.AddItem fxTipoDocumento("FD")
End If

strSQL = "select rtrim(descripcion) as 'ItmX',cod_operadora as 'Idx' from FND_Operadoras"
Call sbCbo_Llena_New(cboOperadora, strSQL, False, True)

strSQL = "select rtrim(RETENCION_CODIGO) as  'IdX', RTRIM(DESCRIPCION) as 'ItmX'" _
       & " From FND_RETENCION_CONCEPTOS  Where ACTIVO = 1"
Call sbCbo_Llena_New(cboRetencion, strSQL, False, True)

strSQL = "select B.id_banco as 'Idx',B.descripcion as 'ItmX'" _
       & " from tes_banco_asg T inner join Tes_Bancos B on T.id_banco = B.id_banco" _
       & " where T.nombre = '" & glogon.Usuario & "' and B.Estado = 'A'"
Call sbCbo_Llena_New(cboBanco, strSQL, False, True)

vPaso = False

cboProceso.Text = "Desembolsar"
cboTipoDocumento.Text = fxTipoDocumento("TE")


Call chkFechas_Click
Call cboOperadora_Click
Call cboRetencion_Click


vScroll = False
     FlatScrollBar.Value = 0
vScroll = True

Call Formularios(Me)
Call RefrescaTags(Me)

End Sub


Private Sub Form_Resize()
'On Error Resume Next
'
'vGrid.Width = Me.Width - 500
'vGrid.Height = Me.Height - (vGrid.top + Frame1.Height + 800)
'Frame1.top = vGrid.top + vGrid.Height + 50
'
'
'imgBanner.Width = Me.Width


On Error Resume Next
 
imgBanner.Width = Me.Width

vGrid.Width = Me.Width - 250
vGrid.Height = Me.Height - (vGrid.top + gbResumen.Height + 800)

lblTitulo.Width = vGrid.Width

gbResumen.top = vGrid.top + vGrid.Height + 50
gbResumen.Width = vGrid.Width

End Sub

Private Sub imgCalcula_Click()
Dim lng As Long, lngCasos As Long
Dim curAportes As Currency, curRendi As Currency
Dim curISR As Currency

Me.MousePointer = vbHourglass

curAportes = 0
curRendi = 0
curISR = 0
lngCasos = 0

For lng = 1 To vGrid.MaxRows
 vGrid.Row = lng
 vGrid.col = 1
 If vGrid.Value = vbChecked Then
   lngCasos = lngCasos + 1
   vGrid.col = 9
   curAportes = curAportes + CCur(vGrid.Text)
   vGrid.col = 10
   curRendi = curRendi + CCur(vGrid.Text)
  
   vGrid.col = 12
   curISR = curISR + CCur(vGrid.Text)

  End If
Next lng

txtCasos.Text = Format(lngCasos, "###,###,###,##0")
txtAportes.Text = Format(curAportes, "Standard")
txtRendimientos.Text = Format(curRendi, "Standard")
txtTotal.Text = Format(curAportes + curRendi, "Standard")
txtISR.Text = Format(curISR, "Standard")
txtNeto.Text = Format(curAportes + curRendi - curISR, "Standard")

Me.MousePointer = vbDefault

End Sub



Private Sub txtCodigo_Change()

vGrid.MaxRows = 0
Call imgCalcula_Click

End Sub



Private Sub txtCodigo_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtDescripcion.SetFocus

If KeyCode = vbKeyF4 Then
   gBusquedas.Columna = "cod_plan"
   gBusquedas.Orden = "cod_plan"
   gBusquedas.Filtro = "And Cod_operadora=" & cboOperadora.ItemData(cboOperadora.ListIndex) & " and Tipo_CDP = 1"
   gBusquedas.Consulta = "select cod_plan,descripcion from fnd_planes"
   frmBusquedas.Show vbModal
   txtDescripcion.SetFocus

   If Trim(gBusquedas.Resultado) <> "" Then
      txtCodigo = Trim(gBusquedas.Resultado)
      txtDescripcion = Trim(gBusquedas.Resultado2)
   End If
   gBusquedas.Resultado = ""
   gBusquedas.Resultado2 = ""
End If

End Sub


Private Sub txtCodigo_LostFocus()
Dim strSQL As String, rs As New ADODB.Recordset

If Trim(txtCodigo) <> "" Then
   strSQL = "Select Descripcion from fnd_planes where cod_operadora=" & cboOperadora.ItemData(cboOperadora.ListIndex)
   strSQL = strSQL & " And cod_plan = '" & Trim(txtCodigo) & "' and Tipo_CDP = 1"
   With rs
     .Open strSQL, glogon.Conection, adOpenStatic
        If .EOF = False Then
           txtDescripcion = Trim(!Descripcion)
        Else
           MsgBox "Codigo incorrecto", vbExclamation
           txtCodigo = ""
           txtDescripcion = ""
           txtCodigo.SetFocus
        End If
     .Close
   End With
Else
  txtDescripcion = ""
End If
End Sub


Private Sub txtDescripcion_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then
   gBusquedas.Columna = "descripcion"
   gBusquedas.Orden = "descripcion"
   gBusquedas.Filtro = "And Cod_operadora=" & cboOperadora.ItemData(cboOperadora.ListIndex) & " and Tipo_CDP = 1"
   gBusquedas.Consulta = "select cod_plan,descripcion from fnd_planes"
   frmBusquedas.Show vbModal
   txtDescripcion.SetFocus
   If Trim(gBusquedas.Resultado) <> "" Then
      txtCodigo = Trim(gBusquedas.Resultado)
      txtDescripcion = Trim(gBusquedas.Resultado2)
   End If
   gBusquedas.Resultado = ""
   gBusquedas.Resultado2 = ""
End If


End Sub


Private Sub vGrid_ButtonClicked(ByVal col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
Dim lng As Long, lngCasos As Long
Dim curAportes As Currency, curRendi As Currency, curISR As Currency

If vPaso Then Exit Sub

If col = 1 Then
   vGrid.Row = Row
   vGrid.col = 1
   If vGrid.Value = vbChecked Then
        lngCasos = 1
        vGrid.col = 9
        curAportes = CCur(vGrid.Text)
        vGrid.col = 10
        curRendi = CCur(vGrid.Text)
        
        vGrid.col = 12
        curISR = CCur(vGrid.Text)
   Else
        lngCasos = -1
        vGrid.col = 9
        curAportes = CCur(vGrid.Text) * -1
        vGrid.col = 10
        curRendi = CCur(vGrid.Text) * -1
        vGrid.col = 12
        curISR = CCur(vGrid.Text) * -1
   End If

    txtCasos.Text = Format(CLng(txtCasos.Text) + lngCasos, "###,###,###,##0")
    txtAportes.Text = Format(CCur(txtAportes.Text) + curAportes, "Standard")
    txtRendimientos.Text = Format(CCur(txtRendimientos.Text) + curRendi, "Standard")
    txtISR.Text = Format(CCur(txtISR.Text) + curISR, "Standard")
    
    txtTotal.Text = Format(CCur(txtAportes.Text) + CCur(txtRendimientos.Text), "Standard")
    txtNeto.Text = Format(CCur(txtAportes.Text) + CCur(txtRendimientos.Text) - CCur(txtISR.Text), "Standard")

End If 'Col = 1

End Sub



