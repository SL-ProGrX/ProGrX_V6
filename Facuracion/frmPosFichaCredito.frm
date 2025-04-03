VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#20.3#0"; "Codejock.Controls.v20.3.0.ocx"
Begin VB.Form frmPosFichaCredito 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Ficha de Crédito (Enlaces)"
   ClientHeight    =   7665
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   10710
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7665
   ScaleWidth      =   10710
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.RadioButton rbTipo 
      Height          =   255
      Index           =   0
      Left            =   2040
      TabIndex        =   29
      Top             =   1440
      Width           =   1935
      _Version        =   1310723
      _ExtentX        =   3408
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Crédito"
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
      Value           =   -1  'True
   End
   Begin XtremeSuiteControls.GroupBox gbFiadores 
      Height          =   2055
      Left            =   240
      TabIndex        =   14
      Top             =   4200
      Width           =   10335
      _Version        =   1310723
      _ExtentX        =   18224
      _ExtentY        =   3619
      _StockProps     =   79
      Caption         =   "Fiadores"
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      BorderStyle     =   1
      Begin XtremeSuiteControls.PushButton cmdGuardarFiadores 
         Height          =   612
         Left            =   8640
         TabIndex        =   22
         Top             =   1200
         Width           =   1452
         _Version        =   1310723
         _ExtentX        =   2561
         _ExtentY        =   1080
         _StockProps     =   79
         Caption         =   "Guardar"
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
         TextAlignment   =   1
         Appearance      =   16
         Picture         =   "frmPosFichaCredito.frx":0000
         ImageAlignment  =   4
      End
      Begin XtremeSuiteControls.FlatEdit txtFia1Ced 
         Height          =   315
         Left            =   1320
         TabIndex        =   41
         Top             =   720
         Width           =   1695
         _Version        =   1310723
         _ExtentX        =   2990
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   8.25
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
      Begin XtremeSuiteControls.FlatEdit txtFia1Nombre 
         Height          =   315
         Left            =   3000
         TabIndex        =   42
         Top             =   720
         Width           =   4455
         _Version        =   1310723
         _ExtentX        =   7858
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   8.25
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
      Begin XtremeSuiteControls.FlatEdit txtFia1Deuda 
         Height          =   315
         Left            =   7440
         TabIndex        =   43
         Top             =   720
         Width           =   495
         _Version        =   1310723
         _ExtentX        =   873
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   "0"
         Alignment       =   1
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtFia2Ced 
         Height          =   315
         Left            =   1320
         TabIndex        =   44
         Top             =   1080
         Width           =   1695
         _Version        =   1310723
         _ExtentX        =   2990
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   8.25
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
      Begin XtremeSuiteControls.FlatEdit txtFia2Nombre 
         Height          =   315
         Left            =   3000
         TabIndex        =   45
         Top             =   1080
         Width           =   4455
         _Version        =   1310723
         _ExtentX        =   7858
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   8.25
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
      Begin XtremeSuiteControls.FlatEdit txtFia2Deuda 
         Height          =   315
         Left            =   7440
         TabIndex        =   46
         Top             =   1080
         Width           =   495
         _Version        =   1310723
         _ExtentX        =   873
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   "0"
         Alignment       =   1
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtFia3Ced 
         Height          =   315
         Left            =   1320
         TabIndex        =   47
         Top             =   1440
         Width           =   1695
         _Version        =   1310723
         _ExtentX        =   2990
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   8.25
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
      Begin XtremeSuiteControls.FlatEdit txtFia3Nombre 
         Height          =   315
         Left            =   3000
         TabIndex        =   48
         Top             =   1440
         Width           =   4455
         _Version        =   1310723
         _ExtentX        =   7858
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   8.25
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
      Begin XtremeSuiteControls.FlatEdit txtFia3Deuda 
         Height          =   315
         Left            =   7440
         TabIndex        =   49
         Top             =   1440
         Width           =   495
         _Version        =   1310723
         _ExtentX        =   873
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   "0"
         Alignment       =   1
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Fiador # 1"
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
         Index           =   7
         Left            =   360
         TabIndex        =   20
         Top             =   720
         Width           =   1332
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Fiador # 2"
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
         Left            =   360
         TabIndex        =   19
         Top             =   1080
         Width           =   1332
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Fiador # 3"
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
         Index           =   9
         Left            =   360
         TabIndex        =   18
         Top             =   1440
         Width           =   1332
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         Caption         =   "Cédula"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   312
         Index           =   10
         Left            =   1320
         TabIndex        =   17
         Top             =   360
         Width           =   1692
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         Caption         =   "Nombre"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   312
         Index           =   11
         Left            =   3000
         TabIndex        =   16
         Top             =   360
         Width           =   4452
      End
      Begin VB.Image imgFichaCliente 
         Height          =   252
         Left            =   8040
         Picture         =   "frmPosFichaCredito.frx":0731
         Stretch         =   -1  'True
         ToolTipText     =   "Ficha de Cliente"
         Top             =   360
         Width           =   252
      End
      Begin VB.Image imgNuevoFia1 
         Height          =   252
         Left            =   8040
         Picture         =   "frmPosFichaCredito.frx":0F00
         Stretch         =   -1  'True
         ToolTipText     =   "Elimina al Fiador"
         Top             =   720
         Width           =   252
      End
      Begin VB.Image imgNuevoFia2 
         Height          =   252
         Left            =   8040
         Picture         =   "frmPosFichaCredito.frx":18AD
         Stretch         =   -1  'True
         ToolTipText     =   "Elimina al Fiador"
         Top             =   1080
         Width           =   252
      End
      Begin VB.Image imgNuevoFia3 
         Height          =   252
         Left            =   8040
         Picture         =   "frmPosFichaCredito.frx":225A
         Stretch         =   -1  'True
         ToolTipText     =   "Elimina al Fiador"
         Top             =   1440
         Width           =   252
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         Caption         =   "% "
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   312
         Index           =   13
         Left            =   7440
         TabIndex        =   15
         Top             =   360
         Width           =   492
      End
   End
   Begin XtremeSuiteControls.GroupBox gbCredito 
      Height          =   1815
      Left            =   240
      TabIndex        =   3
      Top             =   2280
      Width           =   10335
      _Version        =   1310723
      _ExtentX        =   18230
      _ExtentY        =   3201
      _StockProps     =   79
      Caption         =   "Datos del Crédito"
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      BorderStyle     =   1
      Begin XtremeSuiteControls.PushButton cmdGuardar 
         Height          =   612
         Left            =   8640
         TabIndex        =   21
         Top             =   1200
         Width           =   1452
         _Version        =   1310723
         _ExtentX        =   2561
         _ExtentY        =   1080
         _StockProps     =   79
         Caption         =   "Guardar"
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
         TextAlignment   =   1
         Appearance      =   16
         Picture         =   "frmPosFichaCredito.frx":2C07
         ImageAlignment  =   4
      End
      Begin XtremeSuiteControls.FlatEdit txtNotas 
         Height          =   1395
         Left            =   960
         TabIndex        =   35
         Top             =   360
         Width           =   4095
         _Version        =   1310723
         _ExtentX        =   7223
         _ExtentY        =   2461
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
      Begin XtremeSuiteControls.FlatEdit txtMonto 
         Height          =   315
         Left            =   6360
         TabIndex        =   50
         Top             =   360
         Width           =   1935
         _Version        =   1310723
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
         Alignment       =   1
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtCuota 
         Height          =   315
         Left            =   6360
         TabIndex        =   51
         Top             =   1440
         Width           =   1935
         _Version        =   1310723
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
         Alignment       =   1
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtTasa 
         Height          =   315
         Left            =   7200
         TabIndex        =   52
         Top             =   720
         Width           =   1095
         _Version        =   1310723
         _ExtentX        =   1931
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
         Alignment       =   1
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtPlazo 
         Height          =   315
         Left            =   7200
         TabIndex        =   53
         Top             =   1080
         Width           =   1095
         _Version        =   1310723
         _ExtentX        =   1931
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
         Alignment       =   1
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin VB.Label Label1 
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
         Index           =   0
         Left            =   5280
         TabIndex        =   8
         Top             =   360
         Width           =   1092
      End
      Begin VB.Label Label1 
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
         Height          =   252
         Index           =   1
         Left            =   5280
         TabIndex        =   7
         Top             =   720
         Width           =   852
      End
      Begin VB.Label Label1 
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
         Height          =   252
         Index           =   2
         Left            =   5280
         TabIndex        =   6
         Top             =   1080
         Width           =   852
      End
      Begin VB.Label Label1 
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
         Index           =   3
         Left            =   5280
         TabIndex        =   5
         Top             =   1440
         Width           =   852
      End
      Begin VB.Label Label1 
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
         Index           =   12
         Left            =   240
         TabIndex        =   4
         Top             =   360
         Width           =   732
      End
   End
   Begin XtremeSuiteControls.GroupBox gbOperacion 
      Height          =   1215
      Left            =   240
      TabIndex        =   9
      Top             =   6360
      Width           =   10335
      _Version        =   1310723
      _ExtentX        =   18230
      _ExtentY        =   2143
      _StockProps     =   79
      Caption         =   "Operación de Crédito: "
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      BorderStyle     =   1
      Begin VB.ComboBox cboMes 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   7.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "frmPosFichaCredito.frx":3338
         Left            =   4560
         List            =   "frmPosFichaCredito.frx":3363
         Style           =   2  'Dropdown List
         TabIndex        =   12
         ToolTipText     =   "Mes a procesar"
         Top             =   720
         Width           =   1815
      End
      Begin XtremeSuiteControls.PushButton cmdEnvio 
         Height          =   612
         Left            =   6840
         TabIndex        =   23
         Top             =   480
         Width           =   1812
         _Version        =   1310723
         _ExtentX        =   3196
         _ExtentY        =   1080
         _StockProps     =   79
         Caption         =   "Enviar a Cobro"
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
         TextAlignment   =   1
         Appearance      =   16
         Picture         =   "frmPosFichaCredito.frx":33CC
         ImageAlignment  =   4
      End
      Begin XtremeSuiteControls.PushButton cmdReporte 
         Height          =   612
         Left            =   8640
         TabIndex        =   24
         Top             =   480
         Width           =   1452
         _Version        =   1310723
         _ExtentX        =   2561
         _ExtentY        =   1080
         _StockProps     =   79
         Caption         =   "Reporte"
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
         TextAlignment   =   1
         Appearance      =   16
         Picture         =   "frmPosFichaCredito.frx":3C9D
         ImageAlignment  =   4
      End
      Begin XtremeSuiteControls.FlatEdit txtOperacion 
         Height          =   315
         Left            =   2280
         TabIndex        =   39
         Top             =   360
         Width           =   1935
         _Version        =   1310723
         _ExtentX        =   3413
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   16777152
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16777152
         Alignment       =   1
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtSaldo 
         Height          =   315
         Left            =   2280
         TabIndex        =   40
         Top             =   720
         Width           =   1935
         _Version        =   1310723
         _ExtentX        =   3413
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   16777152
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16777152
         Alignment       =   1
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Primer Deducción"
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
         Left            =   4560
         TabIndex        =   13
         Top             =   480
         Width           =   1692
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "No. Operación"
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
         Index           =   15
         Left            =   960
         TabIndex        =   11
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label1 
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
         Index           =   14
         Left            =   960
         TabIndex        =   10
         Top             =   720
         Width           =   1092
      End
   End
   Begin XtremeSuiteControls.ComboBox cboGarantia 
      Height          =   330
      Left            =   7680
      TabIndex        =   25
      Top             =   1800
      Width           =   2055
      _Version        =   1310723
      _ExtentX        =   3625
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
   Begin XtremeSuiteControls.ComboBox cboLinea 
      Height          =   330
      Left            =   1680
      TabIndex        =   28
      Top             =   1800
      Width           =   4815
      _Version        =   1310723
      _ExtentX        =   8493
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
   Begin XtremeSuiteControls.RadioButton rbTipo 
      Height          =   255
      Index           =   1
      Left            =   4080
      TabIndex        =   30
      Top             =   1440
      Width           =   2295
      _Version        =   1310723
      _ExtentX        =   4043
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Cuenta por Cobrar"
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
   End
   Begin XtremeSuiteControls.FlatEdit txtCedula 
      Height          =   315
      Left            =   1680
      TabIndex        =   31
      Top             =   720
      Width           =   2055
      _Version        =   1310723
      _ExtentX        =   3625
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
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtNombre 
      Height          =   315
      Left            =   3720
      TabIndex        =   32
      Top             =   720
      Width           =   6015
      _Version        =   1310723
      _ExtentX        =   10610
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
   Begin XtremeSuiteControls.FlatEdit txtCodInst 
      Height          =   315
      Left            =   1680
      TabIndex        =   33
      Top             =   1080
      Width           =   2055
      _Version        =   1310723
      _ExtentX        =   3625
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
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtDescInst 
      Height          =   315
      Left            =   3720
      TabIndex        =   34
      Top             =   1080
      Width           =   6015
      _Version        =   1310723
      _ExtentX        =   10610
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
   Begin XtremeSuiteControls.FlatEdit txtFactura 
      Height          =   315
      Left            =   1680
      TabIndex        =   36
      Top             =   240
      Width           =   3015
      _Version        =   1310723
      _ExtentX        =   5318
      _ExtentY        =   556
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   16777152
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16777152
      Alignment       =   2
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtTipo 
      Height          =   315
      Left            =   4680
      TabIndex        =   37
      Top             =   240
      Width           =   2535
      _Version        =   1310723
      _ExtentX        =   4471
      _ExtentY        =   556
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   16777152
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16777152
      Alignment       =   2
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtFecha 
      Height          =   315
      Left            =   7200
      TabIndex        =   38
      Top             =   240
      Width           =   2535
      _Version        =   1310723
      _ExtentX        =   4471
      _ExtentY        =   556
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   16777152
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16777152
      Alignment       =   2
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Crd/CxC"
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
      Index           =   1
      Left            =   360
      TabIndex        =   27
      Top             =   1800
      Width           =   975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Garantía"
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
      Left            =   6720
      TabIndex        =   26
      Top             =   1800
      Width           =   1095
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Planilla"
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
      Index           =   0
      Left            =   360
      TabIndex        =   2
      Top             =   1080
      Width           =   855
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Cliente"
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
      Left            =   360
      TabIndex        =   1
      Top             =   720
      Width           =   855
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "No. Factura"
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
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   1095
   End
End
Attribute VB_Name = "frmPosFichaCredito"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vPaso As Boolean


Private Sub sbCargaDatos()
Dim strSQL As String, rs As New ADODB.Recordset
Dim pCxC_Tipo As String, pCxC_Codigo As String

On Error GoTo vError

'Busca la Planilla a la que pertenece el Cliente
strSQL = "select I.cod_institucion,I.descripcion" _
       & " from pv_clientes C inner join instituciones I on C.cod_institucion = I.cod_institucion" _
       & " where C.cedula = '" & txtCedula & "'"
Call OpenRecordSet(rs, strSQL)
If Not rs.EOF And Not rs.BOF Then
  txtCodInst = rs!cod_institucion
  txtDescInst = rs!Descripcion
Else
  MsgBox "Verifique la ficha del Cliente, pues su identificacion de planilla no esta definida...", vbExclamation
End If
rs.Close

'Busca si ya existe la información del crédito

strSQL = "select P.*,isnull(R.saldo,0) as Saldo" _
       & " from pv_preCredito P left join reg_creditos R on P.cxc_operacion = R.id_solicitud" _
       & " where P.cod_factura = '" & txtFactura & "' and P.tipo = '" & txtTipo.Tag & "'"
Call OpenRecordSet(rs, strSQL)
If Not rs.EOF And Not rs.BOF Then
   txtPlazo = rs!Plazo
   txtTasa = rs!Tasa
   txtMonto = Format(rs!Monto, "Standard")
   txtCuota = Format(rs!Cuota, "Standard")
   txtOperacion = rs!cxc_operacion & ""
   txtSaldo = Format(rs!Saldo, "Standard")
   txtNotas = rs!observacion & ""
   If IsNull(rs!cxc_operacion) Then
     cmdEnvio.Enabled = True
     cmdGuardar.Enabled = True
     cmdGuardarFiadores.Enabled = True
   Else
     cmdEnvio.Enabled = False
     cmdGuardar.Enabled = False
     cmdGuardarFiadores.Enabled = False
   End If
   gbFiadores.Enabled = True
   Call sbFiadores_Load
   
   'Carga Linea
   Select Case rs!CxC_Tipo
     Case "C", "CRD"
     Case "CxC"
   End Select
   
   
   
Else
   txtPlazo = 1
   txtTasa = 0
   txtOperacion = 0
   txtSaldo = 0
   txtNotas = ""
   cmdGuardar.Enabled = True
   cmdEnvio.Enabled = False
   gbFiadores.Enabled = False
   cmdGuardarFiadores.Enabled = False
End If

rs.Close

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub sbClientes_CxC_Sincroniza(vCedula As String)
Dim strSQL As String, rs As New ADODB.Recordset

'Verifica Existencia en --Socios--, de lo contrario los crea segun la Base de Cliente

strSQL = "select isnull(count(*),0) as Existe from cxc_personas where cedula = '" & vCedula & "'"
Call OpenRecordSet(rs, strSQL)
If rs!Existe > 0 Then Exit Sub
rs.Close

strSQL = "select * from pv_clientes where cedula = '" & vCedula & "'"
Call OpenRecordSet(rs, strSQL)

If Not rs.EOF And Not rs.BOF Then
   strSQL = "insert into CxC_Personas(cedula,Tipo_Id,nombre,razon_social,celular,telefono1,telefono2,fax,sexo,estadoCivil,fecha_nacimiento" _
          & ",apto_postal,email_01,email_02,webSite,notas,direccion,distrito,provincia,canton,credito_cerrado,Cliente_Exento,cod_categoria,categoria_fecha" _
          & ",ADELANTO_PERMITE, ADELANTO_MODIFICA,ADELANTO_PORCENTAJE, CREDITO_LIMITE, ACTIVO, ADELANTO_COMISION_APL, ADELANTO_COMISION)" _
          & " values('" & rs!Cedula & "'," & rs!Tipo_id & ",'" & rs!Nombre & "','" & rs!Nombre & "','" & rs!Celular _
          & "','" & rs!telefono1 & "','" & rs!telefono1 & "','" & rs!fax & "','" & rs!sexo & "','" & rs!EstadoCivil & "','" _
          & Format(rs!fecha_nacimiento, "yyyy/mm/dd") & "','" & rs!apto_postal & "','" & rs!Email & "','','','" _
          & txtNotas.Text & "','" & rs!Direccion & "','" & rs!distrito & "','" _
          & rs!Provincia & "','" & rs!Canton _
          & "'," & rs!Credito_Cerrado & "," & rs!Cliente_Excento & ",'01',dbo.MyGetdate()," _
          & "0,0,0,0,1,0,0)"
  Call ConectionExecute(strSQL)

End If
rs.Close

End Sub



Private Sub sbClientes_Crd_Sincroniza(vCedula As String)
Dim strSQL As String, rs As New ADODB.Recordset

'Verifica Existencia en --Socios--, de lo contrario los crea segun la Base de Cliente

strSQL = "select isnull(count(*),0) as Existe from socios where cedula = '" & vCedula & "'"
Call OpenRecordSet(rs, strSQL)
If rs!Existe > 0 Then Exit Sub
rs.Close

strSQL = "select * from pv_clientes where cedula = '" & vCedula & "'"
Call OpenRecordSet(rs, strSQL)

If Not rs.EOF And Not rs.BOF Then
  strSQL = "insert socios(cod_institucion,cod_departamento,cod_seccion,id_promotor,cod_sector,cod_profesion" _
         & ",cedula,nombre,boleta,estadoactual,provincia,canton,distrito,direccion,fecha_nac,estadoCivil" _
         & ",hijos,estadoLaboral,fechaIngreso,Apto,Af_Email,Notas,ultimo_estado " _
         & ",ind_liquidacion,bloqueo) values(" & rs!cod_institucion _
         & ",'','',1,1,1,'" & Trim(rs!Cedula) & "','" & Trim(rs!Nombre) & "',0,'N'," & rs!Provincia _
         & "," & rs!Canton & ",'" & rs!distrito & "','" & Trim(rs!Direccion) & "','" _
         & Format(rs!fecha_nacimiento, "yyyy/mm/dd") & "','" & rs!EstadoCivil & "',0,1,dbo.MyGetdate()" _
         & ",'" & Trim(rs!apto_postal) & "','" & Trim(rs!Email) & "','','S',0,0)"
  Call ConectionExecute(strSQL)
  
  strSQL = "insert ahorro_consolidado(cedula,ahorro,aporte) values('" & Trim(rs!Cedula) & "',0,0)"
  Call ConectionExecute(strSQL)
End If
rs.Close

End Sub

Private Function fxCredito_Registro(xCedula As String, xProceso As Currency, xNotas As String _
                            , xPorcentaje As Currency, Optional xTipo As String = "D") As Long

Dim strSQL As String, rs As New ADODB.Recordset
Dim xCodigo As String, xOperacion As Long, xFactura As String, vGarantia As String
Dim vFecha As Date, vMonto As Currency


On Error GoTo vError

vFecha = fxFechaServidor
xFactura = "F" & txtTipo.Tag & "-" & Format(txtFactura, "0000000000")

'Revisar si el Cliente y Crearla si no existe
Call sbClientes_Crd_Sincroniza(xCedula)

''Saca el Codigo de Enlace de Credito
'strSQL = "select cod_credito from PV_PARINSTITUCIONES where cod_institucion = " & txtCodInst
'Call OpenRecordSet(rs, strSQL)
'  xCodigo = Trim(rs!cod_credito)
'rs.Close
'

strSQL = "select * from pv_preCredito where cod_factura = '" & txtFactura _
       & "' and tipo = '" & txtTipo.Tag & "'"
Call OpenRecordSet(rs, strSQL)


xCodigo = rs!Codigo
vGarantia = Trim(rs!Garantia)
vMonto = rs!Monto * (xPorcentaje / 100)
 
If xTipo = "D" Then
    'Registra Credito
    strSQL = "insert reg_creditos(codigo,id_comite,cedula,montoapr,monto_girado" _
           & ",saldo,amortiza,interesc,saldo_mes,cuota,int,interesv,plazo,userrec" _
           & ",userfor,usertesoreria,tesoreria,fechasol,fechaforp" _
           & ",fecha_calculo_int,garantia,primer_cuota,tdocumento,ndocumento" _
           & ",firma_deudor,observacion,estado,prideduc,fecult,estadosol,documento_referido)" _
           & " values('" & UCase(xCodigo) & "',1,'" & Trim(xCedula) & "'," _
           & vMonto & ",0," & vMonto & ",0,0," & vMonto & "," & CCur(fxCalcula_Cuota(CDbl(vMonto), rs!Plazo, rs!Tasa)) & "," & rs!Tasa _
           & "," & rs!Tasa & "," & rs!Plazo & ",'" & glogon.Usuario & "','" & glogon.Usuario _
           & "','" & glogon.Usuario & "','" & Format(vFecha, "yyyy/mm/dd") & "','" _
           & Format(vFecha, "yyyy/mm/dd") & "','" & Format(vFecha, "yyyy/mm/dd") & "','" _
           & Format(vFecha, "yyyy/mm/dd") & "','" & vGarantia & "','F','SC','POS'" _
           & ",1,'TRASLADO DE COBRO DE FACTURA DEL POS','A'," & xProceso _
           & "," & fxFechaProcesoAnterior(xProceso) & ",'F','" & xFactura & "')"
    Call ConectionExecute(strSQL)
Else
    'Registra Credito
    strSQL = "insert reg_creditos(codigo,id_comite,cedula,montoapr,monto_girado" _
           & ",saldo,amortiza,interesc,saldo_mes,cuota,int,interesv,plazo,userrec" _
           & ",userfor,usertesoreria,tesoreria,fechasol,fechaforp" _
           & ",fecha_calculo_int,garantia,primer_cuota,tdocumento,ndocumento" _
           & ",firma_deudor,observacion,estado,prideduc,fecult,estadosol,documento_referido)" _
           & " values('" & UCase(xCodigo) & "',1,'" & Trim(txtCedula) & "'," _
           & vMonto & ",0," & vMonto & ",0,0," & vMonto & "," & CCur(fxCalcula_Cuota(CDbl(vMonto), rs!Plazo, rs!Tasa)) & "," & rs!Tasa _
           & "," & rs!Tasa & "," & rs!Plazo & ",'" & glogon.Usuario & "','" & glogon.Usuario _
           & "','" & glogon.Usuario & "','" & Format(vFecha, "yyyy/mm/dd") & "','" _
           & Format(vFecha, "yyyy/mm/dd") & "','" & Format(vFecha, "yyyy/mm/dd") & "','" _
           & Format(vFecha, "yyyy/mm/dd") & "','N','F','SC','POS'" _
           & ",1,'TRASLADO DE COBRO DE FACTURA DEL POS **** FIADOR ****','A'," & xProceso _
           & "," & fxFechaProcesoAnterior(xProceso) & ",'F','" & xFactura & "')"
    Call ConectionExecute(strSQL)
End If
rs.Close

strSQL = "select max(id_solicitud) as Operacion from reg_creditos where estado = 'A' and cedula = '" _
       & xCedula & "' and codigo = '" & xCodigo & "'"
Call OpenRecordSet(rs, strSQL)
 xOperacion = rs!Operacion
rs.Close


If xTipo = "D" Then
    'Crea a los fiadores y los enlaza al credito
    strSQL = "select cedula from pv_preCredFia where cod_factura = '" & txtFactura _
           & "' and tipo = '" & txtTipo.Tag & "'"
    Call OpenRecordSet(rs, strSQL, 0)
    Do While Not rs.EOF
     Call sbClientes_Crd_Sincroniza(rs!Cedula)
     
     strSQL = "insert into fiadores(id_solicitud,codigo,cedulaf,nombre,firma,estado" _
            & ",salario,devengado,liquidez,interno) values(" & xOperacion & ",'" _
            & xCodigo & "','" & Trim(rs!Cedula) & "','','S','A',0,0,0,1)"
     Call ConectionExecute(strSQL)
     rs.MoveNext
    Loop
    rs.Close
    
    'Actualiza el Estado de la CxC
    strSQL = "update pv_preCredito set cxc_operacion = " & xOperacion _
           & ",cxc_traslado = dbo.MyGetdate(),estado = 'T',cxc_tipo = 'CRD'" _
           & " where cod_factura = '" & txtFactura & "' and tipo = '" & txtTipo.Tag & "'"
    Call ConectionExecute(strSQL)
Else

    'Crea al deudor como fiador de la deuda, para posibles traslados al mismo
     strSQL = "insert into fiadores(id_solicitud,codigo,cedulaf,nombre,firma,estado" _
            & ",salario,devengado,liquidez,interno) values(" & xOperacion & ",'" _
            & xCodigo & "','" & Trim(txtCedula) & "','','S','A',0,0,0,1)"
     Call ConectionExecute(strSQL)
    
    'Actualiza el # Operacion del fiador
    strSQL = "update pv_preCredFia set operacion = " & xOperacion _
           & " where cod_factura = '" & txtFactura & "' and tipo = '" & txtTipo.Tag _
           & "' and cedula = '" & xCedula & "'"
    Call ConectionExecute(strSQL)

End If

'Plan de Pagos
If GLOBALES.SysPlanPagos = 1 Then
    strSQL = "exec spCrdPlanPagos " & xOperacion
    Call ConectionExecute(strSQL)
End If

'Procesa Asiento de Credito
strSQL = "exec spCRDFormalizaAsiento " & xOperacion & ",'" & Format(vFecha, "yyyy/mm/dd") _
            & "'," & xProceso & ",'" & glogon.Usuario & "'"
Call ConectionExecute(strSQL)

fxCredito_Registro = xOperacion

Exit Function

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Function



Private Function fxCxC_Registro(xCedula As String, xProceso As Currency, xNotas As String) As Long

Dim strSQL As String, rs As New ADODB.Recordset
Dim xCodigo As String, xOperacion As Long, xFactura As String
Dim vFecha As Date, vMonto As Currency
Dim vPlazoDias As Long, vPlazo As Long, vTasa As Currency, vCuota As Currency


On Error GoTo vError

vFecha = fxFechaServidor
xFactura = "F" & txtTipo.Tag & "-" & Format(txtFactura, "0000000000")

'Revisar si el Cliente y Crearla si no existe
Call sbClientes_CxC_Sincroniza(xCedula)


strSQL = "select * from pv_preCredito where cod_factura = '" & txtFactura _
       & "' and tipo = '" & txtTipo.Tag & "'"
Call OpenRecordSet(rs, strSQL)


xCodigo = rs!Codigo
vMonto = rs!Monto
vPlazoDias = rs!Plazo * 30
vPlazo = rs!Plazo
vCuota = rs!Cuota
vTasa = rs!Tasa

strSQL = "select isnull(max(Operacion),0) + 1 as 'Operacion' from CxC_Cuentas"
Call OpenRecordSet(rs, strSQL)
  txtOperacion.Text = rs!Operacion
  xOperacion = rs!Operacion
rs.Close



strSQL = "insert CxC_Cuentas(OPERACION,CEDULA,CEDULA_PAGADOR,COD_CONCEPTO,COD_OFICINA,NOTAS,MONTO,SALDO,REBAJOS_TOTAL" _
       & ",EMITIR_TIPO,EMITIR_BANCO,EMITIR_CUENTA,DESEMBOLSO_MONTO,TIPO_PLAZO,TASA_CORRIENTE,TASA_MORA,CUOTA,DIAS_PLAZO, PLAZO,AMORTIZA,INTERESC" _
       & ",ESTADO,NUM_DOCUMENTO,COD_CONTRATO,REGISTRO_FECHA,REGISTRO_USUARIO,FECHA_ULTMOV,AUTORIZA_ESTADO" _
       & ", ADELANTO_MONTO, ADELANTO_PORCENTAJE, DESEMBOLSO_REALIZADO, DESEMBOLSO_PENDIENTE, CEDULA_AUTORIZADO" _
       & ", ADELANTO_COMISION_APL, ADELANTO_COMISION, ADELANTO_COMISION_DIAS" _
       & ", INGRESOS_TOTAL, ACTIVA_FECHA, ACTIVA_USUARIO,Tesoreria_Fecha, Tesoreria_Solicitud, Tesoreria_Estado,Tesoreria_Usuario, FREQ_PAGO  ) " _
       & " VALUES(" & xOperacion & ",'" & txtCedula.Text & "',Null,'" & xCodigo & "','" & GLOBALES.gOficinaTitular _
       & "','" & Trim(txtNotas.Text) & "'," & vMonto & "," & vMonto & ",0,'ND', 1,''," & vMonto & ",'M'," & vTasa _
       & "," & vTasa & "," & vCuota & "," & vPlazoDias & "," & vPlazo & ",0,0,'A','" & xFactura _
       & "',Null,dbo.MyGetdate(),'" & glogon.Usuario & "',dbo.MyGetdate(),'P'" _
       & ",0,0," & vMonto & ",0,'',0,0,0,0, dbo.MyGetdate(),'" & glogon.Usuario & "',dbo.MyGetdate(),0,'C','" & glogon.Usuario & "', 30)"
       
Call ConectionExecute(strSQL)
       

'Plan de Pagos
strSQL = "exec spCxC_CuentaPlanPagos " & xOperacion
Call ConectionExecute(strSQL)

'Procesa Asiento de Credito
strSQL = "exec spCxC_CuentaActivaAsiento " & xOperacion & ",'" & Format(vFecha, "yyyy/mm/dd") _
            & "','" & glogon.Usuario & "'"
Call ConectionExecute(strSQL)


'Actualiza el Estado de la CxC
strSQL = "update pv_preCredito set cxc_operacion = " & xOperacion _
       & ",cxc_traslado = dbo.MyGetdate(),estado = 'T',cxc_tipo = 'CxC'" _
       & " where cod_factura = '" & txtFactura & "' and tipo = '" & txtTipo.Tag & "'"
Call ConectionExecute(strSQL)

fxCxC_Registro = xOperacion

Exit Function

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Function


Private Sub cboLinea_Click()
If vPaso Then Exit Sub

Dim strSQL As String

cboGarantia.Clear

If rbTipo.Item(0).Value Then
        strSQL = "select T.Garantia as 'IdX',rtrim(T.descripcion) as 'ItmX'" _
               & " from crd_catalogo_garantias C inner join crd_garantia_tipos T on C.garantia = T.garantia" _
               & " where C.codigo = '" & cboLinea.ItemData(cboLinea.ListIndex) & "'"
        Call sbCbo_Llena_New(cboGarantia, strSQL, False, True)
End If

End Sub

Private Sub cmdEnvio_Click()
Dim strSQL As String, rs As New ADODB.Recordset
Dim vFecha As Date, vFactura As String, vOperacion As Long
Dim vFechaProceso As Currency, vPorcentaje As Currency
Dim vMes As Integer

On Error GoTo vError


If cboLinea.ListCount = 0 Then
    MsgBox "No Existe una Línea de Credito o Concepto de Cuentas por Cobrar!", vbExclamation
    Exit Sub
End If

If rbTipo.Item(0).Value And cboGarantia.ListCount = 0 Then
    MsgBox "No Existe una Garantía Válida!", vbExclamation
    Exit Sub
End If


vFecha = fxFechaServidor
vMes = fxConvierteMES(cboMes.Text)

'Indica la Primer Deducción
If Month(vFecha) > vMes Then
    vFechaProceso = (Year(vFecha) + 1) & Format(vMes, "00")
Else
    vFechaProceso = Year(vFecha) & Format(vMes, "00")
End If


'Revisar si ya fue enviada
strSQL = "select isnull(count(*),0) as Existe " _
       & " from pv_preCredito where cod_factura = '" & txtFactura _
       & "' and tipo = '" & txtTipo.Tag & "' and cxc_operacion is null"
Call OpenRecordSet(rs, strSQL)
If rs!Existe = 0 Then
  MsgBox "Esta factura ya fue enviada al Cobro anteriormente...", vbInformation
  Exit Sub
End If
rs.Close



'Sincronizar Codigo Institucion, solo se necesita para el Deudor
strSQL = "select cod_institucion from socios where cedula = '" & txtCedula _
       & "' and cod_institucion <> " & txtCodInst
Call OpenRecordSet(rs, strSQL)
If Not rs.EOF And Not rs.BOF Then
    strSQL = "update pv_clientes set cod_institucion = " & rs!cod_institucion _
           & " where cedula = '" & txtCedula & "'"
    Call ConectionExecute(strSQL)
End If
rs.Close



'Validación para Creditos
'If rbTipo.Item(0).Value Then
'        'Revisar que la planilla (institucion) tenga un codigo de credito válido
'        strSQL = "select isnull(count(*),0) as Existe " _
'               & " from PV_PARINSTITUCIONES I inner join Catalogo C on I.cod_credito = C.codigo" _
'               & " where I.cod_institucion = " & txtCodInst
'        Call OpenRecordSet(rs, strSQL)
'        If rs!Existe = 0 Then
'          MsgBox "No se ha establecido el código de Enlace de Crédito para la Planilla del Cliente...", vbInformation
'          Exit Sub
'        End If
'        rs.Close
'
'End If 'Valida Credito

'--------------------------------------

Me.MousePointer = vbHourglass

'Creditos
If rbTipo.Item(0).Value Then
        'Calcular porcentaje de deuda para el deudor
        strSQL = "select isnull(sum(porc_deuda),0) as Porcentaje from pv_preCredFia" _
                   & " where cod_factura = '" & txtFactura & "' and tipo = '" & txtTipo.Tag & "'"
        Call OpenRecordSet(rs, strSQL)
         vPorcentaje = 100 - rs!Porcentaje
        rs.Close
        
        'envia al cobro al deudor
        If vPorcentaje > 0 Then
          vOperacion = fxCredito_Registro(txtCedula, vFechaProceso, "", vPorcentaje, "D")
          txtSaldo = Format(CCur(txtMonto) * (vPorcentaje / 100), "Standard")
          txtOperacion = vOperacion
        Else
            strSQL = "update pv_preCredito set cxc_operacion = " & 0 _
                   & ",cxc_traslado = dbo.MyGetdate(),estado = 'T',cxc_tipo = 'CRD'" _
                   & " where cod_factura = '" & txtFactura & "' and tipo = '" & txtTipo.Tag & "'"
            Call ConectionExecute(strSQL)
            txtOperacion = 0
            txtSaldo = "0.00"
        End If
        
        'Envia a fiadores con % de deudas
        strSQL = "select * from pv_preCredFia" _
               & " where cod_factura = '" & txtFactura & "' and tipo = '" & txtTipo.Tag _
               & "' and isnull(porc_deuda,0) > 0"
        Call OpenRecordSet(rs, strSQL)
        Do While Not rs.EOF
          vOperacion = fxCredito_Registro(rs!Cedula, vFechaProceso, "", rs!porc_deuda, "F")
          rs.MoveNext
        Loop
        rs.Close
End If


'Cuenta x Cobrar
If rbTipo.Item(1).Value Then
          vOperacion = fxCxC_Registro(txtCedula.Text, vFechaProceso, txtNotas.Text)
End If


cmdEnvio.Enabled = False
cmdGuardar.Enabled = False
cmdGuardarFiadores.Enabled = False

Me.MousePointer = vbDefault
MsgBox "Traslada Factura al Sistema de Crédito/Cuentas por Cobrar...", vbInformation

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub cmdGuardar_Click()
Dim strSQL As String, rs As New ADODB.Recordset
Dim vEdita As Boolean, vTipoCxC As String

On Error GoTo vError

vEdita = False

If cboLinea.ListCount = 0 Then
    MsgBox "No se ha Indicado o No tiene Líneas de Crédito o Conceptos de CxC activados!", vbExclamation
    Exit Sub
End If

If rbTipo.Item(0).Value Then
    vTipoCxC = "CRD"
Else
    vTipoCxC = "CxC"
End If

If cboGarantia.ListCount = 0 And vTipoCxC = "CRD" Then
    MsgBox "No se ha Indicado la Garantía!", vbExclamation
    Exit Sub
End If



strSQL = "select isnull(count(*),0) as Existe from pv_preCredito" _
       & " where cod_factura = '" & txtFactura & "' and tipo= '" & txtTipo.Tag & "'"
Call OpenRecordSet(rs, strSQL)
If rs!Existe > 0 Then vEdita = True
rs.Close

If vEdita Then
   strSQL = "update pv_preCredito set plazo = " & txtPlazo _
          & ",cuota = " & CCur(txtCuota) & ",tasa = " & CCur(txtTasa) & ",'" & cboLinea.ItemData(cboLinea.ListIndex) & "'" _
          & ",observacion = '" & txtNotas & ", CxC_Tipo = '" & vTipoCxC & "', Garantia = '" & cboGarantia.ItemData(cboGarantia.ListIndex) _
          & "' where cod_factura = '" & txtFactura & "' and tipo = '" _
          & txtTipo.Tag & "' and cxc_operacion is null"
Else
   strSQL = "insert pv_preCredito(cod_factura,tipo,monto,plazo,tasa,cuota,observacion,codigo,garantia,cxc_Tipo) values('" _
          & txtFactura & "','" & txtTipo.Tag & "'," & CCur(txtMonto) & "," & txtPlazo & "," & txtTasa & "," _
          & CCur(txtCuota) & ",'" & txtNotas & "','" & cboLinea.ItemData(cboLinea.ListIndex) _
          & "','" & cboGarantia.ItemData(cboGarantia.ListIndex) & "','" & vTipoCxC & "')"
   cmdEnvio.Enabled = True
   gbFiadores.Enabled = True
End If

Call ConectionExecute(strSQL)


MsgBox "Información Base para Credito/CxC Actualizada...", vbInformation

Call sbCargaDatos
Call RefrescaTags(Me)

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub cmdGuardarFiadores_Click()
Dim strSQL As String, curDeuda As Currency

On Error GoTo vError

curDeuda = CCur(txtFia1Deuda) + CCur(txtFia2Deuda) + CCur(txtFia3Deuda)
 
If curDeuda > 100 Then
  MsgBox "El porcentaje de las deudas para Asumir de los fiadores es mayor al 100%", vbCritical
  Exit Sub
End If
 
 
strSQL = "delete pv_preCredFia where cod_factura = '" & txtFactura _
       & "' and tipo = '" & txtTipo.Tag & "'"
Call ConectionExecute(strSQL)

If txtFia1Ced <> "" Then
  strSQL = "insert pv_preCredFia(cod_factura,tipo,cedula,porc_deuda) values('" _
         & txtFactura & "','" & txtTipo.Tag & "','" & txtFia1Ced & "'," & CCur(txtFia1Deuda) & ")"
  Call ConectionExecute(strSQL)
End If

If txtFia2Ced <> "" Then
  strSQL = "insert pv_preCredFia(cod_factura,tipo,cedula,porc_deuda) values('" _
         & txtFactura & "','" & txtTipo.Tag & "','" & txtFia2Ced & "'," & CCur(txtFia2Deuda) & ")"
  Call ConectionExecute(strSQL)
End If

If txtFia3Ced <> "" Then
  strSQL = "insert pv_preCredFia(cod_factura,tipo,cedula,porc_deuda) values('" _
         & txtFactura & "','" & txtTipo.Tag & "','" & txtFia3Ced & "'," & CCur(txtFia3Deuda) & ")"
  Call ConectionExecute(strSQL)
End If

MsgBox "Informacion de Fiadores Actualizada Satisfactoriamente...", vbInformation

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical


End Sub

Private Sub cmdReporte_Click()
MsgBox "En Desarrollo Boleta de Orden de Deducción", vbInformation
End Sub

Private Sub Form_Activate()
vModulo = 33
End Sub

Private Sub Form_Load()
Dim strSQL As String, rs As New ADODB.Recordset
Dim iMes As Byte

vModulo = 33

Call Formularios(Me)
Call RefrescaTags(Me)


strSQL = "select F.*,C.nombre" _
       & " from pv_facturacion F inner join pv_clientes C on F.cedula = C.cedula" _
       & " where cod_factura = '" & GLOBALES.gTag & "' and tipo = '" & GLOBALES.gTag2 & "'"
Call OpenRecordSet(rs, strSQL)
If Not rs.BOF And Not rs.EOF Then
    
  Select Case UCase(rs!Tipo)
    Case "A"
      txtTipo.Text = "Automáticas"
    Case "M"
      txtTipo.Text = "Manuales"
  End Select
  
  txtTipo.Tag = rs!Tipo
  txtCedula.Text = rs!Cedula
  txtNombre.Text = rs!Nombre
  txtFactura.Text = rs!cod_Factura
  txtMonto.Text = Format(rs!Total, "Standard")
  
  txtFecha.Text = Format(rs!fecha, "yyyy/mm/dd hh:mm:ss")

End If
rs.Close

iMes = Month(fxFechaServidor)

If iMes = 12 Then
 iMes = 1
Else
 iMes = iMes + 1
End If


cboMes.Text = fxConvierteMES(iMes)

Call sbCargaDatos

End Sub

Private Sub imgFichaCliente_Click()
Call MuestraForms(frmPosFichaCliente)
End Sub


Private Sub imgNuevoFia1_Click()
  txtFia1Ced = ""
  txtFia1Nombre = ""
  txtFia1Deuda = 0
End Sub

Private Sub imgNuevoFia2_Click()
  txtFia2Ced = ""
  txtFia2Nombre = ""
  txtFia2Deuda = 0
End Sub

Private Sub imgNuevoFia3_Click()
  txtFia3Ced = ""
  txtFia3Nombre = ""
  txtFia3Deuda = 0
End Sub

Private Sub sbFiadores_Load()
Dim strSQL As String, rs As New ADODB.Recordset
Dim i As Integer

  i = 1
  txtFia1Ced = ""
  txtFia1Nombre = ""
  txtFia1Deuda = 0
  txtFia1Deuda.ToolTipText = ""
  
  txtFia2Ced = ""
  txtFia2Nombre = ""
  txtFia2Deuda = 0
  txtFia2Deuda.ToolTipText = ""
  
  txtFia3Ced = ""
  txtFia3Nombre = ""
  txtFia3Deuda = 0
  txtFia3Deuda.ToolTipText = ""
  
  strSQL = "select C.cedula,C.nombre,F.porc_deuda,F.operacion" _
         & " from pv_preCredFia F inner join pv_clientes C on F.cedula = C.cedula" _
         & " where F.cod_factura = '" & txtFactura & "' and tipo = '" & txtTipo.Tag & "'"
  Call OpenRecordSet(rs, strSQL, 0)
  Do While Not rs.EOF
   Select Case i
     Case 1
        txtFia1Ced = rs!Cedula
        txtFia1Nombre = rs!Nombre
        txtFia1Deuda = rs!porc_deuda
        txtFia1Deuda.ToolTipText = "Operacion : " & IIf(IsNull(rs!Operacion), 0, rs!Operacion)
     Case 2
        txtFia2Ced = rs!Cedula
        txtFia2Nombre = rs!Nombre
        txtFia2Deuda = rs!porc_deuda
        txtFia2Deuda.ToolTipText = "Operacion : " & IIf(IsNull(rs!Operacion), 0, rs!Operacion)
     Case 3
        txtFia3Ced = rs!Cedula
        txtFia3Nombre = rs!Nombre
        txtFia3Deuda = rs!porc_deuda
        txtFia3Deuda.ToolTipText = "Operacion : " & IIf(IsNull(rs!Operacion), 0, rs!Operacion)
   End Select
   rs.MoveNext
   i = i + 1
  Loop
  rs.Close

End Sub

Private Sub rbTipo_Click(Index As Integer)
Dim strSQL As String

vPaso = True

Select Case Index
  Case 0 'Creditos
        strSQL = "select C.codigo as 'IdX', rtrim(C.Descripcion) as 'Itmx' " _
               & " from Catalogo C" _
               & " " _
               & " where C.FORMA_PAGO_POS = 1" _
               & " group by C.codigo,C.Descripcion "
        Call sbCbo_Llena_New(cboLinea, strSQL, False, True)

  Case 1 'CxC
        'Abierto Todos los Conceptos de CxC Generales
        strSQL = "select cod_concepto as 'IdX', rtrim(Descripcion) as 'Itmx' " _
               & " from CxC_Conceptos " _
               & " Where Activo = 1 and Proceso_Descuento = 0"
        Call sbCbo_Llena_New(cboLinea, strSQL, False, True)
End Select

vPaso = False

Call cboLinea_Click

End Sub

Private Sub txtFia1Ced_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtFia1Nombre.SetFocus

If KeyCode = vbKeyF4 Then
  gBusquedas.Columna = "cedula"
  gBusquedas.Orden = "cedula"
  gBusquedas.Consulta = "select cedula,nombre from pv_clientes"
  gBusquedas.Filtro = ""
  gBusquedas.Resultado = ""
  gBusquedas.Resultado2 = ""
  frmBusquedas.Show vbModal
  txtFia1Ced = gBusquedas.Resultado
  txtFia1Nombre = gBusquedas.Resultado2

End If
End Sub


Private Sub txtFia1Ced_LostFocus()
If Trim(txtFia1Ced) <> Trim(txtCedula) Then
    Call sbXFichaCliente(txtFia1Ced)
    txtFia1Nombre = fxSIFCCodigos("D", txtFia1Ced, "clientes")
Else
  MsgBox "El deudor no puede ser fiador a la vez...", vbExclamation
  txtFia1Ced.Text = ""
End If
End Sub

Private Sub txtFia2Ced_LostFocus()
If Trim(txtFia2Ced) <> Trim(txtCedula) Then
    Call sbXFichaCliente(txtFia2Ced)
    txtFia2Nombre = fxSIFCCodigos("D", txtFia2Ced, "clientes")
Else
  MsgBox "El deudor no puede ser fiador a la vez...", vbExclamation
  txtFia2Ced.Text = ""
End If
End Sub

Private Sub txtFia3Ced_LostFocus()
If Trim(txtFia3Ced) <> Trim(txtCedula) Then
    Call sbXFichaCliente(txtFia3Ced)
    txtFia3Nombre = fxSIFCCodigos("D", txtFia3Ced, "clientes")
Else
  MsgBox "El deudor no puede ser fiador a la vez...", vbExclamation
  txtFia3Ced.Text = ""
End If
End Sub


Private Sub txtFia1Nombre_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then
  gBusquedas.Columna = "nombre"
  gBusquedas.Orden = "nombre"
  gBusquedas.Consulta = "select cedula,nombre from pv_clientes"
  gBusquedas.Filtro = ""
  gBusquedas.Resultado = ""
  gBusquedas.Resultado2 = ""
  frmBusquedas.Show vbModal
  txtFia1Ced = gBusquedas.Resultado
  txtFia1Nombre = gBusquedas.Resultado2
End If
End Sub


Private Sub txtFia2Ced_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtFia2Nombre.SetFocus

If KeyCode = vbKeyF4 Then
  gBusquedas.Columna = "cedula"
  gBusquedas.Orden = "cedula"
  gBusquedas.Consulta = "select cedula,nombre from pv_clientes"
  gBusquedas.Filtro = ""
  gBusquedas.Resultado = ""
  gBusquedas.Resultado2 = ""
  frmBusquedas.Show vbModal
  txtFia2Ced = gBusquedas.Resultado
  txtFia2Nombre = gBusquedas.Resultado2
End If
End Sub


Private Sub txtFia2Nombre_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then
  gBusquedas.Columna = "nombre"
  gBusquedas.Orden = "nombre"
  gBusquedas.Consulta = "select cedula,nombre from pv_clientes"
  gBusquedas.Filtro = ""
  gBusquedas.Resultado = ""
  gBusquedas.Resultado2 = ""
  frmBusquedas.Show vbModal
  txtFia2Ced = gBusquedas.Resultado
  txtFia2Nombre = gBusquedas.Resultado2
End If
End Sub


Private Sub txtFia3Ced_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtFia3Nombre.SetFocus

If KeyCode = vbKeyF4 Then
  gBusquedas.Columna = "cedula"
  gBusquedas.Orden = "cedula"
  gBusquedas.Consulta = "select cedula,nombre from pv_clientes"
  gBusquedas.Filtro = ""
  gBusquedas.Resultado = ""
  gBusquedas.Resultado2 = ""
  frmBusquedas.Show vbModal
  txtFia3Ced = gBusquedas.Resultado
  txtFia3Nombre = gBusquedas.Resultado2
End If
End Sub


Private Sub txtFia3Nombre_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then
  gBusquedas.Columna = "nombre"
  gBusquedas.Orden = "nombre"
  gBusquedas.Consulta = "select cedula,nombre from pv_clientes"
  gBusquedas.Filtro = ""
  gBusquedas.Resultado = ""
  gBusquedas.Resultado2 = ""
  frmBusquedas.Show vbModal
  txtFia3Ced = gBusquedas.Resultado
  txtFia3Nombre = gBusquedas.Resultado2
End If
End Sub


Private Sub txtMonto_Change()
On Error GoTo vError
If CCur(IIf((txtTasa = ""), 0, txtTasa)) >= 0 And CCur(IIf((txtPlazo = ""), 0, txtPlazo)) > 0 _
    And CCur(IIf((txtMonto = ""), 0, txtMonto)) > 0 Then
 txtCuota.Text = Format(fxCalcula_Cuota(CCur(txtMonto), CCur(txtPlazo), CCur(txtTasa)), "Standard")
End If
vError:
End Sub

Private Sub txtPlazo_Change()
On Error GoTo vError
If CCur(IIf((txtTasa = ""), 0, txtTasa)) >= 0 And CCur(IIf((txtPlazo = ""), 0, txtPlazo)) > 0 _
    And CCur(IIf((txtMonto = ""), 0, txtMonto)) > 0 Then
 txtCuota.Text = Format(fxCalcula_Cuota(CCur(txtMonto), CCur(txtPlazo), CCur(txtTasa)), "Standard")
End If
vError:

End Sub

Private Sub txtTasa_Change()
On Error GoTo vError
If CCur(IIf((txtTasa = ""), 0, txtTasa)) >= 0 And CCur(IIf((txtPlazo = ""), 0, txtPlazo)) > 0 _
    And CCur(IIf((txtMonto = ""), 0, txtMonto)) > 0 Then
 txtCuota.Text = Format(fxCalcula_Cuota(CCur(txtMonto), CCur(txtPlazo), CCur(txtTasa)), "Standard")
End If
vError:

End Sub
