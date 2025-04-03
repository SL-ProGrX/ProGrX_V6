VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "Codejock.Controls.v22.1.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "Codejock.ShortcutBar.v22.1.0.ocx"
Object = "{C215CB9A-0AE1-499F-A101-48B3C370D3DF}#22.1#0"; "Codejock.ChartControl.v22.1.0.ocx"
Begin VB.Form frmDSB_Asociados 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Dashboard para Asociados"
   ClientHeight    =   11580
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   17295
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11580
   ScaleWidth      =   17295
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.GroupBox gbScoring 
      Height          =   4335
      Left            =   5760
      TabIndex        =   30
      Top             =   7200
      Width           =   5655
      _Version        =   1441793
      _ExtentX        =   9975
      _ExtentY        =   7646
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      BorderStyle     =   2
      Begin XtremeSuiteControls.FlatEdit txtSc_Morosidad 
         Height          =   375
         Left            =   720
         TabIndex        =   36
         ToolTipText     =   "Operaciones"
         Top             =   1080
         Width           =   975
         _Version        =   1441793
         _ExtentX        =   1720
         _ExtentY        =   661
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   12648447
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   "AA"
         BackColor       =   12648447
         Alignment       =   2
         Locked          =   -1  'True
         Appearance      =   17
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtSc_Endeudamiento 
         Height          =   375
         Left            =   720
         TabIndex        =   37
         ToolTipText     =   "Operaciones"
         Top             =   1560
         Width           =   975
         _Version        =   1441793
         _ExtentX        =   1720
         _ExtentY        =   661
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   12648447
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   "E1"
         BackColor       =   12648447
         Alignment       =   2
         Locked          =   -1  'True
         Appearance      =   17
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtSc_Liquidez 
         Height          =   375
         Left            =   720
         TabIndex        =   38
         ToolTipText     =   "Operaciones"
         Top             =   2040
         Width           =   975
         _Version        =   1441793
         _ExtentX        =   1720
         _ExtentY        =   661
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   12648447
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   "L1"
         BackColor       =   12648447
         Alignment       =   2
         Locked          =   -1  'True
         Appearance      =   17
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtSc_Historial_Pago 
         Height          =   375
         Left            =   720
         TabIndex        =   39
         ToolTipText     =   "Operaciones"
         Top             =   2520
         Width           =   975
         _Version        =   1441793
         _ExtentX        =   1720
         _ExtentY        =   661
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   12648447
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   "H1"
         BackColor       =   12648447
         Alignment       =   2
         Locked          =   -1  'True
         Appearance      =   17
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtSC_Salario_Devengado 
         Height          =   375
         Left            =   240
         TabIndex        =   53
         Top             =   3480
         Width           =   1455
         _Version        =   1441793
         _ExtentX        =   2566
         _ExtentY        =   661
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16777215
         Alignment       =   1
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtSC_Salario_Liquido 
         Height          =   375
         Left            =   240
         TabIndex        =   55
         Top             =   3960
         Width           =   1455
         _Version        =   1441793
         _ExtentX        =   2566
         _ExtentY        =   661
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16777215
         Alignment       =   1
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtSC_Garantia 
         Height          =   375
         Left            =   720
         TabIndex        =   62
         ToolTipText     =   "Operaciones"
         Top             =   3000
         Width           =   975
         _Version        =   1441793
         _ExtentX        =   1720
         _ExtentY        =   661
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   12648447
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   "G1"
         BackColor       =   12648447
         Alignment       =   2
         Locked          =   -1  'True
         Appearance      =   17
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtSC_Fecha 
         Height          =   375
         Left            =   240
         TabIndex        =   51
         Top             =   480
         Width           =   1455
         _Version        =   1441793
         _ExtentX        =   2566
         _ExtentY        =   661
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16777215
         Alignment       =   2
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.Label Label1 
         Height          =   255
         Index           =   7
         Left            =   2160
         TabIndex        =   63
         Top             =   3000
         Width           =   2295
         _Version        =   1441793
         _ExtentX        =   4048
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Calidad de la Garantía"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin XtremeSuiteControls.Label Label9 
         Height          =   255
         Left            =   2160
         TabIndex        =   54
         Top             =   3960
         Width           =   4095
         _Version        =   1441793
         _ExtentX        =   7223
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Salario Liquido"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin XtremeSuiteControls.Label Label8 
         Height          =   255
         Left            =   2160
         TabIndex        =   52
         Top             =   3480
         Width           =   4095
         _Version        =   1441793
         _ExtentX        =   7223
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Salario Devengado"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin XtremeSuiteControls.Label Label7 
         Height          =   255
         Left            =   2160
         TabIndex        =   50
         Top             =   480
         Width           =   4095
         _Version        =   1441793
         _ExtentX        =   7223
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Fecha de Estudio"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin XtremeSuiteControls.Label Label1 
         Height          =   495
         Index           =   5
         Left            =   120
         TabIndex        =   40
         Top             =   -120
         Width           =   4215
         _Version        =   1441793
         _ExtentX        =   7435
         _ExtentY        =   873
         _StockProps     =   79
         Caption         =   "Calificación Créditicia"
         ForeColor       =   4210752
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin XtremeSuiteControls.Label Label1 
         Height          =   375
         Index           =   4
         Left            =   2160
         TabIndex        =   35
         Top             =   2520
         Width           =   2295
         _Version        =   1441793
         _ExtentX        =   4048
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Historial de Pago"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin XtremeSuiteControls.Label Label1 
         Height          =   375
         Index           =   3
         Left            =   2160
         TabIndex        =   34
         Top             =   2040
         Width           =   2655
         _Version        =   1441793
         _ExtentX        =   4683
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Liquidez / Capacidad de Pago"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin XtremeSuiteControls.Label Label1 
         Height          =   255
         Index           =   2
         Left            =   2160
         TabIndex        =   33
         Top             =   1080
         Width           =   2295
         _Version        =   1441793
         _ExtentX        =   4048
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Morosidad"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin XtremeSuiteControls.Label Label1 
         Height          =   375
         Index           =   1
         Left            =   2160
         TabIndex        =   32
         Top             =   1560
         Width           =   2295
         _Version        =   1441793
         _ExtentX        =   4048
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Nivel de Endeudamiento"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin XtremeSuiteControls.GroupBox gbResumen 
      Height          =   2655
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   17295
      _Version        =   1441793
      _ExtentX        =   30506
      _ExtentY        =   4683
      _StockProps     =   79
      BackColor       =   16777215
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      BorderStyle     =   2
      Begin VB.Timer Timer_Access 
         Interval        =   5
         Left            =   15120
         Top             =   120
      End
      Begin XtremeSuiteControls.FlatEdit txtPlan_Aportes 
         Height          =   375
         Left            =   11760
         TabIndex        =   23
         ToolTipText     =   "Principal"
         Top             =   2160
         Width           =   2415
         _Version        =   1441793
         _ExtentX        =   4260
         _ExtentY        =   661
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   16777152
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   "0.00"
         BackColor       =   16777152
         Alignment       =   1
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtRET_Saldos 
         Height          =   375
         Left            =   11760
         TabIndex        =   19
         ToolTipText     =   "Principal"
         Top             =   1680
         Width           =   2415
         _Version        =   1441793
         _ExtentX        =   4260
         _ExtentY        =   661
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   16777152
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   "0.00"
         BackColor       =   16777152
         Alignment       =   1
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtNombre 
         Height          =   435
         Left            =   5520
         TabIndex        =   1
         Top             =   120
         Width           =   7335
         _Version        =   1441793
         _ExtentX        =   12938
         _ExtentY        =   767
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
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtCedula 
         Height          =   435
         Left            =   3240
         TabIndex        =   2
         Top             =   120
         Width           =   2295
         _Version        =   1441793
         _ExtentX        =   4048
         _ExtentY        =   767
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
      Begin MSComctlLib.ImageList imgSemaforos 
         Left            =   2160
         Top             =   1680
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   15
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDSB_Asociados.frx":0000
               Key             =   "verde"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDSB_Asociados.frx":061C
               Key             =   "amarillo"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDSB_Asociados.frx":0C3A
               Key             =   "rojo"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDSB_Asociados.frx":1321
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDSB_Asociados.frx":1BF2
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDSB_Asociados.frx":2319
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDSB_Asociados.frx":293D
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDSB_Asociados.frx":31EA
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDSB_Asociados.frx":38F6
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDSB_Asociados.frx":3FFD
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDSB_Asociados.frx":46FD
               Key             =   ""
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDSB_Asociados.frx":4D19
               Key             =   ""
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDSB_Asociados.frx":544A
               Key             =   ""
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDSB_Asociados.frx":5B47
               Key             =   ""
            EndProperty
            BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDSB_Asociados.frx":624E
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin XtremeSuiteControls.PushButton btnConsulta 
         Height          =   450
         Left            =   12960
         TabIndex        =   11
         Top             =   120
         Width           =   855
         _Version        =   1441793
         _ExtentX        =   1508
         _ExtentY        =   794
         _StockProps     =   79
         Caption         =   "Integral"
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
      Begin XtremeSuiteControls.FlatEdit txtCRD_Saldos 
         Height          =   375
         Left            =   11760
         TabIndex        =   15
         ToolTipText     =   "Principal"
         Top             =   1200
         Width           =   2415
         _Version        =   1441793
         _ExtentX        =   4260
         _ExtentY        =   661
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   16777152
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   "0.00"
         BackColor       =   16777152
         Alignment       =   1
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtCRD_Cuotas 
         Height          =   375
         Left            =   14160
         TabIndex        =   16
         ToolTipText     =   "Cuotas + Seguros"
         Top             =   1200
         Width           =   1815
         _Version        =   1441793
         _ExtentX        =   3201
         _ExtentY        =   661
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   16777088
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   "0.00"
         BackColor       =   16777088
         Alignment       =   1
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtCRD_Operaciones 
         Height          =   375
         Left            =   15960
         TabIndex        =   17
         ToolTipText     =   "Operaciones"
         Top             =   1200
         Width           =   1215
         _Version        =   1441793
         _ExtentX        =   2143
         _ExtentY        =   661
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   16776960
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   "0"
         BackColor       =   16776960
         Alignment       =   1
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtRET_Cuotas 
         Height          =   375
         Left            =   14160
         TabIndex        =   20
         ToolTipText     =   "Cuotas + Seguros"
         Top             =   1680
         Width           =   1815
         _Version        =   1441793
         _ExtentX        =   3201
         _ExtentY        =   661
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   16777088
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   "0.00"
         BackColor       =   16777088
         Alignment       =   1
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtRET_Operaciones 
         Height          =   375
         Left            =   15960
         TabIndex        =   21
         ToolTipText     =   "Operaciones"
         Top             =   1680
         Width           =   1215
         _Version        =   1441793
         _ExtentX        =   2143
         _ExtentY        =   661
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   16776960
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   "0"
         BackColor       =   16776960
         Alignment       =   1
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtPlan_Mensualidad 
         Height          =   375
         Left            =   14160
         TabIndex        =   24
         ToolTipText     =   "Cuotas + Seguros"
         Top             =   2160
         Width           =   1815
         _Version        =   1441793
         _ExtentX        =   3201
         _ExtentY        =   661
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   16777088
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   "0.00"
         BackColor       =   16777088
         Alignment       =   1
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtPlan_Contratos 
         Height          =   375
         Left            =   15960
         TabIndex        =   25
         ToolTipText     =   "Operaciones"
         Top             =   2160
         Width           =   1215
         _Version        =   1441793
         _ExtentX        =   2143
         _ExtentY        =   661
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   16776960
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   "0"
         BackColor       =   16776960
         Alignment       =   1
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.PushButton btnAccess 
         Height          =   450
         Left            =   14160
         TabIndex        =   64
         Top             =   120
         Visible         =   0   'False
         Width           =   855
         _Version        =   1441793
         _ExtentX        =   1508
         _ExtentY        =   794
         _StockProps     =   79
         Caption         =   "Access"
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
      Begin XtremeSuiteControls.Label Label3 
         Height          =   375
         Left            =   9720
         TabIndex        =   22
         Top             =   2160
         Width           =   2055
         _Version        =   1441793
         _ExtentX        =   3625
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Ahorros"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin XtremeSuiteControls.Label Label2 
         Height          =   375
         Left            =   9720
         TabIndex        =   18
         Top             =   1680
         Width           =   2055
         _Version        =   1441793
         _ExtentX        =   3625
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Recaudos"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin XtremeSuiteControls.Label Label1 
         Height          =   375
         Index           =   0
         Left            =   9720
         TabIndex        =   14
         Top             =   1200
         Width           =   2055
         _Version        =   1441793
         _ExtentX        =   3625
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Créditos"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
         Height          =   375
         Index           =   0
         Left            =   0
         TabIndex        =   12
         Top             =   720
         Width           =   9615
         _Version        =   1441793
         _ExtentX        =   16960
         _ExtentY        =   661
         _StockProps     =   14
         Caption         =   "Estado de la Persona:"
         ForeColor       =   4210752
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SubItemCaption  =   -1  'True
         VisualTheme     =   3
         ForeColor       =   4210752
      End
      Begin VB.Image imgInstitucion 
         Height          =   240
         Left            =   3840
         Picture         =   "frmDSB_Asociados.frx":6964
         ToolTipText     =   "Empresa / Deductora?"
         Top             =   1500
         Width           =   240
      End
      Begin VB.Label lblInstitución 
         BackStyle       =   0  'Transparent
         Caption         =   "Empresa / Deductora?"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   4320
         TabIndex        =   10
         Top             =   1500
         Width           =   5175
      End
      Begin VB.Image imgEstado 
         Height          =   240
         Left            =   3840
         Picture         =   "frmDSB_Asociados.frx":6A95
         ToolTipText     =   "Estado de la persona"
         Top             =   1200
         Width           =   240
      End
      Begin VB.Label lblEstado 
         BackStyle       =   0  'Transparent
         Caption         =   "Estado ?"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   4320
         TabIndex        =   9
         Top             =   1200
         Width           =   2655
      End
      Begin VB.Label lblCreditos 
         BackStyle       =   0  'Transparent
         Caption         =   "Estatus Créditos ?"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   600
         TabIndex        =   8
         ToolTipText     =   "Estado General de los Créditos"
         Top             =   1830
         Width           =   2535
      End
      Begin VB.Label lblFianzas 
         BackStyle       =   0  'Transparent
         Caption         =   "Estatus Fianzas ?"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   600
         TabIndex        =   7
         ToolTipText     =   "Estado de las Fianzas"
         Top             =   2130
         Width           =   2415
      End
      Begin VB.Label lblMembresia 
         BackStyle       =   0  'Transparent
         Caption         =   "Membresía ?"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   4320
         TabIndex        =   6
         Top             =   1815
         Width           =   5175
      End
      Begin VB.Image imgCreditos 
         Height          =   255
         Left            =   240
         Picture         =   "frmDSB_Asociados.frx":6BAC
         Stretch         =   -1  'True
         ToolTipText     =   "Estado de los creditos"
         Top             =   1815
         Width           =   255
      End
      Begin VB.Image imgFianzas 
         Height          =   255
         Left            =   240
         Picture         =   "frmDSB_Asociados.frx":737B
         Stretch         =   -1  'True
         ToolTipText     =   "Estado de las fianzas"
         Top             =   2115
         Width           =   255
      End
      Begin VB.Image imgMembresia 
         Height          =   255
         Left            =   3840
         Picture         =   "frmDSB_Asociados.frx":7B4A
         Stretch         =   -1  'True
         ToolTipText     =   "Membresía de la Persona"
         Top             =   1815
         Width           =   255
      End
      Begin VB.Label lblClasificacion 
         BackStyle       =   0  'Transparent
         Caption         =   "Clasificación ?"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   4320
         TabIndex        =   5
         Top             =   2160
         Width           =   5175
      End
      Begin VB.Image imgClasificacion 
         Height          =   255
         Left            =   3840
         Picture         =   "frmDSB_Asociados.frx":8222
         Stretch         =   -1  'True
         ToolTipText     =   "Clasificacion ABCD de la persona"
         Top             =   2160
         Width           =   255
      End
      Begin VB.Image imgEstadoBeneficiarios 
         Height          =   255
         Left            =   240
         Picture         =   "frmDSB_Asociados.frx":89A2
         Stretch         =   -1  'True
         ToolTipText     =   "Estado de Actualización de los Beneficiarios"
         Top             =   1200
         Width           =   255
      End
      Begin VB.Image imgEstadoConsentimiento 
         Height          =   255
         Left            =   240
         Picture         =   "frmDSB_Asociados.frx":9171
         Stretch         =   -1  'True
         ToolTipText     =   "Estado de Autorización de Uso de Información Personal"
         Top             =   1500
         Width           =   255
      End
      Begin VB.Label lblEstadoBeneficiarios 
         BackStyle       =   0  'Transparent
         Caption         =   "Estatus Beneficiarios?"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   600
         TabIndex        =   4
         ToolTipText     =   "Estados de Actualización de los Beneficiarios"
         Top             =   1200
         Width           =   2535
      End
      Begin VB.Label lblEstadoAutInfo 
         BackStyle       =   0  'Transparent
         Caption         =   "Estatus Aut.Info.?"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   600
         TabIndex        =   3
         ToolTipText     =   "Estado de la Autorizacipon de Uso de la Información"
         Top             =   1500
         Width           =   2415
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
         Height          =   375
         Index           =   1
         Left            =   9600
         TabIndex        =   13
         Top             =   720
         Width           =   7695
         _Version        =   1441793
         _ExtentX        =   13573
         _ExtentY        =   661
         _StockProps     =   14
         Caption         =   "Resumen:"
         ForeColor       =   4210752
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SubItemCaption  =   -1  'True
         VisualTheme     =   3
         ForeColor       =   4210752
      End
   End
   Begin XtremeChartControl.ChartControl ccPAT 
      Height          =   4335
      Left            =   0
      TabIndex        =   26
      Top             =   2760
      Width           =   5655
      _Version        =   1441793
      _ExtentX        =   9975
      _ExtentY        =   7646
      _StockProps     =   0
   End
   Begin XtremeChartControl.ChartControl ccCreditos 
      Height          =   4335
      Left            =   5760
      TabIndex        =   27
      Top             =   2760
      Width           =   5655
      _Version        =   1441793
      _ExtentX        =   9975
      _ExtentY        =   7646
      _StockProps     =   0
   End
   Begin XtremeChartControl.ChartControl ccPlanes 
      Height          =   4335
      Left            =   11520
      TabIndex        =   28
      Top             =   2760
      Width           =   5655
      _Version        =   1441793
      _ExtentX        =   9975
      _ExtentY        =   7646
      _StockProps     =   0
   End
   Begin XtremeChartControl.ChartControl ccBeneficios 
      Height          =   4335
      Left            =   0
      TabIndex        =   29
      Top             =   7200
      Width           =   5655
      _Version        =   1441793
      _ExtentX        =   9975
      _ExtentY        =   7646
      _StockProps     =   0
   End
   Begin XtremeSuiteControls.GroupBox gbEstudio 
      Height          =   4335
      Left            =   11520
      TabIndex        =   31
      Top             =   7200
      Width           =   5655
      _Version        =   1441793
      _ExtentX        =   9975
      _ExtentY        =   7646
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      BorderStyle     =   2
      Begin XtremeSuiteControls.FlatEdit txtU_Liquidacion 
         Height          =   375
         Left            =   360
         TabIndex        =   45
         Top             =   3600
         Width           =   1455
         _Version        =   1441793
         _ExtentX        =   2566
         _ExtentY        =   661
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16777215
         Alignment       =   2
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtU_Credito 
         Height          =   375
         Left            =   360
         TabIndex        =   47
         Top             =   2160
         Width           =   1455
         _Version        =   1441793
         _ExtentX        =   2566
         _ExtentY        =   661
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16777215
         Alignment       =   2
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtU_Beneficio 
         Height          =   375
         Left            =   360
         TabIndex        =   49
         Top             =   3120
         Width           =   1455
         _Version        =   1441793
         _ExtentX        =   2566
         _ExtentY        =   661
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16777215
         Alignment       =   2
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtU_Ahorro 
         Height          =   375
         Left            =   360
         TabIndex        =   57
         Top             =   2640
         Width           =   1455
         _Version        =   1441793
         _ExtentX        =   2566
         _ExtentY        =   661
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16777215
         Alignment       =   2
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtDisponible_SA 
         Height          =   375
         Left            =   360
         TabIndex        =   61
         Top             =   1080
         Width           =   1455
         _Version        =   1441793
         _ExtentX        =   2566
         _ExtentY        =   661
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16777215
         Alignment       =   1
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtDisponible_Exc 
         Height          =   375
         Left            =   360
         TabIndex        =   59
         Top             =   1560
         Width           =   1455
         _Version        =   1441793
         _ExtentX        =   2566
         _ExtentY        =   661
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16777215
         Alignment       =   1
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.Label Label12 
         Height          =   255
         Left            =   2040
         TabIndex        =   60
         Top             =   1080
         Width           =   4095
         _Version        =   1441793
         _ExtentX        =   7223
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Disponible Sobre Ahorros"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin XtremeSuiteControls.Label Label11 
         Height          =   255
         Left            =   2040
         TabIndex        =   58
         Top             =   1560
         Width           =   4095
         _Version        =   1441793
         _ExtentX        =   7223
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Disponible Excedentes"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin XtremeSuiteControls.Label Label10 
         Height          =   255
         Left            =   2040
         TabIndex        =   56
         Top             =   2640
         Width           =   4095
         _Version        =   1441793
         _ExtentX        =   7223
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Ultimo Ahorro"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin XtremeSuiteControls.Label Label6 
         Height          =   255
         Left            =   2040
         TabIndex        =   48
         Top             =   3120
         Width           =   4095
         _Version        =   1441793
         _ExtentX        =   7223
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Ultimo Beneficio"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin XtremeSuiteControls.Label Label5 
         Height          =   255
         Left            =   2040
         TabIndex        =   46
         Top             =   2160
         Width           =   4095
         _Version        =   1441793
         _ExtentX        =   7223
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Ultimo Crédito"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin XtremeSuiteControls.Label Label4 
         Height          =   255
         Left            =   2040
         TabIndex        =   44
         Top             =   3600
         Width           =   4095
         _Version        =   1441793
         _ExtentX        =   7223
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Ultima Liquidación"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin XtremeSuiteControls.Label lblIG_Generacion 
         Height          =   255
         Left            =   2640
         TabIndex        =   43
         Top             =   480
         Width           =   4095
         _Version        =   1441793
         _ExtentX        =   7223
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Generación X"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Image imgIG 
         Height          =   255
         Index           =   0
         Left            =   360
         Picture         =   "frmDSB_Asociados.frx":9940
         Stretch         =   -1  'True
         ToolTipText     =   "Estado de Actualización de los Beneficiarios"
         Top             =   480
         Width           =   255
      End
      Begin XtremeSuiteControls.Label lblIG_Edad 
         Height          =   255
         Left            =   1080
         TabIndex        =   42
         Top             =   480
         Width           =   4095
         _Version        =   1441793
         _ExtentX        =   7223
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Edad: 18 años"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin XtremeSuiteControls.Label Label1 
         Height          =   495
         Index           =   6
         Left            =   120
         TabIndex        =   41
         Top             =   -120
         Width           =   4215
         _Version        =   1441793
         _ExtentX        =   7435
         _ExtentY        =   873
         _StockProps     =   79
         Caption         =   "Información General"
         ForeColor       =   4210752
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
End
Attribute VB_Name = "frmDSB_Asociados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim vPaso As Boolean
Dim vRA_Access As Boolean

Dim strSQL As String, rs As New ADODB.Recordset


Dim mChartPallete As String, mChartLabelPosition As Integer
Dim Diagram As ChartDiagram2D
Dim Strip As ChartAxisStrip


Sub CreateSeriesPoint(ByVal pPointCollection As ChartSeriesPointCollection, vArg As String, nValue As Double)
    Dim pPoint As ChartSeriesPoint
    Set pPoint = pPointCollection.Add(vArg, nValue)
     pPoint.LabelText = Format(nValue, "Standard")
     pPoint.LegendText = vArg

End Sub

Public Sub sbChart_3D(pChart As Object, pTitulo As String, Optional pTema As String = "" _
                        , Optional p3dTipo = "3d_Pie", Optional pExpresado As Long = 1 _
                        , Optional Pattern As String = "N" _
                        , Optional pDecimals As Integer = 0)

On Error GoTo vError
    
    If pChart.Content.Series.Count > 0 Then
        pChart.Content.Series.DeleteAll
    End If
    
    Dim Series As ChartSeries
        
        
    pChart.Content.Titles.DeleteAll
    pChart.Content.Titles.Add pTitulo
    
    pChart.Content.Legend.Visible = True
    pChart.Content.Legend.HorizontalAlignment = xtpChartLegendFarOutside
    
    Set Series = pChart.Content.Series.Add(pTema)
    
    
    Call OpenRecordSet(rs, strSQL)
               
    Dim i As Integer, x As Integer, C As Currency

    C = 0
    i = 0
    x = 0

    Do While Not rs.EOF
      C = rs!Value / pExpresado

      CreateSeriesPoint Series.Points, rs!Descripcion, rs!Value / pExpresado

      If rs!Value > C Then
        C = rs!Value / pExpresado
        x = i
      End If

      i = i + 1
      rs.MoveNext
    Loop
    rs.Close

    Series.Points(x).Special = True
                
                
                
    pChart.Content.Appearance.SetPalette mChartPallete
   ' pChart.Content.Series(0).Style.Label.Position mChartLabelPosition
   Select Case p3dTipo
    Case "Pie"
                        
            Dim PieStyle As ChartPieSeriesStyle
            Set PieStyle = New ChartPieSeriesStyle
            Set pChart.Content.Series(0).Style = PieStyle
            
'            PieStyle.Label.Format.Pattern = "{V} Million"
            PieStyle.Label.Format.Category = xtpChartNumber
                       
            'PieStyle.HolePercent = 40
            PieStyle.Rotation = 30
            PieStyle.Label.Visible = True
            'PieStyle.Label.ShowLines = False
            'PieStyle.ExplodedDistancePercent = 30
            
            'cmbPieLabelPosition.ListIndex = PieStyle.Label.Position
            pChart.Content.Series(0).Style.Label.Position = 1
            PieStyle.Label.Antialiasing = False
            
    Case "Pyramid"
            
            Dim PyramidStyle As ChartPyramidSeriesStyle
            Set PyramidStyle = New ChartPyramidSeriesStyle
            Set pChart.Content.Series(0).Style = PyramidStyle
            
            'PyramidStyle.Label.Format.Pattern = "{V} Million"
            PyramidStyle.Label.Format.Category = xtpChartNumber
                       
            PyramidStyle.HeightToWidthRatio = 1.25
            PyramidStyle.PointDistance = 5
            PyramidStyle.Label.Visible = True
            
            'cmbLabelPosition.ListIndex = PyramidStyle.Label.Position
            PyramidStyle.Label.Antialiasing = False
    
    Case "3d_Pie"
                  
            Dim Style3dPie As Chart3dPieSeriesStyle
            
            Set Style3dPie = New Chart3dPieSeriesStyle
            Set pChart.Content.Series(0).Style = Style3dPie

            Style3dPie.Label.Format.Category = xtpChartNumber
                       
            Dim Pie3dRotation As Chart3dRotation
            Set Pie3dRotation = New Chart3dRotation
            
            Pie3dRotation.Yaw = 60
            Pie3dRotation.Pitch = 40
            Pie3dRotation.Roll = 30
            Style3dPie.SetRotation Pie3dRotation
            Style3dPie.Label.Visible = True
            
            Style3dPie.Label.Antialiasing = False
     
     Case "3d_Doughnut"
            Dim Style3dDoughnut As Chart3dPieSeriesStyle
            Set Style3dDoughnut = New Chart3dPieSeriesStyle
            
            Set pChart.Content.Series(0).Style = Style3dDoughnut
            Style3dDoughnut.Label.Format.Category = xtpChartNumber
                       
            Dim Doughnut3dRotation As Chart3dRotation
            Set Doughnut3dRotation = New Chart3dRotation
            Doughnut3dRotation.Yaw = 10
            Doughnut3dRotation.Pitch = 20
            Doughnut3dRotation.Roll = 50
            Style3dDoughnut.SetRotation Doughnut3dRotation
            Style3dDoughnut.Label.Visible = True
            Style3dDoughnut.HolePercent = 60
            Style3dDoughnut.ExplodedDistancePercent = 20
     
     Case "3d_Pyramid"
            Dim Pyramid3dStyle As Chart3dPyramidSeriesStyle
            Set Pyramid3dStyle = New Chart3dPyramidSeriesStyle
            Set pChart.Content.Series(0).Style = Pyramid3dStyle
            
            Pyramid3dStyle.Label.Format.Category = xtpChartNumber
                       
            Dim Pyramid3dRotation As Chart3dRotation
            Set Pyramid3dRotation = New Chart3dRotation
            Pyramid3dRotation.Yaw = 70
            Pyramid3dRotation.Pitch = 20
            Pyramid3dRotation.Roll = 70
            Pyramid3dStyle.SetRotation Pyramid3dRotation
            
            Pyramid3dStyle.HeightToWidthRatio = 2
            Pyramid3dStyle.PointDistance = 2
            Pyramid3dStyle.BaseEdgeCount = 7
            Pyramid3dStyle.SmoothEdges = True
            Pyramid3dStyle.Label.Visible = True
            
     Case "3d_Torus"
            Dim Style3dTorus As Chart3dPieSeriesStyle
            Set Style3dTorus = New Chart3dPieSeriesStyle
            
            Set pChart.Content.Series(0).Style = Style3dTorus
            
            Style3dTorus.Label.Format.Category = xtpChartNumber
                       
            Dim Torus3dRotation As Chart3dRotation
            Set Torus3dRotation = New Chart3dRotation
            Torus3dRotation.Yaw = -20
            Torus3dRotation.Pitch = 0
            Torus3dRotation.Roll = 70
            Style3dTorus.SetRotation Torus3dRotation
            Style3dTorus.Label.Visible = True
            Style3dTorus.IsTorus = True
            Style3dTorus.Depth = Style3dTorus.Depth * 2
            
            Style3dTorus.Label.Antialiasing = False
            
      Case "3d_Funnel"
            'AddFunnelSeries
                        
            Dim Funnel3dStyle As Chart3dFunnelSeriesStyle
            Set Funnel3dStyle = New Chart3dFunnelSeriesStyle
            
            Set pChart.Content.Series(0).Style = Funnel3dStyle
            
            Funnel3dStyle.Label.Format.Category = xtpChartNumber
                       
            Dim Funnel3dRotation As Chart3dRotation
            Set Funnel3dRotation = New Chart3dRotation
            Funnel3dRotation.Yaw = 203
            Funnel3dRotation.Pitch = 355
            Funnel3dRotation.Roll = 79
            Funnel3dStyle.SetRotation Funnel3dRotation
            
            Funnel3dStyle.HeightToWidthRatio = 1.5
            Funnel3dStyle.BaseEdgeCount = 4
            Funnel3dStyle.SmoothEdges = True
            Funnel3dStyle.Label.Visible = True
            
            Funnel3dStyle.Label.Antialiasing = False
    
    
    End Select
                
                
                
    Select Case Pattern
      Case "N" 'Numerico
'            LineStyle.Label.Format.DecimalPlaces = pDecimals
'            LineStyle.Label.Format.UseThousandSeparator = True
'            LineStyle.Label.Format.Category = xtpChartNumber
       
            Set Diagram = pChart.Content.Series(0).Diagram
            Select Case pExpresado
                Case 1
                Case 1000
                    Diagram.AxisY.Title = "Monto en (miles)"
                Case 1000000
                    Diagram.AxisY.Title = "Monto en (millones)"
            End Select
            Diagram.AxisY.Title.Visible = True
            Diagram.AxisX.Title = "Meses"
            Diagram.AxisX.Title.Visible = True
            
            Diagram.AxisY.Label.Format.Category = xtpChartNumber
            Diagram.AxisY.Label.Format.DecimalPlaces = 0
       
       
       Case "P"
'            LineStyle.Label.Format.Category = xtpChartPercentage
       
'            Set Diagram = pChart.Content.Series(0).Diagram
'            Select Case pExpresado
'                Case 1
'                Case 1000
'                    Diagram.AxisY.Title = "Monto en (miles)"
'                Case 1000000
'                    Diagram.AxisY.Title = "Monto en (millones)"
'            End Select
'            Diagram.AxisY.Title.Visible = True
'            Diagram.AxisX.Title = "Meses"
'            Diagram.AxisX.Title.Visible = True
'
            Diagram.AxisY.Label.Format.Category = xtpChartPercentage
            Diagram.AxisY.Label.Format.DecimalPlaces = 0
       
       
    End Select
                
    
Exit Sub
    
vError:
    
End Sub




Private Sub EstadoInicial()
On Error Resume Next

Call Limpia

    txtCedula.Enabled = True
    txtCedula.SetFocus

End Sub

Private Sub Limpia()

 txtCedula.Text = ""
 
 lblFianzas.Caption = "Estado de Fianzas?"
 lblCreditos.Caption = "Estado de Créditos?"
 lblMembresia.Caption = "Membresía?"
 lblMembresia.ToolTipText = ""
  
 lblEstado.Caption = "Estado Persona?"
 lblInstitución.Caption = "Empresa/Deductora?"
 lblInstitución.ToolTipText = ""
 
 lblClasificacion.Caption = "Clasificación?"
  
  
End Sub

Private Function fxLiquidacion(vCedula As String) As String

On Error GoTo vError

Dim pResultado As String

pResultado = ""

glogon.strSQL = "select C.descripcion" _
       & " from liquidacion L inner join Causas_Renuncias C on C.id_causa = L.id_causa" _
       & " where consec in(select max(consec) from liquidacion" _
       & " where cedula = '" & vCedula & "')"
Call OpenRecordSet(glogon.Recordset, glogon.strSQL)
If Not rs.EOF And Not rs.BOF Then
 pResultado = "[CAUSA: " & Trim(glogon.Recordset!Descripcion) & "]"
End If

glogon.Recordset.Close

fxLiquidacion = pResultado

Exit Function

vError:
    fxLiquidacion = pResultado

End Function

Private Sub sbSemaforo(Obj As Object, Color As String)

Select Case LCase(Color)
    Case "rojo"
         Obj.BackColor = RGB(249, 235, 234)
    Case "verde"
         Obj.BackColor = RGB(233, 247, 239)
    Case "amarillo"
        Obj.BackColor = RGB(252, 243, 207)
    Case "blanco"
        Obj.BackColor = vbWhite
 End Select

End Sub


Private Sub sbConsulta(pCedula As String)
Dim vFechaIng As Date, vFianzas As Boolean
Dim rsTmp As New ADODB.Recordset, i As Integer
     
  
vFianzas = False
  
If Not fxSIFValidaCadena(txtCedula.Text) Then
   Exit Sub
End If
  
'Valida Acceso a Expediente
vRA_Access = fxSys_RA_Consulta(Trim(pCedula), glogon.Usuario)
 
If Not vRA_Access Then
    MsgBox "Esta persona se encuentra con -> Expediente Restringido <- Requiere de Autorización para Consultar!", vbExclamation
    txtCedula.Text = ""
    txtNombre.Text = ""
    Exit Sub
End If
  
strSQL = "exec spDSB_Persona_Consulta '" & Trim(pCedula) & "'"
Call OpenRecordSet(rs, strSQL)
 
If Not rs.EOF And Not rs.BOF Then
   
   txtCedula.Text = Trim(rs!Cedulax & "")
   txtNombre.Text = rs!Nombre & ""
   

   lblEstado.Caption = "Estado : " & rs!EstadoX
   
   lblInstitución.Caption = rs!InstitucionX
   lblInstitución.ToolTipText = "Deductora: " & rs!Deductora
    
     
     vFechaIng = IIf(IsNull(rs!FechaIngreso), fxFechaServidor, rs!FechaIngreso)
     
     lblMembresia.FontBold = False
     
     If rs!EstadoActual = "S" Then
        lblMembresia.Caption = "Membresía: " & rs!Membresia   'fxMembresia(vFechaIng)
        lblMembresia.ToolTipText = "[Ing.:" & Format(vFechaIng, "dd/mm/yyyy") & "]"

               
        strSQL = "exec spAFI_ConsultaRenunciaTransito '" & pCedula & "'"
        Call OpenRecordSet(rsTmp, strSQL, 0)
        If Not rsTmp.EOF And Not rsTmp.BOF Then
            lblMembresia.Caption = "Renuncia: " & rsTmp!Cod_Renuncia & " ¦ " & rsTmp!REGISTRO_FECHA & " ¦ " & rsTmp!registro_user
            lblMembresia.ToolTipText = rsTmp!Estado & " ¦ " & rsTmp!Tipo & " ¦ " & Trim(rsTmp!Descripcion)
            lblMembresia.ForeColor = vbRed
            lblMembresia.FontBold = True
        End If
        rsTmp.Close
     
     Else
        lblMembresia.Caption = "Membresía: NADA"
        lblMembresia.ToolTipText = fxLiquidacion(rs!Cedulax)
     End If
     
     'Clasificación de la Persona
     lblClasificacion.Caption = "Clasificación Crediticia : [" & rs!Clasificacion & "]"
     lblClasificacion.Caption = rs!Mora_Antiguedad & ""
    
    'Indica el Estado de las Fianzas
     If rs!IndFianzas = 0 Then
       vFianzas = False
       lblFianzas.Caption = "Fianzas al Día"
       Set imgFianzas.Picture = imgSemaforos.ListImages.Item(1).Picture
     Else
       vFianzas = True
       lblFianzas.Caption = "Fianzas en Mora"
       Set imgFianzas.Picture = imgSemaforos.ListImages.Item(3).Picture
     End If
     
     


     'Indicador de Estado de Beneficiarios
     Select Case rs!IndBeneficiarios
       Case 0 'Rojo
        Set imgEstadoBeneficiarios.Picture = imgSemaforos.ListImages.Item(3).Picture
       Case 1 'Verde
        Set imgEstadoBeneficiarios.Picture = imgSemaforos.ListImages.Item(1).Picture
       Case 2 'Amarillo
        Set imgEstadoBeneficiarios.Picture = imgSemaforos.ListImages.Item(2).Picture
     End Select
     
     lblEstadoBeneficiarios.ToolTipText = "Fecha   .: " & Format(rs!Ben_Update_Fecha & "", "dd/mm/yyyy") & vbCrLf _
                                          & "Usuario .: " & rs!Ben_Update_Usuario & ""



     'Pregunta por el Consentimiento de Uso de la Información Personal para Contacto
    If IsNull(rs!Consentimiento_Contacto_Fecha) Then
       Set imgEstadoConsentimiento.Picture = imgSemaforos.ListImages.Item(3).Picture
    Else
       Set imgEstadoConsentimiento.Picture = imgSemaforos.ListImages.Item(1).Picture
    End If
     
     
    'Resumen
    
    txtCRD_Saldos.Text = Format(rs!Creditos_Saldo, "Standard")
    txtCRD_Cuotas.Text = Format(rs!Creditos_Cuota, "Standard")
    txtCRD_Operaciones.Text = Format(rs!Creditos_Operaciones, "###,##0")
    
    txtRET_Saldos.Text = Format(rs!Retencion_Saldo, "Standard")
    txtRET_Cuotas.Text = Format(rs!Retencion_Cuota, "Standard")
    txtRET_Operaciones.Text = Format(rs!Retencion_Operaciones, "###,##0")
    
    txtPlan_Aportes.Text = Format(rs!Fondos_Acumulado, "Standard")
    txtPlan_Mensualidad.Text = Format(rs!Fondos_Mensualidad, "Standard")
    txtPlan_Contratos.Text = Format(rs!Fondos_Contratos, "###,##0")
    
    
     
     
     
    'Scoring
     txtSC_Fecha.Text = Format(rs!SC_Fecha & "", "yyyy-mm-dd")
     
     txtSc_Morosidad.Text = rs!SC_Morosidad & ""
     txtSc_Endeudamiento.Text = rs!SC_Endeudamiento & ""
     txtSc_Liquidez.Text = rs!SC_Capacidad & ""
     txtSc_Historial_Pago.Text = rs!SC_Historial_Pago & ""
     txtSC_Garantia.Text = rs!SC_Garantia & ""
     
     
     txtSC_Salario_Devengado.Text = Format(rs!SC_Salario_Devengado, "Standard")
     txtSC_Salario_Liquido.Text = Format(rs!SC_Salario_Liquido, "Standard")
     
     Call sbSemaforo(txtSc_Morosidad, rs!SC_Morosidad_Color)
     Call sbSemaforo(txtSc_Endeudamiento, rs!SC_Endeudamiento_Color)
     Call sbSemaforo(txtSc_Liquidez, rs!SC_Capacidad_Color)
     Call sbSemaforo(txtSc_Historial_Pago, rs!SC_Historial_Color)
     Call sbSemaforo(txtSC_Garantia, rs!SC_Garantia_Color)
     
     
     
    'Datos Relevantes
    lblIG_Edad.Caption = "Edad " & rs!Edad & " años"
    lblIG_Generacion.Caption = rs!Generacion_Desc
    
    txtU_Credito.Text = Format(rs!Creditos_Ultimo, "yyyy-mm-dd")
    txtU_Ahorro.Text = Format(rs!Fondos_Ultimo, "yyyy-mm-dd")
    
    txtU_Liquidacion.Text = Format(rs!Fecha_Liquidación & "", "yyyy-mm-dd")
    
    txtU_Beneficio.Text = Format(rs!Ultimo_Beneficio & "", "yyyy-mm-dd")
     
    txtDisponible_SA.Text = Format(rs!Disponible_SobreAhorros, "Standard")
    txtDisponible_Exc.Text = Format(rs!Disponible_Excedentes, "Standard")
     
     
    'Graficos
    Dim pSQL As String
    
    
'    cboChart(0).Clear
'    cboChart(0).AddItem "3d_Pie"
'    cboChart(0).AddItem "3d_Pyramid"
'    cboChart(0).AddItem "3d_Torus"
'    cboChart(0).AddItem "3d_Doughnut"
'    cboChart(0).AddItem "3d_Funnel"
'    cboChart(0).Text = "3d_Torus"
    
    strSQL = "exec spDSB_Asociados_Series '" & txtCedula.Text & "', 'PAT'"
    Call sbChart_3D(ccPAT, "Patrimonio", "PAT", "3d_Doughnut", 1)
    
    strSQL = "exec spDSB_Asociados_Series '" & txtCedula.Text & "', 'CRD'"
    Call sbChart_3D(ccCreditos, "Créditos por Garantía", "CRD", "3d_Pie", 1)

    strSQL = "exec spDSB_Asociados_Series '" & txtCedula.Text & "', 'FND'"
    Call sbChart_3D(ccPlanes, "Planes de Ahorros", "FND", "3d_Torus", 1)

    strSQL = "exec spDSB_Asociados_Series '" & txtCedula.Text & "', 'BEN'"
    Call sbChart_3D(ccBeneficios, "Beneficios", "BEN", "3d_Pie", 1)
     
     
 Else
   MsgBox "No Se encontró registro de la persona solicitada", vbInformation
   Exit Sub
 End If
   
 
End Sub



Private Sub btnConsulta_Click()
Dim frm As Form

 If txtNombre.Text = "" Then Exit Sub
 
 Call sbFormsCall("frmCR_ConsultaCreditos")

For Each frm In Forms
  If UCase(frm.Name) = UCase("frmCR_ConsultaCreditos") Then
    Call frm.sbXConsultaAsistida(txtCedula.Text)
    Exit For
  End If
Next frm
             
End Sub

Private Sub sbBusqueda()

On Error GoTo vError

gBusquedas.Convertir = "N"

    Call Limpia
    gBusquedas.Col1Name = "Identificación"
    gBusquedas.Col2Name = "Id Alterno"
    gBusquedas.Col3Name = "Nombre"
    gBusquedas.Consulta = "Select cedula,cedular,nombre from SOCIOS"
    gBusquedas.Columna = "nombre"
    gBusquedas.Orden = "nombre"
    frmBusquedas.Show vbModal
    txtCedula = Trim(gBusquedas.Resultado)
    gBusquedas.Consulta = ""
    gBusquedas.Columna = ""
    gBusquedas.Orden = ""
    gBusquedas.Resultado = ""
    If Trim(txtCedula) <> "" Then
        Call sbConsulta(txtCedula)
    End If

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical


End Sub


Private Sub Form_Load()

vModulo = 24

mChartPallete = "Danville"
mChartLabelPosition = 3

Call Formularios(Me)

End Sub

Private Sub Timer_Access_Timer()

Timer_Access.Interval = 0

If btnAccess.Tag = "0" Then
   MsgBox "No cuentas con el acceso/permiso para esta opción. Contacte a su administrador!", vbExclamation
   Unload Me
End If

End Sub

Private Sub txtCedula_KeyDown(KeyCode As Integer, Shift As Integer)
Dim vCedTemp As String

On Error GoTo vError

'Busca primer en el Maestro de Socios, de lo contrario revisa si es una operacion
' y regresa la cedula de la operacion

vCedTemp = ""

If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then
 strSQL = "select isnull(count(*),0) as Existe from socios where cedula = '" & txtCedula & "'"
 Call OpenRecordSet(rs, strSQL)
 If rs!Existe = 0 Then
   rs.Close
   strSQL = "select cedula from reg_creditos where id_solicitud = " & txtCedula
   Call OpenRecordSet(rs, strSQL)
   If Not rs.EOF And Not rs.BOF Then
      vCedTemp = Trim(rs!Cedula)
   End If
 End If
 rs.Close
 
    If vCedTemp = "" Then
        Call sbConsulta(txtCedula.Text)
    Else
        Call sbConsulta(vCedTemp)
    End If
End If

If KeyCode = vbKeyF4 Then Call sbBusqueda

Exit Sub

vError:

End Sub



Private Sub txtNombre_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then Call sbBusqueda

End Sub

