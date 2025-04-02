VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.controls.v22.1.0.ocx"
Begin VB.Form frmCO_TrasladoDeuda 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Traspaso de Deuda o Cobro a Fiadores : Personalizado "
   ClientHeight    =   8235
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   11145
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8235
   ScaleWidth      =   11145
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.GroupBox gbTraslado 
      Height          =   1455
      Left            =   120
      TabIndex        =   27
      Top             =   3240
      Width           =   10935
      _Version        =   1441793
      _ExtentX        =   19283
      _ExtentY        =   2561
      _StockProps     =   79
      Caption         =   "Traslado de deuda o Cobro a Fiador"
      ForeColor       =   8421504
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
      BorderStyle     =   1
      Begin XtremeSuiteControls.FlatEdit txtLineaNueva 
         Height          =   312
         Left            =   1800
         TabIndex        =   28
         Top             =   360
         Width           =   1812
         _Version        =   1441793
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtLineaNuevaDesc 
         Height          =   312
         Left            =   3600
         TabIndex        =   29
         Top             =   360
         Width           =   6012
         _Version        =   1441793
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtTasaOriginal 
         Height          =   312
         Left            =   2760
         TabIndex        =   36
         ToolTipText     =   "Original"
         Top             =   1080
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   2
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit txtPlazoOriginal 
         Height          =   312
         Left            =   2760
         TabIndex        =   34
         ToolTipText     =   "Original"
         Top             =   720
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   2
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit txtPlazo 
         Height          =   312
         Left            =   1800
         TabIndex        =   33
         Top             =   720
         Width           =   972
         _Version        =   1441793
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtTasa 
         Height          =   312
         Left            =   1800
         TabIndex        =   35
         Top             =   1080
         Width           =   972
         _Version        =   1441793
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtPorcentajeAsg 
         Height          =   312
         Left            =   8280
         TabIndex        =   39
         Top             =   1080
         Width           =   1332
         _Version        =   1441793
         _ExtentX        =   2350
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
         Left            =   3720
         TabIndex        =   45
         Top             =   1080
         Width           =   2652
      End
      Begin VB.Label Label19 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "% Asignado"
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
         Left            =   6480
         TabIndex        =   40
         Top             =   1080
         Width           =   1692
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
         Left            =   360
         TabIndex        =   32
         Top             =   720
         Width           =   852
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
         Left            =   360
         TabIndex        =   31
         Top             =   1080
         Width           =   852
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Línea Cobro"
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
         TabIndex        =   30
         Top             =   360
         Width           =   1332
      End
   End
   Begin FPSpreadADO.fpSpread vGrid 
      Height          =   1575
      Left            =   120
      TabIndex        =   0
      Top             =   4800
      Width           =   10935
      _Version        =   524288
      _ExtentX        =   19288
      _ExtentY        =   2778
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
      MaxCols         =   496
      ScrollBars      =   2
      SpreadDesigner  =   "frmCO_TrasladoDeuda.frx":0000
      VScrollSpecialType=   2
      AppearanceStyle =   1
   End
   Begin XtremeSuiteControls.GroupBox gbDeuda 
      Height          =   1815
      Left            =   120
      TabIndex        =   1
      Top             =   1440
      Width           =   10935
      _Version        =   1441793
      _ExtentX        =   19288
      _ExtentY        =   3201
      _StockProps     =   79
      Caption         =   "Deuda"
      ForeColor       =   8421504
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
      Begin XtremeSuiteControls.DateTimePicker dtpCalculoIntCorte 
         Height          =   312
         Left            =   2400
         TabIndex        =   2
         Top             =   4560
         Width           =   1212
         _Version        =   1441793
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
      Begin XtremeSuiteControls.FlatEdit txtSaldo 
         Height          =   312
         Left            =   2880
         TabIndex        =   3
         Top             =   360
         Width           =   1812
         _Version        =   1441793
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtIntereses 
         Height          =   312
         Left            =   2880
         TabIndex        =   4
         Top             =   720
         Width           =   1812
         _Version        =   1441793
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtPoliza 
         Height          =   315
         Left            =   2880
         TabIndex        =   5
         Top             =   1440
         Width           =   1815
         _Version        =   1441793
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtCargos 
         Height          =   315
         Left            =   2880
         TabIndex        =   6
         Top             =   1080
         Width           =   1815
         _Version        =   1441793
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtCbrIntereses 
         Height          =   312
         Left            =   2400
         TabIndex        =   7
         Top             =   4080
         Width           =   1812
         _Version        =   1441793
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
         Height          =   315
         Left            =   7560
         TabIndex        =   8
         Top             =   360
         Width           =   1815
         _Version        =   1441793
         _ExtentX        =   3196
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   16777152
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
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
      Begin XtremeSuiteControls.FlatEdit txtDivisa 
         Height          =   315
         Left            =   9360
         TabIndex        =   9
         Top             =   360
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
      Begin XtremeSuiteControls.FlatEdit txtRecuperado 
         Height          =   315
         Left            =   7560
         TabIndex        =   47
         Top             =   720
         Width           =   1815
         _Version        =   1441793
         _ExtentX        =   3196
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   16777152
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
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
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Total Recuperado"
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
         Index           =   2
         Left            =   5640
         TabIndex        =   48
         Top             =   720
         Width           =   1935
      End
      Begin VB.Image imgCalculoInt 
         Height          =   252
         Index           =   1
         Left            =   3960
         Picture         =   "frmCO_TrasladoDeuda.frx":079E
         Stretch         =   -1  'True
         Top             =   4560
         Width           =   252
      End
      Begin VB.Image imgCalculoInt 
         Height          =   252
         Index           =   0
         Left            =   3600
         Picture         =   "frmCO_TrasladoDeuda.frx":0F4A
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
         TabIndex        =   16
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
         TabIndex        =   15
         Top             =   4572
         Width           =   1812
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Total Deuda"
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
         Index           =   17
         Left            =   5640
         TabIndex        =   14
         Top             =   360
         Width           =   1935
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
         Height          =   255
         Index           =   16
         Left            =   960
         TabIndex        =   13
         Top             =   1080
         Width           =   1695
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
         Height          =   255
         Index           =   19
         Left            =   960
         TabIndex        =   12
         Top             =   1440
         Width           =   1695
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Intereses"
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
         Left            =   960
         TabIndex        =   11
         Top             =   720
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
         TabIndex        =   10
         Top             =   360
         Width           =   1332
      End
   End
   Begin XtremeSuiteControls.FlatEdit txtOperacion 
      Height          =   432
      Left            =   2880
      TabIndex        =   17
      Top             =   120
      Width           =   1812
      _Version        =   1441793
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
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtProceso 
      Height          =   432
      Left            =   4680
      TabIndex        =   18
      Top             =   120
      Width           =   2052
      _Version        =   1441793
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
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtOpex 
      Height          =   432
      Left            =   6720
      TabIndex        =   19
      Top             =   120
      Width           =   1092
      _Version        =   1441793
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
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtCodigo 
      Height          =   312
      Left            =   2880
      TabIndex        =   20
      Top             =   600
      Width           =   1812
      _Version        =   1441793
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
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtCedula 
      Height          =   312
      Left            =   2880
      TabIndex        =   21
      Top             =   960
      Width           =   1812
      _Version        =   1441793
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
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtNombre 
      Height          =   312
      Left            =   4680
      TabIndex        =   22
      Top             =   960
      Width           =   6012
      _Version        =   1441793
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
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtDescripcion 
      Height          =   312
      Left            =   4680
      TabIndex        =   23
      Top             =   600
      Width           =   6012
      _Version        =   1441793
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
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtNotas 
      Height          =   792
      Left            =   1920
      TabIndex        =   37
      Top             =   6600
      Width           =   7812
      _Version        =   1441793
      _ExtentX        =   13779
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
   Begin XtremeSuiteControls.PushButton btnPrincipal 
      Height          =   612
      Index           =   0
      Left            =   1920
      TabIndex        =   41
      Top             =   7440
      Width           =   2412
      _Version        =   1441793
      _ExtentX        =   4254
      _ExtentY        =   1080
      _StockProps     =   79
      Caption         =   "Calcular"
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
      Picture         =   "frmCO_TrasladoDeuda.frx":18C7
   End
   Begin XtremeSuiteControls.PushButton btnPrincipal 
      Height          =   612
      Index           =   1
      Left            =   4320
      TabIndex        =   42
      Top             =   7440
      Width           =   2412
      _Version        =   1441793
      _ExtentX        =   4254
      _ExtentY        =   1080
      _StockProps     =   79
      Caption         =   "Trasladar Deuda"
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
      Picture         =   "frmCO_TrasladoDeuda.frx":1F8E
   End
   Begin XtremeSuiteControls.PushButton btnPrincipal 
      Height          =   612
      Index           =   2
      Left            =   6720
      TabIndex        =   43
      Top             =   7440
      Width           =   2412
      _Version        =   1441793
      _ExtentX        =   4254
      _ExtentY        =   1080
      _StockProps     =   79
      Caption         =   "Boleta"
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
      Picture         =   "frmCO_TrasladoDeuda.frx":2766
   End
   Begin XtremeSuiteControls.PushButton btnPrincipal 
      Height          =   612
      Index           =   3
      Left            =   9120
      TabIndex        =   44
      ToolTipText     =   "Cerrar Ventana"
      Top             =   7440
      Width           =   612
      _Version        =   1441793
      _ExtentX        =   1080
      _ExtentY        =   1080
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
      Picture         =   "frmCO_TrasladoDeuda.frx":2F22
   End
   Begin XtremeSuiteControls.FlatEdit txtNTraslado 
      Height          =   435
      Left            =   7800
      TabIndex        =   46
      Top             =   120
      Width           =   2895
      _Version        =   1441793
      _ExtentX        =   5106
      _ExtentY        =   767
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   16777152
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   "TRA-00001"
      BackColor       =   16777152
      Alignment       =   2
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
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
      Index           =   0
      Left            =   600
      TabIndex        =   38
      Top             =   6600
      Width           =   1092
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Operación"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   10.5
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
      TabIndex        =   26
      Top             =   120
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
      TabIndex        =   25
      Top             =   960
      Width           =   1332
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
      TabIndex        =   24
      Top             =   600
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
Attribute VB_Name = "frmCO_TrasladoDeuda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mOperacion As Long, mcurIntCor As Currency, mcurIntMor As Currency, mcurPrincipalMora As Currency
Dim mcurCargos As Currency, mcurPoliza As Currency, mOficina As String
Dim mTasaPts   As Currency, mTasaLiq As Integer, mcurIntPendiente As Currency

Private Sub btnPrincipal_Click(Index As Integer)
Dim i As Byte

Call sbSIFCleanTxtInject(txtNotas)


Select Case Index
  
  Case 0  'Calcular"
    Call sbCalcular
  
  Case 1 'Trasladar"
           i = MsgBox("Esta Seguro de realizar el traslado de la deuda?", vbYesNo)
           If i = vbYes Then
              Call sbTrasladar
           End If
  
  Case 2 'Boleta"
    Call sbBoleta
   
  
  Case 3 'Cerrar"
    Unload Me

End Select
 

End Sub

Private Sub Form_Load()

vModulo = 4

Set imgBanner.Picture = frmContenedor.imgBanner_01.Picture

vGrid.AppearanceStyle = fxGridStyle
 
mOperacion = GLOBALES.gTag
Call sbConsulta
 
End Sub

Private Sub sbConsulta()
Dim strSQL As String, rs As New ADODB.Recordset
Dim i As Integer, rsTmp As New ADODB.Recordset

On Error GoTo vError

txtOperacion.Text = mOperacion

'Se supone que si entra en esta ventana es porque esta previo validado en CO_PRINCIPAL
'Activa toda la barra

btnPrincipal.Item(0).Enabled = True
btnPrincipal.Item(1).Enabled = True
btnPrincipal.Item(2).Enabled = True
vGrid.Enabled = True


'Consulta de Parametros

txtLineaNueva = fxCBRParametro("16")
txtLineaNuevaDesc = fxDescribeCodigo(Trim(txtLineaNueva))

txtPlazo.Text = CStr(fxCBRPlazoRestante(mOperacion))

'Consulta el estado de la operación
If GLOBALES.SysPlanPagos = 1 Then
    strSQL = "select R.cedula,S.nombre,R.saldo,R.proceso,R.Interesv as Tasa,R.plazo,R.Int as TasaOriginal" _
           & ",R.codigo,C.descripcion,isnull(R.liqTasa,0) as LiqTasa, R.Proceso, R.OpEx, isnull(R.Cod_Divisa,'') as 'Divisa'" _
           & ",dbo.fxCRDNumFiadores(R.id_solicitud) as NumFiadores,R.Opex,R.TBP_PuntosAdd" _
           & ",isnull(V.amortiza,0) as MoraAmortiza" _
           & ",R.cod_oficina_r,R.cod_oficina_F,R.cod_oficina_comision, dbo.MyGetdate() as 'FechaServer'" _
           & " from Socios S inner join reg_creditos R on S.cedula = R.cedula inner join Catalogo C on R.codigo = C.codigo" _
           & " left join Vista_morosidad V on R.id_solicitud = V.id_solicitud" _
           & " Where R.id_solicitud = " & mOperacion
Else
    strSQL = "select R.cedula,S.nombre,R.saldo,R.proceso,R.Interesv as Tasa,R.plazo,R.Int as TasaOriginal" _
           & ",isnull(V.intc + V.intm,0) as IntAtrasado,R.codigo,C.descripcion,isnull(R.liqTasa,0) as LiqTasa" _
           & ",dbo.fxCRDNumFiadores(R.id_solicitud) as NumFiadores,R.Opex,R.TBP_PuntosAdd, isnull(R.Cod_Divisa,'') as 'Divisa'" _
           & ",isnull(V.intc,0) as MoraIntC,isnull(V.intm,0) as MoraIntM,isnull(V.amortiza,0) as MoraAmortiza" _
           & ",isnull(V.Cargos,0) as 'Cargos',dbo.fxCRDCalculoIntCorte(R.id_solicitud,dbo.MyGetdate()) as InteresTotal" _
           & ",0 as 'Poliza', dbo.MyGetdate() as 'FechaServer'" _
           & ",R.cod_oficina_r,R.cod_oficina_F,R.cod_oficina_comision" _
           & " from Socios S inner join reg_creditos R on S.cedula = R.cedula inner join Catalogo C on R.codigo = C.codigo" _
           & " left join Vista_morosidad V on R.id_solicitud = V.id_solicitud" _
           & " Where R.id_solicitud = " & mOperacion
End If


Call OpenRecordSet(rs, strSQL)
  
If GLOBALES.SysPlanPagos = 1 Then
  strSQL = "exec spCrdPlanPagosInfoCancelacion " & mOperacion & ",'" & Format(rs!FechaServer, "yyyy/mm/dd") & "'"
  Call OpenRecordSet(rsTmp, strSQL, 0)
    mcurIntCor = rsTmp!IntCor
    mcurIntMor = rsTmp!IntMor
    mcurPrincipalMora = rs!MoraAmortiza
    mcurIntPendiente = 0 'Intereses a hoy
    mcurCargos = rsTmp!Cargos
    mcurPoliza = rsTmp!Poliza
  rsTmp.Close
Else
  mcurIntCor = rs!MoraIntC
  mcurIntMor = rs!MoraIntM
  mcurPrincipalMora = rs!MoraAmortiza
  mcurIntPendiente = rs!InteresTotal - (mcurIntCor + mcurIntMor) 'Intereses a hoy
  mcurCargos = rs!Cargos
  mcurPoliza = rs!Poliza
End If
  mOficina = Trim(rs!cod_oficina_r & "")
  
  txtCedula.Text = rs!Cedula
  txtNombre.Text = rs!Nombre
  
  txtTasa.Text = Format(rs!Tasa, "Standard")
  
  txtSaldo.Text = Format(rs!Saldo, "Standard")
  
  txtCargos.Text = Format(mcurCargos, "Standard")
  txtPoliza.Text = Format(mcurPoliza, "Standard")
  txtIntereses.Text = Format(mcurIntCor + mcurIntMor + mcurIntPendiente, "Standard")
  txtIntereses.ToolTipText = "Int.Venc.Hoy : " & Format(mcurIntPendiente, "Standard")
  
  txtTotal.Text = Format(CCur(txtSaldo.Text) + mcurIntCor + mcurIntMor + mcurIntPendiente + mcurCargos + mcurPoliza, "Standard")
  
  txtOperacion.Tag = rs!opex
  
  txtDivisa.Text = rs!Divisa & ""
  
    Select Case rs!Proceso
      Case "N"
            txtProceso.Text = "Normal"
            
      Case "T" 'Traspaso
            btnPrincipal.Item(0).Enabled = False
            btnPrincipal.Item(1).Enabled = False

            vGrid.Enabled = False
      
            txtProceso.Text = "Traslado"
      
      Case "J", "C" 'Cobro Judicial
            btnPrincipal.Item(0).Enabled = False
            btnPrincipal.Item(1).Enabled = False
            btnPrincipal.Item(2).Enabled = False
            vGrid.Enabled = False
    
            txtProceso.Text = "Cobro Judicial"
      Case Else
            txtProceso.Text = "Otro"
            btnPrincipal.Item(0).Enabled = False
            btnPrincipal.Item(1).Enabled = False
            btnPrincipal.Item(2).Enabled = False
            vGrid.Enabled = False
    End Select
    
   txtPorcentajeAsg.Text = "100"
   txtPorcentajeAsg.Tag = rs!NumFiadores
    
   txtCodigo.Text = rs!Codigo
   txtDescripcion.Text = rs!Descripcion
       
   txtTasaOriginal.Text = Format(rs!TasaOriginal, "Standard")
   txtPlazoOriginal.Text = CStr(rs!Plazo)
       
    If Not IsNull(rs!TBP_PuntosAdd) Then
      lblTasa.Caption = "Tasa (TBP + " & rs!TBP_PuntosAdd & ")"
      mTasaPts = rs!TBP_PuntosAdd
    Else
      lblTasa.Caption = "Tasa %"
      mTasaPts = -1000 'Default para Indicar que es tasa Fija
    End If
    
    If rs!LiqTasa = 1 Then
      lblTasa.Caption = lblTasa.Caption & " + PtsLiq"
    End If
 
    mTasaLiq = rs!LiqTasa

   If rs!opex = 1 Then
        txtOpex.Text = "Sí"
   Else
        txtOpex.Text = "No"
   End If
       
rs.Close



'Consulta la lista de fiadores y Deudores
strSQL = "select 'D' as Tipo,S.cedula,S.Nombre,0 as Porcentaje,0 as Monto, 0 as Cuota,Est.Descripcion, 0, 0" _
       & " from Socios S inner join Reg_Creditos R on S.cedula = R.cedula" _
       & " inner join AFI_ESTADOS_PERSONA Est on S.estadoActual = Est.Cod_Estado" _
       & " Where R.id_solicitud = " & mOperacion _
       & " Union " _
       & " select 'F' as Tipo,S.cedula,S.Nombre," & (100 / txtPorcentajeAsg.Tag) & " as Porcentaje,0 as Monto, 0 as Cuota,Est.Descripcion, 0, 0" _
       & " from Socios S inner join Fiadores F on S.cedula = F.cedulaF" _
       & " inner join AFI_ESTADOS_PERSONA Est on S.estadoActual = Est.Cod_Estado" _
       & " where F.id_solicitud = " & mOperacion
Call sbCargaGrid(vGrid, 9, strSQL)
vGrid.MaxRows = vGrid.MaxRows - 1

Call sbCalcular


Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub sbCalcular()
Dim i As Integer, vPorcentaje As Currency, curMonto As Currency
Dim curTotalPorc As Currency

On Error GoTo vError

curTotalPorc = 0

For i = 1 To vGrid.MaxRows
    vGrid.Row = i
    vGrid.col = 4
    If CCur(vGrid.Text) = 0 Then
        vGrid.col = 5
        vGrid.Text = "0"
        vGrid.col = 6
        vGrid.Text = "0"
    Else
        curTotalPorc = curTotalPorc + CCur(vGrid.Text)
        vPorcentaje = CCur(vGrid.Text) / 100
        vGrid.col = 5
        vGrid.Text = CCur(txtTotal.Text) * vPorcentaje
        curMonto = CCur(txtTotal.Text) * vPorcentaje
        vGrid.col = 6
        vGrid.Text = fxCalcula_Cuota(CDbl(curMonto), txtPlazo, txtTasa)
    End If
Next i

txtPorcentajeAsg.Text = curTotalPorc

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Function fxVerificar() As Boolean
Dim strSQL As String, rs As New ADODB.Recordset
Dim vMensaje As String

On Error GoTo vError

vMensaje = ""

If CCur(txtPorcentajeAsg) <> 100 Then
   vMensaje = vMensaje & " - El porcentaje de Asignación tiene que ser 100%..." & vbCrLf
End If

txtPlazo = CLng(txtPlazo)

If CLng(txtPlazo) > 300 Or CLng(txtPlazo) < 1 Then
   vMensaje = vMensaje & " - El plazo es incorrecto verifique..." & vbCrLf
End If

If CCur(txtTasa) > 100 Or CLng(txtTasa) < 0 Then
   vMensaje = vMensaje & " - La Tasa es incorrecta verifique..." & vbCrLf
End If

If CCur(txtTotal) = 0 Then
   vMensaje = vMensaje & " - No existe un monto a trasladar..." & vbCrLf
End If

If Len(Trim(txtNotas)) = 0 Then
   vMensaje = vMensaje & " - Especifique una Nota para el traslado..." & vbCrLf
End If

'Verifica que exista la línea
strSQL = "select isnull(count(*),0) as Existe from catalogo where codigo = '" & txtLineaNueva & "'"
Call OpenRecordSet(rs, strSQL)
If rs!Existe = 0 Then
   vMensaje = vMensaje & " - La Línea para Traslado de Deudas No Existe..." & vbCrLf
End If
rs.Close

'Verifica que la Operacion se encuentre en Proceso Normal / Para evitar accidentes
strSQL = "select isnull(count(*),0) as Existe from reg_creditos where proceso = 'N' and id_solicitud= '" & txtOperacion.Text & "'"
Call OpenRecordSet(rs, strSQL)
If rs!Existe = 0 Then
   vMensaje = vMensaje & " - La Operación no se encuentra en PROCESO NORMAL para realizar el traslado..." & vbCrLf
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


Private Sub sbTrasladar()
Dim strSQL As String, rs As New ADODB.Recordset
Dim lngPriDeduc As Long, vFecha As Date, i As Integer
Dim curDeuda As Currency, rsTmp As New ADODB.Recordset

Dim vTempCed As String, vTempMonto As Currency, vTempCuota As Currency, vTempOperacion As Long
Dim vTempOpex As String, vTempCtaNormal As String, vTempCtaOpex As String

Dim vTipo As String, pTipoDocum As String, pNumDoc As String, pConcepto As String
Dim pOficina As String, pUnidad As String, pCentroCosto As String, pDivisa As String
Dim strLinea(11) As String, pBaseCalculo As String, pCuota As Currency, pDiaPago As Integer
Dim vTransac As Boolean, pTipoCambio As Currency

On Error GoTo vError

'Verificación
 Call sbCalcular
 
 If Not fxVerificar Then
    Exit Sub
 End If


'Inicia Proceso
Me.MousePointer = vbHourglass


vTransac = False

strSQL = "select O.cod_oficina,O.cod_unidad,O.cod_centro_costo,R.cod_divisa,R.dia_pago, R.Base_calculo" _
       & ", dbo.fxCntXTipoCambio(" & GLOBALES.gEnlace & ",R.cod_divisa , getdate(), 'V') as 'TipoCambio'" _
       & " from reg_creditos R left join sif_oficinas O on R.cod_oficina_R = O.cod_oficina" _
       & " where id_solicitud = " & txtOperacion.Text
Call OpenRecordSet(rs, strSQL)
    pOficina = Trim(rs!COD_OFICINA & "")
    pUnidad = rs!Cod_Unidad & ""
    pCentroCosto = rs!Cod_Centro_Costo & ""
    pDivisa = IIf(IsNull(rs!COD_DIVISA), "COL", rs!COD_DIVISA)
    pDiaPago = IIf(IsNull(rs!dia_pago), 32, rs!dia_pago)
    pBaseCalculo = IIf(IsNull(rs!Base_Calculo), "01", rs!Base_Calculo)
    pTipoCambio = rs!TipoCambio
rs.Close

If pOficina = "" Then
    'Información Base de la Operacion
    strSQL = "exec spCrdOperacionCtas " & txtOperacion.Text
    Call OpenRecordSet(rs, strSQL)
        pOficina = rs!cod_oficina_r
        pUnidad = rs!Cod_Unidad
        pCentroCosto = rs!Cod_Centro_Costo
    rs.Close
End If



'Registro Inicial en Control de Documentos
    vTipo = "TRA"
    pTipoDocum = "TRA"
    pNumDoc = fxDocumentoConsecutivo(vTipo)
    pConcepto = "CBR002"
    vAseDocDetalle = txtNotas.Text
    vAseDocDeposito = ""
    
    
    strLinea(1) = "Saldo Anterior    " & txtSaldo.Text
    strLinea(2) = "Interes Corriente " & Format(mcurIntCor + mcurIntPendiente, "Standard")
    strLinea(3) = "Interes Moratorio " & Format(mcurIntMor, "Standard")
    strLinea(4) = "Cargos            " & Format(mcurCargos, "Standard")
    strLinea(5) = "Amortizacion      " & txtSaldo.Text
    strLinea(6) = "Saldo Actual      " & Format(0, "Standard")
    strLinea(7) = "Operación         " & txtOperacion.Text
    strLinea(8) = "Línea             " & txtCodigo.Text
    strLinea(9) = "Proc.Retencion    " & "NO"
    strLinea(10) = "Usuario           " & glogon.Usuario
    strLinea(11) = "Póliza            " & Format(mcurPoliza, "Standard")
          
    strSQL = "insert SIF_TRANSACCIONES(COD_TRANSACCION,TIPO_DOCUMENTO,REGISTRO_FECHA,REGISTRO_USUARIO,Cliente_IDENTIFICACION,CLIENTE_NOMBRE" _
             & ",cod_concepto,monto,estado,Referencia_01,Referencia_02,Referencia_03,cod_oficina" _
             & ",linea1,linea2,linea3,linea4,linea5,linea6,linea7,linea8,linea9,linea10,detalle,documento,linea11)" _
             & " values('" & pNumDoc & "','" & pTipoDocum & "',dbo.MyGetdate(),'" & glogon.Usuario & "','" & Trim(txtCedula.Text) _
             & "','" & Trim(txtNombre.Text) & "','" & pConcepto & "'," & CCur(txtTotal.Text) & ",'P','" & txtOperacion.Text _
             & "','" & txtCodigo.Text & "','" & vAseDocDeposito & "','" & GLOBALES.gOficinaTitular & "','" & strLinea(1) & "','" _
             & strLinea(2) & "','" & strLinea(3) & "','" & strLinea(4) & "','" _
             & strLinea(5) & "','" & strLinea(6) & "','" & strLinea(7) & "','" _
             & strLinea(8) & "','" & strLinea(9) & "','" & strLinea(10) & "','" _
             & vAseDocDetalle & "','" & vAseDocDeposito & "','" & strLinea(11) & "')"
     Call ConectionExecute(strSQL)


vFecha = fxFechaServidor

strSQL = "select  ISNULL( dbo.fxSIFDateTimeToProceso( dbo.fxCrd_Primer_Deduccion(cod_deductora) ), 0) as 'PriDeduc', cod_deductora" _
       & " From vAfi_Persona_Deductora  WHERE CEDULA = '" & txtCedula.Text & "'"
Call OpenRecordSet(rs, strSQL)
    If rs!PriDeduc > 0 Then
       lngPriDeduc = rs!PriDeduc
    Else
        lngPriDeduc = fxPrimerDeduccion(, rs!cod_Deductora)
    End If
rs.Close

strSQL = "select ctanAmort, ctaOAmort from catalogo where codigo = '" & txtLineaNueva & "'"
Call OpenRecordSet(rs, strSQL)
   vTempCtaNormal = rs!CtaNamort
   vTempCtaOpex = rs!CtaOamort
rs.Close

'Inicia Transacciones
glogon.Conection.BeginTrans
vTransac = True

For i = 1 To vGrid.MaxRows
    vGrid.Row = i
    vGrid.col = 4
    If CCur(vGrid.Text) > 0 Then
       
       vGrid.col = 2
       vTempCed = vGrid.Text
       vGrid.col = 5
       vTempMonto = vGrid.Text
       vGrid.col = 6
       vTempCuota = vGrid.Text
       vGrid.col = 7
       vTempOpex = Trim(vGrid.Text)
       
       If vTempOpex = "S" Then
          vTempOpex = "0"
       Else
          vTempOpex = "1"
       End If
       
        ''TODO OFICINAS y CARGOS
       
        'Registra Nueva Operacion
        strSQL = "insert into reg_creditos(codigo,id_comite,cedula,montosol,estadosol,fechares" _
               & ",plazo,int,interesv,montoapr,prideduc,fechaforp,fechaforf,saldo,amortiza,interesc" _
               & ",cuota,referencia,userrec,userres,userfor,garantia,firma_deudor" _
               & ",monto_girado,cuotas_planilla,cuotas_directas,cuotas_anuladas,Tesoreria,opex,FECULT,observacion,TBP_PuntosAdd" _
               & ",LiqTasa,cod_oficina_R,cod_oficina_f, base_calculo,cod_divisa,cuota_fija,dia_pago) values" _
               & "('" & txtLineaNueva & "',1,'" & vTempCed & "'," & vTempMonto & ",'F','" & Format(vFecha, "yyyy/mm/dd") & "'" _
               & "," & txtPlazo & "," & txtTasa & "," & txtTasa & "," & vTempMonto & "," & lngPriDeduc _
               & ",'" & Format(vFecha, "yyyy/mm/dd") & "','" & Format(vFecha, "yyyy/mm/dd") & "'," & vTempMonto _
               & ",0,0," & vTempCuota & "," & txtOperacion & ",'" & glogon.Usuario & "','" & glogon.Usuario & "','" & glogon.Usuario & "','F',1" _
               & ",0,0,0,0,'" & Format(vFecha, "yyyy/mm/dd") & "'," & vTempOpex & "," & GLOBALES.glngFechaCR & ",'" _
               & Trim(txtNotas.Text) & "'," & IIf((mTasaPts = -1000), "null", mTasaPts) & "," & mTasaLiq & ",'" & pOficina _
               & "','" & GLOBALES.gOficinaTitular & "','" & pBaseCalculo & "','" & pDivisa & "',0," & pDiaPago & ")"
         Call ConectionExecute(strSQL)
        
        strSQL = "select max(id_solicitud) as 'Operacion' from reg_Creditos where cedula = '" & vTempCed & "' and codigo = '" & txtLineaNueva.Text & "'"
        Call OpenRecordSet(rsTmp, strSQL, 0)
           vTempOperacion = rsTmp!Operacion
        rsTmp.Close
        
        If GLOBALES.SysPlanPagos = 1 Then
           'Crea Plan de Pagos para las Nuevas Operaciones
           strSQL = "exec spCrdPlanPagos " & vTempOperacion
           Call ConectionExecute(strSQL)
        End If
        
        strSQL = "exec spSIFDocsAsiento '" & pTipoDocum & "','" & pNumDoc & "'," & vTempMonto & ",'D','" & pDivisa _
               & "',1," & GLOBALES.gEnlace & ",'" & pUnidad & "','','" & IIf((vTempOpex = "1"), vTempCtaOpex, vTempCtaNormal) _
               & "','" & vTempOperacion & "','" & txtLineaNueva.Text & "','" & vAseDocDeposito & "'"
        Call ConectionExecute(strSQL)
         
    End If
Next i

If GLOBALES.SysPlanPagos = 1 Then
    'Actualiza Estado del Plan de Pago
    strSQL = "exec spCrdPlanPagosMoraActualizaOp " & txtOperacion.Text & ",'" & Format(vFecha, "yyyy/mm/dd") & "'"
    Call ConectionExecute(strSQL)
  
    'Aplica Abono de Cancelación
    strSQL = "exec spCrdPlanPagoAbonoCancelacion " & txtOperacion.Text & ",'" & pConcepto & "','" & glogon.Usuario & "','" & pTipoDocum _
           & "','" & pNumDoc & "'," & CCur(txtTotal.Text) & ",'" & Format(vFecha, "yyyy/mm/dd hh:mm:ss") & "',''"
    Call ConectionExecute(strSQL)

    'Cambiado por esta para que aparezca en los reportes de estados de cuenta
    strSQL = "update reg_creditos set estado = 'A', Proceso = 'T'" _
           & ",fecha_enviaproceso = dbo.MyGetdate()" _
           & ",observacion_proceso = '" & txtNotas & "'" _
           & " where id_solicitud = " & txtOperacion.Text
    Call ConectionExecute(strSQL)


Else
    'Cancela Mora Activa del Deudor
    strSQL = "Update morosidad set abintc = intc, abintm = intm, abamortiza = amortiza, abCargo = Cargo" _
           & ", estado = 'C', tcon = '" & vTipo & "',ncon = '" _
           & pNumDoc & "',fecult = dbo.MyGetdate(), cod_concepto = '" & pConcepto & "', usuario = '" & glogon.Usuario _
           & "' where estado = 'A' and id_solicitud = " & txtOperacion
    Call ConectionExecute(strSQL)
    
    'Si se cobraron intereses a hoy entonces registrar linea de detalle como cancelada
    If mcurIntPendiente > 0 Then
       strSQL = "insert MOROSIDAD(CODIGO,ID_SOLICITUD,FECHAP,FECAP ,FECULT,ESTADO,ESTADOI,INTC,INTM,AMORTIZA ,CARGO" _
              & ",ABINTC,ABINTM,ABAMORTIZA ,ABCARGO , TCON,NCON, cod_concepto,usuario,cod_caja) values('" & txtCodigo & "'," & txtOperacion.Text _
              & "," & GLOBALES.glngFechaCR & "," & GLOBALES.glngFechaCR & ",dbo.MyGetdate(),'C','A'," _
              & mcurIntPendiente & ",0,0,0," & mcurIntPendiente & ",0,0,0,'" & vTipo & "','" & pNumDoc & "','" & pConcepto & "','" & glogon.Usuario & "','')"
        Call ConectionExecute(strSQL)
    End If
    
    'INSERT EN CREDITOS DT POR LA DIFERENCIA EN EL SALDO
    strSQL = "insert creditos_dt(CODIGO,ID_SOLICITUD,CUOTA,ABONO,INTCP,AMORTIZA,FECHAS," _
           & "FECHAP,TCON,NCON,ESTADO,ESTADO_ASIENTO, cod_concepto,usuario,cod_caja) values('" & txtCodigo & "'," _
           & txtOperacion & ",0," & CCur(txtSaldo) - mcurPrincipalMora _
           & ",0," & CCur(txtSaldo) - mcurPrincipalMora & ",dbo.MyGetdate()" _
           & "," & GLOBALES.glngFechaCR & ",'" & vTipo & "','" & pNumDoc & "','A','G','" & pConcepto & "','" & glogon.Usuario & "','')"
    Call ConectionExecute(strSQL)


    'Cambiado por esta para que aparezca en los reportes de estados de cuenta
    strSQL = "update reg_creditos set saldo = saldo - " & CCur(txtSaldo) & ", amortiza = amortiza + " & CCur(txtSaldo) _
           & ", interesc = interesc + " & CCur(txtIntereses) & ", estado = 'A', Proceso = 'T'" _
           & ",fecha_enviaproceso = dbo.MyGetdate()" _
           & ",observacion_proceso = '" & txtNotas & "'" _
           & " where id_solicitud = " & txtOperacion
    Call ConectionExecute(strSQL)

End If

'Cierra Transacciones
glogon.Conection.CommitTrans
vTransac = False



'Hacer Asiento Aqui *************************************************************

 If CInt(txtOperacion.Tag) = 0 Then
  strSQL = "select ctanintc as ctaIntc, ctanintm as ctaIntm, ctanamort as ctaAmortiza "
 Else 'cuentas opex
  strSQL = "select ctaointc as ctaIntc, ctaointm as ctaIntm, ctaoamort as ctaAmortiza "
 End If
 strSQL = strSQL & " from catalogo where codigo = '" & txtCodigo & "'"
 Call OpenRecordSet(rs, strSQL)
 
 strSQL = ""
 
'Asiento
If CCur(txtSaldo) > 0 Then
    strSQL = strSQL & Space(10) & "exec spSIFDocsAsiento '" & pTipoDocum & "','" & pNumDoc & "'," & CCur(txtSaldo) & ",'C','" & pDivisa _
           & "',1," & GLOBALES.gEnlace & ",'" & pUnidad & "','','" & rs!ctaamortiza _
           & "','" & txtOperacion.Text & "','" & txtCodigo.Text & "','" & vAseDocDeposito & "'"
End If

If mcurCargos > 0 Then
    strSQL = strSQL & Space(10) & "exec spSIFDocsAsiento '" & pTipoDocum & "','" & pNumDoc & "'," & mcurCargos & ",'C','" & pDivisa _
           & "',1," & GLOBALES.gEnlace & ",'" & pUnidad & "','','" & fxCBRParametro("23") _
           & "','" & txtOperacion.Text & "','" & txtCodigo.Text & "','" & vAseDocDeposito & "'"
End If

'If mcurCargos > 0 Then
'    strSQL = strSQL & Space(10) & "exec spSIFDocsAsiento '" & pTipoDocum & "','" & pNumDoc & "'," & mcurCargos & ",'C','" & pDivisa _
'           & "',1," & GLOBALES.gEnlace & ",'" & pUnidad & "','','" & fxCBRParametro("23") _
'           & "','" & txtOperacion.Text & "','" & txtCodigo.Text & "','" & vAseDocDeposito & "'"
'End If

'--------------------------------------------
 If mcurCargos <> 0 Then
 'Detallar Cargos
   glogon.strSQL = "exec spCrdDocumentoAfectacionCargos '" & pTipoDocum & "','" & pNumDoc & "'"
   Call OpenRecordSet(rsTmp, glogon.strSQL, 0)
   Do While Not rsTmp.EOF
         strSQL = strSQL & Space(10) & "exec spSIFDocsAsiento '" & pTipoDocum & "','" & pNumDoc & "'," & IIf(IsNull(rsTmp!Mov_Monto), mcurCargos, rsTmp!Mov_Monto * fxSys_Tipo_Cambio_Apl(pTipoCambio)) & ",'C','" & rs!COD_DIVISA _
                & "'," & pTipoCambio & "," & GLOBALES.gEnlace & ",'" & rsTmp!Cod_Unidad & "','" & rsTmp!Cod_Centro_Costo & "','" & rsTmp!cod_cuenta _
                & "','" & rsTmp!Id_Solicitud & "','" & rsTmp!Codigo & "','" & vAseDocDeposito & "'"
         rsTmp.MoveNext
   Loop
   rsTmp.Close
 End If
 
 If mcurPoliza <> 0 Then
  
 'Detallar Poliza
   glogon.strSQL = "exec spCrdDocumentoAfectacionPolizas '" & pTipoDocum & "','" & pNumDoc & "'"
   Call OpenRecordSet(rsTmp, glogon.strSQL, 0)
   Do While Not rsTmp.EOF
         strSQL = strSQL & Space(10) & "exec spSIFDocsAsiento '" & pTipoDocum & "','" & pNumDoc & "'," & rsTmp!Mov_Monto * fxSys_Tipo_Cambio_Apl(pTipoCambio) & ",'C','" & rsTmp!COD_DIVISA _
                & "'," & pTipoCambio & "," & GLOBALES.gEnlace & ",'" & rsTmp!Cod_Unidad & "','" & rsTmp!Cod_Centro_Costo & "','" & rsTmp!cod_cuenta _
                & "','" & rsTmp!Id_Solicitud & "','" & rsTmp!Codigo & "','" & vAseDocDeposito & "'"
         rsTmp.MoveNext
   Loop
   rsTmp.Close
   
 End If


'--------------------------------------------

If mcurIntCor > 0 Then
    strSQL = strSQL & Space(10) & "exec spSIFDocsAsiento '" & pTipoDocum & "','" & pNumDoc & "'," & mcurIntCor & ",'C','" & pDivisa _
           & "',1," & GLOBALES.gEnlace & ",'" & pUnidad & "','','" & rs!ctaintc _
           & "','" & txtOperacion.Text & "','" & txtCodigo.Text & "','" & vAseDocDeposito & "'"
End If

If mcurIntPendiente > 0 Then
    strSQL = strSQL & Space(10) & "exec spSIFDocsAsiento '" & pTipoDocum & "','" & pNumDoc & "'," & mcurIntPendiente & ",'C','" & pDivisa _
           & "',1," & GLOBALES.gEnlace & ",'" & pUnidad & "','','" & rs!ctaintc _
           & "','" & txtOperacion.Text & "','" & txtCodigo.Text & "','" & vAseDocDeposito & "'"
End If

If mcurIntMor > 0 Then
    strSQL = strSQL & Space(10) & "exec spSIFDocsAsiento '" & pTipoDocum & "','" & pNumDoc & "'," & mcurIntMor & ",'C','" & pDivisa _
           & "',1," & GLOBALES.gEnlace & ",'" & pUnidad & "','','" & rs!ctaintm _
           & "','" & txtOperacion.Text & "','" & txtCodigo.Text & "','" & vAseDocDeposito & "'"
End If

'Procesa el Lote
If Len(strSQL) > 0 Then
    Call ConectionExecute(strSQL)
End If

rs.Close


'Bitacora de Sistema
Call Bitacora("Aplica", "Traspaso de Deudas de la Operación:" & txtOperacion)

'Bitacora de Cobro
Call sbCBRRegTransac("01", txtCedula, txtOperacion, txtNotas, CCur(txtSaldo), mcurIntCor + mcurIntPendiente, mcurIntMor _
                    , mcurCargos, mcurPoliza, mcurPrincipalMora, pTipoDocum, pNumDoc)

Call sbConsulta
Call sbBoleta
 
Me.MousePointer = vbDefault

MsgBox "Traspaso de Deudas a Realizado Satisfactoriamente..." _
       & vbCrLf & vbCrLf & " - Se la nota de cobro:" & pTipoDocum & "-" & pNumDoc, vbInformation
Call sbImprimeRecibo(pNumDoc, pTipoDocum)

Exit Sub

vError:
 Me.MousePointer = vbDefault
 If vTransac Then
  glogon.Conection.RollbackTrans
 End If
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub sbBoleta()
Dim vTipoDoc As String

Me.MousePointer = vbHourglass

If GLOBALES.SysDocVersion = 1 Then
 vTipoDoc = "4"
Else
 vTipoDoc = "TRA"
End If

With frmContenedor.Crt
    .Reset
    .WindowShowRefreshBtn = True
    .WindowShowPrintSetupBtn = True
    .WindowState = crptMaximized
    .WindowShowSearchBtn = True
    .WindowTitle = "Reportes Módulo de Cobro Administrativo y Judicial"
    
    .Connect = glogon.ConectRPT
    
    .Formulas(0) = "fecha='" & Format(fxFechaServidor, "dd/mm/yyyy") & "'"
    .Formulas(1) = "Empresa='" & GLOBALES.gstrNombreEmpresa & "'"
    .Formulas(2) = "subtitulo='BOLETA DE TRASLADO Y REVERSION DE DEUDAS'"
    .ReportFileName = SIFGlobal.fxPathReportes("Cobro_BoletaTraspasoReversion.rpt")
    .SelectionFormula = "{REG_CREDITOS.ID_SOLICITUD} =" & txtOperacion
    
    .SubreportToChange = "Movimientos"
    .SelectionFormula = "{vCRDsReportesMov.id_solicitud}  = {?Pm-REG_CREDITOS.ID_SOLICITUD} and {vCRDsReportesMov.TCON} = '" & vTipoDoc & "'"
    .PrintReport
End With

Me.MousePointer = vbDefault

End Sub


Private Sub vGrid_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyReturn And vGrid.ActiveCol = 4 Then
   Call sbCalcular
End If

End Sub

