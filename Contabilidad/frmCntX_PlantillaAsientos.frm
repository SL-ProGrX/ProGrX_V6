VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#20.3#0"; "Codejock.Controls.v20.3.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#20.3#0"; "Codejock.ShortcutBar.v20.3.0.ocx"
Begin VB.Form frmCntX_PlantillaAsientos 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Plantillas para Asientos Fijos"
   ClientHeight    =   8055
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   11985
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8055
   ScaleWidth      =   11985
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.GroupBox GroupBox1 
      Height          =   1575
      Left            =   0
      TabIndex        =   13
      Top             =   1080
      Width           =   12015
      _Version        =   1310723
      _ExtentX        =   21193
      _ExtentY        =   2778
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      BorderStyle     =   1
      Begin XtremeSuiteControls.FlatEdit txtCAsiento 
         Height          =   330
         Left            =   1560
         TabIndex        =   14
         Top             =   360
         Width           =   975
         _Version        =   1310723
         _ExtentX        =   1720
         _ExtentY        =   582
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
      Begin XtremeSuiteControls.FlatEdit txtDAsiento 
         Height          =   330
         Left            =   2520
         TabIndex        =   15
         Top             =   360
         Width           =   3255
         _Version        =   1310723
         _ExtentX        =   5741
         _ExtentY        =   582
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
         Height          =   330
         Left            =   1560
         TabIndex        =   17
         Top             =   720
         Width           =   10215
         _Version        =   1310723
         _ExtentX        =   18018
         _ExtentY        =   582
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtDetalle 
         Height          =   330
         Left            =   5640
         TabIndex        =   19
         Top             =   1080
         Width           =   6135
         _Version        =   1310723
         _ExtentX        =   10821
         _ExtentY        =   582
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtAnio 
         Height          =   315
         Left            =   7320
         TabIndex        =   24
         Top             =   360
         Width           =   735
         _Version        =   1310723
         _ExtentX        =   1296
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtPeriodo 
         Height          =   315
         Left            =   8520
         TabIndex        =   25
         Top             =   360
         Width           =   3255
         _Version        =   1310723
         _ExtentX        =   5741
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
      Begin XtremeSuiteControls.FlatEdit txtMes 
         Height          =   315
         Left            =   8040
         TabIndex        =   26
         Top             =   360
         Width           =   495
         _Version        =   1310723
         _ExtentX        =   873
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtDocumento 
         Height          =   330
         Left            =   1560
         TabIndex        =   18
         Top             =   1080
         Width           =   2535
         _Version        =   1310723
         _ExtentX        =   4471
         _ExtentY        =   582
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
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Detalle"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   4
         Left            =   4320
         TabIndex        =   23
         Top             =   1080
         Width           =   1305
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Documento"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   3
         Left            =   120
         TabIndex        =   22
         Top             =   1080
         Width           =   1305
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Descripción"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   2
         Left            =   120
         TabIndex        =   21
         Top             =   720
         Width           =   1305
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Inicia en"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   6480
         TabIndex        =   20
         Top             =   360
         Width           =   1305
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo Asiento"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   120
         TabIndex        =   16
         Top             =   360
         Width           =   1305
      End
   End
   Begin MSComctlLib.Toolbar tlb 
      Align           =   1  'Align Top
      Height          =   330
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11985
      _ExtentX        =   21140
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Style           =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   9
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "nuevo"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "editar"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "borrar"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "guardar"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "deshacer"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "consultar"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "reportes"
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "ayuda"
         EndProperty
      EndProperty
   End
   Begin FPSpreadADO.fpSpread vGrid 
      Height          =   3735
      Left            =   0
      TabIndex        =   7
      Top             =   3240
      Width           =   11895
      _Version        =   524288
      _ExtentX        =   20981
      _ExtentY        =   6588
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
      MaxCols         =   494
      ScrollBars      =   2
      SpreadDesigner  =   "frmCntX_PlantillaAsientos.frx":0000
      VScrollSpecial  =   -1  'True
      VScrollSpecialType=   2
      AppearanceStyle =   1
   End
   Begin MSComCtl2.FlatScrollBar FlatScrollBar 
      Height          =   255
      Left            =   8760
      TabIndex        =   8
      Top             =   600
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   450
      _Version        =   393216
      Arrows          =   65536
      Orientation     =   1638401
   End
   Begin XtremeSuiteControls.FlatEdit txtCodPlantilla 
      Height          =   315
      Left            =   1680
      TabIndex        =   9
      Top             =   600
      Width           =   975
      _Version        =   1310723
      _ExtentX        =   1720
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
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtConsecutivo 
      Height          =   315
      Left            =   9480
      TabIndex        =   10
      Top             =   600
      Width           =   2295
      _Version        =   1310723
      _ExtentX        =   4048
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
      Alignment       =   2
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtDesPlantilla 
      Height          =   315
      Left            =   2640
      TabIndex        =   11
      Top             =   600
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
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtDebito 
      Height          =   315
      Left            =   8040
      TabIndex        =   30
      Top             =   7200
      Width           =   1815
      _Version        =   1310723
      _ExtentX        =   3201
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
   Begin XtremeSuiteControls.FlatEdit txtCredito 
      Height          =   315
      Left            =   9840
      TabIndex        =   31
      Top             =   7200
      Width           =   1815
      _Version        =   1310723
      _ExtentX        =   3201
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
   Begin XtremeSuiteControls.FlatEdit txtDiferencia 
      Height          =   315
      Left            =   5160
      TabIndex        =   32
      Top             =   7200
      Width           =   1575
      _Version        =   1310723
      _ExtentX        =   2778
      _ExtentY        =   556
      _StockProps     =   77
      ForeColor       =   192
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
   Begin XtremeSuiteControls.FlatEdit txtDebitosPro 
      Height          =   315
      Left            =   8040
      TabIndex        =   33
      Top             =   7560
      Width           =   1815
      _Version        =   1310723
      _ExtentX        =   3201
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
   Begin XtremeSuiteControls.FlatEdit txtCreditosPro 
      Height          =   315
      Left            =   9840
      TabIndex        =   34
      Top             =   7560
      Width           =   1815
      _Version        =   1310723
      _ExtentX        =   3201
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
   Begin XtremeSuiteControls.FlatEdit txtDiferenciaPro 
      Height          =   315
      Left            =   5160
      TabIndex        =   35
      Top             =   7560
      Width           =   1575
      _Version        =   1310723
      _ExtentX        =   2778
      _ExtentY        =   556
      _StockProps     =   77
      ForeColor       =   192
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
   Begin XtremeShortcutBar.ShortcutCaption lblDivisa 
      Height          =   375
      Left            =   8760
      TabIndex        =   29
      Top             =   2760
      Width           =   3135
      _Version        =   1310723
      _ExtentX        =   5530
      _ExtentY        =   661
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
      Alignment       =   2
   End
   Begin XtremeShortcutBar.ShortcutCaption lblUnidad 
      Height          =   375
      Left            =   4800
      TabIndex        =   28
      Top             =   2760
      Width           =   3975
      _Version        =   1310723
      _ExtentX        =   7011
      _ExtentY        =   661
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
      Alignment       =   1
   End
   Begin XtremeShortcutBar.ShortcutCaption lblCuenta 
      Height          =   375
      Left            =   0
      TabIndex        =   27
      Top             =   2760
      Width           =   4815
      _Version        =   1310723
      _ExtentX        =   8493
      _ExtentY        =   661
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
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Plantilla"
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
      Left            =   480
      TabIndex        =   12
      Top             =   600
      Width           =   1005
   End
   Begin VB.Label Label8 
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      Caption         =   "Totales:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   6960
      TabIndex        =   6
      Top             =   7590
      Width           =   765
   End
   Begin VB.Label Label10 
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      Caption         =   "Diferencia:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   4200
      TabIndex        =   5
      Top             =   7590
      Width           =   915
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Movimientos Proyectados 1 Periodo "
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   1
      Left            =   360
      TabIndex        =   4
      Top             =   7560
      Width           =   3435
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Movimientos Asiento de Plantilla Base "
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   3
      Top             =   7200
      Width           =   3435
   End
   Begin VB.Label lsblr 
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      Caption         =   "Diferencia:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   4200
      TabIndex        =   2
      Top             =   7230
      Width           =   915
   End
   Begin VB.Label Label8 
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      Caption         =   "Totales:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   6960
      TabIndex        =   1
      Top             =   7230
      Width           =   765
   End
End
Attribute VB_Name = "frmCntX_PlantillaAsientos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Type xUltimos
   Tipo As String
   Valor As Currency
   Divisa             As String
   DivisaDesc         As String
   Unidad             As String
   UnidadDesc         As String
   CC                 As String
   CCDesc             As String
End Type
Dim vEdita As Boolean, vBusca As Integer, vUltimos As xUltimos


Private Sub sbLimpiezaParcial(iCodigo As Integer)
vGrid.MaxRows = 0
vGrid.MaxRows = 1

txtDescripcion = ""

Select Case iCodigo
  Case 1 'Cambia el Tipo de Asiento
    txtDAsiento = ""
   
End Select

End Sub

Private Sub Form_Load()

Set Me.Icon = frmContenedor.Icon

vGrid.AppearanceStyle = fxGridStyle


Call sbToolBarIconos(tlb)
 
If gCntX_Arbol.ArbolActivo Then
  Call sbConsultaPlantilla(Val(gCntX_Arbol.AsientoNumr))
Else
    vEdita = False
    Call sbLimpiaPantalla
    Call sbToolBar(tlb, "activo")
End If
 
 
Call Formularios(Me)
Call RefrescaTags(Me)
 
End Sub

Private Function fxVerificaAsiento() As Boolean
Dim rsX As New ADODB.Recordset, strSQL As String
Dim vMensaje As String, lng As Long

'Verificar Periodo
'Tipo de Asiento
'CntX_Cuentas (En el Detalle)

fxVerificaAsiento = True
vMensaje = ""

If Not vEdita Then
    strSQL = "select isnull(count(*),0) as existe from CntX_Periodos where cod_contabilidad = " & gCntX_Parametros.CodigoConta _
           & " and anio = " & txtAnio & " and mes = " & txtMes & " and estado = 'P'"
    Call OpenRecordSet(rsX, strSQL, 0)
      If rsX!Existe = 0 Then vMensaje = vMensaje & vbCrLf & "- Periodo Indicado se encuentra Cerrado o No se ha creado..."
    rsX.Close
End If

strSQL = "select isnull(count(*),0) as existe from CntX_Tipos_Asientos where cod_contabilidad = " & gCntX_Parametros.CodigoConta _
       & " and tipo_asiento = '" & txtCAsiento & "'"
Call OpenRecordSet(rsX, strSQL, 0)
  If rsX!Existe = 0 Then vMensaje = vMensaje & vbCrLf & "- El tipo de Asiento Indicano no existe..."
rsX.Close


If CCur(txtDiferencia) <> 0 Then vMensaje = vMensaje & vbCrLf & "- El Asiento No se encuentra Balanceado..."
If CCur(txtDiferenciaPro) <> 0 Then vMensaje = vMensaje & vbCrLf & "- El Asiento PROYECTADO No se encuentra Balanceado..."

For lng = 1 To vGrid.MaxRows
 vGrid.Row = lng
 vGrid.Col = 1
 If vGrid.Text <> "" Then
   vGrid.Col = 2
   If vGrid.Text = "" Then
      vGrid.Col = 1
      vMensaje = vMensaje & vbCrLf & "- Cuenta " & vGrid.Text & " No Existe"
   End If
 End If
Next lng

If Len(vMensaje) > 0 Then
  fxVerificaAsiento = False
  MsgBox vMensaje, vbCritical
End If

End Function


Private Sub sbLimpiaPantalla()
vBusca = 1

txtCodPlantilla = ""
txtDesPlantilla = ""

txtCAsiento = ""
txtDAsiento = ""

txtAnio = gCntX_Parametros.PeriodoAnio
txtMes = gCntX_Parametros.PeriodoMes
txtPeriodo = fxCntX_PeriodoDesc(txtAnio, txtMes)

txtDescripcion = ""
txtDetalle = ""
txtDocumento = ""

txtCredito = 0
txtDebito = 0
txtDiferencia = 0
txtCreditosPro = 0
txtDebitosPro = 0
txtDiferenciaPro = 0

vGrid.MaxRows = 0
vGrid.MaxRows = 1
vGrid.MaxCols = 9

lblCuenta.Caption = ""
lblUnidad.Caption = ""
lblDivisa.Caption = ""

End Sub



Private Sub tlb_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim strSQL As String

Select Case UCase(Button.Key)
    Case "INSERTAR", "NUEVO"
      vEdita = False
      Call sbLimpiaPantalla
      Call sbToolBar(tlb, "edicion")
    
      txtDesPlantilla.SetFocus
    
    Case "MODIFICAR", "EDITAR"
        vEdita = True
        txtDesPlantilla.SetFocus
        Call sbToolBar(tlb, "edicion")
    
    Case "BORRAR"
        Call sbBorrar
      
    Case "GUARDAR", "SALVAR"
      Call sbGuardar
    
    Case "DESHACER"
        Call sbLimpiaPantalla
        Call sbToolBar(tlb, "nuevo")
        vEdita = True
    
    Case "CONSULTAR"
       Select Case vBusca
         Case 1, 2 'Tipo de ASiento
            If vBusca = 1 Then
                gBusquedas.Columna = "Tipo_Asiento"
                gBusquedas.Orden = "Tipo_Asiento"
            Else
                gBusquedas.Columna = "Descripcion"
                gBusquedas.Orden = "Descripcion"
            End If
            gBusquedas.Filtro = " and cod_contabilidad = " & gCntX_Parametros.CodigoConta
            gBusquedas.Consulta = "select Tipo_Asiento,descripcion from CntX_Tipos_Asientos"
            frmBusquedas.Show vbModal
            txtCAsiento = gBusquedas.Resultado
            txtDAsiento = gBusquedas.Resultado2
            
         Case 3, 4 'Codigo o Descripcion  de Plantilla
            If vBusca = 3 Then
                gBusquedas.Columna = "cod_plantilla"
                gBusquedas.Orden = "cod_plantilla"
            Else
                gBusquedas.Columna = "Descripcion"
                gBusquedas.Orden = "Descripcion"
            End If
            gBusquedas.Filtro = " and cod_contabilidad = " & gCntX_Parametros.CodigoConta
            gBusquedas.Consulta = "select cod_plantilla,descripcion from CntX_Plantilla_Asientos"
            frmBusquedas.Show vbModal
            txtCodPlantilla = gBusquedas.Resultado
            txtDesPlantilla = gBusquedas.Resultado2
            txtCodPlantilla.SetFocus
       
       End Select

    Case "REPORTES"
      
'      strSQL = "{Cntx_Asientos.cod_contabilidad} = " & gCntX_Parametros.CodigoConta _
'             & " AND {Cntx_Asientos.TIPO_ASIENTO} = '" & txtCAsiento & "' AND " _
'             & " {Cntx_Asientos.NUM_ASIENTO} = '" & txtNAsiento & "'"
'
'      Call sbCntX_Reportes("ASIENTO", strSQL)
    
    Case "AYUDA"
        frmContenedor.CD.HelpContext = Me.HelpContextID
        frmContenedor.CD.ShowHelp
    
    Case "CERRAR"
      UnLoad Me
End Select

End Sub


Private Sub sbCargaGridLocal(vGrid As Object, vGridMaxCol As Integer, strSQL As String)
Dim rs As New ADODB.Recordset, i As Integer

Me.MousePointer = vbHourglass

vGrid.MaxCols = vGridMaxCol
vGrid.MaxRows = 1

vGrid.Row = vGrid.MaxRows

rs.CursorLocation = adUseServer
Call OpenRecordSet(rs, strSQL, 0)

Do While Not rs.EOF
  vGrid.Row = vGrid.MaxRows
  For i = 1 To vGrid.MaxCols
    vGrid.Col = i
    Select Case i
      Case 1 'Cuenta
            vGrid.Text = fxCntX_CuentaFormato(True, CStr(rs.Fields(i - 1).Value))
            vGrid.CellTag = rs!Descripcion

       Case 2 ' Unidad
            vGrid.CellTag = rs!UniDes
            vGrid.Text = CStr(rs!cod_unidad)
       
       Case 3 ' Centro de Costo
            vGrid.Text = CStr(rs!cod_centro_costo)
            vGrid.CellTag = rs!CentroCosto & ""
       
       
       Case 4 ' Divisa
            vGrid.Text = CStr(rs!cod_Divisa)
            vGrid.CellTag = rs!Divisa
       
       Case 5 'Tipo de Cambio
            vGrid.CellTag = CStr(IIf(IsNull(rs!TC), 0, rs!TC))
            vGrid.Text = CStr(IIf(IsNull(rs!TC), 0, rs!TC))
      
       Case 6 'Tipo Incremento
            vGrid.Text = IIf((CStr(rs!inc_tipo) = "P"), "Porcentual", "Monto Adicional")
      
       Case 7 'Incremento valor
            vGrid.Text = CStr(rs!inc_Valor)
      
       Case 8 'debitos
            vGrid.Text = CStr(rs!Debitos)
      
       Case 9 'creditos
            vGrid.Text = CStr(rs!Creditos)
      
      Case Else
            vGrid.Text = CStr(rs.Fields(i - 1).Value)
    End Select
 
  Next i
  vGrid.MaxRows = vGrid.MaxRows + 1
  rs.MoveNext
Loop

rs.Close

Me.MousePointer = vbDefault

End Sub


Private Sub sbConsultaPlantilla(CodPlantilla As Long)
Dim rs As New ADODB.Recordset, strSQL As String

On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = "select * from CntX_Plantilla_Asientos where cod_contabilidad = " & gCntX_Parametros.CodigoConta _
       & " and cod_plantilla = " & CodPlantilla
       
Call OpenRecordSet(rs, strSQL, 0)

If Not rs.BOF And Not rs.EOF Then
  Call sbToolBar(tlb, "activo")
  vEdita = True
 
  'llenar datos en pantalla
  
  txtAnio = rs!anio_inicio
  txtMes = rs!mes_inicio
  txtPeriodo = fxCntX_PeriodoDesc(rs!anio_inicio, rs!mes_inicio)
  
  txtCAsiento.Text = rs!tipo_asiento
  txtDAsiento.Text = fxCntX_TiposAsientos("D", rs!tipo_asiento)
  
  txtCodPlantilla = rs!cod_plantilla
  txtDesPlantilla = rs!Descripcion & ""
  
  txtDescripcion = rs!asiento_descripcion & ""
  txtDetalle = rs!asiento_detalle & ""
  txtDocumento = rs!asiento_documento & ""
  txtConsecutivo = rs!Consecutivo
  
strSQL = "select A.cod_cuenta,B.descripcion,A.cod_unidad,U.descripcion as UniDes,Y.cod_divisa,Y.descripcion as Divisa" _
       & ",A.tc,inc_tipo,inc_valor,debitos,creditos,num_linea,A.cod_centro_costo,Cc.descripcion as CentroCosto" _
       & " from CntX_Plantilla_detalle A inner join CntX_Cuentas B on A.cod_cuenta = B.cod_cuenta" _
       & " and A.cod_contabilidad = B.cod_contabilidad" _
       & " inner join CntX_Unidades U on A.cod_unidad = U.cod_unidad and A.cod_contabilidad = U.cod_contabilidad" _
       & " inner join CntX_Divisas Y on A.cod_divisa = Y.cod_divisa and A.cod_contabilidad = Y.cod_contabilidad" _
       & " left join CntX_Centro_Costos Cc on A.cod_centro_costo = Cc.cod_centro_costo and A.cod_contabilidad = Cc.cod_contabilidad" _
       & " where A.cod_contabilidad = " & gCntX_Parametros.CodigoConta _
       & " and A.cod_plantilla = " & rs!cod_plantilla _
       & " order by num_linea"
   
  Call sbCargaGridLocal(vGrid, 9, strSQL)
 
  Call sbSumaDebitosCreditos

End If

rs.Close
Me.MousePointer = vbDefault

Exit Sub
vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub sbGuardar()
Dim strSQL As String, rs As New ADODB.Recordset, lng As Long

On Error GoTo vError

If fxVerificaAsiento Then
    
    If vEdita Then
      
      strSQL = "update CntX_Plantilla_Asientos set descripcion = '" & UCase(txtDesPlantilla) _
             & "',asiento_descripcion = '" & UCase(txtDescripcion) _
             & "',asiento_detalle = '" & txtDetalle _
             & "',asiento_documento = '" & txtDocumento _
             & "',anio_inicio = " & txtAnio & ",mes_inicio = " & txtMes _
             & ",tipo_asiento = '" & txtCAsiento _
             & "' where cod_contabilidad = " & gCntX_Parametros.CodigoConta _
             & " and cod_plantilla = " & txtCodPlantilla
      Call ConectionExecute(strSQL, 0)
     
      Call Bitacora("Modifica", "Plantilla Asiento : " & txtCodPlantilla & " Conta." & gCntX_Parametros.CodigoConta)
    
    Else 'Inserta
      
      'Saca Consecutivo de Plantilla
       strSQL = "select isnull(max(cod_plantilla),0) as Ultimo from CntX_Plantilla_Asientos" _
              & " where cod_contabilidad = " & gCntX_Parametros.CodigoConta
       Call OpenRecordSet(rs, strSQL, 0)
         
       strSQL = "insert into CntX_Plantilla_Asientos(cod_plantilla,tipo_asiento,cod_contabilidad,anio_inicio,mes_inicio" _
              & ",descripcion,asiento_descripcion,asiento_detalle,asiento_documento,consecutivo) values(" _
              & (rs!ultimo + 1) & ",'" & UCase(txtCAsiento) & "'," & gCntX_Parametros.CodigoConta & "," & txtAnio _
              & "," & txtMes & ",'" & UCase(txtDesPlantilla) & "','" & UCase(txtDescripcion) _
              & "','" & txtDetalle & "','" & txtDocumento & "',0)"
       Call ConectionExecute(strSQL, 0)
       
       txtCodPlantilla = (rs!ultimo + 1)
            
       rs.Close
       
        Call Bitacora("Registra", "Plantilla Asiento : " & txtCodPlantilla & " Conta." & gCntX_Parametros.CodigoConta)
        
    End If 'Si Inserta o Actualiza

        
  'Actualiza el Detalle de la Planilla
  strSQL = "delete CntX_Plantilla_detalle where cod_contabilidad = " _
         & gCntX_Parametros.CodigoConta & " and cod_plantilla = " & txtCodPlantilla

  For lng = 1 To vGrid.MaxRows
    vGrid.Row = lng
    vGrid.Col = 1
    If vGrid.Text <> "" Then
        strSQL = strSQL & Space(10) & "insert into CntX_Plantilla_detalle(cod_plantilla,cod_contabilidad" _
               & ",num_linea,cod_cuenta,cod_unidad,cod_centro_costo,cod_divisa,tc,inc_tipo,inc_valor,debitos,creditos" _
               & ") values(" & txtCodPlantilla & "," & gCntX_Parametros.CodigoConta & "," & lng & ",'"
        vGrid.Row = lng
        vGrid.Col = 1
        strSQL = strSQL & fxCntX_CuentaFormato(False, vGrid.Text) & "','"
        vGrid.Col = 2
        strSQL = strSQL & vGrid.Text & "','"
        vGrid.Col = 3
        strSQL = strSQL & vGrid.Text & "','"
        vGrid.Col = 4
        strSQL = strSQL & vGrid.Text & "',"
        vGrid.Col = 5
        strSQL = strSQL & CCur(vGrid.Text) & ",'"
        vGrid.Col = 6
        If vGrid.Text = "" Then
           strSQL = strSQL & "M',0"
        Else
           If vGrid.Text = "Porcentual" Then
               strSQL = strSQL & "P',"
           Else
               strSQL = strSQL & "M',"
           End If
           vGrid.Col = 7
           strSQL = strSQL & CCur(IIf((vGrid.Text = ""), 0, vGrid.Text)) & ","
        End If
        
        vGrid.Col = 8
        strSQL = strSQL & CCur(IIf((vGrid.Text = ""), 0, vGrid.Text)) & ","
        
        vGrid.Col = 9
        strSQL = strSQL & CCur(IIf((vGrid.Text = ""), 0, vGrid.Text)) & ")" _
      
          
     End If 'vgrid.Text <> ""
   
   Next lng
   
   'Procesa Todo el Detalle
   Call ConectionExecute(strSQL, 0)
        
        
        Call sbToolBar(tlb, "activo")
        Call sbConsultaPlantilla(txtCodPlantilla)
        
        vEdita = True
        
        MsgBox "Información guardada satisfactoriamente...", vbInformation


End If 'Verificacion del Asiento

' Call RefrescaTags(Me)

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical


End Sub

Private Sub sbBorrar()
Dim i As Integer, strSQL As String

On Error GoTo vError

i = MsgBox("Esta Seguro que desea borrar este registro", vbYesNo)

If i = vbYes Then
  strSQL = "delete CntX_Plantilla_detalle where cod_contabilidad = " _
         & gCntX_Parametros.CodigoConta & " and cod_plantilla = " & txtCodPlantilla
  Call ConectionExecute(strSQL, 0)
  
  strSQL = "delete CntX_Plantilla_Asientos where cod_contabilidad = " _
         & gCntX_Parametros.CodigoConta & " and cod_plantilla = " & txtCodPlantilla
  Call ConectionExecute(strSQL, 0)
  

  Call Bitacora("Elimina", "Plantilla Asiento : " & txtCodPlantilla & " Conta." _
                  & gCntX_Parametros.CodigoConta)

  Call sbLimpiaPantalla
  Call sbToolBar(tlb, "nuevo")
End If

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub txtAnio_Change()
On Error GoTo vError
  txtPeriodo = fxCntX_PeriodoDesc(txtAnio, txtMes)
vError:
End Sub

Private Sub txtAnio_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtDescripcion.SetFocus
End Sub

Private Sub txtCAsiento_Change()
Dim rs As New ADODB.Recordset, strSQL As String

strSQL = "select descripcion from CntX_Tipos_Asientos where cod_contabilidad = " _
       & gCntX_Parametros.CodigoConta & " and tipo_asiento = '" _
       & txtCAsiento.Text & "'"
Call OpenRecordSet(rs, strSQL, 0)
If Not rs.EOF And Not rs.BOF Then
  txtDAsiento = rs!Descripcion
End If
rs.Close

End Sub

Private Sub txtCAsiento_GotFocus()
vBusca = 1
End Sub

Private Sub txtCAsiento_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtMes.SetFocus
If KeyCode = vbKeyF4 Then Call tlb_ButtonClick(tlb.Buttons(7))
End Sub

Private Sub txtCAsiento_LostFocus()
Dim rs As New ADODB.Recordset, strSQL As String

strSQL = "select descripcion from CntX_Tipos_Asientos where cod_contabilidad = " _
       & gCntX_Parametros.CodigoConta & " and tipo_asiento = '" _
       & txtCAsiento.Text & "'"
Call OpenRecordSet(rs, strSQL, 0)
If Not rs.EOF And Not rs.BOF Then
  txtDAsiento = rs!Descripcion
End If
rs.Close

End Sub

Private Sub txtCodPlantilla_GotFocus()
vBusca = 3
End Sub

Private Sub txtDAsiento_GotFocus()
vBusca = 2
End Sub

Private Sub txtDAsiento_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtMes.SetFocus
If KeyCode = vbKeyF4 Then Call tlb_ButtonClick(tlb.Buttons(7))
End Sub

Private Sub txtDescripcion_GotFocus()
vBusca = 4
End Sub

Private Sub txtDescripcion_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtDocumento.SetFocus
End Sub

Private Sub txtDesPlantilla_GotFocus()
vBusca = 4
End Sub

Private Sub txtDesPlantilla_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtCAsiento.SetFocus
End Sub

Private Sub txtDetalle_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then vGrid.SetFocus
End Sub

Private Sub txtDocumento_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtDetalle.SetFocus
End Sub

Private Sub txtMes_Change()
On Error GoTo vError
  txtPeriodo = fxCntX_PeriodoDesc(txtAnio, txtMes)
vError:
End Sub

Private Sub txtCodPlantilla_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo vError
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then
 Call sbConsultaPlantilla(txtCodPlantilla)
 txtDesPlantilla.SetFocus
End If
If KeyCode = vbKeyF4 Then Call tlb_ButtonClick(tlb.Buttons(7))
Exit Sub
vError:
  Call sbLimpiaPantalla
End Sub


Private Function fxVerificaCuenta(strCuenta As String) As Boolean
Dim rsX As New ADODB.Recordset, strSQL As String

strSQL = "select isnull(count(*),0) as Existe from CntX_Cuentas where cod_contabilidad = " & gCntX_Parametros.CodigoConta _
       & " and cod_cuenta = '" & strCuenta & "' and acepta_movimientos = 1"

Call OpenRecordSet(rsX, strSQL, 0)
fxVerificaCuenta = IIf((rsX!Existe = 0), False, True)
rsX.Close
End Function

Private Sub sbSumaDebitosCreditos()
Dim x As Long, curValor As Currency
Dim TC As Currency

  txtDebito = 0
  txtCredito = 0
   
  txtDebitosPro = 0
  txtCreditosPro = 0
    
  
  For x = 1 To vGrid.MaxRows
    vGrid.Row = x
    vGrid.Col = 1
    If vGrid.Text <> "" Then
      'Solo en colones en pantalla, por eso no hay que convertir
      TC = 1
      
      vGrid.Col = 8
      txtDebito = CCur(txtDebito) + CCur(IIf(vGrid.Text = "", 0, vGrid.Text))
      vGrid.Col = 9
      txtCredito = CCur(txtCredito) + CCur(IIf(vGrid.Text = "", 0, vGrid.Text))
      vGrid.Col = 6
      If vGrid.Text = "Porcentual" Then
          vGrid.Col = 7
          curValor = 1 + (CCur(vGrid.Text) / 100)
          
          vGrid.Col = 8
          txtDebitosPro = CCur(txtDebitosPro) + CCur(IIf(vGrid.Text = "", 0, vGrid.Text)) * curValor
          vGrid.Col = 9
          txtCreditosPro = CCur(txtCreditosPro) + CCur(IIf(vGrid.Text = "", 0, vGrid.Text)) * curValor
      
      Else
      
          vGrid.Col = 7
          curValor = CCur(vGrid.Text)
          
          vGrid.Col = 8
          txtDebitosPro = CCur(txtDebitosPro) + CCur(IIf(vGrid.Text = "", 0, vGrid.Text)) + curValor
          vGrid.Col = 9
          txtCreditosPro = CCur(txtCreditosPro) + CCur(IIf(vGrid.Text = "", 0, vGrid.Text)) + curValor
      
      End If
     End If 'vGrid.text <> ""
      
  Next x
  txtDiferencia = txtDebito - txtCredito
  txtDebito = Format(txtDebito, "Standard")
  txtCredito = Format(txtCredito, "Standard")
  txtDiferencia = Format(txtDiferencia, "Standard")

  txtDiferenciaPro = txtDebitosPro - txtCreditosPro
  txtDebitosPro = Format(txtDebitosPro, "Standard")
  txtCreditosPro = Format(txtCreditosPro, "Standard")
  txtDiferenciaPro = Format(txtDiferenciaPro, "Standard")
End Sub

Private Sub txtMes_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtAnio.SetFocus
End Sub

Private Sub vGrid_KeyDown(KeyCode As Integer, Shift As Integer)
Dim i As Variant, lng As Long, vTemp(9) As Variant, x As Integer
Dim vTempo As String

If KeyCode = vbKeyDelete Then
  
  vGrid.Row = vGrid.ActiveRow
 
  For lng = vGrid.ActiveRow To vGrid.MaxRows
     vGrid.Row = lng + 1
     For x = 1 To vGrid.MaxCols
        vGrid.Col = x
        vTemp(x) = vGrid.Text
     Next x
     
     vGrid.Row = lng
     For x = 1 To vGrid.MaxCols
       vGrid.Col = x
       vGrid.Text = vTemp(x)
     Next x
  Next lng
  vGrid.MaxRows = vGrid.MaxRows - 1
  If vGrid.MaxRows = 0 Then vGrid.MaxRows = 1
  
  Call sbSumaDebitosCreditos
  
  
End If

'Consulta Cuenta
If KeyCode = vbKeyF4 And vGrid.ActiveCol = 1 Then
  frmCntX_ConsultaCuentas.Show vbModal
  vGrid.Col = vGrid.ActiveCol
  vGrid.Row = vGrid.ActiveRow
  vGrid.Text = gCuenta
End If

'Consulta Unidad
If KeyCode = vbKeyF4 And vGrid.ActiveCol = 2 Then
  gBusquedas.Resultado = ""
  gBusquedas.Resultado2 = ""
  gBusquedas.Columna = "descripcion"
  gBusquedas.Orden = "descripcion"
  gBusquedas.Filtro = " and cod_contabilidad = " & gCntX_Parametros.CodigoConta
  gBusquedas.Consulta = "select cod_unidad,descripcion from CntX_Unidades"
  frmBusquedas.Show vbModal
    
  vGrid.Col = vGrid.ActiveCol
  vGrid.Row = vGrid.ActiveRow
  
  vGrid.Text = gBusquedas.Resultado
  vGrid.CellTag = gBusquedas.Resultado2
  
End If


'Consulta Centro de Costo
If KeyCode = vbKeyF4 And vGrid.ActiveCol = 3 Then
  vGrid.Row = vGrid.ActiveRow
  vGrid.Col = 2
  vTempo = vGrid.Text
  
  gBusquedas.Resultado = ""
  gBusquedas.Resultado2 = ""
  gBusquedas.Columna = "descripcion"
  gBusquedas.Orden = "descripcion"
  gBusquedas.Filtro = " and cod_contabilidad = " & gCntX_Parametros.CodigoConta & " and cod_centro_costo in(select cod_centro_costo" _
                    & " from cntX_unidades_cc where cod_unidad = '" & vTempo & "' and cod_contabilidad = " & gCntX_Parametros.CodigoConta & ")"
  gBusquedas.Consulta = "select cod_centro_costo,descripcion from CntX_Centro_Costos"
  frmBusquedas.Show vbModal
    
  vGrid.Col = vGrid.ActiveCol
  vGrid.Row = vGrid.ActiveRow
  
  vGrid.Text = gBusquedas.Resultado
  vGrid.CellTag = gBusquedas.Resultado2
  vGrid.TextTip = TextTipFixed
  vGrid.CellNote = vGrid.CellTag
  
End If


'Consulta Divisa
If KeyCode = vbKeyF4 And vGrid.ActiveCol = 4 Then
  gBusquedas.Resultado = ""
  gBusquedas.Resultado2 = ""
  gBusquedas.Columna = "descripcion"
  gBusquedas.Orden = "descripcion"
  gBusquedas.Filtro = " and cod_contabilidad = " & gCntX_Parametros.CodigoConta
  gBusquedas.Consulta = "select cod_divisa,descripcion from CntX_Divisas"
  frmBusquedas.Show vbModal
    
  vGrid.Col = vGrid.ActiveCol
  vGrid.Row = vGrid.ActiveRow
  
  vGrid.Text = gBusquedas.Resultado
  vGrid.CellTag = gBusquedas.Resultado2
End If


If (KeyCode = vbKeyReturn Or KeyCode = vbKeyTab) Then
    vGrid.Col = vGrid.ActiveCol
    vGrid.Row = vGrid.ActiveRow
    
    Select Case vGrid.ActiveCol
       Case 1
        vGrid.Text = fxCntX_CuentaFormato(True, vGrid.Text)
        i = fxCntX_CuentaFormato(False, vGrid.Text)
        If fxVerificaCuenta(CStr(i)) Then
          lblCuenta.Caption = fxCntX_Cuenta("D", CStr(i))
          vGrid.CellTag = lblCuenta.Caption
        Else
          MsgBox "Cuenta no es válida : " & vbCrLf & " - No Existe o No Acepta Movimientos" _
                 & vbCrLf & " - VERIFIQUE O MODIFIQUE EN EL CATALAGO DE Cuentas", vbCritical
        End If
        
      Case 2
        'Buscar la unidad
        If fxCntx_UnidadVerifica(vGrid.Text) Then
          vGrid.CellTag = fxCntX_Unidad("D", vGrid.Text)
          vUltimos.Unidad = vGrid.Text
          vUltimos.UnidadDesc = vGrid.CellTag
        Else
          MsgBox "La unidad de negocio no es válida : " & vbCrLf & " - No Existe...", vbCritical
        End If
      
      Case 3 'Verificar el Centro de Costo
        vGrid.Col = 2
        vTempo = vGrid.Text
        vGrid.Col = 3
        
        If fxCntX_CentroCostoVerifica(vGrid.Text, vTempo) Then
          vGrid.TextTip = TextTipFixed
          vGrid.CellTag = fxCntX_CentroCosto("D", vGrid.Text)
          vGrid.CellNote = vGrid.CellTag
          
          vUltimos.CC = vGrid.Text
          vUltimos.CCDesc = vGrid.CellTag
        Else
          MsgBox "El Centro de Costo no es válido y no puede ser utilizada por esta unidad: " & vbCrLf & " - No Existe...", vbCritical
        End If
      
      Case 4 'Divisa
      
        If fxCntX_DivisaVerifica(vGrid.Text) Then
          vGrid.CellTag = fxCntX_Divisas("D", vGrid.Text)
          vUltimos.Divisa = vGrid.Text
          vUltimos.DivisaDesc = vGrid.CellTag
          'Tipo de Cambio
        Else
          MsgBox "La Divisa no es válida : " & vbCrLf & " - No Existe...", vbCritical
        End If
        
      Case 6
        vUltimos.Tipo = vGrid.Text
        
      Case 7
        vUltimos.Valor = vGrid.Text
        
      Case 8 'Debe
        If Val(vGrid.Text) > 0 Then
            vGrid.Col = vGrid.ActiveCol + 1
            vGrid.Row = vGrid.ActiveRow
            vGrid.Text = 0
        
            Call sbSumaDebitosCreditos
            
        End If
      
      Case 9 'Haber
        If Val(vGrid.Text) > 0 Then
            vGrid.Col = vGrid.ActiveCol - 1
            vGrid.Row = vGrid.ActiveRow
            vGrid.Text = 0
        
            Call sbSumaDebitosCreditos
        
        End If
      
        If vGrid.MaxRows = vGrid.Row Then
            vGrid.MaxRows = vGrid.MaxRows + 1
            vGrid.Row = vGrid.MaxRows
            vGrid.Col = 2
            vGrid.Text = vUltimos.Unidad
            vGrid.CellTag = vUltimos.UnidadDesc
            
            vGrid.Col = 3
            vGrid.TextTip = TextTipFixed
            vGrid.Text = vUltimos.CC
            vGrid.CellTag = vUltimos.CCDesc
            vGrid.CellNote = vGrid.CellTag
          
            vGrid.Col = 4
            vGrid.Text = vUltimos.Divisa
            vGrid.CellTag = vUltimos.DivisaDesc
            
            vGrid.Col = 6
            vGrid.Text = vUltimos.Tipo
            vGrid.Col = 7
            vGrid.Text = vUltimos.Valor
        End If
    
    End Select
End If

If KeyCode = vbKeyInsert Then
    vGrid.MaxRows = vGrid.MaxRows + 1
    vGrid.InsertRows vGrid.ActiveRow, 1
    vGrid.Row = vGrid.ActiveRow
    vGrid.Row = vGrid.MaxRows
    
    vGrid.Col = 2
    vGrid.Text = vUltimos.Unidad
    vGrid.CellTag = vUltimos.UnidadDesc
    
    vGrid.Col = 3
    vGrid.TextTip = TextTipFixed
    vGrid.Text = vUltimos.CC
    vGrid.CellTag = vUltimos.CCDesc
    vGrid.CellNote = vGrid.CellTag
    
    vGrid.Col = 3
    vGrid.Text = vUltimos.Divisa
    vGrid.CellTag = vUltimos.DivisaDesc
    vGrid.Col = 6
    vGrid.Text = vUltimos.Tipo
    vGrid.Col = 7
    vGrid.Text = vUltimos.Valor
End If
End Sub


Private Sub vGrid_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
Dim vCuenta As String, vMoneda As String, vMov As Currency, i As Byte

On Error GoTo vError

vGrid.Row = Row


If NewCol = 8 Or NewCol = 9 Then
    vGrid.Col = 4
    lblCuenta.Tag = IIf(fxCntX_DivisaBase(vGrid.Text), "S", "N")
End If

If Col = 4 Then
    'Verificar Tipo de Cambio
    vGrid.Col = 4
    vMoneda = vGrid.Text
    
    vGrid.Col = 5
    If vGrid.Text = "" Or vGrid.Text = "0.00" Then
      vGrid.Col = 1
      vCuenta = fxCntX_CuentaFormato(False, vGrid.Text, 0)
      vGrid.Col = 5
      vGrid.Text = fxCntX_TipoCambio(vMoneda, vCuenta, fxFechaServidor)
    End If
End If

vGrid.Row = NewRow
vGrid.Col = 1
lblCuenta.Caption = vGrid.CellTag
vGrid.Col = 2
lblUnidad.Caption = vGrid.CellTag
vGrid.Col = 4
lblDivisa.Caption = vGrid.CellTag

vGrid.Col = 8
vMov = CCur(vGrid.Text)
vGrid.Col = 9
vMov = vMov + CCur(vGrid.Text)
vGrid.Col = 5
lblDivisa.Caption = lblDivisa.Caption & " [" & Format(vMov / CCur(vGrid.Text), "Standard") & "]"

vError:

End Sub

