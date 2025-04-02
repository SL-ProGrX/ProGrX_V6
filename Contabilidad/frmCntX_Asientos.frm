VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#20.3#0"; "Codejock.Controls.v20.3.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#20.3#0"; "Codejock.ShortcutBar.v20.3.0.ocx"
Begin VB.Form frmCntX_Asientos 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Asientos"
   ClientHeight    =   8145
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12825
   HelpContextID   =   2002
   Icon            =   "frmCntX_Asientos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8145
   ScaleWidth      =   12825
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.PushButton btnProcesos 
      Height          =   330
      Index           =   0
      Left            =   5520
      TabIndex        =   46
      Top             =   40
      Width           =   1455
      _Version        =   1310723
      _ExtentX        =   2566
      _ExtentY        =   582
      _StockProps     =   79
      Caption         =   "Cuentas"
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
      FlatStyle       =   -1  'True
      UseVisualStyle  =   -1  'True
      TextAlignment   =   1
      Appearance      =   17
      Picture         =   "frmCntX_Asientos.frx":6852
      ImageAlignment  =   0
   End
   Begin VB.Frame fraCopia 
      BorderStyle     =   0  'None
      Height          =   5532
      Left            =   12480
      TabIndex        =   12
      Top             =   2280
      Visible         =   0   'False
      Width           =   12612
      Begin XtremeSuiteControls.CheckBox chkCopiaAsReversion 
         Height          =   255
         Left            =   5400
         TabIndex        =   45
         Top             =   960
         Width           =   2295
         _Version        =   1310723
         _ExtentX        =   4048
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Asiento de Reversión ?"
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
      End
      Begin XtremeSuiteControls.PushButton btnCopiar 
         Height          =   615
         Index           =   0
         Left            =   6960
         TabIndex        =   43
         Top             =   4680
         Width           =   1575
         _Version        =   1310723
         _ExtentX        =   2778
         _ExtentY        =   1085
         _StockProps     =   79
         Caption         =   "Copiar"
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
         Picture         =   "frmCntX_Asientos.frx":6F5A
      End
      Begin XtremeSuiteControls.CheckBox chkCopiaDetalles 
         Height          =   615
         Left            =   5280
         TabIndex        =   35
         Top             =   3000
         Width           =   4815
         _Version        =   1310723
         _ExtentX        =   8488
         _ExtentY        =   1080
         _StockProps     =   79
         Caption         =   "Cambiar: Documento y Detalle en las líneas del Asiento   "
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
         TextAlignment   =   1
         Appearance      =   16
         Alignment       =   1
      End
      Begin XtremeSuiteControls.FlatEdit txtCopia_Detalle 
         Height          =   315
         Left            =   7680
         TabIndex        =   36
         Top             =   4080
         Width           =   2415
         _Version        =   1310723
         _ExtentX        =   4254
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
         Appearance      =   2
      End
      Begin XtremeSuiteControls.FlatEdit txtCopia_Documento 
         Height          =   315
         Left            =   7680
         TabIndex        =   37
         Top             =   3720
         Width           =   2415
         _Version        =   1310723
         _ExtentX        =   4254
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
         Appearance      =   2
      End
      Begin XtremeSuiteControls.DateTimePicker dtpCopia_Fecha 
         Height          =   315
         Left            =   8640
         TabIndex        =   38
         Top             =   960
         Width           =   1455
         _Version        =   1310723
         _ExtentX        =   2561
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
      Begin XtremeSuiteControls.FlatEdit txtCopia_Descripcion 
         Height          =   315
         Left            =   2880
         TabIndex        =   39
         Top             =   1320
         Width           =   7215
         _Version        =   1310723
         _ExtentX        =   12726
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
         Appearance      =   2
      End
      Begin XtremeSuiteControls.FlatEdit txtCopia_NAsiento 
         Height          =   312
         Left            =   2880
         TabIndex        =   40
         Top             =   960
         Width           =   2412
         _Version        =   1310723
         _ExtentX        =   4254
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
         Appearance      =   2
      End
      Begin XtremeSuiteControls.FlatEdit txtCopia_Referencia 
         Height          =   315
         Left            =   2880
         TabIndex        =   41
         Top             =   1680
         Width           =   7215
         _Version        =   1310723
         _ExtentX        =   12726
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
         Appearance      =   2
      End
      Begin XtremeSuiteControls.FlatEdit txtCopia_Notas 
         Height          =   915
         Left            =   2880
         TabIndex        =   42
         Top             =   2040
         Width           =   7215
         _Version        =   1310723
         _ExtentX        =   12726
         _ExtentY        =   1614
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
      Begin XtremeSuiteControls.PushButton btnCopiar 
         Height          =   615
         Index           =   1
         Left            =   8520
         TabIndex        =   44
         Top             =   4680
         Width           =   1575
         _Version        =   1310723
         _ExtentX        =   2778
         _ExtentY        =   1085
         _StockProps     =   79
         Caption         =   "Cerrar"
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
         Picture         =   "frmCntX_Asientos.frx":7720
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption2 
         Height          =   492
         Left            =   120
         TabIndex        =   34
         Top             =   120
         Width           =   12372
         _Version        =   1310723
         _ExtentX        =   21823
         _ExtentY        =   868
         _StockProps     =   14
         Caption         =   "Copiar el asiento  a:"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   11.93
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SubItemCaption  =   -1  'True
         VisualTheme     =   6
         Alignment       =   1
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Referencia"
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
         Left            =   1680
         TabIndex        =   20
         Top             =   1680
         Width           =   1368
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Detalle"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6120
         TabIndex        =   18
         Top             =   4080
         Width           =   1545
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Nº Documento"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6120
         TabIndex        =   17
         Top             =   3720
         Width           =   1545
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Notas"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   312
         Index           =   2
         Left            =   1680
         TabIndex        =   16
         Top             =   2040
         Width           =   1308
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   7530
         TabIndex        =   15
         Top             =   960
         Width           =   945
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Nº Asiento"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   312
         Left            =   1680
         TabIndex        =   14
         Top             =   960
         Width           =   1188
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Descripción"
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
         Left            =   1680
         TabIndex        =   13
         Top             =   1320
         Width           =   1368
      End
   End
   Begin MSComCtl2.FlatScrollBar FlatScrollBar 
      Height          =   255
      Left            =   9480
      TabIndex        =   9
      Top             =   600
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   450
      _Version        =   393216
      Arrows          =   65536
      Orientation     =   1638401
   End
   Begin MSComctlLib.StatusBar stBar 
      Align           =   2  'Align Bottom
      Height          =   285
      Left            =   0
      TabIndex        =   8
      Top             =   7860
      Width           =   12825
      _ExtentX        =   22622
      _ExtentY        =   503
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   4304
            MinWidth        =   4304
            Object.ToolTipText     =   "Asiento Creado Por ..."
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   4304
            MinWidth        =   4304
            Object.ToolTipText     =   "Asiento Modificado Por ..."
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   4304
            MinWidth        =   4304
            Object.ToolTipText     =   "Asiento Aplicado Por..."
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   4304
            MinWidth        =   4304
            Object.ToolTipText     =   "Asiento Autorizado Por.."
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.ToolTipText     =   "Módulo que Genera el Asiento"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin FPSpreadADO.fpSpread vGrid 
      Height          =   4455
      Left            =   120
      TabIndex        =   11
      Top             =   2760
      Width           =   12615
      _Version        =   524288
      _ExtentX        =   22246
      _ExtentY        =   7853
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
      SpreadDesigner  =   "frmCntX_Asientos.frx":7EED
      VScrollSpecial  =   -1  'True
      VScrollSpecialType=   2
      AppearanceStyle =   1
   End
   Begin XtremeSuiteControls.FlatEdit txtReferencia 
      Height          =   315
      Left            =   1200
      TabIndex        =   21
      Top             =   1320
      Width           =   11415
      _Version        =   1310723
      _ExtentX        =   20129
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
      Appearance      =   2
   End
   Begin XtremeSuiteControls.FlatEdit txtDescripcion 
      Height          =   315
      Left            =   6240
      TabIndex        =   22
      Top             =   960
      Width           =   6375
      _Version        =   1310723
      _ExtentX        =   11239
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
      Appearance      =   2
   End
   Begin XtremeSuiteControls.FlatEdit txtCAsiento 
      Height          =   315
      Left            =   1200
      TabIndex        =   23
      Top             =   600
      Width           =   975
      _Version        =   1310723
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
      Appearance      =   2
   End
   Begin XtremeSuiteControls.FlatEdit txtDAsiento 
      Height          =   315
      Left            =   2160
      TabIndex        =   24
      Top             =   600
      Width           =   2535
      _Version        =   1310723
      _ExtentX        =   4466
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
   Begin XtremeSuiteControls.FlatEdit txtNAsiento 
      Height          =   315
      Left            =   6240
      TabIndex        =   25
      Top             =   600
      Width           =   2415
      _Version        =   1310723
      _ExtentX        =   4254
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
      Appearance      =   2
   End
   Begin XtremeSuiteControls.DateTimePicker dtpAsientoFecha 
      Height          =   315
      Left            =   11160
      TabIndex        =   26
      Top             =   600
      Width           =   1455
      _Version        =   1310723
      _ExtentX        =   2561
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
   Begin XtremeSuiteControls.FlatEdit txtMes 
      Height          =   315
      Left            =   1200
      TabIndex        =   27
      Top             =   960
      Width           =   375
      _Version        =   1310723
      _ExtentX        =   656
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
      Appearance      =   2
   End
   Begin XtremeSuiteControls.FlatEdit txtPeriodo 
      Height          =   315
      Left            =   2160
      TabIndex        =   28
      Top             =   960
      Width           =   2535
      _Version        =   1310723
      _ExtentX        =   4466
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
   Begin XtremeSuiteControls.FlatEdit txtAnio 
      Height          =   315
      Left            =   1560
      TabIndex        =   29
      Top             =   960
      Width           =   615
      _Version        =   1310723
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
      Appearance      =   2
   End
   Begin XtremeSuiteControls.FlatEdit txtNotas 
      Height          =   555
      Left            =   1200
      TabIndex        =   30
      Top             =   1680
      Width           =   11415
      _Version        =   1310723
      _ExtentX        =   20135
      _ExtentY        =   979
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
   Begin XtremeSuiteControls.FlatEdit txtDebito 
      Height          =   315
      Left            =   8640
      TabIndex        =   31
      Top             =   7440
      Width           =   1935
      _Version        =   1310723
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
      Alignment       =   1
      Locked          =   -1  'True
      Appearance      =   2
   End
   Begin XtremeSuiteControls.FlatEdit txtCredito 
      Height          =   315
      Left            =   10560
      TabIndex        =   32
      Top             =   7440
      Width           =   1935
      _Version        =   1310723
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
      Alignment       =   1
      Locked          =   -1  'True
      Appearance      =   2
   End
   Begin XtremeSuiteControls.FlatEdit txtDiferencia 
      Height          =   315
      Left            =   5040
      TabIndex        =   33
      Top             =   7440
      Width           =   1695
      _Version        =   1310723
      _ExtentX        =   2984
      _ExtentY        =   550
      _StockProps     =   77
      ForeColor       =   255
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
   End
   Begin MSComctlLib.ImageList ImageListMenu 
      Left            =   11520
      Top             =   240
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCntX_Asientos.frx":8643
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCntX_Asientos.frx":EEA5
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCntX_Asientos.frx":15707
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCntX_Asientos.frx":1BF69
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCntX_Asientos.frx":1C61D
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCntX_Asientos.frx":1CDFB
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCntX_Asientos.frx":1D790
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCntX_Asientos.frx":1E11D
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCntX_Asientos.frx":1EAEA
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCntX_Asientos.frx":1F2EF
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCntX_Asientos.frx":1FABC
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin XtremeSuiteControls.PushButton btnProcesos 
      Height          =   330
      Index           =   1
      Left            =   6960
      TabIndex        =   47
      Top             =   40
      Width           =   1455
      _Version        =   1310723
      _ExtentX        =   2566
      _ExtentY        =   582
      _StockProps     =   79
      Caption         =   "Mayorizar"
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
      FlatStyle       =   -1  'True
      UseVisualStyle  =   -1  'True
      TextAlignment   =   1
      Appearance      =   17
      Picture         =   "frmCntX_Asientos.frx":201A4
      ImageAlignment  =   0
   End
   Begin XtremeSuiteControls.PushButton btnProcesos 
      Height          =   330
      Index           =   2
      Left            =   8400
      TabIndex        =   48
      Top             =   40
      Width           =   1455
      _Version        =   1310723
      _ExtentX        =   2566
      _ExtentY        =   582
      _StockProps     =   79
      Caption         =   "Reversar"
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
      FlatStyle       =   -1  'True
      UseVisualStyle  =   -1  'True
      TextAlignment   =   1
      Appearance      =   17
      Picture         =   "frmCntX_Asientos.frx":208CB
      ImageAlignment  =   0
   End
   Begin XtremeSuiteControls.PushButton btnProcesos 
      Height          =   330
      Index           =   3
      Left            =   9840
      TabIndex        =   49
      Top             =   40
      Width           =   1455
      _Version        =   1310723
      _ExtentX        =   2566
      _ExtentY        =   582
      _StockProps     =   79
      Caption         =   "Autorizar"
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
      FlatStyle       =   -1  'True
      UseVisualStyle  =   -1  'True
      TextAlignment   =   1
      Appearance      =   17
      Picture         =   "frmCntX_Asientos.frx":20FCB
      ImageAlignment  =   0
   End
   Begin XtremeSuiteControls.PushButton btnProcesos 
      Height          =   330
      Index           =   4
      Left            =   11280
      TabIndex        =   50
      Top             =   40
      Width           =   1455
      _Version        =   1310723
      _ExtentX        =   2566
      _ExtentY        =   582
      _StockProps     =   79
      Caption         =   "Histórico"
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
      FlatStyle       =   -1  'True
      UseVisualStyle  =   -1  'True
      TextAlignment   =   1
      Appearance      =   17
      Picture         =   "frmCntX_Asientos.frx":215E7
      ImageAlignment  =   0
   End
   Begin XtremeSuiteControls.PushButton btnBarra 
      Height          =   330
      Index           =   0
      Left            =   1200
      TabIndex        =   51
      ToolTipText     =   "Nuevo"
      Top             =   40
      Width           =   1095
      _Version        =   1310723
      _ExtentX        =   1931
      _ExtentY        =   582
      _StockProps     =   79
      Caption         =   "Nuevo"
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FlatStyle       =   -1  'True
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Picture         =   "frmCntX_Asientos.frx":21D00
      ImageAlignment  =   4
   End
   Begin XtremeSuiteControls.PushButton btnBarra 
      Height          =   330
      Index           =   1
      Left            =   2280
      TabIndex        =   52
      ToolTipText     =   "Editar"
      Top             =   40
      Width           =   375
      _Version        =   1310723
      _ExtentX        =   661
      _ExtentY        =   582
      _StockProps     =   79
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
      FlatStyle       =   -1  'True
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Picture         =   "frmCntX_Asientos.frx":22332
      ImageAlignment  =   6
   End
   Begin XtremeSuiteControls.PushButton btnBarra 
      Height          =   330
      Index           =   2
      Left            =   2640
      TabIndex        =   53
      ToolTipText     =   "Eliminar"
      Top             =   40
      Width           =   375
      _Version        =   1310723
      _ExtentX        =   661
      _ExtentY        =   582
      _StockProps     =   79
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
      FlatStyle       =   -1  'True
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Picture         =   "frmCntX_Asientos.frx":2292D
      ImageAlignment  =   6
   End
   Begin XtremeSuiteControls.PushButton btnBarra 
      Height          =   330
      Index           =   3
      Left            =   3120
      TabIndex        =   54
      ToolTipText     =   "Guardar"
      Top             =   45
      Width           =   375
      _Version        =   1310723
      _ExtentX        =   661
      _ExtentY        =   582
      _StockProps     =   79
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
      FlatStyle       =   -1  'True
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Picture         =   "frmCntX_Asientos.frx":22ED1
      ImageAlignment  =   6
   End
   Begin XtremeSuiteControls.PushButton btnBarra 
      Height          =   330
      Index           =   4
      Left            =   3480
      TabIndex        =   55
      ToolTipText     =   "Deshacer"
      Top             =   45
      Width           =   375
      _Version        =   1310723
      _ExtentX        =   661
      _ExtentY        =   582
      _StockProps     =   79
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
      FlatStyle       =   -1  'True
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Picture         =   "frmCntX_Asientos.frx":23602
      ImageAlignment  =   6
   End
   Begin XtremeSuiteControls.PushButton btnBarra 
      Height          =   330
      Index           =   5
      Left            =   3960
      TabIndex        =   56
      ToolTipText     =   "Reporte"
      Top             =   45
      Width           =   375
      _Version        =   1310723
      _ExtentX        =   661
      _ExtentY        =   582
      _StockProps     =   79
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
      FlatStyle       =   -1  'True
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Picture         =   "frmCntX_Asientos.frx":23D02
      ImageAlignment  =   6
   End
   Begin XtremeSuiteControls.PushButton btnBarra 
      Height          =   330
      Index           =   6
      Left            =   4320
      TabIndex        =   57
      ToolTipText     =   "Consultas"
      Top             =   45
      Visible         =   0   'False
      Width           =   375
      _Version        =   1310723
      _ExtentX        =   661
      _ExtentY        =   582
      _StockProps     =   79
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
      FlatStyle       =   -1  'True
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Picture         =   "frmCntX_Asientos.frx":24409
      ImageAlignment  =   6
   End
   Begin XtremeShortcutBar.ShortcutCaption lblCuenta 
      Height          =   375
      Left            =   120
      TabIndex        =   60
      Top             =   2280
      Width           =   4935
      _Version        =   1310723
      _ExtentX        =   8705
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
   Begin XtremeShortcutBar.ShortcutCaption lblUnidad 
      Height          =   375
      Left            =   5040
      TabIndex        =   59
      Top             =   2280
      Width           =   4215
      _Version        =   1310723
      _ExtentX        =   7435
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
   Begin XtremeShortcutBar.ShortcutCaption lblDivisa 
      Height          =   375
      Left            =   9240
      TabIndex        =   58
      Top             =   2280
      Width           =   3495
      _Version        =   1310723
      _ExtentX        =   6165
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
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Referencia"
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
      Left            =   120
      TabIndex        =   19
      Top             =   1320
      Width           =   1005
   End
   Begin VB.Image imgCopia 
      Appearance      =   0  'Flat
      Height          =   240
      Left            =   9090
      Picture         =   "frmCntX_Asientos.frx":24B09
      Stretch         =   -1  'True
      ToolTipText     =   "Copiar Asiento"
      Top             =   600
      Width           =   240
   End
   Begin VB.Image imgBusquedaAdv 
      Height          =   240
      Left            =   8760
      Picture         =   "frmCntX_Asientos.frx":252BF
      Stretch         =   -1  'True
      ToolTipText     =   "Busqueda Avanzada"
      Top             =   600
      Width           =   240
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Notas"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   1
      Left            =   120
      TabIndex        =   10
      Top             =   1680
      Width           =   945
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
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
      Height          =   255
      Left            =   7230
      TabIndex        =   7
      Top             =   7470
      Width           =   1245
   End
   Begin VB.Label lsblr 
      Alignment       =   1  'Right Justify
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
      Height          =   255
      Left            =   3630
      TabIndex        =   6
      Top             =   7470
      Width           =   1275
   End
   Begin VB.Label lblAsientoEstado 
      Caption         =   "Estado del Asiento."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   6840
      Width           =   3795
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Descripción"
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
      Left            =   5010
      TabIndex        =   4
      Top             =   960
      Width           =   1005
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Nº Asiento"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   5040
      TabIndex        =   3
      Top             =   600
      Width           =   1185
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Período"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   945
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   945
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   10500
      TabIndex        =   0
      Top             =   600
      Width           =   585
   End
End
Attribute VB_Name = "frmCntX_Asientos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Type xUltimos
   Anio        As Long
   Mes         As Integer
   TipoAsiento As String
   NumAsiento  As String
   AplAsiento  As Integer
   Detalle     As String
   Documento   As String
   fecha       As Date
   Divisa      As String
   DivisaDesc  As String
   Unidad      As String
   UnidadDesc  As String
   CC          As String
   CCDesc      As String
   Cta         As String
   CtaDesc     As String
   TS          As String
End Type

Dim vEdita As Boolean, vBusca As Integer, vUltimos As xUltimos, vForaneo As String
Dim vScroll As Boolean

Const Id_TaskItem_DatosPersonales = 0
Const Id_TaskItem_RelacionLaboral = 1
Const Id_TaskItem_Redes = 2

Const Id_TaskItem_Telefonos = 3
Const Id_TaskItem_Beneficiarios = 4
Const Id_TaskItem_Cuentas = 5
Const Id_TaskItem_Nombramientos = 6
Const Id_TaskItem_Ingresos = 7
Const Id_TaskItem_Renuncias = 8
Const Id_TaskItem_Liquidaciones = 9
Const Id_TaskItem_Bloqueos = 10
Const Id_TaskItem_Tarjetas = 11
Const Id_TaskItem_Canales = 12
Const Id_TaskItem_Preferencias = 13
Const Id_TaskItem_Bienes = 14
Const Id_TaskItem_Escolaridad = 15
Const Id_TaskItem_Direcciones = 16
Const Id_TaskItem_Salarios = 17
Const Id_TaskItem_Motivos = 18

Private Sub sbBarra_Accion(pAccion As String)

btnBarra.Item(0).Enabled = False 'Nuevo
btnBarra.Item(1).Enabled = False 'Editar
btnBarra.Item(2).Enabled = False 'Borrar
btnBarra.Item(3).Enabled = False 'Guardar
btnBarra.Item(4).Enabled = False 'Deshacer
btnBarra.Item(5).Enabled = True 'Reporte

Select Case UCase(pAccion)
    Case "NUEVO"
        btnBarra.Item(0).Enabled = True 'Nuevo
    
    Case "EDITAR", "EDICION"
    
        btnBarra.Item(3).Enabled = True 'Guardar
        btnBarra.Item(4).Enabled = True 'Deshacer
    
    Case "ACTIVO"
        btnBarra.Item(0).Enabled = True 'Nuevo
        btnBarra.Item(1).Enabled = True 'Editar
        btnBarra.Item(2).Enabled = True 'Borrar
End Select

End Sub



Private Sub sbRefrescaInformacion()
Dim strResultado As String

On Error GoTo vError

txtAnio = Val(txtAnio)
dtpAsientoFecha = CDate(txtAnio & "/" & txtMes & "/01")
  Select Case Val(txtMes)
    Case 1
        strResultado = "ENERO DEL " & txtAnio
    Case 2
        strResultado = "FEBRERO DEL " & txtAnio
    Case 3
        strResultado = "MARZO DEL " & txtAnio
    Case 4
        strResultado = "ABRIL DEL " & txtAnio
    Case 5
        strResultado = "MAYO DEL " & txtAnio
    Case 6
        strResultado = "JUNIO DEL " & txtAnio
    Case 7
        strResultado = "JULIO DEL " & txtAnio
    Case 8
        strResultado = "AGOSTO DEL " & txtAnio
    Case 9
        strResultado = "SETIEMBRE DEL " & txtAnio
    Case 10
        strResultado = "OCTUBRE DEL " & txtAnio
    Case 11
        strResultado = "NOVIEMBRE DEL " & txtAnio
    Case 12
        strResultado = "DICIEMBRE DEL " & txtAnio
  End Select

  txtPeriodo = strResultado

Exit Sub

vError:
End Sub

Private Sub btnBarra_Click(Index As Integer)
Dim strSQL As String, i As Byte

Select Case Index
    Case 0 'INSERTAR", "NUEVO"
      vEdita = False
      Call sbLimpiaPantalla
      txtCAsiento.SetFocus
      Call sbBarra_Accion("edicion")
      
    
    Case 1 '"MODIFICAR", "EDITAR"
      If vUltimos.AplAsiento = 0 Then
        vEdita = True
        txtDescripcion.SetFocus
        Call sbBarra_Accion("edicion")
      Else
        MsgBox "Este Asiento ya Fue Aplicado (No se puede Modificar/Borrar)", vbInformation
      End If
    
    Case 2 '"BORRAR"
      If vUltimos.AplAsiento = 0 Then
        Call sbBorrar
      Else
        MsgBox "Este Asiento ya Fue Aplicado (No se puede Modificar/Borrar)", vbInformation
      End If
      
    Case 3 '"GUARDAR", "SALVAR"
      Call sbGuardar
    
    Case 4 '"DESHACER"
        Call sbLimpiaPantalla
        Call sbBarra_Accion("nuevo")
        vEdita = True
    
    Case 6 ' "CONSULTAR"
       Select Case vBusca
         Case 1 'Tipo de ASiento
            gBusquedas.Columna = "Tipo_Asiento"
            gBusquedas.Orden = "Tipo_Asiento"
            gBusquedas.Filtro = " and cod_contabilidad = " & gCntX_Parametros.CodigoConta
            gBusquedas.Consulta = "select Tipo_Asiento,descripcion from CntX_Tipos_Asientos"
            frmBusquedas.Show vbModal
            txtCAsiento = gBusquedas.Resultado
         Case 2 'Descripcion del Tipo de Asiento
            gBusquedas.Columna = "Descripcion"
            gBusquedas.Orden = "Descripcion"
            gBusquedas.Filtro = " and cod_contabilidad = " & gCntX_Parametros.CodigoConta
            gBusquedas.Consulta = "select Tipo_Asiento,descripcion from CntX_Tipos_Asientos"
            frmBusquedas.Show vbModal
            txtCAsiento = gBusquedas.Resultado
         Case 3 'Numero de Asiento
            gBusquedas.Columna = "Num_Asiento"
            gBusquedas.Orden = "Num_Asiento"
            gBusquedas.Filtro = " and cod_contabilidad = " & gCntX_Parametros.CodigoConta _
                              & " and tipo_asiento = '" & txtCAsiento & "'"
                              
            i = MsgBox("Desea buscar solo asientos del periodo actual?", vbYesNo)
            If i = vbYes Then
                 gBusquedas.Filtro = gBusquedas.Filtro & " and Anio = " & txtAnio & " and mes = " & txtMes
            End If
            
            gBusquedas.Consulta = "select Num_asiento,descripcion from Cntx_Asientos"
            frmBusquedas.Show vbModal
            txtNAsiento = gBusquedas.Resultado
            Call sbConsultaAsiento(txtCAsiento, txtNAsiento)
            
         Case 4 'Descripcion del numero de Asiento
            gBusquedas.Columna = "Descripcion"
            gBusquedas.Orden = "Descripcion"
            gBusquedas.Filtro = " and cod_contabilidad = " & gCntX_Parametros.CodigoConta _
                              & " and tipo_asiento = '" & txtCAsiento & "'"
            i = MsgBox("Desea buscar solo asientos del periodo actual?", vbYesNo)
            If i = vbYes Then
                 gBusquedas.Filtro = gBusquedas.Filtro & " and Anio = " & txtAnio & " and mes = " & txtMes
            End If
            
            gBusquedas.Consulta = "select Num_asiento,descripcion from Cntx_Asientos"
            frmBusquedas.Show vbModal
            txtNAsiento = gBusquedas.Resultado
            Call sbConsultaAsiento(txtCAsiento, txtNAsiento)
       End Select

    Case 5 'REPORTES
      
      strSQL = "{Cntx_Asientos.cod_contabilidad} = " & gCntX_Parametros.CodigoConta _
             & " AND {Cntx_Asientos.TIPO_ASIENTO} = '" & txtCAsiento & "' AND " _
             & " {Cntx_Asientos.NUM_ASIENTO} = '" & txtNAsiento & "'"
      
      Call sbCntX_Reportes("ASIENTO", strSQL)
    
End Select

End Sub

Private Sub btnProcesos_Click(Index As Integer)
Dim vCasos As Long, strSQL As String


Select Case Index
  Case 0  'Cuentas
    Call sbClassCall("Contabilidad", 0, "frmCntX_CatalogoCuentas")
  
  Case 1 'mayorizar
    If Val(txtDiferencia) = 0 And fxCntX_PeriodoVerifica(txtAnio, txtMes) Then
      If fxCntX_AsientoConcurrencia(vUltimos.TipoAsiento, vUltimos.NumAsiento) <> vUltimos.TS Then
        MsgBox "- El Asiento Actual A sido modificado por Otro Usuario/Proceso..."
      Else
        Call sbCntX_Asiento_Mayorizar(txtCAsiento, txtNAsiento, dtpAsientoFecha.Value)
        Call sbConsultaAsiento(txtCAsiento, txtNAsiento)
        MsgBox "Asiento Mayorizado...", vbInformation
      End If
    Else
        MsgBox "No se puede Mayorizar (Causas : 1. El Asiento no está Balanceado" _
                & " o el Periodo se encuentra Cerrado", vbCritical
    End If
    
  Case 2 'reversar
     If fxCntX_PeriodoVerifica(txtAnio, txtMes) Then
      If fxCntX_AsientoConcurrencia(vUltimos.TipoAsiento, vUltimos.NumAsiento) <> vUltimos.TS Then
        MsgBox "- El Asiento Actual A sido modificado por Otro Usuario/Proceso..."
      Else
        Call sbCntX_Asiento_Reversion(txtCAsiento, txtNAsiento, dtpAsientoFecha.Value)
        Call sbConsultaAsiento(txtCAsiento, txtNAsiento)
        MsgBox "Asiento Reversado...", vbInformation
      End If
     Else
        MsgBox "No se puede reversar este asiento porque el periodo se encuentra cerrado...", vbCritical
     End If
     
  Case 3 'Autorizar
     strSQL = "update Cntx_Asientos set user_autoriza = '" & glogon.Usuario _
          & "',fecha_autoriza = getdate()" _
          & " where cod_contabilidad = " & gCntX_Parametros.CodigoConta _
          & " and tipo_asiento = '" & txtCAsiento _
          & "' and num_asiento = '" & txtNAsiento _
          & "' and Fecha_Autoriza is null and Modulo <> 20"
   Call ConectionExecute(strSQL, 0, vCasos)
   If vCasos > 0 Then
        Call sbConsultaAsiento(txtCAsiento, txtNAsiento)
        MsgBox "Asiento Foráneo Autorizado...", vbInformation
   End If

  Case 4 'Historico
     GLOBALES.gTag = vUltimos.Cta
     Call sbFormsCall("frmCntX_CuentaHistorico", , , , False, Me)
     
End Select

End Sub

Private Sub chkCopiaAsReversion_Click()
txtCopia_NAsiento.Text = Replace(txtCopia_NAsiento.Text, "_Rev", "")

If chkCopiaAsReversion.Value = xtpChecked Then
txtCopia_NAsiento.Text = txtCopia_NAsiento.Text & "_Rev"
End If

End Sub

Private Sub chkCopiaDetalles_Click()
If chkCopiaDetalles.Value = vbChecked Then
    txtCopia_Documento.Enabled = True
Else
    txtCopia_Documento.Enabled = False
End If

txtCopia_Detalle.Enabled = txtCopia_Documento.Enabled

End Sub


Private Sub dtpAsientoFecha_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtDescripcion.SetFocus
End Sub

Private Sub sbLimpiezaParcial(iCodigo As Integer)
vGrid.MaxRows = 0
vGrid.MaxRows = 1

txtDescripcion = ""

Select Case iCodigo
  Case 1 'Cambia el Tipo de Asiento
    txtDAsiento = ""
    txtNAsiento = ""
  Case 2 'Cambia el periodo
    txtNAsiento = ""
   
End Select

End Sub

Private Sub btnCopiar_Click(Index As Integer)
Dim strSQL As String


On Error GoTo vError

Select Case Index
  Case 0 'Copiar
  
    strSQL = "exec spCntX_Asiento_Copia " & gCntX_Parametros.CodigoConta & ",'" & txtCAsiento.Text & "','" & txtNAsiento.Text _
           & "','" & txtCopia_NAsiento.Text & "','" & txtCopia_Descripcion.Text & "','" & Format(dtpCopia_Fecha.Value, "yyyy/mm/dd") _
           & "','" & glogon.Usuario & "','" & txtCopia_Notas.Text & "'," & chkCopiaDetalles.Value & ",'" & Mid(txtCopia_Documento.Text, 1, 35) _
           & "','" & Mid(txtCopia_Detalle.Text, 1, 100) & "','" & Mid(txtCopia_Referencia.Text, 1, 200) & "', " & chkCopiaAsReversion.Value
    Call ConectionExecute(strSQL)
    If glogon.error Then
        Exit Sub
    End If
    
    MsgBox "Copia de Asiento realizada satisfactoriamente!", vbInformation
    
    fraCopia.Visible = False
    Call sbConsultaAsiento(txtCAsiento.Text, txtCopia_NAsiento)
       
       
  Case 1 'Cerrar
    fraCopia.Visible = False

End Select

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub FlatScrollBar_Change()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

If vScroll Then
    strSQL = "select Top 1 tipo_asiento,num_asiento from Cntx_Asientos" _
           & " where cod_contabilidad = " & gCntX_Parametros.CodigoConta _
           & " and tipo_asiento = '" & txtCAsiento & "' and mes = " & txtMes _
           & " and anio = " & txtAnio
    
    If FlatScrollBar.Value = 1 Then
       strSQL = strSQL & " and num_asiento > '" & txtNAsiento & "' order by num_asiento asc"
    Else
       strSQL = strSQL & " and num_asiento < '" & txtNAsiento & "' order by num_asiento desc"
    End If
    
    Call OpenRecordSet(rs, strSQL, 0)
    If Not rs.EOF And Not rs.BOF Then
      Call sbConsultaAsiento(rs!Tipo_Asiento, rs!Num_Asiento)
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

Private Function fxNotaCuenta(vCuenta As String)
Dim strSQL As String, rs As New ADODB.Recordset
Dim vNota As String

On Error GoTo vError


strSQL = "select * from vCntX_Mov_Cuentas_General where cod_contabilidad = " & gCntX_Parametros.CodigoConta _
       & " and cod_cuenta = '" & vCuenta & "' and anio = " & txtAnio & " and mes = " & txtMes
Call OpenRecordSet(rs, strSQL, 0)
If rs.EOF And rs.BOF Then
  vNota = "No disponible por el momento"
Else
  vNota = "Estado del Periodo:" & vbCrLf _
        & "___________________" & vbCrLf _
        & " Saldo Inicial : " & Format(rs!saldo_inicial, "Standard") & vbCrLf _
        & " Total Debitos : " & Format(Abs(rs!total_debitos), "Standard") & vbCrLf _
        & " Total Creditos: " & Format(Abs(rs!total_creditos), "Standard") & vbCrLf _
        & " Mensual       : " & Format(rs!total_debitos + rs!total_creditos, "Standard") & vbCrLf _
        & " Acumulado     : " & Format((rs!saldo_inicial + rs!total_debitos + rs!total_creditos), "Standard") & vbCrLf _
        & "___________________"
End If
rs.Close

fxNotaCuenta = vNota

vError:

End Function



Public Sub sbFormReLoad()
 Call Form_Load
End Sub

Private Sub Form_Activate()
vModulo = 20
End Sub

Private Sub Form_Load()

vModulo = 20

Set Me.Icon = frmContenedor.Icon

vGrid.AppearanceStyle = fxGridStyle

 vScroll = False
 FlatScrollBar.Value = 0
 vScroll = True
 
' Call sbToolBarIconos(tlb)
 
 vUltimos.Anio = gCntX_Parametros.PeriodoAnio
 vUltimos.Mes = gCntX_Parametros.PeriodoMes
 

 
 If gCntX_Arbol.ArbolActivo Then
  Call sbConsultaAsiento(gCntX_Arbol.AsientoTipo, gCntX_Arbol.AsientoNumr)
 Else
    vEdita = False
    Call sbLimpiaPantalla
    Call sbBarra_Accion("edicion")
 End If
 
 
Call Formularios(Me)
Call RefrescaTags(Me)
 
End Sub

Private Function fxVerificaAsiento() As Boolean
Dim rsX As New ADODB.Recordset, strSQL As String
Dim vMensaje As String, lng As Long, pDivisa As String

'Verificar Periodo
'Tipo de Asiento
'Fecha del Asiento vrs Periodo
'Numero de Asiento
'Cuentas (En el Detalle)
'Concurrencia
'Autorizacion en Caso de Modificacion

'Nuevo
'Verificar el Tipo de Cambio
'Verificar Oficinas


fxVerificaAsiento = True
vMensaje = ""

strSQL = "select isnull(count(*),0) as existe from CntX_Periodos where cod_contabilidad = " & gCntX_Parametros.CodigoConta _
       & " and anio = " & txtAnio & " and mes = " & txtMes & " and estado = 'P'"
Call OpenRecordSet(rsX, strSQL, 0)
  If rsX!Existe = 0 Then vMensaje = vMensaje & vbCrLf & "- Periodo Indicado se encuentra Cerrado o No se ha creado..."
rsX.Close

strSQL = "select isnull(count(*),0) as existe from CntX_Tipos_Asientos where cod_contabilidad = " & gCntX_Parametros.CodigoConta _
       & " and tipo_asiento = '" & txtCAsiento & "'"
Call OpenRecordSet(rsX, strSQL, 0)
  If rsX!Existe = 0 Then vMensaje = vMensaje & vbCrLf & "- El tipo de Asiento Indicano no existe..."
rsX.Close


strSQL = "select COD_DIVISA  From CNTX_DIVISAS" _
       & " Where DIVISA_LOCAL = 1 And COD_CONTABILIDAD = " & gCntX_Parametros.CodigoConta
Call OpenRecordSet(rsX, strSQL, 0)
  pDivisa = UCase(Trim(rsX!cod_Divisa))
rsX.Close

'Verifica que el asiento no exista
If Not vEdita Then
    strSQL = "select isnull(count(*),0) as existe from CntX_Asientos where cod_contabilidad = " & gCntX_Parametros.CodigoConta _
           & " and Tipo_Asiento = '" & txtCAsiento.Text & "' and num_Asiento = '" & txtNAsiento.Text & "'"
    Call OpenRecordSet(rsX, strSQL, 0)
      If rsX!Existe > 0 Then vMensaje = vMensaje & vbCrLf & "- El Asiento a REGISTRAR ya existe! Consultelo para referencia o cambie el número actual..."
    rsX.Close
End If

If Month(dtpAsientoFecha) <> txtMes Then vMensaje = vMensaje & vbCrLf & "- El Mes del Periodo no se encuentra en la fecha del Asiento..."

If Year(dtpAsientoFecha) <> txtAnio Then vMensaje = vMensaje & vbCrLf & "- El Año del Periodo no se encuentra en la fecha del Asiento..."

If vEdita Then
  With stBar
    If .Panels(5) = "FORANEO" And Len(Trim(.Panels(4))) = 0 Then
      vMensaje = vMensaje & vbCrLf & "- EL ASIENTO ES FORANEO Y NO SE ENCUENTRA AUTORIZADO ..."
    End If
  End With
  If fxCntX_AsientoConcurrencia(vUltimos.TipoAsiento, vUltimos.NumAsiento) <> vUltimos.TS Then
      vMensaje = vMensaje & vbCrLf & "- El Asiento Actual A sido modificado por Otro Usuario/Proceso..."
  End If
End If

For lng = 1 To vGrid.MaxRows
 vGrid.Row = lng
 vGrid.col = 1
 If vGrid.Text <> "" Then
   vGrid.col = 2
   If vGrid.Text = "" Then
      vGrid.col = 1
      vMensaje = vMensaje & vbCrLf & "- Cuenta " & vGrid.Text & " No Existe"
   End If
   
   vGrid.col = 4
   If UCase(vGrid.Text) <> pDivisa Then
      vGrid.col = 5
      If CCur(vGrid.Text) = 1 Then
          vMensaje = vMensaje & vbCrLf & "- Línea " & lng & " : Tipo de Cambio Incorrecto."
      End If
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
txtAnio = vUltimos.Anio
txtMes = vUltimos.Mes
txtPeriodo = ""
txtCAsiento = vUltimos.TipoAsiento
txtCredito = 0
txtDebito = 0
txtDescripcion = ""
txtDiferencia = 0
txtDAsiento = ""
txtNAsiento = ""

txtNotas.Text = ""
txtReferencia.Text = ""

vGrid.MaxRows = 0
vGrid.MaxRows = 1
vGrid.MaxCols = 9


lblCuenta.Caption = ""
lblUnidad.Caption = ""
lblDivisa.Caption = ""

Call sbRefrescaInformacion


btnProcesos(1).Enabled = False
btnProcesos(2).Enabled = False

With stBar
  .Panels(1) = ""
  .Panels(2) = ""
  .Panels(3) = ""
  .Panels(4) = "INTERNO"
End With

End Sub

Private Sub imgBusquedaAdv_Click()
  Call sbFormsCall("frmCntX_AsientosConsultaAdv", 1, , , False, Me)
End Sub

Private Sub imgCopia_Click()
 fraCopia.Left = lblCuenta.Left
 fraCopia.top = lblCuenta.top
 
 fraCopia.Visible = True
 
 txtCopia_Detalle.Text = ""
 txtCopia_Documento.Text = ""
 
 txtCopia_NAsiento.Text = txtNAsiento.Text
 txtCopia_Descripcion.Text = txtDescripcion.Text
 txtCopia_Notas.Text = txtNotas.Text
 
 dtpCopia_Fecha.Value = dtpAsientoFecha.Value
 
 txtCopia_Referencia.Text = txtReferencia.Text
 
 txtCopia_NAsiento.SetFocus
 
End Sub





Private Sub sbCargaGridLocal(vGrid As Object, vGridMaxCol As Integer, strSQL As String)
Dim rs As New ADODB.Recordset, i As Integer
Dim vNota As String

Me.MousePointer = vbHourglass

vGrid.MaxCols = vGridMaxCol
vGrid.MaxRows = 1

vGrid.Row = vGrid.MaxRows

rs.CursorLocation = adUseServer
Call OpenRecordSet(rs, strSQL, 0)

Do While Not rs.EOF
  vNota = "Estado del Periodo:" & vbCrLf _
        & "___________________" & vbCrLf _
        & " Saldo Inicial : " & Format(rs!saldo_inicial, "Standard") & vbCrLf _
        & " Total Debitos : " & Format(Abs(rs!total_debitos), "Standard") & vbCrLf _
        & " Total Creditos: " & Format(Abs(rs!total_creditos), "Standard") & vbCrLf _
        & " Mensual       : " & Format(rs!total_debitos + rs!total_creditos, "Standard") & vbCrLf _
        & " Acumulado     : " & Format((rs!saldo_inicial + rs!total_debitos + rs!total_creditos), "Standard") & vbCrLf _
        & "___________________"
  vGrid.Row = vGrid.MaxRows
  For i = 1 To vGrid.MaxCols
    vGrid.col = i
    Select Case i
       Case 1 ' Cuenta
            vGrid.TextTip = TextTipFixed
            vGrid.TextTipDelay = 1000
            vGrid.CellNote = vNota
            vGrid.CellTag = rs!Descripcion
            vGrid.Text = fxCntX_CuentaFormato(True, CStr(rs!cod_cuenta))
           
       Case 2 ' Unidad
            vGrid.CellTag = rs!UniDes
            vGrid.Text = CStr(rs!Cod_Unidad)
       
       Case 3 ' Centro de Costo
            vGrid.TextTip = TextTipFixed
            vGrid.CellNote = rs!CentroCosto & ""
            
            vGrid.CellTag = rs!CentroCosto & ""
            vGrid.Text = CStr(rs!cod_centro_costo & "")
       
       Case 4 ' Divisa
            vGrid.Text = CStr(rs!cod_Divisa)
            vGrid.CellTag = rs!Divisa
       
       Case 5 'Tipo de Cambio
            vGrid.CellTag = CStr(IIf(IsNull(rs!tc_ajuste), rs!TC, rs!tc_ajuste))
            vGrid.Text = CStr(rs!TC)
       Case 6
            vGrid.Text = CStr(rs!Documento & "")
       Case 7
            vGrid.Text = CStr(rs!Detalle & "")
       Case 8
            vGrid.Text = CStr(rs!monto_debito)
       Case 9
            vGrid.Text = CStr(rs!monto_credito)
    End Select
  
  Next i
  vGrid.MaxRows = vGrid.MaxRows + 1
  rs.MoveNext
Loop

rs.Close

Me.MousePointer = vbDefault

End Sub


Private Sub sbConsultaAsiento(strTipo As String, strNumero As String)
Dim rs As New ADODB.Recordset, strSQL As String

On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = "select * from Cntx_Asientos where cod_contabilidad = " & gCntX_Parametros.CodigoConta _
       & " and tipo_asiento = '" & strTipo & "' and num_asiento = '" & strNumero & "'"

Call OpenRecordSet(rs, strSQL, 0)

If Not rs.BOF And Not rs.EOF Then
  Call sbBarra_Accion("activo")
  vEdita = True
 
  vUltimos.Anio = rs!Anio
  vUltimos.Mes = rs!Mes
  vUltimos.NumAsiento = rs!Num_Asiento
  vUltimos.TipoAsiento = rs!Tipo_Asiento
  vUltimos.fecha = rs!fecha_asiento
  vUltimos.TS = fxTsToHex(IIf(IsNull(rs!TS), "", rs!TS))
  vUltimos.AplAsiento = IIf(IsNull(rs!Fecha_Aplicado), 0, 1)
  
  'llenar datos en pantalla
  
  txtAnio = vUltimos.Anio
  txtMes = vUltimos.Mes
  Call sbRefrescaInformacion
  
  txtCAsiento.Text = vUltimos.TipoAsiento
  txtDAsiento.Text = fxCntX_TiposAsientos("D", vUltimos.TipoAsiento)
  txtDescripcion.Text = IIf(IsNull(rs!Descripcion), "", rs!Descripcion)
  txtNotas.Text = IIf(IsNull(rs!Notas), "", rs!Notas)
  txtReferencia.Text = IIf(IsNull(rs!Referencia), "", rs!Referencia)

  dtpAsientoFecha.Value = vUltimos.fecha
  txtNAsiento = vUltimos.NumAsiento
  
  If vUltimos.AplAsiento = 0 Then
    lblAsientoEstado.Caption = "Este Asiento se Encuentra Pendiente"
  Else
    lblAsientoEstado.Caption = "Este Asiento se Encuentra Mayorizado"
  End If
  
  With stBar
    .Panels(1) = UCase(Trim(IIf(IsNull(rs!user_crea), "", rs!user_crea)))
    .Panels(2) = UCase(Trim(IIf(IsNull(rs!user_modifica), "", rs!user_modifica)))
    .Panels(3) = UCase(Trim(IIf(IsNull(rs!user_aplica), "", rs!user_aplica)))
    .Panels(4) = UCase(Trim(IIf(IsNull(rs!user_autoriza), "", rs!user_autoriza)))
    If rs!Modulo = 20 Then
      .Panels(5) = "INTERNO"
    Else
      .Panels(5) = "FORANEO"
    End If
  End With
  
strSQL = "select A.cod_cuenta,B.descripcion,A.documento,A.detalle,A.monto_debito,A.monto_credito,A.num_linea" _
       & ",isnull(M.saldo_inicial,0) as Saldo_Inicial,isnull(M.total_debitos,0) as Total_debitos" _
       & ",isnull(M.total_creditos,0) as Total_creditos,A.cod_unidad,U.descripcion as UniDes" _
       & ",A.cod_divisa,Y.descripcion as Divisa,isnull(A.Tipo_Cambio,1) as TC,isnull(A.Tipo_Cambio_Ajuste,0) as TC_Ajuste" _
       & ",A.cod_Centro_Costo,Cc.Descripcion as CentroCosto" _
       & " from Cntx_Asientos_detalle A inner join Cntx_Asientos X on A.cod_contabilidad = X.cod_contabilidad" _
       & " and A.tipo_asiento = X.tipo_asiento and A.num_asiento = X.num_asiento" _
       & " inner join CntX_Cuentas B on A.cod_cuenta = B.cod_cuenta and A.cod_contabilidad = B.cod_contabilidad" _
       & " inner join CntX_Unidades U on A.cod_unidad = U.cod_unidad and A.cod_contabilidad = U.cod_contabilidad" _
       & " inner join CntX_Divisas Y on A.cod_divisa = Y.cod_divisa and A.cod_contabilidad = Y.cod_contabilidad" _
       & " left join vCntX_Mov_Cuentas_General M on A.cod_cuenta = M.cod_cuenta and A.cod_contabilidad = M.cod_contabilidad" _
       & " and M.anio = X.anio and M.mes = X.mes" _
       & " left join CntX_Centro_Costos Cc on A.cod_centro_costo = Cc.cod_centro_costo and A.cod_contabilidad = Cc.cod_contabilidad " _
       & " and A.cod_contabilidad = M.cod_contabilidad and X.anio = M.anio and X.mes = M.mes" _
       & " where A.cod_contabilidad = " & gCntX_Parametros.CodigoConta _
       & " and A.tipo_asiento = '" & vUltimos.TipoAsiento & "'" _
       & " and A.num_asiento = '" & vUltimos.NumAsiento & "'" _
       & " order by A.num_linea"
   
  Call sbCargaGridLocal(vGrid, 9, strSQL)
  
  If vGrid.MaxRows > 0 Then
    Call vGrid_LeaveCell(1, 1, 2, 1, True)
  End If
  
  Call sbSumaDebitosCreditos

  If vUltimos.AplAsiento = 1 Then
    btnProcesos(1).Enabled = False
    btnProcesos(2).Enabled = True
    vGrid.Lock = True
  Else
    vGrid.Lock = False
    btnProcesos(1).Enabled = True
    btnProcesos(2).Enabled = False
  End If

End If

rs.Close
Me.MousePointer = vbDefault

Exit Sub
vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub sbGuardar()
Dim strSQL As String, lng As Long
Dim vBitacoraMov As String, vBitacoraDetalle As String

On Error GoTo vError

If fxVerificaAsiento Then
    If vEdita Then
      
      strSQL = "update Cntx_Asientos set descripcion = '" & txtDescripcion.Text _
             & "',anio = " & txtAnio & ", mes = " & txtMes _
             & ",fecha_asiento = '" & Format(dtpAsientoFecha.Value, "yyyy/mm/dd") _
             & "',balanceado = '" & IIf(CCur(txtDiferencia) = 0, "S", "N") _
             & "',user_modifica = '" & UCase(Trim(glogon.Usuario)) _
             & "',notas = '" & Trim(txtNotas) & "', Referencia = '" & Mid(txtReferencia.Text, 1, 200) _
             & "' where cod_contabilidad = " & gCntX_Parametros.CodigoConta _
             & " and tipo_asiento = '" & UCase(txtCAsiento) _
             & "' and num_asiento = '" & txtNAsiento & "'"
'      Call ConectionExecute(strSQL, 0)
      
      vBitacoraMov = "Modifica"
      vBitacoraDetalle = "Asiento : " & txtCAsiento & "-" & txtNAsiento & " Conta." & gCntX_Parametros.CodigoConta
      
'      Call Bitacora("Modifica", "Asiento : " & txtCAsiento & "-" & txtNAsiento & " Conta." & gCntX_Parametros.CodigoConta)
    
    Else 'Inserta
       strSQL = "insert into Cntx_Asientos(tipo_asiento,cod_contabilidad,num_asiento,anio,mes" _
              & ",fecha_asiento,descripcion,balanceado,user_crea,modulo,notas, referencia) values('" & UCase(txtCAsiento) & "'," _
              & gCntX_Parametros.CodigoConta & ",'" & txtNAsiento & "'," & txtAnio _
              & "," & txtMes & ",'" & Format(dtpAsientoFecha.Value, "yyyy/mm/dd") & "','" & txtDescripcion.Text _
              & "','" & IIf(CCur(txtDiferencia) = 0, "S", "N") & "','" & Trim(UCase(glogon.Usuario)) _
              & "',20,'" & Trim(txtNotas) & "','" & Mid(txtReferencia.Text, 1, 200) & "')"
'       Call ConectionExecute(strSQL, 0)
       
      vBitacoraMov = "Registra"
      vBitacoraDetalle = "Asiento : " & txtCAsiento & "-" & txtNAsiento & " Conta." & gCntX_Parametros.CodigoConta
       
'        Call Bitacora("Registra", "Asiento : " & txtCAsiento & "-" & txtNAsiento & " Conta." & gCntX_Parametros.CodigoConta)
        
    End If 'Si Inserta o Actualiza


'Registra el detalle

      strSQL = strSQL & Space(10) & "delete Cntx_Asientos_detalle where cod_contabilidad = " _
             & gCntX_Parametros.CodigoConta & " and tipo_asiento = '" & UCase(txtCAsiento) _
             & "' and num_asiento = '" & txtNAsiento & "'"
 '     Call ConectionExecute(strSQL, 0)
    
      For lng = 1 To vGrid.MaxRows
        vGrid.Row = lng
        vGrid.col = 1
        If vGrid.Text <> "" Then
            strSQL = strSQL & Space(10) & "insert into Cntx_Asientos_detalle(num_linea,tipo_asiento,num_asiento,cod_contabilidad" _
                   & ",cod_cuenta,cod_unidad,cod_centro_costo,cod_divisa,tipo_cambio,documento,detalle,monto_debito,monto_credito)" _
                   & " values(" & lng & ",'" & UCase(txtCAsiento) & "','" & txtNAsiento & "'," _
                   & gCntX_Parametros.CodigoConta & ",'"
            vGrid.col = 1
            strSQL = strSQL & fxCntX_CuentaFormato(False, vGrid.Text) & "','"
            vGrid.col = 2
            strSQL = strSQL & vGrid.Text & "','"
            vGrid.col = 3
            strSQL = strSQL & vGrid.Text & "','"
            vGrid.col = 4
            strSQL = strSQL & vGrid.Text & "',"
            vGrid.col = 5
            strSQL = strSQL & CCur(IIf((vGrid.Text = ""), 0, vGrid.Text)) & ",'"
            vGrid.col = 6
            strSQL = strSQL & vGrid.Text & "','"
            vGrid.col = 7
            strSQL = strSQL & vGrid.Text & "',"
            vGrid.col = 8
            strSQL = strSQL & CCur(IIf((vGrid.Text = ""), 0, vGrid.Text)) & ","
            vGrid.col = 9
            strSQL = strSQL & CCur(IIf((vGrid.Text = ""), 0, vGrid.Text)) & ")"

'            Call ConectionExecute(strSQL, 0)
              
         End If 'vgrid.Text <> ""
       
       Next lng
    
        'Ejecuta el Lote Completo
        Call ConectionExecute(strSQL, 0)
        
        
        Call Bitacora(vBitacoraMov, vBitacoraDetalle)
        Call sbConsultaAsiento(txtCAsiento, txtNAsiento)
        Call sbBarra_Accion("activo")
        
        vEdita = True
        
        MsgBox "Información guardada satisfactoriamente...", vbInformation


End If 'Verificacion del Asiento

Call RefrescaTags(Me)

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
 
End Sub

Private Sub sbBorrar()
Dim i As Integer, strSQL As String

On Error GoTo vError

If fxCntX_AsientoConcurrencia(vUltimos.TipoAsiento, vUltimos.NumAsiento) <> vUltimos.TS Then
   MsgBox "- El Asiento actual a sido modificado por Otro Usuario/Proceso..."
   Exit Sub
End If

If stBar.Panels(5) = "FORANEO" And Len(Trim(stBar.Panels(4))) = 0 Then
  MsgBox "- EL ASIENTO ES FORANEO Y NO SE ENCUENTRA AUTORIZADO ...", vbExclamation
  Exit Sub
End If


i = MsgBox("Esta Seguro que desea borrar este registro", vbYesNo)

If i = vbYes Then
  strSQL = "delete Cntx_Asientos_detalle where cod_contabilidad = " _
         & gCntX_Parametros.CodigoConta & " and tipo_asiento = '" & UCase(txtCAsiento) _
         & "' and num_asiento = '" & txtNAsiento & "'"
  Call ConectionExecute(strSQL, 0)
  
  strSQL = "delete Cntx_Asientos where cod_contabilidad = " _
         & gCntX_Parametros.CodigoConta & " and tipo_asiento = '" & UCase(txtCAsiento) _
         & "' and num_asiento = '" & txtNAsiento & "'"
  Call ConectionExecute(strSQL, 0)
  

  Call Bitacora("Elimina", "Asiento : " & txtCAsiento & "-" & txtNAsiento & " Conta." _
                  & gCntX_Parametros.CodigoConta)

  Call sbLimpiaPantalla
  Call sbBarra_Accion("nuevo")
End If

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub tlbCopiar_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim strSQL As String


On Error GoTo vError

Select Case Button.Key
  Case "Copiar"
  
    strSQL = "exec spCntX_Asiento_Copia " & gCntX_Parametros.CodigoConta & ",'" & txtCAsiento.Text & "','" & txtNAsiento.Text _
           & "','" & txtCopia_NAsiento.Text & "','" & txtCopia_Descripcion.Text & "','" & Format(dtpCopia_Fecha.Value, "yyyy/mm/dd") _
           & "','" & glogon.Usuario & "','" & txtCopia_Notas.Text & "'," & chkCopiaDetalles.Value & ",'" & Mid(txtCopia_Documento.Text, 1, 35) _
           & "','" & Mid(txtCopia_Detalle.Text, 1, 100) & "','" & Mid(txtCopia_Referencia.Text, 1, 200) & "'"
    Call ConectionExecute(strSQL)
    If glogon.error Then
        Exit Sub
    End If
    
    MsgBox "Copia de Asiento realizada satisfactoriamente!", vbInformation
    
    fraCopia.Visible = False
    Call sbConsultaAsiento(txtCAsiento.Text, txtCopia_NAsiento)
       
       
  Case "Cerrar"
    fraCopia.Visible = False

End Select

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub txtAnio_Change()
 Call sbRefrescaInformacion
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
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtNAsiento.SetFocus
If KeyCode = vbKeyF4 Then Call btnBarra_Click(6)
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


Private Sub txtCopia_NAsiento_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF2 And txtCopia_NAsiento.Text = "" Then Call sbConsecutivoAsientoCopia
End Sub

Private Sub txtDAsiento_GotFocus()
vBusca = 2
End Sub

Private Sub txtDAsiento_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtNAsiento.SetFocus
If KeyCode = vbKeyF4 Then Call btnBarra_Click(6)
End Sub

Private Sub txtDescripcion_GotFocus()
vBusca = 4
End Sub

Private Sub txtDescripcion_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then vGrid.SetFocus
If KeyCode = vbKeyF4 Then Call btnBarra_Click(6)
End Sub

Private Sub txtMes_Change()
 Call sbRefrescaInformacion
End Sub

Private Sub txtNAsiento_GotFocus()
vBusca = 3
End Sub

Private Sub sbConsecutivoAsiento()
Dim strSQL As String, rs As New ADODB.Recordset

strSQL = "select isnull(CONSECUTIVO,0) + 1 as 'Consecutivo' " _
       & " From CNTX_TIPOS_ASIENTOS Where COD_CONTABILIDAD = " & gCntX_Parametros.CodigoConta _
       & " and tipo_asiento = '" & txtCAsiento.Text & "'"
       
Call OpenRecordSet(rs, strSQL, 0)
If Not rs.EOF And Not rs.BOF Then
  strSQL = "update CNTX_TIPOS_ASIENTOS set consecutivo = isnull(consecutivo,0) + 1" _
       & " Where COD_CONTABILIDAD = " & gCntX_Parametros.CodigoConta _
       & " and tipo_asiento = '" & txtCAsiento.Text & "'"
  Call ConectionExecute(strSQL, 0)
  
  txtNAsiento.Text = Format(rs!Consecutivo, "00000000")
End If
rs.Close


End Sub

Private Sub sbConsecutivoAsientoCopia()
Dim strSQL As String, rs As New ADODB.Recordset

strSQL = "select isnull(CONSECUTIVO,0) + 1 as 'Consecutivo' " _
       & " From CNTX_TIPOS_ASIENTOS Where COD_CONTABILIDAD = " & gCntX_Parametros.CodigoConta _
       & " and tipo_asiento = '" & txtCAsiento.Text & "'"
       
Call OpenRecordSet(rs, strSQL, 0)
If Not rs.EOF And Not rs.BOF Then
  strSQL = "update CNTX_TIPOS_ASIENTOS set consecutivo = isnull(consecutivo,0) + 1" _
       & " Where COD_CONTABILIDAD = " & gCntX_Parametros.CodigoConta _
       & " and tipo_asiento = '" & txtCAsiento.Text & "'"
  Call ConectionExecute(strSQL, 0)
  
  txtCopia_NAsiento.Text = Format(rs!Consecutivo, "00000000")
End If
rs.Close


End Sub

Private Sub txtNAsiento_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then
 dtpAsientoFecha.SetFocus
 Call sbConsultaAsiento(txtCAsiento, txtNAsiento)
End If
If KeyCode = vbKeyF4 Then Call btnBarra_Click(6)
If KeyCode = vbKeyF2 And txtNAsiento.Text = "" Then Call sbConsecutivoAsiento
End Sub


Private Function fxVerificaCuenta(strCuenta As String) As Boolean
Dim rsX As New ADODB.Recordset, strSQL As String

strSQL = "select isnull(count(*),0) as Existe from CntX_Cuentas where cod_contabilidad = " & gCntX_Parametros.CodigoConta _
       & " and cod_cuenta = '" & strCuenta & "' and acepta_movimientos =1"

Call OpenRecordSet(rsX, strSQL, 0)
 fxVerificaCuenta = IIf((rsX!Existe = 0), False, True)
rsX.Close

End Function


Private Sub sbSumaDebitosCreditos()
Dim x As Long, TC As Currency
  
On Error GoTo vError
  
  txtDebito = 0
  txtCredito = 0
  For x = 1 To vGrid.MaxRows
      vGrid.Row = x
'      vGrid.col = 4
'      TC = CCur(vGrid.Text)
'
'      If TC = 0 Then TC = 1

    'Solo en Colones en Pantalla por Eso no Hay que Convertir Tipo de Cambio
     TC = 1
      
      vGrid.col = 8
      txtDebito = CCur(txtDebito) + (CCur(IIf(vGrid.Text = "", 0, vGrid.Text)) * fxSys_Tipo_Cambio_Apl(TC))
      vGrid.col = 9
      txtCredito = CCur(txtCredito) + (CCur(IIf(vGrid.Text = "", 0, vGrid.Text)) * fxSys_Tipo_Cambio_Apl(TC))
  Next x
  txtDiferencia = txtDebito - txtCredito
  txtDebito = Format(txtDebito, "Standard")
  txtCredito = Format(txtCredito, "Standard")
  txtDiferencia = Format(txtDiferencia, "Standard")

vError:

End Sub


Private Sub vGrid_KeyDown(KeyCode As Integer, Shift As Integer)
Dim i As Variant, lng As Long, vTemp(9) As Variant, x As Integer
Dim vTempo As String

If KeyCode = vbKeyDelete Then
  
  vGrid.Row = vGrid.ActiveRow
  vGrid.col = vGrid.MaxCols
  If vGrid.Text <> "" Then 'Existe en la Base de datos
    'Preguntar y si la respuesta es afirmativa eliminar de la Base de datos
  
  End If
  
  For lng = vGrid.ActiveRow To vGrid.MaxRows
     vGrid.Row = lng + 1
     For x = 1 To vGrid.MaxCols
        vGrid.col = x
        vTemp(x) = vGrid.Text
     Next x
     
     vGrid.Row = lng
     For x = 1 To vGrid.MaxCols
       vGrid.col = x
       vGrid.Text = vTemp(x)
     Next x
  Next lng
  vGrid.MaxRows = vGrid.MaxRows - 1
  If vGrid.MaxRows = 0 Then vGrid.MaxRows = 1
  
  Call sbSumaDebitosCreditos
  
  
End If

'Consulta cuenta
If KeyCode = vbKeyF4 And vGrid.ActiveCol = 1 Then
  frmCntX_ConsultaCuentas.Show vbModal
  vGrid.col = vGrid.ActiveCol
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
    
  vGrid.col = vGrid.ActiveCol
  vGrid.Row = vGrid.ActiveRow
  
  vGrid.Text = gBusquedas.Resultado
  vGrid.CellTag = gBusquedas.Resultado2
  
End If


'Consulta Centro de Costo
If KeyCode = vbKeyF4 And vGrid.ActiveCol = 3 Then
  vGrid.Row = vGrid.ActiveRow
  vGrid.col = 2
  vTempo = vGrid.Text
  
  gBusquedas.Resultado = ""
  gBusquedas.Resultado2 = ""
  gBusquedas.Columna = "descripcion"
  gBusquedas.Orden = "descripcion"
  gBusquedas.Filtro = " and cod_contabilidad = " & gCntX_Parametros.CodigoConta & " and cod_centro_costo in(select cod_centro_costo" _
                    & " from cntX_unidades_cc where cod_unidad = '" & vTempo & "' and cod_contabilidad = " & gCntX_Parametros.CodigoConta & ")"
  gBusquedas.Consulta = "select cod_centro_costo,descripcion from CntX_Centro_Costos"
  frmBusquedas.Show vbModal
    
  vGrid.col = vGrid.ActiveCol
  vGrid.Row = vGrid.ActiveRow
  
  vGrid.Text = gBusquedas.Resultado
  vGrid.CellTag = gBusquedas.Resultado2
  vGrid.TextTip = TextTipFixed
  vGrid.CellNote = vGrid.CellTag
  
End If


'''Consulta Divisa
''If KeyCode = vbKeyF4 And vGrid.ActiveCol = 4 Then
''  gBusquedas.Resultado = ""
''  gBusquedas.Resultado2 = ""
''  gBusquedas.Columna = "descripcion"
''  gBusquedas.Orden = "descripcion"
''  gBusquedas.Filtro = " and cod_contabilidad = " & gCntX_Parametros.CodigoConta
''  gBusquedas.Consulta = "select cod_divisa,descripcion from CntX_Divisas"
''  frmBusquedas.Show vbModal
''
''  vGrid.col = vGrid.ActiveCol
''  vGrid.Row = vGrid.ActiveRow
''
''  vGrid.Text = gBusquedas.Resultado
''  vGrid.CellTag = gBusquedas.Resultado2
''End If


If (KeyCode = vbKeyReturn Or KeyCode = vbKeyTab) Then
    vGrid.col = vGrid.ActiveCol
    vGrid.Row = vGrid.ActiveRow
    
    Select Case vGrid.ActiveCol
      Case 1 'Cuenta
        vGrid.Text = fxCntX_CuentaFormato(True, vGrid.Text)
        i = fxCntX_CuentaFormato(False, vGrid.Text)
        If fxVerificaCuenta(CStr(i)) Then
          vGrid.TextTip = TextTipFixed
          
          
          vGrid.CellNote = fxNotaCuenta(CStr(i))
          
          lblCuenta.Caption = fxCntX_Cuenta("D", CStr(i))
          
          vGrid.CellTag = lblCuenta.Caption
          
          vUltimos.Cta = CStr(i)
          vUltimos.CCDesc = lblCuenta.Caption
          
          
          vTempo = fxCntX_CuentaDivisa(CStr(i))
          
          If vTempo <> "" Then
            vGrid.col = 4
            vGrid.Text = vTempo
          End If

        Else
          MsgBox "Cuenta no es válida : " & vbCrLf & " - No Existe o No Acepta Movimientos" _
                 & vbCrLf & " - VERIFIQUE O MODIFIQUE EN EL CATALAGO DE CUENTAS", vbCritical
        End If
        
      Case 2
        'Buscar la Unidad
        If fxCntx_UnidadVerifica(vGrid.Text) Then
          vGrid.CellTag = fxCntX_Unidad("D", vGrid.Text)
          vUltimos.Unidad = vGrid.Text
          vUltimos.UnidadDesc = vGrid.CellTag
        Else
          MsgBox "La UNIDAD de negocio no es válida : " & vbCrLf & " - No Existe...", vbCritical
        End If
      
      
      Case 3 'Verificar el Centro de Costo
        vGrid.col = 2
        vTempo = vGrid.Text
        vGrid.col = 3
        
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
          MsgBox "La DIVISA no es válida : " & vbCrLf & " - No Existe...", vbCritical
        End If
      
      
      Case 6
        vUltimos.Documento = vGrid.Text
      
      Case 7
        vUltimos.Detalle = vGrid.Text
        
      Case 8 'Debe
        If Val(vGrid.Text) > 0 Then
            vGrid.col = vGrid.ActiveCol + 1
            vGrid.Row = vGrid.ActiveRow
            vGrid.Text = 0
        
            Call sbSumaDebitosCreditos
            
        End If
      
      Case 9 'Haber
        If Val(vGrid.Text) > 0 Then
            vGrid.col = vGrid.ActiveCol - 1
            vGrid.Row = vGrid.ActiveRow
            vGrid.Text = 0
            Call sbSumaDebitosCreditos
        End If
        
        If vGrid.MaxRows = vGrid.ActiveRow Then
            vGrid.MaxRows = vGrid.MaxRows + 1
            vGrid.Row = vGrid.MaxRows
            vGrid.col = 2
            vGrid.Text = vUltimos.Unidad
            vGrid.CellTag = vUltimos.UnidadDesc
            
            vGrid.col = 3
            vGrid.Text = vUltimos.CC
            vGrid.TextTip = TextTipFixed
            vGrid.CellTag = vUltimos.CCDesc
            vGrid.CellNote = vGrid.CellTag
            
            vGrid.col = 4
            vGrid.Text = vUltimos.Divisa
            vGrid.CellTag = vUltimos.DivisaDesc
            vGrid.col = 6
            vGrid.Text = vUltimos.Documento
            vGrid.col = 7
            vGrid.Text = vUltimos.Detalle
        End If
        
    
    End Select


End If

If KeyCode = vbKeyInsert Then
    vGrid.MaxRows = vGrid.MaxRows + 1
    vGrid.InsertRows vGrid.ActiveRow, 1
    vGrid.Row = vGrid.ActiveRow
    vGrid.col = 2
    vGrid.Text = vUltimos.Unidad
    vGrid.CellTag = vUltimos.UnidadDesc
    
    vGrid.col = 3
    vGrid.Text = vUltimos.CC
    vGrid.TextTip = TextTipFixed
    vGrid.CellTag = vUltimos.CCDesc
    vGrid.CellNote = vGrid.CellTag
    
    vGrid.col = 4
    vGrid.Text = vUltimos.Divisa
    vGrid.CellTag = vUltimos.DivisaDesc
    vGrid.col = 6
    vGrid.Text = vUltimos.Documento
    vGrid.col = 7
    vGrid.Text = vUltimos.Detalle
End If

'Activa Ventana para tipo de cambio
'If KeyCode >= 94 And KeyCode <= 105 Then
If (vGrid.ActiveCol = 8 Or vGrid.ActiveCol = 9) _
    And lblCuenta.Tag = "N" And KeyCode = vbKeyF4 Then
       vGrid.col = 1
       gCntX_TipoCambio.Cuenta = fxCntX_CuentaFormato(False, vGrid.Text, 0)
       gCntX_TipoCambio.fecha = dtpAsientoFecha.Value
       vGrid.col = vGrid.ActiveCol
       If vGrid.Text = "" Then
           gCntX_TipoCambio.Monto_Actual = 0
       Else
           gCntX_TipoCambio.Monto_Actual = CCur(vGrid.Text)
       End If
       vGrid.col = 4
       gCntX_TipoCambio.Moneda = vGrid.Text
       vGrid.col = 5
       If vGrid.Text = "" Or vGrid.Text = "0.0000" Then
         gCntX_TipoCambio.TC_Actual = 1
       Else
         gCntX_TipoCambio.TC_Actual = vGrid.Text
       End If
       
       frmCntX_TipoCambio.Show vbModal
       If gCntX_TipoCambio.Paso Then
          vGrid.Row = vGrid.ActiveRow
          vGrid.col = 5
          vGrid.Text = CStr(gCntX_TipoCambio.TC_Nuevo)
          vGrid.col = vGrid.ActiveCol
          vGrid.Text = CStr(gCntX_TipoCambio.Monto_Nuevo)
       End If
    End If
'End If


End Sub

Private Sub vGrid_LeaveCell(ByVal col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
Dim vCuenta As String, vDivisa As String, vMov As Currency, i As Byte

On Error GoTo vError

vGrid.Row = Row


If NewCol = 8 Or NewCol = 9 Then
    vGrid.col = 4
    lblCuenta.Tag = IIf(fxCntX_DivisaBase(vGrid.Text), "S", "N")
End If

If col = 4 Then
    'Verificar Tipo de Cambio
    vGrid.col = 4
    vDivisa = vGrid.Text
    
    vGrid.col = 5
    If vGrid.Text = "" Or vGrid.Text = "0.0000" Then
      vGrid.col = 1
      vCuenta = fxCntX_CuentaFormato(False, vGrid.Text, 0)
      vGrid.col = 5
      vGrid.Text = fxCntX_TipoCambio(vDivisa, vCuenta, dtpAsientoFecha)
    Else
'        vGrid.col = 4
'        i = MsgBox("Desea Cambiar el Tipo de Cambio Actual [" _
'            & vGrid.Text & "] para Esta Linea?", vbYesNo)
'        If i = vbYes Then
'            vGrid.col = 1
'            vCuenta = fxCntX_CuentaFormato(False, vGrid.Text, 0)
'            vGrid.col = 4
'            vGrid.Text = fxCntX_TipoCambio(vDivisa, vCuenta, dtpAsientoFecha)
'        End If
    End If
End If

vGrid.Row = NewRow
vGrid.col = 1
lblCuenta.Caption = vGrid.CellTag

vUltimos.Cta = vGrid.Text
vUltimos.CtaDesc = vGrid.CellTag

vGrid.col = 2
lblUnidad.Caption = vGrid.CellTag
vGrid.col = 4
lblDivisa.Caption = vGrid.CellTag

vGrid.col = 8
vMov = CCur(vGrid.Text)
vGrid.col = 9
vMov = vMov + CCur(vGrid.Text)
vGrid.col = 5
lblDivisa.Caption = lblDivisa.Caption & " [" & Format(vMov / fxSys_Tipo_Cambio_Apl(CCur(vGrid.Text)), "Standard") & "]"

vError:

End Sub

