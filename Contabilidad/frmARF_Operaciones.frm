VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "ComCt332.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.Controls.v24.0.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.ShortcutBar.v24.0.0.ocx"
Begin VB.Form frmARF_Operaciones 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Operaciones de Arrendamientos"
   ClientHeight    =   8490
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   11745
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8490
   ScaleWidth      =   11745
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.TabControl tcMain 
      Height          =   7335
      Left            =   120
      TabIndex        =   8
      Top             =   1080
      Width           =   11535
      _Version        =   1572864
      _ExtentX        =   20346
      _ExtentY        =   12938
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
      ItemCount       =   4
      Item(0).Caption =   "Registro"
      Item(0).ControlCount=   28
      Item(0).Control(0)=   "txtUnidadCod"
      Item(0).Control(1)=   "txtUnidadDesc"
      Item(0).Control(2)=   "txtArrendadorCod"
      Item(0).Control(3)=   "txtArrendadorDesc"
      Item(0).Control(4)=   "cboFrecuencia"
      Item(0).Control(5)=   "Label2(0)"
      Item(0).Control(6)=   "Label2(1)"
      Item(0).Control(7)=   "Label2(2)"
      Item(0).Control(8)=   "txtMonto"
      Item(0).Control(9)=   "Label2(3)"
      Item(0).Control(10)=   "Label2(4)"
      Item(0).Control(11)=   "txtTasaDescuento"
      Item(0).Control(12)=   "Label2(5)"
      Item(0).Control(13)=   "txtTasaInteres"
      Item(0).Control(14)=   "Label2(6)"
      Item(0).Control(15)=   "Label2(7)"
      Item(0).Control(16)=   "Label2(8)"
      Item(0).Control(17)=   "txtPlazo"
      Item(0).Control(18)=   "txtNotas"
      Item(0).Control(19)=   "Label2(9)"
      Item(0).Control(20)=   "dtpInicio"
      Item(0).Control(21)=   "dtpCorte"
      Item(0).Control(22)=   "gbRegistro"
      Item(0).Control(23)=   "Label2(10)"
      Item(0).Control(24)=   "txtIncremento"
      Item(0).Control(25)=   "txtEstado"
      Item(0).Control(26)=   "Label2(13)"
      Item(0).Control(27)=   "txtGarantiaMonto"
      Item(1).Caption =   "Plan"
      Item(1).ControlCount=   1
      Item(1).Control(0)=   "vGrid"
      Item(2).Caption =   "Cierres"
      Item(2).ControlCount=   3
      Item(2).Control(0)=   "lswCierres"
      Item(2).Control(1)=   "scHistorial"
      Item(2).Control(2)=   "btnExport(2)"
      Item(3).Caption =   "Asientos"
      Item(3).ControlCount=   7
      Item(3).Control(0)=   "lswAsiento"
      Item(3).Control(1)=   "lswAsientoMain"
      Item(3).Control(2)=   "GroupBox2"
      Item(3).Control(3)=   "btnExport(5)"
      Item(3).Control(4)=   "btnExport(6)"
      Item(3).Control(5)=   "ShortcutCaption4(2)"
      Item(3).Control(6)=   "scAsientos"
      Begin XtremeSuiteControls.ListView lswAsientoMain 
         Height          =   3135
         Left            =   -69880
         TabIndex        =   46
         Top             =   840
         Visible         =   0   'False
         Width           =   11295
         _Version        =   1572864
         _ExtentX        =   19923
         _ExtentY        =   5530
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
      Begin XtremeSuiteControls.ListView lswAsiento 
         Height          =   1935
         Left            =   -69880
         TabIndex        =   45
         Top             =   4440
         Visible         =   0   'False
         Width           =   11295
         _Version        =   1572864
         _ExtentX        =   19923
         _ExtentY        =   3413
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
      Begin XtremeSuiteControls.ListView lswCierres 
         Height          =   6255
         Left            =   -69880
         TabIndex        =   42
         Top             =   840
         Visible         =   0   'False
         Width           =   11295
         _Version        =   1572864
         _ExtentX        =   19923
         _ExtentY        =   11033
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
      Begin XtremeSuiteControls.GroupBox gbRegistro 
         Height          =   2175
         Left            =   2040
         TabIndex        =   31
         Top             =   4920
         Width           =   7815
         _Version        =   1572864
         _ExtentX        =   13785
         _ExtentY        =   3836
         _StockProps     =   79
         UseVisualStyle  =   -1  'True
         Appearance      =   21
         BorderStyle     =   2
         Begin XtremeSuiteControls.FlatEdit txtR_Fecha 
            Height          =   315
            Left            =   1560
            TabIndex        =   35
            Top             =   720
            Width           =   2295
            _Version        =   1572864
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
         Begin XtremeSuiteControls.FlatEdit txtR_Usuario 
            Height          =   315
            Left            =   3840
            TabIndex        =   36
            Top             =   720
            Width           =   2295
            _Version        =   1572864
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
         Begin XtremeSuiteControls.FlatEdit txtA_Fecha 
            Height          =   315
            Left            =   1560
            TabIndex        =   37
            Top             =   1200
            Width           =   2295
            _Version        =   1572864
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
         Begin XtremeSuiteControls.FlatEdit txtA_Usuario 
            Height          =   315
            Left            =   3840
            TabIndex        =   38
            Top             =   1200
            Width           =   2295
            _Version        =   1572864
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
         Begin XtremeSuiteControls.Label Label3 
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   34
            Top             =   1200
            Width           =   855
            _Version        =   1572864
            _ExtentX        =   1508
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Activa"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin XtremeSuiteControls.Label Label3 
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   33
            Top             =   720
            Width           =   855
            _Version        =   1572864
            _ExtentX        =   1508
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Registra"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
            Height          =   255
            Left            =   0
            TabIndex        =   32
            Top             =   0
            Width           =   7815
            _Version        =   1572864
            _ExtentX        =   13785
            _ExtentY        =   450
            _StockProps     =   14
            Caption         =   "Registro de Control"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9.01
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Alignment       =   1
         End
      End
      Begin XtremeSuiteControls.FlatEdit txtUnidadCod 
         Height          =   315
         Left            =   2040
         TabIndex        =   9
         ToolTipText     =   "Presione F4 para consultar"
         Top             =   840
         Width           =   1815
         _Version        =   1572864
         _ExtentX        =   3196
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
      Begin XtremeSuiteControls.FlatEdit txtUnidadDesc 
         Height          =   315
         Left            =   3840
         TabIndex        =   10
         Top             =   840
         Width           =   6015
         _Version        =   1572864
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
      Begin XtremeSuiteControls.FlatEdit txtArrendadorCod 
         Height          =   315
         Left            =   2040
         TabIndex        =   11
         ToolTipText     =   "Presione F4 para consultar"
         Top             =   1200
         Width           =   1815
         _Version        =   1572864
         _ExtentX        =   3196
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
      Begin XtremeSuiteControls.FlatEdit txtArrendadorDesc 
         Height          =   315
         Left            =   3840
         TabIndex        =   12
         Top             =   1200
         Width           =   6015
         _Version        =   1572864
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
      Begin XtremeSuiteControls.ComboBox cboFrecuencia 
         Height          =   330
         Left            =   2040
         TabIndex        =   13
         Top             =   2040
         Width           =   1815
         _Version        =   1572864
         _ExtentX        =   3201
         _ExtentY        =   582
         _StockProps     =   77
         ForeColor       =   0
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
         Style           =   2
         Appearance      =   6
         UseVisualStyle  =   0   'False
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.FlatEdit txtMonto 
         Height          =   315
         Left            =   2040
         TabIndex        =   17
         Top             =   1680
         Width           =   1815
         _Version        =   1572864
         _ExtentX        =   3196
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
         Alignment       =   1
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtTasaDescuento 
         Height          =   315
         Left            =   6240
         TabIndex        =   20
         Top             =   1680
         Width           =   975
         _Version        =   1572864
         _ExtentX        =   1720
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
      Begin XtremeSuiteControls.FlatEdit txtTasaInteres 
         Height          =   315
         Left            =   6240
         TabIndex        =   22
         Top             =   2040
         Width           =   975
         _Version        =   1572864
         _ExtentX        =   1720
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
      Begin XtremeSuiteControls.FlatEdit txtNotas 
         Height          =   1155
         Left            =   2040
         TabIndex        =   27
         Top             =   3600
         Width           =   7815
         _Version        =   1572864
         _ExtentX        =   13785
         _ExtentY        =   2037
         _StockProps     =   77
         ForeColor       =   0
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
         BackColor       =   16777215
         MultiLine       =   -1  'True
         ScrollBars      =   2
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.DateTimePicker dtpInicio 
         Height          =   330
         Left            =   2040
         TabIndex        =   29
         Top             =   2400
         Width           =   1815
         _Version        =   1572864
         _ExtentX        =   3201
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
      Begin XtremeSuiteControls.DateTimePicker dtpCorte 
         Height          =   330
         Left            =   2040
         TabIndex        =   30
         Top             =   2760
         Width           =   1815
         _Version        =   1572864
         _ExtentX        =   3201
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
      Begin XtremeSuiteControls.FlatEdit txtPlazo 
         Height          =   315
         Left            =   8880
         TabIndex        =   26
         Top             =   1680
         Width           =   975
         _Version        =   1572864
         _ExtentX        =   1720
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
      Begin XtremeSuiteControls.FlatEdit txtIncremento 
         Height          =   315
         Left            =   6240
         TabIndex        =   40
         Top             =   2400
         Width           =   975
         _Version        =   1572864
         _ExtentX        =   1720
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
      Begin FPSpreadADO.fpSpread vGrid 
         Height          =   6975
         Left            =   -70000
         TabIndex        =   41
         Top             =   360
         Visible         =   0   'False
         Width           =   11535
         _Version        =   524288
         _ExtentX        =   20346
         _ExtentY        =   12303
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
         ScrollBarShowMax=   0   'False
         SpreadDesigner  =   "frmARF_Operaciones.frx":0000
         VScrollSpecial  =   -1  'True
         VScrollSpecialType=   2
         ScrollBarTrack  =   1
         AppearanceStyle =   1
         ScrollBarStyle  =   2
      End
      Begin XtremeSuiteControls.PushButton btnExport 
         Height          =   255
         Index           =   2
         Left            =   -58840
         TabIndex        =   44
         ToolTipText     =   "Exportar a Excel"
         Top             =   480
         Visible         =   0   'False
         Width           =   255
         _Version        =   1572864
         _ExtentX        =   444
         _ExtentY        =   444
         _StockProps     =   79
         Appearance      =   7
         Picture         =   "frmARF_Operaciones.frx":0BA6
      End
      Begin XtremeSuiteControls.GroupBox GroupBox2 
         Height          =   735
         Left            =   -69880
         TabIndex        =   47
         Top             =   6480
         Visible         =   0   'False
         Width           =   11295
         _Version        =   1572864
         _ExtentX        =   19923
         _ExtentY        =   1296
         _StockProps     =   79
         UseVisualStyle  =   -1  'True
         BorderStyle     =   2
         Begin XtremeSuiteControls.FlatEdit txtA_Debito 
            Height          =   312
            Left            =   2760
            TabIndex        =   48
            Top             =   120
            Width           =   1932
            _Version        =   1572864
            _ExtentX        =   3408
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
            Alignment       =   1
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtA_Credito 
            Height          =   312
            Left            =   4800
            TabIndex        =   49
            Top             =   120
            Width           =   1932
            _Version        =   1572864
            _ExtentX        =   3408
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
            Alignment       =   1
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtA_Diferencia 
            Height          =   312
            Left            =   8160
            TabIndex        =   50
            Top             =   120
            Width           =   1932
            _Version        =   1572864
            _ExtentX        =   3408
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
            Alignment       =   1
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.Label Label2 
            Height          =   252
            Index           =   12
            Left            =   960
            TabIndex        =   52
            Top             =   120
            Width           =   1572
            _Version        =   1572864
            _ExtentX        =   2773
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Débito / Crédito"
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
         End
         Begin XtremeSuiteControls.Label Label2 
            Height          =   252
            Index           =   11
            Left            =   6480
            TabIndex        =   51
            Top             =   120
            Width           =   1572
            _Version        =   1572864
            _ExtentX        =   2773
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Diferencia"
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
         End
      End
      Begin XtremeSuiteControls.PushButton btnExport 
         Height          =   255
         Index           =   5
         Left            =   -58960
         TabIndex        =   53
         ToolTipText     =   "Exportar a Excel"
         Top             =   4080
         Visible         =   0   'False
         Width           =   255
         _Version        =   1572864
         _ExtentX        =   444
         _ExtentY        =   444
         _StockProps     =   79
         Appearance      =   7
         Picture         =   "frmARF_Operaciones.frx":1477
      End
      Begin XtremeSuiteControls.PushButton btnExport 
         Height          =   255
         Index           =   6
         Left            =   -58960
         TabIndex        =   54
         ToolTipText     =   "Exportar a Excel"
         Top             =   480
         Visible         =   0   'False
         Width           =   255
         _Version        =   1572864
         _ExtentX        =   444
         _ExtentY        =   444
         _StockProps     =   79
         Appearance      =   7
         Picture         =   "frmARF_Operaciones.frx":1D48
      End
      Begin XtremeSuiteControls.FlatEdit txtEstado 
         Height          =   315
         Left            =   8040
         TabIndex        =   57
         Top             =   480
         Width           =   1815
         _Version        =   1572864
         _ExtentX        =   3196
         _ExtentY        =   550
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
      Begin XtremeSuiteControls.FlatEdit txtGarantiaMonto 
         Height          =   315
         Left            =   2040
         TabIndex        =   59
         Top             =   3240
         Width           =   1815
         _Version        =   1572864
         _ExtentX        =   3196
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
         Alignment       =   1
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.Label Label2 
         Height          =   255
         Index           =   13
         Left            =   360
         TabIndex        =   58
         Top             =   3240
         Width           =   1575
         _Version        =   1572864
         _ExtentX        =   2778
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Deposito Garantía"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Transparent     =   -1  'True
         WordWrap        =   -1  'True
      End
      Begin XtremeShortcutBar.ShortcutCaption scAsientos 
         Height          =   375
         Left            =   -69880
         TabIndex        =   56
         Top             =   4080
         Visible         =   0   'False
         Width           =   11295
         _Version        =   1572864
         _ExtentX        =   19923
         _ExtentY        =   661
         _StockProps     =   14
         Caption         =   "Detalle del Asiento"
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
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption4 
         Height          =   375
         Index           =   2
         Left            =   -69880
         TabIndex        =   55
         Top             =   480
         Visible         =   0   'False
         Width           =   11295
         _Version        =   1572864
         _ExtentX        =   19923
         _ExtentY        =   661
         _StockProps     =   14
         Caption         =   "Asientos Relacionados"
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
      Begin XtremeShortcutBar.ShortcutCaption scHistorial 
         Height          =   375
         Left            =   -69880
         TabIndex        =   43
         Top             =   480
         Visible         =   0   'False
         Width           =   11295
         _Version        =   1572864
         _ExtentX        =   19923
         _ExtentY        =   661
         _StockProps     =   14
         Caption         =   "Cierres"
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
      Begin XtremeSuiteControls.Label Label2 
         Height          =   255
         Index           =   10
         Left            =   4080
         TabIndex        =   39
         Top             =   2400
         Width           =   2055
         _Version        =   1572864
         _ExtentX        =   3625
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "[ % ] Incremento anual"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Transparent     =   -1  'True
         WordWrap        =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label2 
         Height          =   255
         Index           =   9
         Left            =   360
         TabIndex        =   28
         Top             =   3600
         Width           =   1575
         _Version        =   1572864
         _ExtentX        =   2778
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Notas"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Transparent     =   -1  'True
         WordWrap        =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label2 
         Height          =   255
         Index           =   8
         Left            =   7320
         TabIndex        =   25
         Top             =   1680
         Width           =   1695
         _Version        =   1572864
         _ExtentX        =   2990
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Plazo en Meses"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Transparent     =   -1  'True
         WordWrap        =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label2 
         Height          =   255
         Index           =   7
         Left            =   360
         TabIndex        =   24
         Top             =   2760
         Width           =   1575
         _Version        =   1572864
         _ExtentX        =   2778
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Termina"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Transparent     =   -1  'True
         WordWrap        =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label2 
         Height          =   255
         Index           =   6
         Left            =   360
         TabIndex        =   23
         Top             =   2400
         Width           =   1575
         _Version        =   1572864
         _ExtentX        =   2778
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Inicia"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Transparent     =   -1  'True
         WordWrap        =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label2 
         Height          =   255
         Index           =   5
         Left            =   4080
         TabIndex        =   21
         Top             =   2040
         Width           =   1575
         _Version        =   1572864
         _ExtentX        =   2778
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Tasa Interés anual"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Transparent     =   -1  'True
         WordWrap        =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label2 
         Height          =   255
         Index           =   4
         Left            =   4080
         TabIndex        =   19
         Top             =   1680
         Width           =   2055
         _Version        =   1572864
         _ExtentX        =   3625
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Tasa Descuento anual"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Transparent     =   -1  'True
         WordWrap        =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label2 
         Height          =   255
         Index           =   3
         Left            =   360
         TabIndex        =   18
         Top             =   2040
         Width           =   1575
         _Version        =   1572864
         _ExtentX        =   2778
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Frecuencia"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Transparent     =   -1  'True
         WordWrap        =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label2 
         Height          =   255
         Index           =   2
         Left            =   360
         TabIndex        =   16
         Top             =   1680
         Width           =   1575
         _Version        =   1572864
         _ExtentX        =   2778
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Monto"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Transparent     =   -1  'True
         WordWrap        =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label2 
         Height          =   255
         Index           =   1
         Left            =   360
         TabIndex        =   15
         Top             =   1200
         Width           =   1575
         _Version        =   1572864
         _ExtentX        =   2778
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Arrendador"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Transparent     =   -1  'True
         WordWrap        =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label2 
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   14
         Top             =   840
         Width           =   1575
         _Version        =   1572864
         _ExtentX        =   2778
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Unidad"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Transparent     =   -1  'True
         WordWrap        =   -1  'True
      End
   End
   Begin VB.Timer TimerX 
      Interval        =   10
      Left            =   8880
      Top             =   600
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   7800
      Top             =   480
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
            Picture         =   "frmARF_Operaciones.frx":2619
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmARF_Operaciones.frx":271A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmARF_Operaciones.frx":2839
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmARF_Operaciones.frx":2959
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmARF_Operaciones.frx":2A6E
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmARF_Operaciones.frx":2B8C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmARF_Operaciones.frx":2CB6
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmARF_Operaciones.frx":2DDC
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmARF_Operaciones.frx":2EF8
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmARF_Operaciones.frx":2FF4
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmARF_Operaciones.frx":310B
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin ComCtl3.CoolBar CoolBarX 
      Align           =   1  'Align Top
      Height          =   390
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11745
      _ExtentX        =   20717
      _ExtentY        =   688
      BandCount       =   2
      BandBorders     =   0   'False
      _CBWidth        =   11745
      _CBHeight       =   390
      _Version        =   "6.7.9839"
      Child1          =   "tlbPrincipal"
      MinHeight1      =   330
      Width1          =   2955
      NewRow1         =   0   'False
      Child2          =   "tlbAux"
      MinHeight2      =   330
      Width2          =   2520
      NewRow2         =   0   'False
      Begin MSComctlLib.Toolbar tlbPrincipal 
         Height          =   330
         Left            =   165
         TabIndex        =   2
         Top             =   30
         Width           =   2760
         _ExtentX        =   4868
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         Style           =   1
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   8
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "nuevo"
               Object.ToolTipText     =   "Inserta (Agrega) un registro nuevo a la Base de Datos"
               Object.Tag             =   "1"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "editar"
               Object.ToolTipText     =   "Modifica (Edita) el registro en pantalla"
               Object.Tag             =   "1"
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "guardar"
               Object.ToolTipText     =   "Guarda la información del registro en la base de datos"
               Object.Tag             =   "1"
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "deshacer"
               Object.ToolTipText     =   "Deshace toda modificación realizada recientemente en el registro actual"
               Object.Tag             =   "1"
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "reportes"
               Object.Tag             =   "1"
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "ayuda"
               Object.ToolTipText     =   "Ayuda General"
               Object.Tag             =   "1"
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "cerrar"
               Object.ToolTipText     =   "Cierra esta ventana"
               Object.Tag             =   "1"
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.Toolbar tlbAux 
         Height          =   330
         Left            =   3120
         TabIndex        =   1
         Top             =   30
         Width           =   8535
         _ExtentX        =   15055
         _ExtentY        =   582
         ButtonWidth     =   1693
         ButtonHeight    =   582
         Style           =   1
         TextAlignment   =   1
         ImageList       =   "ImageList1"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   3
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Activar"
               Key             =   "Activar"
               ImageIndex      =   5
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Anular"
               Key             =   "Anular"
               ImageIndex      =   6
            EndProperty
         EndProperty
      End
   End
   Begin MSComCtl2.FlatScrollBar FlatScrollBar 
      Height          =   252
      Left            =   3480
      TabIndex        =   3
      Top             =   480
      Width           =   492
      _ExtentX        =   873
      _ExtentY        =   450
      _Version        =   393216
      Arrows          =   65536
      Orientation     =   1638401
   End
   Begin XtremeSuiteControls.FlatEdit txtOperacion 
      Height          =   432
      Left            =   1440
      TabIndex        =   4
      Top             =   480
      Width           =   1932
      _Version        =   1572864
      _ExtentX        =   3408
      _ExtentY        =   762
      _StockProps     =   77
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   13.5
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
   Begin XtremeSuiteControls.PushButton btnAdjuntos 
      Height          =   330
      Left            =   11160
      TabIndex        =   5
      ToolTipText     =   "Adjuntar Documentos"
      Top             =   720
      Width           =   495
      _Version        =   1572864
      _ExtentX        =   873
      _ExtentY        =   582
      _StockProps     =   79
      BackColor       =   16777215
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Picture         =   "frmARF_Operaciones.frx":3234
   End
   Begin VB.Label lblNombre 
      BackStyle       =   0  'Transparent
      Caption         =   "..."
      Height          =   252
      Left            =   4800
      TabIndex        =   7
      Top             =   480
      Width           =   4932
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
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
      Height          =   252
      Index           =   0
      Left            =   0
      TabIndex        =   6
      Top             =   480
      Width           =   1212
   End
   Begin VB.Image ImgAutorizacion 
      Height          =   252
      Left            =   4080
      Top             =   480
      Width           =   252
   End
   Begin VB.Image imgId_Cambio 
      Height          =   252
      Left            =   4440
      Picture         =   "frmARF_Operaciones.frx":32BD
      Stretch         =   -1  'True
      ToolTipText     =   "Ajustes y Correcciones"
      Top             =   480
      Width           =   252
   End
End
Attribute VB_Name = "frmARF_Operaciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim strSQL As String, rs As New ADODB.Recordset

Dim vMensaje        As String  'Envia Mensajes en Fallas de Verificacion
Dim vEdita          As Boolean 'Indica si se esta actualizando o insertando
Dim vPaso           As Boolean 'Control de Activacion de Controles en proceso de carga
Dim vScroll         As Boolean
Dim vFecha          As Date, vOperacion As Long

Private Sub btnAdjuntos_Click()
 gGA.Modulo = "ARF"
 gGA.Llave_01 = txtOperacion.Text
 gGA.Llave_02 = ""
 gGA.Llave_03 = ""
 
 Call sbFormsCall("frmGA_Documentos", vbModal, , , False, Me, True)
End Sub

Private Sub sbPlan_Load()

On Error GoTo vError

Me.MousePointer = vbHourglass


strSQL = "exec spARF_Operacion_Plan " & txtOperacion.Text
Call sbCargaGrid(vGrid, vGrid.MaxCols, strSQL)

Me.MousePointer = vbDefault
Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub sbActivar()


On Error GoTo vError

Me.MousePointer = vbHourglass

'Activa la Operacion
strSQL = "exec spARF_Operacion_Activacion " & txtOperacion.Text & ", '" & glogon.Usuario & "'"
Call ConectionExecute(strSQL)

'BITACORA
Call Bitacora("Registra", "Activación de Operación de Arrendamiento No.: " & txtOperacion)

'Imprime Boleta de Activacion
'Call sbReporte_Boleta

Me.MousePointer = vbDefault

MsgBox "Activación aplicada Satisfactoriamente...", vbInformation

Call sbConsulta

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
  
  
End Sub

Private Sub sbAnular()
Dim strSQL As String

On Error GoTo vError

'Me.MousePointer = vbHourglass
'
'
'strSQL = "exec spCxC_Cuenta_Anulacion " & txtOperacion.Text & ",'" & glogon.Usuario & "',''"
'Call ConectionExecute(strSQL)
'
'
''BITACORA
'Call Bitacora("Registra", "Anulación de la Operacion CxC No." & txtOperacion.Text)
'
'Me.MousePointer = vbDefault
'
'MsgBox "Anulación de la Operacion de CxC realizada satisfactoriamente!", vbInformation
'
''Imprime Nota de Reversion
'Call sbImprimeRecibo(Trim(txtOperacion.Text) & "-A", "CxC_FRM")
'
''Imprime Boleta de Activacion/Anulacion
'Call sbReporte_Boleta


Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Function fxVerificaRecepcion() As Boolean
Dim rsX As New ADODB.Recordset, strSQL As String, rs As New ADODB.Recordset
Dim Porcentaje As Currency, pOperacion As Long

fxVerificaRecepcion = True
vMensaje = ""

pOperacion = IIf(IsNumeric(txtOperacion.Text), txtOperacion.Text, 0)

If txtArrendadorCod.Text = "" Then
       vMensaje = vMensaje & vbCrLf & " - No se ha especificado un Arrendador!"
End If

If txtUnidadCod.Text = "" Then
       vMensaje = vMensaje & vbCrLf & " - No se ha especificado una Unidad/Local!"
End If

If Not IsNumeric(txtMonto.Text) Then
       vMensaje = vMensaje & vbCrLf & " - El Monto no es válido!"
End If

If Not IsNumeric(txtTasaDescuento.Text) Then
       vMensaje = vMensaje & vbCrLf & " - La Tasa Descuento no es válida!"
End If

If Not IsNumeric(txtTasaInteres.Text) Then
       vMensaje = vMensaje & vbCrLf & " - Las Tasa de Interés no es válida!"
End If

If Not IsNumeric(txtPlazo.Text) Then
       vMensaje = vMensaje & vbCrLf & " - El Plazo no es válido!"
End If


If Not IsNumeric(txtIncremento.Text) Then
       vMensaje = vMensaje & vbCrLf & " - El Porcentaje de Incremento Anual, no es válido!"
End If
If Not IsNumeric(txtGarantiaMonto.Text) Then
       vMensaje = vMensaje & vbCrLf & " - El dato del depósito de garantía no es válido!"
End If

If dtpInicio.Value >= dtpCorte.Value Then
       vMensaje = vMensaje & vbCrLf & " - Rango de Fechas Erróneo, verificar!..."
End If

 
 
If Len(vMensaje) = 0 Then
    
    If CCur(txtMonto.Text) <= 0 Then
           vMensaje = vMensaje & vbCrLf & " - El Monto no es válido!"
    End If
    
    If CCur(txtGarantiaMonto.Text) < 0 Then
           vMensaje = vMensaje & vbCrLf & " - El dato del depósito de garantía no es válido!"
    End If
    
    If CCur(txtTasaDescuento.Text) > 100 Or CCur(txtTasaDescuento.Text) < 0 Then
           vMensaje = vMensaje & vbCrLf & " - La Tasa Descuento no es válida!"
    End If
    
    If CCur(txtTasaInteres.Text) > 100 Or CCur(txtTasaInteres.Text) < 0 Then
           vMensaje = vMensaje & vbCrLf & " - Las Tasa de Interés no es válida!"
    End If

    If CCur(txtIncremento.Text) > 100 Or CCur(txtIncremento.Text) < 0 Then
           vMensaje = vMensaje & vbCrLf & " - El Porcentaje de Incremento Anual, no es válido!"
    End If

End If


If Len(vMensaje) > 0 Then
    fxVerificaRecepcion = False
    MsgBox vMensaje, vbExclamation
Else
    fxVerificaRecepcion = True
End If

End Function

Private Function fxActivacionVerifica() As Boolean
Dim vMensaje As String

vMensaje = ""

fxActivacionVerifica = True

vFecha = fxFechaServidor

If txtArrendadorCod.Text = "" Then
       vMensaje = vMensaje & vbCrLf & " - No se ha especificado un Arrendador!"
End If

If txtUnidadCod.Text = "" Then
       vMensaje = vMensaje & vbCrLf & " - No se ha especificado una Unidad/Local!"
End If

If Not IsNumeric(txtMonto.Text) Then
       vMensaje = vMensaje & vbCrLf & " - El Monto no es válido!"
End If

If Not IsNumeric(txtTasaDescuento.Text) Then
       vMensaje = vMensaje & vbCrLf & " - La Tasa Descuento no es válida!"
End If

If Not IsNumeric(txtTasaInteres.Text) Then
       vMensaje = vMensaje & vbCrLf & " - Las Tasa de Interés no es válida!"
End If

If Not IsNumeric(txtPlazo.Text) Then
       vMensaje = vMensaje & vbCrLf & " - El Plazo no es válido!"
End If


If Not IsNumeric(txtIncremento.Text) Then
       vMensaje = vMensaje & vbCrLf & " - El Porcentaje de Incremento Anual, no es válido!"
End If
If Not IsNumeric(txtGarantiaMonto.Text) Then
       vMensaje = vMensaje & vbCrLf & " - El dato del depósito de garantía no es válido!"
End If

If dtpInicio.Value >= dtpCorte.Value Then
       vMensaje = vMensaje & vbCrLf & " - Rango de Fechas Erróneo, verificar!..."
End If

 
 
If Len(vMensaje) = 0 Then
    
    If CCur(txtMonto.Text) <= 0 Then
           vMensaje = vMensaje & vbCrLf & " - El Monto no es válido!"
    End If
    
    If CCur(txtGarantiaMonto.Text) < 0 Then
           vMensaje = vMensaje & vbCrLf & " - El dato del depósito de garantía no es válido!"
    End If
    
    If CCur(txtTasaDescuento.Text) > 100 Or CCur(txtTasaDescuento.Text) < 0 Then
           vMensaje = vMensaje & vbCrLf & " - La Tasa Descuento no es válida!"
    End If
    
    If CCur(txtTasaInteres.Text) > 100 Or CCur(txtTasaInteres.Text) < 0 Then
           vMensaje = vMensaje & vbCrLf & " - Las Tasa de Interés no es válida!"
    End If

    If CCur(txtIncremento.Text) > 100 Or CCur(txtIncremento.Text) < 0 Then
           vMensaje = vMensaje & vbCrLf & " - El Porcentaje de Incremento Anual, no es válido!"
    End If

End If

If Len(vMensaje) > 0 Then
 MsgBox vMensaje, vbCritical
 fxActivacionVerifica = False

Else
 fxActivacionVerifica = True
End If

End Function

Private Function fxAnulacionVerifica() As Boolean
Dim rs As New ADODB.Recordset, strSQL As String
Dim lngPriDeduc As Long, vFecha As Date
Dim Porcentaje As Double, vMontoRefunde As Currency
Dim rsTmp As New ADODB.Recordset

vMensaje = ""
fxAnulacionVerifica = True


''0. Verificacion base / Solo se pueden anular las formalizaciones del día
'strSQL = "select fechaforp,datediff(d,fechaforp,dbo.MyGetdate()) as Dias from ARF_OPERACIONES where Operacion = " & Operacion.Operacion
'Call OpenRecordSet(rs, strSQL)
'    vFecha = rs!fechaforp
'    If Abs(rs!Dias) > 0 Then
'      vMensaje = vMensaje & vbCrLf & "- Esta operación fue formalizada un día diferente al actual..."
'    End If
'rs.Close
'
'
''2. Verifica que no se le registren desembolsos, Se deben de anular o eliminar
'strSQL = "select isnull(count(*),0) as Existe from Tes_Transacciones where op = " & Operacion.Operacion _
'       & " and estado <> 'A'"
'Call OpenRecordSet(rs, strSQL)
'If rs!Existe > 0 Then
'  vMensaje = vMensaje & vbCrLf & "- Existen solicitudes o documentos emitidos (Cheques/Transferencias) en Tesorería (Proceda a Anularlos)"
'End If
'rs.Close
'
'If GLOBALES.SysPlanPagos = 0 Then
'    '3. Verificar si se le han realizado movimientos a la Operacion despues de su formalizacion
'    strSQL = "select isnull(count(*),0) as Existe from creditos_dt where Operacion = " & Operacion.Operacion _
'           & " and ncon <> Operacion"
'    Call OpenRecordSet(rs, strSQL)
'    If rs!Existe > 0 Then
'      vMensaje = vMensaje & vbCrLf & "- Existen movimientos a esta operación después de su formalización"
'    End If
'    rs.Close
'
'    'No tiene porque tener ningun registro de mora
'    strSQL = "select isnull(count(*),0) as Existe from MOROSIDAD where Operacion = " & Operacion.Operacion
'    Call OpenRecordSet(rs, strSQL)
'    If rs!Existe > 0 Then
'      vMensaje = vMensaje & vbCrLf & "- Existen movimientos a esta operación después de su formalización"
'    End If
'    rs.Close
'
'
'    '3a. Verificar si se le han realizado movimientos a las refundiciones (Abonadas o Canceladas)
'    strSQL = "select Operacion,consec from creditos_dt where tcon = 3 and ncon = " & Operacion.Operacion
'    Call OpenRecordSet(rs, strSQL)
'    Do While Not rs.EOF
'     strSQL = "select isnull(count(*),0) as Existe from creditos_dt where Operacion = " _
'            & rs!Operacion & " and consec > " & rs!consec
'     Call OpenRecordSet(rsTmp, strSQL, 0)
'        If rsTmp!Existe > 0 Then
'          vMensaje = vMensaje & vbCrLf & "- Existen movimientos realizados a la op:" & rs!Operacion _
'                   & " posterior a su refundicion"
'        End If
'     rsTmp.Close
'     rs.MoveNext
'    Loop
'    rs.Close
'
'    '3a. a la fecha de formalizacion (Doble verificacion para movimientos en mora no reflejados)
'    strSQL = "select isnull(count(*),0) as Existe from creditos_dt" _
'            & " where fechas > '" & Format(vFecha, "yyyy/mm/dd") & "' and Operacion in(select Operacion" _
'            & " from creditos_dt where tcon = 3 and Operacion = " & Operacion.Operacion & ")"
'    Call OpenRecordSet(rs, strSQL)
'    If rs!Existe > 0 Then
'       vMensaje = vMensaje & vbCrLf & "- Existen movimientos realizados a refundiciones posterior a la formalizacion"
'    End If
'    rs.Close
'
'    '3b. a la fecha de formalizacion para Morosidad
'    strSQL = "select isnull(count(*),0) as Existe from morosidad" _
'            & " where Estado = 'C' and fecUlt > '" & Format(vFecha, "yyyy/mm/dd") & "' and Operacion in(select Operacion" _
'            & " from morosidad where estado = 'C' and tcon = 3 and Operacion = " & Operacion.Operacion & ")"
'    Call OpenRecordSet(rs, strSQL)
'    If rs!Existe > 0 Then
'       vMensaje = vMensaje & vbCrLf & "- Existen movimientos realizados a Mora de refundiciones posterior a la formalizacion"
'    End If
'    rs.Close
'
'End If 'SysPlanPagos = 0
'
''4. No puede anular retenciones
'strSQL = "select retencion from catalogo where codigo = '" & Operacion.Codigo & "'"
'Call OpenRecordSet(rs, strSQL)
'If rs!retencion = "S" Then
'   vMensaje = vMensaje & vbCrLf & "- Este es un código de retención No se puede Anular..."
'End If
'rs.Close


If Len(vMensaje) > 0 Then fxAnulacionVerifica = False


End Function





Private Sub FlatScrollBar_Change()

On Error GoTo vError

If txtOperacion.Text = "" Or Not IsNumeric(txtOperacion.Text) Then txtOperacion.Text = "0"

If FlatScrollBar.Tag = "" Then FlatScrollBar.Tag = 0

strSQL = "select Top 1 Operacion from ARF_OPERACIONES"

If FlatScrollBar.Value > CLng(FlatScrollBar.Tag) Then
   strSQL = strSQL & " where Operacion > " & txtOperacion & " order by Operacion asc"
Else
   strSQL = strSQL & " where Operacion < " & txtOperacion & " order by Operacion desc"
End If

FlatScrollBar.Tag = FlatScrollBar.Value

Call OpenRecordSet(rs, strSQL)
If Not rs.EOF And Not rs.BOF Then
  txtOperacion.Text = rs!Operacion
  Call sbConsulta
End If

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub



Private Sub Form_Load()
 vModulo = 20
 
 vFecha = fxFechaServidor
 
 Call sbToolBarIconos(tlbPrincipal, False)
 
 tcMain.Item(0).Selected = True
 
 cboFrecuencia.AddItem "Mensual"
 cboFrecuencia.Text = "Mensual"
 
 Call sbLimpiaDatos

 With tlbPrincipal.Buttons
   .Item(2).Enabled = False
   .Item(3).Enabled = False
   .Item(4).Enabled = False
 End With

Call Formularios(Me)
Call RefrescaTags(Me)

End Sub


Private Sub sbLimpiaDatos()

 
 
'lblAutorizaEstado.Caption = "En Trámite"
'Set ImgAutorizacion.Picture = ImageList1.ListImages.Item(11).Picture
'ImgAutorizacion.ToolTipText = "En Proceso: Consulta/Nuevo"
 
 txtEstado.Text = "Pendiente"
 
 tlbAux.Buttons.Item(1).Enabled = False
 tlbAux.Buttons.Item(3).Enabled = False

 txtArrendadorCod.Text = ""
 txtArrendadorDesc.Text = ""
 lblNombre.Caption = txtArrendadorDesc.Text
 
 txtUnidadCod.Text = ""
 txtUnidadDesc.Text = ""
   
 txtMonto.Text = "0"
 txtPlazo.Text = "1"
 txtTasaDescuento.Text = "0"
 txtTasaInteres.Text = "0"
 txtIncremento.Text = "0"
 
 txtGarantiaMonto.Text = "0"
  
 dtpInicio.Value = vFecha
 dtpCorte.Value = vFecha
  
 txtNotas.Text = ""
 
 txtR_Fecha.Text = ""
 txtR_Usuario.Text = ""
 
 txtA_Fecha.Text = ""
 txtA_Usuario.Text = ""
 
 tcMain.Item(0).Selected = True
 tcMain.Item(1).Enabled = False
 tcMain.Item(2).Enabled = False
 

End Sub


Private Sub sbConsulta()
Dim vFecha As Date, iMes As Integer, lngAnio As Long
Dim i As Integer, vTemp As String

On Error GoTo vError

vPaso = True

strSQL = "select * from vARF_Operacion_Consulta" _
       & " where Operacion = " & txtOperacion.Text
Call OpenRecordSet(rs, strSQL)

If Not rs.EOF And Not rs.BOF Then
 tcMain.Item(1).Enabled = True
 tcMain.Item(2).Enabled = True
 
 
 txtUnidadCod.Text = rs!cod_Local
 txtUnidadDesc.Text = rs!Unidad_Desc

 txtArrendadorCod.Text = rs!cod_Acreedor & ""
 txtArrendadorDesc.Text = rs!ARRENDATARIO_DESC & ""
 
 txtMonto.Text = Format(IIf(IsNull(rs!Cuota), 0, rs!Cuota), "Standard")
 txtPlazo.Text = IIf(IsNull(rs!Plazo), 1, rs!Plazo)
 
 txtTasaDescuento.Text = Format(IIf(IsNull(rs!Tasa_Descuento), 0, rs!Tasa_Descuento), "Standard")
 txtTasaInteres.Text = Format(CStr(IIf(IsNull(rs!Tasa_Interes), 0, rs!Tasa_Interes)), "Standard")
 
 txtIncremento.Text = Format(IIf(IsNull(rs!Incremento_Anual_Porc), 0, rs!Incremento_Anual_Porc), "Standard")
 
 txtNotas.Text = rs!Notas & ""
 
 cboFrecuencia.Text = rs!Periodicidad_Desc
 
 txtGarantiaMonto.Text = Format(rs!Deposito_Garantia_Monto, "Standard")
 
 dtpInicio.Value = rs!Fecha_Inicio
 dtpCorte.Value = rs!Fecha_Finaliza
 
 tlbAux.Buttons.Item(1).Enabled = False
 tlbAux.Buttons.Item(3).Enabled = False

 txtEstado.Text = rs!Estado_Desc
 Select Case rs!Estado
   Case "R"
      
          tlbAux.Buttons.Item(1).Enabled = True
          tlbAux.Buttons.Item(3).Enabled = True
      
      imgId_Cambio.Visible = False
      
   Case "A"
      tlbAux.Buttons.Item(3).Enabled = True
      
      imgId_Cambio.Visible = True
   
   Case "C"
      imgId_Cambio.Visible = False
   
   Case "N"
      
      imgId_Cambio.Visible = False
  
  End Select


txtR_Usuario.Text = IIf(IsNull(rs!REGISTRO_USUARIO), "", rs!REGISTRO_USUARIO)
txtR_Fecha.Text = rs!REGISTRO_FECHA & ""
 
txtA_Usuario.Text = IIf(IsNull(rs!ACTIVA_USUARIO), "", rs!ACTIVA_USUARIO)
txtA_Fecha.Text = IIf(IsNull(rs!Activa_Fecha), rs!Activa_Fecha & "", rs!Activa_Fecha)
 

 With tlbPrincipal.Buttons
   .Item(1).Enabled = True
   .Item(2).Enabled = True
   .Item(3).Enabled = False
   .Item(4).Enabled = False
 End With
 
 
Else
 MsgBox "No existe este número de Operación, verifique!", vbCritical
End If
rs.Close

vPaso = False


Call RefrescaTags(Me)

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

  Resume
End Sub


Private Sub sbGuardar()
Dim pOperacion As Long

On Error GoTo vError
      

Select Case Mid(txtEstado.Text, 1, 1)
  Case "R", "P" 'Recepción
    If Not vEdita Then
       strSQL = "INSERT INTO ARF_OPERACIONES(COD_ACREEDOR, COD_LOCAL, TASA_DESCUENTO, TASA_INTERES, PERIODICIDAD, CUOTA, PLAZO, FECHA_INICIO, FECHA_FINALIZA" _
            & " , CORTE_ULTIMO, PAGO_PROXIMO, NOTAS, ESTADO, DEPOSITO_GARANTIA_MONTO ,DEPOSITO_GARANTIA_IND, INCREMENTO_ANUAL_PORC" _
            & " ,VALOR_PASIVO, DEPRECIACION_ACUM, VALOR_LIBROS, VALOR_INICIAL, REGISTRO_FECHA, REGISTRO_USUARIO)" _
            & " VALUES(" & txtArrendadorCod.Text & ",'" & txtUnidadCod.Text & "', " & CCur(txtTasaDescuento.Text) & ", " & CCur(txtTasaInteres.Text) _
            & ", '" & Mid(cboFrecuencia.Text, 1, 1) & "', " & CCur(txtMonto.Text) & ", " & CLng(txtPlazo.Text) & ", '" & Format(dtpInicio.Value, "yyyy-mm-dd") _
            & "', '" & Format(dtpCorte.Value, "yyyy-mm-dd") & "', Null, Null, '" & txtNotas.Text & "', 'R', " & CCur(txtGarantiaMonto.Text) & ", 0, " & CCur(txtIncremento.Text) _
            & ", 0, 0, 0, 0, getdate(), '" & glogon.Usuario & "')"
       Call ConectionExecute(strSQL)
              
        strSQL = "select isnull(max(Operacion),0) as 'Operacion' from ARF_OPERACIONES"
        Call OpenRecordSet(rs, strSQL)
          pOperacion = rs!Operacion
        rs.Close
              
        txtOperacion.Text = pOperacion
    
    Else
       strSQL = "update ARF_OPERACIONES set COD_ACREEDOR = " & txtArrendadorCod.Text & ", COD_LOCAL = '" & txtUnidadCod.Text _
            & "', TASA_DESCUENTO = " & CCur(txtTasaDescuento.Text) & ", TASA_INTERES = " & CCur(txtTasaInteres.Text) _
            & ", PERIODICIDAD = '" & Mid(cboFrecuencia.Text, 1, 1) & "', CUOTA = " & CCur(txtMonto.Text) & ", PLAZO = " & CLng(txtPlazo.Text) _
            & ", FECHA_INICIO = '" & Format(dtpInicio.Value, "yyyy-mm-dd") & "', FECHA_FINALIZA = '" & Format(dtpCorte.Value, "yyyy-mm-dd") _
            & "', CORTE_ULTIMO = Null, PAGO_PROXIMO = Null, NOTAS = '" & txtNotas.Text & "', ESTADO = 'R'" _
            & ", DEPOSITO_GARANTIA_MONTO  = " & CCur(txtGarantiaMonto.Text) & ", DEPOSITO_GARANTIA_IND = 1, INCREMENTO_ANUAL_PORC = " & CCur(txtIncremento.Text) _
            & " ,VALOR_PASIVO = 0, DEPRECIACION_ACUM = 0, VALOR_LIBROS = 0, VALOR_INICIAL = 0" _
           & " where Operacion = " & txtOperacion.Text
        Call ConectionExecute(strSQL)
      
    End If
  
        
        strSQL = "exec spARF_Operacion_Plan_Add " & txtOperacion.Text
        Call ConectionExecute(strSQL)
        
        
  Case "A" 'Activa
'      If gbFactoreo.Visible Then
'        strSQL = "update ARF_OPERACIONES set notas = '" & txtNotas.Text & "', emitir_tipo = '" & fxTipoDocumento(cboTipoDocumento.Text) & "', Emitir_Banco = " _
'               & cboBanco.ItemData(cboBanco.ListIndex) & ",Emitir_Cuenta = '" & pCuenta _
'               & "', Desembolso_Pendiente = ADELANTO_MONTO, cedula_pagador = " & pPagadorCed _
'               & " where Operacion = " & txtOperacion.Text & " and Tesoreria_Fecha is null"
'      Else
'        strSQL = "update ARF_OPERACIONES set notas = '" & txtNotas.Text & "', emitir_tipo = '" & fxTipoDocumento(cboTipoDocumento.Text) & "', Emitir_Banco = " _
'               & cboBanco.ItemData(cboBanco.ListIndex) & ",Emitir_Cuenta = '" & pCuenta _
'               & "', Desembolso_Pendiente = Desembolso_Monto, cedula_pagador = " & pPagadorCed _
'               & " where Operacion = " & txtOperacion.Text & " and Tesoreria_Fecha is null"
'      End If
'      Call ConectionExecute(strSQL)
'

      
      
End Select

Call sbConsulta

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub



Private Sub tcMain_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)

If Item.Index > 0 Then
  If Not IsNumeric(txtOperacion.Text) Then
        MsgBox "Consulte una Operación!", vbInformation
        tcMain.Item(0).Selected = True
        Exit Sub
  End If
End If

Dim i As Integer, pMonto As Currency
Dim pDebito As Currency, pCredito As Currency


Select Case Item.Index
    Case 0 'General
    Case 1 'Plan
        Call sbPlan_Load
    Case 2 'Cierres
        Call sbARF_Cierres_Load(lswCierres, txtOperacion.Text)
    
    Case 2 'Asiento
        Call sbAsientos_Load
    
        Call sbARF_Asiento_Load(lswAsiento, txtOperacion.Text)

        pDebito = 0
        pCredito = 0

        'Totales
        With lswAsiento.ListItems
        
        For i = 1 To .Count
            pDebito = pDebito + CCur(.Item(i).SubItems(3))
            pCredito = pCredito + CCur(.Item(i).SubItems(4))
        Next i
        
        End With
        
        txtA_Debito.Text = Format(pDebito, "Standard")
        txtA_Credito.Text = Format(pCredito, "Standard")
        txtA_Diferencia.Text = Format(pDebito - pCredito, "Standard")

End Select

End Sub


Private Sub sbAsientos_Load()
Dim pInicio As String, pCorte As String, pFiltro As String, pDetalle As String

pInicio = Format(dtpInicio.Value, "yyyy/mm/dd") & " 00:00:00"
pCorte = Format(fxFechaServidor, "yyyy/mm/dd") & " 23:59:59"
        
pFiltro = ""
pDetalle = "Operacion: " & txtOperacion.Text
        
vPaso = True
Call sbARF_Asientos_Main(lswAsientoMain, pInicio, pCorte, pFiltro, pDetalle)
        
vPaso = False
        
scAsientos.Caption = "Seleccione un asiento"
lswAsiento.ListItems.Clear
        
End Sub

Private Sub tlbAux_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim i As Integer

Select Case Button.Key
  Case "Autorizacion"
'     GLOBALES.gTag = txtOperacion.Text
'     Call sbFormsCall("frmARF_OPERACIONESSGTAutorizacion", 1, , , False, Me)
     
  Case "Activar"
        If fxActivacionVerifica Then
           
            i = MsgBox("Esta seguro que desea >> Activar << esta Operación", vbYesNo)
            If i = vbYes Then
                Call sbActivar
            End If
            
        Else 'Falla Verificacion de Formalizacion
         MsgBox vMensaje, vbCritical
        End If
  
  Case "Anular"

        If fxAnulacionVerifica Then
           i = MsgBox("Esta seguro que desea >> Anular << esta Operación", vbYesNo)
           If i = vbYes Then
              Call sbAnular
           End If
        Else
          MsgBox vMensaje, vbCritical
        End If
End Select

Call sbConsulta


End Sub

Private Sub tlbPrincipal_ButtonClick(ByVal Button As MSComctlLib.Button)

On Error Resume Next

Select Case Button.Key
 Case "nuevo"
  txtOperacion.Text = ""
  txtOperacion.Enabled = False
  
  Call sbLimpiaDatos
  tlbPrincipal.Buttons(1).Enabled = False
  tlbPrincipal.Buttons(2).Enabled = False
  tlbPrincipal.Buttons(3).Enabled = True
  tlbPrincipal.Buttons(4).Enabled = True
  
  txtNotas.Locked = False
  
  vEdita = False
  
  
  txtUnidadCod.SetFocus
  
  
  
 Case "editar"
  If CLng(txtOperacion.Text) > 0 Then  'And Operacion.Estado = "A" Then
      vEdita = True
'      Call Edicion(1)
    
      'Si el Estado Esta en Recepcion o Resolucion puede Cambiar Todos Los Datos
      'Si Esta en Formalización Solo puede Cambiar la Salida
      tlbPrincipal.Buttons(1).Enabled = False
      tlbPrincipal.Buttons(2).Enabled = False
      tlbPrincipal.Buttons(3).Enabled = True
      tlbPrincipal.Buttons(4).Enabled = True
      
      txtOperacion.Enabled = False

      txtNotas.Locked = False
      txtUnidadCod.SetFocus

  End If
 
 Case "guardar"
  
  If fxVerificaRecepcion Then
    Call sbGuardar
    Call sbConsulta
    txtOperacion.Enabled = True
    
    tlbPrincipal.Buttons(1).Enabled = True
    tlbPrincipal.Buttons(2).Enabled = True
    tlbPrincipal.Buttons(3).Enabled = False
    tlbPrincipal.Buttons(4).Enabled = False
    
    txtNotas.Locked = True
  End If
 
 Case "deshacer"
    txtOperacion.Enabled = True
    tlbPrincipal.Buttons(1).Enabled = True
    tlbPrincipal.Buttons(2).Enabled = True
    tlbPrincipal.Buttons(3).Enabled = False
    tlbPrincipal.Buttons(4).Enabled = False
    
    If txtOperacion <> "" Then Call sbConsulta
    txtOperacion.SetFocus
 
 Case "ayuda"
        frmContenedor.CD.HelpContext = Me.HelpContextID
        frmContenedor.CD.ShowHelp
 
 Case "cerrar"
    UnLoad Me

End Select


End Sub




Private Sub txtArrendadorCod_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtArrendadorDesc.SetFocus

If KeyCode = vbKeyF4 Then
  gBusquedas.Convertir = "N"
  gBusquedas.Columna = "COD_ACREEDOR"
  gBusquedas.Orden = "COD_ACREEDOR"
  gBusquedas.Consulta = "select COD_ACREEDOR, Descripcion from ARF_ACREEDORES"
  gBusquedas.Filtro = ""
  frmBusquedas.Show vbModal
  If IsNumeric(gBusquedas.Resultado) Then
    txtArrendadorCod.Text = gBusquedas.Resultado
    txtArrendadorDesc.Text = gBusquedas.Resultado2
  End If

End If
End Sub


Private Sub txtArrendadorDesc_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtMonto.SetFocus
End Sub


Private Sub txtGarantiaMonto_GotFocus()
On Error GoTo vError

txtGarantiaMonto.Text = CCur(txtGarantiaMonto.Text)

vError:

End Sub

Private Sub txtGarantiaMonto_LostFocus()
On Error GoTo vError

txtGarantiaMonto.Text = Format(txtGarantiaMonto.Text, "Standard")

vError:

End Sub


Private Sub txtMonto_GotFocus()
On Error GoTo vError

txtMonto.Text = CCur(txtMonto.Text)

vError:

End Sub

Private Sub txtMonto_LostFocus()
On Error GoTo vError

txtMonto.Text = Format(txtMonto.Text, "Standard")

vError:

End Sub

Private Sub txtTasaDescuento_GotFocus()
On Error GoTo vError

txtTasaDescuento.Text = CCur(txtTasaDescuento.Text)

vError:

End Sub

Private Sub txtTasaDescuento_LostFocus()
On Error GoTo vError

txtTasaDescuento.Text = Format(txtTasaDescuento.Text, "Standard")

vError:

End Sub


Private Sub txtTasaInteres_GotFocus()
On Error GoTo vError

txtTasaInteres.Text = CCur(txtTasaInteres.Text)

vError:

End Sub

Private Sub txtTasaInteres_LostFocus()
On Error GoTo vError

txtTasaInteres.Text = Format(txtTasaInteres.Text, "Standard")

vError:

End Sub

Private Sub txtIncremento_GotFocus()
On Error GoTo vError

txtIncremento.Text = CCur(txtIncremento.Text)

vError:

End Sub

Private Sub txtIncremento_LostFocus()
On Error GoTo vError

txtIncremento.Text = Format(txtIncremento.Text, "Standard")

vError:

End Sub


Private Sub txtUnidadCod_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtUnidadDesc.SetFocus

If KeyCode = vbKeyF4 Then
  gBusquedas.Convertir = "N"
  gBusquedas.Columna = "COD_LOCAL"
  gBusquedas.Orden = "COD_LOCAL"
  gBusquedas.Consulta = "select COD_LOCAL, Descripcion from ARF_UNIDADES"
  gBusquedas.Filtro = ""
  frmBusquedas.Show vbModal
  txtUnidadCod.Text = gBusquedas.Resultado
  txtUnidadDesc.Text = gBusquedas.Resultado2
End If

End Sub
