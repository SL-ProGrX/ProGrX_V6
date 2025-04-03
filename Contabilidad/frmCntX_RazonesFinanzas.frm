VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#20.3#0"; "Codejock.Controls.v20.3.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#20.3#0"; "Codejock.ShortcutBar.v20.3.0.ocx"
Begin VB.Form frmCntX_RazonesFinanzas 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Razones Financieras"
   ClientHeight    =   7365
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   10185
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7365
   ScaleWidth      =   10185
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.TabControl tcMain 
      Height          =   6132
      Left            =   0
      TabIndex        =   2
      Top             =   1080
      Width           =   10092
      _Version        =   1310723
      _ExtentX        =   17801
      _ExtentY        =   10816
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
      Item(0).Caption =   "Tipos de Razones"
      Item(0).ControlCount=   1
      Item(0).Control(0)=   "vGrid"
      Item(1).Caption =   "Razones"
      Item(1).ControlCount=   1
      Item(1).Control(0)=   "vGridFx"
      Item(2).Caption =   "Fórmulas"
      Item(2).ControlCount=   16
      Item(2).Control(0)=   "cboGrupo"
      Item(2).Control(1)=   "txtCodigo"
      Item(2).Control(2)=   "txtNombre"
      Item(2).Control(3)=   "Label3(5)"
      Item(2).Control(4)=   "Label3(6)"
      Item(2).Control(5)=   "FlatScrollBarX"
      Item(2).Control(6)=   "lsw"
      Item(2).Control(7)=   "cboOperador"
      Item(2).Control(8)=   "txtCuentaCod"
      Item(2).Control(9)=   "txtCuentaDesc"
      Item(2).Control(10)=   "Label3(3)"
      Item(2).Control(11)=   "btnBarra(0)"
      Item(2).Control(12)=   "btnBarra(2)"
      Item(2).Control(13)=   "btnBarra(1)"
      Item(2).Control(14)=   "scBarra"
      Item(2).Control(15)=   "tcFx"
      Item(3).Caption =   "Informes"
      Item(3).ControlCount=   16
      Item(3).Control(0)=   "lswRepFx"
      Item(3).Control(1)=   "chkTodos"
      Item(3).Control(2)=   "cmdReporte"
      Item(3).Control(3)=   "cboRepGrupo"
      Item(3).Control(4)=   "cboUnidades"
      Item(3).Control(5)=   "txtPeriodo"
      Item(3).Control(6)=   "txtAnio"
      Item(3).Control(7)=   "txtMes"
      Item(3).Control(8)=   "scTitulo"
      Item(3).Control(9)=   "optRazon(0)"
      Item(3).Control(10)=   "optRazon(1)"
      Item(3).Control(11)=   "chkNotas"
      Item(3).Control(12)=   "Label3(0)"
      Item(3).Control(13)=   "Label3(1)"
      Item(3).Control(14)=   "Label3(2)"
      Item(3).Control(15)=   "lblGenera"
      Begin XtremeSuiteControls.ListView lsw 
         Height          =   1812
         Left            =   -69520
         TabIndex        =   11
         Top             =   4080
         Visible         =   0   'False
         Width           =   9132
         _Version        =   1310723
         _ExtentX        =   16108
         _ExtentY        =   3196
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
         View            =   3
         FullRowSelect   =   -1  'True
         Appearance      =   16
         ShowBorder      =   0   'False
      End
      Begin XtremeSuiteControls.ListView lswRepFx 
         Height          =   2652
         Left            =   -69760
         TabIndex        =   20
         Top             =   1800
         Visible         =   0   'False
         Width           =   9612
         _Version        =   1310723
         _ExtentX        =   16954
         _ExtentY        =   4678
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
         ShowBorder      =   0   'False
      End
      Begin XtremeSuiteControls.TabControl tcFx 
         Height          =   1812
         Left            =   -69520
         TabIndex        =   36
         Top             =   1200
         Visible         =   0   'False
         Width           =   9132
         _Version        =   1310723
         _ExtentX        =   16108
         _ExtentY        =   3196
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
         ItemCount       =   2
         Item(0).Caption =   "Notas"
         Item(0).ControlCount=   1
         Item(0).Control(0)=   "txtNotas"
         Item(1).Caption =   "Fórmula"
         Item(1).ControlCount=   3
         Item(1).Control(0)=   "txtFormula"
         Item(1).Control(1)=   "imgGuardaFormula"
         Item(1).Control(2)=   "imgfxVerifica"
         Begin XtremeSuiteControls.FlatEdit txtNotas 
            Height          =   1392
            Left            =   1440
            TabIndex        =   37
            Top             =   360
            Width           =   7692
            _Version        =   1310723
            _ExtentX        =   13568
            _ExtentY        =   2455
            _StockProps     =   77
            ForeColor       =   16711680
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
            MultiLine       =   -1  'True
            ScrollBars      =   2
            Appearance      =   2
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtFormula 
            Height          =   1392
            Left            =   -68560
            TabIndex        =   38
            Top             =   360
            Visible         =   0   'False
            Width           =   7692
            _Version        =   1310723
            _ExtentX        =   13568
            _ExtentY        =   2455
            _StockProps     =   77
            ForeColor       =   65535
            BackColor       =   12582912
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColor       =   12582912
            Alignment       =   2
            MultiLine       =   -1  'True
            ScrollBars      =   2
            Appearance      =   2
            UseVisualStyle  =   0   'False
         End
         Begin VB.Image imgfxVerifica 
            Height          =   240
            Left            =   -69165
            Picture         =   "frmCntX_RazonesFinanzas.frx":0000
            Top             =   1080
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.Image imgGuardaFormula 
            Height          =   240
            Left            =   -69165
            Picture         =   "frmCntX_RazonesFinanzas.frx":00E9
            Top             =   720
            Visible         =   0   'False
            Width           =   240
         End
      End
      Begin FPSpreadADO.fpSpread vGrid 
         Height          =   5412
         Left            =   1200
         TabIndex        =   3
         Top             =   480
         Width           =   7692
         _Version        =   524288
         _ExtentX        =   13568
         _ExtentY        =   9546
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
         MaxCols         =   493
         ScrollBars      =   2
         SpreadDesigner  =   "frmCntX_RazonesFinanzas.frx":080A
         VScrollSpecial  =   -1  'True
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin FPSpreadADO.fpSpread vGridFx 
         Height          =   5412
         Left            =   -69640
         TabIndex        =   4
         Top             =   480
         Visible         =   0   'False
         Width           =   9612
         _Version        =   524288
         _ExtentX        =   16954
         _ExtentY        =   9546
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
         SpreadDesigner  =   "frmCntX_RazonesFinanzas.frx":0D76
         VScrollSpecial  =   -1  'True
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin XtremeSuiteControls.ComboBox cboGrupo 
         Height          =   312
         Left            =   -68080
         TabIndex        =   5
         Top             =   480
         Visible         =   0   'False
         Width           =   7692
         _Version        =   1310723
         _ExtentX        =   13573
         _ExtentY        =   582
         _StockProps     =   77
         ForeColor       =   1973790
         BackColor       =   16185078
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16185078
         Style           =   2
         Appearance      =   16
         UseVisualStyle  =   0   'False
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.FlatEdit txtCodigo 
         Height          =   312
         Left            =   -68080
         TabIndex        =   6
         Top             =   840
         Visible         =   0   'False
         Width           =   1452
         _Version        =   1310723
         _ExtentX        =   2561
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
         Left            =   -66640
         TabIndex        =   7
         Top             =   840
         Visible         =   0   'False
         Width           =   6252
         _Version        =   1310723
         _ExtentX        =   11028
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
      Begin XtremeSuiteControls.ScrollBar FlatScrollBarX 
         Height          =   252
         Left            =   -60400
         TabIndex        =   10
         Top             =   840
         Visible         =   0   'False
         Width           =   492
         _Version        =   1310723
         _ExtentX        =   868
         _ExtentY        =   0
         _StockProps     =   64
         UseVisualStyle  =   0   'False
         Appearance      =   16
      End
      Begin XtremeSuiteControls.ComboBox cboOperador 
         Height          =   312
         Left            =   -68560
         TabIndex        =   12
         Top             =   3720
         Visible         =   0   'False
         Width           =   2172
         _Version        =   1310723
         _ExtentX        =   3836
         _ExtentY        =   582
         _StockProps     =   77
         ForeColor       =   1973790
         BackColor       =   16185078
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16185078
         Style           =   2
         Appearance      =   16
         UseVisualStyle  =   0   'False
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.FlatEdit txtCuentaCod 
         Height          =   312
         Left            =   -68560
         TabIndex        =   13
         Top             =   3120
         Visible         =   0   'False
         Width           =   2172
         _Version        =   1310723
         _ExtentX        =   3831
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
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtCuentaDesc 
         Height          =   312
         Left            =   -66400
         TabIndex        =   14
         Top             =   3120
         Visible         =   0   'False
         Width           =   6012
         _Version        =   1310723
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
      Begin XtremeSuiteControls.PushButton btnBarra 
         Height          =   312
         Index           =   0
         Left            =   -65680
         TabIndex        =   16
         ToolTipText     =   "Nuevo"
         Top             =   3720
         Visible         =   0   'False
         Width           =   1092
         _Version        =   1310723
         _ExtentX        =   1926
         _ExtentY        =   550
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
         Picture         =   "frmCntX_RazonesFinanzas.frx":12FE
         ImageAlignment  =   4
      End
      Begin XtremeSuiteControls.PushButton btnBarra 
         Height          =   312
         Index           =   2
         Left            =   -64240
         TabIndex        =   17
         ToolTipText     =   "Eliminar"
         Top             =   3720
         Visible         =   0   'False
         Width           =   372
         _Version        =   1310723
         _ExtentX        =   656
         _ExtentY        =   550
         _StockProps     =   79
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
         Picture         =   "frmCntX_RazonesFinanzas.frx":1930
         ImageAlignment  =   6
      End
      Begin XtremeSuiteControls.PushButton btnBarra 
         Height          =   312
         Index           =   1
         Left            =   -64600
         TabIndex        =   18
         ToolTipText     =   "Guardar"
         Top             =   3720
         Visible         =   0   'False
         Width           =   372
         _Version        =   1310723
         _ExtentX        =   656
         _ExtentY        =   550
         _StockProps     =   79
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
         Picture         =   "frmCntX_RazonesFinanzas.frx":1ED4
         ImageAlignment  =   6
      End
      Begin XtremeSuiteControls.CheckBox chkTodos 
         Height          =   216
         Left            =   -69640
         TabIndex        =   21
         Top             =   1500
         Visible         =   0   'False
         Width           =   216
         _Version        =   1310723
         _ExtentX        =   370
         _ExtentY        =   370
         _StockProps     =   79
         BackColor       =   -2147483633
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton cmdReporte 
         Height          =   612
         Left            =   -61720
         TabIndex        =   22
         Top             =   5040
         Visible         =   0   'False
         Width           =   1452
         _Version        =   1310723
         _ExtentX        =   2561
         _ExtentY        =   1080
         _StockProps     =   79
         Caption         =   "Reporte"
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   14
         Picture         =   "frmCntX_RazonesFinanzas.frx":2605
      End
      Begin XtremeSuiteControls.ComboBox cboRepGrupo 
         Height          =   312
         Left            =   -67960
         TabIndex        =   23
         Top             =   960
         Visible         =   0   'False
         Width           =   5532
         _Version        =   1310723
         _ExtentX        =   9763
         _ExtentY        =   582
         _StockProps     =   77
         ForeColor       =   1973790
         BackColor       =   16185078
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16185078
         Style           =   2
         Appearance      =   16
         UseVisualStyle  =   0   'False
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.ComboBox cboUnidades 
         Height          =   312
         Left            =   -67960
         TabIndex        =   24
         Top             =   4560
         Visible         =   0   'False
         Width           =   5532
         _Version        =   1310723
         _ExtentX        =   9763
         _ExtentY        =   582
         _StockProps     =   77
         ForeColor       =   1973790
         BackColor       =   16185078
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16185078
         Style           =   2
         Appearance      =   16
         UseVisualStyle  =   0   'False
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.FlatEdit txtPeriodo 
         Height          =   312
         Left            =   -66880
         TabIndex        =   25
         Top             =   600
         Visible         =   0   'False
         Width           =   4452
         _Version        =   1310723
         _ExtentX        =   7853
         _ExtentY        =   550
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
         Alignment       =   2
         Locked          =   -1  'True
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtAnio 
         Height          =   312
         Left            =   -67960
         TabIndex        =   26
         Top             =   600
         Visible         =   0   'False
         Width           =   624
         _Version        =   1310723
         _ExtentX        =   1101
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
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtMes 
         Height          =   312
         Left            =   -67360
         TabIndex        =   27
         Top             =   600
         Visible         =   0   'False
         Width           =   504
         _Version        =   1310723
         _ExtentX        =   889
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
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.RadioButton optRazon 
         Height          =   252
         Index           =   0
         Left            =   -67960
         TabIndex        =   29
         Top             =   5040
         Visible         =   0   'False
         Width           =   2652
         _Version        =   1310723
         _ExtentX        =   4678
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Comparativo Anual"
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
         Appearance      =   16
         Value           =   -1  'True
         Alignment       =   1
      End
      Begin XtremeSuiteControls.RadioButton optRazon 
         Height          =   252
         Index           =   1
         Left            =   -67960
         TabIndex        =   30
         Top             =   5400
         Visible         =   0   'False
         Width           =   2652
         _Version        =   1310723
         _ExtentX        =   4678
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Comparativo Trimestral"
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
         Appearance      =   16
         Alignment       =   1
      End
      Begin XtremeSuiteControls.CheckBox chkNotas 
         Height          =   372
         Left            =   -64480
         TabIndex        =   31
         Top             =   5160
         Visible         =   0   'False
         Width           =   1932
         _Version        =   1310723
         _ExtentX        =   3408
         _ExtentY        =   656
         _StockProps     =   79
         Caption         =   "Visualizar Notas de la Razón"
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
         Appearance      =   16
         Alignment       =   1
      End
      Begin XtremeSuiteControls.Label lblGenera 
         Height          =   612
         Left            =   -62200
         TabIndex        =   35
         Top             =   600
         Visible         =   0   'False
         Width           =   2172
         _Version        =   1310723
         _ExtentX        =   3831
         _ExtentY        =   1080
         _StockProps     =   79
         Caption         =   "..."
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
         WordWrap        =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label3 
         Height          =   252
         Index           =   2
         Left            =   -69280
         TabIndex        =   34
         Top             =   960
         Visible         =   0   'False
         Width           =   1212
         _Version        =   1310723
         _ExtentX        =   2138
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Tipo Razón"
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
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label3 
         Height          =   252
         Index           =   1
         Left            =   -69280
         TabIndex        =   33
         Top             =   600
         Visible         =   0   'False
         Width           =   1212
         _Version        =   1310723
         _ExtentX        =   2138
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Periodo"
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
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label3 
         Height          =   252
         Index           =   0
         Left            =   -69280
         TabIndex        =   32
         Top             =   4560
         Visible         =   0   'False
         Width           =   1212
         _Version        =   1310723
         _ExtentX        =   2138
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Unidades"
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
         Transparent     =   -1  'True
      End
      Begin XtremeShortcutBar.ShortcutCaption scTitulo 
         Height          =   372
         Left            =   -69760
         TabIndex        =   28
         Top             =   1440
         Visible         =   0   'False
         Width           =   9612
         _Version        =   1310723
         _ExtentX        =   16954
         _ExtentY        =   656
         _StockProps     =   14
         Caption         =   "Razones Disponibles"
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
      Begin XtremeShortcutBar.ShortcutCaption scBarra 
         Height          =   492
         Left            =   -69520
         TabIndex        =   19
         Top             =   3600
         Visible         =   0   'False
         Width           =   9132
         _Version        =   1310723
         _ExtentX        =   16108
         _ExtentY        =   868
         _StockProps     =   14
         Caption         =   "Operador"
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
         VisualTheme     =   3
      End
      Begin XtremeSuiteControls.Label Label3 
         Height          =   252
         Index           =   3
         Left            =   -69520
         TabIndex        =   15
         Top             =   3120
         Visible         =   0   'False
         Width           =   1212
         _Version        =   1310723
         _ExtentX        =   2138
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Cuenta"
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
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label3 
         Height          =   252
         Index           =   6
         Left            =   -69040
         TabIndex        =   9
         Top             =   840
         Visible         =   0   'False
         Width           =   1212
         _Version        =   1310723
         _ExtentX        =   2138
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Razón"
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
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label3 
         Height          =   252
         Index           =   5
         Left            =   -69040
         TabIndex        =   8
         Top             =   480
         Visible         =   0   'False
         Width           =   1212
         _Version        =   1310723
         _ExtentX        =   2138
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Tipo"
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
         Transparent     =   -1  'True
      End
   End
   Begin MSComctlLib.ProgressBar ProgressBarX 
      Align           =   2  'Align Bottom
      Height          =   132
      Left            =   0
      TabIndex        =   0
      Top             =   7236
      Width           =   10188
      _ExtentX        =   17965
      _ExtentY        =   238
      _Version        =   393216
      Appearance      =   0
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Razones Financieras"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   372
      Index           =   6
      Left            =   1800
      TabIndex        =   1
      Top             =   360
      Width           =   8532
   End
   Begin VB.Image imgBanner 
      Appearance      =   0  'Flat
      Height          =   972
      Left            =   0
      Top             =   0
      Width           =   12852
   End
End
Attribute VB_Name = "frmCntX_RazonesFinanzas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strUltimaSeleccion As String, vScroll As Boolean
Dim vPaso As Boolean



Private Sub sbCargaRepLsw()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem


'Saca solo las razones que tienen formula, con detalle Base
strSQL = "select R.cod_razon,R.descripcion" _
       & " from CntX_Razones R inner join CntX_Razones_detalle D" _
       & " on R.cod_razon = D.cod_razon and R.cod_contabilidad = D.cod_contabilidad" _
       & " where R.cod_contabilidad = " & gCntX_Parametros.CodigoConta _
       & " and D.operador = 'B'"

If cboRepGrupo.Text <> "TODOS" Then
  strSQL = strSQL & " and R.cod_grupo = '" & cboRepGrupo.ItemData(cboRepGrupo.ListIndex) & "'"
End If
       

      
strSQL = strSQL & " group by R.cod_razon,R.descripcion"

lswRepFx.ListItems.Clear

Call OpenRecordSet(rs, strSQL)

Do While Not rs.EOF
 Set itmX = lswRepFx.ListItems.Add(, , "")
     itmX.SubItems(1) = Trim(rs!cod_razon)
     itmX.SubItems(2) = rs!Descripcion
     itmX.Checked = chkTodos.Value
 rs.MoveNext
Loop
rs.Close

End Sub

Private Sub btnBarra_Click(Index As Integer)
Dim strSQL As String

Select Case Index
 Case 0 'Nuevo
    txtCuentaCod.Tag = ""
    txtCuentaCod.Text = ""
    txtCuentaDesc.Text = ""
    
 Case 1 'Actualiza
   Call sbRazonActualiza
  
 Case 2 'Borrar
    If txtCuentaCod.Tag = "" Then
      MsgBox "Selecciones Primero la cuenta que desea eliminar...", vbExclamation
    Else
      strSQL = "delete CntX_Razones_detalle where idX = " & txtCuentaCod.Tag _
             & " and cod_contabilidad = " & gCntX_Parametros.CodigoConta _
             & " and cod_razon = '" & txtCodigo & "' and operador <> 'B'"
      Call ConectionExecute(strSQL, 0)
      
      Call sbFxCarga(cboGrupo.ItemData(cboGrupo.ListIndex), txtCodigo)
    End If
End Select
End Sub

Private Sub cboRepGrupo_Click()

If vPaso Then Exit Sub
Call sbCargaRepLsw

End Sub


Private Sub chkTodos_Click()
Dim i As Integer

For i = 1 To lswRepFx.ListItems.Count
  lswRepFx.ListItems.Item(i).Checked = chkTodos.Value
Next i

End Sub


Private Function fxFormula(vRazon As String) As String
Dim strSQL As String, rs As New ADODB.Recordset
Dim vResultado As String, idX As Integer
'Crea la formula Basica

vResultado = ""

strSQL = "select * from CntX_Razones_detalle where cod_contabilidad = " & gCntX_Parametros.CodigoConta _
       & " and cod_razon = '" & vRazon & "' and operador = 'B'"
Call OpenRecordSet(rs, strSQL, 0)
If Not rs.EOF And Not rs.BOF Then
 vResultado = "C" & rs!idX
 idX = rs!idX
End If
rs.Close

strSQL = "select * from CntX_Razones_detalle where cod_contabilidad = " & gCntX_Parametros.CodigoConta _
       & " and cod_razon = '" & vRazon & "' and idX <> " & idX & " Order by IdX"
Call OpenRecordSet(rs, strSQL, 0)
Do While Not rs.EOF
 
 Select Case rs!operador
   Case "R" 'Restar
     vResultado = vResultado & " - "
   Case "M" 'Multiplicar
     vResultado = vResultado & " * "
   Case "D" 'Dividir
     vResultado = vResultado & " / "
   Case "S" 'Sumar
     vResultado = vResultado & " + "
 End Select
 
 vResultado = vResultado & "C" & rs!idX
 
 rs.MoveNext
Loop
rs.Close

fxFormula = vResultado

End Function

Private Sub sbRazonActualiza()
Dim strSQL As String, rs As New ADODB.Recordset
Dim i As Integer

On Error GoTo vError

If txtCuentaCod.Tag = "" Then
 strSQL = "select (isnull(max(idX),0) + 1) as IDx from CntX_Razones_detalle" _
        & " where cod_contabilidad = " & gCntX_Parametros.CodigoConta _
        & " and cod_razon = '" & txtCodigo & "'"
 Call OpenRecordSet(rs, strSQL, 0)
 i = rs!idX
 rs.Close
 
 If Mid(cboOperador.Text, 1, 1) = "B" Then
 'Verificar que no exista ya una cuenta BASE
   strSQL = "select IDx from CntX_Razones_detalle" _
           & " where cod_contabilidad = " & gCntX_Parametros.CodigoConta _
           & " and cod_razon = '" & txtCodigo & "' and operador = 'B'"
    Call OpenRecordSet(rs, strSQL, 0)
    If Not rs.EOF And Not rs.BOF Then
      MsgBox "YA existe una cuenta BASE, no se puede Ingresar esta cuenta...", vbExclamation
      Exit Sub
    End If
    rs.Close
 
 Else
 'Verifica que EXISTA una cuenta BASE
   strSQL = "select IDx from CntX_Razones_detalle" _
           & " where cod_contabilidad = " & gCntX_Parametros.CodigoConta _
           & " and cod_razon = '" & txtCodigo & "' and operador = 'B'"
    Call OpenRecordSet(rs, strSQL, 0)
    If rs.EOF And rs.BOF Then
      MsgBox "NO existe una cuenta BASE, no se puede Ingresar esta cuenta...", vbExclamation
      Exit Sub
    End If
    rs.Close
 
 End If
 
 strSQL = "insert CntX_Razones_detalle(idx,cod_contabilidad,cod_razon,cod_cuenta,operador) values(" _
        & i & "," & gCntX_Parametros.CodigoConta & ",'" & txtCodigo & "','" & fxCntX_CuentaFormato(False, txtCuentaCod) _
        & "','" & Mid(cboOperador.Text, 1, 1) & "')"
 Call ConectionExecute(strSQL, 0)


Else
 'Actualiza
 i = txtCuentaCod.Tag
 If Mid(cboOperador.Text, 1, 1) = "B" Then
 'Verificar que no exista ya una cuenta BASE
   strSQL = "select IDx from CntX_Razones_detalle" _
           & " where cod_contabilidad = " & gCntX_Parametros.CodigoConta _
           & " and cod_razon = '" & txtCodigo & "' and operador = 'B'" _
           & " and idX <> " & i
    Call OpenRecordSet(rs, strSQL, 0)
    If Not rs.EOF And Not rs.BOF Then
      MsgBox "YA existe una cuenta BASE, no se puede Ingresar esta cuenta...", vbExclamation
      Exit Sub
    End If
    rs.Close
 
 Else
 'Verifica que EXISTA una cuenta BASE
   strSQL = "select IDx from CntX_Razones_detalle" _
           & " where cod_contabilidad = " & gCntX_Parametros.CodigoConta _
           & " and cod_razon = '" & txtCodigo & "' and operador = 'B'" _
           & " and idx <> " & i
    Call OpenRecordSet(rs, strSQL, 0)
    If rs.EOF And rs.BOF Then
      MsgBox "NO existe una cuenta BASE, no se puede Ingresar esta cuenta...", vbExclamation
      Exit Sub
    End If
    rs.Close
 
 End If
 
 strSQL = "update CntX_Razones_detalle set cod_cuenta = '" & fxCntX_CuentaFormato(False, txtCuentaCod) _
        & "',operador = '" & Mid(cboOperador.Text, 1, 1) & "' where cod_contabilidad = " & gCntX_Parametros.CodigoConta _
        & " and cod_razon = '" & txtCodigo & "' and idX = " & i
 Call ConectionExecute(strSQL, 0)

End If

strSQL = "update CntX_Razones set notas = '" & txtNotas _
       & "',formula = '" & fxFormula(txtCodigo) _
       & "' where cod_contabilidad = " & gCntX_Parametros.CodigoConta _
       & " and cod_grupo = '" & cboGrupo.ItemData(cboGrupo.ListIndex) _
       & "' and cod_razon = '" & txtCodigo & "'"
Call ConectionExecute(strSQL, 0)




Call sbFxCarga(cboGrupo.ItemData(cboGrupo.ListIndex), txtCodigo)

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub cmdReporte_Click()
Dim strSQL As String, i As Integer, vSubTitulo As String
Dim lngAnio As Long, iMes As Integer
Dim curMonto As Currency, x As Integer
Dim vPeriodo_01 As String, vPeriodo_02 As String, vPeriodo_03 As String

On Error GoTo vError

'Crea la barra de progreso
x = 0
For i = 1 To lswRepFx.ListItems.Count
 If lswRepFx.ListItems.Item(i).Checked Then
       x = x + 1
 End If
Next i

If x = 0 Then
 MsgBox "No se seleccionó ninguna Razón Financiera a evaluar?", vbExclamation
 Exit Sub
End If


Me.MousePointer = vbHourglass
lblGenera.Caption = "Generando Razones..."


ProgressBarX.Visible = True
ProgressBarX.Max = x * 3
ProgressBarX.Value = 0

'Borra Informacion Anterior del Usuario
strSQL = "delete CntX_Razones_Reporte where usuario = '" & glogon.Usuario _
       & "' and cod_contabilidad = " & gCntX_Parametros.CodigoConta
Call ConectionExecute(strSQL, 0)



'Razones Financieras del Periodo Actual
For i = 1 To lswRepFx.ListItems.Count
 If lswRepFx.ListItems.Item(i).Checked Then
   curMonto = fxResRazon(lswRepFx.ListItems.Item(i).SubItems(1), txtAnio, txtMes)
   strSQL = "insert CntX_Razones_Reporte(usuario,cod_contabilidad,cod_razon,monto) values('" _
          & glogon.Usuario & "'," & gCntX_Parametros.CodigoConta & ",'" _
          & lswRepFx.ListItems.Item(i).SubItems(1) & "'," & curMonto & ")"
   Call ConectionExecute(strSQL, 0)
   
   If ProgressBarX.Max > ProgressBarX.Value Then ProgressBarX.Value = ProgressBarX.Value + 1
 End If
Next i


lngAnio = txtAnio
iMes = txtMes

vPeriodo_01 = fxCntX_PeriodoDesc(lngAnio, iMes)
vPeriodo_02 = ""
vPeriodo_03 = ""

For x = 1 To 2
   'Mes Anterior
   Select Case True
      Case optRazon.Item(0).Value 'Comprativo Anual
            lngAnio = lngAnio - 1
            vSubTitulo = optRazon.Item(0).Caption
      
      Case optRazon.Item(1).Value 'Comprativo Ultimo Trimestre
            If iMes = 1 Then
               iMes = 12
               lngAnio = lngAnio - 1
            Else
               iMes = iMes - 1
            End If
            
            vSubTitulo = optRazon.Item(1).Caption
   End Select
   
   
   If x = 1 Then
     vPeriodo_02 = fxCntX_PeriodoDesc(lngAnio, iMes)
   Else
     vPeriodo_03 = fxCntX_PeriodoDesc(lngAnio, iMes)
   End If
   
   
   For i = 1 To lswRepFx.ListItems.Count
        If lswRepFx.ListItems.Item(i).Checked Then
            curMonto = fxResRazon(lswRepFx.ListItems.Item(i).SubItems(1), lngAnio, iMes)
            strSQL = "update CntX_Razones_Reporte set Mes" & Format(x, "00") & " = " & curMonto _
                   & " where usuario = '" & glogon.Usuario & "' and cod_contabilidad = " & gCntX_Parametros.CodigoConta _
                   & " and cod_razon = '" & lswRepFx.ListItems.Item(i).SubItems(1) & "'"
            Call ConectionExecute(strSQL, 0)
        
            If ProgressBarX.Max > ProgressBarX.Value Then ProgressBarX.Value = ProgressBarX.Value + 1
        End If
   Next i 'Razones

Next x 'Periodos a Procesar


lblGenera.Caption = ""
ProgressBarX.Visible = False

If cboUnidades.Text = "TODOS" Then
  vSubTitulo = vSubTitulo & "  ¦  Unidad : " & cboUnidades.Text
Else
  vSubTitulo = vSubTitulo & "  ¦  Contabilidad General"
End If

strSQL = "{CntX_Razones_Reporte.cod_contabilidad} = " & gCntX_Parametros.CodigoConta _
       & " AND {CntX_Razones_Reporte.USUARIO} = '" & glogon.Usuario & "'"

Call sbCntX_Reportes("RAZONES", strSQL, vSubTitulo, vPeriodo_01, vPeriodo_02, vPeriodo_03, chkNotas.Value)

Me.MousePointer = vbDefault

Exit Sub

vError:
 lblGenera.Caption = ""
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub FlatScrollBarX_Change()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

If vScroll Then
    strSQL = "select Top 1 cod_grupo,cod_razon,descripcion from CntX_Razones" _
           & " where cod_contabilidad = " & gCntX_Parametros.CodigoConta _
           & " and cod_grupo = '" & cboGrupo.ItemData(cboGrupo.ListIndex) & "'"
    
    If FlatScrollBarX.Value = 1 Then
       strSQL = strSQL & " and cod_razon > '" & txtCodigo.Text & "' order by cod_razon asc"
    Else
       strSQL = strSQL & " and cod_razon < '" & txtCodigo.Text & "' order by cod_razon desc"
    End If
    
    Call OpenRecordSet(rs, strSQL, 0)
    If Not rs.EOF And Not rs.BOF Then
      txtCodigo.Text = rs!cod_razon
      txtNombre.Text = rs!Descripcion
      Call sbFxCarga(rs!cod_grupo, rs!cod_razon)
    End If
    rs.Close
End If

vScroll = False
 FlatScrollBarX.Value = 0
vScroll = True

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub



Private Sub sbPredeterminados()
Dim strSQL As String, rs As New ADODB.Recordset

'Si no existe ninguno creado, el sistema los ingresa por defecto
strSQL = "select isnull(count(*),0) as Existe from CntX_Razones_Tipos" _
       & " where cod_contabilidad = " & gCntX_Parametros.CodigoConta
Call OpenRecordSet(rs, strSQL, 0)
If rs!Existe > 0 Then
  Exit Sub
End If
rs.Close

'Crea Grupos
strSQL = "insert CntX_Razones_Tipos(cod_contabilidad,cod_grupo,descripcion,activa) values(" _
       & gCntX_Parametros.CodigoConta & ",'IE','Indices de Estabilidad',1)"
Call ConectionExecute(strSQL, 0)

        strSQL = "insert CntX_Razones(cod_contabilidad,cod_grupo,cod_razon,descripcion,notas,resultado) values(" _
               & gCntX_Parametros.CodigoConta & ",'IE','IE.1','Razón Circulante','" _
               & "Evalua la confianza que proporcionan los activos circulantes a los acreedores de corto plazo. f(x) = Activo circulante / Pasivo circulante','')"
        Call ConectionExecute(strSQL, 0)

        strSQL = "insert CntX_Razones(cod_contabilidad,cod_grupo,cod_razon,descripcion,notas,resultado) values(" _
               & gCntX_Parametros.CodigoConta & ",'IE','IE.2','Indice Prueba del Acido','" _
               & "Mide la forma en que los activos de mayor liquidez cubren y garantizan a los pasivos circulantes. f(x) = activos circulantes - Inventario " _
               & " / Pasivo Circulante','')"
        Call ConectionExecute(strSQL, 0)

        strSQL = "insert CntX_Razones(cod_contabilidad,cod_grupo,cod_razon,descripcion,notas,resultado) values(" _
               & gCntX_Parametros.CodigoConta & ",'IE','IE.3','Indice de Deuda','" _
               & "Mide el % de Financiamiento aportado por los acreedores dentro de la empresa. f(x) = Pasivo Total / Activo Total','')"
        Call ConectionExecute(strSQL, 0)

        strSQL = "insert CntX_Razones(cod_contabilidad,cod_grupo,cod_razon,descripcion,notas,resultado) values(" _
               & gCntX_Parametros.CodigoConta & ",'IE','IE.4','Indice de Endeudamiento','" _
               & "Señala la relación entre los fondos que han financiado los acreedores y los recursos que aportan" _
               & " los accionistas de la empresa. f(x) = Pasivo Total / Patrimonio','')"
        Call ConectionExecute(strSQL, 0)

        strSQL = "insert CntX_Razones(cod_contabilidad,cod_grupo,cod_razon,descripcion,notas,resultado) values(" _
               & gCntX_Parametros.CodigoConta & ",'IE','IE.5','Cobertura de Intereses','" _
               & "Capacidad de la Empresa para pagar y cubrir la carga financiera del periodo con sus utilidades" _
               & " . f(x) = Utilidad de Operación / Gastos Financieros" & "','')"
        Call ConectionExecute(strSQL, 0)

        strSQL = "insert CntX_Razones(cod_contabilidad,cod_grupo,cod_razon,descripcion,notas,resultado) values(" _
               & gCntX_Parametros.CodigoConta & ",'IE','IE.6','Indice de Exposición Neta','" _
               & "Tendencia de pasivos en divisa extranjera, ante la expectativa de una devaluación futura de la divisa local" _
               & " . f(x) = Activos Divisa Extranjera / Pasivos Divisa Extranjera" & "','')"
        Call ConectionExecute(strSQL, 0)



strSQL = "insert CntX_Razones_Tipos(cod_contabilidad,cod_grupo,descripcion,activa) values(" _
       & gCntX_Parametros.CodigoConta & ",'IG','Indices de Gestión',1)"
Call ConectionExecute(strSQL, 0)

        strSQL = "insert CntX_Razones(cod_contabilidad,cod_grupo,cod_razon,descripcion,notas,resultado) values(" _
               & gCntX_Parametros.CodigoConta & ",'IG','IG.1','Rotación de Inventario','" _
               & "Mide la Liquidez del Inventario en una empresa. f(x) = Costo de Ventas " _
               & " / Inventario Promedio','')"
        Call ConectionExecute(strSQL, 0)
        
        strSQL = "insert CntX_Razones(cod_contabilidad,cod_grupo,cod_razon,descripcion,notas,resultado) values(" _
               & gCntX_Parametros.CodigoConta & ",'IG','IG.2','Periodo Medio de Inventario','" _
               & "Complementario a la Rotación de Inventario. Analiza la gestión y movimiento de de los inventarios desde una perspectiva diferente." _
               & " . f(x) = (Inventario Promedio  / Costo de Ventas) * 360 = 360/ Rotación de Inventarios" & "','')"
        Call ConectionExecute(strSQL, 0)
        
        
        strSQL = "insert CntX_Razones(cod_contabilidad,cod_grupo,cod_razon,descripcion,notas,resultado) values(" _
               & gCntX_Parametros.CodigoConta & ",'IG','IG.3','Rotación de las Cuentas por Cobrar','" _
               & "Evalúa la velocidad con que son transformadas en efectivo las cuentas x cobrar. f(x) = Ventas netas a crédito" _
               & " / Cuentas x Cobrar Promedio','')"
        Call ConectionExecute(strSQL, 0)
        
        strSQL = "insert CntX_Razones(cod_contabilidad,cod_grupo,cod_razon,descripcion,notas,resultado) values(" _
               & gCntX_Parametros.CodigoConta & ",'IG','IG.4','Periodo Medio de Cobro','" _
               & "Número Promedio de días para recuperar las ventas a crédito. f(x) = (cuentas por Cobrar / Ventas netas a crédito) * 360','')"
        Call ConectionExecute(strSQL, 0)
        
        strSQL = "insert CntX_Razones(cod_contabilidad,cod_grupo,cod_razon,descripcion,notas,resultado) values(" _
               & gCntX_Parametros.CodigoConta & ",'IG','IG.5','Periodo Medio de Pago','" _
               & "Plazo Promedio para pagar compras a crédito. f(x) = (cuentas por Pagar Promedio" _
               & " / Compras Netas a Crédito) * 360','')"
        Call ConectionExecute(strSQL, 0)
        
        strSQL = "insert CntX_Razones(cod_contabilidad,cod_grupo,cod_razon,descripcion,notas,resultado) values(" _
               & gCntX_Parametros.CodigoConta & ",'IG','IG.6','Rotación del Activo Circulante (RAC)','" _
               & "Activo Circulante de Corto Plazo se van Transformando de partidas poco líquidas a otras de mayor liquidez hasta terminar en efectivo." _
               & " f(x) = Ventas Netas / Activo Circulante Promedio','')"
        Call ConectionExecute(strSQL, 0)
        
        
        strSQL = "insert CntX_Razones(cod_contabilidad,cod_grupo,cod_razon,descripcion,notas,resultado) values(" _
               & gCntX_Parametros.CodigoConta & ",'IG','IG.7','Rotación de Activo Fijo','" _
               & "Eficiencia en la utilizacion de los Act.Fijos,Ingresos para las Ventas." _
               & " f(x) = Ventas Netas / Activo Fijo Neto','')"
        Call ConectionExecute(strSQL, 0)
        
        strSQL = "insert CntX_Razones(cod_contabilidad,cod_grupo,cod_razon,descripcion,notas,resultado) values(" _
               & gCntX_Parametros.CodigoConta & ",'IG','IG.8','Rotación de Activo a Largo Plazo','" _
               & "Engloba todas las inversiones a largo plazo que no necesariamente se traducen en ventas." _
               & " f(x) = Ventas Netas / Activo a Largo Plazo','')"
        Call ConectionExecute(strSQL, 0)
        
        
        strSQL = "insert CntX_Razones(cod_contabilidad,cod_grupo,cod_razon,descripcion,notas,resultado) values(" _
               & gCntX_Parametros.CodigoConta & ",'IG','IG.9','Rotación de Activo Total','" _
               & "Eficiencia para utilizar el total de Activos para Generacion de Ventas." _
               & " f(x) = Ventas Netas/ Activo Total Promedio','')"
        Call ConectionExecute(strSQL, 0)


strSQL = "insert CntX_Razones_Tipos(cod_contabilidad,cod_grupo,descripcion,activa) values(" _
       & gCntX_Parametros.CodigoConta & ",'IR','Indices de Rentabilidad',1)"
Call ConectionExecute(strSQL, 0)

    strSQL = "insert CntX_Razones(cod_contabilidad,cod_grupo,cod_razon,descripcion,notas,resultado) values(" _
           & gCntX_Parametros.CodigoConta & ",'IR','IR.1','Margen de Utilidad Bruta MUB','" _
           & "Ideal margen bruto de utilidades alto, y costo relativo de mercancias vendida bajo." _
           & " f(x) = Utilidad Bruta / Ventas Netas','')"
    Call ConectionExecute(strSQL, 0)
    
    strSQL = "insert CntX_Razones(cod_contabilidad,cod_grupo,cod_razon,descripcion,notas,resultado) values(" _
           & gCntX_Parametros.CodigoConta & ",'IR','IR.2','Margen de Utilidad de Operación MUO','" _
           & "Utilidades Puras (Ganadas por Cada cod_unidad monetaria)." _
           & " f(x) = Utilidades de Operacion / Ventas Netas','')"
    Call ConectionExecute(strSQL, 0)
    
    strSQL = "insert CntX_Razones(cod_contabilidad,cod_grupo,cod_razon,descripcion,notas,resultado) values(" _
           & gCntX_Parametros.CodigoConta & ",'IR','IR.3','Margen de Utilidad Neta MUN','" _
           & "Porcentaje Restantes sobre cada unidad monetaria de ventas." _
           & " f(x) = Utilidad Neta / ventas Netas','')"
    Call ConectionExecute(strSQL, 0)
    
    strSQL = "insert CntX_Razones(cod_contabilidad,cod_grupo,cod_razon,descripcion,notas,resultado) values(" _
           & gCntX_Parametros.CodigoConta & ",'IR','IR.4','Rendimiento de Operación sobre Activos (RAT)','" _
           & "(Rend.Inversion) Efectividad totale de la Admin. para generar utilidades con activos disponibles." _
           & " f(x) = Utilidades Netas después de impuestos / activos totales','')"
    Call ConectionExecute(strSQL, 0)
    
    strSQL = "insert CntX_Razones(cod_contabilidad,cod_grupo,cod_razon,descripcion,notas,resultado) values(" _
           & gCntX_Parametros.CodigoConta & ",'IR','IR.5','Rendimiento de Operación Sobre Activos (ROA))','" _
           & "Mide la rentabilidad Generada por los activos sobre las operaciones normales de la empresa." _
           & " f(x) = Utilidad de Operacion / Activo Total','')"
    Call ConectionExecute(strSQL, 0)
    
    
    strSQL = "insert CntX_Razones(cod_contabilidad,cod_grupo,cod_razon,descripcion,notas,resultado) values(" _
           & gCntX_Parametros.CodigoConta & ",'IR','IR.6','Rendimiento sobre la inversion (RSI)','" _
           & "RSI Indicador clave de la eficiencia y eficacia en la utilizacion de los recursos totales para generar gananacias netas." _
           & " f(x) = Utilidad Neta / Activo Total','')"
    Call ConectionExecute(strSQL, 0)
    
    
    strSQL = "insert CntX_Razones(cod_contabilidad,cod_grupo,cod_razon,descripcion,notas,resultado) values(" _
           & gCntX_Parametros.CodigoConta & ",'IR','IR.7','Rendimiento sobre el Patrimonio (RSP)','" _
           & "RSP Rendimiento Final que obtienen los socios de su inversion en la empresa." _
           & " f(x) = Utilidad Neta / Patrimonio','')"
    Call ConectionExecute(strSQL, 0)

End Sub

Private Sub Form_Activate()
vModulo = 20
End Sub

Private Sub Form_Load()

vModulo = 20

Set imgBanner.Picture = frmContenedor.imgBanner_Mantenimiento.Picture

tcMain.Item(0).Selected = True

With lswRepFx.ColumnHeaders
    .Clear
    .Add , , "", 400
    .Add , , "Razón Id", 1250, vbCenter
    .Add , , "Descripción", 6500
End With


With lsw.ColumnHeaders
    .Clear
    .Add , , "[ C ]", 900
    .Add , , "Cuenta", 2150, vbCenter
    .Add , , "Operador", 1200, vbCenter
    .Add , , "Descripción", 4650
End With


vGrid.AppearanceStyle = fxGridStyle
vGridFx.AppearanceStyle = fxGridStyle
lblGenera.Caption = ""

vScroll = False
 FlatScrollBarX.Value = 0
vScroll = True

Call sbPredeterminados
Call sbCargaGrupos

Call Formularios(Me)
Call RefrescaTags(Me)

btnBarra.Item(0).Enabled = vGridFx.Enabled
btnBarra.Item(1).Enabled = vGridFx.Enabled
btnBarra.Item(2).Enabled = vGridFx.Enabled


End Sub

Private Function fxResRazonMnt(vRazon As String, lngAnio As Long, iMes As Integer, vID As Integer) As Currency
Dim strSQL As String, rs As New ADODB.Recordset

'Saca el Movimiento de Cuentas, para la formula


If cboUnidades.Text = "TODOS" Then
    strSQL = "select (M.saldo_inicial + M.total_debitos + M.total_creditos) as Monto" _
           & " from CntX_Razones_detalle R inner join vCntX_Mov_Cuentas_General M" _
           & " on R.cod_contabilidad = M.cod_contabilidad and R.cod_cuenta = M.cod_cuenta"
Else
    strSQL = "select (M.saldo_inicial + M.total_debitos + M.total_creditos) as Monto" _
           & " from CntX_Razones_detalle R inner join vCntX_Mov_Cuentas_Unidad M" _
           & " on R.cod_contabilidad = M.cod_contabilidad and R.cod_cuenta = M.cod_cuenta" _
           & " and M.cod_unidad = '" & cboUnidades.ItemData(cboUnidades.ListIndex) & "'"
End If



strSQL = strSQL & " where R.cod_razon = '" & vRazon & "' and M.Anio = " & lngAnio _
       & " and M.mes = " & iMes & " and R.idX = " & vID


Call OpenRecordSet(rs, strSQL, 0)
If rs.EOF And rs.BOF Then
 fxResRazonMnt = 0
Else
 fxResRazonMnt = rs!Monto
End If
rs.Close

End Function

Private Function fxResRazon(vRazon As String, lngAnio As Long, iMes As Integer) As Currency
Dim strSQL As String, rs As New ADODB.Recordset
Dim x As Integer, vFormula As String
Dim curMontoTmp As Currency, y As Integer
Dim vUltIndice As Currency


On Error GoTo vError

'Extrae la Formula
strSQL = "select formula from CntX_Razones where cod_contabilidad = " & gCntX_Parametros.CodigoConta _
       & " and cod_razon = '" & vRazon & "'"
Call OpenRecordSet(rs, strSQL, 0)
  vFormula = Trim(rs!Formula)
rs.Close


'Formula Simple
strSQL = ""
For x = 1 To Len(vFormula)
  Select Case Mid(vFormula, x, 1)
    Case "C"
     y = 1
     Do While Not IsNumeric(Mid(vFormula, x + y, 1))
       y = y + 1
     Loop
     If y > 1 Then
         vUltIndice = Mid(vFormula, x + 1, y - 1)
     Else
         vUltIndice = Mid(vFormula, x + 1, y)
     End If
     curMontoTmp = fxResRazonMnt(vRazon, lngAnio, iMes, CInt(vUltIndice))
     strSQL = strSQL & curMontoTmp
    
    Case "M"
     y = 1
     Do While IsNumeric(Mid(vFormula, x + y, 1))
       y = y + 1
     Loop
     If y > 1 Then
        curMontoTmp = Mid(vFormula, x + 1, y - 1)
     Else
        curMontoTmp = Mid(vFormula, x + 1, y)
     End If
     strSQL = strSQL & curMontoTmp
    
    Case "(", ")", "-", "+", "*", "/"
     strSQL = strSQL & Mid(vFormula, x, 1)
  
  End Select
  
Next x
  
strSQL = "select " & strSQL & " as Valor"
Call OpenRecordSet(rs, strSQL, 0)
    fxResRazon = rs!Valor
rs.Close

Exit Function
  
vError:
  fxResRazon = 0
  
End Function

Private Sub sbCargaGrupos()
Dim strSQL As String

strSQL = "select cod_grupo,descripcion,activa from CntX_razones_tipos" _
       & " where cod_contabilidad = " & gCntX_Parametros.CodigoConta _
       & " order by cod_grupo"
Call sbCargaGrid(vGrid, 3, strSQL)
End Sub


Private Sub sbCargaCboGrixFx(vCol As Integer, vRow As Long, xGrid As Object)
Dim strResultado As String, rs As New ADODB.Recordset, strSQL As String

xGrid.Col = vCol
xGrid.Row = vRow
xGrid.CellType = CellTypeComboBox

strSQL = "select (cod_grupo + ' - ' + descripcion) as Descripcion from CntX_Razones_Tipos where cod_contabilidad = " _
       & gCntX_Parametros.CodigoConta & " order by cod_grupo"
Call OpenRecordSet(rs, strSQL, 0)

If Not rs.EOF And strUltimaSeleccion = "" Then strUltimaSeleccion = rs!Descripcion

strResultado = ""

Do While Not rs.EOF
  If Len(strResultado) = 0 Then
    strResultado = Chr$(9) & rs!Descripcion
  Else
    strResultado = strResultado & Chr$(9) & rs!Descripcion
  End If
  rs.MoveNext
Loop
rs.Close

xGrid.TypeComboBoxList = strResultado
xGrid.TypeComboBoxEditable = False
xGrid.Text = strUltimaSeleccion

End Sub


Private Function fxGuardarDefinicion() As Long
Dim strSQL As String, rs As New ADODB.Recordset
Dim vCuenta As String, vCuentaMadre As String, vAceptaMovimientos As String
Dim vPresupuesto As String

On Error GoTo vError


vGridFx.Col = 4
strUltimaSeleccion = vGridFx.Text

fxGuardarDefinicion = 0

vGridFx.Row = vGridFx.ActiveRow

vGridFx.Col = 1
rs.Open "select isnull(count(*),0) as Total from CntX_Razones where cod_razon = '" _
        & vGridFx.Text & "' and cod_contabilidad = " & gCntX_Parametros.CodigoConta, glogon.Conection, adOpenStatic

If rs!Total = 0 Then 'Insertar
  strSQL = "insert into CntX_Razones(cod_razon,descripcion,cod_contabilidad,resultado,cod_grupo,notas) values('"
  vGridFx.Col = 1
  strSQL = strSQL & vGridFx.Text & "','"
  vGridFx.Col = 2
  strSQL = strSQL & vGridFx.Text & "'," & gCntX_Parametros.CodigoConta & ",'"
  vGridFx.Col = 3
  strSQL = strSQL & vGridFx.Text & "','"
  vGridFx.Col = 4
  strSQL = strSQL & SIFGlobal.fxCodText(vGridFx.Text) & "','')"
  Call ConectionExecute(strSQL, 0)

  vGridFx.Col = 1
  
  Call Bitacora("Registra", "Razon Financiera Id: " & vGridFx.Text)
  
  fxGuardarDefinicion = 1

Else 'Actualizar

    vGridFx.Col = 2
    strSQL = "update CntX_Razones set descripcion = '" & vGridFx.Text & "', resultado = '"
    vGridFx.Col = 3
    strSQL = strSQL & vGridFx.Text & "',cod_grupo = '"
    vGridFx.Col = 4
    strSQL = strSQL & SIFGlobal.fxCodText(vGridFx.Text) & "'"
    vGridFx.Col = 1
    strSQL = strSQL & " where cod_contabilidad = " & gCntX_Parametros.CodigoConta _
           & " and cod_razon = '" & vGridFx.Text & "'"
    Call ConectionExecute(strSQL, 0)
  
    Call Bitacora("Modifica", "Razon Financiera Id: " & vGridFx.Text)
 
    fxGuardarDefinicion = 1
End If

rs.Close

Exit Function

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Function



Private Sub imgGuardaFormula_Click()
Dim strSQL As String

strSQL = "update CntX_Razones set formula = '" & txtFormula _
       & "' where cod_razon = '" & txtCodigo & "' and cod_contabilidad = " _
       & gCntX_Parametros.CodigoConta
Call ConectionExecute(strSQL, 0)

MsgBox "La formula Basica, fue reemplazada por la actual...", vbInformation


End Sub


Private Sub lsw_ItemClick(ByVal Item As XtremeSuiteControls.ListViewItem)

If lsw.ListItems.Count > 0 Then
 txtCuentaCod.Tag = Item.Text
 txtCuentaCod.Text = Item.SubItems(1)
 txtCuentaDesc.Text = Item.SubItems(3)
 cboOperador.Text = Item.SubItems(2)
End If

End Sub



Private Sub tcMain_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
Select Case Item.Index
  Case 0 'Grupos
    Call sbCargaGrupos
  Case 1 'Definicion
    Call sbCargaDefiniciones
  Case 2 'Formulas
    Call sbCargaFormulas
  Case 3 'Reportes
    Call sbCargaReportes
End Select
End Sub

Private Sub txtAnio_Change()
Call sbRefrescaInformacion
End Sub

Private Sub txtCodigo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtNombre.SetFocus
If KeyCode = vbKeyF4 Then
   gBusquedas.Resultado = ""
   gBusquedas.Resultado2 = ""
   gBusquedas.Columna = "cod_razon"
   gBusquedas.Orden = "cod_razon"
   gBusquedas.Consulta = "select cod_razon,descripcion from CntX_Razones"
   gBusquedas.Filtro = " and cod_contabilidad = " & gCntX_Parametros.CodigoConta _
                     & " and cod_grupo ='" & cboGrupo.ItemData(cboGrupo.ListIndex) & "'"
   frmBusquedas.Show vbModal
   txtCodigo = gBusquedas.Resultado
   txtNombre = gBusquedas.Resultado2
   Call sbFxCarga(cboGrupo.ItemData(cboGrupo.ListIndex), txtCodigo)
End If
End Sub

Private Sub txtCuentaCod_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtCuentaDesc.SetFocus
If KeyCode = vbKeyF4 Then
   frmCntX_ConsultaCuentas.Show vbModal
   txtCuentaCod = gCuenta
   txtCuentaDesc = fxCntX_Cuenta("D", fxCntX_CuentaFormato(False, gCuenta))
End If
End Sub

Private Sub txtCuentaCod_LostFocus()
txtCuentaCod = fxCntX_CuentaFormato(True, txtCuentaCod)
txtCuentaDesc = fxCntX_Cuenta("D", fxCntX_CuentaFormato(False, txtCuentaCod))
End Sub

Private Sub txtCuentaDesc_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then cboOperador.SetFocus
If KeyCode = vbKeyF4 Then
   frmCntX_ConsultaCuentas.Show vbModal
   txtCuentaCod = gCuenta
   txtCuentaDesc = fxCntX_Cuenta("D", fxCntX_CuentaFormato(False, gCuenta))
End If
End Sub

Private Function fxTipoOperador(vTipo As String)

Select Case vTipo
  Case "S"
    fxTipoOperador = "Suma"
  Case "R"
    fxTipoOperador = "Resta"
  Case "M"
    fxTipoOperador = "Multiplica"
  Case "D"
    fxTipoOperador = "Divide"
  Case "B"
    fxTipoOperador = "BASE"
  Case "Suma"
    fxTipoOperador = "S"
  Case "Resta"
    fxTipoOperador = "R"
  Case "Multiplica"
    fxTipoOperador = "M"
  Case "Divide"
    fxTipoOperador = "D"
  Case "BASE"
    fxTipoOperador = "B"
  Case Else
   fxTipoOperador = vTipo
End Select
End Function


Private Sub sbFxCarga(vGrupo As String, vRazon As String)
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem

Me.MousePointer = vbHourglass

On Error GoTo vError

tcFx.Item(0).Selected = True

strSQL = "select notas,formula from CntX_Razones where cod_grupo = '" & vGrupo & "' and cod_razon = '" & vRazon & "'"
Call OpenRecordSet(rs, strSQL, 0)
    txtFormula = rs!Formula & ""
    txtNotas = rs!Notas & ""
rs.Close

txtCuentaCod.Tag = ""
txtCuentaDesc = ""
txtCuentaCod = ""
cboOperador.Text = "BASE"


strSQL = "select R.*,C.descripcion,C.cod_Cuenta_Mask" _
       & " from CntX_Razones_detalle R inner join CntX_Cuentas C on R.cod_contabilidad = C.cod_contabilidad" _
       & " and R.cod_cuenta = C.cod_cuenta" _
       & " where R.cod_razon = '" & vRazon & "'" _
       & " and R.cod_contabilidad = " & gCntX_Parametros.CodigoConta _
       & " order by R.idx"
lsw.ListItems.Clear

Call OpenRecordSet(rs, strSQL)

Do While Not rs.EOF
 Set itmX = lsw.ListItems.Add(, , rs!idX)
     itmX.SubItems(1) = rs!Cod_Cuenta_Mask & ""
     itmX.SubItems(2) = fxTipoOperador(Trim(rs!operador))
     itmX.SubItems(3) = rs!Descripcion
 rs.MoveNext
Loop
rs.Close

Me.MousePointer = vbDefault

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub


Private Sub sbRefrescaInformacion()

On Error GoTo vError

txtAnio.Text = Val(txtAnio)

txtPeriodo.Text = fxCntX_PeriodoDesc(txtAnio, txtMes)

Exit Sub

vError:
End Sub

Private Sub txtMes_Change()
Call sbRefrescaInformacion
End Sub

Private Sub txtNombre_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtFormula.SetFocus
If KeyCode = vbKeyF4 Then
   gBusquedas.Resultado = ""
   gBusquedas.Resultado2 = ""
   gBusquedas.Columna = "descripcion"
   gBusquedas.Orden = "descripcion"
   gBusquedas.Consulta = "select cod_razon,descripcion from CntX_Razones"
   gBusquedas.Filtro = " and cod_contabilidad = " & gCntX_Parametros.CodigoConta _
                     & " and cod_grupo ='" & cboGrupo.ItemData(cboGrupo.ListIndex) & "'"
   frmBusquedas.Show vbModal
   txtCodigo = gBusquedas.Resultado
   txtNombre = gBusquedas.Resultado2
   Call sbFxCarga(cboGrupo.ItemData(cboGrupo.ListIndex), txtCodigo)
End If

End Sub

Private Sub vGridfx_KeyDown(KeyCode As Integer, Shift As Integer)
Dim i As Long

If vGridFx.ActiveCol = vGridFx.MaxCols And (KeyCode = vbKeyReturn Or KeyCode = vbKeyTab) Then
  i = fxGuardarDefinicion
  vGridFx.Row = vGridFx.ActiveRow
  vGridFx.Col = 1
  If vGridFx.MaxRows <= vGridFx.ActiveRow Then
    vGridFx.MaxRows = vGridFx.MaxRows + 1
    vGridFx.Row = vGridFx.MaxRows
    Call sbCargaCboGrixFx(4, vGridFx.MaxRows, vGridFx)
  End If
End If

If KeyCode = vbKeyInsert Then
    vGridFx.MaxRows = vGridFx.MaxRows + 1
    vGridFx.InsertRows vGridFx.ActiveRow, 1
    Call sbCargaCboGrixFx(4, vGridFx.ActiveRow, vGridFx)
End If


End Sub


Private Sub sbCargaGridLocal(xGrid As Object, xGridMaxCol As Integer, strSQL As String)
Dim rs As New ADODB.Recordset, i As Integer, strResultado As String

Me.MousePointer = vbHourglass

xGrid.MaxCols = xGridMaxCol
xGrid.MaxRows = 1
'Call sbCargaComboTiposCuenta(3, xGrid.MaxRows, xGrid)


rs.Open "select (cod_grupo + ' - ' + descripcion) as Descripcion" _
    & " from CntX_Razones_Tipos where cod_contabilidad = " _
    & gCntX_Parametros.CodigoConta & " order by cod_grupo", glogon.Conection, adOpenStatic
    If Not rs.EOF And strUltimaSeleccion = "" Then strUltimaSeleccion = rs!Descripcion
    strResultado = ""
    Do While Not rs.EOF
        If Len(strResultado) = 0 Then
          strResultado = Chr$(9) & rs!Descripcion
        Else
          strResultado = strResultado & Chr$(9) & rs!Descripcion
        End If
      rs.MoveNext
    Loop
rs.Close

xGrid.Row = xGrid.MaxRows

Call OpenRecordSet(rs, strSQL)

Do While Not rs.EOF
  xGrid.Row = xGrid.MaxRows
  
  xGrid.Col = 4
  xGrid.CellType = CellTypeComboBox
  xGrid.TypeComboBoxList = strResultado
  xGrid.TypeComboBoxEditable = False
  xGrid.Text = strUltimaSeleccion
  
  For i = 1 To xGrid.MaxCols
    xGrid.Col = i
    xGrid.Text = CStr(rs.Fields(i - 1).Value)
  Next i
  xGrid.MaxRows = xGrid.MaxRows + 1
  rs.MoveNext
Loop
rs.Close

xGrid.Row = xGrid.MaxRows

xGrid.Col = 4
xGrid.CellType = CellTypeComboBox
xGrid.TypeComboBoxList = strResultado
xGrid.TypeComboBoxEditable = False
xGrid.Text = strUltimaSeleccion


Me.MousePointer = vbDefault

End Sub


Private Sub sbCargaDefiniciones()
Dim strSQL As String

strSQL = "select R.cod_razon,R.descripcion,resultado,(T.cod_grupo + ' - ' + T.descripcion) as Grupo" _
       & " from CntX_Razones_Tipos T inner join CntX_Razones R on T.cod_contabilidad = R.cod_contabilidad" _
       & " and T.cod_grupo = R.cod_grupo" _
       & " where R.cod_contabilidad = " & gCntX_Parametros.CodigoConta _
       & " order by T.cod_grupo,R.cod_razon"
Call sbCargaGridLocal(vGridFx, 4, strSQL)

End Sub

Private Sub sbCargaGrpTipos(pCbo As Object)
Dim strSQL As String

strSQL = "select rtrim(cod_grupo) as  'IdX', rtrim(descripcion) as 'itmX'" _
       & " from CntX_Razones_Tipos" _
       & " where cod_contabilidad = " & gCntX_Parametros.CodigoConta _
       & " order by cod_grupo"

Call sbCbo_Llena_New(pCbo, strSQL, False, True)

End Sub


Private Sub sbCargaFormulas()

Call sbCargaGrpTipos(cboGrupo)

With cboOperador
 .Clear
 .AddItem "BASE"
 .AddItem "Resta"
 .AddItem "Suma"
 .AddItem "Divide"
 .AddItem "Multiplica"
 .Text = "BASE"
End With

txtCodigo.Text = ""
txtNombre.Text = ""
txtFormula.Text = ""

txtCuentaCod.Text = ""
txtCuentaDesc.Text = ""

lsw.ListItems.Clear

End Sub


Private Sub sbCargaReportes()
Dim strSQL As String

strSQL = "select rtrim(cod_unidad) as 'IdX', rtrim(descripcion) as 'ItmX' from CntX_Unidades" _
        & " where cod_contabilidad = " & gCntX_Parametros.CodigoConta
Call sbCbo_Llena_New(cboUnidades, strSQL, True, True)
        
strSQL = "select rtrim(cod_grupo) as 'IdX', rtrim(descripcion) as 'ItmX' from CntX_Razones_Tipos" _
       & " where cod_contabilidad = " & gCntX_Parametros.CodigoConta & " order by cod_grupo"
Call sbCbo_Llena_New(cboRepGrupo, strSQL, True, True)
        
txtAnio.Text = gCntX_Parametros.PeriodoAnio
txtMes.Text = gCntX_Parametros.PeriodoMes

End Sub

Private Function fxGuardarGrupos() As Long
Dim strSQL As String, rs As New ADODB.Recordset
'Guarda la información de la linea
'si es Insert devuelve el codigo, sino devuelve 0

On Error GoTo vError

fxGuardarGrupos = 0
vGrid.Row = vGrid.ActiveRow
vGrid.Col = 1

rs.Open "select isnull(count(*),0) as Total from CntX_Razones_Tipos where cod_grupo = '" _
        & vGrid.Text & "' and cod_contabilidad = " & gCntX_Parametros.CodigoConta, glogon.Conection, adOpenStatic

If rs!Total = 0 Then 'Insertar
  strSQL = "insert into CntX_Razones_Tipos(cod_grupo,cod_contabilidad,descripcion) values('"
  vGrid.Col = 1
  strSQL = strSQL & UCase(vGrid.Text) & "'," & gCntX_Parametros.CodigoConta & ",'"
  vGrid.Col = 2
  strSQL = strSQL & UCase(vGrid.Text) & "')"
  Call ConectionExecute(strSQL, 0)

  vGrid.Col = 2
  
  Call Bitacora("Registra", "Razon Fin. Tipo / Grupo : " & vGrid.Text)
  
  fxGuardarGrupos = 1

Else 'Actualizar

 vGrid.Col = 2
 strSQL = "update CntX_Razones_Tipos set descripcion = '" & UCase(vGrid.Text) & "'"
 strSQL = strSQL & " where cod_contabilidad = " & gCntX_Parametros.CodigoConta _
        & " and cod_grupo = '"
 vGrid.Col = 1
 strSQL = strSQL & vGrid.Text & "'"
 Call ConectionExecute(strSQL, 0)
 
 fxGuardarGrupos = 1
 
 Call Bitacora("Modifica", "Razon Fin. Tipo / Grupo : " & vGrid.Text)

End If

rs.Close

Exit Function

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Function


Private Sub vGrid_KeyDown(KeyCode As Integer, Shift As Integer)
Dim i As Integer

If vGrid.ActiveCol = vGrid.MaxCols And (KeyCode = vbKeyReturn Or KeyCode = vbKeyTab) Then
  i = fxGuardarGrupos
  vGrid.Row = vGrid.ActiveRow
  vGrid.Col = 1
'  vGrid.Text = i
  If vGrid.MaxRows <= vGrid.ActiveRow Then
    vGrid.MaxRows = vGrid.MaxRows + 1
    vGrid.Row = vGrid.MaxRows
  End If
End If

End Sub

