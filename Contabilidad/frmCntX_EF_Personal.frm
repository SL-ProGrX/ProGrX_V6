VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.controls.v22.1.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.shortcutbar.v22.1.0.ocx"
Begin VB.Form frmCntX_EF_Personal 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Estados Financieros Personalizados"
   ClientHeight    =   9075
   ClientLeft      =   30
   ClientTop       =   375
   ClientWidth     =   14265
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   9075
   ScaleWidth      =   14265
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer TimerX 
      Interval        =   10
      Left            =   360
      Top             =   120
   End
   Begin XtremeSuiteControls.ProgressBar prgBar 
      Height          =   135
      Left            =   120
      TabIndex        =   32
      Top             =   8880
      Width           =   14055
      _Version        =   1441793
      _ExtentX        =   24791
      _ExtentY        =   238
      _StockProps     =   93
   End
   Begin XtremeSuiteControls.TabControl tcMain 
      Height          =   7815
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   14055
      _Version        =   1441793
      _ExtentX        =   24786
      _ExtentY        =   13779
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
      ItemCount       =   3
      Item(0).Caption =   "Estados Financieros"
      Item(0).ControlCount=   3
      Item(0).Control(0)=   "vGrid"
      Item(0).Control(1)=   "gItems"
      Item(0).Control(2)=   "scEF"
      Item(1).Caption =   "Items y Cuentas"
      Item(1).ControlCount=   6
      Item(1).Control(0)=   "tcCtas"
      Item(1).Control(1)=   "Label2(0)"
      Item(1).Control(2)=   "cboEF"
      Item(1).Control(3)=   "Label2(1)"
      Item(1).Control(4)=   "lswItems"
      Item(1).Control(5)=   "scItem"
      Item(2).Caption =   "Informes"
      Item(2).ControlCount=   14
      Item(2).Control(0)=   "cboEFRep"
      Item(2).Control(1)=   "Label2(3)"
      Item(2).Control(2)=   "Label2(4)"
      Item(2).Control(3)=   "Label2(5)"
      Item(2).Control(4)=   "txtAnio"
      Item(2).Control(5)=   "Label2(6)"
      Item(2).Control(6)=   "cboRepExpresado"
      Item(2).Control(7)=   "btnInforme"
      Item(2).Control(8)=   "txtMes"
      Item(2).Control(9)=   "lblStatus"
      Item(2).Control(10)=   "cboComparativo"
      Item(2).Control(11)=   "Label2(7)"
      Item(2).Control(12)=   "lsw"
      Item(2).Control(13)=   "Label2(8)"
      Begin XtremeSuiteControls.ListView lswItems 
         Height          =   6252
         Left            =   -69880
         TabIndex        =   9
         Top             =   1440
         Visible         =   0   'False
         Width           =   4692
         _Version        =   1441793
         _ExtentX        =   8276
         _ExtentY        =   11028
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
      Begin XtremeSuiteControls.ListView lsw 
         Height          =   1695
         Left            =   -64360
         TabIndex        =   34
         Top             =   3000
         Visible         =   0   'False
         Width           =   5055
         _Version        =   1441793
         _ExtentX        =   8916
         _ExtentY        =   2990
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
         HideColumnHeaders=   -1  'True
         FullRowSelect   =   -1  'True
         Appearance      =   17
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.TabControl tcCtas 
         Height          =   7452
         Left            =   -65080
         TabIndex        =   3
         Top             =   720
         Visible         =   0   'False
         Width           =   9132
         _Version        =   1441793
         _ExtentX        =   16108
         _ExtentY        =   13144
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
         ItemCount       =   3
         Item(0).Caption =   "Asignar Cuentas"
         Item(0).ControlCount=   6
         Item(0).Control(0)=   "lswCtas"
         Item(0).Control(1)=   "Label2(2)"
         Item(0).Control(2)=   "txtCtaInicial"
         Item(0).Control(3)=   "txtCtaFinal"
         Item(0).Control(4)=   "btnCtas"
         Item(0).Control(5)=   "chkCtaTodas(0)"
         Item(1).Caption =   "Cuentas Registradas"
         Item(1).ControlCount=   3
         Item(1).Control(0)=   "lswCtaAsg"
         Item(1).Control(1)=   "chkCtaTodas(1)"
         Item(1).Control(2)=   "btnCtaElimina"
         Item(2).Caption =   "Funciones"
         Item(2).ControlCount=   1
         Item(2).Control(0)=   "lswFx"
         Begin XtremeSuiteControls.ListView lswCtas 
            Height          =   6252
            Left            =   0
            TabIndex        =   4
            Top             =   720
            Width           =   9012
            _Version        =   1441793
            _ExtentX        =   15896
            _ExtentY        =   11028
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
         Begin XtremeSuiteControls.ListView lswCtaAsg 
            Height          =   6252
            Left            =   -70000
            TabIndex        =   14
            Top             =   720
            Visible         =   0   'False
            Width           =   9012
            _Version        =   1441793
            _ExtentX        =   15896
            _ExtentY        =   11028
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
         Begin XtremeSuiteControls.ListView lswFx 
            Height          =   6615
            Left            =   -70000
            TabIndex        =   33
            Top             =   360
            Visible         =   0   'False
            Width           =   9135
            _Version        =   1441793
            _ExtentX        =   16113
            _ExtentY        =   11668
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
         Begin XtremeSuiteControls.CheckBox chkCtaTodas 
            Height          =   210
            Index           =   0
            Left            =   120
            TabIndex        =   17
            Top             =   360
            Width           =   210
            _Version        =   1441793
            _ExtentX        =   370
            _ExtentY        =   370
            _StockProps     =   79
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.PushButton btnCtas 
            Height          =   312
            Left            =   7800
            TabIndex        =   13
            Top             =   360
            Width           =   1092
            _Version        =   1441793
            _ExtentX        =   1926
            _ExtentY        =   556
            _StockProps     =   79
            Caption         =   "Consultar"
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
         Begin XtremeSuiteControls.FlatEdit txtCtaInicial 
            Height          =   312
            Left            =   2880
            TabIndex        =   11
            Top             =   360
            Width           =   2412
            _Version        =   1441793
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
            Alignment       =   2
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtCtaFinal 
            Height          =   312
            Left            =   5280
            TabIndex        =   12
            Top             =   360
            Width           =   2412
            _Version        =   1441793
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
            Alignment       =   2
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.CheckBox chkCtaTodas 
            Height          =   216
            Index           =   1
            Left            =   -70000
            TabIndex        =   18
            Top             =   360
            Visible         =   0   'False
            Width           =   216
            _Version        =   1441793
            _ExtentX        =   370
            _ExtentY        =   370
            _StockProps     =   79
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.PushButton btnCtaElimina 
            Height          =   312
            Left            =   -64120
            TabIndex        =   19
            Top             =   360
            Visible         =   0   'False
            Width           =   3132
            _Version        =   1441793
            _ExtentX        =   5524
            _ExtentY        =   550
            _StockProps     =   79
            Caption         =   "Eliminar las Seleccionadas"
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
         Begin XtremeSuiteControls.Label Label2 
            Height          =   192
            Index           =   2
            Left            =   840
            TabIndex        =   10
            Top             =   360
            Width           =   1812
            _Version        =   1441793
            _ExtentX        =   3196
            _ExtentY        =   339
            _StockProps     =   79
            Caption         =   "Rango de Cuentas:"
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
            UseMnemonic     =   0   'False
            Transparent     =   -1  'True
            WordWrap        =   -1  'True
         End
      End
      Begin FPSpreadADO.fpSpread vGrid 
         Height          =   3012
         Left            =   840
         TabIndex        =   2
         Top             =   480
         Width           =   12012
         _Version        =   524288
         _ExtentX        =   21188
         _ExtentY        =   5313
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
         MaxCols         =   4
         ScrollBars      =   2
         SpreadDesigner  =   "frmCntX_EF_Personal.frx":0000
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin XtremeSuiteControls.ComboBox cboEF 
         Height          =   312
         Left            =   -69880
         TabIndex        =   6
         Top             =   720
         Visible         =   0   'False
         Width           =   4692
         _Version        =   1441793
         _ExtentX        =   8281
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
      Begin FPSpreadADO.fpSpread gItems 
         Height          =   3735
         Left            =   840
         TabIndex        =   7
         Top             =   4080
         Width           =   12015
         _Version        =   524288
         _ExtentX        =   21193
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
         MaxCols         =   6
         ScrollBars      =   2
         SpreadDesigner  =   "frmCntX_EF_Personal.frx":0683
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin XtremeSuiteControls.ComboBox cboEFRep 
         Height          =   330
         Left            =   -64360
         TabIndex        =   20
         Top             =   840
         Visible         =   0   'False
         Width           =   5055
         _Version        =   1441793
         _ExtentX        =   8916
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
      Begin XtremeSuiteControls.ComboBox cboRepExpresado 
         Height          =   312
         Left            =   -64360
         TabIndex        =   23
         Top             =   1440
         Visible         =   0   'False
         Width           =   1692
         _Version        =   1441793
         _ExtentX        =   2990
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
      Begin XtremeSuiteControls.FlatEdit txtAnio 
         Height          =   312
         Left            =   -64360
         TabIndex        =   25
         Top             =   2040
         Visible         =   0   'False
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtMes 
         Height          =   312
         Left            =   -63520
         TabIndex        =   26
         Top             =   2040
         Visible         =   0   'False
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.PushButton btnInforme 
         Height          =   555
         Left            =   -64360
         TabIndex        =   28
         Top             =   4800
         Visible         =   0   'False
         Width           =   1695
         _Version        =   1441793
         _ExtentX        =   2984
         _ExtentY        =   974
         _StockProps     =   79
         Caption         =   "Procesar Informe"
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
      Begin XtremeSuiteControls.ComboBox cboComparativo 
         Height          =   312
         Left            =   -64360
         TabIndex        =   30
         Top             =   2520
         Visible         =   0   'False
         Width           =   1692
         _Version        =   1441793
         _ExtentX        =   2990
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
      Begin XtremeSuiteControls.Label Label2 
         Height          =   315
         Index           =   8
         Left            =   -66040
         TabIndex        =   35
         Top             =   3000
         Visible         =   0   'False
         Width           =   1815
         _Version        =   1441793
         _ExtentX        =   3196
         _ExtentY        =   550
         _StockProps     =   79
         Caption         =   "Informes:"
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
         UseMnemonic     =   0   'False
         Transparent     =   -1  'True
         WordWrap        =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label2 
         Height          =   312
         Index           =   7
         Left            =   -66040
         TabIndex        =   31
         Top             =   2520
         Visible         =   0   'False
         Width           =   1812
         _Version        =   1441793
         _ExtentX        =   3196
         _ExtentY        =   550
         _StockProps     =   79
         Caption         =   "Compartivo:"
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
         UseMnemonic     =   0   'False
         Transparent     =   -1  'True
         WordWrap        =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblStatus 
         Height          =   495
         Left            =   -62440
         TabIndex        =   29
         Top             =   4800
         Visible         =   0   'False
         Width           =   3135
         _Version        =   1441793
         _ExtentX        =   5530
         _ExtentY        =   873
         _StockProps     =   79
         Caption         =   "..."
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
      Begin XtremeSuiteControls.Label Label2 
         Height          =   312
         Index           =   6
         Left            =   -62560
         TabIndex        =   27
         Top             =   2040
         Visible         =   0   'False
         Width           =   1812
         _Version        =   1441793
         _ExtentX        =   3196
         _ExtentY        =   550
         _StockProps     =   79
         Caption         =   "AAAA/MM:"
         ForeColor       =   8421504
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
         UseMnemonic     =   0   'False
         Transparent     =   -1  'True
         WordWrap        =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label2 
         Height          =   312
         Index           =   5
         Left            =   -66040
         TabIndex        =   24
         Top             =   2040
         Visible         =   0   'False
         Width           =   1812
         _Version        =   1441793
         _ExtentX        =   3196
         _ExtentY        =   550
         _StockProps     =   79
         Caption         =   "Periodo:"
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
         UseMnemonic     =   0   'False
         Transparent     =   -1  'True
         WordWrap        =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label2 
         Height          =   312
         Index           =   4
         Left            =   -66040
         TabIndex        =   22
         Top             =   1440
         Visible         =   0   'False
         Width           =   1812
         _Version        =   1441793
         _ExtentX        =   3196
         _ExtentY        =   550
         _StockProps     =   79
         Caption         =   "Expresado en:"
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
         UseMnemonic     =   0   'False
         Transparent     =   -1  'True
         WordWrap        =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label2 
         Height          =   312
         Index           =   3
         Left            =   -66040
         TabIndex        =   21
         Top             =   840
         Visible         =   0   'False
         Width           =   1812
         _Version        =   1441793
         _ExtentX        =   3196
         _ExtentY        =   550
         _StockProps     =   79
         Caption         =   "Estado Financiero:"
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
         UseMnemonic     =   0   'False
         Transparent     =   -1  'True
         WordWrap        =   -1  'True
      End
      Begin XtremeShortcutBar.ShortcutCaption scItem 
         Height          =   372
         Left            =   -65080
         TabIndex        =   16
         Top             =   360
         Visible         =   0   'False
         Width           =   9132
         _Version        =   1441793
         _ExtentX        =   16108
         _ExtentY        =   656
         _StockProps     =   14
         Caption         =   "Seleccione un Item del Estado Financiero"
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
      Begin XtremeShortcutBar.ShortcutCaption scEF 
         Height          =   372
         Left            =   840
         TabIndex        =   15
         Top             =   3600
         Width           =   12132
         _Version        =   1441793
         _ExtentX        =   21399
         _ExtentY        =   656
         _StockProps     =   14
         Caption         =   "Seleccione un Estado Financiero"
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
         Height          =   192
         Index           =   1
         Left            =   -69880
         TabIndex        =   8
         Top             =   1200
         Visible         =   0   'False
         Width           =   1812
         _Version        =   1441793
         _ExtentX        =   3196
         _ExtentY        =   339
         _StockProps     =   79
         Caption         =   "Items:"
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
         UseMnemonic     =   0   'False
         Transparent     =   -1  'True
         WordWrap        =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label2 
         Height          =   192
         Index           =   0
         Left            =   -69880
         TabIndex        =   5
         Top             =   480
         Visible         =   0   'False
         Width           =   1812
         _Version        =   1441793
         _ExtentX        =   3196
         _ExtentY        =   339
         _StockProps     =   79
         Caption         =   "Estado Financiero:"
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
         UseMnemonic     =   0   'False
         Transparent     =   -1  'True
         WordWrap        =   -1  'True
      End
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Estados Financieros Personalizados"
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
      Height          =   480
      Index           =   2
      Left            =   2160
      TabIndex        =   0
      Top             =   300
      Width           =   7332
   End
   Begin VB.Image imgBanner 
      Height          =   972
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   14532
   End
End
Attribute VB_Name = "frmCntX_EF_Personal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem
Dim vPaso As Boolean

Private Sub btnCtaElimina_Click()
If vPaso Then Exit Sub


Dim i As Long

On Error GoTo vError

With lswCtaAsg.ListItems

Me.MousePointer = vbHourglass

strSQL = ""

For i = 1 To .Count
    If .Item(i).Checked Then
        strSQL = strSQL & Space(10) & "exec spCntX_EF_Cuentas " & gCntX_Parametros.CodigoConta & ", '" & cboEF.ItemData(cboEF.ListIndex) _
               & "', '" & scItem.Tag & "', '" & fxCntX_CuentaFormato(False, .Item(i).Text, 0) _
               & "','" & glogon.Usuario & "','E'"
    End If

    If Len(strSQL) > 20000 Then
        Call ConectionExecute(strSQL)
        strSQL = ""
    End If

Next i

End With

'Lote Final
If Len(strSQL) > 0 Then
    Call ConectionExecute(strSQL)
    strSQL = ""
End If

Me.MousePointer = vbDefault

'Actualiza la Lista
Call sbCuentas_Consulta_Asg

Exit Sub

vError:
    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub btnCtas_Click()
Call sbCuentas_Consulta
End Sub

Private Function fxTituloPeriodo(pMes As Integer, pAnio As Long) As String
Dim pResultado As String

Select Case pMes
    Case 1
        pResultado = "Ene"
    Case 2
        pResultado = "Feb"
    Case 3
        pResultado = "Mar"
    Case 4
        pResultado = "Abr"
    Case 5
        pResultado = "May"
    Case 6
        pResultado = "Jun"
    Case 7
        pResultado = "Jul"
    Case 8
        pResultado = "Ago"
    Case 9
        pResultado = "Set"
    Case 10
        pResultado = "Oct"
    Case 11
        pResultado = "Nov"
    Case 12
        pResultado = "Dic"
End Select

pResultado = UCase(pResultado) & " - " & pAnio

fxTituloPeriodo = pResultado

End Function

Private Sub btnInforme_Click()

On Error GoTo vError

Dim vInforme As Boolean, i As Integer

vInforme = False

For i = 1 To lsw.ListItems.Count
  If lsw.ListItems(i).Checked Then
    vInforme = True
  End If
Next i

If Not vInforme Then
  MsgBox "Seleccione al menos un informe!", vbExclamation
  Exit Sub
End If

Me.MousePointer = vbHourglass


lblStatus.Caption = "Procesando... [Espere]"
DoEvents


Dim pTSaldo As String, pTSaldo_A1 As String, pTSaldo_A2 As String
Dim pTMes As String, pTMes_A1 As String, pTMes_A2 As String

Dim pAnio As Long, pMes As Integer
Dim pExpresado As Long

Select Case cboRepExpresado.Text
    Case "unidad"
        pExpresado = 1
    Case "miles"
        pExpresado = 1000
    Case "millones"
        pExpresado = 1000000
End Select


'Procesa Resultados
strSQL = "exec spCntX_EF_Procesa " & gCntX_Parametros.CodigoConta & ",'" & cboEFRep.ItemData(cboEFRep.ListIndex) _
       & "'," & txtAnio.Text & "," & txtMes.Text & ",'" & glogon.Usuario & "','" & Mid(cboComparativo.Text, 1, 1) _
       & "', " & pExpresado
Call ConectionExecute(strSQL)

pAnio = txtAnio.Text
pMes = txtMes.Text

pTSaldo = fxTituloPeriodo(pMes, pAnio)
pTMes = pTSaldo & Space(10) & "MENSUAL"

If Mid(cboComparativo.Text, 1, 1) = "A" Then
  'Anual
    pAnio = pAnio - 1
    pTSaldo_A1 = fxTituloPeriodo(pMes, pAnio)
    pTMes_A1 = pTSaldo_A1 & Space(10) & "MENSUAL"
       
    pAnio = pAnio - 1
    pTSaldo_A2 = fxTituloPeriodo(pMes, pAnio)
    pTMes_A2 = pTSaldo_A2 & Space(10) & "MENSUAL"

Else
  'Trimestral
    If pMes = 1 Then
       pAnio = pAnio - 1
       pMes = 12
    Else
       pMes = pMes - 1
    End If
    
    pTSaldo_A1 = fxTituloPeriodo(pMes, pAnio)
    pTMes_A1 = pTSaldo_A1 & Space(10) & "MENSUAL"
       
    If pMes = 1 Then
       pAnio = pAnio - 1
       pMes = 12
    Else
       pMes = pMes - 1
    End If
    pTSaldo_A2 = fxTituloPeriodo(pMes, pAnio)
    pTMes_A2 = pTSaldo_A2 & Space(10) & "MENSUAL"
       
End If
 

lblStatus.Caption = "Preparando Informe..."
DoEvents

For i = 1 To lsw.ListItems.Count
  If lsw.ListItems(i).Checked Then
  
  
    With frmContenedor.Crt
     .Reset
     .WindowShowGroupTree = True
     .WindowShowPrintSetupBtn = True
     .WindowShowRefreshBtn = True
     .WindowShowSearchBtn = True
     .WindowState = crptMaximized
     .WindowTitle = "ProGrX: Contabilidad"
     .Formulas(0) = "Fecha='" & Format(fxFechaServidor, "dd/mm/yyyy") & "'"
     .Formulas(1) = "Empresa='" & gCntX_Parametros.NombreEmpresa & "'"
     .Formulas(2) = "Usuario='" & glogon.Usuario & "'"
     .Formulas(3) = "Titulo='" & cboEFRep.Text & "'"
     .Formulas(4) = "SubTitulo='Expresado en: " & cboRepExpresado.Text & "'"
     
     .Formulas(5) = "fxT_Saldo = '" & pTSaldo & "'"
     .Formulas(6) = "fxT_Saldo_A1 = '" & pTSaldo_A1 & "'"
     .Formulas(7) = "fxT_Saldo_A2 = '" & pTSaldo_A2 & "'"
     
     .Formulas(8) = "fxT_Mes = '" & pTMes & "'"
     .Formulas(9) = "fxT_Mes_A1 = '" & pTMes_A1 & "'"
     .Formulas(10) = "fxT_Mes_A2 = '" & pTMes_A2 & "'"
     
     .Connect = glogon.ConectRPT
      
           
     Select Case lsw.ListItems(i).Key
        Case "A01" 'Anual Saldos
            .ReportFileName = SIFGlobal.fxPathReportes("Contabilidad_EF_Personal_A01.rpt")
        Case "A02" 'Anual Saldos + Meses
            .ReportFileName = SIFGlobal.fxPathReportes("Contabilidad_EF_Personal_A02.rpt")
        Case "T01" 'Trimestral Saldos
            .ReportFileName = SIFGlobal.fxPathReportes("Contabilidad_EF_Personal_Trimestral.rpt")
        Case "T02" 'Trimestral Saldos + Meses
            .ReportFileName = SIFGlobal.fxPathReportes("Contabilidad_EF_Personal_Trimestral_T02.rpt")
     End Select
        
     .SelectionFormula = "{vCntX_EF_Resultados.USUARIO} = '" & glogon.Usuario _
                 & "' AND {vCntX_EF_Resultados.COD_CONTABILIDAD} = " & gCntX_Parametros.CodigoConta _
                 & "  AND {vCntX_EF_Resultados.COD_EF} = '" & cboEFRep.ItemData(cboEFRep.ListIndex) & "' AND {vCntX_EF_Resultados.Item_Id_MADRE} = ''"
        
               
     .Action = 1
    
    End With
  
  
  
  
  End If
Next i


Me.MousePointer = vbDefault
lblStatus.Caption = ""

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
 
End Sub

Private Sub cboComparativo_Click()

With lsw.ListItems
    .Clear
    
    If cboComparativo.Text = "Anual" Then
        .Add , "A01", "Informe Anual Saldos"
        .Add , "A02", "Informe Anual Saldos + Mensual"
    Else
        .Add , "T01", "Informe Trimestral Saldos"
        .Add , "T02", "Informe Trimestral Saldos + Mensual"
    End If

End With

End Sub

Private Sub cboEF_Click()
If vPaso Then Exit Sub
If cboEF.ListCount = 0 Then Exit Sub

lswItems.ListItems.Clear
lswCtas.ListItems.Clear
lswCtaAsg.ListItems.Clear

scItem.Caption = "Seleccione un Item para vincular cuentas"
scItem.Tag = ""

With lswItems.ColumnHeaders
    .Clear
    .Add , , "Descripción", lswItems.Width - 150
End With

strSQL = "select ITEM_ID, DESCRIPCION " _
       & " from CNTX_EF_SECCIONES Where ES_TITULO = 0" _
       & " and COD_EF = '" & cboEF.ItemData(cboEF.ListIndex) & "'" _
       & " and COD_CONTABILIDAD = " & gCntX_Parametros.CodigoConta _
       & " order by ITEM_ID_MADRE, PRIORIDAD, ITEM_ID"
Call OpenRecordSet(rs, strSQL)

vPaso = True

Do While Not rs.EOF
    Set itmX = lswItems.ListItems.Add(, , rs!Descripcion)
        itmX.Tag = rs!Item_Id
    rs.MoveNext
Loop
rs.Close

vPaso = False

End Sub

Private Sub chkCtaTodas_Click(Index As Integer)
Dim i As Long

Select Case Index
    Case 0 'Cta Nuevas
        With lswCtas.ListItems
                
        Me.MousePointer = vbHourglass
        
        For i = 1 To .Count
            If .Item(i).Checked <> chkCtaTodas.Item(Index).Value Then
                .Item(i).Checked = chkCtaTodas.Item(Index).Value
            End If
        Next i
                
        Me.MousePointer = vbDefault
        
        End With
    
    Case 1 'Ctas Asignadas
        With lswCtaAsg.ListItems
                
        Me.MousePointer = vbHourglass
        
        For i = 1 To .Count
                .Item(i).Checked = chkCtaTodas.Item(Index).Value
        Next i
                
        Me.MousePointer = vbDefault
        
        End With
End Select

End Sub

Private Sub Form_Activate()
vModulo = 20
End Sub

Private Sub sbConsulta()

tcMain.Item(0).Selected = True

scItem.Tag = ""
scItem.Caption = "Seleccione un Estado Financiero"

gItems.MaxRows = 0

vPaso = True

strSQL = "select COD_EF, DESCRIPCION,ACTIVO,0" _
      & " from CNTX_EF_PERSONAL" _
      & " Where cod_contabilidad = " & gCntX_Parametros.CodigoConta _
      & " order by COD_EF"

Call sbCargaGrid(vGrid, vGrid.MaxCols, strSQL)

vPaso = False

End Sub



Private Sub sbConsulta_Items_Mantenimiento()

tcMain.Item(0).Selected = True

strSQL = "select ITEM_ID, ITEM_ID_MADRE, PRIORIDAD, ES_TITULO, isnull(TOTALES,0), DESCRIPCION" _
      & " from CNTX_EF_SECCIONES" _
      & " Where cod_contabilidad = " & gCntX_Parametros.CodigoConta _
      & "  and cod_EF = '" & scEF.Tag & "'" _
      & " order by ITEM_ID, ITEM_ID_MADRE, PRIORIDAD"
Call sbCargaGrid(gItems, gItems.MaxCols, strSQL)


End Sub




Private Sub Form_Load()

vModulo = 20

vGrid.AppearanceStyle = fxGridStyle
Set imgBanner.Picture = frmContenedor.imgBanner_01.Picture

With lsw.ColumnHeaders
    .Clear
    .Add , , "Informe", lsw.Width - 150
End With

txtAnio.Text = gCntX_Parametros.PeriodoAnio
txtMes.Text = gCntX_Parametros.PeriodoMes

cboRepExpresado.Clear
cboRepExpresado.AddItem "unidad"
cboRepExpresado.AddItem "miles"
cboRepExpresado.AddItem "millones"


cboComparativo.Clear
cboComparativo.AddItem "Anual"
cboComparativo.AddItem "Trimestral"
cboComparativo.Text = "Anual"

Call cboComparativo_Click


Call sbConsulta

Call Formularios(Me)
Call RefrescaTags(Me)

End Sub


Private Function fxGuardar_gItems() As Long

On Error GoTo vError

With gItems

fxGuardar_gItems = 0

.Row = .ActiveRow
.Col = 1

strSQL = "select isnull(count(*),0) as Existe from CNTX_EF_SECCIONES " _
       & " where COD_EF = '" & scEF.Tag & "' AND COD_CONTABILIDAD = " & gCntX_Parametros.CodigoConta _
       & " and ITEM_ID = '" & .Text & "'"
Call OpenRecordSet(rs, strSQL)

If rs!Existe = 0 Then 'Insertar
  If Trim(.Text) = "" Then Exit Function
  
  strSQL = "insert into CNTX_EF_SECCIONES(COD_EF, COD_CONTABILIDAD, ITEM_ID, ITEM_ID_MADRE, PRIORIDAD, ES_TITULO,  TOTALES, DESCRIPCION, REGISTRO_USUARIO, REGISTRO_FECHA)" _
         & " values('" & scEF.Tag & "'," & gCntX_Parametros.CodigoConta & ",'" & .Text & "','"
  .Col = 2
  strSQL = strSQL & .Text & "','"
  .Col = 3
  strSQL = strSQL & .Text & "',"
  .Col = 4
  strSQL = strSQL & .Value & ","
  .Col = 5
  strSQL = strSQL & .Value & ",'"
  
  .Col = 6
  strSQL = strSQL & .Text & "','" & glogon.Usuario & "',dbo.Mygetdate())"

  Call ConectionExecute(strSQL)

  .Col = 1
  Call Bitacora("Registra", "EF Items:  " & .Text)

Else 'Actualizar

  .Col = 2
  strSQL = "update CNTX_EF_SECCIONES set ITEM_ID_MADRE = '" & .Text & "', PRIORIDAD = '"
  .Col = 3
  strSQL = strSQL & .Text & "', ES_TITULO = "
  .Col = 4
  strSQL = strSQL & .Value & ", TOTALES = "
  .Col = 5
  strSQL = strSQL & .Value & ", DESCRIPCION = '"
  .Col = 6
  strSQL = strSQL & .Text & "' where ITEM_ID = '"
  .Col = 1
  strSQL = strSQL & .Text & "' AND COD_EF = '" & scEF.Tag & "' AND COD_CONTABILIDAD = " & gCntX_Parametros.CodigoConta
  
  Call ConectionExecute(strSQL)

 .Col = 1
 Call Bitacora("Modifica", "EF Items:  " & .Text)

End If
rs.Close

End With


fxGuardar_gItems = 1

Exit Function

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Function






Private Function fxGuardar() As Long

On Error GoTo vError

fxGuardar = 0
vGrid.Row = vGrid.ActiveRow
vGrid.Col = 1

strSQL = "select isnull(count(*),0) as Existe from CNTX_EF_PERSONAL " _
       & " where COD_EF = '" & vGrid.Text & "' AND COD_CONTABILIDAD = " & gCntX_Parametros.CodigoConta
Call OpenRecordSet(rs, strSQL)

If rs!Existe = 0 Then 'Insertar
  If Trim(vGrid.Text) = "" Then Exit Function
  
  strSQL = "insert into CNTX_EF_PERSONAL(COD_CONTABILIDAD, COD_EF,DESCRIPCION, ACTIVO, REGISTRO_USUARIO, REGISTRO_FECHA)" _
         & " values(" & gCntX_Parametros.CodigoConta & ",'" & vGrid.Text & "','"
  vGrid.Col = 2
  strSQL = strSQL & vGrid.Text & "',"
  vGrid.Col = 3
  strSQL = strSQL & vGrid.Value & ",'" & glogon.Usuario & "',dbo.Mygetdate())"

  Call ConectionExecute(strSQL)

  vGrid.Col = 1
  Call Bitacora("Registra", "EF Personalizados:  " & vGrid.Text)

Else 'Actualizar

  vGrid.Col = 2
  strSQL = "update CNTX_EF_PERSONAL set descripcion = '" & vGrid.Text & "', ACTIVO = "
  vGrid.Col = 3
  strSQL = strSQL & vGrid.Value & " where COD_EF = '"
  vGrid.Col = 1
  strSQL = strSQL & vGrid.Text & "' AND COD_CONTABILIDAD = " & gCntX_Parametros.CodigoConta
  
  Call ConectionExecute(strSQL)

 vGrid.Col = 1
 Call Bitacora("Modifica", "EF Personalizados:  " & vGrid.Text)

End If
rs.Close

fxGuardar = 1

Exit Function

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Function


Private Sub sbCuentas_Consulta()

Dim pCtaInicio As String, pCtaCorte As String


On Error GoTo vError

Me.MousePointer = vbHourglass


pCtaInicio = Trim(txtCtaInicial.Text)
pCtaCorte = Trim(txtCtaFinal.Text)

strSQL = "select Cta.COD_CUENTA, Cta.COD_CUENTA_MASK, Cta.DESCRIPCION, Cta.COD_DIVISA , Cta.ACEPTA_MOVIMIENTOS " _
       & "  from CntX_Cuentas Cta" _
       & " Where Cta.COD_CONTABILIDAD = " & gCntX_Parametros.CodigoConta _
       & "  and Cta.COD_CUENTA NOT IN(select R.Cuenta " _
       & "          from CntX_EF_Cuentas Efc cross apply dbo.fxCntX_CuentasCascada_Down(Efc.cod_contabilidad,Efc.cod_Cuenta) R" _
       & "          Where Efc.COD_CONTABILIDAD = " & gCntX_Parametros.CodigoConta _
       & "            and Efc.COD_EF = '" & cboEF.ItemData(cboEF.ListIndex) & "'" _
       & "            and Efc.ITEM_ID = '" & scItem.Tag & "'    )"

       
       
If pCtaInicio <> "" And pCtaCorte <> "" Then
    pCtaInicio = fxCntX_CuentaFormato(False, pCtaInicio, 0)
    pCtaCorte = fxCntX_CuentaFormato(False, pCtaCorte, 0)
    
    strSQL = strSQL & "   and Cta.COD_CUENTA BETWEEN '" & pCtaInicio & "' AND '" & pCtaCorte & "'"

Else
    If pCtaInicio <> "" And pCtaCorte = "" Then
        strSQL = strSQL & "   and Cta.COD_CUENTA_MASK like '" & pCtaInicio & "%'"
    End If

End If
       
strSQL = strSQL & " ORDER BY Cta.COD_CUENTA"
       
lswCtas.ListItems.Clear
With lswCtas.ColumnHeaders
    .Clear
    .Add , , "Cuenta", 2500
    .Add , , "Descripción", lswCtas.Width - 2800
End With

Call OpenRecordSet(rs, strSQL)

prgBar.Max = rs.RecordCount + 1
prgBar.Value = 0

vPaso = True

Do While Not rs.EOF
 Set itmX = lswCtas.ListItems.Add(, , rs!Cod_Cuenta_Mask)
     itmX.SubItems(1) = rs!Descripcion
     
     If rs!Acepta_movimientos = 0 Then
        itmX.Bold = True
        itmX.TextBackColor = RGB(176, 211, 238)
     End If
 prgBar.Value = prgBar.Value + 1
 rs.MoveNext
Loop
rs.Close


vPaso = False

prgBar.Value = 0
Me.MousePointer = vbDefault

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical



End Sub



Private Sub sbCuentas_Consulta_Asg()

On Error GoTo vError

Me.MousePointer = vbHourglass


strSQL = "select Cta.COD_CUENTA, Cta.COD_CUENTA_MASK, Cta.DESCRIPCION, Cta.COD_DIVISA , Cta.ACEPTA_MOVIMIENTOS " _
       & " , case when isnull(Efc.COD_CUENTA,'') = '' then 0 else 1 end as 'ASIGNADO' " _
       & "  from CntX_Cuentas Cta" _
       & "  inner join CntX_EF_Cuentas Efc on Cta.COD_CONTABILIDAD = Efc.COD_CONTABILIDAD AND Cta.COD_CUENTA = Efc.COD_CUENTA" _
       & " and Efc.COD_EF = '" & cboEF.ItemData(cboEF.ListIndex) & "'" _
       & " and Efc.ITEM_ID = '" & scItem.Tag & "'" _
       & " Where Cta.COD_CONTABILIDAD = " & gCntX_Parametros.CodigoConta _
       
       
strSQL = strSQL & " ORDER BY Cta.COD_CUENTA"
       
lswCtaAsg.ListItems.Clear
With lswCtaAsg.ColumnHeaders
    .Clear
    .Add , , "Cuenta", 2500
    .Add , , "Descripción", lswCtaAsg.Width - 2800
End With

Call OpenRecordSet(rs, strSQL)

prgBar.Max = rs.RecordCount + 1
prgBar.Value = 0

vPaso = True

Do While Not rs.EOF
 Set itmX = lswCtaAsg.ListItems.Add(, , rs!Cod_Cuenta_Mask)
     itmX.SubItems(1) = rs!Descripcion
     
     If rs!Acepta_movimientos = 0 Then
        itmX.Bold = True
        itmX.TextBackColor = RGB(176, 211, 238)
     End If
     
  prgBar.Value = prgBar.Value + 1
  
 rs.MoveNext
Loop
rs.Close


vPaso = False


prgBar.Value = 0
Me.MousePointer = vbDefault

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical


End Sub


'sbFunciones_Consulta

Private Sub sbFunciones_Consulta()

On Error GoTo vError

Me.MousePointer = vbHourglass


strSQL = "select Cta.COD_FX, Cta.FX_NAME " _
       & " , case when isnull(Efc.COD_FX,'') = '' then 0 else 1 end as 'ASIGNADO' " _
       & "  from CNTX_EF_FUNCIONES Cta" _
       & "   left join CNTX_EF_FX Efc on Cta.COD_CONTABILIDAD = Efc.COD_CONTABILIDAD AND Cta.COD_FX = Efc.COD_FX" _
       & "      and Efc.COD_EF = '" & cboEF.ItemData(cboEF.ListIndex) & "'" _
       & "      and Efc.ITEM_ID = '" & scItem.Tag & "'" _
       & " Where Cta.COD_CONTABILIDAD = " & gCntX_Parametros.CodigoConta _
       & " ORDER BY Cta.COD_FX"
       
lswFx.ListItems.Clear
With lswFx.ColumnHeaders
    .Clear
    .Add , , "Función Id", 1500
    .Add , , "Descripción", lswFx.Width - 1800
End With

Call OpenRecordSet(rs, strSQL)

prgBar.Max = rs.RecordCount + 1
prgBar.Value = 0

vPaso = True

Do While Not rs.EOF
 Set itmX = lswFx.ListItems.Add(, , rs!COD_FX)
     itmX.SubItems(1) = rs!FX_NAME
     
    If rs!Asignado = 1 Then
        itmX.Checked = True
    End If
     
  prgBar.Value = prgBar.Value + 1
  
 rs.MoveNext
Loop
rs.Close


vPaso = False


prgBar.Value = 0
Me.MousePointer = vbDefault

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub






Private Sub lswCtas_ItemCheck(ByVal Item As XtremeSuiteControls.ListViewItem)

If vPaso Then Exit Sub

On Error GoTo vError

If Item.Checked Then
    strSQL = "exec spCntX_EF_Cuentas " & gCntX_Parametros.CodigoConta & ", '" & cboEF.ItemData(cboEF.ListIndex) _
           & "', '" & scItem.Tag & "', '" & fxCntX_CuentaFormato(False, Item.Text, 0) _
           & "','" & glogon.Usuario & "','A'"
    
Else
    strSQL = "exec spCntX_EF_Cuentas " & gCntX_Parametros.CodigoConta & ", '" & cboEF.ItemData(cboEF.ListIndex) _
           & "', '" & scItem.Tag & "', '" & fxCntX_CuentaFormato(False, Item.Text, 0) _
           & "','" & glogon.Usuario & "','E'"
End If

Call ConectionExecute(strSQL)


Exit Sub

vError:
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub lswFx_ItemCheck(ByVal Item As XtremeSuiteControls.ListViewItem)
If vPaso Then Exit Sub

On Error GoTo vError

If Item.Checked Then
    strSQL = "exec spCntX_EF_Fxs " & gCntX_Parametros.CodigoConta & ", '" & cboEF.ItemData(cboEF.ListIndex) _
           & "', '" & scItem.Tag & "', '" & Item.Text _
           & "','" & glogon.Usuario & "','A'"
    
Else
    strSQL = "exec spCntX_EF_Fxs " & gCntX_Parametros.CodigoConta & ", '" & cboEF.ItemData(cboEF.ListIndex) _
           & "', '" & scItem.Tag & "', '" & Item.Text _
           & "','" & glogon.Usuario & "','E'"
End If

Call ConectionExecute(strSQL)


Exit Sub

vError:
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub


Private Sub lswItems_ItemClick(ByVal Item As XtremeSuiteControls.ListViewItem)

If vPaso Then Exit Sub

lswCtas.ListItems.Clear
lswCtaAsg.ListItems.Clear

scItem.Caption = Item.Text
scItem.Tag = Item.Tag

txtCtaInicial.Text = ""
txtCtaFinal.Text = ""

chkCtaTodas(0).Value = xtpUnchecked


Select Case tcCtas.SelectedItem
    Case 0 'Ctas
        Call sbCuentas_Consulta
    Case 1 'Ctas Asignadas
        Call sbCuentas_Consulta_Asg
    Case 2 'Funciones
        Call sbFunciones_Consulta
        
End Select



End Sub




Private Sub PushButton1_Click()

End Sub

Private Sub tcCtas_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)

If vPaso Then Exit Sub

Select Case Item.Index
    Case 0 'Ctas
        Call sbCuentas_Consulta
    Case 1 'Ctas Asignadas
        Call sbCuentas_Consulta_Asg
    Case 2 'Funciones
        Call sbFunciones_Consulta
End Select

End Sub

Private Sub tcMain_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)

Select Case Item.Index

    Case 1
        
        vPaso = True
        
        scItem.Caption = "Seleccione un Item para vincular cuentas"
        scItem.Tag = ""
        
        tcCtas.Item(0).Selected = True
        lswItems.ListItems.Clear
        lswCtas.ListItems.Clear
        lswCtaAsg.ListItems.Clear
        
        strSQL = "select COD_EF AS 'IdX', DESCRIPCION AS 'ItmX' " _
                & " from CNTX_EF_PERSONAL " _
                & " Where COD_CONTABILIDAD = " & gCntX_Parametros.CodigoConta _
                & " Order by Descripcion"
        
        
        Call sbCbo_Llena_New(cboEF, strSQL, False, True)
        
        vPaso = False
    
        Call cboEF_Click
            
    Case 2
    
    
          strSQL = "select COD_EF AS 'IdX', DESCRIPCION AS 'ItmX' " _
                & " from CNTX_EF_PERSONAL " _
                & " Where COD_CONTABILIDAD = " & gCntX_Parametros.CodigoConta _
                & " Order by Descripcion"
        
        vPaso = True
        
        Call sbCbo_Llena_New(cboEFRep, strSQL, False, True)
        
        vPaso = False
    
        
        cboRepExpresado.Text = "unidad"

End Select

End Sub




Private Sub TimerX_Timer()

TimerX.Interval = 0
TimerX.Enabled = False

On Error GoTo vError

strSQL = "exec spCntx_CNTX_EF_FUNCIONES " & gCntX_Parametros.CodigoConta
Call ConectionExecute(strSQL)

Exit Sub

vError:

End Sub

Private Sub txtCtaFinal_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then
     frmCntX_ConsultaCuentas.Show vbModal
     txtCtaFinal = fxCntX_CuentaFormato(True, gCuenta, 0)
End If
End Sub

Private Sub txtCtaInicial_KeyDown(KeyCode As Integer, Shift As Integer)
     
If KeyCode = vbKeyF4 Then
     frmCntX_ConsultaCuentas.Show vbModal
     txtCtaInicial = fxCntX_CuentaFormato(True, gCuenta, 0)
End If

End Sub

Private Sub vGrid_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
If vPaso Then Exit Sub

vGrid.Row = Row
vGrid.Col = 1

scEF.Tag = vGrid.Text

vGrid.Col = 2

scEF.Caption = vGrid.Text

Call sbConsulta_Items_Mantenimiento

End Sub

Private Sub vGrid_KeyDown(KeyCode As Integer, Shift As Integer)
Dim i As Integer


If vGrid.ActiveCol = vGrid.MaxCols And (KeyCode = vbKeyReturn Or KeyCode = vbKeyTab) Then
  i = fxGuardar
  If i = 0 Then Exit Sub
  vGrid.Row = vGrid.ActiveRow
  If vGrid.MaxRows <= vGrid.ActiveRow Then
    vGrid.MaxRows = vGrid.MaxRows + 1
    vGrid.Row = vGrid.MaxRows
  End If
End If

'Inserta Linea
If KeyCode = vbKeyInsert Then
    vGrid.MaxRows = vGrid.MaxRows + 1
    vGrid.InsertRows vGrid.ActiveRow, 1
    vGrid.Row = vGrid.ActiveRow
End If


'Borrar una linea
If KeyCode = vbKeyDelete Then
     i = MsgBox("Esta Seguro que desea borrar este registro", vbYesNo)
     If i = vbYes Then
        
        vGrid.Row = vGrid.ActiveRow
        vGrid.Col = 1
        strSQL = "delete CNTX_EF_PERSONAL where COD_EF = '" & vGrid.Text & "'"
        Call ConectionExecute(strSQL)
        strSQL = vGrid.Text
        vGrid.Col = 1
        Call Bitacora("Elimina", "EF Personalizados:  " & vGrid.Text)
        
        Call sbConsulta
     
     End If
End If

End Sub






Private Sub gItems_KeyDown(KeyCode As Integer, Shift As Integer)
Dim i As Integer

With gItems

If .ActiveCol = .MaxCols And (KeyCode = vbKeyReturn Or KeyCode = vbKeyTab) Then
  i = fxGuardar_gItems
  If i = 0 Then Exit Sub
  .Row = .ActiveRow
  If .MaxRows <= .ActiveRow Then
    .MaxRows = .MaxRows + 1
    .Row = .MaxRows
  End If
End If

'Inserta Linea
If KeyCode = vbKeyInsert Then
    .MaxRows = .MaxRows + 1
    .InsertRows .ActiveRow, 1
    .Row = .ActiveRow
End If


'Borrar una linea
If KeyCode = vbKeyDelete Then
     i = MsgBox("Esta Seguro que desea borrar este registro", vbYesNo)
     If i = vbYes Then
        
        .Row = .ActiveRow
        .Col = 1
        strSQL = "delete CNTX_EF_SECCIONES where ITEM_ID = '" & .Text _
               & "' AND COD_EF = '" & scEF.Tag & "' AND COD_CONTABILIDAD = " & gCntX_Parametros.CodigoConta
        Call ConectionExecute(strSQL)
        strSQL = .Text
        .Col = 1
        Call Bitacora("Elimina", "EF Items: " & .Text)
        
        Call sbConsulta_Items_Mantenimiento
     
     End If
End If

End With

End Sub



