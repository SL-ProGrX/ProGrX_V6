VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "ComCt332.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#20.3#0"; "Codejock.Controls.v20.3.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#20.3#0"; "Codejock.ShortcutBar.v20.3.0.ocx"
Begin VB.Form frmCO_CJ_Tramite 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cobros: Registro del Trámite"
   ClientHeight    =   8820
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9780
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8820
   ScaleWidth      =   9780
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.GroupBox fraNotas 
      Height          =   4095
      Left            =   9240
      TabIndex        =   74
      Top             =   4320
      Visible         =   0   'False
      Width           =   9255
      _Version        =   1310723
      _ExtentX        =   16325
      _ExtentY        =   7223
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      BorderStyle     =   2
      Begin XtremeSuiteControls.FlatEdit txtNotasCambios 
         Height          =   1635
         Left            =   120
         TabIndex        =   76
         Top             =   600
         Width           =   9015
         _Version        =   1310723
         _ExtentX        =   15901
         _ExtentY        =   2884
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
      Begin XtremeSuiteControls.PushButton btnNotas 
         Height          =   375
         Left            =   8040
         TabIndex        =   77
         Top             =   2520
         Width           =   1095
         _Version        =   1310723
         _ExtentX        =   1931
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Guardar"
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
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
         Height          =   375
         Left            =   0
         TabIndex        =   75
         Top             =   0
         Width           =   9255
         _Version        =   1310723
         _ExtentX        =   16325
         _ExtentY        =   661
         _StockProps     =   14
         Caption         =   "Notas de la Gestión de Cambios"
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
   End
   Begin XtremeSuiteControls.TabControl tcMain 
      Height          =   5895
      Left            =   120
      TabIndex        =   17
      Top             =   2640
      Width           =   9495
      _Version        =   1310723
      _ExtentX        =   16748
      _ExtentY        =   10398
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
      Item(0).Caption =   "Estado"
      Item(0).ControlCount=   13
      Item(0).Control(0)=   "Label1(5)"
      Item(0).Control(1)=   "Label1(6)"
      Item(0).Control(2)=   "Label1(7)"
      Item(0).Control(3)=   "txtAbogado"
      Item(0).Control(4)=   "txtAbogadoDesc"
      Item(0).Control(5)=   "txtBuffete"
      Item(0).Control(6)=   "txtBuffeteDesc"
      Item(0).Control(7)=   "Label1(8)"
      Item(0).Control(8)=   "txtNotas"
      Item(0).Control(9)=   "Label1(10)"
      Item(0).Control(10)=   "cboJuzgado"
      Item(0).Control(11)=   "cboJuicio"
      Item(0).Control(12)=   "gbDatos"
      Item(1).Caption =   "Detalle del Trámite"
      Item(1).ControlCount=   6
      Item(1).Control(0)=   "vGrid"
      Item(1).Control(1)=   "dtpFechaExp"
      Item(1).Control(2)=   "Label3(5)"
      Item(1).Control(3)=   "Label1(20)"
      Item(1).Control(4)=   "txtExp"
      Item(1).Control(5)=   "btnExpediente"
      Item(2).Caption =   "Embargables"
      Item(2).ControlCount=   10
      Item(2).Control(0)=   "lsw"
      Item(2).Control(1)=   "txtEmbargo"
      Item(2).Control(2)=   "txtEmbargoDesc"
      Item(2).Control(3)=   "Label1(22)"
      Item(2).Control(4)=   "Label1(25)"
      Item(2).Control(5)=   "txtEmbargoMonto"
      Item(2).Control(6)=   "chkEmbargoAplica"
      Item(2).Control(7)=   "txtEmbargoNotas"
      Item(2).Control(8)=   "Label1(27)"
      Item(2).Control(9)=   "btnEmbargo"
      Begin XtremeSuiteControls.ListView lsw 
         Height          =   2895
         Left            =   -69880
         TabIndex        =   59
         Top             =   2880
         Visible         =   0   'False
         Width           =   9255
         _Version        =   1310723
         _ExtentX        =   16325
         _ExtentY        =   5106
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
      Begin XtremeSuiteControls.CheckBox chkEmbargoAplica 
         Height          =   255
         Left            =   -66520
         TabIndex        =   70
         Top             =   1080
         Visible         =   0   'False
         Width           =   3495
         _Version        =   1310723
         _ExtentX        =   6165
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Aplica Embargo ?"
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
         Appearance      =   17
      End
      Begin XtremeSuiteControls.FlatEdit txtEmbargoDesc 
         Height          =   315
         Left            =   -66760
         TabIndex        =   66
         Top             =   600
         Visible         =   0   'False
         Width           =   5895
         _Version        =   1310723
         _ExtentX        =   10398
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
      Begin XtremeSuiteControls.GroupBox gbDatos 
         Height          =   3135
         Left            =   120
         TabIndex        =   31
         Top             =   2640
         Width           =   9255
         _Version        =   1310723
         _ExtentX        =   16325
         _ExtentY        =   5530
         _StockProps     =   79
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         BorderStyle     =   2
         Begin XtremeSuiteControls.FlatEdit txtMonto 
            Height          =   315
            Left            =   1560
            TabIndex        =   33
            Top             =   240
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
         Begin XtremeSuiteControls.FlatEdit txtIntCorVenc 
            Height          =   315
            Left            =   1560
            TabIndex        =   37
            Top             =   720
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
         Begin XtremeSuiteControls.FlatEdit txtIntCorAtrasado 
            Height          =   315
            Left            =   1560
            TabIndex        =   41
            Top             =   1080
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
         Begin XtremeSuiteControls.FlatEdit txtIntMoratorio 
            Height          =   315
            Left            =   1560
            TabIndex        =   45
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
            Left            =   1560
            TabIndex        =   49
            Top             =   1800
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
         Begin XtremeSuiteControls.FlatEdit txtPolizas 
            Height          =   315
            Left            =   1560
            TabIndex        =   51
            Top             =   2160
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
         Begin XtremeSuiteControls.FlatEdit txtSaldo 
            Height          =   315
            Left            =   1560
            TabIndex        =   53
            Top             =   2640
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
         Begin XtremeSuiteControls.FlatEdit txtTotalDeuda 
            Height          =   315
            Left            =   7200
            TabIndex        =   35
            Top             =   240
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
         Begin XtremeSuiteControls.FlatEdit txtTotalGastos 
            Height          =   315
            Left            =   7200
            TabIndex        =   39
            Top             =   600
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
         Begin XtremeSuiteControls.FlatEdit txtSentenciaMonto 
            Height          =   315
            Left            =   7200
            TabIndex        =   43
            Top             =   1200
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
         Begin XtremeSuiteControls.FlatEdit txtSentenciaFecha 
            Height          =   315
            Left            =   7200
            TabIndex        =   47
            Top             =   1560
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
         Begin XtremeSuiteControls.FlatEdit txtTotalRecuperado 
            Height          =   315
            Left            =   7200
            TabIndex        =   55
            Top             =   2160
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
         Begin XtremeSuiteControls.FlatEdit txtTotalAplicado 
            Height          =   315
            Left            =   7200
            TabIndex        =   57
            Top             =   2520
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
         Begin XtremeSuiteControls.Label Label1 
            Height          =   255
            Index           =   26
            Left            =   5400
            TabIndex        =   56
            Top             =   2520
            Width           =   1455
            _Version        =   1310723
            _ExtentX        =   2566
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Total Aplicado"
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
         Begin XtremeSuiteControls.Label Label1 
            Height          =   255
            Index           =   24
            Left            =   5400
            TabIndex        =   54
            Top             =   2160
            Width           =   1455
            _Version        =   1310723
            _ExtentX        =   2566
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Total Recuperado"
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
         Begin XtremeSuiteControls.Label Label1 
            Height          =   255
            Index           =   23
            Left            =   120
            TabIndex        =   52
            Top             =   2640
            Width           =   1215
            _Version        =   1310723
            _ExtentX        =   2143
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Saldo"
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
         Begin XtremeSuiteControls.Label Label1 
            Height          =   255
            Index           =   21
            Left            =   120
            TabIndex        =   50
            Top             =   2160
            Width           =   1215
            _Version        =   1310723
            _ExtentX        =   2143
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Pólizas"
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
         Begin XtremeSuiteControls.Label Label1 
            Height          =   255
            Index           =   19
            Left            =   120
            TabIndex        =   48
            Top             =   1800
            Width           =   1215
            _Version        =   1310723
            _ExtentX        =   2143
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Cargos"
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
         Begin XtremeSuiteControls.Label Label1 
            Height          =   255
            Index           =   18
            Left            =   5400
            TabIndex        =   46
            Top             =   1560
            Width           =   1455
            _Version        =   1310723
            _ExtentX        =   2566
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Sentencia Fecha"
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
         Begin XtremeSuiteControls.Label Label1 
            Height          =   255
            Index           =   17
            Left            =   120
            TabIndex        =   44
            Top             =   1440
            Width           =   1695
            _Version        =   1310723
            _ExtentX        =   2990
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Int. Moratorios"
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
         Begin XtremeSuiteControls.Label Label1 
            Height          =   255
            Index           =   16
            Left            =   5400
            TabIndex        =   42
            Top             =   1200
            Width           =   1455
            _Version        =   1310723
            _ExtentX        =   2566
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Sentencia Monto"
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
         Begin XtremeSuiteControls.Label Label1 
            Height          =   255
            Index           =   15
            Left            =   120
            TabIndex        =   40
            Top             =   1080
            Width           =   1215
            _Version        =   1310723
            _ExtentX        =   2143
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Int. Cor. Atra."
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
         Begin XtremeSuiteControls.Label Label1 
            Height          =   255
            Index           =   14
            Left            =   5400
            TabIndex        =   38
            Top             =   600
            Width           =   1455
            _Version        =   1310723
            _ExtentX        =   2566
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Total Gastos"
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
         Begin XtremeSuiteControls.Label Label1 
            Height          =   255
            Index           =   13
            Left            =   120
            TabIndex        =   36
            Top             =   720
            Width           =   1215
            _Version        =   1310723
            _ExtentX        =   2143
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Int. Cor. Venc."
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
         Begin XtremeSuiteControls.Label Label1 
            Height          =   255
            Index           =   12
            Left            =   5400
            TabIndex        =   34
            Top             =   240
            Width           =   1455
            _Version        =   1310723
            _ExtentX        =   2566
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Total Adeudado"
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
         Begin XtremeSuiteControls.Label Label1 
            Height          =   255
            Index           =   11
            Left            =   120
            TabIndex        =   32
            Top             =   240
            Width           =   1215
            _Version        =   1310723
            _ExtentX        =   2143
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Monto Inicial"
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
      Begin XtremeSuiteControls.FlatEdit txtBuffeteDesc 
         Height          =   315
         Left            =   3360
         TabIndex        =   24
         Top             =   1320
         Width           =   5895
         _Version        =   1310723
         _ExtentX        =   10398
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
      Begin XtremeSuiteControls.FlatEdit txtAbogadoDesc 
         Height          =   315
         Left            =   3360
         TabIndex        =   21
         Top             =   960
         Width           =   5895
         _Version        =   1310723
         _ExtentX        =   10398
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
      Begin XtremeSuiteControls.FlatEdit txtAbogado 
         Height          =   315
         Left            =   1680
         TabIndex        =   20
         ToolTipText     =   "Preisone F4 para Consultar"
         Top             =   960
         Width           =   1695
         _Version        =   1310723
         _ExtentX        =   2990
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
      Begin XtremeSuiteControls.FlatEdit txtBuffete 
         Height          =   315
         Left            =   1680
         TabIndex        =   23
         Top             =   1320
         Width           =   1695
         _Version        =   1310723
         _ExtentX        =   2990
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
      Begin XtremeSuiteControls.FlatEdit txtNotas 
         Height          =   795
         Left            =   1680
         TabIndex        =   27
         Top             =   1680
         Width           =   7575
         _Version        =   1310723
         _ExtentX        =   13361
         _ExtentY        =   1402
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
      Begin XtremeSuiteControls.ComboBox cboJuzgado 
         Height          =   330
         Left            =   1680
         TabIndex        =   29
         Top             =   600
         Width           =   3735
         _Version        =   1310723
         _ExtentX        =   6588
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
      Begin XtremeSuiteControls.ComboBox cboJuicio 
         Height          =   330
         Left            =   5400
         TabIndex        =   30
         Top             =   600
         Width           =   3855
         _Version        =   1310723
         _ExtentX        =   6800
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
      Begin FPSpreadADO.fpSpread vGrid 
         Height          =   4695
         Left            =   -69880
         TabIndex        =   58
         Top             =   1080
         Visible         =   0   'False
         Width           =   9255
         _Version        =   524288
         _ExtentX        =   16325
         _ExtentY        =   8281
         _StockProps     =   64
         BackColorStyle  =   1
         BorderStyle     =   0
         DisplayRowHeaders=   0   'False
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
         SpreadDesigner  =   "frmCO_CJ_Tramite.frx":0000
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin XtremeSuiteControls.DateTimePicker dtpFechaExp 
         Height          =   315
         Left            =   -63400
         TabIndex        =   60
         Top             =   480
         Visible         =   0   'False
         Width           =   1335
         _Version        =   1310723
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
      Begin XtremeSuiteControls.FlatEdit txtExp 
         Height          =   315
         Left            =   -68320
         TabIndex        =   63
         Top             =   480
         Visible         =   0   'False
         Width           =   3015
         _Version        =   1310723
         _ExtentX        =   5318
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
      Begin XtremeSuiteControls.PushButton btnExpediente 
         Height          =   375
         Left            =   -61960
         TabIndex        =   64
         Top             =   465
         Visible         =   0   'False
         Width           =   1095
         _Version        =   1310723
         _ExtentX        =   1931
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Guardar"
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
      Begin XtremeSuiteControls.FlatEdit txtEmbargo 
         Height          =   315
         Left            =   -68560
         TabIndex        =   65
         Top             =   600
         Visible         =   0   'False
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
      Begin XtremeSuiteControls.FlatEdit txtEmbargoMonto 
         Height          =   315
         Left            =   -68560
         TabIndex        =   69
         Top             =   1080
         Visible         =   0   'False
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtEmbargoNotas 
         Height          =   795
         Left            =   -68560
         TabIndex        =   71
         Top             =   1560
         Visible         =   0   'False
         Width           =   7935
         _Version        =   1310723
         _ExtentX        =   13996
         _ExtentY        =   1402
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
      Begin XtremeSuiteControls.PushButton btnEmbargo 
         Height          =   375
         Left            =   -61720
         TabIndex        =   73
         Top             =   2400
         Visible         =   0   'False
         Width           =   1095
         _Version        =   1310723
         _ExtentX        =   1931
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Guardar"
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
      Begin XtremeSuiteControls.Label Label1 
         Height          =   255
         Index           =   27
         Left            =   -69760
         TabIndex        =   72
         Top             =   1560
         Visible         =   0   'False
         Width           =   1215
         _Version        =   1310723
         _ExtentX        =   2143
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
      Begin XtremeSuiteControls.Label Label1 
         Height          =   255
         Index           =   25
         Left            =   -69760
         TabIndex        =   68
         Top             =   1080
         Visible         =   0   'False
         Width           =   1215
         _Version        =   1310723
         _ExtentX        =   2143
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
      Begin XtremeSuiteControls.Label Label1 
         Height          =   255
         Index           =   22
         Left            =   -69760
         TabIndex        =   67
         Top             =   600
         Visible         =   0   'False
         Width           =   1215
         _Version        =   1310723
         _ExtentX        =   2143
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Tipo Bien"
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
      Begin XtremeSuiteControls.Label Label1 
         Height          =   255
         Index           =   20
         Left            =   -69640
         TabIndex        =   62
         Top             =   480
         Visible         =   0   'False
         Width           =   1215
         _Version        =   1310723
         _ExtentX        =   2143
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "N° Expediente"
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
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha Expediente"
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
         Index           =   5
         Left            =   -65560
         TabIndex        =   61
         Top             =   480
         Visible         =   0   'False
         Width           =   1965
      End
      Begin XtremeSuiteControls.Label Label1 
         Height          =   255
         Index           =   10
         Left            =   120
         TabIndex        =   28
         Top             =   1680
         Width           =   1215
         _Version        =   1310723
         _ExtentX        =   2143
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
      Begin XtremeSuiteControls.Label Label1 
         Height          =   255
         Index           =   8
         Left            =   120
         TabIndex        =   25
         Top             =   1320
         Width           =   1215
         _Version        =   1310723
         _ExtentX        =   2143
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Firma/Buffete"
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
      Begin XtremeSuiteControls.Label Label1 
         Height          =   255
         Index           =   7
         Left            =   120
         TabIndex        =   22
         Top             =   960
         Width           =   1215
         _Version        =   1310723
         _ExtentX        =   2143
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Abogado"
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
      Begin XtremeSuiteControls.Label Label1 
         Height          =   255
         Index           =   6
         Left            =   5400
         TabIndex        =   19
         Top             =   360
         Width           =   1215
         _Version        =   1310723
         _ExtentX        =   2143
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Tipo de Juicio:"
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
      Begin XtremeSuiteControls.Label Label1 
         Height          =   255
         Index           =   5
         Left            =   1680
         TabIndex        =   18
         Top             =   360
         Width           =   1215
         _Version        =   1310723
         _ExtentX        =   2143
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Juzgado:"
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
   End
   Begin MSComctlLib.StatusBar StatusBarX 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   8565
      Width           =   9780
      _ExtentX        =   17251
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   7832
            MinWidth        =   7832
            Object.ToolTipText     =   "Oficina"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   4304
            MinWidth        =   4304
            Object.ToolTipText     =   "Linea"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   4304
            MinWidth        =   4304
            Object.ToolTipText     =   "Recurso"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin ComCtl3.CoolBar CoolBarX 
      Align           =   1  'Align Top
      Height          =   390
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   9780
      _ExtentX        =   17251
      _ExtentY        =   688
      BandCount       =   2
      _CBWidth        =   9780
      _CBHeight       =   390
      _Version        =   "6.7.9839"
      Child1          =   "tlb"
      MinHeight1      =   330
      Width1          =   2955
      NewRow1         =   0   'False
      Child2          =   "tlbAux"
      MinHeight2      =   330
      Width2          =   2520
      NewRow2         =   0   'False
      Begin MSComctlLib.Toolbar tlb 
         Height          =   330
         Left            =   165
         TabIndex        =   3
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
               Object.ToolTipText     =   "Boleta de Registro"
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
         Left            =   3150
         TabIndex        =   2
         Top             =   30
         Width           =   6540
         _ExtentX        =   11536
         _ExtentY        =   582
         ButtonWidth     =   1826
         ButtonHeight    =   582
         Style           =   1
         TextAlignment   =   1
         ImageList       =   "ImageList1"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   5
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Proceso"
               Key             =   "Proceso"
               Object.ToolTipText     =   "Registro del Estado del Trámite"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Gastos"
               Key             =   "Gastos"
               Object.ToolTipText     =   "Registro de Gastos del Trámite"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
               Style           =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
               Caption         =   "Recuperación"
               Key             =   "Recuperacion"
               Object.ToolTipText     =   "Cobros Ejecutados en el Juzgado"
               ImageIndex      =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
               Caption         =   "Aplicado"
               Key             =   "Aplicado"
               Object.ToolTipText     =   "Aplicación a la deuda"
               ImageIndex      =   4
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   9240
      Top             =   -120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCO_CJ_Tramite.frx":15D4
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCO_CJ_Tramite.frx":16EE
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCO_CJ_Tramite.frx":1812
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCO_CJ_Tramite.frx":190F
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComCtl2.FlatScrollBar FlatScrollBar 
      Height          =   255
      Left            =   3480
      TabIndex        =   4
      Top             =   720
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   450
      _Version        =   393216
      Arrows          =   65536
      Orientation     =   1638401
   End
   Begin XtremeSuiteControls.FlatEdit txtOperacion 
      Height          =   405
      Left            =   1560
      TabIndex        =   5
      Top             =   1200
      Width           =   1815
      _Version        =   1310723
      _ExtentX        =   3201
      _ExtentY        =   714
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
      Text            =   "0000"
      Alignment       =   2
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtLineaCod 
      Height          =   315
      Left            =   1560
      TabIndex        =   6
      Top             =   1800
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
   Begin XtremeSuiteControls.FlatEdit txtCedula 
      Height          =   315
      Left            =   1560
      TabIndex        =   7
      Top             =   2160
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
   Begin XtremeSuiteControls.FlatEdit txtLineaDesc 
      Height          =   315
      Left            =   3360
      TabIndex        =   8
      Top             =   1800
      Width           =   6255
      _Version        =   1310723
      _ExtentX        =   11033
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
   Begin XtremeSuiteControls.FlatEdit txtNombre 
      Height          =   315
      Left            =   3360
      TabIndex        =   9
      Top             =   2160
      Width           =   6255
      _Version        =   1310723
      _ExtentX        =   11033
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
   Begin XtremeSuiteControls.FlatEdit txtTramite 
      Height          =   435
      Left            =   1560
      TabIndex        =   10
      Top             =   600
      Width           =   1815
      _Version        =   1310723
      _ExtentX        =   3201
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
      Text            =   "0000"
      BackColor       =   16777152
      Alignment       =   2
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtExpediente 
      Height          =   435
      Left            =   6480
      TabIndex        =   15
      Top             =   600
      Width           =   3135
      _Version        =   1310723
      _ExtentX        =   5530
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
      Text            =   "0000"
      BackColor       =   16777152
      Alignment       =   2
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.ComboBox cboProceso 
      Height          =   405
      Left            =   3360
      TabIndex        =   26
      Top             =   1200
      Width           =   6270
      _Version        =   1310723
      _ExtentX        =   11060
      _ExtentY        =   714
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   12
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
   Begin XtremeSuiteControls.Label Label1 
      Height          =   375
      Index           =   4
      Left            =   5040
      TabIndex        =   16
      Top             =   600
      Width           =   1455
      _Version        =   1310723
      _ExtentX        =   2566
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Expediente:"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Transparent     =   -1  'True
      WordWrap        =   -1  'True
   End
   Begin XtremeSuiteControls.Label Label1 
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   14
      Top             =   1320
      Width           =   1215
      _Version        =   1310723
      _ExtentX        =   2143
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "No. Operación"
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
   Begin XtremeSuiteControls.Label Label1 
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   13
      Top             =   1800
      Width           =   1215
      _Version        =   1310723
      _ExtentX        =   2143
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Línea Crédito"
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
   Begin XtremeSuiteControls.Label Label1 
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   12
      Top             =   2160
      Width           =   1215
      _Version        =   1310723
      _ExtentX        =   2143
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Identificación"
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
   Begin XtremeSuiteControls.Label Label1 
      Height          =   375
      Index           =   3
      Left            =   240
      TabIndex        =   11
      Top             =   600
      Width           =   1215
      _Version        =   1310723
      _ExtentX        =   2143
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Tramite Id:"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Transparent     =   -1  'True
      WordWrap        =   -1  'True
   End
   Begin VB.Image ImgAutorizacion 
      Height          =   255
      Left            =   4200
      Top             =   720
      Width           =   255
   End
End
Attribute VB_Name = "frmCO_CJ_Tramite"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vMensaje        As String  'Envia Mensajes en Fallas de Verificacion
Dim vEdita          As Boolean 'Indica si se esta actualizando o insertando
Dim vPaso           As Boolean 'Control de Activacion de Controles en proceso de carga
Dim vScroll         As Boolean

Dim vFecha          As Date
Dim vLinea As Long, vAbogado As Long
Dim vJuzgado As String, vJuicio As String, vBufete As String

Dim vMensajeCambios As String

Dim strSQL As String, rs As New ADODB.Recordset


Private Function fxValida() As Boolean

fxValida = True
vMensaje = ""
If vEdita = False Then
    strSQL = "select count(*) as 'existe' from CBR_CJ_TRAMITE where id_solicitud = " & txtOperacion & " and SENTENCIA_MONTO is null"
    Call OpenRecordSet(rs, strSQL)
    If rs!Existe > 0 Then vMensaje = vMensaje & vbCrLf & "- La Operación ya cuenta con un tramite!"
End If
    
If Len(txtNotas.Text) <= 10 Then vMensaje = vMensaje & vbCrLf & "- Indique una Nota vailda!"


If Len(vMensaje) > 0 Then
  fxValida = False
  MsgBox vMensaje, vbCritical
End If
End Function



Private Sub btnEmbargo_Click()

On Error GoTo vError

            
strSQL = "exec spCbr_CJ_Tramite_Embargable " & txtTramite.Text & ", " & vLinea & ", '" & txtEmbargo.Text & "', " & chkEmbargoAplica.Value _
       & ", " & CCur(txtEmbargoMonto.Text) & ", '" & Trim(txtEmbargoNotas.Text) & "','" & glogon.Usuario & "'"

Call ConectionExecute(strSQL)

MsgBox "Informacion guardada satsfactoriamente...."

vLinea = 0
 
Exit Sub
vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub btnExpediente_Click()

On Error GoTo vError


strSQL = "exec spCbr_CJ_Tramite_Expediente_Update " & txtTramite.Text & ", '" & Trim(txtExp.Text) & "','" & Format(dtpFechaExp.Value, "yyyy-mm-dd") & "','" & glogon.Usuario & "'"
Call ConectionExecute(strSQL)

MsgBox "Expediente actualizado satisfactoriamente...", vbInformation


Call sbConsulta


Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical


End Sub

Private Sub btnNotas_Click()

fraNotas.Visible = False
Call sbGuardar

End Sub

Private Sub FlatScrollBar_Change()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

tcMain.Item(0).Selected = True

If txtTramite.Text = "" Then txtTramite.Text = "0"
If FlatScrollBar.Tag = "" Then FlatScrollBar.Tag = 0

strSQL = "select Top 1 cod_tramite from CBR_CJ_Tramite"

If FlatScrollBar.Value > CLng(FlatScrollBar.Tag) Then
   strSQL = strSQL & " where cod_tramite > " & txtTramite & " order by cod_tramite asc"
Else
   strSQL = strSQL & " where cod_tramite < " & txtTramite & " order by cod_tramite desc"
End If

FlatScrollBar.Tag = FlatScrollBar.Value

Call OpenRecordSet(rs, strSQL)
If Not rs.EOF And Not rs.BOF Then
  txtTramite.Text = rs!COD_TRAMITE
  Call sbConsulta
End If
rs.Close

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub Form_Activate()
vModulo = 6


End Sub

Private Sub Form_Load()

 vModulo = 6
 
With lsw.ColumnHeaders
    .Clear
    .Add , , "Línea", 1200
    .Add , , "Código", 1200, vbCenter
    .Add , , "Descripción", 3100
    .Add , , "Monto", 1600, vbRightJustify
    .Add , , "Aplica?", 1100, vbCenter
    .Add , , "Notas", 3100
End With
 
 vEdita = True
 
 vFecha = fxFechaServidor
 
 
 Call sbToolBarIconos(tlb)
 Call sbToolBar(tlb, "activo")

 Call Formularios(Me)
 Call RefrescaTags(Me)
 

Call sbLimpiaPantalla


End Sub

Private Sub sbLimpiaPantalla()

Me.MousePointer = vbHourglass
vPaso = True

 tlbAux.Buttons.Item(1).Enabled = False
 tlbAux.Buttons.Item(2).Enabled = False
 tlbAux.Buttons.Item(4).Enabled = False
 tlbAux.Buttons.Item(5).Enabled = False


If cboProceso.ListCount = 0 Then
    strSQL = "select cod_proceso as 'IdX', rtrim(descripcion) as ItmX" _
             & " from  cbr_cj_proceso where activo = 1 and orden = '" & fxMinimoProceso & "'"
    
    Call sbCbo_Llena_New(cboProceso, strSQL, False, True)
End If

If cboJuzgado.ListCount <= 0 Then
    strSQL = "select cod_juzgado as 'IdX',  rtrim(nombre) as ItmX" _
             & " from  cbr_cj_juzgados where activo = 1"
    Call sbCbo_Llena_New(cboJuzgado, strSQL, False, True)
End If

If cboJuicio.ListCount <= 0 Then
    strSQL = "select Tipo_Juicio as 'IdX', rtrim(Descripcion) as ItmX" _
             & " from  cbr_cj_Tipos_Juicios where activo = 1"
    
    Call sbCbo_Llena_New(cboJuicio, strSQL, False, True)
End If


dtpFechaExp.Value = vFecha


txtAbogado.Text = ""
txtAbogadoDesc.Text = ""

txtBuffete.Text = ""
txtBuffeteDesc.Text = ""

txtOperacion.Text = ""

txtCedula.Text = ""
txtNombre.Text = ""
txtLineaCod.Text = ""
txtLineaDesc.Text = ""

txtNotas.Text = ""

txtMonto.Text = "0.00"
txtIntCorVenc.Text = "0.00"
txtIntCorAtrasado.Text = "0.00"
txtIntMoratorio.Text = "0.00"
txtCargos.Text = "0.00"
txtPolizas.Text = "0.00"
txtSaldo.Text = "0.00"
txtTotalDeuda.Text = "0.00"
txtTotalGastos.Text = "0.00"
txtSentenciaMonto.Text = "0.00"
txtSentenciaFecha.Text = ""
txtTotalRecuperado.Text = "0.00"
txtTotalAplicado.Text = "0.00"

txtExp.Text = ""
txtExpediente.Text = ""
 
 tcMain.Item(0).Selected = True
 tcMain.Item(1).Enabled = False

 StatusBarX.Panels(1).Text = ""
 StatusBarX.Panels(2).Text = ""
 StatusBarX.Panels(3).Text = ""
 vPaso = False
 
Me.MousePointer = vbDefault

End Sub



Private Sub sbConsulta()

On Error GoTo vError

vPaso = True


strSQL = "exec spCbr_CJ_Tramite_Consulta " & txtTramite.Text & ""

Call OpenRecordSet(rs, strSQL, 0)

If Not rs.EOF And Not rs.BOF Then
  'Call sbToolBar(tlb, "activo")
  
  If rs!SENTENCIA_INDICA = 0 Then
    tlbAux.Buttons(1).Enabled = True
    tlbAux.Buttons(2).Enabled = True
  Else
    tlbAux.Buttons(1).Enabled = False
    tlbAux.Buttons(2).Enabled = False
  End If
  
  vEdita = True
  
  
  txtOperacion.Text = CStr(rs!Id_Solicitud)
  
  txtCedula.Text = Trim(rs!Cedula)
  txtNombre.Text = rs!Nombre
  txtLineaCod.Text = Trim(rs!Codigo)
  txtLineaDesc.Text = rs!LineaDesc
  
    
  vAbogado = rs!COD_ABOGADO
  vJuzgado = Trim(rs!cod_Juzgado)
  vJuicio = rs!TIPO_JUICIO
  vBufete = IIf(rs!BuffeteId = 0, "", rs!BuffeteId)
  
  vPaso = True
  
  txtAbogado.Text = rs!COD_ABOGADO
  txtAbogadoDesc.Text = rs!AbogadoDesc
  

  txtBuffete.Text = IIf(rs!BuffeteId = 0, "", rs!BuffeteId)
  txtBuffeteDesc.Text = rs!BuffeteDesc
  
    
  Call sbCboAsignaDato(cboJuzgado, rs!JuzgadoDesc, True, rs!cod_Juzgado)
  Call sbCboAsignaDato(cboJuicio, rs!JuicioDesc, True, rs!TIPO_JUICIO)
  Call sbCboAsignaDato(cboProceso, rs!ProcesoDesc, True, rs!cod_Proceso)
    
  txtNotas = IIf(IsNull(rs!notas), "", rs!notas)
  
  txtSentenciaMonto = IIf(IsNull(rs!SENTENCIA_MONTO), 0, rs!SENTENCIA_MONTO)
  txtSentenciaFecha = IIf(IsNull(rs!SENTENCIA_FECHA), "", rs!SENTENCIA_FECHA)
  txtTotalAplicado = Format(rs!TOTAL_APLICADO, "Standard")
  txtTotalGastos = Format(rs!total_gasto, "Standard")
  txtExpediente.Text = IIf(IsNull(rs!EXPEDIENTE_NUMERO), "", rs!EXPEDIENTE_NUMERO)
  
  txtIntCorVenc.Text = Format(rs!base_int_cor_venc, "Standard")
  
  Call sbDatosOperacion
  
  
  txtOperacion.SetFocus
  
  tcMain.Item(1).Enabled = True
  
  vPaso = False
 
Else
    MsgBox "No existe la Trámite, verifique!", vbCritical
    tcMain.Item(1).Enabled = False
End If

Call RefrescaTags(Me)

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
  vPaso = False

End Sub


Private Sub lsw_ItemClick(ByVal Item As XtremeSuiteControls.ListViewItem)

vLinea = Item.Text
txtEmbargo.Text = Item.SubItems(1)
txtEmbargoDesc.Text = Item.SubItems(2)
txtMonto.Text = Format(Item.SubItems(3), "Standard")

If Mid(Item.SubItems(4), 1) = "A" Then
   chkEmbargoAplica.Value = vbChecked
Else
   chkEmbargoAplica.Value = vbUnchecked
End If

txtNotas.Text = Item.SubItems(5)

End Sub



Private Sub tcMain_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)

If Not IsNumeric(txtTramite.Text) Then
    tcMain.Item(0).Selected = True
End If

Select Case Item.Index
  Case 0 'General
  
  Case 1 'Detalles
    vGrid.Enabled = True
    Call vGrid_SheetChanged(1, 1)
  
  Case 2
    strSQL = "select COUNT(*) as cerrado from CBR_CJ_TRAMITE_PROCESO where APLICA_CIERRE_SENTENCIA = 1"
    Call OpenRecordSet(rs, strSQL)
     
    If rs!cerrado = 1 Then
      btnEmbargo.Enabled = False
    Else
      btnEmbargo.Enabled = True
    End If
    rs.Close
    
    vLinea = 0
    Call sbEmbargos_List
    
 Case Else
  
End Select

End Sub

Private Sub tlb_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case UCase(Button.Key)
    Case "INSERTAR", "NUEVO"
      vEdita = False
      Call sbLimpiaPantalla
      txtTramite.Text = ""
      txtOperacion.SetFocus
      Call sbToolBar(tlb, "edicion")
    Case "MODIFICAR", "EDITAR"
      vEdita = True
      cboJuzgado.SetFocus
      Call sbToolBar(tlb, "edicion")
    Case "BORRAR"
'      Call sbBorrar
    Case "GUARDAR", "SALVAR"
     If fxValida Then Call sbGuardar
    Case "DESHACER"
      Call sbToolBar(tlb, "activo")
      If txtTramite.Text = "" Then
        Call sbLimpiaPantalla
        Call sbToolBar(tlb, "nuevo")
        vEdita = True
      Else
        Call sbConsulta
      End If

    Case "CONSULTAR"
'       gBusquedas.Columna = "nombre"
'       gBusquedas.Orden = "nombre"
'       gBusquedas.Consulta = "select cod_abogado,nombre from Cbr_Cj_Abogados"
'       frmBusquedas.Show vbModal
'       txtCodigo.SetFocus
'       txtCodigo = gBusquedas.Resultado
'       txtNombre.SetFocus

    Case "REPORTES"

    Case "AYUDA"
        frmContenedor.CD.HelpContext = Me.HelpContextID
        frmContenedor.CD.ShowHelp

End Select

End Sub


Private Sub txtAbogado_Change()

If txtAbogado.Text = "" Then Exit Sub

If vPaso Then Exit Sub

strSQL = "select COD_BUFETE from CBR_CJ_ABOGADOS where COD_ABOGADO = " & txtAbogado.Text
Call OpenRecordSet(rs, strSQL)

If Not rs.EOF Then
    strSQL = "select isnull(rtrim(cod_bufete),'') as 'IdX', isnull(rtrim(nombre),'') as ItmX" _
             & " from  cbr_cj_bufetes where activo = 1 and cod_bufete = '" & rs!cod_bufete & "'"
Else
    strSQL = "select isnull(rtrim(cod_bufete),'') as 'IdX', isnull(rtrim(nombre),'') as ItmX" _
             & " from  cbr_cj_bufetes where activo = 1 and cod_bufete = rs!COD_BUFETE"
End If

Call OpenRecordSet(rs, strSQL)
If Not rs.EOF And Not rs.BOF Then
    txtBuffete.Text = rs!IdX
    txtBuffeteDesc.Text = rs!itmX
Else
    txtBuffete.Text = ""
    txtBuffeteDesc.Text = ""
End If

rs.Close

End Sub

Private Sub txtAbogado_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyF4 Then
    gBusquedas.Consulta = "select COD_ABOGADO, IDENTIFICACION, NOMBRE FROM CBR_CJ_ABOGADOS"
    gBusquedas.Orden = "COD_ABOGADO"
    gBusquedas.Columna = "COD_ABOGADO"
    gBusquedas.Filtro = " and ACTIVO = 1"
    frmBusquedas.Show vbModal
    txtAbogado.Text = gBusquedas.Resultado
    txtAbogadoDesc.Text = gBusquedas.Resultado3
End If

End Sub

Private Sub txtEmbargo_KeyDown(KeyCode As Integer, Shift As Integer)


If KeyCode = vbKeyF4 Then
    gBusquedas.Consulta = "select COD_EMBARGABLE,Descripcion from CBR_CJ_EMBARGABLES"
    gBusquedas.Orden = "COD_EMBARGABLE"
    gBusquedas.Columna = "COD_EMBARGABLE"
    gBusquedas.Filtro = " and ACTIVO = 1"
    frmBusquedas.Show vbModal
    txtEmbargo.Text = gBusquedas.Resultado
    txtEmbargoDesc.Text = gBusquedas.Resultado2
    
End If

End Sub

Private Sub txtExpediente_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyF4 Then
    gBusquedas.Consulta = "select EXPEDIENTE_NUMERO,cod_tramite from CBR_CJ_TRAMITE"
    gBusquedas.Orden = "EXPEDIENTE_NUMERO"
    gBusquedas.Columna = "EXPEDIENTE_NUMERO"
    gBusquedas.Filtro = " "
    frmBusquedas.Show vbModal
    txtExpediente.Text = gBusquedas.Resultado
    txtTramite.Text = gBusquedas.Resultado2
    If Trim(txtTramite.Text) <> "" Then Call sbConsulta
End If

End Sub

Private Sub txtOperacion_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyF4 Then
    gBusquedas.Col1Name = "No.OPERACION"
    gBusquedas.Consulta = "select  ID_SOLICITUD, CEDULA, NOMBRE,  CODIGO, DESCRIPCION  from vCbr_CJ_Operacion_Consulta"
    gBusquedas.Orden = "id_solicitud"
    gBusquedas.Columna = "id_solicitud"
    gBusquedas.Filtro = ""
    frmBusquedas.Show vbModal
    txtOperacion = gBusquedas.Resultado
    If Trim(txtOperacion.Text) <> "" Then Call sbDatosOperacion
    
End If
End Sub

Private Sub txtTramite_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then Call sbConsulta
End Sub


Private Sub txtTramite_LostFocus()
  Call sbConsulta
End Sub



Private Sub sbGuardar()

Dim vExiste As Integer
Dim strBufete As String
Dim iAplicaCierre As Integer
Dim i As Integer


On Error GoTo vError


'If Mid(txtEstado.Text, 1, 1) <> "P" Then
'    MsgBox "No se puede modificar este trámite porque no se encuentra pendiente", vbExclamation
'    Exit Sub
'End If


If txtBuffete.Text = "" Then
    strBufete = "Null"
Else
    strBufete = "'" & txtBuffete.Text & "'"
End If



If fxUltimoProceso(cboProceso.ItemData(cboProceso.ListIndex)) Then
  iAplicaCierre = 1
Else
  iAplicaCierre = 0
End If


If Not vEdita Then
   strSQL = "insert CBR_CJ_Tramite(CODIGO,ID_SOLICITUD,COD_ABOGADO,COD_JUZGADO,TIPO_JUICIO,COD_PROCESO,BASE_INT_COR," _
           & " BASE_INT_MOR,BASE_CARGOS,BASE_POLIZA,BASE_PRINICIPAL,TOTAL_DEUDA,TOTAL_GASTO,Cedula," _
           & " TOTAL_EJECUTADO,TOTAL_APLICADO,NOTAS,REGISTRO_FECHA,REGISTRO_USUARIO,PROCESO_FECHA,PROCESO_USUARIO,base_int_cor_venc,COD_BUFETE)" _
           & " VALUES('" & txtLineaCod.Text & "'," & txtOperacion.Text & "," & txtAbogado.Text & "," _
           & "'" & cboJuzgado.ItemData(cboJuzgado.ListIndex) & "','" & cboJuicio.ItemData(cboJuicio.ListIndex) & "','" & cboProceso.ItemData(cboProceso.ListIndex) & "'," _
           & " " & CCur(txtIntCorAtrasado.Text) & ", " & CCur(txtIntMoratorio.Text) & "," & CCur(txtCargos.Text) & "," _
           & " " & CCur(txtPolizas.Text) & "," & CCur(txtSaldo.Text) & "," & CCur(txtTotalDeuda.Text) & "," & CCur(txtTotalGastos.Text) & ",'" & Trim(txtCedula.Text) & "'," _
           & " " & CCur(txtTotalRecuperado.Text) & ", " & CCur(txtTotalAplicado.Text) & ",'" & txtNotas.Text & "',dbo.MyGetdate(),'" & glogon.Usuario & "'," _
           & " dbo.MyGetdate(),'" & glogon.Usuario & "'," & CCur(txtIntCorVenc.Text) & "," & strBufete & ")"
   Call ConectionExecute(strSQL)
   
   txtTramite.Text = fxUltimoTramite
   
  strSQL = "exec spCbr_CJ_Tramite_Proceso_Registra " & txtTramite.Text & ", '" & cboProceso.ItemData(cboProceso.ListIndex) _
         & "', " & iAplicaCierre & ", '" & Trim(txtNotas.Text) & "','" & glogon.Usuario & "'"
  Call ConectionExecute(strSQL)
  
  txtOperacion.SetFocus


Else




    'Valida Cambios
    If Not fxValidaCambios() Then
     If txtNotasCambios.Text = "" Then
        
        MsgBox vMensajeCambios, vbInformation
        
        fraNotas.Left = gbDatos.Left
        fraNotas.Visible = True
        
        Exit Sub
     End If
    End If

 
    'Control de Cambios
    strSQL = ""
    
    If vAbogado <> txtAbogado.Text Then
       strSQL = strSQL & Space(10) & "exec spCbr_CJ_Tramite_Cambios " & txtTramite.Text & ", 'Abodado'" _
               & ", '" & vAbogado & "','" & txtAbogado.Text & "','" & txtNotasCambios.Text & "','" & glogon.Usuario & "'"
    End If
    
    If vJuzgado <> cboJuzgado.ItemData(cboJuzgado.ListIndex) Then
       strSQL = strSQL & Space(10) & "exec spCbr_CJ_Tramite_Cambios " & txtTramite.Text & ", 'Juzgado'" _
               & ", '" & vJuzgado & "','" & cboJuzgado.ItemData(cboJuzgado.ListIndex) & "','" & txtNotasCambios.Text & "','" & glogon.Usuario & "'"
    End If
    
    If vJuicio <> cboJuicio.ItemData(cboJuicio.ListIndex) Then
       strSQL = strSQL & Space(10) & "exec spCbr_CJ_Tramite_Cambios " & txtTramite.Text & ", 'Juicio'" _
               & ", '" & vJuicio & "','" & cboJuicio.ItemData(cboJuicio.ListIndex) & "','" & txtNotasCambios.Text & "','" & glogon.Usuario & "'"
    End If
    
    If vBufete <> txtBuffete.Text Then
       strSQL = strSQL & Space(10) & "exec spCbr_CJ_Tramite_Cambios " & txtTramite.Text & ", 'Bufete'" _
               & ", '" & vBufete & "','" & txtBuffete.Text & "','" & txtNotasCambios.Text & "','" & glogon.Usuario & "'"
    
    End If



  strSQL = strSQL & Space(10) & "update CBR_CJ_Tramite set COD_ABOGADO = " & txtAbogado.Text _
         & ",COD_JUZGADO = '" & cboJuzgado.ItemData(cboJuzgado.ListIndex) & "' ,TIPO_JUICIO ='" & cboJuicio.ItemData(cboJuicio.ListIndex) _
         & "',COD_PROCESO = '" & cboProceso.ItemData(cboProceso.ListIndex) & "',COD_BUFETE = " & strBufete & ",notas = '" & txtNotas.Text & "'" _
         & " where cod_tramite = '" & txtTramite.Text & "'"
  Call ConectionExecute(strSQL)

 txtNotasCambios.Text = ""
End If

MsgBox "Registro Realizado Satisfactoriamente. [Tramite: " & txtTramite.Text & "]", vbInformation

Call sbToolBar(tlb, "activo")
Call sbConsulta

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical


End Sub




Private Sub tlbAux_ButtonClick(ByVal Button As MSComctlLib.Button)

On Error GoTo vError

GLOBALES.gTag = txtTramite.Text


Select Case Button.Key
 Case "Proceso"
    
    Call sbSIFForms("frmCO_CJ_Tramite_Proceso", 1, , , False, Me)
 Case "Gastos"
    Call sbSIFForms("frmCO_CJ_Tramite_Gastos", 1, , , False, Me)
    
 Case "Recuperacion"
 
 Case "Aplicado"
 
End Select

Call sbConsulta

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Function fxMinimoProceso() As String

On Error GoTo vError

strSQL = "select MIN(ORDEN) as orden from cbr_cj_proceso "

Call OpenRecordSet(rs, strSQL)
If Not rs.EOF Then
  fxMinimoProceso = IIf(IsNull(rs!Orden), "00", rs!Orden)
Else
 fxMinimoProceso = "00"
End If
rs.Close

Exit Function

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
 fxMinimoProceso = "00"

End Function

Private Sub sbDatosOperacion()

On Error GoTo vError

If txtTramite.Text = "" And txtOperacion.Text <> "" Then
    strSQL = "exec spCbr_CJ_Tramite_Datos_Operacion 0, " & txtOperacion.Text
Else
    strSQL = "exec spCbr_CJ_Tramite_Datos_Operacion " & txtTramite.Text & ", 0"
End If

Call OpenRecordSet(rs, strSQL)
 
If Not rs.EOF Then
    txtCedula.Text = Trim(rs!Cedula)
    txtNombre.Text = rs!Nombre
    txtLineaCod.Text = Trim(rs!Codigo)
    txtLineaDesc.Text = rs!Descripcion

    txtMonto.Text = Format(rs!montoapr, "Standard")
    txtSaldo.Text = Format(rs!Saldo, "Standard")
    txtPolizas.Text = Format(rs!Poliza, "Standard")
    txtCargos.Text = Format(rs!Cargos, "Standard")
    txtIntMoratorio.Text = Format(rs!INT_MOR, "Standard")
    txtIntCorAtrasado.Text = Format(rs!INT_COR, "Standard")
    If vEdita = False Then txtIntCorVenc.Text = Format(rs!InteresTotal - (rs!INT_COR + rs!INT_MOR), "Standard")
Else
    txtMonto.Text = 0
    txtSaldo.Text = 0
    txtPolizas.Text = 0
    txtCargos.Text = 0
    txtIntMoratorio.Text = 0
    txtIntCorAtrasado.Text = 0
    txtIntCorVenc.Text = 0
    
    txtCedula.Text = ""
    txtNombre.Text = ""
    txtLineaCod.Text = ""
    txtLineaDesc.Text = ""
End If
txtTotalDeuda = Format(CCur(txtSaldo.Text) + CCur(txtPolizas.Text) + CCur(txtCargos.Text) _
              + CCur(txtIntMoratorio.Text) + CCur(txtIntCorAtrasado.Text) + CCur(txtIntCorVenc.Text), "Standard")

rs.Close

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical


End Sub


Private Function fxUltimoTramite() As Long

With glogon

 .strSQL = "select isnull(max(cod_tramite),0) as 'Tramite' from cbr_cj_tramite Where Id_Solicitud = " & txtOperacion.Text

 Call OpenRecordSet(.Recordset, .strSQL)
 
    fxUltimoTramite = .Recordset!Tramite
    
 .Recordset.Close
End With

End Function


Private Function fxUltimoProceso(vCodigo As String) As Boolean

With glogon

.strSQL = "select max(ORDEN) as orden from cbr_cj_proceso "
 Call OpenRecordSet(.Recordset, .strSQL)

If Trim(.Recordset!Orden) = Trim(vCodigo) Then
    fxUltimoProceso = True
Else
  fxUltimoProceso = False
End If

 .Recordset.Close
End With


End Function


Private Sub vGrid_SheetChanged(ByVal OldSheet As Integer, ByVal NewSheet As Integer)

Select Case NewSheet
   Case 1 'Proceso
       vGrid.ActiveSheet = NewSheet
       vGrid.Sheet = 1
        strSQL = "select T.NUM_LINEA,case when T.APLICA_CIERRE_SENTENCIA =1" _
               & " then 'Sentencia Aplicada' else 'Sin Sentencia' end as 'Estado', P.DESCRIPCION" _
               & ", T.NOTAS , T.REGISTRO_USUARIO, T.REGISTRO_FECHA" _
               & " From CBR_CJ_TRAMITE_PROCESO T inner join CBR_CJ_PROCESO P on T.cod_Proceso = P.cod_Proceso" _
               & " where T.cod_tramite = " & txtTramite.Text & ""
        Call sbCargaGrid(vGrid, 6, strSQL)
    
    Case 2 'Gastos
        vGrid.ActiveSheet = 2
        vGrid.Sheet = 2

        strSQL = "SELECT NUM_LINEA,T.DESCRIPCION,G.MONTO,G.NOTAS,G.REGISTRO_USUARIO,G.REGISTRO_FECHA," _
              & " isnull(G.TESORERIA_NUMERO,0)as 'Numero' , isnull(G.COD_REMESA,0) as 'Remesa', isnull(G.TESORERIA_FECHA,'') as 'Fecha',G.BENEFICIARIO" _
              & " FROM CBR_CJ_TRAMITE_GASTOS G " _
              & " inner join CBR_CJ_TIPOS_GASTOS T on G.TIPO_GASTO = T.TIPO_GASTO" _
              & " Where G.COD_TRAMITE = " & txtTramite.Text & ""
        Call sbCargaGrid(vGrid, 10, strSQL)
    
    Case 3
        vGrid.ActiveSheet = 3
        vGrid.Sheet = 3

        strSQL = "SELECT LINEA,Cod_tramite,case  TIPO when  'A' then 'Abogado'" _
                & " when 'B' then 'Bufete' when 'J' then 'juzagado'" _
                & " else 'Tipo Juicio' end AS tipo ,COD_ACTUAL," _
                & " COD_ANTERIOR ,REGISTRO_USUARIO,REGISTRO_FECHA, NOTAS" _
                & " FROM CBR_CJ_TRAMITE_CAMBIOS " _
                & " Where COD_TRAMITE = " & txtTramite.Text & " order by linea"
        Call sbCargaGrid(vGrid, 8, strSQL)
      
End Select

End Sub

Public Sub sbConsultaExterna(xTramTemp As String, xOpTemp As String)
 txtTramite = xTramTemp
 If Trim(txtTramite) <> "" Then
    Call txtTramite_KeyDown(vbKeyReturn, 0)
 Else
   txtOperacion = xOpTemp
   vEdita = False
   txtOperacion.SetFocus
   Call sbToolBar(tlb, "edicion")
   Call sbDatosOperacion
 End If
 
End Sub




Private Sub sbEmbargos_List()
Dim itmX As ListViewItem


lsw.ListItems.Clear

If Not IsNumeric(txtTramite.Text) Then Exit Sub

strSQL = "select E.LINEA,E.COD_EMBARGABLE,Em.DESCRIPCION,E.MONTO,E.APLICA,E.NOTAS" _
        & " from CBR_CJ_TRAMITE_EMBARGABLES E inner join" _
        & " CBR_CJ_EMBARGABLES Em on E.COD_EMBARGABLE = Em.COD_EMBARGABLE" _
        & " where E.COD_TRAMITE = " & txtTramite.Text & ""
Call OpenRecordSet(rs, strSQL)

Do While Not rs.EOF
 Set itmX = lsw.ListItems.Add(, , rs!Linea)
     itmX.SubItems(1) = rs!COD_EMBARGABLE
     itmX.SubItems(2) = rs!Descripcion
     itmX.SubItems(3) = Format(rs!Monto, "Standard")
     If rs!Aplica = 1 Then
        itmX.SubItems(4) = "Aplica Embargo"
     Else
        itmX.SubItems(4) = "No aplica Embargo"
     End If
     itmX.SubItems(5) = Trim(rs!notas)
     
 rs.MoveNext
Loop
rs.Close

End Sub



Private Function fxValidaCambios() As Boolean

fxValidaCambios = True
vMensajeCambios = ""


If vAbogado <> txtAbogado.Text Then
    vMensajeCambios = vMensajeCambios & vbCrLf & "Cambio de Abogado"
End If


If vJuzgado <> cboJuzgado.ItemData(cboJuzgado.ListIndex) Then
   vMensajeCambios = vMensajeCambios & vbCrLf & "Cambio de Juzgado"
End If

If vJuicio <> cboJuicio.ItemData(cboJuicio.ListIndex) Then
   vMensajeCambios = vMensajeCambios & vbCrLf & "Cambio de tipo Juicio"
End If

If vBufete <> txtBuffete.Text Then
   vMensajeCambios = vMensajeCambios & vbCrLf & "Cambio de Bufete"
End If

If Len(vMensajeCambios) > 0 Then
  fxValidaCambios = False
End If

End Function


