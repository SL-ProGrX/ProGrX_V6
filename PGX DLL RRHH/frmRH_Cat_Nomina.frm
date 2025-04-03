VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.controls.v22.1.0.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Begin VB.Form frmRH_Cat_Nomina 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "RRHH: Definición de la Nómina"
   ClientHeight    =   8385
   ClientLeft      =   30
   ClientTop       =   390
   ClientWidth     =   9870
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8385
   ScaleWidth      =   9870
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer TimerX 
      Interval        =   5
      Left            =   6600
      Top             =   600
   End
   Begin XtremeSuiteControls.CheckBox chkActivo 
      Height          =   252
      Left            =   8160
      TabIndex        =   0
      Top             =   600
      Width           =   1092
      _Version        =   1441793
      _ExtentX        =   1926
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Activo?"
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
   Begin XtremeSuiteControls.TabControl tcMain 
      Height          =   7092
      Left            =   120
      TabIndex        =   1
      Top             =   1200
      Width           =   9612
      _Version        =   1441793
      _ExtentX        =   16954
      _ExtentY        =   12509
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
      ItemCount       =   5
      Item(0).Caption =   "Nómina"
      Item(0).ControlCount=   18
      Item(0).Control(0)=   "Label1(1)"
      Item(0).Control(1)=   "txtDescripcion"
      Item(0).Control(2)=   "Label1(2)"
      Item(0).Control(3)=   "txtConsecutivo"
      Item(0).Control(4)=   "cboTipoNomina"
      Item(0).Control(5)=   "Label1(7)"
      Item(0).Control(6)=   "cboFrecuencia"
      Item(0).Control(7)=   "Label1(8)"
      Item(0).Control(8)=   "gbAdelanto"
      Item(0).Control(9)=   "GroupBox1"
      Item(0).Control(10)=   "GroupBox2"
      Item(0).Control(11)=   "cboDivisa"
      Item(0).Control(12)=   "Label1(4)"
      Item(0).Control(13)=   "Label1(5)"
      Item(0).Control(14)=   "cboDias"
      Item(0).Control(15)=   "chkSM_Aplica"
      Item(0).Control(16)=   "txtSM_Monto"
      Item(0).Control(17)=   "Label1(11)"
      Item(1).Caption =   "Conceptos"
      Item(1).ControlCount=   2
      Item(1).Control(0)=   "Label1(3)"
      Item(1).Control(1)=   "lsw"
      Item(2).Caption =   "Jornadas"
      Item(2).ControlCount=   2
      Item(2).Control(0)=   "Label1(9)"
      Item(2).Control(1)=   "lswJornadas"
      Item(3).Caption =   "Contratos"
      Item(3).ControlCount=   2
      Item(3).Control(0)=   "Label1(10)"
      Item(3).Control(1)=   "lswContratos"
      Item(4).Caption =   "Prorrateo"
      Item(4).ControlCount=   3
      Item(4).Control(0)=   "chkProrrateo"
      Item(4).Control(1)=   "gProrrateo"
      Item(4).Control(2)=   "txtProrrateo_Sum"
      Begin XtremeSuiteControls.ListView lswContratos 
         Height          =   5892
         Left            =   -68200
         TabIndex        =   34
         Top             =   480
         Visible         =   0   'False
         Width           =   7812
         _Version        =   1441793
         _ExtentX        =   13779
         _ExtentY        =   10393
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
      Begin XtremeSuiteControls.ListView lswJornadas 
         Height          =   5892
         Left            =   -68200
         TabIndex        =   32
         Top             =   480
         Visible         =   0   'False
         Width           =   7812
         _Version        =   1441793
         _ExtentX        =   13779
         _ExtentY        =   10393
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
      Begin XtremeSuiteControls.ListView lsw 
         Height          =   5892
         Left            =   -68200
         TabIndex        =   2
         Top             =   480
         Visible         =   0   'False
         Width           =   7812
         _Version        =   1441793
         _ExtentX        =   13779
         _ExtentY        =   10393
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
      Begin XtremeSuiteControls.GroupBox gbAdelanto 
         Height          =   1092
         Left            =   1800
         TabIndex        =   16
         Top             =   3000
         Width           =   7572
         _Version        =   1441793
         _ExtentX        =   13356
         _ExtentY        =   1926
         _StockProps     =   79
         BackColor       =   -2147483633
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         BorderStyle     =   1
         Begin XtremeSuiteControls.CheckBox chkAguinaldo 
            Height          =   372
            Left            =   0
            TabIndex        =   17
            Top             =   120
            Width           =   3012
            _Version        =   1441793
            _ExtentX        =   5313
            _ExtentY        =   656
            _StockProps     =   79
            Caption         =   "Registra Aguinaldo?"
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
            UseVisualStyle  =   -1  'True
            Appearance      =   16
         End
         Begin XtremeSuiteControls.FlatEdit txtAguinaldoCod 
            Height          =   312
            Left            =   0
            TabIndex        =   18
            ToolTipText     =   "Presione F4 para Consultar"
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
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtAguinaldoDesc 
            Height          =   312
            Left            =   960
            TabIndex        =   19
            Top             =   720
            Width           =   6492
            _Version        =   1441793
            _ExtentX        =   11451
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
      End
      Begin XtremeSuiteControls.FlatEdit txtDescripcion 
         Height          =   312
         Left            =   1800
         TabIndex        =   3
         Top             =   600
         Width           =   7452
         _Version        =   1441793
         _ExtentX        =   13144
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.ComboBox cboTipoNomina 
         Height          =   312
         Left            =   1800
         TabIndex        =   12
         Top             =   960
         Width           =   3132
         _Version        =   1441793
         _ExtentX        =   5530
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
      Begin XtremeSuiteControls.ComboBox cboFrecuencia 
         Height          =   312
         Left            =   1800
         TabIndex        =   14
         Top             =   1320
         Width           =   3132
         _Version        =   1441793
         _ExtentX        =   5530
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
      Begin XtremeSuiteControls.GroupBox GroupBox1 
         Height          =   1092
         Left            =   1800
         TabIndex        =   20
         Top             =   4200
         Width           =   7572
         _Version        =   1441793
         _ExtentX        =   13356
         _ExtentY        =   1926
         _StockProps     =   79
         BackColor       =   -2147483633
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         BorderStyle     =   1
         Begin XtremeSuiteControls.CheckBox chkAdelanto 
            Height          =   372
            Left            =   0
            TabIndex        =   21
            Top             =   120
            Width           =   3012
            _Version        =   1441793
            _ExtentX        =   5313
            _ExtentY        =   656
            _StockProps     =   79
            Caption         =   "Puede Realizar Adelantos de Salario?"
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
            UseVisualStyle  =   -1  'True
            Appearance      =   16
         End
         Begin XtremeSuiteControls.FlatEdit txtAdelantoCod 
            Height          =   312
            Left            =   0
            TabIndex        =   22
            ToolTipText     =   "Presione F4 para Consultar"
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
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtAdelantoDesc 
            Height          =   312
            Left            =   960
            TabIndex        =   23
            Top             =   720
            Width           =   6492
            _Version        =   1441793
            _ExtentX        =   11451
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
         Begin XtremeSuiteControls.FlatEdit txtAdelantoPorc 
            Height          =   315
            Left            =   6600
            TabIndex        =   48
            ToolTipText     =   "Salario Mínimo"
            Top             =   360
            Width           =   855
            _Version        =   1441793
            _ExtentX        =   1508
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
         Begin XtremeSuiteControls.Label Label1 
            Height          =   255
            Index           =   14
            Left            =   4440
            TabIndex        =   49
            Top             =   360
            Width           =   1935
            _Version        =   1441793
            _ExtentX        =   3413
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Porcentaje Máximo"
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
            Alignment       =   5
            Transparent     =   -1  'True
            WordWrap        =   -1  'True
         End
      End
      Begin XtremeSuiteControls.GroupBox GroupBox2 
         Height          =   2652
         Left            =   240
         TabIndex        =   24
         Top             =   5400
         Width           =   9132
         _Version        =   1441793
         _ExtentX        =   16108
         _ExtentY        =   4678
         _StockProps     =   79
         Caption         =   "Cuentas Auxiliares:"
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
         BorderStyle     =   1
         Begin XtremeSuiteControls.FlatEdit txtCuentaPreDesc 
            Height          =   312
            Left            =   3480
            TabIndex        =   46
            Top             =   1080
            Width           =   5532
            _Version        =   1441793
            _ExtentX        =   9758
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
         Begin XtremeSuiteControls.FlatEdit txtCuentaLiqDesc 
            Height          =   312
            Left            =   3480
            TabIndex        =   43
            Top             =   720
            Width           =   5532
            _Version        =   1441793
            _ExtentX        =   9758
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
         Begin XtremeSuiteControls.FlatEdit txtCuentaCod 
            Height          =   312
            Left            =   1560
            TabIndex        =   25
            ToolTipText     =   "Presione F4 para Consultar"
            Top             =   360
            Width           =   1932
            _Version        =   1441793
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
            Alignment       =   2
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtCuentaDesc 
            Height          =   312
            Left            =   3480
            TabIndex        =   26
            Top             =   360
            Width           =   5532
            _Version        =   1441793
            _ExtentX        =   9758
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
         Begin XtremeSuiteControls.FlatEdit txtCuentaLiq 
            Height          =   312
            Left            =   1560
            TabIndex        =   42
            ToolTipText     =   "Presione F4 para Consultar"
            Top             =   720
            Width           =   1932
            _Version        =   1441793
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
            Alignment       =   2
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtCuentaPre 
            Height          =   312
            Left            =   1560
            TabIndex        =   45
            ToolTipText     =   "Presione F4 para Consultar"
            Top             =   1080
            Width           =   1932
            _Version        =   1441793
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
            Alignment       =   2
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.Label Label1 
            Height          =   372
            Index           =   13
            Left            =   0
            TabIndex        =   47
            Top             =   1080
            Width           =   1692
            _Version        =   1441793
            _ExtentX        =   2984
            _ExtentY        =   656
            _StockProps     =   79
            Caption         =   "Preaviso"
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
            Alignment       =   4
            Transparent     =   -1  'True
            WordWrap        =   -1  'True
         End
         Begin XtremeSuiteControls.Label Label1 
            Height          =   372
            Index           =   12
            Left            =   0
            TabIndex        =   44
            Top             =   720
            Width           =   1692
            _Version        =   1441793
            _ExtentX        =   2984
            _ExtentY        =   656
            _StockProps     =   79
            Caption         =   "Pago Liquidación"
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
            Alignment       =   4
            Transparent     =   -1  'True
            WordWrap        =   -1  'True
         End
         Begin XtremeSuiteControls.Label Label1 
            Height          =   372
            Index           =   6
            Left            =   0
            TabIndex        =   27
            Top             =   360
            Width           =   1692
            _Version        =   1441793
            _ExtentX        =   2984
            _ExtentY        =   656
            _StockProps     =   79
            Caption         =   "Pago Planilla"
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
            Alignment       =   4
            Transparent     =   -1  'True
            WordWrap        =   -1  'True
         End
      End
      Begin XtremeSuiteControls.ComboBox cboDivisa 
         Height          =   312
         Left            =   6960
         TabIndex        =   28
         Top             =   960
         Width           =   2292
         _Version        =   1441793
         _ExtentX        =   4048
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
      Begin XtremeSuiteControls.FlatEdit txtConsecutivo 
         Height          =   312
         Left            =   6960
         TabIndex        =   4
         Top             =   1320
         Width           =   2292
         _Version        =   1441793
         _ExtentX        =   4043
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
      Begin XtremeSuiteControls.ComboBox cboDias 
         Height          =   312
         Left            =   6960
         TabIndex        =   31
         Top             =   2040
         Width           =   2292
         _Version        =   1441793
         _ExtentX        =   4048
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
      Begin XtremeSuiteControls.CheckBox chkSM_Aplica 
         Height          =   372
         Left            =   1800
         TabIndex        =   36
         Top             =   2040
         Width           =   3132
         _Version        =   1441793
         _ExtentX        =   5524
         _ExtentY        =   656
         _StockProps     =   79
         Caption         =   "Aplica Protección al Salario Mínimo?"
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
         UseVisualStyle  =   -1  'True
         Appearance      =   16
         Alignment       =   1
      End
      Begin XtremeSuiteControls.FlatEdit txtSM_Monto 
         Height          =   312
         Left            =   3240
         TabIndex        =   37
         ToolTipText     =   "Salario Mínimo"
         Top             =   2520
         Width           =   1692
         _Version        =   1441793
         _ExtentX        =   2984
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.CheckBox chkProrrateo 
         Height          =   1092
         Left            =   -69880
         TabIndex        =   39
         Top             =   360
         Visible         =   0   'False
         Width           =   1572
         _Version        =   1441793
         _ExtentX        =   2773
         _ExtentY        =   1926
         _StockProps     =   79
         Caption         =   "Activar Prorrateo del Gasto en Salarios?"
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
         Appearance      =   16
         Alignment       =   1
      End
      Begin FPSpreadADO.fpSpread gProrrateo 
         Height          =   5172
         Left            =   -67840
         TabIndex        =   40
         Top             =   480
         Visible         =   0   'False
         Width           =   7212
         _Version        =   524288
         _ExtentX        =   12721
         _ExtentY        =   9123
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
         MaxCols         =   483
         ScrollBars      =   2
         SpreadDesigner  =   "frmRH_Cat_Nomina.frx":0000
         VScrollSpecial  =   -1  'True
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin XtremeSuiteControls.FlatEdit txtProrrateo_Sum 
         Height          =   312
         Left            =   -62680
         TabIndex        =   41
         ToolTipText     =   "Salario Mínimo"
         Top             =   5880
         Visible         =   0   'False
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
         Alignment       =   1
         Locked          =   -1  'True
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.Label Label1 
         Height          =   252
         Index           =   11
         Left            =   1800
         TabIndex        =   38
         Top             =   2520
         Width           =   1452
         _Version        =   1441793
         _ExtentX        =   2561
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Salario Mínimo"
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
         Alignment       =   4
         Transparent     =   -1  'True
         WordWrap        =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label1 
         Height          =   1212
         Index           =   10
         Left            =   -69880
         TabIndex        =   35
         Top             =   480
         Visible         =   0   'False
         Width           =   1332
         _Version        =   1441793
         _ExtentX        =   2350
         _ExtentY        =   2138
         _StockProps     =   79
         Caption         =   "Seleccione los Tipos de Contratos admitidos en esta nómina:"
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
         Alignment       =   4
         Transparent     =   -1  'True
         WordWrap        =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label1 
         Height          =   1332
         Index           =   9
         Left            =   -69880
         TabIndex        =   33
         Top             =   480
         Visible         =   0   'False
         Width           =   1332
         _Version        =   1441793
         _ExtentX        =   2350
         _ExtentY        =   2350
         _StockProps     =   79
         Caption         =   "Seleccione los Tipos de Jornadas Laborales admitidas en esta nómina:"
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
         Alignment       =   4
         Transparent     =   -1  'True
         WordWrap        =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label1 
         Height          =   732
         Index           =   5
         Left            =   5400
         TabIndex        =   30
         Top             =   1920
         Width           =   1332
         _Version        =   1441793
         _ExtentX        =   2350
         _ExtentY        =   1291
         _StockProps     =   79
         Caption         =   "Días de Cálculo para Liquidaciónes"
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
         Alignment       =   4
         Transparent     =   -1  'True
         WordWrap        =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label1 
         Height          =   372
         Index           =   4
         Left            =   5400
         TabIndex        =   29
         Top             =   960
         Width           =   1332
         _Version        =   1441793
         _ExtentX        =   2350
         _ExtentY        =   656
         _StockProps     =   79
         Caption         =   "Divisa"
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
         Alignment       =   4
         Transparent     =   -1  'True
         WordWrap        =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label1 
         Height          =   372
         Index           =   8
         Left            =   360
         TabIndex        =   15
         Top             =   1320
         Width           =   1332
         _Version        =   1441793
         _ExtentX        =   2350
         _ExtentY        =   656
         _StockProps     =   79
         Caption         =   "Frec. Pago"
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
         Alignment       =   4
         Transparent     =   -1  'True
         WordWrap        =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label1 
         Height          =   372
         Index           =   7
         Left            =   360
         TabIndex        =   13
         Top             =   960
         Width           =   1332
         _Version        =   1441793
         _ExtentX        =   2350
         _ExtentY        =   656
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
         Alignment       =   4
         Transparent     =   -1  'True
         WordWrap        =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label1 
         Height          =   252
         Index           =   1
         Left            =   360
         TabIndex        =   7
         Top             =   600
         Width           =   1332
         _Version        =   1441793
         _ExtentX        =   2350
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Descripción"
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
         Alignment       =   4
         Transparent     =   -1  'True
         WordWrap        =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label1 
         Height          =   252
         Index           =   2
         Left            =   5400
         TabIndex        =   6
         Top             =   1320
         Width           =   1572
         _Version        =   1441793
         _ExtentX        =   2773
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Consecutivo"
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
         Alignment       =   4
         Transparent     =   -1  'True
         WordWrap        =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label1 
         Height          =   1212
         Index           =   3
         Left            =   -69880
         TabIndex        =   5
         Top             =   480
         Visible         =   0   'False
         Width           =   1332
         _Version        =   1441793
         _ExtentX        =   2350
         _ExtentY        =   2138
         _StockProps     =   79
         Caption         =   "Seleccione los Conceptos admitidos en esta nómina:"
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
         Alignment       =   4
         Transparent     =   -1  'True
         WordWrap        =   -1  'True
      End
   End
   Begin MSComctlLib.Toolbar tlb 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   9870
      _ExtentX        =   17410
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
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
   Begin MSComCtl2.FlatScrollBar FlatScrollBar 
      Height          =   252
      Left            =   3480
      TabIndex        =   9
      Top             =   600
      Width           =   492
      _ExtentX        =   873
      _ExtentY        =   450
      _Version        =   393216
      Arrows          =   65536
      Orientation     =   1638401
   End
   Begin XtremeSuiteControls.FlatEdit txtCodigo 
      Height          =   312
      Left            =   1920
      TabIndex        =   10
      Top             =   600
      Width           =   1452
      _Version        =   1441793
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
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.Label Label1 
      Height          =   252
      Index           =   0
      Left            =   480
      TabIndex        =   11
      Top             =   600
      Width           =   1212
      _Version        =   1441793
      _ExtentX        =   2138
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Nómina"
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
      Alignment       =   1
      Transparent     =   -1  'True
   End
End
Attribute VB_Name = "frmRH_Cat_Nomina"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vScroll As Boolean
Dim vEdita  As Boolean
Dim vCodigo As String, vPaso As Boolean
Dim rs As New ADODB.Recordset, strSQL As String
Dim itmX As ListViewItem



Private Function fxExiste(vCodigo As String) As Boolean

strSQL = "select isnull(count(*),0) as 'Existe'" _
       & " from RH_NOMINAS_CATALOGO where COD_NOMINA =  '" & vCodigo & "' "
Call OpenRecordSet(rs, strSQL)
If rs!Existe = 0 Then
  fxExiste = False
Else
  fxExiste = True
End If
rs.Close
End Function


Private Sub chkProrrateo_Click()

On Error GoTo vError

Me.MousePointer = vbHourglass

  strSQL = "update RH_NOMINAS_CATALOGO set PRORRATEO_APL = " & chkProrrateo.Value _
         & " where COD_NOMINA = '" & vCodigo & "' "
         
  Call ConectionExecute(strSQL)
  Call Bitacora("Modifica", "Catálogo de Nómina: " & vCodigo & ", Prorrateo: " & IIf((chkProrrateo.Value = xtpChecked), "Sí", "No"))

Me.MousePointer = vbDefault

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub FlatScrollBar_Change()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

If vScroll Then

    strSQL = "select Top 1 COD_NOMINA from RH_NOMINAS_CATALOGO"
    
    If FlatScrollBar.Value = 1 Then
       strSQL = strSQL & " where COD_NOMINA > '" & txtCodigo.Text & "' order by COD_NOMINA asc"
    Else
       strSQL = strSQL & " where COD_NOMINA < '" & txtCodigo.Text & "' order by COD_NOMINA desc"
    End If
    
    Call OpenRecordSet(rs, strSQL)
    If Not rs.EOF And Not rs.BOF Then
      txtCodigo.Text = rs!COD_NOMINA
      Call sbConsulta(txtCodigo.Text)
      
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
vModulo = 23
End Sub


Private Sub Form_Load()
vModulo = 23
 
vEdita = True

With lsw.ColumnHeaders
    .Clear
    .Add , , "Id", 950
    .Add , , "Descripción", 6750
End With


With lswContratos.ColumnHeaders
    .Clear
    .Add , , "Id", 950
    .Add , , "Descripción", 6750
End With

With lswJornadas.ColumnHeaders
    .Clear
    .Add , , "Id", 950
    .Add , , "Descripción", 6750
End With


Call sbToolBarIconos(tlb, False)
Call sbToolBar(tlb, "nuevo")

vScroll = False
FlatScrollBar.Value = 0
vScroll = True

Call sbLimpiaDatos

'Pasa Seguridad al Timer
End Sub


Private Sub sbLimpiaDatos()

vCodigo = ""

tcMain.Item(0).Selected = True

txtCodigo.Text = ""
txtDescripcion.Text = ""
txtConsecutivo.Text = "0"

chkAdelanto.Value = xtpChecked
txtAdelantoCod.Text = ""
txtAdelantoDesc.Text = ""

chkAguinaldo.Value = xtpChecked
txtAguinaldoCod.Text = ""
txtAguinaldoDesc.Text = ""

txtCuentaCod.Text = ""
txtCuentaDesc.Text = ""

txtCuentaPre.Text = ""
txtCuentaPreDesc.Text = ""

txtCuentaLiq.Text = ""
txtCuentaLiqDesc.Text = ""

txtSM_Monto.Text = "0"
chkSM_Aplica.Value = xtpUnchecked

cboDias.Text = "30"

chkActivo.Value = xtpChecked

chkProrrateo.Value = xtpUnchecked


txtAdelantoPorc.Text = "0"

End Sub


Private Sub sbProrrateo_Suma()
Dim x As Long, pTotal As Double
  
On Error GoTo vError
  
With gProrrateo
  
  pTotal = 0
  
  For x = 1 To .MaxRows
      .Row = x
      .col = 3
      If IsNumeric(.Text) Then
          pTotal = pTotal + CDbl(.Text)
      End If
  Next x
  
  txtProrrateo_Sum.Text = Format(pTotal, "###,##0.000000000")

End With

vError:

End Sub

Private Sub gProrrateo_KeyUp(KeyCode As Integer, Shift As Integer)

With gProrrateo

If .ActiveCol = .MaxCols Then
  Call sbProrrateo_Suma
End If

End With

End Sub

Private Sub gProrrateo_KeyDown(KeyCode As Integer, Shift As Integer)
Dim i As Integer

With gProrrateo


If .ActiveCol = .MaxCols And (KeyCode = vbKeyReturn Or KeyCode = vbKeyTab) Then
  i = fxProrrateo_Guardar
  
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


'Consulta Centro de Costo
If KeyCode = vbKeyF4 And .ActiveCol = 1 Then
  .Row = .ActiveRow
'  .Col = 2
'  vTempo = .Text
  
  gBusquedas.Resultado = ""
  gBusquedas.Resultado2 = ""
  gBusquedas.Columna = "descripcion"
  gBusquedas.Orden = "descripcion"
  gBusquedas.Consulta = "select cod_centro_costo,descripcion from CntX_Centro_Costos"
  gBusquedas.Filtro = " and cod_contabilidad = " & GLOBALES.gEnlace
'  & " and cod_centro_costo in(select cod_centro_costo" _
'                    & " from cntX_unidades_cc where cod_unidad = '" & vTempo & "' and cod_contabilidad = " & gCntX_Parametros.CodigoConta & ")"
  frmBusquedas.Show vbModal
    
  .Row = .ActiveRow
  .col = 1
  
  .Text = gBusquedas.Resultado
  
  .col = 2
  .Text = gBusquedas.Resultado2
  
End If

If (KeyCode = vbKeyReturn Or KeyCode = vbKeyTab) And .ActiveCol = 1 Then
    Dim pTemp As String
    .Row = .ActiveRow
    .col = 1
    pTemp = .Text
    
    .col = 2
    
    .Text = fxCntX_CentroCosto("D", pTemp)
End If

'Borrar una linea
If KeyCode = vbKeyDelete Then
     i = MsgBox("Esta Seguro que desea borrar este registro", vbYesNo)
     If i = vbYes Then
        
        .Row = .ActiveRow
        .col = 1
        If .Text <> "" Then
            strSQL = "delete RH_NOMINAS_PORRATEO_CC where COD_CENTRO_COSTO = '" _
                    & .Text & "' AND COD_NOMINA = '" & vCodigo & "'"
            Call ConectionExecute(strSQL)
            strSQL = .Text
            .col = 1
            Call Bitacora("Elimina", "Nómina: " & vCodigo & ", Prorrateo Cc: " & .Text)
            
            Call sbProrrateo_Load(vCodigo)
            
        End If
     End If
End If

End With


End Sub




Private Sub lsw_ColumnClick(ByVal ColumnHeader As XtremeSuiteControls.ListViewColumnHeader)
 lsw.SortKey = ColumnHeader.Index - 1
  If lsw.SortOrder = 0 Then lsw.SortOrder = 1 Else lsw.SortOrder = 0
  lsw.Sorted = True
End Sub

Private Sub lsw_ItemCheck(ByVal Item As XtremeSuiteControls.ListViewItem)

On Error GoTo vError

If vPaso Then Exit Sub

If Item.Checked Then
   strSQL = "insert RH_NOMINAS_CATALOGO_CONCEPTOS(COD_CONCEPTO,COD_NOMINA,registro_fecha,registro_usuario)" _
          & " values('" & Item.Text & "','" & vCodigo & "',dbo.MyGetdate(),'" & glogon.Usuario & "')"
Else
   strSQL = "delete RH_NOMINAS_CATALOGO_CONCEPTOS where COD_CONCEPTO = '" & Item.Text & "' and COD_NOMINA = '" & vCodigo & "'"
End If

Call ConectionExecute(strSQL)

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

'--------
Private Sub lswContratos_ColumnClick(ByVal ColumnHeader As XtremeSuiteControls.ListViewColumnHeader)
 lswContratos.SortKey = ColumnHeader.Index - 1
  If lswContratos.SortOrder = 0 Then lswContratos.SortOrder = 1 Else lswContratos.SortOrder = 0
  lswContratos.Sorted = True
End Sub

Private Sub lswContratos_ItemCheck(ByVal Item As XtremeSuiteControls.ListViewItem)

On Error GoTo vError

If vPaso Then Exit Sub

If Item.Checked Then
   strSQL = "insert RH_NOMINAS_CATALOGO_CONTRATOS(CONTRATO_TIPO,COD_NOMINA,registro_fecha,registro_usuario)" _
          & " values('" & Item.Text & "','" & vCodigo & "',dbo.MyGetdate(),'" & glogon.Usuario & "')"
Else
   strSQL = "delete RH_NOMINAS_CATALOGO_CONTRATOS where CONTRATO_TIPO = '" & Item.Text & "' and COD_NOMINA = '" & vCodigo & "'"
End If

Call ConectionExecute(strSQL)

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

'------
Private Sub lswJornadas_ColumnClick(ByVal ColumnHeader As XtremeSuiteControls.ListViewColumnHeader)
 lswJornadas.SortKey = ColumnHeader.Index - 1
  If lswJornadas.SortOrder = 0 Then lswJornadas.SortOrder = 1 Else lswJornadas.SortOrder = 0
  lswJornadas.Sorted = True
End Sub

Private Sub lswJornadas_ItemCheck(ByVal Item As XtremeSuiteControls.ListViewItem)
Dim strSQL As String

On Error GoTo vError

If vPaso Then Exit Sub

If Item.Checked Then
   strSQL = "insert RH_NOMINAS_CATALOGO_JORNADAS(JORNADA_TIPO,COD_NOMINA,registro_fecha,registro_usuario)" _
          & " values('" & Item.Text & "','" & vCodigo & "',dbo.MyGetdate(),'" & glogon.Usuario & "')"
Else
   strSQL = "delete RH_NOMINAS_CATALOGO_JORNADAS where JORNADA_TIPO = '" & Item.Text & "' and COD_NOMINA = '" & vCodigo & "'"
End If

Call ConectionExecute(strSQL)

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub tcMain_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)

If vCodigo = "" Then
    tcMain.Item(0).Selected = True
End If

Select Case Item.Index
 Case 1 'Conceptos
    Call sbConceptos_Load(txtCodigo.Text)
 Case 2 'Jornadas
    Call sbJornadas_Load(txtCodigo.Text)
 Case 3 'Contratos
    Call sbContratos_Load(txtCodigo.Text)
 Case 4 'Prorrateo
    Call sbProrrateo_Load(txtCodigo.Text)
End Select

End Sub

Private Sub TimerX_Timer()
TimerX.Interval = 0
TimerX.Enabled = False

Dim strSQL As String

cboDias.AddItem "30"
cboDias.AddItem "26"
cboDias.Text = "30"

strSQL = "select rtrim(COD_DIVISA) AS 'idX', DESCRIPCION as 'itmX'" _
       & " From vSys_Divisas"
Call sbCbo_Llena_New(cboDivisa, strSQL, False, True)

strSQL = "select rtrim(NOMINA_TIPO) AS 'idX', DESCRIPCION as 'itmX'" _
       & " from RH_NOMINAS_TIPOS where ACTIVO = 1"
Call sbCbo_Llena_New(cboTipoNomina, strSQL, False, True)

strSQL = "select rtrim(COD_FRECPAGO) AS 'idX', DESCRIPCION as 'itmX'" _
       & " from RH_PAGO_FRECUENCIA where ACTIVO = 1"
Call sbCbo_Llena_New(cboFrecuencia, strSQL, False, True)

Call Formularios(Me)
Call RefrescaTags(Me)


End Sub




Private Sub txtAdelantoCod_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then
   gBusquedas.Col1Name = "Concepto Id"
   gBusquedas.Col2Name = "Descripción"
   gBusquedas.Columna = "COD_CONCEPTO"
   gBusquedas.Orden = "COD_CONCEPTO"
   gBusquedas.Consulta = "select COD_CONCEPTO,descripcion from RH_CONCEPTOS"
   frmBusquedas.Show vbModal
   txtAdelantoCod.Text = gBusquedas.Resultado
   txtAdelantoDesc.Text = gBusquedas.Resultado2
End If

End Sub

Private Sub txtAguinaldoCod_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then
   gBusquedas.Col1Name = "Concepto Id"
   gBusquedas.Col2Name = "Descripción"
   gBusquedas.Columna = "COD_CONCEPTO"
   gBusquedas.Orden = "COD_CONCEPTO"
   gBusquedas.Consulta = "select COD_CONCEPTO,descripcion from RH_CONCEPTOS"
   frmBusquedas.Show vbModal
   txtAguinaldoCod.Text = gBusquedas.Resultado
   txtAguinaldoDesc.Text = gBusquedas.Resultado2
End If
End Sub

Private Sub txtCodigo_LostFocus()
If Trim(txtCodigo) <> "" And vEdita = True Then Call sbConsulta(txtCodigo.Text)
End Sub



Private Sub sbConceptos_Load(pCodigo As String)

On Error GoTo vError

vPaso = True

lsw.ListItems.Clear

strSQL = "select R.COD_CONCEPTO AS 'CODIGO', R.DESCRIPCION,ISNULL(A.COD_NOMINA,'') AS 'Idx'" _
       & "  from RH_CONCEPTOS R" _
       & "   LEFT JOIN RH_NOMINAS_CATALOGO_CONCEPTOS A ON R.COD_CONCEPTO = A.COD_CONCEPTO" _
       & "   AND A.COD_NOMINA = '" & pCodigo & "'" _
       & " WHERE R.ACTIVO = 1" _
       & " order by ISNULL(A.COD_NOMINA,'') desc, R.COD_CONCEPTO"
       
Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
  Set itmX = lsw.ListItems.Add(, , rs!Codigo)
      itmX.SubItems(1) = rs!Descripcion
      If rs!IdX <> "" Then
          itmX.Checked = vbChecked
          itmX.ForeColor = vbBlue
      End If
  rs.MoveNext
Loop
rs.Close

vPaso = False

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub sbJornadas_Load(pCodigo As String)

On Error GoTo vError

vPaso = True

lswJornadas.ListItems.Clear

strSQL = "select R.JORNADA_TIPO AS 'CODIGO', R.DESCRIPCION,ISNULL(A.COD_NOMINA,'') AS 'Idx'" _
       & "  from RH_JORNADAS_TIPOS R" _
       & "   LEFT JOIN RH_NOMINAS_CATALOGO_JORNADAS A ON R.JORNADA_TIPO = A.JORNADA_TIPO" _
       & "   AND A.COD_NOMINA = '" & pCodigo & "'" _
       & " WHERE R.ACTIVO = 1" _
       & " order by ISNULL(A.COD_NOMINA,'') desc, R.JORNADA_TIPO"
       
Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
  Set itmX = lswJornadas.ListItems.Add(, , rs!Codigo)
      itmX.SubItems(1) = rs!Descripcion
      If rs!IdX <> "" Then
          itmX.Checked = vbChecked
          itmX.ForeColor = vbBlue
      End If
  rs.MoveNext
Loop
rs.Close

vPaso = False

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub sbContratos_Load(pCodigo As String)

On Error GoTo vError

vPaso = True

lswContratos.ListItems.Clear

strSQL = "select R.CONTRATO_TIPO AS 'CODIGO', R.DESCRIPCION,ISNULL(A.COD_NOMINA,'') AS 'Idx'" _
       & "  from RH_CONTRATOS_TIPOS R" _
       & "   LEFT JOIN RH_NOMINAS_CATALOGO_CONTRATOS A ON R.CONTRATO_TIPO = A.CONTRATO_TIPO" _
       & "   AND A.COD_NOMINA = '" & pCodigo & "'" _
       & " WHERE R.ACTIVO = 1" _
       & " order by ISNULL(A.COD_NOMINA,'') desc, R.CONTRATO_TIPO"
       
Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
  Set itmX = lswContratos.ListItems.Add(, , rs!Codigo)
      itmX.SubItems(1) = rs!Descripcion
      If rs!IdX <> "" Then
          itmX.Checked = vbChecked
          itmX.ForeColor = vbBlue
      End If
  rs.MoveNext
Loop
rs.Close

vPaso = False

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Function fxProrrateo_Guardar() As Long
Dim pExiste As Long, pFactor As String


On Error GoTo vError

fxProrrateo_Guardar = 0

With gProrrateo


.Row = .ActiveRow
.col = 1


strSQL = "SELECT COUNT(*) AS 'EXISTE' FROM RH_NOMINAS_PORRATEO_CC" _
      & " WHERE COD_NOMINA = '" & vCodigo & "' AND COD_CENTRO_COSTO = '" & .Text & "'"
Call OpenRecordSet(rs, strSQL)
 pExiste = rs!Existe
rs.Close


If pExiste = 0 Then  'Insertar
  
  strSQL = "insert into RH_NOMINAS_PORRATEO_CC(COD_NOMINA, COD_CENTRO_COSTO, FACTOR" _
            & ", REGISTRO_USUARIO, REGISTRO_FECHA) values('" & vCodigo & "','" & .Text & "',"
  .col = 3
  
  pFactor = .Text
  
  strSQL = strSQL & CDbl(.Text) & ",'" & glogon.Usuario & "',dbo.Mygetdate())"

  Call ConectionExecute(strSQL)

  .col = 1
  Call Bitacora("Registra", "Nómina: " & vCodigo & ", Prorrateo Cc: " & .Text & ", Fx: " & pFactor)

Else 'Actualizar

 .col = 3
 
  pFactor = .Text
 
 strSQL = "update RH_NOMINAS_PORRATEO_CC set FACTOR = " & CDbl(.Text) _
        & ", ACTUALIZA_FECHA = dbo.mygetdate(), ACTUALIZA_USUARIO = '" & glogon.Usuario & "'" _
        & " where COD_NOMINA = '" & vCodigo & "' AND COD_CENTRO_COSTO = '"
 .col = 1
 strSQL = strSQL & .Text & "'"

 Call ConectionExecute(strSQL)

 .col = 1
 Call Bitacora("Modifica", "Nómina: " & vCodigo & ", Prorrateo Cc: " & .Text & ", Fx: " & pFactor)

End If

End With
fxProrrateo_Guardar = 1

Exit Function

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Function





Private Sub sbProrrateo_Load(pCodigo As String)

On Error GoTo vError

strSQL = "select COD_CENTRO_COSTO, DESCRIPCION, FACTOR" _
       & " FROM vRH_Nominas_Prorrateo WHERE COD_NOMINA = '" & pCodigo _
       & "' ORDER BY DESCRIPCION"
Call sbCargaGrid(gProrrateo, 3, strSQL)

Call sbProrrateo_Suma
Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub sbConsulta(pCodigo As String)

On Error GoTo vError

Me.MousePointer = vbHourglass

 strSQL = "select * " _
        & " from vRH_NOMINA_CATALOGO " _
        & " where COD_NOMINA = '" & pCodigo & "'"
 Call OpenRecordSet(rs, strSQL)

If Not rs.BOF And Not rs.EOF Then
  Call sbToolBar(tlb, "activo")
  tcMain.Item(0).Selected = True
  vEdita = True

  txtCodigo.Text = rs!COD_NOMINA
  vCodigo = rs!COD_NOMINA
  
  txtDescripcion.Text = rs!Descripcion
  chkActivo.Value = rs!ACTIVO
  
  txtConsecutivo.Text = CStr(rs!Consecutivo)
  
  chkAguinaldo.Value = rs!I_AGUINALDO
  chkAdelanto.Value = rs!I_ADELANTO_SALARIO

  txtAdelantoCod.Text = rs!ADELANTO_ID
  txtAdelantoDesc.Text = rs!ADELANTO_DESC
  
  txtAguinaldoCod.Text = rs!AGUINALDO_ID
  txtAguinaldoDesc.Text = rs!AGUINALDO_DESC
  
  txtCuentaCod.Text = rs!CTA_MASK
  txtCuentaDesc.Text = rs!CTA_DESC
  
  txtCuentaLiq.Text = rs!CTA_LIQ_MASK
  txtCuentaLiqDesc.Text = rs!CTA_LIQ_DESC
    
  txtCuentaPre.Text = rs!CTA_PREAVISO_Mask
  txtCuentaPreDesc.Text = rs!CTA_PREAVISO_DESC
    
    
  Call sbCboAsignaDato(cboDivisa, rs!DIVISA_DESC, True, rs!cod_Divisa)
  Call sbCboAsignaDato(cboTipoNomina, rs!TIPO_DESC, True, rs!NOMINA_TIPO)
  Call sbCboAsignaDato(cboFrecuencia, rs!FPAGO_DESC, True, rs!COD_FRECPAGO)
  
  cboDias.Text = CStr(rs!DIAS_CAL_LIQUIDACION)
  
  chkSM_Aplica.Value = rs!SM_Aplica
  txtSM_Monto.Text = Format(rs!SM_Monto, "Standard")
  
  chkProrrateo.Value = rs!PRORRATEO_APL
  txtAdelantoPorc.Text = Format(rs!PORC_ADELANTO, "Standard")
  
Else
  MsgBox "No se encontró registro verifique...", vbInformation
  txtCodigo.Text = ""
  txtCodigo.SetFocus
  Call sbLimpiaDatos
End If

rs.Close
Me.MousePointer = vbDefault
Call RefrescaTags(Me)

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub



Private Sub tlb_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim strSQL As String

Select Case UCase(Button.Key)
    Case "INSERTAR", "NUEVO"
        Call sbLimpiaDatos
        vEdita = False
        txtCodigo.SetFocus
       Call sbToolBar(tlb, "edicion")
    Case "MODIFICAR", "EDITAR"
      vEdita = True
      txtCodigo.SetFocus
      Call sbToolBar(tlb, "edicion")
    Case "BORRAR"
'      Call sbBorrar
    Case "GUARDAR", "SALVAR"
     If fxValida Then Call sbGuardar
     Call sbToolBar(tlb, "activo")
    Case "DESHACER"
      Call sbToolBar(tlb, "activo")
      If vCodigo = "" Then
        Call sbLimpiaDatos
        Call sbToolBar(tlb, "nuevo")
        vEdita = True
      Else
        Call sbConsulta(vCodigo)
      End If

    Case "CONSULTAR"
       gBusquedas.Columna = "descripcion"
       gBusquedas.Orden = "descripcion"
       gBusquedas.Consulta = "select COD_NOMINA,descripcion from RH_NOMINAS_CATALOGO "
       frmBusquedas.Show vbModal
       txtCodigo.SetFocus
       txtCodigo = gBusquedas.Resultado
       
       tcMain.Item(0).Selected = True
       txtDescripcion.SetFocus

End Select


End Sub


Private Sub sbGuardar()

On Error GoTo vError

If fxExiste(txtCodigo.Text) Then
  strSQL = "update RH_NOMINAS_CATALOGO set descripcion = '" & Trim(txtDescripcion.Text) _
         & "', ACTIVO = " & chkActivo.Value & ",COD_DIVISA = '" & cboDivisa.ItemData(cboDivisa.ListIndex) _
         & "', NOMINA_TIPO = '" & cboTipoNomina.ItemData(cboTipoNomina.ListIndex) _
         & "', COD_FRECPAGO = '" & cboFrecuencia.ItemData(cboFrecuencia.ListIndex) _
         & "', I_ADELANTO_SALARIO = " & chkAdelanto.Value & ", C_ADELANTO_SALARIO = '" & txtAdelantoCod.Text _
         & "', I_AGUINALDO = " & chkAguinaldo.Value & ", C_AGUINALDO = '" & txtAguinaldoCod.Text _
         & "', COD_CUENTA = '" & fxgCntCuentaFormato(False, txtCuentaCod.Text, 0) _
         & "', COD_CUENTA_LIQ = '" & fxgCntCuentaFormato(False, txtCuentaLiq.Text, 0) _
         & "', COD_CUENTA_PREAVISO = '" & fxgCntCuentaFormato(False, txtCuentaPre.Text, 0) _
         & "', DIAS_CAL_LIQUIDACION = " & cboDias.Text & ", PRORRATEO_APL = " & chkProrrateo.Value _
         & ", SM_APLICA = " & chkSM_Aplica.Value & ", SM_MONTO = " & CCur(txtSM_Monto.Text) _
         & ", PORC_ADELANTO = " & CCur(txtAdelantoPorc.Text) _
         & " where COD_NOMINA = '" & vCodigo & "' "
         
  Call ConectionExecute(strSQL)
  Call Bitacora("Modifica", "Catálogo de Nómina: " & vCodigo)

Else
  vCodigo = txtCodigo.Text

   strSQL = "insert into RH_NOMINAS_CATALOGO(COD_NOMINA,descripcion,ACTIVO,COD_DIVISA, NOMINA_TIPO,COD_FRECPAGO" _
          & ", I_ADELANTO_SALARIO, C_ADELANTO_SALARIO, I_AGUINALDO, C_AGUINALDO" _
          & ", COD_CUENTA, COD_CUENTA_LIQ, COD_CUENTA_PREAVISO, DIAS_CAL_LIQUIDACION, CONSECUTIVO, PRORRATEO_APL" _
          & ", SM_APLICA, SM_MONTO, PORC_ADELANTO, REGISTRO_USUARIO, REGISTRO_FECHA)" _
          & " values('" & vCodigo & "','" & Trim(txtDescripcion.Text) & "', " & chkActivo.Value _
          & ", '" & cboDivisa.ItemData(cboDivisa.ListIndex) & "', '" & cboTipoNomina.ItemData(cboTipoNomina.ListIndex) _
          & "', '" & cboFrecuencia.ItemData(cboFrecuencia.ListIndex) & "', " & chkAdelanto.Value _
          & ", '" & txtAdelantoCod.Text & "'," & chkAguinaldo.Value & ", '" & txtAguinaldoCod.Text _
          & "', '" & fxgCntCuentaFormato(False, txtCuentaCod.Text, 0) _
          & "', '" & fxgCntCuentaFormato(False, txtCuentaLiq.Text, 0) _
          & "', '" & fxgCntCuentaFormato(False, txtCuentaPre.Text, 0) _
          & "', " & cboDias.Text & "," & txtConsecutivo.Text & ", " & chkProrrateo.Value _
          & ",  " & chkSM_Aplica.Value & ", " & CCur(txtSM_Monto.Text) _
          & ",  " & CCur(txtAdelantoPorc.Text) _
          & ", '" & glogon.Usuario & " ', dbo.MyGetdate())"
   Call ConectionExecute(strSQL)

   Call Bitacora("Registra", "Catálogo de Nómina: " & vCodigo)

End If

MsgBox "Información guardada satisfactoriamente...", vbInformation

Call sbConsulta(vCodigo)

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Function fxValida()

Dim vMensaje As String

vMensaje = ""


If Trim(txtCodigo) = "" Then vMensaje = vMensaje & "- No se ha indicado un código para la Nómina!" & vbCrLf
If Trim(txtDescripcion) = "" Then vMensaje = vMensaje & "- No se ha indicado una descripción!" & vbCrLf


If Not IsNumeric(txtSM_Monto.Text) Then
    vMensaje = vMensaje & "- El Monto del Salario Mínimo no es válido!" & vbCrLf
Else
    If CCur(txtSM_Monto.Text) < 0 Then
        vMensaje = vMensaje & "- El Monto del Salario Mínimo no es válido!" & vbCrLf
    End If
End If

If Not IsNumeric(txtAdelantoPorc.Text) Then
    vMensaje = vMensaje & "- El Monto del Salario Mínimo no es válido!" & vbCrLf
Else
    If CCur(txtAdelantoPorc.Text) < 0 Or CCur(txtAdelantoPorc.Text) > 80 Then
        vMensaje = vMensaje & "- El Porcentaje de Adelantos no puede ser menor a 0 o superio a 80!" & vbCrLf
    End If
End If

If Len(vMensaje) > 0 Then
  MsgBox vMensaje, vbExclamation
  fxValida = False
Else
  fxValida = True
End If


End Function



Private Sub txtCodigo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    tcMain.Item(0).Selected = True
    txtDescripcion.SetFocus
End If


If KeyCode = vbKeyF4 Then
   gBusquedas.Col1Name = "Nomina Id"
   gBusquedas.Col2Name = "Descripción"
   gBusquedas.Columna = "COD_NOMINA"
   gBusquedas.Orden = "COD_NOMINA"
   gBusquedas.Consulta = "select COD_NOMINA,descripcion from RH_NOMINAS_CATALOGO"
   frmBusquedas.Show vbModal
   txtCodigo.Text = gBusquedas.Resultado
   
   tcMain.Item(0).Selected = True
   txtDescripcion.SetFocus
End If

End Sub

Private Sub txtCuentaCod_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtCuentaDesc.SetFocus

If KeyCode = vbKeyF4 Then
   frmCntX_ConsultaCuentas.Show vbModal
   txtCuentaCod.Text = gCuenta
   txtCuentaDesc.Text = fxgCntCuentaDesc(gCuenta)
   txtCuentaCod.Text = fxgCntCuentaFormato(True, txtCuentaCod, 0)
End If

End Sub

Private Sub txtCuentaCod_LostFocus()
   txtCuentaDesc.Text = fxgCntCuentaDesc(fxgCntCuentaFormato(False, txtCuentaCod, 0))
   txtCuentaCod.Text = fxgCntCuentaFormato(True, txtCuentaCod, 0)
End Sub


Private Sub txtCuentaLiq_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtCuentaLiqDesc.SetFocus

If KeyCode = vbKeyF4 Then
   frmCntX_ConsultaCuentas.Show vbModal
   txtCuentaLiq.Text = gCuenta
   txtCuentaLiqDesc.Text = fxgCntCuentaDesc(gCuenta)
   txtCuentaLiq.Text = fxgCntCuentaFormato(True, txtCuentaLiq, 0)
End If
End Sub


Private Sub txtCuentaLiq_LostFocus()
   txtCuentaLiqDesc.Text = fxgCntCuentaDesc(fxgCntCuentaFormato(False, txtCuentaLiq, 0))
   txtCuentaLiq.Text = fxgCntCuentaFormato(True, txtCuentaLiq, 0)
End Sub

Private Sub txtCuentaPre_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtCuentaPreDesc.SetFocus

If KeyCode = vbKeyF4 Then
   frmCntX_ConsultaCuentas.Show vbModal
   txtCuentaPre.Text = gCuenta
   txtCuentaPreDesc.Text = fxgCntCuentaDesc(gCuenta)
   txtCuentaPre.Text = fxgCntCuentaFormato(True, txtCuentaPre, 0)
End If
End Sub

Private Sub txtCuentaPre_LostFocus()
   txtCuentaPreDesc.Text = fxgCntCuentaDesc(fxgCntCuentaFormato(False, txtCuentaPre, 0))
   txtCuentaPre.Text = fxgCntCuentaFormato(True, txtCuentaPre, 0)
End Sub

Private Sub txtDescripcion_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then cboTipoNomina.SetFocus
End Sub


Private Sub txtSM_Monto_GotFocus()
On Error GoTo vError

txtSM_Monto.Text = CCur(txtSM_Monto.Text)

vError:
End Sub

Private Sub txtSM_Monto_LostFocus()
On Error GoTo vError

txtSM_Monto.Text = Format(CCur(txtSM_Monto.Text), "Standard")

vError:

End Sub
