VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.Controls.v24.0.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.ShortcutBar.v24.0.0.ocx"
Begin VB.Form frmPrea_BallonPayment 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Ballon Payment"
   ClientHeight    =   5220
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8145
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5220
   ScaleWidth      =   8145
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin XtremeSuiteControls.TabControl tcMain 
      Height          =   4935
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7935
      _Version        =   1572864
      _ExtentX        =   13996
      _ExtentY        =   8705
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
      Item(0).Caption =   "Condiciones"
      Item(0).ControlCount=   2
      Item(0).Control(0)=   "gbCondiciones"
      Item(0).Control(1)=   "ShortcutCaption1(1)"
      Item(1).Caption =   "Ballon Payment"
      Item(1).ControlCount=   15
      Item(1).Control(0)=   "Label1(0)"
      Item(1).Control(1)=   "Label1(1)"
      Item(1).Control(2)=   "Label1(2)"
      Item(1).Control(3)=   "Label1(3)"
      Item(1).Control(4)=   "Label1(4)"
      Item(1).Control(5)=   "Label1(5)"
      Item(1).Control(6)=   "txtMonto"
      Item(1).Control(7)=   "txtCuota"
      Item(1).Control(8)=   "txtTasa"
      Item(1).Control(9)=   "txtPlazo"
      Item(1).Control(10)=   "txtCuotaBallon"
      Item(1).Control(11)=   "cboTipoPago"
      Item(1).Control(12)=   "btnBallon(0)"
      Item(1).Control(13)=   "btnBallon(1)"
      Item(1).Control(14)=   "ShortcutCaption1(0)"
      Item(2).Caption =   "Tabla de Pagos"
      Item(2).ControlCount=   3
      Item(2).Control(0)=   "vGrid"
      Item(2).Control(1)=   "btnBallon(2)"
      Item(2).Control(2)=   "ShortcutCaption1(2)"
      Item(3).Caption =   "Ahorro Especial"
      Item(3).ControlCount=   11
      Item(3).Control(0)=   "ShortcutCaption1(3)"
      Item(3).Control(1)=   "Label1(6)"
      Item(3).Control(2)=   "FlatEdit1"
      Item(3).Control(3)=   "Label1(7)"
      Item(3).Control(4)=   "FlatEdit2"
      Item(3).Control(5)=   "Label1(8)"
      Item(3).Control(6)=   "FlatEdit3"
      Item(3).Control(7)=   "Label1(9)"
      Item(3).Control(8)=   "FlatEdit4"
      Item(3).Control(9)=   "Label2"
      Item(3).Control(10)=   "Label1(10)"
      Begin XtremeSuiteControls.GroupBox gbCondiciones 
         Height          =   4215
         Left            =   0
         TabIndex        =   1
         Top             =   720
         Width           =   8175
         _Version        =   1572864
         _ExtentX        =   14420
         _ExtentY        =   7435
         _StockProps     =   79
         Caption         =   "Condiciones (Línea de Recuperación de Mora)"
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
         Appearance      =   21
         BorderStyle     =   2
         Begin XtremeSuiteControls.PushButton btnCondiciones 
            Height          =   495
            Left            =   3960
            TabIndex        =   4
            Top             =   2280
            Width           =   1455
            _Version        =   1572864
            _ExtentX        =   2566
            _ExtentY        =   873
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
            Appearance      =   21
            Picture         =   "frmPrea_BallonPayment.frx":0000
         End
         Begin XtremeSuiteControls.CheckBox chkTrasladaSalario 
            Height          =   255
            Left            =   2400
            TabIndex        =   2
            Top             =   600
            Width           =   2295
            _Version        =   1572864
            _ExtentX        =   4048
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Traslado de Salario"
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
            Appearance      =   21
         End
         Begin XtremeSuiteControls.CheckBox chkDeduccionPlanilla 
            Height          =   255
            Left            =   2400
            TabIndex        =   3
            Top             =   1080
            Width           =   2295
            _Version        =   1572864
            _ExtentX        =   4048
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Deducción de Planilla"
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
            Appearance      =   21
         End
      End
      Begin XtremeSuiteControls.FlatEdit txtMonto 
         Height          =   315
         Left            =   -67720
         TabIndex        =   11
         Top             =   960
         Visible         =   0   'False
         Width           =   2055
         _Version        =   1572864
         _ExtentX        =   3619
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
         Alignment       =   1
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtCuota 
         Height          =   315
         Left            =   -67720
         TabIndex        =   12
         Top             =   3240
         Visible         =   0   'False
         Width           =   2055
         _Version        =   1572864
         _ExtentX        =   3619
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
         Alignment       =   1
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtTasa 
         Height          =   315
         Left            =   -66760
         TabIndex        =   13
         Top             =   1800
         Visible         =   0   'False
         Width           =   1095
         _Version        =   1572864
         _ExtentX        =   1926
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
      Begin XtremeSuiteControls.FlatEdit txtPlazo 
         Height          =   315
         Left            =   -66760
         TabIndex        =   14
         Top             =   1440
         Visible         =   0   'False
         Width           =   1095
         _Version        =   1572864
         _ExtentX        =   1926
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
      Begin XtremeSuiteControls.ComboBox cboTipoPago 
         Height          =   330
         Left            =   -67720
         TabIndex        =   15
         Top             =   2280
         Visible         =   0   'False
         Width           =   2055
         _Version        =   1572864
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
      Begin XtremeSuiteControls.FlatEdit txtCuotaBallon 
         Height          =   315
         Left            =   -67720
         TabIndex        =   16
         Top             =   2880
         Visible         =   0   'False
         Width           =   2055
         _Version        =   1572864
         _ExtentX        =   3619
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
      Begin XtremeSuiteControls.PushButton btnBallon 
         Height          =   495
         Index           =   0
         Left            =   -67120
         TabIndex        =   17
         Top             =   3960
         Visible         =   0   'False
         Width           =   1455
         _Version        =   1572864
         _ExtentX        =   2566
         _ExtentY        =   873
         _StockProps     =   79
         Caption         =   "Calcular"
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
         Appearance      =   21
         Picture         =   "frmPrea_BallonPayment.frx":0731
      End
      Begin XtremeSuiteControls.PushButton btnBallon 
         Height          =   495
         Index           =   1
         Left            =   -65680
         TabIndex        =   18
         Top             =   3960
         Visible         =   0   'False
         Width           =   1455
         _Version        =   1572864
         _ExtentX        =   2566
         _ExtentY        =   873
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
         Appearance      =   21
         Picture         =   "frmPrea_BallonPayment.frx":0E16
      End
      Begin FPSpreadADO.fpSpread vGrid 
         Height          =   3975
         Left            =   -70000
         TabIndex        =   21
         Top             =   840
         Visible         =   0   'False
         Width           =   7935
         _Version        =   524288
         _ExtentX        =   13996
         _ExtentY        =   7011
         _StockProps     =   64
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
         MaxCols         =   5
         ScrollBars      =   2
         SpreadDesigner  =   "frmPrea_BallonPayment.frx":1547
         AppearanceStyle =   1
      End
      Begin XtremeSuiteControls.PushButton btnBallon 
         Height          =   375
         Index           =   2
         Left            =   -62680
         TabIndex        =   22
         Top             =   360
         Visible         =   0   'False
         Width           =   615
         _Version        =   1572864
         _ExtentX        =   1085
         _ExtentY        =   661
         _StockProps     =   79
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
         Appearance      =   21
         Picture         =   "frmPrea_BallonPayment.frx":1BCD
      End
      Begin XtremeSuiteControls.FlatEdit FlatEdit1 
         Height          =   315
         Left            =   -67840
         TabIndex        =   26
         Top             =   960
         Visible         =   0   'False
         Width           =   2055
         _Version        =   1572864
         _ExtentX        =   3619
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
         Alignment       =   1
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit FlatEdit2 
         Height          =   315
         Left            =   -67840
         TabIndex        =   28
         Top             =   1320
         Visible         =   0   'False
         Width           =   2055
         _Version        =   1572864
         _ExtentX        =   3619
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
         Alignment       =   1
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit FlatEdit3 
         Height          =   315
         Left            =   -67840
         TabIndex        =   30
         Top             =   1800
         Visible         =   0   'False
         Width           =   2055
         _Version        =   1572864
         _ExtentX        =   3619
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
         Alignment       =   1
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit FlatEdit4 
         Height          =   315
         Left            =   -67840
         TabIndex        =   32
         Top             =   2160
         Visible         =   0   'False
         Width           =   2055
         _Version        =   1572864
         _ExtentX        =   3619
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
         Alignment       =   1
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.Label Label1 
         Height          =   255
         Index           =   10
         Left            =   -65680
         TabIndex        =   34
         Top             =   2160
         Visible         =   0   'False
         Width           =   2175
         _Version        =   1572864
         _ExtentX        =   3836
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "(Según su periodicidad)"
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
         Height          =   615
         Left            =   -69280
         TabIndex        =   33
         Top             =   2880
         Visible         =   0   'False
         Width           =   6255
         _Version        =   1572864
         _ExtentX        =   11033
         _ExtentY        =   1085
         _StockProps     =   79
         Caption         =   "Nota: La información se completa cuando se formalice la operación."
         BackColor       =   12648447
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   2
         WordWrap        =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label1 
         Height          =   255
         Index           =   9
         Left            =   -69520
         TabIndex        =   31
         Top             =   2160
         Visible         =   0   'False
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
         Index           =   8
         Left            =   -69520
         TabIndex        =   29
         Top             =   1800
         Visible         =   0   'False
         Width           =   1575
         _Version        =   1572864
         _ExtentX        =   2778
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Plazo"
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
         Index           =   7
         Left            =   -69520
         TabIndex        =   27
         Top             =   1320
         Visible         =   0   'False
         Width           =   1575
         _Version        =   1572864
         _ExtentX        =   2778
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "No. Contrato"
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
         Index           =   6
         Left            =   -69520
         TabIndex        =   25
         Top             =   960
         Visible         =   0   'False
         Width           =   1575
         _Version        =   1572864
         _ExtentX        =   2778
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Código del Plan"
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
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
         Height          =   375
         Index           =   3
         Left            =   -70000
         TabIndex        =   24
         Top             =   360
         Visible         =   0   'False
         Width           =   7935
         _Version        =   1572864
         _ExtentX        =   13996
         _ExtentY        =   661
         _StockProps     =   14
         Caption         =   "Ahorro Especial: Fondo FCOB"
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
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
         Height          =   375
         Index           =   2
         Left            =   -70000
         TabIndex        =   23
         Top             =   360
         Visible         =   0   'False
         Width           =   7815
         _Version        =   1572864
         _ExtentX        =   13785
         _ExtentY        =   661
         _StockProps     =   14
         Caption         =   "Ballon Payment: Tabla de Pagos"
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
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
         Height          =   375
         Index           =   1
         Left            =   0
         TabIndex        =   20
         Top             =   360
         Width           =   7935
         _Version        =   1572864
         _ExtentX        =   13996
         _ExtentY        =   661
         _StockProps     =   14
         Caption         =   "Condiciones (Línea de Recuperación de Mora)"
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
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
         Height          =   375
         Index           =   0
         Left            =   -70000
         TabIndex        =   19
         Top             =   360
         Visible         =   0   'False
         Width           =   7935
         _Version        =   1572864
         _ExtentX        =   13996
         _ExtentY        =   661
         _StockProps     =   14
         Caption         =   "Ballon Payment: Recuperación Mora"
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
      Begin XtremeSuiteControls.Label Label1 
         Height          =   255
         Index           =   5
         Left            =   -69400
         TabIndex        =   10
         Top             =   3240
         Visible         =   0   'False
         Width           =   1575
         _Version        =   1572864
         _ExtentX        =   2778
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Cuota del Plan"
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
         Index           =   4
         Left            =   -69400
         TabIndex        =   9
         Top             =   2880
         Visible         =   0   'False
         Width           =   1575
         _Version        =   1572864
         _ExtentX        =   2778
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Cuota Final"
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
         Index           =   3
         Left            =   -69400
         TabIndex        =   8
         Top             =   2280
         Visible         =   0   'False
         Width           =   2415
         _Version        =   1572864
         _ExtentX        =   4260
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Tipo de Pago"
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
         Left            =   -69400
         TabIndex        =   7
         Top             =   1800
         Visible         =   0   'False
         Width           =   2415
         _Version        =   1572864
         _ExtentX        =   4260
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Plazo en Meses"
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
         Left            =   -69400
         TabIndex        =   6
         Top             =   1440
         Visible         =   0   'False
         Width           =   2415
         _Version        =   1572864
         _ExtentX        =   4260
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Tasa Anual"
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
         Index           =   0
         Left            =   -69400
         TabIndex        =   5
         Top             =   960
         Visible         =   0   'False
         Width           =   1575
         _Version        =   1572864
         _ExtentX        =   2778
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Monto del Crédito"
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
End
Attribute VB_Name = "frmPrea_BallonPayment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSQL As String, rs As New ADODB.Recordset
Dim vPaso As Boolean, mExpediente As String

Private Sub btnBallon_Click(Index As Integer)

'Guardar
'spCrdPreaRegistraBalloonRecuperaMora

'Leer la TAbla

'Lee Ahorro
'spCrdPreaObtenerBalloonPayment
'spCrdPreaObtenerBalloonPayment_ProGrX
'


On Error GoTo vError

Select Case Index

    Case 0 'Calculo
            strSQL = "exec spCRD_PREA_Calculo_BalloonPayment " & CCur(txtMonto.Text) & ", " & CCur(txtTasa.Text) _
                   & ", " & txtPlazo.Text & ", " & cboTipoPago.ItemData(cboTipoPago.ListIndex) _
                   & ", " & CCur(txtCuotaBallon.Text) & ", '" & mExpediente & "'"
            Call ConectionExecute(strSQL)

            strSQL = "select  Top 1 MONTO_CUOTA  From CRD_PREA_TABLA_PAGOS_BALLOON" _
                   & " Where COD_PREANALISIS = '" & mExpediente & "'" _
                   & "order by ID_CUOTA asc"
            Call OpenRecordSet(rs, strSQL)
            
            txtCuota.Text = Format(rs!Monto_Cuota, "Standard")

    Case 1 'Guardar

        strSQL = "exec spCrdPreaRegistraBalloonRecuperaMora '" & mExpediente & "', " & cboTipoPago.ItemData(cboTipoPago.ListIndex) _
               & ", " & txtTasa.Text & ", " & txtPlazo.Text & ", " & CCur(txtCuotaBallon.Text) & ", " & CCur(txtCuota.Text) _
               & ", " & CCur(txtMonto.Text) & ", '" & glogon.Usuario & "'"
        Call ConectionExecute(strSQL)


        'spCRDCreaFondosPreanalisis

End Select

Me.MousePointer = vbDefault

Exit Sub

vError:
    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub btnCondiciones_Click()

On Error GoTo vError

strSQL = "exec spCrdPreaCondicionesRecuperaMora '" & mExpediente & "', " & chkTrasladaSalario.Value & ", " & chkDeduccionPlanilla.Value
Call ConectionExecute(strSQL)

Call Bitacora("Registra", "BallonPayment Condiciones, Expediente " & mExpediente & ", Traslado Salario [" & chkTrasladaSalario.Value _
        & "] Deduce de Planilla [" & chkDeduccionPlanilla.Value & "]")

MsgBox "Condiciones para Obtener la Linea actualizadas satisfactoriamente!", vbInformation


Exit Sub

vError:

End Sub

Private Sub Form_Load()

vModulo = 3

mExpediente = GLOBALES.gTag

tcMain.Item(0).Selected = True

'spCrdPreaObtenerCondicionesMora
'Checks

strSQL = "SELECT ID_PERIODICIDAD as 'Idx',DESCRIPCION as 'ItmX' FROM CRD_PREA_PERIODICIDAD WHERE ACTIVO = 1"
Call sbCbo_Llena_New(cboTipoPago, strSQL, False, True)

Call sbBallonPayment_Load

End Sub

Private Sub sbBallonPayment_Load()

On Error GoTo vError

strSQL = "exec spCrdPreaObtenerBalloonPayment_ProGrX '" & mExpediente & "'"
Call OpenRecordSet(rs, strSQL)


chkDeduccionPlanilla.Value = rs!DEDUCE_PLANILLA
chkTrasladaSalario.Value = rs!TRASLADA_SALARIO

txtMonto.Text = Format(rs!Monto, "Standard")
txtCuota.Text = Format(rs!Cuota, "Standard")
txtCuotaBallon.Text = Format(rs!CUOTA_BALLOON, "Standard")
txtTasa.Text = Format(rs!TASA, "Standard")
txtPlazo.Text = rs!Plazo

If rs!PERIODICIDAD > 0 Then
    Call sbCboAsignaDato(cboTipoPago, rs!PERIODICIDAD_DESC, False, rs!PERIODICIDAD)
End If

Exit Sub

vError:

End Sub


Private Sub sbTabla()

On Error GoTo vError

'strSQL = "exec spCrdPreaObtenerBalloonPaymentEstudio '" & mExpediente & "'"

strSQL = "SELECT ID_CUOTA,[MONTO_CUOTA],[AMORTIZA],[INTERESES],[MONTO_PRINCIPAL]" _
       & "  From [dbo].[CRD_PREA_TABLA_PAGOS_BALLOON]" _
       & "  WHERE COD_PREANALISIS = '" & mExpediente & "'  ORDER BY ID_CUOTA ASC"
Call sbCargaGrid(vGrid, 5, strSQL, True)

vError:

End Sub

Private Sub tcMain_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)


Select Case Item.Index
    Case 2 'Tabla
        Call sbTabla
    Case 3 'Ahorro
End Select

End Sub


Private Sub txtCuotaBallon_GotFocus()
On Error GoTo vError

txtCuotaBallon.Text = CCur(txtCuotaBallon.Text)

vError:
End Sub

Private Sub txtCuotaBallon_LostFocus()
On Error GoTo vError

txtCuotaBallon.Text = Format(CCur(txtCuotaBallon.Text), "Standard")

vError:

End Sub
