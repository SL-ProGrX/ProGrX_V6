VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.0#0"; "Codejock.Controls.v22.0.0.ocx"
Begin VB.Form frmCR_APA_Movimientos 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consulta y Movimientos de la Operación"
   ClientHeight    =   7965
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11775
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7965
   ScaleWidth      =   11775
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.TabControl tcMain 
      Height          =   6495
      Left            =   0
      TabIndex        =   11
      Top             =   1440
      Width           =   11775
      _Version        =   1441792
      _ExtentX        =   20770
      _ExtentY        =   11456
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
      Item(0).Caption =   "Resumen"
      Item(0).ControlCount=   28
      Item(0).Control(0)=   "Label1(5)"
      Item(0).Control(1)=   "Label1(4)"
      Item(0).Control(2)=   "Label1(3)"
      Item(0).Control(3)=   "Label1(2)"
      Item(0).Control(4)=   "txtPlazo"
      Item(0).Control(5)=   "txtTasa"
      Item(0).Control(6)=   "txtSaldo"
      Item(0).Control(7)=   "txtMonto"
      Item(0).Control(8)=   "txtDivisa"
      Item(0).Control(9)=   "Label1(9)"
      Item(0).Control(10)=   "Label1(8)"
      Item(0).Control(11)=   "Label1(7)"
      Item(0).Control(12)=   "Label1(6)"
      Item(0).Control(13)=   "txtFecFormaliza"
      Item(0).Control(14)=   "txtFecPrimerPago"
      Item(0).Control(15)=   "txtFecProxPago"
      Item(0).Control(16)=   "txtDiaPago"
      Item(0).Control(17)=   "Label1(13)"
      Item(0).Control(18)=   "Label1(12)"
      Item(0).Control(19)=   "Label1(11)"
      Item(0).Control(20)=   "Label1(10)"
      Item(0).Control(21)=   "txtAmortizado"
      Item(0).Control(22)=   "txtIntereses"
      Item(0).Control(23)=   "txtComisiones"
      Item(0).Control(24)=   "txtCargos"
      Item(0).Control(25)=   "Label1(0)"
      Item(0).Control(26)=   "txtNotas"
      Item(0).Control(27)=   "gbMovimiento"
      Item(1).Caption =   "Detalle"
      Item(1).ControlCount=   1
      Item(1).Control(0)=   "vGrid"
      Begin FPSpreadADO.fpSpread vGrid 
         Height          =   5775
         Left            =   -70000
         TabIndex        =   12
         Top             =   480
         Visible         =   0   'False
         Width           =   11775
         _Version        =   524288
         _ExtentX        =   20770
         _ExtentY        =   10186
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
         MaxCols         =   16
         ScrollBarExtMode=   -1  'True
         SpreadDesigner  =   "frmCR_APA_Movimientos.frx":0000
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin XtremeSuiteControls.FlatEdit txtSaldo 
         Height          =   330
         Left            =   1800
         TabIndex        =   19
         Top             =   1920
         Width           =   2055
         _Version        =   1441792
         _ExtentX        =   3625
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
         Alignment       =   1
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtTasa 
         Height          =   330
         Left            =   2880
         TabIndex        =   18
         Top             =   1560
         Width           =   975
         _Version        =   1441792
         _ExtentX        =   1720
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
         Alignment       =   1
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtPlazo 
         Height          =   330
         Left            =   2880
         TabIndex        =   17
         Top             =   1200
         Width           =   975
         _Version        =   1441792
         _ExtentX        =   1720
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
         Alignment       =   1
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtMonto 
         Height          =   330
         Left            =   1800
         TabIndex        =   20
         Top             =   840
         Width           =   2055
         _Version        =   1441792
         _ExtentX        =   3625
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
         Alignment       =   1
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtDivisa 
         Height          =   330
         Left            =   2880
         TabIndex        =   21
         Top             =   480
         Width           =   975
         _Version        =   1441792
         _ExtentX        =   1720
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
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtFecFormaliza 
         Height          =   330
         Left            =   5640
         TabIndex        =   26
         Top             =   840
         Width           =   2055
         _Version        =   1441792
         _ExtentX        =   3625
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
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtFecPrimerPago 
         Height          =   330
         Left            =   5640
         TabIndex        =   27
         Top             =   1200
         Width           =   2055
         _Version        =   1441792
         _ExtentX        =   3625
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
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtFecProxPago 
         Height          =   330
         Left            =   5640
         TabIndex        =   28
         Top             =   1560
         Width           =   2055
         _Version        =   1441792
         _ExtentX        =   3625
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
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtDiaPago 
         Height          =   330
         Left            =   5640
         TabIndex        =   29
         Top             =   1920
         Width           =   2055
         _Version        =   1441792
         _ExtentX        =   3625
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
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtAmortizado 
         Height          =   330
         Left            =   9600
         TabIndex        =   34
         Top             =   840
         Width           =   2055
         _Version        =   1441792
         _ExtentX        =   3625
         _ExtentY        =   582
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
      Begin XtremeSuiteControls.FlatEdit txtIntereses 
         Height          =   330
         Left            =   9600
         TabIndex        =   35
         Top             =   1200
         Width           =   2055
         _Version        =   1441792
         _ExtentX        =   3625
         _ExtentY        =   582
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
      Begin XtremeSuiteControls.FlatEdit txtComisiones 
         Height          =   330
         Left            =   9600
         TabIndex        =   36
         Top             =   1560
         Width           =   2055
         _Version        =   1441792
         _ExtentX        =   3625
         _ExtentY        =   582
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
      Begin XtremeSuiteControls.FlatEdit txtCargos 
         Height          =   330
         Left            =   9600
         TabIndex        =   37
         Top             =   1920
         Width           =   2055
         _Version        =   1441792
         _ExtentX        =   3625
         _ExtentY        =   582
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
      Begin XtremeSuiteControls.FlatEdit txtNotas 
         Height          =   690
         Left            =   1800
         TabIndex        =   39
         Top             =   2400
         Width           =   9855
         _Version        =   1441792
         _ExtentX        =   17383
         _ExtentY        =   1217
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
         MultiLine       =   -1  'True
         ScrollBars      =   2
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.GroupBox gbMovimiento 
         Height          =   2655
         Left            =   120
         TabIndex        =   40
         Top             =   3480
         Width           =   12135
         _Version        =   1441792
         _ExtentX        =   21405
         _ExtentY        =   4683
         _StockProps     =   79
         Caption         =   "Movimiento"
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
         Begin XtremeSuiteControls.RadioButton optMov 
            Height          =   330
            Index           =   0
            Left            =   4080
            TabIndex        =   60
            Top             =   2160
            Width           =   1575
            _Version        =   1441792
            _ExtentX        =   2778
            _ExtentY        =   582
            _StockProps     =   79
            Caption         =   "Abono"
            BackColor       =   12648384
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
            Value           =   -1  'True
         End
         Begin XtremeSuiteControls.FlatEdit txtMovTotal 
            Height          =   330
            Left            =   1680
            TabIndex        =   46
            Top             =   2160
            Width           =   2175
            _Version        =   1441792
            _ExtentX        =   3836
            _ExtentY        =   582
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
         Begin XtremeSuiteControls.FlatEdit txtMovAmortiza 
            Height          =   330
            Left            =   1680
            TabIndex        =   47
            Top             =   360
            Width           =   2175
            _Version        =   1441792
            _ExtentX        =   3836
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
            Alignment       =   1
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtMovIntereses 
            Height          =   330
            Left            =   1680
            TabIndex        =   48
            Top             =   720
            Width           =   2175
            _Version        =   1441792
            _ExtentX        =   3836
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
            Alignment       =   1
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtMovComision 
            Height          =   330
            Left            =   1680
            TabIndex        =   49
            Top             =   1080
            Width           =   2175
            _Version        =   1441792
            _ExtentX        =   3836
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
            Alignment       =   1
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtMovCargos 
            Height          =   330
            Left            =   1680
            TabIndex        =   50
            Top             =   1440
            Width           =   2175
            _Version        =   1441792
            _ExtentX        =   3836
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
            Alignment       =   1
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtCuenta 
            Height          =   330
            Left            =   4080
            TabIndex        =   54
            Top             =   1440
            Width           =   1815
            _Version        =   1441792
            _ExtentX        =   3201
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
         Begin XtremeSuiteControls.FlatEdit txtCuentaDesc 
            Height          =   330
            Left            =   5880
            TabIndex        =   55
            Top             =   1440
            Width           =   5535
            _Version        =   1441792
            _ExtentX        =   9763
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
            BackColor       =   16777215
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtMovDocRef 
            Height          =   690
            Left            =   4080
            TabIndex        =   56
            Top             =   360
            Width           =   1815
            _Version        =   1441792
            _ExtentX        =   3201
            _ExtentY        =   1217
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
            MultiLine       =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtMovNotas 
            Height          =   690
            Left            =   6000
            TabIndex        =   57
            Top             =   360
            Width           =   5415
            _Version        =   1441792
            _ExtentX        =   9551
            _ExtentY        =   1217
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
            MultiLine       =   -1  'True
            ScrollBars      =   2
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.PushButton btnMovimiento 
            Height          =   375
            Index           =   0
            Left            =   8520
            TabIndex        =   58
            Top             =   2160
            Width           =   1455
            _Version        =   1441792
            _ExtentX        =   2566
            _ExtentY        =   661
            _StockProps     =   79
            Caption         =   "Informe"
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
            Picture         =   "frmCR_APA_Movimientos.frx":0A9E
         End
         Begin XtremeSuiteControls.PushButton btnMovimiento 
            Height          =   375
            Index           =   1
            Left            =   9960
            TabIndex        =   59
            Top             =   2160
            Width           =   1455
            _Version        =   1441792
            _ExtentX        =   2566
            _ExtentY        =   661
            _StockProps     =   79
            Caption         =   "Aplicar"
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
            Picture         =   "frmCR_APA_Movimientos.frx":11A5
         End
         Begin XtremeSuiteControls.RadioButton optMov 
            Height          =   330
            Index           =   1
            Left            =   5660
            TabIndex        =   61
            Top             =   2160
            Width           =   1575
            _Version        =   1441792
            _ExtentX        =   2778
            _ExtentY        =   582
            _StockProps     =   79
            Caption         =   "Anulación"
            BackColor       =   12632319
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
         Begin VB.Label Label1 
            Caption         =   "Notas del Movimiento:"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   19
            Left            =   6000
            TabIndex        =   53
            Top             =   120
            Width           =   3135
         End
         Begin VB.Label Label1 
            Caption         =   "Cuenta Contable para Afectar:"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   21
            Left            =   4080
            TabIndex        =   52
            Top             =   1080
            Width           =   3135
         End
         Begin VB.Label Label1 
            Caption         =   "Doc. Referencia:"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   23
            Left            =   4080
            TabIndex        =   51
            Top             =   120
            Width           =   1575
         End
         Begin VB.Label Label1 
            Caption         =   "Total Movimiento"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   20
            Left            =   120
            TabIndex        =   45
            Top             =   2160
            Width           =   1575
         End
         Begin VB.Label Label1 
            Caption         =   "Cargos"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   18
            Left            =   120
            TabIndex        =   44
            Top             =   1440
            Width           =   1575
         End
         Begin VB.Label Label1 
            Caption         =   "Comisión"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   17
            Left            =   120
            TabIndex        =   43
            Top             =   1080
            Width           =   1575
         End
         Begin VB.Label Label1 
            Caption         =   "Intereses"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   16
            Left            =   120
            TabIndex        =   42
            Top             =   720
            Width           =   1575
         End
         Begin VB.Label Label1 
            Caption         =   "Amortización"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   15
            Left            =   120
            TabIndex        =   41
            Top             =   360
            Width           =   1575
         End
      End
      Begin VB.Label Label1 
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
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   38
         Top             =   2400
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Amortizado"
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
         Index           =   10
         Left            =   8040
         TabIndex        =   33
         Top             =   900
         Width           =   1095
      End
      Begin VB.Label Label1 
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
         Height          =   255
         Index           =   11
         Left            =   8040
         TabIndex        =   32
         Top             =   1260
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Comisiones"
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
         Index           =   12
         Left            =   8040
         TabIndex        =   31
         Top             =   1620
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Cargos"
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
         Index           =   13
         Left            =   8040
         TabIndex        =   30
         Top             =   1980
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Formalización"
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
         Index           =   6
         Left            =   4080
         TabIndex        =   25
         Top             =   900
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Primer Pago"
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
         Index           =   7
         Left            =   4080
         TabIndex        =   24
         Top             =   1260
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Próximo Pago"
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
         Index           =   8
         Left            =   4080
         TabIndex        =   23
         Top             =   1620
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Día Pago"
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
         Index           =   9
         Left            =   4080
         TabIndex        =   22
         Top             =   1980
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Monto"
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
         Index           =   2
         Left            =   240
         TabIndex        =   16
         Top             =   840
         Width           =   975
      End
      Begin VB.Label Label1 
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
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   15
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label Label1 
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
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   14
         Top             =   1560
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Saldo"
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
         Index           =   5
         Left            =   240
         TabIndex        =   13
         Top             =   1920
         Width           =   975
      End
   End
   Begin XtremeSuiteControls.GroupBox gbCuenta 
      Height          =   1455
      Left            =   360
      TabIndex        =   0
      Top             =   -120
      Width           =   10935
      _Version        =   1441792
      _ExtentX        =   19288
      _ExtentY        =   2566
      _StockProps     =   79
      BackColor       =   16777215
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      BorderStyle     =   1
      Begin XtremeSuiteControls.CheckBox chkSaldo 
         Height          =   255
         Left            =   7080
         TabIndex        =   10
         Top             =   1080
         Width           =   1815
         _Version        =   1441792
         _ExtentX        =   3201
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Solo con Saldo"
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
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         Value           =   1
      End
      Begin XtremeSuiteControls.FlatEdit txtAcreedor 
         Height          =   330
         Left            =   120
         TabIndex        =   2
         Top             =   600
         Width           =   1815
         _Version        =   1441792
         _ExtentX        =   3201
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
      Begin XtremeSuiteControls.FlatEdit txtAcreedorDesc 
         Height          =   330
         Left            =   1920
         TabIndex        =   3
         Top             =   600
         Width           =   6255
         _Version        =   1441792
         _ExtentX        =   11033
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
         BackColor       =   16777215
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtOperacion 
         Height          =   330
         Left            =   1920
         TabIndex        =   7
         Top             =   1080
         Width           =   2655
         _Version        =   1441792
         _ExtentX        =   4683
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
      Begin XtremeSuiteControls.FlatEdit txtEstado 
         Height          =   330
         Left            =   4560
         TabIndex        =   8
         Top             =   1080
         Width           =   1455
         _Version        =   1441792
         _ExtentX        =   2566
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
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin MSComCtl2.FlatScrollBar FlatScrollBar 
         Height          =   255
         Left            =   6120
         TabIndex        =   9
         Top             =   1080
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   450
         _Version        =   393216
         Arrows          =   65536
         Orientation     =   1638401
      End
      Begin XtremeSuiteControls.FlatEdit txtAcreedorSaldo 
         Height          =   330
         Left            =   8160
         TabIndex        =   5
         Top             =   600
         Width           =   2655
         _Version        =   1441792
         _ExtentX        =   4683
         _ExtentY        =   582
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
      Begin XtremeSuiteControls.Label Label2 
         Height          =   375
         Index           =   2
         Left            =   120
         TabIndex        =   6
         Top             =   1080
         Width           =   1815
         _Version        =   1441792
         _ExtentX        =   3201
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "No. Operación"
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
      Begin XtremeSuiteControls.Label Label2 
         Height          =   375
         Index           =   1
         Left            =   8160
         TabIndex        =   4
         Top             =   240
         Width           =   2415
         _Version        =   1441792
         _ExtentX        =   4260
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Saldo en Operaciones:"
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
      Begin XtremeSuiteControls.Label Label2 
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1815
         _Version        =   1441792
         _ExtentX        =   3201
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Acreedor"
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
   End
End
Attribute VB_Name = "frmCR_APA_Movimientos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vScroll As Boolean, vPaso As Boolean



Public Sub sbConsulta(pAcreedor As String)
Dim rs As New ADODB.Recordset, strSQL As String

On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = "exec spAPA_ConsultaAcreedor '" & pAcreedor & "'"
Call OpenRecordSet(rs, strSQL)

If Not rs.BOF And Not rs.EOF Then
   txtAcreedorDesc.Text = rs!Descripcion & ""
   txtAcreedorSaldo.Text = Format(rs!Saldo, "Standard")
   txtOperacion.SetFocus
Else
   txtAcreedorDesc.Text = ""
   txtAcreedorSaldo.Text = ""
End If

rs.Close

Call sbLimpiaPantalla(True)
Call RefrescaTags(Me)

Me.MousePointer = vbDefault

Exit Sub
vError:
    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub


Public Sub sbConsultaOperacion(pOperacion As String)
Dim rs As New ADODB.Recordset, strSQL As String

On Error GoTo vError

Me.MousePointer = vbHourglass

Call sbLimpiaPantalla(False)

strSQL = "exec spAPA_ConsultaOperacion '" & txtAcreedor.Text & "','" & pOperacion & "'"
Call OpenRecordSet(rs, strSQL)

If Not rs.BOF And Not rs.EOF Then
    
    txtEstado.Text = rs!Estado_desc
    txtDivisa.Text = UCase(rs!cod_divisa & "")
    txtMonto.Text = Format(rs!Monto, "Standard")
    txtPlazo.Text = rs!Plazo
    txtTasa.Text = Format(rs!Tasa, "Standard")
    txtSaldo.Text = Format(rs!Saldo, "Standard")
     
    txtNotas.Text = rs!Notas & ""
    
    txtFecFormaliza.Text = Format(rs!Fecha_Formaliza, "dd/mm/yyyy")
    txtFecPrimerPago.Text = Format(rs!Fecha_Primer_Pago & "", "dd/mm/yyyy")
    txtFecProxPago.Text = Format(rs!Fecha_Prox_Pago & "", "dd/mm/yyyy")
   
    txtDiaPago.Text = rs!dia_de_pago & ""
    
    txtAmortizado.Text = Format(rs!Mov_Amortiza, "Standard")
    txtIntereses.Text = Format(rs!Mov_Intereses, "Standard")
    txtComisiones.Text = Format(rs!Mov_Comision, "Standard")
    txtCargos.Text = Format(rs!Mov_Cargos, "Standard")
    
Else
   txtEstado.Text = ""
End If

rs.Close


Call RefrescaTags(Me)

Me.MousePointer = vbDefault

Exit Sub
vError:
    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub


Private Sub btnMovimiento_Click(Index As Integer)
Dim i As Integer

Select Case Index
  Case 0 'Reporte
  

        With frmContenedor.Crt
            .Reset
            .WindowState = crptMaximized
            .WindowShowGroupTree = True
            .WindowShowPrintSetupBtn = True
            .WindowShowRefreshBtn = True
            .WindowShowSearchBtn = True
            .WindowTitle = "Administración de Garantías"
            .Connect = glogon.ConectRPT
            
            .Formulas(0) = "fxUsuario='" & glogon.Usuario & "'"
            .Formulas(1) = "fxFecha='" & Format(fxFechaServidor, "dd/mm/yyyy") & "'"
            .Formulas(2) = "fxEmpresa='" & GLOBALES.gstrNombreEmpresa & "'"
                 
                 
            .ReportFileName = SIFGlobal.fxPathReportes("Acreedores_OperacionMovimientos.rpt")
            .SelectionFormula = "{CRD_APA_OPERACIONES.COD_ACREEDOR} = '" & txtAcreedor.Text _
                            & "' AND {CRD_APA_OPERACIONES.OPERACION} = '" & txtOperacion.Text & "'"
            
            .SubreportToChange = "sbMovimiento"
            .StoredProcParam(0) = txtAcreedor.Text
            .StoredProcParam(1) = txtOperacion.Text
            .PrintReport
        End With
  
  
  Case 1 'Aplicar
        i = MsgBox("Esta seguro que desea realizar este movimiento?", vbYesNo)
        If i = vbYes Then
             Call sbAplicaMovimiento
        End If
End Select


End Sub

Private Sub FlatScrollBar_Change()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

If vScroll Then
    strSQL = "select Top 1 OPERACION from CRD_APA_OPERACIONES"
    
    If FlatScrollBar.Value = 1 Then
       strSQL = strSQL & " where COD_ACREEDOR = '" & txtAcreedor.Text & "' AND OPERACION  > '" & txtOperacion.Text & "'"
       
       If chkSaldo.Value = vbChecked Then
           strSQL = strSQL & " and Saldo > 0"
       End If
       
       strSQL = strSQL & " order by OPERACION asc"
    Else
       strSQL = strSQL & " where COD_ACREEDOR = '" & txtAcreedor.Text & "' AND OPERACION < '" & txtOperacion.Text & "'"
       
       If chkSaldo.Value = vbChecked Then
           strSQL = strSQL & " and Saldo > 0"
       End If
       
       strSQL = strSQL & " order by OPERACION desc"
    End If
    
    Call OpenRecordSet(rs, strSQL)
    If Not rs.EOF And Not rs.BOF Then
      txtOperacion.Text = rs!Operacion
      Call sbConsultaOperacion(rs!Operacion)
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

Private Sub sbLimpiaPantalla(Optional pTotal As Boolean = True)

If pTotal Then
   txtOperacion.Text = ""
   txtEstado.Text = ""
End If

tcMain.Item(0).Selected = True
    
    txtDivisa.Text = ""
    txtMonto.Text = Format(0, "Standard")
    txtPlazo.Text = 0
    txtTasa.Text = 0
    txtSaldo.Text = Format(0, "Standard")
     
    txtNotas.Text = ""
    
    txtFecFormaliza.Text = ""
    txtFecPrimerPago.Text = ""
    txtFecProxPago.Text = ""
   
    txtDiaPago.Text = ""
    
    
    txtAmortizado.Text = 0
    txtIntereses.Text = 0
    txtComisiones.Text = 0
    txtCargos.Text = 0
    
    txtMovAmortiza.Text = 0
    txtMovIntereses.Text = 0
    txtMovComision.Text = 0
    txtMovCargos.Text = 0
    txtMovTotal.Text = 0
    
    txtMovNotas.Text = ""
    
    Call optMov_Click(0)
End Sub


Private Function fxMovimientoValida() As Boolean
Dim vMensaje As String

On Error GoTo vError

vMensaje = ""

If Len(txtMovDocRef.Text) = 0 Then
   vMensaje = vMensaje & vbCrLf & " -  No se ha indicado un documento de referencia"
End If

If Len(txtMovNotas.Text) <= 10 Then
   vMensaje = vMensaje & vbCrLf & " -  No se ha indicado una Nota válida"
End If

If Not IsNumeric(txtMovAmortiza.Text) Then
   vMensaje = vMensaje & vbCrLf & " -  Monto de Amortización NO ES NUMERICO!"
End If
If Not IsNumeric(txtMovIntereses.Text) Then
   vMensaje = vMensaje & vbCrLf & " -  Monto de Intereses NO ES NUMERICO!"
End If
If Not IsNumeric(txtMovComision.Text) Then
   vMensaje = vMensaje & vbCrLf & " -  Monto de Comisión NO ES NUMERICO!"
End If
If Not IsNumeric(txtMovCargos.Text) Then
   vMensaje = vMensaje & vbCrLf & " -  Monto de Cargos NO ES NUMERICO!"
End If

If Len(vMensaje) = 0 Then
    If CCur(txtMovAmortiza.Text) < 0 Then
          vMensaje = vMensaje & vbCrLf & " -  Monto de Amortización no puede ser negativo!"
    End If
    If CCur(txtMovIntereses.Text) < 0 Then
          vMensaje = vMensaje & vbCrLf & " -  Monto de Intereses no puede ser negativo!"
    End If
    If CCur(txtMovComision.Text) < 0 Then
          vMensaje = vMensaje & vbCrLf & " -  Monto de Comisión no puede ser negativo!"
    End If
    If CCur(txtMovCargos.Text) < 0 Then
          vMensaje = vMensaje & vbCrLf & " -  Monto de Cargos no puede ser negativo!"
    End If
    
    If CCur(txtMovTotal.Text) = 0 Then
      vMensaje = vMensaje & vbCrLf & " -  No se ha indicado un Monto para el movimiento!"
    End If
   
End If


If Len(vMensaje) > 0 Then
   MsgBox vMensaje, vbExclamation
   fxMovimientoValida = False
Else
   fxMovimientoValida = True
End If

Exit Function

vError:
   MsgBox fxSys_Error_Handler(Err.Description), vbCritical
   fxMovimientoValida = False


End Function


Private Sub sbAplicaMovimiento()
Dim strSQL As String, rs As New ADODB.Recordset
Dim vTipo As String, vCuenta As String


On Error GoTo vError

vCuenta = fxgCntCuentaFormato(False, txtCuenta.Text, 0)
If optMov.Item(0).Value Then
   vTipo = "E"
Else
   vTipo = "N"
End If

'VALIDAR DATOS:
If Not fxMovimientoValida Then
   Exit Sub
End If

Me.MousePointer = vbHourglass


strSQL = "exec spAPA_Movimiento '" & txtAcreedor.Text & "','" & txtOperacion.Text & "','" & glogon.Usuario _
       & "','" & vTipo & "'," & CCur(txtMovAmortiza.Text) & "," & CCur(txtMovIntereses.Text) & "," & CCur(txtMovComision) _
       & "," & CCur(txtMovCargos.Text) & ",'" & txtMovNotas.Text & "','" & vCuenta & "','" & txtMovDocRef.Text & "'"
Call OpenRecordSet(rs, strSQL)
If Not rs.EOF And Not rs.BOF Then
  Call sbImprimeRecibo(rs!Cod_Transaccion, rs!Tipo_Documento)
End If
rs.Close

Me.MousePointer = vbDefault
MsgBox "Movimiento Realizado Satisfactoriamente!", vbInformation

Call sbConsultaOperacion(txtOperacion.Text)

Exit Sub

vError:
   Me.MousePointer = vbDefault
   MsgBox fxSys_Error_Handler(Err.Description), vbCritical
   

End Sub


Private Sub sbCuentaCarga()
Dim strSQL As String, rs As New ADODB.Recordset

strSQL = "select D.COD_CUENTA, isnull(C.COD_CUENTA_MASK,C.COD_CUENTA) AS 'CUENTA_MASK', C.DESCRIPCION" _
       & " From SIF_DOCUMENTOS D left join CntX_Cuentas C on C.cod_Contabilidad = " & GLOBALES.gEnlace _
       & " and D.cod_cuenta = C.cod_Cuenta" _
       & " where D.TIPO_DOCUMENTO = 'APA'"
Call OpenRecordSet(rs, strSQL)
If Not rs.EOF And Not rs.BOF Then
   txtCuenta.Text = rs!CUENTA_MASK & ""
   txtCuentaDesc.Text = rs!Descripcion & ""
End If
rs.Close

End Sub

Private Sub Form_Load()
vModulo = 14

 vScroll = False
 FlatScrollBar.Value = 0
 vScroll = True

 Call sbLimpiaPantalla
 Call sbCuentaCarga
 
' Call Formularios(Me)
 Call RefrescaTags(Me)
End Sub

Private Sub sbConsultaDetalle()
Dim strSQL As String

On Error GoTo vError

strSQL = "exec spAPA_ConsultaOperacionDetalle '" & txtAcreedor.Text & "','" & txtOperacion.Text & "'"
Call sbCargaGrid(vGrid, 16, strSQL)

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical


End Sub

Private Sub optMov_Click(Index As Integer)

Select Case True
  Case optMov.Item(0).Value 'Abono
     txtMovAmortiza.ForeColor = vbBlack
  Case optMov.Item(1).Value 'Anulacion
     txtMovAmortiza.ForeColor = vbRed
End Select

txtMovIntereses.ForeColor = txtMovAmortiza.ForeColor
txtMovComision.ForeColor = txtMovAmortiza.ForeColor
txtMovCargos.ForeColor = txtMovAmortiza.ForeColor
txtMovTotal.ForeColor = txtMovAmortiza.ForeColor

End Sub



Private Sub tcMain_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
If txtEstado.Text = "" Or txtAcreedorDesc.Text = "" Then
   tcMain.Item(0).Selected = True

   If txtAcreedorDesc.Text = "" Then
'     txtAcreedor.SetFocus
   End If
End If

If Item.Index = 1 Then
    Call sbConsultaDetalle
End If

End Sub

Private Sub txtAcreedor_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtOperacion.SetFocus

If KeyCode = vbKeyF4 Then
    gBusquedas.Columna = "COD_ACREEDOR"
    gBusquedas.Orden = "COD_ACREEDOR"
    gBusquedas.Consulta = "SELECT COD_ACREEDOR AS 'ACREEDOR', DESCRIPCION FROM CRD_APA_ACREEDORES"
    gBusquedas.Filtro = ""
    frmBusquedas.Show vbModal
    
    If gBusquedas.Resultado <> "" Then
       txtAcreedor.Text = gBusquedas.Resultado
       Call sbConsulta(gBusquedas.Resultado)
    End If
End If


End Sub

Private Sub txtAcreedor_LostFocus()
    Call sbConsulta(txtAcreedor.Text)
End Sub

Private Sub txtAcreedorDesc_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtOperacion.SetFocus

If KeyCode = vbKeyF4 Then
    gBusquedas.Columna = "DESCRIPCION"
    gBusquedas.Orden = "DESCRIPCION"
    gBusquedas.Consulta = "SELECT COD_ACREEDOR AS 'ACREEDOR', DESCRIPCION FROM CRD_APA_ACREEDORES"
    gBusquedas.Filtro = ""
    frmBusquedas.Show vbModal
    
    If gBusquedas.Resultado <> "" Then
       txtAcreedor.Text = gBusquedas.Resultado
       Call sbConsulta(gBusquedas.Resultado)
    End If
End If
End Sub




Private Sub txtCuenta_KeyDown(KeyCode As Integer, Shift As Integer)

On Error GoTo vError

If KeyCode = vbKeyF4 Then
    frmCntX_ConsultaCuentas.Show vbModal
    txtCuenta.Text = gCuenta
    txtCuentaDesc.Text = fxgCntCuentaDesc(gCuenta)
End If

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub txtCuentaDesc_KeyDown(KeyCode As Integer, Shift As Integer)

On Error GoTo vError

If KeyCode = vbKeyF4 Then
    frmCntX_ConsultaCuentas.Show vbModal
    txtCuenta.Text = gCuenta
    txtCuentaDesc.Text = fxgCntCuentaDesc(gCuenta)
End If

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub txtMovAmortiza_GotFocus()
On Error GoTo vError
    txtMovAmortiza.Text = CCur(txtMovAmortiza.Text)
vError:
End Sub

Private Sub txtMovAmortiza_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtMovIntereses.SetFocus

End Sub

Private Sub txtMovAmortiza_KeyUp(KeyCode As Integer, Shift As Integer)
Call sbCalculaTotalMov
End Sub

Private Sub txtMovAmortiza_LostFocus()
On Error GoTo vError
    txtMovAmortiza.Text = Format(CCur(txtMovAmortiza.Text), "Standard")
vError:
End Sub

Private Sub txtMovIntereses_GotFocus()
On Error GoTo vError
    txtMovIntereses.Text = CCur(txtMovIntereses.Text)
vError:
End Sub

Private Sub txtMovIntereses_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtMovComision.SetFocus

End Sub

Private Sub txtMovIntereses_KeyUp(KeyCode As Integer, Shift As Integer)
Call sbCalculaTotalMov
End Sub

Private Sub txtMovIntereses_LostFocus()
On Error GoTo vError
    txtMovIntereses.Text = Format(CCur(txtMovIntereses.Text), "Standard")
vError:
End Sub


Private Sub txtMovComision_GotFocus()
On Error GoTo vError
    txtMovComision.Text = CCur(txtMovComision.Text)
vError:
End Sub

Private Sub txtMovComision_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtMovCargos.SetFocus

End Sub

Private Sub txtMovComision_KeyUp(KeyCode As Integer, Shift As Integer)
Call sbCalculaTotalMov
End Sub

Private Sub txtMovComision_LostFocus()
On Error GoTo vError
    txtMovComision.Text = Format(CCur(txtMovComision.Text), "Standard")
vError:
End Sub



Private Sub txtMovCargos_GotFocus()
On Error GoTo vError
    txtMovCargos.Text = CCur(txtMovCargos.Text)
vError:
End Sub

Private Sub txtMovCargos_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtMovNotas.SetFocus
End Sub

Private Sub txtMovCargos_KeyUp(KeyCode As Integer, Shift As Integer)
Call sbCalculaTotalMov
End Sub

Private Sub txtMovCargos_LostFocus()
On Error GoTo vError
    txtMovCargos.Text = Format(CCur(txtMovCargos.Text), "Standard")
vError:
End Sub



Private Sub sbCalculaTotalMov()

On Error GoTo vError

txtMovTotal.Text = Format(CCur(txtMovAmortiza.Text) + CCur(txtMovIntereses.Text) _
                    + CCur(txtMovComision.Text) + CCur(txtMovCargos.Text), "Standard")

vError:

End Sub


Private Sub txtOperacion_KeyDown(KeyCode As Integer, Shift As Integer)

If txtAcreedorDesc.Text = "" Then
   MsgBox "Debe indicar un Acreedor antes que la operación!", vbExclamation
   txtAcreedor.SetFocus
   Exit Sub
End If

If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then
       Call sbConsultaOperacion(txtOperacion.Text)
       txtEstado.SetFocus
End If

If KeyCode = vbKeyF4 Then
    gBusquedas.Columna = "OPERACION"
    gBusquedas.Orden = "OPERACION"
    gBusquedas.Consulta = "SELECT OPERACION,COD_ACREEDOR AS 'ACREEDOR', MONTO, SALDO, FECHA_FORMALIZA AS 'FORMALIZA'" _
                        & " FROM crd_apa_operaciones"
    gBusquedas.Filtro = " AND COD_ACREEDOR = '" & txtAcreedor.Text & "'"
    frmBusquedas.Show vbModal
    
    If gBusquedas.Resultado <> "" Then
       txtOperacion.Text = gBusquedas.Resultado
       Call sbConsultaOperacion(gBusquedas.Resultado)
    End If
End If


End Sub
