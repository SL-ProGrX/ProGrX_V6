VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.0#0"; "Codejock.Controls.v22.0.0.ocx"
Begin VB.Form frmCR_Configuracion 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Parámetros de Créditos"
   ClientHeight    =   8880
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10320
   HelpContextID   =   3008
   Icon            =   "frmCR_Configuracion.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8880
   ScaleWidth      =   10320
   Begin VB.Timer TimerX 
      Interval        =   10
      Left            =   9240
      Top             =   480
   End
   Begin XtremeSuiteControls.TabControl tcMain 
      Height          =   7572
      Left            =   120
      TabIndex        =   1
      Top             =   1200
      Width           =   10092
      _Version        =   1441792
      _ExtentX        =   17801
      _ExtentY        =   13356
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
      Item(0).Caption =   "Generales"
      Item(0).ControlCount=   1
      Item(0).Control(0)=   "vGrid"
      Item(1).Caption =   "Operativos"
      Item(1).ControlCount=   4
      Item(1).Control(0)=   "cmdGuardar"
      Item(1).Control(1)=   "GroupBox1(0)"
      Item(1).Control(2)=   "GroupBox1(1)"
      Item(1).Control(3)=   "GroupBox1(2)"
      Begin XtremeSuiteControls.GroupBox GroupBox1 
         Height          =   1332
         Index           =   0
         Left            =   -69760
         TabIndex        =   4
         Top             =   480
         Visible         =   0   'False
         Width           =   9612
         _Version        =   1441792
         _ExtentX        =   16954
         _ExtentY        =   2350
         _StockProps     =   79
         Caption         =   "Fecha de Cierre para Cálculo de Intereses (días) Operaciones en Tramite / Políticas"
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
         BorderStyle     =   1
         Begin XtremeSuiteControls.PushButton cmdCambiaFecha 
            Height          =   312
            Left            =   3360
            TabIndex        =   28
            Top             =   480
            Width           =   492
            _Version        =   1441792
            _ExtentX        =   868
            _ExtentY        =   556
            _StockProps     =   79
            BackColor       =   -2147483633
            Transparent     =   -1  'True
            UseVisualStyle  =   -1  'True
            Appearance      =   2
            Picture         =   "frmCR_Configuracion.frx":000C
         End
         Begin XtremeSuiteControls.FlatEdit txtPorcentajeSobreAhorros 
            Height          =   312
            Left            =   6960
            TabIndex        =   25
            Top             =   480
            Visible         =   0   'False
            Width           =   852
            _Version        =   1441792
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
            Appearance      =   2
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtTBP 
            Height          =   312
            Left            =   6960
            TabIndex        =   26
            Top             =   960
            Width           =   852
            _Version        =   1441792
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
            Appearance      =   2
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.DateTimePicker dtpFechaCorte 
            Height          =   312
            Left            =   1920
            TabIndex        =   27
            Top             =   480
            Width           =   1332
            _Version        =   1441792
            _ExtentX        =   2350
            _ExtentY        =   556
            _StockProps     =   68
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CustomFormat    =   "dd/MM/yyyyy"
            Format          =   3
         End
         Begin XtremeSuiteControls.PushButton cmdCambiaTBP 
            Height          =   312
            Left            =   7920
            TabIndex        =   29
            Top             =   960
            Width           =   492
            _Version        =   1441792
            _ExtentX        =   868
            _ExtentY        =   556
            _StockProps     =   79
            BackColor       =   -2147483633
            Transparent     =   -1  'True
            UseVisualStyle  =   -1  'True
            Appearance      =   2
            Picture         =   "frmCR_Configuracion.frx":073D
         End
         Begin XtremeSuiteControls.PushButton cmdPoliticaPago 
            Height          =   312
            Left            =   3360
            TabIndex        =   30
            Top             =   960
            Width           =   492
            _Version        =   1441792
            _ExtentX        =   868
            _ExtentY        =   556
            _StockProps     =   79
            BackColor       =   -2147483633
            Transparent     =   -1  'True
            UseVisualStyle  =   -1  'True
            Appearance      =   2
            Picture         =   "frmCR_Configuracion.frx":0E6E
         End
         Begin XtremeSuiteControls.PushButton cmdCambiaPorcentaje 
            Height          =   312
            Left            =   7920
            TabIndex        =   31
            Top             =   480
            Width           =   492
            _Version        =   1441792
            _ExtentX        =   868
            _ExtentY        =   556
            _StockProps     =   79
            BackColor       =   -2147483633
            Transparent     =   -1  'True
            UseVisualStyle  =   -1  'True
            Appearance      =   2
            Picture         =   "frmCR_Configuracion.frx":1561
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            Caption         =   "Fecha de Corte"
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
            Height          =   312
            Index           =   0
            Left            =   600
            TabIndex        =   10
            Top             =   480
            Width           =   1452
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            Caption         =   "Tasa Básica Pasiva Actual"
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
            Height          =   312
            Index           =   2
            Left            =   4800
            TabIndex        =   9
            Top             =   960
            Width           =   2772
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            Caption         =   "[ % ] Garantía Patrimonial"
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
            Height          =   312
            Index           =   1
            Left            =   4800
            TabIndex        =   8
            Top             =   480
            Width           =   3012
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            Caption         =   "Política de Pago de Créditos"
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
            Height          =   312
            Index           =   4
            Left            =   600
            TabIndex        =   7
            Top             =   960
            Width           =   2652
         End
      End
      Begin FPSpreadADO.fpSpread vGrid 
         Height          =   6855
         Left            =   240
         TabIndex        =   2
         Top             =   480
         Width           =   9615
         _Version        =   524288
         _ExtentX        =   16960
         _ExtentY        =   12091
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
         SpreadDesigner  =   "frmCR_Configuracion.frx":1C54
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin XtremeSuiteControls.PushButton cmdGuardar 
         Height          =   492
         Left            =   -61720
         TabIndex        =   3
         Top             =   6960
         Visible         =   0   'False
         Width           =   1452
         _Version        =   1441792
         _ExtentX        =   2561
         _ExtentY        =   868
         _StockProps     =   79
         Caption         =   "Guardar"
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
         Appearance      =   16
         Picture         =   "frmCR_Configuracion.frx":21E1
      End
      Begin XtremeSuiteControls.GroupBox GroupBox1 
         Height          =   1092
         Index           =   1
         Left            =   -69760
         TabIndex        =   5
         Top             =   2040
         Visible         =   0   'False
         Width           =   9612
         _Version        =   1441792
         _ExtentX        =   16954
         _ExtentY        =   1926
         _StockProps     =   79
         Caption         =   "Cuenta para Abono en Formalización de Credito de la Poliza Saldo Deudores [PSD]"
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
         BorderStyle     =   1
         Begin XtremeSuiteControls.FlatEdit txtPolizaCta 
            Height          =   312
            Left            =   1920
            TabIndex        =   13
            Top             =   360
            Width           =   1932
            _Version        =   1441792
            _ExtentX        =   3408
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
         Begin XtremeSuiteControls.FlatEdit txtPSDMonto 
            Height          =   312
            Left            =   1920
            TabIndex        =   23
            Top             =   720
            Width           =   1932
            _Version        =   1441792
            _ExtentX        =   3408
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
            Appearance      =   2
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtPolizaCtaDesc 
            Height          =   312
            Left            =   3840
            TabIndex        =   14
            Top             =   360
            Width           =   5052
            _Version        =   1441792
            _ExtentX        =   8911
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
         Begin XtremeSuiteControls.FlatEdit FlatEdit1 
            Height          =   312
            Left            =   3840
            TabIndex        =   37
            Top             =   720
            Width           =   5052
            _Version        =   1441792
            _ExtentX        =   8911
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
            Text            =   "Monto sobre millón aprobado (Para Cálculo PSD)"
            Locked          =   -1  'True
            Appearance      =   2
            UseVisualStyle  =   0   'False
         End
         Begin VB.Label Label4 
            Caption         =   "Cta - Polizas (PSD)"
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
            Index           =   3
            Left            =   240
            TabIndex        =   12
            Top             =   360
            Width           =   1692
         End
         Begin VB.Label Label4 
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
            Height          =   252
            Index           =   13
            Left            =   240
            TabIndex        =   11
            Top             =   720
            Width           =   1572
         End
      End
      Begin XtremeSuiteControls.GroupBox GroupBox1 
         Height          =   3492
         Index           =   2
         Left            =   -69760
         TabIndex        =   6
         Top             =   3360
         Visible         =   0   'False
         Width           =   9612
         _Version        =   1441792
         _ExtentX        =   16954
         _ExtentY        =   6159
         _StockProps     =   79
         Caption         =   "Desembolsos de Créditos"
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
         BorderStyle     =   1
         Begin XtremeSuiteControls.CheckBox chkBanRegla 
            Height          =   252
            Left            =   3840
            TabIndex        =   36
            Top             =   3000
            Width           =   5052
            _Version        =   1441792
            _ExtentX        =   8911
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Aplicar Regla para Bancos sin formula continua?   "
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
         Begin XtremeSuiteControls.FlatEdit txtDesembolsosCta 
            Height          =   312
            Left            =   1920
            TabIndex        =   16
            Top             =   360
            Width           =   1932
            _Version        =   1441792
            _ExtentX        =   3408
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
         Begin XtremeSuiteControls.FlatEdit txtMontoMargen 
            Height          =   312
            Left            =   6960
            TabIndex        =   24
            Top             =   960
            Width           =   1932
            _Version        =   1441792
            _ExtentX        =   3408
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
            Appearance      =   2
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.ComboBox cboBanco 
            Height          =   312
            Left            =   2760
            TabIndex        =   32
            Top             =   1800
            Width           =   4212
            _Version        =   1441792
            _ExtentX        =   7435
            _ExtentY        =   582
            _StockProps     =   77
            ForeColor       =   1973790
            BackColor       =   16185078
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColor       =   16185078
            Style           =   2
            Appearance      =   16
            Text            =   "ComboBox1"
         End
         Begin XtremeSuiteControls.ComboBox cboTipoDoc 
            Height          =   312
            Left            =   6960
            TabIndex        =   33
            Top             =   1800
            Width           =   1932
            _Version        =   1441792
            _ExtentX        =   3413
            _ExtentY        =   582
            _StockProps     =   77
            ForeColor       =   1973790
            BackColor       =   16185078
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColor       =   16185078
            Style           =   2
            Appearance      =   16
            Text            =   "ComboBox1"
         End
         Begin XtremeSuiteControls.ComboBox cboMenBan 
            Height          =   312
            Left            =   2760
            TabIndex        =   34
            Top             =   2640
            Width           =   4212
            _Version        =   1441792
            _ExtentX        =   7435
            _ExtentY        =   582
            _StockProps     =   77
            ForeColor       =   1973790
            BackColor       =   16185078
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColor       =   16185078
            Style           =   2
            Appearance      =   16
            Text            =   "ComboBox1"
         End
         Begin XtremeSuiteControls.ComboBox cboMenTipo 
            Height          =   312
            Left            =   6960
            TabIndex        =   35
            Top             =   2640
            Width           =   1932
            _Version        =   1441792
            _ExtentX        =   3413
            _ExtentY        =   582
            _StockProps     =   77
            ForeColor       =   1973790
            BackColor       =   16185078
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColor       =   16185078
            Style           =   2
            Appearance      =   16
            Text            =   "ComboBox1"
         End
         Begin XtremeSuiteControls.FlatEdit txtDesembolsosCtaDesc 
            Height          =   312
            Left            =   3840
            TabIndex        =   17
            Top             =   360
            Width           =   5052
            _Version        =   1441792
            _ExtentX        =   8911
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
         Begin VB.Label Label4 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Banco para Casos Mayores a Monto a [Monto Cambio]"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   252
            Index           =   9
            Left            =   0
            TabIndex        =   22
            Top             =   1440
            Width           =   4932
         End
         Begin VB.Label Label4 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Banco para casos Menores o Iguales al [Monto Cambio]"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   252
            Index           =   10
            Left            =   0
            TabIndex        =   21
            Top             =   2280
            Width           =   4932
         End
         Begin VB.Label Label4 
            Caption         =   "Banco/Tipo Doc."
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
            Index           =   11
            Left            =   1080
            TabIndex        =   20
            Top             =   2640
            Width           =   1452
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BackStyle       =   0  'Transparent
            Caption         =   "Monto para Cambio de Banco, según Monto a Desembolsar del Crédito"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   312
            Left            =   840
            TabIndex        =   19
            Top             =   960
            Width           =   6012
         End
         Begin VB.Label Label4 
            Caption         =   "Banco/Tipo Doc."
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
            Left            =   1080
            TabIndex        =   18
            Top             =   1800
            Width           =   1452
         End
         Begin VB.Label Label4 
            Caption         =   "Cuenta"
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
            Left            =   240
            TabIndex        =   15
            Top             =   360
            Width           =   1452
         End
      End
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Parámetros de Créditos"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   16.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   480
      Index           =   5
      Left            =   1920
      TabIndex        =   0
      Top             =   300
      Width           =   7332
   End
   Begin VB.Image imgBanner 
      Height          =   1092
      Left            =   0
      Top             =   0
      Width           =   10812
   End
End
Attribute VB_Name = "frmCR_Configuracion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCambiaFecha_Click()
Dim strSQL As String

On Error GoTo vError

'Anterior / Por Compatibilidad
strSQL = "update par_ahcr set cr_fecha_calculo = '" & Format(dtpFechaCorte.Value, "yyyy/mm/dd") & "'"
Call ConectionExecute(strSQL)

'Nuevo
Call sbGuardaParametro("09", Format(dtpFechaCorte.Value, "yyyy/mm/dd"))


Call Bitacora("Modifica", "Fecha Calculo Intereses : " & Format(dtpFechaCorte.Value, "yyyy/mm/dd"))
 
MsgBox "La Fecha de Corte Se Cambió Satisfactoriamente", vbInformation


Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub cmdCambiaPorcentaje_Click()
'Dim strSQL As String
'
'On Error GoTo vError
'
'If IsNumeric(txtPorcentajeSobreAhorros) Then
' If txtPorcentajeSobreAhorros <= 200 Then
'   'Anterior / Por Compatibilidad
'   strSQL = "update par_ahcr set CR_POR_AHORRO = " & txtPorcentajeSobreAhorros.Text
'   Call ConectionExecute(strSQL)
'
'    'Nuevo
'    Call sbGuardaParametro("08", txtPorcentajeSobreAhorros.Text)
'
'   Call Bitacora("Modifica", "% Disponible Sobre Ahorros a " & txtPorcentajeSobreAhorros & "%")
'   MsgBox "Actualización Realizada...", vbInformation
'
' End If
'End If
'
'
'Exit Sub
'
'vError:
' MsgBox fxSys_Error_Handler(Err.Description), vbCritical

frmCR_GarantiasPatrimoniales.Show vbModal

End Sub

Private Sub cmdCambiaTBP_Click()
Dim strSQL As String

On Error GoTo vError

If IsNumeric(txtTBP) Then
 If txtTBP <= 100 Then
 
    'Anterior / Compatibilidad
    strSQL = "update par_ahcr set CR_tbp = " & txtTBP
    Call ConectionExecute(strSQL)
   
    'Nuevo
    Call sbGuardaParametro("07", txtTBP.Text, "DEC")
   
   Call Bitacora("Modifica", "% Tasa Básica Pasiva " & txtTBP)
   MsgBox "Actualización Realizada...", vbInformation
 
 End If

End If

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub cmdPoliticaPago_Click()
 frmCR_PoliticaPago.Show vbModal
End Sub

Private Sub Form_Activate()
vModulo = 3
End Sub

Private Sub sbInicializa()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

Me.MousePointer = vbHourglass

tcMain.Item(0).Selected = True

strSQL = "select * from crd_parametros" _
      & " order by cod_parametro"
Call sbCargaGridLocal(vGrid, 3, strSQL)

'Carga Combos de Cuentas Bancarias para Desembolsos de Creditos
strSQL = "Select id_banco as 'IdX',rtrim(Descripcion) as 'ItmX' From Tes_Bancos where Estado = 'A'"

Call sbCbo_Llena_New(cboBanco, strSQL, False, True)
Call sbCbo_Llena_New(cboMenBan, strSQL, False, True)


With cboTipoDoc
  .Clear
  .AddItem "01 - Cheque"
  .AddItem "02 - Transferencia"
End With
  
With cboMenTipo
  .Clear
  .AddItem "01 - Cheque"
  .AddItem "02 - Transferencia"
End With


'Inicializa Parametros
strSQL = "exec spCrdParametros"
Call ConectionExecute(strSQL)


'Se asume que existe un registro en par_ahcr
strSQL = "select * from par_ahcr"
 
Call OpenRecordSet(rs, strSQL)
  dtpFechaCorte.Value = IIf(IsNull(rs!cr_fecha_calculo), Date, rs!cr_fecha_calculo)
  txtPorcentajeSobreAhorros.Text = Format(IIf(IsNull(rs!Cr_Por_ahorro), 100, rs!Cr_Por_ahorro), "Standard")
  txtTBP.Text = Format(IIf(IsNull(rs!cr_tbp), 15, rs!cr_tbp), "Standard")
  
  txtDesembolsosCta.Text = IIf(IsNull(rs!cr_cta_desembolso), "", rs!cr_cta_desembolso)
  txtPolizaCta.Text = IIf(IsNull(rs!cr_cta_polizas), "", rs!cr_cta_polizas)
  
  txtPSDMonto.Text = Format(IIf(IsNull(rs!cr_PsdMnt), 0, rs!cr_PsdMnt), "Standard")
 
  Call sbCboAsignaDato(cboMenBan, Trim(fxDescribeBanco(rs!Cod_Banco_men)), True, rs!Cod_Banco_men)
  Call sbCboAsignaDato(cboBanco, Trim(fxDescribeBanco(rs!cod_banco)), True, rs!cod_banco)
    
  If rs!TipoDoc = "CK" Then
     cboTipoDoc.Text = "01 - Cheque"
  Else
    cboTipoDoc.Text = "02 - Transferencia"
  End If
 
 chkBanRegla.Value = rs!regla_banco
 txtMontoMargen.Text = Format(rs!regla_monto, "Standard")
 
  If rs!cod_tipo_men = "CK" Then
     cboMenTipo.Text = "01 - Cheque"
  Else
    cboMenTipo.Text = "02 - Transferencia"
  End If
    
rs.Close

  
Call CargaLblsDatosMED(txtDesembolsosCta, txtDesembolsosCtaDesc)
Call CargaLblsDatosMED(txtPolizaCta, txtPolizaCtaDesc)

Me.MousePointer = vbDefault

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbExclamation

End Sub


Private Sub Form_Load()


On Error GoTo vError

vModulo = 3

vGrid.AppearanceStyle = fxGridStyle

Set imgBanner.Picture = frmContenedor.imgBanner_Mantenimiento.Picture

Call Formularios(Me)
Call RefrescaTags(Me)

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbExclamation


End Sub

Private Sub CargaLblsDatosMED(pCta As Object, pCtaDesc As Object)
 pCtaDesc.Text = fxgCntCuentaDesc(fxgCntCuentaFormato(False, pCta.Text, 0))
End Sub

Function ValidaDatos(pCta As Object, pCtaDesc As Object) As Boolean

If Not fxgCntCuentaValida(fxgCntCuentaFormato(False, pCta.Text, 0)) Then
     ValidaDatos = False
     MsgBox "No se encontró código de cuenta especificado, Corrijalo...", vbCritical
Else
     ValidaDatos = True
     pCtaDesc.Caption = fxgCntCuentaDesc(fxgCntCuentaFormato(False, pCta.Text, 0))
End If
End Function

Function ValidaDatosGrabar(pCta As Object) As Boolean

ValidaDatosGrabar = fxgCntCuentaValida(fxgCntCuentaFormato(False, pCta.Text, 0))

End Function


Private Sub cmdGuardar_Click()
Dim Grabar As String, strSQL As String

On Error GoTo vError

Grabar = "S"
'valida primero todas las casillas
Select Case False
 Case ValidaDatosGrabar(txtDesembolsosCta)
    Grabar = "N"
 Case ValidaDatosGrabar(txtPolizaCta)
    Grabar = "N"
End Select

If Grabar = "S" Then
    'Anterior por Compatibilidad
    strSQL = "update par_ahcr set " _
           & "CR_cta_desembolso = '" & txtDesembolsosCta.Text & "'," _
           & "CR_cta_polizas = '" & txtPolizaCta.Text & "',CR_PSDMNT = " & CCur(txtPSDMonto) & "," _
           & "cod_banco = " & cboBanco.ItemData(cboBanco.ListIndex) & "," _
           & "tipodoc = '" & IIf(Mid(cboTipoDoc, 1, 2) = "01", "CK", "TE") & "'," _
           & "cod_banco_men = " & cboMenBan.ItemData(cboMenBan.ListIndex) & "," _
           & "cod_tipo_men = '" & IIf(Mid(cboMenTipo.Text, 1, 2) = "01", "CK", "TE") & "'," _
           & "regla_banco = " & chkBanRegla.Value & ",regla_monto = " & CCur(txtMontoMargen)
     Call ConectionExecute(strSQL)
     
     Call Bitacora("Modifica", "Parámetros de Créditos : Cuentas y Desembolsos")

     Call CargaLblsDatosMED(txtDesembolsosCta, txtDesembolsosCtaDesc)
     Call CargaLblsDatosMED(txtPolizaCta, txtPolizaCtaDesc)
    
    
    'Nuevo
    Call sbGuardaParametro("03", txtDesembolsosCta.Text)
    Call sbGuardaParametro("04", txtPolizaCta.Text)
    Call sbGuardaParametro("06", CStr(CCur(txtPSDMonto.Text)))
    Call sbGuardaParametro("10", CStr(CCur(txtMontoMargen)))
    Call sbGuardaParametro("11", CStr(cboBanco.ItemData(cboBanco.ListIndex)))
    Call sbGuardaParametro("12", IIf(Mid(cboTipoDoc, 1, 2) = "01", "CK", "TE"))
    Call sbGuardaParametro("13", CStr(cboMenBan.ItemData(cboMenBan.ListIndex)))
    Call sbGuardaParametro("14", IIf(Mid(cboMenTipo.Text, 1, 2) = "01", "CK", "TE"))
    Call sbGuardaParametro("15", CStr(chkBanRegla.Value))
    
    MsgBox "La Información se guardó satisfactoriamente ...", vbInformation

Else

 MsgBox "No se puede guardar la información, Verifique las cuentas ingresadas...", vbInformation

End If

Call RefrescaTags(Me)
Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbExclamation


End Sub


Private Sub Form_Unload(Cancel As Integer)
 Unload Me
End Sub

Private Sub sbBusqueda(Index As Integer)

On Error GoTo vError

Call sbgCntCuentaConsulta

If gBusquedas.Resultado <> "" Then
    Select Case Index
     Case 1
         txtDesembolsosCta.Text = gBusquedas.Resultado
         txtDesembolsosCta.SetFocus
     Case 3
         txtPolizaCta.Text = gBusquedas.Resultado
         txtPolizaCta.SetFocus
    End Select
End If

Exit Sub
vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbExclamation

End Sub


Private Sub TimerX_Timer()

TimerX.Interval = 0
TimerX.Enabled = False

Call sbInicializa

End Sub

Private Sub txtDesembolsosCta_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next

If KeyCode = vbKeyF4 Then Call sbBusqueda(1)

If KeyCode = vbKeyReturn Then txtDesembolsosCtaDesc.SetFocus

End Sub

Private Sub txtPolizaCta_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = vbKeyF4 Then Call sbBusqueda(3)
If KeyCode = vbKeyReturn Then If ValidaDatos(txtPolizaCta, txtPolizaCtaDesc) Then txtPSDMonto.SetFocus
End Sub


Public Sub sbCargaGridLocal(pGrid As Object, MaxCol As Integer, strSQL As String)
Dim rs As New ADODB.Recordset, i As Integer

With pGrid
    .MaxRows = 0
    .MaxCols = MaxCol
    Call OpenRecordSet(rs, strSQL, 0)
    Do While Not rs.EOF
      .MaxRows = .MaxRows + 1
      .Row = .MaxRows
      
     
      For i = 1 To 3
        .col = i
        Select Case i
          Case 1 'Codigo
            .CellTag = rs!Tipo & ""
            .Text = rs!Cod_Parametro
            .CellNote = "Modificado Por: " & rs!MODIFICA_USUARIO & vbCrLf & "Fecha: " & rs!MODIFICA_FECHA
          
            .TextTip = TextTipFixed
            .TextTipDelay = 1000
          
          Case 2 'Descripcion
            .Text = rs!Descripcion
            .CellNote = rs!Notas & ""
            .TextTip = TextTipFixed
            .TextTipDelay = 1000
          
          Case 3 'Valor
            If UCase(Trim(rs!Tipo)) = "CTA" Then
                .TextTip = TextTipFixed
                .TextTipDelay = 1000
                .CellNoteIndicatorColor = vbBlue
                .CellNoteIndicatorShape = CellNoteIndicatorShapeTriangle
                
                .Text = fxgCntCuentaFormato(True, Trim(rs!Valor), 0)
                .CellNote = fxgCntCuentaDesc(Trim(rs!Valor))
            Else
                .Text = rs!Valor
            End If
            
        End Select
      Next i
      rs.MoveNext
    Loop
    rs.Close

End With

End Sub



Private Sub txtMontoMargen_GotFocus()
On Error GoTo vError
    txtMontoMargen.Text = CCur(txtMontoMargen.Text)
vError:
End Sub

Private Sub txtMontoMargen_LostFocus()
On Error GoTo vError
    txtMontoMargen.Text = Format(CCur(txtMontoMargen.Text), "Standard")
vError:
End Sub

Private Sub txtPSDMonto_GotFocus()
On Error GoTo vError
    txtPSDMonto.Text = CCur(txtPSDMonto.Text)
vError:
End Sub

Private Sub txtPSDMonto_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtDesembolsosCta.SetFocus
End Sub

Private Sub txtPSDMonto_LostFocus()
On Error GoTo vError
    txtPSDMonto.Text = Format(CCur(txtPSDMonto.Text), "Standard")
vError:
End Sub


Private Sub sbGuardaParametro(pParametro As String, pValor As String _
                    , Optional pTipo As String = "DEC")

Dim strSQL As String, rs As New ADODB.Recordset
Dim Validacion As Boolean, vMensaje As String

On Error GoTo vError

Validacion = True
vMensaje = ""

Select Case Trim(pTipo)
  Case "DEC" 'Decimal
    If IsNumeric(pValor) Then
       pValor = CCur(pValor)
    Else
       Validacion = False
       vMensaje = "El valor indicado no es válido...!!!"
    End If
    
  Case "NUM" 'Número Entero
    If IsNumeric(pValor) Then
       pValor = CLng(pValor)
    Else
       Validacion = False
       vMensaje = "El valor indicado no es válido...!!!"
    End If
  
  Case "POR" 'Porcentaje
    If IsNumeric(pValor) Then
       pValor = CCur(pValor)
    Else
       Validacion = False
       vMensaje = "El valor indicado no es válido, suministre un porcentaje ..!!!"
    End If
  
  Case "CTA" 'Cuenta Contable
    Validacion = fxgCntCuentaValida(fxgCntCuentaFormato(False, pValor))
    If Not Validacion Then
        vMensaje = "La Cuenta indicada no es válida, presiones F4 para buscar en el catálogo...!!!"
    Else
        pValor = fxgCntCuentaFormato(False, pValor)
    End If
    
  Case "CHR" 'Caracteres
    If InStr(1, pValor, "'", vbTextCompare) > 0 Then
       Validacion = False
       vMensaje = "El valor indicado contiene caracteres no válidos...!!!"
    End If
    
  Case "PSN" 'Pregunta S ó N
     If UCase(Mid(pValor, 1, 1)) = "S" Or UCase(Mid(pValor, 1, 1)) = "N" Then
       pValor = UCase(Mid(pValor, 1, 1))
     Else
       Validacion = False
       vMensaje = "El valor indicado no es válido > Indique [S] ó [N]...!!!"
     End If
     
  Case "DTS" 'Fecha
    
    If Not IsDate(pValor) Then
       Validacion = False
       vMensaje = "La Fecha indicada no es válida...!!!"
    Else
       pValor = Format(CDate(pValor), "yyyy/mm/dd")
    End If

End Select


If Not Validacion Then
  MsgBox vMensaje, vbExclamation, "Parámetros de Crédito"
  Exit Sub
End If


strSQL = "update crd_parametros set modifica_usuario = '" & glogon.Usuario & "', modifica_Fecha = dbo.MyGetdate()" _
       & ",valor = '" & Trim(pValor) & "' where cod_parametro = '" & pParametro & "'"
Call ConectionExecute(strSQL)

strSQL = "Parámetro de Crédito: " & pParametro & " -> " & pValor

Call Bitacora("Modifica", strSQL)

MsgBox "Parámetro actualizado satisfactoriamente...!", vbInformation

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Function fxGuardar() As Long
Dim vTemp As String

On Error GoTo vError

fxGuardar = 0

vGrid.Row = vGrid.ActiveRow
vGrid.col = 3
vTemp = vGrid.Text


vGrid.col = 1
Call sbGuardaParametro(vGrid.Text, vTemp, vGrid.CellTag)

fxGuardar = 1

Exit Function

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Function

Private Sub vGrid_KeyDown(KeyCode As Integer, Shift As Integer)
Dim i As Integer

If vGrid.ActiveCol = vGrid.MaxCols And (KeyCode = vbKeyReturn Or KeyCode = vbKeyTab) Then
  i = fxGuardar
  If i = 0 Then Exit Sub
  vGrid.Row = vGrid.ActiveRow
End If

If vGrid.ActiveCol = vGrid.MaxCols And KeyCode = vbKeyF4 Then
   vGrid.Row = vGrid.ActiveRow
   vGrid.col = 1
   If vGrid.CellTag = "CTA" Then
      gCuenta = ""
      frmCntX_ConsultaCuentas.Show vbModal
      If gCuenta <> "" Then
        vGrid.col = 3
        vGrid.Text = fxgCntCuentaFormato(True, gCuenta)
        vGrid.TextTip = TextTipFixed
        vGrid.TextTipDelay = 1000
        vGrid.CellNoteIndicatorColor = vbBlue
        vGrid.CellNoteIndicatorShape = CellNoteIndicatorShapeTriangle
        vGrid.CellNote = fxgCntCuentaDesc(gCuenta)
        
        vGrid.col = 1
        Call sbGuardaParametro(vGrid.Text, gCuenta, "CTA")
      End If
      
   
   End If
End If

End Sub
