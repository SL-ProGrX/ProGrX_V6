VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "Codejock.Controls.v22.1.0.ocx"
Begin VB.Form frmAF_Congelar 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Bloquear/Congelar Actividad de Personas"
   ClientHeight    =   7335
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11505
   Icon            =   "frmAF_Congelar.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7335
   ScaleWidth      =   11505
   Begin XtremeSuiteControls.TabControl tcMain 
      Height          =   6252
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   11292
      _Version        =   1441793
      _ExtentX        =   19918
      _ExtentY        =   11028
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
      Item(0).Caption =   "Consultas"
      Item(0).ControlCount=   12
      Item(0).Control(0)=   "lsw"
      Item(0).Control(1)=   "txtConCedula"
      Item(0).Control(2)=   "txtConNombre"
      Item(0).Control(3)=   "Label1(26)"
      Item(0).Control(4)=   "dtpConInicio"
      Item(0).Control(5)=   "Label1(0)"
      Item(0).Control(6)=   "dtpConCorte"
      Item(0).Control(7)=   "chkConFechas"
      Item(0).Control(8)=   "cboConEstado"
      Item(0).Control(9)=   "Label1(4)"
      Item(0).Control(10)=   "btnAccion(0)"
      Item(0).Control(11)=   "btnAccion(1)"
      Item(1).Caption =   "Registro"
      Item(1).ControlCount=   19
      Item(1).Control(0)=   "txtNombre"
      Item(1).Control(1)=   "txtCedula"
      Item(1).Control(2)=   "Label1(1)"
      Item(1).Control(3)=   "cmdGuardar"
      Item(1).Control(4)=   "cmdNuevo"
      Item(1).Control(5)=   "txtCodigo"
      Item(1).Control(6)=   "txtNotas"
      Item(1).Control(7)=   "Label1(19)"
      Item(1).Control(8)=   "cboEstado"
      Item(1).Control(9)=   "Label1(8)"
      Item(1).Control(10)=   "Label1(2)"
      Item(1).Control(11)=   "GroupBox1(0)"
      Item(1).Control(12)=   "GroupBox1(1)"
      Item(1).Control(13)=   "GroupBox1(2)"
      Item(1).Control(14)=   "Label1(5)"
      Item(1).Control(15)=   "cboCausa"
      Item(1).Control(16)=   "dtpInicio"
      Item(1).Control(17)=   "dtpCorte"
      Item(1).Control(18)=   "Label1(9)"
      Item(2).Caption =   "Mantenimiento"
      Item(2).ControlCount=   1
      Item(2).Control(0)=   "vGrid"
      Begin XtremeSuiteControls.ListView lsw 
         Height          =   4695
         Left            =   120
         TabIndex        =   2
         Top             =   1440
         Width           =   11055
         _Version        =   1441793
         _ExtentX        =   19500
         _ExtentY        =   8281
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
         FlatScrollBar   =   -1  'True
         Appearance      =   16
         ShowBorder      =   0   'False
      End
      Begin XtremeSuiteControls.CheckBox chkConFechas 
         Height          =   252
         Left            =   7800
         TabIndex        =   9
         Top             =   960
         Width           =   972
         _Version        =   1441793
         _ExtentX        =   1714
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Todas"
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
         Value           =   1
      End
      Begin XtremeSuiteControls.FlatEdit txtConCedula 
         Height          =   312
         Left            =   1440
         TabIndex        =   3
         Top             =   600
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
         Alignment       =   2
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtConNombre 
         Height          =   312
         Left            =   3120
         TabIndex        =   4
         Top             =   600
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.DateTimePicker dtpConInicio 
         Height          =   312
         Left            =   5040
         TabIndex        =   6
         Top             =   960
         Width           =   1332
         _Version        =   1441793
         _ExtentX        =   2350
         _ExtentY        =   556
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
      Begin XtremeSuiteControls.DateTimePicker dtpConCorte 
         Height          =   312
         Left            =   6360
         TabIndex        =   8
         Top             =   960
         Width           =   1332
         _Version        =   1441793
         _ExtentX        =   2350
         _ExtentY        =   556
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
      Begin FPSpreadADO.fpSpread vGrid 
         Height          =   5532
         Left            =   -69760
         TabIndex        =   10
         Top             =   480
         Visible         =   0   'False
         Width           =   10692
         _Version        =   524288
         _ExtentX        =   18860
         _ExtentY        =   9758
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
         SpreadDesigner  =   "frmAF_Congelar.frx":000C
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin XtremeSuiteControls.FlatEdit txtNombre 
         Height          =   330
         Left            =   -65680
         TabIndex        =   11
         Top             =   1560
         Visible         =   0   'False
         Width           =   4815
         _Version        =   1441793
         _ExtentX        =   8488
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
      Begin XtremeSuiteControls.FlatEdit txtCedula 
         Height          =   330
         Left            =   -67360
         TabIndex        =   12
         Top             =   1560
         Visible         =   0   'False
         Width           =   1695
         _Version        =   1441793
         _ExtentX        =   2984
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
      Begin XtremeSuiteControls.PushButton cmdGuardar 
         Height          =   495
         Left            =   -62200
         TabIndex        =   14
         Top             =   480
         Visible         =   0   'False
         Width           =   1335
         _Version        =   1441793
         _ExtentX        =   2350
         _ExtentY        =   868
         _StockProps     =   79
         Caption         =   "Guardar"
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
         Picture         =   "frmAF_Congelar.frx":0621
      End
      Begin XtremeSuiteControls.PushButton cmdNuevo 
         Height          =   495
         Left            =   -63520
         TabIndex        =   15
         Top             =   480
         Visible         =   0   'False
         Width           =   1335
         _Version        =   1441793
         _ExtentX        =   2350
         _ExtentY        =   868
         _StockProps     =   79
         Caption         =   "Nuevo"
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
         Picture         =   "frmAF_Congelar.frx":0D52
      End
      Begin XtremeSuiteControls.FlatEdit txtCodigo 
         Height          =   570
         Left            =   -67360
         TabIndex        =   16
         Top             =   480
         Visible         =   0   'False
         Width           =   1695
         _Version        =   1441793
         _ExtentX        =   2990
         _ExtentY        =   1005
         _StockProps     =   77
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
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
         Height          =   735
         Left            =   -67360
         TabIndex        =   17
         Top             =   2400
         Visible         =   0   'False
         Width           =   6495
         _Version        =   1441793
         _ExtentX        =   11456
         _ExtentY        =   1296
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
      Begin XtremeSuiteControls.ComboBox cboEstado 
         Height          =   315
         Left            =   -67360
         TabIndex        =   19
         Top             =   2040
         Visible         =   0   'False
         Width           =   1695
         _Version        =   1441793
         _ExtentX        =   2990
         _ExtentY        =   582
         _StockProps     =   77
         ForeColor       =   1973790
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
      Begin XtremeSuiteControls.ComboBox cboConEstado 
         Height          =   312
         Left            =   1440
         TabIndex        =   21
         Top             =   960
         Width           =   1692
         _Version        =   1441793
         _ExtentX        =   2990
         _ExtentY        =   582
         _StockProps     =   77
         ForeColor       =   1973790
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
      Begin XtremeSuiteControls.PushButton btnAccion 
         Height          =   375
         Index           =   0
         Left            =   8760
         TabIndex        =   24
         Top             =   600
         Width           =   495
         _Version        =   1441793
         _ExtentX        =   873
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
         Appearance      =   17
         Picture         =   "frmAF_Congelar.frx":1384
      End
      Begin XtremeSuiteControls.PushButton btnAccion 
         Height          =   375
         Index           =   1
         Left            =   9240
         TabIndex        =   25
         Top             =   600
         Width           =   495
         _Version        =   1441793
         _ExtentX        =   873
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
         Appearance      =   17
         Picture         =   "frmAF_Congelar.frx":1A84
      End
      Begin XtremeSuiteControls.GroupBox GroupBox1 
         Height          =   855
         Index           =   0
         Left            =   -69520
         TabIndex        =   26
         Top             =   3240
         Visible         =   0   'False
         Width           =   4215
         _Version        =   1441793
         _ExtentX        =   7429
         _ExtentY        =   1503
         _StockProps     =   79
         Caption         =   "Acciones para Afiliaciones y Estados de Cuentas"
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
         Appearance      =   16
         BorderStyle     =   1
         Begin XtremeSuiteControls.CheckBox chkPerLiquidacion 
            Height          =   252
            Left            =   360
            TabIndex        =   27
            Top             =   360
            Width           =   3852
            _Version        =   1441793
            _ExtentX        =   6794
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Permite realizar Liquidaciones"
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
         Begin XtremeSuiteControls.CheckBox chkPerMostrarEC 
            Height          =   252
            Left            =   360
            TabIndex        =   28
            Top             =   600
            Width           =   3852
            _Version        =   1441793
            _ExtentX        =   6794
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Permite Mostrar en el Estado de Cuenta"
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
      End
      Begin XtremeSuiteControls.GroupBox GroupBox1 
         Height          =   1935
         Index           =   1
         Left            =   -69520
         TabIndex        =   29
         Top             =   4320
         Visible         =   0   'False
         Width           =   4215
         _Version        =   1441793
         _ExtentX        =   7429
         _ExtentY        =   3408
         _StockProps     =   79
         Caption         =   "Acciones para Crédito y Cobro"
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
         Appearance      =   16
         BorderStyle     =   1
         Begin XtremeSuiteControls.CheckBox chkPerAbonoCajas 
            Height          =   252
            Left            =   360
            TabIndex        =   30
            Top             =   360
            Width           =   3972
            _Version        =   1441793
            _ExtentX        =   7006
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Permite realizar abonos en Cajas"
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
         Begin XtremeSuiteControls.CheckBox chkPerCerrarCreditos 
            Height          =   252
            Left            =   360
            TabIndex        =   31
            Top             =   600
            Width           =   3972
            _Version        =   1441793
            _ExtentX        =   7006
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Permite el acceso a nuevos créditos"
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
         Begin XtremeSuiteControls.CheckBox chkPerCobroJudicial 
            Height          =   252
            Left            =   360
            TabIndex        =   32
            Top             =   840
            Width           =   3972
            _Version        =   1441793
            _ExtentX        =   7006
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Permite realizar Cobros Judiciales"
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
         Begin XtremeSuiteControls.CheckBox chkPerTraspasos 
            Height          =   252
            Left            =   360
            TabIndex        =   33
            Top             =   1080
            Width           =   3972
            _Version        =   1441793
            _ExtentX        =   7006
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Permite realizar traspasos de deudas"
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
         Begin XtremeSuiteControls.CheckBox chkPerReversiones 
            Height          =   252
            Left            =   360
            TabIndex        =   34
            Top             =   1320
            Width           =   3972
            _Version        =   1441793
            _ExtentX        =   7006
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Permite realizar Reversiones de ""Cobros"""
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
         Begin XtremeSuiteControls.CheckBox chkPerReadecuaciones 
            Height          =   252
            Left            =   360
            TabIndex        =   35
            Top             =   1560
            Width           =   3972
            _Version        =   1441793
            _ExtentX        =   7006
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Permite realizar Readecuaciones"
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
      End
      Begin XtremeSuiteControls.GroupBox GroupBox1 
         Height          =   2655
         Index           =   2
         Left            =   -64840
         TabIndex        =   36
         Top             =   3240
         Visible         =   0   'False
         Width           =   5175
         _Version        =   1441793
         _ExtentX        =   9123
         _ExtentY        =   4678
         _StockProps     =   79
         Caption         =   "Acciones para el proceso de Deducciones de Planillas"
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
         Appearance      =   16
         BorderStyle     =   1
         Begin XtremeSuiteControls.CheckBox chkPerCobroFND 
            Height          =   252
            Left            =   360
            TabIndex        =   37
            Top             =   360
            Width           =   4932
            _Version        =   1441793
            _ExtentX        =   8700
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Permite la generación del cobro del Fondo Solidario"
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
         Begin XtremeSuiteControls.CheckBox chkPerCobroCR 
            Height          =   252
            Left            =   360
            TabIndex        =   38
            Top             =   600
            Width           =   3972
            _Version        =   1441793
            _ExtentX        =   7006
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Permite la generación del cobro cuota C.R."
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
         Begin XtremeSuiteControls.CheckBox chkPerDeducionAportes 
            Height          =   252
            Left            =   360
            TabIndex        =   39
            Top             =   1200
            Width           =   4932
            _Version        =   1441793
            _ExtentX        =   8700
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Permite la generación de deducciones de Aportes"
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
         Begin XtremeSuiteControls.CheckBox chkPerDeduccionesCreditos 
            Height          =   252
            Left            =   360
            TabIndex        =   40
            Top             =   1440
            Width           =   4932
            _Version        =   1441793
            _ExtentX        =   8700
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Permite la generación de deducciones de crédito"
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
         Begin XtremeSuiteControls.CheckBox chkPerGeneracionMora 
            Height          =   252
            Left            =   360
            TabIndex        =   41
            Top             =   1680
            Width           =   3972
            _Version        =   1441793
            _ExtentX        =   7006
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Permite la generación de cuotas morosas"
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
      End
      Begin XtremeSuiteControls.ComboBox cboCausa 
         Height          =   330
         Left            =   -67360
         TabIndex        =   43
         Top             =   1080
         Visible         =   0   'False
         Width           =   6495
         _Version        =   1441793
         _ExtentX        =   11456
         _ExtentY        =   582
         _StockProps     =   77
         ForeColor       =   1973790
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
      Begin XtremeSuiteControls.DateTimePicker dtpInicio 
         Height          =   315
         Left            =   -63520
         TabIndex        =   44
         Top             =   2040
         Visible         =   0   'False
         Width           =   1335
         _Version        =   1441793
         _ExtentX        =   2350
         _ExtentY        =   556
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
         Height          =   315
         Left            =   -62200
         TabIndex        =   45
         Top             =   2040
         Visible         =   0   'False
         Width           =   1335
         _Version        =   1441793
         _ExtentX        =   2350
         _ExtentY        =   556
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
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Rango de Fechas"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   9
         Left            =   -65560
         TabIndex        =   46
         Top             =   2040
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Causa"
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
         Left            =   -68440
         TabIndex        =   42
         Top             =   1080
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Id Gestión"
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
         Left            =   -68560
         TabIndex        =   23
         Top             =   600
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Estado"
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
         Index           =   4
         Left            =   480
         TabIndex        =   22
         Top             =   960
         Width           =   852
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Estado"
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
         Left            =   -68440
         TabIndex        =   20
         Top             =   2040
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
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
         Index           =   19
         Left            =   -68440
         TabIndex        =   18
         Top             =   2400
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Identificación"
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
         Index           =   1
         Left            =   -69040
         TabIndex        =   13
         Top             =   1560
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Fechas"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   252
         Index           =   0
         Left            =   3480
         TabIndex        =   7
         Top             =   960
         Width           =   1212
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Identificación"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   252
         Index           =   26
         Left            =   120
         TabIndex        =   5
         Top             =   600
         Width           =   1212
      End
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Bloqueo de Actividad de Personas"
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
      Height          =   495
      Index           =   3
      Left            =   1880
      TabIndex        =   0
      Top             =   240
      Width           =   6255
   End
   Begin VB.Image imgBanner 
      Height          =   852
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   11892
   End
End
Attribute VB_Name = "frmAF_Congelar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vCodigo As String



Private Sub btnAccion_Click(Index As Integer)

Select Case Index
    Case 0 'Buscar
        Call sbBuscar
    Case 1 'Exportar
        Call Excel_Exportar_Lsw(lsw)
End Select

End Sub

Private Sub cboCausa_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then cboEstado.SetFocus
End Sub


Private Sub cboEstado_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then dtpInicio.SetFocus
End Sub


Private Sub chkConFechas_Click()
If chkConFechas.Value = vbChecked Then
  dtpConInicio.Enabled = False
Else
  dtpConInicio.Enabled = True
End If
dtpConCorte.Enabled = dtpConInicio.Enabled
End Sub

Private Sub sbBuscar()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem

On Error GoTo vError

Me.MousePointer = vbHourglass

lsw.ListItems.Clear

strSQL = "select C.cod_congelar,C.cedula,S.nombre,C.estado,C.notas,X.descripcion as Causa,C.fecha_Inicia" _
       & " from afi_congelar C inner join afi_congelar_causas X on C.cod_causa = X.cod_causa" _
       & " inner join socios S on S.cedula = C.cedula" _
       & " where C.cedula like '%" & txtConCedula & "%' and S.nombre like '%" & txtConNombre.Text & "%'" _
       & " and C.Estado = '" & Mid(cboConEstado.Text, 1, 1) & "'"
       
If chkConFechas.Value = vbUnchecked Then
   strSQL = strSQL & " and C.fecha_inicia between '" & Format(dtpConInicio.Value, "yyyy/mm/dd") _
          & " 00:00:00' and '" & Format(dtpConCorte.Value, "yyyy/mm/dd") & " 23:59:59'"
End If

Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
 Set itmX = lsw.ListItems.Add(, , rs!Cedula)
     itmX.SubItems(1) = IIf((rs!Estado = "A"), "Activo", "Inactivo")
     itmX.SubItems(2) = rs!Nombre
     itmX.SubItems(3) = Format(rs!fecha_inicia, "yyyy/mm/dd")
     itmX.SubItems(4) = rs!Causa
     itmX.SubItems(5) = rs!Notas
     itmX.Tag = rs!cod_congelar
 rs.MoveNext
Loop
rs.Close

Me.MousePointer = vbDefault
Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub cmdGuardar_Click()
Dim strSQL As String, rs As New ADODB.Recordset

If Not fxVerifica Then
  Exit Sub
End If

On Error GoTo vError


If txtCodigo.Text = "" Then
   strSQL = "insert afi_congelar(cedula,cod_causa,notas,fecha_crea,usuario_crea" _
          & ",estado,fecha_Inicia,fecha_Finaliza,per_liquidacion,per_mostrar_ec,per_abono_cajas,per_cierra_AcCreditos" _
          & ",per_cobro_judicial,per_traspaso_deudas,per_reversiones,per_readecuaciones,per_deducciones_creditos" _
          & ",per_deducciones_aportes,per_generacion_mora,per_cobro_FndSol,per_cobro_cuotaCr)" _
          & " values('" & txtCedula.Text & "','" & cboCausa.ItemData(cboCausa.ListIndex) & "','" & txtNotas.Text & "',dbo.MyGetdate(),'" _
          & glogon.Usuario & "','" & Mid(cboEstado.Text, 1, 1) & "','" & Format(dtpInicio.Value, "yyyy/mm/dd") _
          & "','" & Format(dtpCorte.Value, "yyyy/mm/dd") & "'," & chkPerLiquidacion.Value _
          & "," & chkPerMostrarEC.Value & "," & chkPerAbonoCajas.Value & "," & chkPerCerrarCreditos.Value _
          & "," & chkPerCobroJudicial.Value & "," & chkPerTraspasos.Value & "," & chkPerReversiones.Value _
          & "," & chkPerReadecuaciones.Value & "," & chkPerDeduccionesCreditos.Value & "," & chkPerDeducionAportes.Value _
          & "," & chkPerGeneracionMora.Value & "," & chkPerCobroFND.Value & "," & chkPerCobroCR.Value & ")"
   Call ConectionExecute(strSQL)
   
   strSQL = "select isnull(max(cod_congelar),0) as Consec from afi_congelar" _
          & " where cedula = '" & txtCedula.Text & "'"
   Call OpenRecordSet(rs, strSQL)
    txtCodigo.Text = CStr(rs!consec)
   rs.Close
   
   Call Bitacora("Registra", "Congelamiento " & txtCodigo.Text & " Cedula = " & txtCedula)

Else
   strSQL = "update afi_congelar set notas = '" & txtNotas.Text & "', cod_causa = '" & cboCausa.ItemData(cboCausa.ListIndex) _
          & "',estado = '" & Mid(cboEstado.Text, 1, 1) & "',per_liquidacion = " & chkPerLiquidacion.Value _
          & ",per_mostrar_Ec = " & chkPerMostrarEC.Value & ",per_abono_cajas = " & chkPerAbonoCajas.Value _
          & ",per_cobro_judicial = " & chkPerCobroJudicial.Value & ",per_traspaso_deudas = " & chkPerTraspasos.Value _
          & ",per_reversiones = " & chkPerReversiones.Value & ",per_readecuaciones = " & chkPerReadecuaciones.Value _
          & ",per_deducciones_creditos = " & chkPerDeduccionesCreditos.Value & ",per_deducciones_aportes = " & chkPerDeducionAportes.Value _
          & ",per_generacion_mora = " & chkPerGeneracionMora.Value & ",per_cobro_FndSol = " & chkPerCobroFND.Value _
          & ",per_cobro_cuotaCr = " & chkPerCobroCR.Value & ",fecha_inicia = '" & Format(dtpInicio.Value, "yyyy/mm/dd") _
          & "',fecha_finaliza = '" & Format(dtpCorte.Value, "yyyy/mm/dd") _
          & "' where cod_congelar = " & txtCodigo.Text
   Call ConectionExecute(strSQL)
   
   Call Bitacora("Modifica", "Congelamiento " & txtCodigo.Text & " Cedula = " & txtCedula.Text)
End If

MsgBox "Datos Actualizados Satisfactoriamente...", vbInformation

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Function fxVerifica() As Boolean

fxVerifica = True

End Function

Private Sub cmdNuevo_Click()
Dim strSQL As String, rs As New ADODB.Recordset

txtCedula = ""
txtNombre = ""
txtCodigo = ""
txtNotas = ""

tcMain.Item(1).Selected = True

cboEstado.Clear
cboEstado.AddItem "Activo"
cboEstado.AddItem "Inactivo"
cboEstado.Text = "Activo"

dtpInicio.MinDate = fxFechaServidor
dtpInicio.Value = dtpInicio.MinDate

dtpCorte.Value = dtpInicio.Value

chkPerAbonoCajas.Value = vbUnchecked
chkPerCerrarCreditos.Value = vbUnchecked
chkPerCobroCR.Value = vbUnchecked
chkPerCobroFND.Value = vbUnchecked
chkPerCobroJudicial.Value = vbUnchecked
chkPerDeduccionesCreditos.Value = vbUnchecked
chkPerDeducionAportes.Value = vbUnchecked
chkPerGeneracionMora.Value = vbUnchecked
chkPerLiquidacion.Value = vbUnchecked
chkPerMostrarEC.Value = vbUnchecked
chkPerReadecuaciones.Value = vbUnchecked
chkPerReversiones.Value = vbUnchecked
chkPerTraspasos.Value = vbUnchecked

'Cargar Aqui las causas
strSQL = "select rtrim(COD_CAUSA) as 'IdX',  rtrim(descripcion) as  'ItmX' from AFI_CONGELAR_CAUSAS" _
       & " where Activa = 1 order by COD_CAUSA"
Call sbCbo_Llena_New(cboCausa, strSQL, False, True)

txtCedula.SetFocus

End Sub

Private Sub dtpCorte_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtNotas.SetFocus
End Sub

Private Sub dtpInicio_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then dtpCorte.SetFocus
End Sub

Private Sub Form_Activate()
vModulo = 1
End Sub

Private Sub Form_Load()
Dim strSQL As String

vModulo = 1
vGrid.AppearanceStyle = fxGridStyle

Set imgBanner.Picture = frmContenedor.imgBanner_01.Picture


With lsw.ColumnHeaders
    .Clear
    .Add , , "Identificación", 1800
    .Add , , "Estado", 1200, vbCenter
    .Add , , "Nombre", 3800
    .Add , , "Inicio", 1200, vbCenter
    .Add , , "Causa", 2800
    .Add , , "Notas", 3800
    

End With

tcMain.Item(0).Selected = True

dtpConInicio.Value = fxFechaServidor
dtpConCorte.Value = dtpConInicio.Value

cboConEstado.Clear
cboConEstado.AddItem "Activo"
cboConEstado.AddItem "Inactivo"
cboConEstado.Text = "Activo"

Call chkConFechas_Click

Call Formularios(Me)
Call RefrescaTags(Me)

End Sub

Private Function fxExiste(vCod As String) As Boolean
Dim strSQL As String, rs As New ADODB.Recordset

strSQL = "select isnull(count(*),0) as 'Existe' from AFI_CONGELAR_CAUSAS" _
       & " where COD_CAUSA = '" & vCod & "'"
Call OpenRecordSet(rs, strSQL)
    fxExiste = IIf((rs!Existe = 1), True, False)
rs.Close
End Function


Private Function fxGuardar() As Long
Dim strSQL As String, rs As New ADODB.Recordset
'Guarda la información de la linea
'si es Insert devuelve el codigo, sino devuelve 0

On Error GoTo vError

fxGuardar = 0
vGrid.Row = vGrid.ActiveRow
vGrid.col = 1

If vGrid.Text = "" Then Exit Function


If Not fxExiste(vGrid.Text) Then
   vGrid.col = 1
   strSQL = "insert AFI_CONGELAR_CAUSAS(cod_causa,descripcion, activa, registro_fecha, registro_usuario)" _
          & " values('" & vGrid.Text & "','"
   vGrid.col = 2
   strSQL = strSQL & vGrid.Text & "',"
   vGrid.col = 3
   strSQL = strSQL & vGrid.Value & ", dbo.Mygetdate(), '" & glogon.Usuario & "')"
   
   
   Call ConectionExecute(strSQL)
   vGrid.col = 1
   Call Bitacora("Registra", "Causa de Congelamiento Cod: " & vGrid.Text)
   
   vGrid.col = 4
   vGrid.Text = Date
   vGrid.col = 5
   vGrid.Text = glogon.Usuario
   
   
 Else 'Actualizar
    vGrid.col = 2
    strSQL = "update AFI_CONGELAR_CAUSAS set descripcion = '" & vGrid.Text & "', Activa = "
    vGrid.col = 3
    strSQL = strSQL & vGrid.Value
    vGrid.col = 1
    strSQL = strSQL & " where cod_causa = '" & vGrid.Text & "'"
   
    Call ConectionExecute(strSQL)
    
    Call Bitacora("Modifica", "Causa de Congelamiento Cod: " & vGrid.Text)
    
End If

Exit Function
   
vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

   
End Function


Private Sub lsw_ColumnClick(ByVal ColumnHeader As XtremeSuiteControls.ListViewColumnHeader)
 lsw.SortKey = ColumnHeader.Index - 1
  If lsw.SortOrder = 0 Then lsw.SortOrder = 1 Else lsw.SortOrder = 0
  lsw.Sorted = True
End Sub

Private Sub lsw_ItemClick(ByVal Item As XtremeSuiteControls.ListViewItem)
On Error GoTo vError

tcMain.Item(1).Selected = True

Call sbCargaCaso(Item.Tag)

vError:

End Sub

Private Sub tcMain_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
Dim strSQL As String


Select Case Item.Index
    Case 1
        Call cmdNuevo_Click
    Case 2
            strSQL = "select COD_CAUSA,descripcion,Activa,registro_fecha,registro_usuario from AFI_CONGELAR_CAUSAS" _
                   & " order by COD_CAUSA"
            Call sbCargaGrid(vGrid, 5, strSQL)

End Select

End Sub

Private Sub txtCedula_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtNombre.SetFocus
If KeyCode = vbKeyF4 Then
   gBusquedas.Resultado = ""
   gBusquedas.Resultado2 = ""
   gBusquedas.Columna = "cedula"
   gBusquedas.Orden = "cedula"
   gBusquedas.Convertir = "N"
   gBusquedas.Consulta = "SELECT CEDULA,NOMBRE FROM SOCIOS"
   gBusquedas.Filtro = ""
   frmBusquedas.Show vbModal
   txtCedula.Text = gBusquedas.Resultado
   txtNombre.Text = gBusquedas.Resultado2
End If

End Sub

Private Sub txtCedula_LostFocus()
txtNombre.Text = fxNombre(txtCedula)
End Sub


Private Sub txtNombre_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtCodigo.SetFocus
If KeyCode = vbKeyF4 Then
   gBusquedas.Resultado = ""
   gBusquedas.Resultado2 = ""
   gBusquedas.Columna = "nombre"
   gBusquedas.Orden = "nombre"
   gBusquedas.Convertir = "N"
   gBusquedas.Consulta = "SELECT CEDULA,NOMBRE FROM SOCIOS"
   gBusquedas.Filtro = ""
   frmBusquedas.Show vbModal
   txtCedula = gBusquedas.Resultado
   txtNombre = gBusquedas.Resultado2
End If

End Sub


Private Sub sbCargaCaso(vCongela As String)
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

strSQL = "select C.*,S.nombre,rtrim(C.cod_causa) as 'CausaId', rtrim(X.descripcion) as 'CausaDesc'" _
       & " from afi_congelar C inner join Socios S on C.cedula = S.cedula" _
       & " inner join afi_congelar_causas X on C.cod_causa = X.cod_causa" _
       & " where C.cod_congelar = " & vCongela
Call OpenRecordSet(rs, strSQL)
If Not rs.EOF And Not rs.BOF Then
    txtCodigo.Text = CStr(rs!cod_congelar)
    
    txtCedula.Text = rs!Cedula
    txtNombre.Text = rs!Nombre
    
    Call sbCboAsignaDato(cboCausa, rs!CausaDesc, True, rs!CausaId)
 
    If rs!Estado = "A" Then
        cboEstado.Text = "Activo"
    Else
        cboEstado.Text = "Inactivo"
    End If
    
    dtpInicio.MinDate = rs!fecha_inicia
    dtpInicio.Value = rs!fecha_inicia
    
    dtpCorte.Value = rs!Fecha_Finaliza
    txtNotas.Text = rs!Notas
    
    chkPerLiquidacion.Value = rs!per_liquidacion
    chkPerMostrarEC.Value = rs!per_mostrar_ec
    
    chkPerAbonoCajas.Value = rs!per_abono_cajas
    chkPerCerrarCreditos.Value = rs!per_cierra_acCreditos
    chkPerCobroJudicial.Value = rs!per_cobro_judicial
    chkPerTraspasos.Value = rs!per_traspaso_deudas
    chkPerReversiones.Value = rs!per_reversiones
    chkPerReadecuaciones.Value = rs!per_reaDecuaciones
    
    chkPerDeduccionesCreditos.Value = rs!per_deducciones_creditos
    chkPerDeducionAportes.Value = rs!per_deducciones_aportes
    chkPerGeneracionMora.Value = rs!per_generacion_mora
    chkPerCobroFND.Value = rs!per_cobro_FndSol
    chkPerCobroCR.Value = rs!per_cobro_cuotaCr

End If
rs.Close

Exit Sub
vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub txtCodigo_KeyDown(KeyCode As Integer, Shift As Integer)

On Error GoTo vError

If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then cboCausa.SetFocus

If KeyCode = vbKeyF4 Then
   gBusquedas.Resultado = ""
   gBusquedas.Resultado2 = ""
   gBusquedas.Columna = "C.cod_congelar"
   gBusquedas.Orden = "C.cod_congelar"
   gBusquedas.Convertir = "N"
   gBusquedas.Consulta = "SELECT C.Cod_congelar,C.estado,C.fecha_finaliza,S.NOMBRE" _
             & " FROM afi_congelar C inner join SOCIOS S on C.cedula = S.cedula"
   gBusquedas.Filtro = " and C.cedula like '" & txtCedula & "%'"
   frmBusquedas.Show vbModal
   If IsNumeric(gBusquedas.Resultado) Then
      Call sbCargaCaso(gBusquedas.Resultado)
   End If
   
End If
vError:

End Sub

Private Sub txtConCedula_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtConNombre.SetFocus
If KeyCode = vbKeyF4 Then
   gBusquedas.Resultado = ""
   gBusquedas.Resultado2 = ""
   gBusquedas.Columna = "cedula"
   gBusquedas.Orden = "cedula"
   gBusquedas.Convertir = "N"
   gBusquedas.Consulta = "SELECT CEDULA,NOMBRE FROM SOCIOS"
   gBusquedas.Filtro = ""
   frmBusquedas.Show vbModal
   txtConCedula = gBusquedas.Resultado
   txtConNombre = gBusquedas.Resultado2
End If
End Sub


Private Sub txtConNombre_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then
   gBusquedas.Resultado = ""
   gBusquedas.Resultado2 = ""
   gBusquedas.Columna = "nombre"
   gBusquedas.Orden = "nombre"
   gBusquedas.Convertir = "N"
   gBusquedas.Consulta = "SELECT CEDULA,NOMBRE FROM SOCIOS"
   gBusquedas.Filtro = ""
   frmBusquedas.Show vbModal
   txtConCedula = gBusquedas.Resultado
   txtConNombre = gBusquedas.Resultado2
End If

End Sub

Private Sub txtNotas_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo vError
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then cmdGuardar.SetFocus
vError:
End Sub

Private Sub vGrid_KeyDown(KeyCode As Integer, Shift As Integer)
Dim i As Long

If vGrid.ActiveCol = 3 And (KeyCode = vbKeyReturn Or KeyCode = vbKeyTab) Then
  i = fxGuardar
  vGrid.Row = vGrid.ActiveRow
  vGrid.col = 1
  If vGrid.MaxRows <= vGrid.ActiveRow Then
    vGrid.MaxRows = vGrid.MaxRows + 1
    vGrid.Row = vGrid.MaxRows
  End If
End If

End Sub



