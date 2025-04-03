VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.5#0"; "comctl32.Ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.controls.v22.1.0.ocx"
Begin VB.Form frmTES_Transacciones 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Transacciones"
   ClientHeight    =   8805
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13155
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8805
   ScaleWidth      =   13155
   Begin XtremeSuiteControls.TabControl tcOpciones 
      Height          =   6375
      Left            =   11040
      TabIndex        =   37
      Top             =   1920
      Width           =   2055
      _Version        =   1441793
      _ExtentX        =   3625
      _ExtentY        =   11245
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
      Color           =   128
      PaintManager.BoldSelected=   -1  'True
      ItemCount       =   1
      Item(0).Caption =   "Opciones"
      Item(0).ControlCount=   10
      Item(0).Control(0)=   "btnOpciones(0)"
      Item(0).Control(1)=   "btnOpciones(1)"
      Item(0).Control(2)=   "btnOpciones(2)"
      Item(0).Control(3)=   "btnOpciones(3)"
      Item(0).Control(4)=   "btnOpciones(4)"
      Item(0).Control(5)=   "btnOpciones(5)"
      Item(0).Control(6)=   "txtRef_01"
      Item(0).Control(7)=   "Label6(3)"
      Item(0).Control(8)=   "txtRef_02"
      Item(0).Control(9)=   "txtRef_03"
      Begin XtremeSuiteControls.PushButton btnOpciones 
         Height          =   420
         Index           =   0
         Left            =   240
         TabIndex        =   38
         Top             =   480
         Width           =   1575
         _Version        =   1441793
         _ExtentX        =   2778
         _ExtentY        =   741
         _StockProps     =   79
         Caption         =   "Anular"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseVisualStyle  =   -1  'True
         TextAlignment   =   1
         Appearance      =   17
         Picture         =   "frmTES_Transacciones.frx":0000
         ImageAlignment  =   4
      End
      Begin XtremeSuiteControls.PushButton btnOpciones 
         Height          =   420
         Index           =   1
         Left            =   240
         TabIndex        =   39
         Top             =   960
         Width           =   1572
         _Version        =   1441793
         _ExtentX        =   2773
         _ExtentY        =   741
         _StockProps     =   79
         Caption         =   "Reclasificar"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseVisualStyle  =   -1  'True
         TextAlignment   =   1
         Appearance      =   17
         Picture         =   "frmTES_Transacciones.frx":0716
         ImageAlignment  =   4
      End
      Begin XtremeSuiteControls.PushButton btnOpciones 
         Height          =   420
         Index           =   2
         Left            =   240
         TabIndex        =   40
         Top             =   1440
         Width           =   1572
         _Version        =   1441793
         _ExtentX        =   2773
         _ExtentY        =   741
         _StockProps     =   79
         Caption         =   "Reimprimir"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseVisualStyle  =   -1  'True
         TextAlignment   =   1
         Appearance      =   17
         Picture         =   "frmTES_Transacciones.frx":0E09
         ImageAlignment  =   4
      End
      Begin XtremeSuiteControls.PushButton btnOpciones 
         Height          =   420
         Index           =   3
         Left            =   240
         TabIndex        =   41
         Top             =   1920
         Width           =   1572
         _Version        =   1441793
         _ExtentX        =   2773
         _ExtentY        =   741
         _StockProps     =   79
         Caption         =   "Fechas"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseVisualStyle  =   -1  'True
         TextAlignment   =   1
         Appearance      =   17
         Picture         =   "frmTES_Transacciones.frx":1510
         ImageAlignment  =   4
      End
      Begin XtremeSuiteControls.PushButton btnOpciones 
         Height          =   420
         Index           =   4
         Left            =   240
         TabIndex        =   42
         Top             =   2400
         Width           =   1572
         _Version        =   1441793
         _ExtentX        =   2773
         _ExtentY        =   741
         _StockProps     =   79
         Caption         =   "Reponer"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseVisualStyle  =   -1  'True
         TextAlignment   =   1
         Appearance      =   17
         Picture         =   "frmTES_Transacciones.frx":1DBC
         ImageAlignment  =   4
      End
      Begin XtremeSuiteControls.PushButton btnOpciones 
         Height          =   420
         Index           =   5
         Left            =   240
         TabIndex        =   43
         Top             =   2880
         Width           =   1572
         _Version        =   1441793
         _ExtentX        =   2773
         _ExtentY        =   741
         _StockProps     =   79
         Caption         =   "Detalle"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseVisualStyle  =   -1  'True
         TextAlignment   =   1
         Appearance      =   17
         Picture         =   "frmTES_Transacciones.frx":24BC
         ImageAlignment  =   4
      End
      Begin XtremeSuiteControls.FlatEdit txtRef_01 
         Height          =   330
         Left            =   240
         TabIndex        =   45
         Top             =   5160
         Width           =   1575
         _Version        =   1441793
         _ExtentX        =   2773
         _ExtentY        =   582
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   7.5
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
      Begin XtremeSuiteControls.FlatEdit txtRef_02 
         Height          =   330
         Left            =   240
         TabIndex        =   47
         Top             =   5520
         Width           =   1575
         _Version        =   1441793
         _ExtentX        =   2773
         _ExtentY        =   582
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   7.5
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
      Begin XtremeSuiteControls.FlatEdit txtRef_03 
         Height          =   330
         Left            =   240
         TabIndex        =   48
         Top             =   5880
         Width           =   1575
         _Version        =   1441793
         _ExtentX        =   2773
         _ExtentY        =   582
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   7.5
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
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         Caption         =   "Referencias"
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
         TabIndex        =   46
         Top             =   4800
         Width           =   1575
      End
   End
   Begin XtremeSuiteControls.TabControl ssTab 
      Height          =   6375
      Left            =   0
      TabIndex        =   7
      Top             =   1920
      Width           =   11055
      _Version        =   1441793
      _ExtentX        =   19500
      _ExtentY        =   11245
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
      PaintManager.BoldSelected=   -1  'True
      ItemCount       =   4
      Item(0).Caption =   "Documento"
      Item(0).Tooltip =   "Información General del Documento"
      Item(0).ControlCount=   40
      Item(0).Control(0)=   "txtBeneficiario"
      Item(0).Control(1)=   "cboTipos"
      Item(0).Control(2)=   "txtTipoCambio"
      Item(0).Control(3)=   "txtRef"
      Item(0).Control(4)=   "cboUnidad"
      Item(0).Control(5)=   "cboConcepto"
      Item(0).Control(6)=   "txtCodBene"
      Item(0).Control(7)=   "txtMonto"
      Item(0).Control(8)=   "Label1(7)"
      Item(0).Control(9)=   "Label6(2)"
      Item(0).Control(10)=   "imgBloqueo"
      Item(0).Control(11)=   "lblRef"
      Item(0).Control(12)=   "Label1(6)"
      Item(0).Control(13)=   "Label1(5)"
      Item(0).Control(14)=   "lblMontoLetra"
      Item(0).Control(15)=   "btnCuenta(0)"
      Item(0).Control(16)=   "Group_Detalle"
      Item(0).Control(17)=   "txtDetalle"
      Item(0).Control(18)=   "txtDivisa"
      Item(0).Control(19)=   "cboCuenta"
      Item(0).Control(20)=   "btnCuenta(1)"
      Item(0).Control(21)=   "cboTipoId"
      Item(0).Control(22)=   "Label7(0)"
      Item(0).Control(23)=   "Label7(1)"
      Item(0).Control(24)=   "Label7(2)"
      Item(0).Control(25)=   "Label7(3)"
      Item(0).Control(26)=   "btnCuenta(2)"
      Item(0).Control(27)=   "txtOrigenNombre"
      Item(0).Control(28)=   "txtOrigenId"
      Item(0).Control(29)=   "Label7(4)"
      Item(0).Control(30)=   "Label7(5)"
      Item(0).Control(31)=   "txtOrigenCta"
      Item(0).Control(32)=   "Label7(6)"
      Item(0).Control(33)=   "txtCorreo"
      Item(0).Control(34)=   "Label7(7)"
      Item(0).Control(35)=   "Label7(8)"
      Item(0).Control(36)=   "Label7(9)"
      Item(0).Control(37)=   "Label7(10)"
      Item(0).Control(38)=   "chkCuentaInterna"
      Item(0).Control(39)=   "lblEstadoSinpe"
      Item(1).Caption =   "Asiento"
      Item(1).ControlCount=   6
      Item(1).Control(0)=   "txtDiferencia"
      Item(1).Control(1)=   "vGrid"
      Item(1).Control(2)=   "Label3(1)"
      Item(1).Control(3)=   "Label3(0)"
      Item(1).Control(4)=   "txtDebito"
      Item(1).Control(5)=   "txtCredito"
      Item(2).Caption =   "Bitácora"
      Item(2).ControlCount=   1
      Item(2).Control(0)=   "vGridDetalle"
      Item(3).Caption =   "SINPE"
      Item(3).ControlCount=   12
      Item(3).Control(0)=   "txtSinpe_Motivo"
      Item(3).Control(1)=   "txtSinpe_Estado"
      Item(3).Control(2)=   "Label7(11)"
      Item(3).Control(3)=   "Label7(12)"
      Item(3).Control(4)=   "Label7(13)"
      Item(3).Control(5)=   "txtSinpe_Referencia"
      Item(3).Control(6)=   "txtSinpe_Fondo"
      Item(3).Control(7)=   "Label7(14)"
      Item(3).Control(8)=   "txtSinpe_DocApl"
      Item(3).Control(9)=   "Label7(15)"
      Item(3).Control(10)=   "txtSinpe_Email"
      Item(3).Control(11)=   "Label7(16)"
      Begin XtremeSuiteControls.CheckBox chkCuentaInterna 
         Height          =   495
         Left            =   9120
         TabIndex        =   71
         Top             =   2760
         Width           =   1335
         _Version        =   1441793
         _ExtentX        =   2355
         _ExtentY        =   873
         _StockProps     =   79
         Caption         =   "Cuenta Interna"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Enabled         =   0   'False
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.GroupBox Group_Detalle 
         Height          =   2415
         Left            =   120
         TabIndex        =   18
         Top             =   4080
         Width           =   10335
         _Version        =   1441793
         _ExtentX        =   18230
         _ExtentY        =   4260
         _StockProps     =   79
         Caption         =   "Afectaciones: "
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
         Begin FPSpreadADO.fpSpread vGrid_Detalle 
            Height          =   1815
            Left            =   0
            TabIndex        =   19
            Top             =   480
            Width           =   10335
            _Version        =   524288
            _ExtentX        =   18230
            _ExtentY        =   3201
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
            MaxCols         =   7
            ScrollBars      =   2
            SpreadDesigner  =   "frmTES_Transacciones.frx":2BD5
            VScrollSpecial  =   -1  'True
            VScrollSpecialType=   2
            AppearanceStyle =   1
         End
         Begin XtremeSuiteControls.PushButton btnDetalle_Cierra 
            Height          =   285
            Left            =   9960
            TabIndex        =   20
            Top             =   0
            Width           =   375
            _Version        =   1441793
            _ExtentX        =   661
            _ExtentY        =   503
            _StockProps     =   79
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial Narrow"
               Size            =   7.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            UseVisualStyle  =   -1  'True
            Appearance      =   17
            Picture         =   "frmTES_Transacciones.frx":3304
         End
      End
      Begin XtremeSuiteControls.PushButton btnCuenta 
         Height          =   315
         Index           =   0
         Left            =   4920
         TabIndex        =   17
         Top             =   2400
         Width           =   975
         _Version        =   1441793
         _ExtentX        =   1714
         _ExtentY        =   550
         _StockProps     =   79
         Caption         =   "Ajuste"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseVisualStyle  =   -1  'True
         Appearance      =   17
      End
      Begin FPSpreadADO.fpSpread vGrid 
         Height          =   5295
         Left            =   -70000
         TabIndex        =   13
         Top             =   480
         Visible         =   0   'False
         Width           =   10695
         _Version        =   524288
         _ExtentX        =   18865
         _ExtentY        =   9340
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
         MaxCols         =   491
         ScrollBars      =   2
         SpreadDesigner  =   "frmTES_Transacciones.frx":38A8
         VScrollSpecial  =   -1  'True
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin FPSpreadADO.fpSpread vGridDetalle 
         Height          =   5895
         Left            =   -69760
         TabIndex        =   16
         Top             =   480
         Visible         =   0   'False
         Width           =   10335
         _Version        =   524288
         _ExtentX        =   18230
         _ExtentY        =   10398
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
         MaxCols         =   5
         MaxRows         =   7
         SpreadDesigner  =   "frmTES_Transacciones.frx":400B
         VScrollSpecial  =   -1  'True
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin XtremeSuiteControls.ComboBox cboTipos 
         Height          =   330
         Left            =   6600
         TabIndex        =   21
         Top             =   840
         Width           =   2415
         _Version        =   1441793
         _ExtentX        =   4260
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
      Begin XtremeSuiteControls.ComboBox cboConcepto 
         Height          =   330
         Left            =   2160
         TabIndex        =   22
         Top             =   840
         Width           =   4215
         _Version        =   1441793
         _ExtentX        =   7435
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
      Begin XtremeSuiteControls.ComboBox cboUnidad 
         Height          =   330
         Left            =   2160
         TabIndex        =   23
         Top             =   480
         Width           =   4215
         _Version        =   1441793
         _ExtentX        =   7435
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
      Begin XtremeSuiteControls.FlatEdit txtBeneficiario 
         Height          =   330
         Left            =   2160
         TabIndex        =   29
         Top             =   1920
         Width           =   6855
         _Version        =   1441793
         _ExtentX        =   12091
         _ExtentY        =   582
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
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
      Begin XtremeSuiteControls.FlatEdit txtDivisa 
         Height          =   330
         Left            =   6480
         TabIndex        =   31
         Top             =   3720
         Width           =   975
         _Version        =   1441793
         _ExtentX        =   1720
         _ExtentY        =   582
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   7.5
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
      Begin XtremeSuiteControls.FlatEdit txtTipoCambio 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0.0000"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   5130
            SubFormatType   =   1
         EndProperty
         Height          =   330
         Left            =   7440
         TabIndex        =   32
         Top             =   3720
         Width           =   1575
         _Version        =   1441793
         _ExtentX        =   2778
         _ExtentY        =   582
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
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
      Begin XtremeSuiteControls.FlatEdit txtMonto 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   5130
            SubFormatType   =   1
         EndProperty
         Height          =   330
         Left            =   2160
         TabIndex        =   30
         Top             =   3720
         Width           =   2415
         _Version        =   1441793
         _ExtentX        =   4260
         _ExtentY        =   582
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
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
      Begin XtremeSuiteControls.FlatEdit txtRef 
         Height          =   330
         Left            =   6480
         TabIndex        =   33
         Top             =   4080
         Width           =   2535
         _Version        =   1441793
         _ExtentX        =   4471
         _ExtentY        =   582
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
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
      Begin XtremeSuiteControls.FlatEdit txtDetalle 
         Height          =   1155
         Left            =   2160
         TabIndex        =   34
         Top             =   5040
         Width           =   8175
         _Version        =   1441793
         _ExtentX        =   14420
         _ExtentY        =   2037
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.ComboBox cboCuenta 
         Height          =   330
         Left            =   2160
         TabIndex        =   35
         Top             =   2400
         Width           =   2415
         _Version        =   1441793
         _ExtentX        =   4260
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.PushButton btnCuenta 
         Height          =   315
         Index           =   1
         Left            =   5880
         TabIndex        =   36
         Top             =   2400
         Width           =   375
         _Version        =   1441793
         _ExtentX        =   656
         _ExtentY        =   550
         _StockProps     =   79
         Caption         =   "..."
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseVisualStyle  =   -1  'True
         Appearance      =   17
      End
      Begin XtremeSuiteControls.FlatEdit txtDiferencia 
         Height          =   315
         Left            =   -68320
         TabIndex        =   49
         Top             =   5940
         Visible         =   0   'False
         Width           =   1815
         _Version        =   1441793
         _ExtentX        =   3196
         _ExtentY        =   550
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
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
      Begin XtremeSuiteControls.FlatEdit txtDebito 
         Height          =   315
         Left            =   -62680
         TabIndex        =   50
         Top             =   5940
         Visible         =   0   'False
         Width           =   1575
         _Version        =   1441793
         _ExtentX        =   2773
         _ExtentY        =   550
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
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
      Begin XtremeSuiteControls.FlatEdit txtCredito 
         Height          =   315
         Left            =   -61120
         TabIndex        =   51
         Top             =   5940
         Visible         =   0   'False
         Width           =   1575
         _Version        =   1441793
         _ExtentX        =   2773
         _ExtentY        =   550
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
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
      Begin XtremeSuiteControls.FlatEdit txtCodBene 
         Height          =   330
         Left            =   2160
         TabIndex        =   28
         Top             =   1440
         Width           =   2415
         _Version        =   1441793
         _ExtentX        =   4260
         _ExtentY        =   582
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
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
      Begin XtremeSuiteControls.ComboBox cboTipoId 
         Height          =   330
         Left            =   6600
         TabIndex        =   54
         Top             =   1440
         Width           =   2415
         _Version        =   1441793
         _ExtentX        =   4260
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
      Begin XtremeSuiteControls.PushButton btnCuenta 
         Height          =   315
         Index           =   2
         Left            =   6480
         TabIndex        =   59
         Top             =   2400
         Width           =   2535
         _Version        =   1441793
         _ExtentX        =   4471
         _ExtentY        =   556
         _StockProps     =   79
         Caption         =   "Valida Sinpe"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseVisualStyle  =   -1  'True
         Appearance      =   17
      End
      Begin XtremeSuiteControls.FlatEdit txtOrigenNombre 
         Height          =   330
         Left            =   2160
         TabIndex        =   60
         Top             =   3240
         Width           =   6855
         _Version        =   1441793
         _ExtentX        =   12091
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
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtOrigenId 
         Height          =   330
         Left            =   2160
         TabIndex        =   61
         Top             =   2880
         Width           =   2415
         _Version        =   1441793
         _ExtentX        =   4260
         _ExtentY        =   582
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
      Begin XtremeSuiteControls.FlatEdit txtOrigenCta 
         Height          =   330
         Left            =   6600
         TabIndex        =   64
         Top             =   2880
         Width           =   2415
         _Version        =   1441793
         _ExtentX        =   4260
         _ExtentY        =   582
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
         Text            =   "CR00000000000000000000"
         BackColor       =   16777152
         Alignment       =   2
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtCorreo 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   5130
            SubFormatType   =   1
         EndProperty
         Height          =   330
         Left            =   2160
         TabIndex        =   66
         Top             =   4080
         Width           =   2415
         _Version        =   1441793
         _ExtentX        =   4260
         _ExtentY        =   582
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
         Alignment       =   1
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtSinpe_Motivo 
         Height          =   330
         Left            =   -66640
         TabIndex        =   73
         Top             =   1560
         Visible         =   0   'False
         Width           =   6855
         _Version        =   1441793
         _ExtentX        =   12091
         _ExtentY        =   582
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
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
      Begin XtremeSuiteControls.FlatEdit txtSinpe_Estado 
         Height          =   330
         Left            =   -66640
         TabIndex        =   74
         Top             =   1080
         Visible         =   0   'False
         Width           =   3015
         _Version        =   1441793
         _ExtentX        =   5318
         _ExtentY        =   582
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
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
      Begin XtremeSuiteControls.FlatEdit txtSinpe_Referencia 
         Height          =   330
         Left            =   -66640
         TabIndex        =   77
         Top             =   2040
         Visible         =   0   'False
         Width           =   3015
         _Version        =   1441793
         _ExtentX        =   5318
         _ExtentY        =   582
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
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
      Begin XtremeSuiteControls.FlatEdit txtSinpe_Fondo 
         Height          =   330
         Left            =   -66640
         TabIndex        =   79
         Top             =   2520
         Visible         =   0   'False
         Width           =   3015
         _Version        =   1441793
         _ExtentX        =   5318
         _ExtentY        =   582
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
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
      Begin XtremeSuiteControls.FlatEdit txtSinpe_DocApl 
         Height          =   330
         Left            =   -66640
         TabIndex        =   81
         Top             =   3000
         Visible         =   0   'False
         Width           =   3015
         _Version        =   1441793
         _ExtentX        =   5318
         _ExtentY        =   582
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
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
      Begin XtremeSuiteControls.FlatEdit txtSinpe_Email 
         Height          =   330
         Left            =   -66640
         TabIndex        =   83
         Top             =   3480
         Visible         =   0   'False
         Width           =   3015
         _Version        =   1441793
         _ExtentX        =   5318
         _ExtentY        =   582
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
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
      Begin XtremeSuiteControls.Label Label7 
         Height          =   255
         Index           =   16
         Left            =   -69640
         TabIndex        =   84
         Top             =   3480
         Visible         =   0   'False
         Width           =   2895
         _Version        =   1441793
         _ExtentX        =   5106
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Email Destino"
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
      Begin XtremeSuiteControls.Label Label7 
         Height          =   255
         Index           =   15
         Left            =   -69640
         TabIndex        =   82
         Top             =   3000
         Visible         =   0   'False
         Width           =   2895
         _Version        =   1441793
         _ExtentX        =   5106
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Documento aplicado"
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
      Begin XtremeSuiteControls.Label Label7 
         Height          =   255
         Index           =   14
         Left            =   -69640
         TabIndex        =   80
         Top             =   2520
         Visible         =   0   'False
         Width           =   2895
         _Version        =   1441793
         _ExtentX        =   5106
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Fondo Aplicado"
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
      Begin XtremeSuiteControls.Label Label7 
         Height          =   255
         Index           =   13
         Left            =   -69640
         TabIndex        =   78
         Top             =   2040
         Visible         =   0   'False
         Width           =   2895
         _Version        =   1441793
         _ExtentX        =   5106
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Código de Referencia"
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
      Begin XtremeSuiteControls.Label Label7 
         Height          =   255
         Index           =   12
         Left            =   -69640
         TabIndex        =   76
         Top             =   1560
         Visible         =   0   'False
         Width           =   2775
         _Version        =   1441793
         _ExtentX        =   4895
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Motivo de Rechazo"
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
      Begin XtremeSuiteControls.Label Label7 
         Height          =   255
         Index           =   11
         Left            =   -69640
         TabIndex        =   75
         Top             =   1080
         Visible         =   0   'False
         Width           =   2895
         _Version        =   1441793
         _ExtentX        =   5106
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Estado de la transferencia"
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
      Begin XtremeSuiteControls.Label lblEstadoSinpe 
         Height          =   330
         Left            =   9120
         TabIndex        =   72
         Top             =   840
         Width           =   1695
         _Version        =   1441793
         _ExtentX        =   2990
         _ExtentY        =   582
         _StockProps     =   79
         Caption         =   "Aceptada"
         BackColor       =   14737632
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   2
         WordWrap        =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label7 
         Height          =   255
         Index           =   10
         Left            =   120
         TabIndex        =   70
         Top             =   5040
         Width           =   1575
         _Version        =   1441793
         _ExtentX        =   2778
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Detalle"
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
      Begin XtremeSuiteControls.Label Label7 
         Height          =   255
         Index           =   9
         Left            =   120
         TabIndex        =   69
         Top             =   4440
         Width           =   1575
         _Version        =   1441793
         _ExtentX        =   2778
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Letras"
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
      Begin XtremeSuiteControls.Label Label7 
         Height          =   255
         Index           =   8
         Left            =   120
         TabIndex        =   68
         Top             =   4080
         Width           =   1575
         _Version        =   1441793
         _ExtentX        =   2778
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Correo"
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
      Begin XtremeSuiteControls.Label Label7 
         Height          =   255
         Index           =   7
         Left            =   120
         TabIndex        =   67
         Top             =   3720
         Width           =   1575
         _Version        =   1441793
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
      Begin XtremeSuiteControls.Label Label7 
         Height          =   255
         Index           =   6
         Left            =   4920
         TabIndex        =   65
         Top             =   2880
         Width           =   1335
         _Version        =   1441793
         _ExtentX        =   2355
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Cuenta Origen"
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
      Begin XtremeSuiteControls.Label Label7 
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   63
         Top             =   2880
         Width           =   1575
         _Version        =   1441793
         _ExtentX        =   2778
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Cédula Origen"
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
      Begin XtremeSuiteControls.Label Label7 
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   62
         Top             =   3240
         Width           =   1575
         _Version        =   1441793
         _ExtentX        =   2778
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Origen Nombre"
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
      Begin XtremeSuiteControls.Label Label7 
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   58
         Top             =   2400
         Width           =   1695
         _Version        =   1441793
         _ExtentX        =   2990
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Cuenta Destino"
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
      Begin XtremeSuiteControls.Label Label7 
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   57
         Top             =   1920
         Width           =   1575
         _Version        =   1441793
         _ExtentX        =   2778
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Beneficiario"
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
      Begin XtremeSuiteControls.Label Label7 
         Height          =   375
         Index           =   1
         Left            =   4800
         TabIndex        =   56
         Top             =   1440
         Width           =   1815
         _Version        =   1441793
         _ExtentX        =   3201
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Tipo de Identificación"
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
      Begin XtremeSuiteControls.Label Label7 
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   55
         Top             =   1440
         Width           =   1935
         _Version        =   1441793
         _ExtentX        =   3413
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Identificación Destino"
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
      Begin XtremeSuiteControls.Label lblMontoLetra 
         Height          =   495
         Left            =   2160
         TabIndex        =   44
         Top             =   4440
         Width           =   8175
         _Version        =   1441793
         _ExtentX        =   14420
         _ExtentY        =   873
         _StockProps     =   79
         Caption         =   "....."
         ForeColor       =   16777215
         BackColor       =   8388608
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
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Totales"
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
         Index           =   0
         Left            =   -63880
         TabIndex        =   15
         Top             =   5940
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "Diferencia"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Index           =   1
         Left            =   -69520
         TabIndex        =   14
         Top             =   5940
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Concepto"
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
         Left            =   960
         TabIndex        =   12
         Top             =   840
         Width           =   1332
      End
      Begin VB.Label Label1 
         Caption         =   "Unidad"
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
         Index           =   6
         Left            =   960
         TabIndex        =   11
         Top             =   480
         Width           =   1332
      End
      Begin VB.Label lblRef 
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
         Left            =   4920
         TabIndex        =   10
         Top             =   4080
         Width           =   1095
      End
      Begin VB.Image imgBloqueo 
         Height          =   480
         Left            =   240
         Picture         =   "frmTES_Transacciones.frx":4D68
         Top             =   540
         Width           =   480
      End
      Begin VB.Label Label6 
         Caption         =   "Divisa/TC"
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
         Left            =   4920
         TabIndex        =   9
         Top             =   3720
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Tipo Origen"
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
         Left            =   6600
         TabIndex        =   8
         Top             =   600
         Width           =   1335
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   12120
      Top             =   240
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTES_Transacciones.frx":5744
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTES_Transacciones.frx":6101
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTES_Transacciones.frx":6A66
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTES_Transacciones.frx":7222
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTES_Transacciones.frx":790A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComCtl2.FlatScrollBar FlatScrollBarX 
      Height          =   252
      Left            =   5160
      TabIndex        =   3
      Top             =   360
      Width           =   492
      _ExtentX        =   873
      _ExtentY        =   450
      _Version        =   393216
      Arrows          =   65536
      Orientation     =   1638401
   End
   Begin MSComCtl2.FlatScrollBar FlatScrollBarY 
      Height          =   255
      Left            =   4560
      TabIndex        =   4
      Top             =   1440
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   450
      _Version        =   393216
      Arrows          =   65536
      Orientation     =   1638401
   End
   Begin MSComctlLib.StatusBar StatusBarx 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   5
      Top             =   8550
      Width           =   13155
      _ExtentX        =   23204
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   6473
            MinWidth        =   6473
            Text            =   "Fecha Solicitud:"
            TextSave        =   "Fecha Solicitud:"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   6473
            MinWidth        =   6473
            Text            =   "Fecha Emisión:"
            TextSave        =   "Fecha Emisión:"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   6473
            MinWidth        =   6473
            Text            =   "Fecha Anulación:"
            TextSave        =   "Fecha Anulación:"
         EndProperty
      EndProperty
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
   Begin MSComctlLib.Toolbar tlb 
      Height          =   330
      Left            =   10680
      TabIndex        =   6
      Top             =   1080
      Width           =   3420
      _ExtentX        =   6033
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
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
   Begin XtremeSuiteControls.ComboBox cboDoc 
      Height          =   312
      Left            =   6360
      TabIndex        =   24
      Top             =   1080
      Width           =   4212
      _Version        =   1441793
      _ExtentX        =   7435
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
   Begin XtremeSuiteControls.ComboBox cbo 
      Height          =   312
      Left            =   1560
      TabIndex        =   25
      Top             =   1080
      Width           =   4692
      _Version        =   1441793
      _ExtentX        =   8281
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
   Begin XtremeSuiteControls.FlatEdit txtDocumento 
      Height          =   330
      Left            =   1560
      TabIndex        =   26
      Top             =   1440
      Width           =   2892
      _Version        =   1441793
      _ExtentX        =   5101
      _ExtentY        =   582
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
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
   Begin XtremeSuiteControls.FlatEdit txtCodigo 
      Height          =   432
      Left            =   2880
      TabIndex        =   27
      Top             =   360
      Width           =   2172
      _Version        =   1441793
      _ExtentX        =   3831
      _ExtentY        =   762
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   12
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
   Begin XtremeSuiteControls.DateTimePicker dtpEmision 
      Height          =   330
      Left            =   7200
      TabIndex        =   53
      Top             =   1440
      Visible         =   0   'False
      Width           =   1332
      _Version        =   1441793
      _ExtentX        =   2350
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
   Begin VB.Label lblEmite 
      BackStyle       =   0  'Transparent
      Caption         =   "Emite"
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
      Left            =   6360
      TabIndex        =   52
      Top             =   1440
      Visible         =   0   'False
      Width           =   1332
   End
   Begin VB.Image imgBloqueado 
      Height          =   192
      Left            =   6600
      Picture         =   "frmTES_Transacciones.frx":8297
      Stretch         =   -1  'True
      Top             =   360
      Visible         =   0   'False
      Width           =   192
   End
   Begin VB.Image imgAutorizado 
      Height          =   192
      Left            =   6240
      Picture         =   "frmTES_Transacciones.frx":8399
      Stretch         =   -1  'True
      Top             =   360
      Visible         =   0   'False
      Width           =   192
   End
   Begin ComctlLib.ImageList imgIconosEstados 
      Left            =   11520
      Top             =   240
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      UseMaskColor    =   0   'False
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   7
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmTES_Transacciones.frx":8482
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmTES_Transacciones.frx":8CD4
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmTES_Transacciones.frx":9526
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmTES_Transacciones.frx":9D78
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmTES_Transacciones.frx":A5CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmTES_Transacciones.frx":AE1C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmTES_Transacciones.frx":B66E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Image imgEstado 
      Height          =   192
      Left            =   5880
      Picture         =   "frmTES_Transacciones.frx":BEC0
      Stretch         =   -1  'True
      Top             =   360
      Visible         =   0   'False
      Width           =   192
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "No. Documento"
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
      TabIndex        =   2
      Top             =   1440
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Cuenta Banco"
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
      Left            =   240
      TabIndex        =   1
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "No. Solicitud"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   0
      Left            =   1440
      TabIndex        =   0
      Top             =   360
      Width           =   1455
   End
   Begin VB.Image imgBanner 
      Height          =   1215
      Left            =   0
      Stretch         =   -1  'True
      Top             =   -240
      Width           =   13815
   End
End
Attribute VB_Name = "frmTES_Transacciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vEdita As Boolean, vCodigo As Long, vPaso As Boolean
Dim vScrollX As Boolean, vScrollY As Boolean
Dim gBanco As String, gConcepto As String, gUnidad As String, gDocumento As String
Dim gTipoCambio As Currency, gDivisa As String, gVariacion As Integer, gDivisaDesc As String, gDivisaCurrency As String
Dim gDivisaLocal As Integer, gVariacionAsiento As Integer, strMovimiento As String
Dim vEstado As String, gFecha As Date, gAutoEmite As Boolean



Private Sub sbDetalle_Consulta()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

Group_Detalle.Visible = True

strSQL = "exec spTes_Consulta_Afectacion_Modulos " & txtCodigo.Text

Call OpenRecordSet(rs, strSQL)

With vGrid_Detalle
   .MaxRows = 0
   Do While Not rs.EOF
     .MaxRows = .MaxRows + 1
     .Row = .MaxRows
     .col = 1
     .Text = rs!cod_Factura
     .col = 2
     .Text = Format(rs!Creacion_Fecha, "dd/mm/yyyy")
     .col = 3
     .Text = Format(rs!Total, "Standard")
     .col = 4
     .Text = CStr(rs!Npago)
     .col = 5
     .Text = Format(rs!Monto, "Standard")
     .col = 6
     .Text = rs!Identificacion & ""
     .col = 7
     .Text = rs!DESCRIPCION & ""
     
     rs.MoveNext
   Loop
   rs.Close
End With

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub btnCuenta_Click(Index As Integer)
Dim strSQL As String

On Error GoTo vError

Select Case Index
  Case 0 'Ajustar
        strSQL = "exec spTes_Cuenta_Bancaria_Cambio " & txtCodigo.Text & ",'" & Trim(cboCuenta.Text) & "','" & glogon.Usuario & "'"
        Call ConectionExecute(strSQL)
        
        If Not glogon.error Then
            MsgBox "Cambio de Cuenta Bancaria realizado satisfactoriamente!", vbInformation
        End If
  Case 1 'Consultar
    GLOBALES.gTag = Trim(txtCodBene.Text)
    GLOBALES.gTag2 = "TES"

    frmCC_Cuentas_Bancarias.Show vbModal

End Select

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub btnDetalle_Cierra_Click()
Group_Detalle.Visible = False
End Sub

Private Sub btnOpciones_Click(Index As Integer)
If vCodigo = 0 Then Exit Sub

GLOBALES.gTag = txtCodigo.Text

Select Case Index
  Case 0 ' "Anular"
    If vEstado = "I" Or vEstado = "E" Or vEstado = "T" Then
       Call sbFormsCall("frmTES_AnulacionDoc", 1, , , False, Me)
    Else
      MsgBox "La solicitud " & txtCodigo & " no puede ser anulada " & vbCrLf & _
             "ya que su estado es solicitada o anulada"
    End If
     
  Case 1 '"Reclasificar"
     Call sbFormsCall("frmTES_Reclasificacion", 1, , , False, Me)
    
  Case 2 '"ReImpresion"
     Call sbFormsCall("frmTES_ReImpresion", 1, , , False, Me)
     
  Case 3 ' "Fechas"
     Call sbFormsCall("frmTES_CambioFechas", 1, , , False, Me)

  Case 4 '"Reponer"
      If vEstado = "I" Or vEstado = "E" Or vEstado = "T" Then
            Call sbFormsCall("frmTES_Reposicion", 1, , , False, Me)
      Else
            MsgBox "La solicitud " & txtCodigo & " no puede ser repuesta " & vbCrLf & _
                   "ya que su estado es solicitada o anulada"
      End If
      
   Case 5 'Detalle
      Call sbDetalle_Consulta
      
End Select

If Index <> 5 Then
    Call sbConsulta(vCodigo)
End If

End Sub

Private Sub cbo_Click()
If vPaso Then Exit Sub

If cbo.ListCount = 0 Then
   cbo.AddItem " "
    If TypeOf cbo Is XtremeSuiteControls.ComboBox Then
        cbo.ItemData(cbo.ListCount - 1) = CStr(0)
    Else
        cbo.ItemData(cbo.NewIndex) = 0
    End If
      cbo.Text = " "
End If

Call sbTesTiposDocsCargaCboAcceso(cboDoc, glogon.Usuario, cbo.ItemData(cbo.ListIndex), "S")
Call sbTesConceptosCargaCbo(cboConcepto, glogon.Usuario, cbo.ItemData(cbo.ListIndex))
Call sbTesUnidadesCargaCbo(cboUnidad, glogon.Usuario, cbo.ItemData(cbo.ListIndex))
Call sbTesControlDivisas(cbo.ItemData(cbo.ListIndex), GLOBALES.gEnlace)

End Sub

Private Sub cbo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then cboDoc.SetFocus
End Sub

Private Sub cboConcepto_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtCodBene.SetFocus
End Sub

Private Sub cboDoc_Click()
If vPaso Then Exit Sub

Dim strSQL As String, rs As New ADODB.Recordset

strSQL = "select dbo.fxTes_DocumentoAutoEmite(" & cbo.ItemData(cbo.ListIndex) & ",'" & cboDoc.ItemData(cboDoc.ListIndex) & "') as 'AutoEmite'"
Call OpenRecordSet(rs, strSQL)
    gAutoEmite = IIf((rs!AutoEmite = 1), True, False)
rs.Close

dtpEmision.Visible = gAutoEmite
lblEmite.Visible = gAutoEmite
dtpEmision.Value = gFecha

End Sub

Private Sub cboDoc_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtDocumento.SetFocus
End Sub


Private Sub cboTipos_Click()
On Error GoTo vError

If Not vEdita Then
    txtCodBene.Text = ""
    txtBeneficiario.Text = ""
    txtCodBene.SetFocus
End If

vError:
End Sub

Private Sub cboTipos_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then txtCodBene.SetFocus
End Sub

Private Sub cboUnidad_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then cboConcepto.SetFocus
End Sub

Private Sub sbSumaDebitosCreditos()
Dim x As Integer, TC As Currency
  
On Error GoTo vError
  
  txtDebito.Text = "0"
  txtCredito.Text = "0"
  For x = 1 To vGrid.MaxRows
     vGrid.Row = x
     TC = 1
      
     vGrid.col = 7
     txtDebito.Text = CCur(txtDebito.Text) + (CCur(IIf(vGrid.Text = "", 0, vGrid.Text)) * fxSys_Tipo_Cambio_Apl(TC))
     vGrid.col = 8
     txtCredito.Text = CCur(txtCredito.Text) + (CCur(IIf(vGrid.Text = "", 0, vGrid.Text)) * fxSys_Tipo_Cambio_Apl(TC))
  Next x
  
  txtDebito.Text = Format(txtDebito.Text, "Standard")
  txtCredito.Text = Format(txtCredito.Text, "Standard")
  txtDiferencia.Text = Format(CCur(txtDebito.Text) - CCur(txtCredito.Text), "Standard")

vError:

End Sub

Private Sub sbConsultaDocumento()
Dim rs As New ADODB.Recordset, strSQL As String

On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = "select Nsolicitud from Tes_Transacciones where ndocumento = '" & txtDocumento.Text & "' and id_banco = " _
       & cbo.ItemData(cbo.ListIndex) & " and Tipo = '" & cboDoc.ItemData(cboDoc.ListIndex) & "'"
Call OpenRecordSet(rs, strSQL)
If Not rs.BOF And Not rs.EOF Then
 Call sbConsulta(rs!NSolicitud)
End If
rs.Close

Me.MousePointer = vbDefault

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub FlatScrollBarX_Change()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

If vScrollX Then
    strSQL = "select Top 1 nsolicitud from Tes_Transacciones"
'           & " where Tipo = '" & fxTipoASEDoc(cboTipo.Text) & "'"
    
    If txtCodigo = "" And FlatScrollBarX.Value = 1 Then txtCodigo.Text = "0"
    If txtCodigo = "" And FlatScrollBarX.Value = 0 Then txtCodigo.Text = "999999999"
    
    If FlatScrollBarX.Value = 1 Then
       strSQL = strSQL & " where nsolicitud > " & txtCodigo & " order by nsolicitud asc"
    Else
       strSQL = strSQL & " where nsolicitud < " & txtCodigo & " order by nsolicitud desc"
    End If
    
    Call OpenRecordSet(rs, strSQL)
    If Not rs.EOF And Not rs.BOF Then
      txtCodigo = rs!NSolicitud
      Call sbConsulta(txtCodigo)
    End If
    rs.Close
End If

vScrollX = False
    FlatScrollBarX.Value = 0
vScrollX = True

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub FlatScrollBarY_Change()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

If vScrollY Then
    strSQL = "select Top 1 ndocumento from Tes_Transacciones" _
           & " where Tipo = '" & cboDoc.ItemData(cboDoc.ListIndex) & "' and id_banco = " & cbo.ItemData(cbo.ListIndex)
    
    If FlatScrollBarY.Value = 1 Then
       strSQL = strSQL & " and ndocumento > '" & txtDocumento & "' order by ndocumento asc"
    Else
       strSQL = strSQL & " and ndocumento < '" & txtDocumento & "' order by ndocumento desc"
    End If
    
    Call OpenRecordSet(rs, strSQL)
    If Not rs.EOF And Not rs.BOF Then
      txtDocumento = rs!nDocumento & ""
      Call sbConsultaDocumento
    End If
    rs.Close
End If

vScrollY = False
FlatScrollBarY.Value = 0
vScrollY = True

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
  MsgBox "Consulte a Su Administrador de Base de Datos, sobre Transacciones con TOP y Record Count", vbInformation

End Sub


Private Sub ssTab_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)

Select Case Item.Index
  Case 0 'Principal
    Group_Detalle.Visible = False
  Case 1 'Asiento
    Call sbCargaAsiento
  Case 2 'Seguimiento
    Call SbSeguimiento
End Select

End Sub

Private Sub tlbAux_ButtonClick(ByVal Button As MSComctlLib.Button)
If vCodigo = 0 Then Exit Sub

GLOBALES.gTag = txtCodigo.Text

Select Case Button.Key
  Case "Anular"
    If vEstado = "I" Or vEstado = "E" Or vEstado = "T" Then
       Call sbFormsCall("frmTES_AnulacionDoc", 1, , , False, Me)
    Else
      MsgBox "La solicitud " & txtCodigo & " no puede ser anulada " & vbCrLf & _
             "ya que su estado es solicitada o anulada"
    End If
     
  Case "Reclasificar"
     Call sbFormsCall("frmTES_Reclasificacion", 1, , , False, Me)
    
  Case "ReImpresion"
     Call sbFormsCall("frmTES_ReImpresion", 1, , , False, Me)
     
  Case "Fechas"
     Call sbFormsCall("frmTES_CambioFechas", 1, , , False, Me)

  Case "Reponer"
      If vEstado = "I" Or vEstado = "E" Or vEstado = "T" Then
            Call sbFormsCall("frmTES_Reposicion", 1, , , False, Me)
      Else
            MsgBox "La solicitud " & txtCodigo & " no puede ser repuesta " & vbCrLf & _
                   "ya que su estado es solicitada o anulada"
      End If

End Select

Call sbConsulta(vCodigo)


End Sub

Private Sub txtDocumento_LostFocus()
If txtDocumento <> "" Then Call sbConsultaDocumento
End Sub

Private Sub txtTipoCambio_KeyDown(KeyCode As Integer, Shift As Integer)
  


If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then cboCuenta.SetFocus



End Sub

Private Sub txtTipoCambio_LostFocus()
' If CCur(txtTipoCambio) > (gTipoCambio + gVariacion) Or CCur(txtTipoCambio) < (gTipoCambio - gVariacion) Then
'   MsgBox "El monto del tipo de cambio no esta dentro del rango de la variación" & vbCrLf & "La variacion es de " & gVariacion
'   txtTipoCambio = gTipoCambio
'   Exit Sub
'End If
End Sub


Private Sub vGrid_KeyDown(KeyCode As Integer, Shift As Integer)
Dim i As Variant, lng As Long, vTemp(8) As Variant, x As Integer
Dim vCuenta As String, strDivisa As String
Dim iColumna As Integer

'No permite Borrar la Primer linea, las demás SI
If KeyCode = vbKeyDelete And vGrid.ActiveRow > 1 Then
  lng = 1
  vGrid.Row = vGrid.ActiveRow
  vGrid.col = vGrid.MaxCols
  If vGrid.Text <> "" Then 'Existe en la Base de datos
    'Preguntar y si la respuesta es afirmativa eliminar de la Base de datos
  
  
  End If
  
  For lng = vGrid.ActiveRow To vGrid.MaxRows - 1
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
  
End If 'KeyCode = vbKeyDelete And vGrid.ActiveRow > 1


'Consulta cuenta / Codigo
If KeyCode = vbKeyF4 And vGrid.ActiveCol = 1 And vGrid.ActiveRow > 1 Then
  Call sbgCntCuentaConsulta("C")
  vGrid.col = vGrid.ActiveCol
  vGrid.Row = vGrid.ActiveRow
  vGrid.Text = gBusquedas.Resultado
End If



'Consulta cuenta / descripcion
If KeyCode = vbKeyF4 And vGrid.ActiveCol = 6 And vGrid.ActiveRow > 1 Then
  Call sbgCntCuentaConsulta("D")
  vGrid.col = 1
  vGrid.Row = vGrid.ActiveRow
  vGrid.Text = gBusquedas.Resultado
End If

If KeyCode = vbKeyF4 And (vGrid.ActiveCol = 5 Or vGrid.ActiveCol = 7 Or vGrid.ActiveCol = 8) And vGrid.ActiveRow > 1 Then
   vGrid.Row = vGrid.ActiveRow
   vGrid.col = 1
   gCntX_TipoCambio.Cuenta = fxgCntCuentaFormato(False, vGrid.Text, 0)
   
   
   gCntX_TipoCambio.fecha = gFecha
   
   
   'If Val(vGrid.Text) > 0 And strMovimiento = "A" Then
   If strMovimiento <> "D" Then
      vGrid.col = 7
      iColumna = 7
      gCntX_TipoCambio.Monto_Actual = IIf(vGrid.Text = "", 0, vGrid.Text)
   Else
      iColumna = 8
      vGrid.col = 8
      gCntX_TipoCambio.Monto_Actual = IIf(vGrid.Text = "", 0, vGrid.Text)
   End If
      vGrid.col = 4
      gCntX_TipoCambio.Moneda = vGrid.Text
      
      If gCntX_TipoCambio.Moneda = gCntX_Parametros.DivisaLocal Then Exit Sub
      vGrid.col = 5
       If vGrid.Text = "" Or vGrid.Text = "0.0000" Then
         gCntX_TipoCambio.TC_Actual = 1
       Else
         gCntX_TipoCambio.TC_Actual = vGrid.Text
       End If
       frmTES_TipoCambio.Show vbModal
       If gCntX_TipoCambio.Paso Then
          vGrid.Row = vGrid.ActiveRow
          vGrid.col = 5
          vGrid.Text = CStr(gCntX_TipoCambio.TC_Nuevo)
          vGrid.col = iColumna
          vGrid.Text = CStr(gCntX_TipoCambio.Monto_Nuevo)
          Call sbCalculaMontoDivisaForanea(vGrid.ActiveRow)
          Call sbSumaDebitosCreditos
       End If
   ' End If
      
End If




'Consulta unidad
If KeyCode = vbKeyF4 And vGrid.ActiveCol = 2 And vGrid.ActiveRow > 1 Then
  gBusquedas.Resultado = ""
  gBusquedas.Resultado2 = ""
  gBusquedas.Columna = "descripcion"
  gBusquedas.Orden = "descripcion"
  gBusquedas.Filtro = " and Activa = 1 and cod_contabilidad = " & GLOBALES.gEnlace
  gBusquedas.Consulta = "select cod_unidad,descripcion from CntX_unidades"
  frmBusquedas.Show vbModal
    
  vGrid.col = vGrid.ActiveCol
  vGrid.Row = vGrid.ActiveRow
  vGrid.Text = gBusquedas.Resultado
  
  vGrid.CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
  vGrid.CellNote = gBusquedas.Resultado2
  vGrid.TextTip = TextTipFixed
  
End If



'Consulta Centro de Costo
If KeyCode = vbKeyF4 And vGrid.ActiveCol = 3 And vGrid.ActiveRow > 1 Then
  
  vGrid.col = 2
  gBusquedas.Resultado = ""
  gBusquedas.Resultado2 = ""
  gBusquedas.Filtro = " and C.cod_Contabilidad = " & GLOBALES.gEnlace
  gBusquedas.Columna = "descripcion"
  gBusquedas.Orden = "descripcion"
  gBusquedas.Consulta = "select C.COD_CENTRO_COSTO,C.descripcion" _
                      & " from CNTX_CENTRO_COSTOS C inner join CNTX_UNIDADES_CC A on C.COD_CENTRO_COSTO = A.COD_CENTRO_COSTO" _
                      & " and C.cod_contabilidad = A.cod_Contabilidad" _
                      & " and A.cod_unidad = '" & vGrid.Text & "'"
  frmBusquedas.Show vbModal
    
  vGrid.col = 3
  vGrid.col = vGrid.ActiveCol
  vGrid.Row = vGrid.ActiveRow
  vGrid.Text = gBusquedas.Resultado
  
  vGrid.CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
  vGrid.CellNote = gBusquedas.Resultado2
  vGrid.TextTip = TextTipFixed
  
End If



If (KeyCode = vbKeyReturn Or KeyCode = vbKeyTab) Then
    vGrid.col = vGrid.ActiveCol
    vGrid.Row = vGrid.ActiveRow
    
    Select Case vGrid.ActiveCol
      Case 1 'Cuenta
        vGrid.Text = fxgCntCuentaFormato(True, vGrid.Text, 0)
        
        If fxgCntCuentaValida(fxgCntCuentaFormato(False, vGrid.Text, 0)) Then
            vCuenta = vGrid.Text
            'CAMBIOS PARA EL TIPO DE CAMBIO Y DE LA DIVISA
            vGrid.col = 6
            vGrid.Text = fxgCntCuentaDesc(fxgCntCuentaFormato(False, vCuenta, 0))
            vGrid.col = 4
            vGrid.Text = fxDivisaCuenta(fxgCntCuentaFormato(False, vCuenta, 0))
            strDivisa = vGrid.Text
            vGrid.col = 5
            
            If vGrid.Row = 1 Then
                vGrid.Text = txtTipoCambio.Text
            Else
                vGrid.Text = fxDivisaTipoCambio(strDivisa)
            End If
            If vGrid.Text = 1 Then
               vGrid.Lock = True
               vGrid.Protect = True
            Else
               vGrid.Lock = False
               vGrid.Protect = False
            End If
        Else
          MsgBox "Cuenta no es válida : " & vbCrLf & " - No Existe o No Acepta Movimientos" _
                 & vbCrLf & " - VERIFIQUE O MODIFIQUE EN EL CATALAGO DE CUENTAS", vbCritical
        End If
        
      Case 2
        'Buscar la unidad
        If fxTesUnidadValida(cbo.ItemData(cbo.ListIndex), glogon.Usuario, cboUnidad.ItemData(cboUnidad.ListIndex)) Then
          vGrid.CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
          vGrid.CellNote = fxgCntUnidad(vGrid.Text)
          vGrid.TextTip = TextTipFixed
        Else
          MsgBox "La unidad de negocio no es válida, o no se tiene asignada a este usuario", vbCritical
        End If
      
      Case 3 'Describe el centro de costo
          vGrid.CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
          vGrid.CellNote = fxgCntCentroCostos(vGrid.Text)
          vGrid.TextTip = TextTipFixed
      
        
      Case 7 'Debe
        If Val(vGrid.Text) > 0 Then
            vGrid.col = vGrid.ActiveCol + 1
            vGrid.Row = vGrid.ActiveRow
            vGrid.Text = 0
        
            Call sbSumaDebitosCreditos
            
        End If
      
      Case 8 'Haber
        If Val(vGrid.Text) > 0 Then
            vGrid.col = vGrid.ActiveCol - 1
            vGrid.Row = vGrid.ActiveRow
            vGrid.Text = 0
        
            Call sbSumaDebitosCreditos
        End If
      
        If vGrid.MaxRows = vGrid.Row Then
            
            vGrid.col = 2
            vTemp(0) = vGrid.Text
            vTemp(1) = vGrid.CellNote
            
            vGrid.col = 3
            vTemp(2) = vGrid.Text
            vTemp(3) = vGrid.CellNote
            
            vGrid.MaxRows = vGrid.MaxRows + 1
            vGrid.Row = vGrid.MaxRows
        
            vGrid.col = 2
            vGrid.Text = vTemp(0)
            vGrid.CellNote = vTemp(1)
            
            vGrid.col = 3
            vGrid.Text = vTemp(2)
            vGrid.CellNote = vTemp(3)
        End If
    
    End Select

End If

If KeyCode = vbKeyInsert And vGrid.ActiveRow > 1 Then
    vGrid.Row = vGrid.ActiveRow
            vGrid.col = 2
            vTemp(0) = vGrid.Text
            vTemp(1) = vGrid.CellNote
            
            vGrid.col = 3
            vTemp(2) = vGrid.Text
            vTemp(3) = vGrid.CellNote
    
    vGrid.MaxRows = vGrid.MaxRows + 1
    vGrid.InsertRows vGrid.ActiveRow, 1
    vGrid.Row = vGrid.ActiveRow
            
            vGrid.col = 2
            vGrid.Text = vTemp(0)
            vGrid.CellNote = vTemp(1)
            
            vGrid.col = 3
            vGrid.Text = vTemp(2)
            vGrid.CellNote = vTemp(3)
    vGrid.col = 1
End If


Call sbCalculaMontoDivisaForanea(vGrid.ActiveRow)

'Bloquea la primer lin
vGrid.Row = 1
For i = 1 To vGrid.MaxCols
    vGrid.col = i
    vGrid.Lock = True
    vGrid.Protect = True
Next i

End Sub

Private Sub Form_Activate()
 vModulo = 9
End Sub

Private Sub Form_Load()

On Error GoTo vError
 
 vModulo = 9
 vGrid.AppearanceStyle = fxGridStyle
 Set imgBanner.Picture = frmContenedor.imgBanner_01.Picture
 
 vScrollX = False
 FlatScrollBarX.Value = 0
 vScrollX = True
 
 vScrollY = False
 FlatScrollBarY.Value = 0
 vScrollY = True
 
 gBanco = ""
 gConcepto = ""
 gUnidad = ""
 gDocumento = ""
 
 
cboTipos.Clear
cboTipos.AddItem "Personas"
cboTipos.ItemData(cboTipos.ListCount - 1) = CStr(1)
cboTipos.AddItem "Bancos"
cboTipos.ItemData(cboTipos.ListCount - 1) = CStr(2)
cboTipos.AddItem "Proveedores"
cboTipos.ItemData(cboTipos.ListCount - 1) = CStr(3)
cboTipos.AddItem "Acreedores"
cboTipos.ItemData(cboTipos.ListCount - 1) = CStr(4)
cboTipos.AddItem "Cuentas"
cboTipos.ItemData(cboTipos.ListCount - 1) = CStr(5)
cboTipos.AddItem "Empleados"
cboTipos.ItemData(cboTipos.ListCount - 1) = CStr(6)
cboTipos.AddItem "Directos"
cboTipos.ItemData(cboTipos.ListCount - 1) = CStr(7)

 
 vEdita = True
 Call sbToolBarIconos(tlb, False)
 Call sbToolBar(tlb, "nuevo")
 Call sbLimpiaPantalla

 Call Formularios(Me)
 Call RefrescaTags(Me)
 
Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbExclamation
 
End Sub

Private Sub sbLimpiaPantalla()
Dim i As Byte

vCodigo = 0
vPaso = True
vEstado = "P"
Call sbTesBancoCargaCboAccesoGestion(cbo, glogon.Usuario, "Solicita")

vPaso = False
Call cbo_Click



cboTipos.Text = "Personas"

StatusBarX.Panels(1).Text = "Fecha Solicitud:"
StatusBarX.Panels(1).ToolTipText = ""
StatusBarX.Panels(2).Text = "Fecha Emisión:"
StatusBarX.Panels(2).ToolTipText = ""
StatusBarX.Panels(3).Text = "Fecha Anulación:"
StatusBarX.Panels(3).ToolTipText = ""

txtDocumento = ""

gFecha = fxFechaServidor

dtpEmision.Visible = gAutoEmite
lblEmite.Visible = gAutoEmite
dtpEmision.Value = gFecha


txtCodBene = ""
txtBeneficiario = ""
txtMonto = 0
lblMontoLetra.Caption = ""
txtDetalle = ""

txtRef.Text = ""
txtRef_01.Text = ""
txtRef_02.Text = ""
txtRef_03.Text = ""


cboCuenta.Clear

txtRef = ""
imgEstado.Visible = False
imgAutorizado.Visible = False
imgBloqueado.Visible = False

ssTab.Item(0).Selected = True
For i = 1 To ssTab.ItemCount - 1
 ssTab.Item(i).Enabled = False
Next i


vGrid.MaxRows = 0
Group_Detalle.Visible = False


End Sub


Private Sub sbCargaAsiento()
Dim strSQL As String, rs As New ADODB.Recordset
Dim i As Integer, rsTmp As New ADODB.Recordset
Dim strDivisa As String, cTipoCambio As Currency
Dim vEstado As String

On Error GoTo vError

Me.MousePointer = vbHourglass
vGrid.MaxCols = 8

cTipoCambio = CCur(txtTipoCambio.Text)  '  fxDivisaTipoCambio(gDivisa)
vEstado = "P"


If Not IsNumeric(txtMonto) Then
  vGrid.MaxRows = 0
  Me.MousePointer = vbDefault
  Exit Sub
End If

If CCur(txtMonto) = 0 Then
  vGrid.MaxRows = 0
  Me.MousePointer = vbDefault
  Exit Sub
End If


If vCodigo > 0 Then
   strSQL = "select C.cod_cuenta_Mask as 'Cod_Cuenta',C.descripcion,D.debehaber,D.monto,D.cod_unidad,Ch.Estado" _
          & ",U.descripcion as UnidadX,D.cod_cc,X.descripcion as CCX,Ch.id_Banco,D.tipo_cambio,D.cod_divisa" _
          & " from Tes_Trans_Asiento D inner join Tes_Transacciones Ch on D.nsolicitud = Ch.Nsolicitud" _
          & " inner join CntX_Cuentas C on D.cuenta_contable = C.cod_cuenta and C.cod_contabilidad = " & GLOBALES.gEnlace _
          & " left join CntX_unidades U on D.cod_unidad = U.cod_unidad and U.cod_contabilidad = " & GLOBALES.gEnlace _
          & " left join CNTX_CENTRO_COSTOS X on D.cod_cc = X.COD_CENTRO_COSTO and X.cod_contabilidad = " & GLOBALES.gEnlace _
          & " where D.nsolicitud = " & vCodigo _
          & " order by D.linea"

    Call OpenRecordSet(rs, strSQL, 0)
    
    
    vGrid.MaxRows = 1
    vGrid.Row = vGrid.MaxRows
    
    
    Do While Not rs.EOF
      vEstado = rs!Estado
      vGrid.Row = vGrid.MaxRows
      
      For i = 1 To vGrid.MaxCols
        vGrid.col = i
        Select Case i
         Case 1
            
            If cbo.ItemData(cbo.ListIndex) = CStr(rs!Id_Banco) Then
             'Se mantiene el banco
             vGrid.Text = rs!cod_cuenta ' fxgCntCuentaFormato(True, CStr(rs!cod_cuenta))
            Else
             'Cambio de Banco
                strSQL = "select C.cod_cuenta_Mask as 'Cod_Cuenta',C.descripcion, C.cod_divisa" _
                       & " from CntX_Cuentas C inner join Tes_Bancos B on C.cod_Cuenta = B.CtaConta" _
                       & " Where B.id_banco = " & cbo.ItemData(cbo.ListIndex) _
                       & " and C.cod_contabilidad = " & GLOBALES.gEnlace
                Call OpenRecordSet(rsTmp, strSQL, 0)
                 vGrid.col = 1
                 vGrid.Text = rsTmp!cod_cuenta ' fxgCntCuentaFormato(True, rsTmp!Cod_Cuenta)
                 vGrid.col = 4
                 vGrid.Text = rsTmp!COD_DIVISA ' fxDivisaCuenta(rsTmp!Cod_Cuenta)
                 strDivisa = vGrid.Text
                 vGrid.col = 5
                 vGrid.Text = fxDivisaTipoCambio(strDivisa)
                 If vGrid.Text = 1 Then
                    vGrid.Lock = True
                    vGrid.Protect = True
                Else
                    vGrid.Lock = False
                    vGrid.Protect = False
                End If
                 vGrid.col = 6
                 vGrid.Text = rsTmp!DESCRIPCION 'fxgCntCuentaDesc(rsTmp!Cod_Cuenta)
                rsTmp.Close
            
            End If
            
         
         Case 2
            vGrid.Text = rs!Cod_Unidad & ""
            vGrid.CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
            vGrid.CellNote = rs!UnidadX & ""
            vGrid.TextTip = TextTipFixed
         
         
         Case 3
            vGrid.Text = rs!cod_cc & ""
            vGrid.CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
            vGrid.CellNote = rs!CCX & ""
            vGrid.TextTip = TextTipFixed
        Case 4
           vGrid.Text = rs!COD_DIVISA
        Case 5
           vGrid.Text = CCur(rs!TIPO_CAMBIO)
         Case 6
            'Si se realiza un cambio de banco, el sistema describe la cuenta en el case 1
            If cbo.ItemData(cbo.ListIndex) = CStr(rs!Id_Banco) Then
                vGrid.Text = CStr(rs!DESCRIPCION)
            End If
         Case 7 'Debitos
           If rs!debehaber = "D" Then
              vGrid.Text = CStr(rs!Monto)
           Else
             vGrid.Text = "0"
           End If
         Case 8 'Creditos
           If rs!debehaber = "D" Then
             vGrid.Text = "0"
           Else
               vGrid.Text = CStr(rs!Monto)
           End If
           
        End Select
        If i = 1 Then strMovimiento = rs!debehaber
      Next i
      vGrid.MaxRows = vGrid.MaxRows + 1
      
      rs.MoveNext
    Loop
    rs.Close
    vGrid.MaxRows = vGrid.MaxRows - 1

Else 'vCodigo > 0
  'Llena Segun Datos en Pantalla
    If vGrid.MaxRows > 1 Then
       i = MsgBox("Ya Existe un Asiento Creado, desea reemplazar el actual", vbYesNo)
       If i = vbNo Then
        Me.MousePointer = vbDefault
        Exit Sub
       End If
    End If
  
  vGrid.MaxRows = 0
  vGrid.MaxRows = 2
  vGrid.Row = 1
  If cbo.Text <> "" Then
'    strSQL = "select ctaConta from Tes_Bancos where id_banco = " & cbo.ItemData(cbo.ListIndex)
    
    strSQL = "select C.cod_cuenta_Mask as 'Cod_Cuenta',C.descripcion, C.cod_divisa" _
           & " from CntX_Cuentas C inner join Tes_Bancos B on C.cod_Cuenta = B.CtaConta" _
           & " Where B.id_banco = " & cbo.ItemData(cbo.ListIndex) _
           & " and C.cod_contabilidad = " & GLOBALES.gEnlace
    
    Call OpenRecordSet(rs, strSQL)
     vGrid.col = 1
     vGrid.Text = rs!cod_cuenta ' fxgCntCuentaFormato(True, rs!ctaConta)
     vGrid.col = 4
     vGrid.Text = rs!COD_DIVISA ' fxDivisaCuenta(rs!ctaConta)
     strDivisa = vGrid.Text
     vGrid.col = 5
     vGrid.Text = cTipoCambio  'fxDivisaTipoCambio(strDivisa)
     If vGrid.Text = 1 Then
        vGrid.Lock = True
        vGrid.Protect = True
    Else
        vGrid.Lock = False
        vGrid.Protect = False
    End If

     vGrid.col = 6
     vGrid.Text = rs!DESCRIPCION ' fxgCntCuentaDesc(rs!ctaConta)
    rs.Close
    
    vGrid.col = 2
    vGrid.Text = cboUnidad.ItemData(cboUnidad.ListIndex)
    strSQL = "select descripcion from CntX_Unidades where cod_unidad = '" & vGrid.Text & "' and cod_contabilidad = " & GLOBALES.gEnlace
    Call OpenRecordSet(rs, strSQL)
    If Not rs.EOF And Not rs.BOF Then
     vGrid.CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
     vGrid.CellNote = rs!DESCRIPCION
     vGrid.TextTip = TextTipFixed
    End If
    rs.Close
     
    vGrid.Row = 2
'    strSQL = "select cod_cuenta from Tes_Conceptos where cod_concepto = '" & cboConcepto.ItemData(cboConcepto.ListIndex) & "'"
    
    strSQL = "select C.cod_cuenta_Mask as 'Cod_Cuenta',C.descripcion, C.cod_divisa" _
           & " from CntX_Cuentas C inner join Tes_Conceptos B on C.cod_Cuenta = B.cod_cuenta" _
           & " Where B.cod_concepto = '" & cboConcepto.ItemData(cboConcepto.ListIndex) & "'" _
           & " and C.cod_contabilidad = " & GLOBALES.gEnlace
    Call OpenRecordSet(rs, strSQL)
    If Not rs.EOF And Not rs.BOF Then
     vGrid.col = 1
     vGrid.Text = rs!cod_cuenta ' fxgCntCuentaFormato(True, rs!Cod_Cuenta)
     vGrid.col = 4
     vGrid.Text = rs!COD_DIVISA ' fxDivisaCuenta(rs!cod_cuenta)
     strDivisa = vGrid.Text
     vGrid.col = 5
     vGrid.Text = fxDivisaTipoCambio(strDivisa)
     If vGrid.Text = 1 Then
         vGrid.Lock = True
         vGrid.Protect = True
     Else
        vGrid.Lock = False
        vGrid.Protect = False
     End If
     
    vGrid.col = 6
     vGrid.Text = rs!DESCRIPCION ' fxgCntCuentaDesc(rs!cod_cuenta)
    End If
    rs.Close
    
    vGrid.col = 2
    vGrid.Text = cboUnidad.ItemData(cboUnidad.ListIndex)
    strSQL = "select descripcion from CntX_Unidades where cod_unidad = '" & vGrid.Text & "' and cod_contabilidad = " & GLOBALES.gEnlace
    Call OpenRecordSet(rs, strSQL)
    If Not rs.EOF And Not rs.BOF Then
     vGrid.CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
     vGrid.CellNote = rs!DESCRIPCION
     vGrid.TextTip = TextTipFixed
    End If
    rs.Close
    
    strMovimiento = fxTesTiposDocAsiento(cboDoc.ItemData(cboDoc.ListIndex))
    If fxTesTiposDocAsiento(cboDoc.ItemData(cboDoc.ListIndex)) = "D" Then
       vGrid.Row = 1
       vGrid.col = 7
'       If cTipoCambio > 1 Then
'         vGrid.Text = CStr(CCur(txtMonto * cTipoCambio))
'       Else
'         vGrid.Text = CStr(CCur(txtMonto))
'       End If
          
       vGrid.Text = CStr(CCur(txtMonto * fxSys_Tipo_Cambio_Apl(cTipoCambio)))
          
       vGrid.col = 8
       vGrid.Text = "0"

       vGrid.Row = 2
       vGrid.col = 7
       vGrid.Text = "0"
       vGrid.col = 8
'       If cTipoCambio > 1 Then
'         vGrid.Text = CStr(CCur(txtMonto * cTipoCambio))
'       Else
'         vGrid.Text = CStr(CCur(txtMonto))
'       End If
       vGrid.Text = CStr(CCur(txtMonto * fxSys_Tipo_Cambio_Apl(cTipoCambio)))

    Else
       vGrid.Row = 1
       vGrid.col = 7
       vGrid.Text = "0"
       vGrid.col = 8
'       If cTipoCambio > 1 Then
'         vGrid.Text = CStr(CCur(txtMonto * cTipoCambio))
'       Else
'         vGrid.Text = CStr(CCur(txtMonto))
'       End If
       
       vGrid.Text = CStr(CCur(txtMonto * fxSys_Tipo_Cambio_Apl(cTipoCambio)))
       
       vGrid.Row = 2
       vGrid.col = 7
       
'       If cTipoCambio > 1 Then
'         vGrid.Text = CStr(CCur(txtMonto * cTipoCambio))
'       Else
'         vGrid.Text = CStr(CCur(txtMonto))
'       End If
       vGrid.Text = CStr(CCur(txtMonto * fxSys_Tipo_Cambio_Apl(cTipoCambio)))

       vGrid.col = 8
       vGrid.Text = "0"
    End If
    
  Else
   vGrid.MaxRows = 0
  End If

End If

'Verifica cambio en el monto
If vGrid.MaxRows > 0 And vEstado = "P" Then
    
    If fxTesTiposDocAsiento(cboDoc.ItemData(cboDoc.ListIndex)) = "D" Then
       vGrid.Row = 1
       vGrid.col = 7
'       If cTipoCambio > 1 Then
'         vGrid.Text = CStr(CCur(txtMonto * cTipoCambio))
'       Else
'         vGrid.Text = CStr(CCur(txtMonto))
'       End If
       vGrid.Text = CStr(CCur(txtMonto * fxSys_Tipo_Cambio_Apl(cTipoCambio)))
       
       vGrid.col = 8
       vGrid.Text = "0"
    Else
       vGrid.Row = 1
       vGrid.col = 7
       vGrid.Text = "0"
       vGrid.col = 8
'       If cTipoCambio > 1 Then
'         vGrid.Text = CStr(CCur(txtMonto * cTipoCambio))
'       Else
'         vGrid.Text = CStr(CCur(txtMonto))
'       End If
    
       vGrid.Text = CStr(CCur(txtMonto * fxSys_Tipo_Cambio_Apl(cTipoCambio)))
    
    End If

End If


Call sbSumaDebitosCreditos
'Call sbCalculaMontoDivisaForanea(0)

'Bloquea la Primer Linea
vGrid.Row = 1
For i = 1 To vGrid.MaxCols
    vGrid.col = i
    vGrid.Lock = True
    vGrid.Protect = True
Next i

Me.MousePointer = vbDefault
Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub SbSeguimiento()
Dim strSQL As String, rs As New ADODB.Recordset


If vCodigo = 0 Then Exit Sub

Me.MousePointer = vbHourglass


        
With vGridDetalle
 Select Case .ActiveSheet
    'Case 1 'Acciones
    Case 1 'Bitacora
        strSQL = "select H.ID, H.FECHA, H.USUARIO,ISNULL(M.DESCRIPCION,'No identificado'),H.DETALLE" _
               & " from TES_HISTORIAL H left join TES_TIPOS_MOVIMIENTOS M on H.COD_MOVIMIENTO = M.COD_MOVIMIENTO" _
               & " WHERE H.NSOLICITUD = " & vCodigo
        Call sbCargaGridFps7(vGridDetalle, 5, strSQL, True, .ActiveSheet)
        
'        strSQL = "select USER_ASIENTO_EMISION,USER_ASIENTO_ANULA,USER_ANULA,USER_GENERA,USER_SOLICITA" _
'               & ",USER_ENTREGA,FECHA_ENTREGA,FECHA_ASIENTO,FECHA_ASIENTO2,FECHA_SOLICITUD,FECHA_EMISION" _
'               & ",FECHA_ANULA,FECHA_AUTORIZACION,USER_AUTORIZA" _
'               & " from Tes_Transacciones where Nsolicitud = " & vCodigo
'        Call OpenRecordSet(rs, strSQL)
'
'        .Sheet = 1
'        .Row = 1
'        .Col = 2
'        .Text = Trim(rs!user_solicita) & ""
'        .Col = 3
'        .Text = rs!fecha_solicitud & ""
'
'        .Row = 2
'        .Col = 2
'        .Text = Trim(rs!USER_AUTORIZA) & ""
'        .Col = 3
'        .Text = rs!fecha_autorizacion & ""
'
'        .Row = 3
'        .Col = 2
'        .Text = Trim(rs!user_genera) & ""
'        .Col = 3
'        .Text = rs!fecha_emision & ""
'
'        .Row = 4
'        .Col = 2
'        .Text = Trim(rs!user_anula) & ""
'        .Col = 3
'        .Text = rs!fecha_anula & ""
'
'        .Row = 5
'        .Col = 2
'        .Text = Trim(rs!user_entrega) & ""
'        .Col = 3
'        .Text = rs!FECHA_Entrega & ""
'
'        .Row = 6
'        .Col = 2
'        .Text = Trim(rs!user_asiento_emision) & ""
'        .Col = 3
'        .Text = rs!fecha_asiento & ""
'
'
'        .Row = 7
'        .Col = 2
'        .Text = Trim(rs!user_asiento_Anula) & ""
'        .Col = 3
'        .Text = rs!fecha_asiento2 & ""
'       rs.Close
   
   Case 2 'Localizacion
        strSQL = "select D.fecha_rec,D.cod_remesa,U.descripcion,D.usuario_rec,D.observacion" _
               & " from Tes_Ubi_RemDet D inner join Tes_ubi_Remesa R on D.cod_Remesa = R.cod_remesa" _
               & " inner join tes_Ubicaciones U on R.cod_ubicacion_destino = U.cod_ubicacion" _
               & " Where D.nsolicitud = " & vCodigo & " And D.estado = 1" _
               & " Order by D.fecha_rec desc"
        Call sbCargaGridFps7(vGridDetalle, 5, strSQL, True, .ActiveSheet)
       
   Case 3 'ReImpresiones
        strSQL = "select Fecha,Usuario,Autoriza,Notas from Tes_reImpresiones where nsolicitud = " & vCodigo _
               & " order by fecha desc"
        Call sbCargaGridFps7(vGridDetalle, 4, strSQL, True, .ActiveSheet)
        
   Case 4 'Cambios de Fecha
        strSQL = "select Id as Idx,Fecha,Usuario,Detalle from tes_historial where nsolicitud = " & vCodigo _
               & " and cod_movimiento = '08' order by fecha desc"
               
        Call sbCargaGridFps7(vGridDetalle, 4, strSQL, True, .ActiveSheet)
        
   
        
 End Select
End With
    
    

Me.MousePointer = vbDefault
End Sub



Private Sub SSTab_Click(PreviousTab As Integer)

Select Case ssTab.SelectedItem
  Case 1 'Asiento
    Call sbCargaAsiento
  Case 2 'Seguimiento
    Call SbSeguimiento
End Select

End Sub

Private Sub tlb_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim strSQL As String

Select Case UCase(Button.Key)
    Case "INSERTAR", "NUEVO"
      vEdita = False
      Call sbLimpiaPantalla
      txtCodigo.Text = ""
      txtCodigo.Enabled = False
      cbo.SetFocus
      ssTab.Item(1).Enabled = True
      Call sbToolBar(tlb, "edicion")
      
      If gBanco <> "" Then
         On Error Resume Next
         cbo.Text = gBanco
         cboConcepto.Text = gConcepto
         cboUnidad.Text = gUnidad
         cboDoc.Text = gDocumento
      End If
      
    Case "MODIFICAR", "EDITAR"
      vEdita = True
      txtCodigo.Enabled = False
      cbo.SetFocus
      ssTab.Item(1).Enabled = True
      Call sbToolBar(tlb, "edicion")
    Case "BORRAR"
      Call sbBorrar
    
    Case "GUARDAR", "SALVAR"
     If fxValida Then Call sbGuardar
     
    Case "DESHACER"
      Call sbToolBar(tlb, "activo")
      If vCodigo = 0 Then
        Call sbLimpiaPantalla
        Call sbToolBar(tlb, "nuevo")
        vEdita = False
        txtCodigo.Enabled = True
        txtCodigo.SetFocus
      Else
        Call sbConsulta(vCodigo)
      End If
      
    Case "CONSULTAR"
    
    Case "REPORTES"
      Call sbReportes
    
    Case "AYUDA"
        frmContenedor.CD.HelpContext = Me.HelpContextID
        frmContenedor.CD.ShowHelp
   
End Select

End Sub


Private Sub sbReportes()

If vCodigo = 0 Then Exit Sub

Me.MousePointer = vbHourglass

With frmContenedor.Crt
    .Reset
    .WindowShowRefreshBtn = True
    .WindowShowPrintSetupBtn = True
    .WindowState = crptMaximized
    .WindowShowSearchBtn = True
    .WindowTitle = "Reportes Módulo de Banking"
    
    .Connect = glogon.ConectRPT
    .Formulas(1) = "fxCodigoBarras = '*" & vCodigo & "*'"
    .ReportFileName = SIFGlobal.fxPathReportes("Banking_BoletaRegistro.rpt")
    .SelectionFormula = "{CHEQUES.NSOLICITUD} = " & vCodigo
    
    .SubreportToChange = "sbDetalle"

    .StoredProcParam(0) = vCodigo

    .PrintReport
End With

Me.MousePointer = vbDefault

End Sub



Private Sub sbSetCombos(cboX As ComboBox, vTexto As String)
Dim vPasoX As Boolean, i As Integer
    
vPasoX = False
For i = 0 To cboX.ListCount
  If Trim(cboX.List(i)) = Trim(vTexto) Then
     vPasoX = True
  End If
Next i

If Not vPasoX Then
   cboX.AddItem vTexto
End If
cboX.Text = vTexto

End Sub

Public Sub sbTESDocConsulta(xCodigo As Long)
 Call sbConsulta(xCodigo)
End Sub

Private Sub sbConsulta(xCodigo As Long)
Dim rs As New ADODB.Recordset, strSQL As String
Dim i As Integer, vPasoX As Boolean

On Error GoTo vError



Me.MousePointer = vbHourglass

strSQL = "select C.*,B.descripcion as Banco,rtrim(X.descripcion) as ConceptoX" _
       & ",rtrim(U.descripcion) as UnidadX" _
       & ",rtrim(T.descripcion) as DocumentoX" _
       & ",B.COD_DIVISA as 'Divisa_Id', dbo.fxSys_Cadena_Capitaliza(di.DESCRIPCION) as 'DIVISA_DESC', isnull(Di.CURRENCY_SIM,B.COD_DIVISA) as 'CURRENCY_SIM'" _
       & " from Tes_Bancos B inner join Tes_Transacciones C on B.id_Banco = C.id_Banco" _
       & " inner join Tes_Conceptos X on C.cod_concepto = X.cod_Concepto" _
       & " inner join Tes_Tipos_doc T on C.tipo = T.tipo" _
       & " inner join CntX_Unidades U on C.cod_unidad = U.cod_unidad and U.cod_contabilidad = " & GLOBALES.gEnlace _
       & "  left join vSys_Divisas Di on B.COD_DIVISA = Di.COD_DIVISA" _
       & " where C.nsolicitud = " & xCodigo
Call OpenRecordSet(rs, strSQL)

If Not rs.BOF And Not rs.EOF Then
  Call sbLimpiaPantalla
  Call sbToolBar(tlb, "activo")
  vEdita = True
  
  vCodigo = rs!NSolicitud
  txtCodigo = rs!NSolicitud
  txtCodigo.Enabled = True
  
  txtCodBene = rs!Codigo & ""
  txtBeneficiario.Text = Trim(rs!Beneficiario & "")
  
  txtMonto = Format(rs!Monto, "Standard")
  lblMontoLetra.Caption = fxMontoLetrasX(rs!Monto, rs!Divisa_Desc)
  
  txtRef_01.Text = Trim(rs!ref_01 & "")
  txtRef_02.Text = Trim(rs!Ref_02 & "")
  txtRef_03.Text = Trim(rs!Ref_03 & "")
  
  txtDocumento = rs!nDocumento & ""
  
  imgEstado.Visible = True
  imgAutorizado.Visible = True
  imgBloqueado.Visible = True
    
  Select Case rs!Estado
    Case "S", "P"
       vEstado = rs!Estado
       gFecha = Format(rs!fecha_solicitud, "dd/mm/yyyy")
       imgEstado.ToolTipText = "Solicitado por " & rs!user_solicita & " - " & gFecha
       Set imgEstado.Picture = imgIconosEstados.ListImages.Item(3).Picture
       
    Case "X" 'Transito
       gFecha = Format(rs!fecha_solicitud, "dd/mm/yyyy")
       vEstado = rs!Estado
       imgEstado.ToolTipText = "Transito " & rs!user_solicita & " - " & gFecha
       Set imgEstado.Picture = imgIconosEstados.ListImages.Item(3).Picture

    Case "I", "E", "T" 'Emitido
       vEstado = rs!Estado
       gFecha = Format(rs!Fecha_Emision, "dd/mm/yyyy")
       imgEstado.ToolTipText = "Emitido por " & rs!user_genera & " - " & gFecha
       Set imgEstado.Picture = imgIconosEstados.ListImages.Item(1).Picture
    Case "A" 'Anulado
       vEstado = rs!Estado
       gFecha = Format(rs!Fecha_Anula, "dd/mm/yyyy")
       imgEstado.ToolTipText = "Anulado por " & vbCrLf & rs!user_anula & " - " & gFecha
       Set imgEstado.Picture = imgIconosEstados.ListImages.Item(2).Picture
  End Select
  

   StatusBarX.Panels(1).Text = StatusBarX.Panels(1).Text & " " & IIf(IsNull(rs!fecha_solicitud), " ", Format(rs!fecha_solicitud, "dd/mm/yyyy"))
   StatusBarX.Panels(1).ToolTipText = IIf(IsNull(rs!user_solicita), "", " " & rs!user_solicita)
   StatusBarX.Panels(2).Text = StatusBarX.Panels(2).Text & " " & IIf(IsNull(rs!Fecha_Emision), " ", Format(rs!Fecha_Emision, "dd/mm/yyyy"))
   StatusBarX.Panels(2).ToolTipText = IIf(IsNull(rs!user_genera), "", " " & rs!user_genera)
   StatusBarX.Panels(3).Text = StatusBarX.Panels(3).Text & " " & IIf(IsNull(rs!Fecha_Anula), " ", Format(rs!Fecha_Anula, "dd/mm/yyyy"))
   StatusBarX.Panels(3).ToolTipText = IIf(IsNull(rs!user_anula), "", " " & rs!user_anula)
   StatusBarX.Refresh
  
  'txtEstado.Tag = rs!Estado
  
'  txtdetalle = rs!detalle & ""

   txtDetalle = rs!Detalle1 & " " & rs!Detalle2 & " " & rs!Detalle3 & " " & rs!Detalle4 & " " & rs!Detalle5
   cboCuenta = rs!Cta_Ahorros & ""
   txtRef = rs!Referencia & ""
   txtTipoCambio.Text = Format(rs!TIPO_CAMBIO, "Standard")
   
   txtDivisa.Text = rs!Divisa_Id & ""
   txtDivisa.Tag = rs!CURRENCY_SIM & ""
   txtDivisa.ToolTipText = rs!Divisa_Desc & ""
   
   'En caso de bloqueo
   If Not IsNull(rs!fecha_hold) Then
       imgBloqueo.Visible = True
       imgBloqueado.Visible = True
       imgBloqueo.ToolTipText = Trim(rs!user_hold) & " - " & Format(rs!fecha_hold, "dd/mm/yyyy")
       Set imgBloqueado.Picture = imgIconosEstados.ListImages.Item(6).Picture
       imgBloqueado.ToolTipText = "Bloqueado por: " & Trim(rs!user_hold) & " - " & rs!fecha_hold
   Else
       imgBloqueo.Visible = False
       Set imgBloqueado.Picture = imgIconosEstados.ListImages.Item(7).Picture
       imgBloqueado.ToolTipText = "Sin Bloqueo"
       
   End If
   'en caso de autorizado
   If Not IsNull(rs!fecha_autorizacion) Then
       Set imgAutorizado.Picture = imgIconosEstados.ListImages.Item(4).Picture
       imgAutorizado.ToolTipText = "Autorizado por: " & Trim(rs!user_autoriza) & " - " & Format(rs!fecha_autorizacion, "dd/mm/yyyy")
   Else
       Set imgAutorizado.Picture = imgIconosEstados.ListImages.Item(5).Picture
       imgAutorizado.ToolTipText = "Sin Autorizar"
   End If
   
   
   'Cargar Combos / Si cada Carga da Error/Se debe de Ingresar al Combo para Consulta
    vPaso = True
    Call sbCboAsignaDato(cbo, rs!Banco, True, CStr(rs!Id_Banco))
    vPaso = False
    cbo.Text = rs!Banco 'Esto Activa el Evento Click del Cbo Tes_Bancos
    
    Call sbCboAsignaDato(cboDoc, rs!DocumentoX, True, rs!Tipo)
    Call sbCboAsignaDato(cboConcepto, rs!ConceptoX, True, rs!COD_CONCEPTO)
    Call sbCboAsignaDato(cboUnidad, rs!UnidadX, True, rs!Cod_Unidad)
    
    cboTipos.ListIndex = IIf(IsNull(rs!Tipo_Beneficiario), 0, rs!Tipo_Beneficiario - 1)
    
    vGrid.MaxRows = 0
    
    ssTab.Item(0).Selected = True
    For i = 0 To ssTab.ItemCount - 1
      ssTab.Item(i).Enabled = True
    Next i
    
Else
  MsgBox "No se encontró registro verifique...", vbInformation
  Call sbLimpiaPantalla
End If

rs.Close

Call cboDoc_Click

Me.MousePointer = vbDefault
Call RefrescaTags(Me)

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Function fxValida() As Boolean
Dim vMensaje As String, i As Integer
Dim curMonto As Currency, vTextTemp As String
Dim curTipoCambio As Currency

Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

vMensaje = ""
fxValida = True

'Validar en los documento que se autoemiten que la fecha no corresponda a un periodo cerrado.

If Not fxTesBancoValida(cbo.ItemData(cbo.ListIndex), glogon.Usuario) Then vMensaje = vMensaje & vbCrLf & " - El Usuario Actual no esta Autorizado a utilizar este Banco..."
If Not fxTesTipoAccesoValida(cbo.ItemData(cbo.ListIndex), glogon.Usuario, cboDoc.ItemData(cboDoc.ListIndex)) Then vMensaje = vMensaje & vbCrLf & " - El Usuario Actual no esta Autorizado a utilizar este Tipo de Documento..."
If Not fxTesConceptoValida(cbo.ItemData(cbo.ListIndex), glogon.Usuario, cboConcepto.ItemData(cboConcepto.ListIndex)) Then vMensaje = vMensaje & vbCrLf & " - El Usuario Actual no esta Autorizado a utilizar este Concepto..."
If Not fxTesUnidadValida(cbo.ItemData(cbo.ListIndex), glogon.Usuario, cboUnidad.ItemData(cboUnidad.ListIndex)) Then vMensaje = vMensaje & vbCrLf & " - El Usuario Actual no esta Autorizado a utilizar esta unidad..."

'Si el documento Se AutoEmite / Revisar si Tiene AutoConsecutivo / Si no es Asi validad el # de Documento
If fxTesBancoDocsValor(cbo.ItemData(cbo.ListIndex), cboDoc.ItemData(cboDoc.ListIndex), "REG_EMISION") = 0 Then
  If fxTesBancoDocsValor(cbo.ItemData(cbo.ListIndex), cboDoc.ItemData(cboDoc.ListIndex), "DOC_AUTO") = 0 Then
     If Len(Trim(txtDocumento)) = 0 Then
        vMensaje = vMensaje & vbCrLf & " - Esta Solicitud se AutoEmite / Digite el #Documento para su Emisión..."
     Else
        'Verificar si Existe el # Documento ya Emitido
        If Not fxTesDocumentoVerifica(cbo.ItemData(cbo.ListIndex), cboDoc.ItemData(cboDoc.ListIndex), txtDocumento) Then
            vMensaje = vMensaje & vbCrLf & " - Esta Solicitud se AutoEmite / El #Documento para su Emisión ya se encuentra registrado..."
        End If
        
     End If
  End If
End If


If IsNumeric(txtMonto) Then
   If CCur(txtMonto) <= 0 Then
        vMensaje = vMensaje & vbCrLf & " - El monto del documento no es válido..."
   End If
Else
    vMensaje = vMensaje & vbCrLf & " - El monto del documento no es válido..."
End If

If txtCodBene = "" Then vMensaje = vMensaje & vbCrLf & " - Codigo del Beneficiario no es válido ..."
If Len(Trim(txtBeneficiario)) = 0 Then vMensaje = vMensaje & vbCrLf & " - Beneficiario no es válido ..."
If Len(Trim(txtDetalle)) = 0 Then vMensaje = vMensaje & vbCrLf & " - El Detalle no es válido ..."

If vEstado <> "P" Then vMensaje = vMensaje & vbCrLf & " - No se puede modificar este Documento porque se encuentra Emitido o Anulado ..."
'Verificar o preguntar si el documento esta autorizado para que no se pueda modificar / ***** ojo consulta

'Revisar Asiento, si no tiene Lineas crear Asiento Básico
If vGrid.MaxRows <= 1 Then Call sbCargaAsiento

'despues de la creacion del asiento revisar #lineas y Balance
If vGrid.MaxRows <= 1 Then vMensaje = vMensaje & vbCrLf & " - El Asiento no se válido..."

Call sbSumaDebitosCreditos
If CCur(txtDiferencia) <> 0 Then vMensaje = vMensaje & vbCrLf & " - El Asiento no se encuentra balanceado..."

'Valida que la Primer linea del Asiento sea igual al monto del documento
vGrid.Row = 1
vGrid.col = 7
curMonto = CCur(vGrid.Text)
vGrid.col = 8
curMonto = curMonto + CCur(vGrid.Text)
If CCur(txtTipoCambio.Text) <> 1 Then
curMonto = curMonto / fxSys_Tipo_Cambio_Apl(CCur(txtTipoCambio.Text))
End If
If Abs(curMonto - CCur(txtMonto)) > 5 Then vMensaje = vMensaje & vbCrLf & " - El Monto Linea 1 del Asiento no corresponde al original..."


'Valida Asiento: Cuentas, Unidad, Centros, Divisas y Tipo de Cambios
Dim pCuenta As String, pUnidad As String, pCentro As String, pDivisa As String
Dim pDebito As Currency, pCredito As Currency, pTipoCambio As Currency


strSQL = ""
For i = 1 To vGrid.MaxRows
   vGrid.Row = i
   vGrid.col = 1
   pCuenta = vGrid.Text
   vGrid.col = 2
   pUnidad = vGrid.Text
   vGrid.col = 3
   pCentro = vGrid.Text
   vGrid.col = 4
   pDivisa = vGrid.Text
   vGrid.col = 5
   pTipoCambio = IIf(IsNumeric(vGrid.Text), vGrid.Text, 0)
   vGrid.col = 7
   pDebito = IIf(IsNumeric(vGrid.Text), vGrid.Text, -1)
   vGrid.col = 8
   pCredito = IIf(IsNumeric(vGrid.Text), vGrid.Text, -1)
   
 strSQL = strSQL & Space(10) & "exec spCntX_Cuentas_Valida_Load " & GLOBALES.gEnlace & ",'" & glogon.Usuario _
        & "','TES', '" & pCuenta & "','" & pDivisa & "','" & pUnidad & "','" & pCentro _
        & "'," & pTipoCambio & "," & pDebito & "," & pCredito & "," & IIf((i = 1), 1, 0)
   
 If Len(strSQL) > 20000 Then
    Call ConectionExecute(strSQL)
    strSQL = ""
 End If
 
Next i

'Ultimo Lote
If Len(strSQL) > 0 Then
   Call ConectionExecute(strSQL)
   strSQL = ""
End If

strSQL = "exec spCntX_Cuentas_Valida_Resultado '" & glogon.Usuario & "', 0"
Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
  vMensaje = vMensaje & vbCrLf & " - " & rs!Resultado & ""
  rs.MoveNext
Loop
rs.Close

''''Valida Cuenta, Unidad de Negocios y Centro de Costo
'''For i = 1 To vGrid.MaxRows
'''  vGrid.Row = i
'''
'''  vGrid.col = 5
'''  If Len(vGrid.Text) = 0 Then vMensaje = vMensaje & vbCrLf & " - La cuenta de la linea : " & i & " no es válida..."
'''
'''  vGrid.col = 2
'''  vTextTemp = vGrid.Text
'''  If Not fxTesUnidadValida(cbo.ItemData(cbo.ListIndex), glogon.Usuario, vGrid.Text) Then
'''     vMensaje = vMensaje & vbCrLf & " - La unidad de negocios no es válida en la línea : " & i & " o el usuario no esta autorizado a esta..."
'''  End If
'''
'''
'''
'''  vGrid.col = 3
'''  If Not fxTesUnidadCCValida(vGrid.Text, vTextTemp) Then
'''     vMensaje = vMensaje & vbCrLf & " - El Centro de Costo no es válido en la línea : " & i & " o no se encuentra asignado a la unidad : " & vTextTemp
'''  End If
'''
''''  vGrid.col = 1
''''  curTipoCambio = Format(fxDivisaTipoCambio(fxDivisaCuenta(fxgCntCuentaFormato(False, vGrid.Text, 0))), "Standard")
''''  vGrid.col = 5
''''  If CCur(vGrid.Text) > (curTipoCambio + gVariacionAsiento) Or CCur(vGrid.Text) < (curTipoCambio - gVariacionAsiento) Then
''''     vMensaje = vMensaje & vbCrLf & " - El valor en del tipo de cambio en la línea : " & i & " es superior o inferior a la variacion permitida : " & gVariacionAsiento
''''     vGrid.Text = curTipoCambio
''''  End If
'''
'''
'''
'''Next i

vError:

If Len(vMensaje) > 0 Then
  fxValida = False
  MsgBox vMensaje, vbCritical
End If

End Function

Private Sub sbGuardar()
Dim strSQL As String, rs As New ADODB.Recordset
Dim i As Integer

On Error GoTo vError

If vEdita Then

 strSQL = "update Tes_Transacciones set id_banco = " & cbo.ItemData(cbo.ListIndex) _
        & ",tipo = '" & cboDoc.ItemData(cboDoc.ListIndex) & "',cod_concepto = '" & cboConcepto.ItemData(cboConcepto.ListIndex) _
        & "',cod_unidad = '" & cboUnidad.ItemData(cboUnidad.ListIndex) _
        & "',codigo = '" & txtCodBene & "',Beneficiario = '" & txtBeneficiario & "',monto = " & CCur(txtMonto) & ",cta_ahorros = '" _
        & Trim(cboCuenta.Text) & "',detalle1 = '" & Mid(txtDetalle, 1, 27) & "',detalle2 = '" & Mid(txtDetalle, 28, 27) _
        & "',detalle3 = '" & Mid(txtDetalle, 55, 27) & "',detalle4 = '" & Mid(txtDetalle, 82, 27) & "',detalle5 = '" _
        & Mid(txtDetalle, 109, 27) & "', tipo_beneficiario = " & cboTipos.ItemData(cboTipos.ListIndex) & ",  tipo_cambio = " _
        & CCur(txtTipoCambio.Text) & ",cod_divisa = '" & gDivisa & "',referencia = "
        
   If txtRef = "" Or Not IsNumeric(txtRef) Then
     strSQL = strSQL & "Null"
   Else
     strSQL = strSQL & txtRef
   End If
        
   If fxTesBancoDocsValor(cbo.ItemData(cbo.ListIndex), cboDoc.ItemData(cboDoc.ListIndex), "REG_AUTORIZACION") = 1 Then
      strSQL = strSQL & ",Autoriza = 'N', fecha_autorizacion = null,user_autoriza = null"
   Else
      strSQL = strSQL & ",Autoriza = 'S',fecha_autorizacion = dbo.MyGetdate(), user_autoriza = '" & glogon.Usuario & "'"
   End If
        
  strSQL = strSQL & " where nsolicitud = " & vCodigo
 
  Call ConectionExecute(strSQL)
  Call Bitacora("Modifica", "0" & vCodigo)

Else
  
  strSQL = "insert Tes_Transacciones(id_banco,tipo,cod_concepto,cod_unidad,codigo,beneficiario,monto,estado,fecha_solicitud,user_solicita" _
         & ",estadoi,modulo,subModulo,cta_ahorros,genera,actualiza,detalle1,detalle2,detalle3,detalle4,detalle5,referencia" _
         & ",op,estado_asiento,entregado,autoriza,fecha_autorizacion,user_autoriza,Ndocumento,tipo_cambio,cod_divisa,tipo_beneficiario, cod_App) values(" _
         & cbo.ItemData(cbo.ListIndex) & ",'" & cboDoc.ItemData(cboDoc.ListIndex) & "','" & cboConcepto.ItemData(cboConcepto.ListIndex) _
         & "','" & cboUnidad.ItemData(cboUnidad.ListIndex) _
         & "','" & txtCodBene & "','" & UCase(txtBeneficiario) & "'," & CCur(txtMonto.Text) & ",'P',dbo.MyGetdate(),'" & glogon.Usuario _
         & "','P','" & vModulo & "','T','" & cboCuenta.Text & "','S','N','"
 
   strSQL = strSQL & Mid(txtDetalle, 1, 27) & "','" & Mid(txtDetalle, 28, 27) & "','" & Mid(txtDetalle, 55, 27) & "','"
   strSQL = strSQL & Mid(txtDetalle, 82, 27) & "','" & Mid(txtDetalle, 109, 27) & "',"
   
   If txtRef = "" Or Not IsNumeric(txtRef) Then
     strSQL = strSQL & "Null,Null"
   Else
     strSQL = strSQL & txtRef & ",Null"
   End If
   
   strSQL = strSQL & ",'P','N',"
   If fxTesBancoDocsValor(cbo.ItemData(cbo.ListIndex), cboDoc.ItemData(cboDoc.ListIndex), "REG_AUTORIZACION") = 1 Then
      strSQL = strSQL & "'N',null,null,'" & txtDocumento & "'," & CCur(txtTipoCambio.Text) _
             & ",'" & gDivisa & "'," & cboTipos.ItemData(cboTipos.ListIndex) & ", 'ProGrX')"
   Else
      strSQL = strSQL & "'S',dbo.MyGetdate(),'" & glogon.Usuario & "','" & txtDocumento _
            & "'," & CCur(txtTipoCambio.Text) & ",'" & gDivisa & "'," & cboTipos.ItemData(cboTipos.ListIndex) _
            & ", 'ProGrX')"
   End If
   
   Call ConectionExecute(strSQL)

  
   strSQL = "select isnull(max(nsolicitud),0) as IDx from Tes_Transacciones where codigo = '" & txtCodBene & "'"
   Call OpenRecordSet(rs, strSQL)
     vCodigo = rs!IdX
   rs.Close
    
   txtCodigo = vCodigo
    
   Call Bitacora("Registra", "Solicitud : " & vCodigo)

End If

'Borra Detalle (Inicia Bloque)
strSQL = "delete Tes_Trans_Asiento where nsolicitud = " & vCodigo
'    Call ConectionExecute(strSQL)

'Guarda Detalle
For i = 1 To vGrid.MaxRows
  vGrid.Row = i
  
  vGrid.col = 4
  If vGrid.Text <> "" Then
    vGrid.col = 1
    strSQL = strSQL & Space(10) & "insert Tes_Trans_Asiento(nSolicitud,Linea,Cuenta_Contable,cod_unidad,cod_cc,cod_divisa,tipo_cambio,DebeHaber,Monto) values(" & vCodigo _
           & "," & i & ",'" & fxgCntCuentaFormato(False, vGrid.Text, 0) & "','"
    vGrid.col = 2
    strSQL = strSQL & vGrid.Text & "','"
    vGrid.col = 3
    strSQL = strSQL & vGrid.Text & "','"
    vGrid.col = 4
    strSQL = strSQL & vGrid.Text & "',"
    vGrid.col = 5
    strSQL = strSQL & vGrid.Text & ",'"
    
    vGrid.col = 7
    If CCur(vGrid.Text) > 0 Then
      strSQL = strSQL & "D'," & CCur(vGrid.Text) & ")"
    Else
      vGrid.col = 8
      strSQL = strSQL & "H'," & CCur(vGrid.Text) & ")"
    End If
'    Call ConectionExecute(strSQL)
  
  End If

Next i

'Registra el Asiento (Una sola Instucción)
Call ConectionExecute(strSQL)



'Preguntar Si Tiene AutoEmision / Generar Documento
If fxTesBancoDocsValor(cbo.ItemData(cbo.ListIndex), cboDoc.ItemData(cboDoc.ListIndex), "REG_EMISION") = 0 Then
   'EMITE DOCUMENTO AQUI
   Call sbCargaArchivosEspeciales(cbo.ItemData(cbo.ListIndex))
   Call sbTesEmitirDocumento(vCodigo, txtDocumento.Text, dtpEmision.Value)
End If

MsgBox "Información guardada satisfactoriamente...", vbInformation

 gBanco = cbo.Text
 gConcepto = cboConcepto.Text
 gUnidad = cboUnidad.Text
 gDocumento = cboDoc.Text


txtCodigo.Enabled = True
txtCodigo.SetFocus
Call txtCodigo_LostFocus

Call sbToolBar(tlb, "activo")

Call RefrescaTags(Me)

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub sbBorrar()
Dim i As Integer, strSQL As String

On Error GoTo vError

If vCodigo = 0 Then Exit Sub

If vEstado <> "P" Then
  MsgBox "Esta solicitud no puede ser eliminada del sistema por que no se encuentra Solicitada...", vbExclamation
  Exit Sub
End If

i = MsgBox("Esta Seguro que desea borrar este registro", vbYesNo)

If i = vbYes Then
  strSQL = "delete Tes_Trans_Asiento where nsolicitud = " & vCodigo
  Call ConectionExecute(strSQL)

  strSQL = "delete Tes_Transacciones where nsolicitud = " & vCodigo
  Call ConectionExecute(strSQL)

  Call Bitacora("Elimina", "Solicitud # : " & vCodigo)
  Call sbLimpiaPantalla
  Call sbToolBar(tlb, "nuevo")
  Call RefrescaTags(Me)
End If

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub tlb_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
'Select Case ButtonMenu.Key
'  Case "LisTipo de Documentos"
'     Call sbReportesInv("Tipo de Documentos", "Tipo de DocumentoS", "Listado", "")
'  Case "InvTipo de Documentos"
'     Call sbReportesInv("InvTipo de Documentos", "Tipo de DocumentoS", "Inventario", "")
'End Select

End Sub

Private Sub txtBeneficiario_KeyDown(KeyCode As Integer, Shift As Integer)

On Error GoTo vError


If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtMonto.SetFocus

If KeyCode = vbKeyF2 Then txtBeneficiario.Text = cbo.Text

If KeyCode = vbKeyF4 Then
  gBusquedas.Convertir = "N"
  Select Case cboTipos.ItemData(cboTipos.ListIndex)
      Case 1
        gBusquedas.Columna = "cedula"
        gBusquedas.Orden = "cedula"
        gBusquedas.Consulta = "select cedula,nombre from socios"
      Case 2
        gBusquedas.Columna = "id_banco"
        gBusquedas.Orden = "id_banco"
        gBusquedas.Consulta = "select id_banco,descripcion from tes_bancos"
      Case 3
        gBusquedas.Columna = "cod_proveedor"
        gBusquedas.Orden = "cod_proveedor"
        gBusquedas.Consulta = "select cod_proveedor,descripcion from cxp_proveedores"

     Case 4
        gBusquedas.Columna = "cod_acreedor"
        gBusquedas.Orden = "cod_acreedor"
        gBusquedas.Consulta = "select cod_acreedor,descripcion from crd_apa_acreedores"
  End Select
  
  gBusquedas.Filtro = ""
  frmBusquedas.Show vbModal
  txtCodBene = gBusquedas.Resultado
  txtBeneficiario = gBusquedas.Resultado2
  If txtCodBene <> "" Then
    txtMonto.SetFocus
  Else
    txtCodBene.SetFocus
  End If
    
End If
gBusquedas.Resultado = ""
gBusquedas.Resultado2 = ""
Exit Sub

vError:
   MsgBox fxSys_Error_Handler(Err.Description)
End Sub

Private Sub txtCodBene_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo vError

If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtBeneficiario.SetFocus


If KeyCode = vbKeyF4 Then
  gBusquedas.Convertir = "N"
  Select Case cboTipos.ItemData(cboTipos.ListIndex)
      Case 1 'Personas
        gBusquedas.Col1Name = "Cedula"
        gBusquedas.Col2Name = "Nombre"
        gBusquedas.Col3Name = "Id Alterno"
        gBusquedas.Columna = "cedula"
        gBusquedas.Orden = "cedula"
        gBusquedas.Consulta = "select cedula,nombre, cedular from socios"
      Case 2 'Bancos
        gBusquedas.Columna = "id_banco"
        gBusquedas.Orden = "id_banco"
        gBusquedas.Consulta = "select id_banco,descripcion from tes_bancos"
      Case 3 'PRoveedores
        gBusquedas.Columna = "cedjur"
        gBusquedas.Orden = "cedjur"
        gBusquedas.Consulta = "select cedjur,cod_proveedor,descripcion from cxp_proveedores"
     Case 4 'Acreedores
        gBusquedas.Columna = "cod_acreedor"
        gBusquedas.Orden = "cod_acreedor"
        gBusquedas.Consulta = "select cod_acreedor,descripcion from crd_apa_acreedores"
  
     Case 5 'Cuentas por Cobrar
        gBusquedas.Columna = "cedula"
        gBusquedas.Orden = "cedula"
        gBusquedas.Consulta = "select cedula,nombre from CXC_PERSONAS"
     
     Case 6 'Empleados
        gBusquedas.Columna = "Identificacion"
        gBusquedas.Orden = "Identificacion"
        gBusquedas.Consulta = "select Identificacion,Nombre_Completo from RH_PERSONAS"
     
     Case 7 'Directos
        gBusquedas.Columna = "Codigo"
        gBusquedas.Orden = "Codigo"
        gBusquedas.Consulta = "select Codigo,Beneficiario from vTes_Beneficiarios"
     
  End Select
  
  gBusquedas.Filtro = ""
  frmBusquedas.Show vbModal
  txtCodBene = gBusquedas.Resultado
  If txtCodBene <> "" Then txtBeneficiario.SetFocus
End If
gBusquedas.Resultado = ""
gBusquedas.Resultado2 = ""

Exit Sub


vError:
  MsgBox fxSys_Error_Handler(Err.Description)
End Sub

Private Sub txtCodBene_LostFocus()
Dim strSQL As String, rs As New ADODB.Recordset
Dim pId As String

On Error GoTo vError

Me.MousePointer = vbHourglass


pId = ".-x-."

Select Case cboTipos.ItemData(cboTipos.ListIndex)
   Case 1 'Personas
    If Trim(txtCodBene) <> "" Then
       strSQL = "Select Cedula,Nombre" _
              & " from Socios Where Cedula='" & txtCodBene & "'"
       Call OpenRecordSet(rs, strSQL)
       If Not rs.EOF And Not rs.BOF Then
          txtBeneficiario = Trim(rs!Nombre & "")
          pId = rs!Cedula
       Else
          txtBeneficiario.Text = ""
          cboCuenta.Text = ""
       End If
       rs.Close
    Else
       txtBeneficiario = ""
       cboCuenta.Text = ""
       
    End If

Case 2 'Bancos
    If Trim(txtCodBene) <> "" Then
       strSQL = "select ID_BANCO,descripcion from TES_BANCOS where ID_BANCO  =" & txtCodBene & ""
       Call OpenRecordSet(rs, strSQL)
       If Not rs.EOF And Not rs.BOF Then
          txtBeneficiario = Trim(rs!DESCRIPCION)
          pId = rs!Id_Banco
       Else
          txtBeneficiario = ""
       End If
       rs.Close
    Else
       txtBeneficiario = ""
    End If

Case 3 'Proveedores
    If Trim(txtCodBene) <> "" Then
       strSQL = "select CEDJUR, DESCRIPCION" _
              & " from CXP_PROVEEDORES where CEDJUR = '" & txtCodBene & "'"
       Call OpenRecordSet(rs, strSQL)
       
       If Not rs.EOF And Not rs.BOF Then
          txtBeneficiario = Trim(rs!DESCRIPCION & "")
          pId = rs!CEDJUR
       Else
          txtBeneficiario = ""
       End If
       rs.Close
    Else
       txtBeneficiario = ""
    End If

Case 5 'Cuentas por Cobrar
    If Trim(txtCodBene) <> "" Then
       strSQL = "select Cod_Acreedor, DESCRIPCION  from CRD_APA_ACREEDORES where cod_acreedor = " & txtCodBene & ""
       Call OpenRecordSet(rs, strSQL)
       
       If Not rs.EOF And Not rs.BOF Then
          txtBeneficiario = Trim(rs!DESCRIPCION)
          pId = rs!Cod_Acreedor
       Else
          txtBeneficiario = ""
       End If
       rs.Close
    Else
       txtBeneficiario = ""
    End If


Case 6 'Empleados
    If Trim(txtCodBene) <> "" Then
       strSQL = "Select IDENTIFICACION, NOMBRE_COMPLETO" _
              & " from RH_PERSONAS Where IDENTIFICACION='" & txtCodBene & "'"
              
              
       Call OpenRecordSet(rs, strSQL)
       If Not rs.EOF And Not rs.BOF Then
          txtBeneficiario = Trim(rs!NOMBRE_COMPLETO & "")
          pId = rs!Identificacion
       Else
          txtBeneficiario.Text = ""
       End If
       rs.Close
    Else
       txtBeneficiario = ""
       
    End If


Case 7 'Directos
    If Trim(txtCodBene) <> "" Then
       strSQL = "Select CODIGO, BENEFICIARIO" _
              & " from vTes_Beneficiarios Where CODIGO ='" & txtCodBene & "'"
              
              
       Call OpenRecordSet(rs, strSQL)
       If Not rs.EOF And Not rs.BOF Then
          txtBeneficiario = Trim(rs!Beneficiario & "")
          pId = rs!Codigo
       End If
       rs.Close
    Else
       txtBeneficiario = ""
       
    End If



End Select

If pId = ".-x-." Then
   pId = Trim(txtCodBene.Text)
End If

'(@Identificacion varchar(30), @BancoId int, @DivisaCheck smallint = 0)"
strSQL = "exec spSys_Cuentas_Bancarias '" & pId & "'," & cbo.ItemData(cbo.ListIndex) & ",1"
Call OpenRecordSet(rs, strSQL)

cboCuenta.Clear
Do While Not rs.EOF
  cboCuenta.AddItem rs!IdX
  rs.MoveNext
Loop
rs.Close

Me.MousePointer = vbDefault

Exit Sub
vError:
   Me.MousePointer = vbDefault
   MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub txtCodigo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then cbo.SetFocus
End Sub

Private Sub txtCodigo_LostFocus()
On Error GoTo vError
 If txtCodigo <> 0 And vEdita Then Call sbConsulta(txtCodigo)
vError:
End Sub

Private Sub cboCuenta_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtRef.SetFocus
End Sub

Private Sub txtDetalle_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo vError
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then ssTab.Tab = 1
vError:
End Sub

Private Sub txtDocumento_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then
 ssTab.Item(0).Selected = True
 cboUnidad.SetFocus
End If
End Sub

Private Sub txtMonto_GotFocus()
'On Error GoTo vError
' txtMonto = CCur(txtMonto)
' txtMonto.SelStart = Len(txtMonto)
'Exit Sub
'
'vError:
' txtMonto = 0

End Sub

Private Sub txtMonto_KeyDown(KeyCode As Integer, Shift As Integer)
If (KeyCode = vbKeyReturn Or KeyCode = vbKeyTab) And Val(txtTipoCambio) > 1 Then
   txtTipoCambio.SetFocus
ElseIf KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then
cboCuenta.SetFocus
End If
End Sub

Private Sub txtMonto_LostFocus()
Dim vCadena As String, i As Integer

On Error GoTo vError

'Limpia Monto
vCadena = ""
For i = 1 To Len(txtMonto)
  Select Case Mid(txtMonto, i, 1)
    Case 1, 2, 3, 4, 5, 6, 7, 8, 9, 0
      vCadena = vCadena & Mid(txtMonto, i, 1)
    Case ".", ","
      vCadena = vCadena & Mid(txtMonto, i, 1)
    Case Else
  End Select
Next i

 txtMonto = Format(CCur(vCadena), "Standard")
 lblMontoLetra.Caption = fxMontoLetrasX(CCur(vCadena), gDivisa)
Exit Sub

vError:
 txtMonto = 0
End Sub

Private Sub txtRef_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtDetalle.SetFocus
End Sub

Private Sub vGridDetalle_SheetChanged(ByVal OldSheet As Integer, ByVal NewSheet As Integer)
Call SbSeguimiento
End Sub


Private Sub sbTesControlDivisas(iCodigo As Integer, iContabilidad As Integer)
Dim strSQL As String, rs As New ADODB.Recordset
Dim pDivisaLocal As Integer

strSQL = "SELECT isnull(D.TC_COMPRA,1) as TC_COMPRA , isnull(D.VARIACION,0) as VARIACION, B.COD_DIVISA, Di.DIVISA_LOCAL" _
        & ", dbo.fxSys_Cadena_Capitaliza(di.DESCRIPCION) as 'DIVISA_DESC', isnull(Di.CURRENCY_SIM,B.COD_DIVISA) as 'CURRENCY_SIM'" _
        & " FROM   TES_BANCOS B left JOIN CNTX_DIVISAS_TIPO_CAMBIO D ON B.COD_DIVISA = D.COD_DIVISA" _
        & " AND D.COD_CONTABILIDAD = " & iContabilidad & " and dbo.MyGetdate() between inicio and corte " _
        & " inner join CNTX_DIVISAS Di on B.COD_DIVISA = Di.COD_DIVISA" _
        & " where B.ID_BANCO = " & iCodigo & ""
Call OpenRecordSet(rs, strSQL)

If Not rs.EOF Then
  
    gTipoCambio = rs!tc_compra
    gVariacion = rs!variacion
    gDivisaDesc = rs!Divisa_Desc & ""
    gDivisa = rs!COD_DIVISA
    gDivisaCurrency = rs!CURRENCY_SIM
    
    pDivisaLocal = rs!divisa_local
    rs.Close
    
    If pDivisaLocal = 0 And gTipoCambio = 1 Then
       strSQL = "SELECT Top 1 D.TC_COMPRA, D.VARIACION,X.Descripcion  from CNTX_DIVISAS_TIPO_CAMBIO D inner join  " _
               & " CNTX_DIVISAS X on D.COD_DIVISA = X.COD_DIVISA where  D.COD_CONTABILIDAD = " & iContabilidad & " " _
                & " and D.cod_divisa = '" & gDivisa & "' order by corte desc"
       Call OpenRecordSet(rs, strSQL)
       If Not rs.EOF Or Not rs.BOF Then
          gTipoCambio = rs!tc_compra
          gVariacion = rs!variacion
       End If
    End If
End If

txtTipoCambio.Text = Format(gTipoCambio, "Standard")
If Val(txtTipoCambio) = 1 Then
    txtTipoCambio.Locked = True
Else
    txtTipoCambio.Locked = False
End If

txtDivisa.Text = gDivisa
txtDivisa.ToolTipText = gDivisaDesc & ""
txtDivisa.Tag = gDivisaCurrency


End Sub



Private Function fxDivisaCuenta(vCuenta As String) As String
Dim strSQL As String, rs As New ADODB.Recordset

strSQL = "Select cod_divisa from CNTX_CUENTAS where cod_cuenta = '" & vCuenta & "' and cod_contabilidad  = " & GLOBALES.gEnlace & ""
Call OpenRecordSet(rs, strSQL)

If Not rs.EOF Then
   fxDivisaCuenta = rs!COD_DIVISA
Else
  fxDivisaCuenta = "COL"
End If
rs.Close

End Function

Private Function fxDivisaTipoCambio(pDivisa As String, Optional pTipo As String = "C") As Currency
Dim strSQL As String, rs As New ADODB.Recordset

Dim pFecha As Date


gVariacionAsiento = 0

pFecha = dtpEmision.Value

fxDivisaTipoCambio = fxCntX_TipoCambio(pDivisa, "V", pFecha)

'strSQL = "Select case when '" & pTipo & "' = 'V' then tc_venta else tc_compra end as 'TipoCambio',variacion" _
'       & "  from CNTX_DIVISAS_TIPO_CAMBIO where cod_divisa = '" & pDivisa & "' and cod_contabilidad  = " & GLOBALES.gEnlace & "" _
'       & "  and dbo.MyGetdate() between inicio and corte"
'
'Call OpenRecordSet(rs, strSQL)
'
'If Not rs.EOF Then
'   fxDivisaTipoCambio = rs!TipoCambio
'   gVariacionAsiento = rs!variacion
'Else
'    If pDivisa = "DOL" Or pDivisa = "EUR" And IsNull(rs!TipoCambio) Then
'     strSQL = "SELECT case when '" & pTipo & "' = 'V' then D.tc_venta else D.tc_compra end as 'TipoCambio', D.VARIACION,X.Descripcion" _
'            & " from CNTX_DIVISAS_TIPO_CAMBIO D inner join  " _
'            & " CNTX_DIVISAS X on D.COD_DIVISA = X.COD_DIVISA where  D.COD_CONTABILIDAD = " & GLOBALES.gEnlace & " " _
'            & " and D.cod_divisa = '" & pDivisa & "' order by D.corte desc"
'
'      rs.Close
'      Call OpenRecordSet(rs, strSQL)
'        If Not rs.EOF Or Not rs.BOF Then
'           fxDivisaTipoCambio = rs!TipoCambio
'           gVariacionAsiento = rs!variacion
'        Else
'          fxDivisaTipoCambio = 1
'          gVariacionAsiento = 0
'        End If
'    Else
'         fxDivisaTipoCambio = 1
'         gVariacionAsiento = 0
'    End If
'End If
'
'rs.Close

End Function




Private Sub sbCalculaMontoDivisaForanea(pRow As Long)
Dim x As Integer, TC As Currency
Dim pMontoForaneo

On Error GoTo vError
 pMontoForaneo = 0
  
'  For X = 1 To vGrid.MaxRows
'     vGrid.Row = X
     
     vGrid.Row = pRow
     vGrid.col = 7
     pMontoForaneo = vGrid.Text
     If pMontoForaneo > 0 Then
        vGrid.col = 5
        pMontoForaneo = pMontoForaneo / fxSys_Tipo_Cambio_Apl(CCur(vGrid.Text))
        vGrid.CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
        vGrid.CellNote = "Importe " & Format(pMontoForaneo, "Standard")
        vGrid.TextTip = TextTipFixed
     End If
     vGrid.col = 8
     pMontoForaneo = vGrid.Text

     If pMontoForaneo > 0 Then
        vGrid.col = 5
        pMontoForaneo = pMontoForaneo / fxSys_Tipo_Cambio_Apl(CCur(vGrid.Text))
        vGrid.CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
        vGrid.CellNote = "Importe " & Format(pMontoForaneo, "Standard")
       vGrid.TextTip = TextTipFixed
     End If
    pMontoForaneo = 0
'  Next X
  
vError:

End Sub
