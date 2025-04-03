VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.Controls.v24.0.0.ocx"
Begin VB.Form frmFNDLiquidacionPlan 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Liquidaciones de Planes"
   ClientHeight    =   8805
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12075
   Icon            =   "frmFNDLiquidacionPlan.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8805
   ScaleWidth      =   12075
   WindowState     =   2  'Maximized
   Begin XtremeSuiteControls.CheckBox chkTarjetaActiva_Valida 
      Height          =   255
      Left            =   4680
      TabIndex        =   67
      Top             =   1680
      Width           =   2775
      _Version        =   1572864
      _ExtentX        =   4895
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Valida Tarjeta Activa ?"
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
   Begin XtremeSuiteControls.GroupBox gbResumen 
      Height          =   1695
      Left            =   0
      TabIndex        =   33
      Top             =   6960
      Width           =   11655
      _Version        =   1572864
      _ExtentX        =   20553
      _ExtentY        =   2984
      _StockProps     =   79
      Caption         =   "Resumen"
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
      BorderStyle     =   1
      Begin XtremeSuiteControls.ComboBox cboTipo 
         Height          =   312
         Left            =   7680
         TabIndex        =   40
         Top             =   480
         Width           =   2412
         _Version        =   1572864
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
      Begin XtremeSuiteControls.DateTimePicker dtpVence 
         Height          =   312
         Left            =   8760
         TabIndex        =   53
         Top             =   840
         Width           =   1332
         _Version        =   1572864
         _ExtentX        =   2350
         _ExtentY        =   550
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
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   3
      End
      Begin XtremeSuiteControls.FlatEdit txtCasos 
         Height          =   312
         Left            =   1920
         TabIndex        =   56
         Top             =   480
         Width           =   1572
         _Version        =   1572864
         _ExtentX        =   2773
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
      Begin XtremeSuiteControls.FlatEdit txtAportes 
         Height          =   312
         Left            =   1920
         TabIndex        =   57
         Top             =   840
         Width           =   1572
         _Version        =   1572864
         _ExtentX        =   2773
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtRendimientos 
         Height          =   312
         Left            =   1920
         TabIndex        =   58
         Top             =   1200
         Width           =   1572
         _Version        =   1572864
         _ExtentX        =   2773
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtTotal 
         Height          =   312
         Left            =   4800
         TabIndex        =   59
         Top             =   480
         Width           =   2052
         _Version        =   1572864
         _ExtentX        =   3619
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtMulta 
         Height          =   312
         Left            =   4800
         TabIndex        =   60
         ToolTipText     =   "Monto en Multa para aplicación masiva por persona"
         Top             =   840
         Width           =   2052
         _Version        =   1572864
         _ExtentX        =   3619
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
         Text            =   "0"
         Alignment       =   1
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.PushButton cmdLiquidar 
         Height          =   615
         Left            =   10200
         TabIndex        =   61
         Top             =   480
         Width           =   1335
         _Version        =   1572864
         _ExtentX        =   2355
         _ExtentY        =   1085
         _StockProps     =   79
         Caption         =   "Procesar"
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
         Picture         =   "frmFNDLiquidacionPlan.frx":000C
         ImageAlignment  =   4
      End
      Begin VB.Label lblVencimiento 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Nuevo Vencimiento"
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
         Height          =   492
         Left            =   7680
         TabIndex        =   47
         Top             =   840
         Width           =   1092
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Casos"
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
         Left            =   600
         TabIndex        =   39
         Top             =   480
         Width           =   1092
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Aportes"
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
         Index           =   1
         Left            =   600
         TabIndex        =   38
         Top             =   840
         Width           =   1092
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Rendimientos"
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
         Index           =   2
         Left            =   600
         TabIndex        =   37
         Top             =   1200
         Width           =   1212
      End
      Begin VB.Label lbl 
         BackStyle       =   0  'Transparent
         Caption         =   "..."
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
         Left            =   4080
         TabIndex        =   36
         Top             =   1200
         Width           =   5892
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Total"
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
         Index           =   3
         Left            =   4080
         TabIndex        =   35
         Top             =   480
         Width           =   972
      End
      Begin VB.Image imgCalcula 
         Height          =   480
         Left            =   6960
         Picture         =   "frmFNDLiquidacionPlan.frx":0733
         ToolTipText     =   "Calcular Ejecución"
         Top             =   480
         Width           =   480
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Multa"
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
         Index           =   5
         Left            =   4080
         TabIndex        =   34
         Top             =   840
         Width           =   972
      End
   End
   Begin XtremeSuiteControls.TabControl tcMain 
      Height          =   4695
      Left            =   0
      TabIndex        =   13
      Top             =   2040
      Width           =   11775
      _Version        =   1572864
      _ExtentX        =   20764
      _ExtentY        =   8276
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
      Item(0).Caption =   "Contratos"
      Item(0).ControlCount=   2
      Item(0).Control(0)=   "chkMarcas"
      Item(0).Control(1)=   "vGrid"
      Item(1).Caption =   "Notas"
      Item(1).ControlCount=   2
      Item(1).Control(0)=   "txtNotas"
      Item(1).Control(1)=   "Label1(5)"
      Item(2).Caption =   "Filtros"
      Item(2).ControlCount=   28
      Item(2).Control(0)=   "chkMontos"
      Item(2).Control(1)=   "txtMntCorte"
      Item(2).Control(2)=   "txtMntInicio"
      Item(2).Control(3)=   "txtContratosSinMovMeses"
      Item(2).Control(4)=   "chkContratosSinMovAportes"
      Item(2).Control(5)=   "chkEstadoPersonaDiferente"
      Item(2).Control(6)=   "chkFondosCero"
      Item(2).Control(7)=   "chkFechas"
      Item(2).Control(8)=   "dtpInicio"
      Item(2).Control(9)=   "txtLinea"
      Item(2).Control(10)=   "txtLineaDesc"
      Item(2).Control(11)=   "chkLineas"
      Item(2).Control(12)=   "dtpCorte"
      Item(2).Control(13)=   "Label3(4)"
      Item(2).Control(14)=   "Label4(5)"
      Item(2).Control(15)=   "Label4(4)"
      Item(2).Control(16)=   "Label4(3)"
      Item(2).Control(17)=   "Label4(2)"
      Item(2).Control(18)=   "Label4(0)"
      Item(2).Control(19)=   "Label4(1)"
      Item(2).Control(20)=   "cboInstitucion"
      Item(2).Control(21)=   "cboEstado"
      Item(2).Control(22)=   "cboCreditos"
      Item(2).Control(23)=   "Label4(6)"
      Item(2).Control(24)=   "txtArchivo"
      Item(2).Control(25)=   "btnArchivo"
      Item(2).Control(26)=   "chkRndSinAporte"
      Item(2).Control(27)=   "chkMensualidad"
      Begin XtremeSuiteControls.FlatEdit txtArchivo 
         Height          =   312
         Left            =   -68800
         TabIndex        =   42
         Top             =   1320
         Visible         =   0   'False
         Width           =   6492
         _Version        =   1572864
         _ExtentX        =   11451
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
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.CheckBox chkLineas 
         Height          =   252
         Left            =   -62080
         TabIndex        =   26
         Top             =   960
         Visible         =   0   'False
         Width           =   1332
         _Version        =   1572864
         _ExtentX        =   2350
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Todas"
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
         Appearance      =   16
         Value           =   1
      End
      Begin FPSpreadADO.fpSpread vGrid 
         Height          =   3495
         Left            =   120
         TabIndex        =   14
         Top             =   720
         Width           =   11535
         _Version        =   524288
         _ExtentX        =   20346
         _ExtentY        =   6165
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
         MaxCols         =   484
         SpreadDesigner  =   "frmFNDLiquidacionPlan.frx":0DEA
         VScrollSpecial  =   -1  'True
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin XtremeSuiteControls.ComboBox cboInstitucion 
         Height          =   312
         Left            =   -68800
         TabIndex        =   23
         Top             =   600
         Visible         =   0   'False
         Width           =   6492
         _Version        =   1572864
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
      Begin XtremeSuiteControls.ComboBox cboEstado 
         Height          =   312
         Left            =   -64960
         TabIndex        =   24
         Top             =   3960
         Visible         =   0   'False
         Width           =   2772
         _Version        =   1572864
         _ExtentX        =   4895
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
      Begin XtremeSuiteControls.ComboBox cboCreditos 
         Height          =   312
         Left            =   -64960
         TabIndex        =   25
         Top             =   4320
         Visible         =   0   'False
         Width           =   2772
         _Version        =   1572864
         _ExtentX        =   4895
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
      Begin XtremeSuiteControls.CheckBox chkFondosCero 
         Height          =   252
         Left            =   -68440
         TabIndex        =   27
         Top             =   2160
         Visible         =   0   'False
         Width           =   5052
         _Version        =   1572864
         _ExtentX        =   8911
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Buscar únicamente Contratos con fondos en Cero"
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
         Appearance      =   16
      End
      Begin XtremeSuiteControls.CheckBox chkContratosSinMovAportes 
         Height          =   252
         Left            =   -68440
         TabIndex        =   28
         Top             =   2520
         Visible         =   0   'False
         Width           =   7212
         _Version        =   1572864
         _ExtentX        =   12721
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Contratos que no tengan movimiento de aportes por más de (x) meses"
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
         Appearance      =   16
      End
      Begin XtremeSuiteControls.CheckBox chkMontos 
         Height          =   252
         Left            =   -62080
         TabIndex        =   29
         Top             =   3240
         Visible         =   0   'False
         Width           =   1932
         _Version        =   1572864
         _ExtentX        =   3408
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Todos"
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
         Appearance      =   16
         Value           =   1
      End
      Begin XtremeSuiteControls.CheckBox chkFechas 
         Height          =   252
         Left            =   -62080
         TabIndex        =   30
         Top             =   3600
         Visible         =   0   'False
         Width           =   1932
         _Version        =   1572864
         _ExtentX        =   3408
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Todas"
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
         Appearance      =   16
         Value           =   1
      End
      Begin XtremeSuiteControls.CheckBox chkEstadoPersonaDiferente 
         Height          =   252
         Left            =   -62080
         TabIndex        =   31
         Top             =   3960
         Visible         =   0   'False
         Width           =   1932
         _Version        =   1572864
         _ExtentX        =   3408
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Diferente de ?"
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
         Appearance      =   16
         Value           =   1
      End
      Begin XtremeSuiteControls.CheckBox chkMarcas 
         Height          =   252
         Left            =   720
         TabIndex        =   32
         Top             =   360
         Width           =   6972
         _Version        =   1572864
         _ExtentX        =   12298
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "&Todas     [Lista de Contratos Localizados con los filtros actuales]"
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
         Appearance      =   16
      End
      Begin XtremeSuiteControls.PushButton btnArchivo 
         Height          =   315
         Left            =   -62080
         TabIndex        =   43
         ToolTipText     =   "Buscar Archivo de Excel (Columa: Identificacion, Hoja: Import)"
         Top             =   1320
         Visible         =   0   'False
         Width           =   372
         _Version        =   1572864
         _ExtentX        =   656
         _ExtentY        =   556
         _StockProps     =   79
         Caption         =   "..."
         BackColor       =   -2147483633
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
      Begin XtremeSuiteControls.CheckBox chkRndSinAporte 
         Height          =   252
         Left            =   -68440
         TabIndex        =   45
         Top             =   1800
         Visible         =   0   'False
         Width           =   5052
         _Version        =   1572864
         _ExtentX        =   8911
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Casos con Rendimiento pero aportes en Cero"
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
         Appearance      =   16
      End
      Begin XtremeSuiteControls.CheckBox chkMensualidad 
         Height          =   252
         Left            =   -62080
         TabIndex        =   46
         Top             =   2880
         Visible         =   0   'False
         Width           =   1932
         _Version        =   1572864
         _ExtentX        =   3408
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Mensualidad en Cero?"
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
         Appearance      =   16
      End
      Begin XtremeSuiteControls.DateTimePicker dtpCorte 
         Height          =   312
         Left            =   -63520
         TabIndex        =   51
         Top             =   3600
         Visible         =   0   'False
         Width           =   1332
         _Version        =   1572864
         _ExtentX        =   2350
         _ExtentY        =   550
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
         Enabled         =   0   'False
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   3
      End
      Begin XtremeSuiteControls.DateTimePicker dtpInicio 
         Height          =   312
         Left            =   -64960
         TabIndex        =   52
         Top             =   3600
         Visible         =   0   'False
         Width           =   1332
         _Version        =   1572864
         _ExtentX        =   2350
         _ExtentY        =   550
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
         Enabled         =   0   'False
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   3
      End
      Begin XtremeSuiteControls.FlatEdit txtLinea 
         Height          =   312
         Left            =   -68800
         TabIndex        =   54
         Top             =   960
         Visible         =   0   'False
         Width           =   1332
         _Version        =   1572864
         _ExtentX        =   2350
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
      Begin XtremeSuiteControls.FlatEdit txtLineaDesc 
         Height          =   312
         Left            =   -67480
         TabIndex        =   55
         Top             =   960
         Visible         =   0   'False
         Width           =   5172
         _Version        =   1572864
         _ExtentX        =   9123
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
      Begin XtremeSuiteControls.FlatEdit txtNotas 
         Height          =   3312
         Left            =   -67000
         TabIndex        =   50
         Top             =   720
         Visible         =   0   'False
         Width           =   7212
         _Version        =   1572864
         _ExtentX        =   12721
         _ExtentY        =   5842
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
      Begin XtremeSuiteControls.FlatEdit txtMntInicio 
         Height          =   315
         Left            =   -64960
         TabIndex        =   64
         Top             =   3240
         Visible         =   0   'False
         Width           =   1335
         _Version        =   1572864
         _ExtentX        =   2355
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
      Begin XtremeSuiteControls.FlatEdit txtMntCorte 
         Height          =   315
         Left            =   -63520
         TabIndex        =   65
         Top             =   3240
         Visible         =   0   'False
         Width           =   1335
         _Version        =   1572864
         _ExtentX        =   2355
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
      Begin XtremeSuiteControls.FlatEdit txtContratosSinMovMeses 
         Height          =   315
         Left            =   -67480
         TabIndex        =   66
         ToolTipText     =   "24"
         Top             =   2880
         Visible         =   0   'False
         Width           =   735
         _Version        =   1572864
         _ExtentX        =   1296
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin VB.Label Label4 
         Caption         =   "Archivo"
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
         Left            =   -69760
         TabIndex        =   41
         Top             =   1320
         Visible         =   0   'False
         Width           =   732
      End
      Begin VB.Label Label4 
         Caption         =   "Línea"
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
         Index           =   1
         Left            =   -69760
         TabIndex        =   22
         Top             =   960
         Visible         =   0   'False
         Width           =   972
      End
      Begin VB.Label Label4 
         Caption         =   "Institución"
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
         Left            =   -69760
         TabIndex        =   21
         Top             =   600
         Visible         =   0   'False
         Width           =   972
      End
      Begin VB.Label Label4 
         Caption         =   "Contratos con fechas de inicio entre...:"
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
         Index           =   2
         Left            =   -67960
         TabIndex        =   20
         Top             =   3600
         Visible         =   0   'False
         Width           =   3012
      End
      Begin VB.Label Label4 
         Caption         =   "Estado de la Persona...:"
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
         Left            =   -67960
         TabIndex        =   19
         Top             =   3960
         Visible         =   0   'False
         Width           =   3012
      End
      Begin VB.Label Label4 
         Caption         =   "Enlace con créditos...:"
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
         Left            =   -67960
         TabIndex        =   18
         Top             =   4320
         Visible         =   0   'False
         Width           =   3012
      End
      Begin VB.Label Label4 
         Caption         =   "Meses..:"
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
         Left            =   -68440
         TabIndex        =   17
         Top             =   2880
         Visible         =   0   'False
         Width           =   972
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Cnt. con Aportes + Rend. entre..:"
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
         Index           =   4
         Left            =   -67960
         TabIndex        =   16
         Top             =   3240
         Visible         =   0   'False
         Width           =   2892
      End
      Begin VB.Label Label1 
         Caption         =   "Notas de la liquidación General"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   732
         Index           =   5
         Left            =   -69040
         TabIndex        =   15
         Top             =   720
         Visible         =   0   'False
         Width           =   1932
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   240
      Top             =   5880
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFNDLiquidacionPlan.frx":1625
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComCtl2.FlatScrollBar FlatScrollBar 
      Height          =   252
      Left            =   8520
      TabIndex        =   2
      Top             =   480
      Width           =   492
      _ExtentX        =   873
      _ExtentY        =   450
      _Version        =   393216
      Arrows          =   65536
      Orientation     =   1638401
   End
   Begin MSComctlLib.ProgressBar prgBar 
      Align           =   2  'Align Bottom
      Height          =   135
      Left            =   0
      TabIndex        =   6
      Top             =   8670
      Visible         =   0   'False
      Width           =   12075
      _ExtentX        =   21299
      _ExtentY        =   238
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin XtremeSuiteControls.ComboBox cboProceso 
      Height          =   312
      Left            =   1800
      TabIndex        =   7
      Top             =   960
      Width           =   1812
      _Version        =   1572864
      _ExtentX        =   3201
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
   Begin XtremeSuiteControls.ComboBox cboRetencion 
      Height          =   330
      Left            =   1800
      TabIndex        =   8
      Top             =   1320
      Width           =   6495
      _Version        =   1572864
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
   Begin XtremeSuiteControls.ComboBox cboBanco 
      Height          =   312
      Left            =   4680
      TabIndex        =   9
      Top             =   1320
      Width           =   3612
      _Version        =   1572864
      _ExtentX        =   6376
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
   Begin XtremeSuiteControls.ComboBox cboTipoDocumento 
      Height          =   312
      Left            =   1800
      TabIndex        =   10
      Top             =   1320
      Width           =   1812
      _Version        =   1572864
      _ExtentX        =   3201
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
   Begin XtremeSuiteControls.ComboBox cboOperadora 
      Height          =   312
      Left            =   1800
      TabIndex        =   11
      Top             =   120
      Width           =   6492
      _Version        =   1572864
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
   Begin XtremeSuiteControls.ComboBox cboCuentaFiltro 
      Height          =   312
      Left            =   4680
      TabIndex        =   12
      Top             =   960
      Width           =   3612
      _Version        =   1572864
      _ExtentX        =   6376
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
   Begin XtremeSuiteControls.FlatEdit txtCodigo 
      Height          =   312
      Left            =   1800
      TabIndex        =   48
      Top             =   480
      Width           =   1332
      _Version        =   1572864
      _ExtentX        =   2350
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
   Begin XtremeSuiteControls.FlatEdit txtDescripcion 
      Height          =   312
      Left            =   3120
      TabIndex        =   49
      Top             =   480
      Width           =   5172
      _Version        =   1572864
      _ExtentX        =   9123
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
   Begin XtremeSuiteControls.PushButton btnAccion 
      Height          =   495
      Index           =   0
      Left            =   8640
      TabIndex        =   62
      Top             =   1200
      Width           =   1215
      _Version        =   1572864
      _ExtentX        =   2143
      _ExtentY        =   873
      _StockProps     =   79
      Caption         =   "Buscar"
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
      Picture         =   "frmFNDLiquidacionPlan.frx":172E
      ImageAlignment  =   4
   End
   Begin XtremeSuiteControls.PushButton btnAccion 
      Height          =   495
      Index           =   1
      Left            =   9840
      TabIndex        =   63
      Top             =   1200
      Width           =   1215
      _Version        =   1572864
      _ExtentX        =   2143
      _ExtentY        =   873
      _StockProps     =   79
      Caption         =   "Exportar"
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
      Picture         =   "frmFNDLiquidacionPlan.frx":1E2E
      ImageAlignment  =   4
   End
   Begin VB.Label lblStatus 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9720
      TabIndex        =   44
      Top             =   360
      Width           =   1335
   End
   Begin VB.Label lblBanco 
      BackStyle       =   0  'Transparent
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
      Height          =   375
      Left            =   3720
      TabIndex        =   5
      Top             =   1320
      Width           =   975
   End
   Begin VB.Label lblTipoDoc 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Tipo Doc."
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
      Left            =   840
      TabIndex        =   4
      Top             =   1320
      Width           =   975
   End
   Begin VB.Label lblConcepto 
      BackColor       =   &H00FFFFFF&
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
      Height          =   255
      Left            =   840
      TabIndex        =   3
      Top             =   1320
      Width           =   975
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Operadora"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
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
      TabIndex        =   1
      Top             =   120
      Width           =   1332
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Plan"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   312
      Left            =   600
      TabIndex        =   0
      Top             =   480
      Width           =   1332
   End
   Begin VB.Image imgBanner 
      Height          =   852
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   15732
   End
End
Attribute VB_Name = "frmFNDLiquidacionPlan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vScroll As Boolean, vPaso As Boolean
Dim vCuentaRet As String, vCuentaRetencion As String
Dim mGrupoBancario As String


Private Sub sbArchivo_Load()
Dim strSQL As String, rs As New ADODB.Recordset
Dim i As Integer, vCampos As Boolean, pLinea As Long


On Error GoTo vError

Me.MousePointer = vbHourglass

lblStatus.Caption = "Cargando Archivo de Filtrado!"
lblStatus.Refresh

Set rs = Excel_Load(txtArchivo.Text, "Import")
    
'Validaciónn del Archivo
vCampos = False
For i = 0 To rs.Fields.Count
     
    If UCase(LCase(rs.Fields(i).Name)) = "IDENTIFICACION" Then
       vCampos = True
    End If
     
     If vCampos Then Exit For
Next i

If Not vCampos Then
   MsgBox "No coincide la estructura del archivo a cargar..." & vbCrLf & _
         "Los campos son Identificacion¦ Nombre de la Hoja = Import"
   Exit Sub
End If


'FIN: Validación del Archivo


'Sube el Archivo

    pLinea = 0
    strSQL = ""
    
    Do While Not rs.EOF
      If Trim(rs!Identificacion) <> "" Then
        pLinea = pLinea + 1
        
        If pLinea = 1 Then
            strSQL = strSQL & Space(10) & "exec spFnd_Archivo_Ref '" & Trim(rs!Identificacion) & "',0, 1"
        Else
            strSQL = strSQL & Space(10) & "exec spFnd_Archivo_Ref '" & Trim(rs!Identificacion) & "',0, 0"
        End If
        
        If Len(strSQL) > 20000 Then
           Call ConectionExecute(strSQL)
           If glogon.error Then
              Exit Sub
           End If
           strSQL = ""
        End If
        
      End If
      rs.MoveNext
    Loop
    rs.Close

'Procesa Ultimo Bloque

If Len(strSQL) > 0 Then
   Call ConectionExecute(strSQL)
   If glogon.error Then
      Exit Sub
   End If
   strSQL = ""
End If

lblStatus.Caption = ""

Exit Sub

vError:

   lblStatus.Caption = ""
    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub


Private Sub sbConsultaPlan(vPlan As String)
Dim strSQL As String, rs As New ADODB.Recordset, itmX As ListItem
Dim mLInterBanca As Integer

On Error GoTo vError

Me.MousePointer = vbHourglass

tcMain.Item(0).Selected = True

If txtArchivo.Text <> "" Then
  Call sbArchivo_Load
Else
 strSQL = "Delete FND_ARCHIVO_REF"
 Call ConectionExecute(strSQL)
End If

'Default
mLInterBanca = 22


If cboTipoDocumento.Text = "Cheque" Or cboProceso.Text = "Retener" Then

    strSQL = "select F.cod_contrato,F.Cedula,S.Nombre,F.Estado,F.Plazo,F.Monto" _
           & ",F.Aportes,F.Rendimiento,F.Fecha_Corte,F.Fecha_Inicio" _
           & ",'' as CuentaAhorroX," & cboBanco.ItemData(cboBanco.ListIndex) & " as BancoX,Est.Descripcion as 'EstadoDesc'" _
           & " from Fnd_Contratos F inner join Socios S on F.Cedula = S.Cedula" _
           & " inner join AFI_ESTADOS_PERSONA Est on S.estadoActual = Est.cod_Estado" _
           & " Where F.Cod_operadora=" & cboOperadora.ItemData(cboOperadora.ListIndex) _
           & " and F.Cod_plan='" & vPlan & "' and F.Estado <> 'L'"
Else
    
    
        strSQL = "select Bg.LCTA_INTERNA, Bg.LCTA_INTERBANCARIA " _
               & " from TES_BANCOS Tb inner join TES_BANCOS_GRUPOS Bg on Tb.COD_GRUPO = Bg.COD_GRUPO" _
               & " Where Tb.ID_BANCO = " & cboBanco.ItemData(cboBanco.ListIndex)
        Call OpenRecordSet(rs, strSQL)
        If rs.EOF And rs.BOF Then
           mLInterBanca = 22
        Else
           mLInterBanca = rs!LCTA_InterBancaria
        End If
        rs.Close
    
    strSQL = "select F.cod_contrato,F.Cedula,S.Nombre,F.Estado,F.Plazo,F.Monto" _
           & ",F.Aportes,F.Rendimiento,F.Fecha_Corte,F.Fecha_Inicio" _
           & ",dbo.fxSys_Cuentas_Bancarias(F.cedula,B.id_Banco,0) as CuentaAhorroX" _
           & ",B.id_Banco as BancoX,B.descripcion as BancoDesc,Est.Descripcion as 'EstadoDesc'" _
           & " from Fnd_Contratos F inner join Socios S on F.Cedula = S.Cedula" _
           & " inner join Fnd_Planes Pln on F.cod_Operadora = Pln.Cod_Operadora and F.cod_Plan = Pln.cod_Plan " _
           & " inner join AFI_ESTADOS_PERSONA Est on S.estadoActual = Est.cod_Estado" _
           & " inner join Tes_Bancos B on B.id_Banco = " & cboBanco.ItemData(cboBanco.ListIndex)
           
    If cboCuentaFiltro.Text <> "TODOS" Then
       If cboCuentaFiltro.Text = "Cuenta Interna" Then
            strSQL = strSQL & " inner join vSys_Personas_Cuenta_Bancaria_Local Cta on F.cedula = Cta.Identificacion" _
                   & " and Cta.cod_Banco = B.cod_Grupo and Cta.cod_Divisa = Pln.Cod_Moneda"
       Else
            strSQL = strSQL & " inner join vSys_Personas_Cuenta_Bancaria_Interbancaria Cta on F.cedula = Cta.Identificacion" _
                   & " and Cta.cod_Banco = B.cod_Grupo and Cta.cod_Divisa = Pln.Cod_Moneda"
       End If
    End If
           
           
    strSQL = strSQL & " Where F.Cod_operadora=" & cboOperadora.ItemData(cboOperadora.ListIndex) _
           & " and F.Cod_plan='" & vPlan & "' and F.Estado <> 'L'" _
           & " and dbo.fxSys_Cuentas_Bancarias(F.cedula,B.id_Banco,0) <>  ''"

    If cboCuentaFiltro.Text = "Interbancaria Mismo Banco" Then
           strSQL = strSQL & " and substring(dbo.fxSys_Cuentas_Bancarias(F.cedula,B.id_Banco,0), 1,10) like '%" & mGrupoBancario & "%'"
    End If


End If
       
       
If cboInstitucion.Text <> "TODOS" Then
   strSQL = strSQL & " and S.cod_institucion = " & cboInstitucion.ItemData(cboInstitucion.ListIndex)
End If
       
If chkLineas.Value = vbUnchecked Then
  strSQL = strSQL & " and F.cedula in(select cedula from reg_creditos where estado in('A','C')" _
          & " and codigo = '" & txtLinea.Text & "')"

End If


'Filtros Adicionales
If chkFechas.Value = vbUnchecked Then
   strSQL = strSQL & " and F.Fecha_inicio between '" & Format(dtpInicio.Value, "yyyy/mm/dd") _
          & " 00:00:00' and '" & Format(dtpCorte.Value, "yyyy/mm/dd") & " 23:59:59'"
End If



If cboEstado.Text <> "TODOS" Then
   If chkEstadoPersonaDiferente.Value = vbChecked Then
      strSQL = strSQL & " and S.EstadoActual not in('" & cboEstado.ItemData(cboEstado.ListIndex) & "')"
   Else
      strSQL = strSQL & " and S.EstadoActual in('" & cboEstado.ItemData(cboEstado.ListIndex) & "')"
   End If
End If

If chkFondosCero.Value = vbChecked Then
    strSQL = strSQL & " and (F.Aportes + F.Rendimiento) = 0"
Else
    If chkMontos.Value = vbUnchecked Then
        strSQL = strSQL & " and (F.Aportes + F.Rendimiento) between " & CCur(txtMntInicio.Text) & " and " & CCur(txtMntCorte.Text)
    End If
End If


If chkContratosSinMovAportes.Value = vbChecked Then
    strSQL = strSQL & " and datediff(m, dbo.fxFndFechaUltAporte(F.cod_operadora,F.cod_plan,F.cod_contrato),dbo.MyGetdate()) > " & txtContratosSinMovMeses.Text
End If

'Con Rendimiento pero Sin Aportes (Inconsistencia)
If chkRndSinAporte.Value = xtpChecked Then
   strSQL = strSQL & " and F.Aportes = 0 and F.Rendimiento > 0"
End If


'Mensualidad en 0 (y no es certificado a plazo)
If chkMensualidad.Value = xtpChecked Then
   strSQL = strSQL & " and F.Monto = 0 and isnull(F.Inversion,0) = 0"
End If


If cboCreditos.Text <> "TODOS" Then
   Select Case cboCreditos.Text
        Case "Persona -> Con créditos activos"
            strSQL = strSQL & " and S.cedula in(select cedula from reg_creditos V " _
                   & " inner join Catalogo C on V.codigo = C.codigo and C.retencion = 'N' and C.poliza = 'N'" _
                   & " where V.saldo > 0 and V.estado = 'A' group by V.cedula)"
        
        Case "Persona -> Con créditos en Mora"
            strSQL = strSQL & " and S.cedula in(select V.cedula from vista_morosidad V " _
                   & " inner join Catalogo C on V.codigo = C.codigo and C.retencion = 'N' and C.poliza = 'N' group by V.cedula)"
                   
                   
        Case "Persona -> Sin créditos activos"
            strSQL = strSQL & " and S.cedula not in(select cedula from reg_creditos where saldo > 0 and estado = 'A' group by cedula)"
        
        Case "Persona -> Sin créditos en Mora"
            strSQL = strSQL & " and S.cedula not in(select v.cedula from vista_morosidad v " _
                   & " inner join Catalogo C on V.codigo = C.codigo and C.retencion = 'N' and C.poliza = 'N' group by V.cedula)"
   
        Case "Persona -> En Cobro Jud y/o Traspaso"
            strSQL = strSQL & " and S.cedula in(select cedula from reg_creditos where saldo > 0 and estado = 'A' and proceso <> 'N' group by cedula)"
   End Select

End If

If txtArchivo.Text <> "" Then
            strSQL = strSQL & " and S.cedula in(select cedula from FND_ARCHIVO_REF)"
End If


Call OpenRecordSet(rs, strSQL)

vPaso = True

vGrid.MaxRows = 0


Do While Not rs.EOF
 vGrid.MaxRows = vGrid.MaxRows + 1
 vGrid.Row = vGrid.MaxRows
 vGrid.Col = 1
 vGrid.Value = chkMarcas.Value
 vGrid.Col = 2
 vGrid.Text = rs!COD_CONTRATO
 vGrid.Col = 3
 vGrid.Text = rs!Cedula
 vGrid.Col = 4
 vGrid.Text = rs!Nombre
 vGrid.Col = 5
 vGrid.Text = Format(rs!APORTES, "standard")
 vGrid.Col = 6
 vGrid.Text = Format(rs!Rendimiento, "standard")
 
     If cboTipoDocumento.Text = "Cheque" Then
        If cboBanco.Text = "TODOS" Then
            vGrid.Col = 7
            vGrid.Text = CStr(rs!BancoX)
        Else
            vGrid.Col = 7
            vGrid.Text = CStr(cboBanco.ItemData(cboBanco.ListIndex))
        End If
        vGrid.Col = 8
        vGrid.Text = "0"
     Else
        vGrid.Col = 7
        vGrid.Text = CStr(rs!BancoX)
        vGrid.Col = 8
        vGrid.Text = Trim(rs!CuentaAhorroX & "")
     End If

 
 
 vGrid.Col = 9
 vGrid.Text = Format(IIf(IsNull(rs!Fecha_Corte), rs!Fecha_Inicio, rs!Fecha_Corte), "yyyy/mm/dd")

 vGrid.Col = 10
 vGrid.Text = rs!EstadoDesc & ""

 rs.MoveNext
Loop
rs.Close

vPaso = False

Me.MousePointer = vbDefault

Call imgCalcula_Click

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub sbExportar()

Dim vHeaders As vGridHeaders
    vHeaders.Columnas = 10
    vHeaders.Headers(1) = "Check"
    vHeaders.Headers(2) = "No. Contrato"
    vHeaders.Headers(3) = "Identificación"
    vHeaders.Headers(4) = "Nombre"
    vHeaders.Headers(5) = "Aportes"
    vHeaders.Headers(6) = "Rendimientos"
    vHeaders.Headers(7) = "Cta. Id."
    vHeaders.Headers(8) = "Cuenta"
    vHeaders.Headers(9) = "Corte"
    vHeaders.Headers(10) = "Estado"

 Call sbSIFGridExportar(vGrid, vHeaders, "ProGrX_Planes_Liquidacion_" & Trim(txtCodigo.Text))

End Sub


Private Sub btnAccion_Click(Index As Integer)

Select Case Index
  Case 0  'Buscar
        Call sbConsultaPlan(Trim(txtCodigo.Text))
  Case 1  'Exportar
        Call sbExportar
End Select

End Sub

Private Sub btnArchivo_Click()

With frmContenedor.CD
        .InitDir = "C:\"
        .DialogTitle = "Localice Archivo de Planilla [Microsoft EXCEL]..."
        .Filter = "Excel|*.xlsx|Excel 97-2003|*.xls"
        .ShowOpen

        If .FileName = "" Then
            MsgBox "Archivo no válido...", vbExclamation
            Exit Sub
        End If

        If UCase(Right(.FileName, 3)) = "XLS" Or UCase(Right(.FileName, 4)) = "XLSX" Then
            'Ok
        Else
            MsgBox "La Extensión del Archivo no es válido...", vbExclamation
            Exit Sub
        End If
        
        txtArchivo.Text = .FileName
    
End With


End Sub



Private Sub sbGrupoBancario(pBancoId As String)
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

strSQL = "select dbo.fxTes_BancoSFN(" & pBancoId & ") as 'Codigo'"
Call OpenRecordSet(rs, strSQL)
  mGrupoBancario = rs!Codigo
rs.Close

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub cboBanco_Click()
If vPaso Then Exit Sub
If cboBanco.ListCount = 0 Then Exit Sub

vGrid.MaxRows = 0

Call sbGrupoBancario(cboBanco.ItemData(cboBanco.ListIndex))

End Sub

Private Sub cboInstitucion_Click()
If vPaso Then Exit Sub
vGrid.MaxRows = 0
End Sub

Private Sub cboInstitucion_KeyDown(KeyCode As Integer, Shift As Integer)

If (KeyCode = vbKeyReturn Or KeyCode = vbKeyTab) And txtLinea.Enabled Then txtLinea.SetFocus

End Sub

Private Sub cboOperadora_Click()
If vPaso Then Exit Sub
vCuentaRet = fxCuentaRetiros

Call txtCodigo_LostFocus
If Trim(txtCodigo) <> "" Then Call sbConsultaPlan(Trim(txtCodigo))
End Sub


Private Sub cboOperadora_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtCodigo.SetFocus
End Sub


Private Function fxCuentaPlan(pTipo As String) As String
Dim strSQL As String, rs As New ADODB.Recordset

fxCuentaPlan = ""

If pTipo = "P" Then
 'Cuenta del Plan : Aportes
  strSQL = "Select Cuenta_Conta as CuentaX from Fnd_Planes Where Cod_Operadora=" & cboOperadora.ItemData(cboOperadora.ListIndex) _
         & " and Cod_Plan='" & txtCodigo & "'"
Else
 'Cuenta del Plan : Rendimiento
  strSQL = "Select Cuenta_Rendimiento as CuentaX from Fnd_Planes Where Cod_Operadora=" & cboOperadora.ItemData(cboOperadora.ListIndex) _
         & " and Cod_Plan='" & txtCodigo & "'"

End If
 
Call OpenRecordSet(rs, strSQL)
If Not rs.EOF And Not rs.BOF Then
    fxCuentaPlan = Trim(rs!CuentaX)
End If
rs.Close

End Function


Private Function fxCuentaRetencion() As String
Dim strSQL As String, rs As New ADODB.Recordset

fxCuentaRetencion = ""
 
strSQL = "select COD_CUENTA  From FND_RETENCION_CONCEPTOS" _
       & " where RETENCION_CODIGO = '" & cboRetencion.ItemData(cboRetencion.ListIndex) & "'"
 
Call OpenRecordSet(rs, strSQL)
If Not rs.EOF And Not rs.BOF Then
    fxCuentaRetencion = Trim(rs!cod_cuenta)
End If
rs.Close

End Function


Private Function fxCuentaRetiros() As String
Dim strSQL As String, rs As New ADODB.Recordset

fxCuentaRetiros = ""

strSQL = "Select cta_retiros from Fnd_operadoras Where Cod_Operadora=" & cboOperadora.ItemData(cboOperadora.ListIndex)
 
Call OpenRecordSet(rs, strSQL)
If Not rs.EOF And Not rs.BOF Then
    fxCuentaRetiros = Trim(rs!Cta_retiros)
End If
rs.Close

End Function

Private Function fxCuentaIngresos() As String
Dim strSQL As String, rs As New ADODB.Recordset

fxCuentaIngresos = ""

strSQL = "Select cta_Ingresos from Fnd_operadoras Where Cod_Operadora = " & cboOperadora.ItemData(cboOperadora.ListIndex)
 
Call OpenRecordSet(rs, strSQL)
If Not rs.EOF And Not rs.BOF Then
    fxCuentaIngresos = Trim(rs!Cta_Ingresos)
End If
rs.Close

End Function


Private Sub sbDocumento(vTipoDoc As String, vDocRef As String, vFecha As Date, vConcepto As String)
Dim strSQL As String, rs As New ADODB.Recordset

Dim vAporteLiq As Currency, vRendiLiq As Currency, vContrato As String, curMulta As Currency
Dim vCuenta As String, vOperadora As Long
Dim vDivisa As String

vDivisa = fxFndDivisa(vOperadora, txtCodigo.Text)
vOperadora = cboOperadora.ItemData(cboOperadora.ListIndex)


strSQL = "Select P.Cod_Operadora,P.Cod_Plan,P.Cuenta_Conta,P.Cuenta_Rendimiento,  L.cod_Cuenta, P.CUENTA_IMPUESTOS as 'ISR_Cta'" _
       & ",max(L.cod_Contrato) as 'Cod_Contrato'" _
       & ",isnull(sum(L.aportes_liq),0) as 'Aporte', isnull(sum(L.Rendi_Liq),0) as 'Rendimiento'" _
       & ", isnull(sum(L.multa_retiro),0) as 'Multa'" _
       & ", isnull(sum(L.ISR_MONTO),0) as 'ISR_Monto'" _
       & " from  Fnd_Liquidacion L inner join Fnd_Planes P on L.cod_operadora = P.cod_Operadora and L.cod_Plan = P.cod_Plan" _
       & " Where L.Cod_Operadora = '" & vOperadora & "' and L.Cod_Plan = '" & txtCodigo.Text _
       & "' and L.LIQ_PLAN = '" & vDocRef & "'" _
       & " group by P.Cod_Operadora,P.Cod_Plan,P.Cuenta_Conta,P.Cuenta_Rendimiento,L.cod_Cuenta, P.CUENTA_IMPUESTOS"
       
'Está configurado el proceso para que se realice plan por plan. Si llegase a cambiar se deben de hacer los ajustes del caso
'y este Ciclo se vuelve determinando porque actualmente solo devuelve un registro
Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF

       vAseDocDetalle = Mid(vAseDocDetalle, 1, 30)

       ' Documentos 2
       Call sbDocumentoMaestro("FLIQ", rs!COD_OPERADORA, rs!COD_PLAN, rs!COD_CONTRATO _
                        , rs!Aporte, vDocRef, vConcepto, rs!COD_PLAN, txtDescripcion.Text, rs!Rendimiento)
       
            If rs!Aporte > 0 Then
                'Registro de Aporte Liquidado
                    strSQL = "exec spSIFDocsAsiento 'FLIQ','" & vDocRef & "'," & rs!Aporte & "" _
                           & ",'D','" & vDivisa & "',1," & GLOBALES.gEnlace & ",'" & GLOBALES.gOficinaUnidad & "'," _
                           & "'','" & Trim(rs!Cuenta_Conta) _
                           & "','" & rs!COD_OPERADORA & "','" & rs!COD_PLAN & "','" & vAseDocDetalle & "'"
    
                Call ConectionExecute(strSQL)
            End If
            
            If rs!Rendimiento > 0 Then
                'Registro de Rendimiento Liquidado
                strSQL = "exec spSIFDocsAsiento 'FLIQ','" & vDocRef & "'," & rs!Rendimiento & "" _
                       & ",'D','" & vDivisa & "',1," & GLOBALES.gEnlace & ",'" & GLOBALES.gOficinaUnidad & "'," _
                       & "'" & GLOBALES.gOficinaCentroCosto & "','" & Trim(rs!Cuenta_Rendimiento) _
                       & "','" & rs!COD_OPERADORA & "','" & rs!COD_PLAN & "','" & vAseDocDetalle & "'"
    
                Call ConectionExecute(strSQL)
            End If

            If rs!ISR_MONTO > 0 Then
                'ISR
                strSQL = "exec spSIFDocsAsiento 'FLIQ','" & vDocRef & "'," & rs!ISR_MONTO & "" _
                       & ",'C','" & vDivisa & "',1," & GLOBALES.gEnlace & ",'" & GLOBALES.gOficinaUnidad & "'," _
                       & "'" & GLOBALES.gOficinaCentroCosto & "','" & Trim(rs!ISR_Cta) _
                       & "','" & rs!COD_OPERADORA & "','" & rs!COD_PLAN & "','" & vAseDocDetalle & "'"
    
                Call ConectionExecute(strSQL)
            End If

            
            
            If rs!Multa > 0 Then
            
                vCuenta = fxCuentaIngresos
                
                strSQL = "exec spSIFDocsAsiento 'FLIQ','" & vDocRef & "'," & rs!Multa _
                       & ",'C','" & vDivisa & "',1," & GLOBALES.gEnlace & ",'" & GLOBALES.gOficinaUnidad & "'," _
                       & "'" & GLOBALES.gOficinaCentroCosto & "','" & vCuenta _
                       & "','" & rs!COD_OPERADORA & "','" & rs!COD_PLAN & "','" & vAseDocDetalle & "'"
                
                Call ConectionExecute(strSQL)
            End If
            
            If rs!Aporte + rs!Rendimiento - (rs!Multa + rs!ISR_MONTO) > 0 Then
                strSQL = "exec spSIFDocsAsiento 'FLIQ','" & vDocRef & "'," & rs!Aporte + rs!Rendimiento - (rs!Multa + rs!ISR_MONTO) _
                       & ",'C','" & vDivisa & "',1," & GLOBALES.gEnlace & ",'" & GLOBALES.gOficinaUnidad & "'," _
                       & "'','" & rs!cod_cuenta _
                       & "','" & rs!COD_OPERADORA & "','" & rs!COD_PLAN & "','" & vAseDocDetalle & "'"
                Call ConectionExecute(strSQL)
            End If

rs.MoveNext
Loop
rs.Close

End Sub


Private Sub sbProceso(lng As Long, vFecha As Date, vTipoDoc As String, vConcepto As String, vDocRef As String, vCuenta)
Dim strSQL As String, rs As New ADODB.Recordset
Dim vOperadora As Long, vPlan As String, vContrato As Long, vTipo As String
Dim vProceso  As Long, vLiq As Long, vMonto As Currency
Dim vRendiLiq As Currency, vAporteLiq As Currency
Dim vMulta As Currency, vCuentaAhorros As String
Dim vBanco As String

On Error GoTo vError

Me.MousePointer = vbHourglass

vTipo = "L"
vOperadora = cboOperadora.ItemData(cboOperadora.ListIndex)
vPlan = Trim(txtCodigo.Text)

vGrid.Row = lng
vGrid.Col = 2
'vContrato = lsw.ListItems.Item(lng).Text

vContrato = vGrid.Text


If IsNumeric(txtMulta.Text) Then
   vMulta = CCur(txtMulta.Text)
Else
    vMulta = 0
End If

vProceso = Year(vFecha) & Format(Month(vFecha), "00")

'vAporteLiq = CCur(lsw.ListItems.Item(lng).SubItems(3))
'vRendiLiq = CCur(lsw.ListItems.Item(lng).SubItems(4))
'vMonto = (CCur(lsw.ListItems.Item(lng).SubItems(3)) + CCur(lsw.ListItems.Item(lng).SubItems(4)))
'vCuentaAhorros = Trim(lsw.ListItems.Item(lng).SubItems(6))


vGrid.Col = 5
vAporteLiq = CCur(vGrid.Text)
vGrid.Col = 6
vRendiLiq = CCur(vGrid.Text)

vMonto = vAporteLiq + vRendiLiq


If vMulta > vMonto Then
   vMulta = vMonto
End If

vGrid.Col = 7
vBanco = vGrid.Text

vGrid.Col = 8
vCuentaAhorros = vGrid.Text


If Mid(cboProceso.Text, 1, 1) = "R" Then
        'Retención
        strSQL = "Insert Fnd_liquidacion(Cod_operadora,Cod_plan,Cod_contrato,Aportes," _
               & "Rendimiento,Aportes_liq,Rendi_liq,Fecha,usuario,Tipo_gestion,Cod_banco,Cta_ahorros," _
               & "Tipo,multa_retiro,traspaso_tesoreria,traspaso_usuario,solicitud_tesoreria" _
               & ",cod_oficina,notas,cod_cuenta,retencion_codigo,LIQ_PLAN)" _
               & " Values(" & vOperadora & ",'" & vPlan & "'," & vContrato & "," _
               & vAporteLiq & "," & vRendiLiq & "," & vAporteLiq & "," _
               & vRendiLiq & ",dbo.MyGetdate(),'" & glogon.Usuario & "','" & vTipo & "'," _
               & "0,'" & vCuentaAhorros & "','" & fxgFNDTipoPago("D", cboTipoDocumento) & "'," & vMulta _
               & ",dbo.MyGetdate(),'" & glogon.Usuario & "',0,'" & GLOBALES.gOficinaTitular & "','" & txtNotas.Text _
               & "','" & vCuenta & "','" & cboRetencion.ItemData(cboRetencion.ListIndex) & "','" & vDocRef & "')"

Else
        'Desembolso
        strSQL = "Insert Fnd_liquidacion(Cod_operadora,Cod_plan,Cod_contrato,Aportes," _
               & "Rendimiento,Aportes_liq,Rendi_liq,Fecha,usuario,Tipo_gestion,Cod_banco,Cta_ahorros," _
               & "Tipo,multa_retiro,cod_oficina,notas,cod_cuenta,LIQ_PLAN)" _
               & " Values(" & vOperadora & ",'" & vPlan & "'," & vContrato & "," _
               & vAporteLiq & "," & vRendiLiq & "," & vAporteLiq & "," _
               & vRendiLiq & ",dbo.MyGetdate(),'" & glogon.Usuario & "','" & vTipo & "'," _
               & vBanco & ",'" & vCuentaAhorros & "','" _
               & fxgFNDTipoPago("D", cboTipoDocumento) & "'," & vMulta & ",'" & GLOBALES.gOficinaTitular & "','" & txtNotas.Text _
               & "','" & vCuenta & "','" & vDocRef & "')"
End If
Call ConectionExecute(strSQL)
 
strSQL = "Select max(Consec) as Consec From Fnd_liquidacion" _
       & " where cod_Operadora = " & vOperadora & " and cod_Plan = '" & vPlan & "' and cod_contrato = " & vContrato
Call OpenRecordSet(rs, strSQL)
    vLiq = rs!consec
rs.Close



strSQL = "Update Fnd_contratos set Estado = 'L', Aportes = 0, Rendimiento = 0, Liq_Tipo = 'L'" _
       & ",Liq_Fecha='" & Format(vFecha, "yyyy/mm/dd") & "',Liq_Monto=" & vMonto & ",Liq_Retiro=" & vMonto _
       & ",Liq_Neto=" & vMonto _
       & " Where cod_operadora = " & vOperadora & " and cod_plan = '" & vPlan & "' and Cod_contrato = " & vContrato

strSQL = strSQL & Space(10) & "Insert fnd_contratos_detalle(Cod_operadora,Cod_plan,Cod_Contrato" _
       & ",Fecha,Monto,Fecha_Proceso,Tcon,Ncon,cod_concepto,usuario,cod_caja,REF_01)" _
       & " values(" & vOperadora & ",'" & vPlan & "'," & vContrato & ",dbo.MyGetdate()," _
       & vMonto * -1 & "," & vProceso & ",'" & vTipoDoc & "','" & vLiq & "','" & vConcepto _
       & "','" & glogon.Usuario & "','','" & vDocRef & "')"
Call ConectionExecute(strSQL)

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub




Private Sub cboProceso_Click()

If vPaso Then Exit Sub

lblBanco.Visible = False
cboBanco.Visible = False
lblTipoDoc.Visible = False
cboTipoDocumento.Visible = False
lblConcepto.Visible = False
cboRetencion.Visible = False
cboCuentaFiltro.Visible = False
chkTarjetaActiva_Valida.Visible = False

If Mid(cboProceso.Text, 1, 1) = "D" Then
    lblBanco.Visible = True
    cboBanco.Visible = True
    lblTipoDoc.Visible = True
    cboTipoDocumento.Visible = True
    cboCuentaFiltro.Visible = True
    chkTarjetaActiva_Valida.Visible = True
Else
    lblConcepto.Visible = True
    cboRetencion.Visible = True
End If

End Sub

Private Sub cboRetencion_Click()

If vPaso Then Exit Sub
If cboRetencion.ListCount <= 0 Then Exit Sub

Dim strSQL As String, rs As New ADODB.Recordset

strSQL = "select cod_cuenta from FND_RETENCION_CONCEPTOS where Retencion_Codigo = '" & cboRetencion.ItemData(cboRetencion.ListIndex) & "'"
Call OpenRecordSet(rs, strSQL)
  vCuentaRetencion = Trim(rs!cod_cuenta)
rs.Close

End Sub

Private Sub cboTipo_Click()

If Mid(cboTipo.Text, 1, 1) = "R" Then
    lblVencimiento.Visible = True
Else
    lblVencimiento.Visible = False
End If

dtpVence.Visible = lblVencimiento.Visible

End Sub

Private Sub cboTipoDocumento_Change()
vGrid.MaxRows = 0
End Sub

Private Sub chkFechas_Click()

If chkFechas.Value = vbChecked Then
   dtpInicio.Enabled = False
Else
   dtpInicio.Enabled = True
End If

dtpCorte.Enabled = dtpInicio.Enabled

End Sub

Private Sub chkLineas_Click()

If chkLineas.Value = vbChecked Then
  txtLinea.Enabled = False
Else
  txtLinea.Enabled = True
End If

txtLineaDesc.Enabled = txtLinea.Enabled
vGrid.MaxRows = 0

End Sub

Private Sub chkMarcas_Click()
Dim lng As Long

'For lng = 1 To lsw.ListItems.Count
'   lsw.ListItems.Item(lng).Checked = chkMarcas.Value
'Next lng

For lng = 1 To vGrid.MaxRows
 vGrid.Row = lng
 vGrid.Col = 1
 vGrid.Value = chkMarcas.Value
Next lng


Call imgCalcula_Click

End Sub


Private Sub chkMontos_Click()
If chkMontos.Value = vbChecked Then
   txtMntInicio.Enabled = False
Else
   txtMntInicio.Enabled = True
End If

txtMntCorte.Enabled = txtMntInicio.Enabled

End Sub

Private Sub cmdLiquidar_Click()
Dim strSQL As String, rs As New ADODB.Recordset
Dim lng As Long, vFecha As Date, vCuenta As String
Dim vConcepto As String, vTipoDoc As String, vDocRef As String


strSQL = MsgBox("Confirma La Liquidacion General?", vbExclamation + vbYesNo)
If strSQL = vbNo Then Exit Sub

Me.MousePointer = vbHourglass

lbl.Caption = "Cargando..."
lbl.Refresh

PrgBar.Max = vGrid.MaxRows + 1
PrgBar.Value = 1

PrgBar.Visible = True

vConcepto = "FND006"
vTipoDoc = "FLIQ"

vFecha = fxFechaServidor

If Mid(cboProceso.Text, 1, 1) = "R" Then
   vCuenta = fxCuentaRetencion 'Cuenta ligada al codigo de Retencion
Else
   vCuenta = fxCuentaRetiros   'Cuenta Default para Retiros por Operadora
End If


'Consecutivo del Documento Creado (Documento General)
strSQL = "exec spFndPlanIdLiqGen " & cboOperadora.ItemData(cboOperadora.ListIndex) & ",'" & txtCodigo.Text & "'"
Call OpenRecordSet(rs, strSQL)
  vFecha = Format(rs!fecha, "yyyy/mm/dd")
  vDocRef = Trim(txtCodigo.Text) & "." & Format(rs!Consecutivo, "000") & "_" & Format(rs!fecha, "yyyy.mm.dd")
rs.Close

'Inicializa Lote
strSQL = ""

Dim vOperadora As Long, vPlan As String, vContrato As Long, vTipo As String
Dim vRendiLiq As Currency, vAporteLiq As Currency, vMulta As Currency
Dim vBanco As String, vCuentaAhorros As String, vBancoTipo As String, vRetencion As String

vTipo = "L"
vOperadora = cboOperadora.ItemData(cboOperadora.ListIndex)
vPlan = Trim(txtCodigo.Text)
vRetencion = cboRetencion.ItemData(cboRetencion.ListIndex)

vBancoTipo = fxTipoDocumento(cboTipoDocumento.Text)

If IsNumeric(txtMulta.Text) Then
   vMulta = CCur(txtMulta.Text)
Else
   vMulta = 0
End If


For lng = 1 To vGrid.MaxRows
 vGrid.Row = lng
 vGrid.Col = 1
 If vGrid.Value = vbChecked Then
    vGrid.Col = 2
    vContrato = vGrid.Text
    vGrid.Col = 5
    vAporteLiq = CCur(vGrid.Text)
    vGrid.Col = 6
    vRendiLiq = CCur(vGrid.Text)

    vGrid.Col = 7
    vBanco = vGrid.Text
    
    vGrid.Col = 8
    vCuentaAhorros = vGrid.Text
    
    
    lbl.Caption = "Procesando Contrato : " & vContrato & " [ " & PrgBar.Value & " / " & PrgBar.Max & " ]"
    lbl.Refresh
    
'    'Procesa Liquidacion
'    Call sbProceso(lng, vFecha, vTipoDoc, vConcepto, vDocRef, vCuenta)
    
'--Fondos: Liquidacion Masiva> Proceso Complementario
    strSQL = strSQL & Space(10) & "exec spFndRetLiq_Masivo_Complemento " & vOperadora & ",'" & vPlan & "'," & vContrato _
           & ",'" & vTipo & "','" & vTipoDoc & "','" & vConcepto & "','" & vDocRef _
           & "'," & vAporteLiq & "," & vRendiLiq & "," & vMulta _
           & ",'" & Mid(txtNotas.Text, 1, 1000) & "','" & glogon.Usuario & "','" & GLOBALES.gOficinaTitular _
           & "','" & Mid(cboProceso.Text, 1, 1) & "','" & vRetencion & "','" & vCuenta & "'," & vBanco & ",'" & vBancoTipo & "','" & vCuentaAhorros _
           & "','ProGrX','" & Mid(cboTipo.Text, 1, 1) & "','" & Format(dtpVence.Value, "yyyy/mm/dd") & "'"
        
    If Len(strSQL) > 30000 Then
       Call ConectionExecute(strSQL)
       strSQL = ""
    End If
    
 End If
 
 PrgBar.Value = PrgBar.Value + 1
 
Next lng

'Cierra Lote
If Len(strSQL) > 0 Then
   Call ConectionExecute(strSQL)
   strSQL = ""
End If


lbl.Caption = "Creando Documento y Asiento General..."
lbl.Refresh

Call sbDocumento(vTipoDoc, vDocRef, vFecha, vConcepto)

lbl.Caption = "Cargando..."
lbl.Refresh

txtArchivo.Text = ""

Call sbConsultaPlan(txtCodigo.Text)

PrgBar.Visible = False
lbl.Caption = ""

Me.MousePointer = vbDefault

End Sub

Private Sub FlatScrollBar_Change()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

If vScroll Then
    strSQL = "select Top 1 cod_plan,descripcion, dateadd(year,1, getdate()) as 'Vence' from fnd_planes" _
           & " where cod_operadora = " & cboOperadora.ItemData(cboOperadora.ListIndex)
    
    If FlatScrollBar.Value = 1 Then
       strSQL = strSQL & " and cod_plan > '" & txtCodigo & "' order by cod_plan asc"
    Else
       strSQL = strSQL & " and cod_plan < '" & txtCodigo & "' order by cod_plan desc"
    End If
    
    Call OpenRecordSet(rs, strSQL)
    If Not rs.EOF And Not rs.BOF Then
      txtCodigo.Text = rs!COD_PLAN
      txtDescripcion.Text = rs!Descripcion
      dtpVence.Value = rs!Vence
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
vModulo = 18 'Fondo de Inversion

End Sub

Private Sub Form_Load()
Dim strSQL As String

vModulo = 18 'Fondo de Inversion

vPaso = True

mGrupoBancario = ""

tcMain.Item(0).Selected = True

Set imgBanner.Picture = frmContenedor.imgBanner_01.Picture

txtMntInicio.Text = "0"
txtMntCorte.Text = "9999999999999"


vGrid.MaxRows = 0
vGrid.MaxCols = 10
vGrid.AppearanceStyle = fxGridStyle

cboTipo.Clear
cboTipo.AddItem "Liquidación"
cboTipo.AddItem "Retiro de Fondos"
cboTipo.Text = "Liquidación"

cboProceso.Clear
cboProceso.AddItem "Desembolsar"
cboProceso.AddItem "Retener"

cboTipoDocumento.Clear
cboTipoDocumento.AddItem fxTipoDocumento("CK")
cboTipoDocumento.AddItem fxTipoDocumento("TE")
cboTipoDocumento.Text = fxTipoDocumento("TE")


cboCuentaFiltro.Clear
cboCuentaFiltro.AddItem "Cuenta Interna"
cboCuentaFiltro.AddItem "Interbancaria"
cboCuentaFiltro.AddItem "Interbancaria Mismo Banco"
cboCuentaFiltro.AddItem "TODOS"
cboCuentaFiltro.Text = "Cuenta Interna"


strSQL = "select rtrim(descripcion) as 'ItmX',cod_operadora as 'IdX' from FND_Operadoras"
Call sbCbo_Llena_New(cboOperadora, strSQL, False, True)


strSQL = "select rtrim(RETENCION_CODIGO) as 'IdX' , RTRIM(DESCRIPCION) as 'ItmX'" _
       & " From FND_RETENCION_CONCEPTOS  Where ACTIVO = 1"
Call sbCbo_Llena_New(cboRetencion, strSQL, False, True)

strSQL = "select rtrim(COD_ESTADO) as 'IdX' , RTRIM(DESCRIPCION) as 'ItmX'" _
       & " From AFI_ESTADOS_PERSONA  Where ACTIVO = 1"
Call sbCbo_Llena_New(cboEstado, strSQL, True, True)


strSQL = "select B.id_banco as 'Idx',B.descripcion as 'ItmX'" _
       & " from tes_banco_asg T inner join Tes_Bancos B on T.id_banco = B.id_banco" _
       & " where T.nombre = '" & glogon.Usuario & "' and B.Estado = 'A'"
Call sbCbo_Llena_New(cboBanco, strSQL, False, True)

strSQL = "select descripcion as 'ItmX',cod_institucion as 'Idx'from instituciones where Activa = 1"
Call sbCbo_Llena_New(cboInstitucion, strSQL, True, True)


cboCreditos.Clear
cboCreditos.AddItem "TODOS"
cboCreditos.Text = "TODOS"
cboCreditos.AddItem "Persona -> Con créditos activos"
cboCreditos.AddItem "Persona -> Con créditos en Mora"
cboCreditos.AddItem "Persona -> Sin créditos activos"
cboCreditos.AddItem "Persona -> Sin créditos en Mora"
cboCreditos.AddItem "Persona -> En Cobro Jud y/o Traspaso"


dtpInicio.Value = fxFechaServidor
dtpCorte.Value = dtpInicio.Value
chkFechas.Value = vbChecked

chkFondosCero.Value = vbUnchecked


vPaso = False

cboProceso.Text = "Desembolsar"
cboTipoDocumento.Text = "Transferencia"

Call cboOperadora_Click
Call cboRetencion_Click
Call cboBanco_Click

txtArchivo.Text = ""

vScroll = False
     FlatScrollBar.Value = 0
vScroll = True

chkLineas_Click
chkMontos_Click

Call Formularios(Me)
Call RefrescaTags(Me)

End Sub


Private Sub Form_Resize()
On Error Resume Next
 
imgBanner.Width = Me.Width

tcMain.Width = Me.Width - 250
tcMain.Height = Me.Height - (tcMain.top + gbResumen.Height + 800)

vGrid.Width = tcMain.Width - 250
vGrid.Height = tcMain.Height - 450

gbResumen.top = tcMain.top + tcMain.Height + 250
gbResumen.Width = tcMain.Width

End Sub

Private Sub imgCalcula_Click()
Dim lng As Long, lngCasos As Long
Dim curAportes As Currency, curRendi As Currency

Me.MousePointer = vbHourglass

curAportes = 0
curRendi = 0
lngCasos = 0

For lng = 1 To vGrid.MaxRows
 vGrid.Row = lng
 vGrid.Col = 1
 If vGrid.Value = vbChecked Then
   lngCasos = lngCasos + 1
   vGrid.Col = 5
   curAportes = curAportes + CCur(vGrid.Text)
   vGrid.Col = 6
   curRendi = curRendi + CCur(vGrid.Text)
  End If
Next lng

txtCasos = Format(lngCasos, "###,###,###,##0")
txtAportes = Format(curAportes, "Standard")
txtRendimientos = Format(curRendi, "Standard")
txtTotal = Format(curAportes + curRendi, "Standard")
  
Me.MousePointer = vbDefault
  
End Sub




Private Sub txtCodigo_Change()

vGrid.MaxRows = 0

End Sub



Private Sub txtCodigo_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtDescripcion.SetFocus

If KeyCode = vbKeyF4 Then
   gBusquedas.Columna = "cod_plan"
   gBusquedas.Orden = "cod_plan"
   gBusquedas.Filtro = "And Cod_operadora=" & cboOperadora.ItemData(cboOperadora.ListIndex)
   gBusquedas.Consulta = "select cod_plan,descripcion from fnd_planes"
   frmBusquedas.Show vbModal
   txtDescripcion.SetFocus
   
   If Trim(gBusquedas.Resultado) <> "" Then
      txtCodigo = Trim(gBusquedas.Resultado)
      txtDescripcion = Trim(gBusquedas.Resultado2)
   End If
   gBusquedas.Resultado = ""
   gBusquedas.Resultado2 = ""
End If

End Sub


Private Sub txtCodigo_LostFocus()
Dim strSQL As String, rs As New ADODB.Recordset

If Trim(txtCodigo) <> "" Then
   strSQL = "Select Descripcion, dateadd(year, 1, getdate()) as 'Vence'" _
          & " from fnd_planes where cod_operadora=" & cboOperadora.ItemData(cboOperadora.ListIndex) _
          & " And cod_plan='" & Trim(txtCodigo) & "'"
   Call OpenRecordSet(rs, strSQL)
        If Not rs.EOF Then
           txtDescripcion.Text = Trim(rs!Descripcion)
           dtpVence.Value = rs!Vence
        Else
           MsgBox "Codigo incorrecto", vbExclamation
           txtCodigo.Text = ""
           txtDescripcion.Text = ""
           txtCodigo.SetFocus
        End If
     rs.Close

Else
  txtDescripcion = ""
End If
End Sub


Private Sub txtDescripcion_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then
   gBusquedas.Columna = "descripcion"
   gBusquedas.Orden = "descripcion"
   gBusquedas.Filtro = "And Cod_operadora=" & cboOperadora.ItemData(cboOperadora.ListIndex)
   gBusquedas.Consulta = "select cod_plan,descripcion from fnd_planes"
   frmBusquedas.Show vbModal
   txtDescripcion.SetFocus
   If Trim(gBusquedas.Resultado) <> "" Then
      txtCodigo = Trim(gBusquedas.Resultado)
      txtDescripcion = Trim(gBusquedas.Resultado2)
   End If
   gBusquedas.Resultado = ""
   gBusquedas.Resultado2 = ""
End If


'If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then cboInstitucion.SetFocus

End Sub



Private Sub txtLinea_Change()
vGrid.MaxRows = 0
End Sub

Private Sub txtLinea_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then txtLineaDesc.SetFocus

If KeyCode = vbKeyF4 Then
  gBusquedas.Resultado = ""
  gBusquedas.Resultado2 = ""
  gBusquedas.Convertir = "N"
  gBusquedas.Resultado = ""
  gBusquedas.Consulta = "select codigo,descripcion from catalogo"
  gBusquedas.Orden = "descripcion"
  gBusquedas.Columna = "descripcion"
  frmBusquedas.Show vbModal
  txtLinea.Text = gBusquedas.Resultado
  txtLineaDesc.Text = gBusquedas.Resultado2
End If

End Sub


Private Sub txtLinea_LostFocus()
 If Len(Trim(txtLinea.Text)) > 0 Then txtLineaDesc.Text = fxDescribeCodigo(Trim(txtLinea.Text))

End Sub

Private Sub txtLineaDesc_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then cboTipoDocumento.SetFocus

If KeyCode = vbKeyF4 Then
  gBusquedas.Resultado = ""
  gBusquedas.Resultado2 = ""
  gBusquedas.Convertir = "N"
  gBusquedas.Resultado = ""
  gBusquedas.Consulta = "select codigo,descripcion from catalogo"
  gBusquedas.Orden = "descripcion"
  gBusquedas.Columna = "descripcion"
  frmBusquedas.Show vbModal
  txtLinea.Text = gBusquedas.Resultado
  txtLineaDesc.Text = gBusquedas.Resultado2
End If
End Sub




Private Sub sbDocumentoMaestro(vTipoDoc As String, ByVal vOperadora As String, vPlan As String, vContrato As Long _
                              , pAportes As Currency, vDocRef As String, vConcepto As String _
                              , vCedula As String, vNombre As String, Optional pRendimiento As Currency = 0)
                              
Dim strSQL As String, rs As New ADODB.Recordset, strLinea(7) As String
Dim lngRecibo As Long, vCuenta As String
        

vAseDocDetalle = ""
txtNotas.Text = fxDepuraString(txtNotas.Text)

strLinea(1) = "Aplica Liquidación General"
strLinea(2) = "Aplicado por..:" & glogon.Usuario
strLinea(3) = "El día        :" & fxFechaServidor



strLinea(4) = ""
strLinea(5) = "Aportes Liq.:" & Format(pAportes, "Standard")
strLinea(6) = "Rendim. Liq.:" & Format(pRendimiento, "Standard")
strLinea(7) = "Total.  Liq.:" & Format(pRendimiento + pAportes, "Standard")



strSQL = "insert SIF_TRANSACCIONES(COD_TRANSACCION,TIPO_DOCUMENTO,REGISTRO_FECHA,REGISTRO_USUARIO,Cliente_IDENTIFICACION,CLIENTE_NOMBRE" _
       & ",cod_concepto,monto,estado,Referencia_01,Referencia_02,cod_oficina" _
       & ",linea1,linea2,linea3,Linea4,linea5,linea6,linea7,detalle,documento)" _
       & " values('" & vDocRef & "','" & vTipoDoc & "',dbo.MyGetdate(),'" & glogon.Usuario & "','" & Trim(vCedula) _
       & "','" & Trim(vNombre) & "','FND006'," & pRendimiento + pAportes & ",'P','" & vOperadora _
       & "','" & vPlan & "','" & GLOBALES.gOficinaTitular & "','" & strLinea(1) & "','" _
       & strLinea(2) & "','" & strLinea(3) & "','" & strLinea(4) & "','" & strLinea(5) & "','" & strLinea(6) & "','" & strLinea(7) & "','" _
       & Mid(txtNotas.Text, 1, 1000) & "','" & vAseDocDeposito & "')"
Call ConectionExecute(strSQL)

End Sub


Private Sub txtMntCorte_GotFocus()
On Error GoTo vError

txtMntCorte.Text = CCur(txtMntCorte.Text)

vError:
End Sub

Private Sub txtMntCorte_LostFocus()
On Error GoTo vError

txtMntCorte.Text = Format(CCur(txtMntCorte.Text), "Standard")

vError:
End Sub

Private Sub txtMntInicio_GotFocus()
On Error GoTo vError

txtMntInicio.Text = CCur(txtMntInicio.Text)

vError:
End Sub

Private Sub txtMntInicio_LostFocus()
On Error GoTo vError

txtMntInicio.Text = Format(CCur(txtMntInicio.Text), "Standard")

vError:
End Sub

Private Sub vGrid_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
Dim lng As Long, lngCasos As Long
Dim curAportes As Currency, curRendi As Currency

If vPaso Then Exit Sub

If Col = 1 Then
   vGrid.Row = Row
   vGrid.Col = 1
   If vGrid.Value = vbChecked Then
        lngCasos = 1
        vGrid.Col = 5
        curAportes = CCur(vGrid.Text)
        vGrid.Col = 6
        curRendi = CCur(vGrid.Text)
   Else
        lngCasos = -1
        vGrid.Col = 5
        curAportes = CCur(vGrid.Text) * -1
        vGrid.Col = 6
        curRendi = CCur(vGrid.Text) * -1
   End If

    txtCasos.Text = Format(CLng(txtCasos.Text) + lngCasos, "###,###,###,##0")
    txtAportes.Text = Format(CCur(txtAportes.Text) + curAportes, "Standard")
    txtRendimientos.Text = Format(CCur(txtRendimientos.Text) + curRendi, "Standard")
    txtTotal.Text = Format(CCur(txtAportes.Text) + CCur(txtRendimientos.Text), "Standard")

End If 'Col = 1

End Sub

Private Sub vGrid_Click(ByVal Col As Long, ByVal Row As Long)
'Dim lng As Long, lngCasos As Long
'Dim curAportes As Currency, curRendi As Currency
'
'
'If Col = 1 Then
'   vGrid.Row = Row
'   vGrid.Col = 1
'   If vGrid.Value = vbChecked Then
'        lngCasos = 1
'        vGrid.Col = 5
'        curAportes = CCur(vGrid.Text)
'        vGrid.Col = 6
'        curRendi = CCur(vGrid.Text)
'   Else
'        lngCasos = -1
'        vGrid.Col = 5
'        curAportes = CCur(vGrid.Text) * -1
'        vGrid.Col = 6
'        curRendi = CCur(vGrid.Text) * -1
'   End If
'
'    txtCasos.Text = Format(CLng(txtCasos.Text) + lngCasos, "###,###,###,##0")
'    txtAportes.Text = Format(CCur(txtAportes.Text) + curAportes, "Standard")
'    txtRendimientos.Text = Format(CCur(txtRendimientos.Text) + curRendi, "Standard")
'    txtTotal.Text = Format(CCur(txtAportes.Text) + CCur(txtRendimientos.Text), "Standard")
'
'End If 'Col = 1

End Sub
