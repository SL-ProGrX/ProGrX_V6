VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.controls.v22.1.0.ocx"
Begin VB.Form frmCR_PolizasSicama 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Control para INS - MAC"
   ClientHeight    =   8085
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   11340
   LinkTopic       =   "Form1"
   ScaleHeight     =   8085
   ScaleWidth      =   11340
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin XtremeSuiteControls.TabControl tcMain 
      Height          =   6972
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   11052
      _Version        =   1441793
      _ExtentX        =   19494
      _ExtentY        =   12298
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
      Item(0).Caption =   "Envío"
      Item(0).Tooltip =   "Generar nuevo Corte para SICAMA"
      Item(0).ControlCount=   11
      Item(0).Control(0)=   "vGrid"
      Item(0).Control(1)=   "Label2(0)"
      Item(0).Control(2)=   "Label2(1)"
      Item(0).Control(3)=   "dtpCorte"
      Item(0).Control(4)=   "Label2(2)"
      Item(0).Control(5)=   "btnCorte"
      Item(0).Control(6)=   "Label2(4)"
      Item(0).Control(7)=   "btnExcel(2)"
      Item(0).Control(8)=   "dtpCorte_Anterior"
      Item(0).Control(9)=   "dtpFactura_Ultima"
      Item(0).Control(10)=   "txtFactura_Ultima"
      Item(1).Caption =   "Recepción"
      Item(1).Tooltip =   "Cargar Archivo de Confirmación del INS"
      Item(1).ControlCount=   13
      Item(1).Control(0)=   "txtArchivo"
      Item(1).Control(1)=   "Label1(2)"
      Item(1).Control(2)=   "Label2(5)"
      Item(1).Control(3)=   "dtpR_Corte"
      Item(1).Control(4)=   "Label2(6)"
      Item(1).Control(5)=   "dtpR_Factura"
      Item(1).Control(6)=   "Label2(7)"
      Item(1).Control(7)=   "txtR_Factura"
      Item(1).Control(8)=   "vGrid_Recepcion"
      Item(1).Control(9)=   "btnRecepcion"
      Item(1).Control(10)=   "btnBuscar"
      Item(1).Control(11)=   "btnCargar"
      Item(1).Control(12)=   "btnInfo"
      Item(2).Caption =   "Consultas"
      Item(2).Tooltip =   "Consulta de Cortes"
      Item(2).ControlCount=   10
      Item(2).Control(0)=   "Label2(8)"
      Item(2).Control(1)=   "btnConsulta(0)"
      Item(2).Control(2)=   "Opt_Consulta(0)"
      Item(2).Control(3)=   "Opt_Consulta(1)"
      Item(2).Control(4)=   "Opt_Consulta(2)"
      Item(2).Control(5)=   "Opt_Consulta(3)"
      Item(2).Control(6)=   "btnExcel(0)"
      Item(2).Control(7)=   "dtpConsulta_Corte"
      Item(2).Control(8)=   "vGrid_Corte"
      Item(2).Control(9)=   "Opt_Consulta(4)"
      Item(3).Caption =   "Consulta de Beneficiarios"
      Item(3).ControlCount=   4
      Item(3).Control(0)=   "vGrid_Beneficiarios"
      Item(3).Control(1)=   "btnConsulta(1)"
      Item(3).Control(2)=   "btnExcel(1)"
      Item(3).Control(3)=   "cboPoliza"
      Begin XtremeSuiteControls.RadioButton Opt_Consulta 
         Height          =   285
         Index           =   0
         Left            =   -67120
         TabIndex        =   20
         Top             =   600
         Visible         =   0   'False
         Width           =   855
         _Version        =   1441793
         _ExtentX        =   1508
         _ExtentY        =   503
         _StockProps     =   79
         Caption         =   "Todo"
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Transparent     =   -1  'True
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         Value           =   -1  'True
      End
      Begin XtremeSuiteControls.DateTimePicker dtpR_Factura 
         Height          =   312
         Left            =   -67360
         TabIndex        =   15
         Top             =   1320
         Visible         =   0   'False
         Width           =   1332
         _Version        =   1441793
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
      Begin XtremeSuiteControls.DateTimePicker dtpCorte_Anterior 
         Height          =   312
         Left            =   1440
         TabIndex        =   11
         Top             =   960
         Width           =   1332
         _Version        =   1441793
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
      Begin XtremeSuiteControls.PushButton btnCorte 
         Height          =   492
         Left            =   8160
         TabIndex        =   9
         Top             =   840
         Width           =   1692
         _Version        =   1441793
         _ExtentX        =   2984
         _ExtentY        =   868
         _StockProps     =   79
         Caption         =   "Genera Corte"
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
         Picture         =   "frmCR_PolizasSicama.frx":0000
         ImageAlignment  =   4
      End
      Begin XtremeSuiteControls.DateTimePicker dtpCorte 
         Height          =   312
         Left            =   1440
         TabIndex        =   5
         Top             =   600
         Width           =   1332
         _Version        =   1441793
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
      Begin XtremeSuiteControls.DateTimePicker dtpFactura_Ultima 
         Height          =   312
         Left            =   4680
         TabIndex        =   7
         Top             =   600
         Width           =   1932
         _Version        =   1441793
         _ExtentX        =   3408
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
      Begin XtremeSuiteControls.DateTimePicker dtpR_Corte 
         Height          =   312
         Left            =   -67360
         TabIndex        =   13
         Top             =   960
         Visible         =   0   'False
         Width           =   1332
         _Version        =   1441793
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
      Begin XtremeSuiteControls.PushButton btnRecepcion 
         Height          =   495
         Left            =   -61600
         TabIndex        =   17
         Top             =   960
         Visible         =   0   'False
         Width           =   2415
         _Version        =   1441793
         _ExtentX        =   4260
         _ExtentY        =   873
         _StockProps     =   79
         Caption         =   "Carga Informe del INS"
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
         Picture         =   "frmCR_PolizasSicama.frx":0719
         ImageAlignment  =   4
      End
      Begin XtremeSuiteControls.PushButton btnConsulta 
         Height          =   492
         Index           =   0
         Left            =   -61960
         TabIndex        =   19
         Top             =   480
         Visible         =   0   'False
         Width           =   1812
         _Version        =   1441793
         _ExtentX        =   3196
         _ExtentY        =   868
         _StockProps     =   79
         Caption         =   "Carga Información"
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
         Picture         =   "frmCR_PolizasSicama.frx":0E39
         ImageAlignment  =   4
      End
      Begin XtremeSuiteControls.RadioButton Opt_Consulta 
         Height          =   285
         Index           =   1
         Left            =   -66160
         TabIndex        =   21
         Top             =   480
         Visible         =   0   'False
         Width           =   1215
         _Version        =   1441793
         _ExtentX        =   2143
         _ExtentY        =   503
         _StockProps     =   79
         Caption         =   "Inclusiones"
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   8.25
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
      Begin XtremeSuiteControls.RadioButton Opt_Consulta 
         Height          =   285
         Index           =   2
         Left            =   -66160
         TabIndex        =   22
         Top             =   840
         Visible         =   0   'False
         Width           =   1215
         _Version        =   1441793
         _ExtentX        =   2138
         _ExtentY        =   508
         _StockProps     =   79
         Caption         =   "Exclusiones"
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   8.25
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
      Begin XtremeSuiteControls.RadioButton Opt_Consulta 
         Height          =   285
         Index           =   3
         Left            =   -64480
         TabIndex        =   23
         Top             =   480
         Visible         =   0   'False
         Width           =   1335
         _Version        =   1441793
         _ExtentX        =   2350
         _ExtentY        =   508
         _StockProps     =   79
         Caption         =   "Modicaciones"
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   8.25
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
      Begin XtremeSuiteControls.PushButton btnExcel 
         Height          =   492
         Index           =   0
         Left            =   -60160
         TabIndex        =   24
         Top             =   480
         Visible         =   0   'False
         Width           =   1092
         _Version        =   1441793
         _ExtentX        =   1926
         _ExtentY        =   868
         _StockProps     =   79
         Caption         =   "Excel"
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
         Picture         =   "frmCR_PolizasSicama.frx":1541
         ImageAlignment  =   4
      End
      Begin XtremeSuiteControls.PushButton btnConsulta 
         Height          =   495
         Index           =   1
         Left            =   -64960
         TabIndex        =   25
         Top             =   480
         Visible         =   0   'False
         Width           =   1815
         _Version        =   1441793
         _ExtentX        =   3196
         _ExtentY        =   868
         _StockProps     =   79
         Caption         =   "Carga Información"
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
         Picture         =   "frmCR_PolizasSicama.frx":1E12
         ImageAlignment  =   4
      End
      Begin XtremeSuiteControls.PushButton btnExcel 
         Height          =   495
         Index           =   1
         Left            =   -63160
         TabIndex        =   27
         Top             =   480
         Visible         =   0   'False
         Width           =   1095
         _Version        =   1441793
         _ExtentX        =   1926
         _ExtentY        =   868
         _StockProps     =   79
         Caption         =   "Excel"
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
         Picture         =   "frmCR_PolizasSicama.frx":251A
         ImageAlignment  =   4
      End
      Begin XtremeSuiteControls.PushButton btnExcel 
         Height          =   492
         Index           =   2
         Left            =   9840
         TabIndex        =   28
         Top             =   840
         Width           =   1092
         _Version        =   1441793
         _ExtentX        =   1926
         _ExtentY        =   868
         _StockProps     =   79
         Caption         =   "Excel"
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
         Picture         =   "frmCR_PolizasSicama.frx":2DEB
         ImageAlignment  =   4
      End
      Begin FPSpreadADO.fpSpread vGrid 
         Height          =   5412
         Left            =   120
         TabIndex        =   3
         Top             =   1440
         Width           =   10812
         _Version        =   524288
         _ExtentX        =   19071
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
         MaxCols         =   41
         MaxRows         =   1000000
         SpreadDesigner  =   "frmCR_PolizasSicama.frx":36BC
         VScrollSpecial  =   -1  'True
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin FPSpreadADO.fpSpread vGrid_Recepcion 
         Height          =   5052
         Left            =   -69880
         TabIndex        =   26
         Top             =   1680
         Visible         =   0   'False
         Width           =   10812
         _Version        =   524288
         _ExtentX        =   19071
         _ExtentY        =   8911
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
         MaxCols         =   38
         MaxRows         =   1000000
         SpreadDesigner  =   "frmCR_PolizasSicama.frx":4943
         VScrollSpecial  =   -1  'True
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin XtremeSuiteControls.PushButton btnBuscar 
         Height          =   372
         Left            =   -61600
         TabIndex        =   29
         Top             =   480
         Visible         =   0   'False
         Width           =   492
         _Version        =   1441793
         _ExtentX        =   868
         _ExtentY        =   656
         _StockProps     =   79
         BackColor       =   -2147483633
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         Picture         =   "frmCR_PolizasSicama.frx":5AD4
      End
      Begin XtremeSuiteControls.PushButton btnCargar 
         Height          =   372
         Left            =   -61120
         TabIndex        =   30
         Top             =   480
         Visible         =   0   'False
         Width           =   492
         _Version        =   1441793
         _ExtentX        =   868
         _ExtentY        =   656
         _StockProps     =   79
         BackColor       =   -2147483633
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         Picture         =   "frmCR_PolizasSicama.frx":61D4
      End
      Begin XtremeSuiteControls.PushButton btnInfo 
         Height          =   372
         Left            =   -60640
         TabIndex        =   31
         Top             =   480
         Visible         =   0   'False
         Width           =   492
         _Version        =   1441793
         _ExtentX        =   868
         _ExtentY        =   656
         _StockProps     =   79
         BackColor       =   -2147483633
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         Picture         =   "frmCR_PolizasSicama.frx":68ED
      End
      Begin XtremeSuiteControls.FlatEdit txtArchivo 
         Height          =   372
         Left            =   -68560
         TabIndex        =   32
         Top             =   480
         Visible         =   0   'False
         Width           =   6852
         _Version        =   1441793
         _ExtentX        =   12086
         _ExtentY        =   656
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
      Begin XtremeSuiteControls.FlatEdit txtFactura_Ultima 
         Height          =   312
         Left            =   4680
         TabIndex        =   33
         Top             =   960
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
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtR_Factura 
         Height          =   312
         Left            =   -64000
         TabIndex        =   34
         Top             =   1080
         Visible         =   0   'False
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
      Begin XtremeSuiteControls.DateTimePicker dtpConsulta_Corte 
         Height          =   312
         Left            =   -68800
         TabIndex        =   35
         Top             =   600
         Visible         =   0   'False
         Width           =   1332
         _Version        =   1441793
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
      Begin FPSpreadADO.fpSpread vGrid_Corte 
         Height          =   5532
         Left            =   -69880
         TabIndex        =   36
         Top             =   1200
         Visible         =   0   'False
         Width           =   10812
         _Version        =   524288
         _ExtentX        =   19071
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
         MaxCols         =   41
         MaxRows         =   1000000
         SpreadDesigner  =   "frmCR_PolizasSicama.frx":7006
         VScrollSpecial  =   -1  'True
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin XtremeSuiteControls.RadioButton Opt_Consulta 
         Height          =   285
         Index           =   4
         Left            =   -64480
         TabIndex        =   37
         Top             =   840
         Visible         =   0   'False
         Width           =   1335
         _Version        =   1441793
         _ExtentX        =   2350
         _ExtentY        =   508
         _StockProps     =   79
         Caption         =   "Sin Cambios"
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   8.25
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
      Begin XtremeSuiteControls.ComboBox cboPoliza 
         Height          =   465
         Left            =   -69880
         TabIndex        =   38
         Top             =   480
         Visible         =   0   'False
         Width           =   4815
         _Version        =   1441793
         _ExtentX        =   8493
         _ExtentY        =   820
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Style           =   2
         Appearance      =   6
         FlatStyle       =   -1  'True
         UseVisualStyle  =   0   'False
         Text            =   "ComboBox1"
      End
      Begin FPSpreadADO.fpSpread vGrid_Beneficiarios 
         Height          =   5775
         Left            =   -70000
         TabIndex        =   39
         Top             =   1080
         Visible         =   0   'False
         Width           =   10815
         _Version        =   524288
         _ExtentX        =   19076
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
         MaxCols         =   32
         MaxRows         =   1000000
         SpreadDesigner  =   "frmCR_PolizasSicama.frx":828D
         VScrollSpecial  =   -1  'True
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Corte:"
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
         Index           =   8
         Left            =   -69640
         TabIndex        =   18
         Top             =   600
         Visible         =   0   'False
         Width           =   852
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "No. Factura:"
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
         Index           =   7
         Left            =   -65440
         TabIndex        =   16
         Top             =   1080
         Visible         =   0   'False
         Width           =   1332
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Factura:"
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
         Left            =   -68560
         TabIndex        =   14
         Top             =   1320
         Visible         =   0   'False
         Width           =   1092
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Ultimo Corte"
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
         Left            =   -68560
         TabIndex        =   12
         Top             =   960
         Visible         =   0   'False
         Width           =   1092
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Anterior:"
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
         TabIndex        =   10
         Top             =   960
         Width           =   1092
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Factura No.:"
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
         Left            =   3120
         TabIndex        =   8
         Top             =   960
         Width           =   1212
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Factura Fecha: "
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
         Left            =   3120
         TabIndex        =   6
         Top             =   600
         Width           =   1332
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Corte:"
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
         Left            =   480
         TabIndex        =   4
         Top             =   600
         Width           =   852
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Archivo"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   2
         Left            =   -69640
         TabIndex        =   2
         Top             =   480
         Visible         =   0   'False
         Width           =   1092
      End
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Control para INS - MAC"
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
      Index           =   0
      Left            =   1440
      TabIndex        =   1
      Top             =   240
      Width           =   6012
   End
   Begin VB.Image imgBanner 
      Height          =   852
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   11532
   End
End
Attribute VB_Name = "frmCR_PolizasSicama"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnBuscar_Click()
With frmContenedor.CD
        .InitDir = "C:\"
        .DialogTitle = "Localice Archivo [Microsoft EXCEL]..."
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

Private Sub btnCargar_Click()
Call sbArchivoCarga
End Sub

Private Sub btnConsulta_Click(Index As Integer)
Dim vTipo As String, strSQL As String

Me.MousePointer = vbHourglass

Select Case Index
  Case 0 'Consulta Corte
     Select Case True
       Case Opt_Consulta.Item(0).Value
          vTipo = "T"
       Case Opt_Consulta.Item(1).Value
          vTipo = "I"
       Case Opt_Consulta.Item(2).Value
          vTipo = "E"
       Case Opt_Consulta.Item(3).Value
          vTipo = "M"
       Case Opt_Consulta.Item(4).Value
          vTipo = "SC"
     End Select
    
    strSQL = "exec spPoliza_Sicama '', '" & Format(dtpConsulta_Corte.Value, "yyyy/MM/dd") & "',1,'" & glogon.Usuario & "','" & vTipo & "'"
    Call sbCargaGrid(vGrid_Corte, 41, strSQL, True)

  Case 1 'Consulta de Beneficiarios
    Call sbConsulta_Beneficiarios
End Select

Me.MousePointer = vbDefault

End Sub

Private Sub sbConsulta_Beneficiarios()
Dim strSQL As String

On Error GoTo vError

Me.MousePointer = vbDefault

strSQL = "exec spPoliza_Beneficiarios_Lista '" & cboPoliza.ItemData(cboPoliza.ListIndex) & "'"
Call sbCargaGrid(vGrid_Beneficiarios, vGrid_Beneficiarios.MaxCols, strSQL, True)

Exit Sub

vError:
  Me.MousePointer = vbHourglass
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
  

End Sub



Private Sub sbCorte_Consulta(pGrid As Object, pCorte As String, Optional pTipoMov As String = "T")
Dim strSQL As String

On Error GoTo vError

Me.MousePointer = vbDefault

strSQL = "exec spPolizas_Sicama_Consulta '" & pCorte & "','" & pTipoMov & "'"
Call sbCargaGrid(pGrid, pGrid.MaxCols, strSQL, True)

Exit Sub

vError:
  Me.MousePointer = vbHourglass
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub sbCorte_Genera()
Dim strSQL As String

On Error GoTo vError

Me.MousePointer = vbDefault

strSQL = "exec spPolizas_Sicama_Genera '" & Format(dtpCorte.Value, "yyyy/mm/dd") & "','" & glogon.Usuario & "'"
Call ConectionExecute(strSQL)

Call sbCorte_Consulta(vGrid, Format(dtpCorte.Value, "yyyy/mm/dd"), "T")

Exit Sub

vError:
  Me.MousePointer = vbHourglass
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub btnCorte_Click()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

Me.MousePointer = vbHourglass



'spPoliza_Sicama(@Poliza varchar(10), @Corte datetime, @Beneficiarios smallint = 1, @Usuario varchar(30)= '')

strSQL = "exec spPoliza_Sicama '', '" & Format(dtpCorte.Value, "yyyy/MM/dd") & "',1,'" & glogon.Usuario & "','T'"
Call sbCargaGrid(vGrid, 41, strSQL, True)

Me.MousePointer = vbDefault
Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub btnExcel_Click(Index As Integer)
Dim vHeaders As vGridHeaders, vFecha As Date, vTipo As String

'Default para Cortes
vHeaders.Columnas = vGrid.MaxCols
vHeaders.Headers(1) = "Corte"
vHeaders.Headers(2) = "Tipo Id"
vHeaders.Headers(3) = "Identificación"
vHeaders.Headers(4) = "Primer Apellido"
vHeaders.Headers(5) = "Segundo Apellido"
vHeaders.Headers(6) = "Primer Nombre"
vHeaders.Headers(7) = "Segundo Nombre"
vHeaders.Headers(8) = "Genero"
vHeaders.Headers(9) = "Fecha Nacimiento"
vHeaders.Headers(10) = "Correo Electrónico"
vHeaders.Headers(11) = "Nacionalidad"
vHeaders.Headers(12) = "Provincia"
vHeaders.Headers(13) = "Canton"
vHeaders.Headers(14) = "Distrito"
vHeaders.Headers(15) = "Direccion Completa"
vHeaders.Headers(16) = "Tipo Telefono"
vHeaders.Headers(17) = "Numero Telefono"
vHeaders.Headers(18) = "Suma Asegurada 1"
vHeaders.Headers(19) = "Suma Asegurada 2"
vHeaders.Headers(20) = "Prima Recaudada"
vHeaders.Headers(21) = "Num.Póliza"
vHeaders.Headers(22) = "Num.Referencia"

vHeaders.Headers(23) = "Bene No.1 Tipo ( ID )"
vHeaders.Headers(24) = "Bene No.1 Identificación"
vHeaders.Headers(25) = "Bene No.1 Nombre Completo"
vHeaders.Headers(26) = "Bene No.1 Parentesco"
vHeaders.Headers(27) = "Bene No.1 Porcentaje"
vHeaders.Headers(28) = "Bene No.2 Tipo ( ID )"
vHeaders.Headers(29) = "Bene No.2 Identificación"
vHeaders.Headers(30) = "Bene No.2 Nombre Completo"
vHeaders.Headers(31) = "Bene No.2 Parentesco"
vHeaders.Headers(32) = "Bene No.2 Porcentaje"
vHeaders.Headers(33) = "Bene No.3 Tipo ( ID )"
vHeaders.Headers(34) = "Bene No.3 Identificación"
vHeaders.Headers(35) = "Bene No.3 Nombre Completo"
vHeaders.Headers(36) = "Bene No.3 Parentesco"
vHeaders.Headers(37) = "Bene No.3 Porcentaje"
vHeaders.Headers(38) = "Porcentaje de Recargo"
'vHeaders.Headers(39) = "Movimiento"
'vHeaders.Headers(40) = "Nacionalidad Id"
'vHeaders.Headers(41) = "Nacionalidad Alterno"
    

Select Case Index
  Case 0 'Consulta de Corte
     vHeaders.Columnas = vGrid_Corte.MaxCols
     Select Case True
       Case Opt_Consulta.Item(0).Value
          vTipo = "TODO"
       Case Opt_Consulta.Item(1).Value
          vTipo = "INCLUSIONES"
       Case Opt_Consulta.Item(2).Value
          vTipo = "EXCLUSIONES"
       Case Opt_Consulta.Item(3).Value
          vTipo = "MODIFICACIONES"
       Case Opt_Consulta.Item(4).Value
          vTipo = "SIN_CAMBIOS"
     End Select
     
    Call sbSIFGridExportar(vGrid_Corte, vHeaders, "INS-MAC_Corte_" & Format(dtpConsulta_Corte.Value, "yyyy-mm-dd") & "_" & vTipo)
  
  Case 1 'Beneficiarios
    vFecha = fxFechaServidor
    
    vHeaders.Columnas = vGrid_Beneficiarios.MaxCols
    vHeaders.Headers(1) = "Identificación"
    vHeaders.Headers(2) = "Nombre"
    vHeaders.Headers(3) = "Bene No.1 Tipo ( ID )"
    vHeaders.Headers(4) = "Bene No.1 Identificación"
    vHeaders.Headers(5) = "Bene No.1 Nombre Completo"
    vHeaders.Headers(6) = "Bene No.1 Parentesco"
    vHeaders.Headers(7) = "Bene No.1 Porcentaje"
    
    vHeaders.Headers(8) = "Bene No.2 Tipo ( ID )"
    vHeaders.Headers(9) = "Bene No.2 Identificación"
    vHeaders.Headers(10) = "Bene No.2 Nombre Completo"
    vHeaders.Headers(11) = "Bene No.2 Parentesco"
    vHeaders.Headers(12) = "Bene No.2 Porcentaje"
    
    vHeaders.Headers(13) = "Bene No.3 Tipo ( ID )"
    vHeaders.Headers(14) = "Bene No.3 Identificación"
    vHeaders.Headers(15) = "Bene No.3 Nombre Completo"
    vHeaders.Headers(16) = "Bene No.3 Parentesco"
    vHeaders.Headers(17) = "Bene No.3 Porcentaje"
    
    vHeaders.Headers(18) = "Bene No.4 Tipo ( ID )"
    vHeaders.Headers(19) = "Bene No.4 Identificación"
    vHeaders.Headers(20) = "Bene No.4 Nombre Completo"
    vHeaders.Headers(21) = "Bene No.4 Parentesco"
    vHeaders.Headers(22) = "Bene No.4 Porcentaje"
    
    vHeaders.Headers(23) = "Bene No.5 Tipo ( ID )"
    vHeaders.Headers(24) = "Bene No.5 Identificación"
    vHeaders.Headers(25) = "Bene No.5 Nombre Completo"
    vHeaders.Headers(26) = "Bene No.5 Parentesco"
    vHeaders.Headers(27) = "Bene No.5 Porcentaje"
    
    vHeaders.Headers(28) = "Bene No.6 Tipo ( ID )"
    vHeaders.Headers(29) = "Bene No.6 Identificación"
    vHeaders.Headers(30) = "Bene No.6 Nombre Completo"
    vHeaders.Headers(31) = "Bene No.6 Parentesco"
    vHeaders.Headers(32) = "Bene No.6 Porcentaje"
    
    Call sbSIFGridExportar(vGrid_Beneficiarios, vHeaders, "Poliza_PC_Beneficiarios_" & Format(vFecha, "yyyy-mm-dd"))
  
  Case 2 'Corte Generado
    vHeaders.Columnas = vGrid.MaxCols
    vHeaders.Headers(39) = "Movimiento"
    vHeaders.Headers(40) = "Nac_Descripción"
    vHeaders.Headers(41) = "Nac_Codigo Alterno"
    
    Call sbSIFGridExportar(vGrid, vHeaders, "INS-MAC_Corte_" & Format(dtpCorte, "yyyy-mm-dd"))
    
End Select

End Sub

Private Sub btnInfo_Click()
'  MsgBox "Archivo de Carga: Microsoft Excel" & vbCrLf _
'        & " - Columnas: CEDULA, NOMBRE, MONTO, TASA, PLAZO, COMISION" & vbCrLf _
'        & " - Nombre de la Hoja: IMPORT" _
'    , vbInformation, "Información del Archivo de Carga"
End Sub

Private Sub Form_Load()
Dim strSQL As String

vModulo = 3

vGrid.AppearanceStyle = fxGridStyle
Set imgBanner.Picture = frmContenedor.imgBanner_01.Picture


dtpCorte.Value = fxFechaServidor
dtpR_Corte.Value = dtpCorte.Value
dtpR_Factura.Value = dtpCorte.Value
dtpFactura_Ultima.Value = dtpCorte.Value
dtpConsulta_Corte.Value = dtpCorte.Value


 strSQL = "select COD_POLIZA as 'Idx', rtrim(Poliza_Desc) as 'ItmX' from vPoliza_Catalogo" _
        & " Where Tipo = 'PC'" _
        & " order by COD_POLIZA"
 Call sbCbo_Llena_New(cboPoliza, strSQL, False, True)

tcMain.Item(0).Selected = True
txtArchivo.Text = ""
vGrid.MaxRows = 0
vGrid_Beneficiarios.MaxRows = 0
vGrid_Recepcion.MaxRows = 0
vGrid_Corte.MaxRows = 0

Call Formularios(Me)
Call RefrescaTags(Me)

End Sub

Private Sub Form_Resize()
On Error Resume Next


imgBanner.Width = Me.Width
tcMain.Width = Me.Width - 350
tcMain.Height = Me.Height - (tcMain.Top + 680)
vGrid.Width = tcMain.Width - 250
vGrid_Beneficiarios.Width = vGrid.Width
vGrid_Recepcion.Width = vGrid.Width
vGrid_Corte.Width = vGrid.Width

vGrid.Height = tcMain.Height - (vGrid.Top + 250)
vGrid_Beneficiarios.Height = tcMain.Height - (vGrid_Beneficiarios.Top + 250)
vGrid_Recepcion.Height = tcMain.Height - (vGrid_Recepcion.Top + 250)
vGrid_Corte.Height = tcMain.Height - (vGrid_Corte.Top + 250)

End Sub

Private Sub Opt_Consulta_Click(Index As Integer)

vGrid_Corte.MaxRows = 0

End Sub

Private Sub tlbX_ButtonClick(ByVal Button As MSComctlLib.Button)
        
Select Case Button.Key
  
  Case "buscar"
    With frmContenedor.CD
        .InitDir = "C:\"
        .DialogTitle = "Localice Archivo [Microsoft EXCEL]..."
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
    
  Case "cargar"
      Call sbArchivoCarga

End Select
End Sub



Private Sub sbArchivoCarga()
Dim strSQL As String, rs As New ADODB.Recordset

Dim pCedula As String, pNombre As String, pFondos As Currency
Dim pOperadora As Integer, pPlan As String, pInstitucion As Long, pLinea As Long

Dim strCadena As String, curMonto As Currency
Dim fn As Long, lCasos As Long
Dim strMonto  As String
Dim strCedula As String
Dim strNombre As String
Dim i As Integer, vCampos As Boolean



On Error GoTo vError


vGrid.MaxRows = 0

If txtArchivo.Text = "" Then
   MsgBox "Seleccione un archivo a procesar...", vbExclamation
   Exit Sub
End If


'If fxAplicada Then
'   MsgBox "Ya se aplico una planilla con esta fecha de proceso para la institución y el plan elegidos"
'   Exit Sub
'End If


Me.MousePointer = vbHourglass


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
         "Los campos son Identificacion... ¦ Nombre de la Hoja = IMPORT"
   Exit Sub
End If

vCampos = False
For i = 0 To rs.Fields.Count
     
    If UCase(LCase(rs.Fields(i).Name)) = "NOMBRE" Then
       vCampos = True
    End If
     
     If vCampos Then Exit For
Next i

If Not vCampos Then
   MsgBox "No coincide la estructura del archivo a cargar..." & vbCrLf & _
         "Los campos sonIdentificacion... ¦ Nombre de la Hoja = IMPORT"
   Exit Sub
End If


'vCampos = False
'For i = 0 To rs.Fields.Count
'
'    If UCase(LCase(rs.Fields(i).Name)) = "FONDOS" Then
'       vCampos = True
'    End If
'
'     If vCampos Then Exit For
'Next i
'
'If Not vCampos Then
'   MsgBox "No coincide la estructura del archivo a cargar..." & vbCrLf & _
'         "Los campos son Identificacion... ¦ Nombre de la Hoja = IMPORT"
'   Exit Sub
'End If

'FIN: Validación del Archivo



'Sube, Revisa y Carga
With vGrid_Recepcion
    
    pLinea = 0
    strSQL = ""
    
'    Do While Not rs.EOF
'      If Trim(rs!Cedula) <> "" Then
'        pCedula = rs!Cedula
'        pNombre = rs!Nombre
'        pFondos = rs!fondos
'        pLinea = pLinea + 1
'
'        If pLinea = 1 Then
'            strSQL = strSQL & Space(10) & "exec spFndPlanillaDirecta_Sube " & pInstitucion & "," & pOperadora & ",'" & pPlan & "','" _
'                   & txtComprobante.Text & "'," & cboProceso.Text & ",'" & pCedula & "','" & pNombre & "'," _
'                   & pFondos & "," & pLinea & "," & 1
'        Else
'            strSQL = strSQL & Space(10) & "exec spFndPlanillaDirecta_Sube " & pInstitucion & "," & pOperadora & ",'" & pPlan & "','" _
'                   & txtComprobante.Text & "'," & cboProceso.Text & ",'" & pCedula & "','" & pNombre & "'," _
'                   & pFondos & "," & pLinea & "," & 0
'        End If
'
'        If Len(strSQL) > 20000 Then
'           Call ConectionExecute(strSQL)
'           If glogon.error Then
'              Exit Sub
'           End If
'           strSQL = ""
'        End If
'
'      End If
'      rs.MoveNext
'    Loop
'    rs.Close
'
''Procesa Ultimo Bloque
'
'If Len(strSQL) > 0 Then
'   Call ConectionExecute(strSQL)
'   If glogon.error Then
'      Exit Sub
'   End If
'   strSQL = ""
'End If
'
''Revisa Lote y lo Carga
'strSQL = "exec spFndPlanillaDirecta_Consulta " & pOperadora & ",'" & pPlan & "','" _
'                   & txtComprobante.Text & "',1"
'Call OpenRecordSet(rs, strSQL)
'If glogon.error Then
'   Exit Sub
'End If
'
'    Do While Not rs.EOF
'            pCedula = rs!Cedula
'            pNombre = rs!Nombre
'            pFondos = rs!fondos
'
'
'            .MaxRows = .MaxRows + 1
'            .Row = .MaxRows
'            .Col = 1
'            .Text = rs!Cedula
'            .Col = 2
'            .Text = rs!Nombre
'            .Col = 3
'            .Value = IIf((rs!Existe_Persona = 1), 0, 1)
'
'            .Col = 4
'            .Value = IIf((rs!Existe_Contrato = 1), 0, 1)
'            .CellTag = rs!cod_contrato
'
'            .Col = 5
'            .Text = Format(rs!fondos, "Standard")
'
'            If rs!Existe_Persona = 0 Then
'               txtSocios.Text = CInt(txtSocios.Text) + 1
'            End If
'
'            If rs!Existe_Contrato = 0 Then
'               txtContratos.Text = CInt(txtContratos.Text) + 1
'            End If
'
'            curMonto = curMonto + rs!fondos
'            txtMonto.Text = Format(curMonto, "Standard")
'            txtCasos = txtCasos + 1
'            txtCasos.Refresh
'
'      rs.MoveNext
'    Loop
'    rs.Close


End With 'vGrid


Exit Sub

vError:
    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub


