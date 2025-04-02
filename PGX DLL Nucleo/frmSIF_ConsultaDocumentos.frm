VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.controls.v22.1.0.ocx"
Begin VB.Form frmSIF_ConsultaDocumentos 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Consulta de Transacciones (Documentos)"
   ClientHeight    =   9045
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   16935
   Icon            =   "frmSIF_ConsultaDocumentos.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9045
   ScaleWidth      =   16935
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin XtremeSuiteControls.CheckBox chkConceptos 
      Height          =   210
      Left            =   2880
      TabIndex        =   128
      Top             =   3960
      Width           =   210
      _Version        =   1441793
      _ExtentX        =   370
      _ExtentY        =   370
      _StockProps     =   79
      Transparent     =   -1  'True
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Value           =   1
   End
   Begin VB.Frame fraDocumento 
      Caption         =   "Buscar Tipo de Documento"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3975
      Left            =   3600
      TabIndex        =   15
      Top             =   5040
      Visible         =   0   'False
      Width           =   10695
      Begin VB.TextBox txtDocNameConsulta 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   3120
         TabIndex        =   18
         Top             =   480
         Width           =   6015
      End
      Begin VB.TextBox txtDocCodConsulta 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   1800
         TabIndex        =   17
         ToolTipText     =   "Código del Tipo de Documento"
         Top             =   480
         Width           =   1335
      End
      Begin MSComctlLib.ListView lswDocConsulta 
         Height          =   2895
         Left            =   1800
         TabIndex        =   16
         Top             =   960
         Width           =   8655
         _ExtentX        =   15266
         _ExtentY        =   5106
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         HideColumnHeaders=   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Object.Width           =   7832
         EndProperty
      End
      Begin XtremeSuiteControls.PushButton btnDocConsultaCerrar 
         Height          =   495
         Left            =   9960
         TabIndex        =   76
         Top             =   0
         Width           =   735
         _Version        =   1441793
         _ExtentX        =   1291
         _ExtentY        =   868
         _StockProps     =   79
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
         FlatStyle       =   -1  'True
         Appearance      =   16
         Picture         =   "frmSIF_ConsultaDocumentos.frx":6852
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Documento"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   26
         Left            =   240
         TabIndex        =   20
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Seleccione el tipo de documento que desea filtrar..:"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   1395
         Index           =   25
         Left            =   240
         TabIndex        =   19
         Top             =   960
         Width           =   1335
      End
   End
   Begin XtremeSuiteControls.TabControl tcMain 
      Height          =   3612
      Left            =   3600
      TabIndex        =   79
      Top             =   5160
      Width           =   13452
      _Version        =   1441793
      _ExtentX        =   23728
      _ExtentY        =   6371
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
      Item(0).Caption =   "Documento"
      Item(0).ControlCount=   24
      Item(0).Control(0)=   "txtCedula"
      Item(0).Control(1)=   "Label1(15)"
      Item(0).Control(2)=   "Label1(13)"
      Item(0).Control(3)=   "Label1(11)"
      Item(0).Control(4)=   "Label1(10)"
      Item(0).Control(5)=   "Label1(9)"
      Item(0).Control(6)=   "Label1(8)"
      Item(0).Control(7)=   "txtCod_Concepto"
      Item(0).Control(8)=   "txtDocumento"
      Item(0).Control(9)=   "txtMonto"
      Item(0).Control(10)=   "txtCaja"
      Item(0).Control(11)=   "txtOficina"
      Item(0).Control(12)=   "Label1(42)"
      Item(0).Control(13)=   "txtUsuarioCaja"
      Item(0).Control(14)=   "txtNombre"
      Item(0).Control(15)=   "txtDescConcepto"
      Item(0).Control(16)=   "txtEstado"
      Item(0).Control(17)=   "Label1(43)"
      Item(0).Control(18)=   "txtFecha"
      Item(0).Control(19)=   "Label1(44)"
      Item(0).Control(20)=   "Label1(12)"
      Item(0).Control(21)=   "chkBloqueado"
      Item(0).Control(22)=   "txtDetalle"
      Item(0).Control(23)=   "btnDocFix"
      Item(1).Caption =   "Asiento"
      Item(1).ControlCount=   6
      Item(1).Control(0)=   "txtDebito"
      Item(1).Control(1)=   "txtCredito"
      Item(1).Control(2)=   "txtDiferencia"
      Item(1).Control(3)=   "vgridAsiento"
      Item(1).Control(4)=   "Label3(0)"
      Item(1).Control(5)=   "Label3(1)"
      Item(2).Caption =   "Afectaciones"
      Item(2).ControlCount=   1
      Item(2).Control(0)=   "vGridAfectaciones"
      Item(3).Caption =   "Seguimiento"
      Item(3).ControlCount=   11
      Item(3).Control(0)=   "txtAnulaFecha"
      Item(3).Control(1)=   "txtTrasladoFecha"
      Item(3).Control(2)=   "txtRegistroFecha"
      Item(3).Control(3)=   "txtAnulaUsuario"
      Item(3).Control(4)=   "txtTrasladoUsuario"
      Item(3).Control(5)=   "txtRegistroUsuario"
      Item(3).Control(6)=   "Label8"
      Item(3).Control(7)=   "Label7"
      Item(3).Control(8)=   "Label6"
      Item(3).Control(9)=   "Label5"
      Item(3).Control(10)=   "Label4"
      Item(4).Caption =   "Valores Registrados"
      Item(4).ControlCount=   1
      Item(4).Control(0)=   "vGridFormasPago"
      Begin XtremeSuiteControls.CheckBox chkBloqueado 
         Height          =   252
         Left            =   2280
         TabIndex        =   116
         Top             =   3000
         Width           =   5892
         _Version        =   1441793
         _ExtentX        =   10393
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Este documento se encuentra bloqueado para Traslado a la Contabilidad?"
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
         Transparent     =   -1  'True
         UseVisualStyle  =   -1  'True
         Appearance      =   16
         Value           =   1
         Alignment       =   1
      End
      Begin FPSpreadADO.fpSpread vgridAsiento 
         Height          =   2412
         Left            =   -69880
         TabIndex        =   80
         Top             =   480
         Visible         =   0   'False
         Width           =   10008
         _Version        =   524288
         _ExtentX        =   17653
         _ExtentY        =   4255
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
         MaxCols         =   12
         SpreadDesigner  =   "frmSIF_ConsultaDocumentos.frx":701F
         VScrollSpecial  =   -1  'True
         AppearanceStyle =   1
      End
      Begin FPSpreadADO.fpSpread vGridAfectaciones 
         Height          =   2892
         Left            =   -70000
         TabIndex        =   83
         Top             =   480
         Visible         =   0   'False
         Width           =   10332
         _Version        =   524288
         _ExtentX        =   18225
         _ExtentY        =   5101
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
         MaxCols         =   12
         SpreadDesigner  =   "frmSIF_ConsultaDocumentos.frx":7965
         VScrollSpecial  =   -1  'True
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin FPSpreadADO.fpSpread vGridFormasPago 
         Height          =   2892
         Left            =   -69880
         TabIndex        =   89
         Top             =   480
         Visible         =   0   'False
         Width           =   10008
         _Version        =   524288
         _ExtentX        =   17653
         _ExtentY        =   5101
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
         SpreadDesigner  =   "frmSIF_ConsultaDocumentos.frx":8720
         AppearanceStyle =   1
      End
      Begin XtremeSuiteControls.FlatEdit txtMonto 
         Height          =   312
         Left            =   1200
         TabIndex        =   99
         Top             =   1800
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
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Locked          =   -1  'True
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtCedula 
         Height          =   312
         Left            =   1200
         TabIndex        =   90
         Top             =   600
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
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   2
         Locked          =   -1  'True
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtCod_Concepto 
         Height          =   312
         Left            =   1200
         TabIndex        =   97
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
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   2
         Locked          =   -1  'True
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtDocumento 
         Height          =   312
         Left            =   1200
         TabIndex        =   98
         Top             =   1440
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
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   2
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtCaja 
         Height          =   312
         Left            =   1200
         TabIndex        =   100
         Top             =   2160
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
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   2
         Locked          =   -1  'True
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtUsuarioCaja 
         Height          =   312
         Left            =   6240
         TabIndex        =   103
         Top             =   2160
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
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   2
         Locked          =   -1  'True
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtNombre 
         Height          =   312
         Left            =   3120
         TabIndex        =   104
         Top             =   600
         Width           =   5052
         _Version        =   1441793
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
      Begin XtremeSuiteControls.FlatEdit txtDescConcepto 
         Height          =   312
         Left            =   3120
         TabIndex        =   105
         Top             =   960
         Width           =   5052
         _Version        =   1441793
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
      Begin XtremeSuiteControls.FlatEdit txtEstado 
         Height          =   312
         Left            =   6240
         TabIndex        =   106
         Top             =   1800
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
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   2
         Locked          =   -1  'True
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtFecha 
         Height          =   312
         Left            =   6240
         TabIndex        =   108
         Top             =   1440
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
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   2
         Locked          =   -1  'True
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtOficina 
         Height          =   312
         Left            =   1200
         TabIndex        =   101
         Top             =   2640
         Width           =   6972
         _Version        =   1441793
         _ExtentX        =   12298
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
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtDetalle 
         Height          =   2712
         Left            =   8280
         TabIndex        =   111
         Top             =   600
         Width           =   5292
         _Version        =   1441793
         _ExtentX        =   9334
         _ExtentY        =   4784
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
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.PushButton btnDocFix 
         Height          =   372
         Left            =   3240
         TabIndex        =   112
         Top             =   1400
         Width           =   492
         _Version        =   1441793
         _ExtentX        =   868
         _ExtentY        =   656
         _StockProps     =   79
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
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         Picture         =   "frmSIF_ConsultaDocumentos.frx":8E29
      End
      Begin XtremeSuiteControls.FlatEdit txtRegistroUsuario 
         Height          =   312
         Left            =   -65800
         TabIndex        =   117
         Top             =   1080
         Visible         =   0   'False
         Width           =   2172
         _Version        =   1441793
         _ExtentX        =   3831
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
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtRegistroFecha 
         Height          =   312
         Left            =   -63640
         TabIndex        =   118
         Top             =   1080
         Visible         =   0   'False
         Width           =   2172
         _Version        =   1441793
         _ExtentX        =   3831
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
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtTrasladoUsuario 
         Height          =   312
         Left            =   -65800
         TabIndex        =   119
         Top             =   1440
         Visible         =   0   'False
         Width           =   2172
         _Version        =   1441793
         _ExtentX        =   3831
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
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtTrasladoFecha 
         Height          =   312
         Left            =   -63640
         TabIndex        =   120
         Top             =   1440
         Visible         =   0   'False
         Width           =   2172
         _Version        =   1441793
         _ExtentX        =   3831
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
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtAnulaUsuario 
         Height          =   312
         Left            =   -65800
         TabIndex        =   121
         Top             =   1800
         Visible         =   0   'False
         Width           =   2172
         _Version        =   1441793
         _ExtentX        =   3831
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
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtAnulaFecha 
         Height          =   312
         Left            =   -63640
         TabIndex        =   122
         Top             =   1800
         Visible         =   0   'False
         Width           =   2172
         _Version        =   1441793
         _ExtentX        =   3831
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
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtDebito 
         Height          =   312
         Left            =   -63040
         TabIndex        =   123
         Top             =   3000
         Visible         =   0   'False
         Width           =   1572
         _Version        =   1441793
         _ExtentX        =   2773
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
         Locked          =   -1  'True
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtCredito 
         Height          =   312
         Left            =   -61480
         TabIndex        =   124
         Top             =   3000
         Visible         =   0   'False
         Width           =   1572
         _Version        =   1441793
         _ExtentX        =   2773
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
         Locked          =   -1  'True
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtDiferencia 
         Height          =   312
         Left            =   -68320
         TabIndex        =   125
         Top             =   3000
         Visible         =   0   'False
         Width           =   1812
         _Version        =   1441793
         _ExtentX        =   3196
         _ExtentY        =   550
         _StockProps     =   77
         ForeColor       =   128
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
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Oficina"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   192
         Index           =   12
         Left            =   240
         TabIndex        =   110
         Top             =   2640
         Width           =   972
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   312
         Index           =   44
         Left            =   5280
         TabIndex        =   109
         Top             =   1440
         Width           =   1212
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Estado"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   312
         Index           =   43
         Left            =   5280
         TabIndex        =   107
         Top             =   1800
         Width           =   1212
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Usuario"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   192
         Index           =   42
         Left            =   5280
         TabIndex        =   102
         Top             =   2160
         Width           =   972
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Cliente"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   312
         Index           =   8
         Left            =   240
         TabIndex        =   96
         Top             =   600
         Width           =   1212
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Concepto"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   312
         Index           =   9
         Left            =   240
         TabIndex        =   95
         Top             =   960
         Width           =   1212
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Detalle"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   312
         Index           =   10
         Left            =   8280
         TabIndex        =   94
         Top             =   360
         Width           =   1212
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Documento"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   312
         Index           =   11
         Left            =   240
         TabIndex        =   93
         Top             =   1440
         Width           =   1212
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Monto"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   312
         Index           =   13
         Left            =   240
         TabIndex        =   92
         Top             =   1800
         Width           =   1212
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Caja"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   192
         Index           =   15
         Left            =   240
         TabIndex        =   91
         Top             =   2160
         Width           =   972
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         Caption         =   "Usuario"
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
         Left            =   -65800
         TabIndex        =   88
         Top             =   840
         Visible         =   0   'False
         Width           =   2160
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         Caption         =   "Fecha"
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
         Left            =   -63760
         TabIndex        =   87
         Top             =   840
         Visible         =   0   'False
         Width           =   2292
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         Caption         =   "Registro"
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
         Height          =   312
         Left            =   -67480
         TabIndex        =   86
         Top             =   1080
         Visible         =   0   'False
         Width           =   1680
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         Caption         =   "Traslado"
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
         Height          =   312
         Left            =   -67480
         TabIndex        =   85
         Top             =   1440
         Visible         =   0   'False
         Width           =   1680
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         Caption         =   "Anulación"
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
         Height          =   312
         Left            =   -67480
         TabIndex        =   84
         Top             =   1800
         Visible         =   0   'False
         Width           =   1680
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
         ForeColor       =   &H00000080&
         Height          =   252
         Index           =   1
         Left            =   -69520
         TabIndex        =   82
         Top             =   3000
         Visible         =   0   'False
         Width           =   1092
      End
      Begin VB.Label Label3 
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
         Height          =   252
         Index           =   0
         Left            =   -64000
         TabIndex        =   81
         Top             =   3000
         Visible         =   0   'False
         Width           =   852
      End
   End
   Begin VB.Frame fraReportes 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5895
      Left            =   6720
      TabIndex        =   11
      Top             =   1200
      Visible         =   0   'False
      Width           =   6015
      Begin VB.OptionButton optReportes 
         Appearance      =   0  'Flat
         Caption         =   "Informe Especial por Bancos"
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
         Index           =   6
         Left            =   360
         TabIndex        =   28
         Top             =   3600
         Width           =   2775
      End
      Begin VB.OptionButton optReportes 
         Appearance      =   0  'Flat
         Caption         =   "Informe Especial para Cierres"
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
         Index           =   5
         Left            =   360
         TabIndex        =   27
         Top             =   2880
         Width           =   2775
      End
      Begin VB.OptionButton optReportes 
         Appearance      =   0  'Flat
         Caption         =   "Agrupado por Concepto"
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
         Index           =   4
         Left            =   360
         TabIndex        =   26
         Top             =   2040
         Width           =   2295
      End
      Begin VB.OptionButton optReportes 
         Appearance      =   0  'Flat
         Caption         =   "Agrupado por Oficina"
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
         Index           =   3
         Left            =   360
         TabIndex        =   25
         Top             =   1680
         Width           =   2295
      End
      Begin VB.OptionButton optReportes 
         Appearance      =   0  'Flat
         Caption         =   "Agrupado por Caja"
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
         Index           =   2
         Left            =   360
         TabIndex        =   24
         Top             =   1320
         Width           =   2295
      End
      Begin VB.OptionButton optReportes 
         Appearance      =   0  'Flat
         Caption         =   "Agrupado por Usuario"
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
         Index           =   1
         Left            =   360
         TabIndex        =   23
         Top             =   960
         Width           =   2295
      End
      Begin VB.OptionButton optReportes 
         Appearance      =   0  'Flat
         Caption         =   "Listado General"
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
         Index           =   0
         Left            =   360
         TabIndex        =   12
         Top             =   600
         Value           =   -1  'True
         Width           =   1695
      End
      Begin XtremeSuiteControls.ComboBox cboBanco 
         Height          =   312
         Left            =   1320
         TabIndex        =   61
         Top             =   4200
         Width           =   4212
         _Version        =   1441793
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
      Begin XtremeSuiteControls.ComboBox cboTipoReporte 
         Height          =   312
         Left            =   1320
         TabIndex        =   62
         Top             =   5040
         Width           =   2172
         _Version        =   1441793
         _ExtentX        =   3836
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
      Begin XtremeSuiteControls.PushButton btnRepInforme 
         Height          =   612
         Left            =   3600
         TabIndex        =   63
         Top             =   5040
         Width           =   1452
         _Version        =   1441793
         _ExtentX        =   2561
         _ExtentY        =   1080
         _StockProps     =   79
         Caption         =   "Informe"
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
         Appearance      =   16
         Picture         =   "frmSIF_ConsultaDocumentos.frx":951C
      End
      Begin XtremeSuiteControls.PushButton btnRepCerrar 
         Height          =   612
         Left            =   5040
         TabIndex        =   64
         Top             =   5040
         Width           =   852
         _Version        =   1441793
         _ExtentX        =   1503
         _ExtentY        =   1080
         _StockProps     =   79
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
         Appearance      =   16
         Picture         =   "frmSIF_ConsultaDocumentos.frx":9CD8
      End
      Begin XtremeSuiteControls.FlatEdit txtReporteUsuario 
         Height          =   312
         Left            =   1320
         TabIndex        =   65
         Top             =   3240
         Width           =   4212
         _Version        =   1441793
         _ExtentX        =   7429
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
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Usuario ...:"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   31
         Left            =   480
         TabIndex        =   31
         Top             =   3240
         Width           =   1095
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Reportes Especiales...:"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   30
         Left            =   120
         TabIndex        =   30
         Top             =   2520
         Width           =   1815
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Cuenta Bancaria...:"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   312
         Index           =   29
         Left            =   480
         TabIndex        =   29
         Top             =   3960
         Width           =   3252
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Reportes ...:"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   312
         Index           =   28
         Left            =   120
         TabIndex        =   22
         Top             =   240
         Width           =   1812
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo ...:"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   27
         Left            =   480
         TabIndex        =   21
         Top             =   5040
         Width           =   855
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FFFFFF&
         Index           =   0
         X1              =   6000
         X2              =   120
         Y1              =   4680
         Y2              =   4680
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   13560
      Top             =   480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSIF_ConsultaDocumentos.frx":A4A5
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSIF_ConsultaDocumentos.frx":AEC3
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSIF_ConsultaDocumentos.frx":B67F
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSIF_ConsultaDocumentos.frx":BE5B
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSIF_ConsultaDocumentos.frx":C628
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSIF_ConsultaDocumentos.frx":CFB5
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Timer TimerX 
      Interval        =   10
      Left            =   12240
      Top             =   120
   End
   Begin VB.Frame fraFilros 
      Caption         =   "Filtros adicionales"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5295
      Left            =   3600
      TabIndex        =   32
      Top             =   1200
      Visible         =   0   'False
      Width           =   7455
      Begin VB.CheckBox chkAsientosDesbalanceados 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "Transacciones con Asientos desbalanceados"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   840
         TabIndex        =   37
         Top             =   4200
         Width           =   4575
      End
      Begin VB.CheckBox chkDocumentosBloqueados 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "Documentos Bloqueados para Traslado Contable "
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   840
         TabIndex        =   33
         Top             =   4560
         Width           =   4575
      End
      Begin XtremeSuiteControls.ComboBox cboFormasPago 
         Height          =   312
         Left            =   1560
         TabIndex        =   58
         Top             =   2160
         Width           =   2652
         _Version        =   1441793
         _ExtentX        =   4683
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
      Begin XtremeSuiteControls.ComboBox cboCajas 
         Height          =   312
         Left            =   1560
         TabIndex        =   59
         Top             =   2640
         Width           =   5052
         _Version        =   1441793
         _ExtentX        =   8916
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
      Begin XtremeSuiteControls.ComboBox cboUsuarios 
         Height          =   312
         Left            =   1560
         TabIndex        =   60
         Top             =   3000
         Width           =   5052
         _Version        =   1441793
         _ExtentX        =   8916
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
      Begin XtremeSuiteControls.FlatEdit txtReferencia_01 
         Height          =   312
         Left            =   5640
         TabIndex        =   70
         Top             =   480
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
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   2
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtReferencia_02 
         Height          =   312
         Left            =   5640
         TabIndex        =   71
         Top             =   840
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
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   2
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtReferencia_03 
         Height          =   312
         Left            =   5640
         TabIndex        =   72
         Top             =   1200
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
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   2
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtFormaPagoNoRef 
         Height          =   312
         Left            =   5040
         TabIndex        =   73
         Top             =   2160
         Width           =   1572
         _Version        =   1441793
         _ExtentX        =   2773
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
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtCajasApertura 
         Height          =   312
         Left            =   1560
         TabIndex        =   74
         Top             =   3480
         Width           =   1572
         _Version        =   1441793
         _ExtentX        =   2773
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
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.PushButton btnFiltrosCerrar 
         Height          =   612
         Left            =   6360
         TabIndex        =   75
         Top             =   4320
         Width           =   852
         _Version        =   1441793
         _ExtentX        =   1503
         _ExtentY        =   1080
         _StockProps     =   79
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
         Appearance      =   16
         Picture         =   "frmSIF_ConsultaDocumentos.frx":D91A
      End
      Begin XtremeSuiteControls.FlatEdit txtCuenta 
         Height          =   312
         Left            =   1560
         TabIndex        =   69
         Top             =   1680
         Width           =   2652
         _Version        =   1441793
         _ExtentX        =   4678
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
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtUsuarioRegistra 
         Height          =   312
         Left            =   1560
         TabIndex        =   68
         Top             =   1200
         Width           =   2652
         _Version        =   1441793
         _ExtentX        =   4678
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
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtNoDocumento 
         Height          =   312
         Left            =   1560
         TabIndex        =   67
         Top             =   840
         Width           =   2652
         _Version        =   1441793
         _ExtentX        =   4678
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
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtNoTransaccion 
         Height          =   312
         Left            =   1560
         TabIndex        =   66
         Top             =   480
         Width           =   2652
         _Version        =   1441793
         _ExtentX        =   4678
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
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Referencia No.3"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   41
         Left            =   4320
         TabIndex        =   46
         Top             =   1200
         Width           =   1455
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Referencia No.2"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   40
         Left            =   4320
         TabIndex        =   45
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Referencia No.1"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   39
         Left            =   4320
         TabIndex        =   44
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Cuenta Contable"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   435
         Index           =   38
         Left            =   120
         TabIndex        =   43
         Top             =   1680
         Width           =   1335
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Usuario"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   37
         Left            =   120
         TabIndex        =   42
         Top             =   1200
         Width           =   1455
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "No. Documento"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   36
         Left            =   120
         TabIndex        =   41
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "No. Transacción"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   35
         Left            =   120
         TabIndex        =   40
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "No. Ref."
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   34
         Left            =   4320
         TabIndex        =   39
         Top             =   2160
         Width           =   1455
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Forma de Pago"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   33
         Left            =   120
         TabIndex        =   38
         Top             =   2160
         Width           =   1455
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "No. Apertura / Cierre de Cajas"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   675
         Index           =   32
         Left            =   120
         TabIndex        =   36
         Top             =   3480
         Width           =   1335
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FFFFFF&
         Index           =   1
         X1              =   7200
         X2              =   120
         Y1              =   3960
         Y2              =   3960
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Cajas"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   21
         Left            =   120
         TabIndex        =   35
         Top             =   2640
         Width           =   1215
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Usuarios de Caja"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   22
         Left            =   120
         TabIndex        =   34
         Top             =   3000
         Width           =   1695
      End
   End
   Begin MSComctlLib.StatusBar StatusBarX 
      Align           =   2  'Align Bottom
      Height          =   252
      Left            =   0
      TabIndex        =   4
      Top             =   8796
      Width           =   16932
      _ExtentX        =   29871
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4304
            MinWidth        =   4304
            Object.ToolTipText     =   "Casos Encontrados..:"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   6068
            MinWidth        =   6068
            Object.ToolTipText     =   "Total Registrado..:"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin FPSpreadADO.fpSpread vGrid 
      Height          =   3372
      Left            =   3600
      TabIndex        =   3
      Top             =   1200
      Width           =   10692
      _Version        =   524288
      _ExtentX        =   18860
      _ExtentY        =   5948
      _StockProps     =   64
      BackColorStyle  =   1
      BorderStyle     =   0
      EditEnterAction =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   12
      SpreadDesigner  =   "frmSIF_ConsultaDocumentos.frx":E0E7
      VScrollSpecial  =   -1  'True
      VScrollSpecialType=   2
      AppearanceStyle =   1
   End
   Begin MSComCtl2.FlatScrollBar FlatScrollBar 
      Height          =   255
      Left            =   13440
      TabIndex        =   10
      Top             =   4800
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   450
      _Version        =   393216
      Arrows          =   65536
      Min             =   -1
      Orientation     =   1638401
   End
   Begin XtremeSuiteControls.ListView lswDocumentos 
      Height          =   3015
      Left            =   120
      TabIndex        =   47
      Top             =   840
      Width           =   3015
      _Version        =   1441793
      _ExtentX        =   5318
      _ExtentY        =   5318
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
      MultiSelect     =   -1  'True
      HideSelection   =   0   'False
      View            =   3
      HideColumnHeaders=   -1  'True
      FullRowSelect   =   -1  'True
      HotTracking     =   -1  'True
      Appearance      =   16
   End
   Begin XtremeSuiteControls.ListView lswConceptos 
      Height          =   1815
      Left            =   120
      TabIndex        =   48
      Top             =   4680
      Width           =   3015
      _Version        =   1441793
      _ExtentX        =   5313
      _ExtentY        =   3196
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
      MultiSelect     =   -1  'True
      HideSelection   =   0   'False
      View            =   3
      HideColumnHeaders=   -1  'True
      FullRowSelect   =   -1  'True
      HotTracking     =   -1  'True
      Appearance      =   16
   End
   Begin XtremeSuiteControls.ComboBox cboEstado 
      Height          =   315
      Left            =   1560
      TabIndex        =   49
      Top             =   6600
      Width           =   1575
      _Version        =   1441793
      _ExtentX        =   2778
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
   Begin XtremeSuiteControls.ComboBox cboFechas 
      Height          =   315
      Left            =   1560
      TabIndex        =   50
      Top             =   6960
      Width           =   1575
      _Version        =   1441793
      _ExtentX        =   2778
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
      Left            =   1560
      TabIndex        =   51
      Top             =   7320
      Width           =   1575
      _Version        =   1441793
      _ExtentX        =   2773
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
      Left            =   1560
      TabIndex        =   52
      Top             =   7680
      Width           =   1575
      _Version        =   1441793
      _ExtentX        =   2773
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
   Begin XtremeSuiteControls.PushButton btnBuscar 
      Height          =   612
      Left            =   8280
      TabIndex        =   53
      Top             =   480
      Width           =   1212
      _Version        =   1441793
      _ExtentX        =   2138
      _ExtentY        =   1080
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
      Picture         =   "frmSIF_ConsultaDocumentos.frx":E986
   End
   Begin XtremeSuiteControls.FlatEdit txtBuscarPor 
      Height          =   312
      Left            =   3600
      TabIndex        =   54
      Top             =   840
      Width           =   4452
      _Version        =   1441793
      _ExtentX        =   7853
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
   Begin XtremeSuiteControls.ComboBox cbo 
      Height          =   312
      Left            =   3600
      TabIndex        =   55
      Top             =   480
      Width           =   4452
      _Version        =   1441793
      _ExtentX        =   7858
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
   Begin XtremeSuiteControls.PushButton btnInforme 
      Height          =   612
      Left            =   9480
      TabIndex        =   56
      Top             =   480
      Width           =   1572
      _Version        =   1441793
      _ExtentX        =   2773
      _ExtentY        =   1080
      _StockProps     =   79
      Caption         =   "Informe"
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
      Picture         =   "frmSIF_ConsultaDocumentos.frx":F3A4
   End
   Begin XtremeSuiteControls.PushButton btnExportar 
      Height          =   612
      Left            =   11040
      TabIndex        =   57
      Top             =   480
      Width           =   1572
      _Version        =   1441793
      _ExtentX        =   2773
      _ExtentY        =   1080
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
      Picture         =   "frmSIF_ConsultaDocumentos.frx":FB60
   End
   Begin XtremeSuiteControls.PushButton btnMain 
      Height          =   372
      Index           =   0
      Left            =   14040
      TabIndex        =   77
      Top             =   4740
      Width           =   1452
      _Version        =   1441793
      _ExtentX        =   2561
      _ExtentY        =   656
      _StockProps     =   79
      Caption         =   "Imprimir"
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
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Picture         =   "frmSIF_ConsultaDocumentos.frx":10365
   End
   Begin XtremeSuiteControls.PushButton btnMain 
      Height          =   372
      Index           =   1
      Left            =   15480
      TabIndex        =   78
      Top             =   4740
      Width           =   1452
      _Version        =   1441793
      _ExtentX        =   2561
      _ExtentY        =   656
      _StockProps     =   79
      Caption         =   "Reversar"
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
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Picture         =   "frmSIF_ConsultaDocumentos.frx":10A6C
   End
   Begin XtremeSuiteControls.FlatEdit txtTransaccion 
      Height          =   312
      Left            =   11040
      TabIndex        =   113
      Top             =   4800
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
      Appearance      =   2
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtDocCod 
      Height          =   312
      Left            =   4800
      TabIndex        =   114
      Top             =   4800
      Width           =   864
      _Version        =   1441793
      _ExtentX        =   1524
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
   Begin XtremeSuiteControls.FlatEdit txtDocName 
      Height          =   312
      Left            =   5640
      TabIndex        =   115
      Top             =   4800
      Width           =   4584
      _Version        =   1441793
      _ExtentX        =   8086
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
      Appearance      =   2
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtFiltraTipoDoc 
      Height          =   315
      Left            =   120
      TabIndex        =   126
      Top             =   480
      Width           =   3015
      _Version        =   1441793
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
   Begin XtremeSuiteControls.FlatEdit txtFiltraConcepto 
      Height          =   315
      Left            =   120
      TabIndex        =   127
      Top             =   4320
      Width           =   3015
      _Version        =   1441793
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
   Begin XtremeSuiteControls.CheckBox chkDocumentos 
      Height          =   210
      Left            =   2880
      TabIndex        =   129
      Top             =   120
      Width           =   210
      _Version        =   1441793
      _ExtentX        =   370
      _ExtentY        =   370
      _StockProps     =   79
      Transparent     =   -1  'True
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Value           =   1
   End
   Begin XtremeSuiteControls.CheckBox chkFiltros 
      Height          =   210
      Left            =   6840
      TabIndex        =   130
      Top             =   120
      Width           =   1170
      _Version        =   1441793
      _ExtentX        =   2064
      _ExtentY        =   370
      _StockProps     =   79
      Caption         =   "+ Filtros ?"
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
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Número"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Index           =   24
      Left            =   10320
      TabIndex        =   14
      Top             =   4800
      Width           =   855
   End
   Begin VB.Label lblConsultaMonto 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   11520
      TabIndex        =   13
      Top             =   720
      Width           =   1455
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Documento"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Index           =   7
      Left            =   3720
      TabIndex        =   9
      Top             =   4800
      Width           =   1215
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Index           =   1
      Left            =   120
      TabIndex        =   8
      Top             =   6960
      Width           =   1215
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Estado"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Index           =   6
      Left            =   120
      TabIndex        =   7
      Top             =   6600
      Width           =   1215
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo Concepto ...:"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Index           =   3
      Left            =   120
      TabIndex        =   6
      Top             =   3960
      Width           =   1815
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Documentos ...:"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Index           =   2
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Corte"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Index           =   5
      Left            =   120
      TabIndex        =   2
      Top             =   7680
      Width           =   1215
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Inicio"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Index           =   4
      Left            =   120
      TabIndex        =   1
      Top             =   7320
      Width           =   1215
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Buscar por ...:"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Index           =   0
      Left            =   3600
      TabIndex        =   0
      Top             =   120
      Width           =   1575
   End
   Begin VB.Image imgBanner 
      Height          =   9390
      Left            =   0
      Picture         =   "frmSIF_ConsultaDocumentos.frx":1116C
      Stretch         =   -1  'True
      Top             =   0
      Width           =   3285
   End
End
Attribute VB_Name = "frmSIF_ConsultaDocumentos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vCodigo As String
Dim vTipoDocu As String
Dim vScroll As Boolean
Dim strSQLReporte As String
Dim strDetalle As String
Dim strSQL As String, rs As New ADODB.Recordset, itmX As ListViewItem



Private Function fxCajasUltimaApertura(pCajas As String) As Long
Dim Resultado As Long

On Error GoTo vError

Resultado = 0

strSQL = "select dbo.fxSIFDocsCajaUltimaApertura('" & pCajas & "') as Resultado"
Call OpenRecordSet(rs, strSQL)
    Resultado = rs!Resultado
rs.Close

vError:

fxCajasUltimaApertura = Resultado

End Function

Private Sub btnBuscar_Click()
    Call sbLimpia
    Call sbBuscar
    Call vGrid_Click(1, 1)
End Sub

Private Sub btnDocConsultaCerrar_Click()
fraDocumento.Visible = False
End Sub

Private Sub btnDocFix_Click()
Dim strSQL As String

On Error GoTo vError

If txtDocumento.Text <> txtDocumento.Tag Then
    strSQL = "update sif_Transacciones set Documento = '" & txtDocumento.Text & "' where tipo_documento = '" _
           & txtDocCod.Text & "' and cod_transaccion = '" & txtTransaccion.Text & "'"
    Call ConectionExecute(strSQL)
   
   strSQL = "TDoc.: " & txtDocCod.Text & " - NDoc.:" & txtTransaccion.Text & " - Act.Doc.:" & txtDocumento.Text & " - Ant.Doc.:" & txtDocumento.Tag
    
   Call Bitacora("Modifica", strSQL)

   txtDocumento.Tag = txtDocumento.Text

   MsgBox "Cambio Satisfactorio!" & vbCrLf & vbCrLf & strSQL, vbInformation

Else
   MsgBox "Realice la modificación del No. de documento y luego presione ajuste!", vbExclamation
End If



Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub btnExportar_Click()
Dim vHeaders As vGridHeaders
    vHeaders.Columnas = 12
    vHeaders.Headers(1) = "# Transaccion"
    vHeaders.Headers(2) = "# Documento"
    vHeaders.Headers(3) = "Tipo.Doc."
    vHeaders.Headers(4) = "Monto"
    vHeaders.Headers(5) = "Estado"
    vHeaders.Headers(6) = "Fec.Registro"
    vHeaders.Headers(7) = "Usr.Registro"
    vHeaders.Headers(8) = "Caja"
    vHeaders.Headers(9) = "Apertura"
    vHeaders.Headers(10) = "Oficina"
    vHeaders.Headers(11) = "Cliente"
    vHeaders.Headers(12) = "Detalle"

 Call sbSIFGridExportar(vGrid, vHeaders, "ProGrX_ConsultaDocumentos")

'Select Case ButtonMenu.Key
'  Case "Excel"
'      Call sbSIFGridExportar(vGrid, vHeaders, "ProGrX_ConsultaDocumentos")
'  Case "HTML"
'      Call sbSIFGridExportar(vGrid, vHeaders, "ProGrX_ConsultaDocumentos", "HTML")
'End Select
End Sub

Private Sub btnFiltrosCerrar_Click()
fraFilros.Visible = False
chkFiltros.Value = vbUnchecked
End Sub

Private Sub btnInforme_Click()
    fraReportes.Top = btnInforme.Top + 480
    fraReportes.Left = btnInforme.Left + 920
    fraReportes.Visible = True
End Sub

Private Sub btnMain_Click(Index As Integer)
 Select Case Index
   Case 0 'Reimprimir
        Call sbImprimeRecibo(vCodigo, vTipoDocu)
   Case 1 'Reversar
        Call sbReversar(vTipoDocu, vCodigo)
 End Select
End Sub

Private Sub btnRepCerrar_Click()
      fraReportes.Visible = False

End Sub

Private Sub btnRepInforme_Click()
Dim vFiltros As String, vFiltrosEtiquetas As String

With frmContenedor.Crt
    .Reset
    .WindowTitle = "Reportes del Módulo: Control de Documentos"
    .WindowState = crptMaximized
    .WindowShowGroupTree = True
    
    .Formulas(1) = "Empresa='" & GLOBALES.gstrNombreEmpresa & "'"
    .Formulas(2) = "Detalle = '" & strDetalle & "'"
    .Formulas(3) = "Usuario = 'Usuario..:" & glogon.Usuario & "'"
    .Formulas(4) = "Fecha = 'Fecha ...:" & fxFechaServidor & "'"
    .Connect = glogon.ConectRPT


    vFiltros = fxReportesFiltros
        Select Case True
            Case optReportes.Item(0).Value 'Reporte General
                If Mid(cboTipoReporte, 1, 1) = "D" Then
                   .ReportFileName = SIFGlobal.fxPathReportes("Sys_DocGeneral.rpt")
                Else
                   .ReportFileName = SIFGlobal.fxPathReportes("Sys_DocGeneralRsm.rpt")
                End If
            Case optReportes.Item(1).Value 'Agrupado por Usuario
                If Mid(cboTipoReporte, 1, 1) = "D" Then
                    .ReportFileName = SIFGlobal.fxPathReportes("Sys_DocUsuario.rpt")
                Else
                    .ReportFileName = SIFGlobal.fxPathReportes("Sys_DocUsuarioRsm.rpt")
                End If
            Case optReportes.Item(2).Value 'Agrupado por Caja
                If Mid(cboTipoReporte, 1, 1) = "D" Then
                    .ReportFileName = SIFGlobal.fxPathReportes("Sys_DocCaja.rpt")
                Else
                    .ReportFileName = SIFGlobal.fxPathReportes("Sys_DocCajaRsm.rpt")
                End If
            Case optReportes.Item(3).Value 'Agrupado por Oficina
                If Mid(cboTipoReporte, 1, 1) = "D" Then
                    .ReportFileName = SIFGlobal.fxPathReportes("Sys_DocOficina.rpt")
                Else
                    .ReportFileName = SIFGlobal.fxPathReportes("Sys_DocOficinaRsm.rpt")
                End If
            Case optReportes.Item(4).Value 'Agrupado por Concepto
                If Mid(cboTipoReporte, 1, 1) = "D" Then
                    .ReportFileName = SIFGlobal.fxPathReportes("Sys_DocConcepto.rpt")
                Else
                    .ReportFileName = SIFGlobal.fxPathReportes("Sys_DocConceptoRsm.rpt")
                End If
        
        
            Case optReportes.Item(5).Value 'Especial Cierre
                'Filtros nuevos no utilizar defaults
                vFiltros = ""
                
                .Formulas(5) = "SUBTITULO='CIERRE DE CAJAS DEL " & Format(dtpInicio.Value, "yyyy/mm/dd") _
                             & " HASTA " & Format(dtpCorte.Value, "yyyy/mm/dd") & "'"
                
                .Formulas(6) = "fxUsuarioCierre = '" & txtReporteUsuario.Text & "'"
                .Formulas(7) = "fxUsuarioNombre = '" & fxUsuarioNombre(txtReporteUsuario.Text) & "'"
                
                
                .ReportFileName = SIFGlobal.fxPathReportes("Sys_DocEspecialCierre.rpt")
                
                
                .SubreportToChange = "subOperaciones"
                .Connect = glogon.ConectRPT
                .SelectionFormula = "({vEspecialCierres.ESTADOSOL} = 'N' OR {vEspecialCierres.ESTADOSOL} ='F')" _
                                  & " AND ({vEspecialCierres.EMITIR} = 'CK' OR {vEspecialCierres.EMITIR} = 'EF' OR {vEspecialCierres.EMITIR} = 'TE')" _
                                  & " AND {vEspecialCierres.USERFOR} = '" & txtReporteUsuario.Text _
                                  & "' AND CDATE({vEspecialCierres.FECHAFORP}) " & fxFechaReportes(1)
                
                .SubreportToChange = "subOperaciones"
                .Connect = glogon.ConectRPT
                .SelectionFormula = "({vEspecialCierres.ESTADOSOL} = 'N' OR {vEspecialCierres.ESTADOSOL} ='F')" _
                                  & " AND ({vEspecialCierres.EMITIR} = 'CK' OR {vEspecialCierres.EMITIR} = 'EF' OR {vEspecialCierres.EMITIR} = 'TE')" _
                                  & " AND {vEspecialCierres.USERFOR} = '" & txtReporteUsuario.Text _
                                  & "' AND CDATE({vEspecialCierres.FECHAFORP}) " & fxFechaReportes(1)
                
                
                .SubreportToChange = "sbCKCajas"
                .Connect = glogon.ConectRPT
                .SelectionFormula = "CDATE({ASE_CK_CAJA.FECHA}) " & fxFechaReportes(1) _
                            & " AND {ASE_CK_CAJA.USUARIO} = '" & txtReporteUsuario.Text & "'"
                
                
                .SubreportToChange = "sbFondos"
                .Connect = glogon.ConectRPT
                .SelectionFormula = "CDATE({FND_LIQUIDACION.FECHA}) " & fxFechaReportes(1) _
                            & " AND {FND_LIQUIDACION.USUARIO} = '" & txtReporteUsuario.Text & "'"
                   
                .SubreportToChange = "subDocumentos"
                .Connect = glogon.ConectRPT
                .SelectionFormula = "CDATE({vSIFDocumentos.REGISTRO_FECHA}) " & fxFechaReportes(1) _
                                  & " AND {vSIFDocumentos.REGISTRO_USUARIO} = '" & txtReporteUsuario.Text & "'" _
                                  & " AND {vSIFDocumentos.APLICA_CIERRE_ESPECIAL} = 1"
                        
                   
        
            Case optReportes.Item(6).Value 'Especial Bancos
                'Filtros nuevos no utilizar defaults
                vFiltros = ""
                
                If Mid(cboTipoReporte, 1, 1) = "D" Then
                    .ReportFileName = SIFGlobal.fxPathReportes("Sys_DocEspecialBancos.rpt")
                Else
                    .ReportFileName = SIFGlobal.fxPathReportes("Sys_DocEspecialBancosRsm.rpt")
                End If
        
        
                  .ReportFileName = SIFGlobal.fxPathReportes("Sys_DocumentoEspecialCierreBanco.rpt")
                  .SelectionFormula = "cdate({CHEQUES.FECHA_EMISION}) " & fxFechaReportes(1) _
                                    & " AND {CHEQUES.ID_BANCO} = " & cboBanco.ItemData(cboBanco.ListIndex)
                 
                
                  .SubreportToChange = "sbEstadistica"
                  .Connect = glogon.ConectRPT
                  .StoredProcParam(0) = cboBanco.ItemData(cboBanco.ListIndex)
                  .StoredProcParam(1) = Format(dtpInicio.Value, "yyyy/mm/dd")
                  
                            
        
        End Select
        
        'Si son los reportes Especiales no aplicar filtros default
        If Not (optReportes.Item(5).Value Or optReportes.Item(6).Value) Then
            .SelectionFormula = vFiltros
        End If
        .PrintReport
           
 
End With

End Sub

Private Sub cboCajas_Click()

If cboCajas.Text <> "TODOS" Then
   strSQL = "select USUARIO as itmx from CAJAS_USUARIOS where COD_CAJA = '" & cboCajas.ItemData(cboCajas.ListIndex) & "'"
   Call sbCbo_Llena_New(cboUsuarios, strSQL, False)
   
   txtCajasApertura.Text = fxCajasUltimaApertura(cboCajas.ItemData(cboCajas.ListIndex))
   
Else
   cboUsuarios.Clear
   cboUsuarios.AddItem "TODOS"
   txtCajasApertura.Text = 1
End If



End Sub


Private Sub chkConceptos_Click()
Dim i As Integer

For i = 1 To lswConceptos.ListItems.Count
  lswConceptos.ListItems.Item(i).Checked = chkConceptos.Value
Next i

End Sub

Private Sub sbConceptos_Load()

On Error GoTo vError

Me.MousePointer = vbHourglass

txtFiltraConcepto.Text = fxSysCleanTxtInject(txtFiltraConcepto.Text)
 
lswConceptos.ListItems.Clear
 
strSQL = "select cod_concepto,DESCRIPCION from sif_conceptos" _
       & " Where descripcion like '%" & txtFiltraConcepto.Text & "%'" _
       & " order by descripcion"
Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
 Set itmX = lswConceptos.ListItems.Add(, , rs!Descripcion)
     itmX.Tag = rs!cod_concepto
     itmX.Checked = chkConceptos.Value
 rs.MoveNext
Loop
rs.Close

Me.MousePointer = vbDefault

Exit Sub

vError:
    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub sbTipoDocs_Load()
On Error GoTo vError

Me.MousePointer = vbHourglass

txtFiltraTipoDoc.Text = fxSysCleanTxtInject(txtFiltraTipoDoc.Text)

lswDocumentos.ListItems.Clear

strSQL = "select tipo_documento as IdX, rtrim(Descripcion) as ItmX" _
       & " from sif_documentos " _
       & " where Activo = 1 and descripcion like '%" & txtFiltraTipoDoc.Text & "%'" _
       & " order by descripcion"
Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
 Set itmX = lswDocumentos.ListItems.Add(, , rs!itmX)
     itmX.Tag = rs!IdX
     itmX.Checked = chkDocumentos.Value
 rs.MoveNext
Loop
rs.Close

Me.MousePointer = vbDefault

Exit Sub

vError:
    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub chkDocumentos_Click()
Dim i As Integer

For i = 1 To lswDocumentos.ListItems.Count
  lswDocumentos.ListItems.Item(i).Checked = chkDocumentos.Value
Next i

End Sub

Private Sub sbBuscar()
Dim strSQL As String, i As Integer
Dim vCadena As String, iCantidad As Integer

On Error GoTo vError

Me.MousePointer = vbHourglass
iCantidad = 0

strSQL = "select Cod_transaccion,isnull(documento,0),Tipo_documento,monto,case when estado = 'I' then 'Emitido'" _
       & " when estado = 'P' then 'Pendiente' when  estado = 'A' then 'Anulado'  end as Estado" _
       & ",isnull(Registro_fecha,'') as Fecha_registro,isnull(Registro_Usuario,'') as 'Usuario', cod_caja,cod_apertura,cod_oficina,Cliente_Nombre" _
       & ",isnull(Detalle,'') from Sif_Transacciones"
 
Select Case Mid(cbo, 1, 2)
    Case "01"
      strSQL = strSQL & " where cliente_Identificacion like '%" & txtBuscarPor.Text & "%'"
   
    Case "02"
      strSQL = strSQL & " where Cliente_Nombre like '%" & txtBuscarPor.Text & "%'"
    
End Select



'Lista de Documentos
If lswDocumentos.ListItems.Count > 0 Then
    vCadena = " and Tipo_Documento in('"
    For i = 1 To lswDocumentos.ListItems.Count
      If lswDocumentos.ListItems.Item(i).Checked Then
        vCadena = vCadena & "','" & lswDocumentos.ListItems.Item(i).Tag
        iCantidad = iCantidad + 1
      End If
    Next i
    strSQL = strSQL & vCadena & "')"
End If

iCantidad = 0
'Lista de Conceptos
If lswConceptos.ListItems.Count > 0 Then
    vCadena = " and Cod_Concepto in('"
    For i = 1 To lswConceptos.ListItems.Count
      If lswConceptos.ListItems.Item(i).Checked Then
        vCadena = vCadena & "','" & lswConceptos.ListItems.Item(i).Tag
        iCantidad = iCantidad + 1
      End If
    Next i
    strSQL = strSQL & vCadena & "')"
End If

Select Case cboFechas.Text
  Case "Registro"
    strSQL = strSQL & " and Registro_Fecha between '" & Format(dtpInicio.Value, "yyyy/mm/dd") _
           & " 00:00:00' and '" & Format(dtpCorte.Value, "yyyy/mm/dd") & " 23:59:59'"
    
  Case "Anulación"
    strSQL = strSQL & " and anulacion_fecha between '" & Format(dtpInicio.Value, "yyyy/mm/dd") _
           & " 00:00:00' and '" & Format(dtpCorte.Value, "yyyy/mm/dd") & " 23:59:59'"
    
  Case "Traslado"
    strSQL = strSQL & " and traspaso_fecha between '" & Format(dtpInicio.Value, "yyyy/mm/dd") _
           & " 00:00:00' and '" & Format(dtpCorte.Value, "yyyy/mm/dd") & " 23:59:59'"
End Select

Select Case cboEstado.Text
  Case "Impreso"
     strSQL = strSQL & " and estado in('I','E')"
     
  Case "Pendiente"
     strSQL = strSQL & " and estado = 'P'"
  
End Select

strSQL = strSQL & " and Traslado_Bloqueo = " & chkDocumentosBloqueados.Value

If Trim(txtUsuarioRegistra.Text) <> "" Then
      strSQL = strSQL & " and Registro_Usuario like '%" & txtUsuarioRegistra.Text & "%'"
End If

If Trim(txtNoTransaccion.Text) <> "" Then
      strSQL = strSQL & " and Cod_Transaccion like '%" & txtNoTransaccion.Text & "%'"
End If


If Trim(txtNoDocumento.Text) <> "" Then
      strSQL = strSQL & " and Documento like '%" & txtNoDocumento.Text & "%'"
End If

If Trim(txtReferencia_01.Text) <> "" Then
      strSQL = strSQL & " and Referencia_01 like '%" & txtReferencia_01.Text & "%'"
End If

If Trim(txtReferencia_02.Text) <> "" Then
      strSQL = strSQL & " and Referencia_02 like '%" & txtReferencia_02.Text & "%'"
End If

If Trim(txtReferencia_03.Text) <> "" Then
      strSQL = strSQL & " and Referencia_03 like '%" & txtReferencia_03.Text & "%'"
End If


If cboCajas.Text <> "" Then
    If cboCajas.Text <> "TODOS" Then
        strSQL = strSQL & " and cod_caja =  '" & cboCajas.ItemData(cboCajas.ListIndex) & "'"
        
        If IsNumeric(txtCajasApertura.Text) Then
            If CLng(txtCajasApertura.Text) > 0 Then strSQL = strSQL & " and cod_Apertura = " & txtCajasApertura.Text
        End If
    End If
End If

If Len(cboUsuarios.Text) > 0 And cboUsuarios.Text <> "TODOS" Then
   strSQL = strSQL & " and registro_usuario = '" & cboUsuarios.Text & "'"
End If
      
      
If txtCuenta.Text <> "" Then
   strSQL = strSQL & " and dbo.fxSIFDocsCuentaExiste(Tipo_Documento,Cod_Transaccion,'" _
          & fxgCntCuentaFormato(False, txtCuenta.Text, 0) & "')" _
          & " = 1"
End If

If cboFormasPago.Text <> "" Then
    If cboFormasPago.Text <> "TODOS" Then
       strSQL = strSQL & " and dbo.fxSIFDocsFormaPagoExiste(Tipo_Documento,Cod_Transaccion,'" _
              & cboFormasPago.ItemData(cboFormasPago.ListIndex) & "','" & txtFormaPagoNoRef.Text & "')" _
              & " = 1"
    End If
End If


If chkAsientosDesbalanceados.Value = vbChecked Then
   strSQL = strSQL & " and dbo.fxSIFDocsAsientoBalanceado(Tipo_Documento,Cod_Transaccion)" _
          & " = 0"
End If

strSQL = strSQL & " Order by Registro_fecha desc, Tipo_Documento, Cod_Transaccion desc"

Call sbCargaGridLocal(vGrid, 12, strSQL)

Me.MousePointer = vbDefault

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub sbCargaGridLocal(vGrid As Object, vGridMaxCol As Integer, strSQL As String)
Dim rs As New ADODB.Recordset, i As Integer
Dim curMonto As Currency

On Error GoTo vError

vGrid.MaxCols = vGridMaxCol
vGrid.MaxRows = 1
vGrid.Row = vGrid.MaxRows
For i = 1 To vGrid.MaxCols
 vGrid.Col = i
 vGrid.Text = ""
Next i

curMonto = 0

Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
  vGrid.Row = vGrid.MaxRows
  For i = 1 To vGrid.MaxCols
    vGrid.Col = i

'    If rs.Fields(i - 1).Type = 135 Then
'        If Year(rs.Fields(i - 1).Value) > 1900 Then
'           vGrid.Text = Format((rs.Fields(i - 1).Value & ""), "dd/mm/yyyy")
'        End If
'    Else
'        vGrid.Text = CStr(rs.Fields(i - 1).Value & "")
'    End If
    vGrid.Text = CStr(rs.Fields(i - 1).Value & "")
  Next i
  vGrid.MaxRows = vGrid.MaxRows + 1
  curMonto = curMonto + rs!Monto
  rs.MoveNext
Loop

StatusBarX.Panels(1).Text = "Casos ..: " & Format(rs.RecordCount, "###,###,##0")
StatusBarX.Panels(2).Text = "Monto ..: " & Format(curMonto, "Standard")

rs.Close

vGrid.MaxRows = vGrid.MaxRows - 1

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub chkFiltros_Click()

On Error Resume Next

fraFilros.Top = 1200
fraFilros.Left = 3480

If chkFiltros.Value = vbChecked Then
    fraFilros.Visible = True
Else
    fraFilros.Visible = False
End If

End Sub

Private Sub FlatScrollBar_Change()
Dim strSQL As String, rs As New ADODB.Recordset
Dim i As Integer

On Error GoTo vError

tcMain.Item(0).Selected = True


If vScroll Then
 If txtTransaccion.Text = "" Then txtTransaccion.Text = 0
    
    i = Len(txtTransaccion.Text)
    
    strSQL = "select Top 1 cod_transaccion from sif_transacciones"
    
'       strSQL = strSQL & " where tipo_documento = '" & Trim(txtDocCod.Text) & "' and (replicate('0', " & i & " - len(cod_transaccion)) + cod_transaccion) > '" & txtTransaccion & "' order by cod_transaccion asc"
    
    If FlatScrollBar.Value = 1 Then
       strSQL = strSQL & " where tipo_documento = '" & Trim(txtDocCod.Text) & "' and cod_transaccion > '" & txtTransaccion & "' order by cod_transaccion asc"
    Else
       strSQL = strSQL & " where tipo_documento = '" & Trim(txtDocCod.Text) & "' and cod_transaccion < '" & txtTransaccion & "' order by cod_transaccion desc"
    End If
    
    Call OpenRecordSet(rs, strSQL)
    If Not rs.EOF And Not rs.BOF Then
      txtTransaccion = rs!Cod_Transaccion
      Call txtTransaccion_KeyDown(vbKeyReturn, 0)
      
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

Private Sub sbDocUltDocumento()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

Me.MousePointer = vbHourglass

txtTransaccion.SetFocus

strSQL = "select  isnull(LTRIM( max( REPLICATE(' ', 40 - len(COD_TRANSACCION)) + COD_TRANSACCION  ) ), '') as 'Transaccion'" _
       & " from Sif_Transacciones" _
       & " where Tipo_Documento = '" & txtDocCod.Text & "'"
Call OpenRecordSet(rs, strSQL)
If Not rs.EOF And Not rs.BOF Then
   txtTransaccion.Text = rs!Transaccion & ""
   Call txtTransaccion_KeyDown(vbKeyReturn, 0)
Else
   txtTransaccion.Text = ""
   Call sbLimpia
End If
rs.Close

Me.MousePointer = vbDefault
Exit Sub

vError:
   Me.MousePointer = vbDefault
   MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub sbAfectaciones(strCodigo As String, strDocumento As String)
Dim strSQL As String, rs As New ADODB.Recordset
Dim vTcon As String, i As Integer, curCargos As Currency
Dim curIntC As Currency, curIntM As Currency, curAmortiza As Currency
Dim curPoliza As Currency

Me.MousePointer = vbHourglass

On Error GoTo vError


curIntC = 0
curIntM = 0
curAmortiza = 0
curCargos = 0
curPoliza = 0



With vGridAfectaciones

Select Case .Sheet
 
  Case 1 'Creditos
        .MaxRows = 0
        .MaxCols = 11
  
        strSQL = "select * from vSIF_CtrlDoc_Crd_Detalle" _
               & " Where TCon= '" & strDocumento & "' And NCon = '" & strCodigo & "'"
        Call OpenRecordSet(rs, strSQL)
        Do While Not rs.EOF
          .MaxRows = .MaxRows + 1
          .Row = .MaxRows
          .Col = 1
          .Text = CStr(rs!Id_solicitud)
          .Col = 2
          .Text = CStr(rs!Codigo)
          .Col = 3
          .Text = Format(rs!Proceso, "####-##")
          
          .Col = 4
          .Text = Format(rs!IntCor, "Standard")
          .Col = 5
          .Text = Format(rs!IntMor, "Standard")
          .Col = 6
          .Text = Format(rs!Cargo, "Standard")
          .Col = 7
          .Text = Format(rs!Poliza, "Standard")
          .Col = 8
          .Text = Format(rs!Principal, "Standard")
          .Col = 9
          .Text = CStr(rs!Cedula)
          .Col = 10
          .Text = CStr(rs!Nombre)
          .Col = 11
          .Text = CStr(rs!Concepto)
          .Col = 12
          .Text = CStr(rs!Institucion)
          
          curAmortiza = curAmortiza + rs!Principal
          curCargos = curCargos + rs!Cargo
          curIntC = curIntC + rs!IntCor
          curIntM = curIntM + rs!IntMor
          curPoliza = curPoliza + rs!Poliza
          rs.MoveNext
        Loop
        rs.Close

        'Totales
        .MaxRows = .MaxRows + 1
        .Row = .MaxRows
        .Col = 3
        .Text = "Totales :"
        .Col = 4
        .Text = Format(curIntC, "Standard")
        .Col = 5
        .Text = Format(curIntM, "Standard")
        .Col = 6
        .Text = Format(curCargos, "Standard")
        .Col = 7
        .Text = Format(curPoliza, "Standard")
        .Col = 8
        .Text = Format(curAmortiza, "Standard")
        .Col = 9
        .Text = Format(curIntC + curIntM + curAmortiza + curCargos + curPoliza, "Standard")

    Case 2 'Fondos
    
        .MaxRows = 0
        .MaxCols = 8
  
        strSQL = "select * from vSIF_CtrlDoc_Fnd_Detalle" _
               & " Where TCon= '" & strDocumento & "' And NCon = '" & strCodigo & "'"
        Call OpenRecordSet(rs, strSQL)
        Do While Not rs.EOF
          .MaxRows = .MaxRows + 1
          .Row = .MaxRows
          .Col = 1
          .Text = Trim(rs!cod_plan)
          .Col = 2
          .Text = CStr(rs!cod_Contrato)
          .Col = 3
          .Text = Trim(rs!Cedula)
          .Col = 4
          .Text = Trim(rs!Nombre)
          .Col = 5
          .Text = Format(rs!Monto, "Standard")
          .Col = 6
          .Text = Trim(rs!Concepto)
          .Col = 7
          .Text = Trim(rs!Institucion)
          .Col = 8
          .Text = Trim(rs!PlanDesc)
          
          curAmortiza = curAmortiza + rs!Monto
          rs.MoveNext
        Loop
        rs.Close

        'Totales
        .MaxRows = .MaxRows + 1
        .Row = .MaxRows
        .Col = 2
        .Text = "Totales :"
        .Col = 5
        .Text = Format(curAmortiza, "Standard")

    
    Case 3 'Patrimonio
        .MaxRows = 0
        .MaxCols = 6
  
        strSQL = "select * from vSIF_CtrlDoc_Pat_Detalle" _
               & " Where TCon= '" & strDocumento & "' And NCon = '" & strCodigo & "'"
        Call OpenRecordSet(rs, strSQL)
        Do While Not rs.EOF
          .MaxRows = .MaxRows + 1
          .Row = .MaxRows
          .Col = 1
          .Text = Trim(rs!Tipo_Aporte)
          .Col = 2
          .Text = Trim(rs!Cedula)
          .Col = 3
          .Text = Trim(rs!Nombre)
          .Col = 4
          .Text = Format(rs!Monto, "Standard")
          .Col = 5
          .Text = Trim(rs!Concepto)
          .Col = 6
          .Text = Trim(rs!Institucion)
          
          curAmortiza = curAmortiza + rs!Monto
          rs.MoveNext
        Loop
        rs.Close

        'Totales
        .MaxRows = .MaxRows + 1
        .Row = .MaxRows
        .Col = 2
        .Text = "Totales :"
        .Col = 4
        .Text = Format(curAmortiza, "Standard")

    
    
  End Select
End With

Me.MousePointer = vbDefault
Exit Sub



vError:
   Me.MousePointer = vbDefault
   MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub


Private Function fxFechaReportes(vTipo As Integer) As String

fxFechaReportes = " in Date(" & Format(dtpInicio.Value, "yyyy,mm,dd") & ")" _
                & " to Date(" & Format(dtpCorte.Value, "yyyy,mm,dd") & ")"

End Function

Private Function fxUsuarioNombre(vUsuario As String) As String
Dim strSQL As String, rs As New ADODB.Recordset

strSQL = "select descripcion from usuarios where nombre = '" & vUsuario & "'"
Call OpenRecordSet(rs, strSQL)
If rs.EOF And rs.BOF Then
 fxUsuarioNombre = "[SIN DESCRIPCION]"
Else
 fxUsuarioNombre = "[" & UCase(Trim(rs!Descripcion)) & "]"

End If
rs.Close
End Function

Private Sub Form_Activate()
 vModulo = 10
End Sub

Private Sub Form_Load()
Dim strSQL As String, rs As New ADODB.Recordset

vModulo = 10

vScroll = False
 FlatScrollBar.Value = 0
vScroll = True

Call Formularios(Me)
Call RefrescaTags(Me)

lswConceptos.ColumnHeaders.Add , , "", 3150
lswDocumentos.ColumnHeaders.Add , , "", 3150

vGrid.AppearanceStyle = fxGridStyle
vgridAsiento.AppearanceStyle = fxGridStyle
vgridAsiento.MaxRows = 0

tcMain.Item(0).Selected = True

cbo.AddItem "01 - Cedula"
cbo.AddItem "02 - Nombre"

cboTipoReporte.Clear
cboTipoReporte.AddItem "Detalle"
cboTipoReporte.AddItem "Resumen"
cboTipoReporte.Text = "Detalle"

cbo.Text = "02 - Nombre"
'
cboFechas.Clear
cboFechas.AddItem "Registro"
cboFechas.AddItem "Anulación"
cboFechas.AddItem "Traslado"
cboFechas.AddItem "[TODAS]"
'
cboFechas.Text = "Registro"

'
cboEstado.Clear
cboEstado.AddItem "Pendiente"
cboEstado.AddItem "Impreso"
cboEstado.AddItem "Anulado"
cboEstado.AddItem "[TODOS]"
cboEstado.Text = "[TODOS]"

dtpInicio.Value = fxFechaServidor
dtpCorte.Value = dtpInicio.Value

Call sbLimpia

End Sub

Private Sub Form_Resize()
On Error Resume Next

imgBanner.Height = Me.Height
 
vGrid.Width = Me.Width - (vGrid.Left + 350)
vGrid.Height = Me.Height - 6145
tcMain.Width = vGrid.Width
tcMain.Top = Me.Height - 4455



Label1(6).Top = Me.Height - 3375
cboEstado.Top = Label1(6).Top
Label1(1).Top = Me.Height - 3015
cboFechas.Top = Label1(1).Top

Label1(4).Top = Me.Height - 2655
dtpInicio.Top = Label1(4).Top
Label1(5).Top = Me.Height - 2295
dtpCorte.Top = Label1(5).Top

'Label1(21).Top = Me.Height - 1935
'cboCajas.Top = Label1(21).Top
'Label1(22).Top = Me.Height - 1575
'cboUsuarios.Top = Label1(22).Top


txtDocCod.Top = Me.Height - 4815
txtDocName.Top = txtDocCod.Top
Label1(7).Top = txtDocCod.Top
Label1(24).Top = txtDocCod.Top

txtTransaccion.Top = txtDocCod.Top
FlatScrollBar.Top = txtDocCod.Top

btnMain.Item(0).Top = txtDocCod.Top - 60
btnMain.Item(1).Top = btnMain.Item(0).Top

lswConceptos.Height = cboEstado.Top - lswConceptos.Top - 200  '  7800

'chkDocumentosBloqueados.Top = cboUsuarios.Top + 360

vGridFormasPago.Width = tcMain.Width - 290
vGridAfectaciones.Width = tcMain.Width - 290
vgridAsiento.Width = tcMain.Width - 490

txtDetalle.Width = tcMain.Width - (txtDetalle.Left + 250)

End Sub








Private Sub lswDocConsulta_DblClick()

If lswDocConsulta.ListItems.Count = 0 Then Exit Sub

fraDocumento.Visible = False

txtDocCod.Text = lswDocConsulta.SelectedItem
txtDocName.Text = lswDocConsulta.SelectedItem.SubItems(1)
Call sbDocUltDocumento

End Sub




Private Sub tcMain_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
Dim strSQL As String


Select Case Item.Index
   Case 0 'Documento
     If vCodigo <> "" Then Call sbCargaDocumento(vCodigo, vTipoDocu)
   Case 1 'Asientos
     If vCodigo <> "" Then Call sbCargaAsiento(vCodigo, vTipoDocu)
   Case 2 'Afectaciones
     If vCodigo <> "" Then Call sbAfectaciones(vCodigo, vTipoDocu)
   Case 3 'Seguimiento
      If vCodigo <> "" Then Call sbSeguimiento(vCodigo, vTipoDocu)
   Case 4 'Formas de Pago
      If vCodigo <> "" Then Call sbCargaFormasPago(vCodigo, vTipoDocu)
End Select

End Sub

Private Sub TimerX_Timer()

TimerX.Interval = 0
TimerX.Enabled = False

Call sbTipoDocs_Load
Call sbConceptos_Load

strSQL = "select cod_caja as 'IdX', rtrim(Descripcion) as 'itmx' from cajas_definicion  where activa = 1"
Call sbCbo_Llena_New(cboCajas, strSQL, True, True)

strSQL = "select cod_forma_pago as 'IdX',  rtrim(Descripcion) as 'itmx' from sif_Formas_Pago  where activa = 1"
Call sbCbo_Llena_New(cboFormasPago, strSQL, True, True)


txtReporteUsuario.Text = glogon.Usuario

cboBanco.Clear
strSQL = "select id_banco as Idx,descripcion as Itmx from Tes_Bancos where estado = 'A'"
Call sbCbo_Llena_New(cboBanco, strSQL, True, True)

Call sbLimpia

End Sub

Private Sub tlb_ButtonClick(ByVal Button As MSComctlLib.Button)

Select Case Button.Key
  Case "Buscar"
   
    Call sbLimpia
    Call sbBuscar
    Call vGrid_Click(1, 1)
  Case "Reporte"
    fraReportes.Top = btnInforme.Top + 480
    fraReportes.Left = btnInforme.Left + 920
    fraReportes.Visible = True
End Select

End Sub



Private Sub sbReversar(pTipo As String, pDocumento As String)
Dim strSQL As String, i As Byte

On Error GoTo vError


i = MsgBox("Esta seguro(a) que desea reversar esta transacción?", vbYesNo)
If i = vbNo Then
   Exit Sub
End If

Me.MousePointer = vbHourglass

'If GLOBALES.SysPlanPagos = 1 Then
'   strSQL = "exec spSIFDocsReversaCrdPlanPago '" & pTipo & "','" & pDocumento & "','" & glogon.Usuario & "'"
'Else
'   strSQL = "exec spSIFDocsReversaCrd '" & pTipo & "','" & pDocumento & "'"
'End If


strSQL = "exec spSIFDocsReversaMain '" & pTipo & "','" & pDocumento & "','" & glogon.Usuario & "'"
Call ConectionExecute(strSQL)

Me.MousePointer = vbDefault
MsgBox "Documento Reversado Satisfactoriamente...", vbExclamation

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical


End Sub




Private Function fxReportesFiltros() As String
Dim vFiltro As String
Dim vCadena As String, iCantidad As Integer, i As Integer


On Error GoTo vError

iCantidad = 0
vFiltro = ""
strDetalle = ""

If txtBuscarPor <> "" Then
    Select Case Mid(cbo, 1, 2)
        Case "01"
            vFiltro = vFiltro & "{vSIFDocumentos.cliente_Identificacion}  like '*" & txtBuscarPor.Text & "*' "
            strDetalle = "Id. cliente .: " & txtBuscarPor.Text
        Case "02"
          
            vFiltro = vFiltro & "{vSIFDocumentos.Cliente_Nombre}  like '*" & txtBuscarPor.Text & "*' "
            strDetalle = "Nombre cliente .: " & txtBuscarPor.Text
    End Select
End If 'txtBuscarPor

'Lista de Documentos
vCadena = ""
For i = 1 To lswDocumentos.ListItems.Count
  If lswDocumentos.ListItems.Item(i).Checked Then
    If vCadena <> "" Then
        vCadena = vCadena & ","
    End If
    vCadena = vCadena & "'" & lswDocumentos.ListItems.Item(i).Tag & "'"
    iCantidad = iCantidad + 1
  End If
Next i

If iCantidad > 2 Then
  strDetalle = strDetalle & " - Doc .: Filtrados"
ElseIf iCantidad = 0 Then
  strDetalle = strDetalle & " - Doc .: Todos"
Else
   strDetalle = strDetalle & " - Doc.: " & Mid(vCadena, 28, Len(vCadena))
End If

iCantidad = 0
  If vFiltro <> Empty And vFiltro <> " and " Then vFiltro = vFiltro & " and "
  vFiltro = vFiltro & "{vSIFDocumentos.Tipo_Documento} in [" & vCadena & "] "

'Lista de Conceptos
vCadena = ""
For i = 1 To lswConceptos.ListItems.Count
  If lswConceptos.ListItems.Item(i).Checked Then
    If vCadena <> "" Then
        vCadena = vCadena & ","
    End If
    vCadena = vCadena & "'" & lswConceptos.ListItems.Item(i).Tag & "'"
    iCantidad = iCantidad + 1
  End If
Next i

If iCantidad > 2 Then
  strDetalle = strDetalle & " - Conceptos .: Filtrados"
ElseIf iCantidad = 0 Then
  strDetalle = strDetalle & " - Conceptos .: Todos"
Else
   strDetalle = strDetalle & " - Concepto.:" & Mid(vCadena, 28, Len(vCadena))
End If

If vFiltro <> Empty And vFiltro <> " and " Then vFiltro = vFiltro & " and "
 vFiltro = vFiltro & "{vSIFDocumentos.Cod_Concepto} in [" & vCadena & "] "

Select Case cboFechas.Text
  Case "Registro"
    
    If vFiltro <> Empty And vFiltro <> " and " Then vFiltro = vFiltro & " and "
    vFiltro = vFiltro & "cdate({vSIFDocumentos.registro_fecha}) in Date(" & Format(dtpInicio.Value, "yyyy,mm,dd")
    vFiltro = vFiltro & ") to Date (" & Format(dtpCorte.Value, "yyyy,mm,dd") & ")"
    
    strDetalle = strDetalle & " - Fecha Registro.: desde " & Format(dtpInicio.Value, "dd/mm/yyyy") & " Hasta " & Format(dtpCorte.Value, "dd/mm/yyyy")
    
  Case "Anulación"
    
    If vFiltro <> Empty And vFiltro <> " and " Then vFiltro = vFiltro & " and "
    vFiltro = vFiltro & "cdate({vSIFDocumentos.anulacion_fecha}) in Date(" & Format(dtpInicio.Value, "yyyy,mm,dd")
    vFiltro = vFiltro & ") to Date (" & Format(dtpCorte.Value, "yyyy,mm,dd") & ")"

    strDetalle = strDetalle & " - Fecha Anulación.: desde " & Format(dtpInicio.Value, "dd/mm/yyyy") & " Hasta " & Format(dtpCorte.Value, "dd/mm/yyyy")
    
  Case "Traslado"
    If vFiltro <> Empty And vFiltro <> " and " Then vFiltro = vFiltro & " and "
    vFiltro = vFiltro & "cdate({vSIFDocumentos.traspaso_fecha}) in Date(" & Format(dtpInicio.Value, "yyyy,mm,dd")
    vFiltro = vFiltro & ") to Date (" & Format(dtpCorte.Value, "yyyy,mm,dd") & ")"
    
    strDetalle = strDetalle & " - Fecha Traslado.: desde " & Format(dtpInicio.Value, "dd/mm/yyyy") & " Hasta " & Format(dtpCorte.Value, "dd/mm/yyyy")
  Case Else
    strDetalle = strDetalle & " - Todas las Fechas"
End Select

If cboCajas.Text <> "TODOS" Then
    
    If vFiltro <> Empty And vFiltro <> " and " Then vFiltro = vFiltro & " and "
    vFiltro = vFiltro & "{vSIFDocumentos.cod_caja} = '" & cboCajas.ItemData(cboCajas.ListIndex) & "' "
    
    If IsNumeric(txtCajasApertura.Text) Then
        If CLng(txtCajasApertura.Text) > 0 Then vFiltro = vFiltro & " and {vSIFDocumentos.cod_Apertura} = " & txtCajasApertura.Text
    End If
    
    strDetalle = strDetalle & " - Caja .:" & cboCajas.ItemData(cboCajas.ListIndex) & ".Ap." & txtCajasApertura.Text
Else
    strDetalle = strDetalle & " - Todas las cajas.ap." & txtCajasApertura.Text
End If

If cboUsuarios.Text <> "" And cboUsuarios.Text <> "TODOS" Then
   If vFiltro <> Empty And vFiltro <> " and" Then vFiltro = vFiltro & " and "
   vFiltro = vFiltro & "{vSIFDocumentos.registro_usuario} = '" & cboUsuarios.Text & "' "
    strDetalle = strDetalle & " - Usuario .:" & cboUsuarios.Text

Else
    strDetalle = strDetalle & " - Todos los Usuarios"
End If


If Trim(txtUsuarioRegistra.Text) <> "" Then
   If vFiltro <> Empty And vFiltro <> " and" Then vFiltro = vFiltro & " and "
      vFiltro = vFiltro & "{vSIFDocumentos.registro_usuario} = '" & txtUsuarioRegistra.Text & "'"
End If


If Trim(txtNoTransaccion.Text) <> "" Then
   If vFiltro <> Empty And vFiltro <> " and" Then vFiltro = vFiltro & " and "
      vFiltro = vFiltro & "{vSIFDocumentos.Cod_Transaccion} = '" & txtNoTransaccion.Text & "'"
End If

If Trim(txtNoDocumento.Text) <> "" Then
   If vFiltro <> Empty And vFiltro <> " and" Then vFiltro = vFiltro & " and "
      vFiltro = vFiltro & "{vSIFDocumentos.Documento} = '" & txtNoDocumento.Text & "'"
End If


        


Select Case cboEstado.Text
  Case "Impreso"
     If vFiltro <> Empty And vFiltro <> " and" Then vFiltro = vFiltro & " and "
     vFiltro = vFiltro & "{vSIFDocumentos.estado}  in ['I','E'] "
  
  Case "Pendiente"
     If vFiltro <> Empty And vFiltro <> " and" Then vFiltro = vFiltro & " and "
     vFiltro = vFiltro & "{vSIFDocumentos.estado}  = 'P' "
  Case Else

End Select
strDetalle = strDetalle & " - Estado..:" & cboEstado.Text

strDetalle = strDetalle & " - Bloqueados..:" & chkDocumentosBloqueados.Value


If vFiltro <> Empty And vFiltro <> " and" Then vFiltro = vFiltro & " and "
vFiltro = vFiltro & "{vSIFDocumentos.Traslado_Bloqueo}  = " & chkDocumentosBloqueados.Value


     
fxReportesFiltros = vFiltro


Exit Function

vError:
  fxReportesFiltros = ""
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical




End Function


Private Sub txtCuenta_KeyDown(KeyCode As Integer, Shift As Integer)

gCuenta = ""
txtCuenta.Text = ""

If KeyCode = vbKeyF4 Then
  frmCntX_ConsultaCuentas.Show vbModal
  
  If gCuenta <> "" Then
      txtCuenta.Text = fxgCntCuentaFormato(True, gCuenta, 0)
  End If
End If


End Sub

Private Sub txtDocCod_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then
   txtDocName.SetFocus
End If

If KeyCode = vbKeyF4 Then
   fraDocumento.Top = txtDocCod.Top
   fraDocumento.Visible = True
   
   txtDocCodConsulta.Text = ""
   txtDocNameConsulta.Text = ""
   
   txtDocCodConsulta.SetFocus
   lswDocConsulta.ListItems.Clear
   
   Call sbDocConsulta
End If

End Sub


Private Sub sbDocCodName()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError
Me.MousePointer = vbHourglass

strSQL = "select Descripcion from SIF_Documentos where Activo = 1" _
       & " and Tipo_Documento like '" & txtDocCod.Text & "'"

Call OpenRecordSet(rs, strSQL)
If Not rs.EOF And Not rs.BOF Then
   txtDocName.Text = Trim(rs!Descripcion)
Else
   txtDocName.Text = "! >> Tipo de Documento no encontrado << !"
End If
rs.Close

Me.MousePointer = vbDefault
Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
  
End Sub



Private Sub sbDocConsulta()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListItem

On Error GoTo vError
Me.MousePointer = vbHourglass

strSQL = "select Tipo_Documento,Descripcion from SIF_Documentos where Activo = 1"

If Len(txtDocCodConsulta.Text) > 0 Then
   strSQL = strSQL & " and Tipo_Documento like '%" & txtDocCodConsulta.Text & "%'"
End If

If Len(txtDocNameConsulta.Text) > 0 Then
   strSQL = strSQL & " and Descripcion like '%" & txtDocNameConsulta.Text & "%'"
End If

strSQL = strSQL & " order by Tipo_documento"

lswDocConsulta.ListItems.Clear

Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
  Set itmX = lswDocConsulta.ListItems.Add(, , rs!Tipo_Documento)
      itmX.SubItems(1) = rs!Descripcion
  rs.MoveNext
Loop
rs.Close

Me.MousePointer = vbDefault
Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
  
End Sub

Private Sub txtDocCod_LostFocus()
Call sbDocCodName
Call sbDocUltDocumento
End Sub

Private Sub txtDocCodConsulta_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
   txtDocNameConsulta.SetFocus
   Call sbDocConsulta
End If
End Sub

Private Sub txtDocName_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then
   txtTransaccion.SetFocus
End If

If KeyCode = vbKeyF4 Then
   fraDocumento.Top = txtDocCod.Top
   fraDocumento.Left = tcMain.Left
   fraDocumento.Visible = True
   
   txtDocCodConsulta.Text = ""
   txtDocNameConsulta.Text = ""
   
   txtDocCodConsulta.SetFocus
   lswDocConsulta.ListItems.Clear
   
   Call sbDocConsulta
   
End If

End Sub

Private Sub txtDocNameConsulta_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
   Call sbDocConsulta
End If
End Sub




Private Sub txtFiltraConcepto_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    Call sbConceptos_Load
End If
End Sub



Private Sub txtFiltraTipoDoc_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    Call sbTipoDocs_Load
End If

End Sub

Private Sub txtReporteUsuario_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyF4 Then
  gBusquedas.Convertir = "N"
  gBusquedas.Columna = "Nombre"
  gBusquedas.Orden = "Nombre"
  gBusquedas.Consulta = "select Nombre,descripcion from Usuarios"
  gBusquedas.Filtro = " and estado = 'A'"
  frmBusquedas.Show vbModal
  txtReporteUsuario.Text = gBusquedas.Resultado
End If

End Sub

Private Sub txtTransaccion_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
    
    vCodigo = txtTransaccion
    vTipoDocu = txtDocCod.Text
    
    Select Case tcMain.Selected.Index
      Case 0
        Call sbCargaDocumento(vCodigo, vTipoDocu)
      Case 1
        Call sbCargaAsiento(vCodigo, vTipoDocu)
    End Select
End If

End Sub

Private Sub vGrid_Click(ByVal Col As Long, ByVal Row As Long)
vGrid.Row = Row
vGrid.Col = 1
vCodigo = vGrid.Text
vGrid.Col = 3
vTipoDocu = vGrid.Text

txtTransaccion.Text = vCodigo
txtDocCod.Text = vTipoDocu

Call sbCargaDocumento(vCodigo, vTipoDocu)

End Sub

Private Sub vGrid_DblClick(ByVal Col As Long, ByVal Row As Long)
'Dim frm As Form
'
'If Row <= 0 Then Exit Sub
'If vGrid.MaxRows <= 0 Then Exit Sub
'
'vGrid.Row = Row
'vGrid.Col = 1
'
'If vGrid.Text = "" Then Exit Sub
'
' Call sbSIFForms("frmTES_Transacciones")
' For Each frm In Forms
'   If UCase(frm.Name) = UCase("frmTES_Transacciones") Then
'     Call frm.sbTESDocConsulta(vGrid.Text)
'     Exit For
'   End If
' Next frm

End Sub


Private Sub sbCargaAsiento(strCodigo As String, strDocumento As String)
Dim strSQL As String, rs As New ADODB.Recordset
Dim i As Integer


 strSQL = "select isnull(C.Cod_Cuenta_Mask, D.cod_cuenta) as 'COD_CUENTA' , isnull(C.descripcion,'--Cuenta No Existe--') as 'Descripcion',D.Cod_divisa,D.tipo_movimiento,D.monto,D.cod_unidad" _
          & ",U.descripcion as UnidadX,D.cod_centro_costo,X.descripcion as CCX,D.Tipo_Cambio" _
          & ",D.Referencia_01,D.Referencia_02,D.Referencia_03, D.Monto / dbo.fxSys_Tipo_Cambio_Apl(D.Tipo_Cambio) as 'IMPORTE_REAL'" _
          & " from Sif_transacciones_asiento D left join CntX_Cuentas C on D.cod_cuenta = C.cod_cuenta and D.cod_contabilidad = C.cod_contabilidad" _
          & " left join cntx_unidades U on D.cod_unidad = U.cod_unidad and D.cod_contabilidad = U.cod_contabilidad" _
          & " left join cntx_centro_costos X on D.cod_centro_costo = X.cod_centro_costo and D.cod_contabilidad = X.cod_contabilidad" _
          & " where D.cod_transaccion= '" & strCodigo & "' and tipo_Documento = '" & strDocumento & "'" _
          & " order by D.Numero_linea"

Call OpenRecordSet(rs, strSQL)
    
    vgridAsiento.MaxRows = 1
    vgridAsiento.Row = vgridAsiento.MaxRows
    
    
    Do While Not rs.EOF
      vgridAsiento.Row = vgridAsiento.MaxRows
      
      For i = 1 To vgridAsiento.MaxCols
        vgridAsiento.Col = i
        Select Case i
         Case 1
            vgridAsiento.Text = rs!cod_cuenta & ""
            
         Case 2
            vgridAsiento.Text = rs!Cod_Unidad & ""
            vgridAsiento.CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
            vgridAsiento.CellNote = rs!UnidadX & ""
            vgridAsiento.TextTip = TextTipFixed
         
         
         Case 3
            vgridAsiento.Text = rs!Cod_Centro_Costo & ""
            vgridAsiento.CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
            vgridAsiento.CellNote = rs!CCX & ""
            vgridAsiento.TextTip = TextTipFixed
         Case 4
           vgridAsiento.Text = UCase(CStr(rs!Cod_Divisa))
         
         Case 5
            vgridAsiento.Text = UCase(CStr(rs!Descripcion & ""))
         Case 6 'Debitos
           If rs!Tipo_Movimiento = "D" Then
             vgridAsiento.Text = Format(CStr(rs!Monto), "Standard")
           Else
             vgridAsiento.Text = "0.00"
           End If
         Case 7 'Creditos
           If rs!Tipo_Movimiento = "D" Then
             vgridAsiento.Text = "0.00"
           Else
             vgridAsiento.Text = Format(CStr(rs!Monto), "Standard")
           End If
           
         Case 8 'Tipo de Cambio
             vgridAsiento.Text = Format(CStr(rs!Tipo_Cambio), "Standard")
         Case 9 'Importe Real
             vgridAsiento.Text = Format(rs!IMPORTE_REAL, "Standard")
           
         Case 10 'Referencia 01
               vgridAsiento.Text = rs!Referencia_01 & ""
         Case 11 'Referencia 02
               vgridAsiento.Text = rs!Referencia_02 & ""
         Case 12 'Referencia 03
               vgridAsiento.Text = rs!Referencia_03 & ""
        End Select
      Next i
      vgridAsiento.MaxRows = vgridAsiento.MaxRows + 1
      
      rs.MoveNext
    Loop
    rs.Close
    vgridAsiento.MaxRows = vgridAsiento.MaxRows - 1
    
    Call sbSumaDebitosCreditos
    
End Sub

Private Sub sbSumaDebitosCreditos()
Dim x As Integer, TC As Currency
  
On Error GoTo vError
  
  txtDebito = 0
  txtCredito = 0
  For x = 1 To vgridAsiento.MaxRows
     vgridAsiento.Row = x
     TC = 1
      
     vgridAsiento.Col = 6
     txtDebito = CCur(txtDebito) + (CCur(IIf(vgridAsiento.Text = "", 0, vgridAsiento.Text)) * TC)
     vgridAsiento.Col = 7
     txtCredito = CCur(txtCredito) + (CCur(IIf(vgridAsiento.Text = "", 0, vgridAsiento.Text)) * TC)
  Next x
  txtDiferencia = txtDebito - txtCredito
  txtDebito = Format(txtDebito, "Standard")
  txtCredito = Format(txtCredito, "Standard")
  txtDiferencia = Format(txtDiferencia, "Standard")

vError:

End Sub


Private Sub sbCargaDocumento(Codigo As String, Documento As String)
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

strSQL = "select T.tipo_documento,T.cod_transaccion, Docs.Descripcion as 'DocumentoDesc' " _
        & ",isnull(T.cliente_identificacion,'') as identificacion, isnull(T.cliente_nombre,'') as nombre,T.monto,T.registro_fecha" _
        & ",T.cod_Concepto,T.registro_usuario, C.descripcion as concepto,Documento,O.Descripcion as oficina,Ca.Descripcion as 'Caja'" _
        & ",case when T.Estado = 'P' then 'Pendiente' when T.Estado = 'I' then 'Impreso' when T.Estado = 'A' then 'Anulado' end as 'Estado'" _
        & ",Linea1,Linea2,Linea3,Linea4,Linea5,Linea6,Linea7,Linea8,Linea9,Linea10,Linea11,Detalle,isnull(T.TRASLADO_BLOQUEO,0) as 'Bloqueo'" _
        & " from sif_transacciones T" _
        & " inner join sif_conceptos C on T.cod_concepto = C.cod_concepto" _
        & " inner join sif_documentos Docs on T.tipo_documento = Docs.Tipo_Documento" _
        & " inner join sif_oficinas O on T.cod_Oficina = O.cod_oficina" _
        & " left join cajas_definicion Ca on T.cod_caja = Ca.cod_caja" _
        & " where T.tipo_documento = '" & Documento & "' and T.cod_transaccion = '" & Codigo & "'"

Call OpenRecordSet(rs, strSQL)


tcMain.Item(0).Selected = True

If Not rs.EOF Then

   txtDocCod.Text = Trim(rs!Tipo_Documento)
   txtDocName.Text = Trim(rs!DocumentoDesc)
   txtTransaccion.Text = rs!Cod_Transaccion
   
   txtCedula.Text = rs!identificacion
   txtNombre.Text = rs!Nombre
   txtCod_Concepto.Text = rs!cod_concepto
   txtDescConcepto.Text = rs!Concepto
   txtDocumento.Text = rs!Documento & ""
   txtDocumento.Tag = rs!Documento & ""
   txtMonto.Text = Format(rs!Monto, "Standard")
   txtCaja.Text = rs!caja & ""
   txtOficina.Text = rs!oficina
   txtEstado.Text = rs!Estado
   
   txtUsuarioCaja.Text = Trim(rs!Registro_Usuario) & ""
   txtFecha.Text = Trim(rs!registro_Fecha & "")
   
   If Not IsNull(rs!Linea1) Then txtDetalle = rs!Linea1 & vbCrLf
   If Not IsNull(rs!linea2) Then txtDetalle = txtDetalle & rs!linea2 & vbCrLf
   If Not IsNull(rs!linea3) Then txtDetalle = txtDetalle & rs!linea3 & vbCrLf
   If Not IsNull(rs!linea4) Then txtDetalle = txtDetalle & rs!linea4 & vbCrLf
   If Not IsNull(rs!linea5) Then txtDetalle = txtDetalle & rs!linea5 & vbCrLf
   If Not IsNull(rs!linea6) Then txtDetalle = txtDetalle & rs!linea6 & vbCrLf
   If Not IsNull(rs!linea7) Then txtDetalle = txtDetalle & rs!linea7 & vbCrLf
   If Not IsNull(rs!linea8) Then txtDetalle = txtDetalle & rs!linea8 & vbCrLf
   If Not IsNull(rs!linea9) Then txtDetalle = txtDetalle & rs!linea9 & vbCrLf
   If Not IsNull(rs!linea10) Then txtDetalle = txtDetalle & rs!linea10 & vbCrLf
   If Not IsNull(rs!linea11) Then txtDetalle = txtDetalle & rs!linea11 & vbCrLf
   If Not IsNull(rs!Detalle) Then txtDetalle = txtDetalle & rs!Detalle & vbCrLf
   
   chkBloqueado.Value = rs!Bloqueo
   
Else
  Call sbLimpia
End If

rs.Close

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub sbLimpia()
   tcMain.Item(0).Selected = True
   txtCedula = ""
   txtNombre = ""
   txtCod_Concepto = ""
   txtDescConcepto = ""
   txtDocumento = ""
   txtMonto = 0
   txtCaja = ""
   txtOficina = ""
   txtEstado = ""
   txtDetalle = ""
   
End Sub



Private Sub sbSeguimiento(Codigo As String, Documento As String)
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

strSQL = "Select registro_fecha,registro_usuario,traspaso_fecha,traspaso_usuario,anulacion_fecha,anulacion_usuario" _
      & " from sif_transacciones where tipo_documento = '" & Documento & "' and cod_transaccion = '" & Codigo & "'"

Call OpenRecordSet(rs, strSQL)

If Not rs.EOF Then
    txtRegistroUsuario = Trim(IIf(IsNull(rs!Registro_Usuario), "", rs!Registro_Usuario))
    txtRegistroFecha = IIf(IsNull(rs!registro_Fecha), "", rs!registro_Fecha)
    txtTrasladoFecha = IIf(IsNull(rs!traspaso_fecha), "", rs!traspaso_fecha)
    txtTrasladoUsuario = Trim(IIf(IsNull(rs!traspaso_usuario), "", rs!traspaso_usuario))
    txtAnulaFecha = IIf(IsNull(rs!anulacion_fecha), "", rs!anulacion_fecha)
    txtAnulaUsuario = Trim(IIf(IsNull(rs!anulacion_usuario), "", rs!anulacion_usuario))
Else
    txtRegistroUsuario = ""
    txtRegistroFecha = ""
    txtTrasladoFecha = ""
    txtTrasladoUsuario = ""
    txtAnulaFecha = ""
    txtAnulaUsuario = ""
End If
rs.Close

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub sbCargaFormasPago(Codigo As String, Documento As String)
Dim strSQL As String

strSQL = "select F.DESCRIPCION, P.Monto, P.COD_DIVISA, P.TIPO_CAMBIO, P.monto / dbo.fxSys_Tipo_Cambio_Apl(P.TIPO_CAMBIO) " _
       & " , case when F.TIPO = 'C' then 'CK.: ' + P.CHEQUE_NUMERO + ' - Emisor.: ' + P.CHEQUE_EMISOR" _
       & "    when F.TIPO = 'D' then 'DOC.: ' + P.NUM_REFERENCIA" _
       & "    when F.TIPO = 'T' then 'TARJ.: ' + P.TARJETA_NUMERO + ' AUT.:' + P.TARJETA_AUTORIZACION + '  TIPO..:' + P.COD_TARJETA" _
       & "    when F.TIPO = 'S' then  S.DOC_TIPO + ' - ' + S.DOC_NUMERO + '     (Id.: ' + CONVERT(VARCHAR(20), S.LINEA) + ') '" _
       & "     else P.NUM_REFERENCIA end" _
       & "  ,ISNULL(P.OBSERVACIONES,'') AS 'NOTAS'" _
       & "  from SIF_TRANSACCIONES_PAGO P inner join SIF_FORMAS_PAGO F on P.COD_FORMA_PAGO = F.COD_FORMA_PAGO" _
       & " left join CAJAS_SALDO_FAVOR S on P.SALDO_FAVOR_ID = S.Linea" _
        & " where P.tipo_documento = '" & Documento & "' and P.cod_transaccion = '" & Codigo & "' order by P.cod_linea"

Call sbCargaGrid(vGridFormasPago, 7, strSQL)
End Sub


Private Sub vGridAfectaciones_SheetChanged(ByVal OldSheet As Integer, ByVal NewSheet As Integer)

vGridAfectaciones.Sheet = NewSheet
Call sbAfectaciones(txtTransaccion.Text, txtDocCod.Text)

End Sub

