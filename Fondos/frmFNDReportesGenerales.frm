VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.Controls.v24.0.0.ocx"
Begin VB.Form frmFNDReportesGenerales 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reportes Generales"
   ClientHeight    =   7485
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10425
   Icon            =   "frmFNDReportesGenerales.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7485
   ScaleWidth      =   10425
   Begin VB.Timer TimerX 
      Interval        =   10
      Left            =   480
      Top             =   120
   End
   Begin XtremeSuiteControls.CheckBox chkPlanes 
      Height          =   210
      Left            =   8640
      TabIndex        =   24
      Top             =   500
      Width           =   210
      _Version        =   1572864
      _ExtentX        =   370
      _ExtentY        =   370
      _StockProps     =   79
      ForeColor       =   16777215
      BackColor       =   12582912
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
   Begin XtremeSuiteControls.GroupBox GroupBox1 
      Height          =   975
      Left            =   120
      TabIndex        =   14
      Top             =   6360
      Width           =   10215
      _Version        =   1572864
      _ExtentX        =   18013
      _ExtentY        =   1714
      _StockProps     =   79
      BackColor       =   16777215
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      BorderStyle     =   1
      Begin XtremeSuiteControls.PushButton btnReporte 
         Height          =   492
         Left            =   6840
         TabIndex        =   33
         Top             =   240
         Width           =   1572
         _Version        =   1572864
         _ExtentX        =   2773
         _ExtentY        =   868
         _StockProps     =   79
         Caption         =   "Reporte"
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
         Picture         =   "frmFNDReportesGenerales.frx":000C
      End
      Begin XtremeSuiteControls.PushButton btnCubo 
         Height          =   492
         Left            =   8400
         TabIndex        =   34
         Top             =   240
         Width           =   1572
         _Version        =   1572864
         _ExtentX        =   2773
         _ExtentY        =   868
         _StockProps     =   79
         Caption         =   "Cubo"
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
         Picture         =   "frmFNDReportesGenerales.frx":07C8
      End
      Begin VB.Label lblStatus 
         BackStyle       =   0  'Transparent
         Caption         =   "Este proceso puede tardar varios minutos, espere el mensaje de proceso concluido."
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   492
         Left            =   0
         TabIndex        =   15
         Top             =   240
         Visible         =   0   'False
         Width           =   5292
      End
   End
   Begin XtremeSuiteControls.TabControl tcFiltros 
      Height          =   5055
      Left            =   4320
      TabIndex        =   3
      Top             =   1320
      Width           =   5895
      _Version        =   1572864
      _ExtentX        =   10398
      _ExtentY        =   8916
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
      Item(0).Caption =   "Filtro: General"
      Item(0).ControlCount=   16
      Item(0).Control(0)=   "cboOperadora"
      Item(0).Control(1)=   "Label2(5)"
      Item(0).Control(2)=   "cboEstado"
      Item(0).Control(3)=   "cboFechaBase"
      Item(0).Control(4)=   "chkFechas"
      Item(0).Control(5)=   "Label2(1)"
      Item(0).Control(6)=   "Label2(2)"
      Item(0).Control(7)=   "chkResumen"
      Item(0).Control(8)=   "dtpInicio"
      Item(0).Control(9)=   "dtpCorte"
      Item(0).Control(10)=   "Label2(3)"
      Item(0).Control(11)=   "Label2(4)"
      Item(0).Control(12)=   "chkContratosSaldos"
      Item(0).Control(13)=   "chkDeduce"
      Item(0).Control(14)=   "cboDivisa"
      Item(0).Control(15)=   "Label2(14)"
      Item(1).Caption =   "Filtro: Movimientos"
      Item(1).ControlCount=   16
      Item(1).Control(0)=   "txtCedula"
      Item(1).Control(1)=   "txtUsuario"
      Item(1).Control(2)=   "txtNDocRef"
      Item(1).Control(3)=   "txtNDoc"
      Item(1).Control(4)=   "cboTipoDoc"
      Item(1).Control(5)=   "cboConcepto"
      Item(1).Control(6)=   "Label2(12)"
      Item(1).Control(7)=   "Label2(11)"
      Item(1).Control(8)=   "Label2(10)"
      Item(1).Control(9)=   "Label2(9)"
      Item(1).Control(10)=   "Label2(8)"
      Item(1).Control(11)=   "Label2(7)"
      Item(1).Control(12)=   "cboInstitucion"
      Item(1).Control(13)=   "Label2(6)"
      Item(1).Control(14)=   "Label2(13)"
      Item(1).Control(15)=   "cboEPersona"
      Begin XtremeSuiteControls.ComboBox cboOperadora 
         Height          =   312
         Left            =   1440
         TabIndex        =   4
         Top             =   480
         Width           =   4332
         _Version        =   1572864
         _ExtentX        =   7646
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
         Left            =   1440
         TabIndex        =   6
         Top             =   1200
         Width           =   2292
         _Version        =   1572864
         _ExtentX        =   4048
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
      Begin XtremeSuiteControls.ComboBox cboFechaBase 
         Height          =   312
         Left            =   1440
         TabIndex        =   7
         Top             =   1560
         Width           =   2292
         _Version        =   1572864
         _ExtentX        =   4048
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
         Height          =   312
         Left            =   2280
         TabIndex        =   10
         Top             =   1920
         Width           =   1452
         _Version        =   1572864
         _ExtentX        =   2561
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
         Height          =   312
         Left            =   2280
         TabIndex        =   11
         Top             =   2280
         Width           =   1452
         _Version        =   1572864
         _ExtentX        =   2561
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
      Begin XtremeSuiteControls.ComboBox cboTipoDoc 
         Height          =   312
         Left            =   -67960
         TabIndex        =   22
         Top             =   1440
         Visible         =   0   'False
         Width           =   3732
         _Version        =   1572864
         _ExtentX        =   6588
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
         Height          =   312
         Left            =   -67960
         TabIndex        =   23
         Top             =   1800
         Visible         =   0   'False
         Width           =   3732
         _Version        =   1572864
         _ExtentX        =   6588
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
      Begin XtremeSuiteControls.CheckBox chkFechas 
         Height          =   252
         Left            =   3960
         TabIndex        =   25
         Top             =   1560
         Width           =   972
         _Version        =   1572864
         _ExtentX        =   1714
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
         Appearance      =   16
      End
      Begin XtremeSuiteControls.CheckBox chkResumen 
         Height          =   252
         Left            =   1560
         TabIndex        =   26
         Top             =   3480
         Width           =   2172
         _Version        =   1572864
         _ExtentX        =   3831
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Informe Resumen?"
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
         Appearance      =   17
         Alignment       =   1
      End
      Begin XtremeSuiteControls.FlatEdit txtNDoc 
         Height          =   312
         Left            =   -67960
         TabIndex        =   35
         Top             =   2280
         Visible         =   0   'False
         Width           =   3732
         _Version        =   1572864
         _ExtentX        =   6583
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
      Begin XtremeSuiteControls.FlatEdit txtNDocRef 
         Height          =   312
         Left            =   -67960
         TabIndex        =   36
         Top             =   2640
         Visible         =   0   'False
         Width           =   3732
         _Version        =   1572864
         _ExtentX        =   6583
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
      Begin XtremeSuiteControls.FlatEdit txtUsuario 
         Height          =   312
         Left            =   -67960
         TabIndex        =   37
         Top             =   3000
         Visible         =   0   'False
         Width           =   3732
         _Version        =   1572864
         _ExtentX        =   6583
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
      Begin XtremeSuiteControls.FlatEdit txtCedula 
         Height          =   312
         Left            =   -67960
         TabIndex        =   38
         Top             =   3360
         Visible         =   0   'False
         Width           =   3732
         _Version        =   1572864
         _ExtentX        =   6583
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
      Begin XtremeSuiteControls.CheckBox chkContratosSaldos 
         Height          =   252
         Left            =   0
         TabIndex        =   41
         Top             =   2760
         Width           =   3732
         _Version        =   1572864
         _ExtentX        =   6583
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Mostrar Contratos con Aportes?"
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
         Appearance      =   17
         Alignment       =   1
      End
      Begin XtremeSuiteControls.CheckBox chkDeduce 
         Height          =   252
         Left            =   0
         TabIndex        =   43
         Top             =   3120
         Width           =   3732
         _Version        =   1572864
         _ExtentX        =   6583
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Contratos sin Deducción Activa?"
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
         Appearance      =   17
         Alignment       =   1
      End
      Begin XtremeSuiteControls.ComboBox cboInstitucion 
         Height          =   312
         Left            =   -67960
         TabIndex        =   44
         Top             =   480
         Visible         =   0   'False
         Width           =   3732
         _Version        =   1572864
         _ExtentX        =   6588
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
      Begin XtremeSuiteControls.ComboBox cboEPersona 
         Height          =   312
         Left            =   -67960
         TabIndex        =   46
         Top             =   840
         Visible         =   0   'False
         Width           =   3732
         _Version        =   1572864
         _ExtentX        =   6588
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
      Begin XtremeSuiteControls.ComboBox cboDivisa 
         Height          =   312
         Left            =   1440
         TabIndex        =   48
         Top             =   840
         Width           =   2292
         _Version        =   1572864
         _ExtentX        =   4048
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
      Begin VB.Label Label2 
         Caption         =   "Divisa"
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
         Index           =   14
         Left            =   240
         TabIndex        =   49
         Top             =   840
         Width           =   972
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Estado Persona"
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
         Left            =   -69520
         TabIndex        =   47
         Top             =   840
         Visible         =   0   'False
         Width           =   1572
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
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
         Index           =   6
         Left            =   -69520
         TabIndex        =   45
         Top             =   480
         Visible         =   0   'False
         Width           =   1572
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
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
         Height          =   252
         Index           =   7
         Left            =   -69520
         TabIndex        =   21
         Top             =   1440
         Visible         =   0   'False
         Width           =   972
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
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
         Index           =   8
         Left            =   -69520
         TabIndex        =   20
         Top             =   1800
         Visible         =   0   'False
         Width           =   972
      End
      Begin VB.Label Label2 
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
         Height          =   252
         Index           =   9
         Left            =   -69520
         TabIndex        =   19
         Top             =   2280
         Visible         =   0   'False
         Width           =   1332
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "No. Doc. Ref."
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
         Left            =   -69520
         TabIndex        =   18
         Top             =   2640
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
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
         Height          =   252
         Index           =   11
         Left            =   -69520
         TabIndex        =   17
         Top             =   3000
         Visible         =   0   'False
         Width           =   1572
      End
      Begin VB.Label Label2 
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
         Height          =   252
         Index           =   12
         Left            =   -69520
         TabIndex        =   16
         Top             =   3360
         Visible         =   0   'False
         Width           =   1572
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Corte"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   312
         Index           =   4
         Left            =   1440
         TabIndex        =   13
         Top             =   2280
         Width           =   732
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Inicio"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   312
         Index           =   3
         Left            =   1440
         TabIndex        =   12
         Top             =   1920
         Width           =   732
      End
      Begin VB.Label Label2 
         Caption         =   "Fecha Base"
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
         Left            =   240
         TabIndex        =   9
         Top             =   1560
         Width           =   972
      End
      Begin VB.Label Label2 
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
         Index           =   1
         Left            =   240
         TabIndex        =   8
         Top             =   1200
         Width           =   972
      End
      Begin VB.Label Label2 
         Caption         =   "Operadora"
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
         Left            =   240
         TabIndex        =   5
         Top             =   480
         Width           =   1092
      End
   End
   Begin XtremeSuiteControls.TabControl tcMain 
      Height          =   5055
      Left            =   120
      TabIndex        =   2
      Top             =   1320
      Width           =   4215
      _Version        =   1572864
      _ExtentX        =   7435
      _ExtentY        =   8916
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
      ItemCount       =   1
      Item(0).Caption =   "Informes:"
      Item(0).ControlCount=   12
      Item(0).Control(0)=   "opt(5)"
      Item(0).Control(1)=   "opt(4)"
      Item(0).Control(2)=   "opt(3)"
      Item(0).Control(3)=   "opt(0)"
      Item(0).Control(4)=   "opt(2)"
      Item(0).Control(5)=   "opt(1)"
      Item(0).Control(6)=   "opt(6)"
      Item(0).Control(7)=   "opt(7)"
      Item(0).Control(8)=   "opt(8)"
      Item(0).Control(9)=   "opt(9)"
      Item(0).Control(10)=   "opt(10)"
      Item(0).Control(11)=   "opt(11)"
      Begin XtremeSuiteControls.RadioButton opt 
         Height          =   252
         Index           =   0
         Left            =   360
         TabIndex        =   27
         Top             =   480
         Width           =   3372
         _Version        =   1572864
         _ExtentX        =   5948
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Movimientos"
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
         Value           =   -1  'True
      End
      Begin XtremeSuiteControls.RadioButton opt 
         Height          =   252
         Index           =   1
         Left            =   360
         TabIndex        =   28
         Top             =   840
         Width           =   3372
         _Version        =   1572864
         _ExtentX        =   5948
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Contratos General"
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
      End
      Begin XtremeSuiteControls.RadioButton opt 
         Height          =   252
         Index           =   2
         Left            =   360
         TabIndex        =   29
         Top             =   1200
         Width           =   3372
         _Version        =   1572864
         _ExtentX        =   5948
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Liquidaciones"
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
      End
      Begin XtremeSuiteControls.RadioButton opt 
         Height          =   252
         Index           =   3
         Left            =   360
         TabIndex        =   30
         Top             =   1560
         Width           =   3372
         _Version        =   1572864
         _ExtentX        =   5948
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Contratos C.D.P."
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
      End
      Begin XtremeSuiteControls.RadioButton opt 
         Height          =   252
         Index           =   4
         Left            =   360
         TabIndex        =   31
         Top             =   1920
         Width           =   3372
         _Version        =   1572864
         _ExtentX        =   5948
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Retiros parciales con aportes en Cero"
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
      End
      Begin XtremeSuiteControls.RadioButton opt 
         Height          =   252
         Index           =   5
         Left            =   360
         TabIndex        =   32
         Top             =   2280
         Width           =   3372
         _Version        =   1572864
         _ExtentX        =   5948
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Movimientos x Documento"
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
      End
      Begin XtremeSuiteControls.RadioButton opt 
         Height          =   255
         Index           =   6
         Left            =   360
         TabIndex        =   50
         Top             =   2640
         Width           =   3375
         _Version        =   1572864
         _ExtentX        =   5948
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Saldos en Negativo"
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
      End
      Begin XtremeSuiteControls.RadioButton opt 
         Height          =   255
         Index           =   7
         Left            =   360
         TabIndex        =   51
         Top             =   3000
         Width           =   3375
         _Version        =   1572864
         _ExtentX        =   5948
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "CashBack Generados"
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
      End
      Begin XtremeSuiteControls.RadioButton opt 
         Height          =   255
         Index           =   8
         Left            =   360
         TabIndex        =   52
         Top             =   3360
         Width           =   3375
         _Version        =   1572864
         _ExtentX        =   5948
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "CashBack Liquidados"
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
      End
      Begin XtremeSuiteControls.RadioButton opt 
         Height          =   255
         Index           =   9
         Left            =   360
         TabIndex        =   53
         Top             =   3720
         Width           =   3375
         _Version        =   1572864
         _ExtentX        =   5948
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "CashBack Vencidos"
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
      End
      Begin XtremeSuiteControls.RadioButton opt 
         Height          =   255
         Index           =   10
         Left            =   360
         TabIndex        =   54
         Top             =   4080
         Width           =   3375
         _Version        =   1572864
         _ExtentX        =   5948
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "IVA Trasfronterizo"
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
      End
      Begin XtremeSuiteControls.RadioButton opt 
         Height          =   255
         Index           =   11
         Left            =   360
         TabIndex        =   55
         Top             =   4440
         Width           =   3375
         _Version        =   1572864
         _ExtentX        =   5948
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Proyección de vencimiento de Cupones"
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
      End
   End
   Begin MSComCtl2.FlatScrollBar FlatScrollBar 
      Height          =   255
      Left            =   8040
      TabIndex        =   1
      Top             =   480
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   450
      _Version        =   393216
      Arrows          =   65536
      Orientation     =   1638401
   End
   Begin XtremeSuiteControls.FlatEdit txtCodigo 
      Height          =   312
      Left            =   1800
      TabIndex        =   39
      Top             =   480
      Width           =   1212
      _Version        =   1572864
      _ExtentX        =   2138
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
      Left            =   3000
      TabIndex        =   40
      Top             =   480
      Width           =   4932
      _Version        =   1572864
      _ExtentX        =   8700
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
   Begin XtremeSuiteControls.Label Label1 
      Height          =   252
      Left            =   9000
      TabIndex        =   42
      Top             =   480
      Width           =   972
      _Version        =   1572864
      _ExtentX        =   1714
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Todos"
      ForeColor       =   16777215
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
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Plan"
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
      Left            =   1800
      TabIndex        =   0
      Top             =   240
      Width           =   1092
   End
   Begin VB.Image imgBanner 
      Height          =   1212
      Left            =   0
      Top             =   0
      Width           =   15732
   End
End
Attribute VB_Name = "frmFNDReportesGenerales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vScroll As Boolean

Private Sub btnCubo_Click()
    lblStatus.Visible = True
    Call sbCubo

End Sub

Private Sub btnReporte_Click()

    lblStatus.Visible = False
    Call sbReportes
 
End Sub

Private Sub cboEstado_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then chkFechas.SetFocus
End Sub


Private Sub cboOperadora_Click()
txtCodigo_LostFocus
End Sub

Private Sub cboOperadora_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then txtCodigo.SetFocus
End Sub


Private Sub chkFechas_Click()

cboFechaBase.Enabled = IIf((chkFechas.Value = vbChecked), False, True)

dtpCorte.Enabled = cboFechaBase.Enabled
dtpInicio.Enabled = cboFechaBase.Enabled

End Sub


Private Sub chkPlanes_Click()
txtCodigo.Enabled = IIf((chkPlanes.Value = vbChecked), False, True)

txtDescripcion.Enabled = txtCodigo.Enabled
End Sub

Private Sub sbReportes()
Dim strSQL As String

Me.MousePointer = vbHourglass

On Error GoTo vError

With frmContenedor.Crt
  .Reset
  .WindowShowGroupTree = True
  .WindowShowPrintSetupBtn = True
  .WindowShowRefreshBtn = True
  .WindowShowSearchBtn = True
  .WindowState = crptMaximized
  .WindowTitle = "Reportes del Módulo de Fondos"

  .Connect = glogon.ConectRPT
    
    Select Case True
      Case opt.Item(0).Value 'Movimientos
      
         If Mid(cboEstado.Text, 1, 1) <> "T" Then strSQL = "{FND_CONTRATOS.ESTADO} = '" & Mid(cboEstado.Text, 1, 1) & "'"
         
     
            If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
            strSQL = strSQL & "{FND_OPERADORAS.COD_OPERADORA} = " & cboOperadora.ItemData(cboOperadora.ListIndex)
            
         If chkPlanes.Value = vbUnchecked Then
            If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
            strSQL = strSQL & "{FND_CONTRATOS.COD_PLAN} ='" & Trim(txtCodigo) & "'"
         End If
                     
         If chkFechas.Value = vbUnchecked Then
            If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
            strSQL = strSQL & "{FND_CONTRATOS_DETALLE.FECHA} in Date(" & Format(dtpInicio.Value, "yyyy,mm,dd") _
                   & ") to Date(" & Format(dtpCorte.Value, "yyyy,mm,dd") & ")"
         End If
                     
                     
         'Filtros Especiales
         If Trim(txtUsuario.Text) <> "" Then
            If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
            strSQL = strSQL & "{FND_CONTRATOS_DETALLE.USUARIO} = '" & Trim(txtUsuario.Text) & "'"
         End If
                     
         If Trim(txtCedula.Text) <> "" Then
            If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
            strSQL = strSQL & "{FND_CONTRATOS.CEDULA} = '" & Trim(txtCedula.Text) & "'"
         End If
                     
         If Trim(txtNDoc.Text) <> "" Then
            If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
            strSQL = strSQL & "{FND_CONTRATOS_DETALLE.NCON} = '" & Trim(txtNDoc.Text) & "'"
         End If
                     
         If Trim(txtNDocRef.Text) <> "" Then
            If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
            strSQL = strSQL & "{FND_CONTRATOS_DETALLE.REF_01} = '" & Trim(txtNDocRef.Text) & "'"
         End If
         
         If cboTipoDoc.Text <> "TODOS" Then
            If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
            strSQL = strSQL & "{FND_CONTRATOS_DETALLE.TCON} = '" & cboTipoDoc.ItemData(cboTipoDoc.ListIndex) & "'"
         End If
         
         If cboConcepto.Text <> "TODOS" Then
            If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
            strSQL = strSQL & "{FND_CONTRATOS_DETALLE.COD_CONCEPTO} = '" & cboConcepto.ItemData(cboConcepto.ListIndex) & "'"
         End If
         
         
         If cboInstitucion.Text <> "TODOS" Then
            If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
            strSQL = strSQL & "{SOCIOS.COD_INSTITUCION} = " & cboInstitucion.ItemData(cboInstitucion.ListIndex)
         End If
                  
                  
         If cboDivisa.Text <> "TODOS" Then
            If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
            strSQL = strSQL & "{FND_PLANES.COD_MONEDA} = '" & cboDivisa.ItemData(cboDivisa.ListIndex) & "'"
         End If
                  
         If cboEPersona.Text <> "TODOS" Then
            If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
            strSQL = strSQL & "{SOCIOS.ESTADOACTUAL} = '" & cboEPersona.ItemData(cboEPersona.ListIndex) & "'"
         End If
                  
         
         If chkResumen.Value = vbUnchecked Then
             .ReportFileName = SIFGlobal.fxPathReportes("Fondos_Movimientos.rpt")
         Else
             .ReportFileName = SIFGlobal.fxPathReportes("Fondos_MovimientosRsm.rpt")
         End If
      
      
      Case opt.Item(1).Value, opt.Item(6).Value 'Contratos
         
         If Mid(cboEstado.Text, 1, 1) <> "T" Then strSQL = "{FND_CONTRATOS.ESTADO} = '" & Mid(cboEstado.Text, 1, 1) & "'"
         
            If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
                strSQL = strSQL & "{FND_CONTRATOS.COD_OPERADORA} = " & cboOperadora.ItemData(cboOperadora.ListIndex)
            
         If chkPlanes.Value = vbUnchecked Then
            If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
            strSQL = strSQL & "{FND_CONTRATOS.COD_PLAN} ='" & Trim(txtCodigo) & "'"
         End If
                     
         If chkContratosSaldos.Value = xtpChecked Then
            If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
            strSQL = strSQL & "{FND_CONTRATOS.APORTES}  > 0"
         End If
                     
         If chkFechas.Value = vbUnchecked Then
            If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
            If cboFechaBase.Text = "Fecha Inicio" Then
                strSQL = strSQL & "{FND_CONTRATOS.FECHA_INICIO} in Date(" & Format(dtpInicio.Value, "yyyy,mm,dd") _
                       & ") to Date(" & Format(dtpCorte.Value, "yyyy,mm,dd") & ")"
            Else
                strSQL = strSQL & "{FND_CONTRATOS.FECHA_CORTE} in Date(" & Format(dtpInicio.Value, "yyyy,mm,dd") _
                       & ") to Date(" & Format(dtpCorte.Value, "yyyy,mm,dd") & ")"
            End If
         End If
                     
         
         'Filtros Especiales
         If Trim(txtCedula.Text) <> "" Then
            If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
            strSQL = strSQL & "{FND_CONTRATOS.CEDULA} = '" & Trim(txtCedula.Text) & "'"
         End If
         
         
         If cboInstitucion.Text <> "TODOS" Then
            If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
            strSQL = strSQL & "{SOCIOS.COD_INSTITUCION} = " & cboInstitucion.ItemData(cboInstitucion.ListIndex)
         End If
         
         
         If chkDeduce.Value = xtpChecked Then
            If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
            strSQL = strSQL & "{FND_CONTRATOS.IND_DEDUCCION} = 0"
         End If
         
         
         If cboDivisa.Text <> "TODOS" Then
            If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
            strSQL = strSQL & "{FND_PLANES.COD_MONEDA} = '" & cboDivisa.ItemData(cboDivisa.ListIndex) & "'"
         End If
                  
         If cboEPersona.Text <> "TODOS" Then
            If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
            strSQL = strSQL & "{SOCIOS.ESTADOACTUAL} = '" & cboEPersona.ItemData(cboEPersona.ListIndex) & "'"
         End If
                           
         If opt.Item(6).Value Then
                   If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
                   strSQL = strSQL & "{FND_CONTRATOS.APORTES} < 0"
         End If
         
        If chkResumen.Value = vbUnchecked Then
            .ReportFileName = SIFGlobal.fxPathReportes("Fondos_Contratos.rpt")
        Else
            .ReportFileName = SIFGlobal.fxPathReportes("Fondos_ContratosRsm.rpt")
        End If
         
         
         
      Case opt.Item(5).Value 'Movimientos x Documento
      
         If Mid(cboEstado.Text, 1, 1) <> "T" Then strSQL = "{FND_CONTRATOS.ESTADO} = '" & Mid(cboEstado.Text, 1, 1) & "'"
         
            If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
            strSQL = strSQL & "{FND_OPERADORAS.COD_OPERADORA} = " & cboOperadora.ItemData(cboOperadora.ListIndex)
            
         If chkPlanes.Value = vbUnchecked Then
            If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
            strSQL = strSQL & "{FND_CONTRATOS.COD_PLAN} ='" & Trim(txtCodigo) & "'"
         End If
                     
         If chkFechas.Value = vbUnchecked Then
            If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
            strSQL = strSQL & "{FND_CONTRATOS_DETALLE.FECHA} in Date(" & Format(dtpInicio.Value, "yyyy,mm,dd") _
                   & ") to Date(" & Format(dtpCorte.Value, "yyyy,mm,dd") & ")"
         End If
                     
         'Filtros Especiales
         If Trim(txtUsuario.Text) <> "" Then
            If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
            strSQL = strSQL & "{FND_CONTRATOS_DETALLE.USUARIO} = '" & Trim(txtUsuario.Text) & "'"
         End If
                     
         If Trim(txtCedula.Text) <> "" Then
            If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
            strSQL = strSQL & "{FND_CONTRATOS.CEDULA} = '" & Trim(txtCedula.Text) & "'"
         End If
                     
         If Trim(txtNDoc.Text) <> "" Then
            If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
            strSQL = strSQL & "{FND_CONTRATOS_DETALLE.NCON} = '" & Trim(txtNDoc.Text) & "'"
         End If
                     
         If Trim(txtNDocRef.Text) <> "" Then
            If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
            strSQL = strSQL & "{FND_CONTRATOS_DETALLE.REF_01} = '" & Trim(txtNDocRef.Text) & "'"
         End If
         
         If cboTipoDoc.Text <> "TODOS" Then
            If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
            strSQL = strSQL & "{FND_CONTRATOS_DETALLE.TCON} = '" & cboTipoDoc.ItemData(cboTipoDoc.ListIndex) & "'"
         End If
         
         If cboConcepto.Text <> "TODOS" Then
            If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
            strSQL = strSQL & "{FND_CONTRATOS_DETALLE.COD_CONCEPTO} = '" & cboConcepto.ItemData(cboConcepto.ListIndex) & "'"
         End If
                     
         
         If cboInstitucion.Text <> "TODOS" Then
            If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
            strSQL = strSQL & "{SOCIOS.COD_INSTITUCION} = " & cboInstitucion.ItemData(cboInstitucion.ListIndex)
         End If
         
         
         If cboDivisa.Text <> "TODOS" Then
            If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
            strSQL = strSQL & "{FND_PLANES.COD_MONEDA} = '" & cboDivisa.ItemData(cboDivisa.ListIndex) & "'"
         End If
                  
         If cboEPersona.Text <> "TODOS" Then
            If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
            strSQL = strSQL & "{SOCIOS.ESTADOACTUAL} = '" & cboEPersona.ItemData(cboEPersona.ListIndex) & "'"
         End If
                           
         
         If chkResumen.Value = vbUnchecked Then
             .ReportFileName = SIFGlobal.fxPathReportes("Fondos_MovimientosXDocumento.rpt")
         Else
             .ReportFileName = SIFGlobal.fxPathReportes("Fondos_MovimientosRsmXDocumento.rpt")
         End If
      
      
      Case opt.Item(1).Value 'Contratos
         
         If Mid(cboEstado.Text, 1, 1) <> "T" Then strSQL = "{FND_CONTRATOS.ESTADO} = '" & Mid(cboEstado.Text, 1, 1) & "'"
         
            If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
                strSQL = strSQL & "{FND_CONTRATOS.COD_OPERADORA} = " & cboOperadora.ItemData(cboOperadora.ListIndex)
            
         If chkPlanes.Value = vbUnchecked Then
            If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
            strSQL = strSQL & "{FND_CONTRATOS.COD_PLAN} ='" & Trim(txtCodigo) & "'"
         End If
                     
         If chkContratosSaldos.Value = xtpChecked Then
            If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
            strSQL = strSQL & "{FND_CONTRATOS.APORTES}  > 0"
         End If
                     
         If chkFechas.Value = vbUnchecked Then
            If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
            If cboFechaBase.Text = "Fecha Inicio" Then
                strSQL = strSQL & "{FND_CONTRATOS.FECHA_INICIO} in Date(" & Format(dtpInicio.Value, "yyyy,mm,dd") _
                       & ") to Date(" & Format(dtpCorte.Value, "yyyy,mm,dd") & ")"
            Else
                strSQL = strSQL & "{FND_CONTRATOS.FECHA_CORTE} in Date(" & Format(dtpInicio.Value, "yyyy,mm,dd") _
                       & ") to Date(" & Format(dtpCorte.Value, "yyyy,mm,dd") & ")"
            End If
         End If
                     
         
         If cboInstitucion.Text <> "TODOS" Then
            If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
            strSQL = strSQL & "{SOCIOS.COD_INSTITUCION} = " & cboInstitucion.ItemData(cboInstitucion.ListIndex)
         End If
         
         If cboDivisa.Text <> "TODOS" Then
            If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
            strSQL = strSQL & "{FND_PLANES.COD_MONEDA} = '" & cboDivisa.ItemData(cboDivisa.ListIndex) & "'"
         End If
                  
         If cboEPersona.Text <> "TODOS" Then
            If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
            strSQL = strSQL & "{SOCIOS.ESTADOACTUAL} = '" & cboEPersona.ItemData(cboEPersona.ListIndex) & "'"
         End If
                           
                           
         If chkResumen.Value = vbUnchecked Then
             .ReportFileName = SIFGlobal.fxPathReportes("Fondos_Contratos.rpt")
         Else
             .ReportFileName = SIFGlobal.fxPathReportes("Fondos_ContratosRsm.rpt")
         End If
         
         
         
      
      Case opt.Item(2).Value 'Liquidaciones
'         If Mid(cboEstado.Text, 1, 1) <> "T" Then strSQL = "{FND_CONTRATOS.ESTADO} = '" & Mid(cboEstado.Text, 1, 1) & "'"
         
        If Len(strSQL) > 0 Then
            strSQL = strSQL & " AND "
            strSQL = strSQL & "{FND_LIQUIDACION.ESTADO} = 'P'"
        End If
        
        If Len(strSQL) > 0 Then
            strSQL = strSQL & " AND "
            strSQL = strSQL & "{FND_LIQUIDACION.COD_OPERADORA} = " & cboOperadora.ItemData(cboOperadora.ListIndex)
        End If
        
         If chkPlanes.Value = vbUnchecked Then
            If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
            strSQL = strSQL & "{FND_LIQUIDACION.COD_PLAN} ='" & Trim(txtCodigo) & "'"
         End If
                     
         If chkFechas.Value = vbUnchecked Then
            If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
            strSQL = strSQL & "{FND_LIQUIDACION.FECHA} in Date(" & Format(dtpInicio.Value, "yyyy,mm,dd") _
                   & ") to Date(" & Format(dtpCorte.Value, "yyyy,mm,dd") & ")"
         End If
                     
         If cboInstitucion.Text <> "TODOS" Then
            If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
            strSQL = strSQL & "{SOCIOS.COD_INSTITUCION} = " & cboInstitucion.ItemData(cboInstitucion.ListIndex)
         End If
                     
                     
         If chkResumen.Value = vbUnchecked Then
             .ReportFileName = SIFGlobal.fxPathReportes("Fondos_Liquidaciones.rpt")
         Else
             .ReportFileName = SIFGlobal.fxPathReportes("Fondos_LiquidacionesRsm.rpt")
         End If
         
         
         
      Case opt.Item(3).Value 'Contratos de CDP
         If Mid(cboEstado.Text, 1, 1) <> "T" Then strSQL = "{FND_CONTRATOS.ESTADO} = '" & Mid(cboEstado.Text, 1, 1) & "'"
         
            If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
                strSQL = strSQL & "{FND_CONTRATOS.COD_OPERADORA} = " & cboOperadora.ItemData(cboOperadora.ListIndex)
            
         If chkPlanes.Value = vbUnchecked Then
            If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
            strSQL = strSQL & "{FND_CONTRATOS.COD_PLAN} ='" & Trim(txtCodigo) & "'"
         End If
                     
         If chkContratosSaldos.Value = xtpChecked Then
            If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
            strSQL = strSQL & "{FND_CONTRATOS.APORTES}  > 0"
         End If
                     
         If chkFechas.Value = vbUnchecked Then
            If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
            If cboFechaBase.Text = "Fecha Inicio" Then
                strSQL = strSQL & "{FND_CONTRATOS.FECHA_INICIO} in Date(" & Format(dtpInicio.Value, "yyyy,mm,dd") _
                       & ") to Date(" & Format(dtpCorte.Value, "yyyy,mm,dd") & ")"
            Else
                strSQL = strSQL & "{FND_CONTRATOS.FECHA_CORTE} in Date(" & Format(dtpInicio.Value, "yyyy,mm,dd") _
                       & ") to Date(" & Format(dtpCorte.Value, "yyyy,mm,dd") & ")"
            End If
         End If
                     
         If cboInstitucion.Text <> "TODOS" Then
            If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
            strSQL = strSQL & "{SOCIOS.COD_INSTITUCION} = " & cboInstitucion.ItemData(cboInstitucion.ListIndex)
         End If
                     
         If cboDivisa.Text <> "TODOS" Then
            If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
            strSQL = strSQL & "{FND_PLANES.COD_MONEDA} = '" & cboDivisa.ItemData(cboDivisa.ListIndex) & "'"
         End If
                  
         If cboEPersona.Text <> "TODOS" Then
            If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
            strSQL = strSQL & "{SOCIOS.ESTADOACTUAL} = '" & cboEPersona.ItemData(cboEPersona.ListIndex) & "'"
         End If
                                       
                     
         If chkResumen.Value = vbUnchecked Then
             .ReportFileName = SIFGlobal.fxPathReportes("Fondos_ContratosCDPs.rpt")
         Else
             .ReportFileName = SIFGlobal.fxPathReportes("Fondos_ContratosRsm.rpt")
         End If
         
         
         
      Case opt.Item(4).Value 'Contratos con Retiros parciales en Cero
         
         If Mid(cboEstado.Text, 1, 1) <> "T" Then strSQL = "{FND_CONTRATOS.ESTADO} = '" & Mid(cboEstado.Text, 1, 1) & "'"
         
            If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
                strSQL = strSQL & "{FND_CONTRATOS.LIQ_TIPO} = 'R' AND {FND_CONTRATOS.APORTES} = 0"
         
         
            If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
                strSQL = strSQL & "{FND_CONTRATOS.COD_OPERADORA} = " & cboOperadora.ItemData(cboOperadora.ListIndex)
            
         If chkPlanes.Value = vbUnchecked Then
            If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
            strSQL = strSQL & "{FND_CONTRATOS.COD_PLAN} ='" & Trim(txtCodigo) & "'"
         End If
                     
         If chkFechas.Value = vbUnchecked Then
            If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
            Select Case cboFechaBase.Text
              Case "Ult.Retiro"
                strSQL = strSQL & "{FND_CONTRATOS.LIQ_FECHA} in Date(" & Format(dtpInicio.Value, "yyyy,mm,dd") _
                       & ") to Date(" & Format(dtpCorte.Value, "yyyy,mm,dd") & ")"
            
              Case "Fecha Inicio"
                strSQL = strSQL & "{FND_CONTRATOS.FECHA_INICIO} in Date(" & Format(dtpInicio.Value, "yyyy,mm,dd") _
                       & ") to Date(" & Format(dtpCorte.Value, "yyyy,mm,dd") & ")"
              Case "Fecha Corte"
                strSQL = strSQL & "{FND_CONTRATOS.FECHA_CORTE} in Date(" & Format(dtpInicio.Value, "yyyy,mm,dd") _
                       & ") to Date(" & Format(dtpCorte.Value, "yyyy,mm,dd") & ")"
            End Select
         End If
                     
                     
         If cboInstitucion.Text <> "TODOS" Then
            If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
            strSQL = strSQL & "{SOCIOS.COD_INSTITUCION} = " & cboInstitucion.ItemData(cboInstitucion.ListIndex)
         End If
         
         If cboDivisa.Text <> "TODOS" Then
            If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
            strSQL = strSQL & "{FND_PLANES.COD_MONEDA} = '" & cboDivisa.ItemData(cboDivisa.ListIndex) & "'"
         End If
                  
         If cboEPersona.Text <> "TODOS" Then
            If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
            strSQL = strSQL & "{SOCIOS.ESTADOACTUAL} = '" & cboEPersona.ItemData(cboEPersona.ListIndex) & "'"
         End If
                           
                           
         If chkResumen.Value = vbUnchecked Then
             .ReportFileName = SIFGlobal.fxPathReportes("Fondos_ContratosRetirosParcialesCero.rpt")
         Else
             .ReportFileName = SIFGlobal.fxPathReportes("Fondos_ContratosRetirosParcialesCero.rpt")
         End If
         
         
      Case opt.Item(7).Value, opt.Item(8).Value, opt.Item(9).Value
      'Cashback
        MsgBox "Informe no localizado!", vbInformation
        
      Case opt.Item(10).Value 'IVA Trasfronterizo
        MsgBox "Informe no localizado!", vbInformation
         
      Case opt.Item(11).Value 'Proyeccion de Cupones
        MsgBox "Informe no localizado!", vbInformation
         
         
    End Select
     
     
     
     If chkFechas.Value = vbUnchecked Then
        .Formulas(0) = "Subtitulo='" & cboFechaBase.Text & " : Incio " & Format(dtpInicio.Value, "dd/mm/yyyy") _
                      & " Corte " & Format(dtpCorte.Value, "dd/mm/yyyy") & " / Contratos : " & cboEstado.Text & "'"
     Else
        .Formulas(0) = "Subtitulo='Todas las Fechas / Contratos : " & cboEstado.Text & "'"
     End If
    
    
    .SelectionFormula = strSQL
    .Formulas(1) = "Fecha='" & Format(fxFechaServidor, "dd/mm/yyyy") & "'"
    .Formulas(2) = "Usuario='" & Trim(glogon.Usuario) & "'"
    .Formulas(3) = "Empresa='" & Trim(GLOBALES.gstrNombreEmpresa) & "'"
    
    
    .PrintReport



End With

Me.MousePointer = vbDefault

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub dtpInicio_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then dtpInicio.SetFocus
End Sub


Private Sub FlatScrollBar_Change()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

If vScroll Then
    strSQL = "select Top 1 cod_plan from fnd_planes" _
           & " where cod_operadora = " & cboOperadora.ItemData(cboOperadora.ListIndex)
    
    If FlatScrollBar.Value = 1 Then
       strSQL = strSQL & " and cod_plan > '" & txtCodigo & "' order by cod_plan asc"
    Else
       strSQL = strSQL & " and cod_plan < '" & txtCodigo & "' order by cod_plan desc"
    End If
    
    Call OpenRecordSet(rs, strSQL)
    If Not rs.EOF And Not rs.BOF Then
      txtCodigo = rs!Cod_Plan
      txtCodigo_LostFocus
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

Private Sub sbCubo()
Dim strSQL As String, rs As New ADODB.Recordset
Dim vFechaInicio As Date, vFechaCorte As Date
Dim vMensaje As String

On Error GoTo vError

Me.MousePointer = vbHourglass

lblStatus.Caption = "Procesando Información Espere!....Este proceso puede durar varios minutos."
lblStatus.Refresh

vMensaje = "Fondos_Movimientos"

If chkFechas.Value = vbChecked Then
  vFechaInicio = "1900/01/01"
  vFechaCorte = fxFechaServidor
Else
  vFechaInicio = dtpInicio.Value
  vFechaCorte = dtpCorte.Value
End If

strSQL = "exec spFndMovAnalisisCubo '" & Format(vFechaInicio, "yyyy/mm/dd") & "','" & Format(dtpCorte, "yyyy/mm/dd") & "'"
Call ConectionExecute(strSQL)

lblStatus.Caption = "Proceso Concluido con éxito, la información puede ser utilizada desde la base de datos de análisis, cubo: " & vMensaje

Me.MousePointer = vbDefault

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub Form_Activate()
vModulo = 18 'Fondo de Inversion

End Sub

Private Sub Form_Load()

vModulo = 18 'Fondo de Inversion


Set imgBanner.Picture = frmContenedor.imgBanner_Reportes.Picture


vScroll = False
 FlatScrollBar.Value = 0
vScroll = True


Call Formularios(Me)
Call RefrescaTags(Me)


 
End Sub



Private Sub opt_Click(Index As Integer)
Dim i As Integer

On Error Resume Next

'chkPlanes.BackColor = txtDescripcion.BackColor

For i = 0 To opt.Count - 1
 opt(i).FontBold = False
 opt(i).ForeColor = vbBlack
Next i

 opt(Index).FontBold = True
 opt(Index).ForeColor = vbBlue

cboFechaBase.Clear
chkFechas.Value = vbChecked

Select Case Index
   Case 0, 5, 7, 8, 9, 10  'Movimientos + Cashback
       cboFechaBase.AddItem "Transacción"
       cboFechaBase.Text = "Transacción"
       cboEstado.Text = "Todos"
   Case 1, 6  'Contratos
       cboFechaBase.AddItem "Fecha Inicio"
       cboFechaBase.AddItem "Fecha Corte"
       cboFechaBase.Text = "Fecha Inicio"
       cboEstado.Text = "Activos"
   Case 2 'Liquidaciones
       cboEstado.Text = "Todos"
       cboFechaBase.AddItem "Transacción"
       cboFechaBase.Text = "Transacción"
   Case 3, 11 'Vencimientos
       cboFechaBase.AddItem "Fecha Inicio"
       cboFechaBase.AddItem "Fecha Corte"
       cboFechaBase.Text = "Fecha Inicio"
       cboEstado.Text = "Activos"
   Case 4 'Retiros Parciales
       cboEstado.Text = "Todos"
       cboFechaBase.AddItem "Ult.Retiro"
       cboFechaBase.AddItem "Fecha Inicio"
       cboFechaBase.AddItem "Fecha Corte"
       cboFechaBase.Text = "Ult.Retiro"
End Select

chkFechas_Click

End Sub

Private Sub TimerX_Timer()
TimerX.Interval = 0
TimerX.Enabled = False

'chkPlanes.BackColor = txtDescripcion.BackColor

Dim strSQL As String

cboEstado.Clear
cboEstado.AddItem "Todos"
cboEstado.AddItem "Activos"
cboEstado.AddItem "Liquidados"
cboEstado.AddItem "Bloqueados"
cboEstado.AddItem "Inactivos"
cboEstado.Text = "Activos"
 
 
strSQL = "select rtrim(cod_estado) as 'IdX', rtrim(descripcion) as 'Itmx'" _
       & " from afi_estados_persona order by descripcion"
Call sbCbo_Llena_New(cboEPersona, strSQL, True, False)
 
strSQL = "select COD_DIVISA AS 'IdX', DESCRIPCION as 'ItmX'" _
       & " From vSys_Divisas"
Call sbCbo_Llena_New(cboDivisa, strSQL, True, True)

dtpInicio.Value = fxFechaServidor
dtpCorte.Value = dtpInicio.Value
 
 
strSQL = "select rtrim(Tipo_Documento) as 'IdX', rtrim(Descripcion) as 'ItmX'" _
       & " from sif_documentos Where Tipo_Documento in('FLIQ','FND','FNC','FRND','PLA','RE','NC','ND','SINPE','TD','PGSP')"
Call sbCbo_Llena_New(cboTipoDoc, strSQL, True, True)
 
strSQL = "select rtrim(cod_Concepto) as 'IdX', rtrim(Descripcion) as ItmX" _
       & " from sif_conceptos Where cod_Concepto like 'FND%'"
Call sbCbo_Llena_New(cboConcepto, strSQL, True, True)

strSQL = "select descripcion as 'itmX',cod_operadora as 'IdX' from FND_Operadoras"
Call sbCbo_Llena_New(cboOperadora, strSQL, False, True)

strSQL = "select cod_institucion as 'IdX', rtrim(descripcion) as 'ItmX' from instituciones where Activa = 1 order by descripcion"
Call sbCbo_Llena_New(cboInstitucion, strSQL, True, True)


 
Call opt_Click(0)



End Sub


Private Sub txtCodigo_KeyDown(KeyCode As Integer, Shift As Integer)

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


Private Sub txtCodigo_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then txtDescripcion.SetFocus
End Sub


Private Sub txtCodigo_LostFocus()
Dim strSQL As String, rs As New ADODB.Recordset

strSQL = "Select Descripcion from Fnd_Planes where Cod_Operadora="
strSQL = strSQL & cboOperadora.ItemData(cboOperadora.ListIndex) & " And "
strSQL = strSQL & "Cod_Plan='" & Trim(txtCodigo) & "'"
With rs
 .Open strSQL, glogon.Conection, adOpenStatic
    If .EOF = False Then
       txtDescripcion = Trim(!Descripcion)
    Else
       txtCodigo = ""
       txtDescripcion = ""
    End If
 .Close
End With

End Sub


Private Sub txtDescripcion_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 115 Then
   gBusquedas.Columna = "descripcion"
   gBusquedas.Orden = "descripcion"
   gBusquedas.Filtro = "And Cod_operadora=" & cboOperadora.ItemData(cboOperadora.ListIndex)
   gBusquedas.Consulta = "select cod_plan,descripcion from fnd_planes"
   frmBusquedas.Show vbModal
   cboEstado.SetFocus
   
   If Trim(gBusquedas.Resultado) <> "" Then
      txtCodigo = Trim(gBusquedas.Resultado)
      txtDescripcion = Trim(gBusquedas.Resultado2)
   End If
   gBusquedas.Resultado = ""
   gBusquedas.Resultado2 = ""
End If
End Sub




