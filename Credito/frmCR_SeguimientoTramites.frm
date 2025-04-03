VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.5#0"; "comctl32.Ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.Controls.v24.0.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.ShortcutBar.v24.0.0.ocx"
Object = "{B8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.TaskPanel.v24.0.0.ocx"
Begin VB.Form frmCR_SeguimientoTramites 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Trámite de Crédito"
   ClientHeight    =   8760
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11730
   HelpContextID   =   3027
   Icon            =   "frmCR_SeguimientoTramites.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8760
   ScaleWidth      =   11730
   Begin XtremeTaskPanel.TaskPanel tpMain 
      Height          =   7005
      Left            =   0
      TabIndex        =   106
      Top             =   1335
      Width           =   2790
      _Version        =   1572864
      _ExtentX        =   4921
      _ExtentY        =   12356
      _StockProps     =   64
      VisualTheme     =   17
      ItemLayout      =   2
      HotTrackStyle   =   1
   End
   Begin XtremeSuiteControls.GroupBox gbReportes 
      Height          =   5055
      Left            =   11160
      TabIndex        =   116
      Top             =   1440
      Visible         =   0   'False
      Width           =   5895
      _Version        =   1572864
      _ExtentX        =   10398
      _ExtentY        =   8916
      _StockProps     =   79
      BackColor       =   16777215
      Appearance      =   16
      Begin XtremeSuiteControls.PushButton btnInforme 
         Height          =   495
         Index           =   0
         Left            =   3840
         TabIndex        =   117
         Top             =   4320
         Width           =   1455
         _Version        =   1572864
         _ExtentX        =   2566
         _ExtentY        =   873
         _StockProps     =   79
         Caption         =   "Informe"
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
         Appearance      =   17
         Picture         =   "frmCR_SeguimientoTramites.frx":6852
      End
      Begin XtremeSuiteControls.RadioButton rbInformes 
         Height          =   255
         Index           =   0
         Left            =   600
         TabIndex        =   118
         Top             =   600
         Width           =   4335
         _Version        =   1572864
         _ExtentX        =   7646
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Boleta de Formalización"
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
         Value           =   -1  'True
      End
      Begin XtremeSuiteControls.RadioButton rbInformes 
         Height          =   255
         Index           =   1
         Left            =   600
         TabIndex        =   119
         Top             =   960
         Width           =   4335
         _Version        =   1572864
         _ExtentX        =   7646
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Emisión de Contrato y Garantía"
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
      End
      Begin XtremeSuiteControls.RadioButton rbInformes 
         Height          =   255
         Index           =   3
         Left            =   600
         TabIndex        =   120
         Top             =   1920
         Width           =   4335
         _Version        =   1572864
         _ExtentX        =   7646
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Solicitud de Crédito"
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
      End
      Begin XtremeSuiteControls.RadioButton rbInformes 
         Height          =   255
         Index           =   4
         Left            =   600
         TabIndex        =   121
         Top             =   2280
         Width           =   4335
         _Version        =   1572864
         _ExtentX        =   7646
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Recibos de Refundiciones"
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
      End
      Begin XtremeSuiteControls.RadioButton rbInformes 
         Height          =   255
         Index           =   5
         Left            =   600
         TabIndex        =   122
         Top             =   2640
         Width           =   4335
         _Version        =   1572864
         _ExtentX        =   7646
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Carátula para Expediente"
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
      End
      Begin XtremeSuiteControls.RadioButton rbInformes 
         Height          =   255
         Index           =   6
         Left            =   600
         TabIndex        =   123
         Top             =   3240
         Width           =   4335
         _Version        =   1572864
         _ExtentX        =   7646
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Boleta de Requisitos"
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
      End
      Begin XtremeSuiteControls.RadioButton rbInformes 
         Height          =   255
         Index           =   7
         Left            =   600
         TabIndex        =   124
         Top             =   3600
         Width           =   4335
         _Version        =   1572864
         _ExtentX        =   7646
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Boleta para Emisión de Cheques"
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
      End
      Begin XtremeSuiteControls.RadioButton rbInformes 
         Height          =   255
         Index           =   8
         Left            =   600
         TabIndex        =   125
         Top             =   4080
         Width           =   4335
         _Version        =   1572864
         _ExtentX        =   7646
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Actas de Resolución"
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
      End
      Begin XtremeSuiteControls.RadioButton rbInformes 
         Height          =   255
         Index           =   2
         Left            =   600
         TabIndex        =   126
         Top             =   1320
         Width           =   4335
         _Version        =   1572864
         _ExtentX        =   7646
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Autorización de Deducción"
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
      End
      Begin XtremeSuiteControls.PushButton btnInforme 
         Height          =   495
         Index           =   1
         Left            =   5280
         TabIndex        =   127
         Top             =   4320
         Width           =   495
         _Version        =   1572864
         _ExtentX        =   873
         _ExtentY        =   873
         _StockProps     =   79
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
         Appearance      =   17
         Picture         =   "frmCR_SeguimientoTramites.frx":6F59
      End
      Begin XtremeSuiteControls.RadioButton rbInformes 
         Height          =   255
         Index           =   9
         Left            =   600
         TabIndex        =   128
         Top             =   4440
         Visible         =   0   'False
         Width           =   3015
         _Version        =   1572864
         _ExtentX        =   5318
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Estudio de Crédito"
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
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
         Height          =   375
         Index           =   1
         Left            =   0
         TabIndex        =   129
         Top             =   0
         Width           =   5895
         _Version        =   1572864
         _ExtentX        =   10398
         _ExtentY        =   661
         _StockProps     =   14
         Caption         =   "Informes:"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   10.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SubItemCaption  =   -1  'True
      End
   End
   Begin XtremeSuiteControls.TabControl tcMain 
      Height          =   7335
      Left            =   2880
      TabIndex        =   4
      Top             =   1440
      Width           =   8895
      _Version        =   1572864
      _ExtentX        =   15690
      _ExtentY        =   12938
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
      PaintManager.Position=   2
      ItemCount       =   5
      Item(0).Caption =   "Recepcion"
      Item(0).ControlCount=   6
      Item(0).Control(0)=   "fraOperacion"
      Item(0).Control(1)=   "txtObservaciones"
      Item(0).Control(2)=   "Label1(19)"
      Item(0).Control(3)=   "chkExpedienteDigital"
      Item(0).Control(4)=   "txtFormularioId"
      Item(0).Control(5)=   "chkPagareManual"
      Item(1).Caption =   "Formalizacion"
      Item(1).ControlCount=   19
      Item(1).Control(0)=   "lsw"
      Item(1).Control(1)=   "chkDeducPlanilla"
      Item(1).Control(2)=   "GroupBox1"
      Item(1).Control(3)=   "lswOpciones"
      Item(1).Control(4)=   "fraEstadoFormalizacion"
      Item(1).Control(5)=   "cboMes"
      Item(1).Control(6)=   "txtPagare"
      Item(1).Control(7)=   "txtDocumento"
      Item(1).Control(8)=   "txtAno"
      Item(1).Control(9)=   "chkPrimera"
      Item(1).Control(10)=   "chkEnviarATesoreria"
      Item(1).Control(11)=   "Label4"
      Item(1).Control(12)=   "Label2"
      Item(1).Control(13)=   "Label16(2)"
      Item(1).Control(14)=   "TituloOpcionesSub"
      Item(1).Control(15)=   "Label5"
      Item(1).Control(16)=   "cboDeductora"
      Item(1).Control(17)=   "cboFrecuencia"
      Item(1).Control(18)=   "chkTrasladoSalario"
      Item(2).Caption =   "Historial"
      Item(2).ControlCount=   25
      Item(2).Control(0)=   "txtAutorizaNota"
      Item(2).Control(1)=   "Label3(0)"
      Item(2).Control(2)=   "Label3(1)"
      Item(2).Control(3)=   "Label3(2)"
      Item(2).Control(4)=   "Label3(3)"
      Item(2).Control(5)=   "Label3(4)"
      Item(2).Control(6)=   "Label3(5)"
      Item(2).Control(7)=   "imgPoolHistorico(0)"
      Item(2).Control(8)=   "imgPoolHistorico(1)"
      Item(2).Control(9)=   "imgPoolHistorico(2)"
      Item(2).Control(10)=   "imgPoolHistorico(3)"
      Item(2).Control(11)=   "imgPoolHistorico(4)"
      Item(2).Control(12)=   "imgPoolHistorico(5)"
      Item(2).Control(13)=   "btnTags"
      Item(2).Control(14)=   "txtRecibe"
      Item(2).Control(15)=   "txtFechaRec"
      Item(2).Control(16)=   "txtResoluciona"
      Item(2).Control(17)=   "txtFechaRes"
      Item(2).Control(18)=   "txtFechaFor"
      Item(2).Control(19)=   "txtTesoreria"
      Item(2).Control(20)=   "txtFechaTes"
      Item(2).Control(21)=   "txtAutorizada"
      Item(2).Control(22)=   "txtFechaAuto"
      Item(2).Control(23)=   "lswHistorial"
      Item(2).Control(24)=   "txtFormaliza"
      Item(3).Caption =   "Consulta"
      Item(3).ControlCount=   4
      Item(3).Control(0)=   "txtConNombre"
      Item(3).Control(1)=   "txtConCedula"
      Item(3).Control(2)=   "lswBusca"
      Item(3).Control(3)=   "Label1(26)"
      Item(4).Caption =   "Monto No Gravable"
      Item(4).ControlCount=   4
      Item(4).Control(0)=   "txtMntNoGravable"
      Item(4).Control(1)=   "Label1(5)"
      Item(4).Control(2)=   "Label1(6)"
      Item(4).Control(3)=   "btnMtnNoGravable"
      Begin XtremeSuiteControls.ListView lswHistorial 
         Height          =   2412
         Left            =   -70000
         TabIndex        =   98
         Top             =   120
         Visible         =   0   'False
         Width           =   8532
         _Version        =   1572864
         _ExtentX        =   15049
         _ExtentY        =   4254
         _StockProps     =   77
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
         View            =   3
         FullRowSelect   =   -1  'True
         FlatScrollBar   =   -1  'True
         Appearance      =   16
         ShowBorder      =   0   'False
      End
      Begin XtremeSuiteControls.ListView lsw 
         Height          =   3372
         Left            =   -67000
         TabIndex        =   47
         Top             =   120
         Visible         =   0   'False
         Width           =   5652
         _Version        =   1572864
         _ExtentX        =   9970
         _ExtentY        =   5948
         _StockProps     =   77
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
         View            =   3
         FullRowSelect   =   -1  'True
         FlatScrollBar   =   -1  'True
         Appearance      =   16
         ShowBorder      =   0   'False
      End
      Begin XtremeSuiteControls.ListView lswBusca 
         Height          =   6132
         Left            =   -70000
         TabIndex        =   99
         Top             =   600
         Visible         =   0   'False
         Width           =   8532
         _Version        =   1572864
         _ExtentX        =   15049
         _ExtentY        =   10816
         _StockProps     =   77
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
         View            =   3
         FullRowSelect   =   -1  'True
         FlatScrollBar   =   -1  'True
         Appearance      =   16
         ShowBorder      =   0   'False
      End
      Begin XtremeSuiteControls.CheckBox chkExpedienteDigital 
         Height          =   255
         Left            =   1080
         TabIndex        =   137
         Top             =   6480
         Width           =   1815
         _Version        =   1572864
         _ExtentX        =   3201
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Expediente Digital"
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
      Begin XtremeSuiteControls.PushButton btnMtnNoGravable 
         Height          =   492
         Left            =   -66640
         TabIndex        =   105
         Top             =   1680
         Visible         =   0   'False
         Width           =   1812
         _Version        =   1572864
         _ExtentX        =   3196
         _ExtentY        =   868
         _StockProps     =   79
         Caption         =   "Actualizar"
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
      Begin VB.Frame fraOperacion 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   5175
         Left            =   120
         TabIndex        =   5
         Top             =   -120
         Width           =   8532
         Begin XtremeSuiteControls.ComboBox cboCalculoAdd 
            Height          =   312
            Left            =   6360
            TabIndex        =   7
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
         End
         Begin XtremeSuiteControls.ComboBox cboGarantia 
            Height          =   312
            Left            =   1200
            TabIndex        =   8
            Top             =   840
            Width           =   1692
            _Version        =   1572864
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
         Begin XtremeSuiteControls.ComboBox cboDestino 
            Height          =   315
            Left            =   4080
            TabIndex        =   9
            Top             =   840
            Width           =   4455
            _Version        =   1572864
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
         Begin XtremeSuiteControls.ComboBox cboActividad 
            Height          =   315
            Left            =   960
            TabIndex        =   10
            Top             =   2040
            Width           =   4215
            _Version        =   1572864
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
         Begin XtremeSuiteControls.ComboBox cboCanal 
            Height          =   315
            Left            =   960
            TabIndex        =   11
            Top             =   2400
            Width           =   4215
            _Version        =   1572864
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
         Begin XtremeSuiteControls.ComboBox cboFondo 
            Height          =   315
            Left            =   960
            TabIndex        =   12
            Top             =   2880
            Width           =   4215
            _Version        =   1572864
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
         Begin XtremeSuiteControls.ComboBox cboFondoContrato 
            Height          =   315
            Left            =   960
            TabIndex        =   13
            Top             =   3240
            Width           =   4215
            _Version        =   1572864
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
         Begin XtremeSuiteControls.ComboBox cboComite 
            Height          =   315
            Left            =   960
            TabIndex        =   14
            Top             =   3720
            Width           =   4215
            _Version        =   1572864
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
         Begin XtremeSuiteControls.ComboBox cboBanco 
            Height          =   315
            Left            =   960
            TabIndex        =   15
            Top             =   4080
            Width           =   4215
            _Version        =   1572864
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
         Begin XtremeSuiteControls.ComboBox cboCuenta 
            Height          =   315
            Left            =   960
            TabIndex        =   16
            Top             =   4440
            Width           =   4215
            _Version        =   1572864
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
         Begin XtremeSuiteControls.ComboBox cboEstado 
            Height          =   315
            Left            =   6360
            TabIndex        =   17
            Top             =   3720
            Width           =   1815
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
         Begin XtremeSuiteControls.ComboBox cboTipoDocumento 
            Height          =   315
            Left            =   6360
            TabIndex        =   18
            Top             =   4080
            Width           =   1815
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
         Begin XtremeSuiteControls.FlatEdit txtNombre 
            Height          =   315
            Left            =   4080
            TabIndex        =   20
            Top             =   120
            Width           =   4455
            _Version        =   1572864
            _ExtentX        =   7853
            _ExtentY        =   550
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
         Begin XtremeSuiteControls.FlatEdit txtDescripcion 
            Height          =   315
            Left            =   4080
            TabIndex        =   21
            Top             =   480
            Width           =   4455
            _Version        =   1572864
            _ExtentX        =   7853
            _ExtentY        =   550
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
         Begin XtremeSuiteControls.FlatEdit txtPromotorNombre 
            Height          =   312
            Left            =   1440
            TabIndex        =   22
            Top             =   1320
            Width           =   3732
            _Version        =   1572864
            _ExtentX        =   6583
            _ExtentY        =   550
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
         Begin XtremeSuiteControls.DateTimePicker dtpFechaSolicitud 
            Height          =   315
            Left            =   6360
            TabIndex        =   76
            Top             =   4440
            Width           =   1815
            _Version        =   1572864
            _ExtentX        =   3196
            _ExtentY        =   550
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
         Begin XtremeSuiteControls.FlatEdit txtMonto 
            Height          =   312
            Left            =   6360
            TabIndex        =   77
            Top             =   1680
            Width           =   1812
            _Version        =   1572864
            _ExtentX        =   3196
            _ExtentY        =   550
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
         Begin XtremeSuiteControls.FlatEdit txtPlazo 
            Height          =   312
            Left            =   7440
            TabIndex        =   78
            Top             =   2040
            Width           =   732
            _Version        =   1572864
            _ExtentX        =   1291
            _ExtentY        =   550
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
         Begin XtremeSuiteControls.FlatEdit txtTasa 
            Height          =   312
            Left            =   7440
            TabIndex        =   79
            Top             =   2400
            Width           =   732
            _Version        =   1572864
            _ExtentX        =   1291
            _ExtentY        =   550
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
         Begin XtremeSuiteControls.FlatEdit txtCuota 
            Height          =   312
            Left            =   6360
            TabIndex        =   80
            Top             =   2880
            Width           =   1812
            _Version        =   1572864
            _ExtentX        =   3196
            _ExtentY        =   550
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
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtCedula 
            Height          =   312
            Left            =   1200
            TabIndex        =   6
            Top             =   120
            Width           =   1692
            _Version        =   1572864
            _ExtentX        =   2984
            _ExtentY        =   550
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
            Alignment       =   2
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtCodigo 
            Height          =   312
            Left            =   1200
            TabIndex        =   19
            Top             =   480
            Width           =   1692
            _Version        =   1572864
            _ExtentX        =   2984
            _ExtentY        =   550
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
            Alignment       =   2
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtPromotorId 
            Height          =   312
            Left            =   960
            TabIndex        =   23
            Top             =   1320
            Width           =   504
            _Version        =   1572864
            _ExtentX        =   889
            _ExtentY        =   550
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
            Alignment       =   2
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtDivisa 
            Height          =   330
            Left            =   5640
            TabIndex        =   107
            Top             =   1320
            Width           =   732
            _Version        =   1572864
            _ExtentX        =   1291
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
            Transparent     =   -1  'True
         End
         Begin XtremeSuiteControls.FlatEdit txtProveedorNombre 
            Height          =   315
            Left            =   1440
            TabIndex        =   130
            Top             =   4800
            Width           =   3735
            _Version        =   1572864
            _ExtentX        =   6583
            _ExtentY        =   550
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
         Begin XtremeSuiteControls.FlatEdit txtProveedorId 
            Height          =   315
            Left            =   960
            TabIndex        =   131
            Top             =   4800
            Width           =   510
            _Version        =   1572864
            _ExtentX        =   889
            _ExtentY        =   550
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
            Alignment       =   2
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.DateTimePicker dtpVence 
            Height          =   315
            Left            =   6360
            TabIndex        =   134
            Top             =   4800
            Width           =   1815
            _Version        =   1572864
            _ExtentX        =   3196
            _ExtentY        =   550
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
         Begin XtremeSuiteControls.ComboBox cboOficina 
            Height          =   315
            Left            =   960
            TabIndex        =   140
            Top             =   1680
            Width           =   4215
            _Version        =   1572864
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
         Begin VB.Label lblFondoDisplay 
            Caption         =   "Ofi.Pres."
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
            Left            =   0
            TabIndex        =   141
            ToolTipText     =   "Oficina/Agencia Presenta"
            Top             =   1680
            Width           =   735
         End
         Begin VB.Label lblVence 
            Caption         =   "Vence"
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
            Left            =   5400
            TabIndex        =   135
            Top             =   4800
            Width           =   735
         End
         Begin VB.Label lblProveedor 
            Caption         =   "Proveedor"
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
            Left            =   0
            TabIndex        =   132
            Top             =   4800
            Width           =   975
         End
         Begin VB.Label lblFondoDisplay 
            Caption         =   "Ejecutivo"
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
            Left            =   0
            TabIndex        =   44
            Top             =   1320
            Width           =   735
         End
         Begin VB.Label lblFondoDisplay 
            Caption         =   "Actividad"
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
            Left            =   0
            TabIndex        =   43
            Top             =   2040
            Width           =   735
         End
         Begin VB.Image imgBullet 
            Height          =   240
            Left            =   5280
            Picture         =   "frmCR_SeguimientoTramites.frx":766F
            ToolTipText     =   "Definir Cuota Bullet"
            Top             =   2880
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.Label lblFondoDisplay 
            Caption         =   "Contrato"
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
            Left            =   0
            TabIndex        =   42
            Top             =   3240
            Width           =   735
         End
         Begin VB.Image imgMonto 
            Height          =   240
            Left            =   8175
            Picture         =   "frmCR_SeguimientoTramites.frx":7D52
            ToolTipText     =   "Calcular Monto para Giro en Cero"
            Top             =   1320
            Width           =   240
         End
         Begin VB.Label lblFondoDisplay 
            Caption         =   "Respaldo"
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
            Left            =   0
            TabIndex        =   41
            Top             =   2880
            Width           =   855
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
            Height          =   252
            Index           =   16
            Left            =   5640
            TabIndex        =   40
            Top             =   1680
            Width           =   612
         End
         Begin VB.Label lblPlazo 
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
            Height          =   252
            Left            =   5640
            TabIndex        =   39
            Top             =   2040
            Width           =   492
         End
         Begin VB.Label lblTasa 
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
            Height          =   252
            Left            =   5640
            TabIndex        =   38
            Top             =   2400
            Width           =   492
         End
         Begin VB.Label Label1 
            Caption         =   "Cuota"
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
            Left            =   5640
            TabIndex        =   37
            Top             =   2880
            Width           =   492
         End
         Begin VB.Label Label1 
            Caption         =   "Banco"
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
            Index           =   15
            Left            =   0
            TabIndex        =   36
            Top             =   4080
            Width           =   495
         End
         Begin VB.Label lblCuentaTitulo 
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
            Height          =   255
            Left            =   0
            TabIndex        =   35
            Top             =   4440
            Width           =   975
         End
         Begin VB.Label Label1 
            Caption         =   "Emitir"
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
            Left            =   5400
            TabIndex        =   34
            Top             =   4080
            Width           =   855
         End
         Begin VB.Label Label1 
            Caption         =   "Evaluado "
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
            Left            =   0
            TabIndex        =   33
            Top             =   3720
            Width           =   855
         End
         Begin VB.Label Label1 
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
            Height          =   255
            Index           =   11
            Left            =   5400
            TabIndex        =   32
            Top             =   4440
            Width           =   735
         End
         Begin VB.Label Label1 
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
            Index           =   10
            Left            =   5400
            TabIndex        =   31
            Top             =   3720
            Width           =   615
         End
         Begin VB.Label Label1 
            Caption         =   "Garantía"
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
            Left            =   0
            TabIndex        =   30
            Top             =   840
            Width           =   1092
         End
         Begin VB.Label Label1 
            Caption         =   "Descripción"
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
            Left            =   3000
            TabIndex        =   29
            Top             =   480
            Width           =   972
         End
         Begin VB.Label Label1 
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
            Index           =   3
            Left            =   0
            TabIndex        =   28
            Top             =   480
            Width           =   1332
         End
         Begin VB.Label Label1 
            Caption         =   "Nombre"
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
            Left            =   3000
            TabIndex        =   27
            Top             =   120
            Width           =   852
         End
         Begin VB.Label Label1 
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
            Index           =   1
            Left            =   0
            TabIndex        =   26
            Top             =   120
            Width           =   1452
         End
         Begin VB.Label Label1 
            Caption         =   "Destino"
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
            Index           =   28
            Left            =   3000
            TabIndex        =   25
            Top             =   840
            Width           =   852
         End
         Begin VB.Label lblFondoDisplay 
            Caption         =   "Canal"
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
            Left            =   0
            TabIndex        =   24
            Top             =   2400
            Width           =   735
         End
      End
      Begin XtremeSuiteControls.FlatEdit txtObservaciones 
         Height          =   1095
         Left            =   1080
         TabIndex        =   45
         Top             =   5280
         Width           =   7215
         _Version        =   1572864
         _ExtentX        =   12726
         _ExtentY        =   1931
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
         ScrollBars      =   2
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.CheckBox chkDeducPlanilla 
         Height          =   252
         Left            =   -69760
         TabIndex        =   48
         Top             =   4080
         Visible         =   0   'False
         Width           =   2172
         _Version        =   1572864
         _ExtentX        =   3831
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Deducir por Planilla"
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
         Alignment       =   1
      End
      Begin XtremeSuiteControls.GroupBox GroupBox1 
         Height          =   1095
         Left            =   -69760
         TabIndex        =   49
         Top             =   5640
         Visible         =   0   'False
         Width           =   5415
         _Version        =   1572864
         _ExtentX        =   9546
         _ExtentY        =   1926
         _StockProps     =   79
         Caption         =   "Recursos:"
         ForeColor       =   4210752
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
         Begin XtremeSuiteControls.ComboBox cboRecursos 
            Height          =   312
            Left            =   0
            TabIndex        =   50
            Top             =   360
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
         Begin XtremeSuiteControls.DateTimePicker dtpDesembolso 
            Height          =   312
            Left            =   3720
            TabIndex        =   82
            Top             =   360
            Width           =   1332
            _Version        =   1572864
            _ExtentX        =   2350
            _ExtentY        =   550
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
         Begin XtremeSuiteControls.FlatEdit txtDisponibleRecursos 
            Height          =   315
            Left            =   2400
            TabIndex        =   143
            Top             =   720
            Width           =   2655
            _Version        =   1572864
            _ExtentX        =   4683
            _ExtentY        =   556
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
         Begin VB.Image imgRecalculoRecurso 
            Height          =   252
            Left            =   5160
            Picture         =   "frmCR_SeguimientoTramites.frx":7E3E
            Stretch         =   -1  'True
            ToolTipText     =   "Recalcula Monto Disponible Recursos"
            Top             =   720
            Width           =   252
         End
         Begin VB.Label Label16 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Disponible Recurso:"
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
            Height          =   330
            Index           =   1
            Left            =   0
            TabIndex        =   51
            Top             =   720
            Width           =   2175
         End
         Begin VB.Image imgGuardaFecDesembolso 
            Height          =   252
            Left            =   5160
            Picture         =   "frmCR_SeguimientoTramites.frx":860D
            Stretch         =   -1  'True
            ToolTipText     =   "Guarda la Fecha de Desembolso"
            Top             =   360
            Width           =   252
         End
      End
      Begin MSComctlLib.ListView lswOpciones 
         Height          =   3012
         Left            =   -69880
         TabIndex        =   52
         Top             =   480
         Visible         =   0   'False
         Width           =   2856
         _ExtentX        =   5027
         _ExtentY        =   5318
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         HideColumnHeaders=   -1  'True
         FlatScrollBar   =   -1  'True
         FullRowSelect   =   -1  'True
         HotTracking     =   -1  'True
         HoverSelection  =   -1  'True
         _Version        =   393217
         SmallIcons      =   "imgFormalizacion"
         ForeColor       =   8388608
         BackColor       =   14737632
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
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Width           =   5029
         EndProperty
      End
      Begin XtremeSuiteControls.GroupBox fraEstadoFormalizacion 
         Height          =   2892
         Left            =   -64120
         TabIndex        =   53
         Top             =   3720
         Visible         =   0   'False
         Width           =   2652
         _Version        =   1572864
         _ExtentX        =   4678
         _ExtentY        =   5101
         _StockProps     =   79
         Caption         =   "Formalización:"
         ForeColor       =   4210752
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
         Begin XtremeSuiteControls.PushButton cmdAplicarFormalizacion 
            Height          =   612
            Left            =   960
            TabIndex        =   54
            Top             =   2160
            Width           =   1572
            _Version        =   1572864
            _ExtentX        =   2773
            _ExtentY        =   1080
            _StockProps     =   79
            Caption         =   "&Aplicar"
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
            Picture         =   "frmCR_SeguimientoTramites.frx":8CE5
            ImageAlignment  =   4
         End
         Begin XtremeSuiteControls.RadioButton optFormalizacion 
            Height          =   252
            Index           =   0
            Left            =   960
            TabIndex        =   55
            Top             =   1320
            Width           =   1452
            _Version        =   1572864
            _ExtentX        =   2561
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Formalizar"
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
            Alignment       =   1
         End
         Begin XtremeSuiteControls.RadioButton optFormalizacion 
            Height          =   252
            Index           =   1
            Left            =   960
            TabIndex        =   56
            Top             =   1680
            Width           =   1452
            _Version        =   1572864
            _ExtentX        =   2561
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Anular"
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
         Begin XtremeSuiteControls.DateTimePicker dtpFechaFormalizacion 
            Height          =   312
            Left            =   1200
            TabIndex        =   83
            Top             =   360
            Width           =   1332
            _Version        =   1572864
            _ExtentX        =   2350
            _ExtentY        =   550
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
            Enabled         =   0   'False
            CustomFormat    =   "dd/MM/yyyy"
            Format          =   3
         End
         Begin XtremeSuiteControls.FlatEdit txtTasaFacial 
            Height          =   315
            Left            =   1200
            TabIndex        =   133
            Top             =   720
            Width           =   1335
            _Version        =   1572864
            _ExtentX        =   2355
            _ExtentY        =   556
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
            Alignment       =   2
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin VB.Label Label16 
            Appearance      =   0  'Flat
            Caption         =   "Tasa Facial.:"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   312
            Index           =   3
            Left            =   0
            TabIndex        =   58
            Top             =   720
            Width           =   1332
         End
         Begin VB.Label Label16 
            Appearance      =   0  'Flat
            Caption         =   "Fecha.:"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   312
            Index           =   0
            Left            =   0
            TabIndex        =   57
            Top             =   360
            Width           =   1332
         End
      End
      Begin XtremeSuiteControls.ComboBox cboMes 
         Height          =   312
         Left            =   -65680
         TabIndex        =   59
         Top             =   4080
         Visible         =   0   'False
         Width           =   1332
         _Version        =   1572864
         _ExtentX        =   2355
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
      Begin XtremeSuiteControls.FlatEdit txtPagare 
         Height          =   312
         Left            =   -66160
         TabIndex        =   60
         Top             =   4800
         Visible         =   0   'False
         Width           =   1812
         _Version        =   1572864
         _ExtentX        =   3196
         _ExtentY        =   550
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
      Begin XtremeSuiteControls.FlatEdit txtDocumento 
         Height          =   312
         Left            =   -66160
         TabIndex        =   61
         Top             =   5160
         Visible         =   0   'False
         Width           =   1812
         _Version        =   1572864
         _ExtentX        =   3196
         _ExtentY        =   550
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
      Begin XtremeSuiteControls.FlatEdit txtAno 
         Height          =   330
         Left            =   -66160
         TabIndex        =   62
         Top             =   4080
         Visible         =   0   'False
         Width           =   492
         _Version        =   1572864
         _ExtentX        =   868
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
      Begin XtremeSuiteControls.CheckBox chkPrimera 
         Height          =   252
         Left            =   -69760
         TabIndex        =   63
         Top             =   4440
         Visible         =   0   'False
         Width           =   2172
         _Version        =   1572864
         _ExtentX        =   3831
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Deducir Primer Cuota"
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
         Alignment       =   1
      End
      Begin XtremeSuiteControls.CheckBox chkEnviarATesoreria 
         Height          =   252
         Left            =   -69760
         TabIndex        =   64
         Top             =   4800
         Visible         =   0   'False
         Width           =   2172
         _Version        =   1572864
         _ExtentX        =   3831
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Enviar a Tesorería"
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
         Alignment       =   1
      End
      Begin XtremeSuiteControls.ComboBox cboDeductora 
         Height          =   312
         Left            =   -68560
         TabIndex        =   85
         Top             =   3720
         Visible         =   0   'False
         Width           =   4212
         _Version        =   1572864
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
      Begin XtremeSuiteControls.PushButton btnTags 
         Height          =   612
         Left            =   -62920
         TabIndex        =   86
         ToolTipText     =   "Etiquetas de Seguimiento"
         Top             =   2760
         Visible         =   0   'False
         Width           =   1452
         _Version        =   1572864
         _ExtentX        =   2561
         _ExtentY        =   1080
         _StockProps     =   79
         Caption         =   "Tag's"
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
         Picture         =   "frmCR_SeguimientoTramites.frx":940C
         ImageAlignment  =   4
      End
      Begin XtremeSuiteControls.FlatEdit txtRecibe 
         Height          =   312
         Left            =   -68320
         TabIndex        =   87
         Top             =   2760
         Visible         =   0   'False
         Width           =   2532
         _Version        =   1572864
         _ExtentX        =   4466
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
         Alignment       =   2
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtFechaRec 
         Height          =   312
         Left            =   -65680
         TabIndex        =   88
         Top             =   2760
         Visible         =   0   'False
         Width           =   2532
         _Version        =   1572864
         _ExtentX        =   4466
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
         Alignment       =   2
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtResoluciona 
         Height          =   312
         Left            =   -68320
         TabIndex        =   89
         Top             =   3120
         Visible         =   0   'False
         Width           =   2532
         _Version        =   1572864
         _ExtentX        =   4466
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
         Alignment       =   2
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtFechaRes 
         Height          =   312
         Left            =   -65680
         TabIndex        =   90
         Top             =   3120
         Visible         =   0   'False
         Width           =   2532
         _Version        =   1572864
         _ExtentX        =   4466
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
         Alignment       =   2
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtFormaliza 
         Height          =   312
         Left            =   -68320
         TabIndex        =   91
         Top             =   3480
         Visible         =   0   'False
         Width           =   2532
         _Version        =   1572864
         _ExtentX        =   4466
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
         Alignment       =   2
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtFechaFor 
         Height          =   312
         Left            =   -65680
         TabIndex        =   92
         Top             =   3480
         Visible         =   0   'False
         Width           =   2532
         _Version        =   1572864
         _ExtentX        =   4466
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
         Alignment       =   2
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtTesoreria 
         Height          =   312
         Left            =   -68320
         TabIndex        =   93
         Top             =   3840
         Visible         =   0   'False
         Width           =   2532
         _Version        =   1572864
         _ExtentX        =   4466
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
         Alignment       =   2
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtFechaTes 
         Height          =   312
         Left            =   -65680
         TabIndex        =   94
         Top             =   3840
         Visible         =   0   'False
         Width           =   2532
         _Version        =   1572864
         _ExtentX        =   4466
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
         Alignment       =   2
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtAutorizada 
         Height          =   312
         Left            =   -68320
         TabIndex        =   95
         Top             =   4320
         Visible         =   0   'False
         Width           =   2532
         _Version        =   1572864
         _ExtentX        =   4466
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
         Alignment       =   2
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtFechaAuto 
         Height          =   312
         Left            =   -65680
         TabIndex        =   96
         Top             =   4320
         Visible         =   0   'False
         Width           =   2532
         _Version        =   1572864
         _ExtentX        =   4466
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
         Alignment       =   2
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtAutorizaNota 
         Height          =   1332
         Left            =   -68320
         TabIndex        =   97
         Top             =   4680
         Visible         =   0   'False
         Width           =   5172
         _Version        =   1572864
         _ExtentX        =   9123
         _ExtentY        =   2350
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
         MultiLine       =   -1  'True
         ScrollBars      =   2
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtConCedula 
         Height          =   312
         Left            =   -68680
         TabIndex        =   100
         Top             =   120
         Visible         =   0   'False
         Width           =   1692
         _Version        =   1572864
         _ExtentX        =   2984
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
         Alignment       =   2
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtConNombre 
         Height          =   312
         Left            =   -67000
         TabIndex        =   101
         Top             =   120
         Visible         =   0   'False
         Width           =   5532
         _Version        =   1572864
         _ExtentX        =   9758
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtMntNoGravable 
         Height          =   312
         Left            =   -66640
         TabIndex        =   102
         Top             =   1080
         Visible         =   0   'False
         Width           =   1812
         _Version        =   1572864
         _ExtentX        =   3196
         _ExtentY        =   550
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
      Begin XtremeSuiteControls.ComboBox cboFrecuencia 
         Height          =   312
         Left            =   -66160
         TabIndex        =   110
         Top             =   4440
         Visible         =   0   'False
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
      Begin XtremeSuiteControls.CheckBox chkPagareManual 
         Height          =   255
         Left            =   4320
         TabIndex        =   138
         Top             =   6480
         Width           =   1815
         _Version        =   1572864
         _ExtentX        =   3201
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Pagaré Manual"
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
      Begin XtremeSuiteControls.FlatEdit txtFormularioId 
         Height          =   315
         Left            =   6480
         TabIndex        =   139
         Top             =   6480
         Width           =   1815
         _Version        =   1572864
         _ExtentX        =   3201
         _ExtentY        =   556
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
         Alignment       =   2
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.CheckBox chkTrasladoSalario 
         Height          =   255
         Left            =   -69760
         TabIndex        =   142
         Top             =   5160
         Visible         =   0   'False
         Width           =   2175
         _Version        =   1572864
         _ExtentX        =   3831
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Traslada Salario ?"
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
         Alignment       =   1
      End
      Begin VB.Label Label1 
         Caption         =   $"frmCR_SeguimientoTramites.frx":9B2C
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   492
         Index           =   6
         Left            =   -69640
         TabIndex        =   104
         Top             =   120
         Visible         =   0   'False
         Width           =   7452
      End
      Begin VB.Label Label1 
         Caption         =   "Monto no Gravable del Crédito"
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
         Left            =   -69520
         TabIndex        =   103
         Top             =   1080
         Visible         =   0   'False
         Width           =   3372
      End
      Begin VB.Label Label5 
         Caption         =   "Deductora"
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
         Left            =   -69760
         TabIndex        =   84
         Top             =   3720
         Visible         =   0   'False
         Width           =   1212
      End
      Begin XtremeShortcutBar.ShortcutCaption TituloOpcionesSub 
         Height          =   360
         Left            =   -69880
         TabIndex        =   81
         Top             =   120
         Visible         =   0   'False
         Width           =   2892
         _Version        =   1572864
         _ExtentX        =   5101
         _ExtentY        =   635
         _StockProps     =   14
         Caption         =   "Opciones"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SubItemCaption  =   -1  'True
         VisualTheme     =   6
      End
      Begin VB.Label Label1 
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
         Left            =   -69880
         TabIndex        =   74
         Top             =   120
         Visible         =   0   'False
         Width           =   1212
      End
      Begin VB.Image imgPoolHistorico 
         Height          =   240
         Index           =   5
         Left            =   -69880
         Picture         =   "frmCR_SeguimientoTramites.frx":9BD7
         Stretch         =   -1  'True
         Top             =   4680
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Image imgPoolHistorico 
         Height          =   240
         Index           =   4
         Left            =   -69880
         Picture         =   "frmCR_SeguimientoTramites.frx":A3A6
         Stretch         =   -1  'True
         Top             =   4320
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Image imgPoolHistorico 
         Height          =   240
         Index           =   3
         Left            =   -69880
         Picture         =   "frmCR_SeguimientoTramites.frx":AB30
         Stretch         =   -1  'True
         Top             =   3840
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Image imgPoolHistorico 
         Height          =   240
         Index           =   2
         Left            =   -69880
         Picture         =   "frmCR_SeguimientoTramites.frx":11382
         Stretch         =   -1  'True
         Top             =   3000
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Image imgPoolHistorico 
         Height          =   240
         Index           =   1
         Left            =   -69880
         Picture         =   "frmCR_SeguimientoTramites.frx":11AE4
         Stretch         =   -1  'True
         Top             =   2640
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Image imgPoolHistorico 
         Height          =   240
         Index           =   0
         Left            =   -69880
         Picture         =   "frmCR_SeguimientoTramites.frx":12448
         Stretch         =   -1  'True
         Top             =   3480
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Nota"
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
         Index           =   5
         Left            =   -69520
         TabIndex        =   73
         Top             =   4680
         Visible         =   0   'False
         Width           =   1212
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Autorizada"
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
         Index           =   4
         Left            =   -69520
         TabIndex        =   72
         Top             =   4320
         Visible         =   0   'False
         Width           =   1212
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Tesoreria"
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
         Index           =   3
         Left            =   -69520
         TabIndex        =   71
         Top             =   3840
         Visible         =   0   'False
         Width           =   1212
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Formaliza"
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
         Index           =   2
         Left            =   -69520
         TabIndex        =   70
         Top             =   3480
         Visible         =   0   'False
         Width           =   1212
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Resoluciona"
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
         Index           =   1
         Left            =   -69520
         TabIndex        =   69
         Top             =   3120
         Visible         =   0   'False
         Width           =   1212
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Recibe"
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
         Index           =   0
         Left            =   -69520
         TabIndex        =   68
         Top             =   2760
         Visible         =   0   'False
         Width           =   1212
      End
      Begin VB.Label Label16 
         Appearance      =   0  'Flat
         Caption         =   "No. Pagaré"
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
         Index           =   2
         Left            =   -67360
         TabIndex        =   67
         Top             =   4800
         Visible         =   0   'False
         Width           =   1212
      End
      Begin VB.Label Label2 
         Caption         =   "Primer Deduc."
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
         Left            =   -67360
         TabIndex        =   66
         Top             =   4080
         Visible         =   0   'False
         Width           =   1212
      End
      Begin VB.Label Label4 
         Caption         =   "Doc. Ref."
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
         Left            =   -67360
         TabIndex        =   65
         Top             =   5160
         Visible         =   0   'False
         Width           =   1092
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
         Index           =   19
         Left            =   120
         TabIndex        =   46
         Top             =   5280
         Width           =   855
      End
   End
   Begin VB.Timer TimerX 
      Interval        =   10
      Left            =   9600
      Top             =   0
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   6600
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_SeguimientoTramites.frx":12BF1
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_SeguimientoTramites.frx":13361
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_SeguimientoTramites.frx":13B3F
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_SeguimientoTramites.frx":14207
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_SeguimientoTramites.frx":14BF3
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_SeguimientoTramites.frx":15365
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_SeguimientoTramites.frx":15AD7
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_SeguimientoTramites.frx":16249
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_SeguimientoTramites.frx":169EC
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_SeguimientoTramites.frx":171CD
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlbPrincipal 
      Height          =   264
      Left            =   9480
      TabIndex        =   2
      Top             =   600
      Visible         =   0   'False
      Width           =   2280
      _ExtentX        =   4022
      _ExtentY        =   476
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Style           =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   6
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "nuevo"
            Object.ToolTipText     =   "Inserta (Agrega) un registro nuevo a la Base de Datos"
            Object.Tag             =   "1"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "editar"
            Object.ToolTipText     =   "Modifica (Edita) el registro en pantalla"
            Object.Tag             =   "1"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "guardar"
            Object.ToolTipText     =   "Guarda la información del registro en la base de datos"
            Object.Tag             =   "1"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "deshacer"
            Object.ToolTipText     =   "Deshace toda modificación realizada recientemente en el registro actual"
            Object.Tag             =   "1"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "reportes"
            Object.ToolTipText     =   "Imprime el listado seleccionado"
            Object.Tag             =   "1"
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   13
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "RepActas"
                  Text            =   "Actas"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "RepPreAnalisis"
                  Text            =   "Estudios"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "RepGarantia"
                  Text            =   "Garantía"
               EndProperty
               BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "-"
               EndProperty
               BeginProperty ButtonMenu5 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "repBoleta"
                  Text            =   "Boleta"
               EndProperty
               BeginProperty ButtonMenu6 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Cheques"
                  Text            =   "Boleta de Cheques"
               EndProperty
               BeginProperty ButtonMenu7 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "RecRef"
                  Text            =   "Recibos de Refundición"
               EndProperty
               BeginProperty ButtonMenu8 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Sep1"
                  Text            =   "-"
               EndProperty
               BeginProperty ButtonMenu9 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Solicitud"
                  Text            =   "Solicitud de Crédito"
               EndProperty
               BeginProperty ButtonMenu10 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Requisitos"
                  Text            =   "Boleta de Requisitos"
               EndProperty
               BeginProperty ButtonMenu11 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Caratula"
                  Text            =   "Carátula para Expediente"
               EndProperty
               BeginProperty ButtonMenu12 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Sep2"
                  Text            =   "-"
               EndProperty
               BeginProperty ButtonMenu13 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "Autorizacion"
                  Text            =   "Autorización de Deducción"
               EndProperty
            EndProperty
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   7200
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_SeguimientoTramites.frx":179BC
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_SeguimientoTramites.frx":1E21E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_SeguimientoTramites.frx":24A80
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_SeguimientoTramites.frx":2B2E2
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_SeguimientoTramites.frx":31B44
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_SeguimientoTramites.frx":383A6
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_SeguimientoTramites.frx":3EC08
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_SeguimientoTramites.frx":4546A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgFormalizacion 
      Left            =   6000
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_SeguimientoTramites.frx":4BCCC
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_SeguimientoTramites.frx":4C46F
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin XtremeSuiteControls.FlatEdit txtOperacion 
      Height          =   372
      Left            =   2760
      TabIndex        =   3
      Top             =   120
      Width           =   2052
      _Version        =   1572864
      _ExtentX        =   3619
      _ExtentY        =   656
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
   Begin MSComctlLib.ImageList imlTaskPanelIcons 
      Left            =   7800
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   65280
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   16
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_SeguimientoTramites.frx":4CC50
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_SeguimientoTramites.frx":4CEC5
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_SeguimientoTramites.frx":4D15E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_SeguimientoTramites.frx":4D2E1
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_SeguimientoTramites.frx":4D479
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_SeguimientoTramites.frx":4D621
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_SeguimientoTramites.frx":4D7C3
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_SeguimientoTramites.frx":4D94E
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_SeguimientoTramites.frx":4DAD9
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_SeguimientoTramites.frx":4DBDD
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_SeguimientoTramites.frx":4DE6C
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_SeguimientoTramites.frx":4DF78
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_SeguimientoTramites.frx":4E1F4
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_SeguimientoTramites.frx":4E2BA
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_SeguimientoTramites.frx":4E45A
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_SeguimientoTramites.frx":4E5F6
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComCtl2.FlatScrollBar FlatScrollBar 
      Height          =   252
      Left            =   4920
      TabIndex        =   109
      Top             =   120
      Width           =   492
      _ExtentX        =   873
      _ExtentY        =   450
      _Version        =   393216
      Arrows          =   65536
      Orientation     =   1638401
   End
   Begin XtremeSuiteControls.PushButton btnBarra 
      Height          =   315
      Index           =   0
      Left            =   8760
      TabIndex        =   111
      ToolTipText     =   "Nuevo"
      Top             =   990
      Width           =   1095
      _Version        =   1572864
      _ExtentX        =   1926
      _ExtentY        =   550
      _StockProps     =   79
      Caption         =   "Nuevo"
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
      Picture         =   "frmCR_SeguimientoTramites.frx":4E6A1
      ImageAlignment  =   4
   End
   Begin XtremeSuiteControls.PushButton btnBarra 
      Height          =   315
      Index           =   1
      Left            =   9840
      TabIndex        =   112
      ToolTipText     =   "Editar"
      Top             =   990
      Width           =   375
      _Version        =   1572864
      _ExtentX        =   656
      _ExtentY        =   550
      _StockProps     =   79
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
      Picture         =   "frmCR_SeguimientoTramites.frx":4ECD3
      ImageAlignment  =   6
   End
   Begin XtremeSuiteControls.PushButton btnBarra 
      Height          =   315
      Index           =   3
      Left            =   10200
      TabIndex        =   113
      ToolTipText     =   "Guardar"
      Top             =   990
      Width           =   375
      _Version        =   1572864
      _ExtentX        =   656
      _ExtentY        =   550
      _StockProps     =   79
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
      Picture         =   "frmCR_SeguimientoTramites.frx":4F2CE
      ImageAlignment  =   6
   End
   Begin XtremeSuiteControls.PushButton btnBarra 
      Height          =   315
      Index           =   4
      Left            =   10560
      TabIndex        =   114
      ToolTipText     =   "Deshacer"
      Top             =   990
      Width           =   375
      _Version        =   1572864
      _ExtentX        =   656
      _ExtentY        =   550
      _StockProps     =   79
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
      Picture         =   "frmCR_SeguimientoTramites.frx":4F9FF
      ImageAlignment  =   6
   End
   Begin XtremeSuiteControls.PushButton btnBarra 
      Height          =   315
      Index           =   5
      Left            =   10920
      TabIndex        =   115
      ToolTipText     =   "Reporte"
      Top             =   990
      Width           =   375
      _Version        =   1572864
      _ExtentX        =   656
      _ExtentY        =   550
      _StockProps     =   79
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
      Picture         =   "frmCR_SeguimientoTramites.frx":500FF
      ImageAlignment  =   6
   End
   Begin XtremeSuiteControls.PushButton btnAdjuntos 
      Height          =   330
      Left            =   11280
      TabIndex        =   136
      ToolTipText     =   "Adjuntar Documentos"
      Top             =   990
      Width           =   495
      _Version        =   1572864
      _ExtentX        =   873
      _ExtentY        =   582
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
      Picture         =   "frmCR_SeguimientoTramites.frx":50806
   End
   Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
      Height          =   372
      Index           =   0
      Left            =   0
      TabIndex        =   108
      Top             =   960
      Width           =   2772
      _Version        =   1572864
      _ExtentX        =   4890
      _ExtentY        =   656
      _StockProps     =   14
      Caption         =   "Detalle:"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   10.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SubItemCaption  =   -1  'True
   End
   Begin XtremeShortcutBar.ShortcutCaption TituloOpcion 
      Height          =   375
      Left            =   2760
      TabIndex        =   75
      Top             =   960
      Width           =   9015
      _Version        =   1572864
      _ExtentX        =   15901
      _ExtentY        =   661
      _StockProps     =   14
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   10.5
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Image imgTags 
      Height          =   240
      Left            =   2880
      Picture         =   "frmCR_SeguimientoTramites.frx":5088F
      Stretch         =   -1  'True
      Top             =   636
      Width           =   240
   End
   Begin VB.Image imgHistorico 
      Height          =   240
      Left            =   3240
      Picture         =   "frmCR_SeguimientoTramites.frx":51015
      Stretch         =   -1  'True
      Top             =   636
      Width           =   240
   End
   Begin VB.Image imgConsulta 
      Height          =   240
      Left            =   3600
      Picture         =   "frmCR_SeguimientoTramites.frx":517E4
      Stretch         =   -1  'True
      Top             =   648
      Width           =   240
   End
   Begin ComctlLib.ImageList imgIconosEstados 
      Left            =   8400
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      UseMaskColor    =   0   'False
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   8
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmCR_SeguimientoTramites.frx":521F2
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmCR_SeguimientoTramites.frx":52A44
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmCR_SeguimientoTramites.frx":53296
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmCR_SeguimientoTramites.frx":53AE8
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmCR_SeguimientoTramites.frx":5433A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmCR_SeguimientoTramites.frx":54B8C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmCR_SeguimientoTramites.frx":553DE
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmCR_SeguimientoTramites.frx":55C30
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Image imgEstado 
      Height          =   240
      Left            =   3960
      Picture         =   "frmCR_SeguimientoTramites.frx":56482
      Stretch         =   -1  'True
      Top             =   636
      Width           =   240
   End
   Begin VB.Label lblNombre 
      BackStyle       =   0  'Transparent
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   252
      Left            =   5640
      TabIndex        =   1
      Top             =   156
      Width           =   5412
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Operación"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   372
      Index           =   0
      Left            =   1320
      TabIndex        =   0
      Top             =   120
      Width           =   1332
   End
   Begin VB.Image imgBanner 
      Height          =   972
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   20892
   End
End
Attribute VB_Name = "frmCR_SeguimientoTramites"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim vMensaje                As String  'Envia Mensajes en Fallas de Verificacion
Dim vEdita                  As Boolean 'Indica si se esta actualizando o insertando
Dim vPasaFormalizacion      As Boolean 'Indica si una formalizacion normal se procesa o no
Dim vDocumentoFormalizacion As Boolean 'Indica si se debe de generar una nota de debito
Dim vPaso                   As Boolean 'Para que Tes_Bancos Click cbo lo ignore
Dim vScroll                 As Boolean, vOperacionLoad As Boolean
Dim vFechaSistema           As Date
'Por incluir una formalizacion que no pasa a Tesoreria o el Monto Girado es Cero

Dim mFrecuenciaPago As String

Const Id_TaskItem_Recepcion = 0
Const Id_TaskItem_Formalizacion = 1
Const Id_TaskItem_Estudio = 2
Const Id_TaskItem_Garantia = 3
Const Id_TaskItem_Coberturas = 4
Const Id_TaskItem_Desembolsos = 5
Const Id_TaskItem_Seguimiento = 6
Const Id_TaskItem_Polizas = 7
Const Id_TaskItem_Requisitos = 8
Const Id_TaskItem_Cargos = 9
Const Id_TaskItem_Causas = 10
Const Id_TaskItem_PlanPagos = 11
Const Id_TaskItem_DatosPersonales = 12
Const Id_TaskItem_PreCalculo = 13
Const Id_TaskItem_Firmas = 14
Const Id_TaskItem_MntNoGravable = 15
Const Id_TaskItem_Historial = 16


Private Sub sbBarra_Accion(pAccion As String)

btnBarra.Item(0).Enabled = False 'Nuevo
btnBarra.Item(1).Enabled = False 'Editar
'btnBarra.Item(2).Enabled = False 'Borrar
btnBarra.Item(3).Enabled = False 'Guardar
btnBarra.Item(4).Enabled = False 'Deshacer
btnBarra.Item(5).Enabled = False 'Reporte

Select Case UCase(pAccion)
    Case "NUEVO"
        btnBarra.Item(0).Enabled = True 'Nuevo
    
    Case "EDITAR", "EDICION"
    
        btnBarra.Item(3).Enabled = True 'Guardar
        btnBarra.Item(4).Enabled = True 'Deshacer
    
    Case "ACTIVO"
        btnBarra.Item(0).Enabled = True 'Nuevo
        btnBarra.Item(1).Enabled = True 'Editar
'        btnBarra.Item(2).Enabled = True 'Borrar
        btnBarra.Item(5).Enabled = True 'Reporte
End Select

End Sub


Private Sub sbTaskPanel_Load()


    Dim Group As TaskPanelGroup
    Dim Item As TaskPanelGroupItem
    
    tpMain.VisualTheme = xtpTaskPanelThemeOffice2016
  
    Set Group = tpMain.Groups.Add(0, "Registro")
    Group.ToolTip = "Información Principal para el Registro y Formalización del Crédito"
    Group.Special = True

    
    Group.Items.Add Id_TaskItem_Recepcion, "Recepción", xtpTaskItemTypeLink, 4
    Group.Items.Add Id_TaskItem_Formalizacion, "Formalización", xtpTaskItemTypeLink, 2
    
    Set Group = tpMain.Groups.Add(0, "Seguimiento")
    Group.ToolTip = "Datos adicionales para el proceso de formalización o control de desembolsos"
    
    Group.Items.Add Id_TaskItem_Historial, "Historial", xtpTaskItemTypeLink, 3
    Group.Items.Add Id_TaskItem_Polizas, "Pólizas", xtpTaskItemTypeLink, 13
    Group.Items.Add Id_TaskItem_Cargos, "Cargos", xtpTaskItemTypeLink, 12
    
    
    Set Group = tpMain.Groups.Add(0, "Garantía")
    Group.Items.Add Id_TaskItem_Garantia, "Garantía", xtpTaskItemTypeLink, 8
    Group.Items.Add Id_TaskItem_Coberturas, "...Coberturas", xtpTaskItemTypeLink, 6
    Group.Items.Add Id_TaskItem_Desembolsos, "...Desembolsos", xtpTaskItemTypeLink, 6
    
    Set Group = tpMain.Groups.Add(0, "Cumplimiento")
    Group.Items.Add Id_TaskItem_Firmas, "Firmas", xtpTaskItemTypeLink, 12
    Group.Items.Add Id_TaskItem_Requisitos, "Requisitos", xtpTaskItemTypeLink, 12
    Group.Items.Add Id_TaskItem_Causas, "Causas", xtpTaskItemTypeLink, 12
    
    Set Group = tpMain.Groups.Add(0, "Info adicional")
    Group.Expanded = False
    Group.Items.Add Id_TaskItem_Estudio, "Estudio de Crédito", xtpTaskItemTypeLink, 10
    Group.Items.Add Id_TaskItem_DatosPersonales, "Datos Personales", xtpTaskItemTypeLink, 1
    Group.Items.Add Id_TaskItem_PlanPagos, "Plan de Pagos", xtpTaskItemTypeLink, 7
    Group.Items.Add Id_TaskItem_PreCalculo, "Pre - Cálculo", xtpTaskItemTypeLink, 2
    Group.Items.Add Id_TaskItem_MntNoGravable, "Monto No Gravable", xtpTaskItemTypeLink, 16
    
   
    tpMain.SetImageList imlTaskPanelIcons
    
'    tpMain.SetMargins 5, 5, 5, 5, 5
    

End Sub




Private Sub btnAdjuntos_Click()
 gGA.Modulo = "CR_01"
 gGA.Llave_01 = txtCedula.Text
 gGA.Llave_02 = txtOperacion.Text
 gGA.Llave_03 = txtCodigo.Text
 
 Call sbFormsCall("frmGA_Documentos", vbModal, , , False, Me, True)
End Sub

Private Sub btnBarra_Click(Index As Integer)


Select Case Index
 Case 0  'nuevo
  txtOperacion.Text = ""
  txtOperacion.Enabled = False
  Call sbLimpiaDatos
  
  Call Edicion(1)
  Call sbBarra_Accion("edicion")
  
  fraOperacion.Enabled = True
  txtObservaciones.Locked = False
  
  vEdita = False
  
  txtCedula.SetFocus
'  Call sbCargaCombos
  
  
  
 Case 1 'editar
  If Operacion.Operacion > 0 Then 'And Operacion.Estado = "A" Then
      vEdita = True
      Call Edicion(1)
    
      Call sbBarra_Accion("edicion")
      
      txtOperacion.Enabled = False
      fraOperacion.Enabled = True
      txtObservaciones.Locked = False

      If txtCedula.Enabled Then
         tcMain.Item(0).Selected = True
         txtCedula.SetFocus
      End If

  End If
 
 Case 3 'guardar
  
  If fxVerificaRecepcion Then
    'Verificar si se cambio el codigo
    If Trim(txtCodigo) <> Operacion.Codigo Then Call ActualizaCodigoOperacion
    Call sbGuardarSolicitud
    Call Edicion(0)
    Call sbCargaOperacion
    txtOperacion.Enabled = True
    
    Call sbBarra_Accion("activo")
    
    fraOperacion.Enabled = False
    txtObservaciones.Locked = True

    
    If vEdita = False Then
        tcMain.Item(0).Selected = True
        'Datos Personales
        Call sbTaskPanel_Accion(Id_TaskItem_DatosPersonales)
    End If
    
    If vEdita = False And cboGarantia.ItemData(cboGarantia.ListIndex) = "F" Then
        tcMain.Item(0).Selected = True
        'Fiadores
        'Call btnOpciones_Click(1)
        Call sbTaskPanel_Accion(Id_TaskItem_Garantia)
    
    End If
    
    If vEdita = False Then
        'Requisitos
        Call sbTaskPanel_Accion(Id_TaskItem_Requisitos)
    
    End If
    
    If Operacion.EstadoSolicitud = "P" Or Operacion.EstadoSolicitud = "D" Then
        'Siempre verifica las causas, por si esta en Pendiente o Denegada
         Call sbTaskPanel_Accion(Id_TaskItem_Causas)
    End If
  
  Else
    MsgBox vMensaje, vbCritical
  End If
 
 Case 4 'deshacer
    txtOperacion.Enabled = True
    
    Call sbBarra_Accion("nuevo")
    
    fraOperacion.Enabled = False
    If txtOperacion <> "" Then Call sbCargaOperacion
    txtOperacion.SetFocus
 
 Case 5 'Reportes
   gbReportes.top = 1440
   gbReportes.Left = 5760
   gbReportes.Visible = IIf((gbReportes.Visible = True), False, True)
   
End Select


End Sub

Private Sub btnInforme_Click(Index As Integer)

If Index = 1 Then 'Exit
    gbReportes.Visible = False
    Exit Sub
End If


Me.MousePointer = vbHourglass

Select Case True
 Case rbInformes.Item(0).Value 'Boleta
    If Operacion.EstadoSolicitud = "F" Or Operacion.EstadoSolicitud = "N" Then
      Call sbCrdSGTBoletaFormaliza(Operacion.Operacion)
    Else
      MsgBox "La Operación # " & Operacion.Operacion & " No se encuentra formalizada", vbInformation
    End If
 
 Case rbInformes.Item(1).Value 'Garantia
   Call sbFormsCall("frmCR_GeneraGarantia", 1, , , False, Me)
 
 Case rbInformes.Item(2).Value 'Autorizacion de Deduccion
    Call sbCrdSGTAutorizacionDeduccion(Operacion.Operacion)
 
 
 Case rbInformes.Item(3).Value 'Solicitud
    Call sbCrdSGTBoletaSolicitud(Operacion.Operacion)
 
 Case rbInformes.Item(4).Value 'Recibos de Refundiciones
   If Operacion.EstadoSolicitud = "F" Or Operacion.EstadoSolicitud = "N" Then
     Call sbCrdSGTReciboRefundicion(Operacion.Operacion)
   End If
 
 Case rbInformes.Item(5).Value 'Caratula para Expediente
    Call sbCrdSGTCaratulaCredito(Operacion.Operacion)
 
 Case rbInformes.Item(6).Value 'Boleta de Requisitos
    Call sbCrdSGTBoletaRequisitos(Operacion.Operacion)
 
 Case rbInformes.Item(7).Value 'Boleta de Cheques
   If Operacion.EstadoSolicitud = "F" Or Operacion.EstadoSolicitud = "N" Then
     Call sbCrdSGTBoletaCK(Operacion.Operacion)
   End If
 
 
 Case rbInformes.Item(8).Value 'Actas de Resolucion
   Call sbFormsCall("frmCR_SolCreacionAgenda", 1, , , False, Me)

 Case rbInformes.Item(9).Value 'Estudio de Credito
   Call sbFormsCall("frmCR_SolicitudesPreAnalisis", 1, , , False, Me)
 
   
    
End Select


Me.MousePointer = vbDefault



End Sub

Private Sub btnMtnNoGravable_Click()
Dim strSQL As String

On Error GoTo vError

If Operacion.Operacion = 0 Then
    MsgBox "No se ha indicado ninguna operación?", vbExclamation
    Exit Sub
End If

'If Operacion.EstadoSolicitud = "F" Then
'    MsgBox "No se puede actualizar porque la operación ya fue formalizada", vbExclamation
'    Exit Sub
'End If

If CCur(txtMntNoGravable.Text) > CCur(txtMonto.Text) Then
    MsgBox "El Monto No Gravable supera al monto del crédito!", vbExclamation
    Exit Sub
End If

strSQL = "update reg_creditos set IVA_Monto = " & CCur(txtMntNoGravable.Text) _
       & " where id_solicitud = " & Operacion.Operacion
Call ConectionExecute(strSQL)


Select Case Operacion.EstadoSolicitud
  Case "R", "P", "A"
    Call sbCargosAdicionales(Operacion.Operacion, txtCodigo, CCur(txtMonto))
  Case Else
    'No Actualiza Cargos
End Select

Call sbBitacoraCredito("28", "Monto: " & txtMntNoGravable.Text, "C", txtOperacion, txtCodigo, "Estado de la Operación [ " & Operacion.EstadoSolicitud & " ]")


MsgBox "Monto No Gravable, actualizado satisfactoriamente!", vbInformation

Call sbCargaOperacion

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub btnTags_Click()
        If Operacion.Operacion > 0 Then
           Call sbFormsCall("frmCR_SeguimientoEtiquetas", 1, , , False, Me)
        End If
End Sub

Private Sub cboActividad_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then cboCanal.SetFocus

If KeyCode = vbKeyF4 Then
   gBusquedas.Convertir = "N"
   gBusquedas.Columna = "cod_actividad"
   gBusquedas.Orden = "cod_actividad"
   gBusquedas.Consulta = "select cod_actividad,descripcion from AFI_ACTIVIDADES_ECO"
   gBusquedas.Filtro = " and activa = 1"
   frmBusquedas.Show vbModal
   cboActividad.Text = gBusquedas.Resultado2
End If
End Sub

Private Sub cboBanco_Click()

If vPaso Or cboBanco.ListCount = 0 Or cboBanco.Text = "" Then Exit Sub

Dim strSQL As String

On Error GoTo vError

strSQL = "exec spSys_Cuentas_Bancarias '" & txtCedula.Text & "'," & cboBanco.ItemData(cboBanco.ListIndex) & ",1"
Call sbCbo_Llena_New(cboCuenta, strSQL, False, True)

vError:


End Sub

Private Sub cboBanco_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
  cboTipoDocumento.SetFocus
End If
End Sub


Private Sub cboCanal_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtMonto.SetFocus

If KeyCode = vbKeyF4 Then
   gBusquedas.Convertir = "N"
   gBusquedas.Columna = "Canal_Tipo"
   gBusquedas.Orden = "Canal_Tipo"
   gBusquedas.Consulta = "select Canal_Tipo,Descripcion from AFI_CANALES_TIPOS"
   gBusquedas.Filtro = " and activo = 1"
   frmBusquedas.Show vbModal
   cboCanal.Text = gBusquedas.Resultado2
End If
End Sub

Private Sub cboComite_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then cboEstado.SetFocus
End Sub


Private Sub cboCuenta_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtObservaciones.SetFocus

End Sub

Private Sub cboDeductora_Click()

If vPaso Then Exit Sub

On Error GoTo vError

Dim strSQL As String, rs As New ADODB.Recordset
Dim vProceso As Currency, pProcesoClean As Long

strSQL = "select rtrim(descripcion) as 'Descripcion', isnull(Frecuencia,'M') as 'Frecuencia_Id'" _
       & " from instituciones " _
       & " where cod_institucion = " & cboDeductora.ItemData(cboDeductora.ListIndex)
Call OpenRecordSet(rs, strSQL)
    mFrecuenciaPago = rs!Frecuencia_ID
rs.Close

cboFrecuencia.Clear
Select Case mFrecuenciaPago
    Case "M" 'Mensual
        cboFrecuencia.AddItem "Mensual"
        cboFrecuencia.ItemData(cboFrecuencia.ListCount - 1) = "0"
        cboFrecuencia.Text = "Mensual"
    
    Case "Q" 'Quincenal
        cboFrecuencia.AddItem "1er Quincena"
        cboFrecuencia.ItemData(cboFrecuencia.ListCount - 1) = "1"
        cboFrecuencia.AddItem "2da Quincena"
        cboFrecuencia.ItemData(cboFrecuencia.ListCount - 1) = "2"
End Select
  
  
vProceso = fxPrimerDeduccion(Operacion.Codigo, cboDeductora.ItemData(cboDeductora.ListIndex))
pProcesoClean = vProceso

cboMes.Text = fxConvierteMES(Val(Mid(pProcesoClean, 5, 2)))
txtAno.Text = Mid(pProcesoClean, 1, 4)

If mFrecuenciaPago = "Q" Then
    If (vProceso - pProcesoClean) = 0.1 Then
        cboFrecuencia.Text = "1er Quincena"
    Else
        cboFrecuencia.Text = "2da Quincena"
    End If
End If

Exit Sub

vError:


End Sub

Private Sub cboDestino_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtPromotorNombre.SetFocus

If KeyCode = vbKeyF4 Then
   gBusquedas.Convertir = "N"
   gBusquedas.Columna = "D.cod_destino"
   gBusquedas.Orden = "D.Cod_Destino"
   gBusquedas.Consulta = "select D.cod_Destino,D.descripcion" _
                        & " from catalogo_destinos D inner join catalogo_destinosASG C on D.cod_destino = C.cod_destino"
   gBusquedas.Filtro = " and C.codigo = '" & txtCodigo.Text & "' "
   frmBusquedas.Show vbModal
   cboDestino.Text = Trim(gBusquedas.Resultado2)
End If

End Sub

Private Sub cboEstado_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then cboBanco.SetFocus
End Sub


Private Sub cboFondo_Click()
Dim strSQL As String, rs As New ADODB.Recordset

If vPaso Then Exit Sub
If cboFondo.ListCount <= 0 Then Exit Sub
If cboFondo.Text = "" Then Exit Sub


If cboGarantia.ItemData(cboGarantia.ListIndex) <> "Y" Then Exit Sub


'Carga Contratos a Plazo

vPaso = True

strSQL = "select cod_contrato,Tasa_Referencia,Aportes, isnull(FECHA_CORTE, getdate()) as 'FECHA_CORTE'" _
       & " from fnd_contratos" _
       & " where cod_plan = '" & cboFondo.ItemData(cboFondo.ListIndex) _
       & "' and estado = 'A' and cedula = '" & txtCedula.Text & "'"
Call OpenRecordSet(rs, strSQL)
cboFondoContrato.Clear
Do While Not rs.EOF
  cboFondoContrato.AddItem "[Cnt: " & rs!COD_CONTRATO & "] [Tasa: " & rs!TASA_REFERENCIA & "] [I: " & Format(rs!Aportes, "Standard") _
        & "] [V: " & Format(rs!fecha_corte, "yyyy-mm-dd") & "]"
'  cboFondoContrato.ItemData(cboFondoContrato.NewIndex) = rs!cod_contrato
  cboFondoContrato.ItemData(cboFondoContrato.ListCount - 1) = CStr(rs!COD_CONTRATO)

  rs.MoveNext
Loop
If rs.RecordCount > 0 Then
   rs.MoveFirst
   cboFondoContrato.Text = "[Cnt: " & rs!COD_CONTRATO & "] [Tasa: " & rs!TASA_REFERENCIA & "] [I: " & Format(rs!Aportes, "Standard") _
        & "] [V: " & Format(rs!fecha_corte, "yyyy-mm-dd") & "]"
End If
rs.Close
vPaso = False

If Not vOperacionLoad Then
    If cboFondoContrato.ListCount <= 0 Then
         strSQL = "exec spCRDGarantiaFNDCalculo '" & txtCedula & "','" & cboFondo.ItemData(cboFondo.ListIndex) & "',0"
    Else
         strSQL = "exec spCRDGarantiaFNDCalculo '" & txtCedula & "','" & cboFondo.ItemData(cboFondo.ListIndex) _
                & "'," & cboFondoContrato.ItemData(cboFondoContrato.ListIndex)
    End If
    
    Call OpenRecordSet(rs, strSQL)
    
    txtMonto.Text = Format(rs!Disponible, "Standard")
    If rs!AplicaTasa = 1 Then
        txtTasa.Text = Format(rs!Tasa, "Standard")
    End If
    
    If rs!AplicaPlazo = 1 Then
        txtPlazo.Text = rs!Plazo
    End If
    
    lblTasa.Tag = rs!AplicaTasa
    lblPlazo.Tag = rs!AplicaPlazo
    
    rs.Close
End If 'Operacion load

End Sub

Private Sub cboFondoContrato_Click()
Dim strSQL As String, rs As New ADODB.Recordset

If vPaso Then Exit Sub
If cboFondo.ListCount <= 0 Then Exit Sub
If cboFondo.Text = "" Then Exit Sub

If cboGarantia.ItemData(cboGarantia.ListIndex) <> "Y" Then Exit Sub

If cboFondoContrato.ListCount <= 0 Then
     strSQL = "exec spCRDGarantiaFNDCalculo '" & txtCedula & "','" & cboFondo.ItemData(cboFondo.ListIndex) & "',0"
Else
     strSQL = "exec spCRDGarantiaFNDCalculo '" & txtCedula & "','" & cboFondo.ItemData(cboFondo.ListIndex) _
            & "'," & cboFondoContrato.ItemData(cboFondoContrato.ListIndex)
End If

Call OpenRecordSet(rs, strSQL)

txtMonto.Text = Format(rs!Disponible, "Standard")
If rs!AplicaTasa = 1 Then
    txtTasa.Text = Format(rs!Tasa, "Standard")
End If

If rs!AplicaPlazo = 1 Then
    txtPlazo.Text = rs!Plazo
End If

lblTasa.Tag = rs!AplicaTasa
lblPlazo.Tag = rs!AplicaPlazo

rs.Close

End Sub



Private Sub cboGarantia_Click()
Dim strSQL As String, rs As New ADODB.Recordset
Dim mGarantiaForm As String

If vPaso Then Exit Sub
If cboGarantia.ListCount <= 0 Then Exit Sub
If cboGarantia.Text = "" Then Exit Sub

On Error GoTo vError

Me.MousePointer = vbHourglass


strSQL = "select FORMULARIO  From CRD_GARANTIA_TIPOS" _
       & " where garantia = '" & cboGarantia.ItemData(cboGarantia.ListIndex) & "'"
Call OpenRecordSet(rs, strSQL)
 mGarantiaForm = Trim(rs!Formulario)
rs.Close

cboFondo.Visible = False
cboFondoContrato.Visible = False
lblFondoDisplay.Item(0).Visible = False
lblFondoDisplay.Item(1).Visible = False

Operacion.PlazoBono = 0



Select Case mGarantiaForm
    Case "F01" 'Sobre Ahorros
        strSQL = "select dbo.fxCrdGarantiaPatMnt('" & txtCedula.Text & "','" & cboGarantia.ItemData(cboGarantia.ListIndex) & "', 'M') as 'Monto'" _
               & ", dbo.fxCrdTasaBonifica_New('" & txtCedula.Text & "','" & txtCodigo.Text & "', '" _
               & cboGarantia.ItemData(cboGarantia.ListIndex) & "', '" & cboDestino.ItemData(cboDestino.ListIndex) & "', " & IIf(txtPlazo.Text = "", 0, txtPlazo.Text) & ") as 'PtsBono'" _
               & ", dbo.fxCrdPlazoBonifica('" & txtCedula.Text & "','" & cboGarantia.ItemData(cboGarantia.ListIndex) & "') as 'PlazoBono'"
        Call OpenRecordSet(rs, strSQL)
          txtMonto.Text = Format(rs!Monto, "Standard")
          Operacion.TasaPtsBono = rs!PtsBono
          txtTasa.ToolTipText = "Pts Bonificación: " & rs!PtsBono
          If rs!PlazoBono > 0 Then
            txtPlazo.Text = rs!PlazoBono
            Operacion.PlazoBono = rs!PlazoBono
          End If
        rs.Close
    
    Case "F05" 'Fondos de Ahorros
        cboFondo.Visible = True
        cboFondoContrato.Visible = True
        lblFondoDisplay.Item(0).Visible = True
        lblFondoDisplay.Item(1).Visible = True
        Call cboFondo_Click

    Case "F06" 'Adelanto de Salario
        strSQL = "select dbo.fxCrdDisponibleAdelantoSalario('" & txtCedula.Text & "', 'M') as 'Monto'" _
               & ", dbo.fxCrdTasaBonifica_New('" & txtCedula.Text & "','" & txtCodigo.Text & "', '" _
               & cboGarantia.ItemData(cboGarantia.ListIndex) & "', '" & cboDestino.ItemData(cboDestino.ListIndex) & ", '" & IIf(txtPlazo.Text = "", 0, txtPlazo.Text) & ") as 'PtsBono'" _
               & ",dbo.fxCrdPlazoBonifica('" & txtCedula.Text & "','" & cboGarantia.ItemData(cboGarantia.ListIndex) & "') as 'PlazoBono'"
        Call OpenRecordSet(rs, strSQL)
          txtMonto.Text = Format(rs!Monto, "Standard")
          Operacion.TasaPtsBono = rs!PtsBono
          txtTasa.ToolTipText = "Pts Bonificación: " & rs!PtsBono
          If rs!PlazoBono > 0 Then
            txtPlazo.Text = rs!PlazoBono
            Operacion.PlazoBono = rs!PlazoBono
          End If
        rs.Close


    Case Else     'Otras Garantias
        strSQL = "select dbo.fxCrdTasaBonifica_New('" & txtCedula.Text & "','" & txtCodigo.Text & "', '" _
               & cboGarantia.ItemData(cboGarantia.ListIndex) & "', '" & cboDestino.ItemData(cboDestino.ListIndex) & "', " & IIf(txtPlazo.Text = "", 0, txtPlazo.Text) & ") as 'PtsBono'" _
             & ",dbo.fxCrdPlazoBonifica('" & txtCedula.Text & "','" & cboGarantia.ItemData(cboGarantia.ListIndex) & "') as 'PlazoBono'"
        Call OpenRecordSet(rs, strSQL)
          Operacion.TasaPtsBono = rs!PtsBono
          txtTasa.ToolTipText = "Pts Bonificación: " & rs!PtsBono
        
          If rs!PlazoBono > 0 Then
            txtPlazo.Text = rs!PlazoBono
            Operacion.PlazoBono = rs!PlazoBono
          End If
        
        rs.Close

End Select

''Fondos y Otros Ahorros
'If mGarantiaForm = "F05" Then
'    cboFondo.Visible = True
'    cboFondoContrato.Visible = True
'    lblFondoDisplay.Item(0).Visible = True
'    lblFondoDisplay.Item(1).Visible = True
'    Call cboFondo_Click
'End If
'
''Sobre Ahorros
'If mGarantiaForm = "F01" Then
'   strSQL = "select dbo.fxCrdGarantiaPatMnt('" & txtCedula.Text & "','" & cboGarantia.ItemData(cboGarantia.ListIndex) & "', 'M') as 'Monto'" _
'          & ",dbo.fxCrdTasaBonifica('" & txtCedula.Text & "','" & txtCodigo.Text & "','" & cboGarantia.ItemData(cboGarantia.ListIndex) & "') as 'PtsBono'" _
'          & ",dbo.fxCrdPlazoBonifica('" & txtCedula.Text & "','" & cboGarantia.ItemData(cboGarantia.ListIndex) & "') as 'PlazoBono'"
'   Call OpenRecordSet(rs, strSQL)
'     txtMonto.Text = Format(rs!Monto, "Standard")
'     Operacion.TasaPtsBono = rs!PtsBono
'     txtTasa.ToolTipText = "Pts Bonificación: " & rs!PtsBono
'     If rs!PlazoBono > 0 Then
'       txtPlazo.Text = rs!PlazoBono
'       Operacion.PlazoBono = rs!PlazoBono
'     End If
'   rs.Close
'
'Else
'   strSQL = "select dbo.fxCrdTasaBonifica('" & txtCedula.Text & "','" & txtCodigo.Text _
'        & "','" & cboGarantia.ItemData(cboGarantia.ListIndex) & "') as 'PtsBono'" _
'        & ",dbo.fxCrdPlazoBonifica('" & txtCedula.Text & "','" & cboGarantia.ItemData(cboGarantia.ListIndex) & "') as 'PlazoBono'"
'   Call OpenRecordSet(rs, strSQL)
'     Operacion.TasaPtsBono = rs!PtsBono
'     txtTasa.ToolTipText = "Pts Bonificación: " & rs!PtsBono
'
'     If rs!PlazoBono > 0 Then
'       txtPlazo.Text = rs!PlazoBono
'       Operacion.PlazoBono = rs!PlazoBono
'     End If
'
'   rs.Close
'
'End If

'Valida Montos, Tasas y Plazos
Call txtMonto_KeyPress(vbKeyReturn)

Me.MousePointer = vbDefault
Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub cboGarantia_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
 If KeyCode = vbKeyReturn Then cboDestino.SetFocus
End Sub

Private Sub cboRecursos_Click()

On Error GoTo vError

txtDisponibleRecursos.Text = "0"
Call imgRecalculoRecurso_Click


vError:
 Me.MousePointer = vbDefault
End Sub


Private Sub cboTipoDocumento_Click()

If vPaso Then Exit Sub
If cboTipoDocumento.ListCount = 0 Then Exit Sub


Dim pTipo As String


pTipo = fxTipoDocumento(cboTipoDocumento.Text)

lblProveedor.top = lblCuentaTitulo.top

If pTipo = "CP" Then
    lblCuentaTitulo.Visible = False
    lblProveedor.Visible = True
Else
    lblCuentaTitulo.Visible = True
    lblProveedor.Visible = False
End If

txtProveedorId.top = lblProveedor.top
txtProveedorNombre.top = lblProveedor.top

txtProveedorId.Visible = lblProveedor.Visible
txtProveedorNombre.Visible = lblProveedor.Visible


cboCuenta.Visible = lblCuentaTitulo.Visible

End Sub

Private Sub cboTipoDocumento_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then

    If cboCuenta.Visible Then cboCuenta.SetFocus
    If Not cboCuenta.Visible And txtPromotorId.Visible Then txtPromotorId.SetFocus

End If
End Sub
 

Private Sub sbFormalizar()
Dim rs As New ADODB.Recordset, strSQL As String
'Dim curRefunde As Currency, curRetencion As Currency, curDesembolsos As Currency
'Dim curIntDias As Currency, curInteres As Currency, curAmortiza As Currency
'Dim lngPriDeduc As Currency, FechaUltima As Currency, curPSD As Currency
'Dim curCuota As Currency, vFecha As Date, curTotal As Currency
'Dim curCargos As Currency, iMes As Integer, lngAnio As Long
'Dim vFechaCalculo As Date, vBoletaCK As Boolean, vTipoCobro As String, vDias As Long
'Dim vTransac As Boolean, vPuntosAdd As String, vTasaPiso As Currency, vBaseCalculo As String, vDiaPago As Integer
'
'Dim vPrimerCuota As Boolean


Dim lngPriDeduc As Currency
Me.MousePointer = vbHourglass
'
'
'curRefunde = 0
'curRetencion = 0
'curIntDias = 0
'curDesembolsos = 0
'curInteres = 0
'curAmortiza = 0
'curPSD = 0
'curCuota = 0
'curCargos = 0
'
'vPrimerCuota = False
'
'vBoletaCK = False
'vTransac = False
'vBaseCalculo = "01" '360/360
'vDiaPago = 32
'
'vFecha = fxFechaServidor
'
'
''Preguntar SI es TBP / revisa si utiliza TBP del Destino
''Extrae Dia de Pago y Base de Calculo
'
'vPuntosAdd = "NULL"
'
'strSQL = "select TBP_Utiliza,TBP_Adicional,Tasa_Destino,Base_Calculo,dbo.fxCRDPoliticaPago(dbo.MyGetdate()) as DiaPago" _
'       & " From catalogo where Codigo = '" & Operacion.Codigo & "'"
'Call OpenRecordSet(rs, strSQL)
'
'vBaseCalculo = Trim(rs!Base_Calculo)
'If chkDeducPlanilla.Value = vbChecked Then
'    vDiaPago = 32
'Else
'    vDiaPago = rs!DiaPago
'End If
'
'If rs!TBP_Utiliza = 1 Then
' vPuntosAdd = rs!TBP_Adicional
'Else
'  If rs!Tasa_Destino = 1 Then
'     rs.Close
'     strSQL = "select Tasa,TBP from catalogo_destinos where cod_destino = '" & cboDestino.ItemData(cboDestino.ListIndex) & "'"
'     Call OpenRecordSet(rs, strSQL)
'     If Not rs.EOF And Not rs.BOF Then
'       If rs!TBP = 1 Then
'           vPuntosAdd = rs!Tasa
'       End If
'     End If
'
'  End If 'Tasa_Destino
'
'End If
'rs.Close
'
'
''Extrae la Tasa Piso para la linea x Garantia (Cero si no tiene)
'vTasaPiso = 0
'strSQL = "select Tasa_Piso" _
'       & " From crd_catalogo_garantias where codigo = '" & Operacion.Codigo & "' and garantia = '" _
'       & cboGarantia.ItemData(cboGarantia.ListIndex) & "' and Utiliza_Tasa_Piso = 1"
'
'Call OpenRecordSet(rs, strSQL)
'If Not rs.EOF And Not rs.BOF Then
' vTasaPiso = rs!Tasa_Piso
'End If
'rs.Close
'
'
'strSQL = "update reg_creditos set fechaforp = '" & Format(vFecha, "yyyy/mm/dd") _
'       & "', fecha_inicio_calculo = '" & Format(dtpDesembolso.Value, "yyyy/mm/dd") _
'       & "', fecha_registro = dbo.MyGetdate(),cod_grupo = '" & cboRecursos.ItemData(cboRecursos.ListIndex) & "',Dia_Pago = " & vDiaPago _
'       & ", categoria_persona = dbo.fxCRDClasificacion('" & Operacion.Cedula & "',dbo.MyGetdate()),cod_oficina_f = '" & GLOBALES.gOficinaTitular _
'       & "', TBP_PuntosAdd = " & vPuntosAdd & ",Tasa_Piso = " & vTasaPiso & ",Base_Calculo = '" & vBaseCalculo _
'       & "',Tasa_Facial = " & CCur(txtTasaFacial.Text) _
'       & ", ind_deduce_Planilla = '" & IIf((chkDeducPlanilla.Value = vbChecked), "S", "N") & "'" _
'       & ", cod_deductora = " & cboDeductora.ItemData(cboDeductora.ListIndex) _
'       & " where id_solicitud = " & Operacion.Operacion
'Call ConectionExecute(strSQL)
'
'vDocumentoFormalizacion = False
'vPasaFormalizacion = True
'
''Se actualiza por Codigo Optimizado el 2010/05/13 por linea siguiente
'curDesembolsos = fxMontoEnGeneral(Operacion.Operacion)
'
'
''Calculo de Intereses de Formalizacion
'lngPriDeduc = txtAno.Text & Format(fxConvierteMES(cboMes.Text), "00") & "." & cboFrecuencia.ItemData(cboFrecuencia.ListIndex)
'
'vFechaCalculo = fxFechaCalculo(Operacion.Codigo, lngPriDeduc, vDiaPago)
'curIntDias = 0
'
'If fxCobraTasaFormaliza(cboDestino.ItemData(cboDestino.ListIndex)) Then
'  curIntDias = fxInteresesHastaFormalizar(, , lngPriDeduc, vDiaPago)  'Ojo Con los Convenios
'End If
'
'
'
''Abonar Intereses + Primer Cuota
'strSQL = "select R.Id_Solicitud,R.Codigo,R.FechaForp,R.Montoapr,R.Int,R.InteresV,R.Plazo" _
'       & ",R.Cuota,R.Primer_Cuota,R.cuota_poliza,R.fecha_inicio_calculo" _
'       & ",Dst.int_form, ISNULL(Dst.TCIntForma, 'A') as 'Tipo_Cobro'" _
'       & " from reg_creditos R" _
'       & " left join Catalogo_Destinos Dst on R.cod_destino = Dst.cod_Destino" _
'       & " where R.id_solicitud =" & Operacion.Operacion
'Call OpenRecordSet(rs, strSQL)
'
'
'On Error GoTo vError
'
'curPSD = IIf(IsNull(rs!cuota_poliza), 0, rs!cuota_poliza)
'vTipoCobro = rs!Tipo_Cobro
'
''Si la Primer Cuota esta Marcada, entonces :
'' 1. Si la Fecha de Formalizacion es el 15 o antes, solo se le cobra la cuota y
''    se le eliminan los intereses de formalización, quedando solo los de la cuota
''    y la fecha de primer deduccion es el mes actual + 1
'' 2. Si es despues del 15, se le cobran los intereses de formalización del dia hasta
''    el ultimo dia del mes en proceso + la primer cuota y la fecha de la primer
''    deducción es el mes de proceso + 2
'' 3. Nota: La fecha de Calculo para el caso 1, tiene que ser un dia menor a la fecha
''    de formalizacion (para Reflejar efecto en la Boleta)
''                                               [*** Modificación al 2002/07/01 ***]
'
''If GLOBALES.SysPlanPagos = 0 Then
'        If rs!PRIMER_CUOTA = "S" Then
'
'          vPrimerCuota = True
'
''          lngPriDeduc = fxPrimerDeduccionCuota(, rs!fecha_inicio_calculo)
'
'          curInteres = rs!montoapr * rs!interesv / 1200
'          curAmortiza = rs!Cuota - curInteres
'          curCuota = rs!Cuota
'
''          'Fecha de Ultima Deducción
''          iMes = Month(rs!fecha_inicio_calculo)
''          lngAnio = Year(rs!fecha_inicio_calculo)
''
''             If iMes = 12 Then
''                iMes = 1
''                lngAnio = lngAnio + 1
''             Else
''                iMes = iMes + 1
''             End If
''
''             'Calcular Intereses Hasta el Ultimo día del Mes
''             vFechaCalculo = CDate(lngAnio & "/" & Format(iMes, "00") & "/01")
''             vFechaCalculo = DateAdd("d", -1, vFechaCalculo)
''
''             If curIntDias > 0 Then 'Esto porque los convenios no cobran intereses
''               curIntDias = ((rs!interesv / 36000) * rs!montoapr * (Abs(DateDiff("d", vFechaCalculo, rs!fecha_inicio_calculo)) + 1))
''             End If
''             curIntDias = curIntDias
''
''          FechaUltima = CLng((lngAnio & Format(iMes, "00")))
''          'Fin del Calculo de la Ultima Deducción
'
'        Else
'
'          FechaUltima = fxFechaProcesoAnterior(lngPriDeduc)
'
'        End If
''Else
''    FechaUltima = fxFechaProcesoAnterior(lngPriDeduc)
''
''End If 'Plan de Pagos  (Primer Cuota..: Deduccion)
'
'
''Verificar Que el Monto Aprobado de la operacion sea mayor a las deducciones que
''Se le van ha aplicar.
'
'curTotal = rs!montoapr - (curInteres + curAmortiza + curIntDias + curRefunde + curRetencion + curDesembolsos + curPSD + curCargos)
'If curTotal <> 0 Then
'   If Abs(curTotal) < 1 Then
'       curIntDias = curIntDias + curTotal
'       curTotal = 0
'   End If
'End If
'rs.Close
'
'If curTotal < 0 Then
'    If Abs(curTotal) > 1 Then
'       vPasaFormalizacion = False
'       Me.MousePointer = vbDefault
'       MsgBox " - No se puede formalizar esta operación porque el Monto de los Rebajos es Mayor al Monto Aprobado...", vbCritical
'       Exit Sub
'    End If
'End If
'
'
''Inicia Transacciones
'
'glogon.Conection.BeginTrans
'vTransac = True
'
'
''Verifica si el monto a girar es Cero, para el cual se debe Generar Forzadamente la ND
'If (curTotal < 0 And curTotal > -0.009) Or curTotal = 0 Then
'  vDocumentoFormalizacion = True
'End If
'
'
''Rebajar PSD Tambien al monto a girar
'strSQL = "update reg_creditos set pagare = " & txtPagare _
'       & ",documento_referido='" & Mid(txtDocumento, 1, 18) & "', Estadosol = 'F',Estado ='A'" _
'       & ",prideduc =" & lngPriDeduc & ",monto_girado = " _
'       & Operacion.MontoAprobado - (curRefunde + curRetencion + curDesembolsos + curInteres + curIntDias + curAmortiza + curPSD + curCargos) _
'       & ",fecult = " & FechaUltima _
'       & ",fecha_calculo_int = '" & Format(vFechaCalculo, "yyyy/mm/dd") & "'" _
'       & ",userfor = '" & glogon.Usuario & "'" _
'       & ",saldo_mes = montoapr, saldo = montoapr, interesc = " & curIntDias
'           If (chkEnviarATesoreria.Value = vbUnchecked) Then
'    strSQL = strSQL & ",tesoreria='" & Format(vFecha, "yyyy/mm/dd") & "'"
'    vDocumentoFormalizacion = True
'
'  End If 'CON SOLO QUE NO SE ENVIA A TESORERIA HAY QUE CREAR NOTA DE DEBITO
'
'  If (chkEnviarATesoreria.Value = vbChecked) And vDocumentoFormalizacion And fxMontoEnDesembolsos(Operacion.Operacion) = 0 Then
'    strSQL = strSQL & ",tesoreria = '" & Format(vFecha, "yyyy/mm/dd") & "'"
'    vDocumentoFormalizacion = True
'  End If 'ES ND Y NO NECESITA ENVIARSE A TESORERIA, PORQUE NO TIENE DESEMBOLSOS
'         'DE LO CONTRATIO SE GENERA DE LA ND Y SE TRASLADAN LOS DESEMBOLSOS A TESORERIA
'         'EN OTRO PROCESO.
'
'  strSQL = strSQL & " where id_solicitud = " & Operacion.Operacion
'  Call ConectionExecute(strSQL)
'
'
'
''Inserta Registro de la Formalización
'
'vDias = DateDiff("d", dtpDesembolso.Value, vFechaCalculo) + 1
'
'If GLOBALES.SysPlanPagos = 1 Then
'      If vTipoCobro = "A" Then
'            strSQL = "exec spCrdPlanPagoAbonoEC " & Operacion.Operacion & ",'CRD000','" & glogon.Usuario & "','FRM','" & Operacion.Operacion _
'                & "'," & vDias & "," & curIntDias & ",0,0,'" & Format(DateAdd("d", vDias, vFecha), "yyyy/mm/dd") & "','',0,0"
'      Else
'            strSQL = "exec spCrdPlanPagoAbonoEC " & Operacion.Operacion & ",'CRD000','" & glogon.Usuario & "','FRM','" & Operacion.Operacion _
'                & "',0," & curIntDias & ",0,0,'" & Format(vFecha, "yyyy/mm/dd") & "','',0,0"
'      End If
'Else
'      'No se incluye PSD porque será referenciada en otro procedimiento
'      strSQL = "insert creditos_dt(CODIGO,ID_SOLICITUD,CUOTA,ABONO,INTCP,AMORTIZA,FECHAS," _
'             & "FECHAP,TCON,NCON,ESTADO,ESTADO_ASIENTO,cod_Concepto,usuario,cod_Caja) values('" & Operacion.Codigo & "'," _
'             & Operacion.Operacion & "," & curCuota & "," & curInteres + curIntDias + curAmortiza _
'             & "," & curInteres + curIntDias & "," & curAmortiza & ",dbo.MyGetdate()" _
'             & "," & GLOBALES.glngFechaCR & ",'FRM'," & Operacion.Operacion & ",'A','G','CRD000','" & glogon.Usuario & "','')"
'End If
'Call ConectionExecute(strSQL)
'
'
''Aplica Abono a Primer Cuota
'If vPrimerCuota And curCuota > 0 Then
'    strSQL = "exec spCrdPlanPagoAbonoEspecial " & Operacion.Operacion & ",'CRD001','" & glogon.Usuario & "','FRM" _
'           & "','" & Operacion.Operacion & "',0," & curInteres & ",0," & curAmortiza _
'           & ",0,0,'" & Format(vFecha, "yyyy/mm/dd hh:mm:ss") & "',''" _
'           & "," & lngPriDeduc & ", 1 ,1,1,1"
''    Call ConectionExecute(strSQL)
'
'
''    strSQL = strSQL & Space(10) & "update reg_creditos set saldo = montoapr - " & curAmortiza _
''           & ", amortiza = isnull(amortiza,0) + " & curAmortiza _
''           & ", interesc = isnull(interesc,0) + " & curInteres _
''           & " where id_solicitud = " & Operacion.Operacion
'      Call ConectionExecute(strSQL)
'
'End If
'
'
''Cambia Los Procesos Anteriores x StoreProcedures
'strSQL = "exec spCRDFormalizaDetalle " & Operacion.Operacion & ",'" & Format(vFecha, "yyyy/mm/dd") & "'," & GLOBALES.glngFechaCR & ",'" & glogon.Usuario & "'"
'Call ConectionExecute(strSQL)
'
''Cierra Transacciones
'glogon.Conection.CommitTrans
'vTransac = False
'
'
''Crea Plan de Pagos
'If GLOBALES.SysPlanPagos = 1 Then
'    strSQL = "exec spCrdPlanPagos " & Operacion.Operacion
'    Call ConectionExecute(strSQL)
'End If
'
'
''Envío a Tesorería
'strSQL = "exec spCrdCreditoEnviaTesoreria_Main " & Operacion.Operacion & ",'" & Operacion.Documento & "'"
'Call OpenRecordSet(rs, strSQL)
'  If rs!BoletaCheques = 1 Then
'    vBoletaCK = True
'  Else
'    vBoletaCK = False
'  End If
'rs.Close




'spCrd_SGT_Formalizacion(@Operacion int, @DeducePlanilla smallint , @Deductora int , @PriDeduc dec(18,2)
'            , @fDesembolso datetime, @Recurso varchar(10), @Oficina varchar(10), @TasaFacial dec(8,4)
'            , @Usuario varchar(30)

'select @Operacion 'Operacion',  @ErrorMsj as 'ErrorMsj'
'    , @DocumentoFormalizacion as 'DocumentoFormalizacion', @PasaFormalizacion as 'PasaFormalizacion'
'    , @BoletaCk as 'BoletaCK'
    
Me.MousePointer = vbHourglass

'Calculo de Intereses de Formalizacion
lngPriDeduc = txtAno.Text & Format(fxConvierteMES(cboMes.Text), "00") & "." & cboFrecuencia.ItemData(cboFrecuencia.ListIndex)

strSQL = "exec spCrd_SGT_Formalizacion " & Operacion.Operacion & ", " & chkDeducPlanilla.Value & ", " & cboDeductora.ItemData(cboDeductora.ListIndex) _
       & ", " & lngPriDeduc & ", '" & Format(dtpDesembolso.Value, "yyyy-mm-dd") & "', '" & cboRecursos.ItemData(cboRecursos.ListIndex) _
       & "', '" & GLOBALES.gOficinaTitular & "', " & CCur(txtTasaFacial.Text) & ", '" & Operacion.Documento & "', '" & glogon.Usuario _
       & "', " & chkEnviarATesoreria.Value & ", " & txtPagare.Text & ", '" & Mid(txtDocumento.Text, 1, 18) & "'"
Call OpenRecordSet(rs, strSQL)

 If rs!PasaFormalizacion = 0 Then
    Me.MousePointer = vbDefault
    MsgBox rs!ErrorMsj, vbExclamation
 End If

 If rs!PasaFormalizacion = 1 Then
    'BITACORA
    Call Bitacora("Registra", "Formalización de la OP: " & Operacion.Operacion)
    
    'Imprime Boleta de Formalizacion
    Call sbCrdSGTBoletaFormaliza(Operacion.Operacion)
 
 
    If rs!BoletaCK = 1 Then
       Call sbCrdSGTBoletaCK(Operacion.Operacion)
    End If
 
    Me.MousePointer = vbDefault
    MsgBox "Formalización Aplicada Satisfactoriamente...", vbInformation
 
 End If

''Tags de Seguimiento
'Call sbCrdOperacionTags(Operacion.Operacion, Operacion.Codigo, "S03", "", txtObservaciones.Text)


Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub sbAnular()
Dim strSQL As String, rs As New ADODB.Recordset
Dim vMensaje As String


On Error GoTo vError

Me.MousePointer = vbHourglass

vMensaje = ""

strSQL = "exec spCRDFormalizaAnulacion " & Operacion.Operacion & ",'" & glogon.Usuario & "',1"
Call OpenRecordSet(rs, strSQL)
If Not glogon.error Then
    vMensaje = rs!Mensaje
    rs.Close

    'BITACORA
    Call Bitacora("Registra", "Anulación de la OP: " & Operacion.Operacion)
    Call sbBitacoraCredito("13", "Monto : " & txtMonto.Text, "C", txtOperacion.Text, txtCodigo.Text, "SGT Anula Formalizacion del Día")
    ''Tags de Seguimiento (Se Aplica desde el Procedure.)
    'Call sbCrdOperacionTags(Operacion.Operacion, Operacion.Codigo, "S09", "", "SGT Anula Formalizacion del Día")
    
    Call sbTrazabilidad_Inserta("09", txtOperacion.Text, txtOperacion.Text)
    vMensaje = vMensaje & vbCrLf & "...Anulación Realizada Satisfactoriamente..."
    
    Me.MousePointer = vbDefault
    If GLOBALES.SysDocVersion = 2 Then
        Call sbImprimeRecibo(Operacion.Operacion, "AFR")
    End If
    
    If Len(vMensaje) > 0 Then MsgBox vMensaje, vbInformation



End If 'Aplica sin Error

Me.MousePointer = vbDefault

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub



Private Sub chkPrimera_Click()
Dim strSQL As String


If lsw.Enabled = True And Operacion.Operacion > 0 And Not vOperacionLoad Then
   strSQL = "update reg_creditos set PRIMER_CUOTA = '" & IIf((chkPrimera.Value = vbChecked), "S", "N") & "'" _
          & " where id_solicitud = " & Operacion.Operacion
   
   Call ConectionExecute(strSQL)
End If

End Sub

Private Sub chkTrasladoSalario_Click()
Dim strSQL As String


If lsw.Enabled = True And Operacion.Operacion > 0 And Not vOperacionLoad Then
   strSQL = "update reg_creditos set IND_APLICA_TRASLADO_SALARIO = " & chkTrasladoSalario.Value _
          & " where id_solicitud = " & Operacion.Operacion
   
   Call ConectionExecute(strSQL)
End If

End Sub

Private Sub cmdAplicarFormalizacion_Click()
Dim strSQL As String, rs As New ADODB.Recordset
Dim i As Integer


If Me.optFormalizacion(0).Value Then
 
    strSQL = "select isnull(count(*),0) as Total from operacion_requisitos where id_solicitud =" & Operacion.Operacion & " and Estado = 0"
    Call OpenRecordSet(rs, strSQL)
    If rs!Total > 0 Then
     Operacion.Ventana = "R"
     Call sbFormsCall("frmCR_SeguimientoReqCar", 1, , , False, Me)
    End If
    rs.Close
 
 If fxVerificaFormalizacion Then
    
     i = MsgBox("Esta seguro que desea >> formalizar << esta Operación", vbYesNo)
     If i = vbYes Then
         Call sbFormalizar
     End If
     
 Else 'Falla Verificacion de Formalizacion
  MsgBox vMensaje, vbCritical
 End If

Else 'Anulacion de la formalizacion
  If fxVerificaAnulacion Then
     i = MsgBox("Esta seguro que desea >> Anular << esta Operación", vbYesNo)
     If i = vbYes Then
        Call sbAnular
     End If
    
  Else
    MsgBox vMensaje, vbCritical
  End If
End If

Call sbCargaOperacion

End Sub




Private Sub dtpFechaFormalizacion_Change()
Dim strSQL As String

strSQL = "update reg_creditos set fechaforp = '" & Format(dtpFechaFormalizacion.Value, "yyyy/mm/dd") & "'" _
       & " where id_solicitud = " & Operacion.Operacion
    
Call ConectionExecute(strSQL)

Call Bitacora("Modifica", "Fecha Formalizacion Operacion " & Operacion.Operacion)

End Sub

Private Sub dtpFechaSolicitud_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = vbKeyReturn Then txtMonto.SetFocus
End Sub


Private Function fxVerificaExisteCodigo(strCodigo As String) As Boolean
Dim rsX As New ADODB.Recordset, strSQL As String
strSQL = "select isnull(count(*),0) as Existe from catalogo where codigo ='" & strCodigo & "'"
Call OpenRecordSet(rsX, strSQL, 0)
fxVerificaExisteCodigo = IIf((rsX!Existe > 0), True, False)
rsX.Close
End Function

Private Function fxVerificaExisteRangoCodigo(strCodigo As String, curMonto As Currency) As Boolean
Dim rsX As New ADODB.Recordset, strSQL As String
 
strSQL = "select isnull(count(*),0) as Existe from rangos"
strSQL = strSQL & " where codigo ='" & strCodigo & "' and " & curMonto & " >=de and " _
        & curMonto & " <=  hasta"
Call OpenRecordSet(rsX, strSQL, 0)
fxVerificaExisteRangoCodigo = IIf((rsX!Existe > 0), True, False)
rsX.Close
End Function

Private Function fxVerificaFiadores() As Boolean
Dim rsX As New ADODB.Recordset, strSQL As String

fxVerificaFiadores = True

vMensaje = ""

strSQL = "select isnull(count(*),0) as Existe from fiadores where estado ='A' and id_solicitud=" & Operacion.Operacion
rsX.CursorLocation = adUseServer
Call OpenRecordSet(rsX, strSQL, 0)
If rsX!Existe = 0 Then vMensaje = "- La garantía de esta operacion es fiduciaria y no se ha especificado ningún fiador"
rsX.Close

If Len(vMensaje) > 0 Then fxVerificaFiadores = False

End Function

Private Function fxVerificaRecepcion() As Boolean
Dim rsX As New ADODB.Recordset, strSQL As String, vPermiteCbrJud As Boolean
Dim Porcentaje As Currency, vEmiteTipo As String

fxVerificaRecepcion = True
vMensaje = ""
vPermiteCbrJud = False


vEmiteTipo = fxTipoDocumento(cboTipoDocumento.Text)
 

If vEmiteTipo = "CP" Then
    If Not IsNumeric(txtProveedorId.Text) Then
        vMensaje = vMensaje & vbCrLf & "- No se ha indicado a ningún Proveedor para la Cuenta por Pagar"
    End If
End If


If dtpVence.Visible Then
  If DateDiff("d", vFechaSistema, dtpVence.Value) <= 0 Then
              vMensaje = vMensaje & vbCrLf & "La fecha de Vencimiento no puede ser igual o menor a la actual"
  End If
End If

If Operacion.Operacion = 0 Or fxEstadoOperacion(cboEstado.Text) = "R" Then
        
        If IsNumeric(txtPlazo) Then
         If txtPlazo < 1 Then vMensaje = vMensaje & vbCrLf & "- El Plazo Solicitado NO es válido"
        Else
           vMensaje = vMensaje & vbCrLf & "- El Plazo Solicitado es Inválido"
        End If
        
        If IsNumeric(txtTasa) Then
         If txtTasa < 0 Then vMensaje = vMensaje & vbCrLf & "- La Tasa solicitada no es válida"
        Else
           vMensaje = vMensaje & vbCrLf & "- El Interés Solicitado es Inválido"
        End If
        
        If IsNumeric(txtMonto.Text) Then
         If txtMonto.Text < 1 Then vMensaje = vMensaje & vbCrLf & "- El Monto Solicitado NO es válido"
        Else
           vMensaje = vMensaje & vbCrLf & "- El Monto Solicitado es Inválido"
        End If
        
        'Verifica Rangos
        If Len(vMensaje) = 0 Then
          strSQL = "exec spCrdFormaliza_Valida_Rangos '" & txtCedula.Text & "','" & txtCodigo.Text & "'," _
                 & CCur(txtMonto.Text) & "," & CCur(txtTasa) & "," & CInt(txtPlazo.Text) _
                 & ",'" & cboDestino.ItemData(cboDestino.ListIndex) & "','" & cboGarantia.ItemData(cboGarantia.ListIndex) _
                 & "'," & Operacion.Operacion
          Call OpenRecordSet(rsX, strSQL)
          If Len(rsX!Mensaje) > 0 Then
              vMensaje = vMensaje & vbCrLf & rsX!Mensaje
          End If
          rsX.Close
        End If
        
        
        'Revision de la Garantia en Fondos / Planes
        If cboGarantia.ItemData(cboGarantia.ListIndex) = "Y" Then
         If cboFondo.ListCount <= 0 Then
                    vMensaje = vMensaje & vbCrLf & " - No existe un PLAN especificado para cobertura de esta garantía"
         Else
         
                If cboFondoContrato.ListCount <= 0 Then
                     If CCur(txtMonto.Text) > fxDisponibleFondos(txtCedula, cboFondo.ItemData(cboFondo.ListIndex), 0) Then
                          vMensaje = vMensaje & vbCrLf & " - El Monto Solicitado excede la cobertura de sus PLANES de ahorros..."
                     End If
                Else
                     If CCur(txtMonto.Text) > fxDisponibleFondos(txtCedula, cboFondo.ItemData(cboFondo.ListIndex), cboFondoContrato.ItemData(cboFondoContrato.ListIndex)) Then
                          vMensaje = vMensaje & vbCrLf & " - El Monto Solicitado excede la cobertura de sus PLAN DE INVERSION..."
                     End If
                End If
         
         End If
        End If


End If 'Recepcion




'Verifica que el Banco de Deposito exista o este asignado (autorizado para el usuario)
If fxEstadoOperacion(cboEstado.Text) = "P" Or fxEstadoOperacion(cboEstado.Text) = "R" Then
    If Not fxBancoAsignado(cboBanco.ItemData(cboBanco.ListIndex), glogon.Usuario) Then
       vMensaje = vMensaje & vbCrLf & "- EL BANCO INDICADO NO SE ENCUENTRA AUTORIZADO AL USUARIO : " & glogon.Usuario
    End If
End If

'VERIFICAR SI TIENE CODIFICACION CONTABLE
'Update_2017m02d22: Simplifica los datos de respuesta y agrega la validacion del Nueva operacion para personas en Cobro Judicial

strSQL = "select ctaNintC,retencion,Poliza, activo, isnull(Permite_PersonaEnCbrJud,0) as 'Permite_Cbr' " _
       & " from catalogo where codigo ='" & txtCodigo & "'"
Call OpenRecordSet(rsX, strSQL, 0)
 If rsX.EOF And rsX.BOF Then
   vMensaje = vMensaje & vbCrLf & "- El código de préstamo no existe"
 Else
 
  If rsX!Permite_Cbr = 1 Then
      vPermiteCbrJud = True
  End If
  
  'Verifica si el codigo tiene codificacion contable
  'Es suficiente con evaluar cualquiera de las 9, pues el sistema
  'solo permite actualizar cuando se especifican todas.
   If IsNull(rsX!ctaNintC) Then vMensaje = vMensaje & vbCrLf & "- El código no se encuentra codificado contablemente"
   
   'No se permiten retenciones ni polizas
   If rsX!retencion = "S" Or rsX!Poliza = "S" Then vMensaje = vMensaje & vbCrLf & "- No se permite guardar porque el código pertenece a una Retencion o Poliza"
  
   'Verificar que el Codigo se encuentre Activo
   If rsX!Activo = 0 Then vMensaje = vMensaje & vbCrLf & "- La Línea de Crédito no se encuentra Activa..."
  
 End If
rsX.Close

'Verifica el estado de la persona vrs los estados permitidos en esta línea de crédito
'Update_2017m02d22: Mejora Consulta
strSQL = "select isnull(count(*),0) as 'Existe'" _
       & " from CRD_CATALOGO_ESTADOS E inner join Socios S on E.cod_Estado = S.EstadoActual and S.cedula = '" & txtCedula.Text _
       & "' where codigo = '" & txtCodigo.Text & "'"
Call OpenRecordSet(rsX, strSQL, 0)
If rsX!Existe = 0 Then vMensaje = vMensaje & vbCrLf & "- Esta Línea de Crédito no Admite El estado actual de la persona (verifique.!)"
rsX.Close

'VERIFICAR COMBOS
If fxCodigoDestino(cboDestino.ItemData(cboDestino.ListIndex), txtCodigo) = 0 Then vMensaje = vMensaje & vbCrLf & "- El Destino No es válido para Esta Línea"
If cboBanco.Text = "" Or cboBanco.ListCount = 0 Then vMensaje = vMensaje & vbCrLf & "- El Banco Especificado NO EXISTE"
If fxCodigoComite(cboComite.Text) = 0 Then vMensaje = vMensaje & vbCrLf & "- El Comité Especificado NO EXISTE"
If fxEstadoOperacion(cboEstado.Text) = "" Then vMensaje = vMensaje & vbCrLf & "- El Estado de la Operación NO ES VALIDO"
If fxTipoDocumento(cboTipoDocumento.Text) = "" Then vMensaje = vMensaje & vbCrLf & "- La emisión de la operación NO ES VALIDA"
If cboGarantia.ItemData(cboGarantia.ListIndex) = "" Then vMensaje = vMensaje & vbCrLf & "- La Garantía especificada NO ES VALIDA"


'Verificar que la persona no tenga prestamos en Cobro Judicial Activos
If Not vPermiteCbrJud Then
    strSQL = "select isnull(count(*),0) as Existe from reg_creditos" _
           & " where estado = 'A' and proceso = 'J' and cedula = '" & txtCedula & "'"
    Call OpenRecordSet(rsX, strSQL, 0)
    If rsX!Existe > 0 Then vMensaje = vMensaje & vbCrLf & "- La persona tiene créditos en Cobro Judicial"
    rsX.Close
End If

'Verificar que el estado del credito
'2019-08-27 Desactivado
'If Operacion.Operacion > 0 Then
'    strSQL = "select isnull(estado,'T') as Estado from reg_creditos" _
'           & " where id_solicitud = " & Operacion.Operacion
'    Call OpenRecordSet(rsX, strSQL, 0)
'    If rsX!Estado <> "T" Then vMensaje = vMensaje & vbCrLf & "- Esta operación ya esta activa, no se pueden guardar los cambios.!"
'    rsX.Close
'End If

If Len(vMensaje) > 0 Then fxVerificaRecepcion = False

End Function

Private Function fxVerificaNivel()
Dim rsX As New ADODB.Recordset, rsX2 As New ADODB.Recordset, strSQL As String

strSQL = "select count(*) as Existe from nivel_miembros A, nivel_derechos B where A.nv_cod_grupo = " _
       & "B.nv_cod_grupo and nombre = '" & glogon.Usuario & "' and codigo = '" _
       & txtCodigo & "'"
Call OpenRecordSet(rsX, strSQL, 0)
rsX.Close

End Function

Private Function fxVerificaFormalizacion() As Boolean
Dim rsX As New ADODB.Recordset, strSQL As String
Dim lngPriDeduc As Currency, vFecha As Date
Dim Porcentaje As Double, vMontoRefunde As Currency
Dim curDisponible As Currency, curGiros As Currency
Dim curMontoTmp As Currency, vPriDeducCorte As Currency

vMensaje = ""
fxVerificaFormalizacion = True


vFecha = fxFechaServidor
If chkDeducPlanilla.Value = vbChecked Then
        strSQL = "select MAX(proceso) as 'Proceso' From PRM_BITACORA" _
               & " where COD_INSTITUCION = " & cboDeductora.ItemData(cboDeductora.ListIndex) _
               & "  and GESTION = 'E' and TRANSACCION = '02'"
        Call OpenRecordSet(rsX, strSQL, 0)
        If IsNull(rsX!Proceso) Then
           vPriDeducCorte = GLOBALES.glngFechaCR
        Else
           vPriDeducCorte = rsX!Proceso
        End If
        rsX.Close
Else
           vPriDeducCorte = GLOBALES.glngFechaCR
End If


If dtpVence.Visible Then
  If DateDiff("d", vFechaSistema, dtpVence.Value) <= 0 Then
              vMensaje = vMensaje & vbCrLf & "La fecha de Vencimiento no puede ser igual o menor a la actual"
  End If
End If

If DateDiff("d", Format(dtpFechaFormalizacion.Value, "yyyy/mm/dd"), Format(dtpDesembolso.Value, "yyyy/mm/dd")) < 0 Then vMensaje = vMensaje & vbCrLf & "- La fecha del desembolsos no puede ser menor que la fecha de formalizacion"
If Not IsNumeric(txtTasaFacial.Text) Then
    vMensaje = vMensaje & vbCrLf & "- Tasa Facial no es correcta!"
End If

'Verifica que si la salida es por transferencia, la cuenta de ahorros no este en blanco
If fxTipoDocumento(cboTipoDocumento.Text) = "TE" Then
  If cboCuenta.ListCount = 0 Or cboCuenta.Text = "" Then
    vMensaje = vMensaje & vbCrLf & "- No se ha especificado una cuenta de ahorros para realizarle la transferencia electrónica..."
  End If
End If

If (Operacion.Estado = "A" Or Operacion.Estado = "C") And Me.optFormalizacion(0).Value = True _
    Then vMensaje = vMensaje & vbCrLf & "- Esta Operación ya fue procesada"

If IsNumeric(txtPagare) Then
  If txtPagare < 0 Then vMensaje = vMensaje & vbCrLf & "- # de Pagaré no es válido"
Else
  vMensaje = vMensaje & vbCrLf & "- # de Pagaré no es válido"
End If

If IsNumeric(txtAno) Then
  If txtAno < Year(vFecha) Then vMensaje = vMensaje & vbCrLf & "- El Año especificado no es válido"
Else
  vMensaje = vMensaje & vbCrLf & "- El Año para la primer deduccion no es válido"
End If

If fxConvierteMES(cboMes.Text) = cboMes.Text Then vMensaje = vMensaje & vbCrLf & "- El Mes para la primer deduccion no es válido"

lngPriDeduc = txtAno.Text & Format(fxConvierteMES(cboMes.Text), "00") & "." & cboFrecuencia.ItemData(cboFrecuencia.ListIndex)

If chkDeducPlanilla.Value = vbChecked Then
    If lngPriDeduc <= vPriDeducCorte Then
        
        If lngPriDeduc = vPriDeducCorte And chkPrimera.Value = xtpChecked Then
            ''Nada
            ' vMensaje = vMensaje & vbCrLf & "- La primer deducción no es válida porque menor a la fecha de proceso actual"
        Else
            vMensaje = vMensaje & vbCrLf & "- La primer deducción no es válida porque es igual o menor a la fecha de proceso actual"
        End If
        
    End If
Else
    If lngPriDeduc < vPriDeducCorte Then
        vMensaje = vMensaje & vbCrLf & "- La primer deducción no es válida porque menor a la fecha de proceso actual"
    End If
End If

If Month(dtpFechaFormalizacion.Value) <> Month(vFecha) Or Year(dtpFechaFormalizacion.Value) <> Year(vFecha) Then
 'Actualiza la fecha de formalizacion
 strSQL = "update reg_creditos set fechaforp = '" & Format(vFecha, "yyyy/mm/dd") _
        & "' where id_solicitud = " & Operacion.Operacion
 Call ConectionExecute(strSQL)
 dtpFechaFormalizacion.Value = vFecha
End If


'Verifica Rangos
If Len(vMensaje) = 0 Then
  strSQL = "exec spCrdFormaliza_Valida_Rangos '" & txtCedula.Text & "','" & txtCodigo.Text & "'," _
         & CCur(txtMonto.Text) & "," & CCur(txtTasa) & "," & CInt(txtPlazo.Text) _
         & ",'" & cboDestino.ItemData(cboDestino.ListIndex) & "','" & cboGarantia.ItemData(cboGarantia.ListIndex) _
         & "'," & Operacion.Operacion
  Call OpenRecordSet(rsX, strSQL)
  If Len(rsX!Mensaje) > 0 Then
      vMensaje = vMensaje & vbCrLf & rsX!Mensaje
  End If
  rsX.Close
End If

'
' STORE PROCEDURE - DE VERIFICACION DE FORMALIZACIONES
'

strSQL = "exec spCRDFormalizaValidacion " & Operacion.Operacion & ", '" & glogon.Usuario & "'"
Call OpenRecordSet(rsX, strSQL, 0)
    If rsX!nivel = 0 Then vMensaje = vMensaje & vbCrLf & "- No existe nivel de formalización de este usuario para la línea " & txtCodigo
    If rsX!refundicion = 0 Then vMensaje = vMensaje & vbCrLf & "- El saldo a refundir vario en la operación : " & Operacion.Operacion
    If rsX!bloqueo = 0 Then vMensaje = vMensaje & vbCrLf & "- Esta persona se encuentra bloqueda, hasta mañana se le podran formalizar operaciones..."
    If rsX!GarAhorro = 0 Then vMensaje = vMensaje & vbCrLf & " - El Monto aprobado excede el porcentaje aprobado de sus ahorros"
    If rsX!MaxOperaciones = 0 Then vMensaje = vMensaje & vbCrLf & "- No se permite sobrepasar el número máximo de operaciones en esta linea"
    If rsX!MaxLinea = 0 Then vMensaje = vMensaje & vbCrLf & "- No se permite sobrepasar el monto máximo de la línea"
    If rsX!MaxGarantia = 0 Then vMensaje = vMensaje & vbCrLf & "- No se permite sobrepasar el monto maximo de la línea x Garantía"
    If rsX!MaxGarantiaTotal = 0 Then vMensaje = vMensaje & vbCrLf & "- No se permite sobrepasar el monto máximo x Garantía"
    
    If rsX!Firmas = 0 Then vMensaje = vMensaje & vbCrLf & "- No se han registrado todas las firmas..."
    If rsX!LineaActiva = 0 Then vMensaje = vMensaje & vbCrLf & "- La línea de crédito no se encuentra Activa..."
    If rsX!DestinoActivo = 0 Then vMensaje = vMensaje & vbCrLf & "- El destino del crédito no se encuentra Activo..."
    If rsX!Cobertura = 0 Then vMensaje = vMensaje & vbCrLf & "- Cobertura de las Hipotecas es inferior al monto del crédito..."
    If rsX!Prendas = 0 Then vMensaje = vMensaje & vbCrLf & "- Cobertura de las Prendas es inferior al monto del crédito..."
    If rsX!EstadoPersona = 0 Then vMensaje = vMensaje & vbCrLf & "- Esta Línea de Crédito no Admite El estado actual de la persona (verifique.!)"
    If rsX!CongeladoCredito = 0 Then vMensaje = vMensaje & vbCrLf & "- La persona tiene un Proceso de Congelamiento de Cuentas (verifique.!)"
    If rsX!Requisitos = 0 Then vMensaje = vMensaje & vbCrLf & "- No se cumplieron los requisitos Obligatorios (verifique.!)"
    If rsX!BaseCalculo = 0 Then vMensaje = vMensaje & vbCrLf & "- No se ha establecido la Base de Calculo para Cuota Balloon!"

rsX.Close

'Verificar la posición de cada Operacion a refundir
strSQL = "exec spCrdSGTRefundicionesValida " & Operacion.Operacion
Call OpenRecordSet(rsX, strSQL, 0)
    If rsX!Cambios > 0 Then vMensaje = vMensaje & vbCrLf & "- " & rsX!Cambios & " Operación a Refundir a Cambiado su Estado ---> Actualice!"
rsX.Close


' hasta aqui el codigo en la base de datos
'

'Revision de Garantia en Excedentes
strSQL = "select ase_codigo from excedentes_parametros"
Call OpenRecordSet(rsX, strSQL, 0)
If UCase(Trim(rsX!ase_codigo)) = UCase(Trim(Operacion.Codigo)) Then
  If CCur(txtMonto.Text) > CCur(Format(fxExcedenteDisponible(Operacion.Cedula, False), "Standard")) Then
     vMensaje = vMensaje & vbCrLf & "- Este es un prestamo sobre excedentes, y el monto aprobado sobrepasa la tabla autorizada..."
  End If
End If
rsX.Close

'Revision de la Garantia en Fondos / Planes
If cboGarantia.ItemData(cboGarantia.ListIndex) = "Y" Then
 If cboFondo.ListCount <= 0 Then
            vMensaje = vMensaje & vbCrLf & " - No existe un PLAN especificado para cobertura de esta garantía"
 Else
 
        If cboFondoContrato.ListCount <= 0 Then
             strSQL = "exec spCRDGarantiaFNDCalculo '" & txtCedula & "','" & cboFondo.ItemData(cboFondo.ListIndex) & "',0"
             
             If CCur(txtMonto.Text) > fxDisponibleFondos(txtCedula, cboFondo.ItemData(cboFondo.ListIndex), 0) Then
                  vMensaje = vMensaje & vbCrLf & " - El Monto Solicitado excede la cobertura de sus PLANES de ahorros..."
             End If
        Else
             If CCur(txtMonto.Text) > fxDisponibleFondos(txtCedula, cboFondo.ItemData(cboFondo.ListIndex), cboFondoContrato.ItemData(cboFondoContrato.ListIndex)) Then
                  vMensaje = vMensaje & vbCrLf & " - El Monto Solicitado excede la cobertura de sus PLAN DE INVERSION..."
             End If
        End If
 
 End If
End If


If cboGarantia.ItemData(cboGarantia.ListIndex) = "H" Then
   'Actualiza Cargos para calculos de cargos con base en el avaluo
    Call sbCargosAdicionales(Operacion.Operacion, txtCodigo, CCur(txtMonto))
End If


'Cambia el 07/junio/2004
'Ahora depende del tipo de documento en el Banco, si es cheque en formula de libreta

If Trim(UCase(cboTipoDocumento.Text)) = "CHEQUE" And Len(vMensaje) = 0 Then
  Select Case fxgTESTipoDocExtraeDato(cboBanco.ItemData(cboBanco.ListIndex), "CK", "Comprobante")
    Case "02"
        Call sbFormsCall("frmCR_SeguimientoDoc", 1, , , False, Me)
        
        If Not Operacion.Valida Then vMensaje = vMensaje & vbCrLf & " - Se registró inconsistencias en la especificación del número del documento a desembolsar."
    Case "ER"
        vMensaje = vMensaje & vbCrLf & " - El Banco asignado no registra el tipo de documento solicitado para desembolsos"
  End Select
End If 'CK

If Operacion.EstadoSolicitud <> "R" Then
     vMensaje = vMensaje & vbCrLf & "- Esta solicitud no se encuentra recibida..."
End If



If Len(vMensaje) = 0 Then
    curDisponible = 0
    strSQL = "exec spCRDDisponibleRecurso '" & cboRecursos.ItemData(cboRecursos.ListIndex) & "','" & Format(dtpDesembolso.Value, "yyyy/mm/dd") & "'"
    Call OpenRecordSet(rsX, strSQL, 0)
    If Not rsX.EOF And Not rsX.BOF Then
        curDisponible = IIf(IsNull(rsX!Disponible), 0, rsX!Disponible)
    End If
    rsX.Close

    Call sbResumenOperacion
    curGiros = CCur(lsw.ListItems.Item(10).SubItems(1)) - CCur(lsw.ListItems.Item(8).SubItems(1)) + CCur(lsw.ListItems.Item(3).SubItems(1)) - CCur(lsw.Tag)
    
    If curGiros > 0 Then
        If curDisponible < curGiros Then
           vMensaje = vMensaje & vbCrLf & " - No Hay disponible en el Recurso, para desembolsar esta Operación..."
           vMensaje = vMensaje & vbCrLf & " - Monto a Girar : " & Format(curGiros, "Standard") & " - Disponible :  " & Format(curDisponible, "Standard")
           vMensaje = vMensaje & vbCrLf & " - Monto Faltante para Girar: " & Format(curGiros - curDisponible, "Standard")
        End If
    
        'Retiros en Cajas> Validacion
        If fxTipoDocumento(cboTipoDocumento.Text) = "RC" Then
          strSQL = "select Valor from CAJAS_PARAMETROS  where cod_parametro = '15'"
          Call OpenRecordSet(rsX, strSQL)
          If IsNumeric(rsX!Valor) Then
                If rsX!Valor < curGiros Then
                    vMensaje = vMensaje & vbCrLf & "- El Monto Máximo para Retiros de Efectivos en Cajas es de " _
                           & Format(rsX!Valor, "Standard") & ", Informe a su Administrador!"
                End If
          Else
            vMensaje = vMensaje & vbCrLf & "- No se ha configurado el Monto para Retiros de Efectivos en Cajas, Informe a su Administrador!"
          End If
        End If
    
    End If
End If

If Len(vMensaje) > 0 Then fxVerificaFormalizacion = False


End Function

Private Function fxVerificaAnulacion() As Boolean
Dim rs As New ADODB.Recordset, strSQL As String
Dim lngPriDeduc As Long, vFecha As Date
Dim Porcentaje As Double, vMontoRefunde As Currency
Dim rsTmp As New ADODB.Recordset

vMensaje = ""
fxVerificaAnulacion = True


If Operacion.EstadoSolicitud <> "F" Then
  vMensaje = vMensaje & vbCrLf & "- Esta Operación no ha sido formalizada! Utilice el Estado de DENEGADA!"
  If Len(vMensaje) > 0 Then fxVerificaAnulacion = False
  Exit Function
End If

strSQL = "select isnull(count(*),0) as Existe" _
   & " from NIVEL_GRUPOS N INNER JOIN nivel_miembros A" _
   & " ON N.NV_COD_GRUPO = A.NV_COD_GRUPO INNER JOIN nivel_derechos B" _
   & " ON N.NV_COD_GRUPO = B.NV_COD_GRUPO Where A.nombre = '" & glogon.Usuario _
   & "' and B.codigo = '" & txtCodigo & "' AND N.nv_tipo = 'N'" _
   & " and (" & CCur(txtMonto.Text) & " between nv_desde and nv_hasta)"

rs.CursorLocation = adUseServer
Call OpenRecordSet(rs, strSQL)
If rs!Existe = 0 Then
  vMensaje = vMensaje & vbCrLf & "- No existe nivel de anulación de este usuario para la línea.: " & txtCodigo
End If
rs.Close


'0. Verificacion base / Solo se pueden anular las formalizaciones del día
'Cambiado el: 22/2/2018 Para que sea en el mismo mes
'strSQL = "select fechaforp,datediff(d,fechaforp,dbo.MyGetdate()) as Resultado"
strSQL = "select fechaforp,  month(fechaforp) - month(dbo.Mygetdate()) + ( year(fechaforp) - year(dbo.Mygetdate()) ) as Resultado" _
       & " from reg_creditos where id_solicitud = " & Operacion.Operacion
Call OpenRecordSet(rs, strSQL)
    vFecha = rs!FechaForp
    If Abs(rs!Resultado) > 0 Then
      vMensaje = vMensaje & vbCrLf & "- Esta operación fue formalizada en un mes diferente..."
    End If
rs.Close

'Busca y Elimina en refundiciones > Inconsistencia de Registro de Operacion = Refundicion
'Antes de continuar con la anulacion
strSQL = "delete refundiciones where (id_solicitud = id_solicitudr) and id_solicitud = " & Operacion.Operacion
Call ConectionExecute(strSQL)


'2. Verifica que no se le registren desembolsos, Se deben de anular o eliminar
strSQL = "select isnull(count(*),0) as Existe from Tes_Transacciones where op = " & Operacion.Operacion _
       & " and estado <> 'A'"
Call OpenRecordSet(rs, strSQL)
If rs!Existe > 0 Then
  vMensaje = vMensaje & vbCrLf & "- Existen solicitudes o documentos emitidos (Cheques/Transferencias) en Tesorería (Proceda a Anularlos)"
End If
rs.Close

If GLOBALES.SysPlanPagos = 0 Then
    '3. Verificar si se le han realizado movimientos a la Operacion despues de su formalizacion
    strSQL = "select isnull(count(*),0) as Existe from creditos_dt where id_solicitud = " & Operacion.Operacion _
           & " and ncon <> '" & Operacion.Operacion & "'"
    Call OpenRecordSet(rs, strSQL)
    If rs!Existe > 0 Then
      vMensaje = vMensaje & vbCrLf & "- Existen movimientos a esta operación después de su formalización"
    End If
    rs.Close
    
    'No tiene porque tener ningun registro de mora
    strSQL = "select isnull(count(*),0) as Existe from MOROSIDAD where id_solicitud = " & Operacion.Operacion
    Call OpenRecordSet(rs, strSQL)
    If rs!Existe > 0 Then
      vMensaje = vMensaje & vbCrLf & "- Existen movimientos a esta operación después de su formalización"
    End If
    rs.Close
    
    
    '3a. Verificar si se le han realizado movimientos a las refundiciones (Abonadas o Canceladas)
    strSQL = "select id_solicitud,consec from creditos_dt where tcon in('3','FRM') and ncon = '" & Operacion.Operacion & "'"
    Call OpenRecordSet(rs, strSQL)
    Do While Not rs.EOF
     strSQL = "select isnull(count(*),0) as Existe from creditos_dt where id_solicitud = " _
            & rs!ID_SOLICITUD & " and consec > " & rs!CONSEC
     Call OpenRecordSet(rsTmp, strSQL, 0)
        If rsTmp!Existe > 0 Then
          vMensaje = vMensaje & vbCrLf & "- Existen movimientos realizados a la op:" & rs!ID_SOLICITUD _
                   & " posterior a su refundicion"
        End If
     rsTmp.Close
     rs.MoveNext
    Loop
    rs.Close
    
    '3a. a la fecha de formalizacion (Doble verificacion para movimientos en mora no reflejados)
    strSQL = "select isnull(count(*),0) as Existe from creditos_dt" _
            & " where fechas > '" & Format(vFecha, "yyyy/mm/dd hh:mm:ss") & "' and id_solicitud <> " & Operacion.Operacion _
            & " and tcon not in('3','FRM') and ncon = '" & Operacion.Operacion & "'"
    Call OpenRecordSet(rs, strSQL)
    If rs!Existe > 0 Then
       vMensaje = vMensaje & vbCrLf & "- Existen movimientos realizados a refundiciones posterior a la formalizacion"
    End If
    rs.Close
    
    '3b. a la fecha de formalizacion para Morosidad
    strSQL = "select isnull(count(*),0) as Existe from morosidad" _
            & " where Estado = 'C' and fecUlt > '" & Format(vFecha, "yyyy/mm/dd hh:mm:ss") & "' and id_solicitud <> " & Operacion.Operacion _
            & " and tcon not in('3','FRM') and ncon = '" & Operacion.Operacion & "'"

    Call OpenRecordSet(rs, strSQL)
    If rs!Existe > 0 Then
       vMensaje = vMensaje & vbCrLf & "- Existen movimientos realizados a Mora de refundiciones posterior a la formalizacion"
    End If
    rs.Close

End If 'SysPlanPagos = 0

'4. No puede anular retenciones
strSQL = "select retencion from catalogo where codigo = '" & Operacion.Codigo & "'"
Call OpenRecordSet(rs, strSQL)
If rs!retencion = "S" Then
   vMensaje = vMensaje & vbCrLf & "- Este es un código de retención No se puede Anular..."
End If
rs.Close


If Len(vMensaje) > 0 Then fxVerificaAnulacion = False


End Function


Private Sub sbConsultaX(vCedula As String)
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem

On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = "Select R.id_solicitud,R.codigo,R.cedula,S.Nombre,R.fechasol,R.montosol,R.estadosol,R.estado,R.proceso" _
       & " FROM REG_CREDITOS R inner join CATALOGO C ON R.CODIGO = C.CODIGO" _
       & " inner join Socios S on R.cedula = S.cedula" _
       & " where C.retencion = 'N' and C.poliza = 'N' and R.cedula like '%" & Trim(txtConCedula.Text) _
       & "%' and S.nombre like '%" & Trim(txtConNombre.Text) & "%'" _
       & " order by R.id_solicitud desc"

lswBusca.ListItems.Clear

Call OpenRecordSet(rs, strSQL, 0)

Do While Not rs.EOF
 Set itmX = lswBusca.ListItems.Add(, , CStr(rs!ID_SOLICITUD))
  itmX.SubItems(1) = rs!Codigo
  itmX.SubItems(2) = rs!Cedula
  itmX.SubItems(3) = rs!Nombre
  
  itmX.SubItems(4) = Format(rs!FechaSol, "yyyy/mm/dd")
  itmX.SubItems(5) = Format(rs!montosol, "Standard")
  
  Select Case rs!estadosol
   Case "R"
    itmX.SubItems(6) = "Recibida"
   Case "P"
    itmX.SubItems(6) = "Pendiente"
   Case "A"
    itmX.SubItems(6) = "Aprobada"
   Case "D"
    itmX.SubItems(6) = "Denegada"
   Case "F"
    itmX.SubItems(6) = "Formalizada"
   Case "N"
    itmX.SubItems(6) = "Anulada"
  End Select
  
 Select Case rs!Estado
   Case "A"
    itmX.SubItems(7) = "Activa"
   Case "C"
    itmX.SubItems(7) = "Cancelada"
   Case Else
    itmX.SubItems(7) = "En Tramite"
 End Select
 
 Select Case rs!Proceso
   Case "J"
    itmX.SubItems(8) = "Cobro Jud"
   Case "N"
    itmX.SubItems(8) = "Normal"
   Case "T"
    itmX.SubItems(8) = "Traspaso"
   Case Else
    itmX.SubItems(8) = "------"
 End Select
 
 rs.MoveNext
Loop
rs.Close

Me.MousePointer = vbDefault
Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub



Public Sub sbGXSegTraIniTlb()
If TimerX.Interval > 0 Then
   Call TimerX_Timer
End If
 Call btnBarra_Click(0)
 txtCedula = GLOBALES.gCedulaActual
 txtCedula_LostFocus
 txtCodigo.SetFocus
End Sub

Private Sub FlatScrollBar_Change()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

If FlatScrollBar.Value = 1 And txtOperacion.Text = "" Then txtOperacion.Text = "0"
If FlatScrollBar.Value = 0 And txtOperacion.Text = "" Then txtOperacion.Text = "999999999999"

If vScroll Then
    strSQL = "select Top 1 R.id_solicitud from reg_creditos R inner join Catalogo C on R.codigo = C.codigo" _
           & " and C.Retencion = 'N' and C.poliza = 'N'"
    
    If FlatScrollBar.Value = 1 Then
       strSQL = strSQL & " where R.id_solicitud > " & txtOperacion & " order by R.id_solicitud asc"
    Else
       strSQL = strSQL & " where R.id_solicitud < " & txtOperacion & " order by R.id_solicitud desc"
    End If
    
    Call OpenRecordSet(rs, strSQL)
    If Not rs.EOF And Not rs.BOF Then
      txtOperacion = rs!ID_SOLICITUD
      Call sbCargaOperacion
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
vModulo = 3
End Sub

Private Sub Form_Load()
 
 vModulo = 3
 
 tcMain.Item(0).Visible = False
 tcMain.Item(1).Visible = False
 tcMain.Item(2).Visible = False
 tcMain.Item(3).Visible = False
 tcMain.Item(4).Visible = False

Set imgBanner.Picture = frmContenedor.imgBanner_Tramites.Picture
 
 mFrecuenciaPago = "M"

 Me.Height = 8805

 Call sbTaskPanel_Load
 
 
 vScroll = False
 FlatScrollBar.Value = 0
 vScroll = True
 
 Call sbToolBarIconos(tlbPrincipal, False)
 
 Call sbBarra_Accion("nuevo")

 
 cboCalculoAdd.AddItem "Monto del Crédito"
 cboCalculoAdd.AddItem "Monto a Girar"
 cboCalculoAdd.AddItem "Giro en Cero"
 cboCalculoAdd.Text = "Monto del Crédito"
 
 
 With cboMes
    .Clear
    .AddItem "Enero"
    .AddItem "Febrero"
    .AddItem "Marzo"
    .AddItem "Abril"
    .AddItem "Mayo"
    .AddItem "Junio"
    .AddItem "Julio"
    .AddItem "Agosto"
    .AddItem "Septiembre"
    .AddItem "Octubre"
    .AddItem "Noviembre"
    .AddItem "Diciembre"
 End With

 With tlbPrincipal.Buttons
   .Item(2).Enabled = False
   .Item(3).Enabled = False
   .Item(4).Enabled = False
 End With

With lswHistorial.ColumnHeaders
    .Clear
    .Add , , "Fecha", 1900
    .Add , , "Usuario", 1900, vbCenter
    .Add , , "Etiqueta", 2500
    .Add , , "Notas", 3500
End With

With lswBusca.ColumnHeaders
    .Clear
    .Add , , "No. Operación", 2000
    .Add , , "Línea", 1100, vbCenter
    .Add , , "Identificación", 1800
    .Add , , "Nombre", 3000
    .Add , , "Fecha", 2100, vbCenter
    .Add , , "Monto", 1800, vbRightJustify
    .Add , , "Estado", 1400, vbCenter
    .Add , , "Activa?", 1400, vbCenter
    .Add , , "Proceso", 1400, vbCenter
End With


vFechaSistema = fxFechaServidor

dtpFechaFormalizacion.Value = vFechaSistema
dtpFechaSolicitud.Value = vFechaSistema
dtpVence.Value = vFechaSistema
dtpDesembolso.Value = vFechaSistema

cboTipoDocumento.Clear
cboTipoDocumento.AddItem fxTipoDocumento("CK")
cboTipoDocumento.AddItem fxTipoDocumento("TE")
cboTipoDocumento.AddItem fxTipoDocumento("TS")
cboTipoDocumento.AddItem fxTipoDocumento("ND")
cboTipoDocumento.AddItem fxTipoDocumento("CD")
cboTipoDocumento.AddItem fxTipoDocumento("CP")
cboTipoDocumento.AddItem fxTipoDocumento("RC")
cboTipoDocumento.Text = fxTipoDocumento("TE")

Call cboTipoDocumento_Click

'Inicializa
tcMain.Item(0).Selected = True
TituloOpcion.Caption = tcMain.Item(0).Caption



Call Formularios(Me)
Call RefrescaTags(Me)

End Sub

Private Sub sbLimpiaDatos()
 Dim i As Integer


If txtCedula.Text = "" And Operacion.Operacion = 0 And tcMain.Item(0).Selected Then Exit Sub

With Operacion
 .Operacion = 0
 .Cedula = ""
 .Codigo = ""
 .EstadoSolicitud = "R"
 .Documento = ""
 .TasaPtsBono = 0
 .PlazoBono = 0
End With

 tcMain.Item(0).Selected = True
 
 cboCalculoAdd.Text = "Monto del Crédito"


 txtCedula = ""
 txtCodigo = ""
 cboCuenta.Clear
 
 txtCuota = ""
 txtCuota = ""
 txtDescripcion = ""
 txtTasa = ""
 txtNombre = ""
 lblNombre.Caption = txtNombre.Text
 txtObservaciones = ""
 txtPagare = ""
 txtPlazo = ""
 txtMonto = ""
 txtPromotorId.Text = ""
 txtPromotorNombre.Text = ""
  
 
 txtProveedorId.Text = ""
 txtProveedorNombre.Text = ""
 
imgEstado.ToolTipText = "Nueva Operación!"
Set imgEstado.Picture = imgIconosEstados.ListImages.Item(3).Picture

 imgBullet.Visible = False
 
 dtpFechaFormalizacion.Value = vFechaSistema
 dtpFechaSolicitud.Value = vFechaSistema
 dtpVence.Value = vFechaSistema
 dtpDesembolso.Value = vFechaSistema
 
 
 lblVence.Visible = False
 dtpVence.Visible = False
 
 cboEstado.Clear
 cboGarantia.Clear
 cboDestino.Clear
 
 txtRecibe.Text = ""
 txtResoluciona.Text = ""
 txtFormaliza.Text = ""
 txtTesoreria.Text = ""
 
 chkEnviarATesoreria.Value = vbChecked
 chkPrimera.Value = vbChecked
 chkDeducPlanilla.Value = vbChecked
 
 chkExpedienteDigital.Value = xtpUnchecked
 chkPagareManual.Value = xtpUnchecked
 txtFormularioId.Text = ""
 
 chkTrasladoSalario.Value = xtpUnchecked
 
 Call cboTipoDocumento_Click
 
 'Asigna Oficina
 Call sbCboAsignaDato(cboOficina, GLOBALES.gOficina, True, GLOBALES.gOficinaTitular)
 
 
' For i = 0 To btnOpciones.Count - 1
'   btnOpciones.Item(i).Enabled = False
' Next i
 
End Sub

Private Sub sbCargaCombos()
Dim strSQL As String

On Error GoTo vError

Me.MousePointer = vbHourglass

vPaso = True

strSQL = "Select id_comite as 'IdX',descripcion as 'ItmX' from comites where estado = 1"
Call sbCbo_Llena_New(cboComite, strSQL, False, True)


'Oficinas
strSQL = "Select COD_OFICINA as 'IdX',descripcion as 'ItmX' from SIF_OFICINAS where estado = 1 ORDER BY DESCRIPCION"
Call sbCbo_Llena_New(cboOficina, strSQL, False, True)


'Asigna Oficina
Call sbCboAsignaDato(cboOficina, GLOBALES.gOficina, True, GLOBALES.gOficinaTitular)

'Carga Garantias de Fondos
strSQL = "exec spCRDGarantiaFND"
Call sbCbo_Llena_New(cboFondo, strSQL, False, True)



strSQL = "select rtrim(cod_actividad) as 'IdX', rtrim(descripcion) as 'ItmX' from AFI_ACTIVIDADES_ECO where activa = 1"
Call sbCbo_Llena_New(cboActividad, strSQL, False, True)

strSQL = "select rtrim(Canal_Tipo) as 'IdX' , rtrim(descripcion) as 'ItmX' from AFI_CANALES_TIPOS where Activo = 1"
Call sbCbo_Llena_New(cboCanal, strSQL, False, True)

'Consulta todas las cuentas Bancarias
strSQL = "exec spCrd_SGT_Bancos '" & glogon.Usuario & "'"
Call sbCbo_Llena_New(cboBanco, strSQL, False, True)

vPaso = False

Me.MousePointer = vbDefault
Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub



Private Sub CargaRecursos(cbo As Object, vCodigo As String, vGrupo As String)
Dim strSQL As String

On Error GoTo vError

strSQL = " select rtrim(G.cod_grupo) as 'IdX', rtrim(G.descripcion) as 'ItmX'" _
       & " from catalogo_grupos G inner join catalogo_asignaGrp A on G.cod_grupo = A.cod_grupo" _
       & " where G.estado = 1 and A.codigo = '" & vCodigo & "'"
Call sbCbo_Llena_New(cbo, strSQL, False, True)

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub ActivaDesActiva(vEstadoSolicitud As String, vEstadoEC As String)
'Activa e Inactiva informacion, en los tabs

lsw.Enabled = True
optFormalizacion(0).Enabled = True
optFormalizacion(0).Value = True
cmdAplicarFormalizacion.Enabled = True


If vEstadoEC = "N" Then
    Select Case UCase(vEstadoSolicitud)
     Case "R"
        tcMain.Item(0).Selected = True
        
     Case "P"
        tcMain.Item(0).Selected = True
     Case "A"
        tcMain.Item(1).Selected = True
     Case "D"
        tcMain.Item(0).Selected = True
     Case "F"
       tcMain.Item(1).Selected = True
       lsw.Enabled = False
       optFormalizacion(1).Value = True
       optFormalizacion(0).Enabled = False
     Case "N"
       tcMain.Item(1).Selected = True
       cmdAplicarFormalizacion.Enabled = False
       lsw.Enabled = False
    End Select

Else
    tcMain.Item(1).Selected = True
    
    Select Case UCase(vEstadoSolicitud)
      Case "F"
            lsw.Enabled = False
            optFormalizacion(1).Value = True
            optFormalizacion(0).Enabled = False
      Case "N"
            cmdAplicarFormalizacion.Enabled = False
            lsw.Enabled = False
     End Select
End If

End Sub

Private Function fxBancoAsignado(vBanco As Integer, vUsuario As String) As Boolean
Dim strSQL As String, rs As New ADODB.Recordset

strSQL = "select isnull(count(*),0) as Existe from tes_banco_asg where id_banco = " _
       & vBanco & " and nombre = '" & vUsuario & "'"
Call OpenRecordSet(rs, strSQL)
If rs!Existe = 0 Then
  fxBancoAsignado = False
Else
  fxBancoAsignado = True
End If
rs.Close

End Function

Private Function fxOperacionDestino(vDestino As String) As String
Dim rs As New ADODB.Recordset, strSQL As String

strSQL = "select rtrim(cod_destino) + ' - ' + descripcion as ItemX from catalogo_destinos where cod_destino = '" & vDestino & "'"
Call OpenRecordSet(rs, strSQL)
If rs.EOF And rs.BOF Then
  fxOperacionDestino = " -"
Else
  fxOperacionDestino = rs!itemx
End If
rs.Close

End Function


Private Function fxGarantiaFondoContrato(pPlan As String, pContrato As Long) As String
Dim rs As New ADODB.Recordset, strSQL As String

fxGarantiaFondoContrato = ""

strSQL = "select cod_contrato,Tasa_Referencia,Aportes, isnull(Fecha_Corte, getdate()) as 'Fecha_Corte'" _
       & " from fnd_contratos " _
       & " where cod_plan = '" & pPlan & "' and cod_contrato = " & pContrato
Call OpenRecordSet(rs, strSQL)
If Not rs.EOF And Not rs.BOF Then
   fxGarantiaFondoContrato = "[Contrato: " & rs!COD_CONTRATO & "]  [Tasa: " & rs!TASA_REFERENCIA _
        & "]  [Inv: " & Format(rs!Aportes, "Standard") & "] [Vence: " & Format(rs!fecha_corte, "yyyy-MM-dd") & "]"
End If
rs.Close


End Function

Private Function fxGarantiaFondo(xGarantia As String) As String
Dim rs As New ADODB.Recordset, strSQL As String

strSQL = "exec spCRDGarantiaFNDDesc '" & xGarantia & "'"
Call OpenRecordSet(rs, strSQL)
  fxGarantiaFondo = RTrim(rs!Descripcion)
rs.Close

End Function



Private Sub sbCargaOperacion()
Dim rs As New ADODB.Recordset, strSQL As String
Dim vFecha As Date, iMes As Integer, lngAnio As Long, vProceso As Currency, pProcesoClean As Long
Dim i As Integer, vTemp As String, dFecha As Date

On Error Resume Next

' For i = 0 To 13
'   btnOpciones.Item(i).Enabled = True
' Next i

vOperacionLoad = True
tcMain.Item(0).Selected = True


strSQL = "exec spCrd_Operacion_Consulta " & txtOperacion.Text

Call OpenRecordSet(rs, strSQL)

If Not rs.EOF And Not rs.BOF Then
  
' Call sbCargaCombos
 vFecha = rs!Fecha_Server

 txtCedula.Text = rs!Cedula
 txtNombre.Text = rs!Nombre
 txtCodigo.Text = rs!Codigo
 txtDivisa.Text = rs!cod_Divisa & ""
 
 lblNombre.Caption = txtNombre.Text
 
 mFrecuenciaPago = "M"
 
 If rs!BaseCalculo = "06" Then
    mFrecuenciaPago = "Q"
 End If
 
 If rs!BulletInd = 1 Or rs!BaseCalculo = "04" Or rs!BaseCalculo = "05" Then
  imgBullet.Visible = True
  imgBullet.Enabled = True
 End If
 
 Operacion.Operacion = rs!ID_SOLICITUD
 Operacion.Cedula = rs!Cedula
 Operacion.Nombre = txtNombre
 Operacion.EstadoSolicitud = rs!estadosol
 Operacion.Codigo = rs!Codigo
 Operacion.Estado = IIf(IsNull(rs!Estado), "N", rs!Estado)
 Operacion.MontoAprobado = IIf(IsNull(rs!montoapr), 0, rs!montoapr)
 Operacion.TS = fxTsToHex(rs!TS)  'TimeStamp
 Operacion.TasaPtsBono = rs!Tasa_Pts_Bono
 
 txtTasa.ToolTipText = "Pts Bonificación: " & rs!Tasa_Pts_Bono
 
' MsgBox Operacion.TS & vbCrLf & rs!TS
 txtDescripcion.Text = rs!CodDesc
 
 txtMntNoGravable.Text = Format(rs!IVA_Monto, "Standard")
 
 
 txtCuota.Text = Format(IIf(IsNull(rs!Cuota), 0, rs!Cuota), "Standard")
 txtPlazo.Text = CStr(IIf(IsNull(rs!Plazo), 0, rs!Plazo))
 txtTasa.Text = CStr(IIf(IsNull(rs!Int), 0, rs!Int))
 
 txtPromotorId.Text = rs!ID_PROMOTOR
 txtPromotorNombre.Text = rs!Ejecutivo
 
 txtMonto.Text = Format(IIf(IsNull(rs!montosol), 0, rs!montosol), "Standard")
 
 txtObservaciones.Text = IIf(IsNull(rs!Observacion), "", rs!Observacion)
 txtPagare.Text = CStr(IIf(IsNull(rs!pagare), 0, rs!pagare))
 
 dtpFechaFormalizacion.Value = IIf(IsNull(rs!FechaForp), vFecha, rs!FechaForp)
 dtpDesembolso.Value = IIf(IsNull(rs!fecha_inicio_calculo), vFecha, rs!fecha_inicio_calculo)
 dtpFechaSolicitud.Value = IIf(IsNull(rs!FechaSol), vFecha, rs!FechaSol)
 
 dtpVence.Value = IIf(IsNull(rs!Fecha_Vence), vFecha, rs!Fecha_Vence)
 
 
 If rs!Base_Calculo = "07" Then
     lblVence.Visible = True
     dtpVence.Visible = True
 Else
     lblVence.Visible = False
     dtpVence.Visible = False
 End If
 
 Call sbCboAsignaDato(cboComite, Trim(rs!Comdesc) & "", True, rs!id_Comite)
  
 'Carga Destino
 Call sbSTCargaCboDestinos(cboDestino, Operacion.Codigo)
 Call sbCboAsignaDato(cboDestino, rs!DestinoDesc, True, rs!cod_destino & "")
 
 'Carga Oficina
 Call sbCboAsignaDato(cboOficina, rs!OficinaDesc, False, rs!OFI_PRESENTA & "")
  
  chkExpedienteDigital.Value = rs!IND_EXPEDIENTE_DIGITAL
  chkPagareManual.Value = rs!PAGARE_MANUAL
  txtFormularioId.Text = CStr(rs!Formulario)
  chkTrasladoSalario.Value = rs!IND_APLICA_TRASLADO_SALARIO
  

 'Si no tiene el banco asignado hay que crearlo pero no puede guardarlo
 'bajo este mismo banco hasta que lo tenga asignado o lo cambie.
 
 vPaso = True
 Call sbCboAsignaDato(cboBanco, rs!BancoDesc, True, IIf(IsNull(rs!cod_banco), 0, rs!cod_banco))
 vPaso = False
 
 'Carga Cuentas de la Persona
 Call cboBanco_Click
 
 'Asigna Cuenta Utilizada
 Call sbCboAsignaDato(cboCuenta, rs!CuentaDesc, True, IIf(IsNull(rs!CTA_BANCO), "", rs!CTA_BANCO))
 
 
 
 chkPrimera.Value = IIf((rs!PRIMER_CUOTA = "S"), 1, 0)
 If rs!Deduccion = 0 Then
    chkDeducPlanilla.Enabled = False
    chkDeducPlanilla.Value = vbUnchecked
 Else
    chkDeducPlanilla.Enabled = True
    chkDeducPlanilla.Value = IIf((rs!ind_deduce_planilla = "S"), 1, 0)
 
 End If
 
 '**
 txtDocumento.Text = IIf(IsNull(rs!documento_referido), "", rs!documento_referido)
 cboTipoDocumento.Text = fxTipoDocumento(IIf(IsNull(rs!emitir), "OT", rs!emitir))
  
 txtProveedorId.Text = IIf((rs!ProveedorId = 0), "", rs!ProveedorId)
 txtProveedorNombre.Text = rs!ProveedorDesc
  
 
 'Carga Deductoras por Institucion
 vPaso = True
     Call sbDeductoras_Load(rs!cod_institucion)
     Call sbCboAsignaDato(cboDeductora, rs!DeductoraDesc, True, rs!cod_Deductora)
    
     cboDeductora.Tag = CStr(rs!cod_Deductora)
 vPaso = False
 
 Call cboDeductora_Click
 
 If IsNull(rs!PriDeduc) Then
'    vProceso = fxPrimerDeduccion(rs!id_solicitud, rs!cod_Deductora)
    
    If chkPrimera.Value = vbChecked Then
        vProceso = fxPrimerDeduccionCuota(Operacion.Codigo)
    
        pProcesoClean = vProceso
        
        cboMes.Text = fxConvierteMES(Val(Mid(pProcesoClean, 5, 2)))
        txtAno.Text = Mid(pProcesoClean, 1, 4)
        
        Select Case (vProceso - pProcesoClean)
          Case 0
              cboFrecuencia.Text = "Mensual"
          Case 0.1
              cboFrecuencia.Text = "1er Quincena"
          Case 0.2
              cboFrecuencia.Text = "2da Quincena"
        End Select
    
    End If 'Primera Cuota Marcada
 Else
    vProceso = rs!PriDeduc
 
    pProcesoClean = vProceso
    
    cboMes.Text = fxConvierteMES(Val(Mid(pProcesoClean, 5, 2)))
    txtAno.Text = Mid(pProcesoClean, 1, 4)
    
    Select Case (vProceso - pProcesoClean)
      Case 0
          cboFrecuencia.Text = "Mensual"
      Case 0.1
          cboFrecuencia.Text = "1er Quincena"
      Case 0.2
          cboFrecuencia.Text = "2da Quincena"
    End Select
 
 End If
 
 Call sbSTCargaCboEstado(cboEstado, rs!estadosol)
 
 Call CargaRecursos(cboRecursos, rs!Codigo, rs!Cod_Grupo & "")
 Call sbCboAsignaDato(cboRecursos, rs!RecursoDesc, False, rs!Cod_Grupo & "")
 
 vPaso = True
        Call sbSTCargaCboGarantia(cboGarantia, rs!Codigo)
        Call sbCboAsignaDato(cboGarantia, rs!GarantiaDesc, False, rs!Garantia)
 vPaso = False
 
' If rs!Garantia = "H" Then
'     btnOpciones.Item(11).Enabled = True
'     btnOpciones.Item(12).Enabled = True
' Else
'     btnOpciones.Item(11).Enabled = False
'     btnOpciones.Item(12).Enabled = False
' End If
 
 If Not IsNull(rs!Cod_actividad) Then
    Call sbCboAsignaDato(cboActividad, rs!ActividadDesc, True, rs!Cod_actividad)
 End If
 
 If Not IsNull(rs!Canal_Tipo) Then
    Call sbCboAsignaDato(cboCanal, rs!CanalDesc, True, rs!Canal_Tipo)
 End If
 
 
 
 
 Call ActivaDesActiva(rs!estadosol, IIf(IsNull(rs!Estado), "N", rs!Estado))

 txtRecibe.Text = IIf(IsNull(rs!userRec), "", Trim(rs!userRec))
 txtResoluciona.Text = IIf(IsNull(rs!userres), "", Trim(rs!userres))
 txtFormaliza.Text = IIf(IsNull(rs!Userfor), "", Trim(rs!Userfor))
 txtTesoreria.Text = IIf(IsNull(rs!usertesoreria), "", Trim(rs!usertesoreria))
 txtAutorizada.Text = IIf(IsNull(rs!Autoriza_user), "", Trim(rs!Autoriza_user))
 txtAutorizaNota.Text = rs!Autoriza_Nota & ""

 txtFechaRec.Text = rs!FechaSol & ""
 txtFechaRes.Text = rs!fechares & ""
 txtFechaFor.Text = IIf(IsNull(rs!FECHA_REGISTRO), rs!FechaForp & "", rs!FECHA_REGISTRO)
 txtFechaTes.Text = rs!tesoreria & ""
 txtFechaAuto.Text = rs!Autoriza_Fecha & ""



  Select Case rs!estadosol
    Case "R"
       dFecha = Format(rs!FechaSol & "", "dd/mm/yyyy")
       imgEstado.ToolTipText = "Solicitado por " & rs!userRec & " - " & dFecha
       Set imgEstado.Picture = imgIconosEstados.ListImages.Item(3).Picture
       
    Case "P" 'Pendiente (HOLD)
       dFecha = Format(rs!FechaSol & "", "dd/mm/yyyy")
       imgEstado.ToolTipText = "Pendiente " & rs!userRec & " - " & dFecha
       Set imgEstado.Picture = imgIconosEstados.ListImages.Item(6).Picture

    Case "D" 'Denegado
       dFecha = Format(rs!FechaSol & "", "dd/mm/yyyy")
       imgEstado.ToolTipText = "Denegado " & rs!userRec & " - " & dFecha
       Set imgEstado.Picture = imgIconosEstados.ListImages.Item(5).Picture

    Case "F" 'Formalizado
       dFecha = Format(IIf(IsNull(rs!FECHA_REGISTRO), rs!FechaForp & "", rs!FECHA_REGISTRO), "dd/mm/yyyy")
       imgEstado.ToolTipText = "Formalizado: " & rs!Userfor & " - " & dFecha
       Set imgEstado.Picture = imgIconosEstados.ListImages.Item(1).Picture
       
    Case "N" 'Anulado
       dFecha = Format(rs!anula_fecha & "", "dd/mm/yyyy")
       imgEstado.ToolTipText = "Anulado por " & vbCrLf & rs!Anula_Usuario & " - " & dFecha
       Set imgEstado.Picture = imgIconosEstados.ListImages.Item(2).Picture
  End Select


 With tlbPrincipal.Buttons
   .Item(1).Enabled = True
   .Item(2).Enabled = True
   .Item(3).Enabled = False
   .Item(4).Enabled = False
 End With
 
 Call sbBarra_Accion("activo")

 
 Me.fraOperacion.Enabled = False


 chkEnviarATesoreria.Value = IIf(fxEnvioTesoreria(rs!cod_destino & ""), vbChecked, vbUnchecked)

    vPaso = True
    If rs!Garantia = "Y" Then
        cboFondo.Visible = True
        cboFondoContrato.Visible = True
        lblFondoDisplay.Item(0).Visible = True
        lblFondoDisplay.Item(1).Visible = True
       
       If Not IsNull(rs!GARANTIA_FND) Then
           cboFondo.Text = fxGarantiaFondo(rs!GARANTIA_FND)
           
           If IIf(IsNull(rs!Garantia_Fnd_Contrato), 0, rs!Garantia_Fnd_Contrato) > 0 Then
               vPaso = False
               Call cboFondo_Click
               vPaso = True
               Call sbCboAsignaDato(cboFondoContrato, fxGarantiaFondoContrato(rs!GARANTIA_FND, rs!Garantia_Fnd_Contrato) _
                                    , True, rs!Garantia_Fnd_Contrato)
           Else
               cboFondoContrato.Clear
           End If
       End If
    
    Else 'No es fondo
        cboFondo.Visible = False
        cboFondoContrato.Visible = False
        lblFondoDisplay.Item(0).Visible = False
        lblFondoDisplay.Item(1).Visible = False
    End If
    vPaso = False

    txtTasaFacial.Text = Format(rs!TasaFacial, "##0.00")


Else
 MsgBox "No existe esta Solicitud", vbCritical
End If
rs.Close

vOperacionLoad = False


' Codigo Corrige Apagado Inecesario del Lsw, porque por error el sistema de seguridad lo toma como de el
i = IIf(lsw.Enabled, 1, 0)
Call RefrescaTags(Me)
imgBullet.Enabled = imgBullet.Visible

lsw.Enabled = IIf((i = 1), True, False)

End Sub

Private Sub sbBusqueda(Index As Integer)
'Set GLOBALES.gfrmFormulario = Me
gBusquedas.Resultado = ""
gBusquedas.Convertir = "N"

Select Case Index
  Case 0 'txtOperacion
    gBusquedas.Convertir = "S"
    gBusquedas.Consulta = "select id_solicitud,codigo,cedula,montoapr,saldo from reg_creditos"
    gBusquedas.Orden = "id_solicitud"
    gBusquedas.Columna = "id_solicitud"
    frmBusquedas.Show vbModal
    txtOperacion = gBusquedas.Resultado
    If Len(Trim(txtOperacion)) > 0 Then
    '  Call ConsultaOperacion
    End If
  
  Case 1 'txtCedula
        gBusquedas.Consulta = "select Cedula,Nombre from socios"
        gBusquedas.Orden = "cedula"
        gBusquedas.Columna = "cedula"
        frmBusquedas.Show vbModal
       ' Call CambiaDatos
        txtCedula = gBusquedas.Resultado
        If Len(Trim(txtCedula)) > 0 Then
          txtNombre = fxNombre(Trim(txtCedula))
        End If
  
  
  Case 2 'txtCodigo
        gBusquedas.Consulta = "select codigo,descripcion from catalogo"
        gBusquedas.Orden = "codigo"
        gBusquedas.Columna = "codigo"
        gBusquedas.Filtro = " and Activo = 1 and Retencion = 'N'"
        frmBusquedas.Show vbModal
       ' Call CambiaDatos
        txtCodigo = gBusquedas.Resultado
        If Len(Trim(txtCodigo)) > 0 Then
          txtDescripcion = fxDescribeCodigo(Trim(txtCodigo))
        End If
  
  Case 3 'txtNombre
        gBusquedas.Consulta = "select Cedula,Nombre from socios"
        gBusquedas.Orden = "nombre"
        gBusquedas.Columna = "nombre"
        frmBusquedas.Show vbModal
       ' Call CambiaDatos
        txtCedula = gBusquedas.Resultado
        If Len(Trim(txtCedula)) > 0 Then
          txtNombre = fxNombre(Trim(txtCedula))
        End If
  
  
  Case 4 'txtDescripcion
        gBusquedas.Consulta = "select codigo,descripcion from catalogo"
        gBusquedas.Orden = "descripcion"
        gBusquedas.Columna = "descripcion"
        frmBusquedas.Show vbModal
       ' Call CambiaDatos
        txtCodigo = gBusquedas.Resultado
        If Len(Trim(txtCodigo)) > 0 Then
          txtDescripcion = fxDescribeCodigo(Trim(txtCodigo))
        End If
  
End Select

End Sub







Private Sub imgBullet_Click()
    Operacion.OperacionConsulta = Operacion.Operacion
    frmCR_OperacionCtaBullet.Show vbModal
End Sub

Private Sub imgConsulta_Click()

tcMain.Item(3).Selected = True

lswBusca.ListItems.Clear
txtConCedula.Text = ""
txtConNombre.Text = ""
    

End Sub

Private Sub imgGuardaFecDesembolso_Click()
Dim strSQL As String

On Error GoTo vError

If Operacion.EstadoSolicitud <> "F" Then

    strSQL = "update reg_creditos set fecha_inicio_calculo = '" & Format(dtpDesembolso.Value, "yyyy/mm/dd") & "'" _
           & " where id_solicitud = " & Operacion.Operacion & " and estadosol not in('F','N')"
    Call ConectionExecute(strSQL)
    
    Call Bitacora("Modifica", "Fecha de Desembolso Operacion " & Operacion.Operacion)
    
    MsgBox "Fecha de Desembolso Actualizada satisfactoriamente...", vbInformation

End If

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub


Private Function fxFirma(xOperacion As Long) As String
Dim strSQL As String, rs As New ADODB.Recordset

strSQL = "select firma_deudor from reg_creditos where id_solicitud = " & xOperacion
Call OpenRecordSet(rs, strSQL)
If rs.EOF And rs.BOF Then
  fxFirma = "NO"
Else
  If rs!firma_deudor = 1 Then
      fxFirma = "SI"
  Else
      fxFirma = "NO"
  End If
End If
rs.Close

End Function


Private Sub imgHistorico_Click()

Call sbHistorial
        
End Sub

Private Sub imgMonto_Click()
Dim strSQL As String, rs As New ADODB.Recordset
Dim curMonto As Currency, curRebajos As Currency, curCargos As Currency, curMntAdd As Currency
Dim curIntereses As Currency, curPrimerCuota As Currency, iDias As Long
Dim curPoliza As Currency, curPolizaBase As Currency

Dim vGarantia As String, vConvenio As String
Dim vCobraTasaFormaliza As Boolean, vCreditoExcedentes As Boolean
Dim i As Integer, vFecha As Date, curTemp As Currency
  
On Error GoTo vError
  
Me.MousePointer = vbHourglass
  
If txtOperacion.Text = "" Or cboCalculoAdd.Text = "Monto del Crédito" Then
  Me.MousePointer = vbDefault
  Exit Sub
End If
  
  
'Calcula valores fijos
curMntAdd = CCur(txtMonto.Text)


'Rebajos Totales - Cargos (los cargos se recalculan de 0 por eso se debe de excluir)
strSQL = "select dbo.fxCrdSGTMontoDeducciones(" & Operacion.Operacion & ") - dbo.fxCrdCargosOperacion(" & Operacion.Operacion & ") as 'Rebajos'"
Call OpenRecordSet(rs, strSQL)
curRebajos = rs!Rebajos
rs.Close

vCobraTasaFormaliza = fxCobraTasaFormaliza(cboDestino.ItemData(cboDestino.ListIndex))
vCreditoExcedentes = fxCreditoExcedente(Operacion.Codigo)
       

strSQL = "select R.Garantia,R.cuota,R.int,C.convenio,R.FECHA_CALCULO_INT,R.FECHA_INICIO_CALCULO" _
       & " from reg_creditos R inner join Catalogo C on R.codigo = C.codigo" _
       & " where R.id_solicitud =" & Operacion.Operacion
Call OpenRecordSet(rs, strSQL)
    vGarantia = rs!Garantia
    vConvenio = rs!Convenio
    If IsNull(rs!fecha_calculo_int) Then
'       vFecha = fxFechaServidor
'       vFecha = DateAdd("m", 1, vFecha)
'       vFecha = DateAdd("d", -1, CDate(Year(vFecha) & "/" & Format(Month(vFecha), "00") & "/01"))
       vFecha = CDate(txtAno.Text & "/" & Format(fxConvierteMES(cboMes.Text), "00") & "/01 23:59:59")
    Else
       vFecha = rs!fecha_calculo_int
    End If
    

    If vFecha < rs!fecha_inicio_calculo Then
     iDias = 0
    Else
     iDias = vFecha - rs!fecha_inicio_calculo '+ 1
    End If
rs.Close
       
  
       
'---------------------------------------------------------------------------------

Dim vPrideduc As Currency, vDiaPago As Integer, vFechaCalculo As Date

vPrideduc = txtAno.Text & Format(fxConvierteMES(cboMes.Text), "00") & "." & cboFrecuencia.ItemData(cboFrecuencia.ListIndex)

If chkDeducPlanilla.Value = vbChecked Then
    vDiaPago = 32
Else
    strSQL = "select dbo.fxCRDPoliticaPago(dbo.MyGetdate()) as DiaPago" _
           & " From catalogo where Codigo = '" & Operacion.Codigo & "'"
    Call OpenRecordSet(rs, strSQL)
        vDiaPago = rs!DiaPago
    rs.Close
End If

vFechaCalculo = fxFechaCalculo(Operacion.Codigo, vPrideduc, vDiaPago)
       
'---------------------------------------------------------------------------------
       
       
'Base de la Poliza
strSQL = "select CR_PSDMNT from par_ahcr"
Call OpenRecordSet(rs, strSQL)
If rs.EOF And rs.BOF Then
  curPolizaBase = 0
Else
  curPolizaBase = IIf(IsNull(rs!cr_PsdMnt), 0, rs!cr_PsdMnt)
End If
rs.Close

curCargos = 0
curMonto = curRebajos
curTemp = 0


' cboCalculoAdd.AddItem "Monto del Crédito"
' cboCalculoAdd.AddItem "Monto a Girar"
' cboCalculoAdd.AddItem "Giro en Cero"

If cboCalculoAdd.Text = "Monto a Girar" Then
  curMonto = curMonto + curMntAdd
End If

i = 5 'Acercamientos
   
'Inicio de Calculos y Variaciones
For i = 1 To 5
    curIntereses = 0
    If vCobraTasaFormaliza Then
         If Operacion.EstadoSolicitud = "F" Then
            If vCreditoExcedentes Then
                 curIntereses = fxInteresesHastaFormalizar(dtpDesembolso.Value, curMonto)
            Else
                 curIntereses = ((curMonto * CCur(txtTasa.Text)) / (36000)) * iDias
            End If
             
             
         Else
             If chkPrimera.Value = vbChecked Then
                 curIntereses = fxInteresesDiasPrimerCuota(dtpDesembolso.Value, curMonto, txtTasa)
             Else
                 curIntereses = fxInteresesHastaFormalizar(dtpDesembolso.Value, curMonto, vPrideduc, vDiaPago)
             End If
         End If
    End If
           
           
    
    curPrimerCuota = IIf((chkPrimera.Value = vbChecked), txtCuota, 0)
    
    
    'Calcula Poliza
    curPoliza = 0
    If vCobraTasaFormaliza Then
        If vGarantia <> "H" And vConvenio = "N" Then
            curPoliza = (curMonto / 1000000) * curPolizaBase
        End If
    End If
    
    
    'Definir el Monto Base del Credito
    
    
    If cboCalculoAdd.Text = "Monto a Girar" Then
        curMonto = Round(curPoliza, 2) + curPrimerCuota + Round(curIntereses, 2) + curRebajos + Round(curCargos, 2) + curMntAdd
    Else
        curMonto = Round(curPoliza, 2) + curPrimerCuota + Round(curIntereses, 2) + curRebajos + Round(curCargos, 2)
    End If
    
    curTemp = curCargos
    
    'Procesar Cargos y Recuperarlos
'    Call sbCargosAdicionales(Operacion.Operacion, Operacion.Codigo, Round(curMonto, 2))
'    curCargos = fxMontoEnCargos(Operacion.Operacion)
    
    strSQL = "select dbo.fxCrd_Operacion_Cargos_Calcula(" & Operacion.Operacion & ",'" & Operacion.Codigo & "'," & Round(curMonto, 2) & ") as 'Cargos'"
    Call OpenRecordSet(rs, strSQL)
        curCargos = rs!Cargos
    rs.Close
    curMonto = curMonto + (curCargos - curTemp)
    txtCuota.Text = fxCalcula_Cuota(CDbl(Round(curMonto, 2)), txtPlazo, txtTasa, mFrecuenciaPago)

Next i

cboCalculoAdd.Text = "Monto del Crédito"
txtMonto.Text = Format(curMonto, "Standard")

txtMonto.SetFocus

Me.MousePointer = vbDefault

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
 
End Sub

Private Sub imgRecalculoRecurso_Click()
Dim strSQL As String, rs As New ADODB.Recordset

Me.MousePointer = vbHourglass



strSQL = "exec spCRDDisponibleRecurso '" & cboRecursos.ItemData(cboRecursos.ListIndex) & "','" & Format(dtpDesembolso.Value, "yyyy/mm/dd") & "'"
Call OpenRecordSet(rs, strSQL)
If Not rs.EOF And Not rs.BOF Then
    txtDisponibleRecursos = Format(rs!Disponible, "Standard")
Else
    txtDisponibleRecursos = 0
End If
rs.Close

Me.MousePointer = vbDefault

End Sub



Private Sub imgTags_Click()
        If Operacion.Operacion > 0 Then
           Call sbFormsCall("frmCR_SeguimientoEtiquetas", 1, , , False, Me)
        End If
End Sub

Private Function fxInteresDiasX()
Dim strSQL As String, rs As New ADODB.Recordset
Dim iDias As Integer



If Not fxCobraTasaFormaliza(cboDestino.ItemData(cboDestino.ListIndex)) Then
   fxInteresDiasX = 0
   Exit Function
End If

strSQL = "select R.FECHA_CALCULO_INT,isnull(R.FECHA_INICIO_CALCULO, R.fecha_Calculo_Int) as 'FECHA_INICIO_CALCULO',C.convenio,C.retencion,C.poliza,R.montoApr,R.int" _
       & " from reg_creditos R inner join Catalogo C on R.codigo = C.codigo" _
       & " where id_solicitud = " & Operacion.Operacion
Call OpenRecordSet(rs, strSQL)
    If rs!fecha_calculo_int < rs!fecha_inicio_calculo Then
     iDias = 0
    Else
     iDias = rs!fecha_calculo_int - rs!fecha_inicio_calculo + 1
    End If
    
    If rs!Convenio = "S" Or rs!retencion = "S" Or rs!Poliza = "S" Then
      fxInteresDiasX = 0
    Else
      fxInteresDiasX = ((rs!montoapr * rs!Int) / (36000)) * iDias
    End If

rs.Close

End Function

Private Sub lsw_ItemCheck(ByVal Item As XtremeSuiteControls.ListViewItem)
Dim strSQL As String

If vPaso Then Exit Sub
 

Select Case Item.Tag
    Case "F" 'Firmas
        If Item.Checked Then
            If Mid(Item.SubItems(1), 1, 1) = "D" Then
                strSQL = "update reg_creditos set firma_deudor = 1,fechaforf = dbo.MyGetdate()" _
                       & " where id_solicitud = " & Operacion.Operacion
            
            Else
                strSQL = "update fiadores set firma = 'S'" _
                       & " where ID_SOLiCITUD = " & Operacion.Operacion & " and CedulaF = '" & Item.Text & "'"
            End If
            
        Else
            If Mid(Item.SubItems(1), 1, 1) = "D" Then
                strSQL = "update reg_creditos set firma_deudor = 0,fechaforf = dbo.MyGetdate()" _
                       & " where id_solicitud = " & Operacion.Operacion
            Else
                strSQL = "update fiadores set firma = 'N'" _
                       & " where ID_SOLiCITUD = " & Operacion.Operacion & " and CedulaF = '" & Item.Text & "'"
            End If
        End If
    
        Call ConectionExecute(strSQL)
        
        
End Select



End Sub


Private Sub lswBusca_ItemClick(ByVal Item As XtremeSuiteControls.ListViewItem)
 
 txtOperacion.Text = Item.Text
 Call sbCargaOperacion

End Sub


Private Sub lswOpciones_Click()
Dim strSQL As String, rs As New ADODB.Recordset
Dim xItem As String, i As Integer, curMonto As Currency
Dim itmX As ListViewItem

xItem = lswOpciones.SelectedItem.Key
With lswOpciones.ListItems
 For i = 1 To .Count
   If .Item(i).Key = xItem Then
      .Item(i).SmallIcon = 2
   Else
      .Item(i).SmallIcon = 1
   End If
 Next i
End With

lsw.ListItems.Clear
lsw.ColumnHeaders.Clear
lsw.HideColumnHeaders = False
lsw.Checkboxes = False

Select Case xItem
 Case "CRD"
   lsw.ColumnHeaders.Add , , "Operación", 980
   lsw.ColumnHeaders.Add , , "Línea", 980, vbCenter
   lsw.ColumnHeaders.Add , , "Monto", 1280, vbRightJustify
   lsw.ColumnHeaders.Add , , "Tipo", 1280, vbCenter
   lsw.ColumnHeaders.Add , , "Descripción", 3980
 
   curMonto = 0
   strSQL = "select R.*,C.descripcion, Isnull(Cargos,0) as 'CargosX'" _
          & " from refundiciones R inner join catalogo C on R.codigo = C.codigo" _
          & " Where R.id_solicitudr = " & Operacion.Operacion
   Call OpenRecordSet(rs, strSQL, 0)
   Do While Not rs.EOF
    Set itmX = lsw.ListItems.Add(, , rs!ID_SOLICITUD)
        itmX.SubItems(1) = rs!Codigo
        itmX.SubItems(2) = Format(rs!Monto, "Standard")
        
        Select Case rs!Tipo
            Case "C"
                itmX.SubItems(3) = "Cancela"
            Case "M"
                itmX.SubItems(3) = "Morosidad"
            Case "P"
                itmX.SubItems(3) = "Pendientes"
        End Select
        
        itmX.SubItems(4) = rs!Descripcion
        curMonto = curMonto + rs!Monto
    rs.MoveNext
   Loop
   rs.Close
   Set itmX = lsw.ListItems.Add(, , "")
       itmX.SubItems(2) = "_______"
   Set itmX = lsw.ListItems.Add(, , "")
       itmX.SubItems(2) = Format(curMonto, "Standard")
 
 
 Case "DES"
   lsw.ColumnHeaders.Add , , "Concepto", 3280
   lsw.ColumnHeaders.Add , , "Monto", 1280, vbRightJustify
   lsw.ColumnHeaders.Add , , "Retiene", 1180, vbCenter
 
 
   curMonto = 0
   strSQL = "select * " _
          & " from desembolsos Where id_solicitud = " & Operacion.Operacion
   Call OpenRecordSet(rs, strSQL, 0)
   Do While Not rs.EOF
    Set itmX = lsw.ListItems.Add(, , rs!CONCEPTO)
        itmX.SubItems(1) = Format(rs!Monto, "Standard")
        itmX.SubItems(2) = IIf((rs!retener = 1), "SI", "NO")
        curMonto = curMonto + rs!Monto
    rs.MoveNext
   Loop
   rs.Close
   Set itmX = lsw.ListItems.Add(, , "")
       itmX.SubItems(1) = "_______"
   Set itmX = lsw.ListItems.Add(, , "")
       itmX.SubItems(1) = Format(curMonto, "Standard")
 
 
 Case "RET"
   lsw.ColumnHeaders.Add , , "Operación", 980
   lsw.ColumnHeaders.Add , , "Línea", 980
   lsw.ColumnHeaders.Add , , "Monto", 1280, vbRightJustify
   lsw.ColumnHeaders.Add , , "Descripción", 3980
 
   curMonto = 0
   strSQL = "select R.*,C.descripcion" _
          & " from REFUNDE_RETENCION R inner join catalogo C on R.codigo = C.codigo" _
          & " Where R.id_solicitudr = " & Operacion.Operacion
   Call OpenRecordSet(rs, strSQL, 0)
   Do While Not rs.EOF
    Set itmX = lsw.ListItems.Add(, , rs!ID_SOLICITUD)
        itmX.SubItems(1) = rs!Codigo
        itmX.SubItems(2) = Format(rs!Monto, "Standard")
        itmX.SubItems(3) = rs!Descripcion
        curMonto = curMonto + rs!Monto
    rs.MoveNext
   Loop
   rs.Close
   Set itmX = lsw.ListItems.Add(, , "")
       itmX.SubItems(2) = "_______"
   Set itmX = lsw.ListItems.Add(, , "")
       itmX.SubItems(2) = Format(curMonto, "Standard")
 
 Case "FIR"
   lsw.ColumnHeaders.Add , , "Identificación", 1800
   lsw.ColumnHeaders.Add , , "Tipo", 1080
   lsw.ColumnHeaders.Add , , "Nombre", 4280
   lsw.ColumnHeaders.Add , , "Firma", 890
       
   lsw.Checkboxes = True
   
   vPaso = True
   
    Set itmX = lsw.ListItems.Add(, , Operacion.Cedula)
        itmX.SubItems(1) = "Deudor"
        itmX.SubItems(2) = Operacion.Nombre
        itmX.SubItems(3) = fxFirma(Operacion.Operacion)
        
        If itmX.SubItems(3) = "SI" Then
            itmX.TextBackColor = RGB(208, 253, 235)
            itmX.Bold = True
            itmX.Checked = True
        Else
            itmX.TextBackColor = RGB(249, 253, 208)
            itmX.Bold = True
        End If
     
     itmX.Tag = "F"
     
   strSQL = "select F.*,S.nombre" _
          & " from Fiadores F inner join Socios S on F.cedulaf = S.cedula" _
          & " Where F.estado = 'A' and F.id_solicitud = " & Operacion.Operacion
   Call OpenRecordSet(rs, strSQL, 0)
   Do While Not rs.EOF
    Set itmX = lsw.ListItems.Add(, , rs!cedulaf)
        If rs!Calidad = "F" Then
            itmX.SubItems(1) = "Fiadores"
        Else
            itmX.SubItems(1) = "Co-Dedudor"
        End If
        
        itmX.SubItems(2) = rs!Nombre
        itmX.SubItems(3) = IIf((rs!Firma = "S"), "SI", "NO")
        If rs!Firma = "S" Then
            itmX.TextBackColor = RGB(208, 253, 235)
            itmX.Bold = True
            itmX.Checked = True
        Else
            itmX.TextBackColor = RGB(249, 253, 208)
            itmX.Bold = True
        End If
        
     'Indica que se trata de firmas
     
     itmX.Tag = "F"
    rs.MoveNext
   Loop
   rs.Close
   
   vPaso = False
   
 Case "REQ"
   lsw.ColumnHeaders.Add , , "Estado", 780
   lsw.ColumnHeaders.Add , , "Descripción", 5280
   lsw.ColumnHeaders.Add , , "Visible", 80
'   lsw.Checkboxes = True
   
   strSQL = "select O.*,R.descripcion, R.visible" _
          & " from requisitos_adicionales R inner join  operacion_requisitos O on R.cod_requisito = O.cod_requisito" _
          & " where O.id_solicitud = " & Operacion.Operacion & " order by O.estado,R.cod_requisito"
   Call OpenRecordSet(rs, strSQL, 0)
   Do While Not rs.EOF
    Set itmX = lsw.ListItems.Add(, , rs!COD_REQUISITO)
        itmX.SubItems(1) = rs!Descripcion
        itmX.SubItems(2) = rs!Visible
        itmX.Tag = "R"
        
        Select Case rs!Estado
           Case 0 'Amarillo
            itmX.TextBackColor = RGB(249, 253, 208)
            itmX.Bold = True
           Case 1 'Verde
            itmX.TextBackColor = RGB(208, 253, 235)
            itmX.Bold = True
            itmX.Checked = True
           Case 2 'Rojo
            itmX.TextBackColor = RGB(236, 62, 99)
            itmX.Bold = True
            itmX.SubItems(2) = 0
        End Select
    rs.MoveNext
   Loop
   rs.Close
          
 Case "CAR"
     lsw.ColumnHeaders.Add , , "Código", 900
     lsw.ColumnHeaders.Add , , "Descripción", 2900
     lsw.ColumnHeaders.Add , , "Monto", 1200, vbRightJustify
     lsw.ColumnHeaders.Add , , "Base", 1200
     lsw.ColumnHeaders.Add , , "Tipo", 1200
     lsw.ColumnHeaders.Add , , "Valor", 1200, vbRightJustify
    
     curMonto = 0
     strSQL = "exec spCrdOperacionFormalizaCargosLista " & Operacion.Operacion
     Call OpenRecordSet(rs, strSQL)
     Do While Not rs.EOF
       Set itmX = lsw.ListItems.Add(, , rs!COD_CARGO)
           itmX.SubItems(1) = rs!Descripcion
           itmX.SubItems(2) = Format(rs!Monto, "Standard")
           Select Case rs!Base
             Case "C"
               itmX.SubItems(3) = "Crédito"
               curMonto = curMonto + rs!Monto
             Case "A"
               itmX.SubItems(3) = "Avalúo"
               curMonto = curMonto + rs!Monto
             Case "P"
               itmX.SubItems(3) = "Prima"
               itmX.ForeColor = vbBlue
           End Select
           itmX.SubItems(4) = IIf((rs!Tipo = "P"), "Porcentaje", "Monto")
           itmX.SubItems(5) = Format(rs!Valor, "Standard")
           
       rs.MoveNext
     Loop
     rs.Close
     
     Set itmX = lsw.ListItems.Add(, , "")
         itmX.SubItems(2) = "________"
     Set itmX = lsw.ListItems.Add(, , "")
         itmX.SubItems(2) = Format(curMonto, "Standard")
 
 Case "FIA"
   lsw.ColumnHeaders.Add , , "Identificación", 1800
   lsw.ColumnHeaders.Add , , "Nombre", 4280
   strSQL = "select S.cedula,S.nombre" _
          & " from Fiadores F inner join Socios S on F.cedulaf = S.cedula" _
          & " Where F.estado = 'A' and F.id_solicitud = " & Operacion.Operacion
   Call OpenRecordSet(rs, strSQL, 0)
   Do While Not rs.EOF
    Set itmX = lsw.ListItems.Add(, , rs!Cedula)
        itmX.SubItems(1) = rs!Nombre
    rs.MoveNext
   Loop
   rs.Close
   
   
 Case "ILI" 'Impacto en Liquidadez
   lsw.ColumnHeaders.Add , , "", 2280
   lsw.ColumnHeaders.Add , , "", 2100, vbRightJustify
   strSQL = "exec spCrd_SGT_Impacto_Liquidez " & Operacion.Operacion
   Call OpenRecordSet(rs, strSQL, 0)
   If Not rs.EOF And Not rs.BOF Then
    
    Set itmX = lsw.ListItems.Add(, , "Cuota Nueva: ")
        itmX.SubItems(1) = Format(rs!Cuota_Nueva, "Standard")
    
    Set itmX = lsw.ListItems.Add(, , "Cuotas Liberadas: ")
        itmX.SubItems(1) = Format(rs!Cuota_Libera, "Standard")
    
    Set itmX = lsw.ListItems.Add(, , "")
        itmX.SubItems(1) = ""
    
    Set itmX = lsw.ListItems.Add(, , "Impacto en Liquidez: ")
        itmX.SubItems(1) = Format(rs!Impacto, "Standard")
        If rs!Impacto < 0 Then
           itmX.ForeColor = vbRed
        End If
        itmX.Bold = True
    
   End If
   rs.Close
   
   
 Case "RSM"
   Call sbResumenOperacion

End Select


End Sub

Private Sub sbResumenOperacion()
Dim strSQL As String, rs As New ADODB.Recordset, rsTotal As New ADODB.Recordset
Dim vPrideduc As Currency, vDiaPago As Integer, vFechaCalculo As Date
Dim xItem As String, i As Integer, curMonto As Currency
Dim itmX As ListViewItem

'Inicializa
lsw.ListItems.Clear
lsw.ColumnHeaders.Clear


lsw.ColumnHeaders.Add , , "", 2480
lsw.ColumnHeaders.Add , , "", 1360, vbRightJustify
lsw.ColumnHeaders.Add , , "", 1520
   
lsw.HideColumnHeaders = True
   
   
vPrideduc = txtAno.Text & Format(fxConvierteMES(cboMes.Text), "00") & "." & cboFrecuencia.ItemData(cboFrecuencia.ListIndex)


If chkDeducPlanilla.Value = vbChecked Then
    vDiaPago = 32
Else
    strSQL = "select dbo.fxCRDPoliticaPago(dbo.MyGetdate()) as DiaPago" _
           & " From catalogo where Codigo = '" & Operacion.Codigo & "'"
    Call OpenRecordSet(rs, strSQL)
        vDiaPago = rs!DiaPago
    rs.Close
End If

strSQL = "exec spCrdSGTResumen " & Operacion.Operacion
Call OpenRecordSet(rsTotal, strSQL)

vFechaCalculo = fxFechaCalculo(Operacion.Codigo, vPrideduc, vDiaPago)
   
   Set itmX = lsw.ListItems.Add(, , "-> Monto Aprobado")
       itmX.SubItems(1) = Format(Operacion.MontoAprobado, "Standard")
       curMonto = Operacion.MontoAprobado
   
   Set itmX = lsw.ListItems.Add(, , "(-) Refundiciones CRD")
'       itmX.SubItems(1) = Format(fxMontoEnRefundiciones(Operacion.Operacion), "Standard")
       itmX.SubItems(1) = Format(rsTotal!REFUNDICIONES, "Standard")
       itmX.ForeColor = vbRed
       curMonto = curMonto - CCur(itmX.SubItems(1))

   Set itmX = lsw.ListItems.Add(, , "(-) Desembolsos y Rebajos")
'       itmX.SubItems(1) = Format(fxMontoEnDesembolsos(Operacion.Operacion), "Standard")
       itmX.SubItems(1) = Format(rsTotal!DESEMBOLSOS, "Standard")
       itmX.ForeColor = vbRed
       curMonto = curMonto - CCur(itmX.SubItems(1))

   Set itmX = lsw.ListItems.Add(, , "(-) Refund.Retenciones")
'       itmX.SubItems(1) = Format(fxMontoEnRetenciones(Operacion.Operacion), "Standard")
       itmX.SubItems(1) = Format(rsTotal!Retenciones, "Standard")
       itmX.ForeColor = vbRed
       curMonto = curMonto - CCur(itmX.SubItems(1))
'       lsw.Tag = fxMontoEnDesembolsosRetenidos(Operacion.Operacion) + fxMontoEnCargos(Operacion.Operacion)
        lsw.Tag = rsTotal!DesembolsosRet + rsTotal!Cargos
       
   Set itmX = lsw.ListItems.Add(, , "(-) Dias de Interes")
       If fxCobraTasaFormaliza(cboDestino.ItemData(cboDestino.ListIndex)) Then
            If Operacion.EstadoSolicitud = "F" Then
               If fxCreditoExcedente(Operacion.Codigo) Then
                    itmX.SubItems(1) = Format(fxInteresesHastaFormalizar(dtpDesembolso.Value, , vPrideduc, vDiaPago), "Standard")
               Else
                    itmX.SubItems(1) = Format(fxInteresDiasX, "Standard")
               End If
                
                
            Else
                
                    itmX.SubItems(1) = Format(fxInteresesHastaFormalizar(dtpDesembolso.Value, , vPrideduc, vDiaPago), "Standard")
                'Eliminado el 2023-12-20
'                If chkPrimera.Value = vbChecked Then
'                    itmX.SubItems(1) = Format(fxInteresesDiasPrimerCuota(dtpDesembolso.Value, Operacion.MontoAprobado, txtTasa), "Standard")
'                Else
'                    itmX.SubItems(1) = Format(fxInteresesHastaFormalizar(dtpDesembolso.Value, , vPrideduc, vDiaPago), "Standard")
'                End If
            End If
               
           itmX.SubItems(2) = "(" & DateDiff("d", dtpDesembolso.Value, vFechaCalculo) + 1 & ") " & Format(vFechaCalculo, "dd/mm/yyyy")
       Else
           itmX.SubItems(1) = "0.00"
       End If
       
       itmX.ForeColor = vbRed
       curMonto = curMonto - CCur(itmX.SubItems(1))

   Set itmX = lsw.ListItems.Add(, , "(-) Primer Cuota")
       itmX.SubItems(1) = Format(IIf((chkPrimera.Value = vbChecked), txtCuota, 0), "Standard")
       itmX.ForeColor = vbRed
       curMonto = curMonto - CCur(itmX.SubItems(1))

    strSQL = "select R.Primer_Cuota,R.Garantia,R.montoapr,R.cuota,R.int,C.convenio" _
           & " from reg_creditos R inner join Catalogo C on R.codigo = C.codigo" _
           & " where R.id_solicitud =" & Operacion.Operacion
    Call OpenRecordSet(rs, strSQL)
    Set itmX = lsw.ListItems.Add(, , "(-) P.S.D.")
    
    If fxCobraTasaFormaliza(cboDestino.ItemData(cboDestino.ListIndex)) Then
        If rs!Garantia <> "H" And rs!Convenio = "N" Then
            itmX.SubItems(1) = Format(fxCuotaPolizaVida(rs!montoapr), "Standard")
        Else
            itmX.SubItems(1) = Format(0, "Standard")
        End If
    
    Else
            itmX.SubItems(1) = Format(0, "Standard")
    End If
    rs.Close
        itmX.ForeColor = vbRed
        curMonto = curMonto - CCur(itmX.SubItems(1))

   
   Set itmX = lsw.ListItems.Add(, , "(-) Cargos Adicionales")
'       itmX.SubItems(1) = Format(fxMontoEnCargos(Operacion.Operacion), "Standard")
        itmX.SubItems(1) = Format(rsTotal!Cargos, "Standard")
       itmX.ForeColor = vbRed
       curMonto = curMonto - CCur(itmX.SubItems(1))


   Set itmX = lsw.ListItems.Add(, , "")
       itmX.SubItems(1) = "___________"
       
   Set itmX = lsw.ListItems.Add(, , "Monto a Girar")
       itmX.SubItems(1) = Format(curMonto, "Standard")



End Sub

Private Sub lswOpciones_DblClick()
Dim strSQL As String, rs As New ADODB.Recordset

'Si la Operacion esta formalizada o Anulada entonces no ingresar a mantenimiento
If Not lsw.Enabled Then Exit Sub


Operacion.FechaDesembolso = dtpDesembolso.Value
Operacion.PriDeduc = txtAno.Text & Format(fxConvierteMES(cboMes.Text), "00")
If chkDeducPlanilla.Value = vbChecked Then
   Operacion.DiaPago = 32
Else
   strSQL = "select dbo.fxCRDPoliticaPago(dbo.MyGetdate()) as DiaPago"
   Call OpenRecordSet(rs, strSQL)
     Operacion.DiaPago = rs!DiaPago
   rs.Close
End If




Select Case lswOpciones.SelectedItem.Key
 
 Case "CRD"
  strSQL = "select refunde from catalogo where codigo = '" & Operacion.Codigo & "'"
  Call OpenRecordSet(rs, strSQL)
  If rs!Refunde = "S" Then
      Call sbFormsCall("frmCR_SeguimientoRefundiciones", 1, , , False, Me)
  Else
    MsgBox "Esta Línea No Permite que se realicen refundiciones con ella...", vbCritical
  End If
  rs.Close
 
 Case "DES"
      Call sbFormsCall("frmCR_SeguimientoDesembolsos", 1, , , False, Me)
 
 Case "RET"
  strSQL = "select refunde from catalogo where codigo = '" & Operacion.Codigo & "'"
  Call OpenRecordSet(rs, strSQL)
  If rs!Refunde = "S" Then
      Call sbFormsCall("frmCR_SeguimientoRetenciones", 1, , , False, Me)
  Else
    MsgBox "Esta Línea No Permite que se realicen refundiciones con ella...", vbCritical
  End If
  rs.Close
  
 Case "FIR"
   Call sbFormsCall("frmCR_SeguimientoFirmas", 1, , , False, Me)

 Case "REQ"
   Operacion.Ventana = "R"
   Call sbFormsCall("frmCR_SeguimientoReqCar", 1, , , False, Me)
 
 Case "CAR"
   Operacion.Ventana = "C"
   Call sbFormsCall("frmCR_SeguimientoReqCar", 1, , , False, Me)

 Case "FIA"
   Call sbFormsCall("frmCR_SolicitudesFiadores", 1, , , False, Me)
 
 Case "RSM" 'Nada
End Select

'Refresca datos en Pantalla
Call lswOpciones_Click

End Sub


Private Sub tcMain_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
TituloOpcion.Caption = Item.Caption

'Carga Opciones: Formalización
If Item.Index = 1 Then
    Call sbSGTInitOpciones(lswOpciones)
    lsw.ListItems.Clear
End If

End Sub

Private Sub TimerX_Timer()
TimerX.Interval = 0
TimerX.Enabled = False


 
 Call sbLimpiaDatos
 Call sbCargaCombos

End Sub

Private Function fxGarantiaForm(pGarantia As String) As String
Dim strSQL As String, rs As New ADODB.Recordset
Dim vResult As String

On Error GoTo vError

vResult = ""

strSQL = "select FORMULARIO" _
       & " From CRD_GARANTIA_TIPOS" _
       & " Where Garantia = '" & pGarantia & "'"
Call OpenRecordSet(rs, strSQL)

vResult = rs!Formulario

fxGarantiaForm = vResult
Exit Function

vError:
  fxGarantiaForm = vResult

End Function


Private Sub sbTaskPanel_Accion(ItemId As Integer)
Dim strSQL As String, rs As New ADODB.Recordset
Dim Modulo_Hipotecario As clsHipotecario


If Operacion.Operacion = 0 Then
    Select Case ItemId
        Case Id_TaskItem_Cargos, Id_TaskItem_Causas, Id_TaskItem_Coberturas _
             , Id_TaskItem_Desembolsos, Id_TaskItem_Firmas, Id_TaskItem_Garantia _
             , Id_TaskItem_MntNoGravable, Id_TaskItem_PlanPagos, Id_TaskItem_Polizas _
             , Id_TaskItem_Requisitos, Id_TaskItem_Seguimiento, Id_TaskItem_PreCalculo _
             , Id_TaskItem_DatosPersonales, Id_TaskItem_Formalizacion, Id_TaskItem_Historial
             
             MsgBox "Esta opción requiere que se haya registrado la operacion!", vbInformation
             Exit Sub
        Case Else
    End Select
End If

Me.MousePointer = vbHourglass

Select Case ItemId
   Case Id_TaskItem_Recepcion
        tcMain.Item(0).Selected = True
   Case Id_TaskItem_Formalizacion
   
        If Operacion.Operacion > 0 Then
            tcMain.Item(1).Selected = True
            Call sbResumenOperacion
        End If

   Case Id_TaskItem_Historial
        Call sbHistorial

   Case Id_TaskItem_Estudio  'Estudio
    
        
        Dim Modulo_Estudio As clsEstudioCrd
        
        Set Modulo_Estudio = New clsEstudioCrd
        Set Modulo_Estudio.vCon = glogon.Conection
        Modulo_Estudio.xOperacion = Operacion.Operacion
        Modulo_Estudio.xkey = glogon.ConectRPT
  
        strSQL = "select cod_preAnalisis from CRD_PREA_PREANALISIS" _
               & " Where id_solicitud = " & Operacion.Operacion
        Call OpenRecordSet(rs, strSQL)
        If rs.EOF And rs.BOF Then
                Call Modulo_Estudio.Menu(glogon.Usuario, glogon.Conection, glogon.BaseDatos, glogon.Servidor _
                            , App.Path, glogon.ConectRPT, 11, glogon.AppName, glogon.AppVersion, glogon.Maquina _
                            , glogon.Portal_Con, glogon.Portal_User, glogon.Portal_Key, gPortal.Empresa_Id, gPortal.Empresa_Name)
        
        Else
            Modulo_Estudio.vSolicitudPreanalisis = rs!cod_PreAnalisis
                Call Modulo_Estudio.Menu(glogon.Usuario, glogon.Conection, glogon.BaseDatos, glogon.Servidor _
                            , App.Path, glogon.ConectRPT, 11, glogon.AppName, glogon.AppVersion, glogon.Maquina _
                            , glogon.Portal_Con, glogon.Portal_User, glogon.Portal_Key, gPortal.Empresa_Id, gPortal.Empresa_Name)
        End If
        rs.Close
        Set Modulo_Estudio = Nothing
   
   Case Id_TaskItem_Garantia  'Garantía

        Select Case fxGarantiaForm(cboGarantia.ItemData(cboGarantia.ListIndex))
            Case "F02" 'Fiadores
                    If Operacion.EstadoSolicitud = "R" Then
                      Call sbFormsCall("frmCR_SolicitudesFiadores", 1, , , False, Me)
                    End If
            
            Case "F03" 'Hipotecaria
                    
                    Set Modulo_Hipotecario = New clsHipotecario
                    Set Modulo_Hipotecario.vCon = glogon.Conection
                    Modulo_Hipotecario.xOperacion = Operacion.Operacion
                    Modulo_Hipotecario.xkey = glogon.ConectRPT
                    Modulo_Hipotecario.xToolBar = gToolBar
                    
                    Call Modulo_Hipotecario.Menu(glogon.Usuario, glogon.Conection, glogon.BaseDatos, glogon.Servidor _
                        , App.Path, glogon.ConectRPT, 2, glogon.AppName, glogon.AppVersion, glogon.Maquina _
                        , glogon.Portal_Con, glogon.Portal_User, glogon.Portal_Key, gPortal.Empresa_Id, gPortal.Empresa_Name)
            
            Case "F07" 'Prendas
            
                    Operacion.GarantiaTipo = "P" 'Prenda
                    Operacion.GarantiaId = 0
                    
                    Operacion.Expendiente = ""
                    Operacion.GarantiaParam = "E" 'Estudio
                    
                    Operacion.Operacion = txtOperacion.Text
                    Operacion.GarantiaParam = "C" 'Credito
                    Operacion.Cedula = Trim(txtCedula.Text)
            
            
'                    If Operacion.EstadoSolicitud = "R" Then
                      Call sbFormsCall("frmCR_Prendas", vbModal, , , False, Me, True)
'                    End If
             

             
            Case Else
                    MsgBox "Esta operación es solo para Prendas, Hipotecas y Fianzas!", vbInformation
        End Select
        
   Case Id_TaskItem_Coberturas 'Garantia - Cobertura
        If cboGarantia.ItemData(cboGarantia.ListIndex) = "H" Then
                Set Modulo_Hipotecario = New clsHipotecario
                Set Modulo_Hipotecario.vCon = glogon.Conection
                Modulo_Hipotecario.xOperacion = Operacion.Operacion
                Modulo_Hipotecario.xkey = glogon.ConectRPT
                Modulo_Hipotecario.xToolBar = gToolBar
                
                Call Modulo_Hipotecario.Menu(glogon.Usuario, glogon.Conection, glogon.BaseDatos, glogon.Servidor _
                    , App.Path, glogon.ConectRPT, 12, glogon.AppName, glogon.AppVersion, glogon.Maquina _
                    , glogon.Portal_Con, glogon.Portal_User, glogon.Portal_Key, gPortal.Empresa_Id, gPortal.Empresa_Name)
           
        Else
            MsgBox "Esta Opción es solo para Garantías Hipotecarias!", vbInformation
           
        End If
   
   Case Id_TaskItem_Desembolsos 'Garantia - Control de Desembolsos
        If cboGarantia.ItemData(cboGarantia.ListIndex) = "H" Then
            Set Modulo_Hipotecario = New clsHipotecario
            Set Modulo_Hipotecario.vCon = glogon.Conection
            Modulo_Hipotecario.xOperacion = Operacion.Operacion
            Modulo_Hipotecario.xkey = glogon.ConectRPT
            Modulo_Hipotecario.xToolBar = gToolBar
            
            Call Modulo_Hipotecario.Menu(glogon.Usuario, glogon.Conection, glogon.BaseDatos, glogon.Servidor _
                , App.Path, glogon.ConectRPT, 4, glogon.AppName, glogon.AppVersion, glogon.Maquina _
                , glogon.Portal_Con, glogon.Portal_User, glogon.Portal_Key, gPortal.Empresa_Id, gPortal.Empresa_Name)
        Else
            MsgBox "Esta Opción es solo para Garantías Hipotecarias con Control de desembolsos!", vbInformation
        End If
   
   Case Id_TaskItem_Seguimiento  'Tag's de Seguimiento
           Call sbFormsCall("frmCR_SeguimientoEtiquetas", 1, , , False, Me)

   Case Id_TaskItem_Polizas  'Pólizas
        If Operacion.EstadoSolicitud = "F" Then
            Operacion.OperacionConsulta = txtOperacion.Text
            Call sbFormsCall("frmCR_PolizasRegistro", 1, , , False, Me)
        Else
              MsgBox "Esta Opción solo se activa cuando la operación se encuentra Formalizada", vbInformation
        End If
   
   Case Id_TaskItem_Requisitos   'Requisitos
        If Operacion.EstadoSolicitud = "R" Then
              Operacion.Ventana = "R"
              Call sbFormsCall("frmCR_SeguimientoReqCar", 1, , , False, Me)
        Else
              MsgBox "Esta Opción solo se activa cuando la operación se encuentra RECIBIDA", vbInformation
        End If
        
        
   Case Id_TaskItem_Cargos   'Cargos
        If Operacion.EstadoSolicitud = "R" Then
              Operacion.Ventana = "C"
              Call sbFormsCall("frmCR_SeguimientoReqCar", 1, , , False, Me)
        Else
              MsgBox "Esta Opción solo se activa cuando la operación se encuentra RECIBIDA", vbInformation
        End If

   Case Id_TaskItem_Causas  'Causas
          If Operacion.EstadoSolicitud = "P" Or Operacion.EstadoSolicitud = "D" Then
            Call sbFormsCall("frmCR_SeguimientoCausas", 1, , , False, Me)
          Else
              MsgBox "Esta Opción solo se activa cuando la operación ha sido DENEGADA o Puesto en estado PENDIENTE", vbInformation
          End If
        
        
   Case Id_TaskItem_PlanPagos  'Plan de Pagos
            Operacion.OperacionConsulta = txtOperacion.Text
            Call sbFormsCall("frmCR_PlanPagos", 1, , , False, Me)
   
   Case Id_TaskItem_DatosPersonales  'Datos Personales
          GLOBALES.gCedulaActual = Operacion.Cedula
          Call sbFormsCall("frmCR_VerificaDatosPersonales", 1, , , False, Me)
   
   Case Id_TaskItem_PreCalculo  'Pre Calculo
            GLOBALES.gCedulaActual = Operacion.Cedula
            Call sbFormsCall("frmCR_CalculoOperacion", 0)
        
   Case Id_TaskItem_Firmas 'Firmas
        If Operacion.Operacion > 0 And Operacion.EstadoSolicitud = "R" Then
          Call sbFormsCall("frmCR_SeguimientoFirmas", 1, , , False, Me)
        Else
            MsgBox "Esta Opción solo se activa cuando la operación se encuentra RECIBIDA", vbInformation
        End If
        
   Case Id_TaskItem_MntNoGravable 'Monto no Gravable
        tcMain.Item(4).Selected = True
        
End Select


Me.MousePointer = vbDefault


End Sub

Private Sub tpMain_ItemClick(ByVal Item As XtremeTaskPanel.ITaskPanelGroupItem)

Call sbTaskPanel_Accion(Item.Id)
  
End Sub

Private Sub txtMntNoGravable_GotFocus()
On Error GoTo vError

txtMntNoGravable.Text = CCur(txtMntNoGravable.Text)

vError:
End Sub

Private Sub txtMntNoGravable_LostFocus()
On Error GoTo vError

txtMntNoGravable.Text = Format(CCur(txtMntNoGravable.Text), "Standard")

vError:
End Sub

Private Sub txtMonto_GotFocus()
On Error GoTo vError

txtMonto.Text = CCur(txtMonto.Text)

vError:

End Sub

Private Sub txtMonto_KeyPress(KeyAscii As Integer)
 If KeyAscii = vbKeyReturn And IsNumeric(txtMonto.Text) Then
   If cboGarantia.ItemData(cboGarantia.ListIndex) = "Y" And cboFondo.ListCount > 0 Then
       If lblPlazo.Tag = 0 Then txtPlazo.Text = fxCatalogoRango(txtCodigo, txtMonto.Text, "P", cboDestino.ItemData(cboDestino.ListIndex), cboGarantia.ItemData(cboGarantia.ListIndex))
       If lblTasa.Tag = 0 Then txtTasa.Text = fxCatalogoRango(txtCodigo, txtMonto.Text, "I", cboDestino.ItemData(cboDestino.ListIndex), cboGarantia.ItemData(cboGarantia.ListIndex)) - Operacion.TasaPtsBono
        txtPlazo.SetFocus
   Else
        If Operacion.PlazoBono = 0 Then
            txtPlazo.Text = fxCatalogoRango(txtCodigo, txtMonto.Text, "P", cboDestino.ItemData(cboDestino.ListIndex), cboGarantia.ItemData(cboGarantia.ListIndex))
        End If
        txtTasa.Text = fxCatalogoRango(txtCodigo, txtMonto.Text, "I", cboDestino.ItemData(cboDestino.ListIndex), cboGarantia.ItemData(cboGarantia.ListIndex)) - Operacion.TasaPtsBono
        txtPlazo.SetFocus
   End If
 End If
End Sub


Private Sub Edicion(intActiva As Integer)
'Activa e inactiva partes a editar

If intActiva = 1 Then
  fraOperacion.Enabled = True
  Select Case Operacion.EstadoSolicitud
   Case "R", "P"
     Me.txtCedula.Enabled = True
     Me.txtCodigo.Enabled = True
     Me.txtMonto.Enabled = True
     Me.txtPlazo.Enabled = True
     Me.txtTasa.Enabled = True
     Me.cboBanco.Enabled = True
     Me.cboComite.Enabled = True
     Me.cboGarantia.Enabled = True
     Me.cboEstado.Enabled = True
     Me.txtObservaciones.Enabled = True
     Me.cboCuenta.Enabled = True
     Me.dtpFechaSolicitud.Enabled = True
     Me.cboTipoDocumento.Enabled = True
     Me.imgMonto.Enabled = True

   Case "A"
     Me.txtCedula.Enabled = True
     Me.txtCodigo.Enabled = True
     Me.cboBanco.Enabled = True
     Me.cboComite.Enabled = True
     Me.cboGarantia.Enabled = True
     Me.txtObservaciones.Enabled = True
     Me.cboCuenta.Enabled = True
     Me.cboTipoDocumento.Enabled = True
     
     Me.dtpFechaSolicitud.Enabled = False
     Me.txtMonto.Enabled = False
     Me.txtPlazo.Enabled = False
     Me.txtTasa.Enabled = False
     Me.cboEstado.Enabled = False
   Case "D", "N"
     Me.txtObservaciones.Enabled = True
     Me.txtCedula.Enabled = False
     Me.txtCodigo.Enabled = False
     Me.cboBanco.Enabled = False
     Me.cboComite.Enabled = False
     Me.cboGarantia.Enabled = False
     Me.cboCuenta.Enabled = False
     Me.dtpFechaSolicitud.Enabled = False
     Me.txtMonto.Enabled = False
     Me.txtPlazo.Enabled = False
     Me.txtTasa.Enabled = False
     Me.cboEstado.Enabled = False
     Me.cboTipoDocumento.Enabled = False
   Case "F"
     Me.txtObservaciones.Enabled = True
     Me.txtCedula.Enabled = False
     Me.txtCodigo.Enabled = False
     Me.cboBanco.Enabled = True
     Me.cboComite.Enabled = False
     Me.cboGarantia.Enabled = False
     Me.cboCuenta.Enabled = True
     Me.dtpFechaSolicitud.Enabled = False
     Me.txtMonto.Enabled = False
     Me.txtPlazo.Enabled = False
     Me.txtTasa.Enabled = False
     Me.cboEstado.Enabled = False
     Me.cboTipoDocumento.Enabled = True
  End Select
Else 'apaga
  fraOperacion.Enabled = False
     Me.txtCedula.Enabled = True
     Me.txtCodigo.Enabled = True
     Me.txtMonto.Enabled = True
     Me.txtPlazo.Enabled = True
     Me.txtTasa.Enabled = True
     Me.cboBanco.Enabled = True
     Me.cboComite.Enabled = True
     Me.cboGarantia.Enabled = True
     Me.cboEstado.Enabled = True
     Me.txtObservaciones.Enabled = True
     Me.cboCuenta.Enabled = True
     Me.dtpFechaSolicitud.Enabled = True
     Me.cboTipoDocumento.Enabled = True
  Select Case Operacion.EstadoSolicitud
   Case "A"
   Case "D"
   Case "N"
   Case "F"
  End Select
End If 'inactiva
End Sub

Private Sub sbCargosAdicionales(vOperacion As Long, vCodigo As String, vMonto As Currency)
Dim strSQL As String

On Error GoTo vError

strSQL = "exec spCRDOperacionCargosAdd " & vOperacion & ",'" & vCodigo & "'," & vMonto & ""
Call ConectionExecute(strSQL)

Exit Sub
vError:



End Sub


Private Sub sbTramiteX(xCodigo As String, xCedula As String, xOperacion As Long)
Dim strSQL As String

On Error GoTo vError

strSQL = "exec spCrd_SGT_Tramite_Rapido " & xOperacion
Call ConectionExecute(strSQL)


Exit Sub

vError:

End Sub

Private Function fxEmite_Actual(pOperacion As Long) As String
Dim pResultado As String

pResultado = "ND"

With glogon
    .strSQL = "select isnull(EMITIR,'ND') as 'Emite' from reg_creditos where id_solicitud = " & pOperacion
    Call OpenRecordSet(.Recordset, .strSQL)
    
    pResultado = .Recordset!Emite
    
End With

fxEmite_Actual = pResultado

End Function


Private Sub sbGuardarSolicitud()
Dim strSQL As String, rs As New ADODB.Recordset
Dim xGarantiaFND As String, xGarantiaFNDContrato As Long
Dim vOperacionTemporal As Long, vActividad As String, vDestino As String, vCanal As String
Dim vProveedor As String, vEmiteTipo As String, vEmiteAnterior As String, vBaseCalculo As String
Dim vPromotor As String, vFormulario As String

'Fondos / Planes
If cboGarantia.ItemData(cboGarantia.ListIndex) = "Y" Then
  xGarantiaFND = cboFondo.ItemData(cboFondo.ListIndex)
  If cboFondoContrato.ListCount > 0 Then
      xGarantiaFNDContrato = cboFondoContrato.ItemData(cboFondoContrato.ListIndex)
  Else
      xGarantiaFNDContrato = 0
  End If

Else
  xGarantiaFND = ""
  xGarantiaFNDContrato = 0
End If


If IsNumeric(txtPromotorId.Text) Then
   vPromotor = txtPromotorId.Text
Else
   vPromotor = "Null"
End If


If IsNumeric(txtFormularioId.Text) Then
   vFormulario = txtFormularioId.Text
Else
   vFormulario = "Null"
End If


vDestino = cboDestino.ItemData(cboDestino.ListIndex)
vEmiteTipo = fxTipoDocumento(cboTipoDocumento.Text)
 

If vEmiteTipo = "CP" Then
    vProveedor = txtProveedorId.Text
Else
    vProveedor = "Null"
End If

vActividad = "Null"
If cboActividad.ListCount > 0 Then
   If cboActividad.Text <> "" Then
       vActividad = "'" & cboActividad.ItemData(cboActividad.ListIndex) & "'"
   End If
End If



vCanal = "Null"
If cboCanal.ListCount > 0 Then
   If cboCanal.Text <> "" Then
       vCanal = "'" & cboCanal.ItemData(cboCanal.ListIndex) & "'"
   End If
End If

strSQL = "select base_Calculo, Moneda from catalogo where codigo = '" & txtCodigo.Text & "'"
Call OpenRecordSet(rs, strSQL)
    vBaseCalculo = rs!Base_Calculo
rs.Close
 
 
Dim vFechaVence As String

vFechaVence = "Null"
If dtpVence.Visible Then
    vFechaVence = "'" & Format(dtpVence.Value, "yyyy-MM-dd") & "'"
End If
 
 
 
' spCrd_Tramite_Operacion_Recepcion_Add(@Operacion int, @Codigo varchar(10), @Destino varchar(10), @Garantia varchar(10), @EstadoSolicitud char(1)
'        , @Cedula varchar(20), @Monto dec(18,2), @Plazo int, @Tasa dec(7,4), @Cuota dec(16,2), @TasaPtsBono dec(7,4), @FSolicita datetime
'        , @Divisa varchar(10), @BaseCalculo varchar(10), @ComiteId int, @PriDeduc dec(7,1),  @Observacion varchar(3000)
'        , @OficinaPresenta varchar(10), @OficinaApoyo varchar(10), @OficinaTitular varchar(10), @Promotor_Id int = Null
'        , @BancoId int, @Cuenta_Bancaria varchar(50), @EmiteTipo varchar(10), @ProveedorId int
'        , @FndGarantia varchar(10), @FndContrato int
'        , @Fecha_Vence datetime = Null
'        , @I_Exp_Digital smallint = 0, @I_Pagare_Manual smallint = 0, @Formulario BigInt = 0
'        , @I_TrasladoSalario smallint = 0, @I_Deduce_Planilla smallint = 1
'        , @Actividad_Id varchar(10) = Null, @Canal_Id varchar(10) = Null
'        , @IVA_Mnt dec(16,2) = 0
'        , @Usuario varchar(30) )
strSQL = "exec spCrd_SGT_Recepcion " & Operacion.Operacion & ", '" & UCase(txtCodigo.Text) _
       & "', '" & vDestino & "', '" & cboGarantia.ItemData(cboGarantia.ListIndex) & "', '" & fxEstadoOperacion(cboEstado.Text) _
       & "', '" & Trim(txtCedula.Text) & "', " & CCur(txtMonto.Text) & ", " & Trim(txtPlazo) & ", " & CCur(txtTasa.Text) & ", " & CCur(txtCuota.Text) _
       & " ,  " & Operacion.TasaPtsBono & ", '" & Format(dtpFechaSolicitud.Value, "yyyy-mm-dd") & "', '" & txtDivisa.Text & "', '" & vBaseCalculo _
       & "',  " & cboComite.ItemData(cboComite.ListIndex) & ", Null, '" & Mid(txtObservaciones.Text, 1, 2000) _
       & "', '" & cboOficina.ItemData(cboOficina.ListIndex) & "', '" & GLOBALES.gOficinaApoyo & "', '" & GLOBALES.gOficinaTitular & "', " & vPromotor _
       & " ,  " & cboBanco.ItemData(cboBanco.ListIndex) & ", '" & IIf((Len(Trim(cboCuenta.ItemData(cboCuenta.ListIndex))) = 0), "0", Mid(Trim(cboCuenta.ItemData(cboCuenta.ListIndex)), 1, 40)) _
       & "', '" & vEmiteTipo & "', " & vProveedor _
       & ",  '" & xGarantiaFND & "', " & xGarantiaFNDContrato _
       & ",   " & vFechaVence _
       & ",   " & chkExpedienteDigital.Value & ", " & chkPagareManual.Value & ", " & vFormulario _
       & ",   " & chkTrasladoSalario.Value & ", " & chkDeducPlanilla.Value _
       & ",   " & vActividad & ", " & vCanal & ", 0 , '" & glogon.Usuario & "'"
Call OpenRecordSet(rs, strSQL)
 
txtOperacion.Text = rs!Operacion
 
 
If rs!Inicial = 1 Then
    Call Bitacora("Registra", "Recepción de la Operacion : " & txtOperacion.Text)
    MsgBox "Solicitud Registrada Satisfactoriamente...", vbInformation
Else
    Call Bitacora("Modifica", "Recepción de la Operacion : " & txtOperacion.Text)
    MsgBox "Solicitud Actualizada Satisfactoriamente...", vbInformation
End If
 
'If vEdita Then
'  Select Case Operacion.EstadoSolicitud
'    Case "R", "P"
'
'      chkPrimera.Value = IIf(fxPrimerCuota(vDestino), vbChecked, vbUnchecked)
'
'      strSQL = "update reg_creditos set cedula = '" & Trim(txtCedula) & "',codigo = '" & Trim(txtCodigo) & "',montosol=" & CCur(txtMonto.Text) _
'         & ",fechasol='" & Format(Me.dtpFechaSolicitud.Value, "yyyy/mm/dd") & "',estadosol='" _
'         & fxEstadoOperacion(cboEstado.Text) & "',id_comite=" & cboComite.ItemData(cboComite.ListIndex) & ",int=" _
'         & Trim(txtTasa) & ",interesv=" & Trim(txtTasa) & ",plazo=" _
'         & Trim(txtPlazo) & ",cuota=" & CCur(txtCuota) & ",garantia='" _
'         & cboGarantia.ItemData(cboGarantia.ListIndex) & "',observacion=" _
'         & IIf((Len(Trim(txtObservaciones)) = 0), "'NADA'", "'" & Mid(Trim(txtObservaciones), 1, 1000) & "'") _
'         & ",acta =null,estado =null,userrec='" & glogon.Usuario & "',cod_banco=" & cboBanco.ItemData(cboBanco.ListIndex) & ",cta_banco= '" _
'         & IIf((Len(Trim(cboCuenta.ItemData(cboCuenta.ListIndex))) = 0), "0", Mid(Trim(cboCuenta.ItemData(cboCuenta.ListIndex)), 1, 30)) _
'         & "',ind_deposito=" & IIf(UCase(cboTipoDocumento.Text) = "TRANSFERENCIA", 1, 0) _
'         & ",emitir='" & vEmiteTipo & "', cod_Proveedor = " & vProveedor _
'         & ",primer_cuota ='" & IIf((chkPrimera.Value = vbChecked), "S", "N") & "',premio=0,tdocumento='ND'" _
'         & ",cod_destino = '" & vDestino & "',cod_oficina_comision = '" & GLOBALES.gOficinaTitular & "',Tasa_Pts_Bono = " & Operacion.TasaPtsBono
'      'Nuevo Plan de Credito
'      strSQL = strSQL & ",montoApr = " & CCur(txtMonto.Text) & ", COD_DIVISA = '" & txtDivisa.Text & "', UserRes = '" & glogon.Usuario _
'             & "',FechaRes = '" & Format(Me.dtpFechaSolicitud.Value, "yyyy/mm/dd") _
'             & "',fecha_inicio_calculo = dbo.MyGetdate(), categoria_persona = dbo.fxCRDClasificacion('" & Trim(txtCedula) & "',dbo.MyGetdate())" _
'             & ",garantia_Fnd = '" & xGarantiaFND & "', GARANTIA_FND_CONTRATO = " & xGarantiaFNDContrato & ", cuota_poliza = "
'
'      If fxCobraTasaFormaliza(vDestino) And cboGarantia.ItemData(cboGarantia.ListIndex) <> "H" And Not fxCodigoConvenio(txtCodigo) Then
'         strSQL = strSQL & fxCuotaPolizaVida(CCur(txtMonto), txtCodigo)
'      Else
'         strSQL = strSQL & "0"
'      End If
'
'
'    Case "A" 'Anulada
'
'      strSQL = "update reg_creditos set codigo = '" & Trim(txtCodigo) & "'" _
'         & ",id_comite=" & cboComite.ItemData(cboComite.ListIndex) & "," _
'         & "garantia='" & cboGarantia.ItemData(cboGarantia.ListIndex) & "',observacion=" _
'         & IIf((Len(Trim(txtObservaciones)) = 0), "'NADA'", "'" & Mid(Trim(txtObservaciones), 1, 500) & "'") _
'         & ",estado =null,cod_banco=" & cboBanco.ItemData(cboBanco.ListIndex) & ",cta_banco= '" _
'         & IIf((Len(Trim(cboCuenta.ItemData(cboCuenta.ListIndex))) = 0), "0", Mid(Trim(cboCuenta.ItemData(cboCuenta.ListIndex)), 1, 30)) _
'         & "',ind_deposito=" & IIf(UCase(cboTipoDocumento.Text) = "TRANSFERENCIA", 1, 0) _
'         & ",primer_cuota ='" & IIf((chkPrimera.Value = vbChecked), "S", "N") & "',garantia_fnd = '" _
'         & xGarantiaFND & "',GARANTIA_FND_CONTRATO = " & xGarantiaFNDContrato
'
'    Case "F"
'  '    strSQL = "update reg_creditos set observacion=" _
'         & IIf((Len(Trim(txtObservaciones)) = 0), "'NADA'", "'" & Mid(Trim(txtObservaciones), 1, 254) & "'") _
'         & ",cod_banco=" & cboBanco.ItemData(cboBanco.ListIndex) & ",cta_banco=" _
'         & IIf((Len(Trim(cboCuenta.ItemData(cboCuenta.ListIndex))) = 0), "0", Mid(Trim(cboCuenta.ItemData(cboCuenta.ListIndex)), 1, 20)) _
'         & ",ind_deposito=" & IIf(UCase(cboTipoDocumento.Text) = "TRANSFERENCIA", 1, 0) _
'         & ",emitir='" & vEmiteTipo & "'"
'
'        If txtAutorizada.Text = "" Then
'          strSQL = "update reg_creditos set observacion=" _
'             & IIf((Len(Trim(txtObservaciones)) = 0), "'NADA'", "'" & Mid(Trim(txtObservaciones), 1, 1000) & "'")
'
'        Else
'          strSQL = "update reg_creditos set observacion= observacion"
'        End If
'
'       Select Case fxEmite_Actual(txtOperacion.Text)
'            Case "TE", "CK", "ND", "TS"
'                    Select Case fxTipoDocumento(cboTipoDocumento.Text)
'                      Case "TE", "CK", "ND", "TS"
'                         strSQL = strSQL & ",cod_banco=" & cboBanco.ItemData(cboBanco.ListIndex) & ",cta_banco= '" _
'                           & IIf((Len(Trim(cboCuenta.ItemData(cboCuenta.ListIndex))) = 0), "0", Mid(Trim(cboCuenta.ItemData(cboCuenta.ListIndex)), 1, 30)) _
'                           & "',ind_deposito=" & IIf(UCase(cboTipoDocumento.Text) = "TRANSFERENCIA", 1, 0) _
'                           & ",emitir='" & vEmiteTipo & "', cod_Proveedor = " & vProveedor
'                      Case Else
'                         'No Aplica para: RC, CP, ND, CD
'                    End Select
'            Case Else
'            'No Aplica para: RC, CP, ND, CD
'        End Select
'
'    Case "N", "D"
'      strSQL = "update reg_creditos set observacion=" _
'         & IIf((Len(Trim(txtObservaciones)) = 0), "'NADA'", "'" & Mid(Trim(txtObservaciones), 1, 1000) & "'")
'
'  End Select
'
'
'
'  If vActividad <> "" Then
'         strSQL = strSQL & ",cod_actividad = '" & vActividad & "'"
'  End If
'
'  If vCanal <> "" Then
'         strSQL = strSQL & ",Canal_Tipo = '" & vCanal & "'"
'  End If
'
'  If txtPromotorId.Text <> "" Then
'         strSQL = strSQL & ", id_Promotor = " & txtPromotorId.Text
'  End If
'
'  strSQL = strSQL & ", BASE_CALCULO = '" & vBaseCalculo & "', IND_DEDUCE_PLANILLA = '" & IIf((chkDeducPlanilla.Value = vbChecked), "S", "N") & "'"
'
'
'  If dtpVence.Visible Then
'      strSQL = strSQL & ", FECHA_VENCE = '" & Format(dtpVence.Value, "yyyy-MM-dd") & "'"
'  Else
'      strSQL = strSQL & ", FECHA_VENCE = Null"
'  End If
'
'  strSQL = strSQL & " Where id_solicitud = " & txtOperacion
'
'
''  strSQL = strSQL & " and TS =  convert(timestamp, '" & Operacion.TS & "')"
''
'  Dim pRows As Long
'
'  'Registra la Actualización
'  Call ConectionExecute(strSQL, 0, pRows)
''
''  If pRows = 0 Then
''    MsgBox "Esta Operación ha sido modificada por otro usuario, debe consultarla nuevamente para poderla modificar!", vbExclamation
''    Exit Sub
''  End If
'
'  'Tags de Seguimiento para Cambios de Estados
'  Select Case fxEstadoOperacion(cboEstado.Text)
'     Case "P"
'       If Operacion.EstadoSolicitud <> "P" Then
'              Call sbCrdOperacionTags(Operacion.Operacion, Operacion.Codigo, "S07", "", "Estado Anterior : " & fxEstadoOperacion(Operacion.EstadoSolicitud))
'       End If
'     Case "D"
'       If Operacion.EstadoSolicitud <> "D" Then
'              Call sbCrdOperacionTags(Operacion.Operacion, Operacion.Codigo, "S08", "", "Estado Anterior : " & fxEstadoOperacion(Operacion.EstadoSolicitud))
'       End If
'     Case "R"
'       If Operacion.EstadoSolicitud <> "R" Then
'           Call sbCrdOperacionTags(Operacion.Operacion, Operacion.Codigo, "S01", "", "Estado Anterior : " & fxEstadoOperacion(Operacion.EstadoSolicitud))
'       Else
'           Call sbCrdOperacionTags(Operacion.Operacion, Operacion.Codigo, "S01", "", "Modifica...: " & txtObservaciones.Text)
'       End If
'
'  End Select
'
'
'  Select Case Operacion.EstadoSolicitud
'    Case "R", "P", "A"
'      Call sbCargosAdicionales(Operacion.Operacion, txtCodigo, CCur(txtMonto))
'
'      'Aplica Regla de Refunndicion Automatica
'      strSQL = "exec spCrd_SGT_Politica_Refundicion_Aplica " & Operacion.Operacion & ",'M'"
'      Call ConectionExecute(strSQL)
'    Case Else
'      'No Actualiza Cargos
'  End Select
'
'
'  If Operacion.EstadoSolicitud = "R" Then
'    Call sbTramiteX(Operacion.Codigo, Operacion.Cedula, Operacion.Operacion)
'  End If
'
'
'  Call Bitacora("Registra", "Actualiza la Solicitud : " & Operacion.Operacion)
'  MsgBox "Solicitud Actualizada Satisfactoriamente...", vbInformation
'
'
'Else 'Inserta
'
'  chkPrimera.Value = IIf(fxPrimerCuota(vDestino), vbChecked, vbUnchecked)
'
'  strSQL = "insert into reg_creditos(cedula, codigo, cod_divisa, montosol, fechasol, estadosol, id_comite, base_Calculo" _
'         & ", int, interesv, plazo, cuota, garantia, observacion, estado, userrec, cod_banco, cta_banco" _
'         & ", ind_deposito, primer_cuota, premio, emitir, tdocumento, cod_destino, montoApr, UserRes, FechaRes" _
'         & ", fecha_inicio_calculo, categoria_persona, cuota_poliza, garantia_fnd, GARANTIA_FND_CONTRATO" _
'         & ", cod_oficina_r, cod_oficina_comision, cod_actividad, canal_Tipo, Tasa_Pts_Bono, Id_Promotor, IND_DEDUCE_PLANILLA, cod_Proveedor, Fecha_Vence)" _
'         & " values('" & Trim(txtCedula) & "', '" & Trim(txtCodigo) & "', '" & txtDivisa.Text & "', " & CCur(txtMonto.Text) _
'         & ", '" & Format(dtpFechaSolicitud.Value, "yyyy/mm/dd") & "','" _
'         & fxEstadoOperacion(cboEstado.Text) & "'," & cboComite.ItemData(cboComite.ListIndex) & ", '" & vBaseCalculo & "', " _
'         & Trim(txtTasa) & ", " & Trim(txtTasa) & ", " & Trim(txtPlazo) & ", " & CCur(txtCuota) & ", '" & cboGarantia.ItemData(cboGarantia.ListIndex) & "'," _
'         & IIf((Len(Trim(txtObservaciones)) = 0), "'NADA'", "'" & Mid(Trim(txtObservaciones), 1, 1000) & "'") _
'         & ",null,'" & glogon.Usuario & "'," & cboBanco.ItemData(cboBanco.ListIndex) & ",'" _
'         & IIf((Len(Trim(cboCuenta.ItemData(cboCuenta.ListIndex))) = 0), "0", Mid(Trim(cboCuenta.ItemData(cboCuenta.ListIndex)), 1, 34)) _
'         & "'," & IIf(UCase(cboTipoDocumento.Text) = "TRANSFERENCIA", 1, 0) & ",'" _
'         & IIf((chkPrimera.Value = vbChecked), "S", "N") & "',0,'" _
'         & vEmiteTipo & "','ND','" & vDestino & "',"
'
' 'Modificacion Nuevo Plan de Credito
' strSQL = strSQL & CCur(txtMonto.Text) & ",'" & glogon.Usuario & "','" _
'        & Format(Me.dtpFechaSolicitud.Value, "yyyy/mm/dd") & "',dbo.MyGetdate(),dbo.fxCRDClasificacion('" & txtCedula.Text & "',dbo.MyGetdate()),"
'
'      If fxCobraTasaFormaliza(vDestino) And cboGarantia.ItemData(cboGarantia.ListIndex) <> "H" And Not fxCodigoConvenio(txtCodigo) Then
'         strSQL = strSQL & fxCuotaPolizaVida(CCur(txtMonto), txtCodigo) & ",'" & xGarantiaFND & "'," & xGarantiaFNDContrato _
'                & ",'" & GLOBALES.gOficinaApoyo & "','" & GLOBALES.gOficinaTitular & "',"
'      Else
'         strSQL = strSQL & "0,'" & xGarantiaFND & "'," & xGarantiaFNDContrato & ",'" & GLOBALES.gOficinaTitular & "','" & GLOBALES.gOficinaApoyo & "',"
'      End If
'
'      If vActividad = "" Then
'         strSQL = strSQL & "Null,"
'      Else
'         strSQL = strSQL & "'" & vActividad & "',"
'      End If
'
'      If vCanal = "" Then
'         strSQL = strSQL & "Null," & Operacion.TasaPtsBono
'      Else
'         strSQL = strSQL & "'" & vCanal & "'," & Operacion.TasaPtsBono
'      End If
'
'      If txtPromotorId.Text = "" Then
'         strSQL = strSQL & ",Null"
'      Else
'         strSQL = strSQL & "," & txtPromotorId.Text
'      End If
'
'      strSQL = strSQL & ",'" & IIf((chkDeducPlanilla.Value = vbChecked), "S", "N") & "', " & vProveedor
'
'      If dtpVence.Visible Then
'        strSQL = strSQL & ", '" & Format(dtpVence.Value, "yyyy-mm-dd") & "')"
'      Else
'        strSQL = strSQL & ", Null )"
'
'      End If
'
'
'   'Verificar si existe la cuenta de ahorros y si no crearla
' Call ConectionExecute(strSQL)
'
' vOperacionTemporal = fxUltimaOperacion(txtCedula)
'
' Call sbCrdOperacionTags(vOperacionTemporal, txtCodigo.Text, "S01", "", txtObservaciones.Text)
'
' Call sbCargosAdicionales(vOperacionTemporal, txtCodigo, CCur(txtMonto))
'
' 'Aplica Regla de Refunndicion Automatica
' strSQL = "exec spCrd_SGT_Politica_Refundicion_Aplica " & vOperacionTemporal & ",'A'"
' Call ConectionExecute(strSQL)
'
'
' Call sbTramiteX(txtCodigo, txtCedula, vOperacionTemporal)
'
' txtOperacion.Text = vOperacionTemporal
'
'
' Call Bitacora("Registra", "Recepción de la Operacion : " & txtOperacion.Text)
' MsgBox "Solicitud Grabada Satisfactoriamente...", vbInformation
'
'End If 'Inserta


End Sub

Private Sub ActualizaCodigoOperacion()
Dim strSQL As String

strSQL = "update fiadores set codigo = '" & txtCodigo & "' where id_solicitud =" & Operacion.Operacion
Call ConectionExecute(strSQL)

strSQL = "update refundiciones set codigor = '" & txtCodigo & "' where id_solicitudr =" & Operacion.Operacion
Call ConectionExecute(strSQL)

strSQL = "update desembolsos set codigo = '" & txtCodigo & "' where id_solicitud =" & Operacion.Operacion
Call ConectionExecute(strSQL)

'strSQL = "update pra_principal set codigo = '" & txtCodigo & "' where id_solicitud =" & Operacion.Operacion
'Call ConectionExecute(strSQL)

End Sub


Private Sub txtMonto_LostFocus()

On Error GoTo vError

txtMonto.Text = Format(txtMonto.Text, "Standard")

vError:


End Sub

Private Sub sbHistorial()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem

On Error GoTo vError

tcMain.Item(2).Selected = True

lswHistorial.ListItems.Clear

strSQL = "select O.*,T.descripcion as Etiqueta" _
       & " from CRD_OPERACION_TAGS O inner join Crd_Tags T on O.Tag_codigo = T.Tag_Codigo" _
       & " where O.id_solicitud = " & txtOperacion.Text & " order by O.registro_fecha"
Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
 Set itmX = lswHistorial.ListItems.Add(, , rs!Registro_Fecha)
     itmX.SubItems(1) = rs!Registro_Usuario
     itmX.SubItems(2) = rs!Etiqueta
     itmX.SubItems(3) = rs!Notas
 rs.MoveNext
Loop
rs.Close

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
  
End Sub



Private Sub tlbPrincipal_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error Resume Next

Select Case Button.Key
 Case "nuevo"
  txtOperacion.Text = ""
  txtOperacion.Enabled = False
  Call sbLimpiaDatos
  tlbPrincipal.Buttons(1).Enabled = False
  tlbPrincipal.Buttons(2).Enabled = False
  tlbPrincipal.Buttons(3).Enabled = True
  tlbPrincipal.Buttons(4).Enabled = True
  fraOperacion.Enabled = True
  txtObservaciones.Locked = False
  
  vEdita = False
  
  
  txtCedula.SetFocus
'  Call sbCargaCombos
  
  
  
 Case "editar"
  If Operacion.Operacion > 0 Then 'And Operacion.Estado = "A" Then
      vEdita = True
      Call Edicion(1)
    
      'Si el Estado Esta en Recepcion o Resolucion puede Cambiar Todos Los Datos
      'Si Esta en Formalización Solo puede Cambiar la Salida
      tlbPrincipal.Buttons(1).Enabled = False
      tlbPrincipal.Buttons(2).Enabled = False
      tlbPrincipal.Buttons(3).Enabled = True
      tlbPrincipal.Buttons(4).Enabled = True
      txtOperacion.Enabled = False
      fraOperacion.Enabled = True
      txtObservaciones.Locked = False


      txtCedula.SetFocus

  End If
 
 Case "guardar"
  
  If fxVerificaRecepcion Then
    'Verificar si se cambio el codigo
    If Trim(txtCodigo) <> Operacion.Codigo Then Call ActualizaCodigoOperacion
    Call sbGuardarSolicitud
    Call Edicion(0)
    Call sbCargaOperacion
    txtOperacion.Enabled = True
    
    tlbPrincipal.Buttons(1).Enabled = True
    tlbPrincipal.Buttons(2).Enabled = True
    tlbPrincipal.Buttons(3).Enabled = False
    tlbPrincipal.Buttons(4).Enabled = False
    
    fraOperacion.Enabled = False
    txtObservaciones.Locked = True

    
    If vEdita = False Then
        tcMain.Item(0).Selected = True
        'Datos Personales
        Call sbTaskPanel_Accion(Id_TaskItem_DatosPersonales)
    End If
    
    If vEdita = False And cboGarantia.ItemData(cboGarantia.ListIndex) = "F" Then
        tcMain.Item(0).Selected = True
        'Fiadores
        'Call btnOpciones_Click(1)
        Call sbTaskPanel_Accion(Id_TaskItem_Garantia)
    
    End If
    
    If vEdita = False Then
        'Requisitos
        Call sbTaskPanel_Accion(Id_TaskItem_Requisitos)
    
    End If
    
    If Operacion.EstadoSolicitud = "P" Or Operacion.EstadoSolicitud = "D" Then
        'Siempre verifica las causas, por si esta en Pendiente o Denegada
         Call sbTaskPanel_Accion(Id_TaskItem_Causas)
    End If
  
  Else
    MsgBox vMensaje, vbCritical
  End If
 
 Case "deshacer"
    txtOperacion.Enabled = True
    tlbPrincipal.Buttons(1).Enabled = True
    tlbPrincipal.Buttons(2).Enabled = True
    tlbPrincipal.Buttons(3).Enabled = False
    tlbPrincipal.Buttons(4).Enabled = False
    fraOperacion.Enabled = False
    If txtOperacion <> "" Then Call sbCargaOperacion
    txtOperacion.SetFocus
 
 Case "ayuda"
        frmContenedor.CD.HelpContext = Me.HelpContextID
        frmContenedor.CD.ShowHelp
 
 Case "cerrar"
  Unload Me

End Select


End Sub



Private Sub tlbPrincipal_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
Me.MousePointer = vbHourglass

Select Case ButtonMenu.Key
 Case "RepActas"
   Call sbFormsCall("frmCR_SolCreacionAgenda", 1, , , False, Me)

 Case "RepPreAnalisis"
   Call sbFormsCall("frmCR_SolicitudesPreAnalisis", 1, , , False, Me)
 
 Case "RepGarantia"
   Call sbFormsCall("frmCR_GeneraGarantia", 1, , , False, Me)
 
 
 Case "repBoleta"
   If Operacion.EstadoSolicitud = "F" Or Operacion.EstadoSolicitud = "N" Then
     Call sbCrdSGTBoletaFormaliza(Operacion.Operacion)
   Else
     MsgBox "La Operación # " & Operacion.Operacion & " No se encuentra formalizada", vbInformation
   End If
   
 Case "Cheques"
   If Operacion.EstadoSolicitud = "F" Or Operacion.EstadoSolicitud = "N" Then
     Call sbCrdSGTBoletaCK(Operacion.Operacion)
   End If
 
 Case "RecRef"
   If Operacion.EstadoSolicitud = "F" Or Operacion.EstadoSolicitud = "N" Then
     Call sbCrdSGTReciboRefundicion(Operacion.Operacion)
   End If
   
 Case "Solicitud"
    Call sbCrdSGTBoletaSolicitud(Operacion.Operacion)
    
 Case "Requisitos"
    Call sbCrdSGTBoletaRequisitos(Operacion.Operacion)

 Case "Caratula"
    Call sbCrdSGTCaratulaCredito(Operacion.Operacion)
    
 Case "Autorizacion"
    Call sbCrdSGTAutorizacionDeduccion(Operacion.Operacion)
End Select


Me.MousePointer = vbDefault

End Sub

Private Sub txtCedula_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then Call sbBusqueda(1)
End Sub

Private Sub txtCedula_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then txtCodigo.SetFocus
End Sub

Private Sub sbDeductoras_Load(pInstitucion As Long)
Dim strSQL As String

strSQL = "select COD_DEDUCTORA AS 'IdX', DESCRIPCION AS 'ItmX'" _
       & " From vAFI_Deductoras" _
       & " Where cod_institucion = " & pInstitucion

Call sbCbo_Llena_New(cboDeductora, strSQL, False, True)

End Sub


Private Sub txtCedula_LostFocus()
Dim strSQL As String, rs As New ADODB.Recordset

strSQL = "select S.nombre, isnull(I.DEDUCCION_PLANILLA,0) as 'Deduccion' " _
       & ",S.cod_institucion, Ed.Cod_Institucion as 'DeductoraCod', Ed.Descripcion as 'DeductoraDesc'" _
       & " from Socios S inner join Instituciones I on S.cod_institucion = I.cod_Institucion" _
       & " left join Instituciones Ed on isnull(S.cod_deductora,S.cod_institucion) = Ed.cod_Institucion" _
       & " Where S.cedula = '" & txtCedula.Text & "'"
       
Call OpenRecordSet(rs, strSQL, 0)
If rs.EOF And rs.BOF Then
    txtNombre.Text = ""
    lblNombre.Caption = ""
    chkDeducPlanilla.Value = vbUnchecked
    chkDeducPlanilla.Enabled = False
Else
    txtNombre.Text = Trim(rs!Nombre)
    lblNombre.Caption = Trim(rs!Nombre)
    

    'Carga Deductoras por Institucion
    Call sbDeductoras_Load(rs!cod_institucion)
    Call sbCboAsignaDato(cboDeductora, rs!DeductoraDesc, True, rs!DeductoraCod)

    cboDeductora.Tag = CStr(rs!DeductoraCod)
    
    
    If rs!Deduccion = 0 Then
        chkDeducPlanilla.Value = vbUnchecked
        chkDeducPlanilla.Enabled = False

    Else
        chkDeducPlanilla.Value = vbChecked
        chkDeducPlanilla.Enabled = True
    End If

End If
rs.Close

 
 Call cboBanco_Click
End Sub

Private Sub DescribeCodigoComite()
Dim rs As New ADODB.Recordset, strSQL As String

On Error GoTo vError

strSQL = "select isnull(id_comite,0) as id_comite from catalogo where codigo='" & txtCodigo & "'"
Call OpenRecordSet(rs, strSQL)
If Not rs.EOF And Not rs.BOF Then
  cboComite.Text = fxDescribeComite(rs!id_Comite)
End If
rs.Close

vError:

End Sub

Private Sub txtCodigo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then Call sbBusqueda(2)
End Sub

Private Sub txtCodigo_KeyPress(KeyAscii As Integer)
On Error Resume Next

If KeyAscii = vbKeyReturn Then
  txtCodigo = UCase(txtCodigo)
  
  vPaso = True
        Call sbSTCargaCboGarantia(cboGarantia, txtCodigo)
        Call sbSTCargaCboEstado(cboEstado, "R")
        Call sbSTCargaCboDestinos(cboDestino, txtCodigo)

  vPaso = False
  
  
  
  'Garantia en Fondos
  If cboGarantia.ItemData(cboGarantia.ListIndex) = "Y" Then
    cboFondo.Visible = True
    cboFondoContrato.Visible = True
    lblFondoDisplay.Item(0).Visible = True
    lblFondoDisplay.Item(1).Visible = True
    Call cboFondo_Click
  Else
    cboFondo.Visible = False
    cboFondoContrato.Visible = False
    lblFondoDisplay.Item(0).Visible = False
    lblFondoDisplay.Item(1).Visible = False
  End If
  
  
  
  If fxCreditoExcedente(txtCodigo) Then
    chkPrimera.Value = vbUnchecked
    txtMonto.Text = fxExcedenteDisponible(txtCedula)
  End If
  
  Call cboGarantia_Click
  
  cboGarantia.SetFocus
  
End If

End Sub

Private Sub txtCodigo_LostFocus()
Dim strSQL As String, rs As New ADODB.Recordset

If txtCodigo.Text = "" Then Exit Sub

On Error GoTo vError

txtDivisa.Text = ""
txtDescripcion.Text = ""

strSQL = "select Cat.CODIGO, Cat.DESCRIPCION, Cat.MONEDA as 'COD_DIVISA', CAT.Base_Calculo" _
       & " , Cat.ID_COMITE, isnull(Com.DESCRIPCION,'') as 'COMITE_DESC'" _
       & " from CATALOGO Cat left join COMITES Com on Cat.ID_COMITE = Com.ID_COMITE" _
       & " where Cat.CODIGO = '" & txtCodigo.Text & "'"




mFrecuenciaPago = "M"

Call OpenRecordSet(rs, strSQL)
If Not rs.BOF And Not rs.BOF Then
    txtDescripcion.Text = rs!Descripcion & ""
    txtDivisa.Text = rs!cod_Divisa
    Call sbCboAsignaDato(cboComite, rs!Comite_Desc, True, rs!id_Comite)
    If rs!Base_Calculo = "06" Then 'Quincenal
        mFrecuenciaPago = "Q"
    End If
    
    
    If rs!Base_Calculo = "07" Then 'Interes Diario
        dtpVence.Visible = True
    Else
        dtpVence.Visible = False
    End If
    lblVence.Visible = dtpVence.Visible
    
End If
rs.Close

strSQL = "exec spCrd_SGT_Bancos '" & glogon.Usuario & "','" & txtDivisa.Text & "'"
Call sbCbo_Llena_New(cboBanco, strSQL, False, True)

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical


End Sub

Private Sub txtConCedula_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then
  txtConNombre = fxNombre(txtConCedula)
  Call sbConsultaX(txtConCedula)
End If

If KeyCode = vbKeyF4 Then
  gBusquedas.Convertir = "N"
  gBusquedas.Columna = "Cedula"
  gBusquedas.Orden = "Cedula"
  gBusquedas.Consulta = "select Cedula,Nombre from socios"
  frmBusquedas.Show vbModal
  txtConCedula = gBusquedas.Resultado
  txtConNombre = gBusquedas.Resultado2
  Call sbConsultaX(txtConCedula)
End If

End Sub

Private Sub txtConNombre_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then Call sbConsultaX(txtConCedula)

If KeyCode = vbKeyF4 Then
  gBusquedas.Convertir = "N"
  gBusquedas.Columna = "Nombre"
  gBusquedas.Orden = "Nombre"
  gBusquedas.Consulta = "select Cedula,Nombre from socios"
  frmBusquedas.Show vbModal
  txtConCedula = gBusquedas.Resultado
  txtConNombre = gBusquedas.Resultado2
  Call sbConsultaX(txtConCedula)
End If

End Sub


Private Sub txtDescripcion_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then Call sbBusqueda(4)
End Sub


Private Sub txtPromotorId_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtMonto.SetFocus
If KeyCode = vbKeyF4 Then
   gBusquedas.Convertir = "N"
   gBusquedas.Columna = "ID_PROMOTOR"
   gBusquedas.Orden = "ID_PROMOTOR"
   gBusquedas.Consulta = "select ID_PROMOTOR as 'Id.',Nombre from promotores"
   gBusquedas.Filtro = " and Estado = 1"
   frmBusquedas.Show vbModal
   txtPromotorId.Text = Trim(gBusquedas.Resultado)
   txtPromotorNombre.Text = Trim(gBusquedas.Resultado2)
End If
End Sub



Private Sub txtPromotorNombre_GotFocus()

If txtPromotorNombre.Text = "" Then
   gBusquedas.Convertir = "N"
   gBusquedas.Columna = "Nombre"
   gBusquedas.Orden = "Nombre"
   gBusquedas.Consulta = "select ID_PROMOTOR as 'Id.' ,Nombre from promotores"
   gBusquedas.Filtro = " and Estado = 1"
   frmBusquedas.Show vbModal
   txtPromotorId.Text = Trim(gBusquedas.Resultado)
   txtPromotorNombre.Text = Trim(gBusquedas.Resultado2)
End If

End Sub

Private Sub txtPromotorNombre_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtMonto.SetFocus

If KeyCode = vbKeyF4 Then
   gBusquedas.Convertir = "N"
   gBusquedas.Columna = "Nombre"
   gBusquedas.Orden = "Nombre"
   gBusquedas.Consulta = "select ID_PROMOTOR as 'Id.' ,Nombre from promotores"
   gBusquedas.Filtro = " and Estado = 1"
   frmBusquedas.Show vbModal
   txtPromotorId.Text = Trim(gBusquedas.Resultado)
   txtPromotorNombre.Text = Trim(gBusquedas.Resultado2)
End If
End Sub



Private Sub txtProveedorId_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtProveedorNombre.SetFocus

If KeyCode = vbKeyF4 Then
  gBusquedas.Col1Name = "Id. Proveedor"
  gBusquedas.Col2Name = "Id. Real"
  gBusquedas.Col3Name = "Nombre"
  gBusquedas.Convertir = "N"
  gBusquedas.Columna = "cod_proveedor"
  gBusquedas.Orden = "cod_proveedor"
  gBusquedas.Consulta = "select cod_proveedor,cedjur,descripcion from cxp_proveedores"
  gBusquedas.Filtro = ""
  frmBusquedas.Show vbModal

  txtProveedorId.Text = gBusquedas.Resultado
  txtProveedorNombre.Text = gBusquedas.Resultado3
End If

End Sub



Private Sub txtProveedorNombre_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtMonto.SetFocus

If KeyCode = vbKeyF4 Then
  gBusquedas.Col1Name = "Id. Proveedor"
  gBusquedas.Col2Name = "Id. Real"
  gBusquedas.Col3Name = "Nombre"
  gBusquedas.Convertir = "N"
  gBusquedas.Columna = "cod_proveedor"
  gBusquedas.Orden = "cod_proveedor"
  gBusquedas.Consulta = "select cod_proveedor,cedjur,descripcion from cxp_proveedores"
  gBusquedas.Filtro = ""
  frmBusquedas.Show vbModal

  txtProveedorId.Text = gBusquedas.Resultado
  txtProveedorNombre.Text = gBusquedas.Resultado3
End If

End Sub

Private Sub txtTasa_Change()
On Error GoTo vError
If CCur(IIf((txtTasa = ""), 0, txtTasa)) >= 0 And CCur(IIf((txtPlazo = ""), 0, txtPlazo)) > 0 _
    And CCur(IIf((txtMonto = ""), 0, txtMonto)) > 0 Then
  txtCuota = fxCalcula_Cuota(CCur(txtMonto), CCur(txtPlazo), CCur(txtTasa), mFrecuenciaPago)
End If
vError:

End Sub

Private Sub txtMonto_Change()
On Error GoTo vError
If CCur(IIf((txtTasa = ""), 0, txtTasa)) > 0 And CCur(IIf((txtPlazo = ""), 0, txtPlazo)) > 0 _
    And CCur(IIf((txtMonto = ""), 0, txtMonto)) > 0 Then
 txtCuota = fxCalcula_Cuota(CCur(txtMonto), CCur(txtPlazo), CCur(txtTasa), mFrecuenciaPago)
End If

vError:
End Sub

Private Sub txtTasa_KeyPress(KeyAscii As Integer)
 If KeyAscii = vbKeyReturn Then
  On Error Resume Next
    If CCur(IIf((txtTasa = ""), 0, txtTasa)) >= 0 And CCur(IIf((txtPlazo = ""), 0, txtPlazo)) > 0 _
        And CCur(IIf((txtMonto = ""), 0, txtMonto)) > 0 Then
      txtCuota = fxCalcula_Cuota(CCur(txtMonto), CCur(txtPlazo), CCur(txtTasa), mFrecuenciaPago)
    End If
   cboComite.SetFocus
End If
End Sub

Private Sub txtTasa_LostFocus()
On Error GoTo vError
If CCur(IIf((txtTasa = ""), 0, txtTasa)) >= 0 And CCur(IIf((txtPlazo = ""), 0, txtPlazo)) > 0 _
    And CCur(IIf((txtMonto = ""), 0, txtMonto)) > 0 Then
  txtCuota = fxCalcula_Cuota(CCur(txtMonto), CCur(txtPlazo), CCur(txtTasa), mFrecuenciaPago)
End If
vError:
End Sub

Private Sub txtNombre_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then Call sbBusqueda(3)
End Sub



Public Sub sbConsultaExterna(xOpTemp As Long)
 txtOperacion = xOpTemp
 If TimerX.Interval > 0 Then
    Call TimerX_Timer
 End If
 Call txtOperacion_KeyDown(vbKeyReturn, 0)
End Sub


Private Sub txtOperacion_Change()
 Call sbLimpiaDatos
  With tlbPrincipal.Buttons
   .Item(2).Enabled = False
   .Item(3).Enabled = False
   .Item(4).Enabled = False
 End With
End Sub

Private Sub txtOperacion_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then Call sbCargaOperacion
If KeyCode = vbKeyF4 Then Call sbBusqueda(0)
End Sub

Private Sub txtPlazo_Change()
On Error GoTo vError
If CCur(IIf((txtTasa = ""), 0, txtTasa)) >= 0 And CCur(IIf((txtPlazo = ""), 0, txtPlazo)) > 0 _
    And CCur(IIf((txtMonto = ""), 0, txtMonto)) > 0 Then
  txtCuota.Text = fxCalcula_Cuota(CCur(txtMonto), CCur(txtPlazo), CCur(txtTasa), mFrecuenciaPago)
End If

vError:
End Sub


Private Sub txtPlazo_KeyDown(KeyCode As Integer, Shift As Integer)
Dim x As Double

If KeyCode = vbKeyReturn And txtPlazo.Text <> "" Then
    
    If cboGarantia.ItemData(cboGarantia.ListIndex) = "Y" And cboFondo.ListCount > 0 Then
       If lblTasa.Tag = 0 Then
            x = fxCatalogoRangoPlz(txtCodigo, txtPlazo, cboDestino.ItemData(cboDestino.ListIndex))
            If x > 0 Then txtTasa.Text = x
       
       End If
    Else
       x = fxCatalogoRangoPlz(txtCodigo, txtPlazo, cboDestino.ItemData(cboDestino.ListIndex))
       If x > 0 Then txtTasa.Text = x - Operacion.TasaPtsBono
    End If

End If
 
End Sub

Private Sub txtPlazo_KeyPress(KeyAscii As Integer)
 If KeyAscii = vbKeyReturn Then txtTasa.SetFocus
End Sub

