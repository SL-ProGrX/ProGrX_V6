VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#20.3#0"; "Codejock.Controls.v20.3.0.ocx"
Begin VB.Form frmVivReportesGarantias 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Garantías Hipotecarias: Informes"
   ClientHeight    =   5295
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   9135
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5295
   ScaleWidth      =   9135
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin XtremeSuiteControls.TabControl tcMain 
      Height          =   2655
      Left            =   120
      TabIndex        =   6
      Top             =   1440
      Width           =   8895
      _Version        =   1310723
      _ExtentX        =   15684
      _ExtentY        =   4678
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
      Item(0).Caption =   "Tiempos de Trámites"
      Item(0).ControlCount=   4
      Item(0).Control(0)=   "Label6(2)"
      Item(0).Control(1)=   "Label6(1)"
      Item(0).Control(2)=   "dtpProdAcum"
      Item(0).Control(3)=   "btnProdAcum"
      Item(1).Caption =   "i Contacto"
      Item(1).ControlCount=   11
      Item(1).Control(0)=   "OptTramitesPendientes"
      Item(1).Control(1)=   "ChkTramitesPendientesTodos"
      Item(1).Control(2)=   "chkContactosDetallado"
      Item(1).Control(3)=   "OptMontoCreditos_Contacto"
      Item(1).Control(4)=   "OptDuracionT_Contacto"
      Item(1).Control(5)=   "cboTipo"
      Item(1).Control(6)=   "cboContacto"
      Item(1).Control(7)=   "Label13"
      Item(1).Control(8)=   "Label12"
      Item(1).Control(9)=   "Label8"
      Item(1).Control(10)=   "cboEmpresas"
      Item(2).Caption =   "i Zonas"
      Item(2).ControlCount=   5
      Item(2).Control(0)=   "chkZonasDetallado"
      Item(2).Control(1)=   "OptDuracionT_Zona"
      Item(2).Control(2)=   "OptMontoCreditos_Zona"
      Item(2).Control(3)=   "cboZonas"
      Item(2).Control(4)=   "Label9"
      Item(3).Caption =   "Desembolsos"
      Item(3).ControlCount=   7
      Item(3).Control(0)=   "optDesembolsosConcluidos"
      Item(3).Control(1)=   "OptAuxDesembolsos"
      Item(3).Control(2)=   "chkIncluirTodos"
      Item(3).Control(3)=   "OptDesembolsosDisponibles"
      Item(3).Control(4)=   "OptFDesembolso"
      Item(3).Control(5)=   "OptFFormalizacion"
      Item(3).Control(6)=   "Label1"
      Item(4).Caption =   "Giros Pendientes"
      Item(4).ControlCount=   6
      Item(4).Control(0)=   "cboTipoContacto2"
      Item(4).Control(1)=   "cboContactos2"
      Item(4).Control(2)=   "cboEstado"
      Item(4).Control(3)=   "Label7"
      Item(4).Control(4)=   "Label4"
      Item(4).Control(5)=   "Label3"
      Begin XtremeSuiteControls.CheckBox chkContactosDetallado 
         Height          =   255
         Left            =   -66880
         TabIndex        =   29
         Top             =   960
         Visible         =   0   'False
         Width           =   1695
         _Version        =   1310723
         _ExtentX        =   2990
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Detallado"
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
         UseVisualStyle  =   -1  'True
         Appearance      =   17
      End
      Begin XtremeSuiteControls.RadioButton OptDuracionT_Contacto 
         Height          =   495
         Left            =   -69760
         TabIndex        =   26
         Top             =   480
         Visible         =   0   'False
         Width           =   2175
         _Version        =   1310723
         _ExtentX        =   3836
         _ExtentY        =   873
         _StockProps     =   79
         Caption         =   "Duración de Trámites"
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
         UseVisualStyle  =   -1  'True
         Appearance      =   17
      End
      Begin VB.ComboBox cboEstado 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   -67360
         Style           =   2  'Dropdown List
         TabIndex        =   22
         Top             =   1920
         Visible         =   0   'False
         Width           =   4215
      End
      Begin VB.ComboBox cboContactos2 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   -67360
         Style           =   2  'Dropdown List
         TabIndex        =   21
         Top             =   720
         Visible         =   0   'False
         Width           =   4215
      End
      Begin VB.ComboBox cboTipoContacto2 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "frmVivReportesGarantias.frx":0000
         Left            =   -67360
         List            =   "frmVivReportesGarantias.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Top             =   1320
         Visible         =   0   'False
         Width           =   4215
      End
      Begin VB.ComboBox cboZonas 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   -68560
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   1680
         Visible         =   0   'False
         Width           =   4932
      End
      Begin VB.ComboBox cboEmpresas 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   -67840
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   1800
         Visible         =   0   'False
         Width           =   4215
      End
      Begin VB.ComboBox cboContacto 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   -67840
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   2280
         Visible         =   0   'False
         Width           =   4215
      End
      Begin VB.ComboBox cboTipo 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "frmVivReportesGarantias.frx":0004
         Left            =   -67840
         List            =   "frmVivReportesGarantias.frx":0006
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   1320
         Visible         =   0   'False
         Width           =   4215
      End
      Begin XtremeSuiteControls.DateTimePicker dtpProdAcum 
         Height          =   312
         Left            =   4560
         TabIndex        =   9
         Top             =   1440
         Width           =   1332
         _Version        =   1310723
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
      Begin XtremeSuiteControls.PushButton btnProdAcum 
         Height          =   612
         Left            =   6720
         TabIndex        =   10
         Top             =   1440
         Width           =   1572
         _Version        =   1310723
         _ExtentX        =   2773
         _ExtentY        =   1080
         _StockProps     =   79
         Caption         =   "Producto Acumulado"
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
      Begin XtremeSuiteControls.RadioButton OptMontoCreditos_Contacto 
         Height          =   495
         Left            =   -67120
         TabIndex        =   27
         Top             =   480
         Visible         =   0   'False
         Width           =   2175
         _Version        =   1310723
         _ExtentX        =   3836
         _ExtentY        =   873
         _StockProps     =   79
         Caption         =   "Monto de Créditos"
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
         UseVisualStyle  =   -1  'True
         Appearance      =   17
      End
      Begin XtremeSuiteControls.RadioButton OptTramitesPendientes 
         Height          =   495
         Left            =   -64480
         TabIndex        =   28
         Top             =   480
         Visible         =   0   'False
         Width           =   2175
         _Version        =   1310723
         _ExtentX        =   3836
         _ExtentY        =   873
         _StockProps     =   79
         Caption         =   "Trámites Pendientes"
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
         UseVisualStyle  =   -1  'True
         Appearance      =   17
      End
      Begin XtremeSuiteControls.CheckBox ChkTramitesPendientesTodos 
         Height          =   255
         Left            =   -64240
         TabIndex        =   30
         Top             =   960
         Visible         =   0   'False
         Width           =   1695
         _Version        =   1310723
         _ExtentX        =   2990
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Incluir Todos"
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
         UseVisualStyle  =   -1  'True
         Appearance      =   17
      End
      Begin XtremeSuiteControls.RadioButton OptDuracionT_Zona 
         Height          =   495
         Left            =   -68560
         TabIndex        =   31
         Top             =   480
         Visible         =   0   'False
         Width           =   2175
         _Version        =   1310723
         _ExtentX        =   3836
         _ExtentY        =   873
         _StockProps     =   79
         Caption         =   "Duración de Trámites"
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
         UseVisualStyle  =   -1  'True
         Appearance      =   17
      End
      Begin XtremeSuiteControls.RadioButton OptMontoCreditos_Zona 
         Height          =   495
         Left            =   -65920
         TabIndex        =   32
         Top             =   480
         Visible         =   0   'False
         Width           =   2175
         _Version        =   1310723
         _ExtentX        =   3836
         _ExtentY        =   873
         _StockProps     =   79
         Caption         =   "Monto de Créditos"
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
         UseVisualStyle  =   -1  'True
         Appearance      =   17
      End
      Begin XtremeSuiteControls.CheckBox chkZonasDetallado 
         Height          =   255
         Left            =   -65680
         TabIndex        =   33
         Top             =   960
         Visible         =   0   'False
         Width           =   1695
         _Version        =   1310723
         _ExtentX        =   2990
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Detallado"
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
         UseVisualStyle  =   -1  'True
         Appearance      =   17
      End
      Begin XtremeSuiteControls.RadioButton OptFDesembolso 
         Height          =   375
         Left            =   -67840
         TabIndex        =   34
         Top             =   480
         Visible         =   0   'False
         Width           =   4935
         _Version        =   1310723
         _ExtentX        =   8705
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Por Fecha Desembolso"
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
         UseVisualStyle  =   -1  'True
         Appearance      =   17
      End
      Begin XtremeSuiteControls.RadioButton OptFFormalizacion 
         Height          =   375
         Left            =   -67840
         TabIndex        =   35
         Top             =   840
         Visible         =   0   'False
         Width           =   4935
         _Version        =   1310723
         _ExtentX        =   8705
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Por Fecha Formalización"
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
         UseVisualStyle  =   -1  'True
         Appearance      =   17
      End
      Begin XtremeSuiteControls.RadioButton OptDesembolsosDisponibles 
         Height          =   375
         Left            =   -67840
         TabIndex        =   36
         Top             =   1200
         Visible         =   0   'False
         Width           =   4935
         _Version        =   1310723
         _ExtentX        =   8705
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Reporte Desembolsos Disponibles"
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
         UseVisualStyle  =   -1  'True
         Appearance      =   17
      End
      Begin XtremeSuiteControls.RadioButton OptAuxDesembolsos 
         Height          =   375
         Left            =   -67840
         TabIndex        =   37
         Top             =   1920
         Visible         =   0   'False
         Width           =   4935
         _Version        =   1310723
         _ExtentX        =   8705
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Auxiliar Desembolsos"
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
         UseVisualStyle  =   -1  'True
         Appearance      =   17
      End
      Begin XtremeSuiteControls.RadioButton optDesembolsosConcluidos 
         Height          =   375
         Left            =   -67840
         TabIndex        =   38
         Top             =   2280
         Visible         =   0   'False
         Width           =   4935
         _Version        =   1310723
         _ExtentX        =   8705
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Operaciones Desembolsos Concluidos"
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
         UseVisualStyle  =   -1  'True
         Appearance      =   17
      End
      Begin XtremeSuiteControls.CheckBox chkIncluirTodos 
         Height          =   255
         Left            =   -67600
         TabIndex        =   39
         Top             =   1560
         Visible         =   0   'False
         Width           =   3855
         _Version        =   1310723
         _ExtentX        =   6800
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Incluir Todos"
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
         UseVisualStyle  =   -1  'True
         Appearance      =   17
      End
      Begin VB.Label Label3 
         Caption         =   "Estado:"
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
         Left            =   -68560
         TabIndex        =   25
         Top             =   1920
         Visible         =   0   'False
         Width           =   1212
      End
      Begin VB.Label Label4 
         Caption         =   "Contacto:"
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
         Left            =   -68560
         TabIndex        =   24
         Top             =   720
         Visible         =   0   'False
         Width           =   1212
      End
      Begin VB.Label Label7 
         Caption         =   "Tipo Contacto:"
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
         Left            =   -68560
         TabIndex        =   23
         Top             =   1320
         Visible         =   0   'False
         Width           =   1212
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Informe de Desembolsos:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   252
         Left            =   -69640
         TabIndex        =   19
         Top             =   480
         Visible         =   0   'False
         Width           =   2532
      End
      Begin VB.Label Label9 
         Caption         =   "Zona:"
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
         Left            =   -69520
         TabIndex        =   18
         Top             =   1680
         Visible         =   0   'False
         Width           =   1092
      End
      Begin VB.Label Label8 
         Caption         =   "Contacto:"
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
         Left            =   -69160
         TabIndex        =   15
         Top             =   2280
         Visible         =   0   'False
         Width           =   1212
      End
      Begin VB.Label Label12 
         Caption         =   "Empresa:"
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
         Left            =   -69160
         TabIndex        =   14
         Top             =   1800
         Visible         =   0   'False
         Width           =   1212
      End
      Begin VB.Label Label13 
         Caption         =   "Tipo Contacto:"
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
         Left            =   -69160
         TabIndex        =   13
         Top             =   1320
         Visible         =   0   'False
         Width           =   1212
      End
      Begin VB.Label Label6 
         Caption         =   "Corte:"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   1
         Left            =   3960
         TabIndex        =   8
         Top             =   1440
         Width           =   612
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Producto Acumulado (Control de Desembolsos) de Créditos Hipotecarios:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   252
         Index           =   2
         Left            =   840
         TabIndex        =   7
         Top             =   960
         Width           =   5052
      End
   End
   Begin XtremeSuiteControls.GroupBox fraFechas 
      Height          =   1092
      Left            =   120
      TabIndex        =   1
      Top             =   4200
      Width           =   8892
      _Version        =   1310723
      _ExtentX        =   15684
      _ExtentY        =   1926
      _StockProps     =   79
      Caption         =   "Informe:"
      ForeColor       =   8388608
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
      Appearance      =   16
      BorderStyle     =   1
      Begin XtremeSuiteControls.PushButton cmdReporte 
         Height          =   612
         Left            =   6720
         TabIndex        =   2
         Top             =   240
         Width           =   1572
         _Version        =   1310723
         _ExtentX        =   2773
         _ExtentY        =   1080
         _StockProps     =   79
         Caption         =   "Reporte"
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
         Picture         =   "frmVivReportesGarantias.frx":0008
      End
      Begin XtremeSuiteControls.DateTimePicker dtpInicio 
         Height          =   312
         Left            =   3360
         TabIndex        =   3
         Top             =   360
         Width           =   1332
         _Version        =   1310723
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
      Begin XtremeSuiteControls.DateTimePicker dtpCorte 
         Height          =   312
         Left            =   4800
         TabIndex        =   4
         Top             =   360
         Width           =   1332
         _Version        =   1310723
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
      Begin VB.Label lblx02 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Fechas"
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
         Left            =   2160
         TabIndex        =   5
         Top             =   360
         Width           =   1212
      End
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Crédito Hipotecario con Avances de Obra: Informes y Estadisticas"
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
      Height          =   732
      Left            =   1920
      TabIndex        =   0
      Top             =   240
      Width           =   6972
   End
   Begin VB.Image imgBanner 
      Height          =   1332
      Left            =   0
      Top             =   0
      Width           =   12012
   End
End
Attribute VB_Name = "frmVivReportesGarantias"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim Reporte As String
    Dim vTitulo As String, vSubTitulo As String, FDesde As String, FCorte As String

Private Sub sbImprimir()


    Dim strSQL As String

    
    Me.MousePointer = vbHourglass
    
    Select Case tcMain.SelectedItem
    Case 0 ' Reporte Garantías por Fechas
    
        FDesde = Format(dtpInicio.Value, "dd/MM/yyyy")
        FCorte = Format(dtpCorte.Value, "dd/MM/yyyy")
        
        vTitulo = "Listado de Duración de Trámites"
        vSubTitulo = "Tiempo en Horas"
        
        Reporte = SIFGlobal.fxPathReportes("Credito_Hipotecario_DuracionGarantias.rpt")
    
        strSQL = strSQL & "cdate({VISTA_ViviendaDuracionGarantiasTotal.RegistroFechaVG}) in Date(" & Format(dtpInicio.Value, "yyyy,mm,dd")
        strSQL = strSQL & ") to Date (" & Format(dtpCorte.Value, "yyyy,mm,dd") & ")"
    
    
    Case 1 ' Reportes por Contacto
            
        FDesde = Format(dtpInicio.Value, "dd/MM/yyyy")
        FCorte = Format(dtpCorte.Value, "dd/MM/yyyy")
        
        If OptDuracionT_Contacto.Value = True Then '' Parámetros Reporte Duración Tramites por Contacto
            
            vTitulo = "Listado de Duración de Trámites"
            vSubTitulo = "Duración por Contacto en Horas"
            Reporte = SIFGlobal.fxPathReportes("Credito_Hipotecario_DuracionGarantiasContacto.rpt")
        
             strSQL = strSQL & "cdate({VISTA_ViviendaDuracionTramitesContacto.AsignacionFecha}) in Date(" & Format(dtpInicio.Value, "yyyy,mm,dd")
             strSQL = strSQL & ") to Date (" & Format(dtpCorte.Value, "yyyy,mm,dd") & ")"
        
            If Mid(Trim(cboTipo.Text), 1, 3) <> "[-T" Then
                If Len(strSQL) > 0 Then
                    strSQL = strSQL & " and "
                End If
                strSQL = strSQL & "{VISTA_ViviendaDuracionTramitesContacto.IdTipo} =  '" & Mid$(cboTipo.Text, 1, 1) & "'"
            End If
        
            If Mid(Trim(cboContacto.Text), 1, 3) <> "[-T" Then
                If Len(strSQL) > 0 Then
                    strSQL = strSQL & " and "
                End If
                strSQL = strSQL & "{VISTA_ViviendaDuracionTramitesContacto.IdContacto} =  " & cboContacto.ItemData(cboContacto.ListIndex)
            End If
            
            If Mid(Trim(cboEmpresas.Text), 1, 3) <> "[-T" Then
                If Len(strSQL) > 0 Then
                    strSQL = strSQL & " and "
                End If
                strSQL = strSQL & "{VISTA_ViviendaDuracionTramitesContacto.IdEmpresa} =  " & cboEmpresas.ItemData(cboEmpresas.ListIndex)
            End If
       End If
        
       If OptMontoCreditos_Contacto.Value = True Then  '' Parámetros Montos Tramites Vivienda por Contacto
            
            vTitulo = "Listado de Montos en Trámites"
            
            If chkContactosDetallado = 1 Then
                vSubTitulo = "Montos por Contacto Detallado"
                Reporte = SIFGlobal.fxPathReportes("Credito_Hipotecario_MontosContactoDet.rpt")
            Else
                vSubTitulo = "Montos por Contacto Resumido"
                Reporte = SIFGlobal.fxPathReportes("Credito_Hipotecario_MontosContactoRes.rpt")
            
            End If
            
            strSQL = strSQL & "cdate({VISTA_ViviendaMontosContactos.AsignacionFecha}) in Date(" & Format(dtpInicio.Value, "yyyy,mm,dd")
            strSQL = strSQL & ") to Date (" & Format(dtpCorte.Value, "yyyy,mm,dd") & ")"
        
            If Mid(Trim(cboTipo.Text), 1, 3) <> "[-T" Then
                If Len(strSQL) > 0 Then
                    strSQL = strSQL & " and "
                End If
                strSQL = strSQL & "{VISTA_ViviendaMontosContactos.TipoProfesional} =  '" & Mid$(cboTipo.Text, 1, 1) & "'"
            End If
        
            If Mid(Trim(cboContacto.Text), 1, 3) <> "[-T" Then
                If Len(strSQL) > 0 Then
                    strSQL = strSQL & " and "
                End If
                strSQL = strSQL & "{VISTA_ViviendaMontosContactos.IdContacto} =  " & cboContacto.ItemData(cboContacto.ListIndex)
            End If
            
            If Mid(Trim(cboEmpresas.Text), 1, 3) <> "[-T" Then
                If Len(strSQL) > 0 Then
                    strSQL = strSQL & " and "
                End If
                strSQL = strSQL & "{VISTA_ViviendaMontosContactos.IdEmpresa} =  " & cboEmpresas.ItemData(cboEmpresas.ListIndex)
            End If
       
        End If
        
       If OptTramitesPendientes.Value = True Then  '' Parámetros Vivienda Tramites pendientes por Contacto
            
            vTitulo = "Listado de Trámites Pendientes"
            vSubTitulo = "Trámites por Contacto"
            Reporte = SIFGlobal.fxPathReportes("Credito_Hipotecario_TramitesPendientesContacto.rpt")
            
            If ChkTramitesPendientesTodos = 0 Then
            
                 strSQL = strSQL & "cdate({VISTA_ViviendaTramitesPContacto.AsignacionFecha}) in Date(" & Format(dtpInicio.Value, "yyyy,mm,dd")
                 strSQL = strSQL & ") to Date (" & Format(dtpCorte.Value, "yyyy,mm,dd") & ")"
             
                 If Mid(Trim(cboTipo.Text), 1, 3) <> "[-T" Then
                     If Len(strSQL) > 0 Then
                         strSQL = strSQL & " and "
                     End If
                     strSQL = strSQL & "{VISTA_ViviendaTramitesPContacto.TipoProfesional} =  '" & Mid$(cboTipo.Text, 1, 1) & "'"
                 End If
             
                 If Mid(Trim(cboContacto.Text), 1, 3) <> "[-T" Then
                     If Len(strSQL) > 0 Then
                         strSQL = strSQL & " and "
                     End If
                     strSQL = strSQL & "{VISTA_ViviendaTramitesPContacto.IdContacto} =  " & cboContacto.ItemData(cboContacto.ListIndex)
                 End If
                 
                 If Mid(Trim(cboEmpresas.Text), 1, 3) <> "[-T" Then
                     If Len(strSQL) > 0 Then
                         strSQL = strSQL & " and "
                     End If
                     strSQL = strSQL & "{VISTA_ViviendaTramitesPContacto.IdEmpresa} =  " & cboEmpresas.ItemData(cboEmpresas.ListIndex)
                 End If
            End If
        End If
        
        
    Case 2 ' Reportes por Zona
    
        FDesde = Format(dtpInicio.Value, "dd/MM/yyyy")
        FCorte = Format(dtpCorte.Value, "dd/MM/yyyy")
    
        If OptDuracionT_Zona = True Then
            
            vTitulo = "Listado de Duración de Trámites"
            vSubTitulo = "Tiempo en Horas por Zona"
            Reporte = SIFGlobal.fxPathReportes("Credito_Hipotecario_DuracionGarantiasZona.rpt")
            
            strSQL = strSQL & "cdate({VISTA_ViviendaDuracionTramitesZona.RegistroFecha}) in Date(" & Format(dtpInicio.Value, "yyyy,mm,dd")
            strSQL = strSQL & ") to Date (" & Format(dtpCorte.Value, "yyyy,mm,dd") & ")"
        
            If Mid(Trim(cboZonas.Text), 1, 3) <> "[-T" Then
                If Len(strSQL) > 0 Then
                    strSQL = strSQL & " and "
                End If
                strSQL = strSQL & "{VISTA_ViviendaDuracionTramitesZona.IdZona} =  " & cboZonas.ItemData(cboZonas.ListIndex)
            End If
        
        Else
        
            vTitulo = "Listado de Monto de Trámites"
            
            If chkZonasDetallado = 1 Then
                vSubTitulo = "Montos por Zona Detallado"
                Reporte = SIFGlobal.fxPathReportes("Credito_Hipotecario_MontosZonaDet.rpt")
            Else
                vSubTitulo = "Montos por Zona Resumido"
                Reporte = SIFGlobal.fxPathReportes("Credito_Hipotecario_MontosZonaRes.rpt")
            End If
            
            strSQL = strSQL & "cdate({VISTA_ViviendaMontosZonas.RegistroFecha}) in Date(" & Format(dtpInicio.Value, "yyyy,mm,dd")
            strSQL = strSQL & ") to Date (" & Format(dtpCorte.Value, "yyyy,mm,dd") & ")"
        
            If Mid(Trim(cboZonas.Text), 1, 3) <> "[-T" Then
                If Len(strSQL) > 0 Then
                    strSQL = strSQL & " and "
                End If
                strSQL = strSQL & "{VISTA_ViviendaMontosZonas.IdZona} =  " & cboZonas.ItemData(cboZonas.ListIndex)
            End If
        
        End If
    
    Case 3 ' Reportes Desembolsos
    
        FDesde = Format(dtpInicio.Value, "dd/MM/yyyy")
        FCorte = Format(dtpCorte.Value, "dd/MM/yyyy")
        
        
        If OptFDesembolso.Value = True Then
        
            vTitulo = "Listado de Desembolsos"
            vSubTitulo = "Desembolsos por Fechas de Desembolso"
            Reporte = SIFGlobal.fxPathReportes("Credito_Hipotecario_Desembolsos.rpt")
            strSQL = strSQL & "cdate({VISTA_ViviendaDesembolsos.RegistroFecha}) in Date(" & Format(dtpInicio.Value, "yyyy,mm,dd")
            strSQL = strSQL & ") to Date (" & Format(dtpCorte.Value, "yyyy,mm,dd") & ")"
            
        End If
        
        If OptFFormalizacion.Value = True Then
            
            vTitulo = "Listado de Desembolso"
            vSubTitulo = "Desembolsos por Fechas de Formalización"
            Reporte = SIFGlobal.fxPathReportes("Credito_Hipotecario_Desembolsos.rpt")
            strSQL = strSQL & "cdate({VISTA_ViviendaDesembolsos.FechaForp}) in Date(" & Format(dtpInicio.Value, "yyyy,mm,dd")
            strSQL = strSQL & ") to Date (" & Format(dtpCorte.Value, "yyyy,mm,dd") & ")"
            
        End If
        
        If OptDesembolsosDisponibles.Value = True Then
        
            vTitulo = "Listado de Desembolsos"
            vSubTitulo = "Desembolsos Disponibles"
            Reporte = SIFGlobal.fxPathReportes("Credito_Hipotecario_DesembolsosDisponible.rpt")
            If chkIncluirTodos.Value = 0 Then
                strSQL = strSQL & "cdate({VISTA_ViviendaDesembolsoDisponible.FechaForp}) in Date(" & Format(dtpInicio.Value, "yyyy,mm,dd")
                strSQL = strSQL & ") to Date (" & Format(dtpCorte.Value, "yyyy,mm,dd") & ")"
            End If
        
        End If
        
        If OptAuxDesembolsos.Value = True Then
            vTitulo = "Listado Auxiliar de Desembolsos"
            vSubTitulo = "Operaciones con Pendientes al Corte"
            Reporte = SIFGlobal.fxPathReportes("Credito_Hipotecario_AuxiliarDesembolsos.rpt")
            
            Call sbImprimirRptAuxDesembolsos
            Exit Sub
        End If
        
        If optDesembolsosConcluidos.Value = True Then
            vTitulo = "Listado de Desembolsos Concluidos"
            vSubTitulo = "Operaciones Disponible en Cero del Periodo"
            Reporte = SIFGlobal.fxPathReportes("Credito_Hipotecario_DesembolsosConcluidos.rpt")
            
            Call sbImprimirRptDesembolsosConcluidos
            Exit Sub
        
        End If
            
    Case 4 ' Reportes Desembolsos Pendientes
    
        FDesde = Format(dtpInicio.Value, "dd/MM/yyyy")
        FCorte = Format(dtpCorte.Value, "dd/MM/yyyy")
        
        vTitulo = "Listado de Desembolsos"
        vSubTitulo = "Desembolsos Pendientes"
        Reporte = SIFGlobal.fxPathReportes("Credito_Hipotecario_DesembolsosPendientes.rpt")
        
        strSQL = strSQL & "cdate({VISTA_ViviendaDesembolsoPendientes.Fecha}) in Date(" & Format(dtpInicio.Value, "yyyy,mm,dd")
        strSQL = strSQL & ") to Date (" & Format(dtpCorte.Value, "yyyy,mm,dd") & ")"
           
        If Mid(Trim(cboTipoContacto2.Text), 1, 3) <> "[-T" Then
            If Len(strSQL) > 0 Then
                strSQL = strSQL & " and "
            End If
            strSQL = strSQL & "{VISTA_ViviendaDesembolsoPendientes.TipoProfesional} =  '" & Mid$(cboTipoContacto2.Text, 1, 1) & "'"
        End If
             
        If Mid(Trim(cboEstado.Text), 1, 3) <> "[-T" Then
            If Len(strSQL) > 0 Then
                strSQL = strSQL & " and "
            End If
            strSQL = strSQL & "{VISTA_ViviendaDesembolsoPendientes.Estado} =  '" & Mid$(cboEstado.Text, 1, 1) & "'"
        End If
             
        If Mid(Trim(cboContactos2.Text), 1, 3) <> "[-T" Then
            If Len(strSQL) > 0 Then
                strSQL = strSQL & " and "
            End If
            strSQL = strSQL & "{VISTA_ViviendaDesembolsoPendientes.IdContacto} =  " & cboContactos2.ItemData(cboContactos2.ListIndex)
        End If
        
        
    End Select
    
    With frmContenedor.Crt
    
      .Reset
      .WindowShowPrintSetupBtn = True
      .WindowShowRefreshBtn = True
      .WindowShowSearchBtn = True
      .WindowState = crptMaximized
      .WindowTitle = "Reportes Admin Créditos Hipotecarios"
      .Connect = glogon.ConectRPT
      
      .ReportFileName = Reporte
      
      .Formulas(0) = "titulo='" & vTitulo & "'"
      .Formulas(1) = "SubTitulo='" & vSubTitulo & "'"
      .Formulas(2) = "Fecha_Desde='" & FDesde & "'"
      .Formulas(3) = "Fecha_corte='" & FCorte & "'"
      .Formulas(4) = "Empresa= '" & GLOBALES.gstrNombreEmpresa & "'"
      .Formulas(5) = "Fecha='" & fxFechaServidor & "'"
      .Formulas(6) = "Usuario='" & glogon.Usuario & "'"
      .SelectionFormula = strSQL
    '  .PrintReport
      .Action = 1
    End With
    
    Me.MousePointer = vbDefault

End Sub

Private Sub sbImprimirRptAuxDesembolsos()
On Error GoTo error

    With frmContenedor.Crt
        .Reset
        .Connect = Empty
        
        .WindowShowGroupTree = True
        .WindowShowPrintSetupBtn = True
        .WindowShowRefreshBtn = True
        .WindowShowSearchBtn = True
        .WindowTitle = "Reportes Admin Créditos Hipotecarios"
        .Destination = crptToWindow
        .WindowState = crptMaximized
        .Connect = glogon.ConectRPT
        
        .ReportFileName = Reporte
        
        .Formulas(0) = "titulo='" & vTitulo & "'"
        .Formulas(1) = "SubTitulo='" & vSubTitulo & "'"
        .Formulas(2) = "Fecha_Desde='" & FDesde & "'"
        .Formulas(3) = "Fecha_corte='" & FCorte & "'"
        .Formulas(4) = "Empresa= '" & GLOBALES.gstrNombreEmpresa & "'"
        .Formulas(5) = "Fecha='" & fxFechaServidor & "'"
        .Formulas(6) = "Usuario='" & glogon.Usuario & "'"
        .StoredProcParam(0) = Format(dtpCorte.Value, "yyyy-MM-dd 23:59:59.000")

        .PrintReport
    End With
    frmContenedor.Crt.StoredProcParam(0) = Empty
    Exit Sub
error:
MsgBox ("Ocurrió un error al imprimir reporte de auxiliar de desembolsos. Error Nativo: " & Err.Description)
End Sub

Private Sub sbImprimirRptDesembolsosConcluidos()
On Error GoTo error

    With frmContenedor.Crt
        .Reset
        .Connect = Empty
        
        .WindowShowGroupTree = True
        .WindowShowPrintSetupBtn = True
        .WindowShowRefreshBtn = True
        .WindowShowSearchBtn = True
        .WindowTitle = "Reportes Admin Créditos Hipotecarios"
        .Destination = crptToWindow
        .WindowState = crptMaximized
        .Connect = glogon.ConectRPT
        
        .ReportFileName = Reporte
        
        .Formulas(0) = "titulo='" & vTitulo & "'"
        .Formulas(1) = "SubTitulo='" & vSubTitulo & "'"
        .Formulas(2) = "Fecha_Desde='" & FDesde & "'"
        .Formulas(3) = "Fecha_corte='" & FCorte & "'"
        .Formulas(4) = "Empresa= '" & GLOBALES.gstrNombreEmpresa & "'"
        .Formulas(5) = "Fecha='" & fxFechaServidor & "'"
        .Formulas(6) = "Usuario='" & glogon.Usuario & "'"
        .StoredProcParam(0) = Format(dtpInicio.Value, "yyyy-MM-dd 00:00:00.000")
        .StoredProcParam(1) = Format(dtpCorte.Value, "yyyy-MM-dd 23:59:59.000")

        .PrintReport
    End With
    frmContenedor.Crt.StoredProcParam(0) = Empty
    frmContenedor.Crt.StoredProcParam(1) = Empty
    Exit Sub
error:
MsgBox ("Ocurrió un error al imprimir reporte de desembolsos concluidos. Error Nativo: " & Err.Description)
End Sub




Private Sub btnProdAcum_Click()

On Error GoTo vError

    With frmContenedor.Crt
        .Reset
       
        .WindowShowGroupTree = True
        .WindowShowPrintSetupBtn = True
        .WindowShowRefreshBtn = True
        .WindowShowSearchBtn = True
        .WindowTitle = "Reportes Admin Créditos Hipotecarios"
        .Destination = crptToWindow
        .WindowState = crptMaximized
        
        .Connect = glogon.ConectRPT
        
        .ReportFileName = SIFGlobal.fxPathReportes("Credito_Hipotecario_AuxiliarProdAcum.rpt")
        
        .Formulas(1) = "fxFechaCorte='" & Format(dtpProdAcum.Value, "dd/mm/yyyy") & "'"
        .Formulas(2) = "fxEmpresa= '" & GLOBALES.gstrNombreEmpresa & "'"
        .Formulas(3) = "fxFecha='" & fxFechaServidor & "'"
        .Formulas(4) = "fxUsuario='" & glogon.Usuario & "'"
        .StoredProcParam(0) = Format(dtpProdAcum.Value, "yyyy-MM-dd 23:59:59.000")

        .PrintReport
    End With
    frmContenedor.Crt.StoredProcParam(0) = Empty

Exit Sub
    
vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub cmdReporte_Click()
    On Error GoTo error
    Me.MousePointer = vbHourglass
    
    Call sbImprimir
    
    Me.MousePointer = vbDefault
    Exit Sub

error:
     Me.MousePointer = vbDefault
     MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub Form_Load()
 
 vModulo = 3
 
Set Me.imgBanner.Picture = frmContenedor.imgBanner_Reportes.Picture

    Call sbgCntParametros
    
    dtpInicio.Value = fxFechaServidor
    dtpCorte.Value = dtpInicio.Value
    dtpProdAcum.Value = dtpInicio.Value
    
    tcMain.Item(0).Selected = True
    Call sbCargaEtiquetaFechas

End Sub

Private Sub sbCargaEtiquetaFechas()

    Select Case tcMain.SelectedItem
        Case 0
            fraFechas.Caption = "Fecha Registro Garantía"
         
        Case 1
            fraFechas.Caption = "Fecha Asignación Trámite"
            
        Case 2
            fraFechas.Caption = "Fecha Registro Garantía"

        Case 3
            If OptFDesembolso.Value = True Then
                fraFechas.Caption = "Fecha Registro Garantía"
            End If
            
            If OptFFormalizacion = True Then
                fraFechas.Caption = "Fecha Formalización Crédito"
            End If
            
            If OptDesembolsosDisponibles = True Then
                fraFechas.Caption = "Fecha Formalización Crédito"
            End If
            
            If OptAuxDesembolsos = True Then
                fraFechas.Caption = "Fecha de Corte"
            End If
        
            If optDesembolsosConcluidos = True Then
                fraFechas.Caption = "Fecha del Último desembolso "
            End If
        
        Case Else
    End Select


End Sub


Private Sub sbCargaCbo(cbo As ComboBox, vTipo As String, Optional vFiltro As String = "")
    Dim strSQL As String, rs As New ADODB.Recordset
    On Error GoTo vError
    
    
    Select Case UCase(vTipo)
        Case "CONTACTOS"
            strSQL = "select IdContacto as xLlave, isnull(Nombre,'') as xDesc from ViviendaContactos where TipoContacto <> 'E'"
        Case "EMPRESAS"
            strSQL = "select IdEmpresa as xLlave, isnull(Nombre,'') as xDesc from ViviendaContactos where TipoContacto = 'E'"
        Case "ZONAS" 'Zonas
            strSQL = "select IdZona as xLlave, isnull(Descripcion,'') as xDesc from viviendaZonas"
       Case Else
    End Select
    
    strSQL = strSQL & vFiltro
    cbo.Clear
    cbo.AddItem Trim("[-Todos-]")
    If execSql(strSQL) Then
     Do While Not glogon.Recordset.EOF
        cbo.AddItem Trim(glogon.Recordset!xDesc)
        cbo.ItemData(cbo.NewIndex) = glogon.Recordset!xLlave
        glogon.Recordset.MoveNext
      Loop
      glogon.Recordset.MoveFirst
      cbo.Text = Trim(glogon.Recordset!xDesc)
      If glogon.Recordset.State <> 0 Then
        glogon.Recordset.Close
      End If
      cbo.ListIndex = 0
      
    End If
    
    Exit Sub
vError:
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub OptAuxDesembolsos_Click()
    Call sbCargaEtiquetaFechas
    dtpInicio.Enabled = False
End Sub

Private Sub optDesembolsosConcluidos_Click()
    Call sbCargaEtiquetaFechas
    dtpInicio.Enabled = True
End Sub

Private Sub OptDesembolsosDisponibles_Click()
    sbCargaEtiquetaFechas
    dtpInicio.Enabled = True
End Sub

Private Sub OptFDesembolso_Click()
    sbCargaEtiquetaFechas
    dtpInicio.Enabled = True
End Sub

Private Sub OptFFormalizacion_Click()
    sbCargaEtiquetaFechas
    dtpInicio.Enabled = True
End Sub

Private Sub tcMain_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)

dtpInicio.Enabled = True

Select Case Item.Index
    Case 0 'Tiempos
    
    
        
    Case 1 'Contacto
    
        OptDuracionT_Contacto.Value = True
        
        cboTipo.Clear
        cboTipo.AddItem Trim("[-Todos-]")
        cboTipo.AddItem "Abogado"
        cboTipo.AddItem "Ingeniero"
        cboTipo.ListIndex = 0
        
        Call sbCargaCbo(cboContacto, "Contactos")
        Call sbCargaCbo(cboEmpresas, "Empresas")
        
    Case 2 'Zonas
    
        OptDuracionT_Zona.Value = True
    
        Call sbCargaCbo(cboZonas, "zonas")
    
    Case 3 'Desembolsos
    
    
        OptFDesembolso.Value = True
    
    Case 4 'Giros Pendientes
                    
        cboEstado.Clear
        cboEstado.AddItem Trim("[-Todos-]")
        cboEstado.AddItem "Pendiente"
        cboEstado.AddItem "Girado"
        cboEstado.ListIndex = 0
        
        cboTipoContacto2.Clear
        cboTipoContacto2.AddItem Trim("[-Todos-]")
        cboTipoContacto2.AddItem "Abogado"
        cboTipoContacto2.AddItem "Ingeniero"
        cboTipoContacto2.ListIndex = 0
        
        Call sbCargaCbo(cboContactos2, "Contactos")

End Select
    
Call sbCargaEtiquetaFechas

End Sub
