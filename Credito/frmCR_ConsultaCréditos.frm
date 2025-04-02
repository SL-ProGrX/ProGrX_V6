VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.Controls.v24.0.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.ShortcutBar.v24.0.0.ocx"
Begin VB.Form frmCR_ConsultaCreditos 
   Caption         =   "Consulta Integrada de la Persona"
   ClientHeight    =   8220
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13815
   HelpContextID   =   3010
   Icon            =   "frmCR_ConsultaCréditos.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8220
   ScaleWidth      =   13815
   Begin XtremeSuiteControls.PushButton isButton1 
      Height          =   375
      Index           =   6
      Left            =   10560
      TabIndex        =   119
      Top             =   2040
      Width           =   975
      _Version        =   1572864
      _ExtentX        =   1720
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Estado"
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
   End
   Begin XtremeSuiteControls.PushButton btnEC_Mail 
      Height          =   375
      Left            =   11520
      TabIndex        =   181
      ToolTipText     =   "Enviar Estado de Cuenta al Correo"
      Top             =   2040
      Width           =   495
      _Version        =   1572864
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
      Appearance      =   6
      Picture         =   "frmCR_ConsultaCréditos.frx":08CA
   End
   Begin XtremeSuiteControls.PushButton btnExport 
      Height          =   375
      Left            =   13200
      TabIndex        =   171
      ToolTipText     =   "Exportar a Excel"
      Top             =   2040
      Width           =   495
      _Version        =   1572864
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
      Appearance      =   6
      Picture         =   "frmCR_ConsultaCréditos.frx":093E
   End
   Begin XtremeSuiteControls.PushButton btnSoS 
      Height          =   330
      Left            =   11880
      TabIndex        =   163
      Top             =   105
      Visible         =   0   'False
      Width           =   735
      _Version        =   1572864
      _ExtentX        =   1296
      _ExtentY        =   582
      _StockProps     =   79
      Caption         =   "SoS"
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
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.PushButton isButton1 
      Height          =   375
      Index           =   9
      Left            =   4920
      TabIndex        =   141
      Top             =   2040
      Width           =   1215
      _Version        =   1572864
      _ExtentX        =   2143
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Beneficios"
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
   End
   Begin XtremeSuiteControls.PushButton isButton1 
      Height          =   375
      Index           =   10
      Left            =   6120
      TabIndex        =   142
      Top             =   2040
      Width           =   1215
      _Version        =   1572864
      _ExtentX        =   2143
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Renuncias"
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
   End
   Begin XtremeSuiteControls.PushButton isButton1 
      Height          =   375
      Index           =   8
      Left            =   8640
      TabIndex        =   138
      Top             =   2040
      Width           =   855
      _Version        =   1572864
      _ExtentX        =   1508
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Correo"
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
   End
   Begin XtremeSuiteControls.PushButton isButton1 
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   113
      Top             =   2040
      Width           =   1215
      _Version        =   1572864
      _ExtentX        =   2143
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Créditos"
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
   End
   Begin XtremeSuiteControls.PushButton isButton1 
      Height          =   375
      Index           =   1
      Left            =   1320
      TabIndex        =   114
      Top             =   2040
      Width           =   1215
      _Version        =   1572864
      _ExtentX        =   2143
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Cobros"
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
   End
   Begin XtremeSuiteControls.PushButton isButton1 
      Height          =   375
      Index           =   2
      Left            =   2520
      TabIndex        =   115
      Top             =   2040
      Width           =   1215
      _Version        =   1572864
      _ExtentX        =   2143
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Ahorros"
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
   End
   Begin XtremeSuiteControls.PushButton isButton1 
      Height          =   375
      Index           =   3
      Left            =   3720
      TabIndex        =   116
      Top             =   2040
      Width           =   1215
      _Version        =   1572864
      _ExtentX        =   2143
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Patrimonio"
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
   End
   Begin XtremeSuiteControls.PushButton isButton1 
      Height          =   375
      Index           =   4
      Left            =   7560
      TabIndex        =   117
      Top             =   2040
      Width           =   1095
      _Version        =   1572864
      _ExtentX        =   1931
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Mensajes"
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
   End
   Begin XtremeSuiteControls.PushButton isButton1 
      Height          =   375
      Index           =   5
      Left            =   9480
      TabIndex        =   118
      Top             =   2040
      Width           =   855
      _Version        =   1572864
      _ExtentX        =   1508
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Info"
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
   End
   Begin XtremeSuiteControls.PushButton isButton1 
      Height          =   375
      Index           =   7
      Left            =   12240
      TabIndex        =   120
      Top             =   2040
      Width           =   975
      _Version        =   1572864
      _ExtentX        =   1720
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Aut/C.I."
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
   End
   Begin VB.Frame fraConsentimiento 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2895
      Left            =   4440
      TabIndex        =   7
      Top             =   2640
      Visible         =   0   'False
      Width           =   6015
      Begin XtremeSuiteControls.PushButton btnConsentimiento 
         Height          =   612
         Index           =   0
         Left            =   3000
         TabIndex        =   108
         Top             =   2280
         Width           =   1572
         _Version        =   1572864
         _ExtentX        =   2773
         _ExtentY        =   1080
         _StockProps     =   79
         Caption         =   "Aprobar"
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
         Picture         =   "frmCR_ConsultaCréditos.frx":0AA8
      End
      Begin XtremeSuiteControls.FlatEdit txtConsentimientoUsuario 
         Height          =   312
         Left            =   2040
         TabIndex        =   105
         Top             =   1200
         Width           =   2532
         _Version        =   1572864
         _ExtentX        =   4466
         _ExtentY        =   556
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
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtConsentimientoFecha 
         Height          =   312
         Left            =   2040
         TabIndex        =   106
         Top             =   1560
         Width           =   2532
         _Version        =   1572864
         _ExtentX        =   4466
         _ExtentY        =   556
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
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.PushButton btnConsentimiento 
         Height          =   612
         Index           =   1
         Left            =   4560
         TabIndex        =   109
         Top             =   2280
         Width           =   732
         _Version        =   1572864
         _ExtentX        =   1291
         _ExtentY        =   1080
         _StockProps     =   79
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
         Picture         =   "frmCR_ConsultaCréditos.frx":1280
      End
      Begin XtremeSuiteControls.PushButton btnConsentimiento 
         Height          =   612
         Index           =   2
         Left            =   5280
         TabIndex        =   110
         Top             =   2280
         Width           =   732
         _Version        =   1572864
         _ExtentX        =   1291
         _ExtentY        =   1080
         _StockProps     =   79
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
         Picture         =   "frmCR_ConsultaCréditos.frx":1A3C
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha .:"
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
         Height          =   315
         Index           =   5
         Left            =   360
         TabIndex        =   10
         Top             =   1560
         Width           =   1335
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Usuario .:"
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
         Height          =   315
         Index           =   4
         Left            =   360
         TabIndex        =   9
         Top             =   1200
         Width           =   1335
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Aprobación del consentimiento de uso de información personal de contacto para productos y servicios."
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
         Height          =   615
         Left            =   240
         TabIndex        =   8
         Top             =   360
         Width           =   5535
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
         Height          =   2052
         Left            =   0
         TabIndex        =   107
         Top             =   120
         Width           =   6012
         _Version        =   1572864
         _ExtentX        =   10604
         _ExtentY        =   3619
         _StockProps     =   14
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin XtremeSuiteControls.PushButton btnSaldosFavor 
      Height          =   340
      Left            =   8892
      TabIndex        =   20
      ToolTipText     =   "Saldos a favor en Cajas?"
      Top             =   100
      Width           =   1800
      _Version        =   1572864
      _ExtentX        =   3175
      _ExtentY        =   600
      _StockProps     =   79
      Caption         =   "0.00"
      BackColor       =   16777215
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
      TextAlignment   =   5
      Appearance      =   17
   End
   Begin VB.Timer TimerX 
      Interval        =   20
      Left            =   12120
      Top             =   1320
   End
   Begin MSComctlLib.StatusBar statusBar 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   0
      Top             =   7905
      Width           =   13815
      _ExtentX        =   24368
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   7
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Bevel           =   0
            Object.Width           =   1834
            MinWidth        =   1834
            TextSave        =   "7/10/2024"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Bevel           =   0
            Object.Width           =   1658
            MinWidth        =   1658
            TextSave        =   "09:40:a. m."
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Bevel           =   0
            Object.Width           =   1129
            MinWidth        =   1129
            TextSave        =   "NUM"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   2892
            MinWidth        =   2892
            Object.ToolTipText     =   "Usuario"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   2541
            MinWidth        =   2541
            Object.ToolTipText     =   "Intereses Pendientes"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   3070
            MinWidth        =   3070
            Object.ToolTipText     =   "Cancelación"
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.ToolTipText     =   "Cargo x Anticipo"
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
   Begin MSComctlLib.ImageList imgSemaforos 
      Left            =   12000
      Top             =   720
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   16
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_ConsultaCréditos.frx":2209
            Key             =   "verde"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_ConsultaCréditos.frx":2825
            Key             =   "amarillo"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_ConsultaCréditos.frx":2E43
            Key             =   "rojo"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_ConsultaCréditos.frx":352A
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_ConsultaCréditos.frx":3DFB
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_ConsultaCréditos.frx":4522
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_ConsultaCréditos.frx":4B46
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_ConsultaCréditos.frx":53F3
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_ConsultaCréditos.frx":5AFF
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_ConsultaCréditos.frx":6206
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_ConsultaCréditos.frx":6906
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_ConsultaCréditos.frx":6F22
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_ConsultaCréditos.frx":7653
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_ConsultaCréditos.frx":7D50
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_ConsultaCréditos.frx":8457
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_ConsultaCréditos.frx":8B6D
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin XtremeSuiteControls.FlatEdit txtNombre 
      Height          =   315
      Left            =   2160
      TabIndex        =   17
      Top             =   120
      Width           =   5892
      _Version        =   1572864
      _ExtentX        =   10393
      _ExtentY        =   556
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
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtRate 
      Height          =   315
      Left            =   8040
      TabIndex        =   18
      ToolTipText     =   "Indice de Riesgo (Deudas vrs Ahorros) (+ 6 Advertencia)"
      Top             =   120
      Width           =   852
      _Version        =   1572864
      _ExtentX        =   1503
      _ExtentY        =   556
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
   Begin XtremeSuiteControls.FlatEdit txtCedula 
      Height          =   315
      Left            =   120
      TabIndex        =   19
      Top             =   120
      Width           =   2052
      _Version        =   1572864
      _ExtentX        =   3619
      _ExtentY        =   556
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
   Begin XtremeSuiteControls.TabControl tcMain 
      Height          =   5535
      Left            =   0
      TabIndex        =   24
      Top             =   2400
      Width           =   13215
      _Version        =   1572864
      _ExtentX        =   23310
      _ExtentY        =   9763
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
      Appearance      =   4
      Color           =   32
      PaintManager.Position=   2
      ItemCount       =   10
      SelectedItem    =   2
      Item(0).Caption =   "Credito"
      Item(0).ControlCount=   20
      Item(0).Control(0)=   "vgCreditos"
      Item(0).Control(1)=   "txtTotalSaldo"
      Item(0).Control(2)=   "txtTotalMonto"
      Item(0).Control(3)=   "txtTotalCuota"
      Item(0).Control(4)=   "dtpCorte"
      Item(0).Control(5)=   "btnCredito(0)"
      Item(0).Control(6)=   "btnCredito(1)"
      Item(0).Control(7)=   "btnCredito(2)"
      Item(0).Control(8)=   "btnCredito(3)"
      Item(0).Control(9)=   "btnCredito(4)"
      Item(0).Control(10)=   "btnCredito(5)"
      Item(0).Control(11)=   "btnCredito(6)"
      Item(0).Control(12)=   "btnCredito(7)"
      Item(0).Control(13)=   "btnCredito(8)"
      Item(0).Control(14)=   "Label1(2)"
      Item(0).Control(15)=   "Label1(1)"
      Item(0).Control(16)=   "Label1(0)"
      Item(0).Control(17)=   "Label1(3)"
      Item(0).Control(18)=   "btnConstancia"
      Item(0).Control(19)=   "gbNotas"
      Item(1).Caption =   "Cobros"
      Item(1).ControlCount=   6
      Item(1).Control(0)=   "vgCobro"
      Item(1).Control(1)=   "isButtonCb(0)"
      Item(1).Control(2)=   "isButtonCb(1)"
      Item(1).Control(3)=   "isButtonCb(2)"
      Item(1).Control(4)=   "rbNotificaEmail(0)"
      Item(1).Control(5)=   "rbNotificaEmail(1)"
      Item(2).Caption =   "Ahorros"
      Item(2).ControlCount=   8
      Item(2).Control(0)=   "lswFND"
      Item(2).Control(1)=   "gbFndContrato"
      Item(2).Control(2)=   "btnFondos_List(0)"
      Item(2).Control(3)=   "btnFondos_List(1)"
      Item(2).Control(4)=   "btnFondos_List(2)"
      Item(2).Control(5)=   "lswFnd_List"
      Item(2).Control(6)=   "btnFondos_Export"
      Item(2).Control(7)=   "btnFondos_List(3)"
      Item(3).Caption =   "Patrimonio"
      Item(3).ControlCount=   29
      Item(3).Control(0)=   "txtPatrimonio"
      Item(3).Control(1)=   "txtCapitalizacion"
      Item(3).Control(2)=   "txtAporte"
      Item(3).Control(3)=   "txtAhorro"
      Item(3).Control(4)=   "txtPAT_Disponible"
      Item(3).Control(5)=   "txtPAT_Saldos"
      Item(3).Control(6)=   "txtPAT_Giro"
      Item(3).Control(7)=   "vgPatrimonio"
      Item(3).Control(8)=   "cboPAT_TipoSaldo"
      Item(3).Control(9)=   "Label4(3)"
      Item(3).Control(10)=   "Label5(0)"
      Item(3).Control(11)=   "Label4(0)"
      Item(3).Control(12)=   "Label3(0)"
      Item(3).Control(13)=   "lblFechaAhorro"
      Item(3).Control(14)=   "lblFechaAporte"
      Item(3).Control(15)=   "lblCapitalizado"
      Item(3).Control(16)=   "Label4(1)"
      Item(3).Control(17)=   "Label4(4)"
      Item(3).Control(18)=   "lblPAT_Saldo"
      Item(3).Control(19)=   "Label4(6)"
      Item(3).Control(20)=   "btnConstanciaAportes"
      Item(3).Control(21)=   "txtCustodia"
      Item(3).Control(22)=   "lblFechaCustodia"
      Item(3).Control(23)=   "txtPat_Divisa"
      Item(3).Control(24)=   "Label4(2)"
      Item(3).Control(25)=   "cboPat_Garantia"
      Item(3).Control(26)=   "btnPatrimonioConsulta"
      Item(3).Control(27)=   "txtPAT_AporteCobro"
      Item(3).Control(28)=   "Label4(5)"
      Item(4).Caption =   "Mensajes"
      Item(4).ControlCount=   7
      Item(4).Control(0)=   "fraMsj"
      Item(4).Control(1)=   "cboMsj"
      Item(4).Control(2)=   "vGrid"
      Item(4).Control(3)=   "imgMsjResuelve"
      Item(4).Control(4)=   "Label6"
      Item(4).Control(5)=   "imgMsjNuevo"
      Item(4).Control(6)=   "imgBorraMsj"
      Item(5).Caption =   "Info"
      Item(5).ControlCount=   26
      Item(5).Control(0)=   "lswDP"
      Item(5).Control(1)=   "btnInfo(0)"
      Item(5).Control(2)=   "btnInfo(1)"
      Item(5).Control(3)=   "btnInfo(2)"
      Item(5).Control(4)=   "btnInfo(4)"
      Item(5).Control(5)=   "btnInfo(5)"
      Item(5).Control(6)=   "btnInfo(6)"
      Item(5).Control(7)=   "btnInfo(8)"
      Item(5).Control(8)=   "btnInfo(9)"
      Item(5).Control(9)=   "btnInfo(10)"
      Item(5).Control(10)=   "btnInfo(3)"
      Item(5).Control(11)=   "btnInfo(7)"
      Item(5).Control(12)=   "btnInfoTrigger(0)"
      Item(5).Control(13)=   "btnInfoTrigger(1)"
      Item(5).Control(14)=   "btnInfoTrigger(2)"
      Item(5).Control(15)=   "btnInfoTrigger(3)"
      Item(5).Control(16)=   "btnInfoTrigger(4)"
      Item(5).Control(17)=   "btnInfoTrigger(5)"
      Item(5).Control(18)=   "btnInfoTrigger(6)"
      Item(5).Control(19)=   "btnInfoTrigger(7)"
      Item(5).Control(20)=   "btnInfoTrigger(8)"
      Item(5).Control(21)=   "btnInfoTrigger(9)"
      Item(5).Control(22)=   "btnInfoTrigger(10)"
      Item(5).Control(23)=   "btnInfoTriggerTag"
      Item(5).Control(24)=   "btnInfo(11)"
      Item(5).Control(25)=   "btnInfoTrigger(11)"
      Item(6).Caption =   "Correo"
      Item(6).ControlCount=   2
      Item(6).Control(0)=   "lswCorreo"
      Item(6).Control(1)=   "scCorreo"
      Item(7).Caption =   "Beneficios"
      Item(7).ControlCount=   2
      Item(7).Control(0)=   "lswBeneficios"
      Item(7).Control(1)=   "gbBeneficios"
      Item(8).Caption =   "Renuncias"
      Item(8).ControlCount=   2
      Item(8).Control(0)=   "lswRenuncias"
      Item(8).Control(1)=   "gbRenuncias"
      Item(9).Caption =   "SoS"
      Item(9).ControlCount=   7
      Item(9).Control(0)=   "lswSoS"
      Item(9).Control(1)=   "lswSoS_Det"
      Item(9).Control(2)=   "Label9"
      Item(9).Control(3)=   "txtSoS_Monto"
      Item(9).Control(4)=   "chkSoS_Exclusion"
      Item(9).Control(5)=   "btnSoS_Export(0)"
      Item(9).Control(6)=   "btnSoS_Export(1)"
      Begin XtremeSuiteControls.ListView lswFND 
         Height          =   1812
         Left            =   120
         TabIndex        =   46
         Top             =   240
         Width           =   12372
         _Version        =   1572864
         _ExtentX        =   21823
         _ExtentY        =   3196
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
         Appearance      =   2
         ShowBorder      =   0   'False
         Arrange         =   2
      End
      Begin XtremeSuiteControls.ListView lswDP 
         Height          =   3372
         Left            =   -67960
         TabIndex        =   69
         Top             =   120
         Visible         =   0   'False
         Width           =   4812
         _Version        =   1572864
         _ExtentX        =   8488
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
         MultiSelect     =   -1  'True
         View            =   3
         FullRowSelect   =   -1  'True
         Appearance      =   16
      End
      Begin XtremeSuiteControls.ListView lswCorreo 
         Height          =   4092
         Left            =   -70000
         TabIndex        =   139
         Top             =   600
         Visible         =   0   'False
         Width           =   12492
         _Version        =   1572864
         _ExtentX        =   22034
         _ExtentY        =   7218
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
         Appearance      =   17
         UseVisualStyle  =   0   'False
         ShowBorder      =   0   'False
         Arrange         =   2
      End
      Begin XtremeSuiteControls.ListView lswBeneficios 
         Height          =   3375
         Left            =   -70000
         TabIndex        =   143
         Top             =   120
         Visible         =   0   'False
         Width           =   12615
         _Version        =   1572864
         _ExtentX        =   22251
         _ExtentY        =   5953
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
         Appearance      =   17
         ShowBorder      =   0   'False
      End
      Begin XtremeSuiteControls.ListView lswRenuncias 
         Height          =   3375
         Left            =   -70000
         TabIndex        =   150
         Top             =   120
         Visible         =   0   'False
         Width           =   12615
         _Version        =   1572864
         _ExtentX        =   22251
         _ExtentY        =   5953
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
         Appearance      =   17
         ShowBorder      =   0   'False
      End
      Begin XtremeSuiteControls.ListView lswFnd_List 
         Height          =   975
         Left            =   120
         TabIndex        =   162
         Top             =   2640
         Width           =   12495
         _Version        =   1572864
         _ExtentX        =   22040
         _ExtentY        =   1720
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
         Appearance      =   2
         ShowBorder      =   0   'False
         Arrange         =   2
      End
      Begin XtremeSuiteControls.ListView lswSoS 
         Height          =   2415
         Left            =   -69880
         TabIndex        =   164
         Top             =   120
         Visible         =   0   'False
         Width           =   12495
         _Version        =   1572864
         _ExtentX        =   22040
         _ExtentY        =   4260
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
         Appearance      =   2
         ShowBorder      =   0   'False
         Arrange         =   2
      End
      Begin XtremeSuiteControls.ListView lswSoS_Det 
         Height          =   1695
         Left            =   -69880
         TabIndex        =   165
         Top             =   3000
         Visible         =   0   'False
         Width           =   12495
         _Version        =   1572864
         _ExtentX        =   22040
         _ExtentY        =   2990
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
         Appearance      =   2
         ShowBorder      =   0   'False
         Arrange         =   2
      End
      Begin XtremeSuiteControls.PushButton btnSoS_Export 
         Height          =   330
         Index           =   0
         Left            =   -63400
         TabIndex        =   169
         Top             =   2640
         Visible         =   0   'False
         Width           =   1815
         _Version        =   1572864
         _ExtentX        =   3201
         _ExtentY        =   582
         _StockProps     =   79
         Caption         =   "Exportar Pagos"
         BackColor       =   -2147483643
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
         Picture         =   "frmCR_ConsultaCréditos.frx":918B
      End
      Begin XtremeSuiteControls.CheckBox chkSoS_Exclusion 
         Height          =   330
         Left            =   -66640
         TabIndex        =   168
         Top             =   2640
         Visible         =   0   'False
         Width           =   3135
         _Version        =   1572864
         _ExtentX        =   5530
         _ExtentY        =   582
         _StockProps     =   79
         Caption         =   "   Excluir del Proceso de Devolución"
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
      Begin XtremeSuiteControls.PushButton btnFondos_List 
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   158
         Top             =   2160
         Width           =   1335
         _Version        =   1572864
         _ExtentX        =   2355
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Movimientos"
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
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         Checked         =   -1  'True
      End
      Begin XtremeSuiteControls.GroupBox gbBeneficios 
         Height          =   855
         Left            =   -70000
         TabIndex        =   144
         Top             =   3720
         Visible         =   0   'False
         Width           =   12255
         _Version        =   1572864
         _ExtentX        =   21616
         _ExtentY        =   1508
         _StockProps     =   79
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         BorderStyle     =   1
         Begin XtremeSuiteControls.PushButton btnBeneficio 
            Height          =   615
            Index           =   0
            Left            =   120
            TabIndex        =   145
            Top             =   240
            Width           =   1215
            _Version        =   1572864
            _ExtentX        =   2143
            _ExtentY        =   1085
            _StockProps     =   79
            Caption         =   "Nuevo"
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
            TextAlignment   =   1
            Appearance      =   17
            Picture         =   "frmCR_ConsultaCréditos.frx":9A5C
            ImageAlignment  =   0
         End
         Begin XtremeSuiteControls.FlatEdit txtBeneCasos 
            Height          =   315
            Left            =   5640
            TabIndex        =   147
            Top             =   360
            Width           =   855
            _Version        =   1572864
            _ExtentX        =   1508
            _ExtentY        =   556
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
            Text            =   "0"
            Alignment       =   1
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtBeneTotal 
            Height          =   315
            Left            =   8160
            TabIndex        =   149
            Top             =   360
            Width           =   1455
            _Version        =   1572864
            _ExtentX        =   2566
            _ExtentY        =   556
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
            Text            =   "0"
            Alignment       =   1
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.PushButton btnBeneficio 
            Height          =   615
            Index           =   1
            Left            =   1320
            TabIndex        =   156
            Top             =   240
            Visible         =   0   'False
            Width           =   1455
            _Version        =   1572864
            _ExtentX        =   2566
            _ExtentY        =   1085
            _StockProps     =   79
            Caption         =   "Consulta Avanzada"
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
            TextAlignment   =   1
            Appearance      =   17
            Picture         =   "frmCR_ConsultaCréditos.frx":A08E
            ImageAlignment  =   0
         End
         Begin XtremeSuiteControls.Label Label2 
            Height          =   255
            Index           =   1
            Left            =   6120
            TabIndex        =   148
            Top             =   360
            Width           =   1935
            _Version        =   1572864
            _ExtentX        =   3413
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Monto Otorgado:"
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
         End
         Begin XtremeSuiteControls.Label Label2 
            Height          =   255
            Index           =   0
            Left            =   3600
            TabIndex        =   146
            Top             =   360
            Width           =   1935
            _Version        =   1572864
            _ExtentX        =   3413
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Cantidad de Beneficios:"
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
         End
      End
      Begin XtremeSuiteControls.RadioButton rbNotificaEmail 
         Height          =   372
         Index           =   0
         Left            =   -69760
         TabIndex        =   136
         Top             =   2160
         Visible         =   0   'False
         Width           =   1212
         _Version        =   1572864
         _ExtentX        =   2138
         _ExtentY        =   656
         _StockProps     =   79
         Caption         =   "Resumen"
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
      Begin XtremeSuiteControls.GroupBox fraMsj 
         Height          =   3135
         Left            =   -67360
         TabIndex        =   97
         Top             =   240
         Visible         =   0   'False
         Width           =   9615
         _Version        =   1572864
         _ExtentX        =   16960
         _ExtentY        =   5530
         _StockProps     =   79
         Caption         =   "Nuevo Mensaje: "
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
         BorderStyle     =   1
         Begin XtremeSuiteControls.CheckBox chkMsjVence 
            Height          =   255
            Left            =   480
            TabIndex        =   177
            Top             =   360
            Width           =   1215
            _Version        =   1572864
            _ExtentX        =   2143
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "No Vence ?  "
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
            TextAlignment   =   1
            Appearance      =   17
            Alignment       =   1
         End
         Begin XtremeSuiteControls.RadioButton optMsj 
            Height          =   255
            Index           =   0
            Left            =   4320
            TabIndex        =   100
            Top             =   360
            Visible         =   0   'False
            Width           =   1335
            _Version        =   1572864
            _ExtentX        =   2350
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "General"
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
            Value           =   -1  'True
         End
         Begin XtremeSuiteControls.DateTimePicker dtpMsjVence 
            Height          =   315
            Left            =   1800
            TabIndex        =   99
            ToolTipText     =   "Fecha de vencimiento del Mensaje "
            Top             =   360
            Width           =   1335
            _Version        =   1572864
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
         Begin XtremeSuiteControls.RadioButton optMsj 
            Height          =   255
            Index           =   1
            Left            =   5760
            TabIndex        =   101
            Top             =   360
            Visible         =   0   'False
            Width           =   1335
            _Version        =   1572864
            _ExtentX        =   2350
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Pendientes"
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
         Begin XtremeSuiteControls.RadioButton optMsj 
            Height          =   255
            Index           =   2
            Left            =   7200
            TabIndex        =   102
            Top             =   360
            Visible         =   0   'False
            Width           =   1335
            _Version        =   1572864
            _ExtentX        =   2350
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Morosidad"
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
         Begin XtremeSuiteControls.FlatEdit txtMsj 
            Height          =   2052
            Left            =   240
            TabIndex        =   103
            Top             =   840
            Width           =   9132
            _Version        =   1572864
            _ExtentX        =   16108
            _ExtentY        =   3619
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
         Begin VB.Image imgGuardaMsj 
            Height          =   252
            Left            =   8640
            Picture         =   "frmCR_ConsultaCréditos.frx":A95F
            Stretch         =   -1  'True
            ToolTipText     =   "Guardar Mensaje"
            Top             =   360
            Width           =   252
         End
         Begin VB.Image imgMsjCierraFrame 
            Height          =   252
            Left            =   9000
            Picture         =   "frmCR_ConsultaCréditos.frx":B13E
            Stretch         =   -1  'True
            ToolTipText     =   "Guardar Mensaje"
            Top             =   360
            Width           =   252
         End
         Begin XtremeSuiteControls.Label Label8 
            Height          =   255
            Index           =   1
            Left            =   3720
            TabIndex        =   98
            Top             =   360
            Visible         =   0   'False
            Width           =   855
            _Version        =   1572864
            _ExtentX        =   1503
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Tipo:"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
      Begin XtremeSuiteControls.GroupBox gbNotas 
         Height          =   855
         Left            =   -69880
         TabIndex        =   94
         Top             =   3840
         Visible         =   0   'False
         Width           =   10935
         _Version        =   1572864
         _ExtentX        =   19288
         _ExtentY        =   1508
         _StockProps     =   79
         Caption         =   "Notas de Bloqueo!"
         ForeColor       =   8421504
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
         Begin XtremeSuiteControls.FlatEdit txtNotas 
            Height          =   528
            Left            =   1680
            TabIndex        =   95
            Top             =   240
            Width           =   8412
            _Version        =   1572864
            _ExtentX        =   14838
            _ExtentY        =   931
            _StockProps     =   77
            ForeColor       =   255
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
            MultiLine       =   -1  'True
            ScrollBars      =   2
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.PushButton btnNotas 
            Height          =   264
            Index           =   0
            Left            =   10560
            TabIndex        =   111
            ToolTipText     =   "Guardar Nota [Principal]"
            Top             =   240
            Width           =   310
            _Version        =   1572864
            _ExtentX        =   547
            _ExtentY        =   466
            _StockProps     =   79
            FlatStyle       =   -1  'True
            UseVisualStyle  =   -1  'True
            Appearance      =   17
            Picture         =   "frmCR_ConsultaCréditos.frx":B8FB
            TextImageRelation=   0
         End
         Begin XtremeSuiteControls.PushButton btnNotas 
            Height          =   264
            Index           =   1
            Left            =   10560
            TabIndex        =   112
            ToolTipText     =   "Registrar Bloqueo de Formalizaciones"
            Top             =   480
            Width           =   310
            _Version        =   1572864
            _ExtentX        =   547
            _ExtentY        =   466
            _StockProps     =   79
            FlatStyle       =   -1  'True
            UseVisualStyle  =   -1  'True
            Appearance      =   17
            Picture         =   "frmCR_ConsultaCréditos.frx":C02C
            TextImageRelation=   0
         End
         Begin VB.Label lblBloqueo 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   528
            Left            =   10200
            TabIndex        =   96
            ToolTipText     =   "Persona Bloqueada o No"
            Top             =   240
            Width           =   252
         End
      End
      Begin XtremeSuiteControls.PushButton btnConstancia 
         Height          =   330
         Left            =   -59320
         TabIndex        =   93
         Top             =   3360
         Visible         =   0   'False
         Width           =   330
         _Version        =   1572864
         _ExtentX        =   582
         _ExtentY        =   582
         _StockProps     =   79
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         Picture         =   "frmCR_ConsultaCréditos.frx":C729
         TextImageRelation=   0
      End
      Begin FPSpreadADO.fpSpread vgCreditos 
         Height          =   3015
         Left            =   -68200
         TabIndex        =   25
         Top             =   240
         Visible         =   0   'False
         Width           =   9135
         _Version        =   524288
         _ExtentX        =   16113
         _ExtentY        =   5318
         _StockProps     =   64
         BackColorStyle  =   1
         BorderStyle     =   0
         DisplayRowHeaders=   0   'False
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
         MaxCols         =   23
         SpreadDesigner  =   "frmCR_ConsultaCréditos.frx":CE30
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin XtremeSuiteControls.FlatEdit txtTotalSaldo 
         Height          =   288
         Left            =   -65320
         TabIndex        =   26
         Top             =   3360
         Visible         =   0   'False
         Width           =   1332
         _Version        =   1572864
         _ExtentX        =   2350
         _ExtentY        =   508
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtTotalMonto 
         Height          =   288
         Left            =   -67480
         TabIndex        =   27
         Top             =   3360
         Visible         =   0   'False
         Width           =   1452
         _Version        =   1572864
         _ExtentX        =   2561
         _ExtentY        =   508
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtTotalCuota 
         Height          =   288
         Left            =   -63280
         TabIndex        =   28
         Top             =   3360
         Visible         =   0   'False
         Width           =   1332
         _Version        =   1572864
         _ExtentX        =   2350
         _ExtentY        =   508
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.DateTimePicker dtpCorte 
         Height          =   312
         Left            =   -60760
         TabIndex        =   29
         Top             =   3360
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   3
      End
      Begin XtremeSuiteControls.PushButton btnCredito 
         Height          =   310
         Index           =   0
         Left            =   -69880
         TabIndex        =   30
         Top             =   240
         Visible         =   0   'False
         Width           =   1452
         _Version        =   1572864
         _ExtentX        =   2561
         _ExtentY        =   547
         _StockProps     =   79
         Caption         =   "Nuevo"
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
      Begin XtremeSuiteControls.PushButton btnCredito 
         Height          =   310
         Index           =   1
         Left            =   -69880
         TabIndex        =   31
         Top             =   600
         Visible         =   0   'False
         Width           =   1452
         _Version        =   1572864
         _ExtentX        =   2561
         _ExtentY        =   547
         _StockProps     =   79
         Caption         =   "Crd. Exc."
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
      Begin XtremeSuiteControls.PushButton btnCredito 
         Height          =   310
         Index           =   2
         Left            =   -69880
         TabIndex        =   32
         Top             =   960
         Visible         =   0   'False
         Width           =   1452
         _Version        =   1572864
         _ExtentX        =   2561
         _ExtentY        =   547
         _StockProps     =   79
         Caption         =   "Cálculo"
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
      Begin XtremeSuiteControls.PushButton btnCredito 
         Height          =   310
         Index           =   3
         Left            =   -69880
         TabIndex        =   33
         Top             =   1320
         Visible         =   0   'False
         Width           =   1452
         _Version        =   1572864
         _ExtentX        =   2561
         _ExtentY        =   547
         _StockProps     =   79
         Caption         =   "Estudio Crédito"
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
      Begin XtremeSuiteControls.PushButton btnCredito 
         Height          =   310
         Index           =   4
         Left            =   -69880
         TabIndex        =   34
         Top             =   1800
         Visible         =   0   'False
         Width           =   1452
         _Version        =   1572864
         _ExtentX        =   2561
         _ExtentY        =   547
         _StockProps     =   79
         Caption         =   "Fianzas"
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
      Begin XtremeSuiteControls.PushButton btnCredito 
         Height          =   310
         Index           =   5
         Left            =   -69880
         TabIndex        =   35
         Top             =   2160
         Visible         =   0   'False
         Width           =   1452
         _Version        =   1572864
         _ExtentX        =   2561
         _ExtentY        =   547
         _StockProps     =   79
         Caption         =   "Morosidad"
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
      Begin XtremeSuiteControls.PushButton btnCredito 
         Height          =   310
         Index           =   6
         Left            =   -69880
         TabIndex        =   36
         Top             =   2640
         Visible         =   0   'False
         Width           =   1452
         _Version        =   1572864
         _ExtentX        =   2561
         _ExtentY        =   547
         _StockProps     =   79
         Caption         =   "Deducciones"
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
      Begin XtremeSuiteControls.PushButton btnCredito 
         Height          =   310
         Index           =   7
         Left            =   -69880
         TabIndex        =   37
         Top             =   3000
         Visible         =   0   'False
         Width           =   1452
         _Version        =   1572864
         _ExtentX        =   2561
         _ExtentY        =   547
         _StockProps     =   79
         Caption         =   "Preliminar"
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
      Begin XtremeSuiteControls.PushButton btnCredito 
         Height          =   310
         Index           =   8
         Left            =   -69880
         TabIndex        =   38
         Top             =   3360
         Visible         =   0   'False
         Width           =   1452
         _Version        =   1572864
         _ExtentX        =   2561
         _ExtentY        =   547
         _StockProps     =   79
         Caption         =   "Transacciones"
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
      Begin FPSpreadADO.fpSpread vgCobro 
         Height          =   4452
         Left            =   -68200
         TabIndex        =   43
         Top             =   240
         Visible         =   0   'False
         Width           =   10572
         _Version        =   524288
         _ExtentX        =   18648
         _ExtentY        =   7853
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
         MaxCols         =   11
         SpreadDesigner  =   "frmCR_ConsultaCréditos.frx":F2D2
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin XtremeSuiteControls.PushButton isButtonCb 
         Height          =   375
         Index           =   0
         Left            =   -69880
         TabIndex        =   44
         Top             =   240
         Visible         =   0   'False
         Width           =   1452
         _Version        =   1572864
         _ExtentX        =   2561
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Expediente"
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
      Begin XtremeSuiteControls.PushButton isButtonCb 
         Height          =   375
         Index           =   1
         Left            =   -69880
         TabIndex        =   45
         Top             =   600
         Visible         =   0   'False
         Width           =   1452
         _Version        =   1572864
         _ExtentX        =   2561
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Advertencias"
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
      Begin XtremeSuiteControls.GroupBox gbFndContrato 
         Height          =   855
         Left            =   120
         TabIndex        =   47
         Top             =   3720
         Width           =   12735
         _Version        =   1572864
         _ExtentX        =   22463
         _ExtentY        =   1508
         _StockProps     =   79
         Caption         =   "Contrato No. "
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
         Begin XtremeSuiteControls.CheckBox chkFndContrato 
            Height          =   492
            Left            =   720
            TabIndex        =   48
            Top             =   240
            Width           =   1692
            _Version        =   1572864
            _ExtentX        =   2984
            _ExtentY        =   868
            _StockProps     =   79
            Caption         =   "Detallado?"
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
            Value           =   1
            Alignment       =   1
         End
         Begin XtremeSuiteControls.PushButton btnFondos 
            Height          =   528
            Index           =   0
            Left            =   2760
            TabIndex        =   49
            Top             =   240
            Width           =   1572
            _Version        =   1572864
            _ExtentX        =   2773
            _ExtentY        =   931
            _StockProps     =   79
            Caption         =   "Estado de Cuenta"
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
            Picture         =   "frmCR_ConsultaCréditos.frx":FF97
         End
         Begin XtremeSuiteControls.PushButton btnFondos 
            Height          =   528
            Index           =   2
            Left            =   6000
            TabIndex        =   50
            Top             =   240
            Width           =   1692
            _Version        =   1572864
            _ExtentX        =   2984
            _ExtentY        =   931
            _StockProps     =   79
            Caption         =   "Registro de Contrato"
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
            Picture         =   "frmCR_ConsultaCréditos.frx":10753
         End
         Begin XtremeSuiteControls.PushButton btnFondos 
            Height          =   528
            Index           =   3
            Left            =   7800
            TabIndex        =   51
            Top             =   240
            Width           =   1452
            _Version        =   1572864
            _ExtentX        =   2561
            _ExtentY        =   931
            _StockProps     =   79
            Caption         =   "Aportes"
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
            Picture         =   "frmCR_ConsultaCréditos.frx":10E07
         End
         Begin XtremeSuiteControls.PushButton btnFondos 
            Height          =   528
            Index           =   1
            Left            =   4320
            TabIndex        =   52
            Top             =   240
            Width           =   1572
            _Version        =   1572864
            _ExtentX        =   2773
            _ExtentY        =   931
            _StockProps     =   79
            Caption         =   "Consulta Avanzada"
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
            Picture         =   "frmCR_ConsultaCréditos.frx":115F6
         End
         Begin XtremeSuiteControls.PushButton btnFondos 
            Height          =   528
            Index           =   4
            Left            =   9240
            TabIndex        =   131
            Top             =   240
            Width           =   1452
            _Version        =   1572864
            _ExtentX        =   2561
            _ExtentY        =   931
            _StockProps     =   79
            Caption         =   "Retiros"
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
            Picture         =   "frmCR_ConsultaCréditos.frx":12014
         End
         Begin XtremeSuiteControls.PushButton btnFondos 
            Height          =   525
            Index           =   5
            Left            =   10800
            TabIndex        =   185
            Top             =   240
            Width           =   1695
            _Version        =   1572864
            _ExtentX        =   2990
            _ExtentY        =   926
            _StockProps     =   79
            Caption         =   "Calculadora"
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
            Picture         =   "frmCR_ConsultaCréditos.frx":12709
         End
      End
      Begin FPSpreadADO.fpSpread vgPatrimonio 
         Height          =   4455
         Left            =   -66400
         TabIndex        =   53
         Top             =   240
         Visible         =   0   'False
         Width           =   8775
         _Version        =   524288
         _ExtentX        =   15478
         _ExtentY        =   7858
         _StockProps     =   64
         BackColorStyle  =   1
         BorderStyle     =   0
         DisplayRowHeaders=   0   'False
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
         MaxCols         =   8
         SpreadDesigner  =   "frmCR_ConsultaCréditos.frx":12A55
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin XtremeSuiteControls.ComboBox cboPAT_TipoSaldo 
         Height          =   288
         Left            =   -68200
         TabIndex        =   54
         Top             =   3480
         Visible         =   0   'False
         Width           =   1692
         _Version        =   1572864
         _ExtentX        =   2990
         _ExtentY        =   582
         _StockProps     =   77
         ForeColor       =   1973790
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   7.5
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
      Begin FPSpreadADO.fpSpread vGrid 
         Height          =   4452
         Left            =   -67240
         TabIndex        =   67
         Top             =   240
         Visible         =   0   'False
         Width           =   9492
         _Version        =   524288
         _ExtentX        =   16743
         _ExtentY        =   7853
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
         SpreadDesigner  =   "frmCR_ConsultaCréditos.frx":14782
         VScrollSpecial  =   -1  'True
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin XtremeSuiteControls.PushButton btnInfo 
         Height          =   288
         Index           =   0
         Left            =   -69760
         TabIndex        =   70
         Top             =   120
         Visible         =   0   'False
         Width           =   1332
         _Version        =   1572864
         _ExtentX        =   2350
         _ExtentY        =   503
         _StockProps     =   79
         Caption         =   "Contacto"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   12
      End
      Begin XtremeSuiteControls.PushButton btnInfo 
         Height          =   288
         Index           =   1
         Left            =   -69760
         TabIndex        =   71
         Top             =   480
         Visible         =   0   'False
         Width           =   1332
         _Version        =   1572864
         _ExtentX        =   2350
         _ExtentY        =   503
         _StockProps     =   79
         Caption         =   "Laboral"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   12
      End
      Begin XtremeSuiteControls.PushButton btnInfo 
         Height          =   288
         Index           =   2
         Left            =   -69760
         TabIndex        =   72
         Top             =   840
         Visible         =   0   'False
         Width           =   1332
         _Version        =   1572864
         _ExtentX        =   2350
         _ExtentY        =   503
         _StockProps     =   79
         Caption         =   "Teléfonos"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   12
      End
      Begin XtremeSuiteControls.PushButton btnInfo 
         Height          =   288
         Index           =   4
         Left            =   -69760
         TabIndex        =   73
         Top             =   1560
         Visible         =   0   'False
         Width           =   1332
         _Version        =   1572864
         _ExtentX        =   2350
         _ExtentY        =   503
         _StockProps     =   79
         Caption         =   "Ingresos"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   12
      End
      Begin XtremeSuiteControls.PushButton btnInfo 
         Height          =   288
         Index           =   5
         Left            =   -69760
         TabIndex        =   74
         Top             =   1920
         Visible         =   0   'False
         Width           =   1332
         _Version        =   1572864
         _ExtentX        =   2350
         _ExtentY        =   503
         _StockProps     =   79
         Caption         =   "Liquidaciones"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   12
      End
      Begin XtremeSuiteControls.PushButton btnInfo 
         Height          =   288
         Index           =   6
         Left            =   -69760
         TabIndex        =   75
         Top             =   2280
         Visible         =   0   'False
         Width           =   1332
         _Version        =   1572864
         _ExtentX        =   2350
         _ExtentY        =   503
         _StockProps     =   79
         Caption         =   "Beneficiarios"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   12
      End
      Begin XtremeSuiteControls.PushButton btnInfo 
         Height          =   288
         Index           =   8
         Left            =   -69760
         TabIndex        =   76
         Top             =   3000
         Visible         =   0   'False
         Width           =   1332
         _Version        =   1572864
         _ExtentX        =   2350
         _ExtentY        =   503
         _StockProps     =   79
         Caption         =   "G y P"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   12
      End
      Begin XtremeSuiteControls.PushButton btnInfo 
         Height          =   288
         Index           =   9
         Left            =   -69760
         TabIndex        =   77
         Top             =   3360
         Visible         =   0   'False
         Width           =   1332
         _Version        =   1572864
         _ExtentX        =   2350
         _ExtentY        =   503
         _StockProps     =   79
         Caption         =   "Bienes"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   12
      End
      Begin XtremeSuiteControls.PushButton btnInfo 
         Height          =   288
         Index           =   10
         Left            =   -69760
         TabIndex        =   78
         Top             =   3720
         Visible         =   0   'False
         Width           =   1332
         _Version        =   1572864
         _ExtentX        =   2350
         _ExtentY        =   503
         _StockProps     =   79
         Caption         =   "Escolaridad"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   12
      End
      Begin XtremeSuiteControls.PushButton btnInfo 
         Height          =   288
         Index           =   3
         Left            =   -69760
         TabIndex        =   79
         Top             =   1200
         Visible         =   0   'False
         Width           =   1332
         _Version        =   1572864
         _ExtentX        =   2350
         _ExtentY        =   503
         _StockProps     =   79
         Caption         =   "Cuentas"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   12
      End
      Begin XtremeSuiteControls.PushButton btnInfo 
         Height          =   288
         Index           =   7
         Left            =   -69760
         TabIndex        =   80
         Top             =   2640
         Visible         =   0   'False
         Width           =   1332
         _Version        =   1572864
         _ExtentX        =   2350
         _ExtentY        =   503
         _StockProps     =   79
         Caption         =   "Canales"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   12
      End
      Begin XtremeSuiteControls.PushButton btnInfoTrigger 
         Height          =   288
         Index           =   0
         Left            =   -68440
         TabIndex        =   81
         Top             =   120
         Visible         =   0   'False
         Width           =   384
         _Version        =   1572864
         _ExtentX        =   677
         _ExtentY        =   508
         _StockProps     =   79
         Caption         =   "..."
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   16
      End
      Begin XtremeSuiteControls.PushButton btnInfoTrigger 
         Height          =   288
         Index           =   1
         Left            =   -68440
         TabIndex        =   82
         Top             =   480
         Visible         =   0   'False
         Width           =   384
         _Version        =   1572864
         _ExtentX        =   677
         _ExtentY        =   508
         _StockProps     =   79
         Caption         =   "..."
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   16
      End
      Begin XtremeSuiteControls.PushButton btnInfoTrigger 
         Height          =   288
         Index           =   2
         Left            =   -68440
         TabIndex        =   83
         Top             =   840
         Visible         =   0   'False
         Width           =   384
         _Version        =   1572864
         _ExtentX        =   677
         _ExtentY        =   508
         _StockProps     =   79
         Caption         =   "..."
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   16
      End
      Begin XtremeSuiteControls.PushButton btnInfoTrigger 
         Height          =   288
         Index           =   3
         Left            =   -68440
         TabIndex        =   84
         Top             =   1200
         Visible         =   0   'False
         Width           =   384
         _Version        =   1572864
         _ExtentX        =   677
         _ExtentY        =   508
         _StockProps     =   79
         Caption         =   "..."
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   16
      End
      Begin XtremeSuiteControls.PushButton btnInfoTrigger 
         Height          =   288
         Index           =   4
         Left            =   -68440
         TabIndex        =   85
         Top             =   1560
         Visible         =   0   'False
         Width           =   384
         _Version        =   1572864
         _ExtentX        =   677
         _ExtentY        =   508
         _StockProps     =   79
         Caption         =   "..."
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   16
      End
      Begin XtremeSuiteControls.PushButton btnInfoTrigger 
         Height          =   288
         Index           =   5
         Left            =   -68440
         TabIndex        =   86
         Top             =   1920
         Visible         =   0   'False
         Width           =   384
         _Version        =   1572864
         _ExtentX        =   677
         _ExtentY        =   508
         _StockProps     =   79
         Caption         =   "..."
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   16
      End
      Begin XtremeSuiteControls.PushButton btnInfoTrigger 
         Height          =   288
         Index           =   6
         Left            =   -68440
         TabIndex        =   87
         Top             =   2280
         Visible         =   0   'False
         Width           =   384
         _Version        =   1572864
         _ExtentX        =   677
         _ExtentY        =   508
         _StockProps     =   79
         Caption         =   "..."
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   16
      End
      Begin XtremeSuiteControls.PushButton btnInfoTrigger 
         Height          =   288
         Index           =   7
         Left            =   -68440
         TabIndex        =   88
         Top             =   2640
         Visible         =   0   'False
         Width           =   384
         _Version        =   1572864
         _ExtentX        =   677
         _ExtentY        =   508
         _StockProps     =   79
         Caption         =   "..."
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   16
      End
      Begin XtremeSuiteControls.PushButton btnInfoTrigger 
         Height          =   288
         Index           =   8
         Left            =   -68440
         TabIndex        =   89
         Top             =   3000
         Visible         =   0   'False
         Width           =   384
         _Version        =   1572864
         _ExtentX        =   677
         _ExtentY        =   508
         _StockProps     =   79
         Caption         =   "..."
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   16
      End
      Begin XtremeSuiteControls.PushButton btnInfoTrigger 
         Height          =   288
         Index           =   9
         Left            =   -68440
         TabIndex        =   90
         Top             =   3360
         Visible         =   0   'False
         Width           =   384
         _Version        =   1572864
         _ExtentX        =   677
         _ExtentY        =   508
         _StockProps     =   79
         Caption         =   "..."
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   16
      End
      Begin XtremeSuiteControls.PushButton btnInfoTrigger 
         Height          =   288
         Index           =   10
         Left            =   -68440
         TabIndex        =   91
         Top             =   3720
         Visible         =   0   'False
         Width           =   384
         _Version        =   1572864
         _ExtentX        =   677
         _ExtentY        =   508
         _StockProps     =   79
         Caption         =   "..."
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   16
      End
      Begin XtremeSuiteControls.PushButton btnInfoTriggerTag 
         Height          =   288
         Left            =   -67960
         TabIndex        =   92
         Top             =   3720
         Visible         =   0   'False
         Width           =   1704
         _Version        =   1572864
         _ExtentX        =   3006
         _ExtentY        =   508
         _StockProps     =   79
         Caption         =   "Trigger Info"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   16
      End
      Begin XtremeSuiteControls.ComboBox cboMsj 
         Height          =   312
         Left            =   -69640
         TabIndex        =   104
         Top             =   600
         Visible         =   0   'False
         Width           =   2172
         _Version        =   1572864
         _ExtentX        =   3836
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
      Begin XtremeSuiteControls.FlatEdit txtPatrimonio 
         Height          =   312
         Left            =   -69040
         TabIndex        =   122
         Top             =   2160
         Visible         =   0   'False
         Width           =   1452
         _Version        =   1572864
         _ExtentX        =   2561
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtPAT_Disponible 
         Height          =   312
         Left            =   -68200
         TabIndex        =   123
         Top             =   3000
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
         Alignment       =   1
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtPAT_Saldos 
         Height          =   312
         Left            =   -68200
         TabIndex        =   124
         Top             =   3840
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
         Alignment       =   1
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtPAT_Giro 
         Height          =   312
         Left            =   -68200
         TabIndex        =   125
         Top             =   4200
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
         Alignment       =   1
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtAhorro 
         Height          =   312
         Left            =   -69040
         TabIndex        =   126
         Top             =   600
         Visible         =   0   'False
         Width           =   1452
         _Version        =   1572864
         _ExtentX        =   2561
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtAporte 
         Height          =   312
         Left            =   -69040
         TabIndex        =   127
         Top             =   960
         Visible         =   0   'False
         Width           =   1452
         _Version        =   1572864
         _ExtentX        =   2561
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtCapitalizacion 
         Height          =   312
         Left            =   -69040
         TabIndex        =   128
         Top             =   1320
         Visible         =   0   'False
         Width           =   1452
         _Version        =   1572864
         _ExtentX        =   2561
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtCustodia 
         Height          =   312
         Left            =   -69040
         TabIndex        =   129
         Top             =   1680
         Visible         =   0   'False
         Width           =   1452
         _Version        =   1572864
         _ExtentX        =   2561
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.PushButton btnConstanciaAportes 
         Height          =   330
         Left            =   -69040
         TabIndex        =   130
         ToolTipText     =   "Constancias e Informes"
         Top             =   120
         Visible         =   0   'False
         Width           =   1215
         _Version        =   1572864
         _ExtentX        =   2143
         _ExtentY        =   582
         _StockProps     =   79
         Caption         =   " Informes"
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
         Picture         =   "frmCR_ConsultaCréditos.frx":14F3F
         ImageAlignment  =   0
         TextImageRelation=   0
      End
      Begin XtremeSuiteControls.FlatEdit txtPat_Divisa 
         Height          =   312
         Left            =   -67600
         TabIndex        =   132
         Top             =   2160
         Visible         =   0   'False
         Width           =   1092
         _Version        =   1572864
         _ExtentX        =   1926
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
      Begin XtremeSuiteControls.ComboBox cboPat_Garantia 
         Height          =   288
         Left            =   -68200
         TabIndex        =   134
         Top             =   2640
         Visible         =   0   'False
         Width           =   1692
         _Version        =   1572864
         _ExtentX        =   2990
         _ExtentY        =   582
         _StockProps     =   77
         ForeColor       =   1973790
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Narrow"
            Size            =   7.5
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
      Begin XtremeSuiteControls.PushButton isButtonCb 
         Height          =   768
         Index           =   2
         Left            =   -69880
         TabIndex        =   135
         Top             =   1200
         Visible         =   0   'False
         Width           =   1452
         _Version        =   1572864
         _ExtentX        =   2561
         _ExtentY        =   1355
         _StockProps     =   79
         Caption         =   "Notificación vía Email"
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
      Begin XtremeSuiteControls.RadioButton rbNotificaEmail 
         Height          =   372
         Index           =   1
         Left            =   -69760
         TabIndex        =   137
         Top             =   2520
         Visible         =   0   'False
         Width           =   1212
         _Version        =   1572864
         _ExtentX        =   2138
         _ExtentY        =   656
         _StockProps     =   79
         Caption         =   "Detalle"
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
      Begin XtremeSuiteControls.GroupBox gbRenuncias 
         Height          =   855
         Left            =   -70000
         TabIndex        =   151
         Top             =   3720
         Visible         =   0   'False
         Width           =   12255
         _Version        =   1572864
         _ExtentX        =   21616
         _ExtentY        =   1508
         _StockProps     =   79
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         BorderStyle     =   1
         Begin XtremeSuiteControls.PushButton btnRenuncia 
            Height          =   615
            Index           =   0
            Left            =   120
            TabIndex        =   152
            Top             =   240
            Width           =   1215
            _Version        =   1572864
            _ExtentX        =   2143
            _ExtentY        =   1085
            _StockProps     =   79
            Caption         =   "Nueva"
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
            TextAlignment   =   1
            Appearance      =   17
            Picture         =   "frmCR_ConsultaCréditos.frx":15646
            ImageAlignment  =   0
         End
         Begin XtremeSuiteControls.FlatEdit txtRenunciasCasos 
            Height          =   315
            Left            =   6000
            TabIndex        =   153
            Top             =   360
            Width           =   855
            _Version        =   1572864
            _ExtentX        =   1508
            _ExtentY        =   556
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
            Text            =   "0"
            Alignment       =   1
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.PushButton btnRenuncia 
            Height          =   615
            Index           =   1
            Left            =   1320
            TabIndex        =   155
            Top             =   240
            Width           =   1575
            _Version        =   1572864
            _ExtentX        =   2778
            _ExtentY        =   1085
            _StockProps     =   79
            Caption         =   "Consulta Avanzada"
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
            TextAlignment   =   1
            Appearance      =   17
            Picture         =   "frmCR_ConsultaCréditos.frx":15C78
            ImageAlignment  =   0
         End
         Begin XtremeSuiteControls.Label Label2 
            Height          =   255
            Index           =   3
            Left            =   3960
            TabIndex        =   154
            Top             =   360
            Width           =   1935
            _Version        =   1572864
            _ExtentX        =   3413
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Cantidad de Renuncias:"
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
         End
      End
      Begin XtremeSuiteControls.PushButton btnPatrimonioConsulta 
         Height          =   330
         Left            =   -67720
         TabIndex        =   157
         ToolTipText     =   "Consulta Avanzada de Patrimonio"
         Top             =   120
         Visible         =   0   'False
         Width           =   1215
         _Version        =   1572864
         _ExtentX        =   2143
         _ExtentY        =   582
         _StockProps     =   79
         Caption         =   "Avanzada"
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
         Picture         =   "frmCR_ConsultaCréditos.frx":16549
         ImageAlignment  =   0
         TextImageRelation=   0
      End
      Begin XtremeSuiteControls.PushButton btnFondos_List 
         Height          =   375
         Index           =   1
         Left            =   1440
         TabIndex        =   159
         Top             =   2160
         Width           =   1335
         _Version        =   1572864
         _ExtentX        =   2355
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Cupones"
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
         UseVisualStyle  =   -1  'True
         Appearance      =   17
      End
      Begin XtremeSuiteControls.PushButton btnFondos_List 
         Height          =   375
         Index           =   2
         Left            =   2760
         TabIndex        =   160
         Top             =   2160
         Width           =   1335
         _Version        =   1572864
         _ExtentX        =   2355
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Bitácora"
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
         UseVisualStyle  =   -1  'True
         Appearance      =   17
      End
      Begin XtremeSuiteControls.PushButton btnFondos_Export 
         Height          =   375
         Left            =   5520
         TabIndex        =   161
         ToolTipText     =   "Exportar a Excel"
         Top             =   2160
         Width           =   495
         _Version        =   1572864
         _ExtentX        =   873
         _ExtentY        =   661
         _StockProps     =   79
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
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         Picture         =   "frmCR_ConsultaCréditos.frx":16C49
      End
      Begin XtremeSuiteControls.FlatEdit txtSoS_Monto 
         Height          =   330
         Left            =   -68560
         TabIndex        =   167
         Top             =   2640
         Visible         =   0   'False
         Width           =   1815
         _Version        =   1572864
         _ExtentX        =   3201
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
         Text            =   "0.00"
         Alignment       =   1
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.PushButton btnSoS_Export 
         Height          =   330
         Index           =   1
         Left            =   -61600
         TabIndex        =   170
         Top             =   2640
         Visible         =   0   'False
         Width           =   2295
         _Version        =   1572864
         _ExtentX        =   4048
         _ExtentY        =   582
         _StockProps     =   79
         Caption         =   "Exportar Movimientos"
         BackColor       =   -2147483643
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
         Picture         =   "frmCR_ConsultaCréditos.frx":16DB3
      End
      Begin XtremeSuiteControls.PushButton btnFondos_List 
         Height          =   375
         Index           =   3
         Left            =   4080
         TabIndex        =   174
         Top             =   2160
         Width           =   1215
         _Version        =   1572864
         _ExtentX        =   2143
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Cierres"
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
         UseVisualStyle  =   -1  'True
         Appearance      =   17
      End
      Begin XtremeSuiteControls.PushButton btnInfo 
         Height          =   285
         Index           =   11
         Left            =   -69760
         TabIndex        =   175
         ToolTipText     =   "Beneficiarios de Pólizas Colectivas"
         Top             =   4080
         Visible         =   0   'False
         Width           =   1335
         _Version        =   1572864
         _ExtentX        =   2350
         _ExtentY        =   503
         _StockProps     =   79
         Caption         =   "Bene. Pólizas"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   12
      End
      Begin XtremeSuiteControls.PushButton btnInfoTrigger 
         Height          =   285
         Index           =   11
         Left            =   -68440
         TabIndex        =   176
         Top             =   4080
         Visible         =   0   'False
         Width           =   390
         _Version        =   1572864
         _ExtentX        =   677
         _ExtentY        =   508
         _StockProps     =   79
         Caption         =   "..."
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   16
      End
      Begin XtremeSuiteControls.FlatEdit txtPAT_AporteCobro 
         Height          =   315
         Left            =   -68200
         TabIndex        =   178
         Top             =   4800
         Visible         =   0   'False
         Width           =   1695
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
         Alignment       =   1
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Info Aporte en cobro: "
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   5
         Left            =   -69880
         TabIndex        =   179
         Top             =   4800
         Visible         =   0   'False
         Width           =   1695
      End
      Begin XtremeSuiteControls.Label Label9 
         Height          =   255
         Left            =   -69760
         TabIndex        =   166
         Top             =   2640
         Visible         =   0   'False
         Width           =   975
         _Version        =   1572864
         _ExtentX        =   1720
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Total: "
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
         Transparent     =   -1  'True
         WordWrap        =   -1  'True
      End
      Begin XtremeShortcutBar.ShortcutCaption scCorreo 
         Height          =   372
         Left            =   -70000
         TabIndex        =   140
         Top             =   240
         Visible         =   0   'False
         Width           =   12612
         _Version        =   1572864
         _ExtentX        =   22246
         _ExtentY        =   656
         _StockProps     =   14
         Caption         =   "Bandeja de Salida: "
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
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Garantías /Patrimonio"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   252
         Index           =   2
         Left            =   -69880
         TabIndex        =   133
         Top             =   2640
         Visible         =   0   'False
         Width           =   1932
      End
      Begin VB.Image imgBorraMsj 
         Height          =   255
         Left            =   -68080
         Picture         =   "frmCR_ConsultaCréditos.frx":17684
         Stretch         =   -1  'True
         ToolTipText     =   "Eliminar Mensajes Marcados"
         Top             =   1200
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Image imgMsjNuevo 
         Height          =   255
         Left            =   -67720
         Picture         =   "frmCR_ConsultaCréditos.frx":17E41
         Stretch         =   -1  'True
         ToolTipText     =   "Crear Nuevo Mensaje"
         Top             =   1200
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label Label6 
         Caption         =   "Tipo de Mensaje ...:"
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
         Left            =   -69640
         TabIndex        =   68
         Top             =   240
         Visible         =   0   'False
         Width           =   1692
      End
      Begin VB.Image imgMsjResuelve 
         Height          =   255
         Left            =   -68440
         Picture         =   "frmCR_ConsultaCréditos.frx":18620
         Stretch         =   -1  'True
         ToolTipText     =   "Quita Pendiente de Mensajes Marcados"
         Top             =   1200
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Giro Maximo: "
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   252
         Index           =   6
         Left            =   -69880
         TabIndex        =   66
         Top             =   4200
         Visible         =   0   'False
         Width           =   1692
      End
      Begin VB.Label lblPAT_Saldo 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "(-) Saldos Refinanciar:"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   252
         Left            =   -69880
         TabIndex        =   65
         Top             =   3840
         Visible         =   0   'False
         Width           =   1692
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Disponible Bruto:"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   252
         Index           =   4
         Left            =   -69880
         TabIndex        =   64
         Top             =   3000
         Visible         =   0   'False
         Width           =   1692
      End
      Begin VB.Label Label4 
         Caption         =   "Capitaliza"
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
         Height          =   252
         Index           =   1
         Left            =   -69880
         TabIndex        =   63
         Top             =   1320
         Visible         =   0   'False
         Width           =   1092
      End
      Begin VB.Label lblCapitalizado 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         Caption         =   "09-1997"
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
         Left            =   -67600
         TabIndex        =   62
         ToolTipText     =   "Fecha de la capitalización de los excedentes"
         Top             =   1320
         Visible         =   0   'False
         Width           =   1092
      End
      Begin VB.Label lblFechaCustodia 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         Caption         =   "10-1998"
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
         Left            =   -67600
         TabIndex        =   61
         ToolTipText     =   "Fecha del último ahorro extraordinario de este socio"
         Top             =   1680
         Visible         =   0   'False
         Width           =   1092
      End
      Begin VB.Label lblFechaAporte 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         Caption         =   "10-1998"
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
         Left            =   -67600
         TabIndex        =   60
         ToolTipText     =   "Fecha del último aporte patronal reportado"
         Top             =   960
         Visible         =   0   'False
         Width           =   1092
      End
      Begin VB.Label lblFechaAhorro 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         Caption         =   "10-1998"
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
         Left            =   -67600
         TabIndex        =   59
         ToolTipText     =   "Fecha del último ahorro obrero reportado"
         Top             =   600
         Visible         =   0   'False
         Width           =   1092
      End
      Begin VB.Label Label3 
         Caption         =   "Obrero"
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
         Height          =   252
         Index           =   0
         Left            =   -69880
         TabIndex        =   58
         Top             =   600
         Visible         =   0   'False
         Width           =   972
      End
      Begin VB.Label Label4 
         Caption         =   "Patronal"
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
         Height          =   252
         Index           =   0
         Left            =   -69880
         TabIndex        =   57
         Top             =   960
         Visible         =   0   'False
         Width           =   1092
      End
      Begin VB.Label Label5 
         Caption         =   "Custodia"
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
         Height          =   252
         Index           =   0
         Left            =   -69880
         TabIndex        =   56
         Top             =   1680
         Visible         =   0   'False
         Width           =   732
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Total:"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   252
         Index           =   3
         Left            =   -69880
         TabIndex        =   55
         Top             =   2160
         Visible         =   0   'False
         Width           =   852
      End
      Begin VB.Label Label1 
         Caption         =   "Corte [Int.]"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   312
         Index           =   3
         Left            =   -61720
         TabIndex        =   42
         ToolTipText     =   "Fecha Corte de Intereses para Cálculo de Cancelaciones"
         Top             =   3360
         Visible         =   0   'False
         Width           =   1692
      End
      Begin VB.Label Label1 
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
         Height          =   312
         Index           =   0
         Left            =   -68200
         TabIndex        =   41
         Top             =   3360
         Visible         =   0   'False
         Width           =   732
      End
      Begin VB.Label Label1 
         Caption         =   "Saldo"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   312
         Index           =   1
         Left            =   -65920
         TabIndex        =   40
         Top             =   3360
         Visible         =   0   'False
         Width           =   732
      End
      Begin VB.Label Label1 
         Caption         =   "Cuota"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   312
         Index           =   2
         Left            =   -63880
         TabIndex        =   39
         Top             =   3360
         Visible         =   0   'False
         Width           =   732
      End
   End
   Begin XtremeSuiteControls.PushButton btnIdentificarDP 
      Height          =   330
      Left            =   10680
      TabIndex        =   172
      ToolTipText     =   "Identificación de Depósitos"
      Top             =   105
      Width           =   735
      _Version        =   1572864
      _ExtentX        =   1296
      _ExtentY        =   582
      _StockProps     =   79
      Caption         =   "DPs"
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
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.ProgressBar ProgressBarX 
      Height          =   135
      Left            =   0
      TabIndex        =   173
      Top             =   1920
      Visible         =   0   'False
      Width           =   13215
      _Version        =   1572864
      _ExtentX        =   23310
      _ExtentY        =   238
      _StockProps     =   93
      Scrolling       =   1
   End
   Begin XtremeSuiteControls.PushButton btnAdjuntos 
      Height          =   330
      Left            =   11400
      TabIndex        =   180
      ToolTipText     =   "Adjuntar Documentos"
      Top             =   105
      Width           =   495
      _Version        =   1572864
      _ExtentX        =   873
      _ExtentY        =   582
      _StockProps     =   79
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
      UseVisualStyle  =   -1  'True
      Picture         =   "frmCR_ConsultaCréditos.frx":18DE8
   End
   Begin VB.Image imgMsjAdvertencia 
      Height          =   255
      Left            =   13320
      Picture         =   "frmCR_ConsultaCréditos.frx":18E71
      Stretch         =   -1  'True
      Top             =   600
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label lblMsjAdvertencias 
      BackStyle       =   0  'Transparent
      Caption         =   "Msj Advertencia?"
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
      Height          =   255
      Left            =   13680
      TabIndex        =   184
      Top             =   600
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Label lblMsjMorosidad 
      BackStyle       =   0  'Transparent
      Caption         =   "Msj Morosidad?"
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
      Height          =   255
      Left            =   13680
      TabIndex        =   183
      Top             =   840
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Image imgMsjMorosidad 
      Height          =   255
      Left            =   13320
      Picture         =   "frmCR_ConsultaCréditos.frx":19640
      Stretch         =   -1  'True
      Top             =   840
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image imgTS 
      Height          =   255
      Left            =   9720
      Picture         =   "frmCR_ConsultaCréditos.frx":19E0F
      Stretch         =   -1  'True
      ToolTipText     =   "Tarjeta Débito"
      Top             =   840
      Width           =   255
   End
   Begin VB.Label lblSalarioTraslada 
      BackStyle       =   0  'Transparent
      Caption         =   "Salario Traslado: No"
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
      Height          =   255
      Left            =   10080
      TabIndex        =   182
      Top             =   840
      Width           =   8895
      WordWrap        =   -1  'True
   End
   Begin XtremeShortcutBar.ShortcutCaption ShortcutCaptionTitle 
      Height          =   375
      Left            =   0
      TabIndex        =   121
      Top             =   2025
      Width           =   13695
      _Version        =   1572864
      _ExtentX        =   24156
      _ExtentY        =   661
      _StockProps     =   14
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      VisualTheme     =   6
   End
   Begin VB.Label lblIBAN 
      BackStyle       =   0  'Transparent
      Caption         =   "Cuenta Ahorros: No"
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
      Height          =   255
      Left            =   10080
      TabIndex        =   23
      Top             =   1500
      Width           =   3135
   End
   Begin VB.Label lblTarjeta 
      BackStyle       =   0  'Transparent
      Caption         =   "Tajeta Debito: No"
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
      Height          =   255
      Left            =   10080
      TabIndex        =   22
      Top             =   1155
      Width           =   3135
   End
   Begin VB.Image Image3 
      Height          =   252
      Index           =   2
      Left            =   9720
      Picture         =   "frmCR_ConsultaCréditos.frx":1A4FF
      Stretch         =   -1  'True
      ToolTipText     =   "Cuenta de Ahorros Sinpe?"
      Top             =   1500
      Width           =   252
   End
   Begin VB.Image Image3 
      Height          =   252
      Index           =   1
      Left            =   9720
      Picture         =   "frmCR_ConsultaCréditos.frx":1ABA3
      Stretch         =   -1  'True
      ToolTipText     =   "Tarjeta Débito"
      Top             =   1155
      Width           =   252
   End
   Begin VB.Image imgAdvertencias 
      Height          =   255
      Index           =   1
      Left            =   9720
      Picture         =   "frmCR_ConsultaCréditos.frx":1B1B5
      Stretch         =   -1  'True
      ToolTipText     =   "Advertencias?"
      Top             =   540
      Width           =   255
   End
   Begin VB.Label lblEstadoAdvertencias 
      BackStyle       =   0  'Transparent
      Caption         =   "Advertencias?"
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
      Height          =   255
      Left            =   10080
      TabIndex        =   21
      Top             =   540
      Width           =   1935
   End
   Begin VB.Image imgMsjGenerales 
      Height          =   255
      Left            =   7680
      Picture         =   "frmCR_ConsultaCréditos.frx":1BB19
      Stretch         =   -1  'True
      Top             =   840
      Width           =   255
   End
   Begin VB.Label lblMsjGenerales 
      BackStyle       =   0  'Transparent
      Caption         =   "Msj Generales?"
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
      Height          =   255
      Left            =   8040
      TabIndex        =   16
      Top             =   840
      Width           =   2055
   End
   Begin VB.Label lblMsjPendientes 
      BackStyle       =   0  'Transparent
      Caption         =   "Msj Pendientes?"
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
      Height          =   255
      Left            =   8040
      TabIndex        =   15
      Top             =   1155
      Width           =   2055
   End
   Begin VB.Image imgMsjPendientes 
      Height          =   255
      Left            =   7680
      Picture         =   "frmCR_ConsultaCréditos.frx":1C2E8
      Stretch         =   -1  'True
      Top             =   1155
      Width           =   255
   End
   Begin VB.Label lblEstadoCobros 
      BackStyle       =   0  'Transparent
      Caption         =   "Gestiones de Cobros?"
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
      Height          =   255
      Left            =   8040
      TabIndex        =   14
      Top             =   1500
      Width           =   1935
   End
   Begin VB.Image Image3 
      Height          =   255
      Index           =   0
      Left            =   7680
      Picture         =   "frmCR_ConsultaCréditos.frx":1CAB7
      Stretch         =   -1  'True
      ToolTipText     =   "Gestiones de Cobro?"
      Top             =   1500
      Width           =   255
   End
   Begin VB.Label lblEstadoMensajes 
      BackStyle       =   0  'Transparent
      Caption         =   "Mensajes?"
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
      Height          =   255
      Left            =   8040
      TabIndex        =   13
      Top             =   540
      Width           =   1935
   End
   Begin VB.Image Image2 
      Height          =   255
      Left            =   7680
      Picture         =   "frmCR_ConsultaCréditos.frx":1D240
      Stretch         =   -1  'True
      ToolTipText     =   "Estado de Actualización de los Beneficiarios"
      Top             =   540
      Width           =   255
   End
   Begin VB.Label lblEstadoAutInfo 
      BackStyle       =   0  'Transparent
      Caption         =   "Estatus Aut.Info.?"
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
      Height          =   252
      Left            =   480
      TabIndex        =   12
      ToolTipText     =   "Estado de la Autorizacipon de Uso de la Información"
      Top             =   840
      Width           =   1692
   End
   Begin VB.Label lblEstadoBeneficiarios 
      BackStyle       =   0  'Transparent
      Caption         =   "Estatus Beneficiarios?"
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
      Height          =   252
      Left            =   480
      TabIndex        =   11
      ToolTipText     =   "Estados de Actualización de los Beneficiarios"
      Top             =   540
      Width           =   1812
   End
   Begin VB.Image imgEstadoConsentimiento 
      Height          =   255
      Left            =   120
      Picture         =   "frmCR_ConsultaCréditos.frx":1DA4D
      Stretch         =   -1  'True
      ToolTipText     =   "Estado de Autorización de Uso de Información Personal"
      Top             =   840
      Width           =   255
   End
   Begin VB.Image imgEstadoBeneficiarios 
      Height          =   255
      Left            =   120
      Picture         =   "frmCR_ConsultaCréditos.frx":1E21C
      Stretch         =   -1  'True
      ToolTipText     =   "Estado de Actualización de los Beneficiarios"
      Top             =   540
      Width           =   255
   End
   Begin VB.Image imgClasificacion 
      Height          =   255
      Left            =   2280
      Picture         =   "frmCR_ConsultaCréditos.frx":1E9EB
      Stretch         =   -1  'True
      ToolTipText     =   "Clasificacion ABCD de la persona"
      Top             =   1500
      Width           =   255
   End
   Begin VB.Label lblClasificacion 
      BackStyle       =   0  'Transparent
      Caption         =   "Clasificación ?"
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
      Height          =   255
      Left            =   2760
      TabIndex        =   6
      Top             =   1500
      Width           =   4815
   End
   Begin VB.Image imgMembresia 
      Height          =   255
      Left            =   2280
      Picture         =   "frmCR_ConsultaCréditos.frx":1F16B
      Stretch         =   -1  'True
      ToolTipText     =   "Membresía de la Persona"
      Top             =   1160
      Width           =   255
   End
   Begin VB.Image imgFianzas 
      Height          =   255
      Left            =   120
      Picture         =   "frmCR_ConsultaCréditos.frx":1F843
      Stretch         =   -1  'True
      ToolTipText     =   "Estado de las fianzas"
      Top             =   1460
      Width           =   255
   End
   Begin VB.Image imgCreditos 
      Height          =   255
      Left            =   120
      Picture         =   "frmCR_ConsultaCréditos.frx":20012
      Stretch         =   -1  'True
      ToolTipText     =   "Estado de los creditos"
      Top             =   1160
      Width           =   255
   End
   Begin VB.Label lblMembresia 
      BackColor       =   &H00000080&
      Caption         =   "Membresía ?"
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
      Height          =   255
      Left            =   2760
      TabIndex        =   5
      Top             =   1200
      Width           =   4815
   End
   Begin VB.Label lblFianzas 
      BackStyle       =   0  'Transparent
      Caption         =   "Estatus Fianzas ?"
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
      Height          =   252
      Left            =   480
      TabIndex        =   4
      ToolTipText     =   "Estado de las Fianzas"
      Top             =   1464
      Width           =   1692
   End
   Begin VB.Label lblCreditos 
      BackStyle       =   0  'Transparent
      Caption         =   "Estatus Créditos ?"
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
      Height          =   252
      Left            =   480
      TabIndex        =   3
      ToolTipText     =   "Estado General de los Créditos"
      Top             =   1164
      Width           =   1812
   End
   Begin VB.Label lblEstado 
      BackStyle       =   0  'Transparent
      Caption         =   "Estado ?"
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
      Height          =   255
      Left            =   2760
      TabIndex        =   2
      Top             =   540
      Width           =   4815
   End
   Begin VB.Image imgEstado 
      Height          =   240
      Left            =   2280
      Picture         =   "frmCR_ConsultaCréditos.frx":207E1
      ToolTipText     =   "Estado de la persona"
      Top             =   540
      Width           =   240
   End
   Begin VB.Label lblInstitución 
      BackStyle       =   0  'Transparent
      Caption         =   "Empresa / Deductora?"
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
      Height          =   375
      Left            =   2760
      TabIndex        =   1
      Top             =   840
      Width           =   4815
   End
   Begin VB.Image imgInstitucion 
      Height          =   240
      Left            =   2280
      Picture         =   "frmCR_ConsultaCréditos.frx":208F8
      ToolTipText     =   "Empresa / Deductora?"
      Top             =   840
      Width           =   240
   End
   Begin VB.Image imgBanner 
      Height          =   1935
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   13215
   End
End
Attribute VB_Name = "frmCR_ConsultaCreditos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vPaso As Boolean
Dim mFecha As Date

Const ID_Menu_Abonos = 10
Const ID_Menu_Anula = 11

Const ID_Menu_General_Cobros = 40
Const ID_Menu_Credito_Estado = 41

Const ID_Menu_Credito_New = 30
Const ID_Menu_Credito_Tramite = 31

Const ID_Menu_Credito_Estudio_New = 32
Const ID_Menu_Credito_Estudio_Open = 33

Const ID_Menu_Credito_PlanPagos = 34

Const ID_Menu_Cancela = 0

Dim vRA_Access As Boolean

Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem

Dim pPlan As String, pContrato As Long, pOperadora As Integer

Sub EstadoInicial()
On Error Resume Next

Call Limpia

    txtCedula.Enabled = True
    txtCedula.SetFocus

End Sub

Sub Limpia()

 txtCedula.Text = ""
 
 lblFianzas.Caption = "Estado de Fianzas?"
 lblCreditos.Caption = "Estado de Créditos?"
 lblMembresia.Caption = "Membresía?"
 lblMembresia.ToolTipText = ""
  
 lblEstado.Caption = "Estado Persona?"
 lblInstitución.Caption = "Empresa/Deductora?"
 lblInstitución.ToolTipText = ""
 
 lblClasificacion.Caption = "Clasificación?"
 
 txtAhorro.Text = 0
 txtAporte.Text = 0
 txtCustodia.Text = 0
 txtCapitalizacion.Text = 0

 lblFechaAhorro.Caption = ""
 lblFechaAporte.Caption = ""
 lblFechaCustodia.Caption = ""
 lblCapitalizado.Caption = ""

 txtNotas = ""
 Call isButton1_Click(0)
  
End Sub


Private Sub sbCreditos(Optional pSheet As Integer = 1)
Dim curCuota As Currency, curMonto As Currency
Dim curSaldo As Currency, vMora As Boolean
Dim i As Integer

'On Error Resume Next

Call isButton1_Click(0)

curCuota = 0
curMonto = 0
curSaldo = 0
vMora = False


txtTotalMonto.Text = ""
txtTotalSaldo.Text = ""
txtTotalCuota.Text = ""

StatusBar.Panels(5).Text = "0.00"
StatusBar.Panels(6).Text = "0.00"

Me.MousePointer = vbHourglass

vPaso = True
vMora = False

With vgCreditos
 .Sheet = pSheet
 .ActiveSheet = pSheet
 
 
 .MaxRows = 0
 strSQL = "exec spSys_Consulta_Integrada_Creditos '" & txtCedula.Text & "','" & Mid(.SheetName, 1, 1) & "'"
 
 Call OpenRecordSet(rs, strSQL)

  Do While Not rs.EOF
    .MaxRows = .MaxRows + 1
    .Row = .MaxRows

    
    For i = 1 To .MaxCols
      .Col = i
      Select Case i
        Case 1 'Status

              .TypePictPicture = imgSemaforos.ListImages.Item(1).Picture
        
             Select Case rs!ProcesoCod
              Case "N"
       
                If Not IsNull(rs!Referencia) Then
                    If rs!MoraCuota = 0 Then
                       .TypePictPicture = imgSemaforos.ListImages.Item(2).Picture
                      .TextTip = TextTipFixed
                      .TextTipDelay = 1000
                      .CellNoteIndicatorShape = CellNoteIndicatorShapeSquare
                      .CellNoteIndicatorColor = vbRed
                      .CellNote = "Referencia: " & rs!Referencia
                    End If
                    .FontBold = True
                End If
        
                If rs!IndicadorCbr > 0 Then
                  .TypePictPicture = imgSemaforos.ListImages.Item(9).Picture
                  .TextTip = TextTipFixed
                  .TextTipDelay = 1000
                
                  .CellNoteIndicatorShape = CellNoteIndicatorShapeSquare
                  .CellNoteIndicatorColor = vbRed
                  
                  .CellNote = "!!! Esta Operación fue Reversada de Cobro Judicial, Revise el Tab de Cobros para mayor información..!!!"
                            
                End If
              
                If rs!CbrExterno = 1 Then
                    .BackColor = RGB(199, 138, 156)
                End If
              
              
              Case "J"
                  .TypePictPicture = imgSemaforos.ListImages.Item(7).Picture
                   vMora = True
                       
                  .TextTip = TextTipFixed
                  .TextTipDelay = 1000
                
                  .CellNoteIndicatorShape = CellNoteIndicatorShapeTriangle
                  .CellNoteIndicatorColor = vbRed
                  
                  .CellNote = ">> Cobro Judicial <<" & vbCrLf _
                            & "Fecha : " & Format(rs!fecha_enviaproceso, "dd/mm/yyyy") & vbCrLf _
                            & "Nota  : " & rs!observacion_proceso & ""
              
              Case "T"
                    If rs!MoraCuota = 0 Then .TypePictPicture = imgSemaforos.ListImages.Item(10).Picture
                    
                    If rs!IndicadorCbr > 0 Then
                       .TypePictPicture = imgSemaforos.ListImages.Item(9).Picture
                    End If
        
             End Select
             
             
             
             If Mid(rs!Estado, 1, 1) = "C" Then
                .TypePictPicture = imgSemaforos.ListImages.Item(6).Picture
             End If

            ' Si esta moroso indicar Mora siempre y cuando no este en cobro Judicial
            If rs!MoraCuota > 0 And rs!ProcesoCod <> "J" Then
              
              .TypePictPicture = imgSemaforos.ListImages.Item(3).Picture
              vMora = True
            
              .TextTip = TextTipFixed
              .TextTipDelay = 1000
            
              .CellNoteIndicatorShape = CellNoteIndicatorShapeTriangle
              .CellNoteIndicatorColor = vbBlue
              
              .CellNote = "Referencia..:" & rs!Referencia & vbCrLf & "Morosidad:  Cuotas: " & rs!MoraCuota & vbCrLf _
                        & "   Intereses : " & Format(rs!MoraInt, "Standard") & vbCrLf _
                        & "   Cargos    : " & Format(rs!MoraCargos, "Standard") & vbCrLf _
                        & "   Póliza    : " & Format(rs!MoraPoliza, "Standard") & vbCrLf _
                        & "   Principal : " & Format(rs!MoraPrincipal, "Standard") & vbCrLf _
                        & "   Cta.+ Vieja : " & Format(rs!MoraAntigua, "####-##") & vbCrLf _
                        & "   Cta. Ultima : " & Format(rs!MoraUltima, "####-##") & vbCrLf & vbCrLf _
                        & "   Total Mora  : " & Format(rs!MoraInt + rs!MoraCargos + rs!MoraPrincipal + rs!MoraPoliza, "Standard") & vbCrLf _
                        & "   Antiguedad  : " & rs!Antiguedad
            
            End If
        
        Case 2 'Check + Currency
           
           .Col = 3
           .CellTag = CStr(rs!Id_Solicitud)
           .Text = CStr(rs!Id_Solicitud)
            
            .Col = 2
           .CellTag = CStr(rs!Id_Solicitud)
'           .Text = CStr(rs!Id_Solicitud)

           If pSheet = 1 Then
                .TypeCheckText = CStr(rs!CURRENCY_SIM)
           Else
                .Text = CStr(rs!CURRENCY_SIM)
           End If
           
'           .CellTag = CStr(rs!Id_Solicitud)
'           If pSheet = 1 Then
'                .TypeCheckText = CStr(rs!Id_Solicitud)
'           Else
'                .Text = CStr(rs!Id_Solicitud)
'           End If
        
        
        Case 3 'Operacion
           .CellTag = CStr(rs!Id_Solicitud)
           .Text = CStr(rs!Id_Solicitud)
           If pSheet = 1 Then
                .TypeCheckText = CStr(rs!Id_Solicitud)
           End If
        
        Case 4 'Linea
            .Text = rs!Codigo
            .TextTip = TextTipFixed
            .TextTipDelay = 1000
            .CellNoteIndicatorShape = CellNoteIndicatorShapeTriangle
            .CellNoteIndicatorColor = vbBlue
  
            .CellNote = Trim(rs!LineaX) & vbCrLf & vbCrLf & "Formaliza: " & Format(rs!FechaForp, "dd/mm/yyyy") & vbCrLf _
                       & "Usuario: " & Trim(rs!Userfor) & vbCrLf _
                       & "Oficina: " & rs!OficinaX & vbCrLf & vbCrLf _
                       & "Deductora: " & rs!Deductora & vbCrLf _
                       & "Deduce Planilla: " & rs!ind_deduce_planilla & vbCrLf _
                       & "Factor cálculo: " & rs!Base_Calculo & vbCrLf _
                       & "Divisa: " & rs!Divisa_Desc & vbCrLf & vbCrLf _
                       & "Doc.Ref: " & rs!nDocumento & vbCrLf _
                       & "Canal: " & rs!CanalDesc & vbCrLf _
                       & "Actividad: " & rs!ActividadDesc
        
        Case 5 'Primer Deduccion
            .Text = Format(rs!PriDeduc, "####-##")
        Case 6 'Monto
            .Text = Format(rs!montoapr, "Standard")
        Case 7 'Saldo
            .Text = Format(rs!Saldo, "Standard")
        Case 8 'Cuota
            .Text = Format(rs!Cuota, "Standard")
        Case 9 'Ultimo Movimiento
            '.Text = Format(rs!FecUlt, "####-##")
            .Text = rs!CtaFechaUltCorte & ""
            
        Case 10 'Garantia
            .Text = rs!Garantia
            .TextTip = TextTipFixed
            .TextTipDelay = 1000
            .CellNote = rs!GarantiaDetalle
        
        Case 11 'Fecha Formaliza
            .Text = Format(rs!FechaForp, "yyyy-mm-dd")
        Case 12 'Termina
            .Text = Format((Year(rs!Termina) & Format(Month(rs!Termina), "00")), "####-##")
        Case 13 'Estado
            .Text = rs!Estado
        Case 14 'Proceso
            .Text = rs!Proceso
        Case 15 'Documento
            .Text = rs!nDocumento & ""
        
        Case 16 'Linea Desc
            .Text = rs!LineaX & ""
        
        Case 17 'Referencia
            .Text = rs!Referencia & ""
        Case 18 'Tasa Original
            .Text = Format(rs!TasaOriginal, "Standard")
        Case 19 'Tasa Actual
            .Text = Format(rs!interesv, "Standard")
        Case 20 'Plazo
            .Text = CStr(rs!Plazo)
      
        Case 21 'Cuotas Atrasadas
            .Text = CStr(rs!MoraCuota)
      
        Case 22 'Deductora
            .Text = CStr(rs!Deductora)
      
        Case 23 'IBAN
            .Text = CStr(rs!IBAN)
      
      
      End Select
    Next i
    
     curMonto = curMonto + rs!montoapr
     curSaldo = curSaldo + rs!Saldo
     curCuota = curCuota + IIf(IsNull(rs!Cuota), 0, rs!Cuota)

    rs.MoveNext
  Loop
  rs.Close
  
End With

  
'Totales
txtTotalMonto.Text = Format(curMonto, "Standard")
txtTotalCuota.Text = Format(curCuota, "Standard")
txtTotalSaldo.Text = Format(curSaldo, "Standard")

'Actualiza Etiqueta del nombre con el estado de la mora
'If vEtiquetas Then
    If vMora Then
        lblCreditos.Caption = "Créditos en Mora"
        Set imgCreditos.Picture = imgSemaforos.ListImages.Item(3).Picture
    Else
        lblCreditos.Caption = "Créditos al Día"
        Set imgCreditos.Picture = imgSemaforos.ListImages.Item(1).Picture
    End If
'End If

Me.MousePointer = vbDefault
vPaso = False

End Sub


Private Sub sbSolicitudes(vCedula As String)
Dim i As Integer


On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = "exec spSIFEstadoSolicitud '" & vCedula & "'"
Call OpenRecordSet(rs, strSQL)

With vgCreditos
    .ActiveSheet = 3
    .Sheet = 3
    .MaxRows = 0
    
    Do While Not rs.EOF
    
    .MaxRows = .MaxRows + 1
    .Row = .MaxRows

    For i = 1 To .MaxCols
      .Col = i
      Select Case i
        Case 1 'Status
            .TypePictPicture = imgSemaforos.ListImages.Item(5).Picture
        
        Case 2 'Divisa
            .Text = CStr(rs!CURRENCY_SIM)
            .TextTip = TextTipFixed
            .TextTipDelay = 1000
            .CellNoteIndicatorShape = CellNoteIndicatorShapeTriangle
            .CellNoteIndicatorColor = vbBlue
  
            .CellNote = rs!Divisa_Desc & ""
                     
        Case 3 'Operacion
            .Text = CStr(rs!Id_Solicitud)
        
        Case 4 'Linea
            .Text = CStr(rs!Codigo)
            .TextTip = TextTipFixed
            .TextTipDelay = 1000
            .CellNoteIndicatorShape = CellNoteIndicatorShapeTriangle
            .CellNoteIndicatorColor = vbBlue
  
            .CellNote = "Solicitado: " & Format(rs!FechaSol, "dd/mm/yyyy") & vbCrLf & "Usuario: " & Trim(rs!userRec)
        
        Case 5 'Cédula
            .Text = CStr(rs!Cedula)
        Case 6 'Solicitud
            .Text = Format(rs!FechaSol, "dd/mm/yyyy")
        Case 7 'Monto
            .Text = Format(rs!montosol, "Standard")
        Case 8 'Estado
            Select Case rs!estadosol
             Case "R"
              .Text = "Recibida"
             Case "P"
              .Text = "Pendiente"
             Case "A"
              .Text = "Aprobada"
             Case "D"
              .Text = "Denegada"
             Case "F"
              .Text = "Formalizada"
             Case "N"
              .Text = "Anulada"
            End Select
      
        Case 9 'Descripcion de Linea
            .Text = Trim(rs!LineaX)
      
        Case 10 'Garantia
            .Text = CStr(rs!Garantia)
            .TextTip = TextTipFixed
            .TextTipDelay = 1000
            .CellNoteIndicatorShape = CellNoteIndicatorShapeTriangle
            .CellNoteIndicatorColor = vbBlue
  
            .CellNote = rs!GarantiaDetalle & ""
      
        Case 11 'Usuario Tramite
            .Text = Trim(rs!userRec & "")
      
        Case 12 'Oficina
            .Text = Trim(rs!OficinaX & "")
      
      End Select
     
     
    Next i

     rs.MoveNext
    Loop
    rs.Close

End With
Me.MousePointer = vbDefault

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub


Private Sub sbPreAnalisis(vCedula As String)
Dim i As Integer


On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = "exec spSIFEstadoPreAnalisis '" & vCedula & "'"
Call OpenRecordSet(rs, strSQL)


With vgCreditos
    .ActiveSheet = 4
    .Sheet = 4
    .MaxRows = 0
    
    Do While Not rs.EOF
    
    .MaxRows = .MaxRows + 1
    .Row = .MaxRows

    For i = 1 To .MaxCols
      .Col = i
      Select Case i
        Case 1 'Status
            .TypePictPicture = imgSemaforos.ListImages.Item(4).Picture
             
        Case 2 'Expediente
            .Text = CStr(rs!cod_PreAnalisis)
        
        Case 3 'Tipo
            .Text = CStr(rs!Tipo)
            
        Case 4 'Linea
            .Text = CStr(rs!cod_linea)
        
        Case 5 'Monto
            .Text = Format(rs!Monto, "Standard")
        
        Case 6 'Estado
            .Text = CStr(rs!Estado)
        
        Case 7 'Operacion
            .Text = CStr(rs!Operacion & "")
        
        Case 8 'Fecha
            .Text = CStr(rs!fecha_creacion & "")
        
        Case 9 'Usuario
            .Text = CStr(rs!Usuario & "")
        
      End Select
     
     
    Next i

     rs.MoveNext
    Loop
    rs.Close

End With
Me.MousePointer = vbDefault

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub


Private Sub sbIncobrable(vCedula As String)
Dim i As Integer


On Error GoTo vError

'
'lswDetalle.ColumnHeaders.Add , , "Operación", 1200
'lswDetalle.ColumnHeaders.Add , , "ID", 400
'lswDetalle.ColumnHeaders.Add , , "Línea", 800
'lswDetalle.ColumnHeaders.Add , , "Monto", 1200, vbRightJustify
'lswDetalle.ColumnHeaders.Add , , "Estado", 1200
'lswDetalle.ColumnHeaders.Add , , "Usuario", 1400
'lswDetalle.ColumnHeaders.Add , , "Fecha", 1200
'lswDetalle.ColumnHeaders.Add , , "Documento", 1400
'lswDetalle.ColumnHeaders.Add , , "Notas", 2400
'

Me.MousePointer = vbHourglass

strSQL = "exec spSIFEstadoIncobrable '" & vCedula & "'"
Call OpenRecordSet(rs, strSQL)


With vgCreditos
    .ActiveSheet = 5
    .Sheet = 5
    .MaxRows = 0
    
    Do While Not rs.EOF
    
    .MaxRows = .MaxRows + 1
    .Row = .MaxRows

    For i = 1 To .MaxCols
      .Col = i
      Select Case i
        Case 1 'Status
            .TypePictPicture = imgSemaforos.ListImages.Item(7).Picture
             
        Case 2 'Operacion
            .Text = CStr(rs!Id_Solicitud)
        
        Case 3 'Linea
            .Text = CStr(rs!Codigo)
        
        Case 4 'Monto
            .Text = Format(rs!Saldo + rs!IntCor + rs!IntMor + rs!Cargos + rs!Poliza, "Standard")
        
        Case 5 'Estado
            .Text = CStr(rs!EstadoX)
        
        Case 6 'Usuario
            .Text = CStr(rs!Registro_Usuario & "")
        
        Case 7 'Fecha
            .Text = Format(rs!Registro_Fecha, "dd/mm/yyyy")
        
        Case 8 'Documento
            .Text = "NC." & rs!genera_documento
        
        Case 9 'Notas
            .Text = "[Id.Incobrable: " & rs!cod_incobrable & " ] " & rs!Notas_Registro
        
      End Select
     
     
    Next i

     rs.MoveNext
    Loop
    rs.Close

End With
Me.MousePointer = vbDefault

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub


Private Sub btnAdjuntos_Click()
 gGA.Modulo = "CL_01"
 gGA.Llave_01 = txtCedula.Text
 gGA.Llave_02 = ""
 gGA.Llave_03 = ""
 
 Call sbFormsCall("frmGA_Documentos", vbModal, , , False, Me, True)
End Sub

Private Sub btnBeneficio_Click(Index As Integer)

Select Case Index
    Case 0 'Nuevo
        Call sbClassCall("Beneficios", 5, "frmAF_BeneficioAsg", txtCedula.Text)
        
       
    Case 1 'Consulta
        'Call sbFormsCall("frmAF_CRSeguimiento", , , , False, Me, False)
        
End Select

End Sub

Private Sub btnConsentimiento_Click(Index As Integer)
Dim strSQL As String

On Error GoTo vError

Select Case Index
 Case 0 'Aprobar
 
   If txtConsentimientoFecha.Text <> "" Then
        MsgBox "El consentimiento de uso de información ya fue aprobado anteriormente!", vbExclamation
        Exit Sub
   End If
   
   
    strSQL = "exec spAFI_Persona_Indicadores '" & Trim(txtCedula.Text) & "', '29', 1, '" & glogon.Usuario & "'"
    Call ConectionExecute(strSQL)
    
    Call Bitacora("Aplica", "Firma Consentimiento Informado a Ced." & txtCedula)

   
   Call sbgAFIBitacora("29", "Id.: " & txtCedula.Text & " - " & txtNombre.Text, txtCedula)

   MsgBox "Consentimiento de Uso de Información de Contacto: Aprobado!", vbInformation
      
   Call sbReporteConsentimiento
   Call sbConsulta(txtCedula.Text)
   
 Case 1 'Imprimir
   If txtConsentimientoFecha.Text = "" Then
        MsgBox "No se ha aprobado - El consentimiento de uso de información - ", vbExclamation
        Exit Sub
   End If

    Call sbReporteConsentimiento
 Case 2 'Cerrar
    fraConsentimiento.Visible = False
    tcMain.Visible = True

End Select

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub btnConstancia_Click()


GLOBALES.gTag = txtCedula.Text
GLOBALES.gTag2 = txtNombre.Text
GLOBALES.gTag3 = Format(dtpCorte.Value, "yyyy-mm-dd")

Call sbFormsCall("frmCR_Constancias", 1, , , False, Me)


End Sub

Private Sub btnEC_Mail_Click()

GLOBALES.gTag = txtCedula.Text
GLOBALES.gTag2 = txtNombre.Text
GLOBALES.gTag3 = Format(dtpCorte.Value, "yyyy-mm-dd")

Call sbFormsCall("frmCC_Estado_Cuenta_Mail", 1, , , False, Me)

End Sub

Private Sub btnExport_Click()
Dim vHeaders As vGridHeaders
 
On Error GoTo vError

Me.MousePointer = vbHourglass


ProgressBarX.Visible = True


Select Case tcMain.SelectedItem
    Case 0 'Creditos
        vHeaders.Columnas = vgCreditos.MaxCols
        Select Case vgCreditos.ActiveSheet
            Case 1, 2
                vHeaders.Columnas = 23
                vHeaders.Headers(1) = "Status"
                vHeaders.Headers(2) = "Divisa"
                vHeaders.Headers(3) = "No.Operación"
                vHeaders.Headers(4) = "Código"
                vHeaders.Headers(5) = "Pri.Deduc"
                vHeaders.Headers(6) = "Monto"
                vHeaders.Headers(7) = "Saldo"
                vHeaders.Headers(8) = "Cuota"
                vHeaders.Headers(9) = "Ult.Mov."
                vHeaders.Headers(10) = "Garantía"
                vHeaders.Headers(11) = "Formaliza"
                vHeaders.Headers(12) = "Termina"
                vHeaders.Headers(13) = "Estado"
                vHeaders.Headers(14) = "Proceso"
                vHeaders.Headers(15) = "Documento"
                vHeaders.Headers(16) = "Linea Descripción"
                vHeaders.Headers(17) = "Op.Referencia"
                vHeaders.Headers(18) = "Tasa Original"
                vHeaders.Headers(19) = "Tasa Actual"
                vHeaders.Headers(20) = "Plazo"
                vHeaders.Headers(21) = "Ctas Atrasadas"
                vHeaders.Headers(22) = "Deductora"
                vHeaders.Headers(23) = "IBAN"
            
            Case 3 'Tramite
                vHeaders.Headers(1) = "Status"
                vHeaders.Headers(2) = "Divisa"
                vHeaders.Headers(3) = "No.Operación"
                vHeaders.Headers(4) = "Código"
                vHeaders.Headers(5) = "Cédula"
                vHeaders.Headers(6) = "Fecha Solicitud"
                vHeaders.Headers(7) = "Monto Solicitud"
                vHeaders.Headers(8) = "Estado"
                vHeaders.Headers(9) = "Línea Descripción"
                vHeaders.Headers(10) = "Garantía"
                vHeaders.Headers(11) = "Usuario Tramita"
                vHeaders.Headers(12) = "Oficina"
            
            Case 4 'Estudio
                vHeaders.Headers(1) = "Status"
                vHeaders.Headers(2) = "Expediente"
                vHeaders.Headers(3) = "Tipo"
                vHeaders.Headers(4) = "Línea"
                vHeaders.Headers(5) = "Monto"
                vHeaders.Headers(6) = "Estado"
                vHeaders.Headers(7) = "Asignado"
                vHeaders.Headers(8) = "Fecha"
                vHeaders.Headers(9) = "Usuario"
        
            Case 5 'Incobrables
                vHeaders.Headers(1) = "Status"
                vHeaders.Headers(2) = "Id"
                vHeaders.Headers(3) = "Línea"
                vHeaders.Headers(4) = "Incobrable"
                vHeaders.Headers(5) = "Estado"
                vHeaders.Headers(6) = "Usuario"
                vHeaders.Headers(7) = "Fecha"
                vHeaders.Headers(8) = "Documento"
                vHeaders.Headers(9) = "Notas"
        End Select
        
        Call sbSIFGridExportar(vgCreditos, vHeaders, "ProGrX_" & Trim(txtCedula.Text) & "_Credito_" & vgCreditos.SheetName)
    
    Case 1 'Cobros
        vHeaders.Columnas = vgCobro.MaxCols
        
        Select Case vgCobro.ActiveSheet
            Case 1 'Gestiones
                vHeaders.Headers(1) = "Status"
                vHeaders.Headers(2) = "Fecha Registro"
                vHeaders.Headers(3) = "Vencimiento"
                vHeaders.Headers(4) = "Gestión"
                vHeaders.Headers(5) = "Notas"
                vHeaders.Headers(6) = "Ejecutivo Registro"
                vHeaders.Headers(7) = "Cargo"
                vHeaders.Headers(8) = "Dias para Comisión"
                vHeaders.Headers(9) = "Tipo de Arreglo"
                vHeaders.Headers(10) = "Fecha o Promesa de Pago"
                vHeaders.Headers(11) = "Causa de Morosidad"
            
            Case 2 'Ejecutivos
                vHeaders.Headers(1) = "Fecha asignación"
                vHeaders.Headers(2) = "Oficial Asignado"
                vHeaders.Headers(3) = "Mantiene Asignación"
                vHeaders.Headers(4) = "Aplica Rebajo Doble"
                vHeaders.Headers(5) = "Aplica Cobro Mora Doble"
        End Select
        
        Call sbSIFGridExportar(vgCobro, vHeaders, "ProGrX_" & Trim(txtCedula.Text) & "_Cobros_" & vgCobro.SheetName)
        
    Case 2 'Ahorros
        Call Excel_Exportar_Lsw(lswFND, ProgressBarX)
    
    Case 3 'Patrimonio
    
        vHeaders.Columnas = vgPatrimonio.MaxCols
        Select Case vgPatrimonio.ActiveSheet
            Case 1, 2, 3, 4, 5
                vHeaders.Headers(1) = "Fecha"
                vHeaders.Headers(2) = "Proceso"
                vHeaders.Headers(3) = "Tipo"
                vHeaders.Headers(4) = "Monto"
                vHeaders.Headers(5) = "Tipo Doc."
                vHeaders.Headers(6) = "Documento"
                vHeaders.Headers(7) = "Concepto"
                vHeaders.Headers(8) = "Usuario"
            Case 6 'Excedentes
                vHeaders.Headers(1) = "Status"
                vHeaders.Headers(2) = "Inicio"
                vHeaders.Headers(3) = "Corte"
                vHeaders.Headers(4) = "Excedente Bruto"
                vHeaders.Headers(5) = "Reservas"
                vHeaders.Headers(6) = "Capitalización Ordinaria"
                vHeaders.Headers(7) = "Impuesto Renta"
                vHeaders.Headers(8) = "Otras Deducciones"
                vHeaders.Headers(9) = "Excedente a Pagar"
        
        End Select
    
        Call sbSIFGridExportar(vgPatrimonio, vHeaders, "ProGrX_" & Trim(txtCedula.Text) & "_Patrimonio_" & vgPatrimonio.SheetName)
    
    Case 4 'Mensajes
    
        vHeaders.Columnas = vGrid.MaxCols
        vHeaders.Headers(1) = "Vencimiento"
        vHeaders.Headers(2) = "Mensaje"
        vHeaders.Headers(3) = "Seleccionar"
        vHeaders.Headers(4) = "Registro Fecha"
        vHeaders.Headers(5) = "Registro Usuario"
        vHeaders.Headers(6) = "Resolución Fecha"
        vHeaders.Headers(7) = "Resolución Usuario"
       
        Call sbSIFGridExportar(vGrid, vHeaders, "ProGrX_" & Trim(txtCedula.Text) & "_Mensajes_" & vGrid.SheetName)
    
    
    Case 5 'Info
        Call Excel_Exportar_Lsw(lswDP, ProgressBarX)
    
    Case 6 'Correos
        Call Excel_Exportar_Lsw(lswCorreo, ProgressBarX)
    
    Case 7 'Beneficios
        Call Excel_Exportar_Lsw(lswBeneficios, ProgressBarX)
    
    Case 8 'Renuncias
        Call Excel_Exportar_Lsw(lswRenuncias, ProgressBarX)
    
    Case 9 'SOS
        Call Excel_Exportar_Lsw(lswSoS, ProgressBarX)

End Select

ProgressBarX.Visible = False

Me.MousePointer = vbDefault

Exit Sub

vError:
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical
 
End Sub

Private Sub btnFondos_Click(Index As Integer)
Dim vCajas As Boolean, sForm As String, frm As Form

On Error GoTo vError


If txtCedula.Text = "" Or txtNombre.Text = "" Then Exit Sub

vCajas = IIf((fxCajasParametros("03") = "S"), True, False)


Select Case Index
 
  Case 0 'Estado
    
    
    Call sbFondoEstado
  
  Case 1 'Consulta Avanzada
        Call sbFormsCall("frmFNDConsultaContratos", , , , False, Me)
  
         For Each frm In Forms
          If UCase(frm.Name) = UCase("frmFNDConsultaContratos") Then
            Call frm.sbConsultaExterna(txtCedula.Text)
            Exit For
          End If
        Next frm
  
  Case 2 'Registro de Contrato
  
        Call sbFormsCall("frmFNDContratos", , , , False, Me)
        
        If pContrato > 0 Then
             For Each frm In Forms
              If UCase(frm.Name) = UCase("frmFNDContratos") Then
                Call frm.sbConsultaExterna(CLng(pOperadora), pPlan, CLng(pContrato))
                Exit For
              End If
            Next frm
        
        Else
        
             For Each frm In Forms
              If UCase(frm.Name) = UCase("frmFNDContratos") Then
                Call frm.sbConsultaExternaNuevoCnt(txtCedula.Text, txtNombre.Text)
                Exit For
              End If
            Next frm

        End If
    
  Case 3 'Aportes
    'Si existe contrato procede
    If pContrato > 0 Then
        gFondos.Operadora = pOperadora
        gFondos.Plan = pPlan
        gFondos.Contrato = pContrato
        gFondos.Cedula = txtCedula.Text
      
        If vCajas Then
           sForm = "frmCajas_FNDAportaciones"
        Else
            sForm = "frmFNDAportaciones"
        End If
      
        Call sbFormsCall(sForm, vbModal, , , False, Me)
   
    Else
        
            MsgBox "Selecciones un Contrato!", vbExclamation
    End If
  
  
  Case 4 'Retiros
    If pContrato > 0 Then
        gFondos.Operadora = pOperadora
        gFondos.Plan = pPlan
        gFondos.Contrato = pContrato
        gFondos.Cedula = txtCedula.Text
      
        Call sbFormsCall("frmFNDRetirosyLiquidaciones", vbModal, , , False, Me)
     Else
    
        MsgBox "Selecciones un Contrato!", vbExclamation
   End If
  
  Case 5 'Calculadora
        GLOBALES.gCedulaActual = txtCedula.Text
        Call sbFormsCall("frmFnd_Calculadora_Inversiones", , , , False, Me)
    
End Select

vError:

End Sub

Private Sub btnFondos_Export_Click()

On Error GoTo vError

Me.MousePointer = vbHourglass

Call Excel_Exportar_Lsw(lswFnd_List, ProgressBarX)

Me.MousePointer = vbDefault

Exit Sub

vError:
    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub btnFondos_List_Click(Index As Integer)


On Error GoTo vError

Me.MousePointer = vbHourglass

Dim i As Integer


For i = 0 To btnFondos_List.Count - 1
    btnFondos_List(i).Checked = False
Next i


btnFondos_List(Index).Checked = True

Select Case Index
    Case 0 'Movimientos
        Call sbFnd_Contratos_Movimientos(pOperadora, pPlan, pContrato, lswFnd_List)
    
    Case 1 'Cupones
        Call sbFnd_Contratos_Cupones(pOperadora, pPlan, pContrato, lswFnd_List)
    
    Case 2 'Bitacora
        Call sbFnd_Contratos_Bitacora(pOperadora, pPlan, pContrato, lswFnd_List)
    
    Case 3 'Cierres
        Call sbFnd_Contratos_Cierres(pOperadora, pPlan, pContrato, lswFnd_List)
    
End Select



Me.MousePointer = vbDefault

Exit Sub

vError:
    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub btnIdentificarDP_Click()

On Error GoTo vError

If txtCedula.Text <> "" Then
    GLOBALES.gTag = txtCedula.Text
    GLOBALES.gTag2 = txtNombre.Text
    Call sbFormsCall("frmCajas_IdentificaSF", vbModal, , , False, Me)
    Call sbConsulta(txtCedula.Text)
    
    strSQL = "select dbo.fxCajas_SaldoaFavor('" & txtCedula.Text & "') as 'Cajas_Saldo_Favor'"
    Call OpenRecordSet(rs, strSQL)
    btnSaldosFavor.Caption = Format(rs!Cajas_Saldo_Favor, "Standard")
    rs.Close
End If

vError:

End Sub

Private Sub btnInfoTrigger_Click(Index As Integer)
Dim frm As Form

GLOBALES.gCedulaActual = Trim(txtCedula)

Select Case Index
  Case 0 'Contacto
        Call sbFormsCall("frmCR_VerificaDatosPersonales", 1, , , False, Me)
        Call btnInfo_Click(0)
     
  Case 1 'Laboral
        Call sbFormsCall("frmCR_VerificaDatosPersonales", 1, , , False, Me)
  
  Case 2 'Teléfonos
        Call sbFormsCall("frmAF_Telefonos", 1, , , False, Me)
        Call btnInfo_Click(2)
  
  Case 3 'Cuentas
        GLOBALES.gTag = Trim(txtCedula.Text)
        GLOBALES.gTag2 = "AFI"
        frmCC_Cuentas_Bancarias.Show vbModal
        Call btnInfo_Click(3)
        
  Case 4 'Ingresos"
        Call sbFormsCall("frmAF_ConsultaMov", 0, , , False)
         For Each frm In Forms
          If UCase(frm.Name) = UCase("frmAF_ConsultaMov") Then
            Call frm.sbConsulaExterna(txtCedula.Text)
            Exit For
          End If
        Next frm
        
  Case 5 'Liquidaciones
        Call sbFormsCall("frmAF_ConsultaMov", 0, , , False)
         
         For Each frm In Forms
          If UCase(frm.Name) = UCase("frmAF_ConsultaMov") Then
            Call frm.sbConsulaExterna(txtCedula.Text)
            Exit For
          End If
        Next frm
  
  Case 6 'Beneficiarios
       Call sbFormsCall("frmAF_Beneficiarios", 1, , , False, Me)
        Call btnInfo_Click(6)

  Case 11 'Beneficiarios de Polizas Colectivas
        GLOBALES.gTag = Trim(txtCedula.Text)
        GLOBALES.gTag2 = txtNombre
        frmCC_Poliza_Beneficiarios.Show vbModal
        Call btnInfo_Click(11)
        
End Select

End Sub

Private Sub btnNotas_Click(Index As Integer)

On Error GoTo vError

Select Case Index
    
  Case 0 'GuardaNota
        If Len(txtNotas.Text) < 10 Then
          MsgBox "No se especificó ninguna nota a registrar!", vbExclamation
          Exit Sub
        End If
        
        strSQL = "update socios set notas = '" & UCase(Trim(txtNotas)) & "',Nota_User = '" _
               & glogon.Usuario & "',Nota_Fecha = dbo.MyGetdate()" _
               & " where cedula = '" & txtCedula & "'"
        'Registra En el Pool de Mensajes Generales
        strSQL = strSQL & Space(10) & "insert socios_mensajes(fecha,cedula,usuario,vencimiento,mensaje,tipo) values(dbo.MyGetdate(),'" _
                & txtCedula & "','" & glogon.Usuario & "','2100/01/01','" _
                & Trim(txtNotas) & "','G')"
        Call ConectionExecute(strSQL)

        
        MsgBox "NOTA REGISTRADA...", vbInformation
  
  Case 1 'Bloqueo
    Call sbCrdBloqueoCreditos(txtCedula.Text, 1)

End Select

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical


End Sub

Private Sub btnPatrimonioConsulta_Click()
 Dim frm As Form
 
 Call sbFormsCall("frmAH_Principal", , , , False, Me, True)

 Call sbFormActivo("frmAH_Principal", frm)
 Call frm.sbConsulta_Externa(txtCedula)

End Sub

Private Sub btnRenuncia_Click(Index As Integer)

Select Case Index
    Case 0 'Nuevo
        Dim frm As Form
        
        Call sbFormsCall("frmAF_CRRenuncia", , , , False, Me, True)
       
        Call sbFormActivo("frmAF_CRRenuncia", frm)
        Call frm.sbConsulta_Externa_Cedula(txtCedula)
       
    Case 1 'Consulta
        Call sbFormsCall("frmAF_CRSeguimiento", , , , False, Me, False)
        
End Select

End Sub

Private Sub btnSaldosFavor_Click()

If btnSaldosFavor.Caption > 0 Then
    ModuloCajas.mClienteId = txtCedula.Text
    ModuloCajas.mCliente = txtNombre.Text
    Call sbFormsCall("frmCajas_TransacSFLiq", vbModal, , , False, Me)
    Call sbConsulta(txtCedula.Text)
End If

End Sub

Private Sub btnSoS_Click()
Dim i As Integer

If Trim(GLOBALES.gTag) = "" Then
  Exit Sub
End If


GLOBALES.gTag = txtCedula.Text

fraConsentimiento.Visible = False
tcMain.Visible = True

For i = 0 To isButton1.Count - 1
   isButton1.Item(i).Checked = False
Next i

tcMain.Item(9).Selected = True


Call sbSoS_Resumen(txtCedula.Text)

End Sub

Private Sub btnSoS_Export_Click(Index As Integer)

On Error GoTo vError

Me.MousePointer = vbHourglass


Select Case Index
    Case 0 'Resumen de Pagos
        Call Excel_Exportar_Lsw(lswSoS)
    Case 1 'Movimientos
        Call Excel_Exportar_Lsw(lswSoS_Det)
End Select

Me.MousePointer = vbDefault

Exit Sub

vError:
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub cboMsj_Click()
Call sbCargaMsj(txtCedula.Text)

If Mid(cboMsj.Text, 1, 1) = "G" Then
    chkMsjVence.Enabled = True
    chkMsjVence.Value = xtpUnchecked
    imgMsjNuevo.Visible = True
Else
    chkMsjVence.Enabled = False
    chkMsjVence.Value = xtpChecked
    imgMsjNuevo.Visible = False
End If

Call chkMsjVence_Click

End Sub

Private Sub sbFondoEstado()

Me.MousePointer = vbHourglass

With frmContenedor.Crt
  .Reset
  .WindowShowPrintSetupBtn = True
  .WindowShowExportBtn = True
  .WindowShowRefreshBtn = True
  .WindowShowSearchBtn = True
  .WindowShowZoomCtl = True
  
  .Connect = glogon.ConectRPT
  
  .WindowTitle = "Fondos de Ahorros e Inversiones"
  .WindowState = crptMaximized
  
  .Formulas(0) = "Fecha='" & Format(fxFechaServidor, "dd/mm/yyyy") & "'"
  .Formulas(1) = "Usuario='" & Trim(glogon.Usuario) & "'"
  .Formulas(2) = "Empresa='" & Trim(GLOBALES.gstrNombreEmpresa) & "'"
  
If chkFndContrato.Value = vbChecked And pContrato > 0 Then
    strSQL = "{FND_CONTRATOS.COD_OPERADORA} =" & pOperadora & "And "
    strSQL = strSQL & "{FND_CONTRATOS.COD_PLAN} ='" & pPlan & "' and "
    strSQL = strSQL & "{FND_CONTRATOS.COD_CONTRATO} = " & pContrato
       
      .ReportFileName = SIFGlobal.fxPathReportes("Fondos_EstadoDetallado.rpt")
      
      .Formulas(3) = "SubTitulo='" & Format(fxFechaServidor, "yyyy/mm/dd") & "'"
      .SelectionFormula = strSQL

Else

      .ReportFileName = SIFGlobal.fxPathReportes("Fondos_EstadoConsolidado.rpt")
      .Formulas(3) = "SubTitulo=' Reporte al " & Format(fxFechaServidor, "yyyy/mm/dd") & "'"
      .SelectionFormula = "{SOCIOS.CEDULA} ='" & Trim(txtCedula) & "'"
  
End If
  .PrintReport
End With


Me.MousePointer = vbDefault

End Sub


Private Sub cboPat_Garantia_Click()
If vPaso Then Exit Sub

On Error GoTo vError

strSQL = "select dbo.fxCrdGarantiaPatMnt(S.cedula, '" & cboPat_Garantia.ItemData(cboPat_Garantia.ListIndex) & "','M') as 'Pat_Garantia_Total'" _
      & " , dbo.fxCrdGarantiaPatMnt(S.cedula, '" & cboPat_Garantia.ItemData(cboPat_Garantia.ListIndex) & "','S') " _
      & " + dbo.fxCrdGarantiaPatMnt_SldTramite(S.cedula,'A')" _
      & " as 'Pat_Garantia_Saldos'" _
      & " from Socios S where S.cedula = '" & txtCedula.Text & "'"
   
Call OpenRecordSet(rs, strSQL)

   txtPAT_Disponible.Text = Format(rs!Pat_Garantia_Total, "Standard")
   txtPAT_Saldos.Text = Format(rs!Pat_Garantia_Saldos, "Standard")
   txtPAT_Saldos.Tag = rs!Pat_Garantia_Saldos

rs.Close

Call cboPAT_TipoSaldo_Click
Call txtPAT_Saldos_Change

vError:

End Sub

Private Sub cboPAT_TipoSaldo_Click()

On Error GoTo vError

lblPAT_Saldo.Caption = "(-) " & cboPAT_TipoSaldo.Text

If cboPAT_TipoSaldo.Text = "Saldos en Garantía" Then
    txtPAT_Saldos.Text = Format(CCur(txtPAT_Saldos.Tag), "Standard")
Else
    txtPAT_Saldos.Text = Format(CCur(StatusBar.Panels(6).Text), "Standard")
End If

vError:

End Sub



Private Sub chkMsjVence_Click()
If chkMsjVence.Value = xtpChecked Then
    dtpMsjVence.Enabled = False
    dtpMsjVence.Value = DateAdd("m", 50 * 12, mFecha)
Else
    dtpMsjVence.Enabled = True
    dtpMsjVence.Value = DateAdd("m", 1, mFecha)
End If
End Sub

Private Sub chkSoS_Exclusion_Click()

If vPaso Then Exit Sub

Dim pMovimiento As String, pDetalle As String

On Error GoTo vError

Me.MousePointer = vbHourglass

pDetalle = "Exclusión del Programa SOS -> Cédula: " & txtCedula.Text

If chkSoS_Exclusion.Value = xtpChecked Then
   pMovimiento = "Registra"
   strSQL = "exec spSOS_Exclusiones_Registro '" & Trim(txtCedula.Text) & "','A','" & glogon.Usuario & "'"
Else
   pMovimiento = "Elimina"
   strSQL = "exec spSOS_Exclusiones_Registro '" & Trim(txtCedula.Text) & "','I','" & glogon.Usuario & "'"
End If

Call ConectionExecute(strSQL)
Call Bitacora(pMovimiento, pDetalle)


Me.MousePointer = vbDefault

Exit Sub

vError:
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub dtpCorte_Change()
Dim i As Integer

StatusBar.Panels(5).Text = "0.00"
StatusBar.Panels(5).Tag = ""
StatusBar.Panels(5).ToolTipText = "Cuota...: "
StatusBar.Panels(6).Text = "0.00"
StatusBar.Panels(7).Text = "0.00"

With vgCreditos
    .Sheet = 1

    For i = 1 To .MaxRows
      .Row = i
      .Col = 2
      If .Value = vbChecked Then
         Call vgCreditos_ButtonClicked(2, i, 0)
      End If
    Next i
End With

End Sub


Private Sub Form_Load()
'Tiene Menú Interno

vModulo = 3

'cbMenuPopUp.Item(1).Visible = False

lblMembresia.BackStyle = 0


tcMain.Item(0).Visible = False
tcMain.Item(1).Visible = False
tcMain.Item(2).Visible = False
tcMain.Item(3).Visible = False
tcMain.Item(4).Visible = False
tcMain.Item(5).Visible = False
tcMain.Item(6).Visible = False

tcMain.Item(7).Visible = False
tcMain.Item(8).Visible = False


tcMain.Item(0).Selected = True


'    Dim Control As CommandBarControl
'    Dim ControlFile As CommandBarPopup
'
'    CommandBars.ActiveMenuBar.Visible = False
'
'    Set ControlFile = CommandBars.ActiveMenuBar.Controls.Add(xtpControlPopup, 0, "&File", -1, False)
'    With ControlFile.CommandBar.Controls
'        .Add xtpControlButton, ID_Menu_Abonos, "&Abonos", -1, False
'        .Add xtpControlButton, ID_Menu_Anula, "An&ula", -1, False
''
''        Set Control = .Add(xtpControlButton, ID_FILE_PRINT, "&Print", -1, False)
''        Control.BeginGroup = True
''        .Add xtpControlButton, ID_FILE_PRINT_SETUP, "Print Set&up...", -1, False
''
''        Set Control = .Add(xtpControlButton, ID_FILE_EXIT, "&Exit", -1, False)
''        Control.BeginGroup = True
'    End With


mFecha = fxFechaServidor

vGrid.AppearanceStyle = fxGridStyle
vGrid.MaxRows = 5

Set imgBanner.Picture = frmContenedor.imgBanner_01.Picture

cboMsj.AddItem "Generales"
cboMsj.AddItem "Bloqueos / Desbloqueo"
'cboMsj.AddItem "Pendientes"
'cboMsj.AddItem "Morosidad"
'cboMsj.AddItem "Resueltos (Pendientes)"

cboPAT_TipoSaldo.AddItem "Saldos en Garantía"
cboPAT_TipoSaldo.AddItem "Saldos Refinanciar"
cboPAT_TipoSaldo.Text = "Saldos en Garantía"


vPaso = True

    strSQL = "select GARANTIA as 'IdX', rtrim(DESCRIPCION) as 'ItmX'" _
           & " from CRD_GARANTIA_TIPOS" _
           & " where FORMULARIO = 'F01' order by Garantia"
    
    Call sbCbo_Llena_New(cboPat_Garantia, strSQL, False, True)

vPaso = False


vgCreditos.ActiveSheet = 1
vgCreditos.Sheet = 1
vgCreditos.MaxRows = 0

Call Formularios(Me)
Call RefrescaTags(Me)

Call EstadoInicial

Call isButton1_Click(0)


If btnNotas.Item(0).Enabled Then txtNotas.Locked = False
btnNotas.Item(1).Enabled = btnNotas.Item(0).Enabled

StatusBar.Panels(4).Text = glogon.Usuario
StatusBar.Panels(5).Text = "0.00"
StatusBar.Panels(6).Text = "0.00"
StatusBar.Panels(7).Text = "0.00"

dtpCorte.Value = fxFechaServidor
dtpCorte.MinDate = dtpCorte.Value

Me.Width = 13935
Me.Height = 8700


End Sub

Private Sub sbBusqueda()

On Error GoTo vError

gBusquedas.Convertir = "N"

    Call Limpia
    gBusquedas.Col1Name = "Identificación"
    gBusquedas.Col2Name = "Id Alterno"
    gBusquedas.Col3Name = "Nombre"
    gBusquedas.Consulta = "Select cedula,cedular,nombre from SOCIOS"
    gBusquedas.Columna = "nombre"
    gBusquedas.Orden = "nombre"
    frmBusquedas.Show vbModal
    txtCedula = Trim(gBusquedas.Resultado)
    gBusquedas.Consulta = ""
    gBusquedas.Columna = ""
    gBusquedas.Orden = ""
    gBusquedas.Resultado = ""
    If Trim(txtCedula) <> "" Then
        Call sbConsulta(txtCedula)
    End If

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical


End Sub



Private Sub Form_Resize()
Dim pLeft As Long, pHeight As Long
On Error Resume Next

imgBanner.Width = Me.Width
ShortcutCaptionTitle.Width = Me.Width

pLeft = 300
pHeight = 770

tcMain.Width = Me.Width - 150
tcMain.Height = Me.Height - (tcMain.top + 300)

ProgressBarX.Width = tcMain.Width

fraConsentimiento.Left = (tcMain.Width - fraConsentimiento.Width) / 2

vgCreditos.Width = tcMain.Width - (vgCreditos.Left + pLeft)
vgCreditos.Height = tcMain.Height - (gbNotas.Height + 1200)

txtTotalMonto.top = vgCreditos.top + vgCreditos.Height + 70
Label1.Item(0).top = txtTotalMonto.top
Label1.Item(1).top = txtTotalMonto.top
Label1.Item(2).top = txtTotalMonto.top
Label1.Item(3).top = txtTotalMonto.top

txtTotalCuota.top = txtTotalMonto.top
txtTotalSaldo.top = txtTotalMonto.top
dtpCorte.top = txtTotalMonto.top
btnConstancia.top = txtTotalMonto.top


gbNotas.top = vgCreditos.top + vgCreditos.Height + 400
gbNotas.Width = tcMain.Width - pLeft

txtNotas.Width = vgCreditos.Width - 450

'Label4.Item(2).top = Me.Height - 1500
'Label4.Item(2).Left = txtNotas.Left - Label4.Item(2).Width + 20

lblBloqueo.Left = txtNotas.Left + txtNotas.Width + 50
btnNotas.Item(0).Left = lblBloqueo.Left + lblBloqueo.Width
btnNotas.Item(1).Left = lblBloqueo.Left + lblBloqueo.Width

vgPatrimonio.Height = tcMain.Height - pHeight
vgPatrimonio.Width = tcMain.Width - (vgPatrimonio.Left + pLeft)

lswFND.Width = tcMain.Width - pLeft
lswFnd_List.Width = tcMain.Width - pLeft

lswFnd_List.Height = tcMain.Height - (lswFnd_List.top + 900 + gbFndContrato.Height)  ' (lswFND.Height + 900 + gbFndContrato.Height)

lswFnd_List.top = btnFondos_List(0).top + btnFondos_List(0).Height + 80

gbFndContrato.top = lswFnd_List.top + lswFnd_List.Height + 50
gbFndContrato.Width = lswFND.Width


'Mensajes

vGrid.Height = tcMain.Height - pHeight
vGrid.Width = tcMain.Width - (vGrid.Left + pLeft)

fraMsj.Width = vGrid.Width
fraMsj.Height = vGrid.Height


vgCobro.Height = tcMain.Height - pHeight
vgCobro.Width = tcMain.Width - (vgCobro.Left + pLeft)
 

lswDP.Width = tcMain.Width - (lswDP.Left + pLeft)
lswDP.Height = tcMain.Height - pHeight

'Correo

lswCorreo.Width = tcMain.Width - (lswCorreo.Left + 100)
scCorreo.Width = lswCorreo.Width

lswCorreo.Height = tcMain.Height - (lswCorreo.top + 250)


'Beneficios

lswBeneficios.Width = tcMain.Width - (lswBeneficios.Left + 100)
lswBeneficios.Height = tcMain.Height - (lswBeneficios.top + gbBeneficios.Height + 750)

gbBeneficios.top = lswBeneficios.top + lswBeneficios.Height + 150

gbBeneficios.Left = lswBeneficios.Left
gbBeneficios.Width = lswBeneficios.Width

'Renuncias

lswRenuncias.Width = tcMain.Width - (lswRenuncias.Left + 100)
lswRenuncias.Height = tcMain.Height - (lswRenuncias.top + gbRenuncias.Height + 750)

gbRenuncias.top = lswRenuncias.top + lswRenuncias.Height + 150

gbRenuncias.Left = lswRenuncias.Left
gbRenuncias.Width = lswRenuncias.Width

'SOS
lswSoS.Width = tcMain.Width - pLeft
lswSoS_Det.Width = tcMain.Width - pLeft

lswSoS_Det.Height = tcMain.Height - (lswSoS_Det.top + 900)

End Sub




Private Sub imgBorraMsj_Click()
Dim i As Integer
Dim msj(2) As String

On Error GoTo vError

With vGrid
    For i = 1 To vGrid.MaxRows
      .Row = i
      .Col = 3
        If .Value = 1 Then
           .Col = 1
           msj(0) = Format(.Text, "yyyy/mm/dd")
           msj(1) = .CellTag
           .Col = 2
           msj(2) = Mid(.Text, 1, 15)
           
           strSQL = "delete socios_mensajes where cedula = '" & txtCedula _
                  & "' and usuario = '" & msj(1) & "' and vencimiento = '" _
                  & msj(0) & "' and substring(mensaje,1,15) = '" _
                  & msj(2) & "'"
           Call ConectionExecute(strSQL)
        End If
    Next i
End With

Call sbCargaMsj(txtCedula)

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbExclamation

End Sub





Private Sub btnConstanciaAportes_Click()

GLOBALES.gTag = txtCedula.Text
GLOBALES.gTag2 = txtNombre.Text

Call sbFormsCall("frmAH_Constancias", 1, , , False, Me)

End Sub


Private Sub imgGuardaMsj_Click()
Dim vTipo As String

On Error GoTo vError

txtMsj.Text = fxSysCleanTxtInject(txtMsj.Text)

If Len(txtMsj.Text) < 50 Then
   MsgBox "El mensaje no contiene la cantidad mínima de caracteres (50)o se encuentra en blanco, por favor introduzca más detalle para guardar.", vbExclamation
   Exit Sub
End If

Select Case True
  Case optMsj.Item(0).Value 'General
       vTipo = "G"
  Case optMsj.Item(1).Value 'Pendiente
       vTipo = "P"
  Case optMsj.Item(2).Value 'Morosidad
       vTipo = "M"
  Case Else
       vTipo = "G"
End Select

strSQL = "insert socios_mensajes(fecha,cedula,usuario,vencimiento,mensaje,Tipo) values(dbo.MyGetdate(),'" _
       & txtCedula & "','" & glogon.Usuario & "','" & Format(dtpMsjVence.Value, "yyyy/mm/dd") & "','" _
       & txtMsj & "','" & vTipo & "')"
Call ConectionExecute(strSQL)

txtMsj = ""
fraMsj.Visible = False
MsgBox "Mensaje Registrado...", vbInformation

Call sbCargaMsj(txtCedula)

Exit Sub
vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub imgMsjCierraFrame_Click()
fraMsj.Visible = False
End Sub

Private Sub imgMsjNuevo_Click()
fraMsj.Visible = True
fraMsj.Width = vGrid.Width
fraMsj.Left = vGrid.Left

chkMsjVence.Value = xtpChecked
Call chkMsjVence_Click

End Sub

Private Sub imgMsjResuelve_Click()
Dim i As Integer
Dim msj(2) As String

On Error GoTo vError

With vGrid
    For i = 1 To vGrid.MaxRows
      .Row = i
      .Col = 3
        If .Value = 1 Then
           .Col = 1
           msj(0) = Format(.Text, "yyyy/mm/dd")
           msj(1) = .CellTag
           .Col = 2
           msj(2) = Mid(.Text, 1, 15)
           
           strSQL = "update socios_mensajes set Resolucion = 'R', Resolucion_Fecha = dbo.MyGetdate()" _
                  & ", Resolucion_Usuario = '" & glogon.Usuario & "'" _
                  & " where cedula = '" & txtCedula _
                  & "' and usuario = '" & msj(1) & "' and vencimiento = '" _
                  & msj(0) & "' and substring(mensaje,1,15) = '" _
                  & msj(2) & "'"
           Call ConectionExecute(strSQL)
        End If
    Next i
End With

Call sbCargaMsj(txtCedula)

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbExclamation

End Sub



Private Sub sbBeneficios_Load()

On Error GoTo vError


Dim curTotal As Currency

Me.MousePointer = vbHourglass


lswBeneficios.ListItems.Clear

With lswBeneficios.ColumnHeaders
    .Clear
    .Add , , "Id Beneficio", 1200
    .Add , , "Código", 1100, vbCenter
    .Add , , "Beneficio", 3200
    .Add , , "Fecha", 1200, vbCenter
    .Add , , "Estado", 1200, vbCenter
    .Add , , "Tipo", 1200, vbCenter
    .Add , , "Monto", 1400, vbRightJustify
    
    .Add , , "Usuario", 2100
    .Add , , "Notas", 3200
    
    .Add , , "Remesa Id", 1800, vbCenter
    .Add , , "Remesa Fecha", 1800, vbCenter
    .Add , , "Remesa Estado", 1800, vbCenter


    .Add , , "Sol.Cédula", 2100, vbCenter
    .Add , , "Sol.Nombre", 3200
End With

curTotal = 0

strSQL = "exec spAFI_Beneficios_Consulta '" & txtCedula.Text & "'"
Call OpenRecordSet(rs, strSQL)


Do While Not rs.EOF
 Set itmX = lswBeneficios.ListItems.Add(, , rs!consec)
     itmX.SubItems(1) = rs!cod_beneficio
     itmX.SubItems(2) = rs!Beneficio_Desc
     itmX.SubItems(3) = Format(rs!Registra_Fecha, "yyyy-mm-dd")
     itmX.SubItems(4) = rs!Estado_Desc
     itmX.SubItems(5) = rs!Tipo_Benefico
     itmX.SubItems(6) = Format(rs!Monto, "Standard")
     itmX.SubItems(7) = rs!REGISTRA_USER & ""
     itmX.SubItems(8) = rs!Notas & ""
     itmX.SubItems(9) = rs!cod_remesa & ""
     itmX.SubItems(10) = rs!Remesa_Fecha & ""
     itmX.SubItems(11) = rs!Remesa_Estado & ""
     itmX.SubItems(12) = rs!Sol_Cedula & ""
     itmX.SubItems(13) = rs!Sol_Nombre & ""
     
     If rs!Estado = "E" Then
         curTotal = curTotal + rs!Monto
     End If
 rs.MoveNext
Loop
rs.Close

txtBeneCasos.Text = lswBeneficios.ListItems.Count
txtBeneTotal.Text = Format(curTotal, "Standard")

Me.MousePointer = vbDefault

Exit Sub

vError:
    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub sbRenuncias_Load()

On Error GoTo vError

Me.MousePointer = vbHourglass

With lswRenuncias.ColumnHeaders
    .Clear
    .Add , , "Id Renuncia", 1200
    .Add , , "Causa", 3200
    .Add , , "Fecha", 1200, vbCenter
    .Add , , "Vencimiento", 1200, vbCenter
    .Add , , "Estado", 1200, vbCenter
    .Add , , "Tipo", 1200, vbCenter
    .Add , , "Desea Volver", 1200, vbCenter
    .Add , , "Apl. Reingreso", 1200, vbCenter
    .Add , , "Ejecutivo", 3200
    
    .Add , , "Notas", 3200
    
    .Add , , "Empresa", 2200
    .Add , , "Provincia", 2200
    .Add , , "Email", 2200
    
    .Add , , "Liq.Id", 1100, vbCenter
    .Add , , "Usuario", 2100

    
End With

strSQL = "exec spAFI_Renuncias_Consulta '" & txtCedula.Text & "'"
Call OpenRecordSet(rs, strSQL)

lswRenuncias.ListItems.Clear

Do While Not rs.EOF
 Set itmX = lswRenuncias.ListItems.Add(, , rs!Cod_Renuncia)
     itmX.SubItems(1) = rs!Causa_Desc
     itmX.SubItems(2) = Format(rs!Registro_Fecha, "yyyy-mm-dd")
     itmX.SubItems(3) = Format(rs!Vencimiento, "yyyy-mm-dd")
     itmX.SubItems(4) = rs!Estado_Desc
     itmX.SubItems(5) = rs!Tipo_Renuncia
     itmX.SubItems(6) = rs!Desea_Volver
     itmX.SubItems(7) = rs!Aplica_Reingreso
     itmX.SubItems(8) = rs!Ejecutivo_Desc
     itmX.SubItems(9) = rs!Notas & ""
     itmX.SubItems(10) = rs!Institucion_Desc
     itmX.SubItems(11) = rs!Provincia_Desc
     itmX.SubItems(12) = rs!Email & ""
     itmX.SubItems(13) = rs!Liquida_Id & ""
     itmX.SubItems(14) = rs!Registro_Usuario & ""
     
 rs.MoveNext
Loop
rs.Close

txtRenunciasCasos.Text = lswRenuncias.ListItems.Count

Me.MousePointer = vbDefault

Exit Sub

vError:
    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub



Private Sub isButton1_Click(Index As Integer)
Dim i As Integer


GLOBALES.gTag = txtCedula.Text

fraConsentimiento.Visible = False
tcMain.Visible = True

For i = 0 To isButton1.Count - 1
   isButton1.Item(i).Checked = False
Next i
isButton1.Item(Index).Checked = True


If Trim(GLOBALES.gTag) = "" Then
  Exit Sub
End If



If Index < 6 Then
    tcMain.Item(Index).Selected = True
End If

Select Case Index
  Case 8
    tcMain.Item(6).Selected = True

  Case 9 'Beneficios
    tcMain.Item(7).Selected = True
  Case 10 'Renuncias
    tcMain.Item(8).Selected = True

End Select

Select Case isButton1.Item(Index).Caption
  
  Case "Créditos"
    
  Case "Beneficios"
    Call sbBeneficios_Load
    
  Case "Renuncias"
    Call sbRenuncias_Load
    
  Case "Cobros"
    Call vgCobro_SheetChanged(1, 1)
  
  Case "Ahorros"
    Call sbFondos(txtCedula)
  
  Case "Patrimonio"
    Call vgPatrimonio_SheetChanged(1, 1)
    Call cboPAT_TipoSaldo_Click
    Call cboPat_Garantia_Click
  
  Case "Mensajes"
    Call sbCargaMsj(txtCedula)
  
  Case "Info"
    Call btnInfo_Click(0)
  
  Case "Correo"
    Call sbCorreo(txtCedula)
    
  Case "Estado"
         Call sbEstadoCuenta(txtCedula)

  Case "Aut/C.I."
     tcMain.Visible = False
     fraConsentimiento.top = tcMain.top
     fraConsentimiento.Visible = True

End Select

End Sub

Private Sub btnCredito_Click(Index As Integer)


GLOBALES.gTag = Trim(txtCedula.Text)

If Trim(txtCedula.Text) = "" Then
  Exit Sub
End If

'     Call sbFormsCall("frmCR_ConsultaPlanFidelidad", 1, 0, 0, False, Me)
'     fraConsentimiento.Left = 1560
'     fraConsentimiento.Top = 720
'     fraConsentimiento.Visible = True


Select Case btnCredito.Item(Index).Caption
  Case "Nuevo"
    Call MDIPrincipal.mnuAccionesSub_Click(6)
  
  Case "Crd. Exc."
     Call sbFormsCall("frmCR_ConsultaCrdExc", 1, 0, 0, False, Me)
     Call vgCreditos_SheetChanged(1, 1)
  
  Case "Cálculo"
        GLOBALES.gCedulaActual = txtCedula.Text
        Call sbFormsCall("frmCR_CalculoOperacion", 0)
  
  Case "Estudio Crédito"
    Call MDIPrincipal.mnuAccionesSub_Click(11)
  
  Case "Fianzas"
     GLOBALES.gCedulaActual = txtCedula.Text
     Call sbFormsCall("frmCR_ConsultaFianzas", 0, , , False, Me, True)
  
  Case "Morosidad"
        GLOBALES.gCedulaActual = txtCedula.Text
        Call sbFormsCall("frmCR_ConsultaCreditosMora", 0, , , False, Me, True)
  
  Case "Deducciones"
      GLOBALES.gCedulaActual = txtCedula.Text
      Call sbFormsCall("frmCR_EnCobroCuotas", 0, , , False, Me, True)
  
  Case "Preliminar"
     Call sbFormsCall("frmCR_ConsultaPlanillaAbonoDist", 0, 0, 0, False, Me, True)
  
  Case "Transacciones"
     Call sbFormsCall("frmCR_ConsultaBitacora", 0, 0, 0, False, Me, True)

End Select

End Sub

Private Sub btnInfo_Click(Index As Integer)
Dim rsList As New ADODB.Recordset

On Error GoTo vError

Me.MousePointer = vbHourglass


lswDP.ListItems.Clear
lswDP.ColumnHeaders.Clear
lswDP.Checkboxes = False

btnInfoTriggerTag.Tag = btnInfo.Item(Index).Caption

Select Case Index
  Case 0 'Contacto
    
    lswDP.ColumnHeaders.Add 1, , "", 2500
    lswDP.ColumnHeaders.Add 2, , "", 4500
    
    strSQL = "select S.direccion,rtrim(isnull(Prov.Descripcion,'')) as 'ProvDesc', rtrim(isnull(Cant.Descripcion,'')) as 'CantonDesc', rtrim(isnull(Dist.Descripcion,'')) as 'DistDesc'" _
           & ",S.sexo,S.fecha_nac,S.estadoCivil, isnull(Ec.Descripcion,'') as 'EstadoCivil_Desc'" _
           & " ,S.AF_EMAIL as 'EMAIL_01', S.EMAIL_02, S.FACEBOOK, S.TWITTER, S.LINKEDIN   " _
           & " ,Ep.DESCRIPCION as 'EstadoPersona', S.FECHAINGRESO, isnull(Nc.DESCRIPCION,'') as 'Nacionalidad' " _
           & " , DATEDIFF(YEAR, isnull(S.FECHA_NAC, getdate()), GETDATE()) as 'Edad'" _
           & " from socios S left join Provincias Prov on S.Provincia = Prov.Provincia" _
           & " left join Cantones Cant on S.Provincia = Cant.Provincia and S.Canton = Cant.Canton" _
           & " left join Distritos Dist on S.Provincia = Dist.Provincia and S.Canton = Dist.Canton and S.distrito = Dist.distrito" _
           & " left join SYS_ESTADO_CIVIL Ec on S.EstadoCivil = Ec.Estado_Civil" _
           & " left join AFI_ESTADOS_PERSONA Ep on S.ESTADOACTUAL = Ep.COD_ESTADO" _
           & " left join SYS_NACIONALIDADES Nc on S.COD_NACIONALIDAD = Nc.COD_NACIONALIDAD" _
           & " where S.cedula = '" & txtCedula & "'"
    Call OpenRecordSet(rs, strSQL)
    If Not rs.EOF And Not rs.BOF Then
      
      Set itmX = lswDP.ListItems.Add(, , "Estado Actual")
          itmX.SubItems(1) = rs!EstadoPersona
      Set itmX = lswDP.ListItems.Add(, , "Fecha Ingreso")
          itmX.SubItems(1) = Format(rs!FechaIngreso, "dd/mm/yyyy")
      
      
      Set itmX = lswDP.ListItems.Add(, , "")
          itmX.SubItems(1) = ""
      Set itmX = lswDP.ListItems.Add(, , "Fecha Nacimiento")
          itmX.SubItems(1) = Format(rs!fecha_nac, "dd/mm/yyyy") & "  Edad: " & rs!Edad & " años."
    
      Set itmX = lswDP.ListItems.Add(, , "Género")
          itmX.SubItems(1) = IIf((rs!sexo = "M"), "Masculino", "Femenino")
    
      Set itmX = lswDP.ListItems.Add(, , "Estado Civil")
          itmX.SubItems(1) = rs!EstadoCivil_Desc
    
      
      Set itmX = lswDP.ListItems.Add(, , "Nacionalidad")
          itmX.SubItems(1) = rs!Nacionalidad
      
      
      Set itmX = lswDP.ListItems.Add(, , "")
          itmX.SubItems(1) = ""
      Set itmX = lswDP.ListItems.Add(, , "Dirección:")
          itmX.Bold = True
          
      Set itmX = lswDP.ListItems.Add(, , "Provincia")
          itmX.SubItems(1) = rs!ProvDesc
      Set itmX = lswDP.ListItems.Add(, , "Cantón")
          itmX.SubItems(1) = rs!CantonDesc
      Set itmX = lswDP.ListItems.Add(, , "Distrito")
          itmX.SubItems(1) = rs!DistDesc
      Set itmX = lswDP.ListItems.Add(, , "Otras Señas")
          itmX.SubItems(1) = rs!direccion & ""
       
      Set itmX = lswDP.ListItems.Add(, , "")
          itmX.SubItems(1) = ""
      Set itmX = lswDP.ListItems.Add(, , "Email No.1")
          itmX.SubItems(1) = rs!Email_01 & ""
      Set itmX = lswDP.ListItems.Add(, , "Email No.2")
          itmX.SubItems(1) = rs!Email_02 & ""
       
      Set itmX = lswDP.ListItems.Add(, , "")
          itmX.SubItems(1) = ""
      Set itmX = lswDP.ListItems.Add(, , "Redes:")
          itmX.Bold = True
      Set itmX = lswDP.ListItems.Add(, , "Facebook")
          itmX.SubItems(1) = rs!Facebook & ""
      Set itmX = lswDP.ListItems.Add(, , "Twitter")
          itmX.SubItems(1) = rs!TWITTER & ""
      Set itmX = lswDP.ListItems.Add(, , "Linkedin")
          itmX.SubItems(1) = rs!Linkedin & ""
       
    End If
    rs.Close
    
  
  Case 1 'Estado Laboral
    lswDP.ColumnHeaders.Add 1, , "", 2500
    lswDP.ColumnHeaders.Add 2, , "", 4500
    
    If Not GLOBALES.SysASEVersion Then
        strSQL = "select I.descripcion as Institucion,D.descripcion as Departamento, X.descripcion as Seccion" _
               & ", S.NOMBRAMIENTO_FECHA as Fecha,DATEDIFF(yyyy,S.NOMBRAMIENTO_FECHA,dbo.MyGetdate())as AniosLaborados" _
               & ", isnull(El.DESCRIPCION,'No Indica') as 'EstadoLaboral'" _
               & " from socios S left join instituciones I on S.cod_institucion = I.cod_institucion" _
               & " left join afDepartamentos D on S.cod_institucion = D.cod_institucion and S.cod_departamento = D.cod_departamento" _
               & " left join afSecciones X on S.cod_institucion = X.cod_institucion" _
               & " and S.cod_departamento = X.cod_departamento and S.cod_seccion = X.cod_Seccion" _
               & " left join AFI_ESTADO_LABORAL El on S.ESTADOLABORAL = El.ESTADO_LABORAL" _
               & " where S.cedula = '" & txtCedula.Text & "'"
        Call OpenRecordSet(rs, strSQL)
        If Not rs.EOF And Not rs.BOF Then
            
            Set itmX = lswDP.ListItems.Add(, , "Institución")
                itmX.SubItems(1) = rs!Institucion
            Set itmX = lswDP.ListItems.Add(, , "Departamento")
                itmX.SubItems(1) = rs!departamento
            Set itmX = lswDP.ListItems.Add(, , "Sección")
                itmX.SubItems(1) = rs!seccion
                
            Set itmX = lswDP.ListItems.Add(, , "")
            
            Set itmX = lswDP.ListItems.Add(, , "Estado Laboral")
                itmX.SubItems(1) = rs!estadoLaboral
           
            Set itmX = lswDP.ListItems.Add(, , "")
           
            Set itmX = lswDP.ListItems.Add(, , "Fecha Nombramiento")
                itmX.SubItems(1) = rs!fecha & ""
            Set itmX = lswDP.ListItems.Add(, , "Años Laborados")
                itmX.SubItems(1) = rs!AniosLaborados & ""
          
        End If
        rs.Close

    Else
        'version con unidades programaticas
        strSQL = "select S.UP,S.UT,I.descripcion as Institucion,D.descripcion as UProgramatica, X.UT_descripcion as UTrabajo" _
               & ", S.NOMBRAMIENTO_FECHA as Fecha,DATEDIFF(yyyy,S.NOMBRAMIENTO_FECHA,dbo.MyGetdate())as AniosLaborados" _
               & ", isnull(El.DESCRIPCION,'No Indica') as 'EstadoLaboral'" _
               & " from socios S left join instituciones I on S.cod_institucion = I.cod_institucion" _
               & " left join uprogramatica D on S.UP = D.codigo" _
               & " left join utrabajo X on S.UT = X.UT_codigo" _
               & " left join AFI_ESTADO_LABORAL El on S.ESTADOLABORAL = El.ESTADO_LABORAL" _
               & " where S.cedula = '" & txtCedula & "'"
        Call OpenRecordSet(rs, strSQL)
        If Not rs.EOF And Not rs.BOF Then
       
            Set itmX = lswDP.ListItems.Add(, , "Institución")
                itmX.SubItems(1) = rs!Institucion
            Set itmX = lswDP.ListItems.Add(, , "Ud. Programatica")
                itmX.SubItems(1) = rs!UProgramatica & ""
            Set itmX = lswDP.ListItems.Add(, , "Ud. Trabajo")
                itmX.SubItems(1) = rs!UTrabajo & ""
                
            Set itmX = lswDP.ListItems.Add(, , "")
            
            Set itmX = lswDP.ListItems.Add(, , "Estado Laboral")
                itmX.SubItems(1) = rs!estadoLaboral
           
            Set itmX = lswDP.ListItems.Add(, , "")
           
            Set itmX = lswDP.ListItems.Add(, , "Fecha Nombramiento")
                itmX.SubItems(1) = rs!fecha & ""
            
            Set itmX = lswDP.ListItems.Add(, , "Años Laborados")
                itmX.SubItems(1) = rs!AniosLaborados & ""
                
                
        End If
        rs.Close
    
    End If
    
  Case 2 'Telefonos
    lswDP.ColumnHeaders.Add 1, , "Numero", 1500
    lswDP.ColumnHeaders.Add 2, , "Tipo", 1500
    lswDP.ColumnHeaders.Add 3, , "Extension", 1500
    lswDP.ColumnHeaders.Add 4, , "Contacto", 2500
    
    
    strSQL = "Select Numero,Tipo,Ext,Contacto From Telefonos where " _
           & "Cedula='" & Trim(txtCedula) & "'"
    Call OpenRecordSet(rs, strSQL)
    Do While Not rs.EOF
       Set itmX = lswDP.ListItems.Add(, , Trim(rs!Numero))
           itmX.SubItems(1) = fxTipoTelefono(rs!Tipo)
           itmX.SubItems(2) = Trim(rs!Ext) & ""
           itmX.SubItems(3) = Trim(rs!contacto) & ""
       rs.MoveNext
    Loop
    rs.Close
  
  

  
  Case 3 'Cuentas Bancarias
        lswDP.ColumnHeaders.Clear
        lswDP.ColumnHeaders.Add 1, , "Cuenta", 2500
        lswDP.ColumnHeaders.Add 2, , "Banco", 3500
        lswDP.ColumnHeaders.Add 3, , "Tipo", 1100, vbCenter
        lswDP.ColumnHeaders.Add 4, , "Divisa", 1100, vbCenter
        lswDP.ColumnHeaders.Add 5, , "Interbanca", 2500
        lswDP.ColumnHeaders.Add 6, , "Activa", 1100, vbCenter
        lswDP.ColumnHeaders.Add 7, , "Fecha", 2500
        lswDP.ColumnHeaders.Add 8, , "Usuario", 2500
       
        strSQL = "select rtrim(B.Descripcion) as 'Banco'" _
               & ",case when C.tipo = 'A' then 'Ahorros' else 'Corriente' end as 'TipoDesc'" _
               & ",C.cod_Divisa,C.CUENTA_INTERNA, C.CUENTA_INTERBANCA, C.ACTIVA, C.REGISTRO_FECHA , C.REGISTRO_USUARIO" _
               & " from SYS_CUENTAS_BANCARIAS C inner join TES_BANCOS_GRUPOS B on C.cod_banco = B.cod_grupo" _
               & " where C.Identificacion = '" & Trim(txtCedula) & "'" 'and C.Modulo = 'AFI'
    
        Call OpenRecordSet(rs, strSQL)
        Do While Not rs.EOF
           Set itmX = lswDP.ListItems.Add(, , rs!CUENTA_INTERNA)
               itmX.SubItems(1) = Trim(rs!Banco)
               itmX.SubItems(2) = rs!TipoDesc
               itmX.SubItems(3) = rs!COD_DIVISA
               itmX.SubItems(4) = rs!CUENTA_INTERBANCA
               itmX.SubItems(5) = IIf(rs!Activa = 1, "Activa", "Cerrada")
               itmX.SubItems(6) = rs!Registro_Fecha & ""
               itmX.SubItems(7) = rs!Registro_Usuario & ""
         
           rs.MoveNext
        Loop
        rs.Close
    
    
    
  Case 4 'Ingresos
    lswDP.Visible = True
    lswDP.ColumnHeaders.Add 1, , "Fecha", 1200
    lswDP.ColumnHeaders.Add 2, , "Usuario", 1500
    lswDP.ColumnHeaders.Add 3, , "Boleta", 900
    lswDP.ColumnHeaders.Add 4, , "Promotor", 3500
    
    'Pregunta si se utilizan unidades programaticas o departamentos ?
    strSQL = "select I.*,P.nombre" _
           & " from afi_ingresos I left join Promotores P on I.id_promotor = P.id_promotor" _
           & " where I.cedula = '" & txtCedula & "'"
    Call OpenRecordSet(rs, strSQL)
    Do While Not rs.EOF
       Set itmX = lswDP.ListItems.Add(, , Format(rs!Fecha_Ingreso, "dd/mm/yyyy"))
           itmX.SubItems(1) = Trim(rs!Usuario & "")
           itmX.SubItems(2) = Trim(rs!Boleta & "")
           itmX.SubItems(3) = Trim(rs!Nombre & "")
       rs.MoveNext
    Loop
    rs.Close
    
  Case 5 'Liquidaciones
    lswDP.Visible = True
   
    With lswDP.ColumnHeaders
      .Clear
      .Add , , "Fecha", 1440
      .Add , , "Liq Id.", 1400
      .Add , , "Tipo", 1440, vbCenter
      .Add , , "Anterior", 1440, vbCenter
      .Add , , "Documento", 1440, vbCenter
      .Add , , "Ubicación", 1440, vbCenter
      .Add , , "Neto", 1640, vbRightJustify
      .Add , , "Estado", 1440, vbCenter
    
    End With
    
    
   With lswDP
     strSQL = "select L.*,E.descripcion as 'EstadoPersona'" _
            & " from liquidacion L inner join afi_estados_persona E on L.EstadoActual = E.cod_estado where L.cedula = '" & Trim(txtCedula) & "' order by L.fecliq"
     Call OpenRecordSet(rs, strSQL, 0)
     .ListItems.Clear
     Do While Not rs.EOF
        Set itmX = .ListItems.Add(, , Format(rs!fecLiq, "yyyy/mm/dd"))
            itmX.SubItems(1) = rs!consec
       
       If rs!estadoactliq = "A" Then
            itmX.SubItems(2) = "Ren.Asociación"
       Else
            itmX.SubItems(2) = "Ren.Patronal"
       End If
                
       Select Case rs!EstadoActual
         Case "S"
            itmX.SubItems(3) = "Asociado"
         Case "A"
            itmX.SubItems(3) = "Ren.Asociación"
         Case "P"
            itmX.SubItems(3) = "Ren.Patronal"
         Case "N"
       End Select
       
       itmX.SubItems(3) = rs!EstadoPersona
       itmX.SubItems(4) = rs!TDOCUMENTO & rs!nDocumento & ""
       
       If rs!ubicacion = "C" Then
           itmX.SubItems(5) = "Contabilidad"
       Else
           itmX.SubItems(5) = "Tesorería"
       End If
       
       itmX.SubItems(6) = Format(rs!TNETO, "Standard")
       
       If rs!Estado = "P" Then
          itmX.SubItems(7) = "Procesada"
       Else
          itmX.SubItems(7) = "Reversada"
       End If
       
       rs.MoveNext
     Loop
     rs.Close
   End With
    
    
    
   Case 6 'Beneficiarios
    lswDP.Visible = True
    lswDP.ColumnHeaders.Add 1, , "Identificación", 1500
    lswDP.ColumnHeaders.Add 2, , "Nombre", 3500
    lswDP.ColumnHeaders.Add 3, , "Porcentaje", 1100, vbRightJustify
    lswDP.ColumnHeaders.Add 4, , "Relación", 1200, vbCenter
    lswDP.ColumnHeaders.Add 5, , "Parentesco", 1100, vbCenter
    

    strSQL = "exec spAFI_PERSONA_BENEFICIARIOS_Consulta '" & Trim(txtCedula) & "',0"
    Call OpenRecordSet(rs, strSQL)
    Do While Not rs.EOF
       Set itmX = lswDP.ListItems.Add(, , rs!cedula_Beneficiario)
           itmX.SubItems(1) = Trim(rs!Nombre)
           itmX.SubItems(2) = Trim(rs!Porcentaje)
           itmX.SubItems(3) = Trim(rs!Relacion_Desc)
           itmX.SubItems(4) = Trim(rs!parentesco)
       
       rs.MoveNext
    Loop
    rs.Close
    
    

  Case 7 'Canal de Comunicacion
  
  
    lswDP.Visible = True
    lswDP.ListItems.Clear
    lswDP.ColumnHeaders.Clear
    lswDP.ColumnHeaders.Add 1, , "Tipo de Canal", 3500
    lswDP.ColumnHeaders.Add 2, , "Usuario", 2100
    lswDP.ColumnHeaders.Add 3, , "Fecha", 2500
    lswDP.Checkboxes = True
    
           
    strSQL = "exec spAFI_Persona_Canales_Consulta '" & txtCedula.Text & "',1"
    Call OpenRecordSet(rs, strSQL)
    
    vPaso = True
    With lswDP.ListItems
       .Clear
       Do While Not rs.EOF
        Set itmX = .Add(, , rs!Descripcion)
            itmX.SubItems(1) = rs!Registro_Usuario & ""
            itmX.SubItems(2) = rs!Registro_Fecha & ""
            itmX.Tag = rs!Canal_Tipo
            
            itmX.Checked = IIf((rs!Asignado = 1), True, False)
            
        rs.MoveNext
       Loop
       rs.Close
    End With
    vPaso = False

  
  Case 8 'Gustos y Preferencias
  
  
    lswDP.Visible = True
    lswDP.ListItems.Clear
    lswDP.ColumnHeaders.Clear
    lswDP.ColumnHeaders.Add 1, , "Gustos y Preferencias", 3500
    lswDP.ColumnHeaders.Add 2, , "Usuario", 2100
    lswDP.ColumnHeaders.Add 3, , "Fecha", 2500
    lswDP.Checkboxes = True

      strSQL = "exec spAFI_Persona_Preferencias_Consulta '" & txtCedula.Text & "',1"
      Call OpenRecordSet(rs, strSQL)
      
      vPaso = True
      With lswDP.ListItems
         .Clear
         Do While Not rs.EOF
          Set itmX = .Add(, , rs!Descripcion)
              itmX.SubItems(1) = rs!Registro_Usuario & ""
              itmX.SubItems(2) = rs!Registro_Fecha & ""
              itmX.Tag = rs!cod_preferencia
              
              itmX.Checked = IIf((rs!Asignado = 1), True, False)
              
          rs.MoveNext
         Loop
         rs.Close
      End With
      vPaso = False

    
    
  
  Case 9 'Bienes
  
  
    lswDP.Visible = True
    lswDP.ColumnHeaders.Clear
    lswDP.ColumnHeaders.Clear
    lswDP.ColumnHeaders.Add 1, , "Bieness", 3500
    lswDP.ColumnHeaders.Add 2, , "Usuario", 2100
    lswDP.ColumnHeaders.Add 3, , "Fecha", 2500
    lswDP.Checkboxes = True
    
  
    strSQL = "exec spAFI_Persona_Bienes_Consulta '" & txtCedula.Text & "',1"
    Call OpenRecordSet(rs, strSQL)
    
    vPaso = True
    With lswDP.ListItems
       .Clear
       Do While Not rs.EOF
        Set itmX = .Add(, , rs!Descripcion)
            itmX.SubItems(1) = rs!Registro_Usuario & ""
            itmX.SubItems(2) = rs!Registro_Fecha & ""
            itmX.Tag = rs!Bien_Tipo
            
            itmX.Checked = IIf((rs!Asignado = 1), True, False)
            
        rs.MoveNext
       Loop
       rs.Close
    End With
    vPaso = False
  
  Case 10 'Escolaridad
  
    lswDP.Visible = True
    lswDP.ListItems.Clear
    lswDP.ColumnHeaders.Clear
    lswDP.ColumnHeaders.Add 1, , "Nivel de Escolaridad", 3500
    lswDP.ColumnHeaders.Add 2, , "Usuario", 2100
    lswDP.ColumnHeaders.Add 3, , "Fecha", 2500
    lswDP.Checkboxes = True
    
 
    strSQL = "exec spAFI_Persona_Escolaridad_Consulta '" & txtCedula.Text & "',1"
    Call OpenRecordSet(rs, strSQL)
    
    vPaso = True
    With lswDP.ListItems
       .Clear
       Do While Not rs.EOF
        Set itmX = .Add(, , rs!Descripcion)
            itmX.SubItems(1) = rs!Registro_Usuario & ""
            itmX.SubItems(2) = rs!Registro_Fecha & ""
            itmX.Tag = rs!Escolaridad_Tipo
            
            itmX.Checked = IIf((rs!Asignado = 1), True, False)
            
        rs.MoveNext
       Loop
       rs.Close
    End With
    vPaso = False
  
    
  Case 11 'Polizas Colectivas
    lswDP.ColumnHeaders.Add 1, , "Tipo Id", 1500
    lswDP.ColumnHeaders.Add 2, , "Cedula", 1500, vbCenter
    lswDP.ColumnHeaders.Add 3, , "Nombre", 3500
    lswDP.ColumnHeaders.Add 4, , "Parentesco", 1500
    lswDP.ColumnHeaders.Add 5, , "Porcentaje", 1000, vbRightJustify
    
    strSQL = "select COD_POLIZA, POLIZA_DESC from vPoliza_Catalogo"
    Call OpenRecordSet(rsList, strSQL)
    Do While Not rsList.EOF
    
           Set itmX = lswDP.ListItems.Add(, , "Poliza:")
                itmX.SubItems(1) = rsList!cod_poliza
                itmX.SubItems(2) = rsList!Poliza_Desc
                itmX.SubItems(3) = ""
                itmX.SubItems(4) = ""
                
               itmX.Bold = True
               itmX.ForeColor = vbWhite
               itmX.TextBackColor = RGB(214, 234, 248)
                
           strSQL = "exec spPoliza_Persona_Beneficiarios '" & Trim(txtCedula) & "', '" & rsList!cod_poliza & "'"
           Call OpenRecordSet(rs, strSQL)
           Do While Not rs.EOF
                Set itmX = lswDP.ListItems.Add(, , rs!Tipo_Id_Desc)
                     itmX.SubItems(1) = rs!Cedula
                     itmX.SubItems(2) = rs!Nombre
                     itmX.SubItems(3) = rs!Parentesco_Desc
                     itmX.SubItems(4) = Format(rs!Porcentaje, "Standard")
             
             rs.MoveNext
           Loop
           rs.Close
           
           Set itmX = lswDP.ListItems.Add(, , "")
                itmX.SubItems(1) = ""
                itmX.SubItems(2) = ""
                itmX.SubItems(3) = ""
                itmX.SubItems(4) = ""
           
      rsList.MoveNext
    Loop
    rsList.Close
 
End Select

Me.MousePointer = vbDefault
Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub lswFND_Click()

On Error GoTo vError


gbFndContrato.Caption = "Plan : " & lswFND.SelectedItem.SubItems(1) & "  Contrato : " & lswFND.SelectedItem.SubItems(2)
gbFndContrato.Tag = lswFND.SelectedItem.SubItems(2)
gbFndContrato.ToolTipText = lswFND.SelectedItem.SubItems(1)

Me.MousePointer = vbDefault
Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub



Private Sub sbFondos(vCedula As String)
Dim curMonto As Currency, curMensualidad As Currency
Dim curAportes As Currency, curRendi As Currency


On Error GoTo vError

Me.MousePointer = vbHourglass

With lswFND
  .ListItems.Clear
  .ColumnHeaders.Clear
  .ColumnHeaders.Add , , "Op", 210
  .ColumnHeaders.Add , , "Plan", 1000
  .ColumnHeaders.Add , , "Contrato", 1100
  .ColumnHeaders.Add , , "Inicio", 1300, vbCenter
  .ColumnHeaders.Add , , "Mensualidad", 1300, vbRightJustify
  .ColumnHeaders.Add , , "Aportes", 1300, vbRightJustify
  .ColumnHeaders.Add , , "Rendimiento", 1300, vbRightJustify
  .ColumnHeaders.Add , , "Total", 1300, vbRightJustify
  .ColumnHeaders.Add , , "Plan", 3500
  .ColumnHeaders.Add , , "Estado", 1100
  .ColumnHeaders.Add , , "Operadora", 3000
  .ColumnHeaders.Add , , "IBAN", 3000
 
End With

pOperadora = 1
pPlan = ""
pContrato = 0

gbFndContrato.Caption = "Contrato No.?"


lswFnd_List.ListItems.Clear

strSQL = "exec spFndContratosConsulta '" & vCedula & "','" & glogon.Usuario & "'"
Call OpenRecordSet(rs, strSQL)

curAportes = 0
curRendi = 0
curMonto = 0
curMensualidad = 0

Do While Not rs.EOF
 Set itmX = lswFND.ListItems.Add(, , rs!COD_OPERADORA)
     itmX.SubItems(1) = rs!Cod_Plan
     itmX.SubItems(2) = rs!COD_Contrato
     itmX.SubItems(3) = Format(rs!Fecha_Inicio, "yyyy-mm-dd")
     itmX.SubItems(4) = Format(rs!Monto, "Standard")
     itmX.SubItems(5) = Format(rs!APORTES, "Standard")
     itmX.SubItems(6) = Format(rs!Rendimiento, "Standard")
     itmX.SubItems(7) = Format(rs!APORTES + rs!Rendimiento, "Standard")
     itmX.SubItems(8) = rs!Plan_Desc
     itmX.SubItems(9) = rs!Estado_Desc
     itmX.SubItems(10) = rs!OPERADORA_DESC
     itmX.SubItems(11) = rs!IBAN

 
 curAportes = curAportes + rs!APORTES
 curRendi = curRendi + rs!Rendimiento
 curMonto = curMonto + rs!APORTES + rs!Rendimiento
 
 rs.MoveNext
Loop
rs.Close


Set itmX = lswFND.ListItems.Add(, , ".")
    itmX.SubItems(4) = "__________________"
    itmX.SubItems(5) = "__________________"
    itmX.SubItems(6) = "__________________"
    itmX.SubItems(7) = "__________________"

Set itmX = lswFND.ListItems.Add(, , ".")
    itmX.SubItems(4) = Format(curMensualidad, "Standard")
    itmX.SubItems(5) = Format(curAportes, "Standard")
    itmX.SubItems(6) = Format(curRendi, "Standard")
    itmX.SubItems(7) = Format(curMonto, "Standard")
    itmX.TextBackColor = RGB(214, 234, 248)
    

Me.MousePointer = vbDefault
Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub sbCargaMsj(vCedula As String)

Dim vTipo As String

On Error GoTo vError

If Len(Trim(txtCedula.Text)) = 0 Then Exit Sub

Me.MousePointer = vbHourglass

'Inicializa Datos y Encabezados
'dtpMsjVence.Value = fxFechaServidor
vGrid.MaxRows = 0
vGrid.MaxCols = 5

vTipo = Mid(cboMsj.Text, 1, 1)

txtMsj = ""
fraMsj.Visible = False


imgMsjResuelve.top = imgBorraMsj.top
imgMsjResuelve.Left = imgBorraMsj.Left

imgMsjResuelve.Visible = False
imgBorraMsj.Visible = False


Select Case vTipo
   Case "P"
        imgMsjResuelve.Visible = True
   
   Case "G", "M", "B", "R"
        imgBorraMsj.Visible = True
   
   Case Else
        imgMsjResuelve.Visible = False
        imgBorraMsj.Visible = False
   
End Select

'2023-12-01 Elimina Opcion
imgMsjResuelve.Visible = False
imgBorraMsj.Visible = False


If vTipo = "R" Then
 strSQL = "select * from socios_mensajes where cedula = '" _
        & vCedula & "' and datediff(d,dbo.MyGetdate(),vencimiento) >= 0 and Tipo = 'P'" _
        & " and Resolucion = 'R' order by Fecha desc"

Else
 strSQL = "select * from socios_mensajes where cedula = '" _
        & vCedula & "' and datediff(d,dbo.MyGetdate(),vencimiento) >= 0 and Tipo = '" & vTipo _
        & "' and isnull(Resolucion, 'P') = 'P'  order by Fecha desc"
End If

Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
  vGrid.MaxRows = vGrid.MaxRows + 1
  vGrid.Row = vGrid.MaxRows
  
  vGrid.Col = 1
  vGrid.Text = Format(rs!Vencimiento, "dd/mm/yyyy")
  vGrid.TextTip = TextTipFixed
  vGrid.TextTipDelay = 1000

  vGrid.CellNote = "Fecha : " & rs!fecha & vbCrLf & "Usuario : " & rs!Usuario
  vGrid.CellTag = rs!Usuario
   
    
  vGrid.Col = 2
  vGrid.Text = rs!Mensaje
      
  vGrid.Col = 4
  vGrid.Text = rs!fecha
      
  vGrid.Col = 5
  vGrid.Text = rs!Usuario
      
'  vGrid.col = 6
'  vGrid.Text = rs!Resolucion_Fecha & ""
'
'  vGrid.col = 7
'  vGrid.Text = rs!Resolucion_Usuario & ""
  vGrid.RowHeight(vGrid.Row) = vGrid.MaxTextRowHeight(vGrid.Row)
  
 rs.MoveNext
Loop
rs.Close


Me.MousePointer = vbDefault

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub sbSoS_Resumen(vCedula As String)
Dim curMonto As Currency

On Error GoTo vError

Me.MousePointer = vbHourglass


strSQL = "exec spSOS_Exclusiones_Consulta '" & vCedula & "', '" & glogon.Usuario & "'"
Call OpenRecordSet(rs, strSQL)

vPaso = True

    If rs!Estado = "A" Then
      chkSoS_Exclusion.Value = xtpChecked
      chkSoS_Exclusion.ToolTipText = "Fecha: " & rs!Registro_Fecha & ".." & rs!Registro_Usuario & ""
    Else
      chkSoS_Exclusion.Value = xtpUnchecked
      chkSoS_Exclusion.ToolTipText = ""
    End If

vPaso = False

rs.Close

lswSoS_Det.ListItems.Clear

With lswSoS
  .ListItems.Clear
  .ColumnHeaders.Clear
  .ColumnHeaders.Add , , "Proceso", 2100
  .ColumnHeaders.Add , , "Devolución", 1800, vbRightJustify
  .ColumnHeaders.Add , , "Tipo", 1000, vbCenter
  .ColumnHeaders.Add , , "IBAN", 2500
  .ColumnHeaders.Add , , "Tesoreria Id", 2100, vbCenter
  .ColumnHeaders.Add , , "Cuenta", 2100, vbCenter
  .ColumnHeaders.Add , , "Estado", 1100, vbCenter
  .ColumnHeaders.Add , , "Emite?", 1800, vbCenter
  .ColumnHeaders.Add , , "No. TF", 1100, vbCenter
  .ColumnHeaders.Add , , "No. Documento", 1800, vbCenter
End With


strSQL = "exec spSOS_Consulta_Resumen '" & vCedula & "','" & glogon.Usuario & "'"
Call OpenRecordSet(rs, strSQL)

curMonto = 0

Do While Not rs.EOF
 Set itmX = lswSoS.ListItems.Add(, , rs!Proceso)
     itmX.SubItems(1) = Format(rs!devolucion, "Standard")
     itmX.SubItems(2) = rs!Tipo
     itmX.SubItems(3) = rs!IBAN & ""
     itmX.SubItems(4) = rs!Tesoreria_Id & ""
     itmX.SubItems(5) = rs!Bancos_Cuenta & ""
     itmX.SubItems(6) = rs!Bancos_Estado & ""
     itmX.SubItems(7) = rs!Bancos_Fecha & ""
     itmX.SubItems(8) = rs!Bancos_TF & ""
     itmX.SubItems(9) = rs!Bancos_Documento & ""

 curMonto = curMonto + rs!devolucion
 
 rs.MoveNext
Loop
rs.Close


'Set itmX = lswSoS.ListItems.Add(, , ".")
'    itmX.SubItems(1) = "__________________"
'
'Set itmX = lswSoS.ListItems.Add(, , ".")
'    itmX.SubItems(1) = Format(curMonto, "Standard")
'    itmX.TextBackColor = RGB(214, 234, 248)
    
txtSoS_Monto.Text = Format(curMonto, "Standard")

Me.MousePointer = vbDefault
Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub




Public Sub sbSoS_Operaciones(pCedula As String, pProceso As Currency)
    
On Error GoTo vError

Me.MousePointer = vbHourglass

 
lswSoS_Det.ListItems.Clear

With lswSoS_Det.ColumnHeaders
    .Clear
    .Add , , "Proceso", 2100, vbCenter
    .Add , , "Operación", 2100, vbCenter
    .Add , , "Código", 1000, vbCenter
    .Add , , "Descripción", 3000
    .Add , , "Garantía", 2000
    .Add , , "Monto Base", 2500, vbRightJustify
    .Add , , "Devolución", 2500, vbRightJustify
    .Add , , "Tipo Doc", 2500
    .Add , , "Num. Doc", 2500
End With

strSQL = "exec spSOS_Consulta_Operaciones '" & pCedula & "', " & pProceso & ", '" & glogon.Usuario & "'"

 Call OpenRecordSet(rs, strSQL)
 
 Do While Not rs.EOF
   Set itmX = lswSoS_Det.ListItems.Add(, , rs!Proceso)
       itmX.SubItems(1) = rs!Id_Solicitud
       itmX.SubItems(2) = rs!Codigo
       itmX.SubItems(3) = rs!Linea_Desc
       itmX.SubItems(4) = rs!Garantia_Desc
       itmX.SubItems(5) = Format(rs!Monto_Base, "Standard")
       itmX.SubItems(6) = Format(rs!devolucion, "Standard")
       itmX.SubItems(7) = rs!Tipo_Doc
       itmX.SubItems(8) = rs!Num_Com
   rs.MoveNext
 Loop
 rs.Close


    

Me.MousePointer = vbDefault

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub




Private Sub sbCorreo(vCedula As String)

On Error GoTo vError

If Len(Trim(txtCedula.Text)) = 0 Then Exit Sub

Me.MousePointer = vbHourglass

lswCorreo.ListItems.Clear

With lswCorreo.ColumnHeaders
    .Clear
    .Add , , "Fecha", 1800
    .Add , , "Estado", 1200, vbCenter
    .Add , , "Asunto", 2800
    .Add , , "Usuario", 1200, vbCenter
    .Add , , "Cuenta", 1000, vbCenter
    .Add , , "Para:", 3200
End With


strSQL = "exec spSys_Mail_Load '" & txtCedula.Text & "'"
Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
  Set itmX = lswCorreo.ListItems.Add(, , rs!fecha)
      itmX.SubItems(1) = rs!EstadoDesc
      itmX.SubItems(2) = rs!Asunto & ""
      itmX.SubItems(3) = rs!Usuario & ""
      itmX.SubItems(4) = rs!COD_SMTP
      itmX.SubItems(5) = rs!Para & ""
 rs.MoveNext
Loop
rs.Close

Me.MousePointer = vbDefault

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub




Private Sub isButtonCb_Click(Index As Integer)

Select Case Index
 Case 0 'Expediente
         Call vgCobro_ButtonClicked(1, 1, 1)
 Case 1 'Advertencias
        GLOBALES.gTag = txtCedula.Text
        Call sbFormsCall("frmCO_AdvertenciasRegistro", 1, , , False, Me)
        
 Case 2 'Notificación
        Call sbCbr_Notifica_Email(txtCedula.Text, IIf(rbNotificaEmail.Item(0).Value, "R", "D"))
 
End Select
  
End Sub

Private Sub lswBeneficios_ItemClick(ByVal Item As XtremeSuiteControls.ListViewItem)


On Error GoTo vError

Call sbClassCall("Beneficios", 5, "frmAF_BeneficioAsg", Item.Text, Item.SubItems(1))

Exit Sub

vError:
 'Nada
End Sub


Private Sub lswDP_ItemCheck(ByVal Item As XtremeSuiteControls.ListViewItem)

If vPaso Then Exit Sub

On Error GoTo vError

Select Case btnInfoTriggerTag.Tag
  Case "Canales"
    strSQL = "exec spAFI_Persona_Canales_Registra '" & txtCedula.Text & "','" & Item.Tag & "','" _
           & IIf((Item.Checked = True), "A", "E") & "','" & glogon.Usuario & "'"
    Call ConectionExecute(strSQL)
  Case "G y P"
    strSQL = "exec spAFI_Persona_Preferencias_Registra '" & txtCedula.Text & "','" & Item.Tag & "','" _
           & IIf((Item.Checked = True), "A", "E") & "','" & glogon.Usuario & "'"
    Call ConectionExecute(strSQL)
  
  Case "Bienes"
    strSQL = "exec spAFI_Persona_Bienes_Registra '" & txtCedula.Text & "','" & Item.Tag & "','" _
           & IIf((Item.Checked = True), "A", "E") & "','" & glogon.Usuario & "'"
    Call ConectionExecute(strSQL)
  
  Case "Escolaridad"
    strSQL = "exec spAFI_Persona_Escolaridad_Registra '" & txtCedula.Text & "','" & Item.Tag & "','" _
           & IIf((Item.Checked = True), "A", "E") & "','" & glogon.Usuario & "'"
    Call ConectionExecute(strSQL)
  
  
  
  Case Else
End Select

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub


Private Sub lswFND_ItemClick(ByVal Item As XtremeSuiteControls.ListViewItem)
Dim curMonto As Currency


On Error GoTo vError

gbFndContrato.Tag = "0"
gbFndContrato.ToolTipText = ""
gbFndContrato.Caption = "Contrato No."



If lswFND.SelectedItem = "." Then Exit Sub

Me.MousePointer = vbHourglass

pOperadora = Item.Text
pPlan = Item.SubItems(1)
pContrato = Item.SubItems(2)



gbFndContrato.Caption = "Plan : " & pPlan & "  Contrato : " & pContrato
gbFndContrato.Tag = pContrato
gbFndContrato.ToolTipText = pPlan


Call btnFondos_List_Click(0)


Me.MousePointer = vbDefault
Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub lswRenuncias_ItemClick(ByVal Item As XtremeSuiteControls.ListViewItem)
Dim frm As Form

On Error GoTo vError

Call sbFormsCall("frmAF_CRRenuncia", , , , False, Me, True)

Call sbFormActivo("frmAF_CRRenuncia", frm)

Call frm.sbConsulta_Externa(Item.Text)

Exit Sub

vError:
 'Nada

End Sub


Private Sub lswSoS_ItemClick(ByVal Item As XtremeSuiteControls.ListViewItem)

Call sbSoS_Operaciones(txtCedula.Text, Item.Text)

End Sub


Private Sub tcMain_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
 Call Form_Resize
End Sub

Private Sub TimerX_Timer()


TimerX.Interval = 0
TimerX.Enabled = False

On Error GoTo vError


strSQL = "select Portal_ID from sif_Empresa"
Call OpenRecordSet(rs, strSQL)
If rs!Portal_Id = 53 Then 'ASOSEJUD
    btnSoS.Visible = True
End If
rs.Close

cboMsj.Text = "Generales"
txtCedula.SetFocus

Exit Sub

vError:

End Sub

Private Sub sbReporteConsentimiento()

Me.MousePointer = vbHourglass

On Error GoTo vError

With frmContenedor.Crt
     .Reset
     .WindowShowRefreshBtn = True
     .WindowShowPrintSetupBtn = True
     .WindowState = crptMaximized
     .WindowShowSearchBtn = True
     .WindowTitle = "Reportes Módulo de Personas"
     
     .Connect = glogon.ConectRPT
     
     
     .ReportFileName = SIFGlobal.fxPathReportes("Personas_ConsentimientoInfo.rpt")
     .SelectionFormula = "{SOCIOS.CEDULA} = '" & txtCedula & "'"

    
      .PrintReport

End With

Me.MousePointer = vbDefault
Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub tlbConsultaCreditos_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim frm As Form
Dim curMonto As Currency

Select Case Button.Key
  Case "fianzas"
     GLOBALES.gCedulaActual = txtCedula.Text
     Call sbFormsCall("frmCR_ConsultaFianzas", 1, , , False, Me)
     
  Case "operacion"
     Call sbFormsCall("frmCR_ConsultaOperaciones")
  
  Case "calculo"
  
        GLOBALES.gCedulaActual = txtCedula.Text
        Call sbFormsCall("frmCR_CalculoOperacion", 0)
        
  Case "estado"
        GLOBALES.gCedulaActual = txtCedula.Text
        Call sbFormsCall("frmCR_ConsultaCreditosMora", 1, , , False, Me)
   
    Case "EstadoCta"
      Call sbEstadoCuenta(txtCedula)
    
    Case "EnCobro"
      GLOBALES.gCedulaActual = txtCedula.Text
      Call sbFormsCall("frmCR_EnCobroCuotas", 1, , , False, Me)
 
 
 End Select
 
End Sub


Private Function fxLiquidacion(vCedula As String) As String

On Error GoTo vError

Dim pResultado As String

pResultado = ""

glogon.strSQL = "select C.descripcion" _
       & " from liquidacion L inner join Causas_Renuncias C on C.id_causa = L.id_causa" _
       & " where consec in(select max(consec) from liquidacion" _
       & " where cedula = '" & vCedula & "')"
Call OpenRecordSet(glogon.Recordset, glogon.strSQL)
If Not rs.EOF And Not rs.BOF Then
 pResultado = "[CAUSA: " & Trim(glogon.Recordset!Descripcion) & "]"
End If

glogon.Recordset.Close

fxLiquidacion = pResultado

Exit Function

vError:
    fxLiquidacion = pResultado

End Function

Private Sub sbConsulta(pCedula As String)
Dim vFechaIng As Date, vFianzas As Boolean
Dim rsTmp As New ADODB.Recordset, i As Integer
     
     
fraConsentimiento.Visible = False
tcMain.Visible = True
     
txtAhorro.Text = 0
txtAporte.Text = 0
txtCustodia.Text = 0
txtCapitalizacion.Text = 0
txtPatrimonio.Text = 0

txtPat_Divisa.Text = ""

lblFechaAhorro.Caption = ""
lblFechaAporte.Caption = ""
lblFechaCustodia.Caption = ""
lblCapitalizado.Caption = ""
  
lblEstadoCobros.Caption = "Gestiones de Cobro?"
lblEstadoAdvertencias.Caption = "Advertencias?"
  
btnSaldosFavor.Caption = "0.00"
  
vFianzas = False
  
If Not fxSIFValidaCadena(txtCedula.Text) Then
   Exit Sub
End If
  
'Valida Acceso a Expediente
vRA_Access = fxSys_RA_Consulta(Trim(pCedula), glogon.Usuario)
 
If Not vRA_Access Then
    MsgBox "Esta persona se encuentra con -> Expediente Restringido <- Requiere de Autorización para Consultar!", vbExclamation
    txtCedula.Text = ""
    txtNombre.Text = ""
    Exit Sub
End If
  
strSQL = "exec spSys_Consulta_Integrada '" & Trim(pCedula) & "'"
Call OpenRecordSet(rs, strSQL)
 
If Not rs.EOF And Not rs.BOF Then
   
   txtRate.Text = Format(IIf(IsNull(rs!Rating), 0, rs!Rating), "Standard")
   
   txtCedula.Text = Trim(rs!Cedulax & "")
   txtNombre.Text = rs!Nombre & ""
   
   btnSaldosFavor.Caption = Format(rs!Cajas_Saldo_Favor, "Standard")
   
   txtAhorro.Text = Format(IIf(IsNull(rs!ahorro), 0, rs!ahorro), "Standard")
   txtAporte.Text = Format(IIf(IsNull(rs!Aporte), 0, rs!Aporte), "Standard")
   txtCustodia.Text = Format(IIf(IsNull(rs!Custodia), 0, rs!Custodia), "Standard")
   txtCapitalizacion.Text = Format(IIf(IsNull(rs!capitaliza), 0, rs!capitaliza), "Standard")
   
   txtPatrimonio.Text = Format(CCur(txtAhorro.Text) + CCur(txtCustodia.Text) + CCur(txtAporte.Text) + CCur(txtCapitalizacion.Text), "Standard")
   
   lblFechaAhorro.Caption = IIf(IsNull(rs!fecAhorro), "", Format(rs!fecAhorro, "dd/mm/yyyy"))
   lblFechaAporte.Caption = IIf(IsNull(rs!fecaporte), "", Format(rs!fecaporte, "dd/mm/yyyy"))
   lblFechaCustodia.Caption = IIf(IsNull(rs!fecCustodia), "", Format(rs!fecCustodia, "dd/mm/yyyy"))
   lblCapitalizado.Caption = IIf(IsNull(rs!fecCapitaliza), "", Format(rs!fecCapitaliza, "dd/mm/yyyy"))
   
   txtPat_Divisa.Text = rs!COD_DIVISA
   txtPAT_Disponible.Text = Format(rs!Pat_Garantia_Total, "Standard")
   txtPAT_Saldos.Text = Format(rs!Pat_Garantia_Saldos, "Standard")
   txtPAT_Saldos.Tag = rs!Pat_Garantia_Saldos
   
   txtPAT_AporteCobro.Text = Format(rs!Pat_Aporte_Manual, "Standard")
   
   cboPAT_TipoSaldo.Text = "Saldos en Garantía"

   txtNotas = rs!Notas & ""
   
   txtNotas.ToolTipText = "Usuario : " & rs!nota_user & " Fecha : " & rs!nota_fecha & ""
   
   If rs!SALARIO_TRASLADA = 1 Then
        lblSalarioTraslada.Caption = "Traslada Salario: Sí"
   Else
        lblSalarioTraslada.Caption = "Traslada Salario ?"
   End If
   
   lblTarjeta.Caption = "Tarjeta: " & rs!tarjeta_Numero & ""
   lblIBAN.Caption = "IBAN: " & rs!IBAN & ""
   
   If rs!bloqueo = 0 Then
     lblBloqueo.BackColor = vbGreen
   Else
     lblBloqueo.BackColor = vbRed
   End If

   lblEstado.Caption = "Estado : " & rs!EstadoX
   
   lblInstitución.Caption = rs!InstitucionX
   lblInstitución.ToolTipText = "Deductora: " & rs!Deductora
    
     
     vFechaIng = IIf(IsNull(rs!FechaIngreso), fxFechaServidor, rs!FechaIngreso)
     
     lblMembresia.ForeColor = vbWhite
     lblMembresia.FontBold = False
     lblMembresia.BackStyle = 0
     
     
     If rs!EstadoActual = "S" Then
        lblMembresia.Caption = "Membresía: " & fxMembresia(vFechaIng)
        lblMembresia.ToolTipText = "[Ing.:" & Format(vFechaIng, "dd/mm/yyyy") & "]"

               
        strSQL = "exec spAFI_ConsultaRenunciaTransito '" & pCedula & "'"
        Call OpenRecordSet(rsTmp, strSQL, 0)
        If Not rsTmp.EOF And Not rsTmp.BOF Then
            lblMembresia.Caption = "Renuncia: " & rsTmp!Cod_Renuncia & " ¦ " & rsTmp!Registro_Fecha & " ¦ " & rsTmp!registro_user
            lblMembresia.ToolTipText = rsTmp!Estado & " ¦ " & rsTmp!Tipo & " ¦ " & Trim(rsTmp!Descripcion)
            
            lblMembresia.BackStyle = 1
            lblMembresia.BackColor = RGB(199, 138, 156)
            
            lblMembresia.FontBold = True
        End If
        rsTmp.Close
     
     Else
        lblMembresia.Caption = "Membresía: NADA"
        lblMembresia.ToolTipText = fxLiquidacion(rs!Cedulax)
     End If
     
     'Clasificación de la Persona
     lblClasificacion.Caption = "Clasificación Crediticia : [" & rs!Clasificacion & "]"
    
    
    'Indica el Estado de las Fianzas
     If rs!IndFianzas = 0 Then
       vFianzas = False
       lblFianzas.Caption = "Fianzas al Día"
       Set imgFianzas.Picture = imgSemaforos.ListImages.Item(1).Picture
     Else
       vFianzas = True
       lblFianzas.Caption = "Fianzas en Mora"
       Set imgFianzas.Picture = imgSemaforos.ListImages.Item(3).Picture
     End If
     
     
'     'Indica los Mensajes
     If rs!IndMensajes = 0 Then
       lblEstadoMensajes.Caption = "Mensajes ?"
     Else
       lblEstadoMensajes.Caption = "Mensajes (" & rs!IndMensajes & ")"
     End If


'    'Indicar de Gestiones de Cobros
     If rs!IndCobro = 0 Then
       lblEstadoCobros.Caption = "Sin Gestión de Cobro"
     Else
       lblEstadoCobros.Caption = "Gestiones de Cobro(" & rs!IndCobro & ")"
     End If
     
'    'Indicar de Advertencias
     If rs!IndAdvertencias = 0 Then
       lblEstadoAdvertencias.Caption = "Sin Advertencias"
     Else
       lblEstadoAdvertencias.Caption = "Advertencias (" & rs!IndAdvertencias & ")"
     End If
     
     
     

     'Indicador de Estado de Beneficiarios
     Select Case rs!IndBeneficiarios
       Case 0 'Rojo
        Set imgEstadoBeneficiarios.Picture = imgSemaforos.ListImages.Item(3).Picture
       Case 1 'Verde
        Set imgEstadoBeneficiarios.Picture = imgSemaforos.ListImages.Item(1).Picture
       Case 2 'Amarillo
        Set imgEstadoBeneficiarios.Picture = imgSemaforos.ListImages.Item(2).Picture
     End Select
     
     lblEstadoBeneficiarios.ToolTipText = "Fecha   .: " & Format(rs!Ben_Update_Fecha & "", "dd/mm/yyyy") & vbCrLf _
                                          & "Usuario .: " & rs!Ben_Update_Usuario & ""



     'Pregunta por el Consentimiento de Uso de la Información Personal para Contacto
    If IsNull(rs!Consentimiento_Contacto_Fecha) Then
       Set imgEstadoConsentimiento.Picture = imgSemaforos.ListImages.Item(3).Picture
       
       imgEstadoConsentimiento.ToolTipText = ""
       
       txtConsentimientoFecha.Text = ""
       txtConsentimientoUsuario.Text = ""
    Else
       Set imgEstadoConsentimiento.Picture = imgSemaforos.ListImages.Item(1).Picture
       txtConsentimientoFecha.Text = rs!Consentimiento_Contacto_Fecha & ""
       txtConsentimientoUsuario.Text = rs!Consentimiento_Contacto_Usuario & ""
       
       imgEstadoConsentimiento.ToolTipText = "Fecha : " & rs!Consentimiento_Contacto_Fecha & "...Usuario: " & rs!Consentimiento_Contacto_Usuario & ""
       
    End If
     
     
     
     If Len(rs!Pat_Advertencia) > 0 Then
            MsgBox rs!Pat_Advertencia, vbExclamation, "Advertencia de Aportes no cotizados"
     End If
     
     'Cierra RecordSet abierto
'     rs.Close
     
     strSQL = "exec spSIFPersonaMensajes '" & pCedula & "'"
     Call OpenRecordSet(rs, strSQL)
     If rs!Pendientes > 0 Then
         lblMsjPendientes.Caption = "Pendientes (" & rs!Pendientes & ")"
         Set imgMsjPendientes.Picture = imgSemaforos.ListImages.Item(9).Picture
     Else
         lblMsjPendientes.Caption = "Msj. Pendientes?"
         Set imgMsjPendientes.Picture = imgSemaforos.ListImages.Item(6).Picture
     End If

     If rs!Advertencias > 0 Then
         lblMsjAdvertencias.Caption = "Advertencias (" & rs!Advertencias & ")"
         Set imgMsjAdvertencia.Picture = imgSemaforos.ListImages.Item(9).Picture
     Else
         lblMsjAdvertencias.Caption = "Msj Advertencias?"
         Set imgMsjAdvertencia.Picture = imgSemaforos.ListImages.Item(6).Picture
     End If

     If rs!Generales > 0 Then
         lblMsjGenerales.Caption = "General (" & rs!Generales & ")"
         Set imgMsjGenerales.Picture = imgSemaforos.ListImages.Item(9).Picture
     Else
         lblMsjGenerales.Caption = "Msj Generales?"
         Set imgMsjGenerales.Picture = imgSemaforos.ListImages.Item(6).Picture
     End If

     If rs!Morosidad > 0 Then
         lblMsjMorosidad.Caption = "Morosidad (" & rs!Morosidad & ")"
         Set imgMsjMorosidad.Picture = imgSemaforos.ListImages.Item(9).Picture
     Else
         lblMsjMorosidad.Caption = "Msj Morosidad?"
         Set imgMsjMorosidad.Picture = imgSemaforos.ListImages.Item(6).Picture
     End If

     If rs!Bloqueos > 0 Then
'         tlbMsj.Buttons.Item(9).Caption = "Bloqueos (" & rs!Bloqueos & ")"
'         Set imgMsjPendientes.Picture = imgSemaforos.ListImages.Item(9).Picture
     Else
'         Set imgMsjPendientes.Picture = imgSemaforos.ListImages.Item(6).Picture
     End If
     
     'Actualiza Disponible Garantia Sobre Ahorros
     Call cboPAT_TipoSaldo_Click

     'Actualiza el Detalle de Creditos
     Call sbCreditos
     
     'Consulta Traslado de Salario
     Call sbTraslado_Salario(txtCedula.Text)
 Else
   MsgBox "No Se encontró registro de la persona solicitada", vbInformation
   Exit Sub
 End If
   
 
End Sub

Public Sub sbXConsultaAsistida(vCedula As String)
  txtCedula = vCedula
  Call txtCedula_KeyDown(vbKeyReturn, 0)
End Sub




Private Sub sbTraslado_Salario(pCedula As String)

Dim strEstadoTramite As String
Dim strRs() As String


On Error GoTo vError

strEstadoTramite = "1>0"
lblSalarioTraslada.Caption = "Sin tramite (Traslado Salario)"
lblSalarioTraslada.ToolTipText = "El asociado no cuenta con un tramite de traslado salario"
Set imgTS.Picture = imgSemaforos.ListImages.Item(16).Picture  'Boli gris

'Solo para ASECCSS
If gPortal.Empresa_Id <> 61 Then
    Exit Sub
End If

'Nota es una Funcion con prefijo de Procedure
strSQL = "select dbo.sp_CONSULTA_ESTADO_TRAMITETS ('" & pCedula & "') as 'Resultado'"
Call OpenRecordSet(rs, strSQL)
If Not rs.BOF And Not rs.BOF Then
    strEstadoTramite = rs!Resultado
End If


strRs = Split(strEstadoTramite, ">")

Select Case CInt(strRs(0))
    Case 1 'Sin Registro
        lblSalarioTraslada.Caption = "Sin tramite (Traslado Salario)"
        lblSalarioTraslada.ToolTipText = "El asociado no cuenta con un tramite de traslado salario con Aseccss"
        Set imgTS.Picture = imgSemaforos.ListImages.Item(16).Picture 'Boli gris
    Case 2 'No tramitado
        lblSalarioTraslada.Caption = "No tramitado (Traslado Salario)"
        lblSalarioTraslada.ToolTipText = "Gestión de tramite de traslado de salario con Aseccss no tramitada"
        Set imgTS.Picture = imgSemaforos.ListImages.Item(16).Picture 'Boli gris
    Case 3 'En trÃ¡mite
        lblSalarioTraslada.Caption = "En trámite (Traslado Salario)"
        lblSalarioTraslada.ToolTipText = "Gestión de tramite de traslado de salario con Aseccss en proceso"
        Set imgTS.Picture = imgSemaforos.ListImages.Item(2).Picture 'Boli amarilla
    Case 4 'Activo
        lblSalarioTraslada.Caption = "Activo (Traslado Salario)"
        lblSalarioTraslada.ToolTipText = "Gestión de tramite de traslado de salario con Aseccss activo"
        Set imgTS.Picture = imgSemaforos.ListImages.Item(1).Picture 'Boli Verde
    Case 5 'Trasladado
        lblSalarioTraslada.Caption = "Trasladado (Traslado Salario)"
        lblSalarioTraslada.ToolTipText = "Asociado traslado su salario a otra entidad financiera"
        Set imgTS.Picture = imgSemaforos.ListImages.Item(16).Picture 'Boli gris
    Case 6 'Permiso sin goce
        lblSalarioTraslada.Caption = "Permiso sin goce (Traslado Salario)"
        lblSalarioTraslada.ToolTipText = "No se tiene recepción de pago de salario por permiso sin goce salarial"
        Set imgTS.Picture = imgSemaforos.ListImages.Item(6).Picture 'Boli Azul
    Case 7 'Sin movimiento
        If CInt(strRs(1) >= 6) Then
            lblSalarioTraslada.Caption = "Sin movimiento (6 meses o mas) (Traslado Salario)"
            lblSalarioTraslada.ToolTipText = "El tramite de traslado de salario con Aseccss tiene 6 meses o mas sin movimiento registrado"
            Set imgTS.Picture = imgSemaforos.ListImages.Item(3).Picture 'Boli Roja
        Else
            lblSalarioTraslada.Caption = "Sin movimiento (entre 1 a 5 meses) (Traslado Salario)"
            lblSalarioTraslada.ToolTipText = "El tramite de traslado de salario con Aseccss tiene de 1 a 5 meses inclusive sin movimiento registrado"
            Set imgTS.Picture = imgSemaforos.ListImages.Item(9).Picture 'Boli Naranja
        End If

End Select


Exit Sub

vError:
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical
    
End Sub

Private Sub txtCedula_KeyDown(KeyCode As Integer, Shift As Integer)
Dim vCedTemp As String

On Error GoTo vError

'Busca primer en el Maestro de Socios, de lo contrario revisa si es una operacion
' y regresa la cedula de la operacion

vCedTemp = ""

If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then
 strSQL = "select isnull(count(*),0) as Existe from socios where cedula = '" & txtCedula & "'"
 Call OpenRecordSet(rs, strSQL)
 If rs!Existe = 0 Then
   rs.Close
   strSQL = "select cedula from reg_creditos where id_solicitud = " & txtCedula
   Call OpenRecordSet(rs, strSQL)
   If Not rs.EOF And Not rs.BOF Then
      vCedTemp = Trim(rs!Cedula)
   End If
 End If
 rs.Close
 
    If vCedTemp = "" Then
        Call sbConsulta(txtCedula.Text)
    Else
        Call sbConsulta(vCedTemp)
    End If
End If

If KeyCode = vbKeyF4 Then Call sbBusqueda

Exit Sub

vError:

End Sub



Private Sub txtNombre_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then Call sbBusqueda

End Sub

Private Sub txtPAT_Saldos_Change()
On Error GoTo vError

txtPAT_Giro.Text = Format(CCur(txtPAT_Disponible.Text) - CCur(txtPAT_Saldos.Text), "Standard")

vError:

End Sub

Private Sub vgCobro_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
Dim frmX As Form

If vPaso Then Exit Sub

Call sbFormsCall("frmCO_ControlSeguimiento", 0, , , False, Me)

For Each frmX In Forms
   If Trim(frmX.Name) = "frmCO_ControlSeguimiento" Then
        Exit For
   End If
Next

Call frmX.sbCargaDatos(txtCedula.Text)
        
End Sub

Private Sub vgCobro_DblClick(ByVal Col As Long, ByVal Row As Long)
Dim frmX As Form

Call sbFormsCall("frmCO_ControlSeguimiento", 0, , , False, Me)

For Each frmX In Forms
   If Trim(frmX.Name) = "frmCO_ControlSeguimiento" Then
        Exit For
   End If
Next

Call frmX.sbCargaDatos(txtCedula.Text)
End Sub

Private Sub vgCobro_SheetChanged(ByVal OldSheet As Integer, ByVal NewSheet As Integer)
Dim i As Integer

On Error GoTo vError

Me.MousePointer = vbHourglass
vPaso = True

With vgCobro
    Select Case NewSheet
      Case 1 'Gestiones
        .ActiveSheet = NewSheet
        .Sheet = NewSheet
        .MaxRows = 0
       strSQL = "select S.*, isnull(G.descripcion,'') as 'Gestion'" _
              & "   , isnull(C.DESCRIPCION,'') as 'Causa'" _
              & "   , isnull(A.descripcion,'') as 'Arreglo'" _
              & " from CBR_Seguimiento S  left join cbr_gestiones G on S.cod_gestion = G.cod_gestion" _
              & "  left join CBR_CAUSAS_MOROSIDAD C on S.COD_CAUSA = C.COD_CAUSA" _
              & "  left join CBR_TIPOS_ARREGLOS A on S.COD_ARREGLO = A.COD_ARREGLO" _
              & " where cedula = '" & txtCedula.Text & "' order by S.cod_seg desc"
        Call OpenRecordSet(rs, strSQL)
        Do While Not rs.EOF
          .MaxRows = .MaxRows + 1
          .Row = .MaxRows
          
          For i = 2 To 11
            .Col = i
            Select Case i
              Case 2 'Fecha
                .Text = Format(rs!fecha, "dd/mm/yyyy")
              Case 3 'vencimiento
                .Text = Format(DateAdd("d", rs!tiempo_resolucion, rs!fecha), "dd/mm/yyyy")
              Case 4 'Gestión
                .Text = rs!Gestion
              Case 5 ' Detalle
                .Text = rs!Notas
                .RowHeight(.Row) = .MaxTextRowHeight(.Row)

              Case 6 ' Ejecutivo
                .Text = rs!Usuario
              Case 7 ' Monto
                .Text = Format(rs!Monto, "Standard")
              Case 8 ' Dias
                .Text = CStr(rs!tiempo_resolucion)
              Case 9  'Arrelgo de Pago
                .Text = rs!Arreglo
              Case 10 'Promesa de Pago
                .Text = Format(rs!Arreglo_Vence & "", "dd/mm/yyyy")
              Case 11 'Causa de Morosidad
                .Text = rs!Causa
                
            End Select
          Next i
          rs.MoveNext
        Loop
        rs.Close
      
      Case 2 'Oficiales
      
        .ActiveSheet = NewSheet
        .Sheet = NewSheet
        .MaxRows = 0
        strSQL = "select * from cbr_asignacion_h where cedula = '" & txtCedula.Text _
               & "' order by fecha_asignacion desc"
        Call OpenRecordSet(rs, strSQL)
        Do While Not rs.EOF
          .MaxRows = .MaxRows + 1
          .Row = .MaxRows
          
          For i = 1 To 5
            .Col = i
            Select Case i
              Case 1 'Fecha
                .Text = Format(rs!fecha_asignacion, "dd/mm/yyyy")
              Case 2 'Oficial
                .Text = UCase(rs!Usuario)
              Case 3 'Mantiene
                .Value = rs!mantener
              Case 4 ' Rebajo 2x
                .Value = rs!rebajo_doble
              Case 5 ' Mora
                .Value = rs!aplica_mora
            End Select
          Next i
          rs.MoveNext
        Loop
        rs.Close
      
    End Select
End With


Me.MousePointer = vbDefault
vPaso = False
Exit Sub

vError:
 vPaso = False
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub sbCancelaPlanPago(pOperacion As Long, Optional pBotonClick As Integer = 0)

Dim rs As New ADODB.Recordset

Me.MousePointer = vbHourglass

On Error GoTo vError

strSQL = "exec spCrdPlanPagosInfoCancelacion " & pOperacion & ",'" & Format(dtpCorte.Value, "yyyy/mm/dd") & "'"


Call OpenRecordSet(rs, strSQL)


If StatusBar.Panels(5).Tag = "" Then StatusBar.Panels(5).Tag = 0

If pBotonClick = vbChecked Then
    StatusBar.Panels(5).Text = Format(CCur(StatusBar.Panels(5).Text) + rs!IntCor + rs!IntMor, "Standard")
    StatusBar.Panels(5).ToolTipText = "Cuota...:" & Format(CCur(StatusBar.Panels(5).Tag) + rs!Cuota, "Standard")
    
    StatusBar.Panels(6).Text = Format(CCur(StatusBar.Panels(6).Text) + rs!Principal + rs!IntCor + rs!IntMor + rs!Cargos + rs!Poliza + rs!CargoAnticipo, "Standard")
    StatusBar.Panels(7).Text = Format(CCur(StatusBar.Panels(7).Text) + (rs!Cargos + rs!Poliza + rs!CargoAnticipo), "Standard")
Else
    StatusBar.Panels(5).Text = Format(CCur(StatusBar.Panels(5).Text) - (rs!IntCor + rs!IntMor), "Standard")
    StatusBar.Panels(5).ToolTipText = "Cuota...:" & Format(CCur(StatusBar.Panels(5).Tag) - rs!Cuota, "Standard")
    
    StatusBar.Panels(6).Text = Format(CCur(StatusBar.Panels(6).Text) - (rs!Principal + rs!IntCor + rs!IntMor + rs!Cargos + rs!Poliza + rs!CargoAnticipo), "Standard")
    StatusBar.Panels(7).Text = Format(CCur(StatusBar.Panels(7).Text) - (rs!Cargos + rs!Poliza + rs!CargoAnticipo), "Standard")
End If
    

Me.MousePointer = vbDefault

Exit Sub
vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub vgCreditos_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
Dim pOperacion As Long

If vPaso Then Exit Sub

If GLOBALES.SysPlanPagos = 1 Then
    With vgCreditos
        .Sheet = .ActiveSheet
        .Row = Row
        .Col = 3
        pOperacion = .Text
        .Col = 2
        Call sbCancelaPlanPago(pOperacion, .Value)
    End With
    Exit Sub
End If


'Sin Plan de Pagos
Me.MousePointer = vbHourglass

Dim rsTmp As New ADODB.Recordset, rs As New ADODB.Recordset
Dim vFecha As Date, vProceso As Long, curInteres As Currency
Dim curCargoAnticipo As Currency, curPorcentaje As Currency
Dim vUltimaCuota As Currency, iMeses As Integer


On Error GoTo vError


With vgCreditos

    .Sheet = .ActiveSheet
    .Row = Row
    
    .Col = 3
    
    curInteres = 0
    curCargoAnticipo = 0
    
    strSQL = "select R.saldo,R.interesv,R.fecUlt,R.fechaforp,isnull(V.intc+V.intm,0) as IntMora,isnull(V.cargos,0) as 'Cargos', C.PORC_CARGO_CANCELACION" _
           & ",isnull(V.cuota,0) as MoraCuota,isnull(V.Amortiza,0) as 'PrincipalAtrasado', R.id_solicitud, dbo.MyGetdate() as FechaActual,R.PriDeduc" _
           & " from reg_creditos R inner join catalogo C on R.codigo = C.codigo" _
           & " left join vista_morosidad V on R.id_solicitud = V.id_solicitud" _
           & " where R.id_solicitud = " & .Text
           
    Call OpenRecordSet(rs, strSQL)
       
       'Fecha Corte
       vFecha = dtpCorte.Value
       vProceso = Year(vFecha) & Format(Month(vFecha), "00")
       curPorcentaje = rs!PORC_CARGO_CANCELACION / 100
       vUltimaCuota = rs!FecUlt
       
       
       'Si existe morosidad, Preguntar si la ultima cuota en mora en igual o mayor al proceso
       ' en caso de ser afirmativo entonces no registrar los dias transcurridos.
       ' si no hay mora el proceso es igual al proceso de mora
       
       If rs!MoraCuota > 0 Then
         strSQL = "select max(fechap) as Proceso from morosidad where estado = 'A' and id_solicitud = " & rs!Id_Solicitud
         Call OpenRecordSet(rsTmp, strSQL, 0)
            If rsTmp!Proceso > vUltimaCuota Then vUltimaCuota = rsTmp!Proceso
         rsTmp.Close
       End If
       
       
       Select Case True
       
         Case vProceso < rs!PriDeduc And vUltimaCuota < rs!PriDeduc
              curInteres = 0
              
         Case vProceso = rs!PriDeduc And vUltimaCuota = rs!PriDeduc
              curInteres = 0
              
         Case vProceso > rs!PriDeduc And vUltimaCuota > vProceso
              curInteres = 0
         
         
         Case vProceso = rs!PriDeduc And vUltimaCuota < vProceso 'Dias
               curInteres = (rs!Saldo * rs!interesv / 36000) * (Day(vFecha))
         
         
         Case (vProceso > rs!PriDeduc And vUltimaCuota = rs!PriDeduc)
                
                iMeses = -1
                Do While vProceso > vUltimaCuota
                   iMeses = iMeses + 1
                   vUltimaCuota = fxFechaProcesoSiguiente(vUltimaCuota)
                Loop
                curInteres = (rs!Saldo * rs!interesv / 36000) * (Day(vFecha) + (iMeses * 30))
       
       
         Case (vProceso > rs!PriDeduc And vProceso > vUltimaCuota)  'Idem Anterior
                
                iMeses = -1
                Do While vProceso > vUltimaCuota
                   iMeses = iMeses + 1
                   vUltimaCuota = fxFechaProcesoSiguiente(vUltimaCuota)
                Loop
                curInteres = (rs!Saldo * rs!interesv / 36000) * (Day(vFecha) + (iMeses * 30))
       
       
         Case Else
              
              curInteres = 0
       
       End Select
       
      If rs!PrincipalAtrasado >= rs!Saldo Then
        curInteres = 0
      End If
      
       curInteres = curInteres + rs!intMora + rs!Cargos
    
    If .Value = vbChecked Then
        StatusBar.Panels(5).Text = Format(CCur(StatusBar.Panels(5).Text) + curInteres, "Standard")
        StatusBar.Panels(6).Text = Format(CCur(StatusBar.Panels(6).Text) + rs!Saldo + curInteres + curCargoAnticipo, "Standard")
        StatusBar.Panels(7).Text = Format(CCur(StatusBar.Panels(7).Text) + curCargoAnticipo, "Standard")
    Else
        StatusBar.Panels(5).Text = Format(CCur(StatusBar.Panels(5).Text) - curInteres, "Standard")
        StatusBar.Panels(6).Text = Format(CCur(StatusBar.Panels(6).Text) - rs!Saldo - curInteres - curCargoAnticipo, "Standard")
        StatusBar.Panels(7).Text = Format(CCur(StatusBar.Panels(7).Text) - curCargoAnticipo, "Standard")
    End If
    
    rs.Close

End With

Me.MousePointer = vbDefault

Exit Sub
vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub vgCreditos_DblClick(ByVal Col As Long, ByVal Row As Long)
Dim frm As Form
Dim x As clsEstudioCrd


On Error GoTo vError

With vgCreditos
    .Sheet = .ActiveSheet
    .Row = .ActiveRow
    

    Select Case .ActiveSheet
      Case 1 'Creditos Activos
         If .MaxRows = 0 Then Exit Sub
      
         .Col = 3
         Operacion.OperacionConsulta = .CellTag
'         frmCR_ConsultaDetalle.Show vbModal
         Call sbFormsCall("frmCR_ConsultaDetalle", , , , , Me, True)
      
      Case 2 'Creditos Cancelados
         If .MaxRows = 0 Then Exit Sub
         .Col = 3
         Operacion.OperacionConsulta = .Text
'         frmCR_ConsultaDetalle.Show vbModal
         Call sbFormsCall("frmCR_ConsultaDetalle", , , , , Me, True)
      
     
      Case 4 'Estudio de Credito
         .Col = 7
            Set x = New clsEstudioCrd
            Set x.vCon = glogon.Conection
            x.xOperacion = .Text
            x.xkey = glogon.ConectRPT
            
            
         .Col = 2
            If .MaxRows = 0 Then
                x.vSolicitudPreanalisis = 0
            Else
                x.vSolicitudPreanalisis = .Text
            End If
            Call x.Menu(glogon.Usuario, glogon.Conection, glogon.BaseDatos, glogon.Servidor _
                        , App.Path, glogon.ConectRPT, 11, glogon.AppName, glogon.AppVersion, glogon.Maquina _
                        , glogon.Portal_Con, glogon.Portal_User, glogon.Portal_Key, gPortal.Empresa_Id, gPortal.Empresa_Name)

            Set x = Nothing
      
    End Select
    

End With

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbExclamation




End Sub


Public Sub sbMenuPopUp_Acciones(Index As Integer)
Dim x As clsEstudioCrd, vOperacion As Long
Dim vExpediente As String, vCajas As Boolean
Dim frmConsultaActiva As Form, frm As Form

On Error GoTo vError

vOperacion = 0
vExpediente = ""
vCajas = False


With vgCreditos
    .Sheet = .ActiveSheet
    .Row = .ActiveRow
    
    Select Case .Sheet
       Case 1 'Activos
          .Col = 2
          If Not IsNumeric(.CellTag) Then Exit Sub
          
          vOperacion = .CellTag
       Case 2, 3 'Cancelados y En Tramite
          .Col = 2
          If Not IsNumeric(.Text) Then Exit Sub
          vOperacion = .Text
       
       Case 4 'PreAnalisis
          .Col = 2
          vExpediente = .Text
          .Col = 7
          If IsNumeric(.Text) Then vOperacion = .Text
       
       Case 5 'Incobrables
          .Col = 2
          If Not IsNumeric(.Text) Then Exit Sub
          vOperacion = .Text
      
    End Select
  
  

 

    Select Case Index
      Case 1 'Abonos
            If vOperacion = 0 Then Exit Sub
            .Col = 6 'Saldo
            If CCur(.Text) = 0 Then Exit Sub
            
            vCajas = IIf((fxCajasParametros("01") = "S"), True, False)
                
                .Col = 18 'Cuotas Morosas
                If CInt(.Text) = 0 Then
                  If vCajas Then
                        ModuloCajas.mRef_01 = vOperacion
                         
                        If GLOBALES.SysPlanPagos = 1 Then
                                 Call sbFormsCall("frmCajas_Crd_AbonosCtP", vbModal, 0, 0, False, Me)
                        Else
                                 Call sbFormsCall("frmCajas_Crd_AbonosStP", vbModal, 0, 0, False, Me)
                        End If
                  
                  Else
                        If GLOBALES.SysPlanPagos = 1 Then
                                 Call sbFormsCall("frmCR_AbonosNew")
                        Else
                                 Call sbFormsCall("frmCR_Abonos")
                        End If
                        
                        For Each frm In Forms
                          If (UCase(frm.Name) = UCase("frmCR_Abonos")) Or (UCase(frm.Name) = UCase("frmCR_AbonosNew")) Then
                            Call frm.sbConsultaExterna(vOperacion)
                            Exit For
                          End If
                        Next frm
                  End If
                Else 'Abonos en Mora
                   
                  If vCajas Then
                        ModuloCajas.mRef_01 = vOperacion
                         
                        If GLOBALES.SysPlanPagos = 1 Then
                                 Call sbFormsCall("frmCajas_Crd_AbonosCtP", vbModal, 0, 0, False, Me)
                        Else
                                 Call sbFormsCall("frmCajas_Crd_AbonosStP", vbModal, 0, 0, False, Me)
                        End If
                  
                  Else
                       If GLOBALES.SysPlanPagos = 1 Then
                                Call sbFormsCall("frmCR_AbonosNew")
                       Else
                                Call sbFormsCall("frmCR_CancelaMorosidad")
                       End If
                    
                        For Each frm In Forms
                          If (UCase(frm.Name) = UCase("frmCR_CancelaMorosidad")) Or (UCase(frm.Name) = UCase("frmCR_AbonosNew")) Then
                            Call frm.sbConsultaExterna(vOperacion)
                            Exit For
                          End If
                        Next frm
                   End If
                
                End If
    
      
      Case 2 'Anulacion de Abonos
            If vOperacion = 0 Then Exit Sub
                
                If GLOBALES.SysPlanPagos = 1 Then
                            Call sbFormsCall("frmCR_AnulaAbonosNew")
                Else
                            Call sbFormsCall("frmCR_AnulaAbonos")
                End If
      
                            For Each frm In Forms
                              If (UCase(frm.Name) = UCase("frmCR_AnulaAbonos")) Or (UCase(frm.Name) = UCase("frmCR_AnulaAbonosNew")) Then
                                Call frm.sbConsultaExterna(vOperacion)
                                Exit For
                              End If
                            Next frm
      
      Case 3 'Gestion de Cobro
            If vOperacion = 0 Then Exit Sub
                
                Call sbFormsCall("frmCO_Principal")
                For Each frm In Forms
                  If UCase(frm.Name) = UCase("frmCO_Principal") Then
                    Call frm.sbConsultaExterna(vOperacion)
                    Exit For
                  End If
                Next frm
      
      
      Case 4 'Movimientos de la Operacion
            
            If vOperacion = 0 Then Exit Sub
            
            Me.MousePointer = vbHourglass
            
            With frmContenedor.Crt
                .Reset
                .WindowShowPrintSetupBtn = True
                .WindowShowRefreshBtn = True
                .WindowShowSearchBtn = True
                .WindowState = crptMaximized
                .WindowTitle = "Reportes del Módulo de Crédito"
                
                .Connect = glogon.ConectRPT
                
                If GLOBALES.SysPlanPagos = 0 Then
                    .ReportFileName = SIFGlobal.fxPathReportes("Cobro_AbonosOperacionFull.rpt")
                    .Formulas(0) = "Empresa='" & GLOBALES.gstrNombreEmpresa & "'"
                    .Formulas(1) = "SubTitulo='ABONOS ORDINARIOS/EXTRAORDINARIOS/MORATORIOS'"
                    .Formulas(2) = "Fecha='" & Format(fxFechaServidor, "dd/mm/yyyy") & "'"
                    .Formulas(3) = "Titulo='MOVIMIENTOS DE LA OPERACION'"
                    .SelectionFormula = "{REG_CREDITOS.ID_SOLICITUD} = " & vOperacion
                    
                    .SubreportToChange = "sbCorte"
                    .StoredProcParam(0) = vOperacion
                    .StoredProcParam(1) = Format(dtpCorte.Value, "yyyy/mm/dd")
                    
                    .SubreportToChange = "sbMovimientos"
                    
                    .StoredProcParam(0) = vOperacion
                    .StoredProcParam(1) = 1
                    
                Else
                     .ReportFileName = SIFGlobal.fxPathReportes("Credito_PlanPagosMov.rpt")
                    
                     .Formulas(0) = "fxFecha='FECHA: " & Format(fxFechaServidor, "dd/mm/yyyy  hh:mm:ss") & "'"
                     .Formulas(1) = "fxEmpresa='" & GLOBALES.gstrNombreEmpresa & "'"
                     .Formulas(2) = "fxUsuario='USER: " & glogon.Usuario & "'"
                     .Formulas(3) = "fxOficina='" & GLOBALES.gOficina & "'"
                     
                     .SelectionFormula = "{REG_CREDITOS.ID_SOLICITUD} = " & vOperacion
                     
                     .SubreportToChange = "sbCorte"
                     .StoredProcParam(0) = vOperacion
                     .StoredProcParam(1) = Format(dtpCorte.Value, "yyyy/mm/dd")
                
                End If

                .PrintReport

               
            End With
            Me.MousePointer = vbDefault
            
      
      
            
            
            
        Case 5 'Nuevo Credito
                GLOBALES.gCedulaActual = txtCedula.Text
                Call sbFormsCall("frmCR_SeguimientoTramites")
                For Each frm In Forms
                  If UCase(frm.Name) = UCase("frmCR_SeguimientoTramites") Then
                    Call frm.sbGXSegTraIniTlb
                    Exit For
                  End If
                Next frm
      
      
      Case 6 'Seguimiento de Tramites
            
            If vOperacion = 0 Then Exit Sub
                
                Call sbFormsCall("frmCR_SeguimientoTramites")
                For Each frm In Forms
                  If UCase(frm.Name) = UCase("frmCR_SeguimientoTramites") Then
                    Call frm.sbConsultaExterna(vOperacion)
                    Exit For
                  End If
                Next frm
      
      Case 7 'Nuevo Estudio Crd
      
            Set x = New clsEstudioCrd
            Set x.vCon = glogon.Conection
            x.xOperacion = vOperacion
            x.xkey = glogon.ConectRPT
      
            x.vSolicitudPreanalisis = 0
            x.vCedula = txtCedula.Text
    
            Call x.Menu(glogon.Usuario, glogon.Conection, glogon.BaseDatos, glogon.Servidor _
            , App.Path, glogon.ConectRPT, 12, glogon.AppName, glogon.AppVersion, glogon.Maquina _
            , glogon.Portal_Con, glogon.Portal_User, glogon.Portal_Key, gPortal.Empresa_Id, gPortal.Empresa_Name)
    
            Set x = Nothing
      
      Case 8 'Estudio Crd
            Set x = New clsEstudioCrd
            Set x.vCon = glogon.Conection
                x.xkey = glogon.ConectRPT
                     
            If .ActiveSheet = 4 Then
                    x.xOperacion = vOperacion
                    x.vSolicitudPreanalisis = vExpediente
                    Call x.Menu(glogon.Usuario, glogon.Conection, glogon.BaseDatos, glogon.Servidor _
                                , App.Path, glogon.ConectRPT, 11, glogon.AppName, glogon.AppVersion, glogon.Maquina _
                                , glogon.Portal_Con, glogon.Portal_User, glogon.Portal_Key, gPortal.Empresa_Id, gPortal.Empresa_Name)
                    
            Else
                    x.xOperacion = vOperacion
                    strSQL = "select cod_preAnalisis from CRD_PREA_PREANALISIS" _
                           & " Where id_solicitud = " & vOperacion
                    Call OpenRecordSet(rs, strSQL)
                    If rs.EOF And rs.BOF Then
                        Call x.Menu(glogon.Usuario, glogon.Conection, glogon.BaseDatos, glogon.Servidor _
                        , App.Path, glogon.ConectRPT, 2, glogon.AppName, glogon.AppVersion, glogon.Maquina _
                        , glogon.Portal_Con, glogon.Portal_User, glogon.Portal_Key, gPortal.Empresa_Id, gPortal.Empresa_Name)
                    
                    Else
                        x.vSolicitudPreanalisis = rs!cod_PreAnalisis
                        Call x.Menu(glogon.Usuario, glogon.Conection, glogon.BaseDatos, glogon.Servidor _
                        , App.Path, glogon.ConectRPT, 11, glogon.AppName, glogon.AppVersion, glogon.Maquina _
                        , glogon.Portal_Con, glogon.Portal_User, glogon.Portal_Key, gPortal.Empresa_Id, gPortal.Empresa_Name)
                    End If
                    rs.Close
                      
            End If
            
      
    
            Set x = Nothing
      
      Case 9 'Plan de Pagos
        If vOperacion = 0 Then Exit Sub
        
        Operacion.OperacionConsulta = vOperacion
        Call sbFormsCall("frmCR_PlanPagos", 1, , , False, Me)
    
      
    End Select

End With

Exit Sub
        
vError:
        Me.MousePointer = vbDefault
        MsgBox fxSys_Error_Handler(Err.Description), vbCritical


End Sub

'Private Sub cbMenuPopUp_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
'    Call sbMenuPopUp_Acciones(Control.Id)
'End Sub

'Private Sub sbMenuPopUp_Show(pItems As Integer)
'Dim oBar As XtremeCommandBars.CommandBar
'Dim oControl As XtremeCommandBars.CommandBarControl
'Dim oPopup As XtremeCommandBars.CommandBarPopup
'Dim i As Integer
'
'Set oBar = cbMenuPopUp.Add("Menú", xtpBarPopup)
'    oBar.EnableAnimation = True
'
'With oBar.Controls
''    'standard items
''    .Add xtpControlButton, 1, "Item1"
''    .Add xtpControlButton, 2, "Item2"
''    .Add(xtpControlButton, 3, "Item3").BeginGroup = True
''    'checkbox
''    Set oControl = .Add(xtpControlCheckBox, 4, "Checkbox")
''    oControl.Checked = True
''    oControl.BeginGroup = True
''    'sub-menu
''    Set oPopup = .Add(xtpControlPopup, 5, "Sub Menu")
''    With oPopup.CommandBar.Controls
''        .Add xtpControlButton, 1, "SubItem1"
''        .Add xtpControlButton, 2, "SubItem2"
''    End With
'
'Select Case pItems
'    Case 1 'Activos
'            .Add xtpControlButton, 1, "Abonos"
'            .Add xtpControlButton, 2, "Anulación de Abonos"
'            .Add(xtpControlButton, 3, "Gestión de Cobros").BeginGroup = True
'            .Add xtpControlButton, 4, "Estado de la Operación"
'            .Add(xtpControlButton, 5, "Nuevo Crédito").BeginGroup = True
'            .Add xtpControlButton, 6, "Tramite de Crédito"
'            .Add(xtpControlButton, 7, "Nuevo Estudio de Crédito").BeginGroup = True
'            .Add xtpControlButton, 8, "Estudio de Crédito"
'            .Add(xtpControlButton, 9, "Plan de Pagos").BeginGroup = True
'            .Add(xtpControlButton, 10, "Cerrar").BeginGroup = True
'
'    Case 2 'Cancelados
'            .Add xtpControlButton, 2, "Anulación de Abonos"
'            .Add xtpControlButton, 4, "Estado de la Operación"
'            .Add(xtpControlButton, 5, "Nuevo Crédito").BeginGroup = True
'            .Add xtpControlButton, 6, "Tramite de Crédito"
'            .Add(xtpControlButton, 7, "Nuevo Estudio de Crédito").BeginGroup = True
'            .Add xtpControlButton, 8, "Estudio de Crédito"
'            .Add(xtpControlButton, 9, "Plan de Pagos").BeginGroup = True
'            .Add(xtpControlButton, 10, "Cerrar").BeginGroup = True
'
'    Case 3  'Seguimiento de Tramites
'            .Add xtpControlButton, 6, "Tramite de Crédito"
'            .Add(xtpControlButton, 7, "Nuevo Estudio de Crédito").BeginGroup = True
'            .Add xtpControlButton, 8, "Estudio de Crédito"
'            .Add(xtpControlButton, 10, "Cerrar").BeginGroup = True
'
'     Case 4 'PreaAnalisis
'            .Add(xtpControlButton, 7, "Nuevo Estudio de Crédito").BeginGroup = True
'            .Add xtpControlButton, 8, "Estudio de Crédito"
'            .Add(xtpControlButton, 10, "Cerrar").BeginGroup = True
'     Case 5 'Incobrables
'            .Add xtpControlButton, 2, "Anulación de Abonos"
'            .Add(xtpControlButton, 10, "Cerrar").BeginGroup = True
'End Select
'
'End With
''show it
'oBar.ShowPopup
'End Sub


Private Sub vgCreditos_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim i As Integer

On Error GoTo vError

Dim frmX As Form

For Each frmX In Forms
   If Mid(frmX.Name, 1, 3) = "MDI" Then
        Exit For
   End If
Next



If Button = 2 Then
   Call PopupMenu(frmX.mnuAcciones, , x, y)
      
'    Call sbMenuPopUp_Show(vgCreditos.ActiveSheet)

End If

vError:

End Sub

Private Sub vgCreditos_SheetChanged(ByVal OldSheet As Integer, ByVal NewSheet As Integer)
Dim i As Integer


With MDIPrincipal.mnuAccionesSub
  For i = 0 To .Count - 1
      .Item(i).Visible = True
  Next i
End With



Select Case NewSheet
    Case 1 'Activos
       Call sbCreditos(NewSheet)
     
     Case 2 'Cancelados
       Call sbCreditos(NewSheet)
    
    
    Case 3  'Seguimiento de Tramites
       With MDIPrincipal.mnuAccionesSub
         For i = 0 To 5
             .Item(i).Visible = False
         Next i
       End With
       
       
       Call sbSolicitudes(txtCedula)
     
     Case 4 'PreaAnalisis
       With MDIPrincipal.mnuAccionesSub
         For i = 0 To 10
             .Item(i).Visible = False
         Next i
       End With
       
       Call sbPreAnalisis(txtCedula)
    
     Case 5 'Incobrables
       With MDIPrincipal.mnuAccionesSub
         For i = 0 To 2
             .Item(i).Visible = False
         Next i
       End With
       
       Call sbIncobrable(txtCedula)


End Select
End Sub

Private Sub sbExcedentes()
Dim vTipo As String

On Error GoTo vError

Me.MousePointer = vbHourglass



strSQL = "exec spEXC_Periodos_Visibles '" & txtCedula.Text & "'"

Call OpenRecordSet(rs, strSQL)

With vgPatrimonio
   .ActiveSheet = 6
   .Sheet = 6
   .MaxRows = 0
   
    Do While Not rs.EOF
       .MaxRows = .MaxRows + 1
       .Row = .MaxRows
        .Col = 1
        .CellTag = CStr(rs!Id_Periodo)
        
       .Col = 2
       .Text = Format((rs!Inicio), "YYYY-MM")
       .Col = 3
       .Text = Format((rs!Corte), "YYYY-MM")
       .Col = 4
       .Text = Format(rs!excedente_bruto, "Standard")
       
       .Col = 5
       .Text = Format(rs!Reserva, "Standard")
       
       .Col = 6
       .Text = Format(rs!capitalizado, "Standard")
       .Col = 7
       .Text = Format(rs!Renta, "Standard")
       .Col = 8
       .Text = Format(rs!excedente_neto, "Standard")
       .Col = 9
        If Not IsNull(rs!mora_aplicada) Then
            .Text = Format(rs!donacion + rs!mora_aplicada + rs!moraopcf_aplicada + rs!saldos_ase_aplicados + rs!capitalizado_individual, "Standard")
        End If
       .Col = 10
       .Text = Format(rs!excedente_final, "Standard")
         
     rs.MoveNext
    Loop
    rs.Close

End With

Me.MousePointer = vbDefault
Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical


End Sub





Private Sub vgPatrimonio_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
On Error GoTo vError

With vgPatrimonio
   .ActiveSheet = 6
   .Sheet = 6
   .Row = Row
   .Col = 1
   
   Call sbEstadoExcedentes(txtCedula.Text, .CellTag)
   
End With

vError:

End Sub

Private Sub vgPatrimonio_SheetChanged(ByVal OldSheet As Integer, ByVal NewSheet As Integer)
Dim vTipo As String

If vPaso Then Exit Sub

If NewSheet = 6 Then
  Call sbExcedentes
  Exit Sub
End If

On Error GoTo vError

Me.MousePointer = vbHourglass

vPaso = True

Select Case NewSheet
  Case 1 'Todos
    vTipo = "('O','P','C','E', 'X')"
  Case 2 'Obrero
    vTipo = "('O')"
  Case 3 'Patronal
    vTipo = "('P')"
  Case 4 'Custodia
    vTipo = "('X')"
  Case 5 'Capitalizacion
    vTipo = "('C')"
End Select

strSQL = "select Top 30 Ah.*, isnull(Doc.Descripcion,'') as 'DocDesc',isnull(Con.Descripcion,'') as 'ConDesc'" _
       & " from ahorro_detallado Ah left join SIF_Documentos Doc On Ah.Tcon = Doc.Tipo_Documento" _
       & " left join SIF_Conceptos Con on Ah.cod_Concepto = Con.cod_Concepto" _
       & " where Ah.cedula = '" & txtCedula & "' and Ah.tipo in" & vTipo & " order by Ah.fecha desc"

Call OpenRecordSet(rs, strSQL)

With vgPatrimonio
   
   .ActiveSheet = NewSheet
   
   
   .Sheet = NewSheet
   .MaxRows = 0
   
    Do While Not rs.EOF
       .MaxRows = .MaxRows + 1
       .Row = .MaxRows
       .Col = 1
       .Text = Format(rs!fecha, "dd/mm/yyyy")
       .Col = 2
       .Text = Format(rs!FechaProc, "####-##")
       .Col = 3
        Select Case rs!Tipo
          Case "O"
               .Text = "Obrero"
          Case "P"
            .Text = "Patronal"
          Case "X"
            .Text = "AP.Custodia"
          Case "C"
            .Text = "Capitalización"
          Case "E"
            .Text = "Extraordinario"
        End Select
        .Col = 4
        .Text = Format(rs!Monto, "Standard")
        .Col = 5
        .Text = rs!DocDesc
        .Col = 6
        .Text = rs!nCon & ""
         
        .Col = 7
        .Text = rs!ConDesc & ""
        .Col = 8
        .Text = rs!Usuario & ""
         

     rs.MoveNext
    Loop
    rs.Close

End With

vPaso = False

Me.MousePointer = vbDefault
Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

