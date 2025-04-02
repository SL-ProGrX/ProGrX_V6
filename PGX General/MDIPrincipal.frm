VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.controls.v22.1.0.ocx"
Object = "{B8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "Codejock.TaskPanel.v22.1.0.ocx"
Begin VB.MDIForm MDIPrincipal 
   BackColor       =   &H80000003&
   ClientHeight    =   8940
   ClientLeft      =   165
   ClientTop       =   810
   ClientWidth     =   15615
   HelpContextID   =   9010
   Icon            =   "MDIPrincipal.frx":0000
   LinkTopic       =   "MDIForm1"
   Picture         =   "MDIPrincipal.frx":071A
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin XtremeTaskPanel.TaskPanel tpContabilidad 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   2
      Top             =   495
      Visible         =   0   'False
      Width           =   15615
      _Version        =   1441793
      _ExtentX        =   27543
      _ExtentY        =   741
      _StockProps     =   64
      VisualTheme     =   13
      ItemLayout      =   2
      HotTrackStyle   =   1
      Begin XtremeSuiteControls.PushButton btnContabilidad 
         Height          =   360
         Index           =   3
         Left            =   11760
         TabIndex        =   8
         Top             =   45
         Width           =   3855
         _Version        =   1441793
         _ExtentX        =   6800
         _ExtentY        =   635
         _StockProps     =   79
         Caption         =   "Contabilidad"
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
         FlatStyle       =   -1  'True
         UseVisualStyle  =   -1  'True
         TextAlignment   =   1
         Appearance      =   17
      End
      Begin XtremeSuiteControls.PushButton btnContabilidad 
         Height          =   360
         Index           =   2
         Left            =   8160
         TabIndex        =   7
         Top             =   45
         Width           =   1455
         _Version        =   1441793
         _ExtentX        =   2566
         _ExtentY        =   635
         _StockProps     =   79
         Caption         =   "Revisión"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FlatStyle       =   -1  'True
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         Picture         =   "MDIPrincipal.frx":89F9
         PushButtonStyle =   2
      End
      Begin XtremeSuiteControls.PushButton btnContabilidad 
         Height          =   360
         Index           =   1
         Left            =   6720
         TabIndex        =   6
         Top             =   45
         Width           =   1455
         _Version        =   1441793
         _ExtentX        =   2566
         _ExtentY        =   635
         _StockProps     =   79
         Caption         =   "Cierres"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FlatStyle       =   -1  'True
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         Picture         =   "MDIPrincipal.frx":8DEF
         ImageAlignment  =   0
         PushButtonStyle =   2
      End
      Begin XtremeSuiteControls.FlatEdit txtMes 
         Height          =   315
         Left            =   840
         TabIndex        =   5
         Top             =   60
         Width           =   495
         _Version        =   1441793
         _ExtentX        =   873
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
         Alignment       =   2
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtAnio 
         Height          =   315
         Left            =   120
         TabIndex        =   4
         Top             =   60
         Width           =   705
         _Version        =   1441793
         _ExtentX        =   1244
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
         Alignment       =   2
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.PushButton btnContabilidad 
         Height          =   360
         Index           =   0
         Left            =   1320
         TabIndex        =   3
         Top             =   45
         Width           =   3615
         _Version        =   1441793
         _ExtentX        =   6376
         _ExtentY        =   635
         _StockProps     =   79
         Caption         =   "Periodos"
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
         FlatStyle       =   -1  'True
         UseVisualStyle  =   -1  'True
         TextAlignment   =   0
         Appearance      =   17
      End
   End
   Begin XtremeTaskPanel.TaskPanel tpMain 
      Align           =   1  'Align Top
      Height          =   495
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   15615
      _Version        =   1441793
      _ExtentX        =   27543
      _ExtentY        =   873
      _StockProps     =   64
      VisualTheme     =   13
      ItemLayout      =   2
      HotTrackStyle   =   1
      Begin XtremeSuiteControls.PushButton btnFavoritos 
         Height          =   360
         Index           =   10
         Left            =   9480
         TabIndex        =   26
         Top             =   75
         Visible         =   0   'False
         Width           =   495
         _Version        =   1441793
         _ExtentX        =   873
         _ExtentY        =   635
         _StockProps     =   79
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FlatStyle       =   -1  'True
         UseVisualStyle  =   -1  'True
         Appearance      =   17
      End
      Begin XtremeSuiteControls.PushButton btnFavoritos 
         Height          =   360
         Index           =   9
         Left            =   9000
         TabIndex        =   25
         Top             =   75
         Visible         =   0   'False
         Width           =   495
         _Version        =   1441793
         _ExtentX        =   873
         _ExtentY        =   635
         _StockProps     =   79
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FlatStyle       =   -1  'True
         UseVisualStyle  =   -1  'True
         Appearance      =   17
      End
      Begin XtremeSuiteControls.PushButton btnEmpresa 
         Height          =   360
         Left            =   13560
         TabIndex        =   24
         Top             =   80
         Width           =   2055
         _Version        =   1441793
         _ExtentX        =   3625
         _ExtentY        =   635
         _StockProps     =   79
         Caption         =   "..."
         BackColor       =   16777215
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
         FlatStyle       =   -1  'True
         UseVisualStyle  =   -1  'True
         TextAlignment   =   1
         Appearance      =   17
      End
      Begin XtremeSuiteControls.PushButton btnBloqueo 
         Height          =   360
         Left            =   9960
         TabIndex        =   23
         Top             =   75
         Width           =   3615
         _Version        =   1441793
         _ExtentX        =   6376
         _ExtentY        =   635
         _StockProps     =   79
         Caption         =   "..."
         BackColor       =   16777215
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
         FlatStyle       =   -1  'True
         UseVisualStyle  =   -1  'True
         Appearance      =   17
      End
      Begin XtremeSuiteControls.PushButton btnFavoritos 
         Height          =   360
         Index           =   8
         Left            =   8520
         TabIndex        =   22
         Top             =   75
         Visible         =   0   'False
         Width           =   495
         _Version        =   1441793
         _ExtentX        =   873
         _ExtentY        =   635
         _StockProps     =   79
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FlatStyle       =   -1  'True
         UseVisualStyle  =   -1  'True
         Appearance      =   17
      End
      Begin XtremeSuiteControls.PushButton btnFavoritos 
         Height          =   360
         Index           =   7
         Left            =   8040
         TabIndex        =   21
         Top             =   75
         Visible         =   0   'False
         Width           =   495
         _Version        =   1441793
         _ExtentX        =   873
         _ExtentY        =   635
         _StockProps     =   79
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FlatStyle       =   -1  'True
         UseVisualStyle  =   -1  'True
         Appearance      =   17
      End
      Begin XtremeSuiteControls.PushButton btnFavoritos 
         Height          =   360
         Index           =   6
         Left            =   7560
         TabIndex        =   20
         Top             =   75
         Visible         =   0   'False
         Width           =   495
         _Version        =   1441793
         _ExtentX        =   873
         _ExtentY        =   635
         _StockProps     =   79
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FlatStyle       =   -1  'True
         UseVisualStyle  =   -1  'True
         Appearance      =   17
      End
      Begin XtremeSuiteControls.PushButton btnFavoritos 
         Height          =   360
         Index           =   5
         Left            =   7080
         TabIndex        =   19
         Top             =   75
         Visible         =   0   'False
         Width           =   495
         _Version        =   1441793
         _ExtentX        =   873
         _ExtentY        =   635
         _StockProps     =   79
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FlatStyle       =   -1  'True
         UseVisualStyle  =   -1  'True
         Appearance      =   17
      End
      Begin XtremeSuiteControls.PushButton btnFavoritos 
         Height          =   360
         Index           =   4
         Left            =   6600
         TabIndex        =   18
         Top             =   75
         Visible         =   0   'False
         Width           =   495
         _Version        =   1441793
         _ExtentX        =   873
         _ExtentY        =   635
         _StockProps     =   79
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FlatStyle       =   -1  'True
         UseVisualStyle  =   -1  'True
         Appearance      =   17
      End
      Begin XtremeSuiteControls.PushButton btnFavoritos 
         Height          =   360
         Index           =   3
         Left            =   6120
         TabIndex        =   17
         Top             =   75
         Visible         =   0   'False
         Width           =   495
         _Version        =   1441793
         _ExtentX        =   873
         _ExtentY        =   635
         _StockProps     =   79
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FlatStyle       =   -1  'True
         UseVisualStyle  =   -1  'True
         Appearance      =   17
      End
      Begin XtremeSuiteControls.PushButton btnFavoritos 
         Height          =   360
         Index           =   2
         Left            =   5640
         TabIndex        =   16
         Top             =   75
         Visible         =   0   'False
         Width           =   495
         _Version        =   1441793
         _ExtentX        =   873
         _ExtentY        =   635
         _StockProps     =   79
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FlatStyle       =   -1  'True
         UseVisualStyle  =   -1  'True
         Appearance      =   17
      End
      Begin XtremeSuiteControls.PushButton btnFavoritos 
         Height          =   360
         Index           =   1
         Left            =   5160
         TabIndex        =   15
         Top             =   75
         Visible         =   0   'False
         Width           =   495
         _Version        =   1441793
         _ExtentX        =   873
         _ExtentY        =   635
         _StockProps     =   79
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FlatStyle       =   -1  'True
         UseVisualStyle  =   -1  'True
         Appearance      =   17
      End
      Begin XtremeSuiteControls.PushButton btnFavoritos 
         Height          =   360
         Index           =   0
         Left            =   4680
         TabIndex        =   14
         Top             =   75
         Visible         =   0   'False
         Width           =   495
         _Version        =   1441793
         _ExtentX        =   873
         _ExtentY        =   635
         _StockProps     =   79
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FlatStyle       =   -1  'True
         UseVisualStyle  =   -1  'True
         Appearance      =   17
      End
      Begin XtremeSuiteControls.PushButton btnMarcas 
         Height          =   360
         Left            =   3240
         TabIndex        =   13
         ToolTipText     =   "Registro de Marcas"
         Top             =   75
         Width           =   1335
         _Version        =   1441793
         _ExtentX        =   2355
         _ExtentY        =   635
         _StockProps     =   79
         Caption         =   "Marcas"
         ForeColor       =   0
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FlatStyle       =   -1  'True
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         Picture         =   "MDIPrincipal.frx":91F8
         PushButtonStyle =   2
      End
      Begin XtremeSuiteControls.PushButton btnMenu 
         Height          =   360
         Index           =   3
         Left            =   2400
         TabIndex        =   12
         ToolTipText     =   "Configuracion de Impresoras del sistema"
         Top             =   75
         Width           =   855
         _Version        =   1441793
         _ExtentX        =   1508
         _ExtentY        =   635
         _StockProps     =   79
         Caption         =   "Prn"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FlatStyle       =   -1  'True
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         Picture         =   "MDIPrincipal.frx":95D6
      End
      Begin XtremeSuiteControls.PushButton btnMenu 
         Height          =   360
         Index           =   2
         Left            =   1680
         TabIndex        =   11
         ToolTipText     =   "Explorador Activos Fijos"
         Top             =   80
         Width           =   735
         _Version        =   1441793
         _ExtentX        =   1296
         _ExtentY        =   635
         _StockProps     =   79
         Caption         =   "AF"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FlatStyle       =   -1  'True
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         Picture         =   "MDIPrincipal.frx":99AE
      End
      Begin XtremeSuiteControls.PushButton btnMenu 
         Height          =   360
         Index           =   1
         Left            =   960
         TabIndex        =   10
         ToolTipText     =   "Explorador Contable"
         Top             =   80
         Width           =   735
         _Version        =   1441793
         _ExtentX        =   1296
         _ExtentY        =   635
         _StockProps     =   79
         Caption         =   "Cnt"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FlatStyle       =   -1  'True
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         Picture         =   "MDIPrincipal.frx":9D7F
      End
      Begin XtremeSuiteControls.PushButton btnMenu 
         Height          =   360
         Index           =   0
         Left            =   120
         TabIndex        =   9
         ToolTipText     =   "Menú Principal"
         Top             =   80
         Width           =   855
         _Version        =   1441793
         _ExtentX        =   1508
         _ExtentY        =   635
         _StockProps     =   79
         Caption         =   "Menú"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FlatStyle       =   -1  'True
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         Picture         =   "MDIPrincipal.frx":A176
      End
   End
   Begin VB.Timer Timer_Load 
      Interval        =   5
      Left            =   600
      Top             =   2010
   End
   Begin VB.Timer TimerSalir 
      Left            =   240
      Top             =   2040
   End
   Begin MSComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   8685
      Width           =   15615
      _ExtentX        =   27543
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   7
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Bevel           =   0
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   "01:50:a. m."
            Object.ToolTipText     =   "Hora"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Bevel           =   0
            Enabled         =   0   'False
            Object.Width           =   989
            MinWidth        =   989
            TextSave        =   "CAPS"
            Object.ToolTipText     =   "Tecla de CAP activa"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Bevel           =   0
            Object.Width           =   988
            MinWidth        =   988
            TextSave        =   "NUM"
            Object.ToolTipText     =   "NumLock"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   3177
            MinWidth        =   3177
            Object.ToolTipText     =   "Usuario Activo"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   8469
            MinWidth        =   8469
            Object.ToolTipText     =   "Fecha de Auxiliar.: "
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   8467
            MinWidth        =   8467
            Object.ToolTipText     =   "Fecha Actual.:"
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.ToolTipText     =   "Tema"
         EndProperty
      EndProperty
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
   Begin MSComctlLib.ImageList imgMenuLista 
      Left            =   240
      Top             =   1320
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   38
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrincipal.frx":A569
            Key             =   "Root"
            Object.Tag             =   "Root"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrincipal.frx":A95C
            Key             =   "Reloj"
            Object.Tag             =   "Reloj"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrincipal.frx":AD65
            Key             =   "Ayuda"
            Object.Tag             =   "Ayuda"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrincipal.frx":B14E
            Key             =   "Calendario"
            Object.Tag             =   "Calendario"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrincipal.frx":B543
            Key             =   "Dinero 2"
            Object.Tag             =   "Dinero 2"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrincipal.frx":B923
            Key             =   "Contacto"
            Object.Tag             =   "Contacto"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrincipal.frx":BD10
            Key             =   "Direccion"
            Object.Tag             =   "Direccion"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrincipal.frx":C0F8
            Key             =   "Libros"
            Object.Tag             =   "Libros"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrincipal.frx":C4D2
            Key             =   "Printer 2"
            Object.Tag             =   "Printer 2"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrincipal.frx":C8AA
            Key             =   "Exportar"
            Object.Tag             =   "Exportar"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrincipal.frx":CCA4
            Key             =   "Agenda"
            Object.Tag             =   "Agenda"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrincipal.frx":D075
            Key             =   "Lupa"
            Object.Tag             =   "Lupa"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrincipal.frx":D468
            Key             =   "Carpeta"
            Object.Tag             =   "Carpeta"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrincipal.frx":D833
            Key             =   "Administrador"
            Object.Tag             =   "Administrador"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrincipal.frx":DC17
            Key             =   "Favorito Add"
            Object.Tag             =   "Favorito Add"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrincipal.frx":DFFA
            Key             =   "Ajustes"
            Object.Tag             =   "Ajustes"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrincipal.frx":E3E4
            Key             =   "Dinero 3"
            Object.Tag             =   "Dinero 3"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrincipal.frx":E7BE
            Key             =   "Documento"
            Object.Tag             =   "Documento"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrincipal.frx":EBA4
            Key             =   "Dinero 1"
            Object.Tag             =   "Dinero 1"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrincipal.frx":EF82
            Key             =   "Grafico"
            Object.Tag             =   "Grafico"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrincipal.frx":F360
            Key             =   "Seguridad"
            Object.Tag             =   "Seguridad"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrincipal.frx":F743
            Key             =   "Compras"
            Object.Tag             =   "Compras"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrincipal.frx":FB24
            Key             =   "Aplicacion"
            Object.Tag             =   "Aplicacion"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrincipal.frx":FF09
            Key             =   "Aplicaciones"
            Object.Tag             =   "Aplicaciones"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrincipal.frx":102EE
            Key             =   "Configuracion"
            Object.Tag             =   "Configuracion"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrincipal.frx":106B8
            Key             =   "Estadistica"
            Object.Tag             =   "Estadistica"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrincipal.frx":10A99
            Key             =   "Analisis"
            Object.Tag             =   "Analisis"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrincipal.frx":10E81
            Key             =   "Explorer"
            Object.Tag             =   "Explorer"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrincipal.frx":11278
            Key             =   "Opciones"
            Object.Tag             =   "Opciones"
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrincipal.frx":11650
            Key             =   "Histograma"
            Object.Tag             =   "Histograma"
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrincipal.frx":11A3C
            Key             =   "Usuario"
            Object.Tag             =   "Usuario"
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrincipal.frx":11E22
            Key             =   "Identificacion"
            Object.Tag             =   "Identificacion"
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrincipal.frx":1220B
            Key             =   "Caja Fuerte"
            Object.Tag             =   "Caja Fuerte"
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrincipal.frx":125EA
            Key             =   "Buscar"
            Object.Tag             =   "Buscar"
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrincipal.frx":129E0
            Key             =   "FastFoward"
            Object.Tag             =   "FastFoward"
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrincipal.frx":12DE6
            Key             =   "Cajas"
            Object.Tag             =   "Cajas"
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrincipal.frx":131B3
            Key             =   "Reportes"
            Object.Tag             =   "Reportes"
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrincipal.frx":1358F
            Key             =   "Aprobacion"
            Object.Tag             =   "Aprobacion"
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuArchivo 
      Caption         =   "&Archivo"
      Begin VB.Menu mnuSeguridad 
         Caption         =   "Seguridad"
         Begin VB.Menu mnuSeguridadSub 
            Caption         =   "Cambiar Contraseña"
            Index           =   0
         End
         Begin VB.Menu mnuSeguridadSub 
            Caption         =   "Actualiza Datos de Contacto"
            Index           =   1
         End
         Begin VB.Menu mnuSeguridadSub 
            Caption         =   "-"
            Index           =   2
         End
         Begin VB.Menu mnuSeguridadSub 
            Caption         =   "Cambiar de Tema"
            Index           =   3
         End
         Begin VB.Menu mnuSeguridadSub 
            Caption         =   "-"
            Index           =   4
         End
         Begin VB.Menu mnuSeguridadSub 
            Caption         =   "Bitácora"
            Index           =   5
         End
      End
      Begin VB.Menu mnuArchivoSeparador1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuParametrosSistemaMenu 
         Caption         =   "Parámetros del Sistema"
         Begin VB.Menu mnuParametrosSistema 
            Caption         =   "Datos de Empresa"
            Index           =   0
         End
         Begin VB.Menu mnuParametrosSistema 
            Caption         =   "Comunicados Generales"
            Index           =   1
         End
         Begin VB.Menu mnuParametrosSistema 
            Caption         =   "-"
            Index           =   2
         End
         Begin VB.Menu mnuParametrosSistema 
            Caption         =   "Encabezado y Pie de Página Estados"
            Index           =   3
         End
         Begin VB.Menu mnuParametrosSistema 
            Caption         =   "-"
            Index           =   4
         End
         Begin VB.Menu mnuParametrosSistema 
            Caption         =   "Consulta de Cola de Asientos"
            Index           =   5
         End
         Begin VB.Menu mnuParametrosSistema 
            Caption         =   "-"
            Index           =   6
         End
         Begin VB.Menu mnuParametrosSistema 
            Caption         =   "Oficinas"
            Index           =   7
         End
         Begin VB.Menu mnuParametrosSistema 
            Caption         =   "Oficinas : Metas de Colocación"
            Index           =   8
         End
         Begin VB.Menu mnuParametrosSistema 
            Caption         =   "-"
            Index           =   9
         End
         Begin VB.Menu mnuParametrosSistema 
            Caption         =   "Variables Globales"
            Index           =   10
         End
         Begin VB.Menu mnuPE_Modo_1 
            Caption         =   "Planilla Especiales (Directas)"
         End
      End
      Begin VB.Menu mnuArchivoSeparador2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCambioEmpresa 
         Caption         =   "Cambio de Empresa"
      End
      Begin VB.Menu mnuArchivoSeparador21 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDashboard 
         Caption         =   "Dashboard Empresarial"
      End
      Begin VB.Menu mnuDashboard_Asociados 
         Caption         =   "Dashboard Asociados"
      End
      Begin VB.Menu mnuArchivoSeparador22 
         Caption         =   "-"
      End
      Begin VB.Menu mnuColaborador 
         Caption         =   "Portal del Colaborador"
      End
      Begin VB.Menu mnuArchivoSeparador23 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSalir 
         Caption         =   "&Salir"
      End
   End
   Begin VB.Menu mnuVer 
      Caption         =   "Ver"
      Begin VB.Menu mnuVerSub 
         Caption         =   "Ordenar por Iconos"
         Index           =   0
      End
      Begin VB.Menu mnuVerSub 
         Caption         =   "Ordenar en Cascada"
         Index           =   1
      End
      Begin VB.Menu mnuVerSub 
         Caption         =   "Ordenar Vertical"
         Index           =   2
      End
      Begin VB.Menu mnuVerSub 
         Caption         =   "Odenar Horizontal"
         Index           =   3
      End
      Begin VB.Menu mnuVerSub 
         Caption         =   "-"
         Index           =   4
      End
      Begin VB.Menu mnuVerSub 
         Caption         =   "Cerrar todas las ventanas"
         Index           =   5
      End
      Begin VB.Menu mnuVerSub 
         Caption         =   "Minimizar todas las ventanas"
         Index           =   6
      End
      Begin VB.Menu mnuVerSub 
         Caption         =   "Restaurar todas la ventanas"
         Index           =   7
      End
      Begin VB.Menu mnuVerSub 
         Caption         =   ""
         Index           =   8
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuAyuda 
      Caption         =   "&?"
      Begin VB.Menu mnuAyudaContenido 
         Caption         =   "Contenido"
      End
      Begin VB.Menu mnuAyudaSoporteTecnico 
         Caption         =   "Soporte Técnico"
      End
      Begin VB.Menu mnuBarraHerramientas 
         Caption         =   "Barra de Herramientas"
      End
      Begin VB.Menu mnuAyudaSeparador1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAyudaAcercaDe 
         Caption         =   "Acerca De..."
      End
   End
   Begin VB.Menu mnuAcciones 
      Caption         =   "Acciones"
      Visible         =   0   'False
      Begin VB.Menu mnuAccionesSub 
         Caption         =   "Abonos"
         Index           =   0
      End
      Begin VB.Menu mnuAccionesSub 
         Caption         =   "Anulación"
         Index           =   1
      End
      Begin VB.Menu mnuAccionesSub 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnuAccionesSub 
         Caption         =   "Gestión de Cobro"
         Index           =   3
      End
      Begin VB.Menu mnuAccionesSub 
         Caption         =   "Estado de la Operación"
         Index           =   4
      End
      Begin VB.Menu mnuAccionesSub 
         Caption         =   "-"
         Index           =   5
      End
      Begin VB.Menu mnuAccionesSub 
         Caption         =   "Nuevo Crédito"
         Index           =   6
      End
      Begin VB.Menu mnuAccionesSub 
         Caption         =   "Trámites"
         Index           =   7
      End
      Begin VB.Menu mnuAccionesSub 
         Caption         =   "-"
         Index           =   8
         Visible         =   0   'False
      End
      Begin VB.Menu mnuAccionesSub 
         Caption         =   "Historial"
         Index           =   9
         Visible         =   0   'False
      End
      Begin VB.Menu mnuAccionesSub 
         Caption         =   "-"
         Index           =   10
      End
      Begin VB.Menu mnuAccionesSub 
         Caption         =   "Nuevo Análisis"
         Index           =   11
      End
      Begin VB.Menu mnuAccionesSub 
         Caption         =   "Análisis ?"
         Index           =   12
      End
      Begin VB.Menu mnuAccionesSub 
         Caption         =   "-"
         Index           =   13
      End
      Begin VB.Menu mnuAccionesSub 
         Caption         =   "Plan de Pagos"
         Index           =   14
      End
      Begin VB.Menu mnuAccionesSub 
         Caption         =   "-"
         Index           =   15
      End
      Begin VB.Menu mnuAccionesSub 
         Caption         =   "Cerrar"
         Index           =   16
      End
   End
   Begin VB.Menu mnuCxC 
      Caption         =   "CxC"
      Visible         =   0   'False
      Begin VB.Menu mnuCxCSub 
         Caption         =   "Abonos"
         Index           =   0
      End
      Begin VB.Menu mnuCxCSub 
         Caption         =   "Anulación"
         Index           =   1
      End
      Begin VB.Menu mnuCxCSub 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnuCxCSub 
         Caption         =   "Nueva Operación"
         Index           =   3
      End
      Begin VB.Menu mnuCxCSub 
         Caption         =   "Tramite"
         Index           =   4
      End
      Begin VB.Menu mnuCxCSub 
         Caption         =   "-"
         Index           =   5
      End
      Begin VB.Menu mnuCxCSub 
         Caption         =   "Movimientos"
         Index           =   6
      End
      Begin VB.Menu mnuCxCSub 
         Caption         =   "Plan de Pagos"
         Index           =   7
      End
      Begin VB.Menu mnuCxCSub 
         Caption         =   "-"
         Index           =   8
      End
      Begin VB.Menu mnuCxCSub 
         Caption         =   "Cerrar"
         Index           =   9
      End
   End
   Begin VB.Menu mnuExplorerContable 
      Caption         =   "Explorador: Contable"
      Visible         =   0   'False
      Begin VB.Menu mnuCntAccionEditar 
         Caption         =   "&Editar"
      End
      Begin VB.Menu mnuCntAccionBorrar 
         Caption         =   "&Borrar"
      End
      Begin VB.Menu mnuCntSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCntAccionRefrescar 
         Caption         =   "Refrescar"
      End
      Begin VB.Menu mnuCntSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCntAccionesImprimir 
         Caption         =   "Imprimir"
      End
      Begin VB.Menu mnuCntSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCntAccionesMayorizar 
         Caption         =   "Mayorizar"
      End
   End
   Begin VB.Menu mnuActivosExplorador 
      Caption         =   "Explorardor: Activos"
      Visible         =   0   'False
      Begin VB.Menu mnuActivosAccionNuevo 
         Caption         =   "Nuevo"
      End
      Begin VB.Menu mnuActivosAccionPropiedades 
         Caption         =   "Propiedades"
      End
      Begin VB.Menu mnuActivosAccionEliminar 
         Caption         =   "Eliminar"
      End
      Begin VB.Menu mnuActivosAccionSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuActivosAccionDepreciacion 
         Caption         =   "Depreciación"
      End
      Begin VB.Menu mnuActivosAccionSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuActivosAccionActualizar 
         Caption         =   "Actualizar"
      End
      Begin VB.Menu mnuActivosAccionImprimir 
         Caption         =   "Imprimir"
      End
   End
   Begin VB.Menu mnuMarcas 
      Caption         =   "Marcas"
      Visible         =   0   'False
      Begin VB.Menu mnuMarcaOpcion 
         Caption         =   "Registro de Marca"
         Index           =   0
      End
      Begin VB.Menu mnuMarcaOpcion 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnuMarcaOpcion 
         Caption         =   "Bitácora de Marcas"
         Index           =   2
      End
      Begin VB.Menu mnuMarcaOpcion 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu mnuMarcaOpcion 
         Caption         =   "Configuración"
         Index           =   4
      End
      Begin VB.Menu mnuMarcaOpcion 
         Caption         =   "Asignación de Usuario"
         Index           =   5
      End
   End
   Begin VB.Menu mnuContaRevision 
      Caption         =   "Revisión de Contabilidad"
      Visible         =   0   'False
      Begin VB.Menu mnuContaRevisionSep 
         Caption         =   "Revisión de Balance "
         Index           =   0
      End
      Begin VB.Menu mnuContaRevisionSep 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnuContaRevisionSep 
         Caption         =   "Verificación de Asientos"
         Index           =   2
      End
   End
   Begin VB.Menu mnuContaCierre 
      Caption         =   "Cierres Contables"
      Visible         =   0   'False
      Begin VB.Menu mnuCierreContable 
         Caption         =   "Periodo"
         Index           =   0
      End
      Begin VB.Menu mnuCierreContable 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnuCierreContable 
         Caption         =   "Asiento de Cierre Fiscal"
         Index           =   2
      End
   End
End
Attribute VB_Name = "MDIPrincipal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mLoad_Inicial As Boolean


Private Sub PopupButtonMenu(btn As PushButton)
    
Select Case btn.Caption
    Case "Cierres"
        Me.PopupMenu mnuContaCierre, , btn.Left + btn.Width, tpContabilidad.top + btn.top
    Case "Revisión"
        Me.PopupMenu mnuContaRevision, , btn.Left + btn.Width, tpContabilidad.top + btn.top

    Case "Marcas"
        Me.PopupMenu mnuMarcas, , btn.Left + btn.Width, tpMain.top + btn.top

End Select

End Sub


Private Sub btnContabilidad_Click(Index As Integer)

Select Case Index
    Case 0 'Periodos
    
        Call sbFormsCall("frmCntX_Periodos", 1, , , False)
    
        txtMes.Text = gCntX_Parametros.PeriodoMes
        txtAnio.Text = gCntX_Parametros.PeriodoAnio

    Case 3 'Contabilidad

      gCntX_Parametros.MuestraTodas = True
      
      Call sbFormsCall("frmCntX_Seleccionar", 1, , , False)
      
      txtAnio = gCntX_Parametros.PeriodoAnio
      txtMes = gCntX_Parametros.PeriodoMes
      
      btnContabilidad(Index).Caption = gCntX_Parametros.NombreEmpresa
    
      Dim frm As Form
       
      Call sbFormsCall("frmCntX_Explorer")
      Call sbFormActivo("frmCntX_Explorer", frm)
        
      Call frm.sbRefrescaArbol
End Select

End Sub

Private Sub btnContabilidad_DropDown(Index As Integer)

Select Case Index
    Case 1 'Cierres
        Call PopupButtonMenu(btnContabilidad(Index))
    
    Case 2 'Revision
        Call PopupButtonMenu(btnContabilidad(Index))
End Select

End Sub

Private Sub btnFavoritos_Click(Index As Integer)
Call sbSIFMenuOptionClick(btnFavoritos(Index).Tag)
End Sub

Private Sub btnMarcas_DropDown()
    Call PopupButtonMenu(btnMarcas)
End Sub

Private Sub btnMenu_Click(Index As Integer)
Dim i As Integer, frmX As Form


Call sbFormActivo("frmCntX_Explorer", frmX)
If Not frmX Is Nothing Then
   frmX.Hide
Else
End If

Call sbFormActivo("frmActivos_Explorador", frmX)
If Not frmX Is Nothing Then
   frmX.Hide
Else
End If

Select Case Index
    Case 0 'Menu Principal

        Call sbFormsCall("frmMenu", 0, 1, 1)
        
    Case 1 'Explorer Contabilidad
        If Not tpContabilidad.Visible Then
            tpContabilidad.Visible = True
            
'            txtAnio.Visible = True
'            txtMes.Visible = True
              
        Else
            tpContabilidad.Visible = False
'            txtAnio.Visible = False
'            txtMes.Visible = False
        End If
        
        Call sbFormsCall("frmCntX_Explorer")
        Call sbFormActivo("frmCntX_Explorer", frmX)
        If Not frmX Is Nothing Then
          frmX.WindowState = vbMaximized
        Else
        End If

        
    Case 2 'Explorer Activos Fijos
    
        Call sbFormsCall("frmActivos_Explorador")
        Call sbFormActivo("frmActivos_Explorador", frmX)
        If Not frmX Is Nothing Then
          frmX.WindowState = vbMaximized
        End If
  
  Case 3 'Configuracion Impresoras
    Call sbFormsCall("frmCC_Impresoras")
  
'  Case 4 'Marcas
'        Dim Marcas As clsMarcas
'
'        Set Marcas = New clsMarcas
'                Call Marcas.Menu(glogon.Usuario, glogon.Conection, glogon.BaseDatos, glogon.Servidor _
'                    , App.Path, glogon.ConectRPT, 1, glogon.AppName, glogon.AppVersion, glogon.Maquina _
'                    , glogon.Portal_Con, glogon.Portal_User, glogon.Portal_Key, gPortal.Empresa_Id, gPortal.Empresa_Name)
'        Set Marcas = Nothing
End Select


End Sub


Private Sub MDIForm_Load()

mLoad_Inicial = True

'Me.BackColor = RGB(36, 113, 163)

Me.BackColor = RGB(78, 111, 178)


btnMenu.Item(0).ForeColor = RGB(78, 111, 178)
btnMenu.Item(1).ForeColor = RGB(78, 111, 178)
btnMenu.Item(2).ForeColor = RGB(78, 111, 178)
btnMenu.Item(3).ForeColor = RGB(78, 111, 178)

btnMarcas.ForeColor = RGB(78, 111, 178)

btnContabilidad.Item(1).ForeColor = RGB(78, 111, 178)
btnContabilidad.Item(2).ForeColor = RGB(78, 111, 178)

StatusBar.Panels(7).Text = glogon.ProGrX_Theme

If glogon.AppStatus = 1 Then
   Call sbFormsCall("frmCC_AppStatus", , , , False, Me)
End If

End Sub


Private Sub MDIForm_Resize()
On Error Resume Next

btnEmpresa.Left = Me.Width - (btnEmpresa.Width + 250)
btnContabilidad(3).Left = Me.Width - (btnContabilidad(3).Width + 250)


End Sub

Private Sub mnuCambioEmpresa_Click()
 Call Main_Reload
End Sub

Private Sub mnuCierreContable_Click(Index As Integer)
Dim iRespuesta As Integer, frmX As Form

Select Case Index
  Case 0 'CierrePeriodo
    iRespuesta = MsgBox("Esta seguro que desea Cerrar este periodo...", vbYesNo)
    If iRespuesta = vbYes Then
        'Reestructura Movimientos
        Set frmX = frmCntX_Procesos
        Call sbCntX_RestructuraMovimientosRSM(txtAnio.Text, txtMes.Text, frmX, False)
        
        'Cierra Periodo (Mensual)
        Me.MousePointer = vbHourglass
            Call sbCntX_PeriodoCierre(txtAnio.Text, txtMes.Text)
        Me.MousePointer = vbDefault
    End If
    
  Case 2 'Asientos de Cierres Fiscal
    iRespuesta = MsgBox("Esta seguro que desea generar Asientos de Cierre Fiscal...", vbYesNo)
    If iRespuesta = vbYes Then
      Set frmX = frmCntX_Procesos
     'No se reestructuran los movimientos porque Para los Asientos de Cierre Fiscal, ya tuvo que realizar el cierre del periodo
     ' Call sbCntX_RestructuraMovimientosRSM(MDIMenu.txtAnio, MDIMenu.txtMes, frmX, False)
      Call sbCntX_CierreFiscal(frmX, txtMes.Text, txtAnio.Text)
    End If

End Select

End Sub

'--- Menu Contextual: Contabilidad

Private Sub mnuCntAccionEditar_Click()
Dim frmX As Form
Call sbFormActivo("frmCntX_Explorer", frmX)
Call frmX.sbButtonPopUp(1)
End Sub

Private Sub mnuCntAccionBorrar_Click()
Dim frmX As Form
Call sbFormActivo("frmCntX_Explorer", frmX)
Call frmX.sbButtonPopUp(2)
End Sub


Private Sub mnuCntAccionesImprimir_Click()
Dim frmX As Form
Call sbFormActivo("frmCntX_Explorer", frmX)
Call frmX.sbButtonPopUp(4)
End Sub

Private Sub mnuCntAccionesMayorizar_Click()
Dim frmX As Form
Call sbFormActivo("frmCntX_Explorer", frmX)
Call frmX.sbButtonPopUp(5)
End Sub

Private Sub mnuCntAccionRefrescar_Click()
Dim frmX As Form
Call sbFormActivo("frmCntX_Explorer", frmX)
Call frmX.sbButtonPopUp(3)
End Sub

'--- Fin: Menu Contextual: Contabilidad



'--- Menu Contextual: Activos Fijos

Private Sub mnuActivosAccionNuevo_Click()
Dim frmX As Form
Call sbFormActivo("frmActivos_Explorador", frmX)
Call frmX.btnAcciones_Click(5)
End Sub

Private Sub mnuActivosAccionPropiedades_Click()
Dim frmX As Form

Call sbFormActivo("frmActivos_Explorador", frmX)
Call frmX.btnAcciones_Click(6)

End Sub

Private Sub mnuActivosAccionEliminar_Click()
Dim frmX As Form

Call sbFormActivo("frmActivos_Explorador", frmX)
Call frmX.btnAcciones_Click(6)

End Sub


Private Sub mnuActivosAccionDepreciacion_Click()
Dim frmX As Form

Call sbFormActivo("frmActivos_Explorador", frmX)

Call frmX.btnAcciones_Click(3)


End Sub

Private Sub mnuActivosAccionActualizar_Click()
Dim frmX As Form

Call sbFormActivo("frmActivos_Explorador", frmX)
Call frmX.btnAcciones_Click(0)

End Sub

Private Sub mnuActivosAccionImprimir_Click()
Dim frmX As Form

Call sbFormActivo("frmActivos_Explorador", frmX)

Call frmX.btnAcciones_Click(1)

End Sub

'---Fin de Menu Contextual: Activos Fijos

Private Sub sbDashboard_Load()
'Dim frm As Form


Call sbFormsCall("frmDSB_Dashboard")

'Call sbFormActivo("frmDSB_Dashboard", frm)
'
'
'frm.top = 240
'frm.Left = Me.Width - (frm.Width + 650)

End Sub


Private Sub MDIForm_Activate()
  
Me.Caption = App.ProductName & " [ " & App.Major & "." & App.Minor & "." & App.Revision & ".r" & GLOBALES.SysVersion & " ]"

txtAnio = gCntX_Parametros.PeriodoAnio
txtMes = gCntX_Parametros.PeriodoMes

btnContabilidad(3).Caption = gCntX_Parametros.NombreEmpresa

End Sub



Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   If UnloadMode = 0 Then
      Cancel = True
      TimerSalir.Interval = 10
   End If
End Sub


Public Sub mnuAccionesSub_Click(Index As Integer)
Dim strSQL As String, rs As New ADODB.Recordset
Dim x As clsEstudioCrd, vOperacion As Long
Dim vExpediente As String, vCajas As Boolean
Dim frmConsultaActiva As Form, frm As Form

On Error GoTo vError

vOperacion = 0
vExpediente = ""
vCajas = False

'Localiza el Frm de Consulta que se encuentre activo para utilizarlo de referencia
For Each frmConsultaActiva In Forms
  If (UCase(frmConsultaActiva.Name) = UCase("frmCR_ConsultaCreditos")) Then
    'lo localiza y sale (se supone que este procedimiento solo puede ser abierto desde esta pantalla
    Exit For
  End If
Next frmConsultaActiva

'Validacion

With frmConsultaActiva.vgCreditos
    .Sheet = .ActiveSheet
    .Row = .ActiveRow
    
    Select Case .Sheet
       Case 1 'Activos
          .col = 3
          If Index = 6 Or Index = 11 Then
            'Nada
           Else
              If Not IsNumeric(.Text) Then Exit Sub
              vOperacion = .Text
           End If
        
       Case 2, 3 'Cancelados y En Tramite
          .col = 3
          
          If Index = 6 Or Index = 11 Then
            'Nada
           Else
              If Not IsNumeric(.Text) Then Exit Sub
              vOperacion = .Text
           End If
       
       Case 4 'PreAnalisis
          .col = 2
          vExpediente = .Text
          .col = 7
          If IsNumeric(.Text) Then vOperacion = .Text
       
       Case 5 'Incobrables
          .col = 2
          If Not IsNumeric(.Text) Then Exit Sub
          vOperacion = .Text
      
    End Select
    

    Select Case Index
      Case 0 'Abonos
            If vOperacion = 0 Then Exit Sub
            .col = 7 'Saldo
            If CCur(.Text) = 0 Then Exit Sub
            
            vCajas = IIf((fxCajasParametros("01") = "S"), True, False)
                
                .col = 19 'Cuotas Morosas
                If CInt(.Text) = 0 Then
                  If vCajas Then
                        ModuloCajas.mRef_01 = vOperacion
                         
                        If GLOBALES.SysPlanPagos = 1 Then
                                 Call sbFormsCall("frmCajas_Crd_AbonosCtP", vbModal, 0, 0, False, Me, True)
                        Else
                                 Call sbFormsCall("frmCajas_Crd_AbonosStP", vbModal, 0, 0, False, Me, True)
                        End If
                  
                  Else
                        If GLOBALES.SysPlanPagos = 1 Then
                                 Call sbFormsCall("frmCR_AbonosNew", vbModal, 0, 0, False, Me, True)
                        Else
                                 Call sbFormsCall("frmCR_Abonos", vbModal, 0, 0, False, Me, True)
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
                                 Call sbFormsCall("frmCajas_Crd_AbonosCtP", vbModal, 0, 0, False, Me, True)
                        Else
                                 Call sbFormsCall("frmCajas_Crd_AbonosStP", vbModal, 0, 0, False, Me, True)
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
    
      
      Case 1 'Anulacion de Abonos
            If vOperacion = 0 Then Exit Sub
                
                If GLOBALES.SysPlanPagos = 1 Then
                            Call sbFormsCall("frmCR_AnulaAbonosNew", 0, 0, 0, False, Me, True)
                Else
                            Call sbFormsCall("frmCR_AnulaAbonos", 0, 0, 0, False, Me, True)
                End If
      
                            For Each frm In Forms
                              If (UCase(frm.Name) = UCase("frmCR_AnulaAbonos")) Or (UCase(frm.Name) = UCase("frmCR_AnulaAbonosNew")) Then
                                Call frm.sbConsultaExterna(vOperacion)
                                Exit For
                              End If
                            Next frm
      
      Case 2 'Sep
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
                    .StoredProcParam(1) = Format(frmConsultaActiva.dtpCorte.Value, "yyyy/mm/dd")
                    
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
                     .StoredProcParam(1) = Format(frmConsultaActiva.dtpCorte.Value, "yyyy/mm/dd")
                
                End If

                .PrintReport
     
               
            End With
            Me.MousePointer = vbDefault
            
      Case 5 'Sep
      Case 6 'Nuevo Credito
                GLOBALES.gCedulaActual = frmConsultaActiva.txtCedula.Text
                Call sbFormsCall("frmCR_SeguimientoTramites")
                For Each frm In Forms
                  If UCase(frm.Name) = UCase("frmCR_SeguimientoTramites") Then
                    Call frm.sbGXSegTraIniTlb
                    Exit For
                  End If
                Next frm
      
      Case 7 'Seguimiento de Tramites
            If vOperacion = 0 Then Exit Sub
                
                Call sbFormsCall("frmCR_SeguimientoTramites")
                For Each frm In Forms
                  If UCase(frm.Name) = UCase("frmCR_SeguimientoTramites") Then
                    Call frm.sbConsultaExterna(vOperacion)
                    Exit For
                  End If
                Next frm
      
      Case 8 'Sep
      
      Case 9 'Historial
            If vOperacion = 0 Then Exit Sub
                
                Call sbFormsCall("frmCR_ConsultaOperaciones")
                For Each frm In Forms
                  If UCase(frm.Name) = UCase("frmCR_ConsultaOperaciones") Then
                    frm.optTipo.Item(0).Value = True
                    Call frm.sbConsultaExterna(vOperacion)
                    Exit For
                  End If
                Next frm
      
      Case 10 'Sep
      Case 11 'Nuevo PreAnalisis
      
            Set x = New clsEstudioCrd
            Set x.vCon = glogon.Conection
            x.xOperacion = vOperacion
            x.xkey = glogon.ConectRPT
      
            x.vSolicitudPreanalisis = 0
            x.vCedula = frmConsultaActiva.txtCedula.Text
    
            Call x.Menu(glogon.Usuario, glogon.Conection, glogon.BaseDatos, glogon.Servidor _
            , App.Path, glogon.ConectRPT, 12, glogon.AppName, glogon.AppVersion, glogon.Maquina _
            , glogon.Portal_Con, glogon.Portal_User, glogon.Portal_Key, gPortal.Empresa_Id, gPortal.Empresa_Name)
    
            Set x = Nothing
      
     
      
      Case 12 'PreAnalisis
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
'                        Call x.Menu(glogon.Usuario, glogon.Conection, glogon.BaseDatos, glogon.Servidor _
'                        , App.Path, glogon.ConectRPT, 2, glogon.AppName, glogon.AppVersion, glogon.Maquina _
'                        , glogon.Portal_Con, glogon.Portal_User, glogon.Portal_Key, gPortal.Empresa_Id, gPortal.Empresa_Name)
                    
                        MsgBox "La Operación No." & vOperacion & " no tiene un estudio de crédito vinculado!", vbInformation
                    Else
                        x.vSolicitudPreanalisis = rs!cod_PreAnalisis
                        Call x.Menu(glogon.Usuario, glogon.Conection, glogon.BaseDatos, glogon.Servidor _
                        , App.Path, glogon.ConectRPT, 11, glogon.AppName, glogon.AppVersion, glogon.Maquina _
                        , glogon.Portal_Con, glogon.Portal_User, glogon.Portal_Key, gPortal.Empresa_Id, gPortal.Empresa_Name)
                    End If
                    rs.Close
                      
            End If
            
      
    
            Set x = Nothing
      
      
      Case 13 'Sep
      
      Case 14 'Plan de Pagos
        If vOperacion = 0 Then Exit Sub
        
        Operacion.OperacionConsulta = vOperacion
        Call sbFormsCall("frmCR_PlanPagos", , , , False, Me)
    
      Case 15 'Sep
      
      Case 16 'Cerrar
      
    End Select

End With

Exit Sub
        
vError:
        Me.MousePointer = vbDefault
        MsgBox fxSys_Error_Handler(Err.Description)


End Sub




Private Sub mnuColaborador_Click()

Call sbFormsCall("frmDSB_Colaborador")

End Sub

Private Sub mnuContaRevisionSep_Click(Index As Integer)


Dim iRespuesta As Integer, frmX As Form

Select Case Index
  
  Case 0 'Revision del Balance
    
    iRespuesta = MsgBox("Esta seguro que desea revisar la balanza de comprobación por inconsistencias?", vbYesNo)
    If iRespuesta = vbYes Then
       Set frmX = frmCntX_Procesos
       Call sbCntX_RestructuraMovimientosRSM(gCntX_Parametros.PeriodoAnio, gCntX_Parametros.PeriodoMes, frmX)
    End If
  
  Case 2 'Verificacion de Asientos
    Call sbFormsCall("frmCntX_UtilVerificaAsientos", vbModal, , , False, Me)

End Select

End Sub

Private Sub mnuCxCSub_Click(Index As Integer)
Dim strSQL As String, rs As New ADODB.Recordset
Dim vOperacion As Long
Dim frmConsultaActiva As Form, frm As Form

On Error GoTo vError

vOperacion = 0

 
'Localiza el Frm de Consulta que se encuentre activo para utilizarlo de referencia
For Each frmConsultaActiva In Forms
  If (UCase(frmConsultaActiva.Name) = UCase("frmCxC_Consulta")) Then
    'lo localiza y sale (se supone que este procedimiento solo puede ser abierto desde esta pantalla
    Exit For
  End If
Next frmConsultaActiva

With frmConsultaActiva.vgCxC
    .Sheet = .ActiveSheet
    .Row = .ActiveRow
    
    Select Case .Sheet
       Case 1 'Activos
          .col = 2
          If Not IsNumeric(.CellTag) Then Exit Sub
          
          vOperacion = .CellTag
       Case 2, 3 'Cancelados y En Tramite
          .col = 2
          If Not IsNumeric(.Text) Then Exit Sub
          vOperacion = .Text
       
    End Select
  

    Select Case Index
      Case 0 'Abonos
            If vOperacion = 0 Then Exit Sub
      
            .col = 6 'Saldo
             If CCur(.Text) = 0 Then Exit Sub

                    Call sbFormsCall("frmCxC_CuentasAbonos", , , , , Me, True)

                    For Each frm In Forms
                      If (UCase(frm.Name) = UCase("frmCxC_CuentasAbonos")) Then
                        Call frm.sbConsultaExterna(vOperacion)
                        Exit For
                      End If
                    Next frm
    
      
      Case 1 'Anulacion de Abonos
            If vOperacion = 0 Then Exit Sub
                
            Call sbFormsCall("frmCxC_CuentasAnulaciones")
            For Each frm In Forms
              If (UCase(frm.Name) = UCase("frmCxC_CuentasAnulaciones")) Then
                Call frm.sbConsultaExterna(vOperacion)
                Exit For
              End If
            Next frm
      
      Case 2 'Sep
      
     
     
      Case 3 'Nueva Operacion
                GLOBALES.gCedulaActual = frmConsultaActiva.txtCedula.Text
                Call sbFormsCall("frmCxC_Cuentas")
                For Each frm In Forms
                  If UCase(frm.Name) = UCase("frmCxC_Cuentas") Then
                    Call frm.sbGXSegTraIniTlb
                    Exit For
                  End If
                Next frm
      
      Case 4 'Tramites
            If vOperacion = 0 Then Exit Sub
                
                Call sbFormsCall("frmCxC_Cuentas")
                For Each frm In Forms
                  If UCase(frm.Name) = UCase("frmCxC_Cuentas") Then
                    Call frm.sbConsultaExterna(vOperacion)
                    Exit For
                  End If
                Next frm
     
      Case 5 'Sep
     
     
      Case 6 'Movimientos de la Operacion
            
            If vOperacion = 0 Then Exit Sub
            
            Me.MousePointer = vbHourglass
            
            With frmContenedor.Crt
                .Reset
                .WindowShowPrintSetupBtn = True
                .WindowShowRefreshBtn = True
                .WindowShowSearchBtn = True
                .WindowState = crptMaximized
                .WindowTitle = "Reportes del Módulo de CxC"
                
                .Connect = glogon.ConectRPT
                
                     .ReportFileName = SIFGlobal.fxPathReportes("CxC_PlanPagosMov.rpt")
                    
                     .Formulas(0) = "fxFecha='FECHA: " & Format(fxFechaServidor, "dd/mm/yyyy  hh:mm:ss") & "'"
                     .Formulas(1) = "fxEmpresa='" & GLOBALES.gstrNombreEmpresa & "'"
                     .Formulas(2) = "fxUsuario='USER: " & glogon.Usuario & "'"
                     .Formulas(3) = "fxOficina='" & GLOBALES.gOficina & "'"
                     
                     .SelectionFormula = "{CXC_CUENTAS.OPERACION} = " & vOperacion
                     
'                     .SubreportToChange = "sbCorte"
'                     .StoredProcParam(0) = vOperacion
'                     .StoredProcParam(1) = Format(frmConsultaActiva.dtpCorte.Value, "yyyy/mm/dd")
                .PrintReport

               
            End With
            Me.MousePointer = vbDefault
            
      Case 7 'Plan de Pagos
        If vOperacion = 0 Then Exit Sub
        
        Operacion.OperacionConsulta = vOperacion
        Call sbFormsCall("frmCxC_PlanPagos", , , , False, Me, True)
      
      Case 5 'Sep
      
      Case 8 'Sep
      
      Case 9 'Cerrar
      
    End Select

End With

Exit Sub
        
vError:
        Me.MousePointer = vbDefault
        MsgBox fxSys_Error_Handler(Err.Description)


End Sub




Private Sub mnuAyudaAcercaDe_Click()
 frmAcercaDe.Show vbModal
End Sub

Private Sub mnuAyudaContenido_Click()
   frmContenedor.CD.HelpCommand = cdlHelpContents
   frmContenedor.CD.ShowHelp
   frmContenedor.CD.HelpCommand = cdlHelpContext
End Sub



Private Sub mnuDashboard_Asociados_Click()
Call sbFormsCall("frmDSB_Asociados")
End Sub

Private Sub mnuDashboard_Click()
Call sbDashboard_Load
End Sub

Private Sub mnuMarcaOpcion_Click(Index As Integer)
Dim clsMarcas As clsMarcas

Set clsMarcas = New clsMarcas
 
Select Case Index
 Case 0 'Marca
        Call clsMarcas.Menu(glogon.Usuario, glogon.Conection, glogon.BaseDatos, glogon.Servidor _
            , App.Path, glogon.ConectRPT, 1, glogon.AppName, glogon.AppVersion, glogon.Maquina _
                    , glogon.Portal_Con, glogon.Portal_User, glogon.Portal_Key, gPortal.Empresa_Id, gPortal.Empresa_Name)
 
 Case 2 'Bitácora
        Call clsMarcas.Menu(glogon.Usuario, glogon.Conection, glogon.BaseDatos, glogon.Servidor _
            , App.Path, glogon.ConectRPT, 4, glogon.AppName, glogon.AppVersion, glogon.Maquina _
                    , glogon.Portal_Con, glogon.Portal_User, glogon.Portal_Key, gPortal.Empresa_Id, gPortal.Empresa_Name)
 Case 4 'Configuracion
        Call clsMarcas.Menu(glogon.Usuario, glogon.Conection, glogon.BaseDatos, glogon.Servidor _
            , App.Path, glogon.ConectRPT, 3, glogon.AppName, glogon.AppVersion, glogon.Maquina _
                    , glogon.Portal_Con, glogon.Portal_User, glogon.Portal_Key, gPortal.Empresa_Id, gPortal.Empresa_Name)

 Case 5 'AsgUsuarios
        Call clsMarcas.Menu(glogon.Usuario, glogon.Conection, glogon.BaseDatos, glogon.Servidor _
            , App.Path, glogon.ConectRPT, 2, glogon.AppName, glogon.AppVersion, glogon.Maquina _
                    , glogon.Portal_Con, glogon.Portal_User, glogon.Portal_Key, gPortal.Empresa_Id, gPortal.Empresa_Name)

End Select

Set clsMarcas = Nothing

End Sub

Private Sub mnuParametrosSistema_Click(Index As Integer)
Dim Nucleo As clsNucleo
  
Set Nucleo = New clsNucleo
  
Select Case Index
  Case 0 'Cofig. Empresa
        Call Nucleo.Menu(glogon.Usuario, glogon.Conection, glogon.BaseDatos, glogon.Servidor _
                    , App.Path, glogon.ConectRPT, 2, glogon.AppName, glogon.AppVersion, glogon.Maquina _
                    , glogon.Portal_Con, glogon.Portal_User, glogon.Portal_Key, gPortal.Empresa_Id, gPortal.Empresa_Name)
  
  Case 1 'Comunicados de Servicio
        Call Nucleo.Menu(glogon.Usuario, glogon.Conection, glogon.BaseDatos, glogon.Servidor _
                    , App.Path, glogon.ConectRPT, 3, glogon.AppName, glogon.AppVersion, glogon.Maquina _
                    , glogon.Portal_Con, glogon.Portal_User, glogon.Portal_Key, gPortal.Empresa_Id, gPortal.Empresa_Name)
  
  Case 2 'Separador
  
  Case 3 'Encabezados y Pie de Pagina EC
        Call Nucleo.Menu(glogon.Usuario, glogon.Conection, glogon.BaseDatos, glogon.Servidor _
                    , App.Path, glogon.ConectRPT, 4, glogon.AppName, glogon.AppVersion, glogon.Maquina _
                    , glogon.Portal_Con, glogon.Portal_User, glogon.Portal_Key, gPortal.Empresa_Id, gPortal.Empresa_Name)
  
  Case 4 'Separador
  
  Case 5 'Consulta Cola de Asientos
        Call Nucleo.Menu(glogon.Usuario, glogon.Conection, glogon.BaseDatos, glogon.Servidor _
                    , App.Path, glogon.ConectRPT, 6, glogon.AppName, glogon.AppVersion, glogon.Maquina _
                    , glogon.Portal_Con, glogon.Portal_User, glogon.Portal_Key, gPortal.Empresa_Id, gPortal.Empresa_Name)
  
  Case 6 'Separador
  
  Case 7 'Oficinas
        Call Nucleo.Menu(glogon.Usuario, glogon.Conection, glogon.BaseDatos, glogon.Servidor _
                    , App.Path, glogon.ConectRPT, 7, glogon.AppName, glogon.AppVersion, glogon.Maquina _
                    , glogon.Portal_Con, glogon.Portal_User, glogon.Portal_Key, gPortal.Empresa_Id, gPortal.Empresa_Name)
  Case 8 'Oficinas Metas
        Call Nucleo.Menu(glogon.Usuario, glogon.Conection, glogon.BaseDatos, glogon.Servidor _
                    , App.Path, glogon.ConectRPT, 8, glogon.AppName, glogon.AppVersion, glogon.Maquina _
                    , glogon.Portal_Con, glogon.Portal_User, glogon.Portal_Key, gPortal.Empresa_Id, gPortal.Empresa_Name)
  Case 9 'Separador
  
  Case 10 'Varibales Globales
        Call Nucleo.Menu(glogon.Usuario, glogon.Conection, glogon.BaseDatos, glogon.Servidor _
                    , App.Path, glogon.ConectRPT, 1, glogon.AppName, glogon.AppVersion, glogon.Maquina _
                    , glogon.Portal_Con, glogon.Portal_User, glogon.Portal_Key, gPortal.Empresa_Id, gPortal.Empresa_Name)

End Select

Set Nucleo = Nothing

End Sub

Private Sub mnuPE_Modo_1_Click()
Call sbFormsCall("frmCC_PE_PlanillaDirecta", , , , False)
End Sub

Private Sub mnuSalir_Click()
On Error Resume Next
 
  'ODBC: [CORE] Crystal Reports
 glogon.DSN = "PGX_Core"
 Call sbLogonDSN(glogon.DSN, True, 0)
   
 'ODBC: [PORTAL] Crystal Reports
 glogon.DSN = "PGX_Portal"
 Call sbLogonDSN(glogon.DSN, True, 1)
 
 'ODBC: [AUXILIARES] Crystal Reports
 glogon.DSN = "PGX_Auxiliar"
 Call sbLogonDSN(glogon.DSN, True, 2)
 
 'ODBC: [AUXILIARES] Crystal Reports
 glogon.DSN = "PGX_Analisis"
 Call sbLogonDSN(glogon.DSN, True, 3)
 
 
 Call sbSEGCuentaLog("11")
' glogon.Conection.Close
 End
End Sub


Private Sub MDIForm_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Shift = 1 And Button = 2 Then MsgBox App.FileDescription & vbCrLf & App.LegalCopyright
End Sub

Private Sub mnuSeguridadSub_Click(Index As Integer)
Dim frmX As Form, pUsuario As String


Select Case Index

 Case 0 'Cambia Contraseña
         frmCambiaClave.Show vbModal
 
 Case 1 'Actualiza Datos de Contacto de Usuario
         frmLogon_Datos_Update.Show vbModal
 
 Case 2 'Sep
 
 Case 3 'Cambiar de Tema
         frmLogon_Theme.Show vbModal
 
 
 Case 4 'Sep
 
 Case 5 'Bitacora
        Dim Nucleo As clsNucleo
        Set Nucleo = New clsNucleo
        
        Call Nucleo.Menu(glogon.Usuario, glogon.Conection, glogon.BaseDatos, glogon.Servidor _
                    , App.Path, glogon.ConectRPT, 5, glogon.AppName, glogon.AppVersion, glogon.Maquina _
                    , glogon.Portal_Con, glogon.Portal_User, glogon.Portal_Key, gPortal.Empresa_Id, gPortal.Empresa_Name)
        
        Set Nucleo = Nothing

End Select

End Sub


Private Sub mnuVerSub_Click(Index As Integer)
Dim frmX As Form

Select Case Index
  Case 0 'Ordenar x Iconos
     Me.Arrange vbArrangeIcons
  Case 1 'Ordenar en Cascada
     Me.Arrange vbCascade
  Case 2 'Ordenar x Titulo Vertical
     Me.Arrange vbTileVertical
  Case 3 'Ordenar x Titulo Horizonal
     Me.Arrange vbTileHorizontal
  
  Case 4 'Separador
  
  Case 5 'Cerrar todas las ventanas
     For Each frmX In Forms
      If Not (frmX Is Me) Then
         Unload frmX
      End If
     Next frmX
   
  Case 6 'Minimizar todas las ventanas
     For Each frmX In Forms
      If Not (frmX Is Me) Then
         frmX.WindowState = vbMinimized
      End If
     Next frmX
   
  Case 7 'Restaurar todas las ventanas
     For Each frmX In Forms
      If Not (frmX Is Me) Then
         frmX.WindowState = vbNormal
      End If
     Next frmX
   

End Select
End Sub

Private Sub sbLogonReconexion()
'Verifica que la conexión se encuentre activa

'Dim i As Integer, vMenu As String
'
'On Error GoTo vError
'
'If glogon.Reconexion = 5 Then Exit Sub
'
'vMenu = Me.Caption
'
'glogon.Conection.CommandTimeout = 10
'
'If glogon.Conection.State = 1 Then glogon.Conection.Close
'
'Me.Caption = "Conexión Caída..Reintentando Conectar al Servidor..(" & glogon.Reconexion & ")"
'glogon.Conection.Open
'
'glogon.Conection.CommandTimeout = 360
'glogon.Reconexion = 1
'Me.Caption = vMenu
'MsgBox "Conexión Reestablecida!", vbInformation
'
'Exit Sub
'
'vError:
' If glogon.Reconexion < 5 And glogon.Conection.State = 0 Then
'       glogon.Reconexion = glogon.Reconexion + 1
'       Me.Caption = vMenu
'       MsgBox "No fue posible la conexión con el servidor...intente nuevamenta la reconexión", vbCritical
' End If
'
End Sub


Public Sub sbFavoritos_Load()
Dim strSQL As String, rs As New ADODB.Recordset
Dim i As Integer, x As Integer



On Error GoTo vError

Screen.MousePointer = vbHourglass

 i = 0

With btnFavoritos
  
 For i = 0 To .Count - 1
    .Item(i).Visible = False
 Next i
  
  i = 0
  strSQL = "exec spSEG_MenuFavoritos " & gPortal.Empresa_Id & ",'" & glogon.Usuario & "'"
  Call OpenRecordSet(rs, strSQL, 1)
  Do While Not rs.EOF
'     Set btnX = .Add(, ("0x0" & rs!menu_nodo), "", tbrDefault, Trim(rs!Icono))
'         btnX.Tag = rs!menu_nodo
'         btnX.ToolTipText = Trim(rs!NODO_DESCRIPCION)
   
    If i <= .Count - 1 Then
        .Item(i).Visible = True
        .Item(i).ToolTipText = Trim(rs!NODO_DESCRIPCION)
        .Item(i).Tag = rs!menu_nodo
        
        For x = 1 To imgMenuLista.ListImages.Count
            If imgMenuLista.ListImages.Item(x).Key = Trim(rs!Icono) Then
                Set .Item(i).Picture = imgMenuLista.ListImages.Item(x).Picture
                    
                Exit For
            End If
        Next x
    End If
    i = i + 1
   rs.MoveNext
  Loop
  rs.Close
End With

Screen.MousePointer = vbDefault

Exit Sub

vError:
  Screen.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical


End Sub


Private Sub Timer_Load_Timer()
Dim strSQL As String, rs As New ADODB.Recordset
Dim pLoad As Boolean

On Error GoTo vErrorTimerLoad

pLoad = False

If Timer_Load.Interval <> 60000 Then
   pLoad = True
End If

Timer_Load.Interval = 60000

'Barra de Favoritos
If pLoad Then
        Call sbFavoritos_Load
End If

btnEmpresa.Caption = gPortal.Empresa_Name


'Revisa si el usuario cumple con requisitos de accesos del Cliente
strSQL = "exec spSEG_Access_Limit " & gPortal.Empresa_Id & ",'" & glogon.Usuario & "','" & glogon.Maquina & "',''"
Call OpenRecordSet(rs, strSQL, 1)

If rs!Indicador = 0 Then '(1 Pasa, 0 No Pasa)
       
   Select Case rs!Indicador
     Case -1
            MsgBox "Su estación de trabajo ha sido desvinculada!", vbExclamation
     Case -2
            MsgBox "Lo sentimos su sesión de Trabajo a Expirado!", vbExclamation
   End Select
   
   'Termina Sesión
   Call mnuSalir_Click

End If
rs.Close




strSQL = "select Fecha_Congela,isnull(fecha_Congela,getdate()) as 'Fecha_Auxiliar', Getdate() as 'Fecha_Actual'" _
        & " from SIF_Empresa"
Call OpenRecordSet(rs, strSQL)

StatusBar.Panels.Item(4).Text = glogon.Usuario
StatusBar.Panels.Item(5).Text = "Fecha Auxiliar: " & Format(rs!Fecha_Auxiliar, "dd/mm/yyyy")
StatusBar.Panels.Item(6).Text = "Fecha Actual: " & Format(rs!Fecha_Actual, "dd/mm/yyyy")

If IsNull(rs!Fecha_Congela) Then
    btnBloqueo.Visible = False
    
Else
    btnBloqueo.Visible = True
    btnBloqueo.Caption = "Fecha Bloqueada: " & Format(rs!Fecha_Auxiliar, "dd/mm/yyyy")
End If

rs.Close


If mLoad_Inicial Then
    strSQL = "exec  spSEG_Logon_Info '" & glogon.Usuario & "','" & glogon.Maquina_MAC & "'"
    Call OpenRecordSet(rs, strSQL, 1)
        
     If Len(Trim(rs!Tel_Cell & "")) = 0 Or Len(Trim(rs!Email & "")) = 0 Then
         frmLogon_Datos_Update.Show vbModal
     End If
    rs.Close
    
    Call sbFormsCall("frmMenu", 0, 1, 1)
'    Call sbDashboard_Load
End If
 
mLoad_Inicial = False

Exit Sub

vErrorTimerLoad:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub TimerSalir_Timer()
TimerSalir.Interval = 0
Call mnuSalir_Click
End Sub


Private Sub txtAnio_Change()
 Call sbCntX_Periodo_Refresh
End Sub

Private Sub txtEmpresa_Change()
 Call sbCntX_Periodo_Refresh
End Sub

Private Sub txtAnio_KeyUp(KeyCode As Integer, Shift As Integer)
On Error GoTo vError
    gCntX_Parametros.PeriodoAnio = txtAnio.Text
vError:
End Sub

Private Sub txtMes_Change()
 Call sbCntX_Periodo_Refresh
End Sub

Private Sub txtMes_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then txtAnio.SetFocus
End Sub

Private Sub txtMes_KeyUp(KeyCode As Integer, Shift As Integer)
On Error GoTo vError
    gCntX_Parametros.PeriodoMes = txtMes.Text
vError:
End Sub

Private Sub sbCntX_Periodo_Refresh()
Dim strSQL As String, rs As New ADODB.Recordset
Dim strResultado As String

On Error GoTo vError

txtAnio = Val(txtAnio)
  
  
gActivos.Anio = txtAnio.Text
gActivos.Mes = txtMes.Text

If txtMes.Text < 12 Then
    gActivos.Periodo = CDate(txtAnio.Text & "/" & Format(txtMes.Text, "00") & "/01")
    gActivos.Periodo = DateAdd("d", -1, DateAdd("m", 1, gActivos.Periodo))
End If
  
btnContabilidad(0).Caption = fxCntX_PeriodoDesc(txtAnio, txtMes)


strSQL = "select estado from CntX_Periodos where cod_contabilidad = " & gCntX_Parametros.CodigoConta _
       & " and anio = " & txtAnio & " and mes = " & txtMes
Call OpenRecordSet(rs, strSQL, 0)
If rs.EOF And rs.BOF Then
 
 If txtMes.Text = "13" Then
    btnContabilidad(0).ToolTipText = "Cierre Fiscal"
    btnContabilidad(0).ForeColor = vbRed
 Else
    btnContabilidad(0).ToolTipText = "Periodo No Definido"
    btnContabilidad(0).ForeColor = vbRed
 End If

Else
  If rs!Estado = "P" Then
    btnContabilidad(0).ToolTipText = "Periodo Pendiente"
    btnContabilidad(0).ForeColor = vbGrayText
  Else
    btnContabilidad(0).ToolTipText = "Periodo Cerrado"
    btnContabilidad(0).ForeColor = vbBlack
  End If
End If
rs.Close
  
Exit Sub

vError:

End Sub

