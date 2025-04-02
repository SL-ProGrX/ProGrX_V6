VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.5#0"; "comctl32.Ocx"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "Codejock.Controls.v22.1.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "Codejock.ShortcutBar.v22.1.0.ocx"
Begin VB.Form frmPreaConsultas 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Estudio Crediticio: Consultas"
   ClientHeight    =   9270
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   12165
   LinkTopic       =   "Form1"
   ScaleHeight     =   9270
   ScaleWidth      =   12165
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer TimerX 
      Interval        =   10
      Left            =   2880
      Top             =   120
   End
   Begin XtremeSuiteControls.GroupBox gbResumen 
      Height          =   2655
      Left            =   3120
      TabIndex        =   16
      Top             =   5400
      Width           =   8895
      _Version        =   1441793
      _ExtentX        =   15690
      _ExtentY        =   4683
      _StockProps     =   79
      Caption         =   "Resumen"
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
      Appearance      =   16
      BorderStyle     =   1
      Begin XtremeSuiteControls.ListView lswRsm 
         Height          =   2055
         Left            =   2160
         TabIndex        =   22
         Top             =   480
         Width           =   6495
         _Version        =   1441793
         _ExtentX        =   11451
         _ExtentY        =   3619
         _StockProps     =   77
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
         MultiSelect     =   -1  'True
         View            =   3
         FullRowSelect   =   -1  'True
         BackColor       =   16777215
         Appearance      =   16
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.RadioButton rbResumen 
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   17
         Top             =   480
         Width           =   1695
         _Version        =   1441793
         _ExtentX        =   2984
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Línea de Crédito"
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
         Value           =   -1  'True
      End
      Begin XtremeSuiteControls.RadioButton rbResumen 
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   18
         Top             =   840
         Width           =   1695
         _Version        =   1441793
         _ExtentX        =   2984
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Destino"
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
      End
      Begin XtremeSuiteControls.RadioButton rbResumen 
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   19
         Top             =   1200
         Width           =   1695
         _Version        =   1441793
         _ExtentX        =   2984
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Garantía"
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
      End
      Begin XtremeSuiteControls.RadioButton rbResumen 
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   20
         Top             =   1560
         Width           =   1695
         _Version        =   1441793
         _ExtentX        =   2984
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Institución"
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
      End
      Begin XtremeSuiteControls.RadioButton rbResumen 
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   21
         Top             =   1920
         Width           =   1695
         _Version        =   1441793
         _ExtentX        =   2984
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Estado"
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
      End
      Begin XtremeSuiteControls.RadioButton rbResumen 
         Height          =   255
         Index           =   5
         Left            =   240
         TabIndex        =   44
         Top             =   2280
         Width           =   1695
         _Version        =   1441793
         _ExtentX        =   2984
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Tendencia"
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
      End
      Begin XtremeSuiteControls.PushButton btnCopy 
         Height          =   255
         Left            =   3000
         TabIndex        =   47
         Top             =   60
         Width           =   255
         _Version        =   1441793
         _ExtentX        =   444
         _ExtentY        =   444
         _StockProps     =   79
         BackColor       =   -2147483633
         Appearance      =   16
         Picture         =   "frmPreaConsultas.frx":0000
      End
      Begin XtremeShortcutBar.ShortcutCaption scMain 
         Height          =   375
         Left            =   0
         TabIndex        =   48
         Top             =   0
         Width           =   6015
         _Version        =   1441793
         _ExtentX        =   10610
         _ExtentY        =   661
         _StockProps     =   14
         Caption         =   "Resumen                                Exportar"
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
         SubItemCaption  =   -1  'True
         ForeColor       =   4210752
      End
   End
   Begin ComctlLib.StatusBar StatusBarX 
      Align           =   2  'Align Bottom
      Height          =   240
      Left            =   0
      TabIndex        =   15
      Top             =   9024
      Width           =   12168
      _ExtentX        =   21458
      _ExtentY        =   423
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   5
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Bevel           =   0
            Object.Width           =   6068
            MinWidth        =   6068
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Bevel           =   0
            Object.Width           =   6068
            MinWidth        =   6068
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   1
            Bevel           =   0
            Enabled         =   0   'False
            TextSave        =   "MAYÚS"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel4 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   2
            Bevel           =   0
            TextSave        =   "NÚM"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel5 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   6
            Bevel           =   0
            TextSave        =   "30/6/2023"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin XtremeSuiteControls.GroupBox gBox_Procesando 
      Height          =   1335
      Left            =   4200
      TabIndex        =   12
      Top             =   1920
      Visible         =   0   'False
      Width           =   5895
      _Version        =   1441793
      _ExtentX        =   10393
      _ExtentY        =   2350
      _StockProps     =   79
      Caption         =   "Procesando: "
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
      Transparent     =   -1  'True
      UseVisualStyle  =   -1  'True
      BorderStyle     =   1
      Begin XtremeSuiteControls.ProgressBar ProgressBar1 
         Height          =   252
         Left            =   960
         TabIndex        =   13
         Top             =   840
         Width           =   3972
         _Version        =   1441793
         _ExtentX        =   7011
         _ExtentY        =   450
         _StockProps     =   93
         BackColor       =   -2147483633
         Scrolling       =   2
         MarqueeDelay    =   60
      End
      Begin VB.Label lblProcesando 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Bajando Información..."
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
         Left            =   -3000
         TabIndex        =   14
         Top             =   1680
         Width           =   3972
      End
   End
   Begin XtremeSuiteControls.GroupBox GroupBox1 
      Height          =   5412
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   2772
      _Version        =   1441793
      _ExtentX        =   4890
      _ExtentY        =   9546
      _StockProps     =   79
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
      Appearance      =   16
      BorderStyle     =   1
      Begin XtremeSuiteControls.FlatEdit FlatEdit_Linea 
         Height          =   336
         Left            =   1320
         TabIndex        =   1
         Top             =   360
         Width           =   1212
         _Version        =   1441793
         _ExtentX        =   2138
         _ExtentY        =   582
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
      Begin XtremeSuiteControls.FlatEdit FlatEdit_Destino 
         Height          =   336
         Left            =   1320
         TabIndex        =   2
         Top             =   720
         Width           =   1212
         _Version        =   1441793
         _ExtentX        =   2138
         _ExtentY        =   582
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
      Begin XtremeSuiteControls.FlatEdit FlatEdit_Institucion 
         Height          =   336
         Left            =   1320
         TabIndex        =   3
         Top             =   1080
         Width           =   1212
         _Version        =   1441793
         _ExtentX        =   2138
         _ExtentY        =   582
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
      Begin XtremeSuiteControls.FlatEdit FlatEdit_Comite 
         Height          =   336
         Left            =   1320
         TabIndex        =   4
         Top             =   1440
         Width           =   1212
         _Version        =   1441793
         _ExtentX        =   2138
         _ExtentY        =   582
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
      Begin XtremeSuiteControls.FlatEdit FlatEdit_Usuario 
         Height          =   336
         Left            =   1320
         TabIndex        =   5
         Top             =   1800
         Width           =   1212
         _Version        =   1441793
         _ExtentX        =   2138
         _ExtentY        =   582
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
      Begin XtremeSuiteControls.ComboBox cboGarantia 
         Height          =   312
         Left            =   480
         TabIndex        =   28
         Top             =   2640
         Width           =   2052
         _Version        =   1441793
         _ExtentX        =   3625
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
      Begin XtremeSuiteControls.ComboBox cboMorosidad 
         Height          =   312
         Left            =   480
         TabIndex        =   29
         Top             =   3240
         Width           =   2052
         _Version        =   1441793
         _ExtentX        =   3625
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
      Begin XtremeSuiteControls.ComboBox cboCapacidad 
         Height          =   312
         Left            =   480
         TabIndex        =   30
         Top             =   3840
         Width           =   2052
         _Version        =   1441793
         _ExtentX        =   3625
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
      Begin XtremeSuiteControls.ComboBox cboEndeudamiento 
         Height          =   312
         Left            =   480
         TabIndex        =   31
         Top             =   4440
         Width           =   2052
         _Version        =   1441793
         _ExtentX        =   3625
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
      Begin XtremeSuiteControls.ComboBox cboHistorial 
         Height          =   312
         Left            =   480
         TabIndex        =   32
         Top             =   5040
         Width           =   2052
         _Version        =   1441793
         _ExtentX        =   3625
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
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Historial Pago:"
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
         Index           =   14
         Left            =   240
         TabIndex        =   27
         Top             =   4800
         Width           =   1212
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Endeudamiento:"
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
         Index           =   13
         Left            =   240
         TabIndex        =   26
         Top             =   4200
         Width           =   1212
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Capacidad:"
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
         Index           =   12
         Left            =   240
         TabIndex        =   25
         Top             =   3600
         Width           =   1212
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Morosidad:"
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
         Index           =   11
         Left            =   240
         TabIndex        =   24
         Top             =   3000
         Width           =   1212
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Garantía:"
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
         Index           =   10
         Left            =   240
         TabIndex        =   23
         Top             =   2400
         Width           =   1212
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Crédito:"
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
         Left            =   240
         TabIndex        =   10
         Top             =   360
         Width           =   972
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Destino:"
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
         Index           =   1
         Left            =   240
         TabIndex        =   9
         Top             =   720
         Width           =   972
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Institución:"
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
         Index           =   2
         Left            =   240
         TabIndex        =   8
         Top             =   1080
         Width           =   972
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Comité:"
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
         Index           =   4
         Left            =   240
         TabIndex        =   7
         Top             =   1440
         Width           =   972
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Usuario:"
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
         Index           =   6
         Left            =   240
         TabIndex        =   6
         ToolTipText     =   "Gustos y Preferencias"
         Top             =   1800
         Width           =   972
      End
      Begin VB.Image imgBanner 
         Height          =   9516
         Index           =   0
         Left            =   0
         Picture         =   "frmPreaConsultas.frx":08D1
         Stretch         =   -1  'True
         Top             =   0
         Width           =   3204
      End
   End
   Begin FPSpreadADO.fpSpread vGrid 
      Height          =   5172
      Left            =   3120
      TabIndex        =   11
      Top             =   120
      Width           =   11052
      _Version        =   524288
      _ExtentX        =   19495
      _ExtentY        =   9123
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
      MaxCols         =   27
      SpreadDesigner  =   "frmPreaConsultas.frx":183B
      VScrollSpecial  =   -1  'True
      VScrollSpecialType=   2
      AppearanceStyle =   1
   End
   Begin XtremeSuiteControls.GroupBox gbFechas 
      Height          =   3012
      Left            =   120
      TabIndex        =   33
      Top             =   5880
      Width           =   2772
      _Version        =   1441793
      _ExtentX        =   4890
      _ExtentY        =   5313
      _StockProps     =   79
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
      Appearance      =   16
      BorderStyle     =   1
      Begin XtremeSuiteControls.PushButton btnBuscar 
         Height          =   735
         Left            =   360
         TabIndex        =   34
         Top             =   2160
         Width           =   1095
         _Version        =   1441793
         _ExtentX        =   1926
         _ExtentY        =   1291
         _StockProps     =   79
         Caption         =   "Buscar"
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
         Picture         =   "frmPreaConsultas.frx":249B
         TextImageRelation=   1
      End
      Begin XtremeSuiteControls.PushButton btnExportar 
         Height          =   735
         Left            =   1440
         TabIndex        =   35
         Top             =   2160
         Width           =   1095
         _Version        =   1441793
         _ExtentX        =   1926
         _ExtentY        =   1291
         _StockProps     =   79
         Caption         =   "Exportar"
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
         Picture         =   "frmPreaConsultas.frx":2EB9
         TextImageRelation=   1
      End
      Begin XtremeSuiteControls.DateTimePicker dtpInicio 
         Height          =   315
         Left            =   1080
         TabIndex        =   36
         Top             =   1320
         Width           =   1455
         _Version        =   1441793
         _ExtentX        =   2566
         _ExtentY        =   556
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
      Begin XtremeSuiteControls.DateTimePicker dtpCorte 
         Height          =   315
         Left            =   1080
         TabIndex        =   37
         Top             =   1680
         Width           =   1455
         _Version        =   1441793
         _ExtentX        =   2566
         _ExtentY        =   556
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
      Begin XtremeSuiteControls.ComboBox cboFecha 
         Height          =   312
         Left            =   840
         TabIndex        =   38
         Top             =   840
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
      Begin XtremeSuiteControls.ComboBox cboEstado 
         Height          =   312
         Left            =   840
         TabIndex        =   41
         Top             =   120
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
      Begin XtremeSuiteControls.ComboBox cboTramite 
         Height          =   312
         Left            =   840
         TabIndex        =   46
         Top             =   480
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
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Tramite:"
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
         Index           =   7
         Left            =   0
         TabIndex        =   45
         Top             =   480
         Width           =   1212
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Fechas:"
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
         Index           =   5
         Left            =   0
         TabIndex        =   43
         Top             =   840
         Width           =   1212
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Estado:"
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
         Index           =   3
         Left            =   0
         TabIndex        =   42
         Top             =   120
         Width           =   1212
      End
      Begin VB.Label Label1 
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
         ForeColor       =   &H00FFFFFF&
         Height          =   435
         Index           =   9
         Left            =   240
         TabIndex        =   40
         Top             =   1320
         Width           =   735
      End
      Begin VB.Label Label1 
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
         ForeColor       =   &H00FFFFFF&
         Height          =   435
         Index           =   8
         Left            =   240
         TabIndex        =   39
         Top             =   1680
         Width           =   735
      End
      Begin VB.Image imgBanner 
         Height          =   9516
         Index           =   1
         Left            =   0
         Picture         =   "frmPreaConsultas.frx":36BE
         Stretch         =   -1  'True
         Top             =   0
         Width           =   3204
      End
   End
   Begin VB.Image imgMainBanner 
      Height          =   9396
      Left            =   -120
      Picture         =   "frmPreaConsultas.frx":4628
      Stretch         =   -1  'True
      Top             =   0
      Width           =   3204
   End
End
Attribute VB_Name = "frmPreaConsultas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vPaso As Boolean, mFecUltMovUpdate As Integer

Private Sub btnBuscar_Click()
Call sbBuscar
End Sub

Private Sub btnCopy_Click()

Dim pTrayIcon As XtremeSuiteControls.TrayIcon


On Error GoTo vError
 

Set pTrayIcon = frmContenedor.TrayIcon
 
On Error GoTo vError

Me.MousePointer = vbHourglass

    
Call Excel_Exportar_Lsw(lswRsm)

 
pTrayIcon.ShowBalloonTip 25, "ProGrX: Notificación" _
            , "Exportación a Excel concluida" _
            , xtpToolTipIconInfo


Me.MousePointer = vbDefault

Exit Sub

vError:
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub btnExportar_Click()
Call sbExportar
End Sub



Private Sub sbExportar()
Dim vHeaders As vGridHeaders

    vHeaders.Columnas = 26
    vHeaders.Headers(1) = "Expediente"
    vHeaders.Headers(2) = "Estado"
    vHeaders.Headers(3) = "Identificación"
    vHeaders.Headers(4) = "Nombre"
    vHeaders.Headers(5) = "Linea Credito"
    vHeaders.Headers(6) = "Destino"
    vHeaders.Headers(7) = "Monto"
    vHeaders.Headers(8) = "Plazo"
    vHeaders.Headers(9) = "Tasa"
    vHeaders.Headers(10) = "Cuota"
    vHeaders.Headers(11) = "Refundiciones"
    vHeaders.Headers(12) = "Cancela Externa"
    vHeaders.Headers(13) = "Monto Colocado"
    vHeaders.Headers(14) = "Institución"
    vHeaders.Headers(15) = "Departamento"
    vHeaders.Headers(16) = "Oficina"
    vHeaders.Headers(17) = "Ejecutivo"
    vHeaders.Headers(18) = "Capacidad"
    vHeaders.Headers(19) = "Endeudamiento"
    vHeaders.Headers(20) = "Historia Pago"
    vHeaders.Headers(21) = "Garantia"
    vHeaders.Headers(22) = "Morosidad"
    vHeaders.Headers(23) = "Fecha Registro"
    vHeaders.Headers(24) = "Fecha Gestión"
    vHeaders.Headers(25) = "SGT Operación"
    vHeaders.Headers(26) = "SGT Estado"
    
    
    Call sbSIFGridExportar(vGrid, vHeaders, "Estudio_Credito_Consulta")


End Sub

Private Sub sbFiltro_Aplica(ByRef pSQL As String)

If cboEstado.Text <> "Todos" Then
   pSQL = pSQL & " and Estado = '" & Mid(cboEstado.Text, 1, 1) & "'"
End If


Select Case cboFecha.Text
  Case "Registro"
       pSQL = pSQL & " and Registro_Fecha between  '" & Format(dtpInicio.Value, "yyyy/mm/dd") _
               & " 00:00:00' and '" & Format(dtpCorte.Value, "yyyy/mm/dd") & " 23:59:59'"
  Case "Gestión"
       pSQL = pSQL & " and Gestion_Fecha between  '" & Format(dtpInicio.Value, "yyyy/mm/dd") _
               & " 00:00:00' and '" & Format(dtpCorte.Value, "yyyy/mm/dd") & " 23:59:59'"
End Select

If Trim(FlatEdit_Linea.Text) <> "" Then
    pSQL = pSQL & " and cod_Linea = '" & FlatEdit_Linea.Text & "'"
End If

If Trim(FlatEdit_Destino.Text) <> "" Then
    pSQL = pSQL & " and cod_destino = '" & FlatEdit_Destino.Text & "'"
End If

If IsNumeric(FlatEdit_Institucion.Text) Then
    pSQL = pSQL & " and cod_institucion = " & FlatEdit_Institucion.Text
End If

If IsNumeric(FlatEdit_Comite.Text) Then
    pSQL = pSQL & " and id_comite = " & FlatEdit_Comite.Text
End If

If cboGarantia.Text <> "TODOS" Then
    pSQL = pSQL & " and clasifica_Garantia = '" & cboGarantia.ItemData(cboGarantia.ListIndex) & "'"
End If
If cboMorosidad.Text <> "TODOS" Then
    pSQL = pSQL & " and clasifica_Morosidad = '" & cboMorosidad.ItemData(cboMorosidad.ListIndex) & "'"
End If
If cboCapacidad.Text <> "TODOS" Then
    pSQL = pSQL & " and clasifica_Capacidad = '" & cboCapacidad.ItemData(cboCapacidad.ListIndex) & "'"
End If
If cboEndeudamiento.Text <> "TODOS" Then
    pSQL = pSQL & " and clasifica_Endeudamiento = '" & cboEndeudamiento.ItemData(cboEndeudamiento.ListIndex) & "'"
End If
If cboHistorial.Text <> "TODOS" Then
    pSQL = pSQL & " and clasifica_Historial = '" & cboHistorial.ItemData(cboHistorial.ListIndex) & "'"
End If


Select Case True
 Case cboTramite.Text = "Todos"
 Case cboTramite.Text = "SGT No Indica"
    pSQL = pSQL & " and Tramite_Estado = 'NA'"
 Case cboTramite.Text = "SGT En Proceso"
    pSQL = pSQL & " and Tramite_Estado = 'P'"
 Case cboTramite.Text = "SGT Formalizada"
    pSQL = pSQL & " and Tramite_Estado = 'F'"
 Case cboTramite.Text = "SGT Anulada"
    pSQL = pSQL & " and Tramite_Estado = 'N'"
End Select

End Sub

Private Sub sbBuscar()
Dim strSQL As String, rs As New ADODB.Recordset
Dim vCadena As String, i As Integer

On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = "select '' as 'Btn',expediente, estado_desc, cedula, nombre, Linea_Desc , Destino_Desc , monto, plazo, tasa,cuota,REFUNDICIONES , DESEMBOLSOS, MONTO_COLOCADO" _
       & ",institucion_desc, departamento_Desc, Oficina_Desc, Usuario , Clasifica_capacidad, Clasifica_endeudamiento, Clasifica_historial" _
       & ", Clasifica_garantia, Clasifica_morosidad" _
       & ", Registro_Fecha , Gestion_Fecha" _
       & ", Operacion, Tramite_Desc" _
       & " From vCrd_Estudio_Crediticio" _
       & " where Usuario like '%" & FlatEdit_Usuario.Text & "%'"
       
       
strSQL = strSQL & " and Registro_Fecha between '" & Format(dtpInicio.Value, "yyyy-mm-dd") _
        & " 00:00:00' and '" & Format(dtpCorte.Value, "yyyy-mm-dd") & " 23:59:59'"
       
If cboEstado.ItemData(cboEstado.ListIndex) <> "T" Then
    strSQL = strSQL & " and Estado = '" & cboEstado.ItemData(cboEstado.ListIndex) & "'"
End If


       
Call sbFiltro_Aplica(strSQL)

gBox_Procesando.Visible = True
lblProcesando.Caption = "Cargando..."

vPaso = True

Call sbCargaGrid(vGrid, 27, strSQL, True)

vPaso = False

StatusBarX.Panels(1).Text = "Casos: " & Format(vGrid.MaxRows, "###,###,##0")

gBox_Procesando.Visible = False

'Carga resumen
Call rbResumen_Click(0)

Me.MousePointer = vbDefault

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub



Private Sub cboEstado_Click()

cboTramite.Text = "Todos"

End Sub

Private Sub cboFecha_Click()
If vPaso Then Exit Sub

If cboFecha.Text = "Todas" Then
   dtpInicio.Enabled = False
   dtpCorte.Enabled = False
Else
   dtpInicio.Enabled = True
   dtpCorte.Enabled = True
End If

End Sub


Private Sub FlatEdit_Comite_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then
    gBusquedas.Convertir = "N"
    gBusquedas.Resultado = ""
    gBusquedas.Consulta = "Select ID_COMITE,DESCRIPCION From COMITES"
    gBusquedas.Columna = "DESCRIPCION"
    gBusquedas.Orden = "DESCRIPCION"
    gBusquedas.Filtro = ""
    frmBusquedas.Show vbModal
    FlatEdit_Comite.Text = Trim(gBusquedas.Resultado)
    FlatEdit_Comite.ToolTipText = Trim(gBusquedas.Resultado2)
End If

End Sub

Private Sub FlatEdit_Destino_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then
    gBusquedas.Convertir = "N"
    gBusquedas.Resultado = ""
    gBusquedas.Consulta = "Select COD_DESTINO,DESCRIPCION From CATALOGO_DESTINOS"
    gBusquedas.Columna = "DESCRIPCION"
    gBusquedas.Orden = "DESCRIPCION"
    gBusquedas.Filtro = ""
    frmBusquedas.Show vbModal
    FlatEdit_Destino.Text = Trim(gBusquedas.Resultado)
    FlatEdit_Destino.ToolTipText = Trim(gBusquedas.Resultado2)
End If
End Sub

Private Sub FlatEdit_Institucion_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then
    gBusquedas.Convertir = "N"
    gBusquedas.Resultado = ""
    gBusquedas.Consulta = "Select COD_INSTITUCION,DESCRIPCION From INSTITUCIONES"
    gBusquedas.Columna = "DESCRIPCION"
    gBusquedas.Orden = "DESCRIPCION"
    gBusquedas.Filtro = ""
    frmBusquedas.Show vbModal
    FlatEdit_Institucion.Text = Trim(gBusquedas.Resultado)
    FlatEdit_Institucion.ToolTipText = Trim(gBusquedas.Resultado2)
End If

End Sub

Private Sub FlatEdit_Usuario_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then
    gBusquedas.Convertir = "N"
    gBusquedas.Resultado = ""
    gBusquedas.Consulta = "Select USUARIO,DESCRIPCION From USUARIO"
    gBusquedas.Columna = "DESCRIPCION"
    gBusquedas.Orden = "DESCRIPCION"
    gBusquedas.Filtro = " AND ESTADO = 'A'"
    frmBusquedas.Show vbModal
    FlatEdit_Usuario.Text = Trim(gBusquedas.Resultado)
    FlatEdit_Usuario.ToolTipText = Trim(gBusquedas.Resultado2)
End If

End Sub

Private Sub Form_Activate()
vModulo = 3

End Sub

Private Sub Form_Load()
Dim strSQL As String

vModulo = 3
vGrid.AppearanceStyle = fxGridStyle

Call Formularios(Me)
Call RefrescaTags(Me)

End Sub



Private Sub Form_Resize()
Dim pHeight As Long, pWidth As Long

On Error Resume Next

imgMainBanner.Height = Me.Height

pHeight = 8712
pWidth = 12264


If Me.Height < pHeight Then
   Me.Height = pHeight
End If

If Me.Width < pWidth Then
   Me.Width = pWidth
End If

vGrid.Width = Me.Width - (vGrid.Left + 160)
vGrid.Height = Me.Height - (vGrid.Top + StatusBarX.Height + gbResumen.Height + 1050)

scMain.Width = vGrid.Width

gbResumen.Top = vGrid.Top + vGrid.Height + 250
gbFechas.Top = gbResumen.Top - 350

gbResumen.Width = vGrid.Width
lswRsm.Width = gbResumen.Width - (lswRsm.Left + 150)


End Sub

Private Sub FlatEdit_Linea_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then
    gBusquedas.Convertir = "N"
    gBusquedas.Resultado = ""
    gBusquedas.Consulta = "Select CODIGO,DESCRIPCION From CATALOGO"
    gBusquedas.Columna = "DESCRIPCION"
    gBusquedas.Orden = "DESCRIPCION"
    gBusquedas.Filtro = " and LINEA_INTERNA = 1 AND RETENCION = 'N' AND POLIZA = 'N'"
    frmBusquedas.Show vbModal
    FlatEdit_Linea.Text = Trim(gBusquedas.Resultado)
    FlatEdit_Linea.ToolTipText = Trim(gBusquedas.Resultado2)
End If
End Sub


Private Sub rbResumen_Click(Index As Integer)
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem
Dim curMonto As Currency, curRefundicion As Currency, curDesembolso As Currency, curColocado As Currency, lngCasos As Long

On Error GoTo vError

Me.MousePointer = vbHourglass

lswRsm.ListItems.Clear
lswRsm.ColumnHeaders.Clear
lswRsm.ColumnHeaders.Add , , "Código", 1200
lswRsm.ColumnHeaders.Add , , "Descripción", 3200
lswRsm.ColumnHeaders.Add , , "No. Casos", 1200, vbCenter
lswRsm.ColumnHeaders.Add , , "Monto", 2400, vbRightJustify
lswRsm.ColumnHeaders.Add , , "Refinanciado", 2400, vbRightJustify
lswRsm.ColumnHeaders.Add , , "Cancela Ext.", 2400, vbRightJustify
lswRsm.ColumnHeaders.Add , , "N.Colocación", 2400, vbRightJustify

strSQL = "select expediente, estado, cedula, nombre, Linea_Desc , Destino_Desc , monto, plazo, tasa,cuota,REFUNDICIONES , DESEMBOLSOS, MONTO_COLOCADO" _
       & ",institucion_desc, departamento_Desc, Oficina_Desc, Usuario , Clasifica_capacidad, Clasifica_endeudamiento, Clasifica_historial" _
       & ", Clasifica_garantia, Clasifica_morosidad" _
       & ", Registro_Fecha , Gestion_Fecha" _
       & " From vCrd_Estudio_Crediticio" _
       & " where Usuario like '%" & FlatEdit_Usuario.Text & "%'"

'Filtros por Seleccion
Select Case True
 Case rbResumen.Item(0).Value 'Linea
       strSQL = "select cod_linea as 'Codigo', linea_desc as 'Descripcion', count(*) as 'Casos'" _
            & ", sum(Monto) as 'Monto', sum(Refundiciones) as 'Refundiciones', sum(Desembolsos) as 'Desembolsos' " _
            & ", sum(Monto - Refundiciones) as 'Monto_Colocado'" _
            & " From vCrd_Estudio_Crediticio" _
            & " where Usuario like '%" & FlatEdit_Usuario.Text & "%'"
 
 Case rbResumen.Item(1).Value 'Destino
       strSQL = "select cod_Destino as 'Codigo', Destino_desc as 'Descripcion', count(*) as 'Casos'" _
            & ", sum(Monto) as 'Monto', sum(Refundiciones) as 'Refundiciones', sum(Desembolsos) as 'Desembolsos' " _
            & ", sum(Monto - Refundiciones) as 'Monto_Colocado'" _
            & " From vCrd_Estudio_Crediticio" _
            & " where Usuario like '%" & FlatEdit_Usuario.Text & "%'"
 
 Case rbResumen.Item(2).Value 'Garantia
       strSQL = "select garantia as 'Codigo', Garantia_desc as 'Descripcion', count(*) as 'Casos'" _
            & ", sum(Monto) as 'Monto', sum(Refundiciones) as 'Refundiciones', sum(Desembolsos) as 'Desembolsos' " _
            & ", sum(Monto - Refundiciones) as 'Monto_Colocado'" _
            & " From vCrd_Estudio_Crediticio" _
            & " where Usuario like '%" & FlatEdit_Usuario.Text & "%'"
 
 Case rbResumen.Item(3).Value 'Institucion
       strSQL = "select cod_institucion as 'Codigo', Institucion_desc as 'Descripcion', count(*) as 'Casos'" _
            & ", sum(Monto) as 'Monto', sum(Refundiciones) as 'Refundiciones', sum(Desembolsos) as 'Desembolsos' " _
            & ", sum(Monto - Refundiciones) as 'Monto_Colocado'" _
            & " From vCrd_Estudio_Crediticio" _
            & " where Usuario like '%" & FlatEdit_Usuario.Text & "%'"
 
 Case rbResumen.Item(4).Value 'Estado
       strSQL = "select Estado as 'Codigo', Estado_desc as 'Descripcion', count(*) as 'Casos'" _
            & ", sum(Monto) as 'Monto', sum(Refundiciones) as 'Refundiciones', sum(Desembolsos) as 'Desembolsos' " _
            & ", sum(Monto - Refundiciones) as 'Monto_Colocado'" _
            & " From vCrd_Estudio_Crediticio" _
            & " where Usuario like '%" & FlatEdit_Usuario.Text & "%'"
 
 Case rbResumen.Item(5).Value 'Tendencia
       strSQL = "select year(registro_Fecha) as 'Codigo'" _
            & ", Case month(Registro_Fecha) when 1 then 'Enero' when 2 then 'Febrero'  when 3 then 'Marzo'  when 4 then 'Abril'" _
            & " when 5 then 'Mayo' when 6 then 'Junio' when 7 then 'Julio' when 8 then 'Agosto'" _
            & " when 9 then 'Setiembre' when 10 then 'Octubre' when 11 then 'Noviembre'  when 12 then 'Diciembre' " _
            & "  end as 'Descripcion', count(*) as 'Casos'" _
            & ", sum(Monto) as 'Monto', sum(Refundiciones) as 'Refundiciones', sum(Desembolsos) as 'Desembolsos' " _
            & ", sum(Monto - Refundiciones) as 'Monto_Colocado'" _
            & " From vCrd_Estudio_Crediticio" _
            & " where Usuario like '%" & FlatEdit_Usuario.Text & "%'"
 
 
End Select

'Aplica Filtros Generales
Call sbFiltro_Aplica(strSQL)

'Agrupamiento
Select Case True
 Case rbResumen.Item(0).Value 'Linea
       strSQL = strSQL & " group by cod_linea, Linea_Desc"
 
 Case rbResumen.Item(1).Value 'Destino
       strSQL = strSQL & " group by cod_Destino, Destino_desc"

 
 Case rbResumen.Item(2).Value 'Garantia
       strSQL = strSQL & " group by garantia, Garantia_desc"
 
 Case rbResumen.Item(3).Value 'Institucion
       strSQL = strSQL & " group by cod_institucion, Institucion_desc"
 
 Case rbResumen.Item(4).Value 'Estado
       strSQL = strSQL & " group by Estado, Estado_desc"
 
 Case rbResumen.Item(5).Value 'Tendencia
       strSQL = strSQL & " group by year(registro_Fecha), month(Registro_Fecha) "

End Select


lngCasos = 0
curMonto = 0
curRefundicion = 0
curDesembolso = 0
curColocado = 0

'Carga Resultados
Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
  Set itmX = lswRsm.ListItems.Add(, , rs!Codigo)
      itmX.SubItems(1) = rs!Descripcion
      itmX.SubItems(2) = Format(rs!Casos, "###,###,##0")
      itmX.SubItems(3) = Format(rs!Monto, "Standard")
      itmX.SubItems(4) = Format(rs!REFUNDICIONES, "Standard")
      itmX.SubItems(5) = Format(rs!DESEMBOLSOS, "Standard")
      itmX.SubItems(6) = Format(rs!Monto_Colocado, "Standard")
      
  lngCasos = lngCasos + 1
  curMonto = curMonto + rs!Monto
  curRefundicion = curRefundicion + rs!REFUNDICIONES
  curDesembolso = curDesembolso + rs!DESEMBOLSOS
  curColocado = curColocado + rs!Monto_Colocado
  rs.MoveNext
Loop
rs.Close

'Totales
  Set itmX = lswRsm.ListItems.Add(, , "")
      itmX.SubItems(1) = ""
      itmX.SubItems(2) = "_____________"
      itmX.SubItems(3) = "_____________"
      itmX.SubItems(4) = "_____________"
      itmX.SubItems(5) = "_____________"
      itmX.SubItems(6) = "_____________"

  Set itmX = lswRsm.ListItems.Add(, , "")
      itmX.SubItems(1) = "TOTALES"
      itmX.SubItems(2) = Format(lngCasos, "###,###,##0")
      itmX.SubItems(3) = Format(curMonto, "Standard")
      itmX.SubItems(4) = Format(curRefundicion, "Standard")
      itmX.SubItems(5) = Format(curDesembolso, "Standard")
      itmX.SubItems(6) = Format(curColocado, "Standard")
      itmX.Bold = True
Me.MousePointer = vbDefault

Exit Sub
  
vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub TimerX_Timer()
Dim strSQL As String

TimerX.Interval = 0
TimerX.Enabled = False

vPaso = True

dtpCorte.Value = fxFechaServidor
dtpInicio.Value = DateAdd("m", -3, dtpCorte.Value)

strSQL = " select M.COD_CAPACIDAD as 'IDX', rtrim(M.COD_CAPACIDAD) + ': ' + RTRIM(R.descripcion) as 'ItmX'" _
       & "   from Crd_Clasificacion_Capacidad M inner join CRD_CLASIFICACION_RAZON R on M.COD_RAZON = R.COD_RAZON"
Call sbCbo_Llena_New(cboCapacidad, strSQL, True, False)


strSQL = "select M.COD_ENDEUDAMIENTO  as 'IDX', rtrim(M.COD_ENDEUDAMIENTO) + ': ' + RTRIM(R.descripcion) as 'ItmX'" _
       & "  from CRD_CLASIFICACION_ENDEUDAMIENTO M inner join CRD_CLASIFICACION_RAZON R on M.COD_RAZON = R.COD_RAZON"
Call sbCbo_Llena_New(cboEndeudamiento, strSQL, True, False)


strSQL = " select M.COD_GARANTIA  as 'IDX', rtrim(M.COD_GARANTIA) + ': ' + RTRIM(R.descripcion) as 'ItmX'" _
       & "  from CRD_CLASIFICACION_GARANTIA M inner join CRD_CLASIFICACION_RAZON R on M.COD_RAZON = R.COD_RAZON"
Call sbCbo_Llena_New(cboGarantia, strSQL, True, False)


strSQL = " select M.COD_HISTORIAL  as 'IDX', rtrim(M.COD_HISTORIAL) + ': ' + RTRIM(M.descripcion) as 'ItmX'" _
       & "  from CRD_CLASIFICACION_HISTORIAL M inner join CRD_CLASIFICACION_RAZON R on M.COD_RAZON = R.COD_RAZON"
Call sbCbo_Llena_New(cboHistorial, strSQL, True, False)


strSQL = "select A.cod_mora as 'IdX', case " _
       & " when A.tipo = 'A' then 'Al Día    '" _
       & " when A.tipo = 'M' then 'Mora      '" _
       & " when A.tipo = 'C' then 'Cbr.Jud   '" _
       & " when A.tipo = 'I' then 'Incobrable' end + ': ' + rtrim(B.descripcion) as 'ItmX'" _
       & " from Cbr_Clasificacion_Mora A inner join Crd_Clasificacion_Razon B on A.cod_Razon = B.Cod_Razon" _
       & " order by A.cod_mora"
Call sbCbo_Llena_New(cboMorosidad, strSQL, True, False)

cboFecha.Clear
cboFecha.AddItem "Registro"
cboFecha.AddItem "Gestión"
cboFecha.AddItem "Todas"
cboFecha.Text = "Registro"

cboTramite.Clear
cboTramite.AddItem "Todos"
cboTramite.AddItem "SGT No Indica"
cboTramite.AddItem "SGT En Proceso"
cboTramite.AddItem "SGT Formalizada"
cboTramite.AddItem "SGT Anulada"


cboEstado.Clear
cboEstado.AddItem "Todos"
cboEstado.ItemData(cboEstado.ListCount - 1) = "T"
 
cboEstado.AddItem "Recibido"
cboEstado.ItemData(cboEstado.ListCount - 1) = "R"
cboEstado.AddItem "Pendiente"
cboEstado.ItemData(cboEstado.ListCount - 1) = "P"
cboEstado.AddItem "Autorizado"
cboEstado.ItemData(cboEstado.ListCount - 1) = "A"
cboEstado.AddItem "Abandonado"
cboEstado.ItemData(cboEstado.ListCount - 1) = "B"
cboEstado.AddItem "Denegado"
cboEstado.ItemData(cboEstado.ListCount - 1) = "D"
cboEstado.Text = "Todos"

vPaso = False

Call cboFecha_Click

Call sbBuscar


End Sub

Private Sub vGrid_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
If vPaso Then Exit Sub

Dim vExpediente As String, frm As Form

vGrid.Row = Row
vGrid.Col = 2

vExpediente = vGrid.Text

If vExpediente = "" Then Exit Sub

Call sbFormsCall("frmPreaEstudio", 0, 0, 0, False, Me, False)
Call sbFormActivo("frmPreaEstudio", frm)

frm.txtExpediente.Text = vExpediente
Call frm.txtExpediente_LostFocus

End Sub
