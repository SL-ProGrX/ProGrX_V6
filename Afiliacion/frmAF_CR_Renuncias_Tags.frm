VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.controls.v22.1.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.shortcutbar.v22.1.0.ocx"
Begin VB.Form frmAF_CR_Renuncias_Tags 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Etiquetas de Renuncias"
   ClientHeight    =   9750
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   14865
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9750
   ScaleWidth      =   14865
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer TimerX 
      Interval        =   5
      Left            =   14520
      Top             =   720
   End
   Begin XtremeSuiteControls.TabControl tcMain 
      Height          =   8415
      Left            =   0
      TabIndex        =   0
      Top             =   1185
      Width           =   14895
      _Version        =   1441793
      _ExtentX        =   26273
      _ExtentY        =   14843
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
      Item(0).Caption =   "Recepción del Documento"
      Item(0).ControlCount=   11
      Item(0).Control(0)=   "Label2(0)"
      Item(0).Control(1)=   "txtRenunciaId"
      Item(0).Control(2)=   "ShortcutCaption1(3)"
      Item(0).Control(3)=   "txtFiltro(2)"
      Item(0).Control(4)=   "btnRecepcionOpciones"
      Item(0).Control(5)=   "btnRefresh(0)"
      Item(0).Control(6)=   "btnExportar(0)"
      Item(0).Control(7)=   "btnRecepcion(0)"
      Item(0).Control(8)=   "chkTodas(0)"
      Item(0).Control(9)=   "lswRecepcion"
      Item(0).Control(10)=   "btnRecepcion(1)"
      Item(1).Caption =   "Revisión Satisfactoria"
      Item(1).ControlCount=   17
      Item(1).Control(0)=   "ShortcutCaption1(0)"
      Item(1).Control(1)=   "ShortcutCaption1(1)"
      Item(1).Control(2)=   "ShortcutCaption1(2)"
      Item(1).Control(3)=   "lswRecibidas"
      Item(1).Control(4)=   "txtFiltro(0)"
      Item(1).Control(5)=   "lswPendientes"
      Item(1).Control(6)=   "txtFiltro(1)"
      Item(1).Control(7)=   "btnRevisado(0)"
      Item(1).Control(8)=   "btnRevisado(1)"
      Item(1).Control(9)=   "btnRevisado(2)"
      Item(1).Control(10)=   "btnRefresh(1)"
      Item(1).Control(11)=   "btnExportar(2)"
      Item(1).Control(12)=   "btnRefresh(2)"
      Item(1).Control(13)=   "btnExportar(1)"
      Item(1).Control(14)=   "chkTodas(1)"
      Item(1).Control(15)=   "chkTodas(2)"
      Item(1).Control(16)=   "txtPendienteNota"
      Item(2).Caption =   "Bitácoras de Etiquetas"
      Item(2).ControlCount=   13
      Item(2).Control(0)=   "Label2(1)"
      Item(2).Control(1)=   "txtbRenunciaId"
      Item(2).Control(2)=   "Label2(2)"
      Item(2).Control(3)=   "txtbEstado"
      Item(2).Control(4)=   "Label2(3)"
      Item(2).Control(5)=   "txtbCedula"
      Item(2).Control(6)=   "Label2(4)"
      Item(2).Control(7)=   "txtbNombre"
      Item(2).Control(8)=   "lswEtiquetas"
      Item(2).Control(9)=   "btnReversa"
      Item(2).Control(10)=   "btnRefresh(3)"
      Item(2).Control(11)=   "btnExportar(3)"
      Item(2).Control(12)=   "txtReversaNota"
      Begin XtremeSuiteControls.ListView lswRecepcion 
         Height          =   6375
         Left            =   3720
         TabIndex        =   29
         Top             =   1230
         Width           =   11055
         _Version        =   1441793
         _ExtentX        =   19500
         _ExtentY        =   11245
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
         View            =   3
         FullRowSelect   =   -1  'True
      End
      Begin XtremeSuiteControls.ListView lswEtiquetas 
         Height          =   6375
         Left            =   -69880
         TabIndex        =   12
         Top             =   1320
         Visible         =   0   'False
         Width           =   14655
         _Version        =   1441793
         _ExtentX        =   25850
         _ExtentY        =   11245
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
         View            =   3
         FullRowSelect   =   -1  'True
      End
      Begin XtremeSuiteControls.ListView lswRecibidas 
         Height          =   6375
         Left            =   -69880
         TabIndex        =   17
         Top             =   1235
         Visible         =   0   'False
         Width           =   7215
         _Version        =   1441793
         _ExtentX        =   12726
         _ExtentY        =   11245
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
         View            =   3
         FullRowSelect   =   -1  'True
      End
      Begin XtremeSuiteControls.ListView lswPendientes 
         Height          =   6375
         Left            =   -62440
         TabIndex        =   19
         Top             =   1235
         Visible         =   0   'False
         Width           =   7215
         _Version        =   1441793
         _ExtentX        =   12726
         _ExtentY        =   11245
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
         View            =   3
         FullRowSelect   =   -1  'True
      End
      Begin XtremeSuiteControls.CheckBox chkTodas 
         Height          =   210
         Index           =   0
         Left            =   3840
         TabIndex        =   36
         Top             =   580
         Width           =   210
         _Version        =   1441793
         _ExtentX        =   370
         _ExtentY        =   370
         _StockProps     =   79
         UseVisualStyle  =   -1  'True
         Appearance      =   17
      End
      Begin XtremeSuiteControls.PushButton btnRecepcion 
         Height          =   375
         Index           =   0
         Left            =   2760
         TabIndex        =   35
         Top             =   840
         Width           =   495
         _Version        =   1441793
         _ExtentX        =   873
         _ExtentY        =   661
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
         Picture         =   "frmAF_CR_Renuncias_Tags.frx":0000
      End
      Begin XtremeSuiteControls.PushButton btnRefresh 
         Height          =   360
         Index           =   1
         Left            =   -63640
         TabIndex        =   24
         Top             =   480
         Visible         =   0   'False
         Width           =   495
         _Version        =   1441793
         _ExtentX        =   873
         _ExtentY        =   635
         _StockProps     =   79
         Appearance      =   6
         Picture         =   "frmAF_CR_Renuncias_Tags.frx":0720
      End
      Begin XtremeSuiteControls.PushButton btnReversa 
         Height          =   495
         Left            =   -57520
         TabIndex        =   13
         Top             =   7800
         Visible         =   0   'False
         Width           =   2295
         _Version        =   1441793
         _ExtentX        =   4048
         _ExtentY        =   873
         _StockProps     =   79
         Caption         =   "Reversar Revisión Satisfactoria"
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
         Picture         =   "frmAF_CR_Renuncias_Tags.frx":0E20
      End
      Begin XtremeSuiteControls.FlatEdit txtbCedula 
         Height          =   375
         Left            =   -65560
         TabIndex        =   9
         Top             =   840
         Visible         =   0   'False
         Width           =   2055
         _Version        =   1441793
         _ExtentX        =   3625
         _ExtentY        =   661
         _StockProps     =   77
         ForeColor       =   0
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
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtbEstado 
         Height          =   375
         Left            =   -67720
         TabIndex        =   7
         Top             =   840
         Visible         =   0   'False
         Width           =   2055
         _Version        =   1441793
         _ExtentX        =   3625
         _ExtentY        =   661
         _StockProps     =   77
         ForeColor       =   0
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
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtbRenunciaId 
         Height          =   375
         Left            =   -69880
         TabIndex        =   5
         ToolTipText     =   "Presione F4 para Consultar"
         Top             =   840
         Visible         =   0   'False
         Width           =   2055
         _Version        =   1441793
         _ExtentX        =   3625
         _ExtentY        =   661
         _StockProps     =   77
         ForeColor       =   0
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtRenunciaId 
         Height          =   375
         Left            =   240
         TabIndex        =   3
         Top             =   840
         Width           =   2415
         _Version        =   1441793
         _ExtentX        =   4260
         _ExtentY        =   661
         _StockProps     =   77
         ForeColor       =   0
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtFiltro 
         Height          =   375
         Index           =   0
         Left            =   -69880
         TabIndex        =   18
         Top             =   860
         Visible         =   0   'False
         Width           =   7215
         _Version        =   1441793
         _ExtentX        =   12726
         _ExtentY        =   661
         _StockProps     =   77
         ForeColor       =   0
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtFiltro 
         Height          =   375
         Index           =   1
         Left            =   -62440
         TabIndex        =   20
         Top             =   860
         Visible         =   0   'False
         Width           =   7215
         _Version        =   1441793
         _ExtentX        =   12726
         _ExtentY        =   661
         _StockProps     =   77
         ForeColor       =   0
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.PushButton btnRevisado 
         Height          =   495
         Index           =   0
         Left            =   -69880
         TabIndex        =   21
         Top             =   7680
         Visible         =   0   'False
         Width           =   1695
         _Version        =   1441793
         _ExtentX        =   2990
         _ExtentY        =   873
         _StockProps     =   79
         Caption         =   "Revisado"
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
         Picture         =   "frmAF_CR_Renuncias_Tags.frx":16DE
         ImageAlignment  =   4
      End
      Begin XtremeSuiteControls.PushButton btnRevisado 
         Height          =   495
         Index           =   1
         Left            =   -68200
         TabIndex        =   22
         Top             =   7680
         Visible         =   0   'False
         Width           =   1695
         _Version        =   1441793
         _ExtentX        =   2990
         _ExtentY        =   873
         _StockProps     =   79
         Caption         =   "Pendiente"
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
         Picture         =   "frmAF_CR_Renuncias_Tags.frx":1E05
         ImageAlignment  =   4
      End
      Begin XtremeSuiteControls.PushButton btnRevisado 
         Height          =   495
         Index           =   2
         Left            =   -59680
         TabIndex        =   23
         Top             =   7680
         Visible         =   0   'False
         Width           =   1695
         _Version        =   1441793
         _ExtentX        =   2990
         _ExtentY        =   873
         _StockProps     =   79
         Caption         =   "Revisado"
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
         Picture         =   "frmAF_CR_Renuncias_Tags.frx":2423
         ImageAlignment  =   4
      End
      Begin XtremeSuiteControls.PushButton btnRefresh 
         Height          =   360
         Index           =   2
         Left            =   -56200
         TabIndex        =   25
         Top             =   480
         Visible         =   0   'False
         Width           =   495
         _Version        =   1441793
         _ExtentX        =   873
         _ExtentY        =   635
         _StockProps     =   79
         Appearance      =   6
         Picture         =   "frmAF_CR_Renuncias_Tags.frx":2B4A
      End
      Begin XtremeSuiteControls.PushButton btnExportar 
         Height          =   360
         Index           =   1
         Left            =   -63160
         TabIndex        =   26
         Top             =   480
         Visible         =   0   'False
         Width           =   495
         _Version        =   1441793
         _ExtentX        =   873
         _ExtentY        =   635
         _StockProps     =   79
         Appearance      =   6
         Picture         =   "frmAF_CR_Renuncias_Tags.frx":324A
      End
      Begin XtremeSuiteControls.PushButton btnExportar 
         Height          =   360
         Index           =   2
         Left            =   -55720
         TabIndex        =   27
         Top             =   480
         Visible         =   0   'False
         Width           =   495
         _Version        =   1441793
         _ExtentX        =   873
         _ExtentY        =   635
         _StockProps     =   79
         Appearance      =   6
         Picture         =   "frmAF_CR_Renuncias_Tags.frx":33B4
      End
      Begin XtremeSuiteControls.FlatEdit txtFiltro 
         Height          =   375
         Index           =   2
         Left            =   3720
         TabIndex        =   30
         Top             =   855
         Width           =   11055
         _Version        =   1441793
         _ExtentX        =   19500
         _ExtentY        =   661
         _StockProps     =   77
         ForeColor       =   0
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.PushButton btnRecepcion 
         Height          =   495
         Index           =   1
         Left            =   7920
         TabIndex        =   31
         Top             =   7680
         Width           =   1695
         _Version        =   1441793
         _ExtentX        =   2990
         _ExtentY        =   873
         _StockProps     =   79
         Caption         =   "Recibido"
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
         Picture         =   "frmAF_CR_Renuncias_Tags.frx":351E
         ImageAlignment  =   4
      End
      Begin XtremeSuiteControls.PushButton btnRecepcionOpciones 
         Height          =   495
         Left            =   9600
         TabIndex        =   32
         Top             =   7680
         Width           =   1695
         _Version        =   1441793
         _ExtentX        =   2990
         _ExtentY        =   873
         _StockProps     =   79
         Caption         =   "Opciones >"
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
         Picture         =   "frmAF_CR_Renuncias_Tags.frx":3B42
         ImageAlignment  =   4
      End
      Begin XtremeSuiteControls.PushButton btnRefresh 
         Height          =   360
         Index           =   0
         Left            =   13800
         TabIndex        =   33
         Top             =   480
         Width           =   495
         _Version        =   1441793
         _ExtentX        =   873
         _ExtentY        =   635
         _StockProps     =   79
         Appearance      =   6
         Picture         =   "frmAF_CR_Renuncias_Tags.frx":425B
      End
      Begin XtremeSuiteControls.PushButton btnExportar 
         Height          =   360
         Index           =   0
         Left            =   14280
         TabIndex        =   34
         Top             =   480
         Width           =   495
         _Version        =   1441793
         _ExtentX        =   873
         _ExtentY        =   635
         _StockProps     =   79
         Appearance      =   6
         Picture         =   "frmAF_CR_Renuncias_Tags.frx":495B
      End
      Begin XtremeSuiteControls.CheckBox chkTodas 
         Height          =   210
         Index           =   1
         Left            =   -69760
         TabIndex        =   37
         Top             =   580
         Visible         =   0   'False
         Width           =   210
         _Version        =   1441793
         _ExtentX        =   370
         _ExtentY        =   370
         _StockProps     =   79
         UseVisualStyle  =   -1  'True
         Appearance      =   17
      End
      Begin XtremeSuiteControls.CheckBox chkTodas 
         Height          =   210
         Index           =   2
         Left            =   -62320
         TabIndex        =   38
         Top             =   580
         Visible         =   0   'False
         Width           =   210
         _Version        =   1441793
         _ExtentX        =   370
         _ExtentY        =   370
         _StockProps     =   79
         UseVisualStyle  =   -1  'True
         Appearance      =   17
      End
      Begin XtremeSuiteControls.PushButton btnRefresh 
         Height          =   360
         Index           =   3
         Left            =   -56200
         TabIndex        =   40
         Top             =   840
         Visible         =   0   'False
         Width           =   495
         _Version        =   1441793
         _ExtentX        =   873
         _ExtentY        =   635
         _StockProps     =   79
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         Picture         =   "frmAF_CR_Renuncias_Tags.frx":4AC5
      End
      Begin XtremeSuiteControls.PushButton btnExportar 
         Height          =   360
         Index           =   3
         Left            =   -55720
         TabIndex        =   41
         Top             =   840
         Visible         =   0   'False
         Width           =   495
         _Version        =   1441793
         _ExtentX        =   873
         _ExtentY        =   635
         _StockProps     =   79
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         Picture         =   "frmAF_CR_Renuncias_Tags.frx":51C5
      End
      Begin XtremeSuiteControls.FlatEdit txtPendienteNota 
         Height          =   495
         Left            =   -66400
         TabIndex        =   42
         ToolTipText     =   "Nota para Pendientes"
         Top             =   7680
         Visible         =   0   'False
         Width           =   3735
         _Version        =   1441793
         _ExtentX        =   6588
         _ExtentY        =   873
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
         MultiLine       =   -1  'True
         ScrollBars      =   2
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtbNombre 
         Height          =   375
         Left            =   -63400
         TabIndex        =   11
         Top             =   840
         Visible         =   0   'False
         Width           =   6015
         _Version        =   1441793
         _ExtentX        =   10610
         _ExtentY        =   661
         _StockProps     =   77
         ForeColor       =   0
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
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtReversaNota 
         Height          =   495
         Left            =   -69880
         TabIndex        =   43
         ToolTipText     =   "Nota para Pendientes"
         Top             =   7800
         Visible         =   0   'False
         Width           =   12015
         _Version        =   1441793
         _ExtentX        =   21193
         _ExtentY        =   873
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
         MultiLine       =   -1  'True
         ScrollBars      =   2
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
         Height          =   375
         Index           =   3
         Left            =   3720
         TabIndex        =   28
         Top             =   480
         Width           =   11055
         _Version        =   1441793
         _ExtentX        =   19500
         _ExtentY        =   661
         _StockProps     =   14
         Caption         =   "Renuncias Pendientes de Recepción"
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
         Alignment       =   1
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
         Height          =   375
         Index           =   2
         Left            =   -63520
         TabIndex        =   16
         Top             =   480
         Visible         =   0   'False
         Width           =   1935
         _Version        =   1441793
         _ExtentX        =   3413
         _ExtentY        =   661
         _StockProps     =   14
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
         Alignment       =   1
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
         Height          =   375
         Index           =   1
         Left            =   -62440
         TabIndex        =   15
         Top             =   480
         Visible         =   0   'False
         Width           =   7215
         _Version        =   1441793
         _ExtentX        =   12726
         _ExtentY        =   661
         _StockProps     =   14
         Caption         =   "Renuncias Pendientes"
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
         Alignment       =   1
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
         Height          =   375
         Index           =   0
         Left            =   -69880
         TabIndex        =   14
         Top             =   480
         Visible         =   0   'False
         Width           =   7215
         _Version        =   1441793
         _ExtentX        =   12726
         _ExtentY        =   661
         _StockProps     =   14
         Caption         =   "Renuncias Recibidas"
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
         Alignment       =   1
      End
      Begin XtremeSuiteControls.Label Label2 
         Height          =   375
         Index           =   4
         Left            =   -63400
         TabIndex        =   10
         Top             =   480
         Visible         =   0   'False
         Width           =   5895
         _Version        =   1441793
         _ExtentX        =   10398
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Nombre"
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
         Transparent     =   -1  'True
         WordWrap        =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label2 
         Height          =   375
         Index           =   3
         Left            =   -65560
         TabIndex        =   8
         Top             =   480
         Visible         =   0   'False
         Width           =   2055
         _Version        =   1441793
         _ExtentX        =   3625
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Identificación"
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
         Transparent     =   -1  'True
         WordWrap        =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label2 
         Height          =   375
         Index           =   2
         Left            =   -67600
         TabIndex        =   6
         Top             =   480
         Visible         =   0   'False
         Width           =   1815
         _Version        =   1441793
         _ExtentX        =   3201
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
         Alignment       =   2
         Transparent     =   -1  'True
         WordWrap        =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label2 
         Height          =   375
         Index           =   1
         Left            =   -69880
         TabIndex        =   4
         Top             =   480
         Visible         =   0   'False
         Width           =   2055
         _Version        =   1441793
         _ExtentX        =   3625
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Código de Renuncia"
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
         Transparent     =   -1  'True
         WordWrap        =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label2 
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   2
         Top             =   480
         Width           =   1575
         _Version        =   1441793
         _ExtentX        =   2778
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Código de Renuncia"
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
   End
   Begin XtremeSuiteControls.ProgressBar ProgressBarX 
      Height          =   135
      Left            =   0
      TabIndex        =   39
      Top             =   9600
      Visible         =   0   'False
      Width           =   14895
      _Version        =   1441793
      _ExtentX        =   26273
      _ExtentY        =   238
      _StockProps     =   93
      BackColor       =   -2147483633
      Scrolling       =   1
   End
   Begin XtremeSuiteControls.PushButton btnAdministra 
      Height          =   375
      Left            =   0
      TabIndex        =   44
      Top             =   720
      Visible         =   0   'False
      Width           =   375
      _Version        =   1441793
      _ExtentX        =   661
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Administra - NoVisible"
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
      Picture         =   "frmAF_CR_Renuncias_Tags.frx":532F
      ImageAlignment  =   4
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Revisión y Control de Renuncias"
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
      Height          =   495
      Index           =   0
      Left            =   2280
      TabIndex        =   1
      Top             =   360
      Width           =   6255
   End
   Begin VB.Image imgBanner 
      Height          =   1125
      Left            =   0
      Top             =   0
      Width           =   15135
   End
End
Attribute VB_Name = "frmAF_CR_Renuncias_Tags"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem
Dim vPaso As Boolean


Private Sub sbListas_Load(Lista As Object, Estado As String, Filtro As String)

On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = "exec spAFI_Renuncias_Control_Consulta '" & Estado & "', '" & Filtro & "'"
Call OpenRecordSet(rs, strSQL)
    
With Lista.ListItems
    .Clear
    Do While Not rs.EOF
        Set itmX = .Add(, , rs!cod_Renuncia)
            itmX.SubItems(1) = rs!Cedula
            itmX.SubItems(2) = rs!Nombre
            itmX.SubItems(3) = Format(rs!registro_Fecha, "yyyy-mm-dd")
            itmX.SubItems(4) = Format(rs!Vencimiento, "yyyy-mm-dd")
            itmX.SubItems(5) = rs!Tipo
            itmX.SubItems(6) = rs!Estado_Desc
            itmX.SubItems(7) = rs!Causa_Desc
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




Private Sub btnExportar_Click(Index As Integer)
On Error GoTo vError

Me.MousePointer = vbHourglass

ProgressBarX.Visible = True

Select Case Index
  Case 0
        Call Excel_Exportar_Lsw(lswRecepcion, ProgressBarX)
  Case 1
        Call Excel_Exportar_Lsw(lswRecibidas, ProgressBarX)
  Case 2
        Call Excel_Exportar_Lsw(lswPendientes, ProgressBarX)
  Case 3
        Call Excel_Exportar_Lsw(lswEtiquetas, ProgressBarX)
End Select

ProgressBarX.Visible = False

Me.MousePointer = vbDefault

Exit Sub

vError:
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub


Private Sub btnRecepcion_Click(Index As Integer)

Dim i As Long

On Error GoTo vError

Me.MousePointer = vbHourglass

ProgressBarX.Visible = True

Select Case Index
  Case 0 'Caso de una Renuncia
    strSQL = "exec spAFI_Renuncia_Recepcion_Aplica " & txtRenunciaId.Text & ", '" & glogon.Usuario _
           & "', 'Recibe Renuncia No. " & txtRenunciaId.Text & "', '" & glogon.Maquina & "', '" & glogon.AppVersion & "'"
    Call ConectionExecute(strSQL)
      
    txtRenunciaId.Text = ""
      
  Case 1 'Listado
    strSQL = ""
    With lswRecepcion.ListItems
    For i = 1 To .Count
        If .Item(i).Checked Then
            strSQL = strSQL & Space(10) & "exec spAFI_Renuncia_Recepcion_Aplica " & .Item(i).Text & ", '" & glogon.Usuario _
                   & "', 'Recibe Renuncia No. " & txtRenunciaId.Text & "', '" & glogon.Maquina & "', '" & glogon.AppVersion & "'"
        End If
        
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
    
    End With
End Select

'Actualiza Lista
Call btnRefresh_Click(0)

ProgressBarX.Visible = False
Me.MousePointer = vbDefault

Exit Sub

vError:
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub btnRefresh_Click(Index As Integer)

Select Case Index
    Case 0 'Recepcion
        Call sbListas_Load(lswRecepcion, "Recepcion", txtFiltro(2).Text)
    Case 1 'Recibidas
        Call sbListas_Load(lswRecibidas, "Recibida", txtFiltro(1).Text)
    Case 2 'Pendientes
        Call sbListas_Load(lswPendientes, "Pendiente", txtFiltro(2).Text)
    Case 3 'Consulta de Tags
        Call txtbRenunciaId_LostFocus
End Select

End Sub

Private Sub sbListas_Check(Lista As Object, Valor As Integer)
Dim i As Long

With Lista.ListItems
   For i = 1 To .Count
       .Item(i).Checked = Valor
   Next i
End With

End Sub

Private Sub btnReversa_Click()
Dim pNota As String

On Error GoTo vError

If Not IsNumeric(txtbRenunciaId.Text) Then Exit Sub

If Len(txtReversaNota.Text) < 10 Then
    MsgBox "Indique una Nota válida para la reversión!", vbExclamation
    Exit Sub
End If


strSQL = "select dbo.fxAFI_Renuncia_Revision_Reversar_Valida(" & txtbRenunciaId.Text & ") as 'Result'"
Call OpenRecordSet(rs, strSQL)
If rs!Result = 0 Then
    MsgBox "No procede la reversión, porque la Renuncia ya fue procesada (Liquidación)", vbExclamation
    Exit Sub
End If

Me.MousePointer = vbHourglass

pNota = Mid(txtReversaNota.Text, 1, 500)

strSQL = "exec spAFI_Renuncia_Revision_Reversar " & txtbRenunciaId.Text & ", '" & glogon.Usuario _
       & "', '" & pNota & "', '" & glogon.Maquina & "', '" & glogon.AppVersion & "'"
Call ConectionExecute(strSQL)

Me.MousePointer = vbDefault
MsgBox "Reversión aplicada satisfactoriamente!", vbInformation

txtReversaNota.Text = ""

'Actualiza Lista
Call txtbRenunciaId_LostFocus

Exit Sub

vError:
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub btnRevisado_Click(Index As Integer)

Dim i As Long, pEstado As String, pRefresh As Integer, pNota As String
Dim vProcesa As Boolean
On Error GoTo vError

i = MsgBox("Esta Seguro que aplicar esta etiqueta? " & btnRevisado(Index).Caption, vbYesNo)
If i = vbNo Then
        Exit Sub
End If

Me.MousePointer = vbHourglass

ProgressBarX.Visible = True

strSQL = ""
vProcesa = False

Select Case Index
  Case 0, 1 'Recibido: Revisado
    If Index = 0 Then
        pEstado = "P" 'Procesa
        pNota = ""
    Else
        pEstado = "E" 'Pendiente
        pNota = Mid(txtPendienteNota.Text, 1, 500)
        
        If Len(pNota) < 10 Then
            Me.MousePointer = vbDefault
            MsgBox "Indique una nota válida para poner el pendiente", vbExclamation
            Exit Sub
        End If
    End If
    
    With lswRecibidas.ListItems
        For i = 1 To .Count
            If .Item(i).Checked Then
                strSQL = strSQL & Space(10) & "exec spAFI_Renuncia_Revision_Aplica " & .Item(i).Text & ", '" & pEstado & "', '" & glogon.Usuario _
                       & "', '" & pNota & "', '" & glogon.Maquina & "', '" & glogon.AppVersion & "'"
            End If
            
            If Len(strSQL) > 20000 Then
                Call ConectionExecute(strSQL)
                strSQL = ""
                vProcesa = True
            End If
        Next i
    
    End With
   
  
  Case 2 'Pendiente : Revisado
    pEstado = "P"
    pNota = ""
    With lswPendientes.ListItems
        For i = 1 To .Count
            If .Item(i).Checked Then
                strSQL = strSQL & Space(10) & "exec spAFI_Renuncia_Revision_Aplica " & .Item(i).Text & ", '" & pEstado & "', '" & glogon.Usuario _
                       & "', '" & pNota & "', '" & glogon.Maquina & "', '" & glogon.AppVersion & "'"
            End If
            
            If Len(strSQL) > 20000 Then
                Call ConectionExecute(strSQL)
                strSQL = ""
                vProcesa = True
            End If
        Next i
    
    End With

End Select


'Ultimo Lote
If Len(strSQL) > 0 Then
    Call ConectionExecute(strSQL)
    strSQL = ""
    vProcesa = True
End If


'Actualiza Lista
Call btnRefresh_Click(1)
If pEstado = "E" Or Index = 2 Then
Call btnRefresh_Click(2)
End If


ProgressBarX.Visible = False
Me.MousePointer = vbDefault

If vProcesa Then
    MsgBox "Casos aplicados satisfactoriamente!", vbInformation
End If

Exit Sub

vError:
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub chkTodas_Click(Index As Integer)

Select Case Index
    Case 0 'Recepcion
        Call sbListas_Check(lswRecepcion, chkTodas(Index).Value)
    Case 1 'Recibidas
        Call sbListas_Check(lswRecibidas, chkTodas(Index).Value)
    Case 2 'Recepcion
        Call sbListas_Check(lswPendientes, chkTodas(Index).Value)
End Select

End Sub

Private Sub Form_Load()

vModulo = 1

Set imgBanner.Picture = frmContenedor.imgBanner_Mantenimiento.Picture

With lswEtiquetas.ColumnHeaders
    .Clear
    .Add , , "Etiqueta", 4000
    .Add , , "Reg. Fecha", 2100
    .Add , , "Reg. Usuario", 2100, vbCenter
    .Add , , "Observaciones", 4000
End With


lswRecepcion.Checkboxes = True
With lswRecepcion.ColumnHeaders
    .Clear
    .Add , , "Id Renuncia", 1800
    .Add , , "Cédula", 1500, vbCenter
    .Add , , "Nombre", 3150
    .Add , , "Fecha", 1800, vbCenter
    .Add , , "Vence", 1800, vbCenter
    .Add , , "Tipo", 1800, vbCenter
    .Add , , "Estado", 1800, vbCenter
    .Add , , "Causa", 3100
End With

lswRecibidas.Checkboxes = True
With lswRecibidas.ColumnHeaders
    .Clear
    .Add , , "Id Renuncia", 1800
    .Add , , "Cédula", 1500, vbCenter
    .Add , , "Nombre", 3150
    .Add , , "Fecha", 1800, vbCenter
    .Add , , "Vence", 1800, vbCenter
    .Add , , "Tipo", 1800, vbCenter
    .Add , , "Estado", 1800, vbCenter
    .Add , , "Causa", 3100
End With

lswPendientes.Checkboxes = True
With lswPendientes.ColumnHeaders
    .Clear
    .Add , , "Id Renuncia", 1800
    .Add , , "Cédula", 1500, vbCenter
    .Add , , "Nombre", 3150
    .Add , , "Fecha", 1800, vbCenter
    .Add , , "Vence", 1800, vbCenter
    .Add , , "Tipo", 1800, vbCenter
    .Add , , "Estado", 1800, vbCenter
    .Add , , "Causa", 3100
End With


Call Formularios(Me)

'Otorga Acceso a traves de btnAdministra.Tag
btnRecepcion(0).Enabled = btnAdministra.Tag
btnRecepcion(1).Enabled = btnAdministra.Tag
btnRecepcionOpciones.Enabled = btnAdministra.Tag
btnRevisado(1).Enabled = btnAdministra.Tag
btnRevisado(2).Enabled = btnAdministra.Tag
btnRevisado(0).Enabled = btnAdministra.Tag

btnAdministra.Visible = False

Call RefrescaTags(Me)

End Sub

Private Sub lswPendientes_ColumnClick(ByVal ColumnHeader As XtremeSuiteControls.ListViewColumnHeader)
 lswPendientes.SortKey = ColumnHeader.Index - 1
  If lswPendientes.SortOrder = 0 Then lswPendientes.SortOrder = 1 Else lswPendientes.SortOrder = 0
  lswPendientes.Sorted = True
End Sub

Private Sub lswRecepcion_ColumnClick(ByVal ColumnHeader As XtremeSuiteControls.ListViewColumnHeader)
 lswRecepcion.SortKey = ColumnHeader.Index - 1
  If lswRecepcion.SortOrder = 0 Then lswRecepcion.SortOrder = 1 Else lswRecepcion.SortOrder = 0
  lswRecepcion.Sorted = True
End Sub


Private Sub lswRecibidas_ColumnClick(ByVal ColumnHeader As XtremeSuiteControls.ListViewColumnHeader)
 lswRecibidas.SortKey = ColumnHeader.Index - 1
  If lswRecibidas.SortOrder = 0 Then lswRecibidas.SortOrder = 1 Else lswRecibidas.SortOrder = 0
  lswRecibidas.Sorted = True
End Sub


Private Sub tcMain_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)

Select Case Item.Index
    Case 0 'Recepcion
        Call sbListas_Load(lswRecepcion, "Recepcion", txtFiltro(2).Text)
    
    Case 1 'Recibidas y Pendientes
        Call sbListas_Load(lswRecibidas, "Recibida", txtFiltro(0).Text)
        Call sbListas_Load(lswPendientes, "Pendiente", txtFiltro(1).Text)
    
    Case 3 'Consulta de Tags

End Select

End Sub

Private Sub TimerX_Timer()

TimerX.Interval = 0
TimerX.Enabled = False

tcMain.Item(0).Selected = True
Call sbListas_Load(lswRecepcion, "Recepcion", txtFiltro(2).Text)

End Sub



Private Sub txtbRenunciaId_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then
   txtbEstado.SetFocus
End If

If KeyCode = vbKeyF4 Then
  gBusquedas.Columna = "COD_RENUNCIA"
  gBusquedas.Orden = "COD_RENUNCIA"
  gBusquedas.Consulta = "select COD_RENUNCIA, CEDULA, NOMBRE, Estado_Desc From vAFI_Renuncias"
  gBusquedas.Filtro = ""
  
  frmBusquedas.Show vbModal
  
  txtbRenunciaId.Text = gBusquedas.Resultado
End If

End Sub

Private Sub sbEtiquetas_Consulta(pRenunciaId As Long)

On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = "exec spAFI_Renuncia_Etiquetas_Consulta " & pRenunciaId
Call OpenRecordSet(rs, strSQL)

With lswEtiquetas.ListItems
    .Clear
    Do While Not rs.EOF
        Set itmX = .Add(, , rs!Tag_Desc)
            itmX.SubItems(1) = rs!Fecha_Format
            itmX.SubItems(2) = rs!Registro_Usuario & ""
            itmX.SubItems(3) = rs!Observacion & ""
            itmX.Tag = rs!Id
     rs.MoveNext
    Loop
End With

rs.Close
Me.MousePointer = vbDefault
Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub txtbRenunciaId_LostFocus()

If IsNumeric(txtbRenunciaId.Text) Then
    strSQL = "select COD_RENUNCIA, CEDULA, NOMBRE, Estado_Desc" _
           & " From vAFI_Renuncias Where cod_Renuncia = " & txtbRenunciaId.Text
    Call OpenRecordSet(rs, strSQL)
    If Not rs.EOF And Not rs.BOF Then
        txtbEstado.Text = rs!Estado_Desc
        txtbCedula.Text = rs!Cedula
        txtbNombre.Text = rs!Nombre
    End If
    rs.Close
    
    Call sbEtiquetas_Consulta(txtbRenunciaId.Text)
End If


End Sub

Private Sub txtFiltro_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyReturn Then
    Select Case Index
        Case 2 'Recepcion
            Call sbListas_Load(lswRecepcion, "Recepcion", txtFiltro(2).Text)
        
        Case 0 'Recibidas
            Call sbListas_Load(lswRecibidas, "Recibidas", txtFiltro(0).Text)
        
        Case 1 'Pendientes
            Call sbListas_Load(lswPendientes, "Pendientes", txtFiltro(1).Text)
    End Select
End If

End Sub

Private Sub txtRenunciaId_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
   Call btnRecepcion_Click(0)
End If

If KeyCode = vbKeyF4 Then
  gBusquedas.Columna = "COD_RENUNCIA"
  gBusquedas.Orden = "COD_RENUNCIA"
  gBusquedas.Consulta = "select COD_RENUNCIA, CEDULA, NOMBRE, Estado_Desc From vAFI_Renuncias_Pendientes_Recibir"
  gBusquedas.Filtro = ""
  
  frmBusquedas.Show vbModal
  
  txtRenunciaId.Text = gBusquedas.Resultado
End If

End Sub
