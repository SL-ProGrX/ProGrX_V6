VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.controls.v22.1.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.shortcutbar.v22.1.0.ocx"
Begin VB.Form frmAF_Reportes 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reportes"
   ClientHeight    =   8295
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10350
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8295
   ScaleWidth      =   10350
   Begin VB.Timer TimerX 
      Interval        =   10
      Left            =   0
      Top             =   0
   End
   Begin XtremeSuiteControls.TabControl tcPrincipal 
      Height          =   7095
      Left            =   0
      TabIndex        =   2
      Top             =   1200
      Width           =   10335
      _Version        =   1441793
      _ExtentX        =   18230
      _ExtentY        =   12515
      _StockProps     =   68
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   4
      Color           =   32
      PaintManager.Position=   2
      PaintManager.BoldSelected=   -1  'True
      ItemCount       =   3
      Item(0).Caption =   "Informes"
      Item(0).ControlCount=   19
      Item(0).Control(0)=   "ArbolExp"
      Item(0).Control(1)=   "dtpInicio"
      Item(0).Control(2)=   "dtpCorte"
      Item(0).Control(3)=   "cboTipo"
      Item(0).Control(4)=   "cboFBase"
      Item(0).Control(5)=   "cboUsuarios"
      Item(0).Control(6)=   "chkFechas"
      Item(0).Control(7)=   "Label1(14)"
      Item(0).Control(8)=   "Label1(8)"
      Item(0).Control(9)=   "Label1(7)"
      Item(0).Control(10)=   "Label1(6)"
      Item(0).Control(11)=   "Label1(5)"
      Item(0).Control(12)=   "Label1(4)"
      Item(0).Control(13)=   "Label1(3)"
      Item(0).Control(14)=   "cboEstado"
      Item(0).Control(15)=   "tcMain"
      Item(0).Control(16)=   "btnReporte"
      Item(0).Control(17)=   "imgSeguridad"
      Item(0).Control(18)=   "lblSeguridad"
      Item(1).Caption =   "Configuración"
      Item(1).ControlCount=   2
      Item(1).Control(0)=   "tcAux"
      Item(1).Control(1)=   "scTitulosTabs(0)"
      Item(2).Caption =   "Seguridad"
      Item(2).ControlCount=   2
      Item(2).Control(0)=   "scTitulosTabs(1)"
      Item(2).Control(1)=   "tcAuxGrpAccs"
      Begin XtremeSuiteControls.TabControl tcAux 
         Height          =   6255
         Left            =   -70000
         TabIndex        =   3
         Top             =   480
         Visible         =   0   'False
         Width           =   10455
         _Version        =   1441793
         _ExtentX        =   18441
         _ExtentY        =   11033
         _StockProps     =   68
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   4
         Color           =   32
         PaintManager.BoldSelected=   -1  'True
         ItemCount       =   3
         Item(0).Caption =   "Grupos"
         Item(0).ControlCount=   2
         Item(0).Control(0)=   "vGrid"
         Item(0).Control(1)=   "Label2(1)"
         Item(1).Caption =   "Miembros"
         Item(1).ControlCount=   4
         Item(1).Control(0)=   "cboMiembros"
         Item(1).Control(1)=   "lswMiembros"
         Item(1).Control(2)=   "Label2(2)"
         Item(1).Control(3)=   "Label2(3)"
         Item(2).Caption =   "Informes"
         Item(2).ControlCount=   4
         Item(2).Control(0)=   "vGridRep"
         Item(2).Control(1)=   "imgAddRep"
         Item(2).Control(2)=   "txtReportes"
         Item(2).Control(3)=   "Label4"
         Begin XtremeSuiteControls.ListView lswMiembros 
            Height          =   5415
            Left            =   -67840
            TabIndex        =   74
            Top             =   840
            Visible         =   0   'False
            Width           =   6375
            _Version        =   1441793
            _ExtentX        =   11245
            _ExtentY        =   9551
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
            Appearance      =   17
            UseVisualStyle  =   0   'False
         End
         Begin FPSpreadADO.fpSpread vGrid 
            Height          =   5775
            Left            =   2160
            TabIndex        =   4
            Top             =   480
            Width           =   6615
            _Version        =   524288
            _ExtentX        =   11668
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
            MaxCols         =   496
            ScrollBars      =   2
            SpreadDesigner  =   "frmAF_Reportes.frx":0000
            VScrollSpecialType=   2
            AppearanceStyle =   1
         End
         Begin FPSpreadADO.fpSpread vGridRep 
            Height          =   5175
            Left            =   -69520
            TabIndex        =   70
            Top             =   960
            Visible         =   0   'False
            Width           =   9135
            _Version        =   524288
            _ExtentX        =   16113
            _ExtentY        =   9128
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
            MaxCols         =   495
            ScrollBars      =   2
            SpreadDesigner  =   "frmAF_Reportes.frx":0507
            VScrollSpecialType=   2
            AppearanceStyle =   1
         End
         Begin XtremeSuiteControls.FlatEdit txtReportes 
            Height          =   375
            Left            =   -67000
            TabIndex        =   71
            Top             =   480
            Visible         =   0   'False
            Width           =   1935
            _Version        =   1441793
            _ExtentX        =   3413
            _ExtentY        =   661
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
            PasswordChar    =   "*"
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.ComboBox cboMiembros 
            Height          =   330
            Left            =   -67840
            TabIndex        =   73
            Top             =   480
            Visible         =   0   'False
            Width           =   6375
            _Version        =   1441793
            _ExtentX        =   11245
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
         Begin XtremeSuiteControls.Label Label4 
            Height          =   375
            Left            =   -69040
            TabIndex        =   72
            Top             =   480
            Visible         =   0   'False
            Width           =   1935
            _Version        =   1441793
            _ExtentX        =   3413
            _ExtentY        =   661
            _StockProps     =   79
            Caption         =   "Clave de Edición (Admin)"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Transparent     =   -1  'True
            WordWrap        =   -1  'True
         End
         Begin VB.Image imgAddRep 
            Height          =   375
            Left            =   -61120
            Picture         =   "frmAF_Reportes.frx":0B05
            Stretch         =   -1  'True
            ToolTipText     =   "Agregar & Actualizar lista de Reportes"
            Top             =   480
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            Caption         =   "Grupo de Usuarios"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   555
            Index           =   1
            Left            =   1080
            TabIndex        =   7
            Top             =   480
            Width           =   1095
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            Caption         =   "Grupo"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   315
            Index           =   2
            Left            =   -69040
            TabIndex        =   6
            Top             =   480
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            Caption         =   "Miembros"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   315
            Index           =   3
            Left            =   -69040
            TabIndex        =   5
            Top             =   840
            Visible         =   0   'False
            Width           =   1215
         End
      End
      Begin XtremeSuiteControls.TabControl tcAuxGrpAccs 
         Height          =   6255
         Left            =   -70000
         TabIndex        =   10
         Top             =   480
         Visible         =   0   'False
         Width           =   10335
         _Version        =   1441793
         _ExtentX        =   18230
         _ExtentY        =   11033
         _StockProps     =   68
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   4
         Color           =   32
         PaintManager.BoldSelected=   -1  'True
         ItemCount       =   3
         Item(0).Caption =   "Grupos"
         Item(0).ControlCount=   2
         Item(0).Control(0)=   "vGridGrpAccss"
         Item(0).Control(1)=   "Label2(6)"
         Item(1).Caption =   "Miembros"
         Item(1).ControlCount=   4
         Item(1).Control(0)=   "cboGrpAccssM"
         Item(1).Control(1)=   "lswGrpAccssM"
         Item(1).Control(2)=   "Label2(8)"
         Item(1).Control(3)=   "Label2(9)"
         Item(2).Caption =   "Informes Autorizados"
         Item(2).ControlCount=   4
         Item(2).Control(0)=   "cboGrpAccssR"
         Item(2).Control(1)=   "lswGrpAccssR"
         Item(2).Control(2)=   "Label2(10)"
         Item(2).Control(3)=   "Label2(11)"
         Begin XtremeSuiteControls.ListView lswGrpAccssR 
            Height          =   5415
            Left            =   -67720
            TabIndex        =   75
            Top             =   840
            Visible         =   0   'False
            Width           =   6615
            _Version        =   1441793
            _ExtentX        =   11668
            _ExtentY        =   9551
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
            Appearance      =   17
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.ListView lswGrpAccssM 
            Height          =   5415
            Left            =   -67720
            TabIndex        =   76
            Top             =   840
            Visible         =   0   'False
            Width           =   6375
            _Version        =   1441793
            _ExtentX        =   11245
            _ExtentY        =   9551
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
            Appearance      =   17
            UseVisualStyle  =   0   'False
         End
         Begin FPSpreadADO.fpSpread vGridGrpAccss 
            Height          =   5655
            Left            =   1560
            TabIndex        =   11
            Top             =   480
            Width           =   7575
            _Version        =   524288
            _ExtentX        =   13361
            _ExtentY        =   9975
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
            MaxCols         =   496
            ScrollBars      =   2
            SpreadDesigner  =   "frmAF_Reportes.frx":1215
            VScrollSpecialType=   2
            AppearanceStyle =   1
         End
         Begin XtremeSuiteControls.ComboBox cboGrpAccssM 
            Height          =   330
            Left            =   -67720
            TabIndex        =   77
            Top             =   480
            Visible         =   0   'False
            Width           =   6375
            _Version        =   1441793
            _ExtentX        =   11245
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
         Begin XtremeSuiteControls.ComboBox cboGrpAccssR 
            Height          =   330
            Left            =   -67720
            TabIndex        =   78
            Top             =   480
            Visible         =   0   'False
            Width           =   6615
            _Version        =   1441793
            _ExtentX        =   11668
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
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            Caption         =   "Grupos de Acceso"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   555
            Index           =   6
            Left            =   480
            TabIndex        =   16
            Top             =   480
            Width           =   1095
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            Caption         =   "Grupo"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   315
            Index           =   8
            Left            =   -68920
            TabIndex        =   15
            Top             =   480
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            Caption         =   "Miembros"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   315
            Index           =   9
            Left            =   -68920
            TabIndex        =   14
            Top             =   840
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            Caption         =   "Grupo"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   315
            Index           =   10
            Left            =   -68920
            TabIndex        =   13
            Top             =   480
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFC0C0&
            Caption         =   "Reportes"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   315
            Index           =   11
            Left            =   -68920
            TabIndex        =   12
            Top             =   840
            Visible         =   0   'False
            Width           =   1215
         End
      End
      Begin MSComctlLib.TreeView ArbolExp 
         Height          =   6720
         Left            =   0
         TabIndex        =   17
         Top             =   0
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   11853
         _Version        =   393217
         HideSelection   =   0   'False
         Indentation     =   176
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
         FullRowSelect   =   -1  'True
         HotTracking     =   -1  'True
         ImageList       =   "imgArbol"
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
      End
      Begin XtremeSuiteControls.DateTimePicker dtpInicio 
         Height          =   315
         Left            =   6480
         TabIndex        =   18
         Top             =   480
         Width           =   1335
         _Version        =   1441793
         _ExtentX        =   2350
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
         Left            =   8520
         TabIndex        =   19
         Top             =   480
         Width           =   1335
         _Version        =   1441793
         _ExtentX        =   2350
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
      Begin XtremeSuiteControls.ComboBox cboTipo 
         Height          =   330
         Left            =   5760
         TabIndex        =   20
         Top             =   120
         Width           =   4095
         _Version        =   1441793
         _ExtentX        =   7223
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
      Begin XtremeSuiteControls.ComboBox cboFBase 
         Height          =   330
         Left            =   6480
         TabIndex        =   21
         Top             =   840
         Width           =   1335
         _Version        =   1441793
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
      Begin XtremeSuiteControls.ComboBox cboUsuarios 
         Height          =   330
         Left            =   5760
         TabIndex        =   22
         Top             =   1680
         Width           =   4095
         _Version        =   1441793
         _ExtentX        =   7223
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
         Height          =   255
         Left            =   8760
         TabIndex        =   23
         Top             =   840
         Width           =   975
         _Version        =   1441793
         _ExtentX        =   1714
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Todas"
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
         Transparent     =   -1  'True
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         Alignment       =   1
      End
      Begin XtremeSuiteControls.ComboBox cboEstado 
         Height          =   330
         Left            =   5760
         TabIndex        =   31
         Top             =   1320
         Width           =   4095
         _Version        =   1441793
         _ExtentX        =   7223
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
      Begin XtremeSuiteControls.TabControl tcMain 
         Height          =   3855
         Left            =   4320
         TabIndex        =   32
         Top             =   2160
         Width           =   6015
         _Version        =   1441793
         _ExtentX        =   10610
         _ExtentY        =   6800
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
         Item(0).Caption =   "General"
         Item(0).ControlCount=   16
         Item(0).Control(0)=   "cboProfesion"
         Item(0).Control(1)=   "cboSector"
         Item(0).Control(2)=   "cboSexo"
         Item(0).Control(3)=   "cboEstadoCivil"
         Item(0).Control(4)=   "Label1(24)"
         Item(0).Control(5)=   "Label1(23)"
         Item(0).Control(6)=   "Label1(22)"
         Item(0).Control(7)=   "Label1(21)"
         Item(0).Control(8)=   "Label1(20)"
         Item(0).Control(9)=   "cboInstitucion"
         Item(0).Control(10)=   "cboPromotor"
         Item(0).Control(11)=   "Label1(11)"
         Item(0).Control(12)=   "Label1(12)"
         Item(0).Control(13)=   "chkComites"
         Item(0).Control(14)=   "chkPromotor"
         Item(0).Control(15)=   "cboEstadoLaboral"
         Item(1).Caption =   "Otros"
         Item(1).ControlCount=   21
         Item(1).Control(0)=   "cboZonas"
         Item(1).Control(1)=   "cboProvincia"
         Item(1).Control(2)=   "cboCanton"
         Item(1).Control(3)=   "cboDistrito"
         Item(1).Control(4)=   "txtDeptCodigo"
         Item(1).Control(5)=   "txtDeptDesc"
         Item(1).Control(6)=   "txtSecCodigo"
         Item(1).Control(7)=   "txtSecDesc"
         Item(1).Control(8)=   "chkProvincias"
         Item(1).Control(9)=   "chkCantones"
         Item(1).Control(10)=   "chkDistritos"
         Item(1).Control(11)=   "chkDepartamento"
         Item(1).Control(12)=   "chkSeccion"
         Item(1).Control(13)=   "Label1(25)"
         Item(1).Control(14)=   "lblSeccion"
         Item(1).Control(15)=   "lblDepartamento"
         Item(1).Control(16)=   "Label9"
         Item(1).Control(17)=   "Label10"
         Item(1).Control(18)=   "Label18"
         Item(1).Control(19)=   "cboTipoId"
         Item(1).Control(20)=   "Label1(0)"
         Begin XtremeSuiteControls.ComboBox cboProfesion 
            Height          =   330
            Left            =   1440
            TabIndex        =   33
            Top             =   1440
            Width           =   4095
            _Version        =   1441793
            _ExtentX        =   7223
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
         Begin XtremeSuiteControls.ComboBox cboSector 
            Height          =   330
            Left            =   1440
            TabIndex        =   34
            Top             =   1800
            Width           =   4095
            _Version        =   1441793
            _ExtentX        =   7223
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
         Begin XtremeSuiteControls.ComboBox cboSexo 
            Height          =   330
            Left            =   1440
            TabIndex        =   35
            Top             =   2400
            Width           =   4095
            _Version        =   1441793
            _ExtentX        =   7223
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
         Begin XtremeSuiteControls.ComboBox cboEstadoCivil 
            Height          =   330
            Left            =   1440
            TabIndex        =   36
            Top             =   2760
            Width           =   4095
            _Version        =   1441793
            _ExtentX        =   7223
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
         Begin XtremeSuiteControls.ComboBox cboEstadoLaboral 
            Height          =   330
            Left            =   1440
            TabIndex        =   37
            Top             =   3360
            Width           =   4095
            _Version        =   1441793
            _ExtentX        =   7223
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
         Begin XtremeSuiteControls.ComboBox cboInstitucion 
            Height          =   330
            Left            =   1440
            TabIndex        =   38
            Top             =   360
            Width           =   4095
            _Version        =   1441793
            _ExtentX        =   7223
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
         Begin XtremeSuiteControls.ComboBox cboPromotor 
            Height          =   330
            Left            =   1440
            TabIndex        =   39
            Top             =   720
            Width           =   4095
            _Version        =   1441793
            _ExtentX        =   7223
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
         Begin XtremeSuiteControls.ComboBox cboZonas 
            Height          =   330
            Left            =   -68560
            TabIndex        =   40
            Top             =   840
            Visible         =   0   'False
            Width           =   3615
            _Version        =   1441793
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
         Begin XtremeSuiteControls.ComboBox cboProvincia 
            Height          =   330
            Left            =   -68560
            TabIndex        =   41
            Top             =   1200
            Visible         =   0   'False
            Width           =   3615
            _Version        =   1441793
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
         Begin XtremeSuiteControls.ComboBox cboCanton 
            Height          =   330
            Left            =   -68560
            TabIndex        =   42
            Top             =   1560
            Visible         =   0   'False
            Width           =   3615
            _Version        =   1441793
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
         Begin XtremeSuiteControls.ComboBox cboDistrito 
            Height          =   330
            Left            =   -68560
            TabIndex        =   43
            Top             =   1920
            Visible         =   0   'False
            Width           =   3615
            _Version        =   1441793
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
         Begin XtremeSuiteControls.FlatEdit txtDeptCodigo 
            Height          =   315
            Left            =   -69640
            TabIndex        =   44
            Top             =   2760
            Visible         =   0   'False
            Width           =   855
            _Version        =   1441793
            _ExtentX        =   1503
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
         Begin XtremeSuiteControls.FlatEdit txtDeptDesc 
            Height          =   315
            Left            =   -68800
            TabIndex        =   45
            Top             =   2760
            Visible         =   0   'False
            Width           =   3855
            _Version        =   1441793
            _ExtentX        =   6800
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
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtSecCodigo 
            Height          =   315
            Left            =   -69640
            TabIndex        =   46
            Top             =   3480
            Visible         =   0   'False
            Width           =   855
            _Version        =   1441793
            _ExtentX        =   1503
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
         Begin XtremeSuiteControls.FlatEdit txtSecDesc 
            Height          =   315
            Left            =   -68800
            TabIndex        =   47
            Top             =   3480
            Visible         =   0   'False
            Width           =   3855
            _Version        =   1441793
            _ExtentX        =   6800
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
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.CheckBox chkProvincias 
            Height          =   255
            Left            =   -64840
            TabIndex        =   48
            Top             =   1200
            Visible         =   0   'False
            Width           =   855
            _Version        =   1441793
            _ExtentX        =   1503
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Todas"
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
            Transparent     =   -1  'True
            UseVisualStyle  =   -1  'True
            Appearance      =   16
            Value           =   1
         End
         Begin XtremeSuiteControls.CheckBox chkCantones 
            Height          =   255
            Left            =   -64840
            TabIndex        =   49
            Top             =   1560
            Visible         =   0   'False
            Width           =   855
            _Version        =   1441793
            _ExtentX        =   1503
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Todos"
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
            Transparent     =   -1  'True
            UseVisualStyle  =   -1  'True
            Appearance      =   16
            Value           =   1
         End
         Begin XtremeSuiteControls.CheckBox chkDistritos 
            Height          =   255
            Left            =   -64840
            TabIndex        =   50
            Top             =   1920
            Visible         =   0   'False
            Width           =   855
            _Version        =   1441793
            _ExtentX        =   1503
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Todos"
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
            Transparent     =   -1  'True
            UseVisualStyle  =   -1  'True
            Appearance      =   16
            Value           =   1
         End
         Begin XtremeSuiteControls.CheckBox chkDepartamento 
            Height          =   255
            Left            =   -64840
            TabIndex        =   51
            Top             =   2760
            Visible         =   0   'False
            Width           =   975
            _Version        =   1441793
            _ExtentX        =   1714
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Todos"
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
            Transparent     =   -1  'True
            UseVisualStyle  =   -1  'True
            Appearance      =   16
            Value           =   1
         End
         Begin XtremeSuiteControls.CheckBox chkSeccion 
            Height          =   255
            Left            =   -64840
            TabIndex        =   52
            Top             =   3480
            Visible         =   0   'False
            Width           =   975
            _Version        =   1441793
            _ExtentX        =   1714
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Todas"
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
            Transparent     =   -1  'True
            UseVisualStyle  =   -1  'True
            Appearance      =   16
            Value           =   1
         End
         Begin XtremeSuiteControls.CheckBox chkPromotor 
            Height          =   252
            Left            =   1440
            TabIndex        =   53
            Top             =   1080
            Width           =   1812
            _Version        =   1441793
            _ExtentX        =   3196
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Promotores"
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
            Transparent     =   -1  'True
            UseVisualStyle  =   -1  'True
            Appearance      =   17
            Value           =   1
         End
         Begin XtremeSuiteControls.CheckBox chkComites 
            Height          =   252
            Left            =   3720
            TabIndex        =   54
            Top             =   1080
            Width           =   1452
            _Version        =   1441793
            _ExtentX        =   2561
            _ExtentY        =   444
            _StockProps     =   79
            Caption         =   "Cómites"
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
            Transparent     =   -1  'True
            UseVisualStyle  =   -1  'True
            Appearance      =   17
            Value           =   1
         End
         Begin XtremeSuiteControls.ComboBox cboTipoId 
            Height          =   330
            Left            =   -68560
            TabIndex        =   79
            Top             =   360
            Visible         =   0   'False
            Width           =   3615
            _Version        =   1441793
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
         Begin VB.Label Label1 
            Caption         =   "Tipo Ids"
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
            Left            =   -69640
            TabIndex        =   80
            Top             =   360
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.Label Label1 
            Caption         =   "Sector"
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
            Index           =   24
            Left            =   240
            TabIndex        =   67
            Top             =   1800
            Width           =   972
         End
         Begin VB.Label Label1 
            Caption         =   "Profesión"
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
            Index           =   23
            Left            =   240
            TabIndex        =   66
            Top             =   1440
            Width           =   972
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Condición Laboral"
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
            Index           =   22
            Left            =   240
            TabIndex        =   65
            Top             =   3360
            Width           =   1092
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Estado Civil"
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
            Index           =   21
            Left            =   240
            TabIndex        =   64
            Top             =   2760
            Width           =   972
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Sexo"
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
            Index           =   20
            Left            =   240
            TabIndex        =   63
            Top             =   2400
            Width           =   972
         End
         Begin VB.Label Label1 
            Caption         =   "Promotor"
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
            Left            =   240
            TabIndex        =   62
            Top             =   720
            Width           =   972
         End
         Begin VB.Label Label1 
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
            Index           =   12
            Left            =   240
            TabIndex        =   61
            Top             =   360
            Width           =   972
         End
         Begin VB.Label Label1 
            Caption         =   "Zonas"
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
            Index           =   25
            Left            =   -69640
            TabIndex        =   60
            Top             =   840
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.Label lblSeccion 
            Caption         =   "Sección"
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
            Left            =   -69640
            TabIndex        =   59
            Top             =   3240
            Visible         =   0   'False
            Width           =   2295
         End
         Begin VB.Label lblDepartamento 
            Caption         =   "Departamento"
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
            Left            =   -69640
            TabIndex        =   58
            Top             =   2520
            Visible         =   0   'False
            Width           =   2295
         End
         Begin VB.Label Label9 
            Caption         =   "Cantón"
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
            Left            =   -69640
            TabIndex        =   57
            Top             =   1560
            Visible         =   0   'False
            Width           =   615
         End
         Begin VB.Label Label10 
            Caption         =   "Provincia"
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
            Left            =   -69640
            TabIndex        =   56
            Top             =   1200
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.Label Label18 
            Caption         =   "Distrito"
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
            Left            =   -69640
            TabIndex        =   55
            Top             =   1920
            Visible         =   0   'False
            Width           =   735
         End
      End
      Begin XtremeSuiteControls.PushButton btnReporte 
         Height          =   495
         Left            =   8280
         TabIndex        =   68
         Top             =   6240
         Width           =   1575
         _Version        =   1441793
         _ExtentX        =   2778
         _ExtentY        =   873
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
         Picture         =   "frmAF_Reportes.frx":178D
         ImageAlignment  =   4
      End
      Begin VB.Label lblSeguridad 
         Caption         =   "[ Requiere Grupo de Acceso Autorizado ]"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   4800
         TabIndex        =   69
         Top             =   6270
         Width           =   3375
      End
      Begin VB.Image imgSeguridad 
         Height          =   255
         Left            =   4440
         Picture         =   "frmAF_Reportes.frx":1E94
         Stretch         =   -1  'True
         ToolTipText     =   "Requiere Grupo de Acceso Autorizado"
         Top             =   6240
         Width           =   255
      End
      Begin VB.Label Label1 
         Caption         =   "Tipo Reporte"
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
         Left            =   4680
         TabIndex        =   30
         Top             =   120
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Fechas"
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
         Left            =   4680
         TabIndex        =   29
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         Caption         =   "Inicio"
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
         Height          =   315
         Index           =   5
         Left            =   5760
         TabIndex        =   28
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         Caption         =   "Corte"
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
         Height          =   315
         Index           =   6
         Left            =   7800
         TabIndex        =   27
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         Caption         =   "Base"
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
         Height          =   330
         Index           =   7
         Left            =   5760
         TabIndex        =   26
         Top             =   840
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Estados"
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
         Index           =   8
         Left            =   4680
         TabIndex        =   25
         Top             =   1320
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Usuarios"
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
         Index           =   14
         Left            =   4680
         TabIndex        =   24
         Top             =   1680
         Width           =   735
      End
      Begin XtremeShortcutBar.ShortcutCaption scTitulosTabs 
         Height          =   375
         Index           =   1
         Left            =   -70000
         TabIndex        =   9
         Top             =   0
         Visible         =   0   'False
         Width           =   10335
         _Version        =   1441793
         _ExtentX        =   18230
         _ExtentY        =   661
         _StockProps     =   14
         Caption         =   "Grupos de Acceso"
         ForeColor       =   4210752
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SubItemCaption  =   -1  'True
         Alignment       =   1
         ForeColor       =   4210752
      End
      Begin XtremeShortcutBar.ShortcutCaption scTitulosTabs 
         Height          =   375
         Index           =   0
         Left            =   -70000
         TabIndex        =   8
         Top             =   0
         Visible         =   0   'False
         Width           =   10455
         _Version        =   1441793
         _ExtentX        =   18441
         _ExtentY        =   661
         _StockProps     =   14
         Caption         =   "Grupos de Trabajo"
         ForeColor       =   4210752
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SubItemCaption  =   -1  'True
         Alignment       =   1
         ForeColor       =   4210752
      End
   End
   Begin MSComctlLib.ImageList imgArbol 
      Left            =   0
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
            Picture         =   "frmAF_Reportes.frx":2590
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAF_Reportes.frx":8DF2
            Key             =   "imgCRD"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAF_Reportes.frx":8F10
            Key             =   "imgCBR"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAF_Reportes.frx":903A
            Key             =   "imgSGT"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAF_Reportes.frx":9160
            Key             =   "imgDetalle"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAF_Reportes.frx":926E
            Key             =   "imgRoot"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAF_Reportes.frx":937B
            Key             =   "imgEspecial"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAF_Reportes.frx":9494
            Key             =   "imgRetenciones"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAF_Reportes.frx":95C2
            Key             =   "imgSeguridad"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAF_Reportes.frx":96CF
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label6 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Reportes de Clientes"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   492
      Left            =   2040
      TabIndex        =   1
      Top             =   240
      Width           =   4572
   End
   Begin VB.Label lblReporte 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "..."
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
      Height          =   252
      Left            =   2820
      TabIndex        =   0
      Top             =   660
      Width           =   6972
   End
   Begin VB.Image imgBanner 
      Height          =   1092
      Left            =   0
      Top             =   0
      Width           =   10332
   End
End
Attribute VB_Name = "frmAF_Reportes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vPaso As Boolean
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem


Private Sub sbInicializa()

Me.MousePointer = vbHourglass

tcPrincipal.Item(0).Selected = True
tcMain.Item(0).Selected = True

imgSeguridad.Visible = False
lblSeguridad.Visible = imgSeguridad.Visible

lblReporte.Tag = ""
lblReporte.Caption = ">>> Seleccione Un Reporte <<<"

dtpInicio.Value = fxFechaServidor
dtpCorte.Value = dtpInicio.Value
chkFechas.Value = vbUnchecked

chkDepartamento.Value = vbChecked
chkSeccion.Value = vbChecked

cboFBase.Clear
cboFBase.AddItem "Ingreso"
cboFBase.AddItem "Nacimiento"
cboFBase.AddItem "Registro"
cboFBase.AddItem "Actualiza"
cboFBase.Text = "Ingreso"

cboTipo.Clear
cboTipo.AddItem "Detalle"
cboTipo.AddItem "Resumen"
cboTipo.Text = "Detalle"


vPaso = True
    'Provincias
    strSQL = "select Provincia as Idx, rtrim(Descripcion) as ItmX from Provincias"
    Call sbCbo_Llena_New(cboProvincia, strSQL, True, True)
    
    'Grupos de Usuarios
    strSQL = "select rtrim(cod_grupo) as 'IdX', rtrim(descripcion) as 'ItmX'" _
             & " from  AFI_grupos"
    Call sbCbo_Llena_New(cboUsuarios, strSQL, True, True)
    
    'Instituciones
    strSQL = "select rtrim(descripcion) as Itmx, cod_institucion as Idx" _
           & " from instituciones order by descripcion"
    Call sbCbo_Llena_New(cboInstitucion, strSQL, True, True)


    'Promotores
    strSQL = "select id_promotor as 'IdX',rtrim(Nombre) as 'ItmX' from promotores order by Nombre"
    Call sbCbo_Llena_New(cboPromotor, strSQL, True, True)
    
    'Estados
    strSQL = "select rtrim(cod_estado) as 'IdX', rtrim(descripcion) as ItmX" _
             & " from  afi_Estados_Persona"
    Call sbCbo_Llena_New(cboEstado, strSQL, True, True)
    
    'Estado Civil
    strSQL = "select Estado_Civil as 'IdX', Descripcion as 'ItmX' from SYS_ESTADO_CIVIL" _
           & " where Activo = 1" _
           & " order by Descripcion asc"
    Call sbCbo_Llena_New(cboEstadoCivil, strSQL, True, True)

    'Profesiones
    strSQL = "select COD_PROFESION as 'Idx',descripcion as 'ItmX' from AFI_PROFESIONES order by descripcion"
    Call sbCbo_Llena_New(cboProfesion, strSQL, True, True)
    
    'Sectores
    strSQL = "select COD_SECTOR as 'Idx',descripcion as 'ItmX' from AFI_SECTORES order by descripcion"
    Call sbCbo_Llena_New(cboSector, strSQL, True, True)
    
    'Zonas
    strSQL = "select rtrim(COD_ZONA) as 'IdX', rtrim(descripcion) as 'ItmX' from AFI_ZONAS order by descripcion"
    Call sbCbo_Llena_New(cboZonas, strSQL, True, True)

    'Tipos Ids
    strSQL = "select TIPO_ID as Idx, rtrim(Descripcion) as ItmX from AFI_TIPOS_IDS" _
           & " order by Tipo_Id"
    Call sbCbo_Llena_New(cboTipoId, strSQL, True, True)


vPaso = False



cboSexo.Clear
cboSexo.AddItem "TODOS"
cboSexo.AddItem "Femenino"
cboSexo.AddItem "Masculino"
cboSexo.Text = "TODOS"

'Estado Laboral
strSQL = "select Estado_Laboral as 'IdX', Descripcion as 'ItmX' from AFI_ESTADO_LABORAL" _
       & " where Activo = 1" _
       & " order by Descripcion asc"
Call sbCbo_Llena_New(cboEstadoLaboral, strSQL, True, True)



Call chkFechas_Click



Call sbRefrescaArbol

Me.MousePointer = vbDefault

End Sub


Private Function fxIndiceCodigo(xkey As String) As String
xkey = Mid(xkey, 4, Len(xkey))
xkey = Mid(xkey, 1, Len(xkey) - 1)
fxIndiceCodigo = xkey
End Function

Private Sub ArbolExp_NodeClick(ByVal Node As MSComctlLib.Node)
Dim i As Integer, rsTmp As New ADODB.Recordset

On Error GoTo vError

If Right(Node.Key, 1) = "Z" Then
  lblReporte.Caption = Node.Text
  lblReporte.Tag = fxIndiceCodigo(Node.Key)
  
  strSQL = "select Tipo,isnull(Seguridad,0) as Seguridad" _
         & " from afi_reportes where id_rep = " & lblReporte.Tag
  Call OpenRecordSet(rs, strSQL)
  
  If rs!seguridad = 1 Then
     imgSeguridad.Visible = True
     
    'Verificar que la persona tenga acceso a este reporte
    strSQL = "select isnull(COUNT(*),0) as Existe" _
           & " From afi_reportes_GRP_AUT where id_rep = " & lblReporte.Tag _
           & " and cod_grupo in(select cod_grupo from afi_reportes_grp_usr where usuario = '" & glogon.Usuario & "')"
    Call OpenRecordSet(rsTmp, strSQL, 0)
    If rsTmp!Existe = 0 Then
       lblSeguridad.Caption = "[ Requiere Grupo de Acceso Autorizado ]"
       lblSeguridad.ForeColor = vbRed
    Else
       lblSeguridad.Caption = "[ Usuario Tiene Acceso Autorizado ]"
       lblSeguridad.ForeColor = vbBlue
    End If
    rsTmp.Close
     
  Else
     imgSeguridad.Visible = False
  End If
  lblSeguridad.Visible = imgSeguridad.Visible
       
       
  tcMain.Item(0).Selected = True
  

  rs.Close
End If

vError:

End Sub



Private Sub btnReporte_Click()
Call sbReporteIng
End Sub

Private Sub cboCanton_Click()

If vPaso Then Exit Sub

    strSQL = "select Distrito as Idx, rtrim(Descripcion) as ItmX from Distritos" _
            & " where provincia = '" & cboProvincia.ItemData(cboProvincia.ListIndex) _
            & "' and Canton = '" & cboCanton.ItemData(cboCanton.ListIndex) _
            & "' order by descripcion"
    Call sbCbo_Llena_New(cboDistrito, strSQL, False, True)

'Agrega Distrito En Limpio, ya que este dato es opcional
cboDistrito.AddItem " "
cboDistrito.Text = " "

End Sub

Private Sub cboCanton_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then cboDistrito.SetFocus
End Sub


Private Sub cboGrpAccssM_Click()

If vPaso Then Exit Sub

If cboGrpAccssM.ListCount <= 0 Then Exit Sub

vPaso = True

With lswGrpAccssM
 .ListItems.Clear
  
 strSQL = "select U.nombre,U.descripcion,A.usuario" _
        & " from Usuarios U left join afi_reportes_GRP_USR A on U.nombre = A.usuario" _
        & " and U.estado = 'A'  and A.cod_grupo = " & cboGrpAccssM.ItemData(cboGrpAccssM.ListIndex) _
        & " order by A.usuario desc,U.nombre asc"
 Call OpenRecordSet(rs, strSQL, 0)
 Do While Not rs.EOF
  Set itmX = .ListItems.Add(, , rs!Nombre)
      itmX.SubItems(1) = rs!Descripcion
      If Not IsNull(rs!Usuario) Then
         itmX.Checked = vbChecked
         itmX.ForeColor = vbBlue
      End If
  rs.MoveNext
 Loop
 rs.Close
End With

vPaso = False

End Sub


Private Sub cboGrpAccssR_Click()

If vPaso Then Exit Sub

If cboGrpAccssR.ListCount <= 0 Then Exit Sub

vPaso = True

With lswGrpAccssR
 .ListItems.Clear
 strSQL = "select R.tipo,R.id_REP,R.reporte,A.cod_grupo" _
        & " from afi_reportes R left join afi_reportes_GRP_AUT A on R.id_REP = A.id_REP" _
        & " and A.cod_grupo = " & cboGrpAccssR.ItemData(cboGrpAccssR.ListIndex) _
        & " order by A.cod_grupo desc,R.tipo asc, R.id_REP asc"
 Call OpenRecordSet(rs, strSQL, 0)
 Do While Not rs.EOF
  Set itmX = .ListItems.Add(, , rs!ID_REP)
      itmX.SubItems(1) = rs!Tipo
      itmX.SubItems(2) = rs!Reporte
      If Not IsNull(rs!Cod_Grupo) Then
         itmX.Checked = vbChecked
         itmX.ForeColor = vbBlue
      End If
  rs.MoveNext
 Loop
 rs.Close
End With

vPaso = False

End Sub

Private Sub cboMiembros_Click()

If vPaso Then Exit Sub

If cboMiembros.ListCount <= 0 Then Exit Sub

vPaso = True

With lswMiembros
 .ListItems.Clear
  
 strSQL = "select U.nombre,U.descripcion,A.usuario" _
        & " from Usuarios U left join afi_grpusers A on U.nombre = A.usuario" _
        & " and U.estado = 'A'  and A.cod_grupo = '" & cboMiembros.ItemData(cboMiembros.ListIndex) & "'" _
        & " order by A.usuario desc,U.nombre asc"
 Call OpenRecordSet(rs, strSQL, 0)
 Do While Not rs.EOF
  Set itmX = .ListItems.Add(, , rs!Nombre)
      itmX.SubItems(1) = rs!Descripcion
      If Not IsNull(rs!Usuario) Then
         itmX.Checked = vbChecked
         itmX.ForeColor = vbBlue
      End If
  rs.MoveNext
 Loop
 rs.Close
End With

vPaso = False

End Sub

Private Sub cboPromotor_Click()

If vPaso Then Exit Sub

If cboPromotor.ListCount <= 0 Then Exit Sub

If cboPromotor.Text = "TODOS" Then
   chkComites.Value = vbChecked
   chkComites.Enabled = True
Else
   chkComites.Value = vbUnchecked
   chkComites.Enabled = False
End If

chkPromotor.Value = chkComites.Value
chkPromotor.Enabled = chkComites.Enabled

End Sub

Private Sub cboProvincia_Click()

If vPaso Then Exit Sub

vPaso = True
    strSQL = "select Canton as Idx, rtrim(Descripcion) as ItmX from Cantones" _
           & " where provincia = '" & cboProvincia.ItemData(cboProvincia.ListIndex) & "' order by descripcion"
    Call sbCbo_Llena_New(cboCanton, strSQL, False, True)
vPaso = False

Call cboCanton_Click

End Sub

Private Sub chkFechas_Click()

If chkFechas.Value = vbChecked Then
   dtpInicio.Enabled = False
Else
   dtpInicio.Enabled = True
End If

dtpCorte.Enabled = dtpInicio.Enabled
cboFBase.Enabled = dtpInicio.Enabled

End Sub


Sub sbCreaNodos(vPadre As String, vTexto As String, vImagen As String, vExpand As Boolean, Optional xkey As String = "N")
Dim nodX As Node, vKey As String
On Error Resume Next

Set nodX = ArbolExp.Nodes.Add(vPadre, tvwChild)
    nodX.Text = vTexto
    nodX.Tag = nodX.Index
    nodX.Image = vImagen
    If xkey = "N" Then
        nodX.Key = vTexto & "0x0" & ArbolExp.Nodes.Count & "ID"
    Else
        nodX.Key = xkey
    End If
    
vKey = nodX.Key

If vExpand Then
    Set nodX = ArbolExp.Nodes.Add(vKey, tvwChild)
        nodX.Key = "F" & vTexto & "0x0" & ArbolExp.Nodes.Count & "ID"
        nodX.Tag = nodX.Index
End If
    
End Sub


Private Sub sbRefrescaArbol()
Dim vNode As Node, strOpciones  As String

With ArbolExp
  .Nodes.Clear
  Set vNode = .Nodes.Add(, , "Reportes", "Reportes", "imgRoot")
  Call sbCreaNodos("Reportes", "Ingreso", "imgCRD", False, "0x0ING")
  Call sbCreaNodos("Reportes", "Renuncias", "imgSGT", False, "0x0REN")
  Call sbCreaNodos("Reportes", "Liquidaciones", "imgCBR", False, "0x0LIQ")
  Call sbCreaNodos("Reportes", "Especiales", "imgEspecial", False, "0x0ESP")
  
  strSQL = "select Id_rep,Reporte,Tipo,isnull(seguridad,0) as Seguridad from afi_reportes order by reporte"
  Call OpenRecordSet(rs, strSQL)
  Do While Not rs.EOF
        If rs!seguridad = 0 Then
            Call sbCreaNodos("0x0" & Trim(UCase(rs!Tipo)), rs!Reporte, "imgDetalle", False, "0x0" & rs!ID_REP & "Z")
        Else
            Call sbCreaNodos("0x0" & Trim(UCase(rs!Tipo)), rs!Reporte, "imgSeguridad", False, "0x0" & rs!ID_REP & "Z")
        End If
    rs.MoveNext
  Loop
  rs.Close
  .Nodes(1).Expanded = True
End With

End Sub


Private Sub cboProvincia_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then cboCanton.SetFocus
End Sub

Private Sub Form_Activate()
vModulo = 1
End Sub

Private Sub Form_Load()

vModulo = 1

Set imgBanner.Picture = frmContenedor.imgBanner_Reportes.Picture

vGrid.AppearanceStyle = fxGridStyle
vGridGrpAccss.AppearanceStyle = vGrid.AppearanceStyle
vGridRep.AppearanceStyle = vGrid.AppearanceStyle

If GLOBALES.SysASEVersion Then
   lblDepartamento.Caption = "Unidad Programatica"
   lblSeccion.Caption = "Unidad de Trabajo"
  
Else
   lblDepartamento.Caption = "Departamento"
   lblSeccion.Caption = "Sección"
End If



With lswMiembros.ColumnHeaders
    .Clear
    .Add , , "Usuario", 2800
    .Add , , "Nombre", lswMiembros.Width - 2900
End With


With lswGrpAccssM.ColumnHeaders
    .Clear
    .Add , , "Usuario", 2800
    .Add , , "Descripción", lswGrpAccssM.Width - 2900
End With

With lswGrpAccssR.ColumnHeaders
    .Clear
    .Add , , "Id", 900
    .Add , , "Tipo", 1000, vbCenter
    .Add , , "Reporte", lswGrpAccssR.Width - 2100
End With

Call Formularios(Me)
Call RefrescaTags(Me)

End Sub


Private Sub imgAddRep_Click()
'Inicializa Base de Datos con Reportes
glogon.Conection.Execute "exec spAFIReportesGen"
MsgBox "Lista de Reportes Actualizada...", vbInformation

tcAux.Item(0).Selected = True

End Sub




Private Sub lswGrpAccssM_ItemCheck(ByVal Item As XtremeSuiteControls.ListViewItem)

On Error GoTo vError

If vPaso Then Exit Sub

If Item.Checked Then
  strSQL = "insert afi_reportes_GRP_USR(cod_grupo,usuario) values(" & cboGrpAccssM.ItemData(cboGrpAccssM.ListIndex) _
         & ",'" & Item.Text & "')"
Else
  strSQL = "delete afi_reportes_GRP_USR where cod_grupo = " & cboGrpAccssM.ItemData(cboGrpAccssM.ListIndex) _
         & " and usuario = '" & Item.Text & "'"
End If
Call ConectionExecute(strSQL)

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub



Private Sub lswGrpAccssR_ItemCheck(ByVal Item As XtremeSuiteControls.ListViewItem)
If vPaso Then Exit Sub

On Error GoTo vError

If Item.Checked Then
  strSQL = "insert afi_reportes_GRP_AUT(cod_grupo,id_REP) values(" & cboGrpAccssR.ItemData(cboGrpAccssR.ListIndex) _
         & "," & Item.Text & ")"
Else
  strSQL = "delete afi_reportes_GRP_AUT where cod_grupo = " & cboGrpAccssR.ItemData(cboGrpAccssR.ListIndex) _
         & " and id_REP = " & Item.Text
End If
Call ConectionExecute(strSQL)

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub lswMiembros_ItemCheck(ByVal Item As XtremeSuiteControls.ListViewItem)

If vPaso Then Exit Sub

On Error GoTo vError


If Item.Checked Then
  'Preguntar si ya Existe el Usuario en Otro Grupo. / de ser asi no continuar
  strSQL = "select isnull(count(*),0) as Existe from afi_grpusers where cod_grupo <> '" _
         & cboMiembros.ItemData(cboMiembros.ListIndex) & "' and usuario = '" & Item.Text & "'"
  Call OpenRecordSet(rs, strSQL)
  If rs!Existe > 0 Then
     rs.Close
     Item.Checked = False
     MsgBox "El Usuario ya ha sido asignado a otro grupo, proceda a excluirlo primero del otro grupo antes de agregarlo a este", vbExclamation
     Exit Sub
  End If
  rs.Close
End If


If Item.Checked Then
  strSQL = "insert afi_grpusers(cod_grupo,usuario) values('" & cboMiembros.ItemData(cboMiembros.ListIndex) _
         & "','" & Item.Text & "')"
Else
  strSQL = "delete afi_grpusers where cod_grupo = '" & cboMiembros.ItemData(cboMiembros.ListIndex) _
         & "' and usuario = '" & Item.Text & "'"
End If
Call ConectionExecute(strSQL)

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
  
End Sub





Private Function fxReporteFile() As String
Dim strSQL As String, rs As New ADODB.Recordset

strSQL = "select prefijo from afi_reportes where id_rep = " & lblReporte.Tag
Call OpenRecordSet(rs, strSQL)
If rs.EOF And rs.BOF Then
  fxReporteFile = ""
Else
  fxReporteFile = Trim(rs!prefijo)
End If
rs.Close

End Function




Private Sub tcAux_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)

Select Case Item.Index
  Case 0 'Grupos
    strSQL = "select cod_grupo,descripcion from afi_grupos order by cod_grupo"
    Call sbCargaGrid(vGrid, 2, strSQL)
    
  Case 1 'Miembros
    vPaso = True
    strSQL = "select cod_grupo as 'Idx', rtrim(descripcion) as ItmX" _
         & " from  afi_grupos"
    Call sbCbo_Llena_New(cboMiembros, strSQL, False, True)
    vPaso = False
    
    Call cboMiembros_Click
    
  Case 2 'Reportes
    strSQL = "select ID_Rep,Tipo,Reporte,Prefijo,isnull(Seguridad,0) from afi_reportes order by tipo,reporte"
    Call sbCargaGrid(vGridRep, 5, strSQL)
End Select

End Sub

Private Sub tcAuxGrpAccs_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)

Select Case Item.Index
  Case 0 'Grupos
    strSQL = "select cod_grupo,descripcion,activo from afi_reportes_grp order by cod_grupo"
    Call sbCargaGrid(vGridGrpAccss, 3, strSQL)
  
  Case 1 'Miembros
    vPaso = True
    strSQL = "select cod_grupo as IdX, rtrim(descripcion) as ItmX" _
         & " from  afi_reportes_grp where activo = 1"
    Call sbCbo_Llena_New(cboGrpAccssM, strSQL, False, True)
    vPaso = False
    
    Call cboGrpAccssM_Click
  
  Case 2 'Reportes
    vPaso = True
    strSQL = "select cod_grupo as IdX, rtrim(descripcion) as ItmX" _
         & " from  afi_reportes_grp where activo = 1"
    Call sbCbo_Llena_New(cboGrpAccssR, strSQL, False, True)
    vPaso = False
    
    Call cboGrpAccssR_Click

End Select

End Sub

Private Sub tcPrincipal_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)

Select Case Item.Index
    Case 0 'Inicial
      Call sbInicializa
      
    Case 1 'Configuracion
      tcAux.Item(0).Selected = True
      
      strSQL = "select cod_grupo,descripcion from afi_grupos order by cod_grupo"
      Call sbCargaGrid(vGrid, 2, strSQL)
    
      lblReporte.Caption = "Configuración de Reportes"
      
    Case 2 'Seguridad
      tcAuxGrpAccs.Item(0).Selected = True
    
      strSQL = "select cod_grupo,descripcion,activo from afi_reportes_grp order by cod_grupo"
      Call sbCargaGrid(vGridGrpAccss, 3, strSQL)
    
      lblReporte.Caption = "Seguridad de Reportes"
    End Select

End Sub

Private Sub TimerX_Timer()
TimerX.Interval = 0
Call sbInicializa
End Sub


Private Sub sbReporteIng()
Dim vTitulo As String, vSubTitulo As String, vFiltro As String, vTemp As String


On Error GoTo vError

If lblReporte.Tag = "" Then Exit Sub

Me.MousePointer = vbHourglass


If imgSeguridad.Visible Then
  'Verificar que la persona tenga acceso a este reporte
  strSQL = "select isnull(COUNT(*),0) as Existe" _
         & " From afi_reportes_GRP_AUT where id_rep = " & lblReporte.Tag _
         & " and cod_grupo in(select cod_grupo from afi_reportes_grp_usr where usuario = '" & glogon.Usuario & "')"
  Call OpenRecordSet(rs, strSQL)
  If rs!Existe = 0 Then
     Me.MousePointer = vbDefault
     rs.Close
     MsgBox "El usuario actual no tiene acceso autorizado a este reporte, verifique...", vbExclamation
     Exit Sub
  End If
  rs.Close
End If

vTitulo = UCase(lblReporte.Caption & ": " & cboTipo.Text)
vSubTitulo = ""
vFiltro = ""
strSQL = ""


With frmContenedor.Crt
 .Reset
 .WindowShowGroupTree = True
 .WindowShowPrintSetupBtn = True
 .WindowShowRefreshBtn = True
 .WindowShowSearchBtn = True
 .WindowState = crptMaximized
 .WindowTitle = "Reportes del Módulo de Personas"
 
 .Connect = glogon.ConectRPT
  
  
 If chkFechas.Value = vbUnchecked Then
    If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
    Select Case Mid(cboFBase.Text, 1, 1)
      Case "I"
        strSQL = strSQL & "{vAFIAfiliacionReportes01.FechaIngreso}"
        vSubTitulo = "Ingreso entre " & Format(dtpInicio.Value, "dd/mm/yyyy") & " y " & Format(dtpCorte.Value, "dd/mm/yyyy")
      Case "N"
        strSQL = strSQL & " ( DAY({vAFIAfiliacionReportes01.Fecha_Nac}) >= " & Day(dtpInicio.Value) & " AND "
        strSQL = strSQL & "   DAY({vAFIAfiliacionReportes01.Fecha_Nac}) <= " & Day(dtpCorte.Value) & ") AND "
        
        strSQL = strSQL & " ( MONTH({vAFIAfiliacionReportes01.Fecha_Nac}) >= " & Month(dtpInicio.Value) & " AND "
        strSQL = strSQL & "   MONTH({vAFIAfiliacionReportes01.Fecha_Nac}) <= " & Month(dtpCorte.Value) & ")"
        
        vSubTitulo = "Nacimiento entre días [" & Day(dtpInicio.Value) & ".." & Day(dtpCorte.Value) & "] Meses [" & Month(dtpInicio.Value) & ".." & Month(dtpCorte.Value) & "]"
      Case "R"
        strSQL = strSQL & "{vAFIAfiliacionReportes01.reg_fecha}"
        vSubTitulo = "Registradas entre " & Format(dtpInicio.Value, "dd/mm/yyyy") & " y " & Format(dtpCorte.Value, "dd/mm/yyyy")
      Case "A"
        strSQL = strSQL & "{vAFIAfiliacionReportes01.ActualizaFecha}"
        vSubTitulo = "Actualizada entre " & Format(dtpInicio.Value, "dd/mm/yyyy") & " y " & Format(dtpCorte.Value, "dd/mm/yyyy")
    End Select
    
    If Mid(cboFBase.Text, 1, 1) <> "N" Then
        strSQL = strSQL & " in Date(" & Format(dtpInicio.Value, "yyyy,mm,dd") & ") to date(" _
               & Format(dtpCorte.Value, "yyyy,mm,dd") & ")"
    End If
 Else
   vSubTitulo = "Historico"
 End If
 
 

 
 
 If cboEstado.Text <> "TODOS" Then
       If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
       strSQL = strSQL & "{vAFIAfiliacionReportes01.estadoactual} = '" & cboEstado.ItemData(cboEstado.ListIndex) & "'"

End If
 vFiltro = vFiltro & "¦ Estado: " & cboEstado.Text
    
 If cboTipoId.Text <> "TODOS" Then
   If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
       strSQL = strSQL & "{vAFIAfiliacionReportes01.TIPO_ID} = " & cboTipoId.ItemData(cboTipoId.ListIndex)
  
       vFiltro = "¦ Tipo Id: " & cboTipoId.Text
 End If
    
    
 If Mid(cboSexo.Text, 1, 1) <> "T" Then
   If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
   strSQL = strSQL & "{vAFIAfiliacionReportes01.sexo} = '" & Mid(cboSexo.Text, 1, 1) & "'"
 End If
 vSubTitulo = vSubTitulo & "¦ Sexo: " & cboSexo.Text
  
  If Mid(cboEstadoCivil.Text, 1, 1) <> "T" Then
   If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
   strSQL = strSQL & "{vAFIAfiliacionReportes01.EstadoCivil} = '" & cboEstadoCivil.ItemData(cboEstadoCivil.ListIndex) & "'"
 End If
 vSubTitulo = vSubTitulo & "¦ Estado Civil: " & cboEstadoCivil.Text
 
 If cboEstadoLaboral.Text <> "TODOS" Then
   If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
     strSQL = strSQL & "{vAFIAfiliacionReportes01.EstadoLaboral} = '" & cboEstadoLaboral.ItemData(cboEstadoLaboral.ListIndex) & "'"
 End If
 vFiltro = vFiltro & "¦ Laboral: " & cboEstadoLaboral.Text
 

 If cboInstitucion.Text <> "TODOS" Then
   If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
   strSQL = strSQL & "{vAFIAfiliacionReportes01.cod_institucion} = " & cboInstitucion.ItemData(cboInstitucion.ListIndex) & ""
 End If
 vFiltro = vFiltro & "¦ Empresa: " & cboInstitucion.Text
  
 If cboPromotor.Text <> "TODOS" Then
   If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
   strSQL = strSQL & "{vAFIAfiliacionReportes01.id_promotor} = " & cboPromotor.ItemData(cboPromotor.ListIndex) & ""
 End If
 vFiltro = vFiltro & "¦ Promotor: " & cboPromotor.Text
 
 If chkComites.Enabled Then
    If chkComites.Value <> chkPromotor.Value Then
       If chkComites.Value = vbChecked Then
           'Solo Comites
           If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
           strSQL = strSQL & "{vAFIAfiliacionReportes01.Comite} = 1"
           
           vFiltro = vFiltro & " [Solo Comités]"
       
       Else
           'Solo Promotores
           If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
           strSQL = strSQL & "{vAFIAfiliacionReportes01.Comite} = 0"
       
           vFiltro = vFiltro & " [Solo Promotores]"
       End If
    End If
 End If
 
 
 If cboSector.Text <> "TODOS" Then
   If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
   strSQL = strSQL & "{vAFIAfiliacionReportes01.cod_sector} = " & cboSector.ItemData(cboSector.ListIndex) & ""
 End If
 vFiltro = vFiltro & "¦ Sector: " & cboSector.Text
 
 
 If cboProfesion.Text <> "TODOS" Then
   If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
   strSQL = strSQL & "{vAFIAfiliacionReportes01.cod_profesion} = " & cboProfesion.ItemData(cboProfesion.ListIndex) & ""
 End If
 vFiltro = vFiltro & "¦ Profesión: " & cboProfesion.Text
 
 'Unidades Departamentales / Programaticas
 If chkDepartamento.Value = vbUnchecked Then
   If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
   strSQL = strSQL & "{vAFIAfiliacionReportes01.CodDepartamento} = '" & txtDeptCodigo.Text & "'"
 
   vFiltro = vFiltro & "¦ Centro: " & txtDeptCodigo.Text
 End If
 
 
 If chkSeccion.Value = vbUnchecked Then
   If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
   strSQL = strSQL & "{vAFIAfiliacionReportes01.CodSeccion} = '" & txtSecCodigo.Text & "'"
 
   vFiltro = vFiltro & "¦ Sección : " & txtSecCodigo.Text
 
 End If
 
 'Geografica
 If cboProvincia.Text <> "TODOS" Then
   If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
   strSQL = strSQL & "{vAFIAfiliacionReportes01.Provincia} = " & cboProvincia.ItemData(cboProvincia.ListIndex) & ""
 
        If cboCanton.Text <> "TODOS" Then
          If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
          strSQL = strSQL & "{vAFIAfiliacionReportes01.Canton} = '" & cboCanton.ItemData(cboCanton.ListIndex) & "'"
        End If
        vFiltro = vFiltro & "¦ Cantón: " & cboCanton.Text
 
 End If
 vFiltro = vFiltro & "¦ Provincia: " & cboProvincia.Text
  

 
 
 .Formulas(0) = "fxFecha='FECHA: " & Format(fxFechaServidor, "dd/mm/yyyy") & "'"
 .Formulas(1) = "fxEmpresa='" & GLOBALES.gstrNombreEmpresa & "'"
 .Formulas(2) = "fxUsuario='USER: " & glogon.Usuario & "'"
 .Formulas(3) = "fxTitulo='" & vTitulo & "'"
 .Formulas(4) = "fxSubTitulo='" & vSubTitulo & "'"
 .Formulas(5) = "fxFiltro='" & vFiltro & "'"
 .ReportFileName = SIFGlobal.fxPathReportes("" & Trim(fxReporteFile) & "_" & Trim(cboTipo.Text) & ".rpt")
 .SelectionFormula = strSQL

 .PrintReport

End With

Me.MousePointer = vbDefault

Call Bitacora("Imprime", lblReporte.Caption)

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub


Private Function fxGuardar() As Long

On Error GoTo vError

fxGuardar = 0
vGrid.Row = vGrid.ActiveRow
vGrid.col = 1

strSQL = "select isnull(count(*),0) as Existe from afi_grupos" _
       & " where cod_grupo = '" & vGrid.Text & "'"
Call OpenRecordSet(rs, strSQL)

If rs!Existe = 0 Then 'Insertar
  If Trim(vGrid.Text) = "" Then Exit Function
  
  strSQL = "insert into afi_grupos(cod_grupo,descripcion) values('" _
         & UCase(vGrid.Text) & "','"
  vGrid.col = 2
  strSQL = strSQL & UCase(vGrid.Text) & "')"

  Call ConectionExecute(strSQL)

  vGrid.col = 1
  Call Bitacora("Registra", "Grupo de Usuarios: " & vGrid.Text)

Else 'Actualizar

 vGrid.col = 2
 strSQL = "update afi_grupos set descripcion = '" & vGrid.Text & "'"
 strSQL = strSQL & " where cod_grupo = '"
 vGrid.col = 1
 strSQL = strSQL & vGrid.Text & "'"
 Call ConectionExecute(strSQL)

 Call Bitacora("Modifica", "Grupo de Usuarios : " & vGrid.Text)


End If
rs.Close

fxGuardar = 1

Exit Function

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Function



Private Sub vGrid_KeyDown(KeyCode As Integer, Shift As Integer)
Dim i As Integer

If vGrid.ActiveCol = vGrid.MaxCols And (KeyCode = vbKeyReturn Or KeyCode = vbKeyTab) Then
  i = fxGuardar
  If i = 0 Then Exit Sub
  vGrid.Row = vGrid.ActiveRow
  If vGrid.MaxRows <= vGrid.ActiveRow Then
    vGrid.MaxRows = vGrid.MaxRows + 1
    vGrid.Row = vGrid.MaxRows
  End If
End If

'Inserta Linea
If KeyCode = vbKeyInsert Then
    vGrid.MaxRows = vGrid.MaxRows + 1
    vGrid.InsertRows vGrid.ActiveRow, 1
    vGrid.Row = vGrid.ActiveRow
End If

End Sub

Private Function fxGuardarRep() As Long

On Error GoTo vError

fxGuardarRep = 0
vGridRep.Row = vGridRep.ActiveRow
vGridRep.col = 1

If vGridRep.Text = "" Then 'Insertar
  vGridRep.col = 2
  strSQL = "insert into afi_reportes(tipo,reporte,prefijo,seguridad) values('" _
         & Trim(vGridRep.Text) & "','"
  vGridRep.col = 3
  strSQL = strSQL & vGridRep.Text & "','"
  vGridRep.col = 4
  strSQL = strSQL & vGridRep.Text & "',"
  vGridRep.col = 5
  strSQL = strSQL & vGridRep.Value & ")"
  
  Call ConectionExecute(strSQL)

  vGridRep.col = 1
  
  strSQL = "select isnull(max(id_Rep),0) as Ultimo from afi_reportes"
  Call OpenRecordSet(rs, strSQL)
   vGridRep.Text = CStr(rs!ultimo)
  rs.Close
  
  Call Bitacora("Registra", "Reportes de Afiliación: " & vGridRep.Text)

Else 'Actualizar

 vGridRep.col = 2
 strSQL = "update afi_reportes set tipo = '" & vGridRep.Text & "',reporte = '"
 vGridRep.col = 3
 strSQL = strSQL & vGridRep.Text & "',prefijo = '"
 vGridRep.col = 4
 strSQL = strSQL & vGridRep.Text & "',seguridad = "
 vGridRep.col = 5
 strSQL = strSQL & vGridRep.Value & " Where ID_rep = "
 vGridRep.col = 1
 strSQL = strSQL & vGridRep.Text
 
 Call ConectionExecute(strSQL)

 Call Bitacora("Modifica", "Reportes de Afiliación : " & vGridRep.Text)


End If

fxGuardarRep = 1

Exit Function

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
 

End Function



Private Function fxGuardarGrpAccss() As Long

On Error GoTo vError

fxGuardarGrpAccss = 0
vGridGrpAccss.Row = vGridGrpAccss.ActiveRow
vGridGrpAccss.col = 1

If vGridGrpAccss.Text = "" Then 'Insertar
  vGridGrpAccss.col = 2
  strSQL = "insert into afi_reportes_grp(descripcion,activo) values('" _
         & Trim(vGridGrpAccss.Text) & "',"
  vGridGrpAccss.col = 3
  strSQL = strSQL & vGridGrpAccss.Value & ")"
  
  Call ConectionExecute(strSQL)

  vGridGrpAccss.col = 1
  
  strSQL = "select isnull(max(cod_grupo),0) as Ultimo from afi_reportes_grp"
  Call OpenRecordSet(rs, strSQL)
   vGridGrpAccss.Text = CStr(rs!ultimo)
  rs.Close
  
  Call Bitacora("Registra", "Reportes > Grupo de Acceso: " & vGridGrpAccss.Text)

Else 'Actualizar

 vGridGrpAccss.col = 2
 strSQL = "update afi_reportes_grp set descripcion = '" & Trim(vGridGrpAccss.Text) & "',activo = "
 vGridGrpAccss.col = 3
 strSQL = strSQL & vGridGrpAccss.Value & " where cod_grupo = "
 vGridGrpAccss.col = 1
 strSQL = strSQL & vGridGrpAccss.Text
 
 Call ConectionExecute(strSQL)

 Call Bitacora("Modifica", "Reportes > Grupo de Acceso: " & vGridGrpAccss.Text)


End If

fxGuardarGrpAccss = 1

Exit Function

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
 

End Function



Private Sub vGridGrpAccss_KeyDown(KeyCode As Integer, Shift As Integer)
Dim i As Integer

If vGridGrpAccss.ActiveCol = vGridGrpAccss.MaxCols And (KeyCode = vbKeyReturn Or KeyCode = vbKeyTab) Then
  i = fxGuardarGrpAccss
  If i = 0 Then Exit Sub
  vGridGrpAccss.Row = vGridGrpAccss.ActiveRow
  If vGridGrpAccss.MaxRows <= vGridGrpAccss.ActiveRow Then
    vGridGrpAccss.MaxRows = vGridGrpAccss.MaxRows + 1
    vGridGrpAccss.Row = vGridGrpAccss.MaxRows
  End If
End If

'Inserta Linea
If KeyCode = vbKeyInsert Then
    vGridGrpAccss.MaxRows = vGridGrpAccss.MaxRows + 1
    vGridGrpAccss.InsertRows vGridGrpAccss.ActiveRow, 1
    vGridGrpAccss.Row = vGridGrpAccss.ActiveRow
End If

End Sub



Private Sub vGridRep_KeyDown(KeyCode As Integer, Shift As Integer)
Dim i As Integer

If vGridRep.ActiveCol = vGridRep.MaxCols And (KeyCode = vbKeyReturn Or KeyCode = vbKeyTab) Then
  If txtReportes = "trick" Then
        i = fxGuardarRep
        If i = 0 Then Exit Sub
        vGridRep.Row = vGridRep.ActiveRow
        If vGridRep.MaxRows <= vGridRep.ActiveRow Then
          vGridRep.MaxRows = vGridRep.MaxRows + 1
          vGridRep.Row = vGridRep.MaxRows
        End If
  Else
    MsgBox "Proporcione la contraseña de Administrador", vbInformation
  End If

End If

'Inserta Linea
If KeyCode = vbKeyInsert Then
    vGridRep.MaxRows = vGridRep.MaxRows + 1
    vGridRep.InsertRows vGridRep.ActiveRow, 1
    vGridRep.Row = vGridRep.ActiveRow
End If

End Sub



Private Sub txtDeptCodigo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtDeptDesc.SetFocus


If KeyCode = vbKeyF4 Then
  If GLOBALES.SysASEVersion Then
        gBusquedas.Columna = "codigo"
        gBusquedas.Orden = "codigo"
        gBusquedas.Consulta = "select codigo,descripcion from uprogramatica"
  Else
        gBusquedas.Columna = "cod_departamento"
        gBusquedas.Orden = "cod_departamento"
        gBusquedas.Consulta = "select cod_departamento as 'Código',descripcion from AFDepartamentos"
  End If
  
  gBusquedas.Filtro = ""
  gBusquedas.Resultado = ""
  gBusquedas.Resultado2 = ""
  frmBusquedas.Show vbModal
  txtDeptCodigo = gBusquedas.Resultado
  txtDeptDesc = gBusquedas.Resultado2
End If

End Sub

Private Sub txtDeptDesc_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtSecCodigo.SetFocus
If KeyCode = vbKeyF4 Then
  If GLOBALES.SysASEVersion Then
        gBusquedas.Columna = "descripcion"
        gBusquedas.Orden = "descripcion"
        gBusquedas.Consulta = "select codigo,descripcion from uprogramatica"
  Else
        gBusquedas.Columna = "descripcion"
        gBusquedas.Orden = "descripcion"
        gBusquedas.Consulta = "select cod_departamento as 'Código',descripcion from AFDepartamentos"
  End If
  
  gBusquedas.Filtro = ""
  gBusquedas.Resultado = ""
  gBusquedas.Resultado2 = ""
  frmBusquedas.Show vbModal
  txtDeptCodigo = gBusquedas.Resultado
  txtDeptDesc = gBusquedas.Resultado2
End If
End Sub

Private Sub txtSecCodigo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtSecDesc.SetFocus
If KeyCode = vbKeyF4 Then
  If GLOBALES.SysASEVersion Then
        gBusquedas.Columna = "ut_codigo"
        gBusquedas.Orden = "ut_codigo"
        gBusquedas.Consulta = "select ut_codigo,ut_descripcion from UTRABAJO"
  Else
        gBusquedas.Columna = "cod_seccion"
        gBusquedas.Orden = "cod_seccion"
        gBusquedas.Consulta = "select cod_seccion as 'Código',descripcion from AFSecciones"
        gBusquedas.Filtro = " and cod_departamento = '" & txtDeptCodigo.Text & "'"
  End If
  
  frmBusquedas.Show vbModal
  frmBusquedas.Show vbModal
  txtSecCodigo = gBusquedas.Resultado
  txtSecDesc = gBusquedas.Resultado2
End If
End Sub

Private Sub txtSecDesc_KeyDown(KeyCode As Integer, Shift As Integer)
'If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then optPropiedad.Item(0).SetFocus
If KeyCode = vbKeyF4 Then
  gBusquedas.Columna = "ut_descripcion "
  gBusquedas.Orden = "ut_descripcion "
  gBusquedas.Consulta = "select ut_codigo,ut_descripcion from UTRABAJO"
  frmBusquedas.Show vbModal
  txtSecCodigo = gBusquedas.Resultado
  txtSecDesc = gBusquedas.Resultado2
End If

End Sub

Private Sub cboPromotor_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then cboSector.SetFocus

If KeyCode = vbKeyF4 Then
  gBusquedas.Resultado = ""
  gBusquedas.Resultado2 = ""
  gBusquedas.Columna = "nombre"
  gBusquedas.Orden = "nombre"
  gBusquedas.Consulta = "select id_promotor,nombre from promotores"
  gBusquedas.Filtro = " and estado = 1"
  frmBusquedas.Show vbModal
  If gBusquedas.Resultado <> "" Then
    cboPromotor.Text = (gBusquedas.Resultado2)
  End If
End If

End Sub
