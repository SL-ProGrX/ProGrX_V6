VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.Controls.v24.0.0.ocx"
Begin VB.Form frmReportesConciliacion 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Reportes de Conciliación"
   ClientHeight    =   8070
   ClientLeft      =   45
   ClientTop       =   210
   ClientWidth     =   9030
   HelpContextID   =   7004
   Icon            =   "frmReportesConciliacion.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8070
   ScaleWidth      =   9030
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.TabControl tcMain 
      Height          =   6855
      Left            =   0
      TabIndex        =   2
      Top             =   1200
      Width           =   9015
      _Version        =   1572864
      _ExtentX        =   15901
      _ExtentY        =   12091
      _StockProps     =   68
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   4
      Color           =   32
      ItemCount       =   2
      Item(0).Caption =   "Auxiliares"
      Item(0).ControlCount=   5
      Item(0).Control(0)=   "GroupBox3"
      Item(0).Control(1)=   "GroupBox1"
      Item(0).Control(2)=   "gbFiltrosCrd"
      Item(0).Control(3)=   "gbFiltros"
      Item(0).Control(4)=   "gbFiltrosFnd"
      Item(1).Caption =   "Especiales"
      Item(1).ControlCount=   2
      Item(1).Control(0)=   "gpEspecial"
      Item(1).Control(1)=   "GroupBox4"
      Begin XtremeSuiteControls.GroupBox gbFiltrosFnd 
         Height          =   1935
         Left            =   0
         TabIndex        =   41
         Top             =   3720
         Width           =   9015
         _Version        =   1572864
         _ExtentX        =   15901
         _ExtentY        =   3413
         _StockProps     =   79
         ForeColor       =   8388608
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
         Begin XtremeSuiteControls.FlatEdit txtPlan 
            Height          =   330
            Left            =   1320
            TabIndex        =   42
            ToolTipText     =   "Presione F4 para Consultar"
            Top             =   1200
            Width           =   975
            _Version        =   1572864
            _ExtentX        =   1720
            _ExtentY        =   582
            _StockProps     =   77
            ForeColor       =   0
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
            BackColor       =   16777215
            Alignment       =   2
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtPlanDesc 
            Height          =   330
            Left            =   2280
            TabIndex        =   43
            Top             =   1200
            Width           =   6015
            _Version        =   1572864
            _ExtentX        =   10610
            _ExtentY        =   582
            _StockProps     =   77
            ForeColor       =   0
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
            BackColor       =   16777215
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.ComboBox cboOperadora 
            Height          =   330
            Left            =   1320
            TabIndex        =   48
            Top             =   360
            Width           =   6975
            _Version        =   1572864
            _ExtentX        =   12303
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
         Begin XtremeSuiteControls.ComboBox cboPlanGrupo 
            Height          =   330
            Left            =   1320
            TabIndex        =   50
            Top             =   720
            Width           =   6975
            _Version        =   1572864
            _ExtentX        =   12303
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
            Caption         =   "Grupo:"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   11
            Left            =   120
            TabIndex        =   51
            Top             =   720
            Width           =   1095
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Operadora:"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   10
            Left            =   120
            TabIndex        =   49
            Top             =   360
            Width           =   1095
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Plan:"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   9
            Left            =   120
            TabIndex        =   44
            Top             =   1200
            Width           =   1095
         End
      End
      Begin XtremeSuiteControls.GroupBox gbFiltros 
         Height          =   1815
         Left            =   0
         TabIndex        =   33
         Top             =   1920
         Width           =   9015
         _Version        =   1572864
         _ExtentX        =   15901
         _ExtentY        =   3201
         _StockProps     =   79
         ForeColor       =   8388608
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
         Begin XtremeSuiteControls.ComboBox cboFiltro 
            Height          =   330
            Left            =   1320
            TabIndex        =   34
            Top             =   360
            Width           =   6975
            _Version        =   1572864
            _ExtentX        =   12303
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
         Begin XtremeSuiteControls.FlatEdit txtInstitucion 
            Height          =   330
            Left            =   1320
            TabIndex        =   36
            ToolTipText     =   "Presione F4 para Consultar"
            Top             =   720
            Width           =   975
            _Version        =   1572864
            _ExtentX        =   1720
            _ExtentY        =   582
            _StockProps     =   77
            ForeColor       =   0
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
            BackColor       =   16777215
            Alignment       =   2
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtInstitucion_Desc 
            Height          =   330
            Left            =   2280
            TabIndex        =   37
            Top             =   720
            Width           =   6015
            _Version        =   1572864
            _ExtentX        =   10610
            _ExtentY        =   582
            _StockProps     =   77
            ForeColor       =   0
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
            BackColor       =   16777215
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtCuenta 
            Height          =   330
            Left            =   1320
            TabIndex        =   45
            ToolTipText     =   "Presione F4 para Consultar"
            Top             =   1080
            Width           =   2055
            _Version        =   1572864
            _ExtentX        =   3625
            _ExtentY        =   582
            _StockProps     =   77
            ForeColor       =   0
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
            BackColor       =   16777215
            Alignment       =   2
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtCuenta_Desc 
            Height          =   330
            Left            =   3360
            TabIndex        =   46
            Top             =   1080
            Width           =   4935
            _Version        =   1572864
            _ExtentX        =   8705
            _ExtentY        =   582
            _StockProps     =   77
            ForeColor       =   0
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
            BackColor       =   16777215
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.ComboBox cboDivisa 
            Height          =   330
            Left            =   1320
            TabIndex        =   53
            Top             =   1440
            Width           =   6975
            _Version        =   1572864
            _ExtentX        =   12303
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
            Caption         =   "Divisa:"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   12
            Left            =   120
            TabIndex        =   52
            Top             =   1440
            Width           =   1095
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Cuenta:"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   7
            Left            =   120
            TabIndex        =   47
            Top             =   1080
            Width           =   1095
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
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   4
            Left            =   120
            TabIndex        =   38
            Top             =   720
            Width           =   1095
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Especiales:"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   2
            Left            =   120
            TabIndex        =   35
            Top             =   360
            Width           =   1095
         End
      End
      Begin XtremeSuiteControls.GroupBox gpEspecial 
         Height          =   4335
         Left            =   -70000
         TabIndex        =   22
         Top             =   360
         Visible         =   0   'False
         Width           =   9135
         _Version        =   1572864
         _ExtentX        =   16113
         _ExtentY        =   7646
         _StockProps     =   79
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         BorderStyle     =   2
         Begin FPSpreadADO.fpSpread vGrid_Creditos 
            Height          =   2655
            Left            =   4920
            TabIndex        =   32
            Top             =   1920
            Visible         =   0   'False
            Width           =   4095
            _Version        =   524288
            _ExtentX        =   7223
            _ExtentY        =   4683
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
            MaxCols         =   62
            SpreadDesigner  =   "frmReportesConciliacion.frx":000C
            VScrollSpecial  =   -1  'True
            VScrollSpecialType=   2
            AppearanceStyle =   1
         End
         Begin FPSpreadADO.fpSpread vGrid_Credito_Incobrables 
            Height          =   2775
            Left            =   1440
            TabIndex        =   31
            Top             =   1920
            Visible         =   0   'False
            Width           =   4095
            _Version        =   524288
            _ExtentX        =   7223
            _ExtentY        =   4895
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
            MaxCols         =   18
            SpreadDesigner  =   "frmReportesConciliacion.frx":1C2C
            VScrollSpecial  =   -1  'True
            VScrollSpecialType=   2
            AppearanceStyle =   1
         End
         Begin XtremeSuiteControls.RadioButton rbEspecial 
            Height          =   255
            Index           =   0
            Left            =   720
            TabIndex        =   24
            Top             =   360
            Width           =   8055
            _Version        =   1572864
            _ExtentX        =   14208
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Informe Integrado por Persona de Productos (Patrimonio/Ahorro/Crédito)"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9.75
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
         Begin XtremeSuiteControls.RadioButton rbEspecial 
            Height          =   255
            Index           =   1
            Left            =   720
            TabIndex        =   25
            Top             =   840
            Width           =   6135
            _Version        =   1572864
            _ExtentX        =   10821
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Base de datos para Analisis de Incobrables (KPMG-Base)"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            UseVisualStyle  =   -1  'True
            Appearance      =   17
         End
         Begin XtremeSuiteControls.RadioButton rbEspecial 
            Height          =   255
            Index           =   2
            Left            =   720
            TabIndex        =   27
            Top             =   1320
            Width           =   6135
            _Version        =   1572864
            _ExtentX        =   10821
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Auxiliar de Crédito Completo"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            UseVisualStyle  =   -1  'True
            Appearance      =   17
         End
         Begin FPSpreadADO.fpSpread vGrid_Persona_Integrado 
            Height          =   2775
            Left            =   120
            TabIndex        =   30
            Top             =   2040
            Visible         =   0   'False
            Width           =   4095
            _Version        =   524288
            _ExtentX        =   7223
            _ExtentY        =   4895
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
            MaxCols         =   93
            SpreadDesigner  =   "frmReportesConciliacion.frx":265E
            VScrollSpecial  =   -1  'True
            VScrollSpecialType=   2
            AppearanceStyle =   1
         End
      End
      Begin XtremeSuiteControls.GroupBox GroupBox3 
         Height          =   1215
         Left            =   0
         TabIndex        =   3
         Top             =   5640
         Width           =   9015
         _Version        =   1572864
         _ExtentX        =   15901
         _ExtentY        =   2143
         _StockProps     =   79
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         BorderStyle     =   1
         Begin XtremeSuiteControls.CheckBox chkSaldos 
            Height          =   375
            Left            =   2280
            TabIndex        =   4
            Top             =   360
            Width           =   3015
            _Version        =   1572864
            _ExtentX        =   5318
            _ExtentY        =   661
            _StockProps     =   79
            Caption         =   "Mostrar solo Líneas con Contenido"
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
            Value           =   1
            Alignment       =   1
         End
         Begin XtremeSuiteControls.PushButton cmdReporte 
            Height          =   615
            Left            =   6600
            TabIndex        =   5
            Top             =   360
            Width           =   1695
            _Version        =   1572864
            _ExtentX        =   2984
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
            Picture         =   "frmReportesConciliacion.frx":4C6F
         End
      End
      Begin XtremeSuiteControls.GroupBox gbFiltrosCrd 
         Height          =   1935
         Left            =   0
         TabIndex        =   6
         Top             =   3720
         Width           =   9015
         _Version        =   1572864
         _ExtentX        =   15901
         _ExtentY        =   3413
         _StockProps     =   79
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
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         BorderStyle     =   1
         Begin XtremeSuiteControls.FlatEdit txtCodigo 
            Height          =   330
            Left            =   1320
            TabIndex        =   7
            ToolTipText     =   "Presione F4 para Consultar"
            Top             =   360
            Width           =   975
            _Version        =   1572864
            _ExtentX        =   1720
            _ExtentY        =   582
            _StockProps     =   77
            ForeColor       =   0
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
            BackColor       =   16777215
            Alignment       =   2
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtCodigo_Desc 
            Height          =   330
            Left            =   2280
            TabIndex        =   8
            Top             =   360
            Width           =   6015
            _Version        =   1572864
            _ExtentX        =   10610
            _ExtentY        =   582
            _StockProps     =   77
            ForeColor       =   0
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
            BackColor       =   16777215
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtDestino 
            Height          =   330
            Left            =   1320
            TabIndex        =   9
            ToolTipText     =   "Presione F4 para Consultar"
            Top             =   1200
            Width           =   975
            _Version        =   1572864
            _ExtentX        =   1720
            _ExtentY        =   582
            _StockProps     =   77
            ForeColor       =   0
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
            BackColor       =   16777215
            Alignment       =   2
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtDestino_Desc 
            Height          =   330
            Left            =   2280
            TabIndex        =   10
            Top             =   1200
            Width           =   6015
            _Version        =   1572864
            _ExtentX        =   10610
            _ExtentY        =   582
            _StockProps     =   77
            ForeColor       =   0
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
            BackColor       =   16777215
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtRecurso 
            Height          =   330
            Left            =   1320
            TabIndex        =   11
            ToolTipText     =   "Presione F4 para Consultar"
            Top             =   1560
            Width           =   975
            _Version        =   1572864
            _ExtentX        =   1720
            _ExtentY        =   582
            _StockProps     =   77
            ForeColor       =   0
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
            BackColor       =   16777215
            Alignment       =   2
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtRecurso_Desc 
            Height          =   330
            Left            =   2280
            TabIndex        =   12
            Top             =   1560
            Width           =   6015
            _Version        =   1572864
            _ExtentX        =   10610
            _ExtentY        =   582
            _StockProps     =   77
            ForeColor       =   0
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
            BackColor       =   16777215
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.ComboBox cboGarantia 
            Height          =   330
            Left            =   1320
            TabIndex        =   40
            Top             =   720
            Width           =   6975
            _Version        =   1572864
            _ExtentX        =   12303
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
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   8
            Left            =   120
            TabIndex        =   39
            Top             =   720
            Width           =   1095
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Código:"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   3
            Left            =   120
            TabIndex        =   15
            Top             =   360
            Width           =   1095
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
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   5
            Left            =   120
            TabIndex        =   14
            Top             =   1200
            Width           =   1095
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Recurso:"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   6
            Left            =   120
            TabIndex        =   13
            Top             =   1560
            Width           =   1095
         End
      End
      Begin XtremeSuiteControls.GroupBox GroupBox1 
         Height          =   1815
         Left            =   0
         TabIndex        =   16
         Top             =   360
         Width           =   9015
         _Version        =   1572864
         _ExtentX        =   15901
         _ExtentY        =   3201
         _StockProps     =   79
         Caption         =   "Auxiliar: "
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
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         BorderStyle     =   2
         Begin XtremeSuiteControls.PushButton btnAuxiliar 
            Height          =   885
            Index           =   0
            Left            =   1320
            TabIndex        =   17
            Top             =   240
            Width           =   2415
            _Version        =   1572864
            _ExtentX        =   4260
            _ExtentY        =   1561
            _StockProps     =   79
            Caption         =   "Patrimonio"
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
            Checked         =   -1  'True
            Picture         =   "frmReportesConciliacion.frx":542B
         End
         Begin XtremeSuiteControls.PushButton btnAuxiliar 
            Height          =   885
            Index           =   1
            Left            =   3720
            TabIndex        =   18
            Top             =   240
            Width           =   2295
            _Version        =   1572864
            _ExtentX        =   4048
            _ExtentY        =   1561
            _StockProps     =   79
            Caption         =   "Fondos"
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
            Picture         =   "frmReportesConciliacion.frx":5DEB
         End
         Begin XtremeSuiteControls.PushButton btnAuxiliar 
            Height          =   885
            Index           =   2
            Left            =   6000
            TabIndex        =   19
            Top             =   240
            Width           =   2295
            _Version        =   1572864
            _ExtentX        =   4048
            _ExtentY        =   1561
            _StockProps     =   79
            Caption         =   "Crédito"
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
            Picture         =   "frmReportesConciliacion.frx":6791
         End
         Begin XtremeSuiteControls.ComboBox cboResultados 
            Height          =   330
            Left            =   1320
            TabIndex        =   20
            Top             =   1200
            Width           =   6975
            _Version        =   1572864
            _ExtentX        =   12303
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
            Caption         =   "Informe:"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   1
            Left            =   120
            TabIndex        =   21
            Top             =   1200
            Width           =   1095
         End
      End
      Begin XtremeSuiteControls.GroupBox GroupBox4 
         Height          =   3375
         Left            =   -70000
         TabIndex        =   23
         Top             =   4680
         Visible         =   0   'False
         Width           =   9135
         _Version        =   1572864
         _ExtentX        =   16113
         _ExtentY        =   5953
         _StockProps     =   79
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         BorderStyle     =   1
         Begin XtremeSuiteControls.PushButton btnExportar 
            Height          =   615
            Left            =   6600
            TabIndex        =   26
            ToolTipText     =   "Exportar a Excel"
            Top             =   1080
            Width           =   2295
            _Version        =   1572864
            _ExtentX        =   4048
            _ExtentY        =   1085
            _StockProps     =   79
            Caption         =   "Exportar a Excel"
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
            Picture         =   "frmReportesConciliacion.frx":70DC
         End
         Begin XtremeSuiteControls.ProgressBar ProgressBarX 
            Height          =   135
            Left            =   0
            TabIndex        =   28
            Top             =   840
            Visible         =   0   'False
            Width           =   9015
            _Version        =   1572864
            _ExtentX        =   15901
            _ExtentY        =   238
            _StockProps     =   93
            BackColor       =   -2147483633
            Scrolling       =   1
         End
         Begin XtremeSuiteControls.Label lblExport 
            Height          =   495
            Left            =   600
            TabIndex        =   29
            Top             =   1200
            Width           =   4815
            _Version        =   1572864
            _ExtentX        =   8493
            _ExtentY        =   873
            _StockProps     =   79
            Caption         =   "..."
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            WordWrap        =   -1  'True
         End
      End
   End
   Begin VB.Timer TimerX 
      Interval        =   10
      Left            =   0
      Top             =   0
   End
   Begin XtremeSuiteControls.ComboBox cboPeriodos 
      Height          =   312
      Left            =   1560
      TabIndex        =   1
      Top             =   480
      Width           =   4812
      _Version        =   1572864
      _ExtentX        =   8493
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
      Caption         =   "Periodo:"
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
      Index           =   0
      Left            =   1560
      TabIndex        =   0
      Top             =   240
      Width           =   1815
   End
   Begin VB.Image imgBanner 
      Height          =   1215
      Left            =   0
      Top             =   0
      Width           =   9135
   End
End
Attribute VB_Name = "frmReportesConciliacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSQL As String, rs As New ADODB.Recordset
Dim vAnio As Long, vMes As Long, vPaso As Boolean
Dim vHeaders As vGridHeaders, vTitulo As String, vEmpresa As String





Private Sub sbReporteAuxCredito()
Dim vMascara As String, i As Integer

On Error GoTo vError

Me.MousePointer = vbHourglass

vMascara = GLOBALES.gstrNiveles

'Verifica los 5 niveles x Cuenta
Do While Len(vMascara) < 5
   vMascara = vMascara & "0"
Loop

With frmContenedor.Crt
    .Reset
    .WindowState = crptMaximized
    .WindowShowGroupTree = True
    .WindowShowPrintSetupBtn = True
    .WindowShowRefreshBtn = True
    .WindowShowSearchBtn = True
    .WindowTitle = "Reportes de Conciliación"
    .Connect = glogon.ConectRPT
    
    .Formulas(0) = "SUBTITULO='PERIODO: " & UCase(cboPeriodos.Text) & "'"
    .Formulas(1) = "FECHA='" & Format(fxFechaServidor, "dd/mm/yyyy") & "'"
    .Formulas(2) = "EMPRESA='" & GLOBALES.gstrNombreEmpresa & "'"
         
         
    Select Case Mid(cboResultados.Text, 1, 2)
       Case "01", "02", "10" 'Resumen y Detalle
            
            .Formulas(3) = "MASCARA='" & vMascara & "'"
            
              Select Case Mid(cboResultados.Text, 1, 2)
                Case "01"
                    If txtInstitucion.Text = "" Then
                         .ReportFileName = SIFGlobal.fxPathReportes("Sys_AuxCreditoResumen.rpt")
                    Else
                         .ReportFileName = SIFGlobal.fxPathReportes("Sys_AuxCreditoResumenIns.rpt")
                    End If
                     .Formulas(4) = "TITULO='AUXILIAR DE CREDITO - RESUMEN DE SALDOS'"
              
                Case "02"
                
                    If txtInstitucion.Text = "" Then
                        .ReportFileName = SIFGlobal.fxPathReportes("Sys_AuxCreditoDetalle.rpt")
                    Else
                        .ReportFileName = SIFGlobal.fxPathReportes("Sys_AuxCreditoDetalleIns.rpt")
                    End If
                    .Formulas(4) = "TITULO='AUXILIAR DE CREDITO - DETALLE DE SALDOS'"
              
                Case "10" 'Resumen (Balance)
                    .ReportFileName = SIFGlobal.fxPathReportes("Sys_AuxCreditoBalanceResumen.rpt")
                    .Formulas(4) = "TITULO='AUXILIAR DE CREDITO - BALANCE RESUMIDO'"
                
              End Select
            
           
            strSQL = "select * from ase_per_historico where id_per_historico = " & cboPeriodos.ItemData(cboPeriodos.ListIndex)
            Call OpenRecordSet(rs, strSQL)
            .SelectionFormula = "{ASE_PER_CERRADOS.ANIO}=" & rs!Anio & " AND {ASE_PER_CERRADOS.MES}=" & rs!Mes _
                              & " AND ({ASE_PER_CERRADOS.ESTADO} = 'A' or {ASE_PER_CERRADOS.ESTADO} = 'C')"
                            
            rs.Close
           
            Select Case Mid(cboFiltro.Text, 1, 2)
                Case "01" 'Solo Créditos
                  .SelectionFormula = .SelectionFormula & " AND {ASE_PER_CATALOGO.RETENCION} = 'N'" _
                                    & " AND {ASE_PER_CATALOGO.POLIZA} = 'N'"
                Case "02" 'Solo Retenciones
                  .SelectionFormula = .SelectionFormula & " AND ({ASE_PER_CATALOGO.RETENCION} = 'S'" _
                                    & " OR {ASE_PER_CATALOGO.POLIZA} = 'S')"
            End Select
       
    
            If Trim(txtRecurso.Text) <> "" Then
                  .SelectionFormula = .SelectionFormula & " AND {REG_CREDITOS.COD_GRUPO} = '" _
                                    & txtRecurso.Text & "'"
            End If
            
            If Trim(txtCodigo.Text) <> "" Then
                  .SelectionFormula = .SelectionFormula & " AND {REG_CREDITOS.CODIGO} = '" _
                                    & txtCodigo.Text & "'"
            End If
            
            If Trim(txtDestino.Text) <> "" Then
                  .SelectionFormula = .SelectionFormula & " AND {REG_CREDITOS.COD_DESTINO} = '" _
                                    & txtDestino.Text & "'"
            End If
            
'            If Trim(txtCuenta.Text) <> "" Then
'                  .SelectionFormula = .SelectionFormula & " AND @Cuenta= '" _
'                                    & fxgCntCuentaFormato(False, txtCuenta.Text) & "'"
'            End If
           
    
       
       Case "03", "04" 'Saldo Inicial Negativo", "Saldo Final Negativo"
         
         If MsgBox("Desea Generar con el reporte general (SI), (NO) Se Generará con el Reporte Auxiliar", vbYesNo) = vbYes Then
            .ReportFileName = SIFGlobal.fxPathReportes("Sys_AuxCreditoDetalle.rpt")
            .Formulas(4) = "MASCARA='" & GLOBALES.gstrNiveles & "'"
         Else
            .ReportFileName = SIFGlobal.fxPathReportes("Sys_AuxCreditoSaldosNegativos.rpt")
         End If
         
         If Mid(cboResultados.Text, 1, 2) = "03" Then
            .Formulas(3) = "TITULO='AUXILIAR DE CREDITO - SALDOS INICIALES NEGATIVOS'"
            .SelectionFormula = "{ASE_PER_CERRADOS.SALDO_INICIAL} < 0"
         Else
            .Formulas(3) = "TITULO='AUXILIAR DE CREDITO - SALDOS FINALES - NEGATIVOS'"
            .SelectionFormula = "{ASE_PER_CERRADOS.SALDO_FINAL} < 0"
         End If
       
            strSQL = "select * from ase_per_historico where id_per_historico = " & cboPeriodos.ItemData(cboPeriodos.ListIndex)
            Call OpenRecordSet(rs, strSQL)
              .SelectionFormula = .SelectionFormula & " AND {ASE_PER_CERRADOS.ANIO}=" & rs!Anio & " AND {ASE_PER_CERRADOS.MES}=" & rs!Mes
            rs.Close
       
      
      Case "05" 'Metodo Contable
            
            .ReportFileName = SIFGlobal.fxPathReportes("Sys_AuxCreditoCuentas.rpt")
            
            .Formulas(3) = "USUARIO='" & UCase(glogon.Usuario) & "'"
            .Formulas(4) = "MASCARA='" & vMascara & "'"
            .Formulas(5) = "Titulo='Auxiliar : Balance Contable'"
            
            strSQL = "select * from ase_per_historico where id_per_historico = " & cboPeriodos.ItemData(cboPeriodos.ListIndex)
            Call OpenRecordSet(rs, strSQL)
                .SelectionFormula = "{ASE_PER_CUENTAS.ANIO}=" & rs!Anio & " AND {ASE_PER_CUENTAS.MES}=" & rs!Mes
            rs.Close
       
       Case "06" 'Cartera x Garantia
            .ReportFileName = SIFGlobal.fxPathReportes("Sys_AuxCreditoResumeXGarantia.rpt")
           
            strSQL = "select * from ase_per_historico where id_per_historico = " & cboPeriodos.ItemData(cboPeriodos.ListIndex)
            Call OpenRecordSet(rs, strSQL)
                .SelectionFormula = "{vSIFAuxCorteRepCredito.ANIO}=" & rs!Anio _
                                  & " AND {vSIFAuxCorteRepCredito.MES}=" & rs!Mes _
                                  & " AND {vSIFAuxCorteRepCredito.SALDO_FINAL} > 0"
            rs.Close
       
            Select Case Mid(cboFiltro.Text, 1, 2)
                Case "01" 'Solo Créditos
                  .SelectionFormula = .SelectionFormula & " AND ({vSIFAuxCorteRepCredito.RETENCION} = 'N'" _
                                    & " AND {vSIFAuxCorteRepCredito.POLIZA} = 'N')"
                Case "02" 'Solo Retenciones
                  .SelectionFormula = .SelectionFormula & " AND ({vSIFAuxCorteRepCredito.RETENCION} = 'S'" _
                                    & " OR {vSIFAuxCorteRepCredito.POLIZA} = 'S')"
            End Select
            
       
       Case "07" 'Complementario
            .Formulas(3) = "MASCARA='" & vMascara & "'"
            .Formulas(4) = "TITULO='AUXILIAR DE CREDITO - COMPLEMENTARIO'"
            
            .ReportFileName = SIFGlobal.fxPathReportes("Sys_AuxCreditoComplementario.rpt")
           
            strSQL = "select * from ase_per_historico where id_per_historico = " & cboPeriodos.ItemData(cboPeriodos.ListIndex)
            Call OpenRecordSet(rs, strSQL)
            .SelectionFormula = "{ASE_PER_CERRADOS.ANIO}=" & rs!Anio & " AND {ASE_PER_CERRADOS.MES}=" & rs!Mes _
                              & " AND ({ASE_PER_CERRADOS.ESTADO} = 'A' or {ASE_PER_CERRADOS.ESTADO} = 'C')"
                            
            rs.Close
           
            Select Case Mid(cboFiltro.Text, 1, 2)
                Case "01" 'Solo Créditos
                  .SelectionFormula = .SelectionFormula & " AND {ASE_PER_CATALOGO.RETENCION} = 'N'" _
                                    & " AND {ASE_PER_CATALOGO.POLIZA} = 'N'"
                Case "02" 'Solo Retenciones
                  .SelectionFormula = .SelectionFormula & " AND ({ASE_PER_CATALOGO.RETENCION} = 'S'" _
                                    & " OR {ASE_PER_CATALOGO.POLIZA} = 'S')"
            End Select
       
      Case "08", "09" 'Producto Acumulado
            
            If Mid(cboResultados.Text, 1, 2) = "08" Then
                .ReportFileName = SIFGlobal.fxPathReportes("Sys_AuxCreditoProdAcum.rpt")
                .Formulas(3) = "USUARIO='" & UCase(glogon.Usuario) & "'"
                .Formulas(4) = "Titulo='Auxiliar: Producto Acumulado [Detalle]'"
            Else
                .ReportFileName = SIFGlobal.fxPathReportes("Sys_AuxCreditoProdAcumRsm.rpt")
                .Formulas(3) = "USUARIO='" & UCase(glogon.Usuario) & "'"
                .Formulas(4) = "Titulo='Auxiliar: Producto Acumulado [Resumen]'"
            End If
            
            
            strSQL = "select * from ase_per_historico where id_per_historico = " & cboPeriodos.ItemData(cboPeriodos.ListIndex)
            Call OpenRecordSet(rs, strSQL)
                .StoredProcParam(0) = rs!Anio
                .StoredProcParam(1) = rs!Mes
                .SelectionFormula = ""
            rs.Close

      
      
      Case "12", "13" 'Ingreso x Int. Cobrados por Adelantado (Resumen)
            
            If Mid(cboResultados.Text, 1, 2) = "12" Then
                .ReportFileName = SIFGlobal.fxPathReportes("Sys_AuxCreditoIntCbrAdelandoRsm.rpt")
                .Formulas(3) = "USUARIO='" & UCase(glogon.Usuario) & "'"
                .Formulas(4) = "Titulo='Ingresos x Intereses Cobrados por Adelantado (Resumen)'"
            Else
                .ReportFileName = SIFGlobal.fxPathReportes("Sys_AuxCreditoIntCbrAdelando.rpt")
                .Formulas(3) = "USUARIO='" & UCase(glogon.Usuario) & "'"
                .Formulas(4) = "Titulo='Ingresos x Intereses Cobrados por Adelantado (Detalle)'"
            End If
            
            
            strSQL = "select * from ase_per_historico where id_per_historico = " & cboPeriodos.ItemData(cboPeriodos.ListIndex)
            Call OpenRecordSet(rs, strSQL)
                .StoredProcParam(0) = rs!Anio
                .StoredProcParam(1) = rs!Mes
                .StoredProcParam(2) = glogon.Usuario
                .StoredProcParam(3) = 0
                .SelectionFormula = ""
            rs.Close
      
      
      Case "14", "15" 'Producto Acumulado en Suspenso
            
            If Mid(cboResultados.Text, 1, 2) = "14" Then
                .ReportFileName = SIFGlobal.fxPathReportes("Sys_AuxCreditoProdAcumSuspenso.rpt")
                .Formulas(3) = "USUARIO='" & UCase(glogon.Usuario) & "'"
                .Formulas(4) = "Titulo='Auxiliar: Producto Acumulado en Suspenso [Detalle]'"
            Else
                .ReportFileName = SIFGlobal.fxPathReportes("Sys_AuxCreditoProdAcumSuspensoRsm.rpt")
                .Formulas(3) = "USUARIO='" & UCase(glogon.Usuario) & "'"
                .Formulas(4) = "Titulo='Auxiliar: Producto Acumulado en Suspenso [Resumen]'"
            End If
            
            
            strSQL = "select * from ase_per_historico where id_per_historico = " & cboPeriodos.ItemData(cboPeriodos.ListIndex)
            Call OpenRecordSet(rs, strSQL)
                .StoredProcParam(0) = rs!Anio
                .StoredProcParam(1) = rs!Mes
                .SelectionFormula = ""
            rs.Close
  
  
      Case "16", "17" 'Gasto Diferido
            
            If Mid(cboResultados.Text, 1, 2) = "16" Then
                .ReportFileName = SIFGlobal.fxPathReportes("Sys_AuxCreditoGastoDiferido.rpt")
                .Formulas(3) = "USUARIO='" & UCase(glogon.Usuario) & "'"
                .Formulas(4) = "Titulo='Cargos / Gastos Diferidos (Detalle)'"
            Else
                .ReportFileName = SIFGlobal.fxPathReportes("Sys_AuxCreditoGastoDiferidoRsm.rpt")
                .Formulas(3) = "USUARIO='" & UCase(glogon.Usuario) & "'"
                .Formulas(4) = "Titulo='Cargos / Gastos Diferidos (Resumen)'"
            End If
            
            
            strSQL = "select * from ase_per_historico where id_per_historico = " & cboPeriodos.ItemData(cboPeriodos.ListIndex)
            Call OpenRecordSet(rs, strSQL)
                .StoredProcParam(0) = rs!Anio
                .StoredProcParam(1) = rs!Mes
                .StoredProcParam(2) = glogon.Usuario
                .StoredProcParam(3) = 0
                .SelectionFormula = ""
            rs.Close
      
    End Select
    .PrintReport
End With


Me.MousePointer = vbDefault

Exit Sub

vError:

Me.MousePointer = vbDefault
MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub sbReporteAuxPatrimonio()
Dim strSQL As String, rs As New ADODB.Recordset
Dim i As Byte, vMascara As String

On Error GoTo vError

Me.MousePointer = vbHourglass

vMascara = GLOBALES.gstrNiveles

'Verifica los 5 niveles x Cuenta
Do While Len(vMascara) < 5
   vMascara = vMascara & "0"
Loop

With frmContenedor.Crt
    .Reset
    .WindowState = crptMaximized
    .WindowShowGroupTree = True
    .WindowShowPrintSetupBtn = True
    .WindowShowRefreshBtn = True
    .WindowShowSearchBtn = True
    .WindowTitle = "Reportes de Conciliación"
    
    .Connect = glogon.ConectRPT
    
    .Formulas(0) = "SUBTITULO='PERIODO: " & UCase(cboPeriodos.Text) & "'"
    .Formulas(1) = "FECHA='" & Format(fxFechaServidor, "dd/mm/yyyy") & "'"
    .Formulas(2) = "EMPRESA='" & GLOBALES.gstrNombreEmpresa & "'"
    .Formulas(3) = "USUARIO='" & UCase(glogon.Usuario) & "'"
    
    Select Case Mid(cboResultados.Text, 1, 2)
       Case "01" 'Resumen
          If Mid(cboFiltro.Text, 1, 1) = "P" Then
              .ReportFileName = SIFGlobal.fxPathReportes("Sys_AuxAportesResumenFCI.rpt")
          Else
              .ReportFileName = SIFGlobal.fxPathReportes("Sys_AuxAportesResumen.rpt")
          End If
       
       
            strSQL = "select * from ase_per_historico where id_per_historico = " & cboPeriodos.ItemData(cboPeriodos.ListIndex)
            Call OpenRecordSet(rs, strSQL)
                .SelectionFormula = "{ASE_PER_APORTES.ANIO}=" & rs!Anio & " AND {ASE_PER_APORTES.MES}=" & rs!Mes
            rs.Close
            
            If txtInstitucion.Text <> "" Then
              .SelectionFormula = .SelectionFormula & " AND {SOCIOS.COD_INSTITUCION} = " & txtInstitucion.Text
            End If
            
            If chkSaldos.Value = vbChecked Then
              .SelectionFormula = .SelectionFormula & " AND ({ASE_PER_APORTES.APORTE} + {ASE_PER_APORTES.AHORRO} + {ASE_PER_APORTES.CAPITALIZA}+  {ASE_PER_APORTES.CUSTODIA} + {ASE_PER_APORTES.EXTRA}) <> 0"
            End If
       
       Case "02" 'Detalle
          If Mid(cboFiltro.Text, 1, 1) = "P" Then
              .ReportFileName = SIFGlobal.fxPathReportes("Sys_AuxAportesDetalleFCI.rpt")
          Else
              .ReportFileName = SIFGlobal.fxPathReportes("Sys_AuxAportesDetalle.rpt")
          End If
            
            strSQL = "select * from ase_per_historico where id_per_historico = " & cboPeriodos.ItemData(cboPeriodos.ListIndex)
            Call OpenRecordSet(rs, strSQL)
                .SelectionFormula = "{ASE_PER_APORTES.ANIO}=" & rs!Anio & " AND {ASE_PER_APORTES.MES}=" & rs!Mes
            rs.Close
            
            If txtInstitucion.Text <> "" Then
              .SelectionFormula = .SelectionFormula & " AND {SOCIOS.COD_INSTITUCION} = " & txtInstitucion.Text
            End If
            
            
            If chkSaldos.Value = vbChecked Then
              .SelectionFormula = .SelectionFormula & " AND ({ASE_PER_APORTES.APORTE} + {ASE_PER_APORTES.AHORRO} + {ASE_PER_APORTES.CAPITALIZA}+  {ASE_PER_APORTES.CUSTODIA} + {ASE_PER_APORTES.EXTRA}) <> 0"
            End If
            
       
       Case "03" 'Resumen x Categoria de la Persona
          If Mid(cboFiltro.Text, 1, 1) = "P" Then
              .ReportFileName = SIFGlobal.fxPathReportes("Sys_AuxAportesRsmCategoriaFCI.rpt")
          Else
              .ReportFileName = SIFGlobal.fxPathReportes("Sys_AuxAportesRsmCategoria.rpt")
          End If
            
            strSQL = "select * from ase_per_historico where id_per_historico = " & cboPeriodos.ItemData(cboPeriodos.ListIndex)
            Call OpenRecordSet(rs, strSQL)
                .SelectionFormula = "{ASE_PER_APORTES.ANIO}=" & rs!Anio & " AND {ASE_PER_APORTES.MES}=" & rs!Mes
            rs.Close
            
            If txtInstitucion.Text <> "" Then
              .SelectionFormula = .SelectionFormula & " AND {SOCIOS.COD_INSTITUCION} = " & txtInstitucion.Text
            End If
      
      
            If chkSaldos.Value = vbChecked Then
              .SelectionFormula = .SelectionFormula & " AND ({ASE_PER_APORTES.APORTE} + {ASE_PER_APORTES.AHORRO} + {ASE_PER_APORTES.CAPITALIZA}+  {ASE_PER_APORTES.CUSTODIA} + {ASE_PER_APORTES.EXTRA}) <> 0"
            End If
      
      Case "04" 'Metodo Contable
            
            .ReportFileName = SIFGlobal.fxPathReportes("Sys_AuxAportesCuentas.rpt")
            
            .Formulas(3) = "USUARIO='" & UCase(glogon.Usuario) & "'"
            .Formulas(4) = "MASCARA='" & vMascara & "'"
            .Formulas(5) = "Titulo='Auxiliar : Balance Contable'"
            
            strSQL = "select * from ase_per_historico where id_per_historico = " & cboPeriodos.ItemData(cboPeriodos.ListIndex)
            Call OpenRecordSet(rs, strSQL)
                'La formula utiliza el alias de la tabla no la real que es : ASE_PER_CUENTAS_PAT
                .SelectionFormula = "{ASE_PER_CUENTAS.ANIO}=" & rs!Anio & " AND {ASE_PER_CUENTAS.MES}=" & rs!Mes
            rs.Close
            
            
    End Select
    
    .Action = 1

    
End With

Me.MousePointer = vbDefault

Exit Sub

vError:
    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub sbReporteAuxFondos()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

Me.MousePointer = vbHourglass

With frmContenedor.Crt
    .Reset
    .WindowState = crptMaximized
    .WindowShowGroupTree = True
    .WindowShowPrintSetupBtn = True
    .WindowShowRefreshBtn = True
    .WindowShowSearchBtn = True
    .WindowTitle = "Reportes de Conciliación"
    
    .Connect = glogon.ConectRPT
    
    .Formulas(0) = "SUBTITULO='PERIODO: " & UCase(cboPeriodos.Text) & " / FILTRO : " & UCase(cboFiltro.Text) & "'"
    .Formulas(1) = "FECHA='" & Format(fxFechaServidor, "dd/mm/yyyy") & "'"
    .Formulas(2) = "EMPRESA='" & GLOBALES.gstrNombreEmpresa & "'"
    .Formulas(3) = "USUARIO='" & UCase(glogon.Usuario) & "'"
    
         
    Select Case Mid(cboResultados.Text, 1, 2)
       Case "01" 'Resumen
           .ReportFileName = SIFGlobal.fxPathReportes("Sys_AuxFondoResumen.rpt")
       Case "02" 'Detalle
           .ReportFileName = SIFGlobal.fxPathReportes("Sys_AuxFondoDetalle.rpt")
    End Select

      .SelectionFormula = "{FND_PER_CERRADOS.ANIO}=" & Mid(cboPeriodos.Text, 1, 4) & " AND {FND_PER_CERRADOS.MES} = " & fxConvierteMES(Mid(cboPeriodos.Text, 8, 10))
    
    Select Case Mid(cboFiltro.Text, 1, 2)
      Case "01" 'Activos
        .SelectionFormula = .SelectionFormula & " AND {FND_PER_CERRADOS.ESTADO} = 'A'"
      Case "02" 'Liquidados
        .SelectionFormula = .SelectionFormula & " AND {FND_PER_CERRADOS.ESTADO} = 'L'"
    End Select


    .PrintReport
End With

vError:

Me.MousePointer = vbDefault

End Sub

Private Sub btnAuxiliar_Click(Index As Integer)
Dim i As Integer

cboResultados.Clear
cboFiltro.Clear

For i = 0 To btnAuxiliar.Count - 1
    If i = Index Then
       btnAuxiliar.Item(i).Checked = True
    Else
       btnAuxiliar.Item(i).Checked = False
    End If
Next i

gbFiltrosFnd.Visible = False
gbFiltrosCrd.Visible = False

Select Case Index
  Case 0 'Patrimonio
       gbFiltrosFnd.Visible = True
       
       cboResultados.AddItem "01 - Aportes (Resumen)"
       cboResultados.AddItem "02 - Aportes (Detalle)"
       cboResultados.AddItem "03 - Aportes (Rsm x Categoria)"
       cboResultados.AddItem "04 - Metodo Contable"
       
       cboResultados.Text = "01 - Aportes (Resumen)"
       
       cboFiltro.AddItem "Patrimonio + FCI"
       cboFiltro.AddItem "Solo Patrimonio"
       cboFiltro.Text = "Solo Patrimonio"
  
  Case 1 'Fondos
       gbFiltrosFnd.Visible = True
       
       cboResultados.AddItem "01 - Fondos (Resumen)"
       cboResultados.AddItem "02 - Fondos (Detalle)"
       cboResultados.Text = "01 - Fondos (Resumen)"
       
       cboFiltro.AddItem "00 - Todos"
       cboFiltro.AddItem "01 - Contratos Activos"
       cboFiltro.AddItem "02 - Liquidados Activos"
       cboFiltro.Text = "00 - Todos"
  
  Case 2 'Creditos
       gbFiltrosCrd.Visible = True
       
       cboResultados.AddItem "01 - Créditos (Resumen)"
       cboResultados.AddItem "02 - Créditos (Detalle)"
       cboResultados.AddItem "03 - Saldo Inicial Negativo"
       cboResultados.AddItem "04 - Saldo Final Negativo"
       cboResultados.AddItem "05 - Metodo Contable"
       cboResultados.AddItem "06 - Cartera x Garantia"
       cboResultados.AddItem "07 - Complementario"
       cboResultados.AddItem "08 - Producto Acumulado"
       cboResultados.AddItem "09 - Producto Acumulado [Resumen]"
       
       
       cboResultados.AddItem "10 - Créditos Balance (Resumen)"
       cboResultados.AddItem "11 - Créditos Balance (Detalle)"
       
       cboResultados.AddItem "12 - Int.Cbr.Adelantado [Resumen]"
       cboResultados.AddItem "13 - Int.Cbr.Adelantado [Detalle]"
       
       cboResultados.AddItem "14 - Prod. Acum. Suspenso"
       cboResultados.AddItem "15 - Prod. Acum. Suspenso [Resumen]"
       
       cboResultados.AddItem "16 - Gastos/Cargos Diferidos"
       cboResultados.AddItem "17 - Gastos/Cargos Diferidos [Resumen]"
       
       
       cboResultados.Text = "01 - Créditos (Resumen)"
       
       cboFiltro.AddItem "00 - Todos"
       cboFiltro.AddItem "01 - Créditos"
       cboFiltro.AddItem "02 - Retenciones"
       cboFiltro.Text = "00 - Todos"

End Select


End Sub

Private Sub sbCargaGrid_Local(vGrid As Object, vGridMaxCol As Integer, strSQL As String, Optional vBorra As Boolean = True)
Dim rs As New ADODB.Recordset, i As Integer

On Error GoTo vErrorLoad

If vBorra Then
    vGrid.MaxCols = vGridMaxCol
    vGrid.MaxRows = 1
    vGrid.Row = vGrid.MaxRows
    For i = 1 To vGrid.MaxCols
     vGrid.Col = i
     vGrid.Text = ""
    Next i
End If

Call OpenRecordSet(rs, strSQL, 0)
  
vGrid.MaxRows = 1

ProgressBarX.Max = rs.RecordCount + 1

Do While Not rs.EOF
  vGrid.Row = vGrid.MaxRows
  
  ProgressBarX.Value = vGrid.MaxRows
  
  For i = 1 To vGrid.MaxCols
  
    vGrid.Col = i
    Select Case vGrid.CellType
        Case CellTypeDate
            vGrid.Text = Format(IIf(IsNull(rs.Fields(i - 1).Value), "", rs.Fields(i - 1)), "dd/mm/yyyy")
        Case Else
            vGrid.Text = CStr(IIf(IsNull(rs.Fields(i - 1).Value), "", rs.Fields(i - 1)))
    End Select
  

  Next i
  vGrid.MaxRows = vGrid.MaxRows + 1
  
  
  
  rs.MoveNext
Loop
rs.Close

Exit Sub

vErrorLoad:
  MsgBox fxSys_Error_Handler(Err.Description), vbExclamation
 
  
End Sub



Private Sub sbEspecial_Persona_Productos()


strSQL = "exec spSys_Aux_Core_Full " & vAnio & ", " & vMes

Call sbCargaGrid_Local(vGrid_Persona_Integrado, 93, strSQL)

'Bitacora
Call Bitacora("Exporta", "Auxiliar Especial: Integral por Persona-Productos Consolidado: " & cboPeriodos.Text)

'Variables del Exporte
vHeaders.Columnas = vGrid_Persona_Integrado.MaxCols
vTitulo = "ProGrX_Aux_Especial_" & vEmpresa & "_" & cboPeriodos.Text & "_Integral_Persona"
    
    vHeaders.Headers(1) = "Corte"
    vHeaders.Headers(2) = "Tipo Id"
    vHeaders.Headers(3) = "Identificación"
    vHeaders.Headers(4) = "Id Alterno"
    vHeaders.Headers(5) = "Nombre"
    vHeaders.Headers(6) = "Fec.Ingreso"
    vHeaders.Headers(7) = "Fec.Nacimiento"
    vHeaders.Headers(8) = "Estado Persona"
    vHeaders.Headers(9) = "Edad"
    vHeaders.Headers(10) = "Membresía"
    vHeaders.Headers(11) = "Género"
    vHeaders.Headers(12) = "Estado Laboral"
    vHeaders.Headers(13) = "Estado Civil"
    vHeaders.Headers(14) = "Id Depart."
    vHeaders.Headers(15) = "Departamento"
    vHeaders.Headers(16) = "Id Sección"
    vHeaders.Headers(17) = "Sección"
    vHeaders.Headers(18) = "Id Oficina"
    vHeaders.Headers(19) = "Oficina"
    vHeaders.Headers(20) = "Institución/Empresa"
    vHeaders.Headers(21) = "Aporte Obrero"
    vHeaders.Headers(22) = "Aporte Patronal"
    vHeaders.Headers(23) = "Capitalización"
    vHeaders.Headers(24) = "Aporte En Custodia"
    vHeaders.Headers(25) = "Primer Deduc."
    vHeaders.Headers(26) = "Profesión"
    vHeaders.Headers(27) = "Sector"
    vHeaders.Headers(28) = "Id Provincia"
    vHeaders.Headers(29) = "Provincia"
    vHeaders.Headers(30) = "Id Cantón"
    vHeaders.Headers(31) = "Cantón"
    vHeaders.Headers(32) = "Id Distrito"
    vHeaders.Headers(33) = "Distrito"
    vHeaders.Headers(34) = "Ejecutivo"
    vHeaders.Headers(35) = "Tipo Ejecutivo"
    vHeaders.Headers(36) = "Apl.Doble Deduc."
    vHeaders.Headers(37) = "Posee Propiedades"
    vHeaders.Headers(38) = "NO Apl.Aportes"
    vHeaders.Headers(39) = "Razón Social"
    vHeaders.Headers(40) = "Tipo Sociedad"
    vHeaders.Headers(41) = "Actividad Económica"
    vHeaders.Headers(42) = "Email No.1"
    vHeaders.Headers(43) = "Email No.2"
    vHeaders.Headers(44) = "Tel. Habitación"
    vHeaders.Headers(45) = "Tel. Trabajo"
    vHeaders.Headers(46) = "Tel. Móvil"
    vHeaders.Headers(47) = "Crd.Gar.S/Ahorros"
    vHeaders.Headers(48) = "Crd.Gar.Excedentes"
    vHeaders.Headers(49) = "Crd.Gar.Fiduciarios"
    vHeaders.Headers(50) = "Crd.Gar.Pagaré"
    vHeaders.Headers(51) = "Crd.Gar.Puente"
    vHeaders.Headers(52) = "Crd.Gar.NoAplica"
    vHeaders.Headers(53) = "Crd.Gar.Comercial"
    vHeaders.Headers(54) = "Crd.Gar.Acciones"
    vHeaders.Headers(55) = "Crd.Gar.BackToBack"
    vHeaders.Headers(56) = "Crd.Gar.Hipotecario"
    vHeaders.Headers(57) = "Total Créditos"
    vHeaders.Headers(58) = "Crd.Operaciones"
    vHeaders.Headers(59) = "Crd.Cuotas"
    vHeaders.Headers(60) = "Crd.Pólizas"
    vHeaders.Headers(61) = "Crd.Saldos"
    vHeaders.Headers(62) = "Recaudos Operaciones"
    vHeaders.Headers(63) = "Recaudos Cuotas"
    vHeaders.Headers(64) = "Recaudos Saldo"
    vHeaders.Headers(65) = "Antiguedad Saldos"
    vHeaders.Headers(66) = "Ahorro a la Vista"
    vHeaders.Headers(67) = "Ahorro Navideño"
    vHeaders.Headers(68) = "Ahorro Escolar"
    vHeaders.Headers(69) = "Ahorro Marchamo"
    vHeaders.Headers(70) = "Ahorro Sal.Escolar"
    vHeaders.Headers(71) = "Ahorro a Plazo 2m"
    vHeaders.Headers(72) = "Fondos Excedentes"
    vHeaders.Headers(73) = "Fondos Exc.Apl.Custodia"
    vHeaders.Headers(74) = "Fondos Otros"
    vHeaders.Headers(75) = "Planes Contratos"
    vHeaders.Headers(76) = "Planes Cuotas"
    vHeaders.Headers(77) = "Planes TOTAL"
    vHeaders.Headers(78) = "Clasificación a Hoy"
    vHeaders.Headers(79) = "Id Deductora"
    vHeaders.Headers(80) = "Deductora Desc."
    vHeaders.Headers(81) = "Deductora Desc.Corta"
    vHeaders.Headers(82) = "Análisis Garantía"
    vHeaders.Headers(83) = "Análisis Endeudamiento"
    vHeaders.Headers(84) = "Análisis Historial Pago"
    vHeaders.Headers(85) = "Análisis Morosidad"
    vHeaders.Headers(86) = "Análisis Liquidez"
    vHeaders.Headers(87) = "Análisis Fecha"
    vHeaders.Headers(88) = "Salario Devengado"
    vHeaders.Headers(89) = "Salario Líquido"
    vHeaders.Headers(90) = "Liquidez Simple"
    vHeaders.Headers(91) = "Liquidez c/Fianzas"
    vHeaders.Headers(92) = "Fecha Liquidación"
    vHeaders.Headers(93) = "Re-Ingreso?"
    
    
   Call sbSIFGridExportar(vGrid_Persona_Integrado, vHeaders, vTitulo)



End Sub


Private Sub sbEspecial_Creditos_Incobrables_KPMG()


strSQL = "exec spSys_Aux_Creditos_Base_Incobrables " & vAnio & ", " & vMes

Call sbCargaGrid(vGrid_Credito_Incobrables, 18, strSQL)

'Bitacora
Call Bitacora("Exporta", "Auxiliar Especial: Creditos Incobrables KMPG: " & cboPeriodos.Text)

'Variables del Exporte
vHeaders.Columnas = vGrid_Credito_Incobrables.MaxCols
vTitulo = "ProGrX_Aux_Especial_" & vEmpresa & "_" & cboPeriodos.Text & "_Creditos_Incobrables_KPMG"
    
    vHeaders.Headers(1) = "Corte"
    vHeaders.Headers(2) = "Fec.Formaliza"
    vHeaders.Headers(3) = "Fec.Termina"
    vHeaders.Headers(4) = "Fec.Ult.Cuota"
    vHeaders.Headers(5) = "No. Documento"
    vHeaders.Headers(6) = "Id Alterno"
    vHeaders.Headers(7) = "Monto"
    vHeaders.Headers(8) = "Saldo"
    vHeaders.Headers(9) = "Periodicidad"
    vHeaders.Headers(10) = "Tasa"
    vHeaders.Headers(11) = "Linea"
    vHeaders.Headers(12) = "Status"
    vHeaders.Headers(13) = "Código"
    vHeaders.Headers(14) = "Descripción"
    vHeaders.Headers(15) = "Divisa"
    vHeaders.Headers(16) = "Cedula"
    vHeaders.Headers(17) = "Nombre"
    vHeaders.Headers(18) = "Tipo Cambio"
   
    
   Call sbSIFGridExportar(vGrid_Credito_Incobrables, vHeaders, vTitulo)



End Sub


Private Sub sbEspecial_Creditos_Completo()


strSQL = "exec spSys_Aux_Creditos_Completo " & vAnio & ", " & vMes

Call sbCargaGrid(vGrid_Creditos, 62, strSQL)

'Bitacora
Call Bitacora("Exporta", "Auxiliar Especial: Creditos Completo: " & cboPeriodos.Text)

'Variables del Exporte
vHeaders.Columnas = vGrid_Creditos.MaxCols
vTitulo = "ProGrX_Aux_Especial_" & vEmpresa & "_" & cboPeriodos.Text & "_Creditos_Completo"
    
    vHeaders.Headers(1) = "Corte"
    vHeaders.Headers(2) = "No. Documento"
    vHeaders.Headers(3) = "Id Alterno"
    vHeaders.Headers(4) = "Cédula"
    vHeaders.Headers(5) = "Nombre"
    vHeaders.Headers(6) = "Divisa"
    vHeaders.Headers(7) = "No.Operación"
    vHeaders.Headers(8) = "L.Código"
    vHeaders.Headers(9) = "L.Descripción"
    vHeaders.Headers(10) = "Garantía"
    vHeaders.Headers(11) = "Destino"
    vHeaders.Headers(12) = "Recurso"
    vHeaders.Headers(13) = "Monto Solicitado"
    vHeaders.Headers(14) = "Tipo Salida"
    vHeaders.Headers(15) = "Monto Girado"
    vHeaders.Headers(16) = "Monto Refinanciado"
    vHeaders.Headers(17) = "Cargos Formaliza"
    vHeaders.Headers(18) = "Monto Desembolsos"
    vHeaders.Headers(19) = "Monto Abono Recaudos"
    vHeaders.Headers(20) = "Otros Rebajos"
    vHeaders.Headers(21) = "Monto Crédito"
    vHeaders.Headers(22) = "Saldo"
    vHeaders.Headers(23) = "Saldo Divisa Local"
    vHeaders.Headers(24) = "Tipo Cambio"
    vHeaders.Headers(25) = "Tasa Original"
    vHeaders.Headers(26) = "Tasa"
    vHeaders.Headers(27) = "Plazo"
    vHeaders.Headers(28) = "Cuota"
    vHeaders.Headers(29) = "Cuota Divisa Local"
    vHeaders.Headers(30) = "Cuota Polizas"
    vHeaders.Headers(31) = "Estado"
    vHeaders.Headers(32) = "Proceso"
    vHeaders.Headers(33) = "Op.Ex.Soc."
    vHeaders.Headers(34) = "Antiguedad"
    vHeaders.Headers(35) = "Mora Cta. + Antigua"
    vHeaders.Headers(36) = "Mora No. Cuotas"
    vHeaders.Headers(37) = "Mora Cargos"
    vHeaders.Headers(38) = "Mora Int.Cor."
    vHeaders.Headers(39) = "Mora Int.Mor."
    vHeaders.Headers(40) = "Mora.Principal"
    vHeaders.Headers(41) = "Fec.Formaliza"
    vHeaders.Headers(42) = "Fec.Pri.Deduc"
    vHeaders.Headers(43) = "Fec.Ult.Cta"
    vHeaders.Headers(44) = "Fec.Termina"
    vHeaders.Headers(45) = "Vencimiento Cartera"
    vHeaders.Headers(46) = "VC R0 Vencidas"
    vHeaders.Headers(47) = "VC R1 0 - 3 Meses"
    vHeaders.Headers(48) = "VC R2 3 - 6 Meses"
    vHeaders.Headers(49) = "VC R3 6 - 12 Meses"
    vHeaders.Headers(50) = "VC R4 Mayor 1 año"
    vHeaders.Headers(51) = "PA Dias"
    vHeaders.Headers(52) = "PA Inicio"
    vHeaders.Headers(53) = "PA Corte"
    vHeaders.Headers(54) = "PA Monto"
    vHeaders.Headers(55) = "Cta Cartera"
    vHeaders.Headers(56) = "Cta Cart. Desc"
    vHeaders.Headers(57) = "Cta Prod.Acum."
    vHeaders.Headers(58) = "Cta Prod. Acum. Desc"
    vHeaders.Headers(59) = "Cta PA Resultados"
    vHeaders.Headers(60) = "Cta PA Resultados Desc"
    vHeaders.Headers(61) = "TITA"
    vHeaders.Headers(62) = "TEA"
  
    
   Call sbSIFGridExportar(vGrid_Creditos, vHeaders, vTitulo)

End Sub




Private Sub btnExportar_Click()

On Error GoTo vError


ProgressBarX.Visible = True

lblExport.Caption = "Cargando Información...Espere!"
DoEvents

Me.MousePointer = vbHourglass

Select Case True
    Case rbEspecial(0).Value 'Integrado por Persona
        Call sbEspecial_Persona_Productos
    Case rbEspecial(1).Value 'BD Incobrables (KPMG-Base)
        Call sbEspecial_Creditos_Incobrables_KPMG
    Case rbEspecial(2).Value 'Credito - Completos
        Call sbEspecial_Creditos_Completo
End Select


Me.MousePointer = vbDefault

lblExport.Caption = ""
ProgressBarX.Visible = False

Exit Sub

vError:
  ProgressBarX.Visible = False
  lblExport.Caption = ""
  
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub cboPeriodos_Click()

            
If vPaso Then Exit Sub

On Error GoTo vError

strSQL = "select * from ase_per_historico where id_per_historico = " & cboPeriodos.ItemData(cboPeriodos.ListIndex)
Call OpenRecordSet(rs, strSQL)

    vAnio = rs!Anio
    vMes = rs!Mes

rs.Close
            
Exit Sub

vError:

End Sub

Private Sub cmdReporte_Click()

If cboPeriodos.Text = "" Then Exit Sub

Select Case True
  Case btnAuxiliar.Item(0).Checked  'Patrimonio
    Call sbReporteAuxPatrimonio
  Case btnAuxiliar.Item(1).Checked 'Fondos
    Call sbReporteAuxFondos
  Case btnAuxiliar.Item(2).Checked  'Creditos
    Call sbReporteAuxCredito
End Select

End Sub

Private Sub Form_Load()

Set imgBanner.Picture = frmContenedor.imgBanner_Reportes.Picture

tcMain.Item(0).Selected = True

Call Formularios(Me)
Call RefrescaTags(Me)



End Sub



Private Sub sbInicial()


strSQL = "select PAG_NOMCORTO from SIF_EMPRESA"
Call OpenRecordSet(rs, strSQL)
vEmpresa = Trim(rs!PAG_NOMCORTO & "")
rs.Close

strSQL = "select * from ase_per_historico order by anio desc,mes desc"
Call OpenRecordSet(rs, strSQL)

vPaso = True

cboPeriodos.Clear
Do While Not rs.EOF
 cboPeriodos.AddItem rs!Anio & " - " & fxConvierteMES(rs!Mes)
 cboPeriodos.ItemData(cboPeriodos.ListCount - 1) = CStr(rs!id_per_historico)
 rs.MoveNext
Loop

If rs.RecordCount > 0 Then
 rs.MoveFirst
 Call sbCboAsignaDato(cboPeriodos, rs!Anio & " - " & fxConvierteMES(rs!Mes), True, rs!id_per_historico)
End If
rs.Close


strSQL = "select GARANTIA as 'IdX', DESCRIPCION as 'ItmX'" _
       & " From CRD_GARANTIA_TIPOS"
Call sbCbo_Llena_New(cboGarantia, strSQL, True, True)


strSQL = "select rtrim(cod_Divisa) as 'IdX', rtrim(Descripcion) as 'ItmX'" _
        & " from CntX_Divisas" _
        & " Where cod_Contabilidad = " & GLOBALES.gEnlace
Call sbCbo_Llena_New(cboDivisa, strSQL, True, True)

strSQL = "select rtrim(descripcion) as 'ItmX',cod_operadora as 'Idx' from FND_Operadoras"
Call sbCbo_Llena_New(cboOperadora, strSQL, False, True)

strSQL = "select rtrim(cod_grupo) as 'IdX', rtrim(descripcion) as ItmX from fnd_grupos"
Call sbCbo_Llena_New(cboPlanGrupo, strSQL, True, True)


vPaso = False

Call cboPeriodos_Click

Call OptX_Click(0)

End Sub



Private Sub TimerX_Timer()
TimerX.Interval = 0
TimerX.Enabled = False
Call sbInicial
End Sub




Private Sub OptX_Click(Index As Integer)

cboResultados.Clear
cboFiltro.Clear

Select Case Index
  Case 0 'Patrimonio
       cboResultados.AddItem "01 - Aportes (Resumen)"
       cboResultados.AddItem "02 - Aportes (Detalle)"
       cboResultados.AddItem "03 - Aportes (Rsm x Categoria)"
       cboResultados.AddItem "04 - Metodo Contable"
       
       cboResultados.Text = "01 - Aportes (Resumen)"
       
       cboFiltro.AddItem "Patrimonio + FCI"
       cboFiltro.AddItem "Solo Patrimonio"
       cboFiltro.Text = "Solo Patrimonio"
  
  Case 1 'Fondos
       cboResultados.AddItem "01 - Fondos (Resumen)"
       cboResultados.AddItem "02 - Fondos (Detalle)"
       cboResultados.Text = "01 - Fondos (Resumen)"
       
       cboFiltro.AddItem "00 - Todos"
       cboFiltro.AddItem "01 - Contratos Activos"
       cboFiltro.AddItem "02 - Liquidados Activos"
       cboFiltro.Text = "00 - Todos"
  
  Case 2 'Creditos
       cboResultados.AddItem "01 - Créditos (Resumen)"
       cboResultados.AddItem "02 - Créditos (Detalle)"
       cboResultados.AddItem "03 - Saldo Inicial Negativo"
       cboResultados.AddItem "04 - Saldo Final Negativo"
       cboResultados.AddItem "05 - Metodo Contable"
       cboResultados.AddItem "06 - Cartera x Garantia"
       cboResultados.AddItem "07 - Complementario"
       cboResultados.AddItem "08 - Producto Acumulado"
       cboResultados.AddItem "09 - Producto Acumulado [Resumen]"
       
       
       cboResultados.AddItem "10 - Créditos Balance (Resumen)"
       cboResultados.AddItem "11 - Créditos Balance (Detalle)"
       
       cboResultados.AddItem "12 - Int.Cbr.Adelantado [Resumen]"
       cboResultados.AddItem "13 - Int.Cbr.Adelantado [Detalle]"
       
       cboResultados.AddItem "14 - Prod. Acum. Suspenso"
       cboResultados.AddItem "15 - Prod. Acum. Suspenso [Resumen]"
       
       cboResultados.AddItem "16 - Gastos/Cargos Diferidos"
       cboResultados.AddItem "17 - Gastos/Cargos Diferidos [Resumen]"
       
       
       cboResultados.Text = "01 - Créditos (Resumen)"
       
       cboFiltro.AddItem "00 - Todos"
       cboFiltro.AddItem "01 - Créditos"
       cboFiltro.AddItem "02 - Retenciones"
       cboFiltro.Text = "00 - Todos"

End Select



End Sub


Private Sub txtCodigo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then
    gBusquedas.Convertir = "N"
    gBusquedas.Resultado = ""
    gBusquedas.Consulta = "Select CODIGO,DESCRIPCION From CATALOGO"
    gBusquedas.Columna = "DESCRIPCION"
    gBusquedas.Orden = "DESCRIPCION"
    
'    gBusquedas.Filtro = " and LINEA_INTERNA = 1 AND RETENCION = 'N' AND POLIZA = 'N'"
    
    frmBusquedas.Show vbModal
    txtCodigo.Text = Trim(gBusquedas.Resultado)
    txtCodigo_Desc.Text = Trim(gBusquedas.Resultado2)
End If
End Sub



Private Sub txtCuenta_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo vError

If KeyCode = vbKeyF4 Then
    frmCntX_ConsultaCuentas.Show vbModal
    txtCuenta.Text = gCuenta
    txtCuenta_Desc.Text = fxgCntCuentaDesc(gCuenta)
End If

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub txtDestino_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then
    gBusquedas.Convertir = "N"
    gBusquedas.Resultado = ""
    gBusquedas.Consulta = "Select COD_DESTINO,DESCRIPCION From CATALOGO_DESTINOS"
    gBusquedas.Columna = "DESCRIPCION"
    gBusquedas.Orden = "DESCRIPCION"
    gBusquedas.Filtro = ""
    frmBusquedas.Show vbModal
    txtDestino.Text = Trim(gBusquedas.Resultado)
    txtDestino_Desc.Text = Trim(gBusquedas.Resultado2)
End If
End Sub

Private Sub txtInstitucion_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then
    gBusquedas.Convertir = "N"
    gBusquedas.Resultado = ""
    gBusquedas.Consulta = "Select COD_INSTITUCION,DESCRIPCION From INSTITUCIONES"
    gBusquedas.Columna = "DESCRIPCION"
    gBusquedas.Orden = "DESCRIPCION"
    gBusquedas.Filtro = ""
    frmBusquedas.Show vbModal
    txtInstitucion.Text = Trim(gBusquedas.Resultado)
    txtInstitucion_Desc.Text = Trim(gBusquedas.Resultado2)
End If
End Sub



Private Sub txtPlan_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then
   gBusquedas.Columna = "cod_plan"
   gBusquedas.Orden = "cod_plan"
   gBusquedas.Filtro = "And Cod_operadora=" & cboOperadora.ItemData(cboOperadora.ListIndex)
   gBusquedas.Consulta = "select cod_plan,descripcion from fnd_planes"
   frmBusquedas.Show vbModal
   txtPlanDesc.SetFocus
   
   If Trim(gBusquedas.Resultado) <> "" Then
      txtPlan.Text = Trim(gBusquedas.Resultado)
      txtPlanDesc.Text = Trim(gBusquedas.Resultado2)
   End If
   gBusquedas.Resultado = ""
   gBusquedas.Resultado2 = ""
End If
End Sub

Private Sub txtRecurso_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then
    gBusquedas.Convertir = "N"
    gBusquedas.Resultado = ""
    gBusquedas.Consulta = "Select COD_GRUPO,DESCRIPCION From catalogo_grupos"
    gBusquedas.Columna = "DESCRIPCION"
    gBusquedas.Orden = "DESCRIPCION"
    gBusquedas.Filtro = ""
    frmBusquedas.Show vbModal
    txtRecurso.Text = Trim(gBusquedas.Resultado)
    txtRecurso_Desc.Text = Trim(gBusquedas.Resultado2)
End If

End Sub
