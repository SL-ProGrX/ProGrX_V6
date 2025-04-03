VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.Controls.v24.0.0.ocx"
Begin VB.Form frmFNDLiquida_Auto_Config 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Configuración para Liquidaciones Automáticas"
   ClientHeight    =   8415
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   10320
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8415
   ScaleWidth      =   10320
   ShowInTaskbar   =   0   'False
   Begin XtremeSuiteControls.TabControl tcMain 
      Height          =   6975
      Left            =   120
      TabIndex        =   0
      Top             =   1320
      Width           =   10095
      _Version        =   1572864
      _ExtentX        =   17806
      _ExtentY        =   12303
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
      ItemCount       =   4
      Item(0).Caption =   "Parámetros"
      Item(0).ControlCount=   1
      Item(0).Control(0)=   "vGrid"
      Item(1).Caption =   "Planes"
      Item(1).ControlCount=   2
      Item(1).Control(0)=   "lsw"
      Item(1).Control(1)=   "gbGarantias"
      Item(2).Caption =   "Planes con Componente Patronal"
      Item(2).ControlCount=   2
      Item(2).Control(0)=   "lswCP"
      Item(2).Control(1)=   "GroupBox1"
      Item(3).Caption =   "Reportes"
      Item(3).ControlCount=   15
      Item(3).Control(0)=   "Label3(0)"
      Item(3).Control(1)=   "Label3(1)"
      Item(3).Control(2)=   "Label3(2)"
      Item(3).Control(3)=   "Label3(3)"
      Item(3).Control(4)=   "GroupBox2"
      Item(3).Control(5)=   "Label2(3)"
      Item(3).Control(6)=   "cboReporteTipo"
      Item(3).Control(7)=   "txtR_Cedula"
      Item(3).Control(8)=   "cboR_Planes"
      Item(3).Control(9)=   "cboR_Proceso"
      Item(3).Control(10)=   "btnReporte"
      Item(3).Control(11)=   "lswR"
      Item(3).Control(12)=   "txtR_Plan"
      Item(3).Control(13)=   "txtR_PlanDesc"
      Item(3).Control(14)=   "btnR_Detallado"
      Begin XtremeSuiteControls.ListView lswCP 
         Height          =   4455
         Left            =   -69880
         TabIndex        =   14
         Top             =   480
         Visible         =   0   'False
         Width           =   9810
         _Version        =   1572864
         _ExtentX        =   17304
         _ExtentY        =   7858
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
         Appearance      =   17
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.ListView lsw 
         Height          =   4575
         Left            =   -69760
         TabIndex        =   3
         Top             =   480
         Visible         =   0   'False
         Width           =   9810
         _Version        =   1572864
         _ExtentX        =   17304
         _ExtentY        =   8070
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
         Appearance      =   17
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.ListView lswR 
         Height          =   3135
         Left            =   -69760
         TabIndex        =   41
         Top             =   2880
         Visible         =   0   'False
         Width           =   9690
         _Version        =   1572864
         _ExtentX        =   17092
         _ExtentY        =   5530
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
         Appearance      =   17
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtR_PlanDesc 
         Height          =   315
         Left            =   -67720
         TabIndex        =   43
         Top             =   6240
         Visible         =   0   'False
         Width           =   5655
         _Version        =   1572864
         _ExtentX        =   9970
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.GroupBox GroupBox2 
         Height          =   1095
         Left            =   -69760
         TabIndex        =   34
         Top             =   1680
         Visible         =   0   'False
         Width           =   9615
         _Version        =   1572864
         _ExtentX        =   16960
         _ExtentY        =   1931
         _StockProps     =   79
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         BorderStyle     =   2
         Begin XtremeSuiteControls.CheckBox chkR_InformeAlDia 
            Height          =   255
            Left            =   4800
            TabIndex        =   36
            Top             =   120
            Width           =   2655
            _Version        =   1572864
            _ExtentX        =   4683
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Reporte al día"
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
         Begin XtremeSuiteControls.ComboBox cboR_Proceso_FA 
            Height          =   330
            Left            =   2400
            TabIndex        =   38
            Top             =   480
            Width           =   1695
            _Version        =   1572864
            _ExtentX        =   2990
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
            Style           =   2
            Appearance      =   6
            UseVisualStyle  =   0   'False
            Text            =   "ComboBox1"
         End
         Begin XtremeSuiteControls.PushButton btnR_Mensual 
            Height          =   375
            Left            =   4800
            TabIndex        =   39
            Top             =   480
            Width           =   2055
            _Version        =   1572864
            _ExtentX        =   3625
            _ExtentY        =   661
            _StockProps     =   79
            Caption         =   "Resumen Mensual"
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
            Picture         =   "frmFNDLiquida_Auto_Config.frx":0000
         End
         Begin XtremeSuiteControls.PushButton btnR_Exportar 
            Height          =   375
            Left            =   6840
            TabIndex        =   40
            Top             =   480
            Width           =   2055
            _Version        =   1572864
            _ExtentX        =   3625
            _ExtentY        =   661
            _StockProps     =   79
            Caption         =   "Exportar"
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
            Picture         =   "frmFNDLiquida_Auto_Config.frx":0707
         End
         Begin XtremeSuiteControls.Label Label3 
            Height          =   255
            Index           =   5
            Left            =   120
            TabIndex        =   37
            Top             =   480
            Width           =   4695
            _Version        =   1572864
            _ExtentX        =   8281
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Proceso a consultar Resumen"
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
         Begin XtremeSuiteControls.Label Label3 
            Height          =   255
            Index           =   4
            Left            =   120
            TabIndex        =   35
            Top             =   120
            Width           =   4695
            _Version        =   1572864
            _ExtentX        =   8281
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Reporte de Fondos Administrativos"
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
      Begin XtremeSuiteControls.PushButton btnReporte 
         Height          =   375
         Left            =   -63880
         TabIndex        =   33
         Top             =   480
         Visible         =   0   'False
         Width           =   1215
         _Version        =   1572864
         _ExtentX        =   2143
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Reporte"
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
      Begin FPSpreadADO.fpSpread vGrid 
         Height          =   6135
         Left            =   240
         TabIndex        =   1
         Top             =   600
         Width           =   9615
         _Version        =   524288
         _ExtentX        =   16960
         _ExtentY        =   10821
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
         SpreadDesigner  =   "frmFNDLiquida_Auto_Config.frx":0871
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin XtremeSuiteControls.GroupBox gbGarantias 
         Height          =   1575
         Left            =   -69280
         TabIndex        =   4
         Top             =   5160
         Visible         =   0   'False
         Width           =   8415
         _Version        =   1572864
         _ExtentX        =   14843
         _ExtentY        =   2778
         _StockProps     =   79
         Caption         =   "Registro"
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
         BorderStyle     =   1
         Begin XtremeSuiteControls.CheckBox chkPatrimonio 
            Height          =   315
            Left            =   960
            TabIndex        =   5
            Top             =   1200
            Width           =   1695
            _Version        =   1572864
            _ExtentX        =   2984
            _ExtentY        =   556
            _StockProps     =   79
            Caption         =   "Patrimonio?   "
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
            Enabled         =   0   'False
            TextAlignment   =   1
            Appearance      =   16
            Alignment       =   1
         End
         Begin XtremeSuiteControls.PushButton btnMov 
            Height          =   315
            Index           =   0
            Left            =   3000
            TabIndex        =   6
            Top             =   1200
            Width           =   375
            _Version        =   1572864
            _ExtentX        =   656
            _ExtentY        =   556
            _StockProps     =   79
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
            FlatStyle       =   -1  'True
            Appearance      =   16
            Picture         =   "frmFNDLiquida_Auto_Config.frx":0DFE
         End
         Begin XtremeSuiteControls.ComboBox cboOperadora 
            Height          =   312
            Left            =   1680
            TabIndex        =   7
            Top             =   360
            Width           =   6612
            _Version        =   1572864
            _ExtentX        =   11668
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
            Style           =   2
            Appearance      =   6
            UseVisualStyle  =   0   'False
            Text            =   "ComboBox1"
         End
         Begin XtremeSuiteControls.FlatEdit txtPlan 
            Height          =   312
            Left            =   1680
            TabIndex        =   8
            Top             =   720
            Width           =   972
            _Version        =   1572864
            _ExtentX        =   1714
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
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtPlanDesc 
            Height          =   312
            Left            =   2640
            TabIndex        =   9
            Top             =   720
            Width           =   5652
            _Version        =   1572864
            _ExtentX        =   9970
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
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.PushButton btnMov 
            Height          =   315
            Index           =   1
            Left            =   3360
            TabIndex        =   10
            Top             =   1200
            Width           =   375
            _Version        =   1572864
            _ExtentX        =   656
            _ExtentY        =   556
            _StockProps     =   79
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
            FlatStyle       =   -1  'True
            Appearance      =   16
            Picture         =   "frmFNDLiquida_Auto_Config.frx":151E
         End
         Begin XtremeSuiteControls.FlatEdit txtLinea 
            Height          =   315
            Left            =   7320
            TabIndex        =   11
            Top             =   1200
            Visible         =   0   'False
            Width           =   975
            _Version        =   1572864
            _ExtentX        =   1714
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
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin VB.Label Label2 
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
            Height          =   252
            Index           =   2
            Left            =   600
            TabIndex        =   13
            Top             =   720
            Width           =   852
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Operadora"
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
            Index           =   4
            Left            =   600
            TabIndex        =   12
            Top             =   360
            Width           =   972
         End
      End
      Begin XtremeSuiteControls.GroupBox GroupBox1 
         Height          =   1815
         Left            =   -69280
         TabIndex        =   15
         Top             =   5040
         Visible         =   0   'False
         Width           =   8415
         _Version        =   1572864
         _ExtentX        =   14843
         _ExtentY        =   3201
         _StockProps     =   79
         Caption         =   "Registro"
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
         BorderStyle     =   1
         Begin XtremeSuiteControls.CheckBox chkComponentePatronal 
            Height          =   315
            Left            =   2760
            TabIndex        =   16
            Top             =   1080
            Width           =   2415
            _Version        =   1572864
            _ExtentX        =   4260
            _ExtentY        =   556
            _StockProps     =   79
            Caption         =   "Componente Patronal?   "
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
            TextAlignment   =   1
            Appearance      =   17
         End
         Begin XtremeSuiteControls.PushButton btnMov 
            Height          =   315
            Index           =   2
            Left            =   2760
            TabIndex        =   17
            Top             =   1560
            Width           =   375
            _Version        =   1572864
            _ExtentX        =   656
            _ExtentY        =   556
            _StockProps     =   79
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
            FlatStyle       =   -1  'True
            Appearance      =   16
            Picture         =   "frmFNDLiquida_Auto_Config.frx":1AC2
         End
         Begin XtremeSuiteControls.ComboBox cboCP_Operadora 
            Height          =   312
            Left            =   1680
            TabIndex        =   18
            Top             =   360
            Width           =   6612
            _Version        =   1572864
            _ExtentX        =   11668
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
            Style           =   2
            Appearance      =   6
            UseVisualStyle  =   0   'False
            Text            =   "ComboBox1"
         End
         Begin XtremeSuiteControls.FlatEdit txtCP_Plan 
            Height          =   312
            Left            =   1680
            TabIndex        =   19
            Top             =   720
            Width           =   972
            _Version        =   1572864
            _ExtentX        =   1714
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
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtCP_PlanDesc 
            Height          =   312
            Left            =   2640
            TabIndex        =   20
            Top             =   720
            Width           =   5652
            _Version        =   1572864
            _ExtentX        =   9970
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
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.PushButton btnMov 
            Height          =   315
            Index           =   3
            Left            =   3120
            TabIndex        =   21
            Top             =   1560
            Width           =   375
            _Version        =   1572864
            _ExtentX        =   656
            _ExtentY        =   556
            _StockProps     =   79
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
            FlatStyle       =   -1  'True
            Appearance      =   16
            Picture         =   "frmFNDLiquida_Auto_Config.frx":21E2
         End
         Begin XtremeSuiteControls.FlatEdit txtCP_Linea 
            Height          =   315
            Left            =   1680
            TabIndex        =   22
            Top             =   1080
            Visible         =   0   'False
            Width           =   975
            _Version        =   1572864
            _ExtentX        =   1714
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
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Operadora"
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
            Index           =   1
            Left            =   600
            TabIndex        =   24
            Top             =   360
            Width           =   972
         End
         Begin VB.Label Label2 
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
            Height          =   252
            Index           =   0
            Left            =   600
            TabIndex        =   23
            Top             =   720
            Width           =   852
         End
      End
      Begin XtremeSuiteControls.ComboBox cboReporteTipo 
         Height          =   330
         Left            =   -66280
         TabIndex        =   26
         Top             =   480
         Visible         =   0   'False
         Width           =   2295
         _Version        =   1572864
         _ExtentX        =   4048
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
         Style           =   2
         Appearance      =   6
         UseVisualStyle  =   0   'False
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.FlatEdit txtR_Cedula 
         Height          =   330
         Left            =   -69760
         TabIndex        =   28
         Top             =   1200
         Visible         =   0   'False
         Width           =   2295
         _Version        =   1572864
         _ExtentX        =   4048
         _ExtentY        =   582
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.ComboBox cboR_Planes 
         Height          =   330
         Left            =   -67480
         TabIndex        =   30
         Top             =   1200
         Visible         =   0   'False
         Width           =   5295
         _Version        =   1572864
         _ExtentX        =   9340
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
         Style           =   2
         Appearance      =   6
         UseVisualStyle  =   0   'False
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.ComboBox cboR_Proceso 
         Height          =   330
         Left            =   -62200
         TabIndex        =   32
         Top             =   1200
         Visible         =   0   'False
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
         Style           =   2
         Appearance      =   6
         UseVisualStyle  =   0   'False
         Text            =   "ComboBox1"
      End
      Begin XtremeSuiteControls.FlatEdit txtR_Plan 
         Height          =   315
         Left            =   -68680
         TabIndex        =   42
         Top             =   6240
         Visible         =   0   'False
         Width           =   975
         _Version        =   1572864
         _ExtentX        =   1714
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
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.PushButton btnR_Detallado 
         Height          =   615
         Left            =   -61960
         TabIndex        =   45
         Top             =   6120
         Visible         =   0   'False
         Width           =   1575
         _Version        =   1572864
         _ExtentX        =   2778
         _ExtentY        =   1085
         _StockProps     =   79
         Caption         =   "Reporte Detallado"
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
         Picture         =   "frmFNDLiquida_Auto_Config.frx":2786
      End
      Begin VB.Label Label2 
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
         Height          =   255
         Index           =   3
         Left            =   -69760
         TabIndex        =   44
         Top             =   6240
         Visible         =   0   'False
         Width           =   855
      End
      Begin XtremeSuiteControls.Label Label3 
         Height          =   255
         Index           =   3
         Left            =   -62200
         TabIndex        =   31
         Top             =   960
         Visible         =   0   'False
         Width           =   1455
         _Version        =   1572864
         _ExtentX        =   2566
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Proceso "
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
      Begin XtremeSuiteControls.Label Label3 
         Height          =   255
         Index           =   2
         Left            =   -67360
         TabIndex        =   29
         Top             =   960
         Visible         =   0   'False
         Width           =   1575
         _Version        =   1572864
         _ExtentX        =   2778
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Plan "
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
      Begin XtremeSuiteControls.Label Label3 
         Height          =   255
         Index           =   1
         Left            =   -69760
         TabIndex        =   27
         Top             =   960
         Visible         =   0   'False
         Width           =   1575
         _Version        =   1572864
         _ExtentX        =   2778
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Cédula Persona"
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
      Begin XtremeSuiteControls.Label Label3 
         Height          =   255
         Index           =   0
         Left            =   -67960
         TabIndex        =   25
         Top             =   480
         Visible         =   0   'False
         Width           =   1575
         _Version        =   1572864
         _ExtentX        =   2778
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Tipo de Reporte"
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
      Left            =   120
      TabIndex        =   46
      Top             =   1200
      Visible         =   0   'False
      Width           =   10095
      _Version        =   1572864
      _ExtentX        =   17806
      _ExtentY        =   238
      _StockProps     =   93
      BackColor       =   -2147483633
      Scrolling       =   1
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Parámetros para Liquidación automática de Fondos"
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
      Height          =   375
      Index           =   3
      Left            =   1680
      TabIndex        =   2
      Top             =   360
      Width           =   7335
   End
   Begin VB.Image imgBanner 
      Height          =   1092
      Left            =   0
      Top             =   0
      Width           =   10932
   End
End
Attribute VB_Name = "frmFNDLiquida_Auto_Config"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem
Dim vPaso As Boolean

Private Sub btnMov_Click(Index As Integer)

Dim strSQL As String, strBitacora As String
Dim pOperadora As Integer, pPlan As String, pGarantia As String, pEstado As String

On Error GoTo vError

Select Case Index
    Case 0, 1
            pOperadora = cboOperadora.ItemData(cboOperadora.ListIndex)
            pPlan = txtPlan.Text

            strSQL = "exec spFnd_LiqAuto_Planes_Add " & pOperadora & ",'" & pPlan & "'," & chkPatrimonio.Value & ",'" & glogon.Usuario & "'"

            strBitacora = "Liquidación Automática: Planes Vincula, Línea: " & txtLinea.Text & ",  Plan: " & pPlan & " Patronal: " _
                        & IIf(chkPatrimonio.Value = xtpChecked, "Sí", "No")
    Case 2, 3
            pOperadora = cboCP_Operadora.ItemData(cboCP_Operadora.ListIndex)
            pPlan = txtCP_Plan.Text
            
            strSQL = "exec spFnd_LiqAuto_Planes_PAT_Add " & pOperadora & ",'" & pPlan & "'," & chkPatrimonio.Value & ",'" & glogon.Usuario & "'"
    
            strBitacora = "Liquidación Automática: Planes Patrimoniales, Línea: " & txtCP_Linea.Text & ",  Plan: " & pPlan & " Patronal: " _
                        & IIf(chkComponentePatronal.Value = xtpChecked, "Sí", "No")
    
End Select


Select Case Index
 Case 0 'Agregar / Modificar
       
    strSQL = strSQL & ",'A'"
    Call ConectionExecute(strSQL)
       
    Call Bitacora("Registra", strBitacora)
    
 
 Case 1 'Elimina

    strSQL = strSQL & ",'B'"
    Call ConectionExecute(strSQL)
    
    Call Bitacora("Elimina", strBitacora)
'----------------------
    
 Case 2 'Agregar / Modificar
       
    strSQL = strSQL & ",'A'"
    Call ConectionExecute(strSQL)
       
    Call Bitacora("Registra", strBitacora)
    
 
 Case 3 'Elimina

    strSQL = strSQL & ",'B'"
    Call ConectionExecute(strSQL)
    
    Call Bitacora("Elimina", strBitacora)
    
End Select



Select Case Index
    Case 0, 1
       txtPlan.Text = ""
       chkPatrimonio.Value = xtpUnchecked
       txtLinea.Text = "0"
       Call sbPlanes_Load
    
    Case 2, 3
       txtCP_Plan.Text = ""
       chkComponentePatronal.Value = xtpUnchecked
       txtCP_Linea.Text = "0"
       Call sbPlanes_CP_Load
End Select


Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub btnR_Exportar_Click()
On Error GoTo vError

Me.MousePointer = vbHourglass

ProgressBarX.Visible = True

Call Excel_Exportar_Lsw(lswR, ProgressBarX)

ProgressBarX.Visible = False

Me.MousePointer = vbDefault

Exit Sub

vError:
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub cboR_Proceso_Click()
Call sbR_Resumen
End Sub

Private Sub Form_Activate()
vModulo = 18
End Sub

Private Sub Form_Load()

On Error GoTo vError

vModulo = 18
vGrid.AppearanceStyle = fxGridStyle

Set imgBanner.Picture = frmContenedor.imgBanner_Mantenimiento.Picture

'Inicializa Parametros
'strSQL = "exec spFnd_LiqAuto_Parametros"
'Call ConectionExecute(strSQL)

tcMain.Item(0).Selected = True

With lsw.ColumnHeaders
    .Clear
    .Add , , "[Id]", 600
    .Add , , "[Op]", 600, vbCenter
    .Add , , "Plan", 900, vbCenter
    .Add , , "Descripción", 3000
End With

With lswCP.ColumnHeaders
    .Clear
    .Add , , "[Id]", 600
    .Add , , "[Op]", 600, vbCenter
    .Add , , "Plan", 900, vbCenter
    .Add , , "Descripción", 3000
    .Add , , "Patrimonio", 1100, vbCenter
End With


cboReporteTipo.AddItem "Detallado"
cboReporteTipo.AddItem "Resumen"
cboReporteTipo.Text = "Detallado"

With lswR.ColumnHeaders
    .Clear
    .Add , , "Plan", 900, vbCenter
    .Add , , "Descripción", 3000
    .Add , , "Cantidad", 1100, vbCenter
    .Add , , "SaldoTotal", 2100, vbRightJustify
    .Add , , "Fecha", 2100, vbCenter
    .Add , , "Usuario", 2100, vbCenter
    .Add , , "Proceso", 1100, vbCenter
End With

strSQL = "select C.CodPlan as 'IdX', P.DESCRIPCION as 'ItmX'" _
       & " from FND_LIQUIDACION_AUTOMATICA_PLANES C inner join FND_PLANES P on C.Operadora  = P.COD_OPERADORA and C.CodPlan = P.COD_PLAN" _
       & " order by C.IdRegistro  "
Call sbCbo_Llena_New(cboR_Planes, strSQL, True, True)


strSQL = " select convert(varchar(4),Anio) +  format(Mes, '00') as 'ItmX', convert(varchar(4),Anio) +  format(Mes, '00') as 'IdX'" _
       & " From FND_LIQUIDACION_AUTOMATICA_RESUMEN" _
       & "  group by anio, mes" _
       & " order by Anio desc, mes desc"
Call sbCbo_Llena_New(cboR_Proceso, strSQL, False, True)

Call sbCbo_Copia(cboR_Proceso, cboR_Proceso_FA)

strSQL = "select rtrim(cod_Operadora) as 'IdX', rtrim(descripcion) as ItmX" _
         & " from  fnd_Operadoras"
Call sbCbo_Llena_New(cboOperadora, strSQL, False, True)

Call sbCbo_Copia(cboOperadora, cboCP_Operadora)

Call sbParametros_Load

Call Formularios(Me)
Call RefrescaTags(Me)

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbExclamation

End Sub



Public Sub sbCargaGridLocal(pGrid As Object, MaxCol As Integer, strSQL As String)
Dim rs As New ADODB.Recordset, i As Integer

With pGrid
    .MaxRows = 0
    .MaxCols = MaxCol
    Call OpenRecordSet(rs, strSQL, 0)
    Do While Not rs.EOF
      .MaxRows = .MaxRows + 1
      .Row = .MaxRows
      For i = 1 To 3
        .Col = i
        Select Case i
          Case 1 'Codigo
            .CellTag = Mid(rs!TipoDato, 1, 3) & ""
            .Text = rs!IdRegistro
            .CellNote = "Modificado Por: " & rs!UsuarioActualiza & vbCrLf & "Fecha: " & rs!FechaActualiza
          
            .TextTip = TextTipFixed
            .TextTipDelay = 1000
          
          Case 2 'Descripcion
            .Text = rs!Descripcion
            .CellNote = ""  'rs!Notas & ""
            .TextTip = TextTipFixed
            .TextTipDelay = 1000
          
          Case 3 'Valor
            If UCase(Mid(Trim(rs!TipoDato), 1, 3)) = "CTA" Then
                .TextTip = TextTipFixed
                .TextTipDelay = 1000
                .CellNoteIndicatorColor = vbBlue
                .CellNoteIndicatorShape = CellNoteIndicatorShapeTriangle
                
                .Text = fxgCntCuentaFormato(True, Trim(rs!Valor), 0)
                .CellNote = fxgCntCuentaDesc(Trim(rs!Valor))
            Else
                .Text = rs!Valor
            End If
            
        End Select
      Next i
      rs.MoveNext
    Loop
    rs.Close

End With

End Sub


Private Sub sbGuardaParametro(pParametro As String, pValor As String _
                    , Optional pTipo As String = "DEC")
Dim strSQL As String, rs As New ADODB.Recordset
Dim Validacion As Boolean, vMensaje As String

On Error GoTo vError

Validacion = True
vMensaje = ""

Select Case UCase(Mid(Trim(pTipo), 1, 3))
  Case "DEC", "INT" 'Decimal
    If IsNumeric(pValor) Then
       pValor = CCur(pValor)
    Else
       Validacion = False
       vMensaje = "El valor indicado no es válido...!!!"
    End If
    
  Case "NUM" 'Número Entero
    If IsNumeric(pValor) Then
       pValor = CLng(pValor)
    Else
       Validacion = False
       vMensaje = "El valor indicado no es válido...!!!"
    End If
  
  Case "POR" 'Porcentaje
    If IsNumeric(pValor) Then
       pValor = CCur(pValor)
    Else
       Validacion = False
       vMensaje = "El valor indicado no es válido, suministre un porcentaje ..!!!"
    End If
  
  Case "CTA" 'Cuenta Contable
    Validacion = fxgCntCuentaValida(fxgCntCuentaFormato(False, pValor))
    If Not Validacion Then
        vMensaje = "La Cuenta indicada no es válida, presiones F4 para buscar en el catálogo...!!!"
    Else
      pValor = fxgCntCuentaFormato(False, pValor)
    End If
    
  Case "CHR" 'Caracteres
    If InStr(1, pValor, "'", vbTextCompare) > 0 Then
       Validacion = False
       vMensaje = "El valor indicado contiene caracteres no válidos...!!!"
    End If
    
  Case "PSN" 'Pregunta S ó N
     If UCase(Mid(pValor, 1, 1)) = "S" Or UCase(Mid(pValor, 1, 1)) = "N" Then
       pValor = UCase(Mid(pValor, 1, 1))
     Else
       Validacion = False
       vMensaje = "El valor indicado no es válido > Indique [S] ó [N]...!!!"
     End If
     
  Case "DTS" 'Fecha
    
    If Not IsDate(pValor) Then
       Validacion = False
       vMensaje = "La Fecha indicada no es válida...!!!"
    Else
       pValor = Format(CDate(pValor), "yyyy/mm/dd")
    End If

End Select


If Not Validacion Then
  MsgBox vMensaje, vbExclamation, "Parámetros de Fondos Liquidación Automática"
  Exit Sub
End If


strSQL = "update FND_LIQUIDACION_AUTOMATICA_PARAMETROS set UsuarioActualiza = '" & glogon.Usuario & "', FechaActualiza = dbo.MyGetdate()" _
       & ",valor = '" & Trim(pValor) & "' where IdRegistro = '" & pParametro & "'"
Call ConectionExecute(strSQL)

strSQL = "Parámetro de Fondos Liquidación Automática: " & pParametro & " -> " & pValor

Call Bitacora("Modifica", strSQL)

MsgBox "Parámetro actualizado satisfactoriamente...!", vbInformation

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Function fxGuardar() As Long
Dim vTemp As String

On Error GoTo vError

fxGuardar = 0

vGrid.Row = vGrid.ActiveRow
vGrid.Col = 3
vTemp = vGrid.Text


vGrid.Col = 1
Call sbGuardaParametro(vGrid.Text, vTemp, vGrid.CellTag)

fxGuardar = 1

Exit Function

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Function

Private Sub sbParametros_Load()

On Error GoTo vError


strSQL = "select IdRegistro, Descripcion, Valor, TipoDato, Operadora,UsuarioActualiza, FechaActualiza  from FND_LIQUIDACION_AUTOMATICA_PARAMETROS" _
      & " order by IdRegistro"
Call sbCargaGridLocal(vGrid, 3, strSQL)

Exit Sub

vError:

End Sub


Private Sub sbPlanes_Load()

On Error GoTo vError

lsw.ListItems.Clear

With lsw.ColumnHeaders
    .Clear
    .Add , , "[Id]", 600
    .Add , , "[Op]", 600, vbCenter
    .Add , , "Plan", 900, vbCenter
    .Add , , "Descripción", 3000
    .Add , , "Patrimonio", 1100, vbCenter
    .Add , , "Reg.Fecha", 2100, vbCenter
    .Add , , "Reg.Usuario", 2100, vbCenter
    
End With

strSQL = "select C.idRegistro, C.Operadora, C.CodPlan, C.FechaRegistro, C.UsuarioRegistro, C.ComponentePatronal" _
       & ", P.DESCRIPCION" _
       & " from FND_LIQUIDACION_AUTOMATICA_PLANES C inner join FND_PLANES P on C.Operadora  = P.COD_OPERADORA and C.CodPlan = P.COD_PLAN" _
       & " order by C.IdRegistro   "
Call OpenRecordSet(rs, strSQL)

Do While Not rs.EOF
 Set itmX = lsw.ListItems.Add(, , rs!IdRegistro)
     itmX.SubItems(1) = rs!Operadora
     itmX.SubItems(2) = rs!CodPlan
     itmX.SubItems(3) = rs!Descripcion
     itmX.SubItems(4) = rs!ComponentePatronal
     itmX.SubItems(5) = rs!FechaRegistro & ""
     itmX.SubItems(6) = rs!UsuarioRegistro & ""
 rs.MoveNext
Loop
rs.Close

Exit Sub

vError:

End Sub


Private Sub sbPlanes_CP_Load()

On Error GoTo vError

lswCP.ListItems.Clear

With lswCP.ColumnHeaders
    .Clear
    .Add , , "[Id]", 600
    .Add , , "[Op]", 600, vbCenter
    .Add , , "Plan", 900, vbCenter
    .Add , , "Descripción", 3000
    .Add , , "Patrimonio", 1100, vbCenter
    .Add , , "Reg.Fecha", 2100, vbCenter
    .Add , , "Reg.Usuario", 2100, vbCenter
    
End With

strSQL = "select C.idRegistro, C.Operadora, C.CodPlan, C.FechaRegistro, C.UsuarioRegistro, C.ComponentePatronal" _
       & ", P.DESCRIPCION" _
       & " from FND_LIQUIDACION_AUTOMATICA_PLANES C inner join FND_PLANES P on C.Operadora  = P.COD_OPERADORA and C.CodPlan = P.COD_PLAN" _
       & " Where C.ComponentePatronal = 1 order by C.IdRegistro   "
Call OpenRecordSet(rs, strSQL)

Do While Not rs.EOF
 Set itmX = lswCP.ListItems.Add(, , rs!IdRegistro)
     itmX.SubItems(1) = rs!Operadora
     itmX.SubItems(2) = rs!CodPlan
     itmX.SubItems(3) = rs!Descripcion
     itmX.SubItems(4) = rs!ComponentePatronal
     itmX.SubItems(5) = rs!FechaRegistro & ""
     itmX.SubItems(6) = rs!UsuarioRegistro & ""
 rs.MoveNext
Loop
rs.Close

Exit Sub

vError:

End Sub


Private Sub sbR_Resumen()

On Error GoTo vError

lswR.ListItems.Clear

strSQL = "select C.Id, C.Anio, C.Mes, C.CodPlan, C.CantidadClientes, C.SaldoTotal , C.FechaInserta, C.UsuarioInserta" _
       & ", P.DESCRIPCION, convert(varchar(4),C.Anio) +  format(C.Mes, '00') as 'Proceso'" _
       & " from FND_LIQUIDACION_AUTOMATICA_RESUMEN C inner join FND_PLANES P on C.CodPlan = P.COD_PLAN" _
       & " Where C.Anio = " & Mid(cboR_Proceso.Text, 1, 4) & " and C.Mes = " & Mid(cboR_Proceso.Text, 5, 2) _
       & " order by C.Id"
Call OpenRecordSet(rs, strSQL)


Do While Not rs.EOF
 Set itmX = lswR.ListItems.Add(, , rs!CodPlan)
     itmX.SubItems(1) = rs!Descripcion
     itmX.SubItems(2) = Format(rs!CantidadClientes, "##0")
     itmX.SubItems(3) = Format(rs!SaldoTotal, "Standard")
     itmX.SubItems(4) = rs!FechaInserta & ""
     itmX.SubItems(5) = rs!UsuarioInserta & ""
     itmX.SubItems(6) = rs!Proceso & ""
 rs.MoveNext
Loop
rs.Close

Exit Sub

vError:

End Sub


Private Sub lsw_ColumnClick(ByVal ColumnHeader As XtremeSuiteControls.ListViewColumnHeader)
 lsw.SortKey = ColumnHeader.Index - 1
  If lsw.SortOrder = 0 Then lsw.SortOrder = 1 Else lsw.SortOrder = 0
  lsw.Sorted = True
End Sub

Private Sub lsw_ItemClick(ByVal Item As XtremeSuiteControls.ListViewItem)

On Error GoTo vError


txtLinea.Text = Item.Text
txtPlan.Text = Item.SubItems(2)
txtPlanDesc.Text = Item.SubItems(3)
chkPatrimonio.Value = Item.SubItems(4)


vError:

End Sub



Private Sub lswCP_ColumnClick(ByVal ColumnHeader As XtremeSuiteControls.ListViewColumnHeader)
 lswCP.SortKey = ColumnHeader.Index - 1
  If lswCP.SortOrder = 0 Then lswCP.SortOrder = 1 Else lswCP.SortOrder = 0
  lswCP.Sorted = True
End Sub

Private Sub lswCP_ItemClick(ByVal Item As XtremeSuiteControls.ListViewItem)

On Error GoTo vError


txtCP_Linea.Text = Item.Text
txtCP_Plan.Text = Item.SubItems(2)
txtCP_PlanDesc.Text = Item.SubItems(3)
chkComponentePatronal.Value = Item.SubItems(4)

vError:

End Sub


Private Sub lswr_ColumnClick(ByVal ColumnHeader As XtremeSuiteControls.ListViewColumnHeader)
 lswR.SortKey = ColumnHeader.Index - 1
  If lswR.SortOrder = 0 Then lswR.SortOrder = 1 Else lswR.SortOrder = 0
  lswR.Sorted = True
End Sub

Private Sub lswR_ItemClick(ByVal Item As XtremeSuiteControls.ListViewItem)

On Error GoTo vError

txtR_Plan.Text = Item.Text
txtR_PlanDesc.Text = Item.SubItems(1)

vError:


End Sub

Private Sub tcMain_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)

Select Case Item.Index
  Case 0 'Parametros
    Call sbParametros_Load
  Case 1 'Planes
    Call sbPlanes_Load
  Case 2 'Planes Componente Patronal
    Call sbPlanes_CP_Load
  Case 3 'Reportes
    Call sbR_Resumen
End Select

End Sub

Private Sub vGrid_KeyDown(KeyCode As Integer, Shift As Integer)
Dim i As Integer

If vGrid.ActiveCol = vGrid.MaxCols And (KeyCode = vbKeyReturn Or KeyCode = vbKeyTab) Then
  i = fxGuardar
  If i = 0 Then Exit Sub
  vGrid.Row = vGrid.ActiveRow
End If

If vGrid.ActiveCol = vGrid.MaxCols And KeyCode = vbKeyF4 Then
   vGrid.Row = vGrid.ActiveRow
   vGrid.Col = 1
   If vGrid.CellTag = "CTA" Then
      gCuenta = ""
      frmCntX_ConsultaCuentas.Show vbModal
      If gCuenta <> "" Then
        vGrid.Col = 3
        vGrid.Text = fxgCntCuentaFormato(True, gCuenta)
        vGrid.TextTip = TextTipFixed
        vGrid.TextTipDelay = 1000
        vGrid.CellNoteIndicatorColor = vbBlue
        vGrid.CellNoteIndicatorShape = CellNoteIndicatorShapeTriangle
        vGrid.CellNote = fxgCntCuentaDesc(gCuenta)
        
        vGrid.Col = 1
        Call sbGuardaParametro(vGrid.Text, gCuenta, "CTA")
      End If
      
   
   End If
End If

End Sub



