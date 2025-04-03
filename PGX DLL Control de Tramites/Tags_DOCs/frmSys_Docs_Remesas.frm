VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.Controls.v24.0.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.ShortcutBar.v24.0.0.ocx"
Begin VB.Form frmSys_Docs_Remesas 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Documentos: Remesas"
   ClientHeight    =   9420
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   12210
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9420
   ScaleWidth      =   12210
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin XtremeSuiteControls.TabControl tcMain 
      Height          =   7575
      Left            =   120
      TabIndex        =   0
      Top             =   1800
      Width           =   12015
      _Version        =   1572864
      _ExtentX        =   21193
      _ExtentY        =   13361
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
      Item(0).Caption =   "Remesa"
      Item(0).ControlCount=   12
      Item(0).Control(0)=   "lsw"
      Item(0).Control(1)=   "btnAplicar"
      Item(0).Control(2)=   "Label3(1)"
      Item(0).Control(3)=   "dtpInicio"
      Item(0).Control(4)=   "dtpCorte"
      Item(0).Control(5)=   "Label3(4)"
      Item(0).Control(6)=   "cboGrupo"
      Item(0).Control(7)=   "cboUsuario"
      Item(0).Control(8)=   "Label3(3)"
      Item(0).Control(9)=   "btnBuscar"
      Item(0).Control(10)=   "txtNotas"
      Item(0).Control(11)=   "Label3(5)"
      Item(1).Caption =   "Reportes"
      Item(1).ControlCount=   10
      Item(1).Control(0)=   "lswRemesas"
      Item(1).Control(1)=   "dtpR_Inicio"
      Item(1).Control(2)=   "dtpR_Corte"
      Item(1).Control(3)=   "Label3(6)"
      Item(1).Control(4)=   "ShortcutCaption1"
      Item(1).Control(5)=   "lblRemesaId"
      Item(1).Control(6)=   "btnRemesa(0)"
      Item(1).Control(7)=   "btnRemesa(1)"
      Item(1).Control(8)=   "gbMicrofilm"
      Item(1).Control(9)=   "btnRemesa(2)"
      Item(2).Caption =   "Consulta"
      Item(2).ControlCount=   4
      Item(2).Control(0)=   "txtCodigoBuscar"
      Item(2).Control(1)=   "Label3(2)"
      Item(2).Control(2)=   "btnConsulta"
      Item(2).Control(3)=   "txtConRemesa"
      Begin XtremeSuiteControls.ListView lsw 
         Height          =   5175
         Left            =   120
         TabIndex        =   1
         Top             =   1680
         Width           =   11775
         _Version        =   1572864
         _ExtentX        =   20770
         _ExtentY        =   9128
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
      Begin XtremeSuiteControls.ListView lswRemesas 
         Height          =   4935
         Left            =   -69880
         TabIndex        =   22
         Top             =   1320
         Visible         =   0   'False
         Width           =   11775
         _Version        =   1572864
         _ExtentX        =   20770
         _ExtentY        =   8705
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
      Begin XtremeSuiteControls.GroupBox gbMicrofilm 
         Height          =   3015
         Left            =   -67120
         TabIndex        =   29
         Top             =   2040
         Visible         =   0   'False
         Width           =   6735
         _Version        =   1572864
         _ExtentX        =   11880
         _ExtentY        =   5318
         _StockProps     =   79
         UseVisualStyle  =   -1  'True
         Appearance      =   21
         Begin XtremeSuiteControls.FlatEdit txtReciboRemesa 
            Height          =   375
            Left            =   2520
            TabIndex        =   34
            Top             =   720
            Width           =   3015
            _Version        =   1572864
            _ExtentX        =   5318
            _ExtentY        =   661
            _StockProps     =   77
            ForeColor       =   0
            BackColor       =   12648447
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColor       =   12648447
            Alignment       =   2
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtReciboUsuario 
            Height          =   375
            Left            =   2520
            TabIndex        =   35
            Top             =   1200
            Width           =   3015
            _Version        =   1572864
            _ExtentX        =   5318
            _ExtentY        =   661
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
            Alignment       =   2
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtReciboFecha 
            Height          =   375
            Left            =   2520
            TabIndex        =   36
            Top             =   1680
            Width           =   3015
            _Version        =   1572864
            _ExtentX        =   5318
            _ExtentY        =   661
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
            Alignment       =   2
            Locked          =   -1  'True
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.PushButton btnMicrofilm 
            Height          =   375
            Index           =   0
            Left            =   4200
            TabIndex        =   37
            Top             =   2400
            Width           =   1335
            _Version        =   1572864
            _ExtentX        =   2355
            _ExtentY        =   661
            _StockProps     =   79
            Caption         =   "Aplicar"
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
            Appearance      =   21
            Picture         =   "frmSys_Docs_Remesas.frx":0000
         End
         Begin XtremeSuiteControls.PushButton btnMicrofilm 
            Height          =   375
            Index           =   1
            Left            =   5520
            TabIndex        =   38
            Top             =   2400
            Width           =   855
            _Version        =   1572864
            _ExtentX        =   1508
            _ExtentY        =   661
            _StockProps     =   79
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
            Appearance      =   21
            Picture         =   "frmSys_Docs_Remesas.frx":0727
         End
         Begin XtremeSuiteControls.Label Label3 
            Height          =   255
            Index           =   9
            Left            =   360
            TabIndex        =   33
            Top             =   1680
            Width           =   2295
            _Version        =   1572864
            _ExtentX        =   4048
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Fecha de recibido:"
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
            WordWrap        =   -1  'True
         End
         Begin XtremeSuiteControls.Label Label3 
            Height          =   255
            Index           =   8
            Left            =   360
            TabIndex        =   32
            Top             =   1200
            Width           =   2295
            _Version        =   1572864
            _ExtentX        =   4048
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Recibido por: (Usuario)"
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
            WordWrap        =   -1  'True
         End
         Begin XtremeSuiteControls.Label Label3 
            Height          =   255
            Index           =   7
            Left            =   360
            TabIndex        =   31
            Top             =   720
            Width           =   2295
            _Version        =   1572864
            _ExtentX        =   4048
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Id. Remesa de Crédito"
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
            WordWrap        =   -1  'True
         End
         Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption2 
            Height          =   375
            Left            =   0
            TabIndex        =   30
            Top             =   0
            Width           =   6735
            _Version        =   1572864
            _ExtentX        =   11880
            _ExtentY        =   661
            _StockProps     =   14
            Caption         =   "Marcar como Recibido en Microfilm (Archivo digital)"
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
      End
      Begin XtremeSuiteControls.DateTimePicker dtpR_Corte 
         Height          =   330
         Left            =   -67240
         TabIndex        =   24
         Top             =   480
         Visible         =   0   'False
         Width           =   1455
         _Version        =   1572864
         _ExtentX        =   2566
         _ExtentY        =   582
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
      Begin XtremeSuiteControls.PushButton btnAplicar 
         Height          =   375
         Left            =   10320
         TabIndex        =   2
         Top             =   6960
         Width           =   1575
         _Version        =   1572864
         _ExtentX        =   2778
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Crear Remesa"
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
         Appearance      =   21
         Picture         =   "frmSys_Docs_Remesas.frx":0D65
      End
      Begin XtremeSuiteControls.PushButton btnRemesa 
         Height          =   375
         Index           =   0
         Left            =   -60760
         TabIndex        =   3
         Top             =   480
         Visible         =   0   'False
         Width           =   1335
         _Version        =   1572864
         _ExtentX        =   2355
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Buscar"
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
         Appearance      =   21
         Picture         =   "frmSys_Docs_Remesas.frx":148C
      End
      Begin XtremeSuiteControls.PushButton btnRemesa 
         Height          =   375
         Index           =   1
         Left            =   -59440
         TabIndex        =   4
         Top             =   480
         Visible         =   0   'False
         Width           =   1335
         _Version        =   1572864
         _ExtentX        =   2355
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Exportar"
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
         Appearance      =   21
         Picture         =   "frmSys_Docs_Remesas.frx":1B8C
      End
      Begin XtremeSuiteControls.FlatEdit txtCodigoBuscar 
         Height          =   330
         Left            =   -68320
         TabIndex        =   5
         Top             =   600
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   2
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.PushButton btnConsulta 
         Height          =   375
         Left            =   -65080
         TabIndex        =   6
         Top             =   600
         Visible         =   0   'False
         Width           =   1335
         _Version        =   1572864
         _ExtentX        =   2355
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Buscar"
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
         Appearance      =   21
         Picture         =   "frmSys_Docs_Remesas.frx":1CF6
      End
      Begin XtremeSuiteControls.ComboBox cboGrupo 
         Height          =   330
         Left            =   1440
         TabIndex        =   12
         Top             =   840
         Width           =   2895
         _Version        =   1572864
         _ExtentX        =   5106
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
      Begin XtremeSuiteControls.DateTimePicker dtpInicio 
         Height          =   330
         Left            =   1440
         TabIndex        =   14
         Top             =   480
         Width           =   1455
         _Version        =   1572864
         _ExtentX        =   2566
         _ExtentY        =   582
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
         Height          =   330
         Left            =   2880
         TabIndex        =   15
         Top             =   480
         Width           =   1455
         _Version        =   1572864
         _ExtentX        =   2566
         _ExtentY        =   582
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
      Begin XtremeSuiteControls.ComboBox cboUsuario 
         Height          =   330
         Left            =   1440
         TabIndex        =   17
         Top             =   1200
         Width           =   2895
         _Version        =   1572864
         _ExtentX        =   5106
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
      Begin XtremeSuiteControls.PushButton btnBuscar 
         Height          =   375
         Left            =   10680
         TabIndex        =   19
         Top             =   360
         Width           =   1215
         _Version        =   1572864
         _ExtentX        =   2143
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Buscar"
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
         Appearance      =   21
         Picture         =   "frmSys_Docs_Remesas.frx":23F6
      End
      Begin XtremeSuiteControls.FlatEdit txtNotas 
         Height          =   735
         Left            =   4440
         TabIndex        =   20
         Top             =   840
         Width           =   7455
         _Version        =   1572864
         _ExtentX        =   13150
         _ExtentY        =   1296
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
         Text            =   "Remesa de Documentación"
         MultiLine       =   -1  'True
         ScrollBars      =   2
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.DateTimePicker dtpR_Inicio 
         Height          =   330
         Left            =   -68680
         TabIndex        =   23
         Top             =   480
         Visible         =   0   'False
         Width           =   1455
         _Version        =   1572864
         _ExtentX        =   2566
         _ExtentY        =   582
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
      Begin XtremeSuiteControls.PushButton btnRemesa 
         Height          =   615
         Index           =   2
         Left            =   -59680
         TabIndex        =   27
         Top             =   6360
         Visible         =   0   'False
         Width           =   1575
         _Version        =   1572864
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
         Picture         =   "frmSys_Docs_Remesas.frx":2AF6
      End
      Begin XtremeSuiteControls.FlatEdit txtConRemesa 
         Height          =   6375
         Left            =   -68320
         TabIndex        =   39
         Top             =   1080
         Visible         =   0   'False
         Width           =   9975
         _Version        =   1572864
         _ExtentX        =   17595
         _ExtentY        =   11245
         _StockProps     =   77
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
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
      Begin XtremeSuiteControls.Label lblRemesaId 
         Height          =   375
         Left            =   -61960
         TabIndex        =   28
         Tag             =   "0"
         Top             =   6480
         Visible         =   0   'False
         Width           =   2175
         _Version        =   1572864
         _ExtentX        =   3836
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Remesa Id: 0"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Transparent     =   -1  'True
         WordWrap        =   -1  'True
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
         Height          =   375
         Left            =   -69880
         TabIndex        =   26
         Top             =   960
         Visible         =   0   'False
         Width           =   11775
         _Version        =   1572864
         _ExtentX        =   20770
         _ExtentY        =   661
         _StockProps     =   14
         Caption         =   "Seleccione una Remesa a Visualizar"
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
      Begin XtremeSuiteControls.Label Label3 
         Height          =   255
         Index           =   6
         Left            =   -69760
         TabIndex        =   25
         Top             =   480
         Visible         =   0   'False
         Width           =   1575
         _Version        =   1572864
         _ExtentX        =   2778
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Fechas:"
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
         WordWrap        =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label3 
         Height          =   255
         Index           =   5
         Left            =   4440
         TabIndex        =   21
         Top             =   600
         Width           =   1575
         _Version        =   1572864
         _ExtentX        =   2778
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Notas:"
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
         WordWrap        =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label3 
         Height          =   255
         Index           =   3
         Left            =   360
         TabIndex        =   18
         Top             =   1200
         Width           =   1575
         _Version        =   1572864
         _ExtentX        =   2778
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Usuario:"
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
         WordWrap        =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label3 
         Height          =   255
         Index           =   4
         Left            =   360
         TabIndex        =   16
         Top             =   480
         Width           =   1575
         _Version        =   1572864
         _ExtentX        =   2778
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Fechas:"
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
         WordWrap        =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label3 
         Height          =   255
         Index           =   1
         Left            =   360
         TabIndex        =   13
         Top             =   840
         Width           =   1575
         _Version        =   1572864
         _ExtentX        =   2778
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Grupos:"
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
         WordWrap        =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label3 
         Height          =   255
         Index           =   2
         Left            =   -69760
         TabIndex        =   7
         Top             =   600
         Visible         =   0   'False
         Width           =   1575
         _Version        =   1572864
         _ExtentX        =   2778
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "No. Documento:"
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
         WordWrap        =   -1  'True
      End
   End
   Begin XtremeSuiteControls.ComboBox cboTipodoc 
      Height          =   345
      Left            =   1800
      TabIndex        =   8
      Top             =   1320
      Width           =   3975
      _Version        =   1572864
      _ExtentX        =   7011
      _ExtentY        =   609
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9.75
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
   Begin XtremeSuiteControls.ProgressBar PrgBar 
      Height          =   135
      Left            =   120
      TabIndex        =   9
      Top             =   1680
      Visible         =   0   'False
      Width           =   12015
      _Version        =   1572864
      _ExtentX        =   21193
      _ExtentY        =   238
      _StockProps     =   93
      BackColor       =   -2147483633
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   8520
      Top             =   1680
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSys_Docs_Remesas.frx":32B2
            Key             =   "IMG1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSys_Docs_Remesas.frx":9B14
            Key             =   "IMG2"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSys_Docs_Remesas.frx":10376
            Key             =   "IMG3"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList3 
      Left            =   8640
      Top             =   2280
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSys_Docs_Remesas.frx":16BD8
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSys_Docs_Remesas.frx":1D43A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSys_Docs_Remesas.frx":23C9C
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSys_Docs_Remesas.frx":23DB6
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSys_Docs_Remesas.frx":23ED4
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSys_Docs_Remesas.frx":2A736
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Remesas de Documentos"
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
      Height          =   480
      Index           =   3
      Left            =   1560
      TabIndex        =   11
      Top             =   360
      Width           =   6252
   End
   Begin XtremeSuiteControls.Label Label3 
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   10
      Top             =   1320
      Width           =   1575
      _Version        =   1572864
      _ExtentX        =   2778
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Tipo Documento"
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
      WordWrap        =   -1  'True
   End
   Begin VB.Image imgBanner 
      Height          =   1092
      Left            =   0
      Top             =   0
      Width           =   13332
   End
End
Attribute VB_Name = "frmSys_Docs_Remesas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem
Dim vPaso As Boolean

Dim mCodigo As String, mTipoDoc As String


Private Sub sbCargaInformacion()

On Error GoTo vError

If cboTipodoc.Text = "TODOS" Then
    strSQL = "SELECT T.COD_TRANSACCION," _
           & "        T.TIPO_DOCUMENTO," _
           & "       ISNULL(T.CLIENTE_IDENTIFICACION, '') AS CLIENTE_IDENTIFICACION," _
           & "       ISNULL(T.CLIENTE_NOMBRE, '') AS CLIENTE_NOMBRE," _
           & "       T.REGISTRO_USUARIO," _
           & "       T.REGISTRO_FECHA" _
           & " FROM SIF_TRANSACCIONES T" _
           & " LEFT JOIN CNT_REMESA_DETALLE R ON T.TIPO_DOCUMENTO = R.TIPO_DOC AND T.COD_TRANSACCION = R.ID_SOLICITUD" _
           & " Where T.ANALISTA_RECEPCION = 1" _
           & "  AND T.ANALISTA_REVISION = 'S'" _
           & "  AND T.TIPO_DOCUMENTO IN ('REA', 'NC', 'ND', 'FND', 'FNC', 'CA', 'RH', 'TCP', 'THAV', 'THCJ', 'TRFA', 'FSL', 'BEAC', 'CBJ', 'TR', 'TRA', 'CD.Liq')" _
           & "  AND T.REGISTRO_FECHA between '" & Format(dtpInicio.Value, "yyyy-mm-dd") & " 00:00:00' and '" & Format(dtpCorte.Value, "yyyy-mm-dd") & " 23:59:59'" _
           & "  AND R.COD_REMESA IS NULL" _
           & " ORDER BY T.TIPO_DOCUMENTO, T.COD_TRANSACCION"
Else
    mTipoDoc = cboTipodoc.ItemData(cboTipodoc.ListIndex)
    
    strSQL = "SELECT T.COD_TRANSACCION," _
           & "        T.TIPO_DOCUMENTO," _
           & "       ISNULL(T.CLIENTE_IDENTIFICACION, '') AS CLIENTE_IDENTIFICACION," _
           & "       ISNULL(T.CLIENTE_NOMBRE, '') AS CLIENTE_NOMBRE," _
           & "       T.REGISTRO_USUARIO," _
           & "       T.REGISTRO_FECHA" _
           & " FROM SIF_TRANSACCIONES T" _
           & " LEFT JOIN CNT_REMESA_DETALLE R ON T.TIPO_DOCUMENTO = R.TIPO_DOC AND T.COD_TRANSACCION = R.ID_SOLICITUD" _
           & " Where T.ANALISTA_RECEPCION = 1" _
           & "  AND T.ANALISTA_REVISION = 'S'" _
           & "  AND T.TIPO_DOCUMENTO = '" & mTipoDoc & "'" _
           & "  AND T.REGISTRO_FECHA between '" & Format(dtpInicio.Value, "yyyy-mm-dd") & " 00:00:00' and '" & Format(dtpCorte.Value, "yyyy-mm-dd") & " 23:59:59'" _
           & "  AND R.COD_REMESA IS NULL" _
           & " ORDER BY T.TIPO_DOCUMENTO, T.COD_TRANSACCION"
    
End If
       
Call OpenRecordSet(rs, strSQL)

lsw.ListItems.Clear

Do While Not rs.EOF
    Set itmX = lsw.ListItems.Add(, , rs!Cod_Transaccion)
    itmX.SubItems(1) = rs!TIPO_DOCUMENTO
    itmX.SubItems(2) = RTrim(rs!CLIENTE_IDENTIFICACION)
    itmX.SubItems(3) = rs!CLIENTE_NOMBRE
    itmX.SubItems(4) = rs!REGISTRO_USUARIO
    itmX.SubItems(5) = Format(rs!REGISTRO_FECHA, "yyyy-mm-dd")
  rs.MoveNext
Loop

rs.Close

Exit Sub
    
vError:
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub sbAplicar()
Dim vRemesa As Long
Dim i As Long, iCasos As Long

On Error GoTo vError

If MsgBox("Está seguro que sea aplicar estas etiquetas y crear Remesa?", vbExclamation + vbYesNo) = vbNo Then
    Exit Sub
End If

'Verifica si hay casos marcados
iCasos = 0
With lsw.ListItems
    For i = 1 To .Count
    
        If .Item(i).Checked Then
            iCasos = iCasos + 1
        End If
    Next i
End With

If iCasos = 0 Then
    MsgBox "No ha seleccionado ningún caso/documento a remesar!", vbExclamation
    Exit Sub
End If

Me.MousePointer = vbHourglass


strSQL = "select isnull(max(COD_REMESA),0) + 1 as 'Remesa' from CNT_REMESAS_DOCS"
Call OpenRecordSet(rs, strSQL)
 vRemesa = rs!Remesa
rs.Close

'Inserta el encabezado de la remesa en CNT_REMESAS_DOCS
strSQL = "insert CNT_REMESAS_DOCS(COD_REMESA,FECHA,USUARIO,NOTAS) values(" & vRemesa & ", dbo.MyGetdate(), '" & glogon.Usuario & "', '" & txtNotas.Text & "')"
Call ConectionExecute(strSQL)


PrgBar.Max = lsw.ListItems.Count + 1
PrgBar.Value = 1
PrgBar.Visible = True


With lsw.ListItems


strSQL = ""
iCasos = 0
For i = 1 To .Count
    
    If .Item(i).Checked Then
        iCasos = iCasos + 1
        strSQL = strSQL & Space(10) & "insert CNT_REMESA_DETALLE(COD_REMESA,id_solicitud,linea,Tipo_Doc)" _
               & " values(" & vRemesa & ",'" & .Item(i).Text & "'," & iCasos & ",'" & .Item(i).SubItems(1) & "'" & ")"

        If Len(strSQL) > 20000 Then
            Call ConectionExecute(strSQL)
            strSQL = ""
        End If
    End If
    
    PrgBar.Value = PrgBar.Value + 1
Next i

.Clear

End With

'Lote Final
If Len(strSQL) > 0 Then
    Call ConectionExecute(strSQL)
    strSQL = ""
End If

PrgBar.Visible = False

Me.MousePointer = vbDefault

MsgBox "Remesa No." & vRemesa & ", creada satisfactoriamente!", vbInformation

Call sbCargaInformacion

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub


Private Sub btnAplicar_Click()
  Call sbAplicar
End Sub

Private Sub btnBuscar_Click()
Call sbCargaInformacion
End Sub

Private Sub btnConsulta_Click()
    Call sbRemesa_Caso_Consulta
End Sub

Private Sub sbRemesa_Exportar()
On Error GoTo vError

Me.MousePointer = vbHourglass

PrgBar.Visible = True

Call Excel_Exportar_Lsw(lswRemesas, PrgBar)

PrgBar.Visible = False

Me.MousePointer = vbDefault

Exit Sub

vError:
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub


Private Sub sbRemesa_Consulta()
On Error GoTo vError

Me.MousePointer = vbHourglass

lblRemesaId.Tag = 0
lblRemesaId.Caption = "<Seleccione una Remesa>"

gbMicrofilm.Visible = False
lswRemesas.Visible = True

strSQL = "select * from CNT_REMESAS_DOCS" _
       & " Where Fecha Between '" & Format(dtpR_Inicio.Value, "yyyy-mm-dd") & " 00:00:00' and '" & Format(dtpR_Corte.Value, "yyyy-mm-dd") & " 23:59:59' order by cod_remesa desc"
Call OpenRecordSet(rs, strSQL)

With lswRemesas.ListItems
  .Clear
  Do While Not rs.EOF
    Set itmX = .Add(, , rs!cod_remesa)
        itmX.SubItems(1) = rs!Usuario & ""
        itmX.SubItems(2) = rs!Fecha & ""
        itmX.SubItems(3) = rs!Notas & ""
        itmX.SubItems(4) = rs!Microfilm_Usuario & ""
        itmX.SubItems(5) = rs!Microfilm_Fecha & ""
    rs.MoveNext
  Loop
  rs.Close

End With

Me.MousePointer = vbDefault

Exit Sub

vError:
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub btnMicrofilm_Click(Index As Integer)

Select Case Index
  Case 0 'Aplicar
  
     If btnMicrofilm(0).Enabled Then
                 
         strSQL = "exec spCNT_Actualiza_Remesa " & txtReciboRemesa.Text & ", '" & glogon.Usuario & "'"
         Call ConectionExecute(strSQL)
         
         MsgBox "Recibo ( Microfilm: Archivo Digital ) Satisfactoriamente...!", vbInformation
         Call sbRemesa_Consulta
     Else
        MsgBox "No tiene los permisos para realizar esta opción, verifique...!!!", vbExclamation
     End If
     
  Case 1 'Cancelar
    'Nada
End Select

gbMicrofilm.Visible = False
lswRemesas.Visible = True
End Sub



Private Sub sbRemesa_Reporte()
Dim vTitulo As String, vSubTitulo As String, vFiltro As String

On Error GoTo vError

If Not IsNumeric(lblRemesaId.Tag) Then Exit Sub

Me.MousePointer = vbHourglass


vSubTitulo = ""
vFiltro = ""

With frmContenedor.Crt
 .Reset
 .WindowShowGroupTree = True
 .WindowShowPrintSetupBtn = True
 .WindowShowRefreshBtn = True
 .WindowShowSearchBtn = True
 .WindowState = crptMaximized
 .WindowTitle = "Reportes del Control de Documentos"

 .Connect = glogon.ConectRPT


  vSubTitulo = "REMESA : " & lblRemesaId.Tag & " LISTADO : DETALLADO"

  .ReportFileName = SIFGlobal.fxPathReportes("Sys_RemesasDocumentos_Detalle.rpt")


 .Formulas(0) = "fxFecha='FECHA: " & Format(fxFechaServidor, "dd/mm/yyyy") & "'"
 .Formulas(1) = "fxEmpresa='" & GLOBALES.gstrNombreEmpresa & "'"
 .Formulas(2) = "fxUsuario='USER: " & glogon.Usuario & "'"
 .Formulas(3) = "fxTitulo='REMESA TRASLADO MICROFILM : DOCUMENTOS'"
 .Formulas(4) = "fxSubTitulo='" & vSubTitulo & "'"
' .SelectionFormula = "{AFI_REMESAS_ING.COD_REMESA} = " & lblRemesa.Tag

 .StoredProcParam(0) = lblRemesaId.Tag
 .PrintReport

End With

Me.MousePointer = vbDefault
Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub sbRemesa_Caso_Consulta()

On Error GoTo vError

txtConRemesa.Text = ""

strSQL = "exec spCNT_Consulta_Documento '" & txtCodigoBuscar.Text & "'"

Call OpenRecordSet(rs, strSQL)

If rs.EOF Then txtConRemesa.Text = "** No se encontró Caso en las remesas registradas **"


Do While Not rs.EOF
 txtConRemesa.Text = txtConRemesa.Text & vbCrLf _
                                  & "Remesa Id      " & vbTab & "...: " & rs!Remesa & vbCrLf
 txtConRemesa.Text = txtConRemesa & "Transacción    " & vbTab & "...: " & rs!TipoDoc & vbCrLf
 txtConRemesa.Text = txtConRemesa & "No. Documento  " & vbTab & "...: " & rs!Transaccion & vbCrLf
 txtConRemesa.Text = txtConRemesa & "Fecha          " & vbTab & "...: " & rs!Fecha & vbCrLf
 txtConRemesa.Text = txtConRemesa & "Usuario        " & vbTab & "...: " & rs!Usuario & vbCrLf
 txtConRemesa.Text = txtConRemesa & "Archivo Fecha  " & vbTab & "...: " & rs!FechaRecibeA & vbCrLf
 txtConRemesa.Text = txtConRemesa & "Archivo Usuario" & vbTab & "...: " & rs!UsuarioArchivo & vbCrLf
 
 rs.MoveNext
Loop
rs.Close

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
 txtConRemesa.Text = ""

End Sub

Private Sub btnRemesa_Click(Index As Integer)

Select Case Index
    Case 0 'Consulta
        Call sbRemesa_Consulta
    
    Case 1 'Exportar
        Call sbRemesa_Exportar

    Case 2 'Reporte
        Call sbRemesa_Reporte

End Select

End Sub

Private Sub cboTipodoc_Click()

If vPaso Then Exit Sub

Select Case tcMain.SelectedItem
    Case 0 'Remesas
        Call sbCargaInformacion
        
    Case 1 'Reportes
        Call sbRemesa_Consulta
        
    Case 2 'Consulta
End Select

End Sub



Private Sub Form_Activate()
vModulo = 8
End Sub

Private Sub Form_Load()

vModulo = 8
   

Set imgBanner.Picture = frmContenedor.imgBanner_Mantenimiento.Picture

With lsw.ColumnHeaders
    .Clear
    .Add , , "Código", 1800
    .Add , , "Tipo", 1800
    .Add , , "Identificación", 1800, vbCenter
    .Add , , "Nombre", 4500
    .Add , , "Usuario", 2800, vbCenter
    .Add , , "Fecha", 1800
End With


With lswRemesas.ColumnHeaders
    .Clear
    .Add , , "Remesa Id", 1200
    .Add , , "Usuario", 2500, vbCenter
    .Add , , "Fecha", 1800
    .Add , , "Notas", 1800
    .Add , , "Archivo Usuario", 2500, vbCenter
    .Add , , "Archivo Fecha", 2100, vbCenter
End With

tcMain.Item(0).Selected = True

dtpInicio.Value = fxFechaServidor
dtpCorte.Value = dtpInicio.Value
dtpInicio.Value = DateAdd("m", -2, dtpInicio.Value)


dtpR_Inicio.Value = dtpInicio.Value
dtpR_Corte.Value = dtpCorte.Value

vPaso = True
    strSQL = "select rtrim(Tipo_Documento) as IdX, rtrim(Descripcion) as 'Itmx'" _
           & " from SIF_Documentos" _
           & " where Tipo_documento in('NC','ND','FND','FNC','CA', 'CD.Liq', 'BEAC', 'CBJ', 'FSL', 'REA', 'RH', 'TCP', 'TRFA', 'TCP', 'THCJ', 'TRA', 'THAV')" _
           & " order by Descripcion"
    Call sbCbo_Llena_New(cboTipodoc, strSQL, True, True)
vPaso = False

Call sbGrupos_Load


Call Formularios(Me)
Call RefrescaTags(Me)

Call cboTipodoc_Click

End Sub


Private Sub lsw_ColumnClick(ByVal ColumnHeader As XtremeSuiteControls.ListViewColumnHeader)
 lsw.SortKey = ColumnHeader.Index - 1
  If lsw.SortOrder = 0 Then lsw.SortOrder = 1 Else lsw.SortOrder = 0
  lsw.Sorted = True
End Sub


Private Sub lswRemesas_DblClick()

If lswRemesas.ListItems.Count <= 0 Then Exit Sub

If Len(lswRemesas.SelectedItem.SubItems(4)) > 0 Then Exit Sub
    lswRemesas.Visible = False
    gbMicrofilm.Visible = True
    txtReciboRemesa.Text = lswRemesas.SelectedItem
    txtReciboUsuario.Text = lswRemesas.SelectedItem.SubItems(4)

End Sub

Private Sub lswRemesas_ItemClick(ByVal Item As XtremeSuiteControls.ListViewItem)
   lblRemesaId.Tag = Item.Text
   lblRemesaId.Caption = "Remesa Id: " & Item.Text
End Sub

Private Sub tcMain_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)

Select Case Item.Index
    Case 0 'Recepcion
        Call sbCargaInformacion
        
    Case 1 'Reportes
        Call sbRemesa_Consulta
        
    Case 2 'Consulta
        txtCodigoBuscar.Text = ""
        txtConRemesa.Text = ""
    
End Select
End Sub


Private Sub sbUsuarios_Load()
If vPaso Then Exit Sub

On Error GoTo vError
    
    Me.MousePointer = vbHourglass
    If cboGrupo.Text = "TODOS" Then
        strSQL = "SELECT UPPER(USUARIO) as 'ItmX', Usuario as 'IdX'" _
               & " from CRD_GRPUSERS group by Usuario order by Usuario"
    Else
        strSQL = "SELECT UPPER(USUARIO) as 'ItmX', Usuario as 'IdX'" _
               & " from CRD_GRPUSERS WHERE COD_GRUPO = '" & cboGrupo.ItemData(cboGrupo.ListIndex) _
               & "' order by Usuario"
    End If
    Call sbCbo_Llena_New(cboUsuario, strSQL, True, True)

    Me.MousePointer = vbDefault
    
Exit Sub

vError:
    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub sbGrupos_Load()

On Error GoTo vError
    
    Me.MousePointer = vbHourglass

    strSQL = "select COD_GRUPO as 'IdX', DESCRIPCION as 'ItmX' from CRD_GRUPOS"
    
    vPaso = True
        Call sbCbo_Llena_New(cboGrupo, strSQL, True, True)
    vPaso = False
    
    Me.MousePointer = vbDefault
    
    Call sbUsuarios_Load
    
Exit Sub

vError:
    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub
