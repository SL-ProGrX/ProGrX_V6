VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.Controls.v24.0.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.ShortcutBar.v24.0.0.ocx"
Begin VB.Form frmFNDTraspasoTesoreria 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Retiros/Liq Ahorros: Traslado a Bancos"
   ClientHeight    =   7050
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   17070
   Icon            =   "frmFNDTraspasoTesoreria.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7050
   ScaleWidth      =   17070
   WindowState     =   2  'Maximized
   Begin XtremeSuiteControls.GroupBox fraFiltros 
      Height          =   4215
      Left            =   9600
      TabIndex        =   29
      Top             =   1560
      Visible         =   0   'False
      Width           =   6135
      _Version        =   1572864
      _ExtentX        =   10821
      _ExtentY        =   7435
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      Begin XtremeSuiteControls.PushButton btnFiltros 
         Height          =   375
         Index           =   0
         Left            =   3360
         TabIndex        =   31
         Top             =   3480
         Width           =   1215
         _Version        =   1572864
         _ExtentX        =   2138
         _ExtentY        =   656
         _StockProps     =   79
         Caption         =   "Buscar"
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
         UseVisualStyle  =   -1  'True
         Appearance      =   17
      End
      Begin XtremeSuiteControls.ComboBox cboBanco 
         Height          =   315
         Left            =   1320
         TabIndex        =   32
         Top             =   600
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
      Begin XtremeSuiteControls.ComboBox cboOficina 
         Height          =   315
         Left            =   1320
         TabIndex        =   33
         Top             =   1080
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
      Begin XtremeSuiteControls.ComboBox cboUsuarios 
         Height          =   315
         Left            =   1320
         TabIndex        =   34
         Top             =   1560
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
      Begin XtremeSuiteControls.ComboBox cboSistema 
         Height          =   315
         Left            =   1320
         TabIndex        =   35
         Top             =   2040
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
      Begin XtremeSuiteControls.PushButton btnFiltros 
         Height          =   375
         Index           =   1
         Left            =   4560
         TabIndex        =   36
         Top             =   3480
         Width           =   1215
         _Version        =   1572864
         _ExtentX        =   2138
         _ExtentY        =   656
         _StockProps     =   79
         Caption         =   "Refrescar"
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
         UseVisualStyle  =   -1  'True
         Appearance      =   17
      End
      Begin XtremeSuiteControls.CheckBox chkMarcas 
         Height          =   375
         Left            =   1320
         TabIndex        =   37
         Top             =   3240
         Visible         =   0   'False
         Width           =   1335
         _Version        =   1572864
         _ExtentX        =   2350
         _ExtentY        =   656
         _StockProps     =   79
         Caption         =   "Solo Casos Marcados"
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
         Appearance      =   16
      End
      Begin XtremeSuiteControls.ComboBox cboTokenConsulta 
         Height          =   315
         Left            =   1320
         TabIndex        =   38
         Top             =   2520
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
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Banco"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   43
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Oficina"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   240
         TabIndex        =   42
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Usuarios"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   240
         TabIndex        =   41
         Top             =   1560
         Width           =   975
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Sistema"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   240
         TabIndex        =   40
         Top             =   2040
         Width           =   975
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Token"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   4
         Left            =   240
         TabIndex        =   39
         Top             =   2520
         Width           =   975
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
         Height          =   375
         Index           =   0
         Left            =   0
         TabIndex        =   30
         Top             =   0
         Width           =   6135
         _Version        =   1572864
         _ExtentX        =   10821
         _ExtentY        =   661
         _StockProps     =   14
         Caption         =   "Filtros adicionalaes:"
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
   Begin XtremeSuiteControls.GroupBox gbReportes 
      Height          =   2895
      Left            =   3840
      TabIndex        =   22
      Top             =   1560
      Visible         =   0   'False
      Width           =   5655
      _Version        =   1572864
      _ExtentX        =   9975
      _ExtentY        =   5106
      _StockProps     =   79
      BackColor       =   16777215
      Appearance      =   16
      Begin XtremeSuiteControls.PushButton btnInforme 
         Height          =   495
         Index           =   0
         Left            =   3600
         TabIndex        =   23
         Top             =   2040
         Width           =   1455
         _Version        =   1572864
         _ExtentX        =   2566
         _ExtentY        =   873
         _StockProps     =   79
         Caption         =   "Informe"
         BackColor       =   -2147483633
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
         Picture         =   "frmFNDTraspasoTesoreria.frx":000C
      End
      Begin XtremeSuiteControls.RadioButton rbInformes 
         Height          =   255
         Index           =   0
         Left            =   600
         TabIndex        =   24
         Top             =   600
         Width           =   4335
         _Version        =   1572864
         _ExtentX        =   7646
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Resumen"
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
         TabIndex        =   25
         Top             =   960
         Width           =   4335
         _Version        =   1572864
         _ExtentX        =   7646
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Detalle"
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
         TabIndex        =   26
         Top             =   1320
         Width           =   4335
         _Version        =   1572864
         _ExtentX        =   7646
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Pendientes"
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
         Left            =   5040
         TabIndex        =   27
         Top             =   2040
         Width           =   495
         _Version        =   1572864
         _ExtentX        =   873
         _ExtentY        =   873
         _StockProps     =   79
         BackColor       =   -2147483633
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
         Picture         =   "frmFNDTraspasoTesoreria.frx":0713
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
         Height          =   375
         Index           =   1
         Left            =   0
         TabIndex        =   28
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
   Begin XtremeSuiteControls.CheckBox chkFiltros 
      Height          =   255
      Left            =   9600
      TabIndex        =   10
      Top             =   1200
      Width           =   1335
      _Version        =   1572864
      _ExtentX        =   2355
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "+ Filtros"
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
      Appearance      =   16
   End
   Begin XtremeSuiteControls.ComboBox cboAccion 
      Height          =   312
      Left            =   2160
      TabIndex        =   5
      Top             =   240
      Width           =   1932
      _Version        =   1572864
      _ExtentX        =   3413
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
   Begin XtremeSuiteControls.DateTimePicker dtpDesde 
      Height          =   312
      Left            =   6120
      TabIndex        =   6
      Top             =   600
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
   Begin XtremeSuiteControls.DateTimePicker dtpHasta 
      Height          =   312
      Left            =   8160
      TabIndex        =   7
      Top             =   600
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
   Begin XtremeSuiteControls.ComboBox cboEstado 
      Height          =   312
      Left            =   2160
      TabIndex        =   8
      Top             =   600
      Width           =   1932
      _Version        =   1572864
      _ExtentX        =   3413
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
   Begin XtremeSuiteControls.ComboBox cboTipo 
      Height          =   312
      Left            =   6120
      TabIndex        =   9
      Top             =   240
      Width           =   3372
      _Version        =   1572864
      _ExtentX        =   5953
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
   Begin XtremeSuiteControls.CheckBox chkTodos 
      Height          =   255
      Left            =   720
      TabIndex        =   11
      Top             =   1200
      Width           =   1335
      _Version        =   1572864
      _ExtentX        =   2355
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Marcar"
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
   End
   Begin XtremeSuiteControls.ComboBox cboToken 
      Height          =   330
      Left            =   9720
      TabIndex        =   16
      Top             =   600
      Width           =   2175
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
   Begin XtremeSuiteControls.PushButton btnTokenNew 
      Height          =   315
      Index           =   2
      Left            =   12000
      TabIndex        =   18
      Top             =   600
      Width           =   855
      _Version        =   1572864
      _ExtentX        =   1508
      _ExtentY        =   556
      _StockProps     =   79
      Caption         =   "Nuevo"
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
      Appearance      =   6
   End
   Begin XtremeSuiteControls.PushButton btnBarra 
      Height          =   375
      Index           =   0
      Left            =   4560
      TabIndex        =   19
      Top             =   1125
      Width           =   1215
      _Version        =   1572864
      _ExtentX        =   2138
      _ExtentY        =   656
      _StockProps     =   79
      Caption         =   "Buscar"
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
      UseVisualStyle  =   -1  'True
      TextAlignment   =   1
      Appearance      =   17
      Picture         =   "frmFNDTraspasoTesoreria.frx":0E29
      ImageAlignment  =   4
   End
   Begin XtremeSuiteControls.PushButton btnBarra 
      Height          =   375
      Index           =   1
      Left            =   7080
      TabIndex        =   20
      Top             =   1125
      Width           =   1215
      _Version        =   1572864
      _ExtentX        =   2138
      _ExtentY        =   656
      _StockProps     =   79
      Caption         =   "Generar"
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
      UseVisualStyle  =   -1  'True
      TextAlignment   =   1
      Appearance      =   17
      Picture         =   "frmFNDTraspasoTesoreria.frx":1529
      ImageAlignment  =   4
   End
   Begin XtremeSuiteControls.PushButton btnBarra 
      Height          =   375
      Index           =   2
      Left            =   8280
      TabIndex        =   21
      Top             =   1125
      Width           =   1215
      _Version        =   1572864
      _ExtentX        =   2138
      _ExtentY        =   656
      _StockProps     =   79
      Caption         =   "Informes"
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
      UseVisualStyle  =   -1  'True
      TextAlignment   =   1
      Appearance      =   17
      Picture         =   "frmFNDTraspasoTesoreria.frx":1C50
      ImageAlignment  =   4
   End
   Begin XtremeSuiteControls.PushButton btnBarra 
      Height          =   375
      Index           =   3
      Left            =   5760
      TabIndex        =   44
      Top             =   1125
      Width           =   1215
      _Version        =   1572864
      _ExtentX        =   2138
      _ExtentY        =   656
      _StockProps     =   79
      Caption         =   "Exportar"
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
      UseVisualStyle  =   -1  'True
      TextAlignment   =   1
      Appearance      =   17
      Picture         =   "frmFNDTraspasoTesoreria.frx":2357
      ImageAlignment  =   4
   End
   Begin XtremeSuiteControls.ProgressBar ProgressBarX 
      Height          =   135
      Left            =   4560
      TabIndex        =   45
      Top             =   1440
      Visible         =   0   'False
      Width           =   2415
      _Version        =   1572864
      _ExtentX        =   4260
      _ExtentY        =   238
      _StockProps     =   93
      BackColor       =   -2147483633
      Scrolling       =   1
   End
   Begin XtremeSuiteControls.ProgressBar prgBar 
      Height          =   135
      Left            =   0
      TabIndex        =   46
      Top             =   6960
      Width           =   1935
      _Version        =   1572864
      _ExtentX        =   3413
      _ExtentY        =   238
      _StockProps     =   93
      BackColor       =   -2147483633
      Value           =   5
   End
   Begin XtremeSuiteControls.TabControl tcMain 
      Height          =   5292
      Left            =   0
      TabIndex        =   12
      Top             =   1560
      Width           =   13692
      _Version        =   1572864
      _ExtentX        =   24151
      _ExtentY        =   9334
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
      Item(0).Caption =   "Listado General"
      Item(0).ControlCount=   1
      Item(0).Control(0)=   "vGrid"
      Item(1).Caption =   "Casos Unificados"
      Item(1).ControlCount=   2
      Item(1).Control(0)=   "lsw"
      Item(1).Control(1)=   "lswDet"
      Begin XtremeSuiteControls.ListView lsw 
         Height          =   3012
         Left            =   -69880
         TabIndex        =   14
         Top             =   360
         Visible         =   0   'False
         Width           =   2892
         _Version        =   1572864
         _ExtentX        =   5101
         _ExtentY        =   5313
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
         Appearance      =   16
         ShowBorder      =   0   'False
      End
      Begin XtremeSuiteControls.ListView lswDet 
         Height          =   1812
         Left            =   -69880
         TabIndex        =   15
         Top             =   3480
         Visible         =   0   'False
         Width           =   4092
         _Version        =   1572864
         _ExtentX        =   7218
         _ExtentY        =   3196
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
         Appearance      =   16
         ShowBorder      =   0   'False
      End
      Begin FPSpreadADO.fpSpread vGrid 
         Height          =   4695
         Left            =   120
         TabIndex        =   13
         Top             =   480
         Width           =   16935
         _Version        =   524288
         _ExtentX        =   29871
         _ExtentY        =   8281
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
         MaxCols         =   15
         SpreadDesigner  =   "frmFNDTraspasoTesoreria.frx":2C28
         VScrollSpecial  =   -1  'True
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Token para Bancos:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   9720
      TabIndex        =   17
      Top             =   360
      Width           =   1575
   End
   Begin VB.Label lblTipoCaso 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo de Caso"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   252
      Left            =   4440
      TabIndex        =   4
      Top             =   240
      Width           =   1572
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Estado"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   2
      Left            =   1320
      TabIndex        =   3
      Top             =   600
      Width           =   615
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Acción"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   1
      Left            =   1320
      TabIndex        =   2
      Top             =   240
      Width           =   615
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Hasta"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   7560
      TabIndex        =   1
      Top             =   600
      Width           =   495
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Desde"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   0
      Left            =   5520
      TabIndex        =   0
      Top             =   600
      Width           =   615
   End
   Begin VB.Image imgBanner 
      Height          =   1092
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   15732
   End
End
Attribute VB_Name = "frmFNDTraspasoTesoreria"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vPaso As Boolean
Dim vDuplicado As Boolean



Private Sub sbExportar()

Dim vHeaders As vGridHeaders

On Error GoTo vError

Me.MousePointer = vbHourglass

ProgressBarX.Visible = True

Select Case tcMain.SelectedItem
    Case 0
    ' Cedula, Nombre, Plan, No. Contrato, Monto, Usuario, Oficina, Tipo Emision, Cuenta, Banco, Fecha Registro, Duplicado?, Token

            vHeaders.Columnas = 15
            vHeaders.Headers(1) = "..."
            vHeaders.Headers(2) = "# RetLiq"
            vHeaders.Headers(3) = "Cédula"
            vHeaders.Headers(4) = "Nombre"
            vHeaders.Headers(5) = "Plan"
            vHeaders.Headers(6) = "No. Contrato"
            vHeaders.Headers(7) = "Monto"
            vHeaders.Headers(8) = "Oficina"
            vHeaders.Headers(9) = "Emite?"
            vHeaders.Headers(10) = "Cuenta"
            vHeaders.Headers(11) = "Banco"
            vHeaders.Headers(12) = "Fecha Registro"
            vHeaders.Headers(13) = "Usuario Registro"
            vHeaders.Headers(14) = "Duplicado?"
            vHeaders.Headers(15) = "Token"
        
         Call sbSIFGridExportar(vGrid, vHeaders, "ProGrX_Fondos_RetLiq_Traslado_Bancos")
    
    Case 1
        Call Excel_Exportar_Lsw(lsw, ProgressBarX)
End Select

ProgressBarX.Visible = False

Me.MousePointer = vbDefault

Exit Sub

vError:
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub btnBarra_Click(Index As Integer)


Select Case Index
  Case 0 'Buscar
    
    If tcMain.SelectedItem = 0 Then
        Call sbBuscar
    Else
        Call sbRevisaDuplicadosEnLaRemesa
    End If
    
  Case 1 'Generar
    If tcMain.SelectedItem = 0 Then
        Call sbGenerar
    Else
        Call sbGenerar_Unificado
    End If

   Case 2 'Informes
   
   gbReportes.top = tcMain.top
   gbReportes.Left = 3840
   gbReportes.Visible = IIf((gbReportes.Visible = True), False, True)
   
   
   Case 3 'Exportar
    Call sbExportar
End Select

End Sub

Private Sub btnFiltros_Click(Index As Integer)
Select Case Index
  Case 0  'Buscar
    Call sbBuscar
    fraFiltros.Visible = False
    chkFiltros.Value = vbUnchecked
  
  Case 1 'Refrescar
    Call sbFiltros
    
End Select
End Sub

Private Sub btnInforme_Click(Index As Integer)

Select Case Index
Case 0 'Informe
    Select Case True
        Case rbInformes.Item(0).Value
            Call sbInformes("Resumen")
        Case rbInformes.Item(1).Value
            Call sbInformes("Detalle")
        Case rbInformes.Item(2).Value
            Call sbInformes("Pendientes")
    End Select

Case 1 'Cerrar
    gbReportes.Visible = False

End Select


End Sub

Private Sub btnTokenNew_Click(Index As Integer)
Dim strSQL As String

On Error GoTo vError

strSQL = "exec spTes_Token_New '" & glogon.Usuario & "'"
Call ConectionExecute(strSQL)

Call sbTokens_Load

Exit Sub

vError:


End Sub

Private Sub cboAccion_Click()
Dim strSQL As String

If vPaso Then Exit Sub

vGrid.MaxRows = 0


vPaso = True

If Mid(cboAccion.Text, 1, 1) = "D" Then
  'D = Desembolsar
  strSQL = "select id_banco as 'Idx', rtrim(descripcion) + '  ' + rtrim(Cta) as 'ItmX' from Tes_Bancos where estado = 'A'"
  Call sbCbo_Llena_New(cboTipo, strSQL, True, True)
  
  lblTipoCaso.Caption = "Bancos...:"
  cboTipo.Enabled = False
  lblTipoCaso.Enabled = False
Else
  'R = Retener
  strSQL = "select RTRIM(RETENCION_CODIGO) as 'IdX', rtrim(descripcion) + ' [' + rtrim(COD_CUENTA) + ']' as 'ItmX'" _
         & " from FND_RETENCION_CONCEPTOS where ACTIVO = 1"
  Call sbCbo_Llena_New(cboTipo, strSQL, False, True)
  
  lblTipoCaso.Caption = "Retener por.:"
  cboTipo.Enabled = True
  lblTipoCaso.Enabled = True

End If

vPaso = False

End Sub


Private Sub cboEstado_Click()

If vPaso Then Exit Sub

vGrid.MaxRows = 0

End Sub


Private Sub cboTipo_Click()
If vPaso Then Exit Sub

If Mid(cboAccion.Text, 1, 1) = "D" Then
    vGrid.MaxRows = 0
End If

End Sub

Private Sub chkFiltros_Click()
If chkFiltros.Value = vbChecked Then
   fraFiltros.Visible = True
   fraFiltros.top = tcMain.top
   fraFiltros.Left = chkFiltros.Left
   Call sbFiltros
Else
   fraFiltros.Visible = False
End If
End Sub

Private Sub chkTodos_Click()
Dim i As Long

If tcMain.SelectedItem = 0 Then
    For i = 1 To vGrid.MaxRows
     vGrid.Row = i
     vGrid.Col = 1
     vGrid.Value = chkTodos.Value
    Next i
Else
    For i = 1 To lsw.ListItems.Count
        lsw.ListItems.Item(i).Checked = chkTodos.Value
    Next i
End If

End Sub

Private Sub sbBuscar()
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'OBJETIVO:      Busca Liquidaciones pendientes o generadas con ubicacion en tesoreria.
'REFERENCIAS:   Ninguna.
'OBSERVACIONES: Ninguna.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim strSQL As String

If dtpHasta.Value < dtpDesde.Value Then
   MsgBox "Verifique el Rango de Fechas", vbInformation, "Error"
   Exit Sub
End If

Me.MousePointer = vbHourglass


strSQL = "exec spFND_TrasladoTesoreria_Fix"
Call ConectionExecute(strSQL)


strSQL = "Select " & chkTodos.Value & " as 'Valor',L.Consec,C.Cedula,S.nombre,L.Cod_Plan,L.Cod_Contrato" _
        & ",case when L.Total_Girar is null then L.Aportes_Liq+L.Rendi_Liq - (isnull(L.multa_retiro,0) + isnull(L.ISR_MONTO,0) + isnull(L.OTROS_REBAJOS,0)) else L.Total_Girar end as 'Total_Girar'" _
        & ",L.Usuario,isnull(L.cod_Oficina,'') as 'Oficina',L.Tipo,L.Cta_Ahorros,B.descripcion,L.Fecha" _
        & ",dbo.fxTesSupervisa(C.cedula,S.nombre,isnull(L.Total_Girar,L.Aportes_Liq+L.Rendi_Liq - isnull(L.multa_retiro,0)),0,'C') as 'Duplicado',TES_SUPERVISION_FECHA" _
        & ",L.PAGO_TERCERO_APL,L.PAGO_TERCERO_TIPO,L.PAGO_TERCERO_ID,L.PAGO_TERCERO_NOMBRE, L.ID_TOKEN" _
        & " From Fnd_Liquidacion L inner join Fnd_Contratos C on L.Cod_Operadora=C.Cod_Operadora " _
        & " and L.Cod_Plan = C.Cod_Plan and L.Cod_Contrato = C.Cod_Contrato" _
        & " inner join Socios S on C.cedula = S.cedula" _
        & " left join Tes_Bancos B on L.cod_Banco = B.id_Banco" _
        & " Where L.Fecha between '" & Format(dtpDesde.Value, "yyyy/mm/dd") & " 00:00:00' and '" _
        & Format(dtpHasta.Value, "yyyy/mm/dd") & " 23:59:59'"
        

If fxSIFParametros("17") = "S" Then
  strSQL = strSQL & " And L.Analista_Revision = 'S'"
End If

If Mid(cboEstado.Text, 1, 1) = "P" Then
  strSQL = strSQL & " And L.Traspaso_tesoreria is Null"
Else
  strSQL = strSQL & " And L.Traspaso_tesoreria is not Null"
End If

'If Mid(cboAccion.Text, 1, 1) = "D" And cboTipo.Text <> "TODOS" Then
'  strSQL = strSQL & " And L.cod_Banco = " & cboTipo.ItemData(cboTipo.ListIndex)
'End If


If chkFiltros.Value = vbChecked Then
    If cboBanco.Text <> "TODOS" Then
      strSQL = strSQL & " And L.cod_banco = " & cboBanco.ItemData(cboBanco.ListIndex)
    End If

    If cboOficina.Text <> "TODOS" Then
      strSQL = strSQL & " And L.cod_oficina = '" & cboOficina.ItemData(cboOficina.ListIndex) & "'"
    End If

    If cboUsuarios.Text <> "TODOS" Then
      strSQL = strSQL & " And L.usuario like '" & cboUsuarios.Text & "%'"
    End If

    If cboSistema.Text <> "TODOS" Then
      strSQL = strSQL & " And isnull(L.cod_app,'" & App.ProductName & "') like '" & cboSistema.Text & "%'"
    End If

    If cboTokenConsulta.Text <> "TODOS" Then
      strSQL = strSQL & " And isnull(L.ID_Token,'') like '" & cboTokenConsulta.Text & "%'"
    End If

End If

Call sbCargaGridLocal(15, strSQL)

If vGrid.MaxRows > 0 Then vGrid.MaxRows = vGrid.MaxRows - 1

Me.MousePointer = vbDefault

Call sbRevisaDuplicadosEnLaRemesa

End Sub


Private Sub sbPersona_Detalle(pCedula As String)
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'OBJETIVO:      Busca Liquidaciones pendientes o generadas con ubicacion en tesoreria.
'REFERENCIAS:   Ninguna.
'OBSERVACIONES: Ninguna.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem

Me.MousePointer = vbHourglass


strSQL = "Select L.Consec,C.Cedula,S.nombre,L.Cod_Plan,L.Cod_Contrato" _
        & ",case when L.Total_Girar is null then L.Aportes_Liq+L.Rendi_Liq - (isnull(L.multa_retiro,0) + isnull(L.ISR_MONTO,0) + isnull(L.OTROS_REBAJOS,0)) else L.Total_Girar end as 'Total_Girar'" _
        & ",L.Usuario,isnull(L.cod_Oficina,'') as 'Oficina',L.Tipo,L.Cta_Ahorros,B.descripcion,L.Fecha" _
        & ",dbo.fxTesSupervisa(C.cedula,S.nombre,isnull(L.Total_Girar,L.Aportes_Liq+L.Rendi_Liq - isnull(L.multa_retiro,0)),0,'C') as 'Duplicado',TES_SUPERVISION_FECHA" _
        & ",L.PAGO_TERCERO_APL,L.PAGO_TERCERO_TIPO,L.PAGO_TERCERO_ID,L.PAGO_TERCERO_NOMBRE" _
        & " From Fnd_Liquidacion L inner join Fnd_Contratos C on L.Cod_Operadora=C.Cod_Operadora " _
        & " and L.Cod_Plan = C.Cod_Plan and L.Cod_Contrato = C.Cod_Contrato" _
        & " inner join Socios S on C.cedula = S.cedula" _
        & " left join Tes_Bancos B on L.cod_Banco = B.id_Banco" _
        & " Where L.Fecha between '" & Format(dtpDesde.Value, "yyyy/mm/dd") & " 00:00:00' and '" _
        & Format(dtpHasta.Value, "yyyy/mm/dd") & " 23:59:59'" _
        & " and C.cedula = '" & pCedula & "'"
        

If fxSIFParametros("17") = "S" Then
  strSQL = strSQL & " And L.Analista_Revision = 'S'"
End If

If Mid(cboEstado.Text, 1, 1) = "P" Then
  strSQL = strSQL & " And L.Traspaso_tesoreria is Null"
Else
  strSQL = strSQL & " And L.Traspaso_tesoreria is not Null"
End If

'If Mid(cboAccion.Text, 1, 1) = "D" And cboTipo.Text <> "TODOS" Then
'  strSQL = strSQL & " And L.cod_Banco = " & cboTipo.ItemData(cboTipo.ListIndex)
'End If


If chkFiltros.Value = vbChecked Then
    If cboBanco.Text <> "TODOS" Then
      strSQL = strSQL & " And L.cod_banco = " & cboBanco.ItemData(cboBanco.ListIndex)
    End If

    If cboOficina.Text <> "TODOS" Then
      strSQL = strSQL & " And L.cod_oficina = '" & cboOficina.ItemData(cboOficina.ListIndex) & "'"
    End If

    If cboUsuarios.Text <> "TODOS" Then
      strSQL = strSQL & " And L.usuario like '" & cboUsuarios.Text & "%'"
    End If

    If cboSistema.Text <> "TODOS" Then
      strSQL = strSQL & " And isnull(L.cod_app,'" & App.ProductName & "') like '" & cboSistema.Text & "%'"
    End If

    If cboTokenConsulta.Text <> "TODOS" Then
      strSQL = strSQL & " And isnull(L.ID_Token,'') like '" & cboTokenConsulta.Text & "%'"
    End If

End If

Call OpenRecordSet(rs, strSQL)

lswDet.ListItems.Clear

Do While Not rs.EOF
 Set itmX = lswDet.ListItems.Add(, , rs!consec)
     itmX.SubItems(1) = rs!Cod_Plan
     itmX.SubItems(2) = rs!COD_Contrato
     itmX.SubItems(3) = rs!Cedula
     itmX.SubItems(4) = rs!Nombre
     itmX.SubItems(5) = rs!TOTAL_GIRAR
     itmX.SubItems(6) = rs!Cta_Ahorros
     itmX.SubItems(7) = rs!Descripcion
 rs.MoveNext
Loop
rs.Close

Me.MousePointer = vbDefault


End Sub



Private Sub sbRevisaDuplicadosEnLaRemesa()
Dim strSQL As String, rs As New ADODB.Recordset
Dim i As Long, vCasos As Boolean
Dim itmX As ListViewItem

Me.MousePointer = vbHourglass


strSQL = "Select count(*) as 'Liquidaciones', C.Cedula,S.nombre,L.Cta_Ahorros,B.descripcion" _
        & ",sum(case when L.Total_Girar is null then L.Aportes_Liq+L.Rendi_Liq - (isnull(L.multa_retiro,0) + isnull(L.ISR_MONTO,0) + isnull(L.OTROS_REBAJOS,0)) else L.Total_Girar end) as 'Total_Girar'" _
        & " From Fnd_Liquidacion L" _
        & "     inner join Fnd_Contratos C on L.Cod_Operadora=C.Cod_Operadora " _
        & "             and L.Cod_Plan = C.Cod_Plan and L.Cod_Contrato = C.Cod_Contrato" _
        & "     inner join Socios S on C.cedula = S.cedula" _
        & "      left join Tes_Bancos B on L.cod_Banco = B.id_Banco" _
        & " Where L.Fecha between '" & Format(dtpDesde.Value, "yyyy/mm/dd") & " 00:00:00' and '" _
        & Format(dtpHasta.Value, "yyyy/mm/dd") & " 23:59:59'"
        

If fxSIFParametros("17") = "S" Then
  strSQL = strSQL & " And L.Analista_Revision = 'S'"
End If

If Mid(cboEstado.Text, 1, 1) = "P" Then
  strSQL = strSQL & " And L.Traspaso_tesoreria is Null"
Else
  strSQL = strSQL & " And L.Traspaso_tesoreria is not Null"
End If


If chkFiltros.Value = vbChecked Then
    If cboBanco.Text <> "TODOS" Then
      strSQL = strSQL & " And L.cod_banco = " & cboBanco.ItemData(cboBanco.ListIndex)
    End If

    If cboOficina.Text <> "TODOS" Then
      strSQL = strSQL & " And L.cod_oficina = '" & cboOficina.ItemData(cboOficina.ListIndex) & "'"
    End If

    If cboUsuarios.Text <> "TODOS" Then
      strSQL = strSQL & " And L.usuario like '" & cboUsuarios.Text & "%'"
    End If

    If cboSistema.Text <> "TODOS" Then
      strSQL = strSQL & " And isnull(L.cod_app,'" & App.ProductName & "') like '" & cboSistema.Text & "%'"
    End If

    If cboTokenConsulta.Text <> "TODOS" Then
      strSQL = strSQL & " And isnull(L.ID_Token,'') like '" & cboTokenConsulta.Text & "%'"
    End If
End If

lsw.ListItems.Clear
lswDet.ListItems.Clear

vCasos = False
strSQL = strSQL & " group by C.Cedula,S.nombre,L.Cta_Ahorros,B.descripcion" _
       & " having count(*) > 1"
Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
  Set itmX = lsw.ListItems.Add(, , rs!Cedula)
      itmX.SubItems(1) = rs!Nombre
      itmX.SubItems(2) = rs!Liquidaciones
      itmX.SubItems(3) = Format(rs!TOTAL_GIRAR, "Standard")
      itmX.SubItems(4) = rs!Cta_Ahorros & ""
      itmX.SubItems(5) = rs!Descripcion & ""
   
   For i = 1 To vGrid.MaxRows
      vGrid.Row = i
      vGrid.Col = 3
      If Trim(rs!Cedula) = Trim(vGrid.Text) Then
             vGrid.Col = 2
             vGrid.BackColor = vbYellow
             vCasos = True
      End If
   Next i
  rs.MoveNext

Loop
rs.Close


Me.MousePointer = vbDefault

If vCasos Then
    MsgBox "Existen Casos con más de un giro en esta misma Remesa (Casos color Amarillo) verifique!!!", vbExclamation
End If

'If lsw.ListItems.Count > 0 Then
'    tcMain.Item(1).Selected = True
'End If


End Sub


Private Sub sbFiltros()
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'OBJETIVO:      Llena combos con filtros adicionales segun los rangos de solicitud base    .
'REFERENCIAS:   Ninguna.
'OBSERVACIONES: Ninguna.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim strSQL As String

If dtpHasta.Value < dtpDesde.Value Then
   MsgBox "Verifique el Rango de Fechas", vbInformation, "Error"
   Exit Sub
End If

Me.MousePointer = vbHourglass


'Cargado de Bancos
strSQL = "Select L.cod_banco as Idx,isnull(B.descripcion,'Sin Banco') as Itmx" _
        & " From Fnd_Liquidacion L left join Tes_Bancos B on L.cod_Banco = B.id_Banco" _
        & " Where L.Fecha between '" & Format(dtpDesde.Value, "yyyy/mm/dd") & " 00:00:00' and '" _
        & Format(dtpHasta.Value, "yyyy/mm/dd") & " 23:59:59'"
If Mid(cboEstado.Text, 1, 1) = "P" Then
  strSQL = strSQL & " And L.Traspaso_tesoreria is Null"
Else
  strSQL = strSQL & " And L.Traspaso_tesoreria is not Null"
End If

strSQL = strSQL & " Group by L.cod_banco,B.descripcion"

Call sbCbo_Llena_New(cboBanco, strSQL, True, True)
        

'Cargado de Usuarios
strSQL = "Select L.USUARIO as 'idX', L.USUARIO as 'Itmx'" _
        & " From Fnd_Liquidacion L" _
        & " Where L.Fecha between '" & Format(dtpDesde.Value, "yyyy/mm/dd") & " 00:00:00' and '" _
        & Format(dtpHasta.Value, "yyyy/mm/dd") & " 23:59:59'"
        
If Mid(cboEstado.Text, 1, 1) = "P" Then
  strSQL = strSQL & " And L.Traspaso_tesoreria is Null"
Else
  strSQL = strSQL & " And L.Traspaso_tesoreria is not Null"
End If
strSQL = strSQL & " Group by L.usuario"
Call sbCbo_Llena_New(cboUsuarios, strSQL, True, True)



'Cargado de Sistemas
strSQL = "Select ISNULL(L.COD_APP,'" & App.ProductName & "') as 'Itmx', ISNULL(L.COD_APP,'" & App.ProductName & "') as 'idX'" _
        & " From Fnd_Liquidacion L" _
        & " Where L.Fecha between '" & Format(dtpDesde.Value, "yyyy/mm/dd") & " 00:00:00' and '" _
        & Format(dtpHasta.Value, "yyyy/mm/dd") & " 23:59:59'"
        
If Mid(cboEstado.Text, 1, 1) = "P" Then
  strSQL = strSQL & " And L.Traspaso_tesoreria is Null"
Else
  strSQL = strSQL & " And L.Traspaso_tesoreria is not Null"
End If
strSQL = strSQL & " Group by ISNULL(L.COD_APP,'" & App.ProductName & "') "
Call sbCbo_Llena_New(cboSistema, strSQL, True, True)


'Cargado de Token
strSQL = "Select ISNULL(L.ID_TOKEN,'') as 'Itmx', ISNULL(L.ID_TOKEN,'') as 'idX'" _
        & " From Fnd_Liquidacion L" _
        & " Where L.Fecha between '" & Format(dtpDesde.Value, "yyyy/mm/dd") & " 00:00:00' and '" _
        & Format(dtpHasta.Value, "yyyy/mm/dd") & " 23:59:59'"
        
If Mid(cboEstado.Text, 1, 1) = "P" Then
  strSQL = strSQL & " And L.Traspaso_tesoreria is Null"
Else
  strSQL = strSQL & " And L.Traspaso_tesoreria is not Null"
End If
strSQL = strSQL & " Group by ISNULL(L.ID_TOKEN,'')"
Call sbCbo_Llena_New(cboTokenConsulta, strSQL, True, True)


'Cargado de Oficinas
strSQL = "Select rtrim(L.cod_Oficina) as 'idX',  isnull(O.descripcion,'') as 'Itmx'" _
        & " From Fnd_Liquidacion L left join SIF_Oficinas O on L.cod_oficina = O.cod_oficina" _
        & " Where L.Fecha between '" & Format(dtpDesde.Value, "yyyy/mm/dd") & " 00:00:00' and '" _
        & Format(dtpHasta.Value, "yyyy/mm/dd") & " 23:59:59'"
If Mid(cboEstado.Text, 1, 1) = "P" Then
  strSQL = strSQL & " And L.Traspaso_tesoreria is Null"
Else
  strSQL = strSQL & " And L.Traspaso_tesoreria is not Null"
End If

strSQL = strSQL & " Group by L.cod_Oficina,O.descripcion"

Call sbCbo_Llena_New(cboOficina, strSQL, True, True)


Me.MousePointer = vbDefault


End Sub


Private Sub sbGenerar()
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'OBJETIVO:      Genera a Tesoreria las liquidaciones.
'REFERENCIAS:   sbTesoreria - (Genera el Asiento de la liquidacion en el
'               modulo de Tesoreria)
'               ProcedimientoErrores - (Registra error en caso de que ocurra uno dentro del
'               Procedimiento)
'OBSERVACIONES: Ninguna.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim strSQL As String, vRetencion As String
Dim i As Long, vFecha As String
Dim iDuplicado As Integer
Dim vToken As String
Dim rs As New ADODB.Recordset

On Error GoTo vError


vToken = cboToken.ItemData(cboToken.ListIndex)


If vGrid.MaxRows = 0 Then Exit Sub
If Mid(cboEstado.Text, 1, 1) = "G" Then Exit Sub



Me.MousePointer = vbHourglass

If Mid(cboAccion.Text, 1, 1) = "R" Then
  vRetencion = cboTipo.ItemData(cboTipo.ListIndex)
Else
  vRetencion = ""
End If

vFecha = Format(fxFechaServidor, "yyyy/mm/dd")
PrgBar.Max = vGrid.MaxRows

strSQL = ""

For i = 1 To vGrid.MaxRows
  vGrid.Row = i
  vGrid.Col = 14
  iDuplicado = IIf(vGrid.Text = "", 0, vGrid.Text)
  vGrid.Col = 1
 
  If vGrid.Value = vbChecked And iDuplicado = 0 Then
     
     If Mid(cboAccion.Text, 1, 1) = "D" Then
            'Desembolsar
'             Call sbTesoreria(i, vFecha, vToken)
             
            vGrid.Col = 2
            strSQL = strSQL & Space(10) & "exec spFNDRetLiqTesoreria " & vGrid.Text & ",'" & glogon.Usuario & "','" & vToken & "'"

     Else
            'Retener
            vGrid.Col = 2
            strSQL = strSQL & Space(10) & "Update Fnd_Liquidacion set Traspaso_Tesoreria= dbo.MyGetdate(),Traspaso_Usuario = '" _
                   & glogon.Usuario & "',Solicitud_Tesoreria= 0, RETENCION_CODIGO = '" & vRetencion & "',NOTAS = ''" _
                   & " Where Consec = " & vGrid.Text
     End If
     
     If Len(strSQL) > 20000 Then
        Call ConectionExecute(strSQL)
        strSQL = ""
     End If
     
  End If
 
 PrgBar.Value = i
 
Next i

'Procesa Lote Final
If Len(strSQL) > 0 Then
   Call ConectionExecute(strSQL)
   strSQL = ""
End If


PrgBar.Value = 0

Me.MousePointer = vbDefault
MsgBox "Proceso Finalizado...", vbExclamation

Call sbBuscar


Exit Sub
vError:
   MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub



Private Sub sbGenerar_Unificado()
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'OBJETIVO:      Genera a Tesoreria las liquidaciones.
'REFERENCIAS:   sbTesoreria - (Genera el Asiento de la liquidacion en el
'               modulo de Tesoreria)
'               ProcedimientoErrores - (Registra error en caso de que ocurra uno dentro del
'               Procedimiento)
'OBSERVACIONES: Ninguna.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim strSQL As String, vRetencion As String
Dim i As Long, vFecha As String
Dim iDuplicado As Integer
Dim vToken As String
Dim rs As New ADODB.Recordset

On Error GoTo vError


vToken = cboToken.ItemData(cboToken.ListIndex)

'strSQL = "select top 1 id_token from tes_tokens where estado = 'A' order by registro_fecha "
'Call OpenRecordSet(rs, strSQL)
'If Not rs.EOF Then
'  vToken = rs!id_token
'Else
'  vToken = fxTesToken
'End If
'rs.Close

If lsw.ListItems.Count = 0 Then Exit Sub
If Mid(cboEstado.Text, 1, 1) = "G" Then Exit Sub



Me.MousePointer = vbHourglass

If Mid(cboAccion.Text, 1, 1) = "R" Then
  vRetencion = cboTipo.ItemData(cboTipo.ListIndex)
Else
  vRetencion = ""
End If

vFecha = Format(fxFechaServidor, "yyyy/mm/dd")
PrgBar.Max = lsw.ListItems.Count + 1

strSQL = ""

With lsw.ListItems

For i = 1 To .Count
  
  If .Item(i).Checked = True Then
     
     If Mid(cboAccion.Text, 1, 1) = "D" Then
            'Desembolsar
            strSQL = strSQL & Space(10) & "exec spFndRetLiqTesoreria_Unificado '" & Trim(.Item(i).Text) _
                                        & "','" & glogon.Usuario & "','" & vToken _
                                        & "','" & Format(dtpDesde.Value, "yyyy/mm/dd") & " 00:00:00" _
                                        & "','" & Format(dtpHasta.Value, "yyyy/mm/dd") & " 23:59:59'"

     Else
            'Retener
            vGrid.Col = 2
'            strSQL = strSQL & Space(10) & "Update Fnd_Liquidacion set Traspaso_Tesoreria= dbo.MyGetdate(),Traspaso_Usuario = '" _
'                   & glogon.Usuario & "',Solicitud_Tesoreria= 0, RETENCION_CODIGO = '" & vRetencion & "',NOTAS = ''" _
'                   & " Where Consec = " & vGrid.Text
     End If
     
     If Len(strSQL) > 20000 Then
        Call ConectionExecute(strSQL)
        strSQL = ""
     End If
     
  End If
 
 PrgBar.Value = i
 
Next i

End With

'Procesa Lote Final
If Len(strSQL) > 0 Then
   Call ConectionExecute(strSQL)
   strSQL = ""
End If


PrgBar.Value = 0

Me.MousePointer = vbDefault
MsgBox "Proceso Finalizado...", vbExclamation

Call sbRevisaDuplicadosEnLaRemesa


Exit Sub
vError:
   MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub dtpDesde_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
   dtpHasta.SetFocus
End If
End Sub


Private Sub Form_Activate()
vModulo = 18 'Fondo de Inversion

End Sub


Private Sub sbTokens_Load()
Dim strSQL As String

On Error GoTo vError

strSQL = "exec spTes_Token_Consulta '', 'A','" & glogon.Usuario & "'"
Call sbCbo_Llena_New(cboToken, strSQL, False, True)

Exit Sub

vError:

End Sub


Private Sub Form_Load()
Dim strSQL As String

vModulo = 18 'Fondo de Inversion

dtpDesde.Value = fxFechaServidor
dtpHasta.Value = dtpDesde.Value

vGrid.AppearanceStyle = fxGridStyle


Set imgBanner.Picture = frmContenedor.imgBanner_01.Picture

vPaso = True
With lswDet.ColumnHeaders
    .Clear
    .Add , , "No. Liq", 1200
    .Add , , "Plan", 1200, vbCenter
    .Add , , "Contrato", 1000, vbCenter
    .Add , , "Identificación", 1400, vbCenter
    .Add , , "Nombre", 4200
    .Add , , "Monto a Girar", 1800, vbRightJustify
    .Add , , "Cuenta", 2100, vbCenter
    .Add , , "Banco", 3200
End With
lswDet.Checkboxes = False

With lsw.ColumnHeaders
    .Clear
    .Add , , "Identificación", 1600
    .Add , , "Nombre", 3600
    .Add , , "No. Liqs", 1200, vbCenter
    .Add , , "Monto", 1600, vbRightJustify
    .Add , , "Cuenta", 2100, vbCenter
    .Add , , "Bancos", 3000
End With


cboOficina.AddItem "TODOS"
cboOficina.Text = "TODOS"

cboBanco.AddItem "TODOS"
cboBanco.Text = "TODOS"

cboUsuarios.AddItem "TODOS"
cboUsuarios.Text = "TODOS"

cboSistema.AddItem "TODOS"
cboSistema.Text = "TODOS"



cboAccion.Clear
cboAccion.AddItem "Desembolsar"
cboAccion.AddItem "Retener"
cboAccion.Text = "Desembolsar"

cboEstado.Clear
cboEstado.AddItem "Pendiente"
cboEstado.AddItem "Generado"
cboEstado.Text = "Pendiente"

Call sbTokens_Load

vPaso = False

Call cboAccion_Click

Call Formularios(Me)
Call RefrescaTags(Me)


End Sub


Private Sub sbTesoreria(Row As Long, vFecha As String, vToken As String)
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError
vGrid.Row = Row

vGrid.Col = 2
strSQL = "exec spFNDRetLiqTesoreria " & vGrid.Text & ",'" & glogon.Usuario & "','" & vToken & "'"
Call ConectionExecute(strSQL)

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub Form_Resize()
On Error Resume Next

imgBanner.Width = Me.Width

'Line1.X2 = Me.Width
tcMain.Width = Me.Width - 400
tcMain.Height = Me.Height - (tcMain.top + 750)

vGrid.Width = tcMain.Width - 100
vGrid.Height = tcMain.Height - (vGrid.top + 100)

lsw.Width = tcMain.Width - 100
lsw.Height = tcMain.Height - (lsw.top + lswDet.Height + 200)

lswDet.Width = lsw.Width
lswDet.top = lsw.top + lsw.Height + 100

PrgBar.Width = Me.Width
PrgBar.top = tcMain.top + tcMain.Height + 20

End Sub


Private Sub lsw_ItemClick(ByVal Item As XtremeSuiteControls.ListViewItem)
 Call sbPersona_Detalle(Item.Text)
End Sub


Private Sub tcMain_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
    If Item.Index = 0 Then
        Call sbBuscar
    Else
        Call sbRevisaDuplicadosEnLaRemesa
    End If
End Sub

Private Sub tlbAcciones_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
  Case "Buscar"
    
    If tcMain.SelectedItem = 0 Then
        Call sbBuscar
    Else
        Call sbRevisaDuplicadosEnLaRemesa
    End If
    
  Case "Generar"
    If tcMain.SelectedItem = 0 Then
        Call sbGenerar
    Else
        Call sbGenerar_Unificado
    End If

End Select
End Sub

Private Sub sbInformes(pInforme As String)
Dim strSQL As String

Me.MousePointer = vbHourglass

With frmContenedor.Crt
 .Reset
 .WindowShowGroupTree = True
 .WindowShowPrintSetupBtn = True
 .WindowShowRefreshBtn = True
 .WindowShowSearchBtn = True
 .WindowState = crptMaximized
 .WindowTitle = "Reportes del Módulo de Fondos de Inversion"
 
 .Connect = glogon.ConectRPT
 
 .Formulas(0) = "Fecha='" & Format(fxFechaServidor, "dd/mm/yyyy") & "'"
 .Formulas(1) = "Empresa='" & GLOBALES.gstrNombreEmpresa & "'"

 
    Select Case pInforme
      Case "Resumen"
        .ReportFileName = SIFGlobal.fxPathReportes("Fondos_LiquidacionesTesoreriaResumen.rpt")
          strSQL = strSQL & "{FND_LIQUIDACION.TRASPASO_TESORERIA} in Date(" & Format(dtpDesde.Value, "yyyy,mm,dd") _
                 & ") to Date(" & Format(dtpHasta.Value, "yyyy,mm,dd") & ")"
      
      Case "Detalle"
        .ReportFileName = SIFGlobal.fxPathReportes("Fondos_LiquidacionesTesoreriaDetalle.rpt")
          strSQL = strSQL & "{FND_LIQUIDACION.TRASPASO_TESORERIA} in Date(" & Format(dtpDesde.Value, "yyyy,mm,dd") _
                 & ") to Date(" & Format(dtpHasta.Value, "yyyy,mm,dd") & ")"
      
      Case "Pendientes"
        .ReportFileName = SIFGlobal.fxPathReportes("Fondos_LiquidacionesTesoreriaPendientes.rpt")
          strSQL = strSQL & "{FND_LIQUIDACION.FECHA} in Date(" & Format(dtpDesde.Value, "yyyy,mm,dd") _
                 & ") to Date(" & Format(dtpHasta.Value, "yyyy,mm,dd") & ") AND ISNULL({FND_LIQUIDACION.TRASPASO_TESORERIA})"
    
    End Select
    

    If cboBanco.Text <> "TODOS" Then
      strSQL = strSQL & " And {FND_LIQUIDACION.COD_BANCO} = " & cboBanco.ItemData(cboBanco.ListIndex)
    End If

    If cboOficina.Text <> "TODOS" Then
      strSQL = strSQL & " And {FND_LIQUIDACION.COD_OFICINA} = '" & cboOficina.ItemData(cboOficina.ListIndex) & "'"
    End If

    If cboUsuarios.Text <> "TODOS" Then
      strSQL = strSQL & " And {FND_LIQUIDACION.USUARIO} = '" & cboUsuarios.Text & "'"
    End If

    If cboSistema.Text <> "TODOS" Then
      strSQL = strSQL & " And {FND_LIQUIDACION.COD_APP} = '" & cboSistema.Text & "'"
    End If
 
    If cboTokenConsulta.Text <> "TODOS" And cboTokenConsulta.Text <> "" Then
      strSQL = strSQL & " And {FND_LIQUIDACION.ID_TOKEN} = '" & cboTokenConsulta.Text & "'"
    End If
 
  .Formulas(2) = "SubTitulo='Del  " & Format(dtpDesde, "dd/mm/yyyy") & "  Al  " & Format(dtpHasta, "dd/mm/yyyy") & "'"
  .Formulas(3) = "Usuario='" & glogon.Usuario & "'"
  
  .ReplaceSelectionFormula strSQL
  .PrintReport
End With



Me.MousePointer = vbDefault


End Sub

Private Sub sbCargaGridLocal(vGridMaxCol As Integer, strSQL As String)
Dim rs As New ADODB.Recordset, i As Integer, strResultado As String
Dim strLista As String

Me.MousePointer = vbHourglass

vGrid.MaxCols = vGridMaxCol
vGrid.MaxRows = 1

vGrid.Row = vGrid.MaxRows
vDuplicado = False
strLista = ""

rs.CursorLocation = adUseServer
Call OpenRecordSet(rs, strSQL, 0)
 
PrgBar.Max = rs.RecordCount + 1
PrgBar.Value = 0

Do While Not rs.EOF
  vGrid.Row = vGrid.MaxRows
  PrgBar.Value = PrgBar.Value + 1
    
  For i = 1 To vGrid.MaxCols
    vGrid.Col = i
    Select Case i
     Case 1
        vGrid.Text = CStr(rs!Valor)
     Case 2
        vGrid.Text = CStr(rs!consec)
     Case 3
        vGrid.Text = CStr(rs!Cedula)
     Case 4
        vGrid.Text = CStr(rs!Nombre)
        If rs!PAGO_TERCERO_APL = 1 Then
              vGrid.TextTip = TextTipFixed
              vGrid.TextTipDelay = 1000
              vGrid.CellNote = "Beneficiario: " & rs!PAGO_TERCERO_ID & " - " & rs!PAGO_TERCERO_NOMBRE
        End If
     Case 5
        vGrid.Text = rs!Cod_Plan
     Case 6
        vGrid.Text = CStr(rs!COD_Contrato)
     Case 7
            vGrid.Text = Format(rs!TOTAL_GIRAR, "Standard")
     Case 8
        vGrid.Text = rs!Usuario
     Case 9
        vGrid.Value = rs!Oficina
     
     Case 10
        vGrid.Value = rs!Tipo
     
     Case 11
        vGrid.Value = rs!Cta_Ahorros
    Case 12
        vGrid.Value = rs!Descripcion & ""
    Case 13
        vGrid.Value = Format(rs!fecha, "dd/mm/yyyy hh:mm:ss")
    Case 14
         If rs!Duplicado = 1 And IsNull(rs!TES_SUPERVISION_FECHA) Then
             vDuplicado = True
             vGrid.Col = 2
             vGrid.Row = vGrid.ActiveRow
             vGrid.BackColor = vbRed
             strLista = strLista & "Cédula " & rs!Cedula & " Contrato " & rs!COD_Contrato & " Plan " & rs!Cod_Plan & " LIQ. " & rs!consec & vbCrLf
         Else
            vGrid.ForeColor = vbBlack
        End If
        vGrid.Col = 14
        vGrid.Value = IIf(vDuplicado, rs!Duplicado, 0)
    
    Case 15 'Token
        vGrid.Value = rs!id_token & ""
    
    End Select
  
  Next i
  
  vGrid.MaxRows = vGrid.MaxRows + 1
  
  rs.MoveNext
Loop

If vDuplicado Then
      MsgBox "Estas liquidaciones necesitan autorización para ser trasladadas ya que cuentan" _
          & "con una transacción por un monto igual en Tesorería " & vbCrLf & vbCrLf & strLista, vbCritical
End If
rs.Close

Me.MousePointer = vbDefault
End Sub

