VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.0#0"; "Codejock.Controls.v22.0.0.ocx"
Begin VB.Form frmFNDBitacora 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Bitacora especial de Fondos"
   ClientHeight    =   7575
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15840
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7575
   ScaleWidth      =   15840
   WindowState     =   2  'Maximized
   Begin XtremeSuiteControls.CheckBox chkFechas 
      Height          =   204
      Left            =   3000
      TabIndex        =   0
      Top             =   600
      Width           =   204
      _Version        =   1441792
      _ExtentX        =   353
      _ExtentY        =   353
      _StockProps     =   79
      BackColor       =   -2147483633
      Transparent     =   -1  'True
      UseVisualStyle  =   -1  'True
      Appearance      =   16
   End
   Begin XtremeSuiteControls.DateTimePicker dtpInicio 
      Height          =   312
      Left            =   1320
      TabIndex        =   1
      Top             =   600
      Width           =   1572
      _Version        =   1441792
      _ExtentX        =   2773
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
   Begin XtremeSuiteControls.DateTimePicker dtpCorte 
      Height          =   312
      Left            =   1320
      TabIndex        =   2
      Top             =   960
      Width           =   1572
      _Version        =   1441792
      _ExtentX        =   2773
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
   Begin XtremeSuiteControls.FlatEdit txtUsuario 
      Height          =   312
      Left            =   1320
      TabIndex        =   3
      Top             =   1320
      Width           =   1572
      _Version        =   1441792
      _ExtentX        =   2773
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
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtCedula 
      Height          =   312
      Left            =   1320
      TabIndex        =   4
      Top             =   1680
      Width           =   1572
      _Version        =   1441792
      _ExtentX        =   2773
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
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.ListView lswMovimientos 
      Height          =   4692
      Left            =   120
      TabIndex        =   5
      Top             =   2520
      Width           =   3732
      _Version        =   1441792
      _ExtentX        =   6583
      _ExtentY        =   8276
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
      MultiSelect     =   -1  'True
      HideSelection   =   0   'False
      View            =   3
      HideColumnHeaders=   -1  'True
      FullRowSelect   =   -1  'True
      HotTracking     =   -1  'True
      Appearance      =   16
   End
   Begin XtremeSuiteControls.CheckBox chkUsuarios 
      Height          =   204
      Left            =   3000
      TabIndex        =   6
      Top             =   1320
      Width           =   204
      _Version        =   1441792
      _ExtentX        =   353
      _ExtentY        =   353
      _StockProps     =   79
      BackColor       =   -2147483633
      Transparent     =   -1  'True
      UseVisualStyle  =   -1  'True
      Appearance      =   16
      Value           =   1
   End
   Begin XtremeSuiteControls.CheckBox chkMovimientos 
      Height          =   204
      Left            =   3600
      TabIndex        =   7
      Top             =   2160
      Width           =   204
      _Version        =   1441792
      _ExtentX        =   353
      _ExtentY        =   353
      _StockProps     =   79
      BackColor       =   -2147483633
      Transparent     =   -1  'True
      UseVisualStyle  =   -1  'True
      Appearance      =   16
      Value           =   1
   End
   Begin XtremeSuiteControls.GroupBox fraRevision 
      Height          =   492
      Left            =   6000
      TabIndex        =   14
      Top             =   1764
      Width           =   5652
      _Version        =   1441792
      _ExtentX        =   9970
      _ExtentY        =   868
      _StockProps     =   79
      BackColor       =   16777215
      Appearance      =   16
      BorderStyle     =   2
      Begin XtremeSuiteControls.CheckBox chkRevision 
         Height          =   252
         Left            =   3120
         TabIndex        =   15
         Top             =   120
         Width           =   2292
         _Version        =   1441792
         _ExtentX        =   4043
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Buscar Usuario/Fecha Revisión"
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
         Appearance      =   16
      End
      Begin XtremeSuiteControls.ComboBox cboRevision 
         Height          =   312
         Left            =   1440
         TabIndex        =   16
         Top             =   120
         Width           =   1572
         _Version        =   1441792
         _ExtentX        =   2778
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
         Caption         =   "Revisión ...:"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   192
         Index           =   3
         Left            =   360
         TabIndex        =   17
         Top             =   120
         Width           =   7692
      End
   End
   Begin XtremeSuiteControls.CheckBox chkRevTodos 
      Height          =   252
      Left            =   4920
      TabIndex        =   18
      Top             =   1848
      Width           =   1212
      _Version        =   1441792
      _ExtentX        =   2138
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Revisados?"
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
      Appearance      =   16
      Value           =   1
   End
   Begin XtremeSuiteControls.PushButton btnBuscar 
      Height          =   372
      Left            =   5400
      TabIndex        =   19
      Top             =   120
      Width           =   1212
      _Version        =   1441792
      _ExtentX        =   2138
      _ExtentY        =   656
      _StockProps     =   79
      Caption         =   "Buscar"
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
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Picture         =   "frmFNDBitacora.frx":0000
   End
   Begin XtremeSuiteControls.PushButton btnInforme 
      Height          =   372
      Left            =   6600
      TabIndex        =   20
      Top             =   120
      Width           =   1572
      _Version        =   1441792
      _ExtentX        =   2773
      _ExtentY        =   656
      _StockProps     =   79
      Caption         =   "Informe"
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
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Picture         =   "frmFNDBitacora.frx":0700
   End
   Begin XtremeSuiteControls.PushButton btnExportar 
      Height          =   372
      Left            =   8160
      TabIndex        =   21
      Top             =   120
      Width           =   1572
      _Version        =   1441792
      _ExtentX        =   2773
      _ExtentY        =   656
      _StockProps     =   79
      Caption         =   "Exportar"
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
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Picture         =   "frmFNDBitacora.frx":0E07
   End
   Begin XtremeSuiteControls.PushButton btnRevisar 
      Height          =   372
      Left            =   10320
      TabIndex        =   22
      Top             =   120
      Width           =   1572
      _Version        =   1441792
      _ExtentX        =   2773
      _ExtentY        =   656
      _StockProps     =   79
      Caption         =   "Revisar"
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
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Picture         =   "frmFNDBitacora.frx":16D8
   End
   Begin XtremeSuiteControls.ComboBox cboOperadora 
      Height          =   312
      Left            =   5400
      TabIndex        =   23
      Top             =   600
      Width           =   6492
      _Version        =   1441792
      _ExtentX        =   11456
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
   Begin XtremeSuiteControls.FlatEdit txtCodigo 
      Height          =   312
      Left            =   5400
      TabIndex        =   24
      Top             =   960
      Width           =   1332
      _Version        =   1441792
      _ExtentX        =   2350
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
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtDescripcion 
      Height          =   312
      Left            =   6720
      TabIndex        =   25
      Top             =   960
      Width           =   5172
      _Version        =   1441792
      _ExtentX        =   9123
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
      Locked          =   -1  'True
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit txtContrato 
      Height          =   312
      Left            =   5400
      TabIndex        =   29
      Top             =   1320
      Width           =   1332
      _Version        =   1441792
      _ExtentX        =   2350
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
   Begin MSComCtl2.FlatScrollBar FlatScrollBar 
      Height          =   252
      Left            =   12000
      TabIndex        =   30
      Top             =   960
      Width           =   492
      _ExtentX        =   873
      _ExtentY        =   450
      _Version        =   393216
      Arrows          =   65536
      Orientation     =   1638401
   End
   Begin XtremeSuiteControls.CheckBox chkTodosCont 
      Height          =   324
      Left            =   6840
      TabIndex        =   31
      Top             =   1320
      Width           =   1404
      _Version        =   1441792
      _ExtentX        =   2476
      _ExtentY        =   572
      _StockProps     =   79
      Caption         =   "Todos"
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
      Appearance      =   16
      Value           =   1
   End
   Begin FPSpreadADO.fpSpread vGrid 
      Height          =   4092
      Left            =   4080
      TabIndex        =   32
      Top             =   2280
      Width           =   12492
      _Version        =   524288
      _ExtentX        =   22035
      _ExtentY        =   7218
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
      MaxCols         =   12
      SpreadDesigner  =   "frmFNDBitacora.frx":1DFF
      VScrollSpecialType=   2
      AppearanceStyle =   1
   End
   Begin XtremeSuiteControls.CheckBox chkTodosPlan 
      Height          =   204
      Left            =   12600
      TabIndex        =   33
      Top             =   1000
      Width           =   1404
      _Version        =   1441792
      _ExtentX        =   2476
      _ExtentY        =   360
      _StockProps     =   79
      Caption         =   "Todos"
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
      Value           =   1
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Contrato"
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
      Height          =   312
      Index           =   1
      Left            =   4200
      TabIndex        =   28
      Top             =   1320
      Width           =   1332
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
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
      ForeColor       =   &H00000000&
      Height          =   312
      Index           =   0
      Left            =   4200
      TabIndex        =   27
      Top             =   960
      Width           =   1332
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
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
      ForeColor       =   &H00000000&
      Height          =   312
      Index           =   4
      Left            =   4200
      TabIndex        =   26
      Top             =   600
      Width           =   1332
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Usuario"
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
      Height          =   312
      Index           =   10
      Left            =   120
      TabIndex        =   13
      Top             =   1320
      Width           =   1092
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Inicio"
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
      Height          =   312
      Index           =   9
      Left            =   120
      TabIndex        =   12
      Top             =   600
      Width           =   1092
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Corte"
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
      Height          =   312
      Index           =   8
      Left            =   120
      TabIndex        =   11
      Top             =   960
      Width           =   1092
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Fechas y Usuario del Movimiento ...:"
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
      Height          =   312
      Index           =   7
      Left            =   120
      TabIndex        =   10
      Top             =   120
      Width           =   3252
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Movimientos ...:"
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
      Height          =   312
      Index           =   0
      Left            =   120
      TabIndex        =   9
      Top             =   2160
      Width           =   1452
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Identificación"
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
      Height          =   312
      Index           =   1
      Left            =   120
      TabIndex        =   8
      Top             =   1680
      Width           =   1092
   End
   Begin VB.Image imgBanner 
      Height          =   9396
      Left            =   0
      Picture         =   "frmFNDBitacora.frx":25F3
      Stretch         =   -1  'True
      Top             =   0
      Width           =   3972
   End
End
Attribute VB_Name = "frmFNDBitacora"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Dim vPaso As Boolean, vScroll As Boolean

Private Sub btnBuscar_Click()
    Call sbBuscar
End Sub

Private Sub btnExportar_Click()
Dim vHeaders As vGridHeaders
    vHeaders.Columnas = 12
    vHeaders.Headers(1) = "Revisado?"
    vHeaders.Headers(2) = "Usuario"
    vHeaders.Headers(3) = "Fecha"
    vHeaders.Headers(4) = "Movimiento"
    vHeaders.Headers(5) = "Operadora"
    vHeaders.Headers(6) = "Plan"
    vHeaders.Headers(7) = "Contrato"
    vHeaders.Headers(8) = "Detalle"
    vHeaders.Headers(9) = "Cédula"
    vHeaders.Headers(10) = "Nombre"
    vHeaders.Headers(11) = "Revisado por"
    vHeaders.Headers(12) = "Revisado Fecha"

Call sbSIFGridExportar(vGrid, vHeaders, "Fondos_BitacoraEspecial")

End Sub

Private Sub btnInforme_Click()
        vGrid.PrintHeader = "Fondos de Ahorros: Bitácora Especial, Fecha : " & fxFechaServidor & " Usuario : " & glogon.Usuario
        vGrid.PrintFooter = "Fechas Rastreo...I:" & Format(dtpInicio.Value, "dd/mm/yyyy") & " C.:" & Format(dtpCorte.Value, "dd/mm/yyyy")
        vGrid.PrintOrientation = PrintOrientationLandscape
        vGrid.PrintSheet
End Sub

Private Sub btnRevisar_Click()
 Call sbRevisar
End Sub

Private Sub cboRevision_Click()
If cboRevision.ListCount = 0 Or vPaso Then Exit Sub
Call sbBuscar
End Sub

Private Sub chkFechas_Click()
If chkFechas.Value = vbChecked Then
  dtpInicio.Enabled = False
  dtpCorte.Enabled = False
Else
  dtpInicio.Enabled = True
  dtpCorte.Enabled = True
End If
End Sub

Private Sub chkMovimientos_Click()
Dim i As Integer

For i = 1 To lswMovimientos.ListItems.Count
  lswMovimientos.ListItems.Item(i).Checked = chkMovimientos.Value
Next i
End Sub

Private Sub chkRevision_Click()
If chkRevision.Value = vbChecked Then
   txtUsuario.BackColor = cboRevision.BackColor
Else
   txtUsuario.BackColor = vbWhite
End If
End Sub

Private Sub chkRevTodos_Click()
Dim i As Long

For i = 1 To vGrid.MaxRows
  vGrid.Row = i
  vGrid.col = 1
  vGrid.Value = chkRevTodos.Value
Next i
End Sub

Private Sub chkTodosCont_Click()
If chkTodosCont.Value = vbChecked Then
   txtContrato.Enabled = False
 Else
   txtContrato.Enabled = True
   txtContrato = Empty
 End If
 
End Sub


Private Sub chkTodosPlan_Click()
If chkTodosPlan.Value = vbChecked Then
   txtCodigo.Enabled = False
   txtCodigo = "(Presione F4)"
   txtDescripcion = Empty
 Else
   txtCodigo.Enabled = True
   txtCodigo = "(Presione F4)"
   txtDescripcion = Empty
 End If
End Sub



Private Sub sbBuscar()
Dim rs As New ADODB.Recordset
On Error GoTo vError

Me.MousePointer = vbHourglass
   
vGrid.MaxRows = 0
vGrid.MaxCols = 12

vPaso = True

Call OpenRecordSet(rs, fxSQL)

Do While Not rs.EOF
  vGrid.MaxRows = vGrid.MaxRows + 1
  vGrid.Row = vGrid.MaxRows
  
  vGrid.col = 1
  vGrid.Text = rs!Revisado
  vGrid.CellTag = rs!id_Bitacora
  
  vGrid.col = 2
  vGrid.Text = rs!Usuario
  vGrid.col = 3
  vGrid.Text = rs!fecha ' Format(rs!Fecha, "dd/mm/yyyy")
  vGrid.col = 4
  vGrid.Text = rs!MovimientoDesc
  vGrid.col = 5
  vGrid.Text = rs!cod_Operadora
  vGrid.col = 6
  vGrid.Text = rs!cod_Plan
  vGrid.col = 7
  vGrid.Text = rs!COD_CONTRATO
  vGrid.col = 8
  vGrid.Text = rs!Detalle
  
  vGrid.col = 9
  vGrid.Text = Trim(rs!Cedula)
  vGrid.col = 10
  vGrid.Text = rs!Nombre
  
  vGrid.col = 11
  vGrid.Text = rs!Revisado_Usuario & ""
  vGrid.col = 12
  vGrid.Text = rs!Revisado_Fecha & ""
  
  rs.MoveNext
  
Loop
rs.Close


vPaso = False
Me.MousePointer = vbDefault

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub chkUsuarios_Click()
 If chkUsuarios.Value = vbChecked Then
   txtUsuario.Enabled = False
 Else
   txtUsuario.Enabled = True
   txtUsuario = "(Presione F4)"
 End If
End Sub

Private Sub FlatScrollBar_Change()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

If vScroll Then
    strSQL = "select Top 1 cod_plan,descripcion, dateadd(year,1, getdate()) as 'Vence' from fnd_planes" _
           & " where cod_operadora = " & cboOperadora.ItemData(cboOperadora.ListIndex)
    
    If FlatScrollBar.Value = 1 Then
       strSQL = strSQL & " and cod_plan > '" & txtCodigo & "' order by cod_plan asc"
    Else
       strSQL = strSQL & " and cod_plan < '" & txtCodigo & "' order by cod_plan desc"
    End If
    
    Call OpenRecordSet(rs, strSQL)
    If Not rs.EOF And Not rs.BOF Then
      txtCodigo.Text = rs!cod_Plan
      txtDescripcion.Text = rs!Descripcion
    End If
    rs.Close
End If

vScroll = False
FlatScrollBar.Value = 0
vScroll = True

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub Form_Load()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem

vModulo = 18

vGrid.AppearanceStyle = fxGridStyle

vScroll = False
     FlatScrollBar.Value = 0
vScroll = True

lswMovimientos.ColumnHeaders.Add , , "", 3150

vPaso = True
    cboRevision.Clear
    cboRevision.AddItem "TODOS"
    cboRevision.AddItem "Pendientes"
    cboRevision.AddItem "Revisados"
    cboRevision.Text = "TODOS"
    
    dtpCorte.Value = fxFechaServidor
    dtpInicio.Value = DateAdd("d", -7, dtpCorte.Value)
    
    strSQL = "select rtrim(descripcion) as 'ItmX',cod_operadora as 'IdX' from FND_Operadoras"
    Call sbCbo_Llena_New(cboOperadora, strSQL, False, True)
    
    
    lswMovimientos.ListItems.Clear
    strSQL = "select MOVIMIENTO,DESCRIPCION from US_MOVIMIENTOS_BE WHERE MODULO = " & vModulo & " ORDER BY MOVIMIENTO"
    Call OpenRecordSet(rs, strSQL)
    Do While Not rs.EOF
     Set itmX = lswMovimientos.ListItems.Add(, , rs!Descripcion)
         itmX.Tag = rs!Movimiento
         itmX.Checked = chkMovimientos.Value
     rs.MoveNext
    Loop
    rs.Close
    
    
    chkTodosCont.Value = vbChecked
    txtContrato.Enabled = False
    
    chkTodosPlan.Value = vbChecked
    txtCodigo.Enabled = False

vPaso = False


Call Formularios(Me)
Call RefrescaTags(Me)

End Sub





Private Function fxSQL() As String
Dim strSQL As String
Dim vCadena As String, i As Integer


strSQL = "select C.*,S.cedula,S.nombre,M.Descripcion as MovimientoDesc,case when C.revisado_fecha is null then 0 else 1 end as 'Revisado'" _
       & " from fnd_contratos_cambios C inner join fnd_contratos X on C.cod_operadora = X.cod_operadora" _
       & " and C.cod_plan = X.cod_plan and C.cod_contrato = X.cod_contrato" _
       & " inner join Socios S on X.cedula = S.cedula" _
       & " inner join US_MOVIMIENTOS_BE M on C.Movimiento = M.Movimiento" _
       & " Where M.Modulo = " & vModulo

If Len(Trim(txtCedula.Text)) > 0 Then
  strSQL = strSQL & " and S.cedula like '%" & txtCedula.Text & "%'"
End If
       
If chkFechas.Value = vbUnchecked Then
   If chkRevision.Value = vbChecked Then
        strSQL = strSQL & " and C.Revisado_fecha between '" & Format(dtpInicio.Value, "yyyy/mm/dd") _
               & " 00:00:00' and '" & Format(dtpCorte.Value, "yyyy/mm/dd") & " 23:59:00'"
   Else
        strSQL = strSQL & " and C.fecha between '" & Format(dtpInicio.Value, "yyyy/mm/dd") _
               & " 00:00:00' and '" & Format(dtpCorte.Value, "yyyy/mm/dd") & " 23:59:00'"
   End If
End If


'Lista de Tipos de Movimientos
vCadena = " and C.movimiento in('"
For i = 1 To lswMovimientos.ListItems.Count
  If lswMovimientos.ListItems.Item(i).Checked Then
    vCadena = vCadena & "','" & lswMovimientos.ListItems.Item(i).Tag
  End If
Next i
strSQL = strSQL & vCadena & "')"

If chkUsuarios.Value = vbUnchecked Then
  If txtUsuario <> "" And txtUsuario <> "(Presione F4)" Then
     If chkRevision.Value = vbChecked Then
             strSQL = strSQL & " and C.Revisado_Usuario = '" & txtUsuario & "'"
     Else
             strSQL = strSQL & " and C.Usuario = '" & txtUsuario & "'"
     End If
  End If
End If


'En caso de que sea un plan
If chkTodosPlan.Value = vbUnchecked Then
  If txtCodigo <> "" And txtCodigo <> "(Presione F4)" Then
     strSQL = strSQL & " and C.cod_plan = '" & txtCodigo & "'"
  End If
End If

'En caso de que sea contrato
If chkTodosCont.Value = vbUnchecked Then
  If txtContrato <> "" Then
     strSQL = strSQL & " and C.cod_contrato = '" & txtContrato & "'"
  End If
End If

strSQL = strSQL & " and C.Cod_operadora = " & cboOperadora.ItemData(cboOperadora.ListIndex)

Select Case Mid(cboRevision.Text, 1, 1)
   Case "P" 'Pendientes
        strSQL = strSQL & " and C.Revisado_Fecha is null"
   Case "R" 'Revisados
        strSQL = strSQL & " and C.Revisado_Fecha is not null"
   Case "T" 'Todos
End Select

If chkRevision.Value = vbChecked Then
    strSQL = strSQL & " order by C.Revisado_fecha"
Else
    strSQL = strSQL & " order by C.fecha"
End If

fxSQL = strSQL

End Function



Private Sub Form_Resize()
On Error Resume Next

imgBanner.Height = Me.Height

vGrid.Width = Me.Width - (520 + vGrid.Left)
vGrid.Height = Me.Height - (vGrid.top + 750)

lswMovimientos.Height = Me.Height - (lswMovimientos.top + 750)

End Sub

Private Sub tlb_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
  Case "Buscar"
    Call sbBuscar
  
  Case "Revisar"
    Call sbRevisar
  
  Case "Reporte"
        vGrid.PrintHeader = "Fondos: Bitácora Especial, Fecha : " & fxFechaServidor & " Usuario : " & glogon.Usuario
        vGrid.PrintFooter = "Fechas Rastreo...I:" & Format(dtpInicio.Value, "dd/mm/yyyy") & " C.:" & Format(dtpCorte.Value, "dd/mm/yyyy")
        vGrid.PrintOrientation = PrintOrientationLandscape
        vGrid.PrintSheet
End Select
End Sub

Private Sub tlb_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
Dim vHeaders As vGridHeaders
    vHeaders.Columnas = 12
    vHeaders.Headers(1) = "Revisado?"
    vHeaders.Headers(2) = "Usuario"
    vHeaders.Headers(3) = "Fecha"
    vHeaders.Headers(4) = "Movimiento"
    vHeaders.Headers(5) = "Operadora"
    vHeaders.Headers(6) = "Plan"
    vHeaders.Headers(7) = "Contrato"
    vHeaders.Headers(8) = "Detalle"
    vHeaders.Headers(9) = "Cédula"
    vHeaders.Headers(10) = "Nombre"
    vHeaders.Headers(11) = "Revisado por"
    vHeaders.Headers(12) = "Revisado Fecha"
    
Select Case ButtonMenu.Key
  Case "Excel"
      Call sbSIFGridExportar(vGrid, vHeaders, "Fondos_BitacoraEspecial")
  Case "HTML"
      Call sbSIFGridExportar(vGrid, vHeaders, "Fondos_BitacoraEspecial", "HTML")
End Select

End Sub



Private Sub txtCodigo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtDescripcion.SetFocus

If KeyCode = vbKeyF4 Then
   gBusquedas.Columna = "cod_plan"
   gBusquedas.Orden = "cod_plan"
   gBusquedas.Filtro = "And Cod_operadora=" & cboOperadora.ItemData(cboOperadora.ListIndex)
   gBusquedas.Consulta = "select cod_plan,descripcion from fnd_planes"
   frmBusquedas.Show vbModal
   txtDescripcion.SetFocus
   
   If Trim(gBusquedas.Resultado) <> "" Then
      txtCodigo = Trim(gBusquedas.Resultado)
      txtDescripcion = Trim(gBusquedas.Resultado2)
   End If
   gBusquedas.Resultado = ""
   gBusquedas.Resultado2 = ""
End If

End Sub



Private Sub txtDescripcion_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then
   gBusquedas.Columna = "descripcion"
   gBusquedas.Orden = "descripcion"
   gBusquedas.Filtro = "And Cod_operadora=" & cboOperadora.ItemData(cboOperadora.ListIndex)
   gBusquedas.Consulta = "select cod_plan,descripcion from fnd_planes"
   frmBusquedas.Show vbModal
   txtDescripcion.SetFocus
   If Trim(gBusquedas.Resultado) <> "" Then
      txtCodigo = Trim(gBusquedas.Resultado)
      txtDescripcion = Trim(gBusquedas.Resultado2)
   End If
   gBusquedas.Resultado = ""
   gBusquedas.Resultado2 = ""
End If
End Sub

Private Sub txtUsuario_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then
    gBusquedas.Convertir = "N"
    gBusquedas.Resultado = Trim(txtUsuario)
    gBusquedas.Consulta = "Select Nombre,Descripcion From Usuarios"
    gBusquedas.Columna = "Nombre"
    gBusquedas.Orden = "Nombre"
    frmBusquedas.Show vbModal
    txtUsuario = Trim(gBusquedas.Resultado)
End If
End Sub



Private Sub sbRevisar()
Dim strSQL As String, i As Long, IdBitacora As Long
Dim vPlan As String, vContrato As String


If vPaso Or Not fraRevision.Enabled Then Exit Sub

On Error GoTo vError

Me.MousePointer = vbHourglass


With vGrid

  For i = 1 To .MaxRows
     .Row = i
     .col = 1
     If .Value = vbChecked Then
        IdBitacora = .CellTag
        .col = 11
        If Trim(.Text) = "" Then
 
           strSQL = "update fnd_contratos_cambios set revisado_usuario = '" & glogon.Usuario & "', revisado_fecha = dbo.MyGetdate()" _
                  & " where id_Bitacora = " & IdBitacora
           Call ConectionExecute(strSQL)
        
           vGrid.col = 11
           vGrid.Text = glogon.Usuario
           vGrid.col = 12
           vGrid.Text = Date
       
            .col = 4
            If .Text = "Registro de Contrato" Then
                 .col = 6
                 vPlan = .Text
                 vGrid.col = 7
                 vContrato = .Text
                 
                 Call sbSIFRegistraTags(vPlan, "S03", "Recepción desde Bitácora", vContrato, "FND", vPlan, vContrato)
            
            End If 'Col 4
         End If 'Col 11 .text = ""
      End If 'Col 1 : Value

   Next i

End With
 
Me.MousePointer = vbDefault
MsgBox "Revisión aplicada satisfactoriamente!", vbInformation
 
Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
 

End Sub


