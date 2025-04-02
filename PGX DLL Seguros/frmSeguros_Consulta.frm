VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.0#0"; "Codejock.Controls.v22.0.0.ocx"
Begin VB.Form frmSeguros_Consulta 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Módulo de Seguros: Consulta General"
   ClientHeight    =   10005
   ClientLeft      =   225
   ClientTop       =   555
   ClientWidth     =   16455
   Icon            =   "frmSeguros_Consulta.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10005
   ScaleWidth      =   16455
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin XtremeSuiteControls.CheckBox chkTiposSeguros 
      Height          =   210
      Left            =   2880
      TabIndex        =   25
      Top             =   840
      Width           =   210
      _Version        =   1441792
      _ExtentX        =   370
      _ExtentY        =   370
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      Value           =   1
   End
   Begin MSComctlLib.StatusBar StatusBarX 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   9750
      Width           =   16455
      _ExtentX        =   29025
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4304
            MinWidth        =   4304
            Object.ToolTipText     =   "Casos Encontrados..:"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   6068
            MinWidth        =   6068
            Object.ToolTipText     =   "Total Registrado..:"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin FPSpreadADO.fpSpread vGrid 
      Height          =   7095
      Left            =   3360
      TabIndex        =   6
      Top             =   1440
      Width           =   10695
      _Version        =   524288
      _ExtentX        =   18865
      _ExtentY        =   12515
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
      MaxCols         =   23
      SpreadDesigner  =   "frmSeguros_Consulta.frx":6852
      VScrollSpecial  =   -1  'True
      VScrollSpecialType=   2
      AppearanceStyle =   1
   End
   Begin VB.Timer TimerX 
      Interval        =   10
      Left            =   12240
      Top             =   120
   End
   Begin XtremeSuiteControls.DateTimePicker dtpInicio 
      Height          =   330
      Left            =   1560
      TabIndex        =   12
      Top             =   8520
      Width           =   1575
      _Version        =   1441792
      _ExtentX        =   2778
      _ExtentY        =   582
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
      Height          =   330
      Left            =   1560
      TabIndex        =   13
      Top             =   8880
      Width           =   1575
      _Version        =   1441792
      _ExtentX        =   2778
      _ExtentY        =   582
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
   Begin XtremeSuiteControls.FlatEdit txtPoliza 
      Height          =   315
      Left            =   3360
      TabIndex        =   14
      Top             =   360
      Width           =   1815
      _Version        =   1441792
      _ExtentX        =   3201
      _ExtentY        =   556
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
   Begin XtremeSuiteControls.FlatEdit txtCedula 
      Height          =   315
      Left            =   5160
      TabIndex        =   15
      Top             =   360
      Width           =   1815
      _Version        =   1441792
      _ExtentX        =   3201
      _ExtentY        =   556
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
   Begin XtremeSuiteControls.FlatEdit txtNombre 
      Height          =   315
      Left            =   6960
      TabIndex        =   16
      Top             =   360
      Width           =   4935
      _Version        =   1441792
      _ExtentX        =   8705
      _ExtentY        =   556
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
   Begin XtremeSuiteControls.FlatEdit txtUsuario 
      Height          =   315
      Left            =   3360
      TabIndex        =   17
      Top             =   960
      Width           =   1815
      _Version        =   1441792
      _ExtentX        =   3201
      _ExtentY        =   556
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
   Begin XtremeSuiteControls.FlatEdit txtClienteCorId 
      Height          =   315
      Left            =   5160
      TabIndex        =   18
      Top             =   960
      Width           =   1815
      _Version        =   1441792
      _ExtentX        =   3201
      _ExtentY        =   556
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
   Begin XtremeSuiteControls.FlatEdit txtClienteCorNombre 
      Height          =   315
      Left            =   6960
      TabIndex        =   19
      Top             =   960
      Width           =   4935
      _Version        =   1441792
      _ExtentX        =   8705
      _ExtentY        =   556
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
   Begin XtremeSuiteControls.ComboBox cboAseguradora 
      Height          =   330
      Left            =   120
      TabIndex        =   20
      Top             =   360
      Width           =   3015
      _Version        =   1441792
      _ExtentX        =   5318
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
   Begin XtremeSuiteControls.ComboBox cboComercializadora 
      Height          =   330
      Left            =   120
      TabIndex        =   21
      Top             =   4440
      Width           =   3015
      _Version        =   1441792
      _ExtentX        =   5318
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
      Height          =   330
      Left            =   1560
      TabIndex        =   22
      Top             =   7800
      Width           =   1575
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
   Begin XtremeSuiteControls.CheckBox chkVendedores 
      Height          =   210
      Left            =   2880
      TabIndex        =   26
      Top             =   4920
      Width           =   210
      _Version        =   1441792
      _ExtentX        =   370
      _ExtentY        =   370
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      Value           =   1
   End
   Begin XtremeSuiteControls.PushButton btnBuscar 
      Height          =   615
      Left            =   12240
      TabIndex        =   27
      Top             =   720
      Width           =   1215
      _Version        =   1441792
      _ExtentX        =   2138
      _ExtentY        =   1080
      _StockProps     =   79
      Caption         =   "Buscar"
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
      Picture         =   "frmSeguros_Consulta.frx":75BB
   End
   Begin XtremeSuiteControls.PushButton btnExportar 
      Height          =   615
      Left            =   13440
      TabIndex        =   28
      Top             =   720
      Width           =   1575
      _Version        =   1441792
      _ExtentX        =   2773
      _ExtentY        =   1080
      _StockProps     =   79
      Caption         =   "Exportar"
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
      Picture         =   "frmSeguros_Consulta.frx":7FD9
   End
   Begin XtremeSuiteControls.CheckBox chkFecha 
      Height          =   210
      Left            =   2880
      TabIndex        =   30
      Top             =   8280
      Width           =   210
      _Version        =   1441792
      _ExtentX        =   370
      _ExtentY        =   370
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      Value           =   1
   End
   Begin XtremeSuiteControls.ListView lswTipoSeguros 
      Height          =   2895
      Left            =   120
      TabIndex        =   31
      Top             =   1200
      Width           =   3015
      _Version        =   1441792
      _ExtentX        =   5318
      _ExtentY        =   5106
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
   Begin XtremeSuiteControls.ListView lswVendedores 
      Height          =   2415
      Left            =   120
      TabIndex        =   32
      Top             =   5280
      Width           =   3015
      _Version        =   1441792
      _ExtentX        =   5318
      _ExtentY        =   4260
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
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Cliente Corporativo Nombre"
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
      Height          =   315
      Index           =   6
      Left            =   6960
      TabIndex        =   29
      Top             =   720
      Width           =   2535
   End
   Begin VB.Label lblComercializadora 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Comercializadora...:"
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
      Left            =   120
      TabIndex        =   24
      Top             =   4200
      Width           =   1815
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Aseguradora...:"
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
      Left            =   120
      TabIndex        =   23
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Cliente Corporativo Id"
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
      Height          =   315
      Index           =   3
      Left            =   5160
      TabIndex        =   11
      Top             =   720
      Width           =   1815
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
      ForeColor       =   &H00000000&
      Height          =   315
      Index           =   8
      Left            =   3360
      TabIndex        =   10
      Top             =   720
      Width           =   1815
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre"
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
      Height          =   315
      Index           =   7
      Left            =   6960
      TabIndex        =   9
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Cédula"
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
      Height          =   315
      Index           =   1
      Left            =   5160
      TabIndex        =   8
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "No. Póliza"
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
      Height          =   315
      Index           =   0
      Left            =   3360
      TabIndex        =   7
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label lblInicio 
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
      Height          =   315
      Left            =   120
      TabIndex        =   5
      Top             =   8520
      Width           =   1215
   End
   Begin VB.Label lblCorte 
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
      Height          =   315
      Left            =   120
      TabIndex        =   4
      Top             =   8880
      Width           =   1215
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Tipos de Seguros ...:"
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
      Index           =   2
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   1815
   End
   Begin VB.Label lblVendedores 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Vendedores"
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
      Left            =   120
      TabIndex        =   2
      Top             =   4920
      Width           =   1815
   End
   Begin VB.Label lblEstado 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Estado"
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
      Left            =   120
      TabIndex        =   1
      Top             =   7800
      Width           =   1215
   End
   Begin VB.Image imgBanner 
      Height          =   9630
      Left            =   0
      Picture         =   "frmSeguros_Consulta.frx":87DE
      Stretch         =   -1  'True
      Top             =   0
      Width           =   3285
   End
End
Attribute VB_Name = "frmSeguros_Consulta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem
Dim vPaso As Boolean



Private Sub btnBuscar_Click()
Call sbBuscar
End Sub

Private Sub btnExportar_Click()
Call sbExportar
End Sub

Private Sub cboAseguradora_Click()

If vPaso Then Exit Sub

On Error GoTo vError

lswTipoSeguros.ListItems.Clear
strSQL = "select COD_ASEGURADORA as IdX,  rtrim(COD_PRODUCTO) + ' - ' + rtrim(Descripcion) as ItmX" _
       & " from SEGUROS_TIPOS_PRODUCTOS" _
       & " where Activo = 1 "
       
If cboAseguradora.Text <> "TODOS" Then
    strSQL = strSQL & " and cod_aseguradora = '" & cboAseguradora.ItemData(cboAseguradora.ListIndex) & "'"
End If
       
strSQL = strSQL & " order by COD_ASEGURADORA, COD_PRODUCTO"
       
       
Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
 Set itmX = lswTipoSeguros.ListItems.Add(, , rs!itmX)
     itmX.Tag = rs!IdX
     itmX.Checked = chkTiposSeguros.Value
 rs.MoveNext
Loop
rs.Close

Exit Sub

vError:
   MsgBox fxSys_Error_Handler(Err.Description), vbCritical
   
End Sub

Private Sub cboComercializadora_Click()

If vPaso Then Exit Sub

On Error GoTo vError

lswVendedores.ListItems.Clear
strSQL = "select  cod_vendedor,Nombre from SEGUROS_Vendedores"

If cboComercializadora.Text <> "TODOS" Then
    strSQL = strSQL & " where cod_comercializadora = '" & cboComercializadora.ItemData(cboComercializadora.ListIndex) & "'"
End If
strSQL = strSQL & " Order by Nombre"


Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
 Set itmX = lswVendedores.ListItems.Add(, , rs!Nombre)
     itmX.Tag = rs!cod_vendedor
     itmX.Checked = chkVendedores.Value
 rs.MoveNext
Loop
rs.Close

Exit Sub

vError:
   MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub



Private Sub chkFecha_Click()

If chkFecha.Value = xtpChecked Then
    dtpInicio.Enabled = True
Else
    dtpInicio.Enabled = False
End If

dtpCorte.Enabled = dtpInicio.Enabled

End Sub

Private Sub chkTiposSeguros_Click()
Dim i As Integer

For i = 1 To lswTipoSeguros.ListItems.Count
  lswTipoSeguros.ListItems.Item(i).Checked = chkTiposSeguros.Value
Next i

End Sub

Private Sub chkVendedores_Click()
Dim i As Integer

For i = 1 To lswVendedores.ListItems.Count
  lswVendedores.ListItems.Item(i).Checked = chkVendedores.Value
Next i

End Sub

Private Sub Form_Activate()
 vModulo = 17
End Sub

Private Sub Form_Load()


vModulo = 17


Call Formularios(Me)
Call RefrescaTags(Me)

vGrid.AppearanceStyle = fxGridStyle



lswTipoSeguros.ColumnHeaders.Add , , "", 3150
lswVendedores.ColumnHeaders.Add , , "", 3150


vGrid.MaxRows = 0

cboEstado.Clear
cboEstado.AddItem "Pendientes"
cboEstado.AddItem "Activadas"
cboEstado.AddItem "Cerradas"
cboEstado.AddItem "[TODOS]"
cboEstado.Text = "[TODOS]"

dtpCorte.Value = fxFechaServidor
dtpInicio.Value = CDate(Year(dtpCorte.Value) & "/" & Format(Month(dtpCorte.Value), "00") & "/01")

vPaso = True

strSQL = "select rtrim(COD_COMERCIALIZADORA) as 'IdX',  rtrim(NOMBRE) as ItmX from SEGUROS_COMERCIALIZADORAS where Activo = 1 order by NOMBRE"
Call sbCbo_Llena_New(cboComercializadora, strSQL, True, True)

strSQL = "select rtrim(COD_ASEGURADORA) as 'IdX', rtrim(NOMBRE) as ItmX from SEGUROS_ASEGURADORAS where Activo = 1 order by NOMBRE"
Call sbCbo_Llena_New(cboAseguradora, strSQL, True, True)

vPaso = False

Call chkFecha_Click


End Sub

Private Sub Form_Resize()
On Error Resume Next

vGrid.Width = Me.Width - 3765
vGrid.Height = Me.Height - (vGrid.Top + StatusBarX.Height + 250)

lswTipoSeguros.Height = (Me.Height / 3.1188589540412) - 500   '- 6685
lswVendedores.Height = lswTipoSeguros.Height

lblComercializadora.Top = lswTipoSeguros.Top + lswTipoSeguros.Height + 205


cboComercializadora.Top = lblComercializadora.Top + lblComercializadora.Height

lblVendedores.Top = cboComercializadora.Top + cboComercializadora.Height + 200
chkVendedores.Top = lblVendedores.Top

lswVendedores.Top = lblVendedores.Top + 360

cboEstado.Top = lswVendedores.Top + lswVendedores.Height + 150
lblEstado.Top = cboEstado.Top

chkFecha.Top = lblEstado.Top + lblEstado.Height + 200

lblInicio.Top = chkFecha.Top + chkFecha.Height + 100
dtpInicio.Top = lblInicio.Top

lblCorte.Top = lblInicio.Top + 360
dtpCorte.Top = lblCorte.Top

imgBanner.Height = Me.Height

End Sub



Private Sub sbBuscar()
Dim i As Integer
Dim vCadena As String, iCantidad As Integer

On Error GoTo vError

Me.MousePointer = vbHourglass
iCantidad = 0


strSQL = "select '',COD_ASEGURADORA, NUM_POLIZA,rtrim(CEDULA),rtrim(NOMBRE),Estado_Desc" _
       & " ,CUOTA,MONTO,REGISTRO_FECHA,ACTIVA_FECHA,CIERRA_FECHA , PAGADO_TOTAL,COBRADO_TOTAL , Balanza_Cobro" _
       & " ,comision_Vendedor_Total,comision_Comercializa_Total,Comision_Interna_Total,isnull(Operacion,0), COD_PRODUCTO_Desc, Vendedor_NOMBRE" _
       & " ,Comercializadora_Nombre,Cliente_Cor_Nombre,Vendedor_Comision_Real" _
       & "  from  vSeguros_ListadoGeneral " _
       & " where Num_Poliza like '%" & txtPoliza.Text & "%'"

If Len(Trim(txtCedula.Text)) > 0 Then
   strSQL = strSQL & " and Cedula like '%" & txtCedula.Text & "%'"
End If

If Len(Trim(txtNombre.Text)) > 0 Then
   strSQL = strSQL & " and Nombre like '%" & txtNombre.Text & "%'"
End If

If Len(Trim(txtUsuario.Text)) > 0 Then
   strSQL = strSQL & " and Registro_Usuario like '%" & txtUsuario.Text & "%'"
End If

If Len(Trim(txtClienteCorId.Text)) > 0 Then
   strSQL = strSQL & " and Cod_Cliente_Corporativo like '%" & txtClienteCorId.Text & "%'"
End If

If Len(Trim(txtClienteCorNombre.Text)) > 0 Then
   strSQL = strSQL & " and Cliente_Cor_Nombre like '%" & txtClienteCorNombre.Text & "%'"
End If


'Tipos de Seguros
iCantidad = 0
For i = 1 To lswTipoSeguros.ListItems.Count
  If lswTipoSeguros.ListItems.Item(i).Checked Then
    iCantidad = iCantidad + 1
  End If
Next i

If iCantidad <> lswTipoSeguros.ListItems.Count Then
    iCantidad = 0
    vCadena = " and COD_PRODUCTO in('"
    For i = 1 To lswTipoSeguros.ListItems.Count
      If lswTipoSeguros.ListItems.Item(i).Checked Then
        vCadena = vCadena & "','" & SIFGlobal.fxCodText(lswTipoSeguros.ListItems.Item(i).Text)
        iCantidad = iCantidad + 1
      End If
    Next i
    strSQL = strSQL & vCadena & "')"
End If


'Lista de Vendedores
If chkVendedores.Value = vbUnchecked Then
    iCantidad = 0
    For i = 1 To lswVendedores.ListItems.Count
      If lswVendedores.ListItems.Item(i).Checked Then
        iCantidad = iCantidad + 1
      End If
    Next i
    
    If iCantidad <> lswVendedores.ListItems.Count Then
        iCantidad = 0
        vCadena = " and Cod_Vendedor in(0"
        For i = 1 To lswVendedores.ListItems.Count
          If lswVendedores.ListItems.Item(i).Checked Then
            vCadena = vCadena & "," & lswVendedores.ListItems.Item(i).Tag
            iCantidad = iCantidad + 1
          End If
        Next i
        strSQL = strSQL & vCadena & ")"
    End If
End If


If dtpInicio.Enabled Then
    Select Case cboEstado.Text
      Case "Pendientes"
        strSQL = strSQL & " and Estado = 'P' and Registro_Fecha between '" & Format(dtpInicio.Value, "yyyy/mm/dd") _
               & " 00:00:00' and '" & Format(dtpCorte.Value, "yyyy/mm/dd") & " 23:59:59'"
        
      Case "Activadas"
        strSQL = strSQL & " and Estado = 'A' and Activa_Fecha between '" & Format(dtpInicio.Value, "yyyy/mm/dd") _
               & " 00:00:00' and '" & Format(dtpCorte.Value, "yyyy/mm/dd") & " 23:59:59'"
        
      Case "Cerradas"
        strSQL = strSQL & " and Estado = 'C' and Cierra_Fecha between '" & Format(dtpInicio.Value, "yyyy/mm/dd") _
               & " 00:00:00' and '" & Format(dtpCorte.Value, "yyyy/mm/dd") & " 23:59:59'"
      Case Else
        strSQL = strSQL & " and Registro_Fecha between '" & Format(dtpInicio.Value, "yyyy/mm/dd") _
               & " 00:00:00' and '" & Format(dtpCorte.Value, "yyyy/mm/dd") & " 23:59:59'"
      
    End Select
End If

If cboAseguradora.Text <> "TODOS" Then
   strSQL = strSQL & " and COD_ASEGURADORA = '" & cboAseguradora.ItemData(cboAseguradora.ListIndex) & "'"
End If

If cboComercializadora.Text <> "TODOS" Then
   strSQL = strSQL & " and COD_COMERCIALIZADORA = '" & cboComercializadora.ItemData(cboComercializadora.ListIndex) & "'"
End If


Call sbCargaGridLocal(vGrid, 23, strSQL)

Me.MousePointer = vbDefault

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub sbCargaGridLocal(vGrid As Object, vGridMaxCol As Integer, strSQL As String)
Dim i As Integer
Dim curMonto As Currency

On Error GoTo vError

vPaso = True

vGrid.MaxCols = vGridMaxCol
vGrid.MaxRows = 1
vGrid.Row = vGrid.MaxRows
For i = 1 To vGrid.MaxCols
 vGrid.Col = i
 vGrid.Text = ""
Next i

curMonto = 0

Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
  vGrid.Row = vGrid.MaxRows
  For i = 1 To vGrid.MaxCols
    vGrid.Col = i

    If rs.Fields(i - 1).Type = 135 Then
        If Year(rs.Fields(i - 1).Value) > 1900 Then
           vGrid.Text = Format((rs.Fields(i - 1).Value & ""), "dd/mm/yyyy")
        End If
    Else
        vGrid.Text = CStr(rs.Fields(i - 1).Value & "")
    End If
    
  Next i
  vGrid.MaxRows = vGrid.MaxRows + 1
  curMonto = curMonto + rs!Cuota
  rs.MoveNext
Loop

StatusBarX.Panels(1).Text = "Casos ..: " & Format(rs.RecordCount, "###,###,##0")
StatusBarX.Panels(2).Text = "Cuotas..: " & Format(curMonto, "Standard")

rs.Close

vGrid.MaxRows = vGrid.MaxRows - 1

vPaso = False

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub



Private Sub TimerX_Timer()
TimerX.Interval = 0
TimerX.Enabled = False

Call cboAseguradora_Click
Call cboComercializadora_Click
End Sub

Private Sub sbExportar()
Dim vHeaders As vGridHeaders
    vHeaders.Columnas = 23
    vHeaders.Headers(1) = "..."
    vHeaders.Headers(2) = "Aseguradora"
    vHeaders.Headers(3) = "No.Póliza"
    vHeaders.Headers(4) = "Cédula"
    vHeaders.Headers(5) = "Nombre"
    vHeaders.Headers(6) = "Estado"
    vHeaders.Headers(7) = "Mensualidad"
    vHeaders.Headers(8) = "Monto"
    vHeaders.Headers(9) = "Fec.Registro"
    vHeaders.Headers(10) = "Fec.Activación"
    vHeaders.Headers(11) = "Fec.Cierre"
    vHeaders.Headers(12) = "Total Pagado"
    vHeaders.Headers(13) = "Total Cobrado"
    vHeaders.Headers(14) = "Balanza Cobraza"
    vHeaders.Headers(15) = "Comisión Vendedor"
    vHeaders.Headers(16) = "Comisión Comercializa"
    vHeaders.Headers(17) = "Comisión Interna"
    vHeaders.Headers(18) = "No. Operación"
    vHeaders.Headers(19) = "Tipo Seguro"
    vHeaders.Headers(20) = "Vendedor"
    vHeaders.Headers(21) = "Comercializadora"
    vHeaders.Headers(22) = "Cliente Corporativo"
    vHeaders.Headers(23) = "Comision Real Vendedor"
    
   Call sbSIFGridExportar(vGrid, vHeaders, "SEGUROS_ConsultaPolizas")
End Sub


Private Sub txtClienteCorId_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then
   gBusquedas.Columna = "Nombre"
   gBusquedas.Orden = "cod_cliente_Corporativo"
   gBusquedas.Filtro = " and activo = 1"
   gBusquedas.Consulta = "Select cod_cliente_Corporativo,Nombre from SEGUROS_CLIENTE_CORPORATIVO"
   frmBusquedas.Show vbModal
   If gBusquedas.Resultado <> "" Then
      txtClienteCorId.Text = gBusquedas.Resultado
      txtClienteCorNombre.Text = gBusquedas.Resultado2
   End If
End If
End Sub

Private Sub vGrid_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
Dim frm As Form

If vPaso Then Exit Sub

Call sbSIFForms("frmSeguros_Registro")

For Each frm In Forms
  If UCase(frm.Name) = UCase("frmSeguros_Registro") Then
    vGrid.Row = Row
    vGrid.Col = 3
    Call frm.sbConsultaExterna(vGrid.Text)
    Exit For
  End If
Next frm
End Sub
