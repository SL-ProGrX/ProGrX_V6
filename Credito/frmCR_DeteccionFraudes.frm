VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.Ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.Controls.v24.0.0.ocx"
Begin VB.Form frmCR_DeteccionFraudes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Detección de Fraudes"
   ClientHeight    =   6135
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10575
   Icon            =   "frmCR_DeteccionFraudes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6135
   ScaleWidth      =   10575
   Begin TabDlg.SSTab ssTab 
      Height          =   5895
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   10335
      _ExtentX        =   18230
      _ExtentY        =   10398
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      ForeColor       =   16711680
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Inicial"
      TabPicture(0)   =   "frmCR_DeteccionFraudes.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Image1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Detección"
      TabPicture(1)   =   "frmCR_DeteccionFraudes.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "optX(3)"
      Tab(1).Control(1)=   "chkXPersona"
      Tab(1).Control(2)=   "chkFechas"
      Tab(1).Control(3)=   "cboEOperacion"
      Tab(1).Control(4)=   "cboEPersona"
      Tab(1).Control(5)=   "cboUsuarios"
      Tab(1).Control(6)=   "cboDestino"
      Tab(1).Control(7)=   "cboRecurso"
      Tab(1).Control(8)=   "chkLineas"
      Tab(1).Control(9)=   "txtCodigo"
      Tab(1).Control(10)=   "cboGarantia"
      Tab(1).Control(11)=   "cboComite"
      Tab(1).Control(12)=   "txtMeses"
      Tab(1).Control(13)=   "optX(2)"
      Tab(1).Control(14)=   "optX(1)"
      Tab(1).Control(15)=   "txtDias"
      Tab(1).Control(16)=   "optX(0)"
      Tab(1).Control(17)=   "cmdReportes"
      Tab(1).Control(18)=   "dtpInicio"
      Tab(1).Control(19)=   "dtpCorte"
      Tab(1).Control(20)=   "Label3(1)"
      Tab(1).Control(21)=   "Label1(4)"
      Tab(1).Control(22)=   "Label1(5)"
      Tab(1).Control(23)=   "Label1(6)"
      Tab(1).Control(24)=   "Label1(8)"
      Tab(1).Control(25)=   "Label1(10)"
      Tab(1).Control(26)=   "Label1(11)"
      Tab(1).Control(27)=   "Label1(12)"
      Tab(1).Control(28)=   "Label1(13)"
      Tab(1).Control(29)=   "Label1(14)"
      Tab(1).Control(30)=   "Label1(15)"
      Tab(1).Control(31)=   "lblDescripcion"
      Tab(1).Control(32)=   "Label1(16)"
      Tab(1).Control(33)=   "Label1(17)"
      Tab(1).Control(34)=   "Line1(0)"
      Tab(1).Control(35)=   "Label3(0)"
      Tab(1).Control(36)=   "Label2"
      Tab(1).ControlCount=   37
      Begin VB.OptionButton optX 
         Appearance      =   0  'Flat
         Caption         =   "Créditos Anulados dias posteriores a su Formalización"
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
         Height          =   255
         Index           =   3
         Left            =   -74280
         TabIndex        =   34
         Top             =   1680
         Width           =   4815
      End
      Begin VB.CheckBox chkXPersona 
         Appearance      =   0  'Flat
         Caption         =   "x Persona"
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
         Height          =   495
         Left            =   -69120
         TabIndex        =   32
         Top             =   480
         Width           =   1095
      End
      Begin VB.CheckBox chkFechas 
         Appearance      =   0  'Flat
         Caption         =   "Todas"
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
         Height          =   255
         Left            =   -69840
         TabIndex        =   18
         Top             =   2520
         Width           =   1095
      End
      Begin VB.ComboBox cboEOperacion 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   330
         Left            =   -72720
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   2880
         Width           =   2775
      End
      Begin VB.ComboBox cboEPersona 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   330
         Left            =   -72720
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   3240
         Width           =   2775
      End
      Begin VB.ComboBox cboUsuarios 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   330
         Left            =   -73680
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   3600
         Width           =   3735
      End
      Begin VB.ComboBox cboDestino 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   330
         Left            =   -68880
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   3960
         Width           =   3975
      End
      Begin VB.ComboBox cboRecurso 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   330
         Left            =   -68880
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   4320
         Width           =   3975
      End
      Begin VB.CheckBox chkLineas 
         Appearance      =   0  'Flat
         Caption         =   "Todas"
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
         Height          =   216
         Left            =   -68760
         TabIndex        =   12
         Top             =   3300
         Width           =   1095
      End
      Begin VB.TextBox txtCodigo 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   -69720
         MaxLength       =   4
         TabIndex        =   11
         Top             =   3600
         Width           =   855
      End
      Begin VB.ComboBox cboGarantia 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   330
         Left            =   -73680
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   3960
         Width           =   3735
      End
      Begin VB.ComboBox cboComite 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   330
         Left            =   -73680
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   4320
         Width           =   3735
      End
      Begin VB.TextBox txtMeses 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   324
         Left            =   -70320
         TabIndex        =   7
         Text            =   "2"
         Top             =   1320
         Width           =   615
      End
      Begin VB.OptionButton optX 
         Appearance      =   0  'Flat
         Caption         =   "Créditos con 1er. deducción superior ?  "
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
         Height          =   255
         Index           =   2
         Left            =   -74280
         TabIndex        =   6
         Top             =   1320
         Width           =   3495
      End
      Begin VB.OptionButton optX 
         Appearance      =   0  'Flat
         Caption         =   "Créditos de personas sin Aportes"
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
         Height          =   255
         Index           =   1
         Left            =   -74280
         TabIndex        =   5
         Top             =   960
         Width           =   3135
      End
      Begin VB.TextBox txtDias 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   324
         Left            =   -70320
         TabIndex        =   3
         Text            =   "30"
         Top             =   600
         Width           =   615
      End
      Begin VB.OptionButton optX 
         Appearance      =   0  'Flat
         Caption         =   "Créditos Renovados en Menos de "
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
         Height          =   255
         Index           =   0
         Left            =   -74280
         TabIndex        =   2
         Top             =   600
         Value           =   -1  'True
         Width           =   3135
      End
      Begin XtremeSuiteControls.PushButton cmdReportes 
         Height          =   612
         Left            =   -66480
         TabIndex        =   35
         Top             =   5160
         Width           =   1572
         _Version        =   1572864
         _ExtentX        =   2773
         _ExtentY        =   1080
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
         Appearance      =   21
         Picture         =   "frmCR_DeteccionFraudes.frx":0342
      End
      Begin XtremeSuiteControls.DateTimePicker dtpInicio 
         Height          =   312
         Left            =   -73080
         TabIndex        =   36
         Top             =   2520
         Width           =   1332
         _Version        =   1572864
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
         Height          =   312
         Left            =   -71280
         TabIndex        =   37
         Top             =   2520
         Width           =   1332
         _Version        =   1572864
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
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "a"
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
         Index           =   1
         Left            =   -71040
         TabIndex        =   33
         Top             =   1320
         Width           =   615
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
         Left            =   -74760
         TabIndex        =   31
         Top             =   2520
         Width           =   735
      End
      Begin VB.Label Label1 
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
         Left            =   -73680
         TabIndex        =   30
         Top             =   2520
         Width           =   615
      End
      Begin VB.Label Label1 
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
         Left            =   -71760
         TabIndex        =   29
         Top             =   2520
         Width           =   615
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
         Left            =   -74760
         TabIndex        =   28
         Top             =   2880
         Width           =   735
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         Caption         =   "Operación"
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
         Index           =   10
         Left            =   -73680
         TabIndex        =   27
         Top             =   2880
         Width           =   975
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         Caption         =   "Persona"
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
         Index           =   11
         Left            =   -73680
         TabIndex        =   26
         Top             =   3240
         Width           =   975
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         Caption         =   "Línea"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   252
         Index           =   12
         Left            =   -69720
         TabIndex        =   25
         Top             =   3288
         Width           =   852
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         Caption         =   "Destino"
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
         Index           =   13
         Left            =   -69720
         TabIndex        =   24
         Top             =   3960
         Width           =   855
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
         Left            =   -74760
         TabIndex        =   23
         Top             =   3600
         Width           =   735
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         Caption         =   "Recurso"
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
         Index           =   15
         Left            =   -69720
         TabIndex        =   22
         Top             =   4320
         Width           =   855
      End
      Begin VB.Label lblDescripcion 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   -68880
         TabIndex        =   21
         Top             =   3600
         Width           =   3975
      End
      Begin VB.Label Label1 
         Caption         =   "Garantía"
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
         Index           =   16
         Left            =   -74760
         TabIndex        =   20
         Top             =   3960
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Comité"
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
         Index           =   17
         Left            =   -74760
         TabIndex        =   19
         Top             =   4320
         Width           =   735
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   0
         X1              =   -74640
         X2              =   -64920
         Y1              =   2280
         Y2              =   2280
      End
      Begin VB.Label Label3 
         Caption         =   "meses desde su formalización"
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
         Index           =   0
         Left            =   -69600
         TabIndex        =   8
         Top             =   1320
         Width           =   3732
      End
      Begin VB.Label Label2 
         Caption         =   "dias"
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
         Left            =   -69600
         TabIndex        =   4
         Top             =   600
         Width           =   1215
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   480
         Picture         =   "frmCR_DeteccionFraudes.frx":0AFE
         Top             =   720
         Width           =   480
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   $"frmCR_DeteccionFraudes.frx":7350
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   1332
         Index           =   0
         Left            =   1080
         TabIndex        =   1
         Top             =   1800
         Width           =   8172
      End
   End
End
Attribute VB_Name = "frmCR_DeteccionFraudes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub sbLlenaCbo(cboX As ComboBox, strSQL As String, Optional vTodos As Boolean = True)
Dim rs As New ADODB.Recordset

cboX.Clear

Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
 cboX.AddItem rs!itmX
 rs.MoveNext
Loop
If rs.RecordCount > 0 Then
   rs.MoveFirst
   cboX.Text = rs!itmX
End If
rs.Close

If vTodos Then
    cboX.AddItem "TODOS"
    cboX.Text = "TODOS"
End If

End Sub


Private Sub chkFechas_Click()

If chkFechas.Value = vbChecked Then
   dtpInicio.Enabled = False
Else
   dtpInicio.Enabled = True
End If

dtpCorte.Enabled = dtpInicio.Enabled

End Sub

Private Sub chkLineas_Click()
Dim strSQL As String, rs As New ADODB.Recordset

If chkLineas.Value = vbChecked Then
  
  txtCodigo.Enabled = False
  
  strSQL = "select cod_grupo + ' - ' + rtrim(descripcion) as ItmX" _
         & " from  catalogo_grupos"
  Call sbLlenaCbo(cboRecurso, strSQL)
  
  strSQL = "select cod_destino + ' - ' + rtrim(descripcion) as ItmX" _
         & " from  catalogo_destinos"
  Call sbLlenaCbo(cboDestino, strSQL)
  
Else
  txtCodigo.Enabled = True

  strSQL = "select (R.cod_grupo) + ' - ' + rtrim(R.descripcion) as ItmX" _
         & " from catalogo_grupos R inner join catalogo_AsignaGrp A on R.cod_grupo = A.cod_grupo" _
         & " where A.codigo = '" & txtCodigo & "'"
  Call sbLlenaCbo(cboRecurso, strSQL)
  
  strSQL = "select (R.cod_destino) + ' - ' + rtrim(R.descripcion) as ItmX" _
         & " from catalogo_destinos R inner join catalogo_destinosAsg A on R.cod_destino = A.cod_destino" _
         & " where A.codigo = '" & txtCodigo & "'"
  Call sbLlenaCbo(cboDestino, strSQL)

End If

End Sub


Private Sub cmdReportes_Click()
Dim vTitulo As String, vSubTitulo As String, vFiltro As String
Dim strSQL As String, vTipoUser As String

On Error GoTo vError

Me.MousePointer = vbHourglass


Select Case True
 Case OptX.Item(0).Value
   vTitulo = UCase(OptX.Item(0).Caption) & " " & txtDias.Text & " dias"
 Case OptX.Item(1).Value
   vTitulo = UCase(OptX.Item(1).Caption)
 Case OptX.Item(2).Value
   vTitulo = UCase(OptX.Item(2).Caption) & " a " & txtMeses.Text & " meses"
 Case OptX.Item(3).Value
   vTitulo = UCase(OptX.Item(3).Caption)
End Select


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
 .WindowTitle = "Reportes del Módulo de Créditos"
 
 .Connect = glogon.ConectRPT
 
 If chkFechas.Value = vbUnchecked Then
    If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
        strSQL = strSQL & "{vCRDCreditosReportes01.fechaforp} in Date(" & Format(dtpInicio.Value, "yyyy,mm,dd") & ") to date(" _
               & Format(dtpCorte.Value, "yyyy,mm,dd") & ")"
        vSubTitulo = "Formalizadas entre " & Format(dtpInicio.Value, "dd/mm/yyyy") & " y " & Format(dtpCorte.Value, "dd/mm/yyyy")
 Else
   vSubTitulo = "Historico"
 End If
 

If cboEOperacion.Enabled Then
   If Mid(cboEOperacion.Text, 1, 1) = "T" Then
     If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
     strSQL = strSQL & "({vCRDCreditosReportes01.estado} = 'A' OR {vCRDCreditosReportes01.estado} = 'C')"
     vSubTitulo = vSubTitulo & " / Estado Operación : " & cboEOperacion.Text
   Else
     If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
     strSQL = strSQL & "{vCRDCreditosReportes01.estado} = '" & Mid(cboEOperacion.Text, 1, 1) & "'"
     vSubTitulo = vSubTitulo & " / Estado Operación : " & cboEOperacion.Text
   End If
End If
    
    
 Select Case Mid(cboEPersona.Text, 1, 2)
     Case "01" 'Socios
       If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
       strSQL = strSQL & "{vCRDCreditosReportes01.estadoactual} = 'S'"
     Case "02" 'Opex
       If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
       strSQL = strSQL & "({vCRDCreditosReportes01.estadoactual} = 'A' OR {vCRDCreditosReportes01.estadoactual} = 'P')"
     Case "03" 'No Socios
       If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
       strSQL = strSQL & "{vCRDCreditosReportes01.estadoactual} = 'N'"
     Case "04" 'Ren.Interna
       If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
       strSQL = strSQL & "{vCRDCreditosReportes01.estadoactual} = 'A'"
     Case "05" 'Ren.Patronal
       If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
       strSQL = strSQL & "{vCRDCreditosReportes01.estadoactual} = 'P'"
 End Select
 vFiltro = vFiltro & "/ ESTADO PERSONA : " & cboEPersona.Text
    
 If Mid(cboGarantia.Text, 1, 1) <> "T" Then
   If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
   strSQL = strSQL & "{vCRDCreditosReportes01.garantia} = '" & Mid(cboGarantia.Text, 1, 1) & "'"
 End If
 vSubTitulo = vSubTitulo & " / Garantía : " & cboGarantia.Text
 
 If cboUsuarios.Text <> "TODOS" Then
   If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
   strSQL = strSQL & "{vCRDCreditosReportes01.COD_GRUPO} = '" & fxCodigoCbo(cboUsuarios) & "'"
 End If
 vFiltro = "/ GRUPOS : " & cboUsuarios
 
 If chkLineas.Value = vbUnchecked Then
   If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
   strSQL = strSQL & "{vCRDCreditosReportes01.Codigo} = '" & Trim(txtCodigo) & "'"
   vFiltro = vFiltro & "/ LINEA : " & UCase(txtCodigo) & " - " & lblDescripcion.Caption
 Else
   vFiltro = vFiltro & "/ TODAS LAS LINEAS"
 End If
 
 If cboRecurso.Text <> "TODOS" Then
   If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
   strSQL = strSQL & "{vCRDCreditosReportes01.recurso} = '" & Trim(fxCodigoCbo(cboRecurso)) & "'"
 End If
 vFiltro = vFiltro & "/ RECURSO : " & cboRecurso.Text
 
 If cboDestino.Text <> "TODOS" Then
   If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
   strSQL = strSQL & "{vCRDCreditosReportes01.cod_destino} = '" & fxCodigoCbo(cboDestino) & "'"
 End If
 vFiltro = vFiltro & "/ DESTINO : " & cboDestino.Text
 
 If cboComite.Text <> "TODOS" Then
   If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
   strSQL = strSQL & "{vCRDCreditosReportes01.id_comite} = " & cboComite.ItemData(cboComite.ListIndex) & ""
 End If
 vFiltro = vFiltro & "/ COMITE : " & cboComite.Text
 
 .Formulas(0) = "fxFecha='FECHA: " & Format(fxFechaServidor, "dd/mm/yyyy") & "'"
 .Formulas(1) = "fxEmpresa='" & GLOBALES.gstrNombreEmpresa & "'"
 .Formulas(2) = "fxUsuario='USER: " & glogon.Usuario & "'"
 .Formulas(3) = "fxTitulo='" & vTitulo & "'"
 .Formulas(4) = "fxSubTitulo='" & vSubTitulo & "'"
 .Formulas(5) = "fxFiltro='" & vFiltro & "'"



  Select Case True
     Case OptX.Item(0).Value
       vTitulo = UCase(OptX.Item(0).Caption)
       'Corre procedimiento almacenado.
       If chkFechas.Value = vbUnchecked Then
          glogon.Conection.Execute "exec spCRDReporteRenovacion '" & Format(dtpInicio.Value, "yyyy/mm/dd") & "','" & Format(dtpCorte.Value, "yyyy/mm/dd") _
                           & "','" & glogon.Usuario & "'," & txtDias.Text
       Else
          glogon.Conection.Execute "exec spCRDReporteRenovacion '1940/01/01','" & Format(Date, "yyyy/mm/dd") & "','" _
                           & glogon.Usuario & "'," & txtDias.Text
       End If
       
        If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
        strSQL = strSQL & "{vCRDCreditosReportes01.estadosol} = 'F' and {vCRDCreditosReportes01.id_solicitud} = {CRD_REPORTES_TMP01.id_solicitud}"
       If chkXPersona.Value = vbUnchecked Then
              .ReportFileName = SIFGlobal.fxPathReportes("Credito_CreditosRenovacion.rpt")
       Else
              .ReportFileName = SIFGlobal.fxPathReportes("Credito_CreditosRenovacionPrs.rpt")
       End If
     
     Case OptX.Item(1).Value
        If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
        strSQL = strSQL & "{vCRDCreditosReportes01.estadosol} = 'F' AND {AHORRO_CONSOLIDADO.AHORRO} = 0"
       .ReportFileName = SIFGlobal.fxPathReportes("Credito_CreditosSinAhorro.rpt")
     
     Case OptX.Item(2).Value
        If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
        strSQL = strSQL & "{vCRDCreditosReportes01.estadosol} = 'F' AND {vCRDCreditosReportes01.DifFecha} > " & txtMeses.Text
       .ReportFileName = SIFGlobal.fxPathReportes("Credito_CreditosPriDeduc.rpt")


     Case OptX.Item(3).Value
       
       vTitulo = UCase(OptX.Item(3).Caption)
       'Corre procedimiento almacenado.
       If chkFechas.Value = vbUnchecked Then
          glogon.Conection.Execute "exec spCRDReporteAnulados '" & Format(dtpInicio.Value, "yyyy/mm/dd") & "','" & Format(dtpCorte.Value, "yyyy/mm/dd") _
                           & "','" & glogon.Usuario & "'"
       Else
          glogon.Conection.Execute "exec spCRDReporteAnulados '1940/01/01','" & Format(Date, "yyyy/mm/dd") & "','" _
                           & glogon.Usuario & "'"
       End If
       
        If Len(strSQL) > 0 Then strSQL = strSQL & " AND "
        strSQL = strSQL & "{vCRDCreditosReportes01.estadosol} = 'N' and {vCRDCreditosReportes01.id_solicitud} = {CRD_REPORTES_TMP01.id_solicitud}"
        .ReportFileName = SIFGlobal.fxPathReportes("Credito_CreditosRenovacion.rpt")


  End Select

 .SelectionFormula = strSQL
 .PrintReport

End With

Me.MousePointer = vbDefault
Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub Form_Activate()
  vModulo = 3
End Sub

Private Sub Form_Load()
 vModulo = 3
 Call Formularios(Me)
 Call RefrescaTags(Me)
 Call sbInicializa
End Sub

Private Sub sbInicializa()
Dim strSQL As String, rs As New ADODB.Recordset

Me.MousePointer = vbHourglass

ssTab.Tab = 0

ssTab.TabEnabled(1) = cmdReportes.Enabled

dtpInicio.Value = fxFechaServidor
dtpCorte.Value = dtpInicio.Value
chkFechas.Value = vbUnchecked
chkLineas.Value = vbChecked

cboGarantia.Clear
cboGarantia.AddItem "A - Sobre Ahorros"
cboGarantia.AddItem "F - Fiduciaria"
cboGarantia.AddItem "H - Hipotecaria"
cboGarantia.AddItem "X - Acciones"
cboGarantia.AddItem "Y - Fondos de Inversion"
cboGarantia.AddItem "N - Sin Garantía"
cboGarantia.AddItem "TODOS"
cboGarantia.Text = "TODOS"

cboEOperacion.Clear
cboEOperacion.AddItem "Activa"
cboEOperacion.AddItem "Cancelada"
cboEOperacion.AddItem "Nulas"
cboEOperacion.AddItem "Todas (Activas/Canceladas)"
cboEOperacion.Text = "Activa"

cboEPersona.Clear
cboEPersona.AddItem "00 - Todos"
cboEPersona.AddItem "01 - Socios"
cboEPersona.AddItem "02 - Ex.Socios"
cboEPersona.AddItem "03 - No Socios"
cboEPersona.AddItem "04 - Ren.Interna"
cboEPersona.AddItem "05 - Ren.Patronal"
cboEPersona.Text = "00 - Todos"

cboComite.Clear

strSQL = "select id_comite,descripcion from comites "
Call OpenRecordSet(rs, strSQL)

Do While Not rs.EOF
   cboComite.AddItem Trim(rs!Descripcion)
   cboComite.ItemData(cboComite.NewIndex) = rs!id_Comite
   rs.MoveNext
Loop
cboComite.AddItem "TODOS"
cboComite.Text = "TODOS"
rs.Close

strSQL = "select cod_grupo + ' - ' + rtrim(descripcion) as ItmX" _
         & " from  crd_grupos"
Call sbLlenaCbo(cboUsuarios, strSQL)


Call chkFechas_Click
Call chkLineas_Click

Me.MousePointer = vbDefault

End Sub



Private Sub txtCodigo_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyReturn Then cboDestino.SetFocus

If KeyCode = vbKeyF4 Then
  gBusquedas.Resultado = ""
  gBusquedas.Resultado2 = ""
  gBusquedas.Convertir = "N"
  gBusquedas.Resultado = ""
  gBusquedas.Consulta = "select codigo,descripcion from catalogo"
  gBusquedas.Orden = "descripcion"
  gBusquedas.Columna = "descripcion"
  frmBusquedas.Show vbModal
  txtCodigo = gBusquedas.Resultado
  lblDescripcion.Caption = gBusquedas.Resultado2
End If

End Sub

Private Sub txtCodigo_LostFocus()
 If Len(Trim(txtCodigo)) > 0 Then lblDescripcion.Caption = fxDescribeCodigo(Trim(txtCodigo))
 Call chkLineas_Click
End Sub

