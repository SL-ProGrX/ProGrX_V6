VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.Controls.v24.0.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#24.0#0"; "Codejock.ShortcutBar.v24.0.0.ocx"
Begin VB.Form frmCO_Cobro_Fiadores 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cobro a Fiadores"
   ClientHeight    =   8640
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   14175
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8640
   ScaleWidth      =   14175
   Begin VB.Timer TimerX 
      Interval        =   5
      Left            =   0
      Top             =   0
   End
   Begin XtremeSuiteControls.ProgressBar ProgressBarX 
      Height          =   135
      Left            =   0
      TabIndex        =   2
      Top             =   1080
      Visible         =   0   'False
      Width           =   14175
      _Version        =   1572864
      _ExtentX        =   25003
      _ExtentY        =   238
      _StockProps     =   93
      BackColor       =   -2147483643
   End
   Begin XtremeSuiteControls.TabControl tcMain 
      Height          =   6855
      Left            =   0
      TabIndex        =   1
      Top             =   1800
      Width           =   14175
      _Version        =   1572864
      _ExtentX        =   25003
      _ExtentY        =   12091
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
      Item(0).Caption =   "Gestión de Cobro a Fiadores"
      Item(0).ControlCount=   10
      Item(0).Control(0)=   "lsw"
      Item(0).Control(1)=   "btnExportar(0)"
      Item(0).Control(2)=   "txtFiltro(0)"
      Item(0).Control(3)=   "scAcuerdos"
      Item(0).Control(4)=   "chkTodos(0)"
      Item(0).Control(5)=   "Label2(0)"
      Item(0).Control(6)=   "cboCuotas"
      Item(0).Control(7)=   "btnBuscar(0)"
      Item(0).Control(8)=   "btnProcesaCobros"
      Item(0).Control(9)=   "btnNotificaAdvertencia"
      Item(1).Caption =   "Activos"
      Item(1).ControlCount=   7
      Item(1).Control(0)=   "btnExportar(1)"
      Item(1).Control(1)=   "txtFiltro(1)"
      Item(1).Control(2)=   "ShortcutCaption1"
      Item(1).Control(3)=   "chkTodos(1)"
      Item(1).Control(4)=   "lswActivos"
      Item(1).Control(5)=   "btnReversa"
      Item(1).Control(6)=   "btnBuscar(1)"
      Begin XtremeSuiteControls.ListView lsw 
         Height          =   5415
         Left            =   120
         TabIndex        =   3
         Top             =   1200
         Width           =   13935
         _Version        =   1572864
         _ExtentX        =   24580
         _ExtentY        =   9551
         _StockProps     =   77
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Checkboxes      =   -1  'True
         View            =   3
         FullRowSelect   =   -1  'True
         Appearance      =   21
      End
      Begin XtremeSuiteControls.ListView lswActivos 
         Height          =   5415
         Left            =   -69880
         TabIndex        =   7
         Top             =   1320
         Visible         =   0   'False
         Width           =   13935
         _Version        =   1572864
         _ExtentX        =   24580
         _ExtentY        =   9551
         _StockProps     =   77
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Checkboxes      =   -1  'True
         View            =   3
         FullRowSelect   =   -1  'True
         Appearance      =   21
      End
      Begin XtremeSuiteControls.PushButton btnBuscar 
         Height          =   375
         Index           =   0
         Left            =   4440
         TabIndex        =   15
         Top             =   435
         Width           =   1215
         _Version        =   1572864
         _ExtentX        =   2143
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Buscar"
         BackColor       =   -2147483643
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
         Picture         =   "frmCO_Cobro_Fiadores.frx":0000
      End
      Begin XtremeSuiteControls.CheckBox chkTodos 
         Height          =   210
         Index           =   0
         Left            =   360
         TabIndex        =   11
         Top             =   920
         Width           =   210
         _Version        =   1572864
         _ExtentX        =   370
         _ExtentY        =   370
         _StockProps     =   79
         Caption         =   "CheckBox1"
         BackColor       =   -2147483643
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
         Appearance      =   21
      End
      Begin XtremeSuiteControls.PushButton btnExportar 
         Height          =   375
         Index           =   0
         Left            =   13560
         TabIndex        =   4
         ToolTipText     =   "Exportar Lista"
         Top             =   840
         Width           =   495
         _Version        =   1572864
         _ExtentX        =   873
         _ExtentY        =   661
         _StockProps     =   79
         BackColor       =   -2147483643
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
         Picture         =   "frmCO_Cobro_Fiadores.frx":0700
      End
      Begin XtremeSuiteControls.FlatEdit txtFiltro 
         Height          =   375
         Index           =   0
         Left            =   2400
         TabIndex        =   5
         Top             =   840
         Width           =   11175
         _Version        =   1572864
         _ExtentX        =   19711
         _ExtentY        =   661
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   16777152
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16777152
         Alignment       =   2
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.PushButton btnExportar 
         Height          =   375
         Index           =   1
         Left            =   -56440
         TabIndex        =   8
         ToolTipText     =   "Exportar Lista"
         Top             =   960
         Visible         =   0   'False
         Width           =   495
         _Version        =   1572864
         _ExtentX        =   873
         _ExtentY        =   661
         _StockProps     =   79
         BackColor       =   -2147483643
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
         Picture         =   "frmCO_Cobro_Fiadores.frx":086A
      End
      Begin XtremeSuiteControls.FlatEdit txtFiltro 
         Height          =   375
         Index           =   1
         Left            =   -67600
         TabIndex        =   9
         Top             =   960
         Visible         =   0   'False
         Width           =   11175
         _Version        =   1572864
         _ExtentX        =   19711
         _ExtentY        =   661
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   16777152
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   16777152
         Alignment       =   2
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.CheckBox chkTodos 
         Height          =   210
         Index           =   1
         Left            =   -69640
         TabIndex        =   12
         Top             =   1040
         Visible         =   0   'False
         Width           =   210
         _Version        =   1572864
         _ExtentX        =   370
         _ExtentY        =   370
         _StockProps     =   79
         Caption         =   "CheckBox1"
         BackColor       =   -2147483643
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
         Appearance      =   21
      End
      Begin XtremeSuiteControls.ComboBox cboCuotas 
         Height          =   330
         Left            =   3120
         TabIndex        =   14
         Top             =   440
         Width           =   1215
         _Version        =   1572864
         _ExtentX        =   2143
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
      Begin XtremeSuiteControls.PushButton btnProcesaCobros 
         Height          =   375
         Left            =   9480
         TabIndex        =   16
         Top             =   435
         Width           =   1815
         _Version        =   1572864
         _ExtentX        =   3201
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Procesa Cobros"
         BackColor       =   -2147483643
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
         Picture         =   "frmCO_Cobro_Fiadores.frx":09D4
      End
      Begin XtremeSuiteControls.PushButton btnNotificaAdvertencia 
         Height          =   375
         Left            =   11280
         TabIndex        =   17
         Top             =   435
         Width           =   2775
         _Version        =   1572864
         _ExtentX        =   4895
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Notifica Advertencia"
         BackColor       =   -2147483643
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
         Picture         =   "frmCO_Cobro_Fiadores.frx":10ED
      End
      Begin XtremeSuiteControls.PushButton btnReversa 
         Height          =   375
         Left            =   -66400
         TabIndex        =   22
         Top             =   480
         Visible         =   0   'False
         Width           =   2775
         _Version        =   1572864
         _ExtentX        =   4895
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Reversa Cobro a Fiadores"
         BackColor       =   -2147483643
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
         Picture         =   "frmCO_Cobro_Fiadores.frx":1258
      End
      Begin XtremeSuiteControls.PushButton btnBuscar 
         Height          =   375
         Index           =   1
         Left            =   -67600
         TabIndex        =   23
         Top             =   480
         Visible         =   0   'False
         Width           =   1215
         _Version        =   1572864
         _ExtentX        =   2143
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Buscar"
         BackColor       =   -2147483643
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
         Picture         =   "frmCO_Cobro_Fiadores.frx":1958
      End
      Begin XtremeSuiteControls.Label Label2 
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   13
         Top             =   440
         Width           =   3015
         _Version        =   1572864
         _ExtentX        =   5318
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Cantidad de Cuotas atrasadas >="
         BackColor       =   -2147483643
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
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
         Height          =   375
         Left            =   -69880
         TabIndex        =   10
         Top             =   960
         Visible         =   0   'False
         Width           =   2295
         _Version        =   1572864
         _ExtentX        =   4048
         _ExtentY        =   661
         _StockProps     =   14
         Caption         =   "Filtrar: "
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
         Alignment       =   2
      End
      Begin XtremeShortcutBar.ShortcutCaption scAcuerdos 
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   840
         Width           =   2295
         _Version        =   1572864
         _ExtentX        =   4048
         _ExtentY        =   661
         _StockProps     =   14
         Caption         =   "Filtrar: "
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
         Alignment       =   2
      End
   End
   Begin XtremeSuiteControls.ComboBox cboInstitucion 
      Height          =   330
      Left            =   1440
      TabIndex        =   19
      Top             =   1320
      Width           =   3255
      _Version        =   1572864
      _ExtentX        =   5741
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
   Begin XtremeSuiteControls.ComboBox cboEstadoPersona 
      Height          =   330
      Left            =   6600
      TabIndex        =   21
      Top             =   1320
      Width           =   2175
      _Version        =   1572864
      _ExtentX        =   3836
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
   Begin XtremeSuiteControls.Label Label2 
      Height          =   255
      Index           =   2
      Left            =   4800
      TabIndex        =   20
      Top             =   1320
      Width           =   1815
      _Version        =   1572864
      _ExtentX        =   3201
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Estado de la Persona"
      BackColor       =   -2147483643
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
   Begin XtremeSuiteControls.Label Label2 
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   18
      Top             =   1320
      Width           =   1335
      _Version        =   1572864
      _ExtentX        =   2355
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Instituciones"
      BackColor       =   -2147483643
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
   Begin XtremeSuiteControls.Label Label1 
      Height          =   495
      Index           =   1
      Left            =   1920
      TabIndex        =   0
      Top             =   360
      Width           =   5175
      _Version        =   1572864
      _ExtentX        =   9128
      _ExtentY        =   873
      _StockProps     =   79
      Caption         =   "Cobro a Fiadores"
      ForeColor       =   16777215
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   14.25
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
      Height          =   1095
      Left            =   0
      Top             =   0
      Width           =   14175
   End
End
Attribute VB_Name = "frmCO_Cobro_Fiadores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem
Dim vPaso As Boolean

Private Sub sbCobros_Pendientes_Load()

On Error GoTo vError

Me.MousePointer = vbHourglass

Dim pInstitucion As String, pEstadoPersona As String, pFiltro As String


If cboInstitucion.Text = "TODOS" Then
  pInstitucion = "Null"
Else
  pInstitucion = cboInstitucion.ItemData(cboInstitucion.ListIndex)
End If


If cboEstadoPersona.Text = "TODOS" Then
  pEstadoPersona = "Null"
Else
  pEstadoPersona = cboEstadoPersona.ItemData(cboEstadoPersona.ListIndex)
End If

pFiltro = "'" & fxSysCleanTxtInject(txtFiltro(0).Text) & "'"

vPaso = True


lsw.ListItems.Clear
'spCBR_Cobro_Fiadores_Pendientes(@Institucion int = Null, @EstadoPersona varchar(10) = Null, @Filtro varchar(200) = '', @NCuotas smallint = 2)
strSQL = "exec spCBR_Cobro_Fiadores_Pendientes " & pInstitucion & ", " & pEstadoPersona _
       & ", " & pFiltro & ", " & cboCuotas.Text
Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
  Set itmX = lsw.ListItems.Add(, , rs!ID_SOLICITUD)
      itmX.SubItems(1) = rs!Codigo
      itmX.SubItems(2) = rs!Cedula
      itmX.SubItems(3) = rs!Nombre
      itmX.SubItems(4) = rs!N_Cuota
      itmX.SubItems(5) = Format(rs!Mora_Financiera, "Standard")
      itmX.SubItems(6) = Format(rs!Saldo, "Standard")
      itmX.SubItems(7) = rs!NOTIFICA_FECHA & ""
      itmX.SubItems(8) = rs!EstadoPersona_Desc
      itmX.SubItems(9) = rs!Linea_Desc
      itmX.SubItems(10) = rs!Institucion_Desc
  rs.MoveNext
Loop
rs.Close

vPaso = True

Me.MousePointer = vbDefault

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub


Private Sub sbCobros_Activos_Load()

On Error GoTo vError

Me.MousePointer = vbHourglass

Dim pInstitucion As String, pEstadoPersona As String, pFiltro As String


If cboInstitucion.Text = "TODOS" Then
  pInstitucion = "Null"
Else
  pInstitucion = cboInstitucion.ItemData(cboInstitucion.ListIndex)
End If


If cboEstadoPersona.Text = "TODOS" Then
  pEstadoPersona = "Null"
Else
  pEstadoPersona = cboEstadoPersona.ItemData(cboEstadoPersona.ListIndex)
End If

pFiltro = "'" & fxSysCleanTxtInject(txtFiltro(1).Text) & "'"

vPaso = True

lswActivos.ListItems.Clear



strSQL = "exec spCBR_Cobro_Fiadores_Activos " & pInstitucion & ", " & pEstadoPersona _
       & ", " & pFiltro
Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
  Set itmX = lswActivos.ListItems.Add(, , rs!ID_SOLICITUD)
      itmX.SubItems(1) = rs!Codigo
      itmX.SubItems(2) = rs!Cedula
      itmX.SubItems(3) = rs!Nombre
      itmX.SubItems(4) = Format(rs!Cuota, "Standard")
      itmX.SubItems(5) = rs!D_Operacion
      itmX.SubItems(6) = rs!D_Codigo
      itmX.SubItems(7) = rs!D_Cedula
      itmX.SubItems(8) = rs!D_Nombre
      itmX.SubItems(9) = rs!EstadoPersona_Desc
      itmX.SubItems(10) = rs!Linea_Desc

  rs.MoveNext
Loop
rs.Close

vPaso = True

Me.MousePointer = vbDefault

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub



Private Sub btnBuscar_Click(Index As Integer)

If Index = 0 Then
    Call sbCobros_Pendientes_Load
Else
    Call sbCobros_Activos_Load
End If

End Sub

Private Sub btnExportar_Click(Index As Integer)
On Error GoTo vError

Me.MousePointer = vbHourglass

ProgressBarX.Visible = True

If Index = 0 Then
    Call Excel_Exportar_Lsw(lsw, ProgressBarX)
Else
    Call Excel_Exportar_Lsw(lswActivos, ProgressBarX)
End If

ProgressBarX.Visible = False

Me.MousePointer = vbDefault

Exit Sub

vError:
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub btnNotificaAdvertencia_Click()
Dim i As Long, iCasos As Long

iCasos = 0
With lsw.ListItems
For i = 1 To .Count
 If .Item(i).Checked Then

    iCasos = iCasos + 1
 End If
Next i
End With

If iCasos = 0 Then
    MsgBox "Debe Seleccionar al menos un caso!", vbExclamation
    Exit Sub
End If


strSQL = "select count(*) as 'Existe'" _
       & "  From CATALOGO " _
       & "  Where codigo in(" _
       & "     select valor From CBR_PARAMETROS" _
       & "     Where COD_PARAMETRO = '25' )"
Call OpenRecordSet(rs, strSQL)
If rs!Existe = 0 Then
    MsgBox "No se encuentra configurada la Línea/Retención para Cobro a Fiador, verifique los parámetros de cobro!", vbExclamation
    Exit Sub
End If
rs.Close

i = MsgBox("Esta seguro que desea Notificar Advertencia de Cobro a Fiadores de los casos seleccionados (Deudor y Fiadores) ?", vbYesNo)
If i = vbNo Then Exit Sub


On Error GoTo vError

Me.MousePointer = vbHourglass

ProgressBarX.Value = 0
ProgressBarX.Max = iCasos
ProgressBarX.Visible = True


iCasos = 0

strSQL = ""

With lsw.ListItems
For i = 1 To .Count
 If .Item(i).Checked Then

    strSQL = "exec spCBR_Cobro_Fiadores_Notifica " & .Item(i).Text & ", '" & glogon.Usuario & "'"
    
    If Len(strSQL) > 20000 Then
        Call ConectionExecute(strSQL)
        strSQL = ""
    End If
    
    iCasos = iCasos + 1
    If ProgressBarX.Value < ProgressBarX.Max Then
        ProgressBarX.Value = ProgressBarX.Value + 1
    End If
 End If
Next i
End With

If Len(strSQL) > 0 Then
    Call ConectionExecute(strSQL)
    strSQL = ""
End If

Me.MousePointer = vbDefault

MsgBox "Notificaciones Enviadas a Deudores y Fiadores!", vbInformation

Call sbCobros_Pendientes_Load

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub btnProcesaCobros_Click()
Dim i As Long, iCasos As Long

iCasos = 0
With lsw.ListItems
For i = 1 To .Count
 If .Item(i).Checked Then

    iCasos = iCasos + 1
 End If
Next i
End With

If iCasos = 0 Then
    MsgBox "Debe Seleccionar al menos un caso!", vbExclamation
    Exit Sub
End If


strSQL = "select count(*) as 'Existe'" _
       & "  From CATALOGO " _
       & "  Where codigo in(" _
       & "     select valor From CBR_PARAMETROS" _
       & "     Where COD_PARAMETRO = '25' )"
Call OpenRecordSet(rs, strSQL)
If rs!Existe = 0 Then
    MsgBox "No se encuentra configurada la Línea/Retención para Cobro a Fiador, verifique los parámetros de cobro!", vbExclamation
    Exit Sub
End If
rs.Close

i = MsgBox("Esta seguro que desea procesar el Cobro a Fiadores de los casos seleccionados?", vbYesNo)
If i = vbNo Then Exit Sub


On Error GoTo vError

Me.MousePointer = vbHourglass

ProgressBarX.Value = 0
ProgressBarX.Max = iCasos
ProgressBarX.Visible = True


iCasos = 0

strSQL = ""

With lsw.ListItems
For i = 1 To .Count
 If .Item(i).Checked Then

    strSQL = "exec spCBR_Cobro_Fiadores_Procesa " & .Item(i).Text & ", '" & glogon.Usuario & "'"
    
    If Len(strSQL) > 20000 Then
        Call ConectionExecute(strSQL)
        strSQL = ""
    End If
    
    iCasos = iCasos + 1
    If ProgressBarX.Value < ProgressBarX.Max Then
        ProgressBarX.Value = ProgressBarX.Value + 1
    End If
 End If
Next i
End With

If Len(strSQL) > 0 Then
    Call ConectionExecute(strSQL)
    strSQL = ""
End If

Me.MousePointer = vbDefault

MsgBox "Cobros a Fiadores activado satisfactoriamente!", vbInformation

Call sbCobros_Pendientes_Load

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub btnReversa_Click()
Dim i As Long, iCasos As Long

iCasos = 0
With lswActivos.ListItems
For i = 1 To .Count
 If .Item(i).Checked Then

    iCasos = iCasos + 1
 End If
Next i
End With

If iCasos = 0 Then
    MsgBox "Debe Seleccionar al menos un caso!", vbExclamation
    Exit Sub
End If


strSQL = "select count(*) as 'Existe'" _
       & "  From CATALOGO " _
       & "  Where codigo in(" _
       & "     select valor From CBR_PARAMETROS" _
       & "     Where COD_PARAMETRO = '25' )"
Call OpenRecordSet(rs, strSQL)
If rs!Existe = 0 Then
    MsgBox "No se encuentra configurada la Línea/Retención para Cobro a Fiador, verifique los parámetros de cobro [25]", vbExclamation
    Exit Sub
End If
rs.Close

strSQL = "select count(*) as 'Existe'" _
       & "  From FND_PLANES " _
       & "  Where COD_PLAN in(" _
       & "     select valor From CBR_PARAMETROS" _
       & "     Where COD_PARAMETRO = '27' )"
Call OpenRecordSet(rs, strSQL)
If rs!Existe = 0 Then
    MsgBox "No se encuentra configurado el Fondo de Devolución para Cobro a Fiador, verifique los parámetros de cobro [27]", vbExclamation
    Exit Sub
End If
rs.Close


i = MsgBox("Esta seguro que desea Reversar el Cobro a Fiadores de los casos seleccionados?", vbYesNo)
If i = vbNo Then Exit Sub


On Error GoTo vError

Me.MousePointer = vbHourglass

ProgressBarX.Value = 0
ProgressBarX.Max = iCasos
ProgressBarX.Visible = True


iCasos = 0

strSQL = ""

With lswActivos.ListItems
For i = 1 To .Count
 If .Item(i).Checked Then

    strSQL = "exec spCBR_Cobro_Fiadores_Reversa " & .Item(i).Text & ", '" & glogon.Usuario & "'"
    
    If Len(strSQL) > 20000 Then
        Call ConectionExecute(strSQL)
        strSQL = ""
    End If
    
    iCasos = iCasos + 1
    If ProgressBarX.Value < ProgressBarX.Max Then
        ProgressBarX.Value = ProgressBarX.Value + 1
    End If
 End If
Next i
End With

If Len(strSQL) > 0 Then
    Call ConectionExecute(strSQL)
    strSQL = ""
End If

Me.MousePointer = vbDefault

MsgBox "Reversión de Cobro a Fiadores, realizada Satisfactoriamente!", vbInformation

Call sbCobros_Activos_Load

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub chkTodos_Click(Index As Integer)

Dim i As Long

vPaso = True

If Index = 0 Then
        For i = 1 To lsw.ListItems.Count
          lsw.ListItems.Item(i).Checked = chkTodos(Index).Value
        Next i
Else
        For i = 1 To lswActivos.ListItems.Count
          lswActivos.ListItems.Item(i).Checked = chkTodos(Index).Value
        Next i
End If

vPaso = False
End Sub

Private Sub Form_Load()
vModulo = 4

Set imgBanner.Picture = frmContenedor.imgBanner_Procesar.Picture

tcMain.Item(0).Selected = True

cboCuotas.AddItem "3"
cboCuotas.AddItem "4"
cboCuotas.AddItem "5"
cboCuotas.AddItem "6"
cboCuotas.Text = "3"


With lsw.ColumnHeaders
    .Clear
    .Add , , "No. Operación", 1200
    .Add , , "Código", 1200, vbCenter
    .Add , , "Cédula", 1800, vbCenter
    .Add , , "Nombre", 3800
    .Add , , "No.Cuotas", 1200, vbCenter
    .Add , , "Mnt. Atraso", 2500, vbRightJustify
    .Add , , "Saldo", 2500, vbRightJustify
    
    .Add , , "Ult.Notifica.", 2500
    .Add , , "Est.Persona", 2800, vbCenter
    .Add , , "Línea Descripción.", 3500
    .Add , , "Institución", 3500
End With



With lswActivos.ColumnHeaders
    .Clear
    .Add , , "No. Operación", 1200
    .Add , , "Código", 1200, vbCenter
    .Add , , "Cédula", 1800, vbCenter
    .Add , , "Nombre", 3800
    .Add , , "Cuota al Cobro", 2500, vbRightJustify
    
    .Add , , "D. Operación", 1200
    .Add , , "D. Código", 1200, vbCenter
    .Add , , "D. Cédula", 1800, vbCenter
    .Add , , "D. Nombre", 3800
    
    .Add , , "Est.Persona", 2800, vbCenter
    .Add , , "Línea Descripción.", 3500
End With



Call Formularios(Me)
Call RefrescaTags(Me)

End Sub


Private Sub TimerX_Timer()

TimerX.Interval = 0
TimerX.Enabled = False

On Error GoTo vError

strSQL = "select COD_INSTITUCION as 'IdX', Descripcion as 'ItmX' from Instituciones" _
       & " order by COD_INSTITUCION asc"
Call sbCbo_Llena_New(cboInstitucion, strSQL, True, True)


strSQL = "SELECT COD_ESTADO AS 'IdX', DESCRIPCION as 'ItmX'" _
       & "  From AFI_ESTADOS_PERSONA  Where ACTIVO = 1" _
       & "  ORDER BY DESCRIPCION"
Call sbCbo_Llena_New(cboEstadoPersona, strSQL, True, True)


Exit Sub

vError:
 

End Sub
