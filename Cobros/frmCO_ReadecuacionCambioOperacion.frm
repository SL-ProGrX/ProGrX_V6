VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Begin VB.Form frmCO_ReadecuacionCambioOperacion 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Readecuación con cambio de Operación"
   ClientHeight    =   6912
   ClientLeft      =   48
   ClientTop       =   216
   ClientWidth     =   9060
   HelpContextID   =   4002
   Icon            =   "frmCO_ReadecuacionCambioOperacion.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6912
   ScaleWidth      =   9060
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtNotas 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2880
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   39
      Top             =   1920
      Width           =   6015
   End
   Begin VB.Frame Frame1 
      Caption         =   "Operación Actual"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   3855
      Left            =   1200
      TabIndex        =   9
      Top             =   2640
      Width           =   3855
      Begin VB.TextBox txtPolizas 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1680
         TabIndex        =   35
         Top             =   2400
         Width           =   1935
      End
      Begin VB.TextBox txtCargos 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1680
         TabIndex        =   33
         Top             =   2040
         Width           =   1935
      End
      Begin VB.TextBox txtIntMoratorio 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1680
         TabIndex        =   31
         Top             =   1680
         Width           =   1935
      End
      Begin VB.TextBox txtIntCorVenc 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1680
         TabIndex        =   29
         Top             =   960
         Width           =   1935
      End
      Begin VB.TextBox txtTotalDeuda 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1680
         TabIndex        =   24
         Top             =   3360
         Width           =   1935
      End
      Begin VB.TextBox txtSaldo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1680
         TabIndex        =   23
         Top             =   2760
         Width           =   1935
      End
      Begin VB.TextBox txtIntCorAtrasado 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1680
         TabIndex        =   22
         Top             =   1320
         Width           =   1935
      End
      Begin VB.TextBox txtMonto 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1680
         TabIndex        =   21
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label Label7 
         Caption         =   "Pólizas"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   7
         Left            =   120
         TabIndex        =   36
         Top             =   2400
         Width           =   1095
      End
      Begin VB.Label Label7 
         Caption         =   "Cargos"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   34
         Top             =   2040
         Width           =   1095
      End
      Begin VB.Label Label7 
         Caption         =   "Int. Moratorios"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   32
         Top             =   1680
         Width           =   1095
      End
      Begin VB.Label Label7 
         Caption         =   "Int. Cor. Venc."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   30
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label7 
         Caption         =   "Monto Inicial"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   13
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label7 
         Caption         =   "Saldo"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   12
         Top             =   2760
         Width           =   975
      End
      Begin VB.Label Label7 
         Caption         =   "Int. Cor. Atra."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   2
         Left            =   120
         TabIndex        =   11
         Top             =   1320
         Width           =   1332
      End
      Begin VB.Label Label7 
         Caption         =   "Total Adeudado"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   10
         Top             =   3360
         Width           =   1455
      End
   End
   Begin VB.Frame fraDatos 
      Caption         =   "Nueva Operación"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   3855
      Left            =   5160
      TabIndex        =   1
      Top             =   2640
      Width           =   3735
      Begin VB.CheckBox chkDiaPago 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "Conservar Día de Pago en la nueva Operación"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   600
         TabIndex        =   40
         Top             =   1920
         Value           =   1  'Checked
         Width           =   2895
      End
      Begin VB.TextBox txtNO_Cuota 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   28
         Top             =   1440
         Width           =   1935
      End
      Begin VB.TextBox txtNO_Tasa 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   27
         Top             =   1080
         Width           =   735
      End
      Begin VB.TextBox txtNO_Plazo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2760
         TabIndex        =   26
         Top             =   720
         Width           =   735
      End
      Begin VB.TextBox txtNO_Monto 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   25
         Top             =   360
         Width           =   1935
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Command1"
         Height          =   615
         Left            =   6840
         TabIndex        =   3
         Top             =   720
         Width           =   855
      End
      Begin VB.CommandButton cmdCambiaOperacion 
         Caption         =   "&Aplicar"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   2400
         Picture         =   "frmCO_ReadecuacionCambioOperacion.frx":000C
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Aplica movimientos"
         Top             =   2880
         Width           =   1095
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   1320
         Picture         =   "frmCO_ReadecuacionCambioOperacion.frx":07D4
         Style           =   1  'Graphical
         TabIndex        =   4
         Tag             =   "1"
         ToolTipText     =   "Cancela"
         Top             =   2880
         Width           =   1095
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         X1              =   3600
         X2              =   120
         Y1              =   2760
         Y2              =   2760
      End
      Begin VB.Label Label8 
         Caption         =   "Monto"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   8
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label8 
         Caption         =   "Plazo"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   7
         Top             =   720
         Width           =   615
      End
      Begin VB.Label lblTasa 
         Caption         =   "Tasa %"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   1080
         Width           =   2415
      End
      Begin VB.Label Label8 
         Caption         =   "Cuota"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   5
         Top             =   1440
         Width           =   615
      End
   End
   Begin VB.TextBox txtOperacion 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   316
      Left            =   2880
      TabIndex        =   0
      ToolTipText     =   "Número de Operación"
      Top             =   240
      Width           =   1695
   End
   Begin MSComctlLib.StatusBar StatusBarX 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   37
      Top             =   6660
      Width           =   9060
      _ExtentX        =   15981
      _ExtentY        =   445
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   7832
            MinWidth        =   7832
            Object.ToolTipText     =   "Oficina"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4304
            MinWidth        =   4304
            Object.ToolTipText     =   "Linea"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4304
            MinWidth        =   4304
            Object.ToolTipText     =   "Recurso"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label3 
      Caption         =   "Notas de la Readecuación"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   2
      Left            =   1200
      TabIndex        =   38
      Top             =   1920
      Width           =   1575
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Operación"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1680
      TabIndex        =   20
      Top             =   240
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "Línea"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   19
      Top             =   1080
      Width           =   615
   End
   Begin VB.Label Label3 
      Caption         =   "Cédula"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   18
      Top             =   1440
      Width           =   615
   End
   Begin VB.Label lblCodigo 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   1200
      TabIndex        =   17
      ToolTipText     =   "Código del Préstamo"
      Top             =   1080
      Width           =   1695
   End
   Begin VB.Label lblCedula 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   1200
      TabIndex        =   16
      ToolTipText     =   "Cédula de la Persona"
      Top             =   1440
      Width           =   1695
   End
   Begin VB.Label lblDescripcion 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   2880
      TabIndex        =   15
      ToolTipText     =   "Descripción del Código"
      Top             =   1080
      Width           =   6015
   End
   Begin VB.Label lblNombre 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   2880
      TabIndex        =   14
      ToolTipText     =   "Nombre de la Persona"
      Top             =   1440
      Width           =   6015
   End
   Begin VB.Image imgReporte 
      Height          =   360
      Left            =   4680
      Picture         =   "frmCO_ReadecuacionCambioOperacion.frx":1151
      Stretch         =   -1  'True
      ToolTipText     =   "Reporte del Cambio con Readecuación"
      Top             =   240
      Width           =   360
   End
   Begin VB.Image imgBanner 
      Height          =   975
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   9135
   End
End
Attribute VB_Name = "frmCO_ReadecuacionCambioOperacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub sbLimpia()

 fraDatos.Enabled = False
 
 txtOperacion = ""
 
 lblCedula.Caption = ""
 lblCodigo.Caption = ""
 
 lblDescripcion.Caption = ""
 lblNombre.Caption = ""
 
 txtMonto.Text = 0
 txtIntCorAtrasado.Text = 0
 txtIntCorVenc.Text = 0
 txtIntMoratorio.Text = 0
 txtSaldo.Text = 0
 txtCargos.Text = 0
 txtPolizas.Text = 0
 txtTotalDeuda.Text = 0
 
 txtNotas.Text = ""
 
 txtNO_Monto = ""
 txtNO_Plazo = ""
 txtNO_Tasa = ""
 txtNO_Cuota = ""
 
 lblTasa.Caption = "Tasa %"

StatusBarX.Panels.Item(1).Text = ""
StatusBarX.Panels.Item(2).Text = ""
StatusBarX.Panels.Item(3).Text = ""


End Sub


Private Sub sbConsultar()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

strSQL = "select R.*,S.nombre,C.descripcion, isnull(V.intC,0) as 'IntMorCor' ,isnull(V.intM,0) as 'IntMorMor', isnull(V.cargos,0) as 'Cargos'" _
       & ",isnull(R.liqTasa,0) as LiqTasaX,dbo.fxCRDCalculoIntCorte(R.id_solicitud,dbo.MyGetdate()) as 'InteresTotal'" _
       & ",0 as 'Poliza', O.descripcion as 'OficinaDesc', R.cod_oficina_r as 'Oficina',R.cod_grupo,Pre.descripcion as 'RecursoDesc'" _
       & ",dbo.MyGetdate() as 'FechaServer'" _
       & " from reg_creditos R inner join Socios S on R.cedula = S.cedula" _
       & " inner join catalogo C on R.codigo = C.codigo" _
       & " left join sif_oficinas O on R.cod_oficina_R = O.cod_Oficina" _
       & " left join Vista_Morosidad V on R.id_solicitud = V.id_solicitud" _
       & " left join CATALOGO_GRUPOS Pre on R.cod_grupo = Pre.cod_grupo" _
       & " Where R.id_solicitud = " & txtOperacion.Text & " and R.estado = 'A'"
    
Call OpenRecordSet(rs, strSQL)

If rs.EOF And rs.BOF Then
 MsgBox "No se encontró el número de operación [Activa]"
 Exit Sub
Else
 fraDatos.Enabled = True
 lblCedula.Caption = rs!Cedula
 lblCodigo.Caption = rs!Codigo
 lblDescripcion.Caption = rs!Descripcion
 lblNombre.Caption = rs!Nombre
 
 txtMonto.Text = Format(rs!montoapr, "Standard")
 txtIntCorAtrasado.Text = Format(rs!intMorCor, "Standard")
 txtIntMoratorio.Text = Format(rs!intMorMor, "Standard")
 txtPolizas.Text = Format(rs!Poliza, "Standard")
 txtCargos.Text = Format(rs!Cargos, "Standard")
 txtSaldo.Text = Format(rs!Saldo, "Standard")
 
 txtIntCorVenc.Text = Format(rs!InteresTotal - (rs!intMorCor + rs!intMorMor), "Standard")
 
 txtTotalDeuda.Text = Format(rs!Saldo + rs!InteresTotal + rs!Cargos + rs!Poliza, "Standard")


 txtNO_Tasa.Locked = True
 
 If Not IsNull(rs!TBP_PuntosAdd) Then
   lblTasa.Caption = "Tasa (TBP + " & rs!TBP_PuntosAdd & ")"
 Else
   lblTasa.Caption = "Tasa %"
   txtNO_Tasa.Locked = False
 End If
 
 If rs!LiqTasaX = 1 Then
   lblTasa.Caption = lblTasa.Caption & " + PtsLiq"
 End If
 
 txtNO_Monto = txtTotalDeuda.Text
 txtNO_Plazo = rs!Plazo
 txtNO_Tasa = rs!interesv
 
 txtNO_Cuota = fxCalcula_Cuota(CCur(txtNO_Monto), CCur(txtNO_Plazo), CCur(txtNO_Tasa))
 txtOperacion.Enabled = False

 StatusBarX.Panels.Item(1).Text = rs!OficinaDesc & ""
 StatusBarX.Panels.Item(1).Tag = rs!Oficina
 StatusBarX.Panels.Item(2).Text = rs!Descripcion & ""
 StatusBarX.Panels.Item(3).Text = rs!RecursoDesc & ""

 If GLOBALES.SysPlanPagos = 1 Then
       strSQL = "exec spCrdPlanPagosInfoCancelacion " & txtOperacion.Text & ", '" & Format(rs!FechaServer, "yyyy/mm/dd") & "'"
       rs.Close
       Call OpenRecordSet(rs, strSQL)
        
        txtIntCorVenc.Text = "0.00"
        txtIntMoratorio.Text = Format(rs!IntMor, "Standard")
        txtIntCorAtrasado.Text = Format(rs!IntCor, "Standard")
        txtCargos.Text = Format(rs!Cargos, "Standard")
        txtPolizas.Text = Format(rs!Poliza, "Standard")
        
        txtTotalDeuda.Text = Format(rs!Principal + rs!IntCor + rs!IntMor + rs!Cargos + rs!Poliza, "Standard")
        txtNO_Monto = txtTotalDeuda.Text
        
        txtNO_Cuota = fxCalcula_Cuota(CCur(txtNO_Monto), CCur(txtNO_Plazo), CCur(txtNO_Tasa))
 End If 'GLOBALES.SysPlanPagos = 1
 
End If

rs.Close

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Function fxVerifica() As Boolean
Dim vMensaje As String

vMensaje = ""

If txtSaldo = "" Or Not IsNumeric(txtSaldo.Text) Then vMensaje = vMensaje & " - Saldo no es válido..." & vbCrLf
 
If txtNO_Monto = "" Or Not IsNumeric(txtNO_Monto.Text) Then vMensaje = vMensaje & " - Monto de la nueva Operación no es válido..." & vbCrLf
If txtNO_Tasa = "" Or Not IsNumeric(txtNO_Tasa.Text) Then vMensaje = vMensaje & " - Tasa no es válida..." & vbCrLf
If txtNO_Plazo = "" Or Not IsNumeric(txtNO_Plazo.Text) Then vMensaje = vMensaje & " - Plazo no es válido..." & vbCrLf
If txtNO_Cuota = "" Or Not IsNumeric(txtNO_Cuota.Text) Then vMensaje = vMensaje & " - Cuota no es válida..." & vbCrLf


If Len(txtNotas.Text) < 10 Then vMensaje = vMensaje & " - La nota para realizar la transacción no es válida..." & vbCrLf

If Len(vMensaje) = 0 Then
  fxVerifica = True
Else
  fxVerifica = False
  MsgBox vMensaje, vbExclamation
End If

End Function


Private Sub sbDocumento(pTipoDoc As String, pNumDoc As String, pConcepto As String, pTipoMov As String)
Dim rs As New ADODB.Recordset, strSQL As String, strLinea(10) As String
Dim curIntC As Currency, curIntM As Currency, curCargo As Currency
Dim curMonto As Currency, curAmortiza As Currency, curPoliza As Currency
Dim vFecha As Date, vCuentaPoliza As String, rsTmp As New ADODB.Recordset



vAseDocDetalle = txtNotas.Text
vFecha = fxFechaServidor

curCargo = 0
curIntC = 0
curIntM = 0
curAmortiza = 0
curPoliza = 0

'Detecta Movimientos Aplicados ( La vista trabaja con y Sin Tabla de Pagos)
If GLOBALES.SysPlanPagos = 1 Then
    strSQL = "exec spCrdDocumentoAfectacion '" & pTipoDoc & "','" & pNumDoc & "','R'"

Else
    strSQL = "select isnull(SUM(intCor),0) as 'IntCor', isnull(SUM(intMor),0) as 'IntMor', isnull(SUM(Cargo),0) as 'Cargos'" _
           & ",isnull(SUM(Poliza),0) as 'Polizas', isnull(SUM(Principal),0) as 'Principal'" _
           & " From dbo.vCRDsReportesMov" _
           & " where Tcon = '" & pTipoMov & "' and Ncon = '" & pNumDoc _
           & "' and id_solicitud = " & txtOperacion.Text
End If

Call OpenRecordSet(rs, strSQL)
If Not rs.EOF And Not rs.BOF Then
    curCargo = rs!Cargos
    curIntC = rs!IntCor
    curIntM = rs!IntMor
    curAmortiza = rs!Principal
    curPoliza = rs!Polizas
End If
rs.Close
curMonto = curCargo + curIntC + curIntM + curAmortiza + curPoliza


'Cuentas
strSQL = "exec spCrdOperacionCtas " & txtOperacion.Text
Call OpenRecordSet(rs, strSQL)

strLinea(1) = "Saldo Anterior    " & txtSaldo.Text
strLinea(2) = "Interes Corriente " & Format(curIntC, "Standard")
strLinea(3) = "Interes Moratorio " & Format(curIntM, "Standard")
strLinea(4) = "Amortización      " & Format(curAmortiza, "Standard")

strLinea(5) = "Saldo Actual      " & Format(CCur(txtSaldo.Text) - curAmortiza, "Standard")
strLinea(6) = "Pólizas           " & Format(curPoliza, "Standard")
strLinea(7) = "Cargos [General]  " & Format(curCargo, "Standard")

strLinea(8) = "Operacion/Línea   " & "Op.:" & txtOperacion.Text & " L.:" & lblCodigo.Caption
strLinea(9) = "Descripción       " & lblDescripcion.Caption
strLinea(10) = " "

If GLOBALES.SysDocVersion = 1 Then
        
        strSQL = "insert asientos_tmp(TMP_TIPO,TMP_USUARIO,TMP_CASO,TMP_CUENTA,TMP_MONTO," _
            & "TMP_DEBEHABER,TMP_FECHA,TMP_ESTADO_ASIENTO) values('TRA','" & glogon.Usuario _
            & "','RA" & txtOperacion & "','" & rs!ctaamortiza & "'," & curMonto & ",'D','" _
            & Format(vFecha, "yyyy/mm/dd") & "','P')"
        Call ConectionExecute(strSQL)
        
        strSQL = "insert asientos_tmp(TMP_TIPO,TMP_USUARIO,TMP_CASO,TMP_CUENTA,TMP_MONTO," _
            & "TMP_DEBEHABER,TMP_FECHA,TMP_ESTADO_ASIENTO) values('TRA','" & glogon.Usuario _
            & "','RA" & txtOperacion & "','" & rs!ctaamortiza & "'," & curAmortiza & ",'H','" _
            & Format(vFecha, "yyyy/mm/dd") & "','P')"
        Call ConectionExecute(strSQL)
        
        If curCargo > 0 Then
            strSQL = "insert asientos_tmp(TMP_TIPO,TMP_USUARIO,TMP_CASO,TMP_CUENTA,TMP_MONTO," _
              & "TMP_DEBEHABER,TMP_FECHA,TMP_ESTADO_ASIENTO) values('TRA','" & glogon.Usuario _
              & "','RA" & txtOperacion & "','" & rs!CtaCargos & "'," & curCargo & ",'H','" _
              & Format(vFecha, "yyyy/mm/dd") & "','P')"
            Call ConectionExecute(strSQL)
        End If
        
        
        If curIntC > 0 Then
            strSQL = "insert asientos_tmp(TMP_TIPO,TMP_USUARIO,TMP_CASO,TMP_CUENTA,TMP_MONTO," _
                & "TMP_DEBEHABER,TMP_FECHA,TMP_ESTADO_ASIENTO) values('TRA','" & glogon.Usuario _
                & "','RA" & txtOperacion & "','" & rs!ctaintc & "'," & curIntC & ",'H','" _
                & Format(vFecha, "yyyy/mm/dd") & "','P')"
            Call ConectionExecute(strSQL)
        End If
        
        
        If curIntM > 0 Then
            strSQL = "insert asientos_tmp(TMP_TIPO,TMP_USUARIO,TMP_CASO,TMP_CUENTA,TMP_MONTO," _
                & "TMP_DEBEHABER,TMP_FECHA,TMP_ESTADO_ASIENTO) values('TRA','" & glogon.Usuario _
                & "','RA" & txtOperacion & "','" & rs!ctaintm & "'," & curIntM & ",'H','" _
                & Format(vFecha, "yyyy/mm/dd") & "','P')"
            Call ConectionExecute(strSQL)
        End If
        
        
        
Else
       'Control de Documentos v2
        strSQL = "insert SIF_TRANSACCIONES(COD_TRANSACCION,TIPO_DOCUMENTO,REGISTRO_FECHA,REGISTRO_USUARIO,Cliente_IDENTIFICACION,CLIENTE_NOMBRE" _
                & ",cod_concepto,monto,estado,Referencia_01,Referencia_02,Referencia_03,cod_oficina" _
                & ",linea1,linea2,linea3,linea4,linea5,linea6,linea7,linea8,linea9,linea10,detalle,documento)" _
                & " values('" & pNumDoc & "','" & pTipoDoc & "',dbo.MyGetdate(),'" & glogon.Usuario & "','" & Trim(lblCedula.Caption) _
                & "','" & Trim(lblNombre.Caption) & "','" & pConcepto & "'," & curMonto & ",'P','" & txtOperacion.Text _
                & "','" & lblCodigo.Caption & "','" & vAseDocDeposito & "','" & GLOBALES.gOficinaTitular & "','" & strLinea(1) & "','" _
                & strLinea(2) & "','" & strLinea(3) & "','" & strLinea(4) & "','" _
                & strLinea(5) & "','" & strLinea(6) & "','" & strLinea(7) & "','" _
                & strLinea(8) & "','" & strLinea(9) & "','" & strLinea(10) & "','" _
                & vAseDocDetalle & "','" & vAseDocDeposito & "')"
'        Call ConectionExecute(strSQL)
        
        'ASIENTO
        If curMonto > 0 Then
          strSQL = strSQL & Space(10) & "exec spSIFDocsAsiento '" & pTipoDoc & "','" & pNumDoc & "'," & curMonto & ",'D','" & rs!Cod_Divisa _
                 & "',1," & GLOBALES.gEnlace & ",'" & rs!cod_unidad & "','" & rs!cod_centro_costo & "','" & rs!ctaamortiza _
                 & "','" & rs!id_solicitud & "','" & rs!Codigo & "','" & vAseDocDeposito & "'"
        End If
        
        If curAmortiza > 0 Then
          strSQL = strSQL & Space(10) & "exec spSIFDocsAsiento '" & pTipoDoc & "','" & pNumDoc & "'," & curAmortiza & ",'C','" & rs!Cod_Divisa _
                 & "',1," & GLOBALES.gEnlace & ",'" & rs!cod_unidad & "','" & rs!cod_centro_costo & "','" & rs!ctaamortiza _
                 & "','" & rs!id_solicitud & "','" & rs!Codigo & "','" & vAseDocDeposito & "'"
        End If
        
        If curIntC > 0 Then
          strSQL = strSQL & Space(10) & "exec spSIFDocsAsiento '" & pTipoDoc & "','" & pNumDoc & "'," & curIntC & ",'C','" & rs!Cod_Divisa _
                 & "',1," & GLOBALES.gEnlace & ",'" & rs!cod_unidad & "','" & rs!cod_centro_costo & "','" & rs!ctaintc _
                 & "','" & rs!id_solicitud & "','" & rs!Codigo & "','" & vAseDocDeposito & "'"
        End If
        
        If curIntM > 0 Then
          strSQL = strSQL & Space(10) & "exec spSIFDocsAsiento '" & pTipoDoc & "','" & pNumDoc & "'," & curIntM & ",'C','" & rs!Cod_Divisa _
                 & "',1," & GLOBALES.gEnlace & ",'" & rs!cod_unidad & "','" & rs!cod_centro_costo & "','" & rs!ctaintm _
                 & "','" & rs!id_solicitud & "','" & rs!Codigo & "','" & vAseDocDeposito & "'"
        End If
        
        If curCargo > 0 And GLOBALES.SysPlanPagos = 0 Then
          strSQL = strSQL & Space(10) & "exec spSIFDocsAsiento '" & pTipoDoc & "','" & pNumDoc & "'," & curCargo & ",'C','" & rs!Cod_Divisa _
                 & "',1," & GLOBALES.gEnlace & ",'" & rs!cod_unidad & "','" & rs!cod_centro_costo & "','" & rs!CtaCargos _
                 & "','" & rs!id_solicitud & "','" & rs!Codigo & "','" & vAseDocDeposito & "'"
        End If

        If curCargo > 0 And GLOBALES.SysPlanPagos = 1 Then
        'Detallar Cargos
          glogon.strSQL = "exec spCrdDocumentoAfectacionCargos '" & pTipoDoc & "','" & pNumDoc & "'"
          Call OpenRecordSet(rsTmp, glogon.strSQL, 0)
          Do While Not rsTmp.EOF
                strSQL = strSQL & Space(10) & "exec spSIFDocsAsiento '" & pTipoDoc & "','" & pNumDoc & "'," & rsTmp!Mov_Monto & ",'C','" & rs!Cod_Divisa _
                       & "',1," & GLOBALES.gEnlace & ",'" & rsTmp!cod_unidad & "','" & rsTmp!cod_centro_costo & "','" & rsTmp!cod_cuenta _
                       & "','" & rsTmp!id_solicitud & "','" & rsTmp!Codigo & "','" & vAseDocDeposito & "'"
'                Call ConectionExecute(strSQL)
                rsTmp.MoveNext
          Loop
          rsTmp.Close
        End If
        
        If curPoliza > 0 And GLOBALES.SysPlanPagos = 1 Then
          glogon.strSQL = "select dbo.fxCrdOperacionCtaContaPolizas(" & rs!id_solicitud & ") as 'Cuenta'"
          Call OpenRecordSet(rsTmp, glogon.strSQL, 0)
            vCuentaPoliza = Trim(rsTmp!Cuenta)
          rsTmp.Close
          
          strSQL = strSQL & Space(10) & "exec spSIFDocsAsiento '" & pTipoDoc & "','" & pNumDoc & "'," & curPoliza & ",'C','" & rs!Cod_Divisa _
                 & "',1," & GLOBALES.gEnlace & ",'" & rs!cod_unidad & "','" & rs!cod_centro_costo & "','" & vCuentaPoliza _
                 & "','" & rs!id_solicitud & "','" & rs!Codigo & "','" & vAseDocDeposito & "'"
'          Call ConectionExecute(strSQL)
        End If


End If 'GLOBALES.SysDocVersion  = 1
rs.Close

'Registra el Documento  + Asiento
 Call ConectionExecute(strSQL)



End Sub

Private Sub sbReadecuacion()
Dim strSQL As String, rs As New ADODB.Recordset
Dim vTransac As Boolean, vOpex As Integer
Dim lngUltimaOperacion As Long, vFecha As Date

Dim vTipoDoc As String, vTipoMov As String, vNumDoc As String
Dim vConcepto As String, vCuenta As String, vDiaPago As Integer, vDeducPlanilla As String

On Error GoTo vError

Me.MousePointer = vbHourglass

vTransac = False
vFecha = fxFechaServidor

'Configuracion del Documento
If GLOBALES.SysDocVersion = 2 Then
    vTipoDoc = "REA"
    vTipoMov = "REA"
    vConcepto = "CBR001"
    vNumDoc = fxDocumentoConsecutivo(vTipoDoc)
Else
    vTipoDoc = "REA"
    vTipoMov = "4"
    vNumDoc = "8889"
    vConcepto = "CBR001"
End If



glogon.Conection.BeginTrans
vTransac = True
If GLOBALES.SysPlanPagos = 1 Then
        strSQL = "exec spCrdPlanPagoAbonoCancelacion " & txtOperacion.Text & ",'" & vConcepto & "','" & glogon.Usuario & "','" & vTipoDoc _
               & "','" & vNumDoc & "'," & CCur(txtNO_Monto.Text) & ",'" & Format(vFecha, "yyyy/mm/dd hh:mm:ss") & "',''"
        Call ConectionExecute(strSQL)
Else
    'Cancelar morosidad sin abonos
     strSQL = "update morosidad set estado = 'C'" _
            & ",abintc = intc,abintm = intm,abamortiza = amortiza,abCargo = cargo" _
            & ",tcon = '" & vTipoMov & "',ncon='" & vNumDoc & "'" _
            & ",fecult = dbo.MyGetdate(), usuario = '" & glogon.Usuario & "',cod_concepto = '" & vConcepto & "', cod_caja = ''" _
            & "where estado = 'A' and id_solicitud = " & txtOperacion
     Call ConectionExecute(strSQL)
    
    strSQL = "select isnull(sum(abamortiza),0) as Amortiza from morosidad where tcon = '" & vTipoMov & "'" _
            & " and ncon = '" & vNumDoc & "' and id_solicitud = " & txtOperacion
    rs.CursorLocation = adUseServer
    Call OpenRecordSet(rs, strSQL)
            'Insertar Registro de Detalle
            strSQL = "insert creditos_dt(CODIGO,ID_SOLICITUD,CUOTA,ABONO,INTCP,AMORTIZA,FECHAS," _
                   & "FECHAP,TCON,NCON,ESTADO,ESTADO_ASIENTO,Usuario,Cod_Concepto,Cod_Caja) values('" & lblCodigo.Caption & "'," _
                   & txtOperacion & "," & CCur(txtIntCorVenc.Text) + CCur(txtSaldo.Text) - rs!Amortiza _
                   & "," & CCur(txtSaldo.Text) - rs!Amortiza & "," & CCur(txtIntCorVenc.Text) & "," _
                   & CCur(txtSaldo.Text) - rs!Amortiza & ",dbo.MyGetdate()," & GLOBALES.glngFechaCR & ",'" _
                   & vTipoMov & "','" & vNumDoc & "','A','G','" & glogon.Usuario & "','" & vConcepto & "','')"
            Call ConectionExecute(strSQL)
    rs.Close
    
    'Cancela la operacion actual
     strSQL = "update reg_creditos set saldo = 0, amortiza = montoapr,saldo_mes = 0," _
            & "estado = 'C',FECHA_ENVIAPROCESO = dbo.MyGetdate(),OBSERVACION_PROCESO='Readecuación de Deuda'" _
            & " where id_solicitud = " & txtOperacion
     Call ConectionExecute(strSQL)
End If
glogon.Conection.CommitTrans
vTransac = False



'Abrir nueva operacion
strSQL = "select * from reg_creditos where id_solicitud = " & txtOperacion
Call OpenRecordSet(rs, strSQL)
 
 vOpex = IIf(IsNull(rs!Opex), 0, rs!Opex)
 
 If chkDiaPago.Value = vbChecked Then
   vDiaPago = rs!dia_pago
 Else
   vDiaPago = Day(vFecha)
 End If
 
 If rs!ind_deduce_planilla = "S" And vDiaPago < 32 Then
    vDiaPago = 32
 End If
 
 strSQL = "insert reg_creditos(codigo,id_comite,cedula,montosol,estadosol,fechasol,fechares,plazo,int,montoapr,prideduc,fechaforp,fechaforf,acta,saldo,amortiza,interesc" _
        & ",cuota,estado,opex,proceso,userrec,userres,userfor,garantia,observacion,firma_deudor,monto_girado,interesv,tesoreria,usertesoreria,primer_cuota" _
        & ",tdocumento,ndocumento,pagare,fecha_calculo_int,premio,cuotas_planilla,cuotas_directas,cuotas_anuladas,FECULT,TBP_PuntosAdd" _
        & ",LiqTasa,cod_oficina_r,cod_oficina_f,cod_oficina_comision,referencia,fecha_registro,DIA_PAGO, IND_DEDUCE_PLANILLA) values(" _
        & "'" & rs!Codigo & "'," & rs!id_Comite & ",'" & rs!Cedula & "'," & CCur(txtNO_Monto) & ",'F','" & Format(vFecha, "yyyy/mm/dd") & "','" & Format(vFecha, "yyyy/mm/dd") & "'," _
        & txtNO_Plazo & "," & txtNO_Tasa & "," & CCur(txtNO_Monto) & "," & fxPrimerDeduccion & ",'" & Format(vFecha, "yyyy/mm/dd") & "','" & Format(vFecha, "yyyy/mm/dd") & "'," _
        & IIf(IsNull(rs!acta), 0, rs!acta) & "," & CCur(txtNO_Monto) & ",0,0," _
        & CCur(txtNO_Cuota) & ",'A'," & rs!Opex & ",'N','" & glogon.Usuario & "','" & glogon.Usuario & "','" & glogon.Usuario & "','" & rs!Garantia & "','" & txtNotas.Text & "'," _
        & "1,0," & txtNO_Tasa & ",'" & Format(vFecha, "yyyy/mm/dd") & "','" & glogon.Usuario & "','N','ND'" _
        & ",'" & rs!id_solicitud & "'," & IIf(IsNull(rs!pagare), 0, rs!pagare) & ",'" & Format(vFecha, "yyyy/mm/dd") & "'," & IIf(IsNull(rs!Premio), 0, rs!Premio) & "," _
        & "0,0,0," & rs!FecUlt & "," & IIf(IsNull(rs!TBP_PuntosAdd), "null", rs!TBP_PuntosAdd) & "," & IIf(IsNull(rs!LiqTasa), "null", rs!LiqTasa) & "," _
        & IIf(IsNull(rs!cod_oficina_r), "Null", "'" & Trim(rs!cod_oficina_r) & "'") & "," & IIf(IsNull(rs!cod_oficina_f), "Null", "'" & Trim(rs!cod_oficina_f) & "'") & "," _
        & IIf(IsNull(rs!cod_oficina_comision), "Null", "'" & Trim(rs!cod_oficina_comision) & "'") & "," & txtOperacion.Text & ",dbo.MyGetdate()," _
        & vDiaPago & ",'" & rs!ind_deduce_planilla & "')"
rs.Close

glogon.Conection.BeginTrans
vTransac = True

Call ConectionExecute(strSQL)
 
'Recuperar la nueva operacion
lngUltimaOperacion = fxUltimaOperacion(lblCedula.Caption)

'Hereda Fiadores Operacion Anterior
strSQL = "insert into fiadores(id_solicitud,codigo,cedulaf,nombre,firma,estado,interno) " _
       & " select " & lngUltimaOperacion & ",codigo,cedulaf,nombre,firma,estado,interno " _
       & " from fiadores where id_solicitud = " & txtOperacion.Text
Call ConectionExecute(strSQL)

glogon.Conection.CommitTrans
vTransac = False

'Crea Plan de Pagos
If GLOBALES.SysPlanPagos = 1 Then
    strSQL = "exec spCrdPlanPagos " & lngUltimaOperacion
    strSQL = strSQL & Space(10) & "exec spCrdPlanPagosActivaCuota " & lngUltimaOperacion
    Call ConectionExecute(strSQL)
End If


'Confecciona el documento
Call sbDocumento(vTipoDoc, vNumDoc, vConcepto, vTipoMov)

'Bitacora General
Call Bitacora("Aplica", ("Readecuacion de Operacion de " & txtOperacion & " A " & lngUltimaOperacion))


'Registro Historial y Expediente
Call sbCBRRegTransac("03", lblCedula.Caption, txtOperacion, txtNotas.Text, CCur(txtSaldo.Text), CCur(txtIntCorAtrasado.Text) + CCur(txtIntCorVenc.Text), CCur(txtIntMoratorio.Text), CCur(txtCargos.Text), CCur(txtPolizas.Text), CCur(txtSaldo.Text), vTipoDoc, vNumDoc)
  
'Imprime Comprobante
If GLOBALES.SysDocVersion = 2 Then
        Call sbImprimeRecibo(vNumDoc, vTipoDoc)
End If

Me.MousePointer = vbDefault

MsgBox ("- La operación No. " & txtOperacion _
        & " fue cancelada y se registró nueva operación No. " _
        & lngUltimaOperacion & vbCrLf & vbCrLf & " - Readecuación No." & vNumDoc), vbInformation
 
Exit Sub

vError:
 If vTransac Then glogon.Conection.RollbackTrans
 
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
 
End Sub

Private Sub cmdCambiaOperacion_Click()
Dim iRespuesta As Integer

  'Verificar Congelamiento
  If fxgCongelamiento(lblCedula.Caption, "per_readecuaciones") Then
    MsgBox "Esta Persona se encuentra CONGELADA, verifique...", vbExclamation
    Exit Sub
  End If


If fxVerifica Then
 ' YES = 6 , NO = 7
 iRespuesta = MsgBox("Esta seguro que desea aplicar la readecuación de esta Operación?", vbYesNo)
  
 Call sbSIFCleanTxtInject(txtNotas)
 
 If iRespuesta = vbYes Then
    Call sbReadecuacion
 End If
  'Reestablece ventana
 Call sbLimpia
 txtOperacion.Enabled = True
 txtOperacion.SetFocus
Else
  MsgBox "Revise la información suministrada no es válida", vbCritical
End If

End Sub

Private Sub cmdCancelar_Click()
  Call sbLimpia
  txtOperacion.Enabled = True
  txtOperacion.SetFocus
End Sub


Private Sub Form_Activate()
 vModulo = 4
End Sub

Private Sub Form_Load()
  
  vModulo = 4
  
  Set imgBanner.Picture = frmContenedor.imgBanner_01.Picture
  
  Call Formularios(Me)
  Call RefrescaTags(Me)
  
  Call sbLimpia

End Sub

Private Sub imgReporte_Click()
Dim rs As New ADODB.Recordset, strSQL As String, vOP As Long

If txtOperacion = "" Then Exit Sub
Me.MousePointer = vbHourglass
strSQL = "select min(id_solicitud) as Operacion from reg_creditos where cedula = '" _
        & lblCedula.Caption & "' and codigo = '" & lblCodigo.Caption & "'" _
        & " and estadosol = 'F' and id_solicitud > " & txtOperacion
rs.CursorLocation = adUseServer
Call OpenRecordSet(rs, strSQL)
vOP = IIf(IsNull(rs!Operacion), 0, rs!Operacion)
rs.Close

Me.MousePointer = vbDefault

If vOP = 0 Then Exit Sub

Me.MousePointer = vbHourglass

With frmContenedor.Crt
 .Reset
 .WindowShowGroupTree = True
 .WindowShowPrintSetupBtn = True
 .WindowShowRefreshBtn = True
 .WindowShowSearchBtn = True
 .WindowState = crptMaximized
 .WindowTitle = "Reportes del Módulo de Cobro"
 
 .Connect = glogon.ConectRPT
 
 .ReportFileName = SIFGlobal.fxPathReportes("Cobro_Readecuacion.rpt")
 .Formulas(0) = "empresa='" & GLOBALES.gstrNombreEmpresa & "'"
 .Formulas(1) = "fecha='" & Format(fxFechaServidor, "dd/mm/yyyy") & "'"
 .Formulas(2) = "usuario='" & glogon.Usuario & "'"
 .SelectionFormula = "{REG_CREDITOS.ID_SOLICITUD} = " & vOP
 
 .SubreportToChange = "Original"
 .SelectionFormula = "{REG_CREDITOS.ID_SOLICITUD} = " & txtOperacion
 
 .PrintReport
End With

Me.MousePointer = vbDefault



End Sub





Private Sub txtNO_Tasa_Change()
On Error GoTo vError

If CCur(IIf((txtNO_Tasa = ""), 0, txtNO_Tasa)) > 0 And CCur(IIf((txtNO_Plazo = ""), 0, txtNO_Plazo)) > 0 _
    And CCur(IIf((txtNO_Monto = ""), 0, txtNO_Monto)) > 0 Then
 txtNO_Cuota = fxCalcula_Cuota(CCur(txtNO_Monto), CCur(txtNO_Plazo), CCur(txtNO_Tasa))
End If

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub txtNO_Monto_Change()
On Error GoTo vError

If CCur(IIf((txtNO_Tasa = ""), 0, txtNO_Tasa)) > 0 And CCur(IIf((txtNO_Plazo = ""), 0, txtNO_Plazo)) > 0 _
    And CCur(IIf((txtNO_Monto = ""), 0, txtNO_Monto)) > 0 Then
 txtNO_Cuota = fxCalcula_Cuota(CCur(txtNO_Monto), CCur(txtNO_Plazo), CCur(txtNO_Tasa))
End If

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub txtNO_Plazo_Change()
On Error GoTo vError

If CCur(IIf((txtNO_Tasa = ""), 0, txtNO_Tasa)) > 0 And CCur(IIf((txtNO_Plazo = ""), 0, txtNO_Plazo)) > 0 _
    And CCur(IIf((txtNO_Monto = ""), 0, txtNO_Monto)) > 0 Then
 txtNO_Cuota = fxCalcula_Cuota(CCur(txtNO_Monto), CCur(txtNO_Plazo), CCur(txtNO_Tasa))
End If

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub txtOperacion_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = vbKeyReturn Then Call sbConsultar
End Sub
