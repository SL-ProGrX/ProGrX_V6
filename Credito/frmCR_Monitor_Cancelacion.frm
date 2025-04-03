VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#20.0#0"; "Codejock.Controls.v20.0.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#20.0#0"; "Codejock.ShortcutBar.v20.0.0.ocx"
Begin VB.Form frmCR_Monitor_Cancelacion 
   Caption         =   "Monitor de Cancelación"
   ClientHeight    =   9684
   ClientLeft      =   48
   ClientTop       =   396
   ClientWidth     =   13452
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9684
   ScaleWidth      =   13452
   WindowState     =   2  'Maximized
   Begin VB.Timer TimerX 
      Interval        =   10
      Left            =   360
      Top             =   360
   End
   Begin XtremeSuiteControls.PushButton btnConsulta 
      Height          =   372
      Left            =   9720
      TabIndex        =   5
      Top             =   1320
      Width           =   1332
      _Version        =   1310720
      _ExtentX        =   2350
      _ExtentY        =   656
      _StockProps     =   79
      Caption         =   "Buscar"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   6
   End
   Begin FPSpreadADO.fpSpread vGrid 
      Height          =   6012
      Left            =   0
      TabIndex        =   0
      Top             =   1800
      Width           =   13812
      _Version        =   524288
      _ExtentX        =   24363
      _ExtentY        =   10605
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
      MaxCols         =   16
      ScrollBars      =   2
      SpreadDesigner  =   "frmCR_Monitor_Cancelacion.frx":0000
      VScrollSpecialType=   2
      AppearanceStyle =   1
   End
   Begin XtremeSuiteControls.FlatEdit txtPorcentaje 
      Height          =   312
      Left            =   8160
      TabIndex        =   6
      Top             =   1380
      Width           =   1212
      _Version        =   1310720
      _ExtentX        =   2138
      _ExtentY        =   556
      _StockProps     =   77
      ForeColor       =   0
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
      Text            =   "20"
      Alignment       =   2
      Appearance      =   2
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.DateTimePicker dtpInicio 
      Height          =   315
      Left            =   2160
      TabIndex        =   7
      Top             =   1380
      Width           =   1330
      _Version        =   1310720
      _ExtentX        =   2346
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
      Left            =   3480
      TabIndex        =   8
      Top             =   1380
      Width           =   1332
      _Version        =   1310720
      _ExtentX        =   2346
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
   Begin XtremeSuiteControls.PushButton btnExportar 
      Height          =   372
      Left            =   11040
      TabIndex        =   9
      Top             =   1320
      Width           =   1332
      _Version        =   1310720
      _ExtentX        =   2350
      _ExtentY        =   656
      _StockProps     =   79
      Caption         =   "Exportar"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   6
   End
   Begin XtremeSuiteControls.Label Label2 
      Height          =   372
      Index           =   1
      Left            =   6120
      TabIndex        =   4
      Top             =   1320
      Width           =   2172
      _Version        =   1310720
      _ExtentX        =   3831
      _ExtentY        =   656
      _StockProps     =   79
      Caption         =   "Porcentaje de Desviación"
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
   End
   Begin XtremeSuiteControls.Label Label2 
      Height          =   372
      Index           =   0
      Left            =   240
      TabIndex        =   3
      Top             =   1320
      Width           =   2172
      _Version        =   1310720
      _ExtentX        =   3831
      _ExtentY        =   656
      _StockProps     =   79
      Caption         =   "Fechas de Cancelación"
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
   End
   Begin XtremeShortcutBar.ShortcutCaption scMain 
      Height          =   612
      Left            =   0
      TabIndex        =   2
      Top             =   1200
      Width           =   13332
      _Version        =   1310720
      _ExtentX        =   23516
      _ExtentY        =   1080
      _StockProps     =   14
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SubItemCaption  =   -1  'True
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Monitor de Cancelación de Operaciones con desviación en Cuota Final"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   13.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   720
      Index           =   0
      Left            =   2160
      TabIndex        =   1
      Top             =   300
      Width           =   10812
   End
   Begin VB.Image imgBanner 
      Height          =   1212
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   14292
   End
End
Attribute VB_Name = "frmCR_Monitor_Cancelacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnConsulta_Click()

If Not IsNumeric(txtPorcentaje.Text) Then
    MsgBox "Error en el Porcentaje: Tiene que ser numérico entre 0 a 100", vbExclamation
    Exit Sub
End If

If CCur(txtPorcentaje.Text) < 0 Or CCur(txtPorcentaje.Text) > 100 Then
    MsgBox "Error en el Porcentaje: Tiene que ser numérico entre 0 a 100", vbExclamation
    Exit Sub
End If

If dtpInicio.Value > dtpCorte.Value Then
    MsgBox "Error en el Rango de Fechas!", vbExclamation
    Exit Sub
End If

Call sbConsulta

End Sub

Private Sub btnExportar_Click()
Dim vHeaders As vGridHeaders
    vHeaders.Columnas = vGrid.MaxCols
    vHeaders.Headers(1) = "Identificación"
    vHeaders.Headers(2) = "Nombre"
    vHeaders.Headers(3) = "No.Operación"
    vHeaders.Headers(4) = "Línea"
    vHeaders.Headers(5) = "Saldo"
    vHeaders.Headers(6) = "Plazo"
    vHeaders.Headers(7) = "Tasa Original"
    vHeaders.Headers(8) = "Tasa Actual"
    vHeaders.Headers(9) = "Cuota Inicial"
    vHeaders.Headers(10) = "Cuota Final"
    vHeaders.Headers(11) = "Fecha Termina"
    vHeaders.Headers(12) = "Fecha Termina Inicial"
    vHeaders.Headers(13) = "Plazo Termina"
    vHeaders.Headers(14) = "Plazo Termina Inicial"
    vHeaders.Headers(15) = "Fecha Formaliza"
    vHeaders.Headers(16) = "No. Documento"

 Call sbSIFGridExportar(vGrid, vHeaders, "ProGrX_Creditos_Cancela_Desviacion")
End Sub

Private Sub Form_Load()

Dim vFecha As Date

vModulo = 3
Set imgBanner.Picture = frmContenedor.imgBanner_Consultas.Picture

vFecha = fxFechaServidor

dtpInicio.Value = DateAdd("m", -2, vFecha)
dtpCorte.Value = DateAdd("m", 3, vFecha)

txtPorcentaje.Text = 20


End Sub


Private Sub sbConsulta()
Dim strSQL As String

On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = "exec spCrd_Monitor_Cancelacion '" & Format(dtpInicio.Value, "yyyy/mm/dd") & " 00:00:00'" _
        & ",'" & Format(dtpCorte.Value, "yyyy/mm/dd") & " 23:59:59', " & txtPorcentaje.Text

Call sbCargaGrid(vGrid, 16, strSQL)

Me.MousePointer = vbDefault
Exit Sub

vError:
    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical
    

End Sub

Private Sub Form_Resize()

On Error Resume Next

imgBanner.Width = Me.Width
scMain.Width = Me.Width

vGrid.Width = Me.Width - 250
vGrid.Height = Me.Height - (vGrid.top + 500)




End Sub

Private Sub TimerX_Timer()
TimerX.Interval = 0
TimerX.Enabled = False

Call sbConsulta


End Sub
