VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#20.2#0"; "Codejock.Controls.v20.2.0.ocx"
Begin VB.Form frmCR_ConveniosOrdenAsiento 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Asiento de la Orden de Liquidación:"
   ClientHeight    =   6825
   ClientLeft      =   30
   ClientTop       =   390
   ClientWidth     =   12840
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6825
   ScaleWidth      =   12840
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer TimerX 
      Interval        =   10
      Left            =   11520
      Top             =   720
   End
   Begin FPSpreadADO.fpSpread vGrid 
      Height          =   5052
      Left            =   120
      TabIndex        =   0
      Top             =   1320
      Width           =   12492
      _Version        =   524288
      _ExtentX        =   22034
      _ExtentY        =   8911
      _StockProps     =   64
      BackColorStyle  =   1
      BorderStyle     =   0
      EditEnterAction =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   9
      SpreadDesigner  =   "frmCR_ConveniosOrdenAsiento.frx":0000
      VScrollSpecial  =   -1  'True
      VScrollSpecialType=   2
      AppearanceStyle =   1
   End
   Begin XtremeSuiteControls.FlatEdit faDebitos 
      Height          =   252
      Left            =   9480
      TabIndex        =   1
      Top             =   6480
      Width           =   1452
      _Version        =   1310722
      _ExtentX        =   2561
      _ExtentY        =   444
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   "0"
      Alignment       =   1
      Appearance      =   2
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit faCreditos 
      Height          =   252
      Left            =   10920
      TabIndex        =   2
      Top             =   6480
      Width           =   1452
      _Version        =   1310722
      _ExtentX        =   2561
      _ExtentY        =   444
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   "0"
      Alignment       =   1
      Appearance      =   2
      UseVisualStyle  =   0   'False
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Estado:"
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
      Index           =   2
      Left            =   4440
      TabIndex        =   4
      Top             =   600
      Width           =   1332
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "No. Orden:"
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
      Index           =   1
      Left            =   1320
      TabIndex        =   5
      Top             =   600
      Width           =   1332
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Convenio:"
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
      Index           =   0
      Left            =   1320
      TabIndex        =   9
      Top             =   240
      Width           =   1332
   End
   Begin VB.Label lblEstado 
      BackStyle       =   0  'Transparent
      Caption         =   "(Estado)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   252
      Left            =   6000
      TabIndex        =   8
      Top             =   600
      Width           =   1332
   End
   Begin VB.Label lblOrden 
      BackStyle       =   0  'Transparent
      Caption         =   "(Orden)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   252
      Left            =   2880
      TabIndex        =   7
      Top             =   600
      Width           =   1932
   End
   Begin VB.Label lblConvenio 
      BackStyle       =   0  'Transparent
      Caption         =   "(Convenio)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   252
      Left            =   2880
      TabIndex        =   6
      Top             =   240
      Width           =   10092
   End
   Begin XtremeSuiteControls.Label Label1 
      Height          =   252
      Left            =   8040
      TabIndex        =   3
      Top             =   6480
      Width           =   1332
      _Version        =   1310722
      _ExtentX        =   2350
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Balance:"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
   End
   Begin VB.Image imgBanner 
      Height          =   1212
      Left            =   0
      Top             =   0
      Width           =   12972
   End
End
Attribute VB_Name = "frmCR_ConveniosOrdenAsiento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()

lblConvenio.Tag = GLOBALES.gTag
lblConvenio.Caption = GLOBALES.gTag3

lblOrden.Caption = GLOBALES.gTag2


Set imgBanner.Picture = frmContenedor.imgBanner_01.Picture

End Sub

Private Sub TimerX_Timer()
TimerX.Interval = 0
TimerX.Enabled = False


Dim strSQL As String, rs As New ADODB.Recordset
Dim curDebitos As Currency, curCreditos As Currency


On Error GoTo vError

Me.MousePointer = vbHourglass

curCreditos = 0
curDebitos = 0

strSQL = "exec spConvenios_Orden_Asiento '" & lblConvenio.Tag & "'," & lblOrden.Caption
Call OpenRecordSet(rs, strSQL)

With vGrid
  .MaxRows = 0
  Do While Not rs.EOF
    curDebitos = curDebitos + rs!Monto_Debito
    curCreditos = curCreditos + rs!Monto_Credito
    
    .MaxRows = .MaxRows + 1
    .Row = .MaxRows
    
    If rs!cod_factura = "(Pendiente)" Then
      lblEstado.Caption = "Abierta"
    Else
      lblEstado.Caption = "Cerrada"
    End If
    
    .Col = 1
    .Text = rs!cod_factura
    .Col = 2
    .Text = rs!cod_Cuenta_Mask
    .Col = 3
    .Text = rs!Descripcion
    .Col = 4
    .Text = rs!cod_unidad
    .Col = 5
    .Text = rs!cod_centro_costo
    .Col = 6
    .Text = rs!Cod_Divisa
    .Col = 7
    .Text = CStr(rs!Tipo_cambio)
    .Col = 8
    .Text = Format(rs!Monto_Debito, "Standard")
    .Col = 9
    .Text = Format(rs!Monto_Credito, "Standard")
    rs.MoveNext
  Loop
  rs.Close
End With

faDebitos.Text = Format(curDebitos, "Standard")
faCreditos.Text = Format(curCreditos, "Standard")

Me.MousePointer = vbDefault


Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub
