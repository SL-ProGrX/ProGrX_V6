VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#19.3#0"; "Codejock.Controls.v19.3.0.ocx"
Begin VB.Form frmPres_Analitico 
   Caption         =   "Análitico Contable"
   ClientHeight    =   6672
   ClientLeft      =   48
   ClientTop       =   396
   ClientWidth     =   12972
   LinkTopic       =   "Form1"
   ScaleHeight     =   6672
   ScaleWidth      =   12972
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer TimerX 
      Interval        =   10
      Left            =   9600
      Top             =   0
   End
   Begin XtremeSuiteControls.FlatEdit feModelo 
      Height          =   312
      Left            =   1200
      TabIndex        =   6
      Top             =   120
      Width           =   2652
      _Version        =   1245187
      _ExtentX        =   4678
      _ExtentY        =   550
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Locked          =   -1  'True
      Appearance      =   2
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit feContabilidad 
      Height          =   312
      Left            =   5640
      TabIndex        =   7
      Top             =   120
      Width           =   3252
      _Version        =   1245187
      _ExtentX        =   5736
      _ExtentY        =   550
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Locked          =   -1  'True
      Appearance      =   2
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit feUnidad 
      Height          =   312
      Left            =   5640
      TabIndex        =   8
      Top             =   600
      Width           =   3252
      _Version        =   1245187
      _ExtentX        =   5736
      _ExtentY        =   550
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Locked          =   -1  'True
      Appearance      =   2
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit feCentroCosto 
      Height          =   312
      Left            =   5640
      TabIndex        =   9
      Top             =   960
      Width           =   3252
      _Version        =   1245187
      _ExtentX        =   5736
      _ExtentY        =   550
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Locked          =   -1  'True
      Appearance      =   2
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit feCuenta 
      Height          =   312
      Left            =   1200
      TabIndex        =   10
      Top             =   600
      Width           =   2652
      _Version        =   1245187
      _ExtentX        =   4678
      _ExtentY        =   550
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Locked          =   -1  'True
      Appearance      =   2
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.FlatEdit fePeriodo 
      Height          =   312
      Left            =   1200
      TabIndex        =   11
      Top             =   960
      Width           =   2652
      _Version        =   1245187
      _ExtentX        =   4678
      _ExtentY        =   550
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Locked          =   -1  'True
      Appearance      =   2
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.PushButton btnExportar 
      Height          =   852
      Left            =   9000
      TabIndex        =   12
      Top             =   480
      Width           =   1332
      _Version        =   1245187
      _ExtentX        =   2350
      _ExtentY        =   1503
      _StockProps     =   79
      Caption         =   "Exportar"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   16
      Picture         =   "frmPres_Analitico.frx":0000
      TextImageRelation=   1
   End
   Begin FPSpreadADO.fpSpread vGrid 
      Height          =   5052
      Left            =   120
      TabIndex        =   13
      Top             =   1440
      Width           =   12612
      _Version        =   524288
      _ExtentX        =   22246
      _ExtentY        =   8911
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
      MaxCols         =   18
      RowHeaderDisplay=   0
      SpreadDesigner  =   "frmPres_Analitico.frx":0805
      VScrollSpecial  =   -1  'True
      VScrollSpecialType=   2
      AppearanceStyle =   1
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "Periodo "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   2
      Left            =   240
      TabIndex        =   5
      Top             =   960
      Width           =   1692
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "Unidad de Negocio"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   5
      Left            =   4080
      TabIndex        =   4
      Top             =   600
      Width           =   1692
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "Centro de Costo"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   6
      Left            =   4080
      TabIndex        =   3
      Top             =   960
      Width           =   1692
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "Modelo"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   0
      Left            =   240
      TabIndex        =   2
      Top             =   120
      Width           =   1692
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "Cuenta"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   1
      Left            =   240
      TabIndex        =   1
      Top             =   600
      Width           =   1692
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "Contabilidad"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   13
      Left            =   4080
      TabIndex        =   0
      Top             =   120
      Width           =   1692
   End
End
Attribute VB_Name = "frmPres_Analitico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSQL As String, rs As New ADODB.Recordset

Private Sub btnExportar_Click()
Dim vHeaders As vGridHeaders
Dim vArchivo As String, Ex As Boolean

    vHeaders.Columnas = 18
    vHeaders.Headers(1) = "Tipo Asiento"
    vHeaders.Headers(2) = "Num. Asiento"
    vHeaders.Headers(3) = "Fecha Asiento"
    vHeaders.Headers(4) = "Usr. Registra"
    vHeaders.Headers(5) = "Usr. Aplica"
    vHeaders.Headers(6) = "Cuenta"
    vHeaders.Headers(7) = "Unidad"
    vHeaders.Headers(8) = "Centro Costo"
    vHeaders.Headers(9) = "Divisa"
    vHeaders.Headers(10) = "Tipo Cambio"
    vHeaders.Headers(11) = "Importe"
    vHeaders.Headers(12) = "Mnt. Débito"
    vHeaders.Headers(13) = "Mnt. Crédito"
    vHeaders.Headers(14) = "Documento"
    vHeaders.Headers(15) = "Detalle"
    vHeaders.Headers(16) = "Referencia"
    vHeaders.Headers(17) = "Descripción"
    vHeaders.Headers(18) = "Anotaciones"
    
vArchivo = "Presupuesto_Analitico_" & Format(fePeriodo.Text, "yyyy-mm-dd") & "_" & feCuenta.Tag
Call sbSIFGridExportar(vGrid, vHeaders, vArchivo, "Excel")
End Sub

Private Sub Form_Load()
On Error GoTo vError
'spPres_Analitico_Descripciones(@Modelo varchar(10), @Contabilidad int
'                , @Cuenta varchar(30), @Unidad varchar(10) = Null, @CentroCosto varchar(10) = Null)
With gCntX_Presupuesto
    feModelo.Tag = .Modelo
    feContabilidad.Tag = .Contabilidad
    feCuenta.Tag = .Cuenta
    feUnidad.Tag = .Unidad
    feCentroCosto.Tag = .Centro
    fePeriodo.Text = Format(.Periodo, "yyyy/mm/dd")
    
    strSQL = "exec spPres_Analitico_Descripciones '" & .Modelo & "'," & .Contabilidad & ",'" & .Cuenta _
           & "','" & .Unidad & "','" & .Centro & "'"
    Call OpenRecordSet(rs, strSQL)
    
    feModelo.Text = rs!Modelo_Desc
    feContabilidad.Text = rs!Conta_Desc
    feUnidad.Text = rs!Unidad_Desc
    feCentroCosto.Text = rs!Centro_Desc
    feCuenta.Text = rs!Cuenta_Mask & ".." & rs!Cuenta_Desc
    
    rs.Close
End With

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub Form_Resize()

On Error Resume Next

vGrid.Width = Me.Width - 250
vGrid.Height = Me.Height - (vGrid.Top + 550)


End Sub

Private Sub TimerX_Timer()
TimerX.Interval = 0
TimerX.Enabled = False

Call sbAnalitico

End Sub



Private Sub sbAnalitico()
Dim pUnidad As String, pCentroCosto As String

Me.MousePointer = vbHourglass

On Error GoTo vError

Select Case feUnidad.Text
 Case "[TODOS]", "[CONSOLIDADO]", "TODOS", "CONSOLIDADO", "-C-"
    pUnidad = "Null"
 Case Else
    pUnidad = "'" & feUnidad.Tag & "'"
End Select


Select Case feCentroCosto.Text
 Case "[TODOS]", "[CONSOLIDADO]", "TODOS", "CONSOLIDADO", "-C-"
    pCentroCosto = "Null"
 Case Else
    pCentroCosto = "'" & feCentroCosto.Tag & "'"
End Select

If pUnidad = "-C-" Then pUnidad = "Null"
If pCentroCosto = "-C-" Then pCentroCosto = "Null"

strSQL = "exec spPres_Analitico " & feContabilidad.Tag & ",'" & Format(fePeriodo.Text, "yyyy/mm/dd") _
       & "','" & feCuenta.Tag & "'," & pUnidad & "," & pCentroCosto

vGrid.MaxRows = 0
Call sbCargaGrid(vGrid, vGrid.MaxCols, strSQL, False)

Me.MousePointer = vbDefault

Exit Sub

vError:
    Me.MousePointer = vbDefault
    MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub
