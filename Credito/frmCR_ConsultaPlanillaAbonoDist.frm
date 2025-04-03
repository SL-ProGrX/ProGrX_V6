VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "Codejock.Controls.v22.1.0.ocx"
Begin VB.Form frmCR_ConsultaPlanillaAbonoDist 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Cálculo de distribución de abono a operaciones"
   ClientHeight    =   6390
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   12960
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6390
   ScaleWidth      =   12960
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.PushButton btnConsulta 
      Height          =   372
      Left            =   10200
      TabIndex        =   4
      Top             =   960
      Width           =   1572
      _Version        =   1441793
      _ExtentX        =   2773
      _ExtentY        =   656
      _StockProps     =   79
      Caption         =   "Calcular"
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
      Picture         =   "frmCR_ConsultaPlanillaAbonoDist.frx":0000
   End
   Begin FPSpreadADO.fpSpread vGrid 
      Height          =   4812
      Left            =   120
      TabIndex        =   2
      Top             =   1440
      Width           =   12732
      _Version        =   524288
      _ExtentX        =   22458
      _ExtentY        =   8488
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
      ScrollBars      =   2
      SpreadDesigner  =   "frmCR_ConsultaPlanillaAbonoDist.frx":0727
      VScrollSpecialType=   2
      AppearanceStyle =   1
   End
   Begin XtremeSuiteControls.FlatEdit txtMonto 
      Height          =   312
      Left            =   1920
      TabIndex        =   3
      Top             =   960
      Width           =   1932
      _Version        =   1441793
      _ExtentX        =   3408
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
      Alignment       =   1
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin XtremeSuiteControls.DateTimePicker dtpCorte 
      Height          =   312
      Left            =   3960
      TabIndex        =   5
      Top             =   960
      Width           =   1332
      _Version        =   1441793
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
   Begin XtremeSuiteControls.ComboBox cboDeductora 
      Height          =   312
      Left            =   5400
      TabIndex        =   6
      Top             =   960
      Width           =   4572
      _Version        =   1441793
      _ExtentX        =   8070
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
      BackStyle       =   0  'Transparent
      Caption         =   "Monto a aplicar: "
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   1692
   End
   Begin VB.Label lblCliente 
      BackStyle       =   0  'Transparent
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   1200
      TabIndex        =   0
      Top             =   240
      Width           =   9255
   End
   Begin VB.Image imgBanner 
      Height          =   855
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12975
   End
End
Attribute VB_Name = "frmCR_ConsultaPlanillaAbonoDist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mCedula As String, mInstitucion As Integer, mProceso As Long, mFecha As Date, vPaso As Boolean


Private Sub btnConsulta_Click()
On Error GoTo vError


txtMonto.Text = Format(CCur(txtMonto.Text), "Standard")
Call sbConsultaInicial

vError:
End Sub

Private Sub cboDeductora_Click()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

'Carga ultimo monto enviado por Planillas
strSQL = "select isnull(sum(Cuota),0) as 'Monto', isnull(max(fecPro)," & mProceso & ") as 'Proceso'" _
       & "  From PRM_ENVIADO_DETALLE" _
       & " where COD_INSTITUCION = " & cboDeductora.ItemData(cboDeductora.ListIndex) & " and cedula = '" & mCedula & "'" _
       & "   and FECPRO in(select max(proceso)" _
       & "                   From PRM_BITACORA" _
       & "                   where COD_INSTITUCION = " & cboDeductora.ItemData(cboDeductora.ListIndex) & " and GESTION = 'E')"
Call OpenRecordSet(rs, strSQL)
If Not rs.EOF And Not rs.BOF Then
       txtMonto.Text = Format(rs!Monto, "Standard")
       mProceso = rs!Proceso
End If
rs.Close

Exit Sub

vError:

End Sub

Private Sub dtpCorte_Change()
Call btnConsulta_Click
End Sub

Private Sub Form_Activate()
vModulo = 3

End Sub

Private Sub Form_Load()
vModulo = 3

imgBanner.Picture = frmContenedor.imgBanner_01.Picture

mCedula = GLOBALES.gTag

txtMonto = "0.00"

Call sbInicializa
Call sbConsultaInicial


End Sub

Private Sub sbInicializa()
Dim strSQL As String, rs As New ADODB.Recordset

Me.MousePointer = vbHourglass

mProceso = GLOBALES.glngFechaCR
mInstitucion = 1
txtMonto.Text = 0

strSQL = "select cod_institucion,cedula,nombre,dbo.MyGetdate() as 'Fecha' from socios where cedula = '" & mCedula & "'"
Call OpenRecordSet(rs, strSQL)
If Not rs.EOF And Not rs.BOF Then
   mInstitucion = rs!cod_institucion
   mFecha = rs!fecha
   dtpCorte.Value = rs!fecha
   lblCliente.Caption = Trim(rs!Cedula) & " - " & Trim(rs!Nombre)
End If
rs.Close

vPaso = True
    strSQL = "exec spAFI_Institucion_Vinculadas " & mInstitucion & ",3"
    Call sbCbo_Llena_New(cboDeductora, strSQL, False, True)
vPaso = False

Call cboDeductora_Click

MousePointer = vbDefault

End Sub


Private Sub sbConsultaInicial()
Dim strSQL As String, rs As New ADODB.Recordset
Dim i As Integer

On Error GoTo vError

Me.MousePointer = vbHourglass


strSQL = "exec spPrmCreditoDetalleAbonos " & cboDeductora.ItemData(cboDeductora.ListIndex) & "," & mProceso & ",'" & mCedula _
        & "'," & CCur(txtMonto.Text) & ",'" & Format(dtpCorte.Value, "yyyy/mm/dd") & "','S',1,1,1"
Call OpenRecordSet(rs, strSQL)
With vGrid
 .MaxRows = 0
 Do While Not rs.EOF
     .MaxRows = .MaxRows + 1
     .Row = .MaxRows
     For i = 1 To 12
       .col = i
       Select Case i
         Case 1 'Operacion
            .Text = CStr(rs!Operacion)
         Case 2 'Linea
            .Text = CStr(rs!Linea)
         Case 3 'Tipo Abono
            .Text = CStr(rs!Tipo)
         Case 4 'Fecha Cuota (Proceso)
            .Text = Format(rs!Proceso, "####-##")
         Case 5 'Int.Cor.
            .Text = Format(rs!AbIntCor, "Standard")
         Case 6 'Int.Mor.
            .Text = Format(rs!AbIntMor, "Standard")
         Case 7 'Cargos
            .Text = Format(rs!AbCargo, "Standard")
         Case 8 'Poliza (Prevista)
            .Text = Format(0, "Standard")
         Case 9 'Amortiza
            .Text = Format(rs!abAmortiza, "Standard")
         Case 10 'Total
            .Text = Format(rs!AbIntCor + rs!AbIntMor + rs!AbCargo + rs!abAmortiza, "Standard")
         Case 11 'Compromiso Pendiente
            If rs!Tipo = "E" Then
                .Text = Format(rs!abAmortiza * -1, "Standard")
            Else
                .Text = Format((rs!IntCor + rs!IntMor + rs!Cargo + rs!Amortiza) - (rs!AbIntCor + rs!AbIntMor + rs!AbCargo + rs!abAmortiza), "Standard")
            End If
         Case 12 'Orden
            .Text = CStr(rs!Orden)
       End Select
     Next i
     rs.MoveNext
 Loop
End With
rs.Close

MousePointer = vbDefault
Exit Sub

vError:
 MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub



Private Sub txtMonto_GotFocus()
On Error GoTo vError
  txtMonto.Text = CCur(txtMonto.Text)
vError:
End Sub

Private Sub txtMonto_KeyDown(KeyCode As Integer, Shift As Integer)

On Error GoTo vError

If KeyCode = vbKeyReturn Then
   txtMonto.Text = Format(CCur(txtMonto.Text), "Standard")
   Call sbConsultaInicial
End If

vError:

End Sub

Private Sub txtMonto_LostFocus()
On Error GoTo vError
  txtMonto.Text = Format(CCur(txtMonto.Text), "Standard")
vError:
End Sub


